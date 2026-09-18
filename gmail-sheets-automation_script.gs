/**
 * ============================================================================
 * AUTOMATIZACIÓN DE REGISTRO DE MOVIMIENTOS FINANCIEROS DESDE GMAIL
 * ============================================================================
 *
 * Qué hace este script:
 * Busca correos de Gmail relacionados a movimientos de dinero (boletas,
 * transferencias, compras, etc.), usa la API de Gemini para extraer los
 * datos estructurados de cada correo, y los guarda en la hoja "BD_Movimientos"
 * de este Google Sheet.
 *
 * Requisitos de configuración antes de ejecutar:
 * 1. En el editor de Apps Script, ir a Configuración del proyecto (ícono de
 *    engranaje) → Propiedades de secuencia de comandos → agregar una
 *    propiedad llamada "GEMINI_API_KEY" con tu clave de API de Gemini.
 * 2. Debe existir una hoja llamada "BD_Movimientos" en este spreadsheet.
 * 3. Se recomienda configurar un trigger (disparador) por tiempo para que
 *    esta función corra automáticamente cada cierto intervalo
 *    (Extensiones → Apps Script → Triggers → Add Trigger).
 *
 * Cómo funciona el manejo de errores (resumen):
 * - RATE_LIMIT: la API de Gemini está saturada (HTTP 429). Se aborta toda
 *   la ejecución para no seguir gastando cuota; se reintenta solo, en la
 *   siguiente corrida programada, sin penalizar al correo.
 * - CONFIGURACION: la API rechaza la clave o el modelo (HTTP 401/403/404).
 *   Esto no se arregla solo — se aborta todo y se envía un correo de
 *   alerta (con cooldown de 24h) para que el dueño del script lo revise.
 * - PERMANENTE: el contenido del correo o la respuesta del modelo tiene un
 *   problema real (bloqueo de seguridad, estructura vacía). Se marca el
 *   mensaje con una estrella y el hilo con la etiqueta de error, para
 *   revisión manual.
 * - TRANSITORIO: fallas de red, timeouts, o datos que no pasaron el
 *   sanity check. Se reintenta hasta 3 veces antes de tratarlo como
 *   PERMANENTE.
 * ============================================================================
 */

// Clave de API leída de forma segura desde las Propiedades del Script
// (nunca hardcodear la clave directamente en el código).
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");

// Nombre del modelo de Gemini a usar.
// IMPORTANTE: verificar en la documentación oficial de Gemini que este
// nombre de modelo exista y esté disponible antes de dejar el script
// corriendo desatendido.
const MODEL_NAME = "gemini-3.6-flash";

/**
 * Función principal. Pensada para ejecutarse periódicamente vía un trigger
 * de tiempo (ej. cada 15-30 minutos).
 *
 * Flujo general:
 * 1. Carga los IDs de mensajes ya guardados en la hoja (para no duplicar).
 * 2. Busca hilos de Gmail candidatos a contener movimientos de dinero.
 * 3. Por cada mensaje nuevo, llama a Gemini para extraer los datos.
 * 4. Según el resultado, guarda el movimiento, reintenta, marca error,
 *    o aborta todo si detecta un problema sistémico (cuota o configuración).
 * 5. Etiqueta cada hilo de Gmail según cómo terminó su procesamiento.
 *
 * No recibe parámetros ni retorna valor: opera directamente sobre el
 * spreadsheet activo y la bandeja de Gmail del usuario.
 */
function procesarMovimientosCorreo() {
  // --- 1. CRONÓMETRO DE EJECUCIÓN ---
  // Apps Script mata las ejecuciones que superan ~6 minutos. Cortamos
  // antes, de forma controlada, para no perder trabajo a mitad de camino.
  const startTime = Date.now();
  const MAX_EXECUTION_TIME = 4.5 * 60 * 1000; // 4.5 minutos (en milisegundos)

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("BD_Movimientos");
  const props = PropertiesService.getScriptProperties();

  // --- LÓGICA ANTI-DUPLICADOS ---
  // Leemos todos los IDs de mensaje ya guardados (columna A) y los
  // metemos en un Set para poder chequear existencia en O(1). Esto evita
  // volver a llamar a la API (y gastar cuota) por un correo ya procesado,
  // y sirve como red de seguridad extra ante cualquier falla de etiquetado.
  const lastRow = sheet.getLastRow();
  let idsGuardados = [];
  if (lastRow > 1) {
    idsGuardados = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  }
  const setIds = new Set(idsGuardados);

  // --- CONFIGURACIÓN DE ETIQUETAS DE GMAIL ---
  // "Movimientos_Procesados": hilo completado sin errores.
  // "Movimientos_Error": hilo con al menos un mensaje que falló de forma
  // permanente (o agotó sus reintentos) y necesita revisión manual.
  const labelName = "Movimientos_Procesados";
  const errorLabelName = "Movimientos_Error";

  let label = GmailApp.getUserLabelByName(labelName) || GmailApp.createLabel(labelName);
  let errorLabel = GmailApp.getUserLabelByName(errorLabelName) || GmailApp.createLabel(errorLabelName);

  // Búsqueda de hilos candidatos: últimos 7 días, que NO tengan ya
  // ninguna de las dos etiquetas (evita reprocesar hilos ya resueltos
  // o ya marcados como error), y que contengan palabras clave típicas
  // de movimientos de dinero.
  const query = `newer_than:7d -label:${labelName} -label:${errorLabelName} (compra OR cargo OR boleta OR transferencia OR abono OR deposito OR "pago recibido" OR devolucion OR "comprobante")`;
  const threads = GmailApp.search(query, 0, 15); // Máx 15 hilos por corrida

  // Bandera de "freno de emergencia": se activa cuando detectamos un
  // problema que afecta a TODAS las llamadas siguientes (cuota agotada
  // o configuración rota), no solo al mensaje actual. Al activarse,
  // se corta tanto el loop de mensajes como el de hilos.
  let abortarGlobal = false;

  for (const thread of threads) {
    // Si ya se acabó el tiempo disponible o se activó el freno de
    // emergencia en una vuelta anterior, no seguimos procesando más
    // hilos. Los hilos no procesados simplemente quedan para la
    // próxima ejecución (no llevan ninguna etiqueta nueva).
    if (Date.now() - startTime > MAX_EXECUTION_TIME || abortarGlobal) {
      Logger.log("Ejecución global interrumpida (Timeout, Cuota o Configuración).");
      break;
    }

    const messages = thread.getMessages();

    // threadProcesadoCompleto: se pone en false apenas un mensaje falla,
    // de cualquier tipo. Solo si sigue en true al final, el hilo se
    // marca como "Procesado".
    let threadProcesadoCompleto = true;

    // threadTieneErrorPermanente: se activa cuando algún mensaje del
    // hilo falló de forma definitiva (o agotó sus 3 reintentos). Se usa
    // para decidir si el hilo se etiqueta como error al final.
    let threadTieneErrorPermanente = false;

    for (const message of messages) {
      const msgId = message.getId();

      // Si este mensaje puntual ya fue guardado en una corrida anterior,
      // lo saltamos sin gastar cuota de la API.
      if (setIds.has(msgId)) {
        continue;
      }

      const bodyText = message.getPlainBody().substring(0, 4000);
      const mov = extraerDatosMovimiento(bodyText);

      // --- CASO: LA EXTRACCIÓN DEVOLVIÓ UN ERROR ---
      if (mov && mov.error_type) {
        threadProcesadoCompleto = false;

        // RATE_LIMIT o CONFIGURACION son errores "sistémicos": no tiene
        // sentido seguir intentando con otros mensajes/hilos en esta
        // misma corrida, porque van a fallar por la misma razón.
        if (mov.error_type === 'RATE_LIMIT' || mov.error_type === 'CONFIGURACION') {
          Logger.log(`Error sistémico (${mov.error_type}) en msg ${msgId}. Tirando del freno de emergencia.`);
          abortarGlobal = true;

          // Solo para errores de configuración (clave inválida, modelo
          // inexistente, permisos) mandamos una alerta por correo, porque
          // a diferencia del rate limit, este tipo de error NO se resuelve
          // solo con el paso del tiempo — requiere intervención humana.
          if (mov.error_type === 'CONFIGURACION') {
            // Cooldown de 24h para no saturar la bandeja si el trigger
            // corre varias veces al día mientras el problema sigue sin
            // resolverse.
            const ultimaAlerta = parseInt(props.getProperty('alerta_config_enviada') || '0');
            const COOLDOWN_24H = 24 * 60 * 60 * 1000;

            if (Date.now() - ultimaAlerta > COOLDOWN_24H) {
              const miCorreo = Session.getActiveUser().getEmail();
              const asunto = "⚠️ ALERTA: Falla en Automatización de Gastos";
              const cuerpo = `Hola,\n\nTu script de registro automático de gastos ha sido detenido por un error de configuración (HTTP 401, 403 o 404).\n\nPosibles causas:\n- Tu API Key de Gemini expiró o es inválida.\n- El modelo (${MODEL_NAME}) no existe o cambió de nombre.\n\nPor favor revisa los logs en Apps Script para solucionarlo.\n\nEsta alerta no se repetirá hasta dentro de 24 horas para evitar saturar tu bandeja.`;

              MailApp.sendEmail(miCorreo, asunto, cuerpo);
              props.setProperty('alerta_config_enviada', Date.now().toString());
              Logger.log("Correo de alerta enviado exitosamente.");
            }
          }

          break; // Corta el loop de mensajes; el loop de hilos se corta arriba
        }

        // PERMANENTE: el problema es específico de este correo/respuesta
        // (contenido bloqueado por seguridad, estructura vacía, etc.).
        // No tiene sentido reintentar — se marca directamente.
        else if (mov.error_type === 'PERMANENTE') {
          Logger.log(`Error de contenido/seguridad en ${msgId}. Descartando inmediatamente.`);
          threadTieneErrorPermanente = true;
          message.star(); // Marca visual sobre el mensaje específico que falló
          props.deleteProperty('retry_' + msgId);
        }

        // TRANSITORIO: falla de red, timeout, o dato que no pasó el
        // sanity check. Vale la pena reintentar unas pocas veces, ya
        // que puede resolverse solo (o el modelo puede acertar en el
        // siguiente intento).
        else if (mov.error_type === 'TRANSITORIO') {
          let intentos = parseInt(props.getProperty('retry_' + msgId) || '0');
          intentos++;

          if (intentos >= 3) {
            // Se agotaron los reintentos: se trata igual que un error
            // permanente a partir de ahora.
            Logger.log(`Fallo transitorio superó 3 intentos en ${msgId}. Marcando como permanente.`);
            threadTieneErrorPermanente = true;
            message.star();
            props.deleteProperty('retry_' + msgId);
          } else {
            // Guardamos el contador para que la próxima ejecución sepa
            // en qué intento va este mensaje específico.
            Logger.log(`Error transitorio en ${msgId}. Intento ${intentos}/3`);
            props.setProperty('retry_' + msgId, intentos.toString());
          }
        }
      }

      // --- CASO: ÉXITO ---
      // mov.es_movimiento !== undefined confirma que recibimos un objeto
      // de datos válido (no un objeto de error).
      else if (mov && mov.es_movimiento !== undefined) {
        if (mov.es_movimiento) {
          sheet.appendRow([
            msgId, mov.fecha, mov.tipo_movimiento, mov.entidad,
            mov.monto, mov.moneda, mov.categoria, mov.metodo_cuenta, mov.detalle
          ]);
          setIds.add(msgId); // Evita duplicados si el mismo hilo se vuelve a leer más abajo
        }
        // Si Gemini determinó que el correo no es un movimiento real
        // (es_movimiento: false), simplemente no se guarda nada, pero
        // igual se considera un mensaje "resuelto" exitosamente.

        props.deleteProperty('retry_' + msgId);
        // Si este mensaje tuvo éxito, reseteamos también el cooldown de
        // alertas de configuración: significa que la API volvió a
        // responder bien, así que un futuro error de configuración
        // debería poder notificarse de inmediato otra vez.
        props.deleteProperty('alerta_config_enviada');
      }

      // Pausa entre llamadas para cuidar la cuota de la API de Gemini.
      Utilities.sleep(15000);
    } // Fin del for de messages

    // --- CLASIFICACIÓN FINAL DEL HILO ---
    // El chequeo "&& !abortarGlobal" evita etiquetar un hilo cuando la
    // razón de no completarse fue un aborto global (cuota/configuración)
    // y no un problema real de ese hilo — así se reintenta completo en
    // la próxima corrida en vez de quedar marcado incorrectamente.
    if (threadProcesadoCompleto && !abortarGlobal) {
      thread.addLabel(label);
    } else if (threadTieneErrorPermanente && !abortarGlobal) {
      thread.addLabel(errorLabel);
    }
  } // Fin del for de threads
}

/**
 * Llama a la API de Gemini para extraer los datos estructurados de un
 * movimiento de dinero a partir del texto de un correo.
 *
 * @param {string} textoCorreo - Cuerpo en texto plano del correo a analizar
 *   (ya truncado a 4000 caracteres antes de llegar acá).
 * @return {Object} Uno de los siguientes:
 *   - Objeto de datos del movimiento (con es_movimiento, monto, fecha, etc.)
 *     si la extracción fue exitosa.
 *   - { error_type: 'RATE_LIMIT' }    si la API está saturada (HTTP 429).
 *   - { error_type: 'CONFIGURACION' } si la clave/modelo/permisos fallan
 *                                     (HTTP 401/403/404).
 *   - { error_type: 'PERMANENTE' }    si el contenido fue bloqueado o la
 *                                     respuesta vino con estructura inválida.
 *   - { error_type: 'TRANSITORIO' }   si hubo un error de red/servidor, o
 *                                     los datos no pasaron el sanity check.
 */
function extraerDatosMovimiento(textoCorreo) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL_NAME}:generateContent?key=${GEMINI_API_KEY}`;

  // Prompt con reglas explícitas de extracción y normalización de datos.
  // Ver el punto 5 en particular: asume formato numérico chileno (punto
  // como separador de miles), ajustar si se usa con correos de otro país.
  const prompt = `
  Eres un analista financiero experto. Tu tarea es extraer información estructurada sobre movimientos de dinero (gastos, compras, ingresos, transferencias) a partir del texto crudo de correos electrónicos bancarios, boletas o comprobantes.
  
  Reglas estrictas de extracción:
  1. "es_movimiento": Evalúa a true SOLO si el correo representa una transacción real, efectiva y consolidada. Evalúa a false si es "ruido financiero" (publicidad, resúmenes de cuenta, transferencias entre mis propias cuentas/traspasos, o notificaciones de inicio de sesión).
  2. "tipo_movimiento": Clasifica obligatoriamente como "Ingreso" (dinero que entra a mi favor) o "Gasto" (dinero que sale).
  3. "fecha": Debe tener el formato exacto YYYY-MM-DD. Infiere la fecha a partir de menciones como "hoy" o del encabezado del correo si es necesario.
  4. "entidad": Limpia el nombre del comercio o contraparte. Si dice "COMPRA EN MP *MELIMAS CONCEPCION", devuelve "Melimas". Si es una transferencia, extrae el nombre limpio de la persona o empresa.
  5. "monto": Extrae solo el valor numérico, siempre positivo. Elimina símbolos ($) y puntos separadores de miles (ej. si dice $15.000, devuelve 15000).
  6. "moneda": Usa el código ISO de 3 letras. Asume "CLP" por defecto para transacciones locales, a menos que se especifique explícitamente "USD" u otra.
  7. "categoria": Analiza la entidad y el contexto para clasificar el movimiento en una de las siguientes categorías fijas:
     - Supermercado / Despensa
     - Comida Fuera / Delivery
     - Transporte / Ciclismo
     - Servicios / Suscripciones
     - Mascotas
     - Educación
     - Salud / Farmacia
     - Compras / Retail
     - Ingresos
     - Otros / No Identificado
  8. "metodo_cuenta": Identifica de dónde salió o a dónde entró el dinero (ej. "Cuenta Corriente", "Tarjeta de Débito", "Tarjeta de Crédito").
  9. "detalle": Crea una descripción muy breve (máximo 5 palabras) del concepto si es evidente (ej. "Repuestos bicicleta", "Comida gatos", "Mensualidad", "Suscripción YouTube").

  Texto del correo a analizar:
  """
  ${textoCorreo}
  """
  `;

  // responseSchema fuerza a Gemini a devolver JSON con esta forma exacta,
  // reduciendo (aunque no eliminando del todo) el riesgo de respuestas
  // mal formadas.
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      responseMimeType: "application/json",
      responseSchema: {
        type: "OBJECT",
        properties: {
          es_movimiento: { type: "BOOLEAN" },
          tipo_movimiento: { type: "STRING", enum: ["Ingreso", "Gasto"] },
          fecha: { type: "STRING", description: "Formato YYYY-MM-DD" },
          entidad: { type: "STRING", description: "Comercio, empresa o persona emisora/receptora" },
          monto: { type: "NUMBER", description: "Monto positivo" },
          moneda: { type: "STRING", description: "Ej: CLP, USD, EUR" },
          categoria: { type: "STRING" },
          metodo_cuenta: { type: "STRING", description: "Ej: Débito, Crédito, Transferencia, Cuenta Corriente" },
          detalle: { type: "STRING" }
        },
        required: ["es_movimiento"]
      }
    }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    // muteHttpExceptions permite leer el body del error en vez de que
    // UrlFetchApp lance una excepción directamente.
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    const json = JSON.parse(response.getContentText());

    // --- 1. VALIDACIÓN DE CÓDIGO HTTP ---
    if (code !== 200) {
      const errorMsg = json.error ? json.error.message : "Error desconocido";
      Logger.log(`API Rechazada (HTTP ${code}): ${errorMsg}`);

      if (code === 429) return { error_type: 'RATE_LIMIT' };

      // 401 (no autorizado), 403 (prohibido) o 404 (modelo no encontrado)
      // son señales de que algo está mal configurado, no del contenido
      // del correo puntual.
      if (code === 401 || code === 403 || code === 404) return { error_type: 'CONFIGURACION' };

      // Cualquier otro código 4xx (ej. 400 Bad Request) se trata como
      // un problema específico de esta solicitud/contenido.
      if (code >= 400 && code < 500) return { error_type: 'PERMANENTE' };

      // 5xx: error del lado del servidor de Google, típicamente transitorio.
      return { error_type: 'TRANSITORIO' };
    }

    // --- 2. FILTRO DE SEGURIDAD DE ENTRADA ---
    // Gemini puede bloquear el prompt completo antes de generar nada.
    if (json.promptFeedback && json.promptFeedback.blockReason) {
      Logger.log(`Bloqueo de seguridad (promptFeedback). Razón: ${json.promptFeedback.blockReason}`);
      return { error_type: 'PERMANENTE' };
    }

    // --- 3. VALIDACIÓN DE ESTRUCTURA DE RESPUESTA ---
    if (!json.candidates || json.candidates.length === 0) {
      Logger.log("La API no devolvió candidatos. Estructura vacía.");
      return { error_type: 'PERMANENTE' };
    }

    const candidate = json.candidates[0];

    // --- 4. FILTRO DE SEGURIDAD DE SALIDA ---
    // finishReason distinto de "STOP" indica que la generación se
    // interrumpió (por seguridad, por límite de tokens, etc.), por lo
    // que el contenido puede venir incompleto o ausente.
    if (candidate.finishReason && candidate.finishReason !== "STOP") {
      Logger.log(`Proceso interrumpido (finishReason). Razón: ${candidate.finishReason}`);
      return { error_type: 'PERMANENTE' };
    }

    const resultado = JSON.parse(candidate.content.parts[0].text);

    // --- 5. SANITY CHECK: validación lógica de los datos ya parseados ---
    // El responseSchema garantiza la FORMA del JSON, pero no que los
    // valores tengan sentido. Estas validaciones atrapan alucinaciones
    // típicas del modelo antes de que lleguen a la hoja de cálculo.
    if (resultado.es_movimiento) {
      // 5.1 Monto: debe ser un número real, finito, mayor a 0.
      // OJO: en JS, typeof NaN === 'number', por eso se valida isNaN()
      // por separado además de typeof.
      if (typeof resultado.monto !== 'number' || isNaN(resultado.monto) || resultado.monto <= 0) {
        Logger.log(`Sanity Check Falló: Monto inválido o NaN (${resultado.monto}). Forzando reintento.`);
        return { error_type: 'TRANSITORIO' };
      }

      // 5.2 Fecha: debe cumplir estrictamente el formato YYYY-MM-DD.
      const regexFecha = /^\d{4}-\d{2}-\d{2}$/;
      if (!resultado.fecha || !regexFecha.test(resultado.fecha)) {
        Logger.log(`Sanity Check Falló: Fecha inválida (${resultado.fecha}). Forzando reintento.`);
        return { error_type: 'TRANSITORIO' };
      }

      // 5.3 Entidad: no puede venir vacía o solo con espacios.
      if (!resultado.entidad || resultado.entidad.trim() === "") {
        Logger.log("Sanity Check Falló: Entidad vacía. Forzando reintento.");
        return { error_type: 'TRANSITORIO' };
      }
    }

    // Si es_movimiento es false, no se aplican los chequeos anteriores
    // (no hay monto/fecha/entidad que validar) y se retorna tal cual.
    return resultado;

  } catch (e) {
    // Captura errores de red (timeout, sin conexión) o de parseo de JSON
    // (respuesta malformada que ni siquiera se pudo interpretar).
    Logger.log("Error en parseo JSON o timeout de red: " + e.toString());
    return { error_type: 'TRANSITORIO' };
  }
}
