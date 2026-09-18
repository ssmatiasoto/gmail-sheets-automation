# Gmail Movimientos → Google Sheets (con Gemini)

Script de Google Apps Script que busca correos de Gmail relacionados a movimientos de dinero (boletas, transferencias, compras, etc.), usa la API de Gemini para extraer los datos estructurados de cada correo, y los guarda automáticamente en una hoja de Google Sheets.

## Qué hace

1. Busca en Gmail correos de los últimos 7 días que contengan palabras clave típicas de movimientos financieros (compra, cargo, boleta, transferencia, abono, depósito, etc.).
2. Por cada correo nuevo, envía el texto a la API de Gemini junto con un set de reglas de extracción (tipo de movimiento, monto, entidad, categoría, etc.).
3. Guarda cada movimiento detectado en la hoja `BD_Movimientos`.
4. Etiqueta los hilos de Gmail según el resultado:
   - `Movimientos_Procesados`: el hilo se procesó sin errores.
   - `Movimientos_Error`: al menos un mensaje del hilo falló de forma permanente y necesita revisión manual (el mensaje específico queda marcado con una ⭐).
5. Evita duplicados y reintenta automáticamente ante errores transitorios (hasta 3 veces por mensaje).
6. Si detecta un problema de configuración (clave inválida, modelo inexistente, permisos), aborta la ejecución y envía un correo de alerta.

## Configuración inicial

### 1. Crear la hoja de destino
En tu Google Sheet, crea una hoja llamada exactamente `BD_Movimientos`. Las columnas esperadas, en orden, son:

| A (ID mensaje) | B (fecha) | C (tipo) | D (entidad) | E (monto) | F (moneda) | G (categoría) | H (método/cuenta) | I (detalle) |
|---|---|---|---|---|---|---|---|---|

No es necesario escribir encabezados manualmente, pero puedes agregarlos en la fila 1 si quieres (el script empieza a leer/escribir desde la fila 2).

### 2. Obtener una API key de Gemini
Consigue tu clave en [Google AI Studio](https://aistudio.google.com/).

### 3. Guardar la clave de forma segura (nunca en el código)
En el editor de Apps Script:
1. Ve a **Configuración del proyecto** (ícono de engranaje, barra lateral izquierda).
2. Baja hasta **Propiedades de secuencia de comandos**.
3. Agrega una propiedad:
   - Nombre: `GEMINI_API_KEY`
   - Valor: tu clave de API

### 4. Verificar el nombre del modelo
Antes de correrlo por primera vez, confirma en la [documentación de Gemini](https://ai.google.dev/gemini-api/docs/models) que el valor de la constante `MODEL_NAME` en el código corresponde a un modelo real y disponible. Si el nombre está mal, la primera ejecución debería fallar rápido con un error de tipo `CONFIGURACION` (HTTP 404) y notificarte por correo — es la forma más simple de detectarlo.

### 5. Pegar el código
Copia el contenido de [`procesarMovimientosCorreo.gs`](./procesarMovimientosCorreo.gs) dentro de un archivo `.gs` en el editor de Apps Script (Extensiones → Apps Script, desde tu Google Sheet).

### 6. Configurar un trigger de tiempo
Para que el script corra solo:
1. En el editor de Apps Script, ve a **Triggers** (ícono de reloj, barra lateral izquierda).
2. **Add Trigger**.
3. Función a ejecutar: `procesarMovimientosCorreo`.
4. Fuente del evento: **Basado en tiempo**.
5. Elige una frecuencia razonable (ej. cada 30 minutos o cada hora), considerando que cada ejecución procesa hasta ~15 hilos con una pausa de 15 segundos entre mensajes.

## Manejo de errores

| Tipo | Causa | Comportamiento |
|---|---|---|
| `RATE_LIMIT` | La API de Gemini devuelve HTTP 429 (cuota excedida) | Se corta toda la ejecución de inmediato. Se reintenta solo en la siguiente corrida programada. |
| `CONFIGURACION` | HTTP 401 / 403 / 404 (clave inválida, sin permisos, o modelo inexistente) | Se corta toda la ejecución y se envía un correo de alerta (máximo una vez cada 24 horas). |
| `PERMANENTE` | Contenido bloqueado por seguridad, respuesta vacía o mal formada | Se marca el mensaje con ⭐ y el hilo con la etiqueta `Movimientos_Error`. No se reintenta. |
| `TRANSITORIO` | Error de red, timeout, o datos que no pasan la validación (monto inválido, fecha con formato incorrecto, entidad vacía) | Se reintenta hasta 3 veces. Al tercer fallo, se trata como `PERMANENTE`. |

## Limitaciones conocidas

- Cada ejecución tiene un límite interno de 4.5 minutos para evitar que Apps Script mate el proceso a la fuerza (límite real: ~6 minutos). Si hay más correos pendientes de los que alcanzan a procesarse en ese tiempo, simplemente quedan para la siguiente corrida.
- El control de tiempo se evalúa entre hilos, no entre mensajes individuales — un hilo con muchos mensajes podría hacer que una ejecución se extienda algo más allá del límite planeado.
- El prompt asume formato de números chileno (punto como separador de miles) y CLP como moneda por defecto. Ajustar la sección de reglas del prompt si se usa en otro país.
- La pausa de 15 segundos entre mensajes es conservadora para cuidar la cuota gratuita de la API; puede ajustarse según el límite real de tu plan de Gemini.

## Seguridad

- La API key nunca debe quedar escrita directamente en el código. Se lee desde `PropertiesService.getScriptProperties()`.
- Si en algún momento una clave quedó expuesta en texto plano (por ejemplo, en un commit anterior), lo más seguro es revocarla y generar una nueva, independientemente de si el repositorio es público o privado.
