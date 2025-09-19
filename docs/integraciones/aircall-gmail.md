# Webhooks de Aircall y Gmail para registrar toques

Esta guía documenta los endpoints disponibles en Apps Script para registrar toques de llamadas (Aircall) y correos (Gmail) en las hojas de leads. También incluye recomendaciones de seguridad y los pasos necesarios para configurar los webhooks en cada origen.

## Endpoint base

| Método | URL                                                       | Autenticación                    |
| ------ | --------------------------------------------------------- | -------------------------------- |
| `POST` | `https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec` | Token compartido por integración |

- El endpoint espera **JSON** en el cuerpo del POST. Para integraciones que no permiten editar el cuerpo (como Aircall) agrega `?action=<nombre>` en la URL.
- Se aceptan dos acciones:
  - `logAircallEvent` registra eventos de llamadas.
  - `logGmailThread` registra interacciones de hilos de correo.
- La respuesta devuelve `ok`, `sheet`, `rowNumber`, el `toque` aplicado y la lista actualizada de `toques`. En caso de evento duplicado (`duplicate: true`) no se escribe un nuevo toque.

### Seguridad con tokens compartidos

El script valida un token almacenado en propiedades del proyecto:

| Integración                  | Propiedad de Script     | Cómo se envía                                              |
| ---------------------------- | ----------------------- | ---------------------------------------------------------- |
| Aircall                      | `AIRCALL_WEBHOOK_TOKEN` | Encabezado `X-Integration-Token` o parámetro/campo `token` |
| Gmail (Apps Script auxiliar) | `GMAIL_WEBHOOK_TOKEN`   | Encabezado `X-Integration-Token` o parámetro/campo `token` |

Pasos para configurarlo:

1. En el editor de Apps Script abre **Project Settings → Script properties**.
2. Crea las propiedades anteriores y asigna un valor secreto.
3. Usa el mismo secreto en cada origen (Aircall o script auxiliar) al invocar el webhook.

> **Tip:** Puedes rotar el token en cualquier momento actualizando la propiedad y el origen.

#### Ejemplo de llamada autenticada

```bash
curl -X POST "https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec?action=logAircallEvent" \
  -H "Content-Type: application/json" \
  -H "X-Integration-Token: ${AIRCALL_WEBHOOK_TOKEN}" \
  -d '{
        "action": "logAircallEvent",
        "sheet": "Lead Recuperados",
        "token": "'"${AIRCALL_WEBHOOK_TOKEN}"'",
        "data": {
          "id": 123456,
          "direction": "inbound",
          "status": "answered",
          "duration": 98,
          "raw_digits": "+521234567890",
          "contact": {
            "name": "Ana Alumna",
            "external_id": "IMP-001234"
          }
        }
      }'
```

Si Aircall enviará firmas HMAC (`X-Aircall-Signature`), puedes conservarlas y extender `verifyIntegrationRequest_` para verificar la firma con la misma propiedad secreta, tal como lo harías con el token compartido.

## Payloads compatibles

### `logAircallEvent`

Los campos se normalizan automáticamente:

- **Identificación del lead:** teléfonos (`raw_digits`, `number.digits`, `contact.phone_numbers`, `from.number`, etc.), correos (`contact.emails`) y `contact.external_id`.
- **Fecha:** `ended_at`, `updated_at`, `started_at` o `timestamp`; se convierte a la zona horaria del proyecto.
- **Estado del toque:** combina `direction`, `status` y `result` (`"Entrante · Answered"`, por ejemplo).
- **Comentario:** agrega agente (`data.user.name`) y duración si están presentes.
- **Metadatos:** guarda en la columna `Metadatos` un registro resumido del evento (`integrationLogs`). Eventos con el mismo `id` y `source` se ignoran como duplicados.

Ejemplo de payload que entrega Aircall (se incluye token mediante encabezado):

```json
{
  "event": "call.ended",
  "data": {
    "id": 8459321,
    "direction": "outbound",
    "status": "voicemail",
    "result": "no_answer",
    "duration": 45,
    "raw_digits": "+521231231234",
    "number": { "digits": "+521231231234" },
    "contact": {
      "name": "Juan Prospecto",
      "external_id": "IMP-00999",
      "phone_numbers": [{ "value": "+52 123 123 1234" }],
      "emails": ["juan@example.com"]
    }
  }
}
```

### `logGmailThread`

Pensado para ejecutarse desde un Apps Script auxiliar que monitorea Gmail y reenvía los datos. Se normalizan:

- **Identificación del lead:** correos presentes en `thread.participants`, `messages[].from`, `messages[].to/cc/bcc`, etc.
- **Fecha:** `body.timestamp`, `thread.updated` o la fecha del último mensaje.
- **Estado del toque:** `Correo enviado`, `Correo recibido` o `Correo registrado` según `direction` (`outgoing`, `incoming`).
- **Comentario:** agrega el asunto (`subject`).
- **Metadatos:** registra `threadId`, `messageId`, participantes y etiquetas en `integrationLogs`.

Ejemplo de payload desde Apps Script:

```json
{
  "action": "logGmailThread",
  "token": "${GMAIL_WEBHOOK_TOKEN}",
  "sheet": "Lead Recuperados",
  "thread": {
    "id": "17c1c9f39ab12345",
    "historyId": "128937465",
    "subject": "Seguimiento UNIDEP",
    "link": "https://mail.google.com/mail/u/0/#inbox/17c1c9f39ab12345",
    "messages": [
      {
        "id": "17c1c9f39ab12346",
        "date": "2024-02-10T16:03:00-06:00",
        "direction": "outgoing",
        "from": "asesor@relead.mx",
        "to": ["ana.alumna@example.com"],
        "cc": [],
        "bcc": [],
        "subject": "Seguimiento UNIDEP",
        "snippet": "Hola Ana, te comparto..."
      }
    ]
  }
}
```

## Comportamiento en la hoja

- Se busca el lead en la base indicada (`sheet`) o, si se omite, en todas las hojas válidas.
- Se completan las columnas `Toque 1` a `Toque 4`. Si todas están ocupadas, se hace _shift_ y el nuevo valor queda en la última columna.
- Se rellena la asignación (`Asignación`/`Asignacion`) y `ActualizadoEl` si estaban vacíos.
- Se añade el comentario del evento solo una vez por registro.

## Configuración en Aircall

1. Inicia sesión en el **Dashboard de Aircall** con permisos de administrador.
2. Ve a **Integrations & API → Webhooks → Add a webhook**.
3. Define:
   - **URL:** `https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec?action=logAircallEvent`
   - **Method:** `POST`
   - **Secret token:** el valor de `AIRCALL_WEBHOOK_TOKEN`
   - Selecciona los eventos `call.ended`, `call.answered`, `call.missed` (según necesidad).
4. Guarda los cambios y realiza una llamada de prueba para verificar que se crea el toque en la hoja.

## Configuración del Apps Script auxiliar para Gmail

1. Crea un proyecto nuevo en Apps Script y habilita la API de Gmail (`Services → Gmail`).
2. Agrega el siguiente fragmento para enviar eventos al webhook:

```js
const WEBHOOK_URL = 'https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec?action=logGmailThread';
const WEBHOOK_TOKEN = PropertiesService.getScriptProperties().getProperty('GMAIL_WEBHOOK_TOKEN');

function notifyThread(thread) {
  const payload = {
    action: 'logGmailThread',
    token: WEBHOOK_TOKEN,
    sheet: 'Lead Recuperados',
    thread: mapThread(thread),
  };
  UrlFetchApp.fetch(WEBHOOK_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'X-Integration-Token': WEBHOOK_TOKEN },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
}
```

3. Implementa `mapThread` para serializar los mensajes necesarios y programa un **trigger** (por ejemplo, cada 5 minutos) que revise la bandeja y llame a `notifyThread` para los hilos nuevos/actualizados.
4. Guarda el token en las propiedades del proyecto auxiliar (`Script properties → GMAIL_WEBHOOK_TOKEN`).

## Validaciones recomendadas

- Automatiza pruebas con `curl` o un `UrlFetchApp.fetch` desde el editor de Apps Script para verificar que recibes `ok: true` y que las columnas de toques se actualizan.
- Conserva registros de respuesta (`response.getContentText()`) en tus integraciones para facilitar el diagnóstico.
- Si deseas migrar a **verificación de firmas** (p. ej. `X-Aircall-Signature` HMAC SHA256), extiende `verifyIntegrationRequest_` utilizando el mismo secreto almacenado en la propiedad.

Con estas configuraciones, todas las llamadas y correos quedarán reflejados en los toques de seguimiento de la hoja correspondiente.
