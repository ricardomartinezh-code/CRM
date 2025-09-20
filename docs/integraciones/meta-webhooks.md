# Webhooks de Meta para Facebook Lead Ads y WhatsApp Cloud API

Esta guía describe cómo exponer el Apps Script de ReLead EDU como webhook para Meta. La integración cubre:

- **Lead Ads (Facebook/Instagram):** descarga los datos completos del lead, crea el registro en la base indicada y adjunta metadatos para seguimiento.
- **WhatsApp Cloud API:** registra automáticamente toques entrantes/salientes cuando Meta envía eventos de mensajes o actualizaciones de estatus.

## Endpoint

| Método | URL                                                       | Descripción                                                             |
| ------ | --------------------------------------------------------- | ----------------------------------------------------------------------- |
| `GET`  | `https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec?action=metaWebhook` | Verificación inicial (`hub.challenge`).                                 |
| `POST` | `https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec?action=metaWebhook` | Recepción de eventos de Facebook Lead Ads y WhatsApp Cloud API.         |

> **Tip:** utiliza el mismo endpoint para ambos productos. Meta diferenciará cada evento mediante `field` (leadgen/messaging) y `messaging_product`.

## Propiedades de script necesarias

Configura las siguientes propiedades en **Apps Script → Project Settings → Script properties**:

| Propiedad | Uso |
| --------- | --- |
| `META_WEBHOOK_VERIFY_TOKEN` | Valor del token usado durante la verificación inicial del webhook (GET). Debe coincidir con el configurado en el dashboard de Meta. |
| `META_APP_SECRET` | Secreto de la app de Meta. Se usa para validar la firma `X-Hub-Signature-256`. |
| `META_LEAD_ACCESS_TOKEN` | Token con permisos `leads_retrieval` para consultar los detalles de cada lead generado. |
| `META_DEFAULT_LEAD_SHEET` | Nombre de la hoja destino donde se insertarán los leads cuando no haya una asignación específica. |
| `META_WHATSAPP_DEFAULT_SHEET` | (Opcional) Hoja preferida para registrar toques de WhatsApp cuando no se proporcione una base explícita. |
| `META_LEAD_SHEET_<FORM_ID>` | (Opcional) Si requieres dirigir diferentes formularios a distintas bases, crea una propiedad por `form_id`. Ejemplo: `META_LEAD_SHEET_123456789012345 = Lead Recuperados`. |

Sin `META_APP_SECRET` la firma se omitirá (no recomendado en producción). Sin `META_LEAD_ACCESS_TOKEN` no será posible descargar los campos del lead.

## Flujo de Facebook Lead Ads

1. Meta envía un evento `leadgen` con el `leadgen_id`.
2. El Apps Script valida la firma `X-Hub-Signature-256` y consulta la API de Graph:
   ```text
   GET https://graph.facebook.com/v19.0/<leadgen_id>?fields=field_data,created_time,form_id,form_name,ad_id,ad_name,adset_id,campaign_id,campaign_name
   ```
3. Con los datos obtenidos:
   - Se determina la base destino (`META_LEAD_SHEET_<FORM_ID>` → `META_DEFAULT_LEAD_SHEET`). También puedes definir un campo oculto `sheet`/`base` en el formulario.
   - Se crea el lead (ID generado `FB-<leadgen_id>`), normalizando nombre, teléfono, correo, campus, modalidad y programa.
   - Se marca como `Nuevo`, asignando la fecha/hora del evento y agregando el comentario “Lead generado en Facebook · <FormName>`.
   - En `Metadatos` se agrega `facebookLead` con `formId`, `formName`, IDs de campaña/anuncio y la copia de `field_data`. Cada evento queda registrado en `integrationLogs` (`source: facebook`).
4. Si el lead ya existe (por ID/telefono/correo) únicamente se actualiza el comentario y se añade el log al metadato. Eventos duplicados (mismo `leadgen_id`) se ignoran.

## Flujo de WhatsApp Cloud API

- Para cada cambio con `messaging_product: whatsapp`:
  - **`messages`:** se registra un toque “WhatsApp recibido” usando el teléfono (`from` o `to`). El comentario incluye el nombre del contacto (si Meta lo envía) y un resumen del mensaje (texto, botón, multimedia, etc.).
  - **`statuses`:** se registra “WhatsApp enviado” cuando Meta envía actualizaciones (`sent`, `delivered`, `read`). El resumen muestra el estatus y se adjuntan errores si existen.
- Todos los eventos guardan un `integrationLog` con `source: whatsapp`, el `wa_id`, el `phone_number_id` y el payload relevante.
- El webhook busca el lead por teléfono en todas las bases (o en la `META_WHATSAPP_DEFAULT_SHEET` si está definida). Si no encuentra coincidencia se omite el evento y se reporta en la respuesta JSON.

## Respuestas

El endpoint siempre devuelve JSON. Ejemplo de respuesta cuando se inserta un lead y se registra un toque:

```json
{
  "ok": true,
  "leads": [
    {
      "ok": true,
      "sheet": "Lead Recuperados",
      "rowNumber": 125,
      "created": true,
      "leadId": "FB-1234567890",
      "source": "facebook"
    }
  ],
  "touches": [
    {
      "ok": true,
      "sheet": "Lead Recuperados",
      "rowNumber": 125,
      "toque": "2024-11-18 10:25 · WhatsApp recibido",
      "toques": ["…"],
      "source": "whatsapp"
    }
  ]
}
```

Si ocurre algún problema (firma inválida, token faltante, lead sin base asignada, etc.) la respuesta incluirá `errors` con el detalle por evento.

## Configuración en Meta

1. Crea una app en [Meta for Developers](https://developers.facebook.com/), habilita **Facebook Login** (si aplica) y **Webhook** para `leadgen` y `whatsapp_business_account`.
2. Define el **Callback URL** con el parámetro `action=metaWebhook` y el **Verify token** igual al valor guardado en `META_WEBHOOK_VERIFY_TOKEN`.
3. Suscribe la app a la página (Lead Ads) y/o a la cuenta de WhatsApp Business.
4. Copia el **App Secret** y el **token de acceso** (con permiso `leads_retrieval`) en las propiedades del script.
5. Realiza una prueba desde el dashboard de Meta. Debes recibir `200 OK` con el `hub.challenge` en la verificación y `ok: true` en los POST.

Con esto podrás centralizar los leads generados en Facebook/Instagram y el historial de conversaciones de WhatsApp directamente en las hojas de ReLead EDU.
