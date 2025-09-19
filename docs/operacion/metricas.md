# Métricas operativas

El endpoint `doGet` admite la acción `operationalMetrics` para obtener un resumen
por base y asesor reutilizando la lógica del reporte estándar. Esta salida está
pensada para integrarse con tableros (por ejemplo, Looker Studio) o con la hoja
de soporte que se actualiza con un activador horario.

## Parámetros

| Parámetro | Tipo | Descripción |
| --- | --- | --- |
| `action` | `string` | Debe establecerse en `operationalMetrics`. |
| `bases` / `base` / `sheet` | `string` | Lista separada por comas, punto y coma o saltos de línea con los nombres de base a incluir. Si no se envía, se consultan todas las bases accesibles para el usuario autenticado. |
| `asesores` / `asesor` | `string` | Lista separada por comas, punto y coma o saltos de línea con los asesores a filtrar. Si se omite, se incluyen todos los asesores detectados por base. |
| `type` | `string` | Igual que en `handleReport_`; admite `estado`, `etapa` o `detalle`. Valor por defecto: `estado`. |
| `etapa` | `string` | Filtra las métricas únicamente a la etapa indicada. |
| `fi` | `string` | Fecha inicial (`Asignación`) en formato ISO 8601. |
| `ff` | `string` | Fecha final (`Asignación`) en formato ISO 8601. |
| `callback` | `string` | Opcional para respuestas JSONP. |

> **Nota:** Los filtros de listas (`bases`, `asesores`) ignoran duplicados y
valores vacíos.

## Estructura de respuesta

La respuesta tiene el siguiente formato:

```json
{
  "data": {
    "generatedAt": "2024-05-01T00:00:00.000Z",
    "type": "estado",
    "filters": {
      "fi": "2024-04-01",
      "ff": "2024-04-30",
      "etapa": "",
      "bases": ["Lead Recuperados"],
      "asesores": []
    },
    "bases": [
      {
        "name": "Lead Recuperados",
        "total": 42,
        "metrics": {
          "Inscrito": 10,
          "No contactado": 12,
          "Contactado": 20
        },
        "asesores": [
          {
            "name": "Ana Pérez",
            "total": 12,
            "metrics": {
              "Inscrito": 4,
              "No contactado": 3,
              "Contactado": 5
            }
          }
        ]
      }
    ],
    "summary": {
      "totalBases": 1,
      "totalLeads": 42,
      "metrics": {
        "Inscrito": 10,
        "No contactado": 12,
        "Contactado": 20
      }
    }
  }
}
```

Los valores en `metrics` representan los conteos calculados por `handleReport_`
para el tipo seleccionado. El campo `total` corresponde a la suma de cada
sección (`base` o `asesor`).

## Automatización y activador horario

El archivo `apps-script.gs` incluye la función `syncOperationalMetricsJob`, que
invoca internamente `handleOperationalMetrics_`, normaliza la salida y la
registra en la hoja **Operational Metrics**. Esta hoja puede conectarse como
fuente a Looker Studio o utilizarse como respaldo histórico.

Para programar la actualización automática:

1. Abre el editor de Apps Script del proyecto.
2. Ejecuta la función `ensureOperationalMetricsTrigger` una vez para crear un
   activador horario que llame a `syncOperationalMetricsJob` cada hora.
3. Verifica en la pestaña de **Activadores** que el disparador se haya creado
   correctamente.

Cada ejecución agrega nuevas filas con las columnas: `Timestamp`, `Base`,
`Asesor`, `Nivel`, `Metric`, `Valor` y `Total`. Puedes duplicar o limpiar la
hoja según tus necesidades de retención de datos.

