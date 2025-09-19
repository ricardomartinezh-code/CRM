# Diagnósticos operativos

## Calendario de auditorías

| Frecuencia | Hora (TZ script)   | Responsable                | Alcance                                                     |
| ---------- | ------------------ | -------------------------- | ----------------------------------------------------------- |
| Diario     | 07:00              | Automatización Apps Script | Base prioritaria configurada en `DIAGNOSTICS_PRIORITY_BASE` |
| Semanal    | Viernes 12:00      | Equipo de datos            | Revisión global de todas las bases                          |
| Mensual    | Primer lunes 10:00 | Operaciones                | Validación de métricas históricas                           |

> **Notas:**
>
> - El horario utiliza la zona configurada en el proyecto de Apps Script.
> - Cualquier ajuste debe registrarse en la propiedad correspondiente del script y comunicarse en el canal de datos.

## Protocolo de respuesta ante incidencias

1. **Recepción del alerta**
   - El sistema envía correo o webhook cuando `sheetsWithErrors > 0`.
   - Confirmar recepción en el canal `#ops-datos` en los primeros 15 minutos.
2. **Clasificación del hallazgo**
   - Registrar en la hoja _Control Diagnósticos_ el estado inicial y responsable asignado.
   - Etiquetar si afecta carga operativa, integridad histórica o indicadores críticos.
3. **Acciones de contención**
   - Bloquear cargas o asignaciones automatizadas sobre la base afectada.
   - Informar a supervisores si el impacto es operativo.
4. **Resolución**
   - Utilizar `handleRepairSheetStructure_` y herramientas de limpieza según corresponda.
   - Documentar pasos y resultados en la bitácora del caso.
5. **Cierre y seguimiento**
   - Actualizar la fila correspondiente en _Control Diagnósticos_ con fecha de cierre y aprendizajes.
   - Revisar causas raíz en la reunión semanal de operaciones de datos.

Mantener este documento actualizado tras cualquier cambio de calendario o proceso.
