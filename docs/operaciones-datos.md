# Procedimientos operativos de datos

## Importar datos: institucionalizar el paso de verificación

1. **Recepción del archivo**
   - Registrar fecha de entrega, responsable y origen del archivo.
   - Clasificar el archivo como "base fría" si proviene de una fuente histórica o externa.
2. **Preparación para la validación**
   - Cargar el archivo en el entorno de staging de Importar datos.
   - Ejecutar controles automáticos de formato (estructura de columnas, tipos de dato, codificación) y documentar los resultados.
3. **Paso de Verificar (obligatorio)**
   - Revisar alertas generadas por los controles automáticos y corregirlas en el archivo fuente.
   - Ejecutar validaciones manuales: conteo de filas vs. remitido, muestreo aleatorio de registros críticos, consistencia de llaves.
   - Registrar la aprobación del analista responsable y archivar evidencia (capturas, reportes de validación).
4. **Autorización para ejecutar**
   - Solo si el paso de Verificar está aprobado, liberar el archivo a la cola de ejecución.
   - Marcar el lote con el identificador de validación y adjuntar el informe de hallazgos.

## Reconciliación de actualizaciones con el flujo de Activos

1. **Ingreso al módulo de Activos**
   - Cargar los nuevos registros o actualizaciones en modo análisis.
   - Asociar cada registro con su identificador único existente en la base maestra.
2. **Revisión de coincidencias**
   - Comparar cada campo relevante con el registro existente y resaltar diferencias.
   - Generar un reporte de discrepancias que incluya: campo, valor actual, valor propuesto y severidad.
3. **Acciones de reconciliación**
   - Validar con el área dueña del dato cualquier discrepancia crítica antes de confirmar el cambio.
   - Documentar en el registro la decisión (aceptar, rechazar, aplazar) y su justificación.
4. **Cierre del análisis**
   - Solo aplicar actualizaciones cuya reconciliación esté aprobada.
   - Archivar el reporte de discrepancias para auditoría.

## Procedimiento de cambios masivos

1. **Solicitud y planeación**
   - Definir alcance (entidades afectadas, campos a modificar, volumen estimado).
   - Establecer responsables de revisión y aprobación.
2. **Controles previos**
   - Ejecutar detección de duplicados (por identificadores clave y combinaciones relevantes).
   - Generar una vista previa del impacto (registros afectados, valores actuales vs. propuestos).
3. **Validación**
   - Revisar la vista previa con los responsables y documentar observaciones.
   - Ajustar el lote hasta obtener aprobación formal.
4. **Ejecución**
   - Programar el cambio en una ventana controlada.
   - Monitorear logs durante la ejecución y registrar incidencias.
5. **Verificación posterior**
   - Comparar el resultado con la vista previa aprobada.
   - Emitir un informe final con métricas de éxito, duplicados resueltos y acciones pendientes.

## Uso regular del módulo de Diagnóstico y plan de solución

1. **Frecuencia**: Ejecutar el Diagnóstico semanalmente o después de cambios masivos.
2. **Actividades**
   - Identificar duplicados, registros inconsistentes y celdas con formatos inválidos.
   - Priorizar hallazgos según impacto operativo.
3. **Plan de solución**
   - Asignar responsables y fechas de resolución por tipo de hallazgo.
   - Documentar acciones correctivas automatizadas (scripts, reglas) y los resultados de cada ejecución.
   - Revisar métricas de salud de las hojas y ajustar el plan según tendencias.

## Coordinación de exportaciones con revisiones de SLA

1. **Calendario operativo**
   - Sincronizar las exportaciones de Datos, Reportes y Resoluciones con el calendario de revisiones de SLA.
   - Definir hitos de preparación, revisión y publicación.
2. **Ejecución de exportaciones**
   - Automatizar la generación de archivos con parámetros de fecha y filtros aprobados.
   - Validar el contenido exportado contra los objetivos del SLA (completitud, puntualidad, exactitud).
3. **Integración con tableros operativos y auditorías**
   - Cargar los datos exportados directamente en los tableros internos sin depender de herramientas externas.
   - Mantener un historial versionado de cada exportación, incluyendo responsables y resultados de revisión.
4. **Retroalimentación continua**
   - Incorporar hallazgos de las auditorías y revisiones de SLA en los procesos de Importar datos y Diagnóstico.
   - Ajustar el calendario y los controles según la retroalimentación para garantizar consistencia.
