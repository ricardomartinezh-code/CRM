# Checklist previo para importaciones de leads

Antes de enviar un archivo al panel ejecuta estas verificaciones para reducir rechazos y reprocesos:

- [ ] **Identificadores críticos completos.** Confirma que cada registro tenga al menos uno de los campos ID, Matrícula, Correo o Teléfono con un valor válido y sin espacios extra.
- [ ] **Duplicados internos eliminados.** Deduplica el archivo utilizando los identificadores anteriores para evitar omisiones por filas repetidas.
- [ ] **Normalización de formatos.** Ajusta teléfonos a un formato homogéneo (10 dígitos), correos en minúsculas y nombres capitalizados para facilitar búsquedas.
- [ ] **Datos requeridos presentes.** Verifica que etapa, estado, campus/modalidad y programa contengan valores compatibles con la base destino.
- [ ] **Comentarios y metadatos limpios.** Sustituye caracteres de control, saltos múltiples o texto irrelevante que pueda interferir con reportes.
- [ ] **Asignaciones verificadas.** Si defines asesores manualmente valida que el identificador exista y sea consistente con el equipo responsable.
- [ ] **Archivo sin filas vacías.** Elimina encabezados duplicados, totales y filas en blanco al final para evitar registros fantasma.
- [ ] **Revisión contra la base actual.** Cruza una muestra con la hoja de destino para confirmar que no existen colisiones inesperadas antes de importar todo el lote.

Mantén este checklist como parte del flujo operativo de datos para minimizar incidencias y aprovechar al máximo la validación automática del sistema.
