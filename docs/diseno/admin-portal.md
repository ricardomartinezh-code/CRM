# Portal administrativo · Lineamientos de UI

Este documento establece directrices para mantener consistencia visual y narrativa en el portal administrativo de ReLead EDU.

## Contenido e historias de datos

- **Contextualizar cada bloque:** acompaña los KPI con descripciones cortas que expliquen qué área del proceso afecta (contacto, integridad de la base, distribución de carga, etc.).
- **Actualizar la cadencia:** especifica en cada métrica la ventana de tiempo y la meta operativa (por ejemplo, "Últimas 24 h" o "Meta < 4 h"). Esto evita ambigüedad en la interpretación de tendencias.
- **Priorizar la acción:** resalta los indicadores que requieren reacción inmediata; agrega notas sobre el disparador de alertas (p. ej., `> meta`, `sin asignar`).
- **Agrupar por flujo:** organiza los KPI siguiendo el embudo operativo (ingreso de leads → depuración → asignación → contacto → resolución) para facilitar la lectura secuencial.
- **Evitar duplicar datos:** cuando un indicador se detalla en otra vista (reportes o tableros), enlaza el origen en vez de repetir tablas completas.

### Indicadores críticos

1. **SLA de contacto:** muestra el promedio del primer contacto tras la asignación. Incluye la meta vigente y resalta desviaciones mayores al 10 % del objetivo.
2. **Duplicados detectados:** expresa el porcentaje de registros con coincidencias activas en las últimas cargas. Añade contexto sobre el umbral de tolerancia y rutas para depuración.
3. **Balance de asignaciones:** calcula el porcentaje de leads distribuidos frente al total disponible. Explica cómo se reparte entre asesores y qué acciones se esperan cuando cae por debajo del 85 %.

## Iconografía y señalización

- Utiliza la librería `bootstrap-icons` ya integrada. Prioriza iconos lineales (`bi-stopwatch`, `bi-layers`, `bi-people`) para mantener uniformidad con el resto del portal.
- Coloca iconos dentro de contenedores circulares con el mismo radio que los KPI para reforzar jerarquía visual. Ajusta el `aria-hidden="true"` cuando sirvan solo de apoyo visual.
- Mantén coherencia semántica: un icono debe reforzar la categoría del dato (tiempo, integridad, colaboración) y no ser decorativo.
- Cuando agregues nuevos indicadores, define un ícono base y reutilízalo en tablas o tooltips para que la asociación sea inmediata.

## Pautas de accesibilidad

- **Contraste:** verifica que texto y fondo mantengan un contraste mínimo de 4.5:1 en ambos temas. Aprovecha las variables `--card` y `--panel-base-*` para lograrlo sin colores arbitrarios.
- **Roles y descripciones:** asigna `role="list"`/`role="listitem"` y atributos `aria-live` cuando las métricas se actualicen dinámicamente. Proporciona descripciones concisas en `metric-desc` para lectores de pantalla.
- **Navegación por teclado:** conserva el orden lógico en el DOM y evita tabindex mayores a `0`. Los botones o enlaces dentro de los KPI deben seguir el orden de lectura.
- **Estados comunicados por texto:** cuando exista una alerta o estado crítico, no dependas únicamente del color; agrega texto o iconografía que explique el cambio.
- **Compatibilidad con zoom:** diseña los contenedores para escalar correctamente hasta 200 % sin pérdida de información o colisiones.
