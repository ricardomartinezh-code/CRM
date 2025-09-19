# Landing ReLead EDU

## Wireframes propuestos

### Sección hero
- **Estructura**: layout en dos columnas; izquierda contenido textual y métricas en carrusel; derecha tarjeta de acceso (login).
- **Jerarquía**:
  1. Kicker con ícono para reforzar categoría.
  2. Titular principal (máximo 2 líneas) con tipografía Manrope 800.
  3. Subtítulo con énfasis en beneficio medible.
  4. Grupo de CTAs primario/secundario.
  5. Carrusel horizontal de métricas repetidas para continuidad.
- **Comportamiento**: carrusel automático con pausa al reducir movimientos (`prefers-reduced-motion`).

### Sección “Casos de éxito”
- **Layout**: bloque full-width con fondo suave (gradiente) y cards individuales.
- **Cards**: emplean variables `--radius` y `--shadow`, contenido con cita + autor + rol.
- **Contenido**: mínimo 3 testimonios simultáneos.

## Copy sugerido

### Hero
- **Kicker**: "Plataforma de reactivación".
- **Titular**: "Impulsa la retención con decisiones basadas en datos".
- **Subtítulo**: "ReLead EDU centraliza seguimiento, priorización y comunicación para que cada plantel recupere leads dormidos en menos tiempo.".
- **CTA primario**: "Solicitar demo guiada" → `mailto:hola@relead.edu?subject=Quiero%20una%20demo`.
- **CTA secundario**: "Contactar a ventas" → `mailto:ventas@relead.edu`.
- **Métricas**:
  - `+62%` leads reactivados (promedio 90 días).
  - `-38%` tiempo de respuesta (SLA de atención).
  - `x1.8` productividad de asesores (leads por jornada).
  - `12` integraciones activas (CRM y sistemas escolares).

### Casos de éxito
- **Título**: "Casos de éxito".
- **Descripción**: "Instituciones que combinan datos, automatización y seguimiento omnicanal con ReLead EDU para sostener su matrícula.".
- **Testimonios**:
  1. Universidad Horizonte — Dirección de Admisiones.
  2. Red Instituto Valle — Coordinación Comercial.
  3. Colegio Delta — Gerencia de Experiencia.

## Especificaciones responsive

| Breakpoint | Layout | Notas |
|------------|--------|-------|
| ≥1200px    | Hero en dos columnas con carrusel ocupando 60% del ancho disponible; tarjeta de login fija a la derecha; bloque de casos de éxito con 3 columnas. | Altura mínima 90vh; CTAs en fila. |
| 768px–1199px | Hero cambia a una columna, login debajo; carrusel conserva scroll automático y ancho completo; casos de éxito en 2 columnas. | CTAs se vuelven botones expandibles (flex-wrap). |
| ≤767px     | Pila vertical: kicker, título, subtítulo, CTAs, carrusel; tarjeta de login ocupa 100%; casos de éxito en 1 columna con padding reducido. | Carrusel permite scroll táctil y se desactiva animación si hay `prefers-reduced-motion`. |

