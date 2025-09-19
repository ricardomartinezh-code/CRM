# Checklist de despliegue operativo

Antes de iniciar un despliegue operativo, confirma los siguientes puntos:

## Verificación funcional (QA)
- [ ] Se ejecutaron las pruebas funcionales relevantes en el entorno de staging y todas pasaron.
- [ ] Se validaron manualmente los flujos críticos del usuario final.
- [ ] Se documentaron y comunicaron las incidencias conocidas pendientes del despliegue.

## Documentación
- [ ] Se actualizaron los manuales de usuario y procedimientos operativos que cambian con este despliegue.
- [ ] Se registraron las notas del lanzamiento en `docs/operacion/changelog.md`.
- [ ] Se adjuntaron instrucciones de reversión o mitigación en caso de incidentes.

## Comunicación interna
- [ ] Se notificó al equipo de soporte y a los stakeholders relevantes sobre la ventana de despliegue.
- [ ] Se confirmaron los responsables on-call para monitorear el despliegue.
- [ ] Se compartió el canal de comunicación a utilizar durante el despliegue y el post-mortem.
