# Incidents — ejemplo destilado: pérdida de datos (Cadete, 2026-09-23)

Este doc es el patrón de ejemplo de la capa release/hotfix/post-mortem
(HR-19 a HR-22). Es una destilación, no el RCA completo: el post-mortem íntegro
vive en la ruta de post-mortems del consumer (parámetro 9 de
`assets/parameters.md`).

## Resumen del incidente

- **Fecha**: 2026-09-23. Consumer: Cadete (OpenShift/Quay).
- **Causa raíz (sistema, blameless)**: un `initContainer` con `rm -rf` sobre el
  PVC se re-ejecutó en un reapply del manifiesto y borró los datos.
- **Impacto**: pérdida de datos; la restauración la hizo IT a mano.
- **Brecha real detectada**: no existían backups automatizados; la recuperación
  dependía de intervención manual.

## Flujo aplicado (lo que el patrón exige)

1. Issue canónica (`type:bug`) con la evidencia del incidente.
2. Fix aterrizado en main por el pipeline normal: eliminación declarativa del
   `initContainer` en el YAML.
3. Deploy inmediato tras el merge.
4. Release hotfix `v1.0.0`: tag anotado, GitHub Release marcada `Latest` y
   notas concisas que enlazan al issue; el RCA nunca vivió dentro de las notas.
5. Post-mortem blameless en la ruta de post-mortems del consumer (parámetro 9)
   con secciones Timeline (UTC) / Impact / Root cause / What worked / What
   failed; causas de sistema, sin personas.
6. Action items abiertos como issues de GitHub con owner: automatización de
   backups de MySQL como brecha sistémica.

## Los cuatro niveles (HR-21, #322)

Un incidente deja cuatro niveles, enlazados por identificadores; el contenido de
los datos personales no entra en el repositorio.

| Nivel | Dónde vive | Qué lleva |
|---|---|---|
| Post-mortem | parámetro 9 (`docs/postmortems/<AAAA-MM-DD>-<slug>.md`) | Qué pasó, causas de sistema y action items; ID de incidente, tickets, SHA y nombres de sistemas |
| Registro operativo | parámetro 24 (por defecto `docs/incidents/<AAAA-MM-DD>-<slug>.md`, en este repositorio) | Detalle técnico para seguir los tickets abiertos: IPs, hostnames, comandos, logs, cronología fina y el siguiente paso de cada ticket |
| Secretos | parámetro 25 (gestor de secretos declarado) | Credenciales y tokens; se citan por su identificador en el gestor. Nunca en git |
| Datos personales y conversaciones con terceros | Sistema de tickets | Se referencian por número de ticket; nunca se copian al repositorio |

La evidencia se cita con su `sha256` y su ubicación de acceso controlado, y lo
redactado se marca de forma explícita (`[IP interna redactada]`), nunca se borra
en silencio. El gate de la fase 5 bloquea secretos, datos personales y evidencia
declarada sin hash ni ubicación; las IPs, los hostnames y los comandos internos
están permitidos (el repositorio es privado y lo ven las mismas personas).

## Lecciones que generaliza

- Un `initContainer` destructivo es código de producción, no bootstrap
  inofensivo: cualquier reapply puede volver a ejecutarlo.
- La brecha no fue el `rm -rf` aislado sino la ausencia de backups automatizados;
  el action item debe atacar la brecha sistémica, no el síntoma.
- Fix, deploy y release PATCH salieron el mismo día; el post-mortem y sus
  action items cerraron el ciclo con evidencia trazable al SHA desplegado.
