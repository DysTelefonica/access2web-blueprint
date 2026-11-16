type: minor

**Qué.** Dos reglas del perfil pre-release, con tests. Un PR que toca el pipeline nocturno necesita un ensayo nocturno en seco en verde para su SHA de cabeza (`assets/prerelease/night_gates.py dry-night`, que solo lee). El ejemplo de CI de la guía de porting añade un job `changes` con un filtro de SQL y migraciones: `upgrade-path` solo corre cuando el PR los cambia.
**Causa.** El ejemplo de CI ejecutaba `upgrade-path` (30 minutos y Postgres) en cada PR, también en uno sin SQL ni migraciones. Y nada exigía un ensayo nocturno antes de tocar el pipeline.
**Arreglo.** El ejemplo de la guía añade el job `changes` y la condición en `upgrade-path`. El gate nuevo sale con `0`, `1` o `2` (fallo cerrado si el registro no se puede leer).
**A partir de ahora.** Un consumer que copie el ejemplo de `upgrade-path` debe copiar también el job `changes`. Un PR que toque el pipeline nocturno necesita un ensayo en seco en verde de su SHA. Slice 8 de #1627.
