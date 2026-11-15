type: minor

**Qué.** Un fragmento de changelog solo podía ser `type: patch` o `type: minor`, así que un cambio incompatible se publicaba como uno más. Así pasó con el authfile por defecto de `push-release.sh` en cadete-devops (#1435): el push de v1.3.2 falló sin que nadie avisara de que había que mover el fichero.
**Causa.** Ni el gate (`check_changelog_fragment.py`) ni el consolidador (`consolidate_changelog.py`) conocían un tipo para lo incompatible, ni pedían decir cómo migrar.
**Arreglo.** Los dos aceptan `type: major` y lo rechazan si no trae una línea `Migración: …` (o `Migration: …`). El consolidador sube la versión mayor y conserva la nota en el CHANGELOG. HR-44 fija la regla: cambiar una ruta o un valor por defecto de un script que ejecuta una persona es incompatible. Tests: los dos casos de `major` en cada asset, y el tipo inexistente pasa a ser `type: huge`.
**A partir de ahora.** Antes de escribir el fragmento se pregunta si alguien que ya usa la skill tiene que cambiar algo. Si la respuesta es sí, el fragmento es `type: major` y lleva su `Migración:`. Fase 0 de CI/CD, punto 0.2.
