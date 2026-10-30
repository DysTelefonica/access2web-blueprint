type: patch

Publica la forma canónica del job que mide la fase 3 dentro del pull request (HR-72, #1027), el segundo eslabón de la cadena: el primero (#1028) hizo exigible el requisito y este publica cómo se cumple.

El fragmento vive en el `references/porting-guide.md`, en el paso G3.5 de la fase 3, porque es lo que lee quien adopta; no en una copia dentro de un test. La suite `assets/templates/tests/test_governance_seals_job.py` lo extrae del documento y lo mide con los gates reales: pasa `check_workflow_policy.py` con la política de jobs requeridos de este repositorio, la fase 3 de `adoption check` lo acepta con el job cableado y sigue acepándolo sin declaración (fail-open), el nombre del check se publica solo desde `pull_request` (HR-30), la CLI se invoca como comando de primer nivel y no encadenada a otro gate (HR-46) y el paso no lleva `|| true` ni `continue-on-error`, que dejarían el job verde sin haber medido. La prosa y el gate no pueden divergir sin que se caiga una prueba.

El mismo paso nombra el destino de cada fase de la 0 a la 6 con el motivo de cada uno. Las que dependen de credenciales del host (la 0), miden checks ya publicados (la 4) o necesitan el PR fusionado (la 6) no son gates de PR por construcción y no se fuerzan; la 3 es la única cuyo rojo puede invertir un PR antes del merge, y por eso es la que se cablea.

Ningún gate existente cambia de alcance ni de nombre: este tramo solo publica la forma y la mide. La adopción en cada consumer sigue siendo un PR por consumer, porque en cuanto declara `governance_seals` el rojo aparece en el PR que declara.