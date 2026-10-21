type: patch

`host-contract.schema.json` no declaraba dos claves que el canon ya
implementa y documenta, así que un consumer que las declarara en su
`.github/host-contract.json` recibía `contract has unknown keys: [...]` y
`rc=2` del gate `check_host_drift` (#956): `governed_path`, la verificación
del camino de merge del detective post-push (#909), y `tracker_branch`, la
regla 7 que exige que lo que llega a la rama protegida venga por el PR del
tracker (#333). Es el mismo hueco que corrigió #946 con
`post_merge_dispatch`: la clave se implementa y se documenta, y el schema —
única fuente de verdad del set `known` desde #946— se queda atrás. Ahora el
schema declara ambas, con la forma que ya valida
`check_push_compliance.governed_path_config` y `tracker_branch()`.

La regresión queda fijada por un test que no hardcodea claves: recorre el AST
de los assets que consumen el host contract, recoge las claves de primer nivel
que se leen de él y exige que todas estén declaradas en el schema. Así la
próxima clave opt-in queda cubierta por el invariante y no por un test escrito
a mano para ella. El recorrido excluye las claves derivadas que el propio gate
escribe al normalizar (`runner_label`) y los dicts que no son el contrato
(respostas de la API, policy file), para no exigir al schema lo que el
consumer no declara.