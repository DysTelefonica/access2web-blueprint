type: minor

El gate `workflow-policy` exige ahora, en todo job, un grupo `concurrency` (declarado en el job o en el workflow), rechaza los servicios que fijan un puerto de host (`5432:5432`, `127.0.0.1:5432:5432`) y exige un `docker info` bajo `timeout` antes del primer `docker run`, en un paso anterior o antes en el mismo paso (un comentario no cuenta). No impone el valor de `cancel-in-progress`: eso es política del consumer. Los consumers cuyos workflows no cumplan verán hallazgos nuevos (#806).
