type: patch

HR-7 pasa a AVISO en el gate de contrato de PR (#1172, decisión del dueño del 2026-10-24; auditoría de fricción de gates en team-skills#1093, comentario `6389037313`).

HR-7 es la palabra de cierre: el PR punta de una cadena lleva `Closes #<issue>` y los intermedios `Refs #<issue>` más la etiqueta de cadena. El gate lo juzgaba como hallazgo y tumbaba el PR; la auditoría lo clasificó como fricción (29 de los 170 job-runs rojos de la ventana de 48 h) porque el enlace a la issue lo sigue exigiendo `issue-spec` —regla del dueño— y la palabra de cierre se deduce de la cadena y de `closingIssuesReferences`. Ahora la punta sin `Closes`, el intermedio con `Closes`, el `chain:partial` sin `Refs` y el commit de un intermedio con `Closes` imprimen su aviso y NO cambian el exit code: el veredicto sigue siendo «PR-CONTRACT OK».

Lo que sigue bloqueando, igual que antes: HR-14 (commit convencional y sin atribución de IA) y el enlace a una issue de otro repositorio (HR-54). HR-53 (sección de evidencia del cierre) ya era aviso desde #1113 y este cambio no lo toca; su test lo fija.

Tests: la suite del gate (`assets/pr-contract/tests/test_pr_contract.py`) pasa los seis casos de HR-7 a aviso, conserva el caso de HR-14 bloqueando en el mismo PR que avisa de HR-7, y fija que HR-53 sigue avisando. Los casos de API viva que usaban el exit code para distinguir el cuerpo vivo del payload usan ahora el aviso, que es lo que sigue distinguiéndolos.
