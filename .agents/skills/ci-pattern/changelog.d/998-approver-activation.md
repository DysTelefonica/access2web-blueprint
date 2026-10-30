type: minor

Team-skills declara sus aprobadores y la puerta dice quién aprobó (DysTelefonica/team-skills#998). Tres cosas que no son el núcleo de la regla y llegan aquí, en la punta de la cadena:

1. `.github/host-contract.json` declara `"approvers": ["ardelperal", "dys-orquestador[bot]"]`: el orquestador aprueba y mergea, el mantenedor aprueba, y la App de los agentes (`dys-agentes[bot]`) no puede aprobar su propio trabajo. Con la declaración, el gate `issue-spec` deja de aceptar cualquier identidad que ponga `status:approved`.
2. El motivo de una aprobación no autorizada deja de perderse en la vía informativa. Una `Refs` a una issue «aprobada» por un actor que no está declarado no cuenta como la aprobación que el PR necesita (#913), pero como una `Refs` es informativa, el motivo va como aviso por `stderr` y no como hallazgo: el PR no se bloquea por una referencia informativa y el operador ve quién puso la etiqueta sin tener autoridad en vez de un «no hay issue aprobada» sin explicación. Salió al ejecutar la puerta contra la API de verdad sobre #998.
3. `SKILL.md` documenta la regla en HR-54: qué se compara, con qué patrón, qué falla cerrado y qué queda exento.

Con esto la puerta queda verificada sobre la issue #998 con el contrato de este PR: `issue-spec: contrato satisfecho` (exit 0), y con `"approvers": ["ardelperal"]` la misma puerta sale con exit 1.