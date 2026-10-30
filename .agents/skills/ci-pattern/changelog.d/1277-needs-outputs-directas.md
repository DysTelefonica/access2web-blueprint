type: patch

El gate `workflow-policy` mide, por job, que toda lectura de `needs.<job>.outputs`
venga de una dependencia **directa**: el contexto `needs` solo expone las outputs de
las directas, así que leer una transitiva devuelve vacío sin error y el workflow queda
verde con el valor equivocado. La excepción es el marcador `if:` de #1192, que lee un
job de primer nivel a propósito como filtro (#1277).
