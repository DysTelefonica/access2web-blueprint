type: minor

`personal/ardelperal/ci-pattern/assets/pr-size/check_pr_size.py` se cablea como workflow dedicado en el catálogo `team-skills` (`.github/workflows/pr-size.yml`) con `pull_request.types = [opened, synchronize, reopened, labeled, unlabeled]` para que la etiqueta `size:exception` reevalúe el gate sin tocar la suite `unit`. El presupuesto sigue siendo 400 líneas, la regla de exclusión `VENDORED_ROOTS = (".agents/skills/", "skills/")` y la única excepción `size:exception` aplicada por el mantenedor: ni el gate ni el flujo del propagador cambian (#900, #623, #658).
