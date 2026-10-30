type: patch

El gate `issue-spec` en modo `pr-event` exime los PR de Dependabot del enlace a issue, igual que `check_branch_name` exime sus ramas (DysTelefonica/team-skills#1113, punto 1 E). La detección vive en un helper compartido (`assets/dependabot/dependabot.py`) y exige las dos cosas a la vez: que `pull_request.user.login` sea `dependabot[bot]` y que la rama tenga la forma de Dependabot. Un PR de otro autor con una rama `dependabot/…` sigue sin exención. La exención no lee el cuerpo ni consulta la API.
