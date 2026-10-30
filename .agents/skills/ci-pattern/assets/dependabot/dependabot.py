"""Detección compartida de los PR y ramas de Dependabot (DysTelefonica/team-skills#1113).

Lo usan check_branch_name.py y check_issue_spec.py (modo pr-event). Un PR cuenta como
de Dependabot solo si lo abre ``dependabot[bot]`` **y** su rama tiene la forma de
Dependabot: el nombre de la rama por sí solo no prueba quién la empujó.
"""

import re

DEPENDABOT_ACTOR = "dependabot[bot]"
DEPENDABOT_BRANCH = re.compile(r"dependabot/(?:pip|uv|npm_and_yarn|github_actions)/[A-Za-z0-9._/-]+")


def is_dependabot_branch(ref: str | None) -> bool:
    """True cuando ``ref`` tiene la forma de una rama que crea Dependabot."""
    return DEPENDABOT_BRANCH.fullmatch(ref or "") is not None


def is_dependabot_pr(actor: str | None, ref: str | None) -> bool:
    """True cuando el PR lo abre Dependabot: actor de la plataforma y rama de Dependabot."""
    return actor == DEPENDABOT_ACTOR and is_dependabot_branch(ref)
