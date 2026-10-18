#!/usr/bin/env python3
# ci-pattern asset — gate issue-spec (DysTelefonica/team-skills#633)
"""Gate de trazabilidad: el cuerpo del PR enlaza una issue aprobada y tipada.

Reglas (cada una es una función; la orquestación vive en validate_body y
validate_remote):

1. El cuerpo lleva `Closes|Fixes|Resolves|Refs #N` fuera de código.
2. La referencia resuelve a una issue, nunca a una pull request: la PR de la
   que se depende va en `depends-on` de Chain Context (HR-54).
3. La issue lleva `status:approved` y exactamente un `type:*`.
4. El cuerpo no cierra y referencia la MISMA issue a la vez (HR-54).
5. La referencia apunta al repositorio del PR (HR-54).
6. Un fallo de la API no es un verde: fail-closed (HR-3).
7. El cuerpo (o el cuerpo de la issue citada) lleva las secciones declaradas
   en la configuración del consumer (DysTelefonica/team-skills#644). Opt-in:
   sin `--required-sections` la regla 7 no se aplica. Las secciones viven
   como títulos Markdown (H1..H6) y se ignoran dentro de un span de código o
   de un bloque cercado, igual que las closing keywords de la regla 1.

Qué cuenta como referencia sigue a GitHub: las palabras de cierre valen en
prosa, pero no dentro de un span de código ni de un bloque cercado. Un bloque
con sangría de 4 espacios no se quita: dentro de una lista es continuación del
ítem, y ocultarlo podría esconder una referencia real. Se cuenta, así que el
fallo es ruidoso y no silencioso. El mismo criterio aplica a la regla 7.

Modos:

- `local --body-file <ruta|-> [--repository owner/repo] [--required-sections <path>]`:
  forma del cuerpo. Sin payload, sin token y sin red. No verifica la issue
  (reglas 2 y 3). Con `--required-sections`, la regla 7 se aplica al cuerpo
  leído.
- `pr-event <event.json> [--required-sections <path>]`: todas las reglas. Usa
  GITHUB_TOKEN y GITHUB_API_URL del entorno si existen. Con
  `--required-sections`, la regla 7 se aplica al cuerpo de cada issue
  consultada.

Exit codes: 0 sin hallazgos · 1 hallazgos · 2 entrada inválida.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import urllib.error
import urllib.request
from collections.abc import Callable, Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import Any

APPROVAL_LABEL = "status:approved"
TYPE_PREFIX = "type:"
DEPENDS_ON_HINT = (
    "Refs/Closes solo apuntan a issues del mismo repositorio; la PR de la que "
    "depende va en depends-on de Chain Context")

# Palabras de cierre que GitHub reconoce, más la referencia no cerrante `Refs`.
CLOSE_KEYWORDS = "close|closes|closed|fix|fixes|fixed|resolve|resolves|resolved"
REFERENCE_RX = re.compile(
    rf"(?i)(?P<kind>\b(?:{CLOSE_KEYWORDS})|\brefs)\s+"
    r"(?P<repo>[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+)?#(?P<number>[0-9]+)\b")
FENCE_RX = re.compile(r"^ {0,3}(`{3,}|~{3,})")
PARAGRAPH_BREAK_RX = re.compile(r"\n\s*\n")
# Un span de código de n backticks se cierra con otro run de exactamente n.
CODE_SPAN_RX = re.compile(r"(?<![\\`])(`+)(?!`)(.+?)(?<!`)\1(?!`)", re.DOTALL)
CODE_PLACEHOLDER = "\x00"  # no es espacio: `Closes `x` #1` no forma una referencia
# ATX heading: 1–6 almohadillas, espacio, texto. La barra final de cierre de
# Markdown y el espacio en blanco alrededor no cuentan para la comparación.
HEADING_RX = re.compile(r"^\s{0,3}(#{1,6})\s+(.+?)\s*#*\s*$")


class GitHubApiError(RuntimeError):
    """Raised when GitHub cannot provide authoritative issue data."""


class IssueSpecConfigError(ValueError):
    """Raised when the required-sections config is unreadable or invalid (regla 7)."""


@dataclass(frozen=True)
class Reference:
    kind: str  # "close" | "refs"
    repo: str  # "" cuando la referencia no lleva owner/repo
    number: int


def strip_code(body: str) -> str:
    """Quita bloques cercados y spans de código: GitHub no cierra desde ahí."""
    kept: list[str] = []
    fence: str | None = None
    for line in body.splitlines():
        if fence is None:
            opened = FENCE_RX.match(line)
            if opened:
                fence = opened.group(1)
            else:
                kept.append(line)
            continue
        closing = line.strip()
        if closing and set(closing) == {fence[0]} and len(closing) >= len(fence):
            fence = None
    paragraphs = PARAGRAPH_BREAK_RX.split("\n".join(kept))
    return "\n\n".join(CODE_SPAN_RX.sub(CODE_PLACEHOLDER, p) for p in paragraphs)


def references(body: str) -> tuple[Reference, ...]:
    """Referencias a issues fuera de código, sin duplicados, en orden estable."""
    found = {
        Reference("refs" if match.group("kind").lower() == "refs" else "close",
                  match.group("repo") or "", int(match.group("number")))
        for match in REFERENCE_RX.finditer(strip_code(body or ""))
    }
    return tuple(sorted(found, key=lambda r: (r.number, r.kind, r.repo)))


# --------------------------------------------------------------------------------------------
# Reglas de forma: solo leen el cuerpo. Son las que corre el modo local.
# --------------------------------------------------------------------------------------------


def rule_has_reference(refs: tuple[Reference, ...]) -> str | None:
    """Regla 1: el cuerpo debe enlazar una issue fuera de código."""
    if refs:
        return None
    return ("el cuerpo no enlaza ninguna issue fuera de código: añada `Closes #N` "
            "(punta de cadena) o `Refs #N` (intermedio). Una PR de la que depende "
            "va en depends-on de Chain Context, no aquí")


# --------------------------------------------------------------------------------------------
# Regla 7 — secciones obligatorias (DysTelefonica/team-skills#644). Opt-in: la
# lista la declara el consumer; sin configuración, validate_body y
# validate_remote no la aplican.
# --------------------------------------------------------------------------------------------


def extract_headings(body: str) -> frozenset[str]:
    """Títulos Markdown (H1..H6) del cuerpo, fuera de spans de código y bloques cercados.

    Reutiliza ``strip_code`` para que el mismo criterio de la regla 1
    (dentro de un span o de un bloque cercado no cuenta) aplique a los
    títulos. La comparación es exacta y case-sensitive, como el contrato
    de formularios de cada consumer publica los nombres literales.
    """
    cleaned = strip_code(body or "")
    found: set[str] = set()
    for line in cleaned.splitlines():
        match = HEADING_RX.match(line)
        if match:
            found.add(match.group(2).strip())
    return frozenset(found)


def extract_sections(body: str) -> dict[str, bool]:
    """Heading Markdown → True si la sección tiene contenido no vacío bajo el heading.

    Recorre el cuerpo fuera de código (vía ``strip_code``, mismo criterio
    que ``extract_headings`` y ``references``: dentro de un span de código o
    de un bloque cercado ni el heading ni el contenido cuentan). Para cada
    heading encontrado registra un flag de contenido que pasa a True en
    cuanto aparece una línea no vacía entre este heading y el siguiente (o
    el final del cuerpo). El caso «heading presente, contenido ausente» es
    lo que la regla 7 llama «sección vacía».
    """
    cleaned = strip_code(body or "")
    sections: dict[str, bool] = {}
    current: str | None = None
    has_content = False
    for line in cleaned.splitlines():
        match = HEADING_RX.match(line)
        if match:
            if current is not None:
                sections[current] = has_content
            current = match.group(2).strip()
            has_content = False
        elif current is not None and line.strip():
            has_content = True
    if current is not None:
        sections[current] = has_content
    return sections


def load_required_sections(path: str) -> tuple[str, ...]:
    """Lee el fichero de configuración de secciones obligatorias. Falla cerrado.

    Esquema:

        {"required_sections": ["Sección A", "Sección B", ...]}

    La clave ``required_sections`` es obligatoria; su valor debe ser una lista
    de strings no vacíos. Cualquier desviación del esquema (fichero ilegible,
    JSON inválido, objeto ausente, clave ausente, valor no lista, elemento no
    string o string vacío) lanza ``IssueSpecConfigError`` y devuelve ``exit 2``
    desde ``main``. El criterio fail-closed es el mismo que la regla 6 (HR-3).
    """
    try:
        text = Path(path).read_text(encoding="utf-8")
    except (OSError, UnicodeDecodeError) as exc:
        raise IssueSpecConfigError(f"no se puede leer {path}: {exc}") from exc
    try:
        data = json.loads(text)
    except json.JSONDecodeError as exc:
        raise IssueSpecConfigError(f"{path} no es JSON válido: {exc}") from exc
    if not isinstance(data, dict):
        raise IssueSpecConfigError(
            f"{path} debe ser un objeto JSON con la clave 'required_sections'")
    sections = data.get("required_sections")
    if not isinstance(sections, list):
        raise IssueSpecConfigError(
            f"{path}: 'required_sections' debe ser una lista de strings no vacíos")
    for item in sections:
        if not isinstance(item, str) or not item.strip():
            raise IssueSpecConfigError(
                f"{path}: cada sección debe ser un string no vacío "
                f"(recibido: {item!r})")
    return tuple(sections)


def rule_required_sections(
    body: str,
    required_sections: tuple[str, ...],
    prefix: str = "",
) -> list[str]:
    """Regla 7: cada sección declarada debe aparecer como heading con contenido.

    Una sección se considera satisfecha cuando su heading está presente fuera
    de código y hay al menos una línea no vacía entre este heading y el
    siguiente (o el final del cuerpo). Si el heading falta, el hallazgo dice
    «falta la sección obligatoria»; si el heading está pero sin contenido,
    el hallazgo dice «sección obligatoria … está vacía». ``prefix`` antepone
    ``#<n>: `` en modo ``validate_remote``. Sin ``required_sections`` la regla
    es no-op (backward compat).
    """
    if not required_sections:
        return []
    sections = extract_sections(body)
    violations: list[str] = []
    for name in required_sections:
        if name not in sections:
            violations.append(f"{prefix}falta la sección obligatoria «{name}»")
        elif not sections[name]:
            violations.append(f"{prefix}sección obligatoria «{name}» está vacía")
    return violations


def rule_closes_and_refs_same_issue(refs: tuple[Reference, ...]) -> list[str]:
    """Regla 4: Closes y Refs del mismo número son una contradicción."""
    closed = {r.number for r in refs if r.kind == "close"}
    referred = {r.number for r in refs if r.kind == "refs"}
    return [
        f"#{number}: el cuerpo cierra y referencia la MISMA issue a la vez; use "
        "Closes o Refs, nunca los dos"
        for number in sorted(closed & referred)
    ]


def rule_same_repository(refs: tuple[Reference, ...], repository: str) -> list[str]:
    """Regla 5: el enlace apunta al repositorio del PR. Sin repositorio, no se juzga."""
    if not repository:
        return []
    return [
        f"{ref.repo}#{ref.number}: la referencia apunta a otro repositorio; el enlace "
        f"va al repositorio del PR ({repository})"
        for ref in refs
        if ref.repo and ref.repo.lower() != repository.lower()
    ]
