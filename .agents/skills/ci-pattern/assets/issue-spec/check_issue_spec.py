#!/usr/bin/env python3
# ci-pattern asset — gate issue-spec (DysTelefonica/team-skills#633)
"""Gate de trazabilidad: el cuerpo del PR enlaza una issue aprobada y tipada.

Reglas (cada una es una función; la orquestación vive en validate_body y
validate_remote):

1. El cuerpo lleva `Closes|Fixes|Resolves|Refs #N` fuera de código.
2. La referencia resuelve a una issue, nunca a una pull request: la PR de la
   que se depende va en `depends-on` de Chain Context (HR-54). Si el cuerpo
   enlaza además otra issue válida, es un aviso (#787); sin ella, un fallo.
3. La issue lleva `status:approved` y exactamente un `type:*`. En un eslabón
   intermedio (`chain:partial`) basta el `type:*`: la aprobación se exige en
   la punta de cadena (#787). Si el consumer declara la lista `approvers` en
   su `.github/host-contract.json`, además se comprueba QUIÉN puso la
   etiqueta: el último evento `labeled` de la aprobación tiene que llevar un
   `actor.login` declarado (#998).
4. Retirada: `Closes #N` junto a `Refs #N` de la misma issue es válido y no
   produce hallazgo (HR-54, #787). El número se conserva para no renumerar las
   reglas citadas.
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
- `pr-event <event.json> [--required-sections <path>] [--host-contract <path>]`:
  todas las reglas. Usa GITHUB_TOKEN y GITHUB_API_URL del entorno si existen. Con
  `--required-sections`, la regla 7 se aplica al cuerpo de cada issue consultada. Con
  `--host-contract`, de la clave opcional `approvers` sale la lista de logins que pueden
  poner `status:approved` (#998).

Los avisos salen por stderr como `issue-spec: aviso: ...` y no cambian el
código de salida.

Exit codes: 0 sin hallazgos · 1 hallazgos · 2 entrada inválida.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

APPROVAL_LABEL = "status:approved"
CHAIN_PARTIAL_LABEL = "chain:partial"
TYPE_PREFIX = "type:"
# Techo de la línea de tiempo que se recorre para resolver QUIÉN puso la aprobación.
# Por encima, el veredicto sería sobre un fragmento del historial (falso verde).
EVENTS_PER_PAGE = 100
EVENTS_MAX_PAGES = 10
NO_EVENTS_READER = "no hay cliente de eventos: no se puede leer quién puso la aprobación"
NO_APPROVERS_WARNING = (
    "el host-contract del consumer no declara «approvers»: se acepta cualquier actor "
    "que ponga la etiqueta de aprobación (declare la lista para comprobar QUIÉN aprobó)")
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


class HostContractError(ValueError):
    """Raised when the host-contract is unreadable or `approvers` unservable (#998).

    Un contrato inservible no puede dejar la puerta en verde: ``main()`` devuelve
    ``exit 2``, el mismo criterio fail-closed que la regla 6 (HR-3).
    """


class HostContractUnreadableError(HostContractError):
    def __init__(self, path: str, exc: BaseException) -> None:
        super().__init__(f"no se puede leer {path}: {exc}")


class HostContractNotJsonError(HostContractError):
    def __init__(self, path: str, exc: BaseException) -> None:
        super().__init__(f"{path} no es JSON válido: {exc}")


class HostContractNotObjectError(HostContractError):
    def __init__(self, path: str) -> None:
        super().__init__(f"{path} debe ser un objeto JSON del contrato del host")


class HostContractApproversError(HostContractError):
    def __init__(self, path: str, detail: str) -> None:
        super().__init__(f"{path}: {detail}")


class ConfigUnreadableError(IssueSpecConfigError):
    def __init__(self, path: str, exc: BaseException) -> None:
        super().__init__(f"no se puede leer {path}: {exc}")


class ConfigNotJsonError(IssueSpecConfigError):
    def __init__(self, path: str, exc: BaseException) -> None:
        super().__init__(f"{path} no es JSON válido: {exc}")


class ConfigNotObjectError(IssueSpecConfigError):
    def __init__(self, path: str) -> None:
        super().__init__(f"{path} debe ser un objeto JSON con la clave 'required_sections'")


class ConfigSectionsNotListError(IssueSpecConfigError):
    def __init__(self, path: str) -> None:
        super().__init__(f"{path}: 'required_sections' debe ser una lista de strings no vacíos")


class ConfigSectionInvalidError(IssueSpecConfigError):
    def __init__(self, path: str, item: object) -> None:
        super().__init__(f"{path}: cada sección debe ser un string no vacío (recibido: {item!r})")


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


def extract_sections(body: str) -> dict[str, bool]:
    """Heading Markdown → True si la sección tiene contenido no vacío bajo el heading.

    Recorre el cuerpo fuera de código (vía ``strip_code``, mismo criterio
    que ``references``: dentro de un span de código o de un bloque cercado
    ni el heading ni el contenido cuentan). Para cada
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
        raise ConfigUnreadableError(path, exc) from exc
    try:
        data = json.loads(text)
    except json.JSONDecodeError as exc:
        raise ConfigNotJsonError(path, exc) from exc
    if not isinstance(data, dict):
        raise ConfigNotObjectError(path)
    sections = data.get("required_sections")
    if not isinstance(sections, list):
        raise ConfigSectionsNotListError(path)
    for item in sections:
        if not isinstance(item, str) or not item.strip():
            raise ConfigSectionInvalidError(path, item)
    return tuple(sections)


def load_approvers(path: str) -> tuple[str, ...]:
    """Lee la lista opcional `approvers` del host-contract del consumer (#998).

    Esquema: ``{"approvers": ["ardelperal", "dys-orquestador[bot]"]}``. Sin la clave se
    devuelve la tupla vacía: el gate conserva su comportamiento de siempre y lo advierte
    en una línea. Con la clave debe ser una lista de logins no vacíos; un contrato
    ilegible o inservible lanza ``HostContractError`` y sale ``exit 2``, porque una
    declaración rota que se ignorase dejaría la puerta en verde sin ella (HR-3).
    """
    try:
        text = Path(path).read_text(encoding="utf-8")
    except (OSError, UnicodeDecodeError) as exc:
        raise HostContractUnreadableError(path, exc) from exc
    try:
        data = json.loads(text)
    except json.JSONDecodeError as exc:
        raise HostContractNotJsonError(path, exc) from exc
    if not isinstance(data, dict):
        raise HostContractNotObjectError(path)
    approvers = data.get("approvers")
    if approvers is None:
        return ()
    if not isinstance(approvers, list) or not approvers:
        raise HostContractApproversError(path, "'approvers' debe ser una lista de logins no vacía")
    for item in approvers:
        if not isinstance(item, str) or not item.strip():
            raise HostContractApproversError(
                path, f"cada aprobador debe ser un login no vacío (recibido: {item!r})")
    return tuple(approvers)


def report_approvers(approvers: tuple[str, ...]) -> None:
    """Una línea por salida: cuántos aprobadores declara el contrato, o el aviso de que no declara ninguno (#998).

    El aviso es el punto 3 de la issue: sin `approvers` el veredicto no cambia, así que la
    puerta lo dice para que el consumer sepa que aún no comprueba QUIÉN aprobó.
    """
    if approvers:
        print(f"issue-spec: aprobadores declarados en el host-contract: {len(approvers)}")
        return
    print(f"issue-spec: aviso: {NO_APPROVERS_WARNING}", file=sys.stderr)


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


def validate_body(
    body: str,
    repository: str = "",
    required_sections: tuple[str, ...] = (),
) -> list[str]:
    """Modo local: forma del cuerpo, sin tocar la red.

    La regla 7 (secciones obligatorias) es opt-in: solo se aplica si
    ``required_sections`` no está vacío.
    """
    refs = references(body)
    violation = rule_has_reference(refs)
    if violation:
        return [violation]
    return rule_same_repository(refs, repository) + rule_required_sections(body, required_sections)


# --------------------------------------------------------------------------------------------
# Reglas de la issue remota: las que necesitan la API.
# --------------------------------------------------------------------------------------------


def rule_reference_is_issue(issue: Mapping[str, Any], number: int) -> str | None:
    """Regla 2: `GET /issues/{n}` devuelve `pull_request` cuando el número es una PR."""
    if "pull_request" not in issue:
        return None
    return (f"#{number}: la referencia resuelve a una pull request, no a una issue; "
            f"muévala a depends-on de Chain Context. {DEPENDS_ON_HINT}")


def _label_names(issue: Mapping[str, Any]) -> set[str]:
    """Conjunto de nombres de etiquetas de la issue. Helper privado."""
    return {
        label["name"] if isinstance(label, Mapping) else str(label)
        for label in issue.get("labels", [])
    }


def rule_missing_labels(
    issue: Mapping[str, Any], number: int, chain_partial: bool = False,
) -> list[str]:
    """Regla 3 (cierre): la issue lleva `status:approved` y un `type:*`.

    En un eslabón intermedio de cadena la aprobación se exige en la punta, no aquí.
    Se aplica solo a las referencias que cierran la issue (`Closes`/`Fixes`/
    `Resolves`); las `Refs` son informativas y las juzga ``validate_remote``
    sin aprobación ni tipo (DysTelefonica/team-skills#913).
    """
    names = _label_names(issue)
    violations: list[str] = []
    if APPROVAL_LABEL not in names and not chain_partial:
        violations.append(f"#{number}: requiere {APPROVAL_LABEL}")
    if not any(name.startswith(TYPE_PREFIX) for name in names):
        violations.append(f"#{number}: requiere una etiqueta {TYPE_PREFIX}*")
    return violations


NO_APPROVED_LINKED_MESSAGE = (
    "ninguna issue enlazada del propio repositorio está aprobada (status:approved "
    "y type:*): el PR debe enlazar al menos una. Las demás referencias son "
    "informativas y no se juzgan (DysTelefonica/team-skills#913)")


# --------------------------------------------------------------------------------------------
# Regla 3 (quién) — quién puso `status:approved` (DysTelefonica/team-skills#998).
# --------------------------------------------------------------------------------------------


def last_approval_actor(events: Sequence[Mapping[str, Any]]) -> str | None:
    """Login del ÚLTIMO evento `labeled` de la etiqueta de aprobación, o None.

    La API devuelve los eventos de más antiguo a más reciente, así que el último
    `labeled` es el que está en pie ahora; los `unlabeled` no cuentan (si la etiqueta se
    retiró, la regla 3 falla antes por su ausencia). Un evento sin `actor.login` legible
    devuelve None, que la regla traduce a «no consta».
    """
    actor: str | None = None
    for event in events:
        if event.get("event") != "labeled":
            continue
        label = event.get("label")
        if not isinstance(label, Mapping) or label.get("name") != APPROVAL_LABEL:
            continue
        who = event.get("actor")
        login = who.get("login") if isinstance(who, Mapping) else None
        actor = login if isinstance(login, str) and login else None
    return actor


def rule_approver(number: int, events: Sequence[Mapping[str, Any]],
                  approvers: Sequence[str]) -> str | None:
    """Regla 3 (cierre): la etiqueta de aprobación la puso un aprobador declarado.

    Sin `approvers` declarados no hay contra quién comparar y la regla no juzga
    (compatibilidad, #998 punto 3). Con la lista declarada el veredicto es fail-closed: o
    el último `labeled` lo puso alguien de la lista, o es un hallazgo que nombra al actor
    y la lista —también cuando el evento no consta o no dice quién lo puso.
    """
    actor = last_approval_actor(events)
    listed = ", ".join(approvers)
    if actor is None:
        return (f"#{number}: {APPROVAL_LABEL} está puesta pero no consta qué evento la puso; "
                f"los aprobadores declarados en el host-contract son: {listed}")
    if actor.casefold() not in {a.casefold() for a in approvers}:
        return (f"#{number}: {APPROVAL_LABEL} la puso «{actor}», que no está en los "
                f"aprobadores declarados en el host-contract: {listed}")
    return None


def rule_api_error(number: int, exc: GitHubApiError, what: str = "the issue") -> str:
    """Regla 6: un error de la API falla cerrado."""
    return f"#{number}: cannot read {what} ({exc})"


def is_chain_partial(payload: Mapping[str, Any]) -> bool:
    """El PR del evento lleva la etiqueta `chain:partial` (eslabón intermedio)."""
    labels = (payload.get("pull_request") or {}).get("labels") or []
    return any(isinstance(label, Mapping) and label.get("name") == CHAIN_PARTIAL_LABEL
               for label in labels)


@dataclass
class PrContext:
    """Contexto del PR que modula el veredicto remoto y recoge sus avisos (#787).

    ``chain_partial``: el PR es un eslabón intermedio (la aprobación se exige en la
    punta). ``warnings``: lo que no bloquea, para que el llamador lo imprima.
    ``approvers`` y ``events_fetcher`` son los logins que el consumer declara como
    aprobadores y cómo se lee la línea de tiempo de una issue (#998); sin ellos la
    regla 3 solo comprueba que la etiqueta esté puesta.
    """

    chain_partial: bool = False
    warnings: list[str] = field(default_factory=list)
    approvers: Sequence[str] = ()
    events_fetcher: Callable[[int], Sequence[Mapping[str, Any]]] | None = None


def _approver_judge(events_get: Callable[[int], Sequence[Mapping[str, Any]]] | None,
                    approvers: Sequence[str],
                    chain_partial: bool) -> Callable[[int], list[str]]:
    """La regla 3 (quién) como juez por número de issue (#998).

    No juzga si el PR es un eslabón `chain:partial` (punto 4: la aprobación se exige en la
    punta) ni si el consumer no declaró `approvers` (punto 3: compatibilidad, y el aviso ya
    lo dio `report_approvers`). Con la lista declarada, lee la línea de tiempo de la issue
    y falla cerrado ante cualquier duda: sin cliente de eventos, evento ausente, actor
    desconocido o API caída son todos hallazgo (HR-3).
    """
    def judge(number: int) -> list[str]:
        if chain_partial or not approvers:
            return []
        if events_get is None:
            return [rule_api_error(number, GitHubApiError(NO_EVENTS_READER), "the label events")]
        try:
            events = events_get(number)
        except GitHubApiError as exc:
            return [rule_api_error(number, exc, "the label events")]
        violation = rule_approver(number, events, approvers)
        return [violation] if violation else []
    return judge


def _judge_closing(
    numbers: list[int],
    get: Callable[[int], Mapping[str, Any]],
    chain_partial: bool,
    required_sections: tuple[str, ...],
    approver: Callable[[int], list[str]],
) -> tuple[list[str], list[str], bool]:
    """Reglas 3 y 7 sobre las referencias que cierran.

    Devuelve (violaciones, avisos de PR, hay alguna issue de cierre válida).
    """
    violations: list[str] = []
    pull_messages: list[str] = []
    valid = False
    for number in numbers:
        try:
            issue = get(number)
        except GitHubApiError as exc:
            violations.append(rule_api_error(number, exc))
            continue
        if message := rule_reference_is_issue(issue, number):
            pull_messages.append(message)
            continue
        valid = True
        violations.extend(rule_missing_labels(issue, number, chain_partial))
        # La aprobación se exige en la punta de cadena, y con ella QUIÉN la puso
        # (#998). Sin la etiqueta puesta o en un eslabón `chain:partial` no se
        # consulta la línea de tiempo: la ausencia ya la dicta la regla anterior.
        if not chain_partial and APPROVAL_LABEL in _label_names(issue):
            violations.extend(approver(number))
        violations.extend(rule_required_sections(
            issue.get("body", ""), required_sections, prefix=f"#{number}: "))
    return violations, pull_messages, valid


def _judge_informative(
    numbers: list[int],
    get: Callable[[int], Mapping[str, Any]],
    required_sections: tuple[str, ...],
    approver: Callable[[int], list[str]],
    context: PrContext,
) -> tuple[list[str], list[str], bool]:
    """Referencias informativas (#913): tipo y secciones, sin exigir aprobación.

    Devuelve (violaciones, avisos de PR, alguna enlazada aprobada y tipada).
    Mantiene la comprobación de tipo y de secciones de siempre para no relajar
    ningún gate existente; la aprobación solo aporta al PR la issue que exige.
    Una issue «aprobada» por un actor que no está en `approvers` no aporta esa aprobación
    (#998): el PR sigue sin la issue aprobada que exige la regla 3. Que el actor no esté
    declarado no es un hallazgo aquí —una `Refs` es informativa, no exige aprobación— pero
    sí un aviso, que es lo que le dice al operador que la etiqueta la puso alguien sin
    autoridad en vez de fallar en silencio.
    """
    violations: list[str] = []
    pull_messages: list[str] = []
    approved = False
    for number in numbers:
        try:
            issue = get(number)
        except GitHubApiError as exc:
            violations.append(rule_api_error(number, exc))
            continue
        if message := rule_reference_is_issue(issue, number):
            pull_messages.append(message)
            continue
        violations.extend(rule_missing_labels(issue, number, chain_partial=True))
        violations.extend(rule_required_sections(
            issue.get("body", ""), required_sections, prefix=f"#{number}: "))
        if rule_missing_labels(issue, number, chain_partial=False):
            continue
        approver_violations = approver(number)
        if approver_violations:
            context.warnings.extend(approver_violations)
        else:
            approved = True
    return violations, pull_messages, approved


def validate_remote(
    body: str,
    repository: str,
    fetch_issue: Callable[[int], Mapping[str, Any]],
    required_sections: tuple[str, ...] = (),
    context: PrContext | None = None,
) -> list[str]:
    """Modo pr-event: todas las reglas. Solo consulta las issues del propio repositorio.

    Regla por PR (#913): el cuerpo debe cerrar al menos una issue del propio
    repositorio (la que lleva `status:approved` y un `type:*`), o ser un
    eslabón intermedio (`chain:partial`, #787 S6). Las `Refs` son
    **informativas**: solo se comprueba que resuelvan a una issue y no a una
    PR; no exigen aprobación, ni tipo, ni las seis secciones canónicas. La
    regla 7 (secciones obligatorias) se aplica solo a las referencias de
    cierre; las informativas no la llevan.

    Advertencias y violaciones:
    - Una referencia a una PR es un aviso si el cuerpo cierra además una
      issue válida (#787); si no hay cierre válido, es bloqueo.

    Fail-closed (HR-3): un error de la API al consultar una issue de cierre
    o informativa es hallazgo, no verde silencioso.
    """
    refs = references(body)
    show_warnings = context is not None  # sin contexto nadie lee los avisos: se queda en fallo
    context = context or PrContext()
    # Regla 7 no se aplica al cuerpo del PR: validate_body se llama con la
    # lista vacía para que las reglas 1–5 sí corran y la 7 no.
    violations = validate_body(body, repository, ())
    if not refs:
        return violations

    same_repo_refs = tuple(
        ref for ref in refs
        if not ref.repo or ref.repo.lower() == repository.lower()
    )
    closing_refs = tuple(ref for ref in same_repo_refs if ref.kind == "close")
    informative_refs = tuple(ref for ref in same_repo_refs if ref.kind == "refs")

    # Una issue referenciada como `Closes` y como `Refs` se trae una sola vez;
    # la verificación de cierre tiene precedencia (es la que exige aprobación).
    fetched: dict[int, Mapping[str, Any]] = {}

    def _get(number: int) -> Mapping[str, Any]:
        if number in fetched:
            return fetched[number]
        issue = fetch_issue(number)
        fetched[number] = issue
        return issue

    # Regla 3 (aprobación + tipo) y 7 (secciones) a las que cierran (#913: las
    # Refs no exigen aprobación, ver _judge_informative). Un eslabón `chain:partial`
    # no exige aprobación en ningún punto, así que el juez de QUIÉN queda apagado
    # (#998 punto 4).
    approver = _approver_judge(context.events_fetcher, context.approvers,
                                context.chain_partial)
    closing_violations, closing_pulls, valid_closing_issue = _judge_closing(
        sorted({ref.number for ref in closing_refs}), _get,
        context.chain_partial, required_sections, approver)
    informative_violations, informative_pulls, approved_linked = _judge_informative(
        sorted({ref.number for ref in informative_refs}), _get, required_sections,
        approver, context)
    violations.extend(closing_violations)
    violations.extend(informative_violations)
    pull_messages = closing_pulls + informative_pulls

    # Sin ningún `Closes` que la regla 3 ya juzgue, el PR debe enlazar una issue
    # aprobada entre sus Refs. Un eslabón intermedio (`chain:partial`) no lo exige (#787 S6).
    if same_repo_refs and not closing_refs and not context.chain_partial and not approved_linked:
        violations.append(NO_APPROVED_LINKED_MESSAGE)

    # Una referencia a una PR es un aviso si el cuerpo cierra además una
    # issue válida (#787); si no hay cierre válido, es bloqueo.
    if valid_closing_issue and show_warnings:
        context.warnings.extend(pull_messages)
    else:
        violations.extend(pull_messages)
    return violations


# --------------------------------------------------------------------------------------------
# GitHub client: solo lectura. `issues: read` basta.
# --------------------------------------------------------------------------------------------


class GitHubIssueFailedError(GitHubApiError):
    def __init__(self, repository: str, number: int, exc: BaseException) -> None:
        super().__init__(f"GET /repos/{repository}/issues/{number} failed: {exc}")


class GitHubPullRequestFailedError(GitHubApiError):
    def __init__(self, repository: str, number: int, exc: BaseException) -> None:
        super().__init__(f"GET /repos/{repository}/pulls/{number} failed: {exc}")


class GitHubIssueEventsFailedError(GitHubApiError):
    def __init__(self, repository: str, number: int, exc: BaseException) -> None:
        super().__init__(f"GET /repos/{repository}/issues/{number}/events failed: {exc}")


class GitHubIssueEventsTruncatedError(GitHubApiError):
    def __init__(self, repository: str, number: int) -> None:
        super().__init__(
            f"GET /repos/{repository}/issues/{number}/events exceeds "
            f"{EVENTS_PER_PAGE * EVENTS_MAX_PAGES} events: la línea de tiempo es más "
            f"larga que el techo del gate y no se puede resolver quién puso la "
            f"última etiqueta")


class UnsupportedApiUrlError(GitHubApiError):
    def __init__(self, api_url: str) -> None:
        super().__init__(f"api url must be https (loopback http only): {api_url}")


def _open(api_url: str, path: str, headers: Mapping[str, str],
          query: Mapping[str, Any] | None = None):
    """Abre el GET con el esquema validado: https, o http solo en loopback.

    Cada rama lleva su literal dentro de la llamada: ruff S310 solo acepta
    una URL con esquema literal.
    """
    parts = urllib.parse.urlsplit(api_url)
    tail = parts.path.rstrip("/") + path
    if query:
        tail = f"{tail}?{urllib.parse.urlencode(query)}"
    if parts.scheme == "https":
        return urllib.request.urlopen(
            urllib.request.Request(f"https://{parts.netloc}{tail}", headers=dict(headers)),
            timeout=30)
    if parts.scheme == "http" and parts.hostname in ("127.0.0.1", "localhost"):
        return urllib.request.urlopen(
            urllib.request.Request(f"http://{parts.netloc}{tail}", headers=dict(headers)),
            timeout=30)
    raise UnsupportedApiUrlError(api_url)


class GitHubClient:
    def __init__(self, token: str = "", api_url: str = "https://api.github.com") -> None:
        self._token = token
        self._api_url = api_url.rstrip("/")

    def _headers(self) -> dict[str, str]:
        headers = {
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
            "User-Agent": "ci-pattern-issue-spec",
        }
        if self._token:
            headers["Authorization"] = f"Bearer {self._token}"
        return headers

    def issue(self, repository: str, number: int) -> Mapping[str, Any]:
        try:
            with _open(self._api_url, f"/repos/{repository}/issues/{number}",
                       self._headers()) as response:
                return json.load(response)
        except (urllib.error.HTTPError, urllib.error.URLError, TimeoutError) as exc:
            raise GitHubIssueFailedError(repository, number, exc) from exc

    def _events_page(self, repository: str, number: int,
                     page: int) -> list[Mapping[str, Any]]:
        """Una página de la línea de tiempo de la issue, de más antigua a más reciente."""
        try:
            with _open(self._api_url, f"/repos/{repository}/issues/{number}/events",
                       self._headers(), {"per_page": EVENTS_PER_PAGE, "page": page}) as response:
                data = json.load(response)
        except (urllib.error.HTTPError, urllib.error.URLError, TimeoutError) as exc:
            raise GitHubIssueEventsFailedError(repository, number, exc) from exc
        return data if isinstance(data, list) else []

    def issue_events(self, repository: str, number: int) -> list[Mapping[str, Any]]:
        """Toda la línea de tiempo de la issue, en orden cronológico (#998).

        La regla 3 (quién) juzga el ÚLTIMO evento `labeled` de la etiqueta de aprobación:
        mirar solo la primera página daría el veredicto sobre un fragmento del historial
        cuando la issue acumula más eventos de los que caben en ella, y ese fragmento
        puede traer como último a un aprobador viejo. Por eso se pagina hasta una página
        corta, y por encima del techo se falla cerrado en vez de dictar veredicto sobre lo
        leído (HR-3).
        """
        events: list[Mapping[str, Any]] = []
        for page in range(1, EVENTS_MAX_PAGES + 1):
            chunk = self._events_page(repository, number, page)
            events.extend(chunk)
            if len(chunk) < EVENTS_PER_PAGE:
                return events
        raise GitHubIssueEventsTruncatedError(repository, number)

    def pull_request(self, repository: str, number: int) -> Mapping[str, Any]:
        # El PR del mode `pr-event` se trae por API (#830): el snapshot del
        # payload queda obsoleto en cuanto el cuerpo se edita, y todo
        # "Re-run jobs" del UI reusa el mismo payload (el workaround de hoy
        # es un commit vacío para forzar `synchronize`, ardelperal/APAP_WEB#2163).
        try:
            with _open(self._api_url, f"/repos/{repository}/pulls/{number}",
                       self._headers()) as response:
                return json.load(response)
        except (urllib.error.HTTPError, urllib.error.URLError, TimeoutError) as exc:
            raise GitHubPullRequestFailedError(repository, number, exc) from exc


def _make_fetcher(repository: str) -> Callable[[int], Mapping[str, Any]]:
    """Fetcher desde `$GITHUB_TOKEN` y `$GITHUB_API_URL`. Los tests lo sustituyen."""
    client = GitHubClient(
        token=os.environ.get("GITHUB_TOKEN", ""),
        api_url=os.environ.get("GITHUB_API_URL", "https://api.github.com"),
    )
    return lambda number: client.issue(repository, number)


def _make_body_fetcher() -> Callable[[str, int], Mapping[str, Any]]:
    """Fetcher del cuerpo del PR desde `$GITHUB_TOKEN` y `$GITHUB_API_URL` (#830)."""
    client = GitHubClient(
        token=os.environ.get("GITHUB_TOKEN", ""),
        api_url=os.environ.get("GITHUB_API_URL", "https://api.github.com"),
    )
    return lambda repository, number: client.pull_request(repository, number)


def _make_events_fetcher(repository: str) -> Callable[[int], Sequence[Mapping[str, Any]]]:
    """Fetcher de la línea de tiempo de una issue (#998), desde el mismo cliente."""
    client = GitHubClient(
        token=os.environ.get("GITHUB_TOKEN", ""),
        api_url=os.environ.get("GITHUB_API_URL", "https://api.github.com"),
    )
    return lambda number: client.issue_events(repository, number)


def _read_body(path: str) -> str:
    if path == "-":
        return sys.stdin.read()
    return Path(path).read_text(encoding="utf-8")


def pr_event(  # noqa: PLR0913 — orquestador del modo remoto; cada kwarg 1:1 con una fuente.
    payload: Mapping[str, Any],
    *,
    body_fetcher: Callable[[str, int], Mapping[str, Any]] | None = None,
    issue_fetcher: Callable[[int], Mapping[str, Any]] | None = None,
    events_fetcher: Callable[[int], Sequence[Mapping[str, Any]]] | None = None,
    required_sections: tuple[str, ...] = (),
    approvers: Sequence[str] = (),
    warnings: list[str] | None = None,
) -> int:
    """Modo ``pr-event``: el cuerpo del PR se lee por API, nunca del payload (#830).

    El snapshot ``payload["pull_request"]["body"]`` queda obsoleto desde la primera
    edición del cuerpo y desde cualquier "Re-run jobs" del UI (porque reusa el
    mismo ``$GITHUB_EVENT_PATH``); el workaround de hoy es un commit vacío para
    forzar ``synchronize`` y un payload fresco — ardelperal/APAP_WEB#2163,
    #2164. La fuente autoritativa del cuerpo es la API: ``GET /repos/{owner}/{repo}/pulls/{n}``
    vía el ``GitHubClient`` del asset. Sin ``GITHUB_TOKEN`` o con la API caída,
    el gate falla cerrado con exit 2 y un motivo explícito (nunca fallback
    silencioso al cuerpo del payload, HR-3).

    ``body_fetcher``, ``issue_fetcher`` y ``events_fetcher`` son inyectables:
    ``main()`` los deja a ``None`` para usar el cliente real desde el entorno, y los tests
    los sustituyen por dobles.

    ``approvers`` son los logins que el consumer declara como aprobadores (#998); sin
    ellos la puerta conserva su veredicto de siempre, y `report_approvers` ya lo dijo.
    """
    if warnings is None:
        warnings = []
    repository = (payload.get("repository") or {}).get("full_name", "")
    if "/" not in repository:
        print("issue-spec: el evento no identifica repository.full_name", file=sys.stderr)
        return 2
    pull = payload.get("pull_request") or {}
    pr_number = pull.get("number")
    if not isinstance(pr_number, int) or pr_number <= 0:
        print("issue-spec: pr-event: el payload no identifica pull_request.number "
              "entero positivo; el cuerpo del PR solo se lee por API", file=sys.stderr)
        return 2

    if body_fetcher is None:
        token = os.environ.get("GITHUB_TOKEN", "")
        if not token:
            print("issue-spec: pr-event: GITHUB_TOKEN es obligatorio para leer el "
                  "cuerpo actual del PR; el cuerpo del payload queda obsoleto tras "
                  "ediciones del cuerpo o re-runs del job", file=sys.stderr)
            return 2
        try:
            pull_data = _make_body_fetcher()(repository, pr_number)
        except GitHubApiError as exc:
            print(f"issue-spec: pr-event: no se pudo leer el cuerpo actual del PR "
                  f"por API ({exc}); el cuerpo del payload queda obsoleto tras "
                  f"ediciones del cuerpo o re-runs del job", file=sys.stderr)
            return 2
        body = pull_data.get("body") or ""
    else:
        body = (body_fetcher(repository, pr_number) or {}).get("body") or ""

    fetcher = issue_fetcher if issue_fetcher is not None else _make_fetcher(repository)
    context = PrContext(chain_partial=is_chain_partial(payload), warnings=warnings,
                        approvers=approvers,
                        events_fetcher=events_fetcher or _make_events_fetcher(repository))
    violations = validate_remote(body, repository, fetcher, required_sections, context)
    success = "issue-spec: contrato satisfecho"
    for warning in warnings:
        print(f"issue-spec: aviso: {warning}", file=sys.stderr)
    for violation in violations:
        print(f"issue-spec: {violation}", file=sys.stderr)
    if violations:
        print(f"issue-spec: fallo con {len(violations)} hallazgo(s)", file=sys.stderr)
        return 1
    print(success)
    return 0


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)
    local = subparsers.add_parser("local", help="forma del cuerpo; sin payload, token ni red")
    local.add_argument("--body-file", required=True, help="cuerpo del PR; '-' lee stdin")
    local.add_argument("--repository", default="", help="owner/repo del PR, si se conoce")
    local.add_argument(
        "--required-sections", default="",
        help=("ruta al JSON con la lista de secciones obligatorias del consumer "
              "(DysTelefonica/team-skills#644); si se omite, la regla 7 no se aplica"))
    event = subparsers.add_parser("pr-event", help="payload de evento de PR, con la API")
    event.add_argument("event", help="ruta al payload JSON del evento")
    event.add_argument(
        "--required-sections", default="",
        help=("ruta al JSON con la lista de secciones obligatorias del consumer "
              "(DysTelefonica/team-skills#644); si se omite, la regla 7 no se aplica"))
    event.add_argument(
        "--host-contract", default="",
        help=("ruta al .github/host-contract.json del consumer; de él se lee la lista "
              "opcional `approvers` que declara QUIÉN puede poner status:approved "
              "(DysTelefonica/team-skills#998); si se omite, el gate acepta cualquier "
              "actor y lo advierte"))
    return parser


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: la salida no debe depender del locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    args = _build_parser().parse_args(argv)
    required: tuple[str, ...] = ()
    approvers: tuple[str, ...] = ()
    warnings: list[str] = []
    # Las dos politicas opcionales del consumer salen del mismo sitio y se niegan juntas:
    # una declaracion inservible es exit 2 (HR-3), no un veredicto sin la regla.
    try:
        if args.required_sections:
            required = load_required_sections(args.required_sections)
        if args.command == "pr-event":
            # El contrato solo se lee en el modo remoto: `local` no verifica la issue, y
            # por tanto no juzga quién la aprobó (#998).
            if args.host_contract:
                approvers = load_approvers(args.host_contract)
            report_approvers(approvers)
    except (IssueSpecConfigError, HostContractError) as exc:
        print(f"issue-spec: {exc}", file=sys.stderr)
        return 2
    if args.command == "local":
        try:
            body = _read_body(args.body_file)
        except (OSError, UnicodeDecodeError) as exc:
            print(f"issue-spec: no se puede leer el cuerpo: {exc}", file=sys.stderr)
            return 2
        violations = validate_body(body, args.repository, required)
        success = ("issue-spec local: forma de las referencias correcta. La aprobación, "
                   "el tipo y si la referencia es una PR los verifica el gate remoto.")
        for violation in violations:
            print(f"issue-spec: {violation}", file=sys.stderr)
        if violations:
            print(f"issue-spec: fallo con {len(violations)} hallazgo(s)", file=sys.stderr)
            return 1
        print(success)
        return 0
    try:
        payload = json.loads(Path(args.event).read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        print(f"issue-spec: payload ilegible: {exc}", file=sys.stderr)
        return 2
    return pr_event(payload, body_fetcher=None, issue_fetcher=None,
                     required_sections=required, approvers=approvers, warnings=warnings)


if __name__ == "__main__":
    sys.exit(main())
