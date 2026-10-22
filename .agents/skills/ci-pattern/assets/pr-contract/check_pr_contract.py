#!/usr/bin/env python3
"""Gate de contrato de PR y commits (#193): HR-6/7/8/14/31. Sin API ni procesos.

El gate no ejecuta `git`: el workflow genera los datos del rango y se los pasa
por fichero (`--changed-files`, `--numstat-file`, `--commits-file`,
`--tracked-files`), de modo que un consumer lo copia byte a byte sin hallazgos
de S603/S607 (DysTelefonica/team-skills#753). Un fichero necesario que falta
sale 2 (HR-3)."""

import argparse
import fnmatch
import importlib.util
import json
import os
import re
import sys
from dataclasses import dataclass
from pathlib import Path

# Fuente única del presupuesto (#931): la regla A1 la comparte con pr-size.
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "budget"))
from vendored_budget import budget_lines, parse_origins  # noqa: E402

MIN_MERGE_PARENTS = 2


class EventFileMissingError(FileNotFoundError):
    def __init__(self, path):
        super().__init__(f"payload del evento ausente: {path}")


class PullRequestMissingError(TypeError):
    def __init__(self):
        super().__init__("el payload no contiene `pull_request`.")


class PolicyFileMissingError(FileNotFoundError):
    def __init__(self, path):
        super().__init__(f"política ausente: {path}")


class ChecksConfigError(ValueError):
    """`checks` declarado pero inválido: falla cerrado (exit 2)."""


class ChecksNotObjectError(ChecksConfigError, TypeError):
    def __init__(self):
        super().__init__("`checks` debe ser un objeto {comprobación: true|false}")


class UnknownChecksError(ChecksConfigError):
    def __init__(self, names, valid):
        super().__init__(
            f"`checks` nombra comprobaciones desconocidas: {', '.join(names)} "
            f"(válidas: {', '.join(valid)})")


class NonBooleanChecksError(ChecksConfigError, TypeError):
    def __init__(self, names):
        super().__init__(f"`checks` exige valores booleanos; no lo son: {', '.join(names)}")


class NoActiveChecksError(ChecksConfigError):
    def __init__(self):
        super().__init__("`checks` apaga todas las comprobaciones: ninguna activa")


class DocsEvidenceConfigError(ValueError):
    """`docs_evidence` declarado pero inválido: falla cerrado (exit 2)."""

    def __init__(self):
        super().__init__(
            "docs_evidence necesita 'pages' (lista de {page, covers}) y 'checklists' (objeto)")


# #938: el gate lee `labels` y `body` por API en vivo. Sin token o con la API
# caída, no se vuelve al payload del evento: se sale 2 (HR-3) con un motivo
# explícito. La llamada al API se delega a `runner-controls/github_rest.py`
# (cliente REST compartido del catálogo), que se carga por ruta con el
# mismo patrón que `check_docs_architecture` (consumers con pin y
# vendorizados lo encuentran igual).

OBSOLETE_NOTE = "payload obsoleto tras ediciones/re-runs del job"
_PAYLOAD_INCOMPLETE_MSG = "payload sin owner/repo o número"
_LABELS_NOT_LIST_MSG = "labels no es lista"
_PULL_NOT_OBJECT_MSG = "pull no es objeto"


def _load_github_rest():
    """Carga `runner-controls/github_rest.py` por ruta. Fail-closed (#938)."""
    rest_path = Path(__file__).resolve().parent.parent / "runner-controls" / "github_rest.py"
    if not rest_path.is_file():
        return None
    spec = importlib.util.spec_from_file_location(
        "pr_contract_github_rest", str(rest_path))
    if spec is None or spec.loader is None:
        return None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


_GITHUB_REST = _load_github_rest()
if _GITHUB_REST is None:
    print("ERROR: runner-controls/github_rest.py no esta en el catalogo "
          "(HR-3 fail-closed).", file=sys.stderr)
    sys.exit(2)


def load_event(path):
    p = Path(path)
    if not p.is_file():
        raise EventFileMissingError(path)
    data = json.loads(p.read_text(encoding="utf-8"))
    pr = data.get("pull_request")
    if not isinstance(pr, dict):
        raise PullRequestMissingError
    return data, pr

def load_policy(path):
    p = Path(path)
    if not p.is_file():
        raise PolicyFileMissingError(path)
    return json.loads(p.read_text(encoding="utf-8"))


def fetch_pr_live(data, pr):
    """Lee labels y body actuales por API; fail-closed si no se puede.

    Delega a `runner-controls/github_rest.py` (cliente REST compartido del
    catálogo): valida el esquema, lee el token, paginata, lanza
    `GitHubApiError` en cualquier fallo. Las plantillas de error viven
    como constantes de módulo para no disparar TRY003 en el call site."""
    repo_full_name = (data.get("repository") or {}).get("full_name")
    number = pr.get("number")
    if not isinstance(repo_full_name, str) or not repo_full_name or \
            not isinstance(number, int):
        raise _GITHUB_REST.GitHubApiError(_PAYLOAD_INCOMPLETE_MSG)
    live_labels = _GITHUB_REST.get_json(
        f"/repos/{repo_full_name}/issues/{number}/labels")
    pull = _GITHUB_REST.get_json(
        f"/repos/{repo_full_name}/pulls/{number}")
    if not isinstance(live_labels, list):
        raise _GITHUB_REST.GitHubApiError(_LABELS_NOT_LIST_MSG)
    if not isinstance(pull, dict):
        raise _GITHUB_REST.GitHubApiError(_PULL_NOT_OBJECT_MSG)
    labels = [{"name": str(item.get("name", ""))} for item in live_labels
              if isinstance(item, dict)]
    return labels, pull.get("body") or ""


def _same_file(a, b):
    """`a` y `b` nombran el mismo fichero (rutas resueltas)."""
    try:
        return Path(a).resolve() == Path(b).resolve()
    except (OSError, RuntimeError):
        return False


def use_live_api(event_file, pr, env):
    """#938/#996: ¿juzgar etiquetas y cuerpo vivos por API en vez de los del payload?

    El payload de un `pull_request` queda congelado al abrir el PR: un PR que se etiqueta
    después falla HR-6 aunque la etiqueta ya esté puesta cuando el job corre. Por eso:

    - `PR_CONTRACT_USE_LIVE_API=0`: opt-out explícito, siempre el payload.
    - `PR_CONTRACT_USE_LIVE_API=1`: siempre por API (sin token es fallo cerrado).
    - Sin la variable: por API si se juzga el evento real de Actions (`--event-file` es
      `GITHUB_EVENT_PATH`), el payload es de un `pull_request` y hay token. Los scripts
      locales, el propagador y la suite de parity pasan su propio fichero de evento y no
      cambian. Con la API caída o rechazada el gate falla cerrado (exit 2).
    """
    flag = env.get("PR_CONTRACT_USE_LIVE_API")
    if flag == "0":
        return False
    if flag == "1":
        return True
    actions_event = env.get("GITHUB_EVENT_PATH")
    has_token = bool(env.get("GITHUB_TOKEN") or env.get("GH_TOKEN"))
    return bool(actions_event and has_token and pr and _same_file(event_file, actions_event))


def read_data_file(path):
    """Texto de un fichero de datos del rango; ausente o ilegible: FileNotFoundError."""
    return Path(path).read_text(encoding="utf-8")


CHECK_NAMES = ("labels", "closes", "chain_context", "issue_links", "closure_tests",
               "docs_evidence", "budget", "commits", "ai_disclosure")


def active_checks(policy):
    """#735: comprobaciones activas. Sin `checks` lo están todas; con `checks`,
    una ausente está activa y `false` la apaga. Valor inválido o ninguna
    activa: ChecksConfigError (HR-3, un gate sin nada que comprobar no da OK)."""
    config = policy.get("checks")
    if config is None:
        return set(CHECK_NAMES)
    if not isinstance(config, dict):
        raise ChecksNotObjectError
    unknown = sorted(set(config) - set(CHECK_NAMES))
    if unknown:
        raise UnknownChecksError(unknown, CHECK_NAMES)
    bad = sorted(name for name, value in config.items() if not isinstance(value, bool))
    if bad:
        raise NonBooleanChecksError(bad)
    active = {name for name in CHECK_NAMES if config.get(name, True)}
    if not active:
        raise NoActiveChecksError
    return active


def _anchor_exists(text, anchor):
    slugs = {re.sub(r"[^\w\- ]", "", m.group(1)).strip().lower().replace(" ", "-")
             for m in re.finditer(r"^#{1,6}\s+(.*?)\s*#*$", text, re.M)}
    return anchor.lower() in slugs


def check_docs_evidence(pr, repo, files, policy, findings):
    """#710 (HR-64): un PR que toca codigo cubierto por una pagina del mapa toca
    la pagina o lleva la seccion «Documentacion» no vacia (`files`: ficheros del
    rango, uno por elemento); las checklists por
    tipo de cambio existen y el cuerpo enlaza la de su tipo. Activacion por
    DECLARACION `docs_evidence` en la politica: sin ella avisa y no falla."""
    config = policy.get("docs_evidence")
    if config is None:
        print("HR-64: docs_evidence no declarado en la politica; la evidencia de "
              "documentacion por PR todavia no se exige.", file=sys.stderr)
        return
    pages = config.get("pages") if isinstance(config, dict) else None
    checklists = config.get("checklists", {}) if isinstance(config, dict) else None
    if not isinstance(pages, list) or not pages or not isinstance(checklists, dict) \
            or not all(isinstance(p, dict) and p.get("page") and p.get("covers") for p in pages):
        raise DocsEvidenceConfigError
    reason = re.sub(r"<!--.*?-->", "", section_body(pr.get("body"), "Documentación"), flags=re.S).strip()
    for entry in pages:
        touched = [f for f in files if any(fnmatch.fnmatch(f, g) for g in entry["covers"])]
        if touched and entry["page"] not in files and not reason:
            findings.append(
                f"HR-64: el PR toca {', '.join(touched)}, cubierto por {entry['page']}, y ni "
                "cambia la pagina ni lleva la seccion «Documentación» con el motivo.")
    body = pr.get("body") or ""
    types = {label["name"] for label in pr.get("labels", [])}
    for label, link in sorted(checklists.items()):
        path, _, anchor = link.partition("#")
        target = Path(repo) / path
        if not target.is_file() or (anchor and not _anchor_exists(target.read_text(encoding="utf-8"), anchor)):
            findings.append(f"HR-64: la checklist de {label} apunta a {link} y no existe.")
        elif label in types and link not in body:
            findings.append(f"HR-64: el cuerpo del PR ({label}) no enlaza su checklist {link}.")


def check_labels(pr, policy, findings):
    required = policy.get("required_labels", [])
    labels = {item["name"] for item in pr.get("labels", [])}
    if not required:
        return
    if not labels:
        findings.append("HR-6: el PR no tiene etiquetas; se requiere al menos una de " + str(required) + ".")
        return
    matching = labels & set(required)
    if not matching:
        findings.append(
            f"HR-6: falta etiqueta obligatoria; se requiere una de {required}, "
            f"el PR tiene {sorted(labels)}."
        )

def check_closes(pr, base, default_branch, policy, findings):
    body = pr.get("body") or ""
    pattern = policy.get("closes_pattern", "Closes|Fixes|Resolves")
    requires = policy.get("chain_tip_requires_closes", True)
    chain_label = policy.get("chain_partial_label", "chain:partial")
    labels = {item["name"] for item in pr.get("labels", [])}
    rx = re.compile(rf"\b(?:{pattern})\s+#\d+", re.IGNORECASE)
    refs_rx = re.compile(r"\bRefs\s+#\d+", re.IGNORECASE)
    is_tip = base == default_branch
    has_closes = bool(rx.search(body))
    has_refs = bool(refs_rx.search(body))
    is_partial = chain_label in labels

    if is_partial:
        # Cadena apilada: Refs obligatorio, Closes prohibido.
        if not has_refs:
            findings.append(
                "HR-7: el PR con etiqueta chain:partial debe llevar 'Refs #<issue>' "
                "en el cuerpo (la issue la cierra el PR punta de cadena)."
            )
        if has_closes:
            findings.append(
                "HR-7: el PR con etiqueta chain:partial no debe llevar "
                "'Closes #<issue>'; solo el PR punta de cadena la cierra."
            )
    elif is_tip and requires and not has_closes:
        findings.append(
            "HR-7: el PR punta de cadena debe llevar 'Closes #<issue>' en el cuerpo."
        )
    elif not is_tip and has_closes:
        findings.append(
            "HR-7: solo el PR punta de cadena lleva 'Closes #<issue>'; "
            "los PR encadenados intermedios no deben cerrar la issue."
        )

def size_exception_applied(pr, policy):
    """A1: la excepción de presupuesto la aplica solo el mantenedor cuando
    partir el PR rompería `main`. El gate verifica que la etiqueta
    `size:exception` está aplicada; la procedencia (mantenedor y motivo)
    vive en el cuerpo del PR bajo `## Excepción de tamaño` y la aplica un
    humano en GitHub, no este gate.

    Devuelve True si la etiqueta está aplicada, False en caso contrario.
    """
    label_name = policy.get("size_exception_label", "size:exception")
    labels = {item.get("name") for item in pr.get("labels", []) if isinstance(item, dict)}
    return label_name in labels


def check_budget(pr, diff_lines, policy, findings):
    """A1: el presupuesto es 400 (adiciones + eliminaciones del PR, sin
    exclusiones). Si el diff lo supera, la única vía válida es la
    etiqueta `size:exception` aplicada por el mantenedor; el campo
    `size-exception-reason:` del cuerpo está retirado."""
    budget = policy.get("review_budget_lines", 400)
    label_name = policy.get("size_exception_label", "size:exception")
    if diff_lines <= budget:
        return
    if not size_exception_applied(pr, policy):
        findings.append(
            f"A1: el diff supera el presupuesto ({diff_lines} > {budget}) y "
            f"no lleva la etiqueta '{label_name}' aplicada por el mantenedor. "
            "Encadene PRs o pida al mantenedor la etiqueta tras responder las "
            "4 preguntas de '## Excepción de tamaño' en el cuerpo."
        )

# #944: los patrones de atribución de IA se aplicaban con `re.search(pat, msg,
# re.IGNORECASE)`, sin límite de palabra ni ancla, así que una palabra que
# CONTIENE el patrón disparaba el hallazgo: el marcador de «generado con» casa
# dentro de «regenerated with» (la redacción natural de un commit que describe
# una regeneración) y el guion hace lo mismo con el de co-autoría tras un `No-`
# o un `hand-`. El límite va en el PATRÓN, compilado una vez, y no en cada
# cadena: así la lista que declara cada consumer queda cubierta sin tocar su
# política. Se exige que la marca NO esté pegada a un carácter de palabra ni a
# un guion; el resto (inicio de línea, espacio, puntuación, emoji previo) sigue
# dando hallazgo, que es lo que cazan los trailers reales. Solo límite izquierdo:
# los tres patrones por defecto acaban en `:` o en letra, y un `\b` final
# dejaría de cazarlos. Los tres marcadores se listan en
# `pr-contract.policy.example.json`, no aquí: citarlos en el código los arrastra
# a los mensajes de commit, que es lo que este gate mide.
AI_ATTR_BOUNDARY = r"(?<![\w-])"
INLINE_FLAGS_RX = re.compile(r"^\(\?[aiLmsux]+\)")

def compile_ai_attribution(patterns):
    """Compila los patrones de `ai_attribution_patterns` exigiendo el límite
    izquierdo de #944. El cuerpo va en un grupo no capturante para que una
    alternancia del patrón (`a|b`) reciba el límite en las dos ramas, y un
    prefijo de flags en línea (`(?i)…`) se conserva delante porque Python exige
    los flags globales al principio de la expresión."""
    compiled = []
    for pat in patterns:
        m = INLINE_FLAGS_RX.match(pat)
        head, body = (pat[:m.end()], pat[m.end():]) if m else ("", pat)
        compiled.append(re.compile(f"{head}{AI_ATTR_BOUNDARY}(?:{body})", re.IGNORECASE))
    return compiled

def check_commits(messages, policy, findings):
    pattern = policy.get("conventional_commit_pattern",
                         r"^(feat|fix|docs|chore|ci|test|perf|refactor|revert)(\([^)]+\))?!?: ")
    ai_patterns = compile_ai_attribution(
        policy.get("ai_attribution_patterns",
                   ["Co-Authored-By:", "Generated-by:", "Generated with"]))
    # %P delante: un commit de merge no tiene asunto conventional —la regla
    # mide los commits de trabajo— y la disciplina de cadenas obliga a
    # sincronizar con `git merge origin/main` en vez de rebase (#333).
    for block in messages.split("---COMMIT-END---"):
        msg = block.strip()
        if not msg:
            continue
        parents, _, message = msg.partition("\n")
        subject = message.split("\n", 1)[0].strip()
        if len(parents.split()) < MIN_MERGE_PARENTS and not re.match(pattern, subject):
            findings.append(f"HR-14: commit no convencional: '{subject[:80]}'.")
        for pat in ai_patterns:
            if pat.search(msg):
                findings.append(f"HR-14: atribución de IA detectada en commit '{subject[:80]}'.")
                break

CLOSES_RX = re.compile(r"\b(?:Closes|Fixes|Resolves)\s+#\d+", re.IGNORECASE)
# #334 (HR-53): el cierre se demuestra con una LISTA ESTRUCTURADA de entradas
# `ruta/al/fichero_de_test.ext` o `ruta::nombre_de_caso`. El formato libre
# dejaba pasar prosa —«test assert» da el token `assert`, presente en
# cualquier fichero de test— y a la vez producía falsos «no existe» con «test
# cubre». La ruta debe estar versionada (git ls-files) y el caso, DEFINIDO en
# ese fichero, no solo aparecer como subcadena.
CLOSURE_ENTRY_RX = re.compile(
    r"`([A-Za-z0-9_][A-Za-z0-9_./-]*\.[A-Za-z0-9]+(?:::[A-Za-z_][A-Za-z0-9_]*)?)`")
TEST_FILE_RX = re.compile(
    r"(?:^|/)(?:tests?|__tests__|spec)/|(?:^|/)test_[^/]+$|_test\.[a-z]+$", re.IGNORECASE)
# Definiciones admitidas del caso, por ecosistema. `\b` y el nombre exacto: la
# subcadena no vale.
CASE_DEF_PATTERNS = (
    r"\bdef\s+{name}\b",                       # Python
    r"\bfunction\s+{name}\b",                  # JS/PHP
    r"\bfunc\s+{name}\b",                      # Go
    r"\b(?:it|test)\s*\(\s*['\"]{name}['\"]",  # JS/TS (Jest, Mocha)
)
DEFAULT_CLOSURE_HEADING = "Tests que prueban el cierre"
# #911: la forma «no aplica» — un cierre SIN test declara, con un encabezado
# estable, la evidencia que cubre el cierre. La evidencia es una ruta
# versionada en git (un documento, un ADR, un registro); no tiene que ser un
# fichero de test. Sin exención por etiqueta: la etiqueta la pone el autor del
# PR y no es una evidencia.
DEFAULT_CLOSURE_NOT_APPLICABLE_HEADING = "No aplica"
# #918: «No aplica» solo vale si TODOS los ficheros cambiados no son ejecutables.
# La política puede declarar `closure_not_applicable_paths`; un workflow o un
# script no entra por defecto. Patrones con `**` (cualquier profundidad) y `*`.
DEFAULT_NOT_APPLICABLE_PATHS = (
    "**/*.md", "docs/**", ".github/ISSUE_TEMPLATE/**", ".github/PULL_REQUEST_TEMPLATE.md")


@dataclass(frozen=True)
class ClosureData:
    """Lo que la comprobación de cierre necesita del repo y del rango (#918).

    `changed` es None si no llegó `--changed-files`: «No aplica» falla cerrado."""
    repo: Path
    tracked: frozenset
    changed: tuple | None
    not_applicable_paths: tuple


def section_body(body, heading):
    """Cuerpo de una sección del PR, por su encabezado (#334).

    Devuelve el texto entre el encabezado ``#... <heading>`` y el siguiente
    encabezado de nivel igual o superior, o ``""`` si la sección no existe.
    El encabezado se compara sin distinguir mayúsculas y sin espacios de
    sobra. Este es el parser COMPARTIDO de secciones del cuerpo: la
    disciplina de cadenas (#333) lo reutiliza para su bloque Chain Context,
    de modo que no haya dos parsers que puedan divergir.
    """
    if not isinstance(body, str) or not body.strip():
        return ""
    target = re.compile(r"^(#{1,6})\s*" + re.escape(heading.strip()) + r"\s*$",
                        re.IGNORECASE)
    lines = body.splitlines()
    start = level = None
    for index, line in enumerate(lines):
        matched = target.match(line.strip())
        if matched:
            start, level = index + 1, len(matched.group(1))
            break
    if start is None:
        return ""
    out = []
    for line in lines[start:]:
        following = re.match(r"^(#{1,6})\s+\S", line.strip())
        if following and len(following.group(1)) <= level:
            break
        out.append(line)
    return "\n".join(out).strip()


def closure_entries(text):
    """Entradas estructuradas de la sección: `ruta` o `ruta::caso`."""
    return CLOSURE_ENTRY_RX.findall(text)


def case_is_defined(repo, rel, case):
    """El caso está DEFINIDO en el fichero (no basta con que la cadena
    aparezca). Se lee solo ese fichero: nada de recorrer el árbol."""
    try:
        text = (Path(repo) / rel).read_text(encoding="utf-8", errors="ignore")
    except OSError:
        return False
    return any(re.search(pattern.format(name=re.escape(case)), text)
               for pattern in CASE_DEF_PATTERNS)


def glob_matches(rel, pattern):
    """Ruta contra un patrón: `**/` cualquier profundidad, `**` cualquier resto, `*` sin barra."""
    regex, i = "", 0
    while i < len(pattern):
        if pattern.startswith("**/", i):
            regex, i = regex + "(?:.*/)?", i + 3
        elif pattern.startswith("**", i):
            regex, i = regex + ".*", i + 2
        elif pattern[i] == "*":
            regex, i = regex + "[^/]*", i + 1
        else:
            regex, i = regex + re.escape(pattern[i]), i + 1
    return re.fullmatch(regex, rel) is not None


def not_applicable_scope_findings(data, heading, marker):
    """#918: «No aplica» solo vale si el PR no cambia nada ejecutable. Sin lista
    de cambios no se puede comprobar: falla cerrado."""
    if data.changed is None:
        return [f"HR-53: la sección '{heading}' declara '{marker}' y el gate no recibió "
                "la lista de ficheros cambiados (--changed-files): sin ella no puede "
                "comprobar que el cambio no es ejecutable."]
    if not data.changed:
        return [f"HR-53: la sección '{heading}' declara '{marker}' y el PR no cambia "
                "ningún fichero."]
    executable = [rel for rel in data.changed
                  if not any(glob_matches(rel, pat) for pat in data.not_applicable_paths)]
    if not executable:
        return []
    listed = ", ".join(f"`{rel}`" for rel in executable)
    return [f"HR-53: la sección '{heading}' declara '{marker}' y el PR cambia ficheros "
            f"ejecutables ({listed}): «{marker}» solo vale para documentación. Añada un "
            "test que pruebe el cambio y nómbrelo en la sección."]


def not_applicable_findings(data, heading, marker, evidence_text):
    """#911: la sección HR-53 declara «No aplica» y nombra la evidencia del
    cierre. La evidencia es una ruta versionada en git; no tiene que ser un
    fichero de test. Sin evidencia, la declaración no cierra nada. #918: y el
    cambio no puede ser ejecutable."""
    scope = not_applicable_scope_findings(data, heading, marker)
    if scope:
        return scope
    repo, tracked = data.repo, data.tracked
    entries = closure_entries(evidence_text)
    if not entries:
        return [
            f"HR-53: la sección '{heading}' declara '{marker}' y no nombra la "
            "evidencia del cierre: escriba la ruta del fichero que lo cubre, en el "
            "formato `ruta/al/fichero.ext`."]
    problems = []
    for entry in entries:
        rel, _, case = entry.partition("::")
        if rel not in tracked:
            problems.append(f"`{entry}` (la evidencia del cierre no está versionada en git)")
            continue
        if case and not case_is_defined(repo, rel, case):
            problems.append(f"`{entry}` (el caso '{case}' no está definido en {rel})")
    if problems:
        return [
            f"HR-53: la sección '{heading}' declara '{marker}' y la evidencia del "
            "cierre no prueba nada: " + "; ".join(problems) + "."]
    return []


def missing_section_finding(body, heading, marker):
    """#334/#911: el hallazgo cuando no hay sección HR-53 dentro del cuerpo."""
    if section_body(body, marker):
        return (f"HR-53: la declaración '{marker}' debe ir dentro de la sección "
                f"'{heading}'.")
    return (f"HR-53: el PR lleva 'Closes #<issue>' y no tiene la sección "
            f"'{heading}': el cierre lo demuestra el test que prueba el "
            "comportamiento, no el merge.")


def check_closure_tests(pr, data, policy, findings):
    """#334 (HR-53): un PR que cierra una issue nombra el test que demuestra
    el comportamiento corregido, en formato estructurado y con el test
    existiendo de verdad (`tracked`: rutas de `git ls-files`). Un «debería estar arreglado» no cierra nada.
    #911: si el cierre no tiene test, la sección declara «No aplica» y nombra la
    evidencia del cierre (una ruta versionada); la etiqueta no exonera nada."""
    body = pr.get("body") or ""
    if not CLOSES_RX.search(body):
        return
    heading = policy.get("closure_tests_heading", DEFAULT_CLOSURE_HEADING)
    marker = policy.get("closure_not_applicable_heading",
                        DEFAULT_CLOSURE_NOT_APPLICABLE_HEADING)
    text = section_body(body, heading)
    marker_text = section_body(text, marker) if text else ""
    if not text:
        findings.append(missing_section_finding(body, heading, marker))
        return
    if marker_text:
        findings.extend(not_applicable_findings(data, heading, marker, marker_text))
        return
    if section_body(body, marker):
        findings.append(
            f"HR-53: la declaración '{marker}' debe ir dentro de la sección "
            f"'{heading}'.")
        return
    entries = closure_entries(text)
    if not entries:
        findings.append(
            f"HR-53: la sección '{heading}' no tiene ninguna entrada en el "
            "formato `ruta/al/fichero_de_test.ext` o `ruta::nombre_de_caso`: "
            "la prosa no nombra un test.")
        return
    findings.extend(structured_test_findings(data.repo, data.tracked, entries, heading))


def structured_test_findings(repo, tracked, entries, heading):
    """HR-53: cada entrada `ruta` o `ruta::caso` es un test versionado que existe."""
    problems = []
    for entry in entries:
        rel, _, case = entry.partition("::")
        if not TEST_FILE_RX.search(rel):
            problems.append(f"`{entry}` (no es un fichero de test)")
            continue
        if rel not in tracked:
            problems.append(f"`{entry}` (la ruta no está versionada en git)")
            continue
        if case and not case_is_defined(repo, rel, case):
            problems.append(f"`{entry}` (el caso '{case}' no está definido en {rel})")
    if not problems:
        return []
    return [f"HR-53: la sección '{heading}' nombra test(s) que no prueban nada: "
            + "; ".join(problems) + "."]


# #333: la sección Chain Context es obligatoria en un PR con `chain:partial` y
# en la punta de una cadena (base == rama por defecto). Sus campos son DATO y
# el gate los verifica: base contra baseRefName, posición y presupuesto con
# forma N/M, y exactamente un 📍 marcando el PR actual.
CHAIN_CONTEXT_HEADING = "Chain Context"
CHAIN_FIELDS = ("chain", "position", "base", "depends-on", "follow-up",
                "starts-at", "ends-with", "review-budget")
CHAIN_FIELD_RX = re.compile(r"^\s*-\s*([a-z][a-z0-9-]*)\s*:\s*(.*\S)\s*$", re.MULTILINE)
PAIR_RX = re.compile(r"^\d+\s*/\s*\d+$")
CHAIN_MARKER = "📍"
# #333 (auditoría de #352): el ENLACE de la issue es `Closes|Fixes|Resolves|Refs
# owner/repo#N`. Una mención informativa del cuerpo —citar una issue de Cadete
# como antecedente, por ejemplo— no es un enlace y no se juzga.
FOREIGN_LINK_RX = re.compile(
    r"\b(?:Closes|Fixes|Resolves|Refs)\s+([A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+)#(\d+)",
    re.IGNORECASE)


def chain_fields(text):
    """Campos `- nombre: valor` de la sección Chain Context."""
    return {m.group(1): m.group(2).strip() for m in CHAIN_FIELD_RX.finditer(text)}


def check_chain_context(pr, base, default_branch, policy, findings):
    """#333: la cadena se declara como dato, no se reconstruye a mano."""
    body = pr.get("body") or ""
    labels = {item["name"] for item in pr.get("labels", [])}
    chain_label = policy.get("chain_partial_label", "chain:partial")
    if chain_label not in labels and base != default_branch:
        return  # ni forma parte de una cadena ni es su punta
    heading = policy.get("chain_context_heading", CHAIN_CONTEXT_HEADING)
    text = section_body(body, heading)
    if not text:
        findings.append(
            f"HR-54: falta la sección '{heading}': un PR con '{chain_label}' y la punta "
            "de una cadena declaran posición, base, dependencias y presupuesto como "
            "datos, y marcan el PR actual con un 📍.")
        return
    fields = chain_fields(text)
    missing = [name for name in CHAIN_FIELDS if not fields.get(name)]
    if missing:
        findings.append(
            f"HR-54: la sección '{heading}' no declara: {', '.join(missing)} "
            f"(campos obligatorios: {', '.join(CHAIN_FIELDS)}).")
    declared_base = fields.get("base", "")
    if declared_base and declared_base != base:
        findings.append(
            f"HR-54: 'base' declara {declared_base!r} y el PR se abre contra {base!r}: "
            "la base de la cadena se verifica contra baseRefName.")
    for name in ("position", "review-budget"):
        value = fields.get(name, "")
        if value and not PAIR_RX.match(value):
            findings.append(
                f"HR-54: '{name}' debe tener la forma N/M (p. ej. 2/4) y declara {value!r}.")
    markers = text.count(CHAIN_MARKER)
    if markers != 1:
        findings.append(
            f"HR-54: el diagrama de '{heading}' marca el PR actual con exactamente un "
            f"📍 y tiene {markers}.")


# #876: la asistencia de IA se declara en el CUERPO del PR; los commits no llevan
# atribución (HR-14). Vacía o solo con el comentario de la plantilla no cuenta.
AI_DISCLOSURE_HEADING = "Asistencia de IA"
AI_DISCLOSURE_TEXT_RX = re.compile(r"[^\W\d_]")  # al menos una letra


def check_ai_disclosure(pr, findings):
    """#876: el cuerpo del PR declara la asistencia de IA en su sección. Ausente,
    vacía o solo con comentarios HTML falla; «ninguna» o una descripción pasan."""
    text = re.sub(r"<!--.*?-->", "", section_body(pr.get("body"), AI_DISCLOSURE_HEADING),
                  flags=re.S).strip()
    if not AI_DISCLOSURE_TEXT_RX.search(text):
        findings.append(
            f"HR-14: falta la sección '{AI_DISCLOSURE_HEADING}' en el cuerpo del PR, o está "
            "vacía: escriba «ninguna» o la herramienta de IA y la parte que ayudó. La "
            "declaración va en el PR, no en los commits.")


def check_issue_links(pr, data, findings):
    """#333: el enlace de la issue no enlaza una issue de otro repositorio; una
    mención informativa del cuerpo no es un enlace. Cerrar y referenciar el
    mismo número no es un fallo aquí ni en issue-spec (#787 retiró el aviso)."""
    body = pr.get("body") or ""
    current = (data.get("repository") or {}).get("full_name") or ""
    for owner_repo, number in sorted(set(FOREIGN_LINK_RX.findall(body))):
        if owner_repo != current:
            findings.append(
                f"HR-54: enlace a una issue de otro repositorio "
                f"({owner_repo}#{number}): el enlace apunta al repositorio del PR.")


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: la salida no debe depender del locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


# Cada comprobación activa que lee el rango git exige su fichero de datos; la
# ausencia sale 2 (HR-3), nunca pasa en silencio.
DATA_FLAGS = (
    ("closure_tests", "tracked_files", "--tracked-files"),
    ("docs_evidence", "changed_files", "--changed-files"),
    ("budget", "numstat_file", "--numstat-file"),
    ("commits", "commits_file", "--commits-file"),
)


def _load_inputs(args):
    """Evento y política; (data, pr, policy, error). `error` es el mensaje de salida 2."""
    try:
        data, pr = load_event(args.event_file)
    except (FileNotFoundError, ValueError, TypeError) as exc:
        return None, None, None, str(exc)
    except json.JSONDecodeError as exc:
        return None, None, None, f"payload JSON inválido: {exc}"
    try:
        policy = load_policy(args.policy_file)
    except (FileNotFoundError, json.JSONDecodeError) as exc:
        return None, None, None, f"política ilegible: {exc}"
    return data, pr, policy, None


def _load_range_data(args, active):
    """Textos de los ficheros del rango que las comprobaciones activas necesitan."""
    texts = {}
    if getattr(args, "changed_files", None) and "docs_evidence" not in active:
        try:
            texts["changed_files"] = read_data_file(args.changed_files)
        except OSError as exc:
            return None, f"--changed-files ilegible: {exc}"
    for check, attr, flag in DATA_FLAGS:
        if check not in active:
            continue
        path = getattr(args, attr)
        if not path:
            return None, f"la comprobación '{check}' está activa y falta {flag} (el gate no ejecuta git)."
        try:
            texts[attr] = read_data_file(path)
        except OSError as exc:
            return None, f"{flag} ilegible: {exc}"
    if "budget" in active and getattr(args, "origins_file", None):
        try:
            texts["origins_file"] = read_data_file(args.origins_file)
        except OSError as exc:
            return None, f"--origins-file ilegible: {exc}"
    return texts, None


def _pr_level_checks(active, data, pr, ctx, findings):
    repo, policy = ctx
    base_ref = pr.get("base", {}).get("ref", "main")
    default_branch = data.get("repository", {}).get("default_branch", "main")
    if "labels" in active:
        check_labels(pr, policy, findings)
    if "closes" in active:
        check_closes(pr, base_ref, default_branch, policy, findings)
    if "chain_context" in active:
        check_chain_context(pr, base_ref, default_branch, policy, findings)
    if "issue_links" in active:
        check_issue_links(pr, data, findings)
    if "ai_disclosure" in active:
        check_ai_disclosure(pr, findings)
    return repo


def closure_data(repo, texts, policy):
    """ClosureData desde los textos del rango (#918). Política inválida: sin rutas
    no ejecutables, así que «No aplica» falla (cerrado)."""
    raw = policy.get("closure_not_applicable_paths", DEFAULT_NOT_APPLICABLE_PATHS)
    paths = tuple(raw) if isinstance(raw, (list, tuple)) and all(
        isinstance(x, str) for x in raw) else ()
    changed = None
    if "changed_files" in texts:
        changed = tuple(line for line in texts["changed_files"].split("\n") if line)
    return ClosureData(Path(repo), frozenset(texts["tracked_files"].split("\n")),
                       changed, paths)


def _range_checks(active, pr, ctx, texts, findings):
    """Comprobaciones sobre el rango git; `DocsEvidenceConfigError` sube a main."""
    repo, policy = ctx
    if "closure_tests" in active:
        check_closure_tests(pr, closure_data(repo, texts, policy), policy, findings)
    if "docs_evidence" in active:
        check_docs_evidence(pr, repo, texts["changed_files"].split(), policy, findings)
    if "budget" in active:
        origins = parse_origins(texts["origins_file"]) if "origins_file" in texts else None
        check_budget(pr, budget_lines(texts["numstat_file"], origins), policy, findings)
    if "commits" in active:
        check_commits(texts["commits_file"], policy, findings)


def main(argv=None):
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description="Gate de contrato de PR y commits (#193).")
    parser.add_argument("--event-file", required=True, help="Payload del evento de PR.")
    parser.add_argument("--policy-file", required=True, help="Política del gate.")
    parser.add_argument("--repo", default=".", help="Raíz del repo, para leer ficheros (no ejecuta git).")
    parser.add_argument("--changed-files", help="`git diff --name-only BASE..HEAD`.")
    parser.add_argument("--numstat-file", help="`git diff --numstat BASE..HEAD`.")
    parser.add_argument("--commits-file",
                        help="`git log --format='%%P%%n%%B%%n---COMMIT-END---' BASE..HEAD`.")
    parser.add_argument("--tracked-files", help="`git ls-files`.")
    parser.add_argument("--origins-file",
                        help="Opcional (#931): origen de cada fichero vendorizado, que genera el runner local.")
    args = parser.parse_args(argv)

    data, pr, policy, error = _load_inputs(args)
    if error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2

    try:
        active = active_checks(policy)
    except ChecksConfigError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    off = [name for name in CHECK_NAMES if name not in active]
    if off:
        print(f"PR-CONTRACT: apagadas por política: {', '.join(off)}", file=sys.stderr)

    # #938/#996: etiquetas y cuerpo se leen por API (vivos) o del payload (congelado al abrir).
    if use_live_api(args.event_file, pr, os.environ):
        try:
            live_labels, live_body = fetch_pr_live(data, pr)
        except _GITHUB_REST.GitHubApiError as exc:
            print(f"ERROR: {exc}; {OBSOLETE_NOTE}", file=sys.stderr)
            return 2
        pr_with_live = dict(pr)
        pr_with_live["labels"] = live_labels
        pr_with_live["body"] = live_body
    else:
        pr_with_live = pr

    texts, error = _load_range_data(args, active)
    if error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2

    ctx = (Path(args.repo), policy)
    findings = []
    _pr_level_checks(active, data, pr_with_live, ctx, findings)
    try:
        _range_checks(active, pr_with_live, ctx, texts, findings)
    except DocsEvidenceConfigError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    # HR-31: el gate procesa el estado actual sin importar la acción
    # (opened, edited, labeled, unlabeled, synchronize).
    return _verdict(findings, pr, action_text=data.get("action", "opened"))


def _verdict(findings, pr, *, action_text):
    """Imprime el veredicto y devuelve 0 (OK) o 1 (hallazgos)."""
    if findings:
        print(f"FAIL: contrato de PR con {len(findings)} hallazgo(s):")
        for f in findings:
            print(f"  ✗ {f}")
        return 1
    print(f"PR-CONTRACT OK: PR #{pr.get('number', '?')} ({action_text}) sin hallazgos.")
    return 0

if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
