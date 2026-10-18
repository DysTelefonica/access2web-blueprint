#!/usr/bin/env python3
"""Gate de contrato de PR y commits (#193): HR-6/7/8/14/31. Sin API ni procesos.

El gate no ejecuta `git`: el workflow genera los datos del rango y se los pasa
por fichero (`--changed-files`, `--numstat-file`, `--commits-file`,
`--tracked-files`), de modo que un consumer lo copia byte a byte sin hallazgos
de S603/S607 (DysTelefonica/team-skills#753). Un fichero necesario que falta
sale 2 (HR-3)."""

import argparse
import fnmatch
import json
import re
import sys
from pathlib import Path

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
        super().__init__("`checks` apaga las ocho comprobaciones: ninguna activa")


class DocsEvidenceConfigError(ValueError):
    """`docs_evidence` declarado pero inválido: falla cerrado (exit 2)."""

    def __init__(self):
        super().__init__(
            "docs_evidence necesita 'pages' (lista de {page, covers}) y 'checklists' (objeto)")


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

def read_data_file(path):
    """Texto de un fichero de datos del rango; ausente o ilegible: FileNotFoundError."""
    return Path(path).read_text(encoding="utf-8")


def numstat_total(numstat):
    """Adiciones + eliminaciones de `git diff --numstat`; los binarios (`-`) no suman."""
    total = 0
    for line in numstat.splitlines():
        added, _, rest = line.partition("\t")
        deleted = rest.partition("\t")[0]
        total += int(added) if added.isdigit() else 0
        total += int(deleted) if deleted.isdigit() else 0
    return total


CHECK_NAMES = ("labels", "closes", "chain_context", "issue_links", "closure_tests",
               "docs_evidence", "budget", "commits")


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

def check_commits(messages, policy, findings):
    pattern = policy.get("conventional_commit_pattern",
                         r"^(feat|fix|docs|chore|ci|test|perf|refactor|revert)(\([^)]+\))?!?: ")
    ai_patterns = policy.get("ai_attribution_patterns",
                             ["Co-Authored-By:", "Generated-by:", "Generated with"])
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
            if re.search(pat, msg, re.IGNORECASE):
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


def check_closure_tests(pr, repo, tracked, policy, findings):
    """#334 (HR-53): un PR que cierra una issue nombra el test que demuestra
    el comportamiento corregido, en formato estructurado y con el test
    existiendo de verdad (`tracked`: rutas de `git ls-files`). Un «debería estar arreglado» no cierra nada."""
    body = pr.get("body") or ""
    if not CLOSES_RX.search(body):
        return
    heading = policy.get("closure_tests_heading", DEFAULT_CLOSURE_HEADING)
    text = section_body(body, heading)
    if not text:
        findings.append(
            f"HR-53: el PR lleva 'Closes #<issue>' y no tiene la sección "
            f"'{heading}': el cierre lo demuestra el test que prueba el "
            "comportamiento, no el merge.")
        return
    entries = closure_entries(text)
    if not entries:
        findings.append(
            f"HR-53: la sección '{heading}' no tiene ninguna entrada en el "
            "formato `ruta/al/fichero_de_test.ext` o `ruta::nombre_de_caso`: "
            "la prosa no nombra un test.")
        return
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
    if problems:
        findings.append(
            f"HR-53: la sección '{heading}' nombra test(s) que no prueban nada: "
            + "; ".join(problems) + ".")


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


def check_issue_links(pr, data, findings):
    """#333: el enlace de la issue no enlaza una issue de otro repositorio; una
    mención informativa del cuerpo no es un enlace. Cerrar y referenciar el
    mismo número ya no es un fallo aquí: lo avisa issue-spec y gana `Closes`
    (#787, una sola evaluación)."""
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
    return repo


def _range_checks(active, pr, ctx, texts, findings):
    """Comprobaciones sobre el rango git; `DocsEvidenceConfigError` sube a main."""
    repo, policy = ctx
    if "closure_tests" in active:
        check_closure_tests(pr, repo, set(texts["tracked_files"].split("\n")), policy, findings)
    if "docs_evidence" in active:
        check_docs_evidence(pr, repo, texts["changed_files"].split(), policy, findings)
    if "budget" in active:
        check_budget(pr, numstat_total(texts["numstat_file"]), policy, findings)
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

    texts, error = _load_range_data(args, active)
    if error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2

    ctx = (Path(args.repo), policy)
    findings = []
    _pr_level_checks(active, data, pr, ctx, findings)
    try:
        _range_checks(active, pr, ctx, texts, findings)
    except DocsEvidenceConfigError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    # HR-31: el gate procesa el estado actual sin importar la acción
    # (opened, edited, labeled, unlabeled, synchronize).

    if findings:
        print(f"FAIL: contrato de PR con {len(findings)} hallazgo(s):")
        for f in findings:
            print(f"  ✗ {f}")
        return 1

    print(f"PR-CONTRACT OK: PR #{pr.get('number', '?')} ({data.get('action', 'opened')}) sin hallazgos.")
    return 0

if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
