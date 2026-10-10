#!/usr/bin/env python3
"""Gate de contrato de PR y commits (#193): HR-6/7/8/14/31. Sin API."""

import argparse
import json
from pathlib import Path
import os
import re
import subprocess
import sys

def load_event(path):
    p = Path(path)
    if not p.is_file():
        raise FileNotFoundError(f"payload del evento ausente: {path}")
    data = json.loads(p.read_text(encoding="utf-8"))
    pr = data.get("pull_request")
    if not isinstance(pr, dict):
        raise ValueError("el payload no contiene `pull_request`.")
    return data, pr

def load_policy(path):
    p = Path(path)
    if not p.is_file():
        raise FileNotFoundError(f"política ausente: {path}")
    return json.loads(p.read_text(encoding="utf-8"))

def git_range(repo, base_sha, head_sha, fmt):
    result = subprocess.run(
        ["git", "-C", str(repo), "log", f"--format={fmt}", f"{base_sha}..{head_sha}"],
        capture_output=True, text=True, check=False,
    )
    return result.stdout if result.returncode == 0 else ""

def git_shortstat(repo, base_sha, head_sha):
    result = subprocess.run(
        ["git", "-C", str(repo), "diff", "--shortstat", f"{base_sha}..{head_sha}"],
        capture_output=True, text=True, check=False,
    )
    # " 5 files changed, 100 insertions(+), 50 deletions(-)"
    m = re.search(r"(\d+) insertion", result.stdout)
    ins = int(m.group(1)) if m else 0
    m = re.search(r"(\d+) deletion", result.stdout)
    dele = int(m.group(1)) if m else 0
    return ins + dele

def check_labels(pr, policy, findings):
    required = policy.get("required_labels", [])
    labels = {l["name"] for l in pr.get("labels", [])}
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
    labels = {l["name"] for l in pr.get("labels", [])}
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

SIZE_PLACEHOLDER = re.compile(r"<[^>]+>|\{\{[^}]+\}\}|^(?:TODO|FIXME)$", re.IGNORECASE)

def size_exception_reason(body, field="size-exception-reason:"):
    """HR-8: la excepción de presupuesto como DATO — una sola línea, una
    sola aparición en todo el cuerpo y valor sin el marcador de plantilla.
    Devuelve la razón válida o None si la excepción no es válida.

    Es la ÚNICA implementación de la regla (SKILL HR-8): la lee también el
    merge gobernado (`assets/runner-controls/check_governed_merge.py`), que
    la importa en proceso en lugar de copiarla."""
    text = body or ""
    needle = field.lower()
    lines = [line for line in text.splitlines() if needle in line.lower()]
    if len(lines) != 1 or text.lower().count(needle) != 1:
        return None
    line = lines[0]
    value = line[line.lower().find(needle) + len(field):].strip()
    if not value or SIZE_PLACEHOLDER.search(value):
        return None
    return value

def check_budget(pr, repo, base_sha, head_sha, policy, findings):
    budget = policy.get("review_budget_lines", 400)
    field = policy.get("size_exception_field", "size-exception-reason:")
    body = pr.get("body") or ""
    diff_lines = git_shortstat(repo, base_sha, head_sha)
    if diff_lines <= budget:
        return
    if size_exception_reason(body, field) is None:
        findings.append(
            f"HR-8: el diff supera el presupuesto ({diff_lines} > {budget}) y "
            f"no existe el campo de excepción '{field}' con reason en el cuerpo."
        )

def check_commits(repo, base_sha, head_sha, policy, findings):
    pattern = policy.get("conventional_commit_pattern",
                         r"^(feat|fix|docs|chore|ci|test|perf|refactor|revert)(\([^)]+\))?!?: ")
    ai_patterns = policy.get("ai_attribution_patterns",
                             ["Co-Authored-By:", "Generated-by:", "Generated with"])
    # %P delante: un commit de merge no tiene asunto conventional —la regla
    # mide los commits de trabajo— y la disciplina de cadenas obliga a
    # sincronizar con `git merge origin/main` en vez de rebase (#333).
    messages = git_range(repo, base_sha, head_sha, "%P%n%B%n---COMMIT-END---")
    for block in messages.split("---COMMIT-END---"):
        msg = block.strip()
        if not msg:
            continue
        parents, _, message = msg.partition("\n")
        subject = message.split("\n", 1)[0].strip()
        if len(parents.split()) < 2 and not re.match(pattern, subject):
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
GIT_TIMEOUT = 60
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


def is_tracked(repo, rel):
    """La ruta está versionada en git. Un fichero sin versionar no demuestra
    nada: no entra en el PR. Timeout explícito y fallo cerrado."""
    try:
        proc = subprocess.run(
            ["git", "-C", str(repo), "ls-files", "--error-unmatch", "--", rel],
            capture_output=True, text=True, check=False, timeout=GIT_TIMEOUT)
    except (OSError, subprocess.TimeoutExpired):
        return False
    return proc.returncode == 0


def case_is_defined(repo, rel, case):
    """El caso está DEFINIDO en el fichero (no basta con que la cadena
    aparezca). Se lee solo ese fichero: nada de recorrer el árbol."""
    try:
        text = (Path(repo) / rel).read_text(encoding="utf-8", errors="ignore")
    except OSError:
        return False
    return any(re.search(pattern.format(name=re.escape(case)), text)
               for pattern in CASE_DEF_PATTERNS)


def check_closure_tests(pr, repo, policy, findings):
    """#334 (HR-53): un PR que cierra una issue nombra el test que demuestra
    el comportamiento corregido, en formato estructurado y con el test
    existiendo de verdad. Un «debería estar arreglado» no cierra nada."""
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
        if not is_tracked(repo, rel):
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
CLOSES_NUMBER_RX = re.compile(r"\b(?:Closes|Fixes|Resolves)\s+#(\d+)", re.IGNORECASE)
REFS_NUMBER_RX = re.compile(r"\bRefs\s+#(\d+)", re.IGNORECASE)
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
    labels = {l["name"] for l in pr.get("labels", [])}
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
    """#333: el enlace de la issue no admite mezclas (cierre y referencia del
    mismo número) ni enlaza una issue de otro repositorio; una mención
    informativa del cuerpo no es un enlace."""
    body = pr.get("body") or ""
    both = sorted(set(CLOSES_NUMBER_RX.findall(body)) & set(REFS_NUMBER_RX.findall(body)))
    if both:
        findings.append(
            "HR-54: el cuerpo cierra y referencia la MISMA issue ("
            + ", ".join(f"#{n}" for n in both)
            + "): use 'Closes #N' o 'Refs #N', nunca las dos para el mismo número.")
    current = (data.get("repository") or {}).get("full_name") or ""
    for owner_repo, number in sorted(set(FOREIGN_LINK_RX.findall(body))):
        if owner_repo != current:
            findings.append(
                f"HR-54: enlace a una issue de otro repositorio "
                f"({owner_repo}#{number}): el enlace apunta al repositorio del PR.")


def main(argv=None):
    parser = argparse.ArgumentParser(description="Gate de contrato de PR y commits (#193).")
    parser.add_argument("--event-file", required=True, help="Payload del evento de PR.")
    parser.add_argument("--policy-file", required=True, help="Política del gate.")
    parser.add_argument("--repo", default=".", help="Ruta del repo para git log/diff.")
    args = parser.parse_args()

    try:
        data, pr = load_event(args.event_file)
    except (FileNotFoundError, ValueError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    except json.JSONDecodeError as exc:
        print(f"ERROR: payload JSON inválido: {exc}", file=sys.stderr)
        return 2

    try:
        policy = load_policy(args.policy_file)
    except (FileNotFoundError, json.JSONDecodeError) as exc:
        print(f"ERROR: política ilegible: {exc}", file=sys.stderr)
        return 2

    base_ref = pr.get("base", {}).get("ref", "main")
    base_sha = pr.get("base", {}).get("sha", "")
    head_sha = pr.get("head", {}).get("sha", "")
    repo = Path(args.repo)
    action = data.get("action", "opened")

    findings = []
    check_labels(pr, policy, findings)
    check_closes(pr, base_ref, data.get("repository", {}).get("default_branch", "main"), policy, findings)
    check_chain_context(pr, base_ref,
                        data.get("repository", {}).get("default_branch", "main"),
                        policy, findings)
    check_issue_links(pr, data, findings)
    check_closure_tests(pr, repo, policy, findings)

    if base_sha and head_sha and base_sha != head_sha:
        check_budget(pr, repo, base_sha, head_sha, policy, findings)
        check_commits(repo, base_sha, head_sha, policy, findings)
    # HR-31: el gate procesa el estado actual sin importar la acción
    # (opened, edited, labeled, unlabeled, synchronize).

    if findings:
        print(f"FAIL: contrato de PR con {len(findings)} hallazgo(s):")
        for f in findings:
            print(f"  ✗ {f}")
        return 1

    print(f"PR-CONTRACT OK: PR #{pr.get('number', '?')} ({action}) sin hallazgos.")
    return 0

if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
