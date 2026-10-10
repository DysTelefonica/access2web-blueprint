#!/usr/bin/env python3
"""Meta-gate de la matriz HR→gate (#191).

Uso: check_hr_matrix.py [<skill_root>]

Verifica que:
  1. Toda HR declarada en SKILL.md (`^- **HR-<n>`) tiene entrada en la
     matriz, y viceversa (sin HRs fantasma).
  2. `enforcement` es `gate` o `manual`.
  3. Una entrada `gate` apunta a un asset y a un test que existen.
  4. Una entrada `manual` lleva `reason` no vacía y `issue` numérico.

El smoke-run (#342): el test declarado de una regla `gate` se ejecuta con
la marca de recursión de la CLI RETIRADA, porque un test no es un gate
anidado — puede llamar a `adoption` en proceso (la suite de HR-45/46/48 lo
hace). El encadenamiento se cierra por IDENTIDAD, no por profundidad: el
test que ya se está ejecutando no se vuelve a smoke-ejecutar (su ruta viaja
en `CI_PATTERN_SMOKE_TESTS`), así que ningún camino puede recursar —el
incidente del 2026-10-05 superó 1.300 procesos y tumbó el VPS dos veces— y
los tests declarados por los fixtures siguen juzgándose normalmente. La
profundidad (`CI_PATTERN_SMOKE`) y su tope quedan como alambre trampa.

Salida: recuento `gate N, manual M` para seguir el avance. Exit 0 si
todo cuadra, 1 con violaciones (listadas), 2 ante entorno insalvable
(SKILL.md o matriz ausentes) — nunca verde en silencio.
"""

import json
import os
import re
import subprocess
import sys
from pathlib import Path

# La CLI es un script, no un módulo importable: el nombre de su marca de
# recursión vive aquí duplicado, y un meta-test de la suite lo compara con
# el de `assets/bin/ci-pattern` (#342).
RECURSION_ENV = "CI_PATTERN_NESTED"
SMOKE_ENV = "CI_PATTERN_SMOKE"
SMOKE_TESTS_ENV = "CI_PATTERN_SMOKE_TESTS"
SMOKE_MAX_DEPTH = 8


def smoke_state(environ=None):
    """(profundidad, tests ya en ejecución) del encadenamiento de smoke-runs.

    Un valor de profundidad que no sea un entero no negativo es duda (HR-3):
    el gate no adivina cuántos smoke-runs lleva encima."""
    env = environ or os.environ
    raw = env.get(SMOKE_ENV, "").strip()
    if raw and not raw.isdigit():
        raise ValueError(f"{SMOKE_ENV}={raw!r} no es un entero no negativo")
    depth = int(raw) if raw else 0
    running = {path for path in env.get(SMOKE_TESTS_ENV, "").split(os.pathsep) if path.strip()}
    return depth, running


def smoke_env(depth, running, test_path, environ=None):
    """Entorno con el que el meta-gate ejecuta el test declarado (#342).

    La marca de recursión de la CLI se RETIRA —un test que llama a
    `adoption` en proceso es legítimo y no es un gate anidado— y la ruta del
    test viaja con las que ya están corriendo: un meta-gate anidado no vuelve
    a smoke-ejecutar un test que ya se está ejecutando."""
    env = {k: v for k, v in (environ or os.environ).items()
           if k not in (RECURSION_ENV, SMOKE_TESTS_ENV)}
    env[SMOKE_ENV] = str(depth + 1)
    env[SMOKE_TESTS_ENV] = os.pathsep.join(sorted(running | {test_path}))
    return env


def skill_md_path(skill_root):
    return Path(skill_root) / "SKILL.md"


def matrix_path(skill_root):
    return Path(skill_root) / "references" / "hr-gate-matrix.json"


def declared_hrs(skill_root):
    """IDs `HR-<n>` declarados como `- **HR-<n> — ...` en SKILL.md."""
    text = skill_md_path(skill_root).read_text(encoding="utf-8")
    return set(re.findall(r"^- \*\*(HR-\d+)\b", text, re.MULTILINE))


def load_matrix(skill_root):
    data = json.loads(matrix_path(skill_root).read_text(encoding="utf-8"))
    rules = data.get("rules")
    if not isinstance(rules, list):
        raise ValueError("la matriz no declara la lista `rules`.")
    return rules


def main(argv):
    args = list(argv)
    open_issues = None
    if "--open-issues-file" in args:
        idx = args.index("--open-issues-file")
        if idx + 1 >= len(args):
            print("ERROR: --open-issues-file requiere un fichero.", file=sys.stderr)
            return 2
        issues_path = Path(args[idx + 1])
        args = args[:idx] + args[idx + 2:]
        if not issues_path.is_file():
            print(f"ERROR: fichero de issues abiertas ausente: {issues_path}", file=sys.stderr)
            return 2
        open_issues = set()
        for line in issues_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line.isdigit():
                open_issues.add(int(line))
    skill_root = Path(args[0]) if args else Path(__file__).resolve().parents[2]
    if not skill_md_path(skill_root).is_file():
        print(f"ERROR: SKILL.md ausente en {skill_root}; entorno insalvable.", file=sys.stderr)
        return 2
    if not matrix_path(skill_root).is_file():
        print(f"ERROR: matriz ausente en {matrix_path(skill_root)}; entorno insalvable.", file=sys.stderr)
        return 2

    hrs = declared_hrs(skill_root)
    try:
        rules = load_matrix(skill_root)
    except (json.JSONDecodeError, ValueError) as exc:
        print(f"ERROR: matriz ilegible: {exc}", file=sys.stderr)
        return 2

    try:
        depth, running = smoke_state()
    except ValueError as exc:
        print(f"ERROR: {exc}; entorno insalvable.", file=sys.stderr)
        return 2
    if depth > SMOKE_MAX_DEPTH:
        print(
            f"ERROR: cadena de smoke-runs de profundidad {depth} (tope "
            f"{SMOKE_MAX_DEPTH}): la guarda anti-recursión del meta-gate está "
            "rota; no se mide nada más (#342).",
            file=sys.stderr,
        )
        return 2
    # Tests que ya se están ejecutando en este encadenamiento: no se repiten.
    skipped = []

    violations = []
    seen = set()
    counts = {"gate": 0, "manual": 0, "process": 0}
    for rule in rules:
        rid = rule.get("id", "?")
        if rid in seen:
            violations.append(f"{rid}: entrada duplicada en la matriz.")
            continue
        seen.add(rid)
        if rid not in hrs:
            violations.append(f"{rid}: la matriz cita una HR que no está en SKILL.md.")
            continue
        enforcement = rule.get("enforcement")
        if enforcement == "gate":
            for field in ("asset", "test"):
                target = rule.get(field)
                if not isinstance(target, str) or not target.strip():
                    violations.append(
                        f"{rid}: gate con '{field}' vacío o ausente; un gate sin ejecutable no gobierna."
                    )
                    continue
                first = target.split()[0]
                target_path = skill_root / first
                if not target_path.is_file():
                    violations.append(
                        f"{rid}: gate con {field} inexistente ({target!r}); un gate que no existe no gobierna."
                    )
                    continue
                if field == "test":
                    resolved = str(target_path.resolve())
                    if resolved in running:
                        # Este test ya se está ejecutando más arriba en la
                        # cadena: volver a smoke-ejecutarlo es la recursión
                        # que tumbó el VPS el 2026-10-05 (#277). Se valida la
                        # estructura y el informe lo declara (#342).
                        skipped.append(f"{rid} ({target})")
                        continue
                    # Smoke-execución del contrato: el comando declarado debe
                    # pasar HOY, no bastar con existir (#191 corrección).
                    if first.endswith(".py"):
                        cmd = [sys.executable, str(target_path)] + target.split()[1:]
                    else:
                        cmd = [str(target_path)] + target.split()[1:]
                    try:
                        run = subprocess.run(
                            cmd, cwd=skill_root, capture_output=True,
                            text=True, timeout=120, check=False,
                            env=smoke_env(depth, running, resolved),
                        )
                    except subprocess.TimeoutExpired:
                        violations.append(f"{rid}: el test '{target}' excede el límite de 120s.")
                        continue
                    except OSError as exc:
                        violations.append(f"{rid}: el test '{target}' no es ejecutable ({exc}).")
                        continue
                    if run.returncode != 0:
                        violations.append(
                            f"{rid}: el test del gate falla (rc={run.returncode}); "
                            "un gate cuyo test está roto no gobierna."
                        )
            counts["gate"] += 1
        elif enforcement == "manual":
            reason = rule.get("reason", "")
            issue = rule.get("issue")
            if not isinstance(reason, str) or not reason.strip():
                violations.append(f"{rid}: manual sin `reason`; una exención sin motivo no gobierna.")
            if not isinstance(issue, int) or issue <= 0:
                violations.append(f"{rid}: manual sin `issue` que cubra su mecanización.")
            elif open_issues is not None and issue not in open_issues:
                violations.append(
                    f"{rid}: la issue {issue} está cerrada o no existe; "
                    "una manual apunta a una issue que no cubre su mecanización."
                )
            counts["manual"] += 1
        elif enforcement == "process":
            reason = rule.get("reason", "")
            if not isinstance(reason, str) or not reason.strip():
                violations.append(
                    f"{rid}: process sin `reason`; una regla de proceso sin motivo documentado no gobierna."
                )
            counts["process"] += 1
        else:
            violations.append(f"{rid}: enforcement desconocido {enforcement!r} (gate|manual|process).")

    for rid in sorted(hrs - seen):
        violations.append(f"{rid}: declarada en SKILL.md y ausente de la matriz.")

    if violations:
        print(f"FAIL: matriz HR→gate con {len(violations)} violación(es):")
        for v in violations:
            print(f"  ✗ {v}")
        return 1

    # #235: el informe declara explicitamente si la comprobación de issues
    # se realizó; sin lista nunca sugiere que las issues fueron verificadas.
    if open_issues is None:
        detail = ("(gate = asset+test; manual = reason+issue; process = reason). "
                  "comprobación de issues no realizada (sin --open-issues-file): "
                  "las issues de las entradas manual no se validaron.")
    else:
        detail = (f"(gate = asset+test; manual = reason+issue abierta verificada; "
                  f"process = reason). issues abiertas verificadas "
                  f"({len(open_issues)} números en la lista).")
    if skipped:
        detail += (f" smoke-run omitido por profundidad {depth} en {len(skipped)} "
                   "regla(s) cuyo test ya se está ejecutando (solo estructura): "
                   + ", ".join(skipped) + ".")
    print(
        f"HR-GATE-MATRIX OK: {len(hrs)} reglas — gate {counts['gate']}, "
        f"manual {counts['manual']}, process {counts['process']} {detail}"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
