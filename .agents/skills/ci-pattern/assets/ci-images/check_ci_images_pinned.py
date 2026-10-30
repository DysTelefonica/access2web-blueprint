#!/usr/bin/env python3
# ci-pattern asset — gate de imágenes de CI fijadas por digest (DysTelefonica/team-skills#1264)
"""check_ci_images_pinned.py <repo>: toda imagen LITERAL de CI lleva `@sha256:<64 hex>`.
P56 activa el gate; P57 excepciona por referencia (motivo no vacío). Una etiqueta que el
workflow construye (-t/--tag) es local: exenta en run y build, no en pull. Con `$` → NO-VERIFICADA.
Exit: 0 limpio/inerte · 1 hallazgos · 2 entrada inválida (fail-closed, HR-3).
"""

from __future__ import annotations

import importlib.util
import re
import sys
from pathlib import Path

# Lector estructural del canon, importado por ruta (sin PyYAML).
_WP_PATH = Path(__file__).resolve().parents[1] / "workflow-policy" / "check_workflow_policy.py"
_spec = importlib.util.spec_from_file_location("check_workflow_policy", _WP_PATH)
WP = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(WP)

DIGEST = re.compile(r"@sha256:[0-9a-f]{64}$")
IMAGE_OPT = re.compile(r"(?:^|[\s,])image=([^\s,]+)")
# `${{ }}` y `$( )` cuentan como un solo token.
EXPR = re.compile(r"\$\{\{.*?\}\}|\$\(.*?\)")
# Flags que consumen valor; el resto (-d, --rm, -it, --privileged...) son booleanas.
_VALUE = "-v -w -u -e -a -p -h --network --name --env-file --mount --entrypoint --label --platform --pull --network-alias --hostname --workdir --user --env --publish --volume --gpus --runtime --memory --memory-swap --cpus --cpuset-cpus --pids-limit --security-opt --cap-add --device --tmpfs --ulimit --stop-signal --shm-size --pid --ipc --uts --userns --cgroupns"
_BUILD = "-f --file -t --tag --build-arg --target --cache-from --secret -o --output --iidfile"
VALUE_FLAGS = set(_VALUE.split())
BUILD_VALUE_FLAGS = VALUE_FLAGS | set(_BUILD.split())


class EntradaInvalidaError(Exception):
    """El gate no puede interpretar el entorno; sale 2, nunca adivina."""


def _pin_output_encoding() -> None:
    """Hace la salida del gate independiente de la locale del runner."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def read_contract(repo: Path):
    """(modo P56, excepciones P57). Fail-closed: contrato ausente, ilegible o inválido."""
    path = repo / "ci-pattern.yaml"
    if not path.is_file():
        msg = f"ci-pattern.yaml ausente en {path}; sin contrato no hay gate"
        raise EntradaInvalidaError(msg)
    try:
        lines = WP.read_lines(path.read_text(encoding="utf-8"))
        doc, i = WP.parse_node(lines, 0, lines[0].indent) if lines else ({}, 0)
    except (WP.InvalidInput, OSError, UnicodeDecodeError) as exc:
        msg = f"ci-pattern.yaml ilegible ({path}): {exc}"
        raise EntradaInvalidaError(msg) from exc
    if i < len(lines) or not isinstance(doc, dict):
        msg = f"ci-pattern.yaml ilegible ({path}): se esperaba un mapa"
        raise EntradaInvalidaError(msg)
    mode = str(doc.get("P56_ci_images_gate", "off"))
    if mode not in ("off", "on"):
        msg = f"P56_ci_images_gate={mode!r} fuera de {{off, on}}"
        raise EntradaInvalidaError(msg)
    exceptions = doc.get("P57_ci_image_exceptions") or {}
    if not isinstance(exceptions, dict):
        msg = "P57_ci_image_exceptions debe ser un mapa referencia→motivo"
        raise EntradaInvalidaError(msg)
    return mode, exceptions


def is_image(token: str) -> bool:
    """Una ruta o URL es contexto de build, no imagen; un token con `$` no es literal."""
    local = token in (".", "..") or token.startswith(("./", "../", "/", "~/")) or "://" in token
    return "$" in token or (not local and (":" in token or "/" in token))


def docker_commands(run):
    """(subcomando, argumentos) de cada docker run|pull|build; `\\` une líneas lógicas."""
    if not isinstance(run, str):
        return
    for line in re.sub(r"\\[ \t]*\n", " ", run).splitlines():
        for segment in re.split(r"&&|\|\||;|\|", line):
            tokens = EXPR.sub(lambda m: m.group(0).replace(" ", ""), segment).split()
            for k in range(len(tokens) - 1):
                if tokens[k] == "docker" and tokens[k + 1] in ("run", "pull", "build"):
                    yield tokens[k + 1], tokens[k + 2 :]
                    break


def first_positional(args: list, value_flags: set):
    it = iter(args)  # `--flag=valor` ocupa un solo token: no consume el siguiente
    for tok in it:
        if tok.startswith("-"):
            if tok in value_flags:
                next(it, None)
        else:
            return tok
    return None


def build_tags(args: list) -> list:
    """Etiquetas que construye el workflow: -t v, --tag v, --tag=v."""
    tags = [args[k + 1] for k, tok in enumerate(args[:-1]) if tok in ("-t", "--tag")]
    return tags + [tok.partition("=")[2] for tok in args if tok.startswith("--tag=")]


def steps_of(job: dict) -> list:
    raw = job.get("steps") if isinstance(job.get("steps"), list) else []
    return [(i, s) for i, s in enumerate(raw) if isinstance(s, dict)]


def built_tags(doc: dict) -> set:
    """Etiquetas que construye cualquier step del workflow."""
    return {
        tag
        for job in doc["jobs"].values()
        for _, step in steps_of(job)
        for sub, args in docker_commands(step.get("run"))
        if sub == "build"
        for tag in build_tags(args)
    }


def step_refs(job_id: str, idx: int, step: dict, built: set) -> list:
    """`uses: docker://`, `with` tras `image=` y `docker run|pull|build` de un step."""
    where = f"steps[{idx}]"
    refs = []
    uses = step.get("uses")
    if isinstance(uses, str) and uses.startswith("docker://"):
        refs.append((job_id, f"{where}.uses", uses[len("docker://") :]))
    with_block = step.get("with") if isinstance(step.get("with"), dict) else {}
    for key, value in with_block.items():
        if isinstance(value, str):
            refs += [(job_id, f"{where}.with.{key}", m) for m in IMAGE_OPT.findall(value)]
    for sub, args in docker_commands(step.get("run")):
        ref = first_positional(args, BUILD_VALUE_FLAGS if sub == "build" else VALUE_FLAGS)
        local = sub != "pull" and ref in built  # pull siempre descarga del registro
        if ref and not local and (sub != "build" or is_image(ref)):
            refs.append((job_id, f"{where}.run", ref))
    return refs


def image_refs(doc: dict) -> list:
    """(job, dónde, referencia) de cada superficie de imagen del workflow."""
    built = built_tags(doc)
    refs = []
    for job_id, job in doc["jobs"].items():
        container = job.get("container")
        image = container.get("image") if isinstance(container, dict) else container
        if image is not None:
            refs.append((job_id, "container.image", str(image)))
        services = job.get("services") if isinstance(job.get("services"), dict) else {}
        for name, svc in services.items():
            if isinstance(svc, dict) and svc.get("image") is not None:
                refs.append((job_id, f"services.{name}.image", str(svc["image"])))
        for idx, step in steps_of(job):
            refs += step_refs(job_id, idx, step, built)
    return refs


def exception_findings(exceptions: dict, literals: set) -> list:
    """P57: motivo vacío y excepción obsoleta son hallazgos, nunca silencio."""
    out = []
    for ref, motivo in exceptions.items():
        if not str(motivo or "").strip():
            out.append(f"EXCEPCION-SIN-MOTIVO {ref}")
        if ref not in literals:
            out.append(f"EXCEPCION-OBSOLETA {ref}")
    return out


def gate(repo: Path, exceptions: dict) -> int:
    excused = {ref for ref, motivo in exceptions.items() if str(motivo or "").strip()}
    wf_dir = repo / ".github" / "workflows"
    files = sorted([*wf_dir.glob("*.yml"), *wf_dir.glob("*.yaml")]) if wf_dir.is_dir() else []
    findings, no_verificadas, literals, total = [], [], set(), 0
    for path in files:
        try:
            doc = WP.parse_workflow(path.read_text(encoding="utf-8"), str(path))
        except (WP.InvalidInput, OSError, UnicodeDecodeError) as exc:
            raise EntradaInvalidaError(str(exc)) from exc
        for job_id, where, ref in image_refs(doc):
            total += 1
            label = f"{path.name}:{job_id}:{where}"
            if "$" in ref:
                no_verificadas.append(f"NO-VERIFICADA {label} -> {ref}")
            elif not DIGEST.search(ref):
                literals.add(ref)
                if ref not in excused:
                    findings.append(f"IMAGEN-SIN-DIGEST {label} -> {ref}")
    findings += exception_findings(exceptions, literals)
    print(f"PINES: on | workflows: {len(files)} | refs: {total} | hallazgos: {len(findings)}")
    for line in findings + no_verificadas:
        print(line)
    return 1 if findings else 0


def main(argv) -> int:
    _pin_output_encoding()
    repo = Path(argv[1] if len(argv) > 1 else ".")
    if not repo.is_dir():
        print(f"ERROR: repo inexistente: {repo}", file=sys.stderr)
        return 2
    try:
        mode, exceptions = read_contract(repo)
        if mode == "off":
            print("PINES: off — gate inerte (P56_ci_images_gate ausente u off)")
            return 0
        return gate(repo, exceptions)
    except EntradaInvalidaError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv))
