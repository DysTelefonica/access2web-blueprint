#!/usr/bin/env python3
"""Local/CI parity gate (HR-4, #196; full contract documented in the
HR-4 amendment and the e2e suite): structural set comparison of the
declared CI job's run commands against the declared local entrypoint,
plus relative order; exclusions are DATA (command fnmatch, side, reason).
Exit: 0 parity, 1 findings, 2 invalid/empty (HR-18).
"""
import argparse
import fnmatch
import json
import os
import re
import sys
from pathlib import Path



def fail2(msg):
    print(f"fail-closed: {msg}", file=sys.stderr)
    return 2


def _logical_lines(text):
    joined, buf = [], ""
    for raw in text.splitlines():
        line = raw.rstrip()
        if line.endswith("\\"):
            buf += line[:-1] + " "
            continue
        buf += line
        joined.append(buf)
        buf = ""
    if buf:
        joined.append(buf)
    return joined


def _strip_noise(line):
    line = re.sub(r"\s+#.*$", "", line).strip().rstrip(";").strip()
    return line.strip('"').strip("'").strip()


def _is_control(line):
    return (not line or '"$(' in line or "$(cd " in line
            or line.startswith(("#", "set ", "export ", "echo ", "exit",
                                "then", "else", "fi", "done", "fail=",
                                "while ", "cd ", "["))
            or line in ("if", "do"))


def _signature(cmd):
    toks = cmd.split()
    while toks and re.match(r"^[A-Za-z_][A-Za-z0-9_]*=", toks[0]):
        toks = toks[1:]
    if toks and toks[0] in ("sudo", "env"):
        toks = toks[1:]
    if len(toks) >= 2:
        return f"{toks[0]} {toks[1]}"
    return " ".join(toks)


def _command_lines(script):
    for raw in _logical_lines(script):
        line = _strip_noise(raw)
        # Strip control prefixes BEFORE the control check, then re-check:
        # `if ! bash x; then` becomes `bash x`, `if [ ! -s f ]; then` becomes
        # `[ ... ]` (control) and is dropped.
        for _ in range(2):
            if _is_control(line):
                break
            if line.startswith("if "):
                line = re.sub(r"^if\s+", "", line)
                line = re.sub(r";\s*then$", "", line)
            elif line.startswith("! "):
                line = re.sub(r"^!\s*", "", line)
            else:
                break
        if _is_control(line) or not line:
            continue
        yield line


def _expand_loops(script, root):
    """Expand `name=( ... )` arrays and `for V in <globs|array>; do` bodies."""
    arrays, commands = {}, []
    lines = _logical_lines(script)
    i = 0
    while i < len(lines):
        line = lines[i]
        m = re.match(r"^(\w+)=\(\s*$", line.strip())
        if m:
            name, i = m.group(1), i + 1
            entries = []
            while i < len(lines) and not lines[i].strip().startswith(")"):
                entries += [_strip_noise(t) for t in lines[i].strip().split()
                            if _strip_noise(t)]
                i += 1
            arrays[name] = entries
            i += 1
            continue
        m = re.match(r"^for\s+(\w+)\s+in\s+(.+?);\s*do\s*$", _strip_noise(line))
        if m:
            var, items_src, i = m.group(1), m.group(2), i + 1
            items = []
            for item in items_src.split():
                clean = item.strip('"\'')
                if clean.startswith("${") and clean.endswith("}"):
                    name = clean[2:-1].split("[")[0]
                    if name not in arrays:
                        raise ValueError(
                            f"el entrypoint expande ${{{name}}} pero el array "
                            f"'{name}' no está declarado; extracción dudosa (HR-3).")
                    items += arrays[name]
                else:
                    items += sorted(os.path.relpath(str(p), str(root))
                                    for p in root.glob(clean))
            while i < len(lines) and not lines[i].strip().startswith("done"):
                body = _strip_noise(lines[i])
                body = re.sub(r"^if\s+!?\s*", "", body)
                body = re.sub(r";\s*then$", "", body)
                if var in body and not _is_control(body):
                    for item in items:
                        commands.append(re.sub(rf'"\$\{{{var}\}}"|"\${var}"|\${var}', item, body))
                i += 1
            i += 1
            continue
        line = _strip_noise(line)
        line = re.sub(r"^if\s+!?\s*", "", line)
        line = re.sub(r";\s*then$", "", line)
        if not _is_control(line) and line:
            commands.append(line)
        i += 1
    return commands


def _side_local(path, root):
    out = []
    for cmd in _expand_loops(Path(path).read_text(encoding="utf-8"), root):
        if cmd:
            sig = _signature(cmd)
            if sig and not sig.startswith("gh "):
                out.append(sig)
    return out


def _run_blocks(workflow_text, job):
    # Accepts ``- run:``/block and named forms deeper than ``steps:``.
    lines = workflow_text.splitlines()
    starts = [i for i, l in enumerate(lines)
              if re.match(rf"^  {re.escape(job)}:\s*(#.*)?$", l)]
    if len(starts) != 1:
        raise ValueError(f"job '{job}' no aparece exactamente una vez en el workflow.")
    head = re.compile(r"^(\s*)(-\s+)?run:\s*(.*)$")
    block_markers = ("|", ">", "|-", ">-", "|+", ">+")
    blocks, current, base = [], None, 0
    steps_indent, i = None, starts[0] + 1
    while i < len(lines):
        line = lines[i]
        if line and not line[0].isspace():
            break
        if re.match(r"^  [^\s#].*:\s*(#.*)?$", line):
            # a known job-internal key at odd indentation does not end the job
            if line.strip().split(":")[0].strip() not in ("name", "needs", "if", "runs-on", "timeout-minutes",
                           "strategy", "concurrency", "env", "defaults", "outputs",
                           "environment", "steps", "services", "container",
                           "permissions", "continue-on-error"):
                break
        if current is not None:
            if line.strip() == "" or (len(line) - len(line.lstrip())) > base:
                current.append(line)
                i += 1
                continue
            blocks.append("\n".join(current))
            current = None  # fall through: re-examine the closing line
        m = re.match(r"^(\s+)steps:\s*$", line)
        if m:
            steps_indent = len(m.group(1))
            i += 1
            continue
        m = head.match(line) if steps_indent is not None else None
        if m and len(m.group(1)) > steps_indent:
            rest = m.group(3).strip()
            # base = column of the `run:` KEY: dash column + 2 when the
            # shorthand `- run:` form is used, so sibling step keys (shell:,
            # env:, working-directory: at the same column) are never
            # absorbed into the script block.
            base = len(m.group(1)) + (2 if m.group(2) else 0)
            if rest and rest not in block_markers:
                blocks.append(rest)
            else:
                current = []
            i += 1
            continue
        i += 1
    if current is not None:
        blocks.append("\n".join(current))
    return blocks


def _side_ci(workflow_path, job, entrypoint, local_sigs):
    text = Path(workflow_path).read_text(encoding="utf-8")
    out = []
    for block in _run_blocks(text, job):
        for cmd in _command_lines(block):
            sig = _signature(cmd)
            if not sig:
                continue
            if sig == f"bash {entrypoint}" or sig.endswith(f" {entrypoint}"):
                out.extend(local_sigs)  # entrypoint == local set, in place
            else:
                out.append(sig)
    return out


def _dedupe(sigs):
    return list(dict.fromkeys(sigs))


def _excluded(sig, exclusions, side, used):
    for exc in exclusions:
        if exc.get("side") in ("any", side) and (
                fnmatch.fnmatch(sig, exc["command"]) or exc["command"] == sig):
            used.add(id(exc))
            return True
    return False


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Read-only local/CI parity gate (HR-4, #196).")
    ap.add_argument("--policy", required=True)
    ap.add_argument("--repo", default=".")
    args = ap.parse_args(argv)
    root = Path(args.repo)
    try:
        policy = json.loads(Path(args.policy).read_text(encoding="utf-8"))
        ci_cfg, local_cfg = policy["ci"], policy["local"]
        workflow, job = ci_cfg["workflow"], ci_cfg["job"]
        entrypoint = local_cfg["entrypoint"]
        exclusions = policy.get("exclusions", [])
        if not isinstance(exclusions, list):
            raise KeyError("exclusions debe ser una lista")
    except (OSError, json.JSONDecodeError, KeyError, TypeError) as exc:
        return fail2(f"política ilegible o incompleta ({args.policy}): {exc}")
    for name in (workflow, entrypoint):
        if not (root / name).is_file():
            return fail2(f"'{name}' no existe en {root}; sujeto ausente es duda (HR-3).")
    try:
        local_sigs = _dedupe(_side_local(root / entrypoint, root))
        ci_sigs = _dedupe(_side_ci(root / workflow, job, entrypoint, local_sigs))
    except ValueError as exc:
        return fail2(str(exc))
    if not local_sigs or not ci_sigs:
        return fail2("sujeto vacío: un lado sin comandos no es paridad (HR-18).")

    findings, used = [], set()
    for exc in exclusions:
        if not isinstance(exc.get("reason"), str) or not exc["reason"].strip():
            findings.append(f"exclusión {exc.get('command')!r} sin 'reason'; "
                            "una exención sin motivo no gobierna (HR-34).")
    local_f = [s for s in local_sigs if not _excluded(s, exclusions, "local", used)]
    ci_f = [s for s in ci_sigs if not _excluded(s, exclusions, "ci", used)]
    for tag, side, other in (("del CI sin entrada local", ci_f, local_f),
                             ("local sin correspondencia en el CI", local_f, ci_f)):
        for sig in sorted(set(side) - set(other)):
            findings.append(f"comando {tag}: {sig!r}")
    common = [s for s in local_f if s in set(ci_f)]
    if common != [s for s in ci_f if s in set(common)]:
        findings.append(
            "el orden relativo de los gates difiere entre local y CI: "
            f"local={common!r} vs ci={[s for s in ci_f if s in set(common)]!r}")
    for exc in exclusions:
        if id(exc) not in used:
            findings.append(f"exclusión declarada que no coincide con ningún comando: "
                            f"{exc.get('command')!r}")

    for sig in sorted(set(local_f) & set(ci_f)):
        print(f"  ok  {sig}")
    if findings:
        print(f"FAIL: paridad local/CI con {len(findings)} hallazgo(s):")
        for f in findings:
            print(f"  ✗ {f}")
        return 1
    print(f"LOCAL/CI PARITY OK: {len(set(local_f) & set(ci_f))} gates comunes "
          f"(exclusiones aplicadas: {len(used)}).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
