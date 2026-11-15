#!/usr/bin/env python3
# ci-pattern asset — in-code baseline relaxation guard (DysTelefonica/team-skills#974, Refs #429)
"""HR-15 guard: in-code gate baselines are compared with the base branch.

Every gate script of a consumer keeps a shrink-only ``BASELINE*`` constant that
the same PR can edit, so a gate that only compares the tree with its own
baseline never stops a PR from raising it. This guard reads each changed
script at ``HEAD`` and at the merge base and rejects any relaxation: a number
that grows, a new dict entry, a new set element. A relaxation is accepted only
when the script also declares a module-level field::

    RELAXATION = {"reason": "<why the baseline must grow>", "reference": "<issue or PR>",
                  "issue": <positive int>}  # or "actor": "<who decided>" (HR-33)

The field must be new in this PR: one that the base copy already carries
justified an earlier change and cannot cover another. Lowering a baseline or
removing entries is always accepted.

The same field covers the other way to loosen a baseline (#982): deleting the
comment that justifies an entry. A justification is a ``#`` comment inside the
span of a ``*BASELINE*`` constant, tied to the entry that follows it (or
written on its line); a comment with no entry after it belongs to the
constant. Dropping or rewording one is rejected unless its entry went away
with it. Moving it, or re-spacing it, is free. Justifications are compared
as ``(entry, text)`` pairs, so the same text on another entry cannot hide a
deleted justification.

A ``*BASELINE*`` whose value is not a literal is reported on stderr as not
protected by this guard. Such a baseline is never silently skipped.

Usage::

    python check_baseline_relaxation.py [--base-ref origin/main] [--gate-glob "scripts/*.py"]

Without ``--base-ref`` the base is ``origin/$GITHUB_BASE_REF``; with neither
the comparison is declared NOT performed and the run exits 0.

Exit codes: 0 clean or not performed, 1 relaxation without field, 2 the base
cannot be read (the gate fails closed, it never skips silently).
"""

from __future__ import annotations

import argparse
import ast
import io
import os
import re
import subprocess
import sys
import tokenize
from collections import Counter
from collections.abc import Iterable
from pathlib import Path

EXIT_OK = 0
EXIT_RELAXED = 1
EXIT_UNREADABLE = 2
BASELINE_NAME = re.compile(r"^[A-Z0-9_]*BASELINE[A-Z0-9_]*$")
RELAXATION_NAME = "RELAXATION"
DEFAULT_GATE_GLOB = "scripts/*.py"


class BaseUnreadableError(RuntimeError):
    """The base branch could not be resolved; the comparison cannot be trusted."""

    def __init__(self, command: str, detail: str) -> None:
        super().__init__(f"{command}: {detail}")


def _unwrap(node: ast.expr) -> ast.expr:
    """``frozenset({...})`` / ``set(...)`` literals are read as their argument."""
    if (
        isinstance(node, ast.Call)
        and isinstance(node.func, ast.Name)
        and node.func.id in {"frozenset", "set"}
    ):
        return node.args[0] if node.args else ast.Set(elts=[])
    return node


def _assigned(source: str) -> Iterable[tuple[str, ast.expr]]:
    for node in ast.parse(source).body:
        if isinstance(node, ast.Assign):
            names = [t.id for t in node.targets if isinstance(t, ast.Name)]
        elif isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name):
            names = [node.target.id]
        else:
            continue
        if node.value is not None:
            for name in names:
                yield name, node.value


def read_baselines(source: str) -> dict[str, object]:
    """Literal value of every module-level ``*BASELINE*`` constant."""
    found: dict[str, object] = {}
    for name, value in _assigned(source):
        if not BASELINE_NAME.match(name):
            continue
        try:
            found[name] = ast.literal_eval(_unwrap(value))
        except (ValueError, TypeError):
            continue
    return found


def read_unreadable_baselines(source: str) -> list[str]:
    """Names of module-level ``*BASELINE*`` constants that are not literals."""
    unreadable: list[str] = []
    for name, value in _assigned(source):
        if not BASELINE_NAME.match(name):
            continue
        try:
            ast.literal_eval(_unwrap(value))
        except (ValueError, TypeError):
            unreadable.append(name)
    return unreadable


def _entry_key(node: ast.expr) -> str:
    try:
        return repr(ast.literal_eval(node))
    except (ValueError, TypeError):
        return ast.dump(node)


def _entries(value: ast.expr) -> list[tuple[int, str]]:
    """``(line, key)`` of each top-level entry or element of a baseline literal."""
    container = _unwrap(value)
    if isinstance(container, ast.Dict):
        return [(k.lineno, _entry_key(k)) for k in container.keys if k is not None]
    if isinstance(container, ast.Set | ast.List | ast.Tuple):
        return [(e.lineno, _entry_key(e)) for e in container.elts]
    return []


def read_justifications(source: str) -> dict[str, tuple[set[str], list[tuple[str | None, str]]]]:
    """Per ``*BASELINE*`` constant: its entry keys and ``(entry, comment)`` pairs.

    A comment belongs to the entry on its own line, else to the next entry; with
    none after it, to the constant (``None``). Text is whitespace-normalised.
    """
    comments = [
        (tok.start[0], " ".join(tok.string.lstrip("#").split()))
        for tok in tokenize.generate_tokens(io.StringIO(source).readline)
        if tok.type == tokenize.COMMENT
    ]
    found: dict[str, tuple[set[str], list[tuple[str | None, str]]]] = {}
    for node in ast.parse(source).body:
        if isinstance(node, ast.Assign):
            names = [t.id for t in node.targets if isinstance(t, ast.Name)]
        elif isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name):
            names = [node.target.id]
        else:
            continue
        if node.value is None:
            continue
        entries = _entries(node.value)
        end = node.end_lineno or node.lineno
        for name in (n for n in names if BASELINE_NAME.match(n)):
            pairs: list[tuple[str | None, str]] = []
            for line, text in comments:
                if node.lineno <= line <= end and text:
                    owner = next((k for ln, k in entries if ln == line), None) or next(
                        (k for ln, k in entries if ln > line), None
                    )
                    pairs.append((owner, text))
            found[name] = ({key for _, key in entries}, pairs)
    return found


def deleted_justifications(base_source: str, head_source: str) -> list[str]:
    """Justifications the base carries that ``head`` dropped while their entry stays."""
    base, head = read_justifications(base_source), read_justifications(head_source)
    deleted: list[str] = []
    for name, (_, pairs) in base.items():
        if name not in head:
            continue
        head_keys, head_pairs = head[name]
        missing = Counter(pairs) - Counter(head_pairs)
        for (owner, text), count in missing.items():
            if owner is not None and owner not in head_keys:
                continue
            owner_text = "" if owner is None else f" for entry {owner}"
            deleted += [
                f"{name} lost the justification {text!r}{owner_text}" for _ in range(count)
            ]
    return deleted


def read_relaxation(source: str) -> object:
    for name, value in _assigned(source):
        if name == RELAXATION_NAME:
            try:
                return ast.literal_eval(value)
            except (ValueError, TypeError):
                return None
    return None


def _is_number(value: object) -> bool:
    return isinstance(value, int | float) and not isinstance(value, bool)


def raised_entries(base: object, head: object) -> list[str]:
    """Describe every way ``head`` is looser than ``base``."""
    if _is_number(base) and _is_number(head):
        return [f"{base!r} -> {head!r}"] if head > base else []
    if isinstance(base, dict) and isinstance(head, dict):
        raised: list[str] = []
        for key, value in head.items():
            if key not in base:
                raised.append(f"new entry {key!r}")
            else:
                raised += [f"{key!r}: {item}" for item in raised_entries(base[key], value)]
        return raised
    if isinstance(base, set | frozenset | list | tuple) and isinstance(
        head, set | frozenset | list | tuple
    ):
        return [f"new element {item!r}" for item in head if item not in base]
    return [] if base == head else [f"value changed {base!r} -> {head!r}"]


def relaxation_problem(head: object, base: object) -> str | None:
    """Why ``head`` is not a valid, new relaxation field; ``None`` when it is."""
    if not isinstance(head, dict):
        return "the script declares no RELAXATION field with 'reason' and 'reference'"
    for key in ("reason", "reference"):
        value = head.get(key)
        if not (isinstance(value, str) and value.strip()):
            return f"RELAXATION needs a non-empty '{key}'"
    actor, issue = head.get("actor"), head.get("issue")
    if not (
        (isinstance(actor, str) and actor.strip())
        or (isinstance(issue, int) and not isinstance(issue, bool) and issue > 0)
    ):
        return "RELAXATION needs a verifiable identity: 'actor' or a positive 'issue' (HR-33)"
    if head == base:
        return "RELAXATION is the one the base already carries; it cannot cover a new change"
    return None


def check_sources(
    path: str,
    base_source: str,
    head_source: str,
    warnings: list[str] | None = None,
) -> list[str]:
    """Violations of one script and warnings for baselines the guard cannot read."""
    base, head = read_baselines(base_source), read_baselines(head_source)
    violations: list[str] = []
    relaxed: list[str] = []
    if warnings is not None:
        unreadable = sorted(
            set(read_unreadable_baselines(base_source))
            | set(read_unreadable_baselines(head_source))
        )
        warnings.extend(
            f"{path}: {name} is not a literal and is not protected by this guard (HR-15)"
            for name in unreadable
        )
    for name, base_value in base.items():
        if name not in head:
            violations.append(f"{path}: {name} is no longer a readable literal")
            continue
        relaxed += [f"{name} {item}" for item in raised_entries(base_value, head[name])]
    relaxed += deleted_justifications(base_source, head_source)
    if relaxed:
        problem = relaxation_problem(read_relaxation(head_source), read_relaxation(base_source))
        if problem is not None:
            violations.append(
                f"{path}: baseline relaxed without justification ({problem}): "
                + "; ".join(relaxed)
                + ". Lower the baseline, restore the dropped justification, or declare RELAXATION = "
                + '{"reason": "...", "reference": "...", "issue": N} in this PR (HR-15).'
            )
    return violations


def _git(*args: str, root: Path) -> str:
    done = subprocess.run(  # noqa: S603 - fixed git argv, no shell
        ["git", *args],  # noqa: S607 - PATH-resolved binary
        cwd=root,
        capture_output=True,
        text=True,
        check=False,
    )
    if done.returncode != 0:
        command = " ".join(("git", *args))
        raise BaseUnreadableError(command, done.stderr.strip())
    return done.stdout


def check_against_base(
    root: Path,
    base_ref: str,
    gate_glob: str,
    warnings: list[str] | None = None,
) -> list[str]:
    merge_base = _git("merge-base", base_ref, "HEAD", root=root).strip()
    changed = _git(
        "diff",
        "--name-only",
        "--diff-filter=M",
        merge_base,
        "HEAD",
        "--",
        gate_glob,
        root=root,
    ).splitlines()
    violations: list[str] = []
    for rel in changed:
        base_source = _git("show", f"{merge_base}:{rel}", root=root)
        head_source = (root / rel).read_text(encoding="utf-8")
        violations += check_sources(rel, base_source, head_source, warnings=warnings)
    return violations


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: output must not depend on the locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--base-ref", default=None)
    parser.add_argument("--gate-glob", default=DEFAULT_GATE_GLOB)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    args = parser.parse_args(argv)
    env_base = os.environ.get("GITHUB_BASE_REF", "")
    base_ref = args.base_ref or (f"origin/{env_base}" if env_base else "")
    if not base_ref:
        print("check_baseline_relaxation: base comparison NOT performed (no base ref).")
        return EXIT_OK
    warnings: list[str] = []
    try:
        violations = check_against_base(
            args.root, base_ref, args.gate_glob, warnings=warnings
        )
    except BaseUnreadableError as exc:
        print(f"FAIL check_baseline_relaxation: cannot read base {base_ref}: {exc}")
        return EXIT_UNREADABLE
    for warning in warnings:
        print(f"WARNING: {warning}", file=sys.stderr)
    for violation in violations:
        print(f"FAIL: {violation}")
    if violations:
        return EXIT_RELAXED
    print(f"check_baseline_relaxation: OK (compared with {base_ref})")
    return EXIT_OK


if __name__ == "__main__":
    sys.exit(main())
