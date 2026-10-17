#!/usr/bin/env python3
# ci-pattern asset — gate de presupuesto de revisión (#623)
"""PR size gate: review budget, with vendored roots excluded.

Counts semantic changed lines between the merge base and ``HEAD``. Paths
under :data:`VENDORED_ROOTS` are excluded from the budget because they
hold propagated skill content that no human reads line by line. A
rename FROM a non-vendored (authored) path INTO a vendored path still
counts: the work was authored outside the vendored subtree and the
budget has to cover it. A rename within the vendored subtree, or a
new file under it, is excluded.

Line-ending churn is not a change a reviewer has to read, so
``--ignore-cr-at-eol`` is enabled by default.

The override is deliberately awkward. ``size:exception`` alone does
nothing; it must be paired with a reason line, because an override
nobody has to justify is not an override, it is the default.

Exit codes:
    0  within budget, or overridden with a stated reason
    1  over budget, or overridden without a reason
    2  usage / missing arguments / unreadable repo state
"""

from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys

# --------------------------------------------------------------------------------------------
# CONFIGURATION
# --------------------------------------------------------------------------------------------

MAX_CHANGED_LINES = 400

#: Vendored roots are excluded from the review budget. They hold propagated
#: skill content (``.agents/skills/`` for the consumer-side install of the
#: catalog, ``skills/`` for legacy mirrors); a reviewer does not read them
#: line by line. A rename from an authored path into a vendored root still
#: counts (see :func:`counts_to_budget`).
VENDORED_ROOTS = (".agents/skills/", "skills/")

#: Generated files a human never reads line by line. Every entry weakens the
#: gate, so keep the list short and justify additions in review.
EXCLUDED_SUFFIXES = ("uv.lock", "poetry.lock", "package-lock.json", "pnpm-lock.yaml")

OVERRIDE_MARKER = "size:exception"
OVERRIDE_REASON = re.compile(r"^\s*size-exception-reason:\s*(?P<reason>\S.*)$", re.MULTILINE)

#: Git numstat rename format: ``{old => new}/rest`` or ``{old => new}``.
#: Captures the unchanged prefix (``a/`` / ``b/``) when present, the
#: arrow, the unchanged suffix, and the new path.
_RENAME_RE = re.compile(
    r"^(?:\{(?P<a>[^}]*?) => (?P<b>[^}]*?)\}(?P<rest>.*))|(?P<plain>.+)$"
)

# --------------------------------------------------------------------------------------------
# PURE HELPERS (no subprocess, easy to unit-test)
# --------------------------------------------------------------------------------------------


def is_under_vendored(path: str) -> bool:
    """Return True when ``path`` is fully under one of :data:`VENDORED_ROOTS`."""
    return any(path.startswith(root) for root in VENDORED_ROOTS)


def counts_to_budget(old_path: str, new_path: str) -> bool:
    """Decide whether a file's changes count toward the review budget.

    Rules (issue #623, regla de BP portada a ci-pattern):

    1. A path whose new location is NOT under a vendored root → counts.
    2. A path whose new location IS under a vendored root AND that was
       not authored outside vendored (i.e. it is new, or was renamed
       from inside the vendored subtree) → does NOT count.
    3. A rename FROM an authored path INTO a vendored root → counts,
       because the work was authored outside the vendored subtree.
    """
    if not is_under_vendored(new_path):
        return True
    if old_path != new_path and not is_under_vendored(old_path):
        return True
    return False


def parse_numstat(path_field: str) -> tuple[str, str] | None:
    """Parse a numstat path field into ``(old_path, new_path)``.

    Plain paths collapse to ``(path, path)``. Git emits a rename in one
    of two forms:

    * ``{old => new}`` — pure rename with no shared suffix.
    * ``{old => new}/rest`` — rename with a shared suffix that is
      identical on both sides (e.g. ``{app => .agents/skills/foo}/auth.py``
      means ``app/auth.py`` → ``.agents/skills/foo/auth.py``).

    Both forms resolve to ``(old_path, new_path)``. Returns ``None`` for
    inputs the parser cannot classify.
    """
    if not path_field or path_field == "-":
        return None
    if path_field.startswith("{") and " => " in path_field and path_field.endswith("}"):
        arrow = path_field.index(" => ")
        old_path = path_field[1:arrow]
        new_path = path_field[arrow + len(" => "):-1]
        return old_path, new_path
    if path_field.startswith("{") and " => " in path_field and "}" in path_field:
        # Rename with a shared suffix: ``{old => new}/rest``. The ``}``
        # marks where the rename notation ends; everything after it is
        # a common suffix that belongs to BOTH old and new paths.
        close_brace = path_field.index("}")
        arrow = path_field.index(" => ")
        common_suffix = path_field[close_brace + 1:]
        old_path = path_field[1:arrow] + common_suffix
        new_path = path_field[arrow + len(" => "):close_brace] + common_suffix
        return old_path, new_path
    return path_field, path_field


# --------------------------------------------------------------------------------------------
# MECHANISM
# --------------------------------------------------------------------------------------------


def _git(*args: str) -> str:
    result = subprocess.run(
        ["git", *args], capture_output=True, text=True, check=False, encoding="utf-8"
    )
    if result.returncode != 0:
        raise RuntimeError(f"git {' '.join(args)} failed: {result.stderr.strip()}")
    return result.stdout


def _strip_diff_prefix(path: str) -> str:
    """Strip the ``a/`` / ``b/`` prefix git uses in diff output."""
    if path.startswith("a/") or path.startswith("b/"):
        return path[2:]
    return path


def original_path(new_path: str, base_ref: str) -> str | None:
    """Return the path where the file at HEAD's ``new_path`` was first
    added in the range, or ``None`` if the history is unknown.

    Cumulative ``git diff main...HEAD`` collapses cross-commit renames
    into a single addition at the final path, so it cannot tell that an
    authored file became vendored. ``git log --follow`` follows the
    rename chain back to the original commit; we read that output in
    reverse and keep the first A or the first pre-rename path we see.

    A vendored file that has no pre-rename path in the range was added
    inside the vendored subtree and is not authored work.
    """
    try:
        raw = _git("log", "--follow", "--name-status", "--pretty=format:",
                   f"{base_ref}..HEAD", "--", new_path)
    except RuntimeError:
        return None
    entries = [line for line in raw.splitlines() if line.strip()]
    if not entries:
        return None
    # Output is newest-first. The earliest entry is the file's origin.
    for entry in reversed(entries):
        parts = entry.split("\t")
        if not parts:
            continue
        status = parts[0]
        if status == "A" and len(parts) == 2:
            return _strip_diff_prefix(parts[1])
        if status.startswith("R") and len(parts) == 3:
            return _strip_diff_prefix(parts[1])
        if status.startswith("C") and len(parts) == 3:
            return _strip_diff_prefix(parts[1])
        # M (modify) and D (delete) do not move the origin.
    return None


def changed_lines(base_ref: str) -> tuple[int, int, list[tuple[str, int]]]:
    """Return ``(counted_total, vendored_excluded, per_file)``.

    ``counted_total`` is the review-budget figure (vendored excluded).
    ``vendored_excluded`` is the number of lines a naive gate would have
    counted but the vendored rule dropped. ``per_file`` is the counted
    breakdown, largest first.
    """
    raw = _git("diff", "--numstat", "--find-renames=50%", "--ignore-cr-at-eol",
               f"{base_ref}...HEAD")
    per_file: list[tuple[str, int]] = []
    counted = 0
    vendored = 0
    for line in raw.splitlines():
        parts = line.split("\t")
        if len(parts) != 3:
            continue
        added, deleted, path = parts
        if added == "-" or deleted == "-":  # binary file, not reviewable line by line
            continue
        if path.endswith(EXCLUDED_SUFFIXES):
            continue
        parsed = parse_numstat(path)
        if parsed is None:
            continue
        old_path, new_path = parsed
        old_path = _strip_diff_prefix(old_path)
        new_path = _strip_diff_prefix(new_path)
        # The diff may have collapsed a cross-commit rename into a
        # single addition at the final path. ``parse_numstat`` cannot
        # tell that the authored origin was outside vendored, so ask
        # git's rename history directly: if the file was first added
        # at a non-vendored path, the work is authored and counts.
        if is_under_vendored(new_path):
            origin = original_path(new_path, base_ref)
            if origin is not None and not is_under_vendored(origin):
                # Use the cumulative numstat lines (additions + deletions)
                # as the figure that counts: the file was carried over a
                # rename, so ``int(added) + int(deleted)`` already reflects
                # only the new work. When the cumulative diff collapsed the
                # rename into a single ``A`` line, the additions column
                # holds the file's HEAD size, which is the correct figure.
                lines = int(added) + int(deleted)
                counted += lines
                per_file.append((new_path, lines))
                continue
        lines = int(added) + int(deleted)
        if counts_to_budget(old_path, new_path):
            counted += lines
            per_file.append((new_path, lines))
        else:
            vendored += lines
    per_file.sort(key=lambda item: item[1], reverse=True)
    return counted, vendored, per_file


def pr_base_ref() -> str | None:
    """Return ``origin/<baseRefName>`` of the PR for the current branch, or None.

    Asks ``gh`` for the same base CI reads from ``github.base_ref``, so a
    stacked PR is measured against its real base and not against
    ``origin/main``. No PR, no ``gh``, or a base that is not fetched locally
    returns None, and the caller falls back to ``origin/main``.
    """
    try:
        found = subprocess.run(
            ["gh", "pr", "view", "--json", "baseRefName"],
            capture_output=True, text=True, check=False, encoding="utf-8",
        )
    except FileNotFoundError:
        return None
    if found.returncode != 0:
        return None
    try:
        name = json.loads(found.stdout).get("baseRefName")
    except (json.JSONDecodeError, AttributeError):
        return None
    if not isinstance(name, str) or not name:
        return None
    ref = f"origin/{name}"
    verify = subprocess.run(
        ["git", "rev-parse", "--verify", "--quiet", ref],
        capture_output=True, check=False,
    )
    return ref if verify.returncode == 0 else None


def override_reason(pr_body: str) -> str | None:
    """Return the stated reason when a valid override is present, else ``None``."""
    if OVERRIDE_MARKER not in pr_body:
        return None
    match = OVERRIDE_REASON.search(pr_body)
    return match.group("reason").strip() if match else None


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8.

    Python picks the output encoding from the platform locale, so the same
    gate emits different bytes on a Windows workstation (cp1252) and a Linux
    runner (utf-8) — and a non-encodable character crashes the write outright.
    A harness that claims determinism cannot let its own output depend on
    where it ran.
    """
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--base-ref",
        default=None,
        help="ref to diff against (default: $BASE_REF, else the PR base from gh, else origin/main)",
    )
    parser.add_argument("--json", action="store_true", help="emit the indicator envelope")
    args = parser.parse_args(argv)
    base_ref = args.base_ref or os.environ.get("BASE_REF") or pr_base_ref() or "origin/main"

    try:
        counted, vendored, per_file = changed_lines(base_ref)
    except RuntimeError as exc:
        if args.json:
            print(json.dumps({"gate": "pr_size", "status": "error", "detail": str(exc)}))
        else:
            print(f"FAIL  {exc}", file=sys.stderr)
        return 1

    if args.json:
        reason = override_reason(os.environ.get("PR_BODY", ""))
        passed = counted <= MAX_CHANGED_LINES or bool(reason)
        envelope = {
            "gate": "pr_size",
            "status": "pass" if passed else "fail",
            "indicators": {
                "changed_lines": counted,
                "vendored_lines": vendored,
                "files_changed": len(per_file),
                "overridden": bool(reason),
            },
            "ceilings": {"changed_lines": MAX_CHANGED_LINES},
            "findings": [
                {"file": path, "line": 0, "detail": f"{count} changed line(s)"}
                for path, count in per_file[:10]
            ] if not passed else [],
        }
        if reason and not passed:
            envelope["override_reason"] = reason
        print(json.dumps(envelope, indent=2))
        return 0 if passed else 1

    if counted <= MAX_CHANGED_LINES:
        suffix = (
            f", {vendored} vendored line(s) excluded"
            if vendored else ""
        )
        print(f"OK    {counted} changed line(s) vs {base_ref}, budget is {MAX_CHANGED_LINES}{suffix}")
        return 0

    pr_body = os.environ.get("PR_BODY", "")
    reason = override_reason(pr_body)
    if reason:
        print(
            f"OK    {counted} changed line(s) vs {base_ref} over budget, "
            f"{vendored} vendored line(s) excluded, overridden: {reason}"
        )
        return 0

    print(f"FAIL  {counted} changed line(s) vs {base_ref}, budget is {MAX_CHANGED_LINES}")
    if vendored:
        print(f"        ({vendored} vendored line(s) excluded)")
    for path, count in per_file[:10]:
        print(f"        {count:>6}  {path}")
    if OVERRIDE_MARKER in pr_body:
        print(
            f"        '{OVERRIDE_MARKER}' found but no 'size-exception-reason: <why>' line in the "
            f"PR body"
        )
    else:
        print(
            f"        split the PR, or add '{OVERRIDE_MARKER}' plus a "
            f"'size-exception-reason: <why>' line to the PR body"
        )
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
