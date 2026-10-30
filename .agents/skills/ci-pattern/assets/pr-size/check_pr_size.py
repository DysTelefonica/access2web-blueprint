#!/usr/bin/env python3
# ci-pattern asset — gate de presupuesto de revisión (#623)
"""PR size gate: review budget, with vendored roots excluded.

Counts semantic changed lines between the merge base and ``HEAD``. Paths
under :data:`VENDORED_ROOTS` are excluded from the budget because they
hold propagated skill content that no human reads line by line. That is
``.agents/skills/`` only: the vendored tree a consumer receives from the
canon. ``skills/`` counts, because it is the authoring catalog of the
repo (#1088): in the canon it holds the 64 skills written here, and in a
consumer it holds that consumer's own catalog. A rename FROM a
non-vendored (authored) path INTO a vendored path still
counts: the work was authored outside the vendored subtree and the
budget has to cover it. A rename within the vendored subtree, or a
new file under it, is excluded.

Line-ending churn is not a change a reviewer has to read, so
``--ignore-cr-at-eol`` is enabled by default. The generated lockfiles
(``uv.lock``, ``package-lock.json``, ``pnpm-lock.yaml``, ``poetry.lock`` and
``requirements*.lock``) are excluded too: the package manager writes them,
not the author (#1177). Nothing else is: any other generated file counts (the
budget is additions + deletions).

Two ways to measure. Locally and in a full checkout the diff against the
base comes from ``git``. With ``--event <pull_request event json>`` the
per-file counts come from the GitHub API (``GITHUB_TOKEN``, optional
``GITHUB_API_URL``: https, or http on loopback only), so the runner needs no
history; any API failure, or a breakdown that still does not match the PR's
``changed_files`` after one re-read, fails closed. If the PR's head moved after
the event (a newer push), the event is superseded: it exits 0, because the newer
event measures the new head (#1113).

Not a CI gate (#1396). In CI the budget is enforced once, by rule A1 of
``assets/pr-contract/check_pr_contract.py``, with one count: running both gave
two red checks with two different numbers on the same PR. This script stays as
a local measurement and as the size diagnostic of ``pr preflight --draft``.

The only override is the ``size:exception`` label, applied by the
maintainer (never by the author, never read from the PR body). With
``--event`` it is read live from the API (``GET issues/N/labels``), because the
event payload is a snapshot from the moment the event fired: a label added right
after the ``opened`` event (cadete#715) would otherwise be missed (#1359). Local
runs take it from ``--labels`` or from ``$PR_LABELS`` (comma separated).

Exit codes:
    0  within budget, or overridden by the ``size:exception`` label
    1  over budget, or unable to measure (fail closed)
    2  usage / missing arguments
"""

from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
import urllib.parse
import urllib.request
from pathlib import Path

# Fuente única del presupuesto (#931): la regla la comparte con pr-contract.
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "budget"))
from vendored_budget import (  # noqa: E402
    LOCKFILE_PATTERNS as LOCKFILE_PATTERNS,  # reexport: sigue siendo nombre público del gate
)
from vendored_budget import (
    VENDORED_ROOTS as VENDORED_ROOTS,  # reexport: sigue siendo nombre público del gate
)
from vendored_budget import (
    GeneratedPatterns,
    _strip_diff_prefix,
    counts_to_budget,
    generated_patterns,
    is_generated,
    is_under_vendored,
    original_path,
    parse_numstat,
)
from vendored_budget import (
    is_generated_lockfile as is_generated_lockfile,  # reexport: sigue siendo nombre público del gate
)

# --------------------------------------------------------------------------------------------
# CONFIGURATION
# --------------------------------------------------------------------------------------------

MAX_CHANGED_LINES = 400

#: The only override: applied by the maintainer, never read from the PR body.
OVERRIDE_LABEL = "size:exception"

FILES_PER_PAGE = 100
#: GitHub lists at most 3000 files of a PR.
MAX_FILE_PAGES = 30
#: Labels are read page by page too (#1359); a PR carries a handful, so one page is the norm.
MAX_LABEL_PAGES = 10
#: Reads of the PR's breakdown before the gate gives up (#1113: one retry).
MEASURE_ATTEMPTS = 2
DEFAULT_API_URL = "https://api.github.com"

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


def base_gitattributes(base_ref: str) -> str:
    """The base's ``.gitattributes`` (#1440), or ``""`` when the base has none.

    Read from the BASE, never the working tree: a PR cannot take its own lines
    out of the budget by marking them ``linguist-generated``.
    """
    try:
        return _git("show", f"{base_ref}:.gitattributes")
    except RuntimeError:
        return ""


def changed_lines(base_ref: str, generated_rules: GeneratedPatterns | None = None,
                  ) -> tuple[int, int, int, list[tuple[str, int]]]:
    """Return ``(counted_total, vendored_excluded, generated_excluded, per_file)``.

    ``counted_total`` is the review-budget figure (vendored and generated
    lockfiles excluded). ``vendored_excluded`` is the number of lines a naive
    gate would have counted but the vendored rule dropped, and
    ``generated_excluded`` the ones the generated-lockfile rule dropped
    (#1177); they are reported apart so the message says what it left out.
    ``per_file`` is the counted breakdown, largest first.
    """
    if generated_rules is None:
        generated_rules = generated_patterns(base_gitattributes(base_ref))
    raw = _git("diff", "--numstat", "--find-renames=50%", "--ignore-cr-at-eol",
               f"{base_ref}...HEAD")
    per_file: list[tuple[str, int]] = []
    counted = 0
    vendored = 0
    generated = 0
    for line in raw.splitlines():
        parts = line.split("\t")
        if len(parts) != 3:
            continue
        added, deleted, path = parts
        if added == "-" or deleted == "-":  # binary file, not reviewable line by line
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
            origin = original_path(new_path, base_ref, _git)
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
        if counts_to_budget(old_path, new_path, generated_rules):
            counted += lines
            per_file.append((new_path, lines))
        elif is_generated(new_path, generated_rules):
            generated += lines
        else:
            vendored += lines
    per_file.sort(key=lambda item: item[1], reverse=True)
    return counted, vendored, generated, per_file


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


def has_override(labels: set[str]) -> bool:
    """True when the maintainer applied the ``size:exception`` label."""
    return OVERRIDE_LABEL in labels


def split_labels(raw: str) -> set[str]:
    """Labels from a comma- or newline-separated string."""
    return {item.strip() for item in re.split(r"[,\n]", raw) if item.strip()}


class MeasureError(RuntimeError):
    """The size cannot be measured; the gate fails closed (exit 1)."""


class EventUnreadableError(MeasureError):
    def __init__(self, path: str, exc: BaseException) -> None:
        super().__init__(f"cannot read the event payload {path}: {exc}")


class EventShapeError(MeasureError):
    def __init__(self) -> None:
        super().__init__("the event payload has no pull_request number or repository")


class MissingCredentialError(MeasureError):
    def __init__(self) -> None:
        super().__init__("GITHUB_TOKEN is not set: the API breakdown cannot be read")


class UnsupportedApiUrlError(MeasureError):
    def __init__(self) -> None:
        super().__init__("GITHUB_API_URL must be https (http only on loopback)")


class ApiFailedError(MeasureError):
    def __init__(self, path: str, exc: BaseException) -> None:
        super().__init__(f"GET {path} failed: {exc}")


class ApiShapeError(MeasureError):
    def __init__(self, path: str) -> None:
        super().__init__(f"GET {path} did not return a file list")


class MalformedFileError(MeasureError):
    def __init__(self, item: object) -> None:
        super().__init__(f"malformed file entry: {item!r}")


class BreakdownMismatchError(MeasureError):
    def __init__(self, got: int, want: int) -> None:
        super().__init__(f"the file breakdown ({got}) does not match the PR's changed_files ({want})")


class GitattributesFileError(RuntimeError):
    """``--gitattributes-file`` could not be read as UTF-8 text (#1440)."""

    def __init__(self, path: str, cause: Exception) -> None:
        super().__init__(f"--gitattributes-file unreadable: {path}: {cause}")


class SupersededError(Exception):
    """The PR got a newer push after this event: its own event measures the new head (#1113).

    Not a MeasureError: it is not a failure of the gate, so it exits 0.
    """

    def __init__(self, event_head: str, head: str) -> None:
        super().__init__(f"superado por un push más reciente (head del evento {event_head[:12]}, "
                         f"head actual {head[:12]})")


def is_int(value: object) -> bool:
    return isinstance(value, int) and not isinstance(value, bool)


def is_file_list(value: object) -> bool:
    return isinstance(value, list) and all(isinstance(item, dict) for item in value)


def read_event(path: str) -> tuple[str, int, set[str], int | None]:
    """``(repository, number, labels, changed_files)`` from a pull_request event payload."""
    try:
        payload = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise EventUnreadableError(path, exc) from exc
    pull = payload.get("pull_request") if isinstance(payload, dict) else None
    repo = (payload.get("repository") or {}).get("full_name") if isinstance(payload, dict) else None
    number = pull.get("number") if isinstance(pull, dict) else None
    if not isinstance(repo, str) or not is_int(number):
        raise EventShapeError
    labels = {lab.get("name") for lab in pull.get("labels") or [] if isinstance(lab, dict)}
    declared = pull.get("changed_files")
    return repo, number, {name for name in labels if isinstance(name, str)}, (
        declared if is_int(declared) else None)


def _open(api_url: str, path: str, token: str):
    """GET ``path`` with the scheme validated: https, or http on loopback only.

    Each branch carries its literal scheme inside the call: ruff S310 only accepts a
    URL with a literal scheme.
    """
    parts = urllib.parse.urlsplit(api_url)
    tail = parts.path.rstrip("/") + path
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/vnd.github+json"}
    if parts.scheme == "https":
        return urllib.request.urlopen(
            urllib.request.Request(f"https://{parts.netloc}{tail}", headers=headers), timeout=30)
    if parts.scheme == "http" and parts.hostname in ("127.0.0.1", "localhost"):
        return urllib.request.urlopen(
            urllib.request.Request(f"http://{parts.netloc}{tail}", headers=headers), timeout=30)
    raise UnsupportedApiUrlError


def api_files(api_url: str, token: str, repo: str, number: int) -> list[dict]:
    """Every file of the PR, page by page (a short page ends the listing)."""
    files: list[dict] = []
    for page in range(1, MAX_FILE_PAGES + 1):
        path = f"/repos/{repo}/pulls/{number}/files?per_page={FILES_PER_PAGE}&page={page}"
        try:
            with _open(api_url, path, token) as response:
                chunk = json.loads(response.read().decode("utf-8"))
        except (OSError, ValueError) as exc:
            raise ApiFailedError(path, exc) from exc
        if not is_file_list(chunk):
            raise ApiShapeError(path)
        files.extend(chunk)
        if len(chunk) < FILES_PER_PAGE:
            break
    return files


def count_api_files(files: list[dict], generated_rules: GeneratedPatterns = (),
                    ) -> tuple[int, int, int, list[tuple[str, int]]]:
    """Same figures as :func:`changed_lines`, from the API's per-file entries."""
    per_file: list[tuple[str, int]] = []
    counted = vendored = generated = 0
    for item in files:
        name, previous = item.get("filename"), item.get("previous_filename") or item.get("filename")
        added, deleted = item.get("additions"), item.get("deletions")
        if not isinstance(name, str) or not is_int(added) or not is_int(deleted):
            raise MalformedFileError(item)
        lines = added + deleted
        if counts_to_budget(previous, name, generated_rules):
            counted += lines
            per_file.append((name, lines))
        elif is_generated(name, generated_rules):
            generated += lines
        else:
            vendored += lines
    per_file.sort(key=lambda entry: entry[1], reverse=True)
    return counted, vendored, generated, per_file


def event_head(event_path: str) -> str | None:
    """``pull_request.head.sha`` from the event payload, when the payload carries it."""
    try:
        payload = json.loads(Path(event_path).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    pull = payload.get("pull_request") if isinstance(payload, dict) else None
    head = pull.get("head") if isinstance(pull, dict) else None
    sha = head.get("sha") if isinstance(head, dict) else None
    return sha if isinstance(sha, str) and sha else None


def api_labels(api_url: str, token: str, repo: str, number: int) -> set[str]:
    """The PR's labels right now, page by page (``GET issues/N/labels``) (#1359).

    Not the event payload: that is a snapshot of the moment the event fired.
    """
    names: set[str] = set()
    for page in range(1, MAX_LABEL_PAGES + 1):
        path = f"/repos/{repo}/issues/{number}/labels?per_page={FILES_PER_PAGE}&page={page}"
        try:
            with _open(api_url, path, token) as response:
                chunk = json.loads(response.read().decode("utf-8"))
        except (OSError, ValueError) as exc:
            raise ApiFailedError(path, exc) from exc
        if not is_file_list(chunk):
            raise ApiShapeError(path)
        names.update(item["name"] for item in chunk if isinstance(item.get("name"), str))
        if len(chunk) < FILES_PER_PAGE:
            break
    return names


def api_pull(api_url: str, token: str, repo: str, number: int) -> tuple[str | None, int | None]:
    """``(head sha, changed_files)`` of the PR right now (``GET pulls/N``)."""
    path = f"/repos/{repo}/pulls/{number}"
    try:
        with _open(api_url, path, token) as response:
            body = json.loads(response.read().decode("utf-8"))
    except (OSError, ValueError) as exc:
        raise ApiFailedError(path, exc) from exc
    if not isinstance(body, dict):
        raise ApiShapeError(path)
    head = body.get("head")
    sha = head.get("sha") if isinstance(head, dict) else None
    changed = body.get("changed_files")
    return (sha if isinstance(sha, str) else None), (changed if is_int(changed) else None)


def api_changed_lines(event_path: str, generated_rules: GeneratedPatterns = (),
                      ) -> tuple[int, int, int, list[tuple[str, int]], set[str]]:
    """``changed_lines`` for a pull_request event, plus the PR's labels read live (#1359).

    The PR's head and its breakdown are read around the file listing, so a push or a
    moved base between the reads shows up as a mismatch instead of a false failure
    (#1113): a moved head means a newer event measures it (superseded, exit 0); a
    mismatch with the head unchanged is retried once, then fails closed.
    """
    repo, number, _payload_labels, _declared = read_event(event_path)
    token = os.environ.get("GITHUB_TOKEN", "")
    if not token:
        raise MissingCredentialError
    api_url = os.environ.get("GITHUB_API_URL") or DEFAULT_API_URL
    expected_head = event_head(event_path)
    got = want = None
    for _ in range(MEASURE_ATTEMPTS):
        head, want = api_pull(api_url, token, repo, number)
        if expected_head and head and head != expected_head:
            raise SupersededError(expected_head, head)
        files = api_files(api_url, token, repo, number)
        head_after, want = api_pull(api_url, token, repo, number)
        if expected_head and head_after and head_after != expected_head:
            raise SupersededError(expected_head, head_after)
        if want is None or len(files) == want:
            counted, vendored, generated, per_file = count_api_files(files, generated_rules)
            return counted, vendored, generated, per_file, api_labels(api_url, token, repo, number)
        got = len(files)
    raise BreakdownMismatchError(got, want)


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


def measure(args: argparse.Namespace) -> tuple[int, int, int, list[tuple[str, int]], set[str], str]:
    """``(counted, vendored, generated, per_file, labels, source)`` from the API (``--event``) or from git."""
    rules = None
    if args.gitattributes_file:
        try:
            text = Path(args.gitattributes_file).read_text(encoding="utf-8")
        except (OSError, UnicodeDecodeError) as exc:
            raise GitattributesFileError(args.gitattributes_file, exc) from exc
        rules = generated_patterns(text)
    if args.event:
        return (*api_changed_lines(args.event, rules or ()), "the GitHub API")
    base_ref = args.base_ref or os.environ.get("BASE_REF") or pr_base_ref() or "origin/main"
    counted, vendored, generated, per_file = changed_lines(base_ref, rules)
    return (counted, vendored, generated, per_file,
            split_labels(args.labels or os.environ.get("PR_LABELS", "")), base_ref)


def envelope_for(counted: int, vendored: int, generated: int, per_file: list,
                 overridden: bool) -> dict:
    passed = counted <= MAX_CHANGED_LINES or overridden
    return {
        "gate": "pr_size",
        "status": "pass" if passed else "fail",
        "indicators": {
            "changed_lines": counted,
            "vendored_lines": vendored,
            "generated_lines": generated,
            "files_changed": len(per_file),
            "overridden": overridden,
        },
        "ceilings": {"changed_lines": MAX_CHANGED_LINES},
        "findings": [
            {"file": path, "line": 0, "detail": f"{count} changed line(s)"}
            for path, count in per_file[:10]
        ] if not passed else [],
    }


def excluded_suffix(vendored: int, generated: int) -> str:
    """El sufijo del mensaje: lo que la regla dejó fuera, nombrado aparte.

    El árbol vendorizado y los lockfiles generados son cosas distintas (HR-13):
    el mensaje no los mezcla ni los llama «vendored» a los segundos (#1177).
    """
    parts = []
    if vendored:
        parts.append(f"{vendored} vendored line(s) excluded")
    if generated:
        parts.append(f"{generated} generated line(s) excluded")
    return f", {', '.join(parts)}" if parts else ""


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--base-ref",
        default=None,
        help="ref to diff against (default: $BASE_REF, else the PR base from gh, else origin/main)",
    )
    parser.add_argument("--event", default=None,
                        help="pull_request event JSON: count through the GitHub API instead of git")
    parser.add_argument("--labels", default=None,
                        help="comma-separated PR labels (default: $PR_LABELS; --event reads the payload)")
    parser.add_argument("--json", action="store_true", help="emit the indicator envelope")
    parser.add_argument("--gitattributes-file", default=None,
                        help="the base's .gitattributes (#1440): paths marked linguist-generated "
                             "do not count (git mode reads the base's own by default)")
    args = parser.parse_args(argv)

    try:
        counted, vendored, generated, per_file, labels, source = measure(args)
    except SupersededError as exc:
        if args.json:
            print(json.dumps({"gate": "pr_size", "status": "superseded", "detail": str(exc)}))
        else:
            print(f"OK    PR {exc}; mide el evento más reciente, no este")
        return 0
    except RuntimeError as exc:
        if args.json:
            print(json.dumps({"gate": "pr_size", "status": "error", "detail": str(exc)}))
        else:
            print(f"FAIL  {exc}", file=sys.stderr)
        return 1

    overridden = has_override(labels)
    if args.json:
        print(json.dumps(envelope_for(counted, vendored, generated, per_file, overridden), indent=2))
        return 0 if counted <= MAX_CHANGED_LINES or overridden else 1

    suffix = excluded_suffix(vendored, generated)
    if counted <= MAX_CHANGED_LINES:
        print(f"OK    {counted} changed line(s) vs {source}, budget is {MAX_CHANGED_LINES}{suffix}")
        return 0
    if overridden:
        print(f"OK    {counted} changed line(s) vs {source} over budget{suffix}, overridden by the "
              f"'{OVERRIDE_LABEL}' label")
        return 0

    print(f"FAIL  {counted} changed line(s) vs {source}, budget is {MAX_CHANGED_LINES}{suffix}")
    for path, count in per_file[:10]:
        print(f"        {count:>6}  {path}")
    print(f"        split the PR; only the maintainer can apply the '{OVERRIDE_LABEL}' label")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
