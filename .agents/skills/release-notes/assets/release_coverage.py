#!/usr/bin/env python3
"""Coverage check for release notes (DysTelefonica/team-skills#1425).

Usage:
  release_coverage.py collect --repo-dir <clone> --from <tag> --to <tag|sha>
                              [--github <owner/name>] [--prs-json <file>] --out <range.json>
  release_coverage.py check --range <range.json> --notes <notes.md>
                            [--exclude-pr N]... [--exclude-issue N]...

``collect`` reads the git range of a local clone and the repository's pull requests
(``gh pr list``, or a saved dump passed with ``--prs-json``) and writes the range file:

* ``prs``: every PR whose head or merge commit is in the range. ``direct`` is true when
  its merge commit is on the first-parent line of ``--to``; otherwise ``integrated_by``
  names the first direct PR whose merge brought its head in (the tip of its chain).
* ``direct_commits``: first-parent commits that are not the merge of a PR (pushes
  without a PR, the usual case in VBA/Access repositories).
* ``issues``: the issues closed by the direct PRs (GitHub closing references).

``check`` fails (exit 1) naming every PR, commit or issue the notes do not account for:
a direct PR must be cited anywhere (``#N``); a chain intermediate is covered when it or
its tip is cited; a direct commit is covered when an issue or PR in its subject, or its
short SHA, is cited; an issue must be cited inside the «Registro de cierres» section.
Exit 2 is an input error. Standard library only.
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from collections.abc import Callable, Iterable
from pathlib import Path

PR_FIELDS = "number,title,state,baseRefName,headRefName,headRefOid,mergeCommit,closingIssuesReferences"
REGISTER_HEADING = "registro de cierres"
CITATION_RE = re.compile(r"(?<![\w&/])#(\d+)\b")
HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")


class InputError(Exception):
    """Bad input: the command cannot produce a verdict."""


# ---------------------------------------------------------------------------
# collect


def _git(repo: Path, *args: str) -> str:
    result = subprocess.run(["git", "-C", str(repo), *args], capture_output=True, text=True,
                            encoding="utf-8", check=False)
    if result.returncode != 0:
        raise InputError(f"git {' '.join(args)}: {result.stderr.strip()}")
    return result.stdout


def classify(range_shas: set[str], first_parent: list[str], prs: Iterable[dict],
             is_ancestor: Callable[[str, str], bool]) -> list[dict]:
    """In-range PRs, each marked direct or integrated by the tip that brought it in.

    ``first_parent`` is oldest first.
    """
    first_parent_set = set(first_parent)
    in_range = []
    for pr in prs:
        head = pr.get("headRefOid") or ""
        merge = (pr.get("mergeCommit") or {}).get("oid") or ""
        if head in range_shas or merge in range_shas:
            in_range.append(pr)
    direct_by_merge = {}
    for pr in in_range:
        merge = (pr.get("mergeCommit") or {}).get("oid") or ""
        if pr.get("state") == "MERGED" and merge in first_parent_set:
            direct_by_merge[merge] = pr["number"]
    records = []
    for pr in sorted(in_range, key=lambda p: p["number"]):
        merge = (pr.get("mergeCommit") or {}).get("oid") or ""
        direct = direct_by_merge.get(merge) == pr["number"]
        integrated_by = None
        if not direct:
            head = pr.get("headRefOid") or ""
            for sha in first_parent:
                if sha in direct_by_merge and head and is_ancestor(head, sha):
                    integrated_by = direct_by_merge[sha]
                    break
        records.append({
            "number": pr["number"],
            "title": pr.get("title", ""),
            "state": pr.get("state", ""),
            "base": pr.get("baseRefName", ""),
            "direct": direct,
            "integrated_by": integrated_by,
            "closes": sorted({i["number"] for i in pr.get("closingIssuesReferences") or []}),
        })
    return records


def collect(repo: Path, from_ref: str, to_ref: str, prs: list[dict]) -> dict:
    for ref in (from_ref, to_ref):
        _git(repo, "rev-parse", "--verify", "--quiet", f"{ref}^{{commit}}")
    range_shas = set(_git(repo, "rev-list", f"{from_ref}..{to_ref}").split())
    first_parent = _git(repo, "rev-list", "--reverse", "--first-parent",
                        f"{from_ref}..{to_ref}").split()

    def is_ancestor(a: str, b: str) -> bool:
        return subprocess.run(["git", "-C", str(repo), "merge-base", "--is-ancestor", a, b],
                              capture_output=True, check=False).returncode == 0

    records = classify(range_shas, first_parent, prs, is_ancestor)
    pr_merges = {(p.get("mergeCommit") or {}).get("oid") for p in prs}
    direct_commits = []
    for sha in first_parent:
        if sha not in pr_merges:
            subject = _git(repo, "log", "-1", "--format=%s", sha).strip()
            direct_commits.append({"sha": sha, "subject": subject})
    issues = sorted({n for r in records if r["direct"] for n in r["closes"]})
    return {
        "from": from_ref,
        "to": to_ref,
        "from_sha": _git(repo, "rev-parse", f"{from_ref}^{{commit}}").strip(),
        "to_sha": _git(repo, "rev-parse", f"{to_ref}^{{commit}}").strip(),
        "prs": records,
        "direct_commits": direct_commits,
        "issues": [{"number": n} for n in issues],
    }


def _load_prs(args: argparse.Namespace, repo: Path) -> list[dict]:
    if args.prs_json:
        return _read_json(Path(args.prs_json))
    if not args.github:
        raise InputError("collect necesita --github <owner/name> o --prs-json <fichero>")
    since = _git(repo, "log", "-1", "--format=%cI", f"{args.from_ref}^{{commit}}").strip()[:10]
    result = subprocess.run(
        ["gh", "pr", "list", "--repo", args.github, "--state", "all", "--limit", "1000",
         "--search", f"updated:>={since}", "--json", PR_FIELDS],
        capture_output=True, text=True, encoding="utf-8", check=False)
    if result.returncode != 0:
        raise InputError(f"gh pr list: {result.stderr.strip()}")
    return json.loads(result.stdout)


# ---------------------------------------------------------------------------
# check


def _read_json(path: Path) -> list | dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise InputError(f"no se puede leer {path}: {exc}") from exc


def citations(text: str) -> set[int]:
    return {int(n) for n in CITATION_RE.findall(text)}


def register_section(notes: str) -> str | None:
    """Text of the «Registro de cierres» section, subsections included."""
    lines = notes.splitlines()
    for start, line in enumerate(lines):
        match = HEADING_RE.match(line)
        if match and REGISTER_HEADING in match.group(2).lower():
            level = len(match.group(1))
            body = []
            for other in lines[start + 1:]:
                nxt = HEADING_RE.match(other)
                if nxt and len(nxt.group(1)) <= level:
                    break
                body.append(other)
            return "\n".join(body)
    return None


def check(data: dict, notes: str, exclude_prs: set[int], exclude_issues: set[int]
          ) -> tuple[list[str], list[str]]:
    """Return (missing, remarks)."""
    cited = citations(notes)
    missing: list[str] = []
    remarks: list[str] = []
    for pr in data.get("prs", []):
        number = pr["number"]
        if number in exclude_prs:
            remarks.append(f"#{number} excluida (--exclude-pr)")
            continue
        if number in cited:
            continue
        tip = pr.get("integrated_by")
        if not pr.get("direct") and tip is not None and tip in cited:
            remarks.append(f"#{number} integrada por la punta #{tip}")
            continue
        detail = "" if pr.get("direct") else " (intermedia sin punta citada)"
        missing.append(f"FALTA PR #{number}{detail}: {pr.get('title', '')}")
    for commit in data.get("direct_commits", []):
        refs = citations(commit.get("subject", ""))
        if refs & cited or commit["sha"][:7] in notes:
            continue
        missing.append(f"FALTA commit {commit['sha'][:7]} (sin PR): {commit.get('subject', '')}")
    issues = [i["number"] for i in data.get("issues", []) if i["number"] not in exclude_issues]
    if issues:
        register = register_section(notes)
        if register is None:
            missing.append("FALTA la sección «Registro de cierres» y el rango cierra issues")
        else:
            in_register = citations(register)
            missing.extend(f"FALTA issue #{n} en «Registro de cierres»" for n in issues
                           if n not in in_register)
    return missing, remarks


# ---------------------------------------------------------------------------
# CLI


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n", 1)[0])
    sub = parser.add_subparsers(dest="command", required=True)
    pc = sub.add_parser("collect", help="escribe el fichero del rango")
    pc.add_argument("--repo-dir", required=True)
    pc.add_argument("--from", dest="from_ref", required=True)
    pc.add_argument("--to", dest="to_ref", required=True)
    pc.add_argument("--github", help="owner/name para gh pr list")
    pc.add_argument("--prs-json", help="volcado de gh pr list --json " + PR_FIELDS)
    pc.add_argument("--out", required=True)
    pk = sub.add_parser("check", help="compara el rango con las notas")
    pk.add_argument("--range", required=True)
    pk.add_argument("--notes", required=True)
    pk.add_argument("--exclude-pr", type=int, action="append", default=[])
    pk.add_argument("--exclude-issue", type=int, action="append", default=[])
    args = parser.parse_args(argv)
    try:
        if args.command == "collect":
            repo = Path(args.repo_dir)
            data = collect(repo, args.from_ref, args.to_ref, _load_prs(args, repo))
            Path(args.out).write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n",
                                      encoding="utf-8")
            print(f"{len(data['prs'])} PR, {len(data['direct_commits'])} commits sin PR y "
                  f"{len(data['issues'])} issues en {args.from_ref}..{args.to_ref} -> {args.out}")
            return 0
        data = _read_json(Path(args.range))
        try:
            notes = Path(args.notes).read_text(encoding="utf-8")
        except OSError as exc:
            raise InputError(f"no se puede leer {args.notes}: {exc}") from exc
        missing, remarks = check(data, notes, set(args.exclude_pr), set(args.exclude_issue))
    except InputError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    for remark in remarks:
        print(f"nota: {remark}")
    if missing:
        print("\n".join(missing))
        print(f"FALLO: {len(missing)} elementos del rango {data.get('from')}..{data.get('to')} "
              "sin contar en las notas")
        return 1
    print(f"OK: {len(data.get('prs', []))} PR, {len(data.get('direct_commits', []))} commits sin PR "
          f"y {len(data.get('issues', []))} issues del rango {data.get('from')}..{data.get('to')} "
          "cubiertos")
    return 0


if __name__ == "__main__":
    sys.exit(main())
