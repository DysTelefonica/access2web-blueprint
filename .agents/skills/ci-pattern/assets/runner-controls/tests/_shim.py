"""Shared shim and fixtures for safe-merge tests (#426).

El shim simula las llamadas a ``gh api`` que hace ``safe_merge.py``.
Los tests de validaciones (test_safe_merge.py) y los de chain/API
(test_safe_merge_chain.py) lo importan para no duplicar el codigo
del shim y los fixtures.
"""
import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

# Constantes locales (los tests no dependen de test_push_compliance
# para que este PR pueda ser revisable y mergable de forma autonoma).
REPO = "DysTelefonica/team-skills"
SHA = "1f4e9a2c7b8d3e6f5a0c9b8d7e6f5a4c3b2d1e0f"
ASSET = Path(__file__).resolve().parent.parent / "safe_merge.py"


class ShimBuilder:
    """Construye el shim gh en Python para simular las llamadas de gh api."""

    def __init__(self) -> None:
        self.lines: list[str] = [
            '#!/usr/bin/env python3',
            'import json, sys, re',
            'STATE_PATH = PLACEHOLDER_STATE',
            'state = json.load(open(STATE_PATH, encoding="utf-8"))',
            'args = sys.argv[1:]',
            'state["calls"].append(" ".join(args))',
            '',
            '# api_error fixture',
            'if state.get("api_error"):',
            '    sys.stderr.write("gh: HTTP 500 (api_error fixture)\\n"); sys.exit(1)',
            '',
            '# POST repos/.../issues',
            'if "-X" in args and "POST" in args:',
            '    if state.get("issue_post_error"):',
            '        sys.stderr.write("gh: HTTP 503\\n"); sys.exit(1)',
            '    endpoint = next((a for a in args if a.startswith("repos/") and a.endswith("/issues")), None)',
            '    if endpoint is None:',
            '        sys.stderr.write("gh: unsupported POST endpoint\\n"); sys.exit(1)',
            '    fields = {k: v for i, a in enumerate(args) if a in ("-f", "-F") and i + 1 < len(args)',
            '               for k, _, v in [args[i + 1].partition("=")]}',
            '    state["created_issues"].append(fields)',
            '    json.dump(state, open(STATE_PATH, "w", encoding="utf-8"))',
            '    repo_name = endpoint.split("/issues")[0].replace("repos/", "")',
            '    print(json.dumps({"number": 42, "html_url": "https://github.com/" + repo_name + "/issues/42", "title": fields.get("title", "")}))',
            '    sys.exit(0)',
            '',
            '# PUT repos/.../merge',
            'if "-X" in args and "PUT" in args:',
            '    endpoint = next((a for a in args if a.startswith("repos/") and a.endswith("/merge")), None)',
            '    if endpoint is None:',
            '        sys.stderr.write("gh: unsupported PUT endpoint\\n"); sys.exit(1)',
            '    if state.get("merge_rejected"):',
            '        sys.stderr.write("gh: HTTP 405 (merge_rejected fixture)\\n"); sys.exit(1)',
            '    if state.get("head_changed") and any(a.startswith("sha=") for a in args):',
            '        sys.stderr.write("gh: HTTP 409 (Head branch was modified)\\n"); sys.exit(1)',
            '    state["merges"].append(endpoint)',
            '    json.dump(state, open(STATE_PATH, "w", encoding="utf-8"))',
            '    print(json.dumps({"sha": "2c7d9b1a4e", "merged": True, "message": "Pull Request successfully merged"}))',
            '    sys.exit(0)',
            '',
            '# DELETE repos/.../git/refs/heads/{branch}',
            'if "-X" in args and "DELETE" in args:',
            '    if state.get("delete_error"):',
            '        sys.stderr.write("gh: HTTP 500 (delete_error)\\n"); sys.exit(1)',
            '    endpoint = next((a for a in args if a.startswith("repos/") and "git/refs/heads/" in a), None)',
            '    if endpoint is None:',
            '        sys.stderr.write("gh: unsupported DELETE endpoint\\n"); sys.exit(1)',
            '    state["deleted_branches"].append(endpoint)',
            '    json.dump(state, open(STATE_PATH, "w", encoding="utf-8"))',
            '    print(json.dumps({"message": "deleted"}))',
            '    sys.exit(0)',
        ]

    def add_get_repos(self) -> None:
        """GET repos/{repo} -> default_branch."""
        self.lines.extend([
            '',
            '# GET repos/{repo} -> default_branch',
            'all_paths = [a for a in args if a.startswith("repos/") and not a.startswith("-")]',
            'path = all_paths[0] if all_paths else None',
            'if path is not None and "/commits/" in path and "/check-runs" in path:',
            '    sha = path.split("/commits/")[1].split("/")[0]',
            '    runs = state["check_runs"].get(sha, [])',
            '    print(json.dumps({"total_count": len(runs), "check_runs": runs}))',
            'elif path is not None and "/pulls/" in path and not path.endswith("/pulls"):',
            '    detail = state["pr_details"].get(path.rsplit("/pulls/", 1)[1])',
            '    if detail is None:',
            '        sys.stderr.write("gh: 404 Not Found\\n"); sys.exit(1)',
            '    print(json.dumps(detail))',
            'elif path is not None and path.endswith("/pulls") and "/commits/" in path:',
            '    sha = path.split("/commits/")[1].split("/")[0]',
            '    print(json.dumps(state["pulls"].get(sha, [])))',
            'elif path is not None and "/compare/" in path:',
            '    print(json.dumps(state.get("compare", {"total_commits": 0, "commits": [], "behind_by": 0, "ahead_by": 0})))',
            'elif path is not None and re.match(r"^repos/[^/]+/[^/]+$", path):',
            '    print(json.dumps({"full_name": PLACEHOLDER_REPO, "default_branch": state.get("default_branch", "main")}))',
            'elif path is not None:',
            '    sys.stderr.write("gh: unsupported endpoint " + path + "\\n"); sys.exit(1)',
            'else:',
            '    sys.stderr.write("gh: no endpoint\\n"); sys.exit(1)',
            'json.dump(state, open(STATE_PATH, "w", encoding="utf-8"))',
        ])

    def build(self, state_path: str, repo: str) -> str:
        code = "\n".join(self.lines)
        code = code.replace("PLACEHOLDER_STATE", repr(state_path))
        code = code.replace("PLACEHOLDER_REPO", repr(repo))
        return code


def pull(**over: object) -> dict:
    base: dict = {
        "number": 7, "state": "open", "mergeable": True, "draft": False,
        "user": {"login": "ardelperal"}, "head": {"ref": "chore/test", "sha": SHA},
        "base": {"ref": "main"}, "title": "feat: change",
        "body": "Closes #426",
        "labels": [], "additions": 10, "deletions": 5,
    }
    base.update(over)
    return base


def contract() -> dict:
    return {"contract_version": 1, "protected_branches": ["main"],
            "required_checks": [{"name": "unit", "class": "documented-only"}]}


def base_state() -> dict:
    return {
        "pulls": {SHA: [{"number": 7, "merged": True, "merged_at": "2026-10-12T10:00:00Z",
                         "user": {"login": "ardelperal"}, "base": {"ref": "main"},
                         "head": {"ref": "chore/test", "sha": SHA}}],
                  "chore/test": pull()},
        "pr_details": {"7": pull()},
        "check_runs": {SHA: [{"id": 1, "name": "unit", "conclusion": "success", "status": "completed"}]},
        "created_issues": [], "calls": [], "merges": [], "deleted_branches": [],
        "compare": {"total_commits": 1, "commits": [{"sha": SHA,
                       "commit": {"author": {"name": "Dev A"}}}], "behind_by": 0, "ahead_by": 1},
        "default_branch": "main",
    }


def run_gate(root: Path, state: dict, pr_num: int = 7,
             required_check: list | None = None, check_only: bool = False) -> subprocess.CompletedProcess:
    """Escribe el shim y el estado; corre el gate."""
    shim = root / "bin" / "gh"
    shim.parent.mkdir(exist_ok=True)
    state_file = root / "state.json"
    state_file.write_text(json.dumps(state), encoding="utf-8")
    builder = ShimBuilder()
    builder.add_get_repos()
    shim.write_text(builder.build(str(state_file), REPO), encoding="utf-8")
    shim.chmod(0o755)
    cmd = [sys.executable, str(ASSET), "--repo", REPO, "--pr", str(pr_num)]
    if required_check:
        for rc in required_check:
            cmd += ["--required-check", rc]
    if check_only:
        cmd.append("--check-only")
    env = {**os.environ, "PATH": f"{root / 'bin'}{os.pathsep}{os.environ.get('PATH', '')}"}
    return subprocess.run(cmd, capture_output=True, text=True, env=env, timeout=60)


class BaseSafeMergeTest(unittest.TestCase):
    """Setup comun para los tests de safe-merge."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def assert_merged(self, proc: subprocess.CompletedProcess, expected: bool) -> None:
        state = json.loads((self.root / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(bool(state["merges"]), expected, proc.stdout + proc.stderr)
