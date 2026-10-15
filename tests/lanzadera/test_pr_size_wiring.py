# HARNESS-PROVENANCE: access2web-blueprint issue #854 — test_pr_size_wiring.py
"""Wiring pin for `scripts/check_pr_size.py` (issue #854): one test per acceptance rule.

The budget is additions + deletions, at most 400. The only override is the
`size:exception` label. Rules 1-3 share one parametrized verdict matrix.
"""

from __future__ import annotations

import importlib.util
import json
import subprocess
import sys
from pathlib import Path

import pytest

SCRIPT = Path(__file__).resolve().parents[2] / "scripts" / "check_pr_size.py"


def _load():
    spec = importlib.util.spec_from_file_location("check_pr_size", SCRIPT)
    module = importlib.util.module_from_spec(spec)
    sys.modules["check_pr_size"] = module
    spec.loader.exec_module(module)
    return module


def _run(*args, cwd=None):
    return subprocess.run(
        [sys.executable, str(SCRIPT), *args],
        cwd=cwd,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )


def _git(repo, *args):
    return subprocess.run(
        ["git", *args], cwd=repo, check=True, capture_output=True, text=True
    ).stdout


def _event(tmp_path, additions, deletions, labels=()):
    payload = {
        "pull_request": {
            "additions": additions,
            "deletions": deletions,
            "labels": [{"name": name} for name in labels],
        }
    }
    path = tmp_path / "event.json"
    path.write_text(json.dumps(payload), encoding="utf-8")
    return str(path)


def _local_repo(tmp_path, insertions, deletions):
    """Repo whose HEAD replaces the `deletions` lines of HEAD~1 with `insertions` new lines."""
    repo = tmp_path / "local"
    repo.mkdir()
    _git(repo, "init", "--initial-branch=main")
    _git(repo, "config", "user.email", "t@e")
    _git(repo, "config", "user.name", "t")
    (repo / "f.txt").write_text("old\n" * deletions, encoding="utf-8")
    _git(repo, "add", "f.txt")
    _git(repo, "commit", "-m", "base")
    (repo / "f.txt").write_text("new\n" * insertions, encoding="utf-8")
    _git(repo, "commit", "-am", "change")
    return repo


def test_ceiling_is_400():
    """The 400-line review budget is fixed."""
    assert _load().MAX_CHANGED_LINES == 400


# Rules 1-3: (mode, additions or insertions, deletions, labels) -> exit code, and the total
# in the output when over budget.
@pytest.mark.parametrize(
    ("mode", "adds", "dels", "labels", "expected_rc", "expected_total"),
    [
        ("pr-event", 100, 50, (), 0, None),  # rule 1: within budget
        ("local", 100, 0, (), 0, None),
        ("pr-event", 350, 200, (), 1, "550"),  # rule 2: over budget, no label
        ("local", 500, 0, (), 1, "500"),
        ("pr-event", 350, 200, ("size:exception",), 0, None),  # rule 3: the label overrides
    ],
    ids=["r1-pr-event", "r1-local", "r2-pr-event", "r2-local", "r3-label"],
)
def test_rules_1_2_3_verdict(mode, adds, dels, labels, expected_rc, expected_total, tmp_path):
    if mode == "pr-event":
        result = _run(mode, _event(tmp_path, adds, dels, labels))
    else:
        result = _run(mode, "HEAD~1", cwd=_local_repo(tmp_path, adds, dels))
    assert result.returncode == expected_rc, result.stderr
    if expected_total:
        assert expected_total in result.stderr


def test_rule_4_local_total_equals_the_api_total(tmp_path):
    """`local` gives the total GitHub reports for the same range: 200 added + 50 removed."""
    repo = _local_repo(tmp_path, insertions=200, deletions=50)
    assert _load().compute_local_total(repo, "HEAD~1") == 250


@pytest.mark.parametrize(
    ("adds", "dels", "expected_rc"),
    [(10, 5, 0), (500, 10, 1)],
    ids=["pass", "fail"],
)
def test_rule_5_binary_output_without_warnings(adds, dels, expected_rc, tmp_path):
    """The verdict is the exit code; stdout stays empty and stderr carries no Python warning."""
    result = _run("pr-event", _event(tmp_path, adds, dels))
    assert result.returncode == expected_rc
    assert result.stdout == ""
    assert "Warning" not in result.stderr


# ---------------------------------------------------------------------------------------------
# Rama vendorizada: las rutas de skills copiadas desde team-skills no cuentan como autoría.
# ---------------------------------------------------------------------------------------------


def _event_with_repo(tmp_path, additions, deletions, labels=()):
    """Evento de pull_request con número y repositorio base: habilita el desglose por fichero."""
    payload = {
        "pull_request": {
            "number": 910,
            "additions": additions,
            "deletions": deletions,
            "labels": [{"name": name} for name in labels],
            "base": {"repo": {"full_name": "DysTelefonica/access2web-blueprint"}},
        }
    }
    path = tmp_path / "event.json"
    path.write_text(json.dumps(payload), encoding="utf-8")
    return str(path)


@pytest.mark.parametrize(
    "path",
    [".agents/skills/ci-pattern/SKILL.md", "skills/documentation-alan-style/SKILL.md"],
)
def test_vendored_paths_are_not_author_lines(path):
    """Rule 6: the two vendored roots are excluded; the path is compared as a prefix."""
    module = _load()
    assert module.is_vendored(module.FileChange(path, added=1, deleted=0)) is True


@pytest.mark.parametrize(
    "path",
    ["scripts/check_pr_size.py", "skillsx/a.md", "myskills/skills/a.md", "docs/skills/a.md"],
)
def test_author_paths_are_not_vendored(path):
    """Rule 6: look-alike paths and nested `skills` directories stay author lines."""
    module = _load()
    assert module.is_vendored(module.FileChange(path, added=1, deleted=0)) is False


def test_rule_7_pr_with_only_vendored_paths_passes(tmp_path, monkeypatch):
    """Rule 7: a PR that only touches vendored skills (#910 shape) is within budget."""
    module = _load()
    monkeypatch.setenv("GITHUB_TOKEN", "t")
    monkeypatch.setattr(
        module,
        "fetch_pr_files",
        lambda repo, number, token: [
            module.FileChange(".agents/skills/ci-pattern/SKILL.md", added=5000, deleted=6000),
            module.FileChange("skills/documentation-alan-style/SKILL.md", added=2000, deleted=3000),
        ],
    )
    assert module.main(["pr-event", _event_with_repo(tmp_path, 7000, 9000)]) == 0


@pytest.mark.parametrize(
    ("author_add", "author_del", "expected_rc", "expected_author"),
    [(300, 0, 0, "300"), (350, 100, 1, "450")],
    ids=["mixed-within", "mixed-over"],
)
def test_rule_8_mixed_pr_counts_only_author_lines(
    author_add, author_del, expected_rc, expected_author, tmp_path, monkeypatch, capsys
):
    """Rule 8: a mixed PR is judged on its author lines, not on the vendored ones."""
    module = _load()
    monkeypatch.setenv("GITHUB_TOKEN", "t")
    monkeypatch.setattr(
        module,
        "fetch_pr_files",
        lambda repo, number, token: [
            module.FileChange("scripts/check_pr_size.py", author_add, author_del),
            module.FileChange(".agents/skills/x/SKILL.md", added=5000, deleted=4000),
        ],
    )
    event = _event_with_repo(tmp_path, author_add + 5000, author_del + 4000)
    assert module.main(["pr-event", event]) == expected_rc
    if expected_rc:
        assert expected_author in capsys.readouterr().err


def test_rule_8_rename_into_vendored_path_still_counts_the_author_side(tmp_path, monkeypatch):
    """Rule 8: a file renamed from an author path into a vendored one keeps counting."""
    module = _load()
    monkeypatch.setenv("GITHUB_TOKEN", "t")
    monkeypatch.setattr(
        module,
        "fetch_pr_files",
        lambda repo, number, token: [
            module.FileChange(
                ".agents/skills/x.md", added=0, deleted=450, previous_path="docs/x.md"
            ),
        ],
    )
    assert module.main(["pr-event", _event_with_repo(tmp_path, 0, 450)]) == 1


@pytest.mark.parametrize("fetch_result", [None, []], ids=["no-breakdown", "empty-breakdown"])
def test_rule_9_fail_closed_counts_everything(fetch_result, tmp_path, monkeypatch):
    """Rule 9: without a breakdown that adds up to the event total, every line counts."""
    module = _load()
    monkeypatch.setenv("GITHUB_TOKEN", "t")
    monkeypatch.setattr(module, "fetch_pr_files", lambda repo, number, token: fetch_result)
    assert module.main(["pr-event", _event_with_repo(tmp_path, 7000, 9000)]) == 1


def test_rule_9_inconsistent_breakdown_counts_everything(tmp_path, monkeypatch, capsys):
    """Rule 9: a breakdown that does not add up to the event total (truncated list) fails closed."""
    module = _load()
    monkeypatch.setenv("GITHUB_TOKEN", "t")
    monkeypatch.setattr(
        module,
        "fetch_pr_files",
        lambda repo, number, token: [module.FileChange("skills/a.md", added=100, deleted=0)],
    )
    assert module.main(["pr-event", _event_with_repo(tmp_path, 7000, 9000)]) == 1
    assert "16000" in capsys.readouterr().err


def test_rule_9_without_token_never_calls_the_api(tmp_path, monkeypatch):
    """Rule 9: with no GITHUB_TOKEN the breakdown is not requested and the total counts."""
    module = _load()
    monkeypatch.delenv("GITHUB_TOKEN", raising=False)

    def must_not_call(repo, number, token):
        raise AssertionError("no debe llamar a la API sin token")

    monkeypatch.setattr(module, "fetch_pr_files", must_not_call)
    assert module.main(["pr-event", _event_with_repo(tmp_path, 7000, 9000)]) == 1


class _FakeResponse:
    def __init__(self, body):
        self._body = body

    def read(self):
        return self._body

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


def test_rule_10_fetch_maps_files_and_sends_the_token_only_to_github():
    """Rule 10: Files API entries map per file; the job token goes only to api.github.com."""
    module = _load()
    seen = []

    def opener(request, timeout):
        seen.append((request.full_url, request.get_header("Authorization")))
        body = [
            {"filename": "skills/a.md", "additions": 3, "deletions": 1},
            {
                "filename": ".agents/skills/b.md",
                "previous_filename": "docs/b.md",
                "additions": 0,
                "deletions": 2,
            },
        ]
        return _FakeResponse(json.dumps(body).encode("utf-8"))

    files = module.fetch_pr_files("DysTelefonica/access2web-blueprint", 910, "t", opener=opener)
    assert files == [
        module.FileChange("skills/a.md", added=3, deleted=1),
        module.FileChange(".agents/skills/b.md", added=0, deleted=2, previous_path="docs/b.md"),
    ]
    url, auth = seen[0]
    assert url.startswith(
        "https://api.github.com/repos/DysTelefonica/access2web-blueprint/pulls/910/files"
    )
    assert auth == "Bearer t"


def _raise_oserror(request, timeout):
    raise OSError("sin red")


def _answer_not_a_list(request, timeout):
    return _FakeResponse(b'{"message": "Not Found"}')


def _answer_full_pages_forever(request, timeout):
    # 100 files per page for every page: the listing never ends within the 3000-file cap.
    body = [{"filename": "a", "additions": 1, "deletions": 0}] * 100
    return _FakeResponse(json.dumps(body).encode("utf-8"))


@pytest.mark.parametrize(
    "opener",
    [_raise_oserror, _answer_not_a_list, _answer_full_pages_forever],
    ids=["network-error", "not-a-list", "more-than-3000-files"],
)
def test_rule_11_fetch_failures_return_none(opener):
    """Rule 11: any failure, malformed answer or incomplete listing means no breakdown."""
    module = _load()
    assert (
        module.fetch_pr_files("DysTelefonica/access2web-blueprint", 910, "t", opener=opener) is None
    )


def test_rule_12_local_mode_excludes_vendored_lines(tmp_path):
    """Rule 12: local mode subtracts the vendored lines of the git diff: 510 total, 10 author."""
    repo = _local_repo_with_files(tmp_path, {".agents/skills/x/SKILL.md": 500, "scripts/a.py": 10})
    assert _run("local", "HEAD~1", cwd=repo).returncode == 0


def test_rule_12_local_mode_counts_author_lines_over_budget(tmp_path):
    """Rule 12: 450 author lines plus 500 vendored lines is still over budget."""
    repo = _local_repo_with_files(tmp_path, {".agents/skills/x/SKILL.md": 500, "scripts/a.py": 450})
    assert _run("local", "HEAD~1", cwd=repo).returncode == 1


def test_rule_13_local_breakdown_outside_a_repository_is_none(tmp_path):
    """Rule 13: no git range, no breakdown (fail-closed in local mode)."""
    module = _load()
    assert module.local_files(tmp_path, "origin/main") is None


def _local_repo_with_files(tmp_path, files):
    repo = tmp_path / "local"
    repo.mkdir()
    _git(repo, "init", "--initial-branch=main")
    _git(repo, "config", "user.email", "t@e")
    _git(repo, "config", "user.name", "t")
    (repo / "README.md").write_text("base\n", encoding="utf-8")
    _git(repo, "add", "README.md")
    _git(repo, "commit", "-m", "base")
    for relative, lines in files.items():
        target = repo / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text("line\n" * lines, encoding="utf-8")
    _git(repo, "add", "-A")
    _git(repo, "commit", "-m", "change")
    return repo
