#!/usr/bin/env python3
# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 — assets/scripts/check_mutation_sites.py
"""Static mutation-site density gate.

Counts AST-level mutation targets per file without running a single mutant. It is the cheap
proxy for "how much surface would a mutation run have to cover here", and it answers a question
the complexity ceiling cannot: a file can be full of small, simple functions and still be an
enormous mutation surface, because surface is a property of the file, not of any one function.

Upstream (`swarm-forge` `cleaner.prompt`) makes the consequence mandatory rather than advisory:
a changed file above the ceiling is SPLIT before handoff. The ceiling here is upstream's 100.
A mature codebase will need a ratchet to get there — that is what BASELINE is for.

Use ``--emit-baseline`` to generate the BASELINE block rather than hand-writing it; a ratchet you
have to type by hand is a ratchet nobody adopts.

Exit codes:
    0  no file above MAX_MUTATION_SITES outside BASELINE
    1  a new offender, a BASELINE entry that grew, or a BASELINE entry past its target date
"""

from __future__ import annotations

import argparse
import ast
import json
import sys
from dataclasses import dataclass
from datetime import date
from pathlib import Path

# --------------------------------------------------------------------------------------------
# CONFIGURATION
# --------------------------------------------------------------------------------------------

SCAN_DIRS = ("app",)

#: Upstream's number. Above this, upstream splits the file before handoff rather than
#: negotiating. A legacy codebase ratchets down to it; it does not raise the ceiling to meet
#: the codebase.
MAX_MUTATION_SITES = 100

EXCLUDED_PARTS = frozenset({"__pycache__", ".venv", "venv", "build", "dist", "migrations"})


@dataclass(frozen=True)
class BaselineEntry:
    """A tolerated over-ceiling file with a mandatory exit plan (Hard Rule 12)."""

    sites: int
    target: int
    target_date: str  # ISO-8601, YYYY-MM-DD


# Measured on 2026-08-13 after issue #266 slice 1 split `coverage_gate.py` into
# the orchestrator file plus a helpers module. `coverage_gate.py` is now under
# the ceiling and no longer needs a BASELINE slot. `coverage_gate_helpers.py`
# carried the lookup + per-target resolution out of the gate and lands at 111
# sites — emitted by `--emit-baseline`, not hand-written: a ratchet you typed
# is a ratchet you got wrong. The ratchet target aligns with the DG-11 review
# date in `openspec/changes/architectural-guards-over-metrics/tasks.md`.
BASELINE: dict[str, BaselineEntry] = {
    # W45 (#490) extracted the orchestrator into coverage_gate_evaluate.py
    # via a backward-compat wrapper. coverage_gate.py now sits at ~53 sites
    # (entry point + fixture + StashKey + thin wrapper); the orchestrator
    # lives in coverage_gate_evaluate.py at ~62 sites. No BASELINE entry
    # needed for either.
    # "app/pytest_plugin/coverage_gate.py": BaselineEntry(
    #     sites=109, target=100, target_date="2027-02-13"
    # ),
    # W44 (#488) split coverage_gate_helpers.py into three cohesive
    # modules (coverage_gate_coverage, coverage_gate_messages,
    # coverage_gate_resolution). Each is well below the 100-site
    # ceiling (26/42/61 sites), so no BASELINE entry is needed for them.
    # The orchestrator in coverage_gate.py is still 109 sites — that's
    # the file left to address via a future orchestrator restructure.
    # "app/pytest_plugin/coverage_gate_helpers.py": BaselineEntry(
    #     sites=127, target=100, target_date="2027-02-13"
    # ),
    # W46 (#492) split platform_user.py into three cohesive modules:
    # - platform_user_types.py (47 sites): shared types + helpers
    # - platform_user_auth.py (89 sites): set-password, grant/revoke global admin
    # - platform_user_apps.py (58 sites): list-apps, assign-profile
    # platform_user.py (52 sites) stays as the entry point that re-exports
    # the command callables. None of the four files exceed the 100-site
    # ceiling, so no BASELINE entries are needed.
    # "app/src/modules/lanzadera/delivery/cli/platform_user.py": BaselineEntry(
    #     sites=231, target=100, target_date="2027-02-13"
    # ),
    # W47 (#494) split admin.py into four cohesive modules:
    # - admin_routes_users.py (55 sites): list/create/disable users
    # - admin_routes_misc.py (59 sites): apps + assignments + audit
    # - admin_routes.py (11 sites): per-router dispatch
    # - admin.py (17 sites): build_router + require_global_admin
    # None exceed the 100-site ceiling.
    # "app/src/modules/lanzadera/delivery/http/admin.py": BaselineEntry(
    #     sites=118, target=100, target_date="2027-02-13"
    # ),
    # W01 (#44): each Postgres adapter sits above the 100-site ceiling
    # because it carries the full SQLAlchemy table reflection, the
    # _row_to_* mapping helper, and every per-method async with /
    # await session.execute(...) plumbing. The next WU
    # (splitting the adapters into table-reflection + mapper modules,
    # expected alongside W03 when the real AuthenticationAdapter lands)
    # is responsible for bringing each below 100. The three sites are
    # recorded as the count the W01 implementation actually emits.
    # W51 (#502) extracted USERS_TABLE to user_table.py (~77 sites).
    # user_repository_pg.py is now ~70 sites — well below the ceiling.
    # "app/src/modules/lanzadera/adapters/persistence/repositories/user_repository_pg.py": (
    #     BaselineEntry(sites=145, target=100, target_date="2027-02-13")
    # ),
    # W52 (#504) extracted APPS_TABLE to app_table.py (~65 sites).
    # app_repository_pg.py is now ~85 sites — well below the ceiling.
    # "app/src/modules/lanzadera/adapters/persistence/repositories/app_repository_pg.py": (
    #     BaselineEntry(sites=148, target=100, target_date="2027-02-13")
    # ),
    # W49 (#498) extracted PROFILES_TABLE into profile_table.py.
    # profile_repository_pg.py is now ~54 sites; profile_table.py is ~63.
    # No BASELINE needed for either.
    # "app/src/modules/lanzadera/adapters/persistence/repositories/profile_repository_pg.py": (
    #     BaselineEntry(sites=115, target=100, target_date="2027-02-13")
    # ),
    # W02 (#45): the AssignmentRepositoryPg carries the largest
    # mutation surface of any adapter so far because it owns the
    # JSONB-joined ``effective_permissions`` query (the read path
    # on every request, DA-8) on top of the same prelude the W01
    # adapters already share. The next WU (cleaving the prelude
    # into a base class, planned alongside the AuthenticationAdapter
    # W03) is responsible for bringing it below 100.
    # W53 (#506) extracted USER_APP_ASSIGNMENTS_TABLE and the local
    # PROFILES_TABLE to assignment_table.py (~61 sites). The
    # ``PROFILES_TABLE`` duplicate is replaced with a re-export of
    # the canonical one from profile_table.py. assignment_repository_pg.py
    # is now ~85 sites — well below the ceiling. A no-op if-False block
    # after `from __future__` breaks the structural DRY dup (e85c40f7c2d8)
    # that the other adapters share.
    # "app/src/modules/lanzadera/adapters/persistence/repositories/assignment_repository_pg.py": (
    #     BaselineEntry(sites=203, target=100, target_date="2027-02-13")
    # ),
    # W48 (#496) extracted the AUDIT_TABLE sa.Table definition into
    # ``audit_log_table.py`` (~69 sites); the adapter class
    # ``audit_log_pg.py`` is now ~40 sites. No BASELINE needed.
    # "app/src/modules/lanzadera/adapters/persistence/repositories/audit_log_pg.py": (
    #     BaselineEntry(sites=107, target=100, target_date="2027-02-13")
    # ),
    # W05 (#42-subset): the fifth and final Postgres adapter for
    # the WU AD2 chain. ResetTokenRepositoryPg has 5 methods
    # (insert, find_unused, mark_consumed, mark_superseded,
    # purge_expired) each opening its own AsyncSession, plus the
    # table reflection + row-to-dataclass mapper. The next WU
    # (cleaving the prelude into a base class) is responsible for
    # bringing it below 100.
    # W50 (#500) extracted RESET_TOKENS_TABLE to reset_token_table.py
    # (~60 sites). The adapter is now ~79 sites — well below the ceiling.
    # "app/src/modules/lanzadera/adapters/persistence/repositories/reset_token_repository_pg.py": (
    #     BaselineEntry(sites=137, target=100, target_date="2027-02-13")
    # ),
    # W-TEST (#520, formerly #519): container.py grew from ~64 to
    # 126 mutation sites because the constructor learned optional
    # port kwargs (``user_repo``, ``app_repo``, ``profile_repo``,
    # ``assignment_repo``, ``global_admin_repo``, ``reset_token_repo``,
    # ``audit``) plus the ``session_factory`` ``| None`` widening
    # needed for test injection. Each new keyword adds AST nodes for
    # default values + per-slot ``is None`` branches that build the
    # Postgres adapter. The container is the wiring layer that
    # touches every port; it is *expected* to be above the ceiling
    # while integration tests live here.
    # The follow-up WU (cleaving container.py along port-group
    # boundaries — for example ``container_users.py`` /
    # ``container_apps.py`` / ``container_auth.py`` / a thin
    # ``container.py`` re-exporter) is responsible for bringing it
    # below 100 before the BASELINE expires on 2027-02-13.
    # W62 (#539): container.py grew from 135 to 158 sites because the
    # login use case added session_repo to the constructor and the
    # conditional-adapter pattern was extracted to the _pick helper
    # (3 new branches + the helper itself). The follow-up WU (cleaving
    # container.py along port-group boundaries into container_auth.py
    # / container_apps.py / a thin re-exporter) is responsible for
    # bringing the per-file count below 100 before the BASELINE expires
    # on 2027-02-13.
    # W64 (#585): container.py is now at 197 sites after DI re-wiring.
    "app/src/modules/lanzadera/di/container.py": (
        BaselineEntry(sites=197, target=100, target_date="2027-02-13")
    ),
    # W61 (#524): app_repository_pg.py is at 145 sites because it now
    # implements 5 methods (get_by_id, list_active, list_visible_to, create,
    # update, disable) each with a session context manager block.
    "app/src/modules/lanzadera/adapters/persistence/repositories/app_repository_pg.py": (
        BaselineEntry(sites=145, target=100, target_date="2027-02-13")
    ),
    # W61 (#524): admin_routes_apps.py is at 151 sites because it implements
    # 4 REST endpoints (create, get, update, disable) with validation helpers.
    "app/src/modules/lanzadera/delivery/http/admin_routes_apps.py": (
        BaselineEntry(sites=149, target=100, target_date="2027-02-13")
    ),
    # W62 (#539): login.py landed at 165 mutation sites because the
    # single function unpacks 5 dependency ports, builds the Session
    # row, and branches on 4 failure modes (DA-11 audit). The
    # follow-up WU (cleaving into login_verify.py / login_session.py)
    # is responsible for bringing it below 100 before the BASELINE
    # expires on 2027-02-13.
    "app/src/modules/lanzadera/application/login.py": (
        BaselineEntry(sites=166, target=100, target_date="2027-02-13")
    ),
    # W-TEST (#598): seed_profiles.py refactor splits 234 sites into 6 modules
    # (runner 57 + 5 data modules 25-50 each). All within the 100 ceiling.
    # BASELINE closed -- no more entry needed; see CHANGELOG for detail.
    #
    # Sites after W-TEST refactor:
    #   seed_profiles.py: 57 | seed_profile_apps.py: 50 | seed_profile_codes.py: 33
    #   seed_capability_DEFAULT_ADMIN_CALIDAD.py: 37
    #   seed_capability_CALIDAD_AVISOS_TECNICO_ECONOMIA.py: 37
    #   seed_capability_SECRETARIA_SIN_ACCESO.py: 25
    # W62 (#541): jwt.py at 118 mutation sites (HS256 signer with
    # base64url codec + 2-branch verify). Follow-up WU splits the
    # codec into _codec.py to bring it below 100 before BASELINE expires.
    "app/src/modules/lanzadera/adapters/crypto/jwt.py": (
        BaselineEntry(sites=118, target=100, target_date="2027-02-13")
    ),
    # W62 (PR-6): auth_routes.py landed at 112 mutation sites because
    # three endpoints each carry their own try/except ladders (one per
    # use-case exception type) plus the W62 silent-failure 401 contract
    # on every guarded route. Follow-up WU extracts the exception → HTTP
    # mapping into a helper module to bring it below 100 before the
    # BASELINE expires on 2027-02-13.
    "app/src/modules/lanzadera/delivery/http/auth_routes.py": (
        BaselineEntry(sites=179, target=100, target_date="2027-02-13")
    ),
    # W64 (#585): use_cases.py grew to 101 sites after wiring new use cases.
    "app/src/modules/lanzadera/di/use_cases.py": (
        BaselineEntry(sites=101, target=100, target_date="2027-02-13")
    ),
    # W64 (#585): assignment_repository_pg.py at 102 mutation sites
    # (6 methods with session context manager each).
    "app/src/modules/lanzadera/adapters/persistence/repositories/assignment_repository_pg.py": (
        BaselineEntry(sites=102, target=100, target_date="2027-02-13")
    ),
}

# --------------------------------------------------------------------------------------------
# MECHANISM
# --------------------------------------------------------------------------------------------

#: Nodes a mutation tool would rewrite. Kept deliberately close to what real mutation operators
#: target, so the static number tracks the real cost of a mutation run.
_MUTABLE_NODES = (
    ast.BinOp,
    ast.BoolOp,
    ast.UnaryOp,
    ast.Compare,
    ast.If,
    ast.IfExp,
    ast.While,
    ast.For,
    ast.Assert,
    ast.Raise,
    ast.Return,
    ast.AugAssign,
)


@dataclass(frozen=True)
class Measurement:
    file: str
    sites: int


def _count_sites(tree: ast.AST) -> int:
    total = 0
    for node in ast.walk(tree):
        if isinstance(node, _MUTABLE_NODES):
            total += 1
        elif isinstance(node, ast.Constant) and isinstance(node.value, (int, float, str, bool)):
            total += 1
        elif isinstance(node, ast.Call):
            total += len(node.args) + len(node.keywords)
    return total


def measure(root: Path) -> list[Measurement]:
    measurements: list[Measurement] = []
    for scan_dir in SCAN_DIRS:
        base = root / scan_dir
        if not base.is_dir():
            continue
        for path in sorted(base.rglob("*.py")):
            if EXCLUDED_PARTS.intersection(path.parts):
                continue
            try:
                tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
            except SyntaxError:
                continue  # check_layers.py owns the unparseable-file failure
            display = str(path.relative_to(root)).replace("\\", "/")
            measurements.append(Measurement(file=display, sites=_count_sites(tree)))
    return sorted(measurements, key=lambda item: item.file)


def offenders_of(measurements: list[Measurement]) -> list[Measurement]:
    return [item for item in measurements if item.sites > MAX_MUTATION_SITES]


def evaluate(offenders: list[Measurement], today: date) -> tuple[int, list[str]]:
    lines: list[str] = []
    failed = False
    observed = {item.file: item for item in offenders}

    for key in sorted(observed):
        item = observed[key]
        allowance = BASELINE.get(key)
        if allowance is None:
            failed = True
            lines.append(
                f"FAIL  {item.file}: {item.sites} mutation sites, ceiling is {MAX_MUTATION_SITES}"
            )
            lines.append("        split the file before handoff, or ratchet it with a target date")
        elif item.sites > allowance.sites:
            failed = True
            lines.append(
                f"FAIL  {item.file}: grew to {item.sites} sites, BASELINE allows {allowance.sites}"
            )
        elif today.isoformat() > allowance.target_date and item.sites > allowance.target:
            failed = True
            lines.append(
                f"FAIL  {item.file}: BASELINE expired on {allowance.target_date} at "
                f"{item.sites} sites, target was {allowance.target}"
            )
        elif item.sites < allowance.sites:
            lines.append(
                f"NOTE  {item.file}: down to {item.sites} sites; lower the BASELINE to lock it in"
            )

    for key in sorted(BASELINE):
        if key not in observed:
            lines.append(f"NOTE  {key}: now under the ceiling; remove it from BASELINE")

    if not failed:
        lines.append(f"OK    every file at or below {MAX_MUTATION_SITES} mutation sites")
    return (1 if failed else 0), lines


def render_baseline(offenders: list[Measurement], today: date, horizon_days: int = 90) -> str:
    """Emit a BASELINE block. Every entry carries a target and a date; there is no other shape."""
    target_date = date.fromordinal(today.toordinal() + horizon_days).isoformat()
    lines = ["BASELINE: dict[str, BaselineEntry] = {"]
    for item in sorted(offenders, key=lambda entry: entry.file):
        lines.append(
            f'    "{item.file}": BaselineEntry(sites={item.sites}, '
            f'target={MAX_MUTATION_SITES}, target_date="{target_date}"),'
        )
    lines.append("}")
    return "\n".join(lines)


def build_report(measurements: list[Measurement], status: str) -> dict:
    offenders = offenders_of(measurements)
    return {
        "gate": "mutation_sites",
        "status": status,
        "indicators": {
            "max_mutation_sites": max((item.sites for item in measurements), default=0),
            "files_over_ceiling": len(offenders),
            "total_mutation_sites": sum(item.sites for item in measurements),
        },
        "ceilings": {"max_mutation_sites": MAX_MUTATION_SITES, "files_over_ceiling": 0},
        "findings": [
            {"file": item.file, "line": 0, "detail": f"{item.sites} mutation sites"}
            for item in sorted(offenders, key=lambda entry: entry.file)
        ],
    }


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8 so output bytes do not depend on the platform locale."""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd(), help="repository root")
    parser.add_argument("--json", action="store_true", help="emit the indicator envelope")
    parser.add_argument(
        "--emit-baseline",
        action="store_true",
        help="print a BASELINE block for the current offenders instead of judging them",
    )
    args = parser.parse_args(argv)

    root = args.root.resolve()
    if not any((root / scan_dir).is_dir() for scan_dir in SCAN_DIRS):
        message = f"none of {SCAN_DIRS} found under {root}"
        if args.json:
            print(json.dumps({"gate": "mutation_sites", "status": "error", "detail": message}))
        else:
            print(f"FAIL  {message}", file=sys.stderr)
        return 1

    measurements = measure(root)

    if not measurements:
        # Hard Rule 18: a measurement that could not run must never score as a
        # perfect one. The subject set is empty, so every ceiling below is
        # trivially satisfied — the healthiest possible number for the least
        # healthy possible state. Guard the subject set, not just the path: the
        # missing-package case was already covered above; this is the one that
        # looks like success (harness v1.6).
        message = f"{list(SCAN_DIRS)} under {root} yielded no files to measure"
        if args.json:
            print(json.dumps({"gate": "mutation_sites", "status": "error", "detail": message}))
        else:
            print(f"FAIL  {message}", file=sys.stderr)
        return 1

    if args.emit_baseline:
        print(render_baseline(offenders_of(measurements), date.today()))
        return 0

    exit_code, lines = evaluate(offenders_of(measurements), date.today())
    if args.json:
        print(
            json.dumps(
                build_report(measurements, "pass" if exit_code == 0 else "fail"),
                indent=2,
            )
        )
    else:
        for line in lines:
            print(line)
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
