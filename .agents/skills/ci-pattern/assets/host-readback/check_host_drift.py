#!/usr/bin/env python3
# ci-pattern asset — fail-closed host readback drift check (DysTelefonica/team-skills#138)
"""Compare the declared host contract against JSON snapshots of the host API.

Every documented rule — label, required check, merge method, ruleset,
protection flag — is classified in the contract (see
``host-contract.example.json``) as ``host-enforced`` (the host API must
show it) or ``documented-only`` (documentation only; MUST NOT be described
as a gate; HR-34). ``host_capabilities`` (#303) declares each area
(``branch_protection``, ``rulesets``, ``required_checks``) as
``available`` or ``unavailable`` and the readback verifies it in both
directions: the real plan-limit 403 body of a private free-tier repo
records ``unavailable`` (it is evidence, not an error), a real payload
records ``available``, and a contradiction is drift. The caller captures host API responses with read-only
GETs and passes them as snapshot files (``--snapshot kind=path``, all four
kinds required); the script never touches the network, never mutates the
host, never probes a write. Drift (exit 1) is either direction: a
``host-enforced`` declaration the host does not show or shows with a
different value, or an enforced required check / merge method / protection
flag the contract does not declare; ``documented-only`` entries are listed
and never judged. A host ruleset that the contract does not declare
host-enforced is drift in any enforcement mode other than ``disabled``
(``evaluate`` records evaluations and usually precedes activation); an
unknown ``enforcement`` value is doubt and fails closed (HR-3). Missing/unreadable/malformed contract or snapshot is
doubt, and doubt fails closed (HR-3). Exit codes: 0 no drift; 1 drift
findings; 2 fail-closed.
"""
import argparse
import json
import sys
from pathlib import Path

# #330/#332: el detector del 403 de plan vive en assets/host_plan.py, modulo
# compartido con el gate de adopcion, para que no existan dos detectores que
# puedan divergir; se importa en el mismo proceso, nunca en subproceso.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from host_plan import PLAN_403_MARKER, is_plan_403  # noqa: E402

HOST_ENFORCED = "host-enforced"
DOCUMENTED_ONLY = "documented-only"
RUNNER_ENFORCED = "runner-enforced"
CLASSES = frozenset({HOST_ENFORCED, DOCUMENTED_ONLY, RUNNER_ENFORCED})
CAPABILITIES = ("branch_protection", "rulesets", "required_checks")
CAPABILITY_VALUES = frozenset({"available", "unavailable"})
CAPABILITY_SNAPSHOT = {"branch_protection": "branch-protection",
                       "required_checks": "branch-protection",
                       "rulesets": "rulesets"}
MERGE_METHOD_KEYS = frozenset({"allow_merge_commit", "allow_squash_merge", "allow_rebase_merge", "allow_update_branch", "allow_auto_merge"})
PROTECTION_KEYS = frozenset({"enforce_admins", "required_linear_history", "required_conversation_resolution", "allow_force_pushes", "allow_deletions"})
RULESET_ENFORCEMENTS = frozenset({"active", "evaluate", "disabled"})
SNAPSHOT_KINDS = ("repo", "labels", "branch-protection", "rulesets")

class ContractError(Exception):
    """A condition the check refuses to interpret; maps to exit 2."""

def load_json(path: str, what: str) -> object:
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except OSError as exc:
        raise ContractError(f"unreadable {what} '{path}': {exc}") from exc
    except json.JSONDecodeError as exc:
        raise ContractError(f"{what} '{path}' is not valid JSON: {exc}") from exc

def validate_contract(data: object) -> dict:
    """Validate the declared contract; every structural doubt raises (exit 2)."""
    if not isinstance(data, dict) or data.get("contract_version") != 1:
        raise ContractError("contract must be an object with contract_version 1")
    known = {"contract_version", "labels", "required_checks", "merge_methods", "rulesets", "protection",
             "host_capabilities", "runner", "protected_branches"}
    protected = data.get("protected_branches")
    if protected is not None and (not isinstance(protected, list) or not protected
                                  or any(not isinstance(b, str) or not b for b in protected)):
        raise ContractError("protected_branches must be a non-empty list of branch names")
    if set(data) - known:
        raise ContractError(f"contract has unknown keys: {sorted(set(data) - known)}")
    caps = data.get("host_capabilities", {})
    if not isinstance(caps, dict) or set(caps) - set(CAPABILITIES):
        raise ContractError(
            f"host_capabilities must be an object with keys from {list(CAPABILITIES)}")
    for cap, value in caps.items():
        if value not in CAPABILITY_VALUES:
            raise ContractError(
                f"host_capabilities[{cap}] = {value!r}; need one of {sorted(CAPABILITY_VALUES)}")
    runner = data.get("runner", {})
    if not isinstance(runner, dict) or set(runner) - {"label"} \
            or ("label" in runner and (not isinstance(runner["label"], str) or not runner["label"])):
        raise ContractError(
            "runner must be an object with a single non-empty string 'label' "
            "(the self-hosted runner that applies runner-enforced rules)")
    contract: dict = {"labels": {}, "required_checks": {}, "merge_methods": {}, "protection": {}, "rulesets": [],
                      "host_capabilities": {cap: caps.get(cap, "available") for cap in CAPABILITIES}}
    for where in ("labels", "required_checks"):
        entries = data.get(where, [])
        if not isinstance(entries, list):
            raise ContractError(f"{where} must be a list")
        for entry in entries:
            if not isinstance(entry, dict) or not isinstance(entry.get("name"), str) or not entry["name"] or entry.get("class") not in CLASSES:
                raise ContractError(f"{where}: every entry needs a non-empty 'name' and 'class' one of {sorted(CLASSES)}")
            if entry["name"] in contract[where]:
                raise ContractError(f"{where}: duplicate name '{entry['name']}'")
            contract[where][entry["name"]] = entry["class"]
    for where, keys in (("merge_methods", MERGE_METHOD_KEYS), ("protection", PROTECTION_KEYS)):
        section = data.get(where, {})
        if not isinstance(section, dict) or set(section) - keys:
            raise ContractError(f"{where} must be an object with keys from {sorted(keys)}")
        for key, value in section.items():
            if not isinstance(value, dict) or not isinstance(value.get("declared"), bool) or value.get("class") not in CLASSES:
                raise ContractError(f"{where}[{key}] must be {{'declared': bool, 'class': {sorted(CLASSES)}}}")
            contract[where][key] = (value["declared"], value["class"])
    rulesets = data.get("rulesets", [])
    if not isinstance(rulesets, list):
        raise ContractError("rulesets must be a list")
    for entry in rulesets:
        if not isinstance(entry, dict) or not isinstance(entry.get("name"), str) or not entry["name"] or entry.get("enforcement") not in RULESET_ENFORCEMENTS or entry.get("class") not in CLASSES:
            raise ContractError(f"rulesets: bad entry {entry!r} (need 'name', "
                                f"'enforcement' one of {sorted(RULESET_ENFORCEMENTS)}, 'class' one of {sorted(CLASSES)})")
        contract["rulesets"].append((entry["name"], entry["enforcement"], entry["class"]))
    contract["runner_label"] = runner.get("label")
    has_runner_rules = any(cls == RUNNER_ENFORCED for cls in contract["labels"].values()) \
        or any(cls == RUNNER_ENFORCED for cls in contract["required_checks"].values()) \
        or any(cls == RUNNER_ENFORCED for _, cls in contract["merge_methods"].values()) \
        or any(cls == RUNNER_ENFORCED for _, cls in contract["protection"].values()) \
        or any(cls == RUNNER_ENFORCED for _, _, cls in contract["rulesets"])
    if has_runner_rules and not contract["runner_label"]:
        raise ContractError(
            "runner-enforced rules require contract key 'runner' with a non-empty 'label' "
            "(the self-hosted runner that applies them); a runner-enforced rule with no "
            "declared runner is unverifiable doubt (HR-3)")
    return contract

def load_snapshots(pairs: list[str]) -> dict:
    """Load the four required JSON snapshots; a missing one is doubt (HR-3)."""
    given: dict[str, str] = {}
    for pair in pairs:
        kind, sep, path = pair.partition("=")
        if not sep or kind not in SNAPSHOT_KINDS or kind in given:
            raise ContractError(f"--snapshot must be <kind>=<path>, kind one of {list(SNAPSHOT_KINDS)}, no duplicates")
        given[kind] = path
    missing = [kind for kind in SNAPSHOT_KINDS if kind not in given]
    if missing:
        raise ContractError(f"missing snapshots: {missing}; a snapshot that cannot be read is doubt, not a pass (HR-3)")
    snaps = {kind: load_json(path, f"snapshot '{kind}'") for kind, path in given.items()}
    for kind in SNAPSHOT_KINDS:
        want = dict if kind in ("repo", "branch-protection") else list
        # El cuerpo real del límite de plan en un kind con capacidad (#303)
        # se registra como 'unavailable': lo interpreta compare() contra el
        # host_capabilities del contrato. Fuera de esos kinds (labels, repo)
        # un cuerpo de error sigue siendo duda.
        if kind in CAPABILITY_SNAPSHOT.values() and is_plan_403(snaps[kind]):
            continue
        if not isinstance(snaps[kind], want):
            raise ContractError(
                f"snapshot '{kind}' ('{given[kind]}') must be a JSON {'object' if want is dict else 'array'}; "
                "got " + type(snaps[kind]).__name__)
    # Un cuerpo de error de la API (p.ej. 404: {'message': 'Not Found',
    # 'documentation_url': ...}) tiene forma de objeto válido pero no
    # contiene ningún campo esperado: la lectura falló y es duda, no un
    # pase (HR-3). Se nombra el fichero para que el operador lo vea.
    # Un valor de 'enforcement' fuera del enum conocido es duda: el host
    # aplica algo que esta version del control no sabe interpretar. Falla
    # cerrado nombrando el fichero, el ruleset y el valor (HR-3).
    if not is_plan_403(snaps["rulesets"]):
        for entry in snaps["rulesets"]:
            if isinstance(entry, dict) and "enforcement" in entry \
                    and entry["enforcement"] not in RULESET_ENFORCEMENTS:
                name = entry.get("name") if isinstance(entry.get("name"), str) else "<sin nombre>"
                raise ContractError(
                    f"snapshot 'rulesets' ('{given['rulesets']}'): unknown enforcement "
                    f"{entry['enforcement']!r} for ruleset {name!r}; a value outside "
                    f"{sorted(RULESET_ENFORCEMENTS)} is doubt, not evidence (HR-3)")
    for kind in ("repo", "branch-protection"):
        if kind in CAPABILITY_SNAPSHOT.values() and is_plan_403(snaps[kind]):
            continue
        snap = snaps[kind]
        expected = MERGE_METHOD_KEYS if kind == "repo" else PROTECTION_KEYS
        if "message" in snap and not (expected & set(snap)):
            raise ContractError(
                f"snapshot '{kind}' ('{given[kind]}') looks like a GitHub API error body "
                f"(has 'message' {snap['message']!r} and none of the expected '{kind}' fields); "
                "a failed API read is doubt, not a pass (HR-3)")
    for where, keys in (("repo", MERGE_METHOD_KEYS), ("branch-protection", PROTECTION_KEYS)):
        for key in keys & set(snaps[where]):
            value = snaps[where][key]
            if not isinstance(value, bool) and not (
                    isinstance(value, dict) and isinstance(value.get("enabled"), bool)):
                raise ContractError(f"snapshot '{where}': unexpected value for '{key}': {value!r}")
    return snaps

def host_flag(value: object) -> bool:
    """Accept a raw boolean or the branch-protection ``{"enabled": bool}`` shape."""
    if isinstance(value, bool):
        return value
    return bool(isinstance(value, dict) and isinstance(value.get("enabled"), bool) and value["enabled"])

def _flags_drift(section: dict, host: dict, keys: frozenset, noun: str,
                 ok: list, documented: list, findings: list) -> None:
    for key in sorted(keys):
        entry = section.get(key)
        if entry is not None and entry[1] == HOST_ENFORCED and key not in host:
            raise ContractError(
                f"snapshot '{noun}': '{key}' is declared host-enforced but absent from the host snapshot; "
                "absence is doubt, not evidence (HR-3)")
        host_val = host_flag(host.get(key))
        if entry is None:
            if host_val:
                findings.append(f"{noun} '{key}' is enabled on the host but not declared in the contract")
        elif entry[1] == DOCUMENTED_ONLY:
            documented.append(f"{noun} '{key}' (declared {entry[0]})")
        elif entry[1] == RUNNER_ENFORCED:
            continue  # juzgado en la pre-pasada runner-enforced
        elif host_val != entry[0]:
            findings.append(f"{noun} '{key}' is declared {entry[0]} but the host reports {host_val}")
        else:
            ok.append(f"{noun} '{key}' matches the host")

def compare(contract: dict, snaps: dict) -> tuple[list[str], list[str], list[str], list[str]]:
    ok: list[str] = []
    documented: list[str] = []
    findings: list[str] = []
    notes: list[str] = []
    caps = contract["host_capabilities"]
    repo, bp = snaps["repo"], snaps["branch-protection"]
    bp_unavailable = is_plan_403(bp)
    rs_unavailable = is_plan_403(snaps["rulesets"])
    seen = {"branch_protection": bp_unavailable, "required_checks": bp_unavailable,
            "rulesets": rs_unavailable}
    # #303: cada capacidad declarada se contrasta con la relectura en las
    # dos direcciones. El 403 de plan es evidencia de 'unavailable'; un
    # payload real es evidencia de 'available'; la contradicción es drift.
    for cap in CAPABILITIES:
        declared = caps[cap]
        noun = CAPABILITY_SNAPSHOT[cap]
        if seen[cap] and declared == "available":
            findings.append(f"capability '{cap}' is declared available but the readback recorded "
                            f"the plan 403 limit (snapshot '{noun}': {PLAN_403_MARKER}...)")
        elif not seen[cap] and declared == "unavailable":
            findings.append(f"capability '{cap}' is declared unavailable but the readback shows "
                            f"the capability (snapshot '{noun}')")
        else:
            notes.append(f"capability '{cap}' matches the readback ({declared})")
    # #303: tercera clase de HR-34. runner-enforced solo es valida en un
    # area cuya capacidad el host declara unavailable; la etiqueta del
    # runner viene de validate_contract (fallo cerrado sin ella).
    runner_entries: list[tuple[str, str, str]] = []
    for name, cls in contract["labels"].items():
        if cls == RUNNER_ENFORCED:
            findings.append(f"label '{name}' is declared runner-enforced but labels are applicable "
                            f"on every plan: declare host-enforced or documented-only")
    for name, cls in contract["required_checks"].items():
        if cls == RUNNER_ENFORCED:
            runner_entries.append(("required check", name, "required_checks"))
    for key, (_, cls) in contract["merge_methods"].items():
        if cls == RUNNER_ENFORCED:
            findings.append(f"merge method '{key}' is declared runner-enforced but repository settings "
                            f"are applicable on every plan: declare host-enforced or documented-only")
    for key, (_, cls) in contract["protection"].items():
        if cls == RUNNER_ENFORCED:
            runner_entries.append(("protection flag", key, "branch_protection"))
    for name, _, cls in contract["rulesets"]:
        if cls == RUNNER_ENFORCED:
            runner_entries.append(("ruleset", name, "rulesets"))
    for noun, name, cap in runner_entries:
        if caps[cap] == "available":
            findings.append(f"{noun} '{name}' is declared runner-enforced but capability '{cap}' "
                            f"is available: the host can enforce it (host-enforced)")
        else:
            notes.append(f"runner-enforced {noun} '{name}' applied by the workflow on runner "
                         f"'{contract['runner_label']}'")
    host_labels = {e["name"] for e in snaps["labels"]
                   if isinstance(e, dict) and isinstance(e.get("name"), str)}
    for name, cls in contract["labels"].items():
        if cls == DOCUMENTED_ONLY:
            documented.append(f"label '{name}'")
        elif cls == RUNNER_ENFORCED:
            continue  # juzgado en la pre-pasada runner-enforced
        elif name not in host_labels:
            findings.append(f"label '{name}' is declared host-enforced but is absent from the host")
        else:
            ok.append(f"label '{name}' matches the host")
    rsc = bp.get("required_status_checks") if not bp_unavailable and isinstance(bp.get("required_status_checks"), dict) else {}
    host_contexts = {c for c in (rsc.get("contexts") or []) if isinstance(c, str)}
    host_contexts |= {c["context"] for c in (rsc.get("checks") or []) if isinstance(c, dict) and isinstance(c.get("context"), str)}
    declared_checks = contract["required_checks"]
    if bp_unavailable:
        # La relectura registró el 403 de plan: el host no puede exigir
        # checks ni aplicar protección. Cada entrada host-enforced del
        # área es contradicción del contrato (drift con nombre).
        for name in sorted(declared_checks):
            if declared_checks[name] == HOST_ENFORCED:
                findings.append(f"required check '{name}' is declared host-enforced but the readback "
                                f"recorded the plan 403 limit for 'required_checks': the host cannot require it")
            elif declared_checks[name] == RUNNER_ENFORCED:
                continue  # juzgado en la pre-pasada runner-enforced
            else:
                documented.append(f"required check '{name}'")
    else:
        for name, cls in declared_checks.items():
            if cls == DOCUMENTED_ONLY:
                documented.append(f"required check '{name}'")
            elif cls == RUNNER_ENFORCED:
                continue  # juzgado en la pre-pasada runner-enforced
            elif name not in host_contexts:
                findings.append(f"required check '{name}' is declared host-enforced but the host does not require it")
            else:
                ok.append(f"required check '{name}' matches the host")
        for name in sorted(host_contexts - set(declared_checks)):
            findings.append(f"required check '{name}' is required on the host but not declared in the contract")
    _flags_drift(contract["merge_methods"], repo, MERGE_METHOD_KEYS, "merge method", ok, documented, findings)
    if bp_unavailable:
        for key in sorted(PROTECTION_KEYS):
            entry = contract["protection"].get(key)
            if entry is None:
                continue
            if entry[1] == HOST_ENFORCED:
                findings.append(f"protection flag '{key}' is declared host-enforced but the readback "
                                f"recorded the plan 403 limit for 'branch_protection': the host cannot enforce it")
            elif entry[1] == RUNNER_ENFORCED:
                continue  # juzgado en la pre-pasada runner-enforced
            else:
                documented.append(f"protection flag '{key}' (declared {entry[0]})")
    else:
        _flags_drift(contract["protection"], bp, PROTECTION_KEYS, "protection flag", ok, documented, findings)
    host_rs = {e["name"]: e.get("enforcement") for e in snaps["rulesets"]
               if isinstance(e, dict) and isinstance(e.get("name"), str)}
    if rs_unavailable:
        for name, enforcement, cls in contract["rulesets"]:
            if cls == HOST_ENFORCED:
                findings.append(f"ruleset '{name}' is declared host-enforced but the readback recorded "
                                f"the plan 403 limit for 'rulesets': the host cannot apply it")
            elif cls == RUNNER_ENFORCED:
                continue  # juzgado en la pre-pasada runner-enforced
            else:
                documented.append(f"ruleset '{name}'")
        return ok, documented, findings, notes
    for name, enforcement, cls in contract["rulesets"]:
        if cls == DOCUMENTED_ONLY:
            documented.append(f"ruleset '{name}'")
        elif cls == RUNNER_ENFORCED:
            continue  # juzgado en la pre-pasada runner-enforced
        elif name not in host_rs:
            findings.append(f"ruleset '{name}' is declared host-enforced but is absent from the host")
        elif host_rs[name] != enforcement:
            findings.append(f"ruleset '{name}' is declared with enforcement '{enforcement}' "
                            f"but the host reports '{host_rs[name]}'")
        else:
            ok.append(f"ruleset '{name}' matches the host")
    # Bidireccional: un ruleset del host que el contrato no declara
    # host-enforced (porque no esta, o porque lo clasifico documented-only)
    # es drift en cualquier modo distinto de 'disabled': 'active' bloquea
    # merges y 'evaluate' registra evaluaciones y suele preceder a la
    # activacion (gobierno no documentado, HR-34).
    declared_host_enforced_rs = {name for name, _, cls in contract["rulesets"] if cls == HOST_ENFORCED}
    for name in sorted(host_rs):
        mode = host_rs[name]
        if mode == "disabled" or name in declared_host_enforced_rs:
            continue
        if mode == "active":
            findings.append(f"ruleset '{name}' is active on the host but not declared host-enforced in the contract")
        elif mode == "evaluate":
            findings.append(f"ruleset '{name}' is in evaluate mode on the host but not declared host-enforced in the contract")
        else:
            findings.append(f"ruleset '{name}' on the host has no known enforcement mode but is not declared host-enforced in the contract")
    return ok, documented, findings, notes

def main() -> int:
    parser = argparse.ArgumentParser(
        description="Read-only drift check: declared host contract vs JSON snapshots of the host API (HR-34).")
    parser.add_argument("--contract", required=True)
    parser.add_argument("--snapshot", action="append", default=[], metavar="KIND=PATH")
    args = parser.parse_args()
    try:
        contract = validate_contract(load_json(args.contract, "contract"))
        snaps = load_snapshots(args.snapshot)
        ok, documented, findings, notes = compare(contract, snaps)
    except ContractError as exc:
        print(f"fail-closed: {exc}", file=sys.stderr)
        return 2
    if not ok and not findings:
        # HR-18 (liveness de la medición): un sujeto vacío o una medición
        # degenerada no es un PASS. Si el contrato no declara nada que el
        # host verifique y el host no muestra nada no declarado, el
        # readback no midió nada: falla cerrado.
        print("fail-closed: vacuous readback: the contract declares no host-enforced rule "
              "and the host snapshots enforce nothing; a degenerate measurement is not a pass (HR-18)",
              file=sys.stderr)
        return 2
    print("HOST READBACK REPORT")
    for line in notes:
        print(f"  {line}")
    for line in ok:
        print(f"  ok  {line}")
    print(f"documented-only rules (listed, not judged): {len(documented)}")
    for line in documented:
        print(f"  documented-only  {line}")
    for line in findings:
        print(f"  DRIFT  {line}")
    print("VERDICT: FAIL" if findings else "VERDICT: PASS")
    return 1 if findings else 0

if __name__ == "__main__":
    sys.exit(main())
