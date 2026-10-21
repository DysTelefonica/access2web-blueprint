#!/usr/bin/env python3
"""host-readback known set derives from the contract schema (DysTelefonica/team-skills#946). stdlib exclusively.

The gate check_host_drift.py rejects contracts whose top-level keys are not
in a hardcoded `known` set (literal). That set fell behind the canon:
HR-817 adds `post_merge_dispatch` to host-contract.json and SKILL.md plus
parameters.md declare it, but the gate's literal does not. This module-level
test catches the gap by importing `validate_contract` and asserting that a
contract declaring a schema-known key is accepted.

Two test cases:
- the RED case: a contract declaring a key that the schema recognises must
  not raise ContractError. RED on main because the literal `known` lacks
  the new key. GREEN after the fix derives known from the schema.
- the fixity case: a contract declaring a key neither in the schema nor in
  any other recognised list keeps failing with the rc=2 fail-closed
  message. Behaviour preserved across the fix.
"""
import sys
import unittest
from pathlib import Path

THIS_DIR = Path(__file__).resolve().parent
GATE = THIS_DIR.parent  # assets/host-readback/
sys.path.insert(0, str(GATE))

from check_host_drift import ContractError, validate_contract  # noqa: E402


def _minimal_contract(**extras):
    """A contract that passes every other validate_contract check,
    leaving only the 'is the key known?' question to drive the result."""
    base = {"contract_version": 1, "protected_branches": ["main"]}
    base.update(extras)
    return base


class KnownDerivesFromSchemaTests(unittest.TestCase):
    """RED on main: 'post_merge_dispatch' is in the canon (HR-817) but not
    in check_host_drift's literal `known`. The fix derives known from the
    schema so the single source of truth is the schema file."""

    def test_post_merge_dispatch_is_a_known_top_level_key(self):
        contract = _minimal_contract(
            post_merge_dispatch=["ci", "push-compliance", "force-push-guard"],
        )
        try:
            validate_contract(contract)
        except ContractError as exc:
            self.fail(
                "validate_contract rejected a contract that declares the "
                "schema-known key post_merge_dispatch (HR-817): "
                + str(exc)
            )

    def test_truly_unknown_keys_still_fail(self):
        """Fixity: a key neither in the schema nor in the legacy literal
        keeps failing with 'contract has unknown keys' and rc=2. This
        behaviour is the contract's fail-closed promise (HR-3)."""
        contract = _minimal_contract(bogus_extra_top_level_key=True)
        with self.assertRaises(ContractError) as cm:
            validate_contract(contract)
        self.assertIn("contract has unknown keys", str(cm.exception))


if __name__ == "__main__":
    unittest.main()
