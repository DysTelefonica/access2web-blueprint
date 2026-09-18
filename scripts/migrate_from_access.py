#!/usr/bin/env python3
"""migrate_from_access.py — one-shot hash migration for Lanzadera MVP (WU E2).

Reads a CSV fixture of legacy Access users with their unsalted SHA256-hex
password hashes (per ``integrations-security.md`` H1) and either:
  - accepts the row (logs + skips; the ``PasswordHasher`` upgrade is
    post-MVP per #625), or
  - aborts non-zero with a precise error when the SHA256 hex has the
    wrong length, non-hex characters, or mismatches the expected format.

CLI:
    python scripts/migrate_from_access.py \\
        --input tests/fixtures/lanzadera/access_users.csv \\
        [--expect-sha256-hash 5e884898da28...]   (reference known-good)

Exit codes:
  0  every row has a valid 64-char hex SHA256 hash (or matches --expect)
  1  CLI / parse error (missing file, bad args)
  2  at least one row had a malformed SHA256 hash

The script is idempotent: running it twice produces the same exit code and
the same dry-run report. It performs no writes; the actual upgrade path
from unsalted SHA256 to Argon2id PHC is implemented by the production
``PasswordHasher`` (DA-2) and is out of scope for this script.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import sys
from pathlib import Path

EXPECTED_SHA256_HEX_LEN = 64


def _parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "One-shot migration script for legacy Access passwords. "
            "Validates that each row's SHA256-hex hash is well-formed; "
            "exits 2 if any row is malformed. No writes."
        )
    )
    parser.add_argument(
        "--input", required=True, type=Path, help="Path to the legacy Access users CSV fixture."
    )
    parser.add_argument(
        "--expect-sha256-hash",
        default=None,
        help="Optional reference 64-char hex SHA256 hash every row must match.",
    )
    return parser.parse_args(argv)


def _is_valid_sha256_hex(s: str, expected: str | None) -> tuple[bool, str]:
    if not s:
        return False, "empty hash"
    if len(s) != EXPECTED_SHA256_HEX_LEN:
        return False, f"hash length {len(s)} != {EXPECTED_SHA256_HEX_LEN}"
    try:
        bytes.fromhex(s)
    except ValueError as exc:
        return False, f"non-hex character: {exc}"
    if expected is not None and s != expected:
        return False, "hash does not match --expect-sha256-hash"
    if len(s) != len(hashlib.sha256(b"").hexdigest()):
        return False, "wrong digest length"
    return True, "ok"


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv if argv is not None else sys.argv[1:])

    if not args.input.exists():
        print(f"error: input file not found: {args.input}", file=sys.stderr)
        return 1

    total, malformed = 0, 0
    with args.input.open(encoding="utf-8") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            total += 1
            email = row.get("email") or "<unknown>"
            digest = (row.get("password_hash") or "").strip()
            ok, reason = _is_valid_sha256_hex(digest, args.expect_sha256_hash)
            verdict = "ok" if ok else "rejected"
            print(f"{verdict}\t{email}\t{reason}")
            if not ok:
                malformed += 1

    print(f"summary: total={total} malformed={malformed}", file=sys.stderr)
    if malformed > 0:
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
