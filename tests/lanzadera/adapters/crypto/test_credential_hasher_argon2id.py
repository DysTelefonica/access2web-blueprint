"""CRITICAL_HELPER tests for :class:`CredentialHasherArgon2id` (DA-2, QC-5).

Coverage target: 100 % for the adapter. Locked at the profile level
(``m=65536, t=3, p=4``) per auth-core/spec.md §Argon2id-only credential
storage. The hash parameter pin is asserted by reading them out of the
PHC string the helper produced — that is the contract Argon2id gives
us and the contract auth-core relies on.
"""

from __future__ import annotations

import pytest

from app.src.modules.lanzadera.adapters.crypto.credential_hasher_argon2id import (
    CredentialHasherArgon2id,
)


def _parse_phc_parameters(phc: str) -> dict[str, int]:
    """Pull the embedded Argon2id parameters out of a PHC string.

    PHC layout: ``$argon2id$v=19$m=<m>,t=<t>,p=<p>[,...]$<salt>$<digest>``.
    We parse manually instead of importing ``argon2.Parameters`` because
    the type moved between argon2-cffi releases and pinning this test
    to one shape would force a dep upgrade just to read parameters
    that the PHC string already encodes textually.
    """
    assert phc.startswith("$argon2id$"), f"not an Argon2id PHC string: {phc!r}"
    # Field 0 is empty before the leading "$", field 1 is "argon2id",
    # field 2 is the version ("v=19"), field 3 is the parameter block.
    parts = phc.split("$")
    assert parts[1] == "argon2id"
    param_block = parts[3]
    return {kv.split("=", 1)[0]: int(kv.split("=", 1)[1]) for kv in param_block.split(",")}


@pytest.fixture
def hasher() -> CredentialHasherArgon2id:
    return CredentialHasherArgon2id()


class TestArgon2idProfile:
    """auth-core/spec.md §Hash uses Argon2id with the pinned profile."""

    def test_hash_is_argon2id_phc_string(self, hasher: CredentialHasherArgon2id) -> None:
        phc = hasher.hash_password("correct horse battery staple")
        # Argon2id PHC strings start with "$argon2id$".
        assert phc.startswith("$argon2id$")

    def test_hash_uses_low_memory_profile(self, hasher: CredentialHasherArgon2id) -> None:
        phc = hasher.hash_password("hunter2hunter2")
        params = _parse_phc_parameters(phc)
        # RFC_9106_LOW_MEMORY profile (m=65536, t=3, p=4). Pin is contract.
        assert params["m"] == 65536  # 64 MiB
        assert params["t"] == 3
        assert params["p"] == 4


class TestHashDistinctness:
    """Salt is random — two hashes of the same plaintext must differ."""

    def test_two_hashes_of_same_plain_differ(self, hasher: CredentialHasherArgon2id) -> None:
        first = hasher.hash_password("same-plaintext")
        second = hasher.hash_password("same-plaintext")
        assert first != second

    def test_two_hashes_both_verify(self, hasher: CredentialHasherArgon2id) -> None:
        # Even though the strings differ, both must verify against the same
        # plaintext — that's the property salts give us.
        plain = "same-plaintext"
        first = hasher.hash_password(plain)
        second = hasher.hash_password(plain)
        assert hasher.verify_password(plain, first) is True
        assert hasher.verify_password(plain, second) is True


class TestVerifyAcceptsAndRejects:
    """auth-core/spec.md §Verify accepts matching plaintext and rejects wrong."""

    def test_verify_accepts_matching_plaintext(self, hasher: CredentialHasherArgon2id) -> None:
        plain = "correct horse battery staple"
        hashed = hasher.hash_password(plain)
        assert hasher.verify_password(plain, hashed) is True

    def test_verify_rejects_wrong_plaintext(self, hasher: CredentialHasherArgon2id) -> None:
        hashed = hasher.hash_password("correct horse battery staple")
        assert hasher.verify_password("wrong", hashed) is False

    def test_verify_rejects_similar_plaintext(self, hasher: CredentialHasherArgon2id) -> None:
        hashed = hasher.hash_password("correct horse battery staple")
        # One character off — must reject.
        assert hasher.verify_password("correct horse battery staplE", hashed) is False

    def test_verify_rejects_empty_plaintext_against_real_hash(
        self, hasher: CredentialHasherArgon2id
    ) -> None:
        hashed = hasher.hash_password("some-real-password")
        assert hasher.verify_password("", hashed) is False


class TestVerifyRejectsEmptyHash:
    """auth-core/spec.md §Empty password attempt is rejected.

    The seed path leaves ``password_hash = NULL``; the helper must not run
    Argon2id against an empty stored hash. The contract is raise, not
    return ``False`` — operators want a loud failure, not a silent
    miss-attempt counter bump.
    """

    def test_verify_rejects_empty_string_hash(self, hasher: CredentialHasherArgon2id) -> None:
        with pytest.raises(ValueError, match="empty"):
            hasher.verify_password("any-plain", "")

    def test_hash_empty_plaintext_is_accepted_but_distinct(
        self, hasher: CredentialHasherArgon2id
    ) -> None:
        # Empty *plaintext* is the caller's responsibility to forbid at the
        # application boundary; the KDF accepts it (Argon2id doesn't refuse
        # empty strings). We just assert it produces a valid PHC that verifies.
        hashed = hasher.hash_password("")
        assert hasher.verify_password("", hashed) is True


class TestVerifyRejectsMalformedHash:
    """Garbage passed as ``hashed`` must return False, not crash."""

    @pytest.mark.parametrize(
        "garbage",
        [
            "not-a-phc-string",
            "$argon2id$v=19$m=1,t=1,p=1$short$short",  # truncated
            "$argon2id$garbage$",
        ],
    )
    def test_malformed_hash_returns_false(
        self, hasher: CredentialHasherArgon2id, garbage: str
    ) -> None:
        assert hasher.verify_password("any-plain", garbage) is False


class TestUnicodeAndLongPlaintext:
    """Real-world passwords are messy — assert the helper keeps working."""

    def test_unicode_plaintext_roundtrip(self, hasher: CredentialHasherArgon2id) -> None:
        plain = "contraseña-ñ-😀-long-enough"
        hashed = hasher.hash_password(plain)
        assert hasher.verify_password(plain, hashed) is True

    def test_long_plaintext_roundtrip(self, hasher: CredentialHasherArgon2id) -> None:
        plain = "x" * 4096  # 4 KiB — well above any sane password max
        hashed = hasher.hash_password(plain)
        assert hasher.verify_password(plain, hashed) is True
