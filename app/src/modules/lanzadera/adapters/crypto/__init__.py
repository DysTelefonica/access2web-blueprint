"""Crypto adapters for the Lanzadera module.

* :class:`CredentialHasherArgon2id` — Argon2id PHC strings, profile
  ``RFC_9106_LOW_MEMORY`` (auth-core/spec.md §Argon2id-only credential
  storage).
* :class:`NationalIdCipher` — AES-256-GCM for ``users.dni_encrypted``,
  keyed via :class:`SecretManagerPort` (CA-S4).
"""
