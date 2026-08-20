"""Adapters implementing :mod:`app.src.modules.lanzadera.domain.ports`.

* :class:`.crypto.CredentialHasherArgon2id` — Argon2id PHC strings.
* :class:`.crypto.NationalIdCipher` — AES-256-GCM for ``users.dni_encrypted``.
* :class:`.repos.AssignmentRepositoryPg` — DA-12, D22, D42, H11.
* :class:`.repos.ResetTokenRepositoryPg` — D90, DA-4.
* :class:`.repos.GlobalAdminRepositoryPg` — D21, D42.
* :class:`.repos.AuditLogPg` — D27, DA-11, D55.
"""
