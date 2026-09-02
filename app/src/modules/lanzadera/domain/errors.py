# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W62 (#539)
# DA-4, D38, D39, D90 — domain exceptions for the auth flows.
"""Domain exceptions for the auth flows (DA-4, D38, D39, D90, D-W62-2)."""

from __future__ import annotations


class NoGlobalAdminError(Exception):
    """Reset-by-email is disabled until at least one global admin exists (D90)."""


class InvalidResetTokenError(Exception):
    """Token is unknown or superseded."""


class ExpiredResetTokenError(Exception):
    """Token exists but is past `expires_at` (D90, 24 h TTL)."""


class ResetTokenAlreadyUsedError(Exception):
    """Token has already been consumed (D90, single-use)."""


class UserNotFoundError(Exception):
    """`issue_reset_token` was called for an email that matches no user."""


class DuplicateEmailError(Exception):
    """`create_user` was called for an email that already maps to a row."""


class ProfileNotFoundError(Exception):
    """`assign_profile` was called for a profile that does not exist on the app."""


class InvalidCredentialsError(Exception):
    """`login` was called with an unknown email or a wrong password.

    The two cases are deliberately not distinguished so the response
    time cannot be used to enumerate users (D-W62-2, D38).
    """


class AccountLockedError(Exception):
    """`login` was called for a user whose `failed_attempts` hit the
    lockout threshold within the last `LockoutPolicy.duration_seconds`
    (D38, D39).

    The caller (delivery) maps this to HTTP 423 Locked.
    """


class AccountNotActiveError(Exception):
    """`login` was called for a user whose status is not ``ACTIVE``
    (PASSWORD_RESET_REQUIRED, DISABLED, etc.).

    The audit row records the attempt but the counter is not incremented;
    the global admin or reset-flow path is the only way out of a
    non-ACTIVE state.
    """


class SessionNotFoundError(Exception):
    """`logout` was called for a session id that does not exist.

    The caller (delivery, PR-5) maps this to HTTP 401 to avoid leaking
    which sessions were once valid; the audit row records the attempt
    so the security log can correlate brute-force probes.
    """


class InvalidTokenError(Exception):
    """`JwtSignerPort.verify` was called with a token that does not
    parse, has an invalid signature, or has a malformed segment
    (header / payload / signature). The auth middleware (PR-5)
    catches this and lets the request through with
    ``request.state.user_id = None`` — the delivery routes' own
    ``require_global_admin`` then returns HTTP 401.
    """


class ExpiredTokenError(Exception):
    """`JwtSignerPort.verify` was called with a token whose payload
    ``exp`` is in the past (``now >= exp``). The auth middleware
    (PR-5) catches this the same way as ``InvalidTokenError``.
    """
