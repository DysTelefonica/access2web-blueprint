# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 42
# DA-4, D90 — domain exceptions for the reset-flow service.
"""Domain exceptions for the auth-reset flow (DA-4, D90)."""

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
