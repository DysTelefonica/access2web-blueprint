# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp PR 2
# DA-1, D21, D42 — `GlobalAdmin` is a frozen value object: the entity is the
# row in `global_admins`, and `user_id` is the only attribute. The presence
# table is the only source of truth (D21); no column on `users` materialises
# the role.
"""Strict TDD — `GlobalAdmin` value object (Phase 1, task 1.6)."""

from __future__ import annotations

import dataclasses
from uuid import uuid4

import pytest

from app.src.modules.lanzadera.domain.global_admin import GlobalAdmin


# ---------------------------------------------------------------------------
# Construction
# ---------------------------------------------------------------------------


class TestGlobalAdminConstruction:
    def test_global_admin_carries_user_id(self) -> None:
        user_id = uuid4()
        admin = GlobalAdmin(user_id=user_id)
        assert admin.user_id == user_id


# ---------------------------------------------------------------------------
# Immutability — value object
# ---------------------------------------------------------------------------


class TestGlobalAdminImmutability:
    def test_global_admin_is_frozen(self) -> None:
        admin = GlobalAdmin(user_id=uuid4())
        with pytest.raises(dataclasses.FrozenInstanceError):  # type: ignore[misc]
            admin.user_id = uuid4()  # type: ignore[misc]

    def test_two_admins_same_user_compare_equal(self) -> None:
        user_id = uuid4()
        a = GlobalAdmin(user_id=user_id)
        b = GlobalAdmin(user_id=user_id)
        assert a == b

    def test_distinct_users_compare_unequal(self) -> None:
        a = GlobalAdmin(user_id=uuid4())
        b = GlobalAdmin(user_id=uuid4())
        assert a != b

    def test_global_admin_is_hashable(self) -> None:
        """Value objects live in sets; membership tests rely on `__hash__`."""
        user_id = uuid4()
        admins = {GlobalAdmin(user_id=user_id), GlobalAdmin(user_id=user_id)}
        assert len(admins) == 1
