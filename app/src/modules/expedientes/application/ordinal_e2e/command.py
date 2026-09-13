"""Commands and outcomes for EXP-CAP-042 E2E ordinal family."""

from dataclasses import dataclass
from typing import Any
from uuid import UUID


class OrdinalE2EError(Exception):
    """Base error for the ordinal use case."""


class OrdinalE2EValidationError(OrdinalE2EError):
    """The family tree is malformed (cycle, unknown parent, multiple roots)."""


class OrdinalE2EAuthorizationError(OrdinalE2EError):
    """The caller has no authenticated actor or lacks permission."""


@dataclass(frozen=True)
class FamilyNode:
    """A node in the E2E ordinal family.

    ``id`` is the stable technical identifier the rest of the E2E
    channel carries. ``parent_id`` is ``None`` for the root; children
    point at their parent. ``ordinal`` is the existing E2E ordinal
    (if any); the service overwrites it with the deterministic
    pre-order value when the family is acyclic.
    """

    id: UUID
    parent_id: UUID | None
    ordinal: int | None = None


@dataclass(frozen=True)
class OrdinalAssignment:
    """The ordinal the service assigned to a single node."""

    node_id: UUID
    ordinal: int


@dataclass(frozen=True)
class OrdinalExpansionResult:
    """The outcome of an ordinal-family expansion.

    ``assignments`` is empty when the family had a cycle; ``cycle`` is
    the set of node IDs that participate in the cycle. Otherwise
    ``root_ordinal`` is the ordinal of the single root.
    """

    assignments: tuple[OrdinalAssignment, ...]
    root_ordinal: int | None
    cycle: frozenset[UUID] = frozenset()


@dataclass(frozen=True)
class OrdinalE2EService:
    """Application-level ordinal/family cycle expansion (CAP-042)."""

    audit_log: Any
    permissions: set[str]
    root_ordinal: int = 1

    def grant(self, permission: str) -> None:
        self.permissions.add(permission)

    async def assign(
        self,
        *,
        actor_id: UUID | None,
        family: list[FamilyNode],
    ) -> OrdinalExpansionResult:
        self._check_actor(actor_id)
        self._check_family_shape(family)
        nodes_by_id: dict[UUID, FamilyNode] = {n.id: n for n in family}
        self._check_known_parents(family, nodes_by_id)
        root = self._find_root(family)
        cycle = self._detect_cycle(nodes_by_id, root.id)
        if cycle:
            return OrdinalExpansionResult((), None, frozenset(cycle))
        assignments = self._walk(root.id, family)
        result = OrdinalExpansionResult(tuple(assignments), self.root_ordinal)
        from app.src.modules.expedientes.application.ordinal_e2e._evidence import (
            ordinal_event,
        )

        assert actor_id is not None
        await self.audit_log.append(ordinal_event(actor_id, result))
        return result

    def _check_actor(self, actor_id: UUID | None) -> None:
        if actor_id is None:
            raise OrdinalE2EAuthorizationError("actor_id is required (deny-by-default)")
        if "e2e.ordinal" not in self.permissions:
            raise OrdinalE2EAuthorizationError("actor lacks permission 'e2e.ordinal'")

    @staticmethod
    def _check_family_shape(family: list[FamilyNode]) -> None:
        if not family:
            raise OrdinalE2EValidationError("family is empty")

    @staticmethod
    def _check_known_parents(family: list[FamilyNode], nodes_by_id: dict[UUID, FamilyNode]) -> None:
        for node in family:
            if node.parent_id is not None and node.parent_id not in nodes_by_id:
                raise OrdinalE2EValidationError(
                    f"unknown parent {node.parent_id!r} for node {node.id!r}"
                )

    @staticmethod
    def _find_root(family: list[FamilyNode]) -> FamilyNode:
        roots = [n for n in family if n.parent_id is None or n.parent_id == n.id]
        if len(roots) != 1:
            raise OrdinalE2EValidationError(f"family must have a single root, found {len(roots)}")
        return roots[0]

    def _walk(self, root_id: UUID, family: list[FamilyNode]) -> list[OrdinalAssignment]:
        assignments: list[OrdinalAssignment] = []
        counter = [self.root_ordinal - 1]

        def visit(node_id: UUID) -> None:
            counter[0] += 1
            assignments.append(OrdinalAssignment(node_id=node_id, ordinal=counter[0]))
            for child in family:
                if child.parent_id == node_id:
                    visit(child.id)

        visit(root_id)
        return assignments

    @staticmethod
    def _detect_cycle(nodes_by_id: dict[UUID, FamilyNode], root_id: UUID) -> set[UUID] | None:
        """Return the set of nodes on the cycle, or ``None`` if acyclic.

        Walks the family from ``root_id`` following parent links; any
        revisit closes a cycle. Raises ``OrdinalE2EValidationError``
        when a node references an unknown parent (so the service can
        reject bad input without producing a silent cycle).
        """
        seen: set[UUID] = set()
        path: set[UUID] = set()

        def visit(node_id: UUID) -> set[UUID] | None:
            if node_id in path:
                return path
            if node_id in seen:
                return None
            path.add(node_id)
            node = nodes_by_id.get(node_id)
            if node is None:
                raise OrdinalE2EValidationError(f"unknown parent {node_id!r} in family")
            if node.parent_id is not None:
                result = visit(node.parent_id)
                if result is not None:
                    return result
            path.discard(node_id)
            seen.add(node_id)
            return None

        return visit(root_id)


__all__ = [
    "FamilyNode",
    "OrdinalAssignment",
    "OrdinalE2EAuthorizationError",
    "OrdinalE2EError",
    "OrdinalE2EService",
    "OrdinalE2EValidationError",
    "OrdinalExpansionResult",
]
