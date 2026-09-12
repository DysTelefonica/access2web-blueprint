"""Composition root for the Expedientes module (F03, issue #224).

D-EXP-3 + issue #224 acceptance: deny/no fake prod. The container refuses
to bind fakes in production and exposes the bound ports as read-only
properties so the delivery layer (HTTP routes, CLI driver) never touches
a port through anything other than the container.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING, cast

from app.src.modules.expedientes.di.config import (
    AppEnv,
    ConfigValidationError,
    ExpedientesConfig,
    ProductionFakeRefused,
)

if TYPE_CHECKING:
    from app.src.modules.expedientes.ports.audit_log import AuditLogPort
    from app.src.modules.expedientes.ports.catalog_repository import CatalogRepositoryPort
    from app.src.modules.expedientes.ports.document_storage import DocumentStoragePort
    from app.src.modules.expedientes.ports.expediente_repository import ExpedienteRepositoryPort
    from app.src.modules.expedientes.ports.hito_repository import HitoRepositoryPort
    from app.src.modules.expedientes.ports.notification_delivery import NotificationDeliveryPort
    from app.src.modules.expedientes.ports.readiness import ReadinessPort


_FAKE_TYPE_MARKERS = (
    "Fake",  # tests/expedientes/adapters/_fakes.py
)


def _is_fake(port: object) -> bool:
    """Detect a port binding that is a deterministic test double.

    Heuristic: type name starts with ``Fake`` (the convention used in
    ``tests/expedientes/adapters/_fakes.py``). This is intentionally
    conservative — adding a real adapter whose class happens to start
    with ``Fake`` is a code smell anyway, and the rename would surface
    it in review.
    """
    cls = type(port)
    return cls.__name__.startswith(_FAKE_TYPE_MARKERS)


@dataclass(frozen=True)
class ExpedientesContainer:
    """The Expedientes composition root (F03, issue #224).

    Built by :meth:`build`. Exposes the 7 ports defined in F01 as
    read-only attributes. The delivery layer (HTTP routes / CLI driver)
    reads them through the properties — never through the constructor
    again, so the production guard runs exactly once at startup.
    """

    config: ExpedientesConfig
    expediente_repo: ExpedienteRepositoryPort = field(repr=False)
    hito_repo: HitoRepositoryPort = field(repr=False)
    catalog_repo: CatalogRepositoryPort = field(repr=False)
    audit_log: AuditLogPort = field(repr=False)
    readiness: ReadinessPort = field(repr=False)
    document_storage: DocumentStoragePort = field(repr=False)
    notification_delivery: NotificationDeliveryPort = field(repr=False)

    @classmethod
    def build(
        cls,
        *,
        config: ExpedientesConfig,
        ports: dict[str, object],
    ) -> ExpedientesContainer:
        """Construct the container, applying the production guard.

        ``ports`` must contain every key in :data:`_REQUIRED_PORTS`.
        In production, every value must be a real adapter (not a fake).
        """
        missing = _REQUIRED_PORTS - set(ports)
        if missing:
            raise ConfigValidationError(f"missing ports for container: {sorted(missing)}")

        if config.app_env is AppEnv.PRODUCTION:
            fakes_found = {
                name: type(port).__name__ for name, port in ports.items() if _is_fake(port)
            }
            if fakes_found:
                raise ProductionFakeRefused(
                    "production refuses to bind fakes (D-EXP-3): "
                    + ", ".join(f"{k}={v}" for k, v in fakes_found.items())
                )

        return cls(
            config=config,
            expediente_repo=cast(ExpedienteRepositoryPort, ports["expediente_repo"]),
            hito_repo=cast(HitoRepositoryPort, ports["hito_repo"]),
            catalog_repo=cast(CatalogRepositoryPort, ports["catalog_repo"]),
            audit_log=cast(AuditLogPort, ports["audit_log"]),
            readiness=cast(ReadinessPort, ports["readiness"]),
            document_storage=cast(DocumentStoragePort, ports["document_storage"]),
            notification_delivery=cast(NotificationDeliveryPort, ports["notification_delivery"]),
        )


_REQUIRED_PORTS = frozenset(
    {
        "expediente_repo",
        "hito_repo",
        "catalog_repo",
        "audit_log",
        "readiness",
        "document_storage",
        "notification_delivery",
    }
)


__all__ = ["ExpedientesContainer"]
