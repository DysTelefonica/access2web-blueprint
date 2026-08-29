# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W55
"""Process-startup bootstrap step (CA-F4, D91, DA-1).

Reads the initial admin list from the configured
``BootstrapAdminSource`` and ensures every address becomes a global
admin (D42 invariant: at least one global admin must exist once
the bootstrap completes; D91: the CLI ``set-password`` is the
exclusive path for the first admin's password).

The function is a pure no-op when the source is empty: the
``bootstrap_global_admins`` use case returns 0 new rows and writes
no audit events, so an environment without ``$GLOBAL_ADMIN_EMAILS``
starts the platform cleanly.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from app.src.modules.lanzadera.di.container import LanzaderaContainer


async def run_bootstrap(container: LanzaderaContainer) -> int:
    """Run the global-admin seed step.

    Returns the number of NEW global-admin rows created (zero on
    every re-run after the first). The actual addresses are read by
    ``container.bootstrap_source``; the use case is invoked with the
    container's own session-scoped adapters and clock so the audit
    row is timestamped with the container's notion of ``now``.
    """
    return await container.bootstrap_global_admins()


__all__ = ["run_bootstrap"]
