"""Admin HTTP delivery — D91 placeholder (DL2, issue #55).

W58 (#515) wired the HTTP routes (``admin_routes.py``,
``admin_routes_users.py``, ``admin_routes_misc.py``) to the
``LanzaderaContainer`` directly. The router is now mounted by
``main.py`` through ``register_routes(router, templates=..., container=...)``
during application startup, so the previous ``build_router`` factory
in this module is no longer reachable from production code.

This module keeps ``require_global_admin`` (the D91 placeholder gate
that ``admin_routes_users.py`` and ``admin_routes_misc.py`` look up via
``_admin.require_global_admin()`` so tests can monkeypatch the attribute
on the ``admin`` module and the destructive commands still propagate).
"""

from __future__ import annotations

__all__ = ["require_global_admin"]


def require_global_admin() -> None:
    """Placeholder for D91's global-admin gate.

    The real implementation checks ``admins.is_global_admin(current_user.id)``
    from a session or header. This WU ships the route shape; the
    gating comes with M02's auth wiring (A01..A03).
    """
