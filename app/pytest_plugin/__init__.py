"""pytest plugin package for the lanzadera-mvp platform.

The ``coverage_gate`` plugin enforces 100% coverage on the `CRITICAL_HELPERS`
list declared inside the plugin module. Phase 0 ships only the declaration;
Phase 4 (PR 4) ships the implementations the gate then validates.
"""

from __future__ import annotations
