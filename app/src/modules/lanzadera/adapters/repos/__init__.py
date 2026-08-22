"""Legacy re-exports — the four adapters that lived at
``app.src.modules.lanzadera.adapters.repos.*`` have all been
migrated to ``app.src.modules.lanzadera.adapters.persistence``
across W02..W05 (#45, #55, #21, #42-subset). This stub
remains for transitional imports; the explicit redirects in
``adapters/__init__.py`` keep the public surface stable.
"""

from __future__ import annotations

from typing import Final

__all__: Final[list[str]] = []
