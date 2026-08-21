"""Admin templates (Mística design system tokens).

Pure HTML + ``mds-*`` CSS classes. The Mística tokens are
the canonical Mistica design system class names (the Mística design
system is documented in ``docs/architecture.md`` §UI tokens). No
JavaScript runtime is required; the templates are static fragments
that the FastAPI router serves with ``text/html``.

The templates are kept in a single file per page so the Mística
classes appear next to the data they render. ``base.html`` is the
shared layout; the other templates extend it via Jinja ``{% block %}``.
"""
from __future__ import annotations