"""Static content of the admin IT handover guide (Citrix + Linux).

Facts come from docs/07-migracion/epica-coexistencia-postgresql.md
(decisions 10, 18-20, 28; H-3, H-4) and docs/architecture.md (D93).
The guide prepares an isolated rehearsal; it never authorises the
production data move. The sections are data, not code: they live in
``guia_it_content.json`` and reach the template as plain mappings.
"""

from __future__ import annotations

from app.src.modules.lanzadera.delivery.http.guia_it_glossary import (
    GLOSSARY,
    GLOSSARY_TERMS,
    load_json,
)

GUIDE_CONTEXT: dict[str, object] = {
    "glossary": GLOSSARY,
    "glossary_terms": GLOSSARY_TERMS,
    **load_json("guia_it_content.json"),
}

__all__ = ["GUIDE_CONTEXT"]
