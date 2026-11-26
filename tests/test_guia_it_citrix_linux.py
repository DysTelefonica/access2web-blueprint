"""Structural checks for the portable IT handover."""

import json
from html.parser import HTMLParser
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
GUIDE = ROOT / "docs/07-migracion/guia-it-citrix-linux.html"
DESIGN = ROOT / "design/guia-it-citrix-linux"


class Inventory(HTMLParser):
    def __init__(self):
        super().__init__()
        self.ids = set()
        self.links = []
        self.external = []
        self.inputs = []

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if "id" in attrs:
            self.ids.add(attrs["id"])
        if tag == "a" and attrs.get("href", "").startswith("#"):
            self.links.append(attrs["href"][1:])
        if tag in {"script", "link", "img", "iframe"} and (attrs.get("src") or attrs.get("href")):
            self.external.append(tag)
        if tag == "input":
            self.inputs.append(attrs.get("type", "text"))


def test_portable_document_and_canvas_are_identical():
    document = GUIDE.read_bytes()
    assert document == (DESIGN / "01-guia.html").read_bytes()
    manifest = json.loads((DESIGN / "design.json").read_text(encoding="utf-8"))
    assert manifest["artboards"][0]["file"] == "01-guia.html"
    page = Inventory()
    page.feed(document.decode("utf-8"))
    assert not page.external
    assert set(page.links) <= page.ids
    assert "password" not in page.inputs
    assert "checkbox" in page.inputs


def test_glossary_defines_linked_terms():
    page = Inventory()
    page.feed(GUIDE.read_text(encoding="utf-8"))
    assert "glosario" in page.ids
    glossary_links = [target for target in page.links if target.startswith("g-")]
    assert "g-bitness" in glossary_links
    assert len(set(glossary_links)) >= 30
    assert set(glossary_links) <= page.ids


def test_handover_contains_scenarios_and_actionable_boundaries():
    text = GUIDE.read_text(encoding="utf-8")
    for requirement in (
        "%APPDATA%",
        "H-3",
        "H-4",
        "verify-full",
        "5432",
        "445",
        "443",
        "preparado ≠ validado",
        "ANEXOS",
        "localStorage",
        "@media print",
    ):
        assert requirement in text
