"""Glossary of the admin IT handover guide (#1810)."""

from app.src.modules.lanzadera.delivery.http.guia_it_glossary import (
    GLOSSARY,
    GLOSSARY_TERMS,
)


def test_glossary_keys_are_unique_and_every_entry_is_defined():
    keys = [entry.key for entry in GLOSSARY]
    assert len(keys) == len(set(keys))
    for entry in GLOSSARY:
        assert entry.key and entry.term and entry.definition.strip()


def test_glossary_defines_bitness_and_indexes_every_term():
    assert "bitness" in GLOSSARY_TERMS
    assert GLOSSARY_TERMS == {entry.key: entry.term for entry in GLOSSARY}
