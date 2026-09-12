"""Strict TDD — AnexoReference value object (CAP-010, CAP-056).

CAP-010 exige "referencia documental opaca" — la referencia que
guarda el sistema es opaca (no expone rutas locales ni del
proveedor). CAP-056 exige "referencia opaca a SharePoint/object
storage con retención y autorización".

Estos tests verifican que el value object encapsule los campos
mínimos (storage_ref, content_type, size_bytes) y las invariantes
básicas que el spec exige:

- La referencia es opaca (string no vacío, sin estructura asumida).
- El tamaño es positivo.
- El content_type sigue el formato MIME estándar.

El límite de tamaño concreto y los niveles de retención son
decisiones de producto (issue #253 H06 + issue #272 R03); este
módulo acepta configuración.
"""

from __future__ import annotations

import pytest

from app.src.modules.expedientes.domain.anexo.reference import (
    AnexoReference,
    AnexoReferenceValidationError,
)


def _valid_ref() -> str:
    return "storage://expedientes/2024/EXP-001/doc-abc123"


def test_constructs_with_minimum_required_fields() -> None:
    """The reference stores opaque ref + content_type + size."""
    ref = AnexoReference(
        storage_ref=_valid_ref(),
        content_type="application/pdf",
        size_bytes=1024,
    )
    assert ref.storage_ref == _valid_ref()
    assert ref.content_type == "application/pdf"
    assert ref.size_bytes == 1024


def test_storage_ref_cannot_be_empty() -> None:
    """CAP-010: reference is opaque; an empty ref is not a valid reference."""
    with pytest.raises(AnexoReferenceValidationError):
        AnexoReference(storage_ref="", content_type="application/pdf", size_bytes=1024)


def test_storage_ref_cannot_be_whitespace_only() -> None:
    with pytest.raises(AnexoReferenceValidationError):
        AnexoReference(storage_ref="   ", content_type="application/pdf", size_bytes=1024)


def test_size_bytes_must_be_positive() -> None:
    """Zero or negative size is rejected — a reference to empty bytes
    is not a valid anexo (the legacy table also enforces Required=True
    on IDDocumento)."""
    with pytest.raises(AnexoReferenceValidationError):
        AnexoReference(storage_ref=_valid_ref(), content_type="application/pdf", size_bytes=0)


def test_size_bytes_must_not_be_negative() -> None:
    with pytest.raises(AnexoReferenceValidationError):
        AnexoReference(storage_ref=_valid_ref(), content_type="application/pdf", size_bytes=-1)


def test_content_type_must_contain_slash() -> None:
    """MIME types have the form ``type/subtype``; an invalid MIME is
    rejected at construction time so the database never stores garbage."""
    with pytest.raises(AnexoReferenceValidationError):
        AnexoReference(storage_ref=_valid_ref(), content_type="application", size_bytes=1024)


def test_content_type_cannot_be_empty() -> None:
    with pytest.raises(AnexoReferenceValidationError):
        AnexoReference(storage_ref=_valid_ref(), content_type="", size_bytes=1024)


def test_content_type_accepts_pdf() -> None:
    ref = AnexoReference(
        storage_ref=_valid_ref(),
        content_type="application/pdf",
        size_bytes=2048,
    )
    assert ref.content_type == "application/pdf"


def test_content_type_accepts_octet_stream() -> None:
    """The generic ``application/octet-stream`` is the right choice for
    binary content whose real type is unknown — common for legacy
    attachments that predate MIME-aware storage."""
    ref = AnexoReference(
        storage_ref=_valid_ref(),
        content_type="application/octet-stream",
        size_bytes=1,
    )
    assert ref.content_type == "application/octet-stream"


def test_content_type_accepts_images_and_documents() -> None:
    """MIME covers a wide range; the constructor only enforces shape."""
    for mime in [
        "image/jpeg",
        "image/png",
        "application/msword",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "application/zip",
    ]:
        ref = AnexoReference(storage_ref=_valid_ref(), content_type=mime, size_bytes=1024)
        assert ref.content_type == mime


def test_content_type_is_case_insensitive_in_storage_but_normalised_lowercase() -> None:
    """RFC 2045 says media types are case-insensitive; we store lower-case
    so equality comparisons are predictable."""
    ref = AnexoReference(
        storage_ref=_valid_ref(),
        content_type="Application/PDF",
        size_bytes=1024,
    )
    assert ref.content_type == "application/pdf"


def test_storage_ref_opaque_property_holds() -> None:
    """CAP-010: the reference is opaque — the constructor does not
    inspect its structure. Local paths, S3 URLs, SharePoint IDs, and
    GUIDs all work because the system treats them as strings."""
    for ref in [
        "s3://bucket/key/path",
        "https://tenant.sharepoint.com/sites/x/Documents/y",
        "simple-uuid-1234",
        "any-opaque-string-with-no-structure",
    ]:
        a = AnexoReference(storage_ref=ref, content_type="application/pdf", size_bytes=1)
        assert a.storage_ref == ref


def test_two_references_with_same_storage_ref_are_equal() -> None:
    """The reference identity is the storage_ref — comparing two
    AnexoReferences with the same ref tells us they're the same
    underlying document regardless of metadata."""
    a = AnexoReference(storage_ref=_valid_ref(), content_type="application/pdf", size_bytes=1024)
    b = AnexoReference(storage_ref=_valid_ref(), content_type="application/pdf", size_bytes=2048)
    assert a == b
    assert hash(a) == hash(b)


def test_different_storage_refs_produce_different_objects() -> None:
    a = AnexoReference(storage_ref="storage://x/a", content_type="application/pdf", size_bytes=1)
    b = AnexoReference(storage_ref="storage://x/b", content_type="application/pdf", size_bytes=1)
    assert a != b
