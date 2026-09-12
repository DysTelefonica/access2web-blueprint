"""AnexoReference — value object para referencias documentales opacas (CAP-010).

CAP-010 exige "referencia documental opaca" — el sistema guarda la
referencia que el adapter de almacenamiento devuelve, sin asumir
estructura (URL S3, GUID SharePoint, ruta local ofuscada, etc.).
El módulo conserva los tres campos mínimos del legacy
``TbExpedientesAnexos`` (IDDocumento / IDExpediente / NombreDocumento)
más el contrato del ``DocumentStoragePort``
(``app/src/modules/expedientes/ports/document_storage.py``).

DA-1: pure domain — no framework imports.
"""

from __future__ import annotations

from dataclasses import dataclass


class AnexoReferenceValidationError(ValueError):
    """Raised when a string is not a valid AnexoReference.

    The error covers the value-object invariants (non-empty storage_ref,
    non-empty MIME type with ``/``, positive size). Semantic validation
    (does this ref exist in the master storage?) is the adapter's job.
    """


@dataclass(frozen=True)
class AnexoReference:
    """A reference to a stored document (CAP-010, CAP-056).

    Storage references are deliberately opaque: the system does not
    interpret the string. Whether the adapter returns an S3 URL,
    a SharePoint ID, or a local GUID, the domain treats them as
    strings. CAP-010 §Camino feliz: "se registra la referencia y su
    retención sin exponer rutas locales".

    Equality compares the storage_ref only — two anexos pointing at
    the same stored document are the same reference regardless of
    any metadata drift in the secondary fields.

    The MIME ``content_type`` is normalised to lowercase on storage
    (RFC 2045: media types are case-insensitive). Equality still
    compares the normalised form.
    """

    storage_ref: str
    content_type: str
    size_bytes: int

    def __post_init__(self) -> None:
        ref = self.storage_ref.strip()
        if not ref:
            raise AnexoReferenceValidationError("storage_ref cannot be empty or whitespace-only")

        mime = self.content_type.strip().lower()
        if "/" not in mime or mime.startswith("/") or mime.endswith("/"):
            raise AnexoReferenceValidationError(
                f"content_type must be a MIME type (type/subtype), got {self.content_type!r}"
            )

        if self.size_bytes <= 0:
            raise AnexoReferenceValidationError(
                f"size_bytes must be positive, got {self.size_bytes}"
            )

        # Commit normalised form back to the frozen instance.
        object.__setattr__(self, "storage_ref", ref)
        object.__setattr__(self, "content_type", mime)

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, AnexoReference):
            return NotImplemented
        return self.storage_ref == other.storage_ref

    def __hash__(self) -> int:
        return hash(self.storage_ref)
