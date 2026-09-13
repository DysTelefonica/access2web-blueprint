"""DocumentStoragePort — CAP-010, D-EXP-6."""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, BinaryIO, Protocol

if TYPE_CHECKING:
    from uuid import UUID


@dataclass
class StoredDocument:
    storage_ref: str  # Opaque.
    content_type: str
    size_bytes: int


class DocumentStoragePort(Protocol):
    """Contenido binario de anexos. Proveedor (S3, SharePoint) queda en adapter."""

    async def prepare_upload(
        self, expediente_id: UUID, filename: str, content_type: str, size_bytes: int
    ) -> str: ...
    async def confirm_upload(self, storage_ref: str) -> StoredDocument: ...
    async def download(self, storage_ref: str) -> BinaryIO: ...
    async def delete(self, storage_ref: str) -> None: ...
