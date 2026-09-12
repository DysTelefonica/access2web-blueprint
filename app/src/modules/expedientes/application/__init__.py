"""Expedientes application layer (C01+ verticals)."""

from app.src.modules.expedientes.application.create_expediente import (
    CreateExpedienteError,
    create_expediente,
)

__all__ = ["CreateExpedienteError", "create_expediente"]
