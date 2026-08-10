# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 — negative fixture, do not "fix"
"""Deliberately broken domain module. Every import below must be reported by check_layers.py."""

import fastapi  # purity:domain — a pure layer must not touch a framework

from app.src.modules.expedientes.domain import Expediente  # slice:lanzadera->expedientes
from app.src.modules.lanzadera.adapters.repo import UserRepo  # direction:domain->adapters


class User:
    def __init__(
        self, repo: UserRepo, expediente: Expediente, app: fastapi.FastAPI
    ) -> None:
        self.repo = repo
        self.expediente = expediente
        self.app = app
