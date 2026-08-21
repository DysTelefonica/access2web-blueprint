"""Expedientes migration package.

Scaffolding for the ``app/src/modules/expedientes/`` module that the
expedientes-web-migration change introduces. F04 (issue #225) ships
the directory structure + the Unit of Work helper; verticals (C01..C04,
R01..R07, Q01..Q02) fill the ``domain/``, ``ports/``, and
``adapters/`` subpackages as the chain progresses.

The empty submodules are placeholders; they import nothing and export
nothing. They are kept here so that imports from later verticals
(``from app.src.modules.expedientes.domain import ...``) resolve to
real packages instead of ``ModuleNotFoundError`` during the rollout.
"""

from __future__ import annotations
