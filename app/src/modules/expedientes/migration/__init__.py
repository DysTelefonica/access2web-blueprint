"""Expedientes migration layer (M01..M04: extractor, staging, promotion).

F04 (issue #225) shipped the ``app/src/modules/expedientes/`` scaffold
without a ``migration/`` subpackage. M01 (issue #257, the
``feat(exp): extractor tranche 01`` WU) populates it with the
read-only Access extractor, the watermark store, the staging writer,
and the orchestration. The other 3 M-WUs (M02 tranche 02, M03 tranche
03, M04 promotion) live in the same subpackage and reuse the primitives
this WU delivers.

The extractor is the boundary between the legacy backend
(``Expedientes_datos.accdb``) and the platform's ``expedientes``
schema. It is a structural walker, not a metric: it neither
consumes nor invalidates the numbers the ``quality_report.py`` gates
track. The watermark store is the idempotency mechanism (D-EXP-9,
"manifest/watermarks") — once a PK lands in staging, the next run
skips it. The promotion layer (M04) reads the watermark and
``count_rows`` to drive the reconciliation gate.
"""

from __future__ import annotations
