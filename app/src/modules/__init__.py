"""All hexagonal modules live under `platform.src.modules.<module>`.

Gate `scripts/check_layers.py` enforces dependency direction between layers and
vertical slicing between modules. `ROOT_PACKAGE = "platform.src.modules"` is
the constant the gate reads.
"""

from __future__ import annotations
