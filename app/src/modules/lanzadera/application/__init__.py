"""Use cases — pure orchestration.

Purity rule: application layer may import `domain`, `ports`, and other
application modules. Adapters and frameworks are forbidden here (DA-1).
"""

from __future__ import annotations
