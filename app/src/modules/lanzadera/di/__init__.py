"""Composition root — wires every port to its adapter.

The `di` layer is allowed to import any other layer because it is where the
hexagonal inversion is performed (DA-1).
"""

from __future__ import annotations
