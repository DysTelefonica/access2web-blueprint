"""Conftest para pytest-socket — fragmento portable (#528).

Activa `--disable-socket` por defecto y abre red solo a los hosts
declarados en `PYTEST_SOCKET_ALLOW` (CSV) o pasados a la fixture
`allow_network`. Tolerante la falta de `pytest-socket` para que el
fichero se pueda copiar sin instalar la dependencia.

Plantilla copy-paste: la auditoría HR-28 sigue siendo el gate
`check_env_isolation.py`; este conftest es la materialización en el
lado del test runner.
"""

from __future__ import annotations

import os

try:  # pytest-socket es opcional.
    import pytest_socket  # type: ignore[import-untyped]
    _HAS_PYTEST_SOCKET = True
except ImportError:
    pytest_socket = None  # type: ignore[assignment]
    _HAS_PYTEST_SOCKET = False

import pytest


def pytest_addoption(parser: pytest.Parser) -> None:
    if not _HAS_PYTEST_SOCKET:
        return
    parser.addini("markers", "socket_enabled: marca un test como autorizado a usar red")


def pytest_collection_modifyitems(config: pytest.Config, items: list[pytest.Item]) -> None:
    if not _HAS_PYTEST_SOCKET:
        return
    allowed = os.environ.get("PYTEST_SOCKET_ALLOW", "").strip()
    if allowed:
        pytest_socket.enable_socket(allowed_hosts=allowed)


@pytest.fixture
def allow_network() -> "_AllowNetwork":
    return _AllowNetwork()


class _AllowNetwork:
    """Permite temporalmente hosts en la allowlist de pytest-socket."""

    def __call__(self, hosts) -> None:
        if not _HAS_PYTEST_SOCKET:
            pytest.skip("pytest-socket not installed; cannot toggle network")
        if isinstance(hosts, str):
            pytest_socket.enable_socket(allowed_hosts=hosts)
        else:
            pytest_socket.enable_socket(allowed_hosts=",".join(hosts))

    def restore(self) -> None:
        if _HAS_PYTEST_SOCKET:
            pytest_socket.disable_socket()