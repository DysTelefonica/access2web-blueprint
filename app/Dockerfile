# Lanzadera MVP — multi-stage Dockerfile.
#
# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 — applies to lanzadera-mvp Phase 0.
# Every `FROM` is pinned by SHA256 digest (Hard Rule 6). Tags mutate silently;
# digests pin byte-for-byte content. Resolved at adoption (2026-08-09) via the
# Docker Hub registry API; re-resolve when bumping Python or switching base OS.

# ---------------------------------------------------------------------------
# Stage 1 — builder: install the package and its dev dependencies into a
# throwaway environment. The final image carries only the runtime wheels.
# ---------------------------------------------------------------------------
FROM python:3.12-slim-bookworm@sha256:4766d8b510c428e595d74b9cc5bbb2fae8e26316fffb4adc89908d79aacd58a2 AS builder

ENV PIP_DISABLE_PIP_VERSION_CHECK=1 \
    PIP_NO_CACHE_DIR=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /build

# System dependencies needed to build argon2-cffi and asyncpg wheels.
RUN apt-get update \
    && apt-get install --no-install-recommends --yes \
        build-essential \
        libffi-dev \
        libssl-dev \
        libpq-dev \
    && rm -rf /var/lib/apt/lists/*

# Copy the package definition first so the wheel cache is reusable.
COPY pyproject.toml ./

# Install runtime + dev dependencies into a virtualenv at /opt/venv.
RUN python -m venv /opt/venv \
    && /opt/venv/bin/pip install --upgrade pip==24.2 \
    && /opt/venv/bin/pip install --no-cache-dir .

# ---------------------------------------------------------------------------
# Stage 2 — runtime: copy only the venv and the source. No build tooling, no
# apt cache, no source tree that includes tests.
# ---------------------------------------------------------------------------
FROM python:3.12-slim-bookworm@sha256:4766d8b510c428e595d74b9cc5bbb2fae8e26316fffb4adc89908d79aacd58a2 AS runtime

ENV PATH="/opt/venv/bin:${PATH}" \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    PIP_NO_CACHE_DIR=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PYTHONHASHSEED=random

# `appuser` exists on the base image; we just assign a non-root UID.
RUN useradd --create-home --uid 10001 --shell /usr/sbin/nologin appuser

WORKDIR /app
COPY --from=builder /opt/venv /opt/venv
COPY src/ /app/src/

USER appuser

EXPOSE 8000

# uvicorn binds to 0.0.0.0:8000. `--proxy-headers` honours X-Forwarded-* when
# fronted by an ingress. `forwarded-allow-ips=*` is the conservative default
# for internal Coolify deploys; tighten per-environment in production.
CMD ["uvicorn", "app.src.main:app", "--host", "0.0.0.0", "--port", "8000", "--proxy-headers", "--forwarded-allow-ips", "*"]