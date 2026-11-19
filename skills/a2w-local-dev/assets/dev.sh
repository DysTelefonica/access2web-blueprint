#!/usr/bin/env bash
# Per-issue local workflow for access2web (#1776): unit tests, integration against Postgres,
# the app with its migrations, Playwright and `make verify`, all inside containers, the same on
# Linux and on Windows through WSL. Mirrors the CI jobs of .github/workflows/ci.yml.
#
#   dev.sh <issue> up | test [pytest -k expr] | integration | e2e [specs] | verify | down | env
#
# On Windows run it from WSL: a worktree created by the Windows git keeps a `gitdir: C:/...`
# that WSL cannot open, so the script exports GIT_DIR and GIT_WORK_TREE. No secret file is
# read: the app gets the same development values the CI uses.
set -euo pipefail

usage() {
    echo "uso: dev.sh <nº issue> up | test [expr] | integration | e2e [specs] | verify | down | env" >&2
    exit 2
}

[ $# -ge 2 ] || usage
ISSUE=$1
CMD=$2
shift 2
[[ "$ISSUE" =~ ^[0-9]+$ ]] || usage
case "$CMD" in up|test|integration|e2e|verify|down|env) ;; *) usage ;; esac

ROOT=${A2W_ROOT:-$(cd "$(dirname "${BASH_SOURCE[0]}")/../../.." && pwd)}

# One stack per issue: its own compose project and ports, so two issues can run at once.
PROJECT="a2w$ISSUE"
APP_PORT=$((28000 + ISSUE % 1000))
PG_PORT=$((25000 + ISSUE % 1000))

WINDOWS=0
if [ -f "$ROOT/.git" ]; then
    gitdir=$(sed -n 's/^gitdir: //p' "$ROOT/.git" | tr -d '\r')
    if [[ "$gitdir" =~ ^([A-Za-z]):/(.*)$ ]]; then
        WINDOWS=1
        drive=$(printf '%s' "${BASH_REMATCH[1]}" | tr '[:upper:]' '[:lower:]')
        export GIT_DIR="/mnt/$drive/${BASH_REMATCH[2]}"
        export GIT_WORK_TREE="$ROOT"
    fi
fi

# Development values, the same the CI uses (ci.yml, integration and e2e-smoke). Not secrets.
PG_URL="postgresql+asyncpg://lanzadera:lanzadera-dev-only@postgres:5432/lanzadera"
DEV_SECRET_KEY="a2w-local-dev-key-not-a-real-secret"
DEV_PLATFORM_KEY="MDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDA="
PYTHON_IMAGE="python:3.12.11"
COMPOSE_FILE="${TMPDIR:-/tmp}/$PROJECT-compose.yml"

write_compose() {
    cat > "$COMPOSE_FILE" <<EOF
name: $PROJECT
services:
  postgres:
    image: postgres:16-alpine
    environment:
      POSTGRES_DB: lanzadera
      POSTGRES_USER: lanzadera
      POSTGRES_PASSWORD: lanzadera-dev-only
    ports: ["$PG_PORT:5432"]
    healthcheck:
      test: ["CMD-SHELL", "pg_isready -U lanzadera -d lanzadera"]
      interval: 2s
      retries: 30
  app:
    build: {context: "$ROOT/app", dockerfile: Dockerfile}
    environment:
      DATABASE_URL: "$PG_URL"
      SECRET_KEY: "$DEV_SECRET_KEY"
      PLATFORM_SECRET_KEY: "$DEV_PLATFORM_KEY"
    ports: ["$APP_PORT:8000"]
    depends_on: {postgres: {condition: service_healthy}}
EOF
}

compose() {
    write_compose
    docker compose -f "$COMPOSE_FILE" "$@"
}

# Python with app/requirements-dev.lock and the CI's tools (jq), built once per lock content and
# recipe version, and reused. Bump TOOLS_RECIPE when the recipe below changes.
TOOLS_RECIPE=2
tools_image() {
    local tag
    tag="a2w-tools:r$TOOLS_RECIPE-$(sha256sum "$ROOT/app/requirements-dev.lock" | cut -c1-12)"
    if ! docker image inspect "$tag" >/dev/null 2>&1; then
        docker build -q -t "$tag" -f - "$ROOT/app" >/dev/null <<EOF
FROM $PYTHON_IMAGE
COPY requirements-dev.lock /tmp/requirements-dev.lock
RUN apt-get update && apt-get install -y --no-install-recommends jq && rm -rf /var/lib/apt/lists/*
RUN pip install --no-cache-dir --require-hashes --requirement /tmp/requirements-dev.lock
EOF
    fi
    printf '%s' "$tag"
}

# Runs a command in the tools image with the worktree mounted; extra docker args go first.
tools() {
    local image args=()
    image=$(tools_image)
    while [ $# -gt 0 ] && [ "$1" != "--" ]; do args+=("$1"); shift; done
    shift
    docker run --rm -u "$(id -u):$(id -g)" -e HOME=/tmp -v "$ROOT:/work" -w /work \
        "${args[@]}" "$image" "$@"
}

wait_for_health() {
    for _ in $(seq 1 60); do
        if curl -fsS "http://localhost:$APP_PORT/health" >/dev/null 2>&1; then return 0; fi
        sleep 2
    done
    echo "FAIL: la app del stack $PROJECT no responde en /health tras 120 s." >&2
    compose logs --tail 80 app >&2
    exit 1
}

# The test suites only allow sockets to 127.0.0.1 (pytest-socket), so the tools container shares
# the host network and reaches the stack's Postgres through its published port, as in the CI.
HOST_PG_URL="postgresql+asyncpg://lanzadera:lanzadera-dev-only@127.0.0.1:$PG_PORT/lanzadera"

migrate() {
    tools --network host -e DATABASE_URL="$HOST_PG_URL" -e PLATFORM_SECRET_KEY="$DEV_PLATFORM_KEY" -- \
        alembic -c app/migrations/alembic.ini upgrade head
}

case "$CMD" in
env)
    echo "PROJECT=$PROJECT"
    echo "APP_PORT=$APP_PORT"
    echo "PG_PORT=$PG_PORT"
    echo "WINDOWS=$WINDOWS"
    echo "GIT_DIR=${GIT_DIR:-}"
    echo "GIT_WORK_TREE=${GIT_WORK_TREE:-}"
    ;;
up)
    compose up -d --wait postgres
    migrate
    compose up -d --build app
    wait_for_health
    echo "Stack $PROJECT: http://localhost:$APP_PORT (Postgres $PG_PORT)"
    ;;
test)
    if [ $# -gt 0 ]; then
        tools -- pytest -c app/pyproject.toml --rootdir=app -m "not integration" -k "$1"
    else
        tools -- pytest -c app/pyproject.toml --rootdir=app -m "not integration"
    fi
    ;;
integration)
    compose up -d --wait postgres
    migrate
    tools --network host -e APAP_INTEGRATION_ENABLED=1 \
        -e APAP_INTEGRATION_URL="$HOST_PG_URL" -e INTEGRATION_DSN="$HOST_PG_URL" -- \
        pytest -c app/pyproject.toml --rootdir=app -m integration
    ;;
e2e)
    wait_for_health
    # The browsers of the image must match the locked @playwright/test, not the range in package.json.
    version=$(grep -A2 '"node_modules/@playwright/test"' "$ROOT/e2e/package-lock.json" | sed -n 's/.*"version": "\([0-9.]*\)".*/\1/p' | head -n1)
    [ -n "$version" ] || { echo "FAIL: no se encuentra la versión de @playwright/test en e2e/package-lock.json." >&2; exit 1; }
    # The global setup seeds the test users with `docker exec` into the app container, so the
    # Playwright container gets the docker CLI and socket; node_modules lives in a volume.
    docker run --rm --network "${PROJECT}_default" -e HOME=/tmp \
        -e BASE_URL=http://app:8000 -e TEST_CONTAINER_NAME="${PROJECT}-app-1" \
        -v /var/run/docker.sock:/var/run/docker.sock -v "$(readlink -f "$(command -v docker)"):/usr/local/bin/docker:ro" \
        -v "${PROJECT}-e2e-node-modules:/e2e/node_modules" -v "$ROOT/e2e:/e2e" -w /e2e \
        "mcr.microsoft.com/playwright:v${version}-noble" \
        sh -c 'npm ci --no-audit --no-fund && npx playwright test "$@"' sh "$@"
    ;;
verify)
    # Only the preflight reads the worktree's git (the diff against origin/main); the test
    # suites create their own repositories, which an exported GIT_DIR would break.
    git_args=(-e GIT_CONFIG_COUNT=1 -e GIT_CONFIG_KEY_0=safe.directory -e GIT_CONFIG_VALUE_0='*')
    if [ "$WINDOWS" = 1 ]; then
        common=$(dirname "$(dirname "$GIT_DIR")")
        git_args+=(-v "$common:$common" -e GIT_DIR="$GIT_DIR" -e GIT_WORK_TREE=/work)
    fi
    tools "${git_args[@]}" -- bash scripts/local-preflight.sh
    ;;
down)
    compose down -v --remove-orphans
    docker volume rm -f "${PROJECT}-e2e-node-modules" >/dev/null
    rm -f "$COMPOSE_FILE"
    ;;
esac
