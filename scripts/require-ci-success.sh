#!/usr/bin/env bash
# CI gate — refuses to release a commit whose CI is not green.
#
# Release and CI are independent workflows over the same commit. Without this
# check the release path never learns that CI was red on the exact bytes it is
# about to publish, and a tag would be created over a known-broken tree.
#
# This closes it for the whole of CI rather than for one job. Adding the E2E
# matrix to the release workflow would cover today's gap and miss the next job
# someone adds; requiring the CI run itself covers every job CI already has and
# every job it grows.
#
# Issue #857, PR-C1c: el criterio pasa al de gentle-ai:
#   1. Toma el run MÁS RECIENTE del workflow ci.yml para ese SHA, esté como
#      esté. Si todavía no está `completed`, falla con «wait for it to finish,
#      then re-run this release from the same tag».
#   2. NO filtra por evento: vale cualquier run de ci.yml sobre el SHA exacto
#      (pull_request, push, schedule, workflow_dispatch), porque los cuatro
#      eventos evalúan el mismo commit.
#   3. Busca por fichero (`actions/workflows/ci.yml/runs?head_sha=…`), no por
#      nombre: el nombre del workflow cambia con el repo, la ruta del fichero
#      no.
#   4. Los mensajes de error dicen qué hacer (push + wait, fix + re-tag,
#      re-run desde el mismo tag).
#
#     GH_TOKEN=... ./scripts/require-ci-success.sh [commit-sha]
#
# Contract: docs/11-releases.md
set -euo pipefail

die() {
  printf 'ci gate: %s\n' "$*" >&2
  exit 1
}

sha=${1:-${GITHUB_SHA:-$(git rev-parse HEAD)}}
repo=${GITHUB_REPOSITORY:-}
workflow_name=ci

if [[ -z "$repo" ]]; then
  repo=$(git remote get-url origin 2>/dev/null |
    sed -E 's#^.*github\.com[:/]##; s#\.git$##') ||
    die 'could not resolve the repository; set GITHUB_REPOSITORY'
fi
[[ -n "$repo" ]] || die 'could not resolve the repository; set GITHUB_REPOSITORY'

command -v gh >/dev/null || die 'the GitHub CLI (gh) is required'
command -v jq >/dev/null || die 'jq is required'

# File-based lookup of `ci.yml` runs for the exact SHA. Per issue #857 PR-C1c
# this drops the per-event filter that the previous version of the script
# used (`event == "merge_group" | "schedule"`): a single SHA can be evaluated
# by `pull_request`, `push`, `schedule` or `workflow_dispatch`, and the
# verdict is the run, not the event.
runs=$(gh api --paginate \
  "repos/${repo}/actions/workflows/ci.yml/runs?head_sha=${sha}&per_page=100" \
  --jq '.workflow_runs[] | [.created_at, .id, .status, (.conclusion // "none"), .html_url] | @tsv') ||
  die "could not read workflow runs for ${sha}"

if [[ -z "$runs" ]]; then
  die "no ${workflow_name} run exists for ${sha}. Push the commit to a branch that runs ${workflow_name}, wait for it to finish, then re-run this release from the same tag"
fi

# A commit can carry several runs of the same workflow: a re-run, or a retry
# after an infrastructure failure. API order is not a contract, so select by
# creation time and use the immutable run ID as a deterministic tie-breaker.
newest=$(printf '%s\n' "$runs" | LC_ALL=C sort -t $'\t' -k1,1 -k2,2n | tail -n 1)
status=$(printf '%s' "$newest" | cut -f3)
conclusion=$(printf '%s' "$newest" | cut -f4)
url=$(printf '%s' "$newest" | cut -f5)

case "$status" in
  completed) ;;
  *) die "${workflow_name} is still ${status} for ${sha} (${url}). Wait for it to finish, then re-run this release from the same tag" ;;
esac

[[ "$conclusion" == "success" ]] ||
  die "${workflow_name} concluded ${conclusion} for ${sha} (${url}). Fix it on main, let ${workflow_name} go green on the exact commit you intend to ship, then move the tag and release again"

printf 'ci gate: %s succeeded for %s\n' "$workflow_name" "$sha"
