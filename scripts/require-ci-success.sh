#!/usr/bin/env bash
# CI gate — refuses to release a commit whose CI is not green.
#
# Release and CI are independent workflows over the same commit. Without this
# check the release path never learns that CI was red on the exact bytes it is
# about to publish, and a tag would be created over a known-broken tree.
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
workflow=${CI_WORKFLOW_NAME:-ci}

if [[ -z "$repo" ]]; then
  repo=$(git remote get-url origin 2>/dev/null |
    sed -E 's#^.*github\.com[:/]##; s#\.git$##') ||
    die 'could not resolve the repository; set GITHUB_REPOSITORY'
fi
[[ -n "$repo" ]] || die 'could not resolve the repository; set GITHUB_REPOSITORY'

command -v gh >/dev/null || die 'the GitHub CLI (gh) is required'
command -v jq >/dev/null || die 'jq is required'

runs=$(gh api --paginate "repos/${repo}/actions/runs?head_sha=${sha}&status=completed" \
  --jq ".workflow_runs[] | select(.name == \"${workflow}\")" |
  jq -s '.') || die "could not query workflow runs for ${sha}"

[[ "$(jq 'length' <<<"$runs")" -gt 0 ]] ||
  die "no completed '${workflow}' run found for ${sha}; wait for CI before tagging"

conclusion=$(jq -r 'sort_by(.run_started_at) | last | .conclusion' <<<"$runs")
[[ "$conclusion" == success ]] ||
  die "the latest '${workflow}' run for ${sha} concluded '${conclusion}', not success"

printf "ci gate: '%s' is green on %s\n" "$workflow" "$sha"
