#!/usr/bin/env bash
# Release preflight — the rule that replaces a staging branch.
#
# A tag may only be published when it names exactly what the trunk already
# says. Run it locally before pushing a tag, and again in CI on the tag itself.
#
#     ./scripts/release-preflight.sh v1.4.0-rc.1
#
# Exit 0 means the tag is publishable. Any other exit means it is not, and the
# reason is on stderr. This script never mutates the repository.
#
# Contract: docs/11-releases.md
set -euo pipefail

die() {
  printf 'release preflight: %s\n' "$*" >&2
  exit 1
}

tag=${1:-}
[[ -n "$tag" ]] || die 'usage: release-preflight.sh vMAJOR.MINOR.PATCH[-rc.N]'

# 1. Canonical version identity. Anything else is not a release.
semver='v(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)'
if [[ "$tag" =~ ^${semver}$ ]]; then
  kind=stable
elif [[ "$tag" =~ ^${semver}-rc\.(0|[1-9][0-9]*)$ ]]; then
  kind=rc
else
  die "tag must be vMAJOR.MINOR.PATCH or vMAJOR.MINOR.PATCH-rc.N, got '$tag'"
fi

# 2. The tag object must be annotated, so it carries an author and a date.
#    A lightweight tag records who released nothing at all.
git rev-parse -q --verify "refs/tags/$tag" >/dev/null ||
  die "tag $tag does not exist here; create it with: git tag -a $tag"
[[ "$(git cat-file -t "refs/tags/$tag")" == tag ]] ||
  die "tag $tag is lightweight; releases must be annotated (git tag -a $tag)"

tag_sha=$(git rev-parse "refs/tags/$tag^{commit}")

# 3. Resolve the trunk. Every rule below is measured against it.
git fetch --no-tags --quiet origin '+refs/heads/main:refs/remotes/origin/main' ||
  die 'could not fetch origin/main'
main_sha=$(git rev-parse 'refs/remotes/origin/main^{commit}')

# 4. The identity rule, per channel.
#
#    An RC is a freeze of the trunk, so it must BE the trunk at cut time.
#    A stable tag is either the same thing, or the promotion of an RC that has
#    already been validated — in which case it legitimately sits behind main,
#    because main kept moving while the candidate was under acceptance.
if [[ "$tag_sha" == "$main_sha" ]]; then
  provenance='exact current origin/main'
elif [[ "$kind" == rc ]]; then
  die "an RC must freeze the exact current origin/main ($main_sha), not $tag_sha"
else
  git merge-base --is-ancestor "$tag_sha" "$main_sha" ||
    die "stable tag $tag is not reachable from origin/main"

  promoted=
  while read -r rc_tag; do
    [[ -n "$rc_tag" ]] || continue
    if [[ "$(git rev-parse "refs/tags/${rc_tag}^{commit}")" == "$tag_sha" ]]; then
      promoted=$rc_tag
      break
    fi
  done < <(git tag --list "${tag}-rc.*")

  [[ -n "$promoted" ]] ||
    die "stable tag $tag is behind main but promotes no RC of its own version"
  provenance="promotion of $promoted"
fi

# 5. Nothing uncommitted may leak into the build.
[[ -z "$(git status --porcelain=v1 --untracked-files=all)" ]] ||
  die 'working tree is dirty; commit or stash before releasing'

printf 'release preflight: %s tag %s verified (%s, commit %s)\n' \
  "$kind" "$tag" "$provenance" "$tag_sha"
