#!/usr/bin/env bash
# Installs the repository skills into the agent directories present locally.
#
# Skills live once in skills/, versioned with the code. Agents each read from
# their own directory, so this copies rather than symlinks: Windows symlinks
# need developer mode or elevation, and a copy behaves the same everywhere.
#
# Re-run after any `git pull` that touches skills/.
#
# Contract: skills/README.md
set -euo pipefail

repo_root=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")/.." && pwd)
source_dir="$repo_root/skills"

[[ -d "$source_dir" ]] || {
  printf 'install-skills: %s does not exist\n' "$source_dir" >&2
  exit 1
}

# Agent directory -> the agent it belongs to. A target is written only when its
# parent already exists, so nobody gets a directory for an agent they do not use.
targets=(".claude:Claude Code" ".opencode:OpenCode" ".codex:Codex")

# Counted apart on purpose: "no agent installed here" and "no skills to install
# yet" are different situations, and reporting one as the other sends the reader
# to fix something that is not broken.
agents_found=0
skills_copied=0

for entry in "${targets[@]}"; do
  dir=${entry%%:*}
  agent=${entry#*:}
  parent="$repo_root/$dir"

  [[ -d "$parent" ]] || continue
  agents_found=$((agents_found + 1))

  dest="$parent/skills"
  mkdir -p "$dest"

  for skill in "$source_dir"/*/; do
    [[ -d "$skill" ]] || continue
    name=$(basename "$skill")
    rm -rf -- "${dest:?}/$name"
    cp -R -- "$skill" "$dest/$name"
    printf 'install-skills: %s -> %s/skills/%s\n' "$name" "$dir" "$name"
    skills_copied=$((skills_copied + 1))
  done

  printf 'install-skills: %s ready at %s\n' "$agent" "$dest"
done

if [[ "$agents_found" -eq 0 ]]; then
  printf 'install-skills: no agent directory found (.claude, .opencode, .codex).\n'
  printf 'install-skills: create the one your agent uses and run this again.\n'
elif [[ "$skills_copied" -eq 0 ]]; then
  printf 'install-skills: skills/ holds no skill yet; nothing to copy.\n'
fi
