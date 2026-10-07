# Migrating a repository onto the pool

Ordered, because the failures are ordered. Everything here was observed moving
`ardelperal/APAP_WEB` from GitHub-hosted runners to this pool on 2026-08-11.

## The decision rule — visibility decides, arithmetic confirms

**Public repo → hosted minutes are free → do NOT migrate the PR-reachable
jobs.** The cost (untrusted PR code reaching the production `docker.sock` as
root, next to Coolify/postgres/adguard/vaultwarden) is not justified against
$0. Evidence: `gh api repos/<o>/<n> --jq '.private'` → `false`, and the
billing report `/orgs/<org>/settings/billing/usage` showing `netAmount: 0` on
every Actions line.

**Private repo → hosted minutes bill (2,000 min/month on an org free plan; 0
if a personal account pays, whose included minutes differ — verify who owns
the billing) → MIGRATING IS THE STANDARD PATH.** Reference arithmetic,
measured on APAP_WEB (2026-10: 23 hosted jobs, 403.5 job-min over 12.3 h ≈
23,500 min/month → a free-plan quota runs out in ~2.6 days). The migration
trigger is the visibility, not preference.

**The sequence is mandatory**: (1) decide visibility, (2) migrate. Never the
reverse (HR-14: PR jobs on self-hosted with a public repo = critical risk).

## The private-repo recipe — seven steps, each with its evidence

Once the repo is private, this is the order. Every step carries the command
that proves it; a claim without its command is an opinion.

### 1. Inventory

Census every `runs-on` per repo and workflow, count the jobs, and take the
widest concurrency tier — that tier is the pool pressure you must absorb.
Then read the actually-billed minutes, classifying each job by
`runner_group_name` (hosted vs SELF-HOSTED) with its duration; Windows jobs
weigh ×2 against the quota:

```sh
gh api repos/<O>/<R>/actions/runs/<RUN_ID>/jobs \
  -q '.jobs[] | "\(.runner_group_name)\t\(.name)\t\(.started_at)\t\(.completed_at)"'
```

### 2. Per-job restrictions

- **Architecture**: this fleet is **aarch64** → verify the index of every
  pinned image (`docker buildx imagetools inspect <image>`) — an amd64-only
  image does not even start (real case: APAP_WEB's replicated minio).
- **Windows never migrates to this VPS** (real case: dysflow — its 2
  `windows-latest` jobs are 87% of its billing; they stay hosted, full stop).
- **`services:` with ports**: on self-hosted the runner container sits on its
  own bridge, so `127.0.0.1` points at the runner, not the host — use
  `host.docker.internal` and TEST reachability instead of assuming it.

### 3. Sizing

A PR needs as many runners as the graph's widest tier, compared against the
runners *usable* for it (the ones matching the job's labels). Headroom 0
means one PR saturates the fleet and the second one queues. Measure `nproc`
and the host memory shared with the rest of the containers — the method is
`pool-sizing.md`.

### 4. Repo gate

If the repo carries a gate like
`scripts/check_workflows.py::check_runner_isolation` (APAP_WEB #782),
PR-reachable jobs do NOT move without an explicit operator decision: that
gate is what protects the production socket. The options are (a) leave the
PR jobs hosted, (b) extend/allowlist the gate accepting the risk, or (c)
split the workflow so the migrated jobs live on `push`/`schedule` triggers.

### 5. Migration order

First the PR-exempt jobs (`schedule` / `workflow_dispatch` / tags — zero
risk), then the `push` ones, and only at the end the PR jobs, after the gate
decision of step 4. APAP_WEB's `deploy` workflow is the example already
proven on this fleet.

### 6. Pool hygiene before widening

- Uniform labels — no single-runner SPOF behind a label (real case: `deploy`
  was 1-wide and absent from the `RUNNER_LABELS` env → one restart deleted it).
- No orphan runners that match no job.
- Healthchecks that cannot report false `unhealthy` (real case:
  `curl https://api.github.com/zen` returns 403 under rate limit and leaves
  the runner permanently `unhealthy`).
- `timeout-minutes` re-sized for aarch64 (+~20% over the hosted timings).
- No fossil systemd units left from pre-container provisioning.

### 7. First guarded run

Equivalence is proven with ONE run observed by the operator on a runner
different from the historical one. The run history records which
`runner_name` each job ran on:

```sh
gh api repos/<O>/<R>/actions/runs/<RUN_ID>/jobs -q '.jobs[] | "\(.name)\t\(.runner_name)"'
```

A job that has never run outside its historical runner has UNPROVEN
equivalence, even if the containers are identical.

## Before you move anything

**Check whether the move is even warranted.** Standard-runner minutes are free
for public repositories and billed for private ones (verified 2026-10-02:
`gh api repos/ardelperal/APAP_WEB --jq '{private,visibility}'` → public; every
Actions minute line in the org usage report shows `netAmount: 0`). A private
repo that exhausts its quota has a real reason; a public one is paying
complexity for nothing. Also check who owns the billing: a repo in an
organisation and a repo under a personal account draw from different quotas, so
a routing decision that is correct in one is not portable to the other. The
working billing endpoint is `/orgs/<org>/settings/billing/usage`
(`/settings/billing/actions` returns 410/403 — dead end).

**Sequence the visibility decision FIRST — before the migration, not after.**
Moving `pull_request`-reachable jobs onto a runner that mounts the production
`/var/run/docker.sock` is a critical risk while the repo is PUBLIC: any
outsider's PR would get root on the prod host, alongside
Coolify/postgres/adguard/vaultwarden. Once the repo is private, the same setup
is only collaborator-level risk. APAP_WEB is currently public, so this is the
live constraint. Repo-side enforcement: `scripts/check_workflows.py::
check_runner_isolation` (APAP_WEB issue #782). Decide visibility, then migrate
— never the reverse.

**Measure the wall clock first.** Write down each job's duration. You will be
asked whether a second runner is worth it, and the answer usually is not: on
APAP_WEB the whole critical path was dominated by a single 123-second test job,
so a second runner bought ~2 minutes. A pool grows for *redundancy*, not speed.

## The move itself

1. Change every `runs-on` to the pool's exact labels. Never `-latest`.
2. **Add `actions/setup-python` to every job that runs a `python` command.** The
   hosted image had an interpreter; this one does not. The jobs that break are
   the ones that never needed the step before, so they are also the ones nobody
   thinks to check.
3. **Grep every `run:` block for `gh`.** Rewrite as curl + jq. See
   `assets/repo-ci-hardening.yml` block 6.
4. Apply blocks 1–5 of that file. Blocks 1 and 4 are prerequisites for ever
   adding a second runner.
5. Update the comments you invalidated. A `runs-on` rationale that still argues
   for the runner you just left is worse than no comment: the next reader trusts
   it.

## After the first merge — do not skip this

**Green PR checks do not mean the merge succeeded.** A `pull_request` run and a
`push`-triggered workflow are different runs with different contexts, and the
second one is where a migration breaks.

On APAP_WEB, four consecutive merges landed with every PR check green while the
`deploy` workflow failed each time. Nobody noticed for two days, because nobody
looks at a workflow that runs *after* the thing they were watching. The cause was
block 6: `gh` was absent, and the error was laundered into a legitimate-looking
"this tree was never proven".

So after the first merge, read the post-merge run list, not the PR:

```sh
gh run list --workflow deploy --limit 5 \
  --json conclusion,headSha -q '.[] | "\(.conclusion) \(.headSha[0:8])"'
```

Compare against the last known-good deploy. If the boundary lines up with your
migration commit, the migration caused it.

## Writing the gates that keep it true

Applying the hygiene by hand is half the work; nothing stops the next job from
omitting it. Every block above should ship with a check, and these three
properties are what separate a check from decoration:

**It must be provable that it can fail.** Reintroduce the defect on the real
tree — not only in a synthetic fixture — and watch the gate reject it. A gate
that has never failed is indistinguishable from one that cannot.

**It must ignore comment lines.** The file usually documents the very command it
stopped using, so a scanner that reads comments reports the explanation as the
offence. That teaches people to delete explanations, which is a worse outcome
than the defect. This bites the tests too: an `assert "|| echo 0" not in text`
fails on the comment describing its removal.

**Word boundaries matter.** A bare search for `gh` matches `through`, `high` and
`enough`. Anchor on invocation position.

## Ordering rule when several checks share a file

Run the *parse-integrity* check first, with its own scanner, and stop if it
fails. A YAML library will happily accept duplicate keys and keep the last one,
so any conclusion drawn past an ambiguous key is arbitrary — the duplicate is
the finding, and the rest must wait.
