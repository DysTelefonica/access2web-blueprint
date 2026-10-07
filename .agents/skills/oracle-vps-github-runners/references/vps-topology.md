# Oracle VPS topology

Facts verified 2026-08-11, corrected 2026-10-02 (visibility, runner census,
Coolify version, access path, billing). Re-verify before citing; the host
changes.

## Host

| | |
|---|---|
| SSH | `ssh oracle` (alias in `~/.ssh/config`, key `~/.ssh/oracle.key`) |
| Hostname | `vinc-ardelperal` |
| Arch / kernel | `aarch64`, Linux 6.17 `-oracle` |
| Capacity | 4 vCPU, 23 GiB RAM |
| Disk | `/dev/sda1` 194 G, was 67 % used |
| Runtimes | Docker 29.6.2 **and** podman 4.9.3 |

From Git Bash on the Windows workstation the alias needs explicit flags,
because that shell resolves a different `HOME`:

```sh
ssh -F /c/Users/adm1/.ssh/config \
    -o UserKnownHostsFile=/c/Users/adm1/.ssh/known_hosts \
    oracle '<command>'
```

## Access path

There is **no `coolify` CLI** on the host. The only supported path is the
Coolify **REST API** at `panel.romancaba.com`, authenticated with
`COOLIFY_ACCESS_TOKEN` from `~/.config/opencode/secrets.env` (verified
2026-10-02: `GET /api/v1/services` → 200, 11 services), typically driven via
the MCP entry in `~/.config/opencode/opencode.json`
(`npx -y @masonator/coolify-mcp@latest`).

Coolify here is **4.3.23** (`GET /api/v1/version` → 200, verified 2026-10-02).
The earlier `Not found` behaviour of `service action=list_containers` and
`logs resource=service` over MCP was observed on 4.1.2 (below 4.2) and may no
longer hold — re-verify at runtime before falling back to reading container
state over SSH.

## Runner inventory

All container runners are Coolify **services** on this single host; one runs
as a systemd unit. **Four runners are registered on APAP_WEB** (verified
census 2026-10-02):

| Runner | Type | Labels | Notes |
|---|---|---|---|
| `apap-coolify-noble` | Coolify service | `self-hosted,Linux,ARM64,apap,oracle,coolify,noble,deploy` | the **ONLY** runner carrying `deploy` |
| `apap-coolify-noble-2` | Coolify service | `self-hosted,Linux,ARM64,apap,oracle,coolify,noble` | no `deploy` |
| `apap-coolify-noble-3` | Coolify service | `self-hosted,Linux,ARM64,apap,oracle,coolify,noble` | no `deploy` |
| `apap-web-oracle-arm64` | systemd unit | `self-hosted,Linux,ARM64,apap,oracle` | no `coolify`/`noble`/`deploy` — matches **ZERO** jobs |

Other services on the same host (unchanged): `a2w-blueprint-runner`
(`myoung34/github-runner:ubuntu-noble`) for `DysTelefonica/access2web-blueprint`
and `planhub-runner` (`myoung34/github-runner:latest`). Both `myoung34` images
read the PAT from `ACCESS_TOKEN` natively; `ghcr.io/actions/actions-runner`
needs the mint-at-start command in `assets/runner-compose.yml`. Either way the
rule is the same: no static `RUNNER_TOKEN`.

**Known defect — `deploy` SPOF:** `deploy` is a 1-wide lane. Every deploy job
serialises through `apap-coolify-noble`, and the lane dies with that single
container. Widening the lane means adding the `deploy` label to more runners,
not adding runners without it.

`apap-web-oracle-arm64` reads as redundancy on the runners page while matching
no job. Treat a runner that cannot match any job as absent.

**Healthcheck noise (cosmetic but misleading):** the three container runners
have shown `unhealthy` for 4 weeks (`FailingStreak 18`): the healthcheck's
`curl https://api.github.com/zen` half fails from the container netns while
the listener half passes and jobs execute normally. The steady state being
`unhealthy` hides real failures — judge runners by job execution, never by the
health badge.

## Workspace paths

One directory per runner under `/data/runners/`. Sharing one between two
instances makes two branches check out over each other.

## Billing context (verified 2026-10-02)

`ardelperal/APAP_WEB` is **PUBLIC** — the earlier claim that it is private is
FALSE and drove a misdirected migration investigation. Evidence:

- `gh api repos/ardelperal/APAP_WEB --jq '{private,visibility}'` →
  `{"private":false,"visibility":"public"}`
- an anonymous `api.github.com` call for the repo returns 200

Standard-runner minutes for public repositories are **FREE**. The org billing
usage report shows every Actions minute line at `netAmount: 0` (8,342 min in
September → $0). The only billed item fleet-wide is ~$0.14/month of artifact
storage (dysflow $0.129 + access2web $0.0072) — self-hosting does not reduce
it.

**Forward-looking migration trigger:** if APAP_WEB ever flips to private, its
~23,500 min/month (403.5 hosted job-min measured in 12.3 h) exhausts the
2,000 min/month free allowance in ~2.6 days — THAT is when this migration
becomes necessary, not before.

**Billing API endpoint:** the working endpoint is
`/orgs/<org>/settings/billing/usage`. The older
`/orgs/<org>/settings/billing/actions` returns 410 (moved), or 403 without
`admin:org` — do not chase it.

`access2web-blueprint` lives in the DysTelefonica org with separate billing, so
a routing decision correct there is not automatically portable to APAP_WEB.
Check repo visibility and the owning account before copying a routing decision
between repos.
