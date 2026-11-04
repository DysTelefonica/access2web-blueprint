# Proposal: Legacy coexistence migration to PostgreSQL

## Intent

The Access/VBA ecosystem (Lanzadera, Expedientes, Brass, Condor, No Conformidades, Gestión de Riesgos, HPS, Tareas) must reach PostgreSQL without a single cutover and without breaking the Access applications still in production.

The problem is structural: each legacy front talks to its own backend, and the backends share tables through links. Two decisions already frame the work:

- **Coexistence is per application.** The Access Lanzadera keeps opening the Access applications that are still alive; the web Lanzadera only opens web applications. No cross-launching. Each application is cut and reverted individually.
- **D92**: during coexistence the web works directly on the legacy-shaped tables in PostgreSQL that Access links through ODBC. One source of truth keeps the legacy alive and makes reversal possible. A modern web schema with synchronization is rejected.

This proposal covers the discovery program that must complete before any legacy remediation, and the rules that govern it.

## Scope

### In Scope

- Phase 0 inventory of every front and backend in scope: linked tables, primary keys, validation rules, relations, orphan rows, risky VBA patterns, and classification of each table.
- Owner rules that constrain discovery: no macros or `AutoExec`; forms are disconnected except Brass; a table belongs in the backend of the front whose own VBA needs it; linked tables are read from the dysflow ERD of both front and backend, not from `MSysObjects`; code usage is traced with CodeGraph; an application with a pending RC1 that has not passed UAT is analyzed on that RC1, not on `main`.
- A catalog of remediation proposals per backend, each with its evidence, prior check and reversal.
- The gate that forbids applying any remediation until discovery covers all applications and dependencies are reconciled.

### Out of Scope

- Any change to legacy `.accdb` files, VBA code or production data. Phase 0 is read-only; the owner applies the approved remediation.
- The modern web schema and its migrations; that arrives after the legacy is fully shut down.
- Per-application cutover implementation, ODBC driver selection and PostgreSQL provisioning.
- Deleting or moving legacy tables or data. Those are owner decisions with an explicit protocol.

## Capabilities

### New Capabilities

- `legacy-discovery`: read-only inventory of fronts, backends, linked tables and VBA usage patterns, with a per-table classification and a catalog of remediation proposals.
- `legacy-remediation-gate`: the rule that no legacy change is applied until discovery covers every application and its dependencies are reconciled.

### Modified Capabilities

None. `openspec/specs/` holds no archived capability yet, so no spec-level requirement changes.

## Approach

1. **Base selection.** For each application, pick the code base under study: the pending RC1 branch when UAT has not passed, otherwise the released baseline. Record branch and SHA per application; never mix versions.
2. **Structural read.** Generate the dysflow ERD of the front and of the backend. Linked tables and their origins come from that output; system tables are not queried.
3. **Code read.** Trace usage with CodeGraph: which tables the owning front actually reads, including dynamic SQL and Brass's connected controls. Absence of static call sites is not proof of dead code.
4. **Cross-check.** Reconcile ERD, code usage and the ownership rule: a table stays in a backend only if that backend's own front needs it. Tables needed by another front are candidates for reassignment; tables nobody needs are candidates for retirement, never automatic deletion.
5. **Classify and price.** Assign every table a classification (`a` indispensable before migration, `b` recommended if cheap, `c` deferred to the final refactor) with its measured cost in code references.
6. **Catalog, do not apply.** Publish proposals with evidence, prior check and reversal. Apply nothing until the gate opens.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `docs/07-migracion/epica-coexistencia-postgresql.md` | Modified | Living epic: phases, entry condition for phase 1, change-of-course log. |
| `docs/07-migracion/fase-0/` | New/Modified | Per-backend inventories and the direct-access inventory for `Lanzadera_Datos`. |
| `skills/coexistence-migration/SKILL.md` | Modified | Repository-local skill: owner rules, decision gates, anti-patterns. |
| `docs/decisions/2026-11-04-d92-tablas-con-forma-legacy-compartidas.md` | New | D92, the shared legacy-shaped table strategy. |
| `scripts/dysflow-erd/extract-erd.ps1` | Modified | ERD extraction extended for the new metadata. |
| `openspec/changes/legacy-coexistence-migration/` | New | This proposal and its downstream specs, design and tasks. |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| A table judged dead is still read by a front outside the analyzed repositories, or through a connected control in Brass. | Med | Discovery gate: no deletion without evidence from every front, and the owner applies it after a quarantine window. |
| Remediating one backend early invalidates another front that links its tables. | Med | No remediation before global discovery and dependency reconciliation. |
| `DROP SCHEMA` used as a rollback path destroys legacy-shaped tables still read by Access. | High | DA-7 and D82 are under review (#1559). Additive schema changes only during coexistence; the rollback path never drops a schema with legacy-shaped tables. |
| Development copies drift from production data, so counts and orphans are wrong. | Med | Counts are re-verified against a recent production copy before any action is approved. |
| A development link points to the production data share, risking writes from development. | Med | Detected during the Expedientes extraction; link hygiene is part of discovery. |
| Review overload from a large inventory in one change. | Med | One application per work unit, under the 400-line review budget. |

## Rollback Plan

Phase 0 produces documents only: revert the commit, or discard the inventory document. Legacy data is never touched, so there is nothing to roll back in the legacy estate.

When remediation is later approved, each action carries its own reversal: quarantine by rename to `zzBAJA_<tabla>`, deletion from a validated backup, and relation removal for added referential integrity. The owner applies and reverses in the legacy; the agent only provides evidence.

## Dependencies

- dysflow MCP for ERD generation on fronts and backends.
- CodeGraph (initialized through Gentle AI, not directly) for VBA usage tracing.
- The RC1 branch of each application with pending UAT, and its SHA.
- Production copies for count verification, read-only.

## Success Criteria

- [ ] Every application in scope has a documented base (branch and SHA) and its front and backend ERD.
- [ ] Every backend has its list of linked tables and origins, obtained without querying system tables.
- [ ] Every table has a classification, a measured cost and a documented ownership verdict.
- [ ] Every front has its inventory of risky VBA patterns, including the Brass connected controls.
- [ ] R-1 and R-2 are recorded as confirmed or refuted, with their exceptions listed.
- [ ] A single catalog of remediation proposals exists, with evidence, prior check and reversal per entry.
- [ ] No legacy file, database or data has been modified during the discovery program.
