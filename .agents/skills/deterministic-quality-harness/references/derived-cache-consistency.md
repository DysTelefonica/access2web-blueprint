# Derived cache consistency check

A derived (denormalized) cache is correct only while **every** write path recomputes it. Reviews
find the missing paths one at a time; the next forgotten path goes stale silently until a person
notices a wrong listing. Hard Rule 20 replaces luck with a detector.

## Write paths that typically forget to recompute

| Path | How the cache goes stale |
|---|---|
| Import ordering | The importer writes the cache before the source rows it derives from, or derives from a partially loaded set. |
| Half-failed import | The source rows commit, the recompute step fails or is skipped, and the run reports success. |
| Interactive reconcile | A manual `reconcile --interactive --apply` edits source rows through a path that bypasses the recompute hook. |
| Web creation paths | A new create endpoint (entity, intake, child record) inserts the source row and never calls the recompute. |
| Backfill gaps | The cache table exists but was never populated, so every row is `missing`. |

Every new write path is another candidate. The check does not trust the list; it compares outcomes.

## The check

- **Read-only.** Open a read-only session or transaction and refuse to run when the server says it
  is writable, or when the connected role can write the cache or its source tables. A safety net
  that can repair the cache hides the bug it exists to report. Repair goes through a separate,
  explicit recompute command.
- **One snapshot.** Read keys, cache, and source in one consistent snapshot (for example
  `REPEATABLE READ READ ONLY`); a row changing mid-run must not read as drift.
- **Every live row.** Iterate the live keys of the source, not the cache: a live row with no cache
  row is drift (`missing`). Sampling misses the one stale row.
- **Same derivation.** Call the function the application uses to derive the value on the fly. A
  second, re-implemented derivation drifts on its own.
- **Exit codes.** `0` no drift, `1` drift, `2` usage, connection, or refused session; an unexpected
  failure gets its own code so a broken run is never read as a clean one (Hard Rule 18).
- **No secrets in output.** Print the destination without user or password; take the DSN from the
  environment, not the command line.

`assets/scripts/cache_consistency.py` ships the pure comparison (`compare`, `exit_code`,
`summary`); the consuming repo supplies the three reads:

```python
from cache_consistency import compare, exit_code, summary

with read_only_snapshot(dsn) as db:          # refuses a writable session or role
    live = db.live_keys()                    # SELECT id FROM entity WHERE active ORDER BY id
    cached = db.cached_values()              # SELECT entity_id, value FROM entity_cache
    drift = compare(live, cached, lambda key: derive_value(db, key))
print(summary(drift))
sys.exit(exit_code(drift))
```

## The single issue

Run the check from a CLI and from a scheduled job (nightly). On drift, the job opens **one**
issue, or comments on it while it stays open; it never opens a second one (pattern of the
APAP_WEB#3502 safety net):

- Find the open issue by label plus exact title through the issues list endpoint, not search: the
  search index lags, and two red runs in a row would open two issues.
- The body or comment carries the count, a bounded sample of keys, and the run link.
- No drift: do nothing. A cancelled run is not drift.
- It is a trace, not a merge gate: nothing required waits for it, and the job declares only the
  permissions it uses (`contents: read`, `issues: write`).

## Required test

The check goes red on a deliberately stale row: seed a consistent dataset, assert exit `0`; corrupt
exactly one cached value, assert exit `1` and that the report names exactly that key. Add the
`missing` case. `assets/tests/test_cache_consistency.py` pins the pure part; the consuming repo
repeats it against an ephemeral database for the read path.

Origin: APAP_WEB#3557 (`animal_current_state`, five stale write paths found by hand in one night),
team-skills#1556.
