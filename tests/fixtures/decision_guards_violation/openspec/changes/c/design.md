# Malformed table fixture — content for malformed_table / malformed_row.
# Today's gate silently skips this (DG-3 fail-closed via no_decision_table).
# The fixture content stays in place so the slice that adds those verdict
# keys finds a ready trigger.

| ID | Decision |
| DG-101 | Missing delimiter row — would emit malformed_table if the gate emitted it. |