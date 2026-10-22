type: patch

Issue-spec (#840): rule 4 is retired. A PR body that closes and references the same issue (`Closes #N` with `Refs #N`) now passes without a warning, since GitHub closes the issue once and the redundant link is harmless. Rule numbers stay stable.
