type: patch

`_changelog_fragment_gate_wired` ahora recorre `.github/workflows/*.{yml,yaml}` (antes sólo `*.yml`). GitHub trata ambas extensiones como workflows del mismo árbol; si el consumer cablea el gate `check_changelog_fragment.py` en un workflow `*.yaml`, la línea HR-44 del doc operativo vuelve a emitirse (#883, Refs #875).
