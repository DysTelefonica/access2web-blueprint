type: patch

El test de issue-spec localiza su gate con la variable CI_PATTERN_ISSUE_SPEC_PATH (override) y, sin ella, con el layout del asset (parents[1]). Un consumer que vendoriza el gate fuera del layout, como APAP_WEB en scripts/, adopta el test byte a byte apuntando la variable a su copia. Tests nuevos fijan ambos caminos.
