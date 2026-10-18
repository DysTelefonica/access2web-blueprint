type: minor

HR-68 (#805, 1/9): `architecture.stack` admite `none` para consumers sin linter de imports (p. ej. PHP). El árbol de módulos y las capas del documento se miden igual, en los dos sentidos; `linter_config` deja de ser obligatorio con `none` y declararlo es una contradicción que falla. Sin `none`, `stack` y `linter_config` siguen siendo obligatorios. Antes CADETE tenía que invocar `check_docs_architecture.py` a mano fuera del contrato.
