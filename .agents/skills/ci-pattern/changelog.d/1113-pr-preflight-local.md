type: minor

Añade el modo API-free `ci-pattern pr preflight --draft --pr-body <archivo> --base <rama> --label <etiqueta>` (#1113). Exige `--base`, admite `--label` repetible y reusa en local `check_issue_spec.validate_body` y el gate `pr-size`; cuerpo y tamaño deben pasar. No llama a la API, a `gh` ni a la red aunque exista `GITHUB_TOKEN`.
