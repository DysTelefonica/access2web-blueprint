type: minor

HR-39 (#1196): la comprobación de comando de `state_dependent_gates` sigue un `uses: ./...` del `pr_job` hacia el workflow reutilizable local que llama, y busca el comando en sus jobs. Si ese workflow no está entre los `--workflow` de entrada, o no ejecuta el comando, el hallazgo lo nombra. Un `uses:` externo sigue dando el hallazgo, sin cambio de comportamiento. Motivo: el job `pip-audit` de Access2Web Blueprint llama a `pip-audit.yml`, y con #1192 era el único hallazgo que quedaba (BP#1414).
