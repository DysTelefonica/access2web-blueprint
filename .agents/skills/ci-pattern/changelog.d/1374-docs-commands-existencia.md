type: minor

HR-65 (#1374): `docs-commands` añade una capa de existencia. Los tres niveles de ejecución prueban la **forma** del comando documentado —que el parser real lo acepte—, y con `dry_run_flag --help` eso es todo lo que se medía: el caso vivo de `DysTelefonica/access2web-blueprint` (`docs/adopcion-fase-2-aislamiento.md:41`) cita `make test` con `--cov` y da el `Makefile` como fuente, pero el target `test` dejó de medir cobertura y el gate no lo veía.

La capa nueva comprueba que **lo que el comando invoca existe**, sin ejecutar nada: el script local que cita (`./scripts/x.sh`) tiene que estar en el repositorio; el target de `make` que cita tiene que existir en el `Makefile` declarado, y los flags que la documentación le atribuye en la misma línea tienen que estar en su receta; el ejecutable del `PATH` que `path_executables` declara permitido tiene que resolverse. La regla es estrecha y determinista: solo mira ficheros, Makefile y `PATH`, y no depende de que la CLI esté en el runner. La activa la declaración `existence` del bloque `docs_commands`, así que ningún consumer cambia de comportamiento hasta que la declare, y `exclude` exime sus citas igual que exime un comando de ejecución.

Tests: `assets/docs-commands/tests/test_docs_commands.py` (`ExistenceTests`).
