# 960-adopt-issue-spec-canon

- Repository: DysTelefonica/access2web-blueprint · Tracking issue: #960 (adoptar el asset canónico issue-spec de ci-pattern — DysTelefonica/team-skills#633)
- Locator: `odd/tasks/960-adopt-issue-spec-canon.md` · Engram mirror topic `odd/960-adopt-issue-spec-canon/tasks`
- Skill under change: `personal/ardelperal/ci-pattern/` (en DysTelefonica/team-skills; el consumer no toca la skill)

## Specs

S1. `scripts/check_issue_spec.py` es byte a byte la copia del canon en `team-skills/personal/ardelperal/ci-pattern/assets/issue-spec/check_issue_spec.py` (sha256 idéntico al asset en origin/main).
S2. `tests/test_check_issue_spec.py` es byte a byte la copia del canon en `team-skills/personal/ardelperal/ci-pattern/assets/issue-spec/tests/test_check_issue_spec.py` (sha256 idéntico al asset en origin/main).
S3. `scripts/local-preflight.sh` invoca el modo `local --body-file` del canon, leyendo el cuerpo desde un fixture versionado; la invocación es determinista y sale con `exit 0` cuando el cuerpo lleva `Closes|Fixes|Resolves|Refs #N` fuera de código.
S4. `scripts/check_issue_spec.py` sigue cubriendo las mismas reglas que aplicaba antes en BP: cuerpo lleva `Closes|Fixes|Resolves|Refs #N`; la referencia resuelve a una issue (no a una PR); la issue lleva `status:approved` y un `type:*`; la API falla cerrada.
S5. El workflow `ci.yml:140` sigue invocando `python scripts/check_issue_spec.py pr-event "$GITHUB_EVENT_PATH"` —sobre la copia canónica—. El job `pr-gate` sigue verde.
S6. HR-54 del canon está cubierto en BP: `Refs`/Closes solo apuntan a issues del mismo repositorio; la PR de la que se depende va en `depends-on` de Chain Context; las referencias dentro de spans de código o bloques no cuentan como closing keywords (gate: `scripts/check_issue_spec.py`).
S7. `.github/ci-pattern-adoption.json` fase 0 lista `scripts/check_issue_spec.py` como artefacto adoptado apuntando al canon, con la descripción mencionando `team-skills #633`.
S8. `.governance-manifest.json` queda igual — el gate canónico y su test no figuran en él (los `scripts/check_*.py` y los `tests/` no se trackean en este manifiesto).
S9. PRs ≤ 400 líneas; si no cabe, cadena con `chain:partial` + `Refs #960` en las intermedias y `Closes #960` en la punta. Nunca `size:exception`.
S10. Commits convencionales en inglés, sin atribución de IA. Sin rebase ni force-push. Una PR en CI a la vez.

## Tasks

- T1 (S1, S4, S6, S10) PR-1, chain 1/3, `chain:partial` + `Refs #960`: sustitución byte a byte de `scripts/check_issue_spec.py` por la copia canónica. El feature doc viaja dentro de este commit.
- T2 (S2, S10) PR-2, chain 2/3, `chain:partial` + `Refs #960`: sustitución byte a byte de `tests/test_check_issue_spec.py` por la suite canónica. `depends-on: PR-1`.
- T3 (S3, S5, S7, S10) PR-3, chain 3/3, `Closes #960`: añadir el modo `local --body-file` al `scripts/local-preflight.sh` (con fixture versionado); actualizar `.github/ci-pattern-adoption.json` fase 0; `make verify` y `adoption check --phase 6` en verde. `depends-on: PR-2`.
- T4 (verificación final) tras los merges: `python3 personal/ardelperal/ci-pattern/assets/bin/ci-pattern adoption check --phase 6 .` desde un clon actualizado de team-skills (origin/main) sobre el árbol real de BP. Exit 0.

Estado: pendiente de arrancar T1.

## Log

L1. Petición original (verbatim): «Tarea nueva: adoptar el asset canónico issue-spec de ci-pattern (team-skills #633, ya en main: personal/ardelperal/ci-pattern/assets/issue-spec/) en DysTelefonica/access2web-blueprint y en DysTelefonica/cadete. APAP_WEB no entra: es de apap3. Por cada repo, una issue nueva con las secciones de su contrato y PRs de 400 líneas como máximo: BP: sustituye scripts/check_issue_spec.py por la copia canónica (entera, byte a byte, sin retoques) y su workflow por la invocación del asset. Añade el modo local --body-file a preflight/local-ci-parity, si el repo lo tiene. Las reglas que ya aplicaba (seis secciones, status:approved, solo issues, HR-54) tienen que seguir aplicándose; si el canon no cubre alguna, para y dímelo con «?» (es un fallo de team-skills), sin parchear el consumidor. CADETE: no tiene gate de issue-spec. Adóptalo igual, con el contrato de issues que ya use CADETE. Actualiza el inventario de la fase 0 y el manifiesto de gobierno de cada repo, y comprueba en un clon completo (no shallow) que «ci-pattern adoption check --phase 6» sigue en verde en los dos. Usa la plantilla de PR canónica con la regla HR-54 (#642).»

L2. Issues creadas: BP #960 (https://github.com/DysTelefonica/access2web-blueprint/issues/960), CADETE #366 (https://github.com/DysTelefonica/cadete/issues/366). Ambas con `type:chore` y `status:approved`.

L3. Estimación de tamaños (diff stat): `scripts/check_issue_spec.py` 167 inserciones + 82 borrados = 249 líneas; `tests/test_check_issue_spec.py` 167 inserciones + 71 borrados = 238 líneas; sustitución combinada 487 líneas, excede 400. Por eso se encadena en tres PRs.