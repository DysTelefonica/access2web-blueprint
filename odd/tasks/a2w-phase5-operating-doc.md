# Fase 5 del patrón de CI en access2web-blueprint (#781)

## Objetivo y alcance

Completar la adopción del patrón de CI en `DysTelefonica/access2web-blueprint`
con los tramos 2/3 y 3/3 de la fase 5: parámetros del patrón, documento
operativo generado, bloque canónico en `AGENTS.md`, declaración
`phases."5".operating_doc` en el contrato de adopción, corrección mínima del
`README.md` (HR-48) y job de deriva del documento en el CI del consumidor.
El tramo 1/3 (re-apuntado de citas) ya está integrado en `main`.

Base: `95b80377cffbb4b3fe59e6c7b8c5cc4c964ac6ef` (`origin/main`).
Rama tramo 2/3: `chore/781-fase-5-doc-operativo-generado`.
Rama tramo 3/3: `chore/781-fase-5-deriva-doc` (base = rama del 2/3).
Pin objetivo del generador y del verificador:
`a1cca27194335424c44bb5d9595fcd6c5d7fb3bd` de `DysTelefonica/team-skills`,
que ya incluye #391: no instala la semántica obsoleta de rebase ni de limpieza
de rama.

## Límites

Sin merge, sin cierre de issue, sin etiquetas protegidas, sin borrado de ramas
ni worktrees, y sin reescritura (`rebase`, `--amend`, `force`) sobre ninguna
rama publicada. Sin propagación y sin `fleet registry`: no se declara
`agents_md: "propagated"`. No se tocan worktrees ajenos ni las fuentes
integradas de #390/#391. Runners hospedados (`ubuntu-24.04`) para este
repositorio público; nunca self-hosted. Presupuesto de 400 líneas por PR; la
cadena usa `chain:partial` + `Refs #781` en el intermedio y `Closes #781` solo
en la punta.

## Tareas

- [x] T1 — Espacio aislado: worktree propio desde `origin/main` fresco
  (`95b80377`), sin tocar los worktrees ajenos; pin verificado y parser
  `P09_labels_type: {}` comprobado contra el pin.
- [x] T2 — Corrección mínima de `README.md`: retirar la ruta absoluta del host
  sin prohibir universalmente la inspección del legacy (HR-48).
- [x] T3 — `ci-pattern.yaml` con los parámetros ratificados y
  `P09_labels_type: {}`; `params validate` en verde.
- [x] T4 — `docs/ci-pattern-flow.md` generado con el CLI del pin `a1cca27`.
- [x] T5 — `AGENTS.md`: enlace al documento y bloque canónico byte a byte
  (read-back idéntico observado).
- [x] T6 — Contrato: declaración `phases."5".operating_doc` verificada por el
  gate. El inventario de `phases."0"` NO admite artefactos fuera del catálogo
  canónico del patrón: añadido descartado (ver Decisiones observadas).
- [ ] T7 — Verificación del tramo 2/3: `adoption check --phase 5` en verde
  (rc=0) y fases previas sin regresión; falta el read-back de determinismo del
  documento y los checks del repositorio afectados por los ficheros tocados.
- [ ] T8 — PR del tramo 2/3 (`chain:partial`, `Refs #781`, base `main`).
- [x] T9 — Tramo 3/3: `.github/workflows/operating-doc-drift.yml` pineado al
  mismo SHA, hosted, y fallo cerrado cuando el secreto falta o está vacío.
  Verificado por `scripts/check_workflows.py` (6 workflow limpios) y por las
  fases 4 y 5 en verde con el workflow presente.
- [x] T10 — Prueba de comportamiento del control de deriva: el control es la
  detección de deriva del patrón, que ya tiene fixture violador en el catálogo
  (`DocsPhaseTests`, `drift_operating_doc`, con casos de deriva de etiqueta, de
  check requerido y de regex de rama); el workflow es su cableado, que HR-32
  excluye como prueba del control. No requiere superficie nueva.
- [x] T11 — Verificación del tramo 3/3 y PR punta con `Closes #781`.
  PRs abiertos: **#826** (2/3, base `main`, `chain:partial`) y **#827**
  (3/3, punta, base = rama del 2/3); issue #781 abierta.

## Decisiones observadas

- **Inventario de fase 0 descartado.** Añadí `ci-pattern.yaml` y
  `docs/ci-pattern-flow.md` como `adopted` en `phases."0".inventory` y el gate
  los rechazó: el inventario solo admite artefactos del catálogo canónico del
  patrón (`.github/**`, plantillas, `AGENTS.md`, `CONTRIBUTING.md`,
  `.team-skills.yaml`, `ruleset:`); la gobernanza de scripts va por referencia
  o por `declared_governed_paths`. Añadido revertido; la declaración de fase 5
  es suficiente y los artefactos nuevos quedan sin entrada en el inventario.
- **HR-55 (`agents_md` no declarado) es informativo.** El gate imprime la
  advertencia y sale en verde; no se declara `agents_md: "propagated"`, que
  queda fuera del encargo (depende de la propagación del gobierno).
- **P18, P21 y P45.** El esquema las exige sin default seguro; el generador no
  las consume. Se declaran con valor de no aplicación (P18 `NO_APLICA` porque
  su patrón exige `^[A-Z][A-Z0-9_]*$`) en vez de omitirlas.
- **P09 `{}`.** Ratificado y declarado: el default del esquema mapea
  formularios (`bug_report.yml`, `feature_request.yml`…) que este repositorio
  no tiene.
- **El workflow SÍ va al inventario.** El gate de exhaustividad de la fase 0
  exige entrada para todo artefacto canónico presente en el árbol: rechazó
  `.github/workflows/operating-doc-drift.yml` por falta de entrada. Es la cara
  opuesta del caso anterior: `.github/workflows/*.yml` es canónico y
  `ci-pattern.yaml`/`docs/*.md` no. La entrada viaja en el tramo que añade el
  artefacto.
- **Incidente registrado: `--out` relativo.** `ci-pattern adoption
  generate-doc` resuelve un `--out` relativo contra el directorio de trabajo,
  no contra `<repo>`: la primera corrida (cwd = clon del catálogo) escribió
  `docs/ci-pattern-flow.md` **dentro del clon compartido de `team-skills`**,
  sobrescribiendo un fichero untracked previo de 190 líneas que era el mismo
  documento generado por una corrida anterior con un pin anterior. Ninguna
  ruta versionada quedó afectada. Mitigación: `--out` absoluto en todas las
  corridas siguientes. Queda a decisión del orquestador el destino de ese
  fichero untracked; este escritor no lo vuelve a tocar.

## Criterios

- El documento generado es reproducible byte a byte con el mismo pin: mismo
  contenido, sin deriva, sin edición manual.
- El bloque canónico de `AGENTS.md` coincide byte a byte con el emitido por el
  CLI entre sus marcas.
- `adoption check --phase 5` en 0 y las fases 0 a 4 sin regresión.
- El job de deriva no publica nombres de check requeridos (solo `push` y
  `schedule`), usa runners hospedados, pinea las acciones por SHA y falla en
  voz alta si el secreto no está disponible, sin imprimirlo.

## Evidencia y estado

Evidencia observada con el pin `a1cca271` (`HEAD` del clon del catálogo
verificado con `git rev-parse HEAD` antes de cada corrida):

- `params validate` → `OK`, rc=0.
- `adoption check --phase 5` → «fases 0..5 en verde», **rc=0**; la línea
  HR-55 sobre `agents_md` es informativa y no cambia el resultado.
- `adoption check --all` → rc=1 únicamente por la fase 6 no declarada, fuera
  de esta unidad.
- Read-back del bloque canónico en `AGENTS.md`: idéntico byte a byte al del
  documento generado; el enlace a `docs/ci-pattern-flow.md` está presente.
- Tramo 3/3: `scripts/check_workflows.py` → 6 workflow limpios, rc=0;
  `adoption check --phase 4` → fases 0..4 en verde, rc=0 (ejecuta los cinco
  gates del contrato, con paridad local/CI y política de workflows entre
  ellos); `adoption check --phase 5` → fases 0..5 en verde, rc=0.
- CI del PR #826: `required`, `quality` (1m30s), `codeql`, `review-budget`,
  `gitleaks`, `pip-audit` y `trivy-config` en verde; `mutation` omitido por
  evento, como está diseñado. CI del PR #827: pendiente al cerrar este
  registro, sin declarar verde.
- `closingIssuesReferences` del PR punta: vacío mientras su base sea la rama
  del 2/3 (GitHub solo registra el cierre contra la rama por defecto); se
  registra al reapuntar la base a `main` tras el merge del 2/3.
- Reconocimiento previo: `main` de a2w en `95b80377`; cadena #785 sin mergear
  (#799 base `main`, #816 y #825 apilados), sin bloquear esta unidad; el
  repositorio es público; el secreto `TEAM_SKILLS_READ_TOKEN` figura en la
  lista de secretos.

Ninguna unidad se declara completada sin evidencia observada.

## Cierre

- PRs: pendientes de creación.
- Commits: pendientes.
- Observación: el documento de coordinación del orquestador vive fuera de las
  superficies autorizadas para este escritor; las transiciones se registran
  aquí y en el espejo, y se informan al orquestador.
