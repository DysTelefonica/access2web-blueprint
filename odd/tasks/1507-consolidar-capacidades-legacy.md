# 1507-consolidar-capacidades-legacy

- Repository: DysTelefonica/access2web-blueprint · Tracking issue: #1507 (docs(cross-cutting): consolidar el índice de capacidades de las ocho aplicaciones legacy y redactar la matriz legacy→web)
- Locator: `odd/tasks/1507-consolidar-capacidades-legacy.md` · Engram mirror topic `odd/1507-consolidar-capacidades-legacy/tasks`
- Método heredado de: `docs/07-migracion/roadmap-descubrimiento-prd.md` (fases 0 a 5) y de los PRD #1073 a #1080

## Specs

S1. «`docs/05-capacidades/capabilities-index.md` enumera las capacidades de las aplicaciones con fila por capacidad atómica: identificador, aplicación, descripción, tier, estado frente al snapshot, fuente, dependencia principal y excepción.» (AC1)
S2. «El tier se asigna con una regla escrita en el propio índice y aplicada de forma uniforme; el índice declara los tres valores (`critical`, `standard`, `minimal`) y a qué corresponde cada uno.» (AC2)
S3. «El índice declara, aplicación por aplicación, si su fase 5 está cerrada y enlaza la issue cuando no lo está. Ningún estudio se da por terminado sin fase 5 en `main`.» (AC3)
S4. «`docs/07-migracion/matriz-legacy-a-web.md` deja de ser un placeholder: una fila por capacidad con componente web candidato, riesgo principal, dependencia, excepción y confianza, más una sección que marque las capacidades cuya lógica vive en un evento de formulario y debe extraerse a servicio o caso de uso.» (AC4)
S5. «Cada fila de la matriz cita la capacidad del índice y la aplicación que la sostiene.» (AC5)
S6. «Los dos documentos declaran sus huecos abiertos: ERD de origen pendiente en las ocho aplicaciones y fases 5 en curso en HPS, HPS Solicitudes y No Conformidades.» (AC6)
S7. «La entrega se encadena en unidades de revisión de hasta 400 líneas; los intermedios llevan `chain:partial` y `Refs #<n>` y solo la punta lleva `Closes #<n>`. Los commits son convencionales, en inglés y sin atribución de IA.» (AC7)

## Tasks

- T1 (S1, S2, S3, S6) docs(capacidades): reescribir el índice maestro con la regla de tier y las capacidades de las aplicaciones con fase 5 en `main` (Expedientes, Lanzadera, Condor, Brass, Gestión de Riesgos). PR-1 de la cadena. Commit: pending.
- T2 (S4, S5) docs(migracion): redactar la matriz legacy→web con una fila por capacidad consolidada. PR-2 de la cadena. Commit: pending.
- T3 (S1, S3, S6) docs(capacidades): incorporar HPS, HPS Solicitudes y No Conformidades cuando su fase 5 aterrice (#1076, #1077, #1080). PR-3 de la cadena. Commit: pending.

Ruta: la unidad de trabajo es la documentación del índice y de la matriz; no hay código de aplicación. La aprobación del merge es del mantenedor; el camino gobernado es `governed-merge.yml`.

## Log

L1. Petición original (verbatim): «dale sigue, qiero trabajar igual que como hace andres, para que lo tengas en cuenta. de ahi el que siempre quiera estar sincronizado y al dia con el.» y «y levantame los desarrollos para ver como va cada cosa».
L2. Decisión de alcance (2026-10-30): el lote de merges de las PRs de documentación (#1077, #1080) se ejecuta por el camino gobernado en paralelo a esta cadena; la consolidación no espera a ese lote porque no toca los mismos ficheros.
L3. Decisión de autoridad (2026-10-30): el índice y la matriz resumen y enlazan; la evidencia (ficha, commit fijado, cita `ruta:línea`) vive en el PRD de cada aplicación. Escribir en el índice lo que el PRD ya dice sería prosa paralela que duplica una autoridad existente.
L4. Hueco conocido (2026-10-30): el ERD de origen de las ocho aplicaciones depende de la extracción `dysflow-erd` en el runner Windows del propietario. Ninguna unidad de esta cadena lo espera; los documentos lo declaran abierto.
