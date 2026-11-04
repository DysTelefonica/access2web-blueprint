# 1581 — Fase 0: descubrimiento global de todas las herramientas

Issue #1581. Épica `docs/07-migracion/epica-coexistencia-postgresql.md` (fase 0, decisión 12).
Rama `odd/1581-fase0-descubrimiento-global`. Solo lectura sobre fronts y backends: no se escribe en ningún `.accdb` ni en producción.

## Specs

S1. «yo creo que los accionables se pueden ir catalogando pero hasta el final del descubrimiento de todas las herramients no se comienza a aplicar por si hay sorpresas en otros fronts».
S2. «y si una tiene un rc1 y aún no se han pasado las penas UAT toma el código de esa rama rc1».
S3. «legacy con rc1 sin haber pasado aun uat esa es la rama a estudiar».
S4. «con dysflow puedes sacar el erd de un front y un backend y ahí se obtienen las tablas vinculadas sin usar tablas de sistema que dysflow no puede».
S5. «sólo brass si tiene controles conectadas y tablas vinculadas de su propio backend y el de lanzadera».
S6. «tabla que el código vba no pida no debe estar en su backend. Ningún otro front debe pedir tablas de un backend que su front no pida».
S7. «nunca hay macros ni auto exec. Para mirar el código te basta codegraph. Todos los formularios y controles de rellenan en tiempo de ejecución Formularios desconectados y no hay vínculo de ninguna tabla ...».
S8. Una ficha por aplicación con la base analizada (rama y SHA), el ERD de su front y de su backend, sus tablas propias y vinculadas con origen, y el uso real trazado con CodeGraph.
S9. Veredicto de propiedad por tabla: una tabla solo debe estar en el backend si el VBA de su propio front la necesita; si la usa otro front, está mal ubicada y se reasigna; si no la usa nadie, es candidata a retirada.
S10. Las acciones del catálogo llevan evidencia, comprobación previa y reversión, y ninguna se aplica: el propietario las aplica tras revisar el conjunto.
S11. R-1 y R-2 constan como confirmadas o refutadas, con sus excepciones explícitas.
S12. Los recuentos se verifican sobre copia reciente de producción antes de aprobar cualquier acción; los de desarrollo no bastan.

## Tasks

- T1 (S1, S10) docs(migracion): gate de descubrimiento global en la épica, A1–A7 como propuestas y skill local actualizada. Commit `e9a82d9c` (PR #1577, merge `208e19f9`).
- T2 (S2, S3, S8) docs(migracion): regla de base RC1 con UAT pendiente, aplicada al inventario de aperturas de `Lanzadera_Datos`. Commit `e9a82d9c` (PR #1577).
- T3 (S4) docs(migracion): método de inventario de vínculos por ERD dysflow de front y backend, sin tablas de sistema, con trazado CodeGraph. Commit `e9a82d9c` (PR #1577).
- T4 (S5, S6, S9) docs(brass): inventario de vínculos y controles conectados de Brass y veredicto de propiedad de sus tablas. Issue #1574. Commit `881bf210` (docs/07-migracion/fase-0/brass.md).
- T5 (S8, S9, S12) docs(migracion): ficha de fase 0 de cada backend restante. HPS `9082b9d5` (base `main`: su UAT está aprobada y liberada), Condor `acfec44b`, Expedientes `7ed18e08`, Gestión de Riesgos `b08824d5`, No Conformidades `7a738a26`. Registros: en curso.
- T6 (S8, S11) docs(migracion): grafo de dependencias entre backends a partir de los vínculos, con confirmación o refutación de R-1 y R-2 y sus excepciones. Commit `4d6fd3ed` (docs/07-migracion/fase-0/grafo-dependencias.md).
- T7 (S10, S12) docs(migracion): catálogo único de accionables proposals con evidencia, comprobación previa sobre copia de producción y reversión. Commit: pendiente.
- T8 (S1, S10) docs(arquitectura): resolver DA-7/D82 (`DROP SCHEMA` sobre tablas legacy) antes de que la web escriba. Issue #1559. Commit: pendiente.
- T9 (S1, S10) chore(ci): registrar la fricción de gobierno detectada en la entrega: `safe_merge` no retira el worktree ni la rama local cuando la rama vive en un worktree (team-skills #1501), falta un helper para esperar a la CI tras el push (team-skills #1502) y este repo no declara qué flujo de seguimiento rige (issue #1583). Commit: pendiente en `odd/tasks/1581-fase0-descubrimiento-global.md`.
- T10 (S4, S8) docs(migracion): anexo de controles conectados con el SQL literal de cada control y el catálogo de consultas por tabla. General `11315f61` (PR #1615) y detalle de Brass `fc9c0c8e` (PR #1617). El detalle de Brass viaja aparte por el presupuesto de revisión de 400 líneas.

Ruta: una unidad de trabajo por aplicación, documentos de inventario bajo el presupuesto de revisión. La aprobación del merge es del mantenedor; el camino gobernado es `safe_merge.py` sin `--admin`, conservando la rama remota.

## Log

L1. Petición original (verbatim): «yo creo que los accionables se pueden ir catalogando pero hasta el final del descubrimiento de todas las herramients no se comienza a aplicar por si hay sorpresas en otros fronts».
L2. Corrección (verbatim): «y si una tiene un rc1 y aún no se han pasado las penas UAT toma el código de esa rama rc1» y «legacy con rc1 sin haber pasado aun uat esa es la rama a estudiar».
L3. Corrección (verbatim): «con dysflow puedes sacar el erd de un front y un backend y ahí se obtienen las tablas vinculadas sin usar tablas de sistema que dysflow no puede».
L4. Corrección (verbatim): «sólo brass si tiene controles conectadas y tablas vinculadas de su propio backend y el de lanzadera».
L5. Corrección (verbatim): «tabla que el código vba no pida no debe estar en su backend. Ningún otro front debe pedir tablas de un backend que su front no pida».
L6. Corrección (verbatim): «nunca hay macros ni auto exec. Para mirar el código te basta codegraph. Todos los formularios y controles de rellenan en tiempo de ejecución Formularios desconectados y no hay vínculo de ninguna tabla ...».
L7. Corrección de método (2026-11-04): el seguimiento de este programa es ODD, no SDD. La propuesta SDD/OpenSpec (`openspec/changes/legacy-coexistence-migration/proposal.md`, PR #1579) se cerró sin fusionar y su issue #1578 con ella; el contenido vive en este documento, en la épica y en el espejo Engram `odd/1581-fase0-descubrimiento-global/tasks`.
L8. Decisión de alcance (2026-11-04): el descubrimiento de cada aplicación incluye su front y su backend. El ERD del backend no sustituye al del front, porque Brass tiene vínculos en el front y R-1 ya tiene excepciones demostradas.
L9. Decisión de alcance (2026-11-04): ninguna acción se aplica en el legacy durante la fase 0, ni siquiera renombrados de cuarentena o relaciones nuevas. «Imprescindible» en la clasificación es una categoría provisional, no una orden de ejecución.
L10. Fricción de gobierno registrada (2026-11-04, a petición del propietario): «si notas alguna fricción con el Gobierno de CI no dudes en generar issue para el coordinadora de teams skill si es común para cualquiera proyecto web o una issue aquí para lo que no sea propagable y que sea específico de este repo». Se abrieron team-skills #1501 (retirada de worktree y rama local por `safe_merge`), team-skills #1502 (helper de espera de CI tras el push) y #1583 en este repo (declarar qué flujo de seguimiento rige).
L11. Base de análisis de Brass (2026-11-04): la rama `2026-004.rc1` (SHA `cb2795128748fa91e79b5cf0a0c2501902692e3f`) es la base porque su UAT sigue pendiente y sin firma (`docs/brass/roadmap.md:53, 93`). La rama `feat/rc1-invoice-sla-gate` no sirve: su trabajo está sin commitear y sin integrar en el binario.
L12. Corrección del propietario (2026-11-04): todos los binarios del ecosistema están protegidos con contraseña y es siempre la misma; el front de lanzadera es la excepción, no está protegido. La extracción usa la variable de entorno `ACCDB_PASSWORD` y el valor no se escribe en ningún artefacto.

L13. Inventarios por aplicación entregados (2026-11-04): HPS `9082b9d5`, Condor `acfec44b`, Expedientes `7ed18e08`, Gestión de Riesgos `b08824d5`, No Conformidades `7a738a26`, grafo de dependencias `4d6fd3ed`. HPS rompe la premisa de la RC1: su UAT está aprobada y liberada, de modo que su base es `main` y no `main` la RC1 abandonada.
L14. El SQL de los formularios con origen de datos se obtiene de los exports `.form.txt` (propiedad `RowSource` con su consulta literal), no de `MSysQueries`; el anexo de controles conectados lo recoge (tarea T10).

L15. Corrección de una regla del propietario por medición (2026-11-04): «salvo Brass, los formularios son desconectados» queda refutada. Brass tiene 106 controles con origen de datos en 41 formularios, Registros 37 en 7 formularios y HPS 3 formularios unidos a recordset contra su caché local. No Conformidades, Condor, Expedientes y Gestión de Riesgos no tienen origen de datos en diseño, salvo casos aislados. La regla general sigue siendo la de formularios desconectados, pero cada aplicación se mide. La skill y la épica quedan corregidas.
L16. Estado de la fase 0 al cierre de esta tanda (2026-11-04): ocho fichas de aplicación, el grafo de dependencias, el catálogo único de 84 accionables y el anexo de SQL están publicados. 71 de las 84 propuestas siguen bloqueadas por falta de la copia de producción, del SQL de 64 consultas guardadas y del inventario de cinco backends sin ficha (Correos, AGEDYS, AGEDO, Solicitudes HPS, Tareas). Nada se ha aplicado en el legacy.
