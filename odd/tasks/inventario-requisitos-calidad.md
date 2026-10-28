# Inventario de requisitos de las aplicaciones de la Lanzadera para Calidad

Issue #1494 · rama `docs/1494-inventario-requisitos-calidad` · worktree `access2web-blueprint-worktrees/docs-1494-inventario-requisitos-calidad`.
Inventario exhaustivo, por aplicación, de lo que hace hoy cada aplicación legacy, para que Calidad decida qué se conserva antes de migrar a web, y para que sirva de backlog.

## Specs

- **S1 — Alcance.** Todas las aplicaciones con icono en las pestañas «En oficina» y «Fuera de la oficina» de la Lanzadera (no la de pruebas), más la propia Lanzadera. Usuario: «pero han de ser todas las aplicaciones, mira a ver si falta alguna»; «revisa el legacy de lanzadera.accdb para ver qué iconos aparecen tanto en fuera de la oficina y dentro ( no la de pruebas)». Excluidas: «AGEDO ya no la quieren , la puedes retirar»; SICA, excluida por el usuario. Resultado: Lanzadera, Gestión de Riesgos, No Conformidades, Condor, Solicitudes HPS, HPS, Brass, Expedientes, Seguridad, Gestión Contratos, AGEDYS técnico (`00_AGEDYS`), AGEDYS económico (`00_AGEDYS_ECO`) y Registro (`00_REGISTROS`).
- **S2 — Rama base.** Usuario: «Ha de ser exahustiva. una por una fijándonos en main cuando no hay rc1 (que se esté realizando una nueva release) si hay una rc1 los requisitos son los del rc1, si no hay eso tienes que mirar en main de ese aplicativo». Una rc1 cuenta si tiene commits que no están en `main`. Cada inventario cita el SHA completo.
- **S3 — Doble propósito.** Usuario: «hay que hacer un levantamiento de requisitos que sirva para dos cosas 1. que se pueda listar con un pequeña descripción para que los miembros de Calidad (no son informaticos) los puedan identificar para decir si se quedan o no 2. que puedas servir para futuro backlog para generar la herramienta web con los que se queden».
- **S4 — Formulario y botón.** Usuario: «Si puede ser identificando el Formulario y el botón donde se aplica ese requisito (si es de ese tipo) para que ellos lo tengan aún más facil para identificarlos». Se usa el texto visible (título del formulario, texto del botón o de la etiqueta); el nombre interno va en la capa técnica.
- **S5 — Feature y reglas de negocio.** Usuario: «y luego lo que hace por detrás ese comando , por ejemplo el comando de publicar edición en gestión de riesgos, la feature es publicar edición y luego las reglas de negocio que se cumplen para poder llegar a ese objetivo»; «por si el requisito se queda pero el alcance cambie». Cada regla es una fila con decisión propia.
- **S6 — Formato para Calidad.** Excel (elección del usuario). Una fila por feature y una por regla; decisión Mantener / Retirar en la feature, Mantener / Cambiar / Quitar en la regla, y un comentario.
- **S7 — Exhaustividad comprobable.** Todo formulario de `src/forms` en el SHA fijado queda cubierto por alguna feature o justificado como no aplicable (subformulario auxiliar, formulario sin uso, etc.), con totales.
- **S8 — Gobierno.** Una decisión nueva sustituye a D85 con el alcance de S1.

## Tasks

| ID | Specs | Ruta | Commit |
|---|---|---|---|
| T1 | S3-S7 | Esquema de la fuente (YAML por aplicación) y generador del Excel con tests (inline) — hecho | 27625a23 |
| T2 | S2-S7 | Piloto: Gestión de Riesgos (`release/2026-008.2.rc1` `99ffa81df3f56a80e6f713e3afc0fb73fd9b3b25`) (worker) — hecho, área Publicación/Ediciones | (este commit) |
| T3 | S3-S6 | Validación del piloto con el usuario — hecho | (este commit) |
| T4 | S1-S7 | Completar Gestión de Riesgos y las 7 aplicaciones restantes con PRD (un worker por aplicación) | |
| T5 | S1-S7 | Las 5 aplicaciones nuevas (un worker por aplicación) | |
| T6 | S8 | Decisión que sustituye a D85 — hecho | f4233dab |

## Log

- L1 — Petición original: «me piden que todas las bases de datos que hay ahora en la lanzadera (access legacy) saquemos un documento de features que llevan, levantando todas las features, para que cuando tengamos que pasar a web, los usuarios le echen un vistazo y vean cuales nunca se ha usado para retirarla.»
- L2 — Ramas base (2026-10-28): Riesgos `release/2026-008.2.rc1` 99ffa81d; No Conformidades `release/2026-009.rc1` 49c257cc; Expedientes `release/2026-008.2.rc1` 7ecda391; Brass `2026-004.rc1` cb279512; Lanzadera `main` f86b4267; Condor `main` 752fde2d; Solicitudes HPS `main` cfa1937a; HPS `main` dc194f4c (su rc1 ya está en main); Seguridad `main` 6dd63d68 (su rc1 no aporta código); Contratos `main` effb968c; AGEDYS técnico `main` 350896d1; AGEDYS económico `main` 424b8bde; Registro `main` 44f36d37.
- L3 — `TbAplicacionesAperturas` deja de registrar a finales de 2024: no sirve como indicador de uso.
- L4 — T1: `openpyxl` no es dependencia del repositorio; el generador usa solo la biblioteca estándar y PyYAML. 12 tests en verde; el fichero abre en openpyxl sin avisos. El commit 27625a23 suma 629 líneas (script 291 + tests 230): la PR de T1 supera el presupuesto de 400 y hay que decidir la estrategia de entrega antes de abrirla. Fallos preexistentes en local, iguales con y sin el cambio: `test_governance_seals`, `test_check_suites_populated`, `test_require_ci_success`, `test_pr_gate_h1` y `test_operating_doc_drift_pin`.
- L5 — Revisión nativa de T1+T6 (`review-b9622c8aab19e7f7`, lentes legibilidad y fiabilidad): aprobada y acusada; 8 hallazgos informativos no bloqueantes (estilo reutilizado en hojas auxiliares, `prefijo` no validado, helper de test que ignora celdas vacías, entre otros) quedan como trabajo posterior.
- L6 — Piloto T2: 63 formularios (11 cubiertos, 5 no aplicables con motivo real, 47 pendientes de piloto), 25 funcionalidades, 180 reglas, 15 posibles defectos marcados. Muestra de evidencias contrastada con el código. Convención de títulos: los formularios cambian su Caption en tiempo de ejecución a «Usuario conectado: …» (`src/classes/Entorno.cls:1448`), así que se usa la etiqueta de cabecera visible o la pestaña. Excel generado en `access2web-blueprint-worktrees/_entregables/requisitos-gestion-riesgos-PILOTO.xlsx`.
- L7 — Usuario: «El piloto estaba muy bien». Formato validado; convenciones del piloto escritas en `docs/07-migracion/inventario-requisitos.md`. Copias de solo lectura del código de cada aplicación en `access2web-blueprint-worktrees/_legacy-src/<app>-<sha8>/`.
- L8 — Lanzadera: 28 formularios (21 cubiertos, 7 no aplicables), 83 funcionalidades, 280 reglas, 25 posibles defectos. Verificado: generador exit 0, recuento de formularios igual al extracto, muestra de evidencias y el acceso sin contraseña desde equipo de administrador (`src/forms/Form_FormLogin.cls:116`). Excel en `_entregables/requisitos-lanzadera.xlsx`.
- L9 — Solicitudes HPS: 34 formularios (32 cubiertos, 2 no aplicables), 75 funcionalidades, 399 reglas, 28 posibles defectos. Verificado con `_tools/check_app.py`: generador exit 0, formularios idénticos al extracto, 791 evidencias en rango, muestra coherente. Los correos solo se encolan; el envío real está fuera del repositorio (`src/classes/CorreoOperaciones.cls:149`).
- L10 — Gestión de Riesgos completa (aec0cdc7): 63 formularios (56 cubiertos, 7 no aplicables), 114 funcionalidades, 619 reglas, 42 posibles defectos; 1354 evidencias en rango.
- L11 — Condor: 38 formularios (todos cubiertos), 85 funcionalidades, 392 reglas, 38 posibles defectos; 721 evidencias en rango. Confirmado: `src/modules/Variables Globales.bas:75` fija `EnDesarrollo = "Sí"` al arrancar, así que en `main` todo correo se redirige al usuario conectado (`src/classes/NotificacionServicio.cls:29`).
- L12 — No Conformidades: 49 formularios (48 cubiertos, 1 no aplicable), 79 funcionalidades, 439 reglas, 33 posibles defectos; 1006 evidencias en rango.
- L13 — 2026-10-28: el usuario pide parar («me estoy quedando sin tokens, mañana seguimos»). Cancelados los workers de HPS, Brass, Expedientes, Seguridad, Contratos, AGEDYS técnico, AGEDYS económico y Registro. Solo HPS dejó fichero parcial (`docs/03-aplicaciones/hps/requisitos.yaml`, 51 funcionalidades, sin commit, sin validar). Para retomar: relanzar esos 8 con la misma plantilla de tarea (HPS: completar el parcial) y verificar cada uno con `access2web-blueprint-worktrees/_tools/check_app.py <slug> _legacy-src/<app>-<sha8>` antes del commit.
