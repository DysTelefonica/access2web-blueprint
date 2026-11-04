---
name: coexistence-migration
description: 'Trigger: migration work, data model, PostgreSQL adapters or DDL, legacy Access backends, coexistence, creating or triaging migration issues. Reglas de la migración por coexistencia progresiva (épica #1545, D92).'
license: Apache-2.0
metadata:
  author: ardelperal
  version: "0.1"
---

## Activation Contract

Load this skill **before** any of:

- Trabajo de migración, de modelo de datos o sobre los backends Access legacy.
- Adaptadores de PostgreSQL, migraciones de Alembic o DDL.
- Crear, redactar o triar una issue de migración.

Las fuentes son la épica [`docs/07-migracion/epica-coexistencia-postgresql.md`](../../docs/07-migracion/epica-coexistencia-postgresql.md) (#1545) y la D92 [`docs/decisions/2026-11-04-d92-tablas-con-forma-legacy-compartidas.md`](../../docs/decisions/2026-11-04-d92-tablas-con-forma-legacy-compartidas.md) (#1547). Esta skill las resume para despachar; ante una discrepancia, prevalecen ellas.

## Hard Rules

1. **Coexistencia progresiva, no corte único.** Las aplicaciones se sustituyen de una en una (strangler fig). La Lanzadera Access sigue abriendo las aplicaciones Access no apagadas; la web contiene solo las aplicaciones web nuevas. Ninguna lanza aplicaciones de la otra y ambas conviven hasta el final.
2. **El corte es por aplicación.** Se deshabilita su botón en la Lanzadera Access y se habilita en la web. La reversión es volver a habilitar el botón de Access: los datos son compartidos.
3. **La Lanzadera coexiste primero** (decisión 12 de la épica). Es la autoridad de identidad: todos los usuarios y el control de permisos de cada herramienta viven en `Lanzadera_Datos`, que comparten las dos Lanzaderas y todas las aplicaciones Access.
4. **R-1 y R-2 son pendientes de verificar en la fase 0.** R-1: cada front consulta solo su backend, que contiene tablas propias y tablas vinculadas desde otros backends. R-2: una tabla vive en el backend de la aplicación que la usa. No las trate como hechos hasta que la fase 0 las confirme.
5. **Tablas con forma legacy durante la coexistencia (D92).** La web trabaja directamente sobre las tablas con forma legacy en PostgreSQL que Access vincula por ODBC. Una sola fuente de verdad mantiene vivo el legacy y hace posible la reversión.
6. **Cambios de esquema solo aditivos.** Columnas nuevas que admiten nulos y tablas nuevas; nunca cambie el significado, el nombre ni el tipo de lo existente: rompería los fronts Access que siguen leyendo esas tablas.
7. **Las mejoras estructurales se hacen en Access.** Se aplican en el backend `.accdb`, con sus fronts adaptados, antes de migrar ese backend. La refactorización a un esquema moderno llega solo tras el apagado completo del legacy.
8. **Nunca `DROP SCHEMA` como reversión** sobre esquemas con tablas con forma legacy: destruiría datos que Access sigue usando. DA-7 y D82 de `docs/architecture.md` están pendientes de revisión por este motivo.
9. **Las tablas que se unen comparten motor.** Una unión mixta `.accdb` + ODBC se resuelve en local dentro de Citrix, con lentitud; los backends cuyas tablas se unen entre sí se trasladan juntos.
10. **La épica es un documento vivo.** Los cambios de rumbo se registran en su sección «Cambios de rumbo»; nunca se editan fases en silencio. Una fase equivocada se marca `descartada` con su motivo.

## Convenciones del legacy confirmadas por el propietario

- No hay macros ni `AutoExec`; no se bloquea un inventario esperando inspeccionarlas.
- Los formularios son desconectados. Sus controles y datos se rellenan en tiempo de ejecución: trace los procedimientos de carga y los eventos, no deduzca ausencia de uso por propiedades de diseño vacías.
- Los fronts no tienen tablas vinculadas de sus backends. Las tablas vinculadas de otros aplicativos están en el backend propio; no proponga trasladar esos vínculos al front.
- Para revisar código utilice CodeGraph como primera herramienta. Use `projectPath` en el MCP multirrepositorio; una limitación del subagente no demuestra que CodeGraph no funcione. Declare las aristas dinámicas y cualquier fallback, sin convertir ausencia de llamadores estáticos en prueba de código muerto.

## Decision Gates

| Situation | Action |
|---|---|
| Tabla sin uso aparente (huérfana) | Aporte evidencia: vínculos de todos los backends y búsqueda en el `src` del front propietario, incluido el SQL dinámico. Proponga cuarentena (copia de seguridad y renombrado a `zzBAJA_<tabla>`, o traslado a un `.accdb` de archivo) y borrado tras un periodo sin incidencias. El propietario borra en el legacy; el agente solo aporta la evidencia. |
| Issue nueva de migración | Etiqueta `estrategia:coexistencia`. |
| Issue previa cuyo alcance cambia con la coexistencia | Etiqueta `coexistencia:ajuste` y comentario que explique qué cambia, con enlace a la épica y a la D92. |
| Issue previa que solo tenía sentido con corte único | Etiqueta `coexistencia:obsoleta` y el mismo comentario. |
| Issue previa válida en ambos modelos o ajena a la migración | Sin etiqueta. |
| Una fase o hipótesis resulta equivocada | Entrada fechada en «Cambios de rumbo» de la épica; la fase pasa a `descartada`. |

## Anti-patterns

| Síntoma | Fix |
|---|---|
| Proponer un esquema web propio con sincronización | Descartado por la D92 (opciones B y C). La web usa las tablas con forma legacy. |
| DDL que renombra o cambia el tipo de una columna legacy durante la coexistencia | Solo cambios aditivos; la mejora se hace antes en Access o tras el apagado. |
| Suponer que la web lanza aplicaciones Access | Cada Lanzadera abre solo las suyas. |
| Borrar una tabla legacy sin el protocolo de evidencia | Evidencia, cuarentena y borrado por el propietario. |

## Output Contract

Antes de cerrar trabajo de migración:

- [ ] El DDL o la migración es aditivo y no altera tablas con forma legacy.
- [ ] La reversión no borra esquemas con tablas con forma legacy.
- [ ] La issue lleva la etiqueta de coexistencia que corresponde, o ninguna si es válida o ajena.
- [ ] Todo cambio de rumbo consta en la épica.

## References

- [`docs/07-migracion/epica-coexistencia-postgresql.md`](../../docs/07-migracion/epica-coexistencia-postgresql.md) (#1545) — decisiones del propietario, fases, reglas del legacy, protocolo de huérfanas e hipótesis.
- [`docs/decisions/2026-11-04-d92-tablas-con-forma-legacy-compartidas.md`](../../docs/decisions/2026-11-04-d92-tablas-con-forma-legacy-compartidas.md) (#1547) — D92 y sus consecuencias.
- `docs/architecture.md` §Persistence — DA-7 y D82, en revisión.
- Skill hermana `architecture-guardrails` — front-door y gestión de D-<n>.
