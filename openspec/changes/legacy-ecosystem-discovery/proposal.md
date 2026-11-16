<!-- Change: legacy-ecosystem-discovery · Project: access2web-blueprint · Date: 2026-11-15 -->

# Propuesta: descubrimiento del ecosistema legacy y su puerta de remediación

> **Frase que organiza**: el programa de migración por coexistencia solo puede tocar el legacy cuando el descubrimiento de todas las herramientas está completo y el propietario ha validado la evidencia; a partir de ahí, cada cambio estructural se aplica en el backend Access dueño de la tabla, de forma reversible y aplicación a aplicación.

## Contexto

La migración del ecosistema Access/VBA a PostgreSQL no puede hacerse de un golpe: hay once aplicaciones en el alcance, cada una con su front y su backend `.accdb`, y los backends comparten tablas por vínculos. Un cambio estructural mal planteado rompe un front que sigue en producción.

Este `change` fija el programa de descubrimiento que precede a toda remediación, y la puerta que lo gobierna. La estrategia de coexistencia, sus decisiones y el estado de cada fase viven en la [épica](https://github.com/DysTelefonica/access2web-blueprint/blob/main/docs/07-migracion/epica-coexistencia-postgresql.md); aquí solo se fija qué capacidades introduce el programa y qué las hace aplicables.

El descubrimiento ya se ejecutó y su evidencia está validada por el propietario (2026-11-15, decisión 28). Lo que este `change` deja established es el contrato que lo hizo posible y que sigue vigente: la puerta de remediación.

## Alcance

### Incluye

- **Capacidad `legacy-discovery`**: inventario de solo lectura de cada front y cada backend del alcance: vínculos y sus orígenes, clave primaria, tipos problemáticos, reglas de validación, relaciones, huérfanos y duplicados, patrones de riesgo del VBA de cada front, clasificación de cada tabla y coste medido en código por tabla y por columna.
- **Capacidad `legacy-remediation-gate`**: la regla de que ningún cambio en el legacy se aplica hasta que el descubrimiento cubre todas las aplicaciones del alcance, las dependencias están reconciliadas y el propietario ha validado el catálogo conjunto de accionables.
- La base de análisis por aplicación: la rama RC1 con UAT pendiente cuando la hay, y `main` en caso contrario, con su SHA registrado y sin mezclar versiones.
- La propiedad de las tablas: una tabla solo permanece en el backend si el VBA de su propio front la necesita; la consumen otros front no la justifica.
- El catálogo de remediación: propuestas con su evidencia, su comprobación previa y su reversión, una issue por aplicación y por repositorio de aplicación.

### No incluye

- Aplicar las propuestas del catálogo. La fase 1 es el trabajo siguiente y lo aplica el propietario en el legacy; el agente aporta la evidencia y la issue.
- El modelo de datos moderno de la web y sus migraciones: llegan tras el apagado completo del legacy.
- El corte de cada aplicación a la web, la elección del driver ODBC y el aprovisionamiento de PostgreSQL.
- El borrado o el traslado de tablas del legacy: son decisiones del propietario, con su protocolo de cuarentena.

## Capacidades

### Nuevas

- `legacy-discovery`: inventario de solo lectura de fronts, backends, vínculos y patrones de VBA, con la clasificación de cada tabla y el catálogo de propuestas de remediación.
- `legacy-remediation-gate`: ningún cambio en el legacy se aplica mientras el descubrimiento no cubra todas las aplicaciones y las dependencias no estén reconciliadas.

### Modificadas

Ninguna: `openspec/specs/` no tiene capacidades archivadas del legacy.

## Enfoque

1. **Base.** Por aplicación, la rama con UAT pendiente si la hay; si no, la base publicada. Se registran rama y SHA; no se mezclan versiones.
2. **Lectura estructural.** ERD del front y del backend con dysflow. Los vínculos y sus orígenes salen de ahí; no se consultan tablas de sistema.
3. **Lectura de código.** Uso real de cada tabla con CodeGraph, siguiendo los-sql dinámico y los controles conectados de Brass. La ausencia de un llamador estático no prueba que el código sea muerto.
4. **Contraste.** ERD, uso en código y regla de propiedad, contra el grafo de dependencias de todos los backends.
5. **Clasificación y precio.** Cada tabla con su clase y su coste medido en referencias de código.
6. **Catálogo, no aplicación.** Las propuestas se publican con evidencia, comprobación previa y reversión. No se aplica ninguna hasta que la puerta se abre.

## Superficies afectadas

| Superficie | Impacto | Descripción |
|---|---|---|
| `docs/07-migracion/epica-coexistencia-postgresql.md` | Modificado | Épica viva: fases, condición de entrada de la fase 1, registro de cambios de rumbo y decisiones del propietario. |
| `docs/07-migracion/fase-0/` | Nuevo | Inventarios por backend y por front, clasificación, grafo, patrones de riesgo, coste y catálogo de accionables. |
| `skills/coexistence-migration/SKILL.md` | Modificado | Skill local: reglas del propietario, puertas de decisión y antipatrones. |
| `docs/decisions/` | Nuevo | D92 (tablas con forma legacy compartidas) y D93 (dos esquemas por aplicación). |
| `scripts/dysflow-erd/` | Nuevo | Scripts de extracción, clasificación, patrones de riesgo y coste. |
| `openspec/changes/legacy-ecosystem-discovery/` | Nuevo | Esta propuesta y su exploración. |

## Riesgos

| Riesgo | Probabilidad | Mitigación |
|---|---|---|
| Una tabla juzgada muerta la lee un front fuera de los repositorios analizados, o a través de un control conectado de Brass. | Media | Puerta de descubrimiento: sin borrado sin evidencia de todos los fronts, y el propietario aplica tras una ventana de cuarentena. |
| Remediar un backend pronto invalida otro front que vincula sus tablas. | Media | Sin remediación antes del descubrimiento global y de la reconciliación de dependencias. |
| `DROP SCHEMA` como reversión destruye tablas con forma legacy que Access sigue leyendo. | Alta | D92 y D93: el esquema legacy no se destruye nunca; DA-7 y D82 siguen en revisión (#1559). |
| Las copias de desarrollo divergen de los datos de producción y los recuentos son falsos. | Media | Los recuentos se verifican contra la copia de producción de referencia antes de aprobar cualquier acción. |
| Un vínculo de desarrollo apunta al share de producción y una escritura de desarrollo llega a los datos reales. | Media | Detectado durante la extracción de Expedientes; la higiene de vínculos es parte del descubrimiento. |
| La revisión se desborda con un inventario grande en un solo cambio. | Media | Una aplicación por unidad de trabajo, por debajo del presupuesto de 400 líneas. |

## Reversión

El programa produce documentos: se revierte el commit o se descarta el documento de inventario. Los datos del legacy nunca se tocan, así que no hay nada que revertir en el legacy.

Cuando se apruebe una remediación, cada acción lleva su reversión: cuarentena por renombrado a `zzBAJA_<tabla>`, borrado desde una copia validada y retirada de relaciones para la integridad referencial. El propietario aplica y revierte en el legacy; el agente aporta la evidencia.

## Dependencias

- dysflow para el ERD de fronts y backends.
- CodeGraph (inicializado a través de Gentle AI) para el uso en código.
- La rama RC1 de cada aplicación con UAT pendiente, y su SHA.
- Copias de producción para verificar recuentos, en solo lectura.

## Criterios de éxito

- [x] Cada aplicación del alcance tiene su base (rama y SHA) y el ERD de su front y su backend.
- [x] Cada backend tiene su lista de vínculos y sus orígenes, sin consultar tablas de sistema.
- [x] Cada tabla tiene clasificación, coste medido y veredicto de propiedad.
- [x] Cada front tiene su inventario de patrones de riesgo, con los controles conectados de Brass.
- [x] R-1 y R-2 constan como confirmadas o refutadas, con sus excepciones.
- [x] Existe un catálogo único de propuestas de remediación, con evidencia, comprobación previa y reversión por entrada.
- [x] Ningún fichero, base de datos ni dato del legacy se ha modificado durante el programa de descubrimiento.
- [ ] El propietario aplica las propuestas del catálogo, aplicación a aplicación, en el orden de la decisión 24.
