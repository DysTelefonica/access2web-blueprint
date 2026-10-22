# Roadmap de descubrimiento y PRD por aplicación

> **Issue:** #1069 · **Estado:** vigente desde 2026-10-21 · **Plantilla:** [plantilla-prd-aplicacion.md](plantilla-prd-aplicacion.md)

## Propósito

Este documento ordena el paso del descubrimiento inicial de las ocho aplicaciones Access/VBA a un análisis suficiente para migrarlas a web. Para cada aplicación separa lo verificado, lo documentado previamente y lo pendiente, y fija la secuencia, el método y los entregables de los estudios.

No declara terminado ningún estudio. Un inventario de ficheros no prueba que una capacidad esté analizada.

## Fuente de verdad

- **El código `src` del repositorio de cada aplicación es la única autoridad sobre el comportamiento.** Engram, las épicas, los walkthroughs y las issues previas son contexto. Una discrepancia entre ese contexto y `src` se registra como discrepancia abierta; no se resuelve por inferencia.
- Cada estudio fija el commit fuente al empezar y lo cita en todos sus entregables. Si el commit cambia durante el estudio, el PRD registra el cambio y vuelve a reconciliar los formularios afectados.
- La trazabilidad estructural (handlers, llamadas, referencias a tablas) se obtiene con CodeGraph-VBA sobre un clon de solo lectura del commit fijado y se contrasta con la lectura del `src`. El índice CodeGraph del blueprint (Python) no es evidencia del legacy.
- Las referencias dinámicas que no se pueden resolver estáticamente (SQL concatenado, `Eval`, nombres de formulario en variables, menús leídos de JSON) quedan como pendientes explícitos.

## Commits de partida (2026-10-21)

Recuento obtenido del árbol git remoto de cada repositorio (`src/forms/*.cls` y `src/modules|classes/*`), comparado con el snapshot `data/staging/<app>/src` del 2026-08-06. Para Expedientes, la decisión del propietario fija la release candidate `release/2026-008.2.rc1`; para el resto, el estudio fija el commit de `main` vigente al empezar.

| Aplicación | Repositorio | Commit de partida | Formularios `src` | Módulos y clases | Formularios snapshot | Delta |
|---|---|---|---|---|---|---|
| Expedientes | `DysTelefonica/Expedientes` | `ec6a8c3` (RC1) | 45 | 142 (82 `.bas`, 60 `.cls`) | 46 | −1; `main` `13af0f6` tiene 42 formularios y 74 módulos y clases |
| Lanzadera | `DysTelefonica/Lanzadera` | `1474e8e` | 28 | 31 | 28 | 0 |
| Condor | `DysTelefonica/condor` | `c88347e` | 38 | 93 | 45 | −7 |
| HPS | `DysTelefonica/HPS` | `5e64cf9` | 38 | 59 | 38 | 0 |
| HPS Solicitudes | `DysTelefonica/HPS_SOLICITUDES` | `cfa1937` | 34 | 40 | 34 | 0 |
| Brass | `DysTelefonica/BRASS` | `3f77d58` | 87 | 44 | 87 | 0 |
| Gestión de Riesgos | `DysTelefonica/GESTION_RIESGOS` | `36fcf60` | 63 | 67 | 62 | +1 |
| No Conformidades | `DysTelefonica/No_conformidades` | `1d7a2f8` | 48 | 57 | 48 | 0 |

La diferencia de Expedientes entre RC1 y `main` (142 frente a 74 módulos y clases) no está explicada; el estudio de Expedientes la resuelve antes de cerrar su inventario.

Un delta 0 no significa que el código sea igual: solo que coincide el número de formularios. Cada estudio compara el contenido contra el snapshot y contra los walkthroughs previos, y clasifica cada capacidad como añadida, modificada, retirada o sin cambios.

Repositorios con actividad reciente (HPS y Condor recibieron pushes el 2026-10-21) pueden avanzar durante el estudio. Por eso cada PRD cita su propio commit y no el de esta tabla.

## Estado por aplicación

Leyenda: **verificado**, comprobado contra `src` en el commit citado; **documentado previamente**, existe en el repositorio pero no se ha revalidado contra el commit de partida; **pendiente**, sin evidencia.

| Aplicación | Verificado | Documentado previamente | Pendiente |
|---|---|---|---|
| Expedientes | Inventario de formularios RC1 (45 `.cls` con sus 45 `.form.txt`); índice CodeGraph-VBA de RC1 construido (232 ficheros VBA) | Épica DRAFT v0.1, 5 walkthroughs (46 formularios, snapshot), 26 fichas `CAP-*`, `data-model.md`, `ERD/schema.sql`, `migration-matrix.md` | Delta RC1 frente al estudio previo, PRD por formulario, 37 fichas de #807, ERD de origen revalidado, matriz campo a campo, coexistencia |
| Lanzadera | Recuento de formularios | Épica DRAFT v0.2, 5 walkthroughs, `ERD/schema.sql`, módulo web en `app/src/modules/lanzadera` | PRD por formulario, reconciliación del MVP web con `src`, ERD de origen, coexistencia |
| Condor | Recuento de formularios | Épica DRAFT v0.1, 5 walkthroughs (45 formularios), `ERD/schema.sql` | Explicar el delta de −7 formularios, PRD, ERD de origen, matriz, coexistencia |
| HPS | Recuento de formularios | Épica DRAFT v0.1, 5 walkthroughs, `ERD/schema.sql` | PRD, auditoría de uso del esquema, PII en caché local, ERD, matriz, coexistencia |
| HPS Solicitudes | Recuento de formularios | Épica DRAFT v0.1, 5 walkthroughs, `ERD/schema.sql` | PRD, auditoría de uso de sus 11 tablas (`epic.md:43,128`), ERD, matriz, coexistencia |
| Brass | Recuento de formularios | Épica DRAFT v0.1, 8 walkthroughs (incluidos los parciales `*-brass-missing`), `ERD/schema.sql` | PRD, cierre de los walkthroughs parciales, ERD (sin backend separado), matriz, coexistencia |
| Gestión de Riesgos | Recuento de formularios | Épica DRAFT v0.1, 5 walkthroughs, `ERD/schema.sql` | Formulario añadido (+1), PRD, máquina de estados, ERD, matriz, coexistencia |
| No Conformidades | Recuento de formularios | Épica DRAFT v0.1, 5 walkthroughs, `ERD/schema.sql`, `data/staging/no-conformidades/docs/schema/` | PRD, ERD revalidado, matriz, coexistencia |

## Secuencia

1. **Expedientes RC1** es el primer estudio, por decisión del propietario. Su PRD fija el formato de referencia.
2. **Las otras siete aplicaciones** se estudian en paralelo, cada una en su propia issue, rama y worktree. Este roadmap no establece prioridades ni fechas entre ellas.
3. **El ERD objetivo PostgreSQL global** se consolida cuando estén entregados los ERD de origen de las ocho aplicaciones. Antes solo pueden proponerse vistas por aplicación.
4. **Las issues de implementación** se derivan del PRD de cada aplicación cuando esté cerrado. No se crean antes ni se fija una cuota de issues.

## Método por aplicación

Cada estudio recorre estas fases en orden. Las fases 2 a 5 se entregan como PR encadenadas de 400 líneas como máximo.

| Fase | Qué se hace | Evidencia mínima |
|---|---|---|
| 0. Fijación | Clonar el commit de partida en solo lectura e indexarlo con CodeGraph-VBA | SHA completo, recuento de ficheros por tipo y estado del índice |
| 1. Inventario | Reconciliar formularios, controles (`.form.txt`), módulos, clases, consultas guardadas e informes | Tabla de inventario con ruta de cada objeto |
| 2. Formularios | Por formulario: navegación de entrada y salida, controles, eventos, handlers, llamadas, reglas, validaciones, cálculos, permisos, consultas, lecturas y escrituras, transacciones, errores, informes e integraciones | Ficha según la plantilla, con `ruta:línea` en cada afirmación |
| 3. Datos | ERD de origen por backend: tablas, columnas, tipos, longitudes, nulabilidad, valores por defecto, claves, índices, relaciones y reglas referenciales; accesos desde `src` | Relaciones declaradas separadas de las inferidas por código; carencias de metadatos listadas |
| 4. Coexistencia | Tablas vinculadas, cadenas de conexión, DAO/ADO, SQL dinámico, sintaxis Jet/ACE y transacciones que deban adaptarse a PostgreSQL vinculado | Lista por formulario de lo que se adapta y de lo que se preserva |
| 5. PRD | Capacidades, flujos normales, alternativos y de error, criterios de aceptación y propuesta de issues futuras | Reconciliación completa del inventario |

### Trazabilidad obligatoria

Cada afirmación del PRD encadena: **control/evento → handler → llamadas → datos/efectos → capacidad → criterio de aceptación → futura issue**. Un eslabón sin evidencia se marca como pendiente y rompe la cadena hasta resolverse.

### Separación de contenidos

El PRD separa siempre tres capas:

- **Comportamiento actual**: lo que hace `src`, aunque parezca incorrecto.
- **Posibles defectos**: comportamiento sospechoso, con la evidencia y sin corregirlo.
- **Propuestas web pendientes de decisión**: cambios para la versión web. No son requisitos hasta que se aprueben.

## Modelo de datos y ERD

### ERD de origen (por aplicación)

Cada estudio entrega el ERD de cada backend de su aplicación a partir de `src` y de `docs/03-aplicaciones/<app>/ERD/schema.sql`. Si `src` no contiene el DDL o los metadatos necesarios, el PRD lista la extracción de metadatos de solo lectura que falta y no inventa la estructura.

### ERD objetivo PostgreSQL (global)

Se entrega en `docs/07-migracion/erd-global-postgresql.md`, con vistas por aplicación, cuando estén los ocho ERD de origen. Debe:

- Identificar entidades compartidas, dependencias entre aplicaciones, colisiones de nombres y el propietario de cada entidad. No fusionar entidades por parecido de nombre.
- Separar esquema actual demostrado, diseño objetivo propuesto y decisiones pendientes de aprobación.
- Resolver la topología: una base con esquemas por aplicación o varias bases. Esta decisión **queda abierta**; si se eligen varias bases, el documento explicita los límites de las claves foráneas entre bases y la alternativa aprobada.

### Matriz campo a campo

Una por aplicación, en `docs/03-aplicaciones/<app>/prd/matriz-campos.md`: aplicación, backend, tabla y columna de origen → base, esquema, tabla y columna de destino, tipo objetivo, transformación, conservación de identificadores y claves foráneas, tratamiento de nulos, duplicados y valores incompatibles, validación y criterio de reconciliación. Toda columna inventariada tiene destino o disposición pendiente justificada.

No se genera ni ejecuta DDL, migraciones ni consolidaciones de datos en estos estudios.

## Coexistencia Access y web

La sustitución es aplicación por aplicación. Al activar la versión web de una aplicación se deshabilita su entrada en la Lanzadera de Access; las aplicaciones no migradas siguen en VBA con tablas PostgreSQL vinculadas que comparten datos con las web. Esto no exige mantener abiertas las dos interfaces de una misma aplicación ni sincronizar dos copias de datos.

Cada PRD incluye, en `prd/coexistencia.md`:

- Qué tablas, consultas y operaciones del formulario deben adaptarse para operar sobre PostgreSQL vinculado y qué comportamiento debe preservarse.
- Claves e identidad de registros actualizables, generación de identificadores, fechas, booleanos, decimales, nulos y valores por defecto. No se afirma compatibilidad sin comprobarla.
- Reglas de negocio que hoy solo aplica VBA y que la web no puede omitir, y viceversa, con la propuesta de restricción compartida pendiente de aprobación.
- Concurrencia: conflictos de edición, bloqueos, refresco de formularios, atomicidad y pérdida de conexión, con política de conflictos explícita.
- Requisitos de conexión ODBC, credenciales, permisos mínimos y auditoría que distinga actor y cliente, sin publicar secretos ni modificar conexiones.
- Pruebas de aceptación bidireccionales (cambio en Access visible en web y al revés, validaciones equivalentes, conflictos concurrentes, integridad referencial, recuperación de errores), a ejecutar después en entorno aislado y nunca sobre datos de producción.
- El corte de la aplicación: carga inicial, reconciliación, cambio de vínculos, activación web, deshabilitación de su entrada en Lanzadera, comprobación de que las demás aplicaciones siguen disponibles y reversión por definir. Deshabilitar la entrada no borra binarios ni datos, ni cierra sesiones abiertas o accesos directos.

## Entregables y ubicación

| Entregable | Ruta |
|---|---|
| PRD de una aplicación | `docs/03-aplicaciones/<app>/prd/README.md` (índice, commit fijado y reconciliación) |
| Ficha por formulario | `docs/03-aplicaciones/<app>/prd/formularios/<Formulario>.md` |
| ERD de origen | `docs/03-aplicaciones/<app>/prd/erd-origen.md` |
| Matriz campo a campo | `docs/03-aplicaciones/<app>/prd/matriz-campos.md` |
| Coexistencia | `docs/03-aplicaciones/<app>/prd/coexistencia.md` |
| Issues propuestas | `docs/03-aplicaciones/<app>/prd/issues-propuestas.md` |
| ERD objetivo global | `docs/07-migracion/erd-global-postgresql.md` |

Las épicas y walkthroughs existentes no se reescriben: el PRD los enlaza y registra dónde los contradice `src`.

## Criterios de cierre de un estudio

- Todos los formularios y controles del inventario están en uno de tres estados: **analizado** con evidencia, **no aplicable** con motivo, o **pendiente** explícito. Los pendientes no cuentan como completados.
- Toda tabla, columna y relación del ERD de origen tiene destino en la matriz o disposición pendiente justificada.
- Las capacidades se clasifican frente al estudio previo como añadidas, modificadas, retiradas o sin cambios.
- Las issues propuestas son unidades implementables, con evidencia, dependencias, criterios de aceptación y etiqueta `app/<slug>`, sin duplicar capacidades compartidas entre aplicaciones.

## Gobierno de los estudios

- Una issue canónica por aplicación, aprobada antes de crear la rama. La issue enlaza este roadmap.
- Cada estudio trabaja en su propio worktree desde `origin/main` y entrega PR encadenadas de 400 líneas como máximo, con `chain:partial` hasta la punta.
- El clon del legacy vive en una ruta ignorada (`.atl/sources/`) y es de solo lectura. No se modifican repositorios legacy, binarios ni bases Access.
