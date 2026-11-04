# Épica: migración por coexistencia progresiva sobre PostgreSQL

> **Issue:** #1545 · **Estado:** documento vivo, vigente desde 2026-11-04 · **Relacionados:** [D14](../decisions/2026-08-04-d14-postgresql-compartido-con-esquemas-por.md), [roadmap de descubrimiento](roadmap-descubrimiento-prd.md), [coexistencia de Lanzadera](../03-aplicaciones/lanzadera/prd/coexistencia.md), [coexistencia de Expedientes](../03-aplicaciones/expedientes/prd/coexistencia.md)

## Propósito

Este documento recoge la estrategia de migración por coexistencia progresiva: las decisiones tomadas, las reglas del legacy, las fases con su estado, las hipótesis pendientes de ensayo y los cambios de rumbo.

No sustituye a los estudios por aplicación. El método de estudio vive en el [roadmap](roadmap-descubrimiento-prd.md); lo que cada aplicación debe adaptar vive en su `prd/coexistencia.md`.

## Cómo se mantiene este documento

El documento es vivo: cambia cuando se descubre que una fase o una hipótesis llevaba a donde no se debía ir.

- **Estados de fase:** `pendiente` (sin empezar), `en curso`, `hecha` (criterio de terminado cumplido con evidencia), `revisada` (el propietario ha validado la evidencia) y `descartada`.
- **Una fase equivocada no se borra.** Se marca `descartada` con el motivo y el enlace a su entrada en [Cambios de rumbo](#cambios-de-rumbo).
- **Las hipótesis se marcan como tales** (`Hipótesis H-n`) y no se tratan como hechos hasta que un ensayo las confirme o las refute.
- **Las decisiones de arquitectura reales se promueven** a `docs/decisions/` como D-n con el formato del [registro](../decisions/README.md). Este documento las enlaza; no las redacta de nuevo.

## Decisiones del propietario (2026-11-04)

| N.º | Decisión |
|---|---|
| 1 | La estrategia pasa de corte único (la web solo entra en servicio cuando todas las aplicaciones legacy estén migradas y probadas) a coexistencia progresiva (patrón strangler fig). |
| 2 | La Lanzadera Access y todas las aplicaciones Access se alojan en un servidor de aplicaciones Windows publicado por Citrix. La plataforma web, en la misma máquina o en una conectada a ella. |
| 3 | Todas las tablas de todos los backends `.accdb` legacy pasan a un único PostgreSQL, con esquemas por aplicación ([D14](../decisions/2026-08-04-d14-postgresql-compartido-con-esquemas-por.md)). Los backends legacy las vinculan por ODBC. |
| 4 | La Lanzadera Access sigue abriendo las aplicaciones Access aún no apagadas; la Lanzadera web contiene solo las aplicaciones web nuevas. Ninguna lanza aplicaciones de la otra. Ambas conviven hasta el final. |
| 5 | El corte es por aplicación: se deshabilita su botón en la Lanzadera Access y se habilita en la web. La reversión consiste en volver a habilitar el botón de Access; los datos son compartidos. |
| 6 | Quedan fuera de alcance Sirdee (pertenece a otro departamento y no figura en la Lanzadera) y AGEDO (no se usa). |
| 7 | Regla R-1 del legacy (ver [Reglas del legacy](#reglas-del-legacy)). |
| 8 | Regla R-2 del legacy (ver [Reglas del legacy](#reglas-del-legacy)). |
| 9 | Este documento es vivo, con estados de fase, hipótesis marcadas y registro de cambios de rumbo. |
| 10 | Durante la coexistencia, la web trabaja directamente sobre las tablas con forma legacy en PostgreSQL: una sola fuente de verdad, para garantizar que el legacy sigue funcionando y que la reversión es posible. No hay esquema web propio con sincronización. Cuando el legacy esté apagado por completo, el modelo se refactoriza a un esquema moderno para dejar de arrastrar la deuda técnica de las tablas legacy. Resuelve P-EXP-3 (ver [P-EXP-3](#p-exp-3)). |
| 11 | La estructura de datos se mejora en el legacy antes de que cada backend pase a PostgreSQL, donde no cueste demasiado trabajo, para que la web nunca empiece sobre tablas con defectos inaceptables. Toda mejora estructural se hace en el backend Access (con sus fronts adaptados) antes de migrar ese backend. Durante la coexistencia, los cambios de esquema son solo aditivos (columnas nuevas que admiten nulos, tablas nuevas) y nunca cambian el significado de lo existente. Ver la [fase 1](#fase-1-saneamiento-en-access). |
| 12 | La Lanzadera es la primera aplicación en coexistir, porque alberga el control de permisos del resto de herramientas y tiene a todos los usuarios y sus permisos (`Lanzadera_Datos`). El ensayo de la fase 3 se hace con una copia de `Lanzadera_Datos` en preproducción, no con una isla. Ver [Cambios de rumbo](#cambios-de-rumbo). |

## Modelo de coexistencia

```text
Servidor de aplicaciones Windows (Citrix)          Plataforma web (misma máquina o conectada)
  Lanzadera Access                                   Lanzadera web
    └── aplicaciones Access no apagadas                └── aplicaciones web nuevas
          │                                                  │
          ▼                                                  │
    backends .accdb (tablas vinculadas por ODBC)             │
          │                                                  │
          └──────────────► PostgreSQL ◄──────────────────────┘
                     (un esquema por aplicación, D14)
```

Cada aplicación está en un solo lado a la vez: su botón está habilitado en una Lanzadera y deshabilitado en la otra. Los datos no se duplican ni se sincronizan entre copias: la web usa las mismas tablas con forma legacy (decisión 10).

## Alcance

- **Dentro:** las aplicaciones del catálogo `TbAplicaciones` que abre la Lanzadera Access, salvo las excluidas.
- **Fuera:** Sirdee y AGEDO.
- **Tablas de AGEDO vinculadas por AGEDYS TEC:** AGEDYS TEC vincula 4 tablas de AGEDO. Viajan con el grupo AGEDYS o se eliminan si no se usan; lo decide la fase 0.

## Reglas del legacy

Ambas reglas proceden del propietario y están **pendientes de verificar en la fase 0**.

**R-1. Cada front consulta solo su propio backend.** Un backend contiene dos clases de tablas: las propias (locales) y las vinculadas desde otros backends que su front también necesita.

Ejemplo: para mostrar el nombre del responsable de un expediente, el front de Expedientes une la tabla de responsables (propia, en `Expedientes_datos`, con el identificador del usuario) con la tabla de usuarios, que `Expedientes_datos` vincula desde `Lanzadera_Datos`.

Excepción conocida: `AGEDYS_DATOS.accdb` es el backend de tres fronts (AGEDYS, AGEDYS ECO y AGEDYS TEC).

Consecuencia, si R-1 se confirma: el grafo de dependencias entre aplicaciones se obtiene leyendo solo los vínculos de los backends, sin abrir los fronts.

**R-2. Cada tabla vive en el backend de la aplicación que la usa.** Si la aplicación B necesita una tabla del backend de A y A no la usa, la tabla pertenece a B.

## Evidencia de partida

Lectura de solo lectura con dysflow sobre `Lanzadera_Datos.accdb` (el backend de la Lanzadera Access), 2026-11-04.

| Hecho | Valor |
|---|---|
| Filas de catálogo en `TbAplicaciones` (front en `NombreEjecutable`, backend en `NombreArchivoDatos`) | 19 |
| Filas de la matriz `TbTablasAVincular` (`IDBBDD`, `IDAplicacion`, `NombreTabla`, `NombreTablaEnLocal`) | 411 |
| Correspondencia de `IDBBDD` | Coincide con el identificador de la aplicación propietaria del backend |

Tablas por backend y aplicaciones que las vinculan, según `TbTablasAVincular`:

| `IDBBDD` | Backend | Tablas | Aplicaciones que las vinculan |
|---|---|---|---|
| 2 | `AGEDYS_DATOS` | 86 | AGEDYS ECO, AGEDYS TEC, AGEDO, Registro, Seguridad |
| 12 | `Lanzadera_Datos` | 15 | Lanzadera, Registro, AGEDYS TEC, AGEDYS ECO, AGEDO, No Conformidades |
| 10 | Registro | 9 | AGEDYS ECO, AGEDYS TEC, AGEDO, Registro |
| 4 | AGEDO | 65 | AGEDO, AGEDYS TEC |
| 6 | Brass | 75 | Solo la propia aplicación (isla) |
| 16 | Seguridad | 30 | Solo la propia aplicación (isla) |
| 1 | Sirdee | 26 | Solo la propia aplicación (isla; fuera de alcance) |

«Isla» significa que ninguna otra aplicación vincula ese backend. Seguridad, además, vincula tablas de `AGEDYS_DATOS`.

### Límites de la evidencia

- La matriz no cubre las aplicaciones modernas (Expedientes, Riesgos, HPS, Solicitudes HPS, Condor y Contratos), que se revinculan por su cuenta (por ejemplo, `BackendResolver.bas` en Expedientes).
- La Lanzadera Access no usa vínculos: `getdb()` abre el backend con `OpenDatabase` y la clave de `TbConfiguracionBackends` (ver [coexistencia de Lanzadera](../03-aplicaciones/lanzadera/prd/coexistencia.md), apartado 1).
- `MSysObjects` del front `Lanzadera.accdb` no es legible (sin permiso de lectura).

### Consultas de reproducción

Sobre `Lanzadera_Datos.accdb`, en modo de solo lectura:

```sql
SELECT COUNT(*) FROM TbAplicaciones;

SELECT COUNT(*) FROM TbTablasAVincular;

SELECT IDBBDD, COUNT(*) AS Tablas
FROM (SELECT DISTINCT IDBBDD, NombreTabla FROM TbTablasAVincular)
GROUP BY IDBBDD;

SELECT DISTINCT IDBBDD, IDAplicacion
FROM TbTablasAVincular
ORDER BY IDBBDD, IDAplicacion;
```

El nombre de cada aplicación se resuelve con su identificador en `TbAplicaciones`.

## Fases

| Fase | Nombre | Estado |
|---|---|---|
| 0 | Inventario | pendiente |
| 1 | Saneamiento en Access | pendiente |
| 2 | Unificar el acceso a datos | pendiente |
| 3 | Ensayo con `Lanzadera_Datos` (antes: ensayo con una isla) | revisada |
| 4 | `Lanzadera_Datos` y núcleos compartidos | pendiente |
| 5 | Expedientes a web | pendiente |

### Fase 0. Inventario

**Estado:** pendiente.

**Objetivo.** Conocer, en solo lectura, el grafo completo de dependencias entre backends, el estado de cada tabla y los patrones de VBA que no sobrevivirían a PostgreSQL.

**Entregables.**

- **Por backend:** las tablas vinculadas (`MSysObjects`, tipo 6), para construir el grafo de dependencias solo desde los backends (gracias a R-1).
- **Por tabla:** clave primaria, tipos problemáticos (multivalor, datos adjuntos, objeto OLE), reglas de validación, valores por defecto, macros de datos, relaciones, y datos huérfanos o duplicados.
- **Por front:** patrones de riesgo en VBA: generadores `Max + 1` (como `DameID`), `LIKE '*…'`, literales de fecha con `#`, `Date()` en SQL, `Seek` y `dbOpenTable`, recordsets sin `dbSeeChanges` y la comparación insensible a mayúsculas de `Option Compare Database`.
- **Clasificación de cada tabla** según el [protocolo de clasificación](#clasificación-de-tablas).
- **Coste de cada cambio estructural:** número de referencias en código por tabla y por columna, sumando todos los fronts. Es el criterio de coste de la fase 1.
- **Verificación de R-1 y R-2**, con sus excepciones listadas.

**Criterio de terminado.** Todo backend dentro de alcance tiene su lista de vínculos; toda tabla tiene clasificación y ficha; todo front tiene su inventario de patrones de riesgo; R-1 y R-2 constan como confirmadas o con sus excepciones.

### Fase 1. Saneamiento en Access

**Estado:** pendiente.

**Objetivo.** Mejorar la estructura de datos en el legacy, backend a backend y antes de migrar cada uno, para que la web no empiece sobre tablas con defectos inaceptables (decisión 11). Los cambios se hacen sobre `.accdb` y son reversibles.

**Reglas.**

- Toda mejora estructural se hace en el backend Access, con sus fronts adaptados, antes de migrar ese backend.
- Durante la coexistencia, los cambios de esquema son solo aditivos: columnas nuevas que admiten nulos y tablas nuevas. Nunca cambian el significado de lo existente.

**Clasificación de cada hallazgo de la fase 0.** El coste se mide en la fase 0 (referencias en código por tabla y columna en todos los fronts).

| Grupo | Qué incluye |
|---|---|
| (a) Imprescindible antes de migrar | Claves primarias; sustituir los generadores `Max + 1` por autonumérico o identidad; columnas de clave foránea con el mismo tipo que la clave referenciada; integridad referencial con limpieza previa de huérfanos; fecha y hora reales en lugar de fechas guardadas como texto; eliminar campos multivalor, de datos adjuntos y objeto OLE |
| (b) Recomendable si es barato | Eliminar columnas muertas; restricciones `NOT NULL` y de unicidad donde los datos ya cumplen; campos `Sí`/`No` coherentes como booleano donde las referencias en código sean pocas |
| (c) Aplazar a la refactorización final, tras el apagado | Renombrar tablas y columnas (nomenclatura `TbX`); normalización profunda; todo lo que toque muchas cadenas de SQL concatenado |

**Entregables.**

- Clasificación (a), (b) o (c) de cada hallazgo de la fase 0, con su coste medido.
- Cambios del grupo (a) aplicados, y los del grupo (b) que se aprueben.
- Huérfanas tratadas según el [protocolo de cuarentena](#protocolo-de-huérfanas).
- SQL compatible con ambos motores y parametrizado.

**Criterio de terminado.** Para cada backend que se va a migrar: los hallazgos del grupo (a) están resueltos, los del grupo (b) resueltos o descartados con motivo, los del grupo (c) registrados para la refactorización final, y sus aplicaciones siguen funcionando sobre `.accdb`.

### Fase 2. Unificar el acceso a datos

**Estado:** pendiente.

**Objetivo.** Que cada front cambie entre `.accdb` y PostgreSQL solo por configuración. Ese mismo cambio es la reversión.

**Entregables.** Adaptación a ODBC sin DSN de los tres mecanismos de acceso existentes:

| Mecanismo | Dónde |
|---|---|
| `getdb()` con `OpenDatabase` y clave | Lanzadera |
| Revinculación desde `TbTablasAVincular` | Aplicaciones antiguas |
| Resolvedor propio de cada aplicación | Aplicaciones modernas (por ejemplo, `BackendResolver.bas` en Expedientes) |

**Criterio de terminado.** Para cada front, el paso de `.accdb` a PostgreSQL y la vuelta se hacen cambiando solo configuración, demostrado en un entorno que no es producción.

### Fase 3. Ensayo con `Lanzadera_Datos`

**Estado:** revisada (2026-11-04, decisión 12). El objetivo anterior, ensayar con una isla (Brass o Seguridad), queda sustituido; ver [Cambios de rumbo](#cambios-de-rumbo).

**Objetivo.** Ensayar la migración de una copia de `Lanzadera_Datos` en preproducción dentro de Citrix, con la Lanzadera Access, la Lanzadera web y al menos una aplicación Access que vincule sus tablas de identidad (por ejemplo, Expedientes con la unión entre responsables y usuarios).

**Entregables.** Informe del ensayo con: conversión de tipos, tablas vinculadas actualizables (requieren clave primaria), rendimiento de las uniones mixtas `.accdb` + ODBC y reversión. Resultado de las [hipótesis H-1 y H-2](#hipótesis-abiertas).

**Criterio de terminado.** Las dos Lanzaderas funcionan sobre la copia en PostgreSQL en preproducción, la reversión se ha ejecutado al menos una vez y H-1 y H-2 están confirmadas o refutadas con evidencia.

### Fase 4. `Lanzadera_Datos` y núcleos compartidos

**Estado:** pendiente.

**Objetivo.** Trasladar `Lanzadera_Datos` a PostgreSQL. Si la fase 3 refuta que las uniones mixtas sean aceptables (H-2), el traslado incluye en el mismo corte los backends que vinculan sus tablas y los núcleos compartidos (`AGEDYS_DATOS` y Registro).

**Entregables.** Plan del corte conjunto, ensayo en preproducción y resultado de la [Hipótesis H-2](#hipótesis-abiertas).

**Criterio de terminado.** Las dos Lanzaderas y todas las aplicaciones Access operan sobre estos backends en PostgreSQL, y la reversión del corte conjunto se ha ensayado.

### Fase 5. Expedientes a web

**Estado:** pendiente.

**Objetivo.** Primer corte de aplicación: Expedientes pasa a la Lanzadera web y se deshabilita su botón en la Lanzadera Access.

**Entregables.** Los que fija la [coexistencia de Expedientes](../03-aplicaciones/expedientes/prd/coexistencia.md) para el corte.

**Criterio de terminado.** Expedientes opera en la web sobre los datos compartidos y la reversión (rehabilitar el botón de Access) se ha probado.

**Arquitectura de datos:** la web opera sobre las tablas con forma legacy (decisión 10; ver [P-EXP-3](#p-exp-3)).

## Clasificación de tablas

Cada tabla del inventario recibe una de estas clases, según R-1 y R-2:

| Clase | Condición | Dueño |
|---|---|---|
| Propia | La usa solo la aplicación de su backend | La aplicación del backend |
| Propia y compartida | La usa la aplicación de su backend y otras la vinculan | La aplicación del backend |
| Mal ubicada | No la usa la aplicación de su backend y la usa una sola aplicación B | B |
| Mal ubicada con dueño por decidir | No la usa la aplicación de su backend y la usan varias | La que escribe; si escriben varias, decide el propietario |
| Huérfana | No la usa ninguna aplicación | Ver el [protocolo de huérfanas](#protocolo-de-huérfanas) |

Las tablas mal ubicadas **no se mueven físicamente en el legacy**. Se registra su dueño real y, en la migración, la tabla se crea directamente en el esquema PostgreSQL de ese dueño.

### Protocolo de huérfanas

El propietario borra en el legacy; el agente aporta la evidencia.

1. **Evidencia de no uso.** Cruce de los vínculos de todos los backends, búsqueda en el `src` del front propietario (CodeGraph o búsqueda de texto, incluido el SQL dinámico) y fecha del último cambio de datos si existe. Aviso expreso de dos riesgos: consumidores fuera de Access (Excel, informes, scripts programados) y tablas de uso anual.
2. **Cuarentena.** Copia de seguridad y renombrado a `zzBAJA_<tabla>`, o traslado a un `.accdb` de archivo.
3. **Borrado definitivo.** Tras un periodo sin incidencias (por ejemplo, un cierre mensual completo) y antes de migrar ese backend.

## Hipótesis abiertas

| Id | Hipótesis | Se ensaya en | Estado |
|---|---|---|---|
| H-1 | Cómo resuelve Access la cadena de vínculos front → vínculo del backend → ODBC: si el front necesita revincularse o no. | Fase 3 | pendiente |
| H-2 | Las uniones entre tablas deben estar en el mismo motor: una unión mixta `.accdb` + ODBC se resuelve en local dentro de Citrix, con lentitud. Por eso los backends cuyas tablas se unen entre sí se trasladan juntos, y `Lanzadera_Datos` (identidad) debe estar en PostgreSQL desde el principio, porque la comparten las dos Lanzaderas y todas las aplicaciones Access. | Fase 3 (decide el alcance de la fase 4) | pendiente |

## P-EXP-3

La pregunta era si la web trabaja sobre las tablas con forma legacy en PostgreSQL o sobre un esquema propio con sincronización. Sus opciones están analizadas en la [coexistencia de Expedientes](../03-aplicaciones/expedientes/prd/coexistencia.md), apartados 11 y 12.

El propietario la resolvió el 2026-11-04 con la decisión 10: tablas con forma legacy compartidas durante la coexistencia y refactorización a un esquema moderno tras apagar el legacy.

Promovida a [D92](../decisions/2026-11-04-d92-tablas-con-forma-legacy-compartidas.md) (issue #1547), que es la fuente vinculante.

## Criterio de clasificación de issues

Las issues abiertas se redactaron bajo el supuesto de corte único. Se propone:

- **Una etiqueta para la nueva estrategia**, por ejemplo `estrategia:coexistencia`. Es una propuesta; la etiqueta no está creada.
- **Tres grupos:**

| Grupo | Significado |
|---|---|
| Vigente en ambos modelos | El trabajo vale igual con corte único y con coexistencia |
| Requiere ajuste | El trabajo sigue siendo necesario, pero su alcance cambia con la coexistencia |
| Obsoleta por corte único | El trabajo solo tenía sentido con corte único |

**Efecto de la decisión 10 ([P-EXP-3](#p-exp-3)).**

- Las issues de adaptadores y DDL de PostgreSQL, como #803, #804 y #813, y las migraciones del esquema propio de identidad de la Lanzadera web (`0001` y `0008`) caen en **requiere ajuste**: deben apuntar a las tablas con forma legacy.
- El dominio, los casos de uso, la interfaz y la capa HTTP siguen vigentes gracias a la arquitectura hexagonal: el cambio queda en los adaptadores.

Este documento no clasifica más issues concretas; la reclasificación se hace en issues aparte.

## Cambios de rumbo

Registro fechado. Cada entrada indica qué se creía, qué se encontró, la evidencia y qué cambia.

### 2026-11-04. De corte único a coexistencia progresiva

- **Qué se creía.** La web entraría en servicio cuando todas las aplicaciones legacy estuvieran migradas y probadas.
- **Qué se encontró.** El propietario decide sustituir las aplicaciones de una en una, con las dos Lanzaderas conviviendo sobre un único PostgreSQL.
- **Evidencia.** Decisiones del propietario del 2026-11-04 (ver [Decisiones del propietario](#decisiones-del-propietario-2026-11-04)); recuentos de `TbAplicaciones` y `TbTablasAVincular` (ver [Evidencia de partida](#evidencia-de-partida)).
- **Qué cambia.** Se abre esta épica con las fases 0 a 5, y las issues abiertas deben clasificarse según el [criterio de clasificación](#criterio-de-clasificación-de-issues).

### 2026-11-04. P-EXP-3: de dos familias de esquemas a tablas compartidas

- **Qué se creía.** La dirección preferida del propietario era la opción C: esquemas modernos para la web y un esquema legacy para Access, sincronizados por triggers en ambos sentidos.
- **Qué se encontró.** Con coexistencia progresiva y reversión por aplicación, una sola fuente de verdad garantiza que el legacy siga funcionando y que la reversión sea posible.
- **Evidencia.** [Coexistencia de Expedientes](../03-aplicaciones/expedientes/prd/coexistencia.md), apartado 12 (costes de las opciones A, B y C); decisión 10 del propietario, que confirmó de forma expresa el 2026-11-04 que la opción A sustituye a la C.
- **Qué cambia.** La web usa las tablas con forma legacy (opción A) hasta apagar el legacy; las issues de adaptadores y DDL pasan a «requiere ajuste».

### 2026-11-04. La Lanzadera, primera en coexistir

- **Qué se creía.** El primer ensayo se haría con una isla (Brass o Seguridad), y `Lanzadera_Datos` se trasladaría después, en un corte conjunto con los núcleos compartidos.
- **Qué se encontró.** La Lanzadera alberga el control de permisos del resto de herramientas y tiene a todos los usuarios y sus permisos: ninguna aplicación web puede coexistir sin compartir esa identidad. Una isla valida la conversión de tipos, pero no las uniones mixtas ni la cadena de vínculos, que son lo que decide el orden del resto.
- **Evidencia.** Decisión 12 del propietario; `Lanzadera_Datos` la vinculan Expedientes, HPS, Condor, No Conformidades, Gestión de Riesgos, la familia AGEDYS y Registro (ver [Evidencia de partida](#evidencia-de-partida) y la [coexistencia de Lanzadera](../03-aplicaciones/lanzadera/prd/coexistencia.md), apartado 6).
- **Qué cambia.** La fase 3 pasa a `revisada`: el ensayo se hace con una copia de `Lanzadera_Datos` en preproducción y valida H-1 y H-2. El resultado de H-2 decide si la fase 4 traslada `Lanzadera_Datos` sola o junto con los backends que la vinculan.
