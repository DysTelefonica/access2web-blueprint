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
| 1 | La estrategia pasa de corte único (la web solo entra en servicio cuando todas las aplicaciones legacy estén migradas y probadas) a coexistencia progresiva (patrón strangler fig). Es una propuesta en exploración: ver la decisión 18. |
| 2 | La Lanzadera Access y todas las aplicaciones Access se alojan en un servidor de aplicaciones Windows publicado por Citrix. La plataforma web, en la misma máquina o en una conectada a ella. Concretada por la decisión 20. |
| 3 | Todas las tablas de todos los backends `.accdb` legacy pasan a un único PostgreSQL, con esquemas por aplicación ([D14](../decisions/2026-08-04-d14-postgresql-compartido-con-esquemas-por.md)). Los backends legacy las vinculan por ODBC. |
| 4 | La Lanzadera Access sigue abriendo las aplicaciones Access aún no apagadas; la Lanzadera web contiene solo las aplicaciones web nuevas. Ninguna lanza aplicaciones de la otra. Ambas conviven hasta el final. |
| 5 | El corte es por aplicación: se deshabilita su botón en la Lanzadera Access y se habilita en la web. La reversión consiste en volver a habilitar el botón de Access; los datos son compartidos. |
| 6 | Quedan fuera de alcance Sirdee (pertenece a otro departamento y no figura en la Lanzadera) y AGEDO (no se usa). |
| 7 | Regla R-1 del legacy (ver [Reglas del legacy](#reglas-del-legacy)). |
| 8 | Regla R-2 del legacy (ver [Reglas del legacy](#reglas-del-legacy)). |
| 9 | Este documento es vivo, con estados de fase, hipótesis marcadas y registro de cambios de rumbo. |
| 10 | Durante la coexistencia, la web trabaja directamente sobre las tablas con forma legacy en PostgreSQL: una sola fuente de verdad, para garantizar que el legacy sigue funcionando y que la reversión es posible. No hay esquema web propio con sincronización. Cuando el legacy esté apagado por completo, el modelo se refactoriza a un esquema moderno para dejar de arrastrar la deuda técnica de las tablas legacy. Resuelve P-EXP-3 (ver [P-EXP-3](#p-exp-3)). |
| 11 | La estructura de datos se mejora en el legacy antes de que cada backend pase a PostgreSQL, donde no cueste demasiado trabajo, para que la web nunca empiece sobre tablas con defectos inaceptables. Toda mejora estructural se hace en el backend Access (con sus fronts adaptados) antes de migrar ese backend. Durante la coexistencia, los cambios de esquema son solo aditivos (columnas nuevas que admiten nulos, tablas nuevas) y nunca cambian el significado de lo existente. Ver la [fase 1](#fase-1-saneamiento-en-access). |
| 13 | AGEDYS Técnico y AGEDYS ECO se sustituyen por una sola aplicación web, **APGEDYS**, con la misma aplicación mostrando vista de técnico o de gestor económico según el usuario que entre. El alcance de la migración son las once aplicaciones del catálogo de la Lanzadera que el propietario marca: las nueve del grupo general más Seguridad y HPS. |
| 14 | AGEDO y Tareas quedan fuera del alcance pero **siguen en uso en Access** de forma indefinida. Son consumidores permanentes de tablas que sí migran: cuando el backend dueño de una de esas tablas pase a PostgreSQL, sus vínculos se reapuntan por ODBC como los de cualquier otro front Access. Las tablas de AGEDO que consumen aplicaciones del alcance tienen que llegar a PostgreSQL para que la web pueda leerlas. Ver [backends fuera del alcance](fase-0/backends-fuera-de-alcance.md#agedo-y-tareas-siguen-en-uso). |
| 15 | La fila 21 de `TbAplicaciones` (AGEDYS, `AGEDYS.accdb`, sin icono en la lanzadera) es un residuo que no se usa. No entra en la migración ni en el reparto de `AGEDYS_DATOS`. Ver el [catálogo de aplicaciones](fase-0/catalogo-aplicaciones.md). |
| 16 | `Correos_datos` es un backend más, no un servicio de la plataforma. Su tabla `TbCorreosEnviados` es la cola de salida: las aplicaciones la cargan y un proceso del servidor la recorre cada cinco minutos para enviar los correos con el SMTP de la empresa. Ese proceso es un consumidor más que hay que reapuntar cuando la cola pase a PostgreSQL. |
| 17 | APGEDYS tiene un único backend común. Las 55 tablas de `AGEDYS_DATOS` que no usa ninguno de los dos fronts pasan por el [protocolo de huérfanas](#protocolo-de-huérfanas) antes de retirarse. |
| 18 | La coexistencia es una prueba de viabilidad en exploración, no la estrategia aprobada. Dirección mantiene el plan de corte único: cuando la última aplicación web esté terminada, se apaga todo el legacy y se enciende lo nuevo, sin compartir nada. La propuesta se presenta a dirección como alternativa. |
| 19 | Tras aplicar en los backends legacy los accionables de la fase 0 (fase 1), se ensaya en Citrix, en preproducción, el paso de **todos** los backends a PostgreSQL, sin ningún backend Access. Superado el ensayo, todos los datos pasan a PostgreSQL de una vez en producción. Ese traslado es **irreversible**. Desde entonces solo hay PostgreSQL: lo no migrado funciona con la Lanzadera Access y lo migrado con la web, las dos contra PostgreSQL. La única reversión que queda es la del botón de cada aplicación. Sustituye el ensayo solo con `Lanzadera_Datos` de la decisión 12 y el traslado por tandas de las fases 3 y 4. |
| 20 | Mientras convivan, los usuarios entran a todo por Citrix: la Lanzadera Access y sus aplicaciones se ejecutan en el servidor de aplicaciones Windows, y la web nueva se abre con un navegador publicado en Citrix. La web nueva y PostgreSQL común están en otra máquina de la misma red. Al terminar la convivencia, la web, ya en contenedores, se lleva a OpenShift, como cadete. |
| 21 | `TbExpedientesAnualidades` vive de forma única en el backend de Expedientes. Hoy hay dos tablas propias con ese nombre, en `AGEDYS_DATOS` y en `Expedientes_datos`, que se vinculan mutuamente y las dos tienen la misma estructura (clave `IDAnualidad`), pero divergen: 223 filas en AGEDYS y 173 en Expedientes, ningún `IDAnualidad` común, 157 pares expediente y año presentes en las dos sin ninguna fila idéntica (por ejemplo, IPSI del 8 % en AGEDYS frente al 4 % en Expedientes), 66 pares solo en AGEDYS y 16 solo en Expedientes. Antes de retirar la de AGEDYS hay que conciliarlas; ver el [catálogo de conflictos](fase-0/catalogo-accionables-conflictos.md#qué-necesita-decidir-el-propietario). |
| 22 | `TbCorreosEnviados.IDCorreo` pasa a autonumérico. `IDCorreo` no es autonumérico: lo asigna cada aplicación al escribir en la cola (Condor lo calcula con `Max(IDCorreo) + 1` en `CorreoRepositorio.bas:18-24`; también lo asignan No Conformidades, HPS, Lanzadera, Solicitudes HPS, Gestión de Riesgos y Expedientes). En la copia de referencia hay 19 valores repetidos, cada uno en una pareja de correos distintos (otro asunto, remitente o aplicación): son colisiones de escrituras simultáneas, no filas duplicadas, y no se puede borrar ninguna. En el saneamiento se renumera el segundo correo de cada pareja por encima del máximo actual, se declara la clave primaria y todas las aplicaciones dejan de calcular el identificador. Antes de renumerar se comprueba que ninguna tabla guarda esos 19 valores. |
| 12 | La Lanzadera es la primera aplicación en coexistir, porque alberga el control de permisos del resto de herramientas y tiene a todos los usuarios y sus permisos (`Lanzadera_Datos`). El ensayo de la fase 3 se hace con una copia de `Lanzadera_Datos` en preproducción, no con una isla. Ver [Cambios de rumbo](#cambios-de-rumbo). El alcance del ensayo queda sustituido por la decisión 19. |

## Modelo de coexistencia

```text
Puesto del usuario: solo el cliente Citrix
        │
        ▼
Servidor de aplicaciones Windows (Citrix)          Máquina web (misma red)
  Navegador publicado ─────────────────────────────► Lanzadera web
  Lanzadera Access                                     └── aplicaciones web nuevas
    └── aplicaciones Access no apagadas                        │
          │ (tablas vinculadas por ODBC)                       │
          └──────────────────────────────────────────► PostgreSQL común
                                                       (un esquema por aplicación, D14)
```

Cada aplicación está en un solo lado a la vez: su botón está habilitado en una Lanzadera y deshabilitado en la otra. Los datos no se duplican ni se sincronizan entre copias: la web usa las mismas tablas con forma legacy (decisión 10). Tras el traslado único de la fase 4 no queda ningún backend `.accdb` (decisión 19).

## Alcance

El alcance lo fija el catálogo de la Lanzadera Access, con la decisión del propietario sobre qué entra de cada grupo (2026-11-04).

| Grupo del catálogo | Aplicaciones | En el alcance |
|---|---|---|
| Fuera de la Oficina CARU | AGEDYS Técnico, AGEDYS ECO, Registro, Gestión de Riesgos, No Conformidades, Brass, Solicitudes HPS, Expedientes, Condor | las nueve |
| Solo Oficina | Seguridad, SICA, HPS, Contratos, AGEDO | Seguridad y HPS |
| En Pruebas | pendientes de confirmar | según la decisión del propietario |

- **Dentro:** las once aplicaciones marcadas. AGEDYS Técnico y AGEDYS ECO se sustituyen por **una sola aplicación web, APGEDYS**, que muestra vista de técnico o de gestor económico según el usuario que entre (decisión 13).
- **Fuera:** Sirdee (decisión 6), AGEDO (decisión 6, confirmada por el propietario el 2026-11-05), SICA y Contratos (decisión del propietario de 2026-11-04) y Tareas (decisión del propietario de 2026-11-05; su backend `Tareas_datos1` no tiene fila en el catálogo de la lanzadera).
- **Backends de aplicaciones fuera del alcance:** sus tablas las consumen aplicaciones dentro del alcance, de modo que se inventarían como dependencias en la fase 0:  Correos, Registro de salida, SICA, Control de cambios, AGEDO y Tareas. No se migran como aplicaciones, pero sí hay que saber qué sobrevive y quién es su dueño.
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
| 6 | Brass | 75 | Solo la propia aplicación (isla por vínculos; abre `Lanzadera_Datos` directamente, ver la [fase 0 de la Lanzadera](fase-0/lanzadera-datos.md#verificación-de-r-1-y-r-2)) |
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
| 0 | Inventario | en curso |
| 1 | Saneamiento en Access | pendiente |
| 2 | Unificar el acceso a datos | pendiente |
| 3 | Ensayo en Citrix con todos los backends (antes: con `Lanzadera_Datos`; antes aún: con una isla) | revisada |
| 4 | Traslado único de todos los datos a PostgreSQL | pendiente |
| 5 | Expedientes a web | pendiente |

### Fase 0. Inventario

**Estado:** en curso. Hecho: [`Lanzadera_Datos`](fase-0/lanzadera-datos.md) (#1560) y [`Brass`](fase-0/brass.md) (#1574), con sus propuestas de acción catalogadas, pendientes de revisión transversal. Pendiente: el resto de backends. Completar un inventario individual no habilita cambios en el legacy.

**Objetivo.** Conocer, en solo lectura, el grafo completo de dependencias entre backends, el estado de cada tabla y los patrones de VBA que no sobrevivirían a PostgreSQL.

**Base de análisis.** Si una aplicación tiene una rama RC1 con UAT pendiente, el descubrimiento se basa en esa RC1, no en `main`. Cada ficha fija rama y SHA y distingue código candidato de versión desplegada. Los inventarios previos hechos solo sobre `main` deben reconciliarse con la RC1 aplicable; las divergencias se registran, no se mezclan silenciosamente. Para el ERD se comprueba la correspondencia del binario con la base analizada.

**Entregables.**

- **Por aplicativo:** obtener con dysflow el ERD del **front y del backend**, identificando tablas propias, tablas vinculadas y sus orígenes sin consultar tablas de sistema como `MSysObjects`. Cruzar ambos ERD con el código mediante CodeGraph para construir el grafo de dependencias. No basta con el backend: Brass tiene vínculos también en el front y R-1 presenta excepciones demostradas.
- **Por tabla:** clave primaria, tipos problemáticos (multivalor, datos adjuntos, objeto OLE), reglas de validación, valores por defecto, macros de datos, relaciones, y datos huérfanos o duplicados.
- **Por front:** patrones de riesgo en VBA: generadores `Max + 1` (como `DameID`), `LIKE '*…'`, literales de fecha con `#`, `Date()` en SQL, `Seek` y `dbOpenTable`, recordsets sin `dbSeeChanges` y la comparación insensible a mayúsculas de `Option Compare Database`.
- **Clasificación de cada tabla** según el [protocolo de clasificación](#clasificación-de-tablas).
- **Coste de cada cambio estructural:** número de referencias en código por tabla y por columna, sumando todos los fronts. Es el criterio de coste de la fase 1.
- **Verificación de R-1 y R-2**, con sus excepciones listadas.

**Avance.** Inventarios publicados, todos de solo lectura y con las propuestas de acción catalogadas sin aplicar:

- [`Lanzadera_Datos`](fase-0/lanzadera-datos.md) y el [inventario de aperturas directas](fase-0/aperturas-directas-lanzadera.md).
- [`Brass`](fase-0/brass.md), la excepción con controles conectados y vínculos en el front.
- [`HPS`](fase-0/hps.md), [`Condor`](fase-0/condor.md), [`Expedientes`](fase-0/expedientes.md), [`Gestión de Riesgos`](fase-0/gestion-riesgos.md) y [`No Conformidades`](fase-0/no-conformidades.md).
- El [catálogo único de accionables](fase-0/catalogo-accionables.md), con los [grupos b y c](fase-0/catalogo-accionables-b.md) y los [conflictos, bloqueos y decisiones](fase-0/catalogo-accionables-conflictos.md): las 159 propuestas de las once aplicaciones del alcance y de los siete backends de dependencias, ordenadas por decisión.
- El [grafo de dependencias entre backends](fase-0/grafo-dependencias.md) en su versión transversal: 17 backends con un único método de recuento.
- El [catálogo de aplicaciones de la lanzadera](fase-0/catalogo-aplicaciones.md): las 22 filas de `TbAplicaciones`, que confirman el alcance y destapan la tercera aplicación AGEDYS, el RegistroAntiguo y el backend de Condor PRUEBA.
- La [copia de producción de referencia](fase-0/copia-referencia.md): la foto fija de los 18 backends, con su SHA-256, contra la que se hacen todas las comprobaciones de datos.
- Las [comprobaciones de datos del grupo (a)](fase-0/comprobaciones-grupo-a.md): las 36 propuestas imprescindibles que admiten medición, con su veredicto contra la copia de referencia.
- El [anexo de controles conectados](fase-0/controles-conectados.md): el SQL literal de cada control con origen de datos y el catálogo de consultas que sobrevivirán a PostgreSQL. La medición refuta la regla «salvo Brass, los formularios son desconectados»: Registros tiene 37 controles conectados y HPS 3 formularios unidos a recordset.

R-1 se confirma en los backends y se refuta en los fronts, con las excepciones medidas. La base de análisis de cada aplicación es su RC1 pendiente de UAT cuando existe; HPS queda en `main` porque su UAT ya está aprobada y liberada.

**Criterio de terminado.** Todo backend dentro de alcance tiene su lista de vínculos; toda tabla tiene clasificación y ficha; todo front tiene su inventario de patrones de riesgo; R-1 y R-2 constan como confirmadas o con sus excepciones.

### Fase 1. Saneamiento en Access

**Estado:** pendiente.

**Objetivo.** Mejorar la estructura de datos en el legacy, backend a backend y antes de migrar cada uno, para que la web no empiece sobre tablas con defectos inaceptables (decisión 11). Los cambios se hacen sobre `.accdb` y son reversibles.

**Condición de entrada.** La fase 1 no comienza hasta terminar el descubrimiento de **todas** las herramientas dentro de alcance, incluidos sus fronts y backends y los controles conectados de cada aplicación, medidos uno por uno. Antes se reconcilian las dependencias y se revisa el catálogo conjunto de accionables con el propietario. Durante la fase 0 se catalogan propuestas, pero no se aplican cambios en el legacy: tampoco renombrados de cuarentena, borrados, movimientos, nuevas claves ni restricciones.

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

**Objetivo.** Que cada front cambie entre `.accdb` y PostgreSQL solo por configuración. Ese cambio permite el ensayo de la fase 3; tras el traslado único de la fase 4 ya no hay vuelta a `.accdb` (decisión 19).

**Entregables.** Adaptación a ODBC sin DSN de los cuatro mecanismos de acceso existentes:

| Mecanismo | Dónde |
|---|---|
| `getdb()` con `OpenDatabase` y clave | Lanzadera |
| Revinculación desde `TbTablasAVincular` | Aplicaciones antiguas. No demostrada: la tabla no tiene ninguna referencia en el `src` de los 13 fronts (fase 0 de la Lanzadera) |
| Resolvedor propio de cada aplicación | Aplicaciones modernas (por ejemplo, `BackendResolver.bas` en Expedientes) |
| Apertura directa de `Lanzadera_Datos` con ruta escrita en el código | 11 fronts, incluido Condor (todos salvo No Conformidades en `main`); se adapta en el mismo corte de la Lanzadera. Inventario y alternativas: [aperturas directas de `Lanzadera_Datos`](fase-0/aperturas-directas-lanzadera.md) (#1570) |

**Criterio de terminado.** Para cada front, el paso de `.accdb` a PostgreSQL y la vuelta se hacen cambiando solo configuración, demostrado en un entorno que no es producción.

### Fase 3. Ensayo en Citrix con todos los backends

**Estado:** revisada (2026-11-05, decisión 19). Sustituye al ensayo solo con `Lanzadera_Datos` (decisión 12), que a su vez sustituyó al ensayo con una isla; ver [Cambios de rumbo](#cambios-de-rumbo).

**Objetivo.** Ensayar en preproducción, dentro de Citrix, el paso de **todos** los backends ya saneados a PostgreSQL, sin ningún backend Access, con la Lanzadera Access y todas las aplicaciones Access funcionando contra él.

**Entregables.** Informe del ensayo con: conversión de tipos, tablas vinculadas actualizables (requieren clave primaria), funcionamiento de cada aplicación, rendimiento en Citrix y resultado de las [hipótesis H-1, H-3 y H-4](#hipótesis-abiertas).

**Criterio de terminado.** Todas las aplicaciones Access funcionan en preproducción sobre PostgreSQL con un rendimiento aceptable, y H-1, H-3 y H-4 están confirmadas o refutadas con evidencia.

### Fase 4. Traslado único de todos los datos

**Estado:** pendiente.

**Objetivo.** Trasladar en producción, de una vez, todos los backends a PostgreSQL. Los ficheros `.accdb` de datos dejan de usarse. Es el **único paso sin vuelta atrás** del plan (decisión 19).

**Entregables.** Plan del traslado, repetición del ensayo de la fase 3 con una copia reciente de producción y reapuntado de los consumidores que no migran (AGEDO, Tareas y el proceso de envío de correos del servidor, decisiones 14 y 16). Los despliegues de prueba (`CONDOR Prueba`, `EXPEDIENTES PRUEBA`, `HPS PRUEBA`, `No Conformidades Prueba`, `00_datos_prueba`) comparten la unidad de red con producción: el plan debe decir qué se hace con cada uno.

**Criterio de terminado.** Todas las aplicaciones Access trabajan en producción sobre PostgreSQL y no queda ningún backend `.accdb` en uso.

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
| H-2 | Las uniones entre tablas deben estar en el mismo motor: una unión mixta `.accdb` + ODBC se resuelve en local dentro de Citrix, con lentitud. Por eso los backends cuyas tablas se unen entre sí se trasladan juntos, y `Lanzadera_Datos` (identidad) debe estar en PostgreSQL desde el principio, porque la comparten las dos Lanzaderas y todas las aplicaciones Access. | — | sin objeto: con el traslado único (decisión 19) no hay uniones mixtas |
| H-3 | La copia del front por usuario que ya hace la Lanzadera (`%APPDATA%\Aplicaciones DYSN\`, `Aplicacion.cls:305` y `:668`) funciona en Citrix sin cambios. Falta confirmar la sincronización de perfiles itinerantes (copiar `.accde` a `%APPDATA%` alarga los inicios de sesión; valorar `%LOCALAPPDATA%`), la ruta de arranque (hoy un acceso directo a una carpeta sincronizada en el perfil), las instalaciones del servidor (Access o Access Runtime con la arquitectura de los fronts y el driver ODBC de PostgreSQL) y cuántas sesiones simultáneas soporta. | Fase 3 | pendiente |
| H-4 | Los documentos adjuntos viven como ficheros en la carpeta `ANEXOS` de cada aplicación en la unidad de red. La web tiene que acceder a esa carpeta o hay que decidir dónde guarda los documentos. | Fase 3 | pendiente |

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

### 2026-11-05. Traslado único de los datos y acceso por Citrix

- **Qué se creía.** Los backends pasaban a PostgreSQL por tandas: primero `Lanzadera_Datos` y los núcleos compartidos, y los demás a medida que sus aplicaciones migraban, con backends `.accdb` conviviendo con PostgreSQL. La coexistencia figuraba como estrategia adoptada.
- **Qué cambia.** Tras el saneamiento, se ensaya en Citrix el paso de todos los backends y después se trasladan todos de una vez; es el único paso irreversible (decisión 19). Los usuarios entran a todo por Citrix mientras convivan, y la web irá después a OpenShift (decisión 20). Ante dirección, la coexistencia es una propuesta en exploración frente al corte único (decisión 18).
- **Motivo.** El propietario lo aclara el 2026-11-05 al preparar la presentación para dirección. Con un solo motor desaparecen las uniones mixtas (H-2) y la convivencia de dos formatos de datos.

### 2026-11-04. Alcance por catálogo de la lanzadera y fusión de AGEDYS

- **Qué se creía.** El alcance era el catálogo `TbAplicaciones` completo, con Sirdee y AGEDO excluidos.
- **Qué cambia.** El propietario delimita el alcance con el catálogo visible de la Lanzadera: entran las nueve aplicaciones del grupo general más Seguridad y HPS; quedan fuera SICA y Contratos. AGEDYS Técnico y AGEDYS ECO se sustituyen por una sola aplicación web, APGEDYS, con vista según rol (decisión 13).
- **Motivo.** El catálogo es lo que el usuario ve y abre; migrar una aplicación que no está en el catálogo no tiene destinatario. La fusión de AGEDYS evita mantener dos aplicaciones web para el mismo dato técnico y económico.

### 2026-11-04. Primero descubrimiento global; después saneamiento

- **Qué se creía.** Los accionables podían aplicarse backend a backend al terminar su inventario.
- **Qué cambia.** Por decisión del propietario (#1576), se catalogan durante la fase 0, pero no se empieza a aplicar ninguno hasta terminar el descubrimiento de todas las herramientas y reconciliar las dependencias entre fronts y backends. La fase 1 tiene esa condición de entrada global.
- **Motivo.** Evitar que una sorpresa en otro front invalide una retirada o modificación ya aplicada, especialmente con la excepción de Brass.

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

### 2026-11-04. R-1 refutada para `Lanzadera_Datos`

- **Qué se creía.** Cada front consulta solo su propio backend (R-1), y Brass era una isla.
- **Qué se encontró.** 10 fronts (y probablemente Condor) abren `Lanzadera_Datos.accdb` directamente con la ruta escrita en el código, incluido Brass; así escriben las aperturas y las conexiones. La revinculación desde `TbTablasAVincular` no aparece en ningún `src`.
- **Evidencia.** [Fase 0 de la Lanzadera](fase-0/lanzadera-datos.md#verificación-de-r-1-y-r-2), con `ruta:línea` por front; revisión con CodeGraph-VBA (#1567).
- **Qué cambia.** La fase 2 suma un cuarto mecanismo de acceso, que debe adaptarse en el mismo corte que `Lanzadera_Datos`; Brass deja de ser isla; R-1 se verifica backend a backend en lugar de darse por cierta.

### 2026-11-04. Aperturas directas de `Lanzadera_Datos`: alcance corregido

- **Qué se creía.** 10 fronts abrían `Lanzadera_Datos.accdb` directamente, Condor probablemente, y por esa conexión escribían las aperturas y las conexiones (entrada anterior).
- **Qué se encontró.** Condor también la abre: son 11 fronts. Todos leen las tablas de identidad; solo cinco tienen código que escribe `TbAplicacionesAperturas`, alcanzable solo en Contratos y Registros. Ningún front escribe `TbConexionesRegistro`. No Conformidades añade en su rama `release/2026-009.rc1` una apertura de solo lectura.
- **Evidencia.** [Aperturas directas de `Lanzadera_Datos`](fase-0/aperturas-directas-lanzadera.md), con SHA y `ruta:línea` por front (#1570).
- **Qué cambia.** El cuarto mecanismo de la fase 2 cubre 11 fronts, con 81 llamadas en 71 procedimientos; la adaptación queda pendiente de decisión entre vínculo del propio backend y conexión ODBC centralizada.
