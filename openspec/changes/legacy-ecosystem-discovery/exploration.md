# Exploración correctiva: descubrimiento del ecosistema legacy

## Resultado y límite de esta pasada

Esta segunda exploración corrige el alcance anterior y establece una **topología general basada en evidencias**, no un inventario completo de capacidades. El alcance exacto contiene únicamente: **Lanzadera, Gestion_Riesgos, No_Conformidades, Condor, HPS_Solicitudes, HPS, Brass y Expedientes**.

**APAP y APAP_WEB quedan explícitamente fuera de alcance.** No se utilizan como nodos del ecosistema ni como evidencia, salvo esta exclusión.

No se han creado propuesta, especificación, diseño ni tareas. Antes de avanzar a esas fases debe completarse la Fase 1 por lotes y obtenerse la aprobación del usuario tras cada aplicación o lote lógico.

## Hechos de entrada confirmados por el usuario

Estos hechos son autoridad de alcance y orientación, pero no sustituyen la comprobación técnica de cada aplicación:

1. **Lanzadera es el origen de los usuarios y los permisos.**
2. **Expedientes crea el expediente del que dependen casi todas las aplicaciones incluidas.**

## Jerarquía de fuentes aplicada

1. `C:\00repos\documentacion` — documentación y especificaciones existentes.
2. Checkout `00_main` de cada aplicación bajo `C:\00repos\codigo` — código, binarios y configuración local.
3. Inspección Dysflow en vivo, únicamente de lectura.
4. Engram — contexto histórico; nunca se usa para elevar una hipótesis a hecho actual.

## Mapeo de aplicaciones, documentación y código

| Aplicación del alcance | Documentación localizada | Checkout principal localizado | Frontend/backend observables | Situación y resolución pendiente |
|---|---|---|---|---|
| **Lanzadera** | `C:\00repos\documentacion\OPENSPEC\00_LANZADERA` | `C:\00repos\codigo\00_LANZADERA\00_main` | `Lanzadera.accdb` / `Lanzadera_Datos.accdb` | Mapeo resuelto. La documentación funcional detallada debe localizarse dentro de `openspec/`. |
| **Gestion_Riesgos** | `C:\00repos\documentacion\OPENSPEC\00_GESTION_RIESGOS` | `C:\00repos\codigo\00_GESTION_RIESGOS\00_main` | `Gestion_Riesgos.accdb` / `Gestion_Riesgos_Datos.accdb` | Mapeo resuelto. Existe `docs/DISCOVERY_MAP.md` y ERD en el checkout principal. |
| **No_Conformidades** | `C:\00repos\documentacion\OPENSPEC\00_No_Conformidades` | `C:\00repos\codigo\00_NO_CONFORMIDADES\00_main` | `NoConformidades.accdb` / `NoConformidades_Datos.accdb` | Mapeo resuelto. El target Dysflow no pudo inventariarse en esta pasada; requiere diagnóstico de acceso. |
| **Condor** | `C:\00repos\documentacion\OPENSPEC\00_CONDOR` | `C:\00repos\codigo\00_CONDOR\00_main` | `CONDOR.accdb` / backend no presente en el checkout inspeccionado | Mapeo de frontend resuelto. Falta confirmar la ruta/nombre del backend de producción y restaurar `.dysflow/project.json`. |
| **HPS_Solicitudes** | `C:\00repos\documentacion\OPENSPEC\00_HPS_SOLICITUDES` (PRD 01–05, ERD y cambios) | **No localizado** como `00_HPS_SOLICITUDES` bajo `C:\00repos\codigo` | No hay frontend/backend Access localizado para este nombre | Es una discrepancia crítica: la documentación describe un sistema Access, pero no existe checkout principal con ese nombre en el directorio inspeccionado. Resolver si corresponde a un repositorio renombrado, a una línea documental sin binario o a una aplicación integrada en HPS. |
| **HPS** | `C:\00repos\documentacion\OPENSPEC\00_HPS` | `C:\00repos\codigo\00_HPS\00_main` | `HPS.accdb` / `HPST.accdb` | Mapeo resuelto. La documentación disponible incluye specs, UAT y releases. |
| **Brass** | `C:\00repos\documentacion\OPENSPEC\00_BRASS` | `C:\00repos\codigo\00_BRASS\00_main` | `Gestion_Brass_Gestion.accdb` / backend no presente en el checkout inspeccionado | Mapeo de frontend resuelto. Falta confirmar backend y configurar Dysflow para inspección reproducible. |
| **Expedientes** | `C:\00repos\documentacion\OPENSPEC\00_EXPEDIENTES` | `C:\00repos\codigo\00_EXPEDIENTES\00_main` | `Expedientes.accdb` / `Expedientes_datos.accdb` (backend observado fuera de `00_main`, en staging) | Mapeo resuelto con advertencia: el `.dysflow/project.json` del main tiene `accessPath` absoluto y estado `path-mismatch`; no se modifica en esta exploración. |

### Fuentes documentales representativas consultadas

- Lanzadera: `00_LANZADERA` y el esquema `00_BRASS/docs/ERD/Lanzadera_Datos.md` / `00_No_Conformidades/docs/ERD/Estructura_Datos_Lanzadera.md` como evidencia de catálogo, usuarios y permisos.
- Gestion_Riesgos: `00_GESTION_RIESGOS/00_main/docs/DISCOVERY_MAP.md`, ERD y `00_Automatizaciones/docs/scripts/GestionRiesgos_bat.md`.
- No Conformidades: `00_No_Conformidades/docs/ERD/` y cambios OpenSpec existentes.
- Condor: `00_CONDOR/docs/PRD/03_Gestion_Solicitudes.md`.
- HPS_Solicitudes: `00_HPS_SOLICITUDES/docs/PRD/01_Dominio_Solicitudes.md`, `04_Usuarios_HPS_Expedientes.md` y `05_Arquitectura_Tecnica.md`.
- HPS, Brass y Expedientes: sus carpetas documentales canónicas; en Expedientes, además, las rutas de entrada indicadas en `00_EXPEDIENTES/00_main/AGENTS.md`.

## Evidencia actual separada por tipo

### Hechos confirmados por documentación/código estático

- Lanzadera contiene el contrato transversal de aplicación/usuario/permisos documentado en `TbAplicaciones`, `TbUsuariosAplicaciones` y `TbUsuariosAplicacionesPermisos`; `TbTablasAVincular` también modela tablas que deben enlazarse por aplicación.
- Gestion_Riesgos contiene referencias estáticas a `getdbLanzadera`, `getdbExpedientes`, `TbExpedientes1` y clases de usuario/expediente. La automatización documental confirma además dependencia de tablas de tareas/correos y de expedientes.
- HPS mantiene `TbExpedientes`, `IDExpediente`, `TbUsuarios` y entidades HPS; su documentación describe `TbSuministradores` vinculado a `Expedientes_datos.accdb` y `IDSolicitud` en la cadena de solicitudes.
- HPS_Solicitudes documenta `TbSolicitudes.IDExpediente` como FK a `TbExpedientes` y flujos de alta, renovación, traspaso y elevación. Esto prueba una dependencia documental, pero no prueba que exista hoy un binario separado.
- Condor documenta `tbSolicitudes.idExpediente` como FK a `TbExpedientes`, y una integración con `TbNoConformidades`; la fuente indica selección obligatoria de expediente en el alta.
- Expedientes documenta y codifica el agregado de expediente: `TbExpedientes`, entidades, suministradores, responsables, estados y operaciones de alta/persistencia.
- No Conformidades contiene entidades `Expediente` y documentos de proyecto/auditoría; la documentación previa identifica una excepción concreta en la que ciertos flujos conservan `Nemotecnico` y no necesitan consultar `TbExpedientes`. Es una excepción de flujo, no una exención global.
- Brass tiene evidencia documental de Lanzadera y código de usuario/permisos, pero el vínculo actual con `IDExpediente`/`CodExp` no queda demostrado por esta pasada.

### Evidencia estática obtenida con CodeGraph

El repositorio documental no tenía índice CodeGraph. Se comprobó previamente la raíz real y se documentó el fallback a herramientas de lectura; no se inicializó un índice nuevo en esta pasada.

Sí existía índice `.codegraph-vba` en `Gestion_Riesgos`, `HPS` y `Expedientes`. Se usó `codegraph-vba.codegraph_explore` antes de ampliar búsquedas de archivos. Hallazgos relevantes:

- Gestion_Riesgos: `getdbLanzadera` y `getdbExpedientes` tienen callers; `UsuarioAplicacionPermisos` accede a `TbUsuariosAplicacionesPermisos` de Lanzadera.
- HPS: `Constructor.getExpediente` consulta `TbExpedientes`; `Entorno.Expedientes` hidrata la colección de expedientes; `UsuarioAplicacionPermisos` resuelve permisos mediante el backend de Lanzadera.
- Expedientes: el grafo confirma clases `UsuarioAplicacionPermisos`, `Expediente`, operaciones de expediente y el patrón de acceso a datos; la relación runtime concreta con cada aplicación consumidora sigue requiriendo búsqueda por símbolo/tabla en cada consumidor.

En los demás repositorios no se encontró `.codegraph` ni `.codegraph-vba` en el `00_main` comprobado. Por contrato, se utilizó únicamente evidencia documental, inventarios de source ya localizados y lectura dirigida de archivos conocidos; no se presenta esa evidencia como grafo completo.

### Evidencia Dysflow de solo lectura

Antes de cada llamada Dysflow se ejecutó `get_capabilities({cwd:<repo-main>})`. El runtime vivo respondió con `adapterVersion: 2.35.2`, `toolsVisible: 94` y workflows de lectura disponibles. No se ejecutaron imports, exports, syncs, SQL, tests, compilación, limpieza ni mutaciones.

Targets y resultado exacto:

| Target Dysflow | `get_capabilities` | `list_objects` | Resultado adicional |
|---|---|---|---|
| `00_LANZADERA\00_main` (`lanzadera`) | `valid`, frontend y backend resueltos | **Éxito**, 87 elementos resumidos | Backend: 35 tablas; frontend: `TbConfiguracionBackends`. |
| `00_GESTION_RIESGOS\00_main` (`00-gestion-riesgos-main`) | `valid`, frontend y backend resueltos | **Éxito**, 193 elementos resumidos | Backend: 68 tablas; frontend: `TbAuxPriorizacion`. |
| `00_NO_CONFORMIDADES\00_main` | `valid`, frontend y backend resueltos | **Falló** con error de runtime no desglosado en la respuesta normalizada | `list_tables` también falló; no se afirma inventario runtime. |
| `00_CONDOR\00_main` | `missing`: no hay `.dysflow/project.json` | **Éxito** usando rutas explícitas de lectura, 172 elementos resumidos | Backend indicado por AGENTS no está localizado; no se pudo validar su esquema. |
| `00_HPS\00_main` (`00-hps-staging`) | `valid`, frontend/backend resueltos | **Éxito**, 130 elementos resumidos | Backend: 22 tablas; frontend: 12 tablas. |
| `00_BRASS\00_main` | `missing`: no hay `.dysflow/project.json` | **Éxito** usando rutas explícitas de lectura, 220 elementos resumidos | Backend no localizado; `list_tables` no pudo validar backend. |
| `00_EXPEDIENTES\00_main` | `path-mismatch`: `accessPath` absoluto heredado; proyecto no write-ready | **Éxito** usando rutas explícitas de lectura, 158 elementos resumidos | Frontend: `TbAuxEstadosMartina`; backend no estaba resoluble desde el main. |

Los recuentos anteriores son inventarios de `list_objects`, no un recuento funcional de capacidades. No se interpretan los campos internos duplicados de esa respuesta como número definitivo de formularios, módulos o informes hasta que se haga una extracción normalizada por aplicación.

## Topología general basada en evidencia

El modelo de trabajo actual es:

`Usuario → Lanzadera (identidad/permisos) → aplicación autorizada → expediente creado/seleccionado en Expedientes → capacidades propias → informes, documentos, correo o procesos batch`.

La flecha hacia Expedientes es **fuerte pero no universal**. La dependencia debe modelarse como una referencia de negocio a `IDExpediente`, `CodExp`, `Nemotecnico` u otro identificador, diferenciando siempre su semántica y cardinalidad. No debe confundirse una tabla vinculada, una copia local, una FK, un código textual o una consulta de selección.

### Identificadores y contratos transversales a confirmar

| Contrato | Evidencia actual | Estado |
|---|---|---|
| Identidad de usuario | `TbUsuariosAplicaciones`, `tbUsuarios`, `CorreoUsuario`, `UsuarioRed`, `VBA.Command` y clases `UsuarioAplicacionPermisos` según aplicación | Confirmado como familia de patrones; contrato único aún no probado. |
| Permisos | `TbUsuariosAplicacionesPermisos`, `IDAplicacion`, roles por aplicación y `TbAplicaciones` | Confirmado en Lanzadera/documentación y consumidores concretos; falta matriz efectiva de roles. |
| Aplicación | `TbAplicaciones.IDAplicacion`, `NombreCorto`, ejecutable, backend y comando | Confirmado como catálogo de Lanzadera; falta extraer las ocho filas/identidades actuales. |
| Expediente | `IDExpediente`, `CodExp`, `CodExpLargo`, `Nemotecnico`, título y estados | Confirmados como campos distintos; falta contrato de equivalencia por aplicación. |
| Backends | `getdb()`, `getdbLanzadera`, `getdbExpedientes`, tablas vinculadas y configuraciones locales | Confirmado como arquitectura distribuida; rutas por entorno todavía incompletas. |

## Matriz inicial de dependencias entre las ocho aplicaciones

`Fuerte` significa evidencia estática/documental directa; `Probable` requiere confirmación dirigida; `Excepción` aplica solo a un flujo identificado.

| Aplicación | Lanzadera: identidad/permisos | Expedientes: expediente | Enlaces adicionales observados | Excepción o incógnita principal |
|---|---|---|---|---|
| Lanzadera | **Origen confirmado** | No consumidora demostrada | Catálogo `TbAplicaciones`, permisos, tablas a vincular, aperturas y tareas | Debe obtenerse el catálogo actual de las otras siete. |
| Gestion_Riesgos | **Fuerte**: `getdbLanzadera`, permisos | **Fuerte**: `getdbExpedientes`, `TbExpedientes1` | Tareas/correos, proyectos, riesgos, suministradores y automatización batch | Confirmar si `TbExpedientes1` es copia, vista o contrato histórico. |
| No_Conformidades | **Probable/fuerte parcial** por clases y documentación | **Fuerte parcial**: entidades de expediente; **Excepción** con `Nemotecnico` en flujo concreto | Proyectos, auditorías, documentos, indicadores y caché | Separar flujos de proyecto y auditoría; medir qué rutas consultan Expedientes realmente. |
| Condor | **Probable**: clases de usuario/permisos; contrato runtime no validado | **Fuerte**: FK `tbSolicitudes.idExpediente` | `TbNoConformidades` para CD_CA, workflow y adjuntos | Backend no localizado/configuración Dysflow ausente. |
| HPS_Solicitudes | Desconocida en runtime; PRD menciona usuarios/permisos | **Fuerte documental**: `TbSolicitudes.IDExpediente` | HPS, suministradores, correos, MARGA/DPS/ONS y automatizaciones | No existe checkout/binary identificado; no mezclarlo con HPS sin resolver identidad. |
| HPS | **Fuerte parcial**: `UsuarioAplicacionPermisos` y acceso a Lanzadera | **Fuerte**: `TbExpedientes`, `IDExpediente`, suministradores | Solicitudes HPS, SICA, históricos, Excel y anexos | Distinguir `HPST.accdb` actual de la consolidación histórica documentada. |
| Brass | **Fuerte documental/parcial estática**: Lanzadera y permisos | **Desconocida/probable**, no demostrada en esta pasada | Eventos, actividades, materiales, equipos, SLA, informes y facturación | Backend ausente y falta rastreo específico de `IDExpediente`/`CodExp`. |
| Expedientes | **Fuerte parcial**: usuarios/permisos en el propio sistema | **Núcleo confirmado**: crea y mantiene el expediente | Suministradores, responsables, E2E/JSON, anexos, informes | Confirmar consumidores actuales y ciclo de baja/cierre visto desde cada aplicación. |

### Excepciones a la dependencia de Expedientes

1. **No Conformidades**: ciertos flujos conservan `Nemotecnico` y pueden operar sin leer `TbExpedientes`; debe documentarse el alcance exacto de la excepción.
2. **Lanzadera**: es el proveedor de identidad/catálogo, no un consumidor de expedientes según la evidencia actual.
3. **Brass**: no se puede afirmar aún dependencia de Expedientes; queda como pendiente, no como dependencia automática.
4. **HPS_Solicitudes**: la dependencia es documental, pero la aplicación ejecutable no está localizada.

## Estructura documental corregida para `access2web-blueprint/`

La documentación debe organizarse por **ecosistema → aplicación → capacidad**, separando topología observada, contratos transversales y futura migración. No se debe introducir todavía una arquitectura web objetivo como si fuera estado actual.

```text
docs/
├── 00-alcance-y-evidencia.md
├── 01-inventario-aplicaciones.md
├── 02-topologia-ecosistema/
│   ├── lanzadera-identidad-permisos.md
│   ├── expedientes-ciclo-de-vida.md
│   ├── matriz-dependencias.md
│   ├── identificadores-y-glosario.md
│   └── topologia-frontends-backends.md
├── 03-aplicaciones/
│   ├── lanzadera/
│   ├── gestion-riesgos/
│   ├── no-conformidades/
│   ├── condor/
│   ├── hps-solicitudes/
│   ├── hps/
│   ├── brass/
│   └── expedientes/
├── 04-integraciones-y-operacion/
│   ├── tablas-vinculadas-y-backends.md
│   ├── identidad-arranque-y-permisos.md
│   ├── informes-exportaciones-y-correo.md
│   ├── procesos-batch-y-automatizaciones.md
│   └── rutas-entornos-y-contingencia.md
├── 05-capacidades/
│   ├── capabilities-index.md
│   └── <una-ficha-por-capacidad>.md
├── 06-seguridad-y-trazabilidad.md
├── 07-migracion/
│   ├── modelo-dominio-agnostico.md
│   └── matriz-legacy-a-web.md
└── 08-decisiones-y-preguntas-abiertas.md
```

Cada ficha de capacidad seguirá `access-vba-capability-docs`: contrato de comportamiento, datos, dependencias, integraciones, informes/batch, excepciones, confianza (`Verified-runtime`, `Verified-static`, `Intended`, `Likely`, `Divergent`) y trazabilidad. Los recuentos de objetos se mantendrán como metadatos de descubrimiento, no como capacidades.

## Fase 1: secuencia de descubrimiento con aprobación por lote

Cada lote debe producir un paquete revisable y detenerse hasta recibir aprobación explícita. La extracción completa cubrirá capacidades, formularios, datos, integraciones, informes y batch.

| Lote | Objetivo y aplicaciones | Entregable mínimo antes de la puerta |
|---|---|---|
| 1 | **Lanzadera** | Catálogo real de `TbAplicaciones`, usuarios, roles, permisos, arranque, comandos, backends, formularios de administración, informes y tareas. |
| 2 | **Expedientes** | Alta, edición, estados, entidades, suministradores, responsables, anexos, consultas, informes, E2E/batch y contratos de `IDExpediente`/`CodExp`/`Nemotecnico`. |
| 3 | **Gestion_Riesgos** | Capacidades de proyectos/riesgos/ediciones, formularios por capability, tablas locales/vinculadas, permisos, informes y automatizaciones; resolver `TbExpedientes1`. |
| 4 | **HPS** | Solicitudes, usuarios HPS, expedientes, SICA, históricos, anexos, indicadores, exportaciones y backends; separar estado actual de consolidaciones. |
| 5 | **No Conformidades** | NC de proyectos/auditorías, cachés, formularios, indicadores, documentos, permisos, consultas a Expedientes y delimitación de la excepción `Nemotecnico`. |
| 6 | **Condor** | Solicitudes, workflow, expediente obligatorio, NC asociada, adjuntos, informes, batch y backend; resolver `.dysflow/project.json` y backend real antes de cerrar. |
| 7 | **Brass** | Eventos, actividades, equipos, materiales, SLA, facturación, informes/batch, permisos y rastreo específico de expediente; resolver backend y configuración Dysflow. |
| 8 | **HPS_Solicitudes** | Solo después de resolver el repositorio/binario: validar si es aplicación independiente, línea histórica o parte de HPS; entonces documentar solicitudes, workflow, correo y automatizaciones. |
| 9 | **Cierre transversal** | Reconciliar identificadores, dependencias, excepciones, rutas, integraciones, criticidad y huecos; registrar divergencias y baseline aprobado. |

### Contenido obligatorio de cada lote

1. Identidad del repositorio, checkout y binarios inspeccionados.
2. Estado Dysflow/CodeGraph y límites de acceso.
3. Inventario normalizado de formularios, informes, módulos, clases, queries y tablas.
4. Capacidades de negocio y flujos principales, no solo nombres técnicos.
5. Modelo de datos y cardinalidades relevantes.
6. Integraciones, archivos, correo, Excel, batch, periféricos y fallos operativos.
7. Dependencias hacia Lanzadera, Expedientes y otras aplicaciones.
8. Ledger de confianza y preguntas que permanecen abiertas.

## Enfoques considerados y recomendación

1. **Ejes transversales primero** — Lanzadera y Expedientes, seguidos por consumidores.
   - Pros: valida primero los contratos que más aplicaciones comparten y reduce duplicación.
   - Contras: puede retrasar el detalle de una aplicación periférica.
   - Esfuerzo: Medio.

2. **Aplicación completa por aplicación** — cerrar una ficha integral antes de cruzar dependencias.
   - Pros: facilita la revisión con el propietario de cada aplicación.
   - Contras: repite identidad/expediente y puede ocultar contratos comunes.
   - Esfuerzo: Alto.

**Recomendación:** enfoque híbrido por lotes: comenzar por Lanzadera y Expedientes, completar después cada consumidor verticalmente y cerrar con reconciliación transversal. La puerta de aprobación después de cada lote evita que una hipótesis propagada se convierta en arquitectura asumida.

## Riesgos y controles

- `HPS_Solicitudes` no tiene checkout principal localizado: no inferir que es HPS ni inventar una inspección runtime.
- Condor y Brass carecen de `.dysflow/project.json`; Expedientes tiene configuración `path-mismatch`. No mutar configuraciones durante esta fase; resolverlas en un lote aprobado.
- No Conformidades no pudo ser inventariada por Dysflow pese a tener configuración válida; el error debe diagnosticarse antes de afirmar recuentos.
- Los backends ausentes en algunos `00_main` pueden ser locales no versionados, rutas compartidas o binarios de otro entorno. Registrar procedencia y fecha de cada ruta.
- Engram contiene observaciones históricas que pueden mezclar legacy, staging y decisiones futuras. Se conserva como contexto, no como prueba de estado actual.
- `IDExpediente`, `CodExp` y `Nemotecnico` pueden no ser intercambiables; cualquier normalización web requiere tabla de correspondencias y cardinalidades verificadas.
- Los recuentos Dysflow obtenidos son inventarios técnicos preliminares y no sustituyen el catálogo funcional.

## Preguntas de mayor valor (máximo cinco)

1. ¿Cuál es el repositorio/binario principal de **HPS_Solicitudes** y debe documentarse como aplicación independiente de HPS?
2. ¿Cuál es la ruta y el nombre del backend vigente de Condor y Brass, y qué entorno debe considerarse baseline?
3. ¿Se autoriza configurar solo lectura los targets Dysflow de Condor, Brass y Expedientes, y diagnosticar el fallo de inventario de No Conformidades?
4. ¿Qué lote debe priorizarse después de aprobar Lanzadera y Expedientes: Gestion_Riesgos, HPS o No Conformidades?
5. ¿Qué catálogo de `TbAplicaciones` y qué versión de backends debe considerarse el baseline operativo actual?

## Ready for Proposal

**No.** La topología general y el plan de Fase 1 ya son suficientemente concretos para solicitar aprobación del primer lote, pero todavía no procede una propuesta de migración. El siguiente paso recomendado es que el usuario confirme las rutas/identidad de HPS_Solicitudes, los backends no localizados y la puerta de acceso para el lote 1; después se completa Lanzadera y se detiene para revisión.
