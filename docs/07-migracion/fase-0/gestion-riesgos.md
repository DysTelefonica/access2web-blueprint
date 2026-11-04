# Fase 0: Gestión de Riesgos (`Gestion_Riesgos`)

Épica #1545 (`docs/07-migracion/epica-coexistencia-postgresql.md`), fase 0. Ficha equivalente de Brass: [`brass.md`](brass.md).
Base de análisis: rama `release/2026-008.2.rc1` (SHA `3ee3394`), worktree `00_GESTION_RIESGOS-worktrees/2026-008.2.rc1`, sobre `fix/publicacion-calidad-ui-y-ediciones` en `f457571`.
Fuentes: ERD del front y del backend con dysflow en solo lectura; trazado del código con CodeGraph sobre el worktree y, para el texto que CodeGraph no indexa (SQL concatenado, cadenas), con `read`/`grep` sobre `src/`. CodeGraph está operativo en este entorno (índice de 473 ficheros y 22.261 nodos); el fallback se declara para lo que el índice no cubre.

## Base de análisis

| Hecho | Evidencia |
|---|---|
| El rc1 del ciclo es `release/2026-008.2.rc1` | `docs/uat/roadmap.md:7` («`staging` ya no existe: el rc1 es `release/2026-008.2.rc1` y todo trabajo parte de ella y vuelve a ella mediante PR») |
| El ciclo está en `mid_sprint`, no en `uat_handoff` | `docs/uat/roadmap.md:14` (compuerta) y `docs/uat/roadmap.md:15` («Próxima gate \| `uat_handoff` (al firmar todas las UAT)») |
| La UAT del rc1 está pendiente de firma | `docs/uat/roadmap.md:170` («UAT del rc1 con Natalia […] Su registro firmado alimenta el `uat_handoff`») y `docs/capabilities/index.md:57` («Alta para modelo funcional; release/UAT pendiente») |
| La web de aceptación del rc1 es la vigente | `docs/uat/uat-rc1-2026-10-28.html:277` (`criteriaVersion: "rc1-2026-11-04.3"`) |
| El trabajo de la RC1 no está todo commiteado | `git status` en el worktree: modificados `src/forms/Form_FormPublicacionCalidadPublicarEjecutar.cls`, `src/forms/Form_FormRiesgosGestion.cls`, `src/modules/modFormCoordinationHelper.bas`; sin seguimiento `src/modules/modRiesgosGestionEdicionesHelper.bas` y `src/modules/Test_RiesgosGestionEdicionTrasPublicar337.bas` |
| `main` no es una base válida | `main` local en `36fcf60`, cuatro commits por detrás de `origin/main`; entre `main` y `release/2026-008.2.rc1` hay 416 ficheros distintos en `src/` (108.909 inserciones y 30.034 borrados) |
| No hay macros ni `AutoExec` | Cero coincidencias de `autoexec` en `src/`; el ERD del front no recoge módulos |

Consecuencia: el inventario de abajo describe el estado del worktree RC1 (rama `fix/publicacion-calidad-ui-y-ediciones` en `f457571` más los cambios locales sin commitear), no `main`. La diferencia con `main` no es un error del inventario.

## Estructura del front

| Hecho | Valor |
|---|---|
| Tablas del front | 2, ambas propias: `TbAuxPriorizacion` (12 filas) y `TbConfiguracionBackends` (1 fila) |
| Vínculos del front | 0: el front no tiene tablas vinculadas |
| Formularios | 61 (`Form_*.form.txt` y `Form_*.cls` parejas) |
| Formularios con `RecordSource` | 0: cero apariciones de `RecordSource` en todo `src/` |
| Controles con `ControlSource` con valor | 0 |
| Controles con `RowSourceType = Table/Query` | 1: `ListaRiesgos` en `src/forms/Form_FormRiesgosEstablecerPrioridades.form.txt:707-710`, sobre `TbAuxPriorizacion`, tabla propia del front |
| Controles con `RowSourceType = Value List` | 75, todos con listas literales |

Los formularios son desconectados, como establece el propietario, y los datos se cargan en tiempo de ejecución: las asignaciones de `RowSource` en código son listas literales (`src/forms/Form_FormAnexos.cls:296` y `:577`, `src/forms/Form_FormCalidadRiesgoMaterializaciones.cls:422`, `src/forms/Form_Form0BDOpciones.cls:439`, que carga `m_ObjEntorno.ListaUsuarios`). El único origen de datos tabular del front es la tabla propia de trabajo `TbAuxPriorizacion`, que la borra y rellena `PriorizacionRepository` (`src/modules/PriorizacionRepository.bas:13`, `:40`, `:129`).

`TbConfiguracionBackends` es la tabla que decide qué backend abre el front: `src/modules/Variables Globales.bas:1338-1339` la lee con `CurrentDb` y `:1348`/`:1352` toman de ella la ruta de producción o la de sandbox, y `:1359` toma de ella la contraseña del backend.

## Aperturas directas

| Apertura | Código | Observación |
|---|---|---|
| Backend propio | `src/modules/Variables Globales.bas:1441` (`getdb()`), que abre con `ws.OpenDatabase(m_URL, False, False, ";PWD=" & m_PasswordBackend)` en `:1555` (producción) y `:1505` (modo pruebas, contra el sandbox) | `m_URL` no está escrito en el código: sale de `Application.TempVars("BackendPathConfigurado")` (`:1539`), que `LeeConfiguracionLocal` rellena desde `TbConfiguracionBackends` (`:1348`, `:1352`) |
| Otros backends | Ninguna | No hay ninguna ruta `.accdb` literal en `src/`; solo dos comentarios que nombran `Expedientes_datos.accdb` (`src/modules/Constructor.bas:5132`, `:5208`) |
| `DoCmd.OpenTable`, `TransferDatabase` | Ninguno | Cero coincidencias en `src/` |
| `CurrentDb` | `src/modules/Variables Globales.bas:1338`, `src/modules/PriorizacionRepository.bas:12`, `:39`, `src/forms/Form_FormPublicacionCalidadPublicarEjecutar.cls:259` | Las tres primeras apuntan a tablas del front (`TbConfiguracionBackends`, `TbAuxPriorizacion`). Las otras tres son la excepción: `src/modules/modMigracionAux.bas:208` y `:244` escriben en `TbIndicadoresRiesgos`, que no existe ni en el front ni en el backend, y `src/modules/IndicadorRiesgosRepositorio.bas:481` resuelve `CurrentDb.TableDefs(tableName)` sobre nombres de tabla del backend (por ejemplo `TbRiesgosAIntegrar` en `:494`), de modo que en producción la comprobación devuelve falso |

La cadena de conexión efectiva es `MS Access;PWD=***;DATABASE=<fichero del backend>`; la credencial no está en el código, pero sí en una tabla del front, en un campo de texto de 50 caracteres (`PasswordBackend`). Los vínculos del backend, en cambio, sí llevan la credencial en su definición: `MS Access;PWD=***;DATABASE=<origen>`.

## Tablas de otros backends usadas por código

Las 14 tablas vinculadas del backend y su uso real en el VBA del front. «Se usa» significa que el código la nombra en SQL o en `DameID`; «no se usa» significa cero referencias activas en `src/`.

| Tabla | Backend de origen | Uso en el front | Evidencia |
|---|---|---|---|
| `TbUsuariosAplicaciones` | `Lanzadera_Datos` | Se usa | `src/modules/Constructor.bas:3900-3901`, `:3992-3993`, `:4217-4221` |
| `TbUsuariosAplicacionesPermisos` | `Lanzadera_Datos` | Se usa | `src/modules/Constructor.bas:3993-3997`, `:4046` |
| `TbExpedientes` | `Expedientes_datos` | Se usa | `src/modules/Constructor.bas:3626`, `:6753`, `:6757` |
| `TbExpedientesResponsables` | `Expedientes_datos` | Se usa (solo lectura) | `src/classes/ExpedienteResponsable.cls:37`, `src/modules/Constructor.bas:3669`, `:4218-4220` |
| `TbExpedientesSuministradores` | `Expedientes_datos` | Se usa | `src/modules/Constructor.bas:5153`, `:5233`, `:5370`, `:5451` |
| `TbSuministradores` | `Expedientes_datos` | Se usa | `src/modules/Constructor.bas:5154`, `:5234`, `:5363`, `:5371`, `:5445` |
| `TbNoConformidades` | `NoConformidades_Datos` | Se usa y se escribe | `src/classes/NC.cls:188`, `:344-346`; inserción en `:338-355` |
| `TbCorreosEnviados` | `Correos_datos` | Se usa y se escribe | `src/classes/Correo.cls:43`, `:271-273`; inserción en `src/modules/modAvisoPublicacionHelper.bas:131` |
| `TbDocumentos` | `AGEDO20_Datos` | Se usa y se escribe | `src/classes/Documento.cls:65`, `:262-264`, `:326`; `src/modules/Constructor.bas:4379` |
| `TbAplicaciones` | `Lanzadera_Datos` | No se usa | 0 referencias; el front toma `IDAplicacion` de `TbConfiguracionBackends` (`src/modules/Variables Globales.bas:1360`) |
| `TbDocumentosID` | `AGEDO20_Datos` | No se usa | 0 referencias |
| `TbDPD` | `AGEDYS_DATOS` | No se usa | 0 referencias |
| `TbResponsablesExpedientes` | `AGEDYS_DATOS` | No se usa | 0 referencias |
| `TbSuministradoresSAP` | `AGEDYS_DATOS` | No se usa | 0 referencias; el código nombra el singular inexistente `TbSuministradorSAP` (`src/classes/Pedido.cls:95`) |

El front no abre `Lanzadera_Datos` por ruta: consume la identidad a través de las tablas vinculadas que el backend ya trae (`TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos`). Eso confirma R-1 para esta aplicación en su forma concreta: un solo backend, con vínculos a seis backends ajenos, y ninguna apertura directa de un backend ajeno.

El punto que R-2 sí marca es la escritura: el front inserta y borra filas en `TbNoConformidades`, `TbCorreosEnviados` y `TbDocumentos`, que pertenecen a otros backends. La tabla debe quedar donde está, pero el modo de escritura es un dato de la dependencia que conviene fijar antes de migrar.

## Backend `Gestion_Riesgos_Datos`

| Hecho | Valor |
|---|---|
| Tablas | 78: 64 propias y 14 vinculadas |
| Relaciones | 19, de las que una es del sistema (`MSysNavPaneGroupCategories` → `MSysNavPaneGroups`); 18 reales |
| Consultas guardadas | 35 |
| Tablas sin clave primaria | 5: `ProyectoRespCalidad`, `TbAuxProyectosRiesgos`, `TbCachePublicabilidadEdicion`, `TbCorreosEnviados`, `TbRiesgosValoracion` |
| Campos multivalor, de datos adjuntos u objeto OLE | 0: los tipos presentes son texto, entero largo, fecha/hora, memo, booleano, entero, doble y byte |

El recuento de 78 tablas coincide con el ERD versionado en el propio repo: `docs/ERD/backend-erd.md:8` («## Tablas (78)»), con fingerprint en `:4` y las relaciones en `:1119-1141`. Los 78 nombres coinciden uno a uno entre ese ERD y el extraído hoy.

Los vínculos del backend apuntan a seis backends: 4 a `Expedientes_datos`, 3 a `Lanzadera_Datos`, 3 a `AGEDYS_DATOS`, 2 a `AGEDO20_Datos`, 1 a `NoConformidades_Datos` y 1 a `Correos_datos`.

### Consultas guardadas

Es el mayor volumen de consultas guardadas del ecosistema (35, frente a 0 en Brass). El tratamiento propuesto se apoya en dos hechos medidos:

- **Ninguna se invoca desde el front.** No hay `OpenQueryDef` en `src/`; los `FROM [` que existen toman el nombre de una tabla constante (`src/modules/MigracionConfiguracionBackend.bas:11` y `:62`), no el de una consulta guardada. El front reimplementa la lógica en SQL embebido. Las únicas coincidencias de nombre son tres constantes de árbol de tareas (`EDICIONESPREPARADASPARAPUBLICAR` en `src/classes/ArbolTareasCalidad.cls:133` y `src/modules/Constantes.bas:302`, `RIESGOSACEPTADOSRECHAZADOS` en `src/classes/ArbolTareasTecnico.cls:172` y `src/modules/Constantes.bas:318`, `RIESGOSRETIRADOSRECHAZADOS` en `src/classes/ArbolTareasTecnico.cls:198` y `src/modules/Constantes.bas:320`), que no son llamadas a consulta guardada.
- **Se dividen en dos familias por el nombre.** Cuatro tienen nombre de trabajo ad hoc: `AntonioNoTieneQueEstar`, `ConsultaFelix`, `Consulta1`, `Consulta4`. Las otras 31 llevan nombre de negocio (`EDICIONESNECESITANPROPUESTAPUBLICACION`, `RIESGOSRETIRADOSNOMOTIVADOS`, `RiesgosIdentificadosEntreDosFechas`, `UltimaEdicionProyectos` y familia `_PorUsuario`).

No se pudo leer el SQL de las 35: `MSysQueries` no es legible con el usuario disponible («no tiene permiso para READ en `MSysQueries'»), y el resto de herramientas de lectura no expone las consultas guardadas. La propuesta R7 no puede cerrarse con esta evidencia: exige una sesión de Access que exporte las definiciones.

## Veredicto de propiedad por tabla

### Se quedan en el backend: 32 tablas propias

Las que el VBA del front nombra. La cifra entre paréntesis es el número de referencias al nombre de la tabla en `src/` (líneas no comentadas, módulos de test incluidos): `TbRiesgos` (757), `TbProyectosEdiciones` (647), `TbProyectos` (417), `TbRiesgosMaterializaciones` (192), `TbRiesgosPlanMitigacionPpal` (108), `TbAnexos` (95), `TbRiesgosPlanMitigacionDetalle` (92), `TbProyectosEdicionesSuministradores` (90), `TbRiesgosPlanContingenciaPpal` (87), `TbRiesgosPlanContingenciaDetalle` (72), `tbCambiosParaPublicacion` (37), `TbProyectoEdicionesCorreoRevision` (34), `TbRiesgosAIntegrar` (32), `TbAuxProyectosRiesgos` (29), `TbLogPublicaciones` (28), `TbCachePublicabilidadEdicion` (26), `TbCacheControlCambiosRow` (18), `TbCacheControlCambiosMeta` (17), `TbBibliotecaRiesgos` (19), `TbConfiguracionVisionRiesgos` (12), `tbCambios` (10), `TbExplicacionCarencias` (8), `TbRiesgosValoracion` (8), `TbUltimoProyecto` (8), `TbRiesgosPlanContingenciaDetalleReversa` (5), `TbRiesgosPlanMitigacionDetalleReversa` (5), `TbRiesgosNC` (4), `TbHerramientaDocAyuda` (2), `TbCambiosExplicacion` (1), `TbConfiguracion` (1), `TbRiesgosAreasImpacto` (1), `TbTareasExplicaciones` (1).

### No las usa el front: 30 tablas propias

Cero referencias activas en `src/`. Se agrupan por su origen:

| Grupo | Tablas | Nota |
|---|---|---|
| Lookup de IDs históricos (14) | `TbIDAccionPlanContingencia`, `TbIDAccionPlanMitigacion`, `TbIDAnexos`, `TbIDBibliotecaRiesgos`, `TbIDCorreosEnviados`, `TbIDDocumentos`, `TbIDEdiciones`, `TbIDLog`, `TbIDPlanContingencia`, `TbIDPlanMitigacion`, `TbIDProyectos`, `TbIDRiesgos`, `TbIDRiesgosAIntegrar`, `TbIDTareas` | El propio repo lo dice: «`TbID*` → tablas lookup de IDs históricos (no se usan en runtime, solo auditoría)» (`docs/codebase/data-model.md:179`) |
| Catálogos de valores posibles (5) | `TbValoresPosiblesContingencia`, `TbValoresPosiblesEstadoRiesgo`, `TbValoresPosiblesMitigacion`, `TbValoresPosiblesPlazoCalidadCosteVulnerabilidad`, `TbValoresPosiblesValoracion` | La app calcula la valoración con SQL propio contra `TbRiesgosValoracion` (`src/modules/Constructor.bas:4756-4759`), no contra los catálogos |
| Resto (11) | `TbAnexosAntigua`, `TbAux`, `TbAuxSumExp`, `TbFelix`, `TbLog`, `TbOrigenesRiesgosDetalles`, `TbProyectosResponsablesCalidad`, `TbProyectosSuministradores`, `TbProyectosTipo`, `TbRiesgosPorTipoProyecto`, `TbTablas` | `TbAnexosAntigua` (2 filas) y `TbProyectosTipo`/`TbRiesgosPorTipoProyecto` (0 filas) son restos; `TbLog` conserva 5.817 filas |

A estas se suman dos que no encajan en el patrón: `ProyectoRespCalidad` (16 filas) solo aparece como constante numérica en `src/modules/Constantes.bas:146`, no como tabla; y `TbTareas` solo se nombra en un módulo de test (`src/modules/Test_CalidadCacheRiesgo.bas:124`).

### No las usa el front: 5 tablas vinculadas

`TbAplicaciones`, `TbDocumentosID`, `TbDPD`, `TbResponsablesExpedientes` y `TbSuministradoresSAP` están vinculadas en el backend y no las nombra el front. No son huérfanas por sí solas: `TbDPD` y `TbResponsablesExpedientes` son tablas de AGEDYS y pueden tener su dueño en el front de AGEDYS. Su retirada del backend de Riesgos es una decisión de ubicación, no una baja de datos, y exige el grafo de vínculos de todos los backends (tarea T6).

### El front usa tablas que no existen

Once tablas no están ni en el front ni en el backend, y el código las consulta:

| Tabla ausente | Evidencia |
|---|---|
| `TbAplicacionesVersiones` | `src/modules/Constructor.bas:6541`, `:7455`, `:7459`, `:7503`, `:7605-7607` |
| `TbVersionCambios` | `src/modules/Constructor.bas:7558`, `:7601`, `:7604-7606` |
| `TbRACS` | `src/classes/RAC.cls:31`, `src/modules/Constructor.bas:7773`, `:7777` |
| `TbExpedientesRACS` | `src/modules/Constructor.bas:7821-7822` |
| `TbProyectosJefesDeProyecto` | `src/classes/Proyecto.cls:4230-4232` |
| `TbOrganosContratacion` | `src/classes/OrganoContratacion.cls:30`, `src/modules/Constructor.bas:7724`, `:7728` |
| `TbJuridicas` | `src/classes/Juridica.cls:30` |
| `TbTipologia` | `src/modules/Constructor.bas:6397` |
| `TbCambioDocumentos` | `src/modules/Constructor.bas:7663` |
| `TbExpedientesConEntidades` | `src/modules/Constructor.bas:3723` |
| `TbIndicadoresRiesgos` | `src/modules/modMigracionAux.bas:172`, `:208`, `:244` |

Hay además dos referencias con nombre mal escrito: `TbExpedientes1` en `src/modules/Constructor.bas:4271-4272` (no es un alias, es el nombre de tabla) y `TbSuministradorSAP` en `src/classes/Pedido.cls:95` (la tabla se llama `TbSuministradoresSAP`).

Los procedimientos afectados tienen llamadores en el código de producción (`getColJPs` ← `src/classes/Entorno.cls:1674`; `getOrganoContratacion` ← `src/classes/Expediente.cls:212`; `getExpedienteRACS` ← `src/classes/Expediente.cls:190`; `getVersionesAplicacion` ← `src/classes/Entorno.cls:257`; `getTiposNC` ← `src/classes/Entorno.cls:2601`; `getCCVersion` ← `src/forms/Form_FormControlCambiosGestion.cls:147`). El error 3070 solo se materializa si esa ruta se ejecuta, así que el coste se mide en la fase 1, no antes.

## Patrones de riesgo en el código

| Patrón | Medida | Evidencia |
|---|---|---|
| Generador `Max + 1` | `DameID` calcula `Max(id) + 1` | `src/modules/Funciones Generales.bas:2362-2401`: SQL en `:2381-2382`, incremento en `:2394` |
| Tablas con identificador generado | 22: 18 propias y 4 de otros backends | `DameID` sobre `TbAnexos`, `tbCambios`, `TbExplicacionCarencias`, `TbProyectosEdiciones`, `TbProyectoEdicionesCorreoRevision`, `TbConfiguracionVisionRiesgos`, `TbRiesgosPlanContingenciaPpal` y `Detalle`, `TbRiesgosPlanMitigacionPpal` y `Detalle`, `TbProyectos`, `TbLogPublicaciones`, `TbRiesgos`, `TbBibliotecaRiesgos`, `TbRiesgosAIntegrar`, `TbRiesgosMaterializaciones`, `TbProyectosEdicionesSuministradores`, `TbUltimoProyecto`, `TbCorreosEnviados`, `TbDocumentos`, `TbExpedientesResponsables`, `TbNoConformidades` |
| Comparación insensible a mayúsculas | 356 de 357 módulos con `Option Compare Database` | Solo `src/modules/JsonConverter.bas` no lo declara |
| Registros sin `dbSeeChanges` | 571 `OpenRecordset` y 0 `dbSeeChanges` | Búsqueda en todo `src/` |
| Literales de fecha con `#` | 45 líneas en 8 ficheros | `src/classes/Edicion.cls:2954`, `:3227-3228`, `:4852`; `src/classes/EdicionCorreoRevision.cls:314`; `src/classes/Proyecto.cls:780`, `:3737`, `:4180`, `:4388`, `:4409`; `src/modules/Funciones Generales.bas:4595`, `:4600`; `src/modules/Instalacion.bas:259`, `:263`, `:732`, `:789`; `src/modules/IndicadorRiesgosRepositorio.bas:114`, `:124`, `:133`, `:142`, `:154`, `:225`, `:240`, `:254`, `:265`, `:278`, `:294-295`, `:353`, `:372`, `:389`, `:407`, `:429`, `:453-454`; `src/modules/mIndicador.bas:463`, `:530`, `:1375`, `:1389`, `:1411`, `:1425`, `:1448`, `:1462`; `src/modules/modAvisoPublicacionHelper.bas:135-136` |
| `Seek`, `dbOpenTable`, `LIKE '%'` | 0 apariciones | Búsqueda en todo `src/` |
| SQL concatenado | Presente en las clases de dominio (`Riesgo.cls`, `Proyecto.cls`, `Edicion.cls`, `NC.cls`) | `src/classes/Riesgo.cls:5800` y `src/classes/Edicion.cls:2952-2955` son ejemplos |

## Propuestas para después del descubrimiento global

Ninguna se aplica: la puerta de descubrimiento global sigue cerrada (#1576). Todas son propuestas con comprobación previa y reversión.

| Id | Propuesta | Grupo (fase 1) | Comprobación previa | Reversión |
|---|---|---|---|---|
| R1 | Sacar la contraseña del backend de la tabla del front: `TbConfiguracionBackends.PasswordBackend` la guarda en un campo de texto de 50 caracteres (`src/modules/Variables Globales.bas:1359`) y las definiciones de vínculo la repiten (`PWD=***`). Propuesta: credencial en la configuración de la aplicación o en el gestor de secretos del servidor, no en la base. | (a) Imprescindible antes de migrar | Inventariar qué aplicaciones leen `TbConfiguracionBackends` y comprobar que el valor no se usa para nada más que abrir el backend. Rotar la credencial en una copia y comprobar que el front sigue arrancando. | Volver a escribir el campo y su valor. |
| R2 | Declarar clave primaria en las 5 tablas sin ella: `ProyectoRespCalidad`, `TbAuxProyectosRiesgos`, `TbCachePublicabilidadEdicion`, `TbCorreosEnviados`, `TbRiesgosValoracion`. Sin clave no se puede actualizar por ODBC en la coexistencia. | (a) Imprescindible antes de migrar | Verificar duplicados en cada tabla antes de fijar la clave. `TbCorreosEnviados` es además tabla de otro backend: la clave la decide su propietario. `TbCachePublicabilidadEdicion` tiene 0 filas y clave única compuesta ya definida (`UX_TbCachePublicabilidadEdicion`). | Quitar la clave. |
| R3 | Poner en cuarentena las 30 tablas propias sin uso en el front: renombrar a `zzBAJA_<tabla>`. | (a) Imprescindible antes de migrar | Cruce de los vínculos de todos los backends (tarea T6) y búsqueda en el `src` de los demás fronts, incluido el SQL dinámico. Verificar que ninguna de las 35 consultas guardadas las usa, algo que hoy no se puede comprobar porque su SQL no es legible. Revisar aparte `ProyectoRespCalidad` (solo aparece como constante) y `TbTareas` (solo la nombra un módulo de test). | Quitar el prefijo. |
| R4 | Retirar del backend los 5 vínculos que el front no usa: `TbAplicaciones`, `TbDocumentosID`, `TbDPD`, `TbResponsablesExpedientes`, `TbSuministradoresSAP`. | (a) Imprescindible antes de migrar | Confirmar con el propietario de AGEDYS que `TbDPD`, `TbResponsablesExpedientes` y `TbSuministradoresSAP` no las necesita ningún front, y que `TbAplicaciones` y `TbDocumentosID` no las usa ningún otro consumidor de `Lanzadera_Datos` y `AGEDO20_Datos`. | Recrear el vínculo. |
| R5 | Resolver las 13 referencias a tablas inexistentes o mal nombradas: `TbAplicacionesVersiones`, `TbVersionCambios`, `TbRACS`, `TbExpedientesRACS`, `TbProyectosJefesDeProyecto`, `TbOrganosContratacion`, `TbJuridicas`, `TbTipologia`, `TbCambioDocumentos`, `TbExpedientesConEntidades`, `TbIndicadoresRiesgos`, `TbExpedientes1` y `TbSuministradorSAP`. O la tabla existe en otro backend y se vincula, o la referencia se elimina. | (a) Imprescindible antes de migrar | Reproducir en runtime los recorridos afectados: panel de responsables del expediente (`getColJPs`), versionado de la aplicación (`getVersionesAplicacion`), control de cambios (`getCCVersion`), ficha de NC (`getTiposNC`) y alta de documentos. | Restaurar la referencia original. |
| R6 | Sustituir el generador `Max + 1` (`DameID`) por identidad en las 22 tablas listadas. | (a) Imprescindible antes de migrar | Verificar que el valor máximo actual no colisiona con el siguiente valor que asignaría la identidad, y medir las 22 llamadas: 18 sobre tablas propias y 4 sobre tablas de otros backends, que requieren acuerdo previo. | Volver a `DameID`. |
| R7 | Decidir el tratamiento de las 35 consultas guardadas: las 4 de nombre ad hoc (`AntonioNoTieneQueEstar`, `ConsultaFelix`, `Consulta1`, `Consulta4`) son candidatas a cuarentena; las 31 de negocio, cuya lógica el front ya reimplementa en SQL, son candidatas a retirada tras confirmar que ningún consumidor externo las usa. | (b) Recomendable si es barato para las 4 ad hoc; (c) Aplazar a la refactorización final para las 31 de negocio | Exportar el SQL de las 35 desde una sesión de Access (hoy `MSysQueries` no es legible) y comprobar consumidores fuera de Access: Excel, informes y scripts programados. | Restaurar las consultas desde la copia de seguridad del backend. |
| R8 | Fijar la semántica de comparación de texto antes de migrar: 356 de 357 módulos declaran `Option Compare Database`, y PostgreSQL no hereda esa comparación insensible a mayúsculas. La web y el legacy no pueden discrepar sobre el mismo dato. | (a) Imprescindible antes de migrar | Inventariar las comparaciones de texto en SQL embebido y en las consultas guardadas, y decidir por columna si la diferencia de comportamiento es aceptable. | Cambiar el criterio acordado y aplicarlo a la vez en el front y en la web; no deja cambios en el legacy que deshacer. |
| R9 | Decidir el nivel de aislamiento de las 571 aperturas de recordset, todas sin `dbSeeChanges`, antes de que el backend pase a ODBC. | (a) Imprescindible antes de migrar | Listar los procedimientos que leen y escriben las mismas filas (publicación de edición, cierre de edición, invalidación de cachés) y decidir para cada uno si necesita lectura serializable. | Ajustar el nivel de aislamiento de la configuración ODBC. |
| R10 | Sustituir los 45 literales de fecha con `#` de los 8 ficheros citados por parámetros o literales compatibles con ambos motores. | (b) Recomendable si es barato | Verificar el formato que produce cada `Format` (hay `mm/dd/yyyy` y `yyyy/mm/dd hh:nn:ss` en el mismo código) y el tratamiento de fechas nulas. | Restaurar los literales originales. |

Lo que **no** se toca ahora: las 14 tablas vinculadas (mientras la web no escriba en ellas), las 2 tablas propias del front, y las 18 relaciones reales del backend.

## Lo que no se pudo verificar

1. El SQL de las 35 consultas guardadas: `MSysQueries` no es legible con el usuario disponible («no tiene permiso para READ en `MSysQueries'») y las herramientas de lectura no exponen las consultas guardadas. Sin ese SQL no se puede distinguir una consulta viva de una muerta, ni saber si alguna de las 30 tablas sin uso aparece en ellas.
2. Si las 30 tablas propias sin uso y las 5 vinculadas sin uso las consume otro front: exige el grafo de vínculos de todos los backends (tarea T6).
3. Si las 11 tablas ausentes se recuperan de otro backend, están heredadas de una versión anterior del modelo, o son código muerto: el error 3070 solo aparece si la ruta se ejecuta, y no se ha ejecutado la aplicación.
4. La paridad entre el binario y `src/`: el worktree tiene el `.accdb` modificado y tres ficheros de `src/` sin commitear. El inventario refleja el código, no una verificación de paridad.
5. El estado de la UAT en el tracker: la evidencia es el roadmap del repo (`docs/uat/roadmap.md`), no la issue de UAT.
6. La divergencia con `docs/codebase/data-model.md`, que describe 81 tablas en el backend y sitúa `TbConfiguracionVisionRiesgos` y `TbAplicaciones` en el front (`:5`, `:23`, `:30`, `:175`), mientras que el ERD versionado el 2026-11-04 (`docs/ERD/backend-erd.md:4`, `:8`) y el ERD extraído hoy coinciden en 78 tablas, con `TbConfiguracionVisionRiesgos` propia del backend y `TbAplicaciones` vinculada a `Lanzadera_Datos`. La página del repo está desfasada respecto del ERD vigente.
