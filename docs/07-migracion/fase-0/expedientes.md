# Fase 0: Expedientes (`Expedientes` / `Expedientes_datos`)

Épica `docs/07-migracion/epica-coexistencia-postgresql.md` (#1545), fase 0. Esta ficha no cita ninguna issue propia: la asigna el coordinador, junto con el resto del descubrimiento global (puerta cerrada por #1576).

Base de análisis: rama `release/2026-008.2.rc1`, SHA `e71d9e7f57512184c6f1c03b94fc075ff07f6a85`, worktree `C:\00repos\codigo\00_EXPEDIENTES-worktrees\2026-008.2.rc1`. Front: el `Expedientes.accdb` de ese worktree. Backend: `C:\00repos\datos\Expedientes_datos.accdb`.

Fuentes: ERD del front y del backend con `scripts/dysflow-erd/extract-erd.ps1` (solo lectura), ya extraídos; trazado del código con CodeGraph sobre el índice del worktree y, en paralelo, lectura y búsqueda literal sobre `src/`. CodeGraph estuvo operativo (`.codegraph/codegraph.db`: 315 ficheros, 13 347 nodos, 29 567 aristas), pero su índice es de símbolos y no de SQL: el inventario por tabla se cerró con búsqueda literal de nombre exacto y CodeGraph se usó para el alcance de las llamadas.

## Base de análisis

| Hecho | Evidencia |
|---|---|
| La rama base es `release/2026-008.2.rc1` | `docs/DELIVERY_ROADMAP.md:32` («`release/2026-008.2.rc1` es el único destino de integración y entrega para todo trabajo futuro de Expedientes. El trabajo ordinario no debe targetear `main` ni `staging`») |
| La UAT de esa RC está pendiente, no superada | `docs/DELIVERY_ROADMAP.md:3` (ciclo UAT 4 «enviado a Calidad» el 2026-11-04, sin respuestas todavía), `:87` («Ejecute el 2026-09-23 la UAT agrupada del operador, incluyendo el Paso 4, y obtener la firma correspondiente»), `:295` (la issue #231 «sigue OPEN como gate cruzado del PR #160 (`release/2026-008.2.rc1 → main`)… Cierra cuando el owner haga la UAT final y promueva a main»), `:7` («freeze rc1 pre-UAT… `main` no se toca») |
| `main` no es la base de estudio | `main` está en `495b965`; `git diff --stat main..release/2026-008.2.rc1 -- src` da 166 ficheros, 44 642 altas y 3 462 bajas |
| El binario del front lleva cambios sin commitear | `git status` del worktree: `M Expedientes.accdb`. El ERD del front refleja ese binario, no el commit `e71d9e7` |

Consecuencia: el inventario de abajo describe la RC1 con el binario que hay en el worktree, no `main`. La diferencia con `main` no es un error del inventario.

## Estructura del front

| Hecho | Valor |
|---|---|
| Tablas del front | 6: 2 propias y 4 vinculadas |
| Tablas propias del front | `TbAuxEstadosMartina` (143 filas) y `TbConfiguracionBackends` (1 fila) |
| Destino de los 4 vínculos | Todos a `C:\00repos\datos\Expedientes_datos.accdb`, el backend en la ruta local, con la credencial en la propia cadena de conexión (`PWD=***`) |
| Vínculos del front a otros backends | Ninguno: la identidad de la Lanzadera y los datos de AGEDYS, HPS, Riesgos, Correos y No Conformidades llegan por los vínculos del backend |
| Formularios | 45 (`src/forms/*.form.txt`), con su code-behind |
| Formularios con `RecordSource` | 0 |
| Controles con `ControlSource` | 0 |
| Controles con `RowSourceType = "Table/View/StoredProc"` | 1: `ListaCola`, en `Form_FormE2ECola` (`src/forms/Form_FormE2ECola.form.txt:288`) |
| Controles con `RowSourceType = "Value List"` | 81; sus `RowSource` son listas literales (`src/forms/Form_Form0BDGestorEntidades.form.txt:949`) |
| Relaciones del front | 0 reales: la única del ERD es `MSysNavPaneGroupCategories` → `MSysNavPaneGroups`, del sistema |
| Consultas guardadas del front | 1: `Consulta1`, exportada en `src/queries/Consulta1.sql` |

La excepción de Brass no se repite aquí: los formularios de Expedientes están desconectados y se rellenan en tiempo de ejecución. La comprobación se hizo en el código, no en las propiedades de diseño: los combos y las listas se llenan con `AddItem` o con SQL que se ejecuta sobre `getdb()`.

Dos matices:

- El único control con origen de datos de tabla es `ListaCola`, cuyo `RowSource` se asigna en tiempo de ejecución (`src/forms/Form_FormE2ECola.cls:138`) con el SQL que compone `src/modules/modE2EColaPanelHelper.bas:74-86` sobre `TbE2ECola`. Ese SQL lo resuelve el motor del front, y el front no vincula `TbE2ECola`: el control no puede resolverlo. El formulario no se abre desde el menú desde que la opción 12 abre el panel HTML (`src/modules/Form0BDGestorEntidadesLogic.bas:36, 55-57`).
- `TbAuxEstadosMartina` existe en el front y en el backend, con las mismas 143 filas y el mismo propósito, pero con tipos distintos en `ID` (texto en el front, entero largo en el backend). Ningún código de `src/` la cita.

## Aperturas directas

| Apertura | Código | Observación |
|---|---|---|
| Backend propio | `src/modules/Variables Globales.bas:637` (`getdb()`) | `wks.OpenDatabase(m_TargetBackendURL, False, False, ";pwd=" & m_PasswordBackend)`. La URL la decide `BackendResolver.ResolveBackendPath` (`src/modules/BackendResolver.bas:5-34`) entre producción, sandbox y test, según `TbConfiguracionBackends.BackendActivo` (`src/modules/Variables Globales.bas:533-534, 541-544`) |
| Otro `.accdb` (Lanzadera, AGEDYS, HPST, Riesgos) | Ninguno | No hay ninguna ruta de `.accdb` en el código: a la Lanzadera se llega por los vínculos del backend. La única ruta absoluta que queda escrita en el código es la carpeta de producción, como constante de un módulo de etiquetado (`src/modules/modOrigenDatos.bas:18`) |
| `DoCmd.OpenTable`, `TransferDatabase`, `dbOpenTable` | Ninguno | 0 apariciones en `src/` |
| `CurrentDb` | 4 usos en producción | `src/modules/Variables Globales.bas:534` (tabla de configuración), `:654` (nombre del `.ini`), `src/modules/FUNCIONES UTILES.bas:2876` (respaldo de `DameID`, camino no recorrido) y `src/modules/M_E2EMigration.bas:38` (ruta de prueba). Ninguno abre datos de negocio |
| Macros | Ninguna | El propietario lo confirma: no hay macros ni `AutoExec` que inspeccionar |

Diferencia con Brass: la credencial del backend no está literal en el código. Vive en la tabla local del front, `TbConfiguracionBackends` (`PasswordBackend`), y el destino se resuelve por configuración. Es el patrón que la fase 2 quiere extender a los demás mecanismos de acceso.

## Tablas de otros backends que usa el front

El front no abre otros backends: alcanza las tablas de otras aplicaciones a través de los 24 vínculos de su propio backend. De esos 24, 10 los usa el código.

| Tabla | Backend de origen | Uso | Evidencia |
|---|---|---|---|
| `TbUsuariosAplicaciones` | `Lanzadera_Datos` | lectura | `src/modules/constructor.bas:49-50, 188, 219-224, 290, 1640, 2607-2609, 5702-5703`; `src/modules/ExpedienteJsonExporter.bas:487, 608-609, 983-992`; `src/modules/modBusquedaVivo.bas:66, 203, 219, 225`; `src/modules/modInformeVivo.bas:63-64, 236`; `src/modules/modResponsablePorRolHelper.bas:205` |
| `TbUsuariosAplicacionesPermisos` | `Lanzadera_Datos` | lectura | `src/modules/constructor.bas:133-134, 220-224` |
| `TbNoConformidades` | `NoConformidades_Datos` | lectura y escritura | `src/modules/FUNCIONES UTILES.bas:2552, 2572, 6390`; `UPDATE` en `:6456` y `:6459` |
| `TbProyectos_Riesgos` | `Gestion_Riesgos_Datos` | lectura y escritura | `src/modules/FUNCIONES UTILES.bas:2388, 5723, 5900, 5904, 5944, 5948, 5973`; `src/modules/constructor.bas:1901, 5392-5393, 5434` |
| `TbProyectosEdiciones_Riesgos` | `Gestion_Riesgos_Datos` | lectura | `src/modules/FUNCIONES UTILES.bas:6145` |
| `TbUsuarios_HPS` | `HPST` | lectura | `src/classes/SuministradorOperaciones.cls:267-270`; `src/modules/FUNCIONES UTILES.bas:2442, 2500` |
| `TbUsuariosHistoricos_HPS` | `HPST` | lectura | `src/classes/SuministradorOperaciones.cls:333-336`; `src/modules/FUNCIONES UTILES.bas:2440, 2498` |
| `TbProyectos_AGEDYS` | `AGEDYS_DATOS` (copia local) | lectura | `src/modules/FUNCIONES UTILES.bas:2346`; `src/modules/Instalador.bas:714` |
| `TbExpedientes1` | `AGEDYS_DATOS` (ruta de producción) | lectura | `src/modules/constructor.bas:1900-1901`; `src/modules/Instalador.bas:713-715` |
| `TbCorreosEnviados` | `Correos_datos` | lectura y escritura | `src/classes/Correo.cls:45`; `src/classes/CorreoOperaciones.cls:46` |

Dos hallazgos de fondo:

**El front escribe en tablas de otros backends.** `TbProyectos_Riesgos` recibe seis `UPDATE` (`src/modules/FUNCIONES UTILES.bas:5723, 5900, 5904, 5944, 5948` y `src/classes/GestionRiesgos.cls:528`), `TbNoConformidades` recibe dos (`src/modules/FUNCIONES UTILES.bas:6456, 6459`) y `TbCorreosEnviados` un `INSERT` (`src/classes/CorreoOperaciones.cls:46`). Son tres tablas cuya dueña es otra aplicación: R-2 exige decidir quién es el dueño y qué ocurre con la escritura cuando Expedientes se apague. `RegistrarCorreoRAC`, que hace el `UPDATE` de `CorreoRAC`, no tiene llamadores estáticos (CodeGraph: «No callers found»); eso no lo convierte en código muerto, porque el resto de las escrituras sí es alcanzable.

**El vínculo `TbExpedientes1` apunta al AGEDYS de producción.** En el backend, `TbExpedientes1` está vinculada a `\\datoste\aplicaciones_dys\Aplicaciones PpD\Proyectos\AGEDYS_DATOS.accdb`, mientras las otras catorce tablas de AGEDYS lo están a la copia local `C:\00repos\datos\AGEDYS_DATOS.accdb`. En este entorno el vínculo no resuelve («no es una ruta de acceso válida»), de modo que la lectura no devuelve datos; donde la ruta sí existe, el front lee el AGEDYS de producción. La cadena viva es `ExpedienteOperaciones.EliminarResponsable` → `Entorno.ColExpedientesEnGestionDeRiesgos` (`src/classes/Entorno.cls:1454`) → `constructor.getExpedientesEnGestionDeRiesgos` (`src/modules/constructor.bas:1900-1901`), que une `TbExpedientes1` con `TbProyectos_Riesgos`.

El front ya tiene la pieza que falta para neutralizar ese vínculo, pero no la usa: `modOrigenDatos` calcula, para cada vínculo, la ruta equivalente en la carpeta del backend activo (`src/modules/modOrigenDatos.bas:139-160`), de modo que en modo prueba un vínculo a producción se reorienta a la carpeta de pruebas. Ni `OrigenDatos_RutaVinculada` ni `OrigenDatos_RutaOrigen` tienen llamadores en `src/`: hoy solo se usa el etiquetado del título y del splash. El mecanismo, por tanto, no protege el vínculo `TbExpedientes1`.

## Backend `Expedientes_datos`

| Hecho | Valor |
|---|---|
| Tablas | 71: 47 propias y 24 vinculadas |
| Relaciones | 15, de las que 14 reales: 13 desde `TbExpedientes` y una desde `TbSuministradores`. La restante es `MSysNavPaneGroupCategories` → `MSysNavPaneGroups`, del sistema |
| Consultas guardadas | 11 (`Consulta1`…`Consulta11`). El extractor conserva nombre y tipo declarado por DAO (48 en `Consulta1`, `Consulta3` y `Consulta4`; 0 en las otras ocho), no el SQL |
| Tablas sin clave primaria | 6: `E2E_export_20260723_142133_lista`, `TbCorreosEnviados`, `TbExpAgedys`, `TbExpAGEDYS1`, `TbExpedientes1`, `TbFacturasPrincipal` |
| Índices únicos no primarios | 47. Entre los que sostienen la unicidad: `TbExpedientesConEntidades(IDExpediente)`, `TbExpedientes(Unico: OrdinalE2E)` y `TbResponsablesPorRol(IDUsuario, Rol)` |
| Reglas de validación | 0 en todo el backend |
| Valores por defecto | 41 campos |
| Campos multivalor, de datos adjuntos u objeto OLE | 0. Los tipos presentes son booleano, entero, entero largo, moneda, fecha/hora, texto y memo, y ningún campo va marcado como complejo |
| Vínculos | 24: 15 a `AGEDYS_DATOS.accdb` (14 a la copia local y uno, `TbExpedientes1`, a la ruta de producción), 3 a `Gestion_Riesgos_Datos.accdb`, 2 a `Lanzadera_Datos.accdb`, 2 a `HPST.accdb`, 1 a `Correos_datos.accdb` y 1 a `NoConformidades_Datos.accdb` |
| Vínculo que no resuelve en este entorno | `TbExpedientes1` |

Precisiones sobre las seis tablas sin clave:

- `TbCorreosEnviados` (vinculada a `Correos_datos.accdb`) no tiene clave primaria ni índice único: sus dos índices, `IDCorreo` e `IDEdicion`, no son únicos. El front la escribe. Access exige un índice único para actualizar una tabla vinculada, de modo que esa ruta de escritura hay que verificarla antes de darla por buena.
- `TbFacturasPrincipal` (vinculada a `AGEDYS_DATOS.accdb`) tiene un índice único sobre `NPedido`, pero ninguna clave primaria.
- `TbExpedientes1` no es evaluable en este entorno: el vínculo roto deja la tabla sin campos ni índices legibles. Su estado real en el AGEDYS de producción no se pudo comprobar.
- `TbExpAgedys` y `TbExpAGEDYS1` (116 filas cada una, el mismo esquema) no tienen ningún índice. `TbExpAgedys` la lee el instalador (`src/modules/Instalador.bas:802`); `TbExpAGEDYS1` no la cita nadie.
- `E2E_export_20260723_142133_lista` (89 filas, `IDExpediente` y `OrdinalE2E`) solo tiene un índice no único sobre `IDExpediente`. No la cita nadie.

## Veredicto de propiedad

**Se quedan en el backend**: 36 de las 47 tablas propias las usa el VBA del front.

| Grupo | Tablas |
|---|---|
| Núcleo del expediente (16) | `TbExpedientes`, `TbExpedientesConEntidades`, `TbExpedientesAnexos`, `TbExpedientesAnualidades`, `TbExpedientesCodigoCompras`, `TbExpedientesComerciales`, `TbExpedientesCPVs`, `TbExpedientesHitos`, `TbExpedientesJefaturas`, `TbExpedientesJuridicas`, `TbExpedientesLugaresEjecucion`, `TbExpedientesModificados`, `TbExpedientesPECAL`, `TbExpedientesRACS`, `TbExpedientesResponsables`, `TbExpedientesSuministradores` |
| Catálogos (13) | `TbComerciales`, `TbCPV`, `TbEjercitos`, `TbEstados`, `TbGradosClasificacion`, `TbJuridicas`, `TbLugaresEjecucion`, `TbOficinasPrograma`, `TbOrganosContratacion`, `TbPECAL`, `TbRACS`, `TbResponsablesPorRol`, `TbSuministradores` |
| Cola y enlaces E2E (2) | `TbE2ECola`, `TbE2EEnlaces` |
| Apoyo (5) | `TbCambios`, `TbConfMostrarEstado`, `TbDatosEconomicosExpedientes`, `TbExpAgedys`, `TbUltimoCambio` |

**Las usa el front y no existen en el backend**: `TbExpedientesDatosAGEDYS`, en `src/classes/ExpedienteAGEDYS.cls:59` (`DameID` sobre `IDExpedienteAGEDYS`), `src/modules/constructor.bas:5003` y `:5007`. No está en el ERD del backend ni en el del front. O es el nombre antiguo de `TbExpAgedys` o de `TbExpAGEDYS1`, o la tabla falta: en runtime, un fallo de runtime lo dirá.

**No las usa el front**: 11 tablas propias y 14 vinculadas.

Propias (11): `Copia de TbExpedientes` (0 filas), `Copia de TbExpedientesConEntidades` (365), `E2E_export_20260723_142133_lista` (89), `ListaPrevia` (99), `TbAusExpPostAGEDO` (203), `TbAuxEstadosMartina` (143), `TbAuxNemotecnico` (0), `TbExpAGEDYS1` (116), `TbExpedientes_antes` (457), `TbExpedientesCadenaContratacion` (0) y `TbJefaturas` (3).

Vinculadas (14): `TbComunicados`, `TbDpDInformeCondicionamiento`, `TbExpedientesAnualidades1`, `TbFacturasDetalle`, `TbFacturasPrincipal`, `TbGestionRiesgos`, `TbGTVContratos`, `TbGTVContratosExpedientes`, `TbNPedido`, `TbResponsablesExpedientes`, `TbSolicitudesOfertasPrevias`, `TbSuministradoresSAP`, `TbVisadoFacturas_Nueva` y `TbVisadosGenerales`.

**Bien ubicadas por su dueño, no por este front**: `TbUsuariosAplicaciones` y `TbUsuariosAplicacionesPermisos` son de la Lanzadera; `TbNoConformidades`, de No Conformidades; `TbProyectos_Riesgos` y `TbProyectosEdiciones_Riesgos`, de Gestión de Riesgos; `TbUsuarios_HPS` y `TbUsuariosHistoricos_HPS`, de HPS; `TbProyectos_AGEDYS` y `TbExpedientes1`, de AGEDYS; `TbCorreosEnviados`, de Correos. Que este front las use no las convierte en suyas: la regla de propiedad va en sentido contrario, y por eso las tres que el front escribe necesitan decisión explícita.

Desde este front no se puede cerrar la clasificación de la épica: saber si una de las 25 tablas sin uso la usa otro front exige el grafo de vínculos de todos los backends (tarea T6). Lo que sí se puede afirmar es que el VBA de Expedientes no las necesita.

## Propuestas para después del descubrimiento global

Ninguna se aplica: la puerta de descubrimiento global sigue cerrada (#1576). Ninguna borra nada: el propietario borra en el legacy; aquí solo se cataloga la evidencia, la comprobación previa y la reversión.

| Id | Propuesta | Grupo (fase 1) | Comprobación previa | Reversión |
|---|---|---|---|---|
| E1 | Aislar el vínculo `TbExpedientes1`, que apunta al AGEDYS de producción desde el backend de desarrollo. Decidir si debe apuntar a la copia local o si el front debe dejar de leer esa tabla. | Imprescindible | Confirmar con el propietario qué entorno debe alimentar `getExpedientesEnGestionDeRiesgos`. En la copia local, comparar el conjunto de expedientes que devuelve la unión con `TbProyectos_Riesgos` antes y después de cambiar el destino del vínculo. Evaluar, en la misma decisión, si se cablea el reorientador de vínculos ya escrito y sin uso (`src/modules/modOrigenDatos.bas:139-160`) o si se retira por superfluo. | Restaurar la cadena de conexión original del vínculo. |
| E2 | Declarar clave primaria en las seis tablas que no la tienen: `TbCorreosEnviados`, `TbFacturasPrincipal`, `TbExpAgedys`, `TbExpAGEDYS1`, `TbExpedientes1` y `E2E_export_20260723_142133_lista`. Tres de ellas (`TbCorreosEnviados`, `TbFacturasPrincipal` y `TbExpedientes1`) son de otros backends: la clave se declara en su backend, no en este. | Imprescindible | Contar duplicados en cada clave candidata antes de fijarla. Para `TbCorreosEnviados`, cuyo `IDCorreo` no es único, medir los duplicados reales y decidir qué índice único puede declarar `Correos_datos`. Para `TbExpedientes1`, resolver antes E1, porque hoy no se leen ni sus campos ni sus índices. | Quitar la clave. |
| E3 | Resolver las tres escrituras del front en tablas de otro backend: `TbProyectos_Riesgos` (seis `UPDATE`), `TbNoConformidades` (dos `UPDATE`) y `TbCorreosEnviados` (un `INSERT`). Cada una necesita un dueño y una regla de escritura explícita antes de que la web coexistir. | Imprescindible | Medir con el propietario de cada aplicación si la tabla se sigue escribiendo desde su propio front o si la escritura de Expedientes es la única. Comprobar que el SQL que se migre no introduce una escritura en un motor distinto del de su lector. | Devolver el `UPDATE` o el `INSERT` a su forma actual y al backend que declare su dueño. |
| E4 | Obtener el SQL de las once consultas guardadas del backend y mapear sus tablas. Es la condición previa de E5: sin ese mapa no se puede afirmar que ninguna consulta guardada usa una de las 25 tablas candidatas. | Imprescindible | El SQL no se pudo leer por permisos: `MSysQueries` responde «no tiene permiso para READ» (dysflow, lectura, modo `read`). Repetir la lectura con un usuario con permiso o extraer las definiciones por otra vía. | No aplica: es una lectura, no un cambio. |
| E5 | Poner en cuarentena las 25 tablas que el VBA del front no usa: renombrarlas a `zzBAJA_<tabla>`. No se borra ninguna. | Imprescindible | Confirmar que ningún otro front ni la web las leen (tarea T6) y que ninguna de las once consultas guardadas las usa (E4). Para las copias (`Copia de TbExpedientes`, `Copia de TbExpedientesConEntidades`, `TbExpedientes_antes`, `TbAuxEstadosMartina`), confirmar antes que no son la copia de trabajo de un proceso de importación o de una persona. | Quitar el prefijo. |
| E6 | Retirar del front los cuatro vínculos que su código no usa: `Copia de TbExpedientes`, `Copia de TbExpedientesConEntidades`, `ListaPrevia` y `TbAusExpPostAGEDO`. Los cuatro apuntan al propio backend, que el front ya abre con `getdb()`, y la regla del propietario sitúa los vínculos de otras aplicaciones en el backend. | Imprescindible | Confirmar que ningún formulario, informe, macro ni herramienta externa del front los usa. Los formularios están desconectados y no tienen `RecordSource`, de modo que el riesgo está fuera de `src/`. | Recrear los vínculos con la misma definición. |
| E7 | Reconciliar `TbExpedientesDatosAGEDYS`: el código la cita en `src/classes/ExpedienteAGEDYS.cls:59` y `src/modules/constructor.bas:5003, 5007`, y no está en el ERD del backend. | Imprescindible | Buscar en el backend de producción si existe con ese nombre y, si no, decidir si el código debe apuntar a `TbExpAgedys` o a `TbExpAGEDYS1`. Comprobar en ambas la columna que usa el `DameID` (`IDExpedienteAGEDYS`). | Sustituir el nombre en el código por el que resulte ser el correcto. |
| E8 | Sustituir el generador `DameID` (máximo más uno) por una identidad. Coste medido: una función (`src/modules/FUNCIONES UTILES.bas:2863-2891`), 30 llamadas en 30 ficheros de producción y 27 pares tabla-campo distintos, más cinco `Max` directos (`src/modules/E2EEnlacesService.bas:110`, `src/modules/E2EFamilyExportHelper.bas:303`, `src/modules/ManualE2EExport.bas:173`, `src/modules/FUNCIONES UTILES.bas:1738, 1754`). | Imprescindible | Comprobar que la tabla destino acepta la identidad y que el alta no necesita el identificador antes de insertar la fila padre: en el alta, el identificador se calcula y se asigna antes de escribir la cabecera (`src/classes/ExpedienteOperaciones.cls:84-86`). Revisar el par que cae en una tabla de otro backend, `TbCorreosEnviados`. | Restaurar `DameID` y los `Max` directos. |
| E9 | Declarar la integridad referencial que falta hacia `TbUsuariosAplicaciones`: `TbExpedientes.IDResponsableCalidad`, `IDResponsableSeguridad`, `IDUsuarioCreacion` e `IDUsuarioUltimoCambio` se unen en SQL (`src/modules/modBusquedaVivo.bas:66, 203`, `src/modules/modInformeVivo.bas:63-64`, `src/modules/constructor.bas:1640, 2607-2609, 5702-5703`) pero no hay ninguna relación declarada en el ERD. | Recomendable | Contar los responsables huérfanos (el caso del usuario borrado ya está documentado en el roadmap de la RC) y comprobar que el tipo de `TbUsuariosAplicaciones.Id` coincide con el de la columna referenciada. | Deshacer la relación y volver a la unión en SQL. |
| E10 | Decidir el destino de `TbAuxEstadosMartina`, duplicada en el front y en el backend con las mismas 143 filas y tipos distintos en `ID`. Ningún código la cita. | Recomendable | Confirmar que ninguna herramienta externa consume la copia del front y que la tabla no es trabajo en curso de una persona: el nombre y las columnas apuntan a estados de desarrollo con nombre propio. | Devolver la copia al front. |
| E11 | Revisar el `RowSource` residual de `Form_FormE2ECola` y las rutinas dormidas del `Instalador`, que borran y copian datos sin llamadores: `RellenoDeDatosDeAGEDYS` borra los expedientes de prueba y los repuebla desde `TbExpedientes1` (`src/modules/Instalador.bas:708-715`), y `RellenarOficinasProgramas` ejecuta `DELETE * FROM TbExpedientes;` (`src/modules/Instalador.bas:799`) antes de copiar desde `TbExpAgedys` (`:802`). | Recomendable | Confirmar con el propietario si el instalador se conserva como herramienta. Si se conserva, que no pueda invocarse por accidente; si no, retirarlo en la fase 1 junto con el formulario. | Restaurar las rutinas y el `RowSource`. |

Lo que **no** se toca ahora: los 24 vínculos del backend, las 14 relaciones reales, las 36 tablas que el front usa y la tabla local de configuración del front.

## Lo que no se pudo verificar

1. El SQL de las once consultas guardadas del backend: `MSysQueries` no es legible con el usuario disponible («no tiene permiso para READ»), de modo que no se pudo mapear qué tablas tocan.
2. El estado real de `TbExpedientes1` (clave, índices y columnas): su vínculo a la ruta de producción no resuelve en este entorno.
3. Si alguna de las 25 tablas sin uso en el front la consume otro front o la web: exige el grafo de vínculos de todos los backends (tarea T6).
4. Si las copias (`Copia de TbExpedientes`, `Copia de TbExpedientesConEntidades`, `TbExpedientes_antes`) o `ListaPrevia` y `TbAusExpPostAGEDO` las usa una herramienta fuera de `src/` (Excel, informes programados, la web): no hay forma de comprobarlo desde el código.
5. La correspondencia entre el binario del front del worktree, que tiene cambios sin commitear, y el commit `e71d9e7`: el ERD del front refleja el binario, no el commit.
6. El estado de la UAT en el tracker: la evidencia es el roadmap del repo, no la issue de UAT.
7. Las diferencias `src/` línea a línea entre la RC1 y `main`: se comprobó el recuento de ficheros y el diff agregado, no el contenido de los 138 módulos de producción.
