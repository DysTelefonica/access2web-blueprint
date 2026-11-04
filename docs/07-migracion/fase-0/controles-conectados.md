# Fase 0: controles conectados del legacy

Épica `docs/07-migracion/epica-coexistencia-postgresql.md` (#1545), fase 0. Anexo de los controles con origen de datos de los siete fronts del ecosistema. No sustituye a las fichas de inventario por aplicación ([`brass.md`](brass.md), [`expedientes.md`](expedientes.md), [`gestion-riesgos.md`](gestion-riesgos.md), [`hps.md`](hps.md), [`condor.md`](condor.md), [`lanzadera-datos.md`](lanzadera-datos.md), [`aperturas-directas-lanzadera.md`](aperturas-directas-lanzadera.md)) ni al [grafo de dependencias entre backends](grafo-dependencias.md): aquí solo se reúne el SQL literal que los controles conectados ejecutan.

Un **control conectado** es un combo, una lista o un formulario cuyo origen de datos es una tabla o una consulta en lugar de una lista de valores fija; importa aquí porque es el punto donde el legacy ata su interfaz a un esquema concreto, y porque es el inventario que decide qué SQL tiene que sobrevivir a PostgreSQL tal cual y qué SQL hay que reescribir antes del corte.

Inventario de solo lectura. No se aplica nada en el legacy.

## Método y bases de análisis

Las fuentes son los exports de formulario (`src/forms/*.form.txt`, que incluyen el code-behind) y los módulos VBA (`src/forms/*.cls`, `src/classes/*.cls`, `src/modules/*.bas`) de cada worktree. Los patrones buscados en el export son `RowSourceType ="Table/Query"` o `"Table/View/StoredProc"` con su `RowSource` inmediatamente después, `RecordSource` a nivel de formulario y `ControlSource` con valor; en el código, las asignaciones a `.RowSource`, `.RecordSource` y `.ControlSource`.

| Aplicación | Base analizada | Worktree | SHA |
|---|---|---|---|
| Brass | `2026-004.rc1` | `00_BRASS-worktrees/2026-004.rc1` | `cb2795128748fa91e79b5cf0a0c2501902692e3f` |
| Expedientes | `release/2026-008.2.rc1` | `00_EXPEDIENTES-worktrees/2026-008.2.rc1` | `e71d9e7f57512184c6f1c03b94fc075ff07f6a85` |
| Gestión de Riesgos | `release/2026-008.2.rc1` | `00_GESTION_RIESGOS-worktrees/2026-008.2.rc1` | `f457571c7d1d2b4eaffef48f0ab8e798553c9355` |
| No Conformidades | `release/2026-009.rc1` | `00_NO_CONFORMIDADES-worktrees/2026-009.rc1` | `f281b4e3ad33e5411ce6e0e6f59ad388698a10b1` |
| Condor | `release/2026-007.rc1` | `00_CONDOR-worktrees/2026-007.rc1` | `09e432ac48786bc2b3b88764239791ced5bf0d27` |
| HPS | `2026-005-rc1` | `00_HPS-worktrees/2026-005-rc1` | no verificable (ver «Lo que no se pudo verificar») |
| Registros | `main` | `00_REGISTROS` | `44f36d3774a4870e7f70d8c8aa71710965f61411` |

Las referencias son relativas al `src/` del worktree de cada aplicación: `src/forms/<formulario>.form.txt:<línea>` para lo puesto en diseño y `src/forms/<formulario>.cls:<línea>` (o `src/modules/<módulo>.bas:<línea>`) para lo asignado en tiempo de ejecución.

Los siete worktrees existen y tienen `src/forms/`. Ninguno tiene macros ni `AutoExec`: la única aparición de `AutoExec` en todo el alcance es un comentario (`src/modules/M_E2EMigration.bas:10`, Expedientes).

## Resumen por aplicación

| Aplicación | Formularios con origen de datos | Controles conectados | SQL distintas | Tablas que aparecen |
|---|---|---|---|---|
| Brass | 41 | 106 | 72 | 31 |
| Registros | 7 | 37 | 17 | 10 |
| HPS | 3 | 3 | 3 | 3 |
| Gestión de Riesgos | 1 | 1 | 1 | 1 |
| Expedientes | 1 | 1 | 0 (el `RowSource` de diseño está vacío) | 0 |
| No Conformidades | 0 | 0 | 0 | 0 |
| Condor | 0 | 0 | 0 | 0 |
| **Total** | **53** | **148** | **93** | **45** |

Los tres controles de HPS son formularios: los tres tienen `RecordSource` a nivel de formulario, no controles con `RowSource`. El total de SQL no cuenta el `RowSource` vacío de Expedientes. El total de tablas son las del diseño; el SQL de tiempo de ejecución añade cinco más, hasta 50, y `TbRiesgos`, que se consulta pero no llega a ningún control, según el [catálogo por tabla](#catálogo-de-sql-por-tabla).

### Un hallazgo que corrige la regla del propietario

La regla «salvo Brass, los formularios son desconectados» se confirma solo en parte. La medición la refuta en dos aplicaciones:

- **Registros** tiene 37 controles con `RowSourceType ="Table/Query"` repartidos en 7 formularios. La propiedad está en el formulario, puesta en diseño.
- **HPS** tiene 3 formularios unidos a recordset (`RecordSource` a nivel de formulario). No es una lista de valores: es SQL.

Lo que sí se confirma para el resto es que los formularios no tienen `RecordSource` y que sus combos y listas se rellenan en tiempo de ejecución con SQL ejecutado sobre `getdb()`, no con una propiedad de diseño. La excepción de Brass es mayor de lo que la regla asumía: Brass no solo tiene combos y listas con origen de datos, sino que es la aplicación donde ese origen de datos está en 41 de sus 88 formularios.

## Controles conectados por aplicación

### Brass

El detalle de Brass, con sus 106 controles, va en un documento aparte por volumen: `controles-conectados-brass.md`.

### Registros

17 SQL distintas sobre 10 tablas del backend `Registro_Datos`, más dos tablas de otros backends: `TbExpedientes` (Expedientes) y `TbProyectos` (Gestión de Riesgos).

| SQL literal | Usos | Tablas | Tipo de origen |
|---|---|---|---|
| `SELECT DISTINCT TbEntradas_RES.Remitente FROM TbEntradas_RES WHERE ((Not (TbEntradas_RES.Remitente) Is Null)) ORDER BY TbEntradas_RES.Remitente;` | Form_FormRegistroEntradaAlta / ComboRemitente (`src/forms/Form_FormRegistroEntradaAlta.form.txt:1475`) | TbEntradas_RES | Table/Query |
| `SELECT DISTINCT tbUsuarios.UsuarioRed, tbUsuarios.Nombre, tbUsuarios.DirCorreo FROM tbUsuarios WHERE (((tbUsuarios.FechaBaja) Is Null)) ORDER BY tbUsuarios.Nombre;` | Form_FormRegistroEntradaAlta / ComboInteresados (`src/forms/Form_FormRegistroEntradaAlta.form.txt:1879`) | tbUsuarios | Table/Query |
| `SELECT distinct TbSalidas_RES.UbicacionFisica FROM TbSalidas_RES WHERE ((Not (TbSalidas_RES.UbicacionFisica) Is Null)) UNION SELECT DISTINCT TbEntradas_RES.UbicacionFisica FROM TbEntradas_RES WHERE ((Not (TbEntradas_RES.UbicacionFisica) Is Null));` | Form_FormRegistroEntradaAlta / ComboUbicacionFisica (`src/forms/Form_FormRegistroEntradaAlta.form.txt:2069`)<br>Form_FormRegistroEntradaEdicion / ComboUbicacionFisica (`src/forms/Form_FormRegistroEntradaEdicion.form.txt:1698`)<br>Form_FormRegistroSalidaAlta / ComboUbicacionFisica (`src/forms/Form_FormRegistroSalidaAlta.form.txt:2027`)<br>Form_FormRegistroSalidaEdicion / ComboUbicacionFisica (`src/forms/Form_FormRegistroSalidaEdicion.form.txt:1765`)<br>Form_FormRegistrosGestion / ComboUbicacionFisica (`src/forms/Form_FormRegistrosGestion.form.txt:4877`) | TbSalidas_RES, TbEntradas_RES | Table/Query |
| `SELECT DISTINCT TbUsuariosAplicaciones.UsuarioRed, TbUsuariosAplicaciones.Nombre, TbUsuariosAplicaciones.CorreoUsuario FROM TbUsuariosAplicaciones WHERE (((TbUsuariosAplicaciones.FechaBaja) Is Null)) ORDER BY TbUsuariosAplicaciones.Nombre;` | Form_FormRegistroEntradaAlta / ComboDestinatarios (`src/forms/Form_FormRegistroEntradaAlta.form.txt:2265`)<br>Form_FormRegistroEntradaEdicion / ComboDestinatarios (`src/forms/Form_FormRegistroEntradaEdicion.form.txt:1894`)<br>Form_FormRegistroInteresados / ComboInteresados (`src/forms/Form_FormRegistroInteresados.form.txt:1183`)<br>Form_FormRegistroSalidaAlta / ComboInteresados (`src/forms/Form_FormRegistroSalidaAlta.form.txt:1838`)<br>Form_FormRegistroSalidaAlta / ComboRemitente (`src/forms/Form_FormRegistroSalidaAlta.form.txt:2260`)<br>Form_FormRegistroSalidaEdicion / ComboRemitente (`src/forms/Form_FormRegistroSalidaEdicion.form.txt:2167`) | TbUsuariosAplicaciones | Table/Query |
| `SELECT TbExpedientes.Expediente AS EXP, TbExpedientes.TITULOEXP AS TÍTULO FROM TbExpedientes ORDER BY TbExpedientes.IdExpediente DESC;` | Form_FormRegistroEntradaAlta / ComboExp (`src/forms/Form_FormRegistroEntradaAlta.form.txt:2332`)<br>Form_FormRegistroEntradaEdicion / ComboExp (`src/forms/Form_FormRegistroEntradaEdicion.form.txt:2035`)<br>Form_FormRegistroSalidaAlta / ComboExp (`src/forms/Form_FormRegistroSalidaAlta.form.txt:2325`)<br>Form_FormRegistroSalidaEdicion / ComboExp (`src/forms/Form_FormRegistroSalidaEdicion.form.txt:1728`)<br>Form_FormRegistrosGestion / ComboExpedienteBusqueda (`src/forms/Form_FormRegistrosGestion.form.txt:3995`) | TbExpedientes | Table/Query |
| `SELECT DISTINCT TbEntradas_RES.ExpActividad FROM TbEntradas_RES WHERE (((TbEntradas_RES.ExpActividad) Is Not Null) AND ((TbEntradas_RES.Juridica)='TSOL')) UNION SELECT DISTINCT TbSalidas_RES.ExpActividad FROM TbSalidas_RES WHERE (((TbSalidas_RES.ExpActividad) Is Not Null) AND ((TbSalidas_RES.Juridica)='TSOL'));` | Form_FormRegistroEntradaAlta / ComboProyecto (`src/forms/Form_FormRegistroEntradaAlta.form.txt:2368`)<br>Form_FormRegistroEntradaEdicion / ComboProyecto (`src/forms/Form_FormRegistroEntradaEdicion.form.txt:2071`)<br>Form_FormRegistroSalidaAlta / ComboProyecto (`src/forms/Form_FormRegistroSalidaAlta.form.txt:2361`)<br>Form_FormRegistroSalidaEdicion / ComboProyecto (`src/forms/Form_FormRegistroSalidaEdicion.form.txt:2034`)<br>Form_FormRegistrosGestion / ComboActividad (`src/forms/Form_FormRegistrosGestion.form.txt:4035`) | TbEntradas_RES, TbSalidas_RES | Table/Query |
| `SELECT TbProyectos.CODPROYECTOS FROM TbProyectos ORDER BY TbProyectos.ID_PROYECTOS DESC;` | Form_FormRegistroEntradaAlta / ComboNDPD (`src/forms/Form_FormRegistroEntradaAlta.form.txt:2435`)<br>Form_FormRegistroEntradaEdicion / ComboNDPD (`src/forms/Form_FormRegistroEntradaEdicion.form.txt:2138`)<br>Form_FormRegistroSalidaAlta / ComboNDPD (`src/forms/Form_FormRegistroSalidaAlta.form.txt:2428`)<br>Form_FormRegistroSalidaEdicion / ComboNDPD (`src/forms/Form_FormRegistroSalidaEdicion.form.txt:2101`) | TbProyectos | Table/Query |
| `SELECT DISTINCT TbEntradas_RES.Remitente FROM TbEntradas_RES WHERE ((Not (TbEntradas_RES.Remitente) Is Null));` | Form_FormRegistroEntradaEdicion / ComboRemitente (`src/forms/Form_FormRegistroEntradaEdicion.form.txt:1468`) | TbEntradas_RES | Table/Query |
| `SELECT DISTINCT TbSalidas_RES.Destinatario FROM TbSalidas_RES WHERE ((Not (TbSalidas_RES.Destinatario) Is Null));` | Form_FormRegistroSalidaAlta / ComboDestinatarios (`src/forms/Form_FormRegistroSalidaAlta.form.txt:2220`) | TbSalidas_RES | Table/Query |
| `SELECT TbSalidas.Destinatario FROM TbSalidas WHERE ((Not (TbSalidas.Destinatario) Is Null)) ORDER BY TbSalidas.Destinatario;` | Form_FormRegistroSalidaEdicion / ComboDestinatarios (`src/forms/Form_FormRegistroSalidaEdicion.form.txt:1470`) | TbSalidas | Table/Query |
| `SELECT Salidas.IDSalida, Salidas.NSalida, 'Salida' AS Origen, Salidas.Extracto, Salidas.Expediente, Salidas.Anexo FROM Salidas LEFT JOIN TbSalidas_Indexacion ON Salidas.IDSalida = TbSalidas_Indexacion.IDSalida WHERE (((Salidas.NSalida) Like '*' Or (Salidas.NSalida) Is Null) AND ((Salidas.Extracto) Like '*' Or (Salidas.Extracto) Is Null) AND ((TbSalidas_Indexacion.TextoArchivo) Like '*' Or (TbSalidas_Indexacion.TextoArchivo) Is Null)) ORDER BY Salidas.IDSalida DESC;` | Form_FormRegistrosAntiguosBusqueda / ListaFiltrados (`src/forms/Form_FormRegistrosAntiguosBusqueda.form.txt:960`) | Salidas, TbSalidas_Indexacion | Table/Query |
| `SELECT [NEntrada] & '/' & [Entradas].[Año] AS NumRegistro FROM Entradas UNION SELECT [NSalida] & '/' & [Salidas].[Año] AS NumRegistro FROM Salidas;` | Form_FormRegistrosAntiguosBusqueda / ComboNregistroBusqueda (`src/forms/Form_FormRegistrosAntiguosBusqueda.form.txt:2414`) | Entradas, Salidas | Table/Query |
| `SELECT TbEntradas_RES.IDEntrada, TbEntradas_RES.Juridica, 'Entrada' AS TipoRegistro, TbEntradas_RES.NEntrada, TbEntradas_RES.Asunto, TbEntradas_RES.ExpActividad, TbEntradas_RES.NRef, TbEntradas_RES.NDPD, TbEntradas_RES.NombreAnexo, TbEntradas_RES.Encriptado FROM TbEntradas_RES WHERE (((TbEntradas_RES.ExpActividad) Like '*' Or (TbEntradas_RES.ExpActividad) Is Null) AND ((TbEntradas_RES.Juridica) Like '*' Or (TbEntradas_RES.Juridica) Is Null) AND ((TbEntradas_RES.Asunto) Like '*' Or (TbEntradas_RES.Asunto) Is Null) AND ((TbEntradas_RES.Verificado) Like '*' Or (TbEntradas_RES.Verificado) Is Null) AND ((TbEntradas_RES.NOTAS) Like '*' Or (TbEntradas_RES.NOTAS) Is Null) AND ((TbEntradas_RES.ExpActividad) Like '*' Or (TbEntradas_RES.ExpActividad) Is Null) AND ((TbEntradas_RES.NRef) Like '*' Or (TbEntradas_RES.NRef) Is Null) AND ((TbEntradas_RES.NDPD) Like '*' Or (TbEntradas_RES.NDPD) Is Null) AND ((TbEntradas_RES.NEntrada) Like '*' Or (TbEntradas_RES.NEntrada) Is Null)AND ((TbEntradas_RES.UbicacionFisica) Like '*' Or (TbEntradas_RES.UbicacionFisica) Is Null) ) ORDER BY TbEntradas_RES.IDEntrada DESC;` | Form_FormRegistrosGestion / ListaFiltrados (`src/forms/Form_FormRegistrosGestion.form.txt:941`) | TbEntradas_RES | Table/Query |
| `SELECT  TbSalidas_RES.NSalida FROM TbSalidas_RES UNION SELECT  TbEntradas_RES.NEntrada FROM TbEntradas_RES;` | Form_FormRegistrosGestion / ComboNregistroBusqueda (`src/forms/Form_FormRegistrosGestion.form.txt:3566`) | TbSalidas_RES, TbEntradas_RES | Table/Query |
| `SELECT DISTINCT TbEntradas_RES.NRef FROM TbEntradas_RES WHERE ((Not (TbEntradas_RES.NRef) Is Null)) UNION SELECT DISTINCT TbSalidas_RES.NRef FROM TbSalidas_RES WHERE ((Not (TbSalidas_RES.NRef) Is Null));` | Form_FormRegistrosGestion / ComboNREF (`src/forms/Form_FormRegistrosGestion.form.txt:4295`) | TbEntradas_RES, TbSalidas_RES | Table/Query |
| `SELECT DISTINCT TbEntradas_RES.NDPD FROM TbEntradas_RES WHERE ((Not (TbEntradas_RES.NDPD) Is Null)) UNION SELECT DISTINCT TbSalidas_RES.NDPD FROM TbSalidas_RES WHERE ((Not (TbSalidas_RES.NDPD) Is Null));` | Form_FormRegistrosGestion / ComboNDPD (`src/forms/Form_FormRegistrosGestion.form.txt:4620`) | TbEntradas_RES, TbSalidas_RES | Table/Query |
| `SELECT DISTINCT Year([FSalida]) AS AñoSalida FROM TbSalidas_RES ORDER BY Year([FSalida]) DESC;` | Form_FormRegistrosGestion / ComboAñoBusqueda (`src/forms/Form_FormRegistrosGestion.form.txt:5353`) | TbSalidas_RES | Table/Query |

Los SQL de Registros se exportan con los saltos de línea reales incrustados como `\015\012`; en la tabla se han sustituido por un espacio para que la SQL sea legible. Es un cambio de presentación, no de contenido.

`TbUsuariosAplicaciones` está en `Lanzadera_Datos` y `tbUsuarios` no aparece en el grafo de vínculos de Registros: es la tabla propia de usuario del front.

### Gestión de Riesgos

| SQL literal | Usos | Tablas | Tipo de origen |
|---|---|---|---|
| `SELECT TbAuxPriorizacion.IDRiesgo, TbAuxPriorizacion.Pri, TbAuxPriorizacion.PriEdAnterior, TbAuxPriorizacion.Codigo, TbAuxPriorizacion.Descripcion, TbAuxPriorizacion.CausaRaiz FROM TbAuxPriorizacion;` | Form_FormRiesgosEstablecerPrioridades / ListaRiesgos (`src/forms/Form_FormRiesgosEstablecerPrioridades.form.txt:708`) | TbAuxPriorizacion | Table/Query |

`TbAuxPriorizacion` es una tabla auxiliar de trabajo del backend de Gestión de Riesgos.

### HPS

Los tres controles conectados de HPS son formularios unidos a recordset: el `RecordSource` está en el formulario, no en un control.

| SQL literal | Usos | Tablas | Tipo de origen |
|---|---|---|---|
| `SELECT TbDatosLocal.* FROM TbDatosLocal;` | Form_FormInicial03ConsultasDatos / formulario (`src/forms/Form_FormInicial03ConsultasDatos.form.txt:289`) | TbDatosLocal | RecordSource |
| `SELECT TbUsuariosSICALocal.* FROM TbUsuariosSICALocal;` | Form_FormInicial07UsuariosSICADatos / formulario (`src/forms/Form_FormInicial07UsuariosSICADatos.form.txt:68`) | TbUsuariosSICALocal | RecordSource |
| `SELECT TbUsuariosHistoricosLocal.* FROM TbUsuariosHistoricosLocal;` | Form_FormInicial08UsuariosHistoricosDatos / formulario (`src/forms/Form_FormInicial08UsuariosHistoricosDatos.form.txt:83`) | TbUsuariosHistoricosLocal | RecordSource |

Las tres son tablas **de la caché local del front**, no del backend `HPST`: `TbDatosLocal` la puebla `src/modules/cacheUsuario.bas:113` (HPS) y la audita `src/modules/CacheConsistencyAudit.bas:184` (HPS); `TbUsuariosSICALocal`, `src/modules/Constructor.bas:1793` (HPS) y `CacheConsistencyAudit.bas:237` (HPS); `TbUsuariosHistoricosLocal`, `src/modules/Constructor.bas:1353` (HPS). HPS es, por tanto, la única aplicación con formularios unidos a recordset, pero la unión es contra la caché local y no contra el backend.

### Expedientes

| SQL literal | Usos | Tablas | Tipo de origen |
|---|---|---|---|
| (vacío) | Form_FormE2ECola / ListaCola (`src/forms/Form_FormE2ECola.form.txt:289`) | ninguna | Table/View/StoredProc |

`ListaCola` declara `RowSourceType ="Table/View/StoredProc"` con `RowSource` vacío: el SQL llega en tiempo de ejecución (ver el apartado siguiente). La ficha de Expedientes ya lo señala.

### No Conformidades y Condor

Ningún control con origen de datos en tabla o consulta. Los 76 `RowSourceType ="Value List"` de No Conformidades y los 19 de Condor son listas de valores literales; su contenido se asigna en tiempo de ejecución y no consulta ninguna tabla.

## Orígenes de datos asignados en tiempo de ejecución

### Diseño frente a tiempo de ejecución

Lo puesto en el formulario es contrato estable: el SQL está en el `.form.txt`, se puede leer, comparar y migrar. Lo asignado en tiempo de ejecución es más frágil: el SQL puede cambiar en cada versión del código y depende de que la rutina que lo compone se mantenga.

Las asignaciones en `.cls` y `.bas` se cuentan así:

| Aplicación | Asignaciones a `RowSource` / `RecordSource` / `ControlSource` | Con SQL literal | Con lista de valores literal | Con variable | Con llamada a función |
|---|---|---|---|---|---|
| Brass | 110 | 0 | 38 | 69 | 3 |
| Expedientes | 83 | 0 | 80 | 3 | 0 |
| Gestión de Riesgos | 60 | 0 | 43 | 9 | 8 |
| No Conformidades | 42 | 0 | 34 | 8 | 0 |
| HPS | 32 | 2 | 26 | 4 | 0 |
| Condor | 19 | 4 | 8 | 6 | 1 |
| Registros | 7 | 0 | 2 | 5 | 0 |

La columna «con lista de valores literal» domina casi todas las aplicaciones: en Expedientes, 80 de las 83 asignaciones rellenan el control con una lista fija de valores separados por `;`. Esas no consultan ninguna tabla y no acotan el esquema.

### Brass

El detalle de Brass en tiempo de ejecución va en `controles-conectados-brass.md`.

### Registros

| SQL en tiempo de ejecución | Asignaciones | Tablas |
|---|---|---|
| `SELECT <tabla>.* FROM <tabla> WHERE (((<tabla>.<campo>)=<id>));` con `<tabla>` y `<campo>` tomados de variables | `src/forms/Form_FormRegistrosGestion.cls:1091` (SQL en `:395`), `src/forms/Form_FormRegistrosAntiguosBusqueda.cls:422` (SQL en `:288`) | TbEntradas_RES o TbSalidas_RES, según el registro |
| `SELECT Entradas.NEntrada AS NumRegistro FROM Entradas UNION SELECT Salidas.NSalida AS NumRegistro FROM Salidas;` | `src/forms/Form_FormRegistrosAntiguosBusqueda.cls:289` (SQL en `:288`) | Entradas, Salidas |
| El mismo SQL con la cláusula `ORDER BY` reescrita por `DameSQLOrdenado` | `src/forms/Form_FormRegistrosGestion.cls:1221` y `src/forms/Form_FormRegistrosAntiguosBusqueda.cls:551` (el SQL de partida se toma de `Me.ListaFiltrados.RowSource`, `src/forms/Form_FormRegistrosGestion.cls:1214`) | las mismas que la consulta de la lista |

### HPS

| SQL en tiempo de ejecución | Asignaciones | Tablas |
|---|---|---|
| `SELECT TbDatosLocal.* FROM TbDatosLocal;` | `src/forms/Form_FormInicial03ConsultasDatos.cls:32` (SQL en `:29`) | TbDatosLocal |
| `SELECT TbUsuariosSICALocal.* FROM TbUsuariosSICALocal <where>` | `src/forms/Form_FormInicial07UsuariosSICA.cls:128` (SQL en `:124`), `src/forms/Form_FormInicial07UsuariosSICADatos.cls:28` | TbUsuariosSICALocal |
| `SELECT * FROM TbUsuariosHistoricosLocal;` | `src/forms/Form_FormInicial08UsuariosHistoricos.cls:140` (SQL en `:124`), `src/forms/Form_FormInicial08UsuariosHistoricosDatos.cls:28` | TbUsuariosHistoricosLocal |

Los `RecordSource` de HPS se asignan al **subformulario** embebido, no al formulario principal: `frm.RecordSource = m_SQL` (`src/forms/Form_FormInicial07UsuariosSICA.cls:128`).

### Expedientes

| SQL en tiempo de ejecución | Asignaciones | Tablas |
|---|---|---|
| `SELECT TOP <n> TbE2ECola.IDExpediente, TbE2ECola.OrdinalE2E, TbE2ECola.Estado, TbE2ECola.TipoOrigen, TbE2ECola.FechaDeteccion, TbE2ECola.FechaProcesado, TbE2ECola.IDCola FROM TbE2ECola <where> ORDER BY TbE2ECola.FechaDeteccion DESC, TbE2ECola.IDCola DESC;` | `src/forms/Form_FormE2ECola.cls:138` (SQL en `src/modules/modE2EColaPanelHelper.bas:68-77`) | TbE2ECola |
| Variante paginada: la misma con `WHERE TbE2ECola.IDCola NOT IN (SELECT TOP <skip> TbE2ECola.IDCola FROM TbE2ECola <where> ORDER BY …)` | `src/modules/modE2EColaPanelHelper.bas:78-90` | TbE2ECola |

`TbE2ECola` no está en el grafo de vínculos de Expedientes: el front no lo vincula y el control no puede resolver su origen, como ya advierte la ficha de Expedientes.

### Condor

Las cuatro asignaciones con SQL literal de Condor son las únicas de todo el alcance que consultan una tabla de sistema de Access.

| SQL en tiempo de ejecución | Asignaciones | Tablas |
|---|---|---|
| `SELECT '(Todos)' AS U FROM MSysObjects UNION SELECT DISTINCT usuario FROM tbLogCambios ORDER BY U;` | `src/forms/frmVerLogCambios.cls:16` | MSysObjects, tbLogCambios |
| `SELECT '(Todos)' AS T FROM MSysObjects UNION SELECT DISTINCT tabla FROM tbLogCambios ORDER BY T;` | `src/forms/frmVerLogCambios.cls:20` | MSysObjects, tbLogCambios |
| `SELECT '(Todos)' AS U FROM MSysObjects UNION SELECT DISTINCT usuario FROM tbLogErrores ORDER BY U;` | `src/forms/frmVerLogErrores.cls:17` | MSysObjects, tbLogErrores |
| `SELECT '(Todos)' AS M FROM MSysObjects UNION SELECT DISTINCT modulo FROM tbLogErrores ORDER BY M;` | `src/forms/frmVerLogErrores.cls:22` | MSysObjects, tbLogErrores |

`tbLogCambios` y `tbLogErrores` son tablas del front de Condor, que las consulta también desde `src/classes/LogServicio.cls:168` y `:182`.

### Gestión de Riesgos y No Conformidades

Ninguna de las dos aplicaciones asigna SQL a un control: sus 60 y 42 asignaciones son listas de valores literales o cadenas compuestas a partir de datos ya leídos (`modProyectoCuadroHelper.CabeceraCuadro`, `src/modules/modProyectoCuadroHelper.bas:133`, devuelve `"ID;Cod.Exp;Nemotécnico;Edición;Pub.;Prox Pub.;F Prev. Cierre;Fecha Cierre;R.Calidad"`).

Una excepción merece mención: `src/forms/Form_FormRiesgoDefinicion.cls:371` y `src/forms/Form_FormRiesgoDefinicionNoBiblioteca.cls:134` asignan al combo el resultado de `modRiesgoPrincipalHelper.ListaCodigosRiesgosPrincipales`, que **ejecuta** `SELECT CodigoRiesgo FROM TbRiesgos WHERE IDEdicion=…` (`src/modules/modRiesgoPrincipalHelper.bas:129`) pero devuelve el resultado como lista de valores, no como SQL. La consulta sí toca `TbRiesgos`; el control, no.

## Catálogo de SQL por tabla

Una fila por tabla, con el número de SQL distintas que la tocan (de diseño y de tiempo de ejecución) y las aplicaciones. Sirve para saber qué sobrevive a PostgreSQL tal cual: el veredicto es de portabilidad del SQL, no de diseño de esquema.

| Tabla | SQL que la tocan | Aplicación | Veredicto de portabilidad |
|---|---|---|---|
| TbEquipos | 25 | Brass | Reescritura de función: `DateDiff`, `IIf`, `Now` |
| TbEventos | 25 | Brass | Reescritura de función: `IIf` |
| TbMaterial | 23 | Brass | Reescritura de función: `IIf` |
| TbSubsistemaBui | 18 | Brass | Reescritura menor: `Like '*'`, `Not (x) = y` |
| TbActividades | 16 | Brass | Reescritura de función: `IIf` |
| TbEntradas_RES | 8 | Registros | Reescritura menor: `Is Not Null`, `Is Null`, `Like '*'` |
| TbNodoBUI | 8 | Brass | Reescritura menor: `Is Null`, `Like '*'` |
| TbPlanificacion | 7 | Brass | Reescritura de función: `DateDiff`, `IIf`, `Now` |
| TbPlanificacionEquipos | 7 | Brass | Reescritura de función: `DateDiff`, `IIf`, `Now` |
| TbSalidas_RES | 7 | Registros | Reescritura menor: `Is Not Null`, `Is Null`, `Year` |
| TbIDEquipos | 5 | Brass | Reescritura menor: `Is Null` |
| MSysObjects | 4 | Condor | Bloqueante: `MSysObjects` |
| TbGastos | 4 | Brass | Reescritura de función: `CDbl`, `Format`, `IIf`, `IsNumeric` |
| TbSubcontrataciones | 4 | Brass | Compatible tal cual |
| Salidas | 3 | Registros | Reescritura menor: `Is Null`, `Like '*'`, concatenación `&` |
| TbBUI | 3 | Brass | Reescritura menor: `Is Null` |
| TbEquiposCalibrables | 3 | Brass | Reescritura menor: `Is Null`, `Like '*'` |
| TbEquiposCalibrablesFechas | 3 | Brass | Reescritura menor: `Is Null`, `Like '*'`, alias entre corchetes `AS [...]` |
| TbFacturaActividadesInvolucradas | 3 | Brass | Reescritura de función: `IIf` |
| TbFacturaMaterialesInvolucrados | 3 | Brass | Compatible tal cual |
| TbTecnicos | 3 | Brass | Reescritura de función: `CDbl`, `Format`, `IIf`, `IsNumeric` |
| Entradas | 2 | Registros | Reescritura menor: concatenación `&` |
| TbAnexos | 2 | Brass | Reescritura de función: `CDbl`, `Format`, `IIf`, `IsNumeric` |
| TbCodActividad | 2 | Brass | Compatible tal cual |
| TbDatosLocal | 2 | HPS | Compatible tal cual |
| TbFacturaEventosInvolucrados | 2 | Brass | Reescritura de función: `IIf` |
| TbFacturaGastosInvolucrados | 2 | Brass | Reescritura de función: `CDbl`, `Format`, `IIf`, `IsNumeric` |
| TbMaterialSeguimiento | 2 | Brass | Reescritura menor: `Is Null`, alias entre corchetes `AS [...]` |
| TbTipoAccion | 2 | Brass | Reescritura menor: `Is Null` |
| TbTipoEvento | 2 | Brass | Compatible tal cual |
| TbUsuariosSICALocal | 2 | HPS | Compatible tal cual |
| tbLogCambios | 2 | Condor | Bloqueante: `MSysObjects` |
| tbLogErrores | 2 | Condor | Bloqueante: `MSysObjects` |
| TbAuxPriorizacion | 1 | Gestión de Riesgos | Compatible tal cual |
| TbCriticidad | 1 | Brass | Compatible tal cual |
| TbE2ECola | 1 | Expedientes | Bloqueante: `SELECT TOP` |
| TbExpedientes | 1 | Registros | Compatible tal cual |
| TbFacturaSubcontratacionesInvolucradas | 1 | Brass | Compatible tal cual |
| TbNodos | 1 | Brass | Compatible tal cual |
| TbOriginador | 1 | Brass | Compatible tal cual |
| TbPlanificacionAnexos | 1 | Brass | Reescritura menor: alias entre corchetes `AS [...]` |
| TbPlanificacionAñadida | 1 | Brass | Compatible tal cual |
| TbProyectos | 1 | Registros | Compatible tal cual |
| TbSalidas | 1 | Registros | Reescritura menor: `Is Null` |
| TbSalidas_Indexacion | 1 | Registros | Reescritura menor: `Is Null`, `Like '*'` |
| TbSellados | 1 | Brass | Compatible tal cual |
| TbUbicacion | 1 | Brass | Compatible tal cual |
| TbUsuariosAplicaciones | 1 | Registros | Reescritura menor: `Is Null` |
| TbUsuariosHistoricosLocal | 1 | HPS | Compatible tal cual |
| tbUsuarios | 1 | Registros | Reescritura menor: `Is Null` |
| TbRiesgos | 1 | Gestión de Riesgos | Compatible tal cual, pero la consulta no llega al control: se ejecuta con `OpenRecordset` y su resultado se pasa como lista de valores |
El catálogo suma 51 tablas: las 45 del diseño, cinco que solo aparecen en el SQL de tiempo de ejecución (`TbPlanificacionAnexos`, `TbE2ECola`, `tbLogCambios`, `tbLogErrores` y `MSysObjects`) y `TbRiesgos`, que se consulta pero no llega a ningún control.

## Riesgos

### Consultas que no compilan en PostgreSQL tal cual

52 de las 93 SQL de diseño no contienen ningún patrón incompatible y se transportan tal cual a PostgreSQL. Las otras 41 requieren al menos una reescritura; la razón común es la sintaxis de Access. Los patrones, con la frecuencia medida sobre las 93:

| Patrón | Dónde | Qué hay que cambiar |
|---|---|---|
| `IIf(...)` | 4 SQL de Brass: `Form_FormEventoGestion.form.txt:440`, `Form_FormEventoGestion.form.txt:645`, `Form_FormGastoGestion.form.txt:980` y `Form_FormPlanificacionesBusqueda.form.txt:1012` | `CASE WHEN` |
| Alias de columna entre corchetes, `AS [Fecha Inicial]` | 8 SQL de diseño (6 de Brass, entre ellas `Form_FormCalibracionGestion.form.txt:977` y `Form_FormMaterialSeguimiento.form.txt:958`; 2 de Registros) | `AS "Fecha Inicial"` |
| `Not (columna) = valor` (el «distinto» de Access) | `Form_FormActividadEdicion.form.txt:625` y `Form_FormSubSistemaAlta.form.txt:540` | `<>` |
| Literales de fecha `#01/01/1900#`, `#1/1/2300#` | `Form_FormEventoGestion.form.txt:440`, `Form_FormGastoGestion.form.txt:980` | `DATE '1900-01-01'` |
| `Between` sobre fecha | `Form_FormEventoGestion.form.txt:440` | `BETWEEN` con `timestamp` |
| `DateDiff('d', Now(), …)` | `Form_FormPlanificacionesBusqueda.form.txt:1012` | `EXTRACT`/`age()` |
| `Now()` | `Form_FormPlanificacionesBusqueda.form.txt:1012` | `CURRENT_TIMESTAMP` |
| `Is Null` / `Is Not Null` | 32 SQL de diseño (21 de Brass, 11 de Registros) | `IS NULL` |
| `Like '*'` | 6 SQL de diseño (4 de Brass: `Form_FormCalibracionEquipoGestion.form.txt:951`, `Form_FormEquipoAlta.form.txt:60` y dos más; 2 de Registros) | `ILIKE '%'` o quitar el filtro |
| `Format(CDbl(…), '#,##0.00 €')` | `Form_FormGastoGestion.form.txt:980` y su gemela de runtime (`Form_FormGastoGestion.cls:506`) | `to_char` |
| `IsNumeric` | las dos anteriores | comprobación previa en el cliente |
| Concatenación con `&` | `Form_FormRegistrosAntiguosBusqueda.form.txt:2415` | `||` |
| `Year([FSalida])` | `Form_FormRegistrosGestion.form.txt:5354` | `EXTRACT(YEAR FROM …)` |
| `SELECT TOP n` | `src/modules/modE2EColaPanelHelper.bas:68-90` | `LIMIT` |
| `MSysObjects` | `src/forms/frmVerLogCambios.cls:16` y `:20`, `src/forms/frmVerLogErrores.cls:17` y `:22` | tabla de catálogo de Access: no existe en PostgreSQL |

`MSysObjects` es el único bloqueante duro: las cuatro consultas de los filtros de los logs de Condor no tienen equivalente en PostgreSQL y hay que reescribirlas contra el catálogo de PostgreSQL o contra el propio `tbLogCambios`/`tbLogErrores`. `SELECT TOP` es un bloqueante mecánico pero fácil: `LIMIT`.

### Tablas de otros backends que aparecen en el SQL

| Tabla | Backend dueño | Aplicación que la consume por SQL |
|---|---|---|
| TbExpedientes | Expedientes | Registros (`src/forms/Form_FormRegistroEntradaAlta.form.txt:2332` y cuatro más) |
| TbProyectos | Gestión de Riesgos | Registros (`src/forms/Form_FormRegistroEntradaAlta.form.txt:2435` y tres más) |
| TbUsuariosAplicaciones | Lanzadera | Registros (`src/forms/Form_FormRegistroEntradaAlta.form.txt:2265` y cinco más) |

El grafo de dependencias ya marca `TbUsuariosAplicaciones` como tabla compartida por los seis backends del ecosistema. Las otras dos no aparecen en él: el grafo se construyó con vínculos, y estos tres accesos son por SQL desde el front, sin vínculo. El hallazgo amplía el grafo: `TbExpedientes` y `TbProyectos` tienen un consumidor más que el que el grafo registra.

### Tablas que el SQL consulta y no existen

- `TbIDEquipos`: la nombran en `FROM` o en un `JOIN` 5 SQL de diseño de Brass (`Form_FormPlanificacionEquipos.form.txt:935`, `:1143`, `:1248` y `Form_FormPlanificacionesBusqueda.form.txt:1287`, `:1420`) y 2 más (`Form_FormEquipoGestion.form.txt:960` y `Form_FormMaterialReparacionEdicion.form.txt:415`). La ficha de Brass ya la marca como tabla inexistente.
- `TbPlanificacionAñadida`: `Form_FormPlanificacionCierre.form.txt:978`.
- `TbE2ECola`: el front de Expedientes no la vincula, de modo que el control `ListaCola` no puede resolver su origen.

Estas consultas fallan en runtime. Un SQL que nombra una tabla inexistente no es un problema de PostgreSQL: es un problema ya presente en el legacy, y conviene resolverlo antes del corte.

### Divergencias entre el export y el binario

El contraste con dysflow se hizo con `verify_code`, que compara el código en disco contra el binario en disco, en solo lectura. El resultado por aplicación:

| Aplicación | Formularios con origen de datos contrastados | Iguales al binario | Con divergencia funcional | Sin medir |
|---|---|---|---|---|
| Brass | 41 | 37 | 4 | 0 |
| Registros | 7 | 7 | 0 | 0 |
| HPS | 5 | 5 | 0 | 0 |
| Gestión de Riesgos | 1 | 1 | 0 | 0 |
| No Conformidades | 2 | 2 | 0 | 0 |
| Condor | 2 | 2 | 0 | 0 |
| Expedientes | 1 | 0 | 1 | 0 |

Las cuatro divergencias de Brass son `Form_FormEventoEdicion`, `Form_FormEventoGestion`, `Form_FormMaterialesGestion` y `Form_FormPlanificacionesBusqueda`. En los cuatro casos el clasificador de dysflow agotó el presupuesto de comparación por líneas y los marcó como `bothChanged` de forma conservadora, con exactamente una línea funcional más en el binario que en el export; la diferencia observada en el primero es de atributos de serialización de formulario (`RecordSelectors` frente a `AutoResize`), no de un `RowSource`. Es decir: la divergencia no está probada sobre el SQL, pero tampoco queda descartada. Los formularios afectados concentran 27 controles conectados: 7 en `Form_FormEventoEdicion`, 8 en `Form_FormEventoGestion`, 7 en `Form_FormMaterialesGestion` y 5 en `Form_FormPlanificacionesBusqueda`.

La divergencia de Expedientes sí es de un control concreto: `Form_FormE2ECola`, con la propiedad `ForeColor` de `lblTitulo` distinta y 57 líneas funcionales frente a 215 en el binario. El `RowSource` de diseño de ese formulario está vacío, de modo que el SQL en tiempo de ejecución no se ve afectado, pero el export del formulario no reproduce el binario.

En HPS, No Conformidades y Condor, la única diferencia entre export y binario es ruido de serialización (`Checksum`, `Attribute VB_Name`) o mayúsculas de identificador, sin efecto sobre el SQL. En Registros no hubo ninguna diferencia.

## Lo que no se pudo verificar

1. **El SHA de la base de HPS.** El worktree `00_HPS-worktrees/2026-005-rc1` tiene los metadatos de git rotos (`C:/00repos/codigo/00_HPS/.git/worktrees/2026-005-rc1` no existe), así que no se pudo confirmar la commit. La rama más cercana en el repositorio es `origin/fix/2026-005-rc1` (`44673692167ee0f7dd534358898b788a605a323b`), pero no está verificado que el árbol del worktree corresponda a esa commit.
2. **Los cuatro formularios de Brass con divergencia funcional.** `verify_code` los declara divergentes de forma conservadora; no se pudo aislar si el `RowSource` de alguno de ellos difiere del binario. Requiere volcar el formulario desde el binario y comparar control a control.
3. **El SQL efectivo de los filtros dinámicos.** En cinco consultas de Brass el `WHERE` se compone en variables aparte del `SELECT` (`FiltrarEventos`, `FiltrarActividades` y `FiltrarMateriales` de `Form_FormEventoGestion`, `Filtrar` de `Form_FormGastoGestion` y `FiltrarMateriales` de `Form_FormMaterialesGestion`); aquí está citado el `SELECT` con la parte del `WHERE` marcada como ausente. El `WHERE` completo no es texto fijo y cambia con el filtro de pantalla.
4. **Los identificadores interpolados.** El SQL en tiempo de ejecución concatena identificadores y cadenas directamente (`'& strIDActividad &'`). Aquí se muestran con `<id>` porque el valor no es fijo; el modo de interpolación no se ha auditado en este anexo.
5. **El valor de los filtros de pantalla.** El SQL de diseño de Brass lleva literales como `'SANTORCAZ'`, `'MADRID'`, `'TELEFÓNICA'` o `538` incrustados en la propia propiedad: son valores de ejemplo guardados con el formulario, no constantes de la aplicación. Con ellos, el control muestra un subconjunto hasta que el código lo sustituye.
6. **Los formularios sin control con origen de datos.** Solo se inventariaron los controles con `RowSourceType` de tabla o consulta, los `RecordSource` de formulario y las asignaciones en runtime. No se revisó si algún formulario consulta por `DLookup`, `Recordset` o `OpenRecordset` para rellenar un control, que es un patrón que no deja rastro en las propiedades del formulario.
7. **El número de consultas por tabla del catálogo.** Cuenta las SQL distintas que mencionan la tabla en su `FROM` o en sus `JOIN`, no las llamadas a procedimiento almacenado ni las consultas que la alcanzan por una vista o una consulta guardada del backend.
