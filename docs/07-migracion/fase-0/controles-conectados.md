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

72 SQL distintas sobre 31 tablas del backend `Gestion_Brass_Gestion_Datos`. La columna de usos indica formulario, control y `ruta:línea` del export.

| SQL literal | Usos | Tablas | Tipo de origen |
|---|---|---|---|
| `SELECT TbUbicacion.UBICACION FROM TbUbicacion ORDER BY TbUbicacion.UBICACION;` | Form_FormActividadEdicion / ComboUbicacion (`src/forms/Form_FormActividadEdicion.form.txt:544`)<br>Form_FormUbicacionGestion / ListaUbicacion (`src/forms/Form_FormUbicacionGestion.form.txt:328`) | TbUbicacion | Table/Query |
| `SELECT TbCodActividad.CODIGO, TbCodActividad.ACTIVIDADNOPROGRAMADA FROM TbCodActividad ORDER BY TbCodActividad.CODIGO;` | Form_FormActividadEdicion / ComboActividad (`src/forms/Form_FormActividadEdicion.form.txt:584`) | TbCodActividad | Table/Query |
| `SELECT TbEventos.IDEvento, TbEventos.DESCRIPCION FROM TbEventos WHERE ((Not (TbEventos.IDEvento)='CCM1301004')) ORDER BY TbEventos.IDEvento;` | Form_FormActividadEdicion / ComboIDEventoGenerado (`src/forms/Form_FormActividadEdicion.form.txt:625`) | TbEventos | Table/Query |
| `SELECT TbActividades.IDActividad, TbActividades.DESCRIPCION AS DESCRIPCIÓN, TbActividades.HorasLaborables AS HLab, TbActividades.IDEvento, TbActividades.FechaAlta AS Alta, TbFacturaActividadesInvolucradas.IDFactura, TbSellados.FechaSello, TbEventos.Franqueado, TbActividades.ALIASTECNICO FROM ((TbActividades LEFT JOIN TbFacturaActividadesInvolucradas ON TbActividades.IDActividad = TbFacturaActividadesInvolucradas.IDActividad) LEFT JOIN TbSellados ON TbActividades.IDActividad = TbSellados.IDActividad) LEFT JOIN TbEventos ON TbActividades.IDEvento = TbEventos.IDEvento WHERE (((TbActividades.IDActividad)='-1'));` | Form_FormActividadGestion / ListaFiltrados (`src/forms/Form_FormActividadGestion.form.txt:394`) | TbActividades, TbFacturaActividadesInvolucradas, TbSellados, TbEventos | Table/Query |
| `SELECT TbActividades.IDActividad, TbActividades.DESCRIPCION FROM TbActividades ORDER BY TbActividades.IDActividad;` | Form_FormActividadGestion / ComboIDActividadBusqueda (`src/forms/Form_FormActividadGestion.form.txt:1640`) | TbActividades | Table/Query |
| `SELECT TbCodActividad.CODIGO AS Código, TbCodActividad.ACTIVIDADNOPROGRAMADA AS Tipo FROM TbCodActividad;` | Form_FormActividadTipoGestion / ListaTiposDeActividad (`src/forms/Form_FormActividadTipoGestion.form.txt:560`) | TbCodActividad | Table/Query |
| `SELECT TbEventos.IDEvento, TbEventos.BUI, TbEventos.SUBSISTEMA, TbEquipos.Equipo FROM TbEventos LEFT JOIN TbEquipos ON TbEventos.IDEquipo = TbEquipos.IDEquipo ORDER BY TbEventos.FechaRegistroAlta DESC;` | Form_FormAnexosNoAlcanzables / ComboIDEventoBusqueda (`src/forms/Form_FormAnexosNoAlcanzables.form.txt:474`)<br>Form_FormEventoGestion / ComboIDEventoBusqueda (`src/forms/Form_FormEventoGestion.form.txt:1919`) | TbEventos, TbEquipos | Table/Query |
| `SELECT TbBUI.BUI FROM TbBUI ORDER BY TbBUI.BUI;` | Form_FormAnexosNoAlcanzables / ComboBUIBusqueda (`src/forms/Form_FormAnexosNoAlcanzables.form.txt:512`)<br>Form_FormBUIGestion / ListaBUI (`src/forms/Form_FormBUIGestion.form.txt:329`)<br>Form_FormCalibracionEquipoAlta / ComboBui (`src/forms/Form_FormCalibracionEquipoAlta.form.txt:635`)<br>Form_FormCalibracionEquipoEdicion / ComboBui (`src/forms/Form_FormCalibracionEquipoEdicion.form.txt:635`)<br>Form_FormEquipoGestion / ComboBUIBusqueda (`src/forms/Form_FormEquipoGestion.form.txt:1380`)<br>Form_FormEventoGestion / ComboBUIBusqueda (`src/forms/Form_FormEventoGestion.form.txt:1997`)<br>Form_FormMaterialesGestion / ComboBUIBusqueda (`src/forms/Form_FormMaterialesGestion.form.txt:3190`)<br>Form_FormPlanificacionesBusqueda / ComboBUIBusqueda (`src/forms/Form_FormPlanificacionesBusqueda.form.txt:1323`) | TbBUI | Table/Query |
| `SELECT DISTINCT TbEquiposCalibrables.IDEquipoCalibrable, TbEquiposCalibrables.BUI, TbEquiposCalibrables.Equipo, TbEquiposCalibrables.Modelo, TbEquiposCalibrables.NS FROM TbEquiposCalibrables LEFT JOIN TbEquiposCalibrablesFechas ON TbEquiposCalibrables.IDEquipoCalibrable = TbEquiposCalibrablesFechas.IDEquipoCalibrable WHERE (((TbEquiposCalibrables.BUI) Like '*' Or (TbEquiposCalibrables.BUI) Is Null) AND ((TbEquiposCalibrables.Equipo) Like '*' Or (TbEquiposCalibrables.Equipo) Is Null));` | Form_FormCalibracionEquipoGestion / ListaFiltrados (`src/forms/Form_FormCalibracionEquipoGestion.form.txt:945`) | TbEquiposCalibrables, TbEquiposCalibrablesFechas | Table/Query |
| `SELECT DISTINCT TbBUI.BUI FROM TbBUI ORDER BY TbBUI.BUI;` | Form_FormCalibracionEquipoGestion / ComboBUI (`src/forms/Form_FormCalibracionEquipoGestion.form.txt:1178`)<br>Form_FormPlanificacionEquipos / ComboBUI (`src/forms/Form_FormPlanificacionEquipos.form.txt:1076`)<br>Form_FormSubSistemaGestion / ComboBUI (`src/forms/Form_FormSubSistemaGestion.form.txt:1223`) | TbBUI | Table/Query |
| `SELECT DISTINCT TbEquiposCalibrables.Equipo, TbEquiposCalibrables.BUI FROM TbEquiposCalibrables ORDER BY TbEquiposCalibrables.Equipo;` | Form_FormCalibracionEquipoGestion / ComboEquipos (`src/forms/Form_FormCalibracionEquipoGestion.form.txt:1245`) | TbEquiposCalibrables | Table/Query |
| `SELECT TbEquiposCalibrablesFechas.IDCalibracion, TbEquiposCalibrablesFechas.IDEquipoCalibrable, TbEquiposCalibrablesFechas.FechaCalibrado AS [Fecha Inicial], TbEquiposCalibrablesFechas.FechaFinCalibrado AS [Fecha Final], TbEquiposCalibrablesFechas.URLDocumentosAnexados FROM TbEquiposCalibrablesFechas WHERE (((TbEquiposCalibrablesFechas.IDEquipoCalibrable)=1));` | Form_FormCalibracionGestion / ListaFiltrados (`src/forms/Form_FormCalibracionGestion.form.txt:977`) | TbEquiposCalibrablesFechas | Table/Query |
| `SELECT DISTINCT TbSubsistemaBui.SUBSISTEMA FROM TbSubsistemaBui WHERE (((TbSubsistemaBui.BUI) = 'TELEFÓNICA')) ORDER BY TbSubsistemaBui.SUBSISTEMA;` | Form_FormEquipoAlta / ComboSubsistema (`src/forms/Form_FormEquipoAlta.form.txt:957`) | TbSubsistemaBui | Table/Query |
| `SELECT TbNodos.NODO FROM TbNodos ORDER BY TbNodos.NODO;` | Form_FormEquipoAlta / ComboNodo (`src/forms/Form_FormEquipoAlta.form.txt:1119`)<br>Form_FormEquipoEdicion / ComboNodo (`src/forms/Form_FormEquipoEdicion.form.txt:536`)<br>Form_FormNodoBUIGestion / ComboNodos (`src/forms/Form_FormNodoBUIGestion.form.txt:606`)<br>Form_FormNodoGestion / ListaNodos (`src/forms/Form_FormNodoGestion.form.txt:325`) | TbNodos | Table/Query |
| `SELECT DISTINCT TbNodoBUI.BUI FROM TbNodoBUI WHERE (((TbNodoBUI.NODO)='MADRID')) ORDER BY TbNodoBUI.BUI;` | Form_FormEquipoAlta / ComboBUI (`src/forms/Form_FormEquipoAlta.form.txt:1157`)<br>Form_FormEquipoEdicion / ComboBUI (`src/forms/Form_FormEquipoEdicion.form.txt:574`) | TbNodoBUI | Table/Query |
| `SELECT DISTINCT TbSubsistemaBui.SUBSISTEMA FROM TbSubsistemaBui WHERE (((TbSubsistemaBui.BUI)='SANTORCAZ')) ORDER BY TbSubsistemaBui.SUBSISTEMA;` | Form_FormEquipoEdicion / ComboSubsistema (`src/forms/Form_FormEquipoEdicion.form.txt:374`) | TbSubsistemaBui | Table/Query |
| `SELECT DISTINCT TbEquipos.IDEquipo, TbEquipos.BUI, TbEquipos.SUBSISTEMA, TbEquipos.Equipo AS Equipo, TbPlanificacionEquipos.EquipoID, TbEquipos.NODO FROM TbEquipos LEFT JOIN TbPlanificacionEquipos ON TbIDEquipos.EquipoID = TbPlanificacionEquipos.EquipoID ORDER BY TbEquipos.BUI, TbEquipos.SUBSISTEMA, TbEquipos.Equipo;` | Form_FormEquipoGestion / ListaFiltrados (`src/forms/Form_FormEquipoGestion.form.txt:957`) | TbEquipos, TbPlanificacionEquipos | Table/Query |
| `SELECT TbSubsistemaBui.SUBSISTEMA FROM TbSubsistemaBui WHERE (((TbSubsistemaBui.BUI)='-1')) ORDER BY TbSubsistemaBui.SUBSISTEMA;` | Form_FormEquipoGestion / ComboSubsistemaBusqueda (`src/forms/Form_FormEquipoGestion.form.txt:1342`) | TbSubsistemaBui | Table/Query |
| `SELECT TbEquipos.Equipo FROM TbEquipos WHERE (((TbEquipos.SUBSISTEMA)='-1')) ORDER BY TbEquipos.IDEquipo;` | Form_FormEquipoGestion / ComboEquipoBusqueda (`src/forms/Form_FormEquipoGestion.form.txt:1500`) | TbEquipos | Table/Query |
| `SELECT DISTINCT TbNodoBUI.BUI, TbNodoBUI.NODO FROM TbNodoBUI INNER JOIN TbSubsistemaBui ON TbNodoBUI.BUI = TbSubsistemaBui.BUI ORDER BY TbNodoBUI.BUI;` | Form_FormEventoAlta / ComboBUI (`src/forms/Form_FormEventoAlta.form.txt:770`)<br>Form_FormEventoEdicion / ComboBUI (`src/forms/Form_FormEventoEdicion.form.txt:778`) | TbNodoBUI, TbSubsistemaBui | Table/Query |
| `SELECT TbTipoEvento.TIPOEVENTO FROM TbTipoEvento ORDER BY TbTipoEvento.TIPOEVENTO;` | Form_FormEventoAlta / ComboTipoEvento (`src/forms/Form_FormEventoAlta.form.txt:839`)<br>Form_FormEventoEdicion / ComboTipoEvento (`src/forms/Form_FormEventoEdicion.form.txt:847`)<br>Form_FormEventoTipoGestion / ListaTipoEventos (`src/forms/Form_FormEventoTipoGestion.form.txt:326`)<br>Form_FormEventosAltaMasivaPorNS / TipoEvento (`src/forms/Form_FormEventosAltaMasivaPorNS.form.txt:597`) | TbTipoEvento | Table/Query |
| `SELECT TbCriticidad.CRITICIDAD, TbCriticidad.DESCRIPCION FROM TbCriticidad;` | Form_FormEventoAlta / ComboCriticidad (`src/forms/Form_FormEventoAlta.form.txt:880`)<br>Form_FormEventoEdicion / ComboCriticidad (`src/forms/Form_FormEventoEdicion.form.txt:888`)<br>Form_FormEventosAltaMasivaPorNS / Criticidad (`src/forms/Form_FormEventosAltaMasivaPorNS.form.txt:638`) | TbCriticidad | Table/Query |
| `SELECT TbOriginador.ORIGINADOR FROM TbOriginador ORDER BY TbOriginador.ORIGINADOR;` | Form_FormEventoAlta / ComboOriginador (`src/forms/Form_FormEventoAlta.form.txt:917`)<br>Form_FormEventoEdicion / ComboOriginador (`src/forms/Form_FormEventoEdicion.form.txt:925`)<br>Form_FormEventosAltaMasivaPorNS / Originador (`src/forms/Form_FormEventosAltaMasivaPorNS.form.txt:675`)<br>Form_FormOriginadorGestion / ListaOriginadores (`src/forms/Form_FormOriginadorGestion.form.txt:328`) | TbOriginador | Table/Query |
| `SELECT DISTINCT TbEventos.CONTACTO FROM TbEventos WHERE ((Not (TbEventos.CONTACTO) Is Null) AND ((TbEventos.BUI)='-1')) ORDER BY TbEventos.CONTACTO;` | Form_FormEventoAlta / Contacto (`src/forms/Form_FormEventoAlta.form.txt:1871`) | TbEventos | Table/Query |
| `SELECT TbSubsistemaBui.SUBSISTEMA FROM TbSubsistemaBui WHERE (((TbSubsistemaBui.BUI)='SANTORCAZ')) ORDER BY TbSubsistemaBui.SUBSISTEMA;` | Form_FormEventoEdicion / ComboSubsistema (`src/forms/Form_FormEventoEdicion.form.txt:676`) | TbSubsistemaBui | Table/Query |
| `SELECT TbEquipos.IDEquipo, TbEquipos.Equipo, TbEquipos.FechaObsoleto FROM TbEquipos WHERE BUI='SANTORCAZ' AND SUBSISTEMA='ANTENA';` | Form_FormEventoEdicion / ComboIDEquipo (`src/forms/Form_FormEventoEdicion.form.txt:996`) | TbEquipos | Table/Query |
| `SELECT DISTINCT TbEventos.CONTACTO FROM TbEventos WHERE ((Not (TbEventos.CONTACTO) Is Null)) ORDER BY TbEventos.CONTACTO;` | Form_FormEventoEdicion / Contacto (`src/forms/Form_FormEventoEdicion.form.txt:2440`) | TbEventos | Table/Query |
| `SELECT IIf(IsNull([IDFactura]),'','*') AS F, TbEventos.IDEvento, TbEventos.BUI, TbEventos.SUBSISTEMA, TbEquipos.Equipo, TbEventos.FECHAALTAEVENTO AS FALTA, TbEventos.DESCRIPCION AS DESCRIPCIÓN, IIf([Franqueado],'Sí','No') AS Franqueo, TbEventos.IDEquipo FROM (TbEventos LEFT JOIN TbEquipos ON TbEventos.IDEquipo = TbEquipos.IDEquipo) LEFT JOIN TbFacturaEventosInvolucrados ON TbEventos.IDEvento = TbFacturaEventosInvolucrados.IDEvento WHERE (((TbEventos.IDEvento) Like '*' Or (TbEventos.IDEvento) Is Null)  AND ((TbEventos.BUI) Like '*' Or (TbEventos.BUI) Is Null)  AND ((TbEventos.SUBSISTEMA) Like '*' Or (TbEventos.SUBSISTEMA) Is Null)  AND ((TbEquipos.Equipo) Like '*' Or (TbEquipos.Equipo) Is Null)  AND ((TbEventos.TIPOEVENTO) Like '*' Or (TbEventos.TIPOEVENTO) Is Null)  AND ((TbEventos.Franqueado)=False) AND ((TbEventos.FECHAALTAEVENTO) Between #01/01/1900# And #01/01/2900#)) ORDER BY TbEventos.FECHAALTAEVENTO DESC;` | Form_FormEventoGestion / ListaEventos (`src/forms/Form_FormEventoGestion.form.txt:429`) | TbEventos, TbEquipos, TbFacturaEventosInvolucrados | Table/Query |
| `SELECT DISTINCT IIf(IsNull([TbMaterial].[IDMaterial]),'','*') AS M, TbActividades.IDActividad AS ID, TbActividades.ALIASTECNICO AS TÉCNICO, TbActividades.FechaAlta AS FALTA, TbActividades.HorasLaborables AS HLab, TbActividades.DESCRIPCION AS DESCRIPCIÓN, TbFacturaActividadesInvolucradas.IDFactura FROM (TbActividades LEFT JOIN TbMaterial ON TbActividades.IDActividad = TbMaterial.IDActividad) LEFT JOIN TbFacturaActividadesInvolucradas ON TbActividades.IDActividad = TbFacturaActividadesInvolucradas.IDActividad WHERE (((TbActividades.IDEvento)='BER2306002')) ORDER BY TbActividades.FechaAlta DESC;` | Form_FormEventoGestion / ListaActividades (`src/forms/Form_FormEventoGestion.form.txt:638`) | TbActividades, TbMaterial, TbFacturaActividadesInvolucradas | Table/Query |
| `SELECT TbMaterial.IDMaterial, TbMaterial.Material, TbMaterial.PN, TbMaterial.NS, TbMaterial.EsReparacion, TbFacturaMaterialesInvolucrados.IDFactura FROM TbMaterial LEFT JOIN TbFacturaMaterialesInvolucrados ON TbMaterial.IDMaterial = TbFacturaMaterialesInvolucrados.IDMaterial WHERE (((TbMaterial.IDActividad)='-1'));` | Form_FormEventoGestion / ListaMateriales (`src/forms/Form_FormEventoGestion.form.txt:846`) | TbMaterial, TbFacturaMaterialesInvolucrados | Table/Query |
| `SELECT TbSubsistemaBui.SUBSISTEMA FROM TbSubsistemaBui;` | Form_FormEventoGestion / ComboSubsistemaBusqueda (`src/forms/Form_FormEventoGestion.form.txt:1960`) | TbSubsistemaBui | Table/Query |
| `SELECT TbEquipos.Equipo FROM TbEquipos ORDER BY TbEquipos.Equipo;` | Form_FormEventoGestion / ComboEquipoBusqueda (`src/forms/Form_FormEventoGestion.form.txt:2115`) | TbEquipos | Table/Query |
| `SELECT TbTipoEvento.TIPOEVENTO FROM TbTipoEvento;` | Form_FormEventoGestion / ComboTipoEventoBusqueda (`src/forms/Form_FormEventoGestion.form.txt:3613`) | TbTipoEvento | Table/Query |
| `SELECT TbSubsistemaBui.SUBSISTEMA FROM TbSubsistemaBui WHERE (((TbSubsistemaBui.BUI)='CECOMMAD')) ORDER BY TbSubsistemaBui.SUBSISTEMA;` | Form_FormEventosAltaMasivaPorNS / Subsistema (`src/forms/Form_FormEventosAltaMasivaPorNS.form.txt:389`) | TbSubsistemaBui | Table/Query |
| `SELECT DISTINCT TbNodoBUI.BUI FROM TbNodoBUI ORDER BY TbNodoBUI.BUI;` | Form_FormEventosAltaMasivaPorNS / BUI (`src/forms/Form_FormEventosAltaMasivaPorNS.form.txt:529`) | TbNodoBUI | Table/Query |
| `SELECT DISTINCT TbEventos.CONTACTO FROM TbEventos WHERE ((Not (TbEventos.CONTACTO) Is Null) AND ((TbEventos.BUI)='MADRID')) ORDER BY TbEventos.CONTACTO;` | Form_FormEventosAltaMasivaPorNS / Contacto (`src/forms/Form_FormEventosAltaMasivaPorNS.form.txt:1545`) | TbEventos | Table/Query |
| `SELECT DISTINCT TbGastos.TipoGasto FROM TbGastos;` | Form_FormGastoAlta / ComboTipoGasto (`src/forms/Form_FormGastoAlta.form.txt:478`)<br>Form_FormGastoEdicion / ComboTipoGasto (`src/forms/Form_FormGastoEdicion.form.txt:715`) | TbGastos | Table/Query |
| `SELECT TbGastos.IDGasto AS ID, TbGastos.TipoGasto AS Gasto, IIf(IsNumeric([TbGastos].[ImporteUnitario]) And IsNumeric([TbGastos].[NumeroUnidades]),Format(CDbl([TbGastos].[ImporteUnitario])*CDbl([TbGastos].[NumeroUnidades]),'#,##0.00 €'),Null) AS [Importe Total], TbGastos.FechaImputacionGasto AS Fecha, TbTecnicos.NOMBRE AS Técnico, TbFacturaGastosInvolucrados.IDFactura, TbAnexos.NombreArchivo FROM ((TbGastos LEFT JOIN TbFacturaGastosInvolucrados ON TbGastos.IDGasto = TbFacturaGastosInvolucrados.IDGasto) LEFT JOIN TbTecnicos ON TbGastos.ALIAS = TbTecnicos.ALIAS) LEFT JOIN TbAnexos ON TbGastos.IDGasto = TbAnexos.IDGasto WHERE (((TbGastos.TipoGasto) Like '*' Or (TbGastos.TipoGasto) Is Null) AND ((TbGastos.FechaImputacionGasto) Between #1/1/1900# And #1/1/2300#));` | Form_FormGastoGestion / ListaFiltrados (`src/forms/Form_FormGastoGestion.form.txt:971`) | TbGastos, TbFacturaGastosInvolucrados, TbTecnicos, TbAnexos | Table/Query |
| `SELECT TbTecnicos.ALIAS, TbTecnicos.NOMBRE FROM TbTecnicos;` | Form_FormGastoGestion / ComboTecnicoBusqueda (`src/forms/Form_FormGastoGestion.form.txt:1208`) | TbTecnicos | Table/Query |
| `SELECT DISTINCT TbGastos.TipoGasto FROM TbGastos WHERE ((Not (TbGastos.TipoGasto) Is Null)) ORDER BY TbGastos.TipoGasto;` | Form_FormGastoGestion / ComboGastoBusqueda (`src/forms/Form_FormGastoGestion.form.txt:1603`) | TbGastos | Table/Query |
| `SELECT DISTINCT TbMaterial.IDMaterial, TbMaterial.Material, TbMaterial.PN, TbMaterial.NS FROM (TbMaterial INNER JOIN TbActividades ON TbMaterial.IDActividad = TbActividades.IDActividad) INNER JOIN TbEventos ON TbActividades.IDEvento = TbEventos.IDEvento WHERE ((Not (TbMaterial.Material) Is Null) And ((TbEventos.EquipoID) =538)) ORDER BY TbMaterial.Material;` | Form_FormMaterialAltaReparacion / ComboMaterial (`src/forms/Form_FormMaterialAltaReparacion.form.txt:447`) | TbMaterial, TbActividades, TbEventos | Table/Query |
| `SELECT DISTINCT TbMaterial.ReparadoPor FROM TbMaterial WHERE ((Not (TbMaterial.ReparadoPor) Is Null)) ORDER BY TbMaterial.ReparadoPor;` | Form_FormMaterialAltaReparacion / ComboReparadoPor (`src/forms/Form_FormMaterialAltaReparacion.form.txt:602`)<br>Form_FormMaterialReparacionEdicion / ComboReparadoPor (`src/forms/Form_FormMaterialReparacionEdicion.form.txt:712`)<br>Form_FormMaterialRepuestoAlta / ComboReparadoPor (`src/forms/Form_FormMaterialRepuestoAlta.form.txt:597`)<br>Form_FormMaterialRepuestoEdicion / ComboReparadoPor (`src/forms/Form_FormMaterialRepuestoEdicion.form.txt:595`) | TbMaterial | Table/Query |
| `SELECT DISTINCT TbTipoAccion.ESTADOREPARABLE FROM TbTipoAccion ORDER BY TbTipoAccion.ESTADOREPARABLE;` | Form_FormMaterialAltaReparacion / ComboTipoAccion (`src/forms/Form_FormMaterialAltaReparacion.form.txt:639`)<br>Form_FormMaterialRepuestoAlta / ComboTipoAccion (`src/forms/Form_FormMaterialRepuestoAlta.form.txt:634`)<br>Form_FormMaterialRepuestoEdicion / ComboTipoAccion (`src/forms/Form_FormMaterialRepuestoEdicion.form.txt:632`) | TbTipoAccion | Table/Query |
| `SELECT DISTINCT TbMaterialSeguimiento.Destino FROM TbMaterialSeguimiento WHERE ((Not (TbMaterialSeguimiento.Destino) Is Null)) ORDER BY TbMaterialSeguimiento.Destino;` | Form_FormMaterialIntervencionAlta / DESTINO (`src/forms/Form_FormMaterialIntervencionAlta.form.txt:838`)<br>Form_FormMaterialIntervencionEdicion / DESTINO (`src/forms/Form_FormMaterialIntervencionEdicion.form.txt:840`) | TbMaterialSeguimiento | Table/Query |
| `SELECT DISTINCT TbMaterial.Material FROM TbMaterial INNER JOIN TbEquipos ON TbMaterial.IDEquipo = TbIDEquipos.IDEquipo WHERE ((Not (TbMaterial.Material) Is Null) AND ((TbEquipos.IDEquipo)=-1)) ORDER BY TbMaterial.Material;` | Form_FormMaterialReparacionEdicion / ComboMaterial (`src/forms/Form_FormMaterialReparacionEdicion.form.txt:413`) | TbMaterial, TbEquipos | Table/Query |
| `SELECT DISTINCT TbTipoAccion.ESTADOREPARABLE FROM TbTipoAccion WHERE ((Not (TbTipoAccion.ESTADOREPARABLE) Is Null)) ORDER BY TbTipoAccion.ESTADOREPARABLE;` | Form_FormMaterialReparacionEdicion / ComboTipoAccion (`src/forms/Form_FormMaterialReparacionEdicion.form.txt:807`) | TbTipoAccion | Table/Query |
| `SELECT DISTINCT TbMaterial.Material, TbMaterial.PN, TbMaterial.NS FROM (TbMaterial INNER JOIN TbActividades ON TbMaterial.IDActividad = TbActividades.IDActividad) INNER JOIN TbEventos ON TbActividades.IDEvento = TbEventos.IDEvento WHERE ((Not (TbMaterial.Material) Is Null) AND ((TbEventos.IDEquipo)=538)) ORDER BY TbMaterial.Material;` | Form_FormMaterialRepuestoAlta / ComboMaterial (`src/forms/Form_FormMaterialRepuestoAlta.form.txt:447`) | TbMaterial, TbActividades, TbEventos | Table/Query |
| `SELECT DISTINCT TbMaterial.Material, TbMaterial.PN, TbMaterial.NS FROM (TbMaterial INNER JOIN TbActividades ON TbMaterial.IDActividad = TbActividades.IDActividad) INNER JOIN TbEventos ON TbActividades.IDEvento = TbEventos.IDEvento WHERE ((Not (TbMaterial.Material) Is Null) AND ((TbEventos.EquipoID)=538)) ORDER BY TbMaterial.Material;` | Form_FormMaterialRepuestoEdicion / ComboMaterial (`src/forms/Form_FormMaterialRepuestoEdicion.form.txt:445`) | TbMaterial, TbActividades, TbEventos | Table/Query |
| `SELECT TbMaterialSeguimiento.IDSeguimiento, TbMaterialSeguimiento.FechaInicioIntervencion AS [F Inicio], TbMaterialSeguimiento.Destino, TbMaterialSeguimiento.FechaFinIntervencion AS [F Fin], TbMaterialSeguimiento.UltimaIntervencion AS Última FROM TbMaterialSeguimiento WHERE (((TbMaterialSeguimiento.IDMaterial)='CGA_130214142101_01')) ORDER BY TbMaterialSeguimiento.FechaInicioIntervencion DESC;` | Form_FormMaterialSeguimiento / ListaIntervenciones (`src/forms/Form_FormMaterialSeguimiento.form.txt:958`) | TbMaterialSeguimiento | Table/Query |
| `SELECT DISTINCT TbMaterial.Material FROM TbMaterial ORDER BY TbMaterial.Material;` | Form_FormMaterialesGestion / ComboMaterialBusqueda (`src/forms/Form_FormMaterialesGestion.form.txt:1324`) | TbMaterial | Table/Query |
| `SELECT DISTINCT TbMaterial.PN, TbMaterial.NS FROM TbMaterial WHERE ((Not (TbMaterial.PN) Is Null)) ORDER BY TbMaterial.PN;` | Form_FormMaterialesGestion / ComboPNBusqueda (`src/forms/Form_FormMaterialesGestion.form.txt:1498`) | TbMaterial | Table/Query |
| `SELECT DISTINCT TbMaterial.NS FROM TbMaterial WHERE ((Not (TbMaterial.NS) Is Null)) ORDER BY TbMaterial.NS;` | Form_FormMaterialesGestion / ComboNSBusqueda (`src/forms/Form_FormMaterialesGestion.form.txt:1537`) | TbMaterial | Table/Query |
| `SELECT TbMaterial.IDMaterial FROM TbMaterial ORDER BY TbMaterial.IDMaterial DESC;` | Form_FormMaterialesGestion / ComboIDMaterialBusqueda (`src/forms/Form_FormMaterialesGestion.form.txt:2945`) | TbMaterial | Table/Query |
| `SELECT TbSubsistemaBui.SUBSISTEMA FROM TbSubsistemaBui WHERE (((TbSubsistemaBui.BUI)='BERMEJA')) ORDER BY TbSubsistemaBui.SUBSISTEMA;` | Form_FormMaterialesGestion / ComboSubsistemaBusqueda (`src/forms/Form_FormMaterialesGestion.form.txt:3151`) | TbSubsistemaBui | Table/Query |
| `SELECT TbEquipos.Equipo FROM TbEquipos WHERE (((TbEquipos.SUBSISTEMA)='-1') AND ((TbEquipos.BUI)='-1')) ORDER BY TbEquipos.Equipo;` | Form_FormMaterialesGestion / ComboEquipoBusqueda (`src/forms/Form_FormMaterialesGestion.form.txt:3282`) | TbEquipos | Table/Query |
| `SELECT TbNodoBUI.BUI, TbNodoBUI.NODO FROM TbNodoBUI WHERE (((TbNodoBUI.NODO)='-1')) ORDER BY TbNodoBUI.BUI;` | Form_FormNodoBUIGestion / ListaBuisEnNodo (`src/forms/Form_FormNodoBUIGestion.form.txt:480`) | TbNodoBUI | Table/Query |
| `SELECT TbBUI.BUI FROM TbBUI LEFT JOIN TbNodoBUI ON TbBUI.BUI = TbNodoBUI.BUI WHERE (((TbNodoBUI.NODO) Is Null));` | Form_FormNodoBUIGestion / ComboBuiParaNodo (`src/forms/Form_FormNodoBUIGestion.form.txt:736`) | TbBUI, TbNodoBUI | Table/Query |
| `SELECT TbEventos.IDEvento, TbEventos.TIPOEVENTO, TbEventos.FECHAALTAEVENTO FROM TbEventos RIGHT JOIN TbPlanificacion ON TbEventos.IDEvento = TbPlanificacion.IDEvento WHERE (((TbEventos.TIPOEVENTO)='MANT. PROGRAMADO') AND ((TbPlanificacion.IDEvento) Is Null) AND ((TbEventos.EquipoID)=94));` | Form_FormPlanificacionCambioIDEvento / ComboIDEventoNuevo (`src/forms/Form_FormPlanificacionCambioIDEvento.form.txt:497`) | TbEventos, TbPlanificacion | Table/Query |
| `SELECT TbEventos.IDEvento, TbEventos.TIPOEVENTO, TbEventos.FECHAALTAEVENTO FROM (TbEventos LEFT JOIN TbPlanificacion ON TbEventos.IDEvento = TbPlanificacion.IDEvento) LEFT JOIN TbIDEquipos ON TbEventos.EquipoID = TbIDEquipos.EquipoID WHERE (((TbEventos.TIPOEVENTO)='mms') AND ((TbEventos.BUI)='bui') AND ((TbEventos.SUBSISTEMA)='subs') AND ((TbIDEquipos.IDEquipo)='id') AND ((TbPlanificacion.IDEvento) Is Null));` | Form_FormPlanificacionCierre / ComboIDEvento (`src/forms/Form_FormPlanificacionCierre.form.txt:904`)<br>Form_FormPlanificacionCierre / ComboIDEventosUsados (`src/forms/Form_FormPlanificacionCierre.form.txt:1101`) | TbEventos, TbPlanificacion, TbIDEquipos | Table/Query |
| `SELECT TbPlanificacionAñadida.ID, TbPlanificacionAñadida.FechaMantoRealizado FROM TbPlanificacionAñadida LEFT JOIN TbPlanificacion ON TbPlanificacionAñadida.ID = TbPlanificacion.IDNoEvento WHERE (((TbPlanificacionAñadida.ID)=1) AND ((TbPlanificacionAñadida.EquipoID)=30));` | Form_FormPlanificacionCierre / ComboIDNoEvento (`src/forms/Form_FormPlanificacionCierre.form.txt:975`) | TbPlanificacionAñadida, TbPlanificacion | Table/Query |
| `SELECT TbEquipos.IDEquipo, TbEquipos.IDEquipo, TbEquipos.BUI, TbEquipos.SUBSISTEMA FROM TbEquipos INNER JOIN TbPlanificacionEquipos ON TbEquipos.IDEquipo = TbPlanificacionEquipos.EquipoID WHERE (((TbEquipos.BUI) Like '*' Or (TbIDEquipos.BUI) Is Null) AND ((TbEquipos.SUBSISTEMA) Like '*' Or (TbEquipos.SUBSISTEMA) Is Null) AND ((TbEquipos.Equipo) Like '*')) ORDER BY TbEquipos.IDEquipo;` | Form_FormPlanificacionEquipos / ListaFiltrados (`src/forms/Form_FormPlanificacionEquipos.form.txt:931`) | TbEquipos, TbPlanificacionEquipos | Table/Query |
| `SELECT TbIDEquipos.IDEquipo FROM TbIDEquipos ORDER BY TbIDEquipos.IDEquipo;` | Form_FormPlanificacionEquipos / ComboEquipos (`src/forms/Form_FormPlanificacionEquipos.form.txt:1143`) | TbIDEquipos | Table/Query |
| `SELECT DISTINCT TbSubsistemaBui.SUBSISTEMA FROM TbSubsistemaBui INNER JOIN TbIDEquipos ON (TbSubsistemaBui.SUBSISTEMA = TbIDEquipos.SUBSISTEMA) AND (TbSubsistemaBui.BUI = TbIDEquipos.BUI) WHERE (((TbIDEquipos.BUI)='-1')) ORDER BY TbSubsistemaBui.SUBSISTEMA;` | Form_FormPlanificacionEquipos / ComboSubSistema (`src/forms/Form_FormPlanificacionEquipos.form.txt:1245`) | TbSubsistemaBui, TbIDEquipos | Table/Query |
| `SELECT TbPlanificacion.IDPlanificacion AS ID, TbEquipos.Equipo, TbPlanificacion.FechaPrevistaCierre AS [Fecha Prevista], TbPlanificacion.FechaCierre AS FCierre, TbPlanificacion.MotivoCierre, TbPlanificacion.IDNuevaPlanificacion AS IDNueva, IIf(IsNull([FechaCierre]),DateDiff('d',Now(),[fechaprevistacierre]),0) AS [Días Restan], TbEquipos.BUI, TbEquipos.SUBSISTEMA, TbPlanificacion.EquipoID, TbPlanificacion.IDEvento FROM (TbPlanificacion INNER JOIN TbPlanificacionEquipos ON TbPlanificacion.EquipoID = TbPlanificacionEquipos.EquipoID) INNER JOIN TbEquipos ON TbPlanificacionEquipos.EquipoID = TbEquipos.IDEquipo WHERE (((TbEquipos.IDEquipo)='-1'));` | Form_FormPlanificacionesBusqueda / ListaFiltrados (`src/forms/Form_FormPlanificacionesBusqueda.form.txt:1004`) | TbPlanificacion, TbPlanificacionEquipos, TbEquipos | Table/Query |
| `SELECT DISTINCT TbIDEquipos.SUBSISTEMA FROM TbIDEquipos INNER JOIN TbPlanificacionEquipos ON TbIDEquipos.EquipoID = TbPlanificacionEquipos.EquipoID WHERE (((TbIDEquipos.BUI)='SANTORCAZ')) ORDER BY SUBSISTEMA;` | Form_FormPlanificacionesBusqueda / ComboSubsistemaBusqueda (`src/forms/Form_FormPlanificacionesBusqueda.form.txt:1285`) | TbIDEquipos, TbPlanificacionEquipos | Table/Query |
| `SELECT TbIDEquipos.IDEquipo, TbIDEquipos.BUI, TbIDEquipos.SUBSISTEMA FROM TbIDEquipos INNER JOIN TbPlanificacionEquipos ON TbIDEquipos.EquipoID = TbPlanificacionEquipos.EquipoID WHERE (((TbPlanificacionEquipos.FechabajaParaPlanificacion) Is Null) AND ((TbIDEquipos.BUI)='SANTORCAZ') AND ((TbIDEquipos.SUBSISTEMA)='TRASMISOR 10085')) ORDER BY TbIDEquipos.IDEquipo;` | Form_FormPlanificacionesBusqueda / ComboEquipoBusqueda (`src/forms/Form_FormPlanificacionesBusqueda.form.txt:1416`) | TbIDEquipos, TbPlanificacionEquipos | Table/Query |
| `SELECT DISTINCT TbPlanificacion.IDPlanificacion FROM TbPlanificacion ORDER BY TbPlanificacion.IDPlanificacion DESC;` | Form_FormPlanificacionesBusqueda / ComboIDMantenimiento (`src/forms/Form_FormPlanificacionesBusqueda.form.txt:2032`) | TbPlanificacion | Table/Query |
| `SELECT DISTINCT TbSubcontrataciones.Empresa FROM TbSubcontrataciones ORDER BY TbSubcontrataciones.Empresa;` | Form_FormSubContratacionAlta / ComboEmpresa (`src/forms/Form_FormSubContratacionAlta.form.txt:560`)<br>Form_FormSubContratacionEdicion / ComboEmpresa (`src/forms/Form_FormSubContratacionEdicion.form.txt:402`)<br>Form_FormSubContratacionGestion / ComboEmpresaBusqueda (`src/forms/Form_FormSubContratacionGestion.form.txt:636`) | TbSubcontrataciones | Table/Query |
| `SELECT TbSubcontrataciones.IDSubcontratacion AS ID, TbSubcontrataciones.Empresa, TbSubcontrataciones.FechaAlta, TbSubcontrataciones.FechaFin, TbSubcontrataciones.Importe, TbFacturaSubcontratacionesInvolucradas.IDFactura FROM TbSubcontrataciones LEFT JOIN TbFacturaSubcontratacionesInvolucradas ON TbSubcontrataciones.IDSubcontratacion = TbFacturaSubcontratacionesInvolucradas.IDSubcontratacion WHERE (((TbSubcontrataciones.IDSubcontratacion)=-1)) ORDER BY TbSubcontrataciones.FechaAlta DESC;` | Form_FormSubContratacionGestion / ListaFiltrados (`src/forms/Form_FormSubContratacionGestion.form.txt:380`) | TbSubcontrataciones, TbFacturaSubcontratacionesInvolucradas | Table/Query |
| `SELECT TbSubcontrataciones.IDSubcontratacion AS ID, TbSubcontrataciones.FechaAlta, TbSubcontrataciones.FechaFin, TbSubcontrataciones.Empresa FROM TbSubcontrataciones ORDER BY TbSubcontrataciones.IDSubcontratacion;` | Form_FormSubContratacionGestion / ComboIDSubContratacion (`src/forms/Form_FormSubContratacionGestion.form.txt:873`) | TbSubcontrataciones | Table/Query |
| `SELECT DISTINCT TbSubsistemaBui.SUBSISTEMA FROM TbSubsistemaBui WHERE ((Not (TbSubsistemaBui.BUI)='CHICLANA')) ORDER BY TbSubsistemaBui.SUBSISTEMA;` | Form_FormSubSistemaAlta / SubSistema (`src/forms/Form_FormSubSistemaAlta.form.txt:540`) | TbSubsistemaBui | Table/Query |
| `SELECT TbSubsistemaBui.SUBSISTEMA FROM TbSubsistemaBui WHERE (((TbSubsistemaBui.BUI)='-1'));` | Form_FormSubSistemaGestion / ListaSubSistema (`src/forms/Form_FormSubSistemaGestion.form.txt:950`) | TbSubsistemaBui | Table/Query |

Dos SQL de Brass citan tablas que no existen en el backend: `TbIDEquipos` (en `Form_FormPlanificacionEquipos.form.txt:935`, `:1143`, `:1248` y en `Form_FormPlanificacionesBusqueda.form.txt:1287`, `:1420`) y `TbPlanificacionAñadida` (`Form_FormPlanificacionCierre.form.txt:978`). La ficha de Brass ya las registra como referencias muertas; aquí se confirma que también entran por el origen de datos de los combos, no solo por el código.

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

| SQL en tiempo de ejecución | Asignaciones | Tablas |
|---|---|---|
| `SELECT TbActividades.* FROM TbActividades WHERE (((TbActividades.IDActividad)='<id>'));` | `src/forms/Form_FormActividadEdicion.cls:521` (SQL en `:430`), `src/forms/Form_FormActividadGestion.cls:1027` (SQL en `:30`) | TbActividades |
| `SELECT TbEquiposCalibrablesFechas.IDEquipoCalibrable FROM TbEquiposCalibrablesFechas WHERE (((TbEquiposCalibrablesFechas.IDEquipoCalibrable)=<id>));` | `src/forms/Form_FormCalibracionEquipoGestion.cls:62` (SQL en `:185`) | TbEquiposCalibrablesFechas |
| `SELECT TbEquiposCalibrables.Equipo, TbEquiposCalibrables.Modelo, TbEquiposCalibrables.NS,TbEquiposCalibrables.BUI, TbEquiposCalibrables.FechaBajaParaCalibracion FROM TbEquiposCalibrables WHERE (((TbEquiposCalibrables.IDEquipoCalibrable)=<id>));` | `src/forms/Form_FormCalibracionGestion.cls:304` (SQL en `:264`) | TbEquiposCalibrables |
| `SELECT DISTINCT TbSubsistemaBui.SUBSISTEMA FROM TbSubsistemaBui WHERE (((TbSubsistemaBui.BUI) Like '*')) ORDER BY TbSubsistemaBui.SUBSISTEMA;` | `src/forms/Form_FormEquipoAlta.cls:71` (SQL en `:60`), `src/forms/Form_FormEquipoEdicion.cls:151` (SQL en `:140`) | TbSubsistemaBui |
| `SELECT distinct TbNodoBUI.BUI FROM TbNodoBUI WHERE (((TbNodoBUI.NODO) Like '*')) ORDER BY TbNodoBUI.BUI;` | `src/forms/Form_FormEquipoAlta.cls:103` (SQL en `:92`), `src/forms/Form_FormEquipoEdicion.cls:180` (SQL en `:170`) | TbNodoBUI |
| `SELECT DISTINCT TbNodoBUI.BUI FROM TbNodoBUI WHERE (((TbNodoBUI.NODO)='<nodo>')) ORDER BY TbNodoBUI.BUI;` | `src/forms/Form_FormEquipoAlta.cls:144` y `:160` (SQL en `:140`) | TbNodoBUI |
| `SELECT TbEquipos.* FROM TbEquipos WHERE (((TbEquipos.IDEquipo)=<id>));` | `src/forms/Form_FormEquipoEdicion.cls:309` y `:318` (SQL en `:286`), `src/forms/Form_FormEquipoGestion.cls:746` (SQL en `:165`) | TbEquipos |
| `SELECT TbEquipos.Equipo FROM TbEquipos WHERE (((TbEquipos.SUBSISTEMA)='-1')) ORDER BY TbEquipos.Equipo;` | `src/forms/Form_FormEquipoGestion.cls:578` y `:584` (SQL en `:574`) | TbEquipos |
| `SELECT TbSubsistemaBui.SUBSISTEMA FROM TbSubsistemaBui WHERE (((TbSubsistemaBui.BUI)='<bui>')) ORDER BY TbSubsistemaBui.SUBSISTEMA;` | `src/forms/Form_FormEquipoGestion.cls:799` y `:820` (SQL en `:795`), `src/forms/Form_FormMaterialesGestion.cls:1353` y `:1374` (SQL en `:1349`) | TbSubsistemaBui |
| `SELECT TbEventos.* FROM TbEventos WHERE (((TbEventos.IDEvento)='<id>'));` | `src/forms/Form_FormEventoAlta.cls:516` y `:530` (SQL en `:431`) | TbEventos |
| `SELECT DISTINCT TbEventos.CONTACTO FROM TbEventos WHERE ((Not (TbEventos.CONTACTO) Is Null) AND ((TbEventos.BUI)='<bui>')) ORDER BY TbEventos.CONTACTO;` | `src/forms/Form_FormEventoAlta.cls:821` (SQL en `:816`), `src/forms/Form_FormEventosAltaMasivaPorNS.cls:620` (SQL en `:615`) | TbEventos |
| `SELECT TbEventos.*, TbEquipos.Equipo FROM TbEquipos RIGHT JOIN TbEventos ON TbEquipos.IDEquipo = TbEventos.IDEquipo WHERE (((TbEventos.IDEvento)='<id>'));` | `src/forms/Form_FormEventoEdicion.cls:875` y `:889` (SQL en `:667`) | TbEquipos, TbEventos |
| `SELECT distinct Equipo FROM TbEquipos ORDER BY TbEquipos.Equipo;` | `src/forms/Form_FormEventoGestion.cls:75` (SQL en `:65`) | TbEquipos |
| `SELECT distinct SUBSISTEMA FROM TbSubsistemaBui ORDER BY SUBSISTEMA;` | `src/forms/Form_FormEventoGestion.cls:122` (SQL en `:116`) | TbSubsistemaBui |
| `SELECT IIf(IsNull([IDFactura]),'','*') AS F, TbEventos.IDEvento, TbEventos.BUI, TbEventos.SUBSISTEMA, TbEquipos.Equipo, TbEventos.FECHAALTAEVENTO AS FALTA, TbEventos.DESCRIPCION AS DESCRIPCIÓN, IIf([Franqueado],'Sí','No') AS Franqueo, TbEventos.IDEquipo FROM (TbEventos LEFT JOIN TbEquipos ON TbEventos.IDEquipo = TbEquipos.IDEquipo) LEFT JOIN TbFacturaEventosInvolucrados ON TbEventos.IDEvento = TbFacturaEventosInvolucrados.IDEvento` más el `WHERE` que compone el filtro | `src/forms/Form_FormEventoGestion.cls:2095` (SQL en `:1241`, procedimiento `FiltrarEventos`) | TbEventos, TbEquipos, TbFacturaEventosInvolucrados |
| `SELECT DISTINCT IIf(IsNull([TbMaterial].[IDMaterial]),'','*') AS M, TbActividades.IDActividad AS ID, TbActividades.ALIASTECNICO AS TÉCNICO, TbActividades.FechaAlta AS FALTA, TbActividades.HorasLaborables AS HLab, TbActividades.DESCRIPCION AS DESCRIPCIÓN, TbFacturaActividadesInvolucradas.IDFactura FROM (TbActividades LEFT JOIN TbMaterial ON TbActividades.IDActividad = TbMaterial.IDActividad) LEFT JOIN TbFacturaActividadesInvolucradas ON TbActividades.IDActividad = TbFacturaActividadesInvolucradas.IDActividad` más el `WHERE` | `src/forms/Form_FormEventoGestion.cls:2103` (SQL en `:1389`, procedimiento `FiltrarActividades`) | TbActividades, TbMaterial, TbFacturaActividadesInvolucradas |
| `SELECT TbMaterial.IDMaterial, TbMaterial.Material, TbMaterial.PN, TbMaterial.NS, TbMaterial.EsReparacion, TbFacturaMaterialesInvolucrados.IDFactura FROM TbMaterial LEFT JOIN TbFacturaMaterialesInvolucrados ON TbMaterial.IDMaterial = TbFacturaMaterialesInvolucrados.IDMaterial` más el `WHERE` | `src/forms/Form_FormEventoGestion.cls:2111` (SQL en `:1427`, procedimiento `FiltrarMateriales`) | TbMaterial, TbFacturaMaterialesInvolucrados |
| `SELECT TbActividades.IDActividad AS ID, TbActividades.ALIASTECNICO AS TÉCNICO, TbActividades.FechaAlta AS FALTA, TbActividades.HorasLaborables AS HLab, TbActividades.DESCRIPCION AS DESCRIPCIÓN FROM TbActividades WHERE (((TbActividades.IDActividad)='-1')) ORDER BY TbActividades.FechaAlta DESC;` | `src/forms/Form_FormEventoGestion.cls:2449`, `:2461` y `:2475` (SQL en `:2444`) | TbActividades |
| `SELECT TbMaterial.IDMaterial, TbMaterial.Material FROM TbMaterial WHERE (((TbMaterial.IDMaterial)='-1'));` | `src/forms/Form_FormEventoGestion.cls:2534` y `:2593` (SQL en `:2458`) | TbMaterial |
| `SELECT TbGastos.IDGasto AS ID, TbGastos.TipoGasto as Gasto, IIf(IsNumeric([TbGastos].[ImporteUnitario]) And IsNumeric([TbGastos].[NumeroUnidades]),Format(CDbl([TbGastos].[ImporteUnitario])*CDbl([TbGastos].[NumeroUnidades]),'#,##0.00 €'),Null) AS [Importe Total], TbGastos.FechaImputacionGasto AS Fecha, TbTecnicos.NOMBRE AS Técnico, TbFacturaGastosInvolucrados.IDFactura , TbAnexos.NombreArchivo FROM ((TbGastos LEFT JOIN TbFacturaGastosInvolucrados ON TbGastos.IDGasto = TbFacturaGastosInvolucrados.IDGasto) LEFT JOIN TbTecnicos ON TbGastos.ALIAS = TbTecnicos.ALIAS) LEFT JOIN TbAnexos ON TbGastos.IDGasto = TbAnexos.IDGasto` más el `WHERE` | `src/forms/Form_FormGastoGestion.cls:584` (SQL en `:506`, procedimiento `Filtrar`) | TbGastos, TbFacturaGastosInvolucrados, TbTecnicos, TbAnexos |
| `SELECT DISTINCT TbMaterial.IDMaterial,TbMaterial.Material, TbMaterial.PN, TbMaterial.NS FROM (TbMaterial INNER JOIN TbActividades ON TbMaterial.IDActividad = TbActividades.IDActividad) INNER JOIN TbEventos ON TbActividades.IDEvento = TbEventos.IDEvento WHERE ((Not (TbMaterial.Material) Is Null) And ((TbEventos.IDEquipo) =<id>)) ORDER BY TbMaterial.Material;` | `src/forms/Form_FormMaterialAltaReparacion.cls:414` (SQL en `:409`) | TbMaterial, TbActividades, TbEventos |
| `SELECT TbMaterial.IDEquipo, TbMaterial.IDActividad,TbMaterial.EsReparacion,TbMaterial.FechaEntrega FROM TbMaterial WHERE (((TbMaterial.IDMaterial)='<id>'));` | `src/forms/Form_FormMaterialReparacionEdicion.cls:522` (SQL en `:296`) | TbMaterial |
| `SELECT DISTINCT TbMaterial.Material, TbMaterial.PN, TbMaterial.NS FROM (TbMaterial INNER JOIN TbActividades ON TbMaterial.IDActividad = TbActividades.IDActividad) INNER JOIN TbEventos ON TbActividades.IDEvento = TbEventos.IDEvento WHERE ((Not (TbMaterial.Material) Is Null) And ((TbEventos.IDEquipo) =<id>)) ORDER BY TbMaterial.Material;` | `src/forms/Form_FormMaterialRepuestoAlta.cls:221` (SQL en `:216`) | TbMaterial, TbActividades, TbEventos |
| `SELECT TbMaterial.IDEquipo, TbMaterial.IDActividad,TbMaterial.EsReparacion FROM TbMaterial WHERE (((TbMaterial.IDMaterial)='<id>'));` | `src/forms/Form_FormMaterialRepuestoEdicion.cls:472` (SQL en `:405`) | TbMaterial |
| `SELECT TbMaterial.Garantia, TbMaterial.Material, TbActividades.IDEvento, TbActividades.IDActividad, TbEquipos.NODO, TbEquipos.BUI, TbEquipos.SUBSISTEMA, TbEquipos.Equipo FROM ((TbMaterial INNER JOIN TbActividades ON TbMaterial.IDActividad = TbActividades.IDActividad) INNER JOIN TbEventos ON TbActividades.IDEvento = TbEventos.IDEvento) INNER JOIN TbEquipos ON TbEventos.IDEquipo = TbEquipos.IDEquipo WHERE (((TbMaterial.IDMaterial)='<id>'));` | `src/forms/Form_FormMaterialSeguimiento.cls:413` (SQL en `:332`) | TbMaterial, TbActividades, TbEventos, TbEquipos |
| `SELECT DISTINCT TbMaterial.IDMaterial, TbMaterial.Material, TbMaterial.FechaIncio, TbMaterial.FechaEntrega, TbEquipos_1.Equipo AS EquipoDeMat1, TbMaterial.EsReparacion, TbMaterial.TipoAccion, TbEquipos.Equipo AS EquipoDeEvento, TbEquipos_1.IDEquipo AS EquipoDeMat, TbEquipos.IDEquipo AS IDEquipoDeEvento, TbEventos.IDEvento, TbEventos.BUI AS BUIEVENTO, TbEventos.SUBSISTEMA AS SUBSISTEMAEVENTO, TbEquipos_1.BUI AS BUIEQUIPO, TbEquipos_1.SUBSISTEMA AS SUBSISTEMAEQUIPO,TbMaterial.NS, TbFacturaMaterialesInvolucrados.IDFactura FROM (((TbEventos RIGHT JOIN (TbMaterial LEFT JOIN TbActividades ON TbMaterial.IDActividad = TbActividades.IDActividad) ON TbEventos.IDEvento = TbActividades.IDEvento) LEFT JOIN TbEquipos ON TbEventos.IDEquipo = TbEquipos.IDEquipo) LEFT JOIN TbEquipos AS TbEquipos_1 ON TbMaterial.IDEquipo = TbEquipos_1.IDEquipo) LEFT JOIN TbFacturaMaterialesInvolucrados ON TbMaterial.IDMaterial = TbFacturaMaterialesInvolucrados.IDMaterial` más el `WHERE` | `src/forms/Form_FormMaterialesGestion.cls:1121` (SQL en `:1127`, procedimiento `FiltrarMateriales`) | TbMaterial, TbActividades, TbEventos, TbEquipos, TbFacturaMaterialesInvolucrados |
| `SELECT TbNodoBUI.BUI, TbNodoBUI.NODO FROM TbNodoBUI WHERE (((TbNodoBUI.NODO)='<nodo>')) ORDER BY TbNodoBUI.BUI;` | `src/forms/Form_FormNodoBUIGestion.cls:115` y `:149` (SQL en `:111` y `:145`) | TbNodoBUI |
| `SELECT TbPlanificacionAnexos.IDAnexoPlanificacion, TbPlanificacionAnexos.IDPlanificacion, TbPlanificacionAnexos.NombreAnexo AS Título, TbPlanificacionAnexos.FechaAnexo AS [Fecha Anexo] FROM TbPlanificacionAnexos WHERE (((TbPlanificacionAnexos.IDPlanificacion)=<id>));` | `src/forms/Form_FormPlanificacionAnexos.cls:139` (SQL en `:135`) | TbPlanificacionAnexos |
| `SELECT TbEventos.IDEvento, TbEventos.TIPOEVENTO, TbEventos.FECHAALTAEVENTO FROM TbEventos LEFT JOIN TbPlanificacion ON TbEventos.IDEvento = TbPlanificacion.IDEvento WHERE (((TbEventos.TIPOEVENTO)='MANT. PROGRAMADO') AND ((TbEventos.IDEquipo)=<id>) AND ((TbPlanificacion.IDEvento) Is Null));` | `src/forms/Form_FormPlanificacionCambioIDEvento.cls:157` (SQL en `:153`) | TbEventos, TbPlanificacion |
| `SELECT  TbPlanificacion.Anio, TbPlanificacion.Mes, TbPlanificacion.Semana, TbPlanificacion.FechaCierre, TbPlanificacion.FechaPrevistaCierre, TbPlanificacion.IDEvento, TbPlanificacion.IDNoEvento, TbPlanificacion.IDEquipo FROM TbPlanificacion INNER JOIN TbEquipos ON TbPlanificacion.IDEquipo = TbEquipos.IDEquipo WHERE (((TbPlanificacion.IDPlanificacion)=<id>));` | `src/forms/Form_FormPlanificacionCierre.cls:120`, `:125` y `:132` (SQL en `:57`) | TbPlanificacion, TbEquipos |
| `SELECT TbEquipos.Equipo FROM TbEquipos ORDER BY TbEquipos.Equipo;` | `src/forms/Form_FormPlanificacionEquipos.cls:158` y `:390` (SQL en `:381`) | TbEquipos |
| `SELECT DISTINCT TbSubsistemaBui.SUBSISTEMA FROM TbSubsistemaBui INNER JOIN TbEquipos ON (TbSubsistemaBui.SUBSISTEMA = TbEquipos.SUBSISTEMA) AND (TbSubsistemaBui.BUI = TbEquipos.BUI) WHERE (((TbEquipos.BUI)='<bui>')) ORDER BY TbSubsistemaBui.SUBSISTEMA;` | `src/forms/Form_FormPlanificacionEquipos.cls:417` (SQL en `:412`) | TbSubsistemaBui, TbEquipos |
| `SELECT DISTINCT TbEquipos.SUBSISTEMA FROM TbEquipos INNER JOIN TbPlanificacionEquipos ON TbEquipos.IDEquipo = TbPlanificacionEquipos.EquipoID WHERE (((TbEquipos.BUI)='<bui>')) ORDER BY SUBSISTEMA;` | `src/forms/Form_FormPlanificacionesBusqueda.cls:38` y `:549` (SQL en `:545`) | TbEquipos, TbPlanificacionEquipos |
| `SELECT TbEquipos.Equipo, TbEquipos.BUI, TbEquipos.SUBSISTEMA FROM TbEquipos INNER JOIN TbPlanificacionEquipos ON TbEquipos.IDEquipo = TbPlanificacionEquipos.EquipoID WHERE (((TbPlanificacionEquipos.FechabajaParaPlanificacion) Is Null)) ORDER BY TbEquipos.Equipo;` | `src/forms/Form_FormPlanificacionesBusqueda.cls:583` (SQL en `:572`) | TbEquipos, TbPlanificacionEquipos |
| `SELECT TbSubcontrataciones.* FROM TbSubcontrataciones WHERE (((TbSubcontrataciones.IDSubcontratacion)=<id>));` | `src/forms/Form_FormSubContratacionGestion.cls:61` (SQL en `:412`) | TbSubcontrataciones |
| `SELECT distinct TbSubsistemaBui.SUBSISTEMA FROM TbSubsistemaBui WHERE ((Not (TbSubsistemaBui.BUI)='<bui>')) ORDER BY TbSubsistemaBui.SUBSISTEMA;` | `src/forms/Form_FormSubSistemaAlta.cls:68` (SQL en `:64`) | TbSubsistemaBui |
| `SELECT TbSubsistemaBui.SUBSISTEMA, TbSubsistemaBui.BUI FROM TbSubsistemaBui WHERE (((TbSubsistemaBui.SUBSISTEMA)='<subsistema>') AND ((TbSubsistemaBui.BUI)='<bui>'));` | `src/forms/Form_FormSubSistemaGestion.cls:318` (SQL en `:267`) | TbSubsistemaBui |

Cuatro asignaciones de Brass se resuelven con SQL construido en un módulo común, no en el propio formulario. Se citan aquí para que el catálogo de SQL no las pierda:

| SQL en tiempo de ejecución | Asignaciones | Tablas |
|---|---|---|
| `SELECT TbSubsistemaBui.SUBSISTEMA FROM TbSubsistemaBui WHERE (((TbSubsistemaBui.BUI)='<bui>')) ORDER BY TbSubsistemaBui.SUBSISTEMA;` (`src/modules/Funciones Generales.bas:1214`, procedimiento `EstablecerComboSubSistema`) | `src/forms/Form_FormEventoAlta.cls:256`, `src/forms/Form_FormEventoEdicion.cls:376`, `src/forms/Form_FormEventosAltaMasivaPorNS.cls:37` | TbSubsistemaBui |
| `SELECT TbEquipos.IDEquipo,TbEquipos.Equipo,TbEquipos.FechaObsoleto FROM TbEquipos WHERE BUI='<bui>' AND SUBSISTEMA='<subsistema>';` (`src/modules/Funciones Generales.bas:1198`, procedimiento `EstablecerComboEquipo`) | `src/forms/Form_FormEventoAlta.cls:336`, `src/forms/Form_FormEventoEdicion.cls:451` y `:702`, `src/forms/Form_FormEventosAltaMasivaPorNS.cls:231` | TbEquipos |
| `SELECT DISTINCT TbMaterial.Material FROM ((TbMaterial INNER JOIN TbActividades ON TbMaterial.IDActividad = TbActividades.IDActividad) INNER JOIN TbEventos ON TbActividades.IDEvento = TbEventos.IDEvento) INNER JOIN TbEquipos ON TbEventos.IDEquipo = TbEquipos.IDEquipo WHERE (((TbEquipos.IDEquipo)=<id>)) ORDER BY TbMaterial.Material;` (`src/modules/Funciones Generales.bas:1231`, procedimiento `EstablecerComboMaterial`) | `src/forms/Form_FormEventosAltaMasivaPorNS.cls:142` | TbMaterial, TbActividades, TbEventos, TbEquipos |
| `SELECT distinct TbEventos.CONTACTO FROM TbEventos WHERE ((Not (TbEventos.CONTACTO) Is Null) AND ((TbEventos.BUI)='<bui>')) ORDER BY TbEventos.CONTACTO;` (`src/modules/Funciones Generales.bas:6675`, procedimiento `EstablecerComboContacto`) | `src/forms/Form_FormEventoEdicion.cls:392` | TbEventos |

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
