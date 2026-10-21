# Proyecto, edición y riesgo — matrices campo a campo

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Issue:** #1079 · Fase 3 · **Matriz:** [matriz-campos.md](../matriz-campos.md)

18 tablas. Cada fila está en estado **propuesta**: la topología PostgreSQL global y el ERD de origen siguen abiertos (D-GR-4 del [README](../README.md)).

### ProyectoRespCalidad

- Columnas: **2**. Filas (snapshot): **16**.
- Evidencia en `src`: **1** fichero(s) de producción — `Constantes.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| Proyecto | TEXT(255) | — |
| UsuarioCalidad | TEXT(255) | personal |

### TbAux

- Columnas: **1**. Filas (snapshot): **13**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDProyecto | LONG(4) | — |

### TbAuxProyectosRiesgos

- Columnas: **5**. Filas (snapshot): **2**.
- Evidencia en `src`: **1** fichero(s) de producción — `mIndicador.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDProyecto | LONG(4) | — |
| Riesgo | TEXT(255) | — |
| FechaDetectado | SHORT_DATE_TIME(8) | — |
| NombreProyectoCompleto | TEXT(255) | — |
| FechaMaterializado | SHORT_DATE_TIME(8) | — |

### TbBibliotecaRiesgos

- Columnas: **9**. Filas (snapshot): **65**.
- Evidencia en `src`: **2** fichero(s) de producción — `RiesgoBiblioteca.cls`, `Constructor.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDRiesgoTipo | LONG(4) | — |
| CODIGO | TEXT(255) | — |
| Tipo | TEXT(255) | — |
| Descripcion | MEMO | — |
| Usuario | TEXT(255) | personal |
| FechaAlta | SHORT_DATE_TIME(8) | — |
| FechaModificacion | SHORT_DATE_TIME(8) | — |
| Familia | TEXT(255) | — |
| Activo | TEXT(2) | — |

### TbOrigenesRiesgosDetalles

- Columnas: **2**. Filas (snapshot): **5**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| Origen | TEXT(255) | — |
| Descripcion | MEMO | — |

### TbProyectoEdicionesCorreoRevision

- Columnas: **5**. Filas (snapshot): **3**.
- Evidencia en `src`: **2** fichero(s) de producción — `EdicionCorreoRevision.cls`, `Constructor.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDEnvioCorreoTecnico | LONG(4) | personal |
| IDEdicion | LONG(4) | — |
| FechaCorreoRevision | SHORT_DATE_TIME(8) | personal |
| UsuarioCalidad | TEXT(255) | personal |
| IDCorreo | LONG(4) | personal |

### TbProyectos

- Columnas: **25**. Filas (snapshot): **81**.
- Evidencia en `src`: **12** fichero(s) de producción — `Edicion.cls`, `Proyecto.cls`, `TareasCalidad.cls`, `TareasTecnico.cls`, `Form_formIndicadorProyectos.cls`, `Form_formIndicadorProyectos.form.txt`…

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDProyecto | LONG(4) | — |
| IDExpediente | LONG(4) | — |
| Proyecto | TEXT(255) | — |
| Juridica | TEXT(255) | — |
| NombreProyecto | TEXT(255) | — |
| Cliente | TEXT(255) | — |
| FechaPrevistaCierre | SHORT_DATE_TIME(8) | — |
| FechaCierre | SHORT_DATE_TIME(8) | — |
| FechaRegistroInicial | SHORT_DATE_TIME(8) | — |
| Elaborado | TEXT(255) | — |
| Revisado | TEXT(255) | — |
| Aprobado | TEXT(255) | — |
| CodigoDocumento | TEXT(255) | — |
| ParaInformeAvisos | TEXT(2) | — |
| FechaFirmaContrato | TEXT(255) | — |
| RiesgosDeLaOferta | TEXT(2) | — |
| RiesgosDelSubContratista | TEXT(2) | — |
| NombreUsuarioCalidad | TEXT(255) | personal |
| EnUTE | TEXT(2) | — |
| FechaMaxProximaPublicacion | SHORT_DATE_TIME(8) | — |
| RequiereRiesgoDeBiblioteca | TEXT(2) | — |
| CorreoRAC | TEXT(255) | personal |
| Ordinal | LONG(4) | — |
| CadenaNombreAutorizados | MEMO | — |
| NombreParaNodo | TEXT(255) | — |

### TbProyectosEdiciones

- Columnas: **22**. Filas (snapshot): **325**.
- Evidencia en `src`: **9** fichero(s) de producción — `Edicion.cls`, `Proyecto.cls`, `TareasCalidad.cls`, `TareasTecnico.cls`, `Constructor.bas`, `IndicadorRiesgosRepositorio.bas`…

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDEdicion | LONG(4) | — |
| IDProyecto | LONG(4) | — |
| FechaEdicion | SHORT_DATE_TIME(8) | — |
| Edicion | INT(2) | — |
| Elaborado | TEXT(255) | — |
| Revisado | TEXT(255) | — |
| Aprobado | TEXT(255) | — |
| FechaPublicacion | SHORT_DATE_TIME(8) | — |
| EntregadoAClienteORAC | TEXT(2) | — |
| Comentarios | MEMO | — |
| PermitidoImprimirExcel | TEXT(2) | — |
| IDDocumentoAGEDO | LONG(4) | — |
| NombreArchivoInforme | TEXT(255) | — |
| FechaMaxProximaPublicacion | SHORT_DATE_TIME(8) | — |
| FechaPreparadaParaPublicar | SHORT_DATE_TIME(8) | — |
| UsuarioProponePublicar | TEXT(255) | personal |
| PropuestaRechazadaPorCalidadFecha | SHORT_DATE_TIME(8) | personal |
| PropuestaRechazadaPorCalidadMotivo | MEMO | personal |
| UsuarioCalidadRechazaPropuesta | TEXT(255) | personal |
| NotasCalidadParaPublicar | MEMO | personal |
| FechaUltimoCambio | SHORT_DATE_TIME(8) | — |
| UsuarioUltimoCambio | TEXT(255) | personal |

### TbProyectosResponsablesCalidad

- Columnas: **2**. Filas (snapshot): **39**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDProyecto | LONG(4) | — |
| UsuarioCalidad | TEXT(255) | personal |

### TbProyectosTipo

- Columnas: **3**. Filas (snapshot): **0**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDTipoProyecto | LONG(4) | — |
| TipoProyecto | TEXT(255) | — |
| Descripcion | MEMO | — |

### TbRiesgos

- Columnas: **47**. Filas (snapshot): **2924**.
- Evidencia en `src`: **11** fichero(s) de producción — `Edicion.cls`, `Riesgo.cls`, `RiesgoBiblioteca.cls`, `RiesgoMaterializacion.cls`, `TareasCalidad.cls`, `TareasTecnico.cls`…

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDRiesgo | LONG(4) | — |
| IDEdicion | LONG(4) | — |
| CodigoUnico | TEXT(255) | — |
| CodigoRiesgo | TEXT(5) | — |
| FechaDetectado | SHORT_DATE_TIME(8) | — |
| DetectadoPor | TEXT(255) | — |
| EntidadDetecta | TEXT(255) | — |
| Plazo | TEXT(255) | — |
| Calidad | TEXT(255) | personal |
| Coste | TEXT(255) | — |
| ImpactoGlobal | TEXT(15) | — |
| Vulnerabilidad | TEXT(15) | — |
| Valoracion | TEXT(15) | — |
| Mitigacion | TEXT(15) | — |
| Contingencia | TEXT(15) | — |
| RequierePlanContingencia | TEXT(2) | — |
| Descripcion | MEMO | — |
| CausaRaiz | MEMO | — |
| Estado | TEXT(255) | — |
| FechaEstado | TEXT(255) | — |
| FechaMaterializado | SHORT_DATE_TIME(8) | — |
| FechaRetirado | SHORT_DATE_TIME(8) | — |
| FechaCerrado | SHORT_DATE_TIME(8) | — |
| FechaMitigacionAceptar | SHORT_DATE_TIME(8) | — |
| JustificacionAceptacionRiesgo | MEMO | — |
| FechaJustificacionAceptacionRiesgo | SHORT_DATE_TIME(8) | — |
| FechaAprobacionAceptacionPorCalidad | SHORT_DATE_TIME(8) | personal |
| FechaRechazoAceptacionPorCalidad | SHORT_DATE_TIME(8) | personal |
| JustificacionRetiroRiesgo | MEMO | — |
| FechaJustificacionRetiroRiesgo | SHORT_DATE_TIME(8) | — |
| FechaAprobacionRetiroPorCalidad | SHORT_DATE_TIME(8) | personal |
| FechaRechazoRetiroPorCalidad | SHORT_DATE_TIME(8) | personal |
| RequiereRiesgoDeBiblioteca | TEXT(2) | — |
| CodRiesgoBiblioteca | TEXT(255) | — |
| RiesgoPendienteRetipificacion | TEXT(2) | — |
| FechaRiesgoParaRetipificar | SHORT_DATE_TIME(8) | — |
| FechaRiesgoRetipificado | SHORT_DATE_TIME(8) | — |
| DiasSinRespuestaCalidadAceptacion | TEXT(255) | personal |
| DiasSinRespuestaCalidadRetiro | TEXT(255) | personal |
| DiasSinRespuestaCalidadRetipificacion | TEXT(255) | personal |
| Origen | TEXT(255) | — |
| HayErrorEnRiesgo | TEXT(255) | — |
| NombreNodoDesc | TEXT(255) | — |
| NombreNodoEstado | TEXT(255) | — |
| NombreIcono | TEXT(255) | — |
| ColorIcono | TEXT(255) | — |
| Priorizacion | LONG(4) | — |

### TbRiesgosAIntegrar

- Columnas: **20**. Filas (snapshot): **175**.
- Evidencia en `src`: **3** fichero(s) de producción — `RiesgoExterno.cls`, `Constructor.bas`, `IndicadorRiesgosRepositorio.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDRiesgoExt | LONG(4) | — |
| CodRiesgo | TEXT(255) | — |
| IDRiesgo | LONG(4) | — |
| Origen | TEXT(255) | — |
| IDEdicion | LONG(4) | — |
| Descripcion | MEMO | — |
| CausaRaiz | TEXT(255) | — |
| FechaDetectado | SHORT_DATE_TIME(8) | — |
| FechaAltaRegistro | SHORT_DATE_TIME(8) | — |
| UsuarioRegistra | TEXT(255) | personal |
| MotivoNoIntegrado | MEMO | — |
| FechaMotivo | SHORT_DATE_TIME(8) | — |
| Trasladar | TEXT(2) | — |
| Suministrador | TEXT(255) | — |
| Pedido | TEXT(255) | — |
| ProveedorPedido | TEXT(255) | — |
| CausaRiesgoPedido | MEMO | — |
| RequiereRiesgoDeBiblioteca | TEXT(2) | — |
| CodRiesgoBiblioteca | TEXT(255) | — |
| RiesgoPendienteRetipificacion | TEXT(2) | — |

### TbRiesgosAreasImpacto

- Columnas: **4**. Filas (snapshot): **15**.
- Evidencia en `src`: **1** fichero(s) de producción — `Constructor.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| Tipo | TEXT(255) | — |
| Indice | TEXT(255) | — |
| Descripcion | MEMO | — |
| ordinal | BYTE(1) | — |

### TbRiesgosMaterializaciones

- Columnas: **11**. Filas (snapshot): **32**.
- Evidencia en `src`: **6** fichero(s) de producción — `RiesgoMaterializacion.cls`, `TareasCalidad.cls`, `Constructor.bas`, `IndicadorRiesgosRepositorio.bas`, `Módulo1.bas`, `modMigracionAux.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| ID | LONG(4) | — |
| IDProyecto | LONG(4) | — |
| CodigoRiesgo | TEXT(5) | — |
| Fecha | SHORT_DATE_TIME(8) | — |
| IDEdicion | LONG(4) | — |
| EsMaterializacion | TEXT(2) | — |
| IDNC | LONG(4) | — |
| ParaNC | TEXT(2) | — |
| FechaDecison | SHORT_DATE_TIME(8) | — |
| Estado | TEXT(255) | — |
| IDPlanContingencia | LONG(4) | — |

### TbRiesgosNC

- Columnas: **6**. Filas (snapshot): **32**.
- Evidencia en `src`: **2** fichero(s) de producción — `Constructor.bas`, `Módulo1.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| ID | LONG(4) | — |
| IDRiesgo | LONG(4) | — |
| IDNC | LONG(4) | — |
| FechaRegistro | SHORT_DATE_TIME(8) | — |
| ParaNC | TEXT(2) | — |
| FechaDecison | SHORT_DATE_TIME(8) | — |

### TbRiesgosPorTipoProyecto

- Columnas: **2**. Filas (snapshot): **0**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDTipoProyecto | LONG(4) | — |
| IDCatalogoRiesgo | LONG(4) | — |

### TbRiesgosValoracion

- Columnas: **3**. Filas (snapshot): **25**.
- Evidencia en `src`: **1** fichero(s) de producción — `Constructor.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| Impacto | TEXT(255) | — |
| Vulnerabilidad | TEXT(255) | — |
| Valoracion | TEXT(255) | — |

### TbUltimoProyecto

- Columnas: **4**. Filas (snapshot): **29**.
- Evidencia en `src`: **3** fichero(s) de producción — `Proyecto.cls`, `UltimoProyecto.cls`, `Constructor.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDUltimoProyecto | LONG(4) | — |
| IDProyecto | LONG(4) | — |
| Usuario | TEXT(255) | personal |
| Fecha | SHORT_DATE_TIME(8) | — |
