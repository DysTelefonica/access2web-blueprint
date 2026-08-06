# ERD - condor_datos

Generado: 2026-05-08 09:25

## Tablas (24)

### tbAdjuntos

| Campo | Tipo | TamaÃ±o | Requerido | PK |
|---|---|---|---|---|
| idAdjunto | Long | 4 | Si | PK |
| idSolicitud | Long | 4 | Si |  |
| etapaWF | Text | 50 | No |  |
| nombreArchivo | Text | 255 | Si |  |
| fechaSubida | Date/Time | 8 | Si |  |
| usuarioSubida | Text | 100 | Si |  |
| descripcion | Memo | - | No |  |
| TipoAccion | Text | 255 | No |  |

### TbCorreosEnviados

| Campo | Tipo | TamaÃ±o | Requerido | PK |
|---|---|---|---|---|

### tbDatosCDCA

| Campo | Tipo | TamaÃ±o | Requerido | PK |
|---|---|---|---|---|
| idDatosCDCA | Long | 4 | Si | PK |
| idSolicitud | Long | 4 | Si |  |
| numContrato | Text | 100 | No |  |
| refSuministrador | Text | 100 | No |  |
| SuministradorNombreDir | Memo | - | No |  |
| refDesviacionesPrevias | Text | 100 | No |  |
| requiereModificacionContrato | Boolean | 1 | No |  |
| identificacionMaterial | Memo | - | No |  |
| numPlanoEspecificacion | Memo | - | No |  |
| cantidadPeriodo | Text | 50 | No |  |
| numSerieLote | Text | 100 | No |  |
| causaNC | Memo | - | No |  |
| descripcionImpactoNC | Memo | - | No |  |
| descripcionImpactoNCCont | Memo | - | No |  |
| afectaPrestaciones | Boolean | 1 | No |  |
| afectaSeguridad | Boolean | 1 | No |  |
| afectaFiabilidad | Boolean | 1 | No |  |
| afectaVidaUtil | Boolean | 1 | No |  |
| afectaMedioambiente | Boolean | 1 | No |  |
| afectaIntercambiabilidad | Boolean | 1 | No |  |
| afectaMantenibilidad | Boolean | 1 | No |  |
| afectaApariencia | Boolean | 1 | No |  |
| afectaOtros | Boolean | 1 | No |  |
| impactoCoste | Text | 50 | No |  |
| clasificacionNC | Text | 50 | No |  |
| esSuministradorAD | Boolean | 1 | No |  |
| identificacionAutoridadDiseno | Text | 100 | No |  |
| efectoFechaEntrega | Memo | - | No |  |
| firmaAprobacionRespIngenieriaNombre | Text | 255 | No |  |
| firmaAprobacionRespProduccionNombre | Text | 255 | No |  |
| firmaAprobacionRespCalidadNombre | Text | 255 | No |  |
| firmaAprobacionRespDisenioNombre | Text | 255 | No |  |
| firmaAprobacionRepresentanteSumNombre | Text | 255 | No |  |
| racCodigo | Text | 50 | No |  |
| observacionesRAC | Memo | - | No |  |
| racNombre | Text | 255 | No |  |
| racDecision | Text | 50 | No |  |
| decisionFinal | Text | 50 | No |  |
| observacionesFinales | Memo | - | No |  |
| NombreFirmanteFinal | Text | 100 | No |  |

### tbDatosCDCASUB

| Campo | Tipo | TamaÃ±o | Requerido | PK |
|---|---|---|---|---|
| idDatosCDCASUB | Long | 4 | Si | PK |
| idSolicitud | Long | 4 | Si |  |
| refSuministrador | Text | 100 | No |  |
| refSubSuministrador | Text | 100 | No |  |
| suministradorPrincipalNombreDir | Memo | - | No |  |
| subSuministradorNombreDir | Memo | - | No |  |
| refDesviacionesPrevias | Text | 100 | No |  |
| requiereModificacionContrato | Boolean | 1 | No |  |
| identificacionMaterial | Memo | - | No |  |
| numPlanoEspecificacion | Text | 100 | No |  |
| cantidadPeriodo | Text | 50 | No |  |
| numSerieLote | Text | 100 | No |  |
| causaNC | Memo | - | No |  |
| descripcionImpactoNC | Memo | - | No |  |
| descripcionImpactoNCCont | Memo | - | No |  |
| afectaPrestaciones | Boolean | 1 | No |  |
| afectaSeguridad | Boolean | 1 | No |  |
| afectaFiabilidad | Boolean | 1 | No |  |
| afectaVidaUtil | Boolean | 1 | No |  |
| afectaMedioambiente | Boolean | 1 | No |  |
| afectaIntercambiabilidad | Boolean | 1 | No |  |
| afectaMantenibilidad | Boolean | 1 | No |  |
| afectaApariencia | Boolean | 1 | No |  |
| afectaOtros | Boolean | 1 | No |  |
| impactoCoste | Text | 50 | No |  |
| clasificacionNC | Text | 50 | No |  |
| esSubSuministradorAD | Boolean | 1 | No |  |
| identificacionAutoridadDiseno | Text | 100 | No |  |
| efectoFechaEntrega | Memo | - | No |  |
| firmaAprobacionRespIngenieriaNombre | Text | 255 | No |  |
| firmaAprobacionRespProduccionNombre | Text | 255 | No |  |
| firmaAprobacionRespCalidadNombre | Text | 255 | No |  |
| firmaAprobacionRespDisenioNombre | Text | 255 | No |  |
| firmaAprobacionRepresentanteSumNombre | Text | 255 | No |  |
| racCodigo | Text | 50 | No |  |
| observacionesRAC | Memo | - | No |  |
| racNombre | Text | 255 | No |  |
| racDecision | Text | 50 | No |  |
| racRechazoMotivos | Memo | - | No |  |
| observacionesRACDelegador | Memo | - | No |  |
| racNombreDelegador | Text | 255 | No |  |
| decisionFinal | Text | 50 | No |  |
| observacionesFinales | Memo | - | No |  |
| NombreFirmanteFinal | Text | 100 | No |  |

### tbDatosPC

| Campo | Tipo | TamaÃ±o | Requerido | PK |
|---|---|---|---|---|
| idDatosPC | Long | 4 | Si | PK |
| idSolicitud | Long | 4 | Si |  |
| refContratoInspeccionOficial | Text | 100 | Si |  |
| refSuministrador | Text | 100 | Si |  |
| denominacionContrato | Memo | - | No |  |
| suministradorNombreDir | Memo | - | No |  |
| objetoContrato | Memo | - | No |  |
| descripcionMaterialAfectado | Memo | - | No |  |
| numPlanoEspecificacion | Memo | - | No |  |
| descripcionPropuestaCambio | Memo | - | No |  |
| descripcionPropuestaCambioCont | Memo | - | No |  |
| motivoCorregirDeficiencias | Boolean | 1 | No |  |
| motivoMejorarCapacidad | Boolean | 1 | No |  |
| motivoAumentarNacionalizacion | Boolean | 1 | No |  |
| motivoMejorarSeguridad | Boolean | 1 | No |  |
| motivoMejorarFiabilidad | Boolean | 1 | No |  |
| motivoMejorarCosteEficacia | Boolean | 1 | No |  |
| motivoOtros | Boolean | 1 | No |  |
| motivoOtrosDetalle | Memo | - | No |  |
| incidenciaCoste | Text | 50 | No |  |
| incidenciaPlazo | Text | 50 | No |  |
| incidenciaSeguridad | Boolean | 1 | No |  |
| incidenciaFiabilidad | Boolean | 1 | No |  |
| incidenciaMantenibilidad | Boolean | 1 | No |  |
| incidenciaIntercambiabilidad | Boolean | 1 | No |  |
| incidenciaVidaUtilAlmacen | Boolean | 1 | No |  |
| incidenciaFuncionamientoFuncion | Boolean | 1 | No |  |
| impactoClasificacion | Text | 255 | No |  |
| CambioAfectaAMaterial | Text | 255 | No |  |
| firmaOficinaTecnicaNombre | Text | 100 | No |  |
| firmaRepSuministradorNombre | Text | 100 | No |  |
| racCodigo | Text | 50 | No |  |
| observacionesRAC | Memo | - | No |  |
| racNombre | Text | 255 | No |  |
| racDecision | Text | 50 | No |  |
| racRechazoMotivos | Memo | - | No |  |
| obsAprobacionAutoridadDiseno | Memo | - | No |  |
| NombreAutoridadDiseno | Text | 100 | No |  |
| decisionFinal | Text | 50 | No |  |
| obsDecisionFinal | Memo | - | No |  |
| NombreFirmanteFinal | Text | 100 | No |  |

### tbDatosPCSUB

| Campo | Tipo | TamaÃ±o | Requerido | PK |
|---|---|---|---|---|
| idDatosPCSUB | Long | 4 | Si | PK |
| idSolicitud | Long | 4 | Si |  |
| refContratoInspeccionOficial | Text | 100 | Si |  |
| refSubSuministrador | Text | 100 | No |  |
| denominacionContrato | Memo | - | No |  |
| SubsuministradorNombreDir | Memo | - | No |  |
| objetoContrato | Memo | - | No |  |
| descripcionMaterialAfectado | Memo | - | No |  |
| numPlanoEspecificacion | Memo | - | No |  |
| descripcionPropuestaCambio | Memo | - | No |  |
| descripcionPropuestaCambioCont | Memo | - | No |  |
| motivoCorregirDeficiencias | Boolean | 1 | No |  |
| motivoMejorarCapacidad | Boolean | 1 | No |  |
| motivoAumentarNacionalizacion | Boolean | 1 | No |  |
| motivoMejorarSeguridad | Boolean | 1 | No |  |
| motivoMejorarFiabilidad | Boolean | 1 | No |  |
| motivoMejorarCosteEficacia | Boolean | 1 | No |  |
| motivoOtros | Boolean | 1 | No |  |
| motivoOtrosDetalle | Memo | - | No |  |
| incidenciaCoste | Text | 50 | No |  |
| incidenciaPlazo | Text | 50 | No |  |
| incidenciaSeguridad | Boolean | 1 | No |  |
| incidenciaFiabilidad | Boolean | 1 | No |  |
| incidenciaMantenibilidad | Boolean | 1 | No |  |
| incidenciaIntercambiabilidad | Boolean | 1 | No |  |
| incidenciaVidaUtilAlmacen | Boolean | 1 | No |  |
| incidenciaFuncionamientoFuncion | Boolean | 1 | No |  |
| impactoClasificacion | Text | 255 | No |  |
| CambioAfectaAMaterial | Text | 255 | No |  |
| firmaOficinaTecnicaSubSuministradorNombre | Text | 100 | No |  |
| firmaRepSubSuministradorNombre | Text | 100 | No |  |
| racCodigo | Text | 50 | No |  |
| observacionesRAC | Memo | - | No |  |
| racNombre | Text | 255 | No |  |
| racDecision | Text | 50 | No |  |
| racRechazoMotivos | Memo | - | No |  |
| observacionesRACDelegador | Memo | - | No |  |
| racNombreDelegador | Text | 255 | No |  |
| obsAprobacionAutoridadDiseno | Memo | - | No |  |
| NombreAutoridadDiseno | Text | 100 | No |  |
| decisionFinal | Text | 50 | No |  |
| obsDecisionFinal | Memo | - | No |  |
| NombreFirmanteFinal | Text | 100 | No |  |

### tbEstados

| Campo | Tipo | TamaÃ±o | Requerido | PK |
|---|---|---|---|---|
| idEstado | Long | 4 | No | PK |
| nombreEstado | Text | 50 | Si |  |
| descripcion | Text | 255 | No |  |
| esEstadoInicial | Boolean | 1 | No |  |
| esEstadoFinal | Boolean | 1 | No |  |
| orden | Long | 4 | No |  |

### TbExpedientes

| Campo | Tipo | TamaÃ±o | Requerido | PK |
|---|---|---|---|---|
| IDExpediente | Long | 4 | Si | PK |
| IDExpedientePadre | Long | 4 | No |  |
| OrdinalE2E | Long | 4 | No |  |
| Nemotecnico | Text | 255 | No |  |
| Titulo | Memo | - | No |  |
| ImporteLicitacion | Double | 8 | No |  |
| ImporteContratacion | Double | 8 | No |  |
| CodProyecto | Text | 255 | No |  |
| CodExp | Text | 255 | No |  |
| CodExpLargo | Text | 255 | No |  |
| CodS4H | Text | 255 | No |  |
| FechaInicioContrato | Date/Time | 8 | No |  |
| FechaFinContrato | Date/Time | 8 | No |  |
| FechaFinGarantia | Date/Time | 8 | No |  |
| EsAM | Text | 2 | No |  |
| EsLote | Text | 2 | No |  |
| EsBasado | Text | 255 | No |  |
| EsExpediente | Text | 2 | No |  |
| Ordinal | Text | 255 | No |  |
| IdGradoClasificacion | Long | 4 | No |  |
| IDOrganoContratacion | Long | 4 | No |  |
| IDOficinaPrograma | Long | 4 | No |  |
| IDEjercito | Long | 4 | No |  |
| AccesoSharepoint | Memo | - | No |  |
| Observaciones | Memo | - | No |  |
| FechaCreacion | Date/Time | 8 | No |  |
| IDUsuarioCreacion | Text | 255 | No |  |
| FechaUltimoCambio | Date/Time | 8 | No |  |
| IDUsuarioUltimoCambio | Text | 255 | No |  |
| IDEstado | Long | 4 | No |  |
| Ambito | Text | 10 | No |  |
| NPedido | Text | 255 | No |  |
| IDResponsableCalidad | Long | 4 | No |  |
| Adjudicado | Text | 2 | No |  |
| EnPeriodoDeAdjudicacion | Text | 2 | No |  |
| Tipo | Text | 255 | No |  |
| TipoInforme | Text | 255 | No |  |
| AGEDYSAplica | Text | 2 | No |  |
| AGEDYSGenerico | Text | 2 | No |  |
| HPSAplica | Text | 2 | No |  |
| CadenaPecal | Text | 255 | No |  |
| Pecal | Text | 2 | No |  |
| POSTAGEDO | Text | 2 | No |  |
| FECHAPREOFERTA | Date/Time | 8 | No |  |
| APLICAESTADO | Text | 2 | No |  |
| FECHAINICIOLICITACION | Date/Time | 8 | No |  |
| FECHAOFERTA | Date/Time | 8 | No |  |
| FECHAADJUDICACION | Date/Time | 8 | No |  |
| FECHAFIRMACONTRATO | Date/Time | 8 | No |  |
| GARANTIAMESES | Text | 255 | No |  |
| FECHACERTIFICACION | Date/Time | 8 | No |  |
| FECHAPERDIDA | Date/Time | 8 | No |  |
| FECHADESESTIMADA | Date/Time | 8 | No |  |
| ESTADO | Text | 255 | No |  |
| CodigoActividad | Text | 255 | No |  |
| AplicaTareaS4H | Text | 2 | No |  |
| ContratistaPrincipal | Text | 2 | No |  |
| IDResponsableSeguridad | Long | 4 | No |  |
| ObjetoContrato | Memo | - | No |  |

### TbExpedientesRACS

| Campo | Tipo | TamaÃ±o | Requerido | PK |
|---|---|---|---|---|

### TbExpedientesResponsables

| Campo | Tipo | TamaÃ±o | Requerido | PK |
|---|---|---|---|---|
| IDExpedienteResponsable | Long | 4 | Si | PK |
| IdExpediente | Long | 4 | Si |  |
| IdUsuario | Long | 4 | Si |  |
| CorreoSiempre | Text | 2 | No |  |
| EsJefeProyecto | Text | 2 | No |  |

### TbExpedientesSuministradores

| Campo | Tipo | TamaÃ±o | Requerido | PK |
|---|---|---|---|---|
| IDExpedienteSuministrador | Long | 4 | Si | PK |
| IDExpediente | Long | 4 | Si |  |
| IDSuministrador | Long | 4 | Si |  |
| IDPadre | Long | 4 | No |  |
| Descripcon | Memo | - | No |  |
| ContratistaPrincipal | Text | 2 | No |  |
| SubContratista | Text | 2 | No |  |

### tbHistorialRechazos

| Campo | Tipo | TamaÃ±o | Requerido | PK |
|---|---|---|---|---|
| ID | Long | 4 | No | PK |
| IdSolicitud | Long | 4 | No |  |
| FechaRechazo | Date/Time | 8 | No |  |
| UsuarioCalidad | Text | 255 | No |  |
| AreaAfectada | Text | 255 | No |  |
| MotivoPrincipal | Memo | - | No |  |
| Comentarios | Memo | - | No |  |
| EstaResuelto | Boolean | 1 | No |  |

### tbLogCambios

| Campo | Tipo | TamaÃ±o | Requerido | PK |
|---|---|---|---|---|
| idLogCambio | Long | 4 | Si | PK |
| fechaHora | Date/Time | 8 | Si |  |
| usuario | Text | 100 | Si |  |
| suplantadoPor | Text | 255 | No |  |
| tabla | Text | 50 | Si |  |
| registro | Long | 4 | Si |  |
| campo | Text | 50 | No |  |
| valorAnterior | Memo | - | No |  |
| valorNuevo | Memo | - | No |  |
| tipoOperacion | Text | 255 | Si |  |

### tbLogErrores

| Campo | Tipo | TamaÃ±o | Requerido | PK |
|---|---|---|---|---|
| idLogError | Long | 4 | Si | PK |
| fechaHora | Date/Time | 8 | Si |  |
| usuario | Text | 100 | No |  |
| suplantadoPor | Text | 255 | No |  |
| modulo | Text | 100 | Si |  |
| procedimiento | Text | 100 | No |  |
| numeroError | Long | 4 | Si |  |
| descripcionError | Memo | - | Si |  |
| contexto | Memo | - | No |  |

### tbLogEstados

| Campo | Tipo | TamaÃ±o | Requerido | PK |
|---|---|---|---|---|
| idLogEstado | Long | 4 | No | PK |
| idSolicitud | Long | 4 | No |  |
| idEstadoAnterior | Long | 4 | No |  |
| idEstadoNuevo | Long | 4 | No |  |
| fechaTransicion | Date/Time | 8 | No |  |
| usuarioTransicion | Text | 255 | No |  |

### TbNoConformidades

| Campo | Tipo | TamaÃ±o | Requerido | PK |
|---|---|---|---|---|

### TbRACS

| Campo | Tipo | TamaÃ±o | Requerido | PK |
|---|---|---|---|---|

### tbRechazos

| Campo | Tipo | TamaÃ±o | Requerido | PK |
|---|---|---|---|---|
| idRechazo | Long | 4 | No | PK |
| idSolicitud | Long | 4 | Si |  |
| fechaRechazo | Date/Time | 8 | No |  |
| motivoPrincipal | Text | 255 | No |  |
| areaAfectada | Text | 255 | No |  |
| comentarios | Memo | - | No |  |
| usuarioRechazo | Text | 255 | No |  |
| esActivo | Boolean | 1 | No |  |
| notasSubsanacion | Memo | - | No |  |
| CambiosTecnico | Memo | - | No |  |

### tbSolicitudes

| Campo | Tipo | TamaÃ±o | Requerido | PK |
|---|---|---|---|---|
| idSolicitud | Long | 4 | Si | PK |
| idExpediente | Long | 4 | Si |  |
| tipoSolicitud | Text | 20 | Si |  |
| codigoSolicitud | Text | 50 | Si |  |
| idNCAsociada | Long | 4 | No |  |
| idEstadoInterno | Long | 4 | Si |  |
| fechaCreacion | Date/Time | 8 | Si |  |
| usuarioCreacion | Text | 100 | Si |  |
| fechaModificacion | Date/Time | 8 | No |  |
| usuarioModificacion | Text | 100 | No |  |
| revisionCalidadEstado | Text | 20 | Si |  |
| revisionCalidadComentarios | Memo | - | No |  |

### TbSuministradores

| Campo | Tipo | TamaÃ±o | Requerido | PK |
|---|---|---|---|---|

### tbTransiciones

| Campo | Tipo | TamaÃ±o | Requerido | PK |
|---|---|---|---|---|
| idTransicion | Long | 4 | Si | PK |
| idEstadoOrigen | Long | 4 | Si |  |
| idEstadoDestino | Long | 4 | Si |  |
| rolRequerido | Text | 50 | Si |  |

### TbUsuariosAplicaciones

| Campo | Tipo | TamaÃ±o | Requerido | PK |
|---|---|---|---|---|

### TbUsuariosAplicacionesPermisos

| Campo | Tipo | TamaÃ±o | Requerido | PK |
|---|---|---|---|---|
| CorreoUsuario | Text | 255 | Si | PK |
| IDAplicacion | Long | 4 | Si | PK |
| EsUsuarioAdministrador | Text | 2 | No |  |
| EsUsuarioCalidad | Text | 2 | No |  |
| EsUsuarioEconomia | Text | 2 | No |  |
| EsUsuarioSecretaria | Text | 2 | No |  |
| EsUsuarioTecnico | Text | 2 | No |  |
| EsUsuarioSinAcceso | Text | 2 | No |  |
| EsUsuarioCalidadAvisos | Text | 2 | No |  |

### tbValidacionRevision

| Campo | Tipo | TamaÃ±o | Requerido | PK |
|---|---|---|---|---|
| Id | Long | 4 | No | PK |
| idSolicitud | Long | 4 | No |  |
| ordinal | Integer | 2 | No |  |
| idAdjunto | Long | 4 | No |  |
| FechaEnvio | Date/Time | 8 | No |  |
| FechaRecepcion | Date/Time | 8 | No |  |
| Resultado | Text | 255 | No |  |
| Comentarios | Memo | - | No |  |
| Usuario | Text | 255 | No |  |
| HashDatos | Text | 255 | No |  |

## Relaciones

| Nombre | Tabla origen | Campo origen | Tabla destino | Campo destino |
|---|---|---|---|---|
| MSysNavPaneGroupCategoriesMSysNavPaneGroups | MSysNavPaneGroupCategories | Id | MSysNavPaneGroups | GroupCategoryID |
| tbSolicitudestbAdjuntos | tbSolicitudes | idSolicitud | tbAdjuntos | idSolicitud |
| tbSolicitudestbDatosCDCA | tbSolicitudes | idSolicitud | tbDatosCDCA | idSolicitud |
| tbSolicitudestbDatosCDCASUB | tbSolicitudes | idSolicitud | tbDatosCDCASUB | idSolicitud |
| tbSolicitudestbDatosPC | tbSolicitudes | idSolicitud | tbDatosPC | idSolicitud |

## Backends vinculados no alcanzados

Las siguientes bases de datos vinculadas no estaban disponibles al generar este ERD.
Sus tablas aparecen en el listado de tablas pero su estructura no pudo verificarse.

- `C:\00repos\datos\0Lanzadera\Lanzadera_Datos.accdb` - tablas vinculadas: TbUsuariosAplicaciones

