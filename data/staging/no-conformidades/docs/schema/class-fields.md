# Class Fields Reference -- `src/classes/*.cls`

> Generated: 2026-06-25 from src/classes/*.cls. Public surface only. Use this as the source of truth for entity properties when writing tests, helpers, or operations code -- DO NOT invent properties.

Total classes: 48

---

## ACAuditoria

### Public fields

| Field | Type |
|---|---|
| IdAccionCorrectiva | String |
| id | String |
| NAccion | String |
| AccionCorrectiva | String |
| FechaAccionCorrectiva | String |
| Estado | String |
| FechaInicialMinima | String |
| FechaFinalUltima | String |
| Notas | String |
| Responsable | String |
| FechaFinPrevistaUltima | String |
| Error | String |

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| NAccionARCalculado | Get | ) As String |
| AlgunaAccionIrregular | Get | ) As EnumSino |
| AlgunaAccionIrregular | Set | p_Valor As Variant) |
| EstadoCalculado | Get | ) As EnumEstadoAC |
| EstadoCalculadoTexto | Get | ) As String |
| EstadoEnum | Get | ) As EnumEstadoAC |
| TieneAcciones | Get | ) As EnumSino |
| AlgunaAccionActiva | Get | ) As EnumSino |
| AlgunaAccionActiva | Set | p_Valor As Variant) |
| TodasAccionesFinalizadas | Get | ) As EnumSino |
| TodasAccionesFinalizadas | Set | p_Valor As Variant) |
| TodasAccionesSinPlanificar | Get | ) As EnumSino |
| TodasAccionesSinPlanificar | Set | p_Valor As Variant) |
| AlgunaAccionPteReplanificar | Get | ) As EnumSino |
| AlgunaAccionPteReplanificar | Set | p_Valor As Variant) |
| FechaInicialMinimaCalculada | Get | ) As String |
| FechaInicialMinimaCalculada | Set | p_Valor As Variant) |
| FechaFinalUltimaCalculada | Get | ) As String |
| FechaFinalUltimaCalculada | Set | p_Valor As Variant) |
| FechaFinPrevistaUltimaCalculada | Get | ) As String |
| FechaFinPrevistaUltimaCalculada | Set | p_Valor As Variant) |
| nc | Get | ) As NCAuditoria |
| IDAccionCorrectivaCalculada | Get | ) As String |
| ARs | Get | ) As Scripting |
| ARs | Set | ByVal objNewValue As Variant) |
| ColCampos | Get | ) As Collection |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| EstadoGrabar | Function | _ Optional p_Estado As String, _ Optional ByRef p_Error As String _) As String |
| FechaInicialMinimaGrabar | Function | _ Optional p_Fecha As String, _ Optional ByRef p_Error As String _) As String |
| FechaFinalUltimaGrabar | Function | _ Optional p_Fecha As String, _ Optional ByRef p_Error As String _) As String |
| FechaFinPrevistaUltimaGrabar | Function | _ Optional p_Fecha As String, _ Optional ByRef p_Error As String _) As String |
| getPropiedad | Function | m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant |
| SetPropiedad | Function | m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String |

---

## ACAuditoriaOperaciones

### Public fields

| Field | Type |
|---|---|
| AC | ACAuditoria |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| MotivoNoOK | Function | _ Optional ByRef p_ObjACAlInicio As ACAuditoria, _ Optional ByRef p_Error As String _) As String |
| Registrar | Function | _ Optional ByRef p_ObjACAlInicio As ACAuditoria, _ Optional ByRef p_Error As String _) As String |
| RegistrarNAccion | Function | _ Optional ByRef p_Error As String _) As String |
| Eliminar | Function | Optional ByRef p_Error As String) As String |
| ActualizarDatosCalculados | Function | _ Optional ByRef p_Error As String _) As String |
| ReplanificarAcciones | Function | _ Optional p_Fecha As String, _ Optional ByRef p_Error As String _) As String |
| RegistrarResponsable | Function | _ Optional ByRef p_Error As String _) As String |

---

## ACProyecto

### Public fields

| Field | Type |
|---|---|
| IdAccionCorrectiva | String |
| IDNoConformidad | String |
| NAccion | String |
| AccionCorrectiva | String |
| Estado | String |
| FechaAccionCorrectiva | String |
| FechaInicialMinima | String |
| FechaFinalUltima | String |
| FechaFinPrevistaUltima | String |
| Notas | String |
| Error | String |

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| ResponsableObj | Get | ) As usuario |
| ResponsableObj | Set | ByVal objNewValue As Variant) |
| NAccionARCalculado | Get | ) As String |
| AlgunaAccionIrregular | Get | ) As EnumSino |
| AlgunaAccionIrregular | Set | p_Valor As Variant) |
| EstadoCalculado | Get | ) As EnumEstadoAC |
| EstadoCalculadoTexto | Get | ) As String |
| EstadoEnum | Get | ) As EnumEstadoAC |
| TieneAcciones | Get | ) As EnumSino |
| AlgunaAccionActiva | Get | ) As EnumSino |
| AlgunaAccionActiva | Set | p_Valor As Variant) |
| TodasAccionesFinalizadas | Get | ) As EnumSino |
| TodasAccionesFinalizadas | Set | p_Valor As Variant) |
| TodasAccionesSinPlanificar | Get | ) As EnumSino |
| TodasAccionesSinPlanificar | Set | p_Valor As Variant) |
| AlgunaAccionPteReplanificar | Get | ) As EnumSino |
| AlgunaAccionPteReplanificar | Set | p_Valor As Variant) |
| FechaInicialMinimaCalculada | Get | ) As String |
| FechaInicialMinimaCalculada | Set | p_Valor As Variant) |
| FechaFinalUltimaCalculada | Get | ) As String |
| FechaFinalUltimaCalculada | Set | p_Valor As Variant) |
| FechaFinPrevistaUltimaCalculada | Get | ) As String |
| FechaFinPrevistaUltimaCalculada | Set | p_Valor As Variant) |
| nc | Get | ) As NCProyecto |
| IDAccionCorrectivaCalculada | Get | ) As String |
| ARs | Get | ) As Scripting |
| ARs | Set | ByVal objNewValue As Variant) |
| ColCampos | Get | ) As Collection |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| EstadoGrabar | Function | _ Optional p_Estado As String, _ Optional ByRef p_Error As String _) As String |
| FechaInicialMinimaGrabar | Function | _ Optional p_Fecha As String, _ Optional ByRef p_Error As String _) As String |
| FechaFinalUltimaGrabar | Function | _ Optional p_Fecha As String, _ Optional ByRef p_Error As String _) As String |
| FechaFinPrevistaUltimaGrabar | Function | _ Optional p_Fecha As String, _ Optional ByRef p_Error As String _) As String |
| getPropiedad | Function | m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant |
| SetPropiedad | Function | m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String |

---

## ACProyectoOperaciones

### Public fields

| Field | Type |
|---|---|
| AC | ACProyecto |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| MotivoNoOK | Function | _ Optional ByRef p_ObjACAlInicio As ACProyecto, _ Optional ByRef p_Error As String _) As String |
| Registrar | Function | _ Optional ByRef p_ObjACAlInicio As ACProyecto, _ Optional ByRef p_Error As String _) As String |
| RegistrarNAccion | Function | _ Optional ByRef p_Error As String _) As String |
| Eliminar | Function | Optional ByRef p_Error As String) As String |
| ActualizarDatosCalculados | Function | _ Optional ByRef p_Error As String _) As String |
| ReplanificarAcciones | Function | _ Optional p_Fecha As String, _ Optional ByRef p_Error As String _) As String |
| RegistrarResponsable | Function | _ Optional ByRef p_Error As String _) As String |

---

## ARAuditoria

### Public fields

| Field | Type |
|---|---|
| IDAccionRealizada | String |
| IdAccionCorrectiva | String |
| NAccion | String |
| AccionRealizada | String |
| FechaAccionRealizada | String |
| FechaInicio | String |
| FechaFinPrevista | String |
| FechaFinReal | String |
| Estado | String |
| Notas | String |
| Error | String |

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| ResponsableObj | Get | ) As usuario |
| Documentos | Get | ) As Scripting |
| Documentos | Set | ByVal objNewValue As Variant) |
| EstadoCalculado | Get | ) As EnumEstadoAR |
| EstadoCalculadoTexto | Get | ) As String |
| EstadoEnum | Get | ) As EnumEstadoAR |
| PendienteDeRetipificar | Get | ) As EnumSino |
| AC | Get | ) As ACAuditoria |
| IDAccionRealizadaCalculada | Get | ) As String |
| ColCampos | Get | ) As Collection |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| EstadoGrabar | Function | _ Optional p_Estado As String, _ Optional ByRef p_Error As String _) As String |
| getPropiedad | Function | m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant |
| SetPropiedad | Function | m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String |

---

## ARAuditoriaOperaciones

### Public fields

| Field | Type |
|---|---|
| AR | ARAuditoria |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| Registrar | Function | _ Optional ByRef p_ObjARAlInicio As ARAuditoria, _ Optional p_Observaciones As String, _ Optional ByRef p_Error As String _) As String |
| Eliminar | Function | Optional ByRef p_Error As String) As String |
| ActualizarDatosCalculados | Function | _ Optional ByRef p_Error As String _) As String |
| Replanificar | Function | _ ByVal p_Fecha As String, _ ByVal p_Observaciones As String, _ Optional ByRef p_Error As String _) As String |
| RegistrarNAccion | Function | _ Optional ByRef p_Error As String _) As String |
| RegistrarResponsable | Function | _ Optional ByRef p_Error As String _) As String |

---

## ARProyecto

### Public fields

| Field | Type |
|---|---|
| IDAccionRealizada | String |
| IdAccionCorrectiva | String |
| NAccion | String |
| AccionRealizada | String |
| FechaAccionRealizada | String |
| FechaInicio | String |
| FechaFinPrevista | String |
| FechaFinReal | String |
| Estado | String |
| Notas | String |
| Error | String |

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| ResponsableObj | Get | ) As usuario |
| ResponsableObj | Let | p_Value As Variant) |
| Documentos | Get | ) As Scripting |
| Documentos | Set | ByVal objNewValue As Variant) |
| EstadoCalculado | Get | ) As EnumEstadoAR |
| EstadoCalculadoTexto | Get | ) As String |
| EstadoEnum | Get | ) As EnumEstadoAR |
| PendienteDeRetipificar | Get | ) As EnumSino |
| AC | Get | ) As ACProyecto |
| IDAccionRealizadaCalculada | Get | ) As String |
| ColCampos | Get | ) As Collection |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| EstadoGrabar | Function | _ Optional p_Estado As String, _ Optional ByRef p_Error As String _) As String |
| getPropiedad | Function | m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant |
| SetPropiedad | Function | m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String |

---

## ARProyectoOperaciones

### Public fields

| Field | Type |
|---|---|
| AR | ARProyecto |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| Registrar | Function | _ Optional ByRef p_ObjARAlInicio As ARProyecto, _ Optional p_Observaciones As String, _ Optional ByRef p_Error As String _) As String |
| Eliminar | Function | Optional ByRef p_Error As String) As String |
| ActualizarDatosCalculados | Function | _ Optional ByRef p_Error As String _) As String |
| Replanificar | Function | _ ByVal p_Fecha As String, _ ByVal p_Observaciones As String, _ Optional ByRef p_Error As String _) As String |

---

## Auditoria

### Public fields

| Field | Type |
|---|---|
| IDAuditoria | String |
| Tipo | String |
| FechaInicio | String |
| FechaFin | String |
| Error | String |

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| NCs | Get | ) As Scripting |
| NCs | Set | ByVal objNewValue As Variant) |
| NombreAuditoria | Get | ) As String |
| URLDirectorio | Get | ) As String |
| IDAuditoriaCalculada | Get | ) As String |
| Documentos | Get | ) As Scripting |
| Documentos | Set | ByVal objNewValue As Variant) |
| ColCampos | Get | ) As Collection |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| NumeroNCCalculado | Function | p_Tipo As String, Optional ByRef p_Error As String) As String |
| getPropiedad | Function | m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant |
| SetPropiedad | Function | m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String |

---

## AuditoriaOperaciones

### Public fields

| Field | Type |
|---|---|
| Auditoria | Auditoria |
| Error | String |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| Registrar | Function | _ Optional ByRef p_AuditoriaAlInicio As Auditoria, _ Optional ByRef p_Error As String _) As String |
| Eliminar | Function | _ Optional ByRef p_Error As String _) As String |
| AnexoMultiple | Function | _ Optional ByRef p_Anexos As Scripting.Dictionary, _ Optional ByRef p_Error As String _) As String |

---

## CacheNCCacheRepositorio

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| UpsertDetalle | Function | ByVal p_IDNC As Long, ByRef p_Valores As Scripting.Dictionary, Optional ByRef p_Error As String) As Boolean |
| ActualizarCampoDetalle | Function | ByVal p_IDNC As Long, ByVal p_Campo As EnumCampoCache, ByVal p_Valor As Variant, Optional ByRef p_Error As String) As Boolean |
| EliminarDetalle | Function | ByVal p_IDNC As Long, Optional ByRef p_Error As String) As Boolean |
| UpsertListado | Function | ByVal p_IDNC As Long, Optional ByRef p_Error As String) As Boolean |
| EliminarListado | Function | ByVal p_IDNC As Long, Optional ByRef p_Error As String) As Boolean |
| GetDetalleValido | Function | ByVal p_IDNC As Long, Optional ByRef p_Error As String) As Boolean |
| GetListadoValido | Function | ByVal p_IDNC As Long, Optional ByRef p_Error As String) As Boolean |

---

## CacheNCCrud

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| NotificarAltaNC | Function | ByVal p_IDNC As Long, Optional ByRef p_Error As String) As Boolean |
| NotificarModificacionNC | Function | ByVal p_IDNC As Long, ByRef p_Campos As Collection, Optional ByRef p_Error As String) As Boolean |
| NotificarEliminacionNC | Function | ByVal p_IDNC As Long, Optional ByRef p_Error As String) As Boolean |
| NotificarCambioACAR | Function | ByVal p_IDNC As Long, Optional ByRef p_Error As String) As Boolean |
| NotificarCambioDocumentos | Function | ByVal p_IDNC As Long, Optional ByRef p_Error As String) As Boolean |
| NotificarCambioRiesgos | Function | ByVal p_IDNC As Long, Optional ByRef p_Error As String) As Boolean |
| NotificarCambioReplanificaciones | Function | ByVal p_IDNC As Long, Optional ByRef p_Error As String) As Boolean |
| NotificarCierreNC | Function | ByVal p_IDNC As Long, Optional ByRef p_Error As String) As Boolean |

---

## CacheNCService

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| NotificarCambioNC | Function | ByVal p_IDNC As Long, ByVal p_Campo As EnumCampoCache, Optional ByRef p_Error As String) As Boolean |
| NotificarCambioMultiCampo | Function | ByVal p_IDNC As Long, ByRef p_Campos As Collection, Optional ByRef p_Error As String) As Boolean |
| NotificarEliminacionNC_Impl | Function | ByVal p_IDNC As Long, Optional ByRef p_Error As String) As Boolean |
| NotificarAltaNC_Impl | Function | ByVal p_IDNC As Long, Optional ByRef p_Error As String) As Boolean |
| SincronizarCacheNC | Function | ByVal p_IDNC As Long, Optional ByRef p_Error As String) As Boolean |

---

## CierreNCService

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| CerrarNC | Function | ByVal p_IDNC As Long, ByRef p_Error As String) As Boolean |

---

## Correo

### Public fields

| Field | Type |
|---|---|
| IDCorreo | String |
| Originador | String |
| DESTINATARIOS | String |
| DestinatariosConCopia | String |
| DestinatariosConCopiaOculta | String |
| Asunto | String |
| Cuerpo | String |
| FechaEnvio | String |
| FechaGrabacion | String |
| IDEdicion | String |
| Error | String |

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| IDCorreoCalculado | Get | ) As String |
| ColCampos | Get | ) As Collection |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| Registrar | Function | Optional ByRef p_Error As String) As String |
| getPropiedad | Function | m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant |
| SetPropiedad | Function | m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String |

---

## DocumentoAuditoria

### Public fields

| Field | Type |
|---|---|
| IDDocumento | String |
| IDNoConformidad | String |
| Documento | String |
| NombreAnexo | String |
| IDAccionRealizada | String |
| IDAuditoria | String |
| IDAuditoriaResultante | String |
| Error | String |

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| IDAuditoriaResultanteCalculado | Get | ) As String |
| AuditoriaResultante | Get | ) As Auditoria |
| URLDirectorioAnexo | Get | ) As String |
| URLAnexo | Get | ) As String |
| IDDocumentoCalculado | Get | ) As String |
| TipoAnexo | Get | ) As EnumTipoAnexoAuditoria |
| AR | Get | ) As ARAuditoria |
| nc | Get | ) As NCAuditoria |
| Auditoria | Get | ) As Auditoria |
| ColCampos | Get | ) As Collection |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| getURLAnexoFinal | Function | _ p_URLLocal As String, _ Optional p_Creandolo As EnumSino = EnumSino.Sí, _ Optional ByRef p_Error As String _) As String |
| getURLDirectorioAnexo | Function | _ p_Creandolo As EnumSino, _ Optional ByRef p_Error As String _) As String |
| getPropiedad | Function | m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant |
| SetPropiedad | Function | m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String |

---

## DocumentoAuditoriaOperaciones

### Public fields

| Field | Type |
|---|---|
| Documento | DocumentoAuditoria |
| Error | String |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| Registrar | Function | _ p_URLArchivoLocal As String, _ Optional ByRef p_Error As String _) As String |
| MotivoNoOK | Function | _ p_URLArchivoLocal As String, _ Optional ByRef p_Error As String _) As String |
| Eliminar | Function | _ Optional ByRef p_Error As String _) As String |
| CambiarNombre | Function | _ p_NombreDocumento As String, _ Optional ByRef p_Error As String _) As String |

---

## DocumentoProyecto

### Public fields

| Field | Type |
|---|---|
| IDDocumento | String |
| IDNoConformidad | String |
| Documento | String |
| NombreAnexo | String |
| IDAccionRealizada | String |
| IDNoConformidadResultante | String |
| Error | String |

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| IDNoConformidadResultanteCalculado | Get | ) As String |
| IDDocumentoCalculado | Get | ) As String |
| URLAnexo | Get | ) As String |
| nc | Get | ) As NCProyecto |
| AR | Get | ) As ARProyecto |
| ColCampos | Get | ) As Collection |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| getURLAnexoFinal | Function | _ p_URLLocal As String, _ Optional ByRef p_Error As String _) As String |
| getPropiedad | Function | m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant |
| SetPropiedad | Function | m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String |

---

## DocumentoProyectoOperaciones

### Public fields

| Field | Type |
|---|---|
| Documento | DocumentoProyecto |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| Eliminar | Function | _ Optional ByRef p_Error As String _) As String |
| MotivoNoOK | Function | _ p_URLArchivoLocal As String, _ Optional ByRef p_Error As String _) As String |
| Registrar | Function | _ p_URLArchivoLocal As String, _ Optional ByRef p_Error As String _) As String |
| CambiarNombre | Function | _ p_NombreDocumento As String, _ Optional ByRef p_Error As String _) As String |

---

## DocumentoService

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| Registrar | Function | ByRef p_Documento As DocumentoProyecto, Optional ByRef p_Error As String) As Boolean |
| Eliminar | Function | ByRef p_Documento As DocumentoProyecto, Optional ByRef p_Error As String) As Boolean |

---

## Entorno

### Public fields

| Field | Type |
|---|---|
| URLUltimoArchivo | String |
| Error | String |

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| ColEnumOrdenOrderBy | Get | ) As Scripting |
| ColEnumOrdenTitulo | Get | ) As Scripting |
| ColEstadosAR | Get | ) As Scripting |
| ColEstadosARTexto | Get | ) As Scripting |
| ColEstadosARTitulo | Get | ) As Scripting |
| ColEstadosAC | Get | ) As Scripting |
| ColEstadosACTexto | Get | ) As Scripting |
| ColEstadosACTitulo | Get | ) As Scripting |
| UsuariosActivos | Get | ) As Scripting |
| UsuariosActivos | Set | p_Valor As Variant) |
| Usuarios | Get | ) As Scripting |
| Usuarios | Set | p_Valor As Variant) |
| TituloUsuarioConectado | Get | ) As String |
| CSS | Get | ) As String |
| CSS1 | Get | ) As String |
| URLDirectorioAplicaciones | Get | ) As String |
| URLArchivoIni | Get | ) As String |
| URLPlantillaInformeNCProyecto | Get | ) As String |
| URLPlantillaInformeNCAuditoria | Get | ) As String |
| URLVideoAyudaNCProyecto | Get | ) As String |
| URLVideoAyudaNCProyectoParaTecnico | Get | ) As String |
| URLVideoAyudaNCAuditoria | Get | ) As String |
| URLDirAplicacion | Get | ) As String |
| URLArchivoCSS | Get | ) As String |
| URLArchivoCSS1 | Get | ) As String |
| URLAyudaIndicadores | Get | ) As String |
| ColAdministradores | Get | ) As Scripting |
| Expedientes | Get | ) As Scripting |
| ColUsuariosCalidad | Get | ) As Scripting |
| CadenaCorreosCalidad | Get | ) As String |
| CadenaCorreosCalidad | Let | p_Valor As String) |
| CadenaCorreosCalidadEnPruebas | Get | ) As String |
| VersionAplicacion | Get | ) As String |
| URLDirectorioDocumentacion | Get | ) As String |
| URLDirectorioPlantillas | Get | ) As String |
| URLDirectorioDocumentacionAuditorias | Get | ) As String |
| URLDirectorioLocal | Get | ) As String |
| URLDirectorioInformesLocal | Get | ) As String |
| URLDirectorioDocumentacionAyuda | Get | ) As String |
| URLDirRecursos | Get | ) As String |
| ColEstadosNC | Get | ) As Scripting |
| ColEstadosNCTexto | Get | ) As Scripting |
| ColEstadosNCTitulo | Get | ) As Scripting |
| ColNCsProyecto | Get | ) As Scripting |
| ColNCsProyecto | Set | ByVal objNewValue As Variant) |
| ColNCsAuditoria | Get | ) As Scripting |
| ColNCsAuditoria | Set | ByVal objNewValue As Variant) |
| ColJuridicasDistintas | Get | ) As Scripting |
| ColJuridicasDistintas | Set | ByVal objNewValue As Variant) |
| ColJefesProyecto | Get | ) As Scripting |
| ColAuditoriaSegTareas | Get | ) As Scripting |
| ColAuditoriaSegTareas | Set | ByVal objNewValue As Variant) |
| ColAuditoriaSegTareasActivas | Get | ) As Scripting |
| ColAuditoriaSegTareasActivas | Set | ByVal objNewValue As Variant) |
| ColAuditoriaSegNC | Get | ) As Scripting |
| ColAuditoriaSegNC | Set | ByVal objNewValue As Variant) |
| ColProyectoSegTareas | Get | ) As Scripting |
| ColProyectoSegTareas | Set | ByVal objNewValue As Variant) |
| ColProyectoSegNC | Get | ) As Scripting |
| ColProyectoSegNC | Set | ByVal objNewValue As Variant) |
| ColSegsTareasProyecto | Get | ) As Scripting |
| ColSegsTareasProyecto | Set | ByVal objNewValue As Variant) |
| ColSegsTareasProyectoActivas | Get | ) As Scripting |
| ColSegsTareasProyectoActivas | Set | ByVal objNewValue As Variant) |
| ColSegsTareasProyectoPteReplanificar | Get | ) As Scripting |
| ColSegsTareasProyectoPteReplanificar | Set | ByVal objNewValue As Variant) |
| ColSegsNCProyectoRegistradas | Get | ) As Scripting |
| ColSegsNCProyectoRegistradas | Set | ByVal objNewValue As Variant) |
| ColSegsNCProyectoTotales | Get | ) As Scripting |
| ColSegsNCProyectoTotales | Set | ByVal objNewValue As Variant) |
| ColSegsNCProyectoAccionesSinTareas | Get | ) As Scripting |
| ColSegsNCProyectoAccionesSinTareas | Set | ByVal objNewValue As Variant) |
| ColSegsNCProyectoPteCE | Get | ) As Scripting |
| ColSegsNCProyectoPteCE | Set | ByVal objNewValue As Variant) |
| ColSegsNCProyectoCECaducada | Get | ) As Scripting |
| ColSegsNCProyectoCECaducada | Set | ByVal objNewValue As Variant) |
| ColSegsNCProyectoCENoConforme | Get | ) As Scripting |
| ColSegsNCProyectoCENoConforme | Set | ByVal objNewValue As Variant) |
| ColSegsTareasAuditoriaTotales | Get | ) As Scripting |
| ColSegsTareasAuditoriaTotales | Set | ByVal objNewValue As Variant) |
| ColSegsTareasAuditoriaActivas | Get | ) As Scripting |
| ColSegsTareasAuditoriaActivas | Set | ByVal objNewValue As Variant) |
| ColSegsTareasAuditoriaPteReplanificar | Get | ) As Scripting |
| ColSegsTareasAuditoriaPteReplanificar | Set | ByVal objNewValue As Variant) |
| ColSegsTareasAuditoriaIrregulares | Get | ) As Scripting |
| ColSegsTareasAuditoriaIrregulares | Set | ByVal objNewValue As Variant) |
| ColSegsNCAuditoriaRegistradas | Get | ) As Scripting |
| ColSegsNCAuditoriaRegistradas | Set | ByVal objNewValue As Variant) |
| ColSegsNCAuditoriaTotales | Get | ) As Scripting |
| ColSegsNCAuditoriaTotales | Set | ByVal objNewValue As Variant) |
| ColSegsNCAuditoriaAccionesSinTareas | Get | ) As Scripting |
| ColSegsNCAuditoriaAccionesSinTareas | Set | ByVal objNewValue As Variant) |
| ColSegsNCAuditoriaPteCE | Get | ) As Scripting |
| ColSegsNCAuditoriaPteCE | Set | ByVal objNewValue As Variant) |
| ColSegsNCAuditoriaCECaducada | Get | ) As Scripting |
| ColSegsNCAuditoriaCECaducada | Set | ByVal objNewValue As Variant) |
| ColSegsNCAuditoriaCENoConforme | Get | ) As Scripting |
| ColSegsNCAuditoriaCENoConforme | Set | ByVal objNewValue As Variant) |
| ColTipos | Get | ) As Scripting |
| ColTipos | Set | ByVal objNewValue As Variant) |
| ColAuditorias | Get | ) As Scripting |
| ColAuditorias | Set | ByVal objNewValue As Variant) |
| ColResponsablesImplantacionDistintos | Get | ) As Scripting |
| ColResponsablesImplantacionDistintos | Set | ByVal objNewValue As Variant) |
| ColAuditoriasNombres | Get | ) As Scripting |
| ColAuditoriasNombres | Set | ByVal objNewValue As Variant) |
| ColPuntosNormaNCAuditorias | Get | ) As Scripting |
| ColPuntosNormaNCAuditorias | Set | p_Valor As Variant) |
| ColItems | Get | ) As Scripting |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| ValidarInfraCritica | Function | Optional ByRef p_Error As String) As Boolean |
| getPropiedad | Function | _ p_NombreCampo As Variant, _ Optional ByRef p_Error As String _) As Variant |
| InvalidateCombosCache | Sub | ) |

---

## Expediente

### Public fields

| Field | Type |
|---|---|
| IDExpediente | String |
| IDExpedientePadre | String |
| Nemotecnico | String |
| Titulo | String |
| ImporteLicitacion | String |
| ImporteContratacion | String |
| CodProyecto | String |
| CodExp | String |
| CodExpLargo | String |
| CodS4H | String |
| FechaInicioContrato | String |
| FechaFinContrato | String |
| FechaFinGarantia | String |
| EsAM | String |
| EsLote | String |
| EsBasado | String |
| EsExpediente | String |
| Ordinal | String |
| IdGradoClasificacion | String |
| IDOrganoContratacion | String |
| IDOficinaPrograma | String |
| IDEjercito | String |
| AccesoSharepoint | String |
| Observaciones | String |
| FechaCreacion | String |
| IDUsuarioCreacion | String |
| FechaUltimoCambio | String |
| IDUsuarioUltimoCambio | String |
| Ambito | String |
| NPedido | String |
| IDResponsableCalidad | String |
| AGEDYSAplica | String |
| AGEDYSGenerico | String |
| HPSAplica | String |
| Tipo | String |
| CadenaPecal | String |
| PECAL | String |
| TipoInforme | String |
| POSTAGEDO | String |
| APLICAESTADO | String |
| FECHAINICIOLICITACION | String |
| FECHAOFERTA | String |
| FECHAADJUDICACION | String |
| FechaFirmaContrato | String |
| GARANTIAMESES | String |
| FECHACERTIFICACION | String |
| FECHAPERDIDA | String |
| FECHADESESTIMADA | String |
| Estado | String |
| CodigoActividad | String |
| Error | String |

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| Riesgos | Get | ) As Scripting |
| TextoExpediente | Get | ) As String |
| EnUTE | Get | ) As EnumSino |
| Juridicas | Get | ) As Scripting |
| CadenaJuridicas | Get | ) As String |
| Responsables | Get | ) As Scripting |
| RESPONSABLECALIDAD | Get | ) As usuario |
| JefeProyecto | Get | ) As usuario |
| JefeProyectoNombre | Get | ) As String |
| ColCampos | Get | ) As Collection |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| getPropiedad | Function | m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant |
| SetPropiedad | Function | m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String |

---

## ExpedienteResponsable

### Public fields

| Field | Type |
|---|---|
| IDExpedienteResponsable | String |
| IDExpediente | String |
| IdUsuario | String |
| CorreoSiempre | String |
| EsJefeProyecto | String |
| Error | String |

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| IDExpedienteResponsableCalculado | Get | ) As String |
| usuario | Get | ) As usuario |
| Expediente | Get | ) As Expediente |
| ColCampos | Get | ) As Collection |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| getPropiedad | Function | m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant |
| SetPropiedad | Function | m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String |

---

## IndicadorServicio

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| CalcularFechasPeriodo | Sub | p_Semestre As String, p_Anio As Integer, ByRef dInicio As Date, ByRef dFin As Date) |
| TotalAbiertasPeriodo | Function | dIni As Date, dFin As Date, Optional sIdsIgnored As String = "") As Long |
| TotalReplanificadas | Function | dIni As Date, dFin As Date, Optional sIdsIgnored As String = "") As Long |
| TotalStockActivo | Function | dIni As Date, dFin As Date, Optional sIdsIgnored As String = "") As Long |
| TotalConRiesgo | Function | dIni As Date, dFin As Date, Optional sIdsIgnored As String = "") As Long |
| ObtenerDetalleIndicador | Function | sTipo As String, dIni As Date, dFin As Date, Optional sIdsIgnored As String = "", Optional ByRef p_Error As String) As Scripting |

---

## Informe

### Public fields

| Field | Type |
|---|---|
| Error | String |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| GenerarWordNoConformidades | Function | _ p_EsDeProyecto As EnumSino, _ Optional p_col As Scripting.Dictionary, _ Optional p_NC As Object, _ Optional p_Error As String _) As String |
| PrepararPlantilla | Function | _ p_EsDeProyecto As EnumSino, _ Optional p_Error As String _) As String |

---

## Juridica

### Public fields

| Field | Type |
|---|---|
| IDJuridica | String |
| Juridica | String |
| Descripcion | String |
| Error | String |

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| IDJuridicaCalculado | Get | ) As String |
| ColCampos | Get | ) As Collection |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| getPropiedad | Function | m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant |
| SetPropiedad | Function | m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String |

---

## LogNCAuditoria

### Public fields

| Field | Type |
|---|---|
| IDLog | String |
| idNC | String |
| idAC | String |
| idAR | String |
| usuario | String |
| FECHA | String |
| Titulo | String |
| Linea | String |
| col | Collection |
| Error | String |

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| nc | Get | ) As NCAuditoria |
| AC | Get | ) As ACAuditoria |
| AR | Get | ) As ARAuditoria |
| UsuarioObj | Get | ) As usuario |
| IDLogCalculado | Get | ) As String |
| LineaResultante | Get | ) As String |
| ColCampos | Get | ) As Collection |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| Alta | Function | _ Optional ByRef p_Error As String _) As String |
| getPropiedad | Function | m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant |
| SetPropiedad | Function | m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String |

---

## LogNCProyecto

### Public fields

| Field | Type |
|---|---|
| IDLog | String |
| idNC | String |
| idAC | String |
| idAR | String |
| usuario | String |
| FECHA | String |
| Titulo | String |
| Linea | String |
| col | Collection |
| Error | String |

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| nc | Get | ) As NCProyecto |
| AC | Get | ) As ACProyecto |
| AR | Get | ) As ARProyecto |
| UsuarioObj | Get | ) As usuario |
| IDLogCalculado | Get | ) As String |
| LineaResultante | Get | ) As String |
| ColCampos | Get | ) As Collection |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| Alta | Function | _ Optional ByRef p_Error As String _) As String |
| getPropiedad | Function | m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant |
| SetPropiedad | Function | m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String |

---

## NCAuditoria

### Public fields

| Field | Type |
|---|---|
| id | String |
| IDAuditoria | String |
| FechaApertura | String |
| Numero | String |
| Descripcion | String |
| CAUSARAIZ | String |
| AccionCorrectiva | String |
| CORRECCION | String |
| FECHACIERRE | String |
| FPREVCIERRE | String |
| RequiereControlEficacia | String |
| ControlEficacia | String |
| FechaControlEficacia | String |
| FechaPrevistaControlEficacia | String |
| ResultadoControlEficacia | String |
| ConformeControlEficacia | String |
| RequiereAccionCorrectiva | String |
| MotivoNoAccionCorrectiva | String |
| MotivoNoRequiereControlEficacia | String |
| Tipo | String |
| PuntoNorma | String |
| Estado | String |
| Borrado | Boolean |
| MotivoBorrado | String |
| Notas | String |
| Cerrada | String |
| Error | String |

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| Particula | Get | ) As String |
| Titulo | Get | ) As String |
| NAccionCalculado | Get | ) As String |
| FECHACIERRECalculada | Get | ) As String |
| FPREVCIERRECalculada | Get | ) As String |
| CerradaCalculada | Get | ) As EnumSino |
| TieneAccionesPorReplanificar | Get | ) As EnumSino |
| EstadoCalculado | Get | ) As EnumEstadoNC |
| EstadoCalculadoTexto | Get | ) As String |
| ACs | Get | ) As Scripting |
| ACs | Set | ByVal objNewValue As Variant) |
| ARsSinFinalizar | Get | ) As Scripting |
| ARsSinFinalizar | Set | ByVal objNewValue As Variant) |
| ACsSinAR | Get | ) As Scripting |
| ACsSinAR | Set | ByVal objNewValue As Variant) |
| TodasLasArsFinalizadas | Get | ) As EnumSino |
| AlgunaACSinAR | Get | ) As EnumSino |
| TodasLasACsSinFechas | Get | ) As EnumSino |
| BorradoCalculado | Get | ) As EnumSino |
| RequiereACRCalculado | Get | ) As EnumSino |
| RequiereControlEficaciaCalculado | Get | ) As EnumSino |
| ConControlDeEficaciaSinRellenar | Get | ) As EnumSino |
| EficaciaOK | Get | ) As EnumSino |
| AccionesOK | Get | ) As EnumSino |
| TieneAcciones | Get | ) As EnumSino |
| Documentos | Get | ) As Scripting |
| Documentos | Set | ByVal objNewValue As Variant) |
| DocumentosCompletos | Get | ) As Scripting |
| DocumentosCompletos | Set | ByVal objNewValue As Variant) |
| Auditoria | Get | ) As Auditoria |
| IDCalculado | Get | ) As String |
| EstadoEnum | Get | ) As EnumEstadoNC |
| EstadoTitulo | Get | ) As String |
| Replanificaciones | Get | ) As Scripting |
| Replanificaciones | Set | p_Valor As Variant) |
| ColCampos | Get | ) As Collection |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| FPREVCIERREGrabar | Function | _ Optional p_Fecha As String, _ Optional ByRef p_Error As String _) As String |
| FECHACIERREGrabar | Function | _ Optional p_Fecha As String, _ Optional ByRef p_Error As String _) As String |
| CIERREGrabar | Function | _ Optional p_Cerrada As EnumSino, _ Optional ByRef p_Error As String _) As String |
| DatosGeneralesOK | Function | _ Optional p_MenosCef As EnumSino = EnumSino.No, _ Optional ByRef p_Error As String _) As EnumSino |
| EstadoGrabar | Function | _ Optional p_Estado As String, _ Optional ByRef p_Error As String _) As String |
| getPropiedad | Function | m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant |
| SetPropiedad | Function | m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String |

---

## NCaUDITORIAOperaciones

### Public fields

| Field | Type |
|---|---|
| nc | NCAuditoria |
| Error | String |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| MotivoDatosUnicosNoOK | Function | _ Optional ByRef p_ObjNCAlInicio As NCAuditoria, _ Optional ByRef p_Error As String, _ Optional ByVal p_MenosCef As EnumSino = EnumSino.No _) As String |
| RegistrarDatosUnicos | Function | _ Optional ByRef p_ObjNCAlInicio As NCAuditoria, _ Optional ByRef p_Error As String, _ Optional ByVal p_MenosCef As EnumSino = EnumSino.No, _ Optional ByRef p_PromptResult As Long = -1, _ Optional ByRef p_MessageText As String _) As String |
| Eliminar | Function | _ Optional ByVal p_SinDejarRastro As EnumSino = EnumSino.No, _ Optional ByRef p_Error As String _) As String |
| Habilitar | Function | _ Optional ByRef p_Error As String _) As String |
| RegistrarControlEficacia | Function | _ Optional ByRef p_Error As String _) As String |
| RegistrarResultadoControlEficacia | Function | _ Optional ByRef p_Error As String _) As String |
| EliminarResultadoControlEficacia | Function | _ Optional ByRef p_Error As String _) As String |
| RegistrarNota | Function | _ Optional ByRef p_Error As String _) As String |
| EliminarNota | Function | _ Optional ByRef p_Error As String _) As String |
| ModificarMotivoBorrado | Function | _ Optional ByRef p_Error As String _) As String |
| ActualizarDatosCalculados | Function | _ Optional ByRef p_Error As String _) As String |

---

## NCProyecto

### Public fields

| Field | Type |
|---|---|
| IDNoConformidad | String |
| Juridica | String |
| CodigoNoConformidad | String |
| EsNoConformidad | Boolean |
| Expediente | String |
| PROYECTO | String |
| VEHICULO | String |
| Descripcion | String |
| CAUSA | String |
| CausaYAnalisRaiz | String |
| EntidadResponsable | String |
| FechaApertura | String |
| FECHACIERRE | String |
| FPREVCIERRE | String |
| Notas | String |
| Borrado | Boolean |
| RequiereACR | Boolean |
| ACR | String |
| MotivoBorrado | String |
| RequiereControlEficacia | String |
| MotivoNoRequiereControlEficacia | String |
| ControlEficacia | String |
| FechaControlEficacia | String |
| FechaPrevistaControlEficacia | String |
| ResultadoControlEficacia | String |
| ConformeControlEficacia | String |
| Cerrada | String |
| IDNCAsociada | String |
| CodigoNoConformidadAsociada | String |
| CodConcesionAsociada | String |
| IDExpediente | String |
| CodExp | String |
| Nemotecnico | String |
| JuridicaExp | String |
| IDTipo | String |
| IDProyecto | String |
| DetectadoPor | String |
| Error | String |

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| Riesgos | Get | ) As Scripting |
| Riesgos | Set | p_Valor As Scripting.Dictionary) |
| CodRiesgosAsociados | Get | ) As String |
| CodRiesgosAsociados | Let | ByVal vNewValue As String) |
| ACsSinAR | Get | ) As Scripting |
| AlgunaACSinAR | Get | ) As EnumSino |
| VinculadoANC | Get | ) As EnumSino |
| NCProyectoAsociada | Get | ) As NCProyecto |
| NCProyectoAsociada | Set | ByVal objNewValue As Variant) |
| FECHACIERRECalculada | Get | ) As String |
| FPREVCIERRECalculada | Get | ) As String |
| CerradaCalculada | Get | ) As EnumSino |
| TipoNCProyecto | Get | ) As TipologiaNCProyectos |
| TipoNCProyecto | Set | ByVal objNewValue As Variant) |
| NAccionCalculado | Get | ) As String |
| TieneAccionesPorReplanificar | Get | ) As EnumSino |
| EstadoCalculado | Get | ) As EnumEstadoNC |
| EstadoCalculadoTexto | Get | ) As String |
| EstadoEnum | Get | ) As EnumEstadoNC |
| EstadoTitulo | Get | ) As String |
| Replanificaciones | Get | ) As Scripting |
| Replanificaciones | Set | p_Valor As Variant) |
| ResponsableCalidadCalculado | Get | ) As String |
| JuridicaCalculada | Get | ) As String |
| CodExpCalculado | Get | ) As String |
| NemotecnicoCalculado | Get | ) As String |
| ExpedienteCalculadoTexto | Get | ) As String |
| ExpedienteObj | Get | ) As Expediente |
| ExpedienteObj | Set | p_Valor As Variant) |
| EsNoConformidadCalculado | Get | ) As EnumSino |
| BorradoCalculado | Get | ) As EnumSino |
| RequiereACRCalculado | Get | ) As EnumSino |
| RequiereControlEficaciaCalculado | Get | ) As EnumSino |
| ConControlDeEficaciaSinRellenar | Get | ) As EnumSino |
| EficaciaOK | Get | ) As EnumSino |
| AccionesOK | Get | ) As EnumSino |
| TieneAcciones | Get | ) As EnumSino |
| CodigoNoConformidadCalculado | Get | ) As String |
| IDNoConformidadCalculado | Get | ) As String |
| ProyectoCalculado | Get | ) As String |
| ResponsableTelefonicaObj | Get | ) As usuario |
| ResponsableTelefonicaObj | Set | ByVal objNewValue As Variant) |
| ResponsableCalidadObj | Get | ) As usuario |
| ResponsableCalidadObj | Set | ByVal objNewValue As Variant) |
| Documentos | Get | ) As Scripting |
| Documentos | Set | ByVal objNewValue As Variant) |
| DocumentosCompletos | Get | ) As Scripting |
| DocumentosCompletos | Set | ByVal objNewValue As Variant) |
| ARsSinFinalizar | Get | ) As Scripting |
| ARsSinFinalizar | Set | ByVal objNewValue As Variant) |
| ACs | Get | ) As Scripting |
| ACs | Set | ByVal objNewValue As Variant) |
| TodasLasArsFinalizadas | Get | ) As EnumSino |
| TodasLasACsSinFechas | Get | ) As EnumSino |
| ColCamposVarianEnVinculacionNC | Get | ) As Collection |
| ColCampos | Get | ) As Collection |
| ColCamposParaCopiarDeVinculada | Get | ) As Collection |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| EstadoGrabar | Function | _ Optional p_Estado As String, _ Optional ByRef p_Error As String _) As String |
| DatosGeneralesOK | Function | _ Optional p_MenosCef As EnumSino = EnumSino.No, _ Optional ByRef p_Error As String _) As EnumSino |
| FECHACIERREGrabar | Function | _ Optional p_Fecha As String, _ Optional ByRef p_Error As String _) As String |
| FPREVCIERREGrabar | Function | _ Optional p_Fecha As String, _ Optional ByRef p_Error As String _) As String |
| CIERREGrabar | Function | _ Optional p_Cerrada As EnumSino, _ Optional ByRef p_Error As String _) As String |
| getPropiedad | Function | m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant |
| SetPropiedad | Function | m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String |

---

## NCProyectoDetailVM

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| IDNoConformidad | Get | ) As Long |
| CodigoNoConformidad | Get | ) As String |
| Estado | Get | ) As String |
| Descripcion | Get | ) As String |
| CAUSA | Get | ) As String |
| ResponsableTelefonica | Get | ) As String |
| RESPONSABLECALIDAD | Get | ) As String |
| Proyecto | Get | ) As String |
| VEHICULO | Get | ) As String |
| Expediente | Get | ) As String |
| FechaApertura | Get | ) As Date |
| FECHACIERRE | Get | ) As Date |
| FechaPrevCierre | Get | ) As Date |
| Juridica | Get | ) As String |
| Tipo | Get | ) As String |
| Notas | Get | ) As String |
| Cerrada | Get | ) As String |
| RequiereACR | Get | ) As Boolean |
| ACR | Get | ) As String |
| RequiereControlEficacia | Get | ) As String |
| ControlEficacia | Get | ) As String |
| FechaControlEficacia | Get | ) As Date |
| ConformeControlEficacia | Get | ) As String |
| IDTipo | Get | ) As Long |
| Tipologia | Get | ) As String |
| IDExpediente | Get | ) As Long |
| CodExp | Get | ) As String |
| Nemotecnico | Get | ) As String |
| CodigoRiesgo | Get | ) As String |
| DetectadoPor | Get | ) As String |
| ResponsableEjecucion | Get | ) As String |
| IDNCAsociada | Get | ) As Long |
| CodigoNoConformidadAsociada | Get | ) As String |
| colARs | Get | ) As Collection |
| colACs | Get | ) As Collection |
| ColDocumentos | Get | ) As Collection |
| ColReplanificaciones | Get | ) As Collection |
| EstaCargado | Get | ) As Boolean |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| CargarPorID | Function | _ ByVal p_IDNoConformidad As Long, _ Optional ByRef p_Error As String) As Boolean |

---

## NCProyectoListItemVM

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| IDNoConformidad | Get | ) As Long |
| IDNoConformidad | Let | ByVal p_Value As Long) |
| CodigoNoConformidad | Get | ) As String |
| CodigoNoConformidad | Let | ByVal p_Value As String) |
| IDExpediente | Get | ) As Long |
| IDExpediente | Let | ByVal p_Value As Long) |
| Descripcion | Get | ) As String |
| Descripcion | Let | ByVal p_Value As String) |
| Expediente | Get | ) As String |
| Estado | Get | ) As String |
| Estado | Let | ByVal p_Value As String) |
| FechaApertura | Get | ) As Date |
| FechaApertura | Let | ByVal p_Value As Date) |
| FECHACIERRE | Get | ) As Date |
| FECHACIERRE | Let | ByVal p_Value As Date) |
| Proyecto | Get | ) As String |
| VEHICULO | Get | ) As String |
| ResponsableTelefonica | Get | ) As String |
| RESPONSABLECALIDAD | Get | ) As String |
| Cerrada | Get | ) As String |
| Cerrada | Let | ByVal p_Value As String) |
| RequiereACR | Get | ) As Boolean |
| ACR | Get | ) As String |
| RequiereControlEficacia | Get | ) As String |
| Nemotecnico | Get | ) As String |
| Nemotecnico | Let | ByVal p_Value As String) |
| CodExp | Get | ) As String |
| CodExp | Let | ByVal p_Value As String) |
| ExpedienteCalculadoTexto | Get | ) As String |
| EstaCargado | Get | ) As Boolean |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| CargarPorID | Function | _ ByVal p_IDNoConformidad As Long, _ Optional ByRef p_Error As String) As Boolean |
| CargarDesdeRecordset | Function | _ ByRef p_rs As Dao.Recordset, _ Optional ByRef p_Error As String) As Boolean |
| CargarDesdeNCProyecto | Function | _ ByVal p_NC As NCProyecto, _ Optional ByRef p_Error As String) As Boolean |

---

## NCProyectoOperaciones

### Public fields

| Field | Type |
|---|---|
| nc | ncProyecto |
| Error | String |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| MotivoAltaDatosUnicosNoOK | Function | _ Optional ByRef p_ObjNCAlInicio As ncProyecto, _ Optional ByRef p_Error As String, _ Optional ByVal p_MenosCef As EnumSino = EnumSino.No _) As String |
| MotivoDatosUnicosNoOK | Function | _ Optional ByRef p_ObjNCAlInicio As ncProyecto, _ Optional ByRef p_Error As String, _ Optional ByVal p_MenosCef As EnumSino = EnumSino.No _) As String |
| RegistrarDatosUnicos | Function | _ Optional ByRef p_ObjNCAlInicio As ncProyecto, _ Optional ByRef p_Error As String, _ Optional ByVal p_MenosCef As EnumSino = EnumSino.No, _ Optional ByRef p_PromptResult As Long = -1, _ Optional ByRef p_MessageText As String _) As String |
| RegistrarAltaDatosUnicosConVinculoNC | Function | _ Optional ByRef p_Error As String, _ Optional ByVal p_MenosCef As EnumSino = EnumSino.No _) As String |
| RegistrarCambiosDatosUnicosConVinculoNC | Function | _ ByRef p_ObjNCAlInicio As ncProyecto, _ Optional ByRef p_Error As String _) As String |
| Eliminar | Function | _ Optional ByVal p_SinDejarRastro As EnumSino = EnumSino.No, _ Optional ByRef p_Error As String _) As String |
| Habilitar | Function | _ Optional ByRef p_Error As String _) As String |
| RegistrarControlEficacia | Function | _ Optional ByRef p_Error As String _) As String |
| EliminarControlEficacia | Function | _ Optional ByRef p_Error As String _) As String |
| RegistrarResultadoControlEficacia | Function | _ Optional ByRef p_Error As String _) As String |
| EliminarResultadoControlEficacia | Function | _ Optional ByRef p_Error As String _) As String |
| RegistrarNota | Function | _ Optional ByRef p_Error As String _) As String |
| EliminarNota | Function | _ Optional ByRef p_Error As String _) As String |
| BorrarTodoLoNecesarioEnVinculados | Function | _ Optional ByRef p_Error As String _) As String |
| ModificarMotivoBorrado | Function | _ Optional ByRef p_Error As String _) As String |
| ActualizarDatosCalculados | Function | _ Optional ByRef p_Error As String _) As String |
| ActualizarRiesgosNC | Sub | p_IDNC As Long, p_ColIDs As Collection, ByRef p_Error As String) |
| MarcarListadoStalePorAccion | Sub | ByVal p_IDNC As Long, Optional ByRef p_Error As String) |

---

## ReplanificacionesAuditoria

### Public fields

| Field | Type |
|---|---|
| IDReplanificacion | String |
| IDNoConformidad | String |
| IDAccionRealizada | String |
| FechaReprogramacion | String |
| FechaPrevistaAlInicio | String |
| FechaPrevistaReplanificada | String |
| Observaciones | String |
| Error | String |

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| IDReplanificacionCalculada | Get | ) As String |
| nc | Get | ) As NCAuditoria |
| AR | Get | ) As ARAuditoria |
| AC | Get | ) As ACAuditoria |
| ColCampos | Get | ) As Collection |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| getPropiedad | Function | m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant |
| SetPropiedad | Function | m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String |

---

## ReplanificacionesAuditoriaOperaciones

### Public fields

| Field | Type |
|---|---|
| ReplanificacionesAuditoria | ReplanificacionesAuditoria |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| Eliminar | Function | _ Optional ByRef p_Error As String _) As String |
| MotivoNoOK | Function | _ Optional ByRef p_Error As String _) As String |
| Registrar | Function | _ Optional ByRef p_Error As String _) As String |

---

## ReplanificacionesProyecto

### Public fields

| Field | Type |
|---|---|
| IDReplanificacion | String |
| IDNoConformidad | String |
| IDAccionRealizada | String |
| FechaReprogramacion | String |
| FechaPrevistaAlInicio | String |
| FechaPrevistaReplanificada | String |
| Observaciones | String |
| Error | String |

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| IDReplanificacionCalculada | Get | ) As String |
| nc | Get | ) As NCProyecto |
| AR | Get | ) As ARProyecto |
| AC | Get | ) As ACProyecto |
| ColCampos | Get | ) As Collection |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| getPropiedad | Function | m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant |
| SetPropiedad | Function | m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String |

---

## ReplanificacionesProyectoOperaciones

### Public fields

| Field | Type |
|---|---|
| ReplanificacionesProyecto | ReplanificacionesProyecto |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| Eliminar | Function | _ Optional ByRef p_Error As String _) As String |
| MotivoNoOK | Function | _ Optional ByRef p_Error As String _) As String |
| Registrar | Function | _ Optional ByRef p_Error As String _) As String |

---

## ReplanificacionesService

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| Registrar | Function | ByRef p_Replanificacion As ReplanificacionesProyecto, Optional ByRef p_Error As String) As Boolean |
| Eliminar | Function | ByRef p_Replanificacion As ReplanificacionesProyecto, Optional ByRef p_Error As String) As Boolean |

---

## Riesgo

### Public fields

| Field | Type |
|---|---|
| idRiesgo | String |
| IDEdicion | String |
| CodigoUnico | String |
| CodigoRiesgo | String |
| FechaDetectado | String |
| DetectadoPor | String |
| EntidadDetecta | String |
| Plazo | String |
| Calidad | String |
| Coste | String |
| ImpactoGlobal | String |
| Vulnerabilidad | String |
| Valoracion | String |
| Mitigacion | String |
| Contingencia | String |
| RequierePlanContingencia | String |
| Descripcion | String |
| CAUSARAIZ | String |
| Estado | String |
| FechaEstado | String |
| Priorizacion | String |
| FechaMaterializado | String |
| FechaRetirado | String |
| FechaCerrado | String |
| FechaMitigacionAceptar | String |
| JustificacionAceptacionRiesgo | String |
| FechaJustificacionAceptacionRiesgo | String |
| FechaAprobacionAceptacionPorCalidad | String |
| FechaRechazoAceptacionPorCalidad | String |
| JustificacionRetiroRiesgo | String |
| FechaJustificacionRetiroRiesgo | String |
| FechaAprobacionRetiroPorCalidad | String |
| FechaRechazoRetiroPorCalidad | String |
| RequiereRiesgoDeBiblioteca | String |
| CodRiesgoBiblioteca | String |
| RiesgoPendienteRetipificacion | String |
| FechaRiesgoParaRetipificar | String |
| FechaRiesgoRetipificado | String |
| DiasSinRespuestaCalidadAceptacion | String |
| DiasSinRespuestaCalidadRetiro | String |
| DiasSinRespuestaCalidadRetipificacion | String |
| Origen | String |
| Error | String |

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| ColCampos | Get | ) As Collection |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| getPropiedad | Function | m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant |
| SetPropiedad | Function | m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String |

---

## RiesgoServicio

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| ListarRiesgosCandidatos | Function | ByRef p_NC As NCProyecto, ByRef p_Error As String) As Scripting |
| ActualizarRiesgosNC | Function | ByVal p_IDNC As Long, ByRef p_ColRiesgosSeleccionados As Collection, ByRef p_Error As String) As Boolean |
| SincronizarCampoTextoNC | Sub | ByVal p_IDNC As Long, ByRef p_Error As String) |

---

## SegNCAuditoria

### Public fields

| Field | Type |
|---|---|
| id | String |
| IDAuditoria | String |
| Auditoria | String |
| Descripcion | String |
| FECHACIERRE | String |
| Numero | String |
| Error | String |

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| nc | Get | ) As NCAuditoria |
| ColCampos | Get | ) As Collection |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| getPropiedad | Function | m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant |
| SetPropiedad | Function | m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String |

---

## SegNCProyecto

### Public fields

| Field | Type |
|---|---|
| IDNoConformidad | String |
| CodigoNoConformidad | String |
| Descripcion | String |
| Nemotecnico | String |
| IDExpediente | String |
| RequiereControlEficacia | String |
| ResultadoControlEficacia | String |
| FECHACIERRE | String |
| Error | String |

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| nc | Get | ) As NCProyecto |
| ColCampos | Get | ) As Collection |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| getPropiedad | Function | m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant |
| SetPropiedad | Function | m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String |

---

## SegTareasAuditoria

### Public fields

| Field | Type |
|---|---|
| IDAccionRealizada | String |
| IdAccionCorrectiva | String |
| id | String |
| NAR | String |
| FechaInicio | String |
| FechaFinPrevista | String |
| FechaFinReal | String |
| TipoNC | String |
| Responsable | String |
| NAccion | String |
| Auditoria | String |
| Error | String |

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| AR | Get | ) As ARAuditoria |
| AC | Get | ) As ACAuditoria |
| nc | Get | ) As NCAuditoria |
| ColCampos | Get | ) As Collection |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| getPropiedad | Function | m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant |
| SetPropiedad | Function | m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String |

---

## SegTareasProyecto

### Public fields

| Field | Type |
|---|---|
| IDAccionRealizada | String |
| IdAccionCorrectiva | String |
| IDNoConformidad | String |
| FechaInicio | String |
| FechaFinPrevista | String |
| FechaFinReal | String |
| TipoNC | String |
| RespCalidad | String |
| IDExpediente | String |
| NAccion | String |
| Error | String |

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| AR | Get | ) As ARProyecto |
| AC | Get | ) As ACProyecto |
| nc | Get | ) As NCProyecto |
| ColCampos | Get | ) As Collection |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| getPropiedad | Function | m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant |
| SetPropiedad | Function | m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String |

---

## TipologiaNCProyectos

### Public fields

| Field | Type |
|---|---|
| IDTipo | String |
| Tipologia | String |
| Error | String |

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| IDTipoCalculado | Get | ) As String |
| ColCampos | Get | ) As Collection |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| Registrar | Function | _ Optional ByRef p_TipologiaAlInicio As TipologiaNCProyectos, _ Optional ByRef p_Error As String _) As String |
| Eliminar | Function | _ Optional ByRef p_Error As String _) As String |
| getPropiedad | Function | m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant |
| SetPropiedad | Function | m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String |

---

## Usuario

### Public fields

| Field | Type |
|---|---|
| CorreoUsuario | String |
| Password | String |
| UsuarioRed | String |
| Nombre | String |
| Matricula | String |
| FechaAlta | String |
| Activado | Boolean |
| FechaProximoCambioContrasenia | String |
| FechaUltimaConexion | String |
| TieneQueCambiarLaPass | Boolean |
| Telefono | String |
| Movil | String |
| Observaciones | String |
| UsuarioImborrable | Boolean |
| EsAdministrador | String |
| PermisosAsignados | Boolean |
| FechaBaja | String |
| PasswordNuncaCaduca | Boolean |
| MantenerLanzaderaAbierta | Boolean |
| PassIncialPlana | String |
| UsuarioSSID | String |
| id | String |
| JefeDelUsuario | String |
| PermisoPruebas | String |
| ParaTareasProgramadas | Boolean |
| FechaBloqueo | String |
| Error | String |

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| EsAdministradorCalculado | Get | ) As EnumSino |
| EsUsuarioTecnicoCalculado | Get | ) As EnumSino |
| EsUsuarioCalidadCalculado | Get | ) As EnumSino |
| EsUsuarioEconomiaCalculado | Get | ) As EnumSino |
| EsUsuarioSecretariaCalculado | Get | ) As EnumSino |
| EsUsuarioCalidadAvisosCalculado | Get | ) As EnumSino |
| ColAplicacionesPermisos | Get | ) As Scripting |
| Permisos | Get | ) As UsuarioAplicacionPermisos |
| ColCampos | Get | ) As Collection |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| getPropiedad | Function | m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant |
| SetPropiedad | Function | m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String |

---

## UsuarioAplicacionPermisos

### Public fields

| Field | Type |
|---|---|
| CorreoUsuario | String |
| IDAplicacion | String |
| EsUsuarioAdministrador | String |
| EsUsuarioCalidad | String |
| EsUsuarioEconomia | String |
| EsUsuarioSecretaria | String |
| EsUsuarioTecnico | String |
| EsUsuarioSinAcceso | String |
| EsUsuarioCalidadAvisos | String |
| Error | String |

### Public Properties

| Property | Kind | Signature |
|---|---|---|
| usuario | Get | ) As usuario |
| usuario | Set | ByVal objNewValue As usuario) |
| EsUsuarioAdministradorCalculado | Get | ) As EnumSino |
| EsUsuarioAdministradorCalculado | Let | ByVal eNewValue As EnumSino) |
| EsUsuarioCalidadCalculado | Get | ) As EnumSino |
| EsUsuarioCalidadCalculado | Let | ByVal eNewValue As EnumSino) |
| EsUsuarioEconomiaCalculado | Get | ) As EnumSino |
| EsUsuarioEconomiaCalculado | Let | ByVal eNewValue As EnumSino) |
| EsUsuarioSecretariaCalculado | Get | ) As EnumSino |
| EsUsuarioSecretariaCalculado | Let | ByVal eNewValue As EnumSino) |
| EsUsuarioTecnicoCalculado | Get | ) As EnumSino |
| EsUsuarioTecnicoCalculado | Let | ByVal eNewValue As EnumSino) |
| EsUsuarioSinAccesoCalculado | Get | ) As EnumSino |
| EsUsuarioSinAccesoCalculado | Let | ByVal eNewValue As EnumSino) |
| EsUsuarioCalidadAvisosCalculado | Get | ) As EnumSino |
| EsUsuarioCalidadAvisosCalculado | Let | ByVal eNewValue As EnumSino) |
| ColCampos | Get | ) As Collection |

### Public Functions / Subs

| Member | Kind | Signature |
|---|---|---|
| getPropiedad | Function | m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant |
| SetPropiedad | Function | m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String |

