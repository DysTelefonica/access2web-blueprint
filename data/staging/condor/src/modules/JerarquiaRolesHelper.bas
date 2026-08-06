Attribute VB_Name = "JerarquiaRolesHelper"
Option Compare Database
Option Explicit

' ==========================================================================
' MODULE: JerarquiaRolesHelper
' PURPOSE: Pure helper that encodes the BR-010 role hierarchy
'          (Admin > Calidad > Tecnico + Ingenieria / Economia / Secretaria /
'          SinAcceso / CalidadAvisos) without touching DAO. Centralizes the
'          table that WorkflowServicio.PermisoSuficiente duplicated inline.
'
' Contract documented in:
'   changes/e2e-methodology-exhaustive-rollout/specs/workflow-precondicion-transicion/spec.md
'
' Constraints (per spec):
'   * Pure helper: NO DAO.Database, NO control access, NO MsgBox.
'   * Internal hash (Select Case) for role resolution.
'   * String inputs (estadoOrigen / estadoDestino) - canonical names like
'     "estadoDesarrolloTecnico", "estadoModificacion", "estadoValidacion",
'     "estadoRevision", "estadoFormalizacion", "estadoAprobada",
'     "estadoRechazada", "estadoPreregistro", "estadoRegistro".
'   * Trim + UCase$ applied to all inputs (defensive).
'   * Returns False for unknown rol or unknown estado destino (safe-by-default).
' ==========================================================================

' Allowed canonical estado destino names. Used to validate p_EstadoDestino
' before any role-specific resolution. Adding a new canonical estado here
' is the only place that must change to keep the helper aligned with
' modEnumeradores.bas.
Private Const JERARQUIA_ESTADOS_VALIDOS As String = _
    "estadoPreregistro|estadoRegistro|estadoDesarrolloTecnico|" & _
    "estadoModificacion|estadoValidacion|estadoRevision|" & _
    "estadoFormalizacion|estadoAprobada|estadoRechazada"

' Etapas (origen) that a Tecnico is allowed to drive transitions from.
' Mirrors CAP-007 §2 transiciones canonicas: estadoDesarrolloTecnico (3)
' and estadoModificacion (4) are the editable workbench.
Private Const JERARQUIA_TECNICO_ORIGENES As String = _
    "estadoDesarrolloTecnico|estadoModificacion"

' Destinations a Tecnico is allowed to transition to (subsanacion loop:
' estadoDesarrolloTecnico <-> estadoModificacion).
Private Const JERARQUIA_TECNICO_DESTINOS As String = _
    "estadoDesarrolloTecnico|estadoModificacion"

' Public seam: returns True iff p_RolUsuario is allowed to drive the
' (origen -> destino) transition under BR-010.
'
' Inputs are Variant to absorb DAO Null without raising 94; trimming and
' UCase$ make the helper case- and whitespace-insensitive (atom #8).
Public Function JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde( _
    ByVal p_RolUsuario As Variant, _
    ByVal p_EstadoOrigen As Variant, _
    ByVal p_EstadoDestino As Variant) As Boolean

    Dim rol As String
    Dim origen As String
    Dim destino As String

    JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde = False

    rol = UCase$(Trim$(Nz(p_RolUsuario, vbNullString)))
    origen = UCase$(Trim$(Nz(p_EstadoOrigen, vbNullString)))
    destino = UCase$(Trim$(Nz(p_EstadoDestino, vbNullString)))

    ' Safe-by-default: unknown rol never transitions.
    If Len(rol) = 0 Then Exit Function

    ' Safe-by-default: unknown destino never transitions.
    If Len(destino) = 0 Then Exit Function
    If Not JerarquiaRolesHelper_EstadoEsValido(destino) Then Exit Function

    Select Case rol
        Case "ADMIN", "ADMINISTRADOR"
            ' BR-010: Admin > all.
            JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde = True

        Case "CALIDAD"
            ' BR-010: Calidad > Validacion / Revision / Rechazada (rechazo
            ' desde formalizacion). Tambien puede devolver la solicitud a
            ' estadoDesarrolloTecnico o estadoModificacion cuando la
            ' subsanacion lo requiere.
            JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde = _
                JerarquiaRolesHelper_CalidadPuedeTransicionar(origen, destino)

        Case "TECNICO", "INGENIERIA"
            ' BR-010: Tecnico/Ingenieria subsanan. Solo desde estados
            ' editables y solo a destinos del loop de subsanacion.
            ' (Per spec escenario, Tecnico NO transiciona a estadoValidacion.)
            JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde = _
                JerarquiaRolesHelper_TecnicoPuedeTransicionar(origen, destino)

        Case "ECONOMIA", "SECRETARIA"
            ' Roles secundarios: sin capacidad de transicionar.
            JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde = False

        Case "CALIDADAVISOS"
            ' Subset mas restrictivo de Calidad: solo a estadoValidacion.
            JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde = _
                (destino = "ESTADOVALIDACION")

        Case "SINACCESO"
            ' SinAcceso nunca transiciona.
            JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde = False

        Case Else
            ' Unknown rol -> False (safe-by-default).
            JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde = False
    End Select
End Function

' ----------------------------------------------------------------------------
' Private role-specific tables
' ----------------------------------------------------------------------------

' Calidad puede transicionar hacia estados propios del perfil:
'   * estadoValidacion       (5) -> aceptar subsanacion
'   * estadoRevision         (6) -> enviar a RAC
'   * estadoRechazada        (9) -> cerrar como rechazada desde formalizacion
'   * estadoDesarrolloTecnico(3) -> devolver para subsanacion
'   * estadoModificacion     (4) -> devolver para subsanacion
Private Function JerarquiaRolesHelper_CalidadPuedeTransicionar( _
    ByVal p_Origen As String, _
    ByVal p_Destino As String) As Boolean

    Select Case p_Destino
        Case "ESTADOVALIDACION", "ESTADOREVISION", "ESTADOREChAZADA", _
             "ESTADODESARROLLOTECNICO", "ESTADOMODIFICACION"
            JerarquiaRolesHelper_CalidadPuedeTransicionar = True
        Case Else
            JerarquiaRolesHelper_CalidadPuedeTransicionar = False
    End Select
End Function

' Tecnico / Ingenieria: solo el loop de subsanacion
'   origen in {estadoDesarrolloTecnico, estadoModificacion}
'   destino in {estadoDesarrolloTecnico, estadoModificacion}
' Cualquier otro par -> False (BR-010 + spec escenario TecnicoNoPuedePasarAValidacion).
Private Function JerarquiaRolesHelper_TecnicoPuedeTransicionar( _
    ByVal p_Origen As String, _
    ByVal p_Destino As String) As Boolean

    Dim origenOK As Boolean
    Dim destinoOK As Boolean

    origenOK = (InStr(1, JERARQUIA_TECNICO_ORIGENES, p_Origen, vbTextCompare) > 0)
    destinoOK = (InStr(1, JERARQUIA_TECNICO_DESTINOS, p_Destino, vbTextCompare) > 0)

    JerarquiaRolesHelper_TecnicoPuedeTransicionar = (origenOK And destinoOK)
End Function

' Validates p_EstadoDestino against the canonical list. Case-insensitive.
Private Function JerarquiaRolesHelper_EstadoEsValido(ByVal p_Estado As String) As Boolean
    JerarquiaRolesHelper_EstadoEsValido = _
        (InStr(1, JERARQUIA_ESTADOS_VALIDOS, p_Estado, vbTextCompare) > 0)
End Function
