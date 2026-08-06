Attribute VB_Name = "modRiesgoEstablecerDatosHelper"
Option Compare Database
Option Explicit

' =============================================================================
' modRiesgoEstablecerDatosHelper.bas
'
' Helper que extrae la logica de negocio de:
'   - Form_FormRiesgo.EstablecerDatos                  (padre)
'   - Form_FormRiesgosGestionRiesgo.EstablecerDatos    (gemelo)
'   - Form_FormRiesgosGestion.EstablecerDatos          (parent LIST form)
'
' Slice HR4 (06120fc): padres + gemelo.
' Slice HR4b (este slice): parent LIST form (FormRiesgosGestion).
'
' Hard rule 1: cero business logic en event handlers o Public methods en forms.
' Hard rule 2: honest signature — solo inyecta lo que la funcion necesita.
' Hard rule 4: Public business methods en forms son anti-patron (EstablecerDatos
'               debe ser Private despues de extraer).
' Hard rule 5: nombres Public globalmente unicos con ModuleName_ prefix.
' Hard rule 7: NO MsgBox/InputBox en helpers. Errores via ByRef p_Error.
'
' Capa:
'   - RiesgoEstablecerDatos_Riesgo(p_Error)               — proxy entry (FormRiesgo)
'   - RiesgoEstablecerDatos_RiesgoGestionRiesgo(p_Error) — proxy entry (gemelo)
'   - RiesgoEstablecerDatos_RiesgosGestion(p_Error)       — proxy entry (parent LIST)
'   - RiesgoEstablecerDatos_Calcular(...)                 — pure compute (atoms, padre/gemelo)
'   - RiesgoEstablecerDatos_CalcularRiesgosGestion(...)   — pure compute (atoms, LIST)
' =============================================================================

' --- Public Type: output del compute (padre/gemelo) ---
Public Type RiesgoEstablecerDatos_Resultado
    PermitidoEditar As EnumSiNo
    PermitidoAlta As Boolean
    Titulo As String
    Particula As String
    NavGeneralTarget As String
    NavRetiradoEnabled As Boolean
    NavMaterializadoEnabled As Boolean
    LblEstadoRiesgoVisible As Boolean
    blnEdicionActiva As Boolean
End Type

' --- Public Type: output del compute (parent LIST form) ---
'   Distinto del padre/gemelo porque opera sobre el proyecto, no sobre
'   un riesgo activo. Solo captura los campos puros que el form
'   renderiza a continuacion en controles (sin tocar DAO).
Public Type RiesgoEstablecerDatos_RiesgosGestion_Resultado
    PermitidoEditar As Boolean
    Caption As String
    ComboVerRetirados As String
    ComboVerDescripcion As String
    EsEdicionActiva As EnumSiNo
End Type

' --- Constants del modulo ---
Private Const NAV_TARGET_CON_BIBLIOTECA As String = "FormRiesgoDefinicion"
Private Const NAV_TARGET_SIN_BIBLIOTECA As String = "FormRiesgoDefinicionNoBiblioteca"

' =============================================================================
' Public Sub RiesgoEstablecerDatos_Riesgo
'   Proxy entry point para modFormCoordinationHelper.Coord_RefreshRiesgo.
'   Hard rule 4 OK: delega al Public Sub Refrescar del form (UI adapter tiny).
' =============================================================================
Public Sub RiesgoEstablecerDatos_Riesgo(Optional ByRef p_Error As String)
    On Error GoTo errores
    p_Error = ""

    If Not FormularioAbierto("FormRiesgo") Then
        ' Estado normal (form no abierto): no-op sin raise.
        Exit Sub
    End If

    Form_FormRiesgo.Refrescar p_Error
    Exit Sub

errores:
    If p_Error = "" Then
        p_Error = "RiesgoEstablecerDatos_Riesgo: " & Err.Description
    End If
End Sub

' =============================================================================
' Public Sub RiesgoEstablecerDatos_RiesgoGestionRiesgo
'   Proxy entry point para modFormCoordinationHelper.Coord_RefreshRiesgoGestionRiesgo.
' =============================================================================
Public Sub RiesgoEstablecerDatos_RiesgoGestionRiesgo(Optional ByRef p_Error As String)
    On Error GoTo errores
    p_Error = ""

    If Not FormularioAbierto("FormRiesgosGestionRiesgo") Then
        Exit Sub
    End If

    Form_FormRiesgosGestionRiesgo.Refrescar p_Error
    Exit Sub

errores:
    If p_Error = "" Then
        p_Error = "RiesgoEstablecerDatos_RiesgoGestionRiesgo: " & Err.Description
    End If
End Sub

' =============================================================================
' Public Sub RiesgoEstablecerDatos_RiesgosGestion
'   Proxy entry point para modFormCoordinationHelper.Coord_RefreshRiesgosGestion.
'   Hard rule 4 OK: delega al Public Sub Refrescar del form (UI adapter tiny).
'
'   Slice HR4b: añade el 3er proxy para el parent LIST form
'   (Form_FormRiesgosGestion) — antes la logica de negocio de su
'   Public Function EstablecerDatos vivia inline.
' =============================================================================
Public Sub RiesgoEstablecerDatos_RiesgosGestion(Optional ByRef p_Error As String)
    On Error GoTo errores
    p_Error = ""

    If Not FormularioAbierto("FormRiesgosGestion") Then
        ' Estado normal (form no abierto): no-op sin raise.
        Exit Sub
    End If

    Form_FormRiesgosGestion.Refrescar p_Error
    Exit Sub

errores:
    If p_Error = "" Then
        p_Error = "RiesgoEstablecerDatos_RiesgosGestion: " & Err.Description
    End If
End Sub

' =============================================================================
' Public Sub RiesgoEstablecerDatos_Calcular — pure compute (atoms test this)
'   Side effects: ninguno (no toca Forms()). Pure compute + carga riesgoAlInicio.
'   Raises: NO. Errores via p_Error ByRef.
'
'   p_EsGemelo: True -> usa CalcularPermitidoEditarGemelo (solo Edicion.EsActivo).
'                False -> usa CalcularPermitidoEditar (EsActivo AND Proyecto.UsuarioAutorizado).
' =============================================================================
Public Sub RiesgoEstablecerDatos_Calcular( _
    ByVal p_ObjRiesgoActivo As Object, _
    ByRef p_ObjEdicionActiva As Object, _
    ByVal p_EsAlta As EnumSiNo, _
    ByVal p_PermitidoEditarActual As Variant, _
    ByVal p_EsGemelo As Boolean, _
    ByRef p_Resultado As RiesgoEstablecerDatos_Resultado, _
    ByRef p_ObjRiesgoAlInicio As Object, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_Error As String)

    On Error GoTo errores

    p_Error = ""
    Set p_ObjRiesgoAlInicio = Nothing
    InicializarResultado p_Resultado

    ' --- Edge case: m_ObjRiesgoActivo Is Nothing ---
    If p_ObjRiesgoActivo Is Nothing Then
        p_Error = "RiesgoEstablecerDatos_Calcular: no hay riesgo activo"
        Exit Sub
    End If

    ' --- Sad path: m_EsAlta=Si pero PermitidoEditar=No ---
    If p_EsAlta = EnumSiNo.Sí Then
        If Nz2(p_PermitidoEditarActual) = CInt(EnumSiNo.No) Then
            p_Error = "Es un alta y el usuario no tiene permitida esa operacion"
            Exit Sub
        End If
        p_Resultado.PermitidoAlta = True
    End If

    ' --- Resolver edicion si no vino seteada ---
    If p_ObjEdicionActiva Is Nothing Then
        Set p_ObjEdicionActiva = p_ObjRiesgoActivo.Edicion
        If p_ObjEdicionActiva Is Nothing Then
            Dim m_IDEdicion As String
            m_IDEdicion = CStr(p_ObjRiesgoActivo.IDEdicion)
            If Len(m_IDEdicion) > 0 Then
                Set p_ObjEdicionActiva = GetCachedEdicion(m_IDEdicion, p_Error)
            End If
            If p_ObjEdicionActiva Is Nothing Then
                If p_Error = "" Then
                    p_Error = "No se puede determinar la edicion activa"
                End If
                Exit Sub
            End If
        End If
    End If

    ' --- PermitidoEditar (computar si viene Empty) ---
    If IsEmpty(p_PermitidoEditarActual) Then
        If p_EsGemelo Then
            p_Resultado.PermitidoEditar = CalcularPermitidoEditarGemelo(p_ObjEdicionActiva)
        Else
            p_Resultado.PermitidoEditar = CalcularPermitidoEditar(p_ObjEdicionActiva)
        End If
    Else
        p_Resultado.PermitidoEditar = CInt(Nz2(p_PermitidoEditarActual))
    End If

    ' --- Carga riesgoAlInicio (solo si NO es alta) ---
    If p_EsAlta <> EnumSiNo.Sí Then
        Dim m_IDRiesgo As String
        m_IDRiesgo = CStr(p_ObjRiesgoActivo.IDRiesgo)
        If Len(m_IDRiesgo) > 0 Then
            Set p_ObjRiesgoAlInicio = GetCachedRiesgoFresh(m_IDRiesgo, p_Error, db)
            If p_Error <> "" Then Exit Sub
            If p_ObjRiesgoAlInicio Is Nothing Then
                p_Error = "Es una edicion y no hay datos"
                Exit Sub
            End If
        End If
    End If

    ' --- EstadoRiesgo + NavRetirado/NavMaterializado ---
    Dim m_EstadoRiesgo As Long
    m_EstadoRiesgo = CLng(p_ObjRiesgoActivo.EstadoEnum)
    If m_EstadoRiesgo = CLng(EnumRiesgoEstado.Incompleto) Or m_EstadoRiesgo = 0 Then
        p_Resultado.NavRetiradoEnabled = False
        p_Resultado.NavMaterializadoEnabled = False
    Else
        p_Resultado.NavRetiradoEnabled = True
        p_Resultado.NavMaterializadoEnabled = True
    End If

    ' --- LblEstadoRiesgoVisible (False en alta, True en edicion) ---
    p_Resultado.LblEstadoRiesgoVisible = (p_EsAlta <> EnumSiNo.Sí)

    ' --- Titulo + Particula + NavGeneralTarget + blnEdicionActiva ---
    CalcularTituloParticulaNavTarget _
        p_ObjEdicionActiva, p_ObjRiesgoAlInicio, p_EsAlta, _
        p_Resultado.Titulo, p_Resultado.Particula, p_Resultado.NavGeneralTarget, _
        p_Resultado.blnEdicionActiva

    Exit Sub

errores:
    If Err.Number <> 1000 Then
        p_Error = "RiesgoEstablecerDatos_Calcular: " & Err.Number & " - " & Err.Description
    End If
End Sub

' =============================================================================
' Public Sub RiesgoEstablecerDatos_CalcularRiesgosGestion — pure compute (LIST form)
'   Slice HR4b: extrae la logica pura del parent LIST form
'   (Form_FormRiesgosGestion.EstablecerDatos). Esta es la porcion
'   deterministica / testeable: NO toca Forms() ni Me ni DAO. El form
'   sigue siendo responsable de:
'     - ComboEdicion.AddItem (iteracion sobre colEdiciones)
'     - Set Me.Arbol, Me.ListaImagenes
'     - RellenarListaImagenes / RegistrarUltimoProyecto / CargarArbol
'     - ActualizarEtiquetaUltimoCambio
'     - Enable/disable de ComandoRiesgoPriorizacion + ComandoPublicacion
'     - ComandoEditarProyecto visibilidad
'
'   Honest signature (Hard rule 2): inyecto SOLO lo que el pure compute
'   necesita: m_ObjEntorno (para TituloUsuarioConectado + flags)
'   y m_ObjEdicionActiva (resuelta). NO recibe db (Hard rule 2 — no la
'   leemos aqui, queda como side effect en el form).
'
'   Raises: NO. Errores via p_Error ByRef.
' =============================================================================
Public Sub RiesgoEstablecerDatos_CalcularRiesgosGestion( _
    ByVal p_ObjEntorno As Object, _
    ByVal p_ObjEdicionActiva As Object, _
    ByRef p_Resultado As RiesgoEstablecerDatos_RiesgosGestion_Resultado, _
    Optional ByRef p_Error As String)

    On Error GoTo errores
    p_Error = ""
    InicializarResultadoRiesgosGestion p_Resultado

    ' --- Sad path: edicion activa requerida ---
    If p_ObjEdicionActiva Is Nothing Then
        p_Error = "No hay edicion activa (padre LIST)"
        Exit Sub
    End If

    ' --- PermitidoEditar (EsActivo AND UsuarioConectadoAutorizado) ---
    If p_ObjEdicionActiva.EsActivo = EnumSiNo.Sí Then
        If p_ObjEdicionActiva.UsuarioConectadoAutorizado = EnumSiNo.Sí Then
            p_Resultado.PermitidoEditar = True
        End If
    End If

    ' --- Caption (del Entorno + edicion/proyecto del edicion) ---
    CalcularCaptionRiesgosGestion _
        p_ObjEntorno, p_ObjEdicionActiva, p_Resultado.Caption

    ' --- ComboVerRetirados / ComboVerDescripcion (defaults derivados del Entorno) ---
    If Not p_ObjEntorno Is Nothing Then
        If p_ObjEntorno.VerSoloRiesgosNoRetirados = EnumSiNo.Sí Then
            p_Resultado.ComboVerRetirados = "No"
        Else
            p_Resultado.ComboVerRetirados = "Sí"
        End If
        If p_ObjEntorno.VerRiesgosDescripcion = EnumSiNo.Sí Then
            p_Resultado.ComboVerDescripcion = "Sí"
        Else
            p_Resultado.ComboVerDescripcion = "No"
        End If
    End If

    ' --- EsEdicionActiva (lo usa el form para flag m_EsEdicionActiva) ---
    p_Resultado.EsEdicionActiva = p_ObjEdicionActiva.EsActivo

    Exit Sub

errores:
    If Err.Number <> 1000 Then
        p_Error = "RiesgoEstablecerDatos_CalcularRiesgosGestion: " & Err.Number & " - " & Err.Description
    End If
End Sub

' =============================================================================
' Helpers privados (no parte del contrato publico)
' =============================================================================

Private Sub InicializarResultado(ByRef p_Resultado As RiesgoEstablecerDatos_Resultado)
    p_Resultado.PermitidoEditar = 0
    p_Resultado.PermitidoAlta = False
    p_Resultado.Titulo = ""
    p_Resultado.Particula = ""
    p_Resultado.NavGeneralTarget = NAV_TARGET_SIN_BIBLIOTECA
    p_Resultado.NavRetiradoEnabled = False
    p_Resultado.NavMaterializadoEnabled = False
    p_Resultado.LblEstadoRiesgoVisible = False
    p_Resultado.blnEdicionActiva = False
End Sub

' Initialize RiesgosGestion-specific Resultado (parent LIST form).
' Slice HR4b.
Private Sub InicializarResultadoRiesgosGestion(ByRef p_Resultado As RiesgoEstablecerDatos_RiesgosGestion_Resultado)
    p_Resultado.PermitidoEditar = False
    p_Resultado.Caption = ""
    p_Resultado.ComboVerRetirados = "Sí"
    p_Resultado.ComboVerDescripcion = "No"
    p_Resultado.EsEdicionActiva = EnumSiNo.No
End Sub

' Compose the form caption: "<TituloUsuario> riesgos de la edicion X de Y".
' Pure compute (no UI side effects). Slice HR4b.
Private Sub CalcularCaptionRiesgosGestion( _
    ByVal p_ObjEntorno As Object, _
    ByVal p_ObjEdicionActiva As Object, _
    ByRef p_Caption As String)

    On Error GoTo errores

    p_Caption = ""
    If p_ObjEntorno Is Nothing Then Exit Sub
    If p_ObjEdicionActiva Is Nothing Then Exit Sub

    p_Caption = p_ObjEntorno.TituloUsuarioConectado & " " & _
                "riesgos de la edición " & p_ObjEdicionActiva.Edicion & _
                " de " & p_ObjEdicionActiva.Proyecto.Proyecto
    Exit Sub

errores:
    ' Silencioso: devuelve "" para que el form renderice sin caption.
End Sub

' Padre (Form_FormRiesgo): edicion.EsActivo AND proyecto.UsuarioAutorizado.
Private Function CalcularPermitidoEditar(ByVal p_ObjEdicionActiva As Object) As EnumSiNo
    On Error GoTo errores

    If p_ObjEdicionActiva Is Nothing Then
        CalcularPermitidoEditar = EnumSiNo.No
        Exit Function
    End If

    If p_ObjEdicionActiva.EsActivo = EnumSiNo.No Then
        CalcularPermitidoEditar = EnumSiNo.No
    Else
        If p_ObjEdicionActiva.Proyecto.UsuarioAutorizado = EnumSiNo.Sí Then
            CalcularPermitidoEditar = EnumSiNo.Sí
        Else
            CalcularPermitidoEditar = EnumSiNo.No
        End If
    End If
    Exit Function

errores:
    CalcularPermitidoEditar = EnumSiNo.No
End Function

' Gemelo (Form_FormRiesgosGestionRiesgo): solo edicion.EsActivo.
Private Function CalcularPermitidoEditarGemelo(ByVal p_ObjEdicionActiva As Object) As EnumSiNo
    On Error GoTo errores

    If p_ObjEdicionActiva Is Nothing Then
        CalcularPermitidoEditarGemelo = EnumSiNo.No
        Exit Function
    End If

    If p_ObjEdicionActiva.EsActivo = EnumSiNo.No Then
        CalcularPermitidoEditarGemelo = EnumSiNo.No
    Else
        CalcularPermitidoEditarGemelo = EnumSiNo.Sí
    End If
    Exit Function

errores:
    CalcularPermitidoEditarGemelo = EnumSiNo.No
End Function

' Computa Titulo, Particula, NavGeneralTarget, blnEdicionActiva.
Private Sub CalcularTituloParticulaNavTarget( _
    ByVal p_ObjEdicionActiva As Object, _
    ByVal p_ObjRiesgoAlInicio As Object, _
    ByVal p_EsAlta As EnumSiNo, _
    ByRef p_Titulo As String, _
    ByRef p_Particula As String, _
    ByRef p_NavGeneralTarget As String, _
    ByRef p_blnEdicionActiva As Boolean)

    On Error GoTo errores

    p_Titulo = ""
    p_Particula = ""
    p_NavGeneralTarget = NAV_TARGET_SIN_BIBLIOTECA
    p_blnEdicionActiva = False

    If p_ObjEdicionActiva Is Nothing Then Exit Sub

    With p_ObjEdicionActiva
        If .EsActivo = EnumSiNo.Sí Then
            p_blnEdicionActiva = True
            p_Particula = "(edicion activa)"
        Else
            p_blnEdicionActiva = False
            p_Particula = "(edicion no Activa)"
        End If
        If p_EsAlta = EnumSiNo.Sí Then
            p_Titulo = "ALTA DE RIESGO PARA LA EDICION " & .Edicion & " " & p_Particula
        Else
            If Not p_ObjRiesgoAlInicio Is Nothing Then
                p_Titulo = "DETALLE DEL RIESGO: " & p_ObjRiesgoAlInicio.CodigoRiesgo & " EDICION " & .Edicion & " " & p_Particula
            Else
                p_Titulo = "DETALLE DEL RIESGO: EDICION " & .Edicion & " " & p_Particula
            End If
        End If
        If .Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí Then
            p_NavGeneralTarget = NAV_TARGET_CON_BIBLIOTECA
        Else
            p_NavGeneralTarget = NAV_TARGET_SIN_BIBLIOTECA
        End If
    End With

    Exit Sub

errores:
    ' Silencioso: defaults sensatos.
End Sub

' Coerce numericos a Long; Empty/Nothing -> 0.
Private Function Nz2(ByVal p_Value As Variant) As Long
    On Error Resume Next
    If IsEmpty(p_Value) Or IsNull(p_Value) Then
        Nz2 = 0
    Else
        Nz2 = CLng(p_Value)
    End If
    On Error GoTo 0
End Function