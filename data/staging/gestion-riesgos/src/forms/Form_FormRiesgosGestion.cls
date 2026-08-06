Option Compare Database
Option Explicit
'````````````````````````````````````````````````````````````````````````````
'`                    Code from Tassos Filoxenidis                          `
'````````````````````````````````````````````````````````````````````````````

Private m_sngLeftResizePos As Integer
Private MeWidth As Integer
Private MeWidthHalf As Integer
Private ArbolLeft As Integer
Private SubFormLeft As Integer
Private Sepleft As Integer
Const SepWMin As Integer = 60
Const SepW As Integer = 60.1

Private WithEvents m_FormRiesgo As Form_FormRiesgo


Public m_NodoSeleccionado As MSComctlLib.Node
Private m_URLInformeRiesgo As String
Public m_ObjSeleccionado As Object
Public m_RiesgoSeleccionado As riesgo
Public m_EstadoRiesgoSeleccionado As EnumRiesgoEstado
Public m_PMSeleccionado As pm
Public m_PMAccionSeleccionado As PMAccion
Public m_PCSeleccionado As pc
Public m_PCAccionSeleccionado As PCAccion
Public m_EsMitigacion As EnumSiNo
Public m_EsEdicionActiva As EnumSiNo
Public m_NodoEdicion As MSComctlLib.Node
Public m_Arbol As MSComctlLib.TreeView
Private m_ListImages As Object
Private m_SinCargarFormulario As EnumSiNo
Public m_ColRiesgosAplicados As Scripting.Dictionary
Public m_RiesgoUltimoSeleccionado As riesgo
Public blnPermitidoEditar As Boolean

Public m_EdicionAlAbrir As Edicion

Private m_Error As String

Private Sub ComandoPublicacion_Click()
    
    Dim m_NombreFormulario As String
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If EsTecnico = EnumSiNo.Sí Then
        m_NombreFormulario = "FormPublicacionTecnico"
    Else
        m_NombreFormulario = "FormPublicacionCalidad"
    End If
    If FormularioAbierto(m_NombreFormulario) Then
        DoCmd.Close acForm, m_NombreFormulario, acSaveNo
    End If
    DoCmd.OpenForm m_NombreFormulario
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoPublicacion_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
Private Sub ComandoRiesgoPriorizacion_Click()
    On Error GoTo errores
    
    
    
    
    m_Error = ""
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If FormularioAbierto("FormRiesgosEstablecerPrioridades") Then
        DoCmd.Close acForm, "FormRiesgosEstablecerPrioridades", acSaveNo
    End If
    DoCmd.OpenForm "FormRiesgosEstablecerPrioridades"
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRiesgoPriorizacion_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

' ----------------------------------------------------------------------------
' Public Sub Refrescar (HR4 UI adapter)
'   Slice HR4b: el proxy de modFormCoordinationHelper.Coord_RefreshRiesgosGestion
'   llama este Sub (UI adapter tiny, Hard rule 4 OK). Delega a Private
'   EstablecerDatos.
' ----------------------------------------------------------------------------
Public Sub Refrescar(Optional ByRef p_Error As String)
    EstablecerDatos p_Error
End Sub

' ----------------------------------------------------------------------------
' Private Function EstablecerDatos (HR4 fix — antes era Public)
'   Slice HR4b: la logica pura (PermitidoEditar, Caption, ComboVerRetirados/
'   ComboVerDescripcion, EsEdicionActiva) se extrajo a
'   modRiesgoEstablecerDatosHelper.RiesgoEstablecerDatos_CalcularRiesgosGestion.
'   Esta funcion queda como thin coordinator: resuelve contexto del form
'   (m_ObjEdicionActiva / m_ObjProyectoActivo), puebla ComboEdicion, llama
'   al helper, y aplica el resultado a los controles UI + side effects
'   (RellenarListaImagenes, RegistrarUltimoProyecto, CargarArbol,
'    ActualizarEtiquetaUltimoCambio).
' ----------------------------------------------------------------------------
Private Function EstablecerDatos(Optional ByRef p_Error As String) As String

    Dim m_ColEdiciones As Scripting.Dictionary
    Dim m_IDEdicion As Variant
    Dim m_Edicion As Edicion
    Dim cmb As ComboBox
    Dim m_ResultadoHelper As RiesgoEstablecerDatos_RiesgosGestion_Resultado
    Dim m_CalcError As String

    On Error GoTo errores
    m_Error = ""
    Set cmb = Me.ComboEdicion
    cmb.RowSource = ""
    Me.ComandoAnexos.Enabled = False
    Me.AllowEdits = False
    On Error GoTo errores

    ' --- Resolver contexto edicion/proyecto ---
    If Not m_EdicionAlAbrir Is Nothing Then
        Set m_ObjProyectoActivo = m_EdicionAlAbrir.Proyecto
        Set m_ObjEdicionActiva = m_EdicionAlAbrir
    Else
        If m_ObjProyectoActivo Is Nothing Then
            If m_ObjEdicionActiva Is Nothing Then
                p_Error = "No se sabe la gestión de riesgos activa"
                Err.Raise 1000
            End If
            Set m_ObjProyectoActivo = m_ObjEdicionActiva.Proyecto
            If m_ObjEdicionActiva Is Nothing Then
                p_Error = "No se sabe la gestión de riesgos activa"
                Err.Raise 1000
            End If
        Else
            Set m_ObjEdicionActiva = m_ObjProyectoActivo.EdicionUltima
        End If
    End If

    ' --- Side effect: poblar ComboEdicion con las ediciones del proyecto ---
    Set m_ColEdiciones = m_ObjProyectoActivo.colEdiciones
    If Not m_ColEdiciones Is Nothing Then
        For Each m_IDEdicion In m_ColEdiciones
            Set m_Edicion = m_ColEdiciones(m_IDEdicion)
            cmb.AddItem m_Edicion.Edicion
            Set m_Edicion = Nothing
        Next
    End If

    ' --- Pure compute: extraer via helper (Hard rule 1) ---
    modRiesgoEstablecerDatosHelper.RiesgoEstablecerDatos_CalcularRiesgosGestion _
        m_ObjEntorno, m_ObjEdicionActiva, m_ResultadoHelper, m_CalcError
    If m_CalcError <> "" Then
        p_Error = "EstablecerDatos (RiesgoEstablecerDatos_CalcularRiesgosGestion): " & m_CalcError
        Err.Raise 1000
    End If

    ' --- Aplicar resultado del helper a controles / flags del form ---
    m_EsEdicionActiva = m_ResultadoHelper.EsEdicionActiva
    Me.ComboEdicion = m_ObjEdicionActiva.Edicion
    Me.Caption = m_ResultadoHelper.Caption
    Me.ComboVerRetirados = m_ResultadoHelper.ComboVerRetirados
    Me.ComboVerDescripcion = m_ResultadoHelper.ComboVerDescripcion
    blnPermitidoEditar = m_ResultadoHelper.PermitidoEditar

    ' --- Side effects (UI rendering) ---
    Set m_Arbol = Me.Arbol.Object
    Set m_ListImages = Me.ListaImagenes.Object
    RellenarListaImagenes m_ListImages, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RegistrarUltimoProyecto m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    CargarArbol p_Refrescando:=EnumSiNo.No, p_Error:=m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If EsTecnico = EnumSiNo.Sí Then
        Me.ComandoEditarProyecto.Visible = False
    Else
        Me.ComandoEditarProyecto.Visible = True
        Me.ComandoEditarProyecto.Enabled = True
    End If
    If blnPermitidoEditar = True Then
        Me.ComandoRiesgoPriorizacion.Enabled = True
    Else
        Me.ComandoRiesgoPriorizacion.Enabled = False
    End If
    Me.ComandoPublicacion.Enabled = True

    ActualizarEtiquetaUltimoCambio

    Me.AllowEdits = True
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error: " & vbNewLine & Err.description
    End If
    Me.AllowEdits = True
End Function
Public Function ActualizarEtiquetaUltimoCambio(Optional ByRef p_Error As String) As String
    If m_ObjEdicionActiva.UsuarioUltimoCambio <> "" Then
        Me.lblUltimoCambio.Visible = True
        If Not m_ObjEdicionActiva.UsuarioUltimoCambioObj Is Nothing Then
            Me.lblUltimoCambio.Caption = "Último cambio " & m_ObjEdicionActiva.FechaUltimoCambio & " por " & m_ObjEdicionActiva.UsuarioUltimoCambioObj.Nombre
        Else
            Me.lblUltimoCambio.Caption = "Último cambio " & m_ObjEdicionActiva.FechaUltimoCambio & " por desconocido"
        End If
        
    Else
        Me.lblUltimoCambio.Visible = False
    End If
End Function

Public Function RefrescarArbolRiesgosScope( _
    ByVal p_Scope As String, _
    Optional ByRef p_Error As String _
) As String
    Dim m_Scope As String
    Dim m_IDEdicion As String
    Dim m_IdRiesgo As String
    Dim m_Plan As Scripting.Dictionary
    Dim m_RefreshService As RiskTreeRefreshService
    Dim m_Nodo As MSComctlLib.Node
    Dim m_NodoRiesgo As MSComctlLib.Node
    Dim m_VerDescripcion As String
    Dim m_Accion As Object
    Dim m_RequestedKey As String
    Dim m_CandidateKeys As String
    Dim m_ExistingKeys As String
    Dim m_FallbackKey As String

    On Error GoTo errores

    p_Error = ""
    Set m_RefreshService = New RiskTreeRefreshService
    m_IDEdicion = vbNullString
    m_IdRiesgo = vbNullString
    If Not m_ObjEdicionActiva Is Nothing Then
        m_IDEdicion = Nz(m_ObjEdicionActiva.IDEdicion, vbNullString)
    End If
    If Not m_ObjRiesgoActivo Is Nothing Then
        m_IdRiesgo = Nz(m_ObjRiesgoActivo.IDRiesgo, vbNullString)
    End If

    Set m_Plan = m_RefreshService.BuildPlan( _
        p_Scope, _
        m_IDEdicion, _
        m_IdRiesgo, _
        vbNullString, _
        vbNullString, _
        vbNullString, _
        p_Error)
    If p_Error <> "" Then Err.Raise 1000

    m_Scope = CStr(m_Plan("normalizedScope"))

    Set m_Nodo = m_Arbol.SelectedItem
    If m_Nodo Is Nothing Then
        Set m_Nodo = m_NodoSeleccionado
    End If

    m_VerDescripcion = Nz(Me.ComboVerDescripcion, "")
    InvalidarCacheArbolPorScope m_Scope, p_Error
    If p_Error <> "" Then Err.Raise 1000

    If CBool(m_Plan("fullReload")) Then
        If p_Error <> "" Then Err.Raise 1000

        RehidratarEdicionActivaTrasInvalidacion p_Error
        If p_Error <> "" Then Err.Raise 1000

        m_RequestedKey = ObtenerClaveNodoSeleccionadoArbol(p_Error)
        If p_Error <> "" Then Err.Raise 1000
        m_CandidateKeys = ObtenerCandidatosFallbackNodoSeleccionadoArbol(p_Error)
        If p_Error <> "" Then Err.Raise 1000

        CargarArbol p_Refrescando:=EnumSiNo.No, p_Error:=p_Error
        If p_Error <> "" Then Err.Raise 1000

        m_ExistingKeys = ObtenerClavesArbolRiesgos(p_Error)
        If p_Error <> "" Then Err.Raise 1000
        Set m_Plan = m_RefreshService.BuildPlan( _
            m_Scope, _
            m_IDEdicion, _
            m_IdRiesgo, _
            m_RequestedKey, _
            m_CandidateKeys, _
            m_ExistingKeys, _
            p_Error)
        If p_Error <> "" Then Err.Raise 1000

        m_FallbackKey = Nz(m_Plan("fallbackKey"), vbNullString)
        If p_Error <> "" Then Err.Raise 1000
        RestaurarSeleccionArbolPorClave m_FallbackKey, p_Error
        If p_Error <> "" Then Err.Raise 1000

        RefrescarArbolRiesgosScope = m_Scope
        Exit Function
    End If

    Select Case m_Scope
        Case "plan"
            CambiaIconoElementoArbolDeRiesgos m_Arbol, p_Error
            If p_Error <> "" Then Err.Raise 1000

        Case "action"
            If Not m_Nodo Is Nothing Then
                If Not m_ObjPMAccionActiva Is Nothing Then
                    Set m_Accion = m_ObjPMAccionActiva
                ElseIf Not m_ObjPCAccionActiva Is Nothing Then
                    Set m_Accion = m_ObjPCAccionActiva
                End If
                If Not m_Accion Is Nothing Then
                    m_Nodo.Image = getNombreIconoAccionArbolDeRiesgos(m_Accion, p_Error)
                    If p_Error <> "" Then Err.Raise 1000
                End If
            End If

            If Not m_ObjPMActivo Is Nothing Then
                ActualizarIconoPlan p_Plan:=m_ObjPMActivo, p_Nodo:=m_Nodo, p_Error:=p_Error
                If p_Error <> "" Then Err.Raise 1000
                Set m_NodoRiesgo = ObtenerNodoRiesgoDesdeNodoPlan(m_Nodo, p_Error)
            ElseIf Not m_ObjPCActivo Is Nothing Then
                ActualizarIconoPlan p_Plan:=m_ObjPCActivo, p_Nodo:=m_Nodo, p_Error:=p_Error
                If p_Error <> "" Then Err.Raise 1000
                Set m_NodoRiesgo = ObtenerNodoRiesgoDesdeNodoPlan(m_Nodo, p_Error)
            End If

            If p_Error = "" And Not m_ObjRiesgoActivo Is Nothing Then
                If m_NodoRiesgo Is Nothing Then Set m_NodoRiesgo = m_Nodo
                ActualizarIconoRiesgo p_Riesgo:=m_ObjRiesgoActivo, p_Nodo:=m_NodoRiesgo, p_VerDescripcion:=m_VerDescripcion, p_Error:=p_Error
                If p_Error <> "" Then Err.Raise 1000
            End If

        Case "risk"
            If Not m_ObjRiesgoActivo Is Nothing Then
                ActualizarIconoRiesgo p_Riesgo:=m_ObjRiesgoActivo, p_Nodo:=m_Nodo, p_VerDescripcion:=m_VerDescripcion, p_Error:=p_Error
                If p_Error <> "" Then Err.Raise 1000
            End If

        Case "action-delete"
            RefrescarArbolPorBorradoAccion p_Error
            If p_Error <> "" Then Err.Raise 1000
    End Select

    RefrescarArbolRiesgosScope = m_Scope
    Exit Function
errores:
    If p_Error = "" Then
        p_Error = "El método RefrescarArbolRiesgosScope ha devuelto el error: " & vbNewLine & Err.description
    End If
    Err.Raise IIf(Err.Number = 0, 1000, Err.Number), "RefrescarArbolRiesgosScope", p_Error
End Function

Public Function RefrescarNodoRiesgoActual( _
    ByVal p_IDRiesgo As String, _
    Optional ByRef p_Error As String _
) As String
    Dim m_NodoRiesgo As MSComctlLib.Node
    Dim m_RiesgoRecargado As riesgo
    Dim m_VerDescripcion As String

    On Error GoTo errores

    p_Error = ""
    RefrescarNodoRiesgoActual = ResolverRutaRefrescoDetalleRiesgo(False, p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Set m_NodoRiesgo = ResolverNodoRiesgoActual(p_IDRiesgo:=p_IDRiesgo, p_Error:=p_Error)
    If p_Error <> "" Then Err.Raise 1000

    If m_NodoRiesgo Is Nothing Then
        RefrescarArbolRiesgosScope p_Scope:="full", p_Error:=p_Error
        If p_Error <> "" Then Err.Raise 1000
        Exit Function
    End If

    Set m_RiesgoRecargado = RehidratarRiesgoActivoNodo(p_IDRiesgo:=p_IDRiesgo, p_Error:=p_Error)
    If p_Error <> "" Then Err.Raise 1000
    If m_RiesgoRecargado Is Nothing Then
        p_Error = "No se pudo rehidratar el riesgo activo para refresco puntual (IDRiesgo=" & p_IDRiesgo & ")"
        Err.Raise 1000
    End If

    m_VerDescripcion = Nz(Me.ComboVerDescripcion, "No")
    If Trim$(m_VerDescripcion) = "" Then m_VerDescripcion = "No"

    ActualizarIconoRiesgo p_Riesgo:=m_RiesgoRecargado, p_Nodo:=m_NodoRiesgo, p_VerDescripcion:=m_VerDescripcion, p_Error:=p_Error
    If p_Error <> "" Then Err.Raise 1000

    m_Arbol.SelectedItem = m_NodoRiesgo
    m_Arbol.Refresh
    RefrescarNodoRiesgoActual = ResolverRutaRefrescoDetalleRiesgo(True, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    Exit Function
errores:
    If p_Error = "" Then
        p_Error = "El método RefrescarNodoRiesgoActual ha devuelto el error: " & vbNewLine & Err.description
    End If
    Err.Raise IIf(Err.Number = 0, 1000, Err.Number), "RefrescarNodoRiesgoActual", p_Error
End Function

Private Function ResolverNodoRiesgoActual( _
    ByVal p_IDRiesgo As String, _
    Optional ByRef p_Error As String _
) As MSComctlLib.Node
    Dim m_ClaveEsperada As String
    Dim m_Nodo As MSComctlLib.Node

    On Error GoTo errores

    p_Error = ""
    m_ClaveEsperada = "RIESGO|" & CStr(p_IDRiesgo)

    Set m_Nodo = m_Arbol.SelectedItem
    If m_Nodo Is Nothing Then Set m_Nodo = m_NodoSeleccionado

    If Not m_Nodo Is Nothing Then
        If StrComp(CStr(m_Nodo.key), m_ClaveEsperada, vbTextCompare) = 0 Then
            Set ResolverNodoRiesgoActual = m_Nodo
            Exit Function
        End If
    End If

    If ExisteNodoArbolPorClave(m_ClaveEsperada, p_Error) Then
        Set ResolverNodoRiesgoActual = m_Arbol.Nodes(m_ClaveEsperada)
    End If
    If p_Error <> "" Then Err.Raise 1000
    Exit Function
errores:
    If p_Error = "" Then
        p_Error = "El método ResolverNodoRiesgoActual ha devuelto el error: " & vbNewLine & Err.description
    End If
    Err.Raise IIf(Err.Number = 0, 1000, Err.Number), "ResolverNodoRiesgoActual", p_Error
End Function

Private Function RehidratarRiesgoActivoNodo( _
    ByVal p_IDRiesgo As String, _
    Optional ByRef p_Error As String _
) As riesgo
    On Error GoTo errores

    p_Error = ""
    InvalidarCacheRiesgo p_IDRiesgo:=CStr(p_IDRiesgo), p_Error:=p_Error
    If p_Error <> "" Then Err.Raise 1000

    Set m_ObjRiesgoActivo = GetCachedRiesgo(p_IDRiesgo:=CStr(p_IDRiesgo), p_Error:=p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Set RehidratarRiesgoActivoNodo = m_ObjRiesgoActivo
    Exit Function
errores:
    If p_Error = "" Then
        p_Error = "El método RehidratarRiesgoActivoNodo ha devuelto el error: " & vbNewLine & Err.description
    End If
    Err.Raise IIf(Err.Number = 0, 1000, Err.Number), "RehidratarRiesgoActivoNodo", p_Error
End Function

Private Function RehidratarEdicionActivaTrasInvalidacion( _
                                                        Optional ByRef p_Error As String _
                                                        ) As String
    Dim m_IDEdicion As String

    On Error GoTo errores

    p_Error = ""
    If m_ObjEdicionActiva Is Nothing Then Exit Function

    m_IDEdicion = CStr(m_ObjEdicionActiva.IDEdicion)
    Set m_ObjEdicionActiva = GetCachedEdicion(p_IDEdicion:=m_IDEdicion, p_Error:=p_Error)
    If p_Error <> "" Then Err.Raise 1000
    If m_ObjEdicionActiva Is Nothing Then
        p_Error = "No se pudo rehidratar la edición activa luego de invalidar caché (IDEdicion=" & m_IDEdicion & ")"
        Err.Raise 1000
    End If

    Set m_ObjProyectoActivo = m_ObjEdicionActiva.Proyecto
    If m_ObjProyectoActivo Is Nothing Then
        p_Error = "No se pudo rehidratar el proyecto activo desde la edición activa (IDEdicion=" & m_IDEdicion & ")"
        Err.Raise 1000
    End If
    Exit Function
errores:
    If p_Error = "" Then
        p_Error = "El método RehidratarEdicionActivaTrasInvalidacion ha devuelto el error: " & vbNewLine & Err.description
    End If
    Err.Raise IIf(Err.Number = 0, 1000, Err.Number), "RehidratarEdicionActivaTrasInvalidacion", p_Error
End Function

Private Function RefrescarArbolPorBorradoAccion( _
                                                Optional ByRef p_Error As String _
                                                ) As String
    Dim m_NodoAccion As MSComctlLib.Node
    Dim m_NodoPlan As MSComctlLib.Node
    Dim m_NodoRiesgo As MSComctlLib.Node
    Dim m_Plan As Object
    Dim m_RequestedKey As String
    Dim m_CandidateKeys As String
    Dim m_ExistingKeys As String
    Dim m_FallbackKey As String
    Dim m_VerDescripcion As String

    On Error GoTo errores

    p_Error = ""
    m_VerDescripcion = Nz(Me.ComboVerDescripcion, "")
    Set m_NodoAccion = m_Arbol.SelectedItem
    If m_NodoAccion Is Nothing Then Set m_NodoAccion = m_NodoSeleccionado

    m_RequestedKey = ObtenerClaveNodoSeleccionadoArbol(p_Error)
    If p_Error <> "" Then Err.Raise 1000
    m_CandidateKeys = ObtenerCandidatosFallbackNodoSeleccionadoArbol(p_Error)
    If p_Error <> "" Then Err.Raise 1000

    If Not m_NodoAccion Is Nothing Then
        Set m_NodoPlan = m_NodoAccion.Parent
        If Not m_NodoPlan Is Nothing Then Set m_NodoRiesgo = m_NodoPlan.Parent
    End If

    If Not m_ObjPMAccionActiva Is Nothing Then
        Set m_Plan = m_ObjPMAccionActiva.Mitigacion
    ElseIf Not m_ObjPCAccionActiva Is Nothing Then
        Set m_Plan = m_ObjPCAccionActiva.Contingencia
    End If

    If Trim$(m_RequestedKey) <> "" Then
        If ExisteNodoArbolPorClave(m_RequestedKey, p_Error) Then
            m_Arbol.Nodes.Remove m_RequestedKey
        End If
        If p_Error <> "" Then Err.Raise 1000
    End If

    If Not m_Plan Is Nothing Then
        ActualizarIconoPlan p_Plan:=m_Plan, p_Nodo:=m_NodoPlan, p_Error:=p_Error
        If p_Error <> "" Then Err.Raise 1000
    End If

    If Not m_ObjRiesgoActivo Is Nothing Then
        ActualizarIconoRiesgo p_Riesgo:=m_ObjRiesgoActivo, p_Nodo:=m_NodoRiesgo, p_VerDescripcion:=m_VerDescripcion, p_Error:=p_Error
        If p_Error <> "" Then Err.Raise 1000
    End If

    m_ExistingKeys = ObtenerClavesArbolRiesgos(p_Error)
    If p_Error <> "" Then Err.Raise 1000
    m_FallbackKey = ResolverClaveFallbackArbolRiesgosConExistentes(m_RequestedKey, m_CandidateKeys, m_ExistingKeys, p_Error)
    If p_Error <> "" Then Err.Raise 1000

    Set m_ObjPMAccionActiva = Nothing
    Set m_ObjPCAccionActiva = Nothing
    RestaurarSeleccionArbolPorClave m_FallbackKey, p_Error
    If p_Error <> "" Then Err.Raise 1000

    m_Arbol.Refresh
    RefrescarArbolPorBorradoAccion = "OK"
    Exit Function
errores:
    If p_Error = "" Then
        p_Error = "El método RefrescarArbolPorBorradoAccion ha devuelto el error: " & vbNewLine & Err.description
    End If
    Err.Raise IIf(Err.Number = 0, 1000, Err.Number), "RefrescarArbolPorBorradoAccion", p_Error
End Function

Private Function ObtenerClavesArbolRiesgos( _
                                           Optional ByRef p_Error As String _
                                           ) As String
    Dim m_Nodo As MSComctlLib.Node
    Dim m_Serialized As String

    On Error GoTo errores

    p_Error = ""
    For Each m_Nodo In m_Arbol.Nodes
        If Trim$(m_Serialized) = "" Then
            m_Serialized = CStr(m_Nodo.key)
        Else
            m_Serialized = m_Serialized & vbLf & CStr(m_Nodo.key)
        End If
    Next m_Nodo

    ObtenerClavesArbolRiesgos = m_Serialized
    Exit Function
errores:
    p_Error = "El método ObtenerClavesArbolRiesgos ha devuelto el error: " & vbNewLine & Err.description
    Err.Raise IIf(Err.Number = 0, 1000, Err.Number), "ObtenerClavesArbolRiesgos", p_Error
End Function

Private Function InvalidarCacheArbolPorScope( _
    ByVal p_Scope As String, _
    Optional ByRef p_Error As String _
) As String
    Dim m_Orden As String
    Dim m_Partes() As String
    Dim m_Item As Variant

    On Error GoTo errores

    p_Error = ""
    m_Orden = ObtenerOrdenInvalidacionRiesgosPorScope(p_Scope, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    If Trim$(m_Orden) = "" Then Exit Function

    m_Partes = Split(m_Orden, "|")
    For Each m_Item In m_Partes
        Select Case LCase$(Trim$(CStr(m_Item)))
            Case "plan"
                If Not m_ObjPMActivo Is Nothing Then
                    InvalidarCachePM m_ObjPMActivo.IDMitigacion
                ElseIf Not m_ObjPCActivo Is Nothing Then
                    InvalidarCachePC m_ObjPCActivo.IDContingencia
                End If
            Case "risk"
                If EsRefrescoCompletoArbolRiesgosScope(p_Scope, p_Error) Then
                    If p_Error <> "" Then Err.Raise 1000
                    InvalidarCacheRiesgosDeEdicionActiva p_Error
                    If p_Error <> "" Then Err.Raise 1000
                Else
                    If Not m_ObjRiesgoActivo Is Nothing Then
                        InvalidarCacheRiesgo m_ObjRiesgoActivo.IDRiesgo
                    End If
                End If
            Case "edition"
                If Not m_ObjEdicionActiva Is Nothing Then
                    InvalidarCacheEdicion m_ObjEdicionActiva.IDEdicion
                End If
        End Select
    Next m_Item

    Exit Function
errores:
    If p_Error = "" Then
        p_Error = "El método InvalidarCacheArbolPorScope ha devuelto el error: " & vbNewLine & Err.description
    End If
    Err.Raise IIf(Err.Number = 0, 1000, Err.Number), "InvalidarCacheArbolPorScope", p_Error
End Function

Private Function InvalidarCacheRiesgosDeEdicionActiva( _
                                                     Optional ByRef p_Error As String _
                                                     ) As String
    Dim m_ColRiesgos As Scripting.Dictionary
    Dim m_IdRiesgo As Variant

    On Error GoTo errores

    p_Error = ""
    If m_ObjEdicionActiva Is Nothing Then Exit Function

    Set m_ColRiesgos = m_ObjEdicionActiva.colRiesgos
    If m_ColRiesgos Is Nothing Then Exit Function

    For Each m_IdRiesgo In m_ColRiesgos
        InvalidarCacheRiesgo CStr(m_IdRiesgo), p_Error
        If p_Error <> "" Then Err.Raise 1000
    Next m_IdRiesgo

    Exit Function
errores:
    If p_Error = "" Then
        p_Error = "El método InvalidarCacheRiesgosDeEdicionActiva ha devuelto el error: " & vbNewLine & Err.description
    End If
    Err.Raise IIf(Err.Number = 0, 1000, Err.Number), "InvalidarCacheRiesgosDeEdicionActiva", p_Error
End Function

Private Function ObtenerClaveNodoSeleccionadoArbol( _
                                                   Optional ByRef p_Error As String _
                                                   ) As String
    Dim m_Nodo As MSComctlLib.Node

    On Error GoTo errores

    p_Error = ""
    Set m_Nodo = m_Arbol.SelectedItem
    If m_Nodo Is Nothing Then Set m_Nodo = m_NodoSeleccionado
    If Not m_Nodo Is Nothing Then ObtenerClaveNodoSeleccionadoArbol = CStr(m_Nodo.key)
    Exit Function
errores:
    p_Error = "El método ObtenerClaveNodoSeleccionadoArbol ha devuelto el error: " & vbNewLine & Err.description
    Err.Raise IIf(Err.Number = 0, 1000, Err.Number), "ObtenerClaveNodoSeleccionadoArbol", p_Error
End Function

Private Function ObtenerCandidatosFallbackNodoSeleccionadoArbol( _
                                                                 Optional ByRef p_Error As String _
                                                                 ) As String
    Dim m_Nodo As MSComctlLib.Node
    Dim m_Serialized As String

    On Error GoTo errores

    p_Error = ""
    Set m_Nodo = m_Arbol.SelectedItem
    If m_Nodo Is Nothing Then Set m_Nodo = m_NodoSeleccionado

    Do While Not m_Nodo Is Nothing
        If Trim$(m_Serialized) = "" Then
            m_Serialized = CStr(m_Nodo.key)
        Else
            m_Serialized = m_Serialized & vbLf & CStr(m_Nodo.key)
        End If
        Set m_Nodo = m_Nodo.Parent
    Loop

    ObtenerCandidatosFallbackNodoSeleccionadoArbol = m_Serialized
    Exit Function
errores:
    p_Error = "El método ObtenerCandidatosFallbackNodoSeleccionadoArbol ha devuelto el error: " & vbNewLine & Err.description
    Err.Raise IIf(Err.Number = 0, 1000, Err.Number), "ObtenerCandidatosFallbackNodoSeleccionadoArbol", p_Error
End Function

Private Function ExisteNodoArbolPorClave( _
                                          ByVal p_Key As String, _
                                          Optional ByRef p_Error As String _
                                          ) As Boolean
    Dim m_Nodo As MSComctlLib.Node

    On Error GoTo errores

    p_Error = ""
    ExisteNodoArbolPorClave = False
    If Trim$(Nz(p_Key, "")) = "" Then Exit Function

    On Error Resume Next
    Set m_Nodo = m_Arbol.Nodes(p_Key)
    If Err.Number <> 0 Then
        Err.Clear
        Set m_Nodo = Nothing
    End If
    On Error GoTo errores

    ExisteNodoArbolPorClave = Not m_Nodo Is Nothing
    Exit Function
errores:
    p_Error = "El método ExisteNodoArbolPorClave ha devuelto el error: " & vbNewLine & Err.description
    Err.Raise IIf(Err.Number = 0, 1000, Err.Number), "ExisteNodoArbolPorClave", p_Error
End Function

Private Function RestaurarSeleccionArbolPorClave( _
                                                 ByVal p_Key As String, _
                                                 Optional ByRef p_Error As String _
                                                 ) As String
    Dim m_Nodo As MSComctlLib.Node

    On Error GoTo errores

    p_Error = ""
    If Trim$(Nz(p_Key, "")) = "" Then Exit Function
    If Not ExisteNodoArbolPorClave(p_Key, p_Error) Then Exit Function
    If p_Error <> "" Then Err.Raise 1000

    Set m_Nodo = m_Arbol.Nodes(p_Key)
    m_Arbol.SelectedItem = m_Nodo
    Arbol_NodeClick m_Nodo
    Exit Function
errores:
    p_Error = "El método RestaurarSeleccionArbolPorClave ha devuelto el error: " & vbNewLine & Err.description
    Err.Raise IIf(Err.Number = 0, 1000, Err.Number), "RestaurarSeleccionArbolPorClave", p_Error
End Function

Public Sub Arbol_NodeClick(ByVal Node As Object)
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    Me.AllowEdits = False
    
    m_URLInformeRiesgo = ""
    Set m_NodoSeleccionado = Node
    Set m_ObjSeleccionado = getObjetoSeleccionadoDeRiesgos(Node.key, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.ComandoAnexos.Enabled = True
    If TypeOf m_ObjSeleccionado Is Edicion Then
        If Me.FormDetalle.SourceObject <> "" Then
            Me.FormDetalle.SourceObject = ""
        End If
        Set m_RiesgoUltimoSeleccionado = Nothing
        Me.lblEstadoRiesgo.Visible = False
        
        Set m_ObjEdicionActiva = m_ObjSeleccionado
        Set m_ObjPMAccionActiva = Nothing
        Set m_ObjPCAccionActiva = Nothing
        If m_ObjEdicionActiva.EsActivo = EnumSiNo.No Then
            blnPermitidoEditar = False
        Else
            If m_ObjEdicionActiva.Proyecto.UsuarioAutorizado = EnumSiNo.Sí Then
                blnPermitidoEditar = True
            
            Else
                blnPermitidoEditar = True
            End If
        End If
        
        
        If m_SinCargarFormulario <> EnumSiNo.Sí Then
            Me.FormDetalle.SourceObject = "FormRiesgosGestionEdicion"
        End If
        
        AvanceCerrar
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Me.AllowEdits = True
        Exit Sub
    End If
    If TypeOf m_ObjSeleccionado Is riesgo Then
        Set m_RiesgoSeleccionado = m_ObjSeleccionado
        m_EstadoRiesgoSeleccionado = m_RiesgoSeleccionado.EstadoEnum
        Set m_ObjRiesgoActivo = m_RiesgoSeleccionado
        Set m_ObjPMAccionActiva = Nothing
        Set m_ObjPCAccionActiva = Nothing
        m_EstadoRiesgoActivo = m_EstadoRiesgoSeleccionado
        If m_SinCargarFormulario <> EnumSiNo.Sí Then
            Me.FormDetalle.SourceObject = "FormRiesgosGestionRiesgo"
        End If
        If m_RiesgoUltimoSeleccionado Is Nothing Then
            If m_ObjEdicionActiva.EsActivo Then
                RellenaEtiquetaRiesgo Me.lblEstadoRiesgo, m_EstadoRiesgoSeleccionado, m_Error
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
                Me.lblEstadoRiesgo.Visible = True
            Else
                Me.lblEstadoRiesgo.Visible = False
            End If
            Set m_RiesgoUltimoSeleccionado = GetCachedRiesgo(m_RiesgoSeleccionado.IDRiesgo)
            
        Else
            If m_RiesgoSeleccionado.IDRiesgo <> m_RiesgoUltimoSeleccionado.IDRiesgo Then
                If m_ObjEdicionActiva.EsActivo Then
                    RellenaEtiquetaRiesgo Me.lblEstadoRiesgo, m_EstadoRiesgoSeleccionado, m_Error
                    If m_Error <> "" Then
                        Err.Raise 1000
                    End If
                    Me.lblEstadoRiesgo.Visible = True
                Else
                    Me.lblEstadoRiesgo.Visible = False
                End If
                Set m_RiesgoUltimoSeleccionado = GetCachedRiesgo(m_RiesgoSeleccionado.IDRiesgo)
            End If
        End If
        AvanceCerrar
        
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Me.AllowEdits = True
        Exit Sub
    End If
    If TypeOf m_ObjSeleccionado Is pm Then
        Set m_PMSeleccionado = m_ObjSeleccionado
        Me.m_EsMitigacion = EnumSiNo.Sí
        Set m_PMSeleccionado = m_ObjSeleccionado
        Set m_ObjPMActivo = m_PMSeleccionado
        Set m_ObjPMAccionActiva = Nothing
        Set m_ObjPCAccionActiva = Nothing
        
        Set m_ObjRiesgoActivo = m_PMSeleccionado.riesgo
        If m_SinCargarFormulario <> EnumSiNo.Sí Then
            Me.FormDetalle.SourceObject = "FormRiesgosGestionPlanPrincipal"
        End If
        If m_RiesgoUltimoSeleccionado Is Nothing Then
            If m_ObjEdicionActiva.EsActivo Then
                RellenaEtiquetaRiesgo Me.lblEstadoRiesgo, m_ObjRiesgoActivo.EstadoEnum, m_Error
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
                Me.lblEstadoRiesgo.Visible = True
            Else
                Me.lblEstadoRiesgo.Visible = False
            End If
           
            Set m_RiesgoUltimoSeleccionado = GetCachedRiesgo(m_ObjRiesgoActivo.IDRiesgo)
        Else
            If m_ObjRiesgoActivo.IDRiesgo <> m_RiesgoUltimoSeleccionado.IDRiesgo Then
                If m_ObjEdicionActiva.EsActivo Then
                    RellenaEtiquetaRiesgo Me.lblEstadoRiesgo, m_ObjRiesgoActivo.EstadoEnum, m_Error
                    If m_Error <> "" Then
                        Err.Raise 1000
                    End If
                    Me.lblEstadoRiesgo.Visible = True
                Else
                    Me.lblEstadoRiesgo.Visible = False
                End If
                Set m_RiesgoUltimoSeleccionado = GetCachedRiesgo(m_ObjRiesgoActivo.IDRiesgo)
            End If
        End If
        
        
        AvanceCerrar
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Me.AllowEdits = True
        Exit Sub
    End If
    If TypeOf m_ObjSeleccionado Is pc Then
        Set m_PCSeleccionado = m_ObjSeleccionado
        Set m_ObjPCActivo = m_PCSeleccionado
        Set m_ObjPMAccionActiva = Nothing
        Set m_ObjPCAccionActiva = Nothing
        Set m_ObjRiesgoActivo = m_PCSeleccionado.riesgo
        Me.m_EsMitigacion = EnumSiNo.No
        
        If m_SinCargarFormulario <> EnumSiNo.Sí Then
            Me.FormDetalle.SourceObject = "FormRiesgosGestionPlanPrincipal"
        End If
        If m_RiesgoUltimoSeleccionado Is Nothing Then
            If m_ObjEdicionActiva.EsActivo Then
                RellenaEtiquetaRiesgo Me.lblEstadoRiesgo, m_ObjRiesgoActivo.EstadoEnum, m_Error
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
                Me.lblEstadoRiesgo.Visible = True
            Else
                Me.lblEstadoRiesgo.Visible = False
            End If
           
            Set m_RiesgoUltimoSeleccionado = GetCachedRiesgo(m_ObjRiesgoActivo.IDRiesgo)
        Else
            If m_ObjRiesgoActivo.IDRiesgo <> m_RiesgoUltimoSeleccionado.IDRiesgo Then
                If m_ObjEdicionActiva.EsActivo Then
                    RellenaEtiquetaRiesgo Me.lblEstadoRiesgo, m_ObjRiesgoActivo.EstadoEnum, m_Error
                    If m_Error <> "" Then
                        Err.Raise 1000
                    End If
                    Me.lblEstadoRiesgo.Visible = True
                Else
                    Me.lblEstadoRiesgo.Visible = False
                End If
                Set m_RiesgoUltimoSeleccionado = GetCachedRiesgo(m_ObjRiesgoActivo.IDRiesgo)
            End If
        End If
        AvanceCerrar
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Me.AllowEdits = True
        Exit Sub
    End If
    If TypeOf m_ObjSeleccionado Is PMAccion Then
        Set m_PMAccionSeleccionado = m_ObjSeleccionado
        Set m_ObjPMAccionActiva = m_PMAccionSeleccionado
        Set m_ObjPCAccionActiva = Nothing
        Set m_ObjRiesgoActivo = m_PMAccionSeleccionado.Mitigacion.riesgo
        Me.m_EsMitigacion = EnumSiNo.Sí
        
        If m_SinCargarFormulario <> EnumSiNo.Sí Then
            Me.FormDetalle.SourceObject = "FormRiesgosGestionPlanAcciones"
        End If
        If m_RiesgoUltimoSeleccionado Is Nothing Then
            If m_ObjEdicionActiva.EsActivo Then
                RellenaEtiquetaRiesgo Me.lblEstadoRiesgo, m_ObjRiesgoActivo.EstadoEnum, m_Error
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
                Me.lblEstadoRiesgo.Visible = True
            Else
                Me.lblEstadoRiesgo.Visible = False
            End If
           
            Set m_RiesgoUltimoSeleccionado = GetCachedRiesgo(m_ObjRiesgoActivo.IDRiesgo)
        Else
            If m_ObjRiesgoActivo.IDRiesgo <> m_RiesgoUltimoSeleccionado.IDRiesgo Then
                If m_ObjEdicionActiva.EsActivo Then
                    RellenaEtiquetaRiesgo Me.lblEstadoRiesgo, m_ObjRiesgoActivo.EstadoEnum, m_Error
                    If m_Error <> "" Then
                        Err.Raise 1000
                    End If
                    Me.lblEstadoRiesgo.Visible = True
                Else
                    Me.lblEstadoRiesgo.Visible = False
                End If
                Set m_RiesgoUltimoSeleccionado = GetCachedRiesgo(m_ObjRiesgoActivo.IDRiesgo)
            End If
        End If
        AvanceCerrar
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Me.AllowEdits = True
        Exit Sub
    End If
    If TypeOf m_ObjSeleccionado Is PCAccion Then
        Set m_PCAccionSeleccionado = m_ObjSeleccionado
        Set m_ObjPCAccionActiva = m_PCAccionSeleccionado
        Set m_ObjPMAccionActiva = Nothing
        Set m_ObjRiesgoActivo = m_PCAccionSeleccionado.Contingencia.riesgo
        Me.m_EsMitigacion = EnumSiNo.No
        If m_SinCargarFormulario <> EnumSiNo.Sí Then
            Me.FormDetalle.SourceObject = "FormRiesgosGestionPlanAcciones"
        End If
        If m_RiesgoUltimoSeleccionado Is Nothing Then
            If m_ObjEdicionActiva.EsActivo Then
                RellenaEtiquetaRiesgo Me.lblEstadoRiesgo, m_ObjRiesgoActivo.EstadoEnum, m_Error
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
                Me.lblEstadoRiesgo.Visible = True
            Else
                Me.lblEstadoRiesgo.Visible = False
            End If
           
            Set m_RiesgoUltimoSeleccionado = GetCachedRiesgo(m_ObjRiesgoActivo.IDRiesgo)
        Else
            If m_ObjRiesgoActivo.IDRiesgo <> m_RiesgoUltimoSeleccionado.IDRiesgo Then
                If m_ObjEdicionActiva.EsActivo Then
                    RellenaEtiquetaRiesgo Me.lblEstadoRiesgo, m_ObjRiesgoActivo.EstadoEnum, m_Error
                    If m_Error <> "" Then
                        Err.Raise 1000
                    End If
                    Me.lblEstadoRiesgo.Visible = True
                Else
                    Me.lblEstadoRiesgo.Visible = False
                End If
                Set m_RiesgoUltimoSeleccionado = GetCachedRiesgo(m_ObjRiesgoActivo.IDRiesgo)
            End If
        End If
        AvanceCerrar
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Me.AllowEdits = True
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al Arbol_NodeClick se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub CmdSalir_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Public Sub ComandoActualizarContador_Click()
    
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If Me.FormDetalle.SourceObject <> "" Then
        Me.FormDetalle.SourceObject = ""
    End If
    Dim m_ScopeRefresco As String
    m_ScopeRefresco = ResolverScopeRefrescoManualRiesgos(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    RefrescarArbolRiesgosScope p_Scope:=m_ScopeRefresco, p_Error:=m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoActualizarContador_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
Public Function getNodoSeleccionado(Optional ByRef p_Error As String) As MSComctlLib.Node
    
    Dim m_Nodo As MSComctlLib.Node
    On Error GoTo errores
    
    For Each m_Nodo In m_Arbol.Nodes
        If m_Nodo.Selected = True Then
            Set getNodoSeleccionado = m_Nodo
            Exit Function
        End If
    Next
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getNodoSeleccionado ha producido el error num: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        
    End If
End Function
Public Function getNodoRiesgo(Optional p_Riesgo As riesgo, Optional ByRef p_Error As String) As MSComctlLib.Node
    
   
    On Error GoTo errores
    If p_Riesgo Is Nothing Then
        If m_ObjSeleccionado Is Nothing Then
            p_Error = "No hay nodo seleccionado"
            Err.Raise 1000
        End If
        If TypeOf m_ObjSeleccionado Is pm Then
            Set p_Riesgo = m_ObjSeleccionado.riesgo
        ElseIf TypeOf m_ObjSeleccionado Is PMAccion Then
            Set p_Riesgo = m_ObjSeleccionado.Mitigacion.riesgo
        ElseIf TypeOf m_ObjSeleccionado Is pc Then
            Set p_Riesgo = m_ObjSeleccionado.riesgo
        ElseIf TypeOf m_ObjSeleccionado Is PCAccion Then
            Set p_Riesgo = m_ObjSeleccionado.Contingencia.riesgo
        Else
            p_Error = "No hay nodo seleccionado"
            Err.Raise 1000
        End If
        
        If p_Riesgo Is Nothing Then
            p_Error = "No hay nodo seleccionado"
            Err.Raise 1000
        End If
    End If
    Set getNodoRiesgo = m_Arbol.Nodes("RIESGO|" & p_Riesgo.IDRiesgo)
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getNodoRiesgo ha producido el error num: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        
    End If
End Function
Public Function getNodoPlan(Optional p_Plan As Object, Optional ByRef p_Error As String) As MSComctlLib.Node
    
   
    On Error GoTo errores
    If p_Plan Is Nothing Then
        If m_ObjSeleccionado Is Nothing Then
            p_Error = "No hay nodo seleccionado"
            Err.Raise 1000
        End If
        If TypeOf m_ObjSeleccionado Is pm Then
            Set p_Plan = m_ObjSeleccionado
        ElseIf TypeOf m_ObjSeleccionado Is PMAccion Then
            Set p_Plan = m_ObjSeleccionado.Mitigacion
        ElseIf TypeOf m_ObjSeleccionado Is pc Then
            Set p_Plan = m_ObjSeleccionado
        ElseIf TypeOf m_ObjSeleccionado Is PCAccion Then
            Set p_Plan = m_ObjSeleccionado.Contingencia
        Else
            p_Error = "No hay nodo seleccionado"
            Err.Raise 1000
        End If
        
        If p_Plan Is Nothing Then
            p_Error = "No hay nodo seleccionado"
            Err.Raise 1000
        End If
    End If
    If TypeOf p_Plan Is pm Then
        Set getNodoPlan = m_Arbol.Nodes("PM|" & p_Plan.riesgo.IDRiesgo & "|" & p_Plan.IDMitigacion)
    Else
        Set getNodoPlan = m_Arbol.Nodes("PC|" & p_Plan.riesgo.IDRiesgo & "|" & p_Plan.IDContingencia)
    End If
    
    
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getNodoPlan ha producido el error num: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        
    End If
End Function


Private Sub ComandoAnexos_Click()
   
    Dim m_Argumento As String
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If TypeOf m_ObjSeleccionado Is Edicion Then
        m_Argumento = "Edición"
        
    Else
        m_Argumento = "Riesgo"
        Set m_ObjRiesgoActivo = m_RiesgoSeleccionado
'    Else
'        m_Error = "El nodo seleccionado no es ni del tipo riesgo ni del tipo edición"
'        Err.Raise 1000
    End If
    If FormularioAbierto("FormAnexos") Then
        DoCmd.Close acForm, "FormAnexos", acSaveNo
    End If
   
    
    DoCmd.OpenForm "FormAnexos", openArgs:=m_Argumento
    
    
    

    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAnexos_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub


Private Sub ComandoAyuda_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    AbrirAyuda m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAyuda_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
  
Private Function getRiesgoSeleccioanado(Optional ByRef p_Error As String) As riesgo
    
    On Error GoTo errores
    If TypeOf m_ObjSeleccionado Is riesgo Then
        Set getRiesgoSeleccioanado = m_ObjSeleccionado
        Exit Function
    End If
    If TypeOf m_ObjSeleccionado Is Edicion Then
        Exit Function
    End If
    If TypeOf m_ObjSeleccionado Is pm Then
        Set getRiesgoSeleccioanado = m_ObjSeleccionado.riesgo
        Exit Function
    End If
    If TypeOf m_ObjSeleccionado Is pc Then
        Set getRiesgoSeleccioanado = m_ObjSeleccionado.riesgo
        Exit Function
    End If
    If TypeOf m_ObjSeleccionado Is PMAccion Then
        Set getRiesgoSeleccioanado = m_ObjSeleccionado.Mitigacion.riesgo
        Exit Function
    End If
    If TypeOf m_ObjSeleccionado Is PCAccion Then
        Set getRiesgoSeleccioanado = m_ObjSeleccionado.Contingencia.riesgo
        Exit Function
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getRiesgoSeleccioanado ha devuelto el error: " & _
                Err.Number & vbNewLine & "Detalle: " & Err.description
        
    End If
End Function

Private Function getColRiesgosAAPlicar( _
                                        Optional p_VerRetirados As String = "No", _
                                        Optional ByRef p_Error As String) As Scripting.Dictionary
    
    On Error GoTo errores
    
    If p_VerRetirados = "Sí" Then
        Set getColRiesgosAAPlicar = m_ObjEdicionActiva.colRiesgos
        
    Else
        Set getColRiesgosAAPlicar = m_ObjEdicionActiva.colRiesgosActivos
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getColRiesgosAAPlicar ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.description
        
    End If
End Function

Private Sub ComandoEditarProyecto_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Set m_ObjProyectoActivo = m_ObjEdicionActiva.Proyecto
    m_EsAlta = EnumSiNo.No
    If FormularioAbierto("FormGestionRiesgos") Then
        DoCmd.Close acForm, "FormGestionRiesgos", acSaveNo
    End If
    DoCmd.OpenForm "FormGestionRiesgos"
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoEditarProyecto_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Public Sub ComboEdicion_BeforeUpdate(Cancel As Integer)
    Dim blnCamEdicionCambiada As Boolean
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    Me.ComandoRiesgoPriorizacion.Enabled = False
    Me.ComandoAnexos.Enabled = False
    Me.ComandoPublicacion.Enabled = False
    Me.ComandoEditarProyecto.Enabled = False
    If Nz(Me.ComboEdicion, "") = "" Then
        Set m_ObjEdicionActiva = Nothing
        blnCamEdicionCambiada = True
        
        CargarArbol p_Refrescando:=EnumSiNo.No, p_Error:=m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        ActualizarEtiquetaUltimoCambio
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Me.AllowEdits = True
        Exit Sub
    End If
    If m_ObjEdicionActiva Is Nothing Then
        If m_ObjProyectoActivo Is Nothing Then
            m_Error = "No se puede determinar ni la edición activa ni la gestión de riesgos activa"
            Err.Raise 1000
        End If
        Set m_ObjEdicionActiva = Constructor.getEdicionV2(m_ObjProyectoActivo.IDProyecto, Me.ComboEdicion, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        blnCamEdicionCambiada = True
    Else
        If Me.ComboEdicion <> m_ObjEdicionActiva.Edicion Then
            Set m_ObjEdicionActiva = Constructor.getEdicionV2(m_ObjEdicionActiva.IDProyecto, Me.ComboEdicion, m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            blnCamEdicionCambiada = True
        End If
    End If
    
    Me.ComandoAnexos.Enabled = True
    Me.ComandoPublicacion.Enabled = True
    Me.ComandoEditarProyecto.Enabled = True
    If blnCamEdicionCambiada = True Then
        m_EsEdicionActiva = m_ObjEdicionActiva.EsActivo
        Me.Caption = m_ObjEntorno.TituloUsuarioConectado & " " & "riesgos de la edición " & m_ObjEdicionActiva.Edicion & " de " & m_ObjEdicionActiva.Proyecto.Proyecto
        CargarArbol p_Refrescando:=EnumSiNo.No, p_Error:=m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        ActualizarEtiquetaUltimoCambio
    End If
    If m_EsEdicionActiva = EnumSiNo.Sí Then
        Me.ComandoRiesgoPriorizacion.Enabled = True
    Else
        Me.ComandoRiesgoPriorizacion.Enabled = False
    End If
    Me.AllowEdits = True
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComboEdicion_BeforeUpdate se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub




Private Sub comboVerDescripcion_BeforeUpdate(Cancel As Integer)
    Dim m_VerDescripcion As String
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    
    m_VerDescripcion = Nz(Me.ComboVerDescripcion, "")
    setModoVerArbolRiesgos p_VerDescripcionActual:=m_VerDescripcion, _
                         p_VerSoloNoRetirados:=Nz(Me.ComboVerRetirados, ""), p_Error:=m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    setTitulosNodosRiesgos tv:=Me.Arbol.Object, p_VerDescripcion:=ResolverModoVerDescripcion(m_VerDescripcion), p_Error:=m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Me.AllowEdits = True
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al comboVerDescripcion_BeforeUpdate se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComboVerRetirados_Change()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    setModoVerArbolRiesgos p_VerDescripcionActual:=Nz(Me.ComboVerDescripcion, ""), _
                        p_VerSoloNoRetirados:=Nz(Me.ComboVerRetirados, ""), p_Error:=m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    CargarArbol p_Refrescando:=EnumSiNo.No, p_Error:=m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComboVerRetirados_Change se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub Form_Load()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    If IsNumeric(Nz(Me.openArgs(), "")) Then
        Set m_EdicionAlAbrir = GetCachedEdicion(p_IDEdicion:=Me.openArgs, p_Error:=m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Me.Sep.Left = Me.FormDetalle.Left - Me.Sep.Width
    If Me.FormDetalle.SourceObject <> "" Then
        Me.FormDetalle.SourceObject = ""
    End If
    m_NodoEdicion.Selected = True
    Arbol_NodeClick m_NodoEdicion
    
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub Sep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        MouseCursor IDC_SIZEWE
        m_sngLeftResizePos = X
        MeWidthHalf = Me.InsideWidth / 2
        MeWidth = Me.InsideWidth - 100
        ArbolLeft = Me.Arbol.Left
        SubFormLeft = Me.FormDetalle.Left
    End If
    
End Sub

Private Sub Sep_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    MouseCursor IDC_SIZEWE
    If Button = 1 Then
        On Error Resume Next
        With Me.Sep
            If .Left > MeWidthHalf Then Exit Sub    '.Left = MeWidthHalf - 1:
            If .Left < SepWMin Then
                .Left = SepW
                Me.Arbol.Width = .Left - ArbolLeft
                Me.FormDetalle.Left = .Left + .Width
                Me.FormDetalle.Width = MeWidth - .Left
               
            Else
                .Move .Left + X - m_sngLeftResizePos
                Me.Arbol.Width = .Left - ArbolLeft
                Me.FormDetalle.Left = .Left + .Width
                Me.FormDetalle.Width = MeWidth - .Left
                
            End If
        End With
    End If
End Sub

Private Sub Sep_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        With Me.Sep
            If .Left > MeWidthHalf Then .Left = MeWidthHalf - 0.1
            Me.Arbol.Width = IIf(Me.Arbol.Width < 60, 60, .Left - ArbolLeft)
            .Left = Me.Arbol.Width + Me.Arbol.Left
            Me.FormDetalle.Left = .Left + .Width
            Me.FormDetalle.Width = MeWidth - .Left
            
        End With
    End If
End Sub

Public Function SeleccionarNodo( _
                                p_Objeto As Object, _
                                Optional p_Error As String) As String
    
    Dim m_Nodo As MSComctlLib.Node
    Dim m_Key As String
    Dim m_IdRiesgo As String
    Dim m_IDPM As String
    Dim m_IDPMA As String
    Dim m_IDPC As String
    Dim m_IDPCA As String
    
    Dim m_Riesgo As riesgo
    Dim m_PM As pm
    Dim m_PC As pc
    Dim m_PMA As PMAccion
    Dim m_PCA As PCAccion
    
    On Error GoTo errores
    
    '   RIESGO1  RIESGO|IDRIESGO1
    '       PM1     PM|IDRIESGO1|IDPM1
    '           ACCIÓN1 PMACCION|IDPM1|IDACCION1
    '           ACCIÓN2 PMACCION|IDPM1|IDACCION2
    '       PM2     PM|IDRIESGO1|IDPM2
    '           ACCIÓN1 PMACCION|IDPM2|IDACCION1
    '   RIESGO2  RIESGO|IDRIESGO2
    
    If p_Objeto Is Nothing Then
        p_Error = "Se ha de indicar el Riesgo"
        Err.Raise 1000
    End If
    If TypeOf p_Objeto Is Edicion Then
        m_Key = "EDICION|" & m_ObjEdicionActiva.IDEdicion
        
    ElseIf TypeOf p_Objeto Is riesgo Then
        m_Key = "RIESGO|" & p_Objeto.IDRiesgo
    ElseIf TypeOf p_Objeto Is pm Then
        m_Key = "PM|" & p_Objeto.riesgo.IDRiesgo & "|" & p_Objeto.IDMitigacion
    ElseIf TypeOf p_Objeto Is pc Then
        m_Key = "PC|" & p_Objeto.riesgo.IDRiesgo & "|" & p_Objeto.IDContingencia
    ElseIf TypeOf p_Objeto Is PMAccion Then
        m_Key = "PMACCION|" & p_Objeto.pm.IDMitigacion & "|" & p_Objeto.IDAccionMitigacion
    ElseIf TypeOf p_Objeto Is PCAccion Then
        m_Key = "PCACCION|" & p_Objeto.pc.IDContingencia & "|" & p_Objeto.IDAccionContingencia
    Else
        p_Error = "No se conoce el nodo seleccionado"
        Err.Raise 1000
    End If
    On Error Resume Next
    Set m_Nodo = m_Arbol.Nodes(m_Key)
    If Err.Number <> 0 Then
        On Error GoTo errores
        p_Error = "No se encuentra el riesgo"
        Err.Raise 1000
    End If
    On Error GoTo errores
    m_Arbol.SelectedItem = m_Nodo
    Arbol_NodeClick m_Nodo
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SeleccionarNodo ha devuelto un error: " & vbNewLine & Err.description
    End If
End Function



Public Function CargarArbol( _
                                Optional p_Refrescando As EnumSiNo = EnumSiNo.No, _
                                Optional ByRef p_Error As String _
                                ) As String

    
    
    
'   EDICION1 EDICION|IDEDICION
'       RIESGO1  RIESGO|IDRIESGO1
'           PM1     PM|IDRIESGO1|IDPM1
'               ACCIÓN1 PMACCION|IDPM1|IDACCION1
'               ACCIÓN2 PMACCION|IDPM1|IDACCION2
'           PM2     PM|IDRIESGO1|IDPM2
'               ACCIÓN1 PMACCION|IDPM2|IDACCION1
'       RIESGO2  RIESGO|IDRIESGO2

    
    
    Dim m_Riesgo As riesgo
    Dim m_Key As String
    Dim m_NombreIcono As String
    Dim m_TituloNodo As String
    Dim m_IdRiesgo As Variant
    Dim m_VerRetirados As String
    Dim m_VerDescripcion As String
    Dim m_DataLoader As RiskTreeDataLoader
    Dim m_PreloadWarnings As String
    
    On Error GoTo errores
    
    With m_Arbol
        .Style = 7 'tvwTreelinesPlusMinusText
        '.Style = 6 'tvwTreelinesPlusMinusText
        .MousePointer = ccArrow
        .LineStyle = 1 'tvwRootLines
        .LabelEdit = tvwManual
        .ImageList = m_ListImages
        '.BorderStyle = ccFixedSingle
        .BorderStyle = 0
        .Appearance = ccFlat
        .OLEDragMode = ccOLEDragManual
        .OLEDropMode = ccOLEDropNone
        .Indentation = 300
        .PathSeparator = "\"
        .HideSelection = False
        .sorted = False
        .FullRowSelect = True
        .Enabled = True
        .CheckBoxes = False
        .SingleSel = True
        .Scroll = True
        .HotTracking = False
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Refresh
        .Nodes.Clear
        
    End With
    If m_ObjEdicionActiva Is Nothing Then
        Me.FormDetalle.SourceObject = ""
        Exit Function
    End If
    
    
    
    m_VerRetirados = Nz(Me.ComboVerRetirados, "")
    If m_VerRetirados <> "Sí" And m_VerRetirados <> "No" Then
        m_VerRetirados = "No"
    End If
    m_VerDescripcion = Nz(Me.ComboVerDescripcion, "")
    If m_VerDescripcion <> "Sí" And m_VerDescripcion <> "No" Then
        m_VerDescripcion = "Sí"
    End If
    Set m_ColRiesgosAplicados = getColRiesgosAAPlicar(m_VerRetirados, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_Key = "EDICION|" & m_ObjEdicionActiva.IDEdicion
    m_NombreIcono = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaCerrada32)
    
    m_TituloNodo = m_ObjEdicionActiva.NombreParaNodo
    Set m_NodoEdicion = m_Arbol.Nodes.Add(, , m_Key, m_TituloNodo, m_NombreIcono)
    m_NodoEdicion.ExpandedImage = fso.GetFileName(m_ObjEntorno.URLIconoCarpetaCompletaAbierta32)
    If m_ColRiesgosAplicados Is Nothing Then
        Exit Function
    End If
    
    ' ============================================================
    ' Cache-aware preload: delegated to RiskTreeDataLoader
    ' (same non-fatal semantics: warnings only; rendering keeps lazy behavior).
    ' ============================================================
    Set m_DataLoader = New RiskTreeDataLoader
    Call m_DataLoader.PreloadForEdition(CStr(Nz(m_ObjEdicionActiva.IDEdicion, "")), m_ColRiesgosAplicados, m_PreloadWarnings)
    If Nz(m_PreloadWarnings, "") <> "" Then
        Debug.Print "CargarArbol preload warning(s): " & m_PreloadWarnings
    End If

    For Each m_IdRiesgo In m_ColRiesgosAplicados
        Set m_Riesgo = m_ColRiesgosAplicados(m_IdRiesgo)
       ' If m_Riesgo.CodigoRiesgo = "R003" Then Stop
        CargarArbolRiesgo p_NodoEdicion:=m_NodoEdicion, _
                                            p_Riesgo:=m_Riesgo, _
                                            p_borrarNodoSeleccionado:=EnumSiNo.Sí, _
                                            p_VerDescripcion:=Nz(ComboVerDescripcion, ""), _
                                            p_Refrescando:=p_Refrescando, p_Error:=p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Set m_Riesgo = Nothing
    Next
    m_NodoEdicion.Expanded = True
    SeleccionarNodo m_ObjEdicionActiva
    m_Arbol.Refresh
    CargarArbol = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CargarArbol ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function EliminarArbolNodos( _
                                    p_NodoSuperior As MSComctlLib.Node, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    Dim m_Nodo As MSComctlLib.Node
    
    
    On Error GoTo errores
    If p_NodoSuperior Is Nothing Then
        Exit Function
    End If
    Me.FormDetalle.SourceObject = ""
    m_Arbol.Nodes.Remove p_NodoSuperior.key
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EliminarArbolNodos ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function BorrarNodosHijos( _
                                    p_NodoPadre As MSComctlLib.Node, _
                                    Optional ByRef p_Error As String _
                                    ) As String
                                    
    Dim i As Long
    Dim NItems As Long
    Dim m_Nodo As MSComctlLib.Node
    
    Dim m_ColNodosABorrar As Scripting.Dictionary
    Dim varItem As Variant
    
    On Error GoTo errores
    
    
    
    If p_NodoPadre Is Nothing Then
        Exit Function
    End If
    NItems = p_NodoPadre.Children
    If NItems = 0 Then
        Exit Function
    End If
    Set m_Nodo = p_NodoPadre.child
    For i = 1 To NItems
        If m_ColNodosABorrar Is Nothing Then
            Set m_ColNodosABorrar = New Scripting.Dictionary
            m_ColNodosABorrar.CompareMode = TextCompare
        End If
        If Not m_ColNodosABorrar.Exists(m_Nodo.key) Then
            m_ColNodosABorrar.Add m_Nodo.key, m_Nodo.key
        End If
        Set m_Nodo = m_Nodo.Next
       
    Next
    Set m_Nodo = Nothing
    If m_ColNodosABorrar Is Nothing Then
        Exit Function
    End If
    For Each varItem In m_ColNodosABorrar
        m_Arbol.Nodes.Remove varItem
    Next
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método BorrarNodosHijos ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function
Public Function CargarArbolRiesgo( _
                                    p_NodoEdicion As MSComctlLib.Node, _
                                    p_Riesgo As riesgo, _
                                    Optional ByRef p_borrarNodoSeleccionado As EnumSiNo = EnumSiNo.Sí, _
                                    Optional ByRef p_VerDescripcion As String = "Sí", _
                                    Optional p_Refrescando As EnumSiNo = EnumSiNo.No, _
                                    Optional ByRef p_Error As String _
                                    ) As String

    
    
    
    Dim m_Key As String
    Dim m_NodoInterno As Object
    
    Dim m_NombreIcono As String
    Dim m_TituloNodo As String
    Dim m_NodoRiesgo As MSComctlLib.Node
    Dim m_ColPMs As Scripting.Dictionary
    Dim m_PM As pm
    Dim m_IDPM As Variant
    Dim m_ColPCs As Scripting.Dictionary
    Dim m_PC As pc
    Dim m_IDPC As Variant
    
    
    
    On Error GoTo errores
    
    If p_NodoEdicion Is Nothing Then
        p_Error = "No se conoce el Nodo edición"
        Err.Raise 1000
    End If
    If p_borrarNodoSeleccionado = Empty Then
        p_borrarNodoSeleccionado = EnumSiNo.Sí
    End If
    'If p_Riesgo.CodigoRiesgo = "R004" Then Stop
'   EDICION1 EDICION|IDEDICION
'       RIESGO1  RIESGO|IDRIESGO1
'           PM1     PM|IDRIESGO1|IDPM1
'               ACCIÓN1 PMACCION|IDPM1|IDACCION1
'               ACCIÓN2 PMACCION|IDPM1|IDACCION2
'           PM2     PM|IDRIESGO1|IDPM2
'               ACCIÓN1 PMACCION|IDPM2|IDACCION1
'       RIESGO2  RIESGO|IDRIESGO2
    m_Key = "RIESGO|" & p_Riesgo.IDRiesgo
    If p_borrarNodoSeleccionado = EnumSiNo.Sí Then
        On Error Resume Next
        m_Arbol.Nodes.Remove m_Key
        On Error GoTo errores
        Err.Clear
    Else
        'hemos de colocarnos en el nodo del riesgo y borrar los hijos
        On Error Resume Next
        Set m_NodoRiesgo = m_Arbol.Nodes(m_Key)
        m_NodoRiesgo.Selected = True
        BorrarNodosHijos m_NodoRiesgo, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    If p_Refrescando = EnumSiNo.Sí Then
        With p_Riesgo
            .NombreIcono = ""
            .NombreNodoDesc = ""
            .NombreNodoEstado = ""
            .ColorIcono = ""
            .HayErrorEnRiesgo = ""
        End With
    End If
    m_NombreIcono = getNombreIconoRiesgoArbolDeRiesgos(p_Riesgo, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If p_VerDescripcion = "Sí" Then
        If p_Riesgo.NombreNodoDesc = "" Then
            p_Riesgo.GrabarNombreNodoDesc p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        m_TituloNodo = p_Riesgo.NombreNodoDesc
    
       
    Else
        If p_Riesgo.NombreNodoEstado = "" Then
            p_Riesgo.GrabarNombreNodoEstado p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        m_TituloNodo = p_Riesgo.NombreNodoEstado
        
        
    End If
   
    If p_borrarNodoSeleccionado = EnumSiNo.Sí Then
        Set m_NodoRiesgo = m_Arbol.Nodes.Add(p_NodoEdicion.key, tvwChild, m_Key, m_TituloNodo, m_NombreIcono)
    Else
        m_NodoRiesgo.Text = m_TituloNodo
        
        m_NodoRiesgo.Image = m_NombreIcono
    End If
    Dim m_Color As String
    With p_Riesgo
        If .ColorIcono = "" Then
            .GrabarColorIcono
        End If
        m_Color = p_Riesgo.ColorIconoCalculado
        If m_Color = "Negro" Then
            m_NodoRiesgo.ForeColor = vbBlack
        ElseIf m_Color = "Rojo" Then
            m_NodoRiesgo.ForeColor = vbRed
        End If
        
    End With
    
    '-------------
    ' PMS
    '-----------------
    Set m_ColPMs = p_Riesgo.colPMs
    If Not m_ColPMs Is Nothing Then
        For Each m_IDPM In m_ColPMs
            Set m_PM = m_ColPMs(m_IDPM)
            CargarArbolPM P_NodoRiesgo:=m_NodoRiesgo, p_Riesgo:=p_Riesgo, _
                        p_PM:=m_PM, p_borrarNodoSeleccionado:=EnumSiNo.Sí, _
                        p_Refrescando:=p_Refrescando, p_Error:=p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        Next
    End If
    '-------------
    ' PCS
    '-----------------
    Set m_ColPCs = p_Riesgo.colPCs
    If Not m_ColPCs Is Nothing Then
        For Each m_IDPC In m_ColPCs
            Set m_PC = m_ColPCs(m_IDPC)
            'If p_Riesgo.CodigoRiesgo = "R016" Then Stop
            CargarArbolPC P_NodoRiesgo:=m_NodoRiesgo, p_Riesgo:=p_Riesgo, _
                        p_PC:=m_PC, p_borrarNodoSeleccionado:=EnumSiNo.Sí, _
                        p_Refrescando:=p_Refrescando, p_Error:=p_Error
        Next
    End If
        
       
    p_NodoEdicion.Expanded = True
    m_Arbol.Refresh
    CargarArbolRiesgo = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CargarArbolRiesgo ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function CargarArbolPM( _
                                P_NodoRiesgo As MSComctlLib.Node, _
                                p_Riesgo As riesgo, _
                                p_PM As pm, _
                                Optional ByRef p_borrarNodoSeleccionado As EnumSiNo = EnumSiNo.Sí, _
                                Optional p_Refrescando As EnumSiNo = EnumSiNo.No, _
                                Optional ByRef p_Error As String _
                                ) As String

    
    
    
        
    Dim m_NombreIcono As String
    Dim m_TituloNodo As String
    Dim m_Nodo As MSComctlLib.Node
    Dim m_Key As String
    Dim m_NodoPM As MSComctlLib.Node
   
    
        
    Dim m_ColPMAS As Scripting.Dictionary
    Dim m_PMA As PMAccion
    Dim m_IDPMA As Variant
    
    
    '   EDICION1 EDICION|IDEDICION
    '       RIESGO1  RIESGO|IDRIESGO1
    '           PM1     PM|IDRIESGO1|IDPM1
    '               ACCIÓN1 PMACCION|IDPM1|IDACCION1
    '               ACCIÓN2 PMACCION|IDPM1|IDACCION2
    '           PM2     PM|IDRIESGO1|IDPM2
    '               ACCIÓN1 PMACCION|IDPM2|IDACCION1
    '       RIESGO2  RIESGO|IDRIESGO2
     
    
    
    On Error GoTo errores
    If P_NodoRiesgo Is Nothing Then
        p_Error = "El nodo está vacío"
        Err.Raise 1000
    End If
    If p_Riesgo Is Nothing Then
        p_Error = "El riesgo está vacío"
        Err.Raise 1000
    End If
    If p_borrarNodoSeleccionado = Empty Then
        p_borrarNodoSeleccionado = EnumSiNo.Sí
    End If
    
   
   
    On Error Resume Next
    If Err.Number <> 0 Then
        On Error GoTo errores
        m_Error = "No se ha podido obtener el nodo del riesgo"
        Err.Raise 1000
    End If
    On Error GoTo errores
    m_TituloNodo = p_PM.CodMitigacion
    
    If p_Refrescando = EnumSiNo.Sí Then
        With p_PM
            .NombreIcono = ""
        End With
    End If
    If p_PM.NombreIcono = "" Then
        p_PM.GrabarNombreIcono p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    m_NombreIcono = p_PM.NombreIcono

    
    m_Key = "PM|" & p_Riesgo.IDRiesgo & "|" & p_PM.IDMitigacion
    If p_borrarNodoSeleccionado = EnumSiNo.Sí Then
        Set m_NodoPM = m_Arbol.Nodes.Add(P_NodoRiesgo.key, tvwChild, m_Key, m_TituloNodo, m_NombreIcono)
        ' Clear local var so next access re-reads from DB (batch pre-load is discarded on full tree refresh)
        Set m_ColPMAS = Nothing
    Else
        m_NodoPM.Text = m_TituloNodo
        m_NodoPM.Image = m_NombreIcono
    End If

    Set m_ColPMAS = p_PM.colAcciones
    If Not m_ColPMAS Is Nothing Then
        For Each m_IDPMA In m_ColPMAS
            Set m_PMA = m_ColPMAS(m_IDPMA)
            m_TituloNodo = m_PMA.CodAccion
            If p_Refrescando = EnumSiNo.Sí Or m_PMA.NombreIcono = "" Then
                m_PMA.NombreIcono = ""
                m_PMA.GrabarNombreIcono p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            End If
            
            m_NombreIcono = m_PMA.NombreIcono
           
            m_Key = "PMACCION|" & p_PM.IDMitigacion & "|" & m_PMA.IDAccionMitigacion
            Set m_Nodo = m_Arbol.Nodes.Add(m_NodoPM.key, tvwChild, m_Key, m_TituloNodo, m_NombreIcono)
            If m_NombreIcono = "Item_Completo_32x32.ico" Then
                m_Nodo.ForeColor = vbBlack
            Else
                m_Nodo.ForeColor = vbRed
            End If
            'm_Nodo.ForeColor = vbBlack
            Set m_PMA = Nothing
        Next
    End If
    m_NodoPM.Expanded = True
    m_Arbol.Refresh
    CargarArbolPM = "ok"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CargarArbolPM ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function
Public Function CargarArbolAccion( _
                                P_NodoPlan As MSComctlLib.Node, _
                                p_Accion As Object, _
                                Optional ByRef p_Error As String _
                                ) As String

    
    
    
        
    Dim m_NombreIcono As String
    Dim m_TituloNodo As String
    Dim m_Nodo As MSComctlLib.Node
    Dim m_Key As String
    
    
    '               ACCIÓN1 PMACCION|IDPM1|IDACCION1
     
    On Error GoTo errores
    If P_NodoPlan Is Nothing Then
        p_Error = "EL nodo está vacío"
        Err.Raise 1000
    End If
    If p_Accion Is Nothing Then
        p_Error = "La acción está vacía"
        Err.Raise 1000
    End If
    m_TituloNodo = p_Accion.CodAccion
    m_NombreIcono = getNombreIconoAccionArbolDeRiesgos(p_Accion, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    If TypeOf p_Accion Is PMAccion Then
        m_Key = "PMACCION|" & p_Accion.Mitigacion.IDMitigacion & "|" & p_Accion.IDAccionMitigacion
    ElseIf TypeOf p_Accion Is PCAccion Then
        m_Key = "PCACCION|" & p_Accion.Contingencia.IDContingencia & "|" & p_Accion.IDAccionContingencia
    Else
        p_Error = "No es acción de contingencia ni de mitigación"
        Err.Raise 1000
    End If
    Set m_Nodo = m_Arbol.Nodes.Add(P_NodoPlan.key, tvwChild, m_Key, m_TituloNodo, m_NombreIcono)
           
        
    
    
    m_Arbol.Refresh
    CargarArbolAccion = "ok"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CargarArbolAccion ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Public Function CargarArbolPC( _
                                P_NodoRiesgo As MSComctlLib.Node, _
                                p_Riesgo As riesgo, _
                                p_PC As pc, _
                                Optional ByRef p_borrarNodoSeleccionado As EnumSiNo = EnumSiNo.Sí, _
                                Optional p_Refrescando As EnumSiNo = EnumSiNo.No, _
                                Optional ByRef p_Error As String _
                                ) As String

    
    
    
        
    Dim m_NombreIcono As String
    Dim m_TituloNodo As String
    Dim m_Nodo As MSComctlLib.Node
    Dim m_Key As String
    Dim m_NodoPC As MSComctlLib.Node
   
    
        
    Dim m_ColPCAS As Scripting.Dictionary
    Dim m_PCA As PCAccion
    Dim m_IDPCA As Variant
    
    
    '   EDICION1 EDICION|IDEDICION
    '       RIESGO1  RIESGO|IDRIESGO1
    '           PC1     PC|IDRIESGO1|IDPC1
    '               ACCIÓN1 PCACCION|IDPC1|IDACCION1
    '               ACCIÓN2 PCACCION|IDPC1|IDACCION2
    '           PC2     PC|IDRIESGO1|IDPC2
    '               ACCIÓN1 PCACCION|IDPC2|IDACCION1
    '       RIESGO2  RIESGO|IDRIESGO2
     
    
    
    On Error GoTo errores
    If P_NodoRiesgo Is Nothing Then
        p_Error = "El nodo está vacío"
        Err.Raise 1000
    End If
    If p_Riesgo Is Nothing Then
        p_Error = "El riesgo está vacío"
        Err.Raise 1000
    End If
    If p_borrarNodoSeleccionado = Empty Then
        p_borrarNodoSeleccionado = EnumSiNo.Sí
    End If
    'If p_Riesgo.CodigoRiesgo = "R009" Then Stop
   
    
    On Error Resume Next
    If Err.Number <> 0 Then
        On Error GoTo errores
        m_Error = "No se ha podido obtener el nodo del riesgo"
        Err.Raise 1000
    End If
    On Error GoTo errores
    m_Key = "PC|" & p_PC.riesgo.IDRiesgo & "|" & p_PC.IDContingencia
    
    If p_borrarNodoSeleccionado = EnumSiNo.Sí Then
        On Error Resume Next
        m_Arbol.Nodes.Remove m_Key
        On Error GoTo errores
        Err.Clear
    Else
        'hemos de colocarnos en el nodo del riesgo y borrar los hijos
        On Error Resume Next
        Set m_NodoPC = m_Arbol.Nodes(m_Key)
        m_NodoPC.Selected = True
        BorrarNodosHijos m_NodoPC, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    m_TituloNodo = p_PC.CodContingencia
    If p_Refrescando = EnumSiNo.Sí Then
        With p_PC
            .GrabarNombreIcono p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End With
    End If
    If p_PC.NombreIcono = "" Then
        p_PC.GrabarNombreIcono p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
   
    m_NombreIcono = p_PC.NombreIcono
    
    m_Key = "PC|" & p_Riesgo.IDRiesgo & "|" & p_PC.IDContingencia
    If p_borrarNodoSeleccionado = EnumSiNo.Sí Then
        Set m_NodoPC = m_Arbol.Nodes.Add(P_NodoRiesgo.key, tvwChild, m_Key, m_TituloNodo, m_NombreIcono)
    Else
        m_NodoPC.Text = m_TituloNodo
        m_NodoPC.Image = m_NombreIcono
    End If

    Set m_ColPCAS = p_PC.colAcciones
    If Not m_ColPCAS Is Nothing Then
        For Each m_IDPCA In m_ColPCAS
            Set m_PCA = m_ColPCAS(m_IDPCA)
            m_TituloNodo = m_PCA.CodAccion
            If p_Refrescando = EnumSiNo.Sí Or m_PCA.NombreIcono = "" Then
                m_PCA.NombreIcono = ""
                m_PCA.GrabarNombreIcono p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            End If
            m_NombreIcono = m_PCA.NombreIcono
            
            m_Key = "PCACCION|" & p_PC.IDContingencia & "|" & m_PCA.IDAccionContingencia
            Set m_Nodo = m_Arbol.Nodes.Add(m_NodoPC.key, tvwChild, m_Key, m_TituloNodo, m_NombreIcono)
            
            Set m_PCA = Nothing
        Next
    End If
    m_NodoPC.Expanded = True
    m_Arbol.Refresh
    CargarArbolPC = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CargarArbolPC ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function
