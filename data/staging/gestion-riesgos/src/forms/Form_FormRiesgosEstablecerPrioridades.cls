Option Compare Database
Option Explicit
Private m_RiesgoSeleccionado As riesgo
Private blnHaHabidoPriorizaciones As Boolean
Private blnPermitidoEscribir As Boolean
Private lngNumeroRiesgosActivos As Long
Private m_ColPriorizacionAlInicio As Scripting.Dictionary
Private m_ColPriorizacionAlFinal As Scripting.Dictionary
Private m_ProyConRiesgosBiblioteca As EnumSiNo
Private m_Error As String

Private Sub ComandoActualizar_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    EstablecerLista EnumSiNo.Sí, m_Error
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
        m_Error = "Al ComandoActualizar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Function SeleccionarSiguiente(Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    Dim i As Long
    
    On Error GoTo errores
    
    Set lst = Me.ListaRiesgos
    If lst.ListCount < 1 Then
        Exit Function
    End If
    If m_RiesgoSeleccionado Is Nothing Then
        lst.Selected(1) = True
        ListaRiesgos_Click
        Exit Function
    End If
    
    For i = 1 To lst.ListCount - 1
        If lst.Column(0, i) = m_RiesgoSeleccionado.IDRiesgo Then
            If i = lst.ListCount - 1 Then
                lst.Selected(1) = True
            Else
                lst.Selected(i + 1) = True
            End If
            
            ListaRiesgos_Click
            Exit Function
        End If
        
    Next
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SeleccionarSiguiente ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Function SeleccionarAnterior(Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    Dim i As Long
    
    On Error GoTo errores
    
    Set lst = Me.ListaRiesgos
    If lst.ListCount < 1 Then
        Exit Function
    End If
    If m_RiesgoSeleccionado Is Nothing Then
        lst.Selected(1) = True
        ListaRiesgos_Click
        Exit Function
    End If
    
    For i = 1 To lst.ListCount - 1
        If lst.Column(0, i) = m_RiesgoSeleccionado.IDRiesgo Then
            If i = 1 Then
                lst.Selected(lst.ListCount - 1) = True
            Else
                lst.Selected(i - 1) = True
            End If
            
            ListaRiesgos_Click
            Exit Function
        End If
        
    Next
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SeleccionarAnterior ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Sub ComandoAsignar_Click()
    Dim lst As ListBox
    Dim m_Priorizacion As String
    
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Set lst = Me.ListaRiesgos
    If lst.ItemsSelected.Count = 0 Then
        m_Error = "Seleccione un elemento de la lista"
        Err.Raise 1000
    End If
    m_Priorizacion = Nz(Me.Priorizacion, "")
    If Not IsNumeric(m_Priorizacion) Then
        m_Error = "Se ha de incluir un número entre el 1 y el " & lngNumeroRiesgosActivos
        Err.Raise 1000
    End If
    If CLng(m_Priorizacion) < 1 Then
        m_Error = "Se ha de incluir un número entre el 1 y el " & lngNumeroRiesgosActivos
        Err.Raise 1000
    End If
    If CLng(m_Priorizacion) > lngNumeroRiesgosActivos Then
        m_Error = "Se ha de incluir un número entre el 1 y el " & lngNumeroRiesgosActivos
        Err.Raise 1000
    End If
    If Not Priorizacion_UpdateTempPriority(CLng(lst.Column(0)), m_Priorizacion, m_Error) Then
        Err.Raise 1000
    End If
    lst.Requery
    On Error Resume Next
    SeleccionarSiguiente
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAsignar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
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
        m_Error = "Al ComandoAyuda_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Sub ComandoGrabarPriorizacion_Click()
   
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    EstablecerPriorizaciones m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGrabarPriorizacion_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    
End Sub

Private Sub ComandoSalir_Click()
    Dim m_Error As String
    On Error Resume Next
    If blnHaHabidoPriorizaciones = True Then
        'RegistrarActualizacionRiesgo p_Edicion:=m_ObjEdicionActiva
        RegistrarActualizacionEdicion p_Edicion:=m_ObjEdicionActiva

        ' HR3d: cross-form CargarArbol -> Coord_CargarArbol (thin proxy).
        Call modFormCoordinationHelper.Coord_CargarArbol(m_Error)
        ' Error es tolerado: si FormRiesgosGestion no esta abierto,
        ' Coord_CargarArbol reporta p_Error legible sin raise.
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub Form_Load()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    AjustarTamaño Me
    EstablecerDatos m_Error
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
        m_Error = "Al Form_Load se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


Private Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
  

    On Error GoTo errores
    Me.ComandoGrabarPriorizacion.Enabled = False
    Me.ComandoAsignar.Enabled = False
    If m_ObjEdicionActiva Is Nothing Then
        p_Error = "No se conoce la edición activa"
        Err.Raise 1000
    End If
    If m_ObjEdicionActiva.EsActivo = EnumSiNo.Sí Then
        If m_ObjEdicionActiva.UsuarioConectadoAutorizado = EnumSiNo.Sí Then
            blnPermitidoEscribir = True
            Me.ComandoGrabarPriorizacion.Enabled = True
        Else
            blnPermitidoEscribir = False
        End If
    Else
        blnPermitidoEscribir = False
    End If
    If m_ObjEdicionActiva.colRiesgosActivos Is Nothing Then
        p_Error = "No hay riesgos no retirados"
        Err.Raise 1000
    End If
    lngNumeroRiesgosActivos = m_ObjEdicionActiva.colRiesgosActivos.Count
    m_ProyConRiesgosBiblioteca = m_ObjEdicionActiva.Proyecto.RequiereRiesgoDeBibliotecaCalculado
    Me.lblInfoOrdinal.Caption = "Introduzca un valor entre " & "1" & " " & "y" & " " & lngNumeroRiesgosActivos & " " & "para la priorización de este riesgo"
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    EstablecerLista EnumSiNo.No, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ColPriorizacionAlInicio = getColDeTabla(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function
Private Function EstablecerLista( _
                                Optional ByRef p_Actualizando As EnumSiNo = EnumSiNo.Sí, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim lst As ListBox
    
    Dim m_SQL As String
    On Error GoTo errores
    
    Set lst = Me.ListaRiesgos
    
    If m_ProyConRiesgosBiblioteca = EnumSiNo.Sí Then
        lst.ColumnCount = 6
        lst.ColumnWidths = "0cm;1,503cm;3cm;2cm;10cm;16cm"
        m_SQL = "SELECT IDRiesgo, Pri, PriEdAnterior, Codigo, Descripcion, CausaRaiz " & _
                "FROM TbAuxPriorizacion;"
        lst.RowSource = "IDRiesgo;Pri.;Pri.Ed.Anterior;Código;Descripción;CausaRaiz"
    Else
        lst.ColumnCount = 5
        lst.ColumnWidths = "0cm;1,503cm;3cm;2cm;25cm"
        m_SQL = "SELECT IDRiesgo, Pri, PriEdAnterior , Codigo, Descripcion " & _
                "FROM TbAuxPriorizacion;"
        
    End If
    If p_Actualizando = EnumSiNo.Sí Then
        Set m_ObjEdicionActiva = GetCachedEdicion(m_ObjEdicionActiva.IDEdicion)
    End If
    RellenarTbAux p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    lst.RowSource = m_SQL
    lst.Requery
    
    lst.Selected(1) = True
    ListaRiesgos_Click
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerLista ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function

Private Function getColDeTabla( _
                                Optional ByRef p_Error As String _
                                ) As Scripting.Dictionary
    
    On Error GoTo errores
    
    Set getColDeTabla = Priorizacion_ReadTempPriorities(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
   
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getColDeTabla ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function



Private Function SeleccionarRiesgo(p_IDRiesgo As String, Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    Dim i As Long
    
    On Error GoTo errores
    Set lst = Me.ListaRiesgos
    
    For i = 1 To lst.ListCount - 1
        If lst.Column(0, i) = p_IDRiesgo Then
            lst.Selected(i) = True
            ListaRiesgos_Click
            Exit Function
        End If
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SeleccionarRiesgo ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function

' HR4-sister (HR3d follow-up #9, #2 - 2026-07-01):
'   Era Public Function. Ahora Private. Body es thin adapter que delega
'   al helper modRiesgosEstablecerPrioridadesHelper.RiesgoPrioridad_VaciarTbAux.
'   1 internal caller in RellenarTbAux (Private) y el original habia sido
'   redundante (Priorizacion_ClearTemp ya es global). Hard rule 4.
Private Function VaciarTbAux(Optional ByRef p_Error As String) As String

    On Error GoTo errores

    ' HR4-sister: delega al helper.
    modRiesgosEstablecerPrioridadesHelper.RiesgoPrioridad_VaciarTbAux p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo VaciarTbAux ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If
End Function
' HR4-sister (HR3d follow-up #9, #2 - 2026-07-01):
'   Era Public Function. Ahora Private. Body es thin adapter que delega
'   al helper. 1 internal caller in EstablecerLista. Hard rule 4.
Private Function RellenarTbAux(Optional ByRef p_Error As String) As String

    On Error GoTo errores

    ' HR4-sister: delega al helper. El helper toma la coleccion + flag
    ' y hace el compute puro (vaciar + iterar + insertar). No toca Forms().
    modRiesgosEstablecerPrioridadesHelper.RiesgoPrioridad_RellenarTbAux _
        m_ObjEdicionActiva.colRiesgosActivos, m_ProyConRiesgosBiblioteca, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo RellenarTbAux ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If
End Function
Private Sub ListaRiesgos_Click()
    
    On Error GoTo errores

    m_Error = ""
    DoCmd.Hourglass True
    Me.Priorizacion = Null
    Me.ComandoAsignar.Enabled = False
    
    Me.Priorizacion = Null
    If Nz(Me.ListaRiesgos.Column(0), "") = "" Then
        Set m_RiesgoSeleccionado = Nothing
        DoCmd.Hourglass False
        Exit Sub
    End If
    If m_RiesgoSeleccionado Is Nothing Then
        Set m_RiesgoSeleccionado = GetCachedRiesgo(p_IDRiesgo:=CStr(Me.ListaRiesgos.Column(0)), p_Error:=m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    
    Else
        If m_RiesgoSeleccionado.IDRiesgo <> Me.ListaRiesgos.Column(0) Then
            Set m_RiesgoSeleccionado = GetCachedRiesgo(p_IDRiesgo:=CStr(Me.ListaRiesgos.Column(0)), p_Error:=m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    If m_RiesgoSeleccionado Is Nothing Then
        DoCmd.Hourglass False
        Exit Sub
    End If
    If IsNumeric(m_RiesgoSeleccionado.Priorizacion) Then
        Me.Priorizacion = m_RiesgoSeleccionado.Priorizacion
    Else
        If blnPermitidoEscribir = True Then
            If IsNumeric(m_RiesgoSeleccionado.PriorizacionEdAnterior) Then
                Me.Priorizacion = m_RiesgoSeleccionado.PriorizacionEdAnterior
            End If
        End If
    End If
    If blnPermitidoEscribir = True Then
        Me.ComandoAsignar.Enabled = True
        
    End If
   
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ListaRiesgos_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub

Private Function EstablecerPriorizaciones(Optional ByRef p_Error As String) As String
    
    Dim m_ID As Variant
    Dim m_Riesgo As riesgo
    Dim m_PriorizacionFinal As String
    Dim m_Service As RiesgoPriorizacionService
    Dim m_ChangedRiskIds As Scripting.Dictionary

    On Error GoTo errores
    If m_ColPriorizacionAlInicio Is Nothing Then
        p_Error = "No se estableció el listado inicial"
        Err.Raise 1000
    End If
    Set m_ColPriorizacionAlFinal = getColDeTabla(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    Set m_Service = New RiesgoPriorizacionService
    If Not m_Service.ValidatePriorities(m_ColPriorizacionAlFinal, lngNumeroRiesgosActivos, p_Error) Then
        Err.Raise 1000
    End If

    Set m_ChangedRiskIds = m_Service.GetChangedRiskIds( _
        m_ColPriorizacionAlInicio, m_ColPriorizacionAlFinal, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    'grabar priorizaciones
    For Each m_ID In m_ChangedRiskIds
        m_PriorizacionFinal = m_ChangedRiskIds(m_ID)
        Set m_Riesgo = GetCachedRiesgo(p_IDRiesgo:=CStr(m_ID), p_Error:=p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_Riesgo Is Nothing Then
            GoTo siguiente
        End If
        m_Riesgo.GrabarPrioridad m_PriorizacionFinal, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
siguiente:
    Next
    ' HR3d: cross-form ComandoActualizarContador_Click ->
    '   Coord_ComandoActualizarContadorRiesgosGestion (thin proxy).
    Call modFormCoordinationHelper.Coord_ComandoActualizarContadorRiesgosGestion(p_Error)
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerPriorizaciones ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Private Sub Priorizacion_KeyDown(KeyCode As Integer, Shift As Integer)
     If Me.ComandoAsignar.Enabled = False Then
        Exit Sub
    End If
    
    If KeyCode = vbKeyReturn Then
        Me.ComandoAsignar.SetFocus
        Me.Priorizacion.SetFocus
        'Debug.Print Me.Pass
        ' Código que deseas ejecutar cuando se pulse Enter
        ComandoAsignar_Click
        
        ' Si deseas evitar que Access realice la acción predeterminada de Enter,
        ' puedes utilizar la siguiente línea:
        KeyCode = 0
    ElseIf KeyCode = 40 Then
        SeleccionarSiguiente
    
    
    ElseIf KeyCode = 38 Then
        SeleccionarAnterior
    End If
End Sub
