Option Compare Database
Option Explicit
Dim strURLAnexo As String, strIDMaterialSeleccionado As String, strEsReparacionMaterialSeleccionado As String, strIDIntervencionSeleccionada As String, strNumeroSeguimientos As String
Dim strIDEquipoAnterior As String, strMaterialBusquedaAnterior As String, strEsReparacionAnterior As String, strEntregaAnterior As String, strPNBusquedaAnterior As String, _
    strNSBusquedaAnterior As String, strIDMaterialBusquedaAnterior As String, strBUIAnterior As String, strSubSistemaAnterior As String, strIrreparableAnterior As String
Public Sub ActualizarLista()
    Dim strTitulo As String, intElementos As Integer, lbl As Label, lst As ListBox, VarItem As Variant
    Dim strIDMaterialSeleccionadoActualmente As String, blnEncontrada As Boolean, i As Integer, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDMaterialSeleccionado = ""
    strEsReparacionMaterialSeleccionado = ""
    Me.ComandoEdicionMaterialSinActividad.Enabled = False
    Me.ComandoEliminar.Enabled = False
    
    Me.ComandoSeguimiento.Enabled = False
    Me.ComandoAdjuntos.Enabled = False
    strIDMaterialSeleccionadoActualmente = Nz(Me.ListaFiltrados.Column(0), "")
    '---------------------------------------------------
    '   -Devuelve:
    '       FiltrarMateriales=strSQLFinal
    '       FiltrarMateriales = "ERR" & "|" &strTextoError
    '---------------------------------------------------
    flag = FiltrarMateriales(Nz(Me.ComboBUIBusqueda, ""), Nz(Me.ComboSubsistemaBusqueda, ""), Nz(Me.ComboEquipoBusqueda.Column(3), ""), Nz(Me.ComboPNBusqueda, ""), Nz(Me.ComboNSBusqueda, ""), _
                            Nz(Me.ComboEsReparacion, ""), Nz(Me.ComboEntrega, ""), Nz(Me.ComboMaterialBusqueda, ""), _
                            Nz(Me.ComboIDMaterialBusqueda, ""), Nz(Me.ComboEquipoBusqueda, ""), Nz(Me.ComboIrreparable, ""))
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método FiltrarMateriales ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Set lst = Me.ListaFiltrados
    intElementos = lst.ListCount - 1
    If intElementos = -1 Then intElementos = 0
    If intElementos = 1 Then
        strIDMaterialSeleccionadoActualmente = lst.Column(0, 1)
        lst.Selected(1) = True
        ListaFiltrados_Click
    Else
        For i = 1 To lst.ListCount - 1
            If lst.Selected(i) = True Then
                strIDMaterialSeleccionadoActualmente = lst.Column(0, i)
                ListaFiltrados_Click
                Exit For
            End If
        Next
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar actualizar la lista  se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Public Sub ComandoActualizarTodo_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    ComandoLimpiarEquipo_Click
    ComandoLimpiarPN_Click
    ComandoLimpiarNS_Click
    ComandoLimpiarEsReparacion_Click
    ComandoLimpiarEntrega_Click
    ComandoLimpiarMaterial_Click
    ComandoLimpiarIDMaterial_Click
    ComandoLimpiarBUI_Click
    ComandoLimpiarSubSistema_Click
    ComandoLimpiarEquipo_Click
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar actualizar las listas se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub ComandoAdjuntos_Click()
    Dim m_Material As Material
    Dim strFormularioQueAbre As String
    Dim strTextoError As String, strIDMaterial As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDMaterial = Nz(Me.ListaFiltrados.Column(0), "")
    If strIDMaterial = "" Then
        strTextoError = "Se ha de Seleccionar un material primero"
        Err.Raise 1000
    End If
    Set m_Material = Constructor.getMaterial(strIDMaterial, strTextoError)
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    If FormularioAbierto("FormAnexos") Then
        DoCmd.Close acForm, "FormAnexos", acSaveNo
    End If
    Set m_ObjEntidadParaAnexoActiva = m_Material
    'strArgumento=strFormularioQueAbre
    DoCmd.OpenForm "FormAnexos", , , , , , strFormularioQueAbre
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver los adjuntos se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub ComandoEdicionMaterialSinActividad_Click()
    Dim strTextoError As String, strFranqueado As String, strEstaFacturado As String, strURLAnexo As String, strNombreArchivo As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDMaterialSeleccionado = Nz(Me.ListaFiltrados.Column(0), "")
    strEsReparacionMaterialSeleccionado = Nz(Me.ListaFiltrados.Column(6), "")
    If strIDMaterialSeleccionado = "" Then
        strTextoError = "No se ha seleccionado ningún material que editar"
        Err.Raise 1000
    End If
    If strEsReparacionMaterialSeleccionado <> "Sí" And strEsReparacionMaterialSeleccionado <> "No" Then
        strTextoError = "No se ha podido determinar si es un material para reparar o no."
        Err.Raise 1000
    End If
    'strArgumento=strIDMaterial
    If strEsReparacionMaterialSeleccionado = "Sí" Then
        '-------------------------------------------------------------------
        '   -Devuelve
        '       FormularioAbierto = true or false
        '       FormularioAbierto =False-->Error
        '------------------------------------------------
        If FormularioAbierto("FormMaterialReparacionEdicion") Then
            DoCmd.Close acForm, "FormMaterialReparacionEdicion", acSaveNo
        End If
        'strArgumento=strIDMaterial  & "|" & strFranqueado & "|" & strEstaFacturado & "|" & strURLAnexo
        'strArgumento=strIDMaterial
        DoCmd.OpenForm "FormMaterialReparacionEdicion", , , , , , strIDMaterialSeleccionado
    Else
        
        strEstaFacturado = Nz(Me.ListaFiltrados.Column(8), "")
        If strEstaFacturado <> "Sí" And strEstaFacturado <> "No" Then
            flag = EstaFacturado("MA", strIDMaterialSeleccionado)
            If InStr(1, flag, "|") <> 0 Then
                dato = Split(flag, "|")
                strTextoError = "El método EstaFranqueado ha devuelto el error: " & vbNewLine & dato(1)
                Err.Raise 1000
            End If
            strFranqueado = flag
        End If
        '-----------------------------------------------
        '   -Devuelve:
        '       strPrefijo=EV,AC,MA,EQ,
        '       EstaFranqueado = strValorFranqueado
        '       EstaFranqueado = "#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = EstaFranqueado("MA", strIDMaterialSeleccionado)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método EstaFranqueado ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strFranqueado = flag
        strNombreArchivo = Nz(Me.ListaFiltrados.Column(9), "")
        If strNombreArchivo <> "" Then
            strURLAnexo = m_ObjEntorno.URLDirectorioDocumentacion & strNombreArchivo
        End If
         '-------------------------------------------------------------------
        '   -Devuelve
        '       FormularioAbierto = true or false
        '       FormularioAbierto =False-->Error
        '------------------------------------------------
        If FormularioAbierto("FormMaterialRepuestoEdicion") Then
            DoCmd.Close acForm, "FormMaterialRepuestoEdicion", acSaveNo
        End If
        'strArgumento=strIDMaterial  & "|" & strFranqueado & "|" & strEstaFacturado  & "|" & strURLAnexo
        DoCmd.OpenForm "FormMaterialRepuestoEdicion", , , , , , strIDMaterialSeleccionado & "|" & strFranqueado & "|" & strEstaFacturado & "|" & strURLAnexo
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar editar un material se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoEliminar_Click()
    Dim strIDEvento As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDMaterialSeleccionado = Nz(Me.ListaFiltrados.Column(0), "")
    strIDEvento = Nz(Me.ListaFiltrados.Column(5), "")
    If strIDMaterialSeleccionado = "" Then
        strTextoError = "No se ha seleccionado ningún material que eliminar"
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("¿Desea eliminar el Material seleccionado: " & strIDMaterialSeleccionado & " ?", _
            vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar material")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If strIDEvento = "" Then
        '-----------------------------------------------
        '   -Devuelve:
        '       EliminarMaterialSinActividad =Descriptivo
        '       EliminarMaterialSinActividad ="#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = EliminarMaterialSinActividad(strIDMaterialSeleccionado, "No")
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método EliminarMaterialSinActividad ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        '-----------------------------------------------
        '   -Devuelve:
        '       EliminarMaterialSinActividad =Descriptivo
        '       EliminarMaterialSinActividad ="#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = EliminarMaterialSinActividad(strIDMaterialSeleccionado, "Sí")
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método EliminarMaterialSinActividad ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
    Else
        '-----------------------------------------------
        '   -Devuelve:
        '       EliminarMaterialEnActividad = Descriptivo
        '       EliminarMaterialEnActividad = "#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = EliminarMaterialEnActividad(strIDMaterialSeleccionado, "No")
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método EliminarMaterialEnActividad ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        '-----------------------------------------------
        '   -Devuelve:
        '       EliminarMaterialEnActividad = Descriptivo
        '       EliminarMaterialEnActividad = "#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = EliminarMaterialEnActividad(strIDMaterialSeleccionado, "Sí")
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método EliminarMaterialEnActividad ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
    End If
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar eliminar un material se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub ComandoExportarExcel_Click()
    Dim m_Error As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    flag = GenerarConsultas("Lista de Materiales", Me.hWnd, , Me.ListaFiltrados, "No")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        m_Error = "El método GenerarConsultas ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoExportarExcel_Click se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
End Sub

Private Sub ComandoLimpiarBUI_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboBUIBusqueda, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboBUIBusqueda = Null
    ActualizarLista
    EstablecerCombo Me.ComboSubsistemaBusqueda
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub

Private Sub ComandoLimpiarEntrega_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboEntrega, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboEntrega = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComandoLimpiarEquipo_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboEquipoBusqueda, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboEquipoBusqueda = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComandoLimpiarEsReparacion_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboIrreparable, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboIrreparable = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComandoLimpiarIDMaterial_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboIDMaterialBusqueda, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboIDMaterialBusqueda = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComandoLimpiarIrreparable_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboIrreparable, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboIrreparable = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub

Private Sub ComandoLimpiarMaterial_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboMaterialBusqueda, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboMaterialBusqueda = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComandoLimpiarNS_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboNSBusqueda, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboNSBusqueda = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComandoLimpiarPN_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboPNBusqueda, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboPNBusqueda = Null
    ActualizarLista
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub

Private Sub ComandoLimpiarSubSistema_Click()
    On Error Resume Next
    LimpiarUltimos
    If Nz(Me.ComboSubsistemaBusqueda, "") = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.ComboSubsistemaBusqueda = Null
    ActualizarLista
    EstablecerCombo Me.ComboEquipoBusqueda
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub ComandoSeguimiento_Click()
    Dim strIDEvento As String, strArgumento As String, strTextoError As String
    Dim strFacturado As String, strEventoFranqueado As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDMaterialSeleccionado = Nz(Me.ListaFiltrados.Column(0), "")
    strIDEvento = Nz(Me.ListaFiltrados.Column(5), "")
    If strIDMaterialSeleccionado = "" Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If Not IsNumeric(strNumeroSeguimientos) Then
        strTextoError = "No se sabe cuántos seguimientos tiene el material"
        Err.Raise 1000
    End If
    
    strFacturado = Nz(Me.ListaFiltrados.Column(7), "")
    If strIDEvento <> "" Then
        '-----------------------------------------------
        '   -Devuelve:
        '       strPrefijo=EV,AC,MA,EQ,
        '       EstaFranqueado = strValorFranqueado
        '       EstaFranqueado = "#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = EstaFranqueado("EV", strIDEvento)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método EstaFranqueado ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strEventoFranqueado = flag
    End If
    '-------------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '-------------------------------------------------
    If FormularioAbierto("FormMaterialSeguimiento") Then
        DoCmd.Close acForm, "FormMaterialSeguimiento", acSaveNo
    End If
    'strID & "|" & strNumeroIntervencionesEnSeguimiento & "|" & strIDEvento
    If CInt(strNumeroSeguimientos) = 0 Then
        If strFacturado = "Sí" Or strEventoFranqueado = "Sí" Then
            strTextoError = "No tiene seguimientos y no se puede generar el primero al estar Facturado o su evento Franqueado"
            Err.Raise 1000
        End If
        '-------------------------------------------------
        '   -Devuelve
        '       FormularioAbierto = true or false
        '       FormularioAbierto =False-->Error
        '-------------------------------------------------
        If FormularioAbierto("FormMaterialIntervencionAlta") Then
            DoCmd.Close acForm, "FormMaterialIntervencionAlta"
        End If
        DoCmd.OpenForm "FormMaterialIntervencionAlta", , , , , , strIDMaterialSeleccionado
    Else
        
        
        strArgumento = strIDMaterialSeleccionado & "|" & strNumeroSeguimientos & "|" & strIDEvento & "|" & strFacturado & "|" & strEventoFranqueado
        DoCmd.OpenForm "FormMaterialSeguimiento", , , , , , strArgumento
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver el seguimiendo del material se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub


Private Sub ComboBUIBusqueda_Change()
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboBUIBusqueda, "")
    If strValor <> strBUIAnterior Then
        EstablecerCombo Me.ComboSubsistemaBusqueda
        ActualizarLista
    End If
    strBUIAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub ComboBUIBusqueda_Exit(Cancel As Integer)
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboBUIBusqueda, "")
    If strValor <> strBUIAnterior Then
        EstablecerCombo Me.ComboSubsistemaBusqueda
        ActualizarLista
    End If
    strBUIAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub ComboEntrega_Change()
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboEntrega, "")
    If strValor <> strEntregaAnterior Then
        ActualizarLista
    End If
    strEntregaAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComboEntrega_Exit(Cancel As Integer)
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboEntrega, "")
    If strValor <> strEntregaAnterior Then
        ActualizarLista
    End If
    strEntregaAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComboEsReparacion_Change()
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboEsReparacion, "")
    If strValor <> strEsReparacionAnterior Then
        ActualizarLista
    End If
    strEsReparacionAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComboEsReparacion_Exit(Cancel As Integer)
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboEsReparacion, "")
    If strValor <> strEsReparacionAnterior Then
        ActualizarLista
    End If
    strEsReparacionAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComboEquipoBusqueda_Change()
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboEquipoBusqueda.Column(0), "")
    If strValor <> strIDEquipoAnterior Then
        ActualizarLista
    End If
    strIDEquipoAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComboEquipoBusqueda_Exit(Cancel As Integer)
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboEquipoBusqueda.Column(0), "")
    If strValor <> strIDEquipoAnterior Then
        ActualizarLista
    End If
    strIDEquipoAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Public Sub ComboIDMaterialBusqueda_Change()
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboIDMaterialBusqueda, "")
    If strValor <> strIDMaterialBusquedaAnterior Then
        ActualizarLista
    End If
    strIDMaterialBusquedaAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComboIDMaterialBusqueda_Exit(Cancel As Integer)
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboIDMaterialBusqueda, "")
    If strValor <> strIDMaterialBusquedaAnterior Then
        ActualizarLista
    End If
    strIDMaterialBusquedaAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub ComboIrreparable_Change()
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboIrreparable, "")
    If strValor <> strIrreparableAnterior Then
        ActualizarLista
    End If
    strIrreparableAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub ComboIrreparable_Exit(Cancel As Integer)
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboIrreparable, "")
    If strValor <> strIrreparableAnterior Then
        ActualizarLista
    End If
    strIrreparableAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub ComboMaterialBusqueda_Change()
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboMaterialBusqueda, "")
    If strValor <> strMaterialBusquedaAnterior Then
        ActualizarLista
    End If
    strMaterialBusquedaAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComboMaterialBusqueda_Exit(Cancel As Integer)
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboMaterialBusqueda, "")
    If strValor <> strMaterialBusquedaAnterior Then
        ActualizarLista
    End If
    strMaterialBusquedaAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComboNSBusqueda_Change()
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboNSBusqueda, "")
    If strValor <> strNSBusquedaAnterior Then
        ActualizarLista
    End If
    strNSBusquedaAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComboNSBusqueda_Exit(Cancel As Integer)
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboNSBusqueda, "")
    If strValor <> strNSBusquedaAnterior Then
        ActualizarLista
    End If
    strNSBusquedaAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComboPNBusqueda_Change()
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboPNBusqueda, "")
    If strValor <> strPNBusquedaAnterior Then
        ActualizarLista
    End If
    strPNBusquedaAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComboPNBusqueda_Exit(Cancel As Integer)
     Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboPNBusqueda, "")
    If strValor <> strPNBusquedaAnterior Then
        ActualizarLista
    End If
    strPNBusquedaAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComboSubsistemaBusqueda_Change()
     Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboSubsistemaBusqueda, "")
    If strValor <> strSubSistemaAnterior Then
        EstablecerCombo Me.ComboEquipoBusqueda
        ActualizarLista
    End If
    strSubSistemaAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComboSubsistemaBusqueda_Exit(Cancel As Integer)
    Dim strTextoError As String, strValor As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValor = Nz(Me.ComboSubsistemaBusqueda, "")
    If strValor <> strSubSistemaAnterior Then
        EstablecerCombo Me.ComboEquipoBusqueda
        ActualizarLista
    End If
    strSubSistemaAnterior = strValor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al aplicar el filtro se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub Form_Open(Canel As Integer)
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.Caption = Me.lblTitulo1.Caption
    AjustarTamaño Me
    
    ActualizarLista
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub
Private Sub cmdSalir_Click()
   On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub ListaFiltrados_Click()
    Dim strIDEventoSeleccionado As String, strFacturado As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strNumeroSeguimientos = ""
    Me.ComandoAdjuntos.Enabled = False
    Me.ComandoEdicionMaterialSinActividad.Enabled = False
    Me.ComandoEliminar.Enabled = False
    
    Me.ComandoAdjuntos.Enabled = False
    strIDMaterialSeleccionado = Nz(Me.ListaFiltrados.Column(0), "")
    strEsReparacionMaterialSeleccionado = Nz(Me.ListaFiltrados.Column(6), "")
    strIDEventoSeleccionado = Nz(Me.ListaFiltrados.Column(5), "")
    strFacturado = Nz(Me.ListaFiltrados.Column(7), "")
    
    If strIDMaterialSeleccionado = "" Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Me.ComandoAdjuntos.Enabled = True
    Me.ComandoSeguimiento.Enabled = True
    If strIDMaterialSeleccionado = "" Then
        Me.ComandoEdicionMaterialSinActividad.Caption = "Editar"
        Me.ComandoEdicionMaterialSinActividad.Enabled = False
        
        Me.ComandoEliminar.Enabled = False
    Else
        If strFacturado = "Sí" Then
            Me.ComandoEdicionMaterialSinActividad.Enabled = True
            Me.ComandoEdicionMaterialSinActividad.Caption = "Ver"
            Me.ComandoEliminar.Enabled = False
            
        Else
            Me.ComandoEdicionMaterialSinActividad.Enabled = True
            Me.ComandoEdicionMaterialSinActividad.Caption = "Editar"
            Me.ComandoEliminar.Enabled = True
            
        End If
    End If
    '-------------------------------------------------------------------------
    '   -Devuelve:
    '       IntervencionNumeroPorMaterial = strNumero & ";" & strIDIntervencion (si es sólo una)
    '       IntervencionNumeroPorMaterial = "#ERR"  & "|" & strTextoError
    '-------------------------------------------------------------------------
    flag = IntervencionNumeroPorMaterial(strIDMaterialSeleccionado)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método IntervencionNumeroPorMaterial ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If InStr(1, flag, ";") <> 0 Then
        dato = Split(flag, ";")
        strNumeroSeguimientos = dato(0)
        strIDIntervencionSeleccionada = dato(1)
    Else
        strNumeroSeguimientos = flag
    End If
    Me.ComandoSeguimiento.Caption = "Seguimiento(" & strNumeroSeguimientos & ")"
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar Seleccionar el material se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Public Function FiltrarMateriales( _
                                Optional strBUI As String, _
                                Optional strSUBSISTEMA As String, _
                                Optional strIDEquipo As String, _
                                Optional strPN As String, _
                                Optional strNS As String, _
                                Optional strEsReparacion As String, _
                                Optional strEntregado As String, _
                                Optional strMaterial As String, _
                                Optional strIDMaterial As String, _
                                Optional strEquipo As String, _
                                Optional strIrreparable As String _
                                ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 30/09/2019
    '   -Modificaciones:
    
    '   -Funcionamiento:
    
    '   -Llamada desde:
    
    '   -Devuelve:
    '       FiltrarMateriales=strSQLFinal
    '       FiltrarMateriales = "ERR" & "|" &strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, stIDEquipoENBBDD As String, strIDEvento As String, strBUIConsulta As String, strSUBSISTEMAConsulta As String
    Dim strWherePN As String, strWhereNS As String, strWhereEsReparacion As String, strFacturado As String, _
        strWhereEntregado As String, strWhereMaterial As String, strWhereIDMaterial As String, strWhereEquipoDeEvento As String, _
        strWhereEquipoDeMaterial As String, strWhereIrreparable As String, lst As ListBox
    Dim strSQLInicial As String, strSQLFinal As String, strParteWhere As String, strTextoError As String
    On Error GoTo errores
    Set lst = Me.ListaFiltrados
    lst.RowSource = "IDMaterial;Material;Inicio;FEntrega;Equipo;IDEvento;Rep.;NS;Fact"
    If strBUI = "" And strSUBSISTEMA = "" And strIDEquipo = "" And strPN = "" And strNS = "" And strEsReparacion = "" And strEntregado = "" _
        And strMaterial = "" And strIDMaterial = "" And strEquipo = "" And strIrreparable = "" Then
        Exit Function
    End If
    Me.ComandoExportarExcel.Enabled = False
    strSQLInicial = "SELECT DISTINCT TbMaterial.IDMaterial, TbMaterial.Material, TbMaterial.FechaIncio, TbMaterial.FechaEntrega, TbEquipos_1.Equipo AS EquipoDeMat1, " & _
                    "TbMaterial.EsReparacion, TbMaterial.TipoAccion, TbEquipos.Equipo AS EquipoDeEvento, TbEquipos_1.IDEquipo AS EquipoDeMat, " & _
                    "TbEquipos.IDEquipo AS IDEquipoDeEvento, TbEventos.IDEvento, TbEventos.BUI AS BUIEVENTO, TbEventos.SUBSISTEMA AS SUBSISTEMAEVENTO, " & _
                    "TbEquipos_1.BUI AS BUIEQUIPO, TbEquipos_1.SUBSISTEMA AS SUBSISTEMAEQUIPO,TbMaterial.NS, TbFacturaMaterialesInvolucrados.IDFactura " & _
                    "FROM (((TbEventos RIGHT JOIN (TbMaterial LEFT JOIN TbActividades ON TbMaterial.IDActividad = TbActividades.IDActividad) ON " & _
                    "TbEventos.IDEvento = TbActividades.IDEvento) LEFT JOIN TbEquipos ON TbEventos.IDEquipo = TbEquipos.IDEquipo) " & _
                    "LEFT JOIN TbEquipos AS TbEquipos_1 ON TbMaterial.IDEquipo = TbEquipos_1.IDEquipo) LEFT JOIN TbFacturaMaterialesInvolucrados ON " & _
                    "TbMaterial.IDMaterial = TbFacturaMaterialesInvolucrados.IDMaterial "
    If strPN = "" Then
        strWherePN = "((TbMaterial.PN) Like '*' Or (TbMaterial.PN) Is Null) "
    Else
        strWherePN = "((TbMaterial.PN)='" & strPN & "')"
    End If
    If strNS = "" Then
        strWhereNS = "((TbMaterial.NS) Like '*' Or (TbMaterial.NS) Is Null) "
    Else
        strWhereNS = "((TbMaterial.NS)='" & strNS & "')"
    End If
    If strEsReparacion <> "Sí" And strEsReparacion <> "No" Then
        strWhereEsReparacion = "((TbMaterial.EsReparacion) Like '*' Or (TbMaterial.EsReparacion) Is Null) "
    Else
        strWhereEsReparacion = "((TbMaterial.EsReparacion)='" & strEsReparacion & "')"
    End If
    If strEntregado <> "Sí" And strEntregado <> "No" Then
        strWhereEntregado = "((TbMaterial.FechaEntrega) Like '*' Or (TbMaterial.FechaEntrega) Is Null) "
    Else
        strWhereEntregado = "((TbMaterial.FechaEntrega)='" & strEntregado & "')"
    End If
    If strMaterial = "" Then
        strWhereMaterial = "((TbMaterial.Material) Like '*' Or (TbMaterial.Material) Is Null) "
    Else
        strWhereMaterial = "((TbMaterial.Material)='" & strMaterial & "')"
    End If
    If strIDMaterial = "" Then
        strWhereIDMaterial = "((TbMaterial.IDMaterial) Like '*' Or (TbMaterial.IDMaterial) Is Null) "
    Else
        strWhereIDMaterial = "((TbMaterial.IDMaterial)='" & strIDMaterial & "')"
    End If
    If strIrreparable <> "Sí" And strIrreparable <> "No" Then
        strWhereIrreparable = "((TbMaterial.TipoAccion) Like '*' Or (TbMaterial.TipoAccion) Is Null) "
    Else
        If strIrreparable = "Sí" Then
            strWhereIrreparable = "((TbMaterial.TipoAccion)='IDENTIFICADO_COMO_IRREPARABLE_OBSOLETO')"
        Else
            strWhereIrreparable = "((TbMaterial.TipoAccion) Like '*' Or (TbMaterial.TipoAccion) Is Null) "
        End If
    End If
    If strEquipo = "" Then
        strParteWhere = "WHERE (" & _
                        strWherePN & " AND " & _
                        strWhereNS & " AND " & _
                        strWhereEsReparacion & " AND " & _
                        strWhereEntregado & " AND " & _
                        strWhereMaterial & " AND " & _
                        strWhereIrreparable & " AND " & _
                        strWhereIDMaterial & ");"
    Else
        strWhereEquipoDeEvento = "((TbEquipos.Equipo)='" & strEquipo & "') "
        strWhereEquipoDeMaterial = "((TbEquipos_1.Equipo)='" & strEquipo & "') "
        strParteWhere = "WHERE ((" & _
                            strWherePN & " AND " & _
                            strWhereNS & " AND " & _
                            strWhereEsReparacion & " AND " & _
                            strWhereEntregado & " AND " & _
                            strWhereMaterial & " AND " & _
                            strWhereIrreparable & " AND " & _
                            strWhereEquipoDeEvento & " AND " & _
                            strWhereIDMaterial & ") " & _
                        "OR (" & _
                            strWherePN & " AND " & _
                            strWhereNS & " AND " & _
                            strWhereEsReparacion & " AND " & _
                            strWhereEntregado & " AND " & _
                            strWhereMaterial & " AND " & _
                            strWhereIrreparable & " AND " & _
                            strWhereEquipoDeMaterial & " AND " & _
                            strWhereIDMaterial & "));"
    End If
    strSQLFinal = strSQLInicial & strParteWhere
    Set rcdDatos = CurrentDb().OpenRecordset(strSQLFinal)
    With rcdDatos
        If Not .EOF Then
            Me.ComandoExportarExcel.Enabled = True
            .MoveFirst
            Do While Not .EOF
                strBUIConsulta = ""
                strSUBSISTEMAConsulta = ""
                strIDEvento = Nz(.Fields("IDEvento"), "")
                strEquipo = Nz(.Fields("EquipoDeMat1"), "")
                If strEquipo = "" Then
                    strEquipo = Nz(.Fields("EquipoDeEvento"), "")
                End If
                If IsNumeric(strIDEquipo) Then
                    stIDEquipoENBBDD = Nz(.Fields("EquipoDeMat"), "")
                    If stIDEquipoENBBDD = "" Then
                        stIDEquipoENBBDD = Nz(.Fields("IDEquipoDeEvento"), "")
                    End If
                    If strIDEquipo <> stIDEquipoENBBDD Then
                        GoTo siguiente
                    End If
                End If
                If strBUI <> "" Then
                    strBUIConsulta = Nz(.Fields("BUIEVENTO"), "")
                    If strBUIConsulta = "" Then
                        strBUIConsulta = Nz(.Fields("BUIEQUIPO"), "")
                    End If
                    If strBUIConsulta <> strBUI Then
                        GoTo siguiente
                    End If
                End If
                If strSUBSISTEMA <> "" Then
                    strSUBSISTEMAConsulta = Nz(.Fields("SUBSISTEMAEVENTO"), "")
                    If strSUBSISTEMAConsulta = "" Then
                        strSUBSISTEMAConsulta = Nz(.Fields("SUBSISTEMAEQUIPO"), "")
                    End If
                    If strSUBSISTEMAConsulta <> strSUBSISTEMA Then
                        GoTo siguiente
                    End If
                End If
                strFacturado = Nz(.Fields("IDFactura"), "")
                If IsNumeric(strFacturado) Then
                    strFacturado = "Sí"
                Else
                    strFacturado = "No"
                End If
                
               ' If Nz(.Fields("Material"), "") = "" Then Stop
                lst.AddItem Nz(.Fields("IDMaterial"), "") & ";" & Nz(.Fields("Material"), "") & ";" & Nz(.Fields("FechaIncio"), "") _
                            & ";" & Nz(.Fields("FechaEntrega"), "") & ";" & strEquipo & ";" & strIDEvento & ";" & Nz(.Fields("EsReparacion"), "") & _
                            ";" & Nz(.Fields("NS"), "") & ";" & strFacturado
siguiente:
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    FiltrarMateriales = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método FiltrarMateriales ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    FiltrarMateriales = "#ERR" & "|" & strTextoError
End Function
Public Sub LimpiarUltimos()
    strIDEquipoAnterior = ""
    strMaterialBusquedaAnterior = ""
    strIDMaterialBusquedaAnterior = ""
    strEsReparacionAnterior = ""
    strEntregaAnterior = ""
    strPNBusquedaAnterior = ""
    strNSBusquedaAnterior = ""
    strBUIAnterior = ""
    strSubSistemaAnterior = ""
End Sub
Public Function IntervencionNumeroPorMaterial(strIDMaterial As String) As String
    '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 03/10/14
    '   -Modificaciones:
    
    '   -Funcionamiento:
   
    '   -Llamada desde
    
    '   -Devuelve:
    '       IntervencionNumeroPorMaterial = cstr(lngNumero) & ";" & strIDIntervencion (si es sólo una)
    '       IntervencionNumeroPorMaterial = "#ERR"  & "|" & strTextoError
    '-------------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, lngNumero As Long, strIDIntervencion As String, strTextoError As String, strResultado As String
    On Error GoTo errores
    If strIDMaterial = "" Then
        strTextoError = "Se ha de indicar el IDMaterial"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbMaterialSeguimiento.IDSeguimiento " & _
                "FROM TbMaterialSeguimiento " & _
                "WHERE (((TbMaterialSeguimiento.IDMaterial)='" & strIDMaterial & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveLast
            .MoveFirst
            strIDIntervencion = .Fields("IDSeguimiento")
            lngNumero = .RecordCount
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If lngNumero = 1 Then
       strResultado = lngNumero & ";" & strIDIntervencion
    Else
        strResultado = lngNumero
    End If
    IntervencionNumeroPorMaterial = strResultado
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Material.IntervencionNumeroPorMaterial ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
     IntervencionNumeroPorMaterial = "#ERR" & "|" & strTextoError
End Function
Public Sub EstablecerCombo(cbo As ComboBox)
    Dim strSQLCombo As String, strBUI As String, strSUBSISTEMA As String, strIDEquipo As String, _
        strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If cbo.Name = Me.ComboSubsistemaBusqueda.Name Then
        strBUI = Nz(Me.ComboBUIBusqueda, "")
        If strBUI = "" Then
            strBUI = "-1"
        End If
        cbo.Value = Null
        Me.ComboSubsistemaBusqueda = Null
        Me.ComboEquipoBusqueda = Null
        strSQLCombo = "SELECT TbSubsistemaBui.SUBSISTEMA " & _
                    "FROM TbSubsistemaBui " & _
                    "WHERE (((TbSubsistemaBui.BUI)='" & strBUI & "')) " & _
                    "ORDER BY TbSubsistemaBui.SUBSISTEMA;"
        cbo.RowSource = strSQLCombo
        cbo.Requery
        If cbo.ListCount = 1 Then
            cbo.Value = cbo.Column(0, 0)
        End If
    ElseIf cbo.Name = Me.ComboEquipoBusqueda.Name Then
        strBUI = Nz(Me.ComboBUIBusqueda, "")
        strSUBSISTEMA = Nz(Me.ComboSubsistemaBusqueda, "")
        If strBUI = "" Then
            strBUI = "-1"
        End If
        If strSUBSISTEMA = "" Then
            strSUBSISTEMA = "-1"
        End If
        cbo.Value = Null
        Me.ComboEquipoBusqueda = Null
        strSQLCombo = "SELECT TbEquipos.Equipo " & _
                        "FROM TbEquipos " & _
                        "WHERE (((TbEquipos.SUBSISTEMA)='" & strSUBSISTEMA & _
                        "') AND ((TbEquipos.BUI)='" & strBUI & "')) " & _
                        "ORDER BY TbEquipos.Equipo;"
        cbo.RowSource = strSQLCombo
        cbo.Requery
        If cbo.ListCount = 1 Then
            cbo.Value = cbo.Column(0, 0)
        End If
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ActualizarLista la lista de Actividades se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub


