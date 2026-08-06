Option Compare Database
Option Explicit
Dim strBUI As String
Dim strSubSistemaAlInicio As String
Private Function HaHabidoCambios() As Boolean
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día  29/01/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -Va a True si hay cambios y False si no
    
    '   -Llamada desde
    
    '   -Devuelve:
    '       HaHabidoCambios = True Or False
    '
    '           strEstado="1"----> True Or False
    '           strEstado="-1"--->False
    '-------------------------------------------------------------------
    Dim strSUBSISTEMA As String
    On Error GoTo errores
    strSUBSISTEMA = Nz(Me.SubSistema, "")
    If strSUBSISTEMA <> strSubSistemaAlInicio Then
        HaHabidoCambios = True
    Else
        HaHabidoCambios = False
    End If
    Exit Function
errores:
    HaHabidoCambios = True
End Function
Private Sub cmdSalir_Click()
    Dim strTextoError As String, strNombreFormulario As String
    On Error GoTo errores
    strNombreFormulario = Me.Name
    If HaHabidoCambios() Then
        pregunta = MsgBox("¿Desea guardar los cambios producidos?", vbExclamation + vbYesNo + vbDefaultButton1, "Guardar Cambios")
        If pregunta = vbYes Then
            VBA.DoEvents
            DoCmd.Hourglass True
            VBA.DoEvents
            '-------------------------------------------------------------------
            '   -Devuelve:
            '       Registrar = Descriptivo
            '       Registrar = "#ERR" & "|" & strTextoError
            '-----------------------------------------------
            flag = Registrar()
            If InStr(1, flag, "|") <> 0 Then
                dato = Split(flag, "|")
                strTextoError = "El método Registrar ha devuelto el error: " & vbNewLine & dato(1)
                Err.Raise 1000
            End If
            '--------------------------------------------------
            '   -Devuelve
            '       FormularioAbierto = true or false
            '       FormularioAbierto =False-->Error
            '--------------------------------------------------
            If FormularioAbierto("FormSubSistemaGestion") Then
                Forms("FormSubSistemaGestion").Controls("ListaSubSistema").Requery
            End If
            VBA.DoEvents
            DoCmd.Hourglass False
            VBA.DoEvents
        End If
    End If
    DoCmd.Close acForm, strNombreFormulario, acSaveNo
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar salir del formulario se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoRegistrar_Click()
    Dim strNombreFormulario As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strNombreFormulario = Me.Name
    If Not HaHabidoCambios() Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        pregunta = MsgBox("No ha habido cambios que guardar", vbCritical, "Sin cambios")
        Exit Sub
    End If
    '-----------------------------------------------
    '   -Devuelve:
    '       Registrar = strEstaFacturado
    '       Registrar = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Registrar()
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Registrar ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    '--------------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '--------------------------------------------------
    If FormularioAbierto("FormSubSistemaGestion") Then
        Forms("FormSubSistemaGestion").Controls("ListaSubSistema").Requery
    End If
    DoCmd.Close acForm, strNombreFormulario, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar registrar el gasto se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Function Registrar() As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día  x
    '   -Modificaciones:
    
    '   -Funcionamiento:
    
    
    '   -Llamada desde
    
    '   -Devuelve:
    '       Registrar = Descriptivo
    '       Registrar = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim strSUBSISTEMA As String, strTextoError As String
    On Error GoTo errores
    strSUBSISTEMA = Nz(Me.SubSistema, "")
    '-----------------------------------------------
    '   -Devuelve:
    '       Edicion = strDescriptivo
    '       Edicion = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Edicion(strSUBSISTEMA)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Edicion ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Registrar = "Edición Correcta"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Registrar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    Registrar = "#ERR" & "|" & strTextoError
End Function
Private Sub Form_Open(Cancel As Integer)
    Dim strArgumento As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    Me.Caption = Me.lblTitulo1.Caption
    'strArgumento=strBUI & "|" & strSubSistemaAlInicio
    strArgumento = Nz(Me.OpenArgs, "")
    If strArgumento = "" Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    If InStr(1, strArgumento, "|") = 0 Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    dato = Split(strArgumento, "|")
    strBUI = Nz(dato(0), "")
    strSubSistemaAlInicio = Nz(dato(1), "")
    
    If strBUI = "" Or strSubSistemaAlInicio = "" Then
        strTextoError = "No se han cargado todos los datos"
        Err.Raise 1000
    End If
    Me.BUI = strBUI
    Me.SubSistema = strSubSistemaAlInicio
   VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar abrir el formulario se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Public Function Edicion(strSUBSISTEMA As String) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 29/12/13
    '   -Modificaciones:
   
    '   -Funcionamiento:
    '       -va a editar el nombre de la tabla TbBUI el nodo con BUI=strBUI
    '       -strBUI<>""
    '       -strBUI Existente
    '       -strSubSistemaInicial<>""
    '       -strBUI, strSubSistemaAlInicio Existente
    '       -strBUI,strSubSistema no existente
    '
    '       -strSubSistemaAlInicio no ha de estar previamente en ningún evento-->me.EsteSubSistemaEstaYaEnEvento="No"
    '   -llamada por:
    '       -Form_FormSubSistemaEdicion.Editar
    '   -Devuelve:
    '       Edicion = Descriptivo
    '       Edicion = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    If strBUI = "" Then
        strTextoError = "No se ha indicado el nombre del BUI del Subsistema"
        Err.Raise 1000
    End If
    If strSubSistemaAlInicio = "" Then
        strTextoError = "No se ha indicado el nombre del SubSistema que queremos que cambiar"
        Err.Raise 1000
    End If
    If strSUBSISTEMA = "" Then
        strTextoError = "No se ha indicado el nombre del SubSistema que queremos que tenga"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbSubsistemaBui.SUBSISTEMA, TbSubsistemaBui.BUI " & _
            "FROM TbSubsistemaBui " & _
            "WHERE (((TbSubsistemaBui.SUBSISTEMA)='" & strSubSistemaAlInicio & _
            "') AND ((TbSubsistemaBui.BUI)='" & strBUI & "'));"
    
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No está registrado el Subsistema : " & strSubSistemaAlInicio & " para el BUI llamado: " & strBUI
            Err.Raise 1000
        End If
    End With
    
    '-----------------------------------------------
    '   -Devuelve:
    '       Dame = strDescriptivo
    '       Dame = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = Dame("TbEventos", "SUBSISTEMA", "SUBSISTEMA", strSubSistemaAlInicio)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag <> "" Then
        strTextoError = "El SubSistema: " & strSubSistemaAlInicio & " para el BUI: " & strBUI & " ha formado parte de algún evento"
        Err.Raise 1000
    End If
    With rcdDatos
        .Edit
            .Fields("SUBSISTEMA") = strSUBSISTEMA
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Edicion = "Se ha editado correctamente el SubSistema"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método SubSistema.Edicion ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    Edicion = "#ERR" & "|" & strTextoError
End Function







