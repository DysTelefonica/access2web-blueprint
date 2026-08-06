Option Compare Database
Option Explicit
Dim strIDPlanificacion As String, strURLAnexo As String
Dim strIDAnexoElegido As String
Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub ComandoAnexar_Click()
    Dim strURL As String, strTitulo As String, strTextoError As String, fso As New FileSystemObject
    On Error GoTo errores
    If strIDPlanificacion = "" Then
        strTextoError = "Se ha abierto el formulario con datos insuficientes"
        Err.Raise 1000
    End If
    strTitulo = Nz(Me.TituloDoc, "")
    If strTitulo = "" Then
        strTextoError = "Se ha de poner un título al documento"
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       Seleccionar = CStr(varFile)
    '       Seleccionar = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = Seleccionar(True, "Seleccione el documento")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Seleccionar ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strURL = flag
    If Not fso.FileExists(strURL) Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       AnexarAPlanificacion = strURLRemoto
    '       AnexarAPlanificacion = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = AnexarAPlanificacion(strIDPlanificacion, strURL, strTitulo)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Eliminar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strURLAnexo = flag
    Me.ListaDocumentos.Requery
    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar adjuntar el anexo se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub ComandoEliminarDocumento_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If Not IsNumeric(strIDPlanificacion) Then
        strTextoError = "No se puede determinar para qué planificación es este anexo"
        Err.Raise 1000
    End If
    If Not IsNumeric(strIDAnexoElegido) Then
        strTextoError = "Ha de seleccionar un anexo de la lista"
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("¿Desea realmente eliminar el anexo seleccionado definitivamente, tanto del disco duro como del registro de la base de datos?", _
                        vbExclamation + vbYesNo + vbDefaultButton2, "Eliminación de anexo")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    DoCmd.Hourglass True
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       EliminarAnexo = Descriptivo
    '       EliminarAnexo = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = EliminarAnexo(strIDAnexoElegido, strIDPlanificacion)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EliminarAnexo ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Me.ListaDocumentos.Requery
    Me.ImagenDocumentoAnexado.Visible = False
    Me.ComandoAnexar.SetFocus
    Me.ComandoEliminarDocumento.Enabled = False
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar eliminar el anexo seleccionado se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub Form_Open(Cancel As Integer)
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    
    Me.Caption = Me.lblTitulo1.Caption
    'strArgumento=strIDPlanificacion
    strIDPlanificacion = Nz(Me.OpenArgs, "")
    If Not IsNumeric(strIDPlanificacion) Then
        strTextoError = "No se ha abierto el formulario con los datos necesarios"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbPlanificacionAnexos.IDAnexoPlanificacion, TbPlanificacionAnexos.IDPlanificacion, " & _
                "TbPlanificacionAnexos.NombreAnexo AS Título, TbPlanificacionAnexos.FechaAnexo AS [Fecha Anexo] " & _
                "FROM TbPlanificacionAnexos " & _
                "WHERE (((TbPlanificacionAnexos.IDPlanificacion)=" & strIDPlanificacion & "));"
    Me.ListaDocumentos.RowSource = m_SQL
    Me.ListaDocumentos.Requery
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
Private Sub ImagenDocumentoAnexado_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       AbrirEnLocal = strURLAnexoLocal
    '       AbrirEnLocal = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = AbrirEnLocal(strURLAnexo, Me.hWnd)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método AbrirEnLocal ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar abrir se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ListaDocumentos_Click()
    Dim fso As New FileSystemObject, strNombreArchivo As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDAnexoElegido = Nz(Me.ListaDocumentos.Column(0), "")
    If Not IsNumeric(strIDPlanificacion) Then
        strTextoError = "Se ha de seleccionar un anexo de la lista"
        Err.Raise 1000
    End If
    '------------------------------------------
    '   -Devuelve:
    '       Dame = strValorObtenido
    '       Dame = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = Dame("TbPlanificacionAnexos", "URLAnexo", "IDAnexoPlanificacion", strIDAnexoElegido)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strURLAnexo = flag
    If strURLAnexo = "" Then
        Me.ImagenDocumentoAnexado.Visible = False
    Else
        Me.ImagenDocumentoAnexado.Visible = True
        If InStr(1, strURLAnexo, ":\") <> 0 Or InStr(1, strURLAnexo, "\\") <> 0 Then
            strNombreArchivo = fso.GetFileName(strURLAnexo)
        Else
            strNombreArchivo = strURLAnexo
        End If
         
        
        strURLAnexo = m_ObjEntorno.URLDirectorioPlanificaciones & strNombreArchivo
    End If
    '-------------------------------------------------------------------
     '   -Devuelve:
    '      EstadoPlanificacion=strEstado
    '      EstadoPlanificacion="-1" Si error
    '       strEstado="Abierta"
    '       strEstado="Cerrada por trabajo realizado"
    '       strEstado="Cerrada por reprogramación"
    '-------------------------------------------------------------------
    flag = EstadoPlanificacion(strIDPlanificacion)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EstadoPlanificacion ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag <> "abierta" Then
        Me.ComandoEliminarDocumento.Enabled = False
    Else
        Me.ComandoEliminarDocumento.Enabled = True
    End If
    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar hacer click en la lista se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Public Function EliminarAnexo(strIDAnexo As String, strIDPlanificacion As String) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 23/06/2014
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -1) strIDAnexo numérico y existente
    '       -2) Se obtiene strIDPlanificacion
    '       -3) Se obtiene estado de la planificacion strIDPlanificacion
    '       -4) Si proviene de cerrada por reprogramación, no se puede eliminar el anexo
    '       -5) Eliminamos el documento del servidor si está
    '       -6) Eliminamos el registro de la tabla TbPlanificacionAnexos
    
    
    '   -llamada por:
    '       -Form_FormPlanificacionAnexos.ComandoEliminarDocumento_Click
    '   -Devuelve:
    '       EliminarAnexo = Descriptivo
    '       EliminarAnexo = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, fso As New FileSystemObject, strNombreArchivo As String, strURLAnexo As String, _
        strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDAnexo) Then
        strTextoError = "El IDAnexo ha de ser un número"
        Err.Raise 1000
    End If
    If Not IsNumeric(strIDPlanificacion) Then
        strTextoError = "El IDPlanificacion ha de ser un número"
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       strEstado = "Abierta"
    '       strEstado="Cerrada por trabajo realizado"
    '       strEstado="Cerrada por reprogramación"
    '       EstadoPlanificacion=strEstado
    '       EstadoPlanificacion="#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = EstadoPlanificacion(strIDPlanificacion)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EstadoPlanificacion ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag = "Cerrada por reprogramación" Then
        strTextoError = "Este anexo pertenece a una planificación que ha sido reprogramada. No se puede eliminar el anexo"
        Err.Raise 1000
    End If
    
    
    m_SQL = "SELECT TbPlanificacionAnexos.IDAnexoPlanificacion,TbPlanificacionAnexos.URLAnexo " & _
            "FROM TbPlanificacionAnexos " & _
            "WHERE (((TbPlanificacionAnexos.IDAnexoPlanificacion)=" & strIDAnexo & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No se ha encontrado ningún anexo con ese ID"
            Err.Raise 1000
        End If
        strURLAnexo = Nz(.Fields("URLAnexo"), "")
        If strURLAnexo <> "" Then
            If InStr(1, strURLAnexo, ":\") Or InStr(1, strURLAnexo, "\\") <> 0 Then
                strNombreArchivo = fso.GetFileName(strURLAnexo)
            Else
                strNombreArchivo = strURLAnexo
            End If
            strURLAnexo = m_ObjEntorno.URLDirectorioPlanificaciones & strNombreArchivo
            If fso.FileExists(strURLAnexo) Then
                fso.DeleteFile strURLAnexo, True
            End If
        End If
        .Delete
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
    EliminarAnexo = "Anexo eliminado"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método EliminarAnexo ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
    EliminarAnexo = "#ERR" & "|" & strTextoError
End Function



