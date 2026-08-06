Option Compare Database
Option Explicit
Dim strIDCalibracionSeleccionada As String, strFechaBajaParaCalibracion As String
Dim strIDEquipoCalibracion As String, strURLAnexo As String
Private Sub ComandoExportarExcel_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_SQL = Me.ListaFiltrados.RowSource
    GenerarConsultasSQL m_SQL, strTextoError
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar obtener el resultado en excel se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoAdjuntarDocumento_Click()
    Dim fso As New FileSystemObject, strURLLocal As String, strTituloBoton As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDCalibracionSeleccionada = Nz(Me.ListaFiltrados.Column(0), "")
    If Not IsNumeric(strIDCalibracionSeleccionada) Then
        strTextoError = "Seleccione un elemento de la lista superior"
        Err.Raise 1000
    End If
    strTituloBoton = Me.ComandoAdjuntarDocumento.Caption
    If strTituloBoton = "Adjuntar" Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       Seleccionar = CStr(varFile)
        '       Seleccionar = "#ERR" & "|" & strTextoError
        '-------------------------------------------------------------------
        flag = Seleccionar(True, "Elija el documento a anexar")
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método Seleccionar ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strURLLocal = flag
        If Not fso.FileExists(strURLLocal) Then
            If Not fso Is Nothing Then
                Set fso = Nothing
            End If
            Exit Sub
        End If
        VBA.DoEvents
        DoCmd.Hourglass True
        VBA.DoEvents
        '-----------------------------------------------
        '   -Devuelve:
        '       AnexarDocumento = strURLAnexo
        '       AnexarDocumento = "#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = AnexarDocumento(strURLLocal)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método AnexarDocumento ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strURLAnexo = flag
        Me.ImagenAnexo.Visible = True
        Me.ComandoAdjuntarDocumento.Caption = "Quitar Adjunto"
    ElseIf strTituloBoton = "Quitar Adjunto" Then
        pregunta = MsgBox("¿Desea eliminar el adjunto?", vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar Anexo")
        If pregunta <> vbYes Then
            If Not fso Is Nothing Then
                Set fso = Nothing
            End If
            VBA.DoEvents
            DoCmd.Hourglass False
            VBA.DoEvents
            Exit Sub
        End If
        '-----------------------------------------------
        '   -Devuelve:
        '       QuitarAnexo = strDescriptivo
        '       QuitarAnexo = "#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = QuitarAnexo()
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método QuitarAnexo ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        Me.ImagenAnexo.Visible = False
        Me.ComandoAdjuntarDocumento.Caption = "Adjuntar"
    Else
        strTextoError = "No se ha programado para el título del botón " & Me.ComandoAdjuntarDocumento.Caption
        Err.Raise 1000
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
        strTextoError = "Al intentar se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoAltaCalibracion_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If Not IsNumeric(strIDEquipoCalibracion) Then
        strTextoError = "No se se puede determinar el equipo"
        Err.Raise 1000
    End If
    '-----------------------------------------------
    '   -Llamada desde
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '-----------------------------------------------
    If FormularioAbierto("FormCalibracionAlta") Then
        DoCmd.Close acForm, "FormCalibracionAlta", acSaveNo
    End If
    DoCmd.OpenForm "FormCalibracionAlta", , , , , , strIDEquipoCalibracion
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar dar de alta una calibración se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoEditarCalibracion_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDCalibracionSeleccionada = Nz(Me.ListaFiltrados.Column(0), "")
    If Not IsNumeric(strIDCalibracionSeleccionada) Then
        strTextoError = "Ha de seleccionar una calibración primero"
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       FormularioAbierto =True or false
    '           strEstado="1"---->True or false
    '           strEstado="-1"--->False
    '-------------------------------------------------------------------
    If FormularioAbierto("FormCalibracionEdicion") Then
        DoCmd.Close acForm, "FormCalibracionEdicion", acSaveNo
    End If
    DoCmd.OpenForm "FormCalibracionEdicion", , , , , , strIDCalibracionSeleccionada
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar editar la calibración se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Sub ComandoEliminarCalibracion_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDCalibracionSeleccionada = Nz(Me.ListaFiltrados.Column(0), "")
    If Not IsNumeric(strIDCalibracionSeleccionada) Then
        strTextoError = "Se ha de elegir una calibración que queramos borrar"
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    pregunta = MsgBox("¿Desea borrar la calibración seleccionada junto con el anexo que pudiera tener?", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "Borrado de Calibración")
    If pregunta <> vbYes Then
       Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       EliminarCalibracion = strDescriptivo
    '       EliminarCalibracion = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = EliminarCalibracion()
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EliminarCalibracion ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    '-----------------------------------------------
    '   -Llamada desde
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '-----------------------------------------------
    If FormularioAbierto("FormCalibracionEquipoGestion") Then
        Form_FormCalibracionEquipoGestion.ActualizarLista
    End If
    
    Me.ListaFiltrados.Requery
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar borrar una calibración se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub Form_Open(Canel As Integer)
    Dim rcdDatos As DAO.Recordset, strArgumento As String, strTextoCalibraciones, strEquipo As String, _
        strModelo As String, strNS As String, strBUI As String, intElementos As Integer, lst As ListBox, VarItem As Variant, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    Me.Caption = Me.lblTitulo1.Caption
    'strArgumento=strIDEquipoCalibracion & "|" & strTextoCalibraciones
    strArgumento = Nz(Me.OpenArgs(), "")
    If InStr(1, strArgumento, "|") <> 0 Then
        dato = Split(strArgumento, "|")
        strIDEquipoCalibracion = CStr(dato(0))
        strTextoCalibraciones = CStr(dato(1))
    Else
        strIDEquipoCalibracion = strArgumento
    End If
    If Not IsNumeric(strIDEquipoCalibracion) Then
        strTextoError = "No se ha abierto con los parámetros adecuados"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbEquiposCalibrables.Equipo, TbEquiposCalibrables.Modelo, TbEquiposCalibrables.NS,TbEquiposCalibrables.BUI, " & _
            "TbEquiposCalibrables.FechaBajaParaCalibracion " & _
            "FROM TbEquiposCalibrables " & _
            "WHERE (((TbEquiposCalibrables.IDEquipoCalibrable)=" & strIDEquipoCalibracion & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No se ha abierto con los parámetros adecuados"
            Err.Raise 1000
        End If
        strEquipo = Nz(.Fields("Equipo"), "")
        strModelo = Nz(.Fields("Modelo"), "")
        strNS = Nz(.Fields("NS"), "")
        strBUI = Nz(.Fields("BUI"), "")
        strFechaBajaParaCalibracion = Nz(.Fields("FechaBajaParaCalibracion"), "")
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If strEquipo <> "" Then
        Me.Equipo = strEquipo
    End If
     If strModelo <> "" Then
        Me.Modelo = strModelo
    End If
     If strNS <> "" Then
        Me.NS = strNS
    End If
    If strBUI <> "" Then
        Me.BUI = strBUI
    End If
    If strTextoCalibraciones <> "" Then
        Me.lblCalibraciones.Caption = strTextoCalibraciones
    End If
    If IsDate(strFechaBajaParaCalibracion) Then
        Me.ComandoAltaCalibracion.Enabled = False
    Else
        Me.ComandoAltaCalibracion.Enabled = True
    End If
    Me.Caption = Me.lblTitulo1.Caption
    Set lst = Me.ListaFiltrados
    lst.RowSource = m_SQL
    lst.Requery
    intElementos = lst.ListCount - 1
    If intElementos = -1 Then intElementos = 0
    If intElementos = 0 Then
        Me.ComandoExportarExcel.Enabled = False
    Else
        If intElementos = 1 Then
            lst.Selected(1) = True
            ListaFiltrados_Click
        Else
            For Each VarItem In lst.ItemsSelected
                lst.Selected(VarItem) = False
            Next
        End If
        Me.ComandoExportarExcel.Enabled = True
    End If
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
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub cmdSalir_Click()
   On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub ImagenAnexo_Click()
    Dim strTextoError As String
    On Error GoTo errores
    DoCmd.Hourglass True
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
    DoCmd.Hourglass False
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar abrir se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Public Sub ListaFiltrados_Click()
    Dim fso As New FileSystemObject, strNombreArchivo As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strIDCalibracionSeleccionada = Nz(Me.ListaFiltrados.Column(0), "")
    strURLAnexo = Nz(Me.ListaFiltrados.Column(4), "")
    If strIDCalibracionSeleccionada = "" Then
        Me.ComandoEditarCalibracion.Enabled = False
        Me.ComandoEliminarCalibracion.Enabled = False
        Me.ComandoAdjuntarDocumento.Caption = "Adjuntar"
        Me.ComandoAdjuntarDocumento.Enabled = False
        Me.ImagenAnexo.Visible = False
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Me.ComandoAdjuntarDocumento.Enabled = True
    If strURLAnexo <> "" Then
        If InStr(1, strURLAnexo, ":\") <> 0 Or InStr(1, strURLAnexo, "\\") <> 0 Then
            strNombreArchivo = fso.GetFileName(strURLAnexo)
        Else
            strNombreArchivo = strURLAnexo
        End If
        
        
        strURLAnexo = m_ObjEntorno.URLDirectorioCalibraciones & strNombreArchivo
        Me.ImagenAnexo.Visible = True
        Me.ComandoAdjuntarDocumento.Caption = "Quitar Adjunto"
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar hacer click sobre la lista de actividades se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Public Function AnexarDocumento( _
                                    strURLDocLocal As String _
                                    ) As String
                                               
      '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 14/03/2014
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -va a añadir en el campo "TbEquiposCalibrablesFechas.URLDocumentosAnexados el resultado final de copiar a remoto lo que viene
    '       -va a Copiar en su ubicación definitiva el documento que viene
    '   -Llamada desde
    '       -Form_FormCalibracionGestion.ComandoAdjuntarDocumento_Click
    '   -Devuelve:
    '       AnexarDocumento = strURLAnexoNuevo
    '       AnexarDocumento = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim fso As New FileSystemObject, strNombreArchivo As String, strExtensionDocumento As String, strURLAnexoNuevo As String, strTextoError As String
    On Error GoTo errores
    
    strIDCalibracionSeleccionada = Nz(Me.ListaFiltrados.Column(0), "")
    strURLAnexo = Nz(Me.ListaFiltrados.Column(4), "")
    If Not IsNumeric(strIDCalibracionSeleccionada) Then
        strTextoError = "Se ha de indicar qué calibración es"
        Err.Raise 1000
    End If
    If Not fso.FileExists(strURLDocLocal) Then
        strTextoError = "Se ha de indicar el anexo local"
        Err.Raise 1000
    Else
        '----------------------------------------------
        '   -Devuelve:
        '       FicheroAbierto = true or false
        '       FicheroAbierto =False-->Error
        '----------------------------------------------
        If FicheroAbierto(strURLDocLocal) Then
            strTextoError = "El anexo está abierto."
            Err.Raise 1000
        End If
    End If
    
    
    If strURLAnexo <> "" Then
        If InStr(1, strURLAnexo, "\\") <> 0 Or InStr(1, strURLAnexo, ":\") <> 0 Then
            strNombreArchivo = fso.GetFileName(strURLAnexo)
        Else
            strNombreArchivo = strURLAnexo
        End If
        strURLAnexo = m_ObjEntorno.URLDirectorioCalibraciones & strNombreArchivo
        If fso.FileExists(strURLAnexo) Then
            '----------------------------------------------
            '   -Devuelve:
            '       FicheroAbierto = true or false
            '       FicheroAbierto =False-->Error
            '----------------------------------------------
            If FicheroAbierto(strURLAnexo) Then
                strTextoError = "Actualmente hay un anexo y está abierto"
                Err.Raise 1000
            End If
        End If
    End If
    strExtensionDocumento = fso.GetExtensionName(strURLDocLocal)
    strNombreArchivo = Format(strIDCalibracionSeleccionada, "0000")
    strNombreArchivo = strNombreArchivo & "." & strExtensionDocumento
    strURLAnexoNuevo = m_ObjEntorno.URLDirectorioCalibraciones & strNombreArchivo
    If fso.FileExists(strURLAnexoNuevo) Then
        '----------------------------------------------
        '   -Devuelve:
        '       FicheroAbierto = true or false
        '       FicheroAbierto =False-->Error
        '----------------------------------------------
        If FicheroAbierto(strURLAnexoNuevo) Then
            strTextoError = "Existía un documento y no ha podido ser sobrescrito, debe estar abierto"
            Err.Raise 1000
        End If
        fso.DeleteFile strURLAnexo, True
    End If
    If fso.FileExists(strURLAnexo) Then
        fso.DeleteFile strURLAnexo, True
    End If
    fso.CopyFile strURLDocLocal, strURLAnexoNuevo, True
    m_SQL = "UPDATE TbEquiposCalibrablesFechas SET TbEquiposCalibrablesFechas.URLDocumentosAnexados = '" & fso.GetFileName(strURLAnexoNuevo) & "' " & _
            "WHERE (((TbEquiposCalibrablesFechas.IDCalibracion)=" & strIDCalibracionSeleccionada & "));"
    DoCmd.SetWarnings False
    DoCmd.RunSQL m_SQL
    DoCmd.SetWarnings True
    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
    AnexarDocumento = strURLAnexoNuevo
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método AnexarDocumento ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
    DoCmd.SetWarnings True
    AnexarDocumento = "#ERR" & "|" & strTextoError
End Function
Public Function QuitarAnexo() As String
     '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 14/03/2014
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -va a borrar en el campo "TbEquiposCalibrablesFechas.URLDocumentosAnexados =Null
    '       -va a eliminar en su ubicación definitiva el documento que viene
    '   -Llamada desde
    '       -Form_FormCalibracionGestion.ComandoAdjuntarDocumento_Click
    '       -Form_FormCalibracionGestion.ComandoEliminarCalibracion_Click
    '   -Devuelve:
    '       QuitarAnexo = Descriptivo
    '       QuitarAnexo = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim fso As New FileSystemObject, strNombreArchivo As String, strTextoError As String
    On Error GoTo errores
    
    strIDCalibracionSeleccionada = Nz(Me.ListaFiltrados.Column(0), "")
    strURLAnexo = Nz(Me.ListaFiltrados.Column(4), "")
    If Not IsNumeric(strIDCalibracionSeleccionada) Then
        strTextoError = "Ha de seleccionar una calibración primero"
        Err.Raise 1000
    End If
    m_SQL = "UPDATE TbEquiposCalibrablesFechas SET TbEquiposCalibrablesFechas.URLDocumentosAnexados = Null " & _
            "WHERE (((TbEquiposCalibrablesFechas.IDCalibracion)=" & strIDCalibracionSeleccionada & "));"
    DoCmd.SetWarnings False
    DoCmd.RunSQL m_SQL
    DoCmd.SetWarnings True
    If strURLAnexo <> "" Then
        
        
        If InStr(1, strURLAnexo, ":\") <> 0 Or InStr(1, strURLAnexo, "\\") <> 0 Then
            strNombreArchivo = fso.GetFileName(strURLAnexo)
        Else
            strNombreArchivo = strURLAnexo
        End If
        strURLAnexo = m_ObjEntorno.URLDirectorioCalibraciones & strNombreArchivo
        '----------------------------------------
        '   -Devuelve:
        '       FicheroAbierto = true or false
        '       FicheroAbierto =False-->Error
        '----------------------------------------
        If fso.FileExists(strURLAnexo) Then
            If (strURLAnexo) Then
                strTextoError = "El anexo está abierto"
                Err.Raise 1000
            End If
        End If
        fso.DeleteFile strURLAnexo, True
    End If
    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
    strURLAnexo = ""
    QuitarAnexo = "Eliminado el anexo correctamente"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método QuitarAnexo ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
    DoCmd.SetWarnings True
    QuitarAnexo = "#ERR" & "|" & strTextoError
End Function
Public Function EliminarCalibracion() As String
     '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 14/11/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
   
    '   -Llamada desde
    '       -Form_FormCalibracionGestion.ComandoEliminarCalibracion_Click
    '   -Devuelve:
    '       EliminarCalibracion = Descripcion
    '       EliminarCalibracion = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim strTextoError As String
    On Error GoTo errores
    
    strIDCalibracionSeleccionada = Nz(Me.ListaFiltrados.Column(0), "")
    If strIDCalibracionSeleccionada = "" Then
        strTextoError = "Seleccione una calibración"
        Err.Raise 1000
    End If
    '-----------------------------------------------
    '   -Devuelve:
    '       QuitarAnexo = Descripcion
    '       QuitarAnexo = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = QuitarAnexo()
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Eliminar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    m_SQL = "DELETE TbEquiposCalibrablesFechas.IDCalibracion " & _
                "FROM TbEquiposCalibrablesFechas " & _
                "WHERE (((TbEquiposCalibrablesFechas.IDCalibracion)=" & strIDCalibracionSeleccionada & "));"
    DoCmd.SetWarnings False
    DoCmd.RunSQL m_SQL
    DoCmd.SetWarnings True
    EliminarCalibracion = "Calibración eliminada correctamente"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método EliminarCalibracion ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.SetWarnings True
    EliminarCalibracion = "#ERR" & "|" & strTextoError
End Function

