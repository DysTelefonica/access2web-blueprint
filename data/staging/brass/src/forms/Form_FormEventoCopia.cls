Option Compare Database
Option Explicit

Public Event AltaEvento(m_Evento As Evento)
Private m_Error As String

Private Sub cmdSalir_Click()
     On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ComandoCopiar_Click()
    Dim m_EventoCopia As Evento
    Dim m_IDEventoCopia As String
    On Error GoTo errores
    
    m_Error = ""
    pregunta = MsgBox("Se va a copiar el evento seleccionado junto con sus actividades,materiales asociados y seguimientos de los mismos" & vbNewLine & _
                    "¿Desea continuar?", vbExclamation + vbYesNo + vbDefaultButton1, "Copiar Evento")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_IDEventoCopia = CopiarEventoCompletoConActividadesYMateriales(m_ObjEventoActivo.IDEVENTO, Nz(Me.FechaAlta, ""), Nz(Me.BUI, ""), m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_EventoCopia = Constructor.getEvento(p_IDEvento:=m_IDEventoCopia, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RaiseEvent AltaEvento(m_EventoCopia)
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoCopiar_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Form_Open(Cancel As Integer)
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    
    If m_ObjEventoActivo Is Nothing Then
        m_Error = "No se conoce el evento a copiar"
        Err.Raise 1000
    End If
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


Private Function EstablecerDatos(Optional ByRef p_Error As String) As String
    Dim m_Centro As Variant
    Dim cmb As ComboBox
    On Error GoTo errores
    Me.lblTitulo1.Caption = "COPIA DE EVENTO " & m_ObjEventoActivo.IDEVENTO
    Me.lblTitulo2.Caption = Me.lblTitulo1.Caption
    Me.Caption = Me.lblTitulo1.Caption
    Set cmb = Me.BUI
    cmb.RowSource = ""
    If m_ObjEntorno.Centros Is Nothing Then
        Exit Function
    End If
    For Each m_Centro In m_ObjEntorno.Centros
        cmb.AddItem m_Centro
    Next
    Me.BUI = m_ObjEventoActivo.BUI
    Me.FechaAlta = Format(m_ObjEventoActivo.FECHAALTAEVENTO, "dd/mm/yyyy")
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function




Public Function CopiarEventoCompletoConActividadesYMateriales( _
                                                                p_IDEventoACopiar As String, _
                                                                Optional p_FechaAlta As String, _
                                                                Optional p_BUI As String, _
                                                                Optional ByRef p_Error As String _
                                                                ) As String
    
    Dim rcdDatosOrigen As DAO.Recordset, rcdDatosDestino As DAO.Recordset, rcdMaterialesOrigen As DAO.Recordset, fld As DAO.Field, ColMateriales As New Collection, VarItem As Variant
    Dim strIDEvento As String, strIDMaterial As String, strIDActividad As String, strIDActividadOriginal As String, strIDMaterialOriginal As String, _
        strAliasTecnico As String
    On Error GoTo errores
    m_Error = ""
    p_Error = ""
    If Not IsDate(p_FechaAlta) Then
        p_FechaAlta = Date
    End If
    
    '------------------------
    ' COPIA DE EVENTO
    '------------------------
        m_SQL = "TbEventos"
        Set rcdDatosDestino = CurrentDb().OpenRecordset(m_SQL)
        m_SQL = "SELECT TbEventos.* " & _
                    "FROM TbEventos " & _
                    "WHERE (((TbEventos.IDEvento)='" & p_IDEventoACopiar & "'));"
        
        Set rcdDatosOrigen = CurrentDb().OpenRecordset(m_SQL)
        If rcdDatosOrigen.EOF Then
            p_Error = "El Evento seleccionado no está registrado"
            Err.Raise 1000
        End If
        If p_BUI = "" Then
            p_BUI = Nz(rcdDatosOrigen("BUI"), "")
        End If
        
        If p_BUI = "" Then
            p_Error = "El Evento seleccionado no tieme BUI registrado"
            Err.Raise 1000
        End If
        '-----------------------------------------------
        '   -Devuelve:
        '       DameIDEvento = strIDEvento
        '       DameIDEvento = "#ERR" & "|" & p_Error
        '-----------------------------------------------
        flag = DameIDEvento(p_BUI, p_FechaAlta)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            p_Error = "El método DameIDEvento ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strIDEvento = flag
        rcdDatosDestino.AddNew
            For Each fld In rcdDatosDestino.Fields
                If fld.Name = "IDEVENTO" Then
                    rcdDatosDestino(fld.Name) = strIDEvento
                ElseIf fld.Name = "FechaRegistroAlta" Then
                    rcdDatosDestino(fld.Name) = p_FechaAlta
                Else
                    If fld.Name <> "HORAFINALEVENTO" And fld.Name <> "CAUSAFIN" And fld.Name <> "Franqueado" And fld.Name <> "IDExportacion" And fld.Name <> "CodImportacion" And _
                        fld.Name <> "IDParte" Then
                        rcdDatosDestino(fld.Name) = rcdDatosOrigen(fld.Name)
                    End If
                End If
            Next
        rcdDatosDestino.Update
        If Not rcdDatosDestino Is Nothing Then
            rcdDatosDestino.Close
            Set rcdDatosDestino = Nothing
        End If
        If Not rcdDatosOrigen Is Nothing Then
            rcdDatosOrigen.Close
            Set rcdDatosOrigen = Nothing
        End If
    '------------------
    ' ACTIVIDADES
    '----------------------
        m_SQL = "TbActividades"
        Set rcdDatosDestino = CurrentDb().OpenRecordset(m_SQL)
        m_SQL = "SELECT TbActividades.* " & _
                "FROM TbEventos INNER JOIN TbActividades ON TbEventos.IDEvento = TbActividades.IDEvento " & _
                "WHERE (((TbEventos.IDEvento)='" & p_IDEventoACopiar & "'));"
        Set rcdDatosOrigen = CurrentDb().OpenRecordset(m_SQL)
        If Not rcdDatosOrigen.EOF Then
            rcdDatosOrigen.MoveFirst
            Do While Not rcdDatosOrigen.EOF
                strAliasTecnico = Nz(rcdDatosOrigen.Fields("ALIASTECNICO"), "")
                p_FechaAlta = Nz(rcdDatosOrigen.Fields("FechaAlta"), "")
                strIDActividadOriginal = Nz(rcdDatosOrigen.Fields("IDActividad"), "")
                If strAliasTecnico = "" Then
                    p_Error = "No se ha encontrado el Alias técnico de una de las actividades de origen"
                    Err.Raise 1000
                End If
                If Not IsDate(p_FechaAlta) Then
                    p_Error = "No se ha encontrado la Fecha de Alta de una de las actividades de origen"
                    Err.Raise 1000
                End If
                
                rcdDatosDestino.AddNew
                    '-----------------------------------------------------------'
                    '   -Devuelve:
                    '       DameIDActividad = strIDActividad
                    '       DameIDActividad = "ERR" & "|" & p_Error
                    '-----------------------------------------------------------'
                    flag = DameIDActividad(strAliasTecnico, p_FechaAlta)
                    If InStr(1, flag, "|") <> 0 Then
                        dato = Split(flag, "|")
                        p_Error = "El método EVE ha devuelto el error: " & vbNewLine & dato(1)
                        Err.Raise 1000
                    End If
                    strIDActividad = flag
                    For Each fld In rcdDatosDestino.Fields
                        If fld.Name = "IDActividad" Then
                            rcdDatosDestino(fld.Name) = strIDActividad
                        ElseIf fld.Name = "IDEvento" Then
                            rcdDatosDestino(fld.Name) = strIDEvento
                        Else
                            If fld.Name <> "IDExportacion" And fld.Name <> "CodImportacion" And fld.Name <> "IDFacturacion" Then
                                rcdDatosDestino(fld.Name) = rcdDatosOrigen(fld.Name)
                            End If
                        End If
                    Next
                rcdDatosDestino.Update
                '------------
                ' copiar todos los materiales de esta actividad
                '-----------------------------------
                m_SQL = "SELECT TbMaterial.* " & _
                        "FROM TbMaterial " & _
                        "WHERE (((TbMaterial.IDActividad)='" & strIDActividadOriginal & "'));"
                Set rcdMaterialesOrigen = CurrentDb().OpenRecordset(m_SQL)
                If Not rcdMaterialesOrigen.EOF Then
                    rcdMaterialesOrigen.MoveFirst
                    Do While Not rcdMaterialesOrigen.EOF
                        strIDMaterialOriginal = Nz(rcdMaterialesOrigen.Fields("IDMaterial"), "")
                        '-----------------------------------------------
                        '   -Devuelve:
                        '       MaterialDeActividadAlta = strIDMaterial
                        '       MaterialDeActividadAlta = "#ERR" & "|" & p_Error
                        '-----------------------------------------------
                        flag = MaterialDeActividadAlta(strIDActividad, _
                                                        Nz(rcdMaterialesOrigen.Fields("Material"), ""), _
                                                        Nz(rcdMaterialesOrigen.Fields("COSTE"), ""), _
                                                        Nz(rcdMaterialesOrigen.Fields("TipoAccion"), ""), _
                                                        Nz(rcdMaterialesOrigen.Fields("PN"), ""), _
                                                        Nz(rcdMaterialesOrigen.Fields("NS"), ""), _
                                                        Nz(rcdMaterialesOrigen.Fields("ReparadoPor"), ""), _
                                                        Nz(rcdMaterialesOrigen.Fields("Descripcion"), ""), _
                                                        Nz(rcdMaterialesOrigen.Fields("FechaIncio"), ""), _
                                                        Nz(rcdMaterialesOrigen.Fields("Garantia"), ""), _
                                                        "Sí", _
                                                        Nz(rcdMaterialesOrigen.Fields("FechaEntrega"), ""))
                        If InStr(1, flag, "|") <> 0 Then
                            dato = Split(flag, "|")
                            p_Error = "El método MaterialDeActividadAlta ha devuelto un error: " & vbNewLine & dato(1)
                            Err.Raise 1000
                        End If
                        strIDMaterial = flag
                        ColMateriales.Add strIDMaterial
                        
                        CopiarAnexoMaterial strIDMaterialOriginal, strIDMaterial, p_Error
                        If p_Error <> "" Then
                            Err.Raise 1000
                        End If
                        CopiarSeguimientoMaterial strIDMaterialOriginal, strIDMaterial, p_Error
                        If p_Error <> "" Then
                            Err.Raise 1000
                        End If
                        
                        
                        rcdMaterialesOrigen.MoveNext
                    Loop
                End If
                
                rcdDatosOrigen.MoveNext
            Loop
        End If
        If Not rcdDatosDestino Is Nothing Then
            rcdDatosDestino.Close
            Set rcdDatosDestino = Nothing
        End If
        If Not rcdDatosOrigen Is Nothing Then
            rcdDatosOrigen.Close
            Set rcdDatosOrigen = Nothing
        End If
        If Not rcdMaterialesOrigen Is Nothing Then
            rcdMaterialesOrigen.Close
            Set rcdMaterialesOrigen = Nothing
        End If
    CopiarEventoCompletoConActividadesYMateriales = strIDEvento
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CopiarEventoCompletoConActividadesYMateriales ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatosDestino Is Nothing Then
        rcdDatosDestino.Close
        Set rcdDatosDestino = Nothing
    End If
    If Not rcdDatosOrigen Is Nothing Then
        rcdDatosOrigen.Close
        Set rcdDatosOrigen = Nothing
    End If
    On Error Resume Next
    If strIDEvento <> "" Then
        m_SQL = "DELETE TbEventos.* " & _
                "FROM TbEventos " & _
                "WHERE (((TbEventos.IDEvento)='" & strIDEvento & "'));"
        DoCmd.SetWarnings False
        DoCmd.RunSQL m_SQL
        DoCmd.SetWarnings True
        m_SQL = "DELETE TbActividades.* " & _
                "FROM TbActividades " & _
                "WHERE (((TbActividades.IDEvento)='" & strIDEvento & "'));"
        DoCmd.SetWarnings False
        DoCmd.RunSQL m_SQL
        DoCmd.SetWarnings True
        If ColMateriales.count > 0 Then
            For Each VarItem In ColMateriales
                m_SQL = "SELECT TbAnexos.IDMaterial " & _
                        "FROM TbAnexos " & _
                        "WHERE (((TbAnexos.IDMaterial)='" & VarItem & "'));"
                DoCmd.SetWarnings False
                DoCmd.RunSQL m_SQL
                DoCmd.SetWarnings True
                m_SQL = "DELETE TbMaterial.* " & _
                        "FROM TbMaterial " & _
                        "WHERE (((TbMaterial.IDMaterial)='" & CStr(VarItem) & "'));"
                DoCmd.SetWarnings False
                DoCmd.RunSQL m_SQL
                DoCmd.SetWarnings True
            Next
        End If
    End If
    
    
    
End Function
Public Function CopiarAnexoMaterial( _
                                        strIDMaterialOriginal As String, _
                                        strIDMaterialDestino As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim rcdDatos As DAO.Recordset, rcdAnexo As DAO.Recordset, strIDAnexo As String, strNombreArchivo As String, strDESCRIPCION As String, strTitulo As String
    On Error GoTo errores
    m_Error = ""
    p_Error = ""
    m_SQL = "TbAnexos"
    Set rcdAnexo = CurrentDb().OpenRecordset(m_SQL)
    m_SQL = "SELECT TbAnexos.Titulo, TbAnexos.Descripcion, TbAnexos.NombreArchivo " & _
            "FROM TbAnexos " & _
            "WHERE (((TbAnexos.IDMaterial)='" & strIDMaterialOriginal & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                strTitulo = Nz(.Fields("Titulo"), "")
                strDESCRIPCION = Nz(.Fields("Descripcion"), "")
                strNombreArchivo = Nz(.Fields("NombreArchivo"), "")
                If strNombreArchivo <> "" Then
                    strIDAnexo = DameID("TbAnexos", "IDAnexo")
                    
                    rcdAnexo.AddNew
                        rcdAnexo.Fields("IDAnexo") = strIDAnexo
                        rcdAnexo.Fields("IDMaterial") = strIDMaterialDestino
                        If strTitulo <> "" Then
                            rcdAnexo.Fields("Titulo") = strTitulo
                        End If
                        If strDESCRIPCION <> "" Then
                            rcdAnexo.Fields("Descripcion") = strDESCRIPCION
                        End If
                        rcdAnexo.Fields("NombreArchivo") = strNombreArchivo
                    rcdAnexo.Update
                End If
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If Not rcdAnexo Is Nothing Then
        rcdAnexo.Close
        Set rcdAnexo = Nothing
    End If
    CopiarAnexoMaterial = "Sí"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CopiarAnexoMaterial ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    If Not rcdAnexo Is Nothing Then
        rcdAnexo.Close
        Set rcdAnexo = Nothing
    End If
    
End Function
Public Function CopiarSeguimientoMaterial( _
                                            p_IDMaterialOrigen As String, _
                                            p_IDMaterialDestino As String, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    Dim rcdDatosDestino As DAO.Recordset
    Dim rcdDatosOrigen As DAO.Recordset
    
    Dim strIDIntervencion As String
   
    On Error GoTo errores
    m_Error = ""
    p_Error = ""
    m_SQL = "TbMaterialSeguimiento"
    Set rcdDatosDestino = CurrentDb().OpenRecordset(m_SQL)
    m_SQL = "SELECT TbMaterialSeguimiento.* " & _
            "FROM TbMaterialSeguimiento " & _
            "WHERE (((TbMaterialSeguimiento.IDMaterial)='" & p_IDMaterialOrigen & "'));"
    Set rcdDatosOrigen = CurrentDb().OpenRecordset(m_SQL)
    If rcdDatosOrigen.EOF Then
        Exit Function
    End If
    rcdDatosOrigen.MoveFirst
    Do While Not rcdDatosOrigen.EOF
        rcdDatosDestino.AddNew
            strIDIntervencion = DameIDIntervencion(p_IDMaterialDestino)
            rcdDatosDestino.Fields("IDSeguimiento") = strIDIntervencion
            rcdDatosDestino.Fields("IDMaterial") = p_IDMaterialDestino
            rcdDatosDestino.Fields("FechaInicioIntervencion") = rcdDatosOrigen.Fields("FechaInicioIntervencion")
            rcdDatosDestino.Fields("Destino") = rcdDatosOrigen.Fields("Destino")
            rcdDatosDestino.Fields("Motivo") = rcdDatosOrigen.Fields("Motivo")
            rcdDatosDestino.Fields("FechaFinIntervencion") = rcdDatosOrigen.Fields("FechaFinIntervencion")
            rcdDatosDestino.Fields("Resultado") = rcdDatosOrigen.Fields("Resultado")
            rcdDatosDestino.Fields("UltimaIntervencion") = rcdDatosOrigen.Fields("UltimaIntervencion")
            
        rcdDatosDestino.Update
        rcdDatosOrigen.MoveNext
    Loop
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CopiarSeguimientoMaterial ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function