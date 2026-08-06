Option Compare Database
Option Explicit
Dim strIDMaterial As String
Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Function Registrar() As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 17/10/2013
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -Va a emplear la función Material.AltaSeguimiento
    '   -Llamada desde:
    '       -me.ComandoRegistrar_Click
    '   -Devuelve:
    '       Registrar = strIDIntervencion
    '       Registrar = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim strFechaInicioIntervencion As String, strMotivo As String, strResultado As String, strDestino As String, strFechaFinIntervencion As String, _
        strUltimaIntervencion As String, strIDIntervencion As String, strTextoError As String
    On Error GoTo errores
    If strIDMaterial = "" Then
        strTextoError = "No se ha podido determinar cuál es el ID material"
        Err.Raise 1000
    End If
    strFechaInicioIntervencion = Nz(Me.FechaInicioIntervencion, "")
    strFechaFinIntervencion = Nz(Me.FechaFinIntervencion, "")
    strUltimaIntervencion = Nz(Me.ComboUltimaIntervencion, "")
    strDestino = Nz(Me.DESTINO, "")
    strMotivo = Nz(Me.MOTIVO, "")
    strResultado = Nz(Me.RESULTADO, "")
    '-----------------------------------------------
    '   -Devuelve:
    '       MaterialIntervencionAlta = strIDSeguimiento
    '       MaterialIntervencionAlta = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = MaterialIntervencionAlta(strIDMaterial, strFechaInicioIntervencion, strDestino, strMotivo, strResultado, strFechaFinIntervencion, strUltimaIntervencion)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método MaterialIntervencionAlta ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Registrar = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El Registrar se ha devuelto el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    Registrar = "#ERR" & "|" & strTextoError
End Function
Private Sub ComandoRegistrar_Click()
    Dim strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '------------------------------------------------------
    '   -Devuelve:
    '       Registrar = strIDIntervencion
    '       Registrar = "#ERR" & "|" & strTextoError
    '------------------------------------------------------
    flag = Registrar()
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Registrar ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    '---------------------------------------------
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '---------------------------------------------
    If FormularioAbierto("FormMaterialSeguimiento") Then
        Form_FormMaterialSeguimiento.EstablecerDatos
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar grabar se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub ComandoVerEvento_Click()
    Dim strTextoError As String, strIDEvento As String
    On Error GoTo errores
    strIDEvento = Nz(Me.IDEVENTO, "")
    If strIDEvento = "" Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    '-------------------------------------------------
    '   -Llamada desde
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '-------------------------------------------------
    If Not FormularioAbierto("FormEventoGestion") Then
        DoCmd.OpenForm "FormEventoGestion"
        Forms("FormEventoGestion").Controls("ComboIDEventoBusqueda").Value = strIDEvento
        Form_FormEventoGestion.ComboIDEventoBusqueda_AfterUpdate
    Else
        Forms("FormEventoGestion").SetFocus
        If Forms("FormEventoGestion").Controls("ComboIDEventoBusqueda").Value <> strIDEvento Then
            Form_FormEventoGestion.ComandoActualizarTodo_Click
            Forms("FormEventoGestion").Controls("ComboIDEventoBusqueda").Value = strIDEvento
            Form_FormEventoGestion.ComboIDEventoBusqueda_AfterUpdate
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver el Evento origen  se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
    strIDMaterial = Nz(Me.OpenArgs, "")
    If strIDMaterial = "" Then
        strTextoError = "Se ha abierto con parámetros inadecuados"
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       EstablecerDatos = strDescriptivo
    '       Eliminar = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = EstablecerDatos()
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EstablecerDatos ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al Abrir el formulario se ha devuelto el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Public Function EstablecerDatos() As String
    
    Dim rcdDatos As DAO.Recordset, strProvieneDeEvento As String, strGarantia As String, strIDEvento As String, strEquipo As String, _
        strMaterial As String, strTextoError As String
    On Error GoTo errores
    If strIDMaterial = "" Then
        strTextoError = "Se ha abierto con parámetros inadecuados"
        Err.Raise 1000
    End If
    '---------------
    ' OBTENCIÓN DE VALORES DE LOS DATOS PREVIOS DEL MATERIAL
    '---------------
    '   -Devuelve:
    '       MaterialProvieneDeActividad = "Sí" "No"
    '       MaterialProvieneDeActividad = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = MaterialProvieneDeActividad(strIDMaterial)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método MaterialProvieneDeActividad ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag = "Sí" Then
        m_SQL = "SELECT TbEquipos.Equipo, TbEventos.BUI, TbEventos.SUBSISTEMA, TbMaterial.Material, TbActividades.IDEvento,TbMaterial.Garantia " & _
                "FROM ((TbMaterial INNER JOIN TbActividades ON TbMaterial.IDActividad = TbActividades.IDActividad) " & _
                "INNER JOIN TbEventos ON TbActividades.IDEvento = TbEventos.IDEvento) " & _
                "INNER JOIN TbEquipos ON TbEventos.IDEquipo = TbEquipos.IDEquipo " & _
                "WHERE (((TbMaterial.IDMaterial)='" & strIDMaterial & "'));"
        Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
        With rcdDatos
            If .EOF Then
                strTextoError = "Se ha abierto con parámetros inadecuados"
                Err.Raise 1000
            End If
            strEquipo = Nz(.Fields("Equipo"), "")
            strGarantia = Nz(.Fields("Garantia"), "")
            strIDEvento = Nz(.Fields("IDEvento"), "")
            strMaterial = Nz(.Fields("Material"), "")
        End With
    Else
        m_SQL = "SELECT TbEquipos.Equipo, TbMaterial.Material ,TbMaterial.Garantia,TbEquipos.BUI, TbEquipos.SUBSISTEMA " & _
                    "FROM TbMaterial INNER JOIN TbEquipos ON TbMaterial.IDEquipo = TbEquipos.IDEquipo " & _
                    "WHERE (((TbMaterial.IDMaterial)='" & strIDMaterial & "'));"
        Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
        With rcdDatos
            If .EOF Then
                strTextoError = "Se ha abierto con parámetros inadecuados"
                Err.Raise 1000
            End If
            strEquipo = Nz(.Fields("Equipo"), "")
            strGarantia = Nz(.Fields("Garantia"), "")
            strMaterial = Nz(.Fields("Material"), "")
        End With
    End If
    rcdDatos.Close
    Set rcdDatos = Nothing
    If strEquipo <> "" Then
        Me.Equipo = strEquipo
    End If
    If strMaterial <> "" Then
        Me.Material = strMaterial
    End If
    If strIDEvento <> "" Then
        Me.IDEVENTO = strIDEvento
        Me.ComandoVerEvento.Enabled = True
    Else
        Me.ComandoVerEvento.Enabled = False
    End If
    If strGarantia <> "" Then
        Me.Garantia = strGarantia
    End If
    EstablecerDatos = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al EstablecerDatos se ha devuelto el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    EstablecerDatos = "#ERR" & "|" & strTextoError
End Function





