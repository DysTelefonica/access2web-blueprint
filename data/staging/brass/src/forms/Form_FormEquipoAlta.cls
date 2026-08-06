Option Compare Database
Option Explicit
Dim strValorAnteriorNodo As String, strValorAnteriorBUI As String, strValorAnteriorSubSistema As String

Private Sub ComandoRegistrar_Click()
    Dim strNodo As String, strBUI As String, strSUBSISTEMA As String, strIDEquipo As String, strDESCRIPCION As String, strEquipo As String, _
        strNombreFormulario As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strNombreFormulario = Me.Name
    strNodo = Nz(Me.ComboNodo, "")
    strBUI = Nz(Me.ComboBui, "")
    strSUBSISTEMA = Nz(Me.ComboSubsistema, "")
    strEquipo = Nz(Me.Equipo, "")
    strDESCRIPCION = Nz(Me.Descripcion, "#")
    '-----------------------------------------------
    '    -Devuelve:
    '       AltaIDEquipo = strIDEquipo
    '       AltaIDEquipo = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = AltaIDEquipo(strNodo, strBUI, strSUBSISTEMA, strEquipo, strDESCRIPCION)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método AltaIDEquipo ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strIDEquipo = flag
    '--------------------------------------------
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '--------------------------------------------
    If FormularioAbierto("FormEquipoGestion") Then
        Form_FormEquipoGestion.ActualizarLista
    End If
    DoCmd.Close acForm, strNombreFormulario, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar dar de alta el IDEquipo se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Public Sub ComboBui_BeforeUpdate(Cancel As Integer)
    Dim strValorBUI As String, strValorActualSubSistema As String, strSQLSubSistema As String
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValorBUI = Nz(Me.ComboBui, "")
    
    If strValorBUI = "" Then
        strSQLSubSistema = "SELECT DISTINCT TbSubsistemaBui.SUBSISTEMA " & _
                            "FROM TbSubsistemaBui " & _
                            "WHERE (((TbSubsistemaBui.BUI) Like '*')) " & _
                            "ORDER BY TbSubsistemaBui.SUBSISTEMA;"
    Else
        
        strSQLSubSistema = "SELECT DISTINCT TbSubsistemaBui.SUBSISTEMA " & _
                            "FROM TbSubsistemaBui " & _
                            "WHERE (((TbSubsistemaBui.BUI) = '" & strValorBUI & "')) " & _
                            "ORDER BY TbSubsistemaBui.SUBSISTEMA;"
    End If
    Me.ComboSubsistema.RowSource = strSQLSubSistema
    Me.ComboSubsistema.Requery
    strValorAnteriorBUI = strValorBUI
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub


Public Sub ComboNodo_BeforeUpdate(Cancel As Integer)
    Dim strValorNodo As String, strValorActualBUI As String, strSQLBUI As String, strValorActualSubSistema As String
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strValorNodo = Nz(Me.ComboNodo, "")
    If strValorNodo <> strValorAnteriorNodo Then
        Me.ComboBui = Null
        Me.ComboSubsistema = Null
    End If
    If strValorNodo = "" Then
        strSQLBUI = "SELECT distinct TbNodoBUI.BUI " & _
                    "FROM TbNodoBUI " & _
                    "WHERE (((TbNodoBUI.NODO) Like '*')) " & _
                    "ORDER BY TbNodoBUI.BUI;"
    Else
       
        strSQLBUI = "SELECT distinct TbNodoBUI.BUI " & _
                    "FROM  TbNodoBUI " & _
                    "WHERE (((TbNodoBUI.NODO)='" & strValorNodo & "')) " & _
                    "ORDER BY TbNodoBUI.BUI;"
    End If
    Me.ComboBui.RowSource = strSQLBUI
    Me.ComboBui.Requery
    strValorAnteriorNodo = strValorNodo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub

Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub Form_Load()
    Dim strNodo As String, strBUI As String, strSUBSISTEMA As String, strArgumento As String, strTextoError As String
    On Error GoTo errores
    'strArgumento=strNodo & "|" & strBUI & "|" & strSubSistema
    'pero es opcional, puede haber 2 barras o ninguna
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    
    Me.Caption = Me.lblTitulo1.Caption
    strArgumento = Nz(Me.OpenArgs, "")
    If strArgumento <> "" Then
        dato = Split(strArgumento, "|")
        If UBound(dato) <> 2 Then
            strTextoError = "Se ha abierto el formulario con parámetros insuficientes"
            Err.Raise 1000
        End If
        strNodo = Nz(dato(0), "")
        strBUI = Nz(dato(1), "")
        strSUBSISTEMA = Nz(dato(2), "")
        Me.Equipo.SetFocus
    End If
    If strNodo <> "" Then
        Me.ComboNodo = strNodo
        m_SQL = "SELECT DISTINCT TbNodoBUI.BUI " & _
                "FROM TbNodoBUI " & _
                "WHERE (((TbNodoBUI.NODO)='" & strNodo & "')) " & _
                "ORDER BY TbNodoBUI.BUI;"
        Me.ComboBui.RowSource = m_SQL
        Me.ComboBui.Requery
        Me.ComboNodo.Enabled = False
        Me.ComboNodo.Locked = True
        Me.ComboNodo.BackColor = lngColorCampoNoEditable
    Else
        Me.ComboNodo.Enabled = True
        Me.ComboNodo.Locked = False
        Me.ComboNodo.BackColor = lngColorCampoEditable
    End If
    If strBUI <> "" Then
        Me.ComboBui = strBUI
        m_SQL = "SELECT DISTINCT TbSubsistemaBui.SUBSISTEMA " & _
                    "FROM TbSubsistemaBui " & _
                    "WHERE (((TbSubsistemaBui.BUI)='" & strBUI & "')) " & _
                    "ORDER BY TbSubsistemaBui.SUBSISTEMA;"
        Me.ComboSubsistema.RowSource = m_SQL
        Me.ComboSubsistema.Requery
        Me.ComboBui.Enabled = False
        Me.ComboBui.Locked = True
        Me.ComboBui.BackColor = lngColorCampoNoEditable
    Else
        Me.ComboBui.Enabled = True
        Me.ComboBui.Locked = False
        Me.ComboBui.BackColor = lngColorCampoEditable
    End If
    If strSUBSISTEMA <> "" Then
        Me.ComboSubsistema = strSUBSISTEMA
        Me.ComboSubsistema.Enabled = False
        Me.ComboSubsistema.Locked = True
        Me.ComboSubsistema.BackColor = lngColorCampoNoEditable
    Else
        
        Me.ComboSubsistema.Enabled = True
        Me.ComboSubsistema.Locked = False
        Me.ComboSubsistema.BackColor = lngColorCampoEditable
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
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Public Function AltaIDEquipo( _
                                strNodo As String, _
                                strBUI As String, _
                                strSUBSISTEMA As String, _
                                strEquipo As String, _
                                Optional strDESCRIPCION As String _
                                ) As String
    '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 25/01/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -va a tomar todos los datos que entran y los va a meter TbEquipos--->
   
    '       -1) Verificación de datos
    '           -strNodo Existente
    '           -strBUI Existente
    '           -strSubsistema Existente
    '           -no haya un registro con mismo strNodo+strBUI+strSubsistema+strIDEquipo
    '   -Llamada desde
    '       -Form_FormEquipoAlta.ComandoAlta_Click
    '   -Devuelve:
    '       AltaIDEquipo = strIDEquipo
    '       AltaIDEquipo = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strIDEquipo As String, strTextoError As String
    On Error GoTo errores
    
    If strNodo = "" Then
        strTextoError = "Es obligatorio indicar el Nodo "
        Err.Raise 1000
    End If
    If strBUI = "" Then
        strTextoError = "Es obligatorio indicar El BUI"
        Err.Raise 1000
    End If
    If strSUBSISTEMA = "" Then
        strTextoError = "Es obligatorio indicar El SubSistema"
        Err.Raise 1000
    End If
    If strEquipo = "" Then
        strTextoError = "Es obligatorio indicar El strEquipo"
        Err.Raise 1000
    End If
    '-----------
    ' AHORA SE VA VER SI ES COMPATIBLE EL NODO,BUI,SUBSISTEMA
    '--------------
    m_SQL = "SELECT TbNodoBUI.NODO, TbNodoBUI.BUI, TbSubsistemaBui.SUBSISTEMA " & _
            "FROM TbNodoBUI INNER JOIN TbSubsistemaBui ON TbNodoBUI.BUI = TbSubsistemaBui.BUI " & _
            "WHERE (((TbNodoBUI.NODO)='" & strNodo & _
            "') AND ((TbNodoBUI.BUI)='" & strBUI & _
            "') AND ((TbSubsistemaBui.SUBSISTEMA)='" & strSUBSISTEMA & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "El subsistema " & strSUBSISTEMA & " no está registrado para el NODO " & strNodo & " y el BUI " & strBUI
            Err.Raise 1000
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    '--------
    ' ALTA DEL IDEQUIPO
    '------------------------------------------
    '   -Devuelve:
    '       DameID = CStr(lngMaxID + 1)
    '       DameID = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = DameID("TbEquipos", "IDEquipo")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameID ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strIDEquipo = flag
    If Not IsNumeric(strIDEquipo) Then
        strTextoError = "La función DameID no ha devuelto un valor numérico válido"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbEquipos.* " & _
            "FROM TbEquipos " & _
            "WHERE (((TbEquipos.IDEquipo)=" & strIDEquipo & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strTextoError = "Ya existe un IDEquipo que tiene el ID: " & strEquipo
            Err.Raise 1000
        End If
    End With
    m_SQL = "SELECT TbEquipos.* " & _
            "FROM TbEquipos " & _
            "WHERE (((TbEquipos.NODO)='" & strNodo & _
            "') AND ((TbEquipos.BUI)='" & strBUI & _
            "') AND ((TbEquipos.SUBSISTEMA)='" & strSUBSISTEMA & _
            "') AND ((TbEquipos.Equipo)='" & strEquipo & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strTextoError = "Ya existe un IDEquipo llamado: " & strEquipo & vbCrLf & " para " & strNodo & vbCrLf & strBUI & vbCrLf & strSUBSISTEMA
            Err.Raise 1000
        End If
        .AddNew
            .Fields("IDEquipo") = strIDEquipo
            .Fields("NODO") = strNodo
            .Fields("BUI") = strBUI
            .Fields("SUBSISTEMA") = strSUBSISTEMA
            .Fields("Equipo") = strEquipo
            If strDESCRIPCION <> "" And strDESCRIPCION <> "#" Then
                .Fields("Descripcion") = strDESCRIPCION
            End If
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    AltaIDEquipo = strIDEquipo
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método AltaIDEquipo ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    AltaIDEquipo = "#ERR" & "|" & strTextoError
End Function







