Option Compare Database
Option Explicit
Dim strBUI As String, strSUBSISTEMA As String, strIDEquipo As String, strIDActividad As String, strIDEvento As String
Dim strIDMaterialUltimo As String, strUltimoMaterialNoEnCombo As String
Public Sub RellenarCombosOrigenDestion()
    Dim rcdDatos As DAO.Recordset, strBUI As String, cbo As ComboBox, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Set cbo = Me.ComboOrigen
    cbo.RowSource = ""
    cbo.AddItem "Telefónica"
    m_SQL = "SELECT distinct TbSubsistemaBui.BUI " & _
                "FROM TbSubsistemaBui " & _
                "ORDER BY TbSubsistemaBui.BUI;"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                cbo.AddItem .Fields("BUI")
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Set cbo = Me.ComboDestino
    cbo.RowSource = Me.ComboOrigen.RowSource
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        strTextoError = "Al intentar rellenar el combo se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Private Function Registrar( _
                            Optional strIDActividad As String, Optional strIDEquipo As String, _
                            Optional strMaterial As String, Optional strPN As String, Optional strNS As String, _
                            Optional strFechaIncio As String, Optional strGarantia As String, _
                            Optional strUltimaIntervencion As String, Optional strFechaFin As String, Optional strDestino As String, _
                            Optional strCOSTE As String, Optional strMotivo As String, Optional strResultado As String, _
                            Optional strReparadoPor As String, Optional strTipoAccion As String _
                            ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día x
    '   -Modificaciones:
    
    '   -Funcionamiento:
    
    '   -Llamada desde:
    
    '   -Devuelve:
    '       Registrar=strSQLFinal
    '       Registrar = "ERR" & "|" &strTextoError
    '-------------------------------------------------------------------
    Dim strIDMaterial As String, strTextoError As String
    On Error GoTo errores
    If strIDActividad = "" And strIDEquipo = "" Then
        strTextoError = "Ha de ser dependiente de una actividad o de un equipo"
        Err.Raise 1000
    ElseIf strIDActividad <> "" And strIDEquipo <> "" Then
        strTextoError = "Ha de ser dependiente de una actividad o de un equipo"
        Err.Raise 1000
    End If
    If strMaterial = "" Then
        strTextoError = "El material es obligatorio"
        Err.Raise 1000
    End If
    If Not IsDate(strFechaIncio) Then
        strTextoError = "Fecha inicial obligatoria"
        Err.Raise 1000
    End If
    If strGarantia <> "Sí" And strGarantia <> "No" Then
        strGarantia = "No"
    End If
    If strUltimaIntervencion <> "Sí" And strUltimaIntervencion <> "No" Then
        strTextoError = "Se ha de indicar si es la última intervención"
        Err.Raise 1000
    End If
    If strUltimaIntervencion = "Sí" Then
        If Not IsDate(strFechaFin) Then
            strTextoError = "Si es la última intervención ha de indicarse la fecha final"
            Err.Raise 1000
        Else
            If CDate(strFechaFin) < CDate(strFechaIncio) Then
                strTextoError = "Fecha inicial mayor que final"
                Err.Raise 1000
            End If
        End If
    End If
    If strDestino = "" Then
        strTextoError = "Destino obligatorio"
        Err.Raise 1000
    End If
    If strMotivo = "" Then
        strTextoError = "Motivo obligatorio"
        Err.Raise 1000
    End If
    If IsDate(strFechaFin) Then
        If strResultado = "" Then
            strTextoError = "El Resultado es obligatorio cuando se rellena la fecha de fin"
            Err.Raise 1000
        End If
    End If
    If strUltimaIntervencion = "Sí" Then
        If strReparadoPor = "" Then
            strTextoError = "Si es la última intervención es obligatorio indicar Reparado Por"
            Err.Raise 1000
        End If
        If strTipoAccion = "" Then
            strTextoError = "Si es la última intervención es obligatorio indicar Tipo Acción"
            Err.Raise 1000
        End If
    End If
    If strIDActividad <> "" Then
        '-----------------------------------------------
        '   -Devuelve:
        '       MaterialDeActividadAlta = strIDMaterial
        '       MaterialDeActividadAlta = "#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = MaterialDeActividadAlta(strIDActividad, strMaterial, strCOSTE, _
                                                 strTipoAccion, strPN, strNS, strReparadoPor, _
                                                strMotivo, strFechaIncio, strGarantia, "Sí", strFechaFin)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método MaterialDeActividadAlta ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
    Else
        '-----------------------------------------------
        '   -Devuelve:
        '       MaterialNODeActividadAlta = strIDMaterial
        '       MaterialNODeActividadAlta = "#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = MaterialNODeActividadAlta(strIDEquipo, strMaterial, strCOSTE, strTipoAccion, strPN, strNS, strReparadoPor, _
                                                strMotivo, strFechaIncio, strGarantia, "Sí", strFechaFin)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método MaterialNODeActividadAlta ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
    End If
    strIDMaterial = flag
    '-----------------------------------------------
    '   -Devuelve:
    '       MaterialIntervencionAlta = strDescriptivo
    '       MaterialIntervencionAlta = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = MaterialIntervencionAlta(strIDMaterial, strFechaIncio, strDestino, strMotivo, strResultado, strFechaFin, strUltimaIntervencion)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método MaterialIntervencionAlta ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Registrar = strIDMaterial
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Registrar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    Registrar = "#ERR" & "|" & strTextoError
End Function
Private Sub ComandoAlta_Click()
    Dim strMaterial As String, strPN As String, strNS As String, _
        strFechaIncio As String, strGarantia As String, strUltimaIntervencion As String, strFechaFin As String, strDestino As String, _
        strCOSTE As String, strMotivo As String, strResultado As String, strReparadoPor As String, strTipoAccion As String, strMaterialExistente As String, _
            strMaterialNoexistente As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    strMaterialExistente = Nz(Me.ComboMaterial.Column(1), "")
    strMaterialNoexistente = Nz(Me.Material, "")
    strCOSTE = Nz(Me.COSTE, "")
    strTipoAccion = Nz(Me.ComboTipoAccion, "")
    strPN = Nz(Me.PN, "")
    strNS = Nz(Me.NS, "")
    strReparadoPor = Nz(Me.ComboReparadoPor, "")
    strMotivo = "Reparación"
    strFechaIncio = Nz(Me.FechaIncio, "")
    strGarantia = Nz(Me.ComboGarantia, "")
    strFechaFin = Nz(Me.FechaFinIntervencion, "")
    
    If strMaterialExistente <> "" Then
        strMaterial = strMaterialExistente
    Else
        strMaterial = strMaterialNoexistente
    End If
    
    If strIDActividad <> "" Then
        '-----------------------------------------------
        '   -Devuelve:
        '       MaterialDeActividadAlta = strIDMaterial
        '       MaterialDeActividadAlta = "#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = MaterialDeActividadAlta(strIDActividad, strMaterial, strCOSTE, _
                                                 strTipoAccion, strPN, strNS, strReparadoPor, _
                                                strMotivo, strFechaIncio, strGarantia, "Sí", strFechaFin)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método MaterialDeActividadAlta ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        '---------------------------------------------
        '   -Devuelve
        '       FormularioAbierto = true or false
        '       FormularioAbierto =False-->Error
        '---------------------------------------------
        If FormularioAbierto("FormEventoGestion") Then
            Form_FormEventoGestion.ListaActividades_Click
        End If
    Else
       '-----------------------------------------------
        '   -Devuelve:
        '       MaterialNODeActividadAlta = strIDMaterial
        '       MaterialNODeActividadAlta = "#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = MaterialNODeActividadAlta(strIDEquipo, strMaterial, strCOSTE, strTipoAccion, strPN, strNS, strReparadoPor, _
                                                strMotivo, strFechaIncio, strGarantia, "Sí", strFechaFin)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método MaterialNODeActividadAlta ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        strTextoError = "Al intentar registrar se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub cmdSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub ComandoVerActividad_Click()
    Dim strTextoError As String
    On Error GoTo errores
    If strIDActividad = "" Then
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
    If Not FormularioAbierto("FormActividadGestion") Then
        DoCmd.OpenForm "FormActividadGestion"
        Forms("FormActividadGestion").Controls("ComboIDActividadBusqueda").Value = strIDActividad
        Form_FormActividadGestion.ComboIDActividadBusqueda_Change
    Else
        Forms("FormActividadGestion").SetFocus
        If Forms("FormActividadGestion").Controls("ComboIDActividadBusqueda").Value <> strIDActividad Then
            'Form_FormActividadGestion.ComandoActualizarTodo_Click
            Forms("FormActividadGestion").Controls("ComboIDActividadBusqueda").Value = strIDActividad
            Form_FormActividadGestion.ComboIDActividadBusqueda_Change
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar ver la actividad se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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

Private Sub ComboMaterial_Change()
    Dim strTextoError As String, strValor As String, strNS As String, strPN As String
    On Error GoTo errores
    strValor = Nz(Me.ComboMaterial.Column(0), "")
    If strValor <> "" Then
        Me.Material = Null
        strNS = Nz(Me.ComboMaterial.Column(2), "")
        strPN = Nz(Me.ComboMaterial.Column(3), "")
        Me.NS = strNS
        Me.PN = strPN
    Else
        If Nz(Me.Material, "") = "" Then
            Me.PN = Null
            Me.NS = Null
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al elegir el material se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub

Private Sub Form_Load()
    Dim strArgumento As String, strTextoError As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AjustarTamaño Me
    
    Me.Caption = Me.lblTitulo1.Caption
    strArgumento = Nz(Me.OpenArgs(), "")
    If InStr(1, strArgumento, "|") = 0 Then
        strTextoError = "El formulario ha sido abierto con parámetros erróneos"
        Err.Raise 1000
    End If
    'strArgumento=strBUI & "|" & strSubSistema & "|" & strIDEquipo & "|" & strIDEquipo
    'strArgumento=strBUI & "|" & strSubSistema & "|" & strIDEquipo & "|" & strIDEquipo & "|" & strIDEvento & "|" & strIDActividad
    dato = Split(strArgumento, "|")
    If UBound(dato) <> 3 And UBound(dato) <> 5 Then
        strTextoError = "El formulario ha sido abierto con parámetros erróneos"
        Err.Raise 1000
    End If
    If UBound(dato) = 3 Then
        strBUI = dato(0)
        strSUBSISTEMA = dato(1)
        strIDEquipo = dato(2)
        strIDEquipo = dato(3)
    Else
        strBUI = dato(0)
        strSUBSISTEMA = dato(1)
        strIDEquipo = dato(2)
        strIDEquipo = dato(3)
        strIDEvento = dato(4)
        strIDActividad = dato(5)
    End If
    If strBUI = "" Or strSUBSISTEMA = "" Or strIDEquipo = "" Or strIDEquipo = "" Then
        strTextoError = "El formulario ha sido abierto con parámetros erróneos"
        Err.Raise 1000
    End If
    If UBound(dato) = 5 Then
        If strIDEvento = "" Or strIDActividad = "" Then
            strTextoError = "El formulario ha sido abierto con parámetros erróneos"
            Err.Raise 1000
        End If
        Me.IDEVENTO = strIDEvento
        Me.IDActividad = strIDActividad
        Me.ComandoVerActividad.Enabled = True
        Me.ComandoVerEvento.Enabled = False
    End If
    Me.BUI = strBUI
    Me.SubSistema = strSUBSISTEMA
    Me.IDEquipo = strIDEquipo
    Me.ComboGarantia = "No"
    m_SQL = "SELECT DISTINCT TbMaterial.IDMaterial,TbMaterial.Material, TbMaterial.PN, TbMaterial.NS " & _
            "FROM (TbMaterial INNER JOIN TbActividades ON TbMaterial.IDActividad = TbActividades.IDActividad) " & _
            "INNER JOIN TbEventos ON TbActividades.IDEvento = TbEventos.IDEvento " & _
            "WHERE ((Not (TbMaterial.Material) Is Null) And ((TbEventos.IDEquipo) =" & strIDEquipo & ")) " & _
            "ORDER BY TbMaterial.Material;"
    Me.ComboMaterial.RowSource = m_SQL
    Me.ComboMaterial.Requery
    RellenarCombosOrigenDestion
    Me.ComboOrigen = Me.BUI
    Me.ComboDestino = Me.BUI
    Me.FechaIncio = Format(Now(), "dd/mm/yyyy")
    VBA.DoEvents
    Me.Repaint
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar abrir el formulario se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
Private Sub Material_Exit(Cancel As Integer)
    Dim strValor As String
    On Error Resume Next
    strValor = Nz(Me.Material, "")
    If strValor <> strUltimoMaterialNoEnCombo And strValor <> "" Then
        Me.ComboMaterial = Null
    End If
    strUltimoMaterialNoEnCombo = strValor
End Sub
Public Function MaterialNODeActividadAlta( _
                                            strIDEquipo As String, _
                                            strMaterial As String, _
                                            Optional strCOSTE As String, _
                                            Optional strTipoAccion As String, _
                                            Optional strPN As String, _
                                            Optional strNS As String, _
                                            Optional strReparadoPor As String, _
                                            Optional strDESCRIPCION As String, _
                                            Optional strFechaInicial As String, _
                                            Optional strGarantia As String, _
                                            Optional strEsReparacion As String, _
                                            Optional strFechaEntrega As String _
                                            ) As String

     '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 10/04/13
    '   -Modificaciones:
    '       -9/7/2013: es opcional strCoste,strFechaEntrega
    '       -24/03/14: Se considera la fecha de obsolescencia en el IDEquipo. Se coteja con la fecha de strFechaAlta
    '       -11/09/2014: Se incluye la FechaInicial y la Garantía
    '       -17/09/2014: No se permite tocar la fecha de entrega. Será consecuencia de los seguimientos
    '       -06/10/2014: Se incluye un campo del tbMaterial -->EsReparacion.
    '   -Funcionamiento:
    '       -va a dar de alta un registro en la tabla TbMaterial
    '       -Nunca va a depender de una IDActividad
    '       -Si no hay coste ha de ser 0
    '       -Se considera acabado cuando está relleno el Reparadopor
   
    '   -Llamada desde
    '       -Form_FormMaterialAltaReparacion.ComandoAlta_Click
    '       -Form_FormMaterialRepuestoAlta.ComandoAlta_Click
    '   -Devuelve:
    '       MaterialNODeActividadAlta = strIDMaterial
    '       MaterialNODeActividadAlta = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strIDMaterial As String, strFechaObsoleto As String, strTextoError As String
    On Error GoTo errores
    
    If Not IsNumeric(strIDEquipo) Then
        strTextoError = "Se ha de indicar el IDEquipo al que pertenece"
        Err.Raise 1000
    End If
    If strCOSTE <> "" Then
       If Not IsNumeric(strCOSTE) Then
            strTextoError = "Se ha de indicar el coste con un valor numérico válido"
            Err.Raise 1000
        End If
    End If
    If Not IsDate(strFechaInicial) Then
        strFechaInicial = Now()
    End If
    If strGarantia <> "Sí" And strGarantia <> "No" Then
        strGarantia = "No"
    End If
    If strEsReparacion <> "Sí" And strEsReparacion <> "No" Then
        strEsReparacion = "No"
    End If
    '------------------------------------------
    '   -Devuelve:
    '       Dame = strValorObtenido
    '       Dame = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = Dame("TbEquipos", "FechaObsoleto", "IDEquipo", strIDEquipo)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strFechaObsoleto = flag
    If IsDate(strFechaObsoleto) Then
        If CDate(strFechaObsoleto) > CDate(Format(Now(), "dd/mm/yyyy")) Then
            strTextoError = "El equipo seleccionado para el evento es obsoleto desde el " & strFechaObsoleto
            Err.Raise 1000
        End If
    End If
    If strEsReparacion = "No" Then
        If strFechaEntrega <> "" Then
            If Not IsDate(strFechaEntrega) Then
                strTextoError = "El campo fecha de entrega no tiene un formato correcto"
                Err.Raise 1000
            End If
        End If
    End If
    'Los ordinales tienen 4 cifras numéricas MNA_XXXX
    '------------------------------------------
    '   -Devuelve:
    '       DameValorSiguienteAlMaximoDeUnCampo = strParteFijaInicial & Format(CStr(lngValorMaximo + 1), String(intNumeroCaracteresDesdeLaDerecha, "0"))
    '       DameValorSiguienteAlMaximoDeUnCampo = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = DameValorSiguienteAlMaximoDeUnCampo("TbMaterial", "IDMaterial", "MNA_", 4)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameValorSiguienteAlMaximoDeUnCampo ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strIDMaterial = flag
    m_SQL = "SELECT TbMaterial.* " & _
            "FROM TbMaterial " & _
            "WHERE (((TbMaterial.IDMaterial)='" & strIDMaterial & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strTextoError = "La función DameIDMaterial ha devuelto un valor de IDMaterial ya existente"
            Err.Raise 1000
        End If
        .AddNew
            .Fields("IDEquipo") = strIDEquipo
            .Fields("IDMaterial") = strIDMaterial
            .Fields("Material") = strMaterial
            .Fields("FechaIncio") = strFechaInicial
            .Fields("Garantia") = strGarantia
            If strPN <> "" Then
                .Fields("PN") = strPN
            End If
            If strNS <> "" Then
                .Fields("NS") = strNS
            End If
            If IsNumeric(strCOSTE) Then
                .Fields("Coste") = strCOSTE
            End If
            If strTipoAccion <> "" Then
                .Fields("TipoAccion") = strTipoAccion
            End If
            If strReparadoPor <> "" Then
                .Fields("ReparadoPor") = strReparadoPor
            End If
            If strDESCRIPCION <> "" Then
                .Fields("Descripcion") = strDESCRIPCION
            End If
            .Fields("EsReparacion") = strEsReparacion
            If strEsReparacion = "No" Then
                If IsDate(strFechaEntrega) Then
                    .Fields("FechaEntrega") = strFechaEntrega
                End If
            End If
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    MaterialNODeActividadAlta = strIDMaterial
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método MaterialNODeActividadAlta ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    MaterialNODeActividadAlta = "#ERR" & "|" & strTextoError
End Function
