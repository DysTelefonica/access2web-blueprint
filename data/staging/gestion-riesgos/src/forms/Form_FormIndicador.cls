Attribute VB_Name = "Form_FormIndicador"
Option Compare Database
Option Explicit
Private strTextoError As String
Private m_IdsCsvUltimoCalculo As String

Private Sub btnSalir_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo errores
    AjustarTamaño Me
    ' Defaults
    Me.Caption = Me.lblTitulo.Caption
    m_IdsCsvUltimoCalculo = ""
    IndicadorState_ResetAll
    Call ResetUI_Indicadores
    Me.cboSemestre.RowSource = "S1;S2;Anual"
    Me.cboSemestre = "S1"
    Me.txtAnio = Format$(Year(Date), "0000")

    ' Estado inicial
    Call LimpiarTempVarsIndicador
    Me.lblProyectosCargados.Caption = "Proyectos: (ninguno)"
    Me.cmdCalcular.Enabled = False

    ' Valores
    Me.lblTile1Valor.Caption = "0"
    Me.lblTile2Valor.Caption = "0"
    Me.lblTile3Valor.Caption = "0"
    Me.lblTile4Valor.Caption = "0"
    Me.lblTile5Valor.Caption = "0"
    Me.lblTile6Valor.Caption = "0"
    Me.cmdCalcular.Enabled = False
    ActualizarBotonesExport
    Call RefrescarEstadoProyectos
    Exit Sub
errores:
    MsgBox "Form_Open: " & Err.Description, vbCritical
End Sub

Private Sub cmdCargarProyectos_Click()
    Dim anio As Long
    Dim sem As String
    Dim openArgs As String

    On Error GoTo errores

    ' Validaciones básicas
    If Not IsNumeric(Nz(Me.txtAnio, "")) Then
        MsgBox "Indica un año válido.", vbExclamation
        Exit Sub
    End If
    anio = CLng(Me.txtAnio)

    sem = Nz(Me.cboSemestre, "S1") ' "S1" / "S2" / "Anual"
    Select Case UCase$(sem)
        Case "S1", "1": sem = "1"
        Case "S2", "2": sem = "2"
        Case "ANUAL", "A", "": sem = ""
        Case Else
            MsgBox "Semestre inválido. Usa S1, S2 o Anual.", vbExclamation
            Exit Sub
    End Select

    IndicadorState_Init
    IndicadorState_MarcarFormAbierto True

    ' OpenArgs: "ANIO=2025;SEM=1"
    openArgs = "ANIO=" & CStr(anio) & ";SEM=" & sem

    DoCmd.OpenForm "FormIndicadorProyectos", acNormal, , , acFormEdit, acDialog, openArgs

    ' Al volver del diálogo, refrescar estado
    RefrescarEstadoProyectos
    Exit Sub

errores:
    MsgBox "Cargar proyectos: " & vbCrLf & Err.Description, vbCritical
End Sub




Private Sub cmdCalcular_Click()
    Dim idsCsv As String
    Dim anio As Long
    Dim sem As String
    Dim dIni As Date, dFin As Date
    Dim rs As DAO.Recordset
    Dim errTxt As String

    On Error GoTo errores

    ' Asegura diccionarios creados
    IndicadorState_Init

    ' Deben estar confirmados (vienen del diálogo)
    If Not IndicadorState_HaConfirmado() Then
        MsgBox "Primero debes cargar proyectos.", vbExclamation
        Exit Sub
    End If

    ' IDs seleccionados (1 sola vez) + cache para export
    errTxt = ""
    idsCsv = IndicadorState_GetIdsCsv(errTxt)
    If errTxt <> "" Or Len(Trim$(idsCsv)) = 0 Then
        MsgBox IIf(errTxt <> "", errTxt, "No hay proyectos seleccionados."), vbExclamation
        Exit Sub
    End If
    m_IdsCsvUltimoCalculo = idsCsv

    ' Año / semestre
    If Not IsNumeric(Nz(Me.txtAnio, "")) Then
        MsgBox "Indica un año válido.", vbExclamation
        Exit Sub
    End If

    anio = CLng(Me.txtAnio)
    sem = Nz(Me.cboSemestre, "S1")

    errTxt = ""
    If Not CalcularRangoSemestre(sem, anio, dIni, dFin, errTxt) Then
        MsgBox errTxt, vbExclamation
        Exit Sub
    End If

    ' Resumen
    errTxt = ""
    Set rs = IndicadorRiesgosV2_GetResumen(dIni, dFin, idsCsv, errTxt)
    If errTxt <> "" Or rs Is Nothing Then Err.Raise 1000

    If Not rs.EOF Then
        Me.lblTile1Valor.Caption = CStr(Nz(rs!RiesgosIdentificados, 0))
        Me.lblTile2Valor.Caption = CStr(Nz(rs!RiesgosRetirados, 0))
        Me.lblTile3Valor.Caption = CStr(Nz(rs!RiesgosEnOferta, 0))
        Me.lblTile4Valor.Caption = CStr(Nz(rs!RiesgosMaterializados, 0))
        Me.lblTile5Valor.Caption = CStr(Nz(rs!RiesgosOfertaPasanGestion, 0))
        Me.lblTile6Valor.Caption = CStr(Nz(rs!RiesgosVigentesEnPeriodo, 0))
    Else
        ' Por si acaso
        Me.lblTile1Valor.Caption = "0"
        Me.lblTile2Valor.Caption = "0"
        Me.lblTile3Valor.Caption = "0"
        Me.lblTile4Valor.Caption = "0"
        Me.lblTile5Valor.Caption = "0"
        Me.lblTile6Valor.Caption = "0"
    End If

    rs.Close
    Set rs = Nothing

    ' Habilita export según valores > 0
    ActualizarBotonesExport

    Exit Sub

errores:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    MsgBox "Calcular: " & vbCrLf & IIf(errTxt <> "", errTxt, Err.Description), vbCritical
End Sub



'========================
' EXPORTS (botones XLS)
'========================
Private Sub cmdTile1Excel_Click(): ExportarTile itIdentificados, "RiesgosIdentificados": End Sub
Private Sub cmdTile2Excel_Click(): ExportarTile itRetirados, "RiesgosRetirados": End Sub
Private Sub cmdTile3Excel_Click(): ExportarTile itEnOferta, "RiesgosEnOferta": End Sub
Private Sub cmdTile4Excel_Click(): ExportarTile itMaterializados, "RiesgosMaterializados": End Sub
Private Sub cmdTile5Excel_Click(): ExportarTile itOfertaTrasladar, "OfertaTrasladar": End Sub
Private Sub cmdTile6Excel_Click(): ExportarTile itVigentesEnPeriodo, "NumeroRiesgosVigentesTotales": End Sub



Private Sub cmdAyuda_Click()
    If Not fso.FileExists(m_ObjEntorno.URLAyudaIndicadores) Then
        MsgBox "Pendiente de Integrar", vbCritical, "Sin Ayuda"
        Exit Sub
    End If
    Ejecutar Me.hWnd, "open", m_ObjEntorno.URLAyudaIndicadores, "", "", 1
End Sub

Private Sub CmdSalir_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub RefrescarEstadoProyectos()
    Dim n As Long

    IndicadorState_Init
    n = IndicadorState_CountSeleccionados()

    If IndicadorState_HaConfirmado() And n > 0 Then
        Me.lblProyectosCargados.Caption = "Proyectos: (" & CStr(n) & ")"
        Me.cmdCalcular.Enabled = True
    Else
        Me.lblProyectosCargados.Caption = "Proyectos: (ninguno)"
        Me.cmdCalcular.Enabled = False
    End If
End Sub
Private Sub ResetUI_Indicadores()
    On Error Resume Next

    ' Estado proyectos
    Me.lblProyectosCargados.Caption = "Proyectos: (ninguno)"

    ' Bloquear calcular hasta que se pase por el diálogo y se acepte
    Me.cmdCalcular.Enabled = False

    ' Poner los tiles a cero
    Me.lblTile1Valor.Caption = "0"
    Me.lblTile2Valor.Caption = "0"
    Me.lblTile3Valor.Caption = "0"
    Me.lblTile4Valor.Caption = "0"
    Me.lblTile5Valor.Caption = "0"
    Me.lblTile6Valor.Caption = "0"

    ' Deshabilitar exportaciones hasta calcular (si tienes esos botones)
    Me.cmdTile1Excel.Enabled = False
    Me.cmdTile2Excel.Enabled = False
    Me.cmdTile3Excel.Enabled = False
    Me.cmdTile4Excel.Enabled = False
    Me.cmdTile5Excel.Enabled = False
    Me.cmdTile6Excel.Enabled = False
End Sub

Private Function ToLongSafe(ByVal v As Variant) As Long
    Dim s As String, i As Long, ch As String, digits As String

    s = Nz(v, "0")
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then digits = digits & ch
    Next

    If Len(digits) = 0 Then
        ToLongSafe = 0
    Else
        ToLongSafe = CLng(digits)
    End If
End Function
'------------------------------------------------------------
' Habilita / deshabilita los botones de export según los números
' AJUSTA nombres de controles a los tuyos
'------------------------------------------------------------
Private Sub ActualizarBotonesExport()
    Dim nId As Long, nRet As Long, nOfe As Long, nMat As Long, nOfeFin As Long, nVig As Long

    ' OJO: usa aquí los controles donde pintas los números (labels o textboxes)
    nId = ToLongSafe(Me.lblTile1Valor.Caption)
    nRet = ToLongSafe(Me.lblTile2Valor.Caption)
    nOfe = ToLongSafe(Me.lblTile3Valor.Caption)
    nMat = ToLongSafe(Me.lblTile4Valor.Caption)
    nOfeFin = ToLongSafe(Me.lblTile5Valor.Caption)
    nVig = ToLongSafe(Me.lblTile6Valor.Caption)

    ' OJO: usa aquí los nombres reales de tus botones de Excel
    Me.cmdTile1Excel.Enabled = (nId > 0)
    Me.cmdTile2Excel.Enabled = (nRet > 0)
    Me.cmdTile3Excel.Enabled = (nOfe > 0)
    Me.cmdTile4Excel.Enabled = (nMat > 0)
    Me.cmdTile5Excel.Enabled = (nOfeFin > 0)
    Me.cmdTile6Excel.Enabled = (nVig > 0)
End Sub
Private Sub ExportarTile(ByVal Tipo As IndicadorTile, ByVal Nombre As String)
    Dim idsCsv As String
    Dim anio As Long
    Dim sem As String
    Dim dIni As Date, dFin As Date
    Dim sql As String
    Dim ok As Boolean

    On Error GoTo errores

    ' 1) Usa los IDs del último cálculo
    idsCsv = Nz(m_IdsCsvUltimoCalculo, "")
    If Len(Trim$(idsCsv)) = 0 Then
        ' fallback por si intentan exportar sin calcular
        idsCsv = IndicadorState_GetIdsCsv(strTextoError)
        If strTextoError <> "" Then
            MsgBox strTextoError, vbExclamation
            Exit Sub
        End If
    End If

    If Not IsNumeric(Nz(Me.txtAnio, "")) Then
        MsgBox "Indica un año válido.", vbExclamation
        Exit Sub
    End If

    anio = CLng(Me.txtAnio)
    sem = Nz(Me.cboSemestre, "S1")

    If Not CalcularRangoSemestre(sem, anio, dIni, dFin, strTextoError) Then
        MsgBox strTextoError, vbExclamation
        Exit Sub
    End If

    sql = IndicadorRiesgosV2_SqlDetalle(Tipo, dIni, dFin, idsCsv)
    If Len(sql) = 0 Then
        MsgBox "No hay SQL para este indicador.", vbExclamation
        Exit Sub
    End If

    ok = ExportarSQLaExcel(sql, Nombre, strTextoError)
    If Not ok Then MsgBox strTextoError, vbExclamation
    Exit Sub

errores:
    MsgBox "Exportar: " & vbCrLf & IIf(strTextoError <> "", strTextoError, Err.Description), vbCritical
End Sub