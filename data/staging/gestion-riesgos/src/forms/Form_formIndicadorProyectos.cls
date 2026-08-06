Attribute VB_Name = "Form_formIndicadorProyectos"
Option Compare Database
Option Explicit

Private m_Error As String
Private flag As String
Private dato

Private m_Anio As Long
Private m_Semestre As String  ' "1" / "2" / "" (si anual lo decides)

Private Sub Form_Close()
    On Error Resume Next
    IndicadorState_MarcarFormAbierto False
End Sub

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo errores

    IndicadorState_Init
    IndicadorState_MarcarFormAbierto True
    
   

    ParseOpenArgs
    CargarLista_RespetandoExcluidos
    Me.cmdQuitar.Enabled = False
    Exit Sub
errores:
    MsgBox "Form_Open (formIndicadorProyectos): " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub
Private Sub CargarLista_RespetandoExcluidos()
    Dim lst As ListBox
    Dim strIDProyecto As String
    Dim strProyecto As String
    Dim strNombreProyecto As String
    Dim strFechaCierre As String
    Dim strFechaAlta As String
    Dim flag As String
    Dim dato
    Dim varItem As Variant
    Dim ColProyectosAbiertos As Collection
    On Error GoTo errores

    Set lst = Me.lstProyectos

    ' Encabezado (fila 0)
    lst.RowSource = "IDProyecto;Código;Proyecto;Cierre;Alta"
    lst.Requery

    ' Si no llegó año, fallo claro
    If m_Anio <= 0 Then Err.Raise 1000, , "Falta el año (OpenArgs ANIO=...)."

    ' Relleno ColProyectosAbiertos (tu función)
   
    Set ColProyectosAbiertos = IndicadorProyectos_GetProyectosAbiertos(CStr(m_Anio), m_Semestre, m_Error)
    If m_Error <> "" Then Err.Raise 1000
    If ColProyectosAbiertos Is Nothing Then
        Exit Sub
    End If
    ' Añadir SOLO los NO excluidos
    If ColProyectosAbiertos.Count > 0 Then
        For Each varItem In ColProyectosAbiertos
            dato = Split(varItem, "|")
            strIDProyecto = Nz(dato(0), "")
            strProyecto = Replace(Nz(dato(1), ""), ";", ":")
            strNombreProyecto = Replace(Nz(dato(2), ""), ";", ":")
            strFechaCierre = Nz(dato(3), "")
            strFechaAlta = Nz(dato(4), "") ' FechaAltaRegistro

            If IsNumeric(strIDProyecto) Then
                If Not IndicadorState_EsExcluido(CLng(strIDProyecto)) Then
                    If IsDate(strFechaCierre) Then strFechaCierre = Format(strFechaCierre, "dd/mm/yyyy")
                    If IsDate(strFechaAlta) Then strFechaAlta = Format(strFechaAlta, "dd/mm/yyyy")
                    lst.AddItem strIDProyecto & ";" & strProyecto & ";" & strNombreProyecto & ";" & strFechaCierre & ";" & strFechaAlta
                End If
            End If
        Next
    End If

    Exit Sub

errores:
    MsgBox "CargarLista_RespetandoExcluidos: " & vbCrLf & Err.Description, vbCritical
End Sub


Private Sub ParseOpenArgs()
    ' Canonical contract: "ANIO=<year>;SEM=<value>" (see
    ' modIndicadorOpenArgsParser.bas). The legacy pipe format
    ' "year|sem" is also accepted defensively.
    '
    ' On malformed input the helper raises Err 1000 with a
    ' descriptive Spanish message — no silent fallback to
    ' Year(Date), which used to mask wrong-project loads when
    ' the parent emitted the named contract but the parser
    ' only knew about pipes.
    Dim errTxt As String

    If Not ParseIndicadorOpenArgs(Nz(Me.openArgs, ""), m_Anio, m_Semestre, errTxt) Then
        Err.Raise 1000, , errTxt
    End If
End Sub



Private Sub lstProyectos_Click()
    On Error Resume Next
    Me.cmdQuitar.Enabled = (SelectedIDProyecto(Me.lstProyectos) > 0)
    
End Sub

Private Function SelectedIDProyecto(ByVal lst As ListBox) As Long
    Dim v As Variant

    SelectedIDProyecto = 0
    If lst.ListIndex < 0 Then Exit Function

    ' Columna 0 = IDProyecto (da igual si hay cabecera falsa o ColumnHeads)
    v = Nz(lst.Column(0, lst.ListIndex), "")

    If IsNumeric(v) Then
        SelectedIDProyecto = CLng(v)
    End If
End Function


Private Sub cmdQuitar_Click()
    Dim lst As ListBox
    Dim idp As Long

    On Error GoTo errores
    Set lst = Me.lstProyectos

    ' Si está en la cabecera (fila 0) o no hay selección, no se quita nada
    If lst.ListIndex <= 0 Then
        MsgBox "Selecciona un proyecto (no la cabecera).", vbExclamation
        Me.cmdQuitar.Enabled = False
        Exit Sub
    End If

    ' ID real del proyecto (columna 0 de la fila seleccionada)
    idp = CLng(lst.Column(0, lst.ListIndex))

    IndicadorState_SetExcluido idp
    lst.RemoveItem lst.ListIndex

    Me.cmdQuitar.Enabled = False
    Exit Sub

errores:
    MsgBox "Quitar: " & vbCrLf & Err.Description, vbCritical
End Sub



Private Sub cmdRecargarTodos_Click()
    On Error GoTo errores
    IndicadorState_ResetExcluidos
    Call CargarLista_RespetandoExcluidos
    Me.cmdQuitar.Enabled = False
    Exit Sub

errores:
    MsgBox "Recargar todos: " & vbCrLf & Err.Description, vbCritical
End Sub


Private Sub cmdAceptar_Click()
    On Error GoTo errores

    IndicadorState_Init
    IndicadorState_SetSeleccionadosDesdeLista Me.lstProyectos
    IndicadorState_MarcarConfirmado True

    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub

errores:
    MsgBox "Aceptar: " & vbCrLf & Err.Description, vbCritical
End Sub



Private Sub cmdCancelar_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
