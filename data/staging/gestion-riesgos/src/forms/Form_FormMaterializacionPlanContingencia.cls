Option Compare Database
Option Explicit

Private m_IdRiesgo As String
Private m_IDMaterializacion As String
Private m_IDPlanInicial As String
Private m_EsModoEdicion As Boolean
Private m_Error As String
Private Const TEMP_MAT_ACEPTADA As String = "MaterializacionPlanContingenciaAceptada"
Private Const TEMP_MAT_FECHA As String = "MaterializacionPlanContingenciaFecha"
Private Const TEMP_MAT_PLAN As String = "MaterializacionPlanContingenciaPlan"
Private Const OPENARGS_EDIT_MATERIALIZACION As String = "EDIT"

Private Sub Form_Load()
    On Error GoTo errores

    m_Error = ""
    ResolverArgumentosApertura
    If m_IdRiesgo = "" Then
        m_Error = "No se ha podido determinar el riesgo a materializar"
        Err.Raise 1000
    End If

    If m_EsModoEdicion Then
        Me.Caption = "Editar materialización"
    Else
        Me.txtFechaMaterializacion = Date
    End If
    CargarPlanesContingenciaElegibles m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    Exit Sub

errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load de FormMaterializacionPlanContingencia se ha producido el error: " & _
            Err.description
    End If
    MsgBox m_Error, vbExclamation, "Materialización"
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub btnAceptar_Click()
    Dim m_FechaMaterializacion As String
    Dim m_IDPlanContingencia As String

    On Error GoTo errores

    m_Error = ""
    m_FechaMaterializacion = Trim$(Nz(Me.txtFechaMaterializacion, ""))
    m_IDPlanContingencia = Trim$(Nz(Me.cmbPlanContingencia, ""))

    If Not IsDate(m_FechaMaterializacion) Then
        m_Error = "Debe indicar una fecha de materialización válida"
        Err.Raise 1000
    End If
    If m_IDPlanContingencia <> "" Then
        If ValidarYResolverPlanContingenciaMaterializacion( _
            m_IdRiesgo, _
            m_IDPlanContingencia, _
            m_Error) <> EnumSiNo.Sí Then
            Err.Raise 1000
        End If
    End If

    If m_EsModoEdicion Then
        If ActualizarMaterializacionVigentePlan( _
            m_IdRiesgo, _
            m_IDMaterializacion, _
            m_FechaMaterializacion, _
            m_IDPlanContingencia, _
            m_Error) <> EnumSiNo.Sí Then
            Err.Raise 1000
        End If
    End If

    SetTempVarMaterializacion TEMP_MAT_ACEPTADA, "Sí"
    SetTempVarMaterializacion TEMP_MAT_FECHA, CStr(CDate(m_FechaMaterializacion))
    SetTempVarMaterializacion TEMP_MAT_PLAN, m_IDPlanContingencia
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub

errores:
    If Err.Number <> 1000 Then
        m_Error = "Al btnAceptar_Click se ha producido el error: " & Err.description
    End If
    MsgBox m_Error, vbExclamation, "Materialización"
End Sub

Private Sub ResolverArgumentosApertura()
    Dim rawArgs As String
    Dim parts As Variant

    rawArgs = Trim$(Nz(Me.openArgs, ""))
    If rawArgs = "" Then Exit Sub

    parts = Split(rawArgs, "|")
    If UBound(parts) < 0 Then Exit Sub
    If StrComp(CStr(parts(0)), OPENARGS_EDIT_MATERIALIZACION, vbTextCompare) <> 0 Then
        m_IdRiesgo = rawArgs
        Exit Sub
    End If

    m_EsModoEdicion = True
    If UBound(parts) >= 1 Then m_IdRiesgo = Trim$(CStr(parts(1)))
    If UBound(parts) >= 2 Then m_IDMaterializacion = Trim$(CStr(parts(2)))
    If UBound(parts) >= 3 Then Me.txtFechaMaterializacion = Trim$(CStr(parts(3)))
    If UBound(parts) >= 4 Then m_IDPlanInicial = Trim$(CStr(parts(4)))
End Sub

Private Sub btnCancelar_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub CargarPlanesContingenciaElegibles(Optional ByRef p_Error As String)
    Dim m_SQL As String

    On Error GoTo errores

    p_Error = ""
    m_SQL = SQLPlanesContingenciaElegiblesMaterializacion(m_IdRiesgo, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    With Me.cmbPlanContingencia
        .RowSourceType = "Table/Query"
        .RowSource = m_SQL
        .ColumnCount = 3
        .BoundColumn = 1
        .ColumnWidths = "0;1814;4535"
        .Requery
        If m_EsModoEdicion And m_IDPlanInicial <> "" Then
            If ComboTieneValor(Me.cmbPlanContingencia, m_IDPlanInicial) Then
                .value = m_IDPlanInicial
                Exit Sub
            End If
            p_Error = "El plan de contingencia asociado ya no es elegible"
            Err.Raise 1000
        End If
        If .ListCount = 0 Then
            p_Error = "El riesgo no tiene planes de contingencia elegibles"
            Err.Raise 1000
        End If
        If .ListCount = 1 Then
            .value = .ItemData(0)
        End If
    End With
    Exit Sub

errores:
    If p_Error = "" Then
        p_Error = "El método CargarPlanesContingenciaElegibles ha devuelto el error: " & _
            Err.description
    End If
End Sub

Private Function ComboTieneValor( _
    ByVal p_Combo As ComboBox, _
    ByVal p_Valor As String _
) As Boolean
    Dim index As Long

    For index = 0 To p_Combo.ListCount - 1
        If StrComp(CStr(p_Combo.ItemData(index)), p_Valor, vbTextCompare) = 0 Then
            ComboTieneValor = True
            Exit Function
        End If
    Next index
End Function

Private Sub SetTempVarMaterializacion( _
    ByVal p_Nombre As String, _
    ByVal p_Valor As String _
)
    On Error Resume Next
    TempVars.Remove p_Nombre
    On Error GoTo 0
    TempVars.Add p_Nombre, p_Valor
End Sub
