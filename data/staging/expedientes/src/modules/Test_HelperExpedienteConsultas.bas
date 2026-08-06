Attribute VB_Name = "Test_HelperExpedienteConsultas"
Option Compare Database
Option Explicit

Public Function Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_CamposVacios_DevuelveTrueTautologias() As String
    Dim logs(0 To 2) As String
    Dim busqueda As ExpedienteBusqueda
    Set busqueda = New ExpedienteBusqueda
    logs(0) = "Arrange: blank ExpedienteBusqueda"
    Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_CamposVacios_DevuelveTrueTautologias = AssertWhereContains(busqueda, "(Estado Is Null or Not Estado Is Null)", logs)
End Function

Public Function Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_FiltraPorID_DevuelveWhere() As String
    Dim logs(0 To 2) As String
    Dim busqueda As ExpedienteBusqueda
    Set busqueda = New ExpedienteBusqueda
    busqueda.CodExp = "EXP-001"
    logs(0) = "Arrange: CodExp filter set"
    Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_FiltraPorID_DevuelveWhere = AssertWhereContains(busqueda, "CodExp='EXP-001'", logs)
End Function

Public Function Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_FiltraPorPalabraClave_DevuelveLike() As String
    Dim logs(0 To 2) As String
    Dim busqueda As ExpedienteBusqueda
    Set busqueda = New ExpedienteBusqueda
    busqueda.CodExp = "EXP-001"
    busqueda.PalabraClave = "PRUEBA"
    logs(0) = "Arrange: CodExp and keyword filters set"
    Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_FiltraPorPalabraClave_DevuelveLike = AssertWhereContains(busqueda, "(Nemotecnico Like '*PRUEBA*' or CodExp Like '*PRUEBA*')", logs)
End Function

Public Function Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_PECALYRAC_DevuelveAndEncadenado() As String
    Dim logs(0 To 2) As String
    Dim busqueda As ExpedienteBusqueda
    Set busqueda = New ExpedienteBusqueda
    busqueda.PECAL = "P1"
    busqueda.RAC = "R1"
    logs(0) = "Arrange: PECAL and RAC filters set"
    Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_PECALYRAC_DevuelveAndEncadenado = AssertWhereContains(busqueda, "TbExpedientesConEntidades.CadenaPecal Like '*P1*'", logs, "CadenaRACs Like '*R1*'")
End Function

Public Function Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_AmbitoDefensa_DevuelveAmbitoSi() As String
    Dim logs(0 To 2) As String
    Dim busqueda As ExpedienteBusqueda
    Set busqueda = New ExpedienteBusqueda
    busqueda.m_EnumAmbito = EnumAmbito.Defensa
    logs(0) = "Arrange: Ambito Defensa set"
    Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_AmbitoDefensa_DevuelveAmbitoSi = AssertWhereContains(busqueda, "Ambito='Sí'", logs)
End Function

Public Function Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_EstadoLargo_PueblaError() As String
    Dim logs(0 To 2) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_EstadoLargo_PueblaError = BuildJsonFail("test did not complete", logs)

    Dim errMsg As String
    Dim busqueda As ExpedienteBusqueda
    Dim whereSql As String

    On Error GoTo HandleError
    Set busqueda = New ExpedienteBusqueda
    busqueda.ESTADO = String$(201, "X")
    logs(0) = "Arrange: overlong ESTADO set"
    whereSql = Helper_ExpedienteConsultas.ConstruirWhereBusqueda(busqueda, errMsg)
    logs(1) = "Act: ConstruirWhereBusqueda executed"
    If errMsg = "" Then
        Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_EstadoLargo_PueblaError = BuildJsonFail("expected p_Error for overlong ESTADO", logs)
        Exit Function
    End If
    If whereSql <> "" Then
        Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_EstadoLargo_PueblaError = BuildJsonFail("expected empty return for overlong ESTADO", logs)
        Exit Function
    End If
    logs(2) = "Assert: overlong ESTADO populates p_Error"
    Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_EstadoLargo_PueblaError = BuildJsonOk("ok", logs)
    Exit Function
HandleError:
    Test_Helper_ExpedienteConsultas_ConstruirWhereBusqueda_EstadoLargo_PueblaError = BuildJsonFail(Err.Description, logs)
End Function

Private Function AssertWhereContains( _
    ByRef p_Busqueda As ExpedienteBusqueda, _
    ByVal p_Expected As String, _
    ByRef p_Logs() As String, _
    Optional ByVal p_Expected2 As String = "") As String

    Dim errMsg As String
    Dim whereSql As String

    On Error GoTo EH
    whereSql = Helper_ExpedienteConsultas.ConstruirWhereBusqueda(p_Busqueda, errMsg)
    p_Logs(1) = "Act: ConstruirWhereBusqueda executed"
    If errMsg <> "" Then
        AssertWhereContains = BuildJsonFail(errMsg, p_Logs)
        Exit Function
    End If
    If Left$(whereSql, 6) <> "WHERE " Or Right$(whereSql, 1) <> ";" Then
        AssertWhereContains = BuildJsonFail("invalid WHERE envelope", p_Logs)
        Exit Function
    End If
    If InStr(1, whereSql, p_Expected, vbTextCompare) = 0 Then
        AssertWhereContains = BuildJsonFail("expected predicate missing: " & p_Expected, p_Logs)
        Exit Function
    End If
    If p_Expected2 <> "" Then
        If InStr(1, whereSql, p_Expected2, vbTextCompare) = 0 Then
            AssertWhereContains = BuildJsonFail("expected predicate missing: " & p_Expected2, p_Logs)
            Exit Function
        End If
    End If
    p_Logs(2) = "Assert: expected predicate is present"
    AssertWhereContains = BuildJsonOk("ok", p_Logs)
    Exit Function
EH:
    AssertWhereContains = BuildJsonFail(Err.Description, p_Logs)
End Function

' ============================================================================
' Tests de renderizado (PRUEBA-003 REFAC-1a — T1a.7+)
' Cubre GenerarFilasAM/Lote/Basado + ContarExpedientes.
' Stateless, no requieren DB — los inputs son Dictionary + Strings.
' ============================================================================

' Helper: crea un ExpedienteCompleto con campos para los tests de render.
' No pisa Err.Number del caller; usa m_Err local.
Private Function BuildExpedienteCompleto( _
    ByVal p_IDExpediente As String, _
    ByVal p_CodExp As String, _
    ByVal p_TipoParaLista As String, _
    ByVal p_Nemotecnico As String, _
    ByVal p_ESTADO As String, _
    ByVal p_Ordinal As String, _
    ByVal p_IDExpedientePadre As String, _
    ByVal p_FechaInicioContrato As String, _
    ByVal p_FechaFinContrato As String, _
    ByVal p_Ejercito As String) As ExpedienteCompleto
    Dim m_Exp As New ExpedienteCompleto
    Dim m_Err As String
    m_Exp.SetPropiedad "IDExpediente", p_IDExpediente, m_Err
    m_Exp.SetPropiedad "CodExp", p_CodExp, m_Err
    m_Exp.SetPropiedad "TipoParaLista", p_TipoParaLista, m_Err
    m_Exp.SetPropiedad "Nemotecnico", p_Nemotecnico, m_Err
    m_Exp.SetPropiedad "ESTADO", p_ESTADO, m_Err
    m_Exp.SetPropiedad "Ordinal", p_Ordinal, m_Err
    m_Exp.SetPropiedad "IDExpedientePadre", p_IDExpedientePadre, m_Err
    m_Exp.SetPropiedad "FechaInicioContrato", p_FechaInicioContrato, m_Err
    m_Exp.SetPropiedad "FechaFinContrato", p_FechaFinContrato, m_Err
    m_Exp.SetPropiedad "Ejercito", p_Ejercito, m_Err
    Set BuildExpedienteCompleto = m_Exp
End Function

' BR-23-RT-01: colección vacía → array vacío
Public Function Test_Helper_ExpedienteConsultas_GenerarFilasAM_CollectionVacia_DevuelveArrayVacio() As String
    Dim logs(0 To 2) As String
    Dim colVacia As Scripting.Dictionary
    Dim errMsg As String
    Dim result As Variant
    Set colVacia = New Scripting.Dictionary
    colVacia.CompareMode = TextCompare
    logs(0) = "Arrange: empty Dictionary"
    result = Helper_ExpedienteConsultas.GenerarFilasAM(colVacia, "Sí", errMsg)
    logs(1) = "Act: GenerarFilasAM executed"
    If errMsg <> "" Then
        Test_Helper_ExpedienteConsultas_GenerarFilasAM_CollectionVacia_DevuelveArrayVacio = BuildJsonFail(errMsg, logs)
        Exit Function
    End If
    If Not IsArray(result) Then
        Test_Helper_ExpedienteConsultas_GenerarFilasAM_CollectionVacia_DevuelveArrayVacio = BuildJsonFail("expected array, got " & TypeName(result), logs)
        Exit Function
    End If
    If UBound(result) < 0 Then
        logs(2) = "Assert: empty array (UBound<0)"
        Test_Helper_ExpedienteConsultas_GenerarFilasAM_CollectionVacia_DevuelveArrayVacio = BuildJsonOk("ok", logs)
    Else
        Test_Helper_ExpedienteConsultas_GenerarFilasAM_CollectionVacia_DevuelveArrayVacio = BuildJsonFail("expected empty array", logs)
    End If
End Function

' BR-23-RT-02: MostrarEstado="Sí" → fila termina en Estado (no FInicial/FFinal)
Public Function Test_Helper_ExpedienteConsultas_GenerarFilasAM_MostrarEstadoSi_DevuelveFilaConEstado() As String
    Dim logs(0 To 2) As String
    Dim col As Scripting.Dictionary
    Dim errMsg As String
    Dim result As Variant
    Set col = New Scripting.Dictionary
    col.CompareMode = TextCompare
    col.Add "EXP-001", BuildExpedienteCompleto("EXP-001", "EXP-COD-1", "AM", "EXP-001", "Vigente", "01", "", "01/01/2024", "31/12/2024", "Tierra")
    logs(0) = "Arrange: 1 expediente ESTADO=Vigente"
    result = Helper_ExpedienteConsultas.GenerarFilasAM(col, "Sí", errMsg)
    logs(1) = "Act: GenerarFilasAM MostrarEstado=Sí"
    If errMsg <> "" Then
        Test_Helper_ExpedienteConsultas_GenerarFilasAM_MostrarEstadoSi_DevuelveFilaConEstado = BuildJsonFail(errMsg, logs)
        Exit Function
    End If
    If UBound(result) <> 0 Then
        Test_Helper_ExpedienteConsultas_GenerarFilasAM_MostrarEstadoSi_DevuelveFilaConEstado = BuildJsonFail("expected 1 row", logs)
        Exit Function
    End If
    If InStr(result(0), ";Vigente") = 0 Then
        Test_Helper_ExpedienteConsultas_GenerarFilasAM_MostrarEstadoSi_DevuelveFilaConEstado = BuildJsonFail("expected 'Estado=Vigente' suffix, got: " & result(0), logs)
        Exit Function
    End If
    logs(2) = "Assert: row ends with Estado='Vigente'"
    Test_Helper_ExpedienteConsultas_GenerarFilasAM_MostrarEstadoSi_DevuelveFilaConEstado = BuildJsonOk("ok", logs)
End Function

' BR-23-RT-03: MostrarEstado="No" → fila termina con FInicial;FFinal
Public Function Test_Helper_ExpedienteConsultas_GenerarFilasAM_MostrarEstadoNo_DevuelveFilaConFInicialFFinal() As String
    Dim logs(0 To 2) As String
    Dim col As Scripting.Dictionary
    Dim errMsg As String
    Dim result As Variant
    Set col = New Scripting.Dictionary
    col.CompareMode = TextCompare
    col.Add "EXP-001", BuildExpedienteCompleto("EXP-001", "EXP-COD-1", "AM", "EXP-001", "Vigente", "01", "", "01/01/2024", "31/12/2024", "Tierra")
    logs(0) = "Arrange: 1 expediente"
    result = Helper_ExpedienteConsultas.GenerarFilasAM(col, "No", errMsg)
    logs(1) = "Act: GenerarFilasAM MostrarEstado=No"
    If errMsg <> "" Then
        Test_Helper_ExpedienteConsultas_GenerarFilasAM_MostrarEstadoNo_DevuelveFilaConFInicialFFinal = BuildJsonFail(errMsg, logs)
        Exit Function
    End If
    If InStr(result(0), "01/01/2024") = 0 Or InStr(result(0), "31/12/2024") = 0 Then
        Test_Helper_ExpedienteConsultas_GenerarFilasAM_MostrarEstadoNo_DevuelveFilaConFInicialFFinal = BuildJsonFail("expected FInicial/FFinal, got: " & result(0), logs)
        Exit Function
    End If
    If InStr(result(0), ";Vigente;") > 0 Then
        Test_Helper_ExpedienteConsultas_GenerarFilasAM_MostrarEstadoNo_DevuelveFilaConFInicialFFinal = BuildJsonFail("Estado should NOT appear when MostrarEstado=No, got: " & result(0), logs)
        Exit Function
    End If
    logs(2) = "Assert: row ends with FInicial;FFinal, no Estado"
    Test_Helper_ExpedienteConsultas_GenerarFilasAM_MostrarEstadoNo_DevuelveFilaConFInicialFFinal = BuildJsonOk("ok", logs)
End Function

' BR-23-RT-04: filtro por p_IdPadre excluye lotes que no son del padre
Public Function Test_Helper_ExpedienteConsultas_GenerarFilasLote_ConFiltroPadre_DevuelveSoloLotesDelPadre() As String
    Dim logs(0 To 2) As String
    Dim col As Scripting.Dictionary
    Dim errMsg As String
    Dim result As Variant
    Set col = New Scripting.Dictionary
    col.CompareMode = TextCompare
    col.Add "L1", BuildExpedienteCompleto("L1", "COD-1", "Lote", "BASE_01", "Vigente", "01", "PADRE-1", "01/01/2024", "31/12/2024", "Tierra")
    col.Add "L2", BuildExpedienteCompleto("L2", "COD-2", "Lote", "BASE_02", "Vigente", "02", "PADRE-1", "01/02/2024", "28/02/2024", "Tierra")
    col.Add "L3", BuildExpedienteCompleto("L3", "COD-3", "Lote", "OTHER_01", "Vigente", "01", "PADRE-2", "01/03/2024", "31/03/2024", "Aire")
    logs(0) = "Arrange: 3 lotes, 2 de PADRE-1, 1 de PADRE-2"
    result = Helper_ExpedienteConsultas.GenerarFilasLote(col, "PADRE-1", "", "Sí", errMsg)
    logs(1) = "Act: GenerarFilasLote filtroPadre=PADRE-1"
    If errMsg <> "" Then
        Test_Helper_ExpedienteConsultas_GenerarFilasLote_ConFiltroPadre_DevuelveSoloLotesDelPadre = BuildJsonFail(errMsg, logs)
        Exit Function
    End If
    If UBound(result) <> 1 Then
        Test_Helper_ExpedienteConsultas_GenerarFilasLote_ConFiltroPadre_DevuelveSoloLotesDelPadre = BuildJsonFail("expected 2 rows (UBound=1), got " & UBound(result) + 1, logs)
        Exit Function
    End If
    If InStr(result(0), "L1;") = 0 Or InStr(result(1), "L2;") = 0 Then
        Test_Helper_ExpedienteConsultas_GenerarFilasLote_ConFiltroPadre_DevuelveSoloLotesDelPadre = BuildJsonFail("expected L1 and L2, got: " & result(0) & " | " & result(1), logs)
        Exit Function
    End If
    logs(2) = "Assert: 2 rows de PADRE-1, PADRE-2 excluido"
    Test_Helper_ExpedienteConsultas_GenerarFilasLote_ConFiltroPadre_DevuelveSoloLotesDelPadre = BuildJsonOk("ok", logs)
End Function

' BR-23-RT-05: Nemotecnico con "_" → trunca en la primera "_" (formato "BASE_ordinal")
Public Function Test_Helper_ExpedienteConsultas_GenerarFilasLote_NemotecnicoConUnderscore_DevuelvePrimeraParte() As String
    Dim logs(0 To 2) As String
    Dim col As Scripting.Dictionary
    Dim errMsg As String
    Dim result As Variant
    Set col = New Scripting.Dictionary
    col.CompareMode = TextCompare
    col.Add "L1", BuildExpedienteCompleto("L1", "COD-1", "Lote", "BASE_NAME_01", "Vigente", "01", "P-1", "01/01/2024", "31/12/2024", "Tierra")
    logs(0) = "Arrange: Nemotecnico='BASE_NAME_01'"
    result = Helper_ExpedienteConsultas.GenerarFilasLote(col, "P-1", "", "Sí", errMsg)
    logs(1) = "Act: GenerarFilasLote"
    If errMsg <> "" Then
        Test_Helper_ExpedienteConsultas_GenerarFilasLote_NemotecnicoConUnderscore_DevuelvePrimeraParte = BuildJsonFail(errMsg, logs)
        Exit Function
    End If
    ' La fila es "L1;01;BASE;Vigente" (truncado en primer "_")
    If InStr(result(0), ";BASE;") = 0 Then
        Test_Helper_ExpedienteConsultas_GenerarFilasLote_NemotecnicoConUnderscore_DevuelvePrimeraParte = BuildJsonFail("expected ';BASE;' (truncated at first _), got: " & result(0), logs)
        Exit Function
    End If
    If InStr(result(0), "NAME_01") > 0 Then
        Test_Helper_ExpedienteConsultas_GenerarFilasLote_NemotecnicoConUnderscore_DevuelvePrimeraParte = BuildJsonFail("should NOT contain 'NAME_01' after truncation, got: " & result(0), logs)
        Exit Function
    End If
    logs(2) = "Assert: Nemotecnico truncado en primer '_'"
    Test_Helper_ExpedienteConsultas_GenerarFilasLote_NemotecnicoConUnderscore_DevuelvePrimeraParte = BuildJsonOk("ok", logs)
End Function

' BR-23-RT-06: Basado incluye Ejercito en la fila
Public Function Test_Helper_ExpedienteConsultas_GenerarFilasBasado_DevuelveFilaConEjercito() As String
    Dim logs(0 To 2) As String
    Dim col As Scripting.Dictionary
    Dim errMsg As String
    Dim result As Variant
    Set col = New Scripting.Dictionary
    col.CompareMode = TextCompare
    col.Add "B1", BuildExpedienteCompleto("B1", "COD-B1", "Basado", "B_NAME", "Vigente", "01", "P-1", "01/01/2024", "31/12/2024", "EjercitoTierra")
    logs(0) = "Arrange: 1 basado con Ejercito='EjercitoTierra'"
    result = Helper_ExpedienteConsultas.GenerarFilasBasado(col, "Sí", errMsg)
    logs(1) = "Act: GenerarFilasBasado"
    If errMsg <> "" Then
        Test_Helper_ExpedienteConsultas_GenerarFilasBasado_DevuelveFilaConEjercito = BuildJsonFail(errMsg, logs)
        Exit Function
    End If
    ' Fila esperada: "B1;COD-B1;EjercitoTierra;01;Vigente"
    If InStr(result(0), "EjercitoTierra") = 0 Then
        Test_Helper_ExpedienteConsultas_GenerarFilasBasado_DevuelveFilaConEjercito = BuildJsonFail("expected 'EjercitoTierra' in row, got: " & result(0), logs)
        Exit Function
    End If
    logs(2) = "Assert: row contains Ejercito='EjercitoTierra'"
    Test_Helper_ExpedienteConsultas_GenerarFilasBasado_DevuelveFilaConEjercito = BuildJsonOk("ok", logs)
End Function

' BR-23-RT-07: collections vacías → todos los counts en 0
Public Function Test_Helper_ExpedienteConsultas_ContarExpedientes_ColVacias_DevuelveCero() As String
    Dim logs(0 To 2) As String
    Dim colAM As Scripting.Dictionary
    Dim colLotes As Scripting.Dictionary
    Dim colBas As Scripting.Dictionary
    Dim nAM As Long, nL As Long, nB As Long, nT As Long
    Dim errMsg As String
    Set colAM = New Scripting.Dictionary
    Set colLotes = New Scripting.Dictionary
    Set colBas = New Scripting.Dictionary
    logs(0) = "Arrange: 3 collections vacías"
    Helper_ExpedienteConsultas.ContarExpedientes colAM, colLotes, colBas, nAM, nL, nB, nT, errMsg
    logs(1) = "Act: ContarExpedientes executed"
    If errMsg <> "" Then
        Test_Helper_ExpedienteConsultas_ContarExpedientes_ColVacias_DevuelveCero = BuildJsonFail(errMsg, logs)
        Exit Function
    End If
    If nAM <> 0 Or nL <> 0 Or nB <> 0 Or nT <> 0 Then
        Test_Helper_ExpedienteConsultas_ContarExpedientes_ColVacias_DevuelveCero = BuildJsonFail("expected 0/0/0/0, got " & nAM & "/" & nL & "/" & nB & "/" & nT, logs)
        Exit Function
    End If
    logs(2) = "Assert: all counts 0"
    Test_Helper_ExpedienteConsultas_ContarExpedientes_ColVacias_DevuelveCero = BuildJsonOk("ok", logs)
End Function

' BR-23-RT-08: counts correctos: 2+3+1=6
Public Function Test_Helper_ExpedienteConsultas_ContarExpedientes_ConColecciones_DevuelveSuma() As String
    Dim logs(0 To 2) As String
    Dim colAM As New Scripting.Dictionary
    Dim colL As New Scripting.Dictionary
    Dim colB As New Scripting.Dictionary
    Dim nAM As Long, nL As Long, nB As Long, nT As Long
    Dim errMsg As String
    Dim i As Long
    For i = 1 To 2: colAM.Add "AM" & i, BuildExpedienteCompleto("AM" & i, "C" & i, "AM", "N" & i, "V", "0" & i, "", "", "", ""): Next
    For i = 1 To 3: colL.Add "L" & i, BuildExpedienteCompleto("L" & i, "C" & i, "Lote", "N" & i, "V", "0" & i, "P", "", "", ""): Next
    colB.Add "B1", BuildExpedienteCompleto("B1", "CB1", "Basado", "NB1", "V", "01", "P", "", "", "Tierra")
    logs(0) = "Arrange: AM=2, Lotes=3, Basados=1"
    Helper_ExpedienteConsultas.ContarExpedientes colAM, colL, colB, nAM, nL, nB, nT, errMsg
    logs(1) = "Act: ContarExpedientes executed"
    If errMsg <> "" Then
        Test_Helper_ExpedienteConsultas_ContarExpedientes_ConColecciones_DevuelveSuma = BuildJsonFail(errMsg, logs)
        Exit Function
    End If
    If nAM <> 2 Or nL <> 3 Or nB <> 1 Or nT <> 6 Then
        Test_Helper_ExpedienteConsultas_ContarExpedientes_ConColecciones_DevuelveSuma = BuildJsonFail("expected 2/3/1/6, got " & nAM & "/" & nL & "/" & nB & "/" & nT, logs)
        Exit Function
    End If
    logs(2) = "Assert: counts 2/3/1, total 6"
    Test_Helper_ExpedienteConsultas_ContarExpedientes_ConColecciones_DevuelveSuma = BuildJsonOk("ok", logs)
End Function
