Attribute VB_Name = "Test_ControlCambios"
Option Compare Database
Option Explicit

Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

Public Function Test_ControlCambios_AlcanceResumen3_MasDeTresEdiciones() As String
    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: maxEd=10, minDisponible=1, alcance=Resumen3"
    logs(1) = "2. Act: calcular edición mínima incluida"
    logs(2) = "3. Assert: debe devolver N-2"

    Dim minEd As Long
    minEd = ControlCambios_CalcularEdicionMinimaParaAlcance(10, 1, EnumControlCambiosAlcanceResumen3)

    logs(3) = "Resultado minEd=" & CStr(minEd)
    If minEd <> 8 Then
        Test_ControlCambios_AlcanceResumen3_MasDeTresEdiciones = BuildFail("Se esperaba minEd=8 y llegó " & CStr(minEd), logs)
        Exit Function
    End If

    Test_ControlCambios_AlcanceResumen3_MasDeTresEdiciones = BuildOk("min_ed_resumen3_ok", logs)
End Function

Public Function Test_ControlCambios_AlcanceResumen3_MenosDeTresEdiciones() As String
    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: maxEd=2, minDisponible=1, alcance=Resumen3"
    logs(1) = "2. Act: calcular edición mínima incluida"
    logs(2) = "3. Assert: no debe ir por debajo de mín disponible"

    Dim minEd As Long
    minEd = ControlCambios_CalcularEdicionMinimaParaAlcance(2, 1, EnumControlCambiosAlcanceResumen3)

    logs(3) = "Resultado minEd=" & CStr(minEd)
    If minEd <> 1 Then
        Test_ControlCambios_AlcanceResumen3_MenosDeTresEdiciones = BuildFail("Se esperaba minEd=1 y llegó " & CStr(minEd), logs)
        Exit Function
    End If

    Test_ControlCambios_AlcanceResumen3_MenosDeTresEdiciones = BuildOk("min_ed_resumen3_short_history_ok", logs)
End Function

Public Function Test_ControlCambios_AlcanceCompleto_NoRecortaHistorial() As String
    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: maxEd=10, minDisponible=1, alcance=Completo"
    logs(1) = "2. Act: calcular edición mínima incluida"
    logs(2) = "3. Assert: debe mantener minDisponible"

    Dim minEd As Long
    minEd = ControlCambios_CalcularEdicionMinimaParaAlcance(10, 1, EnumControlCambiosAlcanceCompleto)

    logs(3) = "Resultado minEd=" & CStr(minEd)
    If minEd <> 1 Then
        Test_ControlCambios_AlcanceCompleto_NoRecortaHistorial = BuildFail("Se esperaba minEd=1 y llegó " & CStr(minEd), logs)
        Exit Function
    End If

    Test_ControlCambios_AlcanceCompleto_NoRecortaHistorial = BuildOk("min_ed_completo_full_history_ok", logs)
End Function

Public Function Test_ControlCambios_ResolverAlcance_ResumidoPorDefecto() As String
    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: selector vacío con default=Resumen3"
    logs(1) = "2. Act: resolver alcance"
    logs(2) = "3. Assert: debe devolver Resumen3"

    Dim alcance As EnumControlCambiosAlcance
    alcance = ControlCambios_ResolverAlcanceDesdeSelector("", EnumControlCambiosAlcanceResumen3)

    logs(3) = "Resultado alcance=" & CStr(alcance)
    If alcance <> EnumControlCambiosAlcanceResumen3 Then
        Test_ControlCambios_ResolverAlcance_ResumidoPorDefecto = BuildFail("Se esperaba alcance Resumen3 y llegó " & CStr(alcance), logs)
        Exit Function
    End If

    Test_ControlCambios_ResolverAlcance_ResumidoPorDefecto = BuildOk("resolver_alcance_resumido_ok", logs)
End Function

Public Function Test_ControlCambios_ResolverAlcance_HistoricoCompleto() As String
    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: selector='HISTORICO'"
    logs(1) = "2. Act: resolver alcance"
    logs(2) = "3. Assert: debe devolver Completo"

    Dim alcance As EnumControlCambiosAlcance
    alcance = ControlCambios_ResolverAlcanceDesdeSelector("HISTORICO", EnumControlCambiosAlcanceResumen3)

    logs(3) = "Resultado alcance=" & CStr(alcance)
    If alcance <> EnumControlCambiosAlcanceCompleto Then
        Test_ControlCambios_ResolverAlcance_HistoricoCompleto = BuildFail("Se esperaba alcance Completo y llegó " & CStr(alcance), logs)
        Exit Function
    End If

    Test_ControlCambios_ResolverAlcance_HistoricoCompleto = BuildOk("resolver_alcance_historico_ok", logs)
End Function

Public Function Test_ControlCambios_DebeUsarHistorico_CompletoTrue_ResumenFalse() As String
    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: alcance Completo y Resumen3"
    logs(1) = "2. Act: evaluar contrato de publicación"
    logs(2) = "3. Assert: Completo=True y Resumen3=False"

    Dim usaCompleto As Boolean
    Dim usaResumen As Boolean
    usaCompleto = ControlCambios_DebeUsarHistoricoCompleto(EnumControlCambiosAlcanceCompleto)
    usaResumen = ControlCambios_DebeUsarHistoricoCompleto(EnumControlCambiosAlcanceResumen3)

    logs(3) = "Resultado completo=" & CStr(usaCompleto)
    logs(4) = "Resultado resumen3=" & CStr(usaResumen)

    If Not usaCompleto Then
        Test_ControlCambios_DebeUsarHistorico_CompletoTrue_ResumenFalse = BuildFail("Se esperaba True para alcance completo", logs)
        Exit Function
    End If

    If usaResumen Then
        Test_ControlCambios_DebeUsarHistorico_CompletoTrue_ResumenFalse = BuildFail("Se esperaba False para alcance resumen3", logs)
        Exit Function
    End If

    Test_ControlCambios_DebeUsarHistorico_CompletoTrue_ResumenFalse = BuildOk("publicacion_alcance_contract_ok", logs)
End Function

Public Function Test_ControlCambios_HTMLScope_Resumen3VsHistoricoContract() As String
    Dim logs(0 To 6) As String
    Dim minResumen As Long
    Dim minHistorico As Long
    Dim maxEdicion As Long
    Dim minDisponible As Long

    logs(0) = "1. Arrange: gestión con cinco ediciones de control de cambios"
    logs(1) = "2. Arrange: no histórico debe usar Resumen3"
    logs(2) = "3. Arrange: histórico debe usar Completo"
    logs(3) = "4. Act: calcular edición mínima incluida por alcance"
    logs(4) = "5. Assert: Resumen3 conserva solo las últimas tres"
    logs(5) = "6. Assert: histórico conserva todo el historial"

    maxEdicion = 5
    minDisponible = 1
    minResumen = ControlCambios_CalcularEdicionMinimaParaAlcance( _
        maxEdicion, _
        minDisponible, _
        EnumControlCambiosAlcanceResumen3)
    minHistorico = ControlCambios_CalcularEdicionMinimaParaAlcance( _
        maxEdicion, _
        minDisponible, _
        EnumControlCambiosAlcanceCompleto)

    logs(6) = "Resultado Resumen3=" & CStr(minResumen) & _
        "; Histórico=" & CStr(minHistorico)

    If minResumen <> 3 Then
        Test_ControlCambios_HTMLScope_Resumen3VsHistoricoContract = _
            BuildFail("Sin histórico debe empezar en edición 3 para un total de 5", logs)
        Exit Function
    End If

    If minHistorico <> 1 Then
        Test_ControlCambios_HTMLScope_Resumen3VsHistoricoContract = _
            BuildFail("Histórico debe empezar en edición 1 para incluir todo", logs)
        Exit Function
    End If

    If minResumen = minHistorico Then
        Test_ControlCambios_HTMLScope_Resumen3VsHistoricoContract = _
            BuildFail("Histórico y no histórico no pueden resolver el mismo corte", logs)
        Exit Function
    End If

    Test_ControlCambios_HTMLScope_Resumen3VsHistoricoContract = _
        BuildOk("html_scope_resumen3_vs_historico_contract_ok", logs)
End Function

Public Function Test_ControlCambios_Diff_IgnoraFormato_NoCambioEfectivo() As String
    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: comparar priorización con espacios/mayúsculas"
    logs(1) = "2. Act: evaluar diferencia efectiva"
    logs(2) = "3. Assert: no debe detectar cambio"

    Dim diff As Boolean
    diff = ControlCambios_DiffValoresDifieren("  ALTA  ", "alta")

    logs(3) = "Resultado diff=" & CStr(diff)
    If diff Then
        Test_ControlCambios_Diff_IgnoraFormato_NoCambioEfectivo = BuildFail("Se detectó cambio falso positivo por formato", logs)
        Exit Function
    End If

    Test_ControlCambios_Diff_IgnoraFormato_NoCambioEfectivo = BuildOk("diff_ignora_formato_ok", logs)
End Function

Public Function Test_ControlCambios_Diff_Priorizacion_CambioReal() As String
    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: comparar priorización ALTA vs MEDIA"
    logs(1) = "2. Act: evaluar diferencia efectiva"
    logs(2) = "3. Assert: debe detectar cambio real"

    Dim diff As Boolean
    diff = ControlCambios_DiffValoresDifieren("Alta", "Media")

    logs(3) = "Resultado diff=" & CStr(diff)
    If Not diff Then
        Test_ControlCambios_Diff_Priorizacion_CambioReal = BuildFail("No detectó cambio real de priorización", logs)
        Exit Function
    End If

    Test_ControlCambios_Diff_Priorizacion_CambioReal = BuildOk("diff_priorizacion_real_ok", logs)
End Function
