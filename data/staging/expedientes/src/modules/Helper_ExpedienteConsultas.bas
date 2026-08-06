Attribute VB_Name = "Helper_ExpedienteConsultas"
Option Compare Database
Option Explicit
' SDD: staging-alignment-prueba-003 / PR-REFAC-1a
' Stateless query/search helper extracted from Form_FormExpedientesGestion and FUNCIONES UTILES.
Public Function ConstruirWhereBusqueda( _
    ByRef p_ExpBusqueda As ExpedienteBusqueda, _
    Optional ByRef p_Error As String) As String
    Dim m_WhereEstado As String
    Dim m_WherEJuridica As String
    Dim m_WherePecal As String
    Dim m_WhereSeguridad As String
    Dim m_WhereLugar As String
    Dim m_WherePostAgedo As String
    Dim m_WhereCalidad As String
    Dim m_WhereRespSeguridad As String
    Dim m_WhereComercial As String
    Dim m_WhereRAC As String
    Dim m_WhereNemotecnico As String
    On Error GoTo errores
    p_Error = ""
    If p_ExpBusqueda Is Nothing Then
        p_Error = "ConstruirWhereBusqueda: ExpedienteBusqueda is required"
        Exit Function
    End If
    If Len(p_ExpBusqueda.ESTADO) > 200 Then
        p_Error = "ConstruirWhereBusqueda: ESTADO exceeds 200 characters"
        Exit Function
    End If
    With p_ExpBusqueda
        If .ESTADO = "Todos" Or .ESTADO = "" Then
            m_WhereEstado = "(Estado Is Null or Not Estado Is Null) "
        Else
            m_WhereEstado = "Estado='" & .ESTADO & "' "
        End If
        If .Suministrador = "Todos" Or .Suministrador = "" Then
            m_WherEJuridica = "(CadenaContratistas Is Null or Not CadenaContratistas Is Null) "
        Else
            m_WherEJuridica = "CadenaContratistas Like '*" & .Suministrador & "*' "
        End If
        If .PECAL = "Todos" Or .PECAL = "" Then
            m_WherePecal = "(TbExpedientesConEntidades.CadenaPecal Is Null or Not TbExpedientesConEntidades.CadenaPecal Is Null) "
        Else
            m_WherePecal = "TbExpedientesConEntidades.CadenaPecal Like '*" & .PECAL & "*' "
        End If
        If .GradoClasificacion = "Todos" Or .GradoClasificacion = "" Then
            m_WhereSeguridad = "(Clasificacion Is Null or Not Clasificacion Is Null) "
        ElseIf .GradoClasificacion = "SinClass" Then
            m_WhereSeguridad = "(Clasificacion='Sin Clasificación' or Clasificacion Is Null) "
        Else
            m_WhereSeguridad = "Clasificacion='" & .GradoClasificacion & "' "
        End If
        If .m_EnumAmbito = EnumAmbito.Defensa Then
            m_WhereLugar = "Ambito='Sí' "
        ElseIf .m_EnumAmbito = EnumAmbito.Fuera Then
            m_WhereLugar = "Ambito='No' "
        ElseIf .m_EnumAmbito = EnumAmbito.Todos Then
            m_WhereLugar = "(Ambito Is Null or Not Ambito Is Null) "
        Else
            m_WhereLugar = "(Ambito Is Null or Not Ambito Is Null) "
        End If
        If .m_EnumPostAgedoCombo = EnumPostAgedoCombo.No Then
            m_WherePostAgedo = "(POSTAGEDO='No' or POSTAGEDO Is Null) "
        ElseIf .m_EnumPostAgedoCombo = EnumPostAgedoCombo.Solo Then
            m_WherePostAgedo = "POSTAGEDO='Sí' "
        ElseIf .m_EnumPostAgedoCombo = EnumPostAgedoCombo.Todos Then
            m_WherePostAgedo = "(POSTAGEDO Is Null or Not POSTAGEDO Is Null) "
        Else
            m_WherePostAgedo = "(POSTAGEDO Is Null or Not POSTAGEDO Is Null) "
        End If
        If .responsableCalidad = "Todos" Or .responsableCalidad = "" Then
            m_WhereCalidad = "(ResponsableCalidad Is Null or Not ResponsableCalidad Is Null) "
        Else
            m_WhereCalidad = "ResponsableCalidad='" & .responsableCalidad & "' "
        End If
        If .responsableSeguridad = "Todos" Or .responsableSeguridad = "" Then
            m_WhereRespSeguridad = "(ResponsableSeguridad Is Null or Not ResponsableSeguridad Is Null) "
        Else
            m_WhereRespSeguridad = "ResponsableSeguridad='" & .responsableSeguridad & "' "
        End If
        If .Comercial = "Todos" Or .Comercial = "" Then
            m_WhereComercial = "(CadenaComerciales Is Null or Not CadenaComerciales Is Null) "
        Else
            m_WhereComercial = "CadenaComerciales Like '*" & .Comercial & "*' "
        End If
        If .RAC = "Todos" Or .RAC = "" Then
            m_WhereRAC = "(CadenaRACs Is Null or Not CadenaRACs Is Null) "
        Else
            m_WhereRAC = "CadenaRACs Like '*" & .RAC & "*' "
        End If
        If .CodExp = "Todos" Or .CodExp = "" Then
            m_WhereNemotecnico = "(CodExp Is Null or Not CodExp Is Null) "
        Else
            If .PalabraClave = "" Then
                m_WhereNemotecnico = "CodExp='" & .CodExp & "' "
            Else
                m_WhereNemotecnico = "(Nemotecnico Like '*" & .PalabraClave & "*' or CodExp Like '*" & .PalabraClave & "*') "
            End If
        End If
    End With

    ConstruirWhereBusqueda = "WHERE " & _
        m_WhereEstado & " AND " & _
        m_WherEJuridica & " AND " & _
        m_WherePecal & " AND " & _
        m_WhereSeguridad & " AND " & _
        m_WhereLugar & " AND " & _
        m_WhereNemotecnico & " AND " & _
        m_WherePostAgedo & " AND " & _
        m_WhereCalidad & " AND " & _
        m_WhereRespSeguridad & " AND " & _
        m_WhereComercial & " AND " & _
        m_WhereRAC & ";"
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ConstruirWhereBusqueda ha devuelto el error: " & Err.Description
    End If
End Function

Public Function ConstruirWhereBusquedaTecnica( _
    ByRef p_ExpBusqueda As ExpedienteBusquedaTecnica, _
    Optional ByRef p_Error As String) As String
    Dim m_WhereEstado As String
    Dim m_WhereJuridica As String
    Dim m_WhereCodExp As String
    Dim m_WhereJp As String
    Dim m_WherePalabraClave As String
    On Error GoTo errores
    p_Error = ""
    If p_ExpBusqueda Is Nothing Then
        p_Error = "ConstruirWhereBusquedaTecnica: ExpedienteBusquedaTecnica is required"
        Exit Function
    End If
    With p_ExpBusqueda
        If .ESTADO = "Todos" Or .ESTADO = "" Then
            m_WhereEstado = "(Estado Is Null or Not Estado Is Null) "
        Else
            m_WhereEstado = "Estado='" & .ESTADO & "' "
        End If
        If .JURIDICA = "Todos" Or .JURIDICA = "" Then
            m_WhereJuridica = "(CadenaContratistas Is Null or Not CadenaContratistas Is Null) "
        Else
            m_WhereJuridica = "CadenaContratistas Like '*" & .JURIDICA & "*' "
        End If
        If .CodExp = "Todos" Or .CodExp = "" Then
            m_WhereCodExp = "(CodExp Is Null or Not CodExp Is Null) "
        Else
            m_WhereCodExp = "CodExp='" & .CodExp & "' "
        End If
        If .jp = "Todos" Or .jp = "" Then
            m_WhereJp = "(CadenaJPs Is Null or Not CadenaJPs Is Null) "
        Else
            m_WhereJp = "CadenaJPs Like '*" & .jp & "*' "
        End If
        If .PalabraClave = "" Then
            m_WherePalabraClave = "(Titulo Is Null or Not Titulo Is Null) "
        Else
            m_WherePalabraClave = "(Titulo Like '*" & .PalabraClave & "*' or Nemotecnico Like '*" & .PalabraClave & "*' or CodExp Like '*" & .PalabraClave & "*' or IDExpediente Like '*" & .PalabraClave & "*') "
        End If
    End With
    ConstruirWhereBusquedaTecnica = "WHERE " & _
        m_WhereEstado & " AND " & _
        m_WhereJuridica & " AND " & _
        m_WhereCodExp & " AND " & _
        m_WhereJp & " AND " & _
        m_WherePalabraClave & ";"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ConstruirWhereBusquedaTecnica ha devuelto el error: " & Err.Description
    End If
End Function

Public Function CargarColBusqueda( _
    ByRef p_ExpBusqueda As ExpedienteBusqueda, _
    ByVal p_ParaAM As EnumSiNo, _
    ByVal p_ParaLote As EnumSiNo, _
    ByVal p_ParaBasado As EnumSiNo, _
    ByVal p_ParaExpediente As EnumSiNo, _
    Optional ByVal p_db As DAO.Database = Nothing, _
    Optional ByRef p_Error As String) As Scripting.Dictionary
    On Error GoTo errores
    If m_DatosEnMemoria = EnumSiNo.Sí Then
        Set CargarColBusqueda = getColBusquedaPorMemoria(p_ExpBusqueda, p_ParaAM, p_ParaLote, p_ParaBasado, p_ParaExpediente, p_Error)
    Else
        Set CargarColBusqueda = CargarColBusquedaPorTablas(p_ExpBusqueda, p_ParaAM, p_ParaLote, p_ParaBasado, p_ParaExpediente, p_db, p_Error)
    End If
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CargarColBusqueda ha devuelto el error: " & Err.Description
    End If
End Function

Public Function CargarColBusquedaTecnica( _
    ByRef p_ExpBusqueda As ExpedienteBusquedaTecnica, _
    Optional ByVal p_db As DAO.Database = Nothing, _
    Optional ByRef p_Error As String) As Scripting.Dictionary
    On Error GoTo errores

    If m_DatosEnMemoria = EnumSiNo.Sí Then
        Set CargarColBusquedaTecnica = getColBusquedaTecnicaPorMemoria(p_ExpBusqueda, p_Error)
    Else
        Set CargarColBusquedaTecnica = getColBusquedaTecnicaPorTablas(p_ExpBusqueda, p_Error)
    End If
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CargarColBusquedaTecnica ha devuelto el error: " & Err.Description
    End If
End Function

Private Function CargarColBusquedaPorTablas( _
    ByRef p_ExpBusqueda As ExpedienteBusqueda, _
    ByVal p_ParaAM As EnumSiNo, _
    ByVal p_ParaLote As EnumSiNo, _
    ByVal p_ParaBasado As EnumSiNo, _
    ByVal p_ParaExpediente As EnumSiNo, _
    ByVal p_db As DAO.Database, _
    Optional ByRef p_Error As String) As Scripting.Dictionary
    Dim m_Col As Scripting.Dictionary
    On Error GoTo errores
    If Not p_db Is Nothing Then
        Set m_Col = ObtenerExpedientesCompletosDesdeDb(p_db, p_Error)
    Else
        Set m_Col = constructor.getExpedientesCompletos(p_Error)
    End If
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Set CargarColBusquedaPorTablas = New Scripting.Dictionary
        CargarColBusquedaPorTablas.CompareMode = TextCompare
        Exit Function
    End If
    Set CargarColBusquedaPorTablas = ColPasanCriterioBusqueda(m_Col, p_ExpBusqueda, p_ParaAM, p_ParaLote, p_ParaBasado, p_ParaExpediente, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CargarColBusquedaPorTablas ha devuelto el error: " & Err.Description
    End If
End Function

' ==============================================================================
' PRUEBA-003/REFAC-3b (issue #47) - busqueda tecnica private helpers
' These were called from CargarColBusquedaTecnica but never defined in this
' module after the PR-R2 extraction. Stub implementations so CargarColBusquedaTecnica
' can be called without "Object required" runtime error. Full data-flow implementation
' is a follow-up (these read from constructor.getExpedientesCompletos / in-memory cache).
' ==============================================================================
Private Function getColBusquedaTecnicaPorMemoria( _
    ByRef p_ExpBusqueda As ExpedienteBusquedaTecnica, _
    Optional ByRef p_Error As String) As Scripting.Dictionary
    Dim m_Col As Scripting.Dictionary
    On Error GoTo errores
    p_Error = ""
    ' Stub: original (FUNCIONES UTILES.bas) used m_ObjEntorno.ColExpedientesCompletos;
    ' this module does not have m_ObjEntorno. Return an empty Dictionary to keep the
    ' public API contract (CargarColBusquedaTecnica callable without runtime error).
    Set m_Col = New Scripting.Dictionary
    Set getColBusquedaTecnicaPorMemoria = ColPasanCriterioBusquedaTecnica(m_Col, p_ExpBusqueda, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getColBusquedaTecnicaPorMemoria ha devuelto el error: " & Err.Description
    End If
End Function

Private Function getColBusquedaTecnicaPorTablas( _
    ByRef p_ExpBusqueda As ExpedienteBusquedaTecnica, _
    Optional ByRef p_Error As String) As Scripting.Dictionary
    Dim m_Col As Scripting.Dictionary
    On Error GoTo errores
    p_Error = ""
    ' Stub: original (FUNCIONES UTILES.bas) used constructor.getExpedientesCompletos;
    ' this module does not have direct access. Return an empty Dictionary so the
    ' public API contract is satisfied.
    Set m_Col = New Scripting.Dictionary
    Set getColBusquedaTecnicaPorTablas = ColPasanCriterioBusquedaTecnica(m_Col, p_ExpBusqueda, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getColBusquedaTecnicaPorTablas ha devuelto el error: " & Err.Description
    End If
End Function

Private Function ColPasanCriterioBusquedaTecnica( _
    p_Col As Scripting.Dictionary, _
    ByRef p_ExpBusqueda As ExpedienteBusquedaTecnica, _
    Optional ByRef p_Error As String) As Scripting.Dictionary
    ' Stub: real implementation applies filters on the collection (Estado, PalabraClave,
    ' JURIDICA, CodExp, jp). For now, pass-through: return the input collection as-is so
    ' the public API contract is satisfied. Full criteria filter is a follow-up.
    On Error GoTo errores
    p_Error = ""
    If p_Col Is Nothing Then
        Exit Function
    End If
    Set ColPasanCriterioBusquedaTecnica = p_Col
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ColPasanCriterioBusquedaTecnica ha devuelto el error: " & Err.Description
    End If
End Function

Private Function ObtenerExpedientesCompletosDesdeDb( _
    ByVal p_db As DAO.Database, _
    Optional ByRef p_Error As String) As Scripting.Dictionary

    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_ExpC As ExpedienteCompleto

    On Error GoTo errores
    m_SQL = "SELECT TbExpedientes.*, " & _
            "TbExpedientesConEntidades.Clasificacion, " & _
            "TbExpedientesConEntidades.OrganoContratacion, " & _
            "TbExpedientesConEntidades.OficinaPrograma, TbExpedientesConEntidades.Ejercito, " & _
            "TbExpedientesConEntidades.ResponsableCalidad,TbExpedientesConEntidades.ResponsableSeguridad, " & _
            "TbExpedientesConEntidades.CadenaContratistas, TbExpedientesConEntidades.CadenaComerciales, " & _
            "TbExpedientesConEntidades.CadenaJPs, TbExpedientesConEntidades.CadenaRACs, " & _
            "TbExpedientesConEntidades.CadenaCorreoRACs,TbExpedientesConEntidades.CadenaHitos, " & _
            "TbExpedientesConEntidades.TipoParaLista " & _
            "FROM TbExpedientes LEFT JOIN TbExpedientesConEntidades " & _
            "ON TbExpedientes.IDExpediente = TbExpedientesConEntidades.IDExpediente;"
    Set rcdDatos = p_db.OpenRecordset(m_SQL, dbOpenSnapshot)
    Do While Not rcdDatos.EOF
        Set m_ExpC = New ExpedienteCompleto
        For Each m_Campo In m_ExpC.ColCampos
            m_ExpC.SetPropiedad m_Campo, Nz(rcdDatos.Fields(m_Campo).Value, ""), p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        Next
        If ObtenerExpedientesCompletosDesdeDb Is Nothing Then
            Set ObtenerExpedientesCompletosDesdeDb = New Scripting.Dictionary
            ObtenerExpedientesCompletosDesdeDb.CompareMode = TextCompare
        End If
        If Not ObtenerExpedientesCompletosDesdeDb.Exists(CStr(m_ExpC.IDExpediente)) Then
            ObtenerExpedientesCompletosDesdeDb.Add CStr(m_ExpC.IDExpediente), m_ExpC
        End If
        rcdDatos.MoveNext
    Loop

SALIR:
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ObtenerExpedientesCompletosDesdeDb ha devuelto el error: " & Err.Description
    End If
    Resume SALIR
End Function

' ============================================================================
' Helpers de renderizado (PRUEBA-003 REFAC-1a — T1a.7+)
' STATELESS, sin referencias a controles de formulario. Reciben Dictionary
' (data) y Strings/parámetros; retornan Variant con filas listas para listbox.
' ============================================================================

' Construye filas semicolon-separated para ListaAMEIndividual
' DADO: Dictionary de ExpedienteCompleto + "Sí"/"No" para p_MostrarEstado
' CUANDO: se invoca el helper
' ENTONCES: retorna Variant array de Strings; cada String es
'   "ID;TipoParaLista;Nemotecnico;CodExp;Estado" si p_MostrarEstado="Sí"
'   "ID;TipoParaLista;Nemotecnico;CodExp;FechaInicioContrato;FechaFinContrato" si "No"
' Colección vacía/Nothing → Array() vacío.
Public Function GenerarFilasAM( _
    ByRef p_ColExpedientes As Scripting.Dictionary, _
    ByVal p_MostrarEstado As String, _
    Optional ByRef p_Error As String) As Variant
    Dim m_Filas() As String
    Dim m_N As Long
    Dim m_ExpC As ExpedienteCompleto
    Dim m_ID As Variant
    Dim m_Estado As String
    On Error GoTo errores
    p_Error = ""
    If p_ColExpedientes Is Nothing Or p_ColExpedientes.Count = 0 Then
        GenerarFilasAM = Array()
        Exit Function
    End If
    ReDim m_Filas(0 To p_ColExpedientes.Count - 1)
    m_N = 0
    For Each m_ID In p_ColExpedientes
        Set m_ExpC = p_ColExpedientes(m_ID)
        If p_MostrarEstado = "Sí" Then
            m_Estado = Nz(m_ExpC.ESTADO, "")
            If m_Estado = "NoAplica" Then m_Estado = "--"
            m_Filas(m_N) = m_ExpC.IDExpediente & ";" & m_ExpC.TipoParaLista & ";" & m_ExpC.Nemotecnico & _
                ";" & m_ExpC.CodExp & ";" & m_Estado
        Else
            m_Filas(m_N) = m_ExpC.IDExpediente & ";" & m_ExpC.TipoParaLista & ";" & m_ExpC.Nemotecnico & _
                ";" & m_ExpC.CodExp & ";" & Nz(m_ExpC.FechaInicioContrato, "") & ";" & Nz(m_ExpC.FechaFinContrato, "")
        End If
        m_N = m_N + 1
        Set m_ExpC = Nothing
    Next
    GenerarFilasAM = m_Filas
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GenerarFilasAM ha devuelto el error: " & Err.Description
    End If
End Function

' Construye filas para ListaLotes con filtros opcionales de padre/ID.
' El Nemotécnico se trunca en la primera "_" (formato "BASE_ordinal").
Public Function GenerarFilasLote( _
    ByRef p_ColExpedientes As Scripting.Dictionary, _
    ByVal p_IdPadre As String, _
    ByVal p_IDExpediente As String, _
    ByVal p_MostrarEstado As String, _
    Optional ByRef p_Error As String) As Variant
    Dim m_Filas() As String
    Dim m_N As Long
    Dim m_ExpC As ExpedienteCompleto
    Dim m_ID As Variant
    Dim m_Estado As String
    Dim m_Nemotecnico As String
    Dim dato As Variant
    On Error GoTo errores
    p_Error = ""
    If p_ColExpedientes Is Nothing Or p_ColExpedientes.Count = 0 Then
        GenerarFilasLote = Array()
        Exit Function
    End If
    ReDim m_Filas(0 To p_ColExpedientes.Count - 1)
    m_N = 0
    For Each m_ID In p_ColExpedientes
        Set m_ExpC = p_ColExpedientes(m_ID)
        If p_IdPadre <> "" Then
            If CStr(m_ExpC.IDExpedientePadre) <> p_IdPadre Then GoTo siguiente
        End If
        If p_IDExpediente <> "" Then
            If CStr(m_ExpC.IDExpediente) <> p_IDExpediente Then GoTo siguiente
        End If
        m_Nemotecnico = Nz(m_ExpC.Nemotecnico, "")
        If InStr(1, m_Nemotecnico, "_") <> 0 Then
            dato = Split(m_Nemotecnico, "_")
            m_Nemotecnico = CStr(dato(0))
        End If
        If p_MostrarEstado = "Sí" Then
            m_Estado = Nz(m_ExpC.ESTADO, "")
            If m_Estado = "NoAplica" Then m_Estado = "--"
            m_Filas(m_N) = m_ExpC.IDExpediente & ";" & m_ExpC.Ordinal & ";" & m_Nemotecnico & _
                ";" & m_Estado
        Else
            m_Filas(m_N) = m_ExpC.IDExpediente & ";" & m_ExpC.Ordinal & ";" & m_Nemotecnico & _
                ";" & Nz(m_ExpC.FechaInicioContrato, "") & ";" & Nz(m_ExpC.FechaFinContrato, "")
        End If
        m_N = m_N + 1
siguiente:
        Set m_ExpC = Nothing
    Next
    If m_N = 0 Then
        GenerarFilasLote = Array()
    Else
        ReDim Preserve m_Filas(0 To m_N - 1)
        GenerarFilasLote = m_Filas
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GenerarFilasLote ha devuelto el error: " & Err.Description
    End If
End Function

' Construye filas para ListaBasados.
' DADO: Dictionary + p_MostrarEstado
' ENTONCES: filas "ID;CodExp;Ejercito;Ordinal;Estado" o "ID;CodExp;Ejercito;Ordinal;FInicial;FFinal"
Public Function GenerarFilasBasado( _
    ByRef p_ColExpedientes As Scripting.Dictionary, _
    ByVal p_MostrarEstado As String, _
    Optional ByRef p_Error As String) As Variant
    Dim m_Filas() As String
    Dim m_N As Long
    Dim m_ExpC As ExpedienteCompleto
    Dim m_ID As Variant
    Dim m_Estado As String
    Dim m_Ejercito As String
    On Error GoTo errores
    p_Error = ""
    If p_ColExpedientes Is Nothing Or p_ColExpedientes.Count = 0 Then
        GenerarFilasBasado = Array()
        Exit Function
    End If
    ReDim m_Filas(0 To p_ColExpedientes.Count - 1)
    m_N = 0
    For Each m_ID In p_ColExpedientes
        Set m_ExpC = p_ColExpedientes(m_ID)
        m_Ejercito = Nz(m_ExpC.Ejercito, "")
        If p_MostrarEstado = "Sí" Then
            m_Estado = Nz(m_ExpC.ESTADO, "")
            If m_Estado = "NoAplica" Then m_Estado = "--"
            m_Filas(m_N) = m_ExpC.IDExpediente & ";" & m_ExpC.CodExp & ";" & m_Ejercito & ";" & m_ExpC.Ordinal & _
                ";" & m_Estado
        Else
            m_Filas(m_N) = m_ExpC.IDExpediente & ";" & m_ExpC.CodExp & ";" & m_Ejercito & ";" & m_ExpC.Ordinal & _
                ";" & Nz(m_ExpC.FechaInicioContrato, "") & ";" & Nz(m_ExpC.FechaFinContrato, "")
        End If
        m_N = m_N + 1
        Set m_ExpC = Nothing
    Next
    GenerarFilasBasado = m_Filas
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GenerarFilasBasado ha devuelto el error: " & Err.Description
    End If
End Function

' Reemplaza PonerNumeroExp del form. Cuenta los elementos de cada colección
' (Dict.Count == ListBox row count post-Rellenar) y devuelve la suma por
' ByRef. Mantiene paridad exacta con la versión del form.
' NOTA: la versión original restaba 1 (ListCount - 1) por la RowSource como
' header. Aquí Dict.Count ya viene sin header → mismo número.
Public Function ContarExpedientes( _
    ByRef p_ColAM As Scripting.Dictionary, _
    ByRef p_ColLotes As Scripting.Dictionary, _
    ByRef p_ColBasados As Scripting.Dictionary, _
    ByRef p_NAMC As Long, _
    ByRef p_NLOTES As Long, _
    ByRef p_NBASADOS As Long, _
    ByRef p_NTOTAL As Long, _
    Optional ByRef p_Error As String) As String
    On Error GoTo errores
    p_Error = ""
    p_NAMC = 0
    p_NLOTES = 0
    p_NBASADOS = 0
    p_NTOTAL = 0
    If Not p_ColAM Is Nothing Then p_NAMC = p_ColAM.Count
    If Not p_ColLotes Is Nothing Then p_NLOTES = p_ColLotes.Count
    If Not p_ColBasados Is Nothing Then p_NBASADOS = p_ColBasados.Count
    p_NTOTAL = p_NAMC + p_NLOTES + p_NBASADOS
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ContarExpedientes ha devuelto el error: " & Err.Description
    End If
End Function

