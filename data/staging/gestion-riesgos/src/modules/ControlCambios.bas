Attribute VB_Name = "ControlCambios"
Option Compare Database
Option Explicit

Private Const C_CC_APP As String = "GESTION_RIESGOS"
Private Const C_CC_SECTION As String = "ControlCambios"
Private Const C_CC_KEY_USE_PREFIX As String = "UseCache_"
Private Const C_CC_ADD_OPEN As String = "{{CC+}}"
Private Const C_CC_ADD_CLOSE As String = "{{/CC+}}"
Private Const C_CC_DEL_OPEN As String = "{{CC-}}"
Private Const C_CC_DEL_CLOSE As String = "{{/CC-}}"
Private Const C_CACHE_LOGIC_VERSION As String = "v2"

Public Function ControlCambios_ConstruirSeccionControlCambiosHTML( _
                                                                    ByVal p_EdicionActual As Edicion, _
                                                                    ByVal p_Proyecto As Proyecto, _
                                                                    Optional ByRef p_Error As String _
                                                                    ) As String
    ControlCambios_ConstruirSeccionControlCambiosHTML = ControlCambios_ConstruirSeccionControlCambiosHTMLConAlcance( _
                                                    p_EdicionActual, _
                                                    p_Proyecto, _
                                                    EnumControlCambiosAlcanceCompleto, _
                                                    p_Error)
End Function

Public Function ControlCambios_ConstruirSeccionControlCambiosHTMLConAlcance( _
                                                                            ByVal p_EdicionActual As Edicion, _
                                                                            ByVal p_Proyecto As Proyecto, _
                                                                            Optional ByVal p_Alcance As EnumControlCambiosAlcance = EnumControlCambiosAlcanceCompleto, _
                                                                            Optional ByRef p_Error As String _
                                                                            ) As String
    Dim s As String
    Dim filas As Collection
    Dim fila As Object
    Dim codigo As String
    Dim ed As String
    Dim estadoHtml As String
    Dim pmHtml As String
    Dim pcHtml As String
    
    On Error GoTo errores
    p_Error = ""
    
    Set filas = ControlCambios_GetFilasHastaEdicionConAlcance( _
                p_EdicionActual, _
                p_Proyecto, _
                ControlCambios_UsarCacheProyecto(ControlCambios_ToLongSafe(p_Proyecto.IDProyecto, 0), p_Error), _
                p_Alcance, _
                p_Error)
    If p_Error <> "" Then Err.Raise 1000
    
    s = "<section class='report-section print-page seccion-vertical'>" & vbCrLf
    s = s & "  <h2>Control de cambios</h2>" & vbCrLf
    s = s & "  <div class='card'>" & vbCrLf
    s = s & "    <table class='report-table report-table-small cc-table'>" & vbCrLf
    s = s & "      <colgroup>" & _
            "<col class='cc-col-codigo'>" & _
            "<col class='cc-col-edicion'>" & _
            "<col class='cc-col-estado'>" & _
            "<col class='cc-col-pm'>" & _
            "<col class='cc-col-pc'>" & _
            "</colgroup>" & vbCrLf
    s = s & "      <thead><tr><th>Código</th><th>Ed.</th><th>Estado</th><th>Planes de mitigación</th><th>Planes de contingencia</th></tr></thead>" & vbCrLf
    s = s & "      <tbody>" & vbCrLf
    
    Dim hayFilas As Boolean
    
    hayFilas = False
    If Not filas Is Nothing Then
        If filas.Count > 0 Then hayFilas = True
    End If
    
    If Not hayFilas Then
        s = s & "        <tr><td colspan='5'>Sin cambios.</td></tr>" & vbCrLf
    Else
        For Each fila In filas
            codigo = Nz(fila("CodigoRiesgo"), "")
            Avance "Control Cambios ..." & codigo
            ed = Nz(fila("Edicion"), "")
            estadoHtml = Nz(fila("EstadoHtml"), "")
            pmHtml = Nz(fila("MitigacionHtml"), "")
            pcHtml = Nz(fila("ContingenciaHtml"), "")
            s = s & "        <tr>" & _
                    "<td>" & ControlCambios_HTMLSafe(codigo) & "</td>" & _
                    "<td>" & ControlCambios_HTMLSafe(ed) & "</td>" & _
                    "<td>" & ControlCambios_NzHtml(estadoHtml) & "</td>" & _
                    "<td>" & ControlCambios_NzHtml(pmHtml) & "</td>" & _
                    "<td>" & ControlCambios_NzHtml(pcHtml) & "</td>" & _
                    "</tr>" & vbCrLf
        Next
    End If
    
    s = s & "      </tbody>" & vbCrLf
    s = s & "      <tfoot>" & vbCrLf
    s = s & "        <tr><td colspan='5' style='height: 18mm; border: none !important; background: transparent !important;'></td></tr>" & vbCrLf
    s = s & "      </tfoot>" & vbCrLf
    s = s & "    </table>" & vbCrLf
    s = s & "  </div>" & vbCrLf
    s = s & "</section>" & vbCrLf
    
    ControlCambios_ConstruirSeccionControlCambiosHTMLConAlcance = s
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en ControlCambios_ConstruirSeccionControlCambiosHTMLConAlcance: " & Err.Description
    End If
End Function
 
Public Function ControlCambios_UsarCache(Optional ByRef p_Error As String) As EnumSiNo
    On Error GoTo errores
    ControlCambios_UsarCache = ControlCambios_UsarCacheProyecto(0, p_Error)
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en ControlCambios_UsarCache: " & Err.Description
    End If
End Function

Public Function ControlCambios_UsarCacheProyecto(ByVal p_IDProyecto As Long, Optional ByRef p_Error As String) As EnumSiNo
    Dim v As String
    
    On Error GoTo errores
    p_Error = ""
    
    v = GetSetting(C_CC_APP, C_CC_SECTION, C_CC_KEY_USE_PREFIX & CStr(p_IDProyecto), "Si")
    If LCase$(Trim$(v)) = "no" Then
        ControlCambios_UsarCacheProyecto = EnumSiNo.No
    Else
        ControlCambios_UsarCacheProyecto = 1
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en ControlCambios_UsarCacheProyecto: " & Err.Description
    End If
End Function

Public Sub ControlCambios_SetUsarCacheProyecto(ByVal p_IDProyecto As Long, ByVal p_UsarCache As EnumSiNo, Optional ByRef p_Error As String)
    On Error GoTo errores
    p_Error = ""
    
    If p_UsarCache = EnumSiNo.No Then
        SaveSetting C_CC_APP, C_CC_SECTION, C_CC_KEY_USE_PREFIX & CStr(p_IDProyecto), "No"
    Else
        SaveSetting C_CC_APP, C_CC_SECTION, C_CC_KEY_USE_PREFIX & CStr(p_IDProyecto), "Si"
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en ControlCambios_SetUsarCacheProyecto: " & Err.Description
    End If
End Sub

Public Sub ControlCambios_LimpiarCache(ByVal p_IDProyecto As Long, Optional ByVal p_Edicion As Long = 0, Optional ByRef p_Error As String)
    Dim db As DAO.Database
    Dim sqlRows As String
    Dim sqlMeta As String
    
    On Error GoTo errores
    p_Error = ""
    
    If p_IDProyecto <= 0 Then Exit Sub
    If Not ControlCambios_CacheSchemaDisponible() Then Exit Sub
    
    Set db = getdb(p_Error)
    If p_Error <> "" Then Err.Raise 1000
    
    If p_Edicion > 0 Then
        sqlRows = "DELETE FROM TbCacheControlCambiosRow WHERE IDProyecto=" & p_IDProyecto & " AND Edicion=" & p_Edicion & ";"
        sqlMeta = "DELETE FROM TbCacheControlCambiosMeta WHERE IDProyecto=" & p_IDProyecto & " AND Edicion=" & p_Edicion & ";"
    Else
        sqlRows = "DELETE FROM TbCacheControlCambiosRow WHERE IDProyecto=" & p_IDProyecto & ";"
        sqlMeta = "DELETE FROM TbCacheControlCambiosMeta WHERE IDProyecto=" & p_IDProyecto & ";"
    End If
    
    db.Execute sqlRows, dbFailOnError
    db.Execute sqlMeta, dbFailOnError
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en ControlCambios_LimpiarCache: " & Err.Description
    End If
End Sub

Public Function ControlCambios_GetCacheDDL() As String
    Dim s As String
    
    s = ""
    
    s = s & "CREATE TABLE TbCacheControlCambiosMeta (" & vbCrLf
    s = s & "  IDProyecto LONG NOT NULL," & vbCrLf
    s = s & "  Edicion LONG NOT NULL," & vbCrLf
    s = s & "  ActiveBuildId LONG NOT NULL," & vbCrLf
    s = s & "  UpdatedAt DATETIME," & vbCrLf
    s = s & "  CacheVersion TEXT(20)," & vbCrLf
    s = s & "  CONSTRAINT PK_TbCacheControlCambiosMeta PRIMARY KEY (IDProyecto, Edicion)" & vbCrLf
    s = s & ");" & vbCrLf & vbCrLf
    
    s = s & "CREATE TABLE TbCacheControlCambiosRow (" & vbCrLf
    s = s & "  IDProyecto LONG NOT NULL," & vbCrLf
    s = s & "  Edicion LONG NOT NULL," & vbCrLf
    s = s & "  BuildId LONG NOT NULL," & vbCrLf
    s = s & "  CodigoRiesgo TEXT(50) NOT NULL," & vbCrLf
    s = s & "  EstadoHtml MEMO," & vbCrLf
    s = s & "  MitigacionHtml MEMO," & vbCrLf
    s = s & "  ContingenciaHtml MEMO," & vbCrLf
    s = s & "  CONSTRAINT PK_TbCacheControlCambiosRow PRIMARY KEY (IDProyecto, Edicion, BuildId, CodigoRiesgo)" & vbCrLf
    s = s & ");" & vbCrLf & vbCrLf
    
    s = s & "CREATE INDEX IX_TbCacheControlCambiosRow_ProjEdBuild ON TbCacheControlCambiosRow (IDProyecto, Edicion, BuildId);" & vbCrLf
    s = s & "CREATE INDEX IX_TbCacheControlCambiosRow_Codigo ON TbCacheControlCambiosRow (CodigoRiesgo);" & vbCrLf
    
    ControlCambios_GetCacheDDL = s
End Function

Public Function ControlCambios_GetFilasHastaEdicion( _
                                                    ByVal p_EdicionActual As Edicion, _
                                                    ByVal p_Proyecto As Proyecto, _
                                                    ByVal p_UsarCache As EnumSiNo, _
                                                    Optional ByRef p_Error As String _
                                                    ) As Collection
    Set ControlCambios_GetFilasHastaEdicion = ControlCambios_GetFilasHastaEdicionConAlcance( _
                                                p_EdicionActual, _
                                                p_Proyecto, _
                                                p_UsarCache, _
                                                EnumControlCambiosAlcanceCompleto, _
                                                p_Error)
End Function

Public Function ControlCambios_GetFilasHastaEdicionConAlcance( _
                                                            ByVal p_EdicionActual As Edicion, _
                                                            ByVal p_Proyecto As Proyecto, _
                                                            ByVal p_UsarCache As EnumSiNo, _
                                                            Optional ByVal p_Alcance As EnumControlCambiosAlcance = EnumControlCambiosAlcanceCompleto, _
                                                            Optional ByRef p_Error As String _
                                                            ) As Collection
    Dim filas As Collection
    Dim edByNum As Scripting.Dictionary
    Dim minEd As Long, maxEd As Long, i As Long, minEdIncluida As Long
    Dim kEd As Variant
    Dim ed As Edicion
    Dim IDProyecto As Long
    Dim cacheOk As Boolean
    Dim tmp As Collection
    
    On Error GoTo errores
    p_Error = ""
    
    If p_Proyecto Is Nothing Or p_EdicionActual Is Nothing Then
        Set ControlCambios_GetFilasHastaEdicionConAlcance = New Collection
        Exit Function
    End If
    
    maxEd = ControlCambios_ToLongSafe(p_EdicionActual.Edicion, 0)
    If maxEd <= 0 Then
        Set ControlCambios_GetFilasHastaEdicionConAlcance = New Collection
        Exit Function
    End If
    
    Set edByNum = New Scripting.Dictionary
    edByNum.CompareMode = TextCompare
    minEd = 2147483647
    
    For Each kEd In p_Proyecto.colEdiciones
        Set ed = p_Proyecto.colEdiciones(kEd)
        If Not ed Is Nothing Then
            If IsNumeric(ed.Edicion) Then
                If CLng(ed.Edicion) < minEd Then minEd = CLng(ed.Edicion)
                edByNum.Add CStr(ed.Edicion), ed
            End If
        End If
    Next kEd
    
    If minEd = 2147483647 Then
        Set ControlCambios_GetFilasHastaEdicionConAlcance = New Collection
        Exit Function
    End If

    minEdIncluida = ControlCambios_CalcularEdicionMinimaParaAlcance(maxEd, minEd, p_Alcance)
    
    IDProyecto = ControlCambios_ToLongSafe(p_Proyecto.IDProyecto, 0)
    cacheOk = (p_UsarCache <> EnumSiNo.No) And (IDProyecto > 0) And ControlCambios_CacheSchemaDisponible()
    
    Set filas = New Collection
    
    For i = minEdIncluida To maxEd
        If edByNum.Exists(CStr(i)) Then
            Set ed = edByNum(CStr(i))
            Dim usarCacheEstaEd As Boolean
            usarCacheEstaEd = cacheOk
            If i = maxEd Then
                If Not IsDate(ed.FechaPublicacion) Then
                    usarCacheEstaEd = False
                End If
            End If
            If usarCacheEstaEd Then
                Set tmp = ControlCambios_Cache_LeerFilas(IDProyecto, i, p_Error)
                If p_Error <> "" Then Err.Raise 1000
                If Not tmp Is Nothing Then
                    ControlCambios_AppendCollection filas, tmp
                    GoTo siguienteEdicion
                End If
            End If
            
            Set tmp = ControlCambios_CalcularFilasEdicion(i, edByNum, p_Error)
            If p_Error <> "" Then Err.Raise 1000
            
            If usarCacheEstaEd Then
                ControlCambios_Cache_GuardarFilas IDProyecto, i, tmp, p_Error
                If p_Error <> "" Then Err.Raise 1000
            End If
            
            ControlCambios_AppendCollection filas, tmp
        End If
siguienteEdicion:
    Next i
    
    Set ControlCambios_GetFilasHastaEdicionConAlcance = filas
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en ControlCambios_GetFilasHastaEdicionConAlcance: " & Err.Description
    End If
End Function

Public Function ControlCambios_CalcularEdicionMinimaParaAlcance( _
                                                            ByVal p_MaxEdicion As Long, _
                                                            ByVal p_MinDisponible As Long, _
                                                            Optional ByVal p_Alcance As EnumControlCambiosAlcance = EnumControlCambiosAlcanceCompleto _
                                                            ) As Long
    Dim minCalculada As Long

    minCalculada = p_MinDisponible
    If p_Alcance = EnumControlCambiosAlcanceResumen3 Then
        minCalculada = p_MaxEdicion - 2
        If minCalculada < p_MinDisponible Then minCalculada = p_MinDisponible
    End If

    If minCalculada < 1 Then minCalculada = 1
    ControlCambios_CalcularEdicionMinimaParaAlcance = minCalculada
End Function

Public Function ControlCambios_ResolverAlcanceDesdeSelector( _
                                                            ByVal p_ValorSelector As String, _
                                                            Optional ByVal p_Default As EnumControlCambiosAlcance = EnumControlCambiosAlcanceResumen3 _
                                                            ) As EnumControlCambiosAlcance
    Dim valor As String

    valor = UCase$(Trim$(Nz(p_ValorSelector, "")))
    Select Case valor
        Case "HISTORICO", "HISTÓRICO", "HISTORICO COMPLETO", "HISTÓRICO COMPLETO"
            ControlCambios_ResolverAlcanceDesdeSelector = EnumControlCambiosAlcanceCompleto
        Case Else
            ControlCambios_ResolverAlcanceDesdeSelector = p_Default
    End Select
End Function

Public Function ControlCambios_DebeUsarHistoricoCompleto( _
                                                        Optional ByVal p_Alcance As EnumControlCambiosAlcance = EnumControlCambiosAlcanceResumen3 _
                                                        ) As Boolean
    ControlCambios_DebeUsarHistoricoCompleto = (p_Alcance = EnumControlCambiosAlcanceCompleto)
End Function

Private Function ControlCambios_CalcularFilasEdicion( _
                                                    ByVal p_EdicionNum As Long, _
                                                    ByVal p_EdByNum As Scripting.Dictionary, _
                                                    Optional ByRef p_Error As String _
                                                    ) As Collection
    Dim filas As Collection
    Dim ed As Edicion
    Dim prevEd As Edicion
    Dim mapAct As Scripting.Dictionary
    Dim mapPrev As Scripting.Dictionary
    Dim codigos() As String
    Dim idx As Long
    Dim codigo As String
    Dim r As riesgo
    Dim rPrev As riesgo
    Dim estadoHtml As String
    Dim pmHtml As String
    Dim pcHtml As String
    Dim isNew As Boolean
    Dim fila As Scripting.Dictionary
    
    On Error GoTo errores
    p_Error = ""
    'If p_EdicionNum = 16 Then Stop
    Set filas = New Collection
    Set ed = p_EdByNum(CStr(p_EdicionNum))
    
    Set mapAct = ControlCambios_MapRiesgosPorCodigo(ed, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    
    If p_EdByNum.Exists(CStr(p_EdicionNum - 1)) Then
        Set prevEd = p_EdByNum(CStr(p_EdicionNum - 1))
        Set mapPrev = ControlCambios_MapRiesgosPorCodigo(prevEd, p_Error)
        If p_Error <> "" Then Err.Raise 1000
    Else
        Set mapPrev = New Scripting.Dictionary
        mapPrev.CompareMode = TextCompare
    End If
    
    codigos = ControlCambios_SortedStringKeys(mapAct)
    For idx = LBound(codigos) To UBound(codigos)
        codigo = codigos(idx)
        'If codigo = "R019" Then Stop
        Set r = mapAct(codigo)
        If r Is Nothing Then GoTo siguienteRiesgo
        
        isNew = (p_EdicionNum = ControlCambios_ToLongSafe(ed.Edicion, p_EdicionNum) And Not mapPrev.Exists(codigo))
        If mapPrev.Exists(codigo) Then
            Set rPrev = mapPrev(codigo)
        Else
            Set rPrev = Nothing
        End If
        
        On Error GoTo errores
        
        estadoHtml = ControlCambios_EstadoHtml(r, rPrev, isNew)
        pmHtml = ControlCambios_PlanesHtml(r, rPrev, True, isNew)
        pcHtml = ControlCambios_PlanesHtml(r, rPrev, False, isNew)
        
        If isNew Or (Trim$(estadoHtml) <> "" Or Trim$(pmHtml) <> "" Or Trim$(pcHtml) <> "") Then
            Set fila = New Scripting.Dictionary
            fila.CompareMode = TextCompare
            fila.Add "CodigoRiesgo", codigo
            fila.Add "Edicion", CStr(p_EdicionNum)
            fila.Add "EstadoHtml", estadoHtml
            fila.Add "MitigacionHtml", pmHtml
            fila.Add "ContingenciaHtml", pcHtml
            filas.Add fila
        End If
        
siguienteRiesgo:
        Set r = Nothing
        Set rPrev = Nothing
    Next idx
    
    Set ControlCambios_CalcularFilasEdicion = filas
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en ControlCambios_CalcularFilasEdicion: " & Err.Description
    End If
End Function

Private Function ControlCambios_MapRiesgosPorCodigo(ByVal p_Edicion As Edicion, Optional ByRef p_Error As String) As Scripting.Dictionary
    Dim col As Scripting.Dictionary
    Dim map As Scripting.Dictionary
    Dim k As Variant
    Dim r As riesgo
    Dim codigo As String
    
    On Error GoTo errores
    p_Error = ""
    
    Set map = New Scripting.Dictionary
    map.CompareMode = TextCompare
    
    If p_Edicion Is Nothing Then
        Set ControlCambios_MapRiesgosPorCodigo = map
        Exit Function
    End If
    
    Set col = p_Edicion.colRiesgos
    If col Is Nothing Then
        Set ControlCambios_MapRiesgosPorCodigo = map
        Exit Function
    End If
    
    For Each k In col
        Set r = col(k)
        If Not r Is Nothing Then
            codigo = Trim$(Nz(r.CodigoRiesgo, ""))
            If codigo <> "" Then
                If Not map.Exists(codigo) Then
                    map.Add codigo, r
                End If
            End If
        End If
        Set r = Nothing
    Next k
    
    Set ControlCambios_MapRiesgosPorCodigo = map
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en ControlCambios_MapRiesgosPorCodigo: " & Err.Description
    End If
End Function

Private Function ControlCambios_EstadoHtml(ByVal r As riesgo, ByVal rPrev As riesgo, ByVal p_Nuevo As Boolean) As String
    Dim s As String
    Dim vAct As String
    Dim vPrev As String
    
    vAct = ControlCambios_NzStr(r.DetectadoPor)
    If rPrev Is Nothing Then
        vPrev = ""
    Else
        vPrev = ControlCambios_NzStr(rPrev.DetectadoPor)
    End If
    ControlCambios_AddLinea s, "Detectado por", vAct, vPrev, p_Nuevo
    
    vAct = ControlCambios_OrigenRiesgo(r)
    vPrev = ControlCambios_OrigenRiesgo(rPrev)
    ControlCambios_AddLinea s, "Origen", vAct, vPrev, p_Nuevo
    
    vAct = ControlCambios_NzStr(r.ImpactoGlobal)
    If rPrev Is Nothing Then
        vPrev = ""
    Else
        vPrev = ControlCambios_NzStr(rPrev.ImpactoGlobal)
    End If
    ControlCambios_AddLinea s, "Impacto global", vAct, vPrev, p_Nuevo
    
    vAct = ControlCambios_NzStr(r.Vulnerabilidad)
    If rPrev Is Nothing Then
        vPrev = ""
    Else
        vPrev = ControlCambios_NzStr(rPrev.Vulnerabilidad)
    End If
    ControlCambios_AddLinea s, "Vulnerabilidad", vAct, vPrev, p_Nuevo
    
    vAct = ControlCambios_NzStr(r.Valoracion)
    If rPrev Is Nothing Then
        vPrev = ""
    Else
        vPrev = ControlCambios_NzStr(rPrev.Valoracion)
    End If
    ControlCambios_AddLinea s, "Valoración", vAct, vPrev, p_Nuevo
    
    vAct = ControlCambios_NzStr(r.Mitigacion)
    If rPrev Is Nothing Then
        vPrev = ""
    Else
        vPrev = ControlCambios_NzStr(rPrev.Mitigacion)
    End If
    ControlCambios_AddLinea s, "Mitigación", vAct, vPrev, p_Nuevo
    
    vAct = ControlCambios_RequiereContingenciaTexto(r)
    vPrev = ControlCambios_RequiereContingenciaTexto(rPrev)
    ControlCambios_AddLinea s, "Contingencia", vAct, vPrev, p_Nuevo
    
    Dim estadoAct As String
    Dim estadoPrev As String
    Dim hayCambioEstado As Boolean

    ' ... (resto del código anterior hasta Materialización)
    vAct = ControlCambios_MaterializacionTexto(r)
    vPrev = ControlCambios_MaterializacionTexto(rPrev)
    ControlCambios_AddLinea s, "Materialización", vAct, vPrev, p_Nuevo
    
    ' Estado
    estadoAct = ControlCambios_EstadoTexto(r)
    estadoPrev = ControlCambios_EstadoTexto(rPrev)
    hayCambioEstado = (estadoAct <> estadoPrev)
    ControlCambios_AddLinea s, "Estado", estadoAct, estadoPrev, p_Nuevo
    
    ' Fecha estado: Solo si hay cambio de estado o es nuevo
    vAct = ControlCambios_FormatoFecha(r.FechaEstado)
    If rPrev Is Nothing Then
        vPrev = ""
    Else
        vPrev = ControlCambios_FormatoFecha(rPrev.FechaEstado)
    End If
    
    If p_Nuevo Or hayCambioEstado Then
        ControlCambios_AddLinea s, "Fecha estado", vAct, vPrev, p_Nuevo
    End If
    
    vAct = ControlCambios_NzStr(r.Priorizacion)
    If rPrev Is Nothing Then
        vPrev = ""
    Else
        vPrev = ControlCambios_NzStr(rPrev.Priorizacion)
    End If
    ControlCambios_AddLinea s, "Priorización", vAct, vPrev, p_Nuevo
    
    If p_Nuevo Then
        s = C_CC_ADD_OPEN & s & C_CC_ADD_CLOSE
    End If
    
    ControlCambios_EstadoHtml = ControlCambios_HTMLSafeLargo(Trim$(s))
End Function

Private Sub ControlCambios_AddLinea(ByRef p_S As String, ByVal p_Label As String, ByVal p_Act As String, ByVal p_Prev As String, ByVal p_Nuevo As Boolean)
    If p_Nuevo Then
        p_S = p_S & p_Label & ": " & p_Act & vbCrLf
    Else
        If ControlCambios_DiffValoresDifieren(p_Act, p_Prev) Then
            p_S = p_S & p_Label & ": " & p_Act & vbCrLf
        End If
    End If
End Sub

Public Function ControlCambios_DiffValoresDifieren(ByVal p_Actual As String, ByVal p_Previa As String) As Boolean
    ControlCambios_DiffValoresDifieren = (ControlCambios_NormalizarDiffValor(p_Actual) <> ControlCambios_NormalizarDiffValor(p_Previa))
End Function

Private Function ControlCambios_NormalizarDiffValor(ByVal p_Valor As String) As String
    Dim t As String

    t = Nz(p_Valor, "")
    t = Replace(t, vbCrLf, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")

    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop

    ControlCambios_NormalizarDiffValor = LCase$(Trim$(t))
End Function

Private Function ControlCambios_PlanesHtml(ByVal r As riesgo, ByVal rPrev As riesgo, ByVal p_Mitigacion As Boolean, ByVal p_NuevoRiesgo As Boolean) As String
    Dim colAct As Scripting.Dictionary
    Dim colPrev As Scripting.Dictionary
    
    On Error GoTo salir
    
    If r Is Nothing Then Exit Function
    
    If p_Mitigacion Then
        Set colAct = r.ColPMs
        If Not rPrev Is Nothing Then Set colPrev = rPrev.ColPMs
    Else
        Set colAct = r.ColPCs
        If Not rPrev Is Nothing Then Set colPrev = rPrev.ColPCs
    End If
    
    ControlCambios_PlanesHtml = ControlCambios_PlanesHtmlDesdeColecciones(colAct, colPrev, p_Mitigacion, p_NuevoRiesgo)
    Exit Function
salir:
    ControlCambios_PlanesHtml = ""
End Function

Private Function ControlCambios_PlanesHtmlDesdeColecciones( _
                                                        ByVal p_ColAct As Scripting.Dictionary, _
                                                        ByVal p_ColPrev As Scripting.Dictionary, _
                                                        ByVal p_Mitigacion As Boolean, _
                                                        ByVal p_NuevoRiesgo As Boolean _
                                                        ) As String
    Dim s As String
    Dim plan As Object
    Dim planPrev As Object
    Dim k As Variant
    Dim idPlanPrev As String
    Dim mapPrevFound As Scripting.Dictionary
    Dim planBlock As String
    Dim isNewPlanFull As Boolean
    
    Dim isActEmpty As Boolean
    Dim isPrevEmpty As Boolean
    
    isActEmpty = True
    If Not p_ColAct Is Nothing Then
        If p_ColAct.Count > 0 Then isActEmpty = False
    End If
    
    isPrevEmpty = True
    If Not p_ColPrev Is Nothing Then
        If p_ColPrev.Count > 0 Then isPrevEmpty = False
    End If
    
    If isActEmpty And isPrevEmpty Then
        ControlCambios_PlanesHtmlDesdeColecciones = ""
        Exit Function
    End If
    
    Set mapPrevFound = New Scripting.Dictionary
    mapPrevFound.CompareMode = TextCompare
    
    ' 1. Procesar planes actuales (Nuevos y Modificados)
    If Not p_ColAct Is Nothing Then
        For Each k In p_ColAct
            Set plan = p_ColAct(k)
            If Not plan Is Nothing Then
                Set planPrev = Nothing
                
                ' Si es un riesgo nuevo, no buscamos plan anterior para forzar modo "Nuevo Completo"
                ' aunque existiera (por corrección de datos o similar)
                If Not p_NuevoRiesgo Then
                    On Error Resume Next
                    If p_Mitigacion Then
                        Set planPrev = plan.PMEdicionAnterior
                    Else
                        Set planPrev = plan.PCEdicionAnterior
                    End If
                    On Error GoTo 0
                End If
                
                If Not planPrev Is Nothing Then
                    idPlanPrev = ControlCambios_GetPlanId(planPrev, p_Mitigacion)
                    If idPlanPrev <> "" Then
                        If Not mapPrevFound.Exists(idPlanPrev) Then mapPrevFound.Add idPlanPrev, True
                    End If
                End If
                
                isNewPlanFull = p_NuevoRiesgo Or (planPrev Is Nothing)
                planBlock = ControlCambios_PlanBlockHtml(plan, planPrev, p_Mitigacion, isNewPlanFull)
                
                If Trim$(planBlock) <> "" Then
                    If s <> "" Then s = s & vbCrLf & vbCrLf
                    s = s & planBlock
                End If
            End If
            Set plan = Nothing
            Set planPrev = Nothing
        Next k
    End If
    
    ' 2. Procesar planes eliminados (Estaban en Prev pero no han sido reclamados por ninguno de Act)
    If Not p_ColPrev Is Nothing Then
        For Each k In p_ColPrev
            Set planPrev = p_ColPrev(k)
            If Not planPrev Is Nothing Then
                idPlanPrev = ControlCambios_GetPlanId(planPrev, p_Mitigacion)
                If idPlanPrev <> "" Then
                    If Not mapPrevFound.Exists(idPlanPrev) Then
                        ' Es un plan eliminado
                        planBlock = ControlCambios_PlanEliminadoBlockHtml(planPrev, p_Mitigacion)
                        If Trim$(planBlock) <> "" Then
                            If s <> "" Then s = s & vbCrLf & vbCrLf
                            s = s & planBlock
                        End If
                    End If
                End If
            End If
            Set planPrev = Nothing
        Next k
    End If
    
    ControlCambios_PlanesHtmlDesdeColecciones = ControlCambios_HTMLSafeLargo(Trim$(s))
End Function

Private Function ControlCambios_PlanBlockHtml(ByVal p_Plan As Object, ByVal p_PlanPrev As Object, ByVal p_Mitigacion As Boolean, ByVal p_NuevoPlan As Boolean) As String
    Dim s As String
    Dim header As String
    Dim body As String
    Dim vAct As String
    Dim vPrev As String
    Dim accionesAct As Scripting.Dictionary
    Dim accionesPrev As Scripting.Dictionary
    Dim k As Variant
    Dim acc As Object
    Dim accPrev As Object
    Dim idAccPrev As String
    Dim accBlock As String
    Dim hayAccionesCambiadas As Boolean
    Dim mapPrevFound As Scripting.Dictionary
    Dim isNewAccFull As Boolean
    
    header = ControlCambios_PlanHeader(p_Plan, p_Mitigacion)
    If header = "" Then Exit Function
    
    body = ""
    
    vAct = ControlCambios_GetPlanDisparador(p_Plan)
    vPrev = ControlCambios_GetPlanDisparador(p_PlanPrev)
    If Not (p_NuevoPlan And vAct = "") Then
        ControlCambios_AddLinea body, "Disparador", vAct, vPrev, p_NuevoPlan
    End If
    
    vAct = ControlCambios_FormatoFecha(ControlCambios_GetPlanFechaActivacion(p_Plan))
    vPrev = ControlCambios_FormatoFecha(ControlCambios_GetPlanFechaActivacion(p_PlanPrev))
    If Not (p_NuevoPlan And vAct = "") Then
        ControlCambios_AddLinea body, "Activación", vAct, vPrev, p_NuevoPlan
    End If
    
    vAct = ControlCambios_FormatoFecha(ControlCambios_GetPlanFechaDesactivacion(p_Plan))
    vPrev = ControlCambios_FormatoFecha(ControlCambios_GetPlanFechaDesactivacion(p_PlanPrev))
    If Not (p_NuevoPlan And vAct = "") Then
        ControlCambios_AddLinea body, "Desactivación", vAct, vPrev, p_NuevoPlan
    End If
    
    Set mapPrevFound = New Scripting.Dictionary
    mapPrevFound.CompareMode = TextCompare
    
    On Error Resume Next
    Set accionesAct = p_Plan.colAcciones
    If Not p_PlanPrev Is Nothing Then Set accionesPrev = p_PlanPrev.colAcciones
    On Error GoTo 0
    
    ' 1. Procesar acciones actuales (Nuevas y Modificadas)
    If Not accionesAct Is Nothing Then
        For Each k In accionesAct
            Set acc = accionesAct(k)
            If Not acc Is Nothing Then
                Set accPrev = Nothing
                On Error Resume Next
                If p_Mitigacion Then
                    Set accPrev = acc.PMAccionEdicionAnterior
                Else
                    Set accPrev = acc.PCAccionEdicionAnterior
                End If
                On Error GoTo 0
                
                If Not accPrev Is Nothing Then
                    idAccPrev = ControlCambios_GetAccionId(accPrev, p_Mitigacion)
                    If idAccPrev <> "" Then
                        If Not mapPrevFound.Exists(idAccPrev) Then mapPrevFound.Add idAccPrev, True
                    End If
                End If
                
                isNewAccFull = p_NuevoPlan Or (accPrev Is Nothing)
                accBlock = ControlCambios_AccionBlockHtml(acc, accPrev, p_Mitigacion, isNewAccFull)
                
                If Trim$(accBlock) <> "" Then
                    hayAccionesCambiadas = True
                    body = body & accBlock
                End If
            End If
            Set acc = Nothing
            Set accPrev = Nothing
        Next k
    End If
    
    ' 2. Procesar acciones eliminadas
    If Not accionesPrev Is Nothing Then
        For Each k In accionesPrev
            Set accPrev = accionesPrev(k)
            If Not accPrev Is Nothing Then
                idAccPrev = ControlCambios_GetAccionId(accPrev, p_Mitigacion)
                If idAccPrev <> "" Then
                    If Not mapPrevFound.Exists(idAccPrev) Then
                        accBlock = ControlCambios_AccionEliminadaBlockHtml(accPrev, p_Mitigacion)
                        If Trim$(accBlock) <> "" Then
                            hayAccionesCambiadas = True
                            body = body & accBlock
                        End If
                    End If
                End If
            End If
            Set accPrev = Nothing
        Next k
    End If
    
    If p_NuevoPlan Then
        ControlCambios_PlanBlockHtml = C_CC_ADD_OPEN & header & vbCrLf & body & C_CC_ADD_CLOSE
        Exit Function
    End If
    
    If Trim$(body) <> "" Then
        ControlCambios_PlanBlockHtml = header & vbCrLf & body
    Else
        ControlCambios_PlanBlockHtml = ""
    End If
End Function

Private Function ControlCambios_PlanEliminadoBlockHtml(ByVal p_PlanPrev As Object, ByVal p_Mitigacion As Boolean) As String
    Dim s As String
    Dim header As String
    
    header = ControlCambios_PlanHeader(p_PlanPrev, p_Mitigacion)
    If header = "" Then Exit Function
    
    s = C_CC_DEL_OPEN & header & " eliminado" & C_CC_DEL_CLOSE
    ControlCambios_PlanEliminadoBlockHtml = s
End Function

Private Function ControlCambios_AccionBlockHtml(ByVal p_Acc As Object, ByVal p_AccPrev As Object, ByVal p_Mitigacion As Boolean, ByVal p_Nuevo As Boolean) As String
    Dim s As String
    Dim body As String
    Dim vAct As String
    Dim vPrev As String
    Dim idAcc As String
    
    idAcc = ControlCambios_GetAccionId(p_Acc, p_Mitigacion)
    If idAcc = "" Then Exit Function
    
    body = ""
    
    vAct = ControlCambios_NzStr(p_Acc.Accion)
    If p_AccPrev Is Nothing Then
        vPrev = ""
    Else
        vPrev = ControlCambios_NzStr(p_AccPrev.Accion)
    End If
    If Not (p_Nuevo And vAct = "") Then
        ControlCambios_AddLinea body, "Descripción", vAct, vPrev, p_Nuevo
    End If
    
    vAct = ControlCambios_NzStr(p_Acc.ResponsableAccion)
    If p_AccPrev Is Nothing Then
        vPrev = ""
    Else
        vPrev = ControlCambios_NzStr(p_AccPrev.ResponsableAccion)
    End If
    If Not (p_Nuevo And vAct = "") Then
        ControlCambios_AddLinea body, "Responsable", vAct, vPrev, p_Nuevo
    End If
    
    vAct = ControlCambios_FormatoFecha(p_Acc.FechaInicio)
    If p_AccPrev Is Nothing Then
        vPrev = ""
    Else
        vPrev = ControlCambios_FormatoFecha(p_AccPrev.FechaInicio)
    End If
    If Not (p_Nuevo And vAct = "") Then
        ControlCambios_AddLinea body, "Fecha inicio", vAct, vPrev, p_Nuevo
    End If
    
    vAct = ControlCambios_FormatoFecha(p_Acc.FechaFinPrevista)
    If p_AccPrev Is Nothing Then
        vPrev = ""
    Else
        vPrev = ControlCambios_FormatoFecha(p_AccPrev.FechaFinPrevista)
    End If
    If Not (p_Nuevo And vAct = "") Then
        ControlCambios_AddLinea body, "Fecha fin prevista", vAct, vPrev, p_Nuevo
    End If
    
    vAct = ControlCambios_FormatoFecha(p_Acc.FechaFinReal)
    If p_AccPrev Is Nothing Then
        vPrev = ""
    Else
        vPrev = ControlCambios_FormatoFecha(p_AccPrev.FechaFinReal)
    End If
    If Not (p_Nuevo And vAct = "") Then
        ControlCambios_AddLinea body, "Fecha fin real", vAct, vPrev, p_Nuevo
    End If
    
    If body = "" And Not p_Nuevo Then
        ControlCambios_AccionBlockHtml = ""
        Exit Function
    End If
    
    s = "- " & idAcc & vbCrLf & body
    
    If p_Nuevo Then
        s = C_CC_ADD_OPEN & s & C_CC_ADD_CLOSE
    End If
    
    ControlCambios_AccionBlockHtml = s
End Function

Private Function ControlCambios_AccionEliminadaBlockHtml(ByVal p_AccPrev As Object, ByVal p_Mitigacion As Boolean) As String
    Dim s As String
    Dim idAcc As String
    
    idAcc = ControlCambios_GetAccionId(p_AccPrev, p_Mitigacion)
    If idAcc = "" Then Exit Function
    
    s = C_CC_DEL_OPEN & "- " & idAcc & " eliminado" & C_CC_DEL_CLOSE & vbCrLf
    ControlCambios_AccionEliminadaBlockHtml = s
End Function

Private Function ControlCambios_PlanHeader(ByVal p_Plan As Object, ByVal p_Mitigacion As Boolean) As String
    Dim idPlan As String
    Dim codPlan As String
    
    idPlan = ControlCambios_GetPlanId(p_Plan, p_Mitigacion)
    If idPlan = "" Then Exit Function
    
    On Error Resume Next
    If Not p_Plan Is Nothing Then
        If p_Mitigacion Then
            codPlan = Trim$(Nz(p_Plan.CodMitigacion, ""))
        Else
            codPlan = Trim$(Nz(p_Plan.CodContingencia, ""))
        End If
    End If
    On Error GoTo 0
    If codPlan <> "" Then
        ControlCambios_PlanHeader = IIf(p_Mitigacion, "PM ", "PC ") & idPlan & " (" & codPlan & ")"
    Else
        ControlCambios_PlanHeader = IIf(p_Mitigacion, "PM ", "PC ") & idPlan
    End If
End Function

Private Function ControlCambios_GetPlanId(ByVal p_Plan As Object, ByVal p_Mitigacion As Boolean) As String
    On Error Resume Next
    If p_Plan Is Nothing Then Exit Function
    If p_Mitigacion Then
        ControlCambios_GetPlanId = Trim$(Nz(p_Plan.IDMitigacion, ""))
    Else
        ControlCambios_GetPlanId = Trim$(Nz(p_Plan.IDContingencia, ""))
    End If
    If ControlCambios_GetPlanId = "" Then
        Dim mCod As String
        On Error Resume Next
        If p_Mitigacion Then
            mCod = Trim$(Nz(p_Plan.CodMitigacion, ""))
        Else
            mCod = Trim$(Nz(p_Plan.CodContingencia, ""))
        End If
        On Error GoTo 0
        ControlCambios_GetPlanId = mCod
    End If
End Function

Private Function ControlCambios_GetAccionId(ByVal p_Acc As Object, ByVal p_Mitigacion As Boolean) As String
    On Error Resume Next
    If p_Acc Is Nothing Then Exit Function
    ControlCambios_GetAccionId = Trim$(Nz(p_Acc.CodAccion, ""))
    If ControlCambios_GetAccionId = "" Then
        If p_Mitigacion Then
            ControlCambios_GetAccionId = Trim$(Nz(p_Acc.IDAccionMitigacion, ""))
        Else
            ControlCambios_GetAccionId = Trim$(Nz(p_Acc.IDAccionContingencia, ""))
        End If
    End If
End Function

Private Function ControlCambios_GetPlanDisparador(ByVal p_Plan As Object) As String
    On Error Resume Next
    If p_Plan Is Nothing Then Exit Function
    ControlCambios_GetPlanDisparador = Trim$(Nz(p_Plan.DisparadorDelPlan, ""))
End Function

Private Function ControlCambios_GetPlanFechaActivacion(ByVal p_Plan As Object) As Variant
    On Error Resume Next
    If p_Plan Is Nothing Then Exit Function
    ControlCambios_GetPlanFechaActivacion = p_Plan.FechaDeActivacion
End Function

Private Function ControlCambios_GetPlanFechaDesactivacion(ByVal p_Plan As Object) As Variant
    On Error Resume Next
    If p_Plan Is Nothing Then Exit Function
    ControlCambios_GetPlanFechaDesactivacion = p_Plan.FechaDesactivacion
End Function

Private Function ControlCambios_HayLineasCambio(ByVal p_S As String, ByVal p_Header As String) As Boolean
    Dim t As String
    t = Replace(p_S, p_Header & vbCrLf, "")
    t = Trim$(Replace(t, vbCrLf, ""))
    ControlCambios_HayLineasCambio = (t <> "")
End Function

Private Function ControlCambios_OrigenRiesgo(ByVal r As riesgo) As String
    On Error Resume Next
    If r Is Nothing Then Exit Function
    ControlCambios_OrigenRiesgo = Trim$(Nz(r.Origen, ""))
    If ControlCambios_OrigenRiesgo = "" Then ControlCambios_OrigenRiesgo = Trim$(Nz(r.CausaRaiz, ""))
End Function

Private Function ControlCambios_RequiereContingenciaTexto(ByVal r As riesgo) As String
    On Error Resume Next
    If r Is Nothing Then Exit Function
    ControlCambios_RequiereContingenciaTexto = Trim$(Nz(r.RequierePlanContingenciaCalculadoTexto, ""))
    If ControlCambios_RequiereContingenciaTexto = "" Then ControlCambios_RequiereContingenciaTexto = Trim$(Nz(r.RequierePlanContingencia, ""))
End Function

Private Function ControlCambios_MaterializacionTexto(ByVal r As riesgo) As String
    Dim e As EnumRiesgoEstado
    On Error Resume Next
    If r Is Nothing Then Exit Function
    e = r.EstadoEnum
    If e = EnumRiesgoEstado.Materializado Then
        ControlCambios_MaterializacionTexto = "Sí"
    Else
        ControlCambios_MaterializacionTexto = "No"
    End If
End Function

Private Function ControlCambios_EstadoTexto(ByVal r As riesgo) As String
    On Error Resume Next
    If r Is Nothing Then Exit Function
    ControlCambios_EstadoTexto = Trim$(Nz(r.ESTADOCalculadoTexto, ""))
    If ControlCambios_EstadoTexto = "" Then ControlCambios_EstadoTexto = Trim$(Nz(r.Estado, ""))
End Function

Private Sub ControlCambios_AppendCollection(ByRef p_Dest As Collection, ByVal p_Src As Collection)
    Dim it As Variant
    If p_Src Is Nothing Then Exit Sub
    For Each it In p_Src
        p_Dest.Add it
    Next
End Sub

Private Function ControlCambios_SortedStringKeys(ByVal p_Dict As Scripting.Dictionary) As String()
    Dim keys() As String
    Dim i As Long
    Dim k As Variant
    
    ReDim keys(0 To 0)
    If p_Dict Is Nothing Then
        ControlCambios_SortedStringKeys = keys
        Exit Function
    End If
    If p_Dict.Count = 0 Then
        ControlCambios_SortedStringKeys = keys
        Exit Function
    End If
    
    ReDim keys(0 To p_Dict.Count - 1)
    i = 0
    For Each k In p_Dict.keys
        keys(i) = CStr(k)
        i = i + 1
    Next k
    
    ControlCambios_QuickSortStrings keys, LBound(keys), UBound(keys)
    ControlCambios_SortedStringKeys = keys
End Function

Private Sub ControlCambios_QuickSortStrings(ByRef arr() As String, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long
    Dim pivot As String
    Dim temp As String
    
    i = first
    j = last
    pivot = arr((first + last) \ 2)
    
    Do While i <= j
        Do While arr(i) < pivot
            i = i + 1
        Loop
        Do While arr(j) > pivot
            j = j - 1
        Loop
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop
    
    If first < j Then ControlCambios_QuickSortStrings arr, first, j
    If i < last Then ControlCambios_QuickSortStrings arr, i, last
End Sub

Private Function ControlCambios_ToLongSafe(ByVal p_Value As Variant, ByVal p_Default As Long) As Long
    If IsNumeric(p_Value) Then
        ControlCambios_ToLongSafe = CLng(p_Value)
    Else
        ControlCambios_ToLongSafe = p_Default
    End If
End Function

Private Function ControlCambios_NzStr(ByVal p_Value As Variant) As String
    ControlCambios_NzStr = Trim$(Nz(p_Value, ""))
End Function

Private Function ControlCambios_FormatoFecha(ByVal p_Valor As Variant) As String
    If IsDate(p_Valor) Then
        ControlCambios_FormatoFecha = Format$(CDate(p_Valor), "dd/mm/yyyy")
    Else
        ControlCambios_FormatoFecha = ""
    End If
End Function

Private Function ControlCambios_NzHtml(ByVal p_HTML As String) As String
    If Trim$(p_HTML) = "" Then
        ControlCambios_NzHtml = "&nbsp;"
    Else
        ControlCambios_NzHtml = p_HTML
    End If
End Function

Private Function ControlCambios_HTMLSafe(ByVal p_Texto As String) As String
    Dim t As String
    t = Nz(p_Texto, "")
    t = Replace(t, "&", "&amp;")
    t = Replace(t, "<", "&lt;")
    t = Replace(t, ">", "&gt;")
    t = Replace(t, """", "&quot;")
    ControlCambios_HTMLSafe = t
End Function

Private Function ControlCambios_HTMLSafeLargo(ByVal p_Texto As String) As String
    Dim t As String
    t = ControlCambios_HTMLSafe(p_Texto)
    t = Replace(t, vbCrLf, "<br>")
    t = Replace(t, vbLf, "<br>")
    t = Replace(t, C_CC_ADD_OPEN, "<span class='cc-added'>")
    t = Replace(t, C_CC_ADD_CLOSE, "</span>")
    t = Replace(t, C_CC_DEL_OPEN, "<span class='cc-deleted'>")
    t = Replace(t, C_CC_DEL_CLOSE, "</span>")
    ControlCambios_HTMLSafeLargo = t
End Function

Private Function ControlCambios_CacheSchemaDisponible() As Boolean
    Dim db As DAO.Database
    Dim errStr As String
    On Error GoTo salir
    Set db = getdb(errStr)
    If errStr <> "" Then GoTo salir
    ControlCambios_CacheSchemaDisponible = ControlCambios_ExisteTabla(db, "TbCacheControlCambiosMeta") And ControlCambios_ExisteTabla(db, "TbCacheControlCambiosRow")
salir:
End Function

Private Function ControlCambios_ExisteTabla(p_db As DAO.Database, p_NombreTabla As String) As Boolean
    Dim tdf As DAO.TableDef
    On Error GoTo errores
    ControlCambios_ExisteTabla = False
    For Each tdf In p_db.TableDefs
        If tdf.Name = p_NombreTabla Then
            ControlCambios_ExisteTabla = True
            Exit Function
        End If
    Next
    Exit Function
errores:
    ControlCambios_ExisteTabla = False
End Function

Private Function ControlCambios_Cache_GetActiveBuildId(p_IDProyecto As Long, p_Edicion As Long, Optional ByRef p_Error As String) As Long
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String
    
    On Error GoTo errores
    p_Error = ""
    ControlCambios_Cache_GetActiveBuildId = 0
    
    Set db = getdb(p_Error)
    If p_Error <> "" Then Err.Raise 1000
    
    sql = "SELECT ActiveBuildId, CacheVersion FROM TbCacheControlCambiosMeta WHERE IDProyecto=" & p_IDProyecto & " AND Edicion=" & p_Edicion & ";"
    Set rcd = db.OpenRecordset(sql)
    If Not rcd.EOF Then
        If Nz(rcd.Fields("CacheVersion"), "") = C_CACHE_LOGIC_VERSION Then
            ControlCambios_Cache_GetActiveBuildId = CLng(Nz(rcd.Fields("ActiveBuildId"), 0))
        Else
            ControlCambios_Cache_GetActiveBuildId = 0
        End If
    End If
    rcd.Close
    Set rcd = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en ControlCambios_Cache_GetActiveBuildId: " & Err.Description
    End If
End Function

Private Function ControlCambios_Cache_SetActiveBuildId(p_IDProyecto As Long, p_Edicion As Long, p_BuildId As Long, Optional ByRef p_Error As String) As String
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String
    
    On Error GoTo errores
    p_Error = ""
    
    Set db = getdb(p_Error)
    If p_Error <> "" Then Err.Raise 1000
    
    sql = "SELECT * FROM TbCacheControlCambiosMeta WHERE IDProyecto=" & p_IDProyecto & " AND Edicion=" & p_Edicion & ";"
    Set rcd = db.OpenRecordset(sql)
    If rcd.EOF Then
        rcd.AddNew
        rcd.Fields("IDProyecto") = p_IDProyecto
        rcd.Fields("Edicion") = p_Edicion
        rcd.Fields("ActiveBuildId") = p_BuildId
        rcd.Fields("UpdatedAt") = Now()
        rcd.Fields("CacheVersion") = C_CACHE_LOGIC_VERSION
        rcd.Update
    Else
        rcd.Edit
        rcd.Fields("ActiveBuildId") = p_BuildId
        rcd.Fields("UpdatedAt") = Now()
        rcd.Fields("CacheVersion") = C_CACHE_LOGIC_VERSION
        rcd.Update
    End If
    rcd.Close
    Set rcd = Nothing
    
    ControlCambios_Cache_SetActiveBuildId = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en ControlCambios_Cache_SetActiveBuildId: " & Err.Description
    End If
End Function

Private Function ControlCambios_Cache_LeerFilas(p_IDProyecto As Long, p_Edicion As Long, Optional ByRef p_Error As String) As Collection
    Dim buildId As Long
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String
    Dim filas As Collection
    Dim fila As Scripting.Dictionary
    
    On Error GoTo errores
    p_Error = ""
    
    buildId = ControlCambios_Cache_GetActiveBuildId(p_IDProyecto, p_Edicion, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    If buildId <= 0 Then
        Set ControlCambios_Cache_LeerFilas = Nothing
        Exit Function
    End If
    
    Set db = getdb(p_Error)
    If p_Error <> "" Then Err.Raise 1000
    
    sql = "SELECT CodigoRiesgo, EstadoHtml, MitigacionHtml, ContingenciaHtml FROM TbCacheControlCambiosRow " & _
          "WHERE IDProyecto=" & p_IDProyecto & " AND Edicion=" & p_Edicion & " AND BuildId=" & buildId & _
          " ORDER BY CodigoRiesgo ASC;"
    Set rcd = db.OpenRecordset(sql)
    
    Set filas = New Collection
    Do While Not rcd.EOF
        Set fila = New Scripting.Dictionary
        fila.CompareMode = TextCompare
        fila.Add "CodigoRiesgo", Nz(rcd.Fields("CodigoRiesgo"), "")
        fila.Add "Edicion", CStr(p_Edicion)
        fila.Add "EstadoHtml", Nz(rcd.Fields("EstadoHtml"), "")
        fila.Add "MitigacionHtml", Nz(rcd.Fields("MitigacionHtml"), "")
        fila.Add "ContingenciaHtml", Nz(rcd.Fields("ContingenciaHtml"), "")
        filas.Add fila
        rcd.MoveNext
    Loop
    rcd.Close
    Set rcd = Nothing
    
    Set ControlCambios_Cache_LeerFilas = filas
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en ControlCambios_Cache_LeerFilas: " & Err.Description
    End If
End Function

Private Sub ControlCambios_Cache_GuardarFilas(p_IDProyecto As Long, p_Edicion As Long, p_Filas As Collection, Optional ByRef p_Error As String)
    Dim db As DAO.Database
    Dim buildId As Long
    Dim fila As Object
    Dim sql As String
    
    On Error GoTo errores
    p_Error = ""
    
    Set db = getdb(p_Error)
    If p_Error <> "" Then Err.Raise 1000
    
    ' Garantizar BuildId único para evitar colisiones
    Do
        buildId = ControlCambios_NuevoBuildId()
        sql = "SELECT Count(*) FROM TbCacheControlCambiosRow WHERE IDProyecto=" & p_IDProyecto & " AND Edicion=" & p_Edicion & " AND BuildId=" & buildId
        If db.OpenRecordset(sql).Fields(0) = 0 Then Exit Do
    Loop
    
    If Not p_Filas Is Nothing Then
        For Each fila In p_Filas
            sql = "INSERT INTO TbCacheControlCambiosRow (IDProyecto, Edicion, BuildId, CodigoRiesgo, EstadoHtml, MitigacionHtml, ContingenciaHtml) VALUES (" & _
                  p_IDProyecto & "," & p_Edicion & "," & buildId & "," & _
                  ControlCambios_SqlText(fila("CodigoRiesgo")) & "," & _
                  ControlCambios_SqlMemo(fila("EstadoHtml")) & "," & _
                  ControlCambios_SqlMemo(fila("MitigacionHtml")) & "," & _
                  ControlCambios_SqlMemo(fila("ContingenciaHtml")) & ");"
            db.Execute sql, dbFailOnError
        Next
    End If
    
    ControlCambios_Cache_SetActiveBuildId p_IDProyecto, p_Edicion, buildId, p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    ' Limpiamos versiones anteriores de la caché para este proyecto/edición para no acumular basura
    sql = "DELETE FROM TbCacheControlCambiosRow WHERE IDProyecto=" & p_IDProyecto & " AND Edicion=" & p_Edicion & " AND BuildId<>" & buildId & ";"
    db.Execute sql
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en ControlCambios_Cache_GuardarFilas: " & Err.Description
    End If
End Sub

Private Function ControlCambios_NuevoBuildId() As Long
    Randomize
    ControlCambios_NuevoBuildId = (CLng(DateDiff("s", #1/1/2020#, Now())) Mod 2000000) * 1000& + CLng((Timer - Fix(Timer)) * 1000)
End Function

Private Function ControlCambios_SqlText(p_Value As Variant) As String
    If IsNull(p_Value) Then
        ControlCambios_SqlText = "NULL"
    Else
        ControlCambios_SqlText = "'" & Replace(CStr(p_Value), "'", "''") & "'"
    End If
End Function

Private Function ControlCambios_SqlMemo(p_Value As Variant) As String
    If IsNull(p_Value) Then
        ControlCambios_SqlMemo = "NULL"
    Else
        ControlCambios_SqlMemo = "'" & Replace(CStr(p_Value), "'", "''") & "'"
    End If
End Function









