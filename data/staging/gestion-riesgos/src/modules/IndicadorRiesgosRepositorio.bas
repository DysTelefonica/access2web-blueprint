Attribute VB_Name = "IndicadorRiesgosRepositorio"
'===============================
' IndicadorRiesgosRepositorioV2.bas
'===============================
Option Compare Database
Option Explicit

' QueryDef temporal (se recrea cada vez)
Private Const QD_NAME As String = "qtmp_IndicadoresRiesgosV2"

' =========================================================
' issue-61: defensa contra inyeccion SQL por p_ListaIDsCsv
'   - Valida y normaliza el CSV a una lista de enteros Long.
'   - Acepta vacio o solo separadores -> devuelve "".
'   - Si encuentra un valor no entero (negativo, decimal, alfanumerico,
'     o un payload tipo "1) OR 1=1 --"), Err.Raise 1000 con mensaje.
'   - NO usar con QueryDef parametrizado: el helper unicamente
'     sanitiza la cadena antes de la concatenacion SQL.
' =========================================================
Public Function ValidarListaIDsCsv(ByVal p_Csv As String) As String
    Dim sTrim As String
    Dim parts() As String
    Dim i As Long
    Dim part As String
    Dim sanitized As String
    Dim count As Long

    sTrim = Trim$(p_Csv)
    If Len(sTrim) = 0 Then
        ValidarListaIDsCsv = ""
        Exit Function
    End If

    parts = Split(sTrim, ",")
    count = 0
    For i = LBound(parts) To UBound(parts)
        part = Trim$(parts(i))
        If Len(part) > 0 Then
            If Not EsEnteroNoNegativo(part) Then
                Err.Raise 1000, "ValidarListaIDsCsv", _
                    "ID invalido en lista CSV: '" & part & "'. " & _
                    "Solo se permiten enteros no negativos."
            End If
            count = count + 1
            If count > 1 Then sanitized = sanitized & ", "
            sanitized = sanitized & CLng(part)
        End If
    Next i

    ValidarListaIDsCsv = sanitized
End Function

Private Function EsEnteroNoNegativo(ByVal p_S As String) As Boolean
    Dim i As Long
    Dim ch As String
    If Len(p_S) = 0 Then Exit Function
    For i = 1 To Len(p_S)
        ch = Mid$(p_S, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    EsEnteroNoNegativo = True
End Function

' Repositorio V2: devuelve recordset con 5 columnas de indicadores (Mario)
'   - Identificados (TbRiesgos.FechaDetectado)
'   - Retirados     (TbRiesgos.FechaRetirado)
'   - EnOferta      (TbRiesgosAIntegrar.FechaDetectado)
'   - Materializados(TbRiesgos.FechaMaterializado)
'   - Oferta->Gestion (TbRiesgosAIntegrar.Trasladar='Sí')
Public Function IndicadorRiesgosRepositorioV2_GetTabla( _
                                                        ByVal p_dIni As Date, _
                                                        ByVal p_dFin As Date, _
                                                        ByVal p_ListaIDsCsv As String, _
                                                        Optional ByRef p_Error As String _
                                                    ) As DAO.Recordset

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim m_SQL As String
    Dim sIni As String, sFin As String
    Dim whereProy As String

    On Error GoTo errores
    p_Error = ""

    ' issue-61: sanitizar CSV contra inyeccion SQL antes de concatenar.
    ' CSV vacio -> whereProy queda vacio (sin filtro de proyectos).
    Dim sanitizedCsv As String
    sanitizedCsv = ValidarListaIDsCsv(p_ListaIDsCsv)

    ' Access requiere literales de fecha en #mm/dd/yyyy#
    sIni = Format$(p_dIni, "mm\/dd\/yyyy")
    sFin = Format$(p_dFin, "mm\/dd\/yyyy")

    whereProy = ""
    If Len(Trim$(sanitizedCsv)) > 0 Then
        whereProy = "WHERE P.IDProyecto In (" & sanitizedCsv & ") "
    End If

    ' NOTA IMPORTANTE:
    ' - NO usamos campos inexistentes (EnOferta/DetectadoEnOferta/PasoAFinal) en TbRiesgos
    ' - La parte "oferta" sale de TbRiesgosAIntegrar (ERD: FechaDetectado, Trasladar)
    m_SQL = ""
    m_SQL = m_SQL & "SELECT " & vbCrLf
    m_SQL = m_SQL & "    P.IDProyecto," & vbCrLf
    m_SQL = m_SQL & "    (P.Proyecto & ' ' & P.NombreProyecto) AS ProyectoCompleto," & vbCrLf

    ' 1) Riesgos identificados (ejecución)
    m_SQL = m_SQL & "    Nz((" & vbCrLf
    m_SQL = m_SQL & "        SELECT Count(*)" & vbCrLf
    m_SQL = m_SQL & "        FROM TbProyectosEdiciones AS E" & vbCrLf
    m_SQL = m_SQL & "        INNER JOIN TbRiesgos AS R ON R.IDEdicion = E.IDEdicion" & vbCrLf
    m_SQL = m_SQL & "        WHERE E.IDProyecto = P.IDProyecto" & vbCrLf
    m_SQL = m_SQL & "          AND R.FechaDetectado Between #" & sIni & "# And #" & sFin & "# " & vbCrLf
    m_SQL = m_SQL & "    ),0) AS RiesgosIdentificados," & vbCrLf

    ' 2) Riesgos retirados (ejecución)
    m_SQL = m_SQL & "    Nz((" & vbCrLf
    m_SQL = m_SQL & "        SELECT Count(*)" & vbCrLf
    m_SQL = m_SQL & "        FROM TbProyectosEdiciones AS E" & vbCrLf
    m_SQL = m_SQL & "        INNER JOIN TbRiesgos AS R ON R.IDEdicion = E.IDEdicion" & vbCrLf
    m_SQL = m_SQL & "        WHERE E.IDProyecto = P.IDProyecto" & vbCrLf
    m_SQL = m_SQL & "          AND R.FechaRetirado Is Not Null" & vbCrLf
    m_SQL = m_SQL & "          AND R.FechaRetirado Between #" & sIni & "# And #" & sFin & "# " & vbCrLf
    m_SQL = m_SQL & "    ),0) AS RiesgosRetirados," & vbCrLf

    ' 3) Riesgos en oferta (oferta)
    m_SQL = m_SQL & "    Nz((" & vbCrLf
    m_SQL = m_SQL & "        SELECT Count(*)" & vbCrLf
    m_SQL = m_SQL & "        FROM TbProyectosEdiciones AS E" & vbCrLf
    m_SQL = m_SQL & "        INNER JOIN TbRiesgosAIntegrar AS RA ON RA.IDEdicion = E.IDEdicion" & vbCrLf
    m_SQL = m_SQL & "        WHERE E.IDProyecto = P.IDProyecto" & vbCrLf
    m_SQL = m_SQL & "          AND RA.FechaDetectado Between #" & sIni & "# And #" & sFin & "# " & vbCrLf
    m_SQL = m_SQL & "    ),0) AS RiesgosEnOferta," & vbCrLf

   ' 4) Riesgos materializados (usar TbRiesgosMaterializaciones: contar EVENTOS)
    m_SQL = m_SQL & "    Nz((" & vbCrLf
    m_SQL = m_SQL & "        SELECT Count(*)" & vbCrLf
    m_SQL = m_SQL & "        FROM TbRiesgosMaterializaciones AS M" & vbCrLf
    m_SQL = m_SQL & "        WHERE M.IDProyecto = P.IDProyecto" & vbCrLf
    m_SQL = m_SQL & "          AND Nz(M.EsMaterializacion,'No')='Sí'" & vbCrLf
    m_SQL = m_SQL & "          AND M.Fecha Between #" & sIni & "# And #" & sFin & "# " & vbCrLf
    m_SQL = m_SQL & "    ),0) AS RiesgosMaterializados," & vbCrLf



   ' 5) Riesgos de oferta que pasan a gestión (oferta -> ejecución): Trasladar='Sí'
    m_SQL = m_SQL & "    Nz((" & vbCrLf
    m_SQL = m_SQL & "        SELECT Count(*)" & vbCrLf
    m_SQL = m_SQL & "        FROM TbProyectosEdiciones AS E" & vbCrLf
    m_SQL = m_SQL & "        INNER JOIN TbRiesgosAIntegrar AS RA ON RA.IDEdicion = E.IDEdicion" & vbCrLf
    m_SQL = m_SQL & "        WHERE E.IDProyecto = P.IDProyecto" & vbCrLf
    m_SQL = m_SQL & "          AND Trim(Nz(RA.Trasladar,''))='Sí'" & vbCrLf
    m_SQL = m_SQL & "          AND RA.FechaAltaRegistro Between #" & sIni & "# And #" & sFin & "# " & vbCrLf
    m_SQL = m_SQL & "    ),0) AS RiesgosOfertaPasanGestion" & vbCrLf


    m_SQL = m_SQL & "FROM TbProyectos AS P" & vbCrLf
    m_SQL = m_SQL & whereProy & vbCrLf
    m_SQL = m_SQL & "ORDER BY P.Proyecto, P.NombreProyecto;"

    Set db = getdb()
    Set rs = db.OpenRecordset(m_SQL, dbOpenSnapshot)

    Set IndicadorRiesgosRepositorioV2_GetTabla = rs
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "IndicadorRiesgosRepositorioV2_GetTabla: " & Err.Number & vbCrLf & Err.Description
    End If
    Set IndicadorRiesgosRepositorioV2_GetTabla = Nothing
End Function


'========================================================
' RESUMEN PARA DASHBOARD (1 fila con 5 números)
'  - Identificados: 1 por CodigoUnico (sin duplicar por ediciones)
'  - Retirados:     1 por CodigoUnico (sin duplicar por ediciones)
'  - En oferta:     1 por IDRiesgoExt (TbRiesgosAIntegrar) (sin duplicar)
'  - Materializados: cuenta TODAS las materializaciones (TbRiesgosMaterializaciones)
'                    SOLO EsMaterializacion='Sí' (pueden repetirse por riesgo)
'  - Oferta->Gestión: 1 por IDRiesgoExt con Trasladar='Sí'
'========================================================
Public Function IndicadorRiesgosV2_GetResumen( _
                                                    ByVal p_dIni As Date, _
                                                    ByVal p_dFin As Date, _
                                                    ByVal p_ListaIDsCsv As String, _
                                                    ByRef p_Error As String _
                                                ) As DAO.Recordset

    Dim db As DAO.Database
    Dim sql As String
    Dim sIni As String, sFin As String

    On Error GoTo errores
    p_Error = ""

    If Len(Trim$(p_ListaIDsCsv & "")) = 0 Then
        p_Error = "Lista de proyectos vacía."
        Err.Raise 1000
    End If

    ' issue-61: sanitizar CSV contra inyeccion SQL antes de concatenar.
    ' Err.Raise 1000 si el CSV contiene valores no enteros.
    Dim sanitizedCsv As String
    sanitizedCsv = ValidarListaIDsCsv(p_ListaIDsCsv)

    ' Access requiere literales de fecha en #mm/dd/yyyy#
    sIni = Format$(p_dIni, "mm\/dd\/yyyy")
    sFin = Format$(p_dFin, "mm\/dd\/yyyy")

    sql = ""
    sql = sql & "SELECT TOP 1 " & vbCrLf

    ' 1) Identificados (UNIQ por CodigoUnico)
    sql = sql & "  Nz((" & vbCrLf
    sql = sql & "    SELECT Count(*)" & vbCrLf
    sql = sql & "    FROM (" & vbCrLf
    sql = sql & "      SELECT R.CodigoUnico" & vbCrLf
    sql = sql & "      FROM TbProyectosEdiciones AS E" & vbCrLf
    sql = sql & "      INNER JOIN TbRiesgos AS R ON R.IDEdicion = E.IDEdicion" & vbCrLf
    sql = sql & "      WHERE E.IDProyecto In (" & sanitizedCsv & ")" & vbCrLf
    sql = sql & "        AND R.CodigoUnico Is Not Null" & vbCrLf
    sql = sql & "        AND R.FechaDetectado Between #" & sIni & "# And #" & sFin & "#" & vbCrLf
    sql = sql & "      GROUP BY R.CodigoUnico" & vbCrLf
    sql = sql & "    ) AS Q" & vbCrLf
    sql = sql & "  ),0) AS RiesgosIdentificados," & vbCrLf

    ' 2) Retirados (UNIQ por CodigoUnico)
    sql = sql & "  Nz((" & vbCrLf
    sql = sql & "    SELECT Count(*)" & vbCrLf
    sql = sql & "    FROM (" & vbCrLf
    sql = sql & "      SELECT R.CodigoUnico" & vbCrLf
    sql = sql & "      FROM TbProyectosEdiciones AS E" & vbCrLf
    sql = sql & "      INNER JOIN TbRiesgos AS R ON R.IDEdicion = E.IDEdicion" & vbCrLf
    sql = sql & "      WHERE E.IDProyecto In (" & sanitizedCsv & ")" & vbCrLf
    sql = sql & "        AND R.CodigoUnico Is Not Null" & vbCrLf
    sql = sql & "        AND R.FechaRetirado Is Not Null" & vbCrLf
    sql = sql & "        AND R.FechaRetirado Between #" & sIni & "# And #" & sFin & "#" & vbCrLf
    sql = sql & "      GROUP BY R.CodigoUnico" & vbCrLf
    sql = sql & "    ) AS Q" & vbCrLf
    sql = sql & "  ),0) AS RiesgosRetirados," & vbCrLf

    ' 3) En oferta (UNIQ por IDRiesgoExt) -> TbRiesgosAIntegrar (join por IDEdicion)
    sql = sql & "  Nz((" & vbCrLf
    sql = sql & "    SELECT Count(*)" & vbCrLf
    sql = sql & "    FROM (" & vbCrLf
    sql = sql & "      SELECT RA.IDRiesgoExt" & vbCrLf
    sql = sql & "      FROM TbProyectosEdiciones AS E" & vbCrLf
    sql = sql & "      INNER JOIN TbRiesgosAIntegrar AS RA ON RA.IDEdicion = E.IDEdicion" & vbCrLf
    sql = sql & "      WHERE E.IDProyecto In (" & sanitizedCsv & ")" & vbCrLf
    sql = sql & "        AND RA.IDRiesgoExt Is Not Null" & vbCrLf
    sql = sql & "        AND RA.FechaAltaRegistro Between #" & sIni & "# And #" & sFin & "#" & vbCrLf
    sql = sql & "      GROUP BY RA.IDRiesgoExt" & vbCrLf
    sql = sql & "    ) AS Q" & vbCrLf
    sql = sql & "  ),0) AS RiesgosEnOferta," & vbCrLf

    ' 4) Materializados (cuenta TODAS las materializaciones) -> TbRiesgosMaterializaciones
    sql = sql & "  Nz((" & vbCrLf
    sql = sql & "    SELECT Count(*)" & vbCrLf
    sql = sql & "    FROM TbRiesgosMaterializaciones AS M" & vbCrLf
    sql = sql & "    WHERE M.IDProyecto In (" & sanitizedCsv & ")" & vbCrLf
    sql = sql & "      AND Nz(M.EsMaterializacion,'')='Sí'" & vbCrLf
    sql = sql & "      AND M.Fecha Between #" & sIni & "# And #" & sFin & "#" & vbCrLf
    sql = sql & "  ),0) AS RiesgosMaterializados," & vbCrLf

    ' 5) Oferta -> Gestión final (UNIQ por IDRiesgoExt) con Trasladar='Sí'
    sql = sql & "  Nz((" & vbCrLf
    sql = sql & "    SELECT Count(*)" & vbCrLf
    sql = sql & "    FROM (" & vbCrLf
    sql = sql & "      SELECT RA.IDRiesgoExt" & vbCrLf
    sql = sql & "      FROM TbProyectosEdiciones AS E" & vbCrLf
    sql = sql & "      INNER JOIN TbRiesgosAIntegrar AS RA ON RA.IDEdicion = E.IDEdicion" & vbCrLf
    sql = sql & "      WHERE E.IDProyecto In (" & sanitizedCsv & ")" & vbCrLf
    sql = sql & "        AND RA.IDRiesgoExt Is Not Null" & vbCrLf
    sql = sql & "        AND Nz(RA.Trasladar,'')='Sí'" & vbCrLf
    sql = sql & "        AND RA.FechaAltaRegistro Between #" & sIni & "# And #" & sFin & "#" & vbCrLf
    sql = sql & "      GROUP BY RA.IDRiesgoExt" & vbCrLf
    sql = sql & "    ) AS Q" & vbCrLf
    sql = sql & "  ),0) AS RiesgosOfertaPasanGestion," & vbCrLf

    ' 6) Vigentes en el periodo (interseca abierto >=1 dia en [dIni, dFin])
    ' issue-127: un riesgo no es distinto a otro por aparecer en otra edicion.
    ' GROUP BY R.CodigoUnico evita que el JOIN con TbProyectosEdiciones multiplique filas.
    sql = sql & "  Nz((" & vbCrLf
    sql = sql & "    SELECT Count(*)" & vbCrLf
    sql = sql & "    FROM (" & vbCrLf
    sql = sql & "      SELECT R.CodigoUnico" & vbCrLf
    sql = sql & "      FROM TbProyectosEdiciones AS E" & vbCrLf
    sql = sql & "      INNER JOIN TbRiesgos AS R ON R.IDEdicion = E.IDEdicion" & vbCrLf
    sql = sql & "      WHERE E.IDProyecto In (" & sanitizedCsv & ")" & vbCrLf
    sql = sql & "        AND R.CodigoUnico Is Not Null" & vbCrLf
    sql = sql & "        AND R.FechaDetectado <= #" & sFin & "#" & vbCrLf
    sql = sql & "        AND (R.FechaRetirado Is Null OR R.FechaRetirado >= #" & sIni & "#)" & vbCrLf
    sql = sql & "      GROUP BY R.CodigoUnico" & vbCrLf
    sql = sql & "    ) AS Q" & vbCrLf
    sql = sql & "  ),0) AS RiesgosVigentesEnPeriodo" & vbCrLf

    ' Tabla “ancla” para que Access no proteste (1 fila)
    sql = sql & "FROM TbProyectos AS X;" & vbCrLf

    Set db = getdb()
    Set IndicadorRiesgosV2_GetResumen = db.OpenRecordset(sql, dbOpenSnapshot)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "IndicadorRiesgosV2_GetResumen: " & Err.Number & vbCrLf & Err.Description
    End If
    Set IndicadorRiesgosV2_GetResumen = Nothing
End Function




Public Function IndicadorRiesgosV2_SqlDetalle( _
                                                ByVal p_Tipo As IndicadorTile, _
                                                ByVal p_dIni As Date, _
                                                ByVal p_dFin As Date, _
                                                ByVal p_ListaIDsCsv As String _
                                            ) As String

    Dim dIni As String, dFin As String
    Dim sanitizedCsv As String

    On Error GoTo errores

    ' issue-61: sanitizar CSV contra inyeccion SQL antes de concatenar.
    ' Err.Raise 1000 si el CSV contiene valores no enteros.
    sanitizedCsv = ValidarListaIDsCsv(p_ListaIDsCsv)

    dIni = Format$(p_dIni, "mm\/dd\/yyyy")
    dFin = Format$(p_dFin, "mm\/dd\/yyyy")

    Select Case p_Tipo

        ' =========================================================
        ' IDENTIFICADOS (1 fila por RIESGO - evita duplicar por edición)
        ' Cuenta/Detalle por CodigoUnico (si se copia entre ediciones)
        ' =========================================================
        Case itIdentificados
            IndicadorRiesgosV2_SqlDetalle = _
                "SELECT " & vbCrLf & _
                "  P.Proyecto  AS CodProyecto," & vbCrLf & _
                "  P.NombreProyecto ," & vbCrLf & _
                "  R.CodigoRiesgo as Código," & vbCrLf & _
                "  Min(R.FechaDetectado) AS Detectado " & vbCrLf & _
                "FROM (TbProyectos AS P" & vbCrLf & _
                "      INNER JOIN TbProyectosEdiciones AS E ON P.IDProyecto = E.IDProyecto)" & vbCrLf & _
                "      INNER JOIN TbRiesgos AS R ON R.IDEdicion = E.IDEdicion" & vbCrLf & _
                "WHERE P.IDProyecto In (" & sanitizedCsv & ")" & vbCrLf & _
                "  AND R.FechaDetectado Between #" & dIni & "# And #" & dFin & "#" & vbCrLf & _
                "GROUP BY P.IDProyecto, P.Proyecto, P.NombreProyecto, R.CodigoRiesgo, R.CodigoUnico" & vbCrLf & _
                "ORDER BY NombreProyecto, Min(R.FechaDetectado), R.CodigoRiesgo;"

        ' =========================================================
        ' RETIRADOS (1 fila por RIESGO - evita duplicar por edición)
        ' =========================================================
        Case itRetirados
            IndicadorRiesgosV2_SqlDetalle = _
                "SELECT " & vbCrLf & _
                "  P.Proyecto  AS CodProyecto," & vbCrLf & _
                "  P.NombreProyecto ," & vbCrLf & _
                "  R.CodigoRiesgo as Código," & vbCrLf & _
                "  Max(R.FechaRetirado) AS Retirado " & vbCrLf & _
                "FROM (TbProyectos AS P" & vbCrLf & _
                "      INNER JOIN TbProyectosEdiciones AS E ON P.IDProyecto = E.IDProyecto)" & vbCrLf & _
                "      INNER JOIN TbRiesgos AS R ON R.IDEdicion = E.IDEdicion" & vbCrLf & _
                "WHERE P.IDProyecto In (" & sanitizedCsv & ")" & vbCrLf & _
                "  AND R.FechaRetirado Is Not Null" & vbCrLf & _
                "  AND R.FechaRetirado Between #" & dIni & "# And #" & dFin & "#" & vbCrLf & _
                "GROUP BY P.IDProyecto, P.Proyecto, P.NombreProyecto, R.CodigoRiesgo, R.CodigoUnico" & vbCrLf & _
                "ORDER BY NombreProyecto, Max(R.FechaRetirado), R.CodigoRiesgo;"

       ' =========================================================
        ' EN OFERTA (TbRiesgosAIntegrar + TbRiesgos)
        ' Campos EXACTOS: NombreProyecto, CodigoRiesgo, FechaDetectado, Trasladar, FechaAltaRegistro
        ' Sin duplicar por ediciones: 1 fila por RA.IDRiesgo (la primera en el rango)
        ' =========================================================
        Case itEnOferta
            IndicadorRiesgosV2_SqlDetalle = _
                "SELECT P.NombreProyecto as Proyecto,P.NombreProyecto, TbRiesgos.CodigoRiesgo as Código, Nz(RA.Trasladar,'') AS Trasladar, RA.FechaAltaRegistro as Alta" & vbCrLf & _
                "FROM ((TbProyectos AS P " & vbCrLf & _
                "  INNER JOIN TbProyectosEdiciones AS E ON P.IDProyecto = E.IDProyecto) " & vbCrLf & _
                "  INNER JOIN TbRiesgosAIntegrar AS RA ON E.IDEdicion = RA.IDEdicion) " & vbCrLf & _
                "  LEFT JOIN TbRiesgos ON RA.IDRiesgo = TbRiesgos.IDRiesgo " & vbCrLf & _
                "WHERE P.IDProyecto In (" & sanitizedCsv & ") " & vbCrLf & _
                "  AND RA.FechaAltaRegistro Between #" & dIni & "# And #" & dFin & "# " & vbCrLf & _
                "ORDER BY P.NombreProyecto, RA.FechaAltaRegistro;"





        ' =========================================================
        ' OFERTA -> PASA A GESTIÓN FINAL (Trasladar='Sí')
        ' =========================================================
        Case itOfertaTrasladar
            IndicadorRiesgosV2_SqlDetalle = _
                "SELECT P.NombreProyecto as Proyecto,P.Proyecto  AS CodProyecto, TbRiesgos.CodigoRiesgo as Código,  Nz(RA.Trasladar,'') AS Trasladar, RA.FechaAltaRegistro as Alta " & vbCrLf & _
                "FROM ((TbProyectos AS P " & vbCrLf & _
                "  INNER JOIN TbProyectosEdiciones AS E ON P.IDProyecto = E.IDProyecto) " & vbCrLf & _
                "  INNER JOIN TbRiesgosAIntegrar AS RA ON E.IDEdicion = RA.IDEdicion) " & vbCrLf & _
                "  INNER JOIN TbRiesgos ON RA.IDRiesgo = TbRiesgos.IDRiesgo " & vbCrLf & _
                "WHERE P.IDProyecto In (" & sanitizedCsv & ") " & vbCrLf & _
                "  AND RA.FechaAltaRegistro Between #" & dIni & "# And #" & dFin & "# " & vbCrLf & _
                "  AND Nz(RA.Trasladar,'')='Sí' " & vbCrLf & _
                "ORDER BY P.NombreProyecto, RA.FechaAltaRegistro;"


        ' =========================================================
        ' MATERIALIZADOS (EVENTOS): 1 fila por materialización (EsMaterializacion='Sí')
        ' Aquí NO agrupamos: Calidad quiere contar repeticiones.
        ' Incluimos Detectado/Retirado del riesgo EN ESA EDICIÓN (para contexto).
        ' =========================================================
        Case itMaterializados
            IndicadorRiesgosV2_SqlDetalle = _
                "SELECT " & vbCrLf & _
                "  P.Proyecto  AS CodProyecto," & vbCrLf & _
                "  P.NombreProyecto ," & vbCrLf & _
                "  M.CodigoRiesgo as Código, " & vbCrLf & _
                "  M.Fecha AS Materialización " & vbCrLf & _
                "FROM (TbProyectos AS P " & vbCrLf & _
                "  INNER JOIN TbRiesgosMaterializaciones AS M ON P.IDProyecto = M.IDProyecto) " & vbCrLf & _
                "  LEFT JOIN TbRiesgos AS R ON (R.IDEdicion = M.IDEdicion AND R.CodigoRiesgo = M.CodigoRiesgo) " & vbCrLf & _
                "WHERE P.IDProyecto In (" & sanitizedCsv & ") " & vbCrLf & _
                "  AND Nz(M.EsMaterializacion,'No')='Sí' " & vbCrLf & _
                "  AND M.Fecha Between #" & dIni & "# And #" & dFin & "# " & vbCrLf & _
                "ORDER BY P.NombreProyecto, M.Fecha, M.CodigoRiesgo;"

        ' =========================================================
        ' VIGENTES EN EL PERIODO (interseca abierto >=1 dia en [dIni, dFin])
        ' 1 fila por TbRiesgos (no agrupamos por CodigoUnico por regla de user)
        ' =========================================================
        Case itVigentesEnPeriodo
            ' issue-127: el mismo riesgo en N ediciones se reporta UNA vez (regla de user).
            ' CodigoUnico es el identificador estable; CodigoRiesgo es por-edicion.
            ' MIN(FechaDetectado) = fecha original; si al menos una edicion sigue vigente
            ' (FechaRetirado Is Null) la columna Retirado queda Null.
            IndicadorRiesgosV2_SqlDetalle = _
                "SELECT " & vbCrLf & _
                "  P.Proyecto  AS CodProyecto," & vbCrLf & _
                "  P.NombreProyecto ," & vbCrLf & _
                "  R.CodigoUnico AS Código," & vbCrLf & _
                "  MIN(R.FechaDetectado) AS Detectado," & vbCrLf & _
                "  IIf(SUM(IIf(R.FechaRetirado Is Null,1,0)) > 0, Null, MAX(R.FechaRetirado)) AS Retirado " & vbCrLf & _
                "FROM (TbProyectos AS P" & vbCrLf & _
                "      INNER JOIN TbProyectosEdiciones AS E ON P.IDProyecto = E.IDProyecto)" & vbCrLf & _
                "      INNER JOIN TbRiesgos AS R ON R.IDEdicion = E.IDEdicion" & vbCrLf & _
                "WHERE P.IDProyecto In (" & sanitizedCsv & ")" & vbCrLf & _
                "  AND R.CodigoUnico Is Not Null" & vbCrLf & _
                "  AND R.FechaDetectado <= #" & dFin & "#" & vbCrLf & _
                "  AND (R.FechaRetirado Is Null OR R.FechaRetirado >= #" & dIni & "#)" & vbCrLf & _
                "GROUP BY P.Proyecto, P.NombreProyecto, R.CodigoUnico" & vbCrLf & _
                "ORDER BY P.NombreProyecto, MIN(R.FechaDetectado), R.CodigoUnico;"


        Case Else
            IndicadorRiesgosV2_SqlDetalle = ""

    End Select
    Exit Function

errores:
    ' issue-61: re-lanzar Err 1000 (validacion de CSV) para que el caller
    ' pueda tratar el caso como input invalido. Otros errores se silencian
    ' (la funcion retorna "" y el caller decide).
    If Err.Number = 1000 Then
        Err.Raise 1000, Err.Source, Err.Description
    End If
    IndicadorRiesgosV2_SqlDetalle = ""
End Function


Public Function FieldExists(ByVal tableName As String, ByVal fieldName As String) As Boolean
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    FieldExists = False
    On Error GoTo salir
    Set tdf = CurrentDb.TableDefs(tableName)
    For Each fld In tdf.Fields
        If StrComp(fld.Name, fieldName, vbTextCompare) = 0 Then
            FieldExists = True
            Exit For
        End If
    Next
    Exit Function
salir:
End Function

Public Function Oferta_KeyField() As String
    ' Devuelve el nombre de campo “clave riesgo” en TbRiesgosAIntegrar
    If FieldExists("TbRiesgosAIntegrar", "CodigoUnico") Then
        Oferta_KeyField = "CodigoUnico"
    ElseIf FieldExists("TbRiesgosAIntegrar", "CodigoRiesgo") Then
        Oferta_KeyField = "CodigoRiesgo"
    ElseIf FieldExists("TbRiesgosAIntegrar", "Riesgo") Then
        Oferta_KeyField = "Riesgo"
    Else
        ' último recurso: no ideal, pero evita romper
        Oferta_KeyField = "IDEdicion"
    End If
End Function

Public Function Oferta_DateField() As String
    ' Devuelve el campo fecha “de alta/detectado” en TbRiesgosAIntegrar
    If FieldExists("TbRiesgosAIntegrar", "FechaAltaRegistro") Then
        Oferta_DateField = "FechaAltaRegistro"
    ElseIf FieldExists("TbRiesgosAIntegrar", "FechaDetectado") Then
        Oferta_DateField = "FechaDetectado"
    Else
        Oferta_DateField = "FechaAltaRegistro" ' fallback
    End If
End Function

Public Function Riesgo_KeyField() As String
    ' Clave estable en TbRiesgos
    If FieldExists("TbRiesgos", "CodigoUnico") Then
        Riesgo_KeyField = "CodigoUnico"
    Else
        Riesgo_KeyField = "CodigoRiesgo"
    End If
End Function








