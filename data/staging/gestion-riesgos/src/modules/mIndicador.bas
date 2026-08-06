Attribute VB_Name = "mIndicador"
Option Compare Database
Option Explicit
Public ColProyectosAbiertos As New Collection
Public ColProyectosParaInforme As New Collection
Public colRiesgosDetalle As New Collection
Public colIndicadorSemestre1ParaAño As New Collection, colIndicadorSemestre2ParaAño As New Collection, colIndicadorResumenAnual As New Collection
Public colSemestre1 As New Collection, colSemestre2 As New Collection
Public Const dblValorAlarma As Double = 15
Public Const dblValorObjetivo As Double = 20
Private flag As String
Private m_SQL As String
Private dato
Public strTextoEnlbl As String
Public Function RellenaColeccionProyectosAbiertos( _
                                                    strAño As String, _
                                                    Optional ByVal strSemestre As String _
                                                    ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 23/04/2020
    '   -Modificaciones:
   
    '   -Funcionamiento:
   
    '   -llamada por:
    
    '   -Devuelve:
    '       colProyectosAbiertos.Add strIDProyecto & "|" & strProyecto & "|" & _
                                    strNombreProyecto & "|" & strFechaCierre & _
                                    "|" & strFechaRegistroInicial
    '       RellenaColeccionProyectosAbiertos = Descriptivo
    '       RellenaColeccionProyectosAbiertos = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strIDProyecto As String, strProyecto As String, strNombreProyecto As String, strFechaCierre As String, strFechaRegistroInicial As String, _
         strEstaEnintervalo As String, strProyectoAExcluir As String, strFechaInicialInforme As String, strFechaFinalInforme As String, intNumero As Integer, intNumeroTotal As Integer, _
         strTextoError As String
    On Error GoTo errores
    
    Set ColProyectosAbiertos = New Collection
    strAño = Year(Now())
    If Not IsNumeric(strAño) Then
        strTextoError = "El año es obligatorio"
        Err.Raise 1000
    End If
    If strSemestre <> "" Then
        If strSemestre <> "1" And strSemestre <> "2" Then
            strTextoError = "Sólo hay dos semestres en un año"
            Err.Raise 1000
        End If
        If strSemestre = "1" Then
            strFechaInicialInforme = "01/01/" & Format(strAño, "0000")
            strFechaFinalInforme = "30/06/" & Format(strAño, "0000")
        Else
            strFechaInicialInforme = "01/07/" & Format(strAño, "0000")
            strFechaFinalInforme = Format("31/12/" & strAño)
        End If
    Else
        strFechaInicialInforme = "01/01/" & Format(strAño, "0000")
        strFechaFinalInforme = Format("31/12/" & strAño)
    End If
    '----------------------------------------------------------------------------------
    ' Expedientes adjudicados
    '----------------------------------------------------------------------------------
    m_SQL = "SELECT IDProyecto, Proyecto, NombreProyecto, FechaCierre,FechaRegistroInicial " & _
            "FROM TbProyectos " & _
            "ORDER BY TbProyectos.IDProyecto;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveLast
            .MoveFirst
            intNumero = 1
            intNumeroTotal = CInt(.RecordCount)
            Do While Not .EOF
                Avance "Revisando Proyectos...(" & intNumero & " de " & intNumeroTotal & ")"
                
                strProyecto = Nz(.fields("Proyecto"), "")
                strIDProyecto = Nz(.fields("IDProyecto"), "")
                strNombreProyecto = Nz(.fields("NombreProyecto"), "")
                strFechaCierre = Nz(.fields("FechaCierre"), "")
                strFechaRegistroInicial = Nz(.fields("FechaRegistroInicial"), "")
                '-------------------------------------------------------------------
                '   -Devuelve:
                '       EstaEnElIntervaloDadoRiesgo = strEstaEnintervalo
                '       EstaEnElIntervaloDadoRiesgo = "#ERR" & "|" & strTextoError
                '-------------------------------------------------------------------
                flag = EstaEnElIntervaloDadoRiesgo(strFechaInicialInforme, strFechaFinalInforme, strFechaRegistroInicial, strFechaCierre)
                If InStr(1, flag, "|") <> 0 Then
                    dato = Split(flag, "|")
                    strTextoError = "El método EstaEnElIntervaloDadoRiesgo ha devuelto el error: " & vbNewLine & dato(1)
                    Err.Raise 1000
                End If
                strEstaEnintervalo = flag
                If strEstaEnintervalo = "Sí" Then
                    ColProyectosAbiertos.Add strIDProyecto & "|" & strProyecto & "|" & strNombreProyecto & "|" & strFechaCierre & "|" & strFechaRegistroInicial
                End If
siguiente:
                intNumero = intNumero + 1
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    RellenaColeccionProyectosAbiertos = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método RellenaColeccionProyectosAbiertos ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    RellenaColeccionProyectosAbiertos = "#ERR" & "|" & strTextoError
End Function
Public Function RellenaColeccionDetalleRiesgos() As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 23/04/2020
    '   -Modificaciones:
   
    '   -Funcionamiento:
   
    '   -llamada por:
    
    '   -Devuelve:
    '   colProyectosAbiertos.Add strIDProyecto & "|" & strProyecto & "|" & _
                                    strNombreProyecto & "|" & strFechaCierre & _
                                    "|" & strFechaRegistroInicial
    '       -colRiesgosDetalle.Add strProyecto & " " & strNombreProyecto & "|" & CadenaRiesgos
    '       CadenaRiesgos="R01:01/01/2019:15/06/2019;R02:
    '       RellenaColeccionDetalleRiesgos = Descriptivo
    '       RellenaColeccionDetalleRiesgos = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, rcdDestino As DAO.Recordset, strIDProyecto As String, strNombreProyecto As String, _
        strFechaDetectado As String, strFechaMaterializado As String, strProyecto As String, strRiesgo As String, _
        strResultado As String, strCadenaRiesgos As String, strProyectoCompleto As String, dato1 As Variant, _
        intNumero As Integer, intNumeroTotal As Integer, _
        strTextoError As String
    On Error GoTo errores
    
    m_SQL = "TbAuxProyectosRiesgos"
    Set rcdDestino = getdb().OpenRecordset(m_SQL)
    With rcdDestino
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                .Delete
                .MoveNext
            Loop
        End If
    End With
'    rcdDestino.Close
'    Set rcdDestino = Nothing
    If ColProyectosParaInforme.count > 0 Then
        intNumeroTotal = CInt(ColProyectosParaInforme.count)
        intNumero = 1
        For Each varItem In ColProyectosParaInforme
            Avance "Rellenando detalle de Proyectos...(" & intNumero & " de " & intNumeroTotal & ")"
            
            strCadenaRiesgos = ""
            dato = Split(varItem, "|")
            strIDProyecto = dato(0)
            strProyecto = dato(1)
            strNombreProyecto = dato(2)
            strProyectoCompleto = strProyecto & " " & strNombreProyecto
            m_SQL = "SELECT DISTINCT TbRiesgos.CodigoRiesgo " & _
                    "FROM TbRiesgos INNER JOIN TbProyectosEdiciones ON TbRiesgos.IDEdicion = TbProyectosEdiciones.IDEdicion " & _
                    "WHERE (((TbProyectosEdiciones.IDProyecto)=" & strIDProyecto & "));"
            Set rcdDatos = getdb().OpenRecordset(m_SQL)
            With rcdDatos
                If Not .EOF Then
                    .MoveFirst
                    Do While Not .EOF
                        strRiesgo = Nz(.fields("CodigoRiesgo"), "")
                        '-------------------------------------------------------------------
                        '   -Devuelve:
                        '       strResultado=strFechaDetectado & ":" & strFechaMaterializado
                        '       RellenaDatosRiesgo = strResultado
                        '       RellenaDatosRiesgo = "#ERR" & "|" & strTextoError
                        '-------------------------------------------------------------------
                        flag = RellenaDatosRiesgo(strIDProyecto, strRiesgo)
                        If InStr(1, flag, "|") <> 0 Then
                            dato = Split(flag, "|")
                            strTextoError = "El método RellenaDatosRiesgo ha devuelto el error: " & vbNewLine & dato(1)
                            Err.Raise 1000
                        End If
                        strResultado = flag
                        dato1 = Split(strResultado, ":")
                        strFechaDetectado = dato1(0)
                        strFechaMaterializado = dato1(1)
                        rcdDestino.AddNew
                            rcdDestino.fields("IDProyecto") = strIDProyecto
                            rcdDestino.fields("NombreProyectoCompleto") = strProyectoCompleto
                            rcdDestino.fields("Riesgo") = strRiesgo
                            If IsDate(strFechaDetectado) Then
                                rcdDestino.fields("FechaDetectado") = strFechaDetectado
                            End If
                            If IsDate(strFechaMaterializado) Then
                                rcdDestino.fields("FechaMaterializado") = strFechaMaterializado
                            End If
                        rcdDestino.Update
                         If strCadenaRiesgos = "" Then
                            strCadenaRiesgos = strRiesgo & ":" & strResultado
                        Else
                            strCadenaRiesgos = strCadenaRiesgos & ";" & strRiesgo & ":" & strResultado
                        End If
                        
                        .MoveNext
                    Loop
                End If
            End With
            rcdDatos.Close
            Set rcdDatos = Nothing
            colRiesgosDetalle.Add strProyectoCompleto & "|" & strCadenaRiesgos
            intNumero = intNumero + 1
        Next
    End If
    
    RellenaColeccionDetalleRiesgos = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método RellenaColeccionDetalleRiesgos ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    RellenaColeccionDetalleRiesgos = "#ERR" & "|" & strTextoError
End Function
Public Function RellenaDatosRiesgo(strIDProyecto As String, strCodRiesgo As String) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 23/04/2020
    '   -Modificaciones:
   
    '   -Funcionamiento:
   
    '   -llamada por:
    
    '   -Devuelve:
    '       strResultado=strFechaDetectado & ":" & strFechaMaterializado
    '       RellenaDatosRiesgo = strResultado
    '       RellenaDatosRiesgo = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strFechaDetectado As String, strFechaMaterializado As String, strResultado As String, strTextoError As String
    On Error GoTo errores
    m_SQL = "SELECT TbRiesgos.FechaDetectado " & _
            "FROM TbRiesgos INNER JOIN TbProyectosEdiciones ON TbRiesgos.IDEdicion = TbProyectosEdiciones.IDEdicion " & _
            "WHERE (((TbRiesgos.CodigoRiesgo)='" & strCodRiesgo & "') AND ((TbProyectosEdiciones.IDProyecto)=" & strIDProyecto & ")) " & _
            "ORDER BY TbRiesgos.FechaDetectado;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            strFechaDetectado = Nz(.fields("FechaDetectado"), "")
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    m_SQL = "SELECT TbRiesgos.FechaMaterializado " & _
            "FROM TbRiesgos INNER JOIN TbProyectosEdiciones ON TbRiesgos.IDEdicion = TbProyectosEdiciones.IDEdicion " & _
            "WHERE (((TbRiesgos.CodigoRiesgo)='" & strCodRiesgo & "') AND ((TbProyectosEdiciones.IDProyecto)=" & strIDProyecto & ") AND (Not (TbRiesgos.FechaMaterializado) Is Null)) " & _
            "ORDER BY TbRiesgos.FechaMaterializado DESC;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            strFechaMaterializado = Nz(.fields("FechaMaterializado"), "")
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
'    getdb().Close
'    Set m_ObjEntorno.dbs = Nothing
    strResultado = strFechaDetectado & ":" & strFechaMaterializado
    RellenaDatosRiesgo = strResultado
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método RellenaDatosRiesgo ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    RellenaDatosRiesgo = "#ERR" & "|" & strTextoError
End Function
Public Function EstaEnElIntervaloDadoRiesgo( _
                                        strFechaInicialIntervalo As String, _
                                        strFechaFinalIntervalo As String, _
                                        Optional strFechaInicial As String, _
                                        Optional strFechaFinal) As String
    '--------------------------------------------------------
    ' Variante específica de mIndicador para procesar riesgos.
    ' Renombrada desde EstaEnElIntervaloDado (issue #59) para evitar
    ' colisión con Funciones Generales.EstaEnElIntervaloDado (que
    ' devuelve EnumSiNo). Esta variante devuelve String con codificación
    ' de error "#ERR|<mensaje>" para preservar el contrato histórico
    ' de mIndicador.bas:87-94.
    '--------------------------------------------------------
    Dim strEstaEnintervalo As String, strTextoError As String
    On Error GoTo errores
    If Not IsDate(strFechaInicialIntervalo) Then
        strTextoError = "Se ha de indiar la fecha inicial del intervalo"
        Err.Raise 1000
    End If
    If Not IsDate(strFechaFinalIntervalo) Then
        strTextoError = "Se ha de indiar la fecha final del intervalo"
        Err.Raise 1000
    End If
    If CDate(strFechaFinalIntervalo) < CDate(strFechaInicialIntervalo) Then
        strTextoError = "Se ha de indiar la fecha final del intervalo posterior o igual a la inicial"
        Err.Raise 1000
    End If
    If IsDate(strFechaInicial) And IsDate(strFechaFinal) Then
        If CDate(strFechaFinal) < CDate(strFechaInicial) Then
            strTextoError = "Las fechas iniciales y finales de expediente no son obligatorias, pero de rellenarse la final ha de ser posterior a la inicial."
            Err.Raise 1000
        End If
    End If
    If IsDate(strFechaInicial) And IsDate(strFechaFinal) Then
        If CDate(strFechaInicial) < CDate(strFechaInicialIntervalo) And CDate(strFechaFinal) < CDate(strFechaInicialIntervalo) Then
            '---------------------
            ' cualquiera de las fechas de inicio y fin de expediente son anteriores de la fecha inicial y final del intervalo
            '---------------------
            strEstaEnintervalo = "No"
        ElseIf (CDate(strFechaInicial) >= CDate(strFechaInicialIntervalo) And CDate(strFechaInicial) <= CDate(strFechaFinalIntervalo)) Or _
                (CDate(strFechaFinal) >= CDate(strFechaInicialIntervalo) And CDate(strFechaFinal) <= CDate(strFechaFinalIntervalo)) Then
            '---------------------
            ' cualquiera de las fechas de inicio y fin de expediente está en el intervalo entre fecha inicial intervalo y final intervalo
            '---------------------
            strEstaEnintervalo = "Sí"
        ElseIf CDate(strFechaInicial) > CDate(strFechaFinalIntervalo) Then
            '---------------------
            ' cualquiera de las fechas de inicio y fin de expediente son posteriores de la fecha inicial y final del intervalo
            '---------------------
            strEstaEnintervalo = "No"
        ElseIf CDate(strFechaInicial) < CDate(strFechaInicialIntervalo) And CDate(strFechaFinal) > CDate(strFechaFinalIntervalo) Then
            '---------------------
            ' El expediente empieza antes del inicio del intervalo y acaba después del fin del intervalo
            '---------------------
            strEstaEnintervalo = "Sí"
        Else

            strTextoError = "Situación de fechas desconocida"
            Err.Raise 1000
        End If
    ElseIf IsDate(strFechaInicial) And Not IsDate(strFechaFinal) Then
        strEstaEnintervalo = "Sí"
    ElseIf Not IsDate(strFechaInicial) And IsDate(strFechaFinal) Then
        '---------------------
        ' la fecha final de expediente ha de estar entre la de inicio y final de intervalo
        '---------------------
        If CDate(strFechaFinal) >= CDate(strFechaInicialIntervalo) And CDate(strFechaFinal) <= CDate(strFechaFinalIntervalo) Then
            strEstaEnintervalo = "Sí"
        Else
            strEstaEnintervalo = "No"
        End If
    Else
        '---------------------
        ' Expediente sin fecha de firmadecontrato y sin fecharecepción adjudicado, cualquier valor de fechas de intervalo entra
        '---------------------
        strEstaEnintervalo = "Sí"
    End If

    EstaEnElIntervaloDadoRiesgo = strEstaEnintervalo
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método EstaEnElIntervaloDadoRiesgo ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.description
    End If
    EstaEnElIntervaloDadoRiesgo = "#ERR" & "|" & strTextoError
End Function
Private Function EstaEnCadenaProyectoAExcluir(strCadenaElementos, strElemento As String) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día X
    '   -Modificaciones:
    
    '   -Funcionamiento:
   
    '   -Llamada desde
   
    '   -Devuelve:
    '       EstaEnCadenaProyectoAExcluir = Sí/No
    '       EstaEnCadenaProyectoAExcluir = "ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim dato1 As Variant, strElementoDeCadena As String, strTextoError As String
    On Error GoTo errores
    If strCadenaElementos <> "" And strElemento <> "" Then
        If InStr(1, strCadenaElementos, ";") <> 0 Then
            dato1 = Split(strCadenaElementos, ";")
            For Each varItem In dato1
                strElementoDeCadena = CStr(varItem)
                If InStr(1, strElemento, strElementoDeCadena) <> 0 Then
                    EstaEnCadenaProyectoAExcluir = "Sí"
                    Exit Function
                End If
            Next
        Else
            If InStr(1, strElemento, strCadenaElementos) <> 0 Then
                EstaEnCadenaProyectoAExcluir = "Sí"
                Exit Function
            End If
        End If
    End If
    EstaEnCadenaProyectoAExcluir = "No"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método EstaEnCadenaProyectoAExcluir ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.description
    End If
    EstaEnCadenaProyectoAExcluir = "ERR" & "|" & strTextoError
End Function
Public Function DameNumRiesgosIdentificados( _
                                                strIDProyecto As String, _
                                                strAño As String, _
                                                Optional strSemestre As String _
                                                ) As String

                                                
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día X
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -Se cuentan el número de Riesgos que hay en el momento de la consulta
    '   -Llamada desde
   
    '   -Devuelve:
    '       DameNumRiesgosIdentificados =cstr(intNumRiesgosIdentificados)
    '       DameNumRiesgosIdentificados = "ERR" & "|" & strTextoError
    '--------------------------------------------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, intNumRiesgosIdentificados As Integer, strFechaInicialConsulta As String, strFechaFinalConsulta As String, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDProyecto) Then
        strTextoError = "Se ha de indicar el IDProyecto"
        Err.Raise 1000
    End If
    If Not IsNumeric(strAño) Then
        strTextoError = "Se ha de indicar el año"
        Err.Raise 1000
    End If
    If strSemestre = "1" Or strSemestre = "2" Then
        If strSemestre = "1" Then
            strFechaInicialConsulta = "01/01/" & strAño
            strFechaFinalConsulta = Format("30/06/" & strAño, "mm/dd/yyyy")
        Else
            strFechaInicialConsulta = Format("01/07/" & strAño, "mm/dd/yyyy")
            strFechaFinalConsulta = Format("31/12/" & strAño, "mm/dd/yyyy")
        End If
    Else
        strFechaInicialConsulta = "01/01/" & strAño
        strFechaFinalConsulta = Format("31/12/" & strAño, "mm/dd/yyyy")
    End If
    m_SQL = "SELECT DISTINCT TbRiesgos.CodigoUnico " & _
            "FROM TbRiesgos INNER JOIN TbProyectosEdiciones ON TbRiesgos.IDEdicion = TbProyectosEdiciones.IDEdicion " & _
            "WHERE (((TbProyectosEdiciones.IDProyecto)=" & strIDProyecto & ") AND ((TbRiesgos.FechaDetectado) Between #" & strFechaInicialConsulta & "# And #" & strFechaFinalConsulta & "#) );"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveLast
            .MoveFirst
            intNumRiesgosIdentificados = .RecordCount
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    DameNumRiesgosIdentificados = CStr(intNumRiesgosIdentificados)
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameNumRiesgosIdentificados ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DameNumRiesgosIdentificados = "#ERR" & "|" & strTextoError
End Function
Public Function DameNumRiesgosMaterializados( _
                                                strIDProyecto As String, _
                                                strAño As String, _
                                                Optional strSemestre As String _
                                                ) As String
                                                
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día X
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -Se cuentan el número de Riesgos que hay en el momento de la consulta
    '   -Llamada desde
   
    '   -Devuelve:
    '       DameNumRiesgosMaterializados =cstr(intNumRiesgosMaterializados)
    '       DameNumRiesgosMaterializados = "ERR" & "|" & strTextoError
    '--------------------------------------------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, intNumRiesgosMaterializados As Integer, strFechaInicialConsulta As String, strFechaFinalConsulta As String, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDProyecto) Then
        strTextoError = "Se ha de indicar el IDProyecto"
        Err.Raise 1000
    End If
    If Not IsNumeric(strAño) Then
        strTextoError = "El año es obligatorio"
        Err.Raise 1000
    End If
    If strSemestre = "1" Or strSemestre = "2" Then
        If strSemestre = "1" Then
            strFechaInicialConsulta = "01/01/" & strAño
            strFechaFinalConsulta = Format("30/06/" & strAño, "mm/dd/yyyy")
        Else
            strFechaInicialConsulta = Format("01/07/" & strAño, "mm/dd/yyyy")
            strFechaFinalConsulta = Format("31/12/" & strAño, "mm/dd/yyyy")
        End If
    Else
        strFechaInicialConsulta = "01/01/" & strAño
        strFechaFinalConsulta = Format("31/12/" & strAño, "mm/dd/yyyy")
    End If
    m_SQL = "SELECT DISTINCT TbRiesgos.CodigoUnico " & _
            "FROM TbRiesgos INNER JOIN TbProyectosEdiciones ON TbRiesgos.IDEdicion = TbProyectosEdiciones.IDEdicion " & _
            "WHERE (((TbProyectosEdiciones.IDProyecto)=" & strIDProyecto & ") AND ((TbRiesgos.FechaMaterializado) Between #" & strFechaInicialConsulta & "# And #" & strFechaFinalConsulta & "#) );"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveLast
            .MoveFirst
            intNumRiesgosMaterializados = .RecordCount
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    DameNumRiesgosMaterializados = CStr(intNumRiesgosMaterializados)
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameNumRiesgosMaterializados ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DameNumRiesgosMaterializados = "#ERR" & "|" & strTextoError
End Function
Private Function DameDatosRiesgosParaIndicador( _
                                                strIDProyecto As String, _
                                                strAño As String, _
                                                Optional strSemestre As String _
                                                ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día X
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -Se cuentan el número de Riesgos que hay en el momento de la consulta
    '   -Llamada desde
   
    '   -Devuelve:
    '       DameDatosRiesgosParaIndicador =cstr(intNumRiesgosIdentificados)  & ";" & _
                                            cstr(intNumRiesgosMaterializados) & ";" & cstr(dblIndicador)
    '       DameDatosRiesgosParaIndicador = "ERR" & "|" & strTextoError
    '--------------------------------------------------------------------------------------------------------
    Dim intNumRiesgosIdentificados As Integer, intNumRiesgosMaterializados As Integer, dblIndicador As Double, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDProyecto) Then
        strTextoError = "Se ha de indicar el IDProyecto"
        Err.Raise 1000
    End If
    If Not IsNumeric(strAño) Then
        strTextoError = "Se ha de indicar el año"
        Err.r1000
    End If
    '------------------------------------------------------------------------------
    '   -Devuelve:
    '       DameNumRiesgosIdentificados =cstr(intNumRiesgosIdentificados)
    '       DameNumRiesgosIdentificados = "ERR" & "|" & strTextoError
    '------------------------------------------------------------------------------
    flag = DameNumRiesgosIdentificados(strIDProyecto, strAño, strSemestre)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameDatosRiesgosParaIndicador ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If Not IsNumeric(flag) Then
        strTextoError = "El método DameDatosRiesgosParaIndicador ha devuelto un resultado con formato incorrecto"
        Err.Raise 1000
    End If
    intNumRiesgosIdentificados = CInt(flag)
    '-----------------------------------------------------------------------------------
     '   -Devuelve:
    '       DameNumRiesgosMaterializados =cstr(intNumRiesgosMaterializados)
    '       DameNumRiesgosMaterializados = "ERR" & "|" & strTextoError
    '-----------------------------------------------------------------------------------
    flag = DameNumRiesgosMaterializados(strIDProyecto, strAño, strSemestre)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameNumRiesgosMaterializados ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If Not IsNumeric(flag) Then
        strTextoError = "El método DameNumRiesgosMaterializados ha devuelto un resultado con formato incorrecto"
        Err.Raise 1000
    End If
    intNumRiesgosMaterializados = CInt(flag)
    If intNumRiesgosIdentificados > 0 Then
        dblIndicador = (intNumRiesgosMaterializados / intNumRiesgosIdentificados) * 100
        dblIndicador = ValorRedondeado(dblIndicador, 2)
    End If
    DameDatosRiesgosParaIndicador = CStr(intNumRiesgosIdentificados) & ";" & CStr(intNumRiesgosMaterializados) & ";" & CStr(dblIndicador)
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameDatosRiesgosParaIndicador ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.description
    End If
    DameDatosRiesgosParaIndicador = "#ERR" & "|" & strTextoError
End Function
Public Function RellenaColeccionIndicador( _
                                            strAño As String, _
                                            Optional strSemestre As String _
                                            ) As String

    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 23/04/2020
    '   -Modificaciones:
    
    '   -Funcionamiento:
    
    '   -Llamada desde
   
    '   -Devuelve:
    '   colIndicador.Add strIDProyecto & "|" & strProyecto & " " & _
                            strNombreProyecto & "|" & intNumRiesgosIdentificados & _
                            "|" & intNumRiesgosMaterializados & "|" & dblIndicador
    '   colIndicador.Add  strProyecto & " " & _
                            strNombreProyecto & "|" & intNumRiesgosIdentificados & _
                            "|" & intNumRiesgosMaterializados & "|" & dblIndicador
    '       RellenaColeccionIndicador =Descriptivo
    '       RellenaColeccionIndicador = "ERR" & "|" & strTextoError
    '--------------------------------------------------------------------------------------------------------
    Dim intNumRiesgosIdentificados As Integer, intNumRiesgosMaterializados As Integer, dblIndicador As Double, strIDProyecto As String, strProyecto As String, _
        strNombreProyecto As String, dato1 As Variant, strResultadoSemestre1 As String, strResultadoSemestre2 As String, strResultadoResumenAnual As String, _
        intIDentificadosSemestre1 As Integer, intIDentificadosSemestre2, intMaterializadosSemestre1 As Integer, intMaterializadosSemestre2 As Integer, dblIndicadorSemestre1 As Double, _
        dblIndicadorSemestre2 As Double, intIDentificadosResumenAnual As Integer, intMaterializadosResumenAnual As Integer, dblIndicadorResumenAnual As Double, intNumero As Integer, _
        intNumeroTotal As Integer, strTextoError As String
    On Error GoTo errores
    
    If Not IsNumeric(strAño) Then
        strTextoError = "Se ha de indicar el año"
        Err.Raise 1000
    End If
    If strSemestre = "1" Or strSemestre = "2" Then
        If ColProyectosParaInforme.count > 0 Then
            intNumeroTotal = CInt(ColProyectosParaInforme.count)
            intNumero = 1
            For Each varItem In ColProyectosParaInforme
                dato = Split(varItem, "|")
                strIDProyecto = dato(0)
                strProyecto = dato(1)
                strNombreProyecto = dato(2)
                Avance "Indicadores de Proyectos...(" & intNumero & " de " & intNumeroTotal & ")"
                
                '--------------------------------------------------------------------------------
                '   -Devuelve:
                '       DameNumeroRiesgosDetectadosMaterializados =intIDentificados & ";" &  intMaterializados
                '       DameNumeroRiesgosDetectadosMaterializados = "ERR" & "|" & strTextoError
                '--------------------------------------------------------------------------------
                flag = DameNumeroRiesgosDetectadosMaterializados(strAño, strSemestre, strIDProyecto)
                If InStr(1, flag, "|") <> 0 Then
                    dato1 = Split(flag, "|")
                    strTextoError = "El método DameNumeroRiesgosDetectadosMaterializados ha devuelto el error: " & vbNewLine & dato1(1)
                    Err.Raise 1000
                End If
                dato1 = Split(flag, ";")
                intNumRiesgosIdentificados = dato1(0)
                intNumRiesgosMaterializados = dato1(1)
                If intNumRiesgosIdentificados = 0 Then
                    dblIndicador = 0
                Else
                    dblIndicador = (intNumRiesgosMaterializados / intNumRiesgosIdentificados) * 100
                End If
                If strSemestre = "1" Then
                    colSemestre1.Add strIDProyecto & "|" & strProyecto & " " & strNombreProyecto & "|" & intNumRiesgosIdentificados & "|" & intNumRiesgosMaterializados & "|" & dblIndicador
                Else
                    colSemestre2.Add strIDProyecto & "|" & strProyecto & " " & strNombreProyecto & "|" & intNumRiesgosIdentificados & "|" & intNumRiesgosMaterializados & "|" & dblIndicador
                End If
                intNumero = intNumero + 1
            Next
        End If
    Else
        
        '--------------------------------------------------------------------------------
        '   -Devuelve:
        '       strResultado=intIDentificados & ";" &  intMaterializados-->strIDProyecto<>""
        '       strResultado=intIDentificadosSemestre1 & ";" & intMaterializadosSemestre1 & "#" & _
                            intIDentificadosSemestre2 & ";" & intMaterializadosSemestre2 & "#" & _
                            intIDentificadosResumenAnual & ";" & intMaterializadosResumenAnual -->strIDProyecto=""
        '       DameNumeroRiesgosDetectadosMaterializados =
        '       DameNumeroRiesgosDetectadosMaterializados = "ERR" & "|" & strTextoError
        '--------------------------------------------------------------------------------
        flag = DameNumeroRiesgosDetectadosMaterializados(strAño)
        If InStr(1, flag, "|") <> 0 Then
            dato1 = Split(flag, "|")
            strTextoError = "El método DameNumeroRiesgosDetectadosMaterializados ha devuelto el error: " & vbNewLine & dato1(1)
            Err.Raise 1000
        End If
        dato = Split(flag, "#")
        strResultadoSemestre1 = dato(0)
        strResultadoSemestre2 = dato(1)
        strResultadoResumenAnual = dato(2)
        dato1 = Split(strResultadoSemestre1, ";")
        intIDentificadosSemestre1 = dato1(0)
        intMaterializadosSemestre1 = dato1(1)
        dato1 = Split(strResultadoSemestre2, ";")
        intIDentificadosSemestre2 = dato1(0)
        intMaterializadosSemestre2 = dato1(1)
        dato1 = Split(strResultadoResumenAnual, ";")
        intIDentificadosResumenAnual = dato1(0)
        intMaterializadosResumenAnual = dato1(1)
        If intIDentificadosSemestre1 = 0 Then
            dblIndicadorSemestre1 = 0
        Else
            dblIndicadorSemestre1 = (intMaterializadosSemestre1 / intIDentificadosSemestre1) * 100
        End If
        colIndicadorSemestre1ParaAño.Add intIDentificadosSemestre1 & "|" & intMaterializadosSemestre1 & "|" & dblIndicadorSemestre1
        If intIDentificadosSemestre2 = 0 Then
            dblIndicadorSemestre2 = 0
        Else
            dblIndicadorSemestre2 = (intMaterializadosSemestre2 / intIDentificadosSemestre2) * 100
        End If
        colIndicadorSemestre2ParaAño.Add intIDentificadosSemestre2 & "|" & intMaterializadosSemestre2 & "|" & dblIndicadorSemestre2
        If intIDentificadosResumenAnual = 0 Then
            dblIndicadorResumenAnual = 0
        Else
            dblIndicadorResumenAnual = (intMaterializadosResumenAnual / intIDentificadosResumenAnual) * 100
        End If
        colIndicadorResumenAnual.Add intIDentificadosResumenAnual & "|" & intMaterializadosResumenAnual & "|" & dblIndicadorResumenAnual
    End If
    RellenaColeccionIndicador = "Relleno"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método RellenaColeccionIndicador ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.description
    End If
    RellenaColeccionIndicador = "#ERR" & "|" & strTextoError
End Function
Public Function DameHTMLSemestre( _
                                    strSemestre As String, _
                                    strAño As String _
                                    ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 23/04/2020
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -colIndicador.Add strIDProyecto & "|" & strProyecto & " " & strNombreProyecto & "|" & intNumRiesgosIdentificados & "|" & intNumRiesgosMaterializados & "|" & dblIndicador
    '   -Llamada desde
   
    '   -Devuelve:
    '       DameHTMLSemestre =strMensaje
    '       DameHTMLSemestre = "ERR" & "||" & strTextoError
    '--------------------------------------------------------------------------------------------------------
    Dim colIndicadorSemestre As New Collection, intNumRiesgosIdentificados As Integer, intNumRiesgosMaterializados As Integer, dblIndicador As Double, strComentarioIndicador As String, strMensaje As String, _
        strProyectoCompleto As String, intNumRiesgosIdentificadosTotales As Integer, intNumRiesgosMaterializadosTotales As Integer, strTextoSemestre As String, strTextoError As String
    On Error GoTo errores
    Set colIndicadorSemestre = New Collection
    If strSemestre = "1" Then
        strTextoSemestre = "SEMESTRE 1 " & strAño
        Set colIndicadorSemestre = colSemestre1
    Else
        strTextoSemestre = "SEMESTRE 2 " & strAño
        Set colIndicadorSemestre = colSemestre2
    End If
    strMensaje = "<table>" & vbNewLine
       strMensaje = strMensaje & "<tr>" & vbNewLine
           strMensaje = strMensaje & "<td colspan='7' class=""ColespanArribaIndicadorRiesgos""> Indicador: Porcentaje de riesgos materializados sobre el total de riesgos detectados de cada proyecto</td>"
       strMensaje = strMensaje & "</tr>" & vbNewLine
       strMensaje = strMensaje & "<tr>" & vbNewLine
           strMensaje = strMensaje & "<td class=""CabeceraSuperiorIndicadorRiesgos"">Proyecto</td>" & vbNewLine
           strMensaje = strMensaje & "<td class=""CabeceraSuperiorIndicadorRiesgos"">Riesgos Identificados</td>" & vbNewLine
           strMensaje = strMensaje & "<td class=""CabeceraSuperiorIndicadorRiesgos"">Riesgos identificados materializados</td>" & vbNewLine
           strMensaje = strMensaje & "<td class=""CabeceraSuperiorIndicadorRiesgos"">Resultado</td>" & vbNewLine
           strMensaje = strMensaje & "<td class=""CabeceraSuperiorIndicadorRiesgos"">Valor Objetivo</td>" & vbNewLine
           strMensaje = strMensaje & "<td class=""CabeceraSuperiorIndicadorRiesgos"">Valor de Alarma</td>" & vbNewLine
           strMensaje = strMensaje & "<td class=""CabeceraSuperiorIndicadorRiesgos"">¿Dentro de los límiteSí</td>" & vbNewLine
       strMensaje = strMensaje & "</tr>" & vbNewLine
    If colIndicadorSemestre.count > 0 Then
        For Each varItem In colIndicadorSemestre
            dato = Split(varItem, "|")
            strProyectoCompleto = dato(1)
            intNumRiesgosIdentificados = dato(2)
            intNumRiesgosMaterializados = dato(3)
            dblIndicador = dato(4)
            strComentarioIndicador = ""
            intNumRiesgosIdentificadosTotales = intNumRiesgosIdentificadosTotales + intNumRiesgosIdentificados
            intNumRiesgosMaterializadosTotales = intNumRiesgosMaterializadosTotales + intNumRiesgosMaterializados
            If dblIndicador < dblValorAlarma Then
                strComentarioIndicador = "El indicador está estable."
            ElseIf dblIndicador >= dblValorAlarma And dblIndicador < dblValorObjetivo Then
                strComentarioIndicador = "El indicador ha sobrepasado el Valor de la Alarma."
            ElseIf dblIndicador >= dblValorObjetivo Then
                strComentarioIndicador = "El indicador ha sobrepasado el Valor Objetivo."
            End If
        strMensaje = strMensaje & "<tr>" & vbNewLine
            strMensaje = strMensaje & "<td class=""CabeceraLateralIndicadorRiesgos"">" & strProyectoCompleto & " </td>" & vbNewLine
            strMensaje = strMensaje & "<td class=""centrado"">" & intNumRiesgosIdentificados & " </td>" & vbNewLine
            strMensaje = strMensaje & "<td class=""centrado"">" & intNumRiesgosMaterializados & " </td>" & vbNewLine
            strMensaje = strMensaje & "<td class=""centrado"">" & ValorRedondeado(dblIndicador, 2) & " </td>" & vbNewLine
            strMensaje = strMensaje & "<td class=""centrado"">" & dblValorObjetivo & " </td>" & vbNewLine
            strMensaje = strMensaje & "<td class=""centrado"">" & dblValorAlarma & " </td>" & vbNewLine
            If strComentarioIndicador = "El indicador ha sobrepasado el Valor Objetivo" Then
                strMensaje = strMensaje & "<td class=""centradoNegrita"">" & strComentarioIndicador & " </td>" & vbNewLine
            Else
                strMensaje = strMensaje & "<td class=""centrado"">" & strComentarioIndicador & " </td>" & vbNewLine
            End If
            
        strMensaje = strMensaje & "</tr>" & vbNewLine
            
        Next
    End If
    If intNumRiesgosIdentificadosTotales = 0 Then
        dblIndicador = 0
    Else
        dblIndicador = (intNumRiesgosMaterializadosTotales / intNumRiesgosIdentificadosTotales) * 100
    End If
    If dblIndicador < dblValorAlarma Then
        strComentarioIndicador = "El indicador está estable."
    ElseIf dblIndicador >= dblValorAlarma And dblIndicador < dblValorObjetivo Then
        strComentarioIndicador = "El indicador ha sobrepasado el Valor de la Alarma."
    ElseIf dblIndicador >= dblValorObjetivo Then
        strComentarioIndicador = "El indicador ha sobrepasado el Valor Objetivo."
    End If
    strMensaje = strMensaje & "<tr>" & vbNewLine
        strMensaje = strMensaje & "<td class=""CabeceraLateralIndicadorRiesgosNegrita"">" & strTextoSemestre & " </td>" & vbNewLine
        strMensaje = strMensaje & "<td class=""ResumenIndicadorRiesgos"">" & intNumRiesgosIdentificadosTotales & " </td>" & vbNewLine
        strMensaje = strMensaje & "<td class=""ResumenIndicadorRiesgos"">" & intNumRiesgosMaterializadosTotales & " </td>" & vbNewLine
        strMensaje = strMensaje & "<td class=""ResumenIndicadorRiesgosNegrita"">" & ValorRedondeado(dblIndicador, 2) & " </td>" & vbNewLine
        strMensaje = strMensaje & "<td class=""ResumenIndicadorRiesgosNegrita"">" & dblValorObjetivo & " </td>" & vbNewLine
        strMensaje = strMensaje & "<td class=""ResumenIndicadorRiesgosNegrita"">" & dblValorAlarma & " </td>" & vbNewLine
        strMensaje = strMensaje & "<td class=""ResumenIndicadorRiesgosNegrita"">" & strComentarioIndicador & " </td>" & vbNewLine
    strMensaje = strMensaje & "</tr>" & vbNewLine
    strMensaje = strMensaje & "</table>" & vbNewLine
    DameHTMLSemestre = strMensaje
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameHTMLSemestre ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.description
    End If
    DameHTMLSemestre = "#ERR" & "||" & strTextoError
End Function
Public Function DameHTMLResumenAnual() As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 23/04/2020
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       colIndicador.Add strIDProyecto & "|" & strProyecto & " " & strNombreProyecto & "|" & intNumRiesgosIdentificados & "|" & intNumRiesgosMaterializados & "|" & dblIndicador
    '   -Llamada desde
   
    '   -Devuelve:
    '       DameHTMLResumenAnual =strMensaje
    '       DameHTMLResumenAnual = "ERR" & "||" & strTextoError
    '--------------------------------------------------------------------------------------------------------
    Dim intIDentificadosResumenAnual As Integer, intMaterializadosResumenAnual As Integer, dblIndicadorResumenAnual As Double, strComentarioResumenAnual As String, _
        intIDentificadosSemestre1 As Integer, intMaterializadosSemestre1 As Integer, dblIndicadorSemestre1 As Double, strComentarioSemestre1 As String, _
        intIDentificadosSemestre2 As Integer, intMaterializadosSemestre2 As Integer, dblIndicadorSemestre2 As Double, strComentarioSemestre2 As String
    Dim strComentarioIndicador As String, strMensaje As String, strTextoError As String
    On Error GoTo errores
    If colIndicadorSemestre1ParaAño.count > 0 Then
        'colIndicadorSemestre1ParaAño.Add intIDentificadosSemestre1 & "|" & intMaterializadosSemestre1 & "|" & dblIndicadorSemestre1
        varItem = colIndicadorSemestre1ParaAño(1)
        dato = Split(varItem, "|")
        intIDentificadosSemestre1 = dato(0)
        intMaterializadosSemestre1 = dato(1)
        dblIndicadorSemestre1 = dato(2)
        If dblIndicadorSemestre1 < dblValorAlarma Then
            strComentarioSemestre1 = "El indicador está estable."
        ElseIf dblIndicadorSemestre1 >= dblValorAlarma And dblIndicadorSemestre1 < dblValorObjetivo Then
            strComentarioSemestre1 = "El indicador ha sobrepasado el Valor de la Alarma."
        ElseIf dblIndicadorSemestre1 >= dblValorObjetivo Then
            strComentarioSemestre1 = "El indicador ha sobrepasado el Valor Objetivo."
        End If
    End If
    If colIndicadorSemestre2ParaAño.count > 0 Then
        'colIndicadorSemestre2ParaAño.Add intIDentificadosSemestre2 & "|" & intMaterializadosSemestre2 & "|" & dblIndicadorSemestre2
        varItem = colIndicadorSemestre2ParaAño(1)
        dato = Split(varItem, "|")
        intIDentificadosSemestre2 = dato(0)
        intMaterializadosSemestre2 = dato(1)
        dblIndicadorSemestre2 = dato(2)
        If dblIndicadorSemestre2 < dblValorAlarma Then
            strComentarioSemestre2 = "El indicador está estable."
        ElseIf dblIndicadorSemestre2 >= dblValorAlarma And dblIndicadorSemestre2 < dblValorObjetivo Then
            strComentarioSemestre2 = "El indicador ha sobrepasado el Valor de la Alarma."
        ElseIf dblIndicadorSemestre2 >= dblValorObjetivo Then
            strComentarioSemestre2 = "El indicador ha sobrepasado el Valor Objetivo."
        End If
    End If
    If colIndicadorResumenAnual.count > 0 Then
        'colIndicadorResumenAnual.Add intIDentificadosResumenAnual & "|" & intMaterializadosResumenAnual & "|" & dblIndicadorResumenAnual
        varItem = colIndicadorResumenAnual(1)
        dato = Split(varItem, "|")
        intIDentificadosResumenAnual = dato(0)
        intMaterializadosResumenAnual = dato(1)
        dblIndicadorResumenAnual = dato(2)
        If dblIndicadorResumenAnual < dblValorAlarma Then
            strComentarioResumenAnual = "El indicador está estable."
        ElseIf dblIndicadorResumenAnual >= dblValorAlarma And dblIndicadorResumenAnual < dblValorObjetivo Then
            strComentarioResumenAnual = "El indicador ha sobrepasado el Valor de la Alarma."
        ElseIf dblIndicadorResumenAnual >= dblValorObjetivo Then
            strComentarioResumenAnual = "El indicador ha sobrepasado el Valor Objetivo."
        End If
    End If
    strMensaje = "<table>" & vbNewLine
        strMensaje = strMensaje & "<tr>" & vbNewLine
            strMensaje = strMensaje & "<td colspan='7' class=""ColespanArribaIndicadorRiesgos""> Indicador: Porcentaje de riesgos materializados sobre el total de riesgos detectados de cada proyecto</td>"
        strMensaje = strMensaje & "</tr>" & vbNewLine
        strMensaje = strMensaje & "<tr>" & vbNewLine
            strMensaje = strMensaje & "<td class=""CabeceraSuperiorIndicadorRiesgos"">Periodo</td>" & vbNewLine
            strMensaje = strMensaje & "<td class=""CabeceraSuperiorIndicadorRiesgos"">Riesgos Identificados</td>" & vbNewLine
            strMensaje = strMensaje & "<td class=""CabeceraSuperiorIndicadorRiesgos"">Riesgos identificados materializados</td>" & vbNewLine
            strMensaje = strMensaje & "<td class=""CabeceraSuperiorIndicadorRiesgos"">Resultado</td>" & vbNewLine
            strMensaje = strMensaje & "<td class=""CabeceraSuperiorIndicadorRiesgos"">Valor Objetivo</td>" & vbNewLine
            strMensaje = strMensaje & "<td class=""CabeceraSuperiorIndicadorRiesgos"">Valor de Alarma</td>" & vbNewLine
            strMensaje = strMensaje & "<td class=""CabeceraSuperiorIndicadorRiesgos"">¿Dentro de los límiteSí</td>" & vbNewLine
        strMensaje = strMensaje & "</tr>" & vbNewLine
        strMensaje = strMensaje & "<tr>" & vbNewLine
            strMensaje = strMensaje & "<td class=""CabeceraLateralIndicadorRiesgos""> Primer Semestre </td>" & vbNewLine
            strMensaje = strMensaje & "<td class=""centrado"">" & intIDentificadosSemestre1 & " </td>" & vbNewLine
            strMensaje = strMensaje & "<td class=""centrado"">" & intMaterializadosSemestre1 & " </td>" & vbNewLine
            strMensaje = strMensaje & "<td class=""centrado"">" & ValorRedondeado(dblIndicadorSemestre1, 2) & " </td>" & vbNewLine
            strMensaje = strMensaje & "<td class=""centrado"">" & dblValorObjetivo & " </td>" & vbNewLine
            strMensaje = strMensaje & "<td class=""centrado"">" & dblValorAlarma & " </td>" & vbNewLine
            If strComentarioSemestre1 = "El indicador ha sobrepasado el Valor Objetivo." Then
                strMensaje = strMensaje & "<td class=""centradoNegrita"">" & strComentarioSemestre1 & " </td>" & vbNewLine
            Else
                strMensaje = strMensaje & "<td class=""centrado"">" & strComentarioSemestre1 & " </td>" & vbNewLine
            End If
        strMensaje = strMensaje & "</tr>" & vbNewLine
        strMensaje = strMensaje & "<tr>" & vbNewLine
            strMensaje = strMensaje & "<td class=""CabeceraLateralIndicadorRiesgos"">" & "Segundo Semestre </td>" & vbNewLine
            strMensaje = strMensaje & "<td class=""centrado"">" & intIDentificadosSemestre2 & " </td>" & vbNewLine
            strMensaje = strMensaje & "<td class=""centrado"">" & intMaterializadosSemestre2 & " </td>" & vbNewLine
            strMensaje = strMensaje & "<td class=""centrado"">" & ValorRedondeado(dblIndicadorSemestre2, 2) & " </td>" & vbNewLine
            strMensaje = strMensaje & "<td class=""centrado"">" & dblValorObjetivo & " </td>" & vbNewLine
            strMensaje = strMensaje & "<td class=""centrado"">" & dblValorAlarma & " </td>" & vbNewLine
            If strComentarioSemestre2 = "El indicador ha sobrepasado el Valor Objetivo." Then
                strMensaje = strMensaje & "<td class=""centradoNegrita"">" & strComentarioSemestre2 & " </td>" & vbNewLine
            Else
                strMensaje = strMensaje & "<td class=""centrado"">" & strComentarioSemestre2 & " </td>" & vbNewLine
            End If
        strMensaje = strMensaje & "</tr>" & vbNewLine
        strMensaje = strMensaje & "<tr>" & vbNewLine
            strMensaje = strMensaje & "<td class=""CabeceraLateralIndicadorRiesgosNegrita"">ANUAL</td>" & vbNewLine
            strMensaje = strMensaje & "<td class=""ResumenIndicadorRiesgos"">" & intIDentificadosResumenAnual & " </td>" & vbNewLine
            strMensaje = strMensaje & "<td class=""ResumenIndicadorRiesgos"">" & intMaterializadosResumenAnual & " </td>" & vbNewLine
            strMensaje = strMensaje & "<td class=""ResumenIndicadorRiesgosNegrita"">" & ValorRedondeado(dblIndicadorResumenAnual, 2) & " </td>" & vbNewLine
            strMensaje = strMensaje & "<td class=""ResumenIndicadorRiesgosNegrita"">" & dblValorObjetivo & " </td>" & vbNewLine
            strMensaje = strMensaje & "<td class=""ResumenIndicadorRiesgosNegrita"">" & dblValorAlarma & " </td>" & vbNewLine
            strMensaje = strMensaje & "<td class=""ResumenIndicadorRiesgosNegrita"">" & strComentarioResumenAnual & " </td>" & vbNewLine
        strMensaje = strMensaje & "</tr>" & vbNewLine
    strMensaje = strMensaje & "</table>" & vbNewLine
    DameHTMLResumenAnual = strMensaje
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameHTMLResumenAnual ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.description
    End If
    DameHTMLResumenAnual = "#ERR" & "||" & strTextoError
End Function
Public Function DameInformeAñoCompleto(strAño As String) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 23/04/2020
    '   -Modificaciones:
    
    '   -Funcionamiento:
    
    '   -Llamada desde
   
    '   -Devuelve:
    '       DameInformeAñoCompleto =Descriptivo
    '       DameInformeAñoCompleto = "ERR" & "|" & strTextoError
    '--------------------------------------------------------------------------------------------------------
    Dim strMensaje As String, strHTMLTablaSemestre1 As String, strHTMLTablaSemestre2 As String, strHTMLDetalleRiesgo As String, strHTMLResumen As String, strTextoError As String
    On Error GoTo errores
    '-------------------------------------------------------
    '   -Devuelve:
    '       DameHTMLSemestre =strMensaje
    '       DameHTMLSemestre = "ERR" & "||" & strTextoError
    '-------------------------------------------------------
    flag = DameHTMLSemestre("1", strAño)
    If InStr(1, flag, "||") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameHTMLSemestre ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strHTMLTablaSemestre1 = flag
    '-------------------------------------------------------
    '   -Devuelve:
    '       DameHTMLSemestre =strMensaje
    '       DameHTMLSemestre = "ERR" & "||" & strTextoError
    '-------------------------------------------------------
    flag = DameHTMLSemestre("2", strAño)
    If InStr(1, flag, "||") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameHTMLSemestre ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strHTMLTablaSemestre2 = flag
    '-----------------------------------------------------------------------
    '   -Devuelve:
    '       DameHTMLResumenAnual =strMensaje
    '       DameHTMLResumenAnual = "ERR" & "||" & strTextoError
    '-----------------------------------------------------------------------
    flag = DameHTMLResumenAnual()
    If InStr(1, flag, "||") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameHTMLResumenAnual ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strHTMLResumen = flag
    '-------------------------------------------------------
    '   -Devuelve:
    '       DameHTMLDetalleRiesgos =strMensaje
    '       DameHTMLDetalleRiesgos = "ERR" & "||" & strTextoError
    '-------------------------------------------------------
    flag = DameHTMLDetalleRiesgos(strAño)
    If InStr(1, flag, "||") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameHTMLDetalleRiesgos ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strHTMLDetalleRiesgo = flag
    
    '------------------------------------------
    '   -Devuelve:
    '       DameCabeceraHTML = strMensaje
    '       DameCabeceraHTML = "-1"
    '------------------------------------------
    flag = DameCabeceraHTML("Indicador")
    If flag = "-1" Then
        strTextoError = "El método DameCabeceraHTML ha devuelto un error desconocido"
        Err.Raise 1000
    End If
    strMensaje = flag & vbNewLine
    strMensaje = strMensaje & strHTMLTablaSemestre1 & vbNewLine
    strMensaje = strMensaje & "<br /><br />" & vbNewLine
    strMensaje = strMensaje & strHTMLTablaSemestre2 & vbNewLine
    strMensaje = strMensaje & "<br /><br />" & vbNewLine
    strMensaje = strMensaje & strHTMLResumen & vbNewLine
    strMensaje = strMensaje & "<br /><br />" & vbNewLine
    strMensaje = strMensaje & strHTMLDetalleRiesgo & vbNewLine
    strMensaje = strMensaje & "<br /><br />" & vbNewLine
    strMensaje = strMensaje & "</body>" & vbNewLine
    strMensaje = strMensaje & "</html>" & vbNewLine
    '---------------------------------------------------
    '   -Devuelve:
    '       HTMLENTXT = strURLCompletaArchivo
    '       HTMLENTXT = "-1"
    '---------------------------------------------------
    flag = HTMLENTXT(strMensaje)
    DameInformeAñoCompleto = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameInformeAñoCompleto ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.description
    End If
    DameInformeAñoCompleto = "#ERR" & "|" & strTextoError
End Function
Public Function DameInformeSemestre( _
                                    strSemestre As String, _
                                    strAño As String _
                                    ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 23/04/2020
    '   -Modificaciones:
    
    '   -Funcionamiento:
    
    '   -Llamada desde
   
    '   -Devuelve:
    '       DameInformeSemestre =Descriptivo
    '       DameInformeSemestre = "ERR" & "|" & strTextoError
    '--------------------------------------------------------------------------------------------------------
    Dim strMensaje As String, strHTMLTablaSemestre As String, strHTMLDetalleRiesgo As String, strTextoError As String
    On Error GoTo errores
    '-------------------------------------------------------
    '   -Devuelve:
    '       DameHTMLSemestre =strMensaje
    '       DameHTMLSemestre = "ERR" & "||" & strTextoError
    '-------------------------------------------------------
    flag = DameHTMLSemestre(strSemestre, strAño)
    If InStr(1, flag, "||") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameHTMLSemestre ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strHTMLTablaSemestre = flag
    '-------------------------------------------------------
    '   -Devuelve:
    '       DameHTMLDetalleRiesgos =strMensaje
    '       DameHTMLDetalleRiesgos = "ERR" & "||" & strTextoError
    '-------------------------------------------------------
    flag = DameHTMLDetalleRiesgos(strAño)
    If InStr(1, flag, "||") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameHTMLDetalleRiesgos ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strHTMLDetalleRiesgo = flag
    '------------------------------------------
    '   -Devuelve:
    '       DameCabeceraHTML = strMensaje
    '       DameCabeceraHTML = "-1"
    '------------------------------------------
    flag = DameCabeceraHTML("Indicador")
    If flag = "-1" Then
        strTextoError = "El método DameCabeceraHTML ha devuelto un error desconocido"
        Err.Raise 1000
    End If
    strMensaje = flag & vbNewLine
    strMensaje = strMensaje & strHTMLTablaSemestre & vbNewLine
    strMensaje = strMensaje & "<br /><br />" & vbNewLine
    strMensaje = strMensaje & strHTMLDetalleRiesgo & vbNewLine
    strMensaje = strMensaje & "<br /><br />" & vbNewLine
    strMensaje = strMensaje & "</body>" & vbNewLine
    strMensaje = strMensaje & "</html>" & vbNewLine
    '---------------------------------------------------
    '   -Devuelve:
    '       HTMLENTXT = strURLCompletaArchivo
    '       HTMLENTXT = "-1"
    '---------------------------------------------------
    flag = HTMLENTXT(strMensaje)
    DameInformeSemestre = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameInformeSemestre ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.description
    End If
    DameInformeSemestre = "#ERR" & "|" & strTextoError
End Function
Public Function ValorRedondeado(dblValor As Double, intNumeroDecimales As Integer) As Double
    Dim strParteDecimal As String, dato, strParteNoDecimal As String, intDecimalSiguiente As Integer, intUltimoDecimal As Integer
    If intNumeroDecimales > 0 Then
        If InStr(1, dblValor, ",") <> 0 Or InStr(1, dblValor, ".") <> 0 Then
            If InStr(1, dblValor, ",") <> 0 Then
                dato = Split(dblValor, ",")
                strParteNoDecimal = dato(0)
                strParteDecimal = dato(1)
            Else
                If InStr(1, dblValor, ".") <> 0 Then
                    dato = Split(dblValor, ".")
                    strParteNoDecimal = dato(0)
                    strParteDecimal = dato(1)
                End If
            End If
            If Len(strParteDecimal) > intNumeroDecimales Then
                intUltimoDecimal = Mid(strParteDecimal, intNumeroDecimales, 1)
                intDecimalSiguiente = Mid(strParteDecimal, intNumeroDecimales + 1, 1)
                If intDecimalSiguiente >= 5 Then
                    intUltimoDecimal = intUltimoDecimal + 1
                End If
                ValorRedondeado = CDbl(strParteNoDecimal & "," & Left(strParteDecimal, intNumeroDecimales - 1) & intUltimoDecimal)
            Else
                ValorRedondeado = dblValor
            End If
        Else
            ValorRedondeado = dblValor
        End If
    Else
        ValorRedondeado = dblValor
    End If
End Function
Public Function DameHTMLDetalleRiesgos(strAño As String) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 23/04/2020
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       CadenaRiesgos="R01:01/01/2019:15/06/2019;R02:
    '       -colRiesgosDetalle.Add strProyecto & " " & strNombreProyecto & "|" & CadenaRiesgos
    '   -Llamada desde
   
    '   -Devuelve:
    '       DameHTMLDetalleRiesgos =strMensaje
    '       DameHTMLDetalleRiesgos = "ERR" & "||" & strTextoError
    '--------------------------------------------------------------------------------------------------------
    Dim strProyectoCompleto As String, strCadenaRiesgos As String, strRiesgo As String, strFechaDetectado As String, strFechaMaterializado As String, strMensaje As String, _
         dato As Variant, dato1 As Variant, dato2 As Variant, VarItem1 As Variant, strClase As String, strTextoError As String
    On Error GoTo errores
    If colRiesgosDetalle.count > 0 Then
        For Each varItem In colRiesgosDetalle
            dato = Split(varItem, "|")
            strProyectoCompleto = dato(0)
            'Debug.Print varItem
            'If strProyectoCompleto = "0186_23 ADQUISICIÓN URGENTE DE EQUIPAMIENTO RADIO SDR DE V/UHF 2023 DEL SCRT." Then Stop
            strCadenaRiesgos = dato(1)
            strMensaje = strMensaje & "<table>" & vbNewLine
            strMensaje = strMensaje & "<tr>" & vbNewLine
               strMensaje = strMensaje & "<td colspan='3' class=""ColespanArribaIndicadorRiesgos"">" & strProyectoCompleto & "</td>"
            strMensaje = strMensaje & "</tr>" & vbNewLine
            strMensaje = strMensaje & "<tr>" & vbNewLine
               strMensaje = strMensaje & "<td class=""CabeceraSuperiorIndicadorRiesgosNegrita"">Riesgos</td>" & vbNewLine
               strMensaje = strMensaje & "<td class=""CabeceraSuperiorIndicadorRiesgosNegrita"">Fecha Detectado</td>" & vbNewLine
               strMensaje = strMensaje & "<td class=""CabeceraSuperiorIndicadorRiesgosNegrita"">Fecha Materializado</td>" & vbNewLine
            strMensaje = strMensaje & "</tr>" & vbNewLine
            If InStr(1, strCadenaRiesgos, ";") <> 0 Then
                dato1 = Split(strCadenaRiesgos, ";")
                For Each VarItem1 In dato1
                    dato2 = Split(VarItem1, ":")
                    strRiesgo = dato2(0)
                    strFechaDetectado = dato2(1)
                    strFechaMaterializado = dato2(2)
                    strMensaje = strMensaje & "<tr>" & vbNewLine
                        strMensaje = strMensaje & "<td class=""CabeceraLateralIndicadorRiesgosNegrita"">" & strRiesgo & " </td>" & vbNewLine
                        If IsDate(strFechaDetectado) Then
                            If Year(CDate(strFechaDetectado)) = strAño Then
                                strClase = "centrado"
                            Else
                                strClase = "centradoNegrita"
                            End If
                        Else
                            strClase = "centrado"
                        End If
                        strClase = "centrado"
                        strMensaje = strMensaje & "<td class=""" & strClase & """>" & TraducirFechaASemestreAño(strFechaDetectado) & " </td>" & vbNewLine
                        If IsDate(strFechaMaterializado) Then
                            If Year(CDate(strFechaMaterializado)) = strAño Then
                                strClase = "centrado"
                            Else
                                strClase = "centradoNegrita"
                            End If
                        Else
                            strClase = "centrado"
                        End If
                        strClase = "centrado"
                        strMensaje = strMensaje & "<td class=""" & strClase & """>" & TraducirFechaASemestreAño(strFechaMaterializado) & " </td>" & vbNewLine
                    strMensaje = strMensaje & "</tr>" & vbNewLine
                Next
            Else
                If InStr(1, strCadenaRiesgos, ":") <> 0 Then
                    dato1 = Split(strCadenaRiesgos, ":")
                    strRiesgo = dato1(0)
                    strFechaDetectado = dato1(1)
                    strFechaMaterializado = dato1(2)
                    strMensaje = strMensaje & "<tr>" & vbNewLine
                        strMensaje = strMensaje & "<td class=""CabeceraLateralIndicadorRiesgosNegrita"">" & strRiesgo & " </td>" & vbNewLine
                            If IsDate(strFechaDetectado) Then
                                If Year(CDate(strFechaDetectado)) = strAño Then
                                    strClase = "centrado"
                                Else
                                    strClase = "centradoNegrita"
                                End If
                            Else
                                strClase = "centrado"
                            End If
                            strClase = "centrado"
                            strMensaje = strMensaje & "<td class=""" & strClase & """>" & TraducirFechaASemestreAño(strFechaDetectado) & " </td>" & vbNewLine
                            If IsDate(strFechaMaterializado) Then
                                If Year(CDate(strFechaMaterializado)) = strAño Then
                                    strClase = "centrado"
                                Else
                                    strClase = "centradoNegrita"
                                End If
                            Else
                                strClase = "centrado"
                            End If
                            strClase = "centrado"
                            strMensaje = strMensaje & "<td class=""" & strClase & """>" & TraducirFechaASemestreAño(strFechaMaterializado) & " </td>" & vbNewLine
                    strMensaje = strMensaje & "</tr>" & vbNewLine
                End If
                
            End If
            strMensaje = strMensaje & "</table>" & vbNewLine
            strMensaje = strMensaje & "<br /><br />" & vbNewLine
        Next
    End If
    DameHTMLDetalleRiesgos = strMensaje
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameHTMLDetalleRiesgos ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.description
    End If
    DameHTMLDetalleRiesgos = "#ERR" & "||" & strTextoError
End Function
Public Function TraducirFechaASemestreAño(strFecha As String) As String
    Dim strAño As String, strTextoTotal As String, strTextoError As String
    On Error GoTo errores
    If Not IsDate(strFecha) Then
        strTextoTotal = "&nbsp;"
    Else
        strAño = Year(CDate(strFecha))
        If CInt(Month(CDate(strFecha))) < 7 Then
            strTextoTotal = "SEMESTRE 1 " & strAño & " (" & Format(strFecha, "dd/mm/yyyy") & ")"
        Else
            strTextoTotal = "SEMESTRE 2 " & strAño & " (" & Format(strFecha, "dd/mm/yyyy") & ")"
        End If
    End If
    TraducirFechaASemestreAño = strTextoTotal
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método TraducirFechaASemestreAño ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.description
    End If
    TraducirFechaASemestreAño = "#ERR" & "|" & strTextoError
End Function
Public Function DameNumeroRiesgosDetectadosMaterializados( _
                                                            strAño As String, _
                                                            Optional strSemestre As String, _
                                                            Optional strIDProyectoOrigen As String _
                                                            ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 23/04/2020
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       CadenaRiesgos="R01:01/01/2019:15/06/2019;R02:
    '       -colRiesgosDetalle.Add strProyecto & " " & strNombreProyecto & "|" & CadenaRiesgos
    '   -Llamada desde
   
    '   -Devuelve:
    '       strResultado=intIDentificados & ";" &  intMaterializados-->strIDProyecto<>""
    '       strResultado=intIDentificadosSemestre1 & ";" & intMaterializadosSemestre1 & "#" & _
                        intIDentificadosSemestre2 & ";" & intMaterializadosSemestre2 & "#" & _
                        intIDentificadosResumenAnual & ";" & intMaterializadosResumenAnual -->strIDProyecto=""
    '       DameNumeroRiesgosDetectadosMaterializados =
    '       DameNumeroRiesgosDetectadosMaterializados = "ERR" & "|" & strTextoError
    '--------------------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strFechaInicialParaConsulta As String, strFechaFinalParaConsulta As String, intIDentificados As Integer, strIDProyecto As String, intMaterializados As Integer, _
        strFechaInicialParaConsultaSemestre1 As String, strFechaInicialParaConsultaSemestre2 As String, intNumero As Integer, intNumeroTotal As Integer, _
        strFechaFinalParaConsultaSemestre1 As String, strFechaFinalParaConsultaSemestre2 As String, intIDentificadosSemestre1 As Integer, intIDentificadosSemestre2 As Integer, _
        intMaterializadosSemestre1 As Integer, intMaterializadosSemestre2 As Integer, intIDentificadosResumenAnual As Integer, intMaterializadosResumenAnual As Integer, strTextoError As String
    On Error GoTo errores
    
    If strAño = "" Then
        strTextoError = "No se puede dejar en blanco el año"
        Err.Raise 1000
    End If
    If strSemestre = "1" Then
        strFechaInicialParaConsulta = "01/01/" & strAño
        strFechaFinalParaConsulta = Format("30/06/" & strAño, "mm/dd/yyyy")
    ElseIf strSemestre = "2" Then
        strFechaInicialParaConsulta = Format("01/07/" & strAño, "mm/dd/yyyy")
        strFechaFinalParaConsulta = Format("31/12/" & strAño, "mm/dd/yyyy")
    Else
        strFechaInicialParaConsultaSemestre1 = "01/01/" & strAño
        strFechaFinalParaConsultaSemestre1 = Format("30/06/" & strAño, "mm/dd/yyyy")
        strFechaInicialParaConsultaSemestre2 = Format("01/07/" & strAño, "mm/dd/yyyy")
        strFechaFinalParaConsultaSemestre2 = Format("31/12/" & strAño, "mm/dd/yyyy")
    End If
    If strIDProyectoOrigen <> "" Then
        Avance "Obteniendo Fecha Detectado"
        
        m_SQL = "SELECT Count(TbAuxProyectosRiesgos.FechaDetectado) AS CuentaDeFechaDetectado " & _
                "FROM TbAuxProyectosRiesgos " & _
                "WHERE (((TbAuxProyectosRiesgos.IDProyecto) =" & strIDProyectoOrigen & ") AND " & _
                "((TbAuxProyectosRiesgos.FechaDetectado)<=#" & strFechaFinalParaConsulta & "#));"
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
        With rcdDatos
            If Not .EOF Then
               intIDentificados = CInt(.fields("CuentaDeFechaDetectado"))
            End If
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
        Avance "Obteniendo Fecha Materializado"
        
        m_SQL = "SELECT Count(TbAuxProyectosRiesgos.FechaMaterializado) AS CuentaDeFechaMaterializado " & _
                "FROM TbAuxProyectosRiesgos " & _
                "WHERE (((TbAuxProyectosRiesgos.IDProyecto) =" & strIDProyectoOrigen & ") AND " & _
                "((TbAuxProyectosRiesgos.FechaMaterializado) Between #" & strFechaInicialParaConsulta & "# And #" & strFechaFinalParaConsulta & "#));"
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
        With rcdDatos
            If Not .EOF Then
               intMaterializados = CInt(.fields("CuentaDeFechaMaterializado"))
            End If
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
        DameNumeroRiesgosDetectadosMaterializados = intIDentificados & ";" & intMaterializados
    Else
        If ColProyectosParaInforme.count > 0 Then
            intNumeroTotal = CInt(ColProyectosParaInforme.count)
            intNumero = 1
            For Each varItem In ColProyectosParaInforme
                dato = Split(varItem, "|")
                strIDProyectoOrigen = dato(0)
                Avance "Fecha Detectado Semestre1 (" & intNumero & " de " & intNumeroTotal & ")"
                
                m_SQL = "SELECT Count(TbAuxProyectosRiesgos.FechaDetectado) AS CuentaDeFechaDetectado " & _
                        "FROM TbAuxProyectosRiesgos " & _
                        "WHERE (((TbAuxProyectosRiesgos.IDProyecto) =" & strIDProyectoOrigen & ") AND " & _
                        "((TbAuxProyectosRiesgos.FechaDetectado)<=#" & strFechaFinalParaConsultaSemestre1 & "#));"
                Set rcdDatos = getdb().OpenRecordset(m_SQL)
                With rcdDatos
                    If Not .EOF Then
                       intIDentificadosSemestre1 = intIDentificadosSemestre1 + CInt(.fields("CuentaDeFechaDetectado"))
                    End If
                End With
                rcdDatos.Close
                Set rcdDatos = Nothing
                Avance "Fecha Detectado Semestre 2 (" & intNumero & " de " & intNumeroTotal & ")"
                
                m_SQL = "SELECT Count(TbAuxProyectosRiesgos.FechaDetectado) AS CuentaDeFechaDetectado " & _
                        "FROM TbAuxProyectosRiesgos " & _
                        "WHERE (((TbAuxProyectosRiesgos.IDProyecto) =" & strIDProyectoOrigen & ") AND " & _
                        "((TbAuxProyectosRiesgos.FechaDetectado)<=#" & strFechaFinalParaConsultaSemestre2 & "#));"
                Set rcdDatos = getdb().OpenRecordset(m_SQL)
                With rcdDatos
                    If Not .EOF Then
                       intIDentificadosSemestre2 = intIDentificadosSemestre2 + CInt(.fields("CuentaDeFechaDetectado"))
                    End If
                End With
                rcdDatos.Close
                Set rcdDatos = Nothing
                intNumero = intNumero + 1
            Next
            intIDentificadosResumenAnual = intIDentificadosSemestre2
            If ColProyectosParaInforme.count > 0 Then
                intNumeroTotal = CInt(ColProyectosParaInforme.count)
                intNumero = 1
                For Each varItem In ColProyectosParaInforme
                    dato = Split(varItem, "|")
                    strIDProyectoOrigen = dato(0)
                    Avance "Fecha Materializado Semestre1 (" & intNumero & " de " & intNumeroTotal & ")"
                    
                    m_SQL = "SELECT Count(TbAuxProyectosRiesgos.FechaMaterializado) AS CuentaDeFechaMaterializado " & _
                            "FROM TbAuxProyectosRiesgos " & _
                            "WHERE (((TbAuxProyectosRiesgos.IDProyecto) =" & strIDProyectoOrigen & ") AND " & _
                            "((TbAuxProyectosRiesgos.FechaMaterializado) Between #" & strFechaInicialParaConsultaSemestre1 & "# And #" & strFechaFinalParaConsultaSemestre1 & "#));"
                    Set rcdDatos = getdb().OpenRecordset(m_SQL)
                    With rcdDatos
                        If Not .EOF Then
                           intMaterializadosSemestre1 = intMaterializadosSemestre1 + CInt(.fields("CuentaDeFechaMaterializado"))
                        End If
                    End With
                    rcdDatos.Close
                    Set rcdDatos = Nothing
                    Avance "Fecha Materializado Semestre2 (" & intNumero & " de " & intNumeroTotal & ")"
                    
                    m_SQL = "SELECT Count(TbAuxProyectosRiesgos.FechaMaterializado) AS CuentaDeFechaMaterializado " & _
                            "FROM TbAuxProyectosRiesgos " & _
                            "WHERE (((TbAuxProyectosRiesgos.IDProyecto) =" & strIDProyectoOrigen & ") AND " & _
                            "((TbAuxProyectosRiesgos.FechaMaterializado) Between #" & strFechaInicialParaConsultaSemestre2 & "# And #" & strFechaFinalParaConsultaSemestre2 & "#));"
                    Set rcdDatos = getdb().OpenRecordset(m_SQL)
                    With rcdDatos
                        If Not .EOF Then
                           intMaterializadosSemestre2 = intMaterializadosSemestre2 + CInt(.fields("CuentaDeFechaMaterializado"))
                        End If
                    End With
                    rcdDatos.Close
                    Set rcdDatos = Nothing
                    intNumero = intNumero + 1
                 Next
                 intMaterializadosResumenAnual = intMaterializadosSemestre1 + intMaterializadosSemestre2
            End If
        End If
        DameNumeroRiesgosDetectadosMaterializados = intIDentificadosSemestre1 & ";" & intMaterializadosSemestre1 & "#" & _
                                                    intIDentificadosSemestre2 & ";" & intMaterializadosSemestre2 & "#" & _
                                                    intIDentificadosResumenAnual & ";" & intMaterializadosResumenAnual
    End If
    getdb().Close
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameNumeroRiesgosDetectadosMaterializados ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DameNumeroRiesgosDetectadosMaterializados = "#ERR" & "|" & strTextoError
End Function
Private Function DameCabeceraHTML(strTitulo As String) As String
    
    Dim strMensaje As String, strTextoError As String
    On Error GoTo errores
    If strTitulo = "" Then
        strTextoError = "Se ha de indicar el Título"
        Err.Raise 1000
    End If
    
    strMensaje = "<html lang=""es"">" & vbNewLine
    strMensaje = strMensaje & "<head>" & vbNewLine
        'strMensaje = strMensaje & "<meta charset=""utf-8"">" & vbNewLine
        strMensaje = strMensaje & "<meta charset=""ISO-8859-1"">" & vbNewLine
        strMensaje = strMensaje & "<title>" & strTitulo & "</title>" & vbNewLine
        strMensaje = strMensaje & "<style type=""text/css"">" & vbNewLine
            strMensaje = strMensaje & m_ObjEntorno.css & vbNewLine
        strMensaje = strMensaje & "</style>" & vbNewLine
    strMensaje = strMensaje & "</head>" & vbNewLine
    strMensaje = strMensaje & "<body>" & vbNewLine
    DameCabeceraHTML = strMensaje
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método HTML.DameCabeceraHTML ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.description
    End If
    DameCabeceraHTML = "ERR" & "|" & strTextoError
End Function
Private Function HTMLENTXT(strMensaje As String) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día x
    '   -Modificaciones:
  
    '   -Funcionamiento:
    '       -Copia el html en un archivo de texto
    '       -Lo renombra a html
    '       -abre el explorador windows de microsoft y muestra el archivo
    '   -Llamada desde
    '       -Form_FormConfAvisoJP.ComandoDetalle_Click
    '   -Devuelve:
    '       HTMLENTXT = strURLCompletaArchivo
    '       HTMLENTXT = "ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim WshShell As Object, IExp As Object, F1 As Object, strURLHTML As String, strNombreHTML As String, strURLTXT As String, _
        strURLCompletaArchivo As String, strNombretxt As String
    Dim strTextoError As String
    On Error GoTo errores
    If strMensaje = "" Then
        strTextoError = "No se ha indicado el HTML"
        Err.Raise 1000
    End If
    Set WshShell = CreateObject("WScript.Shell")
    Set IExp = CreateObject("InternetExplorer.Application")
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       strResultado = strURLTXT & ";" & strURLHTML
    '       DameUntxtYHtml = strResultado
    '       DameUntxtYHtml = "ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = DameUntxtYHtml()
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameUntxtYHtml ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If InStr(1, flag, ";") = 0 Then
        strTextoError = "El método DameUntxtYHtml ha devuelto un resultado con formato desconocido"
        Err.Raise 1000
    End If
    dato = Split(flag, ";")
    strURLTXT = dato(0)
    strURLHTML = dato(1)
    Set F1 = fso.CreateTextFile(strURLTXT, True)
    F1.WriteLine strMensaje
    F1.Close
    fso.GetFile(strURLTXT).Name = fso.GetBaseName(strURLTXT) & ".html"
    IExp.Width = 800
    IExp.Height = 600
    IExp.Left = 0
    IExp.Top = 0
    IExp.Visible = True
    strURLCompletaArchivo = "file:///" & strURLHTML
    IExp.Navigate strURLCompletaArchivo
    If Not WshShell Is Nothing Then
        Set WshShell = Nothing
    End If
    If Not IExp Is Nothing Then
        Set IExp = Nothing
    End If
    
    HTMLENTXT = strURLCompletaArchivo
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método HTMLENTXT ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.description
    End If
    If Not WshShell Is Nothing Then
        Set WshShell = Nothing
    End If
    
    If Not IExp Is Nothing Then
        Set IExp = Nothing
    End If
    
    HTMLENTXT = "ERR" & "|" & strTextoError
End Function
Private Function DameUntxtYHtml() As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día x
    '   -Modificaciones:
  
    '   -Funcionamiento:
    '       -Mira en strURLDirectorioLocal si no hay HTML1.txt si está pasa un bucle de 50 hasta que dé con uno que no esté
    
    '   -Llamada desde
    '       -Form_FormInforme.DameURLHTMLConsulta
    '       -HTML.HTMLENTXT
    '   -Devuelve:
    '       strResultado = strURLTXT & ";" & strURLHTML
    '       DameUntxtYHtml = strResultado
    '       DameUntxtYHtml = "ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim i As Integer, strURLTXT As String, strURLHTML As String, strNombreHTML As String, strNombretxt As String, strResultado As String, strTextoError As String
    On Error GoTo errores
    
    For i = 1 To 50
        strNombretxt = "HTML" & i & ".txt"
        strNombreHTML = "HTML" & i & ".html"
        strURLTXT = m_ObjEntorno.URLDirectorioLocal & strNombretxt
        strURLHTML = m_ObjEntorno.URLDirectorioLocal & strNombreHTML
        If Not fso.FileExists(strURLTXT) And Not fso.FileExists(strURLHTML) Then
            strResultado = strURLTXT & ";" & strURLHTML
            DameUntxtYHtml = strResultado
            
            Exit Function
        End If
    Next
    
    DameUntxtYHtml = ""
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameUntxtYHtml ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.description
    End If
    
    DameUntxtYHtml = "ERR" & "|" & strTextoError
End Function




