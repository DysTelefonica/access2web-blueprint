Attribute VB_Name = "SnapshotHelper"
Option Compare Database
Option Explicit

Private CAMPOS_PC_DICT As Object
Private CAMPOS_CDCA_DICT As Object
Private CAMPOS_CDCASUB_DICT As Object
Private CAMPOS_PCSUB_DICT As Object
Private dictInicializado As Boolean

Private Sub InicializarDicts()
    If dictInicializado Then Exit Sub
    
    Set CAMPOS_PC_DICT = CreateObject("Scripting.Dictionary")
    CAMPOS_PC_DICT.Add "refContratoInspeccionOficial", True
    CAMPOS_PC_DICT.Add "refSuministrador", True
    CAMPOS_PC_DICT.Add "denominacionContrato", True
    CAMPOS_PC_DICT.Add "suministradorNombreDir", True
    CAMPOS_PC_DICT.Add "objetoContrato", True
    CAMPOS_PC_DICT.Add "descripcionMaterialAfectado", True
    CAMPOS_PC_DICT.Add "numPlanoEspecificacion", True
    CAMPOS_PC_DICT.Add "descripcionPropuestaCambio", True
    CAMPOS_PC_DICT.Add "descripcionPropuestaCambioCont", True
    CAMPOS_PC_DICT.Add "motivoCorregirDeficiencias", True
    CAMPOS_PC_DICT.Add "motivoMejorarCapacidad", True
    CAMPOS_PC_DICT.Add "motivoAumentarNacionalizacion", True
    CAMPOS_PC_DICT.Add "motivoMejorarSeguridad", True
    CAMPOS_PC_DICT.Add "motivoMejorarFiabilidad", True
    CAMPOS_PC_DICT.Add "motivoMejorarCosteEficacia", True
    CAMPOS_PC_DICT.Add "motivoOtros", True
    CAMPOS_PC_DICT.Add "motivoOtrosDetalle", True
    CAMPOS_PC_DICT.Add "incidenciaCoste", True
    CAMPOS_PC_DICT.Add "incidenciaPlazo", True
    CAMPOS_PC_DICT.Add "incidenciaSeguridad", True
    CAMPOS_PC_DICT.Add "incidenciaFiabilidad", True
    CAMPOS_PC_DICT.Add "incidenciaMantenibilidad", True
    CAMPOS_PC_DICT.Add "incidenciaIntercambiabilidad", True
    CAMPOS_PC_DICT.Add "incidenciaVidaUtilAlmacen", True
    CAMPOS_PC_DICT.Add "incidenciaFuncionamientoFuncion", True
    CAMPOS_PC_DICT.Add "impactoClasificacion", True
    CAMPOS_PC_DICT.Add "CambioAfectaAMaterial", True
    CAMPOS_PC_DICT.Add "firmaOficinaTecnicaNombre", True
    CAMPOS_PC_DICT.Add "firmaRepSuministradorNombre", True
    CAMPOS_PC_DICT.Add "racCodigo", True
    CAMPOS_PC_DICT.Add "observacionesRAC", True
    CAMPOS_PC_DICT.Add "racNombre", True
    CAMPOS_PC_DICT.Add "racDecision", True
    CAMPOS_PC_DICT.Add "racRechazoMotivos", True
    CAMPOS_PC_DICT.Add "obsAprobacionAutoridadDiseno", True
    CAMPOS_PC_DICT.Add "NombreAutoridadDiseno", True
    CAMPOS_PC_DICT.Add "decisionFinal", True
    CAMPOS_PC_DICT.Add "obsDecisionFinal", True
    CAMPOS_PC_DICT.Add "NombreFirmanteFinal", True
    
    Set CAMPOS_CDCA_DICT = CreateObject("Scripting.Dictionary")
    CAMPOS_CDCA_DICT.Add "numContrato", True
    CAMPOS_CDCA_DICT.Add "refSuministrador", True
    CAMPOS_CDCA_DICT.Add "suministradorNombreDir", True
    CAMPOS_CDCA_DICT.Add "refDesviacionesPrevias", True
    CAMPOS_CDCA_DICT.Add "requiereModificacionContrato", True
    CAMPOS_CDCA_DICT.Add "identificacionMaterial", True
    CAMPOS_CDCA_DICT.Add "numPlanoEspecificacion", True
    CAMPOS_CDCA_DICT.Add "cantidadPeriodo", True
    CAMPOS_CDCA_DICT.Add "numSerieLote", True
    CAMPOS_CDCA_DICT.Add "causaNC", True
    CAMPOS_CDCA_DICT.Add "descripcionImpactoNC", True
    CAMPOS_CDCA_DICT.Add "descripcionImpactoNCCont", True
    CAMPOS_CDCA_DICT.Add "afectaPrestaciones", True
    CAMPOS_CDCA_DICT.Add "afectaSeguridad", True
    CAMPOS_CDCA_DICT.Add "afectaFiabilidad", True
    CAMPOS_CDCA_DICT.Add "afectaVidaUtil", True
    CAMPOS_CDCA_DICT.Add "afectaMedioambiente", True
    CAMPOS_CDCA_DICT.Add "afectaIntercambiabilidad", True
    CAMPOS_CDCA_DICT.Add "afectaMantenibilidad", True
    CAMPOS_CDCA_DICT.Add "afectaApariencia", True
    CAMPOS_CDCA_DICT.Add "afectaOtros", True
    CAMPOS_CDCA_DICT.Add "impactoCoste", True
    CAMPOS_CDCA_DICT.Add "clasificacionNC", True
    CAMPOS_CDCA_DICT.Add "esSuministradorAD", True
    CAMPOS_CDCA_DICT.Add "identificacionAutoridadDiseno", True
    CAMPOS_CDCA_DICT.Add "efectoFechaEntrega", True
    CAMPOS_CDCA_DICT.Add "firmaAprobacionRespIngenieriaNombre", True
    CAMPOS_CDCA_DICT.Add "firmaAprobacionRespProduccionNombre", True
    CAMPOS_CDCA_DICT.Add "firmaAprobacionRespCalidadNombre", True
    CAMPOS_CDCA_DICT.Add "firmaAprobacionRespDisenioNombre", True
    CAMPOS_CDCA_DICT.Add "firmaAprobacionRepresentanteSumNombre", True
    CAMPOS_CDCA_DICT.Add "racCodigo", True
    CAMPOS_CDCA_DICT.Add "observacionesRAC", True
    CAMPOS_CDCA_DICT.Add "racNombre", True
    CAMPOS_CDCA_DICT.Add "racDecision", True
    CAMPOS_CDCA_DICT.Add "decisionFinal", True
    CAMPOS_CDCA_DICT.Add "observacionesFinales", True
    CAMPOS_CDCA_DICT.Add "NombreFirmanteFinal", True
    
    Set CAMPOS_CDCASUB_DICT = CreateObject("Scripting.Dictionary")
    CAMPOS_CDCASUB_DICT.Add "refSuministrador", True
    CAMPOS_CDCASUB_DICT.Add "refSubSuministrador", True
    CAMPOS_CDCASUB_DICT.Add "suministradorPrincipalNombreDir", True
    CAMPOS_CDCASUB_DICT.Add "subSuministradorNombreDir", True
    CAMPOS_CDCASUB_DICT.Add "refDesviacionesPrevias", True
    CAMPOS_CDCASUB_DICT.Add "requiereModificacionContrato", True
    CAMPOS_CDCASUB_DICT.Add "identificacionMaterial", True
    CAMPOS_CDCASUB_DICT.Add "cantidadPeriodo", True
    CAMPOS_CDCASUB_DICT.Add "numPlanoEspecificacion", True
    CAMPOS_CDCASUB_DICT.Add "numSerieLote", True
    CAMPOS_CDCASUB_DICT.Add "causaNC", True
    CAMPOS_CDCASUB_DICT.Add "descripcionImpactoNC", True
    CAMPOS_CDCASUB_DICT.Add "descripcionImpactoNCCont", True
    CAMPOS_CDCASUB_DICT.Add "afectaPrestaciones", True
    CAMPOS_CDCASUB_DICT.Add "afectaSeguridad", True
    CAMPOS_CDCASUB_DICT.Add "afectaFiabilidad", True
    CAMPOS_CDCASUB_DICT.Add "afectaVidaUtil", True
    CAMPOS_CDCASUB_DICT.Add "afectaMedioambiente", True
    CAMPOS_CDCASUB_DICT.Add "afectaIntercambiabilidad", True
    CAMPOS_CDCASUB_DICT.Add "afectaMantenibilidad", True
    CAMPOS_CDCASUB_DICT.Add "afectaApariencia", True
    CAMPOS_CDCASUB_DICT.Add "afectaOtros", True
    CAMPOS_CDCASUB_DICT.Add "impactoCoste", True
    CAMPOS_CDCASUB_DICT.Add "clasificacionNC", True
    CAMPOS_CDCASUB_DICT.Add "esSubSuministradorAD", True
    CAMPOS_CDCASUB_DICT.Add "identificacionAutoridadDiseno", True
    CAMPOS_CDCASUB_DICT.Add "efectoFechaEntrega", True
    CAMPOS_CDCASUB_DICT.Add "firmaAprobacionRespIngenieriaNombre", True
    CAMPOS_CDCASUB_DICT.Add "firmaAprobacionRespProduccionNombre", True
    CAMPOS_CDCASUB_DICT.Add "firmaAprobacionRespCalidadNombre", True
    CAMPOS_CDCASUB_DICT.Add "firmaAprobacionRespDisenioNombre", True
    CAMPOS_CDCASUB_DICT.Add "firmaAprobacionRepresentanteSumNombre", True
    CAMPOS_CDCASUB_DICT.Add "racCodigo", True
    CAMPOS_CDCASUB_DICT.Add "observacionesRAC", True
    CAMPOS_CDCASUB_DICT.Add "racNombre", True
    CAMPOS_CDCASUB_DICT.Add "racDecision", True
    CAMPOS_CDCASUB_DICT.Add "racRechazoMotivos", True
    CAMPOS_CDCASUB_DICT.Add "observacionesRACDelegador", True
    CAMPOS_CDCASUB_DICT.Add "racNombreDelegador", True
    CAMPOS_CDCASUB_DICT.Add "decisionFinal", True
    CAMPOS_CDCASUB_DICT.Add "observacionesFinales", True
    CAMPOS_CDCASUB_DICT.Add "NombreFirmanteFinal", True
    
    ' CAMPOS_PCSUB_DICT: Clone de PC (gemelo arquitectónico)
    Set CAMPOS_PCSUB_DICT = CreateObject("Scripting.Dictionary")
    CAMPOS_PCSUB_DICT.Add "refContratoInspeccionOficial", True
    CAMPOS_PCSUB_DICT.Add "refSubSuministrador", True
    CAMPOS_PCSUB_DICT.Add "denominacionContrato", True
    CAMPOS_PCSUB_DICT.Add "SubsuministradorNombreDir", True
    CAMPOS_PCSUB_DICT.Add "objetoContrato", True
    CAMPOS_PCSUB_DICT.Add "descripcionMaterialAfectado", True
    CAMPOS_PCSUB_DICT.Add "numPlanoEspecificacion", True
    CAMPOS_PCSUB_DICT.Add "descripcionPropuestaCambio", True
    CAMPOS_PCSUB_DICT.Add "descripcionPropuestaCambioCont", True
    CAMPOS_PCSUB_DICT.Add "motivoCorregirDeficiencias", True
    CAMPOS_PCSUB_DICT.Add "motivoMejorarCapacidad", True
    CAMPOS_PCSUB_DICT.Add "motivoAumentarNacionalizacion", True
    CAMPOS_PCSUB_DICT.Add "motivoMejorarSeguridad", True
    CAMPOS_PCSUB_DICT.Add "motivoMejorarFiabilidad", True
    CAMPOS_PCSUB_DICT.Add "motivoMejorarCosteEficacia", True
    CAMPOS_PCSUB_DICT.Add "motivoOtros", True
    CAMPOS_PCSUB_DICT.Add "motivoOtrosDetalle", True
    CAMPOS_PCSUB_DICT.Add "incidenciaCoste", True
    CAMPOS_PCSUB_DICT.Add "incidenciaPlazo", True
    CAMPOS_PCSUB_DICT.Add "incidenciaSeguridad", True
    CAMPOS_PCSUB_DICT.Add "incidenciaFiabilidad", True
    CAMPOS_PCSUB_DICT.Add "incidenciaMantenibilidad", True
    CAMPOS_PCSUB_DICT.Add "incidenciaIntercambiabilidad", True
    CAMPOS_PCSUB_DICT.Add "incidenciaVidaUtilAlmacen", True
    CAMPOS_PCSUB_DICT.Add "incidenciaFuncionamientoFuncion", True
    CAMPOS_PCSUB_DICT.Add "impactoClasificacion", True
    CAMPOS_PCSUB_DICT.Add "CambioAfectaAMaterial", True
    CAMPOS_PCSUB_DICT.Add "firmaOficinaTecnicaSubSuministradorNombre", True
    CAMPOS_PCSUB_DICT.Add "firmaRepSubSuministradorNombre", True
    CAMPOS_PCSUB_DICT.Add "racCodigo", True
    CAMPOS_PCSUB_DICT.Add "observacionesRAC", True
    CAMPOS_PCSUB_DICT.Add "racNombre", True
    CAMPOS_PCSUB_DICT.Add "racDecision", True
    CAMPOS_PCSUB_DICT.Add "racRechazoMotivos", True
    CAMPOS_PCSUB_DICT.Add "observacionesRACDelegador", True
    CAMPOS_PCSUB_DICT.Add "racNombreDelegador", True
    CAMPOS_PCSUB_DICT.Add "obsAprobacionAutoridadDiseno", True
    CAMPOS_PCSUB_DICT.Add "NombreAutoridadDiseno", True
    CAMPOS_PCSUB_DICT.Add "decisionFinal", True
    CAMPOS_PCSUB_DICT.Add "obsDecisionFinal", True
    CAMPOS_PCSUB_DICT.Add "NombreFirmanteFinal", True
    
    dictInicializado = True
End Sub

Private Function ObtenerDictCampos(ByVal tipoSolicitud As String) As Object
    If Not dictInicializado Then InicializarDicts
    Select Case UCase(tipoSolicitud)
        Case "PC"
            Set ObtenerDictCampos = CAMPOS_PC_DICT
        Case "CDCA", "CD_CA"
            Set ObtenerDictCampos = CAMPOS_CDCA_DICT
        Case "CDCASUB", "CD_CA_SUB"
            Set ObtenerDictCampos = CAMPOS_CDCASUB_DICT
        Case "PC_SUB", "PCSUB"
            Set ObtenerDictCampos = CAMPOS_PCSUB_DICT
    End Select
End Function

Public Function CalcularHashFaseValidacion(ByVal idSolicitud As Long, ByVal tipoSolicitud As String, Optional ByRef db As DAO.Database) As String
    Dim jsonSnapshot As String
    
    On Error GoTo Errores
    
    jsonSnapshot = GenerarSnapshotGlobal(idSolicitud, tipoSolicitud, db)
    
    If Len(jsonSnapshot) = 0 Then
        CalcularHashFaseValidacion = ""
        Exit Function
    End If
    
    CalcularHashFaseValidacion = CalcularSHA256(jsonSnapshot)
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SnapshotHelper.CalcularHashFaseValidacion"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud & ", tipo=" & tipoSolicitud
    errObj.Raise
End Function

Public Sub ExportarSnapshotParaDebug(ByVal idSolicitud As Long, ByVal tipoSolicitud As String, ByVal contexto As String, Optional ByVal ordinalCiclo As Long = 0, Optional ByRef db As DAO.Database)
    Dim jsonSnapshot As String
    Dim hash As String
    Dim fso As Object
    Dim ts As Object
    Dim rutaLog As String
    Dim timestamp As String
    Dim rutaLogs As String
    Dim stream As Object
    Dim contenido As String
    Dim jsonPretty As String
    Dim jsonDb As String
    Dim hashDb As String
    Dim jsonDbPretty As String
    Dim hashCiclo As String
    Dim jsonCiclo As String
    Dim jsonCicloPretty As String
    
    On Error Resume Next
    
    jsonSnapshot = GenerarSnapshotGlobal(idSolicitud, tipoSolicitud, db)
    hash = CalcularSHA256(jsonSnapshot)
    jsonPretty = FormatearJsonPretty(jsonSnapshot)
    
    jsonDb = GenerarSnapshotDesdeTabla(idSolicitud, tipoSolicitud)
    hashDb = CalcularSHA256(jsonDb)
    jsonDbPretty = FormatearJsonPretty(jsonDb)
    
    If ordinalCiclo > 0 Then
        hashCiclo = ObtenerHashCiclo(idSolicitud, ordinalCiclo - 1)
        jsonCiclo = GenerarSnapshotDesdeTabla(idSolicitud, tipoSolicitud)
        jsonCicloPretty = FormatearJsonPretty(jsonCiclo)
    End If
    
    timestamp = Format(Now, "yyyyMMdd_HHmmss")
    rutaLogs = CurrentProject.path & "\logs"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(rutaLogs) Then
        fso.CreateFolder rutaLogs
    End If
    
    On Error GoTo Errores
    
    If ordinalCiclo > 0 Then
        rutaLog = rutaLogs & "\ciclo" & ordinalCiclo & "_" & contexto & "_" & timestamp & ".log"
    Else
        rutaLog = rutaLogs & "\snapshot_" & idSolicitud & "_" & contexto & "_" & timestamp & ".log"
    End If
    
    contenido = "=== SNAPSHOT CICLO " & ordinalCiclo & " ===" & vbCrLf
    contenido = contenido & "idSolicitud: " & idSolicitud & vbCrLf
    contenido = contenido & "tipoSolicitud: " & tipoSolicitud & vbCrLf
    contenido = contenido & "contexto: " & contexto & vbCrLf
    contenido = contenido & "timestamp: " & Now & vbCrLf
    contenido = contenido & vbCrLf
    contenido = contenido & "=== HASH EN TABLA (ciclo " & ordinalCiclo & ") ===" & vbCrLf
    contenido = contenido & hashCiclo & vbCrLf
    contenido = contenido & vbCrLf
    contenido = contenido & "=== JSON EN TABLA (ciclo " & ordinalCiclo & ") ===" & vbCrLf
    contenido = contenido & jsonCicloPretty & vbCrLf
    contenido = contenido & vbCrLf
    contenido = contenido & "=== HASH EN MEMORIA (calculo actual) ===" & vbCrLf
    contenido = contenido & hash & vbCrLf
    contenido = contenido & vbCrLf
    contenido = contenido & "=== JSON EN MEMORIA ===" & vbCrLf
    contenido = contenido & jsonPretty & vbCrLf
    contenido = contenido & vbCrLf
    contenido = contenido & jsonDbPretty
    
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText contenido
    stream.SaveToFile rutaLog, 2
    stream.Close
    
    Debug.Print "SNAPSHOT EXPORTADO: " & rutaLog
    Exit Sub
Errores:
    Debug.Print "ERROR ExportarSnapshotParaDebug: " & Err.description
End Sub

Private Function GenerarSnapshotDesdeTabla(ByVal idSolicitud As Long, ByVal tipoSolicitud As String) As String
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String
    Dim jsonSeccion1 As String
    Dim jsonSeccion2 As String
    Dim jsonSeccion3 As String
    
    On Error GoTo Errores
    
    Set db = getdb()
    
    Select Case UCase(tipoSolicitud)
        Case "PC"
            sql = "SELECT * FROM tbDatosPC WHERE idSolicitud = " & idSolicitud
            Set rcd = db.OpenRecordset(sql, dbOpenSnapshot)
            If Not rcd.EOF Then
                jsonSeccion1 = ObtenerJsonDesdeRecordset(rcd, "PC")
            Else
                jsonSeccion1 = "{}"
            End If
            jsonSeccion2 = "{}"
            jsonSeccion3 = "{}"
        Case "CDCA", "CD_CA"
            sql = "SELECT * FROM tbDatosCDCA WHERE idSolicitud = " & idSolicitud
            Set rcd = db.OpenRecordset(sql, dbOpenSnapshot)
            If Not rcd.EOF Then
                jsonSeccion2 = ObtenerJsonDesdeRecordset(rcd, "CDCA")
            Else
                jsonSeccion2 = "{}"
            End If
            jsonSeccion1 = "{}"
            jsonSeccion3 = "{}"
        Case "CDCASUB", "CD_CA_SUB"
            sql = "SELECT * FROM tbDatosCDCASUB WHERE idSolicitud = " & idSolicitud
            Set rcd = db.OpenRecordset(sql, dbOpenSnapshot)
            If Not rcd.EOF Then
                jsonSeccion3 = ObtenerJsonDesdeRecordset(rcd, "CDCASUB")
            Else
                jsonSeccion3 = "{}"
            End If
            jsonSeccion1 = "{}"
            jsonSeccion2 = "{}"
        Case "PC_SUB", "PCSUB"
            sql = "SELECT * FROM tbDatosPCSUB WHERE idSolicitud = " & idSolicitud
            Set rcd = db.OpenRecordset(sql, dbOpenSnapshot)
            If Not rcd.EOF Then
                jsonSeccion1 = ObtenerJsonDesdeRecordset(rcd, "PC_SUB")
            Else
                jsonSeccion1 = "{}"
            End If
            jsonSeccion2 = "{}"
            jsonSeccion3 = "{}"
        Case Else
            jsonSeccion1 = "{}"
            jsonSeccion2 = "{}"
            jsonSeccion3 = "{}"
    End Select
    
    If Not rcd Is Nothing Then rcd.Close
    
    GenerarSnapshotDesdeTabla = "{""datos"":{" & jsonSeccion1 & "},""datos2"":{" & jsonSeccion2 & "},""datos3"":{" & jsonSeccion3 & "}}"
    
    Exit Function
Errores:
    Debug.Print "ERROR GenerarSnapshotDesdeTabla: " & Err.description
    GenerarSnapshotDesdeTabla = "{}"
End Function

Private Function ObtenerJsonDesdeRecordset(ByVal rcd As DAO.Recordset, Optional ByVal tipoSolicitud As String = "") As String
    Dim jsonPairs As String
    Dim fieldName As String
    Dim fieldValue As Variant
    Dim isFirst As Boolean
    Dim fld As DAO.Field
    Dim dictCampos As Object
    
    If rcd Is Nothing Or rcd.EOF Then
        ObtenerJsonDesdeRecordset = ""
        Exit Function
    End If
    
    If tipoSolicitud <> "" Then
        Set dictCampos = ObtenerDictCampos(tipoSolicitud)
    End If
    
    jsonPairs = ""
    isFirst = True
    
    For Each fld In rcd.Fields
        fieldName = fld.name
        
        If dictCampos Is Nothing Then
            fieldValue = Nz(fld.value, "")
            
            If Not isFirst Then
                jsonPairs = jsonPairs & ","
            End If
            
            jsonPairs = jsonPairs & """" & LCase(fieldName) & """:"
            
            If VarType(fieldValue) = vbString Then
                jsonPairs = jsonPairs & """" & EscapeJsonString(CStr(fieldValue)) & """"
            ElseIf IsNumeric(fieldValue) Then
                jsonPairs = jsonPairs & CStr(fieldValue)
            ElseIf VarType(fieldValue) = vbBoolean Then
                jsonPairs = jsonPairs & IIf(fieldValue, "true", "false")
            ElseIf IsNull(fieldValue) Then
                jsonPairs = jsonPairs & "null"
            Else
                jsonPairs = jsonPairs & """" & EscapeJsonString(CStr(fieldValue)) & """"
            End If
            
            isFirst = False
        Else
            If dictCampos.Exists(fieldName) Then
                fieldValue = Nz(fld.value, "")
                
                If Not isFirst Then
                    jsonPairs = jsonPairs & ","
                End If
                
                jsonPairs = jsonPairs & """" & LCase(fieldName) & """:"
                
                If VarType(fieldValue) = vbString Then
                    jsonPairs = jsonPairs & """" & EscapeJsonString(CStr(fieldValue)) & """"
                ElseIf IsNumeric(fieldValue) Then
                    jsonPairs = jsonPairs & CStr(fieldValue)
                ElseIf VarType(fieldValue) = vbBoolean Then
                    jsonPairs = jsonPairs & IIf(fieldValue, "true", "false")
                ElseIf IsNull(fieldValue) Then
                    jsonPairs = jsonPairs & "null"
                Else
                    jsonPairs = jsonPairs & """" & EscapeJsonString(CStr(fieldValue)) & """"
                End If
                
                isFirst = False
            End If
        End If
    Next fld
    
    If Len(jsonPairs) = 0 Then
        jsonPairs = "{}"
    Else
        jsonPairs = "{" & jsonPairs & "}"
    End If
    
    ObtenerJsonDesdeRecordset = jsonPairs
End Function

Private Function ObtenerHashCiclo(ByVal idSolicitud As Long, ByVal ordinal As Long) As String
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String
    
    On Error GoTo Errores
    
    ObtenerHashCiclo = ""
    Set db = getdb()
    
    sql = "SELECT HashDatos FROM tbValidacionRevision WHERE idSolicitud = " & idSolicitud & " AND ordinal = " & ordinal
    Set rcd = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rcd.EOF Then
        ObtenerHashCiclo = Nz(rcd!hashDatos, "")
    End If
    
    rcd.Close
    Exit Function
Errores:
    Debug.Print "ERROR ObtenerHashCiclo: " & Err.description
    ObtenerHashCiclo = ""
End Function

Private Function FormatearJsonPretty(ByVal json As String) As String
    Dim result As String
    Dim i As Long
    Dim indentLevel As Long
    Dim c As String
    Dim inString As Boolean
    Dim prevChar As String
    
    If Len(json) = 0 Then
        FormatearJsonPretty = ""
        Exit Function
    End If
    
    result = ""
    indentLevel = 0
    inString = False
    prevChar = ""
    
    For i = 1 To Len(json)
        c = Mid(json, i, 1)
        
        If c = """" And prevChar <> "\" Then
            inString = Not inString
        End If
        
        If Not inString Then
            Select Case c
                Case "{", "["
                    result = result & c & vbCrLf
                    indentLevel = indentLevel + 1
                    result = result & String(indentLevel * 2, " ")
                Case "}", "]"
                    result = result & vbCrLf
                    indentLevel = indentLevel - 1
                    result = result & String(indentLevel * 2, " ") & c
                Case ","
                    result = result & c & vbCrLf
                    result = result & String(indentLevel * 2, " ")
                Case ":"
                    result = result & ": "
                Case " ", vbTab, vbCr, vbLf, vbCrLf
                Case Else
                    result = result & c
            End Select
        Else
            result = result & c
        End If
        
        prevChar = c
    Next i
    
    FormatearJsonPretty = result
End Function

Public Function GenerarSnapshotGlobal(ByVal idSolicitud As Long, ByVal tipoSolicitud As String, Optional ByRef db As DAO.Database) As String
    Dim jsonSeccion1 As String
    Dim jsonSeccion2 As String
    Dim jsonSeccion3 As String
    
    On Error GoTo Errores
    
    Select Case UCase(tipoSolicitud)
        Case "PC"
            jsonSeccion1 = ObtenerJsonCamposEspecificos("tbDatosPC", idSolicitud, "", tipoSolicitud, db)
            jsonSeccion2 = "{}"
            jsonSeccion3 = "{}"
        Case "CDCA", "CD_CA"
            jsonSeccion1 = "{}"
            jsonSeccion2 = ObtenerJsonCamposEspecificos("tbDatosCDCA", idSolicitud, "", tipoSolicitud, db)
            jsonSeccion3 = "{}"
        Case "CDCASUB", "CD_CA_SUB"
            jsonSeccion1 = "{}"
            jsonSeccion2 = "{}"
            jsonSeccion3 = ObtenerJsonCamposEspecificos("tbDatosCDCASUB", idSolicitud, "", tipoSolicitud, db)
        Case "PC_SUB", "PCSUB"
            jsonSeccion1 = ObtenerJsonCamposEspecificos("tbDatosPCSUB", idSolicitud, "", tipoSolicitud, db)
            jsonSeccion2 = "{}"
            jsonSeccion3 = "{}"
        Case Else
            jsonSeccion1 = "{}"
            jsonSeccion2 = "{}"
            jsonSeccion3 = "{}"
    End Select
    
    Dim jsonDatos As String
    Dim jsonDatos2 As String
    Dim jsonDatos3 As String
    
    jsonDatos = """datos"":{" & jsonSeccion1 & "}"
    jsonDatos2 = """datos2"":{" & jsonSeccion2 & "}"
    jsonDatos3 = """datos3"":{" & jsonSeccion3 & "}"
    
    GenerarSnapshotGlobal = "{" & jsonDatos & "," & jsonDatos2 & "," & jsonDatos3 & "}"
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SnapshotHelper.GenerarSnapshotGlobal"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud & ", tipo=" & tipoSolicitud
    errObj.Raise
End Function

Private Function ObtenerJsonCamposEspecificos(ByVal nombreTabla As String, ByVal idSolicitud As Long, ByVal listaCampos As String, Optional ByVal tipoSolicitud As String = "", Optional ByRef db As DAO.Database) As String
    Dim dbLocal As DAO.Database
    Dim dbCreadoLocal As Boolean
    Dim rcd As DAO.Recordset
    Dim sql As String
    Dim fieldName As String
    Dim fieldValue As Variant
    Dim jsonPairs As String
    Dim isFirst As Boolean
    Dim dictCampos As Object
    Dim fld As DAO.Field
    
    On Error GoTo Errores
    
    If db Is Nothing Then
        Set dbLocal = getdb()
        dbCreadoLocal = True
    Else
        Set dbLocal = db
        dbCreadoLocal = False
    End If
    
    If tipoSolicitud <> "" Then
        Set dictCampos = ObtenerDictCampos(tipoSolicitud)
    End If
    
    sql = "SELECT * FROM " & nombreTabla & " WHERE idSolicitud = " & idSolicitud
    Set rcd = dbLocal.OpenRecordset(sql, dbOpenSnapshot)
    
    If rcd.EOF Then
        ObtenerJsonCamposEspecificos = "{}"
        GoTo LimpiarYSalir
    End If
    
    jsonPairs = ""
    isFirst = True
    
    For Each fld In rcd.Fields
        fieldName = fld.name
        
        If dictCampos Is Nothing Then
            If InStr(1, listaCampos, "|" & fieldName & "|", vbTextCompare) > 0 Then
                fieldValue = Nz(fld.value, "")
                
                If Not isFirst Then
                    jsonPairs = jsonPairs & ","
                End If
                
                jsonPairs = jsonPairs & """" & LCase(fieldName) & """:"
                
                If VarType(fieldValue) = vbString Then
                    jsonPairs = jsonPairs & """" & EscapeJsonString(CStr(fieldValue)) & """"
                ElseIf IsNumeric(fieldValue) Then
                    jsonPairs = jsonPairs & CStr(fieldValue)
                ElseIf VarType(fieldValue) = vbBoolean Then
                    jsonPairs = jsonPairs & IIf(fieldValue, "true", "false")
                ElseIf IsNull(fieldValue) Then
                    jsonPairs = jsonPairs & "null"
                Else
                    jsonPairs = jsonPairs & """" & EscapeJsonString(CStr(fieldValue)) & """"
                End If
                
                isFirst = False
            End If
        Else
            If dictCampos.Exists(fieldName) Then
                fieldValue = Nz(fld.value, "")
                
                If Not isFirst Then
                    jsonPairs = jsonPairs & ","
                End If
                
                jsonPairs = jsonPairs & """" & LCase(fieldName) & """:"
                
                If VarType(fieldValue) = vbString Then
                    jsonPairs = jsonPairs & """" & EscapeJsonString(CStr(fieldValue)) & """"
                ElseIf IsNumeric(fieldValue) Then
                    jsonPairs = jsonPairs & CStr(fieldValue)
                ElseIf VarType(fieldValue) = vbBoolean Then
                    jsonPairs = jsonPairs & IIf(fieldValue, "true", "false")
                ElseIf IsNull(fieldValue) Then
                    jsonPairs = jsonPairs & "null"
                Else
                    jsonPairs = jsonPairs & """" & EscapeJsonString(CStr(fieldValue)) & """"
                End If
                
                isFirst = False
            End If
        End If
    Next fld
    
    If Len(jsonPairs) = 0 Then
        jsonPairs = "{}"
    Else
        jsonPairs = "{" & jsonPairs & "}"
    End If
    
    ObtenerJsonCamposEspecificos = jsonPairs
    
LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    If dbCreadoLocal Then dbLocal.Close
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SnapshotHelper.ObtenerJsonCamposEspecificos"
    errObj.AddToCallStack "nombreTabla=" & nombreTabla & ", idSolicitud=" & idSolicitud
    errObj.Raise
End Function

Private Function EscapeJsonString(ByVal texto As String) As String
    Dim resultado As String
    Dim i As Long
    Dim char As String
    
    resultado = ""
    For i = 1 To Len(texto)
        char = Mid(texto, i, 1)
        Select Case char
            Case """"
                resultado = resultado & "\"""
            Case "\"
                resultado = resultado & "\\"
            Case vbCr
                resultado = resultado & "\n"
            Case vbLf
                resultado = resultado & "\n"
            Case vbTab
                resultado = resultado & "\t"
            Case Else
                If Asc(char) < 32 Then
                Else
                    resultado = resultado & char
                End If
        End Select
    Next i
    
    EscapeJsonString = resultado
End Function

Public Function ObtenerHashDesdeTablaRevision(ByVal idSolicitud As Long, ByVal ordinal As Long) As String
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String
    
    On Error GoTo Errores
    
    Set db = getdb()
    
    sql = "SELECT HashDatos FROM tbValidacionRevision " & _
          "WHERE idSolicitud = " & idSolicitud & " AND ordinal = " & ordinal & ";"
    
    Set rcd = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If rcd.EOF Then
        ObtenerHashDesdeTablaRevision = ""
    Else
        ObtenerHashDesdeTablaRevision = Nz(rcd!hashDatos, "")
    End If
    
LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    If Not db Is Nothing Then db.Close
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SnapshotHelper.ObtenerHashDesdeTablaRevision"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud & ", ordinal=" & ordinal
    errObj.Raise
End Function

Public Sub PruebaDebugLog()
    Call ExportarSnapshotParaDebug(2, "PC", "prueba", 1)
End Sub




