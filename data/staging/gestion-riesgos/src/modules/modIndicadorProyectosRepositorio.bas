Attribute VB_Name = "modIndicadorProyectosRepositorio"
Option Compare Database
Option Explicit

' ============================================================
' modIndicadorProyectosRepositorio
' Pure data-access layer for Form_formIndicadorProyectos.
' Pulled out of Form_formIndicadorProyectos.cls so the form stays a
' thin orchestrator (per access-vba-e2e-methodology). No form
' references, no Access controls - directly unit-testable from a
' Test_*.bas when its inputs come from the sandbox.
'
' Public surface:
'   IndicadorProyectos_GetProyectosAbiertos(anio, semestre, error)
'     -> Collection of "ID|Proyecto|NombreProyecto|FechaCierre|FechaRegistroInicial"
'        for proyectos cuyo solape [FechaRegistroInicial, FechaCierre]
'        interseca el rango del semestre pedido (S1/S2/Anual).
'        Empty collection when there are no proyectos.
'        Returns Nothing on input invalido; p_Error populated.
' ============================================================

Public Function IndicadorProyectos_GetProyectosAbiertos( _
                                        ByVal p_Anio As String, _
                                        Optional ByVal p_Semestre As String, _
                                        Optional ByRef p_Error As String _
                                    ) As Collection
    Dim db As DAO.Database
    Dim qd As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim sql As String

    Dim dIni As Date
    Dim dFinExcl As Date

    Dim idp As Long
    Dim Proyecto As String
    Dim NombreProyecto As String
    Dim FechaCierre As Variant
    Dim fechaRegistroInicial As Variant

    On Error GoTo errores
    p_Error = ""

    If Not IsNumeric(p_Anio) Then
        p_Error = "El anyo es obligatorio"
        Err.Raise 1000
    End If

    Select Case Nz(p_Semestre, "")
        Case "1"
            dIni = DateSerial(CLng(p_Anio), 1, 1)
            dFinExcl = DateSerial(CLng(p_Anio), 7, 1)          ' 01/07 (exclusivo)
        Case "2"
            dIni = DateSerial(CLng(p_Anio), 7, 1)
            dFinExcl = DateSerial(CLng(p_Anio) + 1, 1, 1)      ' 01/01 del anyo siguiente (exclusivo)
        Case ""   ' anual
            dIni = DateSerial(CLng(p_Anio), 1, 1)
            dFinExcl = DateSerial(CLng(p_Anio) + 1, 1, 1)
        Case Else
            p_Error = "Solo hay dos semestres en un anyo"
            Err.Raise 1000
    End Select

    ' Proyectos "abiertos"/activos en el periodo por solape de intervalos:
    ' (InicioProyecto < FinPeriodoExcl) AND (FinProyecto Is Null OR FinProyecto >= InicioPeriodo)
    sql = _
        "PARAMETERS pIni DateTime, pFinExcl DateTime;" & vbCrLf & _
        "SELECT IDProyecto, Proyecto, NombreProyecto, FechaCierre, FechaRegistroInicial" & vbCrLf & _
        "FROM TbProyectos" & vbCrLf & _
        "WHERE (FechaRegistroInicial Is Null OR FechaRegistroInicial < [pFinExcl])" & vbCrLf & _
        "  AND (FechaCierre Is Null OR FechaCierre >= [pIni])" & vbCrLf & _
        "ORDER BY IDProyecto;"

    Set db = getdb()
    Set qd = db.CreateQueryDef("", sql)
    qd.Parameters("pIni").Value = dIni
    qd.Parameters("pFinExcl").Value = dFinExcl

    Set rs = qd.OpenRecordset(dbOpenSnapshot)

    If Not (rs.BOF And rs.EOF) Then
        Set IndicadorProyectos_GetProyectosAbiertos = New Collection
        Do While Not rs.EOF
            idp = CLng(Nz(rs!IDProyecto, 0))
            Proyecto = Nz(rs!Proyecto, "")
            NombreProyecto = Nz(rs!NombreProyecto, "")
            FechaCierre = rs!FechaCierre
            fechaRegistroInicial = rs!fechaRegistroInicial

            IndicadorProyectos_GetProyectosAbiertos.Add _
                CStr(idp) & "|" & Proyecto & "|" & NombreProyecto & "|" & Nz(FechaCierre, "") & "|" & Nz(fechaRegistroInicial, "")

            rs.MoveNext
        Loop
    End If

salir:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set qd = Nothing
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "IndicadorProyectos_GetProyectosAbiertos: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    Resume salir
End Function
