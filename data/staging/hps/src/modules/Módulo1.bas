Attribute VB_Name = "Módulo1"
Option Compare Database
Option Explicit


Public Sub ObtenerEstructuraTabla(ByVal NombreTabla As String)
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    Dim fldIdx As DAO.Field
    Dim strPK As String
    
    On Error GoTo ErrorHandler
    
    Set db = CurrentDb
    Set tdf = db.TableDefs(NombreTabla)
    
    Debug.Print "================================================================="
    Debug.Print " ESTRUCTURA DE LA TABLA: " & UCase(NombreTabla)
    Debug.Print "================================================================="
    
    ' 1. LISTAR CAMPOS
    Debug.Print "--- CAMPOS ---"
    Debug.Print "Nombre", "Tipo", "Tamaño", "Requerido"
    Debug.Print "-----------------------------------------------------------------"
    
    For Each fld In tdf.Fields
        Dim tipoStr As String
        tipoStr = GetTipoDato(fld.Type)
        
        ' Formato: [Nombre] - (Tipo) - [Tamaño] - [EsRequerido?]
        Debug.Print "[" & fld.Name & "]", tipoStr, fld.Size, fld.Required
    Next fld
    
    ' 2. LISTAR ÍNDICES (Clave Primaria)
    Debug.Print vbNewLine & "--- ÍNDICES Y CLAVES ---"
    If tdf.Indexes.Count = 0 Then
        Debug.Print "No hay índices definidos."
    Else
        For Each idx In tdf.Indexes
            Dim strCampos As String
            strCampos = ""
            
            ' Obtener campos del índice
            For Each fldIdx In idx.Fields
                strCampos = strCampos & "[" & fldIdx.Name & "] "
            Next fldIdx
            
            Dim infoIdx As String
            infoIdx = idx.Name
            If idx.Primary Then infoIdx = infoIdx & " (PRIMARY KEY)"
            If idx.Unique Then infoIdx = infoIdx & " (UNIQUE)"
            
            Debug.Print infoIdx & ": " & strCampos
        Next idx
    End If
    
    Debug.Print "=================================================================" & vbNewLine
    
    Exit Sub

ErrorHandler:
    If Err.Number = 3265 Then
        Debug.Print "ERROR: La tabla '" & NombreTabla & "' no existe en esta base de datos."
    Else
        Debug.Print "ERROR inesperado: " & Err.Number & " - " & Err.Description
    End If
End Sub

' Función auxiliar para traducir el código numérico del tipo de dato a texto legible
Private Function GetTipoDato(ByVal intType As Integer) As String
    Select Case intType
        Case dbBoolean: GetTipoDato = "Boolean (Si/No)"
        Case dbByte: GetTipoDato = "Byte"
        Case dbInteger: GetTipoDato = "Integer (Corto)"
        Case dbLong: GetTipoDato = "Long (Entero Largo)"
        Case dbCurrency: GetTipoDato = "Currency"
        Case dbSingle: GetTipoDato = "Single"
        Case dbDouble: GetTipoDato = "Double"
        Case dbDate: GetTipoDato = "Date/Time"
        Case dbText: GetTipoDato = "Text (Corto)"
        Case dbLongBinary: GetTipoDato = "OLE Object"
        Case dbMemo: GetTipoDato = "Memo (Largo)"
        Case dbGUID: GetTipoDato = "GUID"
        Case dbDecimal: GetTipoDato = "Decimal"
        Case 101, 102: GetTipoDato = "Adjunto/Calculado" ' Versiones modernas Access
        Case Else: GetTipoDato = "Desconocido (" & intType & ")"
    End Select
End Function
Public Sub CorregirFechasHPS_DesdeOrigen()
    Dim db As DAO.Database
    Dim strSQL As String
    
    On Error GoTo errores
    
    Set db = CurrentDb
    
    ' Desactivamos avisos para que no pida confirmación 4 veces
    DoCmd.SetWarnings False
    
    Debug.Print "Iniciando corrección de fechas..."
    
    ' ----------------------------------------------------------------
    ' 1. ACTUALIZAR NACIONAL
    ' ----------------------------------------------------------------
    ' Buscamos en TbHPS donde TipoHPS = 'Nacional' y actualizamos las columnas HPS_NAC_
    strSQL = "UPDATE TbUsuariosEntidades AS Dest " & _
             "INNER JOIN TbHPS AS Origen ON Dest.ID = Origen.IDUsuario " & _
             "SET Dest.HPS_NAC_F_Concesion = Origen.F_Concesion, " & _
             "    Dest.HPS_NAC_F_Caducidad = Origen.F_Caducidad " & _
             "WHERE Origen.TipoHPS = 'Nacional';"
    db.Execute strSQL, dbFailOnError
    Debug.Print " - Nacionales actualizados: " & db.RecordsAffected

    ' ----------------------------------------------------------------
    ' 2. ACTUALIZAR OTAN
    ' ----------------------------------------------------------------
    ' Buscamos en TbHPS donde TipoHPS = 'OTAN' y actualizamos las columnas HPS_OTAN_
    strSQL = "UPDATE TbUsuariosEntidades AS Dest " & _
             "INNER JOIN TbHPS AS Origen ON Dest.ID = Origen.IDUsuario " & _
             "SET Dest.HPS_OTAN_F_Concesion = Origen.F_Concesion, " & _
             "    Dest.HPS_OTAN_F_Caducidad = Origen.F_Caducidad " & _
             "WHERE Origen.TipoHPS = 'OTAN';"
    db.Execute strSQL, dbFailOnError
    Debug.Print " - OTAN actualizados: " & db.RecordsAffected

    ' ----------------------------------------------------------------
    ' 3. ACTUALIZAR ESA
    ' ----------------------------------------------------------------
    ' Buscamos en TbHPS donde TipoHPS = 'ESA' y actualizamos las columnas HPS_ESA_
    strSQL = "UPDATE TbUsuariosEntidades AS Dest " & _
             "INNER JOIN TbHPS AS Origen ON Dest.ID = Origen.IDUsuario " & _
             "SET Dest.HPS_ESA_F_Concesion = Origen.F_Concesion, " & _
             "    Dest.HPS_ESA_F_Caducidad = Origen.F_Caducidad " & _
             "WHERE Origen.TipoHPS = 'ESA';"
    db.Execute strSQL, dbFailOnError
    Debug.Print " - ESA actualizados: " & db.RecordsAffected

    ' ----------------------------------------------------------------
    ' 4. ACTUALIZAR UE
    ' ----------------------------------------------------------------
    ' Buscamos en TbHPS donde TipoHPS = 'UE' y actualizamos las columnas HPS_UE_
    strSQL = "UPDATE TbUsuariosEntidades AS Dest " & _
             "INNER JOIN TbHPS AS Origen ON Dest.ID = Origen.IDUsuario " & _
             "SET Dest.HPS_UE_F_Concesion = Origen.F_Concesion, " & _
             "    Dest.HPS_UE_F_Caducidad = Origen.F_Caducidad " & _
             "WHERE Origen.TipoHPS = 'UE';"
    db.Execute strSQL, dbFailOnError
    Debug.Print " - UE actualizados: " & db.RecordsAffected

    ' Reactivamos avisos
    DoCmd.SetWarnings True
    
    MsgBox "Proceso finalizado. Las fechas en TbUsuariosEntidades han sido corregidas basándose en TbHPS.", vbInformation

    Exit Sub

errores:
    DoCmd.SetWarnings True
    MsgBox "Error al corregir datos: " & Err.Description, vbCritical
End Sub
Public Sub ActualizarFechaMinimaConcesion()
    Dim db As DAO.Database
    Dim strSQL As String
    
    On Error GoTo errores
    
    Set db = CurrentDb
    DoCmd.SetWarnings False
    
    Debug.Print "Calculando Fechas Mínimas de Concesión..."
    
    ' ----------------------------------------------------------------
    ' 1. LIMPIEZA PREVIA
    ' ----------------------------------------------------------------
    ' Por si acaso se quedó la tabla temporal de una ejecución fallida anterior
    On Error Resume Next
    db.Execute "DROP TABLE Tmp_MinimasHPS;", dbFailOnError
    On Error GoTo errores
    
    ' ----------------------------------------------------------------
    ' 2. CALCULAR MÍNIMOS (GENERAR TABLA TEMPORAL)
    ' ----------------------------------------------------------------
    ' Creamos una tabla Tmp_MinimasHPS con el ID y la fecha más antigua encontrada.
    ' Filtramos (WHERE F_Concesion Is Not Null) para evitar errores con fechas vacías.
    strSQL = "SELECT IDUsuario, Min(F_Concesion) AS MinimaFecha " & _
             "INTO Tmp_MinimasHPS " & _
             "FROM TbHPS " & _
             "WHERE F_Concesion Is Not Null " & _
             "GROUP BY IDUsuario;"
             
    db.Execute strSQL, dbFailOnError
    
    ' ----------------------------------------------------------------
    ' 3. ACTUALIZAR LA TABLA DE ENTIDADES (PLANCHAR EL DATO)
    ' ----------------------------------------------------------------
    ' Actualizamos TbUsuariosEntidades uniendo con la temporal
    strSQL = "UPDATE TbUsuariosEntidades AS Dest " & _
             "INNER JOIN Tmp_MinimasHPS AS Origen ON Dest.ID = Origen.IDUsuario " & _
             "SET Dest.FechaHPSConcesionMinima = Origen.MinimaFecha;"
             
    db.Execute strSQL, dbFailOnError
    Debug.Print " - Fechas mínimas actualizadas: " & db.RecordsAffected
    
    ' ----------------------------------------------------------------
    ' 4. LIMPIEZA FINAL
    ' ----------------------------------------------------------------
    ' Borramos la tabla temporal para no dejar basura
    db.Execute "DROP TABLE Tmp_MinimasHPS;", dbFailOnError
    
    DoCmd.SetWarnings True
    MsgBox "Fecha Mínima de Concesión actualizada correctamente.", vbInformation
    
    Exit Sub

errores:
    DoCmd.SetWarnings True
    MsgBox "Error al actualizar la fecha mínima: " & Err.Description, vbCritical
    ' Intentamos borrar la temporal en caso de error también
    On Error Resume Next
    db.Execute "DROP TABLE Tmp_MinimasHPS;"
End Sub
