Attribute VB_Name = "RepositorioComun"
' Módulo: RepositorioComun
Option Compare Database
Option Explicit

' --- FUNCIÓN HELPER PRIVADA ---
' Centraliza el acceso a la base de datos.
' Usamos getDB() como indicas para acceder al Back-End.
Private Function DB() As DAO.Database
    Set DB = getdb()
End Function

' --- FUNCIONES PÚBLICAS ---

' Ejecuta una consulta SELECT y devuelve una Colección
Public Function HidratarColeccionDesdeSQL(sql As String, className As String) As Collection
    Dim rs As DAO.Recordset
    Dim col As New Collection
    Dim obj As Object
    
    ' CAMBIO: Usamos DB() en lugar de CurrentDb
    Set rs = DB.OpenRecordset(sql, dbOpenSnapshot)
    
    Do While Not rs.EOF
        Set obj = Factoria.CreateEntity(className)
        LlenarObjetoDesdeRecordset obj, rs
        col.Add obj
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Set HidratarColeccionDesdeSQL = col
End Function

' Ejecuta una consulta SELECT y devuelve UN objeto
Public Function HidratarEntidadDesdeSQL(sql As String, className As String) As Object
    Dim rs As DAO.Recordset
    Dim obj As Object
    
    ' CAMBIO: Usamos DB() en lugar de CurrentDb
    Set rs = DB.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        Set obj = Factoria.CreateEntity(className)
        LlenarObjetoDesdeRecordset obj, rs
        Set HidratarEntidadDesdeSQL = obj
    End If
    
    rs.Close
    Set rs = Nothing
End Function

' Ejecuta consultas de acción (INSERT, UPDATE, DELETE)
Public Sub EjecutarSQL(sql As String)
    ' CAMBIO: Usamos DB() en lugar de CurrentDb
    DB.Execute sql, dbFailOnError
End Sub

' --- FUNCIONES PRIVADAS DE AYUDA ---

Private Sub LlenarObjetoDesdeRecordset(obj As Object, rs As DAO.Recordset)
    Dim fld As DAO.Field
    For Each fld In rs.Fields
        On Error Resume Next
        CallByName obj, fld.Name, VbLet, fld.value
        On Error GoTo 0
    Next fld
End Sub
' --- NUEVO MÉTODO PARA CONSULTAS PARAMETRIZADAS ---
' Permite ejecutar SQL pasando valores sin concatenar cadenas.
' Uso: EjecutarSQLConParametros "INSERT INTO Tbl (A, B) VALUES (p1, p2)", valorA, valorB
Public Sub EjecutarSQLConParametros(sql As String, ParamArray valores() As Variant)
    Dim qdf As DAO.QueryDef
    Dim i As Integer
    
    ' Creamos una QueryDef temporal (sin nombre) apuntando al Back-End
    Set qdf = DB.CreateQueryDef("", sql)
    
    ' Verificamos que el número de parámetros coincida
    ' Nota: qdf.Parameters.Count cuenta cuantos [corchetes] o ? hay en tu SQL
    If qdf.Parameters.Count <> (UBound(valores) + 1) Then
        Err.Raise 51005, "RepositorioComun", "El número de parámetros SQL no coincide con los valores pasados."
    End If
    
    ' Asignamos los valores en orden
    For i = 0 To UBound(valores)
        qdf.Parameters(i).value = valores(i)
    Next i
    
    ' Ejecutamos
    qdf.Execute dbFailOnError
    qdf.Close
    Set qdf = Nothing
End Sub
