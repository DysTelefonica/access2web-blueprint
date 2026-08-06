Attribute VB_Name = "CorreoRepositorio"
' ----- EN: CorreoRepositorio.bas (REEMPLAZAR MÓDULO COMPLETO) -----
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: CorreoRepositorio.bas (Refactorizado con CondorError)
' RESPONSABILIDAD: Acceso a datos para la entidad Correo.
' ==========================================================================

Public Function getSiguienteIDCorreo(ByRef db As DAO.Database) As Long
    Dim rcd As DAO.Recordset
    Dim maxID As Long
    
    On Error GoTo Errores
    maxID = 0
    
    Set rcd = db.OpenRecordset("SELECT Max(IDCorreo) AS MaxID FROM tbCorreosEnviados")
    
    If Not rcd.EOF Then
        maxID = Nz(rcd!maxID, 0)
    End If
    
    getSiguienteIDCorreo = maxID + 1
    
LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Function
    
Errores:
    
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "CorreoRepositorio.getSiguienteIDCorreo"
    errObj.Raise
End Function

'----- EN: CorreoRepositorio.bas (REEMPLAZAR SUBRUTINA) -----

Public Sub GuardarCorreo(ByRef objCorreo As Correo, Optional ByRef db As DAO.Database = Nothing)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    
    On Error GoTo Errores
    ' Fallback: usa BD externa de correo (NO la principal de CONDOR).
    ' IMPORTANTE: esta escritura NO es atómica con la transacción principal de CONDOR.
    ' Si se pasa db (transacción explícita), se comparte el mismo destino.
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    Set rcd = dbConexion.OpenRecordset("tbCorreosEnviados", dbOpenDynaset)
    rcd.AddNew
    
    ' --- REFACTORIZADO ---
    ' Se elimina el bucle manual. La nueva función es más segura.
    ' NOTA: La lógica para omitir 'FechaEnvio' se mantiene si es necesaria,
    ' pero una mejor práctica sería no incluirla en ColCampos si no se debe escribir.
    ' Por ahora, para mantener compatibilidad, hacemos una adaptación.
    Dim tempCol As New Collection
    Dim campo As Variant
    
    ' Creamos una colección temporal sin los campos que no se deben escribir
    For Each campo In objCorreo.ColCampos
        If CStr(campo) <> "FechaEnvio" Then
            tempCol.Add CStr(campo)
        End If
    Next campo
    
    ' Usamos la colección temporal para el relleno
    For Each campo In tempCol
        'Debug.Print campo
        
        rcd.Fields(CStr(campo)).value = objCorreo.getPropiedad(CStr(campo))
    Next campo
    ' --- FIN REFACTORIZADO ---
    
    rcd.Update

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
    
Errores:
    
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "CorreoRepositorio.GuardarCorreo"
    ' Enriquecemos el error si es posible
    If Not campo Is Nothing Then
        errObj.AddToCallStack "Error probable en el campo: " & CStr(campo)
    End If
    errObj.Raise
End Sub
