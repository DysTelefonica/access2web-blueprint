VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NoConformidadServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Public Function getNoConformidades() As Scripting.Dictionary
    On Error GoTo Errores
    
    Set getNoConformidades = NoConformidadRepositorio.getTodas()
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "NoConformidadServicio.getNoConformidades"
    errObj.Raise
End Function
Public Function getNoConformidadesPorExpediente(ByVal idExpediente As Long) As Scripting.Dictionary
    On Error GoTo Errores
    
    If idExpediente <= 0 Then
        ' Si no hay expediente válido, devolvemos una colección vacía
        Set getNoConformidadesPorExpediente = New Scripting.Dictionary
        Exit Function
    End If
    
    Set getNoConformidadesPorExpediente = NoConformidadRepositorio.getPorExpediente(idExpediente)
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "NoConformidadServicio.getNoConformidadesPorExpediente"
    errObj.Raise
End Function

Public Function getNoConformidadPorCodigoCondor(ByVal codigoCondor As String) As NoConformidad
    On Error GoTo Errores
    Set getNoConformidadPorCodigoCondor = NoConformidadRepositorio.getPorCodigoCondor(codigoCondor)
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "NoConformidadServicio.getNoConformidadPorCodigoCondor"
    errObj.Raise
End Function

Public Function getNoConformidadPorID(ByVal idNC As Long) As NoConformidad
    On Error GoTo Errores
    Set getNoConformidadPorID = NoConformidadRepositorio.getPorID(idNC)
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "NoConformidadServicio.getNoConformidadPorID"
    errObj.Raise
End Function

Public Function estaRegistradaEnBaseDatosExternaa(ByVal idNC As Long) As Boolean
    On Error GoTo Errores
    estaRegistradaEnBaseDatosExternaa = NoConformidadRepositorio.existeEnBaseDatosExternaa(idNC)
    Exit Function
Errores:
    estaRegistradaEnBaseDatosExternaa = False
End Function

Public Function actualizarCodigoConcesionEnTransaccion(ByVal idNC As Long, ByVal codigoCondor As String, _
                                                     ByRef dbCondor As DAO.Database, ByRef dbNC As DAO.Database) As Boolean
    On Error GoTo Errores
    
    If idNC > 0 And codigoCondor <> "" Then
        actualizarCodigoConcesionEnTransaccion = NoConformidadRepositorio.actualizarCodigoConcesion(idNC, codigoCondor, dbNC)
    Else
        actualizarCodigoConcesionEnTransaccion = True
    End If
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "NoConformidadServicio.actualizarCodigoConcesionEnTransaccion"
    errObj.Raise
End Function

Public Function limpiarCodigoConcesionEnTransaccion(ByVal idNC As Long, ByRef dbCondor As DAO.Database, _
                                                   ByRef dbNC As DAO.Database) As Boolean
    On Error GoTo Errores
    
    If idNC > 0 Then
        limpiarCodigoConcesionEnTransaccion = NoConformidadRepositorio.limpiarCodigoConcesion(idNC, dbNC)
    Else
        limpiarCodigoConcesionEnTransaccion = True
    End If
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "NoConformidadServicio.limpiarCodigoConcesionEnTransaccion"
    errObj.Raise
End Function


