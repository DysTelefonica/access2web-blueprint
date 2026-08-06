VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstadoServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO DE CLASE: EstadoServicio.cls (REFACTORIZADO CON CondorError)
' ==========================================================================

Public Function getEstadoPorID(ByVal id As Long) As estado
    ' Responsabilidad: Obtiene un único objeto Estado, idealmente desde la caché.
    On Error GoTo Errores
    
    If m_ObjEntorno.estados Is Nothing Then
        Err.Raise 513, , "La colección de estados no está cargada en el Entorno."
    End If
    
    If m_ObjEntorno.estados.Exists(CStr(id)) Then
        Set getEstadoPorID = m_ObjEntorno.estados(CStr(id))
    Else
        Set getEstadoPorID = Nothing ' No se encontró
    End If
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "EstadoServicio.getEstadoPorID"
    errObj.Raise
End Function
