Attribute VB_Name = "MockOperaciones"
Option Compare Database
Option Explicit

' Mock class for testing Helper_EntidadCRUD.EliminarEntidadGenerico and
' the (future) RegistrarEntidadGenerico.
' 1 Object property (the entity) + 1 m_FailOnEliminar flag for error path tests.
' Counter de llamadas para asserts.
'
' Usado por tests en Test_Helper_EntidadCRUD.bas.
' Cierra el gap 'happy path no testeable' del EliminarEntidadGenerico (3er funcion del helper).
'
' Creado: 2026-06-16, REFAC-2a gap closure.

Public Comercial As Object
Public m_FailOnEliminar As Boolean
Private m_EliminarLlamadas As Long

Public Function Eliminar(Optional ByRef p_Error As String) As String
    m_EliminarLlamadas = m_EliminarLlamadas + 1
    If m_FailOnEliminar Then
        p_Error = "MockOperaciones: Eliminar fallo forzado"
        Eliminar = "ERR: " & p_Error
    Else
        Eliminar = ""
    End If
End Function

Public Property Get EliminarLlamadas() As Long
    EliminarLlamadas = m_EliminarLlamadas
End Property

Public Function Registrar(Optional ByRef p_Error As String) As String
    Registrar = ""
End Function
