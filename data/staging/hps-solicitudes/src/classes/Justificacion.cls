VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Justificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Clase: Justificacion.cls
Option Compare Database
Option Explicit

' Variables privadas (prefijo m_)
Private m_idjustificacion As Long
Private m_titulo As String
Private m_descripcion As String
Private m_activa As Boolean

' --- Propiedades (Deben coincidir con los campos de la tabla) ---

Public Property Get IDJustificacion() As Long
    IDJustificacion = m_idjustificacion
End Property
Public Property Let IDJustificacion(ByVal value As Long)
    m_idjustificacion = value
End Property

Public Property Get Titulo() As String
    Titulo = m_titulo
End Property
Public Property Let Titulo(ByVal value As String)
    m_titulo = value
End Property

Public Property Get Descripcion() As String
    Descripcion = m_descripcion
End Property
Public Property Let Descripcion(ByVal value As String)
    m_descripcion = value
End Property

Public Property Get activa() As Boolean
    activa = m_activa
End Property
Public Property Let activa(ByVal value As Boolean)
    m_activa = value
End Property

' --- Métodos standard de arquitectura ---
Public Function getPropiedad(nombrePropiedad As String) As Variant
    getPropiedad = CallByName(Me, nombrePropiedad, VbGet)
End Function

Public Sub setPropiedad(nombrePropiedad As String, valor As Variant)
    CallByName Me, nombrePropiedad, VbLet, valor
End Sub
