Attribute VB_Name = "MockEntidad"
Option Compare Database
Option Explicit

' Mock class for testing Helper_EntidadCRUD.
' 2 String properties + 1 Long property.
'
' Usado por tests en Test_Helper_EntidadCRUD.bas (happy path).
' Reemplaza el Scripting.Dictionary que se usaba antes (que no tenia properties
' y por eso CallByName fallaba).
'
' Creado: 2026-06-16, REFAC-2a gap closure.

Public Comercial As String
Public DESCRIPCION As String
Public IDEjemplo As Long
