Attribute VB_Name = "MockExpediente"
Option Compare Database
Option Explicit

' Mock class for testing Helper_ExpedienteAlta.
' 4 properties (CodExp, EsLote, IDExpedientePadre, ImporteLicitacion)
' leidas por los helpers via CallByName.
'
' Reemplaza el uso de CreateObject("Scripting.Dictionary") que no soporta
' CallByName para property access.
'
' Creado: 2026-06-16, REFAC-3a.

Public CodExp As String
Public EsLote As String
Public IDExpedientePadre As String
Public ImporteLicitacion As Variant
