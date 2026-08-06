Attribute VB_Name = "MockExpedienteEdicion"
Option Compare Database
Option Explicit

' Mock class for testing Helper_ExpedienteEdicion.
' Properties usadas por el helper via CallByName (late binding).
' Cobertura: BR-02-01 (campos sucios), BR-02-02 (bloqueado), BR-04-01 (cambio tipo), BR-05-01 (estado calculado).
'
' Reemplaza el uso de CreateObject("Scripting.Dictionary") que no soporta CallByName.
' Creado: 2026-06-16, REFAC-3a slice 2.

' BR-02-01: SoloCamposSucios necesita todos los campos comparables
Public CodExp As String
Public Titulo As String
Public Nemotecnico As String
Public ImporteLicitacion As Variant
Public ESTADO As String

' BR-02-02: ExpedienteBloqueado
Public Bloqueado As Boolean

' BR-04-01: CambiarTipoExpediente
Public EsAM As String
Public EsLote As String
Public EsExpediente As String
Public EsBasado As String
Public Tipo As String
Public IDExpedientePadre As String

' BR-05-01: EstadoCalculadoTexto
Public FechaInicio As Variant  ' puede ser Date o Null
Public FechaFin As Variant
Public EstadoManual As String
Public EstadoActual As String  ' el campo calculado
