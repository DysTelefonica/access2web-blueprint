VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MapeoServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: MapeoServicio.cls (REFACTORIZADO - Sin instancia de repositorio)
' ==========================================================================

Public Function getMapeoPC() As Scripting.Dictionary
    On Error GoTo Errores
    ' --- CAMBIO: Llamada directa al módulo .bas ---
    Set getMapeoPC = MapeoRepositorio.getMapeoParaPlantilla("PC")
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "MapeoServicio.getMapeoPC"
    errObj.Raise
End Function

Public Function getMapeoCDCA() As Scripting.Dictionary
    On Error GoTo Errores
    ' --- CAMBIO: Llamada directa al módulo .bas ---
    Set getMapeoCDCA = MapeoRepositorio.getMapeoParaPlantilla("CDCA")
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "MapeoServicio.getMapeoCDCA"
    errObj.Raise
End Function

Public Function getMapeoCDCASUB() As Scripting.Dictionary
    On Error GoTo Errores
    ' --- CAMBIO: Llamada directa al módulo .bas ---
    Set getMapeoCDCASUB = MapeoRepositorio.getMapeoParaPlantilla("CDCASUB")
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "MapeoServicio.getMapeoCDCASUB"
    errObj.Raise
End Function

Public Function getMapeoPCSUB() As Scripting.Dictionary
    On Error GoTo Errores
    ' --- PCSUB: Gemelo de PC para subcontratistas ---
    Set getMapeoPCSUB = MapeoRepositorio.getMapeoParaPlantilla("PCSUB")
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "MapeoServicio.getMapeoPCSUB"
    errObj.Raise
End Function

Public Function SembrarBookmarksDesdePlantilla(ByVal nombrePlantilla As String, ByVal rutaPlantilla As String) As Long
    Dim wd As Word.Application, doc As Word.Document
    Dim db As DAO.Database, rcd As DAO.Recordset
    Dim b As Word.Bookmark
    On Error GoTo Errores
    
    SembrarBookmarksDesdePlantilla = 0
    If Not fso.FileExists(rutaPlantilla) Then Err.Raise 513, , "No existe la plantilla: " & rutaPlantilla

    Set db = getdb()
    Set rcd = db.OpenRecordset("tbMapeoCampos", dbOpenDynaset)
    Set wd = New Word.Application: wd.Visible = False
    Set doc = wd.Documents.Open(rutaPlantilla)

    For Each b In doc.Bookmarks
        If Len(Trim$(b.name)) > 0 Then
            rcd.AddNew
            rcd!nombrePlantilla = nombrePlantilla
            rcd!nombreCampoTabla = ""
            rcd!nombreCampoWord = b.name
            rcd!valorAsociado = ""
            rcd.Update
            SembrarBookmarksDesdePlantilla = SembrarBookmarksDesdePlantilla + 1
        End If
    Next

LimpiarYSalir:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close False
    If Not wd Is Nothing Then wd.Quit
    Set rcd = Nothing: Set db = Nothing: Set doc = Nothing: Set wd = Nothing
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "MapeoServicio.SembrarBookmarksDesdePlantilla"
    errObj.Raise
End Function
