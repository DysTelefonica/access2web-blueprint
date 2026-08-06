Attribute VB_Name = "modStub_RebuildNCProyectoListadoCache"
Option Compare Database
Option Explicit

' STUB TEMPORAL — pre-existente issue (no del fix #94).
' NCProyectoGestionListadoHelper.bas:463 llama RebuildNCProyectoListadoCache
' que nunca fue definida. Este stub permite que el helper compile y mi test TDD
' pueda importar el helper. La implementacion real de RebuildNCProyectoListadoCache
' debe abrirse como issue aparte (deuda tecnica pre-existente, no introducida por
' cache-idempotent-warmup-2026-06-17).
'
' Firma deducida de los call sites (helper:463, tests legacy: 2 args).
Public Function RebuildNCProyectoListadoCache(Optional ByVal p_ForceInvalidation As Long = 0, Optional ByRef p_Error As String) As Boolean
    On Error GoTo errores
    p_Error = ""
    ' Stub: retorna True (no-op). La implementacion real debe poblar TbCacheListadoNC
    ' desde el constructor siguiendo A>B>C.
    RebuildNCProyectoListadoCache = True
    Exit Function

errores:
    p_Error = "RebuildNCProyectoListadoCache stub: " & Err.Description
    RebuildNCProyectoListadoCache = False
End Function
