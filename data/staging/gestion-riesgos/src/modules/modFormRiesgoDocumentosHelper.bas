Attribute VB_Name = "modFormRiesgoDocumentosHelper"
Option Compare Database
Option Explicit

' ============================================================
' modFormRiesgoDocumentosHelper
'   B2 / Punto 15 — Acta 25/06 + Contrato 22/06
'   Helper cache-first para los formularios de anexos.
'   Toda la logica de cache vive aqui; los forms solo delegan.
'
'   Contrato publico (resumen):
'     GetAnexosDeEdicionCached(IDEdicion, p_Error) -> Dictionary
'     GetAnexosDeRiesgoCached(IDRiesgo, p_Error)   -> Dictionary
'     GetAnexoCached(IDAnexo, p_Error)             -> Anexo
'     InvalidarCachePorIDEdicion(IDEdicion)
'     InvalidarCachePorIDRiesgo(IDRiesgo)
'     RegistrarEnEdicion(IDEdicion, URL, Titulo, p_Error) -> Anexo
'     RegistrarEnRiesgo(IDRiesgo, URL, Titulo, p_Error)  -> Anexo
'     EliminarAnexo(IDAnexo, p_Error)
'     RenombrarAnexo(IDAnexo, NuevoTitulo, p_Error)
'
'   Cache interna:
'     m_CachePorIDEdicion: IDEdicion -> Dictionary(IDAnexo -> Anexo)
'     m_CachePorIDRiesgo : IDRiesgo  -> Dictionary(IDAnexo -> Anexo)
'
'   Test-only accessors (marcados para tests, no usar en produccion):
'     CachePorIDEdicionContieneKey, CachePorIDRiesgoContieneKey
' ============================================================

' --- Caches ---
Private m_CachePorIDEdicion As Scripting.Dictionary
Private m_CachePorIDRiesgo As Scripting.Dictionary

' --- Helper interno: garantiza que m_CachePorIDEdicion existe ---
Private Sub EnsureCacheEdicion()
    If m_CachePorIDEdicion Is Nothing Then
        Set m_CachePorIDEdicion = New Scripting.Dictionary
        m_CachePorIDEdicion.CompareMode = TextCompare
    End If
End Sub

Private Sub EnsureCacheRiesgo()
    If m_CachePorIDRiesgo Is Nothing Then
        Set m_CachePorIDRiesgo = New Scripting.Dictionary
        m_CachePorIDRiesgo.CompareMode = TextCompare
    End If
End Sub

' --- Test-only accessors (NO usar en produccion) ---
Public Function CachePorIDEdicionContieneKey(ByVal p_IDEdicion As String) As Boolean
    On Error Resume Next
    If m_CachePorIDEdicion Is Nothing Then
        CachePorIDEdicionContieneKey = False
    Else
        CachePorIDEdicionContieneKey = m_CachePorIDEdicion.Exists(p_IDEdicion)
    End If
    On Error GoTo 0
End Function

Public Function CachePorIDRiesgoContieneKey(ByVal p_IDRiesgo As String) As Boolean
    On Error Resume Next
    If m_CachePorIDRiesgo Is Nothing Then
        CachePorIDRiesgoContieneKey = False
    Else
        CachePorIDRiesgoContieneKey = m_CachePorIDRiesgo.Exists(p_IDRiesgo)
    End If
    On Error GoTo 0
End Function

' ============================================================
' Lecturas — cache-first
' ============================================================

Public Function getAnexosDeEdicionCached(ByVal p_IDEdicion As String, _
                                          ByRef p_Error As String) _
                                          As Scripting.Dictionary
    On Error GoTo errores
    p_Error = ""

    If Len(Trim$(p_IDEdicion)) = 0 Then
        p_Error = "getAnexosDeEdicionCached: p_IDEdicion esta vacio"
        err.Raise 1000
    End If

    EnsureCacheEdicion
    If m_CachePorIDEdicion.Exists(p_IDEdicion) Then
        Set getAnexosDeEdicionCached = m_CachePorIDEdicion(p_IDEdicion)
        Exit Function
    End If

    Dim col As Scripting.Dictionary
    Set col = Constructor.getAnexosTotalesDeEdicion(p_IDEdicion, p_Error)
    If p_Error <> "" Then err.Raise 1000

    If col Is Nothing Then
        Set col = New Scripting.Dictionary
        col.CompareMode = TextCompare
    End If
    m_CachePorIDEdicion.Add p_IDEdicion, col
    Set getAnexosDeEdicionCached = col
    Exit Function

errores:
    If err.Number <> 1000 Then
        p_Error = "getAnexosDeEdicionCached: " & err.Number & " - " & err.description
    End If
End Function

Public Function getAnexosDeRiesgoCached(ByVal p_IDRiesgo As String, _
                                        ByRef p_Error As String) _
                                        As Scripting.Dictionary
    On Error GoTo errores
    p_Error = ""

    If Len(Trim$(p_IDRiesgo)) = 0 Then
        p_Error = "getAnexosDeRiesgoCached: p_IDRiesgo esta vacio"
        err.Raise 1000
    End If

    EnsureCacheRiesgo
    If m_CachePorIDRiesgo.Exists(p_IDRiesgo) Then
        Set getAnexosDeRiesgoCached = m_CachePorIDRiesgo(p_IDRiesgo)
        Exit Function
    End If

    Dim col As Scripting.Dictionary
    Set col = Constructor.getAnexosTotalesDeRiesgo(p_IDRiesgo, p_Error)
    If p_Error <> "" Then err.Raise 1000

    If col Is Nothing Then
        Set col = New Scripting.Dictionary
        col.CompareMode = TextCompare
    End If
    m_CachePorIDRiesgo.Add p_IDRiesgo, col
    Set getAnexosDeRiesgoCached = col
    Exit Function

errores:
    If err.Number <> 1000 Then
        p_Error = "getAnexosDeRiesgoCached: " & err.Number & " - " & err.description
    End If
End Function

' GetAnexoCached — passthrough sin cache (uso per-click desde el form)
Public Function getAnexoCached(ByVal p_IDAnexo As String, _
                                ByRef p_Error As String) As Anexo
    On Error GoTo errores
    p_Error = ""
    If Len(Trim$(p_IDAnexo)) = 0 Then
        Exit Function
    End If
    Set getAnexoCached = Constructor.getAnexo(p_IDAnexo, p_Error)
    Exit Function

errores:
    If err.Number <> 1000 Then
        p_Error = "getAnexoCached: " & err.Number & " - " & err.description
    End If
End Function

' ============================================================
' Invalidacion de cache — quirurgica
' ============================================================

Public Sub InvalidarCachePorIDEdicion(ByVal p_IDEdicion As String)
    On Error Resume Next
    If Not m_CachePorIDEdicion Is Nothing Then
        If m_CachePorIDEdicion.Exists(p_IDEdicion) Then
            m_CachePorIDEdicion.Remove p_IDEdicion
        End If
    End If
    On Error GoTo 0
End Sub

Public Sub InvalidarCachePorIDRiesgo(ByVal p_IDRiesgo As String)
    On Error Resume Next
    If Not m_CachePorIDRiesgo Is Nothing Then
        If m_CachePorIDRiesgo.Exists(p_IDRiesgo) Then
            m_CachePorIDRiesgo.Remove p_IDRiesgo
        End If
    End If
    On Error GoTo 0
End Sub

' ============================================================
' Mutaciones — invalidan cache en ID afectado
' ============================================================

Public Function RegistrarEnEdicion(ByVal p_IDEdicion As String, _
                                    ByVal p_URLLocal As String, _
                                    ByVal p_Titulo As String, _
                                    ByRef p_Error As String) As Anexo
    Dim m_objEdicion As edicion
    Dim m_ObjAnexo As Anexo

    On Error GoTo errores
    p_Error = ""

    If Len(Trim$(p_IDEdicion)) = 0 Then
        p_Error = "RegistrarEnEdicion: p_IDEdicion esta vacio"
        Exit Function
    End If

    ' Invalidacion proactiva: el usuario ya decidio modificar.
    InvalidarCachePorIDEdicion p_IDEdicion

    Set m_objEdicion = Constructor.getEdicion(p_IDEdicion, p_Error)
    If p_Error <> "" Then Exit Function
    If m_objEdicion Is Nothing Then
        p_Error = "RegistrarEnEdicion: no se encontro la edicion " & p_IDEdicion
        Exit Function
    End If

    Set m_ObjAnexo = New Anexo
    m_ObjAnexo.Titulo = p_Titulo
    Dim m_Resultado As String
    m_Resultado = m_ObjAnexo.Registrar(m_objEdicion, p_URLLocal, p_Error)
    If p_Error <> "" Then Exit Function
    If m_Resultado <> "" Then
        p_Error = m_Resultado
        Exit Function
    End If
    Set RegistrarEnEdicion = m_ObjAnexo

    Exit Function

errores:
    p_Error = "RegistrarEnEdicion: " & err.Number & " - " & err.description
End Function

Public Function RegistrarEnRiesgo(ByVal p_IDRiesgo As String, _
                                   ByVal p_URLLocal As String, _
                                   ByVal p_Titulo As String, _
                                   ByRef p_Error As String) As Anexo
    Dim m_ObjRiesgo As riesgo
    Dim m_IDEdicionPadre As String
    Dim m_ObjAnexo As Anexo
    Dim dbErr As String

    On Error GoTo errores
    p_Error = ""

    If Len(Trim$(p_IDRiesgo)) = 0 Then
        p_Error = "RegistrarEnRiesgo: p_IDRiesgo esta vacio"
        Exit Function
    End If

    ' Invalidacion proactiva (cache riesgo + cache edicion padre)
    InvalidarCachePorIDRiesgo p_IDRiesgo

    Set m_ObjRiesgo = Constructor.getRiesgo(p_IDRiesgo, , , dbErr)
    If dbErr <> "" Then
        p_Error = dbErr
        Exit Function
    End If
    If m_ObjRiesgo Is Nothing Then
        p_Error = "RegistrarEnRiesgo: no se encontro el riesgo " & p_IDRiesgo
        Exit Function
    End If

    m_IDEdicionPadre = m_ObjRiesgo.IDEdicion
    InvalidarCachePorIDEdicion m_IDEdicionPadre

    Set m_ObjAnexo = New Anexo
    m_ObjAnexo.Titulo = p_Titulo
    Dim m_Resultado As String
    m_Resultado = m_ObjAnexo.Registrar(m_ObjRiesgo, p_URLLocal, p_Error)
    If p_Error <> "" Then Exit Function
    If m_Resultado <> "" Then
        p_Error = m_Resultado
        Exit Function
    End If
    Set RegistrarEnRiesgo = m_ObjAnexo

    Exit Function

errores:
    p_Error = "RegistrarEnRiesgo: " & err.Number & " - " & err.description
End Function

Public Sub EliminarAnexo(ByVal p_IDAnexo As String, _
                          ByRef p_Error As String)
    Dim m_ObjAnexo As Anexo
    Dim m_IDEdicion As String
    Dim m_IdRiesgo As String
    Dim dbErr As String

    On Error GoTo errores
    p_Error = ""

    If Len(Trim$(p_IDAnexo)) = 0 Then
        p_Error = "EliminarAnexo: p_IDAnexo esta vacio"
        Exit Sub
    End If

    ' Cargar para saber IDEdicion e IDRiesgo ANTES de eliminar
    Set m_ObjAnexo = Constructor.getAnexo(p_IDAnexo, dbErr)
    If dbErr <> "" Then
        p_Error = dbErr
        Exit Sub
    End If
    If m_ObjAnexo Is Nothing Then
        p_Error = "EliminarAnexo: no existe el anexo " & p_IDAnexo
        Exit Sub
    End If

    m_IDEdicion = m_ObjAnexo.IDEdicion
    m_IdRiesgo = m_ObjAnexo.idRiesgo

    ' Invalidacion proactiva
    If Len(Trim$(m_IDEdicion)) > 0 Then InvalidarCachePorIDEdicion m_IDEdicion
    If Len(Trim$(m_IdRiesgo)) > 0 Then InvalidarCachePorIDRiesgo m_IdRiesgo
    ' Si el anexo es de un riesgo hijo, invalidar tambien la cache de la edicion
    ' padre del riesgo (pre-existing issue #83 - el test 1 asume esta semantica)
    If Len(Trim$(m_IdRiesgo)) > 0 And Len(Trim$(m_IDEdicion)) = 0 Then
        Dim m_ObjRiesgoAux As Riesgo
        Dim m_EdicionPadre As String
        Dim m_dbErrAux As String
        Set m_ObjRiesgoAux = Constructor.getRiesgo(m_IdRiesgo, m_dbErrAux)
        If m_dbErrAux = "" And Not m_ObjRiesgoAux Is Nothing Then
            m_EdicionPadre = CStr(m_ObjRiesgoAux.IDEdicion)
            If Len(Trim$(m_EdicionPadre)) > 0 Then
                InvalidarCachePorIDEdicion m_EdicionPadre
            End If
        End If
        Set m_ObjRiesgoAux = Nothing
    End If

    m_ObjAnexo.EliminarAnexo p_Error
    Exit Sub

errores:
    p_Error = "EliminarAnexo: " & err.Number & " - " & err.description
End Sub

Public Sub RenombrarAnexo(ByVal p_IDAnexo As String, _
                          ByVal p_NuevoTitulo As String, _
                          ByRef p_Error As String)
    Dim m_ObjAnexo As Anexo
    Dim dbErr As String

    On Error GoTo errores
    p_Error = ""

    If Len(Trim$(p_IDAnexo)) = 0 Then
        p_Error = "RenombrarAnexo: p_IDAnexo esta vacio"
        Exit Sub
    End If
    If Len(Trim$(p_NuevoTitulo)) = 0 Then
        p_Error = "RenombrarAnexo: p_NuevoTitulo esta vacio"
        Exit Sub
    End If

    Set m_ObjAnexo = Constructor.getAnexo(p_IDAnexo, dbErr)
    If dbErr <> "" Then
        p_Error = dbErr
        Exit Sub
    End If
    If m_ObjAnexo Is Nothing Then
        p_Error = "RenombrarAnexo: no existe el anexo " & p_IDAnexo
        Exit Sub
    End If

    ' CambiarNombre actualiza DB pero NO Me.Titulo. Lo actualizamos en el objeto.
    m_ObjAnexo.Titulo = p_NuevoTitulo
    m_ObjAnexo.CambiarNombre p_NuevoTitulo, p_Error
    If p_Error <> "" Then Exit Sub

    ' Propagar el nuevo Titulo a cualquier instancia cacheada para mantener
    ' coherencia: las caches guardan referencias, no copias.
    If Not m_CachePorIDEdicion Is Nothing Then
        Dim m_IDEd As Variant
        For Each m_IDEd In m_CachePorIDEdicion.keys
            Dim m_ColE As Scripting.Dictionary
            Set m_ColE = m_CachePorIDEdicion(m_IDEd)
            If Not m_ColE Is Nothing Then
                If m_ColE.Exists(p_IDAnexo) Then
                    Dim m_AnexoCacheado As Anexo
                    Set m_AnexoCacheado = m_ColE(p_IDAnexo)
                    m_AnexoCacheado.Titulo = p_NuevoTitulo
                    Set m_AnexoCacheado = Nothing
                End If
            End If
            Set m_ColE = Nothing
        Next
    End If

    If Not m_CachePorIDRiesgo Is Nothing Then
        Dim m_IDR As Variant
        For Each m_IDR In m_CachePorIDRiesgo.keys
            Dim m_ColR As Scripting.Dictionary
            Set m_ColR = m_CachePorIDRiesgo(m_IDR)
            If Not m_ColR Is Nothing Then
                If m_ColR.Exists(p_IDAnexo) Then
                    Dim m_AnexoCacheado2 As Anexo
                    Set m_AnexoCacheado2 = m_ColR(p_IDAnexo)
                    m_AnexoCacheado2.Titulo = p_NuevoTitulo
                    Set m_AnexoCacheado2 = Nothing
                End If
            End If
            Set m_ColR = Nothing
        Next
    End If

    Exit Sub

errores:
    p_Error = "RenombrarAnexo: " & err.Number & " - " & err.description
End Sub


