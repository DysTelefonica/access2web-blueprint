Attribute VB_Name = "Helper_ExpedienteAnexos"
Option Compare Database
Option Explicit
' Helper REFAC-3b slice 1: validacion + eliminacion de anexos (documentos adjuntos).
' Stateless, DAO-injectable. Extrae la logica de la eliminacion de anexos
' del form Form_FormExpedienteDocumentacion.ComandoEliminar_Click
' (que delega en ExpedienteOperaciones.EliminarAnexo, p_NombreDocumento).
'
' Cobertura: BR-18-01 (ValidarEliminacionAnexo), BR-18-02 (EliminarAnexo).
'
' Tests en src/modules/Test_Helper_ExpedienteAnexos.bas.
'
' PRUEBA-002 PR-C pendiente: rewire de los 3 handlers ComandoEliminar
' (de FormExpedienteDocumentacion) que hoy llaman a ExpedienteOperaciones.EliminarAnexo.
'
' Spec deviation documentada: EliminarAnexo toma p_PathBase opcional (default "")
' para que los tests puedan apuntar a un sandbox path sin tocar la
' produccion (m_ObjEntorno.URLDirectorioDocumentacion). En produccion se omite
' y se resuelve via el global, manteniendo el comportamiento original.
'
' Tabla: TbExpedientesAnexos (PK IDDocumento Long required, FK IDExpediente
' Long opcional, NombreDocumento Text(255) opcional). La spec tiene typo
' "TbExpedienteAnexos" (singular); se usa el plural real.
'
' Path de archivo: <URLDirectorioDocumentacion>\<Format(IDExpediente,"00000")>\<NombreDocumento>
' Source: getURLDirectorioExpediente en FUNCIONES UTILES.bas:1885.

' Resuelve la DAO: si el caller pasa una, usa esa; si no, getdb() del proyecto.
Private Function ResolveDb( _
    ByVal p_Db As DAO.Database, _
    ByRef p_Error As String) As DAO.Database
    On Error GoTo errores
    If p_Db Is Nothing Then
        Set ResolveDb = getdb()
    Else
        Set ResolveDb = p_Db
    End If
    Exit Function
errores:
    p_Error = "ResolveDb: " & Err.Description
End Function

' Resuelve el path completo de un anexo: base + Format(IDExpediente,"00000") + "\" + filename.
' p_PathBase opcional: si "" usa m_ObjEntorno.URLDirectorioDocumentacion (produccion).
' Si se pasa (tests), se usa ese base. Spec deviation menor para hacer el helper
' deterministico en sandbox.
Private Function ResolverPathAnexo( _
    ByVal p_IDExpediente As Long, _
    ByVal p_NombreDocumento As String, _
    ByVal p_PathBase As String, _
    ByRef p_Error As String) As String
    On Error GoTo errores
    Dim m_Base As String
    If Len(p_PathBase) > 0 Then
        m_Base = p_PathBase
        If Right$(m_Base, 1) <> "\" Then m_Base = m_Base & "\"
    Else
        m_Base = m_ObjEntorno.URLDirectorioDocumentacion
        If Err.Number <> 0 Then
            p_Error = "ResolverPathAnexo: m_ObjEntorno.URLDirectorioDocumentacion fallo: " & Err.Description
            Exit Function
        End If
        If Right$(m_Base, 1) <> "\" Then m_Base = m_Base & "\"
    End If
    ResolverPathAnexo = m_Base & Format(p_IDExpediente, "00000") & "\" & p_NombreDocumento
    Exit Function
errores:
    p_Error = "ResolverPathAnexo: " & Err.Description
End Function

' BR-18-01: Validar que un anexo existe y es elegible para eliminacion.
' DADO un IDDocumento, CUANDO se consulta TbExpedientesAnexos
' ENTONCES retorna True si existe; False + "MSG-ANX-001: anexo invalido" si no.
' DAO-injectable para tests deterministas.
Public Function ValidarEliminacionAnexo( _
    ByVal p_AnexoId As Long, _
    Optional ByVal p_Db As DAO.Database = Nothing, _
    Optional ByRef p_Error As String) As Boolean

    Dim dbUse As DAO.Database
    Dim rs As DAO.Recordset
    Dim m_SQL As String

    On Error GoTo errores

    If p_AnexoId <= 0 Then
        p_Error = "MSG-ANX-001: anexo invalido (ID <= 0)"
        ValidarEliminacionAnexo = False
        Exit Function
    End If

    Set dbUse = ResolveDb(p_Db, p_Error)
    If dbUse Is Nothing Then
        ValidarEliminacionAnexo = False
        Exit Function
    End If

    m_SQL = "SELECT COUNT(*) AS N FROM TbExpedientesAnexos WHERE IDDocumento=" & p_AnexoId
    Set rs = dbUse.OpenRecordset(m_SQL, dbReadOnly)
    If rs!N = 0 Then
        p_Error = "MSG-ANX-001: anexo invalido (no existe en BD)"
        ValidarEliminacionAnexo = False
    Else
        ValidarEliminacionAnexo = True
    End If
    rs.Close

    Exit Function
errores:
    p_Error = "ValidarEliminacionAnexo: " & Err.Description
    ValidarEliminacionAnexo = False
End Function

' BR-18-02: Eliminar un anexo: DB delete (source of truth) + file delete (best-effort).
' DADO un IDDocumento, CUANDO se llama EliminarAnexo
' ENTONCES:
'   - Se valida via ValidarEliminacionAnexo (MSG-ANX-001 si falla)
'   - Se busca el path del archivo (IDExpediente + NombreDocumento)
'   - Se borra el archivo (best-effort, no falla si no existe o esta locked)
'   - Se borra la fila de TbExpedientesAnexos
'   - Retorna True si la fila fue borrada
' Si el file delete falla pero el DB delete tuvo exito, retorna True y deja
' el error del file en p_Error como warning (per spec convention).
' DAO-injectable. p_PathBase opcional para tests (sandbox path).
Public Function EliminarAnexo( _
    ByVal p_Id As Long, _
    Optional ByVal p_Db As DAO.Database = Nothing, _
    Optional ByVal p_PathBase As String = "", _
    Optional ByRef p_Error As String) As Boolean

    Dim dbUse As DAO.Database
    Dim rs As DAO.Recordset
    Dim m_SQL As String
    Dim m_IDExpediente As Long
    Dim m_NombreDocumento As String
    Dim m_PathCompleto As String
    Dim m_FileError As String

    On Error GoTo errores

    ' Pre-validacion
    If Not ValidarEliminacionAnexo(p_Id, p_Db, p_Error) Then
        EliminarAnexo = False
        Exit Function
    End If

    Set dbUse = ResolveDb(p_Db, p_Error)
    If dbUse Is Nothing Then
        EliminarAnexo = False
        Exit Function
    End If

    ' Leer la fila para obtener IDExpediente + NombreDocumento
    m_SQL = "SELECT IDExpediente, NombreDocumento FROM TbExpedientesAnexos WHERE IDDocumento=" & p_Id
    Set rs = dbUse.OpenRecordset(m_SQL, dbReadOnly)
    m_IDExpediente = Nz(rs!IDExpediente, 0)
    m_NombreDocumento = Nz(rs!NombreDocumento, "")
    rs.Close

    ' Resolver path del archivo
    m_PathCompleto = ResolverPathAnexo(m_IDExpediente, m_NombreDocumento, p_PathBase, p_Error)
    If p_Error <> "" Then
        EliminarAnexo = False
        Exit Function
    End If

    ' File delete (best-effort, captura el error por separado)
    On Error Resume Next
    If fso.FileExists(m_PathCompleto) Then
        fso.DeleteFile m_PathCompleto, True
        If Err.Number <> 0 Then m_FileError = Err.Description
    End If
    Err.Clear
    On Error GoTo errores

    ' DB delete (source of truth)
    dbUse.Execute "DELETE * FROM TbExpedientesAnexos WHERE IDDocumento=" & p_Id

    ' Verificar que la fila se borro
    Set rs = dbUse.OpenRecordset("SELECT COUNT(*) AS N FROM TbExpedientesAnexos WHERE IDDocumento=" & p_Id, dbReadOnly)
    If rs!N = 0 Then
        EliminarAnexo = True
        ' Si file delete fallo pero DB ok, dejamos rastro en p_Error (per spec)
        If m_FileError <> "" Then p_Error = "File delete warning: " & m_FileError
    Else
        p_Error = "EliminarAnexo: row not deleted"
        EliminarAnexo = False
    End If
    rs.Close

    Exit Function
errores:
    p_Error = "EliminarAnexo: " & Err.Description
    EliminarAnexo = False
End Function
