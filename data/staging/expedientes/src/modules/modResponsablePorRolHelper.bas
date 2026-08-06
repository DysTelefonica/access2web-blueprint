Attribute VB_Name = "modResponsablePorRolHelper"
Option Compare Database
Option Explicit

' -----------------------------------------------------------------------------
' modResponsablePorRolHelper
'
' Carga los usuarios que ostentan un rol concreto dentro de la aplicacion
' "Expedientes", consultando `TbResponsablesPorRol` (backend
' `Expedientes_datos.accdb`) JOIN `TbUsuariosAplicaciones` (linked desde
' Lanzadera).  Reemplaza las listas hardcodeadas en
' `Form_FormExpedienteGeneral.EstablecerCombos` y `constructor.getUsuariosCalidad`.
'
' Tablas usadas:
'   - TbResponsablesPorRol  (backend).  PK autonumerica; (IDUsuario, Rol)
'                               con indice unico.  Una fila por (IdUsuario, Rol).
'   - TbUsuariosAplicaciones (LINKED desde Lanzadera_Datos.accdb al frontend).
'                               Aporta Id, Nombre, CorreoUsuario, FechaBaja, etc.
'
' Diseno:
'   1. Whitelist CERRADA para el nombre del rol.  Cualquier valor fuera de
'      {Calidad, Seguridad} -> Err.Raise 1000+.  SQL nunca se interpola sin
'      validar.
'   2. `DAO.Database` se inyecta explicitamente.  Si Nothing -> `CurrentDb()`
'      del frontend, que enrutara las linked tables a sus backends
'      respectivos (mismo patron que el resto del repositorio).  Tests
'      inyectan su propio `p_db` (temp .accdb).
'   3. Devuelve `Scripting.Dictionary` keyed por `CStr(IdUsuario)`, value =
'      `Usuario` (mismo patron que `ColUsuariosCalidad` / `ColUsuariosTecnicos`).
'   4. Compatibilidad con expediente "huérfano": si `p_IDResponsableActual`
'      es no-vacio y no figura entre los usuarios devueltos, se añade
'      igualmente (lookup defensivo por Id en `TbUsuariosAplicaciones`) para
'      que un expediente antiguo no quede visualmente vacio.
'   5. Orden determinista: `ORDER BY U.Nombre` (alfabetico por Nombre, case-
'      insensitive en Jet/ACE con el default collation).
'   6. Sin duplicados: la clave del `Scripting.Dictionary` es `CStr(Id)`; un
'      usuario con multiples filas de rol colapsa a una sola entrada.
'   7. Error contract canonico del proyecto: quien falla setea `p_Error` y
'      hace `Err.Raise 1000`; el caller comprueba `If p_Error <> "" Then
'      Err.Raise 1000`.
' -----------------------------------------------------------------------------

' Whitelist cerrada de roles permitidos.
' Mantener sincronizada con los valores validos de `TbResponsablesPorRol.Rol`.
Private Const ROL_CALIDAD    As String = "Calidad"
Private Const ROL_SEGURIDAD  As String = "Seguridad"


' -----------------------------------------------------------------------------
' API publica
' -----------------------------------------------------------------------------

' Devuelve un Dictionary (key=CStr(IdUsuario), value=Usuario) con los usuarios
' ACTIVOS (`FechaBaja Is Null`) que tienen el rol indicado.
'
'   p_Rol                  : "Calidad" | "Seguridad" (whitelist cerrada)
'   p_IDResponsableActual  : ID del responsable actualmente guardado en el
'                            expediente (puede ser ""). Si existe y no esta
'                            en el resultado, se incluye para no perder el
'                            dato y que el combo no quede visualmente vacio.
'   p_db                   : DAO.Database inyectado (tests). Nothing ->
'                            CurrentDb() del frontend.
'   p_Error                : parametro de error estandar del proyecto.
Public Function GetUsuariosPorRol( _
                                    ByVal p_Rol As String, _
                                    ByVal p_IDResponsableActual As String, _
                                    Optional ByRef p_db As DAO.Database, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary

    Dim m_RolValidado As String
    Dim m_db As DAO.Database
    Dim rcd As DAO.Recordset
    Dim m_SQL As String
    Dim m_Usuario As Usuario
    Dim m_Dic As Scripting.Dictionary
    Dim m_ActualKey As String

    ' -- [1] Validacion de la whitelist ---------------------------------
    '  NOTA: este check se hace FUERA de `On Error GoTo errores` a
    '  proposito.  Si lo capturamos localmente con Err.Number = 1000, el
    '  handler lo silencia y el caller pierde la senal.  El caller del
    '  proyecto siempre hace `If p_Error <> "" Then Err.Raise 1000`, asi
    '  que la propagacion directa es lo correcto.
    m_RolValidado = ValidarRol(p_Rol, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    On Error GoTo errores

    ' -- [2] Resolver db ------------------------------------------------
    '  Patron canonico del repo: getdb() apunta al backend correcto
    '  (sandbox en modo testing, produccion en otro caso) y respeta la
    '  configuracion del usuario. CurrentDb() solo sirve cuando la data
    '  esta en el front, lo cual no aplica a TbResponsablesPorRol /
    '  TbUsuariosAplicaciones (viven en el backend).
    If p_db Is Nothing Then
        Set m_db = getdb()
    Else
        Set m_db = p_db
    End If

    ' -- [3] Cargar usuarios activos con el rol pedido ------------------
    Set m_Dic = New Scripting.Dictionary
    m_Dic.CompareMode = TextCompare

    m_SQL = "SELECT U.Id, U.Nombre, U.CorreoUsuario, U.UsuarioRed, " & _
            "       U.Activado, U.FechaBaja " & _
            "FROM TbResponsablesPorRol AS R " & _
            "INNER JOIN TbUsuariosAplicaciones AS U " & _
            "    ON U.Id = R.IDUsuario " & _
            "WHERE R.Rol = '" & Replace(m_RolValidado, "'", "''") & "' " & _
            "  AND U.FechaBaja IS NULL " & _
            "ORDER BY U.Nombre;"

    Set rcd = m_db.OpenRecordset(m_SQL, dbOpenSnapshot)
    With rcd
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                Set m_Usuario = New Usuario
                m_Usuario.CorreoUsuario = Nz(.Fields("CorreoUsuario").value, "")
                m_Usuario.UsuarioRed = Nz(.Fields("UsuarioRed").value, "")
                m_Usuario.Nombre = Nz(.Fields("Nombre").value, "")
                m_Usuario.ID = CStr(Nz(.Fields("Id").value, ""))
                m_ActualKey = m_Usuario.ID
                If Not m_Dic.Exists(m_ActualKey) Then
                    m_Dic.Add m_ActualKey, m_Usuario
                End If
                Set m_Usuario = Nothing
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rcd = Nothing

    ' -- [4] Compatibilidad: incluir ID guardado aunque no tenga el rol -
    If Len(Trim$(p_IDResponsableActual)) > 0 Then
        m_ActualKey = CStr(Trim$(p_IDResponsableActual))
        If Not m_Dic.Exists(m_ActualKey) Then
            Set m_Usuario = LookupUsuarioPorID(CLng(m_ActualKey), m_db)
            If Not m_Usuario Is Nothing Then
                m_Dic.Add m_ActualKey, m_Usuario
            End If
            Set m_Usuario = Nothing
        End If
    End If

    Set GetUsuariosPorRol = m_Dic
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo GetUsuariosPorRol ha devuelto el error n: " & Err.Number & _
                    vbNewLine & "Detalle: " & Err.Description
    End If
End Function


' -----------------------------------------------------------------------------
' Helpers privados
' -----------------------------------------------------------------------------

' Valida `p_Rol` contra la whitelist cerrada.  Setea `p_Error` y devuelve
' "" si no es valido.  Devuelve el nombre validado (identico al literal) si OK.
Private Function ValidarRol( _
                                ByVal p_Rol As String, _
                                Optional ByRef p_Error As String _
                                ) As String
    On Error GoTo errores

    Select Case CStr(p_Rol)
        Case ROL_CALIDAD
            ValidarRol = ROL_CALIDAD
        Case ROL_SEGURIDAD
            ValidarRol = ROL_SEGURIDAD
        Case Else
            p_Error = "ValidarRol: rol no permitido (" & _
                        CStr(p_Rol) & "). Whitelist cerrada."
            ValidarRol = ""
    End Select
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ValidarRol ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


' Lookup defensivo por Id (compatibilidad con responsables "huerfanos").
' Devuelve Nothing si no existe o si FechaBaja es no nulo.
Private Function LookupUsuarioPorID( _
                                        ByVal p_Id As Long, _
                                        ByRef p_db As DAO.Database _
                                        ) As Usuario
    Dim rcd As DAO.Recordset
    Dim m_SQL As String
    Dim m_Usuario As Usuario

    On Error GoTo errores

    m_SQL = "SELECT Id, Nombre, CorreoUsuario, UsuarioRed, Activado, FechaBaja " & _
            "FROM TbUsuariosAplicaciones " & _
            "WHERE Id = " & p_Id & ";"

    Set rcd = p_db.OpenRecordset(m_SQL, dbOpenSnapshot)
    With rcd
        If .EOF Then
            .Close
            Set LookupUsuarioPorID = Nothing
            Exit Function
        End If
        If Not IsNull(.Fields("FechaBaja").value) Then
            .Close
            Set LookupUsuarioPorID = Nothing
            Exit Function
        End If
        Set m_Usuario = New Usuario
        m_Usuario.CorreoUsuario = Nz(.Fields("CorreoUsuario").value, "")
        m_Usuario.UsuarioRed = Nz(.Fields("UsuarioRed").value, "")
        m_Usuario.Nombre = Nz(.Fields("Nombre").value, "")
        m_Usuario.ID = CStr(Nz(.Fields("Id").value, ""))
        Set LookupUsuarioPorID = m_Usuario
        Set m_Usuario = Nothing
        .Close
    End With
    Set rcd = Nothing
    Exit Function

errores:
    Set LookupUsuarioPorID = Nothing
End Function