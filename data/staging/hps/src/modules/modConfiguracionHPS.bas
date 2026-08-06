Attribute VB_Name = "modConfiguracionHPS"
Option Compare Database
Option Explicit

' ============================================================
' modConfiguracionHPS -- preconfigura TbConfiguracionHPS y
' TbVinculosTablas con un perfil (Local o Produccion).
'
' Pensado para llamarse desde la ventana Inmediato del VBE o
' desde atomos de test antes del refactor de Entorno.cls:
'
'     ConfigurarTablaConfiguracion_Perfil "Local"
'     ConfigurarTablaConfiguracion_Perfil "Produccion"
'
' Garantias:
'   - Idempotente: UPSERT por Clave / TablaOrigen.
'   - Crea las tablas si no existen (CREATE TABLE IF NOT EXISTS).
'   - Registra todas las escrituras con Activo=1, FechaModificacion=Now()
'     y UsuarioModificacion=Environ("Username").
'
' Las SQL de CREATE TABLE se pasan tal cual al motor -- los valores
' contenidos (rutas UNC, paths locales) no traen comillas simples,
' asi que Replace(..., "'", "''") cubre la unica inyeccion posible.
' Si en el futuro aparece un valor con apostrofe, agregar escape
' explicito en UpsertConfig / UpsertVinculo.
' ============================================================

' --- Private Const (project rule: Private Const BEFORE first Public) ---

Private Const k_TableConfig As String = "TbConfiguracionHPS"
Private Const k_TableVinculos As String = "TbVinculosTablas"

' Esquema de TbConfiguracionHPS: clave/valor con sello de auditoria.
' Access DAO usa YESNO (no BIT) y NO soporta DEFAULT Now() en CREATE TABLE;
' Activo y FechaModificacion se setean manualmente en los UPSERT.
Private Const k_SchemaConfigSql As String = _
    "CREATE TABLE TbConfiguracionHPS (" & _
    "Clave TEXT(64) PRIMARY KEY, " & _
    "Valor TEXT(255) NOT NULL, " & _
    "Activo YESNO NOT NULL, " & _
    "FechaModificacion DATETIME NOT NULL, " & _
    "UsuarioModificacion TEXT(64))"

' Esquema de TbVinculosTablas: tabla -> backend .accdb de origen.
Private Const k_SchemaVinculosSql As String = _
    "CREATE TABLE TbVinculosTablas (" & _
    "TablaOrigen TEXT(64) PRIMARY KEY, " & _
    "BackendFile TEXT(64) NOT NULL)"

Private Const k_BackendHPST As String = "HPST.accdb"
Private Const k_BackendExpedientes As String = "Expedientes_datos.accdb"
Private Const k_BackendUnknown As String = "Unknown.accdb"

' --- Private helpers (todos antes del primer Public) ---

' Devuelve True si la tabla ya existe en TableDefs (case-insensitive).
Private Function TableDefExists(ByRef p_Db As DAO.Database, ByVal p_TableName As String) As Boolean
    Dim tdf As DAO.TableDef
    For Each tdf In p_Db.TableDefs
        If StrComp(tdf.Name, p_TableName, vbTextCompare) = 0 Then
            TableDefExists = True
            Exit Function
        End If
    Next tdf
    TableDefExists = False
End Function

' CREATE TABLE IF NOT EXISTS -- si la tabla ya esta, no hace nada.
Private Sub ExecuteCreateTableIfNotExists(ByRef p_Db As DAO.Database, ByVal p_TableName As String, ByVal p_Sql As String)
    If Not TableDefExists(p_Db, p_TableName) Then
        p_Db.Execute p_Sql, dbFailOnError
    End If
End Sub

' UPSERT en TbConfiguracionHPS. Reemplaza comillas simples para evitar
' inyeccion SQL basica (las semillas actuales no traen apostrofes,
' pero dejamos el escape como red de seguridad).
Private Sub UpsertConfig(ByRef p_Db As DAO.Database, ByVal p_Clave As String, ByVal p_Valor As String)
    Dim rs As DAO.Recordset
    Dim sqlCheck As String
    Dim claveSafe As String
    Dim valorSafe As String
    Dim usuario As String

    claveSafe = Replace(p_Clave, "'", "''")
    valorSafe = Replace(p_Valor, "'", "''")
    usuario = Replace(Environ("Username"), "'", "''")

    sqlCheck = "SELECT COUNT(*) AS Cantidad FROM " & k_TableConfig & _
               " WHERE Clave='" & claveSafe & "'"
    Set rs = p_Db.OpenRecordset(sqlCheck)
    If rs.Fields("Cantidad").Value = 0 Then
        p_Db.Execute _
            "INSERT INTO " & k_TableConfig & _
            " (Clave, Valor, Activo, FechaModificacion, UsuarioModificacion)" & _
            " VALUES ('" & claveSafe & "', '" & valorSafe & "', 1, Now(), '" & usuario & "')", _
            dbFailOnError
    Else
        p_Db.Execute _
            "UPDATE " & k_TableConfig & _
            " SET Valor='" & valorSafe & _
            "', Activo=1, FechaModificacion=Now(), UsuarioModificacion='" & usuario & _
            "' WHERE Clave='" & claveSafe & "'", _
            dbFailOnError
    End If
    rs.Close
    Set rs = Nothing
End Sub

' UPSERT en TbVinculosTablas (una fila por tabla vinculada no-MSys).
Private Sub UpsertVinculo(ByRef p_Db As DAO.Database, ByVal p_TablaOrigen As String, ByVal p_BackendFile As String)
    Dim rs As DAO.Recordset
    Dim sqlCheck As String
    Dim tablaSafe As String
    Dim backendSafe As String

    tablaSafe = Replace(p_TablaOrigen, "'", "''")
    backendSafe = Replace(p_BackendFile, "'", "''")

    sqlCheck = "SELECT COUNT(*) AS Cantidad FROM " & k_TableVinculos & _
               " WHERE TablaOrigen='" & tablaSafe & "'"
    Set rs = p_Db.OpenRecordset(sqlCheck)
    If rs.Fields("Cantidad").Value = 0 Then
        p_Db.Execute _
            "INSERT INTO " & k_TableVinculos & _
            " (TablaOrigen, BackendFile) VALUES ('" & tablaSafe & "', '" & backendSafe & "')", _
            dbFailOnError
    Else
        p_Db.Execute _
            "UPDATE " & k_TableVinculos & _
            " SET BackendFile='" & backendSafe & _
            "' WHERE TablaOrigen='" & tablaSafe & "'", _
            dbFailOnError
    End If
    rs.Close
    Set rs = Nothing
End Sub

' Heuristica: del Connect string deduce el .accdb de origen por FILENAME,
' no por substring del path. Esto evita falsos positivos cuando el path
' contiene "HPS\" pero el .accdb real es otro (ej. \\server\HPS\
' Expedientes_datos.accdb se confundiria con HPST.accdb y la heuristica
' lo mandaba al backend incorrecto).
'   - Contiene "HPST.accdb"                                          -> HPST.accdb
'   - Contiene "Expedientes_datos.accdb"                             -> Expedientes_datos.accdb
'   - Cualquier otro caso                                            -> Unknown.accdb (con Debug.Print)
Private Function ResolveBackendFileName(ByVal p_Connect As String) As String
    If InStr(1, p_Connect, k_BackendHPST, vbTextCompare) > 0 Then
        ResolveBackendFileName = k_BackendHPST
    ElseIf InStr(1, p_Connect, k_BackendExpedientes, vbTextCompare) > 0 Then
        ResolveBackendFileName = k_BackendExpedientes
    Else
        Debug.Print "modConfiguracionHPS.ResolveBackendFileName: Connect no reconocido, marcando como Unknown.accdb: " & p_Connect
        ResolveBackendFileName = k_BackendUnknown
    End If
End Function

' Recorre TableDefs del frontend y registra cada tabla vinculada
' no-MSys en TbVinculosTablas con el backend inferido.
Private Sub PoblarVinculosDesdeTablas(ByRef p_Db As DAO.Database)
    Dim tdf As DAO.TableDef
    Dim backendFile As String
    For Each tdf In p_Db.TableDefs
        If (tdf.Attributes And dbAttachedTable) = dbAttachedTable Then
            If Left$(tdf.Name, 4) <> "MSys" Then
                backendFile = ResolveBackendFileName(tdf.Connect)
                UpsertVinculo p_Db, tdf.Name, backendFile
            End If
        End If
    Next tdf
End Sub

' Devuelve el Valor de TbConfiguracionHPS para una Clave, o "" si la fila
' no existe o la tabla no esta. Inverso de UpsertConfig -- pensado para
' que AplicarConfiguracionEnVinculos pueda leer los paths de los backends
' en runtime (sin hardcodear constantes). Toleramos que la tabla no exista
' (helper de PR A nunca corrio): en ese caso devolvemos "" para que el
' caller pueda detectar la situacion y raise con un mensaje claro.
Private Function GetConfigValue(ByRef p_Db As DAO.Database, ByVal p_Clave As String) As String
    Dim rs As DAO.Recordset
    Dim claveSafe As String
    Dim tdf As DAO.TableDef

    claveSafe = Replace(p_Clave, "'", "''")

    ' Si la tabla no existe todavia, devolvemos "" en silencio.
    On Error Resume Next
    Set tdf = p_Db.TableDefs(k_TableConfig)
    If Err.Number <> 0 Or tdf Is Nothing Then
        GetConfigValue = ""
        Exit Function
    End If
    On Error GoTo 0

    Set rs = p_Db.OpenRecordset( _
        "SELECT Valor FROM " & k_TableConfig & " WHERE Clave='" & claveSafe & "'", _
        dbOpenSnapshot)
    If rs.EOF Then
        GetConfigValue = ""
    Else
        GetConfigValue = Nz(rs.Fields("Valor").Value, "")
    End If
    rs.Close
    Set rs = Nothing
End Function

' Strip del filename -> devuelve el parent dir de un path. Usa string
' manipulation (Left$ + InStrRev) en vez de fso.GetParentFolderName
' porque esta ultima no maneja bien UNC paths en algunos runtimes. Tolera
' barras finales redundantes y separadores mezclados ("\\" y "/").
Private Function ExtractParentDir(ByVal p_Path As String) As String
    Dim sTrimmed As String
    Dim iSep As Long
    sTrimmed = Trim$(p_Path)
    Do While Len(sTrimmed) > 0 And (Right$(sTrimmed, 1) = "\" Or Right$(sTrimmed, 1) = "/")
        sTrimmed = Left$(sTrimmed, Len(sTrimmed) - 1)
    Loop
    iSep = InStrRev(sTrimmed, "\")
    If iSep <= 0 Then
        ' Sin separador -> el path entero es "parent + filename" donde
        ' parent = "" (caso degenerado, no esperado en este proyecto).
        ExtractParentDir = ""
    Else
        ExtractParentDir = Left$(sTrimmed, iSep - 1)
    End If
End Function

' Devuelve el parent comun que RelinkAllTables debe recibir. Soporta tanto
' backends en el mismo directorio (Local) como backends en subdirectorios
' hermanos (Produccion: HPS\HPST.accdb y EXPEDIENTES\Expedientes_datos.accdb).
Private Function DeriveCommonBackendParent(ByVal p_PathHPST As String, ByVal p_PathExpedientes As String) As String
    Dim sParentHPST As String
    Dim sParentExpedientes As String
    Dim sGrandParentHPST As String
    Dim sGrandParentExpedientes As String

    sParentHPST = ExtractParentDir(p_PathHPST)
    sParentExpedientes = ExtractParentDir(p_PathExpedientes)

    If Len(sParentHPST) = 0 Or Len(sParentExpedientes) = 0 Then
        DeriveCommonBackendParent = ""
        Exit Function
    End If

    If StrComp(sParentHPST, sParentExpedientes, vbTextCompare) = 0 Then
        DeriveCommonBackendParent = sParentHPST
        Exit Function
    End If

    sGrandParentHPST = ExtractParentDir(sParentHPST)
    sGrandParentExpedientes = ExtractParentDir(sParentExpedientes)
    If Len(sGrandParentHPST) > 0 Then
        If StrComp(sGrandParentHPST, sGrandParentExpedientes, vbTextCompare) = 0 Then
            DeriveCommonBackendParent = sGrandParentHPST
            Exit Function
        End If
    End If

    DeriveCommonBackendParent = ""
End Function

' Convierte un path completo en el valor BackendFile que RelinkAllTables concatena
' con el parent comun. Local => HPST.accdb; Produccion => HPS\HPST.accdb.
Private Function BuildBackendRelativePath(ByVal p_FullPath As String, ByVal p_CommonParent As String) As String
    Dim sFullPath As String
    Dim sCommonParent As String

    sFullPath = Trim$(p_FullPath)
    sCommonParent = Trim$(p_CommonParent)
    Do While Len(sCommonParent) > 0 And (Right$(sCommonParent, 1) = "\" Or Right$(sCommonParent, 1) = "/")
        sCommonParent = Left$(sCommonParent, Len(sCommonParent) - 1)
    Loop

    If Len(sCommonParent) > 0 Then
        If StrComp(Left$(sFullPath, Len(sCommonParent) + 1), sCommonParent & "\", vbTextCompare) = 0 Then
            BuildBackendRelativePath = Mid$(sFullPath, Len(sCommonParent) + 2)
            Exit Function
        End If
    End If

    BuildBackendRelativePath = Mid$(sFullPath, InStrRev(sFullPath, "\") + 1)
End Function

' Alinea TbVinculosTablas con los paths de TbConfiguracionHPS sin perder la
' asociacion TablaOrigen -> backend inferida desde los TableDefs actuales.
Private Sub SyncVinculosDesdeConfig(ByRef p_Db As DAO.Database)
    Dim sPathHPST As String
    Dim sPathExpedientes As String
    Dim sCommonParent As String
    Dim sRelHPST As String
    Dim sRelExpedientes As String
    Dim rs As DAO.Recordset
    Dim sTabla As String
    Dim sBackend As String

    sPathHPST = GetConfigValue(p_Db, "HPST_BACKEND_PATH")
    sPathExpedientes = GetConfigValue(p_Db, "EXPEDIENTES_BACKEND_PATH")
    sCommonParent = DeriveCommonBackendParent(sPathHPST, sPathExpedientes)
    If Len(sCommonParent) = 0 Then Exit Sub

    sRelHPST = BuildBackendRelativePath(sPathHPST, sCommonParent)
    sRelExpedientes = BuildBackendRelativePath(sPathExpedientes, sCommonParent)

    Set rs = p_Db.OpenRecordset("SELECT TablaOrigen, BackendFile FROM " & k_TableVinculos, dbOpenSnapshot)
    Do While Not rs.EOF
        sTabla = CStr(Nz(rs.Fields("TablaOrigen").Value, ""))
        sBackend = CStr(Nz(rs.Fields("BackendFile").Value, ""))
        If Len(sTabla) > 0 Then
            If InStr(1, sBackend, k_BackendHPST, vbTextCompare) > 0 Then
                UpsertVinculo p_Db, sTabla, sRelHPST
            ElseIf InStr(1, sBackend, k_BackendExpedientes, vbTextCompare) > 0 Then
                UpsertVinculo p_Db, sTabla, sRelExpedientes
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

' --- Public Subs (VBE-callable) ---

' Preconfigura TbConfiguracionHPS y TbVinculosTablas con el perfil Local.
' Valores pensados para la maquina de desarrollo (rutas locales).
Public Sub ConfigurarTablaConfiguracion_PerfilLocal()
    On Error GoTo EH
    Dim db As DAO.Database
    Dim sBinPath As String
    Set db = CurrentDb

    ExecuteCreateTableIfNotExists db, k_TableConfig, k_SchemaConfigSql
    ExecuteCreateTableIfNotExists db, k_TableVinculos, k_SchemaVinculosSql

    ' [config-as-source-of-truth] Perfil Local: los .accdb viven en el
    ' mismo equipo (C:\00repos\datos). El root APP_ROOT_LOCAL apunta a
    ' la carpeta junto al binario (CurrentProject.path\) y APP_ROOT_REMOTO
    ' apunta al share corporativo \\datoste\aplicaciones_dys\Aplicaciones
    ' PpD\. El resolver (Entorno.ResolveURLCarpetaAnexosCandidate)
    ' escoge automáticamente APP_ROOT_REMOTO si \\datoste es alcanzable
    ' y APP_ROOT_LOCAL si no. Asi "config gobierna" siempre: la config
    ' dice dónde están los backends y los anexos; el código no
    ' sobreescribe con constantes.
    sBinPath = Application.CurrentProject.path
    UpsertConfig db, "APP_ROOT_LOCAL", sBinPath & "\"
    UpsertConfig db, "APP_ROOT_REMOTO", "\\datoste\aplicaciones_dys\Aplicaciones PpD\"
    UpsertConfig db, "HPST_BACKEND_PATH", "C:\00repos\datos\HPST.accdb"
    UpsertConfig db, "EXPEDIENTES_BACKEND_PATH", "C:\00repos\datos\Expedientes_datos.accdb"
    UpsertConfig db, "LANZADERA_BACKEND_PATH", "C:\00repos\datos\Lanzadera_Datos.accdb"
    UpsertConfig db, "CORREOS_BACKEND_PATH", "C:\00repos\datos\Correos_datos.accdb"
    UpsertConfig db, "SOLICITUDES_HPS_BACKEND_PATH", "C:\00repos\datos\Solicitudes_HPS_datos.accdb"
    UpsertConfig db, "MODO_LOCAL_DEFAULT", "S" & Chr$(237)
    UpsertConfig db, "BACKEND_PASSWORD", "dpddpd"

    PoblarVinculosDesdeTablas db
    SyncVinculosDesdeConfig db

    Exit Sub
EH:
    Err.Raise Err.Number, Err.Source, "ConfigurarTablaConfiguracion_PerfilLocal: " & Err.Description
End Sub

' Preconfigura TbConfiguracionHPS y TbVinculosTablas con el perfil Produccion.
' Valores pensados para el entorno productivo (rutas UNC de red corporativa).
Public Sub ConfigurarTablaConfiguracion_PerfilProduccion()
    On Error GoTo EH
    Dim db As DAO.Database
    Set db = CurrentDb

    ExecuteCreateTableIfNotExists db, k_TableConfig, k_SchemaConfigSql
    ExecuteCreateTableIfNotExists db, k_TableVinculos, k_SchemaVinculosSql

    ' [config-as-source-of-truth] Perfil Produccion: en oficina con
    ' VPN al CARU, el root unico es APP_ROOT_REMOTO=\\datoste\aplicaciones_dys\
    ' y todos los .accdb viven debajo. Los anexos viven bajo APP_ROOT_REMOTO
    ' tambien (regla "config gobierna" que vos pediste).
    UpsertConfig db, "APP_ROOT_REMOTO", "\\datoste\aplicaciones_dys\Aplicaciones PpD\"
    UpsertConfig db, "APP_ROOT_LOCAL", "\\datoste\aplicaciones_dys\Aplicaciones PpD\"
    UpsertConfig db, "HPST_BACKEND_PATH", "\\datoste\aplicaciones_dys\Aplicaciones PpD\HPS\HPST.accdb"
    UpsertConfig db, "EXPEDIENTES_BACKEND_PATH", "\\datoste\aplicaciones_dys\Aplicaciones PpD\EXPEDIENTES\Expedientes_datos.accdb"
    UpsertConfig db, "LANZADERA_BACKEND_PATH", "\\datoste\aplicaciones_dys\Aplicaciones PpD\0Lanzadera\Lanzadera_Datos.accdb"
    UpsertConfig db, "CORREOS_BACKEND_PATH", "\\datoste\aplicaciones_dys\Aplicaciones PpD\00Recursos\Correos_datos.accdb"
    UpsertConfig db, "SOLICITUDES_HPS_BACKEND_PATH", "\\datoste\aplicaciones_dys\Aplicaciones PpD\SOLICITUDES HPS\Solicitudes_HPS_datos.accdb"
    UpsertConfig db, "MODO_LOCAL_DEFAULT", "No"
    UpsertConfig db, "BACKEND_PASSWORD", "dpddpd"

    PoblarVinculosDesdeTablas db
    SyncVinculosDesdeConfig db

    Exit Sub
EH:
    Err.Raise Err.Number, Err.Source, "ConfigurarTablaConfiguracion_PerfilProduccion: " & Err.Description
End Sub

' Preconfigura TbConfiguracionHPS y TbVinculosTablas con el perfil
' RemotoPruebas. Análogo a PerfilProduccion pero con la raíz UNC
' \\datoste\aplicaciones_dys\Aplicaciones PpD\00_datos_prueba\ (share de
' pruebas internas). Bajo esa raíz viven los mismos subdirectorios y
' filenames que en producción (HPS\, EXPEDIENTES\, 0Lanzadera\,
' 00Recursos\, SOLICITUDES HPS\) para que el resolver y los tests
' puedan ejercitar el código de producción contra datos sintéticos.
' Los adjuntos (anexos) viven bajo APP_ROOT_REMOTO\ANEXOS\ -- el
' resolver compone esa ruta a partir de APP_ROOT_REMOTO, igual que en
' producción. APP_ROOT_LOCAL mantiene la convención del perfil
' PerfilLocal (CurrentProject.path\) para que el fallback local siga
' funcionando si el share de pruebas no está accesible.
Public Sub ConfigurarTablaConfiguracion_PerfilRemotoPruebas()
    On Error GoTo EH
    Dim db As DAO.Database
    Dim sBinPath As String

    Set db = CurrentDb

    ExecuteCreateTableIfNotExists db, k_TableConfig, k_SchemaConfigSql
    ExecuteCreateTableIfNotExists db, k_TableVinculos, k_SchemaVinculosSql

    sBinPath = Application.CurrentProject.path

    ' [config-as-source-of-truth] Perfil RemotoPruebas: raíz UNC única
    ' APP_ROOT_REMOTO=\\datoste\aplicaciones_dys\Aplicaciones PpD\00_datos_prueba\
    ' y los mismos subdirectorios que Producción. APP_ROOT_LOCAL conserva
    ' la convención del perfil Local (junto al binario) para fallback.
    UpsertConfig db, "APP_ROOT_REMOTO", "\\datoste\aplicaciones_dys\Aplicaciones PpD\00_datos_prueba\"
    UpsertConfig db, "APP_ROOT_LOCAL", sBinPath & "\"
    UpsertConfig db, "HPST_BACKEND_PATH", "\\datoste\aplicaciones_dys\Aplicaciones PpD\00_datos_prueba\HPS\HPST.accdb"
    UpsertConfig db, "EXPEDIENTES_BACKEND_PATH", "\\datoste\aplicaciones_dys\Aplicaciones PpD\00_datos_prueba\EXPEDIENTES\Expedientes_datos.accdb"
    UpsertConfig db, "LANZADERA_BACKEND_PATH", "\\datoste\aplicaciones_dys\Aplicaciones PpD\00_datos_prueba\0Lanzadera\Lanzadera_Datos.accdb"
    UpsertConfig db, "CORREOS_BACKEND_PATH", "\\datoste\aplicaciones_dys\Aplicaciones PpD\00_datos_prueba\00Recursos\Correos_datos.accdb"
    UpsertConfig db, "SOLICITUDES_HPS_BACKEND_PATH", "\\datoste\aplicaciones_dys\Aplicaciones PpD\00_datos_prueba\SOLICITUDES HPS\Solicitudes_HPS_datos.accdb"
    UpsertConfig db, "MODO_LOCAL_DEFAULT", "No"
    UpsertConfig db, "BACKEND_PASSWORD", "dpddpd"

    PoblarVinculosDesdeTablas db
    SyncVinculosDesdeConfig db

    Exit Sub
EH:
    Err.Raise Err.Number, Err.Source, "ConfigurarTablaConfiguracion_PerfilRemotoPruebas: " & Err.Description
End Sub

' Dispatcher por nombre de perfil. Acepta variantes de mayusculas/minusculas
' y la tilde en Produccion. Si el perfil no se reconoce, raise Err 1000.
Public Sub ConfigurarTablaConfiguracion_Perfil(ByVal p_NombrePerfil As String)
    Select Case UCase$(Trim$(p_NombrePerfil))
        Case "LOCAL"
            ConfigurarTablaConfiguracion_PerfilLocal
        Case "PRODUCCION"
            ConfigurarTablaConfiguracion_PerfilProduccion
        Case "REMOTOPRUEBAS"
            ConfigurarTablaConfiguracion_PerfilRemotoPruebas
        Case Else
            Err.Raise 1000, "modConfiguracionHPS.ConfigurarTablaConfiguracion_Perfil", _
                "Perfil desconocido: " & p_NombrePerfil & ". Valores validos: Local, Produccion, RemotoPruebas."
    End Select
End Sub

' --- Workflow helpers (VBE-callable desde la ventana Inmediato) ---

' PrepararProduccion -- workflow de un solo paso para preparar el .accdb
' antes de compilar a .accde. Pensado para correr en dev desde la ventana
' Inmediato del VBE:
'
'     PrepararProduccion
'
' Pasos:
'   1) Popula TbConfiguracionHPS y TbVinculosTablas con valores UNC
'      de Produccion (via ConfigurarTablaConfiguracion_PerfilProduccion).
'   2) Refresca la cache del Entorno (New entorno + CargarConfiguracion)
'      para que GetConfig vea los nuevos paths.
'   3) Re-apunta todas las tablas adjuntas al parent dir
'      \\datoste\aplicaciones_dys\Aplicaciones PpD (via RelinkAllTables).
'   4) Loguea a Debug.Print el resultado.
'
' Los usuarios finales NO corren esto -- lo corre el dev/admin antes de
' compilar a .accde, porque en el .accde los TableDefs no son visibles
' para relink manual desde la UI.
Public Sub PrepararProduccion()
    On Error GoTo EH
    Dim oEnt As entorno
    Dim sParentDir As String
    Dim bRelinkOk As Boolean

    ' 1) Poblar tablas de config + vinculos con perfil Produccion.
    ConfigurarTablaConfiguracion_PerfilProduccion

    ' 2) Refrescar la cache del Entorno para que GetConfig vea los nuevos paths.
    Set oEnt = New entorno
    oEnt.CargarConfiguracion

    ' 3) Parent dir Produccion: lugar donde estan los .accdb de HPST y
    '    Expedientes (los subdirs HPS\ y EXPEDIENTES\ cuelgan de aca).
    sParentDir = "\\datoste\aplicaciones_dys\Aplicaciones PpD"

    ' 4) Re-apuntar todos los TableDefs adjuntos.
    bRelinkOk = oEnt.RelinkAllTables(sParentDir)

    Debug.Print "modConfiguracionHPS.PrepararProduccion: relink " & _
        IIf(bRelinkOk, "OK", "con fallas (best-effort)") & " apuntando a '" & sParentDir & "'."

    Exit Sub
EH:
    Err.Raise Err.Number, Err.Source, "PrepararProduccion: " & Err.Description
End Sub

' PrepararLocal -- equivalente a PrepararProduccion pero con los paths
' del entorno de desarrollo (C:\00repos\datos). Pensado para volver el
' .accdb a modo dev despues de haberlo preparado para produccion.
Public Sub PrepararLocal()
    On Error GoTo EH
    Dim oEnt As entorno
    Dim sParentDir As String
    Dim bRelinkOk As Boolean

    ' 1) Poblar tablas de config + vinculos con perfil Local.
    ConfigurarTablaConfiguracion_PerfilLocal

    ' 2) Refrescar la cache del Entorno.
    Set oEnt = New entorno
    oEnt.CargarConfiguracion

    ' 3) Parent dir Local: donde estan HPST.accdb y Expedientes_datos.accdb
    '    en la maquina de desarrollo (los backends estan ahi, sin subdirs).
    sParentDir = "C:\00repos\datos"

    ' 4) Re-apuntar todos los TableDefs adjuntos.
    bRelinkOk = oEnt.RelinkAllTables(sParentDir)

    Debug.Print "modConfiguracionHPS.PrepararLocal: relink " & _
        IIf(bRelinkOk, "OK", "con fallas (best-effort)") & " apuntando a '" & sParentDir & "'."

    Exit Sub
EH:
    Err.Raise Err.Number, Err.Source, "PrepararLocal: " & Err.Description
End Sub

' PrepararRemotoPruebas -- workflow de un solo paso para preparar el
' .accdb apuntando al share de pruebas internas
' \\datoste\aplicaciones_dys\Aplicaciones PpD\00_datos_prueba\. Análogo a
' PrepararProduccion: misma siembra (vía PerfilRemotoPruebas), misma
' recarga de cache del Entorno y mismo RelinkAllTables sobre el parent
' dir remoto. Pensado para correr desde la ventana Inmediato del VBE
' antes de compilar a .accde, en cualquier máquina que tenga acceso al
' share de pruebas.
Public Sub PrepararRemotoPruebas()
    On Error GoTo EH
    Dim oEnt As entorno
    Dim sParentDir As String
    Dim bRelinkOk As Boolean

    ' 1) Poblar tablas de config + vinculos con perfil RemotoPruebas.
    ConfigurarTablaConfiguracion_PerfilRemotoPruebas

    ' 2) Refrescar la cache del Entorno para que GetConfig vea los nuevos paths.
    Set oEnt = New entorno
    oEnt.CargarConfiguracion

    ' 3) Parent dir RemotoPruebas: lugar donde estan los .accdb de HPST
    '    y Expedientes (los subdirs HPS\ y EXPEDIENTES\ cuelgan de aca).
    sParentDir = "\\datoste\aplicaciones_dys\Aplicaciones PpD\00_datos_prueba"

    ' 4) Re-apuntar todos los TableDefs adjuntos.
    bRelinkOk = oEnt.RelinkAllTables(sParentDir)

    Debug.Print "modConfiguracionHPS.PrepararRemotoPruebas: relink " & _
        IIf(bRelinkOk, "OK", "con fallas (best-effort)") & " apuntando a '" & sParentDir & "'."

    Exit Sub
EH:
    Err.Raise Err.Number, Err.Source, "PrepararRemotoPruebas: " & Err.Description
End Sub

' AplicarConfiguracionEnVinculos -- workflow de un solo paso para
' reapuntar los TableDefs adjuntos usando los paths que el helper de
' PR A ya guardo en TbConfiguracionHPS. A diferencia de PrepararLocal /
' PrepararProduccion (que hardcodean el parent dir segun el perfil),
' este sub LEE los 2 paths de la tabla y verifica alcanzabilidad ANTES
' de tocar TableDefs.
'
' Pasos:
'   1) Lee HPST_BACKEND_PATH y EXPEDIENTES_BACKEND_PATH de TbConfiguracionHPS.
'   2) Si alguno falta o la tabla esta vacia -> raise sin tocar TableDefs.
'   3) Deriva el parent comun de los paths completos. Soporta backends en
'      el mismo directorio (Local) y subdirectorios hermanos (Produccion).
'   4) Verifica que los 2 paths completos existen antes de tocar TableDefs.
'   5) Sincroniza TbVinculosTablas con rutas relativas al parent comun.
'   6) Si todo OK -> New entorno + CargarConfiguracion + RelinkAllTables.
'
' Pensado para correr desde la ventana Inmediato del VBE antes de
' compilar a .accde, en cualquier entorno donde el helper de PR A ya
' haya poblado la tabla.
Public Sub AplicarConfiguracionEnVinculos()
    On Error GoTo EH
    Dim db As DAO.Database
    Dim sPathHPST As String
    Dim sPathExpedientes As String
    Dim sCommonParent As String
    Dim oEnt As entorno
    Dim bRelinkOk As Boolean

    Set db = CurrentDb

    ' 1) Leer los 2 paths de los backends desde TbConfiguracionHPS.
    sPathHPST = GetConfigValue(db, "HPST_BACKEND_PATH")
    sPathExpedientes = GetConfigValue(db, "EXPEDIENTES_BACKEND_PATH")

    If Len(sPathHPST) = 0 Or Len(sPathExpedientes) = 0 Then
        Err.Raise vbObjectError + 9001, _
            "modConfiguracionHPS.AplicarConfiguracionEnVinculos", _
            "TbConfiguracionHPS esta vacia o no contiene HPST_BACKEND_PATH / EXPEDIENTES_BACKEND_PATH. " & _
            "Aplicar 'ConfigurarTablaConfiguracion_PerfilLocal' o 'ConfigurarTablaConfiguracion_PerfilProduccion' primero."
    End If

    ' 2) Derivar el parent comun de los paths completos.
    sCommonParent = DeriveCommonBackendParent(sPathHPST, sPathExpedientes)
    If Len(sCommonParent) = 0 Then
        Err.Raise vbObjectError + 9001, _
            "modConfiguracionHPS.AplicarConfiguracionEnVinculos", _
            "No se pudo derivar un directorio padre comun para los paths: '" & _
            sPathHPST & "' / '" & sPathExpedientes & "'. " & _
            "Corregir las claves HPST_BACKEND_PATH y EXPEDIENTES_BACKEND_PATH en TbConfiguracionHPS."
    End If

    ' 3) Verificar alcanzabilidad de cada backend por su path completo
    '    ANTES de tocar TableDefs. Si falta cualquiera, raise sin relink.
    If Not fso.FileExists(sPathHPST) Then
        Err.Raise vbObjectError + 9001, _
            "modConfiguracionHPS.AplicarConfiguracionEnVinculos", _
            "No se puede hacer relink: archivo '" & sPathHPST & "' no existe. " & _
            "Aplicar 'ConfigurarTablaConfiguracion_PerfilLocal' o 'ConfigurarTablaConfiguracion_PerfilProduccion' primero " & _
            "o corregir la ruta en TbConfiguracionHPS."
    End If
    If Not fso.FileExists(sPathExpedientes) Then
        Err.Raise vbObjectError + 9001, _
            "modConfiguracionHPS.AplicarConfiguracionEnVinculos", _
            "No se puede hacer relink: archivo '" & sPathExpedientes & "' no existe. " & _
            "Aplicar 'ConfigurarTablaConfiguracion_PerfilLocal' o 'ConfigurarTablaConfiguracion_PerfilProduccion' primero " & _
            "o corregir la ruta en TbConfiguracionHPS."
    End If

    ' 4) Poblar TbVinculosTablas con todos los TableDefs adjuntos actuales.
    ' Sin este paso, las tablas que no se catalogaron cuando corrio
    ' PrepararLocal/PrepararProduccion (o cuyo backend nunca se infirio
    ' correctamente) quedan con el Connect viejo y RelinkAllTables las
    ' salta silenciosamente. Idempotente - solo agrega/actualiza, nunca
    ' borra entries.
    PoblarVinculosDesdeTablas db

    ' 5) Alinear TbVinculosTablas con los paths actuales.
    SyncVinculosDesdeConfig db

    ' 6) Refrescar la cache del Entorno para que GetConfig vea los nuevos paths.
    Set oEnt = New entorno
    oEnt.CargarConfiguracion

    ' 7) Re-apuntar todos los TableDefs adjuntos.
    bRelinkOk = oEnt.RelinkAllTables(sCommonParent)

    Debug.Print "modConfiguracionHPS.AplicarConfiguracionEnVinculos: relink " & _
        IIf(bRelinkOk, "OK", "con fallas (best-effort)") & " apuntando a '" & sCommonParent & "'."

    Exit Sub
EH:
    Err.Raise Err.Number, Err.Source, "AplicarConfiguracionEnVinculos: " & Err.Description
End Sub
