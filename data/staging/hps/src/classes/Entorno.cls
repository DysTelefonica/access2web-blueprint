VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Entorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' --- Private Const (PR B: TbConfiguracionHPS / TbVinculosTablas cache layer) ---
' These are the column names the cache layer reads from TbConfiguracionHPS
' (prepopulated by modConfiguracionHPS) and the lookup table the relink
' helper iterates (TbVinculosTablas). Kept module-level so tests in
' Test_HPSEntorno.bas can reference them by name without a magic-string
' literal at the call site.
Private Const k_ConfigTable As String = "TbConfiguracionHPS"
Private Const k_VinculosTable As String = "TbVinculosTablas"
Private Const k_Key_HPST As String = "HPST_BACKEND_PATH"
Private Const k_Key_Expedientes As String = "EXPEDIENTES_BACKEND_PATH"
Private Const k_Key_AnexosHPS As String = "ANEXOS_HPS_PATH"
Private Const k_Key_AnexosHistorico As String = "ANEXOS_HISTORICO_PATH"
Private Const k_Key_AnexosSICA As String = "ANEXOS_SICA_PATH"
Private Const k_Key_ModoLocal As String = "MODO_LOCAL_DEFAULT"
Private Const k_Key_BackendPassword As String = "BACKEND_PASSWORD"
' [binario-driven] Los anexos siempre viven junto al binario de Access
' (Application.CurrentProject.path\ANEXOS\). Las rutas NO se leen de
' TbConfiguracionHPS, así que el operador no puede redirigirlas con un
' path de prueba: el binario es la única verdad para los anexos. Los
' .accdb (HPST, Expedientes, Lanzadera, Correos, SolicitudesHPS) sí
' viven en TbConfiguracionHPS vía GetConfigOrFail.
Private Const k_Key_AppRootLocal As String = "APP_ROOT_LOCAL"
Private Const k_Key_AppRootRemoto As String = "APP_ROOT_REMOTO"
Private Const k_Key_LanzaderaBackend As String = "LANZADERA_BACKEND_PATH"
Private Const k_Key_CorreosBackend As String = "CORREOS_BACKEND_PATH"
Private Const k_Key_SolicitudesHPSBackend As String = "SOLICITUDES_HPS_BACKEND_PATH"

Public URLArchivoUltimo As String
Public SoloLectura As EnumSiNo

Private m_objcolColoresMaster As Scripting.Dictionary
Private m_objcolColoresSlave As Scripting.Dictionary
Private m_lCabeceraBackColor As Long
Private m_lListaBackColor As Long
Private m_lListaForeColor As Long
Private m_lListaBorderColor As Long
Private m_lEtiquetaBackColor  As Long
Private m_lEtiquetaLinkBackColorSeleccionado  As Long
Private m_lEtiquetaLinkForeColor As Long
Private m_lEtiquetaForeColor  As Long
Private m_lEtiquetaForeColorSeleccionado  As Long
Private m_lEtiquetaBorderColor  As Long
Private m_lEtiquetaSeleccionadaForeColor  As Long
Private m_lEtiquetaSeleccionadaBackColor  As Long
Private m_lCampoBackColor  As Long
Private m_lCampoForeColor  As Long
Private m_lCampoBorderColor  As Long
Private m_sURLCarpetaAplicacionesDys As String
Private m_sURLCarpetaRecursos As String
Private m_sURLCarpetaAplicacion As String
Private m_sURLCarpetaAnexos As String
Private m_sURLCarpetaAnexosSICA As String
Private m_sURLCarpetaAnexosHistoricos As String
Private m_sURLDirectorioLocal As String
Private m_sURLCarpetaIconos As String


Private m_sURLIni As String
Private m_sVersionAplicacion As String
Private m_objColEstadosVisiblesHPS As Scripting.Dictionary
Private m_objColVisiblesHPSEstados As Scripting.Dictionary
Private m_objColEstadosHPS As Scripting.Dictionary
Private m_objColHPSEstados As Scripting.Dictionary
Private m_objColTiposConsultas As Scripting.Dictionary
Private m_objColTiposConsultasFormulario As Scripting.Dictionary
Private m_objcolGradosHPS As Scripting.Dictionary
Private m_objcolFiltroEnIndicadordes As Scripting.Dictionary
Private m_objColHPS As Scripting.Dictionary
Private m_eSoloLectura As EnumSiNo
Private m_objColUsuariosEstadoIrregular As Scripting.Dictionary
Private m_objColUsuariosAPuntoCaducar As Scripting.Dictionary
Private m_objColUsuariosCaducados As Scripting.Dictionary
Private m_objColRequiereCursoFaltaPrimeraConvocatoria As Scripting.Dictionary
Private m_objColRequiereCursoFaltaSegundaConvocatoria As Scripting.Dictionary
Private m_objColRequiereCursoFaltaCorreoJefeSeguridad As Scripting.Dictionary
Private m_sDatosEn As String
Private m_sModoEscritura As String
Private m_objColPropiedadesBotonEscritura As Scripting.Dictionary
Private m_objColPropiedadesBotonLectura As Scripting.Dictionary
Private m_objColUsuariosHPSRequierenCurso As Scripting.Dictionary
Private m_objColUsuariosEnSolicitud As Scripting.Dictionary
Private m_objColCamposConsulaUsuariosCompleta As Scripting.Dictionary
Private m_sNumeroUsuariosHPSSinCurso As String
Private m_sNumeroUsuariosEnSolicitud As String

Private m_sURLImagenSemaforoRojo As String
Private m_sURLImagenSemaforoVerde As String

Private m_objColEmpresas As Scripting.Dictionary
Private m_objColEmpresasTramitadoras As Scripting.Dictionary
Private m_objColCadenaContratistas As Scripting.Dictionary
Private m_sURLArchivoExcelPrevio As String
Private m_objColConsultas As Scripting.Dictionary
Private m_objColCamposSuministradores As Scripting.Dictionary
Private m_ObjUsuarioConectadoMaquina As usuario
Private m_objcolUsuariosPermitidosEnLocal As Scripting.Dictionary
Private m_iMesesParaRenovacion As Integer
Private m_lDiasEnPrimeraConvocatoriaSinCurso As Long
Private m_lDiasSinCorreoNoCurso As Long
Private m_lDiasEnSegundaConvocatoriaSinCurso As Long
Private m_objExpedientesUsados As Scripting.Dictionary
Private m_ObjExpedientes As Scripting.Dictionary
Private m_ObjUsuariosHPS As Scripting.Dictionary
Private m_ObjUsuariosHPSNoEntidades As Scripting.Dictionary
Private m_ObjUsuariosEnEntidadNoExistentes As Scripting.Dictionary
Private m_ObjUsuariosNoEnEntidad As Scripting.Dictionary
Private m_ObjMotivos As Scripting.Dictionary

' PR B: lazy-init cache for TbConfiguracionHPS (clave -> valor).
' Populated by CargarConfiguracion (called from EVE in VariablesEntorno
' or lazily from the URL property Gets below). Scripting.Dictionary gives
' us TextCompare lookups so callers can pass the key in any case.
Private m_ConfigCache As Scripting.Dictionary

Public Error As String
Private m_Error As String
Private m_objColPropiedadesEntorno As Scripting.Dictionary

Public Function ResolveURLCarpetaAnexosCandidate(Optional ByVal p_Historicos As Boolean = False) As String
    ' [config-as-source-of-truth] Los anexos viven bajo APP_ROOT_REMOTO
    ' o APP_ROOT_LOCAL según el estado de red real del usuario. NO usa
    ' constantes UNC hardcoded ni Application.CurrentProject.path a
    ' ciegas: la regla es
    '
    '   1) Si APP_ROOT_REMOTO (\\datoste\aplicaciones_dys\Aplicaciones
    '      PpD\) es alcanzable en disco, ese es el root de anexos
    '      porque forma parte del mismo share donde vive el .accdb.
    '   2) Si NO es alcanzable (casa sin VPN), APP_ROOT_LOCAL define
    '      un root local que el operador habrá sembrado con un path
    '      alcanzable (ej. junto al binario).
    '
    ' La verificación es de filesystem (fso.FolderExists) y no depende
    ' de flags globales m_EnOficina/m_AccesoADatosTE que pueden mentir.
    Dim sAppRoot As String
    Dim sRemoto As String
    Dim sLocal As String
    sRemoto = GetConfigOrFail(k_Key_AppRootRemoto)
    sLocal = GetConfigOrFail(k_Key_AppRootLocal)
    If fso.FolderExists(sRemoto) Then
        sAppRoot = sRemoto
    ElseIf fso.FolderExists(sLocal) Then
        sAppRoot = sLocal
    Else
        ' Sin root alcanzable. La config no miente: el operador sembró
        ' un path inalcanzable. Fallamos con mensaje accionable.
        m_Error = "ResolveURLCarpetaAnexosCandidate: ni APP_ROOT_REMOTO ('" & _
            sRemoto & "') ni APP_ROOT_LOCAL ('" & sLocal & _
            "') son alcanzables. Verifica la red VPN o actualiza la " & _
            "configuración con ConfigurarTablaConfiguracion_PerfilLocal."
        Err.Raise 1000
    End If
    ' [config-as-source-of-truth] El resolver devuelve el root puro
    ' de anexos. UsuarioHPS.URLCarpetaAnexosActuales/Historicos
    ' concatena "\<DNI>\Actual\" o "\<DNI>\HISTORICO\" sobre este root.
    ' Asi el root de la config es la unica verdad y el subdirectorio
    ' por usuario lo compone el codigo de modelo de UsuarioHPS.
    ResolveURLCarpetaAnexosCandidate = sAppRoot
End Function

Private Function EnsureSlash(ByVal p_Path As String) As String
    ' Helper local: garantiza que un path de carpeta termina en "\".
    ' Replicado en otros coordinators; este es el del entorno.
    If Len(p_Path) = 0 Then
        EnsureSlash = ""
    ElseIf Right$(p_Path, 1) = "\" Then
        EnsureSlash = p_Path
    Else
        EnsureSlash = p_Path & "\"
    End If
End Function

Public Property Get Motivos() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    If Not m_ObjMotivos Is Nothing Then
        Set Motivos = m_ObjMotivos
        Exit Property
    End If
    Set m_ObjMotivos = Constructor.getMotivos(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set Motivos = m_ObjMotivos
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.Motivos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set Motivos(p_Valor As Variant)
    
    Set m_ObjMotivos = p_Valor
    
End Property

Public Property Get UsuariosHPS() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    If Not m_ObjUsuariosHPS Is Nothing Then
        Set UsuariosHPS = m_ObjUsuariosHPS
        Exit Property
    End If
    Set m_ObjUsuariosHPS = Constructor.getUsuariosHPS(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set UsuariosHPS = m_ObjUsuariosHPS
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.UsuariosHPS ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set UsuariosHPS(p_Valor As Variant)
    
    Set m_ObjUsuariosHPS = p_Valor
    
End Property
Public Property Get UsuariosHPSNoEntidades() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    If Not m_ObjUsuariosHPSNoEntidades Is Nothing Then
        Set UsuariosHPSNoEntidades = m_ObjUsuariosHPSNoEntidades
        Exit Property
    End If
    Set m_ObjUsuariosHPSNoEntidades = Constructor.getUsuariosHPSNoEntidades(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set UsuariosHPSNoEntidades = m_ObjUsuariosHPSNoEntidades
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.UsuariosHPSNoEntidades ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set UsuariosHPSNoEntidades(p_Valor As Variant)
    
    Set m_ObjUsuariosHPSNoEntidades = p_Valor
    
End Property

Public Property Get UsuariosEnEntidadNoExistentes() As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    Dim m_ID As String
    On Error GoTo errores
    Me.Error = ""
    If Not m_ObjUsuariosEnEntidadNoExistentes Is Nothing Then
        Set UsuariosEnEntidadNoExistentes = m_ObjUsuariosEnEntidadNoExistentes
        Exit Property
    End If
    
    
    m_SQL = "SELECT TbUsuariosEntidades.ID " & _
            "FROM TbUsuarios RIGHT JOIN TbUsuariosEntidades ON TbUsuarios.ID = TbUsuariosEntidades.ID " & _
            "WHERE (((TbUsuarios.ID) Is Null));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    If rcdDatos.EOF Then
        rcdDatos.Close
        Set rcdDatos = Nothing
        Exit Property
    End If
    Do While Not rcdDatos.EOF
        m_ID = rcdDatos.Fields("ID")
        If m_ObjUsuariosEnEntidadNoExistentes Is Nothing Then
            Set m_ObjUsuariosEnEntidadNoExistentes = New Scripting.Dictionary
            m_ObjUsuariosEnEntidadNoExistentes.CompareMode = TextCompare
        End If
        If Not m_ObjUsuariosEnEntidadNoExistentes.Exists(m_ID) Then
            m_ObjUsuariosEnEntidadNoExistentes.Add m_ID, m_ID
        End If
        rcdDatos.MoveNext
    Loop
    
    
    
    Set UsuariosEnEntidadNoExistentes = m_ObjUsuariosEnEntidadNoExistentes
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.UsuariosEnEntidadNoExistentes ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set UsuariosEnEntidadNoExistentes(p_Valor As Variant)
    
    Set m_ObjUsuariosEnEntidadNoExistentes = p_Valor
    
End Property

Public Property Get UsuariosNoEnEntidad() As Scripting.Dictionary
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_ID As String
    
    On Error GoTo errores
    Me.Error = ""
    If Not m_ObjUsuariosNoEnEntidad Is Nothing Then
        Set UsuariosNoEnEntidad = m_ObjUsuariosNoEnEntidad
        Exit Property
    End If
    
   
    m_SQL = "SELECT TbUsuarios.ID " & _
            "FROM TbUsuarios LEFT JOIN TbUsuariosEntidades ON TbUsuarios.ID = TbUsuariosEntidades.ID " & _
            "WHERE (((TbUsuariosEntidades.ID) Is Null));"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    If rcdDatos.EOF Then
        rcdDatos.Close
        Set rcdDatos = Nothing
        Exit Property
    End If
    Do While Not rcdDatos.EOF
        m_ID = rcdDatos.Fields("ID")
        If m_ObjUsuariosNoEnEntidad Is Nothing Then
            Set m_ObjUsuariosNoEnEntidad = New Scripting.Dictionary
            m_ObjUsuariosNoEnEntidad.CompareMode = TextCompare
        End If
        If Not m_ObjUsuariosNoEnEntidad.Exists(m_ID) Then
            m_ObjUsuariosNoEnEntidad.Add m_ID, m_ID
        End If
        rcdDatos.MoveNext
    Loop
    
    Set UsuariosNoEnEntidad = m_ObjUsuariosNoEnEntidad
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.UsuariosNoEnEntidad ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set UsuariosNoEnEntidad(p_Valor As Variant)
    
    Set m_ObjUsuariosNoEnEntidad = p_Valor
    
End Property
Public Property Get DiasEnSegundaConvocatoriaSinCurso() As Long
    
    On Error GoTo errores
    Me.Error = ""
    
    If m_lDiasEnSegundaConvocatoriaSinCurso <> 0 Then
        DiasEnSegundaConvocatoriaSinCurso = m_lDiasEnSegundaConvocatoriaSinCurso
        Exit Property
    End If
    m_lDiasEnSegundaConvocatoriaSinCurso = 7
    
    DiasEnSegundaConvocatoriaSinCurso = m_lDiasEnSegundaConvocatoriaSinCurso
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.DiasEnSegundaConvocatoriaSinCurso ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get DiasEnPrimeraConvocatoriaSinCurso() As Long
    
    On Error GoTo errores
    Me.Error = ""
    
    If m_lDiasEnPrimeraConvocatoriaSinCurso <> 0 Then
        DiasEnPrimeraConvocatoriaSinCurso = m_lDiasEnPrimeraConvocatoriaSinCurso
        Exit Property
    End If
    m_lDiasEnPrimeraConvocatoriaSinCurso = 15
    
    DiasEnPrimeraConvocatoriaSinCurso = m_lDiasEnPrimeraConvocatoriaSinCurso
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.DiasEnPrimeraConvocatoriaSinCurso ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get DiasSinCorreoNoCurso() As Long
    
    On Error GoTo errores
    Me.Error = ""
    
    If m_lDiasSinCorreoNoCurso <> 0 Then
        DiasSinCorreoNoCurso = m_lDiasSinCorreoNoCurso
        Exit Property
    End If
    m_lDiasSinCorreoNoCurso = 15
    
    DiasSinCorreoNoCurso = m_lDiasSinCorreoNoCurso
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.DiasSinCorreoNoCurso ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get colUsuariosPermitidosEnLocal() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    If Not m_objcolUsuariosPermitidosEnLocal Is Nothing Then
        Set colUsuariosPermitidosEnLocal = m_objcolUsuariosPermitidosEnLocal
        Exit Property
    End If
    Set m_objcolUsuariosPermitidosEnLocal = CargarUsuariosPermitidosEnLocal(, m_Error)
    If m_Error <> "" Then
        Me.Error = m_Error
    End If
    
    Set colUsuariosPermitidosEnLocal = m_objcolUsuariosPermitidosEnLocal
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.colUsuariosPermitidosEnLocal ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Function CargarUsuariosPermitidosEnLocal( _
                                Optional ByRef p_Db As DAO.Database = Nothing, _
                                Optional ByRef p_Error As String _
                                ) As Scripting.Dictionary
    Dim db As DAO.Database
    Dim rcdDatos As DAO.Recordset
    Dim colUsuarios As Scripting.Dictionary
    Dim m_CampoCorreo As String
    Dim m_CampoActivo As String
    Dim m_SQL As String
    Dim m_Correo As String
    Dim m_Activo As Boolean
    Dim m_TablaCreada As Boolean
    
    On Error GoTo errores
    p_Error = ""
    Me.Error = ""
    Set colUsuarios = New Scripting.Dictionary
    colUsuarios.CompareMode = TextCompare
    
    If p_Db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then
            p_Error = "No se pudo abrir el backend HPST.accdb para leer TbConfiguracionUsuariosAccesoLocal: " & p_Error & ". Se deniega el acceso local por defecto."
            Set CargarUsuariosPermitidosEnLocal = colUsuarios
            GoTo SALIR
        End If
    Else
        Set db = p_Db
    End If
    If db Is Nothing Then
        p_Error = "No se pudo abrir el backend HPST.accdb para leer TbConfiguracionUsuariosAccesoLocal. Se deniega el acceso local por defecto."
        Set CargarUsuariosPermitidosEnLocal = colUsuarios
        GoTo SALIR
    End If
    
    If Not ExisteTabla(db, "TbConfiguracionUsuariosAccesoLocal") Then
        If Not CrearTablaUsuariosPermitidosEnLocal(db, p_Error) Then
            Set CargarUsuariosPermitidosEnLocal = colUsuarios
            GoTo SALIR
        End If
        m_TablaCreada = True
    End If
    
    If m_TablaCreada Then
        If Not SembrarUsuariosPermitidosEnLocalIniciales(db, p_Error) Then
            Set CargarUsuariosPermitidosEnLocal = colUsuarios
            GoTo SALIR
        End If
    End If
    
    m_CampoCorreo = NombreCampoExistente(db, "TbConfiguracionUsuariosAccesoLocal", "CorreoUsuario|Correo|Email|CorreoElectronico")
    If m_CampoCorreo = "" Then
        p_Error = "La tabla TbConfiguracionUsuariosAccesoLocal no tiene columna de correo reconocida. Se deniega el acceso local por defecto."
        Set CargarUsuariosPermitidosEnLocal = colUsuarios
        GoTo SALIR
    End If
    m_CampoActivo = NombreCampoExistente(db, "TbConfiguracionUsuariosAccesoLocal", "Activo|Habilitado|Enabled")
    
    m_SQL = "SELECT * FROM TbConfiguracionUsuariosAccesoLocal"
    Set rcdDatos = db.OpenRecordset(m_SQL, dbOpenSnapshot)
    Do While Not rcdDatos.EOF
        m_Correo = Trim$(Nz(rcdDatos.Fields(m_CampoCorreo).value, ""))
        m_Activo = True
        If m_CampoActivo <> "" Then
            m_Activo = CBool(Nz(rcdDatos.Fields(m_CampoActivo).value, False))
        End If
        If m_Correo <> "" Then
            If m_Activo Then
                If Not colUsuarios.Exists(m_Correo) Then
                    colUsuarios.Add m_Correo, m_Correo
                End If
            End If
        End If
        rcdDatos.MoveNext
    Loop
    
    If colUsuarios.Count = 0 Then
        p_Error = "La tabla TbConfiguracionUsuariosAccesoLocal no contiene usuarios activos. Se deniega el acceso local por defecto."
    End If
    
    Set CargarUsuariosPermitidosEnLocal = colUsuarios

SALIR:
    Me.Error = p_Error
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    If p_Db Is Nothing Then
        Set db = Nothing
    End If
    Exit Function
errores:
    p_Error = "El método Entorno.CargarUsuariosPermitidosEnLocal ha devuelto el error: " & vbNewLine & Err.Description & ". Se deniega el acceso local por defecto."
    Set CargarUsuariosPermitidosEnLocal = colUsuarios
    Resume SALIR
End Function

Private Function CrearTablaUsuariosPermitidosEnLocal(ByRef p_Db As DAO.Database, ByRef p_Error As String) As Boolean
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index

    On Error GoTo errores
    p_Error = ""

    Set tdf = p_Db.CreateTableDef("TbConfiguracionUsuariosAccesoLocal")

    Set fld = tdf.CreateField("ID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("CorreoUsuario", dbText, 255)
    fld.Required = True
    fld.AllowZeroLength = False
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("Activo", dbBoolean)
    tdf.Fields.Append fld


    Set fld = tdf.CreateField("FechaAlta", dbDate)
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("Observaciones", dbMemo)
    tdf.Fields.Append fld


    Set idx = tdf.CreateIndex("pk_TbConfiguracionUsuariosAccesoLocal")
    idx.Primary = True
    idx.Unique = True
    idx.Fields.Append idx.CreateField("ID")
    tdf.Indexes.Append idx

    Set idx = tdf.CreateIndex("ux_TbConfiguracionUsuariosAccesoLocal_CorreoUsuario")
    idx.Unique = True
    idx.Fields.Append idx.CreateField("CorreoUsuario")
    tdf.Indexes.Append idx

    p_Db.TableDefs.Append tdf
    p_Db.TableDefs.Refresh

    CrearTablaUsuariosPermitidosEnLocal = True
    Exit Function
errores:
    p_Error = "No se pudo crear la tabla TbConfiguracionUsuariosAccesoLocal en el backend HPST.accdb: " & Err.Description & ". Se deniega el acceso local por defecto."
    CrearTablaUsuariosPermitidosEnLocal = False
End Function

Private Function SembrarUsuariosPermitidosEnLocalIniciales(ByRef p_Db As DAO.Database, ByRef p_Error As String) As Boolean
    On Error GoTo errores
    p_Error = ""
    
    SembrarUsuarioPermitidoEnLocalSiNoExiste p_Db, "esperanza.delalamoarriba@telefonica.com"
    SembrarUsuarioPermitidoEnLocalSiNoExiste p_Db, "andres.romandelperal@telefonica.com"
    SembrarUsuarioPermitidoEnLocalSiNoExiste p_Db, "martina.torralbarodriguez@telefonica.com"
    SembrarUsuarioPermitidoEnLocalSiNoExiste p_Db, "almudena.cardenasvelloso.ext@telefonica.com"
    
    SembrarUsuariosPermitidosEnLocalIniciales = True
    Exit Function
errores:
    p_Error = "No se pudieron sembrar los usuarios iniciales en TbConfiguracionUsuariosAccesoLocal: " & Err.Description & ". Se deniega el acceso local por defecto."
    SembrarUsuariosPermitidosEnLocalIniciales = False
End Function

Private Sub SembrarUsuarioPermitidoEnLocalSiNoExiste(ByRef p_Db As DAO.Database, ByVal p_CorreoUsuario As String)
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim m_ErrNumber As Long
    Dim m_ErrDescription As String
    
    On Error GoTo errores
    Set qdf = p_Db.CreateQueryDef(vbNullString, _
        "PARAMETERS p_Correo TEXT(255); " & _
        "SELECT ID FROM TbConfiguracionUsuariosAccesoLocal WHERE CorreoUsuario = [p_Correo]")
    qdf.Parameters("p_Correo").value = p_CorreoUsuario
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    If rs.EOF Then
        p_Db.Execute "INSERT INTO TbConfiguracionUsuariosAccesoLocal (CorreoUsuario, Activo, FechaAlta) " & _
                     "VALUES ('" & Replace(p_CorreoUsuario, "'", "''") & "', True, Now())", dbFailOnError
    End If
    
SALIR:
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set qdf = Nothing
    If m_ErrNumber <> 0 Then
        Err.Raise m_ErrNumber, "SembrarUsuarioPermitidoEnLocalSiNoExiste", m_ErrDescription
    End If
    Exit Sub
errores:
    m_ErrNumber = Err.Number
    m_ErrDescription = Err.Description
    Resume SALIR
End Sub

Private Function ExisteTabla(ByRef p_Db As DAO.Database, ByVal p_NombreTabla As String) As Boolean
    Dim tdf As DAO.TableDef
    
    On Error GoTo errores
    p_Db.TableDefs.Refresh
    For Each tdf In p_Db.TableDefs
        If StrComp(tdf.Name, p_NombreTabla, vbTextCompare) = 0 Then
            ExisteTabla = True
            Exit Function
        End If
    Next tdf
    Exit Function
errores:
    ExisteTabla = False
End Function

Private Function NombreCampoExistente(ByRef p_Db As DAO.Database, ByVal p_NombreTabla As String, ByVal p_Candidatos As String) As String
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim arrCandidatos() As String
    Dim i As Long
    
    On Error GoTo errores
    arrCandidatos = Split(p_Candidatos, "|")
    p_Db.TableDefs.Refresh
    Set tdf = p_Db.TableDefs(p_NombreTabla)
    For i = LBound(arrCandidatos) To UBound(arrCandidatos)
        For Each fld In tdf.Fields
            If StrComp(fld.Name, arrCandidatos(i), vbTextCompare) = 0 Then
                NombreCampoExistente = fld.Name
                Exit Function
            End If
        Next fld
    Next i
    Exit Function
errores:
    NombreCampoExistente = ""
End Function

Public Property Get UsuarioConectadoMaquina() As usuario
    
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_ObjUsuarioConectadoMaquina Is Nothing Then
        Set UsuarioConectadoMaquina = m_ObjUsuarioConectadoMaquina
        Exit Property
    End If
    Set m_ObjUsuarioConectadoMaquina = Constructor.getUsuarioConectadoPorMaquina(m_Error)
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    Set UsuarioConectadoMaquina = m_ObjUsuarioConectadoMaquina
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.UsuarioConectadoMaquina ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error


End Property

Public Property Get ColEmpresas() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    
    If Not m_objColEmpresas Is Nothing Then
        Set ColEmpresas = m_objColEmpresas
        Exit Property
    End If
    Set m_objColEmpresas = Constructor.getSuministradores(p_SoloTramitadorasHPS:=EnumSiNo.No, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColEmpresas = m_objColEmpresas
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColEmpresas ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Set ColEmpresas(ByVal objNewValue As Variant)

    Set m_objColEmpresas = objNewValue

End Property

Public Property Get ColEmpresasTramitadoras() As Scripting.Dictionary

    On Error GoTo errores
    Me.Error = ""
    
    If Not m_objColEmpresasTramitadoras Is Nothing Then
        Set ColEmpresasTramitadoras = m_objColEmpresasTramitadoras
        Exit Property
    End If
    Set m_objColEmpresasTramitadoras = Constructor.getSuministradores(p_SoloTramitadorasHPS:=EnumSiNo.Sí, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColEmpresasTramitadoras = m_objColEmpresasTramitadoras
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColEmpresasTramitadoras ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error

End Property

Public Property Set ColEmpresasTramitadoras(ByVal objNewValue As Scripting.Dictionary)

    Set m_objColEmpresasTramitadoras = objNewValue

End Property
Public Property Get ColCadenaContratistas() As Scripting.Dictionary

    On Error GoTo errores
    Me.Error = ""
    
    If Not m_objColCadenaContratistas Is Nothing Then
        Set ColCadenaContratistas = m_objColCadenaContratistas
        Exit Property
    End If
    Set m_objColCadenaContratistas = Constructor.gelCadenaContratistas(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColCadenaContratistas = m_objColCadenaContratistas
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColCadenaContratistas ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error

End Property

Public Property Set ColCadenaContratistas(ByVal objNewValue As Scripting.Dictionary)

    Set m_objColCadenaContratistas = objNewValue

End Property
Public Property Get colColoresMaster() As Scripting.Dictionary
    
    On Error GoTo errores
    Me.Error = ""
    
    If Not m_objcolColoresMaster Is Nothing Then
        Set colColoresMaster = m_objcolColoresMaster
        Exit Property
    End If
    Set m_objcolColoresMaster = New Scripting.Dictionary
    m_objcolColoresMaster.CompareMode = TextCompare
    With m_objcolColoresMaster
        .Add "CabeceraBackColor", 3969911
        .Add "ListaBackColor", 10209467
        .Add "ListaForeColor", 4210752
        .Add "ListaBorderColor", 16777215
        .Add "EtiquetaBackColor", 8107685
        .Add "EtiquetaLinkBackColorSeleccionado", 16777215
        .Add "EtiquetaForeColor", 4210752
        .Add "EtiquetaForeColorSeleccionado", 8210719
        .Add "EtiquetaBorderColor", 16777215
        .Add "EtiquetaLinkForeColor", 16777215
        .Add "CampoBackColor", 14544104
        .Add "CampoForeColor", 4210752
        .Add "CampoBorderColor", 16777215
        
        
        
        
    End With
    
    If Not m_objcolColoresMaster Is Nothing Then
        Set colColoresMaster = m_objcolColoresMaster
        Exit Property
    End If
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.colColoresMaster ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error

End Property

Public Property Set colColoresMaster(ByVal objNewValue As Scripting.Dictionary)

    Set m_objcolColoresMaster = objNewValue

End Property

Public Property Get colColoresSlave() As Scripting.Dictionary

    On Error GoTo errores
    Me.Error = ""
    
    If Not m_objcolColoresSlave Is Nothing Then
        Set colColoresSlave = m_objcolColoresSlave
        Exit Property
    End If
    Set m_objcolColoresSlave = New Scripting.Dictionary
    m_objcolColoresSlave.CompareMode = TextCompare
    With m_objcolColoresSlave
        .Add "CabeceraBackColor", 8210719
        .Add "ListaBackColor", 15720403
        .Add "ListaForeColor", 4210752
        .Add "ListaBorderColor", 16777215
        .Add "EtiquetaBackColor", 14136213
        .Add "EtiquetaLinkBackColorSeleccionado", 16777215
        .Add "EtiquetaForeColor", 4210752
        .Add "EtiquetaForeColorSeleccionado", 8210719
        .Add "EtiquetaBorderColor", 16777215
        .Add "EtiquetaLinkForeColor", 16777215
        .Add "CampoBackColor", 16446961
        .Add "CampoForeColor", 4210752
        .Add "CampoBorderColor", 16777215
    End With
    
    If Not m_objcolColoresSlave Is Nothing Then
        Set colColoresSlave = m_objcolColoresSlave
        Exit Property
    End If
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.colColoresSlave ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error

End Property

Public Property Get CabeceraBackColor() As Long
    
    On Error GoTo errores
    Me.Error = ""
    If Application.TempVars("SoloLectura") = "Sí" Then
        m_lCabeceraBackColor = Me.colColoresSlave("CabeceraBackColor")
    Else
        m_lCabeceraBackColor = Me.colColoresMaster("CabeceraBackColor")
    End If
    
    
    CabeceraBackColor = m_lCabeceraBackColor
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.CabeceraBackColor ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get ListaBackColor() As Long

    On Error GoTo errores
    Me.Error = ""
    
    If Application.TempVars("SoloLectura") = "Sí" Then
        m_lListaBackColor = Me.colColoresSlave("ListaBackColor")
    Else
        m_lListaBackColor = Me.colColoresMaster("ListaBackColor")
    End If
    
    
    ListaBackColor = m_lListaBackColor
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ListaBackColor ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get ListaForeColor() As Long

    On Error GoTo errores
    Me.Error = ""
    
    If Application.TempVars("SoloLectura") = "Sí" Then
        m_lListaForeColor = Me.colColoresSlave("ListaForeColor")
    Else
       m_lListaForeColor = Me.colColoresMaster("ListaForeColor")
    End If
    
   
    ListaForeColor = m_lListaForeColor
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ListaForeColor ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get ListaBorderColor() As Long

    On Error GoTo errores
    Me.Error = ""
    
    If Application.TempVars("SoloLectura") = "Sí" Then
        m_lListaBorderColor = Me.colColoresSlave("ListaBorderColor")
    Else
        m_lListaBorderColor = Me.colColoresMaster("ListaBorderColor")
    End If
    
    ListaBorderColor = m_lListaBorderColor
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ListaBorderColor ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get EtiquetaBackColor() As Long

    On Error GoTo errores
    Me.Error = ""
    
    If Application.TempVars("SoloLectura") = "Sí" Then
        m_lEtiquetaBackColor = Me.colColoresSlave("EtiquetaBackColor")
    Else
        m_lEtiquetaBackColor = Me.colColoresMaster("EtiquetaBackColor")
    End If
        
    EtiquetaBackColor = m_lEtiquetaBackColor
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.EtiquetaBackColor ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get EtiquetaLinkBackColorSeleccionado() As Long

    On Error GoTo errores
    Me.Error = ""
    
    If Application.TempVars("SoloLectura") = "Sí" Then
        m_lEtiquetaLinkBackColorSeleccionado = Me.colColoresSlave("EtiquetaLinkBackColorSeleccionado")
    Else
        m_lEtiquetaLinkBackColorSeleccionado = Me.colColoresMaster("EtiquetaLinkBackColorSeleccionado")
    End If
    
    EtiquetaLinkBackColorSeleccionado = m_lEtiquetaLinkBackColorSeleccionado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.EtiquetaLinkBackColorSeleccionado ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get EtiquetaLinkForeColor() As Long

    On Error GoTo errores
    Me.Error = ""
    
    If Application.TempVars("SoloLectura") = "Sí" Then
         m_lEtiquetaLinkForeColor = Me.colColoresSlave("EtiquetaLinkForeColor")
    Else
        m_lEtiquetaLinkForeColor = Me.colColoresMaster("EtiquetaLinkForeColor")
    End If
    
    EtiquetaLinkForeColor = m_lEtiquetaLinkForeColor
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.EtiquetaLinkForeColor ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property


Public Property Get EtiquetaForeColor() As Long

    On Error GoTo errores
    Me.Error = ""
    
    If Application.TempVars("SoloLectura") = "Sí" Then
         m_lEtiquetaForeColor = Me.colColoresSlave("EtiquetaForeColor")
    Else
        m_lEtiquetaForeColor = Me.colColoresMaster("EtiquetaForeColor")
    End If
        
    EtiquetaForeColor = m_lEtiquetaForeColor
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.EtiquetaForeColor ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get EtiquetaForeColorSeleccionado() As Long

    On Error GoTo errores
    Me.Error = ""
    
    
    If Application.TempVars("SoloLectura") = "Sí" Then
         m_lEtiquetaForeColorSeleccionado = Me.colColoresSlave("EtiquetaForeColorSeleccionado")
    Else
         m_lEtiquetaForeColorSeleccionado = Me.colColoresMaster("EtiquetaForeColorSeleccionado")
    End If
    
    
    EtiquetaForeColorSeleccionado = m_lEtiquetaForeColorSeleccionado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.EtiquetaForeColorSeleccionado ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get EtiquetaBorderColor() As Long

    On Error GoTo errores
    Me.Error = ""
    
    If Application.TempVars("SoloLectura") = "Sí" Then
         m_lEtiquetaBorderColor = Me.colColoresMaster("EtiquetaBorderColor")
    Else
         m_lEtiquetaBorderColor = Me.colColoresSlave("EtiquetaBorderColor")
    End If
    
    EtiquetaBorderColor = m_lEtiquetaBorderColor
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.EtiquetaBorderColor ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get EtiquetaSeleccionadaForeColor() As Long

    On Error GoTo errores
    Me.Error = ""
    
    If Application.TempVars("SoloLectura") = "Sí" Then
         m_lEtiquetaSeleccionadaForeColor = Me.colColoresSlave("EtiquetaSeleccionadaForeColor")
    Else
        m_lEtiquetaSeleccionadaForeColor = Me.colColoresMaster("EtiquetaSeleccionadaForeColor")
    End If
    
    
    EtiquetaSeleccionadaForeColor = m_lEtiquetaSeleccionadaForeColor
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.EtiquetaSeleccionadaForeColor ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property
Public Property Get EtiquetaSeleccionadaBackColor() As Long

    On Error GoTo errores
    Me.Error = ""
    
    If Application.TempVars("SoloLectura") = "Sí" Then
        m_lEtiquetaSeleccionadaBackColor = Me.colColoresSlave("EtiquetaSeleccionadaBackColor")
    Else
         m_lEtiquetaSeleccionadaBackColor = Me.colColoresMaster("EtiquetaSeleccionadaBackColor")
    End If
    
    
    EtiquetaSeleccionadaBackColor = m_lEtiquetaSeleccionadaBackColor
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.EtiquetaSeleccionadaBackColor ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get CampoBackColor() As Long

    On Error GoTo errores
    Me.Error = ""
    
    If Application.TempVars("SoloLectura") = "Sí" Then
        m_lCampoBackColor = Me.colColoresSlave("CampoBackColor")
    Else
         m_lCampoBackColor = Me.colColoresMaster("CampoBackColor")
    End If
    
    
    CampoBackColor = m_lCampoBackColor
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.CampoBackColor ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get CampoForeColor() As Long

    On Error GoTo errores
    Me.Error = ""
    
    If Application.TempVars("SoloLectura") = "Sí" Then
        m_lCampoForeColor = Me.colColoresSlave("CampoForeColor")
    Else
         m_lCampoForeColor = Me.colColoresMaster("CampoForeColor")
    End If
    
    
    CampoForeColor = m_lCampoForeColor
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.CampoForeColor ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get CampoBorderColor() As Long

    On Error GoTo errores
    Me.Error = ""
    
    If Application.TempVars("SoloLectura") = "Sí" Then
        m_lCampoBorderColor = Me.colColoresSlave("CampoBorderColor")
    Else
          m_lCampoBorderColor = Me.colColoresMaster("CampoBorderColor")
    End If
   
    
    CampoBorderColor = m_lCampoBorderColor
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.CampoBorderColor ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property




Public Property Get URLCarpetaAnexos() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sURLCarpetaAnexos <> "" Then
        URLCarpetaAnexos = m_sURLCarpetaAnexos
        Exit Property
    End If
    ' Las rutas de TbConfiguracionHPS solo deben gobernar el modo local
    ' explicito. En modo remoto/VPN (DatosEnLocal <> "Sí") usamos rutas
    ' reales remotas cuando hay acceso a DatosTE, aunque no estemos en oficina.
    m_sURLCarpetaAnexos = ResolveURLCarpetaAnexosCandidate(False)

    ' [anexos subcarpetas] NO validamos fso.FolderExists aca — la URL viene
    ' de env vars + DNI + subfolder (Actual/Historico). La existencia del
    ' directorio se valida cuando el codigo de negocio intenta leer o
    ' escribir el archivo, no aca. Si fso-checkea aqui, un entorno de
    ' test con env vars validas pero sin el folder fisico creado rompe
    ' el helper por razones ajenas a la logica de anexos.
    URLCarpetaAnexos = m_sURLCarpetaAnexos
    Exit Property
errores:
    m_Error = "El método Entorno.URLCarpetaAnexos ha devuelto el error: " & vbNewLine & Err.Description
    CorreoAlAdministrador m_Error
    Me.Error = m_Error
End Property

Public Property Get URLCarpetaAnexosHistoricos() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sURLCarpetaAnexosHistoricos <> "" Then
        URLCarpetaAnexosHistoricos = m_sURLCarpetaAnexosHistoricos
        Exit Property
    End If
    ' Mismo criterio que URLCarpetaAnexos — sin fs-check.
    m_sURLCarpetaAnexosHistoricos = ResolveURLCarpetaAnexosCandidate(True)

    URLCarpetaAnexosHistoricos = m_sURLCarpetaAnexosHistoricos
    Exit Property
errores:
    m_Error = "El método Entorno.URLCarpetaAnexosHistoricos ha devuelto el error: " & vbNewLine & Err.Description
    CorreoAlAdministrador m_Error
    Me.Error = m_Error
End Property

Public Property Get URLCarpetaRecursos() As String
    
    On Error GoTo errores
    Me.Error = ""
    If m_sURLCarpetaRecursos <> "" Then
        URLCarpetaRecursos = m_sURLCarpetaRecursos
        Exit Property
    End If
    If Application.TempVars("DatosEnLocal") = "Sí" Then
        m_sURLCarpetaRecursos = "C:\OneDrive\Telefonica\Aplicaciones_dys.TMETF - Aplicaciones PpD\HPS\Recursos\"
    Else
        m_sURLCarpetaRecursos = "\\datoste\aplicaciones_dys\Aplicaciones PpD\HPS\Recursos\"
    End If
    
    If Not fso.FolderExists(m_sURLCarpetaRecursos) Then
        m_Error = "La carpeta m_sURLCarpetaRecursos no es alcanzable " & vbNewLine & m_sURLCarpetaRecursos
        m_sURLCarpetaRecursos = ""
        Err.Raise 1000
    End If
    
    URLCarpetaRecursos = m_sURLCarpetaRecursos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLCarpetaRecursos ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error

End Property


Public Property Get URLCarpetaAnexosSICA() As String
    On Error GoTo errores
    Me.Error = ""
    If m_sURLCarpetaAnexosSICA <> "" Then
        URLCarpetaAnexosSICA = m_sURLCarpetaAnexosSICA
        Exit Property
    End If
    
    ' --- Mismo patrón que ResolveURLCarpetaAnexosCandidate ---
    ' [config-as-source-of-truth] APP_ROOT_REMOTO si \\datoste está
    ' alcanzable, APP_ROOT_LOCAL si no.
    Dim sSicaRoot As String
    Dim sSicaRemoto As String
    Dim sSicaLocal As String
    sSicaRemoto = GetConfigOrFail(k_Key_AppRootRemoto)
    sSicaLocal = GetConfigOrFail(k_Key_AppRootLocal)
    If fso.FolderExists(sSicaRemoto) Then
        sSicaRoot = sSicaRemoto
    ElseIf fso.FolderExists(sSicaLocal) Then
        sSicaRoot = sSicaLocal
    Else
        m_Error = "URLCarpetaAnexosSICA: ni APP_ROOT_REMOTO ('" & sSicaRemoto & _
            "') ni APP_ROOT_LOCAL ('" & sSicaLocal & "') son alcanzables."
        Err.Raise 1000
    End If
    m_sURLCarpetaAnexosSICA = EnsureSlash(sSicaRoot)
    
    ' Validación y creación si no existe (solo fuera de oficina)
    If Not fso.FolderExists(m_sURLCarpetaAnexosSICA) Then
        If Not fso.FolderExists(fso.GetParentFolderName(m_sURLCarpetaAnexosSICA)) Then
            m_Error = "La carpeta de anexos SICA no es alcanzable " & vbNewLine & m_sURLCarpetaAnexosSICA
            Err.Raise 1000
        End If
        fso.CreateFolder m_sURLCarpetaAnexosSICA
    End If
    
    URLCarpetaAnexosSICA = m_sURLCarpetaAnexosSICA
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLCarpetaAnexosSICA ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get MesesParaRenovacion() As Integer
    
    On Error GoTo errores
    Me.Error = ""
    
    If Not IsNumeric(Application.TempVars("MesesParaRenovacion")) Then
        Application.TempVars("MesesParaRenovacion") = 9
    End If
    m_iMesesParaRenovacion = Application.TempVars("MesesParaRenovacion")
    MesesParaRenovacion = m_iMesesParaRenovacion
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.MesesParaRenovacion ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get URLCarpetaAplicacionesDys() As String
     
    On Error GoTo errores
    
    Me.Error = ""
    If m_sURLCarpetaAplicacionesDys <> "" Then
        URLCarpetaAplicacionesDys = m_sURLCarpetaAplicacionesDys
        Exit Property
    End If
    ' [config-as-source-of-truth] Resolver estricto: APP_ROOT_LOCAL o
    ' APP_ROOT_REMOTO desde TbConfiguracionHPS. No usa TempVars, no usa
    ' CurrentProject.path, no usa m_URLRutaAplicacionesRemotas (que es
    ' una constante hardcoded que también se va a eliminar).
    If Application.TempVars("DatosEnLocal") = "Sí" Then
        m_sURLCarpetaAplicacionesDys = GetConfigOrFail(k_Key_AppRootLocal)
    Else
        m_sURLCarpetaAplicacionesDys = GetConfigOrFail(k_Key_AppRootRemoto)
    End If
    If Not fso.FolderExists(m_sURLCarpetaAplicacionesDys) Then
        m_Error = "No es alcanzable m_sURLCarpetaAplicacionesDys " & vbNewLine & m_sURLCarpetaAplicacionesDys
        m_sURLCarpetaAplicacion = ""
        Err.Raise 1000
    End If
    
    URLCarpetaAplicacionesDys = m_sURLCarpetaAplicacionesDys
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLCarpetaAplicacionesDys ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property



Public Property Get URLCarpetaAplicacion() As String
     
    On Error GoTo errores
    
    Me.Error = ""
    If m_sURLCarpetaAplicacion <> "" Then
        URLCarpetaAplicacion = m_sURLCarpetaAplicacion
        Exit Property
    End If
    ' [config-as-source-of-truth] Resolver estricto: APP_ROOT_LOCAL/REMOTO
    ' desde TbConfiguracionHPS, con sufijo "HPS\". Sin TempVars.
    If Application.TempVars("DatosEnLocal") = "Sí" Then
        m_sURLCarpetaAplicacion = GetConfigOrFail(k_Key_AppRootLocal) & "HPS\"
    Else
        m_sURLCarpetaAplicacion = GetConfigOrFail(k_Key_AppRootRemoto) & "HPS\"
    End If
    If Not fso.FolderExists(m_sURLCarpetaAplicacion) Then
        m_Error = "No es alcanzable m_sURLCarpetaAplicacion " & vbNewLine & m_sURLCarpetaAplicacion
        m_sURLCarpetaAplicacion = ""
        Err.Raise 1000
    End If
    
    URLCarpetaAplicacion = m_sURLCarpetaAplicacion
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLCarpetaAplicacion ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property



Public Property Get URLDirectorioLocal() As String

    On Error GoTo errores
    Me.Error = ""
    If m_sURLDirectorioLocal <> "" Then
        URLDirectorioLocal = m_sURLDirectorioLocal
        Exit Property
    End If
    m_sURLDirectorioLocal = Environ("APPDATA") & "\" & "Aplicaciones DYSN" & "\HPS\"
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not fso.FolderExists(m_sURLDirectorioLocal) Then
        fso.CreateFolder m_sURLDirectorioLocal
    End If
    m_sURLDirectorioLocal = m_sURLDirectorioLocal & "\"
    URLDirectorioLocal = m_sURLDirectorioLocal
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método General.URLDirectorioLocal ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property


Public Property Get URLIni() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    
    If m_sURLIni <> "" Then
        URLIni = m_sURLIni
        Exit Property
    End If
    If m_EnOficina = EnumSiNo.No Then
        Exit Property
    End If
    If Application.TempVars("DatosEnLocal") = "Sí" Then
        m_sURLIni = "C:\OneDrive\Telefonica\Aplicaciones_dys.TMETF - Aplicaciones PpD\HPS\ANEXOS\HPS"
    Else
        m_sURLIni = m_ObjEntorno.URLCarpetaRecursos & "HPS.ini"
    End If
   
    If Not fso.FileExists(m_sURLIni) Then
        m_Error = "No es alcanzable el archivo m_sURLIni: " & m_sURLIni
        m_sURLIni = ""
        Err.Raise 1000
    End If
    URLIni = m_sURLIni
       
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLIni ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get VersionAplicacion() As String
    
    On Error GoTo errores
    
    Me.Error = ""
     
    If m_sVersionAplicacion <> "" Then
        VersionAplicacion = m_sVersionAplicacion
        Exit Property
    End If
    If m_EnOficina = EnumSiNo.Sí Then
    
        m_sVersionAplicacion = LeerIni("VersionAplicacion", "")
        If m_sVersionAplicacion = "" Then
            m_Error = "No se puede determinar la versión de la aplicación"
            Err.Raise 1000
        End If
    Else
        m_sVersionAplicacion = Application.TempVars("VersionAplicacion")
    End If
    If m_sVersionAplicacion <> "" Then
        VersionAplicacion = m_sVersionAplicacion
        Exit Property
    End If
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.VersionAplicacion ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get ColEstadosVisiblesHPS() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    
    If Not m_objColEstadosVisiblesHPS Is Nothing Then
        Set ColEstadosVisiblesHPS = m_objColEstadosVisiblesHPS
        Exit Property
    End If
    Set m_objColEstadosVisiblesHPS = New Scripting.Dictionary
    m_objColEstadosVisiblesHPS.CompareMode = TextCompare
    With m_objColEstadosVisiblesHPS
        .Add "1", "Activo" '--->APuntoCaducarRenovacionNo+APuntoCaducarRenovacionSiSinSolicitud+APuntoCaducarRenovacionSiConSolicitud
        .Add "2", "A punto de Caducar" '---->APuntoCaducarRenovacionSiSinSolicitud+APuntoCaducarRenovacionSiConSolicitud
        .Add "3", "Caducada" '----->CaducadaRenovacionNo+CaducadaRenovacionSiSinSolicitud+CaducadaRenovacionSiConSolicitud
        .Add "4", "Baja"
        
        .Add "5", "Solicitada" '---->ActivoSolicitada+ActivoSinSolicitar+APuntoCaducarRenovacionSiConSolicitud+CaducadaRenovacionSiConSolicitud
        .Add "6", "Irregular"
        .Add "7", "Sin Orden de Renovación"
        .Add "8", "No Posee"
        .Add "9", "Pendiente de Renovación"
    End With
    If Not m_objColEstadosVisiblesHPS Is Nothing Then
        Set ColEstadosVisiblesHPS = m_objColEstadosVisiblesHPS
        Exit Property
    End If
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColEstadosVisiblesHPS ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
    
End Property
Public Property Get ColVisiblesHPSEstados() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    
    If Not m_objColVisiblesHPSEstados Is Nothing Then
        Set ColVisiblesHPSEstados = m_objColVisiblesHPSEstados
        Exit Property
    End If
    Set m_objColVisiblesHPSEstados = New Scripting.Dictionary
    m_objColVisiblesHPSEstados.CompareMode = TextCompare
    With m_objColVisiblesHPSEstados
        .Add "Activo", 1
        .Add "A punto de Caducar", 2
        .Add "Caducada", 3
        .Add "Baja", 4
        
        .Add "Solicitada", 5
        .Add "Irregular", 6
        .Add "Sin Orden de Renovación", 7
        .Add "No Posee", 8
        .Add "Pendiente de Renovación", 9
    End With
    If Not m_objColVisiblesHPSEstados Is Nothing Then
        Set ColVisiblesHPSEstados = m_objColVisiblesHPSEstados
        Exit Property
    End If
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColVisiblesHPSEstados ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
    
End Property
Public Property Get ColEstadosHPS() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    
    If Not m_objColEstadosHPS Is Nothing Then
        Set ColEstadosHPS = m_objColEstadosHPS
        Exit Property
    End If
    Set m_objColEstadosHPS = New Scripting.Dictionary
    m_objColEstadosHPS.CompareMode = TextCompare
    With m_objColEstadosHPS
        .Add "1", "A punto de Caducar con Renovación No"
        .Add "2", "A punto de Caducar con Renovación<>No Sin Solicitud"
        .Add "3", "A punto de Caducar con Renovación<>No Con Solicitud"
        .Add "4", "Caducada con Renovación No"
        .Add "5", "Caducada con Renovación<>No Sin Solicitud"
        .Add "6", "Caducada con Renovación<>No Con Solicitud"
        .Add "7", "Activo con Solicitud"
        .Add "8", "Activo sin Solicitud"
        .Add "9", "Activo Renovación NO"
        .Add "10", "Solicitud Nueva"
        .Add "11", "Pendiente de Renovación"
        .Add "12", "Baja"
        
        .Add "13", "Irregular"
        .Add "14", "No Posee"
    End With
    If Not m_objColEstadosHPS Is Nothing Then
        Set ColEstadosHPS = m_objColEstadosHPS
        Exit Property
    End If
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColEstadosHPS ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
    
End Property
Public Property Get ColHPSEstados() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    
    If Not m_objColHPSEstados Is Nothing Then
        Set ColHPSEstados = m_objColHPSEstados
        Exit Property
    End If
    Set m_objColHPSEstados = New Scripting.Dictionary
    m_objColHPSEstados.CompareMode = TextCompare
    With m_objColHPSEstados
        .Add "A punto de Caducar con Renovación No", 1
        .Add "A punto de Caducar con Renovación<>No Sin Solicitud", 2
        .Add "A punto de Caducar con Renovación<>No Con Solicitud", 3
        .Add "Caducada con Renovación No", 4
        .Add "Caducada con Renovación<>No Sin Solicitud", 5
        .Add "Caducada con Renovación<>No Con Solicitud", 6
        .Add "Activo con Solicitud", 7
        .Add "Activo sin Solicitud", 8
        .Add "Activo Renovación NO", 9
        .Add "Solicitud Nueva", 10
        .Add "Pendiente de Renovación", 11
        .Add "Baja", 12
        
        .Add "Irregular", 13
        .Add "No Posee", 14
    End With
    If Not m_objColHPSEstados Is Nothing Then
        Set ColHPSEstados = m_objColHPSEstados
        Exit Property
    End If
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColHPSEstados ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
    
End Property


Public Property Get ColTiposConsultas() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    
    If Not m_objColTiposConsultas Is Nothing Then
        Set ColTiposConsultas = m_objColTiposConsultas
        Exit Property
    End If
    Set m_objColTiposConsultas = New Scripting.Dictionary
    m_objColTiposConsultas.CompareMode = TextCompare
    With m_objColTiposConsultas
        .Add "1", "Todo"
        .Add "2", "Por Jurídica que Contrata"
        .Add "3", "Por Proyecto Asignado"
        .Add "4", "Por Nombre y Apellidos"
        .Add "5", "Por Grado de Clasificación de HPS"
        .Add "6", "Por Empresa que Pertenece el Usuario"
        .Add "7", "Por Jurídica que Tramita el HPS"
        .Add "8", "Por Motivo de HPS"
        .Add "9", "Requiere Curso"
    End With
    
    If Not m_objColTiposConsultas Is Nothing Then
        Set ColTiposConsultas = m_objColTiposConsultas
        Exit Property
    End If
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColTiposConsultas ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get ColTiposConsultasFormulario() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    
    If Not m_objColTiposConsultasFormulario Is Nothing Then
        Set ColTiposConsultasFormulario = m_objColTiposConsultasFormulario
        Exit Property
    End If
    Set m_objColTiposConsultasFormulario = New Scripting.Dictionary
    m_objColTiposConsultasFormulario.CompareMode = TextCompare
    With m_objColTiposConsultasFormulario
        .Add "1", "FormInicial03ConsultasEstado"
        .Add "2", "FormInicial03ConsultasContratistas"
        .Add "3", "FormInicial03ConsultasProyectoAsignado"
        .Add "4", "FormInicial03ConsultasNombre"
        .Add "5", "FormInicial03ConsultasGrado"
        .Add "6", "FormInicial03ConsultasEmpresaUsuario"
        .Add "7", "FormInicial03ConsultasJuridicaTramitacion"
        .Add "8", "FormInicial03ConsultasMotivoHPS"
        
    End With
    
    If Not m_objColTiposConsultasFormulario Is Nothing Then
        Set ColTiposConsultasFormulario = m_objColTiposConsultasFormulario
        Exit Property
    End If
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColTiposConsultasFormulario ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get ColUsuariosEstadoIrregular() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColUsuariosEstadoIrregular Is Nothing Then
        Set ColUsuariosEstadoIrregular = m_objColUsuariosEstadoIrregular
        Exit Property
    End If
    
    Set m_objColUsuariosEstadoIrregular = getListaDatosLocalPorWhere("Irregular", m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColUsuariosEstadoIrregular = m_objColUsuariosEstadoIrregular
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColUsuariosEstadoIrregular ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Set ColUsuariosEstadoIrregular(value As Variant)

    Set m_objColUsuariosEstadoIrregular = value

End Property

Public Property Get ColUsuariosAPuntoCaducar() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColUsuariosAPuntoCaducar Is Nothing Then
        Set ColUsuariosAPuntoCaducar = m_objColUsuariosAPuntoCaducar
        Exit Property
    End If
    Set m_objColUsuariosAPuntoCaducar = Constructor.getUsuariosParaIndicadores(EnumTipoIndicador.HPSAPuntoDeCaducar, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColUsuariosAPuntoCaducar = m_objColUsuariosAPuntoCaducar
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColUsuariosAPuntoCaducar ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Set ColUsuariosAPuntoCaducar(value As Variant)

    Set m_objColUsuariosAPuntoCaducar = value

End Property

Public Property Get ColUsuariosCaducados() As Scripting.Dictionary

    
     On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColUsuariosCaducados Is Nothing Then
        Set ColUsuariosCaducados = m_objColUsuariosCaducados
        Exit Property
    End If
    Set m_objColUsuariosCaducados = Constructor.getUsuariosParaIndicadores(EnumTipoIndicador.HPSACaducadas, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColUsuariosCaducados = m_objColUsuariosCaducados
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColUsuariosCaducados ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Set ColUsuariosCaducados(value As Variant)

    Set m_objColUsuariosCaducados = value

End Property

Public Property Get ColUsuariosHPSRequierenCurso() As Scripting.Dictionary

    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColUsuariosHPSRequierenCurso Is Nothing Then
        Set ColUsuariosHPSRequierenCurso = m_objColUsuariosHPSRequierenCurso
        Exit Property
    End If
    Set m_objColUsuariosHPSRequierenCurso = Constructor.getUsuariosParaIndicadores(EnumTipoIndicador.PendientesCursoHPS, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColUsuariosHPSRequierenCurso = m_objColUsuariosHPSRequierenCurso
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColUsuariosHPSRequierenCurso ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Set ColUsuariosHPSRequierenCurso(value As Variant)

    Set m_objColUsuariosHPSRequierenCurso = value

End Property


Public Property Get ColRequiereCursoFaltaPrimeraConvocatoria() As Scripting.Dictionary

    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColRequiereCursoFaltaPrimeraConvocatoria Is Nothing Then
        Set ColRequiereCursoFaltaPrimeraConvocatoria = m_objColRequiereCursoFaltaPrimeraConvocatoria
        Exit Property
    End If
    Set m_objColRequiereCursoFaltaPrimeraConvocatoria = Constructor.getUsuariosParaIndicadores(EnumTipoIndicador.PtesPrimeraConvocatoriaCurso, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColRequiereCursoFaltaPrimeraConvocatoria = m_objColRequiereCursoFaltaPrimeraConvocatoria
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColRequiereCursoFaltaPrimeraConvocatoria ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Set ColRequiereCursoFaltaPrimeraConvocatoria(value As Variant)

    Set m_objColRequiereCursoFaltaPrimeraConvocatoria = value

End Property

Public Property Get ColRequiereCursoFaltaSegundaConvocatoria() As Scripting.Dictionary

    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColRequiereCursoFaltaSegundaConvocatoria Is Nothing Then
        Set ColRequiereCursoFaltaSegundaConvocatoria = m_objColRequiereCursoFaltaSegundaConvocatoria
        Exit Property
    End If
    Set m_objColRequiereCursoFaltaSegundaConvocatoria = Constructor.getUsuariosParaIndicadores(EnumTipoIndicador.PtesSegundaConvocatoriaCurso, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColRequiereCursoFaltaSegundaConvocatoria = m_objColRequiereCursoFaltaSegundaConvocatoria
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColRequiereCursoFaltaSegundaConvocatoria ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Set ColRequiereCursoFaltaSegundaConvocatoria(value As Variant)

    Set m_objColRequiereCursoFaltaSegundaConvocatoria = value

End Property


Public Property Get ColRequiereCursoFaltaCorreoJefeSeguridad() As Scripting.Dictionary

    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColRequiereCursoFaltaCorreoJefeSeguridad Is Nothing Then
        Set ColRequiereCursoFaltaCorreoJefeSeguridad = m_objColRequiereCursoFaltaCorreoJefeSeguridad
        Exit Property
    End If
    Set m_objColRequiereCursoFaltaCorreoJefeSeguridad = Constructor.getUsuariosParaIndicadores(EnumTipoIndicador.PtesEnvioCorreoJefeSeguridadCurso, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColRequiereCursoFaltaCorreoJefeSeguridad = m_objColRequiereCursoFaltaCorreoJefeSeguridad
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColRequiereCursoFaltaCorreoJefeSeguridad ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Set ColRequiereCursoFaltaCorreoJefeSeguridad(value As Variant)

    Set m_objColRequiereCursoFaltaCorreoJefeSeguridad = value

End Property

Public Property Get ColUsuariosEnSolicitud() As Scripting.Dictionary

    
    On Error GoTo errores
    
    Me.Error = ""
    If Not m_objColUsuariosEnSolicitud Is Nothing Then
        Set ColUsuariosEnSolicitud = m_objColUsuariosEnSolicitud
        Exit Property
    End If
    
    Set m_objColUsuariosEnSolicitud = Constructor.getUsuariosParaIndicadores(EnumTipoIndicador.HPSASolicitandose, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColUsuariosEnSolicitud = m_objColUsuariosEnSolicitud
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColUsuariosEnSolicitud ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property
Public Property Set ColUsuariosEnSolicitud(value As Variant)

    Set m_objColUsuariosEnSolicitud = value

End Property
Public Property Get NumeroUsuariosEnSolicitud() As String

    On Error GoTo errores
    
    Me.Error = ""
    If Me.ColUsuariosEnSolicitud Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_sNumeroUsuariosEnSolicitud = "0"
        NumeroUsuariosEnSolicitud = m_sNumeroUsuariosEnSolicitud
        Exit Property
    End If
    m_sNumeroUsuariosEnSolicitud = CStr(Me.ColUsuariosEnSolicitud.Count)
    NumeroUsuariosEnSolicitud = m_sNumeroUsuariosEnSolicitud
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.NumeroUsuariosEnSolicitud ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error

End Property

Public Property Get NumeroUsuariosHPSSinCurso() As String

    On Error GoTo errores
    
    Me.Error = ""
    If Me.ColUsuariosHPSRequierenCurso Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_sNumeroUsuariosHPSSinCurso = "0"
        NumeroUsuariosHPSSinCurso = m_sNumeroUsuariosHPSSinCurso
        Exit Property
    End If
    m_sNumeroUsuariosHPSSinCurso = CStr(Me.ColUsuariosHPSRequierenCurso.Count)
    NumeroUsuariosHPSSinCurso = m_sNumeroUsuariosHPSSinCurso
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.NumeroUsuariosHPSSinCurso ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error

End Property

Public Property Get URLCarpetaIconos() As String
    
    Dim m_URLCarpetaRecursos As String
    On Error GoTo errores
    
    Me.Error = ""
    If m_sURLCarpetaIconos <> "" Then
        URLCarpetaIconos = m_sURLCarpetaIconos
        Exit Property
    End If
    m_URLCarpetaRecursos = Me.URLCarpetaRecursos
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_sURLCarpetaIconos = Me.URLCarpetaRecursos & "Iconos\"
    If Not fso.FolderExists(m_sURLCarpetaIconos) Then
        m_Error = "No se puede alcanzar m_sURLCarpetaIconos" & vbNewLine & m_sURLCarpetaIconos
        Err.Raise 1000
    End If
    
    URLCarpetaIconos = m_sURLCarpetaIconos
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLCarpetaIconos ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property


Public Property Get URLImagenSemaforoRojo() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    If m_sURLImagenSemaforoRojo <> "" Then
        URLImagenSemaforoRojo = m_sURLImagenSemaforoRojo
        Exit Property
    End If
    m_sURLImagenSemaforoRojo = Me.URLCarpetaIconos & "Semaforo_Rojo.png"
    If Not fso.FileExists(m_sURLImagenSemaforoRojo) Then
        m_Error = "No se puede alcanzar m_sURLImagenSemaforoRojo" & vbNewLine & m_sURLImagenSemaforoRojo
        Err.Raise 1000
    End If
    URLImagenSemaforoRojo = m_sURLImagenSemaforoRojo
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLImagenSemaforoRojo ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get URLImagenSemaforoVerde() As String

    On Error GoTo errores
    
    Me.Error = ""
    If m_sURLImagenSemaforoVerde <> "" Then
        URLImagenSemaforoVerde = m_sURLImagenSemaforoVerde
        Exit Property
    End If
    m_sURLImagenSemaforoVerde = Me.URLCarpetaIconos & "Semaforo_Verde.png"
    If Not fso.FileExists(m_sURLImagenSemaforoVerde) Then
        m_Error = "No se puede alcanzar m_sURLImagenSemaforoVerde" & vbNewLine & m_sURLImagenSemaforoVerde
        Err.Raise 1000
    End If
    URLImagenSemaforoVerde = m_sURLImagenSemaforoVerde
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLImagenSemaforoVerde ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error

End Property



Public Property Get URLArchivoExcelPrevio() As String
    
    Dim m_NombreArchivoExcelPrevio As String
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    If m_sURLArchivoExcelPrevio <> "" Then
        URLArchivoExcelPrevio = m_sURLArchivoExcelPrevio
        Exit Property
    End If
    m_NombreArchivoExcelPrevio = "ExcelPrevio.xlsx"
    m_sURLArchivoExcelPrevio = Me.URLCarpetaAnexos & m_NombreArchivoExcelPrevio
    
    URLArchivoExcelPrevio = m_sURLArchivoExcelPrevio
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.URLArchivoExcelPrevio ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property



Public Property Get DatosEn() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    If m_EnOficina = EnumSiNo.Sí Then
        m_sDatosEn = "Oficina"
    ElseIf m_EnOficina = EnumSiNo.No Then
        m_sDatosEn = "Local"
    Else
        m_sDatosEn = "#ERR"
    End If
    DatosEn = m_sDatosEn
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.DatosEn ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get ModoEscritura() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    If Me.SoloLectura = EnumSiNo.Sí Then
        m_sModoEscritura = "No"
    ElseIf Me.SoloLectura = EnumSiNo.No Then
        m_sModoEscritura = "Sí"
    Else
        m_sModoEscritura = "#ERR"
    End If
    
    ModoEscritura = m_sModoEscritura
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ModoEscritura ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property


Public Property Get ColPropiedadesBotonEscritura() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    
    If Not m_objColPropiedadesBotonEscritura Is Nothing Then
        Set ColPropiedadesBotonEscritura = m_objColPropiedadesBotonEscritura
        Exit Property
    End If
    Set m_objColPropiedadesBotonEscritura = New Scripting.Dictionary
    m_objColPropiedadesBotonEscritura.CompareMode = TextCompare
    With m_objColPropiedadesBotonEscritura
        .Add "BackColor", 6076508
        .Add "BackShade", 100
        .Add "BackStyle", 1
        .Add "BackTint", 100
        .Add "BorderThemeColorIndex", -1
        .Add "ThemeFontIndex", 1
        .Add "BorderColor", 6076508
        .Add "BorderShade", 100
        .Add "BorderStyle", 1
        .Add "BorderTint", 100
        .Add "CursorOnHover", 1120403456
        .Add "FontSize", 11
        .Add "FontWeight", 800
        .Add "FontBold", 1
        .Add "ForeColor", 16777215
        .Add "ForeShade", 100
        .Add "ForeThemeColorIndex", 1
        .Add "ForeTint", 100
        .Add "Glow", 0
        .Add "Gradient", 0
        .Add "HoverColor", 4695111
        .Add "HoverForeColor", 16777215
        .Add "HoverForeShade", 100
        .Add "HoverForeThemeColorIndex", 1
        .Add "HoverForeTint", 100
        .Add "HoverShade", 100
        .Add "HoverTint", 100
        .Add "PressedColor", 4695111
        .Add "PressedForeColor", 8572084
        .Add "PressedForeShade", 100
        .Add "PressedForeThemeColorIndex", 6
        .Add "PressedShade", 100
        .Add "PressedThemeColorIndex", -1
        .Add "PressedTint", 100
        .Add "Shadow", 0
        .Add "Shape", 1
        
        .Add "Transparent", False
        .Add "UseTheme", True
    End With
    If Not m_objColPropiedadesBotonEscritura Is Nothing Then
        Set ColPropiedadesBotonEscritura = m_objColPropiedadesBotonEscritura
        Exit Property
    End If
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColPropiedadesBotonEscritura ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property
Public Property Get ColPropiedadesBotonLectura() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    
    If Not m_objColPropiedadesBotonLectura Is Nothing Then
        Set ColPropiedadesBotonLectura = m_objColPropiedadesBotonLectura
        Exit Property
    End If
    Set m_objColPropiedadesBotonLectura = New Scripting.Dictionary
    m_objColPropiedadesBotonLectura.CompareMode = TextCompare
    With m_objColPropiedadesBotonLectura
        .Add "BackColor", 13273922
        .Add "BackShade", 100
        .Add "BackStyle", 1
        .Add "BackTint", 100
        .Add "BorderThemeColorIndex", -1
        .Add "ThemeFontIndex", 1
        .Add "BorderColor", 13273922
        .Add "BorderShade", 100
        .Add "BorderStyle", 1
        .Add "BorderTint", 100
        .Add "CursorOnHover", 1120403456
        .Add "FontSize", 11
        .Add "FontWeight", 800
        .Add "FontBold", 1
        .Add "ForeColor", 16777215
        .Add "ForeShade", 100
        .Add "ForeThemeColorIndex", 1
        .Add "ForeTint", 100
        .Add "Glow", 0
        .Add "Gradient", 0
        .Add "HoverColor", 11630130
        .Add "HoverForeColor", 16777215
        .Add "HoverForeShade", 100
        .Add "HoverForeThemeColorIndex", 1
        .Add "HoverForeTint", 100
        .Add "HoverShade", 100
        .Add "HoverTint", 100
        .Add "PressedColor", 11630130
        .Add "PressedForeColor", 13476219
        .Add "PressedForeShade", 100
        .Add "PressedForeThemeColorIndex", 4
        .Add "PressedShade", 100
        .Add "PressedThemeColorIndex", -1
        .Add "PressedTint", 100
        .Add "Shadow", 0
        .Add "Shape", 1
        .Add "Transparent", False
        .Add "UseTheme", True


    End With
    If Not m_objColPropiedadesBotonLectura Is Nothing Then
        Set ColPropiedadesBotonLectura = m_objColPropiedadesBotonLectura
        Exit Property
    End If
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColPropiedadesBotonLectura ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property
Public Property Get colGradosHPS() As Scripting.Dictionary
    
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objcolGradosHPS Is Nothing Then
        Set colGradosHPS = m_objcolGradosHPS
        Exit Property
    End If
    
    Set m_objcolGradosHPS = Constructor.getListaColeccionesGradosHPS(m_Error)
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set colGradosHPS = m_objcolGradosHPS
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.colGradosHPS ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get colFiltroEnIndicadordes() As Scripting.Dictionary
    
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objcolFiltroEnIndicadordes Is Nothing Then
        Set colFiltroEnIndicadordes = m_objcolFiltroEnIndicadordes
        Exit Property
    End If
    Set m_objcolFiltroEnIndicadordes = New Scripting.Dictionary
    m_objcolFiltroEnIndicadordes.CompareMode = TextCompare
    With m_objcolFiltroEnIndicadordes
        .Add "ComandoPendientesCursoHPS", "FILTRO APLICADO: PENDIENTES DE CURSO"
        .Add "ComandoAPuntoDeCaducar", "FILTRO APLICADO: A PUNTO DE CADUCAR"
        .Add "ComandoCaducados", "FILTRO APLICADO: CADUCADOS"
        .Add "ComandoEnSolicitud", "FILTRO APLICADO: CON SOLICITUD ACTIVA"
        .Add "ComandoPendientesConvocatoria1", "FILTRO APLICADO: PENDIENTES DE PRIMERA CONVOCATORIA"
        .Add "ComandoPendientesConvocatoria2", "FILTRO APLICADO: PENDIENTES DE SEGUNDA CONVOCATORIA"
        .Add "ComandoPendientesCorreoJefeSeguridad", "FILTRO APLICADO: PENDIENTES DE CORREO A JEFE DE SEGURIDAD"
    End With
    
    Set colFiltroEnIndicadordes = m_objcolFiltroEnIndicadordes
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.colFiltroEnIndicadordes ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get ColConsultas() As Scripting.Dictionary

    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objColConsultas Is Nothing Then
        Set ColConsultas = m_objColConsultas
        Exit Property
    End If
    
    Set m_objColConsultas = Constructor.getConsultas(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColConsultas = m_objColConsultas
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColConsultas ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error

End Property

Public Property Get ColCamposSuministradores() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    
    If Not m_objColCamposSuministradores Is Nothing Then
        Set ColCamposSuministradores = m_objColCamposSuministradores
        Exit Property
    End If
    Set m_objColCamposSuministradores = New Scripting.Dictionary
    m_objColCamposSuministradores.CompareMode = TextCompare
    With m_objColCamposSuministradores
        .Add "Empresa Usuario", "TbSuministradoresLocal.Nombre"
        .Add "Empresa Tramitadora", "TbSuministradoresLocal_1.Nombre"
        .Add "Jurídica Contratación", "TbSuministradoresLocal_2.Nombre"
    End With
    Set ColCamposSuministradores = m_objColCamposSuministradores
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColCamposSuministradores ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get ExpedientesUsados() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    
    If Not m_objExpedientesUsados Is Nothing Then
        Set ExpedientesUsados = m_objExpedientesUsados
        Exit Property
    End If
    
    
    Set m_objExpedientesUsados = Constructor.getExpedientesUsados(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set ExpedientesUsados = m_objExpedientesUsados
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ExpedientesUsados ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property
Public Property Set ExpedientesUsados(p_Valor As Variant)
    
    Set m_objExpedientesUsados = p_Valor
    
End Property

Public Property Get Expedientes() As Scripting.Dictionary
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Not m_ObjExpedientes Is Nothing Then
        Set Expedientes = m_ObjExpedientes
        Exit Property
    End If
    
    Set m_ObjExpedientes = Constructor.getExpedientes(p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Expedientes = m_ObjExpedientes
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.Expedientes ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Set Expedientes(p_Valor As Variant)

    Set m_ObjExpedientes = p_Valor
    
End Property


Public Property Get ColPropiedadesEntorno() As Scripting.Dictionary

    
    On Error GoTo errores
    Me.Error = ""
    
    If Not m_objColPropiedadesEntorno Is Nothing Then
        Set ColPropiedadesEntorno = m_objColPropiedadesEntorno
        Exit Property
    End If
    Set m_objColPropiedadesEntorno = New Scripting.Dictionary
    m_objColPropiedadesEntorno.CompareMode = TextCompare
    
    With m_objColPropiedadesEntorno
        .Add "CabeceraBackColor", ""
        .Add "CampoBackColor", ""
        .Add "CampoBorderColor", ""
        .Add "CampoForeColor", ""
        .Add "colColoresMaster", "o"
        .Add "colColoresSlave", "o"
        .Add "ColEstadosHPS", "o"
        .Add "ColEstadosVisiblesHPS", "o"
        .Add "colGradosHPS", "o"
        .Add "ColHPSEstados", "o"
        .Add "ColConsultas", "o"
        .Add "ColTiposConsultas", "o"
        .Add "ColTiposConsultasFormulario", "o"
        .Add "ColVisiblesHPSEstados", "o"
        .Add "EtiquetaBackColor", ""
        .Add "EtiquetaBorderColor", ""
        .Add "EtiquetaForeColor", ""
        .Add "EtiquetaForeColorSeleccionado", ""
        .Add "EtiquetaLinkBackColorSeleccionado", ""
        .Add "EtiquetaLinkForeColor", ""
        .Add "EtiquetaSeleccionadaBackColor", ""
        .Add "EtiquetaSeleccionadaForeColor", ""
        .Add "ListaBackColor", ""
        .Add "ListaBorderColor", ""
        .Add "ListaForeColor", ""
        .Add "MesesParaRenovacion", ""
        .Add "SoloLectura", "s"
        .Add "DatosEn", ""
        .Add "ModoEscritura", ""
        .Add "URLCarpetaAplicacionesDys", ""
        .Add "URLCarpetaRecursos", ""
        .Add "URLCarpetaAplicacion", ""
        .Add "URLCarpetaAnexos", ""
        .Add "URLCarpetaAnexosSICA", ""
        .Add "URLCarpetaAnexosHistoricos", ""
        .Add "URLDirectorioLocal", ""
        .Add "URLCarpetaIconos", ""
        .Add "URLImagenSemaforoRojo", ""
        .Add "URLImagenSemaforoVerde", ""
        .Add "URLArchivoExcelPrevio", ""
        .Add "ColEmpresas", "o"
        .Add "ColEmpresasTramitadoras", "o"
        .Add "ColCamposSuministradores", "o"
        .Add "colUsuariosPermitidosEnLocal", "o"
        .Add "colFiltroEnIndicadordes", "o"
        .Add "DiasEnPrimeraConvocatoriaSinCurso", ""
        .Add "DiasSinCorreoNoCurso", ""
        .Add "DiasEnSegundaConvocatoriaSinCurso", ""
    End With
    
    Set ColPropiedadesEntorno = m_objColPropiedadesEntorno
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Entorno.ColPropiedadesEntorno ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Function getValorPropiedad( _
                                p_NombreCampo As String, _
                                Optional ByRef p_Error As String _
                                ) As Variant
    
    On Error GoTo errores
    p_Error = ""
    
    If p_NombreCampo = "CabeceraBackColor" Then
        getValorPropiedad = Me.CabeceraBackColor
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "CampoBackColor" Then
        getValorPropiedad = Me.CampoBackColor
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "CampoBorderColor" Then
        getValorPropiedad = Me.CampoBorderColor
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "CampoForeColor" Then
        getValorPropiedad = Me.CampoForeColor
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_NombreCampo = "colColoresMaster" Then
        Set getValorPropiedad = Me.colColoresMaster
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "colColoresSlave" Then
        Set getValorPropiedad = Me.colColoresSlave
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "ColEstadosHPS" Then
        Set getValorPropiedad = Me.ColEstadosHPS
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "ColEstadosVisiblesHPS" Then
        Set getValorPropiedad = Me.ColEstadosVisiblesHPS
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "ColHPSEstados" Then
        Set getValorPropiedad = Me.ColHPSEstados
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    
    If p_NombreCampo = "ColPropiedadesEntorno" Then
        Set getValorPropiedad = Me.ColPropiedadesEntorno
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_NombreCampo = "ColTiposConsultas" Then
        Set getValorPropiedad = Me.ColTiposConsultas
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "ColTiposConsultasFormulario" Then
        Set getValorPropiedad = Me.ColTiposConsultasFormulario
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    
    
    If p_NombreCampo = "ColVisiblesHPSEstados" Then
        Set getValorPropiedad = Me.ColVisiblesHPSEstados
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    
    
    If p_NombreCampo = "EtiquetaBackColor" Then
        getValorPropiedad = Me.EtiquetaBackColor
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "EtiquetaBorderColor" Then
        getValorPropiedad = Me.EtiquetaBorderColor
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "EtiquetaForeColor" Then
        getValorPropiedad = Me.EtiquetaForeColor
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "EtiquetaSeleccionadaBackColor" Then
        getValorPropiedad = Me.EtiquetaSeleccionadaBackColor
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "EtiquetaSeleccionadaForeColor" Then
        getValorPropiedad = Me.EtiquetaSeleccionadaForeColor
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "EtiquetaLinkForeColor" Then
        getValorPropiedad = Me.EtiquetaLinkForeColor
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "EtiquetaForeColorSeleccionado" Then
        getValorPropiedad = Me.EtiquetaForeColorSeleccionado
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "EtiquetaLinkBackColorSeleccionado" Then
        getValorPropiedad = Me.EtiquetaLinkBackColorSeleccionado
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "ListaBackColor" Then
        getValorPropiedad = Me.ListaBackColor
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "ListaBorderColor" Then
        getValorPropiedad = Me.ListaBorderColor
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "ListaForeColor" Then
        getValorPropiedad = Me.ListaForeColor
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "MesesParaRenovacion" Then
        getValorPropiedad = Me.MesesParaRenovacion
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    
    If p_NombreCampo = "URLCarpetaAplicacionesDys" Then
        getValorPropiedad = Me.URLCarpetaAplicacionesDys
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "URLCarpetaRecursos" Then
        getValorPropiedad = Me.URLCarpetaRecursos
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "URLCarpetaAplicacion" Then
        getValorPropiedad = Me.URLCarpetaAplicacion
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "URLCarpetaAnexos" Then
        getValorPropiedad = Me.URLCarpetaAnexos
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "URLCarpetaAnexosSICA" Then
        getValorPropiedad = Me.URLCarpetaAnexosSICA
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "URLCarpetaAnexosHistoricos" Then
        getValorPropiedad = Me.URLCarpetaAnexosHistoricos
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "URLDirectorioLocal" Then
        getValorPropiedad = Me.URLDirectorioLocal
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "URLCarpetaIconos" Then
        getValorPropiedad = Me.URLCarpetaIconos
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_NombreCampo = "URLIni" Then
        getValorPropiedad = Me.URLIni
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_NombreCampo = "VersionAplicacion" Then
        getValorPropiedad = Me.VersionAplicacion
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "ColTiposConsultasFormulario" Then
        Set getValorPropiedad = Me.ColTiposConsultasFormulario
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_NombreCampo = "colGradosHPS" Then
        Set getValorPropiedad = Me.colGradosHPS
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_NombreCampo = "SoloLectura" Then
        getValorPropiedad = Me.SoloLectura
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If p_NombreCampo = "ColUsuariosEstadoIrregular" Then
        Set getValorPropiedad = Me.ColUsuariosEstadoIrregular
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "ColUsuariosAPuntoCaducar" Then
        Set getValorPropiedad = Me.ColUsuariosAPuntoCaducar
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "ColUsuariosCaducados" Then
        Set getValorPropiedad = Me.ColUsuariosCaducados
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "ColUsuariosEnSolicitud" Then
        Set getValorPropiedad = Me.ColUsuariosEnSolicitud
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
   
    
    
    If p_NombreCampo = "DatosEn" Then
        getValorPropiedad = Me.DatosEn
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "ModoEscritura" Then
        getValorPropiedad = Me.ModoEscritura
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "ColUsuariosHPSRequierenCurso" Then
        Set getValorPropiedad = Me.ColUsuariosHPSRequierenCurso
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "ColRequiereCursoFaltaPrimeraConvocatoria" Then
        Set getValorPropiedad = Me.ColRequiereCursoFaltaPrimeraConvocatoria
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "ColRequiereCursoFaltaSegundaConvocatoria" Then
        Set getValorPropiedad = Me.ColRequiereCursoFaltaSegundaConvocatoria
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "ColRequiereCursoFaltaCorreoJefeSeguridad" Then
        Set getValorPropiedad = Me.ColRequiereCursoFaltaCorreoJefeSeguridad
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
   
    If p_NombreCampo = "URLImagenSemaforoRojo" Then
        getValorPropiedad = Me.URLImagenSemaforoRojo
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "URLImagenSemaforoVerde" Then
        getValorPropiedad = Me.URLImagenSemaforoVerde
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "URLArchivoExcelPrevio" Then
        getValorPropiedad = Me.URLArchivoExcelPrevio
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "ColEmpresas" Then
        Set getValorPropiedad = Me.ColEmpresas
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "ColEmpresasTramitadoras" Then
        Set getValorPropiedad = Me.ColEmpresasTramitadoras
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "ColConsultas" Then
        Set getValorPropiedad = Me.ColConsultas
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "ColCamposSuministradores" Then
        Set getValorPropiedad = Me.ColCamposSuministradores
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "colUsuariosPermitidosEnLocal" Then
        Set getValorPropiedad = Me.colUsuariosPermitidosEnLocal
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "colFiltroEnIndicadordes" Then
        Set getValorPropiedad = Me.colFiltroEnIndicadordes
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "DiasEnPrimeraConvocatoriaSinCurso" Then
        getValorPropiedad = Me.DiasEnPrimeraConvocatoriaSinCurso
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "DiasSinCorreoNoCurso" Then
        getValorPropiedad = Me.DiasSinCorreoNoCurso
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    If p_NombreCampo = "DiasEnSegundaConvocatoriaSinCurso" Then
        getValorPropiedad = Me.DiasEnSegundaConvocatoriaSinCurso
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    p_Error = "La propiedad " & p_NombreCampo & " no está registrada en Entorno.getValorPropiedad "
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Entorno.getValorPropiedad ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Function


' ============================================================
' PR B -- TbConfiguracionHPS cache layer + relink helper.
'
' Con estos miembros, Entorno deja de depender exclusivamente de las
' rutas hardcoded / TempVars: la primera vez que alguien pide una
' propiedad de ruta, se carga TbConfiguracionHPS a un Scripting.Dictionary
' en memoria, y GetConfig devuelve el valor cacheado. Si la tabla no
' existe o esta vacia, el cache queda vacio y los property Gets caen al
' fallback legacy (TempVar DatosEnLocal / hardcoded) -- asi la app sigue
' funcionando aunque nadie haya corrido el helper de PR A.
'
' RelinkAllTables itera los TableDefs adjuntos no-MSys y, para cada uno
' que figure en TbVinculosTablas, reescribe su Connect con el .accdb que
' cuelga del directorio pasado. Es best-effort: si RefreshLink falla en
' una tabla, se loguea con Debug.Print y se sigue; devuelve True solo si
' todas las relinks tuvieron exito.
' ============================================================

' Lee TbConfiguracionHPS y rellena m_ConfigCache. Idempotente: se puede
' llamar varias veces para refrescar tras modificar la tabla (p.ej. desde
' el boton "Aplicar perfil" del Form_FormInicial). Si la tabla no existe,
' registra un warning y deja el cache vacio -- los property Gets haran
' fallback al comportamiento legacy.
Public Sub CargarConfiguracion()
    On Error GoTo EH

    Set m_ConfigCache = New Scripting.Dictionary
    m_ConfigCache.CompareMode = TextCompare

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim rs As DAO.Recordset
    Dim sSql As String

    ' CurrentDb siempre esta disponible en el contexto del frontend.
    Set db = CurrentDb

    ' Si la tabla aun no existe (usuario no ha corrido el helper de PR A),
    ' registramos un warning y dejamos el cache vacio. NO raise: la app
    ' debe seguir funcionando con el fallback legacy.
    On Error Resume Next
    Set tdf = db.TableDefs(k_ConfigTable)
    On Error GoTo EH
    If tdf Is Nothing Then
        Debug.Print "Entorno.CargarConfiguracion: la tabla " & k_ConfigTable & _
                   " no existe todavia. Se usara el fallback hardcoded hasta que se corra el helper de PR A."
        Exit Sub
    End If

    ' Leemos solo filas Activas. En Access/ACE los campos YesNo guardan
    ' True como -1; usar Activo = 1 deja fuera filas creadas como True.
    sSql = "SELECT Clave, Valor FROM " & k_ConfigTable & " WHERE Activo <> 0"
    On Error Resume Next
    Set rs = db.OpenRecordset(sSql, dbOpenSnapshot)
    If Err.Number <> 0 Then
        Debug.Print "Entorno.CargarConfiguracion: no se pudo leer la tabla: " & Err.Description
        On Error GoTo EH
        Exit Sub
    End If
    On Error GoTo EH

    Do While Not rs.EOF
        Dim sClave As String
        Dim sValor As String
        sClave = CStr(Nz(rs.Fields("Clave").value, ""))
        sValor = CStr(Nz(rs.Fields("Valor").value, ""))
        If Len(sClave) > 0 Then
            If Not m_ConfigCache.Exists(sClave) Then
                m_ConfigCache.Add sClave, sValor
            End If
        End If
        rs.MoveNext
    Loop

    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Exit Sub
EH:
    ' Cualquier error inesperado -> cache vacio + log; nunca raise,
    ' porque CargarConfiguracion se invoca desde getters que ya tienen
    ' su propia gestion de errores.
    If Not m_ConfigCache Is Nothing Then
        Set m_ConfigCache = Nothing
    End If
    Debug.Print "Entorno.CargarConfiguracion: error inesperado: " & Err.Description
End Sub

' Devuelve el valor cacheado de TbConfiguracionHPS para la Clave indicada.
' Si la clave no esta en el cache (o el cache esta vacio porque la tabla
' no existe), devuelve "" -- el llamador hace fallback al TempVar/hardcoded.
' Si la cache nunca se cargo, dispara CargarConfiguracion de forma perezosa.
Public Function GetConfig(ByVal p_Clave As String) As String
    On Error Resume Next
    If m_ConfigCache Is Nothing Then
        CargarConfiguracion
    End If
    If m_ConfigCache Is Nothing Then
        GetConfig = ""
        Exit Function
    End If
    If m_ConfigCache.Exists(p_Clave) Then
        GetConfig = CStr(m_ConfigCache(p_Clave))
    Else
        GetConfig = ""
    End If
    On Error GoTo 0
End Function

' [config-as-source-of-truth] Variante estricta del resolver: si la clave
' no está sembrada en TbConfiguracionHPS, Err.Raise 1000 con un mensaje
' que apunta a ConfigurarTablaConfiguracion_PerfilLocal/Produccion. Sin
' fallback hardcoded. Es la única puerta de entrada segura para paths de
' backends y carpetas de anexos; usar GetConfig cuando se acepte
' legitimamente un valor vacío (p.ej. lectura tolerante de features
' opcionales).
Public Function GetConfigOrFail(ByVal p_Clave As String) As String
    On Error GoTo EH
    Dim sValue As String
    sValue = GetConfig(p_Clave)
    If Len(sValue) = 0 Then
        m_Error = "GetConfigOrFail: la clave '" & p_Clave & "' no está sembrada en " & _
            k_ConfigTable & ". Ejecuta ConfigurarTablaConfiguracion_PerfilLocal o " & _
            "ConfigurarTablaConfiguracion_PerfilProduccion desde el menú " & _
            "administración de la aplicación para sembrar el valor canónico."
        Err.Raise 1000
    End If
    GetConfigOrFail = sValue
    Exit Function
EH:
    If Err.Number <> 1000 Then
        m_Error = "El método GetConfigOrFail ha devuelto el error: " & vbNewLine & Err.Description
        Err.Raise 1000
    End If
    Me.Error = m_Error
    Err.Raise 1000
End Function

' Devuelve la ruta del .accdb donde vive TbConfiguracionHPS, o "" si la
' tabla no existe todavia. Util para que el form pueda mostrarla al usuario
' en un MsgBox de confirmacion.
Public Function GetConfigTablePath() As String
    On Error GoTo EH
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim sConnect As String
    Set db = CurrentDb
    Set tdf = db.TableDefs(k_ConfigTable)
    sConnect = tdf.Connect
    ' Connect tiene la forma ";DATABASE=C:\path\file.accdb"
    If Left$(sConnect, 10) = ";DATABASE=" Then
        GetConfigTablePath = Mid$(sConnect, 11)
    Else
        GetConfigTablePath = sConnect
    End If
    Exit Function
EH:
    GetConfigTablePath = ""
End Function

' Reconstruye un string de connect preservando propiedades de conexión existentes
' (proveedor, contraseña si ya existe, etc.) y actualizando únicamente la parte
' de DATABASE. Si no hay connect previo, devuelve ";DATABASE=<nuevoPath>".
' Si p_Password no está vacío, lo añade como ;PWD=<password> (reemplaza
' cualquier token PWD o Jet OLEDB:Database Password existente).
Private Function BuildRelinkConnect(ByVal p_CurrentConnect As String, ByVal p_NewDatabasePath As String, Optional ByVal p_Password As String = "") As String
    Dim sCurrent As String
    Dim sResult As String
    Dim sToken As String
    Dim sUpper As String
    Dim vParts As Variant
    Dim i As Long
    Dim bHasDatabaseToken As Boolean
    Dim bHasPwdToken As Boolean
    Dim bStartsWithSemi As Boolean

    sCurrent = Trim$(p_CurrentConnect)
    If Len(sCurrent) = 0 Then
        BuildRelinkConnect = ";DATABASE=" & p_NewDatabasePath
        GoTo ADD_PWD
    End If

    bStartsWithSemi = (Left$(sCurrent, 1) = ";")
    vParts = Split(sCurrent, ";")

    For i = LBound(vParts) To UBound(vParts)
        sToken = Trim$(CStr(vParts(i)))
        If Len(sToken) = 0 Then
            ' Skip empty tokens produced by leading/trailing separators.
        Else
            sUpper = UCase$(sToken)
            If Left$(sUpper, 9) = "DATABASE=" Then
                sToken = "DATABASE=" & p_NewDatabasePath
                bHasDatabaseToken = True
            ElseIf Left$(sUpper, 4) = "PWD=" Then
                ' Token PWD existente: lo reemplazamos si hay contraseña nueva,
                ' o lo preservamos si no.
                If Len(p_Password) > 0 Then
                    sToken = "PWD=" & p_Password
                End If
                bHasPwdToken = True
            ElseIf InStr(1, sUpper, "DATABASE PASSWORD=", vbTextCompare) > 0 Then
                ' Variante Jet OLEDB:Database Password=<pass>
                If Len(p_Password) > 0 Then
                    sToken = "Jet OLEDB:Database Password=" & p_Password
                End If
                bHasPwdToken = True
            End If

            If Len(sResult) = 0 Then
                sResult = sToken
            Else
                sResult = sResult & ";" & sToken
            End If
        End If
    Next i

    If Not bHasDatabaseToken Then
        If Len(sResult) > 0 Then
            sResult = sResult & ";DATABASE=" & p_NewDatabasePath
        Else
            sResult = "DATABASE=" & p_NewDatabasePath
        End If
    End If

    If bStartsWithSemi Then
        sResult = ";" & sResult
    End If

    ' Anadir PWD si se proporcionó y no hay token ya
    If Len(p_Password) > 0 And Not bHasPwdToken Then
        sResult = sResult & ";PWD=" & p_Password
    End If

    BuildRelinkConnect = sResult
    Exit Function

ADD_PWD:
    If Len(p_Password) > 0 Then
        BuildRelinkConnect = BuildRelinkConnect & ";PWD=" & p_Password
    End If
End Function

' Re-vincula cada TableDef adjunto no-MSys cuyo nombre este en
' TbVinculosTablas, apuntando a "<p_NewBackendDir>\<BackendFile>".
' p_NewBackendDir NO debe incluir la barra final.
'
' Best-effort: si RefreshLink falla en una tabla (p.ej. el .accdb no
' esta ahi), se loguea con Debug.Print y se continua. Devuelve True
' solo si TODAS las relinks tuvieron exito.
'
' Tablas no-MSys que NO esten en TbVinculosTablas se saltan silenciosamente
' (con un Debug.Print) -- asi no rompemos tablas de sistema o auxiliares
' que el helper de PR A no haya catalogado.
Public Function RelinkAllTables(ByVal p_NewBackendDir As String) As Boolean
    On Error GoTo EH
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim rs As DAO.Recordset
    Dim vinculoMap As Object
    Dim sTabla As String
    Dim sBackendFile As String
    Dim sConnect As String
    Dim sNewDatabase As String
    Dim allOk As Boolean
    Dim sTrimmedDir As String

    allOk = True
    sTrimmedDir = Trim$(p_NewBackendDir)
    ' Quitar barra final si la hubiera, para que ";DATABASE=dir\file" salga limpio.
    Do While Len(sTrimmedDir) > 0 And (Right$(sTrimmedDir, 1) = "\" Or Right$(sTrimmedDir, 1) = "/")
        sTrimmedDir = Left$(sTrimmedDir, Len(sTrimmedDir) - 1)
    Loop

    Set db = CurrentDb
    Set vinculoMap = CreateObject("Scripting.Dictionary")
    vinculoMap.CompareMode = 1  ' TextCompare

    ' Cargamos el mapa TablaOrigen -> BackendFile desde TbVinculosTablas.
    ' Toleramos que la tabla no exista (helper de PR A no corrio): en ese
    ' caso el mapa queda vacio y solo se relinkan las tablas que matcheen
    ' contra un mapa vacio, que es ninguna -- asi el call site sigue
    ' funcionando y puede mostrar "tabla no configurada" sin reventar.
    On Error Resume Next
    Set rs = db.OpenRecordset("SELECT TablaOrigen, BackendFile FROM " & k_VinculosTable, dbOpenSnapshot)
    If Err.Number = 0 And Not rs Is Nothing Then
        Do While Not rs.EOF
            sTabla = CStr(Nz(rs.Fields("TablaOrigen").value, ""))
            sBackendFile = CStr(Nz(rs.Fields("BackendFile").value, ""))
            If Len(sTabla) > 0 And Len(sBackendFile) > 0 Then
                If Not vinculoMap.Exists(sTabla) Then
                    vinculoMap.Add sTabla, sBackendFile
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
    Err.Clear
    On Error GoTo EH

    ' Iteramos los TableDefs adjuntos del frontend.
    For Each tdf In db.TableDefs
        If (tdf.Attributes And dbAttachedTable) = dbAttachedTable Then
            ' Saltamos MSys* (tablas de sistema de Access) -- no son datos.
            If Left$(tdf.Name, 4) <> "MSys" Then
                If vinculoMap.Exists(tdf.Name) Then
                    sBackendFile = CStr(vinculoMap(tdf.Name))
                    sNewDatabase = sTrimmedDir & "\" & sBackendFile
                    sConnect = BuildRelinkConnect(CStr(tdf.Connect), sNewDatabase, Me.GetConfig(k_Key_BackendPassword))
                    On Error Resume Next
                    tdf.Connect = sConnect
                    If Err.Number = 0 Then
                        tdf.RefreshLink
                    End If
                    If Err.Number <> 0 Then
                        Debug.Print "Entorno.RelinkAllTables: fallo al relink " & tdf.Name & _
                                   " -> " & sConnect & " | " & Err.Description
                        allOk = False
                        Err.Clear
                        ' RefreshLink can roll back the assigned Connect when the
                        ' target backend is unreachable (for example, an offline UNC
                        ' during production preparation). Keep the intended Connect
                        ' persisted so the compiled frontend points at the configured
                        ' backend as soon as the share is reachable, while still
                        ' returning False to signal that validation failed now.
                        tdf.Connect = sConnect
                        If Err.Number <> 0 Then
                            Debug.Print "Entorno.RelinkAllTables: no se pudo conservar Connect para " & _
                                       tdf.Name & " -> " & sConnect & " | " & Err.Description
                            Err.Clear
                        End If
                    End If
                    On Error GoTo EH
                Else
                    ' Tabla adjunta pero no catalogada: la saltamos sin error.
                    ' Mantener compat con tablas que el helper de PR A no registro.
                    Debug.Print "Entorno.RelinkAllTables: saltando " & tdf.Name & _
                               " (no esta catalogada en " & k_VinculosTable & ")"
                End If
            End If
        End If
    Next tdf

    Set vinculoMap = Nothing
    RelinkAllTables = allOk
    Exit Function
EH:
    ' Cualquier error inesperado (p.ej. dbAttachedTable no soportado) se
    ' considera un fallo global -- devolvemos False y dejamos la traza.
    Debug.Print "Entorno.RelinkAllTables: error inesperado: " & Err.Description
    RelinkAllTables = False
End Function

' Abre el backend (.accdb) cuya ruta cacheada coincida con la clave dada.
' Por defecto abre HPST.accdb (k_Key_HPST). El form puede pasar
' "EXPEDIENTES_BACKEND_PATH" o cualquier otra clave que TbConfiguracionHPS
' tenga catalogada. NO modifica VariablesEntorno.getdb / getdbLanzadera
' (que mantienen la firma legacy para compatibilidad con todos los callers
' existentes) -- este es un punto de entrada opt-in para codigo nuevo.
Public Function getdbFromConfig(Optional ByVal p_BackendKey As String = k_Key_HPST) As DAO.Database
    On Error GoTo EH
    Dim sPath As String
    Dim wks As DAO.Workspace
    Dim dbResult As DAO.Database

    If m_ConfigCache Is Nothing Then
        CargarConfiguracion
    End If
    sPath = GetConfig(p_BackendKey)
    If Len(sPath) = 0 Then
        Err.Raise 1000, "Entorno.getdbFromConfig", _
            "No hay ruta cacheada para la clave '" & p_BackendKey & "'. " & _
            "Ejecuta ConfigurarTablaConfiguracion_PerfilLocal/PerfilProduccion primero."
    End If
    If Not fso.FileExists(sPath) Then
        Err.Raise 1000, "Entorno.getdbFromConfig", _
            "El backend cacheado no existe en disco: " & sPath
    End If

    Set wks = DBEngine.Workspaces(0)
    Set dbResult = wks.OpenDatabase(sPath, False, False, "MS Access;PWD=" & "dpddpd" & "")
    Set getdbFromConfig = dbResult
    Set dbResult = Nothing
    Exit Function
EH:
    Set getdbFromConfig = Nothing
    Debug.Print "Entorno.getdbFromConfig: error al abrir " & p_BackendKey & ": " & Err.Description
End Function

' Devuelve la carpeta padre de p_FullPath. Si p_FullPath esta vacio o
' no tiene padre reconocible, devuelve "". Se usa para extraer el
' directorio comun de los backends cacheados y pasarselo a
' RelinkAllTables (que necesita un dir, no un path de archivo).
Private Function ExtractParentDir(ByVal p_FullPath As String) As String
    On Error Resume Next
    Dim sPath As String
    sPath = Trim$(p_FullPath)
    If Len(sPath) = 0 Then
        ExtractParentDir = ""
        Exit Function
    End If
    ' Quitar barra final si la hay.
    Do While Len(sPath) > 0 And (Right$(sPath, 1) = "\" Or Right$(sPath, 1) = "/")
        sPath = Left$(sPath, Len(sPath) - 1)
    Loop
    ExtractParentDir = fso.GetParentFolderName(sPath)
    If Err.Number <> 0 Or Len(ExtractParentDir) = 0 Then
        ExtractParentDir = ""
    End If
    Err.Clear
    On Error GoTo 0
End Function





