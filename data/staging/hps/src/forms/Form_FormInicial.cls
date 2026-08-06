Attribute VB_Name = "Form_FormInicial"
Option Compare Database
Option Explicit

' --- Local-mode UI constants (project rule: Private Const BEFORE first Private member/Sub) ---
Private Const k_LocalModeTempVar As String = "DatosEnLocal"
Private Const k_LocalModeYes As String = "Sí"
Private Const k_LocalModeNo As String = "No"

Private m_Error As String

Private Sub Form_Timer()
    ' Matamos el timer para que no se repita
    Me.TimerInterval = 0
    
    ' Forzamos a Access a pintar el formulario AHORA MISMO por si acaso
    DoEvents
    
    ' Ejecutamos la carga pesada
    EjecutarCargaInicial
End Sub

Private Sub EjecutarCargaInicial()
    Dim sError As String
    
    On Error GoTo errores
    Me.lblEstado.Visible = True
    ' --- PASO 1: CARGAR ENTORNO (EVE) ---
    ' Avisamos al usuario antes de que se congele los 5 segundos
    AvanceNuevo "Cargando configuración del entorno..."
    
    ' Aquí llamamos a EVE. El usuario verá el formulario y el mensaje anterior
    ' mientras espera esos 5 segundos.
    EVE
    
'    ' Si usabas EstablecerControlesDatos para poner colores bonitos, LLÁMALO AHORA,
'    ' una vez que EVE ya ha cargado los objetos de estilo.
'    EstablecerControlesDatos Me, sError
    Me.lblEstado.Visible = True
    
    ' --- PASO 2: SINCRONIZACIÓN DE DATOS ---
    AvanceNuevo "Sincronizando tablas locales..."
    EjecutarSincronizacionCargaInicialFrontend p_Error:=sError
    
    If sError <> "" Then
        MsgBox "Error en la carga: " & sError, vbCritical
        Me.lblEstado.Caption = "Error."
    Else
        ' --- FINALIZAR ---
        AvanceNuevo "Carga completada."
        ' Habilitamos botones
        Me.ComandoAccesoLectura.Enabled = True
        If m_EnOficina = Empty Then
        sError = "No se puede determinar si estamos o no en la oficina"
            Err.Raise 1000
        End If
        If m_EnOficina = EnumSiNo.No Then
            If m_AccesoADatosTE = 0 Then
                sError = "No se puede determinar si estamos o no con el CARU puesto"
                Err.Raise 1000
            End If
        End If
        If m_EnOficina = EnumSiNo.Sí Then
            Me.SituacionDatos = "En Oficina"
            Me.AccesoRedTelefonica = "Sí"
            Me.ComandoAccesoEscritura.Enabled = True
            Me.ComandoAccesoLectura.Enabled = True
            
        Else
            Me.SituacionDatos = "En Local"
            If m_AccesoADatosTE = EnumSiNo.Sí Then
                Me.AccesoRedTelefonica = "Sí"
            Else
                Me.AccesoRedTelefonica = "No"
            End If
            Me.ComandoAccesoEscritura.Enabled = False
            Me.ComandoAccesoLectura.Enabled = True
        End If
        
        
        
        Me.ComandoAccesoLectura.SetFocus
    End If
    Me.lblEstado.Visible = False
    Exit Sub

errores:
    MsgBox "Error crítico: " & Err.Description, vbCritical
End Sub
Private Sub ComandoAccesoEscritura_Click()
   On Error GoTo errores
    DoCmd.Hourglass True
    m_Error = ""
    CerrarFormulariosAbiertos
    EstablecerSoloLectura p_SoloLectura:=EnumSiNo.No, p_Error:=m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    AvanceCerrar
    DoCmd.OpenForm "FormInicial00Principal"
    DoCmd.Hourglass False
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAccesoLectura_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoAccesoLectura_Click()
     On Error GoTo errores
    DoCmd.Hourglass True
    m_Error = ""
    
    If m_EnOficina = EnumSiNo.No Then
        If m_ObjUsuarioConectado.PermisoEntradaLocalEnCasa <> EnumSiNo.Sí Then
            m_Error = "No está incluido entre los usuarios a los que se les permite entrar desde casa"
            Err.Raise 1000
        End If
    End If
    CerrarFormulariosAbiertos
    EstablecerSoloLectura p_SoloLectura:=EnumSiNo.Sí, p_Error:=m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    'EVE
    
    AvanceCerrar
    DoCmd.OpenForm "FormInicial00Principal"
    DoCmd.Hourglass False
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAccesoLectura_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoSalir_Click()

    On Error Resume Next
    pregunta = MsgBox("¿Deseas salir de la aplicación", vbExclamation + vbYesNo + vbDefaultButton2)
    If pregunta <> vbYes Then
        Exit Sub
    End If

    DoCmd.Quit acQuitSaveNone
End Sub

' [HOTFIX] Source-only local-mode config UI: a single command button that
' reads/writes Application.TempVars("DatosEnLocal") and (optionally) prompts
' for a sandbox backend path. Replaces direct VBE editing of VariablesEntorno.
' The actual re-resolution of m_URLRutaAplicacionesLocal still happens on the
' next EVE/ApplicationRun; this button only flips the TempVar so the gate
' "If Application.TempVars(""DatosEnLocal"") = ""Sí"" Then ..." in
' VariablesEntorno.getdbLanzadera / getdb picks the local backend.
'
' [PR B] El boton ahora ofrece DOS acciones, elegidas con Yes/No/Cancel:
'   - Sí (Yes) = toggle del TempVar DatosEnLocal (comportamiento legacy de
'     PR #8). Si lo activas, el form sigue forzando ruta local hardcoded.
'   - No  = aplicar perfil desde TbConfiguracionHPS. Pregunta Local o
'     Produccion, escribe las 5 claves via ConfigurarTablaConfiguracion_Perfil*,
'     refresca la cache de m_ObjEntorno (CargarConfiguracion) y relinka los
'     TableDefs apuntando al directorio comun de los backends cacheados.
'   - Cancelar = abortar sin tocar nada.
Private Sub ComandoConfigurarModoLocal_Click()
    Dim sCurrentState As String
    Dim sNewState As String
    Dim sBackendPath As String
    Dim sResponse As String
    Dim sParentDir As String
    Dim sBackendHPST As String
    Dim sBackendExp As String
    Dim bRelinkOk As Boolean
    On Error GoTo errores

    ' 1. Read current state from TempVar.
    sCurrentState = Application.TempVars(k_LocalModeTempVar)
    If sCurrentState <> k_LocalModeYes And sCurrentState <> k_LocalModeNo Then
        sCurrentState = "(sin inicializar)"
    End If

    ' 2. Show current state and ask user what to do.
    sResponse = MsgBox( _
        "Modo local actual: " & sCurrentState & vbCrLf & vbCrLf & _
        "Pulsa Sí para TOGGLE del TempVar (activar/desactivar modo local)." & vbCrLf & _
        "Pulsa No para APLICAR PERFIL desde TbConfiguracionHPS (Local o Produccion, con relink).", _
        vbQuestion + vbYesNoCancel + vbDefaultButton2, _
        "Configurar modo local")
    If sResponse = vbCancel Then
        Exit Sub
    End If

    ' --- Branch 1: Toggle del TempVar (legacy de PR #8) ---
    If sResponse = vbYes Then
        If sCurrentState = k_LocalModeYes Then
            sNewState = k_LocalModeNo
        Else
            sNewState = k_LocalModeYes
        End If

        ' Si activamos, pedimos la ruta opcional del backend sandbox.
        sBackendPath = ""
        If sNewState = k_LocalModeYes Then
            sBackendPath = InputBox( _
                "Ruta del backend sandbox (opcional)." & vbCrLf & _
                "Déjalo vacío para usar la ruta por defecto (TbConfiguracionBackends).", _
                "Modo local — backend", "")
        End If

        ' Persistimos el TempVar (esto es lo que el resto de la app lee).
        Application.TempVars(k_LocalModeTempVar) = sNewState

        ' Confirmamos e indicamos al usuario que el cambio se aplica al reiniciar.
        MsgBox _
            "Modo local actualizado a: " & sNewState & vbCrLf & _
            "Backend: " & IIf(Len(sBackendPath) = 0, "(por defecto)", sBackendPath) & vbCrLf & vbCrLf & _
            "El cambio se aplicará al reiniciar la aplicación.", _
            vbInformation, "Modo local"
        Exit Sub
    End If

    ' --- Branch 2: Aplicar perfil desde TbConfiguracionHPS (PR B) ---
    sResponse = MsgBox( _
        "¿Qué perfil quieres aplicar?" & vbCrLf & _
        "Sí = Local (rutas de la maquina de desarrollo)." & vbCrLf & _
        "No  = Produccion (rutas UNC de red corporativa).", _
        vbQuestion + vbYesNoCancel + vbDefaultButton2, _
        "Aplicar perfil")
    If sResponse = vbCancel Then
        Exit Sub
    End If

    ' Escribimos las 5 claves + poblamos TbVinculosTablas con los TableDefs actuales.
    If sResponse = vbYes Then
        ConfigurarTablaConfiguracion_PerfilLocal
    Else
        ConfigurarTablaConfiguracion_PerfilProduccion
    End If

    ' Refrescamos la cache de m_ObjEntorno para que vea los valores nuevos.
    m_ObjEntorno.CargarConfiguracion

    ' Extraemos el directorio comun a partir de las rutas de backend cacheadas.
    ' Si el padre de HPST coincide con el de Expedientes, ese es el dir comun;
    ' si no, usamos el de HPST (es el backend principal de la app).
    sBackendHPST = m_ObjEntorno.GetConfig("HPST_BACKEND_PATH")
    sBackendExp = m_ObjEntorno.GetConfig("EXPEDIENTES_BACKEND_PATH")
    sParentDir = CommonParentDir(sBackendHPST, sBackendExp)
    If Len(sParentDir) = 0 Then
        sParentDir = ParentDirOfPath(sBackendHPST)
    End If

    bRelinkOk = False
    If Len(sParentDir) > 0 Then
        bRelinkOk = m_ObjEntorno.RelinkAllTables(sParentDir)
    End If

    MsgBox _
        "Perfil aplicado." & vbCrLf & _
        "Tabla de configuracion: " & m_ObjEntorno.GetConfigTablePath() & vbCrLf & _
        "Directorio de backends: " & sParentDir & vbCrLf & _
        "Relink: " & IIf(bRelinkOk, "OK", "con errores (ver Debug.Print)"), _
        vbInformation, "Aplicar perfil"
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoConfigurarModoLocal_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

' Devuelve el segmento padre comun entre dos rutas. Si los dos padres no
' coinciden, devuelve "" -- el llamador hace fallback a ParentDirOfPath.
Private Function CommonParentDir(ByVal p_Path1 As String, ByVal p_Path2 As String) As String
    Dim sParent1 As String
    Dim sParent2 As String
    sParent1 = ParentDirOfPath(p_Path1)
    sParent2 = ParentDirOfPath(p_Path2)
    If Len(sParent1) > 0 And StrComp(sParent1, sParent2, vbTextCompare) = 0 Then
        CommonParentDir = sParent1
    Else
        CommonParentDir = ""
    End If
End Function

' Devuelve el directorio padre de p_FullPath usando fso.GetParentFolderName.
' Devuelve "" si la ruta esta vacia o no se puede analizar.
Private Function ParentDirOfPath(ByVal p_FullPath As String) As String
    On Error Resume Next
    Dim sPath As String
    sPath = Trim$(p_FullPath)
    If Len(sPath) = 0 Then
        ParentDirOfPath = ""
        Exit Function
    End If
    ' Quitar barra final si la hay.
    Do While Len(sPath) > 0 And (Right$(sPath, 1) = "\" Or Right$(sPath, 1) = "/")
        sPath = Left$(sPath, Len(sPath) - 1)
    Loop
    ParentDirOfPath = fso.GetParentFolderName(sPath)
    If Err.Number <> 0 Or Len(ParentDirOfPath) = 0 Then
        ParentDirOfPath = ""
    End If
    Err.Clear
    On Error GoTo 0
End Function



Private Sub Form_Load()
    ' [UX] Hide the Access ribbon for a cleaner application look.
    ' Wrapped in On Error Resume Next so a misnamed toolbar (e.g. on a
    ' localized build) doesn't break the form load. The form is a splash
    ' form that opens immediately at app start, so the ribbon gets
    ' hidden before EVE / cache sync blocks the UI thread.
    On Error Resume Next
    DoCmd.ShowToolbar "Ribbon", acToolbarNo
    On Error GoTo 0

    ' 1. Configuración visual básica (sin llamar a objetos globales pesados)
    Me.ComandoAccesoLectura.Enabled = False
    Me.ComandoAccesoEscritura.Enabled = False
    
    ' 2. Mensaje inicial para que el usuario sepa que algo va a pasar
    Me.lblEstado.Caption = "Arrancando sistema..."
    
    ' 3. Iniciamos el temporizador.
    ' 500ms es suficiente para que Access pinte la ventana antes de bloquearse con EVE.
    Me.TimerInterval = 500
End Sub

