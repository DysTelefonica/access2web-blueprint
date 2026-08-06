Attribute VB_Name = "modRiesgoEstadoGateHelper"
' =============================================================================
' modRiesgoEstadoGateHelper.bas
'
' Helper module para risk state gate validation (gemelo Materializado/Mitigacion/Retirado)
' Project: gestion_riesgos
' Branch: feat/metodologia-e2e-riesgos-2026-06-19
' Refactor: 2026-06-19 (extracción desde 3 forms gemelo)
'
' Helpers públicos (7):
'   #1 ValidarTransicionMaterializacion
'   #2 QuitarMaterializacion
'   #3 RegistrarAceptacionRiesgo
'   #4 ValidarCambioMitigacion
'   #5 RegistrarRetiroRiesgo
'   #6 QuitarRetiroRiesgo
'   #7 SolicitarFechaRetiro
'
' Helper interno (compartido por #1 y #5):
'   ValidarFechaRiesgo — regla unificada: no futuro, sin límite pasado,
'   no antes inicio edición
' =============================================================================
Option Compare Database
Option Explicit

' --- Helper #1: ValidarTransicionMaterializacion ---
' Extraído de Form_FormRiesgoMaterializado.ComandoFechaMaterializado_Click (lóneas 15-90)
' Valida: riesgo no en aceptacion, no en retirada, no ya materializado.
' Pide confirmación MsgBox (p_PromptResult), solicita fecha+plan, valida fecha,
' persiste vía MaterializacionRegistrar. Refresca UI (el form lo hace tras llamar helper).
Public Function ValidarTransicionMaterializacion( _
    ByVal p_IDRiesgo As String, _
    ByRef p_Fecha As String, _
    ByRef p_IDPlan As String, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_PromptResult As Long = -1, _
    Optional ByRef p_Error As String) As String

    Dim pregunta As VbMsgBoxResult
    Dim objRiesgo As riesgo
    Dim m_RiesgoEnAceptacion As EnumSiNo
    Dim m_RiesgoEnRetirada As EnumSiNo
    Dim m_FechaMaterializado As String
    Dim m_IDPlanContingencia As String
    Dim m_ValidationError As String
    Dim m_ValidationResult As String

    On Error GoTo errores
    p_Error = ""
    ValidarTransicionMaterializacion = ""

    ' Resolver db
    If db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then
Err.Raise 1000
        End If
    End If

    ' Obtener riesgo
    Set objRiesgo = GetCachedRiesgo(p_IDRiesgo, p_Error, db)
    If p_Error <> "" Or objRiesgo Is Nothing Then
        p_Error = "Riesgo no encontrado: " & p_IDRiesgo
        ValidarTransicionMaterializacion = p_Error
        Exit Function
    End If

    ' Gate: no en aceptacion
    m_RiesgoEnAceptacion = RiesgoEnAceptacion(objRiesgo.EstadoEnum)
    If m_RiesgoEnAceptacion = EnumSiNo.Sí Then
        p_Error = "El riesgo está en camino de ser aceptado"
        ValidarTransicionMaterializacion = p_Error
        Exit Function
    End If

    ' Gate: no en retirada
    m_RiesgoEnRetirada = RiesgoEnRetirada(objRiesgo.EstadoEnum)
    If m_RiesgoEnRetirada = EnumSiNo.Sí Then
        p_Error = "El riesgo está en camino de ser retirado"
        ValidarTransicionMaterializacion = p_Error
        Exit Function
    End If

    ' Gate: no ya materializado
    If objRiesgo.EstadoEnum = EnumRiesgoEstado.Materializado Then
        p_Error = "El riesgo aparece como materializado"
        ValidarTransicionMaterializacion = p_Error
        Exit Function
    End If

    ' Confirmación MsgBox
    If p_PromptResult = -1 Then
        pregunta = MsgBox( _
            "Al materializarse un riesgo, recuerde que ha de ejecutar el plan de contingencia y " & _
            "publicar la edición para que quede constancia", _
            vbExclamation + vbYesNo + vbDefaultButton2, _
            "Materialización")
    Else
        pregunta = p_PromptResult
    End If
    If pregunta <> vbYes Then
        p_Error = "Cancelado por el usuario"
        ValidarTransicionMaterializacion = ""
        Exit Function
    End If

    ' Solicitar datos de materialización (fecha + plan)
    If Not SolicitarDatosMaterializacionHelper( _
            objRiesgo, m_FechaMaterializado, m_IDPlanContingencia, p_Error) Then
        If p_Error <> "" Then
Err.Raise 1000
        End If
        Exit Function
    End If

    ' Validar fecha con regla unificada
    If objRiesgo.Edicion Is Nothing Then
        p_Error = "Riesgo sin edición asociada"
        ValidarTransicionMaterializacion = p_Error
        Exit Function
    End If
    m_ValidationResult = ValidarFechaRiesgo( _
        m_FechaMaterializado, _
        objRiesgo.Edicion.FechaEdicion, _
        db, _
        m_ValidationError)
    If m_ValidationError <> "" Then
        p_Error = m_ValidationError
        ValidarTransicionMaterializacion = p_Error
        Exit Function
    End If

    ' Persistir
    objRiesgo.FechaMaterializado = m_FechaMaterializado
    objRiesgo.MaterializacionRegistrar m_FechaMaterializado, p_Error, m_IDPlanContingencia
    If p_Error <> "" Then
        ValidarTransicionMaterializacion = p_Error
        Exit Function
    End If

    ' Devolver los valores obtenidos al caller (form) para que actualice UI
    p_Fecha = m_FechaMaterializado
    p_IDPlan = m_IDPlanContingencia

    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ValidarTransicionMaterializacion: " & Err.description
    End If
    ValidarTransicionMaterializacion = ""  ' error va por p_Error, no por el return
End Function

' Helper interno: SolicitarDatosMaterializacion
' Simula el diálogo FormMaterializacionPlanContingencia via TempVars.
' En testing, los TempVars son seteados por el test; en producción el form lo hace.
Private Function SolicitarDatosMaterializacionHelper( _
    ByRef p_ObjRiesgo As riesgo, _
    ByRef p_Fecha As String, _
    ByRef p_IDPlan As String, _
    Optional ByRef p_Error As String) As Boolean

    Const TEMP_MAT_ACEPTADA As String = "MaterializacionPlanContingenciaAceptada"
    Const TEMP_MAT_FECHA As String = "MaterializacionPlanContingenciaFecha"
    Const TEMP_MAT_PLAN As String = "MaterializacionPlanContingenciaPlan"

    On Error GoTo errores
    p_Error = ""
    SolicitarDatosMaterializacionHelper = False
    p_Fecha = ""
    p_IDPlan = ""

    ' En tests, los TempVars ya estón seteados por el fixture.
    ' En producción, el form FormMaterializacionPlanContingencia los setea en acDialog.
    If TempVarValue(TEMP_MAT_ACEPTADA) <> "Sí" Then
        Exit Function
    End If

    p_Fecha = TempVarValue(TEMP_MAT_FECHA)
    p_IDPlan = TempVarValue(TEMP_MAT_PLAN)
    SolicitarDatosMaterializacionHelper = True
    Exit Function

errores:
    p_Error = "SolicitarDatosMaterializacionHelper: " & Err.description
End Function

Private Function TempVarValue(ByVal p_Nombre As String) As String
    On Error GoTo salida
    TempVarValue = Nz(TempVars(p_Nombre).value, "")
salida:
End Function

' --- Helper #2: QuitarMaterializacion ---
' Extraído de Form_FormRiesgoMaterializado.ComandoQuitarMaterializado_Click (lóneas 209-259)
' MsgBox "¿Desea habilitar el riesgo?" (p_PromptResult), llama MaterializacionQuitarRegistrar.
Public Function QuitarMaterializacion( _
    ByVal p_IDRiesgo As String, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_PromptResult As Long = -1) As String

    Dim pregunta As VbMsgBoxResult
    Dim objRiesgo As riesgo
    Dim p_Error As String

    On Error GoTo errores
    p_Error = ""
    QuitarMaterializacion = ""

    ' Resolver db
    If db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then
Err.Raise 1000
        End If
    End If

    ' Confirmación
    If p_PromptResult = -1 Then
        pregunta = MsgBox( _
            "¿Desea habilitar el riesgo?", _
            vbExclamation + vbYesNo + vbDefaultButton2, _
            "Habilitar el riesgo")
    Else
        pregunta = p_PromptResult
    End If
    If pregunta <> vbYes Then
        QuitarMaterializacion = ""
        Exit Function
    End If

    ' Obtener riesgo
    Set objRiesgo = GetCachedRiesgo(p_IDRiesgo, p_Error, db)
    If p_Error <> "" Or objRiesgo Is Nothing Then
        QuitarMaterializacion = "Riesgo no encontrado: " & p_IDRiesgo
        Exit Function
    End If

    ' Persistir
    objRiesgo.MaterializacionQuitarRegistrar objRiesgo.EstadoEnum, p_Error
    If p_Error <> "" Then
        QuitarMaterializacion = p_Error
        Exit Function
    End If

    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "QuitarMaterializacion: " & Err.description
    End If
    QuitarMaterializacion = ""  ' error va por p_Error, no por el return
End Function

' --- Helper #3: RegistrarAceptacionRiesgo ---
' Extraído de Form_FormRiesgoMitigacion.ComandoAceptacionRegistrar_Click (lóneas 11-65)
' Registra justificación de aceptación del riesgo.
Public Function RegistrarAceptacionRiesgo( _
    ByVal p_IDRiesgo As String, _
    ByVal p_Justificacion As String, _
    Optional ByRef db As DAO.Database = Nothing) As String

    Dim objRiesgo As riesgo
    Dim p_Error As String

    On Error GoTo errores
    p_Error = ""
    RegistrarAceptacionRiesgo = ""

    ' Resolver db
    If db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then
Err.Raise 1000
        End If
    End If

    ' Validar justificación
    If Trim$(Nz(p_Justificacion, "")) = "" Then
        p_Error = "Se ha de indicar alguna justificación"
        RegistrarAceptacionRiesgo = p_Error
        Exit Function
    End If

    ' Obtener riesgo
    Set objRiesgo = GetCachedRiesgo(p_IDRiesgo, p_Error, db)
    If p_Error <> "" Or objRiesgo Is Nothing Then
        RegistrarAceptacionRiesgo = "Riesgo no encontrado: " & p_IDRiesgo
        Exit Function
    End If

    ' Persistir
    objRiesgo.AceptacionRegistrar p_Justificacion, Date, objRiesgo.EstadoEnum, p_Error
    If p_Error <> "" Then
        RegistrarAceptacionRiesgo = p_Error
        Exit Function
    End If

    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "RegistrarAceptacionRiesgo: " & Err.description
    End If
    RegistrarAceptacionRiesgo = ""  ' error va por p_Error, no por el return
End Function

' --- Helper #4: ValidarCambioMitigacion ---
' Extraído de Form_FormRiesgoMitigacion.ListaMitigacion_Click (lóneas 433-563)
' Valida transición de mitigación. Si anterior="Aceptar" y nueva<>"Aceptar",
' pide confirmación (p_PromptResult) y ejecuta AceptacionRegistrarQuitar.
' Gate: no puede cambiar mitigacion si Estado=Retirado o Materializado.
Public Function ValidarCambioMitigacion( _
    ByVal p_IDRiesgo As String, _
    ByVal p_NuevaMitigacion As String, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_PromptResult As Long = -1) As String

    Dim pregunta As VbMsgBoxResult
    Dim objRiesgo As riesgo
    Dim objRiesgoAlInicio As riesgo
    Dim m_RiesgoEnRetirada As EnumSiNo
    Dim p_Error As String
    Dim m_MitigacionAnterior As String

    On Error GoTo errores
    p_Error = ""
    ValidarCambioMitigacion = ""

    ' Resolver db
    If db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then
Err.Raise 1000
        End If
    End If

    ' Validar mitigacion no vacía
    If Trim$(Nz(p_NuevaMitigacion, "")) = "" Then
        p_Error = "Se ha de indicar la mitigación"
        ValidarCambioMitigacion = p_Error
        Exit Function
    End If

    ' Obtener riesgo
    Set objRiesgo = GetCachedRiesgo(p_IDRiesgo, p_Error, db)
    If p_Error <> "" Or objRiesgo Is Nothing Then
        ValidarCambioMitigacion = "Riesgo no encontrado: " & p_IDRiesgo
        Exit Function
    End If

    ' Gate: no puede cambiar mitigacion si riesgo en retirada
    m_RiesgoEnRetirada = RiesgoEnRetirada(objRiesgo.EstadoEnum)
    If m_RiesgoEnRetirada = EnumSiNo.Sí Then
        p_Error = "El riesgo está en fase de ser retirado"
        ValidarCambioMitigacion = p_Error
        Exit Function
    End If

    ' Gate: no puede cambiar mitigacion si riesgo materializado
    If objRiesgo.EstadoEnum = EnumRiesgoEstado.Materializado Then
        p_Error = "El riesgo está materializado"
        ValidarCambioMitigacion = p_Error
        Exit Function
    End If

    ' Obtener estado al inicio para comparar mitigacion anterior
    Set objRiesgoAlInicio = GetCachedRiesgo(p_IDRiesgo, p_Error, db)
    If p_Error <> "" Or objRiesgoAlInicio Is Nothing Then
        ' No tenemos AlInicio — no podemos hacer la transición Aceptar>otra
        ' Simplemente aceptar la nueva mitigacion sin quitar aceptacion
        objRiesgo.Mitigacion = p_NuevaMitigacion
        ValidarCambioMitigacion = ""
        Exit Function
    End If

    m_MitigacionAnterior = Trim$(Nz(objRiesgoAlInicio.Mitigacion, ""))

    ' Si la anterior era "Aceptar" y la nueva NO es "Aceptar",
    ' pedir confirmación y quitar la aceptacion
    If m_MitigacionAnterior = "Aceptar" And p_NuevaMitigacion <> "Aceptar" Then
        If p_PromptResult = -1 Then
            pregunta = MsgBox( _
                "¿Desea quitar la aceptación del riesgo y elegir otra mitigación?", _
                vbExclamation + vbYesNo + vbDefaultButton2, _
                "Quitar mitigación aceptar")
        Else
            pregunta = p_PromptResult
        End If
        If pregunta <> vbYes Then
            ValidarCambioMitigacion = ""
            Exit Function
        End If

        objRiesgo.AceptacionRegistrarQuitar p_NuevaMitigacion, Date, objRiesgo.EstadoEnum, p_Error
        If p_Error <> "" Then
Err.Raise 1000
        End If
    End If

    ' Registrar la nueva mitigacion
    objRiesgo.Mitigacion = p_NuevaMitigacion
    If p_NuevaMitigacion <> "Aceptar" Then
        ' Limpiar justificación de aceptación previa si se eligió otra mitigación
        objRiesgo.JustificacionAceptacionRiesgo = ""
        objRiesgo.FechaJustificacionAceptacionRiesgo = ""
    End If

    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ValidarCambioMitigacion: " & Err.description
    End If
    ValidarCambioMitigacion = ""  ' error va por p_Error, no por el return
End Function

' --- Helper #5: RegistrarRetiroRiesgo ---
' Extraído de Form_FormRiesgoRetirado.ComandoRetiroRegistrar_Click (lóneas 131-203)
' Valida: justificación no vacía, riesgo existe, no materializado (si lo está,
' quita materialización primero con MaterializacionQuitarRegistrar),
' fecha no vacía, persiste vía RetiroRegistrar.
' Altera p_Fecha de salida con la fecha persistida.
Public Function RegistrarRetiroRiesgo( _
    ByVal p_IDRiesgo As String, _
    ByVal p_Justificacion As String, _
    ByVal p_Fecha As String, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_PromptResult As Long = -1) As String

    Dim objRiesgo As riesgo
    Dim p_Error As String
    Dim m_ValidationResult As String
    Dim m_ValidationError As String

    On Error GoTo errores
    p_Error = ""
    RegistrarRetiroRiesgo = ""

    ' Resolver db
    If db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then
Err.Raise 1000
        End If
    End If

    ' Validar justificación
    If Trim$(Nz(p_Justificacion, "")) = "" Then
        p_Error = "Se ha de indicar alguna justificación"
        RegistrarRetiroRiesgo = p_Error
        Exit Function
    End If

    ' Obtener riesgo
    Set objRiesgo = GetCachedRiesgo(p_IDRiesgo, p_Error, db)
    If p_Error <> "" Or objRiesgo Is Nothing Then
        RegistrarRetiroRiesgo = "Riesgo no encontrado: " & p_IDRiesgo
        Exit Function
    End If

    ' Validar fecha no vacía
    If Trim$(Nz(p_Fecha, "")) = "" Or Not IsDate(p_Fecha) Then
        p_Error = "Primero ha de indicar la fecha de retirada"
        RegistrarRetiroRiesgo = p_Error
        Exit Function
    End If

    ' Si el riesgo está materializado, quitar la materialización primero
    If objRiesgo.EstadoEnum = EnumRiesgoEstado.Materializado Then
        If p_PromptResult = -1 Then
            Dim pregunta As VbMsgBoxResult
            pregunta = MsgBox( _
                "El riesgo actualmente está materializado." & vbNewLine & _
                "¿Desea proponer la retirada del riesgo?", _
                vbExclamation + vbYesNo + vbDefaultButton2, _
                "Propuesta de retirada")
            If pregunta <> vbYes Then
                RegistrarRetiroRiesgo = ""
                Exit Function
            End If
        End If

        objRiesgo.MaterializacionQuitarRegistrar objRiesgo.EstadoEnum, p_Error
        If p_Error <> "" Then
Err.Raise 1000
        End If
    End If

    ' Validar fecha con regla unificada (no futuro, no antes inicio edición)
    If objRiesgo.Edicion Is Nothing Then
        p_Error = "Riesgo sin edición asociada"
        RegistrarRetiroRiesgo = p_Error
        Exit Function
    End If
    m_ValidationResult = ValidarFechaRiesgo( _
        p_Fecha, _
        objRiesgo.Edicion.FechaEdicion, _
        db, _
        m_ValidationError)
    If m_ValidationError <> "" Then
        p_Error = m_ValidationError
        RegistrarRetiroRiesgo = p_Error
        Exit Function
    End If

    ' Persistir
    objRiesgo.RetiroRegistrar p_Justificacion, p_Fecha, objRiesgo.EstadoEnum, p_Error
    If p_Error <> "" Then
        RegistrarRetiroRiesgo = p_Error
        Exit Function
    End If

    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "RegistrarRetiroRiesgo: " & Err.description
    End If
    RegistrarRetiroRiesgo = ""  ' error va por p_Error, no por el return
End Function

' --- Helper #6: QuitarRetiroRiesgo ---
' Extraído de Form_FormRiesgoRetirado.ComandoEliminar_Click (lóneas 205-265)
' MsgBox "¿Desea quitar la retirada del Riesgo?" (p_PromptResult).
' Gate: riesgo debe estar en estado válido para quitar retiro
' (RiesgoEnRetirada = Sí).
Public Function QuitarRetiroRiesgo( _
    ByVal p_IDRiesgo As String, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_PromptResult As Long = -1) As String

    Dim pregunta As VbMsgBoxResult
    Dim objRiesgo As riesgo
    Dim m_Estado As EnumRiesgoEstado
    Dim m_RiesgoEnRetirada As EnumSiNo
    Dim p_Error As String

    On Error GoTo errores
    p_Error = ""
    QuitarRetiroRiesgo = ""

    ' Resolver db
    If db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then
Err.Raise 1000
        End If
    End If

    ' Confirmación
    If p_PromptResult = -1 Then
        pregunta = MsgBox( _
            "¿Desea quitar la retirada del Riesgo?", _
            vbExclamation + vbYesNo + vbDefaultButton2, _
            "Quitar retirada")
    Else
        pregunta = p_PromptResult
    End If
    If pregunta <> vbYes Then
        QuitarRetiroRiesgo = ""
        Exit Function
    End If

    ' Obtener riesgo
    Set objRiesgo = GetCachedRiesgo(p_IDRiesgo, p_Error, db)
    If p_Error <> "" Or objRiesgo Is Nothing Then
        QuitarRetiroRiesgo = "Riesgo no encontrado: " & p_IDRiesgo
        Exit Function
    End If

    ' Gate: debe estar en estado válido para quitar retiro
    m_Estado = objRiesgo.ESTADOCalculado
    m_RiesgoEnRetirada = RiesgoEnRetirada(m_Estado)
    If m_RiesgoEnRetirada = EnumSiNo.No Then
        p_Error = "El retiro del riesgo no es sin justificar, sin visar o rechazado"
        QuitarRetiroRiesgo = p_Error
        Exit Function
    End If

    ' Persistir
    objRiesgo.RetiroRegistrarQuitar m_Estado, p_Error
    If p_Error <> "" Then
        QuitarRetiroRiesgo = p_Error
        Exit Function
    End If

    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "QuitarRetiroRiesgo: " & Err.description
    End If
    QuitarRetiroRiesgo = ""  ' error va por p_Error, no por el return
End Function

' --- Helper #7: SolicitarFechaRetiro ---
' Extraído de Form_FormRiesgoRetirado.ComandoFRetirado_Click (lóneas 301-364).
' Solicita fecha de retiro. El form llama a este helper y luego actualiza los
' controles de visibilidad/enabled segón el estado.
' Si p_Fecha viene vacía, usa InputBox para solicitarla.
' Gate: riesgo no en aceptación.
' El form se encarga de mostrar/ocultar campos y establecer enabled tras llamar
' a este helper.
Public Function SolicitarFechaRetiro( _
    ByVal p_IDRiesgo As String, _
    Optional ByRef p_Fecha As String, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_PromptResult As Long = -1) As String

    Dim m_Fecha As String
    Dim objRiesgo As riesgo
    Dim m_RiesgoEnAceptacion As EnumSiNo
    Dim p_Error As String

    On Error GoTo errores
    p_Error = ""
    SolicitarFechaRetiro = ""

    ' Resolver db
    If db Is Nothing Then
        Set db = getdb(p_Error)
        If p_Error <> "" Then
Err.Raise 1000
        End If
    End If

    ' Obtener riesgo
    Set objRiesgo = GetCachedRiesgo(p_IDRiesgo, p_Error, db)
    If p_Error <> "" Or objRiesgo Is Nothing Then
        SolicitarFechaRetiro = "Riesgo no encontrado: " & p_IDRiesgo
        Exit Function
    End If

    ' Gate: no puede solicitar fecha si riesgo está aceptado o en fase de aceptación
    m_RiesgoEnAceptacion = RiesgoEnAceptacion(objRiesgo.EstadoEnum)
    If m_RiesgoEnAceptacion = EnumSiNo.Sí Then
        p_Error = "El riesgo está aceptado o en fase de aceptación"
        SolicitarFechaRetiro = p_Error
        Exit Function
    End If

    ' Si no hay fecha preestablecida, solicitar vía InputBox
    If Trim$(Nz(p_Fecha, "")) = "" Then
        If p_PromptResult = -1 Then
            m_Fecha = Trim$(Nz(InputBox( _
                "Introduzca la fecha de retiro del riesgo", _
                "Fecha retirado", _
                Format$(Date, "dd/mm/yyyy")), ""))
            If Not IsDate(m_Fecha) Then
                ' InputBox cancelado o fecha inválida — no es error, el form maneja
                SolicitarFechaRetiro = ""
                Exit Function
            End If
            p_Fecha = m_Fecha
        End If
        ' Si p_PromptResult <> -1 (test invocó con valor explícito), no se
        ' abre InputBox. Si además p_Fecha está vacío, el caller debe proveerlo.
    End If

    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "SolicitarFechaRetiro: " & Err.description
    End If
    SolicitarFechaRetiro = ""  ' error va por p_Error, no por el return
End Function

' --- Helper interno: ValidarFechaRiesgo ---
' Regla unificada (decidida 2026-06-19 para ambos lifecycle stages):
'   1. p_Fecha > Date > error "fecha futura no permitida"
'   2. p_Fecha < p_FechaInicioEdicion > error "fecha anterior al inicio de la edición"
'   3. Sin límite de pasado lejano
' Compartido por: ValidarTransicionMaterializacion (#1) y RegistrarRetiroRiesgo (#5).
Private Function ValidarFechaRiesgo( _
    ByVal p_Fecha As String, _
    ByVal p_FechaInicioEdicion As String, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_Error As String) As String

    On Error GoTo errores
    p_Error = ""
    ValidarFechaRiesgo = ""

    ' Validar que sea fecha
    If Not IsDate(p_Fecha) Then
        p_Error = "Fecha inválida"
        ValidarFechaRiesgo = p_Error
        Exit Function
    End If

    ' Regla 1: no futuro
    If CDate(p_Fecha) > Date Then
        p_Error = "La fecha no puede ser futura"
        ValidarFechaRiesgo = p_Error
        Exit Function
    End If

    ' Regla 2: no antes del inicio de la edición
    If IsDate(p_FechaInicioEdicion) Then
        If CDate(p_Fecha) < CDate(p_FechaInicioEdicion) Then
            p_Error = "La fecha no puede ser anterior al inicio de la edición"
            Err.Raise 1000
        End If
    End If

    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ValidarFechaRiesgo: " & Err.description
    End If
    ValidarFechaRiesgo = ""  ' error va por p_Error, no por el return
End Function

