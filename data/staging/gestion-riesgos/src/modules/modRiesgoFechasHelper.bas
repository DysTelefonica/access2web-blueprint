Attribute VB_Name = "modRiesgoFechasHelper"
Option Compare Database
Option Explicit

' =============================================================================
' modRiesgoFechasHelper.bas
'
' Helper module para validación centralizada de fechas del riesgo.
' Project: gestion_riesgos
' Branch: feat/metodologia-e2e-riesgos-2026-06-19
' SDD:     e2e-form-by-form-2026-06-22 - Bloque 4 - REQ-CAL-04
'
' Helper público (1):
'   ValidarFechaCampo - valida que la fecha del campo sea aceptable para
'                       el formulario que la persiste. Regla centralizada.
'
' DECISIÓN 2026-06-22 (sdd-apply):
'   * Helper puro (no DAO): la regla es función pura de (p_NombreCampo,
'     p_Valor). Se incluye `Optional ByRef db As DAO.Database = Nothing`
'     por consistencia con el patrón del proyecto (skill
'     access-vba-e2e-methodology §hard rule 2).
'   * Devuelve Boolean (no String JSON). Mismo estilo que
'     modRiesgoPriorizacionHelper.bas y modPlanAccionAuthorizationHelper.bas.
'   * p_Error ByRef es OBLIGATORIO (convención Telefónica D&S); el caller lo
'     inspecciona y muestra el mensaje al usuario.
'   * Para FechaMaterializado el helper preserva la semántica de
'     Riesgo.ValidarFechaMaterializacionPermitida (Riesgo.cls:6219):
'     fecha futura ? False. Esto MANTIENE VERDES los 4 átomos
'     Test_RiesgoMaterializacion_FechaFuturaBloqueada /
'     _FechaHoyPermitida / _FechaPasadaPermitida /
'     _RegistrarFechaFuturaNoPersiste.
'   * La rama "acepta fechas pasadas" es la lectura operativa propuesta
'     para refrendar con Calidad. Si Calidad no refrenda esta lectura,
'     el átomo 3 (Edge) se cae pero el helper sigue siendo válido para
'     el resto.
'
' Regla de negocio (lectura operativa propuesta):
'   - FechaInicio y FechaFinPrevista: PUEDEN ser futuras (planificación).
'   - FechaMaterializado: NO PUEDE ser futura (operación ya ocurrida).
'   - Resto de campos de fecha del riesgo (FechaDetectado, FechaRetirado,
'     FechaAprobacionAceptacionPorCalidad, FechaRechazoAceptacionPorCalidad,
'     FechaAprobacionRetiroPorCalidad, FechaRechazoRetiroPorCalidad,
'     FechaPublicacion, FechaCierre, FechaProximaPublicacion,
'     FechaPreparadaParaPublicar, FechaEdicion): NO PUEDEN ser futuras.
'   - Fechas pasadas: se aceptan donde tengan sentido.
'   - Vacío / no-fecha: no es nuestro problema (la regla de "campo
'     requerido" la aplica el formulario).
' =============================================================================

' -----------------------------------------------------------------------------
' UserLabelFor
'
' Traduce el nombre canónico del campo a una etiqueta legible en español para
' usar en el mensaje de error que verá el usuario.
'
' Parámetros:
'   p_NombreCampo (String) - nombre del campo (canonical uppercase).
'
' Retorna:
'   Etiqueta en español (lowercase) para mensajes al usuario.
' -----------------------------------------------------------------------------
Private Function UserLabelFor(ByVal p_NombreCampo As String) As String
    Dim m_Nombre As String
    m_Nombre = UCase$(Trim$(Nz(p_NombreCampo, "")))

    Select Case m_Nombre
        Case "FECHADETECTADO"
            UserLabelFor = "fecha de detectado"
        Case "FECHARETIRADO"
            UserLabelFor = "fecha de retirado"
        Case "FECHAAPROBACIONACEPTACIONPORCALIDAD"
            UserLabelFor = "fecha de aprobación de aceptación por calidad"
        Case "FECHARECHAZOACEPTACIONPORCALIDAD"
            UserLabelFor = "fecha de rechazo de aceptación por calidad"
        Case "FECHAAPROBACIONRETIROPORCALIDAD"
            UserLabelFor = "fecha de aprobación de retiro por calidad"
        Case "FECHARECHAZORETIROPORCALIDAD"
            UserLabelFor = "fecha de rechazo de retiro por calidad"
        Case "FECHAPUBLICACION"
            UserLabelFor = "fecha de publicación"
        Case "FECHACIERRE"
            UserLabelFor = "fecha de cierre"
        Case "FECHAPROXIMAPUBLICACION"
            UserLabelFor = "fecha de próxima publicación"
        Case "FECHAPREPARADAPARAPUBLICAR"
            UserLabelFor = "fecha de preparada para publicar"
        Case "FECHAEDICION"
            UserLabelFor = "fecha de edición"
        Case "FECHAMATERIALIZADO"
            UserLabelFor = "fecha de materialización"
        Case Else
            ' Fallback: usar el nombre tal como llegó (lowercase).
            UserLabelFor = LCase$(Trim$(Nz(p_NombreCampo, "")))
    End Select
End Function

' -----------------------------------------------------------------------------
' ValidarFechaCampo
'
' Valida que el valor sea una fecha aceptable para el campo cuyo nombre se
' pasa como argumento. La regla concreta:
'   - Vacío / no-fecha  ? True  (no es nuestro problema)
'   - FechaInicio, FechaFinPrevista ? True (permiten futuro)
'   - FechaMaterializado > hoy ? False (delega semántica existente)
'   - Resto > hoy ? False con etiqueta en español
'
' Parámetros:
'   p_NombreCampo (String)            - nombre canónico del campo
'                                       (case-insensitive).
'   p_Valor (String)                  - valor de la fecha (formato VB:
'                                       dd/mm/yyyy o Variant convertible).
'   db (DAO.Database, opcional)       - presente por convención del
'                                       proyecto; no se usa.
'   p_Error (out String, opcional)    - mensaje legible si retorna False;
'                                       vacío si True.
'
' Retorna:
'   True  si la fecha es aceptable para ese campo.
'   False si la fecha es futura y NO está en la whitelist, con p_Error
'         poblado en español.
' -----------------------------------------------------------------------------
Public Function ValidarFechaCampo( _
    ByVal p_NombreCampo As String, _
    ByVal p_Valor As String, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_Error As String) As Boolean

    Dim m_Nombre As String
    Dim m_valor As String
    Dim m_Fecha As Date

    On Error GoTo errores

    p_Error = ""
    ValidarFechaCampo = False

    ' db está presente por convención del proyecto. La regla no consulta la
    ' BD; este helper es una función pura sobre (p_NombreCampo, p_Valor).
    ' El If explícito evita el warning "unused parameter" y deja claro que
    ' la decisión arquitectónica es deliberada.
    If db Is Nothing Then
        ' no-op
    End If

    m_Nombre = UCase$(Trim$(Nz(p_NombreCampo, "")))
    m_valor = Trim$(Nz(p_Valor, ""))

    ' 1. Vacío no es nuestro problema (la regla de "requerido" la aplica
    '    el formulario). Devolvemos True sin error.
    If Len(m_valor) = 0 Then
        ValidarFechaCampo = True
        Exit Function
    End If

    ' 2. No-fecha tampoco es nuestro problema (otro validador maneja
    '    tipos). Devolvemos True sin error para no duplicar mensajes.
    If Not IsDate(m_valor) Then
        ValidarFechaCampo = True
        Exit Function
    End If

    m_Fecha = CDate(m_valor)

    ' 3. Whitelist: FechaInicio y FechaFinPrevista pueden ser futuras.
    If m_Nombre = "FECHAINICIO" Or m_Nombre = "FECHAFINPREVISTA" Then
        ValidarFechaCampo = True
        Exit Function
    End If

    ' 4. FechaMaterializado delega en la semántica existente de
    '    Riesgo.ValidarFechaMaterializacionPermitida (Riesgo.cls:6219).
    '    Esto MANTIENE VERDES los 4 átomos Test_RiesgoMaterializacion_*
    '    existentes.
    If m_Nombre = "FECHAMATERIALIZADO" Then
        If DateValue(m_Fecha) > Date Then
            p_Error = "La " & UserLabelFor(m_Nombre) & _
                " no puede ser posterior a la fecha actual"
            ValidarFechaCampo = False
            Exit Function
        End If
        ValidarFechaCampo = True
        Exit Function
    End If

    ' 5. Resto de campos: no se permiten fechas futuras.
    If DateValue(m_Fecha) > Date Then
        p_Error = "La " & UserLabelFor(m_Nombre) & _
            " no puede ser posterior a la fecha actual"
        ValidarFechaCampo = False
        Exit Function
    End If

    ' 6. Fecha pasada o igual a hoy: aceptada.
    ValidarFechaCampo = True
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ValidarFechaCampo: " & Err.Number & " - " & Err.description
    End If
    ValidarFechaCampo = False
End Function

