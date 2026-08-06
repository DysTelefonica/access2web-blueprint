Attribute VB_Name = "modFormInteractionHelper"
Option Compare Database
Option Explicit

' ----------------------------------------------------------------------------
' Helper de interaccion con forms y controles.
'   Skill ref: access-vba-e2e-methodology v1.2-draft #10 (helpers de
'   infraestructura) y #11 (forms delgados). Exports con prefijo
'   FormInteraction_ para evitar colision con Public Sub/Function ya
'   existentes en Funciones Generales.bas y otros legacy.
'
'   Que hace: todas las operaciones que un form .cls necesita para leer,
'   escribir, refrescar y navegar controles y subforms. Centraliza el
'   patron repetido "Forms("X")", "Me.Controls(...)", "DoCmd.OpenForm" y
'   similares que aparecian 26 veces en los 26 forms en scope.
'
'   Que NO hace: logica de negocio, validacion, persistencia, DAO.
'   Eso vive en helpers de feature (mod<Concepto>Helper.bas). Este
'   helper NO toca la BD.
'
'   Naming: FormInteraction_<Verbo> por convencion del proyecto.
' ----------------------------------------------------------------------------

' Re-export de legacy ya existente en Funciones Generales.bas.
' Si la firma legacy cambia, este wrapper se actualiza. NO redefinir.
Public Function FormInteraction_FormularioAbierto(ByVal p_NombreForm As String) As Boolean
    FormInteraction_FormularioAbierto = FormularioAbierto(p_NombreForm)
End Function

Public Sub FormInteraction_CerrarFormulario(ByVal p_NombreForm As String)
    If FormInteraction_FormularioAbierto(p_NombreForm) Then
        DoCmd.Close acForm, p_NombreForm, acSaveNo
    End If
End Sub

' ----------------------------------------------------------------------------
' Acceso tipado a controles
' ----------------------------------------------------------------------------
Public Function FormInteraction_ObtenerTextoControl( _
    ByVal p_Form As Form, _
    ByVal p_NombreControl As String, _
    Optional ByRef p_Error As String _
) As String
    Dim ctl As Control
    Dim valor As Variant

    p_Error = ""
    On Error GoTo errores

    If p_Form Is Nothing Then
        p_Error = "FormInteraction_ObtenerTextoControl: el formulario no esta disponible"
        Exit Function
    End If

    If Len(Trim$(p_NombreControl)) = 0 Then
        p_Error = "FormInteraction_ObtenerTextoControl: el nombre del control esta vacio"
        Exit Function
    End If

    Set ctl = p_Form.Controls(p_NombreControl)
    valor = ctl.value
    FormInteraction_ObtenerTextoControl = Nz(valor, "")
    Exit Function

errores:
    p_Error = "FormInteraction_ObtenerTextoControl: no se pudo leer el control '" & _
        p_NombreControl & "'. Detalle: " & Err.Description
End Function

Public Function FormInteraction_EstablecerValorControl(ByVal p_Form As Form, ByVal p_NombreControl As String, ByVal p_Valor As Variant) As Boolean
    On Error GoTo errores
    p_Form.Controls(p_NombreControl).value = p_Valor
    FormInteraction_EstablecerValorControl = True
    Exit Function
errores:
    FormInteraction_EstablecerValorControl = False
End Function

' ----------------------------------------------------------------------------
' Carga de combos
' ----------------------------------------------------------------------------
Public Function FormInteraction_CargarCombo(ByVal p_Combo As ComboBox, ByVal p_RowSource As String, Optional ByVal p_Default As String = "") As Boolean
    On Error GoTo errores
    p_Combo.RowSource = p_RowSource
    p_Combo.Requery
    If Len(p_Default) > 0 Then
        p_Combo.value = p_Default
    End If
    FormInteraction_CargarCombo = True
    Exit Function
errores:
    FormInteraction_CargarCombo = False
End Function

' ----------------------------------------------------------------------------
' Limpieza de varios controles a la vez
' ----------------------------------------------------------------------------
Public Sub FormInteraction_LimpiarControles(ByVal p_Form As Form, ParamArray p_NombresControles() As Variant)
    Dim i As Long
    For i = LBound(p_NombresControles) To UBound(p_NombresControles)
        Call FormInteraction_EstablecerValorControl(p_Form, CStr(p_NombresControles(i)), Null)
    Next i
End Sub

' ----------------------------------------------------------------------------
' Refresco de subforms
' ----------------------------------------------------------------------------
Public Sub FormInteraction_RefrescarSubform(ByVal p_Form As Form, ByVal p_NombreSubform As String)
    On Error Resume Next
    If Len(p_Form.Controls(p_NombreSubform).SourceObject) > 0 Then
        p_Form.Controls(p_NombreSubform).Requery
    End If
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Mensaje (wrapper de MsgBox para atomos testeables)
' ----------------------------------------------------------------------------
Public Function FormInteraction_Mensaje( _
    ByVal p_Mensaje As String, _
    ByVal p_Estilo As VbMsgBoxStyle, _
    ByVal p_Titulo As String, _
    Optional ByRef p_Respuesta As VbMsgBoxResult = vbOK _
) As Boolean
    ' El parametro ByRef p_Respuesta permite a los atomos testear
    ' la eleccion del usuario sin abrir un MsgBox real.
    p_Respuesta = MsgBox(p_Mensaje, p_Estilo, p_Titulo)
    FormInteraction_Mensaje = True
End Function
