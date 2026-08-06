Attribute VB_Name = "Helper_EntidadCRUD"
Option Compare Database
Option Explicit
' Helper REFAC-2a: CRUD generico para single-record entity forms.
' Stateless, no UI, no popup. Extrae el patron comun de los 11 forms
' (Form_FormComercial, Form_FormCPV, Form_FormEjercito, Form_FormLugarEjecucion,
'  Form_FormOficinaPrograma, Form_FormOrganoContratacion, Form_FormPECAL,
'  Form_FormRAC, Form_FormGradoClasificacion, Form_FormSuministrador, etc.).
'
' Los 11 forms tienen esta estructura duplicada (200 LOC cada uno):
'   - ComandoRegistrar_Click -> copia form fields a m_ObjXxxActivo, llama .Registrar
'   - HaHabidoCambios -> compara form fields a m_ObjXxxAlInicio
'   - EstablecerDatos -> carga m_ObjXxxAlInicio y popula form
'   - Form_Open -> setup
'
' Este helper extrae las 2 funciones mas reusables (Copiar y HaHabidoCambios)
' usando un patron generico:
'   - Input: Scripting.Dictionary de field name -> value (NO Form, NO Me)
'   - Output: aplica al objeto via CallByName (late binding)
'   - Asi es testeable sin instanciar la entity class ni abrir UI
'
' Convenciones:
'   - Sin Form_*, sin Me.*, sin MsgBox, sin DoCmd
'   - Optional ByRef p_Error As String ultimo parametro
'   - Retorno String: "" on success, motivo on failure

' BR-2a-01: Copia los valores de p_Valores (Dictionary) al objeto p_Obj via CallByName.
' p_NombresCampos es Variant array de String con los nombres de propiedades a copiar.
' Retorna "" on success, motivo on failure (campo no encontrado, etc).
Public Function CopiarCamposAObjeto( _
    ByVal p_Obj As Object, _
    ByVal p_Valores As Scripting.Dictionary, _
    ByVal p_NombresCampos As Variant, _
    Optional ByRef p_Error As String) As String
    
    Dim i As Long
    Dim m_Nombre As String
    Dim m_Valor As Variant
    Dim m_ErrNumber As Long
    Dim m_ErrDescription As String
    
    On Error GoTo errores
    
    If p_Obj Is Nothing Then
        p_Error = "CopiarCamposAObjeto: p_Obj is Nothing"
        CopiarCamposAObjeto = p_Error
        Exit Function
    End If
    If p_Valores Is Nothing Then
        p_Error = "CopiarCamposAObjeto: p_Valores is Nothing"
        CopiarCamposAObjeto = p_Error
        Exit Function
    End If
    
    For i = LBound(p_NombresCampos) To UBound(p_NombresCampos)
        m_Nombre = CStr(p_NombresCampos(i))
        If Not p_Valores.Exists(m_Nombre) Then
            p_Error = "CopiarCamposAObjeto: campo '" & m_Nombre & "' no esta en p_Valores"
            CopiarCamposAObjeto = p_Error
            Exit Function
        End If
        m_Valor = p_Valores(m_Nombre)
        ' CallByName es late binding - funciona con cualquier clase
        CallByName p_Obj, m_Nombre, VbLet, m_Valor
    Next i
    
    CopiarCamposAObjeto = ""
    Exit Function
    
errores:
    m_ErrNumber = Err.Number
    m_ErrDescription = Err.Description
    p_Error = "CopiarCamposAObjeto: " & m_ErrDescription & " (campo: " & m_Nombre & ")"
    CopiarCamposAObjeto = p_Error
End Function

' BR-2a-02: Compara los valores de p_Valores (Dictionary) con las propiedades de p_ObjInicial.
' Retorna True si AL MENOS UN campo difiere (hay cambios que guardar), False si todos coinciden.
' Si p_ObjInicial es Nothing, retorna True (es un alta, todo es "cambio").
Public Function HaHabidoCambiosGenerico( _
    ByVal p_ObjInicial As Object, _
    ByVal p_Valores As Scripting.Dictionary, _
    ByVal p_NombresCampos As Variant, _
    Optional ByRef p_Error As String) As Boolean
    
    Dim i As Long
    Dim m_Nombre As String
    Dim m_ValorForm As Variant
    Dim m_ValorObj As Variant
    
    On Error GoTo errores
    
    If p_Valores Is Nothing Then
        p_Error = "HaHabidoCambiosGenerico: p_Valores is Nothing"
        HaHabidoCambiosGenerico = True
        Exit Function
    End If
    
    ' Si no hay inicial (es un alta), todo es "cambio"
    If p_ObjInicial Is Nothing Then
        HaHabidoCambiosGenerico = True
        Exit Function
    End If
    
    For i = LBound(p_NombresCampos) To UBound(p_NombresCampos)
        m_Nombre = CStr(p_NombresCampos(i))
        If Not p_Valores.Exists(m_Nombre) Then
            ' Campo no provisto -> no podemos comparar, asumimos cambio
            HaHabidoCambiosGenerico = True
            Exit Function
        End If
        m_ValorForm = p_Valores(m_Nombre)
        ' CallByName VbGet lee la propiedad via late binding
        On Error Resume Next
        m_ValorObj = CallByName(p_ObjInicial, m_Nombre, VbGet)
        On Error GoTo errores
        If m_ValorForm <> m_ValorObj Then
            HaHabidoCambiosGenerico = True
            Exit Function
        End If
    Next i
    
    HaHabidoCambiosGenerico = False
    Exit Function

errores:
    p_Error = "HaHabidoCambiosGenerico: " & Err.Description & " (campo: " & m_Nombre & ")"
    HaHabidoCambiosGenerico = True
End Function

' BR-2b-01: Elimina una entidad via su operations class. Patron generico via late binding.
' p_Operaciones es la clase XxxOperaciones (ComercialOperaciones, CPVOperaciones, etc.)
' p_NombrePropEntidad es el nombre de la propiedad que recibe la entidad (ej. "Comercial", "CPV", "Ejercito")
' p_Entidad es el objeto a eliminar (Comercial, CPV, etc.)
' p_Confirmacion es el texto del MsgBox (puede ser "" para skip confirm)
' Retorna:
'   ""   - el usuario cancelo (pregunta = No) o la eliminacion tuvo exito
'   "OK" - elimino correctamente
'   motivo - el usuario decidio no eliminar (pregunta = No) o hubo error
' El helper oculta el patron repetido de "set propiedad + llamar Eliminar" de los 10 gestion forms.
Public Function EliminarEntidadGenerico( _
    ByVal p_Operaciones As Object, _
    ByVal p_NombrePropEntidad As String, _
    ByVal p_Entidad As Object, _
    ByVal p_Confirmacion As String, _
    Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    If p_Operaciones Is Nothing Then
        p_Error = "EliminarEntidadGenerico: p_Operaciones is Nothing"
        EliminarEntidadGenerico = p_Error
        Exit Function
    End If
    If p_Entidad Is Nothing Then
        p_Error = "EliminarEntidadGenerico: p_Entidad is Nothing"
        EliminarEntidadGenerico = p_Error
        Exit Function
    End If
    
    ' Confirmacion (opcional)
    If Len(p_Confirmacion) > 0 Then
        If MsgBox(p_Confirmacion, vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar") <> vbYes Then
            ' Usuario cancelo - retorno "" (no error, no eliminacion)
            EliminarEntidadGenerico = ""
            Exit Function
        End If
    End If
    
    ' Set propiedad (ej. m_ComercialOp.Comercial = m_Entidad)
    CallByName p_Operaciones, p_NombrePropEntidad, VbSet, p_Entidad
    ' Llamar Eliminar (todos los XxxOperaciones tienen un metodo Eliminar(p_Error))
    EliminarEntidadGenerico = CallByName(p_Operaciones, "Eliminar", VbMethod, p_Error)
    Exit Function
    
errores:
    p_Error = "EliminarEntidadGenerico: " & Err.Description
    EliminarEntidadGenerico = p_Error
End Function
