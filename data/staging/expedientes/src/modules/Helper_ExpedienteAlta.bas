Attribute VB_Name = "Helper_ExpedienteAlta"
Option Compare Database
Option Explicit
' Helper REFAC-3a: validaciones + logica del alta de expedientes.
' Stateless, no popup, no UI. Extrae las reglas de negocio del alta
' (que viven en ExpedienteOperaciones.MotivoNoOKAlta + RegistrarAlta) para que
' sean testables de forma aislada.
'
' Cobertura: BR-01-01..05 (PR-C del coverage matrix)
' Pendiente: BR-01-01 happy path (requiere ExpedienteDTO completo, multi-step)
'           BR-01-05 rollback on FK violation (requiere setup de FK invalida + assert rollback)
'
' Tests en src/modules/Test_Helper_ExpedienteAlta.bas.
'
' Convenciones:
'   - Sin Form_*, sin Me.*, sin popup, sin MsgBox
'   - Optional ByRef p_Error As String ultimo parametro
'   - Retorno String: "" on success, motivo on validation fail

' BR-01-02: CodExp es obligatorio (salvo para Lotes que se generan automaticamente)
' Retorna "" si OK; motivo si CodExp vacio y NO es Lote.
Public Function ValidarCodExpObligatorio( _
    ByVal p_Expediente As Object, _
    ByRef p_Motivo As String, _
    Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    If p_Expediente Is Nothing Then
        p_Error = "ValidarCodExpObligatorio: p_Expediente is Nothing"
        ValidarCodExpObligatorio = p_Error
        Exit Function
    End If
    
    Dim m_CodExp As String
    Dim m_EsLote As String
    m_CodExp = CallByName(p_Expediente, "CodExp", VbGet)
    m_EsLote = CallByName(p_Expediente, "EsLote", VbGet)
    
    If Len(m_CodExp) = 0 And m_EsLote <> "Sí" Then
        p_Motivo = "El Codigo de Expediente es obligatorio (salvo para Lotes)"
        ValidarCodExpObligatorio = "NO"
        Exit Function
    End If
    
    ValidarCodExpObligatorio = "OK"
    Exit Function
    
errores:
    p_Error = "ValidarCodExpObligatorio: " & Err.Description
    ValidarCodExpObligatorio = "ERR"
End Function

' BR-01-03: Lote requiere IDExpedientePadre
Public Function ValidarLoteRequierePadre( _
    ByVal p_Expediente As Object, _
    ByRef p_Motivo As String, _
    Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    If p_Expediente Is Nothing Then
        p_Error = "ValidarLoteRequierePadre: p_Expediente is Nothing"
        ValidarLoteRequierePadre = p_Error
        Exit Function
    End If
    
    Dim m_EsLote As String
    Dim m_IDExpedientePadre As String
    m_EsLote = CallByName(p_Expediente, "EsLote", VbGet)
    m_IDExpedientePadre = CallByName(p_Expediente, "IDExpedientePadre", VbGet)
    
    If m_EsLote = "Sí" And Len(m_IDExpedientePadre) = 0 Then
        p_Motivo = "Un Lote requiere un IDExpedientePadre (Acuerdo Marco o Expediente padre)"
        ValidarLoteRequierePadre = "NO"
        Exit Function
    End If
    
    ValidarLoteRequierePadre = "OK"
    Exit Function
    
errores:
    p_Error = "ValidarLoteRequierePadre: " & Err.Description
    ValidarLoteRequierePadre = "ERR"
End Function

' BR-01-04: Importe no puede ser negativo
Public Function ValidarImporteNoNegativo( _
    ByVal p_Expediente As Object, _
    ByRef p_Motivo As String, _
    Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    If p_Expediente Is Nothing Then
        p_Error = "ValidarImporteNoNegativo: p_Expediente is Nothing"
        ValidarImporteNoNegativo = p_Error
        Exit Function
    End If
    
    Dim m_Importe As Variant
    m_Importe = CallByName(p_Expediente, "ImporteLicitacion", VbGet)
    
    If IsNumeric(m_Importe) Then
        If CDbl(m_Importe) < 0 Then
            p_Motivo = "El Importe de Licitacion no puede ser negativo (recibido: " & m_Importe & ")"
            ValidarImporteNoNegativo = "NO"
            Exit Function
        End If
    End If
    
    ValidarImporteNoNegativo = "OK"
    Exit Function
    
errores:
    p_Error = "ValidarImporteNoNegativo: " & Err.Description
    ValidarImporteNoNegativo = "ERR"
End Function

' BR-01-02..04: Wrapper de las 3 validaciones. Retorna el primer motivo.
' Si todas pasan, retorna "OK".
Public Function ValidarAlta( _
    ByVal p_Expediente As Object, _
    ByRef p_Motivo As String, _
    Optional ByRef p_Error As String) As String
    
    Dim m_Resultado As String
    Dim m_MotivoLocal As String
    
    m_Resultado = ValidarCodExpObligatorio(p_Expediente, m_MotivoLocal, p_Error)
    If m_Resultado <> "OK" Then
        p_Motivo = m_MotivoLocal
        ValidarAlta = m_Resultado
        Exit Function
    End If
    
    m_Resultado = ValidarLoteRequierePadre(p_Expediente, m_MotivoLocal, p_Error)
    If m_Resultado <> "OK" Then
        p_Motivo = m_MotivoLocal
        ValidarAlta = m_Resultado
        Exit Function
    End If
    
    m_Resultado = ValidarImporteNoNegativo(p_Expediente, m_MotivoLocal, p_Error)
    If m_Resultado <> "OK" Then
        p_Motivo = m_MotivoLocal
        ValidarAlta = m_Resultado
        Exit Function
    End If
    
    ValidarAlta = "OK"
End Function
