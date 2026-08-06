Attribute VB_Name = "Helper_ExpedienteEdicion"
Option Compare Database
Option Explicit
' Helper REFAC-3a slice 2: validaciones + comparaciones de edicion de expedientes.
' Stateless, no popup, no UI. Extrae las reglas de negocio de la edicion.
'
' Cobertura: BR-02-01..02, BR-04-01, BR-05-01 (PR-C del coverage matrix).
' Las funciones que tocan DB (SELECT FROM TbExpedientes WHERE Bloqueado=True) se difieren
' porque requieren DAO inyectable + datos sembrados. Esta slice cubre solo las funciones
' puras (CallByName sobre 2 entidades en memoria, o logica condicional).
'
' Tests en src/modules/Test_Helper_ExpedienteEdicion.bas.

' BR-02-01: Compara 2 entidades y retorna la lista de campos que cambiaron.
' p_Anterior y p_Nuevo son 2 objetos con las mismas properties (via CallByName).
' p_NombresCampos es Variant array de String con los nombres de properties a comparar.
' Retorna "" on success + p_CamposSucios poblado con array de Strings;
'         motivo on failure (p_Obj Nothing, etc).
Public Function SoloCamposSucios( _
    ByVal p_Anterior As Object, _
    ByVal p_Nuevo As Object, _
    ByVal p_NombresCampos As Variant, _
    ByRef p_CamposSucios As Variant, _
    Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    If p_Anterior Is Nothing Then
        p_Error = "SoloCamposSucios: p_Anterior is Nothing"
        SoloCamposSucios = p_Error
        Exit Function
    End If
    If p_Nuevo Is Nothing Then
        p_Error = "SoloCamposSucios: p_Nuevo is Nothing"
        SoloCamposSucios = p_Error
        Exit Function
    End If
    
    Dim m_Sucios As String
    m_Sucios = ""
    Dim m_ValorAnt As Variant
    Dim m_ValorNue As Variant
    Dim i As Long
    Dim m_Nombre As String
    
    For i = LBound(p_NombresCampos) To UBound(p_NombresCampos)
        m_Nombre = CStr(p_NombresCampos(i))
        On Error Resume Next
        m_ValorAnt = CallByName(p_Anterior, m_Nombre, VbGet)
        m_ValorNue = CallByName(p_Nuevo, m_Nombre, VbGet)
        On Error GoTo errores
        If CStr(m_ValorAnt) <> CStr(m_ValorNue) Then
            If Len(m_Sucios) > 0 Then m_Sucios = m_Sucios & ","
            m_Sucios = m_Sucios & m_Nombre
        End If
    Next i
    
    p_CamposSucios = m_Sucios
    SoloCamposSucios = ""
    Exit Function
    
errores:
    p_Error = "SoloCamposSucios: " & Err.Description & " (campo: " & m_Nombre & ")"
    SoloCamposSucios = "ERR"
End Function

' BR-04-01: Cambiar tipo de Regular a Lote requiere IDExpedientePadre.
' p_EsAM/EsLote/EsExpediente/EsBasado son los flags actuales; p_TipoNuevo es "AM"/"Lote"/"Expediente"/"Basado".
' Retorna "" if OK; motivo if cambio de tipo requiere padre y no lo tiene.
Public Function CambiarTipoRequierePadre( _
    ByVal p_EsAM As String, _
    ByVal p_EsLote As String, _
    ByVal p_EsExpediente As String, _
    ByVal p_EsBasado As String, _
    ByVal p_TipoNuevo As String, _
    ByVal p_IDExpedientePadre As String, _
    ByRef p_Motivo As String, _
    Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    ' Determinar el tipo actual (asumimos exactamente uno es "Sí")
    Dim m_TipoActual As String
    If p_EsAM = "Sí" Then m_TipoActual = "AM"
    If p_EsLote = "Sí" Then m_TipoActual = "Lote"
    If p_EsExpediente = "Sí" Then m_TipoActual = "Expediente"
    If p_EsBasado = "Sí" Then m_TipoActual = "Basado"
    
    ' Si el tipo no cambia, OK (no es un "cambio" de tipo)
    If m_TipoActual = p_TipoNuevo Then
        CambiarTipoRequierePadre = "OK"
        Exit Function
    End If
    
    ' Reglas de cambio
    If p_TipoNuevo = "Lote" Then
        ' Cualquier cosa -> Lote requiere padre
        If Len(p_IDExpedientePadre) = 0 Then
            p_Motivo = "Cambiar a Lote requiere un IDExpedientePadre (Acuerdo Marco o Expediente padre)"
            CambiarTipoRequierePadre = "NO"
            Exit Function
        End If
    ElseIf p_TipoNuevo = "Basado" Then
        ' Cualquier cosa -> Basado requiere padre
        If Len(p_IDExpedientePadre) = 0 Then
            p_Motivo = "Cambiar a Basado requiere un IDExpedientePadre"
            CambiarTipoRequierePadre = "NO"
            Exit Function
        End If
    ElseIf p_TipoNuevo = "AM" Or p_TipoNuevo = "Expediente" Then
        ' AM/Expediente no requieren padre
        ' (cualquiera -> AM: OK; cualquiera -> Expediente: OK)
    End If
    
    CambiarTipoRequierePadre = "OK"
    Exit Function
    
errores:
    p_Error = "CambiarTipoRequierePadre: " & Err.Description
    CambiarTipoRequierePadre = "ERR"
End Function

' BR-05-01: Deriva el estado del expediente desde FechaInicio, FechaFin, y EstadoManual.
' Estados posibles: "No iniciado" (sin FechaInicio), "En curso" (entre fechas), "Finalizado" (FechaFin pasada).
Public Function EstadoCalculadoTexto( _
    ByVal p_FechaInicio As Variant, _
    ByVal p_FechaFin As Variant, _
    ByVal p_EstadoManual As String, _
    ByRef p_Estado As String, _
    Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    ' Si el usuario setea estado manual, ese gana (override)
    If Len(p_EstadoManual) > 0 Then
        p_Estado = p_EstadoManual
        EstadoCalculadoTexto = "OK"
        Exit Function
    End If
    
    ' Si no hay FechaInicio, el estado es "No iniciado"
    ' Usamos `& ""` para convertir Variant a String de forma segura (Null/Empty -> "")
    If IsNull(p_FechaInicio) Or Len(p_FechaInicio & "") = 0 Then
        p_Estado = "No iniciado"
        EstadoCalculadoTexto = "OK"
        Exit Function
    End If

    ' Si no hay FechaFin, esta "En curso" (iniciado pero no finalizado)
    If IsNull(p_FechaFin) Or Len(p_FechaFin & "") = 0 Then
        p_Estado = "En curso"
        EstadoCalculadoTexto = "OK"
        Exit Function
    End If
    
    ' Ambos fechas presentes: comparar con la fecha actual
    If CDate(p_FechaFin) < Date Then
        p_Estado = "Finalizado"
    ElseIf CDate(p_FechaInicio) > Date Then
        p_Estado = "No iniciado"
    Else
        p_Estado = "En curso"
    End If
    
    EstadoCalculadoTexto = "OK"
    Exit Function
    
errores:
    p_Error = "EstadoCalculadoTexto: " & Err.Description
    EstadoCalculadoTexto = "ERR"
End Function
