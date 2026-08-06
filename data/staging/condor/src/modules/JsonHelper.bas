Attribute VB_Name = "JsonHelper"

Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: JsonHelper
' DESCRIPCIÓN: Utilidades para la gestión de JSON de trazabilidad de cambios.
' DEPENDENCIA: Requiere JsonConverter.bas
' ==========================================================================

Public Sub RegistrarCambio(ByRef jsonDict As Object, ByVal nombreBloque As String, ByVal nombreCampo As String, ByVal valorOriginal As Variant, ByVal valorFinal As Variant)
    ' RESPONSABILIDAD: Actualizar el diccionario de cambios con un nuevo valor.
    ' LÓGICA:
    ' 1. Si el campo NO existe en el JSON -> Se crea. Valor Inicial = valorOriginal.
    ' 2. Si el campo YA existe -> Se preserva Valor Inicial. Se actualiza Valor Final.
    
    ' Validar inputs nulos
    Dim sValOriginal As String
    Dim sValFinal As String
    sValOriginal = Nz(valorOriginal, "")
    sValFinal = Nz(valorFinal, "")
    
    ' Si no hay cambio real respecto al original (y no existía registro previo), ignorar
    ' PERO: Si ya existía registro, puede que esté revirtiendo al valor original.
    ' En ese caso, ¿borramos la entrada o dejamos constancia de que volvió al origen?
    ' Decisión: Si vuelve al valor original, se mantiene el registro para mostrar que hubo "dudas" o ediciones,
    ' salvo que queramos limpiar el JSON. Por ahora, registramos todo cambio.
    
    ' Ignorar si no hay cambio efectivo en este momento (vs lo que se intenta guardar)
    ' Esto se debe filtrar antes de llamar a esta función, pero por seguridad:
    If sValOriginal = sValFinal And Not ExisteCampo(jsonDict, nombreBloque, nombreCampo) Then Exit Sub
    
    Dim bloqueDict As Object
    Dim campoDict As Object
    
    ' 1. Obtener o crear el bloque (ej: "PC_Propuesta")
    If jsonDict.Exists(nombreBloque) Then
        Set bloqueDict = jsonDict(nombreBloque)
    Else
        Set bloqueDict = JsonConverter.ParseJson("{}")
        jsonDict.Add nombreBloque, bloqueDict
    End If
    
    ' 2. Obtener o crear el campo
    If bloqueDict.Exists(nombreCampo) Then
        ' YA EXISTE: Actualizamos solo el valor final y fecha
        Set campoDict = bloqueDict(nombreCampo)
        campoDict("valor_final") = sValFinal
        campoDict("fecha_ult_mod") = Format(Now, "yyyy-mm-dd hh:nn:ss")
        
        If Not m_ObjUsuarioActivo Is Nothing Then
            campoDict("usuario_mod") = m_ObjUsuarioActivo.nombre
        Else
            campoDict("usuario_mod") = "Desconocido"
        End If
    Else
        ' NO EXISTE: Creamos la entrada completa
        ' Solo si hay diferencia real entre original y final (ahora sí filtramos trivialidades)
        If sValOriginal <> sValFinal Then
            Set campoDict = JsonConverter.ParseJson("{}")
            campoDict.Add "valor_inicial", sValOriginal
            campoDict.Add "valor_final", sValFinal
            campoDict.Add "fecha_ult_mod", Format(Now, "yyyy-mm-dd hh:nn:ss")
            
            If Not m_ObjUsuarioActivo Is Nothing Then
                campoDict.Add "usuario_mod", m_ObjUsuarioActivo.nombre
            Else
                campoDict.Add "usuario_mod", "Desconocido"
            End If
            
            bloqueDict.Add nombreCampo, campoDict
        End If
    End If
    
End Sub

Private Function ExisteCampo(ByRef jsonDict As Object, ByVal nombreBloque As String, ByVal nombreCampo As String) As Boolean
    If Not jsonDict.Exists(nombreBloque) Then
        ExisteCampo = False
        Exit Function
    End If
    
    Dim bloque As Object
    Set bloque = jsonDict(nombreBloque)
    ExisteCampo = bloque.Exists(nombreCampo)
End Function

Public Function ObtenerJsonRechazo(ByVal idSolicitud As Long) As Object
    ' Recupera el JSON actual de BBDD y lo devuelve como Diccionario.
    ' Si está vacío o error, devuelve Diccionario vacío.
    ' Nota: Busca el último rechazo (activo o inactivo) para obtener los cambios del técnico
    ' incluso después de que el rechazo haya sido desactivado al avanzar de fase.
    
    Dim rechazo As rechazo
    Set rechazo = RechazoRepositorio.GetUltimoRechazo(idSolicitud)
    
    If rechazo Is Nothing Then
        Set ObtenerJsonRechazo = JsonConverter.ParseJson("{}")
        Exit Function
    End If
    
    If Len(rechazo.CambiosTecnico) = 0 Then
        Set ObtenerJsonRechazo = JsonConverter.ParseJson("{}")
    Else
        On Error Resume Next
        Set ObtenerJsonRechazo = JsonConverter.ParseJson(rechazo.CambiosTecnico)
        If Err.Number <> 0 Then Set ObtenerJsonRechazo = JsonConverter.ParseJson("{}")
        On Error GoTo 0
    End If
End Function

Public Sub GuardarJsonRechazo(ByVal idSolicitud As Long, ByRef jsonDict As Object)
    Dim jsonString As String
    jsonString = JsonConverter.ConvertToJson(jsonDict)
    RechazoRepositorio.ActualizarJSON idSolicitud, jsonString
End Sub


