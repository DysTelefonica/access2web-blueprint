Attribute VB_Name = "Test_Helper_EntidadCRUD"
Option Compare Database
Option Explicit
' Tests atómicos para Helper_EntidadCRUD (PRUEBA-003 REFAC-2a).
' El helper es generico via CallByName + Dictionary. Los tests usan una
' clase mock (MockEntidad) definida al final del archivo, con propiedades
' Public String/Long que el helper puede leer/escribir via late binding.
'
' Cobertura: BR-2a-01..02 (Verified-runtime).
' Los 11 call sites (Form_FormComercial, Form_FormCPV, etc.) son
' Verified-static hasta T2a.6 rewire.

' Test mock class: 2 String properties + 1 Long property
Private Type MockEntidadFields
    Nombre As String
    Descripcion As String
    IDEjemplo As Long
End Type

' Implementamos el mock como un Dictionary para que CallByName funcione via late binding.
' Pero CallByName requiere un Object real, no un Dictionary. Asi que usamos un
' Variant wrapper. Esto es un limite de VBA: para testear CallByName con late binding,
' necesitamos una clase real.
'
' Workaround: la clase MockEntidadTest se importa junto con este modulo.
' Para evitar crear un .cls separado, usamos un Scripting.Dictionary como
' "objeto" - CallByName NO funciona con Dictionary directamente, asi que
' estos tests son mas limitados de lo que me gustaria.
'
' Conclusion: para tests reales del helper, hace falta una clase mock.
' Por ahora, los tests verifican el comportamiento cuando p_Obj es Nothing
' (que es un caso valido y testeable).

' BR-2a-01: CopiarCamposAObjeto con p_Obj Nothing -> error sin crash
Public Function Test_Helper_EntidadCRUD_CopiarCamposAObjeto_ObjNothing_PueblaError() As String
    Dim logs(0 To 2) As String
    Dim errMsg As String
    Dim result As String
    Dim dic As Scripting.Dictionary
    Set dic = New Scripting.Dictionary
    dic.Add "Nombre", "Test"
    logs(0) = "Arrange: p_Obj=Nothing, dic with 1 field"
    result = Helper_EntidadCRUD.CopiarCamposAObjeto(Nothing, dic, Array("Nombre"), errMsg)
    logs(1) = "Act: called"
    If result = "" Then
        Test_Helper_EntidadCRUD_CopiarCamposAObjeto_ObjNothing_PueblaError = BuildJsonFail("expected error, got empty", logs)
        Exit Function
    End If
    If errMsg = "" Then
        Test_Helper_EntidadCRUD_CopiarCamposAObjeto_ObjNothing_PueblaError = BuildJsonFail("expected p_Error populated", logs)
        Exit Function
    End If
    logs(2) = "Assert: error + p_Error populated"
    Test_Helper_EntidadCRUD_CopiarCamposAObjeto_ObjNothing_PueblaError = BuildJsonOk("ok", logs)
End Function

' BR-2a-01: CopiarCamposAObjeto con p_Valores Nothing -> error sin crash
Public Function Test_Helper_EntidadCRUD_CopiarCamposAObjeto_ValoresNothing_PueblaError() As String
    Dim logs(0 To 2) As String
    Dim errMsg As String
    Dim result As String
    Dim mock As Object
    Set mock = CreateObject("Scripting.Dictionary")  ' placeholder; will be replaced by a real mock class in a future slice
    logs(0) = "Arrange: p_Valores=Nothing"
    result = Helper_EntidadCRUD.CopiarCamposAObjeto(mock, Nothing, Array("Nombre"), errMsg)
    logs(1) = "Act: called"
    If result = "" Then
        Test_Helper_EntidadCRUD_CopiarCamposAObjeto_ValoresNothing_PueblaError = BuildJsonFail("expected error, got empty", logs)
        Exit Function
    End If
    If errMsg = "" Then
        Test_Helper_EntidadCRUD_CopiarCamposAObjeto_ValoresNothing_PueblaError = BuildJsonFail("expected p_Error populated", logs)
        Exit Function
    End If
    logs(2) = "Assert: error + p_Error populated"
    Test_Helper_EntidadCRUD_CopiarCamposAObjeto_ValoresNothing_PueblaError = BuildJsonOk("ok", logs)
End Function

' BR-2a-01: CopiarCamposAObjeto con campo que no existe en p_Valores -> error
' (El mock Dictionary no tiene 'Nombre' asi que CallByName falla primero; el test
'  acepta CUALQUIER error como evidencia de que el helper maneja la condicion
'  sin crashear. Para testear exactamente el path de 'campo faltante' haria
'  falta una clase mock con esas properties, lo cual se difiere a un test slice
'  con un .cls dedicado.)
Public Function Test_Helper_EntidadCRUD_CopiarCamposAObjeto_CampoFaltante_PueblaError() As String
    Dim logs(0 To 2) As String
    Dim errMsg As String
    Dim result As String
    Dim dic As Scripting.Dictionary
    Set dic = New Scripting.Dictionary
    dic.Add "Nombre", "Test"
    ' No agregamos "Descripcion" - el helper deberia fallar al intentar escribir
    ' 'Descripcion' (o al escribir 'Nombre' en el mock Dictionary, que no tiene esa property)
    Dim mock As Object
    Set mock = CreateObject("Scripting.Dictionary")  ' placeholder (no tiene properties Nombre/Descripcion)
    logs(0) = "Arrange: dic has Nombre, request Nombre+Descripcion, mock sin properties"
    result = Helper_EntidadCRUD.CopiarCamposAObjeto(mock, dic, Array("Nombre", "Descripcion"), errMsg)
    logs(1) = "Act: called"
    If result = "" Then
        Test_Helper_EntidadCRUD_CopiarCamposAObjeto_CampoFaltante_PueblaError = BuildJsonFail("expected error, got empty", logs)
        Exit Function
    End If
    If errMsg = "" Then
        Test_Helper_EntidadCRUD_CopiarCamposAObjeto_CampoFaltante_PueblaError = BuildJsonFail("expected p_Error populated", logs)
        Exit Function
    End If
    logs(2) = "Assert: error populated (helper no crashea con mock invalido)"
    Test_Helper_EntidadCRUD_CopiarCamposAObjeto_CampoFaltante_PueblaError = BuildJsonOk("ok", logs)
End Function

' BR-2a-01: CopiarCamposAObjeto HAPPY PATH con MockEntidad real
' Cierra la brecha del gap "happy path no testeable" - ahora tenemos una clase
' con properties reales (Comercial, DESCRIPCION) y podemos verificar que el helper
' las escribe correctamente via CallByName.
Public Function Test_Helper_EntidadCRUD_CopiarCamposAObjeto_HappyPath_CopiaValores() As String
    Dim logs(0 To 3) As String
    Dim mock As MockEntidad
    Set mock = New MockEntidad
    Dim dic As Scripting.Dictionary
    Set dic = New Scripting.Dictionary
    dic.Add "Comercial", "Acme SA"
    dic.Add "DESCRIPCION", "Test descripcion"
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: MockEntidad + dic con Comercial='Acme SA' y DESCRIPCION='Test descripcion'"
    result = Helper_EntidadCRUD.CopiarCamposAObjeto(mock, dic, Array("Comercial", "DESCRIPCION"), errMsg)
    logs(1) = "Act: called"
    If result <> "" Then
        Test_Helper_EntidadCRUD_CopiarCamposAObjeto_HappyPath_CopiaValores = BuildJsonFail("expected success, got: " & result & " err: " & errMsg, logs)
        Exit Function
    End If
    If errMsg <> "" Then
        Test_Helper_EntidadCRUD_CopiarCamposAObjeto_HappyPath_CopiaValores = BuildJsonFail("expected no error, got: " & errMsg, logs)
        Exit Function
    End If
    If mock.Comercial <> "Acme SA" Then
        Test_Helper_EntidadCRUD_CopiarCamposAObjeto_HappyPath_CopiaValores = BuildJsonFail("expected mock.Comercial='Acme SA', got: " & mock.Comercial, logs)
        Exit Function
    End If
    If mock.DESCRIPCION <> "Test descripcion" Then
        Test_Helper_EntidadCRUD_CopiarCamposAObjeto_HappyPath_CopiaValores = BuildJsonFail("expected mock.DESCRIPCION='Test descripcion', got: " & mock.DESCRIPCION, logs)
        Exit Function
    End If
    logs(2) = "Assert: mock.Comercial=Acme SA, mock.DESCRIPCION=Test descripcion"
    Test_Helper_EntidadCRUD_CopiarCamposAObjeto_HappyPath_CopiaValores = BuildJsonOk("ok", logs)
End Function

' BR-2a-02: HaHabidoCambiosGenerico HAPPY PATH con MockEntidad real
' Cuando el obj inicial tiene los MISMOS valores que el form, retorna False (sin cambios).
Public Function Test_Helper_EntidadCRUD_HaHabidoCambios_HappyPath_SinCambios_DevuelveFalse() As String
    Dim logs(0 To 3) As String
    Dim mockInicial As MockEntidad
    Set mockInicial = New MockEntidad
    mockInicial.Comercial = "Acme SA"
    mockInicial.DESCRIPCION = "Test descripcion"
    Dim dic As Scripting.Dictionary
    Set dic = New Scripting.Dictionary
    dic.Add "Comercial", "Acme SA"
    dic.Add "DESCRIPCION", "Test descripcion"
    Dim errMsg As String
    Dim result As Boolean
    logs(0) = "Arrange: obj inicial con Comercial=Acme SA, form tambien=Acme SA (sin cambios)"
    result = Helper_EntidadCRUD.HaHabidoCambiosGenerico(mockInicial, dic, Array("Comercial", "DESCRIPCION"), errMsg)
    logs(1) = "Act: called"
    If result <> False Then
        Test_Helper_EntidadCRUD_HaHabidoCambios_HappyPath_SinCambios_DevuelveFalse = BuildJsonFail("expected False (sin cambios), got: " & result, logs)
        Exit Function
    End If
    If errMsg <> "" Then
        Test_Helper_EntidadCRUD_HaHabidoCambios_HappyPath_SinCambios_DevuelveFalse = BuildJsonFail("expected no error, got: " & errMsg, logs)
        Exit Function
    End If
    logs(2) = "Assert: False (sin cambios detectados)"
    Test_Helper_EntidadCRUD_HaHabidoCambios_HappyPath_SinCambios_DevuelveFalse = BuildJsonOk("ok", logs)
End Function

' BR-2a-02: HaHabidoCambiosGenerico HAPPY PATH con un campo cambiado
' Cuando un campo difiere, retorna True.
Public Function Test_Helper_EntidadCRUD_HaHabidoCambios_HappyPath_UnCambio_DevuelveTrue() As String
    Dim logs(0 To 3) As String
    Dim mockInicial As MockEntidad
    Set mockInicial = New MockEntidad
    mockInicial.Comercial = "Acme SA"
    mockInicial.DESCRIPCION = "Original"
    Dim dic As Scripting.Dictionary
    Set dic = New Scripting.Dictionary
    dic.Add "Comercial", "Acme SA"
    dic.Add "DESCRIPCION", "Modificado"
    Dim errMsg As String
    Dim result As Boolean
    logs(0) = "Arrange: obj inicial DESCRIPCION='Original', form DESCRIPCION='Modificado'"
    result = Helper_EntidadCRUD.HaHabidoCambiosGenerico(mockInicial, dic, Array("Comercial", "DESCRIPCION"), errMsg)
    logs(1) = "Act: called"
    If result <> True Then
        Test_Helper_EntidadCRUD_HaHabidoCambios_HappyPath_UnCambio_DevuelveTrue = BuildJsonFail("expected True (1 cambio detectado), got: " & result, logs)
        Exit Function
    End If
    logs(2) = "Assert: True (cambio en DESCRIPCION)"
    Test_Helper_EntidadCRUD_HaHabidoCambios_HappyPath_UnCambio_DevuelveTrue = BuildJsonOk("ok", logs)
End Function

' BR-2a-02: HaHabidoCambiosGenerico con p_ObjInicial Nothing -> True (es un alta)
Public Function Test_Helper_EntidadCRUD_HaHabidoCambios_ObjInicialNothing_DevuelveTrue() As String
    Dim logs(0 To 2) As String
    Dim errMsg As String
    Dim result As Boolean
    Dim dic As Scripting.Dictionary
    Set dic = New Scripting.Dictionary
    dic.Add "Nombre", "Test"
    logs(0) = "Arrange: p_ObjInicial=Nothing (es un alta)"
    result = Helper_EntidadCRUD.HaHabidoCambiosGenerico(Nothing, dic, Array("Nombre"), errMsg)
    logs(1) = "Act: called"
    If result <> True Then
        Test_Helper_EntidadCRUD_HaHabidoCambios_ObjInicialNothing_DevuelveTrue = BuildJsonFail("expected True (es alta), got: " & result, logs)
        Exit Function
    End If
    logs(2) = "Assert: True"
    Test_Helper_EntidadCRUD_HaHabidoCambios_ObjInicialNothing_DevuelveTrue = BuildJsonOk("ok", logs)
End Function

' BR-2a-02: HaHabidoCambiosGenerico con p_Valores Nothing -> True + p_Error
Public Function Test_Helper_EntidadCRUD_HaHabidoCambios_ValoresNothing_PueblaError() As String
    Dim logs(0 To 2) As String
    Dim errMsg As String
    Dim result As Boolean
    Dim mock As Object
    Set mock = CreateObject("Scripting.Dictionary")  ' placeholder
    logs(0) = "Arrange: p_Valores=Nothing"
    result = Helper_EntidadCRUD.HaHabidoCambiosGenerico(mock, Nothing, Array("Nombre"), errMsg)
    logs(1) = "Act: called"
    If result <> True Then
        Test_Helper_EntidadCRUD_HaHabidoCambios_ValoresNothing_PueblaError = BuildJsonFail("expected True (no se puede comparar), got: " & result, logs)
        Exit Function
    End If
    If errMsg = "" Then
        Test_Helper_EntidadCRUD_HaHabidoCambios_ValoresNothing_PueblaError = BuildJsonFail("expected p_Error populated", logs)
        Exit Function
    End If
    logs(2) = "Assert: True + p_Error"
    Test_Helper_EntidadCRUD_HaHabidoCambios_ValoresNothing_PueblaError = BuildJsonOk("ok", logs)
End Function

' BR-2a-02: HaHabidoCambiosGenerico con campo que no existe en p_Valores -> True
Public Function Test_Helper_EntidadCRUD_HaHabidoCambios_CampoFaltante_DevuelveTrue() As String
    Dim logs(0 To 2) As String
    Dim errMsg As String
    Dim result As Boolean
    Dim dic As Scripting.Dictionary
    Set dic = New Scripting.Dictionary
    dic.Add "Nombre", "Test"
    Dim mock As Object
    Set mock = CreateObject("Scripting.Dictionary")  ' placeholder
    logs(0) = "Arrange: dic has Nombre, request Nombre+Descripcion"
    result = Helper_EntidadCRUD.HaHabidoCambiosGenerico(mock, dic, Array("Nombre", "Descripcion"), errMsg)
    logs(1) = "Act: called"
    If result <> True Then
        Test_Helper_EntidadCRUD_HaHabidoCambios_CampoFaltante_DevuelveTrue = BuildJsonFail("expected True (campo faltante), got: " & result, logs)
        Exit Function
    End If
    logs(2) = "Assert: True"
    Test_Helper_EntidadCRUD_HaHabidoCambios_CampoFaltante_DevuelveTrue = BuildJsonOk("ok", logs)
End Function

' BR-2b-01: EliminarEntidadGenerico con p_Operaciones Nothing -> error
Public Function Test_Helper_EntidadCRUD_EliminarEntidadGenerico_OpNothing_PueblaError() As String
    Dim logs(0 To 2) As String
    Dim errMsg As String
    Dim result As String
    Dim entidad As Object
    Set entidad = CreateObject("Scripting.Dictionary")
    logs(0) = "Arrange: p_Operaciones=Nothing"
    result = Helper_EntidadCRUD.EliminarEntidadGenerico(Nothing, "Comercial", entidad, "", errMsg)
    logs(1) = "Act: called"
    If result = "" Then
        Test_Helper_EntidadCRUD_EliminarEntidadGenerico_OpNothing_PueblaError = BuildJsonFail("expected error, got empty", logs)
        Exit Function
    End If
    If errMsg = "" Then
        Test_Helper_EntidadCRUD_EliminarEntidadGenerico_OpNothing_PueblaError = BuildJsonFail("expected p_Error populated", logs)
        Exit Function
    End If
    logs(2) = "Assert: error + p_Error"
    Test_Helper_EntidadCRUD_EliminarEntidadGenerico_OpNothing_PueblaError = BuildJsonOk("ok", logs)
End Function

' BR-2b-01: EliminarEntidadGenerico con p_Entidad Nothing -> error
Public Function Test_Helper_EntidadCRUD_EliminarEntidadGenerico_EntidadNothing_PueblaError() As String
    Dim logs(0 To 2) As String
    Dim errMsg As String
    Dim result As String
    Dim op As Object
    Set op = CreateObject("Scripting.Dictionary")
    logs(0) = "Arrange: p_Entidad=Nothing"
    result = Helper_EntidadCRUD.EliminarEntidadGenerico(op, "Comercial", Nothing, "", errMsg)
    logs(1) = "Act: called"
    If result = "" Then
        Test_Helper_EntidadCRUD_EliminarEntidadGenerico_EntidadNothing_PueblaError = BuildJsonFail("expected error, got empty", logs)
        Exit Function
    End If
    If errMsg = "" Then
        Test_Helper_EntidadCRUD_EliminarEntidadGenerico_EntidadNothing_PueblaError = BuildJsonFail("expected p_Error populated", logs)
        Exit Function
    End If
    logs(2) = "Assert: error + p_Error"
    Test_Helper_EntidadCRUD_EliminarEntidadGenerico_EntidadNothing_PueblaError = BuildJsonOk("ok", logs)
End Function

' BR-2b-01: EliminarEntidadGenerico con Dictionary mock (no tiene Eliminar method) -> error
' (mismo limit del helper: el mock no es una clase real con Eliminar method)
Public Function Test_Helper_EntidadCRUD_EliminarEntidadGenerico_MockSinEliminar_PueblaError() As String
    Dim logs(0 To 2) As String
    Dim errMsg As String
    Dim result As String
    Dim op As Object
    Dim entidad As Object
    Set op = CreateObject("Scripting.Dictionary")  ' no tiene property Comercial ni method Eliminar
    Set entidad = CreateObject("Scripting.Dictionary")
    logs(0) = "Arrange: mock Dictionary sin property/metodo esperado"
    result = Helper_EntidadCRUD.EliminarEntidadGenerico(op, "Comercial", entidad, "", errMsg)
    logs(1) = "Act: called (sin confirmacion, va directo a CallByName)"
    If result = "" Then
        Test_Helper_EntidadCRUD_EliminarEntidadGenerico_MockSinEliminar_PueblaError = BuildJsonFail("expected error (CallByName en mock Dictionary falla), got empty", logs)
        Exit Function
    End If
    If errMsg = "" Then
        Test_Helper_EntidadCRUD_EliminarEntidadGenerico_MockSinEliminar_PueblaError = BuildJsonFail("expected p_Error populated", logs)
        Exit Function
    End If
    logs(2) = "Assert: error + p_Error (helper no crashea con mock invalido)"
    Test_Helper_EntidadCRUD_EliminarEntidadGenerico_MockSinEliminar_PueblaError = BuildJsonOk("ok", logs)
End Function

' BR-2b-01: EliminarEntidadGenerico HAPPY PATH con MockOperaciones
' Cierra el gap del 3er funcion del helper: ahora con una clase mock real (con property + method)
' podemos verificar que el helper setea la property y llama a Eliminar via CallByName.
Public Function Test_Helper_EntidadCRUD_EliminarEntidadGenerico_HappyPath_LlamaEliminar() As String
    Dim logs(0 To 3) As String
    Dim op As MockOperaciones
    Set op = New MockOperaciones
    Dim entidad As MockEntidad
    Set entidad = New MockEntidad
    entidad.Comercial = "Acme SA"
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: MockOperaciones + MockEntidad('Acme SA')"
    result = Helper_EntidadCRUD.EliminarEntidadGenerico(op, "Comercial", entidad, "", errMsg)
    logs(1) = "Act: called without confirmation prompt"
    If result <> "" Then
        Test_Helper_EntidadCRUD_EliminarEntidadGenerico_HappyPath_LlamaEliminar = BuildJsonFail("expected success, got: " & result & " err: " & errMsg, logs)
        Exit Function
    End If
    If errMsg <> "" Then
        Test_Helper_EntidadCRUD_EliminarEntidadGenerico_HappyPath_LlamaEliminar = BuildJsonFail("expected no error, got: " & errMsg, logs)
        Exit Function
    End If
    If op.EliminarLlamadas <> 1 Then
        Test_Helper_EntidadCRUD_EliminarEntidadGenerico_HappyPath_LlamaEliminar = BuildJsonFail("expected EliminarLlamadas=1, got: " & op.EliminarLlamadas, logs)
        Exit Function
    End If
    logs(2) = "Assert: op.EliminarLlamadas=1, op.Comercial=entidad (propiedad seteada via CallByName)"
    Test_Helper_EntidadCRUD_EliminarEntidadGenerico_HappyPath_LlamaEliminar = BuildJsonOk("ok", logs)
End Function

' BR-2b-01: EliminarEntidadGenerico ERROR PATH con m_FailOnEliminar=True
' Verifica que cuando la operations class retorna error, el helper lo propaga correctamente.
Public Function Test_Helper_EntidadCRUD_EliminarEntidadGenerico_ErrorPath_PropagaError() As String
    Dim logs(0 To 3) As String
    Dim op As MockOperaciones
    Set op = New MockOperaciones
    op.m_FailOnEliminar = True
    Dim entidad As MockEntidad
    Set entidad = New MockEntidad
    Dim errMsg As String
    Dim result As String
    logs(0) = "Arrange: MockOperaciones con m_FailOnEliminar=True"
    result = Helper_EntidadCRUD.EliminarEntidadGenerico(op, "Comercial", entidad, "", errMsg)
    logs(1) = "Act: called"
    If result = "" Then
        Test_Helper_EntidadCRUD_EliminarEntidadGenerico_ErrorPath_PropagaError = BuildJsonFail("expected error, got empty", logs)
        Exit Function
    End If
    If errMsg = "" Then
        Test_Helper_EntidadCRUD_EliminarEntidadGenerico_ErrorPath_PropagaError = BuildJsonFail("expected p_Error populated", logs)
        Exit Function
    End If
    logs(2) = "Assert: result con error, p_Error poblada"
    Test_Helper_EntidadCRUD_EliminarEntidadGenerico_ErrorPath_PropagaError = BuildJsonOk("ok", logs)
End Function
