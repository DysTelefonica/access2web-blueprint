Attribute VB_Name = "modNCHelper"
Option Compare Database
Option Explicit

' =============================================================================
' modNCHelper.bas
'
' Helper module para NoConformidades (NC) en el contexto de riesgos materializados.
' Project: gestion_riesgos
' Branch: feat/metodologia-e2e-riesgos-2026-06-19
' Refactor: 2026-06-19
'
' Helpers públicos (2):
'   #1 RegistrarNoConformidad — crea NC y vincula a materialización
'   #2 NC_HaHabidoCambios — compara 14 campos de NC para detectar cambios
'
' NOTA: NC_HaHabidoCambios retorna Boolean (NO String JSON). Excepción al
' contrato canónico documentada en átomos (Test_NCHelper.bas).
' =============================================================================

' --- Referencias usadas en este módulo ---
' Constructor.bas        — getRiesgoMaterializado(), getNC()
' NC.cls                — clase de No Conformidad
' RiesgoMaterializacion — VincularNC()
' Test_Helper.bas       — BuildJsonOk(), BuildJsonFail() (solo para el caller)

' =============================================================================
' Helper #1: RegistrarNoConformidad
'
' Orchestrates creation of a new NC linked to a materialized risk.
'
' Comportamiento extraído de Form_FormRiesgoNC.ComandoRegistrar_Click (9-68):
'   1. Resolver la materialización por ID
'   2. Validar que existe y está materializada
'   3. Verificar que no tiene NC vinculada (ParaNC <> 'Sí')
'   4. Pedir confirmación vía MsgBox (bypass con p_PromptResult)
'   5. Crear nuevo NC, setear campos desde el form
'   6. NC.Registrar() — persiste en TbNoConformidades
'   7. RiesgoMaterializacion.VincularNC() — vincula en TbRiesgosMaterializaciones
'   8. p_NCRegistrado = m_NC (para que el test verifique)
'
' Parámetros:
'   p_IDRiesgoMaterializado  — ID de la materialización a vincular
'   p_NCRegistrado          — (out) objeto NC persisted (para tests y caller)
'   db                       — DAO.Database explícito (para tests con sandbox)
'   p_PromptResult           — si <> 0, usa este valor en lugar de MsgBox
'
' Retorna: JSON string (BuildJsonOk / BuildJsonFail)
' =============================================================================
Public Function RegistrarNoConformidad( _
    ByVal p_IDRiesgoMaterializado As String, _
    Optional ByRef p_NCRegistrado As nc, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_PromptResult As Long) As String

    Dim m_RiesgoMat As RiesgoMaterializacion
    Dim m_NC As nc
    Dim m_Proyecto As Proyecto
    Dim m_PromptResult As Long
    Dim m_Salida As String
    Dim m_Error As String

    RegistrarNoConformidad = ""
    Set p_NCRegistrado = Nothing

    On Error GoTo errores

    ' --- 1. Resolver la materialización ---
    m_Error = ""
    Set m_RiesgoMat = Constructor.getRiesgoMaterializado(Trim$(p_IDRiesgoMaterializado), m_Error)
    If m_Error <> "" Then
        m_Salida = BuildJsonFail("Materializacion no encontrada: " & m_Error, GetZeroLogs())
        RegistrarNoConformidad = m_Salida
        Exit Function
    End If
    If m_RiesgoMat Is Nothing Then
        m_Salida = BuildJsonFail("Materializacion no encontrada: ID " & p_IDRiesgoMaterializado & " no existe", GetZeroLogs())
        RegistrarNoConformidad = m_Salida
        Exit Function
    End If

    ' --- 2. Validar estado: debe estar materializada ---
    If m_RiesgoMat.EsMaterializacionCalcuado <> EnumSiNo.Sí Then
        m_Salida = BuildJsonFail("La materializacion no esta activa para registrar una NC", GetZeroLogs())
        RegistrarNoConformidad = m_Salida
        Exit Function
    End If

    ' --- 3. Verificar que no tiene NC ya vinculada ---
    ' Si ParaNC='Sí' (calculado desde IDNC<>''), ya hay una NC vinculada.
    ' Intentar crear otra NC duplicada no es comportamiento válido.
    If m_RiesgoMat.ParaNCCalculado = EnumSiNo.Sí Then
        m_Salida = BuildJsonFail("La materializacion ya tiene una NC vinculada", GetZeroLogs())
        RegistrarNoConformidad = m_Salida
        Exit Function
    End If

    ' --- 4. Confirmación del usuario ---
    If p_PromptResult <> 0 Then
        m_PromptResult = p_PromptResult
    Else
        m_PromptResult = MsgBox( _
            "Se va a registrar una No Conformidad para el riesgo materializado seleccionado." & vbCrLf & vbCrLf & _
            "¿Desea continuar?", _
            vbYesNo + vbQuestion, _
            "Registrar No Conformidad")
    End If

    If m_PromptResult <> vbYes Then
        m_Salida = BuildJsonFail("Operacion cancelada por el usuario", GetZeroLogs())
        RegistrarNoConformidad = m_Salida
        Exit Function
    End If

    ' --- 5. Obtener el proyecto para setear Juridica ---
    Set m_Proyecto = m_RiesgoMat.Proyecto
    m_Error = m_RiesgoMat.Error
    If m_Error <> "" Then
        m_Salida = BuildJsonFail("Error al obtener el proyecto: " & m_Error, GetZeroLogs())
        RegistrarNoConformidad = m_Salida
        Exit Function
    End If
    If m_Proyecto Is Nothing Then
        m_Salida = BuildJsonFail("No se encontro el proyecto asociado a la materializacion", GetZeroLogs())
        RegistrarNoConformidad = m_Salida
        Exit Function
    End If

    ' --- 6. Crear nuevo NC y setear campos ---
    Set m_NC = New nc
    m_NC.Juridica = m_Proyecto.Juridica
    m_NC.EsNoConformidad = True

    ' Los campos de datos (Descripcion, Causa, ACR, Expediente, Tipo, etc.)
    ' se esperan pre-setados en la NC que entrega el form.
    ' Aqui solo se setea lo que el form hace en RellenarDatosDeColeDeFormulario
    ' que requiere el proyecto: Juridica + EsNoConformidad=True.
    ' Los demas campos los setea el form antes de llamar al helper
    ' o bien se esperan vacios (NC.Registrar valida los obligatorios).

    ' --- 7. Registrar la NC (persiste en TbNoConformidades) ---
    ' p_ObjNCAlInicio = Nothing indica que es NUEVA NC (no update)
    m_Error = ""
    m_Salida = m_NC.Registrar(Nothing, m_Error)
    If m_Error <> "" Then
        m_Salida = BuildJsonFail("Error al registrar la NC: " & m_Error, GetZeroLogs())
        RegistrarNoConformidad = m_Salida
        Exit Function
    End If

    ' --- 8. Vincular la NC a la materializacion ---
    ' El return de VincularNC sobrescribiria m_Error (ByRef), por eso
    ' usamos el pattern: ignorar retorno, chequear Error property post-llamada.
    m_Salida = m_RiesgoMat.VincularNC(m_NC, m_Error)
    If m_RiesgoMat.Error <> "" Then
        ' La NC ya esta persistida pero la vinculacion fallo.
        ' No lanzamos error 1000 para no perder el ID ya generado.
        m_Salida = BuildJsonFail("NC creada pero error al vincular: " & m_RiesgoMat.Error, GetZeroLogs())
        RegistrarNoConformidad = m_Salida
        Exit Function
    End If

    ' --- 9. Devolver la NC creada ---
    Set p_NCRegistrado = m_NC

    ' JSON payload con los datos de la NC creada
    Dim jsonResult As Object
    Set jsonResult = New Scripting.Dictionary
    jsonResult("nc_id") = m_NC.IDNoConformidad
    jsonResult("codigo") = m_NC.CodigoNoConformidad
    Dim jsonPayload As String
    jsonPayload = JsonConverter.ConvertToJson(jsonResult, 2)

    RegistrarNoConformidad = BuildJsonOk(jsonPayload, GetZeroLogs())
    Exit Function

errores:
    If Err.Number <> 1000 Then
        m_Salida = BuildJsonFail("Error inesperado en RegistrarNoConformidad: " & Err.Number & " - " & Err.description, GetZeroLogs())
    Else
        m_Salida = BuildJsonFail("Error registrado: " & Err.description, GetZeroLogs())
    End If
    RegistrarNoConformidad = m_Salida
End Function

' =============================================================================
' Helper #2: NC_HaHabidoCambios
'
' Detecta cambios en una NC comparando 14 campos específicos entre el snapshot
' en memoria (p_NCInicial) y el estado actual en la base de datos.
'
' Comportamiento extraído de Form_FormRiesgoNC.HaHabidoCambios (342-415).
' Son 14 campos: Descripcion, Causa, ACR, Expediente, Tipo, Proyecto, Vehiculo,
' EntidadResponsable, FechaApertura, FPREVCIERRE, ResponsableTelefonica, Notas,
' RequiereControlEficacia, FechaPrevistaControlEficacia.
'
' NOTA: Retorna Boolean (NO JSON). Excepción al contrato canónico documentada
' en Test_NCHelper.bas — el átomo lo testea con "If result Then".
'
' Parámetros:
'   p_NCInicial        — snapshot de la NC al inicio (desde BD o form open)
'   p_IDNoConformidad  — ID de la NC actual en BD para recargar y comparar
'
' Retorna: Boolean — True si al menos 1 campo cambió, False si todos iguales
' =============================================================================
Public Function NC_HaHabidoCambios( _
    ByVal p_NCInicial As nc, _
    ByVal p_IDNoConformidad As String) As Boolean

    Dim m_NCActual As nc
    Dim m_Error As String

    NC_HaHabidoCambios = False
    m_Error = ""

    On Error GoTo errores

    ' Si no hay snapshot, se asume que hay cambios (nueva NC)
    If p_NCInicial Is Nothing Then
        NC_HaHabidoCambios = True
        Exit Function
    End If

    ' Cargar el estado actual de la NC desde la BD
    Set m_NCActual = Constructor.getNC(Trim$(p_IDNoConformidad), m_Error)
    If m_Error <> "" Then
        ' No se pudo cargar la NC desde BD: no hay cambios detectables
        NC_HaHabidoCambios = False
        Exit Function
    End If
    If m_NCActual Is Nothing Then
        ' La NC ya no existe en BD
        NC_HaHabidoCambios = False
        Exit Function
    End If

    ' Comparar los 14 campos editables de la NC
    ' (el orden y la lógica deben coincidir con Form_FormRiesgoNC.HaHabidoCambios)

    If p_NCInicial.Descripcion <> Nz(m_NCActual.Descripcion, "") Then
        NC_HaHabidoCambios = True
        Exit Function
    End If

    If p_NCInicial.Causa <> Nz(m_NCActual.Causa, "") Then
        NC_HaHabidoCambios = True
        Exit Function
    End If

    If p_NCInicial.ACR <> Nz(m_NCActual.ACR, "") Then
        NC_HaHabidoCambios = True
        Exit Function
    End If

    If p_NCInicial.Expediente <> Nz(m_NCActual.Expediente, "") Then
        NC_HaHabidoCambios = True
        Exit Function
    End If

    If p_NCInicial.Tipo <> Nz(m_NCActual.Tipo, "") Then
        NC_HaHabidoCambios = True
        Exit Function
    End If

    If p_NCInicial.Proyecto <> Nz(m_NCActual.Proyecto, "") Then
        NC_HaHabidoCambios = True
        Exit Function
    End If

    If p_NCInicial.Vehiculo <> Nz(m_NCActual.Vehiculo, "") Then
        NC_HaHabidoCambios = True
        Exit Function
    End If

    If p_NCInicial.EntidadResponsable <> Nz(m_NCActual.EntidadResponsable, "") Then
        NC_HaHabidoCambios = True
        Exit Function
    End If

    If p_NCInicial.FechaApertura <> Nz(m_NCActual.FechaApertura, "") Then
        NC_HaHabidoCambios = True
        Exit Function
    End If

    If p_NCInicial.FPREVCIERRE <> Nz(m_NCActual.FPREVCIERRE, "") Then
        NC_HaHabidoCambios = True
        Exit Function
    End If

    If p_NCInicial.ResponsableTelefonica <> Nz(m_NCActual.ResponsableTelefonica, "") Then
        NC_HaHabidoCambios = True
        Exit Function
    End If

    If p_NCInicial.Notas <> Nz(m_NCActual.Notas, "") Then
        NC_HaHabidoCambios = True
        Exit Function
    End If

    If p_NCInicial.RequiereControlEficacia <> Nz(m_NCActual.RequiereControlEficacia, "") Then
        NC_HaHabidoCambios = True
        Exit Function
    End If

    If p_NCInicial.FechaPrevistaControlEficacia <> Nz(m_NCActual.FechaPrevistaControlEficacia, "") Then
        NC_HaHabidoCambios = True
        Exit Function
    End If

    ' Todos los campos iguais > no hubo cambios
    NC_HaHabidoCambios = False
    Exit Function

errores:
    ' Cualquier error no recuperable: no se pueden detectar cambios
    NC_HaHabidoCambios = False
End Function

' =============================================================================
' Helper privado: GetZeroLogs
' Devuelve un array de strings vacio para usar con BuildJsonOk/Fail
' =============================================================================
Private Function GetZeroLogs() As String()
    Dim logs(0 To 0) As String
    logs(0) = ""
    GetZeroLogs = logs
End Function

