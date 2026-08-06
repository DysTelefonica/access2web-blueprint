Attribute VB_Name = "FormulariosPadreAuxiliares"

Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: FormulariosPadreAuxiliares.bas
' RESPONSABILIDAD: Helpers compartidos entre frmDatosPC, frmDatosCDCA, frmDatosCDCASUB y frmDatosPCSUB.
'                 Evita duplicación de código en métodos casi idénticos.
' VERSIÓN: 1.1 — agregado TipoForm_PCSUB = 4
' ==========================================================================

' NOTE: enumTipoFormulario ahora está en modEnumeradores.bas (shared)

' ==========================================================================
' MÉTODO: SonEntidadesIguales
' DESCRIPCIÓN: Compara dos entidades cualquiera campo a campo usando ColCampos.
'             Funciona con cualquier entidad que implemente getPropiedad/SetPropiedad.
' ==========================================================================
Public Function SonEntidadesIguales(ByVal objA As Object, ByVal objB As Object) As Boolean
    Dim campo As Variant
    Dim valA As Variant, valB As Variant
    
    On Error GoTo Diferentes
    
    For Each campo In objA.ColCampos
        valA = objA.getPropiedad(CStr(campo))
        valB = objB.getPropiedad(CStr(campo))
        
        ' Usamos CStr para normalizar la comparación (ej: Null vs "")
        If CStr(Nz(valA, "")) <> CStr(Nz(valB, "")) Then
            SonEntidadesIguales = False
            Exit Function
        End If
    Next campo
    
    SonEntidadesIguales = True
    Exit Function

Diferentes:
    SonEntidadesIguales = False
End Function

' ==========================================================================
' MÉTODO: ClonarEntidad
' DESCRIPCIÓN: Clona cualquier entidad que implemente ColCampos y getPropiedad.
'              Usa CallByName para invocar el tipo correcto en el llamador.
' PARAMetROS:
'   - origen: Entidad origen a clonar
'   - tipoForm: Tipo de formulario (1=PC, 2=CDCA, 3=CDCASUB)
' RETORNA: Nueva instancia del tipo corresponding
' ==========================================================================
Public Function ClonarEntidad(ByVal origen As Object, ByVal tipoForm As enumTipoFormulario) As Object
    Dim destino As Object
    Dim campo As Variant
    Dim cloneErr As Long
    Dim cloneErrDesc As String
    
    On Error GoTo CloneError
    
    ' Crear la instancia destino según el tipo
    Select Case tipoForm
        Case TipoForm_PC
            Set destino = New DatosPC
        Case TipoForm_CDCA
            Set destino = New DatosCDCA
        Case TipoForm_CDCASUB
            Set destino = New DatosCDCASUB
        Case TipoForm_PCSUB
            Set destino = New DatosPCSUB
        Case Else
            Err.Raise 513, "ClonarEntidad", "Tipo de formulario no reconocido: " & tipoForm
    End Select
    
    ' Copiar campos usando getPropiedad/SetPropiedad (funciona en todas las entidades)
    For Each campo In origen.ColCampos
        CallByName destino, "SetPropiedad", VbMethod, CStr(campo), CallByName(origen, "getPropiedad", VbMethod, CStr(campo))
    Next campo
    
    Set ClonarEntidad = destino
    Exit Function
    
CloneError:
    cloneErr = Err.Number: cloneErrDesc = Err.description
    Dim errObj As New CondorError
    errObj.Create cloneErr, cloneErrDesc, "FormulariosPadreAuxiliares.ClonarEntidad"
    errObj.AddToCallStack "Campo: " & CStr(campo)
    errObj.Raise
End Function

' ==========================================================================
' MÉTODO: GetSubformNameFromTabName
' DESCRIPCIÓN: Traduce el nombre de un botón de navegación (pestaña) al nombre
'              de su subformulario correspondiente.
' PARÁMETROS:
'   - tabName: Nombre del tab/pestaña (ej: "tabGeneral", "tabPropuesta")
'   - tipoForm: Tipo de formulario (1=PC, 2=CDCA, 3=CDCASUB)
' RETORNA: Nombre del subformulario o "" si no se encuentra
' ==========================================================================
Public Function GetSubformNameFromTabName(ByVal tabName As String, ByVal tipoForm As enumTipoFormulario) As String
    Select Case tabName
        Case "tabGeneral"
            Select Case tipoForm
                Case TipoForm_PC:         GetSubformNameFromTabName = "subfrmDatosPC_Generales"
                Case TipoForm_CDCA:       GetSubformNameFromTabName = "subfrmDatosCDCA_Generales"
                Case TipoForm_CDCASUB:    GetSubformNameFromTabName = "subfrmDatosCDCASUB_Generales"
                Case TipoForm_PCSUB:      GetSubformNameFromTabName = "subfrmDatosPCSUB_Generales"
            End Select
            
        Case "tabPropuesta"
            Select Case tipoForm
                Case TipoForm_PC:         GetSubformNameFromTabName = "subfrmDatosPC_Propuesta"
                Case TipoForm_CDCA:        GetSubformNameFromTabName = "subfrmDatosCDCA_Propuesta"
                Case TipoForm_CDCASUB:    GetSubformNameFromTabName = "subfrmDatosCDCASUB_Propuesta"
                Case TipoForm_PCSUB:      GetSubformNameFromTabName = "subfrmDatosPCSUB_Propuesta"
            End Select
            
        Case "tabImpacto"
            Select Case tipoForm
                Case TipoForm_PC:         GetSubformNameFromTabName = "subfrmDatosPC_Impacto"
                Case TipoForm_CDCA:       GetSubformNameFromTabName = "subfrmDatosCDCA_Impacto"
                Case TipoForm_CDCASUB:    GetSubformNameFromTabName = "subfrmDatosCDCASUB_Impacto"
                Case TipoForm_PCSUB:      GetSubformNameFromTabName = "subfrmDatosPCSUB_Impacto"
            End Select
            
        Case "tabAprobacionSuministrador"
            Select Case tipoForm
                Case TipoForm_PC:         GetSubformNameFromTabName = "subfrmDatosPC_AprobacionSuministrador"
                Case TipoForm_CDCA:       GetSubformNameFromTabName = "subfrmDatosCDCA_AprobacionSuministrador"
                Case TipoForm_CDCASUB:    GetSubformNameFromTabName = "subfrmDatosCDCASUB_AprobacionSuministrador"
                Case TipoForm_PCSUB:      GetSubformNameFromTabName = "subfrmDatosPCSUB_AprobacionSuministrador"
            End Select
            
        Case "tabDictamenRAC"
            Select Case tipoForm
                Case TipoForm_PC:         GetSubformNameFromTabName = "subfrmDatosPC_DictamenRAC"
                Case TipoForm_CDCA:       GetSubformNameFromTabName = "subfrmDatosCDCA_DictamenRAC"
                Case TipoForm_CDCASUB:    GetSubformNameFromTabName = "subfrmDatosCDCASUB_DictamenRAC"
                Case TipoForm_PCSUB:      GetSubformNameFromTabName = "subfrmDatosPCSUB_DictamenRAC"
            End Select
            
        Case "tabDecisionFinal"
            Select Case tipoForm
                Case TipoForm_PC:         GetSubformNameFromTabName = "subfrmDatosPC_DecisionFinal"
                Case TipoForm_CDCA:       GetSubformNameFromTabName = "subfrmDatosCDCA_DecisionFinal"
                Case TipoForm_CDCASUB:    GetSubformNameFromTabName = "subfrmDatosCDCASUB_DecisionFinal"
                Case TipoForm_PCSUB:      GetSubformNameFromTabName = "subfrmDatosPCSUB_DecisionFinal"
            End Select
            
        Case Else
            GetSubformNameFromTabName = ""
    End Select
End Function

' ==========================================================================
' MÉTODO: MapearBloquePorNombreSub
' DESCRIPCIÓN: Traduce el nombre del subformulario al enumerado enumBloqueFormulario.
' PARÁMETROS:
'   - nombreSub: Nombre del subformulario (ej: "subfrmDatosPC_Generales")
'   - tipoForm: Tipo de formulario (1=PC, 2=CDCA, 3=CDCASUB)
' RETORNA: Valor enumBloqueFormulario o 0 si no se reconoce
' ==========================================================================
Public Function MapearBloquePorNombreSub(ByVal nombreSub As String, ByVal tipoForm As enumTipoFormulario) As enumBloqueFormulario
    ' Sin On Error — Select Case puro, no puede fallar. El Case Else devuelve 0.
    
    Select Case nombreSub
        ' --- PREFIJO PC ---
        Case "subfrmDatosPC_Generales"
            MapearBloquePorNombreSub = Bloque_Generales
        Case "subfrmDatosPC_Propuesta"
            MapearBloquePorNombreSub = Bloque_Propuesta
        Case "subfrmDatosPC_Impacto"
            MapearBloquePorNombreSub = Bloque_Impacto
        Case "subfrmDatosPC_AprobacionSuministrador"
            MapearBloquePorNombreSub = Bloque_AprobacionSuministrador
        Case "subfrmDatosPC_DictamenRAC"
            MapearBloquePorNombreSub = Bloque_DictamenRAC
        Case "subfrmDatosPC_DecisionFinal"
            MapearBloquePorNombreSub = Bloque_DecisionFinal
            
        ' --- PREFIJO CDCA ---
        Case "subfrmDatosCDCA_Generales"
            MapearBloquePorNombreSub = Bloque_Generales
        Case "subfrmDatosCDCA_Propuesta"
            MapearBloquePorNombreSub = Bloque_Propuesta
        Case "subfrmDatosCDCA_Impacto"
            MapearBloquePorNombreSub = Bloque_Impacto
        Case "subfrmDatosCDCA_AprobacionSuministrador"
            MapearBloquePorNombreSub = Bloque_AprobacionSuministrador
        Case "subfrmDatosCDCA_DictamenRAC"
            MapearBloquePorNombreSub = Bloque_DictamenRAC
        Case "subfrmDatosCDCA_DecisionFinal"
            MapearBloquePorNombreSub = Bloque_DecisionFinal
            
        ' --- PREFIJO CDCASUB ---
        Case "subfrmDatosCDCASUB_Generales"
            MapearBloquePorNombreSub = Bloque_Generales
        Case "subfrmDatosCDCASUB_Propuesta"
            MapearBloquePorNombreSub = Bloque_Propuesta
        Case "subfrmDatosCDCASUB_Impacto"
            MapearBloquePorNombreSub = Bloque_Impacto
        Case "subfrmDatosCDCASUB_AprobacionSuministrador"
            MapearBloquePorNombreSub = Bloque_AprobacionSuministrador
        Case "subfrmDatosCDCASUB_DictamenRAC"
            MapearBloquePorNombreSub = Bloque_DictamenRAC
        Case "subfrmDatosCDCASUB_DecisionFinal"
            MapearBloquePorNombreSub = Bloque_DecisionFinal
            
        ' --- PREFIJO PCSUB ---
        Case "subfrmDatosPCSUB_Generales"
            MapearBloquePorNombreSub = Bloque_Generales
        Case "subfrmDatosPCSUB_Propuesta"
            MapearBloquePorNombreSub = Bloque_Propuesta
        Case "subfrmDatosPCSUB_Impacto"
            MapearBloquePorNombreSub = Bloque_Impacto
        Case "subfrmDatosPCSUB_AprobacionSuministrador"
            MapearBloquePorNombreSub = Bloque_AprobacionSuministrador
        Case "subfrmDatosPCSUB_DictamenRAC"
            MapearBloquePorNombreSub = Bloque_DictamenRAC
        Case "subfrmDatosPCSUB_DecisionFinal"
            MapearBloquePorNombreSub = Bloque_DecisionFinal
            
        Case Else
            MapearBloquePorNombreSub = 0
    End Select
End Function

' ==========================================================================
' MÉTODO: GetNombreAmigableBloque
' DESCRIPCIÓN: Devuelve el nombre visible de un bloque para el usuario,
'              adaptando el mensaje según el rol actual.
' PARÁMETROS:
'   - bloque: Valor enumBloqueFormulario
'   - rol: Rol del usuario actual (usar variable global rolUsuario)
' RETORNA: String con el nombre amigable
' ==========================================================================
Public Function GetNombreAmigableBloque(ByVal bloque As enumBloqueFormulario) As String
    ' Sin On Error — Select Case puro, no puede fallar. El Case Else devuelve default.
    
    Select Case bloque
        Case Bloque_Generales
            GetNombreAmigableBloque = "General"
            
        Case Bloque_Propuesta
            If rolUsuario = rol.Tecnico Then
                GetNombreAmigableBloque = "Detalle y Descripción (1 de 2)"
            Else
                GetNombreAmigableBloque = "Detalle y Descripción"
            End If
            
        Case Bloque_Impacto
            If rolUsuario = rol.Tecnico Then
                GetNombreAmigableBloque = "Motivos (2 de 2)"
            Else
                GetNombreAmigableBloque = "Motivos"
            End If
            
        Case Bloque_AprobacionSuministrador
            GetNombreAmigableBloque = "Aprobación Suministrador"
            
        Case Bloque_DictamenRAC
            GetNombreAmigableBloque = "RAC"
            
        Case Bloque_DecisionFinal
            GetNombreAmigableBloque = "Autoridad Competente"
            
        Case Else
            GetNombreAmigableBloque = "Sección actual"
    End Select
    
    ' Fallback por si viene un valor inesperado
    If GetNombreAmigableBloque = "" Then GetNombreAmigableBloque = "SECCIÓN actual"
End Function

' ==========================================================================
' MÉTODO: HabilitarGuardarEnSubform
' DESCRIPCIÓN: Activa o desactiva los botones de acción del subformulario.
' PARÁMETROS:
'   - frmSub: Referencia al formulario subformulario
'   - habilitar: True para habilitar, False para deshabilitar
' NOTA: Cada formulario puede tener controles ligeramente diferentes
'       (cmdCargarDefaults vs cmdCargarDatosPredeterminados), se ignoran
'       los errores si un control no existe.
' ==========================================================================
Public Sub HabilitarGuardarEnSubform(ByVal frmSub As Access.Form, ByVal habilitar As Boolean)
    On Error Resume Next
    frmSub.Controls("cmdGuardar").Enabled = habilitar
    frmSub.Controls("cmdCargarDefaults").Enabled = habilitar
    frmSub.Controls("cmdCargarDatosPredeterminados").Enabled = habilitar
End Sub
