
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: Form_subfrmDatosCDCA_AprobacionSuministrador
' CORRECCIÓN: Respetar flags booleanos 'Aplica' y no inferir por contenido de texto.
' ==========================================================================

Private m_ViewModel As DatosCDCAViewModel

' Snapshot para restaurar datos
Private m_original_IngNombre As String
Private m_original_CalNombre As String
Private m_original_ProdNombre As String
Private m_original_DisNombre As String
Private m_original_RepNombre As String

Public Sub PoblarControles(ByRef vm As DatosCDCAViewModel)
    On Error GoTo Errores
    Set m_ViewModel = vm
    If vm Is Nothing Then Exit Sub

    ' 1. Firmas Obligatorias
    Me.firmaAprobacionRespIngenieriaNombre = Nz(vm.vm_firmaAprobacionRespIngenieriaNombre, "")
    Me.firmaAprobacionRespCalidadNombre = Nz(vm.vm_firmaAprobacionRespCalidadNombre, "")
    
    ' 2. Firmas Opcionales (Datos)
    Me.firmaAprobacionRespProduccionNombre = Nz(vm.vm_firmaAprobacionRespProduccionNombre, "")
    Me.firmaAprobacionRespDisenioNombre = Nz(vm.vm_firmaAprobacionRespDisenioNombre, "")
    Me.firmaAprobacionRepresentanteSumNombre = Nz(vm.vm_firmaAprobacionRepresentanteSumNombre, "")
    
    ' Guardar valores originales para dirty check
    m_original_IngNombre = Nz(vm.vm_firmaAprobacionRespIngenieriaNombre, "")
    m_original_CalNombre = Nz(vm.vm_firmaAprobacionRespCalidadNombre, "")
    m_original_ProdNombre = Nz(vm.vm_firmaAprobacionRespProduccionNombre, "")
    m_original_DisNombre = Nz(vm.vm_firmaAprobacionRespDisenioNombre, "")
    m_original_RepNombre = Nz(vm.vm_firmaAprobacionRepresentanteSumNombre, "")
    
    ' 3. Ajustar UI (Toggles y Habilitación) CORREGIDO
    ' Usamos los booleanos explícitos del VM, no inferimos por texto.
    
    ' Producción
    If vm.vm_aplicaRespProduccion Then Me.grpRespProduccion.value = 1 Else Me.grpRespProduccion.value = 2
    Call ActualizarUI_Produccion
    
    ' Diseño
    If vm.vm_aplicaRespDisenio Then Me.grpRespDisenio.value = 1 Else Me.grpRespDisenio.value = 2
    Call ActualizarUI_Disenio
    
    ' Representante
    If vm.vm_aplicaRepresentanteSum Then Me.grpRepresentanteSum.value = 1 Else Me.grpRepresentanteSum.value = 2
    Call ActualizarUI_Representante
    
    ' 4. Visibilidad y Habilitación forzadas
    Dim puedeEditar As Boolean
    puedeEditar = vm.PuedeGuardar_AprobacionSuministrador
    
    Me.cmdCargarDatosPredeterminados.Visible = True
    Me.cmdCargarDatosPredeterminados.Enabled = puedeEditar
    Me.cmdGuardar.Enabled = puedeEditar
    Call QuitarFocoControl(Me, "cmdRestaurar")
    Me.cmdRestaurar.Visible = False
    
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("subfrmDatosCDCA_AprobacionSuministrador.PoblarControles", Me)
End Sub

Public Sub RecogerValores(ByRef vm As DatosCDCAViewModel)
    On Error GoTo Errores
    If vm Is Nothing Then Exit Sub
    
    vm.vm_firmaAprobacionRespIngenieriaNombre = Nz(Me.firmaAprobacionRespIngenieriaNombre.value, "")
    vm.vm_firmaAprobacionRespCalidadNombre = Nz(Me.firmaAprobacionRespCalidadNombre.value, "")
    
    vm.vm_firmaAprobacionRespProduccionNombre = Nz(Me.firmaAprobacionRespProduccionNombre.value, "")
    vm.vm_firmaAprobacionRespDisenioNombre = Nz(Me.firmaAprobacionRespDisenioNombre.value, "")
    vm.vm_firmaAprobacionRepresentanteSumNombre = Nz(Me.firmaAprobacionRepresentanteSumNombre.value, "")

    ' Guardamos el estado de los toggles explícitamente
    vm.vm_aplicaRespProduccion = (Nz(Me.grpRespProduccion.value, 2) = 1)
    vm.vm_aplicaRespDisenio = (Nz(Me.grpRespDisenio.value, 2) = 1)
    vm.vm_aplicaRepresentanteSum = (Nz(Me.grpRepresentanteSum.value, 2) = 1)

LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("RecogerValores", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdGuardar_Click()
    On Error GoTo Errores
    Me.Parent.GuardarDesdeSubform Me.name
    Call QuitarFocoControl(Me, "cmdRestaurar")
    Me.cmdRestaurar.Visible = False
    
    ' Reactivar carga de defaults si sigue siendo editable
    'Me.cmdCargarDatosPredeterminados.Enabled = m_ViewModel.PuedeGuardar_AprobacionSuministrador
    
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdGuardar_Click", Me)
End Sub

Private Sub cmdCargarDatosPredeterminados_Click()
    Dim cdcaServ As New DatosCDCAServicio
    Dim cdcaInicial As DatosCDCA
    On Error GoTo Errores
    If m_ViewModel Is Nothing Then Exit Sub
    
    Me.Parent.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    
    ' 1. Snapshot
    m_original_IngNombre = Nz(Me.firmaAprobacionRespIngenieriaNombre.value, "")
    m_original_CalNombre = Nz(Me.firmaAprobacionRespCalidadNombre.value, "")
    
    ' 2. Obtener defaults
    Set cdcaInicial = cdcaServ.GetDatosAprobacionIniciales(m_ViewModel.Solicitud.idSolicitud)
    
    ' 3. Aplicar defaults
    Me.firmaAprobacionRespIngenieriaNombre = cdcaInicial.firmaAprobacionRespIngenieriaNombre
    Me.firmaAprobacionRespCalidadNombre = cdcaInicial.firmaAprobacionRespCalidadNombre
    Me.cmdRestaurar.Visible = True
    Me.cmdRestaurar.Enabled = True
    ' 4. Gestión botones
    Call QuitarFocoControl(Me, "cmdCargarDatosPredeterminados")
    Me.cmdCargarDatosPredeterminados.Enabled = False
    
    
    
LimpiarYSalir:
    On Error Resume Next
    Me.Parent.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Set cdcaServ = Nothing
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdCargarDatosPredeterminados_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdRestaurar_Click()
    On Error GoTo Errores
    
    ' Restaurar desde Snapshot
    Me.firmaAprobacionRespIngenieriaNombre = m_original_IngNombre
    Me.firmaAprobacionRespCalidadNombre = m_original_CalNombre
    
    ' Restaurar botones
    Me.cmdCargarDatosPredeterminados.Enabled = True
    Me.cmdCargarDatosPredeterminados.SetFocus
    Call QuitarFocoControl(Me, "cmdRestaurar")
    Me.cmdRestaurar.Visible = False
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdRestaurar_Click", Me)
    Resume LimpiarYSalir
End Sub


' --- GESTIÓN DINÁMICA DE UI (Toggles) ---

Private Sub grpRespProduccion_AfterUpdate(): Call ActualizarUI_Produccion: End Sub
Private Sub grpRespDisenio_AfterUpdate(): Call ActualizarUI_Disenio: End Sub
Private Sub grpRepresentanteSum_AfterUpdate(): Call ActualizarUI_Representante: End Sub

Private Sub ActualizarUI_Produccion()
    ' 1 = Sí, 2 = No
    Dim aplica As Boolean
    aplica = (Nz(Me.grpRespProduccion.value, 2) = 1)
    
    Me.firmaAprobacionRespProduccionNombre.Enabled = aplica
    
    If Not aplica Then
        Me.firmaAprobacionRespProduccionNombre.value = Null
    End If
End Sub

Private Sub ActualizarUI_Disenio()
    Dim aplica As Boolean
    aplica = (Nz(Me.grpRespDisenio.value, 2) = 1)
    Me.firmaAprobacionRespDisenioNombre.Enabled = aplica
    If Not aplica Then
        Me.firmaAprobacionRespDisenioNombre.value = Null
    End If
End Sub

Private Sub ActualizarUI_Representante()
    Dim aplica As Boolean
    aplica = (Nz(Me.grpRepresentanteSum.value, 2) = 1)
    Me.firmaAprobacionRepresentanteSumNombre.Enabled = aplica
    If Not aplica Then
        Me.firmaAprobacionRepresentanteSumNombre.value = Null
    End If
End Sub

Public Function hayCambios() As Boolean
    On Error GoTo Errores
    Dim cambioIng As Boolean, cambioCal As Boolean, cambioProd As Boolean
    Dim cambioDis As Boolean, cambioRep As Boolean
    cambioIng = Nz(Me.firmaAprobacionRespIngenieriaNombre.value, "") <> m_original_IngNombre
    cambioCal = Nz(Me.firmaAprobacionRespCalidadNombre.value, "") <> m_original_CalNombre
    cambioProd = Nz(Me.firmaAprobacionRespProduccionNombre.value, "") <> m_original_ProdNombre
    cambioDis = Nz(Me.firmaAprobacionRespDisenioNombre.value, "") <> m_original_DisNombre
    cambioRep = Nz(Me.firmaAprobacionRepresentanteSumNombre.value, "") <> m_original_RepNombre
    hayCambios = cambioIng Or cambioCal Or cambioProd Or cambioDis Or cambioRep
    Exit Function
Errores:
    hayCambios = True
End Function

Public Sub GuardarAprobacionSuministrador(ByRef vm As DatosCDCAViewModel)
    On Error GoTo Errores
    If Not hayCambios() Then Exit Sub
    Call RecogerValores(vm)
    Exit Sub
Errores:
    ManejadorErroresFormulario "GuardarAprobacionSuministrador", Me
End Sub