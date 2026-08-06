
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: Form_subfrmDatosCDCASUB_AprobacionSuministrador
' ==========================================================================

Private m_ViewModel As DatosCDCASUBViewModel

' Snapshot
Private m_original_IngNombre As String
Private m_original_CalNombre As String
Private m_original_ProdNombre As String
Private m_original_DisNombre As String
Private m_original_RepNombre As String

Public Sub PoblarControles(ByRef vm As DatosCDCASUBViewModel)
    On Error GoTo Errores
    Set m_ViewModel = vm
    If vm Is Nothing Then Exit Sub

    ' Guardar valores originales para dirty check
    m_original_IngNombre = Nz(vm.vm_firmaAprobacionRespIngenieriaNombre, "")
    m_original_CalNombre = Nz(vm.vm_firmaAprobacionRespCalidadNombre, "")
    m_original_ProdNombre = Nz(vm.vm_firmaAprobacionRespProduccionNombre, "")
    m_original_DisNombre = Nz(vm.vm_firmaAprobacionRespDisenioNombre, "")
    m_original_RepNombre = Nz(vm.vm_firmaAprobacionRepresentanteSumNombre, "")

    ' 1. Firmas Obligatorias
    Me.firmaAprobacionRespIngenieriaNombre = Nz(vm.vm_firmaAprobacionRespIngenieriaNombre, "")
    Me.firmaAprobacionRespCalidadNombre = Nz(vm.vm_firmaAprobacionRespCalidadNombre, "")
    
    ' 2. Firmas Opcionales
    Me.firmaAprobacionRespProduccionNombre = Nz(vm.vm_firmaAprobacionRespProduccionNombre, "")
    Me.firmaAprobacionRespDisenioNombre = Nz(vm.vm_firmaAprobacionRespDisenioNombre, "")
    Me.firmaAprobacionRepresentanteSumNombre = Nz(vm.vm_firmaAprobacionRepresentanteSumNombre, "")
    
    ' 3. Ajustar UI (Toggles)
    If vm.vm_aplicaRespProduccion Then Me.grpRespProduccion.value = 1 Else Me.grpRespProduccion.value = 2
    Call ActualizarUI_Produccion
    
    If vm.vm_aplicaRespDisenio Then Me.grpRespDisenio.value = 1 Else Me.grpRespDisenio.value = 2
    Call ActualizarUI_Disenio
    
    If vm.vm_aplicaRepresentanteSum Then Me.grpRepresentanteSum.value = 1 Else Me.grpRepresentanteSum.value = 2
    Call ActualizarUI_Representante
    
    ' 4. Visibilidad
    Dim puedeEditar As Boolean
    puedeEditar = vm.PuedeGuardar_AprobacionSuministrador
    
    Me.cmdCargarDatosPredeterminados.Visible = True
    Me.cmdCargarDatosPredeterminados.Enabled = puedeEditar
    Me.cmdGuardar.Enabled = puedeEditar
    If Me.cmdCargarDatosPredeterminados.Enabled = True Then Me.cmdCargarDatosPredeterminados.SetFocus
    
    Call QuitarFocoControl(Me, "cmdRestaurar")
    
    Me.cmdRestaurar.Visible = False
    
    Exit Sub
Errores:
    ManejadorErroresFormulario "subfrmDatosCDCASUB_AprobacionSuministrador.PoblarControles", Me
End Sub

Public Sub RecogerValores(ByRef vm As DatosCDCASUBViewModel)
    On Error GoTo Errores
    If vm Is Nothing Then Exit Sub
    
    vm.vm_firmaAprobacionRespIngenieriaNombre = Nz(Me.firmaAprobacionRespIngenieriaNombre.value, "")
    vm.vm_firmaAprobacionRespCalidadNombre = Nz(Me.firmaAprobacionRespCalidadNombre.value, "")
    
    vm.vm_firmaAprobacionRespProduccionNombre = Nz(Me.firmaAprobacionRespProduccionNombre.value, "")
    vm.vm_firmaAprobacionRespDisenioNombre = Nz(Me.firmaAprobacionRespDisenioNombre.value, "")
    vm.vm_firmaAprobacionRepresentanteSumNombre = Nz(Me.firmaAprobacionRepresentanteSumNombre.value, "")

    vm.vm_aplicaRespProduccion = (Nz(Me.grpRespProduccion.value, 2) = 1)
    vm.vm_aplicaRespDisenio = (Nz(Me.grpRespDisenio.value, 2) = 1)
    vm.vm_aplicaRepresentanteSum = (Nz(Me.grpRepresentanteSum.value, 2) = 1)

LimpiarYSalir:
    Exit Sub
Errores:
    ManejadorErroresFormulario "RecogerValores", Me
End Sub

Private Sub cmdGuardar_Click()
    On Error GoTo Errores
    Me.Parent.GuardarDesdeSubform Me.name
    Call QuitarFocoControl(Me, "cmdRestaurar")
    Me.cmdRestaurar.Visible = False
    Me.cmdCargarDatosPredeterminados.Enabled = m_ViewModel.PuedeGuardar_AprobacionSuministrador
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdGuardar_Click", Me)
End Sub

Private Sub cmdCargarDatosPredeterminados_Click()
    Dim srv As New DatosCDCASUBServicio
    Dim def As DatosCDCASUB
    On Error GoTo Errores
    If m_ViewModel Is Nothing Then Exit Sub
    
    Me.Parent.lblBloqueo.Visible = True: DoCmd.Hourglass True
    
    m_original_IngNombre = Nz(Me.firmaAprobacionRespIngenieriaNombre.value, "")
    m_original_CalNombre = Nz(Me.firmaAprobacionRespCalidadNombre.value, "")
    
    Set def = srv.GetDatosAprobacionIniciales(m_ViewModel.Solicitud.idSolicitud)
    
    Me.firmaAprobacionRespIngenieriaNombre = def.firmaAprobacionRespIngenieriaNombre
    Me.firmaAprobacionRespCalidadNombre = def.firmaAprobacionRespCalidadNombre
    
    Me.cmdCargarDatosPredeterminados.Enabled = False
    Me.cmdRestaurar.Visible = True: Me.cmdRestaurar.Enabled = True
    
LimpiarYSalir:
    Me.Parent.lblBloqueo.Visible = False: DoCmd.Hourglass False
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdCargarDatosPredeterminados_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdRestaurar_Click()
    On Error GoTo Errores
    Me.firmaAprobacionRespIngenieriaNombre = m_original_IngNombre
    Me.firmaAprobacionRespCalidadNombre = m_original_CalNombre
    
    Me.cmdCargarDatosPredeterminados.Enabled = True
    Me.cmdCargarDatosPredeterminados.SetFocus
    Call QuitarFocoControl(Me, "cmdRestaurar")
    Me.cmdRestaurar.Visible = False
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdRestaurarDatos_Click", Me)
End Sub



Private Sub grpRespProduccion_AfterUpdate(): Call ActualizarUI_Produccion: End Sub
Private Sub grpRespDisenio_AfterUpdate(): Call ActualizarUI_Disenio: End Sub
Private Sub grpRepresentanteSum_AfterUpdate(): Call ActualizarUI_Representante: End Sub

Private Sub ActualizarUI_Produccion()
    Dim aplica As Boolean: aplica = (Nz(Me.grpRespProduccion.value, 2) = 1)
    Me.firmaAprobacionRespProduccionNombre.Enabled = aplica
    If Not aplica Then
        Me.firmaAprobacionRespProduccionNombre.value = Null
    End If
End Sub

Private Sub ActualizarUI_Disenio()
    Dim aplica As Boolean: aplica = (Nz(Me.grpRespDisenio.value, 2) = 1)
    Me.firmaAprobacionRespDisenioNombre.Enabled = aplica
    If Not aplica Then
        Me.firmaAprobacionRespDisenioNombre.value = Null
    End If
End Sub

Private Sub ActualizarUI_Representante()
    Dim aplica As Boolean: aplica = (Nz(Me.grpRepresentanteSum.value, 2) = 1)
    Me.firmaAprobacionRepresentanteSumNombre.Enabled = aplica
    If Not aplica Then
        Me.firmaAprobacionRepresentanteSumNombre.value = Null
    End If
End Sub

Public Function hayCambios() As Boolean
    On Error GoTo Errores
    
    Dim cambioIng As Boolean
    Dim cambioCal As Boolean
    Dim cambioProd As Boolean
    Dim cambioDis As Boolean
    Dim cambioRep As Boolean
    
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

Public Sub GuardarAprobacionSuministrador(ByRef vm As DatosCDCASUBViewModel)
    On Error GoTo Errores
    
    If Not hayCambios() Then
        Exit Sub
    End If
    
    Call RecogerValores(vm)
    Exit Sub
Errores:
    ManejadorErroresFormulario "GuardarAprobacionSuministrador", Me
End Sub