
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: subfrmDatosPCSUB_AprobacionSuministrador.cls
' VERSIÓN: Clone de Form_subfrmDatosPC_AprobacionSuministrador adaptado para PCSUB.
'           Usa campos sub: firmaOficinaTecnicaSubSuministradorNombre,
'           firmaRepSubSuministradorNombre.
' ==========================================================================

' -- El ViewModel del formulario padre es la fuente de verdad --
Private m_ViewModel As DatosPCSUBViewModel

' -- Variables para guardar el estado previo a la carga de defaults --
Private m_original_firmaOficinaTecnicaSubSuministradorNombre As String
Private m_original_firmaRepSubSuministradorNombre As String

Public Sub PoblarControles(ByRef vm As DatosPCSUBViewModel)
    On Error GoTo Errores
    Set m_ViewModel = vm
    If vm Is Nothing Then Exit Sub
    
    ' Volcar datos
    Me.firmaOficinaTecnicaSubSuministradorNombre = Nz(vm.vm_firmaOficinaTecnicaSubSuministradorNombre, "")
    Me.firmaRepSubSuministradorNombre = Nz(vm.vm_firmaRepSubSuministradorNombre, "")
    
    ' Guardar valores originales para dirty check
    m_original_firmaOficinaTecnicaSubSuministradorNombre = Nz(vm.vm_firmaOficinaTecnicaSubSuministradorNombre, "")
    m_original_firmaRepSubSuministradorNombre = Nz(vm.vm_firmaRepSubSuministradorNombre, "")
    
    ' CORRECCIÓN: Visibilidad y Habilitación forzadas
    Dim puedeEditar As Boolean
    puedeEditar = vm.PuedeGuardar_AprobacionSuministrador
    
    Me.cmdCargarDefaults.Visible = True ' Siempre visible si estamos en el form
    Me.cmdCargarDefaults.Enabled = puedeEditar
    Me.cmdGuardar.Enabled = puedeEditar
    
    ' El botón restaurar solo debe verse si se ha pulsado cargar defaults
    Call QuitarFocoControl(Me, "cmdRestaurarDatos")
    Me.cmdRestaurarDatos.Visible = False

    Exit Sub
Errores:
    ManejadorErroresFormulario "subfrmDatosPCSUB_AprobacionSuministrador.PoblarControles", Me
End Sub

Public Sub RecogerValores(ByRef vm As DatosPCSUBViewModel)
    On Error GoTo Errores
    If vm Is Nothing Then Exit Sub
    
    vm.vm_firmaOficinaTecnicaSubSuministradorNombre = Nz(Me.firmaOficinaTecnicaSubSuministradorNombre.value, "")
    vm.vm_firmaRepSubSuministradorNombre = Nz(Me.firmaRepSubSuministradorNombre.value, "")
    
    Exit Sub
Errores:
    ManejadorErroresFormulario "subfrmDatosPCSUB_AprobacionSuministrador.RecogerValores", Me
End Sub

' --------------------------------------------------------------------------
' EVENTOS DE BOTONES
' --------------------------------------------------------------------------

Private Sub cmdGuardar_Click()
    On Error GoTo Errores
    Me.Parent.GuardarDesdeSubform Me.name
    Call QuitarFocoControl(Me, "cmdRestaurarDatos")
    Me.cmdRestaurarDatos.Visible = False
    Exit Sub
Errores:
    ManejadorErroresFormulario "subfrmDatosPCSUB_AprobacionSuministrador.cmdGuardar_Click", Me
End Sub

Private Sub cmdCargarDefaults_Click()
    Dim pcServ As New DatosPCSUBServicio
    Dim pcInicial As DatosPCSUB
    
    On Error GoTo Errores
    If m_ViewModel Is Nothing Then Exit Sub
    
    Me.Parent.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    
    ' 1. Guardar estado actual
    m_original_firmaOficinaTecnicaSubSuministradorNombre = Nz(Me.firmaOficinaTecnicaSubSuministradorNombre.value, "")
    m_original_firmaRepSubSuministradorNombre = Nz(Me.firmaRepSubSuministradorNombre.value, "")
    
    ' 2. Llamar al servicio para obtener los datos iniciales
    Set pcInicial = pcServ.GetDatosAprobacionIniciales(m_ViewModel.idSolicitud)
    
    ' 3. Volcar datos en los controles
    If Not pcInicial Is Nothing Then
        Me.firmaOficinaTecnicaSubSuministradorNombre.value = pcInicial.firmaOficinaTecnicaSubSuministradorNombre
        Me.firmaRepSubSuministradorNombre.value = pcInicial.firmaRepSubSuministradorNombre
    End If
    
    ' 4. Gestionar visibilidad/habilitación de botones
    Me.cmdCargarDefaults.Enabled = False
    Me.cmdRestaurarDatos.Visible = True
    Me.cmdRestaurarDatos.Enabled = True
LimpiarYSalir:
    Me.Parent.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Set pcServ = Nothing
    Set pcInicial = Nothing
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("subfrmDatosPCSUB_AprobacionSuministrador.cmdCargarDefaults_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdRestaurarDatos_Click()
    On Error GoTo Errores
    
    ' 1. Restaurar valores desde las variables de módulo
    Me.firmaOficinaTecnicaSubSuministradorNombre.value = m_original_firmaOficinaTecnicaSubSuministradorNombre
    Me.firmaRepSubSuministradorNombre.value = m_original_firmaRepSubSuministradorNombre
     ' 3. Volver a habilitar el botón de cargar
    Me.cmdCargarDefaults.Enabled = True
    Me.cmdCargarDefaults.SetFocus
    ' 2. Ocultar el botón de restaurar
    Call QuitarFocoControl(Me, "cmdRestaurarDatos")
    Me.cmdRestaurarDatos.Visible = False
    
   
    
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("subfrmDatosPCSUB_AprobacionSuministrador.cmdRestaurarDatos_Click", Me)
End Sub

Public Function hayCambios() As Boolean
    On Error GoTo Errores
    Dim cambioOT As Boolean, cambioRep As Boolean
    cambioOT = Nz(Me.firmaOficinaTecnicaSubSuministradorNombre.value, "") <> m_original_firmaOficinaTecnicaSubSuministradorNombre
    cambioRep = Nz(Me.firmaRepSubSuministradorNombre.value, "") <> m_original_firmaRepSubSuministradorNombre
    hayCambios = cambioOT Or cambioRep
    Exit Function
Errores:
    hayCambios = True
End Function

Public Sub GuardarAprobacionSuministrador(ByRef vm As DatosPCSUBViewModel)
    On Error GoTo Errores
    If Not hayCambios() Then Exit Sub
    Call RecogerValores(vm)
Errores:
    ManejadorErroresFormulario "GuardarAprobacionSuministrador", Me
End Sub
