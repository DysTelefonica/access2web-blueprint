
' ==========================================================================
' MÓDULO: Form_subfrmDatosPC_AprobacionSuministrador.cls
' VERSIÓN: DEFINITIVA con carga de defaults, restauración y lógica de habilitación.
' ==========================================================================
Option Compare Database
Option Explicit

' -- El ViewModel del formulario padre es la fuente de verdad --
Private m_ViewModel As DatosPCViewModel

' -- Variables para guardar el estado previo a la carga de defaults --
Private m_original_firmaOficinaTecnicaNombre As String
Private m_original_firmaRepSuministradorNombre As String

' EN: Form_subfrmDatosPC_AprobacionSuministrador.cls

Public Sub PoblarControles(ByRef vm As DatosPCViewModel)
    On Error GoTo Errores
    Set m_ViewModel = vm
    If vm Is Nothing Then Exit Sub
    
    ' Volcar datos
    Me.firmaOficinaTecnicaNombre = Nz(vm.vm_firmaOficinaTecnicaNombre, "")
    Me.firmaRepSuministradorNombre = Nz(vm.vm_firmaRepSuministradorNombre, "")
    
    ' Guardar valores originales para dirty check
    m_original_firmaOficinaTecnicaNombre = Nz(vm.vm_firmaOficinaTecnicaNombre, "")
    m_original_firmaRepSuministradorNombre = Nz(vm.vm_firmaRepSuministradorNombre, "")
    
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
    ManejadorErroresFormulario "subfrmDatosPC_AprobacionSuministrador.PoblarControles", Me
End Sub

Public Sub RecogerValores(ByRef vm As DatosPCViewModel)
    On Error GoTo Errores
    If vm Is Nothing Then Exit Sub
    
    vm.vm_firmaOficinaTecnicaNombre = Nz(Me.firmaOficinaTecnicaNombre.value, "")
    vm.vm_firmaRepSuministradorNombre = Nz(Me.firmaRepSuministradorNombre.value, "")
    
    Exit Sub
Errores:
    ManejadorErroresFormulario "subfrmDatosPC_AprobacionSuministrador.RecogerValores", Me
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
    ManejadorErroresFormulario "subfrmDatosPC_AprobacionSuministrador.cmdGuardar_Click", Me
End Sub

Private Sub cmdCargarDefaults_Click()
    Dim pcServ As New DatosPCServicio
    Dim pcInicial As DatosPC
    
    On Error GoTo Errores
    If m_ViewModel Is Nothing Then Exit Sub
    
    Me.Parent.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    
    ' 1. Guardar estado actual
    m_original_firmaOficinaTecnicaNombre = Nz(Me.firmaOficinaTecnicaNombre.value, "")
    m_original_firmaRepSuministradorNombre = Nz(Me.firmaRepSuministradorNombre.value, "")
    
    ' 2. Llamar al servicio para obtener los datos iniciales
    Set pcInicial = pcServ.GetDatosAprobacionSuministradorIniciales(m_ViewModel.idSolicitud)
    
    ' 3. Volcar datos en los controles
    If Not pcInicial Is Nothing Then
        Me.firmaOficinaTecnicaNombre.value = pcInicial.firmaOficinaTecnicaNombre
        Me.firmaRepSuministradorNombre.value = pcInicial.firmaRepSuministradorNombre
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
    Call ManejadorErroresFormulario("subfrmDatosPC_AprobacionSuministrador.cmdCargarDefaults_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdRestaurarDatos_Click()
    On Error GoTo Errores
    
    ' 1. Restaurar valores desde las variables de módulo
    Me.firmaOficinaTecnicaNombre.value = m_original_firmaOficinaTecnicaNombre
    Me.firmaRepSuministradorNombre.value = m_original_firmaRepSuministradorNombre
     ' 3. Volver a habilitar el botón de cargar
    Me.cmdCargarDefaults.Enabled = True
    Me.cmdCargarDefaults.SetFocus
    ' 2. Ocultar el botón de restaurar
    Call QuitarFocoControl(Me, "cmdRestaurarDatos")
    Me.cmdRestaurarDatos.Visible = False
    
   
    
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("subfrmDatosPC_AprobacionSuministrador.cmdRestaurarDatos_Click", Me)
End Sub

Public Function hayCambios() As Boolean
    On Error GoTo Errores
    Dim cambioOT As Boolean, cambioRep As Boolean
    cambioOT = Nz(Me.firmaOficinaTecnicaNombre.value, "") <> m_original_firmaOficinaTecnicaNombre
    cambioRep = Nz(Me.firmaRepSuministradorNombre.value, "") <> m_original_firmaRepSuministradorNombre
    hayCambios = cambioOT Or cambioRep
    Exit Function
Errores:
    hayCambios = True
End Function

Public Sub GuardarAprobacionSuministrador(ByRef vm As DatosPCViewModel)
    On Error GoTo Errores
    If Not hayCambios() Then Exit Sub
    Call RecogerValores(vm)
Errores:
    ManejadorErroresFormulario "GuardarAprobacionSuministrador", Me
End Sub