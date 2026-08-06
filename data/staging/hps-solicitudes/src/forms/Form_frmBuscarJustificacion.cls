' -------------------------------------------------------------------------
' FORMULARIO: frmBuscarJustificacion
' DESCRIPCIÓN: Buscador y Gestor de Justificaciones.
'              Soporta modo "Selección" (para Solicitudes) y modo "Gestión".
' -------------------------------------------------------------------------
Option Compare Database
Option Explicit

' --- EVENTOS PÚBLICOS ---
' Este evento se lanza cuando el usuario elige una justificación (ya sea manual o automática tras alta)
Public Event JustificacionElegida(id As Long, Titulo As String)

' --- VARIABLES PRIVADAS ---
' Referencia al formulario hijo para capturar sus eventos (WithEvents es clave aquí)
Private WithEvents f_Detalle As Form_frmJustificacion

' Servicio para lógica de negocio (borrados, etc.)
Private m_Servicio As New JustificacionServicio

' Bandera para saber en qué modo estamos
Private m_ModoSeleccion As Boolean

Private Sub cmdSalir_Click()
    DoCmd.Close acForm, Me.Name
End Sub

' -------------------------------------------------------------------------
' INICIALIZACIÓN Y CARGA
' -------------------------------------------------------------------------
Private Sub Form_Load()
    Ajustar Me
    ' Leemos OpenArgs para determinar el modo de apertura
    ' Si el formulario padre nos pasó "MODO_SELECCION", activamos el botón Elegir
    If Nz(Me.OpenArgs, "") = "MODO_SELECCION" Then
        m_ModoSeleccion = True
        Me.cmdElegir.Visible = True
    Else
        m_ModoSeleccion = False
        Me.cmdElegir.Visible = False
    End If
    Me.cboFiltroEstado.value = "Activas"
    ConfigurarLista
    CargarLista
End Sub
Private Sub txtBuscar_AfterUpdate()
    ' 1. PRIMERO: Ejecutamos el filtro con lo que haya escrito el usuario
    '    Usamos .Value porque en AfterUpdate el valor ya está confirmado
    CargarLista Nz(Me.txtBuscar.value, "")
    
    ' 2. SEGUNDO: Aplicamos la lógica de "Auto-selección" si el resultado es único
    If Me.lstResultados.ListCount = 2 Then
        ' (ListCount 2 = 1 Cabecera + 1 Resultado)
        
        ' Seleccionamos la fila
        Me.lstResultados.Selected(1) = True
        
        ' Ejecutamos la acción según el modo
        If m_ModoSeleccion Then
            cmdElegir_Click
        Else
            btnEditar_Click
        End If
        
    ElseIf Me.lstResultados.ListCount > 2 Then
        ' Si hay varios resultados, mandamos el foco a la lista para navegar con flechas
        Me.lstResultados.SetFocus
    End If
End Sub
Private Sub ConfigurarLista()
    ' Configuración visual básica por código (por si acaso)
    Me.lstResultados.ColumnCount = 5
    Me.lstResultados.ColumnWidths = "0cm;1cm;4cm;14cm;0cm" ' Ocultamos ID
End Sub
Private Sub txtBuscar_Change()
    ' El evento Change se dispara con cada pulsación de tecla.
    ' IMPORTANTE: Mientras escribimos, el valor real (.Value) no se ha actualizado aún,
    ' por eso debemos usar la propiedad .Text para saber qué está viendo el usuario.
    
End Sub

Private Sub cboFiltroEstado_AfterUpdate()
    ' Cuando el usuario cambia "Activas" a "Todas", recargamos la lista.
    ' IMPORTANTE: Le pasamos lo que haya en el buscador de texto (txtBuscar)
    ' para no perder el filtro que el usuario ya hubiera escrito.
    
    CargarLista Nz(Me.txtBuscar.value, "")
End Sub
Private Sub CargarLista(Optional sFiltroTexto As String = "")
    Dim col As Collection
    Dim just As Justificacion
    Dim sFiltroCombo As String
    Dim bCumpleEstado As Boolean
    Dim sSimbolo As String
    
    ' Leemos el valor del combo
    sFiltroCombo = Nz(Me.cboFiltroEstado.value, "Activas")
    
    Set col = m_ObjEntorno.Justificaciones
    
    Me.lstResultados.RowSource = ""
    ' Añadimos cabecera (incluyendo la columna oculta extra)
    Me.lstResultados.AddItem "ID;Est.;Justificación;Descripción;RawEstado"
    
    If col Is Nothing Then Exit Sub
    
    For Each just In col
        ' 1. FILTRO COMBO
        Select Case sFiltroCombo
            Case "Activas":   bCumpleEstado = just.activa
            Case "Inactivas": bCumpleEstado = Not just.activa
            Case "Todas":     bCumpleEstado = True
        End Select
        
        If bCumpleEstado Then
            ' 2. FILTRO TEXTO
            If Len(sFiltroTexto) = 0 Or _
               InStr(1, just.Titulo, sFiltroTexto, vbTextCompare) > 0 Or _
               InStr(1, just.Descripcion, sFiltroTexto, vbTextCompare) > 0 Then
                
                ' 3. SÍMBOLO (Usamos ChrW para evitar errores en el editor VBA)
                If just.activa Then
                    sSimbolo = ChrW(10004) ' ? (Check Heavy)
                Else
                    sSimbolo = ChrW(10006) ' ? (Cross Heavy)
                End If
                
                ' 4. AÑADIR A LISTA
                ' Observa el último valor: CStr(just.activa). Es la columna oculta 4.
                Me.lstResultados.AddItem just.IDJustificacion & ";" & _
                                         sSimbolo & ";" & _
                                         just.Titulo & ";" & _
                                         just.Descripcion & ";" & _
                                         CStr(just.activa)
            End If
        End If
    Next just
    
    ActualizarEstadoBotones
End Sub
'---------------------------------------------------------------
' GESTIÓN (ALTA / EDICIÓN / BAJA)
' -------------------------------------------------------------------------

Private Sub btnAlta_Click()
    ' Abrimos el formulario de detalle.
    ' IMPORTANTE: NO usamos acDialog aquí. Si usáramos acDialog, el código se detendría
    ' y no podríamos capturar el evento "JustificacionGuardada".
    DoCmd.OpenForm "frmJustificacion"
    
    ' Enlazamos la variable con eventos al formulario recién abierto
    Set f_Detalle = Forms("frmJustificacion")
End Sub

Private Sub btnEditar_Click()
    ' Validar selección
    If IsNull(Me.lstResultados) Then
        MsgBox "Por favor, selecciona un registro para editar.", vbExclamation, "Atención"
        Exit Sub
    End If
    
    Dim idSel As Long
    idSel = Me.lstResultados.Column(0)
    
    ' Abrimos el formulario pasando el ID en OpenArgs
    DoCmd.OpenForm "frmJustificacion", , , , , , idSel
    
    ' También capturamos eventos al editar, para refrescar la lista si cambian el nombre
    Set f_Detalle = Forms("frmJustificacion")
End Sub

Private Sub btnEliminar_Click()
    If IsNull(Me.lstResultados) Then
        MsgBox "Selecciona un registro para eliminar.", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("¿Estás seguro de eliminar (desactivar) esta justificación?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    
    Dim idSel As Long
    idSel = Me.lstResultados.Column(0)
    
    ' Delegamos al servicio (Arquitectura limpia)
    On Error GoTo ErrorEliminar
    m_Servicio.EliminarJustificacion idSel
    
    ' Refrescamos (El servicio ya invalidó la caché, así que CargarLista traerá datos frescos)
    CargarLista
    Exit Sub

ErrorEliminar:
    MsgBox "No se pudo eliminar: " & Err.Description, vbCritical
End Sub

' -------------------------------------------------------------------------
' LÓGICA DE EVENTOS CRUZADOS (DAISY CHAIN)
' -------------------------------------------------------------------------

' Este subrutina se dispara AUTOMÁTICAMENTE cuando el formulario hijo (frmJustificacion)
' lanza su evento "JustificacionGuardada" al terminar de guardar.
' --- CAPTURA DEL EVENTO DEL HIJO ---
Private Sub f_Detalle_JustificacionGuardada(id As Long, Titulo As String)
    
    ' 1. Invalidar caché global para asegurar que tenemos los datos nuevos
    m_ObjEntorno.RecargarJustificaciones
    
    ' 2. Refrescar la lista APLICANDO EL FILTRO ACTUAL
    '    CAMBIO CLAVE: Pasamos el valor de txtBuscar a CargarLista.
    '    Usamos Nz() por si el cuadro de texto está vacío (Null).
    CargarLista Nz(Me.txtBuscar.value, "")
    
    ' 3. Lógica inteligente:
    If m_ModoSeleccion Then
        ' CASO A: Modo Selección (Venimos de Solicitud)
        ' Si el registro sigue cumpliendo el filtro (existe en la lista), lo elegimos.
        ' Si desapareció por el filtro, cerramos igual devolviendo el ID (Access lo manejará).
        RaiseEvent JustificacionElegida(id, Titulo)
        DoCmd.Close acForm, Me.Name
    Else
        ' CASO B: Modo Gestión (Desde menú Entidades)
        ' Intentamos resaltar la fila editada.
        ' Nota: Si cambiaste el nombre y ya no coincide con el filtro,
        ' ResaltarFila simplemente no encontrará el ID y no seleccionará nada (correcto).
        ResaltarFila id
        
        ' IMPORTANTE: Forzamos actualización de botones (Editar/Eliminar)
        ' porque al recargar la lista se habrán deshabilitado.
        ActualizarEstadoBotones
    End If
    
    ' Limpiamos la referencia
    Set f_Detalle = Nothing
End Sub
' -------------------------------------------------------------------------
' SELECCIÓN MANUAL
' -------------------------------------------------------------------------

Private Sub cmdElegir_Click()
    If IsNull(Me.lstResultados) Then
        MsgBox "Debes seleccionar una justificación de la lista.", vbExclamation
        Exit Sub
    End If
    
    ' CORRECCIÓN DE COLUMNAS:
    ' Param 1 (ID)   -> Column(0)
    ' Param 2 (Texto)-> Column(2) (Es el Título. La Column(1) es el símbolo ?)
    
    RaiseEvent JustificacionElegida(Me.lstResultados.Column(0), Me.lstResultados.Column(2))
    
    ' Cerramos este buscador
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub lstResultados_DblClick(Cancel As Integer)
    ' Doble clic actúa inteligentemente según el modo
    If m_ModoSeleccion Then
        cmdElegir_Click ' Seleccionar
    Else
        btnEditar_Click ' Editar
    End If
End Sub



' -------------------------------------------------------------------------
' AYUDAS PRIVADAS
' -------------------------------------------------------------------------
Private Sub ResaltarFila(id As Long)
    ' Busca el ID en la lista y lo selecciona visualmente
    Dim i As Integer
    For i = 1 To Me.lstResultados.ListCount - 1
        If CLng(Me.lstResultados.Column(0, i)) = id Then
            Me.lstResultados.Selected(i) = True
            Exit For
        End If
    Next i
End Sub
' -------------------------------------------------------------------------
' GESTIÓN DE ESTADO DE BOTONES
' -------------------------------------------------------------------------
Private Sub ActualizarEstadoBotones()
    Dim idSeleccionado As Long
    Dim estaEnUso As Boolean
    
    ' Si no hay selección válida (saltamos cabecera o nulos)
    If IsNull(Me.lstResultados) Or Nz(Me.lstResultados.Column(0)) = 0 Then
        Me.btnEditar.Enabled = False
        Me.btnEliminar.Enabled = False
        If m_ModoSeleccion Then Me.cmdElegir.Enabled = False
        Exit Sub
    End If
    
    idSeleccionado = Nz(Me.lstResultados.Column(0))
    If idSeleccionado = 0 Then
        Exit Sub
    End If
    ' --- REGLA 1: Solo permitir ELEGIR si está Activa ---
    ' Leemos la columna OCULTA (Índice 4) que tiene el valor real True/False
    Dim esActiva As Boolean
    esActiva = CBool(Me.lstResultados.Column(4))
    
    If m_ModoSeleccion Then
        If esActiva Then
            Me.cmdElegir.Enabled = True
            Me.cmdElegir.ControlTipText = "Seleccionar esta justificación."
        Else
            Me.cmdElegir.Enabled = False
            Me.cmdElegir.ControlTipText = "No se puede seleccionar: Está inactiva."
        End If
    End If
    
    ' --- REGLA 2: EDITAR siempre permitido ---
    Me.btnEditar.Enabled = True
    
    ' --- REGLA 3: ELIMINAR solo si no está en uso ---
    estaEnUso = m_Servicio.VerificarUso(idSeleccionado)
    
    If estaEnUso Then
        Me.btnEliminar.Enabled = False
        Me.btnEliminar.ControlTipText = "En uso. No se puede eliminar."
    Else
        Me.btnEliminar.Enabled = True
        Me.btnEliminar.ControlTipText = "Eliminar (Desactivar)."
    End If
End Sub
Private Sub lstResultados_AfterUpdate()
    ActualizarEstadoBotones
End Sub

' También es recomendable ponerlo en el evento Click por seguridad
Private Sub lstResultados_Click()
    ActualizarEstadoBotones
End Sub

