Attribute VB_Name = "modGestionRiesgosDatosGeneralesHelper"
' ============================================================
' modGestionRiesgosDatosGeneralesHelper.bas
'
' Helper que extrae la logica de negocio de:
'   - Form_FormGestionRiesgosDatosGenerales.RellenarDatosExpediente (line 344)
'   - Form_FormGestionRiesgosDatosGenerales.RellenarDatosAlObjeto   (line 434)
'
' Slice HR4-sister (HR3d follow-up #9, #5 - 2026-07-01).
'
' Hard rule 1: cero business logic en event handlers o Public methods en forms.
' Hard rule 2: honest signature — solo inyecta lo que la funcion necesita.
' Hard rule 4: Public business methods en forms son anti-patron.
' Hard rule 5: nombres Public globalmente unicos con ModuleName_ prefix.
' Hard rule 7: NO MsgBox/InputBox en helpers. Errores via ByRef p_Error.
'
' Diseno (matching modFormInteractionHelper pattern - takes form Me as Object):
'   - RellenarDatosExpediente: el form pasa Me. El helper escribe
'     los valores del Expediente a los Me.* controls. El form queda
'     como thin adapter que solo delega al helper.
'   - RellenarDatosAlObjeto: el form pasa Me. El helper lee los Me.*
'     controls y los escribe al proyecto. El form queda como thin
'     adapter que solo delega al helper.
' ============================================================
Option Compare Database
Option Explicit

' =============================================================================
' Public Sub DatosGenerales_RellenarDatosExpediente
'   Escribe los valores del Expediente a los Me.* controls del form.
'   Si p_Expediente es Nothing, limpia todos los controls (matching
'   el comportamiento original).
'   Raises: NO. Errores via p_Error ByRef.
' =============================================================================
Public Sub DatosGenerales_RellenarDatosExpediente( _
    ByVal p_Form As Object, _
    ByVal p_Expediente As Expediente, _
    ByVal p_ObjProyectoAlInicio As Proyecto, _
    Optional ByRef p_Error As String)

    On Error GoTo errores
    p_Error = ""

    If p_Expediente Is Nothing Then
        p_Form.idExpediente = Null
        p_Form.Juridica = Null
        p_Form.CodExp = Null
        p_Form.Nemotecnico = Null
        p_Form.Titulo = Null
        p_Form.FechaInicioContrato = Null
        p_Form.FechaFinContrato = Null
        p_Form.OrganoContratacion = Null
        p_Form.ResponsableCalidad = Null
        p_Form.CodigoDocumento = Null
        p_Form.CorreoRAC = Null
        p_Form.EnUTE = Null
        Exit Sub
    End If

    With p_Expediente
        p_Form.idExpediente = .idExpediente
        p_Form.Juridica = .CadenaJuridicas
        p_Form.CodExp = .CodExp
        If Not p_ObjProyectoAlInicio Is Nothing Then
            If p_ObjProyectoAlInicio.NombreProyecto <> "" Then
                p_Form.Nemotecnico = p_ObjProyectoAlInicio.NombreProyecto
            Else
                p_Form.Nemotecnico = Null
            End If
        Else
            If .Nemotecnico <> "" Then
                p_Form.Nemotecnico = .Nemotecnico
            Else
                p_Form.Nemotecnico = .Titulo
            End If
        End If
        p_Form.Titulo = .Titulo

        If IsDate(.FechaInicioContrato) Then
            p_Form.FechaInicioContrato = .FechaInicioContrato
        End If

        If IsDate(.FechaFinContrato) Then
            p_Form.FechaFinContrato = .FechaFinContrato
        End If
        If Not .OrganoContratacion Is Nothing Then
            p_Form.OrganoContratacion = .OrganoContratacion.OrganoContratacion
        End If
        If Not .ResponsableCalidad Is Nothing Then
            p_Form.ResponsableCalidad = .ResponsableCalidad.Nombre
        End If
        If Not p_ObjProyectoAlInicio Is Nothing Then
            If p_ObjProyectoAlInicio.EnUTE <> "" Then
                p_Form.EnUTE = p_ObjProyectoAlInicio.EnUTE
            Else
                p_Form.EnUTE = Null
            End If
        Else
            If .EnUTE = EnumSiNo.Sí Then
                p_Form.EnUTE = "Sí"
            Else
                p_Form.EnUTE = "No"
            End If

        End If

        If Not p_ObjProyectoAlInicio Is Nothing Then
            If p_ObjProyectoAlInicio.CodigoDocumento <> "" Then
                p_Form.CodigoDocumento = p_ObjProyectoAlInicio.CodigoDocumento
            Else
                p_Form.CodigoDocumento = Null
            End If
        Else
            If .CodigoDocParteExpediente <> "" Then
                p_Form.CodigoDocumento = .CodigoDocParteExpediente
            Else
                p_Form.CodigoDocumento = Null
            End If
        End If
        p_Form.CorreoRAC = .CadenaCorreoRACs

    End With
    Exit Sub

errores:
    If Err.Number <> 1000 Then
        p_Error = "DatosGenerales_RellenarDatosExpediente: " & Err.Description
    End If
End Sub

' =============================================================================
' Public Sub DatosGenerales_RellenarDatosAlObjeto
'   Lee los Me.* controls del form y los escribe al Proyecto.
'   Si el Proyecto es Nothing, no-op (matching el comportamiento original).
'   Raises: NO. Errores via p_Error ByRef.
' =============================================================================
Public Sub DatosGenerales_RellenarDatosAlObjeto( _
    ByVal p_Form As Object, _
    ByVal p_ObjProyectoActivo As Proyecto, _
    Optional ByRef p_Error As String)

    On Error GoTo errores
    p_Error = ""

    If p_ObjProyectoActivo Is Nothing Then
        Exit Sub
    End If

    With p_ObjProyectoActivo
        .idExpediente = Nz(p_Form.idExpediente, "")
        .CodigoDocumento = Nz(p_Form.CodigoDocumento, "")
        .ParaInformeAvisos = Nz(p_Form.ParaInformeAvisos, "")
        .EnUTE = Nz(p_Form.EnUTE, "")
        .CorreoRAC = Nz(p_Form.CorreoRAC, "")
        .Elaborado = Nz(p_Form.Elaborado, "")
        .Revisado = Nz(p_Form.Revisado, "")
        .Aprobado = Nz(p_Form.Aprobado, "")

    End With
    Exit Sub

errores:
    If Err.Number <> 1000 Then
        p_Error = "DatosGenerales_RellenarDatosAlObjeto: " & Err.Description
    End If
End Sub
