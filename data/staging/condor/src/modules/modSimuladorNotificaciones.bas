Attribute VB_Name = "modSimuladorNotificaciones"

Option Compare Database
Option Explicit

Public Sub SimularTodasLasNotificaciones()
    Call Test_Notif_Rechazo_Tecnico
    Call Test_Notif_Revision_Calidad
    Call Test_Notif_Avance_Validacion
    Call Test_Notif_Envio_RAC
    Call Test_Notif_Reset_Formalizacion
    Call Test_Notif_Reset_Tecnico
End Sub

Public Sub Test_Notif_Rechazo_Tecnico()
    Dim vm As SolicitudViewModel
    Dim asunto As String
    Set vm = getViewModelEjemplo()
    If vm Is Nothing Then Exit Sub
    asunto = "CONDOR - RECHAZO | " & vm.Expediente.Nemotecnico & " | " & vm.TipoSolicitudNombre & " | " & vm.Solicitud.codigoSolicitud
    Call DispararNotificacionTest(vm, asunto, vm.Expediente.EmailResponsableTecnico)
End Sub

Public Sub Test_Notif_Revision_Calidad()
    Dim vm As SolicitudViewModel
    Dim asunto As String
    Set vm = getViewModelEjemplo()
    If vm Is Nothing Then Exit Sub
    asunto = "CONDOR - REVIS" & Chr(211) & "N | " & vm.Expediente.Nemotecnico & " | " & vm.TipoSolicitudNombre & " | " & vm.Solicitud.codigoSolicitud
    Call DispararNotificacionTest(vm, asunto, vm.Expediente.EmailResponsableCalidad)
End Sub

Public Sub Test_Notif_Avance_Validacion()
    Dim vm As SolicitudViewModel
    Dim asunto As String
    Set vm = getViewModelEjemplo()
    If vm Is Nothing Then Exit Sub
    asunto = "CONDOR - AVANCE | " & vm.Expediente.Nemotecnico & " | " & vm.TipoSolicitudNombre & " | " & vm.Solicitud.codigoSolicitud
    Call DispararNotificacionTest(vm, asunto, vm.Expediente.EmailResponsableCalidad)
End Sub

Public Sub Test_Notif_Envio_RAC()
    Dim vm As SolicitudViewModel
    Dim asunto As String
    Set vm = getViewModelEjemplo()
    If vm Is Nothing Then Exit Sub
    asunto = "CONDOR - ENVIO RAC | " & vm.Expediente.Nemotecnico & " | " & vm.TipoSolicitudNombre & " | " & vm.Solicitud.codigoSolicitud
    Call DispararNotificacionTest(vm, asunto, vm.Expediente.EmailResponsableCalidad)
End Sub

Public Sub Test_Notif_Reset_Formalizacion()
    Dim vm As SolicitudViewModel
    Dim asunto As String
    Set vm = getViewModelEjemplo()
    If vm Is Nothing Then Exit Sub
    asunto = "CONDOR - RESET FORMALIZACION | " & vm.Expediente.Nemotecnico & " | " & vm.TipoSolicitudNombre & " | " & vm.Solicitud.codigoSolicitud
    Call DispararNotificacionTest(vm, asunto, vm.Expediente.EmailResponsableCalidad)
End Sub

Public Sub Test_Notif_Reset_Tecnico()
    Dim vm As SolicitudViewModel
    Dim asunto As String
    Set vm = getViewModelEjemplo()
    If vm Is Nothing Then Exit Sub
    asunto = "CONDOR - RESET TECNICO | " & vm.Expediente.Nemotecnico & " | " & vm.TipoSolicitudNombre & " | " & vm.Solicitud.codigoSolicitud
    Call DispararNotificacionTest(vm, asunto, vm.Expediente.EmailResponsableTecnico)
End Sub

Public Sub DispararNotificacionTest(ByRef vm As SolicitudViewModel, ByVal asunto As String, ByVal destinatarioOriginal As String)
    Dim notifServ As New NotificacionServicio
    Dim destinatarioTest As String
    Dim copia As String
    Dim extraHtml As String
    destinatarioTest = "andres.romandelperal@telefonica.com"
    copia = ""
    extraHtml = "<div class='dev-notice'><b>SIMULACION:</b> Destinatario original -> " & Nz(destinatarioOriginal, "") & "</div>"
    Call notifServ.EnviarNotificacionConDetalle(vm, asunto, destinatarioTest, copia, extraHtml)
End Sub

Private Function getViewModelEjemplo() As SolicitudViewModel
    Dim solServ As New SolicitudServicio
    Dim idEjemplo As Long
    idEjemplo = getIDSolicitudEjemplo()
    If idEjemplo <= 0 Then
        Set getViewModelEjemplo = Nothing
        Exit Function
    End If
    Set getViewModelEjemplo = solServ.getSolicitudViewModelPorID(idEjemplo)
End Function

Private Function getIDSolicitudEjemplo() As Long
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rcd As DAO.Recordset
    Dim sql As String
    Set db = getdb()
    sql = "SELECT TOP 1 idSolicitud FROM tbSolicitudes ORDER BY fechaCreacion DESC;"
    Set qdf = db.CreateQueryDef("", sql)
    Set rcd = qdf.OpenRecordset(dbOpenSnapshot)
    If Not rcd.EOF Then
        getIDSolicitudEjemplo = Nz(rcd.Fields(0).value, 0)
    Else
        getIDSolicitudEjemplo = 0
    End If
    rcd.Close
End Function

