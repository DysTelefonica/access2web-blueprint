VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Public frm As Form
Public OrdinalMenuDesplegable As Single
Public NumeroDeMenus As Single
Public NumeroDeSubMenus As Single
Public NombreSubForm As String
Public TieneSubMenus As EnumSino


Public Sub OcultarSubMenus()
    
    Dim ctlSubMenu As Control
    Dim ctlLabel As Control
    Dim ctl As Control
    Dim i As Single
    
    If TieneSubMenus = EnumSino.No Then
        Exit Sub
    End If
    For i = 1 To NumeroDeMenus
        Set ctlSubMenu = Nothing
        Set ctlSubMenu = getControl(frm, "SUBMENU" & i)
        If Not ctlSubMenu Is Nothing Then
            ctlSubMenu.Visible = False
        End If
        Set ctlLabel = Nothing
        Set ctlLabel = getControl(frm, "lblSUBMENU" & i)
        If Not ctlLabel Is Nothing Then
            ctlLabel.Visible = False
        End If

    Next
    
End Sub
Public Sub MostrarSubMenus()
    Dim ctlSubMenu As Control
    Dim ctlLabel As Control
    Dim ctl As Control
    Dim i As Single
    
    
    If TieneSubMenus = EnumSino.No Then
        Exit Sub
    End If
    For i = 1 To NumeroDeMenus
        Set ctlSubMenu = Nothing
        Set ctlSubMenu = getControl(frm, "SUBMENU" & i)
        If Not ctlSubMenu Is Nothing Then
            ctlSubMenu.BackColor = m_ObjEntorno.ColorMenuDesactivado
            ctlSubMenu.Visible = True
        End If
        Set ctlLabel = Nothing
        Set ctlLabel = getControl(frm, "lblSUBMENU" & i)
        If Not ctlLabel Is Nothing Then
            ctlLabel.Visible = True
        End If

    Next
    
End Sub

Public Sub ActDesActMenu(p_Numero As Single, Optional p_EsActivar As EnumSino = EnumSino.Sí)
    Dim ctlInd As Control
    Dim ctlMenu As Control
    Dim i As Single
    
    
   
    If p_EsActivar = EnumSino.No Then
        Set ctlInd = getControl(frm, "IND" & i)
        If Not ctlInd Is Nothing Then
            If ctlInd.Visible = True Then
                ctlInd.Visible = False
            End If
        End If
        Set ctlMenu = getControl(frm, "MENU" & i)
        If Not ctlMenu Is Nothing Then
            If ctlMenu.BackColor = m_ObjEntorno.ColorMenuActivado Then
                ctlMenu.BackColor = m_ObjEntorno.ColorMenuDesactivado
            End If
        End If
        Exit Sub
    End If
    
    For i = 1 To NumeroDeMenus
        Set ctlInd = Nothing
        Set ctlInd = getControl(frm, "IND" & i)
        Set ctlMenu = Nothing
        Set ctlMenu = getControl(frm, "MENU" & i)
        
        If i = p_Numero Then
            If Not ctlInd Is Nothing Then
                If ctlInd.Visible = False Then
                    ctlInd.Visible = True
                    ctlMenu.BackColor = m_ObjEntorno.ColorMenuActivado
                End If
            End If
        Else
            If Not ctlInd Is Nothing Then
                If ctlInd.Visible = True Then
                    ctlInd.Visible = False
                    If Not ctlMenu Is Nothing Then
                        ctlMenu.BackColor = m_ObjEntorno.ColorMenuDesactivado
                        If TieneSubMenus = EnumSino.Sí Then
                            
                            If i = OrdinalMenuDesplegable Then
                                OcultarSubMenus
                            End If
                        End If
                    End If
                End If
            End If
            
        End If

    Next
   
End Sub
Private Function getControl(frm As Form, p_NombreControl As String) As Control
    
    On Error Resume Next
    Set getControl = frm.Controls(p_NombreControl)
    
    
End Function

Public Sub ActivarSubMenu(p_Numero As Single)
   
    Dim ctlSubMenu As Control
    Dim i As Single
    
    
    If TieneSubMenus = EnumSino.No Then
        Exit Sub
    End If
    
    For i = 1 To NumeroDeSubMenus
        Set ctlSubMenu = Nothing
        Set ctlSubMenu = getControl(frm, "SUBMENU" & i)
        
        If i = p_Numero Then
            If Not ctlSubMenu Is Nothing Then
                If ctlSubMenu.BackColor <> m_ObjEntorno.ColorMenuActivado Then
                    ctlSubMenu.BackColor = m_ObjEntorno.ColorMenuActivado
                End If
            End If
            
        Else
            If Not ctlSubMenu Is Nothing Then
                If ctlSubMenu.BackColor <> m_ObjEntorno.ColorMenuDesactivado Then
                    ctlSubMenu.BackColor = m_ObjEntorno.ColorMenuDesactivado
                End If
            End If
        End If

    Next
   
End Sub

Public Sub RatonSobrePANEL1()
    Dim i As Single
    Dim ctl As Control
    
    
    If TieneSubMenus = EnumSino.Sí Then
        
        Set ctl = Nothing
        Set ctl = getControl(frm, "Menu" & OrdinalMenuDesplegable)
        If Not ctl Is Nothing Then
            If ctl.BackColor = m_ObjEntorno.ColorMenuActivado Then
                Exit Sub
            End If
        End If
    End If
    
    For i = 1 To NumeroDeMenus
        ActDesActMenu i, EnumSino.No
    Next
    
    
End Sub

Public Sub RatonSobreDetalle()
    Dim i As Single
    Dim ctl As Control
    
    
    For i = 1 To NumeroDeMenus
        ActDesActMenu i, EnumSino.No
    Next
    
    If TieneSubMenus = EnumSino.Sí Then
        Set ctl = getControl(frm, "Desplegar")
        
        If Not ctl Is Nothing Then
            If ctl.Visible = False Then
                OcultarSubMenus
            End If
        End If
        
    End If
End Sub
Public Sub EstablecerDatosIniciales(p_frm As Form, m_Error As String)
    
    Dim ctl As Control
    Dim ctlMenu As Control
    Dim ctlInd As Control
    
    Dim i As Single
   
    Set Me.frm = p_frm
    
    For Each ctl In frm.Controls
        If ctl.ControlType = acLabel Then
            If Left(ctl.Name, 7) = "lblMENU" Then
                NumeroDeMenus = NumeroDeMenus + 1
            End If
            If Left(ctl.Name, 10) = "lblSUBMENU" Then
                NumeroDeSubMenus = NumeroDeSubMenus + 1
                If InStr(1, ctl.Tag, "MENU") <> 0 And OrdinalMenuDesplegable = 0 Then
                    dato = Split(ctl.Tag, "MENU")
                    OrdinalMenuDesplegable = CSng(dato(1))
                End If
            End If
        End If
    Next
    If NumeroDeSubMenus = 0 Then
        TieneSubMenus = EnumSino.No
    Else
        TieneSubMenus = EnumSino.Sí
    End If
    Set ctl = getControl(frm, "PANEL1")
    If Not ctl Is Nothing Then
        ctl.BackColor = m_ObjEntorno.ColorMenuDesactivado
    End If
    For i = 1 To NumeroDeMenus
        Set ctlMenu = Nothing
        Set ctlMenu = getControl(frm, "MENU" & i)
        If Not ctlMenu Is Nothing Then
            ctlMenu.BackColor = m_ObjEntorno.ColorMenuDesactivado
        End If
        Set ctlInd = Nothing
        Set ctlInd = getControl(frm, "IND" & i)
        If Not ctlInd Is Nothing Then
            ctlInd.BackColor = RGB(255, 255, 255)
            ctlInd.Visible = False
        End If
    Next
    If TieneSubMenus = EnumSino.Sí Then
        OcultarSubMenus
    End If
End Sub


