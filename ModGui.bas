Attribute VB_Name = "ModGui"
'basicly this module just hides shows unhover and hovers graphics
Public Sub HidePanel()
frmMain.shpPanel1.Width = 9
frmMain.shpImage.Visible = False
frmMain.imgLogo.Visible = False
frmMain.Line1.Visible = False
frmMain.shpMin.Visible = True
frmMain.shpMin1.Visible = True
frmMain.cmdShow.Visible = True
frmMain.shpMin3.Visible = True
frmMain.shpPanel2.Visible = False
frmMain.shpPanel3.Visible = False
frmMain.shpPanel4.Visible = False
frmMain.shpPanel5.Visible = False
frmMain.shpPanel6.Visible = False
frmMain.Shape10.Visible = False
frmMain.Shape9.Visible = False
frmMain.Shape12.Visible = False
End Sub
Public Sub ShowPanel()
frmMain.Shape9.Visible = True
frmMain.shpPanel6.Visible = True
frmMain.shpPanel1.Width = 177
frmMain.shpImage.Visible = True
frmMain.imgLogo.Visible = True
frmMain.Line1.Visible = True
frmMain.shpMin.Visible = True
frmMain.shpMin1.Visible = True
frmMain.cmdShow.Visible = False
frmMain.shpMin3.Visible = True
frmMain.shpPanel2.Visible = True
frmMain.shpPanel3.Visible = True
frmMain.shpPanel4.Visible = True
frmMain.shpPanel5.Visible = True
frmMain.Shape10.Visible = True
frmMain.Shape12.Visible = True
End Sub
Public Sub HoverLabelShow()
frmMain.cmdHide.ForeColor = vbGreen
End Sub
Public Sub UnHoverLabelShow()
frmMain.cmdHide.ForeColor = vbRed
End Sub
Public Sub HoverLabelHide()
frmMain.cmdShow.ForeColor = vbGreen
End Sub
Public Sub UnHoverLabelHide()
frmMain.cmdShow.ForeColor = vbRed
End Sub
Public Sub UnHoverAll()
UnHoverLabelShow
UnHoverLabelHide
End Sub
Public Function HoverCmd(cmd As Integer)
If cmd = 1 Then
UnhoverCmd4
UnHoverCmd3
frmMain.cmdShowLogs.ForeColor = vbGreen
UnHoverTitle
UnHoverCmd2
UnHoverLabelShow
ElseIf cmd = 2 Then
UnhoverCmd4
UnHoverCmd1
UnHoverCmd3
UnHoverLabelShow
frmMain.cmdShowRoom.ForeColor = vbGreen
UnHoverTitle
ElseIf cmd = 3 Then
UnhoverCmd4
UnHoverCmd1
UnHoverLabelShow
frmMain.cmdShowPublic.ForeColor = vbGreen
UnHoverTitle
UnHoverCmd2
ElseIf cmd = 4 Then
UnHoverCmd3
UnHoverCmd1
UnHoverLabelShow
frmMain.cmdShowBasic.ForeColor = vbGreen
UnHoverTitle
UnHoverCmd2
End If
End Function
Public Sub UnHoverCmd1()
frmMain.cmdShowLogs.ForeColor = vbRed
End Sub
Public Sub UnHoverCmd2()
frmMain.cmdShowRoom.ForeColor = vbRed
End Sub
Public Sub UnHoverCmd3()
frmMain.cmdShowPublic.ForeColor = vbRed
End Sub
Public Sub UnhoverCmd4()
frmMain.cmdShowBasic.ForeColor = vbRed
End Sub
Public Sub UnHoverAllCmd()
UnHoverCmd1
UnHoverCmd2
UnHoverCmd3
UnhoverCmd4
End Sub
Public Sub HoverTitle()
frmMain.Label2.ForeColor = vbGreen
End Sub
Public Sub UnHoverTitle()
frmMain.Label2.ForeColor = vbRed
End Sub
Public Sub HideInfo()
frmMain.lblName.Caption = ""
frmMain.lblEmail.Caption = ""
frmMain.lblFigureNum.Text = ""
frmMain.lblLastAccess.Caption = ""
frmMain.lblIP.Caption = ""
frmMain.lblFilm.Caption = ""
frmMain.lblTickets.Caption = ""
frmMain.lblBirth.Caption = ""
frmMain.lblAccess.Caption = ""
frmMain.lblBan.Caption = ""
frmMain.lblClientID.Text = ""
End Sub
