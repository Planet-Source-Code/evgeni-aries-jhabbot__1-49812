Attribute VB_Name = "ModMap"
Dim RoomModel As String
Public Sub GetModel(Class As String)
If Class = "@v" Then 'this is the indetifier of the room packet
    RoomModel = Split(SckBuffer, "t=")(1) 'after t= goes the info about the shap that it is
    If RoomModel = "model_e#" Then 'all next code is self explonitary
    frmMap.frmModelE.Visible = True
    frmMap.frmModelA.Visible = False
    frmMap.frmModelC.Visible = False
    frmMap.frmModelD.Visible = False
    frmMap.frmModelB.Visible = False
    frmMap.frmModelF.Visible = False
    ElseIf RoomModel = "model_b#" Then
    frmMap.frmModelE.Visible = False
    frmMap.frmModelA.Visible = False
    frmMap.frmModelC.Visible = False
    frmMap.frmModelD.Visible = False
    frmMap.frmModelB.Visible = True
    frmMap.frmModelF.Visible = False
    ElseIf RoomModel = "model_f#" Then
    frmMap.frmModelE.Visible = False
    frmMap.frmModelA.Visible = False
    frmMap.frmModelC.Visible = False
    frmMap.frmModelD.Visible = False
    frmMap.frmModelB.Visible = False
    frmMap.frmModelF.Visible = True
    ElseIf RoomModel = "model_a#" Then
    frmMap.frmModelE.Visible = False
    frmMap.frmModelA.Visible = True
    frmMap.frmModelC.Visible = False
    frmMap.frmModelD.Visible = False
    frmMap.frmModelB.Visible = False
    frmMap.frmModelF.Visible = False
    ElseIf RoomModel = "model_c#" Then
    frmMap.frmModelE.Visible = False
    frmMap.frmModelA.Visible = False
    frmMap.frmModelC.Visible = True
    frmMap.frmModelD.Visible = False
    frmMap.frmModelB.Visible = False
    frmMap.frmModelF.Visible = False
    ElseIf RoomModel = "model_d#" Then
    frmMap.frmModelE.Visible = False
    frmMap.frmModelA.Visible = False
    frmMap.frmModelC.Visible = False
    frmMap.frmModelD.Visible = True
    frmMap.frmModelB.Visible = False
    frmMap.frmModelF.Visible = False
    End If
End If
'when i reload the room i think if i remeber correctly instead of @v a AE appears
If Class = "AE" Then
    RoomModel = Split(SckBuffer, "AE")(1)
    RoomModel = Split(RoomModel, " ")(0)
    If RoomModel = "model_e" Then
    frmMap.frmModelE.Visible = True
    frmMap.frmModelA.Visible = False
    frmMap.frmModelC.Visible = False
    frmMap.frmModelD.Visible = False
    frmMap.frmModelB.Visible = False
    frmMap.frmModelF.Visible = False
    ElseIf RoomModel = "model_b" Then
    frmMap.frmModelE.Visible = False
    frmMap.frmModelA.Visible = False
    frmMap.frmModelC.Visible = False
    frmMap.frmModelD.Visible = False
    frmMap.frmModelB.Visible = True
    frmMap.frmModelF.Visible = False
    ElseIf RoomModel = "model_f" Then
    frmMap.frmModelE.Visible = False
    frmMap.frmModelA.Visible = False
    frmMap.frmModelC.Visible = False
    frmMap.frmModelD.Visible = False
    frmMap.frmModelB.Visible = False
    frmMap.frmModelF.Visible = True
    ElseIf RoomModel = "model_a" Then
    frmMap.frmModelE.Visible = False
    frmMap.frmModelA.Visible = True
    frmMap.frmModelC.Visible = False
    frmMap.frmModelD.Visible = False
    frmMap.frmModelB.Visible = False
    frmMap.frmModelF.Visible = False
    ElseIf RoomModel = "model_c" Then
    frmMap.frmModelE.Visible = False
    frmMap.frmModelA.Visible = False
    frmMap.frmModelC.Visible = True
    frmMap.frmModelD.Visible = False
    frmMap.frmModelB.Visible = False
    frmMap.frmModelF.Visible = False
    ElseIf RoomModel = "model_d" Then
    frmMap.frmModelE.Visible = False
    frmMap.frmModelA.Visible = False
    frmMap.frmModelC.Visible = False
    frmMap.frmModelD.Visible = True
    frmMap.frmModelB.Visible = False
    frmMap.frmModelF.Visible = False
    End If
End If
'=======================
'@R is the packet that means your were kicked
'iof kicked then room disapears if room disapears heightmap disapears
If Class = "@R" Then
    frmMap.frmModelF.Visible = False
    frmMap.frmModelB.Visible = False
    frmMap.frmModelE.Visible = False
    frmMain.lstHobbas.Visible = False
    For i = 1 To 25
    'and since we are kicked
    'you would like to refresh your settings of people and hobbas in the next room
    People(i) = Empty
    Hobbas(i) = Empty
    Next i
End If
If Class = "@_" Then
'@_ is the heightmap code but i use it to also tell me that the room is loaded
'since the room is loaded i will pop up the panel
    frmMain.lstHobbas.Visible = True
    frmMain.lstHobba.Clear
    frmMain.lstPeopleRights.Clear
    For i = 1 To 25
    'refresh
    People(i) = Empty
    Hobbas(i) = Empty
    Next i
End If
End Sub
