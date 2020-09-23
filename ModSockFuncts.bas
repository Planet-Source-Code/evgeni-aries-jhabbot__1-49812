Attribute VB_Name = "ModSockFuncts"
Public SckBuffer As String 'the packet is put in to the sckbuff string
Public Indentify As String 'used for the 2 charachters of the packet to determine
Public Packet As String 'real packet
'next variables speak for themselves
Public AccountName As String '
Public Email As String
Public LastAccess As String
Public Access As String
Public Figure As String
Public IP As String
Public PHTickets As Integer
Public Birthday As String
Public AccessCount As Integer
Public Film As Integer
Public sdata As String 'client data
Public Client As Boolean ' check if client send
Public WaveHabbo As Boolean
Public FlickerHabbo As Boolean
Public DanceHabbo As Boolean
Public Ping As Boolean
Public Tile As String
Public PeopleRights As String
Public DetermineRights As String
Public ServerPort As Single
Public ServerHost As String
Public ClientHost As String
Public ClientPort As Single
Public UnloadPage As Boolean
'Drinks
Public Cam As Boolean
Public Carry As String
Public EndPacket As String
'Hobbbas and people
Public Hobbas(1 To 26) As String 'there can only be 26 hobbas
Public Hobba As Boolean
Public People(1 To 26) As String 'there can only be 25 people and 1 hobba
Public Rights As Boolean
'this is the sub used to send packets and recieve packets
Public Sub UpdateStatus(PacketStatus As String)
On Error Resume Next 'if error like socket not connected then just go on without outputing and closing program
' this sub updates and outputs the info
If PacketStatus = "@@#" Then 'special defined value tthat is recieved firsdt when connected
    frmMain.lblStatus = "Getting data"
ElseIf PacketStatus = "@A" Then 'all values go one by one till the habbohotel.dcr is loaded and ready to load rooms
    frmMain.lblStatus = "Key ID Recieved."
    Packet = Split(SckBuffer, "@A")(1)
    Packet = Split(Packet, "#")(0)
    frmMain.lblClientID.Text = Packet 'its not client id just the key was to lazy to fix it
ElseIf PacketStatus = "@B" Then
    'check the middle if its "can" this means that the user is allowed else banned
    Packet = Mid$(SckBuffer, 11, 9) 'the code is buggy because if you are banned it still
    If Packet = "can_trade" Then 'shows yes and then disconnects
        frmMain.lblBan.Caption = "No"
    Else
        frmMain.lblBan.Caption = "Yes"
    End If
frmMain.lblStatus = "Checks if banned."
ElseIf PacketStatus = "@E" Then ' this is the packet with all the basic info
    frmMain.lblStatus = "Getting Info."
    Packet = Split(SckBuffer, "name=")(1) '(name)email=
    Packet = Split(Packet, "email=")(0)
    Packet = Split(Packet, Chr(13))(0) 'now we split email and we get just the name
    'put the specified packet which inour case is the name in the variable
    AccountName = Packet
    frmMain.lblName.Caption = AccountName 'show name output
    Packet = Split(SckBuffer, "email=")(1) 'now we get (email)figure=1212121212
    Packet = Split(Packet, "figure=")(0) 'now we split this packet and just get the email
    Email = Packet 'put packet into email
    frmMain.lblEmail.Caption = Email 'show output
    Packet = Split(SckBuffer, "figure=")(1) 'figure=12121phonenumber= +44
    Packet = Split(Packet, "phoneNumber=")(0) 'split it we get figure
    Figure = Packet 'put packet into the variable
    frmMain.lblFigureNum.Text = Figure 'show out put
    Packet = Split(SckBuffer, "last_access_time=")(1) 'lass_access_time=212has_read
    Packet = Split(Packet, "has_read_agreement=")(0) 'split it but has and get acess
    Access = Packet 'put packet into access
    frmMain.lblLastAccess.Caption = Access ' show output
    Packet = Split(SckBuffer, "last_ip=")(1) 'last_ip=111.11.111.11ph_ticker=0
    Packet = Split(Packet, "ph_tickets=0")(0) 'split it by ph or if would like ph doesnt
    'make a diff effenct you could do that to all the others also and get the data
    IP = Packet 'put packet into variable
    frmMain.lblIP.Caption = IP 'show output
    Packet = Split(SckBuffer, "ph_tickets=")(1) '(tickets)][birthday=11.11.11
    Packet = Split(Packet, "birthday=")(0) ' split it and get phtickets
    PHTickets = Packet 'put tickets in a variable
    frmMain.lblTickets.Caption = PHTickets 'show out put
    Packet = Split(SckBuffer, "birthday=")(1) '(birth)directmail=
    Packet = Split(Packet, "directMail=")(0) 'split [(birth)][directmail]
    Birthday = Packet 'put packet into variable
    frmMain.lblBirth.Caption = Birthday 'show output
    Packet = Split(SckBuffer, "access_count=")(1) '(access)has_special
    Packet = Split(Packet, "has_special_rights=")(0) '[(access)][has_special]
    AccessCount = Packet 'put packt into variable
    frmMain.lblAccess.Caption = AccessCount 'show output
    Packet = Split(SckBuffer, "photo_film=")(1) '(film)#
    Packet = Split(Packet, "#")(0) '[(film)][#]
    Film = Packet 'put packet into a variable
    frmMain.lblFilm.Caption = Film 'show output
ElseIf PacketStatus = "@F" Then 'this packet loads the credits this means its loaded
    frmMain.lblStatus.Caption = "[Connected]" 'show output
End If
End Sub
Public Sub Wave()
On Error Resume Next 'just incase if error like socket aint connected it wont close
frmMain.sckserver.SendData "€Þ€€€" ' this is the value for wave
End Sub
Public Sub Flicker()
On Error Resume Next
frmMain.sckserver.SendData "€À€€€" 'value for flickr
'you cant recieve these values threw winsock packet editor you must
'create liek a textbox and put the data inside there
'i had this feature in dhabbox another program made by me but like had bad layout
'so i thougth of jhabbot and know since i got to bored i just didnt wont to implent
'anythign new or old
End Sub
Public Sub Dance()
On Error Resume Next
frmMain.sckserver.SendData "€]€€€" 'dance code
End Sub
Public Sub GetTile()
On Error Resume Next
If Left(SckBuffer, 2) = "@b" Then 'since @b is the status data
    Tile = Split(SckBuffer, "@b")(1) 'this removes @b
    AccountName = Split(Tile, " ")(0) 'this splits it for a accoutname
    If AccountName = frmMain.lblName.Caption Then 'if the account data is your name
        Tile = Split(Tile, " ")(1) 'then get the tile data namehere 1,1,1,1/sit.ect
        Tile = Split(Tile, "/")(0) 'tile is now equal to 1,1,1,1/sit.ect and now split it /
        frmMain.lblTile.Caption = Tile 'tile equal 1,1,1,1,1
'Useful comment:1,1 of the 2 values are the moving values of the tiles
'Useful comment:if you actually go on habbo you may notice that your figure changes positions on a tile
'Useful comment: this are the 3 last 1,1,1 values
    End If
End If
End Sub
Public Sub GetPeopleWithRights()
On Error Resume Next
PeopleRights = Split(SckBuffer, "@b")(1) 'like i said before this is the status report
If PeopleRights = SckBuffer Then 'status reports have what the user is doing and if the person
    Exit Sub 'isa admin and by doing that it there codes liek flatctrl admin flatctrl user and frlctrl furniture(hobbas rights)
End If 'so since all start with flatctrl made the spliting code easier
Packet = Split(PeopleRights, "flatctrl")(0) 'name 1,1,1,1,1 /sit/flatctrl
If Packet = PeopleRights Then 'we split it and get the name by the space
    Exit Sub
End If
PeopleRights = Split(PeopleRights, " ")(0) 'and to be sure now split it again
If Left(PeopleRights, 1) = "@" Then 'finally if there a @ character that means something is wrong and
    Exit Sub 'exit subs
Else
    CancelOutRights 'otherwise now that we got this lets canceloutrights
End If
End Sub
Public Sub CancelOutRights()
    For i = 1 To 25 '25 people in a room
        If People(i) = PeopleRights Then ' i as in the index of the people array
        Rights = True 'if people(index) string = to peoplerights this means its all
        'ready been modified and the same code would be easier to exit sub but this
        'gave me some buggy problems so i just put rights as a boolen
        ElseIf People(i) <> PeopleRights And People(i) = Empty Then
            'if index not the same as peoplerights and the index is empty in other words not modified then
            If Rights = False Then 'check if rights = true see above code
            People(i) = PeopleRights 'if rights = false then people index = peoplerights
            frmMain.lstPeopleRights.AddItem PeopleRights 'add the rights now
            Rights = True 'since we modified the index now we can only add one name
            'per finished loop so we put rights = true
            End If
        End If
    Next i
Rights = False 'now at the end so that the loop can go on next time peoplerights was hit
'we put rights = false
End Sub
Public Sub HobbaWave()
If Indentify = "@b" And frmMain.chkHobWaveflt.Value = Checked Then 'check if filter loaded
Dim Badge As String
    Badge = Split(SckBuffer, "/")(2) 'name 1,1,1,1,1/sit/mod 1,2,3.ect
    If Left(Badge, 5) = "mod 1" Then 'well thisi s pretty easy to understand
    'calculate the 5 spaces of left and you will see what it is
        'if mod1 then considired of being a hobba
        Badge = Split(SckBuffer, " ")(0) 'now we get the name of the hobba
        HobbaCancelOut 'cancel out samething as peoplerights
        Badge = Split(SckBuffer, "mod 1")(0) 'split mod1 from the status report
        frmMain.sckclient.SendData Badge & "wave/#" 'then add wave which will make
        'the hobba wave
        'same thing goes for the other commands
    ElseIf Left(Badge, 5) = "mod 2" Then
        Badge = Split(SckBuffer, " ")(0)
        HobbaCancelOut
        Badge = Split(SckBuffer, "mod 2")(0)
        frmMain.sckclient.SendData Badge & "wave/#"
    ElseIf Left(Badge, 5) = "mod A" Then
        Badge = Split(SckBuffer, " ")(0)
        HobbaCancelOut
        Badge = Split(SckBuffer, "mod A")(0)
        frmMain.sckclient.SendData Badge & "wave/#"
    End If
End If
End Sub
Public Sub HobbaCancelOut()
    For i = 1 To 26 '26 hobbas in a room
        If Hobbas(i) = Badge Then ' i as in the index of the hobbas array
        Hobba = True 'if hobbas(index) string = to badge this means its all
        'ready been modified and the same code would be easier to exit sub but this
        'gave me some buggy problems so i just put hobbas as a boolean
        ElseIf Hobbas(i) <> Badge And Hobbas(i) = Empty Then
            'if index not the same as badge and the index is empty in other words not modified then
            If Hobba = False Then 'check if hobba = true see above code
            Hobbas(i) = Badge 'if hobba = false then hobbas index = peoplerights
            frmMain.lstHobba.AddItem Badge 'add the badge name now
            Hobba = True 'since we modified the index now we can only add one name
            'per finished loop so we put hobba = true
            End If
        End If
    Next i
Hobba = False 'now at the end so that the loop can go on next time badge was hit
'we put hobba = false
End Sub
Public Sub GetCommands()
On Error Resume Next
'there are few ways to talk on habbo shout say whisper but it doesnt matter
'the packets are usualy @Z or @X then name then what was said
'lets begin splitting the packet
'the packet is @X name:fdfsfdfdsfdsf#
Packet = Split(SckBuffer, AccountName)(1) 'now its done to this :sdsdsd#
Packet = Split(Packet, "#")(0) 'now we remove #
Packet = Split(Packet, ":")(1) 'now we remove :
If Packet = "Panel" Or Packet = "panel" Then 'check for the keywords for the specified function
    frmMain.lstHobbas.Visible = True
ElseIf Packet = "HidePanel" Or Packet = "hidepanel" Then
    frmMain.lstHobbas.Visible = False
End If
End Sub
