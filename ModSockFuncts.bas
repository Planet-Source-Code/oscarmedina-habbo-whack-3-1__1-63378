Attribute VB_Name = "ModSockFuncts"
Public SckBuffer As String
Public comienzo As String
Public Coordenadas As String
Public Indentify As String
Public Packet As String
Public nChars As Integer
Public Element As String
Public AccountName As String
Public NombredeHabbo As String
Public credits As String
Public Email As String
Public LastAccess As String
Public Access As String
Public Letra1 As String
Public Letra2 As String
Public Users As String
Public Figure As String
Public furni As String
Public Mission As String
Public PHTickets As Integer
Public Birthday As String
Public AccessCount As Integer
Public Film As Integer
Public sdata As String
Public Client As Boolean
Public WaveHabbo As Boolean
Public FlickerHabbo As Boolean
Public DanceHabbo As Boolean
Public Tile As String
Public Chat As String
Public PeopleRights As String
Public DetermineRights As String
Public ServerPort As Single
Public Variable As String
Public ServerHost As String
Public ClientHost As String
Public ClientPort As Single
Public UnloadPage As Boolean
Public ID As Integer
'This is the sub used to send packets and recieve packets
Public Sub UpdateStatus(PacketStatus As String)
On Error Resume Next
If PacketStatus = "@@" Then
    frmMain.lblStatus = "Getting Data"
  'Message on Habbo Telling we are in
    frmMain.sckClient.SendData "@amod_warn/         Connection Established" & vbCrLf & "- OscarMedina [OscarMedina.Tk] -"
    DesblockearFunciones
ElseIf PacketStatus = "@A" Then
    frmMain.lblStatus = "Id Key recived"
    Packet = Split(SckBuffer, "@A")(1)
    Packet = Split(Packet, "")(0)
    frmMain.lblClientID.Text = Packet
ElseIf PacketStatus = "Ce" Then
    Packet = Split(SckBuffer, "@F")(1)
    Packet = Split(Packet, ".0")(0)
    credits = Packet
    frmMain.lblcredits.Caption = credits
ElseIf PacketStatus = "@E" Then
    frmMain.lblStatus = "Writing Info."
    Packet = Split(SckBuffer, "name=")(1)
    Packet = Split(Packet, "email=")(0)
    Packet = Split(Packet, Chr(13))(0)
    AccountName = Packet
    frmMain.lblName.Caption = AccountName
    
    Packet = Split(SckBuffer, "email=")(1)
    Packet = Split(Packet, "figure=")(0)
    Email = Packet
    frmMain.lblEmail.Caption = Email
    
    Packet = Split(SckBuffer, "figure=")(1)
    Packet = Split(Packet, "sex=")(0)
    Figure = Packet
    frmMain.lblFigureNum.Text = Figure
    
    Packet = Split(SckBuffer, "sex=")(1)
    Packet = Split(Packet, "customData=")(0)
    sex = Packet
    frmMain.lblsex.Caption = sex
    
    If frmMain.lblsex.Caption = "m" Then
    frmMain.lblsex.Caption = "Male"
    End If
    If frmMain.lblsex.Caption = "f" Then
    frmMain.lblsex.Caption = "Female"
    End If
    
    Packet = Split(SckBuffer, "customData=")(1)
    Packet = Split(Packet, "ph_tickets=")(0)
    Mission = Packet
    frmMain.lblmission.Caption = Mission
    
    Packet = Split(SckBuffer, "ph_tickets=")(1)
    Packet = Split(Packet, "ph_figure=")(0)
    PHTickets = Packet
    frmMain.lbltickets.Caption = PHTickets
    
    Packet = Split(SckBuffer, "birthday=")(1)
    Packet = Split(Packet, "photo_film=")(0)
    Birthday = Packet
    frmMain.lblBirth.Caption = Birthday
    
    Packet = Split(SckBuffer, "photo_film=")(1)
    Packet = Split(Packet, "directMail=")(0)
    Film = Packet
    frmMain.lblfilm.Caption = Film
    frmMain.lblStatus.Caption = "[Connected]"
ElseIf PacketStatus = "@_" Then
Form6.lstusers.Clear
Form6.lstnombres.Clear
Call UsuariosA
Call Furnilist

ElseIf PacketStatus = "A^" Then
Dim Found() As String
Found() = Split(SckBuffer, "@]")
nChars = UBound(Found)

If nChars > 0 Then
Call Form6.retirarusuario
End If

ElseIf PacketStatus = "Ab" Then
Found() = Split(SckBuffer, "@]")
nChars = UBound(Found)

If nChars > 0 Then
Call Form6.retirarusuario
End If

ElseIf PacketStatus = "@]" Then
Call Form6.retirarusuario
End If

End Sub
Public Sub GetTile()
On Error Resume Next
If Left(SckBuffer, 2) = "@b" Then
    Tile = Split(SckBuffer, "@b")(1)
    AccountName = Split(Tile, " ")(0)
    Tile = Split(Tile, " ")(1)
    Tile = Split(Tile, "/")(0)
    frmMain.lblTile.Caption = Tile
End If
End Sub
Public Sub GetChat(PacketStatus As String)
'Checks for chat, i spent lots of time making this and it eint compleate yet...
On Error Resume Next
If PacketStatus = "@Y" Then

If frmMain.prueba.Text = "0" Then
frmMain.prueba.Text = "1"
Else
frmMain.prueba.Text = "0"
GoTo Sig:
End If
    
    Packet = Split(SckBuffer, "@Y")(1)
    
    Letra1 = Left(Packet, 1)
    Packet = Split(SckBuffer, Letra1)(1)
    
    Packet = Mid(Packet, 1, InStr(1, Packet, "") - 2)
    
ChecaPakete
Form6.ChecarNombre (ID)
    
    frmMain.lstChatLog.AddItem (NombredeHabbo & " Whispered: " & Packet)
    
Sig:
ElseIf PacketStatus = "@Z" Then
    Packet = Split(SckBuffer, "@Z")(1)
    
    Letra1 = Left(Packet, 1)
    Packet = Split(SckBuffer, Letra1)(1)
    
    Packet = Mid(Packet, 1, InStr(1, Packet, "") - 2)
    ID = ""
    
ChecaPakete
Form6.ChecarNombre (ID)
    
    frmMain.lstChatLog.AddItem (NombredeHabbo & " Shouted: " & Packet)
ElseIf PacketStatus = "@X" Then
    Packet = Split(SckBuffer, "@X")(1)
   
    Letra1 = Left(Packet, 1)
    Packet = Split(SckBuffer, Letra1)(1)
    
    Packet = Mid(Packet, 1, InStr(1, Packet, "") - 2)
    ID = ""
    
ChecaPakete
Form6.ChecarNombre (ID)

    frmMain.lstChatLog.AddItem (NombredeHabbo & " Said: " & Packet)

ElseIf PacketStatus = "@\" Then
Call Userslist
End If
End Sub

Public Sub DesblockearFunciones()
frmMain.cmdShowLogs.Enabled = True
frmMain.cmdShowRoom.Enabled = True
frmMain.imgLogo.Enabled = True
frmMain.Image4.Enabled = True
frmMain.Label7.Enabled = True
End Sub

Public Sub BlockearFunciones()
frmMain.cmdShowLogs.Enabled = False
frmMain.cmdShowRoom.Enabled = False
frmMain.imgLogo.Enabled = False
frmMain.Image4.Enabled = False
frmMain.Label7.Enabled = False
End Sub
Public Sub ChecaPakete()
Dim RID As Boolean
Letra2 = Left(Packet, 1)
RID = False

Select Case Letra2
Case "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
RID = True
End Select

If RID = True Then
IDS = Letra1 & Letra2
Packet = Split(Packet, Letra2)(1)
Else
IDS = Letra1
End If

Form7.txtcodificada.Text = IDS
Call Form7.EmpezarCodificacion
ID = Form7.txtID.Text
End Sub

Public Function submitname(UserName As String)
If UserName = "not found" Then
IDS = Letra1

Form7.txtcodificada.Text = IDS
Call Form7.EmpezarCodificacion
ID = Form7.txtID.Text
Form6.ChecarNombre (ID)
Packet = Letra2 & Packet
GoTo salir
End If
NombredeHabbo = UserName
salir:
End Function

Public Sub Furnilist()
Letra = "FurniList"
Form4.lstfurni.Clear
On Error Resume Next
    Packet = Split(SckBuffer, "@`")(1)
    Packet = Split(Packet, "")(0)
    furni = Packet
Do Until Letra = ""
    Letra = Left(furni, 1)
    
    If Letra = Chr(13) Then
    furni = Right$(furni, Len(furni) - 1)
    Letra = Left(furni, 1)
    End If
    
    Element = Split(furni, Letra)(1)
    Element = Split(furni, Chr(13))(0)

    Form4.lstfurni.AddItem (Element)
    furni = Form4.FindAndReplace(furni, Element, "")
Loop
End Sub

Public Sub modifyfurni(furni2 As String)
On Error Resume Next
Dim Letra1 As String
Dim Letra2 As String
Dim Variables As String

Form3.Variables.Clear
    
    Letra1 = Right$(furni2, Len(furni2) - 1)
    Form3.ID.Text = Split(furni2, Letra1)(1)
    Form3.ID.Text = Split(furni2, ",")(0)
    
    furni2 = Form3.FindAndReplace(furni2, Form3.ID.Text + ",", "")

    Letra1 = Right$(furni2, Len(furni2) - 1)
    Form3.sprite.Text = Split(furni2, Letra1)(1)
    Form3.sprite.Text = Split(furni2, " ")(0)
    
    furni2 = Form3.FindAndReplace(furni2, Form3.sprite.Text + " ", "")
    
    'Obtener Coordenadas
    Letra1 = Right$(furni2, Len(furni2) - 1)
    Coordenadas = Split(furni2, Letra1)(1)
    Coordenadas = Split(furni2, "/")(0)
    'Coordenadas + Color
    
    Letra1 = Right$(furni2, Len(furni2) - 1)
    Form3.X.Text = Split(furni2, Letra1)(1)
    Form3.X.Text = Split(furni2, " ")(0)
    
    Form3.Text1.Text = furni2
    Form3.FindAndHighlight Form3.Text1, Form3.X.Text + " ", False
    Form3.ReplaceAndHighLight Form3.Text1, ""
    
    furni2 = Form3.Text1.Text
    
    Letra1 = Right$(furni2, Len(furni2) - 1)
    Form3.Y.Text = Split(furni2, Letra1)(1)
    Form3.Y.Text = Split(furni2, " ")(0)
    
    furni2 = Form3.FindAndReplace(furni2, Form3.Y.Text + " ", "")
    
    Letra1 = Right$(furni2, Len(furni2) - 1)
    Form3.L.Text = Split(furni2, Letra1)(1)
    Form3.L.Text = Split(furni2, " ")(0)
    
    furni2 = Form3.FindAndReplace(furni2, Form3.L.Text + " ", "")
    
    Letra1 = Right$(furni2, Len(furni2) - 1)
    Form3.R.Text = Split(furni2, Letra1)(1)
    Form3.R.Text = Split(furni2, " ")(0)
    
    furni2 = Form3.FindAndReplace(furni2, Form3.R.Text + " ", "")
    
    Letra1 = Right$(furni2, Len(furni2) - 1)
    Form3.Z.Text = Split(furni2, Letra1)(1)
    Form3.Z.Text = Split(furni2, " ")(0)
    
    furni2 = Form3.FindAndReplace(furni2, Form3.Z.Text + " ", "")
    
    Letra1 = Right$(furni2, Len(furni2) - 1)
    Form3.color.Text = Split(furni2, Letra1)(1)
    Form3.color.Text = Split(furni2, "/")(0)
    
    furni2 = Form3.FindAndReplace(furni2, Form3.color.Text + "/", "")

    Letra1 = Right$(furni2, Len(furni2) - 1)
    Variables = Split(furni2, Letra1)(1)
    Variables = Split(furni2, "/")(0)
    Form3.Variables.AddItem (Variables)
    
    Letra2 = Left$(furni2, Len(furni2) - 1)
    
    FindChar (furni2)
    
    If nChars = "0" Then
    furni2 = Form3.FindAndReplace(furni2, Variables, "")
    Else
    furni2 = Form3.FindAndReplace(furni2, Variables + "/", "")
    End If

If Not furni2 = "" Then
    FindChar (furni2)
If Not nChars = "0" Then
Do Until furni2 = ""
    Letra1 = Right$(furni2, Len(furni2) - 1)
    Variables = Split(furni2, Letra1)(1)
    Letra2 = Left$(furni2, Len(furni2) - 1)
    Variables = Split(furni2, "/")(0)
    
    Form3.Variables.AddItem (Variables)
    
    FindChar (furni2)
If Not nChars = "0" Then
    furni2 = Form3.FindAndReplace(furni2, Variables + "/", "")
Else
    furni2 = Form3.FindAndReplace(furni2, Variables, "")
End If
Loop


Else
    
    
Do Until furni2 = ""
    Letra1 = Right$(furni2, Len(furni2) - 1)
    Variable = Split(furni2, Letra1)(1)
    Letra2 = Left$(furni2, Len(furni2) - 1)
    If Letra2 = "" Then
    Form3.Variables.AddItem (furni2)
    furni2 = ""
    Else
    Form3.Variables.AddItem (Variable)
    furni2 = Form3.FindAndReplace(furni2, Variables, "")
    End If
    Loop
End If
End If
Form3.ArreglarPackete
Form3.Show

Coordenadas = Form3.FindAndReplace(Coordenadas, " " & Form3.color.Text, "")

    Letra1 = Right$(Coordenadas, Len(Coordenadas) - 1)
    Form3.X.Text = Split(Coordenadas, Letra1)(1)
    Form3.X.Text = Split(Coordenadas, " ")(0)
    
    Form3.Text1.Text = Coordenadas
    Form3.FindAndHighlight Form3.Text1, Form3.X.Text + " ", False
    Form3.ReplaceAndHighLight Form3.Text1, ""
    
    Coordenadas = Form3.Text1.Text
    
    Letra1 = Right$(Coordenadas, Len(Coordenadas) - 1)
    Form3.Y.Text = Split(Coordenadas, Letra1)(1)
    Form3.Y.Text = Split(Coordenadas, " ")(0)
    
    Coordenadas = Form3.FindAndReplace(Coordenadas, Form3.Y.Text + " ", "")
    
    Letra1 = Right$(Coordenadas, Len(Coordenadas) - 1)
    Form3.L.Text = Split(Coordenadas, Letra1)(1)
    Form3.L.Text = Split(Coordenadas, " ")(0)
    
    Form3.Text1.Text = Coordenadas
    Form3.FindAndHighlight Form3.Text1, Form3.L.Text + " ", False
    Form3.ReplaceAndHighLight Form3.Text1, ""
    
    Coordenadas = Form3.Text1.Text
    
    Letra1 = Right$(Coordenadas, Len(Coordenadas) - 1)
    Form3.W.Text = Split(Coordenadas, Letra1)(1)
    Form3.W.Text = Split(Coordenadas, " ")(0)
    
    Form3.Text1.Text = Coordenadas
    Form3.FindAndHighlight Form3.Text1, Form3.W.Text + " ", False
    Form3.ReplaceAndHighLight Form3.Text1, ""
    
    Coordenadas = Form3.Text1.Text
    
    Letra1 = Right$(Coordenadas, Len(Coordenadas) - 1)
    Form3.R.Text = Split(Coordenadas, Letra1)(1)
    Form3.R.Text = Split(Coordenadas, " ")(0)
    
    Coordenadas = Form3.FindAndReplace(Coordenadas, Form3.R.Text + " ", "")
    
    Form3.Z.Text = Coordenadas
        
    Coordenadas = Form3.FindAndReplace(Coordenadas, Form3.Z.Text, "")
End Sub
Public Function FindChar(texto As String) As String
Dim Found() As String
Found() = Split(texto, "/")
nChars = UBound(Found)
End Function
Public Sub Userslist()
Dim Borrar As String
Letra = "UsersList"
On Error Resume Next
    Packet = Split(SckBuffer, "@\")(1)
    Packet = Split(Packet, "c:")(0)
    Users = Packet
    Element = Split(Users, "i:")(1)
    Element = Split(Users, "f:")(0)

    Form6.separarid (Element & "f:")
    
    Users = ""
End Sub

Public Function UsuariosA()
Dim Borrar As String
Dim Final As String
Letra = "UsersList"
Element = "Elemento"
On Error Resume Next
    Packet = Split(SckBuffer, "@\")(1)
    Packet = Split(Packet, "")(0)
    Users = Packet & ""

If Not Users = "" Then
Do Until Letra = ""
    Letras = Left(Users, 2)
    Letra = Left(Users, 1)
    
    If Letras = "f:" Then
    Borrar = Split(Users, "f:")(1)
    Borrar = Split(Users, "i:")(0)
    EncontrarChars (Users)
    If nChars = 0 Then
    Users = ""
    GoTo salir:
    Else
    Users = Form4.FindAndReplace(Users, Borrar, "")
    End If
    End If
    
    If Not Element = "" Then
    Element = Split(Users, "i:")(1)
    Element = Split(Users, "f:")(0)

    Form6.separarid (Element & "f:")
    End If
    
    Users = Form4.FindAndReplace(Users, Element, "")
    
        
    Letra = Left(Users, 2)
Loop
End If
salir:
End Function
Public Function EncontrarChars(texto As String) As String
Dim Found() As String
Found() = Split(texto, "i:")
nChars = UBound(Found)
End Function
