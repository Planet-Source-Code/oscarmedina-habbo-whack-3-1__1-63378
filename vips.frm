VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form5 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Habbo Whack Connection - OscarMedina"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4980
   Icon            =   "vips.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   4980
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   4320
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox url 
      Height          =   285
      Left            =   960
      TabIndex        =   8
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txthostS 
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Text            =   "<param name=""sw4"" value=""connection.mus.host="
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txthostC 
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Text            =   "<param name=""sw1"" value=""connection.info.host="
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtpuerto 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Text            =   "<param name=""sw2"" value=""connection.info.port="
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox antiresize 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   $"vips.frx":058A
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox lstServer 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      ItemData        =   "vips.frx":0635
      Left            =   1200
      List            =   "vips.frx":0645
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   240
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000C&
      Height          =   195
      Left            =   1470
      TabIndex        =   7
      Top             =   1350
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   255
      Left            =   1440
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Shape shpConnect 
      BorderColor     =   &H000000FF&
      Height          =   255
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label cmdConnect 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CONNECT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Habbo Whack Connection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub GuardarConexion(lugar As String, html As String)
Open lugar For Output As #1
Print #1, html
Close #1
End Sub

Private Sub cmdConnect_Click()
On Error Resume Next
Dim popup As String

If lstServer.Text = "" Then
MsgBox ("Select the Hotel to Connect(In Blue)")
GoTo salir:
Else
cmdConnect.Enabled = False
cmdConnect.Caption = "CONNECTING"

Form5.Height = "2115"

Label2.Visible = True

popup = Inet2.OpenURL("http://www.habbohotel.co.uk/habbo/en/")

Label2.Width = "500"

popup = Mid(popup, InStr(popup, "/habbo/en/common/habbo_client/client_popup/?t=") + Len("/habbo/en/common/habbo_client/client_popup/?t="))
popup = Mid(popup, 1, InStr(1, popup, " target=") - 2)

Label2.Width = "680"

url.Text = "http://www." & lstServer.Text & "common/habbo_client/client_popup/?t=" & popup

Call abrir
End If

salir:
End Sub

Private Sub abrir()
Dim codigo As String
hostS = txthostS.Text
hostC = txthostC.Text
puerto = txtpuerto.Text


codigo = Inet1.OpenURL(url.Text)

Label2.Width = "800"

ServerHost = Mid(codigo, InStr(codigo, hostS) + Len(hostS))
ServerHost = Mid(ServerHost, 1, InStr(1, ServerHost, ">") - 2)
frmMain.sckServer.RemoteHost = ServerHost

Label2.Width = "1000"

ClientHost = Mid(codigo, InStr(codigo, hostC) + Len(hostC))
ClientHost = Mid(ClientHost, 1, InStr(1, ClientHost, ">") - 2)
frmMain.sckClient.RemoteHost = ClientHost

Label2.Width = "1200"

conectport = Mid(codigo, InStr(codigo, puerto) + Len(puerto))
conectport = Mid(conectport, 1, InStr(1, conectport, ">") - 2)

Label2.Width = "1500"

frmMain.sckServer.RemotePort = conectport
frmMain.sckClient.RemotePort = conectport
frmMain.sckClient.LocalPort = conectport

Label2.Width = "1800"

codigo = Replace(codigo, "connection.info.host=" & ClientHost, "connection.info.host=127.0.0.1")
codigo = Replace(codigo, antiresize.Text, "")

Label2.Width = "2000"

GuardarConexion "C:\tmp.html", codigo
frmMain.txtnavegar.Text = "C:\tmp.html"

Label2.Width = "2120"

frmMain.Show
Unload Form5

End Sub
