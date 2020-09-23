VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send Packets Client/Server"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      Height          =   195
      Left            =   1900
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reset"
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
      Left            =   1900
      TabIndex        =   3
      Top             =   2520
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   195
      Left            =   3120
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Send to Server"
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
      Left            =   3120
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H000000FF&
      Height          =   195
      Left            =   240
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Send to Client"
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
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
frmMain.sckServer.SendData Text1.Text
End Sub

Private Sub Label2_Click()
Text1.Text = ""
End Sub

Private Sub Label4_Click()
frmMain.sckClient.SendData Text1.Text
End Sub
