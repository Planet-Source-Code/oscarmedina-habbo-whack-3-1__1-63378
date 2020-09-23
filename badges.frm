VERSION 5.00
Begin VB.Form form8 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client Side Badges - Click A Badge"
   ClientHeight    =   660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4515
   Icon            =   "badges.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "badges.frx":058A
   ScaleHeight     =   660
   ScaleWidth      =   4515
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image UK2 
      Height          =   735
      Left            =   3960
      Top             =   0
      Width           =   615
   End
   Begin VB.Image UK1 
      Height          =   735
      Left            =   3345
      Top             =   0
      Width           =   615
   End
   Begin VB.Image VIP 
      Height          =   735
      Left            =   2760
      Top             =   0
      Width           =   495
   End
   Begin VB.Image ADM 
      Height          =   735
      Left            =   2230
      Top             =   0
      Width           =   500
   End
   Begin VB.Image HBA 
      Height          =   735
      Left            =   1680
      Top             =   0
      Width           =   525
   End
   Begin VB.Image NWB 
      Height          =   735
      Left            =   1150
      Top             =   0
      Width           =   525
   End
   Begin VB.Image HC2 
      Height          =   735
      Left            =   580
      Top             =   0
      Width           =   525
   End
   Begin VB.Image HC1 
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   530
   End
End
Attribute VB_Name = "form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ADM_Click()
frmMain.sckClient.SendData "CeIADM"
form8.Hide
End Sub

Private Sub HBA_Click()
frmMain.sckClient.SendData "CeIHBA"
form8.Hide
End Sub

Private Sub HC1_Click()
frmMain.sckClient.SendData "CeIHC1"
form8.Hide
End Sub

Private Sub HC2_Click()
frmMain.sckClient.SendData "CeIHC2"
form8.Hide
End Sub

Private Sub NWB_Click()
frmMain.sckClient.SendData "CeINWB"
form8.Hide
End Sub

Private Sub UK1_Click()
frmMain.sckClient.SendData "CeIUK1"
form8.Hide
End Sub

Private Sub UK2_Click()
frmMain.sckClient.SendData "CeIUK2"
form8.Hide
End Sub

Private Sub VIP_Click()
frmMain.sckClient.SendData "CeIVIP"
form8.Hide
End Sub
