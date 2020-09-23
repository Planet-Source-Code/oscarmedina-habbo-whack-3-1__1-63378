VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create/Edit Furni"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   Icon            =   "furni.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox send 
      Height          =   285
      Left            =   840
      TabIndex        =   18
      Top             =   3720
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   17
      Top             =   3120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ListBox Variables 
      Appearance      =   0  'Flat
      Height          =   1590
      ItemData        =   "furni.frx":058A
      Left            =   4800
      List            =   "furni.frx":058C
      TabIndex        =   16
      Top             =   0
      Width           =   2055
   End
   Begin VB.TextBox color 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      TabIndex        =   15
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      Caption         =   "Create / Modify Furni"
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox R 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   13
      Text            =   "0"
      Top             =   1920
      Width           =   255
   End
   Begin VB.TextBox W 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   6
      Text            =   "1"
      Top             =   1440
      Width           =   255
   End
   Begin VB.TextBox L 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Text            =   "1"
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox Z 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4200
      TabIndex        =   4
      Text            =   "0.0"
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox Y 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      Text            =   "5"
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox X 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3240
      TabIndex        =   2
      Text            =   "5"
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox sprite 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox ID 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "0"
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label cmdfurnilist 
      BackStyle       =   0  'Transparent
      Caption         =   "Click Here to see Furni List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4440
      TabIndex        =   19
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   1185
      Left            =   1440
      Picture         =   "furni.frx":058E
      Top             =   960
      Width           =   3300
   End
   Begin VB.Label Label6 
      Caption         =   "Rotation:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "X         Y          Z"
      Height          =   255
      Left            =   3285
      TabIndex        =   11
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Sprite Code"
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "ID"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Width:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Lenght:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   615
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public num As String
Public Letras As String
Public mVariables As Boolean

Private Sub cmdfurnilist_Click()
Form4.Show
End Sub

Private Sub Command1_Click()
Dim nVariables As Integer
Dim V As Integer

mVariables = False

send.Text = "A]" & ID.Text & "," & sprite.Text & " " & X.Text & " " & Y.Text & " " & L.Text & " " & W.Text & " " & R.Text & " " & Z.Text & " " & color.Text & "/"

nVariables = Variables.ListCount

If nVariables = "0" Then
send.Text = send.Text & ""
Else
For V = 0 To nVariables - 1
Text1.Text = V
Variables.ListIndex = Text1.Text
send.Text = send.Text & Variables.Text & "/"
Next V
send.Text = send.Text & ""
End If

frmMain.sckClient.SendData (send.Text)
Form3.Hide

mVariables = True
End Sub

Function FindAndReplace(InputString As String, FindString As String, ReplaceString As String) As String
Dim Found() As String
Found() = Split(InputString, FindString)
If UBound(Found) = 0 Then FindAndReplace = InputString: Exit Function
For I = 0 To UBound(Found)
FindAndReplace = FindAndReplace & Found(I) & ReplaceString
Next I
End Function

Public Function FindAndHighlight(txt1 As TextBox, SearchString As String, CaseSensitive As Boolean, Optional StartIndex As Integer)
Dim X As Integer
On Error GoTo err
Dim xSelStart As Integer
Dim xSelLength As Integer

If StartIndex <= 0 Then X = 1 Else X = StartIndex

    If CaseSensitive = True Then
        xSelStart = InStr(X, txt1.Text, SearchString) - 1
    Else
        xSelStart = InStr(X, LCase(txt1.Text), LCase(SearchString)) - 1
    End If

xSelLength = Len(SearchString)

txt1.SelStart = xSelStart
txt1.SelLength = xSelLength
err:
End Function
Public Function ReplaceAndHighLight(txt1 As TextBox, ReplaceWith As String)
Dim xSelStart As Integer
Dim xSelLength As Integer
On Error GoTo err

xSelStart = txt1.SelStart
xSelLength = Len(ReplaceWith)

txt1.SelText = ReplaceWith
txt1.SelStart = xSelStart
txt1.SelLength = xSelLength
err:
End Function
Public Sub ArreglarPackete()
Dim numero As Integer
num = Str$(Len(Z.Text))
numero = num
If numero > 3 Then
Letras = Right$(Z.Text, Len(Z.Text) - 2)
Z.Text = FindAndReplace(Z.Text, Letras, "") + "0"
End If

num = Str$(Len(R.Text))
numero = num
If numero > 1 Then
Letras = Right$(R.Text, Len(R.Text) - 1)
R.Text = FindAndReplace(R.Text, Letras, "")
End If

FindChar (color.Text)
If nChars >= 1 Then
Letras = Right$(color.Text, Len(color.Text) - 2)
color.Text = Letras
End If
End Sub
Public Function FindChar(texto As String) As String
Dim Found() As String
Found() = Split(texto, ".")
nChars = UBound(Found)
End Function

Private Sub Form_Load()
mVariables = True
End Sub

Private Sub Variables_Click()
If mVariables = True Then
Dim nvo As String
nvo = InputBox("Modify: " & Variables.Text)
Variables.List(Variables.ListIndex) = nvo
End If
End Sub
