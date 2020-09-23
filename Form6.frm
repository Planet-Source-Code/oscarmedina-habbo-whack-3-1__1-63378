VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2895
   LinkTopic       =   "Form6"
   ScaleHeight     =   4545
   ScaleWidth      =   2895
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstnombres 
      Height          =   4545
      Left            =   1680
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.ListBox lstusers 
      Height          =   4545
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public remover As String
Public vEncontrado As Integer
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form6.Hide
End Sub
Public Function BuscarChars(texto As String, Buscar As String) As String
Dim Found() As String
Found() = Split(Buscar, texto)
vEncontrado = UBound(Found)
End Function
Public Function retirarusuario()
Dim id1 As Integer
Dim id2 As Integer

remover = Split(SckBuffer, "@]")(1)
remover = Split(remover, "")(0)

id1 = remover

nVariables = lstusers.ListCount

For V = 0 To nVariables - 1
lstusers.ListIndex = V

id2 = lstusers.Text
If id1 = id2 Then
lstusers.RemoveItem (V)
lstnombres.RemoveItem (V)
GoTo salir:
End If
Next V
salir:
End Function
Public Function separarid(packete As String)
Dim ID As String
Dim nombre As String
Dim Found() As String

ID = Split(packete, "i:")(1)
ID = Split(ID, "n:")(0)

nombre = Split(packete, "n:")(1)
nombre = Split(nombre, "f:")(0)

Found() = Split(nombre, "")
nChars = UBound(Found)

If nChars > 0 Then
nombre = Split(nombre, "")(1)
lstnombres.AddItem ("(Pet)" & nombre)
Else
lstnombres.AddItem (nombre)
End If

lstusers.AddItem (ID)

End Function
Public Function ChecarNombre(UsersID As Integer)
Dim id1 As Integer
Dim id2 As Integer
Dim HabboName As String

id1 = UsersID

nVariables = lstusers.ListCount

For V = 0 To nVariables - 1
lstusers.ListIndex = V

id2 = lstusers.Text
If id1 = id2 Then
lstnombres.ListIndex = V
HabboName = lstnombres.Text
GoTo salir:
End If
Next V
HabboName = "not found"
salir:
Call ModSockFuncts.submitname(HabboName)
End Function
