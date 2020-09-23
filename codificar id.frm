VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2415
   LinkTopic       =   "Form7"
   ScaleHeight     =   885
   ScaleWidth      =   2415
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtID 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtcodificada 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public dec2letra As Integer
Public numerodecimal As Integer
Public nElemento As Integer
Public texto As String
Public decodificada As String

Public Sub decimalizar(Text As String)
Dim TextStart As Single

TextStart = 1

Select Case (Mid(decodificada, TextStart, 1))
Case "A"
setcapt "65"
Case "B"
setcapt "66"
Case "C"
setcapt "67"
Case "D"
setcapt "68"
Case "E"
setcapt "69"
Case "F"
setcapt "70"
Case "G"
setcapt "71"
Case "H"
setcapt "72"
Case "I"
setcapt "73"
Case "J"
setcapt "74"
Case "K"
setcapt "75"
Case "L"
setcapt "76"
Case "M"
setcapt "77"
Case "N"
setcapt "78"
Case "O"
setcapt "79"
Case "P"
setcapt "80"
Case "Q"
setcapt "81"
Case "R"
setcapt "82"
Case "S"
setcapt "83"
Case "T"
setcapt "84"
Case "U"
setcapt "85"
Case "V"
setcapt "86"
Case "W"
setcapt "87"
Case "X"
setcapt "88"
Case "Y"
setcapt "89"
Case "Z"
setcapt "90"
End Select
End Sub
Sub setcapt(Text As String)
numerodecimal = Text
End Sub

Public Sub EmpezarCodificacion()
If txtcodificada.Text = "H" Then
txtID.Text = "0"
ElseIf txtcodificada.Text = "I" Then
txtID.Text = "1"
ElseIf txtcodificada.Text = "J" Then
txtID.Text = "2"
ElseIf txtcodificada.Text = "K" Then
txtID.Text = "3"
Else
decodificada = txtcodificada.Text

texto = Left$(decodificada, Len(decodificada) - 1)

If Not texto = "" Then
Call sacarelemento(texto)
End If

decodificada = Replace(decodificada, texto, "")
Call decimalizar(decodificada)
dec2letra = numerodecimal
texto = dec2letra - 64
texto = texto * 4 + nElemento

txtID.Text = texto
End If
End Sub
Public Sub sacarelemento(Elemento As String)
Dim TextStart As Single

TextStart = 1

Select Case (Mid(txtcodificada.Text, TextStart, 1))
Case "P"
elem "0"
Case "Q"
elem "1"
Case "R"
elem "2"
Case "S"
elem "3"
End Select
End Sub
Sub elem(Elemento As String)
nElemento = Elemento
End Sub
