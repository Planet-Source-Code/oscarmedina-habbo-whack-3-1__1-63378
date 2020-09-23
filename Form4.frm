VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Furni Lister - Click the Furni to Modify"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9240
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstfurni 
      Height          =   4155
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lstfurni_Click()
modifyfurni (lstfurni.Text)
End Sub
Function FindAndReplace(InputString As String, FindString As String, ReplaceString As String) As String
Dim Found() As String
Found() = Split(InputString, FindString)
If UBound(Found) = 0 Then FindAndReplace = InputString: Exit Function
For I = 0 To UBound(Found)
FindAndReplace = FindAndReplace & Found(I) & ReplaceString
Next I
End Function
