VERSION 5.00
Begin VB.Form Secret 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form3"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1590
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   1590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Secret"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()

If Form1.WindowState <> vbNormal And Form2.WindowState <> vbNormal Then
Form1.Show
Form2.Show
End If
End Sub

'Private Sub Form_GotFocus()
'
'
'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 21 Then  'ctrl+U
Form1.Show
Form2.Show
   End If

End Sub
