VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TaskTracker"
   ClientHeight    =   2175
   ClientLeft      =   2760
   ClientTop       =   3705
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      Caption         =   "Save on Desktop"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Save as files"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Save as shortcuts"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   3855
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Do you want to save this virtual folder as a file folder?"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Option2_Click()

End Sub
