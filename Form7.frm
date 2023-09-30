VERSION 5.00
Begin VB.Form Splash 
   BorderStyle     =   0  'None
   ClientHeight    =   4320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6510
   ClipControls    =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSplash 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4345
      Left            =   0
      ScaleHeight     =   4350
      ScaleWidth      =   6540
      TabIndex        =   0
      Top             =   0
      Width           =   6535
      Begin VB.Frame Frame1 
         BackColor       =   &H80000005&
         Height          =   975
         Left            =   130
         TabIndex        =   1
         Top             =   3240
         Visible         =   0   'False
         Width           =   6240
         Begin VB.CommandButton btRegNow 
            Caption         =   "Register Now"
            Height          =   375
            Left            =   3260
            TabIndex        =   3
            Top             =   180
            Width           =   2400
         End
         Begin VB.CommandButton btRegLater 
            Caption         =   "Register Later"
            Height          =   375
            Left            =   480
            TabIndex        =   2
            Top             =   180
            Width           =   2400
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000005&
            Caption         =   "You Have Been Using TaskTracker for 0 Days"
            Height          =   270
            Left            =   1200
            TabIndex        =   4
            Top             =   660
            Width           =   3735
         End
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   6840
         Top             =   3120
      End
      Begin VB.Label txLoading 
         BackColor       =   &H80000005&
         Caption         =   " Loading..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   200
         TabIndex        =   5
         Top             =   200
         Visible         =   0   'False
         Width           =   1400
      End
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ©  2004 Michael M. Ross, Wordwise Solutions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Private splashcnt As Integer

Private Sub btRegNow_Click()
On Error GoTo errlog
'   blLoading = True
   blExpired = True
   Unload Me
   Unload About
   About.Show 1
errlog:
   subErrLog ("Splash:btRegNow_Click")
End Sub

Private Sub btRegLater_Click()
On Error GoTo errlog
   blTempStatus = True
   blExpired = False
'   blLoading = True
   Unload Me
errlog:
   subErrLog ("Splash:btRegLater_Click")
End Sub

Private Sub Form_Load()
On Error GoTo errlog
   Dim ti As Integer
   
   ti = Val(Format$(Time, "ss"))
   
   If ti = 0 Then ti = 60     'possible bugfix
   
   If ti Mod 2 = 0 Then       'randomize buttons
      btRegLater.Left = 380
      btRegNow.Left = 3260
   Else
      btRegLater.Left = 3260
      btRegNow.Left = 380
   End If
   
   RandomMinutesTillExpiration = ti
   
'   If blNoCache = True Then
'      Timer1.Enabled = True
'   Else
'      Timer1.Enabled = False
'   End If
   
   Timer1_Timer
      
errlog:
   subErrLog ("Splash:Form_Load")
End Sub

Public Sub Timer1_Timer()
On Error GoTo errlog

   splashcnt = splashcnt + 1
   
'   Me.Refresh
   
   If blNoCache = False And Form1.blNewTTDir = False Then
      txLoading.Visible = True
      txLoading.ZOrder
      If splashcnt < 10 Then  'don't show 100%
         txLoading = " Loading... " + vbNewLine + " " + Str(splashcnt * 10) + "%"
      End If
   Else
      txLoading = " Loading... "
   End If
   txLoading.Refresh
   
   If splashcnt > 10 Then
      If blRegStatus = False And iDaysUsed > 14 Then
         Splash.Frame1.Visible = True
      Else
         If blNoCache = True Then
            Unload Me
         End If
      End If
   End If
   
errlog:
   subErrLog ("Splash:Timer1_Timer")
End Sub

