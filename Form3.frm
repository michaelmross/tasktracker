VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form About 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About TaskTracker"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5280
   ClipControls    =   0   'False
   Icon            =   "Form3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRegPage 
      Caption         =   "&Get Registration Code"
      Height          =   375
      Left            =   400
      TabIndex        =   14
      Top             =   2475
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.CommandButton cmdSponsor 
      Caption         =   "&Enter Registration Code"
      Height          =   375
      Left            =   2750
      TabIndex        =   13
      Top             =   2475
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.TextBox txSystem 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1575
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1680
      Width           =   800
   End
   Begin VB.TextBox txNetwork 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4050
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1680
      Width           =   800
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
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
      Left            =   3960
      TabIndex        =   6
      Top             =   3160
      Width           =   1215
   End
   Begin VB.TextBox txHistory 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2060
      Width           =   5055
   End
   Begin VB.Label lbBeta 
      Caption         =   "BETA"
      Height          =   255
      Left            =   4580
      TabIndex        =   15
      Top             =   180
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lbUpdate 
      Caption         =   "Check for Update"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3600
      MousePointer    =   4  'Icon
      TabIndex        =   12
      Top             =   840
      Width           =   1455
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   20
      Left            =   100
      Top             =   1560
      Width           =   5070
   End
   Begin VB.Label lbNetwork 
      Caption         =   "Nonsystem/Network"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   2505
      MousePointer    =   4  'Icon
      TabIndex        =   10
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lbSystem 
      Caption         =   "System Drive(s)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   360
      MousePointer    =   4  'Icon
      TabIndex        =   11
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   7
      Top             =   3650
      Width           =   4300
      WordWrap        =   -1  'True
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4440
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   128
      ImageHeight     =   128
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form3.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form3.frx":AF24
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbVersion 
      Caption         =   "Version 1.1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      MousePointer    =   4  'Icon
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Keep Your Files at Your Fingertips™"
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
      Left            =   1920
      TabIndex        =   3
      Top             =   450
      Width           =   2805
   End
   Begin VB.Label Label2 
      Caption         =   "TaskTracker™"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1905
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lbURL 
      Caption         =   "TaskTracker Web Site"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   150
      MousePointer    =   4  'Icon
      TabIndex        =   1
      Top             =   3200
      Width           =   3375
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1335
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "© 2003-2007 Wordwise Solutions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      Top             =   1200
      Width           =   2700
   End
   Begin VB.Menu mnLogging 
      Caption         =   "Logging Menu"
      Visible         =   0   'False
      Begin VB.Menu mLogErrors 
         Caption         =   "&Log Errors"
         Index           =   1
      End
      Begin VB.Menu mSnapshot 
         Caption         =   "Take &Snapshot"
      End
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' © 2003-2005 Michael M. Ross, Wordwise Solutions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" _
   (ByVal lpszUrlName As String) As Long

Private Declare Function URLDownloadToFile Lib "urlmon" _
   Alias "URLDownloadToFileA" _
  (ByVal pCaller As Long, _
   ByVal szURL As String, _
   ByVal szFileName As String, _
   ByVal dwReserved As Long, _
   ByVal lpfnCB As Long) As Long
   
Private strVersion As String
Private blReady As Boolean
Private blRegMode As Boolean
Private blLogErrors As Boolean
Private blLogSnapshot As Boolean

Private Sub cmdRegPage_Click()
On Error GoTo errlog

   Dim sSourceUrl As String, sLocalFile As String
   sSourceUrl = "http://tasktracker.wordwisesolutions.com/download/ttversion"
   sLocalFile = TTpath + "ttversion"
   If About.DownloadFile(sSourceUrl, sLocalFile) = True Then      'use this function (from About box) as test of connectedness
      fnOpenURL ("http://tasktracker.wordwisesolutions.com/update/")
   Else
      MsgBox "Cannot open the Registration web page at this time." _
      + vbNewLine + "You may not be connected to the Internet.", vbInformation
   End If
   
errlog:
   subErrLog ("About:cmdRegPage_Click")
End Sub

Private Sub cmdSponsor_Click()
   RegEnter
End Sub

Private Sub Command1_Click()
On Error GoTo errlog

   If blRegStatus = False Then
      If blExpired = True Then
         If iDaysUsed > 45 Then     'just expire
            ClearAutoRun            'don't autorestart
            'this removes the icon from the system tray
            Shell_NotifyIcon NIM_DELETE, nid
            End
         Else
            Form1.blStartNag = False
            blExpired = False
         End If
      End If
      Form1.Timer1.Enabled = True
   End If
   
errlog:
   Unload Me
   subErrLog ("About:OK")
End Sub

Private Sub RegEnter()
On Error GoTo errlog

   With txHistory
      .Enabled = True
      .Locked = False
      .Text = "Type or paste your registration code here"
      .SelLength = 41
      .SetFocus
      cmdSponsor.Visible = False
      cmdRegPage.Visible = False
      blRegMode = True
   End With

errlog:
   subErrLog ("About:RegEnter")
End Sub

Private Sub Form_Load()
On Error GoTo errlog

Dim wc As Integer, tc As Integer
   
'   Form1.mHideShow(1).Enabled = False
   strVersion = Trim$(Str$(VB.App.Major)) + "." + Trim$(Str$(VB.App.Minor)) + "." + Trim$(Str$(VB.App.Revision))
   lbVersion = "Version " + strVersion
   
   If VB.App.Comments = "Beta Release" Then
      lbBeta.Visible = True
   End If
   
   Set Image1.Picture = ImageList1.ListImages(1).Picture
   lbURL.MouseIcon = ImageList1.ListImages(2).Picture
   lbURL.MousePointer = vbCustom
   lbUpdate.MouseIcon = ImageList1.ListImages(2).Picture
   lbUpdate.MousePointer = vbCustom
   
   subGetFixedDrives
   
   wc = Form1.FilesCountAll(sRecent, "*.lnk")
   tc = Form1.FilesCountAll(TTpath, "*.lnk")
   
   Label4.Caption = "Protected by Copyright Law and International Treaties."
   
   If blExpired = False Then
   
      If blRegStatus = False Then
         txHistory.Text = "Unregistered User: Your TaskTracker file history contains " + Trim$(Str$(tc)) + " items, " _
          + "versus " + Trim$(Str$(wc)) + " items in your Windows file history. " _
          + "The longer you use your TaskTracker the more files and folders " _
          + "it will track and the more useful it will be. " _
          + "Days used: " + Trim$(Str$(iDaysUsed))
      Else
'         Dim iDays As String
'         If iDaysUsed = 1 Then
'            iDays = "1 day "
'         Else
'            iDays = Trim$(Str$(iDaysUsed)) + " days "
'         End If
         txHistory.Text = RegName + " - " + RegStatus + " since " + _
         RegDate + ". (Days since your last update: " + Trim$(Str$(iDaysUsed)) + ")" + vbNewLine + _
         "Your TaskTracker file history contains " + Trim$(Str$(tc)) + " items, " + _
         "versus " + Trim$(Str$(wc)) + " items in your Windows file history. "
      End If
   
      If Len(sDrives) > 0 Then
         lbSystem.Enabled = True
         lbSystem.MouseIcon = ImageList1.ListImages(2).Picture
         lbSystem.MousePointer = vbCustom
      End If
      If Len(nDrives) > 0 Then
         lbNetwork.Enabled = True
         lbNetwork.MouseIcon = ImageList1.ListImages(2).Picture
         lbNetwork.MousePointer = vbCustom
      End If
          
   Else
      cmdSponsor.Visible = True
      cmdRegPage.Visible = True
      txHistory.Text = "Thank you for evaluating TaskTracker!"
   End If
      
   If LenB(GetSetting("TaskTracker", "Settings", "Logging")) > 0 Then
      About.mLogErrors(1).Checked = True
   Else
      About.mLogErrors(1).Checked = False 'default
   End If

   With Form1
      .Label1.ToolTipText = "Tracking " + Trim$(Str$(modData.TotalFiles)) + " files (" + Trim$(Str$(tc)) + " shortcuts)"
      .StatusBar1.Panels(1).ToolTipText = "Tracking " + Trim$(Str$(modData.TotalFiles)) + " files (" + Trim$(Str$(tc)) + " shortcuts)"
   End With
   
errlog:
   subErrLog ("About:FormLoad")
End Sub

Private Function RegStatus() As String

   Select Case RegType
      Case "T"
         RegStatus = "Temporary User"
      Case "C"
         RegStatus = "Registered User"
      Case "S"
         RegStatus = "Registered User"
      Case "B"
         RegStatus = "Registered Business User"
      Case "P"
         RegStatus = "Registered Business User"
   End Select
          
End Function

Private Sub Form_Mouseup(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo errlog

   If Button = 2 Then
      PopupMenu mnLogging
   End If

errlog:
   subErrLog ("About:Form_Mouseup")
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errlog

   If blReady = True Then
      If txHistory.Text = "Your Name, Your Company" Then
         txHistory.SelLength = Len(txHistory.Text)
         MsgBox "Please customize your TaskTracker with your name as you would like it to appear in the About box.", vbExclamation
         txHistory.SetFocus
         Cancel = 1
         Exit Sub
      Else
         Dim Response As Integer
         Response = MsgBox("Are your name (and company name) correct?" _
                           + vbNewLine + vbNewLine + txHistory.Text, 52)
         If Response = vbYes Then
            blJustRegd = True
            SaveRegDetails
            Form1.Timer1.Enabled = True
         Else
            txHistory.SetFocus
            Cancel = 1
         End If
         blRegMode = False
      End If
      Exit Sub
   End If
   
errlog:
   subErrLog ("About:Form_Unload")
End Sub

Private Sub SaveRegDetails()
On Error Resume Next

   Dim iMax As String, LastDate As Integer
   Dim DaysUsed As Integer
   
   GetLocalTime SYS_TIME
   LastDate = SYS_TIME.wDay
   iMax = "noexpiry"
   DaysUsed = Str$(iDaysUsed)
   RegDate = Str$(Date)
   RegName = txHistory.Text
      
   Call Form1.RegWrite(iMax, LastDate, DaysUsed)
   
   blReady = False

errlog:
   subErrLog ("About:SaveRegDetails")
End Sub

Private Sub lbSystem_Click()
   MsgBox "System drives that TaskTracker is monitoring:" _
   + vbNewLine + sDrives, vbInformation
End Sub

Private Sub lbNetwork_Click()
   MsgBox "Nonsystem and network drives that TaskTracker is monitoring:" _
   + vbNewLine + nDrives, vbInformation
End Sub

Private Sub lbURL_Click()
On Error GoTo errlog
   
   fnOpenURL ("http://tasktracker.wordwisesolutions.com/")

errlog:
   subErrLog ("About:lbURL_Click")
End Sub

Public Sub lbUpdate_Click()
On Error GoTo errlog
   
   Dim sSourceUrl As String
   Dim sLocalFile As String
   Dim hfile As Long
   Dim CurVer As String
   Dim duce As Boolean
     
   sSourceUrl = "http://tasktracker.wordwisesolutions.com/download/V11/ttversion"
   sLocalFile = TTpath + "ttversion"
      
   duce = DeleteUrlCacheEntry(sSourceUrl)

   If DownloadFile(sSourceUrl, sLocalFile) Then
   
      hfile = FreeFile
      Open sLocalFile For Input As #hfile
         CurVer = Trim$(Input$(LOF(hfile), hfile))
      Close #hfile
'      CurVer = "1.1.99"   'test
      strVersion = Trim$(Str$(VB.App.Major)) + "." + Trim$(Str$(VB.App.Minor)) + "." + Trim$(Str$(VB.App.Revision))
      'doesn't evaluate numbers correctly with two dots
      If Val(Right$(CurVer, Len(CurVer) - 4)) > Val(Right$(strVersion, Len(strVersion) - 4)) Or _
         VB.App.Comments = "Beta Release" And Val(Right$(CurVer, Len(CurVer) - 4)) >= Val(Right$(strVersion, Len(strVersion) - 4)) Then
         Dim Response As Integer
         If VB.App.Comments = "Beta Release" Then
            If blLoading = False Then
               Response = MsgBox("This beta release has expired. The final release of this TaskTracker version is available." _
               + vbNewLine + "Click OK to download now.", 65)
            Else
               Response = MsgBox("This beta release has expired. The final release of this TaskTracker version is available." _
               + vbNewLine + "Click OK to download now." + vbNewLine + vbNewLine + vbNewLine _
               + "(Choose Preferences > Startup Options to prevent update checks.)", 65)
            End If
         Else
            If blLoading = False Then
               Response = MsgBox("A newer version of TaskTracker is available." _
               + vbNewLine + "Click OK to download now.", 65)
            Else
               Response = MsgBox("A newer version of TaskTracker is available." _
               + vbNewLine + "Click OK to download now." + vbNewLine + vbNewLine + vbNewLine _
               + "(Choose Preferences > Startup Options to prevent update checks.)", 65)
            End If
         End If
         If Response = vbOK Then
            fnOpenURL ("http://tasktracker.wordwisesolutions.com/update/update.htm")
            If blLoading = True Then
               MsgBox "Restart TaskTracker after installing the update.", vbInformation
'              ***********************************************
               Shell_NotifyIcon NIM_DELETE, nid
               End
'              ***********************************************
            Else
               MsgBox "Exit TaskTracker before installing the update.", vbInformation
            End If
         End If
'        Beta Release Terminate
'        ***********************************************
         If Response = vbCancel Then
            If VB.App.Comments = "Beta Release" Then
               Shell_NotifyIcon NIM_DELETE, nid
               End
            End If
         End If
'        ***********************************************
      Else
         If blLoading = False Then
            If VB.App.Comments = "Beta Release" Then
               MsgBox "You have a beta release of the next version of TaskTracker.", vbInformation
            Else
               MsgBox "You have the current version of TaskTracker.", vbInformation
            End If
         End If
      End If
      
   Else
   
      If blLoading = False Then
         MsgBox "There may be a problem with your Internet connection. " _
         + vbNewLine + "Please try again later when you're online.", vbInformation
'         + vbNewLine + "(Use the link at the bottom of the About box to go to the TaskTracker home page.)", vbInformation
      End If
      
   End If
   
   Exit Sub
   
errlog:
   MsgBox "Please try updating TaskTracker later.", vbInformation
   subErrLog ("About:lbUpdate_Click")
End Sub

Public Function DownloadFile(URL As String, LocalFilename As String) As Boolean
    Dim lngRetVal As Long
    lngRetVal = URLDownloadToFile(0, URL, LocalFilename, 0, 0)
    If lngRetVal = 0 Then DownloadFile = True
End Function

Private Sub mLogErrors_Click(Index As Integer)
On Error GoTo errlog

   If mLogErrors(1).Checked = True Then
      mLogErrors(1).Checked = False
      DeleteSetting "TaskTracker", "Settings", "Logging"
   Else
      mLogErrors(1).Checked = True
      If blLogErrors = False Then
         MsgBox "A log of errors will be saved to your " + App.path + " folder.", vbInformation
         SaveSetting "TaskTracker", "Settings", "Logging", "1"
         blLogErrors = True
      End If
   End If
   
errlog:
   subErrLog ("About:mLogErrors_Click")
End Sub

Private Sub mSnapshot_Click()
On Error GoTo errlog

   If mSnapshot.Checked = True Then
      mSnapshot.Checked = False
      DeleteSetting "TaskTracker", "Settings", "Snapshot"
   Else
      mSnapshot.Checked = True
      If blLogSnapshot = False Then
         MsgBox "A onetime snapshot of the files TaskTracker is tracking on you system will be saved to your " _
         + vbNewLine + App.path + " folder the next time you start TaskTracker.", vbInformation
         SaveSetting "TaskTracker", "Settings", "Snapshot", "1"
         blLogSnapshot = True
      End If
   End If
   
errlog:
   subErrLog ("About:mSnapErrors_Click")
End Sub

Private Sub txHistory_Change()
On Error GoTo errlog

   If blRegMode = True And blRegStatus = False Then
      Dim ct As String
      With txHistory
         ct = Trim$(TrimNull(.Text))
         If Len(ct) = 9 Then
            If Left$(ct, 1) = "C" Or Left$(ct, 1) = "S" Or Left$(ct, 1) = "B" Or Left$(ct, 1) = "P" Or Left$(ct, 1) = "T" Then
               If Right$(ct, 1) = "%" Then
                  If InStr(CurrentCode, .Text) > 0 Then     'CurrentCode function gets ttcodes file and checks first 9 chars of provided code against it
                      RegType = Left$(Trim$(.Text), 1)
                      .Text = "Your Name, Your Company"
                      .SelLength = Len(.Text)
                      MsgBox "Thank you for registering!" + vbNewLine + _
                         "Type your name (and company name) the way you would like them to appear." + vbNewLine + _
                         "When you are ready, click OK to close the About box.", vbInformation
                      .SetFocus
                      blExpired = False
                      blRegStatus = True
                      blReady = True
                  End If
               End If
            End If
         End If
      End With
   End If
   
errlog:
   subErrLog ("About:txHistory_Change")
End Sub

Private Function CurrentCode() As String
On Error GoTo errlog
   
   Dim sSourceUrl As String
   Dim sLocalFile As String
   Dim hfile As Long
   Dim duce As Boolean
   Dim NiceLongKey As String
     
   sSourceUrl = "http://tasktracker.wordwisesolutions.com/download/ttcodes"
   sLocalFile = TTpath + "ttcodes"
      
   duce = DeleteUrlCacheEntry(sSourceUrl)
   
   NiceLongKey = "© 2003-2005 Michael M. Ross, Wordwise Solutions"

   If DownloadFile(sSourceUrl, sLocalFile) Then
      
     hfile = FreeFile
     
   hfile = FreeFile
   Open sLocalFile For Input As #hfile
      CurrentCode = DeCrypt(Trim$(Input$(LOF(hfile), hfile)), NiceLongKey)
   Close #hfile

                  
   Else        'if connection problem use hard-coded codes if necessary
   
     Dim cfile As String

     cfile = "236 106 85 160 118 129 150 101 85 131 101 113 182 190 203 223 179 138 174 75 187 99 126 170 181 220 152 124 115 144 148 224 199 195 169 152 185 82 198 167 172 183 226 189 148"

     CurrentCode = DeCrypt(cfile, NiceLongKey)

'      MsgBox "You may need to disable your firewall while you complete your registration. " _
'      + vbNewLine + "Please contact support@wordwisesolutions.com if the problem persists.", vbInformation
'
'      blRegMode = False
'
'      txHistory.Text = "Registration unsuccessful. Please try again later."
'
'      blReady = False
      
   End If
   
   
errlog:
   subErrLog ("About:lbUpdate_Click")
End Function
