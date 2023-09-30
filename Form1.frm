VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.Ocx"
Begin VB.Form Form1 
   Caption         =   "TaskTracker"
   ClientHeight    =   6760
   ClientLeft      =   169
   ClientTop       =   559
   ClientWidth     =   2652
   ClipControls    =   0   'False
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6760
   ScaleWidth      =   2652
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   4560
   End
   Begin VB.PictureBox tticon 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   4
      Left            =   840
      ScaleHeight     =   234
      ScaleWidth      =   234
      TabIndex        =   11
      Top             =   2640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox tticon 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   840
      ScaleHeight     =   234
      ScaleWidth      =   234
      TabIndex        =   10
      Top             =   2280
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox tticon 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   840
      ScaleHeight     =   234
      ScaleWidth      =   234
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox tticon 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   840
      ScaleHeight     =   234
      ScaleWidth      =   234
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1680
      Top             =   3720
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   95
      TabIndex        =   0
      Top             =   5775
      Visible         =   0   'False
      Width           =   2390
      _ExtentX        =   4409
      _ExtentY        =   479
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Frame frTemp 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   380
      Visible         =   0   'False
      Width           =   2400
      Begin VB.Label lbTemp 
         BackColor       =   &H80000005&
         Caption         =   "Initializing data... please wait"
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   2415
      End
   End
   Begin VB.PictureBox tticon 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   840
      ScaleHeight     =   481
      ScaleWidth      =   481
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   3600
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   6465
      Width           =   2655
      _ExtentX        =   4673
      _ExtentY        =   527
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7056
            MinWidth        =   7056
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   10583
            MinWidth        =   10583
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.1509
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   5595
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5955
      _ExtentX        =   10975
      _ExtentY        =   10304
      View            =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDropMode     =   1
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.1509
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin VB.PictureBox pixSmall 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1680
      ScaleHeight     =   234
      ScaleWidth      =   234
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   240
   End
   Begin ComctlLib.ImageList ImageList3 
      Left            =   240
      Top             =   4680
      _ExtentX        =   1006
      _ExtentY        =   1006
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0442
            Key             =   "TT16"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":061C
            Key             =   "Virtual"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   1200
      Top             =   2160
      _ExtentX        =   1006
      _ExtentY        =   1006
      BackColor       =   -2147483633
      MaskColor       =   -2147483633
      _Version        =   327682
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   2640
      _ExtentX        =   1006
      _ExtentY        =   1006
      BackColor       =   -2147483633
      MaskColor       =   -2147483628
      _Version        =   327682
   End
   Begin VB.Label Label1 
      Caption         =   "Click a file type to view recently used files of this type."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Top             =   5680
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Menu mnOptions 
      Caption         =   "Options Menu"
      Visible         =   0   'False
      Begin VB.Menu mOpenType 
         Caption         =   "&View Files..."
         Index           =   1
      End
      Begin VB.Menu mVirtual 
         Caption         =   "&Virtual Folders..."
         Shortcut        =   ^U
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mRefresh 
         Caption         =   "Re&fresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mReload 
         Caption         =   "Rel&oad"
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mSort 
         Caption         =   "&Sort"
         Begin VB.Menu mAlpha 
            Caption         =   "&Alphabetic"
            Checked         =   -1  'True
            Shortcut        =   ^S
         End
         Begin VB.Menu mNumber 
            Caption         =   "&Frequency"
         End
         Begin VB.Menu mLatest 
            Caption         =   "&Most Recent"
         End
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mHideType 
         Caption         =   "&Hide This File Type"
      End
      Begin VB.Menu mShowAll 
         Caption         =   "Sho&w All File Types"
         Checked         =   -1  'True
      End
      Begin VB.Menu mNetwork 
         Caption         =   "Show &Network Files as Type"
         Checked         =   -1  'True
      End
      Begin VB.Menu mRemovable 
         Caption         =   "Show Removab&le Files as Type"
         Checked         =   -1  'True
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mRegister2 
         Caption         =   "&Register..."
      End
      Begin VB.Menu mPreferences 
         Caption         =   "&Preferences"
         Shortcut        =   ^N
      End
      Begin VB.Menu mHelp2 
         Caption         =   "&Did You Know?"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mPopAbout2 
         Caption         =   "&About TaskTracker"
         Shortcut        =   ^B
      End
      Begin VB.Menu mPopExit2 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnPreferences 
      Caption         =   "Preferences Menu"
      Visible         =   0   'False
      Begin VB.Menu mFade 
         Caption         =   "&Fade In and Out"
      End
      Begin VB.Menu mTaskbar 
         Caption         =   "Show in Tas&kbar"
      End
      Begin VB.Menu mTop 
         Caption         =   "&TaskTracker on Top"
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mAutoPreview 
         Caption         =   "&Autoshow Image Previews"
      End
      Begin VB.Menu mHideTips 
         Caption         =   "Hide T&ooltips"
      End
      Begin VB.Menu mIcons 
         Caption         =   "Show &Icons Only"
      End
      Begin VB.Menu mGrid 
         Caption         =   "Show &Gridlines"
         Begin VB.Menu mGridType 
            Caption         =   "File &Type List"
         End
         Begin VB.Menu mGridName 
            Caption         =   "File &Name List"
         End
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mSaveSize 
         Caption         =   "Remember My TaskTracker La&yout"
      End
      Begin VB.Menu mSync 
         Caption         =   "Synchroni&ze TaskTracker Windows"
         Begin VB.Menu mSingleClick 
            Caption         =   "&Single-Click Open/Close"
         End
         Begin VB.Menu mAutohide 
            Caption         =   "&Autohide on File Open"
         End
         Begin VB.Menu mDefaultFocus 
            Caption         =   "&Default Focus"
            Begin VB.Menu mFocusTypes 
               Caption         =   "On &Type List"
            End
            Begin VB.Menu mFocusFiles 
               Caption         =   "On &File List"
            End
         End
         Begin VB.Menu mGlue 
            Caption         =   "&Glue Together"
            Begin VB.Menu mGlueRight 
               Caption         =   "File Types on &Right"
            End
            Begin VB.Menu mGlueLeft 
               Caption         =   "File Types on &Left"
            End
         End
      End
      Begin VB.Menu mStartup 
         Caption         =   "Startup Options"
         Begin VB.Menu mNoSplash 
            Caption         =   "&No Splash Screen"
         End
         Begin VB.Menu mCheckUpdates 
            Caption         =   "&Check for Updates"
         End
         Begin VB.Menu mWinStart 
            Caption         =   "&Start with Windows"
            Begin VB.Menu mMin 
               Caption         =   "&Minimized"
            End
            Begin VB.Menu mNorm 
               Caption         =   "&Normal"
            End
         End
      End
      Begin VB.Menu mCacheSettings 
         Caption         =   "Performance Settings"
         Begin VB.Menu mNoCache 
            Caption         =   "&No Caching"
         End
         Begin VB.Menu mAlwaysValidate 
            Caption         =   "&Always Validate Files"
         End
         Begin VB.Menu mTrueAccessDates 
            Caption         =   "&True Accessed Dates"
         End
         Begin VB.Menu mSynchronization 
            Caption         =   "&Synchronization"
            Begin VB.Menu mSyncShow 
               Caption         =   "On First S&howing"
            End
            Begin VB.Menu mSyncStartup 
               Caption         =   "On S&tartup"
            End
            Begin VB.Menu mNeverSync 
               Caption         =   "&Never"
            End
         End
      End
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mHideShow 
         Caption         =   "&Hide TaskTracker"
         Index           =   1
      End
      Begin VB.Menu mPopExit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu separator 
         Caption         =   "-"
      End
      Begin VB.Menu mRegister 
         Caption         =   "&Register..."
      End
      Begin VB.Menu mPreferences2 
         Caption         =   "&Preferences"
      End
      Begin VB.Menu mHelp 
         Caption         =   "&Did You Know?"
      End
      Begin VB.Menu mPopAbout 
         Caption         =   "&About TaskTracker"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' © 2003-2007 Michael M. Ross, Wordwise Solutions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
  
'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
'         (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" _
         (ByVal hWnd As Long) As Long
         
Private Declare Function SetProcessWorkingSetSize Lib "kernel32.dll" _
   (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, _
    ByVal dwMaximumWorkingSetSize As Long) As Long
    
Private Declare Function OpenProcess Lib "kernel32" _
   (ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
        
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private Declare Function GetWindow Lib "user32" _
    (ByVal hWnd As Long, ByVal wCmd As Long) As Long
    
Private Declare Function MapVirtualKey Lib "user32" _
    Alias "MapVirtualKeyA" (ByVal wCode As Long, _
    ByVal wMapType As Long) As Long

Private Declare Function GetWindowText Lib "user32" _
    Alias "GetWindowTextA" (ByVal hWnd As Long, _
    ByVal lpString As String, ByVal cch As Long) _
    As Long
    
Private Declare Function GetClassName Lib "user32" _
    Alias "GetClassNameA" (ByVal hWnd As Long, _
    ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    
Private Const PROCESS_ALL_ACCESS = &H1F0FFF

Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const KD_EXTENDED = &H1000000
Private Const KD_DOWN = 0
Private Const KD_UP = &HC0000000

Public Enum VIRTUAL_KEY
    VK_F24 = &H87&
End Enum

Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDNEXT = 2

Private Const ILC_MASK = 1&
Private Const ILC_COLOR8 = &H8&

Private Const vbKeyF24 As Integer = vbKeyF16 + 8

'Private TaskBarID As Long
Private himgSmall As Long
Private lnkcnt As Long
Private filecnt As Long

Private wBottom As Integer
Private HD As Integer
Private iRefreshTime As Integer
Private meleft As Integer
Private metop As Integer
Private meheight As Integer
Private pnt As Integer
Private ProgPercent As Integer
Private LastwBottom As Integer
Private LastwWidth As Integer

Private fNoShow() As String
Private fname As String
'Private sLastFileName As String
Private KeepType As String

Public blNewType As Boolean
Public blStartNag As Boolean
Public blNewTTDir As Boolean
Public blReloading As Boolean
Public blNoSysTray As Boolean
Public blSetTempCaption As Boolean

Private blWait As Boolean
Private blRightClick As Boolean
Private blHideType As Boolean
Private blShift As Boolean
Private blTaskbar As Boolean
Private blFade As Boolean
Public blNewShortcuts As Boolean
'Private blLockWin As Boolean
Private blKeyClick As Boolean
Private blFirstGoodFile As Boolean
'Private blUpdatedSinceCacheStart As Boolean
'Private blLoadedVirtualType As Boolean
Private blNeedUpdate As Boolean
Private blLoadModal As Boolean
Private FirstFolderCheck As Boolean
'Private blAlwaysValidate As Boolean
Private blSaveLayout As Boolean
Private blReposition As Boolean
'Private blTTTemp As Boolean

Private SecsAfterMidnight As Single

Private itmT As ListItem

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Label1.ToolTipText = "Tracking " + Trim$(Str$(modData.TotalFiles)) + " files (" + Trim$(Str$(fCount)) + " shortcuts)"
End Sub

Private Sub StatusBar1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   StatusBar1.Panels(1).ToolTipText = "Tracking " + Trim$(Str$(modData.TotalFiles)) + " files (" + Trim$(Str$(fCount)) + " shortcuts)"
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo errlog
   
   If mHideTips.Checked = True Then Exit Sub
   
   Dim lvhti As LVHITTESTINFO
   Dim lItemIndex As Long
   
   lvhti.pt.x = x \ Screen.TwipsPerPixelX
   lvhti.pt.y = y \ Screen.TwipsPerPixelY
   lItemIndex = SendMessage(ListView1.hWnd, LVM_HITTEST, 0, lvhti) + 1
   
   If m_lCurItemIndex <> lItemIndex Then
      m_lCurItemIndex = lItemIndex
      If m_lCurItemIndex = 0 Then   ' no item under the mouse pointer
         TT.Destroy
      Else
         If blLoading = False Then
            If ListView1.ListItems(m_lCurItemIndex).Text <> "Virtual Folders" Then
               Dim LatestFile As String
               If Len(ListView1.ListItems(m_lCurItemIndex).SubItems(2)) > 0 Then
                  LatestFile = ListView1.ListItems(m_lCurItemIndex).SubItems(2)
               Else
                  LatestFile = "Unknown"
               End If
               TT.Title = ListView1.ListItems(m_lCurItemIndex).Text
               TT.TipText = "File Count: " + LTrim$(ListView1.ListItems(m_lCurItemIndex).SubItems(1)) + vbNewLine + _
                            "Latest File: " + LatestFile
'            Else
'               Dim AVs As Variant
'               Dim av As Integer
'               Dim vn As String
'               AVs = GetAllSettings("TaskTracker", "TTI\Virtual")
'
'               For av = UBound(AVs, 1) To LBound(AVs, 1) Step -1
'                  vn = AVs(av, 1) + ", " + vn
'               Next av
'
'               TT.Title = Str$(Trim$(UBound(AVs, 1) + 1)) + " Virtual Folders"
'               TT.TipText = Left$(vn, Len(vn) - 2)
            End If
            TT.Create ListView1.hWnd
            If mIcons.Checked = False Then
               TT.SetDelayTime (sdtInitial), 1000
            Else
               TT.SetDelayTime (sdtInitial), 500
            End If
         End If
      End If
   End If
errlog:
   subErrLog ("Form1:ListView1_MouseMove")
End Sub

Private Sub mautohide_Click()
On Error GoTo errlog
   If mAutohide.Checked = True Then
      mAutohide.Checked = False
   Else
      mAutohide.Checked = True
   End If
errlog:
   subErrLog ("Form1:mautohide_Click")
End Sub

Private Sub mCheckUpdates_Click()
On Error GoTo errlog
   If mCheckUpdates.Checked = True Then
      mCheckUpdates.Checked = False
   Else
      mCheckUpdates.Checked = True
   End If
errlog:
   subErrLog ("Form1:mCheckUpdates_Click")
End Sub

Private Sub mNoCache_Click()
On Error GoTo errlog
   If mNoCache.Checked = True Then
      mNoCache.Checked = False
      blNoCache = False
      mSynchronization.Enabled = True
   Else
      If blLoading = False Then
         Dim Response As Integer
         Response = MsgBox("This option should only be selected if TaskTracker has apparently stopped tracking files." + vbNewLine _
         + "(This selection will be removed after you restart TaskTracker.) Do you really want to continue?", 308)
         If Response = vbNo Then
            Exit Sub
         End If
      End If
      If blLoading = False Then
         mNoCache.Checked = True    'only user selected - not when loading
         mSynchronization.Enabled = False
      End If
      blNoCache = True
   End If
errlog:
   subErrLog ("Form1:mNoCache_Click")
End Sub

Private Sub mFade_Click()
On Error GoTo errlog
   If mFade.Checked = False Then
      mFade.Checked = True
      If blFade = False Then
         MsgBox "This option autohides TaskTracker when you are using another program." + vbNewLine _
         + "To restore TaskTracker, simply click it or the system tray icon.", vbInformation
      End If
      If blTaskbarStatus = False Then  'need to avoid the situation where taskbar status is still true but unchecked, enabling this option
         blNoFade = False
         Timer2.Enabled = True
      End If
   Else
      mFade.Checked = False
      subRestore
      If mGlueRight.Checked = False And mGlueLeft.Checked = False Then
         Timer2.Enabled = False
      End If
   End If
   blFade = True
errlog:
   subErrLog ("Form1:mFade_Click")
End Sub

Private Sub mFocusFiles_Click()
On Error GoTo errlog
   If mFocusFiles.Checked = True Then
      mFocusFiles.Checked = False
      mFocusTypes.Checked = True
      Form1.SetFocus
   Else
      mFocusFiles.Checked = True
      mFocusTypes.Checked = False
      Form2.SetFocus
   End If
errlog:
   subErrLog ("Form1:mFocusFiles_Click")
End Sub

Private Sub mFocusTypes_Click()
On Error GoTo errlog
   If mFocusTypes.Checked = True Then
      mFocusTypes.Checked = False
      mFocusFiles.Checked = True
      Form2.SetFocus
   Else
      mFocusTypes.Checked = True
      mFocusFiles.Checked = False
      Form1.SetFocus
   End If
errlog:
   subErrLog ("Form1:mFocusTypes_Click")
End Sub

Private Sub mGridName_Click()
On Error GoTo errlog
   If mGridName.Checked = False Then
      mGridName.Checked = True
   Else
      mGridName.Checked = False
   End If
   Call SendMessage(Form2.ListView2.hWnd, _
                    LVM_SETEXTENDEDLISTVIEWSTYLE, _
                    LVS_EX_GRIDLINES, ByVal mGridName.Checked)
errlog:
   subErrLog ("Form1:mGridName_Click")
End Sub

Public Sub mHelp_Click()
On Error GoTo errlog

   Dim sSourceUrl As String, sLocalFile As String
   sSourceUrl = "http://tasktracker.wordwisesolutions.com/download/ttversion"
   sLocalFile = TTpath + "ttversion"
   If About.DownloadFile(sSourceUrl, sLocalFile) = True Then      'use this function (from About box) as test of connectedness
      fnOpenURL ("http://tasktracker.wordwisesolutions.com/didyouknow/")
   Else
      MsgBox "Cannot display " + Chr$(34) + "TaskTracker - Did You Know?" + Chr$(34) + " at this time." _
      + vbNewLine + "You may not be connected to the Internet.", vbInformation
   End If
   
errlog:
   subErrLog ("Form1:mHelp_Click")
End Sub

Private Sub mHelp2_Click()
   mHelp_Click
End Sub

Private Sub mMin_Click()
On Error GoTo errlog
   If mMin.Checked = True Then
      mMin.Checked = False
      mNorm.Checked = False
      ClearAutoRun
   Else
      mMin.Checked = True
      mNorm.Checked = False
      SetAutoRun
   End If
errlog:
   subErrLog ("Form1:mMin_Click")
End Sub

Private Sub mNorm_Click()
On Error GoTo errlog
   If mNorm.Checked = True Then
      mNorm.Checked = False
      mMin.Checked = False
      ClearAutoRun
   Else
      mNorm.Checked = True
      mMin.Checked = False
      SetAutoRun
   End If
errlog:
   subErrLog ("Form1:mNorm_Click")
End Sub

Private Sub mNetwork_Click()
On Error GoTo errlog
   Dim nw As Long
   Dim ni As Long
   Dim lv1 As Integer
   Dim lgx As ListItem
   
   If mNetwork.Checked = True Then
      mNetwork.Checked = False
      
      Dim blAddtheType As Boolean
      For ni = 0 To UBound(NetworkDrives)
         If Len(NetworkDrives(ni)) > 0 Then
            For nw = 0 To UBound(fType)
               If nw = UBound(fType) Then Exit For
               If Left$(aSFileName(nw), 3) = NetworkDrives(ni) Or Left$(aSFileName(nw), 2) = "\\" Then
                  blNotValidated(nw) = True
                  fType(nw) = tType(nw)
                  
                  'Add any missing icons
                  If Len(fType(nw)) > 0 Then
                     If Left$(fType(nw), 5) <> "hide-" Then
                        blAddtheType = True
                        For lv1 = 1 To ListView1.ListItems.Count
                          If fType(nw) = ListView1.ListItems(lv1).Text Then
                             blAddtheType = False
                             Exit For
                          End If
                        Next
                        If blAddtheType = True Then
                           Call GetNewIcon(nw, fType(nw))
                        End If
                     End If
                  End If
                  
               End If
            Next nw
         End If
      Next ni
      
      For lv1 = 1 To ListView1.ListItems.Count
         If ListView1.ListItems(lv1).Text = "Network Drive" Then
            ListView1.ListItems.Remove ("Network Drive")
            Exit For
         End If
      Next lv1
      
   Else
   
      mNetwork.Checked = True
      Dim blAddNetwork As Boolean
      
     'First determine which files are on network drives or have UNC paths
     For ni = 0 To UBound(NetworkDrives)
         If Len(NetworkDrives(ni)) > 0 Then
            For nw = 0 To UBound(fType)
               If nw = UBound(fType) Then Exit For
               If Left$(aSFileName(nw), 3) = NetworkDrives(ni) Then
                  blNotValidated(nw) = True
                  fType(nw) = "Network Drive"
                  blAddNetwork = True
               End If
            Next nw
         End If
      Next ni
      
      For nw = 0 To UBound(fType)
         If nw = UBound(fType) Then Exit For
         If Left$(aSFileName(nw), 2) = "\\" Then
            blNotValidated(nw) = True
            fType(nw) = "Network Drive"
            blAddNetwork = True
         End If
      Next nw
      
      'Then add the Network type and delete any empty types
      If blAddNetwork = True Then
         Dim blRemoveType As Boolean
         
         Set lgx = ListView1.ListItems.Add(, "Network Drive", "Network Drive")
         lgx.SmallIcon = ImageList1.ListImages("Network Drive").Key
         mNetwork.Checked = True
         
         For lv1 = 1 To ListView1.ListItems.Count
            If lv1 = ListView1.ListItems.Count Then Exit For
            If ListView1.ListItems(lv1).Text <> "Network Drive" Then
               blRemoveType = True
               For nw = 0 To UBound(fType)
                  If nw = UBound(fType) Then Exit For
                  If Len(fType(nw)) > 0 Then
                     If fType(nw) = ListView1.ListItems(lv1).Text Then
                         blRemoveType = False
                         Exit For
                     End If
                  End If
               Next nw
               If blRemoveType = True Then
                  ListView1.ListItems.Remove (lv1)
               End If
            End If
         Next lv1

      Else
         MsgBox "No networked drives are currently connected, or there are no files to show.", vbExclamation
         Exit Sub
      End If
      
   End If
   
   subClickList
   ListView1_Click

errlog:
   subErrLog ("Form1:mNetwork_click")
End Sub

Private Sub GetNewIcon(gni As Long, gtype As String)
On Error GoTo errlog
   Dim img2 As ListImage
   Dim lgx As ListItem

'   AddFileItemIcons (TTpath + aShortcut(gni))
'
'   Set img1 = ImageList1.ListImages.Add(, gtype + "g", pixSmall.Picture)
'   Set lgx = ListView1.ListItems.Add(, gtype + "g", gtype)
'   lgx.SmallIcon = ImageList1.ListImages(gtype + "g").Key
   
   AddFileItemIcons (TTpath + aShortcut(gni))
   Set img2 = ImageList2.ListImages.Add(, gtype + "g", pixSmall.Picture)
   lgx.SmallIcon = Form1.ImageList2.ListImages(gtype + "g").Key
   
errlog:
   subErrLog ("Form1:GetNewIcon")
End Sub

Private Sub mAlwaysValidate_Click()
On Error GoTo errlog
   If mAlwaysValidate.Checked = True Then
      mAlwaysValidate.Checked = False
   Else
'      If blAlwaysValidate = False Then
'         MsgBox "This option can increase accuracy but may slow performance.", vbInformation
'         blAlwaysValidate = True
'      End If
      Erase blTrueDates
      mAlwaysValidate.Checked = True
   End If
errlog:
   subErrLog ("Form1:mAlwaysValidate_Click")
End Sub

Private Sub mPopExit2_Click()
On Error GoTo errlog
   mPopExit_Click
errlog:
   subErrLog ("Form1:mPopExit2_Click")
End Sub

Public Sub mPreferences_Click()
On Error GoTo errlog
   blMenu = True
   Timer3.Enabled = True
errlog:
   subErrLog ("Form1:mPreferences_Click")
End Sub

Private Sub mPreferences2_Click()
On Error GoTo errlog
   blMenu = True
   Timer3.Enabled = True
errlog:
   subErrLog ("Form1:mPreferences2_Click")
End Sub

Private Sub mRegister_Click()
On Error GoTo errlog
   Unload About
   blExpired = True
   About.Show 1
errlog:
   subErrLog ("Form1:mRegister_Click")
End Sub

Private Sub mRegister2_Click()
   mRegister_Click
End Sub

Private Sub mReload_Click()
On Error GoTo errlog

   Dim Response As Integer
   Response = MsgBox("Do you really want to reload TaskTracker?" + vbNewLine _
   + "(This can take a long time.)", 292)
   If Response = vbYes Then
       subReload
   End If
   
errlog:
   subErrLog ("Form1:mReload_Click")
End Sub

Private Sub mRemovable_Click()
On Error GoTo errlog
   Dim rw As Long
   Dim ri As Long
   Dim lv1 As Integer
   Dim lgx As ListItem
   
   If mRemovable.Checked = True Then
      mRemovable.Checked = False
      
      Dim blAddtheType As Boolean
      For ri = 0 To UBound(RemovableDrives)
         If Len(RemovableDrives(ri)) > 0 Then
            For rw = 0 To UBound(fType)
               If rw = UBound(fType) Then Exit For
               If Left$(aSFileName(rw), 3) = RemovableDrives(ri) Then
                  blNotValidated(rw) = True
                  fType(rw) = tType(rw)
                  
                  'Add any missing icons
                  If Len(fType(rw)) > 0 Then
                     If Left$(fType(rw), 5) <> "hide-" Then
                        blAddtheType = True
                        For lv1 = 1 To ListView1.ListItems.Count
                          If fType(rw) = ListView1.ListItems(lv1).Text Then
                             blAddtheType = False
                             Exit For
                          End If
                        Next
                        If blAddtheType = True Then
                           Call GetNewIcon(rw, fType(rw))
                        End If
                     End If
                  End If
                  
               End If
            Next rw
         End If
      Next ri
      
      For lv1 = 1 To ListView1.ListItems.Count
         If ListView1.ListItems(lv1).Text = "Removable Drive" Then
            ListView1.ListItems.Remove ("Removable Drive")
            Exit For
         End If
      Next lv1
      
      For ri = 0 To UBound(CDDrives)
         If Len(CDDrives(ri)) > 0 Then
            For rw = 0 To UBound(fType)
               If rw = UBound(fType) Then Exit For
               If Left$(aSFileName(rw), 3) = CDDrives(ri) Then
                  blNotValidated(rw) = True
                  fType(rw) = tType(rw)
                  
                  'Add any missing icons
                  If Len(fType(rw)) > 0 Then
                     If Left$(fType(rw), 5) <> "hide-" Then
                        blAddtheType = True
                        For lv1 = 1 To ListView1.ListItems.Count
                          If fType(rw) = ListView1.ListItems(lv1).Text Then
                             blAddtheType = False
                             Exit For
                          End If
                        Next
                        If blAddtheType = True Then
                           Call GetNewIcon(rw, fType(rw))
                        End If
                     End If
                  End If
                  
               End If
            Next rw
         End If
      Next ri
         
      For lv1 = 1 To ListView1.ListItems.Count
         If ListView1.ListItems(lv1).Text = "CD Drive" Then
            ListView1.ListItems.Remove ("CD Drive")
            Exit For
         End If
      Next lv1
   
   Else
   
      Dim blAddRemovable As Boolean
      Dim blRemoveType As Boolean
      
      'First determine which files are on removable drives
      mRemovable.Checked = True
      For ri = 0 To UBound(RemovableDrives)
         If Len(RemovableDrives(ri)) > 0 Then
            For rw = 0 To UBound(fType)
               If rw = UBound(fType) Then Exit For
               If Left$(aSFileName(rw), 3) = RemovableDrives(ri) Then
                  blNotValidated(rw) = True
                  fType(rw) = "Removable Drive"
                  blAddRemovable = True
               End If
            Next rw
         End If
      Next ri
            
      If blAddRemovable = True Then
         
         'Then add the Removable type and delete any empty types
         Set lgx = ListView1.ListItems.Add(, "Removable Drive", "Removable Drive")
         lgx.SmallIcon = ImageList1.ListImages("Removable Drive").Key
         mNetwork.Checked = True
         
         For lv1 = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(lv1).Text <> "Removable Drive" Then
               blRemoveType = True
               For rw = 0 To UBound(fType)
                  If rw = UBound(fType) Then Exit For
                  If Len(fType(rw)) > 0 Then
                     If fType(rw) = ListView1.ListItems(lv1).Text Then
                         blRemoveType = False
                         Exit For
                     End If
                  End If
               Next rw
               If blRemoveType = True Then
                  ListView1.ListItems.Remove (lv1)
               End If
            End If
         Next lv1
      End If
      
      'First determine which files are on CD drives
      Dim blAddCD As Boolean
      For ri = 0 To UBound(CDDrives)
         If Len(CDDrives(ri)) > 0 Then
            For rw = 0 To UBound(fType)
               If rw = UBound(fType) Then Exit For
               If Left$(aSFileName(rw), 3) = CDDrives(ri) Then
                  blNotValidated(rw) = True
                  fType(rw) = "CD Drive"
                  blAddCD = True
               End If
            Next rw
         End If
      Next ri
      
      If blAddCD = True Then
         
         'Then add the CD type and delete any empty types
         Set lgx = ListView1.ListItems.Add(, "CD Drive", "CD Drive")
         lgx.SmallIcon = ImageList1.ListImages("CD Drive").Key
         mNetwork.Checked = True
         
         For lv1 = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(lv1).Text <> "CD Drive" Then
               blRemoveType = True
               For rw = 0 To UBound(fType)
                  If rw = UBound(fType) Then Exit For
                  If Len(fType(rw)) > 0 Then
                     If fType(rw) = ListView1.ListItems(lv1).Text Then
                         blRemoveType = False
                         Exit For
                     End If
                  End If
               Next rw
               If blRemoveType = True Then
                  ListView1.ListItems.Remove (lv1)
               End If
            End If
         Next lv1
      End If
      
      If blAddRemovable = False And blAddCD = False Then
         MsgBox "No removable or CD drives are currently connected, or there are no files to show.", vbExclamation
         Exit Sub
      ElseIf blAddRemovable = False And blAddCD = True Then
         MsgBox "No removable drive is currently connected, or there are no files to show.", vbExclamation
      ElseIf blAddRemovable = True And blAddCD = False Then
         MsgBox "No CD drive is currently connected, or there are no files to show.", vbExclamation
      End If
      
   End If
   
   subClickList
   ListView1_Click

   
errlog:
   subErrLog ("Form1:mRemovable_click")
End Sub

Public Sub subClickList()
On Error GoTo errlog
   If mAlpha.Checked = True Then
      mAlpha_Click
   ElseIf mNumber.Checked = True Then
      mNumber_Click
   ElseIf mLatest.Checked = True Then
      mLatest_Click
   End If
errlog:
   subErrLog ("Form1:subClickList")
End Sub

Private Sub mSaveSize_Click()
On Error GoTo errlog
   
   If mSaveSize.Checked = False Then
      mSaveSize.Checked = True
      subSaveSize
   Else
      If blSaveLayout = False Then
         MsgBox "Default settings will be restored when you restart TaskTracker. ", vbInformation
         blSaveLayout = True
      End If
      mSaveSize.Checked = False
      SaveSetting "TaskTracker", "Settings", "Form1Left", ""
      SaveSetting "TaskTracker", "Settings", "Form1Top", ""
      SaveSetting "TaskTracker", "Settings", "Form1Height", ""
      SaveSetting "TaskTracker", "Settings", "Form1Width", ""
      SaveSetting "TaskTracker", "Settings", "Form2Left", ""
      SaveSetting "TaskTracker", "Settings", "Form2Top", ""
      SaveSetting "TaskTracker", "Settings", "Form2Height", ""
      SaveSetting "TaskTracker", "Settings", "Form2Width", ""
      SaveSetting "TaskTracker", "Settings", "Form4Left", ""
      SaveSetting "TaskTracker", "Settings", "Form4Top", ""
      SaveSetting "TaskTracker", "Settings", "Form5Left", ""
      SaveSetting "TaskTracker", "Settings", "Form5Top", ""
      SaveSetting "TaskTracker", "Settings", "Form5Height", ""
      SaveSetting "TaskTracker", "Settings", "Form5Width", ""
      SaveSetting "TaskTracker", "Settings", "Form6Left", ""
      SaveSetting "TaskTracker", "Settings", "Form6Top", ""
      SaveSetting "TaskTracker", "Settings", "Form6Height", ""
      SaveSetting "TaskTracker", "Settings", "Form6Width", ""
      SaveSetting "TaskTracker", "Settings", "col1width", ""
      SaveSetting "TaskTracker", "Settings", "col2width", ""
'      SaveSetting "TaskTracker", "Settings", "colSwidth", ""
      SaveSetting "TaskTracker", "Settings", "col1Awidth", ""
      SaveSetting "TaskTracker", "Settings", "col1NAwidth", ""
      SaveSetting "TaskTracker", "Settings", "col2Awidth", ""
'      SaveSetting "TaskTracker", "Settings", "col2NAwidth", ""
      SaveSetting "TaskTracker", "Settings", "col3Awidth", ""
      SaveSetting "TaskTracker", "Settings", "col3NAwidth", ""
      SaveSetting "TaskTracker", "Settings", "col1Pwidth", ""
      SaveSetting "TaskTracker", "Settings", "col2Pwidth", ""
      SaveSetting "TaskTracker", "Settings", "col3Pwidth", ""
      SaveSetting "TaskTracker", "Settings", "col1PAwidth", ""
      SaveSetting "TaskTracker", "Settings", "col1NPAwidth", ""
      SaveSetting "TaskTracker", "Settings", "col2PAwidth", ""
      SaveSetting "TaskTracker", "Settings", "col2NPAwidth", ""
      SaveSetting "TaskTracker", "Settings", "col3PAwidth", ""
'      SaveSetting "TaskTracker", "Settings", "col3NPAwidth", ""
      SaveSetting "TaskTracker", "Settings", "col4PAwidth", ""
      SaveSetting "TaskTracker", "Settings", "col4NPAwidth", ""
      SaveSetting "TaskTracker", "Settings", "col4Swidth", ""
   End If
errlog:
   subErrLog ("Form1:mSaveSize")
End Sub

Private Sub mSingleClick_Click()
On Error GoTo errlog
   If mSingleClick.Checked = True Then
      mSingleClick.Checked = False
      mDefaultFocus.Enabled = False
      mAutohide.Enabled = False
   Else
      mSingleClick.Checked = True
      mDefaultFocus.Enabled = True
      mAutohide.Enabled = True
   End If
errlog:
   subErrLog ("Form1:mSingleClick_Click")
End Sub

Private Sub mGlueRight_Click()
On Error GoTo errlog
   If mGlueRight.Checked = True Then
      mGlueRight.Checked = False
      If mFade.Checked = False And mGlueLeft.Checked = False Then
         Timer2.Enabled = False
      End If
   Else
      mGlueRight.Checked = True
'      mSingleClick.Checked = True
      mGlueLeft.Checked = False
      ListView1_Click
      Timer2.Enabled = True
      'adjust if form1 is right of desktop
      If (Form2.Left + Form2.Width + Form1.Width) > wWidth Then
         Form1.Left = wWidth - Form1.Width
      End If
      Form_Resize
   End If
errlog:
   subErrLog ("Form1:mGlueRight_Click")
End Sub

Private Sub mGlueLeft_Click()
On Error GoTo errlog
   If mGlueLeft.Checked = True Then
      mGlueLeft.Checked = False
      If mFade.Checked = False And mGlueRight.Checked = False Then
         Timer2.Enabled = False
      End If
   Else
      mGlueLeft.Checked = True
'      mSingleClick.Checked = True
      mGlueRight.Checked = False
      ListView1_Click
      Timer2.Enabled = True
      'adjust if form1 is left of desktop
      If Form2.Left < Form1.Width Then
         Form1.Left = 0
      End If
      Form_Resize
   End If
errlog:
   subErrLog ("Form1:mGlueLeft_Click")
End Sub

Private Sub mSyncStartup_Click()
On Error GoTo errlog
   If mSyncStartup.Checked = False Then
      Dim Response As Integer
      Response = MsgBox("This option significantly increases startup time but provides better performance subsequently." + vbNewLine _
      + "(It's a good option if you seldom need to restart TaskTracker.) Do you really want to continue?", 292)
      If Response = vbNo Then
         Exit Sub
      End If
      mSyncStartup.Checked = True
      mSyncShow.Checked = False
      mNeverSync.Checked = False
   End If
errlog:
   subErrLog ("Form1:mGlueLeft_Click")
End Sub

Private Sub mSyncShow_Click()
On Error GoTo errlog
   If mSyncShow.Checked = False Then
      mSyncShow.Checked = True
      mSyncStartup.Checked = False
      mNeverSync.Checked = False
   End If
errlog:
   subErrLog ("Form1:mGlueLeft_Click")
End Sub

Private Sub mNeverSync_Click()
On Error GoTo errlog
   If mNeverSync.Checked = False Then
      Dim Response As Integer
      Response = MsgBox("This option reduces startup time but reduces accuracy and is not recommended for most users." + vbNewLine _
      + "Do you really want to continue?", 292)
      If Response = vbNo Then
         Exit Sub
      End If
      mNeverSync.Checked = True
      mSyncShow.Checked = False
      mSyncStartup.Checked = False
   End If
errlog:
   subErrLog ("Form1:mGlueLeft_Click")
End Sub

Private Sub mTaskbar_Click()
On Error GoTo errlog
   If mTaskbar.Checked = False Then
      mTaskbar.Checked = True
      mFade.Enabled = False
      If blTaskbar = False Then
         MsgBox "This will take effect after you exit and restart TaskTracker. " + vbNewLine _
         + "By choosing this option, you can restore TaskTracker by using Alt+Tab.", vbInformation
         blTaskbar = True
      End If
   Else
      mTaskbar.Checked = False
      mFade.Enabled = True
      If blTaskbar = False Then
         blNoSysTray = True
         MsgBox "This will take effect after you exit and restart TaskTracker. ", vbInformation
      End If
   End If
errlog:
   subErrLog ("Form1:mGlueLeft_Click")
End Sub

Private Sub mHideTips_Click()
On Error GoTo errlog
   If mHideTips.Checked = False Then
      mHideTips.Checked = True
   Else
      mHideTips.Checked = False
   End If
errlog:
   subErrLog ("Form1:mGlueLeft_Click")
End Sub

Private Sub mNoSplash_Click()
On Error GoTo errlog
   If mNoSplash.Checked = False Then
      mNoSplash.Checked = True
   Else
      mNoSplash.Checked = False
   End If
errlog:
   subErrLog ("Form1:mGlueLeft_Click")
End Sub

Private Sub mAutoPreview_Click()
On Error GoTo errlog
   
   If mAutoPreview.Checked = False Then
      mAutoPreview.Checked = True
   Else
      mAutoPreview.Checked = False
      Unload Form6
   End If

errlog:
subErrLog ("Form1:mAutoPreview_Click")
End Sub

Public Sub mVirtual_Click()
On Error GoTo errlog

   If mVirtual.Checked = True Then
     mVirtual.Checked = False
     Unload Form5
   Else
     mVirtual.Checked = True
'     Form5.blForm5Load = True
     Form5.Show
     Form5.SetFocus
     If Form2.mFilter.Checked = True Then
         blSearchAll = False
         Form4.Form_Unload (1)
     End If
   End If
   
errlog:
subErrLog ("Form1:mVirtual_click")
End Sub

Public Sub Timer1_Timer()
On Error GoTo errlog

   If GetInputState() <> 0 Then
       DoEvents
   End If
        
   Reposition
   
   If blSetTempCaption = True Then
      blSetTempCaption = False
      Form2.Caption = "TaskTracker - Loading..."
   End If
   
   If blStartNag = True Then
      blStartNag = False
      If Me.Visible = True Then
         About.Show 1
      End If
   End If
      
   If Form5.blVirtualFolderDrag = True Then
      Form5.VirtualFolderDrag
   End If
   
   If blVirtualView = True Then Exit Sub

'*******************************************************
   'timer is used to activate string filter searches to reduce jumpy typing
   '(only when all files are selected)
'   If blSearchAdd = True Then
'      blSearchAdd = False
'      If blLoading = False Then
'         With Form4
'            .SearchAdd
'            Form2.subForm2Init
'            .SetFocus
'            .cbSearch.SelStart = Len(.cbSearch.Text)
'         End With
'      End If
'   End If
'*******************************************************
   
'*******************************************************
   Timer1.Interval = iRefreshTime * 1000     'default ~1 second
'   If Me.WindowState = vbMinimized Or Me.Visible = False Then
'      Timer1.Interval = iRefreshTime * 5000     'default ~5 second
'   Else
'      Timer1.Interval = iRefreshTime * 1000     'default ~1 second
'   End If
'   Debug.Print Str(Format$(Time, "ss"))
'*******************************************************
   
   blWait = False 'this is set each time the timer is called
   
   RandomExpiration
   
'   If RegisterWindowMessage(TaskbarCreatedString) <> TaskBarID Then
'      TaskBarID = RegisterWindowMessage(TaskbarCreatedString)
   If mTaskbar.Checked = False And blNoSysTray = False Then
      subSysTray
   End If
'   End If
   
   If tmCount > 12 Then    'used to mean something (~60s) when timer interval was 5s
      tmCount = 0
      
'*******************************************************
      'date rollover - update idaysused if TT is not exited overnight
      If Timer < SecsAfterMidnight Then   'it must be a new day!
         iDaysUsed = iDaysUsed + 1
         Unload About
         Load About
      End If
      SecsAfterMidnight = Timer
'*******************************************************
      
   End If
   
'*******************************************************
   If blLoading = True Then    'no cache = true
      InitializeTimer
      If Command$ = "-m" Then
         blLoading = False
      End If
   End If
'*******************************************************
    
   tmCount = tmCount + 1
   
   KeepType = sType        'the filetype selected in listview1
      
   frmNotify.CheckNotifyList
            
'*******************************************************
   If RecentFolderChange = True Or FirstFolderCheck = True Then    'this function returns true if Recent folder mod date has changed
      
'      If DoesFileExist(TTpath + "TTTemp") Then     'put this check for another TT running in TTpath and check here to minimize overhead
'         Kill TTpath + "TTTemp"
'         subShowOnly
'      End If
      
      FirstFolderCheck = False
      NormalSequence       'Updates when TT file count has changed
      Exit Sub

   Else
         
      StopState            'user-initiated refresh
      
   End If
'*******************************************************
             
errlog:
subErrLog ("Form1:Timer1_Timer")
End Sub

Private Sub TrimMem()
On Error GoTo errlog

    Dim lRetVal As Long
    Dim hProcess As Long
    Dim lpMinimumWorkingSetSize As Long
    Dim lpMaximumWorkingSetSize As Long
    '
    ' Open the Process to get the Processes Handle
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, GetCurrentProcessId)

    '
    ' If both dwMinimumWorkingSetSize and dwMaximumWorkingSetSize
    ' have the value -1, the function temporarily trims the working
    ' set of the specified process to zero. This essentially swaps
    ' the process out of physical RAM memory.
    lpMinimumWorkingSetSize = -1
    lpMaximumWorkingSetSize = -1
    '
    ' Lets trim the working set now
    lRetVal = SetProcessWorkingSetSize(hProcess, lpMinimumWorkingSetSize, lpMaximumWorkingSetSize)
    '
    ' Remember to close the Handle when we are done with it.
    hProcess = CloseHandle(hProcess)
    
errlog:
subErrLog ("Form1:TrimMem")
End Sub

Private Sub Reposition()
On Error GoTo errlog:
   'reposition according to desktop sidebars, such as Google Desktop
   If mSaveSize.Checked = False Then
      If Me.WindowState = vbNormal Then
         wBottom = flActualScreenHeight(1)
         wWidth = flActualScreenWidth(1)
         If blReposition = True Or LastwBottom <> wBottom Then
            LastwBottom = wBottom
            Me.Top = wBottom - Me.Height
         End If
         If blReposition = True Or LastwWidth <> wWidth Then
            LastwWidth = wWidth
            Me.Left = wWidth - Me.Width
         End If
      End If
   End If
   Timer2_Timer   'sync with glue left/right
errlog:
   blReposition = False
subErrLog ("Form1:Reposition")
End Sub


Private Sub RandomExpiration()
On Error GoTo errlog:
   blStartNag = False
   If blRegStatus = False Then
      If iDaysUsed > 29 Then
         Dim currenttime As String
         Dim currentminutes As Integer
         currenttime = Format$(Time, "h:m:s")
         currentminutes = Val(Mid$(currenttime, InStr(currenttime, ":") + 1, 2))
         If currentminutes = RandomMinutesTillExpiration Then
            Unload About
            blExpired = True
            blStartNag = True     'now wait for window to be visible in timer1
'            blLoading = False
'            About.Show 1
         End If
      End If
   End If
errlog:
subErrLog ("Form1:RandomExpiration")
End Sub

Private Sub LoadSequence()
On Error Resume Next
   
   blBegin = True
      
   subErase
      
   '***********
   GetFileSystem           'needed in case FAT32
   subGetFixedDrives       'determine system drives
   GetShortDateFormat      'get localization format
   '***********
      
   If mCheckUpdates.Checked = True Or VB.App.Comments = "Beta Release" Then
      About.lbUpdate_Click       'needs to follow subStart for TTPath
   End If
   
   subStart
   
   StopState      'initializes the dc value
   
'   subFirstSel
   
   If Command$ <> "-m" Then
      blLoading = False
   End If
   
   blWait = True
   Timer1.Enabled = True
   FirstFolderCheck = True
   
'   Shell_NotifyIcon n, nid    'switch systray caption from loading
   
   If mGlueLeft.Checked = True Or mGlueRight.Checked = True Then
      Timer2.Enabled = True
   End If
   
   TrimMem
   
errlog:
   subErrLog ("Form1:LoadSequence")
End Sub

'clean out all the arrays
Private Sub subErase()
On Error Resume Next

   Erase fType
   Erase tType
   Erase PrevType
   Erase blExists
   Erase blNotValidated
   Erase aLastAcc
   Erase aLastMod
   Erase aCreated
   Erase aLastDate
   Erase aShortcut
   Erase aSFileName
   Erase aNumShort
   Erase aNumPath
   Erase blTrueIcons
   Erase blTrueDates
   Erase KeepSelected
   
   lnkcnt = 0
   fCount = 0
   TTCnt = 0
   pnt = 0
   filecnt = 0
   rf = 0
   vf = 0
   
End Sub

Public Sub NormalSequence()
On Error GoTo errlog
  
   If blWait = False Then                                   'execute update
   
      If Form2.mFilter.Checked = True Then Exit Sub         'unfortunate but necessary
   
      blBegin = True
      blRefresh = True
      blUserRefresh = False
      
      If Len(TTpath) > 0 Then
      
'           blDateSwitch = True
                   
      '    ****************
           subCopyShortcuts (sRecent & "*.lnk")     'calls BeginNewRecentShortcuts > AddFileItemDetails > ...
      '    ****************
            
           subGetFixedDrives       'check for change in system drives
           
           If blNewShortcuts = True Then
                       
               ' ****************
'                 If Form2.blDontPopAgain Then
'                    Form2.blDontPopAgain = False
'                 Else
                    PopFromMem
                    Form2.subLbCount
'                 End If
               ' ****************
            
           End If
      
      End If
      
      blWait = True
      
      If mLatest.Checked = True Then   'autoupdate listview1 with latest-used file type
         mLatest_Click
      End If
   
   End If
   
errlog:
   blNewShortcuts = False
   blRefresh = False
   blLoading = False
   blBegin = False
   subErrLog ("Form1:NormalSequence")
End Sub

Private Sub StopState()
On Error GoTo errlog

'   Dim dc As Integer

   If Len(sType) = 0 Then
      PopsType
   End If
   
'   dc = Form2.ListView2.ListItems.Count
   
   If blLoading = True Or blUserRefresh = True Then
   
      blRefresh = True
      blDateSwitch = True
      PopFromMem
      Form2.subLbCount
'      Form2.blDontPopAgain = True
   
      'only show this message if refresh was user-initiated
      If blUserRefresh = True Then
'        If Form2.ListView2.ListItems.Count = dc Then 'only show if no change (all files still exist)
         blUserRefresh = False      'prevent mysterious double msg by clearing these before
         blRefresh = False
'            MsgBox "No new files found by TaskTracker.", vbInformation
'        End If
      End If
      
   End If
   
errlog:
   blUserRefresh = False
   blRefresh = False
   subErrLog ("Form1:StopState")
End Sub

Public Sub subStart()
On Error GoTo errlog
   'Start with the My Recent Documents folder. First time the app runs it will read the files from this.
   'Once the TaskTracker copy of the Recent shortcuts is made, subsequent starts will read from the \TT folder
   '(which will start accumulating more shortcuts than the Recent folder).
      
   sRecent = fGetSpecialFolder(CSIDL_RECENT)
   
'  ******************
   GetTTPath
'  ******************
   
'  ******************
   RecentFolderChange      'just to initialize dLastWrite
'  ******************

'  ***********
   GetTheFPath
'  ***********

'  ***********
   DateCheck
'  ***********

   If blReloading = False Then
      If blRegStatus = True Or iDaysUsed < 15 Then
         If Command$ <> "-m" Then
            LoadSplash
         End If
      End If
   End If
   
   If blJustRegd = True Then
      Form1.Visible = True
      Form1.WindowState = vbNormal
   End If
   
'  ***********
   If blNewTTDir = True Then
      subCopyShortcuts (sRecent & "*.lnk")     'calls BeginNewRecentShortcuts > AddFileItemDetails > ...
'      subCopyShortcuts ("*.lnk")
   End If
'  ***********
         
   fCount = FilesCountAll(TTpath, "*.lnk")

'  ***************
   If blNoCache = False And blLoadedAllTypes = False And blAbort = False Then
      InstantLoad
      Exit Sub
   End If
'  ***************

'   Unload Splash
     
'  ***************
   GetTheFileList        'calls BeginFileItemDetails > AddFileItemDetails > ...
'  ***************

'  ***************
   If blNoCache = True Then
      SetLoadTTFileData       'Just to retrieve virtual folders
   End If
'  ***************

errlog:
   subErrLog ("Form1:subStart")
End Sub

Private Sub GetTTPath()
   If DoesFolderExist(fGetSpecialFolder(CSIDL_PROFILE) + "TaskTracker\") Then
      TTpath = fGetSpecialFolder(CSIDL_PROFILE) + "TaskTracker\"            'admin
   ElseIf DoesFolderExist(fGetSpecialFolder(CSIDL_LOCAL_APPDATA) + "TaskTracker\") Then
      TTpath = fGetSpecialFolder(CSIDL_LOCAL_APPDATA) + "TaskTracker\"      'limited
   Else  'if neither folder exists, begin by assuming the path will be for admin
      TTpath = fGetSpecialFolder(CSIDL_PROFILE) + "TaskTracker\"            'admin
   End If
End Sub
     
Private Sub GetTheFPath()

   'Running TT first time
   If DoesFolderExist(TTpath) = False Then
      CreateTTDir
      blNewTTDir = True
   End If
           
   If FilesCountAll(sRecent, "*.lnk") = 0 And FilesCountAll(TTpath, "*.lnk") = 0 Then
   
      fnOpenURL ("http://tasktracker.wordwisesolutions.com/support/noshortcuts.htm")

      MsgBox "Exiting application." + vbNewLine + vbNewLine + "TaskTracker did not find any shortcut files in your " _
      + "Windows' Recent Documents Folder for the current user account. TaskTracker begins by building " _
      + "a file history from the contents of this system folder and cannot get started without at least one shortcut file. " _
      + vbNewLine + vbNewLine + "You will probably need to make a change in the system's registry to enable your recent document " _
      + "history for task tracking. Go to http://tasktracker.wordwisesolutions.com/support/noshortcuts for more information." _
      + vbNewLine + vbNewLine + "Ask a system administrator if you're not sure. ", vbCritical
      mPopExit_Click
   End If
   
End Sub

Private Sub CreateTTDir()
On Error Resume Next
  
   blNewTTDir = True

   MkDir TTpath
       
   fCount = 0
   
   If DoesFolderExist(TTpath) = False Then    'if no admin privs, make dir here
      
      TTpath = fGetSpecialFolder(CSIDL_LOCAL_APPDATA) + "TaskTracker\"
      
      MkDir TTpath
      
      fCount = 0
      
      If DoesFolderExist(TTpath) = False Then
         MsgBox "Exiting application." + vbNewLine + vbNewLine + "TaskTracker could not create a working directory " _
         + "due to your account privilege or security settings. " _
         + vbNewLine + "Ask a system administrator if these can be changed. ", vbCritical
         mPopExit_Click
      End If
   
   End If
          
   Call WriteTT(90, Now, 0)
   
errlog:
subErrLog ("Form1:CreateTTDir")
End Sub

Private Sub DateCheck()
On Error GoTo errlog

   Dim iMax As Variant, LastDate As Variant, DaysUsed As Variant
   Dim intF As Integer
   Dim blUnregistered As Boolean
   Dim blCrypt As Boolean
   Dim KeyName As String
   
   GetLocalTime SYS_TIME
      
'  TT file removed***********************************************
   If DoesFolderExist(TTpath) = True Then
      If DoesFileExist(TTpath + "TaskTracker") = False Then
         If FilesCountAll(TTpath, "*.lnk") > 0 Then
            iMax = "1.1eval"
            DaysUsed = 30
            LastDate = SYS_TIME.wDay
            Call WriteTT(iMax, LastDate, DaysUsed)
         End If
      End If
   End If
   
   
   'Cracked test***************************
   'have to open and close the file separately for this, or breaks the Read TT step next
   intF = FreeFile
   Open TTpath + "TaskTracker" For Binary Access Read As #intF
     
   Dim all As String, test As String
   all = Input(LOF(intF), intF)
   test = Mid$(all, 15, 1)
   If AscB(test) = 16 Then
      MsgBox "TaskTracker has detected an illegally cracked, pirated registration file on this computer." + vbNewLine _
           + "Please mail support@wordwisesolutions.com if you paid for TaskTracker." + vbNewLine _
           + "You may register TaskTracker at http://tasktracker.wordwisesolutions.com/update/", vbCritical
      End
   End If
   
   Close #intF
   '***************************************

   
'  Read TT file***********************************************
   intF = FreeFile
   Open TTpath + "TaskTracker" For Binary Access Read As #intF
     
      Get #intF, , iMax
      Get #intF, , LastDate
      Get #intF, , DaysUsed
      If iMax = "noexpiry" Then     'legacy
         Get #intF, , RegType
         Get #intF, , RegDate
         Get #intF, , RegName
      Else
         If Right$(iMax, 1) = Chr$(128) Then
            iMax = DeCrypt(iMax, "Status 10041958 PseudoCrypt")
            blCrypt = True
            If iMax = "noexpiry" Then
               Get #intF, , RegType
               Get #intF, , RegDate
               Get #intF, , RegName
            End If
         End If
      End If
   Close #intF
     
   If blCrypt = True Then
      If iMax = "noexpiry" Then
         KeyName = DeCrypt(RegName, "RegName 10041958 PseudoCrypt")
         RegName = KeyName
      Else
         KeyName = "No RegName 10041958 PseudoCrypt"
      End If
      LastDate = DeCrypt(LastDate, KeyName)
      DaysUsed = DeCrypt(DaysUsed, KeyName)
      RegType = DeCrypt(RegType, KeyName)
      RegDate = DeCrypt(RegDate, KeyName)
   End If
      
   iDaysUsed = Val(DaysUsed)
         
'  Registered user***********************************************
   If RegType = "C" Or RegType = "S" Or RegType = "B" Or RegType = "P" Or RegType = "T" Then
      blRegStatus = True
      If Len(GetSetting("TaskTracker", "Settings", "Revision")) > 0 Then
         If Trim$(GetSetting("TaskTracker", "Settings", "Revision")) <> Trim$(Str$(VB.App.Revision)) Then

            iDaysUsed = 0   'downloaded new version
         Else
            If LastDate <> SYS_TIME.wDay Then
               iDaysUsed = iDaysUsed + 1
            End If
         End If
      Else
         iDaysUsed = 0
      End If
      
      DaysUsed = iDaysUsed
      LastDate = SYS_TIME.wDay
      
      Call RegWrite(iMax, LastDate, DaysUsed)
      
      SaveSetting "TaskTracker", "Settings", "Revision", Trim$(Str$(VB.App.Revision))
                  
   Else
      blUnregistered = True
   End If
   
   'Reset Temporary license
   If RegType = "T" Then
      If DateDiff("d", RegDate, Now) > 30 Then
         blUnregistered = True
         RegType = vbNullString
         iMax = "1.1eval"
         Call WriteTT(iMax, LastDate, DaysUsed)
      End If
   End If
      
   'Once only - reset days used to 0 if V1 (or new) user*************
   If iMax <> "1.1eval" Then
      Dim blOnceOnly As Boolean
      iMax = "1.1eval"
      blOnceOnly = True
   End If
   
'  Evaluate once a day***********************************************
   If LastDate <> SYS_TIME.wDay Then
   
      If blOnceOnly = True Then
         iDaysUsed = 0
      Else
         iDaysUsed = iDaysUsed + 1     'increment days used
      End If
      
      DaysUsed = iDaysUsed
      
      LastDate = SYS_TIME.wDay

      Call WriteTT(iMax, LastDate, DaysUsed)
      
      SaveSetting "TaskTracker", "Settings", "Revision", Trim$(Str$(VB.App.Major)) _
      + Trim$(Str$(VB.App.Minor)) + Trim$(Str$(VB.App.Revision))

   End If
   
   'Load Nag Splash after 14 days*********************
    If blReloading = False Then
      If blUnregistered = True Then
         If iDaysUsed > 14 Then
            blLoadModal = True
            If Command$ <> "-m" Then
               LoadSplash
            End If
         End If
      End If
   End If

errlog:
subErrLog ("Form1:DateCheck")
End Sub

Private Sub WriteTT(iMax As Variant, LastDate As Variant, DaysUsed As Variant)
   Dim intF As Integer
   Dim piMax As Variant, pLastDate As Variant, pDaysUsed As Variant
   Dim KeyName As String
   
   intF = FreeFile
   
   If blRegStatus = True Then
      KeyName = RegName
   Else
      KeyName = "No RegName 10041958 PseudoCrypt"
   End If
   
   piMax = CVar(EnCrypt(iMax, "Status 10041958 PseudoCrypt")) + Chr$(128)
   pLastDate = CVar(EnCrypt(LastDate, KeyName))
   pDaysUsed = CVar(EnCrypt(DaysUsed, KeyName))
   
   Open TTpath + "TaskTracker" For Binary Access Write As #intF
      Put #intF, , piMax
      Put #intF, , pLastDate
      Put #intF, , pDaysUsed
   Close #intF
End Sub

Public Sub RegWrite(iMax As Variant, LastDate As Variant, DaysUsed As Variant)
   Dim intF As Integer
   Dim piMax As Variant, pLastDate As Variant, pDaysUsed As Variant
   Dim pRegType As Variant, pRegDate As Variant, pRegName As Variant

   piMax = CVar(EnCrypt(iMax, "Status 10041958 PseudoCrypt")) + Chr$(128)
   pLastDate = CVar(EnCrypt(LastDate, RegName))
   pDaysUsed = CVar(EnCrypt(DaysUsed, RegName))
   pRegType = CVar(EnCrypt(RegType, RegName))
   pRegDate = CVar(EnCrypt(RegDate, RegName))
   pRegName = CVar(EnCrypt(RegName, "RegName 10041958 PseudoCrypt"))

   intF = FreeFile
   Open TTpath + "TaskTracker" For Binary Access Write As #intF
   
      Put #intF, , piMax
      Put #intF, , pLastDate
      Put #intF, , pDaysUsed
      Put #intF, , pRegType
      Put #intF, , pRegDate
      Put #intF, , pRegName
   Close #intF
End Sub

Public Sub Form_Load()
On Error Resume Next
     
   If App.PrevInstance Then
      Timer1.Enabled = False
      WakeUpApp Me
   End If
   
   StartLogging
            
   blLoading = True
   blThisType = True
   blFirstLoad = True
   wBottom = flActualScreenHeight(1)
   wWidth = flActualScreenWidth(1)
  
   GetSettings
   
   Recovery
   
   TaskbarStatus
         
   'default start by tray
    If mSaveSize.Checked = False Then
       Me.Top = wBottom - Me.Height
       Me.Left = wWidth - Me.Width
    End If
  
   Set TT = New CTooltip
   
   If mTaskbar.Checked = False And blNoSysTray = False Then
      subSysTray
   End If
   
   subLoadIcons
   
   If mIcons.Checked = True Then
      IconsOnly
   End If
   
   If blNoCache = False And blInstantStarted = False Then
      LoadSequence
      InitializeTimer
   End If
            
   Timer1.Enabled = True  'Timer1 starts things rolling if not cached start...
   Me.Show 0
   If mNoSplash.Checked = False Then
      Splash.SetFocus
   End If
   
'   If blNoCache = False Then
      If mSingleClick = True Then
         ListView1_Click
      End If
'   End If

    frmNotify.Hide
   
errlog:
   subErrLog ("Form1:Form_Load")
End Sub

Private Sub WakeUpApp(Main As Form)
On Error GoTo exitapp
    Dim hnd As Long
    Dim ScanF24 As Long
    Dim Title As String

    Title = App.FileDescription
    Main.Caption = "... duplicate instance ..."

    hnd = GetWindow(Main.hWnd, GW_HWNDFIRST)
    Do While hnd <> 0
        If WakeUpComp(hnd, Title) Then

            ScanF24 = MapVirtualKey(VK_F24, 0) * 65536
            SendMessage hnd, WM_KEYDOWN, VK_F24, 1 Or ScanF24 Or KD_EXTENDED Or KD_DOWN
            SendMessage hnd, WM_KEYUP, VK_F24, 1 Or ScanF24 Or KD_EXTENDED Or KD_UP
        End If
        hnd = GetWindow(hnd, GW_HWNDNEXT)
    Loop
exitapp:
    End
End Sub

Private Function WakeUpComp(hnd As Long, Title) As Boolean
    Dim buf As String
    Dim s As String

    WakeUpComp = False

    buf = String(255, 0)
    GetWindowText hnd, buf, Len(buf)
    s = ApiTextStrip(buf)
    If InStr(1, s, Title, vbTextCompare) <= 0 Then
        Exit Function
    End If

    buf = String(255, 0)
    GetClassName hnd, buf, Len(buf)
    s = ApiTextStrip(buf)
    If InStr(1, s, "Thunder", vbTextCompare) <= 0 Then
        Exit Function
    End If

    WakeUpComp = True
End Function

'Private Sub ActivatePrevInstance()
'On Error GoTo exitapp
'
'   Dim PrevHndl As Long
'   Dim result As Long
'
'   App.Title = "closing..."
'
'   PrevHndl = FindWindow("ThunderRT6Main", "TaskTracker")
'
'   result = SetForegroundWindow(PrevHndl)
''   result = BringWindowToTop(PrevHndl)
''   result = ShowWindow(PrevHndl, SW_SHoW)
''   result = ShowWindow(PrevHndl, SW_SHoWNORMAL)
''   result = ShowWindow(PrevHndl, SW_RESTORE)
'
''   If LenB(GetSetting("TaskTracker", "Settings", "Taskbar")) = 0 Or _
''     CBool(GetSetting("TaskTracker", "Settings", "Taskbar")) = False Then
''     MsgBox "TaskTracker is already running. Use the system tray icon to restore it." _
''     + vbNewLine + "(To avoid this message, choose the " + Chr$(34) + "Show in Taskbar" + Chr$(34) _
''     + " option.)", vbInformation, "TaskTracker"
'''     + vbNewLine + "If no system tray icon is visible, use the Windows Task Manager " + Chr$(34) + "End Process" + Chr$(34) _
'''     + vbNewLine + "command to shut down TaskTracker."
''   End If
'   SendKeys "{F16}"
'
'   blExit = True
'   blEndNewInstance = True
'
'exitapp:
'
'   End
'
'End Sub

'Private Sub subWriteTTTemp()
'On Error GoTo errlog
'
'   Dim intF As Integer
'   Dim TTTemp As String
'
'   GetTTPath
'
'   TTTemp = TTpath + "TTTemp"
'   intF = FreeFile
'   Open TTTemp For Append Access Write As #intF
'
''      Print #intF, ""
'
'errlog:
'   Close #intF
'   subErrLog ("Form1:InstantLoad")
'End Sub

Private Sub LoadSplash()
On Error GoTo errlog

   Timer2.Enabled = False     'bugfix

   If LenB(GetSetting("TaskTracker", "Settings", "NoSplash")) > 0 Then
      mNoSplash.Checked = CBool(GetSetting("TaskTracker", "Settings", "NoSplash"))
   Else
      mNoSplash.Checked = False 'default
   End If
   If mNoSplash.Checked = True Then
      If blRegStatus = True Or (blRegStatus = False And iDaysUsed < 15) Then
         Exit Sub
      End If
   End If
   
   ImageList1.ListImages.Add 1, "splash", LoadResPicture("SPLASH", 0)
   With Splash
      .Timer1.Enabled = True
      .picSplash.Picture = ImageList1.ListImages("splash").Picture
      If blLoadModal = True Then
        If iDaysUsed = 1 Then
           .Label1.Caption = "You Have Been Using TaskTracker for " + Trim$(Str$(iDaysUsed)) + " Day"
        Else
           .Label1.Caption = "You Have Been Using TaskTracker for " + Trim$(Str$(iDaysUsed)) + " Days"
        End If
        .Show 1
      Else
        .Show
      End If
      .Refresh
  End With
  
errlog:
   Timer2.Enabled = True
   subErrLog ("Form1:LoadSplash")
End Sub

Private Sub InstantLoad()
On Error GoTo errlog

   LoadTTFiles
   
   If blAbort = True Then
      Exit Sub
   End If
      
   LoadTTTypesData
   
   SyncFileIcons
   
   DeleteSetting "TaskTracker", "Settings", "Recovery"
   
errlog:
   subErrLog ("Form1:InstantLoad")
End Sub

Private Sub TaskbarStatus()
On Error GoTo errlog

   If mTaskbar.Checked = False Then
      SetWindowLong Me.hWnd, GWL_EXSTYLE, (GetWindowLong(hWnd, _
        GWL_EXSTYLE) And Not WS_EX_APPWINDOW)
      SetWindowLong Form2.hWnd, GWL_EXSTYLE, (GetWindowLong(hWnd, _
        GWL_EXSTYLE) And Not WS_EX_APPWINDOW)
      blTaskbarStatus = False    'use this instead of mTaskbar check to be consistent with current config, rather than "next time" config
   Else
      Dim ll_Style As Long
      SetWindowLong Me.hWnd, GWL_EXSTYLE, (GetWindowLong(hWnd, _
        GWL_EXSTYLE) Or WS_EX_APPWINDOW)
      'Get window style
      ll_Style = GetWindowLong(Me.hWnd, GWL_STYLE)
      'Add the minimize button
      Call SetWindowLong(Me.hWnd, GWL_STYLE, ll_Style Or WS_MINIMIZEBOX)
      blTaskbarStatus = True
   End If

errlog:
   subErrLog ("Form1:TaskbarStatus")
End Sub

Private Sub Recovery()
On Error GoTo errlog

   If LenB(GetSetting("TaskTracker", "Settings", "Recovery")) = 1 Then
      Dim Response As Integer
      Response = MsgBox("TaskTracker encountered an error the last time it started." + vbNewLine + vbNewLine _
      + "Do you want to rebuild its cache? (This can take a long time.)", 292)
      If Response = vbYes Then
         blScratchCache = True
         blNewTTDir = True
      End If
   Else
      If Command$ = "-n" Then
         blScratchCache = True
      End If
   End If
   SaveSetting "TaskTracker", "Settings", "Recovery", "1"
   
errlog:
   subErrLog ("Form1:Recovery")
End Sub

Private Sub InitializeTimer()
On Error GoTo errlog
   Timer1.Enabled = True
   Timer1.Interval = iRefreshTime * 1000     'default ~1 second
'   Timer1_Timer      'bugfix
   If (blNoCache = True And blScratchCache = False) Or blBegin = True Then
      LoadSequence
      If mSingleClick = True Then
         ListView1_Click
      End If
   End If
errlog:
   subErrLog ("Form1:InitializeTimer")
End Sub

Private Sub subLoadIcons()
On Error GoTo errlog
   With ImageList1
      .ImageHeight = 16
      .ImageWidth = 16
      .ListImages.Add 1, "Network Drive", LoadResPicture("NET", 1)
      .ListImages.Add 2, "Deleted or Renamed", LoadResPicture("DEL", 1)
      .ListImages.Add 3, "CD Drive", LoadResPicture("CD", 1)
      .ListImages.Add 4, "Removable Drive", LoadResPicture("REM", 1)
      .ListImages.Add 5, "Unknown", LoadResPicture("UNK", 1)
'      .ListImages.Add 6, "TT16", LoadResPicture("TT", 1)     'causes distortion!
'      .ListImages.Add 7, "Virtual", LoadResPicture("VIR", 1)  'causes distortion!
   End With
   With ImageList2
      .ImageHeight = 16
      .ImageWidth = 16
      .MaskColor = ILC_MASK Or ILC_COLOR8
      .ListImages.Add 1, "UP", LoadResPicture("UP", 1)
      .ListImages.Add 2, "DOWN", LoadResPicture("DOWN", 1)
   End With
errlog:
   subErrLog ("Form1:subLoadIcons")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errlog

   Dim blF10 As Boolean
   Dim blF5 As Boolean
   Dim alttest As Integer

'   If blCtrlKey = True And KeyCode = vbKeyT Then
'      mHideShow_Click (0)
'      Exit Sub
'   End If
   If KeyCode = vbKeyF5 Then
      blF5 = True
   End If
   If blF5 = True And blShift = False Then
      mRefresh_Click
      Exit Sub
   End If
   If mVirtual.Checked = False Then
      If blF5 = True And blShift = True Then
         mReload_Click
         Exit Sub
      End If
   End If
   If KeyCode = vbKeyReturn Then
      mOpenType_Click (1)
   End If
   If KeyCode = vbKeyControl Then
      blCtrlKey = True
      blMultSelect = True
   End If
   If KeyCode = vbKeyShift Then
      blMultSelect = True
      blShift = True
   End If
   If KeyCode = vbKeyF10 Then
      blF10 = True
   End If
   If blShift = True And blF10 = True Then
      subRC1Menu
      blF10 = False
   End If
   If KeyCode = vbKeyF1 Then
      mHelp_Click
   End If
   If KeyCode = vbKeyEscape Then
      Call Form_QueryUnload(1, 0)
   End If
   If KeyCode = vbKeyF16 Then
      subShowOnly
   End If
   If KeyCode = vbKeyF24 Then
      subShowOnly
   End If
   
   alttest = Shift And 7
   If alttest = 4 Then
      blAltPress = True
'      use as toggle to prevent fade when fade option is active
      If Form1.mFade.Checked = True Then
         If blSuspendFade = False Then
            blSuspendFade = True
         Else
            blSuspendFade = False
         End If
      End If
   Else
      blAltPress = False
   End If
   
errlog:
   blF10 = False
   blF5 = False
   subErrLog ("Form1:KeyDown")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo errlog

   If blLoading = True Then Exit Sub
   
   If KeyAscii = 1 Then    'ctrl+A - select all
      mSelectAll_Click
   End If
'   If KeyAscii = 16 Then  'ctrl+P
'      mAlpha_Click
'   End If
'   If KeyAscii = 13 Then  'ctrl+M
'      mLatest_Click
'   End If
'   If KeyAscii = 17 Then  'ctrl+Q
'      mNumber_Click
'   End If
   If KeyAscii = 6 Then    'ctrl+F
      Form2.mFilter_Click
   End If
   If KeyAscii = 14 Then  'ctrl+n
      mPreferences_Click
   End If
   If KeyAscii = 19 Then   'ctrl+S     - 3 way
      If mLatest.Checked = True Then
        mAlpha_Click
      ElseIf mAlpha.Checked = True Then
        mNumber_Click
      Else
        mLatest_Click
      End If
   End If
   If KeyAscii = 2 Then    'ctrl+B
      mPopAbout_Click
   End If
   If KeyAscii = 21 Then  'ctrl+U
      mVirtual_Click
   End If
   If KeyAscii = 24 Then  'ctrl+X
      mPopExit2_Click
   End If
   
   blCtrlKey = False
   
errlog:
   subErrLog ("Form1:KeyPress")
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo errlog
   If blAddType = True Then
      SelectionStatus
      ListView1_Click
   End If
   blShift = False
   If KeyCode = vbKeyF24 Then
      subShowOnly
   End If
errlog:
subErrLog ("Form1:Form_KeyUp")
End Sub

Public Sub SelectionStatus()
On Error GoTo errlog
   Dim lv1 As Integer, ts As Integer
   For lv1 = 1 To Form1.ListView1.ListItems.Count
      If ListView1.ListItems(lv1).Selected = True Then
         ListView1.ListItems(lv1).EnsureVisible
         ts = ts + 1
      End If
   Next lv1
   If ts < 2 Then
'      blShiftKey = False
      blCtrlKey = False
      blMultSelect = False
      If mSelectAll.Checked = False Then
         blSelectAll = False
      End If
   End If
errlog:
subErrLog ("Form1:SelectionStatus")
End Sub

Public Sub subSysTray()
On Error GoTo errlog
   
   With nid
   
      .cbSize = Len(nid)
      .hWnd = Me.hWnd
      .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
      .uCallBackMessage = WM_MOUSEMOVE
      .hIcon = ImageList3.ListImages("TT16").ExtractIcon
      
      If blLoading = False Then
         .szTip = "TaskTracker" & vbNullChar
      Else
      
         If Form1.Visible = False Then
            If blNoCache = False And blReloading = False Then   'cached data
               If blNoSplash = True Then                        'no splash
                  If fc > 0 And TotalFiles > 0 Then
                     If fc Mod 10 = 0 Then
                        ProgPercent = Int(fc / TotalFiles * 100)
                        If ProgPercent > 99 Then
                           ProgPercent = 99
                        End If
                     End If
                  End If
               End If
            End If
            If blReloading = True Or blNoCache = True Then
               If lnkcnt > 0 And fCount > 0 Then         'getfiles data
                  If lnkcnt Mod 10 = 0 Then
                     ProgPercent = Int(lnkcnt / fCount * 100)
                     If ProgPercent > 99 Then
                        ProgPercent = 99
                     End If
                  End If
               End If
            End If
         End If
         
         If Form1.Visible = True Or ProgPercent = 0 Then
            .szTip = "TaskTracker - Loading..." & vbNullChar
         Else
            .szTip = "TaskTracker - Loading..." & Str$(Trim$(ProgPercent)) + "%" & vbNullChar
         End If
         
      End If
      
   End With
   
   Shell_NotifyIcon NIM_MODIFY, nid
   Shell_NotifyIcon NIM_ADD, nid
   
   Exit Sub
   
errlog:
   nid.szTip = "TaskTracker" & vbNullChar
   subErrLog ("Form1:subSysTray")
End Sub

'Private Sub cmdChangeTooltip()
'    nid.szTip = "A new tooltip text" & vbNullChar
'    Shell_NotifyIcon NIM_MODIFY, nid
'End Sub

Private Sub Form_Mouseup(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo errlog
   If blLoading = False Then
      If Button = 2 Then
         subRC1Menu
      Else
         blRightClick = False
      End If
   End If
errlog:
subErrLog ("Form1:Form_Mouseup")
End Sub

Public Sub Form_Resize()
On Error Resume Next

   Form1Sizing       'yes, this is called again in this procedure - this code is a mess
   
   If blMinStart = False And blBegin = False Then
      If blForm2Resize = True Then Exit Sub     'designed to stop feedback loop with form2
      If blLoading = True Then
         blLoadResize = True
         SysTrayLoad
         Exit Sub
      Else
         SysTrayStatus
         blForm1Resize = True
      End If
   Else
      If blLoading = False Then
         subMinFix
      End If
   End If

'   If blLoading = False Or blReloading = True Then
   If blMinStart = False Then
      If mSingleClick.Checked = True Then
         If Me.Visible = False Or Me.WindowState = vbMinimized Then
            TrimMem
            HideForms
         Else
            subSingleClick
         End If
      Else
         If Me.Visible = False Or Me.WindowState = vbMinimized Or _
            Form2.Visible = False Then
            TrimMem
         End If
      End If
   End If
   
   Form1Sizing
   
'   If blLoading = False Or blReloading = True Then
   If blMinStart = False Then
      If Form1.WindowState <> vbMinimized And Form2.WindowState <> vbMinimized Then
         If Form1.mGlueLeft.Checked = True Or Form1.mGlueRight.Checked = True Then
            Form2.Height = Form1.Height
            Form2.Top = Form1.Top
         End If
      End If
   End If
   Me.Refresh
   
   blForm1Resize = False
   
errlog:
'   If Err.Number <> 0 And Err.Number <> 384 Then subErrLog ("Form1:Resize")
   subErrLog ("Form1:Resize")
End Sub

Public Sub HideForms()
On Error GoTo errlog
   Form2.Hide
   Form4.Hide
   If Form1.mVirtual.Checked Then
      Form5.Hide
   End If
   If blForm6Load Then
      Form6.Hide
   End If
errlog:
'   If Err.Number <> 0 And Err.Number <> 384 Then subErrLog ("Form1:HideForms")
   subErrLog ("Form1:HideForms")
End Sub

Private Sub subSingleClick()
On Error GoTo errlog

   If mSingleClick.Checked = True Or blForm2Show = True Then
      blForm2Show = False
''      If blReloading = True Then
''         Form1.Show
''      Else
'       Form2.Show
''      End If
      If blLoading = False Then
         Form2.Visible = True
         Form2.WindowState = vbNormal
      End If
      ReshowOtherWindows
   End If
   
errlog:
'   If Err.Number <> 0 And Err.Number <> 384 Then subErrLog ("Form1:subSingleClick")
   subErrLog ("Form1:subSingleClick")
End Sub

Private Sub Form1Sizing()
On Error GoTo errlog

   If Me.Visible <> False And Me.WindowState <> vbMinimized Then
   
'      If mSaveSize.Checked = True Then
'         SaveFormPositions
'      End If
      
      If mIcons.Checked = False Then
      
         If Me.Width > 4000 Then
            Me.Width = 4000
         End If
                        
         If Me.Width < 1800 Then
            Label1.Width = 1500
         Else
            Label1.Width = Me.Width - 300
         End If
         
         If Me.Width > 2530 Then
            Label1.Top = Me.Height - 1325
            Label1.Height = 500
         Else
            Label1.Top = Me.Height - 1425
            Label1.Height = 750
         End If
         
         With ListView1
            .Width = Me.Width - 100
            .Height = Form1.Height - 1500
            If ListView1.ColumnHeaders.Count > 0 Then
               VertScrollbar
            End If
         End With

         ProgressBar1.Width = Me.Width - 300
         ProgressBar1.Top = ListView1.Height + 250
            
      Else
      
         Me.Width = 600    'this is what I want, but the real min is 900
'         ListView1.Width = Me.Width
         ListView1.Height = Me.Height
         
      End If
   End If
   
errlog:
   subErrLog ("Form1:Form1Sizing")
End Sub

Private Sub SysTrayLoad()
On Error GoTo errlog
  'special case for running minimized on win startup
   If Me.Visible = True Then
      If Me.WindowState = vbMinimized Then
         If mTaskbar.Checked = False And blNoSysTray = False Then
            Me.Hide
         End If
         Form2.Hide
      End If
   End If
errlog:
   subErrLog ("Form1:SysTrayLoad")
End Sub

Private Sub VertScrollbar()
On Error GoTo errlog
   If (GetWindowLong(ListView1.hWnd, GWL_STYLE) And WS_VSCROLL) = False Then
      ListView1.ColumnHeaders(1).Width = ListView1.Width - 350
   Else
      ListView1.ColumnHeaders(1).Width = ListView1.Width - 600
   End If
errlog:
   subErrLog ("Form1:VertScrollbar")
End Sub

Public Sub SysTrayStatus()
On Error GoTo errlog
   If blTaskbarStatus = False Then
      If Me.Visible = False Then
         mHideShow(1).Caption = "&Show TaskTracker"
      Else
         mHideShow(1).Caption = "&Hide TaskTracker"
      End If
   End If
errlog:
   subErrLog ("Form1:SysTrayStatus")
End Sub

Public Sub Form_Unload(Cancel As Integer)
On Error Resume Next
   
   Shell_NotifyIcon NIM_DELETE, nid
   blExiting = True
   
   If blEndNewInstance = False Then
      SaveSettings
      Me.Hide
      Form2.Hide
      If Form2.mFilter.Checked = True Then
         Form4.Hide
      End If
      If Form1.mVirtual.Checked = True Then
         Form5.Hide
      End If
      If blPreviewOpen = True Then
         Form6.Hide
      End If
      About.Hide
   End If
   
   Unload Form2
   Unload frmNotify
   Unload Me
   If Me.mVirtual.Checked = False Then
      SaveTTInfo
   Else
      Unload Form5
   End If
'   Unload Form4
'   Unload Form6
'   Unload About

   End
End Sub

Public Sub InitListview1()
On Error GoTo errlog

   Dim clm As ColumnHeader
      
   With ListView1
      .View = lvwReport
      .ColumnHeaders.Clear
      If mIcons.Checked = False Then
         Set clm = .ColumnHeaders.Add(, sType, "File Types", ListView1.Width - 400)
      Else
         Set clm = .ColumnHeaders.Add(, sType, "File Types", 325)
      End If
      Set clm = .ColumnHeaders.Add(, , "Frequency", 0)  'hidden
      Set clm = .ColumnHeaders.Add(, , "Most Recent", 0)  'hidden
      Set clm = .ColumnHeaders.Add(, , , 0) 'hidden
      Set .SmallIcons = ImageList1
            
      If blRefresh = True Or blLoading = True Then
      
         .SortKey = TypeSortCol
         If .SortKey = 1 Then           'frequency
            .SortOrder = lvwDescending
            mNumber.Checked = True
            mAlpha.Checked = False
            mLatest.Checked = False
            
         ElseIf .SortKey = 0 Then       'alpha
            .SortOrder = lvwAscending
            mAlpha.Checked = True
            mNumber.Checked = False
            mLatest.Checked = False
         
         ElseIf .SortKey = 2 Then       'most recent
            .SortOrder = lvwAscending
            mAlpha.Checked = False
            mNumber.Checked = False
            mLatest.Checked = True
            
         End If
         
         subTypeSort
                  
      End If
            
   End With
    
   Exit Sub
   
errlog:
subErrLog ("Form1:InitListview1")
End Sub

Private Sub BeginFileItemDetails(WFD As WIN32_FIND_DATA)
On Error GoTo errlog

   Dim shFilename As String
   Dim sfilename As String
   
'   If blInstantOn = True And blInstantStarted = False Then
''     use the shfilename
'   Else
      shFilename = TrimNull(WFD.cFileName)
'   End If
   
   If shFilename = "." Or shFilename = ".." Then
      Exit Sub
   End If
   
   If blSnapshot = True Then
       Call subStatusLog("TT shortcut: " + shFilename)
   End If
   
   sType = vbNullString
   
   sfilename = fnResolveLink(TTpath + shFilename)
                   
''   Debug.Print sFilename
'   If Len(sFilename) > 0 Then
'      If sFilename = sLastFileName Then
'         sLastFileName = sFilename
'         Kill TTpath + shFilename      'delete the dupe shortcut
'         Exit Sub
'      End If
'   End If
'   sLastFileName = sFilename
     
   Call AddFileItemDetails(shFilename, sfilename)
   
errlog:
subErrLog ("Form1:BeginFileItemDetails")
End Sub
   
Public Sub AddFileItemDetails(ByVal shFilename As String, ByVal sfilename As String)
On Error GoTo errlog

   Dim di As Integer
   
   ReDim Preserve fType(TTCnt)
   ReDim Preserve blExists(TTCnt)
   ReDim Preserve blNotValidated(TTCnt)
'   ReDim Preserve aFType(TTCnt)
   ReDim Preserve aSFileName(TTCnt)
   ReDim Preserve tType(TTCnt)
'   ReDim Preserve aSmallIcon(TTCnt)
   ReDim Preserve aShortcut(TTCnt)
   
   'File by file, work through all the drive letters, starting with SystemDrives, to determine
 'whether to validate the file automatically or to put into one of unvalidated categories
 '(NetworkDrives, RemovableDrives, CDDrives)
         
   If iSysDrv > 0 Then
      For di = 0 To UBound(SystemDrives)
      
         If ParseSystemFile(SystemDrives(di), sfilename) Then
            If blSnapshot = True Then
               Call subStatusLog("System: " + shFilename + ", " + sfilename)
            End If
            If GetInputState() <> 0 Then
                DoEvents
            End If
            Call ContinueWithDetails(shFilename, sfilename)
            Exit Sub
         End If
         
      Next di
   End If
  
   If iNetDrv > 0 Then
      For di = 0 To UBound(NetworkDrives)
   
         If ParseNetworkFile(NetworkDrives(di), shFilename, sfilename) Then
            If blSnapshot = True Then
               Call subStatusLog("Network: " + shFilename + ", " + sfilename)
            End If
            If GetInputState() <> 0 Then
                DoEvents
            End If
            Call ContinueWithDetails(shFilename, sfilename)
            Exit Sub
         End If
   
      Next di
   End If
   
   If iRemDrv > 0 Then
      For di = 0 To UBound(RemovableDrives)
   
         If ParseRemovableFile(RemovableDrives(di), shFilename, sfilename) Then
            If blSnapshot = True Then
               Call subStatusLog("Removable: " + shFilename + ", " + sfilename)
            End If
            If GetInputState() <> 0 Then
                DoEvents
            End If
            Call ContinueWithDetails(shFilename, sfilename)
            Exit Sub
         End If
              
      Next di
   End If
   
   If iCDDrv > 0 Then
      For di = 0 To UBound(CDDrives)
      
         If ParseCDFile(CDDrives(di), shFilename, sfilename) Then
            If blSnapshot = True Then
               Call subStatusLog("CD: " + shFilename + ", " + sfilename)
            End If
            If GetInputState() <> 0 Then
                DoEvents
            End If
            Call ContinueWithDetails(shFilename, sfilename)
            Exit Sub
         End If
         
      Next di
   End If
        
errlog:
subErrLog ("Form1:AddFileItemDetails")
End Sub

Private Function ParseSystemFile(SystemDrive As String, sfilename As String) As Boolean
On Error GoTo errlog
   
   sfilename = Trim$(TrimNull(sfilename))
   
   Call NullLetterFix(SystemDrive, sfilename)
   
   If Left$(sfilename, 3) = SystemDrive Then
   
      'here I give the filename and get the icon associated with it and the file type
      himgSmall& = SHGetFileInfo(sfilename, _
                          0&, shinfo, Len(shinfo), _
                          BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
                          
      fType(TTCnt) = Trim$(TrimNull(shinfo.szTypeName))
      
      If Len(fType(TTCnt)) = 0 Then
         fType(TTCnt) = EmptySystemType(sfilename)
      End If
      
      If Len(fType(TTCnt)) > 0 Then
         vf = vf + 1    'counter for About box - just a raw count of shortcut "items"
      End If
           
      tType(TTCnt) = vbNullString
      
      blNotValidated(TTCnt) = False
 
      ParseSystemFile = True
      
   Else
   
      ParseSystemFile = False
   
   End If

errlog:
   subErrLog ("Form1:ParseSystemFile")
End Function

Private Function ParseNetworkFile(NetworkDrive As String, shFilename As String, sfilename As String) As Boolean
On Error GoTo errlog

   sfilename = Trim$(TrimNull(sfilename))
   
   Call NullLetterFix(NetworkDrive, sfilename)

   If Left$(sfilename, 3) = NetworkDrive Or Left$(sfilename, 2) = "\\" Then
   
      himgSmall& = SHGetFileInfo(sfilename, _
                          0&, shinfo, Len(shinfo), _
                          BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
                          
       tType(TTCnt) = Trim$(TrimNull(shinfo.szTypeName))
       
      If Len(fType(TTCnt)) = 0 Then
         tType(TTCnt) = EmptyNetworkOrRemovableType(sfilename)
      End If
          
      'The file is not local, so I get the shortcut icon to avoid locking and timeouts
      himgSmall& = SHGetFileInfo(TTpath & shFilename, _
                 0&, shinfo, Len(shinfo), _
                 BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
   
       'get the target type of the file
       If SHGetFileInfo(sfilename, 0&, _
           shinfo, Len(shinfo), _
           SHGFI_TYPENAME Or SHGFI_USEFILEATTRIBUTES) Then
   
          tType(TTCnt) = TrimNull(shinfo.szTypeName)
                       
       Else
       
         If MaybeAFolder(sfilename) Then
            tType(TTCnt) = "File Folder"
         Else
            tType(TTCnt) = "Unknown"
         End If
         
       End If
                
      blNotValidated(TTCnt) = True
      If mNetwork.Checked = True Then
         fType(TTCnt) = "Network Drive"
      Else
         fType(TTCnt) = tType(TTCnt)
      End If
   
      'eliminate drive letters only from file list
      If Right$(sfilename, 2) = ":\" Then
         Exit Function
      End If
   
      rf = rf + 1    'counter for About box - just a raw count of shortcut "items"
      
      ParseNetworkFile = True
         
   Else
   
      ParseNetworkFile = False
      
   End If
   
errlog:
   subErrLog ("Form1:ParseNetworkFile")
End Function

Private Function ParseRemovableFile(RemovableDrive As String, shFilename As String, sfilename As String) As Boolean
On Error GoTo errlog

   sfilename = Trim$(TrimNull(sfilename))
   
   Call NullLetterFix(RemovableDrive, sfilename)
   
   If Left$(sfilename, 3) = RemovableDrive Then
   
      himgSmall& = SHGetFileInfo(sfilename, _
                          0&, shinfo, Len(shinfo), _
                          BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
                          
       tType(TTCnt) = Trim$(TrimNull(shinfo.szTypeName))
       
       If Len(fType(TTCnt)) = 0 Then
         tType(TTCnt) = EmptyNetworkOrRemovableType(sfilename)
       End If
             
      'The file is not local, so I get the shortcut icon to avoid locking and timeouts
      himgSmall& = SHGetFileInfo(TTpath & shFilename, _
                 0&, shinfo, Len(shinfo), _
                 BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
   
       'get the target type of the file
       If SHGetFileInfo(sfilename, 0&, _
               shinfo, Len(shinfo), _
               SHGFI_TYPENAME Or SHGFI_USEFILEATTRIBUTES) Then
   
          tType(TTCnt) = TrimNull(shinfo.szTypeName)
                       
       Else
      
          If MaybeAFolder(sfilename) Then
            tType(TTCnt) = "File Folder"
          Else
            tType(TTCnt) = "Unknown"
          End If
          
       End If
               
      'eliminate drive letters only from file list
      If Right$(sfilename, 2) = ":\" Then
         Exit Function
      End If
      
      blNotValidated(TTCnt) = True
      If mRemovable.Checked = True Then
         fType(TTCnt) = "Removable Drive"
      Else
         fType(TTCnt) = tType(TTCnt)
      End If
         
      rf = rf + 1    'counter for About box - just a raw count of shortcut "items"
      
      ParseRemovableFile = True
         
   Else
   
      ParseRemovableFile = False
      
   End If
   
errlog:
   subErrLog ("Form1:ParseRemovableFile")
End Function

Private Function ParseCDFile(CDDrive As String, shFilename As String, sfilename As String) As Boolean
On Error GoTo errlog

   sfilename = Trim$(TrimNull(sfilename))
   
   Call NullLetterFix(CDDrive, sfilename)
   
   If Left$(sfilename, 3) = CDDrive Then
   
      himgSmall& = SHGetFileInfo(sfilename, _
                          0&, shinfo, Len(shinfo), _
                          BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
                          
     tType(TTCnt) = Trim$(TrimNull(shinfo.szTypeName))
     
     If Len(fType(TTCnt)) = 0 Then
         tType(TTCnt) = EmptyNetworkOrRemovableType(sfilename)
     End If
       
      'The file is not local, so I get the shortcut icon to avoid locking and timeouts
      himgSmall& = SHGetFileInfo(TTpath & shFilename, _
                 0&, shinfo, Len(shinfo), _
                 BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
   
      'get the target type of the file
      If SHGetFileInfo(sfilename, 0&, _
               shinfo, Len(shinfo), _
               SHGFI_TYPENAME Or SHGFI_USEFILEATTRIBUTES) Then
   
          tType(TTCnt) = TrimNull(shinfo.szTypeName)
                       
      Else
      
          If MaybeAFolder(sfilename) Then
            tType(TTCnt) = "File Folder"
          Else
            tType(TTCnt) = "Unknown"
          End If
          
      End If
               
      
      'eliminate drive letters only from file list
      If Right$(sfilename, 2) = ":\" Then
         Exit Function
      End If
         
      blNotValidated(TTCnt) = True
      If mRemovable.Checked = True Then
         fType(TTCnt) = "CD Drive"
      Else
         fType(TTCnt) = tType(TTCnt)
      End If
      
      rf = rf + 1    'counter for About box - just a raw count of shortcut "items"
      
      ParseCDFile = True
         
   Else
   
      ParseCDFile = False
      
   End If
   
errlog:
   subErrLog ("Form1:ParseCDFile")
End Function

Private Function NullLetterFix(DriveLetter As String, sfilename As String) As String
On Error GoTo errlog
'truncated path fix - needed if drive letter is replaced with a null char
'   sFilename = "~" + Right$(sFilename, Len(sFilename) - 1)   'test
   If LetterTest(Left$(sfilename, 1)) = True Then
      NullLetterFix = sfilename
      Exit Function
   Else
      If InStr(sfilename, ":") > 0 Then
         Dim ColonPos As Integer
         Dim RestoreName As String
         RestoreName = sfilename
         ColonPos = InStr(sfilename, ":") - 1
         sfilename = Left$(DriveLetter, 1) + Right$(sfilename, Len(sfilename) - ColonPos)
         If DoesFileExist(sfilename) = False Then
            NullLetterFix = RestoreName
         Else
            NullLetterFix = sfilename
         End If
      End If
   End If
errlog:
   subErrLog ("Form1:NullLetterFix")
End Function

Private Function LetterTest(DL As String) As Boolean
On Error GoTo errlog
   Select Case UCase(DL)
      Case "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", _
           "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
         LetterTest = True
      Case Else
         LetterTest = False
   End Select
errlog:
   subErrLog ("Form1:LetterTest")
End Function

Private Function EmptySystemType(ByVal sfilename As String)
On Error GoTo errlog
'On W2K, SHGetFileInfo can fail for some file types, so if nothing is returned by szTypeName,
'see if the file has an extension and if the file exists, give it a file type according
'to the extension - works well for XML file, not so well for a Word file (ie, DOC file).
'If we don't do this, these "untyped types" would wrongly end up in the Deleted Folder.

   If InStr(sfilename, ".") > 0 Then
      If DoesFileExist(sfilename) = True Then
          Dim dot As Integer
          dot = InStrRev(sfilename, ".")
          EmptySystemType = UCase$(Right$(sfilename, Len(sfilename) - dot)) + " File"
      End If
   End If

errlog:
   subErrLog ("Form1:EmptySystemType")
End Function

Private Function EmptyNetworkOrRemovableType(ByVal sfilename As String)
On Error GoTo errlog
'On W2K, SHGetFileInfo can fail for some file types, so if nothing is returned by szTypeName,
'give it a file type according to the extension without even checking existence - better to
'show a deleted file as live than a live file as deleted...

   If InStr(sfilename, ".") > 0 Then
      Dim dot As Integer
      dot = InStrRev(sfilename, ".")
      EmptyNetworkOrRemovableType = UCase$(Right$(sfilename, Len(sfilename) - dot)) + " File"
   End If

errlog:
   subErrLog ("Form1:EmptyNetworkOrRemovableType")
End Function

Private Sub ContinueWithDetails(shFilename As String, sfilename As String)
On Error GoTo errlog

   Dim img1 As ListImage
   Dim img2 As ListImage
   Dim ThisType As String
   Dim dp As Integer
   Dim ft As Integer
   Dim DateChangeCnt As Long
      
   'Deleted or renamed
   If Len(Trim$(fType(TTCnt))) = 0 Then
   
      If blSnapshot = True Then
         Call subStatusLog("Deleted: " + shFilename + ", " + sfilename)
      End If
   
      Dim blDead As Boolean
      blDead = True
               
      'the file does not exist, so here I give only the shortcut name to get an Icon
      himgSmall& = SHGetFileInfo(TTpath & shFilename, _
                 0&, shinfo, Len(shinfo), _
                 BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
              
      'get the target type of the nonexistent file - this works!
      If SHGetFileInfo(sfilename, 0&, _
              shinfo, Len(shinfo), _
              SHGFI_TYPENAME Or SHGFI_USEFILEATTRIBUTES) Then
              
         tType(TTCnt) = TrimNull(shinfo.szTypeName)

         fType(TTCnt) = "Deleted or Renamed"
                     
      Else
      
         Rollback
         Exit Sub
         
      End If
       
   'eliminate drive letters only from file list
   ElseIf fType(TTCnt) = "Local Disk" Or Right$(sfilename, 2) = ":\" Then
           
      Rollback
      Exit Sub
      
   End If
   
   'Eliminate duplicate files during loading - different shortcuts pointing to the same target.
   'This iterates through the files of the current type, comparing the current filename
   'against the array of filenames of this type. After loading, flag existing files with date change.
    For ft = 0 To UBound(fType())
      If Len(sfilename) > 0 Then
         If sfilename = aSFileName(ft) Then
            If blLoading = True Then
               If InStr(shFilename, "(") > 0 And InStr(shFilename, ")") > 0 Then
                  Kill TTpath + shFilename      'delete the (incremented) dupe shortcuts
               End If
               Kill sRecent + shFilename
               Rollback
               Exit Sub
            Else
               DateChangeCnt = ft      'coming from beginnewrecentshortcuts - the date has changed
   '            blDateChange = True
            End If
         End If
      End If
    Next ft
   
   'folder type fix - English only
   Dim blFolderFix As Boolean   'evaluate individually for speed
'   If Len(fType(TTCnt)) = 0 Then blFolderFix = True
'   If Len(tType(TTCnt)) = 0 Then blFolderFix = True
   If tType(TTCnt) = "File" Then blFolderFix = True
   If fType(TTCnt) = "File" Then blFolderFix = True
   
   If blFolderFix Then
      Dim blFolderExists As Boolean
      If DoesFolderExist(sfilename) = True Then
         blFolderExists = True
      Else
         blFolderExists = False
      End If
      If tType(TTCnt) = "File" Or tType(TTCnt) = "" Then
         If blFolderExists = True Then
            tType(TTCnt) = "File Folder"
         Else
            tType(TTCnt) = "Unknown"
         End If
      End If
      If fType(TTCnt) = "File" Or fType(TTCnt) = "" Then
         If blFolderExists = True Then
            fType(TTCnt) = "File Folder"
         Else
            fType(TTCnt) = "Unknown"
         End If
      End If
   End If
 
   If blDead = True Or tType(TTCnt) = "Unknown" Then
      blExists(TTCnt) = False
      blNotValidated(TTCnt) = True
   Else
      blExists(TTCnt) = True
   End If
         
   'Use ttype to mark unwanted types
   ReDim Preserve fNoShow(HD)
   For dp = 0 To UBound(fNoShow)
      If blHideType = True Then
         If fType(TTCnt) = fNoShow(dp) Then
            tType(TTCnt) = "hide-" + tType(TTCnt)
         End If
      End If
   Next dp
   
   ThisType = fType(TTCnt)
   
   'Only want one representative icon of each type
   'So eliminate duplicate types from filetype list, but continue processing
   ReDim Preserve PrevType(TTCnt)
   Dim pt As Integer
   For pt = 0 To TTCnt
      If PrevType(pt) = ThisType Then
      
         blNewType = False
                  
         GoTo DupeFileType
         
      End If
   Next pt
   
   blNewType = True
   
   PrevType(pnt) = ThisType
   pnt = pnt + 1
          
   If blLoading = True Or blBegin = True Then
        
      'cannot do the icon thing any later than this, or they get discombobulated
      Call AddFileItemIcons(, himgSmall&)
               
      'Imagelist1 doubles up for use by Listview1 (types) and Listview2 (files)is used by List1.
      'shFilename (shortcut filename) is, by definition, unique within the TT or Recent folder on
      'first load. But reloading poses a problem. If I clear the imagelist, listview2 appears without
      'icons for a noticeable period. However, when reloading the shFilenames are duplicates and generate
      'Err.Number = 35602 (key is not unique). However, since the icons are already loaded for existing files
      'and types, it doesn't impact the integrity of the lists. (There's no need to rollback, just continue.)
      
      On Error Resume Next
      Set img1 = ImageList1.ListImages.Add(, fType(TTCnt), pixSmall.Picture)
      On Error GoTo 0
      On Error GoTo errlog
      
      blFirstGoodFile = True
      
   Else
   
      If blFirstGoodFile = False Then
         Rollback
         Exit Sub
      End If
      GoTo DupeFileType

   End If
     
'   LockWindowUpdate ListView1.hwnd  'loses this setting
              
   'Add the unique file type to the list - exclude hidden
   If Left$(tType(TTCnt), 5) <> "hide-" Then
      Set itmT = ListView1.ListItems.Add(, fType(TTCnt), fType(TTCnt))
   End If
      
   itmT.SubItems(1) = "1"  'initialize with 0 for file count
   
   itmT.SmallIcon = ImageList1.ListImages(fType(TTCnt)).Key
   ListView1.ListItems(ListView1.ListItems.Count).Selected = False
   
'   LockWindowUpdate 0
   
'   If Left(tType, 5) = "Adobe" Then
'      Debug.Print aSFileName(TTCnt)
'   End If
   
DupeFileType:
   
'   Debug.Print Str(TTCnt)
   AddtoFileCount
   
   If DateChangeCnt = 0 Then
   
      aSFileName(TTCnt) = sfilename
     
      aShortcut(TTCnt) = shFilename

   Else
   
      aSFileName(DateChangeCnt) = sfilename  ' prevents creating multiple instances of same file
     
      aShortcut(DateChangeCnt) = shFilename
     
   End If
      
   If sfilename = DroppedName Then
      DroppedName = vbNullString
      SelectThisFile = shFilename
   End If
   
   On Error Resume Next
   Call AddFileItemIcons(, himgSmall&) 'unfortunately, I need to do this over or the icons go nuts
   Set img2 = ImageList2.ListImages.Add(, shFilename, pixSmall.Picture)
   On Error GoTo 0
   On Error GoTo errlog
      
   If blExists(TTCnt) = True Then
        If Form2.mAccessed.Checked = True Then
          Call DateTheFile(TTpath + shFilename)   'shortcut Accessed time works better than file access
        Else
          Call DateTheFile(aSFileName(TTCnt))
        End If
        If blLoading = False Then
          If modData.blNoAccess = True Then           'this defends against a newly added locked shortcut
             Call DateTheFile(aSFileName(TTCnt))
             modData.blNoAccess = False
          End If
        End If
   Else
        Call DateTheFile(TTpath + shFilename)     'use shortcut if doesn't exist
   End If
   
   ReDim Preserve blTrueDates(TTCnt)
   blTrueDates(TTCnt) = True
   
   If GetInputState() <> 0 Then
      DoEvents
   End If
      
errlog:
'   LockWindowUpdate 0
   Set img1 = Nothing
   Set itmT = Nothing
   subErrLog ("Form1:ContinueWithDetails")
End Sub

Private Sub Rollback()
On Error GoTo errlog
   If TTCnt = 0 Then Exit Sub
   TTCnt = TTCnt - 1              'roll back the counter and do not ContinueWithDetails
   DupeFileCount = DupeFileCount + 1    'value is stored in registry and only checked on next cached startup
   ReDim Preserve fType(TTCnt)
   ReDim Preserve blExists(TTCnt)
   ReDim Preserve aSFileName(TTCnt)
   ReDim Preserve tType(TTCnt)
errlog:
subErrLog ("Form1:Rollback")
End Sub

Private Sub AddtoFileCount()
On Error GoTo errlog

'Debug.Print fType(TTCnt) + "  " + Str(TTCnt)

   If ListView1.ListItems.Count = 0 Then Exit Sub
   
   If blNewType = True Then
'      ListView1.ListItems(fType(TTCnt)).SubItems(1) = "1"
      Exit Sub
   End If
   
   Dim fc As Integer
   fc = Val(Trim$(ListView1.ListItems(fType(TTCnt)).SubItems(1)))
   ListView1.ListItems(fType(TTCnt)).SubItems(1) = Trim$(Str$(fc + 1))

   filecnt = filecnt + 1
   
errlog:
   If Err.Number <> 35601 Then   'Element not found
      subErrLog ("Form1:AddtoFileCount")
   End If
End Sub

Public Sub GetTheFileList()
On Error Resume Next

   Dim hfile As Long
   Dim WFD As WIN32_FIND_DATA
 
   lbTemp.Visible = True
   frTemp.Visible = True
   Me.MousePointer = vbArrowHourglass
   ListView1.ToolTipText = "Please wait..."
  
'   DoEvents
   
   If Len(TTpath) = 0 Then
      GetTheFPath    'failsafe
   End If
    
   If Len(TTpath) > 0 Then
      
       fname = TTpath & "*.lnk"

       InitListview1
       
       hfile = FindFirstFile(fname, WFD)
       
       If mIcons.Checked = False Then
          Label1.Visible = False
          ProgressBar1.Max = fCount
          ProgressBar1.Value = 0
          ProgressBar1.Visible = True
          StatusBar1.Panels(1) = "Loading file types..."
       End If
       
'       If blReloading Then
'         LockWindowUpdate ListView1.hwnd  'lock updates to prevent clearing and reload flicker
'       End If
       
       ListView1.ListItems.Clear
       
       If hfile > 0 Then
          ReDim Preserve fType(fCount)
          ReDim Preserve blExists(fCount)
          ReDim Preserve blNotValidated(fCount)
          ReDim Preserve aSFileName(fCount)
          ReDim Preserve tType(fCount)
          ReDim Preserve aShortcut(fCount)

          BeginFileItemDetails WFD
          
          lbTemp.Visible = False
          frTemp.Visible = False
          
          If GetInputState() <> 0 Then
             DoEvents
          End If
          
          Do While FindNextFile(hfile, WFD)
          
            If Form1.Visible = True Then
               ProgressBar1.Value = lnkcnt
               ProgressBar1.ToolTipText = Str$(lnkcnt) + " shortcuts processed"
               StatusBar1.Panels(1) = "Loading file types... " + Str$(Trim$(Int(lnkcnt / fCount * 100))) + "%"
               StatusBar1.Panels(1).ToolTipText = Str$(lnkcnt) + " shortcuts processed"
            Else
               If lnkcnt Mod 10 = 0 Then
                  If mTaskbar.Checked = False And blNoSysTray = False Then
                     subSysTray
                  End If
               End If
            End If
            
            lnkcnt = lnkcnt + 1
            TTCnt = TTCnt + 1
                                    
            
            BeginFileItemDetails WFD
            
            If ListView1.ListItems.Count > 0 Then
               ListView1.ListItems(1).Selected = True    'first select the first item
               ListView1.ListItems(1).Selected = False   'then remove selection highlight
            End If
            
'            subUnlock (lnkcnt)
            
            If mIcons.Checked = False Then VertScrollbar
            
            If lnkcnt Mod 25 = 0 Then
               ListView1.Refresh
               Me.Refresh
'               If GetInputState() <> 0 Then
'                  DoEvents
'               End If
            End If
            
          Loop
          
      Else     'empty TTDir, so start over
      
         blNoCache = True
         blNewTTDir = True
         subStart
         
      End If
      
      FindClose hfile
          
   End If
   
   ListView1.Sorted = True
   
   subTypeSort

   RestoreNormalState
   
   If Form2.Visible = True Then
      If Len(KeepType) > 0 Then
         ListView1.ListItems(KeepType).Selected = True
      End If
      ListView1_Click
   End If
         
errlog:
'   If blReloading Then
'      LockWindowUpdate 0
'   End If
   subErrLog ("Form1:GetTheFileList")
End Sub

Public Sub RestoreNormalState()
On Error GoTo errlog

'   ListView1.ListItems(1).Selected = True    'first select the first item
'   ListView1.ListItems(1).Selected = False   'then remove selection highlight

   If LenB(GetSetting("TaskTracker", "Settings", "Recovery")) > 0 Then
      DeleteSetting "TaskTracker", "Settings", "Recovery"
   End If
   
   ProgressBar1.Visible = False
   StatusBar1.Panels(1).Text = Str$(ListView1.ListItems.Count) + " file types loaded."
   
   If mIcons.Checked = False Then
      Label1.Visible = True
      Label1.ForeColor = &H80000012
   End If
   
   ListView1.ToolTipText = vbNullString
   If blLoading = True Then
      Label1.ToolTipText = "Tracking " + Trim$(Str$(filecnt)) + " files (" + Trim$(Str$(lnkcnt)) + " shortcuts)"
      StatusBar1.Panels(1).ToolTipText = "Tracking " + Trim$(Str$(filecnt)) + " files (" + Trim$(Str$(lnkcnt)) + " shortcuts)"
   End If
   
   'bandaid for minstart with no caching to keep systray startup minimized
   If blMinStart = True Then
      If mTaskbar.Checked = False Then
         Me.WindowState = vbMinimized
      End If
   End If

errlog:
   Me.MousePointer = vbDefault
   subErrLog ("Form1:RestoreNormalState")
End Sub

Public Sub subTypeSort()
On Error GoTo errlog
   'Set type sort between Alpha (default), Frequency, and Most Recent
   If TypeSortCol = 1 Then
      mNumber_Click
   ElseIf TypeSortCol = 2 Then
      mLatest_Click
   End If
   'else Alpha
errlog:
   subErrLog ("Form1:subTypeSort")
End Sub

Public Function FilesCountAll(sSource As String, sFileType As String) As Long
On Error GoTo errlog

   Dim WFD As WIN32_FIND_DATA
   Dim hfile As Long
   
  'Start searching for files in
  'sSource by obtaining a file
  'handle to the first file matching
  'the filespec passed
   hfile = FindFirstFile(sSource & sFileType, WFD)
   
   If hfile <> INVALID_HANDLE_VALUE Then
   
     'must have at least one, so ...
      Do
      
        'increment the counter and
        'find the next file
         FilesCountAll = FilesCountAll + 1
               
      Loop Until FindNextFile(hfile, WFD) = 0
      
   End If
   
  'Close the search handle
   Call FindClose(hfile)
   
errlog:
   subErrLog ("Form1:FilesCountAll")
End Function

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo errlog
  If blLoading = False Then
    If Button = 2 Then
      subRC1Menu
    End If
  End If
errlog:
   subErrLog ("Form1:Label1_MouseUp")
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
On Error GoTo errlog
   
   If mAlpha.Checked = True Then
      ListView1.ColumnHeaders(1).Text = "File Types - Frequency"
      mAlpha.Checked = False
      mNumber.Checked = True
      mLatest.Checked = False
      mNumber_Click
   ElseIf mNumber.Checked = True Then
      ListView1.ColumnHeaders(1).Text = "File Types - Most Recent"
      mAlpha.Checked = False
      mNumber.Checked = False
      mLatest.Checked = True
      mLatest_Click
   ElseIf mLatest.Checked = True Then
      ListView1.ColumnHeaders(1).Text = "File Types - Alphabetic"
      mAlpha.Checked = True
      mNumber.Checked = False
      mLatest.Checked = False
      mAlpha_Click
   End If
   
errlog:
   subErrLog ("Form1:ListView1_ColumnClick")
End Sub

'Private Sub ListView1_KeyPress(KeyAscii As Integer)
'   If blKeyDown = True Or blKeyUp = True Then
'      ListView1.ListItems(ListView1.SelectedItem.Index).Selected = True
'   End If
'End Sub

Private Sub ListView1_Keyup(KeyCode As Integer, Shift As Integer)
On Error GoTo errlog

'   If KeyCode <> vbKeyShift And KeyCode <> vbKeyControl Then
'      blSelectAll = False
'   End If
   If KeyCode = vbKeyF5 Then
      Exit Sub
   End If
   If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
      ListView1.ListItems(ListView1.SelectedItem.Index).Selected = True
      blKeyClick = True
      ListView1_Click
      blKeyClick = False
   End If
'   If KeyCode = vbKeyDown Then
'      blKeyDown = True
'   End If
'   If KeyCode = vbKeyUp Then
'      blKeyUp = True
'   End If
'   blKeyDown = False
'   blKeyUp = False
   If mGlueLeft.Checked = False Then
      If KeyCode = vbKeyLeft Then
         Form2.ListView2.SetFocus
      End If
   ElseIf mGlueLeft.Checked = True Then
      If KeyCode = vbKeyRight Then
         Form2.ListView2.SetFocus
      End If
   End If
   
errlog:
subErrLog ("Form1:ListView1_Keyup")
End Sub

Private Sub ListView1_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo errlog
   blform2 = False
   Call DragDropRoutine(Data, Effect, blform2)
   Form2.ListView2.MousePointer = ccDefault
errlog:
   subErrLog ("Form1:ListView1_OLEDragDrop")
End Sub

Private Sub mGridType_Click()
On Error GoTo errlog

   If mGridType.Checked = False Then
      mGridType.Checked = True
   Else
      mGridType.Checked = False
   End If
   Call SendMessage(ListView1.hWnd, _
                    LVM_SETEXTENDEDLISTVIEWSTYLE, _
                    LVS_EX_GRIDLINES, ByVal mGridType.Checked)
   
errlog:
   subErrLog ("Form1:mGridType_Click")
End Sub

Private Sub mSelectAll_Click()
On Error GoTo errlog

   If mSelectAll.Checked = False Then
      mSelectAll.Checked = True
'      Form4.ckAll.Value = 1
'      Form2.StatusBar1.Panels(1).Text = "Loading all files..."
      Form2.Caption = "TaskTracker - Loading..."
      Form2.Refresh
      blSelectAll = True
      ListView1_Click
   Else
      blSelectAll = False
      SelectOne
      mSelectAll.Checked = False
'      Form4.ckAll.Value = 0
      ListView1_Click
   End If
   
errlog:
   subErrLog ("Form1:mSelectAll_Click")
End Sub

Public Sub mTop_Click()
On Error GoTo errlog

   If mTop.Checked = False Then
      Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME)
      Call SetWindowPos(Form2.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME)
      Call SetWindowPos(About.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME)
      Call SetWindowPos(Form4.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME)
      Call SetWindowPos(Form5.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME)
      Call SetWindowPos(Form6.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME)
      mTop.Checked = True
   Else
      Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME)
      Call SetWindowPos(Form2.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME)
      Call SetWindowPos(About.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME)
      Call SetWindowPos(Form4.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME)
      Call SetWindowPos(Form5.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME)
      Call SetWindowPos(Form6.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME)
      mTop.Checked = False
   End If
   
errlog:
   subErrLog ("Form1:mTop_Click")
End Sub

Private Sub mHideType_Click()
On Error GoTo errlog

   If TestVFs = False Then
      mShowAll.Checked = False
      blHideType = True
      HD = HD + 1
      ReDim Preserve fNoShow(HD)
      fNoShow(HD) = ListView1.SelectedItem.Text
      ListView1.ListItems.Remove (ListView1.SelectedItem.Text)
      StatusBar1.Panels(1).Text = Str$(ListView1.ListItems.Count) + " file types loaded."
      If Form2.Visible = True Then
         ListView1_Click
      End If
   Else
      MsgBox "This file type cannot be hidden at the moment because files of this type are in a virtual folder.", vbExclamation
   End If
errlog:
   subErrLog ("Form1:mHideType_Click")
End Sub

Private Function TestVFs() As Boolean
On Error GoTo errlog
  Dim vtp As Long
  For vtp = 0 To UBound(fType)
    If InStr(tType(vtp), "Virtual Folder") > 0 Then
       If fType(vtp) = ListView1.SelectedItem.Text Then
          TestVFs = True
          Exit Function
       End If
    End If
  Next
errlog:
   subErrLog ("Form1:TestVFs")
End Function


Private Sub mIcons_Click()
On Error GoTo errlog
   
   If mIcons.Checked = False Then
      mIcons.Checked = True
      IconsOnly
   Else
      mIcons.Checked = False
      StatusBar1.Visible = True
      If blLoading = True Then
         ProgressBar1.Visible = True
         Label1.Visible = False
      End If
      Form1.Width = 2800
      Label1.Visible = True
      VertScrollbar
      Form_Resize
      ListView1.HideColumnHeaders = False
      Me.Caption = "TaskTracker"
      mGridType.Enabled = True
      ListView1.ForeColor = vbWindowText
      Call SendMessage(ListView1.hWnd, _
                    LVM_SETEXTENDEDLISTVIEWSTYLE, _
                    LVS_EX_GRIDLINES, ByVal mGridType.Checked)
                    
      Me.Top = wBottom - Me.Height
      Me.Left = wWidth - Me.Width
      If Form2.WindowState <> vbMinimized Then
         Form2.Left = Form1.Left - Form2.Width
         Form2.Top = Form1.Top
      End If
   End If
   Me.Refresh
   
errlog:
   subErrLog ("Form1:mIcons_Click")
End Sub

Private Sub IconsOnly()
On Error GoTo errlog
'   Set mclsStyle = New clsStyle
'   mclsStyle.Titlebar = False
   ListView1.HideColumnHeaders = True
   Me.Width = 600
   ListView1.Width = 600
   Me.Caption = vbNullString
   mGridType.Enabled = False
'   ListView1.GridLines = False
   ListView1.Height = Me.Height
   Label1.Visible = False
   StatusBar1.Visible = False
   ProgressBar1.Visible = False
   ListView1.ForeColor = ListView1.BackColor
   
   If blLoading = False Then
      ListView1.ColumnHeaders(1).Width = 325
   End If
   
   Me.Top = wBottom - Me.Height
   Me.Left = wWidth - Me.Width
   Form2.Left = Form1.Left - Form2.Width
   Form2.Top = Form1.Top

errlog:
   subErrLog ("Form1:IconsOnly")
End Sub

Private Sub mShowAll_Click()
On Error Resume Next
Dim hi As Integer
Dim hw As Integer
Dim lgx As ListItem
'Dim blAdd As Boolean

   blHideType = False
   
   For hi = 0 To UBound(fNoShow)
      For hw = 0 To UBound(fType)
         If hw = UBound(fType) Then Exit For
         If Len(tType(hw)) > 0 Then
            If fType(hw) = fNoShow(hi) Then
               If Left$(tType(TTCnt), 14) <> "Virtual Folder" Then
                  tType(hw) = Right$(tType(hw), Len(tType(hw)) - 5)  'remove "hide-" - only added with no-cache start
'                  blAdd = True
               End If
            End If
         End If
      Next hw
      If Len(fType(hw)) > 0 And Len(fNoShow(hi)) > 0 Then
         Set lgx = ListView1.ListItems.Add(, fNoShow(hi), fNoShow(hi))
         lgx.SmallIcon = ImageList1.ListImages(fNoShow(hi)).Key
      End If
   Next hi

   Erase fNoShow
   HD = 0
'   ListView1_Click

   StatusBar1.Panels(1).Text = Str$(ListView1.ListItems.Count) + " file types loaded."

errlog:
   mShowAll.Checked = True
   mShowAll.Enabled = False
   subErrLog ("Form1:mShowAll_click")
End Sub

Private Sub mLatest_Click()

   mLatest.Checked = True
   mNumber.Checked = False
   mAlpha.Checked = False
   ListView1.Sorted = True
   ListView1.SortKey = 2
   ListView1.SortOrder = lvwDescending
   lv1 = True
   'api callback routine required for dates only
   SendMessage ListView1.hWnd, _
                LVM_SORTITEMS, _
                ListView1.hWnd, _
                ByVal FARPROC(AddressOf CompareDates)
   lv1 = False
   ListView1.SortKey = 3  'this trick resets the listview indexes after the callback sort
   TypeSortCol = 2
   If ListView1.ListItems.Count > 0 Then
      ListView1.ColumnHeaders(1).Text = "File Types - Most Recent"
      ListView1.ListItems(ListView1.SelectedItem.Index).EnsureVisible
   End If
      
errlog:
   subErrLog ("Form1:mLatest_Click")
End Sub

Private Sub mNumber_Click()
On Error GoTo errlog

   mNumber.Checked = True
   mLatest.Checked = False
   mAlpha.Checked = False
   ListView1.Sorted = True
   ListView1.SortKey = 1
   ListView1.SortOrder = lvwAscending

   'api callback routine for value sort
   SendMessage ListView1.hWnd, _
                LVM_SORTITEMS, _
                ListView1.hWnd, _
                ByVal FARPROC(AddressOf CompareValues)

   ListView1.SortKey = 3  'this trick resets the listview indexes after the callback sort
   TypeSortCol = 1
   If ListView1.ListItems.Count > 0 Then
      ListView1.ColumnHeaders(1).Text = "File Types - Frequency"
      ListView1.ListItems(ListView1.SelectedItem.Index).EnsureVisible
   End If
      
errlog:
subErrLog ("Form1:mNumber_Click")
End Sub

Private Sub mTrueAccessDates_click()
On Error GoTo errlog
   If mTrueAccessDates.Checked = True Then
      mTrueAccessDates.Checked = False
   Else
      mTrueAccessDates.Checked = True
   End If
   If Form2.mAccessed.Checked = True Then
      Form2.mAccessed_Click
   End If
errlog:
   subErrLog ("Form1:mTrueAccessDates_click")
End Sub

Private Sub mAlpha_Click()
On Error GoTo errlog

   mNumber.Checked = False
   mLatest.Checked = False
   mAlpha.Checked = True
   ListView1.SortOrder = lvwAscending
   ListView1.SortKey = 0
   TypeSortCol = 0
   If ListView1.ListItems.Count > 0 Then
      ListView1.ColumnHeaders(1).Text = "File Types - Alphabetic"
      ListView1.ListItems(ListView1.SelectedItem.Index).EnsureVisible
   End If
      
errlog:
subErrLog ("Form1:mAlpha_Click")
End Sub

Private Sub mOpenType_Click(Index As Integer)
   ListView1_Click
End Sub

Public Sub mPopAbout_Click()
On Error GoTo errlog

   blExpired = False
'   Call Form1.Form_MouseMove(0, 0, 7680, 0)
   If blLoading = False Then
      About.txSystem = Trim$(Str$(vf)) + " Files"
      About.txNetwork = Trim$(Str$(rf)) + " Files"
   End If
'   If blExpired = True Then
'      About.Show 1
'   Else
'      About.Show 0
'   End If
   About.Show
   About.ZOrder 0
   
errlog:
subErrLog ("Form1:mPopAbout_click")
End Sub

Private Sub mPopAbout2_click()
   mPopAbout_Click
End Sub

Public Sub mRefresh_Click()
On Error GoTo errlog

   If blLoading = True Then
      Exit Sub
   End If
   
   blWait = False
   
   subGetFixedDrives
   
   If blDriveChange = True Then
      Dim Response As Integer
      Response = MsgBox("A network or removable drive has changed." + vbNewLine _
      + "Do you want to reload TaskTracker? (This can take a long time.)", 292)
      If Response = vbNo Then
         blDriveChange = False
      End If
   End If
   
   If blNeedUpdate Or blDriveChange = True Then
   
      subReload
      
   Else
      blUserRefresh = True
      blRefresh = True
      blWait = False
   End If

errlog:
   subErrLog ("Form1:mRefresh_click")
End Sub

Public Sub subReload()
On Error Resume Next

   SaveTTInfo     'be cautious

   Timer1.Enabled = False
   blNeedUpdate = False
   blDriveChange = False
   blVirtualView = False
   blLoading = True
   blReloading = True
   blAbort = True
   blAddType = False
   blThisType = True
   blAllTypes = False
   blMultSelect = False
   
   SaveSetting "TaskTracker", "Settings", "Recovery", "1"
   
   Label1.Visible = False
   Me.Refresh
   If mIcons.Checked = False Then
      ProgressBar1.Value = 0
      ProgressBar1.Visible = True
   End If
      
   Erase modData.KeepSelected
   
   ListView1.SmallIcons = Nothing
   Form2.ListView2.SmallIcons = Nothing
   Form5.ListView3.SmallIcons = Nothing
   ImageList1.ListImages.Clear
   ImageList2.ListImages.Clear
   Unload Form5
   
   subLoadIcons
   
'   Shell_NotifyIcon NIM_DELETE, nid
   If mTaskbar.Checked = False And blNoSysTray = False Then
      subSysTray
   End If
   
   LoadSequence
   If mSingleClick = True Then
      ListView1_Click
   End If
   
   'Do this to reload virtual folders
   blNoCache = True        'temporarily set this
   SetLoadTTFileData
   
errlog:
   blLoading = False
   blReloading = False
   blNoCache = Form1.mNoCache.Checked
   subErrLog ("Form1:subReload")
End Sub
  
Private Sub Listview1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo errlog

  If Button = 1 Then blSelectAll = False

  If blLoading = False Then
      If Button = 2 Then
         subRC1Menu
      Else
         blRightClick = False
      End If
  End If
  
'   If mIcons.Checked = True Then
'      ListView1.ForeColor = vbWindowText
'   End If
errlog:
   subErrLog ("Form1:Listview1_MouseUp")
End Sub

Private Sub subRC1Menu()
On Error Resume Next

      If ListView1.ListItems.Count > 1 Then
         If Trim$(ListView1.SelectedItem.Text) <> "Removable Drive" And _
            Trim$(ListView1.SelectedItem.Text) <> "Network Drive" And _
            Trim$(ListView1.SelectedItem.Text) <> "CD Drive" Then
            mHideType.Visible = True
            mHideType.Enabled = True
            mHideType.Caption = "Hide " + ListView1.SelectedItem.Text + " Type"
         Else
            mHideType.Visible = False
         End If
      Else
         mHideType.Visible = False
      End If
      
'      If Trim$(ListView1.SelectedItem.Text) = "Virtual Folders" Then
'         mOpenVirtual.Visible = True
'         mSelectVirtual.Visible = True
'         mHideVirtual.Visible = True
'         virsep.Visible = True
'      Else
'         mOpenVirtual.Visible = False
'         mSelectVirtual.Visible = False
'         mHideVirtual.Visible = False
'         virsep.Visible = False
'      End If

      If mVirtual.Checked = True Then
         mReload.Enabled = False
      Else
         mReload.Enabled = True
      End If
      
      If UBound(fNoShow) = 0 Then
         mShowAll.Enabled = False
      Else
         mShowAll.Enabled = True
      End If
      
      If blRegStatus = False Then
         mRegister2.Visible = True
      Else
         mRegister2.Visible = False
      End If
      
      PopupMenu mnOptions, , , , mOpenType(1)
            
      blRightClick = True
      
errlog:
   subErrLog ("Form1:subRC1Menu")
End Sub

Public Sub ListView1_Click()
On Error GoTo errlog

   blDragOut = False
   blVirtualView = False
      
'   If Not Me.ActiveControl Is ListView1 Then Exit Sub
   
   If blLoading = True Then
      ListView1.SelectedItem.Selected = False
      Exit Sub
   End If
      
'   If Me.ActiveControl Is ListView1 = False Then
'      Exit Sub
'   End If

   If blRightClick = True Then
       blRightClick = False
       Exit Sub
   End If
   
'   LVS = 0  'clear last virtual folder selection
     
'   ListView1.ColumnHeaders(1).Text = "File Types"  'restore default caption
      
   If blCtrlKey = False Then
      SelectionStatus
   Else
      blCtrlKey = False
   End If
   
   If blLoading = False Then
      blDontRepeatClick = True
   End If
   
   If blSelectAll = False Then
      If blMultSelect = False Then
         mSelectAll.Checked = False
'         Form4.ckAll.Value = 0
      End If
   End If
   
   blClicked = False
   
'  The 3 file type selection modes are set here
'  ********************************************
   If blMultSelect = True And blSelectAll = False Then 'add
      blAddType = True
      blThisType = False
      blAllTypes = False
   ElseIf blSelectAll = True Then   'all
      blAddType = False
      blThisType = False
      blAllTypes = True
      Form4.blWasAddType = False
   Else                             'this
      blAddType = False
      blThisType = True
      blAllTypes = False
      Form4.blWasAddType = False
   End If
'   If Form2.mFilter.Checked = True Then
'      Form4.subCBState
'   End If
'   *******************************************
         
   'the user-selected type
   sType = ListView1.SelectedItem.Text
   
   blSetTempCaption = True   'wait for timer to set loading caption - it may not be needed
   
   PopsType
   
   blForm2Show = True
   
   Me.Refresh
'   If ListView1.SelectedItem.Text = "Virtual Folders" Then
'      mVirtual.Checked = False 'make sure you open it, since it's a toggle
'      mVirtual_Click
'      Form5.mOpen_Click
'      Exit Sub
'   End If
   
   Form2.subForm2Init
             
   If mFocusFiles.Checked = True Then
      If blKeyClick = False Then
         Form2.SetFocus
      End If
   End If
   
   blDontRepeatClick = False
      
errlog:
   subErrLog ("Form1:ListView1_click")
End Sub

Public Sub PopsType()
On Error GoTo errlog

'   sType is usually the selected file type in listview1 but when a new file is dragged in...
   If Len(sRegType) > 0 Then
      Dim lst As Integer
      With Form1.ListView1
         For lst = 1 To .ListItems.Count
            .ListItems(lst).Selected = False
         Next lst
      End With
      'if nonsystem file (but still registered type)...
      ReDim Preserve fType(TTCnt)   'bugfix
      If mNetwork.Checked And fType(TTCnt) = "Network Drive" Or _
         mRemovable.Checked And fType(TTCnt) = "Removable Drive" Or _
         mRemovable.Checked And fType(TTCnt) = "CD Drive" Then
         sType = fType(TTCnt)
      Else  'all system files
         sType = sRegType
      End If
      Form2.SetTitle
   Else
   
      If Len(sType) = 0 Then     'just in case, give it something
         If Len(sLastType) > 0 Then
            sType = sLastType
         Else
            sType = ListView1.ListItems(1).Text
         End If
      End If
         
      If blThisType = True Then
         ListView1.ListItems(sType).Selected = True
      End If
      
   End If
   
errlog:
   subErrLog ("Form1:PopsType")
End Sub

Public Sub SelectAllForm1()
On Error GoTo errlog

   Dim lst As Integer
   With Form1.ListView1
      For lst = 1 To .ListItems.Count
         Set .SelectedItem = .ListItems(lst)
      Next lst
'      blSelectAll = True
   End With
   
errlog:
   subErrLog ("Form1:SelectAllForm1")
End Sub

Public Sub SelectOne()
On Error GoTo errlog
   
   Dim lst As Integer
'   blWasMultSel = False
   With ListView1
      For lst = 1 To .ListItems.Count
         .ListItems(lst).Selected = False
      Next lst
      If Len(sType) > 0 Then
         Set .SelectedItem = .ListItems(sType)
      Else
         Set .SelectedItem = .ListItems(1)
      End If
   End With
   
errlog:
subErrLog ("Form1:SelectOne")
End Sub

Public Sub subShortStart()
   blRefresh = True
   blWait = False
   Timer1_Timer
End Sub

Private Sub subMinFix()
On Error GoTo errlog
  'special case for running minimized on win startup
   Form2.subSetSize
   PopFromMem      'bandaid to fix missing icons on min startup
   If blVFShow = True Then    'reshow on startup
      Form1.mVirtual_Click
   End If
   blMinStart = False
errlog:
   subErrLog ("Form1:subMinFix")
End Sub

Public Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo errlog

   If blLoading Then Exit Sub

   'This procedure receives the callbacks from the System Tray icon.
   Dim result As Long
   Dim msg As Long
   'the value of X will vary depending upon the scalemode setting
   If Me.ScaleMode = vbPixels Then
      msg = x
   Else
      msg = x \ Screen.TwipsPerPixelX
   End If
   Select Case msg
   
      Case WM_LBUTTONUP, WM_LBUTTONDBLCLK   '515 restore form window
         blForm1Resize = True
         
         subShowOnly
      
      Case WM_RBUTTONUP        '517 display popup menu
      
'         subMinFix
      
         If blMinStart = True Then
            If mTaskbar.Checked = False And blNoSysTray = False Then
               mHideShow(1).Caption = "&Show TaskTracker"
            End If
'            If blVFShow = True Then    'reshow on startup
'               Form1.mVirtual_Click
'            End If
         End If
         
'         If Form2.blPoponRestore = True Then
'            Form2.blPoponRestore = False
'            Call Form2.ListView2_ColumnClick(Form2.ListView2.ColumnHeaders(1))  'fix sortkey bug
'         End If

         result = SetForegroundWindow(Me.hWnd)
         'sets bolding
         If blRegStatus = False Then
            mRegister.Visible = True
         Else
            mRegister.Visible = False
         End If
         If mHideShow(1).Caption = "&Show TaskTracker" Then
            PopupMenu mPopupSys, , , , mHideShow(1)
         Else
            PopupMenu mPopupSys
         End If
         
      Case WM_MOUSEMOVE
         
      Case Else
      
   End Select
     
   If Me.WindowState <> vbMinimized Then
      meleft = Me.Left
      metop = Me.Top
      meheight = Me.Height
'      mewidth = Me.Width
   End If

errlog:
'   If Err.Number <> 0 And Err.Number <> 5 Then
'      subErrLog ("Form1:Form_MouseMove")
'   End If
      subErrLog ("Form1:Form_MouseMove")
End Sub

'System Tray menu
Public Sub mHideShow_Click(Index As Integer)
On Error GoTo errlog
   'called when the user clicks the popup menu Restore command
   If blLoading = False Then
      subHideShow
   End If
errlog:
   subErrLog ("Form1:mHideShow_Click")
End Sub

Public Sub subHideShow()
On Error Resume Next
   
   If blTaskbarStatus = False Then       'systray
      If Me.Visible = False Then
      
         subShowOnly
         
       Else
         Me.Hide
         TrimMem
         If VB.Forms.Count > 1 Then
           If Form2.Visible = True Then
              Form2.Hide
              blForm2Show = True
           End If
           Form4.Hide      'resume next avoids skipping out due to object unloaded error
           If Me.mVirtual.Checked = True Then
               Form5.Hide
           End If
           Form6.Hide
         End If
         mHideShow(1).Caption = "&Show TaskTracker"
       End If
   End If
      
errlog:
   subErrLog ("Form1:subHideShow")
End Sub

Private Sub subShowOnly()
On Error GoTo errlog

   blReposition = True
   
   Dim result As Long
   Me.WindowState = vbNormal

   Me.Show 0
   subSingleClick

   result = SetForegroundWindow(Me.hWnd)
   If blMinStart = True Then
      subMinFix
   End If
   mHideShow(1).Caption = "&Hide TaskTracker"
   If mFade.Checked = True Then
      subRestore
   End If
   
   If Form1.mFocusTypes.Checked = True Then
      Form1.SetFocus
   Else
      Form2.SetFocus
   End If

errlog:
   subErrLog ("Form1:subShowOnly")
End Sub

Public Sub mPopExit_Click()
On Error Resume Next

   blExit = True
   'this removes the icon from the system tray
   Shell_NotifyIcon NIM_DELETE, nid
   Unload Me
   End
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo errlog

   If blBuilding Then
      blStop = True
      Exit Sub
   End If

   If blTaskbarStatus = False Then
   
      If blExit = False And UnloadMode <> vbAppWindows Then
         Cancel = 1
      Else
         Exit Sub
      End If
      
'      If blLoading = False Or blReloading = True Then
      
'      If blTaskbarStatus = False Then
          Me.Hide
'      Else
'         Me.WindowState = vbMinimized
'      End If
'      End If
      
      Form_Resize
          
   End If
   
errlog:
subErrLog ("Form1:Form_QueryUnload")
End Sub

Public Sub subCopyShortcuts(LnkSearch As String)
On Error Resume Next

   Dim shFilename As String, NewShName As String
   Dim sfilename As String
   Dim NumFiles As Integer
   Dim RecentCount As Long, TTCount As Long
   
   RecentCount = FilesCountAll(sRecent, "*.lnk")
   TTCount = FilesCountAll(TTpath, "*.lnk")
              
   If RecentCount > TTCount Then    'if for some reason this occurs, do a global copy
'       blNewTTDir = True
'       blAbort = True
'       blNoCache = False
   Else
      If blLoading = True Then
        If blNoCache = False And blInstantStarted = True Then
           Exit Sub
        End If
      End If
   End If

   NumFiles = 0      'returns # files copied
   
   blComparetoTTDate = True

   If blNewTTDir = True Then
      ProgressBar1.Max = RecentCount
      lbTemp.Visible = True
      frTemp.Visible = True
      Me.MousePointer = vbArrowHourglass
      ListView1.ToolTipText = "Please wait..."
      Form1.Visible = True
   End If

   'loop through Recent folder for all shortcuts
   shFilename = Dir$(LnkSearch)
   Do While shFilename <> ""
      
      NumFiles = NumFiles + 1
      
      If blSnapshot = True Then
          Call subStatusLog("Recent shortcut: " + shFilename)
      End If
      
      'only applies with newly created TaskTracker dir (or global copy)
      If blNewTTDir = True Then
      
         FileCopy (sRecent & shFilename), (TTpath & shFilename)
         
         If GetInputState() <> 0 Then
            DoEvents
         End If
         
         If mIcons.Checked = False Then
            ProgressBar1.Visible = True
            ProgressBar1.Value = NumFiles
            ProgressBar1.ToolTipText = Str$(NumFiles) + " shortcuts processed"
            StatusBar1.Panels(1).Text = "Building File List..." + Str$(Trim$(Int(NumFiles / RecentCount * 100))) + "%"
            StatusBar1.Panels(1).ToolTipText = Str$(NumFiles) + " shortcuts processed"
         End If

      'ShortcutStatus sees if shortcut exists in TT and has the same mod date as in Recent - if not...
      Else
      
         sfilename = fnResolveLink(sRecent + shFilename)
         If GetInputState() <> 0 Then
             DoEvents
         End If
         
         If ShortcutStatus(sfilename, shFilename, NewShName) = False Then
         
            If Len(NewShName) = 0 Then
               FileCopy (sRecent & shFilename), (TTpath & shFilename)
               If GetInputState() <> 0 Then
                   DoEvents
               End If
            Else
               FileCopy (sRecent & shFilename), (TTpath & NewShName)
               shFilename = NewShName
               If GetInputState() <> 0 Then
                   DoEvents
               End If
            End If
            
            NewShName = vbNullString
            dLastWrite = ComparetoTTDate(sRecent)     'update this to defend against "double-entry" errors
            
            If blLoading = False Then
               If Len(shFilename) > 0 Then
                   blNewShortcuts = True    'proceed with PopFromMem after...
                  '***********************
                   BeginNewRecentShortcuts (shFilename)   '>AddFileItemDetails>ContinueWithDetails
                  '***********************
               End If
            End If
            
         End If
      End If
            
      shFilename = Dir$        'get next matching file

      If GetInputState() <> 0 Then
          DoEvents
      End If
       
   Loop

errlog:
   Erase DroppedList
   blComparetoTTDate = False
   blNewTTDir = False
   subErrLog ("Form1:subCopyShortcuts")
End Sub

Private Function ShortcutStatus(sfilename As String, shFilename As String, Optional ByRef NewShName As String) As Boolean
On Error GoTo errlog

   If DoesFileExist(TTpath & shFilename) Then            'same-name shortcut in TT folder...
   
      If fnResolveLink(TTpath & shFilename) = fnResolveLink(sRecent & shFilename) Then    'and points to same file...
      
         If blNoCache = False And blInstantStarted = False Then
            ShortcutStatus = True      'so don't copy
            Exit Function
            
         Else
         
            If ComparetoTTDate(sRecent & shFilename) = ComparetoTTDate(TTpath & shFilename) Then   'if dates are the same...
               ShortcutStatus = True   'don't copy
               
            Else
            
                '*may* need to rename or it will wipe out the existing shortcut
                Call RenameTheShortcut(shFilename, NewShName, sfilename)
                ShortcutStatus = False      'copy - this is a reopened or updated file
                              
            End If
         End If
         
      Else  'determine if the shortcut points to a file that's already being tracked...
'         If CheckTTStatus(shFilename, True) = False Then     'if not...
         
               '*may* need to rename or it will wipe out the existing shortcut
               Call RenameTheShortcut(shFilename, NewShName, sfilename)
               ShortcutStatus = False                 'copy
               
'         End If
      End If
   Else
        '*may* need to rename or it will wipe out the existing shortcut
        Call RenameTheShortcut(shFilename, NewShName, sfilename)
        ShortcutStatus = False                 'copy
   End If
   
   If ShortcutStatus = False Then
       Call CheckTTStatus(shFilename, False)   'set blTrueDates to false if already being tracked
   End If
   
errlog:
subErrLog ("Form1:ShortcutStatus")
End Function

'This function does double-duty for ShortcutStatus, doing either:
'   "False" action: set blTrueDates to false if file is already being tracked  (nothing returned)
'   "True" action: determine if a duplicate shortcut points to an existing file (returns True)
Private Function CheckTTStatus(shFilename As String, ST As Boolean)
On Error GoTo errlog

   Dim ResolvedFile As String
   Dim cf As Long
   ResolvedFile = fnResolveLink(sRecent + shFilename)
'   Debug.Print Str(ST) + " - " + ResolvedFile
   For cf = 0 To UBound(aSFileName)
      If ResolvedFile = aSFileName(cf) Then
         If ResolvedFile <> fnResolveLink(TTpath + shFilename) Then
            If ST = True Then
               CheckTTStatus = True
               Exit For
            Else
               blTrueDates(cf) = False
               Exit For
            End If
         End If
      End If
   Next cf
   
errlog:
subErrLog ("Form1:CheckTTStatus")
End Function

Private Function RenameTheShortcut(shFilename As String, Optional ByRef NewShName As String, Optional sfilename As String) As Boolean
On Error GoTo errlog

   'if incremented in Recent, leave alone
   If InStr(shFilename, "(") > 0 And InStr(shFilename, ")") > 0 Then
      Exit Function
   End If

   If DoesFileExist(TTpath & shFilename) = True Then
      If fnResolveLink(TTpath & shFilename) = sfilename Then
          RenameTheShortcut = False
          Exit Function
       End If
       Dim ni As Integer
       ni = 1
       NewShName = Left$(shFilename, Len(shFilename) - 4) + " (1).lnk"
       RenameTheShortcut = True
   Else
       RenameTheShortcut = False
       Exit Function  'doesn't exist yet, so no need to rename
   End If
   
   Do While DoesFileExist(TTpath & NewShName) = True      'if shortcut name exists, increment (<n>) suffix
   
'     If name exists and does not resolve to the resolved Recent filename, increment the name
      If fnResolveLink(TTpath & NewShName) <> sfilename Then
          NewShName = Left$(shFilename, Len(shFilename) - 4) + " (" + (Trim$(Str$(ni))) + ").lnk"
          RenameTheShortcut = False
      Else
'          NewShName = vbNullString
          RenameTheShortcut = True
          Exit Do   'if the resolved name corresponds with a resolved Recent name, don't rename
      End If
      
      ni = ni + 1
   Loop

errlog:
   subErrLog ("Form1:RenameTheShortcut")
End Function

Public Sub BeginNewRecentShortcuts(ByVal NewRecentShortcut As String, Optional ByVal dfile As String)
On Error GoTo errlog

   Dim sfilename As String
   
   blBegin = True
   blDateSwitch = True
                     
   sType = vbNullString
   If Len(dfile) = 0 Then
      sfilename = fnResolveLink(TTpath + NewRecentShortcut)
'      sLastFileName = sfilename
   Else
      sfilename = dfile
   End If
   
   TTCnt = TTCnt + 1    'increment shortcut count by 1
         
   Call AddFileItemDetails(NewRecentShortcut, sfilename)
      
errlog:
subErrLog ("Form1:BeginNewRecentShortcuts")
End Sub

Private Sub GetSettings()
On Error GoTo errlog

   Dim mHideTypes As String
   Dim ln As Integer
      
   If LenB(GetSetting("TaskTracker", "Settings", "IconsOnly")) > 0 Then
      mIcons.Checked = CBool(GetSetting("TaskTracker", "Settings", "IconsOnly"))
   Else
      mIcons.Checked = False 'default
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "TypeSortCol")) = 0 Then
      TypeSortCol = 0   'default (alpha sort)
   Else
      TypeSortCol = Val(GetSetting("TaskTracker", "Settings", "TypeSortCol"))
   End If
         
   If LenB(GetSetting("TaskTracker", "Settings", "RegSortCol")) = 0 Then
      RegSortCol = 1  'default
   Else
      RegSortCol = Val(GetSetting("TaskTracker", "Settings", "RegSortCol"))
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "RegPSortCol")) = 0 Then
      RegPSortCol = 2  'default
   Else
      RegPSortCol = Val(GetSetting("TaskTracker", "Settings", "RegPSortCol"))
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "DelSortCol")) = 0 Then
      DelSortCol = 2  'default
   Else
      DelSortCol = Val(GetSetting("TaskTracker", "Settings", "DelSortCol"))
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "DelPSortCol")) = 0 Then
      DelPSortCol = 3  'default
   Else
      DelPSortCol = Val(GetSetting("TaskTracker", "Settings", "DelPSortCol"))
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "MultSortCol")) = 0 Then
      MultSortCol = 2 'default
   Else
      MultSortCol = Val(GetSetting("TaskTracker", "Settings", "MultSortCol"))
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "MultPSortCol")) = 0 Then
      MultPSortCol = 3  'default
   Else
      MultPSortCol = Val(GetSetting("TaskTracker", "Settings", "MultPSortCol"))
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "RegSortOrder")) = 0 Then
      RegSortOrder = 1  'default
   Else
      RegSortOrder = Val(GetSetting("TaskTracker", "Settings", "RegSortOrder"))
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "RegPSortOrder")) = 0 Then
      RegPSortOrder = 1  'default
   Else
      RegPSortOrder = Val(GetSetting("TaskTracker", "Settings", "RegPSortOrder"))
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "DelSortOrder")) = 0 Then
      DelSortOrder = 1  'default
   Else
      DelSortOrder = Val(GetSetting("TaskTracker", "Settings", "DelSortOrder"))
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "DelPSortOrder")) = 0 Then
      DelPSortOrder = 1  'default
   Else
      DelPSortOrder = Val(GetSetting("TaskTracker", "Settings", "DelPSortOrder"))
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "MultSortOrder")) = 0 Then
      MultSortOrder = 1  'default
   Else
      MultSortOrder = Val(GetSetting("TaskTracker", "Settings", "MultSortOrder"))
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "MultPSortOrder")) = 0 Then
      MultPSortOrder = 1  'default
   Else
      MultPSortOrder = Val(GetSetting("TaskTracker", "Settings", "MultPSortOrder"))
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "OnTop")) > 0 Then
      If CBool(GetSetting("TaskTracker", "Settings", "OnTop")) = True Then
         mTop_Click
      Else
         mTop.Checked = False   'default
      End If
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "Grid1")) > 0 Then
      mGridType.Checked = CBool(GetSetting("TaskTracker", "Settings", "Grid1"))
   Else
      mGridType.Checked = False   'default
   End If
   
   Call SendMessage(ListView1.hWnd, _
                    LVM_SETEXTENDEDLISTVIEWSTYLE, _
                    LVS_EX_GRIDLINES, ByVal mGridType.Checked)

   If LenB(GetSetting("TaskTracker", "Settings", "SingleClick")) > 0 Then
      mSingleClick.Checked = CBool(GetSetting("TaskTracker", "Settings", "SingleClick"))
      mDefaultFocus.Enabled = mSingleClick.Checked
      mAutohide.Enabled = mSingleClick.Checked
   Else
      mSingleClick.Checked = True   'default
      mDefaultFocus.Enabled = True
      mAutohide.Enabled = True
   End If

   If LenB(GetSetting("TaskTracker", "Settings", "GlueRight")) > 0 Then
      Form1.mGlueRight.Checked = CBool(GetSetting("TaskTracker", "Settings", "GlueRight"))
   Else
      Form1.mGlueRight.Checked = False   'default
   End If

   If LenB(GetSetting("TaskTracker", "Settings", "GlueLeft")) > 0 Then
      Form1.mGlueLeft.Checked = CBool(GetSetting("TaskTracker", "Settings", "GlueLeft"))
   Else
      Form1.mGlueLeft.Checked = False   'default
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "HiddenTypes")) > 0 Then
      mHideTypes = GetSetting("TaskTracker", "Settings", "HiddenTypes")
      'parse the string for each filetype
      For ln = 1 To Len(mHideTypes)
        If Mid$(mHideTypes, ln, 1) = ";" Then
           ReDim Preserve fNoShow(HD)
           fNoShow(HD) = Left$(mHideTypes, InStr(mHideTypes, ";") - 1)
           mHideTypes = Right$(mHideTypes, Len(mHideTypes) - ln)
           HD = HD + 1
           ln = 1
        End If
      Next ln
      If HD > 0 Then
         blHideType = True
         mShowAll.Checked = False
      End If
   Else
      ReDim Preserve fNoShow(0)
      blHideType = False
      mShowAll.Checked = True
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "Form1Width")) > 0 Then
      Form1.Width = Val(GetSetting("TaskTracker", "Settings", "Form1Width"))
   Else
      Form1.Width = 2800
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "Form1Height")) > 0 Then
      Form1.Height = Val(GetSetting("TaskTracker", "Settings", "Form1Height"))
   Else
      Form1.Height = 7200
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "Form1Top")) > 0 Then
      Form1.Top = Val(GetSetting("TaskTracker", "Settings", "Form1Top"))
      Form1.mSaveSize.Checked = True
   End If
   
   If Len(GetSetting("TaskTracker", "Settings", "Form1Left")) > 0 Then
      Form1.Left = Val(GetSetting("TaskTracker", "Settings", "Form1Left"))
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "Network")) > 0 Then
      mNetwork.Checked = CBool(GetSetting("TaskTracker", "Settings", "Network"))
   Else
      mNetwork.Checked = False 'default
   End If

   If LenB(GetSetting("TaskTracker", "Settings", "Removable")) > 0 Then
      mRemovable.Checked = CBool(GetSetting("TaskTracker", "Settings", "Removable"))
   Else
      mRemovable.Checked = False 'default
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "Taskbar")) > 0 Then
      mTaskbar.Checked = CBool(GetSetting("TaskTracker", "Settings", "Taskbar"))
   Else
      mTaskbar.Checked = False 'default
   End If
      
   If LenB(GetSetting("TaskTracker", "Settings", "Fade")) > 0 Then
      mFade.Checked = CBool(GetSetting("TaskTracker", "Settings", "Fade"))
      If mTaskbar.Checked = False Then    'can be checked but not active if taskbar is true
         blNoFade = False
         Timer2.Enabled = True
      Else
         mFade.Enabled = False
      End If
   Else
      mFade.Checked = False 'default
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "HideTips")) > 0 Then
      mHideTips.Checked = CBool(GetSetting("TaskTracker", "Settings", "HideTips"))
   Else
      mHideTips.Checked = False 'default
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "Refresh")) > 0 Then
      iRefreshTime = Int(Val(GetSetting("TaskTracker", "Settings", "Refresh")))
   Else
      iRefreshTime = 1     'default
   End If

   If LenB(GetSetting("TaskTracker", "Settings", "Autorun")) > 0 Then
      If GetSetting("TaskTracker", "Settings", "Autorun") = "Minimized" Then
         mMin.Checked = True
         SetAutoRun
      ElseIf GetSetting("TaskTracker", "Settings", "Autorun") = "Normal" Then
         mNorm.Checked = True
         SetAutoRun
      Else
         mMin.Checked = False
         mNorm.Checked = False
         ClearAutoRun
      End If
   Else
      mMin.Checked = False     'default
'      SetAutoRun
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "Snapshot")) > 0 Then
      blSnapshot = True
      DeleteSetting "TaskTracker", "Settings", "Snapshot"   'once-only setting
   Else
      blSnapshot = False 'default
   End If
     
   If LenB(GetSetting("TaskTracker", "Settings", "TypesFocus")) > 0 Then
      mFocusTypes.Checked = CBool(GetSetting("TaskTracker", "Settings", "TypesFocus"))
   Else
      mFocusTypes.Checked = True 'default
   End If
   
   If mFocusTypes.Checked = True Then
      mFocusFiles.Checked = False   'default
   Else
      mFocusFiles.Checked = True
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "NoCache")) > 0 Then
      If CBool(GetSetting("TaskTracker", "Settings", "NoCache")) = True Then
         mNoCache_Click
      End If
   Else
      mNoCache.Checked = False
   End If
'   SaveSetting "TaskTracker", "Settings", "NoCache", "False"
   
   If blSnapshot = True Then               'override caching if snapshot'ing
      blNoCache = True
      mSynchronization.Enabled = False
   End If

   If LenB(GetSetting("TaskTracker", "Settings", "Caching")) > 0 Then
      If GetSetting("TaskTracker", "Settings", "Caching") = "Never" Then
         mNeverSync.Checked = True
         mSyncShow.Checked = False
         mSyncStartup.Checked = False
      ElseIf GetSetting("TaskTracker", "Settings", "Caching") = "FirstShow" Then
         mNeverSync.Checked = False
         mSyncShow.Checked = True
         mSyncStartup.Checked = False
      ElseIf GetSetting("TaskTracker", "Settings", "Caching") = "Startup" Then
         mNeverSync.Checked = False
         mSyncShow.Checked = False
         mSyncStartup.Checked = True
      End If
   Else  'default
      mNeverSync.Checked = False
      mSyncShow.Checked = True
      mSyncStartup.Checked = False
   End If

'  Get splash setting before the rest
'   If LenB(GetSetting("TaskTracker", "Settings", "NoSplash")) > 0 Then
'      mNoSplash.Checked = CBool(GetSetting("TaskTracker", "Settings", "NoSplash"))
'   Else
'      mNoSplash.Checked = False 'default
'   End If

   If LenB(GetSetting("TaskTracker", "Settings", "AutoPreview")) > 0 Then
      mAutoPreview.Checked = CBool(GetSetting("TaskTracker", "Settings", "AutoPreview"))
   Else
      mAutoPreview.Checked = False 'default
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "AlwaysValidate")) > 0 Then
      mAlwaysValidate.Checked = CBool(GetSetting("TaskTracker", "Settings", "AlwaysValidate"))
   Else
      mAlwaysValidate.Checked = True 'default
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "autohide")) > 0 Then
      mAutohide.Checked = CBool(GetSetting("TaskTracker", "Settings", "autohide"))
   Else
      mAutohide.Checked = False 'default
   End If

   If LenB(GetSetting("TaskTracker", "Settings", "CheckUpdate")) > 0 Then
      mCheckUpdates.Checked = CBool(GetSetting("TaskTracker", "Settings", "CheckUpdate"))
   Else
      mCheckUpdates.Checked = True 'default
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "VFShow")) > 0 Then
      If CBool(GetSetting("TaskTracker", "Settings", "VFShow")) = True Then
         blVFShow = True
      End If
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "AccessDates")) > 0 Then
      mTrueAccessDates.Checked = CBool(GetSetting("TaskTracker", "Settings", "AccessDates"))
   Else
      mTrueAccessDates.Checked = False 'default
   End If

errlog:
   subErrLog ("Form1:GetSettings")
End Sub

Public Sub SaveSettings()
On Error GoTo errlog

   Dim sHideTypes As String
   Dim ht As Integer
   
   If blLoading = False Then
      Form2.Form2ColumnOrder  'get latest col order
'      SetListview2Columns     'get latest col sizes
   End If
   
   SaveSetting "TaskTracker", "Settings", "IconsOnly", Trim$(Str$(mIcons.Checked))
   SaveSetting "TaskTracker", "Settings", "TypeSortCol", TypeSortCol
   SaveSetting "TaskTracker", "Settings", "RegSortCol", Trim$(Str$(RegSortCol))
   SaveSetting "TaskTracker", "Settings", "RegPSortCol", Trim$(Str$(RegPSortCol))
   SaveSetting "TaskTracker", "Settings", "DelSortCol", Trim$(Str$(DelSortCol))
   SaveSetting "TaskTracker", "Settings", "DelPSortCol", Trim$(Str$(DelPSortCol))
   SaveSetting "TaskTracker", "Settings", "MultSortCol", Trim$(Str$(MultSortCol))
   SaveSetting "TaskTracker", "Settings", "MultPSortCol", Trim$(Str$(MultPSortCol))
   SaveSetting "TaskTracker", "Settings", "RegSortOrder", Trim$(Str$(RegSortOrder))
   SaveSetting "TaskTracker", "Settings", "RegPSortOrder", Trim$(Str$(RegPSortOrder))
   SaveSetting "TaskTracker", "Settings", "DelSortOrder", Trim$(Str$(DelSortOrder))
   SaveSetting "TaskTracker", "Settings", "DelPSortOrder", Trim$(Str$(DelPSortOrder))
   SaveSetting "TaskTracker", "Settings", "MultPSortOrder", Trim$(Str$(MultPSortOrder))
   SaveSetting "TaskTracker", "Settings", "MultPSortOrder", Trim$(Str$(MultPSortOrder))
   SaveSetting "TaskTracker", "Settings", "FullPath", Trim$(Str$(Form2.mFullPath.Checked))
   SaveSetting "TaskTracker", "Settings", "ContainerFolder", Trim$(Str$(Form2.mContainerFolder.Checked))
   SaveSetting "TaskTracker", "Settings", "FileExt", Trim$(Str$(Form2.mExt.Checked))
   SaveSetting "TaskTracker", "Settings", "Grid1", Trim$(Str$(mGridType.Checked))
   SaveSetting "TaskTracker", "Settings", "Grid2", Trim$(Str$(mGridName.Checked))
   SaveSetting "TaskTracker", "Settings", "OnTop", Trim$(Str$(mTop.Checked))
'   SaveSetting "TaskTracker", "Settings", "RefreshTime", Trim$(Str$(iRefreshTime))
   SaveSetting "TaskTracker", "Settings", "SingleClick", Trim$(Str$(Form1.mSingleClick.Checked))
   SaveSetting "TaskTracker", "Settings", "GlueRight", Trim$(Str$(mGlueRight.Checked))
   SaveSetting "TaskTracker", "Settings", "GlueLeft", Trim$(Str$(mGlueLeft.Checked))
   SaveSetting "TaskTracker", "Settings", "Explorer", Trim$(Str$(Form2.blExplorer))
   SaveSetting "TaskTracker", "Settings", "Network", Trim$(Str$(mNetwork.Checked))
   SaveSetting "TaskTracker", "Settings", "Removable", Trim$(Str$(mRemovable.Checked))
   SaveSetting "TaskTracker", "Settings", "Simple", Trim$(Str$(Form2.mSimple.Checked))
   SaveSetting "TaskTracker", "Settings", "Taskbar", Trim$(Str$(mTaskbar.Checked))
   SaveSetting "TaskTracker", "Settings", "Fade", Trim$(Str$(mFade.Checked))
   SaveSetting "TaskTracker", "Settings", "HideTips", Trim$(Str$(mHideTips.Checked))
   SaveSetting "TaskTracker", "Settings", "TypesFocus", Trim$(Str$(mFocusTypes.Checked))
   SaveSetting "TaskTracker", "Settings", "NoCache", Trim$(Str$(mNoCache.Checked))
   SaveSetting "TaskTracker", "Settings", "NoSplash", Trim$(Str$(mNoSplash.Checked))
   SaveSetting "TaskTracker", "Settings", "AutoPreview", Trim$(Str$(mAutoPreview.Checked))
   SaveSetting "TaskTracker", "Settings", "AlwaysValidate", Trim$(Str$(mAlwaysValidate.Checked))
   SaveSetting "TaskTracker", "Settings", "autohide", Trim$(Str$(mAutohide.Checked))
   SaveSetting "TaskTracker", "Settings", "CheckUpdate", Trim$(Str$(mCheckUpdates.Checked))
   SaveSetting "TaskTracker", "Settings", "VFShow", Trim$(Str$(Form1.mVirtual.Checked))
   SaveSetting "TaskTracker", "Settings", "AccessDates", Trim$(Str$(Form1.mTrueAccessDates.Checked))

   If Form2.mModified.Checked = True Then
      SaveSetting "TaskTracker", "Settings", "LastDate", "Modified"
   ElseIf Form2.mCreated.Checked = True Then
      SaveSetting "TaskTracker", "Settings", "LastDate", "Created"
   ElseIf Form2.mAccessed.Checked = True Then
      SaveSetting "TaskTracker", "Settings", "LastDate", "Accessed"
   End If
   
   If blHideType = True Then
      For ht = 0 To UBound(fNoShow)
         If Len(fNoShow(ht)) > 0 Then
            sHideTypes = sHideTypes + fNoShow(ht) + ";"
         End If
      Next ht
      SaveSetting "TaskTracker", "Settings", "HiddenTypes", sHideTypes
   Else
      SaveSetting "TaskTracker", "Settings", "HiddenTypes", ""
   End If
   
   If mSaveSize.Checked = True And blExit = True Then
     If Form2.WindowState = vbNormal And Form2.Visible = True Then
         subSaveSize
     End If
   End If

   If mMin.Checked = True Then
      SaveSetting "TaskTracker", "Settings", "AutoRun", "Minimized"
   ElseIf mNorm.Checked = True Then
      SaveSetting "TaskTracker", "Settings", "AutoRun", "Normal"
   Else
      SaveSetting "TaskTracker", "Settings", "AutoRun", "False"
   End If
   
   If mNeverSync.Checked = True Then
      SaveSetting "TaskTracker", "Settings", "Caching", "Never"
   ElseIf mSyncShow.Checked = True Then
      SaveSetting "TaskTracker", "Settings", "Caching", "FirstShow"
   Else
      SaveSetting "TaskTracker", "Settings", "Caching", "Startup"
   End If
         
errlog:
   subErrLog ("Form1:SaveSettings")
End Sub

Private Sub subSaveSize()
On Error GoTo errlog

'defend against column collapse by not saving zero widths

   If col1width = 0 And col2width = 0 Then
   'skip
   Else
      SaveSetting "TaskTracker", "Settings", "col1width", Trim$(Str$(col1width))
      SaveSetting "TaskTracker", "Settings", "col2width", Trim$(Str$(col2width))
'      SaveSetting "TaskTracker", "Settings", "colSwidth", Trim$(Str$(colSwidth))
   End If
   
   If col1Awidth = 0 And col2Awidth = 0 And col3Awidth = 0 Then
   'skip
   Else
      SaveSetting "TaskTracker", "Settings", "col1Awidth", Trim$(Str$(col1Awidth))
      SaveSetting "TaskTracker", "Settings", "col1NAwidth", Trim$(Str$(col1NAwidth))
      SaveSetting "TaskTracker", "Settings", "col2Awidth", Trim$(Str$(col2Awidth))
'      SaveSetting "TaskTracker", "Settings", "col2NAwidth", Trim$(Str$(col2NAwidth))
      SaveSetting "TaskTracker", "Settings", "col3Awidth", Trim$(Str$(col3Awidth))
      SaveSetting "TaskTracker", "Settings", "col3NAwidth", Trim$(Str$(col3NAwidth))
   End If
   
   If col1Pwidth = 0 And col2Pwidth = 0 And col3Pwidth = 0 Then
   'skip
   Else
      SaveSetting "TaskTracker", "Settings", "col1Pwidth", Trim$(Str$(col1Pwidth))
      SaveSetting "TaskTracker", "Settings", "col2Pwidth", Trim$(Str$(col2Pwidth))
      SaveSetting "TaskTracker", "Settings", "col3Pwidth", Trim$(Str$(col3Pwidth))
      SaveSetting "TaskTracker", "Settings", "col4Swidth", Trim$(Str$(col4Swidth))
   End If
   
   If col1PAwidth = 0 And col2PAwidth = 0 And col3PAwidth = 0 Then
   'skip
   Else
      SaveSetting "TaskTracker", "Settings", "col1PAwidth", Trim$(Str$(col1PAwidth))
      SaveSetting "TaskTracker", "Settings", "col1NPAwidth", Trim$(Str$(col1NPAwidth))
      SaveSetting "TaskTracker", "Settings", "col2PAwidth", Trim$(Str$(col2PAwidth))
      SaveSetting "TaskTracker", "Settings", "col2NPAwidth", Trim$(Str$(col2NPAwidth))
      SaveSetting "TaskTracker", "Settings", "col3PAwidth", Trim$(Str$(col3PAwidth))
'      SaveSetting "TaskTracker", "Settings", "col3NPAwidth", Trim$(Str$(col3NPAwidth))
      SaveSetting "TaskTracker", "Settings", "col4PAwidth", Trim$(Str$(col4PAwidth))
      SaveSetting "TaskTracker", "Settings", "col4NPAwidth", Trim$(Str$(col4NPAwidth))
   End If
    
   'This ensures both forms are inside the desktop
   If mGlueRight.Checked = True Then
      If Form2.Left < 0 Then Form2.Left = 0
      If Form1.Left < Form2.Width - Form1.Width Then
         Form1.Left = Form2.Width
      End If
      If Form1.Left > wWidth - Form1.Width Then
         Form1.Left = wWidth - Form1.Width
         Form2.Left = Form1.Left - Form2.Width
      End If
   ElseIf mGlueLeft.Checked = True Then
      If Form1.Left < 0 Then Form1.Left = 0
      If Form2.Left < Form1.Width - Form2.Width Then
         Form2.Left = Form1.Width
      End If
      If Form2.Left > wWidth - Form2.Width Then
         Form2.Left = wWidth - Form2.Width
         Form1.Left = Form2.Left - Form1.Width
      End If
   End If
   If Form1.Top < 0 Then Form1.Top = 0
   If Form1.Top + Form1.Height > wBottom Then
      If Form1.Height < wBottom Then
         Form1.Top = wBottom - Form1.Height
      Else
         Form1.Top = 0
         Form1.Height = wBottom
      End If
   End If
   If Form2.Top < 0 Then Form2.Top = 0
   If Form2.Top + Form2.Height > wBottom Then
      If Form2.Height < wBottom Then
         Form2.Top = wBottom - Form2.Height
      Else
         Form2.Top = 0
         Form2.Height = wBottom
      End If
   End If
   
   SaveFormPositions
   
errlog:
   subErrLog ("Form1:fnSaveSize")
End Sub

Private Sub SaveFormPositions()
On Error GoTo errlog

   If Form1.WindowState = vbNormal Then
      SaveSetting "TaskTracker", "Settings", "Form1Left", Trim$(Str$(Form1.Left))
      SaveSetting "TaskTracker", "Settings", "Form1Top", Trim$(Str$(Form1.Top))
      SaveSetting "TaskTracker", "Settings", "Form1Height", Trim$(Str$(Form1.Height))
      SaveSetting "TaskTracker", "Settings", "Form1Width", Trim$(Str$(Form1.Width))
   End If
   
   If Form2.WindowState = vbNormal Then
      SaveSetting "TaskTracker", "Settings", "Form2Left", Trim$(Str$(Form2.Left))
      SaveSetting "TaskTracker", "Settings", "Form2Top", Trim$(Str$(Form2.Top))
      SaveSetting "TaskTracker", "Settings", "Form2Height", Trim$(Str$(Form2.Height))
      SaveSetting "TaskTracker", "Settings", "Form2Width", Trim$(Str$(Form2.Width))
   End If

   If Form4.WindowState = vbNormal Then
      SaveSetting "TaskTracker", "Settings", "Form4Left", Trim$(Str$(Form4.Left))
      SaveSetting "TaskTracker", "Settings", "Form4Top", Trim$(Str$(Form4.Top))
   End If
   
'   If Form5.WindowState = vbNormal Then
'      SaveSetting "TaskTracker", "Settings", "Form5Left", Trim$(Str$(Form5.Left))
'      SaveSetting "TaskTracker", "Settings", "Form5Top", Trim$(Str$(Form5.Top))
'      SaveSetting "TaskTracker", "Settings", "Form5Height", Trim$(Str$(Form5.Height))
'      SaveSetting "TaskTracker", "Settings", "Form5Width", Trim$(Str$(Form5.Width))
'   End If
   
   If Form6.WindowState = vbNormal Then
      SaveSetting "TaskTracker", "Settings", "Form6Left", Trim$(Str$(Form6.Left))
      SaveSetting "TaskTracker", "Settings", "Form6Top", Trim$(Str$(Form6.Top))
      SaveSetting "TaskTracker", "Settings", "Form6Height", Trim$(Str$(Form6.Height))
      SaveSetting "TaskTracker", "Settings", "Form6Width", Trim$(Str$(Form6.Width))
   End If
   
errlog:
   subErrLog ("Form1:SaveFormPositions")
End Sub

Private Sub Form_Click()
On Error GoTo errlog
   Call Form_KeyDown(0, 0)    'this is to get the Ctrl and Shift key status when switching focus back to Form 1
   If blLoading = False Then
      blRefresh = True
      RecentFolderChange
   End If
errlog:
   subErrLog ("Form1:Form_Click")
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As ComctlLib.Panel)
On Error GoTo errlog
   blRefresh = True
   RecentFolderChange
errlog:
   subErrLog ("Form1:StatusBar1_PanelClick")
End Sub

Private Sub ProgressBar1_Click()
On Error GoTo errlog
   blRefresh = True
   RecentFolderChange
errlog:
   subErrLog ("Form1:ProgressBar1_Click")
End Sub

Private Sub Label1_Click()
On Error GoTo errlog
   blRefresh = True
   RecentFolderChange
errlog:
   subErrLog ("Form1:Label1_Click")
End Sub

Public Sub Timer2_Timer()
On Error GoTo errlog

'Timer2.Interval = 5000

   If Form1.Visible = False Then Exit Sub
   If Form1.WindowState = vbMinimized Then Exit Sub
   
   If mGlueRight.Checked = True Then
      If meleft <> Me.Left Then
         Form2.Left = Form1.Left - Form2.Width
      Else
         Form1.Left = Form2.Left + Form2.Width
      End If
   ElseIf mGlueLeft.Checked = True Then
      If meleft <> Me.Left Then
         Form2.Left = Form1.Left + Form1.Width
      Else
         Form1.Left = Form2.Left - Form1.Width
      End If
   End If
   
   If mGlueLeft.Checked = True Or mGlueRight.Checked = True Then
      If metop <> Me.Top Then
         Form2.Top = Form1.Top
      Else
         Form1.Top = Form2.Top
      End If
      If meheight <> Me.Height Then
         Form2.Height = Form1.Height
      Else
         Form1.Height = Form2.Height
      End If
'      If mewidth <> Me.Width Then
'         Form2.Width = Form1.Width
'      Else
'         Form1.Width = Form2.Width
'      End If
      meleft = Form1.Left
      metop = Form1.Top
      meheight = Form1.Height
'      mewidth = Form1.Width
   End If
   
   If blTaskbarStatus = False Then
      If mFade.Checked = True Then
         If blSuspendFade = False Then
            IsTTActive
         End If
      End If
   End If

errlog:
   subErrLog ("Form1:Timer2")
End Sub

Public Sub Timer3_Timer()
On Error GoTo errlog
   Timer3.Enabled = False
   If blMenu = True Then
      PopupMenu mnPreferences
      blMenu = False
   End If
   If Form2.ListView2.Width > 7000 And Form2.ListView2.Width < 7200 Then 'only for default autosize then
      If blScroll = True And mIcons.Checked = False And mSaveSize.Checked = False Then
'         If sType <> sLastType And blVirtualView = False Then
            ScrollBarCompensation
            blScroll = False
'         End If
      End If
   End If
errlog:
   subErrLog ("Form1:Timer3_Timer")
End Sub

Private Sub StartLogging()
On Error GoTo errlog
   If Command$ = "-t" Then
      blTraceLog = True
      If DoesFileExist(App.path + "\Trace.log") Then
         Kill App.path + "\Trace.log"
      End If
   End If
   If Command$ = "-l" Then
         SaveSetting "TaskTracker", "Settings", "Logging", "1"
   End If
   If Command$ = "-s" Then
         SaveSetting "TaskTracker", "Settings", "Snapshot", "1"
   End If
errlog:
   subErrLog ("Form1:LogStart")
End Sub

Public Function ApiTextStrip(ByVal s As String) As String
    Dim i As Long

    i = InStr(s, Chr(0)) ' Find first Null byte

    If i > 0 Then
        ApiTextStrip = Left(s, i - 1)
    Else
        ApiTextStrip = s
    End If
End Function



