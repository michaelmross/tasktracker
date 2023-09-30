VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.Ocx"
Begin VB.Form Form2 
   Caption         =   "TaskTracker"
   ClientHeight    =   6760
   ClientLeft      =   169
   ClientTop       =   559
   ClientWidth     =   6981
   ClipControls    =   0   'False
   Icon            =   "Form2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6760
   ScaleWidth      =   6981
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ListView ListView2 
      Height          =   5550
      Left            =   -15
      TabIndex        =   1
      Top             =   0
      Width           =   7020
      _ExtentX        =   12940
      _ExtentY        =   10232
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
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
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "File Name"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Last Modified"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Shortcut"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   299
      Left            =   0
      TabIndex        =   0
      Top             =   6461
      Width           =   6981
      _ExtentX        =   12317
      _ExtentY        =   527
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   14111
            MinWidth        =   14111
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
   Begin VB.CommandButton btOpen 
      Caption         =   "&Open"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   5700
      Width           =   1215
   End
   Begin VB.Label lbTip 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   120
      TabIndex        =   3
      Top             =   5750
      Width           =   5295
   End
   Begin VB.Menu mnOptions 
      Caption         =   "Options Menu"
      Visible         =   0   'False
      Begin VB.Menu mOpen 
         Caption         =   "&Open"
         Index           =   1
      End
      Begin VB.Menu mOpenWith 
         Caption         =   "Open &With..."
         Shortcut        =   ^W
      End
      Begin VB.Menu mExplorer 
         Caption         =   "Open in E&xplorer"
         Shortcut        =   ^R
      End
      Begin VB.Menu mContainer 
         Caption         =   "Open Cont&aining Folder"
         Shortcut        =   ^G
      End
      Begin VB.Menu mViewBin 
         Caption         =   "Open Rec&ycle Bin"
      End
      Begin VB.Menu mCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mPaste 
         Caption         =   "Past&e"
         Shortcut        =   ^V
      End
      Begin VB.Menu mCopyPath 
         Caption         =   "Cop&y Path"
         Shortcut        =   ^Y
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mAddVirtual 
         Caption         =   "Add to Virt&ual Folder"
         Shortcut        =   ^I
         Visible         =   0   'False
      End
      Begin VB.Menu mDelVirtual 
         Caption         =   "Remove from Virt&ual Folder"
         Shortcut        =   ^X
         Visible         =   0   'False
      End
      Begin VB.Menu mFilter 
         Caption         =   "&Filter Dates and Names..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mSearchFile 
         Caption         =   "Searc&h in File System"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mSimple 
         Caption         =   "Simple List &View"
      End
      Begin VB.Menu mPreview 
         Caption         =   "View &Image"
         Shortcut        =   ^M
         Visible         =   0   'False
      End
      Begin VB.Menu mExt 
         Caption         =   "View Ex&tensions"
         Checked         =   -1  'True
      End
      Begin VB.Menu mFileTypes 
         Caption         =   "View F&ile Types"
      End
      Begin VB.Menu mFolders 
         Caption         =   "View &Folders"
         Begin VB.Menu mFullPath 
            Caption         =   "&Folder Paths"
            Shortcut        =   ^T
         End
         Begin VB.Menu mContainerFolder 
            Caption         =   "&Containing Folders"
         End
      End
      Begin VB.Menu mDateSelect 
         Caption         =   "&Show Dates"
         Begin VB.Menu mCreated 
            Caption         =   "When C&reated"
            Shortcut        =   ^D
         End
         Begin VB.Menu mModified 
            Caption         =   "When Last &Modified"
            Checked         =   -1  'True
         End
         Begin VB.Menu mAccessed 
            Caption         =   "When Last &Accessed"
         End
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mDelShort 
         Caption         =   "&Delete This Shortcut"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mRename 
         Caption         =   "Re&name File"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mProps 
         Caption         =   "P&roperties"
         Shortcut        =   ^P
      End
      Begin VB.Menu mShowExpMenu 
         Caption         =   "Show Explorer Menu"
         Shortcut        =   ^E
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mPreferences 
         Caption         =   "&Preferences"
         Shortcut        =   ^N
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' © 2003-2007 Michael M. Ross, Wordwise Solutions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function ShellExecute Lib _
        "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hWnd As Long, ByVal lpOperation _
        As String, ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long


Private Declare Function GetSystemDirectory Lib _
        "kernel32" Alias "GetSystemDirectoryA" _
        (ByVal lpBuffer As String, ByVal nSize _
        As Long) As Long
        
Private Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long
Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters  As String
    lpDirectory   As String
    nShow As Long
    hInstApp As Long
    lpIDList      As Long
    lpClass       As String
    hkeyClass     As Long
    dwHotKey      As Long
    hIcon         As Long
    hProcess      As Long
End Type

Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
   
Private Declare Function DoExplorerMenu Lib "expmnu.dll" (ByVal hWnd As Long, ByVal sFilePath As String, ByVal x As Long, ByVal y As Long) As Boolean
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
   
Private Const LVM_FIRST = &H1000
'Private Const LVM_GETHEADER = (LVM_FIRST + 31)

Private Const HDM_FIRST = &H1200
'Private Const HDM_SETITEM = (HDM_FIRST + 4)

Private Const LVS_EX_HEADERDRAGDROP As Long = &H10
Private Const LVM_GETCOLUMNORDERARRAY = (LVM_FIRST + 59)
Private Const LVM_SETCOLUMNORDERARRAY = (LVM_FIRST + 58)

Private clsheader As cHeadIcons
Private TTS As CTooltip

Private PosArrayAddTypes() As Long
Private PosArraySpecTypes() As Long
Private PosArrayRegTypes() As Long
Private PrevArrayAddTypes() As Long
Private PrevArraySpecTypes() As Long
Private PrevArrayRegTypes() As Long
'Private KeepArrayRegTypes() As Integer
'Private KeepArraySpecTypes() As Integer
'Private KeepArrayAddTypes() As Integer

Private ColOrderReg As String
Private ColOrderSpec As String
Private ColOrderAdd As String

Private LastSubItem As Integer
Private iSortKey As Integer
'Private iSortOrder As Integer

'Public blPoponRestore As Boolean
Public blExplorer As Boolean
'Public blSpecialType As Boolean
'Public blDontPopAgain As Boolean
Private blRightClick As Boolean
Private blContainer As Boolean
Private blOpenWith As Boolean
Private blMultSel As Boolean
Private blDuds As Boolean, blGoods As Boolean
Private blShift As Boolean
Private blColAdd As Boolean
Private blColSpec As Boolean
Private blColReg As Boolean
Private blChangeComboHeight As Boolean
Private blPaste As Boolean
Private blRegTips As Boolean
Private blLoadingFiles As Boolean
Private blTemp As Boolean

Public Sub Form_Click()
   Me.SetFocus
End Sub

'Private m_lpszToolA As Long

Private Sub ListView2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo errlog

   If Form1.mHideTips.Checked = True Then Exit Sub
      
   Dim lvhti As LVHITTESTINFO
   lvhti.pt.x = x \ Screen.TwipsPerPixelX
   lvhti.pt.y = y \ Screen.TwipsPerPixelY
   Dim lItemIndex As Long
   Dim itip As Long
'   Call GetCursorPos(lvhti.pt)
'   Call ScreenToClient(ListView2.hwnd, lvhti.pt)
   lItemIndex = SendMessage(ListView2.hWnd, LVM_SUBITEMHITTEST, 0, lvhti) + 1
   
   If m_lCurItemIndex <> lItemIndex Or LastSubItem <> lvhti.iSubItem Then
      m_lCurItemIndex = lItemIndex
      LastSubItem = lvhti.iSubItem
      If m_lCurItemIndex = 0 Then  ' no item under the mouse pointer
         TT.Destroy
      Else
'         If ObscuredText(lvhti.iItem, lvhti.iSubItem, m_lpszToolA) = True Then  - not working...still trying...
         If lvhti.iSubItem = 0 Then
            Dim SIP As String
            Dim PathWithoutFilename As String
            Dim sDate As String
            itip = m_lCurItemIndex   'byref cannot be used directly
            SIP = fnSelectedItemsPath(itip)
            PathWithoutFilename = Left$(SIP, Len(SIP) - (Len(SIP) - InStrRev(SIP, "\")))
            If mModified.Checked = True Then
               sDate = "Last Modified: "
            ElseIf mAccessed.Checked = True Then
               sDate = "Last Accessed: "
            Else
               sDate = "Created: "
            End If
            If Left$(PathWithoutFilename, 2) <> "\\" Then
               Dim ResourceName As String
               ResourceName = GetNetResourceName(Left$(PathWithoutFilename, 2))
               If Len(Trim$(ResourceName)) > 0 Then
                  ResourceName = "(UNC) " + Trim$(ResourceName)
               End If
            End If
            
            'order of next 4 lines is critical - e.g. if Create is first, tips are displaced by 1 lisitem!
            TT.Title = ListView2.ListItems(itip).Text

            If AllorAddorSpecial = True Or blSearchAll = True Then
               
                If Form2.mFileTypes.Checked = False Or mSimple.Checked = True Then
                
                   blRegTips = False
                
                Else
               
                   blRegTips = True
                  
                End If
               
            Else
            
                blRegTips = True
            
            End If
            
            If blRegTips = True Then
            
                  If Len(ResourceName) > 0 Then
                     TT.TipText = sDate + ListView2.ListItems(itip).SubItems(iShortCol - 1) + vbNewLine + PathWithoutFilename _
                     + vbNewLine + ResourceName
                  Else
                     TT.TipText = sDate + ListView2.ListItems(itip).SubItems(iShortCol - 1) + vbNewLine + PathWithoutFilename
                  End If
                  
            Else
            
               '  Identify file types when mixed types can occur and File Types column is not shown
                  Call IsRegisteredType(SIP)    'to get filetype (sRegType)
                  
                  If Len(sRegType) = 0 Or sRegType = "File" Then
                     If DoesFolderExist(SIP) = True Then
                        sRegType = "File Folder"
                     Else
                        sRegType = "Unknown"
                     End If
                  End If
                  
                  If Len(ResourceName) > 0 Then
                     TT.TipText = sRegType + vbNewLine + sDate + ListView2.ListItems(itip).SubItems(iShortCol - 1) _
                     + vbNewLine + PathWithoutFilename + vbNewLine + ResourceName
                  Else
                     TT.TipText = sRegType + vbNewLine + sDate + ListView2.ListItems(itip).SubItems(iShortCol - 1) _
                     + vbNewLine + PathWithoutFilename
                  End If
                  sRegType = vbNullString

            End If
            
            TT.Create ListView2.hWnd
            TT.SetDelayTime (sdtInitial), 1000
            
         Else
            TT.Destroy
         End If
      End If
   End If
   
errlog:
   subErrLog ("Form2:ListView2_MouseMove")
End Sub

Public Sub subForm2Init()
On Error GoTo errlog

   blLoadingFiles = True
   Form2.Caption = "TaskTracker"   'temporarily set
'   Form1.blSetTempCaption = True   'wait for timer to set loading caption - it may not be needed
   
   SetContext

   If blLoading = False Then
      Me.WindowState = vbNormal
      subSetSize
      If Me.Visible = False Then
         Form1.Timer2_Timer
         Me.Show 0
         ReshowOtherWindows
      End If
   End If
   Form1.SetFocus      'let the user up/down key through the list, left/right keys switch forms

   If blLoading = False Then
'      If blDontPopAgain Then
''         blDontPopAgain = False
''      Else
         PopFromMem
'      End If
   End If
   
   Form1.blSetTempCaption = False
   
errlog:
   SetTitle
'   subLbCount - in popfrommem
   blLoadingFiles = False
   subErrLog ("Form2:subForm2Init")
End Sub

Public Sub SetContext()
On Error GoTo errlog

   If blAllTypes = True Then
      Form1.SelectAllForm1
'   ElseIf blThisType = True Then
'      blWasMultSel = False
   End If
   
errlog:
   subErrLog ("Form2:SetContext")
End Sub

Public Sub SetTitle(Optional newstring As String)
On Error GoTo errlog

   Dim Form2Caption As String
      
   sLastType = sType
   
   If blVirtual = True And blVirtualView = True Then 'virtual folder view
      
      If Len(newstring) = 0 Then
         Form2Caption = "TaskTracker - " + Form5.ListView3.SelectedItem.Text
      Else
         Form2Caption = "Tasktracker - " + newstring
      End If
      If Form4.ckSave.Value = 1 And Form4.ckSave.Enabled = True Then
         If Form5.mIgnoreFilters.Checked = False Then
            Form2.Caption = Form2Caption + " [Filtered]"
         Else
            Form2.Caption = Form2Caption
         End If
      Else
         Form2.Caption = Form2Caption
      End If
      
   Else
   
'   Set Form2 Caption
'   ******************************************************
      If blAllTypes = True Then
'         Form1.SelectAllForm1
         Form2Caption = "TaskTracker - All file types"
      ElseIf blThisType = True And blSearchAll = False Then
'         blWasMultSel = False
         Form2Caption = "TaskTracker - " + sType
      ElseIf blThisType = True And blSearchAll = True Then
'         blWasMultSel = False
         Form2Caption = "TaskTracker - Selected files"
      ElseIf blAddType = True Or (blThisType = True And blSearchAll = True) Then
'         Form4.ckAll.Value = 2
         Form2Caption = "TaskTracker - Selected file types"
      End If
'   ******************************************************
      
      If Form2.mFilter.Checked = False Then
         If Form4.ckSave.Value = 1 And Form4.ckSave.Enabled = True Then
            Form2.Caption = Form2Caption + " [Filtered]"
         Else
            Form2.Caption = Form2Caption
         End If
      Else
         Form2.Caption = Form2Caption
      End If
      
   End If
   
'   subSelTypes
   SetCaption
            
errlog:
   subErrLog ("Form2:SetTitle")
End Sub

Private Function SetCaption()
On Error GoTo errlog

   If sType = "Deleted or Renamed" And blThisType = True And blVirtualView = False Then
      btOpen.Enabled = False
'      cbDates.Enabled = False
      lbTip.Caption = "These files no longer exist or have been renamed. " + _
      "Right click on this list or the File Types list for a menu of commands and options. "
   ElseIf sType = "Network Drive" And blThisType = True And blVirtualView = False Then
      btOpen.Enabled = True
'      cbDates.Enabled = False
      lbTip.Caption = "These files are on a network drive. " + _
      "Right click on this list or the File Types list for a menu of commands and options.  "
   ElseIf sType = "CD Drive" And blThisType = True And blVirtualView = False Then
      btOpen.Enabled = True
'      cbDates.Enabled = False
      lbTip.Caption = "These files are on a CD drive. " + _
      "Right click on this list or the File Types list for a menu of commands and options. "
   ElseIf sType = "Removable Drive" And blThisType = True And blVirtualView = False Then
      btOpen.Enabled = True
'      cbDates.Enabled = False
      lbTip.Caption = "These files are on a removable drive. " + _
      "Right click on this list or the File Types list for a menu of commands and options. "
   Else
      btOpen.Enabled = True
'      cbDates.Enabled = True
      lbTip.Caption = "You can open one or several files or folders at a time. " + _
      "Right click on this list or the File Types list for a menu of commands and options. "
   End If
   
errlog:
subErrLog ("Form2:SetCaption")
End Function


'default button - captures Enter key
Private Sub btOpen_Click()

  If blClicked = False Then
    Exit Sub
  End If
'  blContainer = False
  OpenFiles
   
End Sub

Private Sub OpenFiles()
On Error GoTo errlog

  Dim sh As New Shell
  Dim sPath As String
  Dim sl As Long
  Dim blFailed As Boolean
  Dim blOpenedOne As Boolean
  
  With ListView2
      For sl = 1 To .ListItems.Count
         If .ListItems(sl).Selected = True Then
            If Len(.ListItems(sl).Text) > 0 Then
            
               sPath = fnSelectedItemsPath(sl)
               
'********************************************Open File or Folder******************************************************
               If blContainer = False Then
                  If DoesFileExist(sPath) = False Then
                     If sType = "Network Drive" Then
                        MsgBox "TaskTracker cannot open " + sPath + "." + vbNewLine + _
                               "The network drive is not available.", vbInformation
                        blFailed = True
'                        Exit For
                     ElseIf sType = "Deleted or Renamed" Then
                        blFailed = True
'                        Exit For
                     Else
                        MsgBox "TaskTracker cannot find " + sPath + "." + vbNewLine + _
                               "You may have just moved, renamed, or deleted it.", vbInformation
                        subForm2Init
                        subLbCount
                        blFailed = True
'                        Exit For
                     End If
                  Else
                     blOpenedOne = True
                     If blOpenWith = False Then
                        If blExplorer = False Then
                           sh.Open sPath
                        ElseIf DoesFolderExist(sPath) = False Then
                           sh.Open sPath
                        Else
                           Shell Environ("WINDIR") + "\explorer.exe  /e," + sPath, vbNormalFocus
                        End If
                     Else
                        OpenWithDialog (sPath)
                     End If
                  End If
'********************************************Open Container Folder******************************************************
               Else
                  If blOpenWith = False Then
                    
                     If DoesFileExist(sPath) = False Then   'if container is for a deleted file
                     
                        sPath = Left$(sPath, Len(sPath) - (Len(sPath) - InStrRev(sPath, "\")))
                        
                        If DoesFolderExist(sPath) = True Then
                           blOpenedOne = True
                           If blExplorer = True Then
                              Shell Environ("WINDIR") + "\explorer.exe  /e," + sPath, vbNormalFocus
                           Else
                              sh.Open Left$(sPath, Len(sPath) - (Len(sPath) - InStrRev(sPath, "\")))
                           End If
                        Else
                           MsgBox "The containing folder has been moved, renamed, or deleted.", vbInformation
                           blFailed = True
'                           Exit For
                        End If
                        
                     Else

                        blOpenedOne = True
                        If blExplorer = True Then 'Open with Explorer
                           Shell Environ("WINDIR") + "\explorer.exe  /e,/select," + sPath, vbNormalFocus
                        Else
                           sh.Open Left$(sPath, Len(sPath) - (Len(sPath) - InStrRev(sPath, "\")))
                        End If
                        
                     End If
                  End If
               End If
            Else
               MsgBox "Select a file to open.", vbInformation
               blFailed = True
            End If
         End If
         If blLoadingFiles = True Then
            If blOpenedOne = True Then    'while loading just open one file
               Exit For
            End If
         End If
      Next sl
      'Hide with ALT key or Autohide but not both
      If Form1.mSingleClick.Checked = True Then
         If blPaste = False Then
            If blFailed = False Then
               If blAltPress = True And Form1.mAutohide.Checked = False Or _
                  blAltPress = False And Form1.mAutohide.Checked = True Then
                  If blTaskbarStatus = False Then
                     Form1.subHideShow
                  Else
                     Form1.WindowState = vbMinimized
                     Form1.HideForms
                  End If
               End If
            End If
         End If
      End If
   End With
      
errlog:
   blAltPress = False
   blContainer = False
   blPaste = False
   blShift = False
   subErrLog ("Form2:OpenFiles")
End Sub

Public Function fnSelectedItemsPath(sl As Long) As String
On Error GoTo errlog
Dim nS As Integer
   For nS = 0 To UBound(aNumShort)
      If ListView2.ListItems(sl).SubItems(iShortCol) = aNumShort(nS) Then
         fnSelectedItemsPath = aNumPath(nS)
         Exit For
      End If
   Next nS
errlog:
subErrLog ("Form2:fnSelectedItemsPath")
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errlog

   Dim blF10 As Boolean
   Dim alttest As Integer

   If KeyCode = vbKeyF5 Then
      Form1.mRefresh_Click
   End If
'   If blAltPress = False Then
      If KeyCode = vbKeyReturn Then
         mOpen_Click (1)
      End If
'   End If
   If KeyCode = vbKeyDelete Then
      If blVirtualView = True Then
         mDelVirtual_Click
      Else
         If Me.ActiveControl Is ListView2 Then
            mDelShort_Click
         End If
      End If
   End If
   If sType <> "Deleted or Renamed" And sType <> "Unknown" Then
      If KeyCode = vbKeyF2 Then
         ListView2.StartLabelEdit
      End If
   End If
   If KeyCode = vbKeyShift Then
     blShift = True
   End If
   If KeyCode = vbKeyF10 Then
      blF10 = True
   End If
   If blShift = True And blF10 = True Then
      subRC2Menu
      blF10 = False
   End If
   If KeyCode = vbKeyEscape Then
      Call Form_QueryUnload(1, 0)
   End If
   If KeyCode = vbKeyF1 Then
      Form1.mHelp_Click
   End If
   
'   If blAltPress Then
'      If KeyCode = vbKeyReturn Then
'         mProps_Click
'      End If
'   End If
   
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
subErrLog ("Form2:Form_KeyDown")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo errlog

   If blLoading = True Then Exit Sub

   If KeyAscii = 1 Then   'ctrl+A - select all
      SelectAllForm2
   End If
   If KeyAscii = 3 Then    'ctrl+C
      mCopy_Click
   End If
   If KeyAscii = 5 Then    'ctrl+E
      If blRegStatus = True Then
         mShowExpMenu_Click
      End If
   End If
   If KeyAscii = 6 Then    'ctrl+F
      mFilter_Click
   End If
   If KeyAscii = 7 Then   'ctrl+G
'      If sType <> "File Folder" Then
         blExplorer = False   'restore default container view
'      End If
      mContainer_click
   End If
   If KeyAscii = 13 Then  'ctrl+M
      mPreview_Click
   End If
   If KeyAscii = 14 Then  'ctrl+n
      Form1.mPreferences_Click
   End If
   If KeyAscii = 16 Then  'ctrl+P
      mProps_Click
   End If
   If KeyAscii = 18 Then  'ctrl+R
'      If sType <> "File Folder" Then
         blContainer = True      'open containing folder in explorer view
'      End If
      mExplorer_Click
   End If
   If KeyAscii = 22 Then  'ctrl+V
      mPaste_Click
      DoEvents             'bugfix only needed for keyboard
      SendKeys "+{INSERT}" 'second paste does the trick
   End If
   If KeyAscii = 23 Then  'ctrl+W
      mOpenWith_Click
   End If
   If KeyAscii = 25 Then  'ctrl+Y
      mCopyPath_Click
   End If
'   If KeyAscii = 20 Then  'ctrl+T
'      mCreated_Click
'   End If
'   If KeyAscii = 13 Then  'ctrl+M
'      mModified_Click
'   End If
'   If KeyAscii = 5 Then  'ctrl+E
'      mAccessed_Click
'   End If
   If KeyAscii = 24 Then  'ctrl+X
      If blVirtualView = True Then
         mDelVirtual_Click
      End If
   End If
   If KeyAscii = 21 Then  'ctrl+U
      Form1.mVirtual_Click
   End If
   If KeyAscii = 9 Then  'ctrl+I
      If Form5.Visible = True Then
         mAddVirtual_Click
      End If
   End If
   If KeyAscii = 20 Then  'ctrl+T   - 3 way
      If mFullPath.Checked = False And mContainerFolder.Checked = False Then
         mContainerFolder_Click
      ElseIf mContainerFolder.Checked = True Then
         mFullPath_Click
      Else
         mFullPath_Click   'unselect both - no folder view
      End If
   End If
   If KeyAscii = 4 Then   'ctrl+D   - 3 way
      If mModified.Checked = True Then
        mAccessed_Click
      ElseIf mAccessed.Checked = True Then
        mCreated_Click
      Else
        mModified_Click
      End If
   End If
   
   blCtrlKey = False
   
errlog:
   subErrLog ("Form2:KeyPress")
End Sub

Public Sub SelectAllForm2()
On Error GoTo errlog

   Dim lst As Integer
   With ListView2
      For lst = 1 To .ListItems.Count
         Set .SelectedItem = .ListItems(lst)
      Next lst
   End With
   
errlog:
   subErrLog ("Form2:SelectAllForm2")
End Sub

'Public Sub UnSelectAllForm2()
'On Error GoTo errlog
'
'   Dim lst As Integer
'   With ListView2
'      For lst = 1 To .ListItems.Count
'         .SelectedItem.Selected = False
'      Next lst
'   End With
'
'errlog:
'   subErrLog ("Form2:UnSelectAllForm2")
'End Sub

'Private Sub Form_Initialize()
'    InitCommonControls
'   'set the new listview style to full-row select
'End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If blAddType = True Then
      Form1.SelectionStatus
      Form1.ListView1_Click
   End If
   blShift = False
   blAltPress = False
End Sub
    
Private Sub Form_Load()
On Error GoTo errlog

   If Command$ = "-m" Then
      Me.WindowState = vbMinimized
      blMinStart = True
   End If
 
   GetSettings
   
'   subSetSize
   
   Dim lOldStyle As Long, lNewStyle As Long
   Dim lresult As Long
   
   'Set full row select style
   lOldStyle = SendMessage(ListView2.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, ByVal 0&)
   lNewStyle = lOldStyle Or LVS_EX_FULLROWSELECT
   lresult = SendMessage(ListView2.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, ByVal lNewStyle)
         
   Set TTS = New CTooltip
   
'   Set clsheader = New cHeadIcons
'   Set clsheader.ListView = ListView2
'   Call clsheader.SetHeaderIcons(0, Ascn)
   
'   m_lpszToolA = LocalAlloc(LPTR, MAX_LVITEM)
'   If (m_lpszToolA = 0) Then Call LocalFree(m_lpszToolA)
      
errlog:
   subErrLog ("Form2:Form_Load")
End Sub

'Private Sub TaskbarStatus()
'   If blTaskbarStatus = False Then
'      SetWindowLong Form2.hwnd, GWL_EXSTYLE, (GetWindowLong(hwnd, _
'        GWL_EXSTYLE) And Not WS_EX_APPWINDOW)
'   Else
''       Dim ll_Style As Long
''
''      SetWindowLong Form2.hwnd, GWL_EXSTYLE, (GetWindowLong(hwnd, _
''        GWL_EXSTYLE) Or WS_EX_APPWINDOW)
''      'Get window style
''      ll_Style = GetWindowLong(Form2.hwnd, GWL_STYLE)
''      'Add the minimize button
''      Call SetWindowLong(Form2.hwnd, GWL_STYLE, ll_Style Or WS_MINIMIZEBOX)
'   End If
''   Me.Show 0
''   Me.Visible = False
'End Sub

Private Sub Form_Mouseup(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then
      subRC2Menu
   Else
      blRightClick = False
   End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next

   If blExit = True Or UnloadMode = vbAppWindows Then
      Exit Sub
   End If
   
   If blBuilding Then
      blStop = True
      Exit Sub
   End If
   
   If blTaskbarStatus = False Then
      Cancel = 1
'      If blLoading = False Or Form1.blReloading = True Then
         Me.Hide
         Form4.Hide
         If Form1.mVirtual.Checked Then
            Form5.Hide
         End If
         Form6.Hide
         If Form1.mSingleClick.Checked = True Then
             Form1.Hide
         End If
'      End If
   Else
      Cancel = 1
      If blLoading = False Then
         Me.Hide
         If Form1.mSingleClick.Checked = True Then
            Form1.WindowState = vbMinimized
         Else
            Exit Sub
         End If
      End If
   End If
   
   Form1.Form_Resize
   
errlog:
subErrLog ("Form2:Form_QueryUnload")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
End Sub

Private Sub lbTip_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then
      subRC2Menu
   End If
End Sub

Private Sub ListView2_AfterLabelEdit(Cancel As Integer, newstring As String)
On Error GoTo errlog
   
   If Len(newstring) = 0 Then
      MsgBox "You must enter a valid file name.", vbExclamation
      Cancel = True
      Exit Sub
   End If

   Dim fn As Integer
   Dim sPath As String, sNewPath As String, sOldPath As String, NameString As String, NoExtString As String
   
   NameString = newstring  'have to change var name to avoid extension when no extension wanted
   
   If mExt.Checked = False Then
      If Len(sExt) > 0 Then
         NoExtString = NameString
         NameString = NameString + "." + sExt
      End If
   End If

   sPath = fnGetPath
  
   sNewPath = Left$(sPath, Len(sPath) - (Len(sPath) - InStrRev(sPath, "\"))) + NameString
   
   sOldPath = fnRenamePath
   
   If DoesFileExist(sOldPath) = False Then
      MsgBox "This file has been moved, deleted, or already renamed.", vbExclamation
      Cancel = True
      Exit Sub
   ElseIf DoesFileExist(sNewPath) = True Then
      MsgBox "A file with this name already exists in this folder." + vbNewLine + _
      "Choose another name or move/copy to another folder.", vbInformation
      Cancel = True
      Exit Sub
   End If
        
   'update the path array
   Dim nS As Integer
   For nS = 0 To UBound(aNumShort)
      If ListView2.ListItems(nSelected).SubItems(iShortCol) = aNumShort(nS) Then
         aNumPath(nS) = sNewPath
         Exit For
      End If
   Next nS
   
   Name sOldPath As sNewPath
   
'   Dim lResult As Long
'   lResult = SHAddToRecentDocs(SHARD_PATH, sNewPath)

'  Search for existing name in filename array and replace
   For fn = 0 To TTCnt
      If aSFileName(fn) = sOldPath Then
         If mExt.Checked = True Then
            aSFileName(fn) = sNewPath
         Else
            aSFileName(fn) = NoExtString
         End If
         Exit For
      End If
   Next fn
'   If Len(NoExtString) > 0 Then
'       ListView2.SelectedItem.Text = NoExtString
'   End If
   
Exit Sub
errlog:
   MsgBox "An error occurred. Check the validity of your filename.", vbExclamation
   Cancel = True
   subErrLog ("Form2:Listview2_AfterLabelEdit")
End Sub

Public Sub ListView2_Click()
On Error GoTo errlog

   blDragOut = False
   
   If blMultSel = True Then
      blMultSel = False
      Exit Sub
   End If
   If ListView2.ListItems.Count > 0 Then
      If blRightClick = True Then
'         ListView2.SelectedItem.Selected = True
         blRightClick = False
      Else
         blClicked = True
      End If
      SaveSelected
   End If
   
errlog:
subErrLog ("Form2:ListView2_Click")
End Sub

Private Sub LVXHeaderDragDrop()

   Dim lOldStyle As Long, lNewStyle As Long, lresult As Long
   
   lOldStyle = SendMessage(ListView2.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, ByVal 0&)
   lNewStyle = lOldStyle Or LVS_EX_HEADERDRAGDROP
   lresult = SendMessage(ListView2.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, ByVal lNewStyle)
                        
End Sub

Public Sub ListView2_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
On Error Resume Next

   If blLoading = True Then
      Exit Sub
   End If
   
   Dim blSpecialType As Boolean
      
   blColumnClick = True
    
   LVXHeaderDragDrop       'rearranging columns
   
   With ListView2
           
      .Sorted = True
      If blAllTypes = True Or blAddType = True And blVirtualView = False Then
         If blFolderSort = True Then
            .SortKey = MultPSortCol
         Else
            .SortKey = MultSortCol
         End If
      ElseIf blThisType = True Then    'And blVirtualView = False
         If sType = "Deleted or Renamed" Or sType = "Network Drive" _
             Or sType = "Removable Drive" Or sType = "CD Drive" Or blMultSelect = True Then
            blSpecialType = True
            If blFolderSort = True Then
               .SortKey = DelPSortCol
            Else
               .SortKey = DelSortCol
            End If
         Else
            If blFolderSort = True Then
               .SortKey = RegPSortCol
            Else
               .SortKey = RegSortCol
            End If
         End If
      End If

      '0 is the name column, 1 or 2 is the date column or type column or folder column
      'first time showing form2 or event requiring resort
      If blFirstSort = True Or blDateSwitch = True Then
         
         'retain sort order
         If blSpecialType = True Then
            If blFolderSort = True Then
               .SortOrder = DelPSortOrder
            Else
               .SortOrder = DelSortOrder
            End If
         Else
           If blAllTypes = True Or blAddType = True Then
              If blFolderSort = True Then
                .SortOrder = MultPSortOrder
              Else
                .SortOrder = MultSortOrder
              End If
           ElseIf blThisType = True Then
              If blFolderSort = True Then
                .SortOrder = RegPSortOrder
              Else
                .SortOrder = RegSortOrder
              End If
           End If
         End If
      
      Else     'type is the same
            
            'when switching columns...
            If .SortKey <> ColumnHeader.Index - 1 Then   'if column clicked is not sortkey
            
               If Left$(.ColumnHeaders(ColumnHeader.Index).Key, 4) = "date" Then
                  .SortOrder = lvwDescending    'reverse order for date
               Else
                  .SortOrder = lvwAscending     'for alpha columns
               End If
           Else      'alternate ascending/descending for same type
               If .SortOrder = lvwDescending Then        'if sortkey column is clicked twice
                  .SortOrder = lvwAscending
               Else
                  .SortOrder = lvwDescending
               End If
            End If
            .SortKey = ColumnHeader.Index - 1
                                              
      End If
      
      FileSortOrder = .SortOrder
      iSortKey = .SortKey
      
      ApplyHeaderIcon
      
      If blAllTypes = True Or blAddType = True And blVirtualView = False Then
         If blFolderSort = False Then
            MultSortOrder = .SortOrder
            MultSortCol = .SortKey
         Else
            MultPSortOrder = .SortOrder
            MultPSortCol = .SortKey
         End If
      ElseIf blThisType = True Then
         If sType = "Deleted or Renamed" Or sType = "Network Drive" _
            Or sType = "Removable Drive" Or sType = "CD Drive" Or blMultSelect = True Then
            If blFolderSort = True Then
               DelPSortOrder = .SortOrder
               DelPSortCol = .SortKey
            Else
               DelSortOrder = .SortOrder
               DelSortCol = .SortKey
            End If
         Else
            If blFolderSort = True Then
               RegPSortOrder = .SortOrder
               RegPSortCol = .SortKey
            Else
               RegSortOrder = .SortOrder
               RegSortCol = .SortKey
            End If
         End If
      End If

      If Left$(.ColumnHeaders(.SortKey + 1).Key, 4) = "date" Then
          DateSortWrapper
      End If
          
      .ListItems(.SelectedItem.Index).EnsureVisible
   
   End With
   
   If blTemp Then
      blTemp = False
      ListView2.ColumnHeaders.Remove ("temp")
   End If
   
errlog:
   blColumnClick = False
   blDateSwitch = False
   blFirstSort = False
   subErrLog ("Form2:Listview2_ColumnClick")
End Sub

Private Sub DateSortWrapper()
On Error GoTo errlog

  With ListView2
      .Sorted = True
   
      'api callback routine required for dates only
      SendMessage ListView2.hWnd, _
                   LVM_SORTITEMS, _
                   ListView2.hWnd, _
                   ByVal FARPROC(AddressOf CompareDates)
                   
      'this trick resets the listview indexes after the callback sort
      Dim clm As ColumnHeader
      Set clm = .ColumnHeaders.Add(, "temp", , 0)
      blTemp = True
      If .SortKey = 1 Then
         .SortKey = 3
      ElseIf .SortKey = 2 Then
         .SortKey = 4  'must be empty subitem
      ElseIf .SortKey = 3 Then
         .SortKey = 5   'ditto - if you don't do this, OpenFiles will open the wrong files and folders!
      ElseIf .SortKey = 4 Then
         .SortKey = 6  'ditto...
      End If
  End With
  
errlog:
   subErrLog ("Form2:DateSortWrapper")
End Sub

Public Sub subLbCount()
On Error GoTo errlog
Dim speriod As String

   If Form4.blFiltered = True Or Form4.blSaveFilter = True Then
      speriod = " for this period."
   Else
      speriod = "."
   End If

   If blVirtual = False Then
      If Len(sType) > 0 Then
        If ListView2.ListItems.Count = 0 Then
            StatusBar1.Panels(1).Text = "No files found" + speriod + sInterrupted
            btOpen.Enabled = False
        Else
            SetCaption
        End If
        If blSearchAll = True Then
            If ListView2.ListItems.Count > 1 Then
              StatusBar1.Panels(1).Text = Str$(ListView2.ListItems.Count) & " selected files found" + speriod + sInterrupted
            ElseIf ListView2.ListItems.Count = 1 Then
              StatusBar1.Panels(1).Text = "1 selected file found" + speriod + sInterrupted
            End If
        ElseIf blThisType Then
            Dim sFiles As String
            If Right(LCase(sType), 4) = "file" Then
               If ListView2.ListItems.Count = 1 Then
                  sFiles = ""
               Else
                  sFiles = "s"
               End If
            ElseIf Right(LCase(sType), 6) = "folder" Then
               If ListView2.ListItems.Count = 1 Then
                  sFiles = ""
               Else
                  sFiles = "s"
               End If
            Else
               If ListView2.ListItems.Count = 1 Then
                  sFiles = " file"
               Else
                  sFiles = " files"
               End If
            End If
            If ListView2.ListItems.Count > 1 Then
   
               StatusBar1.Panels(1).Text = Str$(ListView2.ListItems.Count) & " " & sType & sFiles & " found" + speriod + sInterrupted
              
            ElseIf ListView2.ListItems.Count = 1 Then
            
               StatusBar1.Panels(1).Text = "1 " & sType & sFiles + " found" + speriod + sInterrupted

            End If
        ElseIf blAllTypes = True Then
            StatusBar1.Panels(1).Text = Str$(ListView2.ListItems.Count) & " files of all types found" + speriod + sInterrupted
        ElseIf blAddType = True Then
            StatusBar1.Panels(1).Text = Str$(ListView2.ListItems.Count) & " files of selected types found" + speriod + sInterrupted
        End If
        If blThisType = True Then
            Form1.ListView1.ListItems(sType).SubItems(1) = Trim$(Str$(ListView2.ListItems.Count))
        End If
      End If
   Else
      If ListView2.ListItems.Count > 1 Then
         StatusBar1.Panels(1).Text = Str$(ListView2.ListItems.Count) & " files in virtual folder."
      ElseIf ListView2.ListItems.Count = 1 Then
         StatusBar1.Panels(1).Text = "1 file in virtual folder."
      Else
         StatusBar1.Panels(1).Text = "The virtual folder is empty."
      End If
   End If
   
errlog:
   subErrLog ("Form2:subLbCount")
End Sub
   
Private Sub Listview2_DblClick()
   
   btOpen_Click
   
End Sub
   
Public Sub Form_Resize()
   On Error Resume Next
   
   Form2Sizing
   If Form2.Visible = False Then
      Form4.Hide
      If Form1.mVirtual.Checked = True Then
         Form5.Hide
      End If
      Form6.Hide
   '   Else
   '      Form4.Visible = True
   End If
   
   If blForm1Resize = True Then Exit Sub     'designed to stop feedback loop with form1
   
   If blLoading = True Then
      Exit Sub
   End If
   
   blForm2Resize = True
   
   If Me.Visible = False Or Me.WindowState = vbMinimized Then
      If Form1.mSingleClick.Checked = True Then
         blForm2Resize = True
         Form1.WindowState = Me.WindowState
      End If
      Form1.SysTrayStatus
      blForm2Resize = False
      Exit Sub
   End If
   
   If Me.WindowState = vbNormal Then
      Me.Show
      If Form1.mSingleClick.Checked = True Then
         Form1.WindowState = Me.WindowState
      End If
   End If
   
   Form1.SysTrayStatus
   
   Form2Sizing
        
   If blLoading = False Then
      If Form1.WindowState <> vbMinimized And Form2.WindowState <> vbMinimized Then
         If Form1.mGlueLeft.Checked = True Or Form1.mGlueRight.Checked = True Then
            If blLoadResize = False Then
               Form1.Height = Form2.Height
               Form1.Top = Form2.Top
            Else
               Form2.Height = Form1.Height
               Form2.Top = Form1.Top
               blLoadResize = False
            End If
         End If
      End If
   End If
   Me.Refresh
   
   '   Form4.WindowState = Form2.WindowState
   
   blForm2Resize = False
   
errlog:
   '   If Err.Number <> 0 And Err.Number <> 380 Then
   '      subErrLog ("Form2:Form_Resize")
   '   End If
      subErrLog ("Form2:Form_Resize")
End Sub

Private Sub Form2Sizing()
On Error GoTo errlog

   ListView2.Width = Me.Width - 100
   If Me.Width > 7110 Then
      btOpen.Left = Me.Width - 1450
'      cbSearch.Left = Me.Width - 4148
'      cbSearch.Width = 2550
'      cbDates.Width = 2550
'      ckAll.Left = cbSearch.Left + cbSearch.Width + 150
      StatusBar1.Panels(1).Width = Me.Width
      lbTip.Width = Me.Width - 1800
      lbTip.Height = 500
      lbTip.Top = Me.Height - 1325
   ElseIf Me.Width > 5300 Then
'      If Form1.mSaveSize.Checked = False Then
'         Dim ch As Integer
'         With ListView2
'            For ch = 1 To .ColumnHeaders.Count
'            .ColumnHeaders(ch).Width = .Width\.ColumnHeaders(ch).Width * .Width
'            Next ch
'         End With
'      End If
      btOpen.Left = Me.Width - 1450
'      cbDates.Width = Form2.Width\2.788
'      cbSearch.Width = Form2.Width\2.788
'      cbSearch.Left = Form2.Width\2.4
'      ckAll.Left = cbSearch.Left + cbSearch.Width + 150
      lbTip.Width = Me.Width - 1750
      If Me.Width < 6650 Then
         lbTip.Top = Me.Height - 1425
         lbTip.Height = 750
      Else
         lbTip.Height = 500
         lbTip.Top = Me.Height - 1325
      End If
   Else
      btOpen.Left = 3850
'      cbDates.Width = 1901
'      cbSearch.Width = 1901
'      cbSearch.Left = 2208
'      ckAll.Left = 4259
      lbTip.Width = 3500
      lbTip.Height = 750
      lbTip.Top = Me.Height - 1425
   End If
      
   If Me.Height > 2500 Then
'      ListView2.Top = 600
      ListView2.Height = Me.Height - 1500
      btOpen.Top = Me.Height - 1365
   Else
      ListView2.ZOrder
      StatusBar1.ZOrder
   End If
   If Me.Height < 2000 Then
      lbTip.Visible = False
   Else
      lbTip.Visible = True
   End If
   
'   Dim colw As Long
'   colw = ListView2.ColumnHeaders.Count - 1     'ignore shortcut col
''   Call SendMessage(ListView2.hwnd, LVM_SETCOLUMNWIDTH, ListView2.ColumnHeaders.Count, ByVal 0)
'   Call SendMessage(ListView2.hwnd, LVM_SETCOLUMNWIDTH, colw, ByVal LVSCW_AUTOSIZE_USEHEADER)
''   Call SendMessage(ListView2.hwnd, LVM_SETCOLUMNWIDTH, colw + 1, 0)
   
errlog:
   subErrLog ("Form2:Form2Sizing")
End Sub


  ' See: "FIX: Problem with ListView's ColumnHeader Width Property"
  ' http://support.microsoft.com/support/kb/articles/q179/9/88.asp
  ' and: "HOWTO: Set the Column Width of Columns in a ListView Control"
  ' http://support.microsoft.com/support/kb/articles/q147/6/66.asp

'other form2sizing "formulas"
'      cbDates.Width = Form2.Width \ (2.62 * 7110 \ Me.Width)
'      cbSearch.Width = Form2.Width \ (2.62 * 7110 \ Me.Width)
'      cbSearch.Left = Form2.Width \ (2.32 * 7110 \ Me.Width)
'      cbDates.Width = Form2.Width - 4360
'      cbSearch.Width = Form2.Width - 4360
'      cbSearch.Left = Form2.Width - 4050
'      cbSearch.Width = Form2.Width \ 2.58
'      cbSearch.Left = Form2.Width \ 2.32
'      cbSearch.Width = Form2.Width - (Form2.Width \ 1.6)
'      cbDates.Width = Form2.Width \ 2.58
'      If Me.Width > 6500 Then
'         cbDates.Width = Form2.Width \ 2.58
'         cbSearch.Width = Form2.Width \ 2.58
'         cbSearch.Left = Form2.Width \ 2.32
'      Else
'         cbDates.Width = Form2.Width \ 2.7
'         cbSearch.Width = Form2.Width \ 2.7
'         cbSearch.Left = Form2.Width \ 2.7
'      End If
'      cbDates.Width = Form2.Width \ 2.58
'      cbSearch.Width = Form2.Width \ 2.58
'      cbSearch.Left = Form2.Width \ 2.32

Private Sub ListView2_GotFocus()
   ListView2_Click
End Sub

Private Sub ListView2_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo errlog

   SaveSelected
   
   If blLoading = False Then
      If Form1.mAutoPreview.Checked = True Or mPreview.Checked = True Then
         blForm6Load = True
         Form6.Timer1.Enabled = True
      End If
   End If
   
   If Form1.mGlueLeft.Checked = False Then
      If KeyCode = vbKeyRight Then
         Form1.SetFocus
      End If
   ElseIf Form1.mGlueLeft.Checked = True Then
      If KeyCode = vbKeyLeft Then
         Form1.SetFocus
      End If
   End If
   
errlog:
subErrLog ("Form2:ListView2_KeyUp")
End Sub

Public Function fnGetPath(Optional Index As Integer) As String
On Error GoTo errlog

  Dim sl As Long
  With ListView2
      If Index > 0 Then
         sl = Index
      Else
         sl = nSelected
      End If
      
      If Len(.ListItems(sl).Text) > 0 Then
      
         fnGetPath = fnSelectedItemsPath(sl)
         
      End If
      
   End With
   
errlog:
subErrLog ("Form2:fnGetPath")
End Function

'bugfix function doesn't use nSelected which is unreliable with LabelEdit
Private Function fnRenamePath() As String
On Error GoTo errlog

  With ListView2
      If Len(.SelectedItem.Text) > 0 Then
      
          fnRenamePath = fnSelectedItemsPath(nSelected)
          
      End If
   End With
   
errlog:
subErrLog ("Form2:fnGetPath")
End Function

Private Sub ListView2_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo errlog
   blform2 = True
   Call DragDropRoutine(Data, Effect, blform2)
   Form2.SetFocus
'   ListView2.MousePointer = ccDefault
errlog:
   subErrLog ("Form2:ListView2_OLEDragDrop")
End Sub

'just open the target folder
Public Sub PerhapsaMailAttachment()
On Error GoTo errlog

   If VerifyFolderType = False Then
      blContainer = True
   End If
   OpenFiles
   
errlog:
   subErrLog ("Form2:PerhapsaMailAttachment")
End Sub

Private Sub ListView2_OLEDragOver(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, state As Integer)
On Error GoTo NoFileDrop

   If Data.Files.Count > 0 Then     'raises error for mail attachments
      Exit Sub
   End If
   
NoFileDrop:    'only interested in nonfile drops (most likely mail attachments)

   Dim lvhti As LVHITTESTINFO       'track the drop target
   Dim lItemIndex As Long
   Dim lv2 As Integer
   
   lvhti.pt.x = x \ Screen.TwipsPerPixelX
   lvhti.pt.y = y \ Screen.TwipsPerPixelY
   lItemIndex = SendMessage(ListView2.hWnd, LVM_HITTEST, 0, lvhti) + 1
      
   If lItemIndex > 0 Then
      For lv2 = 1 To ListView2.ListItems.Count
         ListView2.ListItems(lv2).Selected = False
      Next lv2
      ListView2.ListItems(lItemIndex).Selected = True    'selected folder/file (container) for drop
   End If
   
'   Call DeterminePointer(Data, Effect)
   
   subErrLog ("Form5: ListView2_OleDragOver")
End Sub

'Private Sub DeterminePointer(Data As ComctlLib.DataObject, Effect As Long)
'On Error Resume Next
'
'   ListView2.MousePointer = ccNoDrop
'
'   If Data.GetData(vbCFText) = True Then
'      ListView2.OLEDropMode = ccOLEDropManual
'   End If
'
'   If Err.Number <> 461 Then
'      If InStr(Data.GetData(vbCFText), "redir?url=file%3A%2F%2F") > 0 Then   'google desktop file
'      ListView2.OLEDropMode = ccOLEDropManual
'      End If
'   End If
'End Sub

Private Sub ListView2_OLESetData(Data As ComctlLib.DataObject, DataFormat As Integer)
On Error GoTo errlog

   If blVirtualDrag = True Then Exit Sub  'bugfix

   Dim sl As Integer
   For sl = 1 To ListView2.ListItems.Count
      If ListView2.ListItems(sl).Selected = True Then
         If sType <> "Deleted or Renamed" And Len(ListView2.ListItems(sl).SubItems(1)) > 0 _
            And ListView2.ListItems(sl).SubItems(1) <> "Deleted or Renamed" Then
            Dim gp As String
            gp = fnGetPath(sl)
            If DoesFileExist(gp) Then
               Data.Files.Add gp
               blGoods = True
            Else
               blDuds = True
            End If
         Else
            blDuds = True
         End If
     End If
   Next sl
   
errlog:
   subErrLog ("Form2:ListView2_OLESetData")
End Sub

Private Sub ListView2_OLEStartDrag(Data As ComctlLib.DataObject, AllowedEffects As Long)
   If sType = "Deleted or Renamed" Then
      Exit Sub
   End If
   blDragOut = True
   AllowedEffects = vbDropEffectCopy
   Data.SetData , vbCFFiles
   Form5.KeepForm5Visible
End Sub

'Dragging out
Private Sub ListView2_OLECompleteDrag(Effect As Long)
On Error GoTo errlog
   If blVirtualDrag = True Then Exit Sub 'bugfix
   If sType = "Deleted or Renamed" Then
      Exit Sub
   End If
   Screen.MousePointer = vbDefault
   If blDuds = True And blGoods = True Then
      MsgBox "TaskTracker cannot find some of the files you are dragging."
      Screen.MousePointer = vbDefault
   ElseIf blDuds = True And blGoods = False Then
      MsgBox "TaskTracker cannot find the file(s) you are dragging."
      Screen.MousePointer = vbDefault
   End If
   blGoods = False
   blDuds = False
   If Form5.blTransparent Then
      Call SetWindowPos(Form5.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME)
      subRestore
      Form5.blTransparent = False
   End If
errlog:
   blDragOut = False
   subErrLog ("Form2:ListView2_OLECompleteDrag")
End Sub

Public Sub mAccessed_Click()
On Error GoTo errlog
   mModified.Checked = False
   mAccessed.Checked = True
   mCreated.Checked = False
   blDateSwitch = True
'   subFirstSel
   Erase blTrueDates
   subForm2Init
errlog:
   subErrLog ("Form2:mAccessed_Click")
End Sub

Private Sub mAddVirtual_Click()
On Error GoTo errlog
    blClickVirtualAdd = True
    Form5.InitVirtualFolder
    blClickVirtualAdd = False
errlog:
   subErrLog ("Form2:mAddVirtual_Click")
End Sub

Private Sub mContainerFolder_Click()
On Error GoTo errlog
   If mContainerFolder.Checked = False Then
      mContainerFolder.Checked = True
      blFolderSort = True
      mFullPath.Checked = False
   Else
      mContainerFolder.Checked = False
      blFolderSort = False
   End If
   blDontSaveCols = True
   FolderSort
   subForm2Init
errlog:
   subErrLog ("Form2:mContainerFolder_Click")
End Sub

Private Sub mCopy_Click()
On Error GoTo errlog

   Dim sl As Integer
   Dim DL As Integer
   Dim PathstoCopy() As String
   For sl = 1 To ListView2.ListItems.Count
      If ListView2.ListItems(sl).Selected = True Then
         
         ReDim Preserve PathstoCopy(sl - 1)
         PathstoCopy(DL) = fnGetPath(sl)
         DL = DL + 1
      End If
   Next sl
   ClipboardCopyFiles PathstoCopy
   
errlog:
   subErrLog ("Form2:mCopy_Click")
End Sub

Public Function ShowFileProperties(filename As String, OwnerhWnd As Long, props As Boolean) As Long
On Error GoTo errlog
    
    'Call API Function to show properties or open file.
    Dim SEI As SHELLEXECUTEINFO
    With SEI
        .cbSize = Len(SEI)
        .fMask = 64 Or 12 Or 1024
        .hWnd = OwnerhWnd
        If props Then .lpVerb = "properties" Else .lpVerb = "open"
        .lpFile = filename
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        If props Then .nShow = 0 Else .nShow = 1
        .hInstApp = 0
        .lpIDList = 0
    End With
    ShellExecuteEX SEI
    ShowFileProperties = SEI.hInstApp
    
errlog:
   subErrLog ("Form2:ShowFileProperties")
End Function

Private Sub mCopyPath_Click()
On Error GoTo errlog
   Clipboard.Clear
   Clipboard.SetText fnGetPath
errlog:
   subErrLog ("Form2:mCopyPath_Click")
End Sub

Private Sub mDelShort_Click()
On Error GoTo errlog

   Dim DS As Integer, delcnt As Integer, ls As Integer, dint As Integer
   delcnt = ListView2.ListItems.Count
   dint = 0
   For DS = 1 To delcnt  'listitems.count goes down with each delete
'      If ListView2.SelectedItem.Index = DS Then
'      Debug.Print Str(DS)
         If ListView2.ListItems(DS).Selected = True Then
            Call subDelShort(DS, delcnt)
            ls = DS - dint    'dint is needed to count the last selected item - # deleted items
            dint = dint + 1
         End If
'      End If
      If DS = delcnt + 1 Then
         Exit For
      End If
   Next DS
   
   If ls = 0 Then Exit Sub
      
'   subFirstSel
   blDateSwitch = True
   PopFromMem
   subLbCount

   'this gets the selection rectangle on the listitem before the last selected one
   With ListView2
      If ls > 1 Then
         .ListItems(1).Selected = False
         .ListItems(ls).Selected = True
         .SelectedItem.Selected = False
      Else
         If .ListItems.Count > 0 And delcnt = 1 Then
            .SelectedItem.Selected = False
         End If
      End If
   End With
   
   'when last shortcut is removed for a filetype, the type is removed from form1
   If ListView2.ListItems.Count = 0 Then
      Dim delindx As Integer
      Dim pt As Integer
      With Form1
         For pt = 0 To UBound(PrevType)   'delete previous type if deleted type, so type can be readded (unlikely)
            If .ListView1.SelectedItem.Text = PrevType(pt) Then
               PrevType(pt) = vbNullString
               Exit For
            End If
         Next pt
         delindx = .ListView1.SelectedItem.Index
'         .ListView1.SelectedItem.Selected = False
         .ListView1.ListItems.Remove (.ListView1.SelectedItem.Text)
         .StatusBar1.Panels(1).Text = Str$(.ListView1.ListItems.Count) + " file types loaded."
         blExit = False
         If Form1.mSingleClick.Checked = False Then
            Unload Me
         Else
            If delindx > 1 Then
                .ListView1.ListItems(delindx - 1).Selected = True
            Else
                .ListView1.SelectedItem.Selected = True
            End If
            Form1.ListView1_Click
         End If
      End With
   End If

errlog:
   subErrLog ("Form2:mDelShort_Click")
End Sub

Private Sub subDelShort(DS As Integer, ByRef delcnt As Integer)
On Error GoTo errlog
      
   Dim SHFileOp As SHFILEOPSTRUCT
   Dim lFlags   As Long
   Dim lresult As Long
   
   lFlags = lFlags Or FOF_SILENT
   lFlags = lFlags Or FOF_NOCONFIRMATION
   
   'Delete in \Recent folder
   If DoesFileExist(sRecent + Trim$(ListView2.ListItems(DS).SubItems(iShortCol))) = True Then
      With SHFileOp
          .wFunc = FO_DELETE
          .pFrom = sRecent + Trim$(ListView2.ListItems(DS).SubItems(iShortCol)) & vbNullChar & vbNullChar
          .fFlags = lFlags
      End With
      lresult = SHFileOperation(SHFileOp)
      If lresult <> 0 Then
         GoTo delerr
      End If
'   Else
'      delcnt = delcnt - 1     'if shortcut doesn't exist, get rid of the item
   End If

   'Delete in \TaskTracker folder
   If DoesFileExist(TTpath + Trim$(ListView2.ListItems(DS).SubItems(iShortCol))) = True Then
      With SHFileOp
          .wFunc = FO_DELETE
          .pFrom = TTpath + Trim$(ListView2.ListItems(DS).SubItems(iShortCol)) & vbNullChar & vbNullChar
          .fFlags = lFlags
      End With
      lresult = SHFileOperation(SHFileOp)
      If lresult = 0 Then
         delcnt = delcnt - 1
      Else
         GoTo delerr
      End If
'   Else
'      delcnt = delcnt - 1     'if shortcut doesn't exist, get rid of the item
   End If
   
'  Remove from filename arrays
   Dim sh As Integer
   
   For sh = 0 To TTCnt

      If aShortcut(sh) = ListView2.ListItems(DS).SubItems(iShortCol) Then
         fType(sh) = vbNullString
         tType(sh) = vbNullString
         blExists(sh) = vbNull
'         aSmallIcon(sh) = vbNullString
'         aFType(sh) = vbNullString
         aLastAcc(sh) = vbNullString
         aLastMod(sh) = vbNullString
         aCreated(sh) = vbNullString
         aShortcut(sh) = vbNullString
         aSFileName(sh) = vbNullString
         Exit For
      End If
      
   Next sh
   
Exit Sub
delerr:
   MsgBox "The shortcut file (" + Trim$(ListView2.ListItems(DS).SubItems(iShortCol)) + ") " _
          + "could not be deleted.", vbInformation
errlog:
   subErrLog ("Form2:subDelShort")
End Sub

Private Sub mDelVirtual_Click()
   On Error GoTo errlog
   Dim dv As Integer
   For dv = 1 To ListView2.ListItems.Count
'     If dv <= ListView2.ListItems.Count Then
        If ListView2.ListItems(dv).Selected = True Then
           subDelVirtual (dv)
           dv = dv - 1
        End If
'     End If
   Next dv
errlog:
   subErrLog ("Form2:mDelVirtual_Click")
End Sub

Private Sub subDelVirtual(SelItm As Integer)
    On Error GoTo errlog
    Dim vtp As String
    Dim ash As String
    Dim df As Integer
'    Dim blFound As Boolean

    'Get the shortcut
    With ListView2
        If iShortCol = 2 Then
            ash = .ListItems(SelItm).SubItems(2)
        ElseIf iShortCol = 3 Then
            ash = .ListItems(SelItm).SubItems(3)
        ElseIf iShortCol = 4 Then
            ash = .ListItems(SelItm).SubItems(4)
        End If
        
        'case where the shortcut is missing because the "real" TT item has been deleted
        If Len(ash) = 0 Then
           Dim blNameOnly As Boolean
           Dim fNameOnly As String
           blNameOnly = True
           ash = .ListItems(SelItm)
        End If
        
    End With
    
   'and the virtual ttype
    With Form5.ListView3
       vtp = .ListItems(.SelectedItem.Index).Key
    End With
    
   'then use both to get the right array counter
    TTCnt = UBound(fType)
        
    For df = 0 To TTCnt
      If tType(df) = vtp Then
         If blNameOnly = True Then
            fNameOnly = Right$(aSFileName(df), Len(aSFileName(df)) - InStrRev(aSFileName(df), "\"))
         End If
         If blNameOnly = False And aShortcut(df) = ash Or _
            blNameOnly = True And fNameOnly = ash Then
            fType(df) = vbNullString
            tType(df) = vbNullString
            blExists(df) = vbNull
            aLastDate(df) = vbNullString
            aShortcut(df) = vbNullString
            aSFileName(df) = vbNullString
'            blFound = True
            Exit For
        End If
      End If
    Next df
    
'    If blFound = False Then
'      For df = 0 To TTCnt
'        If tType(df) = vtp Then
'           fType(df) = vbNullString
'           tType(df) = vbNullString
'           blExists(df) = vbNull
'           aLastDate(df) = vbNullString
'           aShortcut(df) = vbNullString
'           aSFileName(df) = vbNullString
'           blFound = True
'           Exit For
'        End If
'      Next df
'    End If

errlog:
   ListView2.ListItems.Remove (SelItm)
   If ListView2.ListItems.Count = 0 Then
      Form5.mDelete_Click
   End If
   subErrLog ("Form2:subDelVirtual")
End Sub

Private Sub mExplorer_Click()
   blExplorer = True
   blContainer = True
   OpenFiles
End Sub

Private Sub mFileTypes_Click()
On Error GoTo errlog
   If mFileTypes.Checked = False Then
      mFileTypes.Checked = True
      If blVirtualView = True Then
         SaveSetting "TaskTracker", "Settings", "VirtualViewTypes", True
         blVirtual = True
         PopFromMem
         blVirtual = False
      Else
         SaveSetting "TaskTracker", "Settings", "OtherViewTypes", True
         PopFromMem
      End If
      Exit Sub
   Else
      mFileTypes.Checked = False
      If blVirtualView = True Then
         SaveSetting "TaskTracker", "Settings", "VirtualViewTypes", False
         blVirtual = True
         PopFromMem
         blVirtual = False
      Else
         SaveSetting "TaskTracker", "Settings", "OtherViewTypes", False
         PopFromMem
      End If
   End If
   Form2.Refresh
errlog:
   subErrLog ("Form2:mFileTypes_Click")
End Sub

Public Sub mFilter_Click()
On Error GoTo errlog

  With Form4
    If mFilter.Checked = True Then
      mFilter.Checked = False
      .Form_Unload (1)
    Else
      mFilter.Checked = True
      blForm4Loading = True
      Form5.Form_Unload (1)
'      Form2.subForm2Init
      If .blForm4Loaded = False Then
         .SetPosition
      End If
      If blChangeComboHeight = False Then
         .ChangeComboHeight
         blChangeComboHeight = True
      End If
      .Show
'      .blForm4Showing = True
      .cbDates.SetFocus    'close dropdown
      .cbSearch.SetFocus
'      .cbSearch_Change
      .cmdVirtual.Enabled = False
'      .cbSearch_click
      If Right$(Form2.Caption, Len(Form2.Caption) - 14) <> sType Then   'in case a virtual folder
         Form1.ListView1_Click
      End If
      Form4.SetFocus
    End If
  End With
  
errlog:
subErrLog ("Form2:mFilter_click")
End Sub

Public Sub FolderSort()
On Error GoTo errlog
Dim ch As Integer
   
   With ListView2
   
     ch = .ColumnHeaders.Count
      
      If blAllTypes = True Or blAddType = True Then
         If blFolderSort Then
            .SortOrder = MultSortOrder
            If MultSortCol > ch Then .SortKey = MultSortCol
         Else
            .SortOrder = MultPSortOrder
            If MultPSortCol > ch Then .SortKey = MultPSortCol
         End If
      ElseIf blThisType = True Then
         If sType = "Deleted or Renamed" Or sType = "Network Drive" _
            Or sType = "Removable Drive" Or sType = "CD Drive" Or blMultSelect = True Then
            If blFolderSort Then
               .SortOrder = DelPSortOrder
               If DelPSortCol > ch Then .SortKey = DelPSortCol
            Else
               .SortOrder = DelSortOrder
               If DelSortCol > ch Then .SortKey = DelSortCol
            End If
         Else
            If blFolderSort Then
               .SortOrder = RegPSortOrder
               If RegPSortCol > ch Then .SortKey = RegPSortCol
            Else
               .SortOrder = RegSortOrder
               If RegSortCol > ch Then .SortKey = RegSortCol
            End If
         End If
      End If
   
   End With
   
   blDateSwitch = True
         
errlog:
subErrLog ("Form2:mFolders_sort")
End Sub

Private Sub mExt_Click()
On Error GoTo errlog
   If mExt.Checked = True Then
      mExt.Checked = False
   Else
      mExt.Checked = True
   End If
   subForm2Init
errlog:
subErrLog ("Form2:mExt_Click")
End Sub

Private Sub Listview2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo errlog
   
   SaveSelected
   If Button = 2 Then
      subRC2Menu
   Else
      blRightClick = False
      If blLoading = False Then
         If Form1.mAutoPreview.Checked = True Or mPreview.Checked = True Then
            blForm6Load = True
            Form6.Timer1.Enabled = True
         End If
      End If
   End If
   
errlog:
subErrLog ("Form2:ListView2_MouseUp")
End Sub

Private Sub subRC2Menu()
On Error GoTo errlog

   Dim hDrop As Long
   Dim nS As Integer
   Dim ts As Integer
   Dim ms As Integer

   ListView2_Click
   
   nS = nSelected

   'Enable if clipboard has a file
   hDrop = IsClipboardFormatAvailable(CF_HDROP)
   If hDrop > 0 Then
       mPaste.Enabled = True
   Else
       mPaste.Enabled = False
   End If
   
   If ListView2.ListItems.Count > 0 And nS > 0 Then
   
      If blThisType = True Then
   
         ThisTypeCase
         
      Else
         
   '      mPaste.Visible = False
         mContainer.Visible = True
         If ListView2.ListItems(nSelected).SubItems(1) = "Deleted or Renamed" Then
            mOpen(1).Enabled = False
            mContainer.Enabled = False
         Else
            mOpen(1).Enabled = True
            mContainer.Enabled = True
         End If
         
      End If
      
   Else
         mOpen(1).Enabled = False
         mOpenWith.Enabled = False
         mContainer.Enabled = False
         mExplorer.Visible = False
         mCopy.Enabled = False
         mCopyPath.Enabled = False
         mPaste.Enabled = False
         mPaste.Caption = "&Paste to Containing Folder"
         mExt.Enabled = True
         mProps.Enabled = False
         mShowExpMenu.Enabled = False
         mRename.Enabled = False
         mDelShort.Enabled = False
   End If
   
   If AllorAddorSpecial = True Then
       mFileTypes.Visible = True
   Else
       mFileTypes.Visible = False
   End If

   If blVirtualView = True Then
      mDelShort.Enabled = False
      mAddVirtual.Visible = False
      mDelVirtual.Visible = True
   Else
      If Form1.mVirtual.Checked = True Then
         mAddVirtual.Visible = True
      Else
         mAddVirtual.Visible = False
      End If
      mDelVirtual.Visible = False
   End If

   For ts = 1 To ListView2.ListItems.Count
      If ListView2.ListItems(ts).Selected = True Then
         ms = ms + 1
      End If
   Next ts
   
   If ms > 1 Then
      mDelShort.Caption = "&Delete Shortcuts"
   Else
      mDelShort.Caption = "&Delete Shortcut"
   End If
   
errlog:
   PopupMenu mnOptions, , , , mOpen(1)
   blRightClick = True
   
   subErrLog ("Form2:subRC2Menu")
End Sub

Private Sub ThisTypeCase()
On Error GoTo errlog

   Dim sPath As String
   Dim TestType As String
   
   If VerifyFolderType = True Then
      TestType = "File Folder"
   ElseIf blVirtualView = True Then
      TestType = "Virtual"
   Else
      TestType = sType
   End If
   
   If blRegStatus = True Then
      If DoesFileExist(App.path + "\expmnu.dll") Then
         mShowExpMenu.Visible = True
      Else
         mShowExpMenu.Visible = False
      End If
   Else
      mShowExpMenu.Visible = False
   End If
   
   'set available commands on context menu
   Select Case TestType
   
      Case "Network Drive", "Removable Drive", "CD Drive", "Virtual"
         mOpen(1).Enabled = True
         mOpenWith.Enabled = True
         mContainer.Enabled = True
         mExplorer.Visible = False
         mCopy.Enabled = True
         mCopyPath.Enabled = True
         mPaste.Caption = "&Paste to Containing Folder"
         mExt.Enabled = True
         mProps.Enabled = True
         mShowExpMenu.Enabled = True
         mRename.Caption = "Re&name File/Folder"
         mRename.Enabled = True
         mDelShort.Enabled = True
         If CheckPreview = True Then
            mPreview.Visible = True
         End If
         
      Case "File Folder"
         mOpen(1).Enabled = True
         mContainer.Enabled = True
         mOpenWith.Enabled = False
         mCopy.Enabled = True
         mCopyPath.Enabled = True
         mExplorer.Visible = True
         mPaste.Caption = "&Paste"
         mExt.Enabled = False
         mProps.Enabled = True
         mShowExpMenu.Enabled = True
         mRename.Caption = "Re&name Folder"
         mRename.Enabled = True
         mDelShort.Enabled = True
        
      Case "Deleted or Renamed", "Unknown"
         mOpen(1).Enabled = False
         mOpenWith.Enabled = False
         mCopy.Enabled = False
         mCopyPath.Enabled = True
         mExplorer.Visible = False
         mPaste.Enabled = False
         mExt.Enabled = True
         mProps.Enabled = False
         mShowExpMenu.Enabled = False
         mRename.Enabled = False
         mDelShort.Enabled = True
         mViewBin.Visible = True
         mPaste.Caption = "&Paste to Same Folder"
         'If the folder still exists, enable "open containing folder" command
         sPath = fnGetPath
         sPath = Left$(sPath, Len(sPath) - (Len(sPath) - InStrRev(sPath, "\")))
         If DoesFolderExist(sPath) Then
            mContainer.Enabled = True
         Else
            mContainer.Enabled = False
         End If
      
      Case Else   'all other types
         mOpen(1).Enabled = True
         mOpenWith.Enabled = True
         mContainer.Enabled = True
         mExplorer.Visible = True
         mCopy.Enabled = True
         mCopyPath.Enabled = True
         mPaste.Caption = "&Paste to Containing Folder"
         mExt.Enabled = True
         mProps.Enabled = True
         mShowExpMenu.Enabled = True
         mRename.Caption = "Re&name File"
         mRename.Enabled = True
         mDelShort.Enabled = True
         mViewBin.Visible = False
         If CheckPreview = True Then
            mPreview.Visible = True
         End If

   End Select

errlog:
   subErrLog ("Form2:ThisTypeCase")
End Sub

Public Function CheckPreview() As Boolean
On Error GoTo errlog
   Dim fp As String
   Dim sExt As String
'   Dim blPreview As Boolean
   fp = fnGetPath
   sExt = LCase$(Right$(fp, Len(fp) - InStrRev(fp, ".")))

   mPreview.Visible = False
   Select Case sExt
   
      Case "bmp", "gif", "jpg", "jpeg", "jpe", "jfif", "tif", "tiff", "png", "ico", "cur"
         CheckPreview = True
               
      Case Else
         CheckPreview = False
         
         If mPreview.Checked = True Then
            mPreview.Checked = False
         End If
         
   End Select
errlog:
   subErrLog ("Form2:CheckPreview")
End Function

Private Sub mFullPath_Click()
On Error GoTo errlog
   If mFullPath.Checked = False Then
      mFullPath.Checked = True
      blFolderSort = True
      mContainerFolder.Checked = False
   Else
      mFullPath.Checked = False
      blFolderSort = False
   End If
   blDontSaveCols = True
   FolderSort
   subForm2Init
errlog:
   subErrLog ("Form2:mFullPath_Click")
End Sub

Private Sub mOpen_Click(Index As Integer)
On Error GoTo errlog
   blExplorer = False
   OpenFiles
errlog:
   subErrLog ("Form2:mOpen_Click")
End Sub

Private Sub mContainer_click()
On Error GoTo errlog
   blContainer = True
   btOpen_Click
errlog:
   subErrLog ("Form2:mContainer_click")
End Sub

Private Sub mOpenWith_Click()
On Error GoTo errlog
   blOpenWith = True
   OpenFiles
   blOpenWith = False
errlog:
   subErrLog ("Form2:mOpenWith_Click")
End Sub

Private Sub OpenWithDialog(sfile As String)
On Error GoTo errlog
Dim lRet As Long
Dim sDir As String

    sDir = Space(260)
    lRet = GetSystemDirectory(sDir, Len(sDir))

    sDir = Left$(sDir, lRet)

    Call ShellExecute(GetDesktopWindow, _
          vbNullString, "RUNDLL32.EXE", _
          "shell32.dll,OpenAs_RunDLL " & _
          sfile, sDir, vbNormalFocus)

errlog:
   subErrLog ("Form2:OpenWithDialog")
End Sub

Private Sub mModified_Click()
On Error GoTo errlog

   mAccessed.Checked = False
   mCreated.Checked = False
   mModified.Checked = True
   blDateSwitch = True
'   subFirstSel
   Erase blTrueDates
   subForm2Init
   
errlog:
   subErrLog ("Form2:mModified_Click")
End Sub

Private Sub mCreated_Click()
On Error GoTo errlog

   mModified.Checked = False
   mAccessed.Checked = False
   mCreated.Checked = True
   blDateSwitch = True
'   subFirstSel
   Erase blTrueDates
   subForm2Init

errlog:
   subErrLog ("Form2:mCreated_Click")
End Sub

Private Sub mPaste_Click()
On Error GoTo errlog

   If VerifyFolderType = False Then
      blContainer = True
   End If
   blPaste = True
   OpenFiles
   If VerifyFolderType = False Then
      blContainer = True
   End If
   OpenFiles                 'twice ensures the explorer window gets focus before sending keys
   SendKeys "+{INSERT}"        ' Paste
   
errlog:
   subErrLog ("Form2:mPaste_Click")
End Sub

Private Sub mPreferences_Click()
On Error GoTo errlog

   blMenu = True
   Form1.Timer3.Enabled = True

errlog:
   subErrLog ("Form2:mPreferences_Click")
End Sub

Private Sub mPreview_Click()
On Error GoTo errlog

   If mPreview.Checked = False Then
      If CheckPreview = False Then
         Exit Sub
      End If
      mPreview.Checked = True
      blForm6Load = True
      Form6.Timer1.Enabled = True
   Else
      mPreview.Checked = False
      Set Form6.picPreview.Picture = LoadPicture()
      Form6.Hide
      blPreviewOpen = False
   End If
   
errlog:
   subErrLog ("Form2:mPreview_click")
End Sub

Private Sub mProps_Click()
On Error GoTo errlog
   Call ShowFileProperties(fnGetPath, 0, True)
errlog:
subErrLog ("Form2:mProps_Click")
End Sub

Private Sub mRename_Click()
On Error GoTo errlog
   ListView2.StartLabelEdit
errlog:
subErrLog ("Form2:mRename_Click")
End Sub

Public Sub subSetSize()
On Error GoTo errlog
   
   If Form1.mSaveSize.Checked = False Then 'if no settings
     If Form1.mIcons.Checked = False Then
         If Form2.Left + (Form2.Width * 2) < Form1.Left + (Form1.Width * 2) Then 'Form2 is left of Form1
            If Form1.Left > Form2.Width Then
               Form2.Left = Form1.Left - Form2.Width  'form2 on left if space
            Else
               Form2.Left = Form1.Left + Form1.Width  'else form2 on right
            End If
         ElseIf Form2.Left + (Form2.Width * 2) > Form1.Left + (Form1.Width * 2) Then   'Form1 is left of Form2
            If wWidth - (Form1.Left + Form1.Width) > Form2.Width Then
               Form2.Left = Form1.Left + Form1.Width  'form2 on right if space
            Else
               Form2.Left = Form1.Left - Form2.Width  'else form2 on left
            End If
         End If
      Else
         Form2.Left = Form1.Left - Form2.Width
      End If
      Form2.Height = Form1.Height
      Form2.Top = Form1.Top
   Else
      If blFirstLoad = True Then   'use settings on load
         If Len(GetSetting("TaskTracker", "Settings", "Form2Top")) > 0 Then
            Form2.Top = Val(GetSetting("TaskTracker", "Settings", "Form2Top"))
         End If
         If Len(GetSetting("TaskTracker", "Settings", "Form2Height")) > 0 Then
            Form2.Height = Val(GetSetting("TaskTracker", "Settings", "Form2Height"))
         End If
         If Len(GetSetting("TaskTracker", "Settings", "Form2Width")) > 0 Then
            Form2.Width = Val(GetSetting("TaskTracker", "Settings", "Form2Width"))
         End If
         If Len(GetSetting("TaskTracker", "Settings", "Form2Left")) > 0 Then
            Form2.Left = Val(GetSetting("TaskTracker", "Settings", "Form2Left"))
         End If
      End If
   End If
   
errlog:
blFirstLoad = False
subErrLog ("Form2:subSetSize")
End Sub

Public Sub GetListView2Settings()
   If Len(GetSetting("TaskTracker", "Settings", "col1width")) > 0 Then
      col1width = Val(GetSetting("TaskTracker", "Settings", "col1width"))
   Else
      col1width = 3850
   End If
   If Len(GetSetting("TaskTracker", "Settings", "col2width")) > 0 Then
      col2width = Val(GetSetting("TaskTracker", "Settings", "col2width"))
   Else
      col2width = 2500
   End If
'   If Len(GetSetting("TaskTracker", "Settings", "colSwidth")) > 0 Then
'      colSwidth = Val(GetSetting("TaskTracker", "Settings", "colSwidth"))
'   Else
'      colSwidth = 0
'   End If
   If Len(GetSetting("TaskTracker", "Settings", "col1Pwidth")) > 0 Then
      col1Pwidth = Val(GetSetting("TaskTracker", "Settings", "col1Pwidth"))
   Else
      col1Pwidth = 2000
   End If
   If Len(GetSetting("TaskTracker", "Settings", "col2Pwidth")) > 0 Then
      col2Pwidth = Val(GetSetting("TaskTracker", "Settings", "col2Pwidth"))
   Else
      col2Pwidth = 2300
   End If
   If Len(GetSetting("TaskTracker", "Settings", "col3Pwidth")) > 0 Then
      col3Pwidth = Val(GetSetting("TaskTracker", "Settings", "col3Pwidth"))
   Else
      col3Pwidth = 2050
   End If
   If Len(GetSetting("TaskTracker", "Settings", "col4Swidth")) > 0 Then
      col4Swidth = Val(GetSetting("TaskTracker", "Settings", "col4Swidth"))
   Else
      col4Swidth = 0
   End If
   If Len(GetSetting("TaskTracker", "Settings", "col1Awidth")) > 0 Then
      col1Awidth = Val(GetSetting("TaskTracker", "Settings", "col1Awidth"))
   Else
      col1Awidth = 2300
   End If
   If Len(GetSetting("TaskTracker", "Settings", "col1NAwidth")) > 0 Then
      col1NAwidth = Val(GetSetting("TaskTracker", "Settings", "col1NAwidth"))
   Else
      col1NAwidth = 3850
   End If
   If Len(GetSetting("TaskTracker", "Settings", "col2Awidth")) > 0 Then
      col2Awidth = Val(GetSetting("TaskTracker", "Settings", "col2Awidth"))
   Else
      col2Awidth = 1900
   End If
'   If Len(GetSetting("TaskTracker", "Settings", "col2NAwidth")) > 0 Then
'      col2NAwidth = Val(GetSetting("TaskTracker", "Settings", "col2NAwidth"))
'   Else
'      col2NAwidth = 2500
'   End If
   If Len(GetSetting("TaskTracker", "Settings", "col3Awidth")) > 0 Then
      col3Awidth = Val(GetSetting("TaskTracker", "Settings", "col3Awidth"))
   Else
      col3Awidth = 1900
   End If
   If Len(GetSetting("TaskTracker", "Settings", "col3NAwidth")) > 0 Then
      col3NAwidth = Val(GetSetting("TaskTracker", "Settings", "col3NAwidth"))
   Else
      col3NAwidth = 2500
   End If
   If Len(GetSetting("TaskTracker", "Settings", "col1PAwidth")) > 0 Then
      col1PAwidth = Val(GetSetting("TaskTracker", "Settings", "col1PAwidth"))
   Else
      col1PAwidth = 1890
   End If
   If Len(GetSetting("TaskTracker", "Settings", "col1NPAwidth")) > 0 Then
      col1NPAwidth = Val(GetSetting("TaskTracker", "Settings", "col1NPAwidth"))
   Else
      col1NPAwidth = 2000
   End If
   If Len(GetSetting("TaskTracker", "Settings", "col2PAwidth")) > 0 Then
      col2PAwidth = Val(GetSetting("TaskTracker", "Settings", "col2PAwidth"))
   Else
      col2PAwidth = 1300
   End If
   If Len(GetSetting("TaskTracker", "Settings", "col2NPAwidth")) > 0 Then
      col2NPAwidth = Val(GetSetting("TaskTracker", "Settings", "col2NPAwidth"))
   Else
      col2NPAwidth = 2000
   End If
   If Len(GetSetting("TaskTracker", "Settings", "col3PAwidth")) > 0 Then
      col3PAwidth = Val(GetSetting("TaskTracker", "Settings", "col3PAwidth"))
   Else
      col3PAwidth = 1300
   End If
'   If Len(GetSetting("TaskTracker", "Settings", "col3NPAwidth")) > 0 Then
'      col3NPAwidth = Val(GetSetting("TaskTracker", "Settings", "col3NPAwidth"))
'   Else
'      col3NPAwidth = 2500
'   End If
   If Len(GetSetting("TaskTracker", "Settings", "col4PAwidth")) > 0 Then
      col4PAwidth = Val(GetSetting("TaskTracker", "Settings", "col4PAwidth"))
   Else
      col4PAwidth = 1300
   End If
   If Len(GetSetting("TaskTracker", "Settings", "col4NPAwidth")) > 0 Then
      col4NPAwidth = Val(GetSetting("TaskTracker", "Settings", "col4NPAwidth"))
   Else
      col4NPAwidth = 2600
   End If
End Sub

Private Sub mShowExpMenu_Click()
On Error GoTo errlog
    Dim pt As POINTAPI
    Dim DL As Long
    Dim sl As Long
    Dim sFilePath As String
        
    DL = GetCursorPos(pt)
    
    For sl = 1 To ListView2.ListItems.Count
      If ListView2.ListItems(sl).Selected = True Then
         If Len(ListView2.ListItems(sl).Text) > 0 Then
            sFilePath = fnSelectedItemsPath(sl)
         End If
      End If
    Next
'    If DoesFolderExist(sFilePath) Then
'      If Right$(sFilePath, 1) <> "\" Then sFilePath = sFilePath & "\"
'    End If
    sFilePath = sFilePath & Chr$(0)
    
    DL = DoExplorerMenu((Me.hWnd), sFilePath, pt.x - 300, pt.y - 300)
errlog:
   subErrLog ("Form2:mShowExpMenu_Click")
End Sub

Private Sub mSimple_Click()
   If mSimple.Checked = False Then
      mSimple.Checked = True
   Else
      mSimple.Checked = False
   End If
   subSimple
End Sub

Public Sub subSimple()
   If mSimple.Checked = True Then
      mFolders.Visible = False
      Form1.mGridName.Enabled = False
      ListView2.View = lvwList
   Else
      mSimple.Checked = False
      mFolders.Visible = True
      Form1.mGridName.Enabled = True
      ListView2.View = lvwReport
   End If
End Sub

'Public Sub subSelTypes()
'On Error GoTo errlog
'   SelTypes =vbnullstring
'   TTS.Destroy
'   If Form1.mHideTips.Checked = True Then Exit Sub
'   If blThisType = True Or blAllTypes = True Then
'      Exit Sub
'   End If
'
'   Dim ts As Integer
'   With Form1.ListView1
'      For ts = 1 To .ListItems.Count
'         If .ListItems(ts).Selected = True Then
'            SelTypes = SelTypes + .ListItems(ts).Text + ", "
'         End If
'      Next ts
'      SelTypes = Left$(SelTypes, Len(SelTypes) - 2)
'   End With
'   With TTS
''      .HwndParentControl = Form4.ckAll.hwnd
'      .Title = "Selected file types"
'      .TipText = SelTypes
'      .Create Form4.ckAll.hwnd
'      .MaxTipWidth = IIf(1, 225, -1)
'      .SetDelayTime (sdtInitial), 0
'    End With
'errlog:
'   subErrLog ("Form2:subSelTypes")
'End Sub

Private Sub GetSettings()
On Error GoTo errlog
      
   If LenB(GetSetting("TaskTracker", "Settings", "FileExt")) > 0 Then
      Form2.mExt.Checked = CBool(GetSetting("TaskTracker", "Settings", "FileExt"))
   Else
      Form2.mExt.Checked = True 'default
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "LastDate")) > 0 Then
      If GetSetting("TaskTracker", "Settings", "LastDate") = "Created" Then
         mAccessed.Checked = False
         mModified.Checked = False
         mCreated.Checked = True
      ElseIf GetSetting("TaskTracker", "Settings", "LastDate") = "Modified" Then
         mAccessed.Checked = False
         mModified.Checked = True
         mCreated.Checked = False
      ElseIf GetSetting("TaskTracker", "Settings", "LastDate") = "Accessed" Then
         mAccessed.Checked = True
         mModified.Checked = False
         mCreated.Checked = False
      End If
   Else  'default
      mAccessed.Checked = False
      mModified.Checked = True
      mCreated.Checked = False
   End If
   
   'setting in form1
   If LenB(GetSetting("TaskTracker", "Settings", "Grid2")) > 0 Then
      Form1.mGridName.Checked = CBool(GetSetting("TaskTracker", "Settings", "Grid2"))
   Else
      Form1.mGridName.Checked = True   'default
   End If
   Call SendMessage(ListView2.hWnd, _
                    LVM_SETEXTENDEDLISTVIEWSTYLE, _
                    LVS_EX_GRIDLINES, ByVal Form1.mGridName.Checked)

   If LenB(GetSetting("TaskTracker", "Settings", "Explorer")) > 0 Then
      blExplorer = CBool(GetSetting("TaskTracker", "Settings", "Explorer"))
   Else
      blExplorer = False 'default
   End If
        
   If LenB(GetSetting("TaskTracker", "Settings", "Simple")) > 0 Then
      mSimple.Checked = CBool(GetSetting("TaskTracker", "Settings", "Simple"))
      subSimple
   Else
      mSimple.Checked = False 'default
   End If
   
   If LenB(GetSetting("TaskTracker", "Settings", "FullPath")) > 0 Then
      mFullPath.Checked = CBool(GetSetting("TaskTracker", "Settings", "FullPath"))
   Else
      mFullPath.Checked = False 'default
   End If
      
   If LenB(GetSetting("TaskTracker", "Settings", "ContainerFolder")) > 0 Then
      mContainerFolder.Checked = CBool(GetSetting("TaskTracker", "Settings", "ContainerFolder"))
   Else
      mContainerFolder.Checked = False 'default
   End If
      
   LoadColPos

errlog:
   subErrLog ("Form2:GetSettings")
End Sub

Private Sub LoadColPos()
On Error GoTo errlog

   Dim Ord As Integer

   If LenB(GetSetting("TaskTracker", "Settings", "ColOrderAdd")) > 0 And Form1.mSaveSize.Checked = True Then
      ColOrderAdd = GetSetting("TaskTracker", "Settings", "ColOrderAdd")
   Else
      ColOrderAdd = "012345"  'default
   End If
   
   ReDim Preserve PosArrayAddTypes(5)
   ReDim Preserve PrevArrayAddTypes(5)
   For Ord = 0 To 5
      PosArrayAddTypes(Ord) = Mid$(ColOrderAdd, Ord + 1, 1)
      PrevArrayAddTypes(Ord) = Mid$(ColOrderAdd, Ord + 1, 1)
'      Debug.Print Mid$(ColOrderAdd, Ord + 1, 1)
   Next Ord
  
   If LenB(GetSetting("TaskTracker", "Settings", "ColOrderSpec")) > 0 And Form1.mSaveSize.Checked = True Then
      ColOrderSpec = GetSetting("TaskTracker", "Settings", "ColOrderSpec")
   Else
      ColOrderSpec = "012345"  'default
   End If
   
   ReDim Preserve PosArraySpecTypes(5)
   ReDim Preserve PrevArraySpecTypes(5)
   For Ord = 0 To 5
      PosArraySpecTypes(Ord) = Mid$(ColOrderSpec, Ord + 1, 1)
      PrevArraySpecTypes(Ord) = Mid$(ColOrderSpec, Ord + 1, 1)
'      Debug.Print Mid$(ColOrderSpec, Ord + 1, 1)
   Next Ord
   
   If LenB(GetSetting("TaskTracker", "Settings", "ColOrderReg")) > 0 And Form1.mSaveSize.Checked = True Then
      ColOrderReg = GetSetting("TaskTracker", "Settings", "ColOrderReg")
   Else
      ColOrderReg = "012345"  'default
   End If
   
   ReDim Preserve PosArrayRegTypes(5)
   ReDim Preserve PrevArrayRegTypes(5)
   For Ord = 0 To 5
      PosArrayRegTypes(Ord) = Mid$(ColOrderReg, Ord + 1, 1)
      PrevArrayRegTypes(Ord) = Mid$(ColOrderReg, Ord + 1, 1)
'      Debug.Print Mid$(ColOrderReg, Ord + 1, 1)
   Next Ord

errlog:
   subErrLog ("Form2:LoadColPos")
End Sub


'Private Sub LVW_BINDHEADERICONS(lvwCtrl As ListView, clsHeader)
'   Set clsHeader = New cHeadIcons
'   Set clsHeader.ListView = lvwCtrl
'   Call clsHeader.SetHeaderIcons(0, Ascn)
'End Sub

'Private Function ObscuredText(iItem As Long, iSubItem As Long, slpz As Long) As Boolean
'   Dim cxText As Long
'   Dim rcLV As RECT
'   Dim cxCol As Long
'   Dim rcItem As RECT
'   Dim fRet As Boolean
'
'   ' Get the specified item's text width and item rect, and the ListView's client rect.
'   cxText = SendMessage(ListView2.hwnd, LVM_GETSTRINGWIDTHA, 0, slpz)
'   Call ListView_GetSubItemRect(ListView2.hwnd, iItem, iSubItem, LVIR_LABEL, rcItem)
'   Call GetClientRect(ListView2.hwnd, rcLV)
'
''   If iSubItem = 0 Then
''      cxCol = SendMessage(ListView2.hwnd, LVM_GETCOLUMNWIDTH, 0, 0)
''      fRet = ((cxText + 4) > (rcItem.Right - rcItem.Left))
''      Call InflateRect(rcItem, -2, -2)
''      ObscuredText = fRet Or (RectInRect(rcLV, rcItem) = False)
''   Else
'      cxCol = SendMessage(ListView2.hwnd, LVM_GETCOLUMNWIDTH, ByVal iSubItem, 0)
'      Call InflateRect(rcItem, -6, -2)
'      ObscuredText = ((cxText + 12) > cxCol) Or (RectInRect(rcLV, rcItem) = False)
''   End If
'End Function
'
'Private Function ListView_GetSubItemRect(hwnd As Long, iItem As Long, iSubItem As Long, _
'                                                                    code As Long, prc As RECT) As Boolean
'  prc.Top = iSubItem
'  prc.Left = code
'  ListView_GetSubItemRect = SendMessage(hwnd, LVM_GETSUBITEMRECT, ByVal iItem, prc)
'End Function
'
'Private Function RectInRect(rc1 As RECT, rc2 As RECT) As Boolean
'  RectInRect = PtInRect(rc1, rc2.Left, rc2.Top) And PtInRect(rc1, rc2.Right, rc2.Bottom)
'End Function
'

Public Sub GetColPos()
On Error GoTo errlog

   Dim totalCols As Long
   Dim pos As Integer
   Dim blDiff As Boolean
   
   totalCols = ListView2.ColumnHeaders.Count
   
   If blAllTypes = True Or blAddType = True Then
   
      If blVirtual = False Then
         If blColAdd = True Then
            blColSame = True
         Else
            blColSame = False
         End If
         blColAdd = True
         blColReg = False
         blColSpec = False
   
         totalCols = ListView2.ColumnHeaders.Count
         
         Call SendMessage(ByVal ListView2.hWnd, _
                          ByVal LVM_GETCOLUMNORDERARRAY, _
                          ByVal totalCols, _
                          PosArrayAddTypes(0))
                               
         blDiff = False
         If blColSame = True Then
            For pos = 0 To UBound(PosArrayAddTypes)
               If PosArrayAddTypes(pos) <> PrevArrayAddTypes(pos) Then
                  blDiff = True
                  Exit For
               End If
            Next pos
            If blDiff = True Then
               For pos = 0 To UBound(PosArrayAddTypes)
                 PrevArrayAddTypes(pos) = PosArrayAddTypes(pos)
               Next pos
            Else
               For pos = 0 To UBound(PrevArrayAddTypes)
                 PosArrayAddTypes(pos) = PrevArrayAddTypes(pos)
               Next pos
            End If
         Else
            For pos = 0 To UBound(PrevArrayAddTypes)
              PosArrayAddTypes(pos) = PrevArrayAddTypes(pos)
            Next pos
         End If
         Exit Sub
      End If
   End If
  
   If blThisType = True And blExtraCol Then
      
      If blVirtual = False Then
         If blColSpec = True Then
            blColSame = True
         Else
            blColSame = False
         End If
         blColSpec = True
         blColAdd = False
         blColReg = False
               
         Call SendMessage(ByVal ListView2.hWnd, _
                          ByVal LVM_GETCOLUMNORDERARRAY, _
                          ByVal totalCols, _
                          PosArraySpecTypes(0))
           
         If blColSame = True Then
            blDiff = False
            For pos = 0 To UBound(PosArraySpecTypes)
               If PosArraySpecTypes(pos) <> PrevArraySpecTypes(pos) Then
                  blDiff = True
                  Exit For
               End If
            Next pos
            If blDiff = True Then
               For pos = 0 To UBound(PosArraySpecTypes)
                 PrevArraySpecTypes(pos) = PosArraySpecTypes(pos)
               Next pos
            Else
               For pos = 0 To UBound(PrevArraySpecTypes)
                 PosArraySpecTypes(pos) = PrevArraySpecTypes(pos)
               Next pos
            End If
         Else
            For pos = 0 To UBound(PrevArraySpecTypes)
               PosArraySpecTypes(pos) = PrevArraySpecTypes(pos)
            Next pos
         End If
         Exit Sub
      End If
   End If
   
   'Regular
      
   If blColReg = True Then
      blColSame = True
   Else
      blColSame = False
   End If
   blColReg = True
   blColAdd = False
   blColSpec = False
        
   Call SendMessage(ByVal ListView2.hWnd, _
                    ByVal LVM_GETCOLUMNORDERARRAY, _
                    ByVal totalCols, _
                    PosArrayRegTypes(0))
                                           
   If blColSame = True Then
      blDiff = False
      For pos = 0 To UBound(PosArrayRegTypes)
         If PosArrayRegTypes(pos) <> PrevArrayRegTypes(pos) Then
            blDiff = True
            Exit For
         End If
      Next pos
      If blDiff = True Then
         For pos = 0 To UBound(PosArrayRegTypes)
           PrevArrayRegTypes(pos) = PosArrayRegTypes(pos)
         Next pos
      Else
         For pos = 0 To UBound(PrevArrayRegTypes)
           PosArrayRegTypes(pos) = PrevArrayRegTypes(pos)
         Next pos
      End If
   Else
      For pos = 0 To UBound(PrevArrayRegTypes)
         PosArrayRegTypes(pos) = PrevArrayRegTypes(pos)
      Next pos
   End If
  
errlog:
   subErrLog ("Form2:GetColPos")
End Sub

Public Sub SetColPos()
On Error GoTo errlog
 
  Dim totalCols As Long
  totalCols = ListView2.ColumnHeaders.Count
                            
  If blAllTypes = True Or blAddType = True Then
  
    Call SendMessage(ByVal ListView2.hWnd, _
                       ByVal LVM_SETCOLUMNORDERARRAY, _
                       ByVal totalCols, PosArrayAddTypes(0))
      
  ElseIf blThisType = True Then
  
    If blExtraCol Then
     
       Call SendMessage(ByVal ListView2.hWnd, _
                         ByVal LVM_SETCOLUMNORDERARRAY, _
                         ByVal totalCols, PosArraySpecTypes(0))
    Else
       GoTo Regular
    End If
    
  Else
  
Regular:

     Call SendMessage(ByVal ListView2.hWnd, _
                      ByVal LVM_SETCOLUMNORDERARRAY, _
                      ByVal totalCols, PosArrayRegTypes(0))
                      
'      Dim strcolorder As String
'      Dim co As Integer
'      For co = 0 To UBound(PosArrayRegTypes)
'         strcolorder = strcolorder + Trim$(Str(PosArrayRegTypes(co)))
'      Next co
'      Debug.Print strcolorder
  End If
  

errlog:
   subErrLog ("Form2:SetColPos")
End Sub

Public Sub Form2ColumnOrder()
On Error GoTo errlog

   GetColPos
   
   Dim co As Integer
   Dim strcolorder As String
   
   For co = 0 To UBound(PosArrayAddTypes)
      strcolorder = strcolorder + Trim$(Str(PosArrayAddTypes(co)))
   Next co
   SaveSetting "TaskTracker", "Settings", "ColOrderAdd", strcolorder
   
   strcolorder = vbNullString
   For co = 0 To UBound(PosArrayRegTypes)
      strcolorder = strcolorder + Trim$(Str(PosArrayRegTypes(co)))
   Next co
   SaveSetting "TaskTracker", "Settings", "ColOrderReg", strcolorder
   
   strcolorder = vbNullString
   For co = 0 To UBound(PosArraySpecTypes)
      strcolorder = strcolorder + Trim$(Str(PosArraySpecTypes(co)))
   Next co
   SaveSetting "TaskTracker", "Settings", "ColOrderSpec", strcolorder
   
errlog:
   subErrLog ("Form2:Form2ColumnOrder")
End Sub

Private Sub mViewBin_Click()
On Error GoTo errlog

   Dim skey As String
'   Clipboard.Clear
'   Clipboard.SetText ListView2.ListItems(ListView2.SelectedItem.Index).Text
   ShellExecute 0, "Open", "explorer.exe", _
        "/root,::{645FF040-5081-101B-9F08-00AA002F954E}", 0&, vbNormalFocus
   Sleep (2000)
   skey = Left$(ListView2.ListItems(ListView2.SelectedItem.Index).Text, 1)
   SendKeys skey
'   SendKeys "{F3}"
'   DoEvents
'   SendKeys "+{INSERT}"
'   SendKeys "+{INSERT}"
'   SendKeys "{ENTER}"
errlog:
   subErrLog ("Form2:mViewBin_Click")
End Sub

Private Sub mSearchFile_Click()
On Error GoTo errlog
   
   Dim fNameOnly As String
   Dim fpathonly As String
   fNameOnly = Right$(fnGetPath, Len(fnGetPath) - InStrRev(fnGetPath, "\"))
   fpathonly = Left$(fnGetPath, Len(fnGetPath) - Len(fNameOnly))
'   ShellExecute 0, "Find", "explorer.exe", "/" + fPathOnly, 0&, vbNormalFocus
   ShellExecute hWnd, "find", fpathonly, vbNullString, vbNullString, vbNormalFocus
'   Sleep (2000)
'   SendKeys "{ENTER}"
errlog:
   subErrLog ("Form2:mSearchFile_Click")
End Sub

Public Sub ApplyHeaderIcon()
On Error GoTo errlog

   Dim ia As Integer
   
   'have to keep doing this to keep it transparent
   Set clsheader = New cHeadIcons
   Set clsheader.ListView = ListView2

'  Call clsheader.SetHeaderIcons(0, Ascn)
   For ia = 0 To ListView2.ColumnHeaders.Count - 1
      If ia = iSortKey Then
          clsheader.SetHeaderIcons iSortKey, FileSortOrder
          Exit For
      End If
   Next ia

errlog:
   subErrLog ("Form2:ApplyHeaderIcon")
End Sub


