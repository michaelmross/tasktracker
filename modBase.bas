Attribute VB_Name = "modBase"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ©  2003-2005 Michael M. Ross, Wordwise Solutions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function CreateFile Lib "kernel32" _
   Alias "CreateFileA" _
  (ByVal lpFileName As String, _
   ByVal dwDesiredAccess As Long, _
   ByVal dwShareMode As Long, _
   ByVal lpSecurityAttributes As Long, _
   ByVal dwCreationDisposition As Long, _
   ByVal dwFlagsAndAttributes As Long, _
   ByVal hTemplateFile As Long) As Long

Public Declare Function GetFileTime Lib "kernel32" _
  (ByVal hfile As Long, _
   lpCreationTime As FILETIME, _
   lpLastAccessTime As FILETIME, _
   lpLastWriteTime As FILETIME) As Long
   
Public Declare Function OpenFile Lib "kernel32" _
   (ByVal lpFileName As String, _
    lpReOpenBuff As OFSTRUCT, _
    ByVal wStyle As Long) As Long
    
Public Declare Function GetSystemMetrics Lib "user32" _
   (ByVal nIndex As Long) As Long
   
Public Declare Function GetInputState Lib "user32" () As Long

Public Declare Sub InitCommonControls Lib "comctl32.dll" ()

Public Declare Function FindFirstFile Lib "kernel32" _
   Alias "FindFirstFileA" _
  (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
   
Public Declare Function FindNextFile Lib "kernel32" _
  Alias "FindNextFileA" _
  (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
(ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
      
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter _
    As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long) As Long
    
Public Declare Function ImageList_Draw Lib "comctl32" _
  (ByVal himl&, _
   ByVal i&, _
   ByVal hDCDest&, _
   ByVal x&, _
   ByVal y&, _
   ByVal flags&) As Long

Public Declare Function SHGetFileInfo Lib "shell32" _
   Alias "SHGetFileInfoA" _
  (ByVal pszPath As String, _
   ByVal dwFileAttributes As Long, _
   psfi As SHFILEINFO, _
   ByVal cbSizeFileInfo As Long, _
   ByVal uFlags As Long) As Long
   
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'user defined type required by Shell_NotifyIcon API call
Public Type NOTIFYICONDATA
 cbSize As Long
 hWnd As Long
 uId As Long
 uFlags As Long
 uCallBackMessage As Long
 hIcon As Long
 szTip As String * 64
End Type

'constants required by Shell_NotifyIcon API call:
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONUP = &H205

Public Const GWL_STYLE = (-16)
Public Const WS_MINIMIZEBOX = &H20000

Public Declare Function Shell_NotifyIcon Lib "shell32" _
Alias "Shell_NotifyIconA" _
(ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public nid As NOTIFYICONDATA

Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_SYSICONINDEX = &H4000  'system icon index
Public Const SHGFI_SMALLICON = &H1        'small icon
Public Const ILD_TRANSPARENT = &H1        'display transparent
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_TYPENAME = &H400
Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or _
             SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or _
             SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Public Const MAX_NPATH = 260

Public Const SPI_GETWORKAREA& = 48

Public Const WS_VSCROLL = &H200000

Public Type SHFILEINFO
   hIcon          As Long
   iIcon          As Long
   dwAttributes   As Long
   szDisplayName  As String * MAX_NPATH
   szTypeName     As String * 80
End Type

Public shinfo As SHFILEINFO
Public OFS As OFSTRUCT

Public Const GENERIC_READ = &H80000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
Public Const OPEN_EXISTING = 3
Public Const OFS_MAXPATHNAME = 260

Public Type FILETIME
  dwLowDateTime     As Long
  dwHighDateTime    As Long
End Type

Public Type OFSTRUCT
   cBytes      As Byte
   fFixedDisk  As Byte
   nErrCode    As Integer
   Reserved1   As Integer
   Reserved2   As Integer
   szPathName(OFS_MAXPATHNAME) As Byte
End Type

Public Type SYSTEMTIME
  wYear          As Integer
  wMonth         As Integer
  wDayOfWeek     As Integer
  wDay           As Integer
  wHour          As Integer
  wMinute        As Integer
  wSecond        As Integer
  wMilliseconds  As Long
End Type
   
Public Const SM_TABLETPC = 86
   
Public Type WIN32_FIND_DATA
  dwFileAttributes  As Long
  ftCreationTime    As FILETIME
  ftLastAccessTime  As FILETIME
  ftLastWriteTime   As FILETIME
  nFileSizeHigh     As Long
  nFileSizeLow      As Long
  dwReserved0       As Long
  dwReserved1       As Long
  cFileName         As String * MAX_PATH
  cAlternate        As String * 14
End Type

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST As Long = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_SHOWME = SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE

Public Const LVM_FIRST = &H1000
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54)
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 55)
Public Const LVS_EX_FULLROWSELECT = &H20
Public Const LVS_EX_GRIDLINES = &H1

Public Const LVM_HITTEST = LVM_FIRST + 18
Public Const LVM_SUBITEMHITTEST = (LVM_FIRST + 57)
Public Const LVM_FINDITEM As Long = (LVM_FIRST + 13)
Public Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 45)
Public Const LVM_SORTITEMS As Long = (LVM_FIRST + 48)
Public Const LVM_GETHEADER = (LVM_FIRST + 31)

Public Type LVHITTESTINFO
   pt As POINTAPI
   flags As Long
   iItem As Long
   iSubItem As Long
End Type

Public TT As CTooltip
Public m_lCurItemIndex As Long

   
Public Enum enmScale
   Twips = 1
'   Pixels = 2
End Enum

Public Type RECT   ' rct
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

'Autosize headers
Public Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Public Const LVSCW_AUTOSIZE As Long = -1
Public Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

'Global Vars
Public sType As String        'the file type being displayed
Public TTpath As String
Public sExt As String
'Public ThisFType As String
Public fType() As String      'the parsed file type, including Deleted or Renamed
Public tType() As String      'file type info if D or N, also virtual folder names
Public PrevType() As String
'Public aSmallIcon() As String
'Public aFType() As String
Public aShortcut() As String
Public aLastAcc() As String
Public aLastMod() As String
Public aCreated() As String
Public DroppedList() As String
Public sLastType As String
Public sRecent As String

Public TTCnt As Long    'the main counter for tracked items - shortcuts, files, cached, or virtual
Public vf As Long
Public rf As Long
Public fCount As Long   'the TT folder shortcut count
Public m_lTTHwnd As Long ' hwnd of the tooltip

Public iDaysUsed As Integer
Public wWidth As Integer
Public col1width As Integer
Public col2width As Integer
'Public colSwidth As Integer
Public col1Pwidth As Integer
Public col2Pwidth As Integer
Public col3Pwidth As Integer
Public col4Swidth As Integer
Public col1Awidth As Integer
Public col1NAwidth As Integer
Public col2Awidth As Integer
'Public col2NAwidth As Integer
Public col3Awidth As Integer
Public col3NAwidth As Integer
Public col1PAwidth As Integer
Public col1NPAwidth As Integer
Public col2PAwidth As Integer
Public col2NPAwidth As Integer
Public col3PAwidth As Integer
'Public col3NPAwidth As Integer
Public col4PAwidth As Integer
Public col4NPAwidth As Integer
Public iShortCol As Integer
Public tmCount As Integer
Public RandomMinutesTillExpiration As Integer

Public dLastWrite As Date

Public blExists() As Boolean
Public blNotValidated() As Boolean
Public blDontCopy() As Boolean
Public blExit As Boolean
Public blLoading As Boolean
Public blAllTypes As Boolean
Public blThisType As Boolean
Public blAddType As Boolean
Public blRefresh As Boolean
Public blClicked As Boolean
Public blForm2Show As Boolean
Public blFirstSort As Boolean
Public blExpired As Boolean
Public blDragOut As Boolean
Public blSelectAll As Boolean
Public blMultSelect As Boolean
Public blCtrlKey As Boolean
'Public blWasMultSel As Boolean
Public blUserRefresh As Boolean
Public blDateSwitch As Boolean
Public blComparetoTTDate As Boolean
Public blAltPress As Boolean
Public blFirstLoad As Boolean
Public blBegin As Boolean
Public blSuspendFade As Boolean
Public blcbSearch As Boolean
Public blRegStatus As Boolean
Public blTempStatus As Boolean
Public blForm1Resize As Boolean
Public blForm2Resize As Boolean
Public blTaskbarStatus As Boolean
Public blEndNewInstance As Boolean
Public blMultFileDrop As Boolean
Public blDriveChange As Boolean
Public blSnapshot As Boolean
Public blMenu As Boolean
Public blMinStart As Boolean
Public blColSame As Boolean
Public blFolderSort As Boolean
Public blExiting As Boolean
Public blLoadResize As Boolean
Public blDontRepeatClick As Boolean
'Public blDragDrop As Boolean
Public blColumnClick As Boolean
Public blJustRegd As Boolean
Public blVirtualView As Boolean
Public blClickVirtualAdd As Boolean
Public blVirtualDrag As Boolean
Public blForm6Load As Boolean
Public blPreviewOpen As Boolean
Public blForm5Loaded As Boolean
'Public blVirtualLoadedOnce As Boolean
'Public blSearchAdd As Boolean
Public blTraceLog As Boolean
Public blForm4Loading As Boolean        'for form5 unload only
Public blNotified As Boolean

Public RegType As Variant
Public RegDate As Variant
Public RegName As Variant

'Filled in GetSettings from Registry values
Public FileSortOrder As Byte
Public TypeSortCol As Byte

Public Function flActualScreenHeight(ScaleType As enmScale) As Long
On Error GoTo errlog

   Dim lReturn    As Long
   Dim tRect      As RECT

   lReturn = SystemParametersInfo(SPI_GETWORKAREA, 0&, tRect, 0&)

   If ScaleType = Twips Then
      flActualScreenHeight = tRect.Bottom * Screen.TwipsPerPixelX
   Else
      flActualScreenHeight = tRect.Bottom
   End If

errlog:
   subErrLog ("modBase:flActualScreenHeight")
End Function

Public Function flActualScreenWidth(ScaleType As enmScale) As Long
On Error GoTo errlog

   Dim lReturn    As Long
   Dim tRect      As RECT

   lReturn = SystemParametersInfo(SPI_GETWORKAREA, 0&, tRect, 0&)

   'The data is returned in Pixels, if the user wants Twips, then convert it
   If ScaleType = Twips Then
      flActualScreenWidth = tRect.Right * Screen.TwipsPerPixelY
   Else
      flActualScreenWidth = tRect.Right
   End If

errlog:
   subErrLog ("modBase:flActualScreenWidth")
End Function

Sub Main()
   CheckOS
'   Load Form1
   If Command$ = "-m" Then
      If Form1.mTaskbar.Checked = False Then
         Form1.Hide
      Else
         Form1.WindowState = vbMinimized
      End If
   Else
      Load Form1
   End If
End Sub

Public Sub CheckOS()
   If Str(GetSystemMetrics(SM_TABLETPC)) <> 0 Then
   
      If GetSetting("TaskTracker", "Settings", "COMCTL32") <> "6.0.81.6" Then    'set by installer

         fnOpenURL ("http://tasktracker.wordwisesolutions.com/support/TabletPC.htm")
      
         MsgBox "TaskTracker does not support Tablet PC Edition.", vbCritical
   
         End
      
      End If
      
   End If
End Sub

