VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTooltip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private mlngStyle As ttStyleEnum
'Private mvarStyle As Long

''Windows API Functions
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
  (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

''Windows API Constants
Private Const WM_USER = &H400
Private Const CW_USEDEFAULT = &H80000000

''Windows API Types
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'Tooltip Window Constants
Private Const TTM_ADDTOOLA = (WM_USER + 4)
Private Const TTM_UPDATETIPTEXTA = (WM_USER + 12)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTM_SETTITLE = (WM_USER + 32)
Private Const TTS_BALLOON = &H40
Private Const TTS_ALWAYSTIP = &H1
Private Const TTF_SUBCLASS = &H10
Private Const TTF_IDISHWND = &H1
Private Const TTM_SETDELAYTIME = WM_USER + 3
Private mlngHwndParentControl As Long
Private Const TOOLTIPS_CLASSA = "tooltips_class32"
Private mvarIcon As ttIconType

'Tooltip Window Types
Private Type TOOLINFO
    lSize As Long
    lFlags As Long
    hwnd As Long
    lId As Long
    lprect As RECT
    hInstance As Long
    lpStr As String
    lParam As Long
End Type

Public Enum ttIconType
    TTNoIcon = 0
    TTIconInfo = 1
    TTIconWarning = 2
    TTIconError = 3
End Enum

Public Enum ttStyleEnum
    TTStandard
    TTBalloon
End Enum

Public Enum TooltipFlagConstants
    ttfalwaystip = &H1
    ttfNoPrefix = &H2
    ttfNoAnimate = &H10
    ttfNoFade = &H20
    ttfballoon = &H40
End Enum

Public Enum SetDelayTimeConstants
    sdtAutoPop = &H2
    sdtInitial = &H3
    sdtReshow = &H1
    sdtAutomatic = &H0
End Enum

Private mvarTitle As String

'private data
Private m_lParentHwnd As Long ' hwnd of the window the tooltip attached to
Private ti As TOOLINFO

Public Function Create(ByVal ParentHwnd As Long, Optional ByVal flags As TooltipFlagConstants) As Boolean
   Dim lWinStyle As Long
   
   If m_lTTHwnd <> 0 Then
      DestroyWindow m_lTTHwnd
   End If
   
   m_lParentHwnd = ParentHwnd
   
   lWinStyle = TTS_ALWAYSTIP
   
    If flags = ttfballoon Then
        lWinStyle = lWinStyle Or TTS_BALLOON
    End If
       
   m_lTTHwnd = CreateWindowEx(0&, _
      TOOLTIPS_CLASSA, _
      vbNullString, _
      lWinStyle, _
      CW_USEDEFAULT, _
      CW_USEDEFAULT, _
      CW_USEDEFAULT, _
      CW_USEDEFAULT, _
      0&, _
      0&, _
      App.hInstance, _
      0&)
               
   With ti
      .lFlags = TTF_SUBCLASS Or TTF_IDISHWND
      .hwnd = m_lParentHwnd
      .lId = m_lParentHwnd '0
      If blcbSearch = True Then
        .lId = fnFixCombo(Form4.cbSearch) 'fix for dropdown
         blcbSearch = False
      Else
        .lId = m_lParentHwnd '0
      End If
     .hInstance = App.hInstance
      .lSize = Len(ti)
   End With
   
   ''add the tooltip structure
   SendMessage m_lTTHwnd, TTM_ADDTOOLA, 0&, ti

   ''if we want a title or we want an icon
   If mvarTitle <> vbNullString Or mvarIcon <> TTNoIcon Then
      SendMessage m_lTTHwnd, TTM_SETTITLE, CLng(mvarIcon), ByVal mvarTitle
   End If
   
End Function

Private Function fnFixCombo(ByVal ctrl As Control) As Long
Dim cbHwnd As Long
   cbHwnd = ctrl.hwnd
   cbHwnd = FindWindowEx(ctrl.hwnd, 0&, vbNullString, vbNullString)
   fnFixCombo = cbHwnd
End Function

Public Property Let MaxTipWidth(ByVal NewWidth As Long)
    If m_lTTHwnd = 0 Then Exit Property
    
    SendMessage m_lTTHwnd, TTM_SETMAXTIPWIDTH, 0, ByVal NewWidth
End Property

Public Property Let Title(ByVal vData As String)
   mvarTitle = vData
   If m_lTTHwnd <> 0 And mvarTitle <> Empty And mvarIcon <> TTNoIcon Then
      SendMessage m_lTTHwnd, TTM_SETTITLE, CLng(mvarIcon), ByVal mvarTitle
   End If
End Property

Public Property Let TipText(ByVal vData As String)
'   mvarTipText = vData
   ti.lpStr = vData
   If m_lTTHwnd <> 0 Then
      SendMessage m_lTTHwnd, TTM_UPDATETIPTEXTA, 0&, ti
   End If
End Property

Private Sub Class_Initialize()
   InitCommonControls
    Me.HwndParentControl = 0
    Me.Style = TTStandard
End Sub

'Set the control you want the tool tip to apply to
Public Property Let HwndParentControl(ByVal lhWnd As Long)
    mlngHwndParentControl = lhWnd
End Property
Public Property Get HwndParentControl() As Long
    HwndParentControl = mlngHwndParentControl
End Property

Private Sub Class_Terminate()
   If m_lTTHwnd = 0 Then Destroy
End Sub

Public Sub Destroy()
   If m_lTTHwnd <> 0 Then
      DestroyWindow m_lTTHwnd
   End If
End Sub

'Style of the ToolTip
Public Property Let Style(ByVal lngToolTipStyle As ttStyleEnum)
    mlngStyle = lngToolTipStyle
End Property

Public Property Get Style() As ttStyleEnum
    Style = mlngStyle
End Property

Public Property Let Icon(ByVal vData As ttIconType)
   mvarIcon = vData
   If m_lTTHwnd <> 0 And mvarTitle <> Empty And mvarIcon <> TTNoIcon Then
      SendMessage m_lTTHwnd, TTM_SETTITLE, CLng(mvarIcon), ByVal mvarTitle
   End If
End Property

Public Sub SetDelayTime(ByVal TimeMode As SetDelayTimeConstants, ByVal Time As Long)
    If m_lTTHwnd = 0 Then Exit Sub
    
    SendMessage m_lTTHwnd, TTM_SETDELAYTIME, TimeMode, ByVal Time
End Sub




