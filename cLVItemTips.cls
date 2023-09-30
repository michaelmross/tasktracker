VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLVItemTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' Copyright © 1997-1999 Brad Martinez, http://www.mvps.org
'
' Demostrates how to show tooltips for ListItems or SubItems whose label text is not
' fully visible in all ListView views. Works with both Comctl32.ocx and Mscomctl.ocx
' ListViews, the real listview, and on Comctl32.dll v4.70 (IE3) and greater.
'
' Code was written in and formatted for 8pt MS San Serif
'
' tooltip window defaults attribute values:
'   DelayInitial = 500  (1/2 sec)
'   DelayAutoPopup = 5000  (5 secs)
'   DelayReshow = 100 (1/10 sec)
'   MaxTipWidth = -1  (no wordwrap)
'   all Margins = 0 (includes a 2 pixel default padding around text)
'
' todo:
' add Enabled, ReshowDistance (nMouseMoves), delaytime props
'
' ===========================================================================
' module level variables

Private m_hwndTT As Long
Private m_hwndTTOld As Long
Private m_hwndLV As Long

Private m_ti As TOOLINFO

' ANSI and Unicode string buffers filled on TTN_NEEDTEXT
Private m_lpszToolA As Long
Private m_lpszToolW As Long

' set to True if the current listview is using pre IE4 definitions
' (including the statically linked Mscomctl.ocx ListView)
Private m_fIsLVPreIE4 As Boolean

' Set if the tooltip is a Unicode window (NT only), and expects Unicode text
Private m_fIsTTUnicode As Boolean

' ===========================================================================
' general definitions

' using IE4 definitions
#Const WIN32_IE = &H400

Private Const WM_SIZE = &H5
Private Const WM_SETFONT = &H30
Private Const WM_GETFONT = &H31
Private Const WM_NOTIFY = &H4E
Private Const WM_MOUSEMOVE = &H200
Private Const WM_USER = &H400

Private Type POINTAPI   ' pt
  x As Long
  y As Long
End Type

Private Type RECT   ' rct
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

' The NMHDR structure contains information about a notification message. The pointer
' to this structure is specified as the lParam member of the WM_NOTIFY message.
Private Type NMHDR
  hwndFrom As Long   ' Window handle of control sending message
  idFrom As Long        ' Identifier of control sending message
  code  As Long          ' Specifies the notification code
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As Any) As Long   ' lpPoint As POINTAPI
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As Any) As Long   ' lpPoint As POINTAPI
Private Declare Function PtInRect Lib "user32" (lprc As RECT, ByVal x As Long, ByVal y As Long) As Boolean
Private Declare Function InflateRect Lib "user32" (lprc As RECT, ByVal dx As Long, ByVal dy As Long) As Long

Private Declare Function IsWindowUnicode Lib "user32" (ByVal hWnd As Long) As Long

' SetWindowPos hwndInsertAfter
Private Const HWND_TOP = 0

' SetWindowPos uFlags
Private Const SWP_NOSIZE = &H1
'Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40

' All coordinates for child windows are client coordinates (relative
' to the upper-left corner of the parent window's client area).
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hwndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                            (ByVal hWnd As Long, _
                            ByVal wMsg As Long, _
                            ByVal wParam As Long, _
                            lParam As Any) As Long   ' <---

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" _
                            (ByVal dwExStyle As Long, ByVal lpClassName As String, _
                             ByVal lpWindowName As String, ByVal dwStyle As Long, _
                             ByVal x As Long, ByVal y As Long, _
                             ByVal nWidth As Long, ByVal nHeight As Long, _
                             ByVal hwndParent As Long, ByVal hMenu As Long, _
                             ByVal hInstance As Long, lpParam As Any) As Long

Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Const WS_DISABLED = &H8000000

Private Const GWL_STYLE = (-16)
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
'Private Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long
'Private Declare Function lstrlenW Lib "kernel32" (lpString As Any) As Long

' LocalAlloc uFlags
Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40
Private Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)

Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long

' Converts a ANSI string to a Unicode string.
' Specify -1 for cchMultiByte and 0 for cchWideChar to return string length.
Private Declare Function MultiByteToWideChar Lib "kernel32" _
                            (ByVal CodePage As Long, _
                            ByVal dwFlags As Long, _
                            lpMultiByteStr As Any, _
                            ByVal cchMultiByte As Long, _
                            lpWideCharStr As Any, _
                            ByVal cchWideChar As Long) As Long
' CodePage
Private Const CP_ACP = 0        ' ANSI code page

' ===========================================================================
' listview definitions

Private Const MAX_LVITEM = 2048

Private Enum LVViewStyles
  LVS_ICON = &H0
  LVS_REPORT = &H1
  LVS_SMALLICON = &H2
  LVS_LIST = &H3
  LVS_TYPEMASK = &H3
End Enum

' messages
Private Const LVM_FIRST = &H1000
Private Const LVM_GETIMAGELIST = (LVM_FIRST + 2)
Private Const LVM_GETITEMA = (LVM_FIRST + 5)
Private Const LVM_GETSTRINGWIDTHA = (LVM_FIRST + 17)
Private Const LVM_GETCOLUMNWIDTH = (LVM_FIRST + 29)
#If (WIN32_IE >= &H300) Then
Private Const LVM_GETSUBITEMRECT = (LVM_FIRST + 56)
Private Const LVM_SUBITEMHITTEST = (LVM_FIRST + 57)
#End If
#If (WIN32_IE >= &H400) Then
Private Const LVM_SETTOOLTIPS = (LVM_FIRST + 74)
Private Const LVM_GETTOOLTIPS = (LVM_FIRST + 78)
#End If

' LVM_GETIMAGELIST wParam
Private Const LVSIL_NORMAL = 0
Private Const LVSIL_SMALL = 1

' LVM_GET/SETITEM lParam
Private Type LVITEM   ' was LV_ITEM
  mask As Long
  iItem As Long
  iSubItem As Long
  state As Long
  stateMask As Long
  pszText As Long  ' if String, must be pre-allocated before before filled
  cchTextMax As Long
  iImage As Long
  lParam As Long
#If (WIN32_IE >= &H300) Then
  iIndent As Long
#End If
End Type

' LVITEM mask
Private Const LVIF_TEXT = &H1

' LVM_GETITEMRECT rct.Left
Private Enum LVIR_Flags
  LVIR_BOUNDS = 0
  LVIR_ICON = 1
  LVIR_LABEL = 2
  LVIR_SELECTBOUNDS = 3
End Enum

' LVM_HITTEST lParam
Private Type LVHITTESTINFO   ' was LV_HITTESTINFO
  pt As POINTAPI
  flags As Long
  iItem As Long
#If (WIN32_IE >= &H300) Then
  iSubItem As Long    ' this is was NOT in win95.  valid only for LVM_SUBITEMHITTEST
#End If
End Type
 
' LVHITTESTINFO flags
Private Const LVHT_ONITEMICON = &H2
Private Const LVHT_ONITEMLABEL = &H4

' ===========================================================================
' tooltip definitions

Private Const TOOLTIPS_CLASS = "tooltips_class32"

' messages
Private Enum TT_Msgs
  TTM_SETDELAYTIME = (WM_USER + 3)     ' wParam = TTDT_*, lParam = millisecs
  
  TTM_ADDTOOLA = (WM_USER + 4)            ' wParam = 0, lParam = lpti, rtns T/F
  TTM_DELTOOLA = (WM_USER + 5)             ' wParam = 0, lParam = lpti
  TTM_NEWTOOLRECTA = (WM_USER + 6)  ' wParam = 0, lParam = lpti
  TTM_SETTOOLINFOA = (WM_USER + 9)    ' wParam = 0, lParam = lpti
  TTM_ENUMTOOLSA = (WM_USER + 14)     ' wParam = iTool, lParam = lpti, rtns T/F
  TTM_GETCURRENTTOOLA = (WM_USER + 15)  ' wParam = 0, lParam = lpti, rtns T/F

  TTM_ADDTOOLW = (WM_USER + 50)
  TTM_DELTOOLW = (WM_USER + 51)
  TTM_NEWTOOLRECTW = (WM_USER + 52)
  TTM_SETTOOLINFOW = (WM_USER + 54)
  TTM_ENUMTOOLSW = (WM_USER + 58)
  TTM_GETCURRENTTOOLW = (WM_USER + 59)

#If (WIN32_IE >= &H300) Then
  TTM_SETMAXTIPWIDTH = (WM_USER + 24)     ' wParam = 0, lParam = pixels, rtns prev width, -1 = no wordwrap
  TTM_SETMARGIN = (WM_USER + 26)                ' wParam = 0, lParam = lprc (rc members = respective margin distance)
#End If
  TTM_POP = (WM_USER + 28)
End Enum   ' TT_Msgs

' TTM_SETDELAYTIME wParam
Private Enum TTDT_Flags
  TTDT_AUTOMATIC = 0
  TTDT_RESHOW = 1
  TTDT_AUTOPOP = 2
  TTDT_INITIAL = 3
End Enum

' lParam for many tooltip messages
Private Type TOOLINFO
  cbSize As Long
  uFlags As Long
  hWnd As Long
  uId As Long
  lprc As RECT
  hinst As Long
  lpszText As Long
#If (WIN32_IE >= &H300) Then
  lParam As Long
#End If
End Type   ' TOOLINFO

' TOOLINFO uFlags
Private Const TTF_SUBCLASS = &H10
#If (WIN32_IE >= &H300) Then
Private Const TTF_TRANSPARENT = &H100
#End If

' TOOLINFO lpszText
Private Const LPSTR_TEXTCALLBACK = -1

' notifications
Private Const TTN_FIRST = -520&     '   (0U-520U)
Private Const TTN_NEEDTEXTA = (TTN_FIRST - 0)    ' is now TTN_GETDISPINFO
Private Const TTN_NEEDTEXTW = (TTN_FIRST - 10)
Private Const TTN_SHOW = (TTN_FIRST - 1)
Private Const TTN_POP = (TTN_FIRST - 2)

' TTN_NEEDTEXTA lParam
Private Type NMTTDISPINFOA
  hdr As NMHDR
  lpszText As Long
  szText(1 To 80) As Byte   'As String * 80
  hinst As Long
  uFlags As Long
#If (WIN32_IE >= &H300) Then
  lParam As Long
#End If
End Type

' TTN_NEEDTEXTW lParam
Private Type NMTTDISPINFOW
  hdr As NMHDR
  lpszText As Long
  szText(1 To 160) As Byte   ' As String * 160
  hinst As Long
  uFlags As Long
#If (WIN32_IE >= &H300) Then
  lParam As Long
#End If
End Type
'

' Clean up...

Private Sub Class_Terminate()
  
  Call UnSubClass(m_hwndTT)
  
  ' Re-assign any tooltip the ListView had, and destroy our tooltip.
  If m_hwndTTOld Then Call ListView_SetToolTips(m_hwndLV, m_hwndTTOld)
  If m_hwndTT Then Call DestroyWindow(m_hwndTT)
  
  ' free the text buffers
  If m_lpszToolA Then LocalFree (m_lpszToolA)
  If m_lpszToolW Then LocalFree (m_lpszToolW)

End Sub

' ========================================================================
' public methods

Public Function Attach(hwndLV As Long, Optional hFont As Long = 0) As Boolean
  Dim lvhti As LVHITTESTINFO

  ' Instead of checking the window's style to verfiy that it is indeed a listview,
  ' we'll just send it a listview specific message. Set the point way the heck off
  ' in the boonies so that we don't hit test any item, if the handle doesn't belong
  ' to a listview, the call will return 0, otherwise it should return -1 indicating that
  ' the call succeeded but mouse is not over an item.
  lvhti.pt.x = &H80000000
  lvhti.pt.y = &H80000000
  If (ListView_SubItemHitTest(hwndLV, lvhti) = 0) Then Exit Function
  
  ' Reinitialize
  Call Class_Terminate
  
  ' Allocate our global tooltip text buffer pointers (they *must* be global in order
  ' for the tooltip to see it when specified on TTN_NEEDTEXT, see Q180646 )
  m_lpszToolA = LocalAlloc(LPTR, MAX_LVITEM)
  If (m_lpszToolA = 0) Then Exit Function
  
  m_lpszToolW = LocalAlloc(LPTR, MAX_LVITEM)
  If (m_lpszToolW = 0) Then
    If (m_lpszToolA = 0) Then Call LocalFree(m_lpszToolA)
    Exit Function
  End If
  
'  Call InitCommonControls()   ' we are working with a listview after all...
  m_hwndLV = hwndLV
  
  ' Create a new tooltip window using all default values (the IE3 listview does
  ' not have a tooltip, or at least one that's easily retrieved anyway...).
  m_hwndTT = CreateWindowEx(0, TOOLTIPS_CLASS, vbNullString, 0, _
                                                    0, 0, 0, 0, m_hwndLV, 0, App.hInstance, 0&)
  If m_hwndTT Then
    
    ' Add a new transparent tool whose bounding rect is the size of the ListView,
    ' specifying for it to ask for text on TTN_NEEDTEXT notifications.
    m_ti.cbSize = Len(m_ti)
    m_ti.uFlags = TTF_SUBCLASS Or TTF_TRANSPARENT
    m_ti.hWnd = m_hwndLV
    m_ti.lParam = -1
    Call GetClientRect(m_hwndLV, m_ti.lprc)
    m_ti.lpszText = LPSTR_TEXTCALLBACK
    If Tooltip_AddTool(m_hwndTT, m_ti) Then
      
      Call Tooltip_SetDelayTime(m_hwndTT, TTDT_INITIAL, 0)
      Call Tooltip_SetDelayTime(m_hwndTT, TTDT_RESHOW, 0)
  
      ' Sets the flag if the tooltip is a Unicode window (NT only), and wants
      ' only Unicode text (most of the time...)
      m_fIsTTUnicode = IsWindowUnicode(m_hwndTT)
      
      ' Set any specified font
      Call UpdateFont(hFont)
      
      ' Subclass the tooltip to prevent it from seeing the TTM_POP that the
      ' ListView sends when we display the tooltip over ListItems that the
      ' ListView doesn't want the tooltip displayed for...
      Call SubClass(m_hwndTT, AddressOf WndProc)
      
      ' Assign our tooltip to the ListView while remove any existing tooltip, saving
      ' it's handle, it will be re-assigned back to the ListView in Class_Terminate
      ' (this message will fail on pre-IE4 listviews since they don't have tooltips).
      m_hwndTTOld = ListView_SetToolTips(m_hwndLV, m_hwndTT)
  
      ' If m_hwndTTOld is 0, then we may be working with a pre IE4 listview,
      ' if the message again returns 0, then we're certain and the flag will be set.
      ' (otheriwse someone already removed the listview's original tooltip...)
      If (m_hwndTTOld = 0) Then
        m_fIsLVPreIE4 = (ListView_SetToolTips(m_hwndLV, m_hwndTT) = 0)
      End If
      
      Attach = CBool(m_hwndTT)
    
    End If   ' Tooltip_AddTool
  End If   ' hwndTT

End Function

Public Sub UpdateFont(hFont As Long)
  
  ' Exit if we're not initialized
  If (m_hwndTT = 0) Then Exit Sub
  
  If hFont Then Call SendMessage(m_hwndTT, WM_SETFONT, hFont, ByVal 0&)

End Sub

Friend Function DoTTMessage(hwndTT As Long, uMsg As Long, wParam As Long, lParam As Long) As Long
  
  ' Exit if we're not initialized
  If (m_hwndTT = 0) Or (m_hwndLV = 0) Or (hwndTT <> m_hwndTT) Then Exit Function

  Select Case uMsg
  
    ' The real listview will automatically send this message and dismisses its tooltip
    ' if the listview sees that the tooltip is being shown when it doesn't normally show
    ' it, even in large icon view, so we'll have things our way and eat the message...
    ' EXCEPT on NT4/Comctl32.dll v4.xx (IE4) and in large icon view, where we
    ' will process the message and let the listview handle positioning of the tooltip.
    
    Case TTM_POP
      If (m_fIsTTUnicode = False) Or m_fIsLVPreIE4 Or (m_fIsTTUnicode And _
           ((GetWindowLong(m_hwndLV, GWL_STYLE) And LVS_TYPEMASK) <> LVS_ICON)) Then
        DoTTMessage = 1
      End If
    
    End Select
    
End Function

Friend Function DoLVMessage(hwndLV As Long, uMsg As Long, wParam As Long, lParam As Long) As Long
  Static fShown As Boolean
  Static nMouseMoves As Integer
  Static fRectReset As Boolean
  Static fNT4Reset As Boolean
  
  Dim dwStyle  As Long
  Dim nmh As NMHDR
  Dim lvhti As LVHITTESTINFO
  Dim rc As RECT
  
  ' Exit if we're not initialized
  If (m_hwndTT = 0) Or (m_hwndLV = 0) Or (hwndLV <> m_hwndLV) Then Exit Function

  ' ============================================================
  ' On NT4/Comctl32.dll v4.xx (IE4) and in large icon view, let the listview handle
  ' sizing and positioning of the tooltip. For some reason the listview will not size
  ' the tooltip window, and it will be single line if we try to position it...
  
  If m_fIsTTUnicode And (m_fIsLVPreIE4 = False) Then
    If ((GetWindowLong(m_hwndLV, GWL_STYLE) And LVS_TYPEMASK) = LVS_ICON) Then
      If (fNT4Reset = False) Then
        fNT4Reset = True
        Call GetClientRect(m_hwndLV, m_ti.lprc)
        Call Tooltip_SetToolInfo(m_hwndTT, m_ti)
      End If
      Exit Function
    Else
      fNT4Reset = False
    End If
  End If
  
  Select Case uMsg
  
  ' ============================================================
  ' Since we're toggling the tooltip tool's bounding rect between the ListView's
  ' client rect and the current ListItem's rect, when the tooltip is set to a current
  ' ListItem's rect and is dismissed by being clicked or a TTDT_AUTOPOP time
  ' out, though its rect was reset to the ListView's client rect on WM_SIZE or
  ' TTM_POP below, the only way found to cause the tooltip to forget the current
  ' ListItem's rect and become set to the ListView's client rect (allowing it to display
  ' over other ListItems), is to remove, and re-add the tool.

  Case WM_MOUSEMOVE

    ' If the tooltip is not shown and has not yet been reset since the last time it
    ' was shown
    If (fShown = False) And (fRectReset = False) Then
      
      ' Reset the tooltip tool on the 5th mouse move
      nMouseMoves = nMouseMoves + 1
      If (nMouseMoves = 5) Then
        nMouseMoves = 0
        fRectReset = True
'Debug.Print "reset"

        ' Get the zero-index tool's info, remove it, set it's rect to the ListView's
        ' client rect, and add it back (TTM_SETTOOLINFOA seemed to work
        ' most, but not all of the time...?).
        Call Tooltip_DeleteTool(m_hwndTT, m_ti)
        Call GetClientRect(m_hwndLV, m_ti.lprc)
        Call Tooltip_AddTool(m_hwndTT, m_ti)

      End If
    End If

  ' ============================================================
  ' Reset the tooltip tool's bounding rect to that of the ListView's client rect so we
  ' get TTN_NEEDTEXT notifications when the mouse is anyhere in the ListView.

  Case WM_SIZE

    ' If the tooltip is not shown and *has* been reset since the last time it was
    ' shown, reset its rect.
    If (fShown = False) And (fRectReset = False) Then
      Call GetClientRect(m_hwndLV, m_ti.lprc)
      Call Tooltip_SetToolInfo(m_hwndTT, m_ti)
    End If
    
  ' ============================================================
  ' Process tooltip notification messages
  
  Case WM_NOTIFY
  
    ' Fill the NMHDR struct from its pointer, exit if we have a notification
    ' from something other than our tooltip
    MoveMemory nmh, ByVal lParam, Len(nmh)
    If (nmh.hwndFrom <> m_hwndTT) Then Exit Function
    
'' debugging code...
'Debug.Print GetCtrlNotifyCodeStr(nmh.code)
    
    Select Case nmh.code
        
      ' ============================================================
      ' Occurs when the tooltip window is requesting the text it is to display.
      
      Case TTN_NEEDTEXTA, TTN_NEEDTEXTW  ' lParam = lp NMTTDISPIFNOA/W
        Dim ttdiA As NMTTDISPINFOA
        Dim ttdiW As NMTTDISPINFOA
          
        ' Get the index of the item or subitem the mouse is over
        Call GetCursorPos(lvhti.pt)
        Call ScreenToClient(m_hwndLV, lvhti.pt)
        Call ListView_SubItemHitTest(m_hwndLV, lvhti)   ' also hit tests iItem
        
        ' If we're over an item or subitem's icon or label...
        If (lvhti.flags And (LVHT_ONITEMLABEL Or LVHT_ONITEMICON)) Then
          
          ' If we got the item or subitem's text pointer in our allocated buffer...
          If GetLVItemTextPtr(m_hwndLV, lvhti.iItem, lvhti.iSubItem, m_lpszToolA) Then
            
            ' Get the listview's current view.
            dwStyle = (GetWindowLong(m_hwndLV, GWL_STYLE) And LVS_TYPEMASK)
            
            ' If the item's text is not fully visible...
            If IsLVItemTextObscured(m_hwndLV, lvhti.iItem, lvhti.iSubItem, dwStyle = LVS_REPORT, m_lpszToolA) Then
              
              ' Reduce the tooltip's top margin if in large icon view or if the imagelist
              ' has no small icon imagelist assigned to it, or shift the tooltip text up 2
              ' pixels if in small icon view.
              If (dwStyle = LVS_ICON) Or (ListView_GetImageList(m_hwndLV, LVSIL_SMALL) = 0) Then
                rc.Top = -2
              ElseIf (dwStyle = LVS_SMALLICON) Then
                rc.Top = -2
                rc.Bottom = 2
              End If
              ' Also reduce the tooltip's left and right 2 pixel default margin by 1 pixel
              rc.Left = -1
              rc.Right = -1
              Call Tooltip_SetMargin(m_hwndTT, rc)

              ' Set the tooltip's new rect to the item's icon and label bounding rect,
              ' that way when the mouse leaves the rect, the tooltip will be hidden,
              ' and we'll get subsequent TTN_NEEDTEXT notifications
              Call ListView_GetSubItemRect(m_hwndLV, lvhti.iItem, lvhti.iSubItem, LVIR_BOUNDS, m_ti.lprc)
              Call Tooltip_NewToolRect(m_hwndTT, m_ti)
'Debug.Print "here"

              ' If in large icon view, and running a pre IE4 listview, set the max tip width
              ' to the item's label width, otherwise set to no wordwarp (pre IE4 listviews
              ' don't automatically size, format and wordwrap the tooltip in large icon view
              ' since they don't do tooltips, it won't be pretty on IE3, but it will work...).
              If m_fIsLVPreIE4 Then
                If (dwStyle = LVS_ICON) Then
                  Call Tooltip_SetMaxTipWidth(m_hwndTT, (m_ti.lprc.Right - m_ti.lprc.Left) + 2)
                Else
                  Call Tooltip_SetMaxTipWidth(m_hwndTT, -1)
                End If
              End If
              
              ' Fill the respective NMTTDISPINFO struct, set the tooltip woindow text
              ' pointer and copy the struct back to it's pointer. On NT the Unicode tooltip
              ' wants to see Unicode text, except on Comctl32.dll 4.70, where it asks for
              ' ANSI text a couple of times before finally using all Unicode text... (!!??)
              If (nmh.code = TTN_NEEDTEXTA) Then
                MoveMemory ttdiA, ByVal lParam, Len(ttdiA)
                ttdiA.lpszText = m_lpszToolA
                MoveMemory ByVal lParam, ttdiA, Len(ttdiA)
              Else   ' (nmh.code = TTN_NEEDTEXTW)
                MoveMemory ttdiW, ByVal lParam, Len(ttdiW)
                Call MultiByteToWideChar(CP_ACP, 0, ByVal m_lpszToolA, -1, ByVal m_lpszToolW, MAX_LVITEM)
                ttdiW.lpszText = m_lpszToolW
                MoveMemory ByVal lParam, ttdiW, Len(ttdiW)
              End If
              
              ' Cancel the notification preventing the tooltip from doing its default processing
              DoLVMessage = 1
            
            End If   ' IsLVItemTextObscured
          End If   ' GetLVItemTextPtr
        End If   ' (lvhti.flags And (LVHT_ONITEMLABEL Or LVHT_ONITEMICON))
      
      ' ============================================================
      ' Occurs when the tooltip window is about to be shown.
  
      Case TTN_SHOW:   ' lParam = lp NMHDR

        ' Set the show flag
        fShown = True

        ' Prevents tooltip window from repeatedly appearing and disappearing when the
        ' TreeView is in modally shown window of compiled executable, see Q189608.
        dwStyle = GetWindowLong(m_hwndTT, GWL_STYLE)
        If (dwStyle And WS_DISABLED) Then
          Call SetWindowLong(m_hwndTT, GWL_STYLE, dwStyle And Not WS_DISABLED)
        End If

        ' Get the index of the item or subitem the mouse is over.
        Call GetCursorPos(lvhti.pt)
        Call ScreenToClient(m_hwndLV, lvhti.pt)
        Call ListView_SubItemHitTest(m_hwndLV, lvhti)   ' also hit tests iItem

        ' If we're over an item's icon or label... (this is redundant since
        ' the tooltip won't be shown if we didn't specify text for it above...)
        If (lvhti.flags And (LVHT_ONITEMICON Or LVHT_ONITEMLABEL)) Then
          
          ' Get the label rect of the item or subitem
          If ListView_GetSubItemRect(m_hwndLV, lvhti.iItem, lvhti.iSubItem, LVIR_LABEL, rc) Then

            ' Convert the rect to screen coords (the tooltip is WS_POPUP)
            Call ClientToScreen(m_hwndLV, rc)
            
            ' If a subitem, shift the rect 4 pixels to the right (the text inside subitem
            ' labels are padded with 6 pixels on either side, less one pixel each for
            ' the tooltip's left window border and margin.
            If lvhti.iSubItem Then rc.Left = rc.Left + 4

            ' Position the tooltip window and make it higher in the z-order than the
            ' listview window, the tooltip will show itself after this notification returns.
            Call SetWindowPos(m_hwndTT, HWND_TOP, rc.Left, rc.Top, 0, 0, _
                                           SWP_NOSIZE Or SWP_NOACTIVATE) ' Or SWP_SHOWWINDOW)

            ' Cancel the notification preventing the tooltip from doing its default processing
            DoLVMessage = 1

          End If   ' ListView_GetsubItemRect
        End If   ' lvhti.flags

      ' ============================================================
      ' Occurs when the tooltip window is about to be hidden

      Case TTN_POP

        ' Clear the show and reset flags
        fShown = False
        fRectReset = False

        ' Resets the tooltip's rect to the ListView's entire client rect, that way we'll
        ' get TTN_NEEDTEXT when the mouse is anyhere else in the ListView.
        Call GetClientRect(m_hwndLV, m_ti.lprc)
        Call Tooltip_NewToolRect(m_hwndTT, m_ti)
  
    End Select   ' WM_NOTIFY
  End Select   ' uMsg

End Function

' ========================================================================
' private listview calls

' Determines if the text in the specied item or subitem's label is partially truncated with
' an ellipsis, or is not fully visible within the ListView's client rect.

'   hwndLV        - listview window handle
'   iItem              - item index
'   iSubItem       - subitem index, is non zero only for report view subitems
'   fReportView  - flag set if listview is in report view
'   lpszText        - long pointer to item's text

' Returns True if the specified item's label text is truncated, returns False otherwise.

Private Function IsLVItemTextObscured(hwndLV As Long, _
                                                                iItem As Long, _
                                                                iSubItem As Long, _
                                                                fReportView As Boolean, _
                                                                lpszText As Long) As Boolean
  Dim cxText As Long
  Dim rcLV As RECT
  Dim cxCol As Long
  Dim rcItem As RECT
  Dim fRet As Boolean
  
  ' Get the specified item's text width and item rect, and the ListView's client rect.
  cxText = ListView_GetStringWidthA(hwndLV, lpszText)
  Call ListView_GetSubItemRect(hwndLV, iItem, iSubItem, LVIR_LABEL, rcItem)
  Call GetClientRect(hwndLV, rcLV)
  
  ' If a subitem's index is specified (indicating report view)...
  If iSubItem Then
    ' The label rect of report view subitems extend the width of the column
    ' and the text inside the label is padded with 6 pixels left and right, and
    ' a 1 pixel margin and a 1 pixel label border top and bottom.
    cxCol = ListView_GetColumnWidth(hwndLV, iSubItem)
    Call InflateRect(rcItem, -6, -2)
    IsLVItemTextObscured = ((cxText + 12) > cxCol) Or (RectInRect(rcLV, rcItem) = False)
  
  Else
    ' Either the main item in report view, or in large icon, small icon or list views...

    If fReportView Then
      ' The label rect of report view main items extend the width of the column
      ' and the text is surrounded by a 1 pixel margin and a 1 pixel label border.
      cxCol = ListView_GetColumnWidth(hwndLV, 0)
      fRet = ((rcItem.Left + cxText + 4) > cxCol)
      Call InflateRect(rcItem, -2, -2)
      IsLVItemTextObscured = fRet Or (RectInRect(rcLV, rcItem) = False)
    
    Else
      ' The label rect in large icon, small icon and list views is the width of the text and
      ' the text also surrounded by 2 pixel margin as are report view main item above.
      fRet = ((cxText + 4) > (rcItem.Right - rcItem.Left))
      Call InflateRect(rcItem, -2, -2)
      IsLVItemTextObscured = fRet Or (RectInRect(rcLV, rcItem) = False)
    
    End If   ' fReportView
  End If   ' iSubItem
  
'Debug.Print "cxCol: " & cxCol, " cxText: " & cxText ', " cxItem: " & rcItem.Right - rcItem.Left

End Function

' Returns True if rc2 lies entirely inside rc1, returns False otherwise.

Private Function RectInRect(rc1 As RECT, rc2 As RECT) As Boolean
  RectInRect = PtInRect(rc1, rc2.Left, rc2.Top) And PtInRect(rc1, rc2.Right, rc2.Bottom)
End Function

' Fills the passed pointer to a text buffer with the ANSI text of the specified listview item or subitem

'   hwndLV    - listview window handle
'   iItem         - item index
'   iSubItem   - subitem index, is non zero only for report view subitems
'   lpszBuf     - allocated text buffer to be filled

' Returns True if lpszBuf was filled successfully, returns False otherwise.

Private Function GetLVItemTextPtr(hwndLV As Long, _
                                                       iItem As Long, _
                                                       iSubItem As Long, _
                                                       lpszBuf As Long) As Boolean
  Dim lvi As LVITEM
  
  lvi.cchTextMax = LocalSize(lpszBuf)   ' lstrlenW(ByVal lpszBuf)   ' for an all vbNullChar string buffer
  If lvi.cchTextMax Then
    lvi.mask = LVIF_TEXT
    lvi.iItem = iItem
    lvi.iSubItem = iSubItem
    lvi.pszText = lpszBuf
    GetLVItemTextPtr = ListView_GetItemA(hwndLV, lvi)
  End If
  
End Function
 
' ===========================================================================
' listview macros
 
Private Function ListView_GetImageList(hWnd As Long, iImageList As Long) As Long
  ListView_GetImageList = SendMessage(hWnd, LVM_GETIMAGELIST, ByVal iImageList, 0)
End Function
 
Private Function ListView_GetItemA(hWnd As Long, pitem As LVITEM) As Boolean
  ListView_GetItemA = SendMessage(hWnd, LVM_GETITEMA, 0, pitem)
End Function
 
Private Function ListView_GetStringWidthA(hwndLV As Long, psz As Long) As Long
  ListView_GetStringWidthA = SendMessage(hwndLV, LVM_GETSTRINGWIDTHA, 0, ByVal psz)
End Function
 
Private Function ListView_GetColumnWidth(hWnd As Long, iCol As Long) As Long
  ListView_GetColumnWidth = SendMessage(hWnd, LVM_GETCOLUMNWIDTH, ByVal iCol, 0)
End Function

#If (WIN32_IE >= &H300) Then

Private Function ListView_SubItemHitTest(hWnd As Long, plvhti As LVHITTESTINFO) As Long
  ListView_SubItemHitTest = SendMessage(hWnd, LVM_SUBITEMHITTEST, 0, plvhti)
End Function

Private Function ListView_GetSubItemRect(hWnd As Long, iItem As Long, iSubItem As Long, _
                                                                    code As Long, prc As RECT) As Boolean
  prc.Top = iSubItem
  prc.Left = code
  ListView_GetSubItemRect = SendMessage(hWnd, LVM_GETSUBITEMRECT, ByVal iItem, prc)
End Function
'
#End If  ' ' WIN32_IE >= &H300
'

#If (WIN32_IE >= &H400) Then

Private Function ListView_SetToolTips(hwndLV As Long, hwndTooltip As Long) As Long
  ListView_SetToolTips = SendMessage(hwndLV, LVM_SETTOOLTIPS, ByVal hwndTooltip, 0)
End Function

Private Function ListView_GetToolTips(hwndLV As Long) As Long
  ListView_GetToolTips = SendMessage(hwndLV, LVM_GETTOOLTIPS, 0, 0)
End Function
'
#End If  ' ' WIN32_IE >= &H400
'

' ===========================================================================
' tooltip macros (all are user defined)

' For some reason it appears that all tooltips are created only as Unicode windows
' on all versions of NT. Though they appear to correctly translate their ANSI equivilent
' messages, we'll be safe and just send them Unicde messages anyway...


Private Function Tooltip_SetDelayTime(hwndTT As Long, flag As TTDT_Flags, iMillisecs As Long) As Boolean
  Tooltip_SetDelayTime = SendMessage(hwndTT, TTM_SETDELAYTIME, flag, ByVal iMillisecs)
End Function

Private Function Tooltip_AddTool(hwndTT As Long, ti As TOOLINFO) As Boolean
  If m_fIsTTUnicode Then
    Tooltip_AddTool = SendMessage(hwndTT, TTM_ADDTOOLW, 0, ti)
  Else
    Tooltip_AddTool = SendMessage(hwndTT, TTM_ADDTOOLA, 0, ti)
  End If
End Function

Private Sub Tooltip_DeleteTool(hwndTT As Long, ti As TOOLINFO)
  If m_fIsTTUnicode Then
    Call SendMessage(hwndTT, TTM_DELTOOLW, 0, ti)
  Else
    Call SendMessage(hwndTT, TTM_DELTOOLA, 0, ti)
  End If
End Sub

Private Sub Tooltip_NewToolRect(hwndTT As Long, ti As TOOLINFO)
  If m_fIsTTUnicode Then
    Call SendMessage(hwndTT, TTM_NEWTOOLRECTW, 0, ti)
  Else
    Call SendMessage(hwndTT, TTM_NEWTOOLRECTA, 0, ti)
  End If
End Sub

Private Sub Tooltip_SetToolInfo(hwndTT As Long, ti As TOOLINFO)
  If m_fIsTTUnicode Then
    Call SendMessage(hwndTT, TTM_SETTOOLINFOW, 0, ti)
  Else
    Call SendMessage(hwndTT, TTM_SETTOOLINFOA, 0, ti)
  End If
End Sub

Private Function Tooltip_EnumTools(hwndTT As Long, iTool As Long, ti As TOOLINFO) As Boolean
  If m_fIsTTUnicode Then
    Tooltip_EnumTools = SendMessage(hwndTT, TTM_ENUMTOOLSW, ByVal iTool, ti)
  Else
    Tooltip_EnumTools = SendMessage(hwndTT, TTM_ENUMTOOLSA, ByVal iTool, ti)
  End If
End Function

Private Function Tooltip_GetCurrentTool(hwndTT As Long, ti As TOOLINFO) As Boolean
  If m_fIsTTUnicode Then
    Tooltip_GetCurrentTool = SendMessage(hwndTT, TTM_GETCURRENTTOOLW, 0, ti)
  Else
    Tooltip_GetCurrentTool = SendMessage(hwndTT, TTM_GETCURRENTTOOLA, 0, ti)
  End If
End Function

#If (WIN32_IE >= &H300) Then

Private Function Tooltip_SetMaxTipWidth(hwndTT As Long, cx As Long) As Boolean
  Tooltip_SetMaxTipWidth = SendMessage(hwndTT, TTM_SETMAXTIPWIDTH, 0, ByVal cx)
End Function

Private Sub Tooltip_SetMargin(hwndTT As Long, lprc As RECT)
  Call SendMessage(hwndTT, TTM_SETMARGIN, 0, lprc)
End Sub
'
#End If   ' (WIN32_IE >= &H300)


