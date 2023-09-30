Attribute VB_Name = "mWndProc"
Option Explicit
'
' Copyright © 1997-1999 Brad Martinez, http://www.mvps.org
'
' Code was written in and formatted for 8pt MS San Serif

#Const WIN32_IE = &H400

' The NMHDR structure contains information about a notification message. The pointer
' to this structure is specified as the lParam member of the WM_NOTIFY message.
Public Type NMHDR
  hwndFrom As Long   ' Window handle of control sending message
  idFrom As Long        ' Identifier of control sending message
  code  As Long          ' Specifies the notification code
End Type

Private Const WM_DESTROY = &H2
Private Const WM_SIZE = &H5
Private Const WM_NOTIFY = &H4E
Private Const WM_MOUSEMOVE = &H200

Private Const WM_USER = &H400
#If (WIN32_IE >= &H300) Then
Private Const TTM_POP = (WM_USER + 28)
#End If

Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Public Enum GWL_nIndex
  GWL_WNDPROC = (-4)
'  GWL_HWNDPARENT = (-8)
  GWL_ID = (-12)
  GWL_STYLE = (-16)
  GWL_EXSTYLE = (-20)
'  GWL_USERDATA = (-21)
End Enum

Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As GWL_nIndex) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As GWL_nIndex, ByVal dwNewLong As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const OLDWNDPROC = "OldWndProc"
Private Const OBJECTPTR = "ObjectPtr"

#If DEBUGWINDOWPROC Then
  ' maintains a WindowProcHook reference for each subclassed window.
  ' the window's handle is the collection item's key string.
  Private m_colWPHooks As New Collection
#End If
'

Public Function SubClass(hWnd As Long, _
                                         lpfnNew As Long, _
                                         Optional objNotify As Object = Nothing) As Boolean
  Dim lpfnOld As Long
  Dim fSuccess As Boolean
  On Error GoTo Out
  
  If GetProp(hWnd, OLDWNDPROC) Then
    SubClass = True
    Exit Function
  End If
  
#If (DEBUGWINDOWPROC = 0) Then
    lpfnOld = SetWindowLong(hWnd, GWL_WNDPROC, lpfnNew)

#Else
    Dim objWPHook As WindowProcHook
    
    Set objWPHook = CreateWindowProcHook
    m_colWPHooks.Add objWPHook, CStr(hWnd)
    
    With objWPHook
      Call .SetMainProc(lpfnNew)
      lpfnOld = SetWindowLong(hWnd, GWL_WNDPROC, .ProcAddress)
      Call .SetDebugProc(lpfnOld)
    End With

#End If
  
  If lpfnOld Then
    fSuccess = SetProp(hWnd, OLDWNDPROC, lpfnOld)
    If (objNotify Is Nothing) = False Then
      fSuccess = fSuccess And SetProp(hWnd, OBJECTPTR, ObjPtr(objNotify))
    End If
  End If
  
Out:
  If fSuccess Then
    SubClass = True
  
  Else
    If lpfnOld Then Call SetWindowLong(hWnd, GWL_WNDPROC, lpfnOld)
    MsgBox "Error subclassing window &H" & Hex(hWnd) & vbCrLf & vbCrLf & _
                  "Err# " & Err.Number & ": " & Err.Description, vbExclamation
  End If
  
End Function

Public Function UnSubClass(hWnd As Long) As Boolean
  Dim lpfnOld As Long
  
  lpfnOld = GetProp(hWnd, OLDWNDPROC)
  If lpfnOld Then
    
    If SetWindowLong(hWnd, GWL_WNDPROC, lpfnOld) Then
      Call RemoveProp(hWnd, OLDWNDPROC)
      Call RemoveProp(hWnd, OBJECTPTR)

#If DEBUGWINDOWPROC Then
      ' remove the WindowProcHook reference from the collection
      m_colWPHooks.Remove CStr(hWnd)
#End If
      
      UnSubClass = True
    
    End If   ' SetWindowLong
  End If   ' lpfnOld

End Function

' Returns the specified object reference stored in the subclassed
' window's OBJECTPTR window property.
' The object reference is valid for only as long as the calling proc holds it.

Public Function GetObj(hWnd As Long) As Object
  Dim Obj As Object
  Dim pObj As Long
  pObj = GetProp(hWnd, OBJECTPTR)
  If pObj Then
    MoveMemory Obj, pObj, 4
    Set GetObj = Obj
    MoveMemory Obj, 0&, 4
  End If
End Function

Public Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  
'Dim s As String
's = GetTTMsgStr(uMsg)
'If Len(s) Then Debug.Print s

  Select Case uMsg
  
    ' ============================================================
    ' Messages that must be handled by the cLVItemTips class
    
    Case WM_SIZE, WM_MOUSEMOVE, WM_NOTIFY
      Dim dwRet As Long

      ' If dwRet is non-zero, eat the message (don't call CallWindowProc,
      ' DoLVMessage is a Friend procedure in Form1 exposing it's cLVItemTips).
      dwRet = Form1.ListViewItemTip.DoLVMessage(hWnd, uMsg, wParam, lParam)
      If dwRet Then
        WndProc = dwRet
        Exit Function
      End If
    
    ' The real listview automatically sends this message to dismiss its tooltip
    ' if the listview sees that its internal code has not explicitly shown the
    ' tooltip, so we'll have things our way and eat the message...
    Case TTM_POP
      dwRet = Form1.ListViewItemTip.DoTTMessage(hWnd, uMsg, wParam, lParam)
      If dwRet Then
        WndProc = dwRet
        Exit Function
      End If
      
    ' ======================================================
    ' Unsubclass the window.
    
    Case WM_DESTROY
      ' OLDWNDPROC will be gone after UnSubClass is called!
      Call CallWindowProc(GetProp(hWnd, OLDWNDPROC), hWnd, uMsg, wParam, lParam)
      Call UnSubClass(hWnd)
      Exit Function
      
  End Select
  
  WndProc = CallWindowProc(GetProp(hWnd, OLDWNDPROC), hWnd, uMsg, wParam, lParam)
  
End Function

' Returns a common control's generic notification code string from it's value.

Public Function GetCtrlNotifyCodeStr(code As Long) As String
  Dim sMsg As String
  Const NM_FIRST = -0&        ' (0U-  0U)
  
  Select Case code
    Case (NM_FIRST - 1): sMsg = "NM_OUTOFMEMORY"
    Case (NM_FIRST - 2): sMsg = "NM_CLICK"
    Case (NM_FIRST - 3): sMsg = "NM_DBLCLK"
    Case (NM_FIRST - 4): sMsg = "NM_RETURN"
    Case (NM_FIRST - 5): sMsg = "NM_RCLICK"
    Case (NM_FIRST - 6): sMsg = "NM_RDBLCLK"
    Case (NM_FIRST - 7): sMsg = "NM_SETFOCUS"
    Case (NM_FIRST - 8): sMsg = "NM_KILLFOCUS"

#If (WIN32_IE >= &H300) Then
    Case (NM_FIRST - 12): sMsg = "NM_CUSTOMDRAW"
    Case (NM_FIRST - 13): sMsg = "NM_HOVER"
#End If

#If (WIN32_IE >= &H400) Then
    Case (NM_FIRST - 14): sMsg = "NM_NCHITTEST"
    Case (NM_FIRST - 15): sMsg = "NM_KEYDOWN"
    Case (NM_FIRST - 16): sMsg = "NM_RELEASEDCAPTURE"
    Case (NM_FIRST - 17): sMsg = "NM_SETCURSOR"
    Case (NM_FIRST - 18): sMsg = "NM_CHAR"
#End If
    
    Case Else: sMsg = GetTTNotifyCodeStr(code)
  End Select
  
  GetCtrlNotifyCodeStr = sMsg

End Function

' Returns a tooltip control's notification code string from it's value

Public Function GetTTNotifyCodeStr(code As Long) As String
  Dim sMsg As String
   Const TTN_FIRST = -520&    '   (0U-520U)
  
  Select Case code
    Case (TTN_FIRST - 0): sMsg = "TTN_NEEDTEXTA"
    Case (TTN_FIRST - 10): sMsg = "TTN_NEEDTEXTW"
    Case (TTN_FIRST - 1): sMsg = "TTN_SHOW"
    Case (TTN_FIRST - 2): sMsg = "TTN_POP"
    
    Case Else: sMsg = "&H" & Hex(code) & "(" & code & "&)" & " (unknown notification)"
  End Select
  
  GetTTNotifyCodeStr = sMsg

End Function

' Returns a tooltip window's message string from it's value

Public Function GetTTMsgStr(uMsg As Long) As String
  Dim sMsg As String
  Select Case uMsg
  
    Case (WM_USER + 1): sMsg = "TTM_ACTIVATE"
    Case (WM_USER + 3): sMsg = "TTM_SETDELAYTIME"
    Case (WM_USER + 7): sMsg = "TTM_RELAYEVENT"
    Case (WM_USER + 13): sMsg = "TTM_GETTOOLCOUNT"
    Case (WM_USER + 16): sMsg = "TTM_WINDOWFROMPOINT"
    
    Case (WM_USER + 4): sMsg = "TTM_ADDTOOLA"
    Case (WM_USER + 5): sMsg = "TTM_DELTOOLA"
    Case (WM_USER + 6): sMsg = "TTM_NEWTOOLRECTA"
    Case (WM_USER + 8): sMsg = "TTM_GETTOOLINFOA"
    Case (WM_USER + 9): sMsg = "TTM_SETTOOLINFOA"
    Case (WM_USER + 10): sMsg = "TTM_HITTESTA"
    Case (WM_USER + 11): sMsg = "TTM_GETTEXTA"
    Case (WM_USER + 12): sMsg = "TTM_UPDATETIPTEXTA"
    Case (WM_USER + 14): sMsg = "TTM_ENUMTOOLSA"
    Case (WM_USER + 15): sMsg = "TTM_GETCURRENTTOOLA"

    Case (WM_USER + 50): sMsg = "TTM_ADDTOOLW"
    Case (WM_USER + 51): sMsg = "TTM_DELTOOLW"
    Case (WM_USER + 52): sMsg = "TTM_NEWTOOLRECTW"
    Case (WM_USER + 53): sMsg = "TTM_GETTOOLINFOW"
    Case (WM_USER + 54): sMsg = "TTM_SETTOOLINFOW"
    Case (WM_USER + 55): sMsg = "TTM_HITTESTW"
    Case (WM_USER + 56): sMsg = "TTM_GETTEXTW"
    Case (WM_USER + 57): sMsg = "TTM_UPDATETIPTEXTW"
    Case (WM_USER + 58): sMsg = "TTM_ENUMTOOLSW"
    Case (WM_USER + 59): sMsg = "TTM_GETCURRENTTOOLW"

#If (WIN32_IE >= &H300) Then
    Case (WM_USER + 17): sMsg = "TTM_TRACKACTIVATE"
    Case (WM_USER + 18): sMsg = "TTM_TRACKPOSITION"
    Case (WM_USER + 19): sMsg = "TTM_SETTIPBKCOLOR"
    Case (WM_USER + 20): sMsg = "TTM_SETTIPTEXTCOLOR"
    Case (WM_USER + 21): sMsg = "TTM_GETDELAYTIME"
    Case (WM_USER + 22): sMsg = "TTM_GETTIPBKCOLOR"
    Case (WM_USER + 23): sMsg = "TTM_GETTIPTEXTCOLOR"
    Case (WM_USER + 24): sMsg = "TTM_SETMAXTIPWIDTH"
    Case (WM_USER + 25): sMsg = "TTM_GETMAXTIPWIDTH"
    Case (WM_USER + 26): sMsg = "TTM_SETMARGIN"
    Case (WM_USER + 27): sMsg = "TTM_GETMARGIN"
    Case (WM_USER + 28): sMsg = "TTM_POP"
#End If

#If (WIN32_IE >= &H400) Then
    Case (WM_USER + 29): sMsg = "TTM_UPDATE"
#End If
    
'    Case Else: sMsg = "WM_USER + " & uMsg - WM_USER '& " (unknown control specific msg)"
  
  End Select
  
  If Len(sMsg) Then GetTTMsgStr = sMsg

End Function
