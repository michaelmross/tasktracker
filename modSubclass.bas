Attribute VB_Name = "mSubclass"
Option Explicit

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
   (ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
   (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_WNDPROC = -4

Private Const WM_MOUSEMOVE = &H200

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
   (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type HDHITTESTINFO
    pt As POINTAPI
    flags As Long
    iItem As Long
End Type
Private Const HDM_FIRST = &H1200
Private Const HDM_HITTEST = HDM_FIRST + 6
Private Const HDM_GETITEMA = (HDM_FIRST + 3)
Private Const HDM_GETITEM = HDM_GETITEMA

Private Type HD_ITEM
    mask As Long
    cxy As Long
    pszText As String
    hbm As Long
    cchTextMax As Long
    fmt As Long
    lParam As Long
    ' 4.70:
    iImage As Long
    iOrder As Long
End Type
Private Const HDI_LPARAM = &H8

Private Type TLoHiLong
   Lo As Integer
   Hi As Integer
End Type
Private Type TAllLong
   All As Long
End Type
Dim mLH As TLoHiLong, mAL As TAllLong


Private m_lPrevWndProc As Long
Private m_lCurHdrItem As Long

Public Sub Hook(ByVal pHwnd As Long)
   m_lPrevWndProc = SetWindowLong(pHwnd, GWL_WNDPROC, AddressOf WindowProc)
   m_lCurHdrItem = -1
End Sub

Public Sub Unhook(ByVal pHwnd As Long)
   SetWindowLong pHwnd, GWL_WNDPROC, m_lPrevWndProc
End Sub

Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, _
      ByVal wParam As Long, ByVal lParam As Long) As Long
   
   Dim hti As HDHITTESTINFO
   Dim lCol As Long
   
   If uMsg = WM_MOUSEMOVE Then
      ' The low and high words of lParam contains x and y coordinates
      ' of the mouse pointer respectively:
      mAL.All = lParam
      LSet mLH = mAL
      hti.pt.x = mLH.Lo
      hti.pt.y = mLH.Hi
      ' retrieving the index of the header item under the mouse pointer:
      SendMessage hwnd, HDM_HITTEST, 0&, hti
      ' if the current header changed...
      If hti.iItem <> m_lCurHdrItem Then
         m_lCurHdrItem = hti.iItem
         Form1.TT.RemoveToolTip
         If m_lCurHdrItem <> -1 Then
            Form1.TT.InitToolTip hwnd, "Multiline tooltip" & vbCrLf & "for " & Form1.ListView1.ColumnHeaders(m_lCurHdrItem + 1)
         End If
      End If
   End If
   
   WindowProc = CallWindowProc(m_lPrevWndProc, hwnd, uMsg, wParam, lParam)
End Function


