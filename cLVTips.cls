VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLVTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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

