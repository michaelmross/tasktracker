VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form5 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "TaskTracker - Virtual Folders"
   ClientHeight    =   2835
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   2835
   ClipControls    =   0   'False
   Icon            =   "Form5.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   2835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   2535
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   0
            Object.Width           =   5292
            MinWidth        =   5292
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frTemp 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   240
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   460
      Visible         =   0   'False
      Width           =   2175
      Begin VB.Label lbTemp 
         BackColor       =   &H80000005&
         Caption         =   "Drag files to this window to create a new virtual folder."
         ForeColor       =   &H8000000C&
         Height          =   375
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   2
         Top             =   6
         Visible         =   0   'False
         Width           =   1935
         WordWrap        =   -1  'True
      End
   End
   Begin ComctlLib.ListView ListView3 
      Height          =   3015
      Left            =   -20
      TabIndex        =   0
      Top             =   -25
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5318
      View            =   1
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      OLEDropMode     =   1
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Virtual Folders"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnVirtual 
      Caption         =   "Virtual Folder Menu"
      Visible         =   0   'False
      Begin VB.Menu mOpen 
         Caption         =   "&View Files..."
      End
      Begin VB.Menu mRename 
         Caption         =   "&Rename Virtual Folder"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mDelete 
         Caption         =   "&Delete Virtual Folder"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mIgnoreFilters 
         Caption         =   "&Ignore Filters"
         Checked         =   -1  'True
         Enabled         =   0   'False
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mSaveFolder 
         Caption         =   "Save as File Folder..."
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ©  2004-2006 Michael M. Ross, Wordwise Solutions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Const FOF_ALLOWUNDO = &H40 ' Allow to undo rename, delete ie sends to recycle bin
Private Const FOF_NOCONFIRMATION = &H10  ' No File Delete or Overwrite Confirmation Dialog

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const LB_SETSEL = &H185&

'Public blForm5Load As Boolean
Private Const LVS_EX_TRACKSELECT As Long = &H8

Public lItemIndex As Long

Private Const CSIDL_DESKTOP = &H0

Public blVirtualFolderDrag As Boolean
Public blTransparent As Boolean
Private vchange As Boolean
Private bdPath As String
Private bdFolder As String
Private bdFolderPath As String

Public Sub Form_Click()
    Me.ZOrder
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo errlog
   If KeyAscii = 19 Then  'ctrl+S
     mOpen_Click
     VirtualFolderDrag
   End If
   If KeyAscii = 6 Then    'ctrl+F
      Form2.mFilter_Click
   End If
errlog:
   subErrLog ("Form5:Form_KeyPress")
End Sub

Private Sub Form_Resize()
On Error GoTo errlog
   Dim colw As Long
   ListView3.Width = Me.Width - 50
   ListView3.Height = Me.Height - (StatusBar1.Height + 330)
   StatusBar1.Panels(1).Width = StatusBar1.Width
   
   colw = ListView3.ColumnHeaders.Count - 1
   Call SendMessage(ListView3.hwnd, LVM_SETCOLUMNWIDTH, colw, ByVal LVSCW_AUTOSIZE_USEHEADER)
   
errlog:
   subErrLog ("Form5:Form_Resize")
End Sub

Private Sub ListView3_AfterLabelEdit(Cancel As Integer, newstring As String)
   blVirtual = True
   Form2.SetTitle (newstring)
   blVirtual = False
End Sub

'Private blNewFolder As Boolean

Private Sub listview3_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errlog
   If KeyCode = vbKeyReturn Then
      mOpen_Click
   End If
   If KeyCode = vbKeyF2 Then
      ListView3.StartLabelEdit
   End If
   If KeyCode = vbKeyDelete Then
      mDelete_Click
   End If
errlog:
   subErrLog ("Form5:listview3_KeyDown")
End Sub

Private Sub Form_Load()
On Error GoTo errlog
   
'   If blForm5Load = False Then
'      Unload Me
'      Exit Sub
'   End If

   Call SendMessage(ListView3.hwnd, _
                    LVM_SETEXTENDEDLISTVIEWSTYLE, _
                    LVS_EX_TRACKSELECT, _
                    ByVal True)
                    
   Set ListView3.SmallIcons = Form1.ImageList3
   
   ListView3.View = lvwReport
   
   LoadVirtualFolders
      
   If Form1.mSaveSize.Checked = True Or blForm5Loaded = True Then
      If LenB(GetSetting("TaskTracker", "Settings", "form5Width")) > 0 Then
         Form5.Width = Val(GetSetting("TaskTracker", "Settings", "form5Width"))
      Else
         Form5.Width = 2925
      End If
      
      If LenB(GetSetting("TaskTracker", "Settings", "form5Height")) > 0 Then
         Form5.Height = Val(GetSetting("TaskTracker", "Settings", "form5Height"))
      Else
         Form5.Height = 3180
      End If
      
      If LenB(GetSetting("TaskTracker", "Settings", "form5Top")) > 0 Then
         Form5.Top = Val(GetSetting("TaskTracker", "Settings", "form5Top"))
'         form5.mSaveSize.Checked = True
      Else
         StandardPos
         Exit Sub
      End If
      
      If Len(GetSetting("TaskTracker", "Settings", "form5Left")) > 0 Then
         Form5.Left = Val(GetSetting("TaskTracker", "Settings", "form5Left"))
      End If
         
   Else
   
      If blExiting = False Then
         StandardPos
      End If
      
   End If
   
   blForm5Loaded = True
'   blVirtualLoadedOnce = True
   
   If LenB(GetSetting("TaskTracker", "Settings", "IgnoreFilters")) > 0 Then
      mIgnoreFilters.Checked = CBool(GetSetting("TaskTracker", "Settings", "IgnoreFilters"))
   Else
      mIgnoreFilters.Checked = True 'default
   End If
   
'   SortVirtualView
   
errlog:
   subErrLog ("Form5:Form_Load")
End Sub

Private Sub StandardPos()
On Error GoTo errlog
   Form2.Show
   If Form2.Left > Form5.Width Then
      Form5.Left = Form2.Left - Form5.Width
      Form5.Top = Form2.Top
   Else
      If Form2.Top < Form5.Height Then
         Form5.Top = Form2.Top + Height
         Form5.Left = Form2.Left
      Else
        Form5.Top = Form2.Top - Form5.Height
        Form5.Left = (Form2.Left + Form2.Width) - Form5.Width
      End If
   End If
errlog:
   subErrLog ("Form5:StandardPos")
End Sub

Private Sub LoadVirtualFolders()
On Error Resume Next
   
   If DoesFileExist(TTpath + "TTVirtual") Then
      LoadVirtualFoldersData
      Exit Sub
   End If
   
'Legacy....
   
   Dim AllFolders As Variant
   Dim agx As ListItem
   Dim av As Integer
   Dim fn As String, fn1 As String, fn2 As String
   
   If SavedVirtualFolders = False Then
      frTemp.Visible = True
      lbTemp.Visible = True
      StatusBar1.Panels(1).Text = "Drag from TaskTracker's file list..."
      Exit Sub
   End If
   
   AllFolders = GetAllSettings("TaskTracker", "TTI\Virtual")

   For av = LBound(AllFolders, 1) To UBound(AllFolders, 1)
      fn = AllFolders(av, 1)
      fn1 = Left$(fn, InStr(fn, "*/") - 1)
      fn2 = Mid$(fn, InStr(fn, "*/") + 2)
      Set agx = Form5.ListView3.ListItems.Add(, fn2, fn1)
      agx.SmallIcon = Form1.ImageList3.ListImages("Virtual").Key
   Next av
      
errlog:
   subErrLog ("Form5:LoadVirtualFolders")
End Sub

Private Function SavedVirtualFolders() As Boolean
On Error GoTo errlog

   Dim cr As New cRegistry

   cr.ClassKey = HKEY_CURRENT_USER
   cr.SectionKey = "Software\VB and VBA Program Settings\TaskTracker\TTI\Virtual"
   cr.ValueKey = "TaskTracker"
   If cr.KeyExists Then
      SavedVirtualFolders = True
   End If

errlog:
   subErrLog ("Form5:SavedVirtualFolders")
End Function

Private Sub LoadVirtualFoldersData()
On Error Resume Next
   Dim agx As ListItem
   Dim lc As Integer
   Dim fn1 As String, fn2 As String
   Dim NextLine As String
   Dim LastLine As String
   Dim intF As Integer
            
   intF = FreeFile()
   Open TTpath + "TTVirtual" For Binary As #intF
   
   Do Until EOF(intF)

      Line Input #intF, NextLine
      
      If Len(NextLine) = 0 Then Exit Do

      fn1 = Left$(NextLine, InStr(NextLine, "*/") - 1)
      fn2 = Mid$(NextLine, InStr(NextLine, "*/") + 2)
      
      If fn1 + fn2 = LastLine Then GoTo skipdupe
      
      If Len(fn1) > 0 And Len(fn2) > 0 Then
         Set agx = ListView3.ListItems.Add(, fn2, fn1)
         agx.SmallIcon = Form1.ImageList3.ListImages("Virtual").Key
         LastLine = fn1 + fn2
         lc = lc + 1
      Else
         Exit Do
      End If
      
skipdupe:
   
   Loop
   
   'empty file
   If lc = 0 Then
      If Len(fn1) = 0 Or Len(fn2) = 0 Then
         frTemp.Visible = True
         lbTemp.Visible = True
         StatusBar1.Panels(1).Text = "Drag from TaskTracker's file list..."
      End If
   Else
      StatusBar1.Panels(1).Text = "Click a folder to view its contents..."
   End If
   
errlog:
   Close #intF
   subErrLog ("Form5:LoadVirtualFoldersData")
End Sub

Public Sub Form_Unload(Cancel As Integer)
On Error GoTo errlog

   Form1.mVirtual.Checked = False
'   Form2.mFilter.Enabled = True
   blVirtualView = False
   SaveForm5Pos
   SaveSetting "TaskTracker", "Settings", "IgnoreFilters", Trim$(Str$(mIgnoreFilters.Checked))
   Me.Hide
   If vchange Then
      vchange = False
      SaveTTInfo
   End If
   
errlog:
   subErrLog ("Form5:Form_Unload")
End Sub

Private Sub SaveForm5Pos()
   If Form5.WindowState = vbNormal Then
      SaveSetting "TaskTracker", "Settings", "Form5Left", Trim$(Str$(Form5.Left))
      SaveSetting "TaskTracker", "Settings", "Form5Top", Trim$(Str$(Form5.Top))
      SaveSetting "TaskTracker", "Settings", "Form5Height", Trim$(Str$(Form5.Height))
      SaveSetting "TaskTracker", "Settings", "Form5Width", Trim$(Str$(Form5.Width))
   End If
End Sub

Private Sub ListView3_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo errlog
   If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
      mOpen_Click
   End If
errlog:
   subErrLog ("Form5:ListView3_KeyUp")
End Sub

Private Sub ListView3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo errlog

   Dim lvhti As LVHITTESTINFO
   Dim lItemIndex As Long
   Dim vi As Integer
   Dim vn As String
   Dim vr As String
   Dim vttext As String
      
   lvhti.pt.x = x \ Screen.TwipsPerPixelX
   lvhti.pt.y = y \ Screen.TwipsPerPixelY
   lItemIndex = SendMessage(ListView3.hwnd, LVM_HITTEST, 0, lvhti) + 1
   
   If m_lCurItemIndex <> lItemIndex Then
      m_lCurItemIndex = lItemIndex
      If m_lCurItemIndex = 0 Then   ' no item under the mouse pointer
         TT.Destroy
      Else
      
         If ListView3.ListItems(m_lCurItemIndex).Key <> ListView3.ListItems(m_lCurItemIndex).Text Then
            TT.Title = ListView3.ListItems(m_lCurItemIndex).Text
            vttext = "(" + ListView3.ListItems(m_lCurItemIndex).Key + ")"
         Else
            TT.Title = ListView3.ListItems(m_lCurItemIndex).Key
         End If
         
         vn = vbNullString
         For vi = 0 To UBound(tType)
            If tType(vi) = ListView3.ListItems(m_lCurItemIndex).Key Then
               vr = Right$(aSFileName(vi), Len(aSFileName(vi)) - InStrRev(aSFileName(vi), "\"))
               If Len(vr) > 0 Then
                  vn = vr + ", " + vn
               End If
            End If
         Next vi
         
         If Len(vttext) > 0 Then
'         Debug.Print vn
            If Len(vn) > 1000 Then
                TT.TipText = vttext + vbNewLine + Left$(vn, 998) + "..."
            Else
                TT.TipText = vttext + vbNewLine + Left$(vn, Len(vn) - 2)
            End If
         Else
            If Len(vn) > 1000 Then
                TT.TipText = Left$(vn, 998) + "..."
            Else
                TT.TipText = Left$(vn, Len(vn) - 2)
            End If
         End If
         
         TT.Create ListView3.hwnd
         TT.MaxTipWidth = 300
         TT.SetDelayTime (sdtInitial), 1000
      End If
   End If

errlog:
   subErrLog ("Form5:ListView3_MouseMove")
End Sub

Private Sub ListView3_OLECompleteDrag(Effect As Long)
On Error GoTo errlog

    If Effect = vbDropEffectCopy Then
        mOpen_Click
        ListView3.OLEDropMode = ccOLEDropManual
        Form1.ListView1.OLEDropMode = ccOLEDropManual
        Form2.ListView2.OLEDropMode = ccOLEDropManual
        VirtualFolderDrag
  End If
    
errlog:
   subErrLog ("Form5:ListView3_OleCompleteDrag")
End Sub

Public Sub VirtualFolderDrag()
On Error GoTo errlog

    Unload VFDrag
    Load VFDrag
    VFDrag.Caption = "Save Virtual Folder - " + ListView3.SelectedItem
    VFDrag.ChangeIllegalChars (ListView3.SelectedItem)
    VFDrag.txFolder.SelStart = 0
    VFDrag.txFolder.SelLength = Len(VFDrag.txFolder)
    
'    Call SetWindowPos(VFDrag.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME)
    VFDrag.Show 1
    
errlog:
   subErrLog ("Form5:VirtualFolderDrag")
End Sub

Private Sub ListView3_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
 On Error GoTo errlog
 blVirtualDrag = True

 Call SetWindowPos(Form5.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME)
 subRestore
 
 If blDragOut = False Then GoTo errlog
 
 If Len(Data.GetData(vbCFText)) > 0 Then
   
   If blVirtualView = True Then
     Exit Sub
   End If
  
   InitVirtualFolder
    
 End If
 
 blVirtualDrag = False
 
 Exit Sub

errlog:
   MsgBox "Add items to virtual folders by dragging them from the TaskTracker file list.", vbExclamation
   subErrLog ("Form5:ListView3_OleDragDrop")
End Sub

Private Sub VDragDrop(Optional Data As DataObject, Optional Effect As Long, Optional Button As Integer, Optional Shift As Integer)
 On Error GoTo errlog
 blVirtualDrag = True
 
 If Len(Data.GetData(vbCFText)) > 0 Then
   
   If blVirtualView = True Then
     Exit Sub
   End If
  
   InitVirtualFolder
    
 End If
 
 blVirtualDrag = False
 
 Exit Sub

errlog:
   MsgBox "Add items to virtual folders by dragging them from the TaskTracker file list.", vbExclamation
   subErrLog ("Form5:VDragDrop")
End Sub


Private Sub ListView3_OLESetData(Data As ComctlLib.DataObject, DataFormat As Integer)
On Error GoTo errlog

 DataFormat = vbCFText
 
errlog:
   subErrLog ("Form5:ListView2_OLESetData")
End Sub

Private Sub ListView3_OLEDragOver(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, state As Integer)
On Error GoTo errlog

   Dim lvhti As LVHITTESTINFO
      
   lvhti.pt.x = x \ Screen.TwipsPerPixelX
   lvhti.pt.y = y \ Screen.TwipsPerPixelY
   lItemIndex = SendMessage(ListView3.hwnd, LVM_HITTEST, 0, lvhti) + 1
      
   Form5.SetFocus
   If lItemIndex > 0 Then
      ListView3.ListItems(lItemIndex).Selected = True
   Else
      Dim lv3 As Integer
      For lv3 = 1 To ListView3.ListItems.Count
         ListView3.ListItems(lv3).Selected = False
      Next lv3
      Call SendMessageLong(ListView3.hwnd, LB_SETSEL, False, -1)
   End If
'   If blVirtualView = True Then
'      ListView3.MousePointer = vbNoDrop
'   End If
   
'ListView3.ListItems(ListView3.SelectedItem.Index).Selected = True
'ListView3.LabelEdit = 1

errlog:
   subErrLog ("Form5:ListView3_OleDragOver")
End Sub

Private Sub Listview3_Mouseup(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo errlog
   If blLoading = False Then
      If Button = 2 Then
         If Form4.ckSave.Value = 1 Then
            mIgnoreFilters.Enabled = True
         Else
            mIgnoreFilters.Enabled = False
         End If
         PopupMenu mnVirtual, , , , mOpen
      Else
         mOpen_Click
         If blTransparent Then
             Call SetWindowPos(Form5.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME)
             subRestore
             blTransparent = False
         End If
         StatusBar1.Panels(1).Text = "Click here to save folder contents..."
      End If
   End If
errlog:
   subErrLog ("Form5:Listview3_Mouseup")
End Sub

Private Sub ListView3_OLEStartDrag(Data As ComctlLib.DataObject, AllowedEffects As Long)
On Error GoTo errlog
   AllowedEffects = vbDropEffectCopy
   ListView3.OLEDropMode = ccOLEDropNone
   Form1.ListView1.OLEDropMode = ccOLEDropNone
   Form2.ListView2.OLEDropMode = ccOLEDropNone
   Data.SetData , vbCFFiles
errlog:
   subErrLog ("Form5:ListView3_OLEStartDrag")
End Sub

Public Sub mDelete_Click()
On Error Resume Next

   If ListView3.ListItems.Count = 0 Then Exit Sub

   Dim df As Integer
   Dim si As Integer
   
   For df = 0 To TTCnt
      If tType(df) = ListView3.SelectedItem.Key Then
         fType(df) = vbNullString
         tType(df) = vbNullString
         blExists(df) = vbNull
         aLastDate(df) = vbNullString
         aShortcut(df) = vbNullString
         aSFileName(df) = vbNullString
      End If
   Next df

   With ListView3
      si = .SelectedItem.Index
      .ListItems.Remove (.SelectedItem.Key)
      If .ListItems.Count > 1 Then
         If si < .ListItems.Count Then
            .ListItems(si).Selected = True
         Else
            .ListItems(si - 1).Selected = True
         End If
      Else
         .ListItems(1).Selected = True
      End If
      
      .Sorted = True
'      SortVirtualView
      
      If blVirtualView = True Then
         mOpen_Click
      End If
      If .ListItems.Count = 0 Then
         frTemp.Visible = True
         lbTemp.Visible = True
         StatusBar1.Panels(1).Text = "Drag from TaskTracker's file list..."
         Form1.ListView1_Click
      End If
   End With
   
   vchange = True
   
errlog:
   subErrLog ("Form5:mDelete_Click")
End Sub

Private Sub mIgnoreFilters_Click()
On Error GoTo errlog

   If mIgnoreFilters.Checked = True Then
      mIgnoreFilters.Checked = False
   Else
      mIgnoreFilters.Checked = True
   End If
   PopFromMem
   Form2.subLbCount

errlog:
   subErrLog ("Form5:mIgnoreFilters_Click")
End Sub

Public Sub mOpen_Click()
On Error GoTo errlog

  With ListView3
      If .ListItems.Count = 0 Then Exit Sub
      If .SelectedItem.Index = 0 Then Exit Sub
      
      blVirtual = True        'set just long enough for selection in popfrommem
      
      blVirtualView = True
      blNoCache = False
      
      Form2.subForm2Init
      
      blNoCache = Form1.mNoCache.Checked
      
      blVirtual = False       'reset
      
      Me.SetFocus
  End With
  
errlog:
   subErrLog ("Form5:mOpen_Click")
End Sub

Private Sub mRename_Click()
On Error GoTo errlog
   ListView3.StartLabelEdit
errlog:
   subErrLog ("Form5:mRename_Click")
End Sub

'Private Sub Timer1_Timer()
'   Timer1.Enabled = False
'   ListView3.StartLabelEdit
'End Sub

Public Sub InitVirtualFolder()
On Error GoTo errlog

  If sType = "Deleted or Renamed" Then
      MsgBox "Deleted items cannot be added to virtual folders.", vbInformation
      Exit Sub
  End If
   
  Dim vf As String
  Dim sl As Integer
  Dim blAddFolderOnce As Boolean
  
  With Form2.ListView2
      For sl = 1 To .ListItems.Count
         If .ListItems(sl).Selected = True Then
            If ListView3.ListItems.Count = 0 Then
            
               If blAddFolderOnce = False Then
                  vf = AddVirtualFolder
                  blAddFolderOnce = True
               End If
               
            ElseIf lItemIndex = 0 Then
            
               If blClickVirtualAdd = False Then
                  If blAddFolderOnce = False Then
                     vf = AddVirtualFolder
                     blAddFolderOnce = True
                  End If
               Else
               
                 If blAddFolderOnce = False Then
                    vf = ListView3.ListItems(ListView3.SelectedItem.Index).Key
                 End If
                 
               End If
            Else
            
              If blAddFolderOnce = False Then
                 vf = ListView3.ListItems(ListView3.SelectedItem.Index).Key
              End If
              
            End If
            Call AddVirtualFileItem(vf, sl)
         End If
       Next sl
   End With
  
errlog:
   StatusBar1.Panels(1).Text = "Click a folder to view its contents..."
   subErrLog ("Form5:InitVirtualFolder")
End Sub

Private Function AddVirtualFolder() As String
On Error GoTo errlog

   Dim lvf As ListItem
   Dim blNotUnique As Boolean
   Dim vn As Integer
   Dim lc As Integer
   Dim vf As String
   
   If ListView3.ListItems.Count = 0 Then
      frTemp.Visible = False
      lbTemp.Visible = False
'      StatusBar1.Panels(1).Text = "Drag from TaskTracker's file list..."
      vf = "Virtual Folder 1"
   Else

      'Determine first unique "Virtual Folder"+int
      vn = 0
      blNotUnique = True
      Do While blNotUnique = True
         vn = vn + 1
         For lc = 1 To ListView3.ListItems.Count
            vf = "Virtual Folder" + Str$(Trim$(vn))
            If vf = ListView3.ListItems(lc).Key Then
               blNotUnique = True
               Exit For
            Else
               blNotUnique = False
            End If
         Next lc
      Loop

   End If
   
   AddVirtualFolder = vf
   
   Set lvf = ListView3.ListItems.Add(, vf, vf)
   lvf.SmallIcon = Form1.ImageList3.ListImages("Virtual").Key
   ListView3.ListItems(vf).Selected = True
   
   StatusBar1.Panels(1).Text = "Click a folder to view its contents..."
   Me.SetFocus
   
   vchange = True
   
errlog:
   subErrLog ("Form5:AddVirtualFolder")
End Function

Private Sub AddVirtualFileItem(vfname As String, SelItm As Integer)
On Error GoTo errlog

   Dim sf As Integer
   Dim TP As String
   Dim LD As String
   Dim SD As String
   
   TTCnt = UBound(fType)
   
   With Form2.ListView2
   
      If iShortCol = 2 Then
           LD = .ListItems(SelItm).SubItems(1)
           SD = .ListItems(SelItm).SubItems(2)
      ElseIf iShortCol = 3 Then
           LD = .ListItems(SelItm).SubItems(2)
           SD = .ListItems(SelItm).SubItems(3)
      ElseIf iShortCol = 4 Then
           LD = .ListItems(SelItm).SubItems(3)
           SD = .ListItems(SelItm).SubItems(4)
      End If
       
      'Check that it's not already added to the virtual folder
      For sf = 0 To TTCnt
         If aShortcut(sf) = SD Then
            If tType(sf) = vfname Then
               Exit Sub
            End If
         End If
      Next sf
   
      TTCnt = TTCnt + 1
      ReDim Preserve fType(TTCnt)
      ReDim Preserve blExists(TTCnt)
      ReDim Preserve blNotValidated(TTCnt)
      ReDim Preserve aSFileName(TTCnt)
      ReDim Preserve aLastDate(TTCnt)
      ReDim Preserve tType(TTCnt)
      ReDim Preserve aShortcut(TTCnt)

      aSFileName(TTCnt) = Form2.fnGetPath(SelItm)
      fType(TTCnt) = RestoreFileType(aSFileName(TTCnt), TTCnt) 'pops the ftype
      tType(TTCnt) = vfname
      aLastDate(TTCnt) = LD
      aShortcut(TTCnt) = SD
      If fType(TTCnt) = "Network Drive" Or fType(TTCnt) = "CD Drive" Or fType(TTCnt) = "Removable Drive" _
          Or fType(TTCnt) = "Deleted or Renamed" Then
          blExists(TTCnt) = False
          blNotValidated(TTCnt) = True
      Else
         If DoesFileExist(aSFileName(TTCnt)) = True Then
            blExists(TTCnt) = True
            blNotValidated(TTCnt) = False
         Else
            blExists(TTCnt) = False
            blNotValidated(TTCnt) = True
         End If
      End If
'      blTrueIcons(TTCnt) = False
   End With
   
'   TTCnt = TTCnt

errlog:
   subErrLog ("Form5:AddVirtualFileItem")
End Sub

Private Sub frTemp_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo errlog
   Call VDragDrop(Data, Effect, Button, Shift)
errlog:
   subErrLog ("Form5:frTemp_OleDragDrop")
End Sub

Private Sub lbTemp_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo errlog
   Call VDragDrop(Data, Effect, Button, Shift)
errlog:
   subErrLog ("Form5:lbTemp_OleDragDrop")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errlog
  If KeyCode = vbKeyEscape Then
     Form_Unload (0)
  End If
  If KeyCode = vbKeyF2 Then
     ListView3.StartLabelEdit
  End If
errlog:
   subErrLog ("Form5:Form_KeyDown")
End Sub

Public Sub FolderPath()
On Error GoTo errlog
   
 bdFolder = VFDrag.txFolder
 If VFDrag.cbTarget.Text = "Desktop" Then
    bdPath = fGetSpecialFolder(CSIDL_DESKTOP)
    bdFolderPath = bdPath + bdFolder
 Else
    bdPath = VFDrag.cbTarget
    If Right$(bdPath, 1) <> "\" Then
      bdFolderPath = bdPath + "\" + bdFolder
    Else
      bdFolderPath = bdPath + bdFolder      'if root
    End If
 End If
 
 If Len(bdFolderPath) > 0 Then
    If DoesFolderExist(bdFolderPath) = True Then
       Dim Response As Integer
       Response = MsgBox("Would you like to replace the existing " + bdFolder + " folder?", vbOKCancel + vbInformation)
       If Response = vbOK Then
          Dim SHFileOp As SHFILEOPSTRUCT
          Dim lresult As Long
          With SHFileOp
              .wFunc = FO_DELETE
              .pFrom = bdFolderPath & vbNullChar & vbNullChar
              .fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION    'put in recycle bin
          End With
          lresult = SHFileOperation(SHFileOp)
       Else
          GoTo errlog
       End If
    End If

    MkDir bdFolderPath
        
    If VFDrag.optShorts.Value = True Then
        SaveFolderShortcuts
    Else
        SaveFolderFiles
    End If
End If

errlog:
  If Err.Number <> 0 Then
     MsgBox "An error occurred when trying to save the folder.", vbExclamation
     Form5.mSaveFolder_Click
  Else
     Unload VFDrag
     Me.SetFocus
  End If
  subErrLog ("Form5:FolderPath")
End Sub

Private Sub SaveFolderFiles()
On Error GoTo errlog

Dim vi As Long
Dim iFail As Integer
Dim blFolder As Boolean
Dim blNoFile As Boolean
Dim ItemPath As String
Dim fNameOnly As String
Dim strFailed As String

  With Form2
     For vi = 1 To .ListView2.ListItems.Count
     
        ItemPath = (.fnSelectedItemsPath(vi))
        If DoesFolderExist(ItemPath) = True Then
           blFolder = True
           iFail = iFail + 1
        ElseIf DoesFileExist(ItemPath) = True Then
           fNameOnly = Right$(ItemPath, Len(ItemPath) - InStrRev(ItemPath, "\"))
           On Error Resume Next
           FileCopy ItemPath, bdFolderPath + "\" + fNameOnly
           If Err.Number > 0 Then
              strFailed = strFailed + ItemPath + vbNewLine
              blNoFile = True
              iFail = iFail + 1
           End If
           On Error GoTo 0
        Else
           strFailed = strFailed + ItemPath + vbNewLine
           blNoFile = True
           iFail = iFail + 1
        End If
        
     Next vi
     
     If iFail = .ListView2.ListItems.Count Then
        RmDir bdFolderPath
        If blFolder = True And blNoFile = False Then
           MsgBox "Folders cannot be copied this way and have been ignored." + vbNewLine + _
                  "Try saving by creating shortcuts instead.", vbInformation
        ElseIf blNoFile = True And blFolder = False Then
           MsgBox "The files could not be found.", vbExclamation
        ElseIf blFolder = True And blNoFile = True Then
           MsgBox "Folders cannot be copied this way, and the files could not be found." + vbNewLine + _
                  "Try saving by creating shortcuts instead.", vbInformation
        End If
        Exit Sub
     End If
     
  End With
      
  If InStr(bdFolderPath, "Desktop") = 0 Then
     Shell Environ("WINDIR") + "\explorer.exe  /e," + bdFolderPath, vbNormalFocus
  End If
  DoEvents
  Sleep (500)
  
  Dim SuccessString As String
  If InStr(bdFolderPath, "Desktop") = 0 Then
    SuccessString = "Folder " + Chr$(34) + bdFolder + Chr$(34) + " successfully created in " + bdPath + "."
  Else
    SuccessString = "Folder " + Chr$(34) + bdFolder + Chr$(34) + " successfully created on your Desktop."
  End If
  
  If blFolder = True And blNoFile = False Then
     MsgBox SuccessString + vbNewLine + "However, some items were folders and have been ignored." + vbNewLine + _
            "Try saving by creating shortcuts instead.", vbInformation
  ElseIf blNoFile = True And blFolder = False Then
     MsgBox SuccessString + vbNewLine + "However, some files could not be found or could not be copied:" + vbNewLine + vbNewLine + strFailed, vbInformation
  ElseIf blFolder = True And blNoFile = True Then
     MsgBox SuccessString + vbNewLine + "However, some items were folders and have been ignored." + vbNewLine + _
            "Some files could not be found or could not be copied:" + vbNewLine + vbNewLine + strFailed, vbInformation
  Else
     MsgBox SuccessString, vbInformation
  End If
  
  Exit Sub
   
errlog:
   MsgBox "An error occurred. The folder may not be complete.", vbExclamation
   subErrLog ("Form5:mSaveFolderFiles")
End Sub

Private Sub SaveFolderShortcuts()
On Error GoTo errlog

Dim blNoFile As Boolean
Dim iFail As Integer
Dim df As String
Dim vi As Long
 
  With Form2
     For vi = 1 To .ListView2.ListItems.Count
        blNoFile = False
        df = .fnSelectedItemsPath(vi)
        If DoesFileExist(df) = False Then
           blNoFile = True
           iFail = iFail + 1
        End If
        If DoesFileExist(TTpath + .ListView2.ListItems(vi).SubItems(iShortCol)) = False Then
            Call MakeAnotherShortcut(bdFolder, bdFolderPath, df)
        Else
          If blNoFile = False Then
             FileCopy TTpath + .ListView2.ListItems(vi).SubItems(iShortCol), bdFolderPath + "\" + .ListView2.ListItems(vi).SubItems(iShortCol)
          End If
        End If
     Next vi
  End With
  
  If iFail = Form2.ListView2.ListItems.Count Then
    RmDir bdFolderPath
    MsgBox "The files could not be found.", vbExclamation
    Exit Sub
  End If
  
  If InStr(bdFolderPath, "Desktop") = 0 Then
     Shell Environ("WINDIR") + "\explorer.exe  /e," + bdFolderPath, vbNormalFocus
  End If
  DoEvents
  Sleep (500)
  
  If iFail > 0 Then
     If InStr(bdFolderPath, "Desktop") = 0 Then
        MsgBox "Folder " + Chr$(34) + bdFolder + Chr$(34) + " successfully created in " + bdPath + "." + vbNewLine + _
               "However, some shortcuts point to files that cannot be found.", vbInformation
     Else
        MsgBox "Folder " + Chr$(34) + bdFolder + Chr$(34) + " successfully created on your Desktop." + vbNewLine + _
               "However, some shortcuts point to files that cannot be found.", vbInformation
     End If
  Else
     If InStr(bdPath, "Desktop") = 0 Then
       MsgBox "Folder " + Chr$(34) + bdFolder + Chr$(34) + " successfully created in " + bdPath + ".", vbInformation
     Else
       MsgBox "Folder " + Chr$(34) + bdFolder + Chr$(34) + " successfully created on your Desktop.", vbInformation
     End If
  End If
   
   Exit Sub
   
errlog:
   MsgBox "An error occurred. The folder may not be complete.", vbExclamation
   subErrLog ("Form5:mSaveFolderShortcuts")
End Sub

Private Sub MakeAnotherShortcut(bdFolder As String, bdPath As String, df As String)
On Error GoTo errlog
Dim fno As String
Dim kr As String
  fno = Right$(df, Len(df) - InStrRev(df, "\"))
  kr = sRecent
  sRecent = bdPath + "\"
  Call ShellShortcut(df, fno)
errlog:
  sRecent = kr
subErrLog ("Form5:MakeAnotherShortcut")
End Sub

Public Sub mSaveFolder_Click()
    mOpen_Click
    VirtualFolderDrag
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As ComctlLib.Panel)
    mOpen_Click
    VirtualFolderDrag
End Sub

Public Sub KeepForm5Visible()
   If Me.Left > Form2.Left And Me.Left < Form2.Left + Form2.Width Then
      If Me.Top > Form2.Top And Me.Top < Form2.Top + Form2.Height Then
         blTransparent = True
         Dim lOldStyle5 As Long
         Dim bTrans As Integer
         Call SetWindowPos(Form5.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME)
         lOldStyle5 = GetWindowLong(Form5.hwnd, GWL_EXSTYLE)
         SetWindowLong Form5.hwnd, GWL_EXSTYLE, lOldStyle5 Or WS_EX_LAYERED
         For bTrans = 255 To 200 Step -1
             SetLayeredWindowAttributes Form5.hwnd, 0, bTrans, LWA_ALPHA
             DoEvents
         Next bTrans
      End If
   End If
End Sub


