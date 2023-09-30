VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form4 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TaskTracker Filter"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4065
   ClipControls    =   0   'False
   Icon            =   "Form4.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdVirtual 
      Caption         =   "Save Filter Results as &Virtual Folder..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1400
      Width           =   3730
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   1905
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   8819
            MinWidth        =   8819
            Text            =   "Select a date range and/or enter part of a filename."
            TextSave        =   "Select a date range and/or enter part of a filename."
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Default         =   -1  'True
      Height          =   350
      Left            =   3000
      TabIndex        =   1
      Top             =   200
      Width           =   855
   End
   Begin VB.CheckBox ckSave 
      Caption         =   "&Save"
      Height          =   255
      Left            =   3050
      TabIndex        =   4
      Top             =   1000
      Width           =   855
   End
   Begin VB.CheckBox ckAll 
      Caption         =   "&All Files"
      Height          =   375
      Left            =   3050
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.ComboBox cbSearch 
      ForeColor       =   &H80000011&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "Enter search string"
      Top             =   200
      Width           =   2700
   End
   Begin VB.ComboBox cbDates 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   800
      Width           =   2700
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' © 2004-2007 Michael M. Ross, Wordwise Solutions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_GETITEMHEIGHT = &H154
         
'Private Type POINTAPI
'   X As Long
'   Y As Long
'End Type

'Private Type RECT
'   Left As Long
'   Top As Long
'   Right As Long
'   Bottom As Long
'End Type

'Private Declare Function SendMessage Lib "user32" _
'   Alias "SendMessageA" _
'  (ByVal hwnd As Long, _
'   ByVal wMsg As Long, _
'   ByVal wParam As Long, _
'   lParam As Any) As Long

Private Declare Function MoveWindow Lib "user32" _
  (ByVal hwnd As Long, _
   ByVal x As Long, ByVal y As Long, _
   ByVal nWidth As Long, _
   ByVal nHeight As Long, _
   ByVal bRepaint As Long) As Long

Private Declare Function GetWindowRect Lib "user32" _
  (ByVal hwnd As Long, _
   lprect As RECT) As Long

Private Declare Function ScreenToClient Lib "user32" _
  (ByVal hwnd As Long, _
   lpPoint As POINTAPI) As Long
   
Private TTC As CTooltip

Public blForm4Loaded As Boolean
'Public blForm4Showing As Boolean
Public blWasAddType As Boolean
Public blFiltered As Boolean
Public blSaveFilter As Boolean

Private blForm4Exit As Boolean
Private blDontAdd As Boolean
Private blVFNew As Boolean
Private StringLengthNow As Integer
Private StringLengthBefore As Integer

Public Sub cbSearch_Change()
On Error GoTo errlog

    If blLoading = True Then Exit Sub
    If blForm4Exit = True Then
       blForm4Exit = False
       Exit Sub
    End If
    
    With cbSearch

        If .Text <> "Enter search string" Then
           .ForeColor = &H80000008
        Else
           .ForeColor = &H80000011
        End If
        
        Dim css As Integer
        css = .SelStart
        
        subCBState
        
        If fnIgnore Then Exit Sub
        
'        If Len(.Text) = 1 Then
'             If .Text = "," Or .Text = ";" Or .Text = "*" Or .Text = " " Then
'                 Exit Sub
'             End If
'        End If
              
        If .Text <> "Enter search string" Or .Text <> "" Then
           blFiltered = True
        Else
           blFiltered = False
        End If

        
'        If fnIgnore Then
'           If cbDates.ListIndex = 0 Then
'              Exit Sub
'           End If
'        End If
'
'        blFiltered = True
        
        If blForm4Loaded = True Then
           If InStr(.Text, " ") > 0 Or _
           InStr(.Text, ",") > 0 Or _
           InStr(.Text, ";") > 0 Or _
           .SelStart = 0 Then
              subAdvice
           End If
        End If
        
        StringLengthBefore = StringLengthNow
        StringLengthNow = Len(cbSearch.Text)
        
'        If blForm4Showing = True Then
'           blForm4Showing = False
'        Else
'           If blSearchAll = False Then
'              If fnNewString = True Or fnStringIsShorter = True Then      'minimize unnecessary jitters
'                 Form2.subForm2Init
'              End If
'           End If
'        End If
        
        Form4.SetFocus
        .SelStart = css
   
    End With
   
errlog:
   subErrLog ("Form4:cbSearch_Change")
End Sub

'Private Function fnStringIsShorter() As Boolean
'   If StringLengthNow > 1 Then
'      If (StringLengthNow <= StringLengthBefore) Then
'         fnStringIsShorter = True
'      End If
'   Else
'      fnStringIsShorter = True
'   End If
'End Function

Public Sub subCBState()
On Error GoTo errlog

   If fnNewString = True Or cbDates.Text <> "All Dates" Then   'ignore wildcard and delimiter
      blFiltered = True
      ckSave.Enabled = True
   Else
      If cbDates.Text = "All Dates" Then
         ckSave.Enabled = False
         ckSave.Value = 0
      End If
   End If
   If Form2.mFilter.Checked = True Then   'make sure this is only the case when dialog is open
      If ckAll.Value = 1 Then
         If fnNewString = True Or cbDates.Text <> "All Dates" Then
            If blAddType = True Then
               blWasAddType = True
            End If
            blAddType = True
         Else
            If blWasAddType = True Then
               blAddType = True
               blWasAddType = False
            Else
               blAddType = False
               blThisType = True
            End If
         End If
      End If
   End If

errlog:
   subErrLog ("Form4:cbSearch_Change")
End Sub

Public Function fnNewString() As Boolean     'eliminate these from searches before going further
   If Len(cbSearch.Text) > 0 Then
      If cbSearch.Text = "Enter search string" Then Exit Function
      If Len(cbSearch.Text) = 1 Then
         'ignore wildcard, delimiter, etc. if that's all there is...
         If cbSearch.Text = "*" Or cbSearch.Text = ":" Or cbSearch.Text = "." Then
            cmdVirtual.Enabled = False
            Exit Function
         End If
      ElseIf Len(cbSearch.Text) = 2 Then
         If cbSearch.Text = "*." Then
            cmdVirtual.Enabled = False
            Exit Function
         End If
      Else
         If cbSearch.Text = "*.*" Then
            cmdVirtual.Enabled = False
            Exit Function
         End If
      End If
      fnNewString = True
   End If
End Function

Private Sub cbSearch_GotFocus()
On Error GoTo errlog
'   If blDontSelectSearch = False Then
      cbSearch.SelLength = Len(cbSearch.Text)
'   End If
    If cbSearch.Text = "Enter search string" Then
      subAdvice
    End If
errlog:
   subErrLog ("Form4:cbSearch_GotFocus")
End Sub

Public Sub cbSearch_click()
On Error GoTo errlog

   cbSearch.SelLength = Len(cbSearch.Text)
   cbSearch_Change
errlog:
   subErrLog ("Form4:cbSearch_click")
End Sub

Private Sub cbDates_Click()
On Error GoTo errlog
   If Form2.mFilter.Checked = True Then
'      If cbDates.ListIndex <> 0 Then
'         blFiltered = True
'      End If
      subCBState
      SearchScope
      Me.Refresh
      Form2.subForm2Init
      cmdVirtual.Enabled = True
   End If
errlog:
   subErrLog ("Form4:cbDates_Click")
End Sub

Private Sub cbSearch_KeyPress(KeyAscii As Integer)
   If Left$(cbSearch.Text, 1) <> "*" Then
      blDontAdd = True
   End If
End Sub

Private Sub cbSearch_KeyUp(KeyCode As Integer, Shift As Integer)
   If blDontAdd Then
      With cbSearch
         If Len(.Text) < 5 Then
            If InStr(.Text, "*") = 0 Then
               .Text = "*" + .Text
               .SelStart = Len(.Text)
            End If
         End If
      End With
      blDontAdd = False
   End If
End Sub

Private Sub ckAll_Click()
On Error GoTo errlog

   If blDontRepeatClick = True Then
      blDontRepeatClick = False
      Exit Sub
   End If
   If ckSave.Value = 1 Then
      ckAll.Value = 1
      subAllAdvice (2)
      Exit Sub
   Else
      subAllAdvice (1)
   End If
   
   subApply
   cbDates_Click
   
errlog:
   subErrLog ("Form4:ckAll_Click")
End Sub

Private Sub ckSave_Click()
On Error GoTo errlog
'   correct - just a marker for form_unload to check
   If ckSave.Value = 1 Then
      ckAll.Value = 1
   End If
   subSaveAdvice
errlog:
   subErrLog ("Form4:ckSave_Click")
End Sub

Private Sub cmdVirtual_Click()
On Error Resume Next

   Dim ft As String
   Dim sr As String
   If ckAll.Value = 1 Then
      ft = " (All Files)"
   Else
      ft = " (" + Right$(Form2.Caption, Len(Form2.Caption) - 14) + ")"
   End If
   If cbSearch <> "Enter search string" Then
      sr = cbSearch
      sr = cbDates + "; " + sr
   Else
      sr = cbDates
   End If
   ckSave.Value = 1
   blVFNew = True
   Form_Unload (0)
   Form5.Show
   Form2.SelectAllForm2
   Form5.InitVirtualFolder
   Form4.Visible = False
   ckSave.Value = 0
   blSearchAll = False
   Form_Unload (0)
   Form5.ZOrder
   Form5.SetFocus
   Form5.ListView3.SelectedItem.Text = "Filter Results: " + sr + ft
   Form5.ListView3.SelectedItem.Selected = True
   Form5.mOpen_Click

errlog:
   blVFNew = False
   subErrLog ("Form4:cmdVirtual_Click")
End Sub

Private Sub cmdApply_Click()
On Error GoTo errlog
   subApply
   cbDates_Click
errlog:
   subErrLog ("Form4:cmdApply_Click")
End Sub

Private Sub subApply()
On Error GoTo errlog
   If blLoading = False Then
      If fnIgnore Then Exit Sub
      If cbSearch.Text = "Enter search string" Then
          Exit Sub
      End If
      SearchScope
      SearchAdd
      Form2.subForm2Init
      Me.SetFocus
      cbSearch.SelStart = Len(cbSearch.Text)
      cmdVirtual.Enabled = True
   End If
errlog:
   subErrLog ("Form4:subApply")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo errlog
   If KeyAscii = 21 Then  'ctrl+U
      Form1.mVirtual_Click
   End If
errlog:
   subErrLog ("Form4:Form_KeyPress")
End Sub

Private Sub Form_Load()
On Error GoTo errlog

  With cbDates
      .AddItem "All Dates"                   '0
      .AddItem "Today"                       '1
      .AddItem "Yesterday"                   '2
      .AddItem "In Last Week"                '3
      .AddItem "Two Weeks Ago"               '4
      .AddItem "In Last 2 Weeks"             '5
      .AddItem "Three Weeks Ago"             '6
      .AddItem "In Last 3 Weeks"             '7
      .AddItem "Four Weeks Ago"              '8
      .AddItem "In Last Month"               '9
      .AddItem "Two Months Ago"              '10
      .AddItem "Three Months Ago"            '11
      .AddItem "In Last 3 Months"            '12
      .AddItem "Three to 6 Months Ago"       '13
      .AddItem "In Last 6 Months"            '14
      .AddItem "Six Months to 9 Months Ago"  '15
      .AddItem "Nine Months to 1 Year Ago"   '16
      .AddItem "Six Months to 1 Year Ago"    '17
      .AddItem "In Last Year"                '18
      .AddItem "More Than 1 Year Ago"        '19
      .ListIndex = 0
   End With
   
   Set TTC = New CTooltip
   
   TestforFilter
   
errlog:
   subErrLog ("Form4:Form_Load")
End Sub

Public Function TestforFilter() As Boolean
On Error GoTo errlog

   If Len(GetSetting("TaskTracker", "Settings", "SearchFilter")) > 0 Then
      cbSearch.Text = GetSetting("TaskTracker", "Settings", "SearchFilter")
      ckSave.Value = 1
   End If
   
   If Len(GetSetting("TaskTracker", "Settings", "DateFilter")) > 0 Then
      cbDates.ListIndex = Val(GetSetting("TaskTracker", "Settings", "DateFilter"))
      ckSave.Value = 1
   End If
   
   If Len(GetSetting("TaskTracker", "Settings", "Searchall")) > 0 Then
      ckAll.Value = Val(GetSetting("TaskTracker", "Settings", "Searchall"))
   Else
      ckAll.Value = 0
   End If

errlog:
   subErrLog ("Form4:TestFilter")
End Function

Public Sub SetPosition()
On Error GoTo errlog

   If Form1.mSaveSize.Checked = True Then
      If LenB(GetSetting("TaskTracker", "Settings", "form4Top")) > 0 Then
         Form4.Top = Val(GetSetting("TaskTracker", "Settings", "form4Top"))
      Else
         StandardPos
         Exit Sub
      End If
      
      If Len(GetSetting("TaskTracker", "Settings", "form4Left")) > 0 Then
         Form4.Left = Val(GetSetting("TaskTracker", "Settings", "form4Left"))
      End If
   Else
      If blForm4Loaded = False Then
         StandardPos
      End If
   End If
   blForm4Loaded = True
   
errlog:
   subErrLog ("Form4:SetPosition")
End Sub

Private Sub StandardPos()
On Error GoTo errlog
   Form4.Left = Form2.Left
   If Form2.Top > Form4.Height Then
      Form4.Top = Form2.Top - Form4.Height
   Else
      If Form2.Left > Form4.Width Then
         Form4.Top = Form2.Top
         Form4.Left = Form2.Left - Form4.Width
      Else
         Form4.Top = Form2.Top
      End If
   End If
errlog:
   subErrLog ("Form4:StandardPos")
End Sub

'Private Sub SetShapeAndPosition()
'   Me.Height = 700
'   Me.Top = Form2.Top + 700
'   Me.Width = Form2.Width
'   Me.Left = Form2.Left
'   cbDates.Left = Me.Left + 100
'   cbDates.Width = 1000
'   cbSearch.Left = Me.Width - 1100
'   cbSearch.Width = 1000
'   ckAll.Left = Me.Width - 600
'   ckAll.Width = 500
'End Sub

Public Sub ChangeComboHeight()
On Error GoTo errlog

   Dim pt As POINTAPI
   Dim rc As RECT
   Dim newHeight As Long
   Dim oldScaleMode As Long
   Dim numItemsToDisplay As Long
   Dim itemHeight As Long
   
  'how many items should appear in the dropdown?
   numItemsToDisplay = 20

  'Save the current form scalemode, then
  'switch to pixels
   oldScaleMode = Form4.ScaleMode
   Form4.ScaleMode = vbPixels
     
  'get the system height of a single
  'combo box list item
   itemHeight = SendMessage(cbDates.hwnd, CB_GETITEMHEIGHT, 0, ByVal 0)
   
  'Calculate the new height of the combo box. This
  'is the number of items times the item height
  'plus two. The 'plus two' is required to allow
  'the calculations to take into account the size
  'of the edit portion of the combo as it relates
  'to item height. In other words, even if the
  'combo is only 21 px high (315 twips), if the
  'item height is 13 px per item (as it is with
  'small fonts), we need to use two items to
  'achieve this height.
   newHeight = itemHeight * (numItemsToDisplay + 2)
   
  'get the co-ordinates of the combo box
  'relative to the screen
   Call GetWindowRect(cbDates.hwnd, rc)
   pt.x = rc.Left
   pt.y = rc.Top

  'then translate into co-ordinates
  'relative to the form.
   Call ScreenToClient(Form4.hwnd, pt)

  'using the values returned and set above,
  'call MoveWindow to reposition the combo box
   Call MoveWindow(cbDates.hwnd, pt.x, pt.y, cbDates.Width, newHeight, True)
   
  'it's done, so show the new combo height
   Call SendMessage(cbDates.hwnd, CB_SHOWDROPDOWN, True, ByVal 0)
   
  'restore the original form scalemode
  'before leaving
   Form4.ScaleMode = oldScaleMode
   
errlog:
   subErrLog ("Form4:ChangeComboHeight")
End Sub

Public Sub SearchAdd()
On Error GoTo errlog

   Dim blNewString As Boolean
   Dim cs As Integer
   
   blNewString = fnNewString
   
   For cs = 0 To cbSearch.ListCount
      If Len(cbSearch.Text) > 0 Then
         If blNewString = True Then
            If cbSearch.Text <> cbSearch.List(cs) Then
               cbSearch.AddItem cbSearch.Text, 0
               Exit For
            End If
         End If
      End If
   Next cs

errlog:
   subErrLog ("Form4:SearchAdd")
End Sub

'Private Sub Form_Resize()
'   If Me.Height > LastMeHeight Or Me.Width < LastMeWidth Then
'      Me.Width = 4155
'      Me.Height = 1605
'      cbDates.Left = 120
'      cbDates.Top = 200
'      cbDates.Width = 2700
'      cbSearch.Left = 120
'      cbSearch.Top = 720
'      cbSearch.Width = 2700
'      ckAll.Left = 3000
'      ckAll.Top = 440
'      ckAll.Width = 975
'   End If
'   If Me.Width > LastMeWidth Or Me.Height < LastMeHeight Then
'      Me.Height = 800
'      Me.Top = Form2.Top + 700
'      Me.Width = Form2.Width
'      Me.Left = Form2.Left
'      cbDates.Top = 50
'      cbSearch.Left = 3000
'      cbSearch.Top = cbDates.Top
'      ckAll.Left = Me.Width - 600
'      ckAll.Width = 500
'   End If
'   LastMeWidth = Me.Width
'   LastMeHeight = Me.Height
'End Sub

Public Sub Form_Unload(Cancel As Integer)
On Error GoTo errlog

   Dim Response As Integer
   Cancel = 1     'prevent unload
   
   blStop = True
   
   blForm4Exit = True
   Form4.Hide
   If blExiting = False Then
      Form2.mFilter.Checked = False
   End If
   If ckSave.Value = 1 Then
      If blSaveFilter = False Then
         blSaveFilter = True
      End If
      SaveSetting "TaskTracker", "Settings", "DateFilter", (Trim$(Str$(cbDates.ListIndex)))
      SaveSetting "TaskTracker", "Settings", "SearchFilter", cbSearch.Text
   Else
      blSaveFilter = False
      blFiltered = False
      cbDates.ListIndex = 0
      cbSearch.Text = "Enter search string"
      On Error Resume Next
      DeleteSetting "TaskTracker", "Settings", "DateFilter"
      DeleteSetting "TaskTracker", "Settings", "SearchFilter"
   End If
   SaveSetting "TaskTracker", "Settings", "SearchAll", (Trim$(Str$(ckAll.Value)))
'   blStop = False
   subCBState
   If blExiting = False Then
      Form2.SetTitle
      If Form1.mFocusFiles.Checked = True Then
         Form2.SetFocus
      End If
   End If
   
   If blVFNew Then Exit Sub

errlog:
   blSearchAll = False
   blForm4Exit = False
   Form1.ListView1_Click
   Form1.NormalSequence       'need to initiate it since on hold while dialog was open
   subErrLog ("Form4:Form_Unload")
End Sub

Private Sub SearchScope()
On Error GoTo errlog
   If Form4.Visible = True Then
      If ckAll.Value = 0 Then
         blSearchAll = False
         If blWasAddType = True Then
            blAddType = True
            blWasAddType = False
         Else
            blAddType = False
            blThisType = True
         End If
         subCBState
      Else
         If cbDates.Text <> "All Dates" Or fnIgnore = False Then
'            If fnIgnore = False Then
               blSearchAll = True
'            End If
            If blAddType = True Then
               blWasAddType = True
            End If
            blAddType = True
         End If
      End If
   End If
errlog:
   subErrLog ("Form4:SearchScope")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errlog
   If KeyCode = vbKeyEscape Then
      Form_Unload (1)
   End If
errlog:
   subErrLog ("Form4:Form_KeyDown")
End Sub

Private Sub subAdvice()
On Error GoTo errlog

   TTC.Destroy
   If Form1.mHideTips.Checked = True Then Exit Sub
   With TTC
'      .HwndParentControl = Form4.cbSearch.hwnd
      .Icon = 1
      .Title = "Search Advice"
      .TipText = "The automatic wildcard (" + Chr$(34) + "*" + Chr$(34) + ") means the search string can occur anywhere in the filename. " _
                 + "If you delete the first asterisk, only filenames beginning with the string are shown. " _
                 + "You can also add asterisks before, after, or inside a string (one or more times). For example, " + Chr$(34) + "a*bc*.gif" + Chr$(34) + vbNewLine + "" _
                 + "Use a " + Chr$(34) + ":" + Chr$(34) + " (with no spaces) to search for more than one filename or file type at a time. For example, " + Chr$(34) + "*.gif:*.jpg" + Chr$(34) + vbNewLine + "" _
                 + "(TaskTracker treats " + Chr$(34) + "," + Chr$(34) + " " + Chr$(34) + ";" + Chr$(34) + " and " + Chr$(34) + " " + Chr$(34) + " literally as being commas, semicolons, and spaces in a filename.)"
       blcbSearch = True
      .Create cbSearch.hwnd, ttfballoon
      .MaxTipWidth = IIf(1, 300, -1)
      .SetDelayTime (sdtAutoPop), 30000
   End With
   
errlog:
   subErrLog ("Form4:subAdvice")
End Sub

Private Sub subAllAdvice(i As Byte)
On Error GoTo errlog
   TTC.Destroy
   If Form1.mHideTips.Checked = True Then Exit Sub
   With TTC
      .Icon = 1
      .Title = "Search Advice"
      If i = 1 Then
         .TipText = "This option applies a date or string filter to all file types. " _
                 + "If no filters are selected, it has no effect. " _
                 + "(To select all files without applying a filter, use the File Type " _
                 + "context menu's Select All (Ctrl+A) command.)"
      Else
         .TipText = "When you save a filter, it's applied to all file types."
      End If
      .Create ckAll.hwnd, ttfballoon
      .MaxTipWidth = IIf(1, 300, -1)
      .SetDelayTime (sdtAutoPop), 30000
   End With
errlog:
   subErrLog ("Form4:subAllAdvice")
End Sub

Private Sub subSaveAdvice()
On Error GoTo errlog
   TTC.Destroy
   If Form1.mHideTips.Checked = True Then Exit Sub
   With TTC
      .Icon = 1
      .Title = "Search Advice"
      .TipText = "This option saves a date or string filter (applying it to all file types) after closing this dialog."
      .Create ckSave.hwnd, ttfballoon
      .MaxTipWidth = IIf(1, 300, -1)
      .SetDelayTime (sdtAutoPop), 30000
   End With
errlog:
   subErrLog ("Form4:subSaveAdvice")
End Sub

Private Function fnIgnore() As Boolean
   With cbSearch
      If Len(.Text) = 1 Then
           If .Text = "," Or .Text = ";" Or .Text = "*" Or .Text = " " Then
               fnIgnore = True
               Exit Function
           End If
      ElseIf Len(.Text) = 0 Then
         fnIgnore = True
         Exit Function
      End If
      If InStr(.Text, "Enter search string") > 0 Then
          fnIgnore = True
      End If
   End With
End Function
