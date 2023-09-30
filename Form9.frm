VERSION 5.00
Begin VB.Form VFDrag 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Save Virtual Folder"
   ClientHeight    =   2055
   ClientLeft      =   2760
   ClientTop       =   3705
   ClientWidth     =   4320
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txFolder 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1755
   End
   Begin VB.ComboBox cbTarget 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1320
      Width           =   2610
   End
   Begin VB.OptionButton optFiles 
      Caption         =   "Copy &Files"
      Height          =   375
      Left            =   2865
      TabIndex        =   5
      Top             =   1560
      Width           =   1695
   End
   Begin VB.OptionButton optShorts 
      Caption         =   "Copy &Shortcuts"
      Height          =   255
      Left            =   2865
      TabIndex        =   4
      Top             =   1225
      Width           =   1665
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   585
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   105
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "S&pecify Location"
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   1005
      Width           =   2115
   End
   Begin VB.Label Label1 
      Caption         =   "Sa&ve as File Folder"
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   2970
   End
End
Attribute VB_Name = "VFDrag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type BROWSEINFO
   hOwner           As Long
   pidlRoot         As Long
   pszDisplayName   As String
   lpszTitle        As String
   ulFlags          As Long
   lpfn             As Long
   lParam           As Long
   iImage           As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = &H1
'Private Const BIF_DONTGOBELOWDOMAIN = &H2
'Private Const BIF_STATUSTEXT = &H4
'Private Const BIF_RETURNFSANCESTORS = &H8
'Private Const BIF_BROWSEFORCOMPUTER = &H1000
'Private Const BIF_BROWSEFORPRINTER = &H2000
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" _
   Alias "SHBrowseForFolderA" _
  (lpBrowseInfo As BROWSEINFO) As Long
  
Private Declare Function SHGetPathFromIDList Lib "shell32" _
   Alias "SHGetPathFromIDListA" _
  (ByVal pidl As Long, _
   ByVal pszPath As String) As Long
   
Private Declare Sub CoTaskMemFree Lib "ole32" _
   (ByVal pv As Long)

Private Sub CancelButton_Click()
On Error GoTo errlog
    Unload VFDrag
errlog:
   subErrLog ("VFDrag:CancelButton_Click")
End Sub

Private Sub cbTarget_Click()
On Error GoTo errlog

  If cbTarget.Text = "Browse..." Then
    
    With VFDrag.cbTarget
'      Call SetWindowPos(Me.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME)

'      Me.Hide
'      Me.Refresh
'      Form5.Refresh
      Dim obd As String
      obd = OpenBrowseDialog
      If LenB(obd) > 0 Then
        .AddItem obd, 0
      End If
'      Me.Show
      
'      Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME)
      
      If LenB(obd) > 0 Then
        If .List(1) <> "Desktop" Then
        .RemoveItem (1)
        ElseIf .List(2) <> "Desktop" And .List(2) <> "Browse..." Then
          .RemoveItem (2)
        End If
      End If
      .ListIndex = 0
    End With
    
    SaveSetting "TaskTracker", "Settings", "VFDragPath", VFDrag.cbTarget.Text
  End If
  
errlog:
   subErrLog ("VFDrag:cbTarget_Click")
End Sub

Private Sub Form_Load()
On Error GoTo errlog
'   If LenB(GetSetting("TaskTracker", "Settings", "VFDragDesktop")) = 0 Then
'      cbTarget.Text = "Desktop"
'   Else
'      ckDesktop.Value = Val(GetSetting("TaskTracker", "Settings", "VFDragDesktop"))
'   End If
  If LenB(GetSetting("TaskTracker", "Settings", "VFDragFolder")) = 0 Then
     optShorts.Value = 1
     optFiles.Value = 0
  Else
     If CStr(GetSetting("TaskTracker", "Settings", "VFDragFolder")) = "True" Then
         optShorts.Value = 0
         optFiles.Value = 1
     Else
         optShorts.Value = 1
         optFiles.Value = 0
     End If
  End If
  If LenB(GetSetting("TaskTracker", "Settings", "VFDragPath")) > 0 Then
      If GetSetting("TaskTracker", "Settings", "VFDragDesktop") = "True" Then
         cbTarget.AddItem "Desktop"
         cbTarget.AddItem GetSetting("TaskTracker", "Settings", "VFDragPath")
      Else
         cbTarget.AddItem GetSetting("TaskTracker", "Settings", "VFDragPath")
         cbTarget.AddItem "Desktop"
      End If
  Else
     cbTarget.AddItem "Desktop"
  End If
  cbTarget.AddItem "Browse..."
  cbTarget.ListIndex = 0
    
errlog:
   subErrLog ("VFDrag:Form_Load")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "TaskTracker", "Settings", "VFDragFolder", Trim$(Str$(optFiles.Value))
End Sub

Private Sub OKButton_Click()
On Error GoTo errlog
  If cbTarget.Text = "Desktop" Then
    SaveSetting "TaskTracker", "Settings", "VFDragDesktop", "True"
  Else
    SaveSetting "TaskTracker", "Settings", "VFDragDesktop", "False"
  End If
  Me.Hide
  Form5.FolderPath
errlog:
   subErrLog ("VFDrag:OKButton_Click")
End Sub

Public Function OpenBrowseDialog() As String
On Error GoTo errlog

  Dim BI As BROWSEINFO
  Dim pidl As Long
  Dim path As String
  Dim pos As Long
    
  With BI
      
     'hwnd of the window that receives messages
     'from the call. Can be your application
     'or the handle from GetDesktopWindow()
      .hOwner = Me.hwnd

     'pointer to the item identifier list specifying
     'the location of the "root" folder to browse from.
     'If NULL, the desktop folder is used.
'      If LenB(GetSetting("TaskTracker", "Settings", "VFDragPath")) = 0 Then
         .pidlRoot = 0&
'      Else
'         .pidlRoot = Val(GetSetting("TaskTracker", "Settings", "VFDragPath")) + vbNullChar
'      End If
     
     'message to be displayed in the Browse dialog
     .lpszTitle = "Select a location to create the " + Chr$(34) + Form5.ListView3.SelectedItem + Chr$(34) + " folder"

     'the type of folder to return.
      .ulFlags = BIF_RETURNONLYFSDIRS
   End With
    
  'show the browse for folders dialog
   pidl = SHBrowseForFolder(BI)
 
  'the dialog has closed, so parse & display the
  'user's returned folder selection contained in pidl
   path = Space$(MAX_PATH)
    
   If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
      pos = InStr(path, Chr$(0))
      OpenBrowseDialog = Left(path, pos - 1)
   End If

   Call CoTaskMemFree(pidl)
  
errlog:
   subErrLog ("VFDrag: OpenBrowseDialog")
End Function

Private Sub txFolder_Change()
    ChangeIllegalChars (txFolder)
End Sub

Public Sub ChangeIllegalChars(thestring As String)
Dim chars As Integer
   txFolder = thestring
   For chars = 1 To Len(thestring)
      If Mid$(thestring, chars, 1) = "\" Or _
         Mid$(thestring, chars, 1) = "/" Or _
         Mid$(thestring, chars, 1) = ":" Or _
         Mid$(thestring, chars, 1) = "*" Or _
         Mid$(thestring, chars, 1) = "?" Or _
         Mid$(thestring, chars, 1) = Chr(34) Or _
         Mid$(thestring, chars, 1) = "<" Or _
         Mid$(thestring, chars, 1) = ">" Or _
         Mid$(thestring, chars, 1) = "|" Then
         txFolder = Left$(thestring, chars - 1) + " -" + Right$(thestring, Len(thestring) - chars)
      End If
   Next
End Sub
