VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form6 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Image Preview"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   3885
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   241
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   259
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   3240
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   ""
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3615
      ScaleWidth      =   4215
      TabIndex        =   1
      Top             =   0
      Width           =   4215
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   3480
         Top             =   3000
      End
      Begin VB.Label Label1 
         Caption         =   "Unable to Display Image"
         Height          =   615
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   3615
      End
   End
   Begin VB.PictureBox picLoad 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   2280
      ScaleHeight     =   1335
      ScaleWidth      =   1335
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ©  2004 Michael M. Ross, Wordwise Solutions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private blForm6Loaded As Boolean

Private Type POINTAPI
    x  As Long
    y  As Long
End Type

Private Const SRCCOPY           As Long = &HCC0020
Private Const STRETCH_HALFTONE  As Long = &H4&

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lpPt As POINTAPI) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function UnrealizeObject Lib "gdi32" (ByVal hObject As Long) As Long

'Private Form6Inst() As Form6
'Private FormCount As Integer

Private Sub CreateThumbPic(picSource As PictureBox, picPreview As PictureBox)
'Defer error handling to LoadSelectedFilePicture
'Uses the halftone stretch mode to produce a high-quality resizable image from picLoad.

Dim lRet            As Long
Dim lLeft           As Long
Dim lTop            As Long
Dim lWidth          As Long
Dim lHeight         As Long
Dim lForeColor      As Long
Dim hBrush          As Long
Dim hDummyBrush     As Long
Dim lOrigMode       As Long
Dim fScale          As Single
Dim uBrushOrigPt    As POINTAPI

    picPreview.Width = Form6.Width \ Screen.TwipsPerPixelX - 5
    picPreview.Height = Form6.Height \ Screen.TwipsPerPixelY - 23
    picPreview.Left = -1
    picPreview.Top = -1
    picPreview.BackColor = vbButtonFace
    picPreview.AutoRedraw = True
    picPreview.Cls
    
    If picSource.Width <= picPreview.Width - 2 And picSource.Height <= picPreview.Height - 2 Then
        fScale = 1
    Else
        fScale = IIf(picSource.Width > picSource.Height, (picPreview.Width - 2) / picSource.Width, (picPreview.Height - 2) / picSource.Height)
    End If
    lWidth = picSource.Width * fScale
    lHeight = picSource.Height * fScale
    lLeft = Int((picPreview.Width - lWidth) \ 2)
    lTop = Int((picPreview.Height - lHeight) \ 2)
    
    'Store the original ForeColor
    lForeColor = picPreview.ForeColor
    
    'Set picEdit's stretch mode to halftone (this may cause misalignment of the brush)
    lOrigMode = SetStretchBltMode(picPreview.hDC, STRETCH_HALFTONE)
    
    'Realign the brush...
    'Get picEdit's brush by selecting a dummy brush into the DC
    hDummyBrush = CreateSolidBrush(lForeColor)
    hBrush = SelectObject(picPreview.hDC, hDummyBrush)
    'Reset the brush (This will force windows to realign it when it's put back)
    lRet = UnrealizeObject(hBrush)
    'Set picEdit's brush alignment coordinates to the left-top of the bitmap
    lRet = SetBrushOrgEx(picPreview.hDC, lLeft, lTop, uBrushOrigPt)
    'Now put the original brush back into the DC at the new alignment
    hDummyBrush = SelectObject(picPreview.hDC, hBrush)
    
    'Stretch the bitmap
    lRet = StretchBlt(picPreview.hDC, lLeft, lTop, lWidth, lHeight, _
            picSource.hDC, 0, 0, picSource.Width, picSource.Height, SRCCOPY)
    
    'Set the stretch mode back to it's original mode
    lRet = SetStretchBltMode(picPreview.hDC, lOrigMode)
    
    'Reset the original alignment of the brush...
    'Get picEdit's brush by selecting the dummy brush back into the DC
    hBrush = SelectObject(picPreview.hDC, hDummyBrush)
    'Reset the brush (This will force windows to realign it when it's put back)
    lRet = UnrealizeObject(hBrush)
    'Set the brush alignment back to the original coordinates
    lRet = SetBrushOrgEx(picPreview.hDC, uBrushOrigPt.x, uBrushOrigPt.y, uBrushOrigPt)
    'Now put the original brush back into picEdit's DC at the original alignment
    hDummyBrush = SelectObject(picPreview.hDC, hBrush)
    'Get rid of the dummy brush
    lRet = DeleteObject(hDummyBrush)
    
    'Restore the original ForeColor
    picPreview.ForeColor = lForeColor

'    picPreview.Line (lLeft - 1, lTop - 1)-Step(lWidth + 1, lHeight + 1), &H0&, B
        
End Sub


Public Sub LoadSelectedFilePicture(FilePath As String)
On Error GoTo errlog

   Label1.Visible = False
'   Me.Caption = "Image Preview - Loading..."
   StatusBar1.Panels(1).Text = "Loading..."
   
   Set picLoad.Picture = LoadPicture()
   Set picLoad.Picture = LoadPicture(FilePath, vbLPLargeShell, vbLPDefault)
   
   If InStr(LCase$(FilePath), ".ico") > 0 _
              Or InStr(LCase$(FilePath), ".cur") > 0 Then
                Set picLoad.Picture = LoadPicture(FilePath, vbLPLargeShell, vbLPDefault)
   Else
       Set picLoad.Picture = LoadPicture(FilePath)
   End If '   MsgBox Str(picPreview.Image)
   
   Call CreateThumbPic(picLoad, picPreview)
         
'   Me.Caption = "Image Preview"
   
   StatusBar1.Panels(1).Text = Form2.ListView2.ListItems(Form2.ListView2.SelectedItem.Index).Text

   If Form2.Visible Then
      Me.Show
      KeepForm6Visible
   End If
   
   blPreviewOpen = True
   
   Exit Sub
   
errlog:
   If Form1.mAutoPreview.Checked = True Then
      Me.Hide
   ElseIf Form2.mPreview.Checked = True Then
      If Form2.CheckPreview = False Then
         Me.Hide
      Else
         Me.Show
         Me.Caption = "Image Preview"
         Label1.Width = Me.Width - 300
         Label1.Caption = "Unable to Display Image: " + Form2.ListView2.ListItems(Form2.ListView2.SelectedItem.Index).Text
         Label1.Visible = True
      End If
   End If
   Set picPreview.Picture = LoadPicture()
   
   blPreviewOpen = False
   
   subErrLog ("Form6:LoadSelectedFilePicture")
End Sub

'Private Sub Form_Click()
'   Load Me
'End Sub

Private Sub Form_Load()
On Error GoTo errlog
   If blForm6Load = False Then
      Unload Me
      Exit Sub
   End If
   If Form1.mSaveSize.Checked = True Then
      If LenB(GetSetting("TaskTracker", "Settings", "Form6Width")) > 0 Then
         Form6.Width = Val(GetSetting("TaskTracker", "Settings", "form6Width"))
      Else
         Form6.Width = 4000
      End If
      
      If LenB(GetSetting("TaskTracker", "Settings", "Form6Height")) > 0 Then
         Form6.Height = Val(GetSetting("TaskTracker", "Settings", "form6Height"))
      Else
         Form6.Height = 4000
      End If
      
      If LenB(GetSetting("TaskTracker", "Settings", "Form6Top")) > 0 Then
         Form6.Top = Val(GetSetting("TaskTracker", "Settings", "form6Top"))
'         Form6.mSaveSize.Checked = True
      Else
         StandardPos
         Exit Sub
      End If
      
      If Len(GetSetting("TaskTracker", "Settings", "Form6Left")) > 0 Then
         Form6.Left = Val(GetSetting("TaskTracker", "Settings", "form6Left"))
      End If
      
   Else
      If blLoading = False Then
         If blForm6Loaded = False Then
            StandardPos
            blForm6Loaded = True
         End If
      End If
   End If
errlog:
   subErrLog ("Form6:Form_Load")
End Sub

Private Sub StandardPos()
On Error GoTo errlog
   If Form2.Left > Form6.Width Then
      Form6.Left = Form2.Left - Form6.Width
'         Form6.Top = (-(Form2.Top - Form2.Height)) + Form6.Height    'pos diff + height = bottom
      Form6.Top = (Form2.Top + Form2.Height) - Form6.Height
    Else
      If Form2.Top < Form6.Height Then
         Form6.Top = Form2.Top + Form2.Height
         Form6.Left = (Form2.Left + Form2.Width) - Form6.Width
      Else
        Form6.Top = Form2.Top - Form6.Height
        Form6.Left = (Form2.Left + Form2.Width) - Form6.Width
      End If
    End If
errlog:
   subErrLog ("Form6:StandardPos")
End Sub

Private Sub Form_Resize()
On Error GoTo errlog
   StatusBar1.Panels(1).Width = StatusBar1.Width
   Call CreateThumbPic(picLoad, picPreview)  'redraw image to size of form up to 100% of pic size
errlog:
   subErrLog ("Form6:Form_Resize")
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errlog
   If blForm6Load = True Then
      Cancel = 1  'prevent unload
      Set picPreview.Picture = LoadPicture()
      If Form2.mPreview.Checked = True Then
         Form2.mPreview.Checked = False
      End If
      Form6.Hide
      blPreviewOpen = False
   End If
errlog:
   subErrLog ("Form6:Form_unload")
End Sub

Public Sub Timer1_Timer()
On Error GoTo errlog
   Dim fp As String
   fp = Form2.fnGetPath
   If Len(fp) > 0 Then
      LoadSelectedFilePicture (fp)
   Else
      Me.Hide
   End If
   Timer1.Enabled = False
errlog:
   subErrLog ("Form6:Timer1_Timer")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errlog
   If KeyCode = vbKeyEscape Then
      Form_Unload (0)
   End If
errlog:
   subErrLog ("Form6:Form_KeyDown")
End Sub

Private Sub KeepForm6Visible()
   If Me.Left > Form2.Left And Me.Left < Form2.Left + Form2.Width Then
      If Me.Top > Form2.Top And Me.Top < Form2.Top + Form2.Height Then
         Me.ZOrder
         Exit Sub
      End If
   End If
   Form2.SetFocus
End Sub

