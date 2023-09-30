Attribute VB_Name = "modActivation"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ©  2004 Michael M. Ross, Wordwise Solutions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
'
Public Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" _
  (ByVal hwnd As Long, _
   ByVal nIndex As Long, _
   ByVal wNewWord As Long) As Long
 
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long) As Long
 
Public Declare Function SetLayeredWindowAttributes Lib "user32" _
 (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, _
 ByVal dwFlags As Long) As Long

Public Const WS_EX_LAYERED = &H80000
Public Const LWA_ALPHA = &H2&

Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_APPWINDOW = &H40000

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Public blNoFade As Boolean
Public blFadeCheck As Boolean
Private blLastActive As Boolean

Public Function IsTTActive() As Boolean
On Error GoTo errlog

    Dim lRet As Long, lhWnd As Long
    Dim sWindowText As String
    
    'Get the window handle of the foreground window
    lhWnd = GetForegroundWindow()
    If (lhWnd <> 0) Then
        'Get the window caption
        sWindowText = Space(255)
        lRet = GetWindowText(lhWnd, sWindowText, 255)
        sWindowText = Left(sWindowText, lRet)
        
        If blFadeCheck = True Then
'           Debug.Print "fadecheck"
           blFadeCheck = False
           If Left$(sWindowText, 11) = "TaskTracker" Then      'Left$(sWindowText, 11) because form2 has dynamic title
               IsTTActive = True
               blLastActive = True
           End If
           Exit Function
        Else
           If Left$(sWindowText, 11) = "TaskTracker" Then
               blLastActive = True
           End If
        End If
        
        If Left$(sWindowText, 11) <> "TaskTracker" Then
            If blLastActive = False Then
               Exit Function
            End If
            If blNoFade = False Then
               subFadeOut
            End If
            blLastActive = False
        Else
           If blLastActive = True Then
               Exit Function
           End If
           If blNoFade = True Then
               subFadeIn
            End If
            blLastActive = True
        End If
    End If
errlog:
   subErrLog ("modActivation: IsTTActive")
End Function

Public Sub subFadeIn()
   On Error Resume Next
   
   Dim lOldStyle1 As Long, lOldStyle2 As Long, lOldStyle4 As Long, lOldStyle5 As Long, lOldStyle6 As Long
   Dim bTrans As Integer ' The level of transparency (0 - 255)

   lOldStyle1 = GetWindowLong(Form1.hwnd, GWL_EXSTYLE)
   lOldStyle2 = GetWindowLong(Form2.hwnd, GWL_EXSTYLE)
   lOldStyle4 = GetWindowLong(Form4.hwnd, GWL_EXSTYLE)
   lOldStyle5 = GetWindowLong(Form5.hwnd, GWL_EXSTYLE)
   lOldStyle6 = GetWindowLong(Form6.hwnd, GWL_EXSTYLE)
   SetWindowLong Form1.hwnd, GWL_EXSTYLE, lOldStyle1 Or WS_EX_LAYERED
   SetWindowLong Form2.hwnd, GWL_EXSTYLE, lOldStyle2 Or WS_EX_LAYERED
   SetWindowLong Form4.hwnd, GWL_EXSTYLE, lOldStyle4 Or WS_EX_LAYERED
   SetWindowLong Form5.hwnd, GWL_EXSTYLE, lOldStyle5 Or WS_EX_LAYERED
   SetWindowLong Form6.hwnd, GWL_EXSTYLE, lOldStyle6 Or WS_EX_LAYERED
   For bTrans = 0 To 255 Step 5
       SetLayeredWindowAttributes Form1.hwnd, 0, bTrans, LWA_ALPHA
       SetLayeredWindowAttributes Form2.hwnd, 0, bTrans, LWA_ALPHA
       SetLayeredWindowAttributes Form4.hwnd, 0, bTrans, LWA_ALPHA
       SetLayeredWindowAttributes Form5.hwnd, 0, bTrans, LWA_ALPHA
       SetLayeredWindowAttributes Form6.hwnd, 0, bTrans, LWA_ALPHA
       DoEvents
   Next bTrans
   blNoFade = False
   
errlog:
   subErrLog ("modActivation: subFadeIn")
End Sub

Private Sub subFadeOut()
   On Error Resume Next
   
   Dim lOldStyle1 As Long, lOldStyle2 As Long, lOldStyle4 As Long, lOldStyle5 As Long, lOldStyle6 As Long
   Dim bTrans As Integer ' The level of transparency (0 - 255)

   lOldStyle1 = GetWindowLong(Form1.hwnd, GWL_EXSTYLE)
   lOldStyle2 = GetWindowLong(Form2.hwnd, GWL_EXSTYLE)
   lOldStyle4 = GetWindowLong(Form4.hwnd, GWL_EXSTYLE)
   lOldStyle5 = GetWindowLong(Form5.hwnd, GWL_EXSTYLE)
   lOldStyle6 = GetWindowLong(Form6.hwnd, GWL_EXSTYLE)
   SetWindowLong Form1.hwnd, GWL_EXSTYLE, lOldStyle1 Or WS_EX_LAYERED
   SetWindowLong Form2.hwnd, GWL_EXSTYLE, lOldStyle2 Or WS_EX_LAYERED
   SetWindowLong Form4.hwnd, GWL_EXSTYLE, lOldStyle4 Or WS_EX_LAYERED
   SetWindowLong Form5.hwnd, GWL_EXSTYLE, lOldStyle5 Or WS_EX_LAYERED
   SetWindowLong Form6.hwnd, GWL_EXSTYLE, lOldStyle6 Or WS_EX_LAYERED
   For bTrans = 255 To 0 Step -1
       SetLayeredWindowAttributes Form1.hwnd, 0, bTrans, LWA_ALPHA
       SetLayeredWindowAttributes Form2.hwnd, 0, bTrans, LWA_ALPHA
       SetLayeredWindowAttributes Form4.hwnd, 0, bTrans, LWA_ALPHA
       SetLayeredWindowAttributes Form5.hwnd, 0, bTrans, LWA_ALPHA
       SetLayeredWindowAttributes Form6.hwnd, 0, bTrans, LWA_ALPHA
       If bTrans Mod 5 = 0 Then
          blFadeCheck = True
          If IsTTActive = True Then
            subRestore
            Exit Sub
          End If
       End If
       If GetInputState() <> 0 Then
          DoEvents
       End If
   Next bTrans
   
   Form1.Visible = False
   Form1.SysTrayStatus
   Form2.Visible = False
   Form4.Visible = False
   Form5.Visible = False
   Form6.Visible = False
   
errlog:
   subErrLog ("modActivation: subFadeOut")
End Sub

Public Sub subRestore()
   On Error Resume Next
   
   Dim lOldStyle1 As Long, lOldStyle2 As Long, lOldStyle4 As Long, lOldStyle5 As Long, lOldStyle6 As Long

   lOldStyle1 = GetWindowLong(Form1.hwnd, GWL_EXSTYLE)
   lOldStyle2 = GetWindowLong(Form2.hwnd, GWL_EXSTYLE)
   SetWindowLong Form1.hwnd, GWL_EXSTYLE, lOldStyle1 Or WS_EX_LAYERED
   SetWindowLong Form2.hwnd, GWL_EXSTYLE, lOldStyle2 Or WS_EX_LAYERED
   SetLayeredWindowAttributes Form1.hwnd, 0, 255, LWA_ALPHA
   SetLayeredWindowAttributes Form2.hwnd, 0, 255, LWA_ALPHA
   
   lOldStyle4 = GetWindowLong(Form4.hwnd, GWL_EXSTYLE)
   lOldStyle5 = GetWindowLong(Form5.hwnd, GWL_EXSTYLE)
   lOldStyle6 = GetWindowLong(Form6.hwnd, GWL_EXSTYLE)
   SetWindowLong Form4.hwnd, GWL_EXSTYLE, lOldStyle4 Or WS_EX_LAYERED
   SetWindowLong Form5.hwnd, GWL_EXSTYLE, lOldStyle5 Or WS_EX_LAYERED
   SetWindowLong Form6.hwnd, GWL_EXSTYLE, lOldStyle6 Or WS_EX_LAYERED
   SetLayeredWindowAttributes Form4.hwnd, 0, 255, LWA_ALPHA
   SetLayeredWindowAttributes Form5.hwnd, 0, 255, LWA_ALPHA
   SetLayeredWindowAttributes Form6.hwnd, 0, 255, LWA_ALPHA
   
errlog:
   blLastActive = True
   subErrLog ("modActivation: subRestore")
End Sub
