VERSION 5.00
Begin VB.Form frmNotify 
   Caption         =   "Form3"
   ClientHeight    =   3692
   ClientLeft      =   104
   ClientTop       =   416
   ClientWidth     =   7579
   LinkTopic       =   "Form3"
   ScaleHeight     =   3692
   ScaleWidth      =   7579
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ListBox List1 
      Height          =   2587
      ItemData        =   "frmNotify.frx":0000
      Left            =   585
      List            =   "frmNotify.frx":0002
      TabIndex        =   0
      Top             =   234
      Width           =   6682
   End
   Begin VB.Timer tmNotify 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   0
   End
   Begin VB.Label txt 
      Caption         =   "Label1"
      Height          =   481
      Left            =   585
      TabIndex        =   1
      Top             =   3042
      Width           =   2821
   End
End
Attribute VB_Name = "frmNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Declare Function PathIsFileSpec Lib "shlwapi" _
'   Alias "PathIsFileSpecA" _
'  (ByVal pszPath As String) As Long
         
Private Sub Form_Load()

   If SubClass(hWnd) Then

'      If IsIDE Then

'      Text1.Text = "**IMPORTANT**" & vbCrLf & _
'              "This window is subclassed. Do not close it from" & vbCrLf & _
'              "either VB's End button or End menu command," & vbCrLf & _
'              "or VB will blow up. Close this window only from" & vbCrLf & _
'              "the system menu above!" & vbCrLf & vbCrLf & Text1
'      End If

      Call SHNotify_Register(hWnd)

'   Else
'      Text1.Text = "Well, it is supposed to work."
   End If
   Me.Show
End Sub


'Private Sub Form_Resize()
'
'  On Error GoTo Out
'  Text1.Move 0, 0, ScaleWidth, ScaleHeight
'
'Out:
'End Sub


Private Sub Form_Unload(Cancel As Integer)

   Call SHNotify_Unregister
   Call UnSubClass(hWnd)

End Sub


Private Function IsIDE() As Boolean

   On Error GoTo Out
'   Debug.Print 1 / 0
  
Out:
   IsIDE = Err
End Function


Public Sub NotificationReceipt(wParam As Long, lParam As Long)

   Dim sOut As String
   Dim shns As SHNOTIFYSTRUCT
   
   sOut = SHNotify_GetEventStr(lParam) & vbCrLf
   
'   List1.AddItem "abc"
   
  'Fill the SHNOTIFYSTRUCT from its pointer.
   CopyMemory shns, ByVal wParam, Len(shns)
       
  'lParam is the ID of the notification event,
  'one of the SHCN_EventIDs.
   Select Case lParam
      
     '----------------------------------------------------
     'For the SHCNE_FREESPACE event, dwItem1 points
     'to what looks like a 10 byte struct. The first
     'two bytes are the size of the struct, and the
     'next two members equate to SHChangeNotify's
     'dwItem1 and dwItem2 params.
    
     'The dwItem1 member is a bitfield indicating which
     'drive(s) had its (their) free space changed.
     'The bitfield is identical to the bitfield returned
     'from a GetLogicalDrives call, i.e., bit 0 = A:\, bit
     '1 = B:\, 2, = C:\, etc. Since VB does DWORD alignment
     'when CopyMemory'ing to a struct, we'll extract the
     'bitfield directly from its memory location.
    
      Case SHCNE_FREESPACE
      
         Dim dwDriveBits As Long
         Dim wHighBit As Integer
         Dim wBit As Integer
         
         CopyMemory dwDriveBits, ByVal shns.dwItem1 + 2, 4
   
        'Get the zero based position of the highest
        'bit set in the bitmask (essentially determining
        'the value's highest complete power of 2).
        'Use floating point division (we want the exact
        'values from the Logs) and remove the fractional
        'value (the fraction indicates the value of
        'the last incomplete power of 2, which means the
        'bit isn't set).
        
         wHighBit = Int(Log(dwDriveBits) / Log(2))
         
         For wBit = 0 To wHighBit
           
          'If the bit is set...
           If (2 ^ wBit) And dwDriveBits Then
             
            '... get its drive string
             sOut = sOut & Chr$(vbKeyA + wBit) & ":\" & vbCrLf
   
           End If
         Next
      
     '----------------------------------------------------
     'shns.dwItem1 also points to a 10 byte struct. The
     'struct's second member (after the struct's first
     'WORD size member) points to the system imagelist
     'index of the image that was updated.
      Case SHCNE_UPDATEIMAGE
      
         Dim iImage As Long
      
         CopyMemory iImage, ByVal shns.dwItem1 + 2, 4
'         sOut = iImage & vbCrLf
    
     '----------------------------------------------------
     'Everything else except SHCNE_ATTRIBUTES is the
     'pidl(s) of the changed item(s). For SHCNE_ATTRIBUTES,
     'neither item is used. See the description of the
     'values for the wEventId parameter of the
     'SHChangeNotify API function for more info.
      Case Else
         Dim sDisplayname As String
         
         If shns.dwItem1 Then
         
            sDisplayname = GetDisplayNameFromPIDL(shns.dwItem1)
            
            If Len(sDisplayname) Then
'             sOut = sDisplayname & vbCrLf
             sOut = GetPathFromPIDL(shns.dwItem1) & vbCrLf
'            Debug.Print "1 " + sOut
            Else
             Exit Sub
            End If
            
         End If
         
         If shns.dwItem2 Then
         
            sDisplayname = GetDisplayNameFromPIDL(shns.dwItem2)
           
            If Len(sDisplayname) Then
'               sOut = sOut & "second item displayname: " & sDisplayname & vbCrLf
               sOut = GetPathFromPIDL(shns.dwItem2) & vbCrLf
'            Debug.Print "2 " + sOut
            Else
               Exit Sub
            End If
         End If
  
  End Select
  
  If InStr(sOut, "ntuser.dat") > 0 Then Exit Sub
  If InStr(sOut, ".lnk") > 0 Then Exit Sub
  If InStr(sOut, "\Desktop") > 0 Then Exit Sub
  If InStr(sOut, "FREESPACE") > 0 Then Exit Sub
  If InStr(sOut, ".tmp") > 0 Then Exit Sub
  If InStr(sOut, "TaskTracker.log") Then Exit Sub

'  Debug.Print "1 " + sOut
  
'  If DoesFileExist(sOut) Then

'    Debug.Print "a " + sOut
'    Call subAddShortcut(sOut)
'  End If
  List1.AddItem sOut

End Sub

Public Sub CheckNotifyList()
On Error GoTo errlog
    Dim sfile As String
    Dim lsts As Integer
    With List1
        If .ListCount = 0 Then Exit Sub
        For lsts = 1 To .ListCount
'           DoEvents
           sfile = Trim$(.List(lsts))
           .RemoveItem (lsts)
           If Len(Trim$(sfile)) > 0 Then
'              If InStr(sfile, "\") > 0 And InStr(sfile, ".") > 0 Then
              sfile = GetStrFromBufferA(sfile)
'              If FileExists(sfile) Then
'                Debug.Print sfile
                 blNotified = True
'                Call ShowFileProperties(sfile, 0, True)
                 Call subAddShortcut(sfile, 0, 1)
                 blNotified = False
'              End If
           End If
        Next lsts
    End With
errlog:
    blNotified = False
'    With List1
'        If Len(.List(lsts)) > 0 Then
''            .RemoveItem (lsts)
'        End If
'    End With
End Sub


'Private Sub tmNotify_Timer()

  'initial settings: Interval = 1, Enabled = False
  
'   Static nCount As Integer
'
'   If nCount = 0 Then tmNotify.Interval = 200
'
'   nCount = nCount + 1
'   Call FlashWindow(hWnd, True)
   
  'Reset everything after 3 flash cycles
'   If nCount = 6 Then
'      nCount = 0
'      tmNotify.Interval = 1
'      tmNotify = False
'   End If

'End Sub

'Public Function IsPathAFile(ByVal sPath As String) As Boolean
'
'  'given a path or file, determines
'  'if the string is a file or path.
'
'  'The function searches a path for
'  'any path delimiting characters
'  '(for example, ':' or '\' ). If
'  'there are no path delimiting
'  'characters present, the path is
'  'considered to be a File Spec path.
'
'  IsPathAFile = PathIsFileSpec(sPath) = 1
'
'End Function

'Public Function FileExists(sSource As String) As Boolean
'On Error GoTo errlog
'
'   Dim lFile As Long
'
'   lFile = CreateFile(sSource, 0&, FILE_SHARE_READ, 0&, _
'                      OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0&)
'
'   If lFile <> 0 Then
'      FileExists = True
'   End If
'
'   Call CloseHandle(lFile)
'
'errlog:
'subErrLog ("modFileFuncs: DoesFileExist")
'End Function

Public Sub CreateLink(sfile As String, dfile As String)

   Dim SHFileOp As SHFILEOPSTRUCT
   Dim sSource As String, sDestination As String
  
  'terminate passed strings with a null
   sSource = sfile & Chr$(0)
   sDestination = dfile & Chr$(0)
  
  'set up the options
   With SHFileOp
     .wFunc = FO_COPY
     .pFrom = sSource
     .pTo = sDestination
     .fFlags = FOF_SILENT Or FOF_NOCONFIRMATION
   End With
   
  'and perform the copy
   Call SHFileOperation(SHFileOp)
  
End Sub




