Attribute VB_Name = "modShortcut"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ©  2003-2007 Michael M. Ross, Wordwise Solutions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
  
Public Const SHARD_PATH = &H2&
Private Const ERROR_SUCCESS As Long = 0
Private Const vbDot As Long = 46 ' Asc(".") = 46
Private Const SHGFI_USEFILEATTRIBUTES As Long = &H10
Private Const SHGFI_TYPENAME As Long = &H400

Public Declare Function SHAddToRecentDocs Lib "shell32" _
  (ByVal dwFlags As Long, _
   ByVal dwData As String) As Long
          
Private Declare Function RegEnumKeyEx Lib "advapi32" _
   Alias "RegEnumKeyExA" _
   (ByVal hKey As Long, _
   ByVal dwIndex As Long, _
   ByVal lpName As String, _
   lpcbName As Long, _
   ByVal lpReserved As Long, _
   ByVal lpClass As String, _
   lpcbClass As Long, _
   lpftLastWriteTime As FILETIME) As Long
   
Private Declare Function FindExecutable Lib "shell32" _
   Alias "FindExecutableA" _
  (ByVal lpFile As String, _
   ByVal lpDirectory As String, _
   ByVal sresult As String) As Long

'Private Const MAX_PATH As Long = 260
Private Const ERROR_FILE_NO_ASSOCIATION As Long = 31

Public blform2 As Boolean
      
Private blOnce1 As Boolean
Private blOnce2 As Boolean
Private blOnce3 As Boolean
Private blOnce4 As Boolean
Private blOnce5 As Boolean

Private iDFiles As Integer
Private DragErr() As Integer
Private kn As Integer
Private et As Integer
Public lsc As Integer

Private DragErrFile() As String
Private ExcludeType() As String

Public sRegType As String
Public KeepRegType As String
Public SelectThisFile As String
Public SelectThisType As String
Public blFindSelected As Boolean

Public KeepTheName() As String
Public DroppedName As String

Public Sub subAddShortcut(dfile As String, Optional dcnt As Integer, Optional dfiles As Integer)
On Error GoTo Cleanup
'Debug.Print dfile
'Debug.Print Str(dcnt)
'Debug.Print Str(dfiles)

   If blLoading = True Then
      Exit Sub
   End If
      
   Dim fNameOnly As String
   Dim frootnname As String
   Dim lp As Integer
   Dim cc As Integer
   Dim blcreatednew As Boolean
   Dim blRegd As Boolean
   
   'Determine filename only
   For lp = 1 To Len(dfile)
       If Mid$(dfile, lp, 1) = "\" Then
          cc = cc + 1
       End If
   Next lp
   If InStr(dfile, "\\") = 0 Then
      If cc > 1 Then
         fNameOnly = Right$(dfile, Len(dfile) - InStrRev(dfile, "\"))
         frootnname = Left$(dfile, 3) + "...\" + fNameOnly
      Else
         frootnname = dfile
      End If
   Else
      If cc > 2 Then
         fNameOnly = Right$(dfile, Len(dfile) - InStrRev(dfile, "\"))
         frootnname = Left$(dfile, 3) + "...\" + fNameOnly
      Else
         frootnname = dfile
      End If
   End If
   
   If blNotified Then
'        Call CreateNewShortcut(blRegd, dfile, fNameOnly, frootnname, dcnt)
        Call frmNotify.CreateLink(dfile, sRecent + fNameOnly + ".lnk")
        Exit Sub
    End If

   'Immediate Error msgs
'   If blMultFileDrop = False Then
'      If LCase$(Right$(dfile, 4)) = ".lnk" Or LCase$(Right$(dfile, 4)) = ".url" Then
'         Form2.Form_Click
'         MsgBox "Drag files, not shortcuts, into TaskTracker." _
'         + vbNewLine + vbNewLine + "(" + frootnname + ")", vbExclamation
'         GoTo Cleanup
'      End If
      
'  Handle duplicate file names
'  ******************************************************
   Dim knc As Integer
   Dim blRenameTheShortcut As Boolean
   ReDim Preserve KeepTheName(kn)
   For knc = 0 To UBound(KeepTheName)
      If frootnname = KeepTheName(knc) Then
         blRenameTheShortcut = True
      End If
   Next knc
   KeepTheName(kn) = frootnname
   kn = kn + 1
   If kn = dfiles Then
      Erase KeepTheName
      kn = 0
   End If
'     ******************************************************
      
''   Errors for Summary
'   Else
   If LCase$(Right$(dfile, 4)) = ".lnk" Or LCase$(Right$(dfile, 4)) = ".url" Then
      ReDim Preserve DragErr(dcnt)
      ReDim Preserve DragErrFile(dcnt)
      DragErr(dcnt) = 1
      DragErrFile(dcnt) = frootnname
      Exit Sub
   End If
            
'   If blNotified Then
'       Call SyncShortcut(dfile, fNameOnly, "")
'       Debug.Print dfile
'       Exit Sub
'   End If
      
   ReDim Preserve blDontCopy(dcnt)
   Dim lresult As Long
   Dim shFlag As Long
   
   shFlag = SHARD_PATH
   
   'Create the Recent shortcut
'   lresult = SHAddToRecentDocs(shFlag, dfile)
   
'   If blNotified Then
'       Call SyncShortcut(dfile, fNameOnly, "")
'       Exit Sub
'   End If
   
   If GetInputState() <> 0 Then
       DoEvents
   End If
   
'   Form2.ApplyHeaderIcon
'   If blMultFileDrop = False Then
'      If Right$(dfile, 2) = ":\" Then
'         Form2.Form_Click
'         MsgBox "The root of a shared drive cannot be tracked. Drag a shared folder instead." _
'         + vbNewLine + vbNewLine + "(" + frootnname + ")", vbExclamation
'         blDontCopy(dcnt) = False   '!actually, it's True, but that would cause fntrackfiles to say it's being tracked!
'         dcnt = dcnt + 1
'         GoTo Cleanup
'      End If
'   Else
   If Right$(dfile, 2) = ":\" Then
      ReDim Preserve DragErr(dcnt)
      ReDim Preserve DragErrFile(dcnt)
      DragErr(dcnt) = 2
      DragErrFile(dcnt) = frootnname
      GoTo Cleanup
   End If

   If IsRegisteredType(dfile) = True Then
      blRegd = True
   Else
      blRegd = False
   End If
   
   If CompareToMemArray(dfile) = False Then
   
      blDontCopy(dcnt) = True
      dcnt = dcnt + 1
      
   Else

      With Form1
         .ProgressBar1.Value = dcnt
         .ProgressBar1.ToolTipText = Str$(Trim$((dcnt))) + " files processed"
         .Refresh
         
         If CreateNewShortcut(blRegd, dfile, fNameOnly, frootnname, dcnt) = False Then
            If blMultFileDrop = True Then
               .StatusBar1.Panels(1) = "Adding" + Str$(TrimNull((dfiles - dcnt) - 1)) + " files... " _
               + Str$(Trim$(Int(dcnt / dfiles * 100))) + "%"
            End If
            sRegType = vbNullString
            GoTo Cleanup
         Else
            blcreatednew = True
         End If
         
   '     continue if blRegd = True Or CreateNewShortcut = True
         
         If blMultFileDrop = True Then
            .StatusBar1.Panels(1) = "Adding" + Str$(TrimNull(dfiles)) + " files... " _
            + Str$(Trim$(Int(dcnt / dfiles * 100))) + "%"
         End If
         If Len(sRegType) > 0 Then
            KeepRegType = sRegType
   '         Form1.ListView1.ListItems(sRegType).Selected = True
         End If
         
      End With
      
      blDontCopy(dcnt) = False
      dcnt = dcnt + 1
      If blRenameTheShortcut = True Then
         Form1.subCopyShortcuts (sRecent & "*.lnk")
'         blMultFileDrop = False
      End If
      
   End If
   
'   If blNotified Then Exit Sub
             
   If blcreatednew = True Then
   
      Dim SHFileOp As SHFILEOPSTRUCT
      Dim lFlags   As Long
      
      lFlags = lFlags Or FOF_SILENT
      lFlags = lFlags Or FOF_NOCONFIRMATION
      
      'Delete in \Recent folder
      If DoesFileExist(sRecent + fNameOnly + ".lnk") = True Then
         With SHFileOp
             .wFunc = FO_DELETE
             .pFrom = sRecent + fNameOnly + ".lnk" & vbNullChar & vbNullChar
             .fFlags = lFlags
         End With
         lresult = SHFileOperation(SHFileOp)
      End If
      
      If GetInputState() <> 0 Then
          DoEvents
      End If
      
      KeepRegType = vbNullString
      
  End If

  Exit Sub

Cleanup:
'   If blNotified = False Then
'      Form1.subShortStart
'   End If
   subErrLog ("modShortcut:subAddShortcut")
End Sub

Public Function IsRegisteredType(dfile As String) As Boolean
On Error GoTo errlog

   Dim dwIndex As Long
   Dim sSubkey As String * MAX_PATH
   Dim sClass As String * MAX_PATH
   Dim ft As FILETIME
   Dim GFT As String
   
'   Form2.ApplyHeaderIcon
   
   IsRegisteredType = False
      
   If DoesFolderExist(dfile) = True Then
      IsRegisteredType = True
      sRegType = "File Folder"
      Exit Function
   End If
   
   GFT = GetFileType(dfile)
   
   Do While RegEnumKeyEx(HKEY_CLASSES_ROOT, _
                         dwIndex, _
                         sSubkey, _
                         MAX_PATH, _
                         0, sClass, _
                         MAX_PATH, ft) = ERROR_SUCCESS
      
      If Asc(sSubkey) = vbDot Then
                
       'Pass the returned string to get the file type
         If GFT = GetFileType(sSubkey) Then
            
            sRegType = GFT
            IsRegisteredType = True
            Exit Do
            
         End If
            
      End If

      dwIndex = dwIndex + 1
   
   Loop

errlog:
   subErrLog ("modShortcut:IsRegisteredType")
End Function

Private Function GetFileType(dfile As String) As String
On Error GoTo errlog

   Dim sfi As SHFILEINFO
   If SHGetFileInfo(dfile, 0&, _
                    sfi, Len(sfi), _
                    SHGFI_TYPENAME Or SHGFI_USEFILEATTRIBUTES) Then
      GetFileType = TrimNull(sfi.szTypeName)
   End If

errlog:
   subErrLog ("modShortcut:GetFileType")
End Function

Public Sub ShowTypesFiles()
On Error GoTo errlog

   If Form2.Visible = True Then
      If Len(sRegType) > 0 Then
         Form1.ListView1.ListItems(sRegType).Selected = True
      End If
      Form1.ListView1_Click
   End If
   
errlog:
   subErrLog ("modShortcut:ShowTypesFiles")
End Sub

Public Sub TrackStatus(dfcount As Integer)
On Error GoTo errlog

   Dim dc As Integer, Dudfiles As Integer
   If dfcount = 1 Then
'      Form2.ApplyHeaderIcon
      If Len(SelectThisType) > 0 Then
         sType = SelectThisType
      End If
      Form1.PopsType
      ShowTypesFiles
'      LeaveSelected
      
      If UBound(blDontCopy()) = 0 And blDontCopy(0) = False Then
         'nothing
      Else
         If blFindSelected = True Then
            MsgBox "TaskTracker is already tracking this file.", vbInformation
         Else
            If blDontCopy(0) = True Then
               Dim Response As Integer
               Response = MsgBox("TaskTracker may already be tracking this file but needs reloading for you to see it." + vbNewLine _
               + "Do you want to reload TaskTracker? (This can take a long time.)", 292)
               If Response = vbYes Then
                  Form1.subReload
               End If
            End If
         End If
      End If
      
   Else
      For dc = 0 To UBound(blDontCopy)
         If blDontCopy(dc) = True Then
            Dudfiles = Dudfiles + 1
         End If
      Next dc
      
      Form1.PopsType
      ShowTypesFiles
         
      If Dudfiles = dfcount Then
   '      Form2.ApplyHeaderIcon
         MsgBox "TaskTracker is already tracking these files.", vbInformation
         Exit Sub
      ElseIf Dudfiles > 0 Then
         If blMultFileDrop = True Then
            MsgBox "TaskTracker is already tracking some of these files.", vbInformation
         End If
      End If
   End If
   
errlog:
   sRegType = vbNullString
   Form1.Timer1.Enabled = True
   SelectThisFile = vbNullString
   SelectThisType = vbNullString
   subErrLog ("modShortcut:TrackStatus")
End Sub

Private Function CompareToMemArray(dfile As String) As Boolean
On Error GoTo errlog
   Dim cm As Long

   With Form1.ListView1
      CompareToMemArray = True
      For cm = 0 To UBound(fType()) - 1
         If LCase(dfile) = LCase(aSFileName(cm)) Then
            If blExists(cm) = True Then
               If fType(cm) <> "Deleted or Renamed" Then
                  CompareToMemArray = False     'file exists, don't copy
'                  If blMultFileDrop = False Then
                     SelectThisType = GetFileType(dfile)    'ensure right file type - cached can be wrong
                     SelectThisFile = aShortcut(cm)      'this selects existing file if present already
'                  End If
                  Exit Function
               End If
            End If
         End If
      Next cm
   End With
   
   'This is for dropped files that are new only
   TTCnt = TTCnt + 1
   ReDim Preserve fType(TTCnt)
   ReDim Preserve blExists(TTCnt)
   SelectThisType = GetFileType(dfile)    'ensure right file type - cached can be wrong
   DroppedName = dfile
           
errlog:
   subErrLog ("modShortcut:CompareToMemArray")
End Function

Public Function InitDrag(dfiles As Integer) As Boolean
On Error GoTo errlog
   blMultFileDrop = False
   blOnce1 = False
   blOnce2 = False
   blOnce3 = False
   blOnce4 = False
   blOnce5 = False
   If blDragOut = True Then
      blDragOut = False
      InitDrag = False
      Exit Function
   End If
   If dfiles > 1 Then
      blMultFileDrop = True
         With Form1
            If .mIcons.Checked = False Then
               .Label1.Visible = False
               .ProgressBar1.Max = dfiles
               .ProgressBar1.Value = 0
               .ProgressBar1.Visible = True
               .Refresh
               .StatusBar1.Panels(1) = "Preparing to add " + Trim$(Str$(dfiles)) + " files..."
            End If
         End With
      InitDrag = True
      iDFiles = dfiles
   End If
   InitDrag = True
'   blDragDrop = True
   Form2.SetFocus
errlog:
   subErrLog ("modShortcut:InitDrag")
End Function

Public Sub SummaryMessage()
On Error Resume Next
   Dim sm As Integer
   Dim str1 As String, str2 As String
   Dim msg1 As String, msg2 As String
   Dim msgstr As String
   
   For sm = 0 To iDFiles
      If DragErr(sm) = 1 Then
         str1 = str1 + ", " + DragErrFile(sm)
      ElseIf DragErr(sm) = 2 Then
         str2 = str2 + ", " + DragErrFile(sm)
      End If
   Next sm
   
   str1 = Right$(str1, Len(str1) - 1)
   str2 = Right$(str2, Len(str2) - 1)
   
   If Len(str1) > 0 Then
      msg1 = "Drag files, not shortcuts, into TaskTracker." + vbNewLine + Chr$(13) + "(" + Trim$(str1) + ")" + Chr$(13) + Chr$(13)
   End If
   If Len(str2) > 0 Then
      msg2 = "The root of a shared drive cannot be tracked. Drag a shared folder instead." + vbNewLine + Chr$(13) + "(" + Trim$(str2) + ")" + Chr$(13) + Chr$(13)
   End If
   
   msgstr = msg1 + msg2
   
   If Len(msgstr) > 0 Then
   
      Form2.Form_Click
      MsgBox msgstr, vbExclamation
      
   End If
   
errlog:
   Erase DragErr
   Erase DragErrFile
   subErrLog ("modShortcut:SummaryMessage")
End Sub

Private Function CreateNewShortcut(ByVal blRegd As String, ByVal dfile As String, ByVal fNameOnly As String, _
   ByVal frootnname As String, ByVal dcnt As Integer) As Boolean
   
On Error GoTo errlog

   Dim sresult As String
   Dim fpathonly As String
   Dim blNewType As Boolean
   Dim ft As Long
   Dim Response As Integer
   Dim SetReg As String
   Dim etc As Integer
   
   Form2.Form_Click
   
   blNewType = True
   
   'If not a registered type, take the raw extension
   If blRegd = False Then
      sRegType = UCase$(Right$(frootnname, Len(frootnname) - InStrRev(frootnname, "."))) + " File"
   End If
      
   'if ftype already exists, not a new type
   If Len(sRegType) > 0 Then
      For ft = 1 To UBound(fType)
        If sRegType = fType(ft) Then
           blNewType = False
           Exit For
        End If
      Next
   End If
   
   'Exclude rejected types, unless individually added
   If blNewType = True Then
      If blMultFileDrop = True Then
         If et > 0 Then
            For etc = 0 To UBound(ExcludeType)
               If sRegType = ExcludeType(etc) Then
                  blDontCopy(dcnt) = False
                  dcnt = dcnt + 1
                  CreateNewShortcut = False
                  Exit Function
               End If
            Next etc
         End If
      End If
   End If

   If blNotified = False Then
        Select Case sRegType
        
           Case ""
           
              If blNewType = True Then
                 On Error Resume Next
                    If DragErr(dcnt) = 0 Then
                       If blOnce1 = False Then
                          Response = MsgBox("This file is not a registered file type on your system." _
                             + vbNewLine + vbNewLine + frootnname _
                             + vbNewLine + vbNewLine + "You can add this file to TaskTracker, but new files of this type will not be tracked automatically." _
                             + vbNewLine + "Do you want to continue?", 52)
                          blOnce1 = True
                          SetReg = UCase$(Right$(frootnname, Len(frootnname) - InStrRev(frootnname, "."))) + " File"
                          GoTo Response
                       End If
                    End If
                 On Error GoTo 0
              End If
              
           Case "File Folder"
           
                 ' no message
              
           Case Else
           
              'Exe or Dll - do this way to get round localized names "Application","Application Extension"
               If blRegd = True Then
                 If Right$(frootnname, Len(frootnname) - InStrRev(frootnname, ".")) = "exe" Then
                    If blNewType = True Then
                       If blOnce2 = False Then
                          Response = MsgBox("This file is an application." _
                             + vbNewLine + vbNewLine + frootnname _
                             + vbNewLine + vbNewLine + "You can add this file to TaskTracker, but new files of this type will not be tracked automatically." _
                             + vbNewLine + "Do you want to continue?", 52)
                          blOnce2 = True
                          GoTo Response
                       End If
                    End If
                 ElseIf Right$(frootnname, Len(frootnname) - InStrRev(frootnname, ".")) = "dll" Then
                    If blNewType = True Then
                       If blOnce3 = False Then
                          Response = MsgBox("This file is an application extension." _
                             + vbNewLine + vbNewLine + frootnname _
                             + vbNewLine + vbNewLine + "You can add this file to TaskTracker, but new files of this type will not be tracked automatically." _
                             + vbNewLine + "Do you want to continue?", 52)
                          blOnce3 = True
                          GoTo Response
                       End If
                    End If
                 End If
              End If
                                   
              'No Assocation - reg'd or unreg'd
              If blNewType = True Then
                 sresult = Space$(MAX_PATH)
                 fpathonly = Left$(dfile, Len(dfile) - Len(fNameOnly))
                 If FindExecutable(fNameOnly, fpathonly, sresult) = ERROR_FILE_NO_ASSOCIATION Then
                    If blOnce4 = False Then
                       If blRegd Then
                          Response = MsgBox("This file is registered but not currently associated with an application on your system." _
                             + vbNewLine + vbNewLine + frootnname _
                             + vbNewLine + vbNewLine + "You can add this file to TaskTracker, but new files of this type will not be tracked automatically." _
                             + vbNewLine + "Do you want to continue?", 52)
                       Else
                          Response = MsgBox("This file is not registered or associated with an application on your system." _
                             + vbNewLine + vbNewLine + frootnname _
                             + vbNewLine + vbNewLine + "You can add this file to TaskTracker, but new files of this type will not be tracked automatically." _
                             + vbNewLine + "Do you want to continue?", 52)
                       End If
                       blOnce4 = True
                       GoTo Response
                    End If
                 End If
                 
                 If blOnce5 = False Then
                    If blRegd Then
                       'just add the file - it's registered and associated...
                    Else
                       Response = MsgBox(frootnname _
                          + vbNewLine + vbNewLine + "You can add this file to TaskTracker, but new files of this type will not be tracked automatically." _
                          + vbNewLine + "Do you want to continue?", 52)
                    End If
                 End If
                 blOnce5 = True
                 GoTo Response
              End If
                       
        End Select
    End If
   
Response:
            
   If blNewType = True Then
      If Response = vbNo Then
         blDontCopy(dcnt) = False
         dcnt = dcnt + 1
         CreateNewShortcut = False
         
         Dim ext As String
         If Len(SetReg) > 0 Then
            ext = SetReg
         Else
            ext = sRegType
         End If
         
         ReDim Preserve ExcludeType(et)
         ExcludeType(et) = ext
         et = et + 1
         
         Exit Function
      Else
         CreateNewShortcut = True
         Call SyncShortcut(dfile, fNameOnly, SetReg)
      End If
   Else
'      If blRegd = True Then
         CreateNewShortcut = True
         Call SyncShortcut(dfile, fNameOnly, SetReg)
'      End If
   End If

errlog:
   blNotified = False
   SetReg = vbNullString
   subErrLog ("modShortcut:CreateNewShortcut")
End Function

Private Sub SyncShortcut(dfile As String, fNameOnly As String, SetReg As String)
On Error GoTo errlog
   Call ShellShortcut(dfile, fNameOnly)
   Form1.subCopyShortcuts (sRecent & "*.lnk")
   If Len(SetReg) > 0 Then
      sRegType = SetReg
   End If
errlog:
   subErrLog ("modShortcut:SyncShortcut")
End Sub

Public Sub ShellShortcut(dfile As String, fNameOnly As String)
On Error GoTo errlog
   Dim oShellLink As Object
   Dim WshShell As Object
   Set WshShell = CreateObject("WScript.Shell")
   Set oShellLink = WshShell.CreateShortcut(sRecent + fNameOnly + ".lnk")
   oShellLink.TargetPath = dfile
   oShellLink.Save
   If GetInputState() <> 0 Then
       DoEvents
   End If
errlog:
   subErrLog ("modShortcut:ShellShortcut")
End Sub

Public Sub DragDropRoutine(Data As ComctlLib.DataObject, Effect As Long, blform2 As Boolean)
On Error Resume Next
   Dim blGDS As Boolean

   If Data.GetData(vbCFText) = True Then
      GoTo filedrop 'test for error 461
   End If

   If Err.Number <> 461 Then
         Dim thefile As String
      If InStr(Data.GetData(vbCFText), "redir?url=file%3A%2F%2F") > 0 Then   'google desktop file
         blGDS = True
         Erase blDontCopy
         Form1.Timer1.Enabled = False
         Effect = vbDropEffectCopy
         thefile = ParseGDString(Data.GetData(vbCFText))
         If DoesFileExist(thefile) Then
            Call subAddShortcut(thefile, 0, 1)
         Else
            MsgBox "File does not exist at the last location recorded by Google Desktop.", vbInformation
            Exit Sub
         End If
'      ElseIf InStr(data.GetData(vbCFText), "redir?url=http%3A%2F%2F") > 0 Then   'google mail file
'         blGDS = True
'         Erase blDontCopy
'         Form1.Timer1.Enabled = False
'         Effect = vbDropEffectCopy
''         Dim thefile As String
'         thefile = data.GetData(vbCFText)
''         If DoesFileExist(thefile) Then
'            Call subAddShortcut(thefile, 0, 1)
''         Else
''            MsgBox "File does not exist at the last location recorded by Google Desktop.", vbInformation
''            Exit Sub
''         End If
      Else
         GoTo errlog
      End If
    Else
filedrop:
'On Error GoTo 0
      CheckPreviousTypes      'need to refresh PrevType() with what's actually in Form1.Listview1

      If InitDrag(Data.Files.Count) = True Then
         Dim i As Integer
         Erase blDontCopy
         Form1.Timer1.Enabled = False
         Effect = vbDropEffectCopy
'         Erase KeepSelected
         ReDim Preserve DroppedList(Data.Files.Count)
         For i = 1 To Data.Files.Count
            DroppedList(i - 1) = CStr(Data.Files(i))
            Call subAddShortcut(CStr(Data.Files(i)), i - 1, Data.Files.Count)
         Next i
      End If
      If blform2 Then
         blform2 = False
         If InitDrag(Data.Files.Count) = False Then
            Form2.PerhapsaMailAttachment     'lame but can be dropped to a folder location
            Exit Sub
         End If
      End If
   End If

   SummaryMessage
   If blMultFileDrop = True Then
      Form1.RestoreNormalState
      TrackStatus (Data.Files.Count)
      Form2.subLbCount
   Else
      If blGDS Then
         TrackStatus (1)
      ElseIf Data.Files.Count > 0 Then
         TrackStatus (Data.Files.Count)
      End If
   End If
      
errlog:
   Effect = vbDropEffectNone
   sRegType = vbNullString
subErrLog ("modShortcut:DragDropRoutine")
End Sub

Private Sub CheckPreviousTypes()
On Error GoTo errlog
   lsc = 0
   Erase PrevType
   For lsc = 0 To Form1.ListView1.ListItems.Count - 1
      ReDim Preserve PrevType(lsc)
      PrevType(lsc) = Form1.ListView1.ListItems(lsc + 1).Text
   Next
errlog:
subErrLog ("modShortcut:CheckPreviousTypes")
End Sub
 
Private Function ParseGDString(filename As String) As String
On Error GoTo errlog
Dim thestring As String
Dim thedrive As String
Dim chars As Integer
Dim thestart As Integer, theend As Integer

   thestart = InStr(filename, "2F%2F") + 4
   theend = InStr(filename, "%3Fevent%")
   thestring = Mid$(filename, thestart, theend - thestart)
   thestring = Right$(thestring, Len(thestring) - 7)
   thedrive = Left(thestring, 1)
   thestring = Right$(thestring, Len(thestring) - 1)
   
   For chars = 1 To Len(thestring)
      If Mid$(thestring, chars, 3) = "%5C" Then
         thestring = Left$(thestring, chars - 1) + "\" + Right$(thestring, Len(thestring) - chars - 2)
      End If
   Next
   For chars = 1 To Len(thestring)
      If Mid$(thestring, chars, 3) = "%27" Then
         thestring = Left$(thestring, chars - 1) + "'" + Right$(thestring, Len(thestring) - chars - 2)
      End If
   Next
   For chars = 1 To Len(thestring)
      If Mid$(thestring, chars, 3) = "%28" Then
         thestring = Left$(thestring, chars - 1) + "(" + Right$(thestring, Len(thestring) - chars - 2)
      End If
   Next
   For chars = 1 To Len(thestring)
      If Mid$(thestring, chars, 3) = "%29" Then
         thestring = Left$(thestring, chars - 1) + ")" + Right$(thestring, Len(thestring) - chars - 2)
      End If
   Next
   For chars = 1 To Len(thestring)
      If Mid$(thestring, chars, 3) = "%2D" Then
         thestring = Left$(thestring, chars - 1) + "-" + Right$(thestring, Len(thestring) - chars - 2)
      End If
   Next
   For chars = 1 To Len(thestring)
      If Mid$(thestring, chars, 3) = "%5F" Then
         thestring = Left$(thestring, chars - 1) + "_" + Right$(thestring, Len(thestring) - chars - 2)
      End If
   Next
   For chars = 1 To Len(thestring)
      If Mid$(thestring, chars, 3) = "%2E" Then
         thestring = Left$(thestring, chars - 1) + "." + Right$(thestring, Len(thestring) - chars - 2)
      End If
   Next
   For chars = 1 To Len(thestring)
      If Mid$(thestring, chars, 3) = "%3B" Then
         thestring = Left$(thestring, chars - 1) + ";" + Right$(thestring, Len(thestring) - chars - 2)
      End If
   Next
   For chars = 1 To Len(thestring)
      If Mid$(thestring, chars, 1) = "+" Then
         thestring = Left$(thestring, chars - 1) + " " + Right$(thestring, Len(thestring) - chars)
      End If
   Next

   ParseGDString = thedrive + ":\" + thestring
   
errlog:
   subErrLog ("modShortcut:ParseGDString")
End Function


