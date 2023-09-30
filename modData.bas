Attribute VB_Name = "modData"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' © 2003-2007 Michael M. Ross, Wordwise Solutions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Declare Function CloseHandle Lib "kernel32" _
   (ByVal hfile As Long) As Long
   
Public Declare Function GetLocaleInfo Lib "kernel32" _
   Alias "GetLocaleInfoA" _
  (ByVal Locale As Long, _
   ByVal LCType As Long, _
   ByVal lpLCData As String, _
   ByVal cchData As Long) As Long
   
Private Declare Function GetDiskFreeSpace Lib "kernel32" _
   Alias "GetDiskFreeSpaceA" _
  (ByVal lpRootPathName As String, _
   lpSectorsPerCluster As Long, _
   lpBytesPerSector As Long, _
   lpNumberOfFreeClusters As Long, _
   lpTtoalNumberOfClusters As Long) As Long
   
Public Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
   
Public Declare Function GetUserDefaultLCID Lib "kernel32" () As Long

Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Public SYS_TIME As SYSTEMTIME

Public Const SHGFI_USEFILEATTRIBUTES As Long = &H10
'Public Const SHGFI_TYPENAME As Long = &H400
   
Private Const LOCALE_SSHORTDATE As Long = &H1F    'short date format string
Private Const LVNI_SELECTED As Long = &H2
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_GETNEXTITEM  As Long = (LVM_FIRST + 12)
Private Const LVM_GETSELECTEDCOUNT As Long = (LVM_FIRST + 50)

Public Const OF_READWRITE = &H2

Public KeepSelected() As String    'this will be redim'd as 2 dimensions
Public aSFileName() As String
Public aLastDate() As String   'Gets values of aLastAcc or aLastMod or aCreated
Public strShortDateFormat As String
Public strSavedString As String
Public sInterrupted As String

Public blTrueIcons() As Boolean
Public blTrueDates() As Boolean
Public aNumShort() As String, aNumPath() As String

Public RegSortCol As Byte, DelSortCol As Byte, MultSortCol As Byte
Public RegSortOrder As Byte, DelSortOrder As Byte, MultSortOrder As Byte
Public RegPSortCol As Byte, DelPSortCol As Byte, MultPSortCol As Byte
Public RegPSortOrder As Byte, DelPSortOrder As Byte, MultPSortOrder As Byte

Public DupeFileCount As Integer

Public blNoAccess As Boolean
Public blNoCache As Boolean
Public blInstantStarted As Boolean
Public blLoadedAllTypes As Boolean
Public blAbort As Boolean
Public blVirtual As Boolean
Public blFAT32 As Boolean
Public blScratchCache As Boolean
Public blScroll As Boolean
Public blVFShow As Boolean
Public blDontSaveCols As Boolean
Public blSearchAll As Boolean
Public blBuilding As Boolean
Public blStop As Boolean
Public blExtraCol As Boolean

Public TotalFiles As Long
Public fc As Long
Public dcx As Integer
'Public FndCnt As Long

'Private KeepCount() As Integer
Private lv1x As Integer
Private LastTrueCnt As Long

Private ict As Integer
Private nct As Integer
Private NotConnected() As String
Private IsConnected() As String

Public blVirtualViewTypes As Boolean
Public blOtherViewTypes As Boolean
Public blNoSplash As Boolean

'Private blPA As Boolean
Private blGetLV2Settings As Boolean
Private blUseNeverSync As Boolean
Private blLoadLegacy As Boolean
Private blElse As Boolean
Private blFolderType As Boolean
'Private blUnknownVirtual As Boolean

Public Function RecentFolderChange() As Boolean
On Error GoTo errlog

   If blLoading = False Then
      If Form1.ListView1.ListItems.Count = 0 Then     'failsafe recovery
         RecentFolderChange = True
         Exit Function
      End If
   End If

'Not using watched folder technique to minimize resource use
   Dim FT_CREATE As FILETIME
   Dim FT_ACCESS As FILETIME
   Dim FT_WRITE As FILETIME
   Dim dNewWrite As Date
   Dim lFile As Long

   lFile = CreateFile(sRecent, 0&, FILE_SHARE_READ, 0&, _
                      OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0&)
                      
'   If lFile = 0 Then ???
   
   If lFile > 0 Then
   
      blFAT32 = False
      
      Call GetFileTime(lFile, FT_CREATE, FT_ACCESS, FT_WRITE)
      
     'NTFS folder dates are modified each time a file inside changes, but FAT32 folder modified dates never change -
      'same as created, so need to check the TT folder every time.
         
      If blFAT32 = False Then    'NTFS
      
         dNewWrite = CDate(GetFileDate(FT_WRITE))
         
      Else  'FAT32 only
         If tmCount Mod 2 = 0 Then     'not too bad on CPUs
            RecentFolderChange = True
            Exit Function
         End If
      End If
      
      If dNewWrite > dLastWrite Then
            
         RecentFolderChange = True
   
      Else
      
         RecentFolderChange = False
      
      End If
   
      dLastWrite = dNewWrite
   
   End If
   
errlog:
   Call CloseHandle(lFile)
   subErrLog ("modData: RecentFolderChange")
End Function

Public Sub GetShortDateFormat()
   Dim LCID As Long
   
  'get the locale for the user
   LCID = GetUserDefaultLCID()
   
   If LCID <> 0 Then
      
     'return the short date format once
      strShortDateFormat = GetUserLocaleInfo(LCID, LOCALE_SSHORTDATE)
            
   End If
   
End Sub


Private Function GetUserLocaleInfo(ByVal dwLocaleID As Long, _
                                   ByVal dwLCType As Long) As String

   Dim sReturn As String
   Dim r As Long

  'call the function passing the Locale type
  'variable to retrieve the required size of
  'the string buffer needed
   r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))

  'if successful..
   If r Then

     'pad the buffer with r spaces
      sReturn = Space$(r)

     'and call again passing the buffer
      r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))

     'if successful (r > 0)
      If r Then

        'r holds the size of the string
        'including the terminating null
         GetUserLocaleInfo = Left$(sReturn, r - 1)

      End If

   End If

End Function

Public Function GetFileDate(ft As FILETIME) As String
On Error GoTo errlog
  
'   convert the FILETIME to LOCALTIME, then to SYSTEMTIME type
   Dim ST As SYSTEMTIME
   Dim LT As FILETIME
   Dim t As Long
   Dim DS As Double
   Dim ts As Double
   t = FileTimeToLocalFileTime(ft, LT)
   t = FileTimeToSystemTime(LT, ST)
   If t Then
       DS = DateSerial(ST.wYear, ST.wMonth, ST.wDay)
       ts = TimeSerial(ST.wHour, ST.wMinute, ST.wSecond)
       DS = DS + ts
       If blComparetoTTDate = False Then
         If DS > 0 Then
             GetFileDate = Format$(DS, strShortDateFormat + " hh:mm:ss")
         Else
             GetFileDate = vbNullString
         End If
       Else
         GetFileDate = Format$(DS, strShortDateFormat + " hh:mm:ss")
       End If
   End If
errlog:
subErrLog ("modData:GetFileDate")
End Function

Public Sub PopFromMem()
On Error Resume Next

   Dim ts As Integer
   Dim tc As Long
   Dim blSearchString As Boolean
   Dim CurTypeDate As String
   Dim CurTypeCnt As Integer

   If blColumnClick = True Then Exit Sub
           
   blStop = False
   blSearchString = False
   blBuilding = True
   sInterrupted = ""
   ict = 0
   nct = 0
   Erase NotConnected
   Erase IsConnected

'   If Form2.mFilter.Checked = True Then
      With Form4     'apply filter if any
         If Len(.cbSearch.Text) > 0 Then
            If Form4.fnNewString = True Then
               If blVirtual = True Then
                  If Form5.mIgnoreFilters.Checked = False Then
                     blSearchString = True
                  End If
               Else
                  blSearchString = True
               End If
            End If
         End If
      End With
'   End If
   
   If blSearchString Then
      strSavedString = Form4.cbSearch.Text
   Else
      strSavedString = vbNullString
   End If
   
   If Form2.mFullPath.Checked = True Then
      blFolderSort = True
      Form2.FolderSort
   End If
      
   If Form2.mContainerFolder.Checked = True Then
      blFolderSort = True
      Form2.FolderSort
   End If

   If blLoading = False Then
      Form1.SaveSettings
   End If
   
   blExtraCol = ExtraColumn
   
   SetListview2Columns
   
   Form2.ListView2.ListItems.Clear
   Form2.Refresh
   
'   If Form1.mSyncShow.Checked = True And blVirtual = False Or blVirtual = True Then
   If Form1.mSyncShow.Checked = True Or blVirtual = True Then
      ReDim Preserve blTrueIcons(UBound(fType))
   End If
   
   ReDim Preserve blTrueDates(UBound(fType))

   With Form1.ListView1
      CurTypeDate = .ListItems(.SelectedItem.Index).SubItems(2)
      CurTypeCnt = .ListItems(.SelectedItem.Index).SubItems(1)
   End With
   
'   FndCnt = 0
   
'   Form2.Caption = "TaskTracker - Loading..."
   Form2.StatusBar1.Panels(1).Text = "Building list..."
            
   If (blAddType = True Or blAllTypes = True Or blSearchAll = True) And blVirtual = False Then
   'Iterate through each bolded type, then iterate through the files of this type
      With Form1.ListView1
      
         For ts = 1 To .ListItems.Count
         
'            Debug.Print Str(ts)
            
            If ts > .ListItems.Count Then Exit For    'bugfix perhaps
         
            If .ListItems(ts).Selected = True Then
            
'            Debug.Print .ListItems(ts).Text
            
               For TTCnt = 0 To UBound(fType)
               
'                  Debug.Print Str(TTCnt)
               
                  If Len(fType(TTCnt)) > 0 Then
                  
                    If fType(TTCnt) = .ListItems(ts).Text Or (blThisType = True And blSearchAll = True) Then
                    
                       If blUserRefresh = True Then
                          blTrueDates(TTCnt) = False
                       End If
                       
                       If blSearchString = True Then
                          If StringFilter(TTCnt) = False Then
                              GoTo filter1
                          End If
                       End If
                       
                       If Left$(tType(TTCnt), 14) = "Virtual Folder" Then    'otherwise same file will appear twice
                          GoTo filter1
                       End If
                       
                       If blNotValidated(TTCnt) = True Then
                          tc = tc + 1
'                         **********
                          Call PopForm2(TTCnt, tc)
'                         **********
                       Else
                                              
                          If Form1.mAlwaysValidate.Checked = True Or blUserRefresh _
                             Or blLoading Or Form1.blNewShortcuts Then
                             If DoesFileExist(aSFileName(TTCnt)) = True Then
                                tc = tc + 1
'                               **********
                                Call PopForm2(TTCnt, tc)
'                               **********
                             Else
                                blNotValidated(TTCnt) = True
                                fType(TTCnt) = "Deleted or Renamed"
                                tType(TTCnt) = "Unknown"
                             End If
                          Else
                              tc = tc + 1
'                            **********
                             Call PopForm2(TTCnt, tc)
'                            **********
                          End If
                          
                       End If
                    End If
                  End If
        '         DoEvents - screws everything up
                  If blStop Then
                     blStop = False
                     sInterrupted = " (Interrupted)"
                     Exit Sub
                  End If
filter1:
               Next TTCnt
            End If
         Next ts
      End With
      
   Else  'blThisType
   
      If Len(sType) = 0 Then
         Form1.PopsType
      End If
   
      'Iterate through the files of the selected type
      For TTCnt = 0 To UBound(fType)
         If Len(fType(TTCnt)) > 0 Then
         '  Virtual Folders only
         '  *************************************************
            If Left$(tType(TTCnt), 14) = "Virtual Folder" Then
            
               If blVirtual = True Then
               
                   If tType(TTCnt) = Form5.ListView3.SelectedItem.Key Then
                   
                        If blUserRefresh = True Then
                           blTrueDates(TTCnt) = False
                        End If
                        
                        If blSearchString = True Then
                           If StringFilter(TTCnt) = False Then
                              GoTo filter2
                           End If
                        End If

                        If blNotValidated(TTCnt) = True Then
                           tc = tc + 1
'                          **********
                           Call PopForm2(TTCnt, tc)
'                          **********
                        Else
                           If Form1.mAlwaysValidate.Checked = True Or blUserRefresh _
                               Or blLoading Or Form1.blNewShortcuts Then
                               tc = tc + 1
'                              **********
                               Call PopForm2(TTCnt, tc)
'                              **********
                           Else
                               tc = tc + 1
'                              **********
                               Call PopForm2(TTCnt, tc)
'                              **********
                           End If
                        End If
                           
                  End If
               End If
         '  *************************************************
            Else
               If blVirtual = False Then
               
                  If fType(TTCnt) = sType Then
                  
                     If blUserRefresh = True Then
                        blTrueDates(TTCnt) = False
                     End If
               
                     If blSearchString = True Then
                        If StringFilter(TTCnt) = False Then
                           GoTo filter2
                        End If
                     End If
                     
                     If blNotValidated(TTCnt) = True Then
                        tc = tc + 1
'                      **********
                       Call PopForm2(TTCnt, tc)
'                      **********
                     Else
                        If Form1.mAlwaysValidate.Checked = True Or blUserRefresh Or blLoading Then
                            If DoesFileExist(aSFileName(TTCnt)) = True Then
                                tc = tc + 1
'                               **********
                                Call PopForm2(TTCnt, tc)
'                               **********
                             Else
                                blNotValidated(TTCnt) = True
                                fType(TTCnt) = "Deleted or Renamed"
                                tType(TTCnt) = "Unknown"
                             End If
                        Else
                            tc = tc + 1
'                           **********
                            Call PopForm2(TTCnt, tc)
'                           **********
                        End If
                     End If
                  End If
               End If
            End If
         End If
'        DoEvents - screws everything up
         If blStop Then
            blStop = False
            sInterrupted = " (Interrupted)"
            Exit Sub
         End If
filter2:
      Next TTCnt
      
   End If
            
   With Form2
      If .ListView2.ListItems.Count > 0 Then
         blFirstSort = True
'         Call .ListView2_ColumnClick(.ListView2.ColumnHeaders(1))
         .ListView2.ListItems(1).Selected = True
         .ListView2.ListItems(1).EnsureVisible
      Else
         If blVirtualView = False Then
            If blSearchString = False And Form4.cbDates.ListIndex = 0 Then     'do not remove empty types unless no filters
               With Form1
                  .ListView1.ListItems.Remove (.ListView1.SelectedItem.Text)
                  .ListView1_Click
               End With
            End If
         End If
      End If
      If blMultFileDrop = False Then
         Form1.StatusBar1.Panels(1).Text = Str$(Form1.ListView1.ListItems.Count) + " file types loaded."
      End If
   End With
   
   'only when loading
   If blLoading = True Then
      blFirstSort = True
      
      If Form1.mSingleClick.Checked = True Then
         Form1.ListView1_Click
         Form2.SetTitle
'         If blMinStart = True Then
'            If Form1.mTaskbar.Checked = True Then
'               Form1.WindowState = vbMinimized
'            End If
'         End If
'         Form2.Show
         Form2.WindowState = Form1.WindowState
         Form2.subSetSize
         If blMinStart = False Then
            If blVFShow = True Then    'reshow on startup
               If Form1.blReloading = False Then
                  Form1.mVirtual_Click
               End If
            End If
         End If
      End If
      
      Form2.ListView2.ListItems(1).Selected = False
      If Form1.mFocusTypes.Checked = True Then
         Form1.Visible = True
         Form1.SetFocus
      Else
         Form2.SetFocus
      End If
      
            
      'sort, then select first type
      Form1.subClickList
'      If Form1.blReloading = False Then
         Form1.ListView1.ListItems(1).Selected = True
'      End If
      
      Splash.Timer1.Enabled = False
      Unload Splash
      
'      If blUseNeverSync = True Then
'         blUseNeverSync = False
'         If Form1.WindowState = vbNormal Then
'            If Command$ <> "-m" Then
'                  MsgBox "Loading was delayed because some files could not be accessed. " + vbNewLine + _
'                  "TaskTracker finished loading without synchronizing all file data, including icons." + vbNewLine + vbNewLine _
'                  + "(If you see this message often, change your Preferences to: Performance " + vbNewLine + _
'                  "Settings > Synchronization > On First Showing.)", vbInformation
'            End If
'         End If
'      End If
            
   Else
   
      'check for change in date or count - and update Listview1 sort order
      With Form1.ListView1
         If .ListItems(.SelectedItem.Index).SubItems(2) <> CurTypeDate Or _
            .ListItems(.SelectedItem.Index).SubItems(1) <> CurTypeCnt Then
               Form1.subClickList
         End If
      End With

   End If
   
   If Form1.mIcons.Checked = False Then
      blScroll = True
      Form1.Timer3.Enabled = True   'bugfix - can't call ScrollBarCompensation directly so use timer
   End If
   
   If VB.Forms.Count = 4 Then
      About.txSystem = Trim$(Str$(vf)) + " Files"
      About.txNetwork = Trim$(Str$(rf)) + " Files"
   End If
   
   If blSnapshot = True Then
      Call subStatusLog("Types")
   End If
   
   Call Form2.ListView2_ColumnClick(Form2.ListView2.ColumnHeaders(1))
   
   If Form1.mAutoPreview.Checked = True Or Form2.mPreview.Checked = True Then
      If blLoading = False Then
        blForm6Load = True
        Form6.Timer1.Enabled = True
      End If
   End If

errlog:
   Form2.subLbCount
   LeaveSelected
   blBuilding = False
   blInstantStarted = True
   blClicked = True
   blBegin = False
   blSearchAll = False
   subErrLog ("modData: PopFromMem")
End Sub

'string search filter
Private Function StringFilter(TTCnt As Long) As Boolean
On Error GoTo errlog

   Dim strSearch As String
   Dim SearchPhrase As String
   
   strSearch = strSavedString
   
   If InStr(strSearch, ":") > 0 Then
   
      Do While InStr(strSearch, ":") > 0  'multiphrase
      
         If Left$(strSearch, 1) = ":" Then
            strSearch = Right$(strSearch, Len(strSearch) - 1)
         End If
         
         If InStr(strSearch, ":") = 0 Then
            Exit Do
         End If
         
         SearchPhrase = Left$(strSearch, InStr(strSearch, ":") - 1)
         
         If SearchExpression(SearchPhrase, TTCnt) = True Then
            StringFilter = True
            Exit Function
         End If
            
         strSearch = Right$(strSearch, Len(strSearch) - Len(SearchPhrase) - 1)
         If strSearch = ":" Or strSearch = ": " Then Exit Do
         If Len(strSearch) = 0 Then Exit Do
            
      Loop
      
   End If
   
   If Len(strSearch) = 0 Then
      StringFilter = False
      Exit Function
   End If
   
   If Len(strSearch) > 0 And strSearch <> "*" And strSearch <> "." And _
      strSearch <> "*." And strSearch <> "*.*" Then
      If SearchExpression(strSearch, TTCnt) = True Then     'single phrase or last phrase
         StringFilter = True
      Else
         StringFilter = False
      End If
   Else
      StringFilter = False
   End If
      
errlog:
   subErrLog ("modData:StringFilter")
End Function

Private Function SearchExpression(strSearch As String, TTCnt As Long) As Boolean
On Error GoTo errlog
   
   Dim fNameOnly As String
   Dim searchpart As String
   
   fNameOnly = Right$(aSFileName(TTCnt), Len(aSFileName(TTCnt)) - InStrRev(aSFileName(TTCnt), "\"))
   
'   If Left$(strSearch, 1) = "*" Then
'      strSearch = Right$(strSearch, Len(strSearch) - 1)
'   End If
   
'   Debug.Print Right$(strSearch, Len(strSearch) - 1)
   
   
    If InStr(strSearch, "*") > 0 Then
              
      If InStr(strSearch, "*") > 1 Then
         Dim firstpos As Integer
         firstpos = InStr(strSearch, "*") - 1
         If Left$(LCase$(strSearch), firstpos) <> Left$(LCase$(fNameOnly), firstpos) Then
            SearchExpression = False
            Exit Function
         End If
      End If
      
      Do While InStr(strSearch, "*") > 0
               
         searchpart = Left$(strSearch, InStr(strSearch, "*") - 1)
         
'         If PerformSearch(searchpart) = False Then Exit Do
                  
         If InStr(1, fNameOnly, searchpart, vbTextCompare) = 0 Then
            SearchExpression = False
            Exit Do
         Else
            SearchExpression = True
         End If
         
         strSearch = Right$(strSearch, Len(strSearch) - Len(searchpart) - 1)
'         If PerformSearch(strSearch) = False Then Exit Do
                  
      Loop
      
      If Len(strSearch) = 0 Then
         Exit Function
      End If
            
      'remaining - must contain string
      If InStr(1, fNameOnly, strSearch, vbTextCompare) = 0 Then
         SearchExpression = False
      Else
         SearchExpression = True
      End If
   
   Else
   
      'regular - must begin with string
      If InStr(1, Left$(fNameOnly, Len(strSearch)), strSearch, vbTextCompare) = 0 Then
         SearchExpression = False
      Else
         SearchExpression = True
      End If
      
   End If
      
errlog:
   subErrLog ("modData:SearchExpression")
End Function

'Private Function PerformSearch(SP As String) As Boolean
''   If Len(cbSearch.Text) > 0 And cbSearch.Text <> "Enter search string" And _
''      Right$(cbSearch.Text, 1) <> "*" And Right$(cbSearch.Text, 1) <> ":" Then 'ignore wildcard and delimiter
''      PerformSearch = True
''   End If
'   If SP = "" Then
'      PerformSearch = False
'      Exit Function
'   End If
''   Debug.Print SP
'   If blSearchAll = True Then
'      If Right$(SP, 1) <> "*" And Right$(SP, 1) <> "." Then
'         PerformSearch = True
'      Else
'         PerformSearch = False
'      End If
'   Else
'      PerformSearch = True
'   End If
'End Function

Public Sub ReshowOtherWindows()
On Error GoTo errlog
      If Form2.mFilter.Checked = True Then
         Form4.Show
      End If
      
      If Form1.mVirtual.Checked = True Then
         Form5.Show
      End If
      
'      If Form2.mPreview.Checked = True Then
      If blPreviewOpen = True Then
         Form6.Show
      End If
      
errlog:
   subErrLog ("modData:ReshowOtherWindows")
End Sub

Public Sub SetListview2Columns()
On Error Resume Next

   Dim clm As ColumnHeader
      
   If blLoading = True Then
      With Form2.ListView2
         If blNoCache = True Then
            .ListItems.Clear
            Set .SmallIcons = Nothing
         End If
      ' ********************************************************
         Set .SmallIcons = Form1.ImageList2     'initializes imagelist2
      ' ********************************************************
      End With
      If blGetLV2Settings = True Then Exit Sub
      blGetLV2Settings = True
      Form2.GetListView2Settings
'      blFirstThisType = True
   Else
      If Form1.mSaveSize.Checked = True Then
         If blDontSaveCols = False Then
            SaveColumnWidths
         Else
            blDontSaveCols = False
         End If
      End If
   End If

   With Form2.ListView2
   
      .ColumnHeaders.Clear
      
         If AllorAddorSpecial = True Then
         
            SetViewTypes
            
            If blFolderSort = True Then
            
                If blVirtualViewTypes And blVirtualView Or _
                  blVirtualView = False And blOtherViewTypes Then
                  
                  Set clm = .ColumnHeaders.Add(, "string1", "File Name", col1PAwidth)
                
                  If Form2.mFullPath.Checked = True Then
                    Set clm = .ColumnHeaders.Add(, "string2", "Folder Path", col2PAwidth)
                  Else
                    Set clm = .ColumnHeaders.Add(, "string2", "Containing Folder", col2PAwidth)
                  End If
                  
                  Set clm = .ColumnHeaders.Add(, "string3", "Type", col3PAwidth)
                  
                  If Form2.mModified.Checked = True Then
                     Set clm = .ColumnHeaders.Add(, "date1", "Last Modified", col4PAwidth)
                  ElseIf Form2.mCreated.Checked = True Then
                     Set clm = .ColumnHeaders.Add(, "date2", "Created", col4PAwidth)
                  ElseIf Form2.mAccessed.Checked = True Then
                     Set clm = .ColumnHeaders.Add(, "date3", "Last Accessed", col4PAwidth)
                  End If
                  iShortCol = 4
                  
                Else
                
                  Set clm = .ColumnHeaders.Add(, "string1", "File Name", col1NPAwidth)
                  
                  If Form2.mFullPath.Checked = True Then
                    Set clm = .ColumnHeaders.Add(, "string2", "Folder Path", col2NPAwidth)
                  Else
                    Set clm = .ColumnHeaders.Add(, "string2", "Containing Folder", col2NPAwidth)
                  End If
                                    
                  If Form2.mModified.Checked = True Then
                     Set clm = .ColumnHeaders.Add(, "date1", "Last Modified", col4NPAwidth)
                  ElseIf Form2.mCreated.Checked = True Then
                     Set clm = .ColumnHeaders.Add(, "date2", "Created", col4NPAwidth)
                  ElseIf Form2.mAccessed.Checked = True Then
                     Set clm = .ColumnHeaders.Add(, "date3", "Last Accessed", col4NPAwidth)
                  End If
                  
                  iShortCol = 3
                End If
                
                Set clm = .ColumnHeaders.Add(, "shortcut", "Shortcut", 0)
            
             Else
                
                If blVirtualViewTypes And blVirtualView Or _
                  blVirtualView = False And blOtherViewTypes Then
                  
                  Set clm = .ColumnHeaders.Add(, "string1", "File Name", col1Awidth)
                  Set clm = .ColumnHeaders.Add(, "string2", "Type", col2Awidth)
                  If Form2.mModified.Checked = True Then
                     Set clm = .ColumnHeaders.Add(, "date1", "Last Modified", col3Awidth)
                  ElseIf Form2.mCreated.Checked = True Then
                     Set clm = .ColumnHeaders.Add(, "date2", "Created", col3Awidth)
                  ElseIf Form2.mAccessed.Checked = True Then
                     Set clm = .ColumnHeaders.Add(, "date3", "Last Accessed", col3Awidth)
                  End If
                  iShortCol = 3
                  
                Else
                
                  Set clm = .ColumnHeaders.Add(, "string1", "File Name", col1NAwidth)
                  If Form2.mModified.Checked = True Then
                     Set clm = .ColumnHeaders.Add(, "date1", "Last Modified", col3NAwidth)
                  ElseIf Form2.mCreated.Checked = True Then
                     Set clm = .ColumnHeaders.Add(, "date2", "Created", col3NAwidth)
                  ElseIf Form2.mAccessed.Checked = True Then
                     Set clm = .ColumnHeaders.Add(, "date3", "Last Accessed", col3NAwidth)
                  End If
                  iShortCol = 2
                  
                End If
               
                Set clm = .ColumnHeaders.Add(, "shortcut", "Shortcut", 0)
            
            End If
            
   '        All or Add: more than one type
   '        ********************************************************
            If blAllTypes = True Or (blAddType = True And blVirtualView = False) Then
          
               Set clm = .ColumnHeaders.Add(, , "", 0)
               Form2.ListView2.SortOrder = MultSortOrder
               Form2.ListView2.SortKey = MultSortCol
                           
   '        Special types: Deleted or Renamed and Network or Removable Disk
   '        ********************************************************
            ElseIf blThisType = True And blVirtualView = False And blExtraCol Or _
               blThisType = True And blExtraCol Then
                                       
               If blFolderSort = True Then
                  .SortOrder = DelPSortOrder
                  .SortKey = DelPSortCol
               Else
                  .SortOrder = DelSortOrder
                  .SortKey = DelSortCol
               End If
               
'               blPA = True
            End If
            
'        All the other types when selected individually
'        ********************************************************
        Else
                    
            If blFolderSort = True Then
            
               Set clm = .ColumnHeaders.Add(, "string1", "File Name", col1Pwidth)
               If Form2.mFullPath.Checked = True Then
                  Set clm = .ColumnHeaders.Add(, "string2", "Folder Path", col2Pwidth)
               Else
                  Set clm = .ColumnHeaders.Add(, "string2", "Containing Folder", col2Pwidth)
               End If
               If Form2.mModified.Checked = True Then
                  Set clm = .ColumnHeaders.Add(, "date1", "Last Modified", col3Pwidth)
               ElseIf Form2.mCreated.Checked = True Then
                  Set clm = .ColumnHeaders.Add(, "date2", "Created", col3Pwidth)
               ElseIf Form2.mAccessed.Checked = True Then
                  Set clm = .ColumnHeaders.Add(, "date3", "Last Accessed", col3Pwidth)
               End If
               Set clm = .ColumnHeaders.Add(, "shortcut", "Shortcut", 0)
               iShortCol = 3
            
            Else
            
               Set clm = .ColumnHeaders.Add(, "string1", "File Name", col1width)
               If Form2.mModified.Checked = True Then
                  Set clm = .ColumnHeaders.Add(, "date1", "Last Modified", col2width)
               ElseIf Form2.mCreated.Checked = True Then
                  Set clm = .ColumnHeaders.Add(, "date2", "Created", col2width)
               ElseIf Form2.mAccessed.Checked = True Then
                  Set clm = .ColumnHeaders.Add(, "date3", "Last Accessed", col2width)
               End If
               Set clm = .ColumnHeaders.Add(, "shortcut", "Shortcut", 0)
               iShortCol = 2
            
            End If
            
            If blFolderSort = True Then
               .SortOrder = RegPSortOrder
               .SortKey = RegPSortCol
            Else
               .SortOrder = RegSortOrder
               .SortKey = RegSortCol
            End If
        End If
        
   End With
   
'  ***************
   Form2.SetColPos
'  ***************

errlog:
   subErrLog ("modData:Listview2Columns")
End Sub

Private Sub SetViewTypes()
On Error GoTo errlog
   If LenB(GetSetting("TaskTracker", "Settings", "VirtualViewTypes")) > 0 Then
      blVirtualViewTypes = CBool(GetSetting("TaskTracker", "Settings", "VirtualViewTypes"))
   Else
      blVirtualViewTypes = True 'default
   End If
   If LenB(GetSetting("TaskTracker", "Settings", "OtherViewTypes")) > 0 Then
      blOtherViewTypes = CBool(GetSetting("TaskTracker", "Settings", "OtherViewTypes"))
   Else
      blOtherViewTypes = True 'default
   End If
   If blVirtualViewTypes = True And blVirtualView = True _
      Or blOtherViewTypes = True And blVirtualView = False Then
      Form2.mFileTypes.Checked = True
   Else
      Form2.mFileTypes.Checked = False
   End If
errlog:
   subErrLog ("modData:SetViewTypes")
End Sub

'Saves the settings for the alltype/addtype ("A") config and the views with and without a folder ("P") column
Private Sub SaveColumnWidths()
On Error GoTo errlog

   With Form2.ListView2
   
        If AllorAddorSpecial = True Then
                                    
            If blElse = True Then      'preserves separate columnwidths for blthistype
               blElse = False
               Exit Sub
            End If
            
            If blFolderSort = True Then
            
               If blVirtualViewTypes And blVirtualView Or _
                  blVirtualView = False And blOtherViewTypes Then
                  col1PAwidth = .ColumnHeaders(1).Width
                  col2PAwidth = .ColumnHeaders(2).Width
                  col3PAwidth = .ColumnHeaders(3).Width
                  col4PAwidth = .ColumnHeaders(4).Width
'                  colSwidth = .ColumnHeaders(5).Width
               Else
                  col1NPAwidth = .ColumnHeaders(1).Width
                  col2NPAwidth = .ColumnHeaders(2).Width
                  col4NPAwidth = .ColumnHeaders(3).Width
'                  colSwidth = .ColumnHeaders(4).Width
               End If
            
             Else
               
               If blVirtualViewTypes And blVirtualView Or _
                  blVirtualView = False And blOtherViewTypes Then
                  col1Awidth = .ColumnHeaders(1).Width
                  col2Awidth = .ColumnHeaders(2).Width
                  col3Awidth = .ColumnHeaders(3).Width
'                  colSwidth = .ColumnHeaders(4).Width
               Else
                  col1NAwidth = .ColumnHeaders(1).Width
                  col3NAwidth = .ColumnHeaders(2).Width
'                  colSwidth = .ColumnHeaders(3).Width
               End If
            
            End If
                        
'        All the other types when selected individually
'        ********************************************************
        Else
            If blElse = False Then
               blElse = True
               Exit Sub
            End If
                    
            If blFolderSort = True Then
            
               col1Pwidth = .ColumnHeaders(1).Width
               col2Pwidth = .ColumnHeaders(2).Width
               col3Pwidth = .ColumnHeaders(3).Width
               col4Swidth = .ColumnHeaders(4).Width
            
             Else
             
              col1width = .ColumnHeaders(1).Width
              col2width = .ColumnHeaders(2).Width
'              colSwidth = .ColumnHeaders(3).Width

            End If
      End If
   End With
errlog:
   subErrLog ("modData:SaveColumnWidths")
End Sub

Private Function ListView_GetSelectedCount(hWnd As Long) As Long

   ListView_GetSelectedCount = SendMessage(hWnd, LVM_GETSELECTEDCOUNT, 0&, ByVal 0&)

End Function

Private Function ListView_GetNextItem(hWnd As Long, Index As Long, flags As Long) As Long

   ListView_GetNextItem = SendMessage(hWnd, LVM_GETNEXTITEM, Index, ByVal flags)

End Function

Public Function nSelected() As Integer
On Error GoTo errlog

   Dim Index As Long
   Dim numSelected As Long
   
   numSelected = ListView_GetSelectedCount(Form2.ListView2.hWnd)

   If numSelected <> 0 Then
         
      Do
      
         Index = ListView_GetNextItem(Form2.ListView2.hWnd, Index, LVNI_SELECTED)
         
         If Index > -1 Then
            nSelected = CInt(Index) + 1
            Exit Do
         Else
            nSelected = 1
            Exit Do
         End If
      
      Loop Until Index = -1
      
   End If
   
errlog:
   subErrLog ("modData: nSelected")
End Function

Public Sub SaveSelected()
On Error GoTo errlog
'Save selected item indexes in Listview2 for each type in Listview1
   Dim lv2x As Integer
   dcx = 1
   lv1x = Form1.ListView1.SelectedItem.Index
   For lv2x = 1 To Form2.ListView2.ListItems.Count
      If Form2.ListView2.ListItems(lv2x).Selected = True Then
         KeepSelected(lv1x, dcx) = Form2.ListView2.ListItems(lv2x).SubItems(iShortCol)
      Else
         KeepSelected(lv1x, dcx) = vbNullString
      End If
      dcx = dcx + 1
   Next lv2x

errlog:
'If Err.Number <> 0 And Err.Number <> 9 Then subErrLog ("modData: SaveSelected")
   subErrLog ("modData: SaveSelected")
End Sub

Public Sub LeaveSelected()
On Error GoTo errlog
   
   Dim blLeave As Boolean
   
   With Form2.ListView2
   
      On Error Resume Next
      '2-Dim Array (Listview1 index, Listview2 index)
      ReDim Preserve KeepSelected(Form1.ListView1.ListItems.Count, fCount)
      On Error GoTo 0
      
      If .ListItems.Count > 0 Then
      
         Dim LVS As Integer

         blFindSelected = False
         For LVS = 1 To .ListItems.Count
               
            If Len(SelectThisFile) = 0 Then
               If KeepSelected(Form1.ListView1.SelectedItem.Index, LVS) = Form2.ListView2.ListItems(LVS).SubItems(iShortCol) Then
                  .ListItems(LVS).Selected = True
                  .ListItems(LVS).EnsureVisible
                  blLeave = True
               Else
                  .ListItems(LVS).Selected = False
               End If
            Else
'               If sType = SelectThisType Then   'selection after dropping file
                  If Len(SelectThisFile) > 0 Then
                     If SelectThisFile = .ListItems(LVS).SubItems(iShortCol) Then
                        .ListItems(LVS).Selected = True
                        .ListItems(LVS).EnsureVisible
                        blFindSelected = True
                        SaveSelected
                        blLeave = True
                        Exit For
                     Else
                        .ListItems(LVS).Selected = False
                     End If
'                  End If
               End If
            End If
         
         Next LVS
                
      End If
      
      
errlog:

      If blLeave = False Then
         For LVS = 2 To .ListItems.Count
            .ListItems(LVS).Selected = False
         Next LVS
         .ListItems(1).Selected = True
         .ListItems(1).EnsureVisible
      End If
      
   End With
   
'   If Err.Number <> 0 And Err.Number <> 9 Then subErrLog ("modData: LeaveSelected")
   subErrLog ("modData: LeaveSelected")
End Sub

Private Sub PopForm2(TTCnt As Long, tc As Long)
On Error Resume Next

   If Len(aSFileName(TTCnt)) = 0 Then
      Exit Sub
   End If
   
   If Left$(tType(TTCnt), 5) = "hide-" Then
      Exit Sub
   End If
   
'   If blSearchAll = False And blSelectAll = False Then
         CallDateTheFile1              'to speed things up, don't use if date filter set
'   End If
            
   If Len(aLastDate(TTCnt)) = 0 Then
      If Err.Number = 9 Then
         GoTo nodate
      End If
      If Form4.cbDates.ListIndex = 0 Then 'All
         GoTo nodate
      Else
         Exit Sub
      End If
   End If
      
   'fnDateRange evaluates file dates against selected date range
   If fnDateRange(CDate(aLastDate(TTCnt))) = False Then
      Exit Sub
   End If

   Dim itmx As ListItem
   Dim FormCnt As Long
   Dim fNameOnly As String
   Dim fpathonly As String
   Dim blExtorFolder As Boolean
'   Dim blDoEvents As Boolean
   
nodate:  'if no date found, include in list for "All Dates"
      
      'Populate Listview2 - first the filename, preserving path data in aNumShort/aNumPath
   With Form2
   
        ReDim Preserve aNumShort(tc)
        aNumShort(tc) = aShortcut(TTCnt)
        
        ReDim Preserve aNumPath(tc)
        aNumPath(tc) = aSFileName(TTCnt)
        
        'aSFilename array parsed according to Folder/Ext options
        fNameOnly = Right$(aSFileName(TTCnt), Len(aSFileName(TTCnt)) - InStrRev(aSFileName(TTCnt), "\"))
        fpathonly = Left$(aSFileName(TTCnt), Len(aSFileName(TTCnt)) - Len(fNameOnly))
        
        If Form2.mFullPath.Checked = True Then
            fpathonly = Left$(aSFileName(TTCnt), Len(aSFileName(TTCnt)) - Len(fNameOnly))
        ElseIf Form2.mContainerFolder.Checked = True Then
            fpathonly = Left$(fpathonly, Len(fpathonly) - 1)
            fpathonly = Right$(fpathonly, Len(fpathonly) - InStrRev(fpathonly, "\"))
        End If

        If .mExt.Checked = True Then
            blExtorFolder = True
        Else
            blExtorFolder = False
            If VerifyFolderType = True Then
               blExtorFolder = True
            End If
        End If
            
        If blExtorFolder = True Then
        
            Set itmx = .ListView2.ListItems.Add(, , fNameOnly)

        Else
        
            sExt = Right$(aSFileName(TTCnt), Len(aSFileName(TTCnt)) - InStrRev(aSFileName(TTCnt), "."))
                
            If InStrRev(fNameOnly, ".") > 0 Then
               Set itmx = .ListView2.ListItems.Add(, , _
                     Left$(fNameOnly, Len(fNameOnly) - (Len(fNameOnly) - InStrRev(fNameOnly, ".")) - 1))
            Else    'folder or folder+file with no ext
               Set itmx = .ListView2.ListItems.Add(, , Left$(fNameOnly, Len(fNameOnly)))
            End If
         
        End If
        
   End With
   
   If sType = "Unknown" And Form1.mAlwaysValidate.Checked = True Then
      blTrueIcons(TTCnt) = False
   End If
   
   'Calls GetTrueIcon when using default Synchronize > On First Showing scheme
   'Each file runs through this once unless reloading or always validated
   If blNoCache = False Then
      If Form1.mSyncShow.Checked = True Or blVirtual = True Then
         If blTrueIcons(TTCnt) = False Then
'           Debug.Print Str(TTCnt)
            GetTrueIcon (TTCnt)
            blTrueIcons(TTCnt) = True
            blRefresh = True
            FormCnt = Form2.ListView2.ListItems.Count
            'status bar first load feedback
            If FormCnt > 10 Then
               If blThisType And blSearchAll = False Then
                  Form2.StatusBar1.Panels(1).Text = "Building list: " + Trim$(Str$(FormCnt)) & " " & sType & " files found..."
               Else
                  Form2.StatusBar1.Panels(1).Text = "Building list: " + Trim$(Str$(FormCnt)) & " files found..."
               End If
            End If
         End If
      End If
   End If
         
   'Add the unique icon of each file
   If blNoCache = False Then
      If Form1.mSyncShow.Checked = True And blVirtual = False Then
         itmx.SmallIcon = Form1.ImageList2.ListImages(aShortcut(TTCnt)).Key
      ElseIf blVirtual = True Then
         itmx.SmallIcon = Form1.ImageList2.ListImages(aShortcut(TTCnt) + tType(TTCnt)).Key
      Else
         itmx.SmallIcon = Form1.ImageList2.ListImages(aShortcut(TTCnt)).Key
      End If
   Else
      itmx.SmallIcon = Form1.ImageList2.ListImages(aShortcut(TTCnt)).Key
   End If

  'Populate Listview2 - remaining data
  If AllorAddorSpecial = True Then
       
       If blFolderSort = False Then
                 
          If blVirtualView = True And blVirtualViewTypes = True Or _
             blVirtualView = False And blOtherViewTypes = True Then
             Form2.mFileTypes.Checked = True
             If blVirtualView Or blAddType Or blAllTypes Then
               itmx.SubItems(1) = fType(TTCnt)
             Else
               itmx.SubItems(1) = tType(TTCnt)
             End If
            itmx.SubItems(2) = CStr(aLastDate(TTCnt))
            itmx.SubItems(3) = aShortcut(TTCnt)
          Else
            itmx.SubItems(1) = CStr(aLastDate(TTCnt))
            itmx.SubItems(2) = aShortcut(TTCnt)
          End If
          
       Else
       
          itmx.SubItems(1) = fpathonly
          
          If blVirtualView = True And blVirtualViewTypes = True Or _
             blVirtualView = False And blOtherViewTypes = True Then
             Form2.mFileTypes.Checked = True
             If blVirtualView Or blAddType Or blAllTypes Then
               itmx.SubItems(2) = fType(TTCnt)
             Else
               itmx.SubItems(2) = tType(TTCnt)
             End If
            itmx.SubItems(3) = CStr(aLastDate(TTCnt))
            itmx.SubItems(4) = aShortcut(TTCnt)
            
          Else
            itmx.SubItems(2) = CStr(aLastDate(TTCnt))
            itmx.SubItems(3) = aShortcut(TTCnt)
          End If
         
       End If
       
    Else
    
       If blFolderSort = False Then
          itmx.SubItems(1) = CStr(aLastDate(TTCnt))
          itmx.SubItems(2) = aShortcut(TTCnt)
       Else
          itmx.SubItems(1) = fpathonly
          itmx.SubItems(2) = CStr(aLastDate(TTCnt))
          itmx.SubItems(3) = aShortcut(TTCnt)
       End If
   
    End If
    
    If blRefresh Then
      blRefresh = False
      If FormCnt = 1 Then
         Form1.Refresh
      ElseIf FormCnt < 4 Then
         Form2.Refresh
      ElseIf FormCnt < 50 Then          'smooth the loading
         If FormCnt Mod 4 = 0 Then
            Form2.Refresh
         End If
         If FormCnt Mod 10 = 0 Then
            If GetInputState() <> 0 Then
               DoEvents
            End If
         End If
      ElseIf FormCnt Mod 50 = 0 Then
         If GetInputState() <> 0 Then
            DoEvents
         End If
      End If
    End If
           
errlog:
   subErrLog ("modData: PopForm2")
End Sub

Private Sub CallDateTheFile1()
On Error GoTo errlog

   'This applies to syncshow and neversync - syncstart handled in LoadTTFileData
   If Form1.mAlwaysValidate.Checked = False Then
      If blTrueDates(TTCnt) = False Then
         CallDateTheFile2
         blTrueDates(TTCnt) = True
      Else
         If Len(aLastDate(TTCnt)) = 0 Then
            CallDateTheFile2
         End If
      End If
   Else
      CallDateTheFile2
   End If

errlog:
   subErrLog ("modData: CallDateTheFile1")
End Sub

Private Sub CallDateTheFile2()
On Error GoTo errlog
      
   blFolderType = False
   
   'This applies to syncshow and neversync - syncstart handled in LoadTTFileData
   If blNotValidated(TTCnt) = False Then
          
       If VerifyFolderType = True Then
          blFolderType = True

'          TT's Accessed (pseudo or true) and Modified dates are all the same for folders.
'          Can't use shortcut modification for pseudo-accessed because folder shortcut mod
'          dates never change. Can't use true folder accessed dates because always touched.
          
           Call DateTheFile(aSFileName(TTCnt))   'use true file
             
       Else
       
         If Form2.mAccessed.Checked = True Then
         
            If Form1.mTrueAccessDates.Checked = True Then
       
               Call DateTheFile(aSFileName(TTCnt))      'use true file
               
            Else
            
               Call DateTheFile(TTpath + aShortcut(TTCnt))    'use the TT folder\shortcut
            
            End If
            
         Else
                   
               Call DateTheFile(aSFileName(TTCnt))    'use the TT folder\shortcut
               
         End If
         
       End If

   ElseIf blNotValidated(TTCnt) = True Then
   
      Call DateTheFile(TTpath + aShortcut(TTCnt))    'use the TT folder\shortcut
      
   End If

errlog:
   subErrLog ("modData: CallDateTheFile2")
End Sub

Public Sub DateTheFile(sfilename As String, Optional DateCnt As Long)
On Error GoTo errlog

'   Debug.Print aLastAcc(TTCnt)
'   Debug.Print aLastMod(TTCnt)
'   Debug.Print aCreated(TTCnt)

   If DateCnt > 0 Then     'only when optional arg is passed from LoadTTFileData
      If blLoading = True Then      'only on syncshow startup
         If Form1.blReloading = False Then
            If Form1.mAlwaysValidate.Checked = False Then
               If blNoCache = False Then
                  If Form1.mSyncStartup.Checked = True Then
                     TTCnt = DateCnt
                  End If
               End If
            End If
         End If
      End If
   End If

   ReDim Preserve aLastAcc(TTCnt)
   ReDim Preserve aLastMod(TTCnt)
   ReDim Preserve aCreated(TTCnt)
   
   Dim FT_CREATE As FILETIME
   Dim FT_ACCESS As FILETIME
   Dim FT_WRITE As FILETIME
   Dim lFile As Long
   
'   If blFolderType = False Then
      lFile = CreateFile(sfilename, GENERIC_READ, FILE_SHARE_READ, 0&, OPEN_EXISTING, 0&, 0&)
'   Else
'      lFile = CreateFile(sFilename, 0&, FILE_SHARE_READ, 0&, _
'                   OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0&)    'only for folders - touches the file
'   End If
                   
   If lFile = -1 Then
   
      lFile = OpenFile(sfilename, OFS, 0&)   'alternate call for locked (open) files

   End If
   
   If lFile = -1 Then
   
        lFile = CreateFile(sfilename, 0&, FILE_SHARE_READ, 0&, _
                   OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0&)    'only for folders - touches the file
   End If

'   If lFile = -1 Then
'         lFile = CreateFile(shFilename, GENERIC_READ, FILE_SHARE_READ, 0&, OPEN_EXISTING, 0&, 0&) 'mainly for deleted
'   End If

   If lFile <> -1 Then

      'get the FILETIME info for the created,
      'accessed and last write info
      Call GetFileTime(lFile, FT_CREATE, FT_ACCESS, FT_WRITE)
      
      If Form2.mModified.Checked = True Then
         'MODIFIED
         aLastMod(TTCnt) = GetFileDate(FT_WRITE)
      ElseIf Form2.mCreated.Checked = True Then
         'CREATED
         aCreated(TTCnt) = GetFileDate(FT_CREATE)
      Else
         'ACCESSED   - accessed dates pose problems: dead files don't have accessed dates,
'         and getting Folder accessed dates touches them, so....
         If blNotValidated(TTCnt) = True Then
            'Deleted/Renamed, Networked, Removable, CD
            If Len(aLastMod(TTCnt)) > 0 Then
               aLastAcc(TTCnt) = aLastMod(TTCnt)
            Else
               aLastAcc(TTCnt) = GetFileDate(FT_WRITE)   'NOT AN ERROR! - I'm getting the modification date of the shortcut
            End If
         Else
            If blFolderType = False Then
            
               If Form1.mTrueAccessDates.Checked = True Then
            
                  aLastAcc(TTCnt) = GetFileDate(FT_ACCESS)
                  
               Else
               
                  aLastAcc(TTCnt) = GetFileDate(FT_WRITE)   'NOT AN ERROR! - I'm getting the modification date of the shortcut
               
               End If
               
            Else
               aLastAcc(TTCnt) = GetFileDate(FT_WRITE)   'NOT AN ERROR! - I'm getting the modification date of the shortcut
            End If
         End If
      End If
            
      PopaLastDate (TTCnt)
       
      LatestDateForType

   Else
   
      blNoAccess = True
   
   End If
   
errlog:
   Call CloseHandle(lFile)
   subErrLog ("Form1:DateTheFile")
End Sub

Private Sub PopaLastDate(Optional TTCnt As Long)
On Error GoTo errlog

   ReDim Preserve aLastDate((UBound(fType)))
   'Populate aLastDate with either aLastMod or aCreated
   If Form2.mModified.Checked = True Then
         If Len(aLastMod(TTCnt)) > 0 Then       'avoid blank dates when files are in use and locked
            aLastDate(TTCnt) = aLastMod(TTCnt)
         End If
   ElseIf Form2.mCreated.Checked = True Then
         If Len(aCreated(TTCnt)) > 0 Then
            aLastDate(TTCnt) = aCreated(TTCnt)
         End If
   ElseIf Form2.mAccessed.Checked = True Then
         If Len(aLastAcc(TTCnt)) > 0 Then
            aLastDate(TTCnt) = aLastAcc(TTCnt)
         End If
   End If
   
errlog:
   subErrLog ("modData: PopaLastDate")
End Sub

'This is what the "most recent" listview1 sort uses
Private Sub LatestDateForType()
On Error GoTo errlog

   If Form1.ListView1.ListItems.Count = 0 Then Exit Sub
'Debug.Print fType(TTCnt) + "  " + Str(TTCnt)
   If Form1.blNewType = True Then
       Form1.ListView1.ListItems(fType(TTCnt)).SubItems(2) = Trim$(aLastDate(TTCnt))
       Exit Sub
   End If
   If Len(Form1.ListView1.ListItems(fType(TTCnt)).SubItems(2)) = 0 Then
      If Len(Trim$(aLastDate(TTCnt))) > 0 Then
         Form1.ListView1.ListItems(fType(TTCnt)).SubItems(2) = Trim$(aLastDate(TTCnt))
      End If
   ElseIf CDate(aLastDate(TTCnt)) >= CDate(Form1.ListView1.ListItems(fType(TTCnt)).SubItems(2)) Then
       Form1.ListView1.ListItems(fType(TTCnt)).SubItems(2) = Trim$(aLastDate(TTCnt))
   End If
   
errlog:
   If Err.Number <> 35601 Then   'Element not found
      subErrLog ("modData: LatestDateforType")
   End If
End Sub

Public Function fnDateRange(fDate As Date) As Boolean
On Error GoTo errlog

   Dim sDate As String
   Dim SelDate1 As Date, SelDate2 As Date

   If Len(Trim$(fDate)) = 0 Then
      If Form4.cbDates.Index = 0 Then
         fnDateRange = True
         Exit Function
      End If
      Exit Function
   End If

   If InStr(fDate, ":") = 0 Then 'time doesn't need to be removed for evaluation
      sDate = Trim$(fDate)
   Else
      sDate = CStr(fDate)
'      sDate = Left$(sDate, InStr(sDate, " "))
      sDate = Trim$(sDate)
   End If
   
   If blVirtual = True Then
      If Form5.mIgnoreFilters.Checked = True Then
         fnDateRange = True
         Exit Function
      End If
   End If

   Select Case Form4.cbDates.ListIndex

     Case 0  'Everything
             fnDateRange = True
             Exit Function
             
     Case 1  'today
         SelDate1 = DateAdd("d", 0, Date)
         If sDate > SelDate1 Then
            fnDateRange = True
            Exit Function
         End If
         
     Case 2   'yesterday
         SelDate1 = DateAdd("d", -1, Date)
         If sDate > SelDate1 Then
            fnDateRange = True
            Exit Function
         End If
         
     Case 3   'last week
         SelDate1 = DateAdd("ww", -1, Date)
         If sDate > SelDate1 Then
            fnDateRange = True
            Exit Function
         End If
         
     Case 4   '2 weeks ago
         SelDate1 = DateAdd("ww", -2, Date)
         SelDate2 = DateAdd("ww", -1, Date)
         If sDate > SelDate1 And sDate < SelDate2 Then
            fnDateRange = True
            Exit Function
         End If
         
     Case 5   'Last 2 weeks
         SelDate1 = DateAdd("ww", -2, Date)
         If sDate > SelDate1 Then
            fnDateRange = True
            Exit Function
         End If
     
     Case 6   '3 weeks ago
        SelDate1 = DateAdd("ww", -3, Date)
        SelDate2 = DateAdd("ww", -2, Date)
        If sDate > SelDate1 And sDate < SelDate2 Then
            fnDateRange = True
            Exit Function
         End If
      
     Case 7   'Last 3 weeks
         SelDate1 = DateAdd("ww", -3, Date)
         If sDate > SelDate1 Then
            fnDateRange = True
            Exit Function
         End If
     
     Case 8   '4 weeks ago
        SelDate1 = DateAdd("ww", -4, Date)
        SelDate2 = DateAdd("ww", -3, Date)
        If sDate > SelDate1 And sDate < SelDate2 Then
            fnDateRange = True
            Exit Function
         End If
      
     Case 9   'last month
         SelDate1 = DateAdd("m", -1, Date)
         If sDate > SelDate1 Then
            fnDateRange = True
            Exit Function
         End If
      
      Case 10   '2 months ago
         SelDate1 = DateAdd("m", -2, Date)
         SelDate2 = DateAdd("m", -1, Date)
         If sDate > SelDate1 And sDate < SelDate2 Then
            fnDateRange = True
            Exit Function
         End If
         
      Case 11   '3 months ago
         SelDate1 = DateAdd("m", -3, Date)
         SelDate2 = DateAdd("m", -2, Date)
         If sDate > SelDate1 And sDate < SelDate2 Then
            fnDateRange = True
            Exit Function
         End If
         
      Case 12   'last 3 months
         SelDate1 = DateAdd("m", -3, Date)
         If sDate > SelDate1 Then
            fnDateRange = True
            Exit Function
         End If
         
      Case 13   '3 to 6 months ago
         SelDate1 = DateAdd("m", -6, Date)
         SelDate2 = DateAdd("m", -3, Date)
         If sDate > SelDate1 And sDate < SelDate2 Then
            fnDateRange = True
            Exit Function
         End If
      
      Case 14   'last 6 months
         SelDate1 = DateAdd("m", -6, Date)
         If sDate > SelDate1 Then
            fnDateRange = True
            Exit Function
         End If
         
      Case 15   '6 months to 9 months ago
         SelDate1 = DateAdd("m", -9, Date)
         SelDate2 = DateAdd("m", -6, Date)
         If sDate > SelDate1 And sDate < SelDate2 Then
            fnDateRange = True
            Exit Function
         End If
     
      Case 16   '9 months to 1 year ago
         SelDate1 = DateAdd("m", -12, Date)
         SelDate2 = DateAdd("m", -9, Date)
         If sDate > SelDate1 And sDate < SelDate2 Then
            fnDateRange = True
            Exit Function
         End If
         
      Case 17   '6 months to 1 year ago
         SelDate1 = DateAdd("m", -12, Date)
         SelDate2 = DateAdd("m", -6, Date)
         If sDate > SelDate1 And sDate < SelDate2 Then
            fnDateRange = True
            Exit Function
         End If
       
      Case 18   'Last Year
         SelDate1 = DateAdd("yyyy", -1, Date)
         If sDate > SelDate1 Then
            fnDateRange = True
            Exit Function
         End If
               
      Case 19   'More than 1 year ago
         SelDate1 = DateAdd("yyyy", -1, Date)
         If sDate < SelDate1 Then
            fnDateRange = True
            Exit Function
         End If
                 
    End Select
        
errlog:
subErrLog ("modData: fnDateRange")
End Function

Public Function ComparetoTTDate(fname As String) As Date
On Error GoTo errlog

   Dim FT_CREATE As FILETIME
   Dim FT_ACCESS As FILETIME
   Dim FT_WRITE As FILETIME
'   Dim dNewWrite As Date
   Dim lFile As Long

   lFile = CreateFile(fname, 0&, FILE_SHARE_READ, 0&, _
                      OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0&)
                      
'   lFile = CreateFile(fname, GENERIC_READ, READ_CONTROL, 0&, _
'                      OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0&)
   
'   lFile = OpenFile(fname, OFS, OF_READWRITE)
   
   If lFile > 0 Then
      
      Call GetFileTime(lFile, FT_CREATE, FT_ACCESS, FT_WRITE)
      
      ComparetoTTDate = GetFileDate(FT_WRITE)
         
   End If
   
errlog:
   Call CloseHandle(lFile)
   subErrLog ("modData: RecentFolderChange")
End Function

Public Sub ScrollBarCompensation()
On Error GoTo errlog

      With Form2.ListView2
         
         If AllorAddorSpecial = True Then
                                                
            If blFolderSort = True Then
            
                If VertScrollbar = True Then    'scrollbar compensation
                  If blVirtualViewTypes And blVirtualView Or _
                    blVirtualView = False And blOtherViewTypes Then
                    .ColumnHeaders(4).Width = .ColumnHeaders(4).Width - 250
                  Else
                    .ColumnHeaders(3).Width = .ColumnHeaders(3).Width - 250
                  End If
                End If
            
             Else
               
                If VertScrollbar = True Then
                   If blVirtualViewTypes And blVirtualView Or _
                    blVirtualView = False And blOtherViewTypes Then
                    .ColumnHeaders(3).Width = .ColumnHeaders(3).Width - 250
                  Else
                    .ColumnHeaders(2).Width = .ColumnHeaders(2).Width - 250
                  End If
               End If
            
            End If
            
'        All the other types when selected individually
'        ********************************************************
        Else
                    
            If blFolderSort = True Then
            
               If VertScrollbar = True Then
                  .ColumnHeaders(3).Width = .ColumnHeaders(3).Width - 250
               End If
            
            Else
            
               If VertScrollbar = True Then
                  .ColumnHeaders(2).Width = .ColumnHeaders(2).Width - 250
               End If
            
            End If
            
        End If
   End With
      
errlog:
   subErrLog ("modData: ScrollbarCompensation")
End Sub

Private Function VertScrollbar() As Boolean
   If (GetWindowLong(Form2.ListView2.hWnd, GWL_STYLE) And WS_VSCROLL) = False Then
      VertScrollbar = False
   Else
      VertScrollbar = True
   End If
End Function

Public Sub SaveTTInfo()
'On Error Resume Next

   If blForm4Loading = False Then

        SaveTypes
        
        SaveFileDetails (False)
        
        SaveFileDetails (True)
        
   Else
      blForm4Loading = False
   End If
   
   SaveVirtualFolders
   
'errlog:
'   subErrLog ("modData: SaveTTInfo")
End Sub

Private Sub SaveTypes()
On Error Resume Next

   Dim intF As Integer

   Dim lsts As Integer
   Dim sinfo As String
   
   If DoesFileExist(TTpath + "TTItems") Then
      Kill TTpath + "TTItems"
   Else
      DeleteSetting "TaskTracker", "TTI"
   End If
   
   intF = FreeFile
   Open TTpath + "TTItems" For Append Access Write As #intF
   With Form1.ListView1
      If .ListItems.Count = 0 Then Exit Sub
      For lsts = 1 To .ListItems.Count
         If Len(.ListItems(lsts).Text) > 0 Then
            sinfo = .ListItems(lsts).Text + "*/" + .ListItems(lsts).SubItems(1) + "*/" + .ListItems(lsts).SubItems(2)
                        
            Print #intF, sinfo
         
         End If
      Next lsts
   End With
   Close #intF
   
errlog:
   Close #intF
   subErrLog ("modData: SafeTypes")
End Sub

Private Sub SaveVirtualFolders()
On Error Resume Next

   Dim intF As Integer
   Dim lsts As Integer
   Dim sinfo As String
   
   If Form1.mVirtual.Checked = False Then
      Load Form5     'need to do this to persist virtual folders
   End If
      
   If DoesFileExist(TTpath + "TTVirtual") Then
      FileCopy TTpath + "TTVirtual", TTpath + "TTVirtual_bak"
      Kill TTpath + "TTVirtual"
   Else
      DeleteSetting "TaskTracker", "TTI\Virtual"
   End If
   
   intF = FreeFile
   Open TTpath + "TTVirtual" For Append Access Write As #intF
   
   With Form5.ListView3
      If .ListItems.Count = 0 Then Exit Sub
      For lsts = 1 To .ListItems.Count
         If Len(.ListItems(lsts).Text) > 0 Then
            sinfo = .ListItems(lsts).Text + "*/" + .ListItems(lsts).Key
            
'            SaveSetting "TaskTracker", "TTI\Virtual", ElimSpace(Str$(Trim$(lsts))), sinfo

            Print #intF, sinfo
         
         End If
      Next lsts
   End With
   
   Close #intF
   Exit Sub
   
errlog:
   Close #intF
   FileCopy TTpath + "TTVirtual_bak", TTpath + "TTVirtual"  'recover backup if an error
   subErrLog ("modData: SaveVirtualFolders")
End Sub

Private Sub SaveFileDetails(blSaveVirtual As Boolean)
On Error Resume Next
   Dim sdata As String
   Dim LastData As String
   Dim totc As Long
   Dim truetot As Long
   Dim wo As Long
   Dim vo As Long
   Dim intF As Integer
   Dim writeonce() As String
   Dim virtualonce() As String
   Dim blWriteOnce As Boolean
   Dim blVirtualOnce As Boolean
   Dim blDoOnce As Boolean
   Dim TTData As String
   
   RegCleanup
   
   If blSaveVirtual = False Then
      TTData = "TTData"
   Else
      TTData = "TTVData"
      If DoesFileExist(TTpath + TTData) Then
         FileCopy TTpath + "TTVData", TTpath + "TTVData_bak"
      End If
   End If
           
   ReDim Preserve writeonce(0)
   ReDim Preserve virtualonce(0)
   
   For totc = 0 To UBound(fType)
   
       ReDim Preserve writeonce(totc)
       If Len(aSFileName(totc)) > 0 Then
       
       '     Dupe check - too slow for shutdown
'            For ck = 0 To totc - 1
'               If LCase(aSFileName(totc)) = LCase(aSFileName(ck)) Then
'                  GoTo writeonce      'it's a duplicate
'                  Exit For
'               End If
'            Next
       
            If blSaveVirtual = False Then
            
               For wo = 0 To totc
                  If aSFileName(totc) = writeonce(wo) Then
                     blWriteOnce = True      'now need to check it's not a virtual item
                     Exit For
                  Else
                     blWriteOnce = False
                  End If
               Next wo
       
               If Left$(tType(totc), 14) <> "Virtual Folder" Then
                  writeonce(totc) = aSFileName(totc)
               End If
               
            End If
       
            sdata = aSFileName(totc) + "*/"
                                                                  
            If Len(fType(totc)) > 0 Then
               sdata = sdata + fType(totc) + "*/"
            Else
               sdata = sdata + "*/"
            End If
            
            If Len(tType(totc)) > 0 Then
               sdata = sdata + tType(totc) + "*/"
            Else
               sdata = sdata + "*/"
            End If
                        
            If Left$(tType(totc), 14) <> "Virtual Folder" Then
            
               If blSaveVirtual = False Then
                  If blWriteOnce = True Then
                     GoTo writeonce
                  End If
               Else
                  GoTo writeonce    'ignore if not a VF when saving VFs
               End If
               
            Else
            
               If blSaveVirtual = False Then    'ignore if a VF when not saving VFs
                  GoTo writeonce
               Else
                  For vo = 0 To UBound(virtualonce)   'need to prevent duplicates in same VF
                     If tType(totc) + aSFileName(totc) = virtualonce(vo) Then
                        blVirtualOnce = True
                        Exit For
                     Else
                        blVirtualOnce = False
                     End If
                  Next vo
                  
                  If blVirtualOnce = True Then
                     GoTo writeonce
                  End If
                  
                  vo = vo + 1
                  ReDim Preserve virtualonce(vo)
                  virtualonce(vo) = tType(totc) + aSFileName(totc)
                  
              End If
            End If
            
            sdata = sdata + Str$(CInt(blExists(totc))) + "*/" + Trim$(Str$(CInt(blNotValidated(totc)))) + "*/"
            
            '********************************
            'since out-of-subscript errors occur here regularly, isolate this and pop with delimiters
            sdata = sdata + aLastAcc(totc) + "*/" + _
                                                                  aLastMod(totc) + "*/" + _
                                                                  aCreated(totc) + "*/" + _
                                                                  aLastDate(totc)
            If Err.Number = 9 Then
                sdata = sdata + "*/*/"
            End If
            '********************************
                                                                  
            sdata = sdata + "*/" + aShortcut(totc)
                        
            If Len(sdata) > 0 Then  'guard against empty
               If sdata <> LastData Then  'bugfix
               
                  If blDoOnce = False Then
                     blDoOnce = True
                     'Only delete now...
                     If DoesFileExist(TTpath + TTData) Then
                        Kill TTpath + TTData
                     End If
                     intF = FreeFile
                     Open TTpath + TTData For Append Access Write As #intF
                  End If
                  
                  Print #intF, sdata
                  truetot = truetot + 1
               End If
            End If
            
            LastData = sdata
'       Else
'            Empties = Empties + 1
       End If
       
writeonce:
       
   Next totc
   
   Close #intF
   If blSaveVirtual = False Then
      SaveSetting "TaskTracker", "Settings", "TrueCnt", Str$(truetot)
   End If
   Exit Sub
   
errlog:
   Close #intF
   FileCopy TTpath + "TTVData_bak", TTpath + "TTVData"      'recover backup if an error
   If blSaveVirtual = False Then
      SaveSetting "TaskTracker", "Settings", "TrueCnt", Str$(truetot)
   End If
   subErrLog ("modData: SaveFileDetails")
End Sub


Private Sub RegCleanup()
On Error GoTo errlog

   Dim cr As New cRegistry

   With cr
      .ClassKey = HKEY_CURRENT_USER
      .ValueKey = "TaskTracker"
   
      .SectionKey = "Software\VB and VBA Program Settings\TaskTracker\TTInfo\FileData"
      If .KeyExists Then
         .DeleteKey
      End If

      .SectionKey = "Software\VB and VBA Program Settings\TaskTracker\TTInfo"
      If .KeyExists Then
         .DeleteKey
      End If
      
      .SectionKey = "Software\VB and VBA Program Settings\TaskTracker\TTI\FileData"
      If .KeyExists Then
         .DeleteKey
      End If
   
      .SectionKey = "Software\VB and VBA Program Settings\TaskTracker\TTI"
      If .KeyExists Then
         .DeleteKey
      End If
   End With
   
errlog:
   subErrLog ("modData: RegCleanup")
End Sub

Public Sub LoadTTFiles()
On Error Resume Next

   Dim iFiles As Integer
   Dim iLinks As Integer
   Dim iTypes As Integer
         
   If Len(GetSetting("TaskTracker", "Settings", "TrueCnt")) > 0 Then
      LastTrueCnt = Val(GetSetting("TaskTracker", "Settings", "TrueCnt"))
   End If
   
   iFiles = LinesInFile(TTpath + "TTData")
   iTypes = LinesInFile(TTpath + "TTItems")
   iLinks = Form1.FilesCountAll(TTpath, "*.lnk")
   TotalFiles = iFiles
   
   If iTypes = 0 Then                     'if zero types
      blScratchCache = True
   ElseIf iFiles / iLinks > 1.5 Then      'if cached files is >150% of TTDir shortcuts
      blScratchCache = True
      blNoCache = True
      Form1.blNewTTDir = True
      Form1.subStart
      Exit Sub
   ElseIf iLinks / iFiles > 1.5 Then      'if TTDir shortcuts is >150% of cached files
      blScratchCache = True
   ElseIf iFiles < LastTrueCnt Then       'if cached files are somehow fewer than recorded - ??
      blScratchCache = True
   ElseIf DoesFileExist(TTpath + "TTItems") = False Or _
      DoesFileExist(TTpath + "TTData") = False Then      'old version upgrade
      blScratchCache = True
   End If
      
   If blScratchCache = True Then
      blNoCache = True
      blInstantStarted = True
      blAbort = True
      Form1.Visible = True
      Form1.subStart
      Exit Sub
   End If
   
   Set Form1.ListView1.SmallIcons = Form1.ImageList1
   
   SetLoadTTFileData
   
errlog:
   Form1.Label1.Visible = True
   subErrLog ("modData: LoadTTFiles")
End Sub

Public Sub SetLoadTTFileData()
On Error Resume Next
   If DoesFileExist(TTpath + "TTVData") = True Then
   
      If blNoCache = True Then
         LoadTTFileData (True)
      Else
         LoadTTFileData (True)
         LoadTTFileData (False)
      End If
      
   Else
      blLoadLegacy = True
      LoadTTFileData (False)
   End If
errlog:
   subErrLog ("modData: SetLoadTTFileData")
End Sub

Private Sub LoadTTFileData(blLoadVirtual As Boolean)
On Error Resume Next

   Dim all As String
   Dim ft As Long
   Dim ftc As Integer
   Dim iFiles As Long
   Dim VS As String, FS As String
   Dim himgSmall As Long
   Dim img1 As ListImage
   Dim blTypeSet As Boolean
   Dim fTypeSet() As String
   Dim intF As Integer
   Dim NextLine As String
   Dim di As Integer
   Dim TTData As String
      
   blNoSplash = CBool(GetSetting("TaskTracker", "Settings", "NoSplash"))
   
   If blLoadVirtual = False Then
      TTData = "TTData"
   Else
      TTData = "TTVData"
   End If
   
   If blNoCache = False Then
      iFiles = LinesInFile(TTpath + "TTData")
   Else
      iFiles = UBound(fType)     'get ubound from uncached data when checking just for virtual folders
   End If
   
   If iFiles = 0 Then Exit Sub
   
'   Debug.Print Format$(Time, "ss")
   ReDim fTypeSet(0)

   intF = FreeFile()
   Open TTpath + TTData For Binary As #intF
   
   all = Input(LOF(intF), intF)
   
'  ****************************************
   Do Until Len(all) = 0
      NextLine = Left(all, InStr(all, Chr(13)) - 1)
      all = Right(all, Len(all) - Len(NextLine) - 2)
      
'     **********************
      If blLoadVirtual = False Then
         If blNoCache = False Then
            If Form1.mNoSplash.Checked = False Then
               If blRegStatus = True Or iDaysUsed < 15 Then
                  If fc > 0 Then
                     If fc Mod iFiles / 10 = 0 Then
                        Splash.Timer1_Timer
                     End If
                  End If
               End If
            End If
         End If
      End If
'     **********************
   
      fc = fc + 1
      
      FS = NextLine
      
'     ****************************************
'     Legacy Check for virtual folders
      If blLoadLegacy = True Then
         If blNoCache = True Then
            Dim vtest As String
            VS = (Mid$(NextLine, InStr(NextLine, "*/") + 1))
            VS = (Mid$(VS, InStr(VS, "*/") + 2))
            vtest = Left$(VS, InStr(VS, "*/") - 1)
            If Left$(vtest, 14) <> "Virtual Folder" Then
               GoTo trynext   'ignore
            End If
            If fc > UBound(fType) Then
               RedimPreserve (fc)
            Else
               RedimPreserve (UBound(fType) + 1)
            End If
            tType(fc) = Left$(VS, InStr(VS, "*/") - 1)
         Else
            RedimPreserve (fc)
         End If
      Else
         If fc > UBound(fType) Then
            RedimPreserve (fc)
         Else
            RedimPreserve (UBound(fType) + 1)
         End If
      End If
'     ****************************************
      
'     ****************************************
'     Parse the line
      aSFileName(fc) = Left$(FS, InStr(FS, "*/") - 1)
'     Dupe check - too time-consuming on startup
'      For ck = 0 To fc - 1
'         If LCase(aSFileName(ck)) = LCase(aSFileName(fc)) Then
'            GoTo trynext      'it's a duplicate
'            Exit For
'         End If
'      Next
      FS = (Mid$(FS, InStr(FS, "*/") + 2))
'      Debug.Print FS
      fType(fc) = Left$(FS, InStr(FS, "*/") - 1)
      FS = (Mid$(FS, InStr(FS, "*/") + 2))
'      Debug.Print FS
      PrevType(fc) = fType(fc)
      tType(fc) = Left$(FS, InStr(FS, "*/") - 1)
      FS = (Mid$(FS, InStr(FS, "*/") + 2))
'      Debug.Print FS
      blExists(fc) = CBool(Left$(FS, InStr(FS, "*/") - 1))
      FS = (Mid$(FS, InStr(FS, "*/") + 2))
'      Debug.Print FS
'      Debug.Print Left$(FS, InStr(FS, "*/") - 1)
      blNotValidated(fc) = CBool(Left$(FS, InStr(FS, "*/") - 1))
'      Debug.Print Str(blNotValidated(fc))
      FS = (Mid$(FS, InStr(FS, "*/") + 2))
'      MsgBox FS
      aLastAcc(fc) = Left$(FS, InStr(FS, "*/") - 1)
      FS = (Mid$(FS, InStr(FS, "*/") + 2))
'      MsgBox FS
      aLastMod(fc) = Left$(FS, InStr(FS, "*/") - 1)
      FS = (Mid$(FS, InStr(FS, "*/") + 2))
      aCreated(fc) = Left$(FS, InStr(FS, "*/") - 1)
      FS = (Mid$(FS, InStr(FS, "*/")) + 2)
'      MsgBox FS
      aLastDate(fc) = Left$(FS, InStr(FS, "*/") - 1)
      FS = (Mid$(FS, InStr(FS, "*/") + 2))
'      MsgBox FS
      If InStr(FS, "*/") > 0 Then
         aShortcut(fc) = Right$(FS, Len(FS) - InStr(FS, "*/") - 1)
      Else
         aShortcut(fc) = FS
      End If
'      Debug.Print aShortcut(fc)

      If blLoadVirtual = False Then
               
   '     ****************************************
         'if type is loaded don't load again
         blTypeSet = False
         For ftc = 0 To ft
            If fType(fc) = fTypeSet(ftc) Then
               blTypeSet = True
               Exit For
            End If
         Next ftc
         
         If blTypeSet = False Then  'load up each type image just once,
            himgSmall& = SHGetFileInfo(TTpath + aShortcut(fc), _
                          0&, shinfo, Len(shinfo), _
                          BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
            If himgSmall& <> 0 Then    'but make sure the image is loaded
               Call AddFileItemIcons(, himgSmall&)
               Set img1 = Form1.ImageList1.ListImages.Add(, fType(fc), Form1.pixSmall.Picture)
               ft = ft + 1
               ReDim Preserve fTypeSet(ft)
               fTypeSet(ft) = fType(fc)
            End If
         End If
   '     ****************************************
       
   '     ****************************************
         For di = 0 To UBound(SystemDrives)
         
            If Left$(aSFileName(fc), 1) = Left$(SystemDrives(di), 1) Then
               Exit For
            End If
            
            rf = rf + 1       'counter for nonsystem/network files
            
         Next di
         
         If fType(fc) = "Network Drive" Then
            If Form1.mNetwork.Checked = False Then
               fType(TTCnt) = tType(TTCnt)
            End If
         ElseIf fType(fc) = "Removable Drive" Then
            If Form1.mRemovable.Checked = False Then
               fType(TTCnt) = tType(TTCnt)
            End If
         ElseIf fType(fc) = "CD Drive" Then
            If Form1.mRemovable.Checked = False Then
               fType(TTCnt) = tType(TTCnt)
            End If
         ElseIf fType(fc) = "Deleted or Renamed" Then
            vf = vf + 1
         ElseIf fType(fc) = "Unknown" Then
            vf = vf + 1
         Else    'everything else
            vf = vf + 1
         End If
   '     ****************************************
      End If
      
'  This is required for Synchronize > On Startup option - but not with Always Validate Files
   '     ****************************************
      If Form1.mAlwaysValidate.Checked = False Then
         If Form1.mSyncStartup.Checked = True Then
            ReDim Preserve blTrueDates(UBound(fType))
            blTrueDates(fc) = True
            If blNotValidated(fc) = False Then
               
'               If Form2.mAccessed.Checked = True Then
'                  Call DateTheFile((TTpath + aShortcut(fc)), fc) 'shortcut Accessed time works better than file access
'               Else
                  Call DateTheFile((aSFileName(fc)), fc)
'               End If
         
            ElseIf blNotValidated(fc) = True Then
            
               Call DateTheFile((TTpath + aShortcut(fc)), fc)   'use the TT folder\shortcut's dates
               blTrueDates(TTpath) = True
            End If
         End If
      End If
   '     ****************************************
   
      If blNoSplash = True Then
         If Form1.mTaskbar.Checked = False And Form1.blNoSysTray = False Then
            If fc Mod 50 = 0 Then
               Form1.subSysTray
            End If
         End If
      End If
      If GetInputState() <> 0 Then
          DoEvents
      End If
trynext:
      
   Loop
'  ****************************************
   
   If blLoadVirtual = False Then
      If blNoCache = False Then
         TTCnt = fc    'only want this to happen when loading from cache
      End If
   End If
      
errlog:
   Close #intF
   Set img1 = Nothing
   subErrLog ("modData: LoadTTData")
End Sub

Private Sub RedimPreserve(fc As Long)
   ReDim Preserve aSFileName(fc)
   ReDim Preserve fType(fc)
   ReDim Preserve PrevType(fc)
   ReDim Preserve tType(fc)
   ReDim Preserve blExists(fc)
   ReDim Preserve blNotValidated(fc)
   ReDim Preserve aLastAcc(fc)
   ReDim Preserve aLastMod(fc)
   ReDim Preserve aCreated(fc)
   ReDim Preserve aLastDate(fc)
   ReDim Preserve aShortcut(fc)
End Sub

Public Sub LoadTTTypesData()
On Error Resume Next
   
   Dim lgx As ListItem
   Dim tf As String
   Dim DS As String
   Dim intF As Integer
         
   Form1.lbTemp.Visible = False
   Form1.InitListview1
   
   intF = FreeFile()
   Open TTpath + "TTItems" For Binary As #intF
   
   Do Until EOF(intF)

      Line Input #intF, DS

      tf = Left$(DS, InStr(DS, "*/") - 1)
 
      Set lgx = Form1.ListView1.ListItems.Add(, tf, tf)
      lgx.SmallIcon = Form1.ImageList1.ListImages(tf).Key
      If Err.Number = 35601 Then
        Form1.ListView1.ListItems.Remove (tf)
      End If
      If InStr(DS, "*/") > 0 Then
        DS = (Mid$(DS, InStr(DS, "*/") + 2))
        lgx.SubItems(1) = Left$(DS, InStr(DS, "*/") - 1)
        DS = (Mid$(DS, InStr(DS, "*/") + 2))
        lgx.SubItems(2) = Left$(DS, Len(DS))
      End If
      If GetInputState() <> 0 Then
          DoEvents
      End If
   Loop
             
errlog:
   Close #intF
   subErrLog ("modData: LoadTTTypesData")
End Sub

Public Sub SyncFileIcons()
On Error Resume Next
   Dim fn As Integer
   Dim scnt As Integer
   Dim iShortcuts As Integer
   Dim himgSmall As Long
   Dim img2 As ListImage
   Dim StartTime As Integer, EndTime As Integer
   Dim OneTenth As Long
   Dim IgnoreSync As Boolean

   iShortcuts = UBound(aShortcut)
   
   OneTenth = iShortcuts / 10
   
   If Form1.mSyncStartup.Checked = True Then

      For fn = 0 To iShortcuts
           
        StartTime = Format$(Time, "ss")

        If blExists(fn) = True And blNotValidated(fn) = False Then
           If IsLocalFile(aSFileName(fn)) = True Then
               AddFileItemIcons (aSFileName(fn))
               Set img2 = Form1.ImageList2.ListImages.Add(, aShortcut(fn), Form1.pixSmall.Picture)
           End If
        Else
           AddFileItemIcons (TTpath + aShortcut(fn))
'            Set img2 = Form1.ImageList1.ListImages(fType(fn)).ExtractIcon
           Set img2 = Form1.ImageList2.ListImages.Add(, aShortcut(fn), Form1.pixSmall.Picture)
        End If
        
        EndTime = Format$(Time, "ss")
        
'       Delayed startup opt-out
        If blUseNeverSync = False And IgnoreSync = False Then
          If Form1.WindowState = vbNormal Then
             If Command$ <> "-m" Then
                 If EndTime - StartTime > 1 Then   'if more than a 1 second delay caused by shfgetfileinfo, use imagelist1 icons
                    Dim Response As Integer
                    Response = MsgBox("Loading was delayed because some files could not be accessed. " + vbNewLine + _
                     "Click Yes if you would like to continue loading, or click No to finish loading without synchronizing " + vbNewLine + _
                     "all file data, including icons." + vbNewLine + vbNewLine + _
                     "(If you see this message often, change your Preferences to: Performance " + vbNewLine + _
                     "Settings > Synchronization > On First Showing.)", 52)
                 End If
                 If Response = vbNo Then
                     blJustRegd = True
                     blUseNeverSync = True
                     GoTo UseNeverSync
                 Else
                     IgnoreSync = True
                 End If
             End If
          End If
        End If

        If blNoCache = False Then
            
            If fn Mod OneTenth = 0 Then
               scnt = scnt + 1
               If scnt < 10 Then
                  If blNoSplash = False Then
                     Splash.txLoading = " Synchronizing... " + vbNewLine + " " + Str(scnt * 10) + "%"
                     Splash.Refresh
                  Else  'for systray
                     nid.szTip = "TaskTracker - Synchronizing..." & Str$(Trim$(scnt * 10)) + "%" & vbNullChar
                     Shell_NotifyIcon NIM_MODIFY, nid
                     Shell_NotifyIcon NIM_ADD, nid
                  End If
               End If
            End If
        End If
      
        'Form1 is not usually visible when this happens
        If fn Mod (iShortcuts \ 100) = 0 Then

            Form1.StatusBar1.Panels(1).Text = "Synchronizing data... " + Str$(Trim$(Int(fn / iShortcuts * 100))) + "%"

        End If

       Next fn

    End If
        
UseNeverSync:

    If Form1.mNeverSync.Checked = True Or blUseNeverSync = True Then
    
       For fn = 0 To iShortcuts
           Set img2 = Form1.ImageList2.ListImages.Add(, aShortcut(fn), Form1.ImageList1.ListImages(fType(fn)).ExtractIcon)
           
           If Err.Number = 35603 Then
               Set img2 = Form1.ImageList2.ListImages.Add(, aShortcut(fn), Form1.ImageList1.ListImages("Unknown").ExtractIcon)
           End If
       Next fn
       
    End If
    
errlog:
   IgnoreSync = False
   Set img2 = Nothing
   subErrLog ("modData: SyncFileIcons")
End Sub

Private Sub GetTrueIcon(cf As Long)
On Error GoTo errlog

   Dim img2 As ListImage
   
   If blExists(cf) = True And blNotValidated(cf) = False Then
   
      If IsLocalFile(aSFileName(cf)) = True Then
      
          AddFileItemIcons (aSFileName(cf))
                    
          If blVirtual = False Then
            Set img2 = Form1.ImageList2.ListImages.Add(, aShortcut(cf), Form1.pixSmall.Picture)
          Else
            If DoesFileExist(TTpath + aShortcut(cf)) = False Then
               Set img2 = Form1.ImageList2.ListImages.Add(, aShortcut(cf) + tType(cf), LoadResPicture("UNK", 1))
            Else
               Set img2 = Form1.ImageList2.ListImages.Add(, aShortcut(cf) + tType(cf), Form1.pixSmall.Picture)
            End If
          End If
          
      Else
         GoTo Remote
      End If
      
   Else
Remote:
      If IsDriveConnected(aSFileName(cf)) = True Then
      
         AddFileItemIcons (TTpath + aShortcut(cf))

         If Form1.mAlwaysValidate.Checked = True Then
         
             Call RestoreFileType(aSFileName(cf), cf)
             
         End If
         
         If blVirtual = False Then
            Set img2 = Form1.ImageList2.ListImages.Add(, aShortcut(cf), Form1.pixSmall.Picture)
         Else
            If DoesFileExist(TTpath + aShortcut(cf)) = False Then
               Set img2 = Form1.ImageList2.ListImages.Add(, aShortcut(cf) + tType(cf), LoadResPicture("UNK", 1))
            Else
               Set img2 = Form1.ImageList2.ListImages.Add(, aShortcut(cf) + tType(cf), Form1.pixSmall.Picture)
            End If
         End If
         
      Else
      
         If blVirtual = False Then
            If tType(cf) = "Unknown" Then
               If MaybeAFolder(aShortcut(cf)) Then
                  Set img2 = Form1.ImageList2.ListImages.Add(, aShortcut(cf), Form1.ImageList1.ListImages("File Folder").ExtractIcon)
                  fType(cf) = "File Folder"
                  tType(cf) = "File Folder"
               Else
                  Set img2 = Form1.ImageList2.ListImages.Add(, aShortcut(cf) + tType(cf), LoadResPicture("UNK", 1))
               End If
            Else
               If Len(tType(cf)) > 0 Then
                  Set img2 = Form1.ImageList2.ListImages.Add(, aShortcut(cf), Form1.ImageList1.ListImages(tType(cf)).ExtractIcon)
               Else
                  Set img2 = Form1.ImageList2.ListImages.Add(, aShortcut(cf) + tType(cf), LoadResPicture("UNK", 1))
               End If
            End If
         Else
            If DoesFileExist(TTpath + aShortcut(cf)) = False Then
               Set img2 = Form1.ImageList2.ListImages.Add(, aShortcut(cf) + tType(cf), LoadResPicture("UNK", 1))
            Else
geticon:
               AddFileItemIcons (TTpath + aShortcut(cf))
               Set img2 = Form1.ImageList2.ListImages.Add(, aShortcut(cf) + tType(cf), Form1.pixSmall.Picture)
            End If
         End If
         
      End If
   End If
   
errlog:
   If Err.Number = 35601 Then
      Set img2 = Form1.ImageList2.ListImages.Add(, aShortcut(cf), LoadResPicture("UNK", 1))
   End If
'   blUnknownVirtual = False
   Set img2 = Nothing
   If Err.Number <> 35601 Then
      subErrLog ("modData: GetTrueIcon")
   End If
End Sub

Private Function IsLocalFile(fpath) As Boolean
Dim ni As Integer
   For ni = 0 To UBound(SystemDrives)
      If Len(SystemDrives(ni)) > 0 Then
         If Left$(fpath, 3) = SystemDrives(ni) Then
            IsLocalFile = True
            Exit Function
         End If
      End If
   Next ni
End Function

Private Function IsDriveConnected(fpath) As Boolean
On Error GoTo errlog
Dim lct As Integer

   If Form1.mAlwaysValidate.Checked = False Then
      If Left$(fpath, 2) = "\\" Then
         IsDriveConnected = False
         Exit Function
      End If
   Else
      IsDriveConnected = True    'assume
   End If

   If ict > 0 Then
      For lct = 0 To UBound(IsConnected)
         If Left$(fpath, 3) = IsConnected(lct) Then
            IsDriveConnected = True
            Exit Function
         End If
      Next
   End If
   If nct > 0 Then
      For lct = 0 To UBound(NotConnected)
         If Left$(fpath, 3) = NotConnected(lct) Then
            Exit Function
         End If
      Next
   End If
   'create connected and not-connected arrays, so each drive is checked only once for each popfrommem
   If DriveExists(Left$(fpath, 3)) Then
      ReDim Preserve IsConnected(ict)
      IsConnected(ict) = Left$(fpath, 3)
      ict = ict + 1
      IsDriveConnected = True
   Else
      ReDim Preserve NotConnected(nct)
      NotConnected(nct) = Left$(fpath, 3)
      nct = nct + 1
   End If
   
errlog:
   subErrLog ("modData: IsDriveConnected")
End Function

Private Function DriveExists(drvName As String) As Boolean
'use diskspace because it doesn't time out
   Dim nS As Long
   Dim nB As Long
   Dim nF As Long
   Dim nT As Long
   If GetDiskFreeSpace(drvName, nS, nB, nF, nT) <> 0 Then
      DriveExists = True
   End If
End Function

Public Function MaybeAFolder(sf As String) As Boolean
   sf = Left$(sf, Len(sf) - 3)
   If Len(sf) - InStrRev(sf, ".") < 5 Then      'just assume it's a folder
      MaybeAFolder = True
   End If
End Function

Private Function ExtraColumn()
   If Form1.ListView1.SelectedItem.Text = "Deleted or Renamed" Or _
      Form1.ListView1.SelectedItem.Text = "Network Drive" Or _
      Form1.ListView1.SelectedItem.Text = "Removable Drive" Or _
      Form1.ListView1.SelectedItem.Text = "CD Drive" Then
      ExtraColumn = True
   Else
      ExtraColumn = False
   End If
End Function

'restore file info - usually "unknown" or virtual ones - this is a messy afterthought...
Public Function RestoreFileType(filename As String, cf As Long) As String
On Error GoTo errlog
   If SHGetFileInfo(aSFileName(cf), 0&, _
       shinfo, Len(shinfo), _
       SHGFI_TYPENAME Or SHGFI_USEFILEATTRIBUTES) Then
   
      tType(cf) = TrimNull(shinfo.szTypeName)
      If Form1.mNetwork.Checked = True Then
           fType(cf) = "Network Drive"
      Else
           fType(cf) = tType(cf)
      End If
      If Form1.mRemovable.Checked = True Then
           fType(cf) = "Removable Drive"
      Else
           fType(cf) = tType(cf)
      End If
'      If Form1.mRemovable.Checked = True Then     -oh well, who the hell cares...
'        fType(cf) = "CD Drive"
'      Else
'        fType(cf) = tType(cf)
'      End If
      RestoreFileType = tType(cf)
    End If
errlog:
   subErrLog ("modData: RestoreFileType")
End Function

