Attribute VB_Name = "modFileFuncs"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ©  2003 Michael M. Ross, Wordwise Solutions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260

Private Declare Function FindFirstFile Lib "kernel32" _
   Alias "FindFirstFileA" _
  (ByVal lpFileName As String, _
   lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal _
   lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
  (ByVal hwndOwner As Long, ByVal nFolder As Long, _
  pidl As ITEMIDLIST) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
  Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
  ByVal pszPath As String) As Long

Public Declare Function FindClose Lib "kernel32" _
  (ByVal hFindFile As Long) As Long

Public Type SH_ITEMID
    cb As Long
    abID As Byte
End Type

Public Type ITEMIDLIST
    mkid As SH_ITEMID
End Type

'Private Type HD_ITEM
'   mask As Long
'   cxy As Long
'   pszText As String
'   hbm As Long
'   cchTextMax As Long
'   fmt As Long
'   lParam As Long
'   iImage As Long
'   iOrder As Long
'End Type

Private Const LVM_FIRST = &H1000
'Private Const HDI_IMAGE = &H20
'Private Const HDI_FORMAT = &H4
'
'Private Const HDM_FIRST = &H1200
'Private Const HDM_SETITEM = (HDM_FIRST + 4)
'Private Const HDM_SETIMAGELIST = (HDM_FIRST + 8)
'Private Const HDM_GETIMAGELIST = (HDM_FIRST + 9)

'Private Const HDF_BITMAP_ON_RIGHT = &H1000
'Private Const HDF_BITMAP = &H2000
'Private Const HDF_IMAGE = &H800
'Private Const HDF_STRING = &H4000

Public Const CSIDL_RECENT = 8
Public Const CSIDL_PROFILE = 40
Public Const CSIDL_LOCAL_APPDATA = 28

Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10

Public Const FO_DELETE = &H3
Public Const FOF_SILENT = &H4
Public Const FOF_NOCONFIRMATION = &H10
 
Public Type SHFILEOPSTRUCT
    hWnd      As Long
    wFunc     As Long
    pFrom     As String
    pTo       As String
    fFlags    As Integer
    fAborted  As Boolean
    hNameMaps As Long
    sProgress As String
End Type

Public Declare Function SHFileOperation Lib "shell32.dll" _
    Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
    
'Public Type WINDOWPLACEMENT
'  Length            As Long
'  flags             As Long
'  showCmd           As Long
'  ptMinPosition     As POINTAPI
'  ptMaxPosition     As POINTAPI
'  rcNormalPosition  As RECT
'End Type
'
'Public Const SW_SHOWNORMAL = 1
'Public Const SW_SHOWMINIMIZED = 2
'Public Const SW_SHOWMAXIMIZED = 3
'Public Const SW_SHOWNOACTIVATE = 4

Public Function fGetSpecialFolder(CSIDL As Long) As String
On Error GoTo errlog

   Dim sPath As String
   Dim IDL As ITEMIDLIST
    '
    ' Retrieve info about system folders such as the
    ' "Desktop" folder. Info is stored in the IDL structure.
    '
    fGetSpecialFolder = vbNullString
    If SHGetSpecialFolderLocation(Form1.hWnd, CSIDL, IDL) = 0 Then
        '
        ' Get the path from the ID list, and return the folder.
        '
        sPath = Space$(MAX_PATH)
        If SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath) Then
            fGetSpecialFolder = Left$(sPath, InStr(sPath, vbNullChar) - 1) & "\"
        End If
    End If
    
errlog:
   subErrLog ("modFileFuncs: fGetSpecialFolder")
End Function

Public Function DoesFileExist(ByVal sSource As String) As Boolean
On Error GoTo errlog

   Dim lFile As Long

   lFile = CreateFile(sSource, 0&, FILE_SHARE_READ, 0&, _
                      OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0&)
   
   If lFile > 0 Then
      DoesFileExist = True
   End If
   
   Call CloseHandle(lFile)
   
errlog:
subErrLog ("modFileFuncs: DoesFileExist")
End Function

Public Function DoesFolderExist(ByVal sFolder As String) As Boolean
On Error GoTo errlog

   Dim hfile As Long
   Dim WFD As WIN32_FIND_DATA

   sFolder = UnQualifyPath(sFolder)

   hfile = FindFirstFile(sFolder, WFD)

   DoesFolderExist = (hfile <> INVALID_HANDLE_VALUE) And _
                  (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY)
                  
   Call FindClose(hfile)
   
errlog:
subErrLog ("modFileFuncs: DoesFolderExist")
End Function

Private Function UnQualifyPath(ByVal sFolder As String) As String

  'trim and remove any trailing slash
   sFolder = Trim$(sFolder)
   
   If Right$(sFolder, 1) = "\" Then
         UnQualifyPath = Left$(sFolder, Len(sFolder) - 1)
   Else: UnQualifyPath = sFolder
   End If
   
End Function

Public Function fnResolveLink(strTestname As String) As String
On Error GoTo errlog
  
   Dim LinkShell As New WshShell
   Dim LinkShortCut As WshShortcut 'new ???
   Set LinkShortCut = LinkShell.CreateShortcut(strTestname)

   fnResolveLink = LinkShortCut.TargetPath
        
   Set LinkShell = Nothing
   Set LinkShortCut = Nothing

   Exit Function
errlog:
   'old version of ie (wshom.ocx is old), so try this
   fnResolveLink = TryAlternate(strTestname)
   subErrLog ("modFileFuncs:fnResolveLink")
End Function

Private Function TryAlternate(ByVal strTestname As String) As String
On Error GoTo errlog

' doesn't work for everyone
  Dim sh As New Shell
  Dim fold As Folder
  Dim folditm As FolderItem
  Dim SLO As ShellLinkObject
  Set fold = sh.NameSpace(ssfDESKTOP)
  Set folditm = fold.ParseName(strTestname)
  If folditm.IsLink Then
    Set SLO = folditm.GetLink
    TryAlternate = SLO.path
  Else
    TryAlternate = vbNullString
  End If
  
errlog:
   subErrLog ("modFileFuncs:TryAlternate")
End Function

Public Sub AddFileItemIcons(Optional filename As String, Optional himgSmall As Long)
On Error GoTo errlog

   Dim r As Long
   
   If Len(filename) > 0 Then     'if filename is supplied, need to acquire image
   
      himgSmall& = SHGetFileInfo(filename, _
                             0&, shinfo, Len(shinfo), _
                             BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
   End If
   
   With Form1
   
     .pixSmall.Picture = LoadPicture()
     
     r& = ImageList_Draw(himgSmall&, shinfo.iIcon, .pixSmall.hDC, 0, 0, ILD_TRANSPARENT)
     .pixSmall.Picture = .pixSmall.Image
     
   End With
    
errlog:
    subErrLog ("modFileFuncs:AddFileItemIcons")
End Sub

Public Sub SetAutoRun()
On Error GoTo errlog

   Dim sExe As String
   Dim cr As New cRegistry
   
   sExe = App.path
   If (Right$(sExe, 1) <> "\") Then sExe = sExe & "\"
   sExe = sExe & "TaskTracker.exe"
   
   With cr
          
      If App.EXEName = "TaskTracker" Then
         .ClassKey = HKEY_CURRENT_USER
         .SectionKey = "Software\Microsoft\Windows\CurrentVersion\Run"
         .ValueKey = "TaskTracker"
         .ValueType = REG_SZ
         If Form1.mNorm.Checked = True Then
            .Value = Chr$(34) + sExe + Chr$(34)
         ElseIf Form1.mMin.Checked = True Then
            .Value = Chr$(34) + sExe + Chr$(34) + " -m"
         End If
      End If
      
       'delete, just in case
      .ClassKey = HKEY_LOCAL_MACHINE
      .SectionKey = "Software\Microsoft\Windows\CurrentVersion\Run"
      .ValueKey = "tasktracker"
      .DeleteValue
      
   End With
   
errlog:
   subErrLog ("modFileFuncs:SetAutoRun")
End Sub

Public Sub ClearAutoRun()
On Error GoTo errlog

   Dim cr As New cRegistry
   cr.ClassKey = HKEY_CURRENT_USER
   cr.SectionKey = "Software\Microsoft\Windows\CurrentVersion\Run"
   cr.ValueKey = "TaskTracker"
   On Error Resume Next
   cr.DeleteValue
   Err.Clear
        
errlog:
   subErrLog ("modFileFuncs:ClearAutoRun")
End Sub

Public Sub fnOpenURL(ByVal URL As String)

   Call ShellExecute(0, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
   
End Sub

Public Function TrimNull(ByVal item As String) As String
On Error GoTo errlog

    Dim pos As Integer
    
    pos = InStr(item, Chr$(0))
    If pos Then item = Left$(item, pos - 1)
    TrimNull = item
    
errlog:
subErrLog ("modeFileFuncs:TrimNull")
End Function

'Public Function ElimSpace(ByVal item As String)
'  If Left$(item, 1) = Chr$(32) Then
'    ElimSpace = Right$(item, Len(item) - 1)
'  Else
'    ElimSpace = item
'  End If
'End Function

Public Function LinesInFile(ByVal file_name As String) As _
    Long
On Error GoTo errlog:
Dim fnum As Integer
Dim lines As Long
Dim one_line As String

    fnum = FreeFile
    Open file_name For Input As #fnum
    Do While Not EOF(fnum)
        Line Input #fnum, one_line
        lines = lines + 1
    Loop
    Close #fnum

    LinesInFile = lines
errlog:
subErrLog ("modeFileFuncs:LinesInFile")
End Function

Public Function EnCrypt(ByVal Text As String, ByVal Key As String) As String
   Dim i As Long, KeyLen As Long, KeyPtr As Long, KeyChr As Integer
   Dim numcrypt As String, TextChr As Integer
   
   KeyLen = Len(Key)
      
   For i = 1 To Len(Text)
   
      TextChr = Asc(Mid(Text, i, 1))
      KeyChr = Asc(Mid(Key, KeyPtr + 1, 1))
      KeyPtr = ((KeyPtr + 1) Mod KeyLen)
      numcrypt = numcrypt + Trim(Str(TextChr + KeyChr)) + " "
      
   Next i
   
   EnCrypt = numcrypt
End Function

Public Function DeCrypt(ByVal nText As String, ByVal Key As String) As String
   Dim i As Long, KeyLen As Long, KeyPtr As Long, KeyChr As Integer
   Dim uncrypt As String, lastnum As Integer, charnum As Integer
   
   lastnum = 0
   
   KeyLen = Len(Key)
   
   For i = 1 To Len(nText)
   
      If Mid(nText, i, 1) = " " Then
         lastnum = lastnum + 1
         charnum = Mid(nText, lastnum, i - lastnum)
         If charnum > 0 Then
            KeyChr = Asc(Mid(Key, KeyPtr + 1, 1))
            uncrypt = uncrypt + Chr$(charnum - KeyChr)
         End If
         KeyPtr = ((KeyPtr + 1) Mod KeyLen)
         lastnum = i
      End If
   
   Next i
   
   DeCrypt = uncrypt
End Function

Public Function VerifyFolderType() As Boolean
On Error GoTo errlog

   VerifyFolderType = False
   If sType = "File Folder" Then
      VerifyFolderType = True
   Else
      If sType <> "Deleted or Renamed" And sType <> "Network Drive" And _
         sType <> "Removable Drive" And sType <> "CD Drive" Then
         If DoesFolderExist(Form2.fnSelectedItemsPath(Form2.ListView2.SelectedItem.Index)) = True Then
            VerifyFolderType = True
            sRegType = vbNullString
         End If
      End If
   End If
   
errlog:
   subErrLog ("modFileFuncs:VerifyFolderType")
End Function

Public Function AllorAddorSpecial() As Boolean
On Error GoTo errlog:

'  Test for All or Add + Special types
'  ********************************************************
   If blAllTypes = True Or blAddType = True Or blVirtualView = True Or _
       blThisType = True And blExtraCol Then
       
       AllorAddorSpecial = True
    Else
       AllorAddorSpecial = False
    End If
    
errlog:
   subErrLog ("modFileFuncs:AllorAddorSpecial")
End Function


'Public Sub SetHeaderIcon(colNo As Integer, _
'                          imgIconNo As Byte, _
'                          Optional showImage As Long)
'
'   Dim hHeader As Long
'   Dim HD As HD_ITEM
'
'  'get a handle to the listview header component
'   hHeader = SendMessage(Form2.ListView2.hwnd, LVM_GETHEADER, 0, ByVal 0)
'
'  'set up the required structure members
'   With HD
'      .mask = HDI_IMAGE Or HDI_FORMAT
'      .pszText = Form2.ListView2.ColumnHeaders(colNo + 1).Text
'
'       If showImage Then
'         .fmt = HDF_STRING Or HDF_IMAGE Or HDF_BITMAP_ON_RIGHT
'         .iImage = imgIconNo
'       Else
'         .fmt = HDF_STRING
'      End If
'
'
'   End With
'
'  'modify the header
'   Call SendMessage(hHeader, HDM_SETITEM, colNo, HD)
'
'End Sub


