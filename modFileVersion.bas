Attribute VB_Name = "modFileVersion"
Option Explicit

Public Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersion As Long     'e.g. 0x00000042 = "0.42"
   dwFileVersionMS As Long    'e.g. 0x00030075 = "3.75"
   dwFileVersionLS As Long    'e.g. 0x00000031 = "0.31"
   dwProductVersionMS As Long 'e.g. 0x00030010 = "3.10"
   dwProductVersionLS As Long 'e.g. 0x00000031 = "0.31"
   dwFileFlagsMask As Long    'e.g. 0x3F for version "0.42"
   dwFileFlags As Long        'e.g. VFF_DEBUG Or VFF_PRERELEASE
   dwFileOS As Long           'e.g. VOS_DOS_WINDOWS16
   dwFileType As Long         'e.g. VFT_DRIVER
   dwFileSubtype As Long      'e.g. VFT2_DRV_KEYBOARD
   dwFileDateMS As Long       'e.g. 0
   dwFileDateLS As Long       'e.g. 0
End Type

Public Declare Function GetFileVersionInfoSize Lib "version.dll" _
   Alias "GetFileVersionInfoSizeA" _
  (ByVal lptstrFilename As String, _
   lpdwHandle As Long) As Long

Public Declare Function GetFileVersionInfo Lib "version.dll" _
   Alias "GetFileVersionInfoA" _
  (ByVal lptstrFilename As String, _
   ByVal dwHandle As Long, _
   ByVal dwLen As Long, _
   lpData As Any) As Long
   
Public Declare Function VerQueryValue Lib "version.dll" _
   Alias "VerQueryValueA" _
  (pBlock As Any, _
   ByVal lpSubBlock As String, _
   lplpBuffer As Any, nVerSize As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (Destination As Any, _
   Source As Any, _
   ByVal Length As Long)

Public Declare Function GetWindowsDirectory Lib "kernel32" _
   Alias "GetWindowsDirectoryA" _
  (ByVal lpBuffer As String, _
   ByVal nSize As Long) As Long

Public Declare Function FindClose Lib "kernel32" _
  (ByVal hFindFile As Long) As Long
     
Public Function HiWord(dw As Long) As Long
  
   If dw And &H80000000 Then
         HiWord = (dw \ 65535) - 1
   Else: HiWord = dw \ 65535
   End If
    
End Function
  

Public Function LoWord(dw As Long) As Long
  
   If dw And &H8000& Then
         LoWord = &H8000& Or (dw And &H7FFF&)
   Else: LoWord = dw And &HFFFF&
   End If
    
End Function

Public Function GetFileVersion(sDriverFile As String) As String
   
   Dim FI As VS_FIXEDFILEINFO
   Dim sBuffer() As Byte
   Dim nBufferSize As Long
   Dim lpBuffer As Long
   Dim nVerSize As Long
   Dim nUnused As Long
   Dim tmpVer As String
   
   nBufferSize = GetFileVersionInfoSize(sDriverFile, nUnused)
   
   If nBufferSize > 0 Then
   
      ReDim sBuffer(nBufferSize)
      Call GetFileVersionInfo(sDriverFile, 0&, nBufferSize, sBuffer(0))
      
      Call VerQueryValue(sBuffer(0), "\", lpBuffer, nVerSize)
      Call CopyMemory(FI, ByVal lpBuffer, Len(FI))
     
     'extract the file version from the FI structure
      tmpVer = Format$(HiWord(FI.dwFileVersionMS)) & "." & _
               Format$(LoWord(FI.dwFileVersionMS), "00") & "."
         
      If FI.dwFileVersionLS > 0 Then
         tmpVer = tmpVer & Format$(HiWord(FI.dwFileVersionLS), "00") & "." & _
                           Format$(LoWord(FI.dwFileVersionLS), "00")
      Else
         tmpVer = tmpVer & Format$(FI.dwFileVersionLS, "0000")
      End If
      
      End If
   
   GetFileVersion = tmpVer
   
End Function

Public Sub CheckFileVersions()
   If FileExists(GetWinDir + "\system32\COMCTL32.OCX") = True Then
      If Val(GetFileVersion(GetWinDir + "\system32\COMCTL32.OCX")) < 6 Then
'         MsgBox "TaskTracker requires a newer version of a Windows system file than is present on your system." _
'         + vbNewLine + "You can install this file from http://tasktracker.wordwisesolutions.com/support/.", vbCritical
      Else
         Exit Sub    'COMCTL32.OCX is OK
      End If
   Else
'      MsgBox "TaskTracker requires a Windows system file that is not present on your system." _
'      + vbNewLine + "You can install this file from http://tasktracker.wordwisesolutions.com/support/.", vbCritical
   End If
   
   fnOpenURL ("http://tasktracker.wordwisesolutions.com/support/")
   
   MsgBox "TaskTracker requires a Windows system file that is out of date or not present on your system." _
   + vbNewLine + "You can install this file from http://tasktracker.wordwisesolutions.com/support/.", vbCritical
   
   End
   
End Sub

'I don't fgetspecialfolder, as elsewhere, because this would start form1
Private Function GetWinDir() As String

    Dim nSize As Long
    Dim tmp As String
   
   'pad the string for the return value and
   'set nSize equal to the size of the string
    tmp = Space$(256)
    nSize = Len(tmp)

   'call the API
    Call GetWindowsDirectory(tmp, nSize)
    
   'trim off the trailing null added by the API
    GetWinDir = TrimNull(tmp)
    
End Function

