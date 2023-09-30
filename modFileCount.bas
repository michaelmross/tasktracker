Attribute VB_Name = "modFileCount"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2003 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const INVALID_HANDLE_VALUE = -1
'Public Const MAX_PATH = 260
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800

Public Type SECURITY_ATTRIBUTES
   nLength As Long
   lpSecurityDescriptor As Long
   bInheritHandle As Long
End Type

Public Declare Function CreateDirectory Lib "kernel32" _
   Alias "CreateDirectoryA" _
  (ByVal lpPathName As String, _
   lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long

Public Declare Function CopyFile Lib "kernel32" _
   Alias "CopyFileA" _
  (ByVal lpExistingFileName As String, _
   ByVal lpNewFileName As String, _
   ByVal bFailIfExists As Long) As Long

Public Declare Function FindFirstFile1 Lib "kernel32" _
   Alias "FindFirstFile1A" _
  (ByVal lpFileName As String, _
   lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function FindNextFile1 Lib "kernel32" _
   Alias "FindNextFile1A" _
  (ByVal hFindFile As Long, _
   lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function FindClose Lib "kernel32" _
  (ByVal hFindFile As Long) As Long

Public Declare Function GetFileAttributes Lib "kernel32" _
   Alias "GetFileAttributesA" _
  (ByVal lpFileName As String) As Long

Public Declare Function LockWindowUpdate Lib "user32" _
  (ByVal hwndLock As Long) As Long


