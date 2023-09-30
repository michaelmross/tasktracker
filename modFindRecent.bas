Attribute VB_Name = "modFindRecent"
Option Explicit

Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
  (ByVal hwndOwner As Long, ByVal nFolder As Long, _
  pidl As ITEMIDLIST) As Long

Declare Function SHGetPathFromIDList Lib "shell32.dll" _
  Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
  ByVal pszPath As String) As Long

Public Type SH_ITEMID
    cb As Long
    abID As Byte
End Type

Public Type ITEMIDLIST
    mkid As SH_ITEMID
End Type

Private Const MAX_PATH As Integer = 260
Public Const CSIDL_RECENT = 8
Public Const CSIDL_APPDATA = 26

Public Function fGetSpecialFolder(CSIDL As Long) As String
On Error GoTo errlog
   Dim sPath As String
   Dim IDL As ITEMIDLIST
    '
    ' Retrieve info about system folders such as the
    ' "Desktop" folder. Info is stored in the IDL structure.
    '
    fGetSpecialFolder = ""
    If SHGetSpecialFolderLocation(Form1.hwnd, CSIDL, IDL) = 0 Then
        '
        ' Get the path from the ID list, and return the folder.
        '
        sPath = Space$(MAX_PATH)
        If SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath) Then
            fGetSpecialFolder = Left$(sPath, InStr(sPath, vbNullChar) - 1) & "\"
        End If
    End If
errlog:
If Err.Number <> 0 Then subErrLog ("modFindRecent: fGetSpecialFolder")
End Function


