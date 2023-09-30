Attribute VB_Name = "modSave"
Option Explicit
Public gINIFile As String   'this global contains the name of the INI file assigned to this application

'API Function calls
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


'This public method is used to retrieve a Key value from the INI file)
Public Function GetINIString(strINIFile As String, strSection As String, strKey As String, _
                                strDefault As String) As String
   
    Dim strTemp As String * 256     'set string max length to 256 chars
    Dim intLength As Integer
    strTemp = ""
    strTemp = Space$(256)           'initialize string with spaces
    intLength = GetPrivateProfileString(strSection, strKey, strDefault, strTemp, 255, strINIFile)
    GetINIString = Left$(strTemp, intLength)
                               
   
End Function

'This public method is used to write a Key value to the INI file
Public Sub WriteINIString(strINIFile As String, strSection As String, strKey As String, strvalue As String)
    Dim indx As Integer
    Dim strTemp As String
   
    strTemp = strvalue
   
    'a key value must not contain either a carriage return or a line feed, therefore check for these in
    'the passed string, and substitute a " " if you find one.  This is purely precautionary.
    For indx = 1 To Len(strvalue)
        If Mid$(strvalue, indx, 1) = vbCr Or Mid$(strvalue, indx, 1) = vbLf Then
            Mid$(strvalue, indx) = " "
        End If
    Next indx
   
    indx = WritePrivateProfileString(strSection, strKey, strTemp, strINIFile)
End Sub

