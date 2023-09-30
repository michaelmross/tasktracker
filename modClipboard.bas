Attribute VB_Name = "modClipboard"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright © James Crowley
'http://www.developerfusion.com/show/224/.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Required data structures
Private Type POINTAPI
X As Long
Y As Long
End Type

' Clipboard Manager Functions
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Public Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long

' Other required Win32 APIs
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Const CF_HDROP = 15

' Global Memory Flags
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)

Private Type DROPFILES
   pFiles As Long
   pt As POINTAPI
   fNC As Long
   fWide As Long
End Type

Public Function ClipboardCopyFiles(Files() As String) As Boolean
On Error GoTo errlog

   Dim data As String
   Dim df As DROPFILES
   Dim hGlobal As Long
   Dim lpGlobal As Long
   Dim i As Long
   
   ' Open and clear existing crud off clipboard.
   If OpenClipboard(0&) Then
   Call EmptyClipboard
   
   ' Build double-null terminated list of files.
   For i = LBound(Files) To UBound(Files)
      data = data & Files(i) & vbNullChar
   Next
   
   data = data & vbNullChar
   Dim ld As Integer
   For ld = 1 To Len(data)
      If Left$(data, 1) = vbNullChar Then
         data = Right$(data, Len(data) - 1)
      End If
   Next ld
         
   ' Allocate and get pointer to global memory,
   ' then copy file list to it.
   hGlobal = GlobalAlloc(GHND, Len(df) + Len(data))
   If hGlobal Then
      lpGlobal = GlobalLock(hGlobal)
      
      ' Build DROPFILES structure in global memory.
      df.pFiles = Len(df)
      Call CopyMem(ByVal lpGlobal, df, Len(df))
      Call CopyMem(ByVal (lpGlobal + Len(df)), ByVal data, Len(data))
      Call GlobalUnlock(hGlobal)
      
      ' Copy data to clipboard, and return success.
      If SetClipboardData(CF_HDROP, hGlobal) Then
         ClipboardCopyFiles = True
      End If
   End If
   
   ' Clean up
   Call CloseClipboard
   End If
   
errlog:
   subErrLog ("modClipboard: ClipboardCopyFiles")
End Function
   
