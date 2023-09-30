Attribute VB_Name = "modWrite"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ©  2003 Michael M. Ross, Wordwise Solutions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private blLogOnce As Boolean
Private blLogTypes As Boolean
Private blVersionPrinted As Boolean

Public Sub subErrLog(ByVal strSource As String)

   If blTraceLog = True Then
      subTraceLog (strSource)
   End If
   
   If Err.Number = 0 Then Exit Sub

   If GetSetting("TaskTracker", "Settings", "Logging") = "1" Or _
      GetSetting("TaskTracker", "Settings", "Logging") = "2" Then
   Else
      Exit Sub
   End If
   
   Dim strErrLog As String
   Dim strErrNum As String, strErrDes As String
   Dim intF As Integer

   strErrNum = Err.Number          'This is necessary, otherwise
   strErrDes = Err.Description     'ErrorLogExit will be tripped
   
   If GetSetting("TaskTracker", "Settings", "Logging") = "2" Then
      MsgBox strErrNum + ": " + strErrDes + vbNewLine + strSource
   End If

On Error GoTo ErrorLogExit
    
    strErrLog = App.path + "\TaskTracker.log"
    
'   Log entry ***************************************************
    intF = FreeFile
    Open strErrLog For Append Access Write As #intF
            
        Print #intF, Str$(Now)
        
        If blVersionPrinted = False Then
            blVersionPrinted = True
            Print #intF, "Version: " + Trim$(Str$(VB.App.Major)) + "." + Trim$(Str$(VB.App.Minor)) + "." + Trim$(Str$(VB.App.Revision))
            If Str(GetSystemMetrics(SM_TABLETPC)) <> 0 Then
               Print #intF, "Tablet PC"
            End If
            If blFAT32 = True Then
               Print #intF, "File System: FAT32"
            Else
               Print #intF, "File System: NTFS"
            End If
        End If
        
        If LenB(strErrNum) > 0 Then
            Print #intF, "Error Number: " + strErrNum
        End If
        
        If LenB(strSource) > 0 Then
            Print #intF, "Source: " + strSource
        End If
        
        If LenB(strErrDes) > 0 Then
            Print #intF, "Description: " + strErrDes
        End If
        
        Print #intF, ""
               
    Close #intF
'   *************************************************************
    
ErrorLogExit:
        
End Sub

Public Sub subStatusLog(ByVal strSource As String)
   
   Dim strSnapshot As String
   Dim intF As Integer

On Error GoTo StatusLogExit
   
   strSnapshot = App.path + "\Snapshot.log"
   
   If blLogOnce = False Then
   
      blLogOnce = True
   
      If DoesFileExist(App.path + "\Snapshot.log") Then
         Kill App.path + "\Snapshot.log"
      End If
      
      intF = FreeFile
      Open strSnapshot For Append Access Write As #intF
              
      Print #intF, Str$(Now)
      Print #intF, "Version: " + Trim$(Str$(VB.App.Major)) + "." + Trim$(Str$(VB.App.Minor)) + "." + Trim$(Str$(VB.App.Revision))
      Print #intF, ""
           
   End If
   
   Close #intF
   
   If blLoading = True Then
      intF = FreeFile
      Open strSnapshot For Append Access Write As #intF
      
      Print #intF, strSource
      Print #intF, ""

   End If
   
   Close #intF
    
   If blLoading = False And blLogTypes = False Then
      blLogTypes = True
      intF = FreeFile
      Open strSnapshot For Append Access Write As #intF
      
      Dim lv1 As Integer
       With Form1.ListView1
         For lv1 = 1 To .ListItems.Count
            Print #intF, .ListItems(lv1).Text + ", " + .ListItems(lv1).SubItems(1) + ", " + .ListItems(lv1).SubItems(2)
            Print #intF, ""
         Next lv1
      End With

   End If
                 
   Close #intF
    
StatusLogExit:
        
End Sub

'Public Sub subCacheLog(ByVal strSource As String)
'
''   If GetSetting("TaskTracker", "Settings", "Caching") = "1" Then
''   Else
''      Exit Sub
''   End If
'
'   Dim strCachelog As String
'   Dim intF As Integer
'
'On Error GoTo CacheLogExit
'
'   strCachelog = App.path + "\Caching.log"
'
'   If blLogOnce = False Then
'
'      blLogOnce = True
'
'      intF = FreeFile
'      Open strCachelog For Append Access Write As #intF
'
'      Print #intF, Str$(Now)
'      Print #intF, ""
'      Print #intF, strSource
'      Print #intF, ""
'
'   End If
'
'   Close #intF
'
'CacheLogExit:
'
'End Sub

Public Sub subTraceLog(ByVal strSource As String)
On Error GoTo TraceLogExit

   Dim intF As Integer
   Dim strTraceLog As String
   
   strTraceLog = App.path + "\Trace.log"

   intF = FreeFile
   Open strTraceLog For Append Access Write As #intF

   Print #intF, strSource

   Close #intF

TraceLogExit:

End Sub
