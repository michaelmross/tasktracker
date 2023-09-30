Attribute VB_Name = "modDrives"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2003 VBnet, Randy Birch, All Rights Reserved.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Declare Function GetLogicalDriveStrings Lib "kernel32" _
     Alias "GetLogicalDriveStringsA" _
    (ByVal nBufferLength As Long, _
     ByVal lpBuffer As String) As Long

Public Declare Function GetDriveType Lib "kernel32" _
     Alias "GetDriveTypeA" _
    (ByVal nDrive As String) As Long

Private Declare Function WNetGetConnection Lib "mpr.dll" _
   Alias "WNetGetConnectionA" _
  (ByVal lpszLocalName As String, _
   ByVal lpszRemoteName As String, _
   cbRemoteName As Long) As Long
  
Private Declare Function GetVolumeInformation Lib "kernel32.dll" _
   Alias "GetVolumeInformationA" _
   (ByVal lpRootPathName As String, _
   ByVal lpVolumeNameBuffer As String, _
   ByVal nVolumeNameSize As Integer, _
   lpVolumeSerialNumber As Long, _
   lpMaximumComponentLength As Long, _
   lpFileSystemFlags As Long, _
   ByVal lpFileSystemNameBuffer As String, _
   ByVal nFileSystemNameSize As Long) As Long
   
'Private Type SHARE_INFO_2
'  shi2_netname       As Long
'  shi2_type          As Long
'  shi2_remark        As Long
'  shi2_permissions   As Long
'  shi2_max_uses      As Long
'  shi2_current_uses  As Long
'  shi2_path          As Long
'  shi2_passwd        As Long
'End Type

Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_FIXED = 3
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_CDROM = 5
Public Const DRIVE_RAMDISK = 6

Public SystemDrives() As String
Public NetworkDrives() As String
Public RemovableDrives() As String
Public CDDrives() As String

Public iSysDrv As Integer
Public iNetDrv As Integer
Public iRemDrv As Integer
Public iCDDrv As Integer

Public sDrives As String
Public nDrives As String

Public oldDrives As String

Public Sub GetFileSystem()
On Error GoTo errlog

Dim lRet As Long
Dim lans As Long
Dim sDrive As String, sVolumeName As String
Dim SDriveType As String

 SDriveType = String$(255, Chr$(0))
    
 lRet = GetVolumeInformation(sDrive, sVolumeName, 255, lans, 0, 0, SDriveType, 255)
   
 If Left$(Trim$(SDriveType), 4) <> "NTFS" Then
   blFAT32 = True
 End If

errlog:
   subErrLog ("modDrives: GetFileSystem")
End Sub

Public Sub subGetFixedDrives()
On Error GoTo errlog

   Dim allDrives As String
   Dim eachDrive As String
   Dim currDrive As String
   
   sDrives = vbNullString
   nDrives = vbNullString
 
  'get the list of all available drives
   allDrives = GetDriveString()
   eachDrive = allDrives
 
  'separate the drive strings and retrieve the drive type
   Do Until eachDrive = Chr$(0)
   
     'strip off one drive from the string allDrives
      currDrive = StripNulls(eachDrive)
   
     'get the drive type
      rgbDrvType (currDrive)
            
   Loop
   
   If blLoading = False Then
      If oldDrives <> allDrives Then
         blDriveChange = True
      End If
   End If
   oldDrives = allDrives
   
errlog:
   subErrLog ("modDrives: subGetFixedDrives")
End Sub


Private Function rgbDrvType(RootPathName As String) As String
On Error GoTo errlog
 
  'Passed is the drive to check (currDrive)
  'Returned is the type of drive for the About box  (rgbDrvType)
  'Result is the following 4 arrays, which are used in Form1: AddFileItemDetails
   ReDim Preserve SystemDrives(iSysDrv)     'system
   ReDim Preserve RemovableDrives(iRemDrv)  'removable, including floppy, ramdisk, unknown
   ReDim Preserve NetworkDrives(iNetDrv)    'network
   ReDim Preserve CDDrives(iCDDrv)          'CD
   Dim RN As String
   
   Select Case GetDriveType(RootPathName)
   
      Case DRIVE_REMOVABLE:
      
          Select Case Left$(RootPathName, 1)
              Case "a", "b": rgbDrvType = "Floppy drive"
              Case Else: rgbDrvType = "Removable drive"
          End Select
          
          RemovableDrives(iRemDrv) = RootPathName
          iRemDrv = iRemDrv + 1
          
          nDrives = nDrives + vbNewLine + RootPathName + "  " + rgbDrvType
 
      Case DRIVE_RAMDISK: rgbDrvType = "RAM drive"
      
          RemovableDrives(iRemDrv) = RootPathName
          iRemDrv = iRemDrv + 1
          
          nDrives = nDrives + vbNewLine + RootPathName + "  " + rgbDrvType
 
      Case DRIVE_CDROM:   rgbDrvType = "CD drive"
 
         CDDrives(iCDDrv) = RootPathName
         iCDDrv = iCDDrv + 1
 
          nDrives = nDrives + vbNewLine + RootPathName + "  " + rgbDrvType
                  
      Case DRIVE_FIXED:  rgbDrvType = "System drive"
      
   '     SystemDrives is an array of fixed drive letters
   '     (excludes CDs, Floppies, Removables, Network Drives)
         SystemDrives(iSysDrv) = RootPathName
         iSysDrv = iSysDrv + 1
                  
         sDrives = sDrives + vbNewLine + RootPathName + "  " + rgbDrvType
         
         RN = GetNetResourceName(Left$(RootPathName, 2))
         If Len(RN) > 0 Then sDrives = sDrives + " (" + RN + ")"
      
      Case DRIVE_REMOTE:  rgbDrvType = "Network drive"
      
         NetworkDrives(iNetDrv) = RootPathName
         iNetDrv = iNetDrv + 1
            
         nDrives = nDrives + vbNewLine + RootPathName + "  " + rgbDrvType
            
         RN = GetNetResourceName(Left$(RootPathName, 2))
         If Len(RN) > 0 Then nDrives = nDrives + " (" + RN + ")"
               
      Case Else:      rgbDrvType = "Unknown"
      
         RemovableDrives(iRemDrv) = RootPathName
         iRemDrv = iRemDrv + 1
          
         nDrives = nDrives + vbNewLine + RootPathName + "  " + rgbDrvType
 
   End Select
  
errlog:
   subErrLog ("modDrives: rgbDrvType")
End Function

Private Function GetDriveString() As String

  'returns string of available
  'drives each separated by a null
   Dim sBuffer As String
   
  'possible 26 drives, four chars each, plus one termininating char.
   sBuffer = Space$(26 * 4 + 1)
  
  If GetLogicalDriveStrings(Len(sBuffer), sBuffer) Then
  
     'do not trim off trailing null!
      GetDriveString = Trim$(sBuffer)
      
  End If

End Function

Public Function StripNulls(startstr As String) As String

 'Take a string separated by chr$(0)
 'and split off 1 item, shortening the
 'string so next item is ready for removal.
  Dim pos As Long

  pos = InStr(startstr$, Chr$(0))
  
  If pos Then
      
      StripNulls = Mid$(startstr, 1, pos - 1)
      startstr = Mid$(startstr, pos + 1, Len(startstr))
    
  End If

End Function

Public Function GetNetResourceName(sShare As String) As String

  'Returns the UNC name of share passed if
  'the user has logged on to a network.
  'The default return value is an empty string,
  'meaning either the share didn't exist or
  'there was no net connection.
   Dim buff As String
   Dim nSize As Long

   buff = Space$(MAX_PATH)
   nSize = Len(buff)
   
  'get name of resource associated with
  'the passed drive. Returns 0 on success,
  'or the error code
   If WNetGetConnection(sShare, buff, nSize) = 0 Then
    
      GetNetResourceName = TrimNull(buff)
    
   End If
   
End Function


