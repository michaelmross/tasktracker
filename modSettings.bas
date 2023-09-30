Attribute VB_Name = "modSettings"
Option Explicit

Public Sub GetForm1Settings()
On Error GoTo errlog

   Dim mHideTypes As String
   Dim ln As Integer
   Dim c As New cRegistry
    
   With c
   
      .ClassKey = HKEY_CURRENT_USER
      .SectionKey = "Software\TaskTracker\General"
      If .KeyExists = False Then
         .SectionKey = "Software\VB and VBA Program Settings\TaskTracker\Settings"
      End If
      .ValueType = REG_SZ
      
      .ValueKey = "IconsOnly"
      If LenB(.Value) > 0 Then
         Form1.mIcons.Checked = CBool(.Value)
      Else
         Form1.mIcons.Checked = False 'default
      End If
      
      .ValueKey = "TypeSortCol"
      If LenB(.Value) = 0 Then
         TypeSortCol = 0   'default (alpha sort)
      Else
         TypeSortCol = Val(.Value)
      End If
            
      .ValueKey = "RegSortCol"
      If LenB(.Value) = 0 Then
         RegSortCol = 1  'default
      Else
         RegSortCol = Val(.Value)
      End If
      
      .ValueKey = "RegPSortCol"
      If LenB(.Value) = 0 Then
         RegPSortCol = 2  'default
      Else
         RegPSortCol = Val(.Value)
      End If
      
      .ValueKey = "DelSortCol"
      If LenB(.Value) = 0 Then
         DelSortCol = 2  'default
      Else
         DelSortCol = Val(.Value)
      End If
      
      .ValueKey = "DelPSortCol"
      If LenB(.Value) = 0 Then
         DelPSortCol = 3  'default
      Else
         DelPSortCol = Val(.Value)
      End If
      
      .ValueKey = "MultSortCol"
      If LenB(.Value) = 0 Then
         MultSortCol = 2 'default
      Else
         MultSortCol = Val(.Value)
      End If
      
      .ValueKey = "MultPSortCol"
      If LenB(.Value) = 0 Then
         MultPSortCol = 3  'default
      Else
         MultPSortCol = Val(.Value)
      End If
      
      .ValueKey = "RegSortOrder"
      If LenB(.Value) = 0 Then
         RegSortOrder = 1  'default
      Else
         RegSortOrder = Val(.Value)
      End If
      
      .ValueKey = "RegPSortOrder"
      If LenB(.Value) = 0 Then
         RegPSortOrder = 1  'default
      Else
         RegPSortOrder = Val(.Value)
      End If
      
      .ValueKey = "DelSortOrder"
      If LenB(.Value) = 0 Then
         DelSortOrder = 1  'default
      Else
         DelSortOrder = Val(.Value)
      End If
      
      .ValueKey = "DelSortOrder"
      If LenB(.Value) = 0 Then
         DelPSortOrder = 1  'default
      Else
         DelPSortOrder = Val(.Value)
      End If
      
      .ValueKey = "MultSortOrder"
      If LenB(.Value) = 0 Then
         MultSortOrder = 1  'default
      Else
         MultSortOrder = Val(.Value)
      End If
      
      .ValueKey = "MultPSortOrder"
      If LenB(.Value) = 0 Then
         MultPSortOrder = 1  'default
      Else
         MultPSortOrder = Val(.Value)
      End If
      
      .ValueKey = "OnTop"
      If LenB(.Value) > 0 Then
         If CBool(.Value) = True Then
            Form1.mTop_Click
         Else
            Form1.mTop.Checked = False   'default
         End If
      End If
      
      .ValueKey = "Grid1"
      If LenB(.Value) > 0 Then
         Form1.mGridType.Checked = CBool(.Value)
      Else
         Form1.mGridType.Checked = False   'default
      End If
      
      Call SendMessage(Form1.ListView1.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, _
                       LVS_EX_GRIDLINES, ByVal Form1.mGridType.Checked)

      .ValueKey = "SingleClick"
      If LenB(.Value) > 0 Then
         Form1.mSingleClick.Checked = CBool(.Value)
         Form1.mDefaultFocus.Enabled = CBool(.Value)
      Else
         Form1.mSingleClick.Checked = True   'default
         Form1.mDefaultFocus.Enabled = True
      End If
   
      .ValueKey = "GlueRight"
      If LenB(.Value) > 0 Then
         Form1.mGlueRight.Checked = CBool(.Value)
      Else
         Form1.mGlueRight.Checked = False   'default
      End If
   
      .ValueKey = "GlueLeft"
      If LenB(.Value) > 0 Then
         Form1.mGlueLeft.Checked = CBool(.Value)
      Else
         Form1.mGlueLeft.Checked = False   'default
      End If
      
      .ValueKey = "HiddenTypes"
      If LenB(.Value) > 0 Then
         mHideTypes = .Value
         'parse the string for each filetype
         For ln = 1 To Len(mHideTypes)
           If Mid$(mHideTypes, ln, 1) = ";" Then
              ReDim Preserve fNoShow(Form1.hd)
              fNoShow(Form1.hd) = Left$(mHideTypes, InStr(mHideTypes, ";") - 1)
              mHideTypes = Right$(mHideTypes, Len(mHideTypes) - ln)
              Form1.hd = Form1.hd + 1
              ln = 1
           End If
         Next ln
         If Form1.hd > 0 Then
            Form1.blHideType = True
            Form1.mShowAll.Checked = False
         End If
      Else
         ReDim Preserve fNoShow(0)
         Form1.blHideType = False
         Form1.mShowAll.Checked = True
      End If
      
      .ValueKey = "Network"
      If LenB(.Value) > 0 Then
         Form1.mNetwork.Checked = CBool(.Value)
      Else
         Form1.mNetwork.Checked = False 'default
      End If
   
      .ValueKey = "Removable"
      If LenB(.Value) > 0 Then
         Form1.mRemovable.Checked = CBool(.Value)
      Else
         Form1.mRemovable.Checked = False 'default
      End If
      
      .ValueKey = "Taskbar"
      If LenB(.Value) > 0 Then
         Form1.mTaskbar.Checked = CBool(.Value)
      Else
         Form1.mTaskbar.Checked = False 'default
      End If
         
      .ValueKey = "Fade"
      If LenB(.Value) > 0 Then
         Form1.mFade.Checked = CBool(.Value)
         blNoFade = False
         Form1.Timer2.Enabled = True
      Else
         Form1.mFade.Checked = False 'default
      End If
      
      .ValueKey = "HideTips"
      If LenB(.Value) > 0 Then
         Form1.mHideTips.Checked = CBool(.Value)
      Else
         Form1.mHideTips.Checked = False 'default
      End If
      
      .ValueKey = "HideTips"
      If LenB(.Value) = 0 Then 'Or _
'         Int(Val(GetSetting("TaskTracker", "Settings", "RefreshTime"))) < 2 Or _
'         Int(Val(GetSetting("TaskTracker", "Settings", "RefreshTime"))) > 10 Then
         Form1.iRefreshTime = 5   'default
      Else
         Form1.iRefreshTime = Int(Val(.Value))
      End If
   
      .ValueKey = "Autorun"
      If LenB(.Value) > 0 Then
         If .Value = "Minimized" Then
            Form1.mMin.Checked = True
            SetAutoRun
         ElseIf .Value = "Normal" Then
            Form1.mNorm.Checked = True
            SetAutoRun
         Else
            Form1.mMin.Checked = False
            Form1.mNorm.Checked = False
            ClearAutoRun
         End If
      Else
         Form1.mMin.Checked = True     'default
         SetAutoRun
      End If
      
      .ValueKey = "Snapshot"
      If LenB(.Value) > 0 Then
         blSnapshot = True
         .ValueKey = "Snapshot"
         .DeleteValue   'once-only setting
      Else
         blSnapshot = False 'default
      End If
        
      .ValueKey = "TypesFocus"
      If LenB(.Value) > 0 Then
         Form1.mFocusTypes.Checked = CBool(.Value)
      Else
         Form1.mFocusTypes.Checked = True 'default
      End If
      
      If Form1.mFocusTypes.Checked = True Then
         Form1.mFocusFiles.Checked = False   'default
      Else
         Form1.mFocusFiles.Checked = True
      End If
      
      .ValueKey = "NoCache"
      If LenB(.Value) > 0 Then
         If CBool(.Value) = True Then
            Form1.mNoCache_Click
         End If
      Else
         Form1.mNoCache.Checked = False
      End If
      If blSnapshot = True Then               'override caching if snapshot'ing
         blNoCache = True
         Form1.mSynchronization.Enabled = False
      End If
   
      .ValueKey = "Caching"
      If LenB(.Value) > 0 Then
         If .Value = "Never" Then
            Form1.mNeverSync.Checked = True
            Form1.mSyncShow.Checked = False
            Form1.mSyncStartup.Checked = False
         ElseIf .Value = "FirstShow" Then
            Form1.mNeverSync.Checked = False
            Form1.mSyncShow.Checked = True
            Form1.mSyncStartup.Checked = False
         ElseIf .Value = "Startup" Then
            Form1.mNeverSync.Checked = False
            Form1.mSyncShow.Checked = False
            Form1.mSyncStartup.Checked = True
         End If
      Else  'default
         Form1.mNeverSync.Checked = False
         Form1.mSyncShow.Checked = True
         Form1.mSyncStartup.Checked = False
      End If
      
      .ValueKey = "AutoPreview"
      If LenB(.Value) > 0 Then
         Form1.mAutoPreview.Checked = CBool(.Value)
      Else
         Form1.mAutoPreview.Checked = False 'default
      End If
'
      .SectionKey = "Software\TaskTracker\Layout"
      If .KeyExists = False Then
         .SectionKey = "Software\VB and VBA Program Settings\TaskTracker"
      End If
      
      .ValueKey = "Form1Width"
      If LenB(.Value) > 0 Then
         Form1.Width = Val(.Value)
      Else
         Form1.Width = 2800
      End If
      
      .ValueKey = "Form1Height"
      If LenB(.Value) > 0 Then
         Form1.Height = Val(.Value)
      Else
         Form1.Height = 7200
      End If
      
      .ValueKey = "Form1Top"
      If LenB(.Value) > 0 Then
         Form1.Top = Val(.Value)
         Form1.mSaveSize.Checked = True
      End If
      
      .ValueKey = "Form1Left"
      If Len(.Value) > 0 Then
         Form1.Left = Val(.Value)
      End If
   
   End With

errlog:
   If Err.Number <> 0 Then subErrLog ("modSettings:GetForm1Settings")
End Sub

Public Sub GetForm2Settings()
On Error GoTo errlog
      
   Dim c As New cRegistry
    
   With c
   
      .ClassKey = HKEY_CURRENT_USER
      .SectionKey = "Software\TaskTracker\General"
      If .KeyExists = False Then
         .SectionKey = "Software\VB and VBA Program Settings\TaskTracker"
      End If
      .ValueType = REG_SZ
   
      .ValueKey = "FileExt"
      If LenB(.Value) > 0 Then
         Form2.mExt.Checked = CBool(.Value)
      Else
         Form2.mExt.Checked = True 'default
      End If
      
      .ValueKey = "LastDate"
      If LenB(.Value) > 0 Then
         If .Value = "Created" Then
            Form2.mAccessed.Checked = False
            Form2.mModified.Checked = False
            Form2.mCreated.Checked = True
         ElseIf .Value = "Modified" Then
            Form2.mAccessed.Checked = False
            Form2.mModified.Checked = True
            Form2.mCreated.Checked = False
         ElseIf .Value = "Accessed" Then
            Form2.mAccessed.Checked = True
            Form2.mModified.Checked = False
            Form2.mCreated.Checked = False
         End If
      Else  'default
         Form2.mAccessed.Checked = False
         Form2.mModified.Checked = True
         Form2.mCreated.Checked = False
      End If
      
      'setting in form1
      .ValueKey = "Grid2"
      If LenB(.Value) > 0 Then
         Form1.mGridName.Checked = CBool(.Value)
      Else
         Form1.mGridName.Checked = True   'default
      End If
      Call SendMessage(Form2.ListView2.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, _
                       LVS_EX_GRIDLINES, ByVal Form1.mGridName.Checked)
   
      .ValueKey = "Explorer"
      If LenB(.Value) > 0 Then
         Form2.blExplorer = CBool(.Value)
      Else
         Form2.blExplorer = False 'default
      End If
           
      .ValueKey = "Simple"
      If LenB(.Value) > 0 Then
         Form2.mSimple.Checked = CBool(.Value)
         Form2.subSimple
      Else
         Form2.mSimple.Checked = False 'default
      End If
      
      .ValueKey = "FullPath"
      If LenB(.Value) > 0 Then
         Form2.mFullPath.Checked = CBool(.Value)
      Else
         Form2.mFullPath.Checked = False 'default
      End If
         
      .ValueKey = "ContainerFolder"
      If LenB(.Value) > 0 Then
         Form2.mContainerFolder.Checked = CBool(.Value)
      Else
         Form2.mContainerFolder.Checked = False 'default
      End If
      
      Form2.LoadColPos
      
   End With

errlog:
   If Err.Number <> 0 Then subErrLog ("modSettings:GetForm2Settings")
End Sub

Public Sub GetListView2Settings()

   Dim c As New cRegistry
    
   With c
   
      .ClassKey = HKEY_CURRENT_USER
      .SectionKey = "Software\TaskTracker\Layout"
      If .KeyExists = False Then
         .SectionKey = "Software\VB and VBA Program Settings\TaskTracker"
      End If
      .ValueType = REG_SZ

      .ValueKey = "col1width"
      If Len(.Value) > 0 Then
         col1width = Val(.Value)
      Else
         col1width = 3850
      End If
      
      .ValueKey = "col2width"
      If Len(.Value) > 0 Then
         col2width = Val(.Value)
      Else
         col2width = 2500
      End If
      
      .ValueKey = "colSwidth"
      If Len(.Value) > 0 Then
         colSwidth = Val(.Value)
      Else
         colSwidth = 0
      End If
      
      .ValueKey = "col1Pwidth"
      If Len(.Value) > 0 Then
         col1Pwidth = Val(.Value)
      Else
         col1Pwidth = 2000
      End If
      
      .ValueKey = "col2Pwidth"
      If Len(.Value) > 0 Then
         col2Pwidth = Val(.Value)
      Else
         col2Pwidth = 2300
      End If
      
      .ValueKey = "col3Pwidth"
      If Len(.Value) > 0 Then
         col3Pwidth = Val(.Value)
      Else
         col3Pwidth = 2050
      End If
      
      .ValueKey = "col4Swidth"
      If Len(.Value) > 0 Then
         col4Swidth = Val(.Value)
      Else
         col4Swidth = 0
      End If
      
      .ValueKey = "col1Awidth"
      If Len(.Value) > 0 Then
         col1Awidth = Val(.Value)
      Else
         col1Awidth = 1900
      End If
      
      .ValueKey = "col2Awidth"
      If Len(.Value) > 0 Then
         col2Awidth = Val(.Value)
      Else
         col2Awidth = 2300
      End If
      
      .ValueKey = "col3Awidth"
      If Len(.Value) > 0 Then
         col3Awidth = Val(.Value)
      Else
         col3Awidth = 1900
      End If
      
      .ValueKey = "col1PAwidth"
      If Len(.Value) > 0 Then
         col1PAwidth = Val(.Value)
      Else
         col1PAwidth = 1890
      End If
      
      .ValueKey = "col2PAwidth"
      If Len(.Value) > 0 Then
         col2PAwidth = Val(.Value)
      Else
         col2PAwidth = 1300
      End If
      
      .ValueKey = "col3PAwidth"
      If Len(.Value) > 0 Then
         col3PAwidth = Val(.Value)
      Else
         col3PAwidth = 1300
      End If
      
      .ValueKey = "col4PAwidth"
      If Len(.Value) > 0 Then
         col4PAwidth = Val(.Value)
      Else
         col4PAwidth = 1300
      End If
      
   End With
   
errlog:
   If Err.Number <> 0 Then subErrLog ("modSettings:GetListView2Settings")
End Sub
