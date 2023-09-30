Attribute VB_Name = "modCallback"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2003 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public objFind As LV_FINDINFO
Public objItem As LV_ITEM
  
Public Type POINTAPI
  x As Long
  y As Long
End Type

Public Type LV_FINDINFO
  flags As Long
  psz As String
  lParam As Long
  pt As POINTAPI
  vkDirection As Long
End Type

Public Type LV_ITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type
 
Public Const LVFI_PARAM As Long = &H1
Public Const LVIF_TEXT As Long = &H1

'Public Const LVM_FIRST As Long = &H1000
     
Public Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
   
Public lv1 As Boolean
   
Public Function CompareDates(ByVal lParam1 As Long, _
                             ByVal lParam2 As Long, _
                             ByVal hWnd As Long) As Long
On Error GoTo errlog

   Dim dDate1 As Date
   Dim dDate2 As Date
   Dim lvSortOrder As Integer
   
  'Obtain the item names and dates corresponding to the
  'input parameters
   dDate1 = ListView_GetItemDate(hWnd, lParam1)
   dDate2 = ListView_GetItemDate(hWnd, lParam2)
     
   If lv1 = True Then
      lvSortOrder = Form1.ListView1.SortOrder
   Else
      lvSortOrder = Form2.ListView2.SortOrder
   End If
     
   Select Case lvSortOrder
      Case 0: 'sort ascending
            
            If dDate1 < dDate2 Then
               CompareDates = 0
            ElseIf dDate1 = dDate2 Then
               CompareDates = 1
            Else
               CompareDates = 2
            End If
      
      Case 1: 'sort descending
   
            If dDate1 > dDate2 Then
               CompareDates = 0
            ElseIf dDate1 = dDate2 Then
               CompareDates = 1
            Else
               CompareDates = 2
            End If
   
   End Select

errlog:
   subErrLog ("modSortOrder: CompareDates")
End Function

Public Function CompareValues(ByVal lParam1 As Long, _
                              ByVal lParam2 As Long, _
                              ByVal hWnd As Long) As Long
     
   Dim val1 As Long
   Dim val2 As Long
     
  'Obtain the item names and values corresponding
  'to the input parameters
   val1 = ListView_GetItemValueStr(hWnd, lParam1)
   val2 = ListView_GetItemValueStr(hWnd, lParam2)
     
   'sort descending
            
   If val1 > val2 Then
      CompareValues = 0
   ElseIf val1 = val2 Then
      CompareValues = 1
   Else
      CompareValues = 2
   End If

End Function

Public Function ListView_GetItemDate(hWnd As Long, lParam As Long) As Date
On Error GoTo errlog
   
   Dim hIndex As Long
   Dim r As Long
  
  'Convert the input parameter to an index in the list view
   objFind.flags = LVFI_PARAM
   objFind.lParam = lParam
   hIndex = SendMessage(hWnd, LVM_FINDITEM, -1, objFind)
     
  'Obtain the value of the specified list view item.
   objItem.mask = LVIF_TEXT
   If lv1 = True Then
      objItem.iSubItem = Form1.ListView1.SortKey
   Else
      objItem.iSubItem = Form2.ListView2.SortKey
   End If
   objItem.pszText = Space$(32)
   objItem.cchTextMax = Len(objItem.pszText)
     
  'convert it into a date and exit
   r = SendMessage(hWnd, LVM_GETITEMTEXT, hIndex, objItem)
   If r > 0 Then
      ListView_GetItemDate = CDate(Left$(objItem.pszText, r))
   End If
  
errlog:
   subErrLog ("modSortOrder: ListView_GetItemDate")
End Function

Public Function ListView_GetItemValueStr(hWnd As Long, lParam As Long) As Long

   Dim hIndex As Long
   Dim r As Long
  
  'Convert the input parameter to an index in the list view
   objFind.flags = LVFI_PARAM
   objFind.lParam = lParam
   hIndex = SendMessage(hWnd, LVM_FINDITEM, -1, objFind)
     
  'Obtain the value of the specified list view item.
  'The objItem.iSubItem member is set to the index
  'of the column that is being retrieved.
   objItem.mask = LVIF_TEXT
   objItem.iSubItem = 1
   objItem.pszText = Space$(32)
   objItem.cchTextMax = Len(objItem.pszText)
     
   r = SendMessage(hWnd, LVM_GETITEMTEXT, hIndex, objItem)
   If r > 0 Then
      ListView_GetItemValueStr = CLng(Left$(objItem.pszText, r))
   End If

End Function



Public Function FARPROC(ByVal pfn As Long) As Long
  
  'A procedure that receives and returns
  'the value of the AddressOf operator.
 
  FARPROC = pfn

End Function


