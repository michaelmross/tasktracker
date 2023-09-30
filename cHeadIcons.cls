VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHeadIcons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private ListviewCtrl As ListView
Private LvwHwnd As Long

Public Enum SORTORDERENUM
  Ascn = 0
  Desc = 1
End Enum

Private Const LVM_FIRST = &H1000
Private Const HDM_FIRST = &H1200
Private Const HDM_SETITEM = (HDM_FIRST + 4)
Private Const HDM_SETIMAGELIST = (HDM_FIRST + 8)
Private Const HDI_IMAGE = &H20
Private Const HDI_FORMAT = &H4
Private Const HDF_BITMAP_ON_RIGHT = &H1000
Private Const HDF_IMAGE = &H800
Private Const HDF_STRING = &H4000

Private Type HDITEM
  mask As Long
  cxy As Long
  pszText As String
  hbm As Long
  cchTextMax As Long
  fmt As Long
  lParam As Long
  iImage As Long
  iOrder As Long
End Type

' HDITEM fmt
Private Const HDF_LEFT = 0
Private Const HDF_RIGHT = 1
'Private Const HDF_IMAGE = &H800
'Private Const HDF_BITMAP_ON_RIGHT = &H1000
'Private Const HDF_STRING = &H4000

' imagelist definitions
Private Const ILC_MASK = &H1
Private Const ILC_COLOR8 = &H8

Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Boolean
Private Declare Function ImageList_ReplaceIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hIcon As Long) As Long

Private Sub Class_Initialize()
  LvwHwnd = ImageList_Create(16, 16, ILC_MASK Or ILC_COLOR8, 2, 0)
  If LvwHwnd Then
    ' Load the icons into the image so that their zero-based
    ' indicescorrespond to the SORTORDERENUM values.
    Call ImageList_AddIcon(LvwHwnd, LoadResPicture("UP", 1))
    Call ImageList_AddIcon(LvwHwnd, LoadResPicture("DOWN", 1))
    Call ImageList_AddIcon(LvwHwnd, LoadResPicture("NOTHING", 1))
  End If
End Sub

Private Sub Class_Terminate()
  If LvwHwnd Then Call ImageList_Destroy(LvwHwnd)
End Sub

Public Property Get ListView() As ListView
  Set ListView = ListviewCtrl
End Property

Public Property Set ListView(lv As ListView)
  Set ListviewCtrl = lv
End Property

Private Function ListView_GetHeader(hwnd As Long) As Long
  ListView_GetHeader = SendMessage(hwnd, LVM_GETHEADER, 0, 0)
End Function

Private Function Header_SetItem(hwndHD As Long, i As Long, phdi As HDITEM) As Boolean
  Header_SetItem = SendMessage(hwndHD, HDM_SETITEM, i, phdi)
End Function

Private Function Header_SetImageList(hwnd As Long, himl As Long) As Long
  Header_SetImageList = SendMessage(hwnd, HDM_SETIMAGELIST, 0, ByVal himl)
End Function

Private Function ImageList_AddIcon(himl As Long, hIcon As Long) As Long
  ImageList_AddIcon = ImageList_ReplaceIcon(himl, -1, hIcon)
End Function

Public Function SetHeaderIcons(ByVal iActiveColumn As Integer, ByVal iSortOrder As SORTORDERENUM) As Boolean
  Static hwndHdr As Long
  Dim i As Long
  Dim fShow As Boolean
  Dim fAlignRight As Boolean
  Dim HDI As HDITEM
  
  If (LvwHwnd = 0) Or (ListviewCtrl Is Nothing) Then Exit Function
  If (ListviewCtrl.View <> lvwReport) Then Exit Function
  
  ' The ListView's header is created *after* the first
  ' ColumnHeader is added.
  If (hwndHdr = 0) Then
    hwndHdr = ListView_GetHeader(ListviewCtrl.hwnd)
    Call Header_SetImageList(hwndHdr, LvwHwnd)
  End If
  
  If (hwndHdr = 0) Then Exit Function
    
  With ListviewCtrl.ColumnHeaders
    For i = 0 To .Count - 1
      HDI.mask = HDI_FORMAT Or HDI_IMAGE
      
      fAlignRight = .item(i + 1).Alignment = lvwColumnRight
      
      ' Since we're setting the header's format, we have to
      ' specify the string flag
      HDI.fmt = HDF_STRING Or (fAlignRight And HDF_RIGHT)   ' HDF_LEFT = 0
      
      ' If the active column, add the sort icon with the appropriate
      ' alignment (the icon is removed if HDF_IMAGE is not set).
      If (i = iActiveColumn) Then
        HDI.fmt = HDI.fmt Or HDF_IMAGE Or ((fAlignRight = False) And HDF_BITMAP_ON_RIGHT)
         ' If not the Asc icon index (0), then set to
         ' the Desc (1) icon index.
         HDI.iImage = Abs(CBool(iSortOrder))
      Else
        HDI.iImage = 2    'for Vista
      End If
      
      
      Call Header_SetItem(hwndHdr, i, HDI)
    Next
  End With
  
  SetHeaderIcons = True
  
End Function


