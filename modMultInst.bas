Attribute VB_Name = "modMultInst"
Option Explicit

'''''''Declare in .bas Module''''''''''''''''''''''''''''''''
Public MgWind2() As frmMsgWind     '''Form Object Array
Public MgWinState() As Integer     '''Integer array parallels and holds val 1=loaded, 0=Not Loaded
Public mgFrm2Ix%                   '''holds the total Top Dimension Number Of Both Arrays


Public Function MakeAMsgWind() As Integer
'''create a Function to call to create indexed instances of the MgWind2()

'''To Use in a project copy the MsgWind2.frm and MsgWind2.frx files
'''into the project's dir and add to the Project.

'''''''Declare in .bas Module''''''''''''''''''''''''''''''''
'''Public MgWind2() As frmMsgWind     '''Form Object Array
'''Public MgWinState() As Integer     '''Integer array parallels and holds val 1=loaded, 0=Not Loaded
'''Public mgFrm2Ix%                   '''holds the total Top Dimension Number Of Both Arrays
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''in Form_Load Or Sub_Main of Project '''Initialize Arrays With First Dim, 0
'''ReDim MgWind2(0)
'''ReDim MgWinState(0)
'''MgWinState(0) = 0                   ''' =0 To Indicate Form Not yet Loaded
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''Use At Calling Sub, To Load The Form's Text1.Text And Make The Form Visible'''''
'''RetClk% = MakeAMsgWind
'''MgWind2(RetClk%).Text1.Text = "This Instance Of MgWind2(" & Trim$(Str$(RetClk%)) & ") Is Loaded."
'''MgWind2(RetClk%).Visible = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''Copy This Code into The .bas Public Function MakeAMsgWind '''''''''''''

For ClnIx% = mgFrm2Ix% To 0 Step -1      '''Check To See If There Are Empty Array Elements At Top End
 If MgWinState(ClnIx%) = 1 Then
 GoTo GotaLst
 End If
Next ClnIx%
ClnIx% = 0

GotaLst:
If ClnIx% < mgFrm2Ix% Then               '''If Yes Then Remove Them With Redim = Top Loaded Element
ReDim Preserve MgWind2(ClnIx%)
ReDim Preserve MgWinState(ClnIx%)
mgFrm2Ix% = ClnIx%
End If

For CkMgfrmIx% = 0 To mgFrm2Ix%          '''Then See If There Is An Intermediate Element Not Loaded
 If MgWinState(CkMgfrmIx%) = 0 Then      '''If Yes (=0) Then GoTo GotaFrm And Use It
 GoTo GotaFrm
 End If
Next CkMgfrmIx%
mgFrm2Ix% = CkMgfrmIx%                   '''Else Use The Inc. Value To Create A New Top Element To Use

ReDim Preserve MgWind2(CkMgfrmIx%)
ReDim Preserve MgWinState(CkMgfrmIx%)

GotaFrm:

Set MgWind2(CkMgfrmIx%) = New frmMsgWind                        '''Then Load Up The New Form Element
MakeAMsgWind = CkMgfrmIx%                                      '''Set Function's Return Value = The New Form's Index
MgWind2(CkMgfrmIx%).Label1.Caption = Trim$(Str$(CkMgfrmIx%))  '''Store The Index In The Form In An Invisible Label
MgWinState(CkMgfrmIx%) = 1                                   '''Load The Array To Keep Each Elements Loaded State

End Function

