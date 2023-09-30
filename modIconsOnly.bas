Attribute VB_Name = "Module1"
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRECT As RECT) As Long


Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRECT As RECT) As Long


Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long


Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long


Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
    Public Const RGN_AND = 1
    Public Const RGN_COPY = 5
    Public Const RGN_DIFF = 4
    Public Const RGN_OR = 2
    Public Const RGN_XOR = 3


Type POINTAPI
    x As Long
    Y As Long
    End Type


Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
    End Type
            

