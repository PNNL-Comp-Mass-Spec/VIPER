Attribute VB_Name = "Module23"
'module contains overlay drawing functions
'created: 12/20/2002 nt
'last modified: 02/12/2003 nt
'-------------------------------------------------------------------
Option Explicit

Public Const HI_MM = 100       'each logical unit is mapped in 0.01 millimters

'LOGICAL COORDINATE SYSTEM CONSTANTS-SIZE
Public Const LoX0 = 0          'logical coordinates defaults
Public Const LoY0 = 0          '(X0,Y0)-(XE,YE) define real
Public Const LoX1 = 3000       'logical window; (X1,Y1) defines
Public Const LoY1 = 3000       'small offset from the coordinate
Public Const LoX2 = 97000      'logical window; (X1,Y1)-(X2,Y2)
Public Const LoY2 = 97000      'defines small offset from the
Public Const LoXE = 100000     'coordinate axes
Public Const LoYE = 100000
'LOGICAL COORDINATE SYSTEM CONSTANTS-INDENTS
Public Const loSXPercent = 0.03
Public Const loSYPercent = 0.03
Public Const loLXPercent = 0.07
Public Const loLYPercent = 0.07
'LOGICAL COORDINATE SYSTEM CONSTANTS-OTHER
Public Const LoWndW As Long = (LoXE - LoX0) / (1 - loSXPercent - loLXPercent)
Public Const LoWndH As Long = (LoYE - LoY0) / (1 - loSYPercent - loLYPercent)
Public Const LoSX As Long = loSXPercent * LoWndW
Public Const loSY As Long = loSYPercent * LoWndH
'''Public Const loLX As Long = loLXPercent * LoWndW
'''Public Const loLY As Long = loLYPercent * LoWndH


'overlay drawing structures
Public hOlyBackClrBrush As Long
Public hOlyForeClrPen As Long
Public hOlyGridClrPen As Long

Private hOlyClrPen() As Long             'we want ready drawing tools
Private hOlyClrPenStick() As Long        'we want ready drawing tools
Private hOlyClrBrush() As Long

'-------------------------------------------------------------------
'drawing objects maintenance functions

Public Sub CreateOlyBackClrObject(ByVal BackColor As Long)
On Error Resume Next
If hOlyBackClrBrush <> 0 Then DeleteObject (hOlyBackClrBrush)
hOlyBackClrBrush = CreateSolidBrush(BackColor)
End Sub

Public Sub CreateOlyForeClrObject(ByVal ForeColor As Long)
On Error Resume Next
If hOlyForeClrPen <> 0 Then DeleteObject (hOlyForeClrPen)
hOlyForeClrPen = CreatePen(PS_SOLID, 1, ForeColor)

If hOlyGridClrPen <> 0 Then DeleteObject (hOlyGridClrPen)
hOlyGridClrPen = CreatePen(OlyOptions.GRID.LineStyle, 1, ForeColor)
End Sub

Public Function AddEditOlyClr(ByVal OlyInd As Long, ByVal OlyColor As Long) As Boolean
'---------------------------------------------------------------------------------------
'if call to OlyClr object arrays fail we need to initialize; otherwise just change color
'---------------------------------------------------------------------------------------
On Error GoTo err_AddEditOlyClr
If hOlyClrBrush(OlyInd) <> 0 Then DeleteObject (hOlyClrBrush(OlyInd))
hOlyClrBrush(OlyInd) = CreateSolidBrush(OlyColor)
If hOlyClrPen(OlyInd) <> 0 Then DeleteObject (hOlyClrPen(OlyInd))
hOlyClrPen(OlyInd) = CreatePen(PS_SOLID, 1, OlyColor)
If hOlyClrPenStick(OlyInd) <> 0 Then DeleteObject (hOlyClrPenStick(OlyInd))
hOlyClrPenStick(OlyInd) = CreatePen(PS_SOLID, OlyOptions.DefStickWidth, OlyColor)
AddEditOlyClr = True
Exit Function

err_AddEditOlyClr:
If Err.Number = 9 Then
   ReDim Preserve hOlyClrPen(OlyInd)
   ReDim Preserve hOlyClrBrush(OlyInd)
   ReDim Preserve hOlyClrPenStick(OlyInd)
   Resume
End If
End Function

Public Function RemoveOlyClr(ByVal OlyInd As Long) As Boolean
Dim OlyClrCnt As Long
Dim I As Long
On Error Resume Next
OlyClrCnt = UBound(hOlyClrBrush) - 1
If hOlyClrBrush(OlyInd) <> 0 Then DeleteObject (hOlyClrBrush(OlyInd))
If hOlyClrPen(OlyInd) <> 0 Then DeleteObject (hOlyClrPen(OlyInd))
If hOlyClrPenStick(OlyInd) <> 0 Then DeleteObject (hOlyClrPenStick(OlyInd))
For I = OlyInd To OlyClrCnt - 1
    hOlyClrBrush(I) = hOlyClrBrush(I + 1)
    hOlyClrPen(I) = hOlyClrPen(I + 1)
    hOlyClrPenStick(I) = hOlyClrPenStick(I + 1)
Next I
ReDim Preserve hOlyClrPen(OlyClrCnt - 1)
ReDim Preserve hOlyClrPenStick(OlyClrCnt - 1)
ReDim Preserve hOlyClrBrush(OlyClrCnt - 1)
End Function
'-------------------------------------------------------------------



'drawing functions
Public Sub OlyDrawBox(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
On Error Resume Next
I = SelectObject(hDC, hOlyClrBrush(OlyInd))
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
      If .Visible(I) Then
        If Not .OutOfScope(I) Then
           Rectangle hDC, .X(I) - .XL(I), .Y(I) - .YL(I), _
                       .X(I) + .XU(I), .Y(I) + .YU(I)
        End If
      End If
   Next I
End With
End Sub


Public Sub OlyDrawEmptyBox(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
On Error Resume Next
I = SelectObject(hDC, GetStockObject(NULL_BRUSH))
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
      If .Visible(I) Then
        If Not .OutOfScope(I) Then
           Rectangle hDC, .X(I) - .XL(I), .Y(I) - .YL(I), _
                       .X(I) + .XU(I), .Y(I) + .YU(I)
        End If
      End If
   Next I
End With
End Sub


Public Sub OlyDrawSpot(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
Dim h As Long, v As Long
On Error Resume Next
I = SelectObject(hDC, hOlyClrBrush(OlyInd))
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
      If .Visible(I) Then
         If Not .OutOfScope(I) Then
            h = .R(I) \ 2
            v = .R(I) \ 2
            Ellipse hDC, .X(I) - h, .Y(I) - v, .X(I) + h, .Y(I) + v
         End If
      End If
   Next I
End With
End Sub


Public Sub OlyDrawEmptySpot(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
Dim h As Long, v As Long
On Error Resume Next
I = SelectObject(hDC, GetStockObject(NULL_BRUSH))
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
      If .Visible(I) Then
         If Not .OutOfScope(I) Then
            h = .R(I) \ 2
            v = .R(I) \ 2
            Ellipse hDC, .X(I) - h, .Y(I) - v, .X(I) + h, .Y(I) + v
         End If
      End If
   Next I
End With
End Sub


Public Sub OlyDrawTriangle(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
Dim h As Long, v As Long
Dim TriPoints(2) As POINTAPI
On Error Resume Next
I = SelectObject(hDC, hOlyClrBrush(OlyInd))
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
      If .Visible(I) Then
         If Not .OutOfScope(I) Then
            h = .R(I) \ 2
            v = .R(I) \ 2
            TriPoints(0).X = .X(I) - h:    TriPoints(0).Y = .Y(I) - v / 2
            TriPoints(1).X = .X(I):      TriPoints(1).Y = .Y(I) + v / 2
            TriPoints(2).X = .X(I) + h:      TriPoints(2).Y = .Y(I) - v / 2
            Polygon hDC, TriPoints(0), 3
         End If
      End If
   Next I
End With
End Sub


Public Sub OlyDrawEmptyTriangle(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
Dim h As Long, v As Long
Dim TriPoints(2) As POINTAPI
On Error Resume Next
I = SelectObject(hDC, GetStockObject(NULL_BRUSH))
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
      If .Visible(I) Then
         If Not .OutOfScope(I) Then
            h = .R(I) \ 2
            v = .R(I) \ 2
            TriPoints(0).X = .X(I) - h:    TriPoints(0).Y = .Y(I) - v / 2
            TriPoints(1).X = .X(I):      TriPoints(1).Y = .Y(I) + v / 2
            TriPoints(2).X = .X(I) + h:      TriPoints(2).Y = .Y(I) - v / 2
            Polygon hDC, TriPoints(0), 3
         End If
      End If
   Next I
End With
End Sub


Public Sub OlyDrawStick(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
Dim Res As Long
Dim OldPen As Long
Dim ptPoint As POINTAPI
On Error Resume Next
Res = SelectObject(hDC, GetStockObject(NULL_BRUSH))
OldPen = SelectObject(hDC, hOlyClrPenStick(OlyInd))
Res = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
      If .Visible(I) Then
         If Not .OutOfScope(I) Then
            Res = MoveToEx(hDC, .X(I) - .XL(I), .Y(I), ptPoint)
            Res = LineTo(hDC, .X(I) + .XU(I), .Y(I))
         End If
      End If
   Next I
End With
Res = SelectObject(hDC, OldPen)
End Sub


Public Sub OlyDrawTriStar(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long, Res As Long
Dim h As Long, v As Long
Dim XC As Long, YC As Long
Dim tp(2) As POINTAPI
Dim ptPoint As POINTAPI
On Error Resume Next
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
       If Not .OutOfScope(I) Then
          h = .R(I) \ 2
          v = .R(I) \ 2
          tp(0).X = .X(I) - h:    tp(0).Y = .Y(I) - v / 2
          tp(1).X = .X(I):    tp(1).Y = .Y(I) + v / 2
          tp(2).X = .X(I) + h:     tp(2).Y = .Y(I) - v / 2
          XC = CLng((tp(0).X + tp(1).X + tp(2).X) / 3)
          YC = CLng((tp(0).Y + tp(1).Y + tp(2).Y) / 3)
          Res = MoveToEx(hDC, tp(0).X, tp(0).Y, ptPoint)
          Res = LineTo(hDC, XC, YC)
          Res = LineTo(hDC, tp(1).X, tp(1).Y)
          Res = MoveToEx(hDC, tp(2).X, tp(2).Y, ptPoint)
          Res = LineTo(hDC, XC, YC)
       End If
   Next I
End With
End Sub


Public Sub OlyDrawBoxNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
On Error Resume Next
I = SelectObject(hDC, hOlyClrBrush(OlyInd))
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
      If .Visible(I) Then
         If Not .OutOfScope(I) Then
            Rectangle hDC, .Y(I) - .YL(I), .X(I) - .XL(I), _
                        .Y(I) + .YU(I), .X(I) + .XU(I)
         End If
      End If
   Next I
End With
End Sub


Public Sub OlyDrawEmptyBoxNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
On Error Resume Next
I = SelectObject(hDC, GetStockObject(NULL_BRUSH))
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
      If .Visible(I) Then
         If Not .OutOfScope(I) Then
            Rectangle hDC, .Y(I) - .YL(I), .X(I) - .XL(I), _
                        .Y(I) + .YU(I), .X(I) + .XU(I)
         End If
      End If
   Next I
End With
End Sub


Public Sub OlyDrawSpotNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
Dim h As Long, v As Long
On Error Resume Next
I = SelectObject(hDC, hOlyClrBrush(OlyInd))
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
      If .Visible(I) Then
         If Not .OutOfScope(I) Then
            h = .R(I) \ 2
            v = .R(I) \ 2
            Ellipse hDC, .Y(I) - v, .X(I) - h, .Y(I) + v, .X(I) + h
         End If
      End If
   Next I
End With
End Sub


Public Sub OlyDrawEmptySpotNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
Dim h As Long, v As Long
On Error Resume Next
I = SelectObject(hDC, GetStockObject(NULL_BRUSH))
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
      If .Visible(I) Then
         If Not .OutOfScope(I) Then
            h = .R(I) \ 2
            v = .R(I) \ 2
            Ellipse hDC, .Y(I) - v, .X(I) - h, .Y(I) + v, .X(I) + h
         End If
      End If
   Next I
End With
End Sub


Public Sub OlyDrawStickNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
Dim Res As Long
Dim OldPen As Long
Dim ptPoint As POINTAPI
On Error Resume Next
Res = SelectObject(hDC, GetStockObject(NULL_BRUSH))
OldPen = SelectObject(hDC, hOlyClrPenStick(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
      If .Visible(I) Then
         If Not .OutOfScope(I) Then
            Res = MoveToEx(hDC, .Y(I), .X(I) - .XL(I), ptPoint)
            Res = LineTo(hDC, .Y(I), .X(I) + .XU(I))
         End If
      End If
   Next I
End With
Res = SelectObject(hDC, OldPen)
End Sub


Public Sub OlyDrawTriangleNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
Dim h As Long, v As Long
Dim TriPoints(2) As POINTAPI
On Error Resume Next
I = SelectObject(hDC, hOlyClrBrush(OlyInd))
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
      If .Visible(I) Then
         If Not .OutOfScope(I) Then
            h = .R(I) \ 2
            v = .R(I) \ 2
            TriPoints(0).X = .Y(I) - v / 2:    TriPoints(0).Y = .X(I) - h
            TriPoints(1).X = .Y(I) + v / 2:      TriPoints(1).Y = .X(I)
            TriPoints(2).X = .Y(I) - v / 2:      TriPoints(2).Y = .X(I) + h
            Polygon hDC, TriPoints(0), 3
         End If
      End If
   Next I
End With
End Sub


Public Sub OlyDrawEmptyTriangleNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
Dim h As Long, v As Long
Dim TriPoints(2) As POINTAPI
On Error Resume Next
I = SelectObject(hDC, GetStockObject(NULL_BRUSH))
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
      If .Visible(I) Then
         If Not .OutOfScope(I) Then
            h = .R(I) \ 2
            v = .R(I) \ 2
            TriPoints(0).X = .Y(I) - v / 2:    TriPoints(0).Y = .X(I) - h
            TriPoints(1).X = .Y(I) + v / 2:      TriPoints(1).Y = .X(I)
            TriPoints(2).X = .Y(I) - v / 2:      TriPoints(2).Y = .X(I) + h
            Polygon hDC, TriPoints(0), 3
         End If
      End If
   Next I
End With
End Sub


Public Sub OlyDrawTriStarNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long, Res As Long
Dim h As Long, v As Long
Dim XC As Long, YC As Long
Dim tp(2) As POINTAPI
Dim ptPoint As POINTAPI
On Error Resume Next
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
       If Not .OutOfScope(I) Then
          h = .R(I) \ 2
          v = .R(I) \ 2
          tp(0).X = .Y(I) - v / 2:    tp(0).Y = .X(I) - h
          tp(1).X = .Y(I) + v / 2:    tp(1).Y = .X(I)
          tp(2).X = .Y(I) - v / 2:     tp(2).Y = .X(I) + h
          XC = CLng((tp(0).X + tp(1).X + tp(2).X) / 3)
          YC = CLng((tp(0).Y + tp(1).Y + tp(2).Y) / 3)
          Res = MoveToEx(hDC, tp(0).Y, tp(0).X, ptPoint)
          Res = LineTo(hDC, YC, XC)
          Res = LineTo(hDC, tp(1).Y, tp(1).X)
          Res = MoveToEx(hDC, tp(2).Y, tp(2).X, ptPoint)
          Res = LineTo(hDC, YC, XC)
       End If
   Next I
End With
End Sub


Public Sub OlyUMCDrawBox(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
On Error Resume Next
I = SelectObject(hDC, hOlyClrBrush(OlyInd))
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
       If Not .OutOfScope(I) Then Rectangle hDC, .XL(I), .Y(I), .XU(I), .YU(I)
   Next I
End With
End Sub


Public Sub OlyUMCDrawEmptyBox(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
On Error Resume Next
I = SelectObject(hDC, GetStockObject(NULL_BRUSH))
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
       If Not .OutOfScope(I) Then Rectangle hDC, .XL(I), .Y(I), .XU(I), .YU(I)
   Next I
End With
End Sub


Public Sub OlyUMCDrawSpot(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
On Error Resume Next
I = SelectObject(hDC, hOlyClrBrush(OlyInd))
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
       If Not .OutOfScope(I) Then Ellipse hDC, .XL(I), .Y(I), .XU(I), .YU(I)
   Next I
End With
End Sub


Public Sub OlyUMCDrawEmptySpot(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
On Error Resume Next
I = SelectObject(hDC, GetStockObject(NULL_BRUSH))
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
       If Not .OutOfScope(I) Then Ellipse hDC, .XL(I), .Y(I), .XU(I), .YU(I)
   Next I
End With
End Sub


Public Sub OlyUMCDrawTriangle(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
Dim TriPoints(2) As POINTAPI
On Error Resume Next
I = SelectObject(hDC, hOlyClrBrush(OlyInd))
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
       If Not .OutOfScope(I) Then
          TriPoints(0).X = .XL(I):      TriPoints(0).Y = .YL(I)
          TriPoints(1).X = .X(I):      TriPoints(1).Y = .Y(I)
          TriPoints(2).X = .XU(I):      TriPoints(2).Y = .YU(I)
          Polygon hDC, TriPoints(0), 3
       End If
   Next I
End With
End Sub


Public Sub OlyUMCDrawEmptyTriangle(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
Dim TriPoints(2) As POINTAPI
On Error Resume Next
I = SelectObject(hDC, GetStockObject(NULL_BRUSH))
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
       If Not .OutOfScope(I) Then
          TriPoints(0).X = .XL(I):      TriPoints(0).Y = .YL(I)
          TriPoints(1).X = .X(I):      TriPoints(1).Y = .Y(I)
          TriPoints(2).X = .XU(I):      TriPoints(2).Y = .YU(I)
          Polygon hDC, TriPoints(0), 3
       End If
   Next I
End With
End Sub


Public Sub OlyUMCDrawTriStar(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long, Res As Long
Dim XC As Long, YC As Long
Dim ptPoint As POINTAPI
On Error Resume Next
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
       If Not .OutOfScope(I) Then
          XC = CLng((.XL(I) + .X(I) + .XU(I)) / 3)
          YC = CLng((.YL(I) + .Y(I) + .YU(I)) / 3)
          Res = MoveToEx(hDC, .XL(I), .YL(I), ptPoint)
          Res = LineTo(hDC, XC, YC)
          Res = LineTo(hDC, .XU(I), .YU(I))
          Res = MoveToEx(hDC, .X(I), .Y(I), ptPoint)
          Res = LineTo(hDC, XC, YC)
       End If
   Next I
End With
End Sub


Public Sub OlyUMCDrawStick(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
Dim Res As Long
Dim OldPen As Long
Dim ptPoint As POINTAPI
On Error Resume Next
Res = SelectObject(hDC, GetStockObject(NULL_BRUSH))
OldPen = SelectObject(hDC, hOlyClrPenStick(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
       If Not .OutOfScope(I) Then
          Res = MoveToEx(hDC, .XL(I), .YL(I), ptPoint)
          Res = LineTo(hDC, .XU(I), .YU(I))
       End If
   Next I
End With
Res = SelectObject(hDC, OldPen)
End Sub


Public Sub OlyUMCDrawBoxNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
On Error Resume Next
I = SelectObject(hDC, hOlyClrBrush(OlyInd))
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
       If Not .OutOfScope(I) Then Rectangle hDC, .Y(I), .XL(I), .YU(I), .XU(I)
   Next I
End With
End Sub


Public Sub OlyUMCDrawEmptyBoxNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
On Error Resume Next
I = SelectObject(hDC, GetStockObject(NULL_BRUSH))
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
       If Not .OutOfScope(I) Then Rectangle hDC, .Y(I), .XL(I), .YU(I), .XU(I)
   Next I
End With
End Sub


Public Sub OlyUMCDrawSpotNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
On Error Resume Next
I = SelectObject(hDC, hOlyClrBrush(OlyInd))
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
       If Not .OutOfScope(I) Then Ellipse hDC, .Y(I), .XL(I), .YU(I), .XU(I)
   Next I
End With
End Sub


Public Sub OlyUMCDrawEmptySpotNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
On Error Resume Next
I = SelectObject(hDC, GetStockObject(NULL_BRUSH))
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
       If Not .OutOfScope(I) Then Ellipse hDC, .Y(I), .XL(I), .YU(I), .XU(I)
   Next I
End With
End Sub

Public Sub OlyUMCDrawStickNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
Dim Res As Long
Dim OldPen As Long
Dim ptPoint As POINTAPI
On Error Resume Next
Res = SelectObject(hDC, GetStockObject(NULL_BRUSH))
OldPen = SelectObject(hDC, hOlyClrPenStick(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
       If Not .OutOfScope(I) Then
          Res = MoveToEx(hDC, .YL(I), .XL(I), ptPoint)
          Res = LineTo(hDC, .YU(I), .XU(I))
       End If
   Next I
End With
Res = SelectObject(hDC, OldPen)
End Sub


Public Sub OlyUMCDrawTriangleNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
Dim TriPoints(2) As POINTAPI
On Error Resume Next
I = SelectObject(hDC, hOlyClrBrush(OlyInd))
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
       If Not .OutOfScope(I) Then
          TriPoints(0).X = .YL(I):      TriPoints(0).Y = .XL(I)
          TriPoints(1).X = .Y(I):      TriPoints(1).Y = .X(I)
          TriPoints(2).X = .YU(I):      TriPoints(2).Y = .XU(I)
          Polygon hDC, TriPoints(0), 3
       End If
   Next I
End With
End Sub

Public Sub OlyUMCDrawEmptyTriangleNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long
Dim TriPoints(2) As POINTAPI
On Error Resume Next
I = SelectObject(hDC, GetStockObject(NULL_BRUSH))
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
       If Not .OutOfScope(I) Then
          TriPoints(0).X = .YL(I):      TriPoints(0).Y = .XL(I)
          TriPoints(1).X = .Y(I):      TriPoints(1).Y = .X(I)
          TriPoints(2).X = .YU(I):      TriPoints(2).Y = .XU(I)
          Polygon hDC, TriPoints(0), 3
       End If
   Next I
End With
End Sub

Public Sub OlyUMCDrawTriStarNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim I As Long, Res As Long
Dim XC As Long, YC As Long
Dim ptPoint As POINTAPI
On Error Resume Next
I = SelectObject(hDC, hOlyClrPen(OlyInd))
I = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
       If Not .OutOfScope(I) Then
          XC = CLng((.XL(I) + .X(I) + .XU(I)) / 3)
          YC = CLng((.YL(I) + .Y(I) + .YU(I)) / 3)
          Res = MoveToEx(hDC, .YL(I), .XL(I), ptPoint)
          Res = LineTo(hDC, YC, XC)
          Res = LineTo(hDC, .YU(I), .XU(I))
          Res = MoveToEx(hDC, .Y(I), .X(I), ptPoint)
          Res = LineTo(hDC, YC, XC)
       End If
   Next I
End With
End Sub



Public Sub OlyDrawTextHorz(ByVal OlyInd As Long, ByVal hDC As Long, _
                           ByVal TextHeight As Long, ByVal TextWidth As Long)
Dim I As Long, Res As Long
Dim lNewFont As Long, lOldFont As Long
Dim lfFnt As LOGFONT
Dim lOldBrush As Long
Dim szText As Size
On Error Resume Next
lOldFont = SelectObject(hDC, GetStockObject(SYSTEM_FONT))
Res = GetObjectAPI(lOldFont, Len(lfFnt), lfFnt)
Res = SelectObject(hDC, lOldFont)
lfFnt.lfWidth = TextWidth
lfFnt.lfHeight = TextHeight
If lNewFont <> 0 Then DeleteObject (lNewFont)
lNewFont = CreateFontIndirect(lfFnt)
Res = SetBkMode(hDC, TRANSPARENT)
Res = SetTextColor(hDC, Oly(OlyInd).Color)
lOldBrush = SelectObject(hDC, GetStockObject(NULL_BRUSH))
lOldFont = SelectObject(hDC, lNewFont)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
      If Not .OutOfScope(I) Then
         If .Visible(I) And Len(.Text(I)) > 0 Then
            Res = GetTextExtentPoint32(hDC, .Text(I), Len(.Text(I)), szText)
            Res = TextOut(hDC, .X(I) - CLng(szText.cx / 2), .Y(I) - CLng(TextHeight / 12), .Text(I), Len(.Text(I)))
         End If
      End If
   Next I
End With
Res = SelectObject(hDC, lOldBrush)
Res = SelectObject(hDC, lOldFont)
If lNewFont <> 0 Then DeleteObject (lNewFont)
End Sub


Public Sub OlyDrawTextH_MM(ByVal OlyInd As Long, ByVal hDC As Long, ByVal TextHeight As Double, _
                           ByVal TextWidth As Double, ByVal TextOffset As Double)
'------------------------------------------------------------------------------------------------
'TextHeight is text height in millimeters, TextWidth determines the text width
'------------------------------------------------------------------------------------------------
Dim I As Long, Res As Long
Dim lNewFont As Long, lOldFont As Long
Dim lfFnt As LOGFONT
Dim lOldBrush As Long
Dim szText As Size
On Error Resume Next
lOldFont = SelectObject(hDC, GetStockObject(SYSTEM_FONT))
Res = GetObjectAPI(lOldFont, Len(lfFnt), lfFnt)
Res = SelectObject(hDC, lOldFont)
'here set font properties if you wish
lfFnt.lfHeight = CLng(TextHeight)
lfFnt.lfWidth = CLng(TextWidth)
If lNewFont <> 0 Then DeleteObject (lNewFont)
lNewFont = CreateFontIndirect(lfFnt)
Res = SetBkMode(hDC, TRANSPARENT)
Res = SetTextColor(hDC, Oly(OlyInd).Color)
lOldBrush = SelectObject(hDC, GetStockObject(NULL_BRUSH))
lOldFont = SelectObject(hDC, lNewFont)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
      If Not .OutOfScope(I) Then
         If .Visible(I) And Len(.Text(I)) > 0 Then                         'print under the actual spot
            Res = GetTextExtentPoint32(hDC, .Text(I), Len(.Text(I)), szText)
            Res = TextOut(hDC, .X(I) - CLng(szText.cx / 2), .Y(I) - TextOffset, .Text(I), Len(.Text(I)))
         End If
      End If
   Next I
End With
Res = SelectObject(hDC, lOldBrush)
Res = SelectObject(hDC, lOldFont)
If lNewFont <> 0 Then DeleteObject (lNewFont)
End Sub


Public Sub OlyDrawTextVert(ByVal OlyInd As Long, ByVal hDC As Long, _
                           ByVal TextHeight As Long, ByVal TextWidth As Long)
Dim I As Long, Res As Long
Dim lNewFont As Long, lOldFont As Long
Dim lfFnt As LOGFONT
Dim lOldBrush As Long
Dim szText As Size
On Error Resume Next
lOldFont = SelectObject(hDC, GetStockObject(SYSTEM_FONT))
Res = GetObjectAPI(lOldFont, Len(lfFnt), lfFnt)
Res = SelectObject(hDC, lOldFont)
lfFnt.lfWidth = TextWidth
lfFnt.lfHeight = TextHeight
lfFnt.lfEscapement = 900
'here set font properties if you wish
If lNewFont <> 0 Then DeleteObject (lNewFont)
lNewFont = CreateFontIndirect(lfFnt)
Res = SetBkMode(hDC, TRANSPARENT)
Res = SetTextColor(hDC, Oly(OlyInd).Color)
lOldBrush = SelectObject(hDC, GetStockObject(NULL_BRUSH))
lOldFont = SelectObject(hDC, lNewFont)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
      If Not .OutOfScope(I) Then
         If .Visible(I) And Len(.Text(I)) > 0 Then
            Res = GetTextExtentPoint32(hDC, .Text(I), Len(.Text(I)), szText)
            Res = TextOut(hDC, .Y(I) - CLng(TextHeight / 12), .X(I) + CLng(szText.cx / 2), .Text(I), Len(.Text(I)))
         End If
      End If
   Next I
End With
Res = SelectObject(hDC, lOldBrush)
Res = SelectObject(hDC, lOldFont)
If lNewFont <> 0 Then DeleteObject (lNewFont)
End Sub



Public Sub OlyDrawTextV_MM(ByVal OlyInd As Long, ByVal hDC As Long, ByVal TextHeight As Double, _
                           ByVal TextWidth As Double, ByVal TextOffset As Double)
Dim I As Long, Res As Long
Dim lNewFont As Long, lOldFont As Long
Dim lfFnt As LOGFONT
Dim lOldBrush As Long
Dim szText As Size
On Error Resume Next
lOldFont = SelectObject(hDC, GetStockObject(SYSTEM_FONT))       'get currently selected font
Res = GetObjectAPI(lOldFont, Len(lfFnt), lfFnt)                 'fill data to LOGFONT structure
Res = SelectObject(hDC, lOldFont)                               'release stock object
'here set font properties if you wish
lfFnt.lfHeight = CLng(TextHeight)
lfFnt.lfWidth = CLng(TextWidth)
lfFnt.lfEscapement = 900                                       'vertical font drawing
If lNewFont <> 0 Then DeleteObject (lNewFont)
lNewFont = CreateFontIndirect(lfFnt)
Res = SetBkMode(hDC, TRANSPARENT)
Res = SetTextColor(hDC, Oly(OlyInd).Color)
lOldBrush = SelectObject(hDC, GetStockObject(NULL_BRUSH))
lOldFont = SelectObject(hDC, lNewFont)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
      If Not .OutOfScope(I) Then
         If .Visible(I) And Len(.Text(I)) > 0 Then              'print left of actual spot
            Res = GetTextExtentPoint32(hDC, .Text(I), Len(.Text(I)), szText)
            Res = TextOut(hDC, .Y(I) - TextOffset, .X(I) + CLng(szText.cx / 2), .Text(I), Len(.Text(I)))
         End If
      End If
   Next I
End With
Res = SelectObject(hDC, lOldBrush)
Res = SelectObject(hDC, lOldFont)
DeleteObject (lNewFont)
End Sub


'patch to work metric drawing for stick without changing pen weight
Public Sub OlyDrawStickNETVert_MM(ByVal OlyInd As Long, ByVal hDC As Long, ByVal HalfWidth As Long)
Dim I As Long
Dim Res As Long
Dim OldPen As Long
Dim OldBrush As Long
On Error Resume Next
OldBrush = SelectObject(hDC, hOlyClrBrush(OlyInd))
OldPen = SelectObject(hDC, hOlyClrPenStick(OlyInd))
Res = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
      If .Visible(I) Then
         If Not .OutOfScope(I) Then
            Rectangle hDC, .Y(I) - HalfWidth, .X(I) - .XL(I), .Y(I) + HalfWidth, .X(I) + .XL(I)
         End If
      End If
   Next I
End With
Res = SelectObject(hDC, OldPen)
Res = SelectObject(hDC, OldBrush)
End Sub

Public Sub OlyDrawStickNETHorz_MM(ByVal OlyInd As Long, ByVal hDC As Long, ByVal HalfWidth As Long)
Dim I As Long
Dim Res As Long
Dim OldPen As Long
Dim OldBrush As Long
On Error Resume Next
OldBrush = SelectObject(hDC, hOlyClrBrush(OlyInd))
OldPen = SelectObject(hDC, hOlyClrPenStick(OlyInd))
Res = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
      If .Visible(I) Then
         If Not .OutOfScope(I) Then
            Rectangle hDC, .X(I) - .XL(I), .Y(I) - HalfWidth, .X(I) + .XL(I), .Y(I) + HalfWidth
         End If
      End If
   Next I
End With
Res = SelectObject(hDC, OldPen)
Res = SelectObject(hDC, OldBrush)
End Sub

Public Sub OlyUMCDrawStickHorz_MM(ByVal OlyInd As Long, ByVal hDC As Long, ByVal HalfWidth As Long)
Dim I As Long
Dim Res As Long
Dim OldPen As Long
Dim OldBrush As Long
On Error Resume Next
OldBrush = SelectObject(hDC, hOlyClrBrush(OlyInd))
OldPen = SelectObject(hDC, hOlyClrPenStick(OlyInd))
Res = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
       If Not .OutOfScope(I) Then
          Rectangle hDC, .XL(I), .Y(I) - HalfWidth, .XU(I), .Y(I) + HalfWidth
       End If
   Next I
End With
Res = SelectObject(hDC, OldPen)
Res = SelectObject(hDC, OldBrush)
End Sub

Public Sub OlyUMCDrawStickVert_MM(ByVal OlyInd As Long, ByVal hDC As Long, ByVal HalfWidth As Long)
Dim I As Long
Dim Res As Long
Dim OldPen As Long
Dim OldBrush As Long
On Error Resume Next
OldBrush = SelectObject(hDC, hOlyClrBrush(OlyInd))
OldPen = SelectObject(hDC, hOlyClrPenStick(OlyInd))
Res = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For I = 0 To .DataCnt - 1
       If Not .OutOfScope(I) Then
          Rectangle hDC, .Y(I) - HalfWidth, .XL(I), .Y(I) + HalfWidth, .XU(I)
       End If
   Next I
End With
Res = SelectObject(hDC, OldPen)
Res = SelectObject(hDC, OldBrush)
End Sub

