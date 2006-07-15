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

' Unused Function (May 2003)
'''Public Sub DestroyOlyDrawingObjects()
'''On Error Resume Next
'''Dim i As Integer
'''If hOlyBackClrBrush <> 0 Then DeleteObject (hOlyBackClrBrush)
'''If hOlyForeClrPen <> 0 Then DeleteObject (hOlyForeClrPen)
'''If hOlyGridClrPen <> 0 Then DeleteObject (hOlyGridClrPen)
'''For i = 0 To OlyCnt - 1
'''    If hOlyClrBrush(i) <> 0 Then DeleteObject (hOlyClrBrush(i))
'''    If hOlyClrPen(i) <> 0 Then DeleteObject (hOlyClrPen(i))
'''    If hOlyClrPenStick(i) <> 0 Then DeleteObject (hOlyClrPenStick(i))
'''Next i
'''End Sub

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
Dim i As Long
On Error Resume Next
OlyClrCnt = UBound(hOlyClrBrush) - 1
If hOlyClrBrush(OlyInd) <> 0 Then DeleteObject (hOlyClrBrush(OlyInd))
If hOlyClrPen(OlyInd) <> 0 Then DeleteObject (hOlyClrPen(OlyInd))
If hOlyClrPenStick(OlyInd) <> 0 Then DeleteObject (hOlyClrPenStick(OlyInd))
For i = OlyInd To OlyClrCnt - 1
    hOlyClrBrush(i) = hOlyClrBrush(i + 1)
    hOlyClrPen(i) = hOlyClrPen(i + 1)
    hOlyClrPenStick(i) = hOlyClrPenStick(i + 1)
Next i
ReDim Preserve hOlyClrPen(OlyClrCnt - 1)
ReDim Preserve hOlyClrPenStick(OlyClrCnt - 1)
ReDim Preserve hOlyClrBrush(OlyClrCnt - 1)
End Function
'-------------------------------------------------------------------



'drawing functions
Public Sub OlyDrawBox(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
On Error Resume Next
i = SelectObject(hDC, hOlyClrBrush(OlyInd))
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
      If .Visible(i) Then
        If Not .OutOfScope(i) Then
           Rectangle hDC, .x(i) - .XL(i), .y(i) - .YL(i), _
                       .x(i) + .XU(i), .y(i) + .YU(i)
        End If
      End If
   Next i
End With
End Sub


Public Sub OlyDrawEmptyBox(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
On Error Resume Next
i = SelectObject(hDC, GetStockObject(NULL_BRUSH))
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
      If .Visible(i) Then
        If Not .OutOfScope(i) Then
           Rectangle hDC, .x(i) - .XL(i), .y(i) - .YL(i), _
                       .x(i) + .XU(i), .y(i) + .YU(i)
        End If
      End If
   Next i
End With
End Sub


Public Sub OlyDrawSpot(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
Dim H As Long, V As Long
On Error Resume Next
i = SelectObject(hDC, hOlyClrBrush(OlyInd))
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
      If .Visible(i) Then
         If Not .OutOfScope(i) Then
            H = .R(i) \ 2
            V = .R(i) \ 2
            Ellipse hDC, .x(i) - H, .y(i) - V, .x(i) + H, .y(i) + V
         End If
      End If
   Next i
End With
End Sub


Public Sub OlyDrawEmptySpot(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
Dim H As Long, V As Long
On Error Resume Next
i = SelectObject(hDC, GetStockObject(NULL_BRUSH))
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
      If .Visible(i) Then
         If Not .OutOfScope(i) Then
            H = .R(i) \ 2
            V = .R(i) \ 2
            Ellipse hDC, .x(i) - H, .y(i) - V, .x(i) + H, .y(i) + V
         End If
      End If
   Next i
End With
End Sub


Public Sub OlyDrawTriangle(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
Dim H As Long, V As Long
Dim TriPoints(2) As POINTAPI
On Error Resume Next
i = SelectObject(hDC, hOlyClrBrush(OlyInd))
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
      If .Visible(i) Then
         If Not .OutOfScope(i) Then
            H = .R(i) \ 2
            V = .R(i) \ 2
            TriPoints(0).x = .x(i) - H:    TriPoints(0).y = .y(i) - V / 2
            TriPoints(1).x = .x(i):      TriPoints(1).y = .y(i) + V / 2
            TriPoints(2).x = .x(i) + H:      TriPoints(2).y = .y(i) - V / 2
            Polygon hDC, TriPoints(0), 3
         End If
      End If
   Next i
End With
End Sub


Public Sub OlyDrawEmptyTriangle(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
Dim H As Long, V As Long
Dim TriPoints(2) As POINTAPI
On Error Resume Next
i = SelectObject(hDC, GetStockObject(NULL_BRUSH))
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
      If .Visible(i) Then
         If Not .OutOfScope(i) Then
            H = .R(i) \ 2
            V = .R(i) \ 2
            TriPoints(0).x = .x(i) - H:    TriPoints(0).y = .y(i) - V / 2
            TriPoints(1).x = .x(i):      TriPoints(1).y = .y(i) + V / 2
            TriPoints(2).x = .x(i) + H:      TriPoints(2).y = .y(i) - V / 2
            Polygon hDC, TriPoints(0), 3
         End If
      End If
   Next i
End With
End Sub


Public Sub OlyDrawStick(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
Dim Res As Long
Dim OldPen As Long
Dim ptPoint As POINTAPI
On Error Resume Next
Res = SelectObject(hDC, GetStockObject(NULL_BRUSH))
OldPen = SelectObject(hDC, hOlyClrPenStick(OlyInd))
Res = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
      If .Visible(i) Then
         If Not .OutOfScope(i) Then
            Res = MoveToEx(hDC, .x(i) - .XL(i), .y(i), ptPoint)
            Res = LineTo(hDC, .x(i) + .XU(i), .y(i))
         End If
      End If
   Next i
End With
Res = SelectObject(hDC, OldPen)
End Sub


Public Sub OlyDrawTriStar(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long, Res As Long
Dim H As Long, V As Long
Dim XC As Long, YC As Long
Dim TP(2) As POINTAPI
Dim ptPoint As POINTAPI
On Error Resume Next
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
       If Not .OutOfScope(i) Then
          H = .R(i) \ 2
          V = .R(i) \ 2
          TP(0).x = .x(i) - H:    TP(0).y = .y(i) - V / 2
          TP(1).x = .x(i):    TP(1).y = .y(i) + V / 2
          TP(2).x = .x(i) + H:     TP(2).y = .y(i) - V / 2
          XC = CLng((TP(0).x + TP(1).x + TP(2).x) / 3)
          YC = CLng((TP(0).y + TP(1).y + TP(2).y) / 3)
          Res = MoveToEx(hDC, TP(0).x, TP(0).y, ptPoint)
          Res = LineTo(hDC, XC, YC)
          Res = LineTo(hDC, TP(1).x, TP(1).y)
          Res = MoveToEx(hDC, TP(2).x, TP(2).y, ptPoint)
          Res = LineTo(hDC, XC, YC)
       End If
   Next i
End With
End Sub


Public Sub OlyDrawBoxNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
On Error Resume Next
i = SelectObject(hDC, hOlyClrBrush(OlyInd))
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
      If .Visible(i) Then
         If Not .OutOfScope(i) Then
            Rectangle hDC, .y(i) - .YL(i), .x(i) - .XL(i), _
                        .y(i) + .YU(i), .x(i) + .XU(i)
         End If
      End If
   Next i
End With
End Sub


Public Sub OlyDrawEmptyBoxNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
On Error Resume Next
i = SelectObject(hDC, GetStockObject(NULL_BRUSH))
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
      If .Visible(i) Then
         If Not .OutOfScope(i) Then
            Rectangle hDC, .y(i) - .YL(i), .x(i) - .XL(i), _
                        .y(i) + .YU(i), .x(i) + .XU(i)
         End If
      End If
   Next i
End With
End Sub


Public Sub OlyDrawSpotNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
Dim H As Long, V As Long
On Error Resume Next
i = SelectObject(hDC, hOlyClrBrush(OlyInd))
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
      If .Visible(i) Then
         If Not .OutOfScope(i) Then
            H = .R(i) \ 2
            V = .R(i) \ 2
            Ellipse hDC, .y(i) - V, .x(i) - H, .y(i) + V, .x(i) + H
         End If
      End If
   Next i
End With
End Sub


Public Sub OlyDrawEmptySpotNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
Dim H As Long, V As Long
On Error Resume Next
i = SelectObject(hDC, GetStockObject(NULL_BRUSH))
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
      If .Visible(i) Then
         If Not .OutOfScope(i) Then
            H = .R(i) \ 2
            V = .R(i) \ 2
            Ellipse hDC, .y(i) - V, .x(i) - H, .y(i) + V, .x(i) + H
         End If
      End If
   Next i
End With
End Sub


Public Sub OlyDrawStickNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
Dim Res As Long
Dim OldPen As Long
Dim ptPoint As POINTAPI
On Error Resume Next
Res = SelectObject(hDC, GetStockObject(NULL_BRUSH))
OldPen = SelectObject(hDC, hOlyClrPenStick(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
      If .Visible(i) Then
         If Not .OutOfScope(i) Then
            Res = MoveToEx(hDC, .y(i), .x(i) - .XL(i), ptPoint)
            Res = LineTo(hDC, .y(i), .x(i) + .XU(i))
         End If
      End If
   Next i
End With
Res = SelectObject(hDC, OldPen)
End Sub


Public Sub OlyDrawTriangleNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
Dim H As Long, V As Long
Dim TriPoints(2) As POINTAPI
On Error Resume Next
i = SelectObject(hDC, hOlyClrBrush(OlyInd))
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
      If .Visible(i) Then
         If Not .OutOfScope(i) Then
            H = .R(i) \ 2
            V = .R(i) \ 2
            TriPoints(0).x = .y(i) - V / 2:    TriPoints(0).y = .x(i) - H
            TriPoints(1).x = .y(i) + V / 2:      TriPoints(1).y = .x(i)
            TriPoints(2).x = .y(i) - V / 2:      TriPoints(2).y = .x(i) + H
            Polygon hDC, TriPoints(0), 3
         End If
      End If
   Next i
End With
End Sub


Public Sub OlyDrawEmptyTriangleNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
Dim H As Long, V As Long
Dim TriPoints(2) As POINTAPI
On Error Resume Next
i = SelectObject(hDC, GetStockObject(NULL_BRUSH))
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
      If .Visible(i) Then
         If Not .OutOfScope(i) Then
            H = .R(i) \ 2
            V = .R(i) \ 2
            TriPoints(0).x = .y(i) - V / 2:    TriPoints(0).y = .x(i) - H
            TriPoints(1).x = .y(i) + V / 2:      TriPoints(1).y = .x(i)
            TriPoints(2).x = .y(i) - V / 2:      TriPoints(2).y = .x(i) + H
            Polygon hDC, TriPoints(0), 3
         End If
      End If
   Next i
End With
End Sub


Public Sub OlyDrawTriStarNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long, Res As Long
Dim H As Long, V As Long
Dim XC As Long, YC As Long
Dim TP(2) As POINTAPI
Dim ptPoint As POINTAPI
On Error Resume Next
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
       If Not .OutOfScope(i) Then
          H = .R(i) \ 2
          V = .R(i) \ 2
          TP(0).x = .y(i) - V / 2:    TP(0).y = .x(i) - H
          TP(1).x = .y(i) + V / 2:    TP(1).y = .x(i)
          TP(2).x = .y(i) - V / 2:     TP(2).y = .x(i) + H
          XC = CLng((TP(0).x + TP(1).x + TP(2).x) / 3)
          YC = CLng((TP(0).y + TP(1).y + TP(2).y) / 3)
          Res = MoveToEx(hDC, TP(0).y, TP(0).x, ptPoint)
          Res = LineTo(hDC, YC, XC)
          Res = LineTo(hDC, TP(1).y, TP(1).x)
          Res = MoveToEx(hDC, TP(2).y, TP(2).x, ptPoint)
          Res = LineTo(hDC, YC, XC)
       End If
   Next i
End With
End Sub


Public Sub OlyUMCDrawBox(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
On Error Resume Next
i = SelectObject(hDC, hOlyClrBrush(OlyInd))
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
       If Not .OutOfScope(i) Then Rectangle hDC, .XL(i), .y(i), .XU(i), .YU(i)
   Next i
End With
End Sub


Public Sub OlyUMCDrawEmptyBox(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
On Error Resume Next
i = SelectObject(hDC, GetStockObject(NULL_BRUSH))
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
       If Not .OutOfScope(i) Then Rectangle hDC, .XL(i), .y(i), .XU(i), .YU(i)
   Next i
End With
End Sub


Public Sub OlyUMCDrawSpot(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
On Error Resume Next
i = SelectObject(hDC, hOlyClrBrush(OlyInd))
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
       If Not .OutOfScope(i) Then Ellipse hDC, .XL(i), .y(i), .XU(i), .YU(i)
   Next i
End With
End Sub


Public Sub OlyUMCDrawEmptySpot(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
On Error Resume Next
i = SelectObject(hDC, GetStockObject(NULL_BRUSH))
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
       If Not .OutOfScope(i) Then Ellipse hDC, .XL(i), .y(i), .XU(i), .YU(i)
   Next i
End With
End Sub


Public Sub OlyUMCDrawTriangle(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
Dim TriPoints(2) As POINTAPI
On Error Resume Next
i = SelectObject(hDC, hOlyClrBrush(OlyInd))
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
       If Not .OutOfScope(i) Then
          TriPoints(0).x = .XL(i):      TriPoints(0).y = .YL(i)
          TriPoints(1).x = .x(i):      TriPoints(1).y = .y(i)
          TriPoints(2).x = .XU(i):      TriPoints(2).y = .YU(i)
          Polygon hDC, TriPoints(0), 3
       End If
   Next i
End With
End Sub


Public Sub OlyUMCDrawEmptyTriangle(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
Dim TriPoints(2) As POINTAPI
On Error Resume Next
i = SelectObject(hDC, GetStockObject(NULL_BRUSH))
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
       If Not .OutOfScope(i) Then
          TriPoints(0).x = .XL(i):      TriPoints(0).y = .YL(i)
          TriPoints(1).x = .x(i):      TriPoints(1).y = .y(i)
          TriPoints(2).x = .XU(i):      TriPoints(2).y = .YU(i)
          Polygon hDC, TriPoints(0), 3
       End If
   Next i
End With
End Sub


Public Sub OlyUMCDrawTriStar(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long, Res As Long
Dim XC As Long, YC As Long
Dim ptPoint As POINTAPI
On Error Resume Next
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
       If Not .OutOfScope(i) Then
          XC = CLng((.XL(i) + .x(i) + .XU(i)) / 3)
          YC = CLng((.YL(i) + .y(i) + .YU(i)) / 3)
          Res = MoveToEx(hDC, .XL(i), .YL(i), ptPoint)
          Res = LineTo(hDC, XC, YC)
          Res = LineTo(hDC, .XU(i), .YU(i))
          Res = MoveToEx(hDC, .x(i), .y(i), ptPoint)
          Res = LineTo(hDC, XC, YC)
       End If
   Next i
End With
End Sub


Public Sub OlyUMCDrawStick(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
Dim Res As Long
Dim OldPen As Long
Dim ptPoint As POINTAPI
On Error Resume Next
Res = SelectObject(hDC, GetStockObject(NULL_BRUSH))
OldPen = SelectObject(hDC, hOlyClrPenStick(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
       If Not .OutOfScope(i) Then
          Res = MoveToEx(hDC, .XL(i), .YL(i), ptPoint)
          Res = LineTo(hDC, .XU(i), .YU(i))
       End If
   Next i
End With
Res = SelectObject(hDC, OldPen)
End Sub


Public Sub OlyUMCDrawBoxNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
On Error Resume Next
i = SelectObject(hDC, hOlyClrBrush(OlyInd))
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
       If Not .OutOfScope(i) Then Rectangle hDC, .y(i), .XL(i), .YU(i), .XU(i)
   Next i
End With
End Sub


Public Sub OlyUMCDrawEmptyBoxNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
On Error Resume Next
i = SelectObject(hDC, GetStockObject(NULL_BRUSH))
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
       If Not .OutOfScope(i) Then Rectangle hDC, .y(i), .XL(i), .YU(i), .XU(i)
   Next i
End With
End Sub


Public Sub OlyUMCDrawSpotNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
On Error Resume Next
i = SelectObject(hDC, hOlyClrBrush(OlyInd))
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
       If Not .OutOfScope(i) Then Ellipse hDC, .y(i), .XL(i), .YU(i), .XU(i)
   Next i
End With
End Sub


Public Sub OlyUMCDrawEmptySpotNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
On Error Resume Next
i = SelectObject(hDC, GetStockObject(NULL_BRUSH))
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
       If Not .OutOfScope(i) Then Ellipse hDC, .y(i), .XL(i), .YU(i), .XU(i)
   Next i
End With
End Sub

Public Sub OlyUMCDrawStickNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
Dim Res As Long
Dim OldPen As Long
Dim ptPoint As POINTAPI
On Error Resume Next
Res = SelectObject(hDC, GetStockObject(NULL_BRUSH))
OldPen = SelectObject(hDC, hOlyClrPenStick(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
       If Not .OutOfScope(i) Then
          Res = MoveToEx(hDC, .YL(i), .XL(i), ptPoint)
          Res = LineTo(hDC, .YU(i), .XU(i))
       End If
   Next i
End With
Res = SelectObject(hDC, OldPen)
End Sub


Public Sub OlyUMCDrawTriangleNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
Dim TriPoints(2) As POINTAPI
On Error Resume Next
i = SelectObject(hDC, hOlyClrBrush(OlyInd))
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
       If Not .OutOfScope(i) Then
          TriPoints(0).x = .YL(i):      TriPoints(0).y = .XL(i)
          TriPoints(1).x = .y(i):      TriPoints(1).y = .x(i)
          TriPoints(2).x = .YU(i):      TriPoints(2).y = .XU(i)
          Polygon hDC, TriPoints(0), 3
       End If
   Next i
End With
End Sub

Public Sub OlyUMCDrawEmptyTriangleNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long
Dim TriPoints(2) As POINTAPI
On Error Resume Next
i = SelectObject(hDC, GetStockObject(NULL_BRUSH))
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
       If Not .OutOfScope(i) Then
          TriPoints(0).x = .YL(i):      TriPoints(0).y = .XL(i)
          TriPoints(1).x = .y(i):      TriPoints(1).y = .x(i)
          TriPoints(2).x = .YU(i):      TriPoints(2).y = .XU(i)
          Polygon hDC, TriPoints(0), 3
       End If
   Next i
End With
End Sub

Public Sub OlyUMCDrawTriStarNETVert(ByVal OlyInd As Long, ByVal hDC As Long)
Dim i As Long, Res As Long
Dim XC As Long, YC As Long
Dim ptPoint As POINTAPI
On Error Resume Next
i = SelectObject(hDC, hOlyClrPen(OlyInd))
i = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
       If Not .OutOfScope(i) Then
          XC = CLng((.XL(i) + .x(i) + .XU(i)) / 3)
          YC = CLng((.YL(i) + .y(i) + .YU(i)) / 3)
          Res = MoveToEx(hDC, .YL(i), .XL(i), ptPoint)
          Res = LineTo(hDC, YC, XC)
          Res = LineTo(hDC, .YU(i), .XU(i))
          Res = MoveToEx(hDC, .y(i), .x(i), ptPoint)
          Res = LineTo(hDC, YC, XC)
       End If
   Next i
End With
End Sub



Public Sub OlyDrawTextHorz(ByVal OlyInd As Long, ByVal hDC As Long, _
                           ByVal TextHeight As Long, ByVal TextWidth As Long)
Dim i As Long, Res As Long
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
   For i = 0 To .DataCnt - 1
      If Not .OutOfScope(i) Then
         If .Visible(i) And Len(.Text(i)) > 0 Then
            Res = GetTextExtentPoint32(hDC, .Text(i), Len(.Text(i)), szText)
            Res = TextOut(hDC, .x(i) - CLng(szText.cx / 2), .y(i) - CLng(TextHeight / 12), .Text(i), Len(.Text(i)))
         End If
      End If
   Next i
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
Dim i As Long, Res As Long
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
   For i = 0 To .DataCnt - 1
      If Not .OutOfScope(i) Then
         If .Visible(i) And Len(.Text(i)) > 0 Then                         'print under the actual spot
            Res = GetTextExtentPoint32(hDC, .Text(i), Len(.Text(i)), szText)
            Res = TextOut(hDC, .x(i) - CLng(szText.cx / 2), .y(i) - TextOffset, .Text(i), Len(.Text(i)))
         End If
      End If
   Next i
End With
Res = SelectObject(hDC, lOldBrush)
Res = SelectObject(hDC, lOldFont)
If lNewFont <> 0 Then DeleteObject (lNewFont)
End Sub


Public Sub OlyDrawTextVert(ByVal OlyInd As Long, ByVal hDC As Long, _
                           ByVal TextHeight As Long, ByVal TextWidth As Long)
Dim i As Long, Res As Long
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
   For i = 0 To .DataCnt - 1
      If Not .OutOfScope(i) Then
         If .Visible(i) And Len(.Text(i)) > 0 Then
            Res = GetTextExtentPoint32(hDC, .Text(i), Len(.Text(i)), szText)
            Res = TextOut(hDC, .y(i) - CLng(TextHeight / 12), .x(i) + CLng(szText.cx / 2), .Text(i), Len(.Text(i)))
         End If
      End If
   Next i
End With
Res = SelectObject(hDC, lOldBrush)
Res = SelectObject(hDC, lOldFont)
If lNewFont <> 0 Then DeleteObject (lNewFont)
End Sub



Public Sub OlyDrawTextV_MM(ByVal OlyInd As Long, ByVal hDC As Long, ByVal TextHeight As Double, _
                           ByVal TextWidth As Double, ByVal TextOffset As Double)
Dim i As Long, Res As Long
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
   For i = 0 To .DataCnt - 1
      If Not .OutOfScope(i) Then
         If .Visible(i) And Len(.Text(i)) > 0 Then              'print left of actual spot
            Res = GetTextExtentPoint32(hDC, .Text(i), Len(.Text(i)), szText)
            Res = TextOut(hDC, .y(i) - TextOffset, .x(i) + CLng(szText.cx / 2), .Text(i), Len(.Text(i)))
         End If
      End If
   Next i
End With
Res = SelectObject(hDC, lOldBrush)
Res = SelectObject(hDC, lOldFont)
DeleteObject (lNewFont)
End Sub


'patch to work metric drawing for stick without changing pen weight
Public Sub OlyDrawStickNETVert_MM(ByVal OlyInd As Long, ByVal hDC As Long, ByVal HalfWidth As Long)
Dim i As Long
Dim Res As Long
Dim OldPen As Long
Dim OldBrush As Long
On Error Resume Next
OldBrush = SelectObject(hDC, hOlyClrBrush(OlyInd))
OldPen = SelectObject(hDC, hOlyClrPenStick(OlyInd))
Res = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
      If .Visible(i) Then
         If Not .OutOfScope(i) Then
            Rectangle hDC, .y(i) - HalfWidth, .x(i) - .XL(i), .y(i) + HalfWidth, .x(i) + .XL(i)
         End If
      End If
   Next i
End With
Res = SelectObject(hDC, OldPen)
Res = SelectObject(hDC, OldBrush)
End Sub

Public Sub OlyDrawStickNETHorz_MM(ByVal OlyInd As Long, ByVal hDC As Long, ByVal HalfWidth As Long)
Dim i As Long
Dim Res As Long
Dim OldPen As Long
Dim OldBrush As Long
On Error Resume Next
OldBrush = SelectObject(hDC, hOlyClrBrush(OlyInd))
OldPen = SelectObject(hDC, hOlyClrPenStick(OlyInd))
Res = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
      If .Visible(i) Then
         If Not .OutOfScope(i) Then
            Rectangle hDC, .x(i) - .XL(i), .y(i) - HalfWidth, .x(i) + .XL(i), .y(i) + HalfWidth
         End If
      End If
   Next i
End With
Res = SelectObject(hDC, OldPen)
Res = SelectObject(hDC, OldBrush)
End Sub

Public Sub OlyUMCDrawStickHorz_MM(ByVal OlyInd As Long, ByVal hDC As Long, ByVal HalfWidth As Long)
Dim i As Long
Dim Res As Long
Dim OldPen As Long
Dim OldBrush As Long
On Error Resume Next
OldBrush = SelectObject(hDC, hOlyClrBrush(OlyInd))
OldPen = SelectObject(hDC, hOlyClrPenStick(OlyInd))
Res = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
       If Not .OutOfScope(i) Then
          Rectangle hDC, .XL(i), .y(i) - HalfWidth, .XU(i), .y(i) + HalfWidth
       End If
   Next i
End With
Res = SelectObject(hDC, OldPen)
Res = SelectObject(hDC, OldBrush)
End Sub

Public Sub OlyUMCDrawStickVert_MM(ByVal OlyInd As Long, ByVal hDC As Long, ByVal HalfWidth As Long)
Dim i As Long
Dim Res As Long
Dim OldPen As Long
Dim OldBrush As Long
On Error Resume Next
OldBrush = SelectObject(hDC, hOlyClrBrush(OlyInd))
OldPen = SelectObject(hDC, hOlyClrPenStick(OlyInd))
Res = SetROP2(hDC, R2_MASKPEN)
With OlyCoo(OlyInd)
   For i = 0 To .DataCnt - 1
       If Not .OutOfScope(i) Then
          Rectangle hDC, .y(i) - HalfWidth, .XL(i), .y(i) + HalfWidth, .XU(i)
       End If
   Next i
End With
Res = SelectObject(hDC, OldPen)
Res = SelectObject(hDC, OldBrush)
End Sub

