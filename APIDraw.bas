Attribute VB_Name = "Module7"
'API drawing to the device context functions
'last modified 02/12/2003 nt
'-------------------------------------------
Option Explicit

Private Const DEFAULT_MINIMUM_WARNING_INTERVAL_SECONDS As Long = 5

Private hForeColorPen As Long
Private hBackColorBrush As Long
Private hBackColorPen As Long

Private hIsoColorBrush As Long
Private hCSColorBrush As Long
Private hIsoColorPen As Long
Private hCSColorPen As Long

'selection color objects
Private hSelColorBrush As Long
Private hSelColorPen As Long

'differential display and charge state map colors
'colors 51 to 56 are reserved for charge state map colors
'although it can differentiate only six colors different
'charge states can be painted with those colors
'This is for now only reserved for future use; it tracks
'for now only charge states 1,2,3,4,5; everything else is
'painted as charge state 6
Private hDDClrBrushes(-50 To 56) As Long
Private hDDClrPens(-50 To 56) As Long

Public APIDrawingAborted As Boolean
Public APIDrawStartTime As Date
Private mDrawTimeWarningIntervalSeconds As Single
Private mDrawContinuationClickCount As Long

Public Sub DestroyDrawingObjects()
On Error Resume Next
Dim i As Integer
If hBackColorBrush <> 0 Then DeleteObject (hBackColorBrush)
If hBackColorPen <> 0 Then DeleteObject (hBackColorPen)
If hForeColorPen <> 0 Then DeleteObject (hForeColorPen)
If hIsoColorBrush <> 0 Then DeleteObject (hIsoColorBrush)
If hCSColorBrush <> 0 Then DeleteObject (hCSColorBrush)
If hCSColorPen <> 0 Then DeleteObject (hCSColorPen)
If hIsoColorPen <> 0 Then DeleteObject (hIsoColorPen)
If hSelColorPen <> 0 Then DeleteObject (hSelColorPen)
If hSelColorBrush <> 0 Then DeleteObject (hSelColorBrush)
For i = -50 To 56
    If hDDClrBrushes(i) <> 0 Then DeleteObject (hDDClrBrushes(i))
    If hDDClrPens(i) <> 0 Then DeleteObject (hDDClrPens(i))
Next i
End Sub

Public Sub SetBackForeColorObjects()
On Error Resume Next
'background brush and pen
If hBackColorBrush <> 0 Then DeleteObject (hBackColorBrush)
If hBackColorPen <> 0 Then DeleteObject (hBackColorPen)
hBackColorBrush = CreateSolidBrush(glBackColor)
hBackColorPen = CreatePen(PS_SOLID, 1, glBackColor)
'foreground pen
If hForeColorPen <> 0 Then DeleteObject (hForeColorPen)
hForeColorPen = CreatePen(PS_SOLID, 1, glForeColor)
End Sub

Public Sub GelColorsChange(ByVal ChangePreferences As Boolean, ByVal ChangeDRDefinition As Integer)
On Error Resume Next
Dim i As Integer
For i = 1 To UBound(GelBody)
    If Not GelStatus(i).Deleted Then
       If ChangePreferences Then
          GelData(i).Preferences = glPreferences
          If ChangeDRDefinition < 0 Then InitDrawERColors i
       End If
       GelBody(i).picGraph.Refresh
    End If
Next i
End Sub

Private Sub GelMetrics(ByVal nInd As Long, ByVal dcw As Long, ByVal dch As Long)

If GelBody(nInd).csMyCooSys Is Nothing Then Exit Sub

With GelBody(nInd).csMyCooSys
    .SXOffset = dcw * lDfSXPercent
    .SYOffset = dch * lDfSYPercent
    .LXOffset = dcw * lDfLXPercent
    .LYOffset = dch * lDfLYPercent
    
    Select Case .csOriginXY
    Case glOriginBL             'axes intersection coordinates
        .XYX0 = .LXOffset
        .XYY0 = dch - .LYOffset
        .XYXE = dcw - .SXOffset
        .XYYE = .SYOffset
    Case glOriginBR
        .XYX0 = dcw - .LXOffset
        .XYY0 = dch - .LYOffset
        .XYXE = .SXOffset
        .XYYE = .SYOffset
    Case glOriginTL
        .XYX0 = .LXOffset
        .XYY0 = .LYOffset
        .XYXE = dcw - .SXOffset
        .XYYE = dch - .SYOffset
    Case glOriginTR
        .XYX0 = dcw - .LXOffset
        .XYY0 = .LYOffset
        .XYXE = .SXOffset
        .XYYE = dch - .SYOffset
    End Select
    'origin coordinates - see explanation in the documentation
    If (.csOrigin + .csOriginXY) Mod 2 = 0 Then
       .OrX0 = .XYX0
       .OrXE = .XYXE
    Else
       .OrX0 = .XYXE
       .OrXE = .XYX0
    End If
        
    If (.csOrigin < 3 And .csOriginXY < 3) Or (.csOrigin > 2 And .csOriginXY > 2) Then
       .OrY0 = .XYY0
       .OrYE = .XYYE
    Else
       .OrY0 = .XYYE
       .OrYE = .XYY0
    End If
        
    'viewport coordinates
    .VX0 = CLng(.OrX0)
    .VXE = CLng(.OrXE - .OrX0)
    .VY0 = CLng(.OrY0)
    .VYE = CLng(.OrYE - .OrY0)
End With
End Sub

Public Sub GelCooSys(ByVal Ind As Long, ByVal hDC As Long)
Dim Res As Long
Dim ptPoint As POINTAPI
Dim szSize As Size
On Error Resume Next
Res = SetMapMode(hDC, MM_ANISOTROPIC)
If Not GelBody(Ind).csMyCooSys Is Nothing Then
    With GelBody(Ind).csMyCooSys
         'logical window
         Res = SetWindowOrgEx(hDC, LDfX0, LDfY0, ptPoint)
         Res = SetWindowExtEx(hDC, LDfXE, LDfYE, szSize)
         'viewport
         Res = SetViewportOrgEx(hDC, .VX0, .VY0, ptPoint)
         Res = SetViewportExtEx(hDC, .VXE, .VYE, szSize)
    End With
End If
End Sub

Private Sub GelDrawCooSys(ByVal Ind As Long, ByVal hDC As Long)
Dim OldPen As Long
Dim ptPoint As POINTAPI
Dim Res As Long
On Error Resume Next
OldPen = SelectObject(hDC, hForeColorPen)
Res = SetROP2(hDC, R2_MASKPEN)
Res = SetBkMode(hDC, TRANSPARENT)
With GelBody(Ind).csMyCooSys
    'horizontal axis
    If (.csOrigin < 3 And .csOriginXY < 3) Or _
       (.csOrigin > 2 And .csOriginXY > 2) Then
       Res = MoveToEx(hDC, LDfX0, LDfY0, ptPoint)
       Res = LineTo(hDC, LDfX0 + LDfXE, LDfY0)
    Else
       Res = MoveToEx(hDC, LDfX0, LDfY0 + LDfYE, ptPoint)
       Res = LineTo(hDC, LDfX0 + LDfXE, LDfY0 + LDfYE)
    End If
    'vertical axis
    If (.csOrigin + .csOriginXY) Mod 2 = 0 Then
       Res = MoveToEx(hDC, LDfX0, LDfY0, ptPoint)
       Res = LineTo(hDC, LDfX0, LDfY0 + LDfYE)
    Else
       Res = MoveToEx(hDC, LDfX0 + LDfXE, LDfY0, ptPoint)
       Res = LineTo(hDC, LDfX0 + LDfXE, LDfY0 + LDfYE)
    End If
End With
Res = SelectObject(hDC, OldPen)
End Sub

Private Sub GelDrawBackColor(ByVal hDC As Long, ByVal w As Long, ByVal h As Long)
Dim OldBrush As Long
Dim Res As Long
On Error Resume Next

OldBrush = SelectObject(hDC, hBackColorBrush)
Res = PatBlt(hDC, 0, 0, w, h, PATCOPY)
Res = SelectObject(hDC, OldBrush)
End Sub

Private Sub GelDrawLegend(ByVal Ind As Long, ByVal hDC As Long)
'----------------------------------------------------------------
'draws legend and markings on the coordinate axes; this function
'just redirects drawing to the appropriate function
'----------------------------------------------------------------
Select Case GelBody(Ind).fgDisplay
Case glNormalDisplay
     DrawLegendNormal Ind, hDC
Case glDifferentialDisplay
     DrawLegendDifferential Ind, hDC
Case glChargeStateMapDisplay
     DrawLegendChargeStateMap Ind, hDC
End Select
End Sub

Private Sub DrawLegendShape(ByVal hDC As Long, ByVal ThisShape As Integer, ByVal cx As Long, ByVal cy As Long)
Dim j As Integer
Dim vptAPIs As Variant
Dim ptAPIs() As POINTAPI
On Error Resume Next
Select Case ThisShape
Case glShapeEli
    Ellipse hDC, cx - 100, cy - 100, cx + 100, cy + 100
Case glShapeRec
    Rectangle hDC, cx - 100, cy - 100, cx + 100, cy + 100
Case glShapeRRe
    RoundRect hDC, cx - 100, cy - 100, cx + 100, cy + 100, 100, 100
Case glShapeTri
    ReDim ptAPIs(2)
    vptAPIs = GetTrianglePoints(cx, cy, 100, 1)
    For j = 0 To 2
        ptAPIs(j).X = vptAPIs(j, 0)
        ptAPIs(j).Y = vptAPIs(j, 1)
    Next j
    Polygon hDC, ptAPIs(0), 3
Case glShapeSta
    ReDim ptAPIs(7)
    vptAPIs = Get4StarPoints(cx, cy, 100, 1)
    For j = 0 To 7
        ptAPIs(j).X = vptAPIs(j, 0)
        ptAPIs(j).Y = vptAPIs(j, 1)
    Next j
    Polygon hDC, ptAPIs(0), 8
Case glShapeHex
    ReDim ptAPIs(5)
    vptAPIs = GetHexagonPoints(cx, cy, 100, 1)
    For j = 0 To 5
        ptAPIs(j).X = vptAPIs(j, 0)
        ptAPIs(j).Y = vptAPIs(j, 1)
    Next j
    Polygon hDC, ptAPIs(0), 6
End Select
End Sub


Public Sub SetCSIsoColorObjects()
On Error Resume Next
If hCSColorBrush <> 0 Then DeleteObject (hCSColorBrush)
hCSColorBrush = CreateSolidBrush(glCSColor)
If hCSColorPen <> 0 Then DeleteObject (hCSColorPen)
hCSColorPen = CreatePen(PS_SOLID, 1, glCSColor)
If hIsoColorBrush <> 0 Then DeleteObject (hIsoColorBrush)
hIsoColorBrush = CreateSolidBrush(glIsoColor)
If hIsoColorPen <> 0 Then DeleteObject (hIsoColorPen)
hIsoColorPen = CreatePen(PS_SOLID, 1, glIsoColor)
End Sub

Public Sub SetSelColorObjects()
On Error Resume Next
If hSelColorBrush <> 0 Then DeleteObject (hSelColorBrush)
hSelColorBrush = CreateSolidBrush(glSelColor)
If hSelColorPen <> 0 Then DeleteObject (hSelColorPen)
hSelColorPen = CreatePen(PS_SOLID, 1, glSelColor)
End Sub

Public Sub SetDDRColorObjects()
'prepares brushes and pens for Differential Display
Dim i As Integer
On Error Resume Next
For i = -50 To 50
  If hDDClrBrushes(i) <> 0 Then DeleteObject (hDDClrBrushes(i))
  hDDClrBrushes(i) = CreateSolidBrush(aDDColors(i))
  If hDDClrPens(i) <> 0 Then DeleteObject (hDDClrPens(i))
  hDDClrPens(i) = CreatePen(PS_SOLID, 1, aDDColors(i))
Next i
'charge state map colors we need to create just once
If hDDClrBrushes(51) = 0 Then hDDClrBrushes(51) = CreateSolidBrush(vbBlue)
If hDDClrPens(51) = 0 Then hDDClrPens(51) = CreatePen(PS_SOLID, 1, vbBlue)
If hDDClrBrushes(52) = 0 Then hDDClrBrushes(52) = CreateSolidBrush(vbRed)
If hDDClrPens(52) = 0 Then hDDClrPens(52) = CreatePen(PS_SOLID, 1, vbRed)
If hDDClrBrushes(53) = 0 Then hDDClrBrushes(53) = CreateSolidBrush(RGB(0, 155, 0))      ' Green
If hDDClrPens(53) = 0 Then hDDClrPens(53) = CreatePen(PS_SOLID, 1, RGB(0, 155, 0))      ' Green
If hDDClrBrushes(54) = 0 Then hDDClrBrushes(54) = CreateSolidBrush(vbMagenta)
If hDDClrPens(54) = 0 Then hDDClrPens(54) = CreatePen(PS_SOLID, 1, vbMagenta)
If hDDClrBrushes(55) = 0 Then hDDClrBrushes(55) = CreateSolidBrush(RGB(0, 200, 255))    ' Light blue
If hDDClrPens(55) = 0 Then hDDClrPens(55) = CreatePen(PS_SOLID, 1, RGB(0, 200, 255))    ' Light blue
If hDDClrBrushes(56) = 0 Then hDDClrBrushes(56) = CreateSolidBrush(vbYellow)
If hDDClrPens(56) = 0 Then hDDClrPens(56) = CreatePen(PS_SOLID, 1, vbYellow)
End Sub

Private Sub GelDrawDataNormal(ByVal Ind As Long, ByVal hDC As Long)
'draws data in normal view for FN CooSys to the device context
Dim OldBrush As Long
Dim OldPen As Long
Dim Res As Long
On Error Resume Next
OldBrush = SelectObject(hDC, GetStockObject(NULL_BRUSH))
OldPen = SelectObject(hDC, GetStockObject(NULL_PEN))
Select Case GelBody(Ind).fgZOrder
Case glCSOnTop
    DrawIsoData Ind, hDC
    DrawCSData Ind, hDC
Case glIsoOnTop
    DrawCSData Ind, hDC
    DrawIsoData Ind, hDC
End Select
If GelUMCDraw(Ind).Visible Then
   If GelUMCDraw(Ind).Count > 0 Then DrawUniqueMassClasses Ind, hDC
End If
Res = SelectObject(hDC, OldBrush)
Res = SelectObject(hDC, OldPen)
End Sub

Private Sub GelDrawDataDiff(ByVal Ind As Long, ByVal hDC As Long)
'draws data in differential view for FN CooSys to the device context
Dim OldBrush As Long
Dim OldPen As Long
Dim Res As Long
On Error Resume Next
OldBrush = SelectObject(hDC, GetStockObject(NULL_BRUSH))
OldPen = SelectObject(hDC, GetStockObject(NULL_PEN))
Select Case GelBody(Ind).fgZOrder
Case glCSOnTop
    DrawDiffDataIso Ind, hDC
    DrawDiffDataCS Ind, hDC
Case glIsoOnTop
    DrawDiffDataCS Ind, hDC
    DrawDiffDataIso Ind, hDC
End Select
If GelUMCDraw(Ind).Visible Then
   If GelUMCDraw(Ind).Count > 0 Then DrawUniqueMassClasses Ind, hDC
End If
Res = SelectObject(hDC, OldBrush)
Res = SelectObject(hDC, OldPen)
End Sub

Private Function CheckSlowDrawing() As Boolean
    ' Checks for slow plot drawing
    ' Returns True if OK to keep drawing
    ' Returns False if user requested to stop drawing
    
    Dim eResponse As VbMsgBoxResult
    
    APIDrawingAborted = False
    If DateDiff("s", APIDrawStartTime, Now()) >= mDrawTimeWarningIntervalSeconds Then
        If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            eResponse = vbYes
        Else
            eResponse = MsgBox("Drawing the image appears to be slow; continue drawing?", vbQuestion + vbYesNoCancel + vbDefaultButton1, "Continue")
        End If
        
        If eResponse <> vbYes Then
            APIDrawingAborted = True
            APIDrawStartTime = Now()
            
            CheckSlowDrawing = False
            Exit Function
        End If
        
        mDrawContinuationClickCount = mDrawContinuationClickCount + 1
        
        mDrawTimeWarningIntervalSeconds = mDrawTimeWarningIntervalSeconds * 1.75
        If mDrawTimeWarningIntervalSeconds < 1 Then mDrawTimeWarningIntervalSeconds = 1
        
        APIDrawStartTime = Now()
    End If
    
    CheckSlowDrawing = True
    
End Function

Private Sub DrawIsoData(ByVal Ind As Long, ByVal hDC As Long)
Dim Res As Long
On Error Resume Next
With GelData(Ind)
  If .IsoLines > 0 And GelDraw(Ind).IsoVisible Then
     Res = SelectObject(hDC, hIsoColorBrush)
     If .Preferences.BorderClrSameAsInt Then
        Res = SelectObject(hDC, hIsoColorPen)
     Else
        Res = SelectObject(hDC, hForeColorPen)
     End If
     Select Case glIsoShape
     Case glShapeEli
          DrawEliIso Ind, hDC, .Preferences.AbuAspectRatio
     Case glShapeRec
          DrawRecIso Ind, hDC, .Preferences.AbuAspectRatio
     Case glShapeRRe
          DrawRReIso Ind, hDC, .Preferences.AbuAspectRatio
     Case glShapeTri
          DrawTriIso Ind, hDC, .Preferences.AbuAspectRatio
     Case glShapeSta
          DrawStaIso Ind, hDC, .Preferences.AbuAspectRatio
     Case glShapeHex
          DrawHexIso Ind, hDC, .Preferences.AbuAspectRatio
     End Select
  End If
End With
End Sub

Private Sub DrawCSData(ByVal Ind As Long, ByVal hDC As Long)
Dim Res As Long
On Error Resume Next
With GelData(Ind)
  If .CSLines > 0 And GelDraw(Ind).CSVisible Then
     Res = SelectObject(hDC, hCSColorBrush)
     If .Preferences.BorderClrSameAsInt Then
        Res = SelectObject(hDC, hCSColorPen)
     Else
        Res = SelectObject(hDC, hForeColorPen)
     End If
     Select Case glCSShape
     Case glShapeEli
          DrawEliCS Ind, hDC, .Preferences.AbuAspectRatio
     Case glShapeRec
          DrawRecCS Ind, hDC, .Preferences.AbuAspectRatio
     Case glShapeRRe
          DrawRReCS Ind, hDC, .Preferences.AbuAspectRatio
     Case glShapeTri
          DrawTriCS Ind, hDC, .Preferences.AbuAspectRatio
     Case glShapeSta
          DrawStaCS Ind, hDC, .Preferences.AbuAspectRatio
     Case glShapeHex
          DrawHexCS Ind, hDC, .Preferences.AbuAspectRatio
     End Select
  End If
End With
End Sub


Public Function ClippingRegionA(hDC As Long) As Long
Dim ptOrg As POINTAPI
Dim szExt As Size
Dim Res As Long
Dim hCR As Long
On Error Resume Next
Res = GetViewportOrgEx(hDC, ptOrg)
Res = GetViewportExtEx(hDC, szExt)
hCR = CreateRectRgn(ptOrg.X, ptOrg.Y, ptOrg.X + szExt.cx, ptOrg.Y + szExt.cy)
Res = SelectClipRgn(hDC, hCR)
ClippingRegionA = hCR
End Function

Private Sub DrawRecCS(ByVal Ind As Long, ByVal hDC As Long, ByVal ar As Double)
'this is called only if .CSCount>0
Dim i As Long
Dim h As Long, v As Long
On Error Resume Next
ResetDrawTime
With GelDraw(Ind)
   For i = 1 To .CSCount
      If .CSID(i) > 0 And .CSR(i) > 0 Then
        h = .CSR(i) \ 2
        v = CLng(.CSR(i) / (2 * ar))
        Rectangle hDC, .CSX(i) - h, .CSY(i) - v, .CSX(i) + h, .CSY(i) + v
      End If
      
      If i Mod 10000 = 0 Then
          If Not CheckSlowDrawing() Then Exit Sub
      End If
   Next i
End With
End Sub

Private Sub DrawRReCS(ByVal Ind As Long, ByVal hDC As Long, ByVal ar As Double)
Dim i As Long
Dim h As Long, v As Long
On Error Resume Next
ResetDrawTime
With GelDraw(Ind)
   For i = 1 To .CSCount
      If .CSID(i) > 0 And .CSR(i) > 0 Then
        h = .CSR(i) \ 2
        v = CLng(.CSR(i) / (2 * ar))
        RoundRect hDC, .CSX(i) - h, .CSY(i) - v, .CSX(i) + h, .CSY(i) + v, h, v
      End If
      
      If i Mod 10000 = 0 Then
          If Not CheckSlowDrawing() Then Exit Sub
      End If
   Next i
End With
End Sub

Private Sub DrawRecIso(ByVal Ind As Long, ByVal hDC As Long, ByVal ar As Double)
Dim i As Long
Dim h As Long, v As Long
On Error Resume Next
ResetDrawTime
With GelDraw(Ind)
   For i = 1 To .IsoCount
      If .IsoID(i) > 0 And .IsoR(i) > 0 Then
        h = .IsoR(i) \ 2
        v = CLng(.IsoR(i) / (2 * ar))
        Rectangle hDC, .IsoX(i) - h, .IsoY(i) - v, .IsoX(i) + h, .IsoY(i) + v
      End If
      
      If i Mod 10000 = 0 Then
          If Not CheckSlowDrawing() Then Exit Sub
      End If
   Next i
End With
End Sub

Private Sub DrawRReIso(ByVal Ind As Long, ByVal hDC As Long, ByVal ar As Double)
Dim i As Long
Dim h As Long, v As Long
On Error Resume Next
ResetDrawTime
With GelDraw(Ind)
   For i = 1 To .IsoCount
      If .IsoID(i) > 0 And .IsoR(i) > 0 Then
        h = .IsoR(i) \ 2
        v = CLng(.IsoR(i) / (2 * ar))
        RoundRect hDC, .IsoX(i) - h, .IsoY(i) - v, .IsoX(i) + h, .IsoY(i) + v, h, v
      End If
      
      If i Mod 10000 = 0 Then
          If Not CheckSlowDrawing() Then Exit Sub
      End If
   Next i
End With
End Sub

Private Sub DrawEliCS(ByVal Ind As Long, ByVal hDC As Long, ByVal ar As Double)
Dim i As Long
Dim h As Long, v As Long
On Error Resume Next
ResetDrawTime
With GelDraw(Ind)
   For i = 1 To .CSCount
      If .CSID(i) > 0 And .CSR(i) > 0 Then
        h = .CSR(i) \ 2
        v = CLng(.CSR(i) / (2 * ar))
        Ellipse hDC, .CSX(i) - h, .CSY(i) - v, .CSX(i) + h, .CSY(i) + v
      End If
      
      If i Mod 10000 = 0 Then
          If Not CheckSlowDrawing() Then Exit Sub
      End If
   Next i
End With
End Sub

Private Sub DrawEliIso(ByVal Ind As Long, ByVal hDC As Long, ByVal ar As Double)
Dim i As Long
Dim h As Long, v As Long
On Error Resume Next
ResetDrawTime
With GelDraw(Ind)
   For i = 1 To .IsoCount
      If .IsoID(i) > 0 And .IsoR(i) > 0 Then
        h = .IsoR(i) \ 2
        v = CLng(.IsoR(i) / (2 * ar))
        Ellipse hDC, .IsoX(i) - h, .IsoY(i) - v, .IsoX(i) + h, .IsoY(i) + v
      End If
      
      If i Mod 10000 = 0 Then
          If Not CheckSlowDrawing() Then Exit Sub
      End If
   Next i
End With
End Sub

Private Sub DrawTriCS(ByVal Ind As Long, ByVal hDC As Long, ByVal ar As Double)
'this 0.87 is approx. SQRT(3)\2
Dim vptAPIs As Variant
Dim ptAPIs(2) As POINTAPI
Dim i As Long, j As Long
On Error Resume Next
ResetDrawTime
With GelDraw(Ind)
   For i = 1 To .CSCount
       If .CSID(i) > 0 And .CSR(i) > 0 Then
         vptAPIs = GetTrianglePoints(.CSX(i), .CSY(i), .CSR(i) \ 2, ar)
         For j = 0 To 2
             ptAPIs(j).X = vptAPIs(j, 0)
             ptAPIs(j).Y = vptAPIs(j, 1)
         Next j
         Polygon hDC, ptAPIs(0), 3
       End If
      
      If i Mod 10000 = 0 Then
          If Not CheckSlowDrawing() Then Exit Sub
      End If
   Next i
End With
End Sub

Private Sub DrawTriIso(ByVal Ind As Long, ByVal hDC As Long, ByVal ar As Double)
Dim vptAPIs As Variant
Dim ptAPIs(2) As POINTAPI
Dim i As Long, j As Long
On Error Resume Next
ResetDrawTime
With GelDraw(Ind)
   For i = 1 To .IsoCount
       If .IsoID(i) > 0 And .IsoR(i) > 0 Then
         vptAPIs = GetTrianglePoints(.IsoX(i), .IsoY(i), .IsoR(i) \ 2, ar)
         For j = 0 To 2
             ptAPIs(j).X = vptAPIs(j, 0)
             ptAPIs(j).Y = vptAPIs(j, 1)
         Next j
         Polygon hDC, ptAPIs(0), 3
       End If
   
      If i Mod 10000 = 0 Then
          If Not CheckSlowDrawing() Then Exit Sub
      End If
   Next i
End With
End Sub

Private Sub DrawStaCS(ByVal Ind As Long, ByVal hDC As Long, ByVal ar As Double)
Dim vptAPIs As Variant
Dim ptAPIs(7) As POINTAPI
Dim i As Long, j As Long
On Error Resume Next
ResetDrawTime
With GelDraw(Ind)
   For i = 1 To .CSCount
       If .CSID(i) > 0 And .CSR(i) > 0 Then
         vptAPIs = Get4StarPoints(.CSX(i), .CSY(i), .CSR(i) \ 2, ar)
         For j = 0 To 7
             ptAPIs(j).X = vptAPIs(j, 0)
             ptAPIs(j).Y = vptAPIs(j, 1)
         Next j
         Polygon hDC, ptAPIs(0), 8
       End If
   
      If i Mod 10000 = 0 Then
          If Not CheckSlowDrawing() Then Exit Sub
      End If
   Next i
End With
End Sub

Private Sub DrawStaIso(ByVal Ind As Long, ByVal hDC As Long, ByVal ar As Double)
Dim vptAPIs As Variant
Dim ptAPIs(7) As POINTAPI
Dim i As Long, j As Long
On Error Resume Next
ResetDrawTime
With GelDraw(Ind)
   For i = 1 To .IsoCount
       If .IsoID(i) > 0 And .IsoR(i) > 0 Then
         vptAPIs = Get4StarPoints(.IsoX(i), .IsoY(i), .IsoR(i) \ 2, ar)
         For j = 0 To 7
             ptAPIs(j).X = vptAPIs(j, 0)
             ptAPIs(j).Y = vptAPIs(j, 1)
         Next j
         Polygon hDC, ptAPIs(0), 8
       End If
   
      If i Mod 10000 = 0 Then
          If Not CheckSlowDrawing() Then Exit Sub
      End If
   Next i
End With
End Sub

Private Sub DrawHexIso(ByVal Ind As Long, ByVal hDC As Long, ByVal ar As Double)
Dim vptAPIs As Variant
Dim ptAPIs(5) As POINTAPI
Dim i As Long, j As Long
On Error Resume Next
ResetDrawTime
With GelDraw(Ind)
   For i = 1 To .IsoCount
       If .IsoID(i) > 0 And .IsoR(i) > 0 Then
         vptAPIs = GetHexagonPoints(.IsoX(i), .IsoY(i), .IsoR(i) \ 2, ar)
         For j = 0 To 5
             ptAPIs(j).X = vptAPIs(j, 0)
             ptAPIs(j).Y = vptAPIs(j, 1)
         Next j
         Polygon hDC, ptAPIs(0), 6
       End If
   
      If i Mod 10000 = 0 Then
          If Not CheckSlowDrawing() Then Exit Sub
      End If
   Next i
End With
End Sub

Private Sub DrawHexCS(ByVal Ind As Long, ByVal hDC As Long, ByVal ar As Double)
Dim vptAPIs As Variant
Dim ptAPIs(5) As POINTAPI
Dim i As Long, j As Long
On Error Resume Next
ResetDrawTime
With GelDraw(Ind)
   For i = 1 To .CSCount
       If .CSID(i) > 0 And .CSR(i) > 0 Then
         vptAPIs = GetHexagonPoints(.CSX(i), .CSY(i), .CSR(i) \ 2, ar)
         For j = 0 To 5
             ptAPIs(j).X = vptAPIs(j, 0)
             ptAPIs(j).Y = vptAPIs(j, 1)
         Next j
         Polygon hDC, ptAPIs(0), 6
       End If
   
      If i Mod 10000 = 0 Then
          If Not CheckSlowDrawing() Then Exit Sub
      End If
   Next i
End With
End Sub

Private Sub DrawDiffDataCS(ByVal Ind As Long, ByVal hDC As Long)
Dim vptAPIs As Variant
Dim ptAPIs() As POINTAPI
Dim j As Integer
Dim Res As Long, i As Long
Dim h As Long, v As Long
Dim ar As Double
On Error Resume Next
ResetDrawTime
ar = GelData(Ind).Preferences.AbuAspectRatio
With GelDraw(Ind)
   If .CSCount > 0 Then
      If GelData(Ind).Preferences.BorderClrSameAsInt Then
      Select Case glCSShape
      Case glShapeEli
        For i = 1 To .CSCount
            If .CSERClr(i) <> glDONT_DISPLAY And .CSID(i) > 0 And .CSR(i) > 0 Then
               Res = SelectObject(hDC, hDDClrBrushes(.CSERClr(i)))
               Res = SelectObject(hDC, hDDClrPens(.CSERClr(i)))
               h = .CSR(i) \ 2
               v = CLng(.CSR(i) / (2 * ar))
               Ellipse hDC, .CSX(i) - h, .CSY(i) - v, .CSX(i) + h, .CSY(i) + v
            End If
        
            If i Mod 10000 = 0 Then
                If Not CheckSlowDrawing() Then Exit Sub
            End If
        Next i
      Case glShapeRec
        For i = 1 To .CSCount
            If .CSERClr(i) <> glDONT_DISPLAY And .CSID(i) > 0 And .CSR(i) > 0 Then
               Res = SelectObject(hDC, hDDClrBrushes(.CSERClr(i)))
               Res = SelectObject(hDC, hDDClrPens(.CSERClr(i)))
               h = .CSR(i) \ 2
               v = CLng(.CSR(i) / (2 * ar))
               Rectangle hDC, .CSX(i) - h, .CSY(i) - v, .CSX(i) + h, .CSY(i) + v
            End If
        
            If i Mod 10000 = 0 Then
                If Not CheckSlowDrawing() Then Exit Sub
            End If
        Next i
      Case glShapeRRe
        For i = 1 To .CSCount
            If .CSERClr(i) <> glDONT_DISPLAY And .CSID(i) > 0 And .CSR(i) > 0 Then
               Res = SelectObject(hDC, hDDClrBrushes(.CSERClr(i)))
               Res = SelectObject(hDC, hDDClrPens(.CSERClr(i)))
               h = .CSR(i) \ 2
               v = CLng(.CSR(i) / (2 * ar))
               RoundRect hDC, .CSX(i) - h, .CSY(i) - v, .CSX(i) + h, .CSY(i) + v, h, v
            End If
        
            If i Mod 10000 = 0 Then
                If Not CheckSlowDrawing() Then Exit Sub
            End If
        Next i
      Case glShapeTri
        ReDim ptAPIs(2)
        For i = 1 To .CSCount
            If .CSERClr(i) <> glDONT_DISPLAY And .CSID(i) > 0 And .CSR(i) > 0 Then
               Res = SelectObject(hDC, hDDClrBrushes(.CSERClr(i)))
               Res = SelectObject(hDC, hDDClrPens(.CSERClr(i)))
               vptAPIs = GetTrianglePoints(.CSX(i), .CSY(i), .CSR(i) \ 2, ar)
               For j = 0 To 2
                 ptAPIs(j).X = vptAPIs(j, 0)
                 ptAPIs(j).Y = vptAPIs(j, 1)
               Next j
               Polygon hDC, ptAPIs(0), 3
            End If
        
            If i Mod 10000 = 0 Then
                If Not CheckSlowDrawing() Then Exit Sub
            End If
        Next i
      Case glShapeSta
        ReDim ptAPIs(7)
        For i = 1 To .CSCount
            If .CSERClr(i) <> glDONT_DISPLAY And .CSID(i) > 0 And .CSR(i) > 0 Then
               Res = SelectObject(hDC, hDDClrBrushes(.CSERClr(i)))
               Res = SelectObject(hDC, hDDClrPens(.CSERClr(i)))
               vptAPIs = Get4StarPoints(.CSX(i), .CSY(i), .CSR(i) \ 2, ar)
               For j = 0 To 7
                 ptAPIs(j).X = vptAPIs(j, 0)
                 ptAPIs(j).Y = vptAPIs(j, 1)
               Next j
               Polygon hDC, ptAPIs(0), 8
            End If
        
            If i Mod 10000 = 0 Then
                If Not CheckSlowDrawing() Then Exit Sub
            End If
        Next i
      Case glShapeHex
        ReDim ptAPIs(5)
        For i = 1 To .CSCount
            If .CSERClr(i) <> glDONT_DISPLAY And .CSID(i) > 0 And .CSR(i) > 0 Then
               Res = SelectObject(hDC, hDDClrBrushes(.CSERClr(i)))
               Res = SelectObject(hDC, hDDClrPens(.CSERClr(i)))
               vptAPIs = GetHexagonPoints(.CSX(i), .CSY(i), .CSR(i) \ 2, ar)
               For j = 0 To 5
                 ptAPIs(j).X = vptAPIs(j, 0)
                 ptAPIs(j).Y = vptAPIs(j, 1)
               Next j
               Polygon hDC, ptAPIs(0), 6
            End If
        
            If i Mod 10000 = 0 Then
                If Not CheckSlowDrawing() Then Exit Sub
            End If
        Next i
      End Select
   Else
      Res = SelectObject(hDC, hForeColorPen)
      Select Case glCSShape
      Case glShapeEli
        For i = 1 To .CSCount
            If .CSERClr(i) <> glDONT_DISPLAY And .CSID(i) > 0 And .CSR(i) > 0 Then
               Res = SelectObject(hDC, hDDClrBrushes(.CSERClr(i)))
               h = .CSR(i) \ 2
               v = CLng(.CSR(i) / (2 * ar))
               Ellipse hDC, .CSX(i) - h, .CSY(i) - v, .CSX(i) + h, .CSY(i) + v
            End If
        
            If i Mod 10000 = 0 Then
                If Not CheckSlowDrawing() Then Exit Sub
            End If
        Next i
      Case glShapeRec
        For i = 1 To .CSCount
            If .CSERClr(i) <> glDONT_DISPLAY And .CSID(i) > 0 And .CSR(i) > 0 Then
               Res = SelectObject(hDC, hDDClrBrushes(.CSERClr(i)))
               h = .CSR(i) \ 2
               v = CLng(.CSR(i) / (2 * ar))
               Rectangle hDC, .CSX(i) - h, .CSY(i) - v, .CSX(i) + h, .CSY(i) + v
            End If
        
            If i Mod 10000 = 0 Then
                If Not CheckSlowDrawing() Then Exit Sub
            End If
        Next i
      Case glShapeRRe
        For i = 1 To .CSCount
            If .CSERClr(i) <> glDONT_DISPLAY And .CSID(i) > 0 And .CSR(i) > 0 Then
               Res = SelectObject(hDC, hDDClrBrushes(.CSERClr(i)))
               h = .CSR(i) \ 2
               v = CLng(.CSR(i) / (2 * ar))
               RoundRect hDC, .CSX(i) - h, .CSY(i) - v, .CSX(i) + h, .CSY(i) + v, h, v
            End If
        
            If i Mod 10000 = 0 Then
                If Not CheckSlowDrawing() Then Exit Sub
            End If
        Next i
      Case glShapeTri
        ReDim ptAPIs(2)
        For i = 1 To .CSCount
            If .CSERClr(i) <> glDONT_DISPLAY And .CSID(i) > 0 And .CSR(i) > 0 Then
               Res = SelectObject(hDC, hDDClrBrushes(.CSERClr(i)))
               vptAPIs = GetTrianglePoints(.CSX(i), .CSY(i), .CSR(i) \ 2, ar)
               For j = 0 To 2
                 ptAPIs(j).X = vptAPIs(j, 0)
                 ptAPIs(j).Y = vptAPIs(j, 1)
               Next j
               Polygon hDC, ptAPIs(0), 3
            End If
        
            If i Mod 10000 = 0 Then
                If Not CheckSlowDrawing() Then Exit Sub
            End If
        Next i
      Case glShapeSta
        ReDim ptAPIs(7)
        For i = 1 To .CSCount
            If .CSERClr(i) <> glDONT_DISPLAY And .CSID(i) > 0 And .CSR(i) > 0 Then
               Res = SelectObject(hDC, hDDClrBrushes(.CSERClr(i)))
               vptAPIs = Get4StarPoints(.CSX(i), .CSY(i), .CSR(i) \ 2, ar)
               For j = 0 To 7
                 ptAPIs(j).X = vptAPIs(j, 0)
                 ptAPIs(j).Y = vptAPIs(j, 1)
               Next j
               Polygon hDC, ptAPIs(0), 8
            End If
        
            If i Mod 10000 = 0 Then
                If Not CheckSlowDrawing() Then Exit Sub
            End If
        Next i
      Case glShapeHex
        ReDim ptAPIs(5)
        For i = 1 To .CSCount
            If .CSERClr(i) <> glDONT_DISPLAY And .CSID(i) > 0 And .CSR(i) > 0 Then
               Res = SelectObject(hDC, hDDClrBrushes(.CSERClr(i)))
               vptAPIs = GetHexagonPoints(.CSX(i), .CSY(i), .CSR(i) \ 2, ar)
               For j = 0 To 5
                 ptAPIs(j).X = vptAPIs(j, 0)
                 ptAPIs(j).Y = vptAPIs(j, 1)
               Next j
               Polygon hDC, ptAPIs(0), 6
            End If
        
            If i Mod 10000 = 0 Then
                If Not CheckSlowDrawing() Then Exit Sub
            End If
        Next i
      End Select
      End If
   End If
End With
End Sub

Private Sub DrawDiffDataIso(ByVal Ind As Long, ByVal hDC As Long)
Dim vptAPIs As Variant
Dim ptAPIs() As POINTAPI
Dim j As Integer
Dim Res As Long, i As Long
Dim h As Long, v As Long
Dim ar As Double
On Error Resume Next
ResetDrawTime
ar = GelData(Ind).Preferences.AbuAspectRatio
With GelDraw(Ind)
   If .IsoCount > 0 Then
      If GelData(Ind).Preferences.BorderClrSameAsInt Then
      Select Case glIsoShape
      Case glShapeEli
        For i = 1 To .IsoCount
            If .IsoERClr(i) <> glDONT_DISPLAY And .IsoID(i) > 0 And .IsoR(i) > 0 Then
               Res = SelectObject(hDC, hDDClrBrushes(.IsoERClr(i)))
               Res = SelectObject(hDC, hDDClrPens(.IsoERClr(i)))
               h = .IsoR(i) \ 2
               v = CLng(.IsoR(i) / (2 * ar))
               Ellipse hDC, .IsoX(i) - h, .IsoY(i) - v, .IsoX(i) + h, .IsoY(i) + v
            End If
        
            If i Mod 10000 = 0 Then
                If Not CheckSlowDrawing() Then Exit Sub
            End If
        Next i
      Case glShapeRec
        For i = 1 To .IsoCount
            If .IsoERClr(i) <> glDONT_DISPLAY And .IsoID(i) > 0 And .IsoR(i) > 0 Then
               Res = SelectObject(hDC, hDDClrBrushes(.IsoERClr(i)))
               Res = SelectObject(hDC, hDDClrPens(.IsoERClr(i)))
               h = .IsoR(i) \ 2
               v = CLng(.IsoR(i) / (2 * ar))
               Rectangle hDC, .IsoX(i) - h, .IsoY(i) - v, .IsoX(i) + h, .IsoY(i) + v
            End If
        
            If i Mod 10000 = 0 Then
                If Not CheckSlowDrawing() Then Exit Sub
            End If
        Next i
      Case glShapeRRe
        For i = 1 To .IsoCount
            If .IsoERClr(i) <> glDONT_DISPLAY And .IsoID(i) > 0 And .IsoR(i) > 0 Then
               Res = SelectObject(hDC, hDDClrBrushes(.IsoERClr(i)))
               Res = SelectObject(hDC, hDDClrPens(.IsoERClr(i)))
               h = .IsoR(i) \ 2
               v = CLng(.IsoR(i) / (2 * ar))
               RoundRect hDC, .IsoX(i) - h, .IsoY(i) - v, .IsoX(i) + h, .IsoY(i) + v, h, v
            End If
        
            If i Mod 10000 = 0 Then
                If Not CheckSlowDrawing() Then Exit Sub
            End If
        Next i
      Case glShapeTri
        ReDim ptAPIs(2)
        For i = 1 To .IsoCount
            If .IsoERClr(i) <> glDONT_DISPLAY And .IsoID(i) > 0 And .IsoR(i) > 0 Then
               Res = SelectObject(hDC, hDDClrBrushes(.IsoERClr(i)))
               Res = SelectObject(hDC, hDDClrPens(.IsoERClr(i)))
               vptAPIs = GetTrianglePoints(.IsoX(i), .IsoY(i), .IsoR(i) \ 2, ar)
               For j = 0 To 2
                 ptAPIs(j).X = vptAPIs(j, 0)
                 ptAPIs(j).Y = vptAPIs(j, 1)
               Next j
               Polygon hDC, ptAPIs(0), 3
            End If
        
            If i Mod 10000 = 0 Then
                If Not CheckSlowDrawing() Then Exit Sub
            End If
        Next i
      Case glShapeSta
        ReDim ptAPIs(7)
        For i = 1 To .IsoCount
            If .IsoERClr(i) <> glDONT_DISPLAY And .IsoID(i) > 0 And .IsoR(i) > 0 Then
               Res = SelectObject(hDC, hDDClrBrushes(.IsoERClr(i)))
               Res = SelectObject(hDC, hDDClrPens(.IsoERClr(i)))
               vptAPIs = Get4StarPoints(.IsoX(i), .IsoY(i), .IsoR(i) \ 2, ar)
               For j = 0 To 7
                 ptAPIs(j).X = vptAPIs(j, 0)
                 ptAPIs(j).Y = vptAPIs(j, 1)
               Next j
               Polygon hDC, ptAPIs(0), 8
            End If
        
            If i Mod 10000 = 0 Then
                If Not CheckSlowDrawing() Then Exit Sub
            End If
        Next i
      Case glShapeHex
        ReDim ptAPIs(5)
        For i = 1 To .IsoCount
            If .IsoERClr(i) <> glDONT_DISPLAY And .IsoID(i) > 0 And .IsoR(i) > 0 Then
               Res = SelectObject(hDC, hDDClrBrushes(.IsoERClr(i)))
               Res = SelectObject(hDC, hDDClrPens(.IsoERClr(i)))
               vptAPIs = GetHexagonPoints(.IsoX(i), .IsoY(i), .IsoR(i) \ 2, ar)
               For j = 0 To 5
                 ptAPIs(j).X = vptAPIs(j, 0)
                 ptAPIs(j).Y = vptAPIs(j, 1)
               Next j
               Polygon hDC, ptAPIs(0), 6
            End If
        
            If i Mod 10000 = 0 Then
                If Not CheckSlowDrawing() Then Exit Sub
            End If
        Next i
      End Select
   Else
      Res = SelectObject(hDC, hForeColorPen)
      Select Case glIsoShape
      Case glShapeEli
        For i = 1 To .IsoCount
            If .IsoERClr(i) <> glDONT_DISPLAY And .IsoID(i) > 0 And .IsoR(i) > 0 Then
               Res = SelectObject(hDC, hDDClrBrushes(.IsoERClr(i)))
               h = .IsoR(i) \ 2
               v = CLng(.IsoR(i) / (2 * ar))
               Ellipse hDC, .IsoX(i) - h, .IsoY(i) - v, .IsoX(i) + h, .IsoY(i) + v
            End If
        
            If i Mod 10000 = 0 Then
                If Not CheckSlowDrawing() Then Exit Sub
            End If
        Next i
      Case glShapeRec
        For i = 1 To .IsoCount
            If .IsoERClr(i) <> glDONT_DISPLAY And .IsoID(i) > 0 And .IsoR(i) > 0 Then
               Res = SelectObject(hDC, hDDClrBrushes(.IsoERClr(i)))
               h = .IsoR(i) \ 2
               v = CLng(.IsoR(i) / (2 * ar))
               Rectangle hDC, .IsoX(i) - h, .IsoY(i) - v, .IsoX(i) + h, .IsoY(i) + v
            End If
        
            If i Mod 10000 = 0 Then
                If Not CheckSlowDrawing() Then Exit Sub
            End If
        Next i
      Case glShapeRRe
        For i = 1 To .IsoCount
            If .IsoERClr(i) <> glDONT_DISPLAY And .IsoID(i) > 0 And .IsoR(i) > 0 Then
               Res = SelectObject(hDC, hDDClrBrushes(.IsoERClr(i)))
               h = .IsoR(i) \ 2
               v = CLng(.IsoR(i) / (2 * ar))
               RoundRect hDC, .IsoX(i) - h, .IsoY(i) - v, .IsoX(i) + h, .IsoY(i) + v, h, v
            End If
        
            If i Mod 10000 = 0 Then
                If Not CheckSlowDrawing() Then Exit Sub
            End If
        Next i
      Case glShapeTri
        ReDim ptAPIs(2)
        For i = 1 To .IsoCount
            If .IsoERClr(i) <> glDONT_DISPLAY And .IsoID(i) > 0 And .IsoR(i) > 0 Then
               Res = SelectObject(hDC, hDDClrBrushes(.IsoERClr(i)))
               vptAPIs = GetTrianglePoints(.IsoX(i), .IsoY(i), .IsoR(i) \ 2, ar)
               For j = 0 To 2
                 ptAPIs(j).X = vptAPIs(j, 0)
                 ptAPIs(j).Y = vptAPIs(j, 1)
               Next j
               Polygon hDC, ptAPIs(0), 3
            End If
        
            If i Mod 10000 = 0 Then
                If Not CheckSlowDrawing() Then Exit Sub
            End If
        Next i
      Case glShapeSta
        ReDim ptAPIs(7)
        For i = 1 To .IsoCount
            If .IsoERClr(i) <> glDONT_DISPLAY And .IsoID(i) > 0 And .IsoR(i) > 0 Then
               Res = SelectObject(hDC, hDDClrBrushes(.IsoERClr(i)))
               vptAPIs = Get4StarPoints(.IsoX(i), .IsoY(i), .IsoR(i) \ 2, ar)
               For j = 0 To 7
                 ptAPIs(j).X = vptAPIs(j, 0)
                 ptAPIs(j).Y = vptAPIs(j, 1)
               Next j
               Polygon hDC, ptAPIs(0), 8
            End If
        
            If i Mod 10000 = 0 Then
                If Not CheckSlowDrawing() Then Exit Sub
            End If
        Next i
      Case glShapeHex
        ReDim ptAPIs(5)
        For i = 1 To .IsoCount
            If .IsoERClr(i) <> glDONT_DISPLAY And .IsoID(i) > 0 And .IsoR(i) > 0 Then
               Res = SelectObject(hDC, hDDClrBrushes(.IsoERClr(i)))
               vptAPIs = GetHexagonPoints(.IsoX(i), .IsoY(i), .IsoR(i) \ 2, ar)
               For j = 0 To 5
                 ptAPIs(j).X = vptAPIs(j, 0)
                 ptAPIs(j).Y = vptAPIs(j, 1)
               Next j
               Polygon hDC, ptAPIs(0), 6
            End If
        
            If i Mod 10000 = 0 Then
                If Not CheckSlowDrawing() Then Exit Sub
            End If
        Next i
      End Select
      End If
   End If
End With
End Sub

Private Sub ResetDrawTime()
    APIDrawStartTime = Now()
    APIDrawingAborted = False
    mDrawTimeWarningIntervalSeconds = DEFAULT_MINIMUM_WARNING_INTERVAL_SECONDS + mDrawContinuationClickCount / 2#
End Sub

Private Sub GelDrawVAxisNumbers(ByVal Ind As Long, ByVal hDC As Long)

Const TICK_MARK_COUNT = 11

Dim i As Integer, iSign As Integer
Dim sNumber As String
Dim szNumber As Size
Dim rStep As Double
Dim lStep As Long, ly As Long
Dim NumL As Long
Dim MarkO As Long   'X coordinates of the thick mark (out)
Dim MarkL As Long   'X coordinates of the thick mark (on axis)
Dim ptPoint As POINTAPI
Dim vSelVec As Variant
Dim Res As Long
Dim strFormat As String

On Error Resume Next

With GelBody(Ind).csMyCooSys
    If (.csOrigin + .csOriginXY) Mod 2 = 0 Then
       MarkO = LDfX0 - 30
       MarkL = LDfX0
    Else
       MarkO = LDfXE + 30
       MarkL = LDfXE
    End If
    If .csOrigin > 2 Then
       iSign = -1
    Else
       iSign = 1
    End If
    rStep = (.CurrRYMax - .CurrRYMin) / (TICK_MARK_COUNT - 1)
    lStep = (LDfY2 - LDfY1) \ (TICK_MARK_COUNT - 1)
    
    If rStep >= 10 Then
        strFormat = "#,###,##0"
    ElseIf rStep >= 1 Then
        strFormat = "#,###,##0.0"
    ElseIf rStep >= 0.1 Then
        strFormat = "#,###,##0.00"
    ElseIf rStep >= 0.01 Then
        strFormat = "#,###,##0.000"
    ElseIf rStep >= 0.001 Then
        strFormat = "#,###,##0.0000"
    Else
        strFormat = "#,###,##0.00000"
    End If
    
    For i = 0 To TICK_MARK_COUNT - 1
        ly = LDfY1 + i * lStep
        Select Case .csYScale
        Case glVAxisLin
            sNumber = Format$(.CurrRYMin + i * rStep, strFormat)
        Case glVAxisLog
            sNumber = Format$(10 ^ (.CurrRYMin + i * rStep), strFormat)
        End Select
        Res = GetTextExtentPoint32(hDC, sNumber, Len(sNumber), szNumber)
        Res = MoveToEx(hDC, MarkO, ly, ptPoint)
        Res = LineTo(hDC, MarkL, ly)
        vSelVec = GetJobVector(Ind, vAxXSelectMatrix, vXVNumJobMatrix)
        NumL = SP(vSelVec, MarkL, 30, LDfSX \ 4, szNumber.cx)
        If bIncludeTextLabels Then
            Res = TextOut(hDC, NumL, ly + iSign * szNumber.cy \ 2, sNumber, Len(sNumber))
        End If
    Next i
End With
End Sub

Private Sub GelDrawpiHAxisNumbers(ByVal Ind As Long, ByVal hDC As Long)
Dim i As Integer, iSign As Integer
Dim sNumber As String
Dim szNumber As Size
Dim lStep As Long, lx As Long
Dim lNumOfSteps As Long
Dim NumT As Long
Dim MarkO As Long   'X coordinates of the thick mark (out)
Dim MarkL As Long   'X coordinates of the thick mark (on axis)
Dim ptPoint As POINTAPI
Dim vSelVec As Variant
Dim sCurrpI As Double
Dim lCurrFN As Long
Dim Res As Long
On Error Resume Next

With GelBody(Ind).csMyCooSys
    If (.csOrigin < 3 And .csOriginXY < 3) Or _
       (.csOrigin > 2 And .csOriginXY > 2) Then
       MarkO = LDfY0 - 30
       MarkL = LDfY0
    Else
       MarkO = LDfYE + 30
       MarkL = LDfYE
    End If
    If .csOrigin Mod 2 = 0 Then
       iSign = 1
    Else
       iSign = -1
    End If
    lNumOfSteps = .NumOfStepsX
    lStep = (LDfX2 - LDfX1) \ lNumOfSteps
    Res = GetTextExtentPoint32(hDC, "0123456789", 10, szNumber)
    vSelVec = GetJobVector(Ind, vAxYSelectMatrix, vYHNumJobMatrix)
    NumT = SP(vSelVec, MarkL, 30, lDfSY \ 8, szNumber.cy)
    For i = 0 To lNumOfSteps
        lx = LDfX1 + i * lStep
        lCurrFN = .CurrRXMin + i * .csXStep
        sCurrpI = GelData(Ind).ScanInfo(GetDFIndex(Ind, CLng(lCurrFN))).ScanPI
        If Err Then
           Err.Clear
        Else
           sNumber = Format$(sCurrpI, "0.0000")
           Res = GetTextExtentPoint32(hDC, sNumber, Len(sNumber), szNumber)
           Res = MoveToEx(hDC, lx, MarkO, ptPoint)
           Res = LineTo(hDC, lx, MarkL)
           If bIncludeTextLabels Then
               Res = TextOut(hDC, lx + iSign * szNumber.cx \ 2, NumT, sNumber, Len(sNumber))
           End If
        End If
    Next i
End With
End Sub

Private Sub GelDrawFNHAxisNumbers(ByVal Ind As Long, ByVal hDC As Long, ByVal blnShowNETValues As Boolean)
Dim i As Integer, iSign As Integer
Dim sNumber As String
Dim szNumber As Size
Dim lNumOfSteps As Long
Dim lStep As Long, lx As Long
Dim lngScanNumber As Long
Dim NumT As Long
Dim MarkO As Long   'X coordinates of the thick mark (out)
Dim MarkL As Long   'X coordinates of the thick mark (on axis)
Dim ptPoint As POINTAPI
Dim vSelVec As Variant
Dim Res As Long

On Error Resume Next

With GelBody(Ind).csMyCooSys
    If (.csOrigin < 3 And .csOriginXY < 3) Or _
       (.csOrigin > 2 And .csOriginXY > 2) Then
       MarkO = LDfY0 - 30
       MarkL = LDfY0
    Else
       MarkO = LDfYE + 30
       MarkL = LDfYE
    End If
    If .csOrigin Mod 2 = 0 Then
       iSign = 1
    Else
       iSign = -1
    End If
    lNumOfSteps = .NumOfStepsX
    lStep = (LDfX2 - LDfX1) \ (lNumOfSteps)
    Res = GetTextExtentPoint32(hDC, "0123456789", 10, szNumber)
    vSelVec = GetJobVector(Ind, vAxYSelectMatrix, vYHNumJobMatrix)
    NumT = SP(vSelVec, MarkL, 30, lDfSY \ 8, szNumber.cy)
    
    ' Yes, this needs to range from 0 to lNumOfSteps and not from 0 to lNumOfSteps - 1
    For i = 0 To lNumOfSteps
        lx = LDfX1 + i * lStep
        lngScanNumber = .CurrRXMin + i * .csXStep
        If blnShowNETValues Then
            sNumber = Format$(ScanToGANET(Ind, lngScanNumber), "0.000")
        Else
            sNumber = Format$(lngScanNumber, "#,##0")
        End If
        Res = GetTextExtentPoint32(hDC, sNumber, Len(sNumber), szNumber)
        Res = MoveToEx(hDC, lx, MarkO, ptPoint)
        Res = LineTo(hDC, lx, MarkL)
        If bIncludeTextLabels Then
            Res = TextOut(hDC, lx + iSign * szNumber.cx \ 2, NumT, sNumber, Len(sNumber))
        End If
    Next i
End With
End Sub

Private Sub DrawDateFileName(ByVal Ind As Long, ByVal hDC As Long, ByVal TextT As Long)
'used only when printing
Dim sFNDt As String
Dim szFNDt As Size
Dim TextL As Long
Dim vSelVec As Variant
Dim Res As Long
On Error Resume Next

sFNDt = CompactPathString(GelBody(Ind).Caption, 85) & " - " & Format(Now(), "m/dd/yyyy h:nn am/pm")
Res = GetTextExtentPoint32(hDC, sFNDt, Len(sFNDt), szFNDt)
vSelVec = GetJobVector(Ind, vAxXSelectMatrix, vFNDJobMatrix)
TextL = SP(vSelVec, LDfX0, LDfXE, lDfLX, szFNDt.cx)
Res = TextOut(hDC, TextL, TextT, sFNDt, Len(sFNDt))
End Sub

Private Function Get4StarPoints(ByVal X As Long, ByVal Y As Long, ByVal R As Long, ByVal ar As Double) As Variant
Dim pt(7, 1) As Long
Dim i As Integer
On Error Resume Next
For i = 0 To 7
    Select Case i
    Case 0, 4
        pt(i, 0) = X
    Case 1, 3
        pt(i, 0) = X - R \ 4
    Case 5, 7
        pt(i, 0) = X + R \ 4
    Case 2
        pt(i, 0) = X - R
    Case 6
        pt(i, 0) = X + R
    End Select
    Select Case i
    Case 2, 6
        pt(i, 1) = Y
    Case 1, 7
        pt(i, 1) = Y - CLng(R / (4 * ar))
    Case 3, 5
        pt(i, 1) = Y + CLng(R / (4 * ar))
    Case 0
        pt(i, 1) = Y - CLng(R / ar)
    Case 4
        pt(i, 1) = Y + CLng(R / ar)
    End Select
Next i
Get4StarPoints = pt
End Function

Private Function GetTrianglePoints(ByVal X As Long, ByVal Y As Long, ByVal R As Long, ByVal ar As Double) As Variant
Dim pt(2, 1) As Long
Dim HShift As Long
On Error Resume Next

HShift = CLng(0.87 * R)
pt(0, 0) = X
pt(1, 0) = X - HShift
pt(2, 0) = X + HShift
pt(0, 1) = Y - CLng(R / ar)
pt(1, 1) = Y + CLng(R / (2 * ar))
pt(2, 1) = pt(1, 1)
GetTrianglePoints = pt
End Function

Private Function GetHexagonPoints(ByVal X As Long, ByVal Y As Long, ByVal R As Long, ByVal ar As Double) As Variant
Dim pt(5, 1) As Long
Dim VShift As Long
On Error Resume Next

pt(0, 0) = X - R \ 2
pt(1, 0) = X - R
pt(2, 0) = pt(0, 0)
pt(3, 0) = X + R \ 2
pt(4, 0) = X + R
pt(5, 0) = pt(3, 0)
VShift = CLng(0.87 * R / ar)
pt(0, 1) = Y - VShift
pt(1, 1) = Y
pt(2, 1) = Y + VShift
pt(3, 1) = pt(2, 1)
pt(4, 1) = Y
pt(5, 1) = pt(0, 1)
GetHexagonPoints = pt
End Function

Public Sub GelDrawScreen(ByVal Ind As Long, Optional blnIncludeFileNameAndDate As Boolean = False, Optional blnIncludeTextLabels As Boolean = True)
Dim IndDC As Long
Dim OldDC As Long
Dim OldBMP As Long
Dim MemDC As Long
Dim MemBMP As Long
Dim hClipRgn As Long
Dim BMPWidth As Long
Dim BMPHeight As Long
Dim Res As Long

On Error GoTo GelDrawScreenErrorHandler

bSetFileNameDate = blnIncludeFileNameAndDate
bIncludeTextLabels = blnIncludeTextLabels
IndDC = GelBody(Ind).picGraph.hDC
BMPWidth = GelBody(Ind).picGraph.ScaleWidth
BMPHeight = GelBody(Ind).picGraph.ScaleHeight
'drawing is done to memory device, and then BitBlted to screen device
'create compatible bitmaps and device context
MemBMP = CreateCompatibleBitmap(IndDC, BMPWidth, BMPHeight)
MemDC = CreateCompatibleDC(IndDC)
'select bitmap to device context
OldBMP = SelectObject(MemDC, MemBMP)
OldDC = SaveDC(MemDC)
'do the drawing
GelMetrics Ind, BMPWidth, BMPHeight
GelDrawBackColor MemDC, BMPWidth, BMPHeight
GelCooSys Ind, MemDC
GelDrawCooSys Ind, MemDC
GelDrawLegend Ind, MemDC

hClipRgn = ClippingRegionA(MemDC)
Select Case GelBody(Ind).fgDisplay
Case glNormalDisplay
     GelDrawDataNormal Ind, MemDC
Case glDifferentialDisplay, glChargeStateMapDisplay
     GelDrawDataDiff Ind, MemDC
End Select
GelDrawSelection Ind, MemDC
Res = RestoreDC(MemDC, OldDC)
'copy memory drawing to the screen(picture box)
Res = BitBlt(IndDC, 0, 0, BMPWidth, BMPHeight, MemDC, 0, 0, SRCCOPY)
MemBMP = SelectObject(MemDC, OldBMP)
'delete device GDI objects and destroy memory device context
DeleteObject (hClipRgn)
DeleteObject (MemBMP)
DeleteObject (MemDC)

Exit Sub

GelDrawScreenErrorHandler:
Debug.Print "Error in GelDrawScreen: " & Err.Description
Debug.Assert False
Resume Next

End Sub

Public Sub GelDrawPrinter(ByVal Ind As Long)
Dim Res As Long
Dim PrtDC As Long
Dim OldDC As Long
Dim hClipRgn As Long
On Error Resume Next

PrtDC = Printer.hDC
OldDC = SaveDC(PrtDC)
Printer.Orientation = vbPRORLandscape
Printer.ScaleMode = vbPixels
Printer.Print " "

GelMetrics Ind, Printer.ScaleWidth, Printer.ScaleHeight
GelDrawBackColor PrtDC, Printer.ScaleWidth, Printer.ScaleHeight
GelCooSys Ind, PrtDC
GelDrawCooSys Ind, PrtDC
GelDrawLegend Ind, PrtDC

hClipRgn = ClippingRegionA(PrtDC)
Select Case GelBody(Ind).fgDisplay
Case glNormalDisplay
     GelDrawDataNormal Ind, PrtDC
Case glDifferentialDisplay, glChargeStateMapDisplay
     GelDrawDataDiff Ind, PrtDC
End Select
Printer.EndDoc
Res = RestoreDC(PrtDC, OldDC)
If hClipRgn <> 0 Then DeleteObject (hClipRgn)
End Sub

Public Function GelDrawMetafile(ByVal Ind As Long, Optional blnSaveToDisk As Boolean, Optional strFilePath As String = "", Optional blnCreateEnhancedMetaFile As Boolean = False, Optional blnIncludeFileNameAndCurrentTime As Boolean = True, Optional blnIncludeTextLabels As Boolean = True) As Long
' Returns 0 if success, the error number if an error

Dim Res As Long

Dim lngErrorCode As Long
Dim lngReturn As Long
Dim objMetaFilePic As METAFILEPICT
Dim hGlbMemory As Long, lngGlobalAddress As Long

Dim bytBuffer() As Byte

Dim lngBufferSizeRequired As Long
Dim hMemWinMetaFile As Long
Dim hFileWinMetaFile As Long
    
Dim OldDC As Long
Dim hRefDC As Long
Dim emfDC As Long       'metafile device context
Dim emfHandle As Long   'metafile handle
Dim hClipRgn As Long

Dim iWidthMM As Long
Dim iHeightMM As Long
Dim iWidthPels As Long
Dim iHeightPels As Long
Dim iMMPerPelX As Double
Dim iMMPerPelY As Double
Dim rcRef As Rect           'reference rectangle

On Error GoTo GelDrawMetaFileErrorHandler

hRefDC = GelBody(Ind).picGraph.hDC

iWidthMM = GetDeviceCaps(hRefDC, HORZSIZE)
iHeightMM = GetDeviceCaps(hRefDC, VERTSIZE)
iWidthPels = GetDeviceCaps(hRefDC, HORZRES)
iHeightPels = GetDeviceCaps(hRefDC, VERTRES)

iMMPerPelX = (iWidthMM * 100) / iWidthPels
iMMPerPelY = (iHeightMM * 100) / iHeightPels

rcRef.Top = 0
rcRef.Left = 0
rcRef.Bottom = GelBody(Ind).picGraph.ScaleHeight
rcRef.Right = GelBody(Ind).picGraph.ScaleWidth
'convert to himetric units
rcRef.Left = rcRef.Left * iMMPerPelX
rcRef.Top = rcRef.Top * iMMPerPelY
rcRef.Right = rcRef.Right * iMMPerPelX
rcRef.Bottom = rcRef.Bottom * iMMPerPelY

bSetFileNameDate = blnIncludeFileNameAndCurrentTime
bIncludeTextLabels = blnIncludeTextLabels
If blnSaveToDisk And blnCreateEnhancedMetaFile Then
    ' Only send strFilePath to CreateEnhMetaFile if saving to disk and creating an enhanced meta file
    ' We'll take care of creating a standard meta file on disk below
    emfDC = CreateEnhMetaFile(hRefDC, strFilePath, rcRef, vbNullString)
    If emfDC = 0 Then
        lngErrorCode = 75       ' Path/File access error
    Else
        OldDC = SaveDC(emfDC)
    End If
Else
    emfDC = CreateEnhMetaFile(hRefDC, vbNullString, rcRef, vbNullString)
    OldDC = SaveDC(emfDC)
End If
Debug.Assert bSetFileNameDate = blnIncludeFileNameAndCurrentTime
Debug.Assert bIncludeTextLabels = blnIncludeTextLabels

GelMetrics Ind, GelBody(Ind).picGraph.ScaleWidth, GelBody(Ind).picGraph.ScaleHeight
GelDrawBackColor emfDC, GelBody(Ind).picGraph.ScaleWidth, GelBody(Ind).picGraph.ScaleHeight
GelCooSys Ind, emfDC
GelDrawCooSys Ind, emfDC
GelDrawLegend Ind, emfDC
hClipRgn = ClippingRegionA(emfDC)
Select Case GelBody(Ind).fgDisplay
Case glNormalDisplay
     GelDrawDataNormal Ind, emfDC
Case glDifferentialDisplay, glChargeStateMapDisplay
     GelDrawDataDiff Ind, emfDC
End Select
GelDrawSelection Ind, emfDC
'If Not blnSaveToDisk Then Res = RestoreDC(emfDC, OldDC)
Res = RestoreDC(emfDC, OldDC)
If hClipRgn <> 0 Then DeleteObject (hClipRgn)
emfHandle = CloseEnhMetaFile(emfDC)

If blnCreateEnhancedMetaFile Then
    If Not blnSaveToDisk Then
        ' Copy to clipboard
        Res = OpenClipboard(GelBody(Ind).picGraph.hwnd)
        Res = EmptyClipboard()
        Res = SetClipboardData(CF_ENHMETAFILE, emfHandle)
        Res = CloseClipboard
    End If
Else
    lngBufferSizeRequired = GetWinMetaFileBits(emfHandle, 0, ByVal CLng(0), MM_ANISOTROPIC, hRefDC)
    
    If lngBufferSizeRequired = 0 Then
        ' GetWinMetaFileBits() sometimes returns 0 on the first call; try again
        lngBufferSizeRequired = GetWinMetaFileBits(emfHandle, 0, ByVal CLng(0), MM_ANISOTROPIC, hRefDC)
        If lngBufferSizeRequired = 0 Then
            MsgBox "GetWinMetaFileBits() returned 0; unable to copy the picture to the clipboard"
            lngReturn = DeleteEnhMetaFile(emfHandle)
            lngReturn = DeleteObject(emfDC)
            GelDrawMetafile = lngErrorCode
            bIncludeTextLabels = True
            Exit Function
        End If
    End If

    ReDim bytBuffer(0 To lngBufferSizeRequired - 1)
    
    lngBufferSizeRequired = GetWinMetaFileBits(emfHandle, lngBufferSizeRequired, bytBuffer(0), MM_ANISOTROPIC, hRefDC)

    hMemWinMetaFile = SetMetaFileBitsEx(lngBufferSizeRequired, bytBuffer(0))
    
    If blnSaveToDisk Then
        ' Copy the data in hMemWinMetaFile to disk
        hFileWinMetaFile = CopyMetaFile(hMemWinMetaFile, strFilePath)
        If hFileWinMetaFile = 0 Then lngErrorCode = 75      ' Path/File access error
        
        DeleteMetaFile hFileWinMetaFile
    Else
        ' Copy to clipboard
        objMetaFilePic.mm = MM_ANISOTROPIC
        objMetaFilePic.xExt = GelBody(Ind).picGraph.ScaleWidth * Screen.TwipsPerPixelX * Sqr(3.5)
        objMetaFilePic.yExt = GelBody(Ind).picGraph.ScaleHeight * Screen.TwipsPerPixelY * Sqr(3.5)
        objMetaFilePic.hMF = hMemWinMetaFile
        
        ' Take out hardcoded sizes
        hGlbMemory = GlobalAlloc(GMEM_MOVEABLE, Len(objMetaFilePic))
        
        lngGlobalAddress = GlobalLock(hGlbMemory)
        agCopyData objMetaFilePic, ByVal lngGlobalAddress&, Len(objMetaFilePic)
        Res = GlobalUnlock(hGlbMemory)
        
        Res = OpenClipboard(GelBody(Ind).picGraph.hwnd)
        Res = EmptyClipboard()
        Res = SetClipboardData(CF_METAFILEPICT, hGlbMemory)
        Res = CloseClipboard
    End If
    
    DeleteMetaFile hMemWinMetaFile

End If

lngReturn = DeleteEnhMetaFile(emfHandle)
Debug.Assert lngReturn <> 0

' The following should actually return a non-zero value, and it's not
' Thus, it's probably inappropriate
lngReturn = DeleteObject(emfDC)

bIncludeTextLabels = True
GelDrawMetafile = lngErrorCode

Exit Function

GelDrawMetaFileErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then MsgBox "An error occurred while copying/saving the current picture: " & Err.Description, vbExclamation + vbOKOnly, "Error"
    GelDrawMetafile = AssureNonZero(Err.Number)
    bIncludeTextLabels = True
End Function

Public Sub GelDrawSelectionAdd(ByVal Ind As Long, _
                               ByVal SelType As Integer, _
                               ByVal SelID As Long)
'this procedure is isolated to make drawings faster and with less
'flickering; if only new point is added to selection there is no need
'to redraw everything; there is no need here to draw to memory
Dim IndDC As Long
Dim OldDC As Long
Dim hClipRgn As Long
Dim BMPWidth As Long
Dim BMPHeight As Long
Dim OldPen As Long
Dim OldBrush As Long
Dim OldFont As Long
Dim NewFont As Long
Dim SX As Long, SY As Long
Dim Res As Long
Dim Sel() As Long
Dim AspRatio As Double
On Error Resume Next

bSetFileNameDate = False
bIncludeTextLabels = True
IndDC = GelBody(Ind).picGraph.hDC
BMPWidth = GelBody(Ind).picGraph.ScaleWidth
BMPHeight = GelBody(Ind).picGraph.ScaleHeight
OldDC = SaveDC(IndDC)
'do the drawing
GelMetrics Ind, BMPWidth, BMPHeight
GelCooSys Ind, IndDC
If glSelColor = -1 Then     'mark selection with flag
   OldPen = SelectObject(IndDC, hForeColorPen)
   SetGraphFont Ind, IndDC, OldFont, NewFont
   Res = SetROP2(IndDC, R2_MASKPEN)
   Res = SetBkMode(IndDC, TRANSPARENT)
   hClipRgn = ClippingRegionA(IndDC)
   'draw the selection here
   GetDrawSelectionOffsets Ind, SX, SY
   Select Case GelBody(Ind).fgDisplay
   Case glNormalDisplay
     Select Case SelType
     Case glCSType
          NormalDisplayCSSelected Ind, IndDC, SelID, "C" & SelID, SX, SY
     Case glIsoType
          NormalDisplayIsoSelected Ind, IndDC, SelID, "I-" & SelID, SX, SY
     End Select
   Case glDifferentialDisplay, glChargeStateMapDisplay
     Select Case SelType
     Case glCSType
          DiffDisplayCSSelected Ind, IndDC, SelID, "C" & SelID, SX, SY
     Case glIsoType
          DiffDisplayIsoSelected Ind, IndDC, SelID, "I-" & SelID, SX, SY
     End Select
   End Select
   'restore old GDI objects
   Res = SelectObject(IndDC, OldFont)
   If NewFont <> 0 Then DeleteObject (NewFont)
   Res = SelectObject(IndDC, OldPen)
Else            'mark selection with color
   OldPen = SelectObject(IndDC, hSelColorPen)
   OldBrush = SelectObject(IndDC, hSelColorBrush)
   Res = SetROP2(IndDC, R2_MASKPEN)
   Res = SetBkMode(IndDC, TRANSPARENT)
   ReDim Sel(1 To 1)
   Sel(1) = SelID
   AspRatio = GelData(Ind).Preferences.AbuAspectRatio
   Select Case SelType
   Case glCSType
      Select Case glCSShape
      Case glShapeEli
         SelDrawEliCS Ind, IndDC, AspRatio, Sel
      Case glShapeRec
         SelDrawRecCS Ind, IndDC, AspRatio, Sel
      Case glShapeRRe
         SelDrawRReCS Ind, IndDC, AspRatio, Sel
      Case glShapeTri
         SelDrawTriCS Ind, IndDC, AspRatio, Sel
      Case glShapeSta
         SelDrawStaCS Ind, IndDC, AspRatio, Sel
      Case glShapeHex
         SelDrawHexCS Ind, IndDC, AspRatio, Sel
      End Select
   Case glIsoType
      Select Case glIsoShape
      Case glShapeEli
         SelDrawEliIso Ind, IndDC, AspRatio, Sel
      Case glShapeRec
         SelDrawRecIso Ind, IndDC, AspRatio, Sel
      Case glShapeRRe
         SelDrawRReIso Ind, IndDC, AspRatio, Sel
      Case glShapeTri
         SelDrawTriIso Ind, IndDC, AspRatio, Sel
      Case glShapeSta
         SelDrawStaIso Ind, IndDC, AspRatio, Sel
      Case glShapeHex
         SelDrawHexIso Ind, IndDC, AspRatio, Sel
      End Select
   End Select
   Res = SelectObject(IndDC, OldBrush)
   Res = SelectObject(IndDC, OldPen)
End If
Res = RestoreDC(IndDC, OldDC)
'delete device GDI objects and destroy memory device context
DeleteObject (hClipRgn)
End Sub


Private Sub GelDrawSelection(ByVal Ind As Long, ByVal hDC As Long)
Dim OldPen As Long
Dim OldFont As Long
Dim OldBrush As Long
Dim NewFont As Long
Dim Res As Long
On Error Resume Next
If glSelColor = -1 Then         'mark selection with flags
   OldPen = SelectObject(hDC, hForeColorPen)
   SetGraphFont Ind, hDC, OldFont, NewFont
   Res = SetROP2(hDC, R2_MASKPEN)
   Res = SetBkMode(hDC, TRANSPARENT)
   Select Case GelBody(Ind).fgDisplay
   Case glNormalDisplay
        DrawSelectionNormal Ind, hDC
   Case glDifferentialDisplay, glChargeStateMapDisplay
        DrawSelectionDifferential Ind, hDC
   End Select
   Res = SelectObject(hDC, OldFont)
   If NewFont <> 0 Then DeleteObject (NewFont)
   Res = SelectObject(hDC, OldPen)
Else                            'mark selection with color
   OldPen = SelectObject(hDC, hSelColorPen)
   OldBrush = SelectObject(hDC, hSelColorBrush)
   Res = SetROP2(hDC, R2_MASKPEN)
   Res = SetBkMode(hDC, TRANSPARENT)
   DrawSelectionColor Ind, hDC
   Res = SelectObject(hDC, OldBrush)
   Res = SelectObject(hDC, OldPen)
End If
End Sub

Private Sub GetSelFlagPoints(pts() As POINTAPI)
'flag is something like letter F flag in 300x360 rectangle
On Error Resume Next
pts(1).X = pts(0).X + 83    '100\6
pts(2).X = pts(0).X + 300
pts(3).X = pts(0).X + 100   '300\3
pts(4).X = pts(2).X
pts(1).Y = pts(0).Y + 330   '11*360\12
pts(2).Y = pts(1).Y
pts(3).Y = pts(0).Y + 360
pts(4).Y = pts(3).Y
End Sub

Private Sub DrawSelectedFlag(ByVal hDC As Long, _
                             pts() As POINTAPI, _
                             ByVal Lbl As String, _
                             ByVal SLblX As Long, _
                             ByVal SlblY As Long)
'hdc - handle to device context, pts - array of points used for drawing,
'lbl - text label to draw next to flag, SLblX, SLblY shifts used to
'position the drawing
Dim SomePenPos  As POINTAPI
Dim Res As Long
MoveToEx hDC, pts(0).X, pts(0).Y, SomePenPos
LineTo hDC, pts(3).X, pts(3).Y
LineTo hDC, pts(4).X, pts(4).Y
MoveToEx hDC, pts(1).X, pts(1).Y, SomePenPos
LineTo hDC, pts(2).X, pts(2).Y
'write label text
Res = TextOut(hDC, pts(4).X + SLblX, pts(4).Y + SlblY, Lbl, Len(Lbl))
End Sub


Private Sub SetGraphFont(ByVal Ind As Long, ByVal hDC As Long, lOldFont As Long, lNewFont As Long)
'sets font appropriate for drawing on the Graph
Dim ldffont As Long
Dim lfLogFont As LOGFONT
Dim Res As Long
On Error Resume Next
'get the font from the graph picture box control (Arial Narrow)
ldffont = SelectObject(GelBody(Ind).picGraph.hDC, GetStockObject(SYSTEM_FONT))
Res = GetObjectAPI(ldffont, Len(lfLogFont), lfLogFont)
Res = SelectObject(GelBody(Ind).picGraph.hDC, ldffont)

'create new logical font
lfLogFont.lfWidth = 75          ' Was 65
lfLogFont.lfHeight = 290        ' Was 300
lNewFont = CreateFontIndirect(lfLogFont)

'select newly created logical font to DC
lOldFont = SelectObject(hDC, lNewFont)
Res = SetTextColor(hDC, glForeColor)
End Sub


Private Sub DrawSelectionNormal(ByVal Ind As Long, ByVal hDC As Long)
'draws selection to the device context
Dim i As Long, iSel As Long
Dim SX As Long, SY As Long    'used to position output
On Error Resume Next
GetDrawSelectionOffsets Ind, SX, SY
With GelBody(Ind).GelSel
  If .CSSelCnt > 0 Then
     For i = 1 To .CSSelCnt
         iSel = .Value(i, glCSType)
         NormalDisplayCSSelected Ind, hDC, iSel, "C" & iSel, SX, SY
     Next i
  End If
  If .IsoSelCnt > 0 Then
     For i = 1 To .IsoSelCnt
         iSel = .Value(i, glIsoType)
         NormalDisplayIsoSelected Ind, hDC, iSel, "I-" & iSel, SX, SY
     Next i
  End If
End With
End Sub

Private Sub NormalDisplayCSSelected(ByVal Ind As Long, _
                                    ByVal hDC As Long, _
                                    ByVal SelID As Long, _
                                    ByVal SelText As String, _
                                    ByVal SX As Integer, _
                                    ByVal SY As Integer)
Dim ptAPIs(4) As POINTAPI
Dim szLbl As Size
Dim Res As Long
With GelDraw(Ind)
  If .CSID(SelID) > 0 And .CSR(SelID) > 0 Then
     ptAPIs(0).X = .CSX(SelID)
     ptAPIs(0).Y = .CSY(SelID)
     GetSelFlagPoints ptAPIs
     Res = GetTextExtentPoint32(hDC, SelText, Len(SelText), szLbl)
     DrawSelectedFlag hDC, ptAPIs, SelText, szLbl.cx * SX, szLbl.cy * SY
  End If
End With
End Sub

Private Sub NormalDisplayIsoSelected(ByVal Ind As Long, _
                                     ByVal hDC As Long, _
                                     ByVal SelID As Long, _
                                     ByVal SelText As String, _
                                     ByVal SX As Integer, _
                                     ByVal SY As Integer)
Dim ptAPIs(4) As POINTAPI
Dim szLbl As Size
Dim Res As Long
With GelDraw(Ind)
  If .IsoID(SelID) > 0 And .IsoR(SelID) > 0 Then
     ptAPIs(0).X = .IsoX(SelID)
     ptAPIs(0).Y = .IsoY(SelID)
     GetSelFlagPoints ptAPIs
     Res = GetTextExtentPoint32(hDC, SelText, Len(SelText), szLbl)
     DrawSelectedFlag hDC, ptAPIs, SelText, szLbl.cx * SX, szLbl.cy * SY
  End If
End With
End Sub

Private Sub DrawSelectionDifferential(ByVal Ind As Long, ByVal hDC As Long)
'draws selection to the device context
Dim i As Long, iSel As Long
Dim SX As Long, SY As Long    'used to position output
On Error Resume Next
GetDrawSelectionOffsets Ind, SX, SY
With GelBody(Ind).GelSel
  If .CSSelCnt > 0 Then
     For i = 1 To .CSSelCnt
        iSel = .Value(i, glCSType)
        DiffDisplayCSSelected Ind, hDC, iSel, "C" & iSel, SX, SY
     Next i
  End If
  If .IsoSelCnt > 0 Then
     For i = 1 To .IsoSelCnt
        iSel = .Value(i, glIsoType)
        DiffDisplayIsoSelected Ind, hDC, iSel, "I-" & iSel, SX, SY
     Next i
  End If
End With
End Sub

Private Sub DiffDisplayCSSelected(ByVal Ind As Long, _
                                  ByVal hDC As Long, _
                                  ByVal SelID As Long, _
                                  ByVal SelText As String, _
                                  ByVal SX As Integer, _
                                  ByVal SY As Integer)
Dim ptAPIs(4) As POINTAPI
Dim szLbl As Size
Dim Res As Long
With GelDraw(Ind)
   If .CSID(SelID) > 0 And .CSR(SelID) > 0 And .CSERClr(SelID) <> glDONT_DISPLAY Then
      ptAPIs(0).X = .CSX(SelID)
      ptAPIs(0).Y = .CSY(SelID)
      GetSelFlagPoints ptAPIs
      Res = GetTextExtentPoint32(hDC, SelText, Len(SelText), szLbl)
      DrawSelectedFlag hDC, ptAPIs, SelText, szLbl.cx * SX, szLbl.cy * SY
   End If
End With
End Sub

Private Sub DiffDisplayIsoSelected(ByVal Ind As Long, _
                                   ByVal hDC As Long, _
                                   ByVal SelID As Long, _
                                   ByVal SelText As String, _
                                   ByVal SX As Integer, _
                                   ByVal SY As Integer)
Dim ptAPIs(4) As POINTAPI
Dim szLbl As Size
Dim Res As Long
With GelDraw(Ind)
   If .IsoID(SelID) > 0 And .IsoR(SelID) > 0 And .IsoERClr(SelID) >= 0 Then
      ptAPIs(0).X = .IsoX(SelID)
      ptAPIs(0).Y = .IsoY(SelID)
      GetSelFlagPoints ptAPIs
      Res = GetTextExtentPoint32(hDC, SelText, Len(SelText), szLbl)
      DrawSelectedFlag hDC, ptAPIs, SelText, szLbl.cx * SX, szLbl.cy * SY
   End If
End With
End Sub

Private Sub GetDrawSelectionOffsets(ByVal Ind As Long, OffsX As Long, OffsY As Long)
'shortcut - for drawing selection - picks directions of selection
'label drawings depending on the coordinate system orientation
With GelBody(Ind).csMyCooSys
   If .csXOrient = glNormal Then
      OffsX = -1
   Else
      OffsX = 0
   End If
   If .csYOrient = glNormal Then
      OffsY = 1
   Else
      OffsY = 0
   End If
End With
End Sub

Private Sub DrawSelectionColor(ByVal Ind As Long, ByVal hDC As Long)
'draws selection to the device context as colored spots;
'first pick the selection and then draw all together
Dim CSSel() As Long
Dim IsoSel() As Long
Dim AspRatio As Double
On Error Resume Next
AspRatio = GelData(Ind).Preferences.AbuAspectRatio
With GelBody(Ind).GelSel
   If .CSSelCnt > 0 Then
      .GetCSSel CSSel()
      'draw cs selection
      Select Case glCSShape
      Case glShapeEli
         SelDrawEliCS Ind, hDC, AspRatio, CSSel()
      Case glShapeRec
         SelDrawRecCS Ind, hDC, AspRatio, CSSel()
      Case glShapeRRe
         SelDrawRReCS Ind, hDC, AspRatio, CSSel()
      Case glShapeTri
         SelDrawTriCS Ind, hDC, AspRatio, CSSel()
      Case glShapeSta
         SelDrawStaCS Ind, hDC, AspRatio, CSSel()
      Case glShapeHex
         SelDrawHexCS Ind, hDC, AspRatio, CSSel()
      End Select
   End If
   If .IsoSelCnt > 0 Then
      .GetIsoSel IsoSel()
      'draw iso selection
      Select Case glIsoShape
      Case glShapeEli
         SelDrawEliIso Ind, hDC, AspRatio, IsoSel()
      Case glShapeRec
         SelDrawRecIso Ind, hDC, AspRatio, IsoSel()
      Case glShapeRRe
         SelDrawRReIso Ind, hDC, AspRatio, IsoSel()
      Case glShapeTri
         SelDrawTriIso Ind, hDC, AspRatio, IsoSel()
      Case glShapeSta
         SelDrawStaIso Ind, hDC, AspRatio, IsoSel()
      Case glShapeHex
         SelDrawHexIso Ind, hDC, AspRatio, IsoSel()
      End Select
   End If
End With
End Sub


Private Sub SelDrawEliCS(ByVal Ind As Long, ByVal hDC As Long, ByVal ar As Double, Sel() As Long)
Dim h As Long, v As Long
Dim i As Long
Dim SelCnt As Long
On Error Resume Next
SelCnt = UBound(Sel)
If Err Then Exit Sub
If SelCnt > 0 Then
   With GelDraw(Ind)            'no checking for visibility here (selected is visible)
     For i = 1 To SelCnt
       h = .CSR(Sel(i)) \ 2
       v = CLng(.CSR(Sel(i)) / (2 * ar))
       Ellipse hDC, .CSX(Sel(i)) - h, .CSY(Sel(i)) - v, .CSX(Sel(i)) + h, .CSY(Sel(i)) + v
     Next i
   End With
End If
End Sub

Private Sub SelDrawEliIso(ByVal Ind As Long, ByVal hDC As Long, ByVal ar As Double, Sel() As Long)
Dim h As Long, v As Long
Dim i As Long
Dim SelCnt As Long
On Error Resume Next
SelCnt = UBound(Sel)
If Err Then Exit Sub
If SelCnt > 0 Then
   With GelDraw(Ind)
     For i = 1 To SelCnt
        h = .IsoR(Sel(i)) \ 2
        v = CLng(.IsoR(Sel(i)) / (2 * ar))
        Ellipse hDC, .IsoX(Sel(i)) - h, .IsoY(Sel(i)) - v, .IsoX(Sel(i)) + h, .IsoY(Sel(i)) + v
     Next i
   End With
End If
End Sub

Private Sub SelDrawTriCS(ByVal Ind As Long, ByVal hDC As Long, ByVal ar As Double, Sel() As Long)
'this 0.87 is approx. SQRT(3)\2
Dim vptAPIs As Variant
Dim ptAPIs(2) As POINTAPI
Dim i As Long, j As Long
Dim SelCnt As Long
On Error Resume Next
SelCnt = UBound(Sel)
If Err Then Exit Sub
If SelCnt > 0 Then
   With GelDraw(Ind)
     For i = 1 To SelCnt
        vptAPIs = GetTrianglePoints(.CSX(Sel(i)), .CSY(Sel(i)), .CSR(Sel(i)) \ 2, ar)
        For j = 0 To 2
           ptAPIs(j).X = vptAPIs(j, 0)
           ptAPIs(j).Y = vptAPIs(j, 1)
        Next j
        Polygon hDC, ptAPIs(0), 3
     Next i
   End With
End If
End Sub

Private Sub SelDrawTriIso(ByVal Ind As Long, ByVal hDC As Long, ByVal ar As Double, Sel() As Long)
Dim vptAPIs As Variant
Dim ptAPIs(2) As POINTAPI
Dim i As Long, j As Long
Dim SelCnt As Long
On Error Resume Next
SelCnt = UBound(Sel)
If Err Then Exit Sub
If SelCnt > 0 Then
   With GelDraw(Ind)
     For i = 1 To SelCnt
        vptAPIs = GetTrianglePoints(.IsoX(Sel(i)), .IsoY(Sel(i)), .IsoR(Sel(i)) \ 2, ar)
        For j = 0 To 2
            ptAPIs(j).X = vptAPIs(j, 0)
            ptAPIs(j).Y = vptAPIs(j, 1)
        Next j
        Polygon hDC, ptAPIs(0), 3
     Next i
   End With
End If
End Sub

Private Sub SelDrawStaCS(ByVal Ind As Long, ByVal hDC As Long, ByVal ar As Double, Sel() As Long)
Dim vptAPIs As Variant
Dim ptAPIs(7) As POINTAPI
Dim i As Long, j As Long
Dim SelCnt As Long
On Error Resume Next
SelCnt = UBound(Sel)
If Err Then Exit Sub
If SelCnt > 0 Then
   With GelDraw(Ind)
     For i = 1 To SelCnt
        vptAPIs = Get4StarPoints(.CSX(Sel(i)), .CSY(Sel(i)), .CSR(Sel(i)) \ 2, ar)
        For j = 0 To 7
            ptAPIs(j).X = vptAPIs(j, 0)
            ptAPIs(j).Y = vptAPIs(j, 1)
        Next j
        Polygon hDC, ptAPIs(0), 8
     Next i
   End With
End If
End Sub

Private Sub SelDrawStaIso(ByVal Ind As Long, ByVal hDC As Long, ByVal ar As Double, Sel() As Long)
Dim vptAPIs As Variant
Dim ptAPIs(7) As POINTAPI
Dim i As Long, j As Long
Dim SelCnt As Long
On Error Resume Next
SelCnt = UBound(Sel)
If Err Then Exit Sub
If SelCnt > 0 Then
   With GelDraw(Ind)
     For i = 1 To SelCnt
        vptAPIs = Get4StarPoints(.IsoX(Sel(i)), .IsoY(Sel(i)), .IsoR(Sel(i)) \ 2, ar)
        For j = 0 To 7
            ptAPIs(j).X = vptAPIs(j, 0)
            ptAPIs(j).Y = vptAPIs(j, 1)
        Next j
        Polygon hDC, ptAPIs(0), 8
     Next i
   End With
End If
End Sub

Private Sub SelDrawHexIso(ByVal Ind As Long, ByVal hDC As Long, ByVal ar As Double, Sel() As Long)
Dim vptAPIs As Variant
Dim ptAPIs(5) As POINTAPI
Dim i As Long, j As Long
Dim SelCnt As Long
On Error Resume Next
SelCnt = UBound(Sel)
If Err Then Exit Sub
If SelCnt > 0 Then
   With GelDraw(Ind)
     For i = 1 To SelCnt
        vptAPIs = GetHexagonPoints(.IsoX(Sel(i)), .IsoY(Sel(i)), .IsoR(Sel(i)) \ 2, ar)
        For j = 0 To 5
            ptAPIs(j).X = vptAPIs(j, 0)
            ptAPIs(j).Y = vptAPIs(j, 1)
        Next j
        Polygon hDC, ptAPIs(0), 6
     Next i
   End With
End If
End Sub

Private Sub SelDrawHexCS(ByVal Ind As Long, ByVal hDC As Long, ByVal ar As Double, Sel() As Long)
Dim vptAPIs As Variant
Dim ptAPIs(5) As POINTAPI
Dim i As Long, j As Long
Dim SelCnt As Long
On Error Resume Next
SelCnt = UBound(Sel)
If Err Then Exit Sub
If SelCnt > 0 Then
   With GelDraw(Ind)
     For i = 1 To SelCnt
        vptAPIs = GetHexagonPoints(.CSX(Sel(i)), .CSY(Sel(i)), .CSR(Sel(i)) \ 2, ar)
        For j = 0 To 5
            ptAPIs(j).X = vptAPIs(j, 0)
            ptAPIs(j).Y = vptAPIs(j, 1)
        Next j
        Polygon hDC, ptAPIs(0), 6
     Next i
   End With
End If
End Sub

Private Sub SelDrawRecCS(ByVal Ind As Long, ByVal hDC As Long, ByVal ar As Double, Sel() As Long)
'this is called only if .CSCount>0
Dim i As Long
Dim h As Long, v As Long
Dim SelCnt As Long
On Error Resume Next
SelCnt = UBound(Sel)
If Err Then Exit Sub
If SelCnt > 0 Then
   With GelDraw(Ind)
     For i = 1 To SelCnt
         h = .CSR(Sel(i)) \ 2
         v = CLng(.CSR(Sel(i)) / (2 * ar))
         Rectangle hDC, .CSX(Sel(i)) - h, .CSY(Sel(i)) - v, .CSX(Sel(i)) + h, .CSY(Sel(i)) + v
     Next i
   End With
End If
End Sub

Private Sub SelDrawRReCS(ByVal Ind As Long, ByVal hDC As Long, ByVal ar As Double, Sel() As Long)
Dim i As Long
Dim h As Long, v As Long
Dim SelCnt As Long
'On Error Resume Next
SelCnt = UBound(Sel)
If Err Then Exit Sub
If SelCnt > 0 Then
   With GelDraw(Ind)
     For i = 1 To SelCnt
         h = .CSR(Sel(i)) \ 2
         v = CLng(.CSR(Sel(i)) / (2 * ar))
         RoundRect hDC, .CSX(Sel(i)) - h, .CSY(Sel(i)) - v, .CSX(Sel(i)) + h, .CSY(Sel(i)) + v, h, v
     Next i
   End With
End If
End Sub

Private Sub SelDrawRecIso(ByVal Ind As Long, ByVal hDC As Long, ByVal ar As Double, Sel() As Long)
Dim i As Long
Dim h As Long, v As Long
Dim SelCnt As Long
On Error Resume Next
SelCnt = UBound(Sel)
If Err Then Exit Sub
If SelCnt > 0 Then
   With GelDraw(Ind)
     For i = 1 To SelCnt
         h = .IsoR(Sel(i)) \ 2
         v = CLng(.IsoR(Sel(i)) / (2 * ar))
         Rectangle hDC, .IsoX(Sel(i)) - h, .IsoY(Sel(i)) - v, .IsoX(Sel(i)) + h, .IsoY(Sel(i)) + v
     Next i
   End With
End If
End Sub

Private Sub SelDrawRReIso(ByVal Ind As Long, ByVal hDC As Long, ByVal ar As Double, Sel() As Long)
Dim i As Long
Dim h As Long, v As Long
Dim SelCnt As Long
On Error Resume Next
SelCnt = UBound(Sel)
If Err Then Exit Sub
If SelCnt > 0 Then
   With GelDraw(Ind)
     For i = 1 To SelCnt
         h = .IsoR(Sel(i)) \ 2
         v = CLng(.IsoR(Sel(i)) / (2 * ar))
         RoundRect hDC, .IsoX(Sel(i)) - h, .IsoY(Sel(i)) - v, .IsoX(Sel(i)) + h, .IsoY(Sel(i)) + v, h, v
     Next i
   End With
End If
End Sub

Private Sub DrawLegendNormal(ByVal Ind As Long, ByVal hDC As Long)
'-------------------------------------------------------------------
'draws legend for normal display
'-------------------------------------------------------------------
Dim lOldFont As Long, lNewFont As Long
Dim OldBrush As Long
Dim OldPen As Long
Dim Res As Long
Dim ShH As Long, ShW As Long
Dim ShCX As Long, ShCY As Long
Dim X01 As Long, XE1 As Long, XH1 As Long

Dim HLblX As Long, HLblY As Long
Dim VLblX As Long, VLblY As Long
Dim sLgndLblLN As String, sLgndLblRN As String

Dim sHLbl As String, sVLbl As String
Dim szHLbl As Size, szVLbl As Size
Dim szLLblLN As Size, szLLblRN As Size
Dim vSelVec As Variant

On Error Resume Next
'create new font suitable for writing on the graph
SetGraphFont Ind, hDC, lOldFont, lNewFont

Res = SetROP2(hDC, R2_MASKPEN)
Res = SetBkMode(hDC, TRANSPARENT)

'set labels
sLgndLblLN = "Charge State"
sLgndLblRN = "Isotopic"

Select Case GelBody(Ind).csMyCooSys.csType
Case glPICooSys
     sHLbl = "pI"
Case glFNCooSys
     sHLbl = "Scan number"
Case glNETCooSys
     sHLbl = "NET"
End Select

Select Case GelBody(Ind).csMyCooSys.csYScale
Case glVAxisLin
    sVLbl = "Monoisotopic Mass"
Case glVAxisLog
    sVLbl = "Monoisotopic Mass (Log Scale)"
End Select

'get the sizes of the axes and legend labels
Res = GetTextExtentPoint32(hDC, sHLbl, Len(sHLbl), szHLbl)
Res = GetTextExtentPoint32(hDC, sVLbl, Len(sVLbl), szVLbl)
Res = GetTextExtentPoint32(hDC, sLgndLblLN, Len(sLgndLblLN), szLLblLN)
Res = GetTextExtentPoint32(hDC, sLgndLblRN, Len(sLgndLblRN), szLLblRN)

'draw coordinate axes labels
vSelVec = GetJobVector(Ind, vAxXSelectMatrix, vHXJobMatrix)
HLblX = SP(vSelVec, LDfX0, LDfXE, LDfSX \ 2, szHLbl.cx)
vSelVec = GetJobVector(Ind, vAxXSelectMatrix, vVXJobMatrix)
VLblX = SP(vSelVec, LDfX0, LDfXE, LDfSX \ 2, szVLbl.cx)
vSelVec = GetJobVector(Ind, vAxYSelectMatrix, vHYJobMatrix)
HLblY = SP(vSelVec, LDfY0, LDfYE, lDfLY \ 2, szHLbl.cy)
vSelVec = GetJobVector(Ind, vAxYSelectMatrix, vVYJobMatrix)
VLblY = SP(vSelVec, LDfY0, LDfYE, 0, szVLbl.cy)
If bIncludeTextLabels Then
    Res = TextOut(hDC, HLblX, HLblY, sHLbl, Len(sHLbl))
    Res = TextOut(hDC, VLblX, VLblY, sVLbl, Len(sVLbl))
End If

   'draw legend labels and boxes

   OldBrush = SelectObject(hDC, GetStockObject(NULL_BRUSH)) 'prepare transparent brush
   OldPen = SelectObject(hDC, hForeColorPen)
   X01 = LDfX0 + 1100
   XE1 = LDfXE - 1100
   XH1 = (XE1 - X01) \ 2

   'draw shape and charge state/isotopic colors legend
   ShH = szVLbl.cy
   vSelVec = GetJobVector(Ind, vAxYSelectMatrix, vVYJobMatrix)
   ShCY = SP(vSelVec, LDfY0, LDfYE, ShH \ 2, 0)
   ShW = 100
   vSelVec = GetJobVector(Ind, vAxXSelectMatrix, vLgndCSMatrix)

   If GelDraw(Ind).CSVisible Then
       'charge state legend
       ShCX = SP(vSelVec, X01 + XH1, XE1 - XH1, 3 * LDfSX \ 4 + 200, szLLblLN.cx)
       If bIncludeTextLabels Then
           Res = TextOut(hDC, ShCX, VLblY, sLgndLblLN, Len(sLgndLblLN))
       End If
       ShCX = SP(vSelVec, X01 + XH1, XE1 - XH1, LDfSX \ 4 + 100, 0)
       Res = SelectObject(hDC, hCSColorBrush)
       DrawLegendShape hDC, glCSShape, ShCX, ShCY
   End If

   If GelDraw(Ind).IsoVisible And GelDraw(Ind).CSVisible Then
       'isotopic legend
       ShCX = SP(vSelVec, X01 + XH1, XE1 - XH1, -(3 * LDfSX \ 4 + 200), 0)
       If bIncludeTextLabels Then
           Res = TextOut(hDC, ShCX, VLblY, sLgndLblRN, Len(sLgndLblRN))
       End If
       ShCX = SP(vSelVec, X01 + XH1, XE1 - XH1, -(LDfSX \ 4 + 100), 0)
       Res = SelectObject(hDC, hIsoColorBrush)
       DrawLegendShape hDC, glIsoShape, ShCX, ShCY
   End If

Res = SelectObject(hDC, hForeColorPen)
GelDrawVAxisNumbers Ind, hDC
Select Case GelBody(Ind).csMyCooSys.csType
Case glFNCooSys
     GelDrawFNHAxisNumbers Ind, hDC, False
Case glNETCooSys
     GelDrawFNHAxisNumbers Ind, hDC, True
Case glPICooSys
     GelDrawpiHAxisNumbers Ind, hDC
End Select

If bSetFileNameDate Then DrawDateFileName Ind, hDC, HLblY

Res = SelectObject(hDC, OldPen)
Res = SelectObject(hDC, OldBrush)
Res = SelectObject(hDC, lOldFont)
If lNewFont <> 0 Then DeleteObject (lNewFont)
End Sub


Private Sub DrawLegendDifferential(ByVal Ind As Long, ByVal hDC As Long)
'-------------------------------------------------------------------------
'draws legend and markings on the coordinate axes for differential display
'-------------------------------------------------------------------------
Dim lOldFont As Long, lNewFont As Long
Dim OldBrush As Long
Dim OldPen As Long
Dim Res As Long
Dim ShH As Long, ShW As Long
Dim ShT As Long, ShB As Long
Dim ShCX As Long, ShCY As Long
Dim X01 As Long, XE1 As Long
Dim i As Integer

Dim HLblX As Long, HLblY As Long
Dim VLblX As Long, VLblY As Long
Dim sLgndLblLN As String, sLgndLblRN As String
Dim sLgndLblLD As String, sLgndLblRD As String

Dim sHLbl As String, sVLbl As String
Dim szHLbl As Size, szVLbl As Size
Dim szLLblLN As Size, szLLblRN As Size
Dim szLLblLD As Size, szLLblRD As Size
Dim vSelVec As Variant

On Error Resume Next
'create new font suitable for writing on the graph
SetGraphFont Ind, hDC, lOldFont, lNewFont

Res = SetROP2(hDC, R2_MASKPEN)
Res = SetBkMode(hDC, TRANSPARENT)

'set labels
sLgndLblLN = "Charge State"
sLgndLblRN = "Isotopic"
Select Case GelData(Ind).Preferences.DRDefinition
Case glNormal
     sLgndLblLD = "Suppressed"
     sLgndLblRD = "Induced"
Case glReverse
     sLgndLblLD = "Induced"
     sLgndLblRD = "Suppressed"
End Select

Select Case GelBody(Ind).csMyCooSys.csType
Case glPICooSys
     sHLbl = "pI"
Case glFNCooSys
     sHLbl = "Scan number"
Case glNETCooSys
     sHLbl = "NET"
End Select

Select Case GelBody(Ind).csMyCooSys.csYScale
Case glVAxisLin
    sVLbl = "Monoisotopic Mass"
Case glVAxisLog
    sVLbl = "Monoisotopic Mass (Log Scale)"
End Select

'get the sizes of the axes and legend labels
Res = GetTextExtentPoint32(hDC, sHLbl, Len(sHLbl), szHLbl)
Res = GetTextExtentPoint32(hDC, sVLbl, Len(sVLbl), szVLbl)
Res = GetTextExtentPoint32(hDC, sLgndLblLN, Len(sLgndLblLN), szLLblLN)
Res = GetTextExtentPoint32(hDC, sLgndLblRN, Len(sLgndLblRN), szLLblRN)
Res = GetTextExtentPoint32(hDC, sLgndLblLD, Len(sLgndLblLD), szLLblLD)
Res = GetTextExtentPoint32(hDC, sLgndLblRD, Len(sLgndLblRD), szLLblRD)

'draw coordinate axes labels
vSelVec = GetJobVector(Ind, vAxXSelectMatrix, vHXJobMatrix)
HLblX = SP(vSelVec, LDfX0, LDfXE, LDfSX \ 2, szHLbl.cx)
vSelVec = GetJobVector(Ind, vAxXSelectMatrix, vVXJobMatrix)
VLblX = SP(vSelVec, LDfX0, LDfXE, LDfSX \ 2, szVLbl.cx)
vSelVec = GetJobVector(Ind, vAxYSelectMatrix, vHYJobMatrix)
HLblY = SP(vSelVec, LDfY0, LDfYE, lDfLY \ 2, szHLbl.cy)
vSelVec = GetJobVector(Ind, vAxYSelectMatrix, vVYJobMatrix)
VLblY = SP(vSelVec, LDfY0, LDfYE, 0, szVLbl.cy)
If bIncludeTextLabels Then
    Res = TextOut(hDC, HLblX, HLblY, sHLbl, Len(sHLbl))
    Res = TextOut(hDC, VLblX, VLblY, sVLbl, Len(sVLbl))
End If

   'draw legend labels and boxes

   OldBrush = SelectObject(hDC, GetStockObject(NULL_BRUSH)) 'prepare transparent brush
   OldPen = SelectObject(hDC, hForeColorPen)
   X01 = LDfX0 + 1100
   XE1 = LDfXE - 1100

   'draw shape and charge state/isotopic colors legend
   ShH = szVLbl.cy
   vSelVec = GetJobVector(Ind, vAxYSelectMatrix, vVYJobMatrix)
   ShCY = SP(vSelVec, LDfY0, LDfYE, ShH \ 2, 0)
   ShW = 100
   vSelVec = GetJobVector(Ind, vAxXSelectMatrix, vLgndCSMatrix)

   If GelDraw(Ind).CSVisible Then
       'charge state legend
       ShCX = SP(vSelVec, X01, XE1, 3 * LDfSX \ 4 + 200, szLLblLN.cx)
       If bIncludeTextLabels Then
           Res = TextOut(hDC, ShCX, VLblY, sLgndLblLN, Len(sLgndLblLN))
       End If
       ShCX = SP(vSelVec, X01, XE1, LDfSX \ 4 + 100, 0)
       DrawLegendShape hDC, glCSShape, ShCX, ShCY
   End If
   
   If GelDraw(Ind).IsoVisible And GelDraw(Ind).CSVisible Then
       'isotopic legend
       ShCX = SP(vSelVec, X01, XE1, -(3 * LDfSX \ 4 + 200), 0)
       If bIncludeTextLabels Then
           Res = TextOut(hDC, ShCX, VLblY, sLgndLblRN, Len(sLgndLblRN))
       End If
       ShCX = SP(vSelVec, X01, XE1, -(LDfSX \ 4 + 100), 0)
       DrawLegendShape hDC, glIsoShape, ShCX, ShCY
   End If

vSelVec = GetJobVector(Ind, vAxYSelectMatrix, vVYJobMatrix)
ShH = 3 * szVLbl.cy \ 4
ShT = SP(vSelVec, LDfY0, LDfYE, ShH \ 6, ShH)
vSelVec = GetJobVector(Ind, vAxYSelectMatrix, vVYBJobMatrix)
ShB = SP(vSelVec, LDfY0, LDfYE, ShH \ 6, ShH)
ShW = 36
  
Res = SelectObject(hDC, GetStockObject(NULL_PEN))
vSelVec = GetJobVector(Ind, vAxXSelectMatrix, vLgndCSMatrix)
If bIncludeTextLabels Then
    ShCX = SP(vSelVec, LDfXE \ 2, LDfXE \ 2, LDfSX \ 4 + 1800, szLLblLD.cx)
    Res = TextOut(hDC, ShCX, VLblY, sLgndLblLD, Len(sLgndLblLD))
    ShCX = SP(vSelVec, LDfXE \ 2, LDfXE \ 2, -(LDfSX \ 4 + 1800), 0)
    Res = TextOut(hDC, ShCX, VLblY, sLgndLblRD, Len(sLgndLblRD))
End If
Select Case GelData(Ind).Preferences.DRDefinition
Case glNormal
    For i = 0 To 50
       Res = SelectObject(hDC, hDDClrPens(-i))
       Res = SelectObject(hDC, hDDClrBrushes(-i))
       ShCX = SP(vSelVec, LDfXE \ 2, LDfXE \ 2, i * ShW, 0)
       Res = Rectangle(hDC, ShCX - 18, ShT, ShCX + 18, ShB)
       Res = SelectObject(hDC, hDDClrPens(i))
       Res = SelectObject(hDC, hDDClrBrushes(i))
       ShCX = SP(vSelVec, LDfXE \ 2, LDfXE \ 2, -i * ShW, 0)
       Res = Rectangle(hDC, ShCX - 18, ShT, ShCX + 18, ShB)
    Next i
Case glReverse
    For i = 0 To 50
       Res = SelectObject(hDC, hDDClrPens(-i))
       Res = SelectObject(hDC, hDDClrBrushes(-i))
       ShCX = SP(vSelVec, LDfXE \ 2, LDfXE \ 2, -i * ShW, 0)
       Res = Rectangle(hDC, ShCX - 18, ShT, ShCX + 18, ShB)
       Res = SelectObject(hDC, hDDClrPens(i))
       Res = SelectObject(hDC, hDDClrBrushes(i))
       ShCX = SP(vSelVec, LDfXE \ 2, LDfXE \ 2, i * ShW, 0)
       Res = Rectangle(hDC, ShCX - 18, ShT, ShCX + 18, ShB)
    Next i
End Select

Res = SelectObject(hDC, hForeColorPen)
GelDrawVAxisNumbers Ind, hDC
Select Case GelBody(Ind).csMyCooSys.csType
Case glFNCooSys
     GelDrawFNHAxisNumbers Ind, hDC, False
Case glNETCooSys
     GelDrawFNHAxisNumbers Ind, hDC, True
Case glPICooSys
     GelDrawpiHAxisNumbers Ind, hDC
End Select


If bSetFileNameDate Then DrawDateFileName Ind, hDC, HLblY

Res = SelectObject(hDC, OldPen)
Res = SelectObject(hDC, OldBrush)
Res = SelectObject(hDC, lOldFont)
If lNewFont <> 0 Then DeleteObject (lNewFont)
End Sub


Private Sub DrawLegendChargeStateMap(ByVal Ind As Long, ByVal hDC As Long)
'-----------------------------------------------------------------------------
'draws legend and markings on the coordinate axes for charge state map display
'-----------------------------------------------------------------------------
Dim lOldFont As Long, lNewFont As Long
Dim OldBrush As Long
Dim OldPen As Long
Dim Res As Long
Dim ShH As Long, ShW As Long
Dim ShT As Long, ShB As Long
Dim ShCX As Long, ShCY As Long
Dim X01 As Long, XE1 As Long
Dim i As Integer

Dim HLblX As Long, HLblY As Long
Dim VLblX As Long, VLblY As Long
Dim sLgndLblLN As String, sLgndLblRN As String

Dim sHLbl As String, sVLbl As String
Dim szHLbl As Size, szVLbl As Size
Dim szLLblLN As Size, szLLblRN As Size
Dim vSelVec As Variant

Dim CSLbl(5) As String      'labels for charge states
Dim szLLblCS(5) As Size
Dim MaxLblW As Long

Dim iSign As Long

On Error Resume Next
'create new font suitable for writing on the graph
SetGraphFont Ind, hDC, lOldFont, lNewFont

Res = SetROP2(hDC, R2_MASKPEN)
Res = SetBkMode(hDC, TRANSPARENT)

'set labels
sLgndLblLN = "Charge State"
sLgndLblRN = "Isotopic"

Select Case GelBody(Ind).csMyCooSys.csType
Case glPICooSys
     sHLbl = "pI"
Case glFNCooSys
     sHLbl = "Scan number"
Case glNETCooSys
     sHLbl = "NET"
End Select

Select Case GelBody(Ind).csMyCooSys.csYScale
Case glVAxisLin
    sVLbl = "Monoisotopic Mass"
Case glVAxisLog
    sVLbl = "Monoisotopic Mass (Log Scale)"
End Select

'get the sizes of the axes and legend labels
Res = GetTextExtentPoint32(hDC, sHLbl, Len(sHLbl), szHLbl)
Res = GetTextExtentPoint32(hDC, sVLbl, Len(sVLbl), szVLbl)
Res = GetTextExtentPoint32(hDC, sLgndLblLN, Len(sLgndLblLN), szLLblLN)
Res = GetTextExtentPoint32(hDC, sLgndLblRN, Len(sLgndLblRN), szLLblRN)

'draw coordinate axes labels
vSelVec = GetJobVector(Ind, vAxXSelectMatrix, vHXJobMatrix)
HLblX = SP(vSelVec, LDfX0, LDfXE, LDfSX \ 2, szHLbl.cx)
vSelVec = GetJobVector(Ind, vAxXSelectMatrix, vVXJobMatrix)
VLblX = SP(vSelVec, LDfX0, LDfXE, LDfSX \ 2, szVLbl.cx)
vSelVec = GetJobVector(Ind, vAxYSelectMatrix, vHYJobMatrix)
HLblY = SP(vSelVec, LDfY0, LDfYE, lDfLY \ 2, szHLbl.cy)
vSelVec = GetJobVector(Ind, vAxYSelectMatrix, vVYJobMatrix)
VLblY = SP(vSelVec, LDfY0, LDfYE, 0, szVLbl.cy)
If bIncludeTextLabels Then
    Res = TextOut(hDC, HLblX, HLblY, sHLbl, Len(sHLbl))
    Res = TextOut(hDC, VLblX, VLblY, sVLbl, Len(sVLbl))
End If

   'draw legend labels and boxes

   OldBrush = SelectObject(hDC, GetStockObject(NULL_BRUSH)) 'prepare transparent brush
   OldPen = SelectObject(hDC, hForeColorPen)
   X01 = LDfX0 + 1100
   XE1 = LDfXE - 1100

   'draw shape and charge state/isotopic colors legend
   ShH = szVLbl.cy
   vSelVec = GetJobVector(Ind, vAxYSelectMatrix, vVYJobMatrix)
   ShCY = SP(vSelVec, LDfY0, LDfYE, ShH \ 2, 0)
   ShW = 100
   vSelVec = GetJobVector(Ind, vAxXSelectMatrix, vLgndCSMatrix)

   If GelDraw(Ind).CSVisible Then
       'charge state legend
       ShCX = SP(vSelVec, X01, XE1, 3 * LDfSX \ 4 + 200, szLLblLN.cx)
       If bIncludeTextLabels Then
           Res = TextOut(hDC, ShCX, VLblY, sLgndLblLN, Len(sLgndLblLN))
       End If
       ShCX = SP(vSelVec, X01, XE1, LDfSX \ 4 + 100, 0)
       DrawLegendShape hDC, glCSShape, ShCX, ShCY
   End If
    
   If GelDraw(Ind).IsoVisible And GelDraw(Ind).CSVisible Then
       'isotopic legend
       ShCX = SP(vSelVec, X01, XE1, -(3 * LDfSX \ 4 + 200), 0)
       If bIncludeTextLabels Then
           Res = TextOut(hDC, ShCX, VLblY, sLgndLblRN, Len(sLgndLblRN))
       End If
       ShCX = SP(vSelVec, X01, XE1, -(LDfSX \ 4 + 100), 0)
       DrawLegendShape hDC, glIsoShape, ShCX, ShCY
   End If


For i = 0 To 4          'make labels as short as possible
    If glCS1(i + 1) = glCS2(i + 1) Then
       CSLbl(i) = CStr(glCS1(i + 1))
    Else
       CSLbl(i) = CStr(glCS1(i + 1)) & "-" & CStr(glCS2(i + 1))
    End If
Next i
CSLbl(5) = "Other"

For i = 0 To 5         'determine the size of longest label
    Res = GetTextExtentPoint32(hDC, CSLbl(i), Len(CSLbl(i)), szLLblCS(i))
    If Res > MaxLblW Then MaxLblW = Res
Next i

vSelVec = GetJobVector(Ind, vAxYSelectMatrix, vVYJobMatrix)
ShH = 3 * szVLbl.cy \ 4
ShT = SP(vSelVec, LDfY0, LDfYE, ShH \ 6, ShH)
vSelVec = GetJobVector(Ind, vAxYSelectMatrix, vVYBJobMatrix)
ShB = SP(vSelVec, LDfY0, LDfYE, ShH \ 6, ShH)
ShW = ShH
MaxLblW = MaxLblW + ShW + 250
  
If GelBody(Ind).csMyCooSys.csOrigin Mod 2 = 0 Then
   iSign = 1
Else
   iSign = -1
End If

vSelVec = GetJobVector(Ind, vAxYSelectMatrix, vYHNumJobMatrix)
For i = 0 To 5
    Res = SelectObject(hDC, hDDClrBrushes(51 + i))
    ShCX = LDfXE \ 2 - iSign * (i - 4) * MaxLblW
    If bIncludeTextLabels Then
        Res = TextOut(hDC, ShCX, ShT, CSLbl(i), Len(CSLbl(i)))
    End If
    ShCX = ShCX - iSign * (szLLblCS(i).cx + 75)
    Res = Rectangle(hDC, ShCX - ShW \ 4, ShT, ShCX + ShW \ 4, ShB)
Next i

GelDrawVAxisNumbers Ind, hDC
Select Case GelBody(Ind).csMyCooSys.csType
Case glFNCooSys
     GelDrawFNHAxisNumbers Ind, hDC, False
Case glNETCooSys
     GelDrawFNHAxisNumbers Ind, hDC, True
Case glPICooSys
     GelDrawpiHAxisNumbers Ind, hDC
End Select

If bSetFileNameDate Then DrawDateFileName Ind, hDC, HLblY

Res = SelectObject(hDC, OldPen)
Res = SelectObject(hDC, OldBrush)
Res = SelectObject(hDC, lOldFont)
If lNewFont <> 0 Then DeleteObject (lNewFont)
End Sub


Private Sub DrawUniqueMassClasses(ByVal Ind As Long, hDC As Long)
'---------------------------------------------------------------
'draws Unique Mass Classes as rectangles bounding all its points
'---------------------------------------------------------------
Dim i As Long
Dim Res As Long
Dim OldBrush As Long
Dim OldPen As Long
Dim ptDummy As POINTAPI
On Error Resume Next
OldBrush = SelectObject(hDC, GetStockObject(NULL_BRUSH))
OldPen = SelectObject(hDC, hForeColorPen)
With GelUMCDraw(Ind)
    For i = 0 To .Count - 1
        If .ClassID(i) >= 0 Then
           Res = Rectangle(hDC, .X1(i), .Y1(i), .x2(i), .Y2(i))
           'draw also line so we see something when rectangle is not visible
           Res = MoveToEx(hDC, .X1(i), .Y1(i), ptDummy)
           Res = LineTo(hDC, .x2(i), .Y2(i))
        End If
    Next i
End With
Res = SelectObject(hDC, OldBrush)
Res = SelectObject(hDC, OldPen)
End Sub


