VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CooSysR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'coordinate system for the GelBody
'last modified 03/26/2003 nt
'-------------------------------------------------------------------------------------
Option Explicit

Const MAXZOOMDEPTH = 100

Public CSIndex As Long         'connection between coosys and GelBody and GelData
Public csZmLvl As Integer      'up to 100 zoom levels
Public BuildingCS As Boolean   'true only when coo sys is initializing or reinitializing

Private mcsType As Integer       'pI or FN type
Private mcsOriginXY As Integer   'XY axes intersection position
Private mcsXOrient As Integer    'horizontal orientation
Private mcsYOrient As Integer    'vertical orientation
Private mcsOrigin As Integer     '(XMin,YMin) position - read only
Private mcsYScale As Integer     'y scale type (linear or logarithmic)

Private mcsXStep As Long         'integer step for the x axis
Private mcsXNumOfSteps As Long   'number of steps for selected XStep

Private mcsScaleXLP As Double    'scale to calculate logical coordinates from real
Private mcsScaleYLP As Double

Private mcsScaleXRP As Double    'scale to calculate real coordinates from logical
Private mcsScaleYRP As Double

Public VX0 As Long               'viewport-physical device coordinates
Public VXE As Long
Public VY0 As Long
Public VYE As Long

Public XYX0 As Double    'intersection of the XY axes
Public XYXE As Double    'point coordinates
Public XYY0 As Double
Public XYYE As Double

Public OrX0 As Double    'origin (as defined above)
Public OrXE As Double    'coordinates
Public OrY0 As Double
Public OrYE As Double

Private RX1(101) As Long     'real coordinates stack - for x axis its index in the .ScanInfo array (discrete)
Private RX2(101) As Long     'RX1(i)<RX2(i)
Private RY1(101) As Double   '
Private RY2(101) As Double   'RY1(i)<RY2(i)

Private RX1Prev As Long      'used to go step back after coordinate system reset
Private RX2Prev As Long
Private RY1Prev As Double
Private RY2Prev As Double


Public RZ1 As Double
Public RZ2 As Double

Private mcsMatrix(1 To 4, 1 To 4) As Long 'read only as variant array

Public SXOffset As Double  'small and large offsets
Public LXOffset As Double  'for the current coo sys
Public SYOffset As Double
Public LYOffset As Double

Public Event CooSysChanged()
Public Event MWScaleChange()

Public Sub ZoomIn(ByVal lX1 As Long, ByVal lY1 As Long, ByVal lX2 As Long, ByVal lY2 As Long)
'input - logical coordinates of the box to zoom-in
Dim ReX1 As Double, ReX2 As Double
Dim lReX1 As Long, lReX2 As Long
Dim ReY1 As Double, ReY2 As Double

If lX1 < lX2 Then
   ReX1 = CDbl(lX1)
   ReX2 = CDbl(lX2)
Else
   ReX1 = CDbl(lX2)
   ReX2 = CDbl(lX1)
End If
If lY1 < lY2 Then
   ReY1 = CDbl(lY1)
   ReY2 = CDbl(lY2)
Else
   ReY1 = CDbl(lY2)
   ReY2 = CDbl(lY1)
End If

LPToRP ReX1, ReY1, ReX2, ReY2

lReX1 = CLng(ReX1)
lReX2 = CLng(ReX2)
If lReX1 >= lReX2 Then
   lReX2 = lReX1 + 1
   lReX1 = lReX1 - 1
End If
If ReY1 >= ReY2 Then
   ReY2 = ReY1 + 1
   ReY1 = ReY1 - 1
End If

ZoomInStorePrevious lReX1, ReY1, lReX2, ReY2

mcsXStepPicker
CoordinateDraw
RaiseEvent CooSysChanged
End Sub

Public Sub ZoomInR(ByVal lX1 As Long, ByVal dY1 As Double, ByVal lX2 As Long, ByVal dY2 As Double)

ZoomInStorePrevious lX1, dY1, lX2, dY2

mcsXStepPicker
CoordinateDraw
RaiseEvent CooSysChanged
End Sub

Private Sub ZoomInStorePrevious(ByVal lX1 As Long, ByVal dY1 As Double, ByVal lX2 As Long, ByVal dY2 As Double)

Dim i As Long

If csZmLvl < MAXZOOMDEPTH Then
   csZmLvl = csZmLvl + 1
Else
    ' Shift all of the zoom level values down one location
    For i = 2 To MAXZOOMDEPTH
        RX1(i) = RX1(i + 1)
        RX2(i) = RX2(i + 1)
        RY1(i) = RY1(i + 1)
        RY2(i) = RY2(i + 1)
    Next i
End If
RX1(csZmLvl) = lX1
RX2(csZmLvl) = lX2
RY1(csZmLvl) = dY1
RY2(csZmLvl) = dY2

RX1Prev = RX1(csZmLvl)
RX2Prev = RX2(csZmLvl)
RY1Prev = RY1(csZmLvl)
RY2Prev = RY2(csZmLvl)

End Sub

Public Sub ZoomOut()
csZmLvl = 1
mcsXStepPicker
CoordinateDraw
RaiseEvent CooSysChanged
End Sub

Public Sub ZoomOut1()
If csZmLvl > 1 Then
   csZmLvl = csZmLvl - 1
   mcsXStepPicker
   CoordinateDraw
   RaiseEvent CooSysChanged
End If
End Sub


Public Sub ZoomPrevious()
'----------------------------------------------------------------------------
'this is ZoomOut1 if there is anything too-zoom out and zoom-in to remembered
'----------------------------------------------------------------------------
If csZmLvl > 1 Then
   Call ZoomOut1
Else
   Call ZoomInR(RX1Prev, RY1Prev, RX2Prev, RY2Prev)
End If
End Sub


Public Sub CoordinateDraw()
Dim IsoF As Integer     'isotopic mw field (just a shortcut)
Dim ScaleAbu As Double
Dim FNSpan As Double
Dim lMaxLogAbu As Long
Dim lMinLogAbu As Long
Dim LogMW As Double
Dim LogMWMin As Double
Dim LogMWMax As Double
Dim UMCMWAbsTol As Double
Dim UMCMWAbu As Long    'size of class used as width of the UMC box
Dim i As Long

On Error GoTo exit_CoordinateDraw

mcsScaleXLP = (LDfX2 - LDfX1) / (RX2(csZmLvl) - RX1(csZmLvl))
mcsScaleYLP = (LDfY2 - LDfY1) / (RY2(csZmLvl) - RY1(csZmLvl))

With GelData(CSIndex)
    'if auto sizing adjust points
    If glbPreferencesExpanded.AutoAdjSize Then
        If glbPreferencesExpanded.AutoSizeMultiplier <= 0 Then glbPreferencesExpanded.AutoSizeMultiplier = 1
       .Preferences.MaxPointFactor = (RX2(csZmLvl) - RX1(csZmLvl)) / 100 * glbPreferencesExpanded.AutoSizeMultiplier
       If .Preferences.MaxPointFactor < 1 Then .Preferences.MaxPointFactor = 1
       .Preferences.MinPointFactor = .Preferences.MaxPointFactor / 4
       If .Preferences.MinPointFactor < 0.5 Then .Preferences.MinPointFactor = 0.5
    End If
    
    lMinLogAbu = mcsScaleXLP * .Preferences.MinPointFactor
    lMaxLogAbu = mcsScaleXLP * .Preferences.MaxPointFactor
    IsoF = .Preferences.IsoDataField
    If RZ1 <> RZ2 Then
       ScaleAbu = (lMaxLogAbu - lMinLogAbu) / (RZ2 - RZ1)
    Else
       ScaleAbu = 0
    End If
    If csYScale = glVAxisLin Then
       If .CSLines > 0 Then
          For i = 1 To .CSLines
            If .CSData(i).ScanNumber < RX1(csZmLvl) Or .CSData(i).ScanNumber > RX2(csZmLvl) Or _
               .CSData(i).AverageMW < RY1(csZmLvl) Or .CSData(i).AverageMW > RY2(csZmLvl) Then
               GelDraw(CSIndex).CSX(i) = CLngNoOverflow(LDfX1 + mcsScaleXLP * (.CSData(i).ScanNumber - RX1(csZmLvl))) 'FN
               GelDraw(CSIndex).CSY(i) = CLngNoOverflow(LDfY1 + mcsScaleYLP * (.CSData(i).AverageMW - RY1(csZmLvl))) 'MW
               GelDraw(CSIndex).CSR(i) = -1     'out of scope; don't draw
'               GelDraw(CSIndex).CSR(i) = CLngNoOverflow((.CSData(i).Abundance - RZ1) * ScaleAbu + lMinLogAbu)   'Intensity
            Else
               GelDraw(CSIndex).CSX(i) = CLngNoOverflow(LDfX1 + mcsScaleXLP * (.CSData(i).ScanNumber - RX1(csZmLvl))) 'FN
               GelDraw(CSIndex).CSY(i) = CLngNoOverflow(LDfY1 + mcsScaleYLP * (.CSData(i).AverageMW - RY1(csZmLvl))) 'MW
               GelDraw(CSIndex).CSR(i) = CLngNoOverflow((.CSData(i).Abundance - RZ1) * ScaleAbu + lMinLogAbu)   'Intensity
            End If
          Next i
       End If
       If .IsoLines > 0 Then
          For i = 1 To .IsoLines
            If .IsoData(i).ScanNumber < RX1(csZmLvl) Or .IsoData(i).ScanNumber > RX2(csZmLvl) Or _
               GetIsoMass(.IsoData(i), IsoF) < RY1(csZmLvl) Or GetIsoMass(.IsoData(i), IsoF) > RY2(csZmLvl) Then
               GelDraw(CSIndex).IsoX(i) = CLngNoOverflow(LDfX1 + mcsScaleXLP * (.IsoData(i).ScanNumber - RX1(csZmLvl)))
               GelDraw(CSIndex).IsoY(i) = CLngNoOverflow(LDfY1 + mcsScaleYLP * (GetIsoMass(.IsoData(i), IsoF) - RY1(csZmLvl)))
               GelDraw(CSIndex).IsoR(i) = -1     'out of scope; don't draw
'               GelDraw(CSIndex).IsoR(i) = CLngNoOverflow((.IsoData(i).Abundance - RZ1) * ScaleAbu + lMinLogAbu)
            Else
               GelDraw(CSIndex).IsoX(i) = CLngNoOverflow(LDfX1 + mcsScaleXLP * (.IsoData(i).ScanNumber - RX1(csZmLvl)))
               GelDraw(CSIndex).IsoY(i) = CLngNoOverflow(LDfY1 + mcsScaleYLP * (GetIsoMass(.IsoData(i), IsoF) - RY1(csZmLvl)))
               GelDraw(CSIndex).IsoR(i) = CLngNoOverflow((.IsoData(i).Abundance - RZ1) * ScaleAbu + lMinLogAbu)
            End If
          Next i
       End If
    Else
       If .CSLines > 0 Then
          For i = 1 To .CSLines
            LogMW = GelDraw(CSIndex).CSLogMW(i)
            If .CSData(i).ScanNumber < RX1(csZmLvl) Or .CSData(i).ScanNumber > RX2(csZmLvl) Or _
               LogMW < RY1(csZmLvl) Or LogMW > RY2(csZmLvl) Or LogMW = -glHugeOverExp Then
               GelDraw(CSIndex).CSX(i) = CLngNoOverflow(LDfX1 + mcsScaleXLP * (.CSData(i).ScanNumber - RX1(csZmLvl)))
               GelDraw(CSIndex).CSY(i) = CLngNoOverflow(LDfY1 + mcsScaleYLP * (LogMW - RY1(csZmLvl)))
               GelDraw(CSIndex).CSR(i) = -1     'out of scope; don't draw
'               GelDraw(CSIndex).CSR(i) = CLngNoOverflow((.CSData(i).Abundance - RZ1) * ScaleAbu + lMinLogAbu)
            Else
               GelDraw(CSIndex).CSX(i) = CLngNoOverflow(LDfX1 + mcsScaleXLP * (.CSData(i).ScanNumber - RX1(csZmLvl)))
               GelDraw(CSIndex).CSY(i) = CLngNoOverflow(LDfY1 + mcsScaleYLP * (LogMW - RY1(csZmLvl)))
               GelDraw(CSIndex).CSR(i) = CLngNoOverflow((.CSData(i).Abundance - RZ1) * ScaleAbu + lMinLogAbu)
            End If
          Next i
       End If
       If .IsoLines > 0 Then
          For i = 1 To .IsoLines
            LogMW = GelDraw(CSIndex).IsoLogMW(i)
            If .IsoData(i).ScanNumber < RX1(csZmLvl) Or .IsoData(i).ScanNumber > RX2(csZmLvl) Or _
               LogMW < RY1(csZmLvl) Or LogMW > RY2(csZmLvl) Or LogMW = -glHugeOverExp Then
               GelDraw(CSIndex).IsoX(i) = CLngNoOverflow(LDfX1 + mcsScaleXLP * (.IsoData(i).ScanNumber - RX1(csZmLvl)))
               GelDraw(CSIndex).IsoY(i) = CLngNoOverflow(LDfY1 + mcsScaleYLP * (LogMW - RY1(csZmLvl)))
               GelDraw(CSIndex).IsoR(i) = -1     'out of scope; don't draw
'               GelDraw(CSIndex).IsoR(i) = CLngNoOverflow((.IsoData(i).Abundance - RZ1) * ScaleAbu + lMinLogAbu)
            Else
               GelDraw(CSIndex).IsoX(i) = CLngNoOverflow(LDfX1 + mcsScaleXLP * (.IsoData(i).ScanNumber - RX1(csZmLvl)))
               GelDraw(CSIndex).IsoY(i) = CLngNoOverflow(LDfY1 + mcsScaleYLP * (LogMW - RY1(csZmLvl)))
               GelDraw(CSIndex).IsoR(i) = CLngNoOverflow((.IsoData(i).Abundance - RZ1) * ScaleAbu + lMinLogAbu)
            End If
          Next
       End If
    End If
End With
'if Unique Mass Classes are visible we need to recalculate DrawUMC structure
'class is drawn if any of points making class rectangle fits in current scope
If GelUMCDraw(CSIndex).Visible Then
   With GelUMC(CSIndex)
      If .UMCCnt > 0 Then
         Select Case GelUMCDraw(CSIndex).DrawType
         Case umcdt_ActualUMC
            If csYScale = glVAxisLin Then
               For i = 0 To .UMCCnt - 1
                 If (IsInScope(.UMCs(i).MinScan, .UMCs(i).MinMW) Or IsInScope(.UMCs(i).MinScan, .UMCs(i).MaxMW) _
                   Or IsInScope(.UMCs(i).MaxScan, .UMCs(i).MinMW) Or IsInScope(.UMCs(i).MaxScan, .UMCs(i).MaxMW)) Then
                   GelUMCDraw(CSIndex).ClassID(i) = Abs(GelUMCDraw(CSIndex).ClassID(i))
                   GelUMCDraw(CSIndex).X1(i) = CLngNoOverflow(LDfX1 + mcsScaleXLP * (.UMCs(i).MinScan - RX1(csZmLvl)))
                   GelUMCDraw(CSIndex).x2(i) = CLngNoOverflow(LDfX1 + mcsScaleXLP * (.UMCs(i).MaxScan - RX1(csZmLvl)))
                   GelUMCDraw(CSIndex).Y1(i) = CLngNoOverflow(LDfY1 + mcsScaleYLP * (.UMCs(i).MinMW - RY1(csZmLvl)))
                   GelUMCDraw(CSIndex).Y2(i) = CLngNoOverflow(LDfY1 + mcsScaleYLP * (.UMCs(i).MaxMW - RY1(csZmLvl)))
                 Else
                   GelUMCDraw(CSIndex).ClassID(i) = -Abs(GelUMCDraw(CSIndex).ClassID(i))
                 End If
               Next i
            Else
               For i = 0 To .UMCCnt - 1
                 LogMWMin = Log(.UMCs(i).MinMW) / Log(10#)
                 LogMWMax = Log(.UMCs(i).MaxMW) / Log(10#)
                 If (IsInScope(.UMCs(i).MinScan, LogMWMin) Or IsInScope(.UMCs(i).MinScan, LogMWMax) Or _
                   IsInScope(.UMCs(i).MaxScan, LogMWMin) Or IsInScope(.UMCs(i).MaxScan, LogMWMax)) Then
                   GelUMCDraw(CSIndex).ClassID(i) = Abs(GelUMCDraw(CSIndex).ClassID(i))
                   GelUMCDraw(CSIndex).X1(i) = CLngNoOverflow(LDfX1 + mcsScaleXLP * (.UMCs(i).MinScan - RX1(csZmLvl)))
                   GelUMCDraw(CSIndex).x2(i) = CLngNoOverflow(LDfX1 + mcsScaleXLP * (.UMCs(i).MaxScan - RX1(csZmLvl)))
                   GelUMCDraw(CSIndex).Y1(i) = CLngNoOverflow(LDfY1 + mcsScaleYLP * (LogMWMin - RY1(csZmLvl)))
                   GelUMCDraw(CSIndex).Y2(i) = CLngNoOverflow(LDfY1 + mcsScaleYLP * (LogMWMax - RY1(csZmLvl)))
                 Else
                   GelUMCDraw(CSIndex).ClassID(i) = -Abs(GelUMCDraw(CSIndex).ClassID(i))
                 End If
               Next i
            End If
         Case umcdt_FullUMCRegion
            If csYScale = glVAxisLin Then
               For i = 0 To .UMCCnt - 1
                 Select Case .def.TolType
                 Case gltPPM
                      UMCMWAbsTol = .UMCs(i).ClassMW * .def.Tol * glPPM
                 Case gltABS
                      UMCMWAbsTol = .def.Tol
                 End Select
                 If (IsInScope(.UMCs(i).MinScan, .UMCs(i).MinMW) Or IsInScope(.UMCs(i).MinScan, .UMCs(i).MaxMW) _
                   Or IsInScope(.UMCs(i).MaxScan, .UMCs(i).MinMW) Or IsInScope(.UMCs(i).MaxScan, .UMCs(i).MaxMW)) Then
                   GelUMCDraw(CSIndex).ClassID(i) = Abs(GelUMCDraw(CSIndex).ClassID(i))
                   GelUMCDraw(CSIndex).X1(i) = CLngNoOverflow(LDfX1 + mcsScaleXLP * (.UMCs(i).MinScan - RX1(csZmLvl)))
                   GelUMCDraw(CSIndex).x2(i) = CLngNoOverflow(LDfX1 + mcsScaleXLP * (.UMCs(i).MaxScan - RX1(csZmLvl)))
                   GelUMCDraw(CSIndex).Y1(i) = CLngNoOverflow(LDfY1 + mcsScaleYLP * (.UMCs(i).ClassMW - UMCMWAbsTol - RY1(csZmLvl)))
                   GelUMCDraw(CSIndex).Y2(i) = CLngNoOverflow(LDfY1 + mcsScaleYLP * (.UMCs(i).ClassMW + UMCMWAbsTol - RY1(csZmLvl)))
                 Else
                   GelUMCDraw(CSIndex).ClassID(i) = -Abs(GelUMCDraw(CSIndex).ClassID(i))
                 End If
               Next i
            Else
               For i = 0 To .UMCCnt - 1
'                 Select Case .def.TolType                      'test with log of size as well
'                 Case gltPPM
'                      UMCMWAbsTol = Log(.UMCs(i).ClassMW * .def.Tol * glPPM) / Log(10#)
'                 Case gltABS
'                      UMCMWAbsTol = Log(.def.Tol) / Log(10#)
'                 End Select
                 LogMW = .UMCs(i).ClassMW
                 Select Case .def.TolType
                 Case gltPPM
                      UMCMWAbsTol = .UMCs(i).ClassMW * .def.Tol * glPPM
                 Case gltABS
                      UMCMWAbsTol = .def.Tol
                 End Select
                 If (IsInScope(.UMCs(i).MinScan, LogMWMin) Or IsInScope(.UMCs(i).MinScan, LogMWMax) Or _
                   IsInScope(.UMCs(i).MaxScan, LogMWMin) Or IsInScope(.UMCs(i).MaxScan, LogMWMax)) Then
                   GelUMCDraw(CSIndex).ClassID(i) = Abs(GelUMCDraw(CSIndex).ClassID(i))
                   GelUMCDraw(CSIndex).X1(i) = CLngNoOverflow(LDfX1 + mcsScaleXLP * (.UMCs(i).MinScan - RX1(csZmLvl)))
                   GelUMCDraw(CSIndex).x2(i) = CLngNoOverflow(LDfX1 + mcsScaleXLP * (.UMCs(i).MaxScan - RX1(csZmLvl)))
                   GelUMCDraw(CSIndex).Y1(i) = CLngNoOverflow(LDfY1 + mcsScaleYLP * (LogMW - UMCMWAbsTol - RY1(csZmLvl)))
                   GelUMCDraw(CSIndex).Y2(i) = CLngNoOverflow(LDfY1 + mcsScaleYLP * (LogMW + UMCMWAbsTol - RY1(csZmLvl)))
                 Else
                   GelUMCDraw(CSIndex).ClassID(i) = -Abs(GelUMCDraw(CSIndex).ClassID(i))
                 End If
               Next i
            End If
         Case umcdt_UMCIntensity
            If csYScale = glVAxisLin Then
               For i = 0 To .UMCCnt - 1
                 If (IsInScope(.UMCs(i).MinScan, .UMCs(i).MinMW) Or IsInScope(.UMCs(i).MinScan, .UMCs(i).MaxMW) _
                   Or IsInScope(.UMCs(i).MaxScan, .UMCs(i).MinMW) Or IsInScope(.UMCs(i).MaxScan, .UMCs(i).MaxMW)) Then
                   GelUMCDraw(CSIndex).ClassID(i) = Abs(GelUMCDraw(CSIndex).ClassID(i))
                   UMCMWAbu = CLngNoOverflow(((.UMCs(i).ClassAbundance - RZ1) * ScaleAbu + lMinLogAbu) / 2)
                   GelUMCDraw(CSIndex).X1(i) = CLngNoOverflow(LDfX1 + mcsScaleXLP * (.UMCs(i).MinScan - RX1(csZmLvl)))
                   GelUMCDraw(CSIndex).x2(i) = CLngNoOverflow(LDfX1 + mcsScaleXLP * (.UMCs(i).MaxScan - RX1(csZmLvl)))
                   GelUMCDraw(CSIndex).Y1(i) = CLngNoOverflow(LDfY1 + mcsScaleYLP * (.UMCs(i).ClassMW - RY1(csZmLvl)) - UMCMWAbu)
                   GelUMCDraw(CSIndex).Y2(i) = CLngNoOverflow(LDfY1 + mcsScaleYLP * (.UMCs(i).ClassMW - RY1(csZmLvl)) + UMCMWAbu)
                 Else
                   GelUMCDraw(CSIndex).ClassID(i) = -Abs(GelUMCDraw(CSIndex).ClassID(i))
                 End If
               Next i
            Else
               For i = 0 To .UMCCnt - 1
                 LogMW = Log(.UMCs(i).ClassMW) / Log(10#)
                 If (IsInScope(.UMCs(i).MinScan, LogMWMin) Or IsInScope(.UMCs(i).MinScan, LogMWMax) Or _
                   IsInScope(.UMCs(i).MaxScan, LogMWMin) Or IsInScope(.UMCs(i).MaxScan, LogMWMax)) Then
                   GelUMCDraw(CSIndex).ClassID(i) = Abs(GelUMCDraw(CSIndex).ClassID(i))
                   UMCMWAbu = CLngNoOverflow(((.UMCs(i).ClassAbundance - RZ1) * ScaleAbu + lMinLogAbu) / 2)
                   GelUMCDraw(CSIndex).X1(i) = CLngNoOverflow(LDfX1 + mcsScaleXLP * (.UMCs(i).MinScan - RX1(csZmLvl)))
                   GelUMCDraw(CSIndex).x2(i) = CLngNoOverflow(LDfX1 + mcsScaleXLP * (.UMCs(i).MaxScan - RX1(csZmLvl)))
                   GelUMCDraw(CSIndex).Y1(i) = CLngNoOverflow(LDfY1 + mcsScaleYLP * (LogMW - RY1(csZmLvl)) - UMCMWAbu)
                   GelUMCDraw(CSIndex).Y2(i) = CLngNoOverflow(LDfY1 + mcsScaleYLP * (LogMW - RY1(csZmLvl)) + UMCMWAbu)
                 Else
                   GelUMCDraw(CSIndex).ClassID(i) = -Abs(GelUMCDraw(CSIndex).ClassID(i))
                 End If
               Next i
            End If
         End Select
      End If
   End With
End If
Exit Sub

exit_CoordinateDraw:
Debug.Assert False
Debug.Print "Error in CooSysR.CoordinateDraw: " & Err.Description
LogErrors Err.Number, "CooSysR->CoordinateDraw"

End Sub

Private Static Function CLngNoOverflow(dblValue As Double) As Long
    On Error GoTo OutOfRange
    
    CLngNoOverflow = CLng(dblValue)
    Exit Function
    
OutOfRange:
    If dblValue > 0 Then
        CLngNoOverflow = 2147483647
    Else
        CLngNoOverflow = -2147483647
    End If

End Function

Public Sub InitFNType()
'initializes coordinate system of the FN type
'should not come to here if .ScanInfo does not have at least 1 element
On Error GoTo InitFNTypeErrorHandler
With GelData(CSIndex)
    RX1(0) = .ScanInfo(1).ScanNumber
    RX2(0) = .ScanInfo(UBound(.ScanInfo)).ScanNumber
    Select Case mcsYScale
    Case glVAxisLin
         RY1(0) = .MinMW
         RY2(0) = .MaxMW
    Case glVAxisLog
         If .MinMW > 0 Then
            RY1(0) = CDbl(Log(.MinMW) / Log(10#))
         Else
            RY1(0) = -glHugeOverExp
         End If
         If .MaxMW > 0 Then
            RY2(0) = CDbl(Log(.MaxMW) / Log(10#))
         Else
            RY2(0) = -glHugeOverExp
         End If
    End Select
    RX1(1) = RX1(0)
    RX2(1) = RX2(0)
    If RX2(1) <= RX1(1) Then
       RX2(1) = RX1(1) + 1
       RX1(1) = RX1(1) - 1
    End If
    RY1(1) = RY1(0)
    RY2(1) = RY2(0)
    If RY2(1) <= RY1(1) Then
       RY2(1) = RY1(1) + 1
       RY1(1) = RY1(1) - 1
    End If
    RZ1 = .MinAbu
    RZ2 = .MaxAbu
End With
csZmLvl = 1
mcsXStepPicker

Exit Sub
InitFNTypeErrorHandler:
Debug.Print "Error in CooSysR->InitFNType: " & Err.Description
LogErrors Err.Number, "CooSysR->InitFNType"
End Sub

Public Property Get csType() As Integer
csType = mcsType
End Property

Public Property Let csType(ByVal iNewType As Integer)
mcsType = iNewType
If Not BuildingCS Then
   mcsXStepPicker
   CoordinateDraw
   RaiseEvent CooSysChanged
End If
End Property

Public Sub LPToRP(X1 As Double, Y1 As Double, x2 As Double, Y2 As Double)
'converts logical coordinates to the real coordinates
mcsScaleYRP = (RY2(csZmLvl) - RY1(csZmLvl)) / (LDfY2 - LDfY1)
mcsScaleXRP = (RX2(csZmLvl) - RX1(csZmLvl)) / (LDfX2 - LDfX1)

X1 = mcsScaleXRP * (X1 - LDfX1) + RX1(csZmLvl)
x2 = mcsScaleXRP * (x2 - LDfX1) + RX1(csZmLvl)
Y1 = mcsScaleYRP * (Y1 - LDfY1) + RY1(csZmLvl)
Y2 = mcsScaleYRP * (Y2 - LDfY1) + RY1(csZmLvl)
End Sub

Public Property Get csOriginXY() As Integer
csOriginXY = mcsOriginXY
End Property

Public Property Let csOriginXY(ByVal iOXY As Integer)
mcsOriginXY = iOXY
If mcsXOrient > 0 And mcsYOrient > 0 Then
   mcsOrigin = mcsXOrient + 2 * (mcsYOrient - 1)
End If
End Property

Public Property Get csXOrient() As Integer
csXOrient = mcsXOrient
End Property

Public Property Let csXOrient(ByVal iXOrient As Integer)
mcsXOrient = iXOrient
If mcsYOrient > 0 Then
   mcsOrigin = mcsXOrient + 2 * (mcsYOrient - 1)
End If
End Property

Public Property Get csYOrient() As Integer
csYOrient = mcsYOrient
End Property

Public Property Let csYOrient(ByVal iYOrient As Integer)
mcsYOrient = iYOrient
If mcsXOrient > 0 Then
   mcsOrigin = mcsXOrient + 2 * (mcsYOrient - 1)
End If
End Property

Public Property Get csOrigin() As Integer
csOrigin = mcsOrigin
End Property

Public Property Get CurrRXMin() As Double
CurrRXMin = RX1(csZmLvl)
End Property

Public Property Get CurrRXMax() As Double
CurrRXMax = RX2(csZmLvl)
End Property

Public Property Get CurrRYMin() As Double
CurrRYMin = RY1(csZmLvl)
End Property

Public Property Get CurrRYMax() As Double
CurrRYMax = RY2(csZmLvl)
End Property

Private Sub mcsXStepPicker()
'--------------------------------------------------------------
'ExtDir defines direction in which coordinate system will be
'adjusted to achieve harmony with horizontal scaling and nature
'--------------------------------------------------------------
Dim ExtL As Long
Dim ExtR As Long
Dim bDone As Boolean
Dim xSpan As Long
ExtL = 0
ExtR = 0
Do While Not bDone
   RX1(csZmLvl) = RX1(csZmLvl) - ExtL
   If RX1(csZmLvl) < 0 Then RX1(csZmLvl) = 0
   RX2(csZmLvl) = RX2(csZmLvl) + ExtR
   xSpan = RX2(csZmLvl) - RX1(csZmLvl)
   If xSpan < 20 Then
      mcsXStep = 1
      mcsXNumOfSteps = xSpan
      bDone = True
   Else
      mcsXNumOfSteps = FactorN(xSpan, 15)
      If mcsXNumOfSteps < 10 Or mcsXNumOfSteps > 19 Then
         If ExtR = 0 Then
            ExtL = 0
            ExtR = 1
         Else
            ExtL = 1
            ExtR = 0
         End If
      Else
         mcsXStep = xSpan \ mcsXNumOfSteps
         bDone = True
      End If
   End If
Loop
End Sub


Private Sub mcsXStepPickerEnforcer(NewRX1 As Long, NewRX2 As Long, _
                                   ByVal ExtDir As Long, ByVal ExtMode As Long)
'-------------------------------------------------------------------------------
'this function is necessary to allow aproximately correct behavior of horizontal
'fine tuning
'ExtDir=0,1 Left,Right; ExtMode=-1 or 1 for shrink and expand(depending on side)
'adjustment is calculated so that actual StepPicker will have nothing to adjust
'-------------------------------------------------------------------------------
Dim bDone As Boolean
Dim xSpan As Long
Dim xStepsCnt As Long
Select Case ExtDir
Case glTuneLT
     Do While Not bDone
        NewRX1 = NewRX1 + ExtMode
        If NewRX1 < 0 Then NewRX1 = 0
        xSpan = NewRX2 - NewRX1
        If xSpan < 20 Then
           bDone = True
        Else
           xStepsCnt = FactorN(xSpan, 15)
           If Not (xStepsCnt < 10 Or xStepsCnt > 19) Then bDone = True
        End If
     Loop
Case glTuneRB
     Do While Not bDone
        NewRX2 = NewRX2 + ExtMode
        xSpan = NewRX2 - NewRX1
        If xSpan < 20 Then
           bDone = True
        Else
           xStepsCnt = FactorN(xSpan, 15)
           If Not (xStepsCnt < 10 Or xStepsCnt > 19) Then bDone = True
        End If
     Loop
End Select
End Sub

Public Sub GetVisibleDimensions(ByRef FNMin As Long, ByRef FNMax As Long, ByRef MassMin As Double, ByRef MassMax As Double)

    FNMin = Me.CurrRXMin
    FNMax = Me.CurrRXMax
    
    If Me.csYScale = glVAxisLin Then
       MassMin = Round(Me.CurrRYMin, 6)
       MassMax = Round(Me.CurrRYMax, 6)
    Else
       MassMin = Round(10 ^ Me.CurrRYMin, 6)
       MassMax = Round(10 ^ Me.CurrRYMax, 6)
    End If

End Sub


Public Sub GetViewPortRectangle(ByVal paX As Long, ByVal paY As Long, rcT As Long, rcL As Long, rcB As Long, rcR As Long)
If VXE < 0 Then
   rcL = paX + VX0 + VXE
   rcR = paX + VX0
Else
   rcL = paX + VX0
   rcR = paX + VX0 + VXE
End If
If VYE < 0 Then
   rcT = paY + VY0 + VYE
   rcB = paY + VY0
Else
   rcT = paY + VY0
   rcB = paY + VY0 + VYE
End If
End Sub

Public Property Get csXStep() As Long
    csXStep = mcsXStep
End Property

Public Sub FineTuneH(ByVal TuneDirection As Integer, ByVal TuneSide As Integer)
Dim NewR1 As Long, NewR2 As Long
Select Case TuneSide
Case glTuneLT
     NewR1 = RX1(csZmLvl)
     NewR2 = RX2(csZmLvl)
     Select Case TuneDirection
     Case glExpand
        mcsXStepPickerEnforcer NewR1, NewR2, glTuneLT, -1
        If NewR1 >= RX1(0) Then ZoomInR NewR1, RY1(csZmLvl), NewR2, RY2(csZmLvl)
     Case glShrink
        mcsXStepPickerEnforcer NewR1, NewR2, glTuneLT, 1
        If NewR1 < NewR2 Then ZoomInR NewR1, RY1(csZmLvl), NewR2, RY2(csZmLvl)
     End Select
Case glTuneRB
     NewR1 = RX1(csZmLvl)
     NewR2 = RX2(csZmLvl)
     Select Case TuneDirection
     Case glExpand
        mcsXStepPickerEnforcer NewR1, NewR2, glTuneRB, 1
        If NewR2 <= RX2(0) Then ZoomInR NewR1, RY1(csZmLvl), NewR2, RY2(csZmLvl)
     Case glShrink
        mcsXStepPickerEnforcer NewR1, NewR2, glTuneRB, -1
        If NewR1 < NewR2 Then ZoomInR NewR1, RY1(csZmLvl), NewR2, RY2(csZmLvl)
     End Select
Case glTuneMoveL
     NewR1 = RX1(csZmLvl) - 1
     NewR2 = RX2(csZmLvl) - 1
     If NewR1 >= RX1(0) Then ZoomInR NewR1, RY1(csZmLvl), NewR2, RY2(csZmLvl)
Case glTuneMoveR
     NewR1 = RX1(csZmLvl) + 1
     NewR2 = RX2(csZmLvl) + 1
     If NewR1 <= RX2(0) Then ZoomInR NewR1, RY1(csZmLvl), NewR2, RY2(csZmLvl)
End Select
End Sub

Public Sub FineTuneV(ByVal TuneDirection As Integer, ByVal TuneSide As Integer)
Dim NewR1 As Double, NewR2 As Double
Dim VShift As Double
VShift = (RY2(csZmLvl) - RY1(csZmLvl)) / 10
Select Case TuneSide
Case glTuneLT
     Select Case TuneDirection
     Case glExpand
          NewR1 = RY1(csZmLvl) - VShift
          If NewR1 > RY1(0) Then
             ZoomInR RX1(csZmLvl), NewR1, RX2(csZmLvl), RY2(csZmLvl)
          Else
             ZoomInR RX1(csZmLvl), RY1(0), RX2(csZmLvl), RY2(csZmLvl)
          End If
     Case glShrink
          NewR1 = RY1(csZmLvl) + VShift
          If NewR1 < RY2(csZmLvl) Then ZoomInR RX1(csZmLvl), NewR1, RX2(csZmLvl), RY2(csZmLvl)
     End Select
Case glTuneRB
     Select Case TuneDirection
     Case glExpand
          NewR2 = RY2(csZmLvl) + VShift
          If NewR2 < RY2(0) Then
             ZoomInR RX1(csZmLvl), RY1(csZmLvl), RX2(csZmLvl), NewR2
          Else
             ZoomInR RX1(csZmLvl), RY1(csZmLvl), RX2(csZmLvl), RY2(0)
          End If
     Case glShrink
          NewR2 = RY2(csZmLvl) - VShift
          If NewR2 > RY1(csZmLvl) Then ZoomInR RX1(csZmLvl), RY1(csZmLvl), RX2(csZmLvl), NewR2
     End Select
Case glTuneMoveU
     NewR1 = RY1(csZmLvl) + VShift
     NewR2 = RY2(csZmLvl) + VShift
     If NewR2 < RY2(0) Then ZoomInR RX1(csZmLvl), NewR1, RX2(csZmLvl), NewR2
Case glTuneMoveD
     NewR1 = RY1(csZmLvl) - VShift
     NewR2 = RY2(csZmLvl) - VShift
     If NewR1 >= RY1(0) Then ZoomInR RX1(csZmLvl), NewR1, RX2(csZmLvl), NewR2
End Select
End Sub

Public Property Get csYScale() As Integer
csYScale = mcsYScale
End Property

Public Property Let csYScale(ByVal iNewValue As Integer)
mcsYScale = iNewValue
RaiseEvent MWScaleChange
End Property

Public Property Get NumOfStepsX() As Long
NumOfStepsX = mcsXNumOfSteps
End Property

Public Property Get ScaleX_LToR() As Double
ScaleX_LToR = mcsScaleXRP
End Property

Public Property Get ScaleY_LToR() As Double
ScaleY_LToR = mcsScaleYRP
End Property

Public Property Get ScaleX_RToL() As Double
ScaleX_RToL = mcsScaleXLP
End Property

Public Property Get ScaleY_RToL() As Double
ScaleY_RToL = mcsScaleYLP
End Property


Public Function IsInScope(x As Long, y As Double) As Boolean
'--------------------------------------------------------------
'returns True if point X,Y is in current scope, False otherwise
'--------------------------------------------------------------
If x < RX1(csZmLvl) Then Exit Function
If x > RX2(csZmLvl) Then Exit Function
If y < RY1(csZmLvl) Then Exit Function
If y > RY2(csZmLvl) Then Exit Function
IsInScope = True
End Function
