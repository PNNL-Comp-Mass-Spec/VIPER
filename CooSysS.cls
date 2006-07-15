VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CooSysS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'more generic version of coordinate system; only keeps zoom levels and
'information about coordinate system as well as secret "T2S" codes
'---------------------------------------------------------------------
'last modified 04/07/2003 nt
'---------------------------------------------------------------------
Option Explicit

Const MAXZOOMDEPTH = 100

Public OwnerIndex As Long        'index of GelData caller structure
Public csZmLvl As Integer        'up to 100 zoom levels
Public BuildingCS As Boolean     'true only when coo sys is initializing or reinitializing

Dim mcsOriginXY As Integer       'XY axes intersection position
Dim mcsXOrient As Integer        'horizontal orientation
Dim mcsYOrient As Integer        'vertical orientation
Dim mcsOrigin As Integer         '(XMin,YMin) position - read only
Dim mcsYScale As Integer         'y scale type (linear or logarithmic)
Dim mcsType As Integer           'scan numbers or something else

Dim RX1(MAXZOOMDEPTH) As Long    'real coordinates stack - for x axis its index in the .ScanInfo array (discrete)
Dim RX2(MAXZOOMDEPTH) As Long    'RX1(i)<RX2(i)
Dim RY1(MAXZOOMDEPTH) As Double  '
Dim RY2(MAXZOOMDEPTH) As Double  'RY1(i)<RY2(i)

Dim RX1Prev As Long              'used to go step back after coordinate system reset
Dim RX2Prev As Long
Dim RY1Prev As Double
Dim RY2Prev As Double

'coordinate system does not need to know anything about device(viewport)
'but it should allow for logical coordinates to be set
Dim mLogX0 As Long
Dim mLogXE As Long
Dim mLogY0 As Long
Dim mLogYE As Long

'This is experimental aspect ratio; assumption is that the "best view" in
'this coo.sys. you get when window width and height are in ratio 20:1
Dim mIdealXYAspRat As Double     'ideal XY aspect ratio
Dim mCurrXYAspRat As Double      'current XY aspect ratio relative to ideal

'scales for conversion from logical to real and real to logical coordinates
Dim mScaleXLPRP As Double
Dim mScaleYLPRP As Double

Dim mScaleXRPLP As Double
Dim mScaleYRPLP As Double

Public Event CooSysChanged()

Public Sub ZoomInR(ByVal lX1 As Long, ByVal dY1 As Double, ByVal lX2 As Long, ByVal dY2 As Double)
'-------------------------------------------------------------------------------------------------
'input - real coordinates of the box to zoom-in
'-------------------------------------------------------------------------------------------------
If csZmLvl < MAXZOOMDEPTH Then
   csZmLvl = csZmLvl + 1
Else
   RX1(MAXZOOMDEPTH) = RX1(MAXZOOMDEPTH - 1):     RX2(MAXZOOMDEPTH) = RX2(MAXZOOMDEPTH - 1)
   RY1(MAXZOOMDEPTH) = RY1(MAXZOOMDEPTH - 1):     RY2(MAXZOOMDEPTH) = RY2(MAXZOOMDEPTH - 1)
End If
RX1(csZmLvl - 1) = lX1:       RX2(csZmLvl - 1) = lX2
RY1(csZmLvl - 1) = dY1:       RY2(csZmLvl - 1) = dY2
RX1Prev = RX1(csZmLvl - 1):   RX2Prev = RX2(csZmLvl - 1)
RY1Prev = RY1(csZmLvl - 1):   RY2Prev = RY2(csZmLvl - 1)
CalcConversionScales
CalcAspectRatio
RaiseEvent CooSysChanged
End Sub

Public Sub ZoomInRFirst(ByVal lX1 As Long, ByVal dY1 As Double, _
                        ByVal lX2 As Long, ByVal dY2 As Double)
'----------------------------------------------------------------------
'input - real coordinates of the box to zoom-in - first level of zoom
'----------------------------------------------------------------------
csZmLvl = 1
RX1(csZmLvl - 1) = lX1:       RX2(csZmLvl - 1) = lX2
RY1(csZmLvl - 1) = dY1:       RY2(csZmLvl - 1) = dY2
RX1Prev = RX1(csZmLvl - 1):   RX2Prev = RX2(csZmLvl - 1)
RY1Prev = RY1(csZmLvl - 1):   RY2Prev = RY2(csZmLvl - 1)
CalcConversionScales
CalcAspectRatio
RaiseEvent CooSysChanged
End Sub


Public Sub ZoomInL(ByVal lX1 As Long, ByVal lY1 As Long, _
                   ByVal lX2 As Long, ByVal lY2 As Long)
'---------------------------------------------------------
'input - logical coordinates of the box to zoom-in
'---------------------------------------------------------
Dim ReX1 As Double, ReX2 As Double
Dim lReX1 As Long, lReX2 As Long
Dim ReY1 As Double, ReY2 As Double

If lX1 < lX2 Then
   ReX1 = CDbl(lX1):   ReX2 = CDbl(lX2)
Else
   ReX1 = CDbl(lX2):   ReX2 = CDbl(lX1)
End If
If lY1 < lY2 Then
   ReY1 = CDbl(lY1):   ReY2 = CDbl(lY2)
Else
   ReY1 = CDbl(lY2):   ReY2 = CDbl(lY1)
End If

LPToRP ReX1, ReY1, ReX2, ReY2

lReX1 = CLng(ReX1):     lReX2 = CLng(ReX2)
If lReX1 >= lReX2 Then
   lReX2 = lReX1 + 1:   lReX1 = lReX1 - 1
End If
If ReY1 >= ReY2 Then
   ReY2 = ReY1 + 1:     ReY1 = ReY1 - 1
End If

ZoomInStorePrevious lReX1, ReY1, lReX1, ReY2

CalcConversionScales
CalcAspectRatio
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

Public Sub LPToRP(X1 As Double, Y1 As Double, x2 As Double, Y2 As Double)
'------------------------------------------------------------------------
'converts logical coordinates to the real coordinates
'------------------------------------------------------------------------
On Error GoTo LPToRPErrorHandler

If csZmLvl <= 0 Then
    X1 = mScaleXLPRP * (X1 - mLogX0)
    x2 = mScaleXLPRP * (x2 - mLogX0)
    Y1 = mScaleYLPRP * (Y1 - mLogY0)
    Y2 = mScaleYLPRP * (Y2 - mLogY0)
Else
    X1 = mScaleXLPRP * (X1 - mLogX0) + RX1(csZmLvl - 1)
    x2 = mScaleXLPRP * (x2 - mLogX0) + RX1(csZmLvl - 1)
    Y1 = mScaleYLPRP * (Y1 - mLogY0) + RY1(csZmLvl - 1)
    Y2 = mScaleYLPRP * (Y2 - mLogY0) + RY1(csZmLvl - 1)
End If
Exit Sub

LPToRPErrorHandler:
Debug.Assert False
Debug.Print "Error in CooSysS.LPToRP"
Resume Next
End Sub


Public Sub RPToLP(X1 As Double, Y1 As Double, x2 As Double, Y2 As Double)
'------------------------------------------------------------------------
'converts real coordinates to the logical coordinates
'------------------------------------------------------------------------
X1 = mScaleXRPLP * (X1 - RX1(csZmLvl - 1)) + mLogX0
x2 = mScaleXRPLP * (x2 - RX1(csZmLvl - 1)) + mLogX0
Y1 = mScaleYRPLP * (Y1 - RY1(csZmLvl - 1)) + mLogY0
Y2 = mScaleYRPLP * (Y2 - RY1(csZmLvl - 1)) + mLogY0
End Sub


Public Sub ZoomOut()
'-------------------------------------------------
'zoom out all the way
'-------------------------------------------------
csZmLvl = 1
CalcConversionScales
CalcAspectRatio
RaiseEvent CooSysChanged
End Sub

Public Sub ZoomOut1()
'-------------------------------------------------
'zoom out one level
'-------------------------------------------------
If csZmLvl > 1 Then
   csZmLvl = csZmLvl - 1
   CalcConversionScales
   CalcAspectRatio
   RaiseEvent CooSysChanged
End If
End Sub

Public Sub ZoomPrevious()
'----------------------------------------------------------------------------
'this is ZoomOut1 if there is anything to zoom-out and zoom-in to remembered
'----------------------------------------------------------------------------
If csZmLvl > 1 Then
   Call ZoomOut1
Else
   Call ZoomInR(RX1Prev, RY1Prev, RX2Prev, RY2Prev)
End If
End Sub

Public Property Get csType() As Integer
csType = mcsType
End Property

Public Property Let csType(ByVal iNewType As Integer)
mcsType = iNewType
If Not BuildingCS Then RaiseEvent CooSysChanged
End Property

Public Property Get csOriginXY() As Integer
csOriginXY = mcsOriginXY
End Property

Public Property Let csOriginXY(ByVal iOXY As Integer)
mcsOriginXY = iOXY
If mcsXOrient > 0 And mcsYOrient > 0 Then mcsOrigin = mcsXOrient + 2 * (mcsYOrient - 1)
End Property

Public Property Get csXOrient() As Integer
csXOrient = mcsXOrient
End Property

Public Property Let csXOrient(ByVal iXOrient As Integer)
mcsXOrient = iXOrient
If mcsYOrient > 0 Then mcsOrigin = mcsXOrient + 2 * (mcsYOrient - 1)
End Property

Public Property Get csYOrient() As Integer
csYOrient = mcsYOrient
End Property

Public Property Let csYOrient(ByVal iYOrient As Integer)
mcsYOrient = iYOrient
If mcsXOrient > 0 Then mcsOrigin = mcsXOrient + 2 * (mcsYOrient - 1)
End Property

Public Property Get csOrigin() As Integer
csOrigin = mcsOrigin
End Property

Public Property Get csYScale() As Integer
csYScale = mcsYScale
End Property

Public Property Let csYScale(ByVal iNewValue As Integer)
mcsYScale = iNewValue
RaiseEvent CooSysChanged
End Property


Public Property Get csIdealXYAspRat() As Double
csIdealXYAspRat = mIdealXYAspRat
End Property

Public Property Let csIdealXYAspRat(ByVal NewIdealAspectRatio As Double)
If NewIdealAspectRatio > 0 Then
   mIdealXYAspRat = NewIdealAspectRatio
Else                                    'should raise error but don't feel like it
   mIdealXYAspRat = 1
End If
End Property


Public Property Let LogX0(ByVal NewLogX0 As Long)
mLogX0 = NewLogX0
End Property

Public Property Get LogX0() As Long
LogX0 = mLogX0
End Property


Public Property Let LogXE(ByVal NewLogXE As Long)
mLogXE = NewLogXE
End Property

Public Property Get LogXE() As Long
LogXE = mLogXE
End Property


Public Property Let LogY0(ByVal NewLogY0 As Long)
mLogY0 = NewLogY0
End Property

Public Property Get LogY0() As Long
LogY0 = mLogY0
End Property


Public Property Let LogYE(ByVal NewLogYE As Long)
mLogYE = NewLogYE
End Property

Public Property Get LogYE() As Long
LogYE = mLogYE
End Property

'read only properties

'conversion scales
Public Property Get ScaleY_LToR() As Double
ScaleY_LToR = mScaleYLPRP
End Property

Public Property Get ScaleX_LToR() As Double
ScaleX_LToR = mScaleXLPRP
End Property

Public Property Get ScaleY_RToL() As Double
ScaleY_RToL = mScaleYRPLP
End Property

Public Property Get ScaleX_RToL() As Double
ScaleX_RToL = mScaleXRPLP
End Property

'calculated aspect ratio; relative to the ideal aspect ratio
Public Property Get csXYAspRat() As Double
csXYAspRat = mCurrXYAspRat
End Property

'largest zoom window(first level of zoom)
Public Property Get TtlRXMin() As Double
On Error Resume Next
TtlRXMin = RX1(0)
End Property

Public Property Get TtlRXMax() As Double
On Error Resume Next
TtlRXMax = RX2(0)
End Property

Public Property Get TtlRYMin() As Double
On Error Resume Next
TtlRYMin = RY1(0)
End Property

Public Property Get TtlRYMax() As Double
On Error Resume Next
TtlRYMax = RY2(0)
End Property

'current zoom window
Public Property Get CurrRXMin() As Double
On Error Resume Next
CurrRXMin = RX1(csZmLvl - 1)
End Property

Public Property Get CurrRXMax() As Double
On Error Resume Next
CurrRXMax = RX2(csZmLvl - 1)
End Property

Public Property Get CurrRYMin() As Double
On Error Resume Next
CurrRYMin = RY1(csZmLvl - 1)
End Property

Public Property Get CurrRYMax() As Double
On Error Resume Next
CurrRYMax = RY2(csZmLvl - 1)
End Property

'---------------------------------------------------------------------------------
Public Function IsInScope(x As Long, y As Double) As Boolean
'-------------------------------------------------------------------
'returns True if point X,Y is in current scope, False otherwise
'-------------------------------------------------------------------
If x < RX1(csZmLvl - 1) Then Exit Function    'look current zoom level
If x > RX2(csZmLvl - 1) Then Exit Function
If y < RY1(csZmLvl - 1) Then Exit Function
If y > RY2(csZmLvl - 1) Then Exit Function
IsInScope = True
End Function

Public Function IsInLargeScope(x As Long, y As Double) As Boolean
'----------------------------------------------------------------
'returns True if point X,Y is in total scope, False otherwise
'----------------------------------------------------------------
If x < RX1(0) Then Exit Function        'look first zoom level
If x > RX2(0) Then Exit Function
If y < RY1(0) Then Exit Function
If y > RY2(0) Then Exit Function
IsInLargeScope = True
End Function

Private Sub CalcAspectRatio()
'----------------------------------------------------------------
'calculates aspect ratio of the current view
'----------------------------------------------------------------
On Error Resume Next
mCurrXYAspRat = ((RX2(csZmLvl - 1) - RX1(csZmLvl - 1)) / _
            (RY2(csZmLvl - 1) - RY1(csZmLvl - 1))) / mIdealXYAspRat
End Sub

Private Sub Class_Initialize()
mIdealXYAspRat = 1
End Sub


Public Function CalcConversionScales() As Boolean
'---------------------------------------------------------------------
'calculates scales for conversion between real and logical coordinates
'returns True if successful
'---------------------------------------------------------------------
On Error GoTo exit_CalcConversionScales
mScaleYLPRP = (RY2(csZmLvl - 1) - RY1(csZmLvl - 1)) / (LogYE - LogY0)
mScaleYRPLP = 1 / mScaleYLPRP
mScaleXLPRP = (RX2(csZmLvl - 1) - RX1(csZmLvl - 1)) / (LogXE - LogX0)
mScaleXRPLP = 1 / mScaleXLPRP
CalcConversionScales = True
exit_CalcConversionScales:
End Function
