VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CooSysO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'coordinate system for the overlay
'last modified: 12/27/2002 nt
'------------------------------------------------------------------
Option Explicit

Public csZmLvl As Integer           'up to 100 zoom levels
Public BuildingCS As Boolean        'true only when coo sys is initializing or reinitializing

Private mcsOriginXY As Integer      'XY axes intersection position
Private mcsXOrient As Integer       'horizontal orientation
Private mcsYOrient As Integer       'vertical orientation
Private mcsOrigin As Integer        '(XMin,YMin) position - read only
Private mcsYScale As Integer        'y scale type (linear or logarithmic)

Private mcsScaleXLP As Single       'scale to calculate logical coordinates from real
Private mcsScaleYLP As Single

Private mcsScaleXRP As Single       'scale to calculate real coordinates from logical
Private mcsScaleYRP As Single

Private mcsScaleR As Single         'determined by auto sizing

Private mcsLMinSz As Single         'minimum spot size
Private mcsLMaxSz As Single         'maximum spot size

Private mcsUniSpotSz As Single      'spot size for uniform size display

Private mcsAspectRatio As Single

Private mcsResolutionRX As Single
Private mcsResolutionRY As Single

Public VX0 As Long               'viewport-physical device coordinates
Public VXE As Long
Public VY0 As Long
Public VYE As Long

Public XYX0 As Single    'intersection of the XY axes
Public XYXE As Single    'point coordinates
Public XYY0 As Single
Public XYYE As Single

Public OrX0 As Single    'origin (as defined above)
Public OrXE As Single    'coordinates
Public OrY0 As Single
Public OrYE As Single

Private RX1(101) As Single     'real coordinates stack
Private RX2(101) As Single     'RX1(i)<RX2(i)
Private RY1(101) As Single     '
Private RY2(101) As Single     'RY1(i)<RY2(i)

Public RZ1 As Single
Public RZ2 As Single

Public SXOffset As Single  'small and large offsets
Public LXOffset As Single  'for the current coo sys
Public SYOffset As Single
Public LYOffset As Single

Public Event CooSysChanged()
Public Event YScaleChange()

Public Sub ZoomIn(ByVal lX1 As Long, ByVal lY1 As Long, _
                  ByVal lX2 As Long, ByVal lY2 As Long)
'--------------------------------------------------------
'input - logical coordinates of the box to zoom-in
'--------------------------------------------------------
Dim ReX1 As Single, ReX2 As Single
Dim ReY1 As Single, ReY2 As Single

If lX1 < lX2 Then
   ReX1 = CSng(lX1):   ReX2 = CSng(lX2)
Else
   ReX1 = CSng(lX2):   ReX2 = CSng(lX1)
End If
If lY1 < lY2 Then
   ReY1 = CSng(lY1):   ReY2 = CSng(lY2)
Else
   ReY1 = CSng(lY2):   ReY2 = CSng(lY1)
End If

LPToRP ReX1, ReY1, ReX2, ReY2

If ReX1 >= ReX2 Then
   ReX2 = ReX1 + 0.05:   ReX1 = ReX1 - 0.05
End If
If ReY1 >= ReY2 Then
   ReY2 = ReY1 + 1:   ReY1 = ReY1 - 1
End If

If csZmLvl < 100 Then
   csZmLvl = csZmLvl + 1
Else
   RX1(101) = RX1(100):   RX2(101) = RX2(100)
   RY1(101) = RY1(100):   RY2(101) = RY2(100)
End If
RX1(csZmLvl) = ReX1:   RX2(csZmLvl) = ReX2
RY1(csZmLvl) = ReY1:   RY2(csZmLvl) = ReY2
CalculateScales
RaiseEvent CooSysChanged
End Sub

Public Sub ZoomInR(ByVal dX1 As Single, ByVal dY1 As Single, _
                   ByVal dX2 As Single, ByVal dY2 As Single)
'-------------------------------------------------------------
'input - real coordinates of the box to zoom-in
'-------------------------------------------------------------
If csZmLvl < 100 Then
   csZmLvl = csZmLvl + 1
Else
   RX1(101) = RX1(100):   RX2(101) = RX2(100)
   RY1(101) = RY1(100):   RY2(101) = RY2(100)
End If
RX1(csZmLvl) = dX1:     RX2(csZmLvl) = dX2
RY1(csZmLvl) = dY1:     RY2(csZmLvl) = dY2
CalculateScales
RaiseEvent CooSysChanged
End Sub

Public Sub ZoomOut()
csZmLvl = 1
CalculateScales
RaiseEvent CooSysChanged
End Sub

Public Sub ZoomOut1()
If csZmLvl > 1 Then
   csZmLvl = csZmLvl - 1
   CalculateScales
   RaiseEvent CooSysChanged
End If
End Sub

Public Sub InitCooSys(ByVal MinRX As Single, ByVal MinRY As Single, ByVal MinRZ As Single, _
                      ByVal MaxRX As Single, ByVal MaxRY As Single, ByVal MaxRZ As Single)
'-------------------------------------------------------------------------------------------
'initializes coordinate system
'-------------------------------------------------------------------------------------------
RX1(0) = MinRX:    RX2(0) = MaxRX
Select Case mcsYScale
Case glVAxisLin
     RY1(0) = MinRY:         RY2(0) = MaxRY
Case glVAxisLog
     RY1(0) = Log(MinRY) / Log(10#)
     RY2(0) = Log(MaxRY) / Log(10#)
End Select
RX1(1) = RX1(0):    RX2(1) = RX2(0)
RY1(1) = RY1(0):    RY2(1) = RY2(0)
RZ1 = MinRZ:    RZ2 = MaxRZ
csZmLvl = 1
CalculateScales
RaiseEvent CooSysChanged
End Sub

Public Sub LPToRP(X1 As Single, Y1 As Single, x2 As Single, Y2 As Single)
'------------------------------------------------------------------------
'converts logical coordinates to the real coordinates
'------------------------------------------------------------------------
X1 = mcsScaleXRP * (X1 - LoX1) + RX1(csZmLvl)
x2 = mcsScaleXRP * (x2 - LoX1) + RX1(csZmLvl)
Y1 = mcsScaleYRP * (Y1 - LoY1) + RY1(csZmLvl)
Y2 = mcsScaleYRP * (Y2 - LoY1) + RY1(csZmLvl)
End Sub


Public Sub LPToRP1(X1 As Single, Y1 As Single)
'------------------------------------------------------------------------
'converts logical coordinates to the real coordinates
'------------------------------------------------------------------------
X1 = mcsScaleXRP * (X1 - LoX1) + RX1(csZmLvl)
Y1 = mcsScaleYRP * (Y1 - LoY1) + RY1(csZmLvl)
End Sub


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

Public Sub GetViewPortRectangle(ByVal paX As Long, ByVal paY As Long, _
                                rcT As Long, rcL As Long, rcB As Long, rcR As Long)
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

Public Property Get csYScale() As Integer
csYScale = mcsYScale
End Property

Public Property Let csYScale(ByVal iNewValue As Integer)
mcsYScale = iNewValue
If Not BuildingCS Then RaiseEvent YScaleChange
End Property

Public Property Get ScaleX_LToR() As Single
ScaleX_LToR = mcsScaleXRP
End Property

Public Property Get ScaleY_LToR() As Single
ScaleY_LToR = mcsScaleYRP
End Property

Public Property Get ScaleX_RToL() As Single
ScaleX_RToL = mcsScaleXLP
End Property

Public Property Get ScaleY_RToL() As Single
ScaleY_RToL = mcsScaleYLP
End Property

Public Sub CalculateScales()
On Error Resume Next
mcsScaleYRP = CSng(RY2(csZmLvl) - RY1(csZmLvl)) / (LoY2 - LoY1)
mcsScaleXRP = CSng(RX2(csZmLvl) - RX1(csZmLvl)) / (LoX2 - LoX1)
mcsScaleXLP = CSng(LoX2 - LoX1) / (RX2(csZmLvl) - RX1(csZmLvl))
mcsScaleYLP = CSng(LoY2 - LoY1) / (RY2(csZmLvl) - RY1(csZmLvl))
If mcsScaleYLP > 0 Then
   mcsAspectRatio = mcsScaleXLP / mcsScaleYLP
Else
   mcsAspectRatio = -1
End If

'calculate size for uniform size display
'mcsUniSpotSz = mcsScaleXLP * mcsResolutionRX
'If mcsUniSpotSz < lDfMinSz Then
'   mcsUniSpotSz = lDfMinSz
'ElseIf mcsUniSpotSz > lDfMaxSz Then
'   mcsUniSpotSz = lDfMaxSz
'End If

'next statement needs correction
If RZ1 <> RZ2 Then
   'Const lSzScale = 10
   'mcsScaleR = lSzScale * (mcsLMaxSz - mcsLMinSz) / (RZ2 - RZ1)
   'mcsScaleR = (mcsLMaxSz - mcsLMinSz) / (RZ2 - RZ1)
   mcsScaleR = (lDfMaxSz - lDfMinSz) / (RZ2 - RZ1)
Else
   mcsScaleR = 0
End If
End Sub

Public Property Get AspectRatio() As Single
AspectRatio = mcsAspectRatio
End Property

Public Property Get ScaleR() As Single
ScaleR = mcsScaleR
End Property

Private Sub Class_Initialize()
mcsAspectRatio = -1
End Sub

Public Property Get LMinSz() As Single
LMinSz = mcsLMinSz
End Property

Public Property Let LMinSz(ByVal NewLMinSz As Single)
mcsLMinSz = NewLMinSz * (LoXE - LoX0)
mcsUniSpotSz = (mcsLMinSz + mcsLMaxSz) / 2
If Not BuildingCS Then RaiseEvent CooSysChanged
End Property

Public Property Get LMaxSz() As Single
LMaxSz = mcsLMaxSz
End Property

Public Property Let LMaxSz(ByVal NewLMaxSz As Single)
mcsLMaxSz = NewLMaxSz * (LoXE - LoX0)
mcsUniSpotSz = (mcsLMinSz + mcsLMaxSz) / 2
If Not BuildingCS Then RaiseEvent CooSysChanged
End Property

Public Property Get UniSpotSize() As Single
UniSpotSize = mcsUniSpotSz
End Property

Public Property Get ResolutionRX() As Single
ResolutionRX = mcsResolutionRX
End Property

Public Property Let ResolutionRX(ByVal NewResolutionRX As Single)
mcsResolutionRX = NewResolutionRX
If Not BuildingCS Then RaiseEvent CooSysChanged
End Property

Public Property Get ResolutionRY() As Single
ResolutionRY = mcsResolutionRY
End Property

Public Property Let ResolutionRY(ByVal NewResolutionRY As Single)
mcsResolutionRY = NewResolutionRY
If Not BuildingCS Then RaiseEvent CooSysChanged
End Property
