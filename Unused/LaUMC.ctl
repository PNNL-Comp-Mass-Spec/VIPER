VERSION 5.00
Begin VB.UserControl LaUMC 
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4845
   ScaleHeight     =   3255
   ScaleWidth      =   4845
   Begin VB.PictureBox picD 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   317
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
   Begin VB.Menu mnuF 
      Caption         =   "Function"
      Visible         =   0   'False
      Begin VB.Menu mnuFViewCoo 
         Caption         =   "View Coordinates"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFSelectAll 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnuFClearSelection 
         Caption         =   "Cl&ear Selection"
      End
      Begin VB.Menu mnuFSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFS 
         Caption         =   "Sp&ot Shape"
         Begin VB.Menu mnuFShape 
            Caption         =   "&Oval"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuFShape 
            Caption         =   "&Rectangle"
            Index           =   1
         End
         Begin VB.Menu mnuFShape 
            Caption         =   "Ro&und Rectangle"
            Index           =   2
         End
         Begin VB.Menu mnuFShape 
            Caption         =   "&Star"
            Index           =   3
         End
      End
      Begin VB.Menu mnuFSSz 
         Caption         =   "Spot Si&ze"
         Begin VB.Menu mnuFSSzIncrease 
            Caption         =   "&Increase"
         End
         Begin VB.Menu mnuFSSzDecrease 
            Caption         =   "&Decrease"
         End
      End
      Begin VB.Menu mnuFBackColor 
         Caption         =   "Selection &Color"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuFSpotColor 
         Caption         =   "&Spot Color"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDFCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "LaUMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'function spot selection user control
'--------------------------------------------------------
'last modified: 10/03/2002 nt
'--------------------------------------------------------
Option Explicit

Const MSG_TITLE = "2D Displays - Unique Mass Classes"

Const MAXDATACOUNT = 25000              'maximum data points in this display

'Public Enum sSpotsShape
'    sCircle = 0
'    sRectangle = 1
'    sRoundRectangle = 2
'    sStar = 3
'End Enum

Public Enum umcAbuScale
    umcAbuLin = 0
    umcAbuLog = 1
End Enum

Dim DisplayInd As Long          'index of display to which this UMC belongs
Dim UMCInd As Long              'inde



'public properties
Public VLabel As String     'label on vertical axis
Public VNumFmt As String    'numerical format for vertical axis
Public HLabel As String     'label on horizontal axis
Public HNumFmt As String    'numerical format for horizontal axis

Public SpotShape As Long

Public RegColor As Long             'regular spot color
Public SelColor As Long             'selection spot color

Dim mViewCoordinates As Boolean

Dim mViewWindow As sSpotsView       'window can be fixed or variable depending on
                                    'loaded data
                                                                        
'fixed coordinates
Dim mFixedMinX As Double
Dim mFixedMaxX As Double
Dim mFixedMinY As Double
Dim mFixedMaxY As Double
'current data coordinates
Dim mCurrMinX As Double
Dim mCurrMaxX As Double
Dim mCurrMinY As Double
Dim mCurrMaxY As Double
'working coordinatges
Dim mMinX As Double            'minimum of data range X
Dim mMaxX As Double            'maximum of data range X
Dim mMinY As Double            'minimum of data range Y
Dim mMaxY As Double            'maximum of data range Y
                            
Dim mDataCnt As Long                'count of data points
Dim mDataID() As Long
Dim mDataType() As Long
Dim mDataX() As Double              'data that has to be drawn
Dim mDataY() As Double
Dim mDataSelected() As Boolean      'data selection

Dim mSpotSize As Long
Dim mHalfSize As Long

Dim mScaleX As Double         'scales used to draw
Dim mScaleY As Double         'on logical window

Dim mIsFree As Boolean        'read only properety; True if control
                              'is free for assignment; False if it is
                              'filled with data

Dim mHotSpotInd As Long       'index of hot spot; -1 if none
Dim mCurrX As String          'current coordinates (formated)
Dim mCurrY As String

'coordinate system-viewport coordinates
Dim VPX0 As Long
Dim VPY0 As Long
Dim VPXE As Long
Dim VPYE As Long

'actual drawing points
Dim mgX() As Long
Dim mgY() As Long


Dim hRegBrush As Long          'regular brush
Dim hSelBrush As Long          'selected brush
Dim hBackPen As Long           'back color pen
Dim paPoints() As POINTAPI          'used to calculate dev to log coordinate conversion


Public Function AddSpotOne(SpotID As Long, SpotX As Boolean, SpotY As Double) As Boolean
'------------------------------------------------------------------------------------
'adds spot to current data
'------------------------------------------------------------------------------------
On Error GoTo err_AddSpotOne
If mDataCnt >= MAXDATACOUNT Then
   MsgBox "Too many data points, cannot accept more!", vbOKOnly, MSG_TITLE
   Exit Function
End If
mDataCnt = mDataCnt + 1
ReDim Preserve mDataID(mDataCnt - 1)
ReDim Preserve mDataX(mDataCnt - 1)
ReDim Preserve mDataY(mDataCnt - 1)
ReDim Preserve mDataSelected(mDataCnt - 1)
mDataID(mDataCnt - 1) = SpotID
mDataX(mDataCnt - 1) = SpotX
mDataY(mDataCnt - 1) = SpotY
If SpotX < mCurrMinX Then mCurrMinX = SpotX
If SpotX > mCurrMaxX Then mCurrMaxX = SpotX
If SpotY < mCurrMinY Then mCurrMinY = SpotY
If SpotY > mCurrMaxY Then mCurrMaxY = SpotY
AddSpotOne = True
SpotsRefresh
err_AddSpotOne:
End Function

Public Function AddSpotsMany(SpotID() As Long, SpotX() As Double, SpotY() As Double) As Boolean
'-----------------------------------------------------------------------------------------------
'adds spot to current data
'-----------------------------------------------------------------------------------------------
Dim i As Long, FirstSpotInd As Long
Dim Resp As Long
On Error GoTo err_AddSpotsMany
FirstSpotInd = LBound(SpotID)
mDataCnt = UBound(SpotID) - FirstSpotInd + 1
If mDataCnt > MAXDATACOUNT Then
   mDataCnt = MAXDATACOUNT
   Resp = MsgBox("Too many spots; only first " & MAXDATACOUNT & " will be presented!", vbOKCancel, MSG_TITLE)
   If Resp <> vbOK Then Exit Function
End If
ReDim mDataID(mDataCnt - 1)
ReDim mDataX(mDataCnt - 1)
ReDim mDataY(mDataCnt - 1)
ReDim mDataSelected(mDataCnt - 1)
mCurrMinX = 1E+308:     mCurrMaxX = -1E+308
mCurrMinY = 1E+308:     mCurrMaxY = -1E+308
picD.Cls
For i = FirstSpotInd To UBound(SpotID)
    mDataID(i - FirstSpotInd) = SpotID(i)
    mDataX(i - FirstSpotInd) = SpotX(i)
    mDataY(i - FirstSpotInd) = SpotY(i)
    If SpotX(i) < mCurrMinX Then mCurrMinX = SpotX(i)
    If SpotX(i) > mCurrMaxX Then mCurrMaxX = SpotX(i)
    If SpotY(i) < mCurrMinY Then mCurrMinY = SpotY(i)
    If SpotY(i) > mCurrMaxY Then mCurrMaxY = SpotY(i)
Next i
mHotSpotInd = -1
mIsFree = False
Call SpotsRefresh
AddSpotsMany = True
err_AddSpotsMany:
End Function


Public Sub Clear()
'---------------------------------------------------
'erase arrays and clean picture; declare object free
'---------------------------------------------------
On Error Resume Next
Erase mDataID
Erase mDataX
Erase mDataY
Erase mDataSelected
mDataCnt = 0
mCurrMinX = 1E+308:     mCurrMaxX = -1E+308
mCurrMinY = 1E+308:     mCurrMaxY = -1E+308
mHotSpotInd = -1
picD.Cls
mIsFree = True
End Sub


Public Sub Draw()
'-----------------------------------------------------
'draws spots on picture device context
'-----------------------------------------------------
Dim OldDC As Long
Dim Res As Long
OldDC = SaveDC(picD.hDC)
DGCooSys picD.hDC, picD.ScaleWidth, picD.ScaleHeight
DGDrawCooSys picD.hDC
WriteLabels picD.hDC
If mDataCnt > 0 Then
   DrawSpots picD.hDC
   DrawSelectedSpots picD.hDC
End If
Res = RestoreDC(picD.hDC, OldDC)
End Sub


Private Sub DGCooSys(ByVal hDC As Long, ByVal dcw As Long, ByVal dch As Long)
'----------------------------------------------------------------------------
'establishes coordinate system on device context
'----------------------------------------------------------------------------
Dim Res As Long
Dim ptPoint As POINTAPI
Dim szSize As Size
On Error Resume Next
VPX0 = 35
VPY0 = dch - 25
VPXE = dcw - 40
VPYE = 50 - dch
Res = SetMapMode(hDC, MM_ANISOTROPIC)
'logical window
Res = SetWindowOrgEx(hDC, LDfX0, LDfY0, ptPoint)
Res = SetWindowExtEx(hDC, LDfXE, LDfYE, szSize)
'viewport
Res = SetViewportOrgEx(hDC, VPX0, VPY0, ptPoint)
Res = SetViewportExtEx(hDC, VPXE, VPYE, szSize)
'set working window coordinates
Select Case mViewWindow
Case sFixedWindow
     mMinX = mFixedMinX:        mMaxX = mFixedMaxX
     mMinY = mFixedMinY:        mMaxY = mFixedMaxY
Case sVariableWindow
     mMinX = mCurrMinX:        mMaxX = mCurrMaxX
     mMinY = mCurrMinY:        mMaxY = mCurrMaxY
End Select
End Sub

Public Sub DGDrawCooSys(ByVal hDC As Long)
'-----------------------------------------
'draws coordinate system on device context
'-----------------------------------------
Dim ptPoint As POINTAPI
Dim Res As Long
On Error Resume Next
'horizontal
Res = MoveToEx(hDC, LDfX0, LDfY0, ptPoint)
Res = LineTo(hDC, LDfX0 + LDfXE, LDfY0)
'vertical
Res = MoveToEx(hDC, LDfX0, LDfY0, ptPoint)
Res = LineTo(hDC, LDfX0, LDfY0 + LDfYE)
End Sub

Private Sub mnuDFCopy_Click()
Call CopyFD
End Sub

Private Sub mnuFClearSelection_Click()
On Error Resume Next
ReDim mDataSelected(mDataCnt - 1)
SpotsRefresh
End Sub

Private Sub mnuFSelectAll_Click()
Dim i As Long
On Error Resume Next
For i = 0 To mDataCnt - 1
    mDataSelected(i) = True
Next i
SpotsRefresh
End Sub

Private Sub mnuFShape_Click(Index As Integer)
Dim i As Long
SpotShape = Index
For i = mnuFShape.LBound To mnuFShape.UBound
    If i = Index Then
       mnuFShape(i).Checked = True
    Else
       mnuFShape(i).Checked = False
    End If
Next i
SpotsRefresh
End Sub

Private Sub mnuFSSzDecrease_Click()
If mSpotSize > 10 Then
   mSpotSize = mSpotSize - mSpotSize \ 3
   mHalfSize = mSpotSize \ 2
SpotsRefresh
End If
End Sub

Private Sub mnuFSSzIncrease_Click()
mSpotSize = mSpotSize + mSpotSize \ 3
mHalfSize = mSpotSize \ 2
SpotsRefresh
End Sub

Private Sub mnuFViewCoo_Click()
'toggle visibility of coordinates
mViewCoordinates = Not mViewCoordinates
mnuFViewCoo.Checked = mViewCoordinates
If Not mViewCoordinates Then WriteCoordinates picD.hDC, picD.ScaleWidth, picD.ScaleHeight  'this will clear coordinates
End Sub

Private Sub picD_Click()
On Error Resume Next
If mHotSpotInd >= 0 Then
   mDataSelected(mHotSpotInd) = Not mDataSelected(mHotSpotInd)
   SpotsRefresh
End If
End Sub

Private Sub picD_KeyDown(KeyCode As Integer, Shift As Integer)
'-------------------------------------------------------------
'copy as metafile on system clipboard on Ctrl+C combination
'-------------------------------------------------------------
Dim CtrlDown As Boolean
CtrlDown = (Shift And vbCtrlMask) > 0
If KeyCode = vbKeyC And CtrlDown Then CopyFD
End Sub

Private Sub picD_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then UserControl.PopupMenu mnuF
End Sub

Private Sub picD_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Lbl As String
Dim szLbl As Size
Dim Res As Long
On Error Resume Next
If mDataCnt > 0 Then
   ReDim paPoints(0)
   paPoints(0).x = x
   paPoints(0).y = y
   DevLogConversion 0, 1
   TrackHotSpot paPoints(0).x, paPoints(0).y
   If mViewCoordinates Then WriteCoordinates picD.hDC, picD.ScaleWidth, picD.ScaleHeight
End If
End Sub

Private Sub picD_Paint()
SpotsRefresh
End Sub

Private Sub UserControl_Initialize()
VLabel = "?"
HLabel = "?"
HNumFmt = "0.00"
VNumFmt = "0"
mIsFree = True
mSpotSize = LDfXE / 50
mHalfSize = mSpotSize / 2
SpotShape = sCircle
RegColor = vbGreen
SelColor = vbRed
mViewCoordinates = mnuFViewCoo.Checked
mViewWindow = sFixedWindow
mFixedMinX = 0:    mFixedMaxX = 1
mFixedMinY = 0:     mFixedMaxY = 10000
SetDrawingObjects
End Sub

Private Sub UserControl_Resize()
picD.width = UserControl.ScaleWidth
picD.Height = UserControl.ScaleHeight
End Sub

Private Sub CalculateSpots()
Dim i As Long
On Error Resume Next

mScaleY = LDfYE / mMaxY
mScaleX = LDfXE / mMaxX

ReDim mgX(mDataCnt - 1)
ReDim mgY(mDataCnt - 1)
For i = 0 To mDataCnt - 1
    mgX(i) = CLng(mDataX(i) * mScaleX)
    mgY(i) = CLng(mDataY(i) * mScaleY)
Next i
End Sub

Public Sub DrawSpots(ByVal hDC As Long)
'--------------------------------------------
'draws spots to device context
'--------------------------------------------
Dim ptPoint As POINTAPI
Dim OldBrush As Long
Dim i As Long
On Error Resume Next
OldBrush = SelectObject(hDC, hRegBrush)
Select Case SpotShape
Case sCircle
    For i = 0 To mDataCnt - 1
        Ellipse hDC, mgX(i) - mHalfSize, mgY(i) - mHalfSize, mgX(i) + mHalfSize, mgY(i) + mHalfSize
    Next i
Case sRectangle
    For i = 0 To mDataCnt - 1
        Rectangle hDC, mgX(i) - mHalfSize, mgY(i) - mHalfSize, mgX(i) + mHalfSize, mgY(i) + mHalfSize
    Next i
Case sRoundRectangle
    For i = 0 To mDataCnt - 1
        RoundRect hDC, mgX(i) - mHalfSize, mgY(i) - mHalfSize, mgX(i) + mHalfSize, mgY(i) + mHalfSize, mHalfSize, mHalfSize
    Next i
Case sStar
    Dim ThisStar() As Long
    Dim ptAPIs(7) As POINTAPI
    Dim j As Long
    For i = 0 To mDataCnt - 1
        GetStarPoints mgX(i), mgY(i), ThisStar()
        For j = 0 To 7
            ptAPIs(j).x = ThisStar(j, 0)
            ptAPIs(j).y = ThisStar(j, 1)
        Next j
        Polygon hDC, ptAPIs(0), 8
    Next i
End Select
Call SelectObject(hDC, OldBrush)
End Sub

Public Sub DrawSelectedSpots(ByVal hDC As Long)
'--------------------------------------------
'draws spots to device context
'--------------------------------------------
Dim ptPoint As POINTAPI
Dim Res As Long
Dim i As Long
Dim OldBrush As Long
Dim HalfSize As Long
On Error Resume Next
OldBrush = SelectObject(hDC, hSelBrush)
Select Case SpotShape
Case sCircle
    For i = 0 To mDataCnt - 1
        If mDataSelected(i) Then
           Ellipse hDC, mgX(i) - mHalfSize, mgY(i) - mHalfSize, mgX(i) + mHalfSize, mgY(i) + mHalfSize
        End If
    Next i
Case sRectangle
    For i = 0 To mDataCnt - 1
        If mDataSelected(i) Then
           Rectangle hDC, mgX(i) - mHalfSize, mgY(i) - mHalfSize, mgX(i) + mHalfSize, mgY(i) + mHalfSize
        End If
    Next i
Case sRoundRectangle
    For i = 0 To mDataCnt - 1
        If mDataSelected(i) Then
           RoundRect hDC, mgX(i) - mHalfSize, mgY(i) - mHalfSize, mgX(i) + mHalfSize, mgY(i) + mHalfSize, mHalfSize, mHalfSize
        End If
    Next i
Case sStar
    Dim ThisStar() As Long
    Dim ptAPIs(7) As POINTAPI
    Dim j As Long
    For i = 0 To mDataCnt - 1
        If mDataSelected(i) Then
           GetStarPoints mgX(i), mgY(i), ThisStar()
           For j = 0 To 7
               ptAPIs(j).x = ThisStar(j, 0)
               ptAPIs(j).y = ThisStar(j, 1)
           Next j
           Polygon hDC, ptAPIs(0), 8
        End If
    Next i
End Select
Call SelectObject(hDC, OldBrush)
End Sub

Public Property Get IsFree() As Boolean
IsFree = mIsFree
End Property

Private Sub WriteLabels(ByVal hDC As Long)
Dim lfLogFont As LOGFONT
Dim lOldFont As Long
Dim lNewFont As Long
Dim lFont As Long
Dim Res As Long
On Error Resume Next
'get the font from the picture box control (Arial Narrow)
lFont = SelectObject(hDC, GetStockObject(SYSTEM_FONT))
Res = GetObjectAPI(lFont, Len(lfLogFont), lfLogFont)
Res = SelectObject(hDC, lFont)
'create new logical font
lfLogFont.lfWidth = 100
lfLogFont.lfHeight = 600
lNewFont = CreateFontIndirect(lfLogFont)
'select newly created logical font to DC
lOldFont = SelectObject(hDC, lNewFont)
'draw coordinate axes labels
Res = TextOut(hDC, 9600, -800, HLabel, Len(HLabel))
Res = TextOut(hDC, 50, 10700, VLabel, Len(VLabel))
'draw numeric markers on vertical axes
WriteVMarkers hDC
WriteHMarkers hDC
'restore old font
Res = SelectObject(hDC, lOldFont)
DeleteObject (lNewFont)
End Sub


Public Sub CopyFD()
Dim Res As Long

Dim OldSize As Size
Dim OldDC As Long
Dim hRefDC As Long
Dim emfDC As Long       'metafile device context
Dim emfHandle As Long   'metafile handle

Dim iWidthMM As Long
Dim iHeightMM As Long
Dim iWidthPels As Long
Dim iHeightPels As Long
Dim iMMPerPelX As Double
Dim iMMPerPelY As Double
Dim rcRef As RECT           'reference rectangle
On Error Resume Next


hRefDC = picD.hDC
iWidthMM = GetDeviceCaps(hRefDC, HORZSIZE)
iHeightMM = GetDeviceCaps(hRefDC, VERTSIZE)
iWidthPels = GetDeviceCaps(hRefDC, HORZRES)
iHeightPels = GetDeviceCaps(hRefDC, VERTRES)

iMMPerPelX = (iWidthMM * 100) / iWidthPels
iMMPerPelY = (iHeightMM * 100) / iHeightPels

rcRef.Top = 0
rcRef.Left = 0
rcRef.Bottom = picD.ScaleHeight
rcRef.Right = picD.ScaleWidth
'convert to himetric units
rcRef.Left = rcRef.Left * iMMPerPelX
rcRef.Top = rcRef.Top * iMMPerPelY
rcRef.Right = rcRef.Right * iMMPerPelX
rcRef.Bottom = rcRef.Bottom * iMMPerPelY

emfDC = CreateEnhMetaFile(hRefDC, vbNullString, rcRef, vbNullString)
OldDC = SaveDC(emfDC)

DGCooSys emfDC, picD.ScaleWidth, picD.ScaleHeight
DGDrawCooSys emfDC
WriteLabels emfDC
DrawSpots emfDC
DrawSelectedSpots emfDC
If mViewCoordinates Then WriteCoordinates emfDC, picD.ScaleWidth, picD.ScaleHeight
Res = RestoreDC(emfDC, OldDC)
emfHandle = CloseEnhMetaFile(emfDC)

Res = OpenClipboard(picD.hWnd)
Res = EmptyClipboard()
Res = SetClipboardData(CF_ENHMETAFILE, emfHandle)
Res = CloseClipboard
End Sub

Private Sub WriteVMarkers(ByVal hDC As Long)
Dim VDelta As Double
Dim LDelta As Long
Dim VLbl As String
Dim szLbl As Size
Dim ptPoint As POINTAPI
Dim Res As Long
Dim i As Long
VDelta = (mMaxY - mMinY) / 3
LDelta = LDfYE / 3
For i = 0 To 4
    VLbl = Format$(mMinY + i * VDelta, VNumFmt)
    Res = GetTextExtentPoint32(hDC, VLbl, Len(VLbl), szLbl)
    Res = TextOut(hDC, -szLbl.cx - 100, i * LDelta + 400, VLbl, Len(VLbl))
    Res = MoveToEx(hDC, -50, i * LDelta, ptPoint)
    Res = LineTo(hDC, 50, i * LDelta)
Next i
'write also number of points
VLbl = "Spots count: " & mDataCnt
Res = GetTextExtentPoint32(hDC, VLbl, Len(VLbl), szLbl)
Res = TextOut(hDC, 2000, 10800, VLbl, Len(VLbl))
End Sub


Private Sub WriteHMarkers(ByVal hDC As Long)
Dim HDelta As Double
Dim LDelta As Long
Dim HLbl As String
Dim szLbl As Size
Dim ptPoint As POINTAPI
Dim Res As Long
Dim i As Long
LDelta = LDfXE / 3
HDelta = (mMaxX - mMinX) / 3
For i = 0 To 4
   HLbl = Format(mMinX + i * HDelta, HNumFmt)
   Res = GetTextExtentPoint32(hDC, HLbl, Len(HLbl), szLbl)
   Res = TextOut(hDC, i * LDelta - szLbl.cy \ 2, -200, HLbl, Len(HLbl))
   Res = MoveToEx(hDC, i * LDelta, -50, ptPoint)
   Res = LineTo(hDC, i * LDelta, 100)
Next i
End Sub


Private Sub WriteCoordinates(hDC As Long, w As Long, H As Long)
Dim OldDC As Long
Dim lfLogFont As LOGFONT
Dim lOldFont As Long
Dim lNewFont As Long
Dim Lbl As String
Dim lFont As Long
Dim szLbl As Size
Dim OldPen As Long
Dim Res As Long
On Error Resume Next
OldDC = SaveDC(hDC)
DGCooSys hDC, w, H
'get the font from the picture box control (Arial Narrow)
lFont = SelectObject(hDC, GetStockObject(SYSTEM_FONT))
Res = GetObjectAPI(lFont, Len(lfLogFont), lfLogFont)
Res = SelectObject(hDC, lFont)
'create new logical font
lfLogFont.lfWidth = 100
lfLogFont.lfHeight = 600
lNewFont = CreateFontIndirect(lfLogFont)
'select newly created logical font to DC
lOldFont = SelectObject(hDC, lNewFont)
'write new coordinates
Lbl = "Coordinates: " & mCurrX & ",  " & mCurrY
If mHotSpotInd >= 0 Then Lbl = Lbl & "   ID: " & mDataID(mHotSpotInd)
Res = GetTextExtentPoint32(hDC, Lbl, Len(Lbl), szLbl)
'clear whatever was written before; just draw regular rectangle
OldPen = SelectObject(hDC, hBackPen)             'invisible pen
Rectangle hDC, 4000, 10800, 11000, 10800 - szLbl.cy
Res = SelectObject(hDC, OldPen)
If mViewCoordinates Then Res = TextOut(hDC, 4000, 10800, Lbl, Len(Lbl))
'restore old font
Res = SelectObject(hDC, lOldFont)
DeleteObject (lNewFont)
Res = RestoreDC(hDC, OldDC)
End Sub


Public Sub SpotsRefresh()
'-------------------------------------------
'rebuild coordinate system and redraws spots
'-------------------------------------------
picD.Cls
Call CalculateSpots
Call Draw
End Sub


Public Function WriteSToFile(ByVal FileName As String) As Boolean
'---------------------------------------------------------------
'write bins results to text semicolon delimited file
'returns True if successful, False if not or any error occurs
'NOTE: bins are appended to a file; create file if not found
'---------------------------------------------------------------
Dim i As Long
Dim hfile As Integer
On Error GoTo exit_WriteSToFile

hfile = FreeFile
Open FileName For Append As hfile
Print #hfile, "Presenting - " & HLabel & ", " & VLabel
Print #hfile, "Range - [" & mMinX & ", " & mMaxX & "]X[" & mMinY & ", " & mMaxY & "]"
Print #hfile, vbCrLf
Print #hfile, "ID;X;Y"
For i = 0 To mDataCnt - 1
    Print #hfile, mDataID(i) & ";" & mDataX(i) & ";" & mDataY(i)
Next i
Close hfile
WriteSToFile = True
exit_WriteSToFile:
End Function

Public Sub DestroyDrawingObjects()
On Error Resume Next
If hRegBrush <> 0 Then DeleteObject (hRegBrush)
If hSelBrush <> 0 Then DeleteObject (hSelBrush)
If hBackPen <> 0 Then DeleteObject (hBackPen)
End Sub

Public Sub SetDrawingObjects()
On Error Resume Next
hRegBrush = CreateSolidBrush(RegColor)
hSelBrush = CreateSolidBrush(SelColor)
hBackPen = CreatePen(PS_SOLID, 1, picD.BackColor)
End Sub

Private Sub UserControl_Terminate()
DestroyDrawingObjects
End Sub

Private Sub DevLogConversion(ByVal ConversionType As Integer, ByVal NumOfPoints As Integer)
Dim OldDC As Long, Res As Long
OldDC = SaveDC(picD.hDC)
DGCooSys picD.hDC, picD.ScaleWidth, picD.ScaleHeight
Select Case ConversionType
Case 0
    Res = DPtoLP(picD.hDC, paPoints(0), NumOfPoints)
Case 1
    Res = LPtoDP(picD.hDC, paPoints(0), NumOfPoints)
End Select
Res = RestoreDC(picD.hDC, OldDC)
End Sub


Private Sub TrackHotSpot(lx As Long, ly As Long)
Dim i As Long
On Error Resume Next
mCurrX = Format$(lx / mScaleX, "0.0000")            'current coordinates
mCurrY = Format$(ly / mScaleY, "0.0000")
For i = mDataCnt - 1 To 0 Step -1
    If Abs(lx - mgX(i)) <= mHalfSize Then
       If Abs(ly - mgY(i)) <= mHalfSize Then
          mHotSpotInd = i
          mCurrX = Format$(mDataX(i), "0.0000")     'change current coordinates to
          mCurrY = Format$(mDataY(i), "0.0000")     'exact data location if hot spot
          Exit Sub
       End If
    End If
Next i
mHotSpotInd = -1
End Sub


Public Function GetSelection(Sel() As Long) As Long
'----------------------------------------------------
'fills array Sel with ids of currently selected spots
'returns number of spots; -1 on any error
'----------------------------------------------------
Dim SelCnt As Long
Dim i As Long
On Error GoTo err_GetSelection
ReDim Sel(mDataCnt - 1)
For i = 0 To mDataCnt - 1
    If mDataSelected(i) Then
       SelCnt = SelCnt + 1
       Sel(SelCnt - 1) = mDataID(i)
    End If
Next i
If SelCnt > 0 Then
   If SelCnt < mDataCnt - 1 Then ReDim Preserve Sel(SelCnt - 1)
Else
   Erase Sel
End If
GetSelection = SelCnt
Exit Function

err_GetSelection:
Erase Sel
GetSelection = -1
End Function


Public Function ToggleSpotSelection(ByVal Ind As Long) As Boolean
'------------------------------------------------------------------------
'toggles spot selection and returns current state of selection for a spot
'------------------------------------------------------------------------
On Error Resume Next
mDataSelected(Ind) = Not mDataSelected(Ind)
SpotsRefresh
End Function

Public Property Get ViewWindow() As sSpotsView
ViewWindow = mViewWindow
End Property

Public Property Let ViewWindow(ByVal NewViewWindow As sSpotsView)
mViewWindow = NewViewWindow
SpotsRefresh
End Property

Public Sub SetFixedWindow(ByVal XMin As Double, ByVal XMax As Double, _
                          ByVal YMin As Double, ByVal YMax As Double)
'----------------------------------------------------------------------
'sets fixed window coordinates to be used for coordinate system
'----------------------------------------------------------------------
mFixedMinX = XMin:  mFixedMaxX = XMax
mFixedMinY = YMin:  mFixedMaxY = YMax
End Sub


Private Sub GetStarPoints(ByVal x As Long, ByVal y As Long, StarPoints() As Long)
Dim i As Long
On Error Resume Next
ReDim StarPoints(7, 1)
For i = 0 To 7
    Select Case i
    Case 0, 4
        StarPoints(i, 0) = x
    Case 1, 3
        StarPoints(i, 0) = x - mHalfSize \ 4
    Case 5, 7
        StarPoints(i, 0) = x + mHalfSize \ 4
    Case 2
        StarPoints(i, 0) = x - mHalfSize
    Case 6
        StarPoints(i, 0) = x + mHalfSize
    End Select
    Select Case i
    Case 2, 6
        StarPoints(i, 1) = y
    Case 1, 7
        StarPoints(i, 1) = y - CLng(mHalfSize / 4)
    Case 3, 5
        StarPoints(i, 1) = y + CLng(mHalfSize / 4)
    Case 0
        StarPoints(i, 1) = y - CLng(mHalfSize)
    Case 4
        StarPoints(i, 1) = y + CLng(mHalfSize)
    End Select
Next i
End Sub


Public Property Get SpotSize() As Long
SpotSize = mSpotSize
End Property

Public Property Let SpotSize(ByVal NewSize As Long)
mSpotSize = NewSize
mHalfSize = mSpotSize / 2
End Property
