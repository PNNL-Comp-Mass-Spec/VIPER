VERSION 5.00
Begin VB.UserControl La2DGrid 
   ClientHeight    =   5310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7050
   ScaleHeight     =   5310
   ScaleWidth      =   7050
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
      Height          =   5055
      Left            =   0
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   461
      TabIndex        =   0
      Top             =   120
      Width           =   6975
   End
   Begin VB.Menu mnuF 
      Caption         =   "Function"
      Visible         =   0   'False
      Begin VB.Menu mnuFViewCoo 
         Caption         =   "View &Coordinates"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFViewGrid 
         Caption         =   "View &Grid"
      End
      Begin VB.Menu mnuFSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFSM 
         Caption         =   "Selection Mode"
         Begin VB.Menu mnuFSMode 
            Caption         =   "Toggle"
            Index           =   0
         End
         Begin VB.Menu mnuFSMode 
            Caption         =   "Select Not selected"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuFSMode 
            Caption         =   "Unselect Selected"
            Index           =   2
         End
      End
      Begin VB.Menu mnuFS 
         Caption         =   "Select"
         Begin VB.Menu mnuFSSpecial 
            Caption         =   "All"
            Index           =   0
         End
         Begin VB.Menu mnuFSSpecial 
            Caption         =   "Checkers 1"
            Index           =   1
         End
         Begin VB.Menu mnuFSSpecial 
            Caption         =   "Checkers 2"
            Index           =   2
         End
         Begin VB.Menu mnuFSSpecial 
            Caption         =   "Positive Bins"
            Index           =   3
         End
         Begin VB.Menu mnuFSSpecial 
            Caption         =   "Fill"
            Index           =   4
         End
         Begin VB.Menu mnuFSSpecial 
            Caption         =   "Invert Selection"
            Index           =   5
         End
         Begin VB.Menu mnuFSSpecial 
            Caption         =   "Random Selection"
            Index           =   6
         End
      End
      Begin VB.Menu mnuFClearSelection 
         Caption         =   "Cl&ear Selection"
      End
      Begin VB.Menu mnuFSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDFCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "La2DGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'discrete 2D chart user control
'---------------------------------------------------------------------------
'last modified: 11/04/2002 nt
'---------------------------------------------------------------------------
Option Explicit

Const MSG_TITLE = "2D Displays - MW/ET GRID"

Public Enum MoveDirection       'describes basic movement in 2D grid
    mdNone = 0
    mdLeft = 1
    mdRight = 2
    mdUp = 3
    mdDown = 4
    mdLeftUp = 5                'diagonal movements
    mdLeftDown = 6
    mdRightUp = 7
    mdRightDown = 8
End Enum

Public Enum AddDataResult
    AddDataError = -1
    AddDataOutOfScope = 0
    AddDataOK = 1
End Enum

Public Enum ColorScale
    ColorScaleLin = 0
    ColorScaleLog = 1
End Enum

Public Enum SelectMode
    SelectModeToggle = 0            'toggle selection state of the 2D bin
    SelectModeSelect = 1            'select non-selected 2D bins
    SelectModeDeselect = 2          'deselect selected 2D bins
End Enum

Const MAX_X_RES = 10000
Const MAX_Y_RES = 10000
Const MAX_CLR_RES = 256

'LOGICAL COORDINATE SYSTEM CONSTANTS-SIZE
Const LgX0 = 0          'logical window
Const LgY0 = 0          '(X0,Y0)-(XE,YE)
Const LgXE = 10000
Const LgYE = 10000
'LOGICAL COORDINATE SYSTEM CONSTANTS-INDENTS
Const lgSXPct = 0.025
Const lgSYPct = 0.025
Const lgLXPct = 0.075
Const lgLYPct = 0.075
Const LgFntWPct = 0.005
Const LgFntHPct = 0.025
'LOGICAL COORDINATE SYSTEM CONSTANTS-OTHER
'Const LgWndW As Long = (LgXE - LgX0) / (1 - lgSXPct - lgLXPct)
'Const LgWndH As Long = (LgYE - LgY0) / (1 - lgSYPct - lgLYPct)
'''Const LgSX As Long = lgSXPct * LgWndW
'''Const LgSY As Long = lgSYPct * LgWndH
'''Const LgLX As Long = lgLXPct * LgWndW
'''Const LgLY As Long = lgLYPct * LgWndH

Const LgFntW As Long = LgXE * LgFntWPct
Const LgFntH As Long = LgYE * LgFntHPct

Dim mMinX As Double
Dim mMaxX As Double
Dim mMinY As Double
Dim mMaxY As Double

Dim mXRes As Long
Dim mYRes As Long
Dim mClrRes As Long

Dim mXStep As Double            'real width of horizontal bin
Dim mYStep As Double            'real width of vertical bin

Dim mLXStep As Long             'logical steps used to draw grid
Dim mLYStep As Long

Dim mDataCnt As Long            'total number of data points
Dim mDataVal As Double          'total value of data points
Dim mGridVal() As Double        'value of each 2D bin
Dim mGridClrInd() As Integer    '2D bins color index
Dim mGridSel() As Boolean       '2D bins selection state

Dim mBackColor As Long            'background color
Dim mForeColor As Long            'foreground color
Dim mValColor As Long             'color to shade to mark 2D bin value
Dim mSelColor As Long             'selection color

'public properties
Public DisplayInd As Long           'index of display to which this grid belongs
Public SelMode As SelectMode        'selection mode

Public VLabel As String     'label on vertical axis
Public VNumFmt As String    'numerical format for vertical axis
Public HLabel As String     'label on horizontal axis
Public HNumFmt As String    'numerical format for horizontal axis

Dim mViewCoordinates As Boolean
Dim mViewGrid As Boolean

Dim mXHotInd As Long       'index of hot bin; -1 if none
Dim mYHotInd As Long       'index of hot bin; -1 if none

'coordinate system-viewport coordinates
Dim VPX0 As Long
Dim VPY0 As Long
Dim VPXE As Long
Dim VPYE As Long

Dim mMinGridVal As Double      'have to track minimum and maximum values for
Dim mMaxGridVal As Double      'fast color assignments

Dim hValBrush() As Long        'brushes to indicate 2D bin value
Dim hSelBrush As Long          'selected brush
Dim hBackBrush As Long         'background color brush
Dim hForePen As Long           'foreground color pen
Dim hBackPen As Long           'background color pen
Dim paPoints() As POINTAPI     'used to calculate dev to log coordinate conversion

Dim StopSelPlay As Boolean


Public Function AddData(x As Double, y As Double, ByVal XYVal As Double) As Long
'----------------------------------------------------------------------------------
'adds data point with coordinate X, Y to the chart; coordinates of the bins where
'spot belongs are calculated and value of the spot is added to the total  bin value
'returns 1 if OK, 0 if data is out of current boundaries; -1 on any error
'----------------------------------------------------------------------------------
Dim CurrXInd As Long
Dim CurrYInd As Long
On Error GoTo err_Adddata

If x = mMaxX Then
   CurrXInd = mXRes - 1
Else
   CurrXInd = Int((x - mMinX) / mXStep)
End If
If y = mMaxY Then
   CurrYInd = mYRes - 1
Else
   CurrYInd = Int((y - mMinY) / mYStep)
End If

If CurrXInd < 0 Or CurrXInd > mXRes - 1 Or CurrYInd < 0 Or CurrYInd > mYRes - 1 Then
   AddData = AddDataOutOfScope
Else
   mGridVal(CurrXInd, CurrYInd) = mGridVal(CurrXInd, CurrYInd) + XYVal
   mDataCnt = mDataCnt + 1
   mDataVal = mDataVal + XYVal
   AddData = AddDataOK
End If
Exit Function

err_Adddata:
AddData = AddDataError
End Function

Private Sub mnuFClearSelection_Click()
On Error Resume Next
ReDim mGridSel(mXRes - 1, mYRes - 1)
Call Refresh
End Sub

Private Sub mnuFSelectAll_Click()
Dim i As Long, j As Long
For i = 0 To mXRes - 1
    For j = 0 To mYRes - 1
        mGridSel(i, j) = True
    Next j
Next i
Call Refresh
End Sub

Private Sub mnuFSMode_Click(Index As Integer)
Dim i As Long
On Error Resume Next
For i = 0 To mnuFSMode.Count - 1
    If i = Index Then
       mnuFSMode(i).Checked = True
    Else
       mnuFSMode(i).Checked = False
    End If
Next i
SelMode = Index
End Sub

Private Sub mnuFSSpecial_Click(Index As Integer)
Dim i As Long
Dim j As Long
On Error Resume Next
Select Case Index
Case 0              'select all
    For i = 0 To mXRes - 1
        For j = 0 To mYRes - 1
            mGridSel(i, j) = True
        Next j
    Next i
Case 1              'checkers 1
    For i = 0 To mXRes - 1
        For j = 0 To mYRes - 1
            If (i + j) Mod 2 = 0 Then
               mGridSel(i, j) = True
            Else
               mGridSel(i, j) = False
            End If
        Next j
    Next i
Case 2              'checkers 2
    For i = 0 To mXRes - 1
        For j = 0 To mYRes - 1
            If (i + j) Mod 2 = 0 Then
               mGridSel(i, j) = False
            Else
               mGridSel(i, j) = True
            End If
        Next j
    Next i
Case 3             'select all bins with positive values
    For i = 0 To mXRes - 1
        For j = 0 To mYRes - 1
            If mGridVal(i, j) > 0 Then
               mGridSel(i, j) = True
            Else
               mGridSel(i, j) = False
            End If
        Next j
    Next i
Case 4             'select or unselect by filling the regions of selection
    If mXHotInd >= 0 And mYHotInd >= 0 Then
       If SelMode = SelectModeToggle Then
          MsgBox "Fill option not applicable on Toggle selection mode!", vbOKOnly, MSG_TITLE
       Else
          Call SelectFill(mXHotInd, mYHotInd, mdNone)
       End If
    End If
Case 5
    For i = 0 To mXRes - 1
        For j = 0 To mYRes - 1
            Toggle2DBinSelection i, j
        Next j
    Next i
Case 6
    Call RandomSelection
End Select
Call Refresh
End Sub

Private Sub mnuFViewGrid_Click()
'-------------------------------------------------------
'toggle visibility of coordinates
'-------------------------------------------------------
mViewGrid = Not mViewGrid
mnuFViewGrid.Checked = mViewGrid
Call Refresh
End Sub

Private Sub UserControl_Initialize()
VLabel = "MW"
HLabel = "Scan"
HNumFmt = "0.00"
VNumFmt = "0"
mXRes = 16
mYRes = 32
mClrRes = 16
BackColor = vbWhite         'this call is property assignment which will
ForeColor = vbBlack         'trigger also creation of drawing objects
ValueColor = vbBlue
SelectionColor = vbRed
mViewCoordinates = mnuFViewCoo.Checked
mViewGrid = mnuFViewGrid.Checked
mMinX = 0:      mMaxX = 9889.236
mMinY = 0:      mMaxY = 9889.256
mXHotInd = -1:  mYHotInd = -1
SelMode = SelectModeSelect
End Sub

Private Sub UserControl_Resize()
picD.width = UserControl.ScaleWidth
picD.Height = UserControl.ScaleHeight
End Sub

Private Sub UserControl_Terminate()
DestroyDrawingObjects
End Sub

'---------------------------------------------------------------------------
'resolution properties
'---------------------------------------------------------------------------
Public Property Get XRes() As Long
XRes = mXRes
End Property

Public Property Let XRes(ByVal NewXRes As Long)
If NewXRes <= MAX_X_RES Then
   mXRes = NewXRes
Else
   MsgBox "Maximum X resolution " & MAX_X_RES & ".", vbOKOnly, MSG_TITLE
End If
End Property

Public Property Get YRes() As Long
YRes = mYRes
End Property

Public Property Let YRes(ByVal NewYRes As Long)
If NewYRes <= MAX_Y_RES Then
   mYRes = NewYRes
Else
   MsgBox "Maximum Y resolution " & MAX_Y_RES & ".", vbOKOnly, MSG_TITLE
End If
End Property

Public Property Get COLORRES() As Long
COLORRES = mClrRes
End Property

Public Property Let COLORRES(ByVal NewClrRes As Long)
If NewClrRes <= MAX_CLR_RES Then
   mClrRes = NewClrRes
Else
   MsgBox "Maximum color resolution " & MAX_CLR_RES & ".", vbOKOnly, MSG_TITLE
End If
End Property
'----------------------------------------------End Resolution Properties


'-----------------------------------------------------------------------
'properties dealing with boundaries
'-----------------------------------------------------------------------
Public Property Get MinX() As Double
MinX = mMinX
End Property

Public Property Let MinX(ByVal NewMinX As Double)
mMinX = NewMinX
End Property

Public Property Get MaxX() As Double
MaxX = mMaxX
End Property

Public Property Let MaxX(ByVal NewMaxX As Double)
mMaxX = NewMaxX
End Property

Public Property Get MinY() As Double
MinY = mMinY
End Property

Public Property Let MinY(ByVal NewMinY As Double)
mMinY = NewMinY
End Property

Public Property Get MaxY() As Double
MaxY = mMaxY
End Property

Public Property Let MaxY(ByVal NewMaxY As Double)
mMaxY = NewMaxY
End Property
'-------------------------------------------------------End Boundary Properties

Public Property Get DataCount() As Long
DataCount = mDataCnt
End Property

Public Property Get DataValue() As Double
DataValue = mDataVal
End Property


'----------------------------------------------------------------
'color public properties
'----------------------------------------------------------------
Public Property Get BackColor() As Long
BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal NewBackColor As Long)
On Error Resume Next
mBackColor = NewBackColor
If hBackBrush <> 0 Then DeleteObject (hBackBrush)
hBackBrush = CreateSolidBrush(mBackColor)
If hBackPen <> 0 Then DeleteObject (hBackPen)
hBackPen = CreatePen(PS_SOLID, 1, mBackColor)
ValueColor = mValColor         'this will trigger recalculation of value color
                               'necessary since value color depends on back color
End Property

Public Property Get ForeColor() As Long
ForeColor = mForeColor
End Property

Public Property Let ForeColor(ByVal NewForeColor As Long)
On Error Resume Next
mForeColor = NewForeColor
If hForePen <> 0 Then DeleteObject (hForePen)
hForePen = CreatePen(PS_SOLID, 1, mForeColor)
Call Refresh
End Property

Public Property Get SelectionColor() As Long
SelectionColor = mSelColor
End Property

Public Property Let SelectionColor(ByVal NewSelectionColor As Long)
On Error Resume Next
mSelColor = NewSelectionColor
If hSelBrush <> 0 Then DeleteObject (hSelBrush)
hSelBrush = CreateSolidBrush(mSelColor)
Call Refresh
End Property

Public Property Get ValueColor() As Long
ValueColor = mValColor
End Property

Public Property Let ValueColor(ByVal NewValueColor As Long)
Dim i As Long
Dim dClrStep As Double
Dim lClr As Long
On Error Resume Next
mValColor = NewValueColor
dClrStep = MAX_CLR_RES / mClrRes
For i = 0 To mClrRes - 1
    If hValBrush(i) <> 0 Then DeleteObject (hValBrush(i))
Next i
ReDim hValBrush(mClrRes - 1)
For i = 0 To mClrRes - 1
    lClr = (&H10101 * (Int(dClrStep * i) + 1)) And (mValColor Xor mBackColor) Xor mBackColor
    hValBrush(i) = CreateSolidBrush(lClr)
Next i
Call Refresh
End Property
'----------------------------------------------------------end color properties


Public Sub InitData()
'----------------------------------------------------
'initializes data structures
'----------------------------------------------------
ReDim mGridVal(mXRes - 1, mYRes - 1)
ReDim mGridClrInd(mXRes - 1, mYRes - 1)
ReDim mGridSel(mXRes - 1, mYRes - 1)
mDataCnt = 0
mDataVal = 0
mXStep = (mMaxX - mMinX) / mXRes
mYStep = (mMaxY - mMinY) / mYRes
mLXStep = Int((LgXE - LgX0) / mXRes)
mLYStep = Int((LgYE - LgY0) / mYRes)
End Sub

Public Sub CalcColors(ByVal ColorScaleType As ColorScale)
'--------------------------------------------------
'calculate color indexes for all 2DBins
'--------------------------------------------------
Dim i As Long, j As Long
Dim ValStep As Double
Call GetValueRange
Select Case ColorScaleType
Case ColorScaleLin
     ValStep = (mMaxGridVal - mMinGridVal) / mClrRes
     If ValStep = 0 Then ValStep = 0.01
     For i = 0 To mXRes - 1
         For j = 0 To mYRes - 1
             mGridClrInd(i, j) = Int((mGridVal(i, j) - mMinGridVal) / ValStep)
             If mGridClrInd(i, j) >= mClrRes Then mGridClrInd(i, j) = mClrRes - 1
         Next j
     Next i
Case ColorScaleLog
End Select
End Sub

Private Sub DrawGrid(hDC As Long)
Dim ptPoint As POINTAPI
Dim Res As Long
Dim OldPen As Long
Dim i As Long
Dim CurrOffset As Long
On Error Resume Next
OldPen = SelectObject(hDC, hForePen)
For i = 1 To mYRes                     'horizontal
    CurrOffset = i * mLYStep
    Res = MoveToEx(hDC, LgX0, LgY0 + CurrOffset, ptPoint)
    Res = LineTo(hDC, LgX0 + LgXE, LgY0 + CurrOffset)
Next i
For i = 1 To mXRes                     'vertical
    CurrOffset = i * mLXStep
    Res = MoveToEx(hDC, LgX0 + CurrOffset, LgY0, ptPoint)
    Res = LineTo(hDC, LgX0 + CurrOffset, LgY0 + LgYE)
Next i
Res = SelectObject(hDC, OldPen)
End Sub

Private Sub Draw2DBin(hDC As Long, i As Long, j As Long)
'-------------------------------------------------------
'draws (i,j) 2D bin to device context in Value color
'-------------------------------------------------------
Dim OldBrush As Long
Dim Res As Long
On Error Resume Next
OldBrush = SelectObject(hDC, hValBrush(mGridClrInd(i, j)))
Res = PatBlt(hDC, LgX0 + i * mLXStep, LgY0 + j * mLYStep, mLXStep, mLYStep, PATCOPY)
Call SelectObject(hDC, OldBrush)
End Sub


Private Sub Draw2DBinComplete(i As Long, j As Long)
'----------------------------------------------------------------------
'draws (i,j) 2D bin to the screen(no need to use this on anything else)
'----------------------------------------------------------------------
Dim OldDC As Long
Dim OldBrush As Long
Dim Res As Long
On Error Resume Next
With picD
    OldDC = SaveDC(.hDC)
    OldBrush = SelectObject(.hDC, hValBrush(mGridClrInd(i, j)))
    DGCooSys .hDC, .ScaleWidth, .ScaleHeight
    Res = PatBlt(.hDC, LgX0 + i * mLXStep, LgY0 + j * mLYStep, mLXStep, mLYStep, PATCOPY)
    Call SelectObject(.hDC, OldBrush)
    If i = 0 Or j = 0 Then DGDrawCooSys .hDC        'have to redraw coordinate system
    If mViewGrid Then DrawGrid .hDC
    Res = RestoreDC(.hDC, OldDC)
End With
End Sub


Private Sub Draw2DBinBack(hDC As Long, i As Long, j As Long)
'-------------------------------------------------------
'draws (i,j) 2D bin to device context in Value color
'-------------------------------------------------------
Dim OldBrush As Long
Dim Res As Long
On Error Resume Next
OldBrush = SelectObject(hDC, hBackBrush)
Res = PatBlt(hDC, LgX0 + i * mLXStep, LgY0 + j * mLYStep, mLXStep, mLYStep, PATCOPY)
Call SelectObject(hDC, OldBrush)
End Sub


Private Sub Draw2DBinBackComplete(i As Long, j As Long)
'----------------------------------------------------------------------
'draws (i,j) 2D bin to the screen(no need to use this on anything else)
'----------------------------------------------------------------------
Dim OldDC As Long
Dim OldBrush As Long
Dim Res As Long
On Error Resume Next
With picD
    OldDC = SaveDC(.hDC)
    OldBrush = SelectObject(.hDC, hBackBrush)
    DGCooSys .hDC, .ScaleWidth, .ScaleHeight
    Res = PatBlt(.hDC, LgX0 + i * mLXStep, LgY0 + j * mLYStep, mLXStep, mLYStep, PATCOPY)
    Call SelectObject(.hDC, OldBrush)
    If i = 0 Or j = 0 Then DGDrawCooSys .hDC        'have to redraw coordinate system
    If mViewGrid Then DrawGrid .hDC
    Res = RestoreDC(.hDC, OldDC)
End With
End Sub


Private Sub Draw2DBinSel(hDC As Long, i As Long, j As Long)
'----------------------------------------------------------
'draws (i,j) 2D bin to device context in selection color
'----------------------------------------------------------
Dim OldBrush As Long
Dim Res As Long
On Error Resume Next
OldBrush = SelectObject(hDC, hSelBrush)
Res = PatBlt(hDC, LgX0 + i * mLXStep, LgY0 + j * mLYStep, mLXStep, mLYStep, PATCOPY)
Call SelectObject(hDC, OldBrush)
End Sub

Private Sub Draw2DBinSelComplete(i As Long, j As Long)
'----------------------------------------------------------------------
'draws (i,j) 2D bin to the screen(no need to use this on anything else)
'----------------------------------------------------------------------
Dim OldDC As Long
Dim OldBrush As Long
Dim Res As Long
On Error Resume Next
With picD
    OldDC = SaveDC(.hDC)
    OldBrush = SelectObject(.hDC, hSelBrush)
    DGCooSys .hDC, .ScaleWidth, .ScaleHeight
    Res = PatBlt(.hDC, LgX0 + i * mLXStep, LgY0 + j * mLYStep, mLXStep, mLYStep, PATCOPY)
    Call SelectObject(.hDC, OldBrush)
    If i = 0 Or j = 0 Then DGDrawCooSys .hDC        'have to redraw coordinate system
    If mViewGrid Then DrawGrid .hDC
    Res = RestoreDC(.hDC, OldDC)
End With
End Sub

Public Sub Clear()
'---------------------------------------------------
'erase arrays and clean picture
'---------------------------------------------------
On Error Resume Next
Erase mGridVal
Erase mGridClrInd
Erase mGridSel
mDataCnt = 0
mDataVal = 0
picD.Cls
End Sub

Public Sub Draw()
'-----------------------------------------------------
'draws spots on picture device context
'-----------------------------------------------------
Dim OldDC As Long
Dim Res As Long
Dim i As Long
Dim j As Long
OldDC = SaveDC(picD.hDC)
DrawBackColor picD.hDC
DGCooSys picD.hDC, picD.ScaleWidth, picD.ScaleHeight
WriteLabels picD.hDC
DrawDataColors picD.hDC
If mDataCnt > 0 Then
   For i = 0 To mXRes - 1
       For j = 0 To mYRes - 1
           If mGridSel(i, j) Then
              Call Draw2DBinSel(picD.hDC, i, j)
           Else
              If mGridClrInd(i, j) >= 0 Then Call Draw2DBin(picD.hDC, i, j)
           End If
       Next j
   Next i
End If
'draw lines at the end so they are visible
DGDrawCooSys picD.hDC
If mViewGrid Then DrawGrid picD.hDC
Res = RestoreDC(picD.hDC, OldDC)
End Sub


Private Sub DGCooSys(ByVal hDC As Long, ByVal dcw As Long, ByVal dch As Long)
'----------------------------------------------------------------------------
'establishes coordinate system on device context; this coordinate system can
'only have origin in the bottom left of the logical window
'----------------------------------------------------------------------------
Dim Res As Long
Dim ptPoint As POINTAPI
Dim szSize As Size
On Error Resume Next
VPX0 = CLng(dcw * lgLXPct / 2)
VPY0 = CLng(dch * (1 - 3 * lgLYPct / 4))
VPXE = CLng(dcw * (1 - lgLXPct))
VPYE = CLng((4 * lgLYPct / 3 - 1) * dch)
Res = SetMapMode(hDC, MM_ANISOTROPIC)
'logical window
Res = SetWindowOrgEx(hDC, LgX0, LgY0, ptPoint)
Res = SetWindowExtEx(hDC, LgXE, LgYE, szSize)
'viewport
Res = SetViewportOrgEx(hDC, VPX0, VPY0, ptPoint)
Res = SetViewportExtEx(hDC, VPXE, VPYE, szSize)
End Sub


Public Sub DGDrawCooSys(ByVal hDC As Long)
'-----------------------------------------
'draws coordinate system on device context
'-----------------------------------------
Dim ptPoint As POINTAPI
Dim OldPen As Long
Dim Res As Long
On Error Resume Next
OldPen = SelectObject(hDC, hForePen)
'horizontal
Res = MoveToEx(hDC, LgX0, LgY0, ptPoint)
Res = LineTo(hDC, LgX0 + LgXE, LgY0)
'vertical
Res = MoveToEx(hDC, LgX0, LgY0, ptPoint)
Res = LineTo(hDC, LgX0, LgY0 + LgYE)
Res = SelectObject(hDC, OldPen)
End Sub


Private Sub mnuDFCopy_Click()
Call CopyFD
End Sub


Private Sub mnuFViewCoo_Click()
'-------------------------------------------------------
'toggle visibility of coordinates
'-------------------------------------------------------
mViewCoordinates = Not mViewCoordinates
mnuFViewCoo.Checked = mViewCoordinates
If Not mViewCoordinates Then WriteCoordinates picD.hDC, picD.ScaleWidth, picD.ScaleHeight  'this will clear coordinates
End Sub


Private Sub picD_Click()
'select 2DBin here
On Error Resume Next
If (mXHotInd >= 0 And mYHotInd >= 0) Then Toggle2DBinSelection mXHotInd, mYHotInd
End Sub


Private Sub picD_KeyDown(KeyCode As Integer, Shift As Integer)
'-------------------------------------------------------------
'copy as metafile on system clipboard on Ctrl+C combination
'-------------------------------------------------------------
Dim CtrlDown As Boolean
CtrlDown = (Shift And vbCtrlMask) > 0
Select Case KeyCode
Case vbKeyC
     If CtrlDown Then CopyFD
Case vbKeyD
     If CtrlDown Then StopSelPlay = True
Case vbKeyS
     If CtrlDown Then SelectionPlay
End Select
End Sub


Private Sub picD_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then UserControl.PopupMenu mnuF
End Sub


Private Sub picD_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If mDataCnt > 0 Then
   ReDim paPoints(0)
   paPoints(0).x = x
   paPoints(0).y = y
   DevLogConversion 0, 1
   TrackHotSpot paPoints(0).x, paPoints(0).y
   If Button = vbLeftButton Then
      If mXHotInd >= 0 And mYHotInd >= 0 Then
         Select Case SelMode
         Case SelectModeToggle
            Toggle2DBinSelection mXHotInd, mYHotInd
         Case SelectModeSelect
            If Not mGridSel(mXHotInd, mYHotInd) Then
               Toggle2DBinSelection mXHotInd, mYHotInd
            End If
         Case SelectModeDeselect
            If mGridSel(mXHotInd, mYHotInd) Then
               Toggle2DBinSelection mXHotInd, mYHotInd
            End If
         End Select
      End If
   End If
   If mViewCoordinates Then WriteCoordinates picD.hDC, picD.ScaleWidth, picD.ScaleHeight
End If
End Sub


Private Sub picD_Paint()
Call Refresh
End Sub


Private Sub WriteLabels(ByVal hDC As Long)
Dim lfLogFont As LOGFONT
Dim lOldFont As Long
Dim lNewFont As Long
Dim lOldPen As Long
Dim lFont As Long
Dim Res As Long
On Error Resume Next
lOldPen = SelectObject(hDC, hForePen)
'get the font from the picture box control (Arial Narrow)
lFont = SelectObject(hDC, GetStockObject(SYSTEM_FONT))
Res = GetObjectAPI(lFont, Len(lfLogFont), lfLogFont)
Res = SelectObject(hDC, lFont)
'create new logical font
lfLogFont.lfWidth = LgFntW
lfLogFont.lfHeight = LgFntH
lNewFont = CreateFontIndirect(lfLogFont)
'select newly created logical font to DC
lOldFont = SelectObject(hDC, lNewFont)
'draw coordinate axes labels
Res = TextOut(hDC, (1 - 0.03) * LgXE, -0.03 * LgYE, HLabel, Len(HLabel))
Res = TextOut(hDC, LgX0 + 0.01 * LgXE, (1 + 0.03) * LgYE, VLabel, Len(VLabel))
'draw numeric markers on vertical axes
WriteVMarkers hDC
WriteHMarkers hDC
'restore old font
Res = SelectObject(hDC, lOldFont)
DeleteObject (lNewFont)
Res = SelectObject(hDC, lOldPen)
End Sub


Private Sub WriteVMarkers(ByVal hDC As Long)
Dim VDelta As Double
Dim LDelta As Long
Dim VLbl As String
Dim szLbl As Size
Dim ptPoint As POINTAPI
Dim MarkSize As Long
Dim Res As Long
Dim i As Long
VDelta = (mMaxY - mMinY) / 3
LDelta = LgYE / 3
MarkSize = CLng(0.01 * LgXE)
For i = 0 To 4
    VLbl = Format$(mMinY + i * VDelta, VNumFmt)
    Res = GetTextExtentPoint32(hDC, VLbl, Len(VLbl), szLbl)
    Res = TextOut(hDC, -szLbl.cx - MarkSize, i * LDelta + 4 * MarkSize, VLbl, Len(VLbl))
    If Not mViewGrid Then
       Res = MoveToEx(hDC, -MarkSize \ 2, i * LDelta, ptPoint)
       Res = LineTo(hDC, MarkSize \ 2, i * LDelta)
    End If
Next i
'write also number of 2D bins and data points
VLbl = "2D bins count: " & (mXRes * mYRes)
Res = GetTextExtentPoint32(hDC, VLbl, Len(VLbl), szLbl)
Res = TextOut(hDC, LgXE * 0.05, (1 + 0.035) * LgYE, VLbl, Len(VLbl))
VLbl = "Data count: " & mDataCnt
Res = GetTextExtentPoint32(hDC, VLbl, Len(VLbl), szLbl)
Res = TextOut(hDC, LgXE * 0.2, (1 + 0.035) * LgYE, VLbl, Len(VLbl))
End Sub


Private Sub WriteHMarkers(ByVal hDC As Long)
Dim HDelta As Double
Dim LDelta As Long
Dim HLbl As String
Dim szLbl As Size
Dim ptPoint As POINTAPI
Dim MarkSize As Long
Dim Res As Long
Dim i As Long
LDelta = LgXE / 3
HDelta = (mMaxX - mMinX) / 3
MarkSize = CLng(0.01 * LgXE)
For i = 0 To 4
   HLbl = Format(mMinX + i * HDelta, HNumFmt)
   Res = GetTextExtentPoint32(hDC, HLbl, Len(HLbl), szLbl)
   Res = TextOut(hDC, i * LDelta - szLbl.cy \ 2, -MarkSize, HLbl, Len(HLbl))
   If Not mViewGrid Then
      Res = MoveToEx(hDC, i * LDelta, -MarkSize \ 2, ptPoint)
      Res = LineTo(hDC, i * LDelta, MarkSize)
   End If
Next i
End Sub


Private Sub DrawDataColors(hDC As Long)
Dim lWdth As Long               'space where drawing will take place
Dim lW As Long
Dim lH As Long
Dim OldBrush As Long
Dim OldPen As Long
Dim lStartPosX As Long, lStartPosY As Long
Dim lStartX As Long, lStartY As Long
Dim MarkSize As Long
Dim i As Long
Dim Res As Long
Dim ptPoint As POINTAPI
On Error Resume Next
lWdth = CLng(0.25 * LgXE)         '25% of width of logical window
lW = CLng(lWdth / mClrRes)        'width of individual color box
lH = CLng(LgYE * 0.01)            '1% of height of logical window
lStartPosX = CLng(0.4 * LgXE)     '40% of width of logical window
lStartPosY = CLng((1 + 0.02) * LgYE)
MarkSize = CLng(0.01 * LgXE)
OldBrush = SelectObject(hDC, hValBrush(0))
OldPen = SelectObject(hDC, hForePen)
For i = 0 To mClrRes - 1
    Res = SelectObject(hDC, hValBrush(i))
    lStartX = lStartPosX + i * lW
    lStartY = lStartPosY + lH
    PatBlt hDC, lStartX, lStartPosY, lW, lH, PATCOPY
    Res = MoveToEx(hDC, lStartX, lStartY, ptPoint)
    Res = LineTo(hDC, lStartX, lStartY + MarkSize)
Next i
'draw first and last mark a bit different
Res = MoveToEx(hDC, lStartPosX, lStartPosY, ptPoint)
Res = LineTo(hDC, lStartPosX, lStartPosY + lH + MarkSize)
Res = MoveToEx(hDC, lStartX + lW, lStartPosY, ptPoint)
Res = LineTo(hDC, lStartX + lW, lStartPosY + lH + MarkSize)
Res = SelectObject(hDC, OldPen)
Res = SelectObject(hDC, OldBrush)
End Sub


Private Sub WriteCoordinates(hDC As Long, w As Long, h As Long)
Dim OldDC As Long
Dim lfLogFont As LOGFONT
Dim lOldFont As Long
Dim lNewFont As Long
Dim Lbl As String
Dim lFont As Long
Dim szLbl As Size
Dim OldPen As Long
Dim OldBrush As Long
Dim Res As Long
On Error Resume Next
OldDC = SaveDC(hDC)
DGCooSys hDC, w, h
'get the font from the picture box control (Arial Narrow)
lFont = SelectObject(hDC, GetStockObject(SYSTEM_FONT))
Res = GetObjectAPI(lFont, Len(lfLogFont), lfLogFont)
Res = SelectObject(hDC, lFont)
'create new logical font
lfLogFont.lfWidth = LgFntW
lfLogFont.lfHeight = LgFntH
lNewFont = CreateFontIndirect(lfLogFont)
'select newly created logical font to DC
lOldFont = SelectObject(hDC, lNewFont)
'write new coordinates
Lbl = "2DBin: (" & mXHotInd & ",  " & mYHotInd & ") Value: " & mGridVal(mXHotInd, mYHotInd) & " Clr.Ind: " & mGridClrInd(mXHotInd, mYHotInd)
Res = GetTextExtentPoint32(hDC, Lbl, Len(Lbl), szLbl)
'clear whatever was written before; just draw regular rectangle
OldPen = SelectObject(hDC, hBackPen)             'invisible pen
OldBrush = SelectObject(hDC, hBackBrush)
Rectangle hDC, LgXE * 0.7, (1 + 0.035) * LgYE, (1 + 0.1) * LgXE, (1 + 0.035) * LgYE - szLbl.cy
If mViewCoordinates Then
   If mXHotInd >= 0 And mYHotInd >= 0 Then
      Res = TextOut(hDC, LgXE * 0.7, (1 + 0.035) * LgYE, Lbl, Len(Lbl))
   End If
End If
Res = SelectObject(hDC, OldPen)
Res = SelectObject(hDC, OldBrush)
'restore old font
Res = SelectObject(hDC, lOldFont)
DeleteObject (lNewFont)
Res = RestoreDC(hDC, OldDC)
End Sub


Public Sub Refresh()
'-------------------------------------------
'rebuild coordinate system and redraws spots
'-------------------------------------------
picD.Cls
Call Draw
End Sub


Public Sub CopyFD()
'-----------------------------------------------------------
'draw current display to the enhanced metafile and then copy
'it to the clipboard
'-----------------------------------------------------------
Dim Res As Long

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
Dim rcRef As Rect           'reference rectangle
Dim i As Long, j As Long
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
DrawBackColor emfDC
DGCooSys emfDC, picD.ScaleWidth, picD.ScaleHeight
WriteLabels emfDC
DrawDataColors emfDC
If mDataCnt > 0 Then
   For i = 0 To mXRes - 1
       For j = 0 To mYRes - 1
           If mGridSel(i, j) Then
              Call Draw2DBinSel(emfDC, i, j)
           Else
              If mGridClrInd(i, j) >= 0 Then Call Draw2DBin(emfDC, i, j)
           End If
       Next j
   Next i
End If
DGDrawCooSys emfDC
If mViewGrid Then DrawGrid emfDC
If mViewCoordinates Then
   If mXHotInd >= 0 And mYHotInd >= 0 Then
      WriteCoordinates emfDC, picD.ScaleWidth, picD.ScaleHeight
   End If
End If
Res = RestoreDC(emfDC, OldDC)
emfHandle = CloseEnhMetaFile(emfDC)

Res = OpenClipboard(picD.hwnd)
Res = EmptyClipboard()
Res = SetClipboardData(CF_ENHMETAFILE, emfHandle)
Res = CloseClipboard
End Sub


Public Function WriteSToFile(ByVal FileName As String) As Boolean
'----------------------------------------------------------------
'write bins results to text semicolon delimited file
'returns True if successful, False if not or any error occurs
'----------------------------------------------------------------
Dim i As Long, j As Long
Dim hfile As Integer
On Error GoTo exit_WriteSToFile

hfile = FreeFile
Open FileName For Append As hfile
Print #hfile, "Presenting - " & HLabel & ", " & VLabel
Print #hfile, "Range - [" & mMinX & ", " & mMaxX & "]X[" & mMinY & ", " & mMaxY & "]"
Print #hfile, "X Resolution: " & mXRes & " - Y Resolution: " & mYRes
Print #hfile, vbCrLf
Print #hfile, "Min X Range;Max X Range;Min Y Range;Max Y Range; Value"
For i = 0 To mXRes - 1
  For j = 0 To mYRes - 1
    Print #hfile, mMinX + i * mXStep & ";" & mMinX + (i + 1) * mXStep & ";" _
       & mMinY + i * mYStep & ";" & mMinY + (i + 1) * mYStep & ";" & mGridVal(i, j)
  Next j
Next i
Close hfile
WriteSToFile = True
exit_WriteSToFile:
End Function

Public Sub DestroyDrawingObjects()
Dim i As Long
On Error Resume Next
For i = 0 To mClrRes - 1
    If hValBrush(i) <> 0 Then DeleteObject (hValBrush(i))
Next i
Erase hValBrush
If hSelBrush <> 0 Then DeleteObject (hSelBrush)
If hBackBrush <> 0 Then DeleteObject (hBackBrush)
If hForePen <> 0 Then DeleteObject (hForePen)
If hBackPen <> 0 Then DeleteObject (hBackPen)
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
On Error Resume Next
mXHotInd = Int(lx / mLXStep)
If mXHotInd < 0 Or mXHotInd > mXRes - 1 Then mXHotInd = -1
mYHotInd = Int(ly / mLYStep)
If mYHotInd < 0 Or mYHotInd > mYRes - 1 Then mYHotInd = -1
End Sub

Public Sub DrawBackColor(ByVal hDC As Long)
Dim OldBrush As Long
Dim Res As Long
On Error Resume Next
OldBrush = SelectObject(hDC, hBackBrush)
Res = PatBlt(hDC, 0, 0, picD.ScaleWidth, picD.ScaleWidth, PATCOPY)
Res = SelectObject(hDC, OldBrush)
End Sub

Public Function Toggle2DBinSelection(ByVal XInd As Long, ByVal YInd As Long) As Boolean
'--------------------------------------------------------------------------------------
'toggles 2D bin selection and returns current state of selection for a 2D bin
'--------------------------------------------------------------------------------------
On Error Resume Next
mGridSel(XInd, YInd) = Not mGridSel(XInd, YInd)
If mGridSel(XInd, YInd) Then
   Draw2DBinSelComplete XInd, YInd
Else
   If mGridClrInd(XInd, YInd) < 0 Then
      Draw2DBinBackComplete XInd, YInd
   Else
      Draw2DBinComplete XInd, YInd
   End If
End If
End Function

Public Sub GetValueRange()
'---------------------------------------------------------
'fills mMinGridVal and mMaxGridVal; this has to be done in
'separate procedure to make sure we visited all 2D bins
'---------------------------------------------------------
Dim i As Long, j As Long
mMinGridVal = glHugeOverExp
mMaxGridVal = -glHugeOverExp
For i = 0 To mXRes - 1
    For j = 0 To mYRes - 1
        If mGridVal(i, j) < mMinGridVal Then mMinGridVal = mGridVal(i, j)
        If mGridVal(i, j) > mMaxGridVal Then mMaxGridVal = mGridVal(i, j)
    Next j
Next i
End Sub


Public Function IsInSelection(ByVal x As Double, ByVal y As Double) As Boolean
'-----------------------------------------------------------------------------
'returns True if X, Y fits in current selection, False otherwise
'-----------------------------------------------------------------------------
Dim XInd As Long
Dim YInd As Long
On Error Resume Next

'figure out which 2D bin it belongs and check is it selected
If x = mMaxX Then
   XInd = mXRes - 1
Else
   XInd = Int((x - mMinX) / mXStep)
End If
If y = mMaxY Then
   YInd = mYRes - 1
Else
   YInd = Int((y - mMinY) / mYStep)
End If
If (XInd < 0 Or XInd > mXRes - 1 Or YInd < 0 Or YInd > mYRes - 1) Then
   Exit Function
Else
   IsInSelection = mGridSel(XInd, YInd)
End If
End Function


Private Sub SelectFill(ByVal Indx As Long, ByVal IndY As Long, ByVal MoveDir As MoveDirection)
On Error Resume Next
DoEvents
If StopSelPlay Then Exit Sub
Select Case SelMode
Case SelectModeToggle           'not applicable
Case SelectModeSelect
     If (Indx < 0 Or Indx > mXRes - 1 Or IndY < 0 Or IndY > mYRes - 1) Then Exit Sub
     If Not mGridSel(Indx, IndY) Then Toggle2DBinSelection Indx, IndY
     'do not go backward
     If Not (MoveDir = mdRight Or mGridSel(Indx - 1, IndY)) Then SelectFill Indx - 1, IndY, mdLeft
     If Not (MoveDir = mdLeft Or mGridSel(Indx + 1, IndY)) Then SelectFill Indx + 1, IndY, mdRight
     If Not (MoveDir = mdDown Or mGridSel(Indx, IndY + 1)) Then SelectFill Indx, IndY + 1, mdUp
     If Not (MoveDir = mdUp Or mGridSel(Indx, IndY - 1)) Then SelectFill Indx, IndY - 1, mdDown
Case SelectModeDeselect
     If (Indx < 0 Or Indx > mXRes - 1 Or IndY < 0 Or IndY > mYRes - 1) Then Exit Sub
     If mGridSel(Indx, IndY) Then Toggle2DBinSelection Indx, IndY
     'do not go backward
     If ((Not MoveDir = mdRight) And mGridSel(Indx - 1, IndY)) Then SelectFill Indx - 1, IndY, mdLeft
     If ((Not MoveDir = mdLeft) And mGridSel(Indx + 1, IndY)) Then SelectFill Indx + 1, IndY, mdRight
     If ((Not MoveDir = mdDown) And mGridSel(Indx, IndY + 1)) Then SelectFill Indx, IndY + 1, mdUp
     If ((Not MoveDir = mdUp) And mGridSel(Indx, IndY - 1)) Then SelectFill Indx, IndY - 1, mdDown
End Select
End Sub

Private Sub RandomSelection()
Dim i As Long, j As Long
Dim RandomDensity As Long
Dim RandomDirection As Long
Randomize
RandomDensity = 2 + CLng(Rnd() * 3)
RandomDirection = CLng(Rnd() * 100) Mod 4
Select Case RandomDirection
Case 0
  For i = 0 To mXRes - 1
      For j = 0 To mYRes - 1
        DoEvents
        If (i + j + CLng(Rnd() * 100)) Mod RandomDensity = 0 Then Toggle2DBinSelection i, j
        If StopSelPlay Then Exit Sub
      Next j
  Next i
Case 1
  For i = mXRes - 1 To 0 Step -1
      For j = mYRes - 1 To 0 Step -1
        DoEvents
        If (i + j + CLng(Rnd() * 100)) Mod RandomDensity = 0 Then Toggle2DBinSelection i, j
        If StopSelPlay Then Exit Sub
      Next j
  Next i
Case 2
  For j = 0 To mYRes - 1
      For i = 0 To mXRes - 1
        DoEvents
        If (i + j + CLng(Rnd() * 100)) Mod RandomDensity = 0 Then Toggle2DBinSelection i, j
        If StopSelPlay Then Exit Sub
      Next i
  Next j
Case 3
  For j = mYRes - 1 To 0 Step -1
      For i = mXRes - 1 To 0 Step -1
        DoEvents
        If (i + j + CLng(Rnd() * 100)) Mod RandomDensity = 0 Then Toggle2DBinSelection i, j
        If StopSelPlay Then Exit Sub
      Next i
  Next j
End Select
End Sub

Private Sub SelectionPlay()
Dim SUSRatio As Double
Dim PlaySectionDone As Boolean
Do Until StopSelPlay
   RandomSelection
   PlaySectionDone = False
   Do Until PlaySectionDone
      PlaySelect
      SUSRatio = GetSelUnSelRatio()
      If (SUSRatio < 0.2 Or SUSRatio > 5 Or StopSelPlay) Then PlaySectionDone = True
   Loop
Loop
StopSelPlay = False
End Sub

Private Sub PlaySelect()
Dim OldHotX As Long
Dim OldHotY As Long
Dim OldSelMode As Long
On Error GoTo exit_PlaySelect
OldHotX = mXHotInd
OldHotY = mYHotInd
OldSelMode = SelMode
mXHotInd = Int(Rnd() * (mXRes - 1))
mYHotInd = Int(Rnd() * (mYRes - 1))
If mGridSel(mXHotInd, mYHotInd) Then
   SelMode = SelectModeDeselect
Else
   SelMode = SelectModeSelect
End If
Call SelectFill(mXHotInd, mYHotInd, mdNone)
exit_PlaySelect:
mXHotInd = OldHotX
mYHotInd = OldHotY
SelMode = OldSelMode
End Sub


Private Function GetSelUnSelRatio() As Double
'-----------------------------------------------------------------
'return ratio of selected and unselected cells
'-----------------------------------------------------------------
Dim i As Long, j As Long
Dim SelCnt As Long, UnSelCnt As Long
For i = 0 To mXRes - 1
    For j = 0 To mYRes - 1
        If mGridSel(i, j) Then
            SelCnt = SelCnt + 1
        Else
            UnSelCnt = UnSelCnt + 1
        End If
    Next j
Next i
If UnSelCnt > 0 Then
   GetSelUnSelRatio = SelCnt / UnSelCnt
Else
   GetSelUnSelRatio = 1000
End If
End Function
