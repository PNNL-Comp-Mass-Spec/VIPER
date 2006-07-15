VERSION 5.00
Begin VB.UserControl LaSpotsRWC 
   ClientHeight    =   2475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4845
   ScaleHeight     =   2475
   ScaleWidth      =   4845
   Begin VB.PictureBox picD 
      BackColor       =   &H00FFFFFF&
      DrawMode        =   10  'Mask Pen
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      ScaleHeight     =   157
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   317
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
   Begin VB.Menu mnuF 
      Caption         =   "Function"
      Visible         =   0   'False
      Begin VB.Menu mnuFZoomOut 
         Caption         =   "Zoom Out"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuFZoomOut1 
         Caption         =   "Zoom Out 1 Level"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFViewNormal 
         Caption         =   "View Normal"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFViewCSMap 
         Caption         =   "View Charge State Map"
      End
      Begin VB.Menu mnuFSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFViewCoo 
         Caption         =   "View Coordinates"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFViewCSSpots 
         Caption         =   "View C.S. Spots"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFViewIsoSpots 
         Caption         =   "View Isotopic Spots"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFViewBoxes 
         Caption         =   "View Boxes"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFViewScanLines 
         Caption         =   "View Scan Lines"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFO 
         Caption         =   "Options"
         Begin VB.Menu mnuFOForeColor 
            Caption         =   "ForeColor"
         End
         Begin VB.Menu mnuFOBackColor 
            Caption         =   "BackColor"
         End
         Begin VB.Menu mnuFOCSColor 
            Caption         =   "Charge State Color"
         End
         Begin VB.Menu mnuFOIsoColor 
            Caption         =   "Isotopic Color"
         End
         Begin VB.Menu mnuFOCSS 
            Caption         =   "Charge State Shape"
            Begin VB.Menu mnuFOCSShape 
               Caption         =   "Ellipse"
               Index           =   0
            End
            Begin VB.Menu mnuFOCSShape 
               Caption         =   "Rectangle"
               Index           =   1
            End
            Begin VB.Menu mnuFOCSShape 
               Caption         =   "Round Rectangle"
               Index           =   2
            End
            Begin VB.Menu mnuFOCSShape 
               Caption         =   "Triangle"
               Index           =   3
            End
            Begin VB.Menu mnuFOCSShape 
               Caption         =   "Star"
               Index           =   4
            End
            Begin VB.Menu mnuFOCSShape 
               Caption         =   "Hexagon"
               Index           =   5
            End
            Begin VB.Menu mnuFOCSShape 
               Caption         =   "Gradient Fill Rectangle"
               Index           =   6
            End
         End
         Begin VB.Menu mnuFOIsoS 
            Caption         =   "Isotopic Shape"
            Begin VB.Menu mnuFOIsoShape 
               Caption         =   "Ellipse"
               Index           =   0
            End
            Begin VB.Menu mnuFOIsoShape 
               Caption         =   "Rectangle"
               Index           =   1
            End
            Begin VB.Menu mnuFOIsoShape 
               Caption         =   "Round Rectangle"
               Index           =   2
            End
            Begin VB.Menu mnuFOIsoShape 
               Caption         =   "Triangle"
               Index           =   3
            End
            Begin VB.Menu mnuFOIsoShape 
               Caption         =   "Star"
               Index           =   4
            End
            Begin VB.Menu mnuFOIsoShape 
               Caption         =   "Hexagon"
               Index           =   5
            End
            Begin VB.Menu mnuFOIsoShape 
               Caption         =   "Gradient Fill Rectangle"
               Index           =   6
            End
         End
      End
      Begin VB.Menu mnuFSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDFCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "LaSpotsRWC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'LaSpots Royal With Cheese
'---------------------------------------------------------------------------
'Used to present small portions of display in close-up;
'Data are loaded in arrays as spots and boxes; method View is then used
'to focus on any portion of loaded window, but not out of it
'---------------------------------------------------------------------------
'last modified: 04/08/2003 nt
'---------------------------------------------------------------------------
Option Explicit

Const MSG_TITLE = "2D Displays - La Boxed Spots Royal With Cheese"

Const L_X0 = 0
Const L_XE = 10000
Const L_Y0 = 0
Const L_YE = 10000
Const L_MIN_R = 100
Const L_MAX_R = 1000

Const ConvLPDP = 1
Const ConvDPLP = 2

Const MAXSPOTCOUNT = 1000              'maximum data points in this display
Const MAXBOXCOUNT = 1000               'maximum boxes in this view

'public properties
Dim mBackClr As Long
Dim mForeClr As Long
Dim mCSClr As Long
Dim mIsoClr As Long

Dim mCSShapeType As Long
Dim mIsoShapeType As Long

Public WithEvents MyCooSys As CooSysS
Attribute MyCooSys.VB_VarHelpID = -1

'following four variables bound the total viewing window
Dim mTtlFirstScan As Long          'first scan extracted from provided data
Dim mTtlLastScan As Long           'last scan extracted from provided data
Dim mTtlMinMW As Double            'minimum of mass extracted from the provided data
Dim mTtlmaxmw As Double            'maximum of mass extracted from the provided data
Dim mTtlMinAbu As Double           'minimum of abundance extracted from the provided data
Dim mTtlMaxAbu As Double           'maximum of abundance extracted from the provided data

Public OwnerInd As Long            'index of owner display

'charge state spots
Dim mSpotCSCnt As Long
Dim mSpotCSInd() As Long           'index in original display
'isotopic spots
Dim mSpotIsoCnt As Long
Dim mSpotIsoInd() As Long          'index in original display
'boxes
Dim mBoxCnt As Long
Dim mBoxInd() As Long              'index in UMC structure of original display

'following variables are actual properties that could be changed through pop-up menu
'commands so there will not be exposed as public properties
Dim mViewCoordinates As Boolean
Dim mViewCSSpots As Boolean
Dim mViewIsoSpots As Boolean
Dim mViewBoxes As Boolean
Dim mViewScanLines As Boolean
                            
Dim mScaleR As Double
Dim mMinR As Long                  'min size of spot in a view (depends on the number of scans)
Dim mMaxR As Long                  'max size of spot in a view (depends on the number of scans)

Dim mCurrX As String               'current coordinates (formated)
Dim mCurrY As String

'coordinate system-viewport coordinates
Dim VPX0 As Long
Dim VPY0 As Long
Dim VPXE As Long
Dim VPYE As Long

'actual drawing spots and boxes
Dim mCSX() As Long
Dim mCSY() As Long
Dim mCSRX() As Long
Dim mCSRY() As Long

Dim mIsoX() As Long
Dim mIsoY() As Long
Dim mIsoRX() As Long
Dim mIsoRY() As Long

Dim mBoxX1() As Long
Dim mBoxY1() As Long
Dim mBoxX2() As Long
Dim mBoxY2() As Long

Dim hBrushCS As Long             'foreground color brush for Charge State data
Dim hBrushIso As Long            'foreground color brush for Isotopic data
Dim hBackBrush As Long           'background color brush
Dim hBackPen As Long             'background color pen
Dim hForePen As Long             'foreground color pen
Dim hForeDotPen As Long          'foreground color dotted pen

Dim paPoints() As POINTAPI       'used to calculate dev to log coordinate conversion

Public Event MenuClosed()        'indication something happened

Dim mAction As Long              'action that is currently in progress

Dim mZoomX1 As Double      'private variables to help with zooming,
Dim mZoomY1 As Double      'clipping, and mouse tracking features
Dim mZoomX2 As Double
Dim mZoomY2 As Double


Public Sub DrawPic()
'---------------------------------------------------
'draws spots and boxes on picture device context
'---------------------------------------------------
Dim OldDC As Long
Dim OldForePen As Long
Dim Res As Long
On Error Resume Next
OldDC = SaveDC(picD.hDC)
OldForePen = SelectObject(picD.hDC, hForePen)
'Paint_Background picD.hDC, picD.ScaleWidth, picD.ScaleHeight
DGCooSys picD.hDC, picD.ScaleWidth, picD.ScaleHeight
'Res = SetROP2(picD.hdc, R2_MASKPEN)
Res = SetROP2(picD.hDC, R2_COPYPEN)
Res = SetBkMode(picD.hDC, TRANSPARENT)
If mViewCSSpots Then DrawCSSpotsEx picD.hDC
If mViewIsoSpots Then DrawIsoSpotsEx picD.hDC
If mViewBoxes Then DrawBoxesEx picD.hDC
If mViewScanLines Then DrawScanLines picD.hDC
Res = SelectObject(picD.hDC, OldForePen)
Res = RestoreDC(picD.hDC, OldDC)
End Sub

Public Sub DrawPrt()
'---------------------------------------------------
'draws spots and boxes on printer device context
'---------------------------------------------------
End Sub

Public Function DrawEMF() As Long
'----------------------------------------------------------
'draws spots and boxes on enhanced metafile; returns handle
'to enhanced metafile if successful; 0 on any error
'----------------------------------------------------------
Dim Res As Long

Dim OldForePen As Long
Dim OldDC As Long
Dim hRefDC As Long
Dim emfDC As Long                   'metafile device context
Dim emfHandle As Long               'metafile handle

Dim iWidthMM As Long, iHeightMM As Long
Dim iWidthPels As Long, iHeightPels As Long
Dim iMMPerPelX As Double, iMMPerPelY As Double
Dim rcRef As Rect           'reference rectangle
On Error GoTo exit_DrawEMF

hRefDC = picD.hDC
iWidthMM = GetDeviceCaps(hRefDC, HORZSIZE):         iHeightMM = GetDeviceCaps(hRefDC, VERTSIZE)
iWidthPels = GetDeviceCaps(hRefDC, HORZRES):        iHeightPels = GetDeviceCaps(hRefDC, VERTRES)

iMMPerPelX = (iWidthMM * 100) / iWidthPels:         iMMPerPelY = (iHeightMM * 100) / iHeightPels

rcRef.Top = 0:                                      rcRef.Left = 0
rcRef.Bottom = picD.ScaleHeight:                    rcRef.Right = picD.ScaleWidth
'convert to himetric units
rcRef.Left = rcRef.Left * iMMPerPelX:               rcRef.Top = rcRef.Top * iMMPerPelY
rcRef.Right = rcRef.Right * iMMPerPelX:             rcRef.Bottom = rcRef.Bottom * iMMPerPelY

emfDC = CreateEnhMetaFile(hRefDC, vbNullString, rcRef, vbNullString)
OldDC = SaveDC(emfDC)
OldForePen = SelectObject(emfDC, hForePen)
Paint_Background emfDC, picD.ScaleWidth, picD.ScaleHeight
DGCooSys emfDC, picD.ScaleWidth, picD.ScaleHeight
'Res = SetROP2(emfDC, R2_MASKPEN)
Res = SetROP2(emfDC, R2_COPYPEN)
Res = SetBkMode(emfDC, TRANSPARENT)
If mViewCSSpots Then DrawCSSpotsEx emfDC
If mViewIsoSpots Then DrawIsoSpotsEx emfDC
If mViewBoxes Then DrawBoxesEx emfDC
If mViewScanLines Then DrawScanLines emfDC
Res = SelectObject(emfDC, OldForePen)
Res = RestoreDC(emfDC, OldDC)
emfHandle = CloseEnhMetaFile(emfDC)

exit_DrawEMF:
DrawEMF = emfHandle
End Function

Private Sub DGCooSys(ByVal hDC As Long, ByVal dcw As Long, ByVal dch As Long)
'----------------------------------------------------------------------------
'establishes coordinate system on device context
'----------------------------------------------------------------------------
Dim Res As Long
Dim ptPoint As POINTAPI
Dim szSize As Size
On Error Resume Next
VPX0 = 5                'make small offset from the edges
VPY0 = dch - 5
VPXE = dcw - 5
VPYE = 5 - dch
Res = SetMapMode(hDC, MM_ANISOTROPIC)
'logical window
Res = SetWindowOrgEx(hDC, L_X0, L_Y0, ptPoint)
Res = SetWindowExtEx(hDC, L_XE, L_YE, szSize)
'viewport
Res = SetViewportOrgEx(hDC, VPX0, VPY0, ptPoint)
Res = SetViewportExtEx(hDC, VPXE, VPYE, szSize)
End Sub

Private Sub mnuDFCopy_Click()
'----------------------------------------------------------------------
'draws on enhanced metafile device context and copy it to the clipboard
'----------------------------------------------------------------------
Dim hEmf As Long
Dim Res As Long
On Error Resume Next
hEmf = DrawEMF()
If hEmf <> 0 Then
   Res = OpenClipboard(picD.hwnd)
   Res = EmptyClipboard()
   Res = SetClipboardData(CF_ENHMETAFILE, hEmf)
   Res = CloseClipboard
End If
End Sub

Private Sub mnuFOBackColor_Click()
Dim TmpClr As Long
On Error GoTo exit_FOBackColor
TmpClr = mBackClr
Call GetColorAPIDlg(picD.hwnd, TmpClr)
BackColor = TmpClr                        'this will trigger Property Let
exit_FOBackColor:
RaiseEvent MenuClosed
End Sub

Private Sub mnuFOCSColor_Click()
Dim TmpClr As Long
On Error GoTo exit_FOCSColor
TmpClr = mCSClr
Call GetColorAPIDlg(picD.hwnd, TmpClr)
CSColor = TmpClr                          'this will trigger Property Let
exit_FOCSColor:
RaiseEvent MenuClosed
End Sub

Private Sub mnuFOCSShape_Click(Index As Integer)
CSShape = Index         'this will trigger property Let
End Sub

Private Sub mnuFOForeColor_Click()
Dim TmpClr As Long
On Error GoTo exit_FOForeColor
TmpClr = mForeClr
Call GetColorAPIDlg(picD.hwnd, TmpClr)
ForeColor = TmpClr                         'this will trigger Property Let
exit_FOForeColor:
RaiseEvent MenuClosed
End Sub

Private Sub mnuFOIsoColor_Click()
Dim TmpClr As Long
On Error GoTo exit_FOIsoColor
TmpClr = mIsoClr
Call GetColorAPIDlg(picD.hwnd, TmpClr)
IsoColor = TmpClr                          'this will trigger Property Let
exit_FOIsoColor:
RaiseEvent MenuClosed
End Sub

Private Sub mnuFOIsoShape_Click(Index As Integer)
IsoShape = Index                           'this will trigger property Let
End Sub

Private Sub mnuFViewBoxes_Click()
On Error Resume Next
mViewBoxes = Not mViewBoxes
mnuFViewBoxes.Checked = mViewBoxes
RaiseEvent MenuClosed
End Sub

Private Sub mnuFViewCoo_Click()
'----------------------------------------------------------------------------------------
'toggle visibility of coordinates
'----------------------------------------------------------------------------------------
On Error Resume Next
mViewCoordinates = Not mViewCoordinates
mnuFViewCoo.Checked = mViewCoordinates
If Not mViewCoordinates Then WriteCoordinates picD.hDC, picD.ScaleWidth, picD.ScaleHeight  'this will clear coordinates
RaiseEvent MenuClosed
End Sub

Private Sub mnuFViewCSSpots_Click()
On Error Resume Next
mViewCSSpots = Not mViewCSSpots
mnuFViewCSSpots.Checked = mViewCSSpots
RaiseEvent MenuClosed
End Sub

Private Sub mnuFViewIsoSpots_Click()
On Error Resume Next
mViewIsoSpots = Not mViewIsoSpots
mnuFViewIsoSpots.Checked = mViewIsoSpots
RaiseEvent MenuClosed
End Sub

Private Sub mnuFViewScanLines_Click()
On Error Resume Next
mViewScanLines = Not mViewScanLines
mnuFViewScanLines.Checked = mViewScanLines
RaiseEvent MenuClosed
End Sub

Private Sub mnuFZoomOut_Click()
MyCooSys.ZoomOut
End Sub

Private Sub mnuFZoomOut1_Click()
MyCooSys.ZoomPrevious
End Sub

Private Sub MyCooSys_CooSysChanged()
If Not MyCooSys.BuildingCS Then DrawRefresh
End Sub

Private Sub picD_KeyDown(KeyCode As Integer, Shift As Integer)
'-------------------------------------------------------------
'copy as metafile on system clipboard on Ctrl+C combination
'-------------------------------------------------------------
Dim CtrlDown As Boolean
CtrlDown = (Shift And vbCtrlMask) > 0
If CtrlDown Then
   Select Case KeyCode
   Case vbKeyA                  'zoom-out one level
        MyCooSys.ZoomPrevious
   Case vbKeyC
        Call mnuDFCopy_Click
   Case vbKeyZ                  'zoom out
        MyCooSys.ZoomOut
   End Select
End If
End Sub

Private Sub picD_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Res As Long
Dim paClip As POINTAPI
Dim rcClip As Rect
On Error Resume Next
ReDim paPoints(0)
Select Case Button
Case vbLeftButton
     mAction = glActionZoom
Case vbRightButton
     mAction = glNoAction
     UserControl.PopupMenu mnuF
End Select
If mAction = glActionZoom Then
   mZoomX1 = x:                 mZoomY1 = y
   mZoomX2 = mZoomX1:           mZoomY2 = mZoomY1
   'clip cursor on viewport; lengthy but straightforward sequence
   paPoints(0).x = mZoomX1:     paPoints(0).y = mZoomY1
   DevLogConversion ConvDPLP, 1
   If (paPoints(0).x < L_X0) Or (paPoints(0).x > L_XE) Then
      mAction = glNoAction
      Exit Sub
   End If
   If (paPoints(0).y < L_Y0) Or (paPoints(0).y > L_YE) Then
      mAction = glNoAction
      Exit Sub
   End If
   paClip.x = 0:     paClip.y = 0
   Res = ClientToScreen(picD.hwnd, paClip)
   If VPXE < 0 Then
      rcClip.Left = paClip.x + VPX0 + VPXE:      rcClip.Right = paClip.x + VPX0
   Else
      rcClip.Left = paClip.x + VPX0:             rcClip.Right = paClip.x + VPX0 + VPXE
   End If
   If VPYE < 0 Then
      rcClip.Top = paClip.y + VPY0 + VPYE:       rcClip.Bottom = paClip.y + VPY0
   Else
      rcClip.Top = paClip.y + VPY0:              rcClip.Bottom = paClip.y + VPY0 + VPYE
   End If
   Res = ClipCursor(rcClip)
   picD.DrawStyle = vbDot
End If
End Sub

Private Sub picD_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
ReDim paPoints(0)
paPoints(0).x = x
paPoints(0).y = y
DevLogConversion ConvLPDP, 1
TrackHotSpot paPoints(0).x, paPoints(0).y
If mAction = glActionZoom Then
   picD.Line (mZoomX1, mZoomY1)-(mZoomX2, mZoomY2), , B
   mZoomX2 = x
   mZoomY2 = y
   picD.Line (mZoomX1, mZoomY1)-(mZoomX2, mZoomY2), , B
End If
If mViewCoordinates Then WriteCoordinates picD.hDC, picD.ScaleWidth, picD.ScaleHeight
End Sub

Private Sub picD_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Res As Long
ReDim paPoints(1)
Select Case mAction
Case glNoAction                         'remove restrictions on cursor
  If Button = vbLeftButton Then Res = ClipCursorByNum(&O0)
Case glActionZoom
  If Button = vbLeftButton Then
     Res = ClipCursorByNum(&O0)
     picD.Line (mZoomX1, mZoomY1)-(mZoomX2, mZoomY2), , B
     picD.DrawStyle = vbSolid
     If (Abs(mZoomX1 - mZoomX2) > 10) And (Abs(mZoomY1 - mZoomY2) > 10) Then
        paPoints(0).x = mZoomX1:        paPoints(0).y = mZoomY1
        paPoints(1).x = mZoomX2:        paPoints(1).y = mZoomY2
        DevLogConversion ConvLPDP, 2
        MyCooSys.ZoomInL paPoints(0).x, paPoints(0).y, paPoints(1).x, paPoints(1).y
     End If
  End If
End Select
mAction = glNoAction
End Sub

Private Sub picD_Paint()
DrawRefresh
End Sub

Private Sub UserControl_Initialize()
'following properties can be overwritten by client; these are like a defaults
Set MyCooSys = New CooSysS
MyCooSys.LogX0 = L_X0:      MyCooSys.LogXE = L_XE
MyCooSys.LogY0 = L_Y0:      MyCooSys.LogYE = L_YE
mViewCoordinates = mnuFViewCoo.Checked
mViewCSSpots = mnuFViewCSSpots.Checked
mViewIsoSpots = mnuFViewIsoSpots.Checked
mViewBoxes = mnuFViewBoxes.Checked
mViewScanLines = mnuFViewScanLines.Checked
'SetDrawingObjects
CSShape = glShapeEli
IsoShape = glShapeGradRec
End Sub

Private Sub UserControl_Resize()
picD.width = UserControl.ScaleWidth
picD.Height = UserControl.ScaleHeight
End Sub


Private Sub CalculatePrepareScaling()
On Error Resume Next
mMaxR = CLng((L_XE - L_Y0) / (2 * (MyCooSys.CurrRXMax - MyCooSys.CurrRXMin + 1 + 1)))
mMinR = CLng(mMaxR / 10)
mScaleR = (mMaxR - mMinR) / (mTtlMaxAbu - mTtlMinAbu)
End Sub


Private Sub CalculateSpots()
'-----------------------------------------------------------------------
'calculate coordinates of all points that should be displayed
'NOTE: this procedure binds this user control to this application
'-----------------------------------------------------------------------
Dim i As Long
Dim CurrScan As Long
Dim CurrMW As Double
Dim CurrAbu As Double
Dim CurrOInd As Long                'original index in CS/Iso Num arrays
On Error Resume Next

With GelData(OwnerInd)
   If mSpotCSCnt > 0 Then
      ReDim mCSX(mSpotCSCnt - 1):             ReDim mCSY(mSpotCSCnt - 1)
      ReDim mCSRX(mSpotCSCnt - 1):            ReDim mCSRY(mSpotCSCnt - 1)
      For i = 0 To mSpotCSCnt - 1
          CurrOInd = Abs(mSpotCSInd(i))
          CurrScan = .CSData(CurrOInd).ScanNumber
          CurrMW = .CSData(CurrOInd).AverageMW
          CurrAbu = .CSData(CurrOInd).Abundance
          If MyCooSys.IsInScope(CurrScan, CurrMW) Then
             mCSX(i) = CLng(MyCooSys.ScaleX_RToL * (CurrScan - MyCooSys.CurrRXMin))
             mCSY(i) = CLng(MyCooSys.ScaleY_RToL * (CurrMW - MyCooSys.CurrRYMin))
             mCSRX(i) = CLng((CurrAbu - mTtlMinAbu) * mScaleR + mMinR)
             mCSRY(i) = CLng(mCSRX(i) * MyCooSys.csXYAspRat + mMinR)
             mSpotCSInd(i) = CurrOInd
          Else
             mSpotCSInd(i) = -CurrOInd         'negative index as indication not to draw
          End If
      Next i
   Else
      Erase mCSX:   Erase mCSY:   Erase mCSRX:   Erase mCSRY
   End If

   If mSpotIsoCnt > 0 Then
      ReDim mIsoX(mSpotIsoCnt - 1):           ReDim mIsoY(mSpotIsoCnt - 1)
      ReDim mIsoRX(mSpotIsoCnt - 1):          ReDim mIsoRY(mSpotIsoCnt - 1)
      For i = 0 To mSpotIsoCnt - 1
          CurrOInd = Abs(mSpotIsoInd(i))
          CurrScan = .IsoData(CurrOInd).ScanNumber
          CurrMW = GetIsoMass(.IsoData(CurrOInd), .Preferences.IsoDataField)
          CurrAbu = .IsoData(CurrOInd).Abundance
          If MyCooSys.IsInScope(CurrScan, CurrMW) Then
             mIsoX(i) = CLng(MyCooSys.ScaleX_RToL * (CurrScan - MyCooSys.CurrRXMin))
             mIsoY(i) = CLng(MyCooSys.ScaleY_RToL * (CurrMW - MyCooSys.CurrRYMin))
             mIsoRX(i) = CLng((CurrAbu - mTtlMinAbu) * mScaleR + mMinR)
             mIsoRY(i) = CLng(mIsoRX(i) * MyCooSys.csXYAspRat + mMinR)
             mSpotIsoInd(i) = CurrOInd
          Else
             mSpotIsoInd(i) = -CurrOInd
          End If
      Next i
   Else
      Erase mIsoX:   Erase mIsoY:   Erase mIsoRX:   Erase mIsoRY
   End If
End With
End Sub


Private Sub CalculateBoxes()
Dim i As Long
On Error Resume Next
If mBoxCnt > 0 Then
   ReDim mBoxX1(mBoxCnt - 1):      ReDim mBoxY1(mBoxCnt - 1)
   ReDim mBoxX2(mBoxCnt - 1):      ReDim mBoxY2(mBoxCnt - 1)
   With GelUMC(OwnerInd)
        For i = 0 To mBoxCnt - 1
            mBoxX1(i) = CLng(MyCooSys.ScaleX_RToL * (.UMCs(mBoxInd(i)).MinScan - MyCooSys.CurrRXMin))
            mBoxY1(i) = CLng(MyCooSys.ScaleY_RToL * (.UMCs(mBoxInd(i)).MinMW - MyCooSys.CurrRYMin))
            mBoxX2(i) = CLng(MyCooSys.ScaleX_RToL * (.UMCs(mBoxInd(i)).MaxScan - MyCooSys.CurrRXMin))
            mBoxY2(i) = CLng(MyCooSys.ScaleY_RToL * (.UMCs(mBoxInd(i)).MaxMW - MyCooSys.CurrRYMin))
        Next i
   End With
Else
   Erase mBoxX1:    Erase mBoxY1:   Erase mBoxX2:   Erase mBoxY2
End If
End Sub


Private Sub WriteCoordinates(hDC As Long, w As Long, h As Long)
Dim OldDC As Long
Dim lfLogFont As LOGFONT
Dim lOldFont As Long
Dim lNewFont As Long
Dim Lbl As String
Dim lFont As Long
Dim szLbl As Size
Dim Res As Long
Dim OldBkMode As Long
On Error Resume Next
OldDC = SaveDC(hDC)
DGCooSys hDC, w, h
lFont = SelectObject(hDC, GetStockObject(SYSTEM_FONT))
Res = GetObjectAPI(lFont, Len(lfLogFont), lfLogFont)
Res = SelectObject(hDC, lFont)
'create new logical font
lfLogFont.lfWidth = CLng(L_XE / 80)
lfLogFont.lfHeight = CLng(lfLogFont.lfWidth * 8)
lNewFont = CreateFontIndirect(lfLogFont)
'select newly created logical font to DC
lOldFont = SelectObject(hDC, lNewFont)
'write new coordinates
Lbl = "Coordinates: " & mCurrX & ",  " & mCurrY
Res = SetTextColor(hDC, mForeClr)
Res = SetBkColor(hDC, mBackClr)
OldBkMode = SetBkMode(hDC, OPAQUE)
Res = GetTextExtentPoint32(hDC, Lbl, Len(Lbl), szLbl)
Res = TextOut(hDC, CLng(L_XE / 32), L_YE, Lbl, Len(Lbl))
Res = SetBkMode(hDC, OldBkMode)
Res = SelectObject(hDC, lOldFont)
DeleteObject (lNewFont)
Res = RestoreDC(hDC, OldDC)
End Sub


Public Sub DrawRefresh()
'-----------------------------------------------------
'rebuild coordinate system and redraws spots and boxes
'-----------------------------------------------------
picD.Cls
Call CalculatePrepareScaling
Call CalculateSpots
Call CalculateBoxes
Call DrawPic
End Sub


Public Function WriteSToFile(ByVal Filename As String) As Boolean
'---------------------------------------------------------------
'write bins results to text semicolon delimited file
'returns True if successful, False if not or any error occurs
'NOTE: bins are appended to a file; create file if not found
'---------------------------------------------------------------
Dim i As Long
Dim hfile As Integer
On Error GoTo exit_WriteSToFile

'hfile = FreeFile
'Open FileName For Append As hfile
'Print #hfile, "Presenting - " & HLabel & ", " & VLabel
'Print #hfile, "Range - [" & mMinX & ", " & mMaxX & "]X[" & mMinY & ", " & mMaxY & "]"
'Print #hfile, vbCrLf
'Print #hfile, "ID;X;Y"
'For i = 0 To mDataCnt - 1
'    Print #hfile, mDataID(i) & ";" & mDataX(i) & ";" & mDataY(i)
'Next i
'Close hfile
'WriteSToFile = True
exit_WriteSToFile:
End Function

Public Sub DestroyDrawingObjects()
On Error Resume Next
If hBrushCS <> 0 Then DeleteObject (hBrushCS)
If hBrushIso <> 0 Then DeleteObject (hBrushIso)
If hBackBrush <> 0 Then DeleteObject (hBackBrush)
If hBackPen <> 0 Then DeleteObject (hBackPen)
If hForePen <> 0 Then DeleteObject (hForePen)
If hForeDotPen <> 0 Then DeleteObject (hForeDotPen)
End Sub

Public Sub SetDrawingObjects()
On Error Resume Next
If hBrushCS <> 0 Then DeleteObject (hBrushCS)
hBrushCS = CreateSolidBrush(mCSClr)
If hBrushIso <> 0 Then DeleteObject (hBrushIso)
hBrushIso = CreateSolidBrush(mIsoClr)
If hBackBrush <> 0 Then DeleteObject (hBackBrush)
hBackBrush = CreateSolidBrush(mBackClr)
If hBackPen <> 0 Then DeleteObject (hBackPen)
hBackPen = CreatePen(PS_SOLID, 1, mBackClr)
If hForePen <> 0 Then DeleteObject (hForePen)
hForePen = CreatePen(PS_SOLID, 1, mForeClr)
If hForeDotPen <> 0 Then DeleteObject (hForeDotPen)
hForeDotPen = CreatePen(PS_DOT, 1, mForeClr)
End Sub

Private Sub UserControl_Terminate()
DestroyDrawingObjects
End Sub

Private Sub DevLogConversion(ByVal ConversionType As Integer, ByVal NumOfPoints As Integer)
Dim OldDC As Long, Res As Long
On Error Resume Next
OldDC = SaveDC(picD.hDC)
DGCooSys picD.hDC, picD.ScaleWidth, picD.ScaleHeight
Select Case ConversionType
Case ConvLPDP               'logical to device coordinates
    Res = DPtoLP(picD.hDC, paPoints(0), NumOfPoints)
Case ConvDPLP               'device to logical coordinates
    Res = LPtoDP(picD.hDC, paPoints(0), NumOfPoints)
End Select
Res = RestoreDC(picD.hDC, OldDC)
End Sub


Private Sub TrackHotSpot(lx As Long, ly As Long)
On Error Resume Next
With MyCooSys
     mCurrX = Format$(.CurrRXMin + lx / .ScaleX_RToL, "0.0000")  'current coordinates
     mCurrY = Format$(.CurrRYMin + ly / .ScaleY_RToL, "0.0000")
End With
End Sub


Public Function AddBoxes(UMCInd() As Long) As Boolean
'-------------------------------------------------------------
'adds new boxes to array of boxes; returns True if successful
'-------------------------------------------------------------
On Error GoTo err_AddBoxes
mBoxCnt = UBound(UMCInd) + 1
mBoxInd = UMCInd                'array assignment
AddBoxes = True
exit_AddBoxes:
Exit Function

err_AddBoxes:
mBoxCnt = 0
Erase mBoxInd
Resume exit_AddBoxes
End Function


Public Function AddSpotsCS(CSInd() As Long) As Boolean
'----------------------------------------------------------------------
'adds charge state spots to array of spots ; returns True if successful
'----------------------------------------------------------------------
On Error GoTo err_AddSpotsCS
mSpotCSCnt = UBound(CSInd) + 1
mSpotCSInd = CSInd                'array assignment
AddSpotsCS = True
exit_AddSpotsCS:
Exit Function

err_AddSpotsCS:
mSpotCSCnt = 0
Erase mSpotCSInd
Resume exit_AddSpotsCS
End Function


Public Function AddSpotsIso(IsoInd() As Long) As Boolean
'----------------------------------------------------------------------
'adds charge state spots to array of spots ; returns True if successful
'----------------------------------------------------------------------
On Error GoTo err_AddSpotsIso
mSpotIsoCnt = UBound(IsoInd) + 1
mSpotIsoInd = IsoInd                'array assignment
AddSpotsIso = True
exit_AddSpotsIso:
Exit Function

err_AddSpotsIso:
mSpotIsoCnt = 0
Erase mSpotIsoInd
Resume exit_AddSpotsIso
End Function


Public Sub DrawCSSpotsEx(ByVal hDC As Long)
Dim i As Long
Dim OldBrush As Long
Dim ptAPIs() As POINTAPI
Dim OffsetX As Long, OffsetY As Long
On Error Resume Next
OldBrush = SelectObject(hDC, hBrushCS)
Select Case mCSShapeType
Case glShapeEli
  For i = 0 To mSpotCSCnt - 1
      If mSpotCSInd(i) > 0 Then
         Ellipse hDC, mCSX(i) - mCSRX(i), mCSY(i) - mCSRY(i), mCSX(i) + mCSRX(i), mCSY(i) + mCSRY(i)
      End If
  Next i
Case glShapeRec
  For i = 0 To mSpotCSCnt - 1
      If mSpotCSInd(i) > 0 Then
         Rectangle hDC, mCSX(i) - mCSRX(i), mCSY(i) - mCSRY(i), mCSX(i) + mCSRX(i), mCSY(i) + mCSRY(i)
      End If
  Next i
Case glShapeRRe
  For i = 0 To mSpotCSCnt - 1
      If mSpotCSInd(i) > 0 Then
         RoundRect hDC, mCSX(i) - mCSRX(i), mCSY(i) - mCSRY(i), mCSX(i) + mCSRX(i), mCSY(i) + mCSRY(i), mCSRX(i), mCSRY(i)
      End If
  Next i
Case glShapeTri
  ReDim ptAPIs(2) As POINTAPI
  For i = 0 To mSpotCSCnt - 1
      If mSpotCSInd(i) > 0 Then
         OffsetX = CLng(mCSRX(i) * 0.87):          OffsetY = CLng(mCSRY(i) * 0.87)
         ptAPIs(0).x = mCSX(i) - OffsetX:          ptAPIs(0).y = mCSY(i) - OffsetY
         ptAPIs(1).x = mCSX(i):                    ptAPIs(1).y = mCSY(i) + OffsetY
         ptAPIs(2).x = mCSX(i) + OffsetX:          ptAPIs(2).y = mCSY(i) - OffsetY
         Polygon hDC, ptAPIs(0), 3
      End If
  Next i
Case glShapeSta
  ReDim ptAPIs(7) As POINTAPI
  For i = 0 To mSpotCSCnt - 1
      If mSpotCSInd(i) > 0 Then
         OffsetX = CLng(mCSRX(i) / 4):             OffsetY = CLng(mCSRY(i) / 4)
         ptAPIs(0).x = mCSX(i) - mCSRX(i):         ptAPIs(0).y = mCSY(i)
         ptAPIs(1).x = mCSX(i) - OffsetX:          ptAPIs(1).y = mCSY(i) - OffsetY
         ptAPIs(2).x = mCSX(i):                    ptAPIs(2).y = mCSY(i) - mCSRY(i)
         ptAPIs(3).x = mCSX(i) + OffsetX:          ptAPIs(3).y = mCSY(i) - OffsetY
         ptAPIs(4).x = mCSX(i) + mCSRX(i):         ptAPIs(4).y = mCSY(i)
         ptAPIs(5).x = mCSX(i) + OffsetX:          ptAPIs(5).y = mCSY(i) + OffsetY
         ptAPIs(6).x = mCSX(i):                    ptAPIs(6).y = mCSY(i) + mCSRY(i)
         ptAPIs(7).x = mCSX(i) - OffsetX:          ptAPIs(7).y = mCSY(i) + OffsetY
         Polygon hDC, ptAPIs(0), 8
      End If
  Next i
Case glShapeHex
  ReDim ptAPIs(5) As POINTAPI
  For i = 0 To mSpotCSCnt - 1
      If mSpotCSInd(i) > 0 Then
         OffsetX = CLng(mCSRX(i) * 0.87):          OffsetY = CLng(mCSRY(i) * 0.87)
         ptAPIs(0).x = mCSX(i) - mCSRX(i):         ptAPIs(0).y = mCSY(i)
         ptAPIs(1).x = mCSX(i) - OffsetX:          ptAPIs(1).y = mCSY(i) - OffsetY
         ptAPIs(2).x = mCSX(i) + OffsetX:          ptAPIs(2).y = mCSY(i) - OffsetY
         ptAPIs(3).x = mCSX(i) + mCSRX(i):         ptAPIs(3).y = mCSY(i)
         ptAPIs(4).x = mCSX(i) + OffsetX:          ptAPIs(4).y = mCSY(i) + OffsetY
         ptAPIs(5).x = mCSX(i) - OffsetX:          ptAPIs(5).y = mCSY(i) + OffsetY
         Polygon hDC, ptAPIs(0), 6
      End If
  Next i
Case glShapeGradRec
  Dim Vert(1) As TRIVERTEX
  Dim gRec As GRADIENT_RECT
  For i = 0 To mSpotCSCnt - 1
      If mSpotCSInd(i) > 0 Then
         With Vert(0)                      'set color at upper left corner
              .x = mCSX(i) - mCSRX(i):              .y = mCSY(i) - mCSRY(i)
              .Alpha = 0
              .Red = LongToSignedShort((mCSClr And &HFF) * 256)
              .Green = LongToSignedShort(((mCSClr And &HFF00) \ &H100) * 256)
              .Blue = LongToSignedShort(((mCSClr And &HFF0000) \ &H10000) * 256)
         End With
         With Vert(1)                      'set color at bottom right corner
              .x = mCSX(i) + mCSRX(i):              .y = mCSY(i) + mCSRY(i)
              .Alpha = 0
              .Red = 0:              .Green = 0:              .Blue = 0
         End With
         With gRec
              .LowerRight = 0
              .UpperLeft = 1
         End With
         Call GradientFill(hDC, Vert(0), 2, gRec, 1, GRADIENT_FILL_RECT_H)
      End If
  Next i
End Select
i = SelectObject(hDC, OldBrush)
End Sub


Public Sub DrawIsoSpotsEx(ByVal hDC As Long)
Dim i As Long
Dim OldBrush As Long
Dim ptAPIs() As POINTAPI
Dim OffsetX As Long, OffsetY As Long
On Error Resume Next
OldBrush = SelectObject(hDC, hBrushIso)
Select Case mIsoShapeType
Case glShapeEli
  For i = 0 To mSpotIsoCnt - 1
      If mSpotIsoInd(i) > 0 Then
         Ellipse hDC, mIsoX(i) - mIsoRX(i), mIsoY(i) - mIsoRY(i), mIsoX(i) + mIsoRX(i), mIsoY(i) + mIsoRY(i)
      End If
  Next i
Case glShapeRec
  For i = 0 To mSpotIsoCnt - 1
      If mSpotIsoInd(i) > 0 Then
         Rectangle hDC, mIsoX(i) - mIsoRX(i), mIsoY(i) - mIsoRY(i), mIsoX(i) + mIsoRX(i), mIsoY(i) + mIsoRY(i)
      End If
  Next i
Case glShapeRRe
  For i = 0 To mSpotIsoCnt - 1
      If mSpotIsoInd(i) > 0 Then
         RoundRect hDC, mIsoX(i) - mIsoRX(i), mIsoY(i) - mIsoRY(i), mIsoX(i) + mIsoRX(i), mIsoY(i) + mIsoRY(i), mIsoRX(i), mIsoRY(i)
      End If
  Next i
Case glShapeTri
  ReDim ptAPIs(2) As POINTAPI
  For i = 0 To mSpotIsoCnt - 1
      If mSpotIsoInd(i) > 0 Then
         OffsetX = CLng(mIsoRX(i) * 0.87):          OffsetY = CLng(mIsoRY(i) * 0.87)
         ptAPIs(0).x = mIsoX(i) - OffsetX:          ptAPIs(0).y = mIsoY(i) - OffsetY
         ptAPIs(1).x = mIsoX(i):                    ptAPIs(1).y = mIsoY(i) + OffsetY
         ptAPIs(2).x = mIsoX(i) + OffsetX:          ptAPIs(2).y = mIsoY(i) - OffsetY
         Polygon hDC, ptAPIs(0), 3
      End If
  Next i
Case glShapeSta
  ReDim ptAPIs(7) As POINTAPI
  For i = 0 To mSpotIsoCnt - 1
      If mSpotIsoInd(i) > 0 Then
         OffsetX = CLng(mIsoRX(i) / 4):             OffsetY = CLng(mIsoRY(i) / 4)
         ptAPIs(0).x = mIsoX(i) - mIsoRX(i):        ptAPIs(0).y = mIsoY(i)
         ptAPIs(1).x = mIsoX(i) - OffsetX:          ptAPIs(1).y = mIsoY(i) - OffsetY
         ptAPIs(2).x = mIsoX(i):                    ptAPIs(2).y = mIsoY(i) - mIsoRY(i)
         ptAPIs(3).x = mIsoX(i) + OffsetX:          ptAPIs(3).y = mIsoY(i) - OffsetY
         ptAPIs(4).x = mIsoX(i) + mIsoRX(i):        ptAPIs(4).y = mIsoY(i)
         ptAPIs(5).x = mIsoX(i) + OffsetX:          ptAPIs(5).y = mIsoY(i) + OffsetY
         ptAPIs(6).x = mIsoX(i):                    ptAPIs(6).y = mIsoY(i) + mIsoRY(i)
         ptAPIs(7).x = mIsoX(i) - OffsetX:          ptAPIs(7).y = mIsoY(i) + OffsetY
         Polygon hDC, ptAPIs(0), 8
      End If
  Next i
Case glShapeHex
  ReDim ptAPIs(5) As POINTAPI
  For i = 0 To mSpotIsoCnt - 1
      If mSpotIsoInd(i) > 0 Then
         OffsetX = CLng(mIsoRX(i) * 0.87):          OffsetY = CLng(mIsoRY(i) * 0.87)
         ptAPIs(0).x = mIsoX(i) - mIsoRX(i):        ptAPIs(0).y = mIsoY(i)
         ptAPIs(1).x = mIsoX(i) - OffsetX:          ptAPIs(1).y = mIsoY(i) - OffsetY
         ptAPIs(2).x = mIsoX(i) + OffsetX:          ptAPIs(2).y = mIsoY(i) - OffsetY
         ptAPIs(3).x = mIsoX(i) + mIsoRX(i):        ptAPIs(3).y = mIsoY(i)
         ptAPIs(4).x = mIsoX(i) + OffsetX:          ptAPIs(4).y = mIsoY(i) + OffsetY
         ptAPIs(5).x = mIsoX(i) - OffsetX:          ptAPIs(5).y = mIsoY(i) + OffsetY
         Polygon hDC, ptAPIs(0), 6
      End If
  Next i
Case glShapeGradRec
  Dim Vert(1) As TRIVERTEX
  Dim gRec As GRADIENT_RECT
  For i = 0 To mSpotIsoCnt - 1
      If mSpotIsoInd(i) > 0 Then
         With Vert(0)                      'set color at upper left corner
              .x = mIsoX(i) - mIsoRX(i):              .y = mIsoY(i) - mIsoRY(i)
              .Alpha = 0
              .Red = LongToSignedShort((mIsoClr And &HFF) * 256)
              .Green = LongToSignedShort(((mIsoClr And &HFF00) \ &H100) * 256)
              .Blue = LongToSignedShort(((mIsoClr And &HFF0000) \ &H10000) * 256)
         End With
         With Vert(1)                      'set color at bottom right corner
              .x = mIsoX(i) + mIsoRX(i):              .y = mIsoY(i) + mIsoRY(i)
              .Alpha = 0
              .Red = 0:              .Green = 0:              .Blue = 0
         End With
         With gRec
              .LowerRight = 0
              .UpperLeft = 1
         End With
         Call GradientFill(hDC, Vert(0), 2, gRec, 1, GRADIENT_FILL_RECT_H)
      End If
  Next i
End Select
i = SelectObject(hDC, OldBrush)
End Sub


Public Sub DrawBoxesEx(ByVal hDC As Long)
Dim OldBrush As Long
Dim i As Long
Dim Res As Long
Dim ptDummy As POINTAPI
On Error Resume Next
OldBrush = SelectObject(hDC, GetStockObject(NULL_BRUSH))
For i = 0 To mBoxCnt - 1
    Rectangle hDC, mBoxX1(i), mBoxY1(i), mBoxX2(i), mBoxY2(i)
    Res = MoveToEx(hDC, mBoxX1(i), mBoxY1(i), ptDummy)          'draw line since box
    Res = LineTo(hDC, mBoxX2(i), mBoxY2(i))                     'is not always visible
Next i
i = SelectObject(hDC, OldBrush)
End Sub


Private Sub DrawScanLines(ByVal hDC As Long)
Dim OldForePen As Long
Dim Res As Long, i As Long
Dim FirstScan As Long, LastScan As Long
Dim CurrScanLX As Long
Dim ptDummy As POINTAPI
On Error Resume Next
OldForePen = SelectObject(hDC, hForeDotPen)
FirstScan = CLng(MyCooSys.CurrRXMin):       LastScan = CLng(MyCooSys.CurrRXMax)
For i = FirstScan To LastScan
    CurrScanLX = CLng(MyCooSys.ScaleX_RToL * (i - MyCooSys.CurrRXMin))
    Res = MoveToEx(hDC, CurrScanLX, L_Y0, ptDummy)
    Res = LineTo(hDC, CurrScanLX, L_YE)
Next i
Res = SelectObject(hDC, OldForePen)
End Sub


Public Function InitCoordinateSystem() As Boolean
'-----------------------------------------------------------------------
'this procedure has to be called after spots are loaded to initiate 1st
'zoom-level for coordinate system; returns True if successful
'-----------------------------------------------------------------------
On Error Resume Next
Call FindRangeLimits
If mTtlFirstScan <= mTtlLastScan Then
   If mTtlMinMW <= mTtlmaxmw Then
      MyCooSys.ZoomOut
      InitCoordinateSystem = True
   End If
End If
End Function


Private Sub FindRangeLimits()
'----------------------------------------------------------------------
'establishes range limits for loaded data
'----------------------------------------------------------------------
Dim Abu As Double
Dim i As Long
On Error Resume Next
mTtlMinAbu = glHugeDouble:              mTtlMaxAbu = -glHugeDouble

With GelData(OwnerInd)
   For i = 0 To mSpotCSCnt - 1
       Abu = .CSData(mSpotCSInd(i)).Abundance
       If Abu < mTtlMinAbu Then mTtlMinAbu = Abu
       If Abu > mTtlMaxAbu Then mTtlMaxAbu = Abu
   Next i
   For i = 0 To mSpotIsoCnt - 1
       Abu = .IsoData(mSpotIsoInd(i)).Abundance
       If Abu < mTtlMinAbu Then mTtlMinAbu = Abu
       If Abu > mTtlMaxAbu Then mTtlMaxAbu = Abu
   Next i
End With
End Sub


Public Sub SetRangeLimits(Scan1 As Long, Scan2 As Long, MW1 As Double, MW2 As Double)
mTtlFirstScan = Scan1:          mTtlLastScan = Scan2
mTtlMinMW = MW1:                mTtlmaxmw = MW2
MyCooSys.BuildingCS = True
MyCooSys.ZoomInRFirst mTtlFirstScan, mTtlMinMW, mTtlLastScan, mTtlmaxmw
MyCooSys.BuildingCS = False
End Sub


Private Sub Paint_Background(ByVal hDC As Long, w As Long, h As Long)
Dim OldBrush As Long
Dim Res As Long
OldBrush = SelectObject(hDC, hBackBrush)
Res = PatBlt(hDC, 0, 0, w, h, PATCOPY)
Res = SelectObject(hDC, OldBrush)
End Sub

Public Property Get BackColor() As Long
BackColor = picD.BackColor
End Property

Public Property Let BackColor(ByVal NewBackColor As Long)
On Error Resume Next
mBackClr = NewBackColor
picD.BackColor = mBackClr
Call SetDrawingObjects
If Not MyCooSys.BuildingCS Then picD.Refresh
End Property

Public Property Get ForeColor() As Long
ForeColor = mForeClr
End Property

Public Property Let ForeColor(ByVal NewForeColor As Long)
On Error Resume Next
mForeClr = NewForeColor
Call SetDrawingObjects
If Not MyCooSys.BuildingCS Then picD.Refresh
End Property

Public Property Get CSColor() As Long
CSColor = mCSClr
End Property

Public Property Let CSColor(ByVal NewCSColor As Long)
On Error Resume Next
mCSClr = NewCSColor
Call SetDrawingObjects
If Not MyCooSys.BuildingCS Then picD.Refresh
End Property

Public Property Get IsoColor() As Long
IsoColor = mIsoClr
End Property

Public Property Let IsoColor(ByVal NewIsoColor As Long)
On Error Resume Next
mIsoClr = NewIsoColor
Call SetDrawingObjects
If Not MyCooSys.BuildingCS Then picD.Refresh
End Property

Public Property Get CSShape() As Long
CSShape = mCSShapeType
End Property

Public Property Let CSShape(ByVal NewCSShape As Long)
On Error Resume Next
mCSShapeType = NewCSShape
Call SyncCSShapeMenu
If Not MyCooSys.BuildingCS Then         'redraw (no need to recalculate)
   picD.Cls
   Call DrawPic
End If
End Property

Private Sub SyncCSShapeMenu()
Dim i As Long
On Error Resume Next
For i = 0 To mnuFOCSShape.Count - 1
    If i = mCSShapeType Then
       mnuFOCSShape(i).Checked = True
    Else
       mnuFOCSShape(i).Checked = False
    End If
Next i
End Sub


Public Property Get IsoShape() As Long
IsoShape = mIsoShapeType
End Property

Public Property Let IsoShape(ByVal NewIsoShape As Long)
On Error Resume Next
mIsoShapeType = NewIsoShape
Call SyncIsoShapeMenu
If Not MyCooSys.BuildingCS Then         'redraw (no need to recalculate)
   picD.Cls
   Call DrawPic
End If
End Property

Private Sub SyncIsoShapeMenu()
Dim i As Long
On Error Resume Next
For i = 0 To mnuFOIsoShape.Count - 1
    If i = mIsoShapeType Then
       mnuFOIsoShape(i).Checked = True
    Else
       mnuFOIsoShape(i).Checked = False
    End If
Next i
End Sub


