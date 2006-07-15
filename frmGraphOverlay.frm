VERSION 5.00
Begin VB.Form frmGraphOverlay 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Overlay"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8010
   ClipControls    =   0   'False
   DrawStyle       =   2  'Dot
   Icon            =   "frmGraphOverlay.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   352
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   534
   Begin VB.PictureBox picGraph 
      BackColor       =   &H00FFFFFF&
      DrawMode        =   10  'Mask Pen
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5250
      Left            =   0
      ScaleHeight     =   346
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   530
      TabIndex        =   0
      Top             =   0
      Width           =   8010
      Begin VB.PictureBox picCoordinates 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5280
         ScaleHeight     =   300
         ScaleWidth      =   2115
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   2115
         Begin VB.Label lblVCoo 
            Alignment       =   1  'Right Justify
            Caption         =   "0 "
            Height          =   225
            Left            =   1080
            TabIndex        =   3
            Top             =   30
            Width           =   975
         End
         Begin VB.Label lblHCoo 
            Alignment       =   1  'Right Justify
            Caption         =   "0 "
            Height          =   225
            Left            =   60
            TabIndex        =   2
            Top             =   30
            Width           =   975
         End
      End
   End
   Begin VB.Menu mnuF 
      Caption         =   "&File"
      HelpContextID   =   105
      Begin VB.Menu mnuFManager 
         Caption         =   "Overlay Manager"
      End
      Begin VB.Menu mnuFSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFSave 
         Caption         =   "&Save"
         HelpContextID   =   105
      End
      Begin VB.Menu mnuFSaveAs 
         Caption         =   "Save &As"
         HelpContextID   =   105
      End
      Begin VB.Menu mnuFSaveAsDisplay 
         Caption         =   "Save As Display"
      End
      Begin VB.Menu mnuFSavePic 
         Caption         =   "Sa&ve Picture"
      End
      Begin VB.Menu mnuFSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFPrint 
         Caption         =   "&Print"
         HelpContextID   =   105
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFPrintSetup 
         Caption         =   "Print Set&up"
      End
      Begin VB.Menu mnuFPrintMetric 
         Caption         =   "Print &Metric"
      End
      Begin VB.Menu mnuFSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFClose 
         Caption         =   "&Close"
         HelpContextID   =   105
      End
   End
   Begin VB.Menu mnuE 
      Caption         =   "&Edit"
      HelpContextID   =   106
      Begin VB.Menu mnuECopyBMP 
         Caption         =   "Copy As &BMP"
      End
      Begin VB.Menu mnuECopyWMF 
         Caption         =   "Copy As &WMF"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuESep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEFillMassTagNames 
         Caption         =   "Fill MT tag Names"
      End
      Begin VB.Menu mnuESep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEComment 
         Caption         =   "Comm&ent"
      End
   End
   Begin VB.Menu mnuV 
      Caption         =   "&View"
      HelpContextID   =   107
      Begin VB.Menu mnuVChargeState 
         Caption         =   "&Charge State"
         HelpContextID   =   107
      End
      Begin VB.Menu mnuVIsotopic 
         Caption         =   "I&sotopic"
         HelpContextID   =   107
      End
      Begin VB.Menu mnuVSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVOrientation 
         Caption         =   "NET Horizontal"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuVOrientation 
         Caption         =   "NET Vertical"
         Index           =   1
      End
      Begin VB.Menu mnuVSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVCoordinates 
         Caption         =   "Coor&dinates"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuVSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVDeviceCaps 
         Caption         =   "Device Caps"
      End
   End
   Begin VB.Menu mnuT 
      Caption         =   "&Tools"
      HelpContextID   =   108
      Begin VB.Menu mnuTGetJiggy 
         Caption         =   "Get Jiggy"
      End
      Begin VB.Menu mnuTSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTZoomIn 
         Caption         =   "Zoom &In"
      End
      Begin VB.Menu mnuTZoomOut 
         Caption         =   "Zoom &Out"
         HelpContextID   =   108
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuTZoomOutOneLevel 
         Caption         =   "Zoom Out One &Level"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuW 
      Caption         =   "&Window"
      HelpContextID   =   109
      WindowList      =   -1  'True
      Begin VB.Menu mnuWTileH 
         Caption         =   "Tile &Horizontally"
         HelpContextID   =   109
      End
      Begin VB.Menu mnuWTileV 
         Caption         =   "Tile &Vertically"
         HelpContextID   =   109
      End
      Begin VB.Menu mnuWCascade 
         Caption         =   "&Cascade"
         HelpContextID   =   109
      End
      Begin VB.Menu mnuWArangeIcons 
         Caption         =   "&Arrange Icons"
         HelpContextID   =   109
      End
   End
   Begin VB.Menu mnuA 
      Caption         =   "&About"
      HelpContextID   =   110
   End
End
Attribute VB_Name = "frmGraphOverlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Last modified: 02/12/2003 nt
'---------------------------------------------------------------------
Option Explicit

Const SMALL_VALUE = 0.00001     'used to fix single arithmetic bug

Const ConvLPDP = 1
Const ConvDPLP = 2

Dim bResize As Boolean      'if True allow resize event
Dim bPaint As Boolean       'if True allow paint event
Dim OldMeScaleW As Long
Dim OldMeScaleH As Long

Dim lAction As Long         'mouse action

Dim gbZoomX1 As Double      'private variables to help with
Dim gbZoomY1 As Double      'zooming, clipping, and mouse
Dim gbZoomX2 As Double      'tracking features
Dim gbZoomY2 As Double

'''Dim guShiftX As Single
'''Dim guShiftY As Single

Dim paPoints() As POINTAPI  'used to calculate Dev to Log coordinates

'coordinate system variable for the overlay
Public WithEvents MyCooSys As CooSysO
Attribute MyCooSys.VB_VarHelpID = -1

Dim MyMetricPrint  As New frmGraphOverlayPrintMetric
Attribute MyMetricPrint.VB_VarHelpID = -1

Dim mtrVX0 As Long      'viewport origin
Dim mtrVY0 As Long

Dim ScreenUpdateEnabled As Boolean

Private Sub mnuA_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnuECopyBMP_Click()
On Error GoTo err_CopyGraph
Exit Sub

err_CopyGraph:
MsgBox "Unexpected error." & vbCrLf & sErrLogReference, vbOKOnly
LogErrors Err.Number, "Copy Overlay Graph"
End Sub

Private Sub mnuECopyWMF_Click()
OverlayDrawMetafile
End Sub

Private Sub mnuEFillMassTagNames_Click()
'---------------------------------------------------------------
'fills MT tag names to appropriate structures; this works only
'if one of the overlaid displays is MT tags display
'---------------------------------------------------------------
Dim i As Long, j As Long
Dim MyNamesCnt As Long
Dim MyNames() As String
Dim MyID As Long
Dim MyNamer As mtdbMTNames
On Error Resume Next
For i = 0 To OlyCnt - 1
    If Oly(i).DisplayInd = 0 Then
       Me.MousePointer = vbHourglass
       'loading MT tag names
       MDIStatus True, "Loading MT tag names ..."
       Set MyNamer = New mtdbMTNames
       MyNamer.DBConnectionString = GelAnalysis(1).MTDB.cn.ConnectionString
       MyNamer.RetrieveSQL = glbPreferencesExpanded.MTSConnectionInfo.sqlGetMTNames
       If MyNamer.FillData Then
          If MyNamer.DataStatus = dsLoaded Then
             For j = 0 To OlyCoo(i).DataCnt - 1
'MT tags display is always loaded in the same order as MT tag arrays and always
'contains all elements of it therefore we can simply use same order here
                 MyID = AMTData(j + 1).ID
                 MyNamesCnt = MyNamer.GetNamesForMTID(MyID, MyNames)
                 If MyNamesCnt >= 0 Then
                    OlyCoo(i).Text(j) = GetMassTagNameDisplay(MyNames)
                 End If
             Next j
             MyNamer.DeleteData
             Set MyNamer = Nothing
          End If
       Else
          MsgBox "MT tag names not found.", vbOKOnly, glFGTU
       End If
       MDIStatus True, ""
       Me.MousePointer = vbDefault
       Exit Sub
    End If
Next i
MsgBox "MT tags display not found.", vbOKOnly, glFGTU
End Sub


Private Sub mnuFPrint_Click()
'---------------------------------------------------------------------
'send to printer to scale to its device context
'---------------------------------------------------------------------
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
OverlayMetrics Printer.ScaleWidth, Printer.ScaleHeight
'OverlayPaintBackColor PrtDC, Printer.ScaleWidth, Printer.ScaleHeight
OverlayCooSys PrtDC
OverlayDrawCooSys PrtDC
If OlyOptions.GRID.VertGridVisible Then OverlayDrawVGrid PrtDC
If OlyOptions.GRID.HorzGridVisible Then OverlayDrawHGrid PrtDC
OverlayDrawTextNumbersLegend PrtDC
hClipRgn = ClippingRegionA(PrtDC)
OverlayDraw PrtDC
Printer.EndDoc
Res = RestoreDC(PrtDC, OldDC)
DeleteObject (hClipRgn)
End Sub


Private Sub mnuFPrintMetric_Click()
'---------------------------------------------------------------------
'display metric printing dialog and print if User doesn't cancel
'---------------------------------------------------------------------
With MyMetricPrint
    .MinMW = MyCooSys.CurrRYMin:        .MaxMW = MyCooSys.CurrRYMax
    .MinET = MyCooSys.CurrRXMin:        .MaxET = MyCooSys.CurrRXMax
    .MinAbu = MyCooSys.RZ1:             .MaxAbu = MyCooSys.RZ2
    .IsCancel = False
    .Show vbModal
    If Not .IsCancel Then
       'here comes metric print; first zoom in if neccessary
       If (.MinMW <> MyCooSys.CurrRYMin) Or (.MaxMW <> MyCooSys.CurrRYMax) Or _
          (.MinET <> MyCooSys.CurrRXMin) Or (.MaxET <> MyCooSys.CurrRXMax) Then
          Select Case OlyOptions.Orientation
          Case OrientMWVrtETHrz
             MyCooSys.ZoomInR .MinET, .MinMW, .MaxET, .MaxMW
          Case OrientMWHrzETVrt
             MyCooSys.ZoomInR .MinMW, .MinET, .MaxMW, .MaxET
          End Select
       End If
       Printer.Orientation = vbPRORLandscape
       Printer.ScaleMode = vbPixels
       Printer.Print " "
       OverlayDrawMetric Printer.hDC
       Printer.EndDoc
    End If
End With
End Sub

Private Sub mnuFPrintSetup_Click()
PrinterSetupAPIDlg (Me.hwnd)
End Sub


Private Sub mnuFSavePic_Click()
Dim sSaveFileName As String
Dim PicSaveType As pftPictureFileTypeConstants
Dim m_cDIB As New cDIBSection   'structure used to save JPG

On Error Resume Next
sSaveFileName = FileSaveProc(Me.hwnd, Me.Caption, fstFileSaveTypeConstants.fstPIC, PicSaveType)
If sSaveFileName <> "" Then
   Select Case PicSaveType
   Case pftPictureFileTypeConstants.pftBMP
       picGraph.AutoRedraw = True
       Call picGraph_Paint
       SavePicture picGraph.Image, sSaveFileName
       picGraph.Cls
       picGraph.AutoRedraw = False
   Case pftPictureFileTypeConstants.pftJPG
       picGraph.AutoRedraw = True
       Call picGraph_Paint
       m_cDIB.LoadFromBMP picGraph.Image
       SaveJPGToFile m_cDIB, sSaveFileName
       picGraph.Cls
       picGraph.AutoRedraw = False
   Case pftPictureFileTypeConstants.pftWMF, pftPictureFileTypeConstants.pftEMF, pftPictureFileTypeConstants.pftPNG
       ''GelSaveMetafile nMyIndex, sSaveFileName
   Case Else
       MsgBox "Save picture - unknown format.", vbOKOnly
   End Select
End If
Set m_cDIB = Nothing
End Sub


Private Sub mnuTGetJiggy_Click()
frmOverlayJiggy.Show vbModal
End Sub


Private Sub mnuTZoomIn_Click()
frmZoomIn.Tag = OlyCallerID
frmZoomIn.Show vbModal
End Sub


Private Sub mnuTZoomOut_Click()
MyCooSys.ZoomOut
End Sub


Private Sub mnuTZoomOutOneLevel_Click()
MyCooSys.ZoomOut1
End Sub


Private Sub mnuVCoordinates_Click()
mnuVCoordinates.Checked = Not mnuVCoordinates.Checked
If mnuVCoordinates.Checked Then
   picCoordinates.Visible = True
Else
   picCoordinates.Visible = False
End If
End Sub

Private Sub mnuVDeviceCaps_Click()
Dim MyCaps As New frmDeviceCaps
MyCaps.ParentDC = picGraph.hDC
MyCaps.Show vbModal
Set MyCaps = Nothing
End Sub


Private Sub mnuVOrientation_Click(Index As Integer)
mnuVOrientation(Index).Checked = True
mnuVOrientation((Index + 1) Mod 2).Checked = False
OlyOptions.Orientation = Index
Me.picGraph.Refresh
End Sub



Private Sub MyCooSys_CooSysChanged()
'whatever we do with coordinate system update the grid parameters
OlyOptions.GRID.UpdateGrid MyCooSys.CurrRXMin, MyCooSys.CurrRXMax, _
                           MyCooSys.CurrRYMin, MyCooSys.CurrRYMax
CalcOverlayData
Me.picGraph.Refresh
End Sub


Private Sub MyCooSys_YScaleChange()
If MyCooSys.csYScale = glVAxisLog Then
'logarithmic scale might need to be initialized
End If
CalcOverlayData
Me.picGraph.Refresh
End Sub


Private Sub Form_Activate()
If OlyCnt > 0 Then
   MDIForm1.ProperToolbar False
   MyCooSys.LMinSz = OlyOptions.DefMinSize          'in case something changed
   MyCooSys.LMaxSz = OlyOptions.DefMaxSize
Else
   Unload Me
End If
End Sub


Private Sub Form_Load()
mnuVOrientation(OlyOptions.Orientation).Checked = True
mnuVOrientation((OlyOptions.Orientation + 1) Mod 2).Checked = False
If OlyCnt <= 0 Then frmOverlayManager.Show vbModal
Me.width = 8375
Me.Height = 6100 - GetSystemMetrics(SM_CYMENU) * Screen.TwipsPerPixelY
bResize = True
Me.Visible = True
BuildGraph
ScreenUpdateEnabled = True
Form_Resize
End Sub


Private Sub Form_Resize()
If (bResize And (Me.ScaleHeight > 0) And (Me.ScaleWidth > 0)) Then
   If Me.ScaleHeight <> OldMeScaleH Then bPaint = False
   Me.picGraph.width = Me.ScaleWidth
   bPaint = True
   Me.picGraph.Height = Me.ScaleHeight
End If
OldMeScaleH = Me.ScaleHeight
OldMeScaleW = Me.ScaleWidth
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set MyCooSys = Nothing
MDIForm1.ProperToolbar False
Unload MyMetricPrint
Set MyMetricPrint = Nothing
End Sub


Private Sub mnuWArangeIcons_Click()
MDIForm1.Arrange vbArrangeIcons
End Sub


Private Sub mnuWCascade_Click()
MDIForm1.Arrange vbCascade
End Sub


Private Sub mnuWTileH_Click()
MDIForm1.Arrange vbTileHorizontal
End Sub


Private Sub mnuWTileV_Click()
MDIForm1.Arrange vbTileVertical
End Sub


Public Function BuildGraph() As Boolean
'------------------------------------------------------------------------------------
'this procedure should be called every time we add new or remove display from overlay
'------------------------------------------------------------------------------------
Dim olyMinNET As Double, olyMaxNET As Double
Dim olyMinMW As Double, olyMaxMW As Double
Dim olyMinAbu As Double, olyMaxAbu As Double
BuildCooSys
GetOverlayLimits olyMinNET, olyMinMW, olyMinAbu, olyMaxNET, olyMaxMW, olyMaxAbu
MyCooSys.InitCooSys olyMinNET, olyMinMW, olyMinAbu, olyMaxNET, olyMaxMW, olyMaxAbu
CalcOverlayData
BuildGraph = True
End Function

Private Sub mnuFClose_Click()
Unload Me
End Sub


Private Sub mnuFManager_Click()
On Error Resume Next
frmOverlayManager.Show vbModal
MyCooSys.LMinSz = OlyOptions.DefMinSize          'in case something changed
MyCooSys.LMaxSz = OlyOptions.DefMaxSize
End Sub


Private Sub picGraph_GotFocus()
If bPaint Then Me.picGraph.Refresh
bPaint = False
End Sub


Private Sub picGraph_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Res As Long
Dim rcClip As Rect
Dim paClip As POINTAPI

ReDim paPoints(0)
paPoints(0).x = x:          paPoints(0).y = y
DevLogConversion ConvDPLP, 1
lAction = glActionZoom
Select Case Button
Case vbLeftButton
  If lAction = glActionZoom Then
     gbZoomX1 = x:           gbZoomY1 = y
     gbZoomX2 = gbZoomX1:    gbZoomY2 = gbZoomY1
    'clip cursor on viewport
     paPoints(0).x = gbZoomX1:     paPoints(0).y = gbZoomY1
     DevLogConversion ConvDPLP, 1
     If (paPoints(0).x < LoX0) Or (paPoints(0).x > LoXE) Or _
        (paPoints(0).y < LoY0) Or (paPoints(0).y > LoYE) Then
        lAction = glNoAction
        Exit Sub
     End If
     paClip.x = 0:     paClip.y = 0
     Res = ClientToScreen(picGraph.hwnd, paClip)
     MyCooSys.GetViewPortRectangle paClip.x, paClip.y, rcClip.Top, rcClip.Left, rcClip.Bottom, rcClip.Right
     Res = ClipCursor(rcClip)
     picGraph.DrawStyle = vbDot
  End If
Case Else
  lAction = glNoAction
End Select
End Sub


Private Sub picGraph_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ReDim paPoints(0)
Dim sngX As Single, sngY As Single
paPoints(0).x = x:      paPoints(0).y = y
DevLogConversion ConvDPLP, 1
If picCoordinates.Visible Then
   sngX = CSng(paPoints(0).x)
   sngY = CSng(paPoints(0).y)
   Select Case OlyOptions.Orientation
   Case OrientMWHrzETVrt
        Call MyCooSys.LPToRP1(sngY, sngX)
   Case OrientMWVrtETHrz
        Call MyCooSys.LPToRP1(sngX, sngY)
   End Select
   lblHCoo.Caption = Format$(sngX, "0.0000") & Chr$(32)
   lblVCoo.Caption = Format$(sngY, "0.0000") & Chr$(32)
End If
If lAction = glActionZoom Then
   picGraph.Line (gbZoomX1, gbZoomY1)-(gbZoomX2, gbZoomY2), , B
   gbZoomX2 = x
   gbZoomY2 = y
   picGraph.Line (gbZoomX1, gbZoomY1)-(gbZoomX2, gbZoomY2), , B
End If
End Sub


Private Sub picGraph_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Res As Long
ReDim paPoints(1)
Select Case lAction
Case glNoAction             'remove restrictions on cursor
  If Button = vbLeftButton Then Res = ClipCursorByNum(&O0)
Case glActionZoom
  If Button = vbLeftButton Then
     Res = ClipCursorByNum(&O0)
     picGraph.Line (gbZoomX1, gbZoomY1)-(gbZoomX2, gbZoomY2), , B
     picGraph.DrawStyle = vbSolid
     If (Abs(gbZoomX1 - gbZoomX2) > 10) And (Abs(gbZoomY1 - gbZoomY2) > 10) Then
        paPoints(0).x = gbZoomX1:        paPoints(0).y = gbZoomY1
        paPoints(1).x = gbZoomX2:        paPoints(1).y = gbZoomY2
        DevLogConversion ConvDPLP, 2
        Select Case OlyOptions.Orientation
        Case OrientMWVrtETHrz
             MyCooSys.ZoomIn paPoints(0).x, paPoints(0).y, paPoints(1).x, paPoints(1).y
        Case OrientMWHrzETVrt
             MyCooSys.ZoomIn paPoints(0).y, paPoints(0).x, paPoints(1).y, paPoints(1).x
        End Select
     End If
  End If
End Select
lAction = glNoAction
End Sub


Private Sub picGraph_Paint()
On Error Resume Next
If ScreenUpdateEnabled Then OverlayDrawScreen
End Sub


Private Sub picGraph_Resize()
On Error Resume Next
If bPaint Then Me.picGraph.Refresh
Me.picCoordinates.Left = Me.picGraph.ScaleWidth - Me.picCoordinates.width
End Sub


Public Sub BuildCooSys()
Set MyCooSys = New CooSysO
MyCooSys.BuildingCS = True
MyCooSys.csOriginXY = glOriginBL
MyCooSys.csXOrient = glNormal
MyCooSys.csYOrient = glNormal
MyCooSys.csYScale = glVAxisLin
MyCooSys.LMinSz = OlyOptions.DefMinSize
MyCooSys.LMaxSz = OlyOptions.DefMaxSize
MyCooSys.BuildingCS = False
End Sub


Public Function CalcOverlayData() As Boolean
'----------------------------------------------------------------
'recalculates all overlay data based on current coordinate system
'----------------------------------------------------------------
Dim i As Long
For i = 0 To OlyCnt - 1
    If Oly(i).Visible Then CalcOverlayDataOne (i)
Next i
End Function


Public Function CalcOverlayDataOne(OlyInd) As Boolean
'------------------------------------------------------------------------
'recalculates coordinates of spots for Oly(OlyInd)
'------------------------------------------------------------------------
Dim CurrMW As Double, CurrNET As Double, CurrAbu As Double
Dim i As Long, TmpCnt As Long
Dim LogRect As Rect
Dim ChP() As LaV2DGPoint
On Error Resume Next      'there will be a whole bunch of overflow errors
'------------------------------------------------------------------------
'point is visible if within logical borders; unique mass class is visible
'if any of its characteristic points within logical borders
LogRect.Left = LoX0:       LogRect.Top = LoY0
LogRect.Right = LoXE:      LogRect.Bottom = LoYE
'------------------------------------------------------------------------
Select Case Oly(OlyInd).Type
Case olySolo
  Select Case Oly(OlyInd).Shape
  Case olyStick             'stick is unique since it does not care about size and has adjustment for NET
     TmpCnt = 0
     With Oly(OlyInd)
        For i = 1 To GelData(.DisplayInd).CSLines
            TmpCnt = TmpCnt + 1
            CurrMW = GelData(.DisplayInd).CSData(i).AverageMW
            CurrNET = .NETSlope * GelData(.DisplayInd).CSData(i).ScanNumber + .NETIntercept
            OlyCoo(OlyInd).x(TmpCnt - 1) = CSng(LoX1 + MyCooSys.ScaleX_RToL * (CurrNET - MyCooSys.CurrRXMin))
            OlyCoo(OlyInd).y(TmpCnt - 1) = CSng(LoY1 + MyCooSys.ScaleY_RToL * (CurrMW - MyCooSys.CurrRYMin))
            OlyCoo(OlyInd).XL(TmpCnt - 1) = Abs(CSng(MyCooSys.ScaleX_RToL * OlyAdj(OlyInd).NETL(TmpCnt - 1)))
            OlyCoo(OlyInd).XU(TmpCnt - 1) = Abs(CSng(MyCooSys.ScaleX_RToL * OlyAdj(OlyInd).NETU(TmpCnt - 1)))
            OlyCoo(OlyInd).YL(TmpCnt - 1) = OlyCoo(OlyInd).y(TmpCnt - 1)
            OlyCoo(OlyInd).YU(TmpCnt - 1) = OlyCoo(OlyInd).y(TmpCnt - 1)
            If PtInRect(LogRect, OlyCoo(OlyInd).x(TmpCnt - 1), OlyCoo(OlyInd).y(TmpCnt - 1)) = 0 Then
               OlyCoo(OlyInd).OutOfScope(TmpCnt - 1) = True
            Else
               OlyCoo(OlyInd).OutOfScope(TmpCnt - 1) = False
            End If
        Next i
        For i = 1 To GelData(.DisplayInd).IsoLines
            TmpCnt = TmpCnt + 1
            CurrMW = GetIsoMass(GelData(.DisplayInd).IsoData(i), GelData(.DisplayInd).Preferences.IsoDataField)
            CurrNET = .NETSlope * GelData(.DisplayInd).IsoData(i).ScanNumber + .NETIntercept
            OlyCoo(OlyInd).x(TmpCnt - 1) = CSng(LoX1 + MyCooSys.ScaleX_RToL * (CurrNET - MyCooSys.CurrRXMin))
            OlyCoo(OlyInd).y(TmpCnt - 1) = CSng(LoY1 + MyCooSys.ScaleY_RToL * (CurrMW - MyCooSys.CurrRYMin))
            OlyCoo(OlyInd).XL(TmpCnt - 1) = Abs(CSng(MyCooSys.ScaleX_RToL * OlyAdj(OlyInd).NETL(TmpCnt - 1)))
            OlyCoo(OlyInd).XU(TmpCnt - 1) = Abs(CSng(MyCooSys.ScaleX_RToL * OlyAdj(OlyInd).NETU(TmpCnt - 1)))
            OlyCoo(OlyInd).YL(TmpCnt - 1) = OlyCoo(OlyInd).y(TmpCnt - 1)
            OlyCoo(OlyInd).YU(TmpCnt - 1) = OlyCoo(OlyInd).y(TmpCnt - 1)
            If PtInRect(LogRect, OlyCoo(OlyInd).x(TmpCnt - 1), OlyCoo(OlyInd).y(TmpCnt - 1)) = 0 Then
               OlyCoo(OlyInd).OutOfScope(TmpCnt - 1) = True
            Else
               OlyCoo(OlyInd).OutOfScope(TmpCnt - 1) = False
            End If
        Next i
     End With
  Case olyBox, olyBoxEmpty, olySpot, olySpotEmpty, olyTriangle, olyTriangleEmpty, olyTriStar
     TmpCnt = 0
     With Oly(OlyInd)
        If .UniformSize Then
           For i = 1 To GelData(.DisplayInd).CSLines
               TmpCnt = TmpCnt + 1
               CurrMW = GelData(.DisplayInd).CSData(i).AverageMW
               CurrNET = .NETSlope * GelData(.DisplayInd).CSData(i).ScanNumber + .NETIntercept
               CurrAbu = GelData(.DisplayInd).CSData(i).Abundance
               OlyCoo(OlyInd).x(TmpCnt - 1) = CSng(LoX1 + MyCooSys.ScaleX_RToL * (CurrNET - MyCooSys.CurrRXMin))
               OlyCoo(OlyInd).y(TmpCnt - 1) = CSng(LoY1 + MyCooSys.ScaleY_RToL * (CurrMW - MyCooSys.CurrRYMin))
               OlyCoo(OlyInd).R(TmpCnt - 1) = MyCooSys.UniSpotSize
               If PtInRect(LogRect, OlyCoo(OlyInd).x(TmpCnt - 1), OlyCoo(OlyInd).y(TmpCnt - 1)) = 0 Then
                  OlyCoo(OlyInd).OutOfScope(TmpCnt - 1) = True
               Else
                  OlyCoo(OlyInd).OutOfScope(TmpCnt - 1) = False
               End If
           Next i
           For i = 1 To GelData(.DisplayInd).IsoLines
               TmpCnt = TmpCnt + 1
               CurrMW = GetIsoMass(GelData(.DisplayInd).IsoData(i), GelData(.DisplayInd).Preferences.IsoDataField)
               CurrNET = .NETSlope * GelData(.DisplayInd).IsoData(i).ScanNumber + .NETIntercept
               CurrAbu = GelData(.DisplayInd).IsoData(i).Abundance
               OlyCoo(OlyInd).x(TmpCnt - 1) = CSng(LoX1 + MyCooSys.ScaleX_RToL * (CurrNET - MyCooSys.CurrRXMin))
               OlyCoo(OlyInd).y(TmpCnt - 1) = CSng(LoY1 + MyCooSys.ScaleY_RToL * (CurrMW - MyCooSys.CurrRYMin))
               OlyCoo(OlyInd).R(TmpCnt - 1) = MyCooSys.UniSpotSize
               If PtInRect(LogRect, OlyCoo(OlyInd).x(TmpCnt - 1), OlyCoo(OlyInd).y(TmpCnt - 1)) = 0 Then
                  OlyCoo(OlyInd).OutOfScope(TmpCnt - 1) = True
               Else
                  OlyCoo(OlyInd).OutOfScope(TmpCnt - 1) = False
               End If
           Next i
        Else
           For i = 1 To GelData(.DisplayInd).CSLines
               TmpCnt = TmpCnt + 1
               CurrMW = GelData(.DisplayInd).CSData(i).AverageMW
               CurrNET = .NETSlope * GelData(.DisplayInd).CSData(i).ScanNumber + .NETIntercept
               CurrAbu = GelData(.DisplayInd).CSData(i).Abundance
               OlyCoo(OlyInd).x(TmpCnt - 1) = CSng(LoX1 + MyCooSys.ScaleX_RToL * (CurrNET - MyCooSys.CurrRXMin))
               OlyCoo(OlyInd).y(TmpCnt - 1) = CSng(LoY1 + MyCooSys.ScaleY_RToL * (CurrMW - MyCooSys.CurrRYMin))
               OlyCoo(OlyInd).R(TmpCnt - 1) = CSng((CurrAbu - MyCooSys.RZ1) * MyCooSys.ScaleR + lDfMinSz)
               If PtInRect(LogRect, OlyCoo(OlyInd).x(TmpCnt - 1), OlyCoo(OlyInd).y(TmpCnt - 1)) = 0 Then
                  OlyCoo(OlyInd).OutOfScope(TmpCnt - 1) = True
               Else
                  OlyCoo(OlyInd).OutOfScope(TmpCnt - 1) = False
               End If
           Next i
           For i = 1 To GelData(.DisplayInd).IsoLines
               TmpCnt = TmpCnt + 1
               CurrMW = GetIsoMass(GelData(.DisplayInd).IsoData(i), GelData(.DisplayInd).Preferences.IsoDataField)
               CurrNET = .NETSlope * GelData(.DisplayInd).IsoData(i).ScanNumber + .NETIntercept
               CurrAbu = GelData(.DisplayInd).IsoData(i).Abundance
               OlyCoo(OlyInd).x(TmpCnt - 1) = CSng(LoX1 + MyCooSys.ScaleX_RToL * (CurrNET - MyCooSys.CurrRXMin))
               OlyCoo(OlyInd).y(TmpCnt - 1) = CSng(LoY1 + MyCooSys.ScaleY_RToL * (CurrMW - MyCooSys.CurrRYMin))
               OlyCoo(OlyInd).R(TmpCnt - 1) = CSng((CurrAbu - MyCooSys.RZ1) * MyCooSys.ScaleR + lDfMinSz)
               If PtInRect(LogRect, OlyCoo(OlyInd).x(TmpCnt - 1), OlyCoo(OlyInd).y(TmpCnt - 1)) = 0 Then
                  OlyCoo(OlyInd).OutOfScope(TmpCnt - 1) = True
               Else
                  OlyCoo(OlyInd).OutOfScope(TmpCnt - 1) = False
               End If
           Next i
        End If
     End With
   End Select
Case OlyUMC
   With Oly(OlyInd)
     If .UniformSize Then
        For i = 0 To GelUMC(.DisplayInd).UMCCnt - 1
            If fUMCCharacteristicPoints(.DisplayInd, i, ChP()) Then
               CurrNET = .NETSlope * ChP(0).Scan + .NETIntercept
               OlyCoo(OlyInd).XL(i) = CSng(LoX1 + MyCooSys.ScaleX_RToL * (CurrNET - MyCooSys.CurrRXMin))
               CurrNET = .NETSlope * ChP(2).Scan + .NETIntercept
               OlyCoo(OlyInd).XU(i) = CSng(LoX1 + MyCooSys.ScaleX_RToL * (CurrNET - MyCooSys.CurrRXMin))
               CurrNET = .NETSlope * ChP(1).Scan + .NETIntercept
               OlyCoo(OlyInd).x(i) = CSng(LoX1 + MyCooSys.ScaleX_RToL * (CurrNET - MyCooSys.CurrRXMin))
               'end points draw as class representative mass and class representative size higher
               OlyCoo(OlyInd).YL(i) = CSng(LoY1 + MyCooSys.ScaleY_RToL * (ChP(1).MW - MyCooSys.CurrRYMin))
               OlyCoo(OlyInd).YU(i) = OlyCoo(OlyInd).YL(i)
               OlyCoo(OlyInd).y(i) = OlyCoo(OlyInd).YL(i) + MyCooSys.UniSpotSize
               'UMC should be drawn if any of its points is visible(optimized for speed)
               OlyCoo(OlyInd).OutOfScope(i) = True
               If PtInRect(LogRect, OlyCoo(OlyInd).x(i), OlyCoo(OlyInd).y(i)) <> 0 Then
                  OlyCoo(OlyInd).OutOfScope(i) = False
               Else
                  If PtInRect(LogRect, OlyCoo(OlyInd).XL(i), OlyCoo(OlyInd).YL(i)) <> 0 Then
                     OlyCoo(OlyInd).OutOfScope(i) = False
                  Else
                     If PtInRect(LogRect, OlyCoo(OlyInd).XU(i), OlyCoo(OlyInd).YU(i)) <> 0 Then OlyCoo(OlyInd).OutOfScope(i) = False
                  End If
               End If
            Else
               OlyCoo(OlyInd).R(i) = -1
            End If
        Next i
     Else
        Dim MyScaleSize As Single
        Dim MaxScaleSize As Single
        MaxScaleSize = ((Log(MyCooSys.RZ2) / Log(10) - Log(MyCooSys.RZ1) / Log(10)))
        For i = 0 To GelUMC(.DisplayInd).UMCCnt - 1
            If fUMCCharacteristicPoints(.DisplayInd, i, ChP()) Then
               CurrNET = .NETSlope * ChP(0).Scan + .NETIntercept
               OlyCoo(OlyInd).XL(i) = CSng(LoX1 + MyCooSys.ScaleX_RToL * (CurrNET - MyCooSys.CurrRXMin))
               CurrNET = .NETSlope * ChP(2).Scan + .NETIntercept
               OlyCoo(OlyInd).XU(i) = CSng(LoX1 + MyCooSys.ScaleX_RToL * (CurrNET - MyCooSys.CurrRXMin))
               CurrNET = .NETSlope * ChP(1).Scan + .NETIntercept
               OlyCoo(OlyInd).x(i) = CSng(LoX1 + MyCooSys.ScaleX_RToL * (CurrNET - MyCooSys.CurrRXMin))
               'end points draw as class representative mass and class representative size higher
               OlyCoo(OlyInd).YL(i) = CSng(LoY1 + MyCooSys.ScaleY_RToL * (ChP(1).MW - MyCooSys.CurrRYMin))
               OlyCoo(OlyInd).YU(i) = OlyCoo(OlyInd).YL(i)
               'OlyCoo(OlyInd).R(i) = MyCooSys.ScaleR * (ChP(1).Abu - MyCooSys.RZ1)
               'If OlyCoo(OlyInd).R(i) < MyCooSys.LMinSz Then OlyCoo(OlyInd).R(i) = MyCooSys.LMinSz
               'If OlyCoo(OlyInd).R(i) > MyCooSys.LMaxSz Then OlyCoo(OlyInd).R(i) = MyCooSys.LMaxSz
               'OlyCoo(OlyInd).y(i) = OlyCoo(OlyInd).YL(i) + OlyCoo(OlyInd).R(i)
               
               'scalling logarithmic
               MyScaleSize = ((Log(ChP(1).Abu) / Log(10) - Log(MyCooSys.RZ1) / Log(10)))
               OlyCoo(OlyInd).R(i) = MyCooSys.LMinSz + (MyScaleSize / MaxScaleSize) * MyCooSys.LMaxSz
               OlyCoo(OlyInd).y(i) = OlyCoo(OlyInd).YL(i) + OlyCoo(OlyInd).R(i)
               
               'UMC should be drawn if any of its points is visible(optimized for speed)
               OlyCoo(OlyInd).OutOfScope(i) = True
               If PtInRect(LogRect, OlyCoo(OlyInd).x(i), OlyCoo(OlyInd).y(i)) <> 0 Then
                  OlyCoo(OlyInd).OutOfScope(i) = False
               Else
                  If PtInRect(LogRect, OlyCoo(OlyInd).XL(i), OlyCoo(OlyInd).YL(i)) <> 0 Then
                     OlyCoo(OlyInd).OutOfScope(i) = False
                  Else
                     If PtInRect(LogRect, OlyCoo(OlyInd).XU(i), OlyCoo(OlyInd).YU(i)) <> 0 Then OlyCoo(OlyInd).OutOfScope(i) = False
                  End If
               End If
            Else
               OlyCoo(OlyInd).R(i) = -1
            End If
        Next i
     End If
   End With
End Select
End Function


Private Sub OverlayCooSys(ByVal hDC As Long)
Dim Res As Long
Dim ptPoint As POINTAPI
Dim szSize As Size
On Error Resume Next
Res = SetMapMode(hDC, MM_ANISOTROPIC)
'logical window
Res = SetWindowOrgEx(hDC, LoX0, LoY0, ptPoint)
Res = SetWindowExtEx(hDC, LoXE, LoYE, szSize)
'viewport
Res = SetViewportOrgEx(hDC, MyCooSys.VX0, MyCooSys.VY0, ptPoint)
Res = SetViewportExtEx(hDC, MyCooSys.VXE, MyCooSys.VYE, szSize)
End Sub


Private Sub OverlayMetrics(ByVal dcw As Long, ByVal dch As Long)
With MyCooSys
    .SXOffset = dcw * loSXPercent:    .SYOffset = dch * loSYPercent
    .LXOffset = dcw * loLXPercent:    .LYOffset = dch * loLYPercent
    Select Case .csOriginXY
    Case glOriginBL             'axes intersection coordinates
        .XYX0 = .LXOffset:        .XYY0 = dch - .LYOffset
        .XYXE = dcw - .SXOffset:  .XYYE = .SYOffset
    Case glOriginBR
        .XYX0 = dcw - .LXOffset:  .XYY0 = dch - .LYOffset
        .XYXE = .SXOffset:        .XYYE = .SYOffset
    Case glOriginTL
        .XYX0 = .LXOffset:        .XYY0 = .LYOffset
        .XYXE = dcw - .SXOffset:  .XYYE = dch - .SYOffset
    Case glOriginTR
        .XYX0 = dcw - .LXOffset:  .XYY0 = .LYOffset
        .XYXE = .SXOffset:        .XYYE = dch - .SYOffset
    End Select
    'origin coordinates - see explanation in the documentation
    If (.csOrigin + .csOriginXY) Mod 2 = 0 Then
       .OrX0 = .XYX0:       .OrXE = .XYXE
    Else
       .OrX0 = .XYXE:       .OrXE = .XYX0
    End If
    If (.csOrigin < 3 And .csOriginXY < 3) Or (.csOrigin > 2 And .csOriginXY > 2) Then
       .OrY0 = .XYY0:       .OrYE = .XYYE
    Else
       .OrY0 = .XYYE:       .OrYE = .XYY0
    End If
    'viewport coordinates
    .VX0 = CLng(.OrX0):    .VXE = CLng(.OrXE - .OrX0)
    .VY0 = CLng(.OrY0):    .VYE = CLng(.OrYE - .OrY0)
End With
End Sub


Private Sub OverlayPaintBackColor(ByVal hDC As Long, ByVal w As Long, ByVal h As Long)
Dim OldBrush As Long
Dim Res As Long
On Error Resume Next
OldBrush = SelectObject(hDC, hOlyBackClrBrush)
Res = PatBlt(hDC, 0, 0, w, h, PATCOPY)
Res = SelectObject(hDC, OldBrush)
End Sub


Private Sub OverlayDrawCooSys(ByVal hDC As Long)
Dim OldPen As Long
Dim ptPoint As POINTAPI
Dim Res As Long
On Error Resume Next
OldPen = SelectObject(hDC, hOlyForeClrPen)
Res = SetROP2(hDC, R2_COPYPEN)
Res = SetBkMode(hDC, TRANSPARENT)
With MyCooSys
    'horizontal axis
    If (.csOrigin < 3 And .csOriginXY < 3) Or _
       (.csOrigin > 2 And .csOriginXY > 2) Then
       Res = MoveToEx(hDC, LoX0, LoY0, ptPoint)
       Res = LineTo(hDC, LoX0 + LoXE, LoY0)
    Else
       Res = MoveToEx(hDC, LoX0, LoY0 + LoYE, ptPoint)
       Res = LineTo(hDC, LoX0 + LoXE, LoY0 + LoYE)
    End If
    'vertical axis
    If (.csOrigin + .csOriginXY) Mod 2 = 0 Then
       Res = MoveToEx(hDC, LoX0, LoY0, ptPoint)
       Res = LineTo(hDC, LoX0, LoY0 + LoYE)
    Else
       Res = MoveToEx(hDC, LoX0 + LoXE, LoY0, ptPoint)
       Res = LineTo(hDC, LoX0 + LoXE, LoY0 + LoYE)
    End If
End With
Res = SelectObject(hDC, OldPen)
End Sub


Private Sub OverlayDrawHNumbers(ByVal hDC As Long)
Dim i As Long, iSign As Long
Dim sNumber As String
Dim szNumber As Size
Dim rStep As Single
Dim lStep As Long, lx As Long
Dim NumT As Long
Dim MarkO As Long               'X coordinates of the thick mark (out)
Dim MarkL As Long               'X coordinates of the thick mark (on axis)
Dim ptPoint As POINTAPI
Dim vSelVec As Variant
Dim Res As Long
On Error Resume Next

If (MyCooSys.csOrigin < 3 And MyCooSys.csOriginXY < 3) Or _
   (MyCooSys.csOrigin > 2 And MyCooSys.csOriginXY > 2) Then
    MarkO = LoY0 - 30:       MarkL = LoY0
Else
    MarkO = LoYE + 30:       MarkL = LoYE
End If
If MyCooSys.csOrigin Mod 2 = 0 Then
    iSign = 1
Else
    iSign = -1
End If
rStep = (MyCooSys.CurrRXMax - MyCooSys.CurrRXMin) / 8
lStep = (LoX2 - LoX1) \ 8
Res = GetTextExtentPoint32(hDC, "0123456789", 10, szNumber)
vSelVec = OverlayGetJobVector(vAxYSelectMatrix, vYHNumJobMatrix)
NumT = SP(vSelVec, MarkL, 30, loSY \ 8, szNumber.cy)
For i = 0 To 8
    lx = LoX1 + i * lStep
    sNumber = Format$(MyCooSys.CurrRXMin + i * rStep, "0.0000")
    Res = GetTextExtentPoint32(hDC, sNumber, Len(sNumber), szNumber)
    Res = MoveToEx(hDC, lx, MarkO, ptPoint)
    Res = LineTo(hDC, lx, MarkL)
    Res = TextOut(hDC, lx + iSign * szNumber.cx \ 2, NumT, sNumber, Len(sNumber))
Next i
End Sub


Private Sub OverlayDrawVNumbers(ByVal hDC As Long)
Dim i As Long, iSign As Long
Dim sNumber As String
Dim szNumber As Size
Dim rStep As Single
Dim lStep As Long, ly As Long
Dim NumL As Long
Dim MarkO As Long               'X coordinates of the thick mark (out)
Dim MarkL As Long               'X coordinates of the thick mark (on axis)
Dim ptPoint As POINTAPI
Dim vSelVec As Variant
Dim Res As Long
On Error Resume Next

If (MyCooSys.csOrigin + MyCooSys.csOriginXY) Mod 2 = 0 Then
    MarkO = LoX0 - 30:       MarkL = LoX0
Else
    MarkO = LoXE + 30:       MarkL = LoXE
End If
If MyCooSys.csOrigin > 2 Then
    iSign = -1
Else
    iSign = 1
End If
rStep = (MyCooSys.CurrRYMax - MyCooSys.CurrRYMin) / 8
lStep = (LoY2 - LoY1) \ 8
For i = 0 To 8
    ly = LoY1 + i * lStep
    Select Case MyCooSys.csYScale
    Case glVAxisLin
        sNumber = Format$(MyCooSys.CurrRYMin + i * rStep, "#,###,##0.0000")
    Case glVAxisLog
        sNumber = Format$(10 ^ (MyCooSys.CurrRYMin + i * rStep), "#,###,##0.0000")
    End Select
    Res = GetTextExtentPoint32(hDC, sNumber, Len(sNumber), szNumber)
    Res = MoveToEx(hDC, MarkO, ly, ptPoint)
    Res = LineTo(hDC, MarkL, ly)
    vSelVec = OverlayGetJobVector(vAxXSelectMatrix, vXVNumJobMatrix)
    NumL = SP(vSelVec, MarkL, 30, LoSX \ 4, szNumber.cx)
    Res = TextOut(hDC, NumL, ly + iSign * szNumber.cy \ 2, sNumber, Len(sNumber))
Next i
End Sub


Private Sub OverlayDrawLegend(ByVal hDC As Long)
Dim sHLbl As String, sVLbl As String
Dim szHLbl As Size, szVLbl As Size
Dim Res As Long
On Error Resume Next
sHLbl = "NET"
Select Case MyCooSys.csYScale
Case glVAxisLin
    sVLbl = "MW(Linear scale)"
Case glVAxisLog
    sVLbl = "MW(Logarithmic scale)"
End Select
'get the sizes of the axes and legend labels
Res = GetTextExtentPoint32(hDC, sHLbl, Len(sHLbl), szHLbl)
Res = GetTextExtentPoint32(hDC, sVLbl, Len(sVLbl), szVLbl)
'draw coordinate axes labels
Res = TextOut(hDC, LoXE - szHLbl.cx, LoY0 - szHLbl.cy - CLng(loSY / 2), sHLbl, Len(sHLbl))
Res = TextOut(hDC, LoX0 + CLng(LoSX / 2), LoYE + szVLbl.cy, sVLbl, Len(sVLbl))
End Sub


Private Sub OverlayDrawLegendRev(ByVal hDC As Long)
Dim sHLbl As String, sVLbl As String
Dim szHLbl As Size, szVLbl As Size
Dim Res As Long
On Error Resume Next
sVLbl = "NET"
Select Case MyCooSys.csYScale
Case glVAxisLin
    sHLbl = "MW(Linear scale)"
Case glVAxisLog
    sHLbl = "MW(Logarithmic scale)"
End Select
'get the sizes of the axes and legend labels
Res = GetTextExtentPoint32(hDC, sHLbl, Len(sHLbl), szHLbl)
Res = GetTextExtentPoint32(hDC, sVLbl, Len(sVLbl), szVLbl)
'draw coordinate axes labels
Res = TextOut(hDC, LoXE - szHLbl.cx, LoY0 - szHLbl.cy - CLng(loSY / 2), sHLbl, Len(sHLbl))
Res = TextOut(hDC, LoX0 + CLng(LoSX / 2), LoYE + szVLbl.cy, sVLbl, Len(sVLbl))
End Sub


Private Sub OverlayDrawTextNumbersLegend(ByVal hDC As Long)
'-------------------------------------------------------------------
'draws anything that includes text/numbers for the overlay display
'-------------------------------------------------------------------
Dim ldffont As Long, lOldFont As Long, lNewFont As Long
Dim lfLogFont As LOGFONT
Dim OldBrush As Long
Dim OldPen As Long
Dim Res As Long
On Error Resume Next
'create new font suitable for writing on the graph
ldffont = SelectObject(Me.picGraph.hDC, GetStockObject(SYSTEM_FONT))
Res = GetObjectAPI(ldffont, Len(lfLogFont), lfLogFont)
Res = SelectObject(Me.picGraph.hDC, ldffont)
'create new logical font
lfLogFont.lfWidth = CLng(OlyOptions.DefFontWidth * (LoXE - LoX0))
lfLogFont.lfHeight = CLng(OlyOptions.DefFontHeight * (LoYE - LoY0))
lNewFont = CreateFontIndirect(lfLogFont)
'select newly created logical font to DC
lOldFont = SelectObject(hDC, lNewFont)
Res = SetTextColor(hDC, OlyOptions.ForeColor)
Res = SetROP2(hDC, R2_COPYPEN)
Res = SetBkMode(hDC, TRANSPARENT)
Res = SelectObject(hDC, hOlyForeClrPen)
Select Case OlyOptions.Orientation
Case OrientMWVrtETHrz
     OverlayDrawLegend hDC
     OverlayDrawHNumbers hDC
     OverlayDrawVNumbers hDC
Case OrientMWHrzETVrt
     OverlayDrawLegendRev hDC
     OverlayDrawHNumbersRev hDC
     OverlayDrawVNumbersRev hDC
End Select
Res = SelectObject(hDC, OldPen)
Res = SelectObject(hDC, OldBrush)
Res = SelectObject(hDC, lOldFont)
If lNewFont <> 0 Then DeleteObject (lNewFont)
End Sub


Private Function OverlayGetJobVector(ByVal SelectMatrix As Variant, _
                                     ByVal JobMatrix As Variant) As Variant
'----------------------------------------------------------------------------
'returns vector (variant array) appropriate for this job and coosys selection
'----------------------------------------------------------------------------
Dim JobColumn As Integer
Dim i As Long
Dim aRes(1 To 4) As Long
On Error GoTo err_GetJobVector
JobColumn = SelectMatrix(MyCooSys.csOriginXY, MyCooSys.csOrigin)
For i = 1 To 4
    aRes(i) = JobMatrix(i, JobColumn)
Next i
OverlayGetJobVector = aRes
Exit Function
err_GetJobVector:
OverlayGetJobVector = Null
End Function


Private Sub DevLogConversion(ByVal ConversionType As Integer, ByVal NumOfPoints As Integer)
Dim OldDC As Long, Res As Long
OldDC = SaveDC(picGraph.hDC)
OverlayCooSys picGraph.hDC
Select Case ConversionType
Case ConvDPLP
    Res = DPtoLP(picGraph.hDC, paPoints(0), NumOfPoints)
Case ConvLPDP
    Res = LPtoDP(picGraph.hDC, paPoints(0), NumOfPoints)
End Select
Res = RestoreDC(picGraph.hDC, OldDC)
End Sub


Public Sub OverlayDrawScreen()
Dim MyDC As Long
Dim OldDC As Long
Dim hClipRgn As Long
Dim Res As Long
On Error Resume Next
Me.MousePointer = vbHourglass
MyDC = Me.picGraph.hDC
OldDC = SaveDC(MyDC)
OverlayMetrics Me.picGraph.ScaleWidth, Me.picGraph.ScaleHeight
OverlayPaintBackColor MyDC, Me.picGraph.ScaleWidth, Me.picGraph.ScaleHeight
OverlayCooSys MyDC
OverlayDrawCooSys MyDC
If OlyOptions.GRID.HorzGridVisible Then OverlayDrawHGrid MyDC
If OlyOptions.GRID.VertGridVisible Then OverlayDrawVGrid MyDC
OverlayDrawTextNumbersLegend MyDC
hClipRgn = ClippingRegionA(MyDC)
OverlayDraw MyDC
Res = RestoreDC(MyDC, OldDC)
DeleteObject (hClipRgn)
Me.MousePointer = vbDefault
End Sub


Public Sub OverlayDrawMetafile()
'-------------------------------------------------------------
'draws overlay on enhanced metafile and copy it at clipboard
'-------------------------------------------------------------
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
Dim Res As Long
On Error Resume Next

hRefDC = picGraph.hDC
iWidthMM = GetDeviceCaps(hRefDC, HORZSIZE)
iHeightMM = GetDeviceCaps(hRefDC, VERTSIZE)
iWidthPels = GetDeviceCaps(hRefDC, HORZRES)
iHeightPels = GetDeviceCaps(hRefDC, VERTRES)

iMMPerPelX = (iWidthMM * 100) / iWidthPels
iMMPerPelY = (iHeightMM * 100) / iHeightPels

rcRef.Top = 0:                            rcRef.Left = 0
rcRef.Bottom = picGraph.ScaleHeight:      rcRef.Right = picGraph.ScaleWidth
'convert to himetric units
rcRef.Left = rcRef.Left * iMMPerPelX:     rcRef.Top = rcRef.Top * iMMPerPelY
rcRef.Right = rcRef.Right * iMMPerPelX:   rcRef.Bottom = rcRef.Bottom * iMMPerPelY

emfDC = CreateEnhMetaFile(hRefDC, vbNullString, rcRef, vbNullString)
OldDC = SaveDC(emfDC)
OverlayMetrics picGraph.ScaleWidth, picGraph.ScaleHeight
'OverlayPaintBackColor emfDC, picGraph.ScaleWidth, picGraph.ScaleHeight
OverlayCooSys emfDC
OverlayDrawCooSys emfDC
If OlyOptions.GRID.VertGridVisible Then OverlayDrawVGrid emfDC
If OlyOptions.GRID.HorzGridVisible Then OverlayDrawHGrid emfDC
OverlayDrawTextNumbersLegend emfDC
hClipRgn = ClippingRegionA(emfDC)
OverlayDraw emfDC
Res = RestoreDC(emfDC, OldDC)
If hClipRgn <> 0 Then DeleteObject (hClipRgn)
emfHandle = CloseEnhMetaFile(emfDC)
Res = OpenClipboard(picGraph.hwnd)
Res = EmptyClipboard()
Res = SetClipboardData(CF_ENHMETAFILE, emfHandle)
Res = CloseClipboard
End Sub


Private Function OverlayDraw(ByVal hDC As Long) As Boolean
Dim TextHeight As Long
Dim TextWidth As Long
Dim i As Long
Dim ZOrder() As Long
If GetOlyZOrder(ZOrder()) Then
   Select Case OlyOptions.Orientation
   Case OrientMWVrtETHrz
     For i = 0 To OlyCnt - 1
       If Oly(ZOrder(i)).Visible Then
          TextHeight = CLng(OlyOptions.DefFontHeight * (LoYE - LoY0) * Oly(ZOrder(i)).TextHeightPct)
          TextWidth = CLng(OlyOptions.DefFontWidth * (LoXE - LoX0) * Oly(ZOrder(i)).TextHeightPct)
          Select Case Oly(ZOrder(i)).Type
          Case olySolo
               Select Case Oly(ZOrder(i)).Shape
               Case olyBox
                    OlyDrawBox ZOrder(i), hDC
               Case olyBoxEmpty
                    OlyDrawEmptyBox ZOrder(i), hDC
               Case olySpot
                    OlyDrawSpot ZOrder(i), hDC
               Case olySpotEmpty
                    OlyDrawEmptySpot ZOrder(i), hDC
               Case olyStick
                    OlyDrawStick ZOrder(i), hDC
               Case olyTriangle
                    OlyDrawTriangle ZOrder(i), hDC
               Case olyTriangleEmpty
                    OlyDrawEmptyTriangle ZOrder(i), hDC
               Case olyTriStar
                    OlyDrawTriStar ZOrder(i), hDC
               End Select
          Case OlyUMC
               Select Case Oly(ZOrder(i)).Shape
               Case olyBox
                    OlyUMCDrawBox ZOrder(i), hDC
               Case olyBoxEmpty
                    OlyUMCDrawEmptyBox ZOrder(i), hDC
               Case olySpot
                    OlyUMCDrawSpot ZOrder(i), hDC
               Case olySpotEmpty
                    OlyUMCDrawEmptySpot ZOrder(i), hDC
               Case olyStick
                    OlyUMCDrawStick ZOrder(i), hDC
               Case olyTriangle
                    OlyUMCDrawTriangle ZOrder(i), hDC
               Case olyTriangleEmpty
                    OlyUMCDrawEmptyTriangle ZOrder(i), hDC
               Case olyTriStar
                    OlyUMCDrawTriStar ZOrder(i), hDC
               End Select
          End Select
          If Oly(ZOrder(i)).ShowText Then OlyDrawTextHorz ZOrder(i), hDC, TextHeight, TextWidth
       End If
     Next i
   Case OrientMWHrzETVrt
     For i = 0 To OlyCnt - 1
       If Oly(ZOrder(i)).Visible Then
          TextHeight = CLng(OlyOptions.DefFontHeight * (LoYE - LoY0) * Oly(ZOrder(i)).TextHeightPct)
          TextWidth = CLng(OlyOptions.DefFontWidth * (LoXE - LoX0) * Oly(ZOrder(i)).TextHeightPct)
          Select Case Oly(ZOrder(i)).Type
          Case olySolo
               Select Case Oly(ZOrder(i)).Shape
               Case olyBox
                    OlyDrawBoxNETVert ZOrder(i), hDC
               Case olyBoxEmpty
                    OlyDrawEmptyBoxNETVert ZOrder(i), hDC
               Case olySpot
                    OlyDrawSpotNETVert ZOrder(i), hDC
               Case olySpotEmpty
                    OlyDrawEmptySpotNETVert ZOrder(i), hDC
               Case olyStick
                    OlyDrawStickNETVert ZOrder(i), hDC
               Case olyTriangle
                    OlyDrawTriangleNETVert ZOrder(i), hDC
               Case olyTriangleEmpty
                    OlyDrawEmptyTriangleNETVert ZOrder(i), hDC
               Case olyTriStar
                    OlyDrawTriStarNETVert ZOrder(i), hDC
               End Select
          Case OlyUMC
               Select Case Oly(ZOrder(i)).Shape
               Case olyBox
                    OlyUMCDrawBoxNETVert ZOrder(i), hDC
               Case olyBoxEmpty
                    OlyUMCDrawEmptyBoxNETVert ZOrder(i), hDC
               Case olySpot
                    OlyUMCDrawSpotNETVert ZOrder(i), hDC
               Case olySpotEmpty
                    OlyUMCDrawEmptySpotNETVert ZOrder(i), hDC
               Case olyStick
                    OlyUMCDrawStickNETVert ZOrder(i), hDC
               Case olyTriangle
                    OlyUMCDrawTriangleNETVert ZOrder(i), hDC
               Case olyTriangleEmpty
                    OlyUMCDrawEmptyTriangleNETVert ZOrder(i), hDC
               Case olyTriStar
                    OlyUMCDrawTriStarNETVert ZOrder(i), hDC
               End Select
          End Select
          If Oly(ZOrder(i)).ShowText Then OlyDrawTextVert ZOrder(i), hDC, TextHeight, TextWidth
       End If
     Next i
   End Select
End If
End Function


Private Sub OverlayDrawVGrid(hDC As Long)
Dim i As Long
Dim rStep As Single
Dim lStep As Long, lx As Long
Dim ptPoint As POINTAPI
Dim OldPen As Long
Dim Res As Long
On Error Resume Next
OldPen = SelectObject(hDC, hOlyGridClrPen)
rStep = (MyCooSys.CurrRXMax - MyCooSys.CurrRXMin) / OlyOptions.GRID.VertBinsCount
lStep = (LoX2 - LoX1) \ OlyOptions.GRID.VertBinsCount
For i = 0 To OlyOptions.GRID.VertBinsCount
    lx = LoX1 + i * lStep
    Res = MoveToEx(hDC, lx, LoY0, ptPoint)
    Res = LineTo(hDC, lx, LoYE)
Next i
Res = SelectObject(hDC, OldPen)
End Sub

Private Sub OverlayDrawHGrid(hDC As Long)
Dim i As Long
Dim rStep As Single
Dim lStep As Long, ly As Long
Dim ptPoint As POINTAPI
Dim OldPen As Long
Dim Res As Long
On Error Resume Next
OldPen = SelectObject(hDC, hOlyGridClrPen)
rStep = (MyCooSys.CurrRYMax - MyCooSys.CurrRYMin) / OlyOptions.GRID.HorzBinsCount
lStep = (LoY2 - LoY1) \ OlyOptions.GRID.HorzBinsCount
For i = 0 To OlyOptions.GRID.HorzBinsCount
    ly = LoY1 + i * lStep
    Res = MoveToEx(hDC, LoX0, ly, ptPoint)
    Res = LineTo(hDC, LoXE, ly)
Next i
Res = SelectObject(hDC, OldPen)
End Sub

Private Sub OverlayDrawHNumbersRev(ByVal hDC As Long)
Dim i As Long, iSign As Long
Dim sNumber As String
Dim szNumber As Size
Dim rStep As Single
Dim lStep As Long, lx As Long
Dim NumT As Long
Dim MarkO As Long               'X coordinates of the thick mark (out)
Dim MarkL As Long               'X coordinates of the thick mark (on axis)
Dim ptPoint As POINTAPI
Dim vSelVec As Variant
Dim Res As Long
On Error Resume Next

If (MyCooSys.csOrigin < 3 And MyCooSys.csOriginXY < 3) Or _
   (MyCooSys.csOrigin > 2 And MyCooSys.csOriginXY > 2) Then
    MarkO = LoY0 - 30:       MarkL = LoY0
Else
    MarkO = LoYE + 30:       MarkL = LoYE
End If
If MyCooSys.csOrigin Mod 2 = 0 Then
    iSign = 1
Else
    iSign = -1
End If
rStep = (MyCooSys.CurrRYMax - MyCooSys.CurrRYMin) / 8
lStep = (LoX2 - LoX1) \ 8
Res = GetTextExtentPoint32(hDC, "0123456789", 10, szNumber)
vSelVec = OverlayGetJobVector(vAxYSelectMatrix, vYHNumJobMatrix)
NumT = SP(vSelVec, MarkL, 30, loSY \ 8, szNumber.cy)
For i = 0 To 8
    lx = LoX1 + i * lStep
    sNumber = Format$(MyCooSys.CurrRYMin + i * rStep, "0.0000")
    Res = GetTextExtentPoint32(hDC, sNumber, Len(sNumber), szNumber)
    Res = MoveToEx(hDC, lx, MarkO, ptPoint)
    Res = LineTo(hDC, lx, MarkL)
    Res = TextOut(hDC, lx + iSign * szNumber.cx \ 2, NumT, sNumber, Len(sNumber))
Next i
End Sub


Private Sub OverlayDrawVNumbersRev(ByVal hDC As Long)
Dim i As Long, iSign As Long
Dim sNumber As String
Dim szNumber As Size
Dim rStep As Single
Dim lStep As Long, ly As Long
Dim NumL As Long
Dim MarkO As Long               'X coordinates of the thick mark (out)
Dim MarkL As Long               'X coordinates of the thick mark (on axis)
Dim ptPoint As POINTAPI
Dim vSelVec As Variant
Dim Res As Long
On Error Resume Next

If (MyCooSys.csOrigin + MyCooSys.csOriginXY) Mod 2 = 0 Then
    MarkO = LoX0 - 30:       MarkL = LoX0
Else
    MarkO = LoXE + 30:       MarkL = LoXE
End If
If MyCooSys.csOrigin > 2 Then
    iSign = -1
Else
    iSign = 1
End If
rStep = (MyCooSys.CurrRXMax - MyCooSys.CurrRXMin) / 8
lStep = (LoY2 - LoY1) \ 8
For i = 0 To 8
    ly = LoY1 + i * lStep
    sNumber = Format$(MyCooSys.CurrRXMin + i * rStep, "#,###,##0.0000")
    Res = GetTextExtentPoint32(hDC, sNumber, Len(sNumber), szNumber)
    Res = MoveToEx(hDC, MarkO, ly, ptPoint)
    Res = LineTo(hDC, MarkL, ly)
    vSelVec = OverlayGetJobVector(vAxXSelectMatrix, vXVNumJobMatrix)
    NumL = SP(vSelVec, MarkL, 30, LoSX \ 4, szNumber.cx)
    Res = TextOut(hDC, NumL, ly + iSign * szNumber.cy \ 2, sNumber, Len(sNumber))
Next i
End Sub



'-------------------------------------M E T R I C   D R A W I N G-----------

Private Sub OverlayDrawMetric(hDC As Long)
Dim hClipRgn As Long
Dim tb As TextBoxGraph
On Error Resume Next
Me.MousePointer = vbHourglass
'metric printing is always landscape
mtrVX0 = GetDeviceCaps(hDC, HORZRES)      'this should be in device mode (preferably pixels)
mtrVY0 = GetDeviceCaps(hDC, VERTRES)

OverlayCooSys_MM hDC                      'includes painting of coordinate axes
If OlyOptions.GRID.VertGridVisible Then OverlayDrawVGrid_MM hDC
If OlyOptions.GRID.HorzGridVisible Then OverlayDrawHGrid_MM hDC
hClipRgn = ClippingRegionA_MM(hDC)
OverlayDraw_MM hDC
DeleteObject (hClipRgn)         '
Call SelectClipRgn(hDC, 0&)         'remove clipping regin from dc
'text boxes could be wherever
For Each tb In MyMetricPrint.TextBoxes
    tb.Draw hDC
Next
Me.MousePointer = vbDefault
End Sub


Public Sub OverlayCooSys_MM(ByVal hDC As Long)
Dim Res As Long
Dim ptPoint As POINTAPI
Dim OldPen As Long
On Error Resume Next
'set appropriate mapping mode; this mapping mode also moves coordinate origin
'in lower left corner(which is exactly what we want)
Res = SetMapMode(hDC, MM_HIMETRIC)
With MyMetricPrint
    Call SetViewportOrgEx(hDC, 0, mtrVY0, ptPoint)                        'viewport
    Call SetWindowOrgEx(hDC, -.OriginHorzL, -.OriginVertL, ptPoint)         'logical window
    
    OldPen = SelectObject(hDC, hOlyForeClrPen)
    Res = SetROP2(hDC, R2_COPYPEN)
    
    Select Case MyMetricPrint.Orientation
    Case OrientMWVrtETHrz
        Res = MoveToEx(hDC, 0, .MWRangeL, ptPoint)
        Res = LineTo(hDC, 0, 0)
        Res = LineTo(hDC, .ETRangeL, 0)
    Case OrientMWHrzETVrt
        Res = MoveToEx(hDC, .MWRangeL, 0, ptPoint)
        Res = LineTo(hDC, 0, 0)
        Res = LineTo(hDC, 0, .ETRangeL)
    End Select
    
    Res = SelectObject(hDC, OldPen)
End With
End Sub


Public Function CalcOverlayDataOne_MM(OlyInd) As Boolean
'------------------------------------------------------------------------
'recalculates coordinates of spots for Oly(OlyInd)
'------------------------------------------------------------------------
Dim CurrMW As Double, CurrNET As Double, CurrAbu As Double
Dim i As Long, TmpCnt As Long
Dim ChP() As LaV2DGPoint
Dim ChPNET() As Double
Dim ScaleET As Double
Dim ScaleMW As Double
On Error Resume Next      'there will be a whole bunch of overflow errors
'------------------------------------------------------------------------
'point is visible if within logical borders; unique mass class is visible
'if any of its characteristic points within logical borders
'------------------------------------------------------------------------
ScaleMW = MyMetricPrint.mmPerDa * HI_MM
ScaleET = MyMetricPrint.mmPerET * HI_MM
Select Case Oly(OlyInd).Type
Case olySolo
  Select Case Oly(OlyInd).Shape
  Case olyStick             'stick is unique since it does not care about size and has adjustment for NET
     TmpCnt = 0
     With Oly(OlyInd)
        For i = 1 To GelData(.DisplayInd).CSLines
            TmpCnt = TmpCnt + 1
            CurrMW = GelData(.DisplayInd).CSData(i).AverageMW
            CurrNET = .NETSlope * GelData(.DisplayInd).CSData(i).ScanNumber + .NETIntercept
            If MyMetricPrint.IsInScope(CurrNET, CurrMW) Then
               OlyCoo(OlyInd).OutOfScope(TmpCnt - 1) = False
               OlyCoo(OlyInd).x(TmpCnt - 1) = CSng((CurrNET - MyMetricPrint.MinET) * ScaleET)
               OlyCoo(OlyInd).y(TmpCnt - 1) = CSng((CurrMW - MyMetricPrint.MinMW) * ScaleMW)
               OlyCoo(OlyInd).XL(TmpCnt - 1) = Abs(CSng(OlyAdj(OlyInd).NETL(TmpCnt - 1) * ScaleET))
               OlyCoo(OlyInd).XU(TmpCnt - 1) = Abs(CSng(OlyAdj(OlyInd).NETU(TmpCnt - 1) * ScaleET))
               OlyCoo(OlyInd).YL(TmpCnt - 1) = OlyCoo(OlyInd).y(TmpCnt - 1)
               OlyCoo(OlyInd).YU(TmpCnt - 1) = OlyCoo(OlyInd).y(TmpCnt - 1)
            Else
               OlyCoo(OlyInd).OutOfScope(TmpCnt - 1) = True
            End If
        Next i
        For i = 1 To GelData(.DisplayInd).IsoLines
            TmpCnt = TmpCnt + 1
            CurrMW = GetIsoMass(GelData(.DisplayInd).IsoData(i), GelData(.DisplayInd).Preferences.IsoDataField)
            CurrNET = .NETSlope * GelData(.DisplayInd).IsoData(i).ScanNumber + .NETIntercept
            If MyMetricPrint.IsInScope(CurrNET, CurrMW) Then
               OlyCoo(OlyInd).OutOfScope(TmpCnt - 1) = False
               OlyCoo(OlyInd).x(TmpCnt - 1) = CSng((CurrNET - MyMetricPrint.MinET) * ScaleET)
               OlyCoo(OlyInd).y(TmpCnt - 1) = CSng((CurrMW - MyMetricPrint.MinMW) * ScaleMW)
               OlyCoo(OlyInd).XL(TmpCnt - 1) = Abs(CSng(OlyAdj(OlyInd).NETL(TmpCnt - 1) * ScaleET))
               OlyCoo(OlyInd).XU(TmpCnt - 1) = Abs(CSng(OlyAdj(OlyInd).NETU(TmpCnt - 1) * ScaleET))
               OlyCoo(OlyInd).YL(TmpCnt - 1) = OlyCoo(OlyInd).y(TmpCnt - 1)
               OlyCoo(OlyInd).YU(TmpCnt - 1) = OlyCoo(OlyInd).y(TmpCnt - 1)
            Else
               OlyCoo(OlyInd).OutOfScope(TmpCnt - 1) = True
            End If
        Next i
     End With
  Case olyBox, olyBoxEmpty, olySpot, olySpotEmpty, olyTriangle, olyTriangleEmpty, olyTriStar
     TmpCnt = 0
     With Oly(OlyInd)
        If MyMetricPrint.ScaleSizeType = stNone Then
           For i = 1 To GelData(.DisplayInd).CSLines
               TmpCnt = TmpCnt + 1
               CurrMW = GelData(.DisplayInd).CSData(i).AverageMW
               CurrNET = .NETSlope * GelData(.DisplayInd).CSData(i).ScanNumber + .NETIntercept
               CurrAbu = GelData(.DisplayInd).CSData(i).Abundance
               If MyMetricPrint.IsInScope(CurrNET, CurrMW) Then
                  OlyCoo(OlyInd).OutOfScope(TmpCnt - 1) = False
                  OlyCoo(OlyInd).x(TmpCnt - 1) = CSng((CurrNET - MyMetricPrint.MinET) * ScaleET)
                  OlyCoo(OlyInd).y(TmpCnt - 1) = CSng((CurrMW - MyMetricPrint.MinMW) * ScaleMW)
                  OlyCoo(OlyInd).R(TmpCnt - 1) = MyMetricPrint.SpotHeightL
               Else
                  OlyCoo(OlyInd).OutOfScope(TmpCnt - 1) = True
               End If
           Next i
           For i = 1 To GelData(.DisplayInd).IsoLines
               TmpCnt = TmpCnt + 1
               CurrMW = GetIsoMass(GelData(.DisplayInd).IsoData(i), GelData(.DisplayInd).Preferences.IsoDataField)
               CurrNET = .NETSlope * GelData(.DisplayInd).IsoData(i).ScanNumber + .NETIntercept
               CurrAbu = GelData(.DisplayInd).IsoData(i).Abundance
               If MyMetricPrint.IsInScope(CurrNET, CurrMW) Then
                  OlyCoo(OlyInd).OutOfScope(TmpCnt - 1) = False
                  OlyCoo(OlyInd).x(TmpCnt - 1) = CSng((CurrNET - MyMetricPrint.MinET) * ScaleET)
                  OlyCoo(OlyInd).y(TmpCnt - 1) = CSng((CurrMW - MyMetricPrint.MinMW) * ScaleMW)
                  OlyCoo(OlyInd).R(TmpCnt - 1) = MyMetricPrint.SpotHeightL
               Else
                  OlyCoo(OlyInd).OutOfScope(TmpCnt - 1) = True
               End If
           Next i
        Else
           For i = 1 To GelData(.DisplayInd).CSLines
               TmpCnt = TmpCnt + 1
               CurrMW = GelData(.DisplayInd).CSData(i).AverageMW
               CurrNET = .NETSlope * GelData(.DisplayInd).CSData(i).ScanNumber + .NETIntercept
               CurrAbu = GelData(.DisplayInd).CSData(i).Abundance
               If MyMetricPrint.IsInScope(CurrNET, CurrMW) Then
                  OlyCoo(OlyInd).OutOfScope(TmpCnt - 1) = False
                  OlyCoo(OlyInd).x(TmpCnt - 1) = CSng((CurrNET - MyMetricPrint.MinET) * ScaleET)
                  OlyCoo(OlyInd).y(TmpCnt - 1) = CSng((CurrMW - MyMetricPrint.MinMW) * ScaleMW)
                  OlyCoo(OlyInd).R(TmpCnt - 1) = MyMetricPrint.GetSizeL(CurrAbu)
               Else
                  OlyCoo(OlyInd).OutOfScope(TmpCnt - 1) = True
               End If
           Next i
           For i = 1 To GelData(.DisplayInd).IsoLines
               TmpCnt = TmpCnt + 1
               CurrMW = GetIsoMass(GelData(.DisplayInd).IsoData(i), GelData(.DisplayInd).Preferences.IsoDataField)
               CurrNET = .NETSlope * GelData(.DisplayInd).IsoData(i).ScanNumber + .NETIntercept
               CurrAbu = GelData(.DisplayInd).IsoData(i).Abundance
               If MyMetricPrint.IsInScope(CurrNET, CurrMW) Then
                  OlyCoo(OlyInd).OutOfScope(TmpCnt - 1) = False
                  OlyCoo(OlyInd).x(TmpCnt - 1) = CSng((CurrNET - MyMetricPrint.MinET) * ScaleET)
                  OlyCoo(OlyInd).y(TmpCnt - 1) = CSng((CurrMW - MyMetricPrint.MinMW) * ScaleMW)
                  OlyCoo(OlyInd).R(TmpCnt - 1) = MyMetricPrint.GetSizeL(CurrAbu)
               Else
                  OlyCoo(OlyInd).OutOfScope(TmpCnt - 1) = True
               End If
           Next i
        End If
     End With
   End Select
Case OlyUMC
   With Oly(OlyInd)
     If MyMetricPrint.ScaleSizeType = stNone Then
        For i = 0 To GelUMC(.DisplayInd).UMCCnt - 1
            If fUMCCharacteristicPoints(.DisplayInd, i, ChP()) Then
               ReDim ChPNET(2)
               ChPNET(0) = .NETSlope * ChP(0).Scan + .NETIntercept
               OlyCoo(OlyInd).XL(i) = CSng((ChPNET(0) - MyMetricPrint.MinET) * ScaleET)
               ChPNET(2) = .NETSlope * ChP(2).Scan + .NETIntercept
               OlyCoo(OlyInd).XU(i) = CSng((ChPNET(2) - MyMetricPrint.MinET) * ScaleET)
               ChPNET(1) = .NETSlope * ChP(1).Scan + .NETIntercept
               OlyCoo(OlyInd).x(i) = CSng((ChPNET(1) - MyMetricPrint.MinET) * ScaleET)
               'end points draw as class mass and class representative size higher
               OlyCoo(OlyInd).YL(i) = CSng((ChP(1).MW - MyMetricPrint.MinMW) * ScaleMW)
               OlyCoo(OlyInd).YU(i) = OlyCoo(OlyInd).YL(i)
               OlyCoo(OlyInd).y(i) = OlyCoo(OlyInd).YL(i) + MyMetricPrint.SpotHeightL
               'UMC should be drawn if any of its points is visible(optimized for speed)
               OlyCoo(OlyInd).OutOfScope(i) = True
               If MyMetricPrint.IsInScope(ChPNET(1), ChP(1).MW) Then
                  OlyCoo(OlyInd).OutOfScope(i) = False
               Else
                  If MyMetricPrint.IsInScope(ChPNET(0), ChP(0).MW) Then
                     OlyCoo(OlyInd).OutOfScope(i) = False
                  Else
                     If MyMetricPrint.IsInScope(ChPNET(2), ChP(2).MW) Then OlyCoo(OlyInd).OutOfScope(i) = False
                  End If
               End If
            Else
               OlyCoo(OlyInd).R(i) = -1
            End If
        Next i
     Else
        For i = 0 To GelUMC(.DisplayInd).UMCCnt - 1
            If fUMCCharacteristicPoints(.DisplayInd, i, ChP()) Then
               ReDim ChPNET(2)
               ChPNET(0) = .NETSlope * ChP(0).Scan + .NETIntercept
               OlyCoo(OlyInd).XL(i) = CSng((ChPNET(0) - MyMetricPrint.MinET) * ScaleET)
               ChPNET(2) = .NETSlope * ChP(2).Scan + .NETIntercept
               OlyCoo(OlyInd).XU(i) = CSng((ChPNET(2) - MyMetricPrint.MinET) * ScaleET)
               ChPNET(1) = .NETSlope * ChP(1).Scan + .NETIntercept
               OlyCoo(OlyInd).x(i) = CSng((ChPNET(1) - MyMetricPrint.MinET) * ScaleET)
               'end points draw as class mass and class representative size higher
               OlyCoo(OlyInd).YL(i) = CSng((ChP(1).MW - MyMetricPrint.MinMW) * ScaleMW)
               OlyCoo(OlyInd).YU(i) = OlyCoo(OlyInd).YL(i)
               OlyCoo(OlyInd).R(i) = MyMetricPrint.GetSizeL(ChP(0).Abu)
               OlyCoo(OlyInd).y(i) = OlyCoo(OlyInd).YL(i) + OlyCoo(OlyInd).R(i)
               'UMC should be drawn if any of its points is visible(optimized for speed)
               OlyCoo(OlyInd).OutOfScope(i) = True
               If MyMetricPrint.IsInScope(ChPNET(1), ChP(1).MW) Then
                  OlyCoo(OlyInd).OutOfScope(i) = False
               Else
                  If MyMetricPrint.IsInScope(ChPNET(0), ChP(0).MW) Then
                     OlyCoo(OlyInd).OutOfScope(i) = False
                  Else
                     If MyMetricPrint.IsInScope(ChPNET(2), ChP(2).MW) Then OlyCoo(OlyInd).OutOfScope(i) = False
                  End If
               End If
            Else
               OlyCoo(OlyInd).R(i) = -1
            End If
        Next i
     End If
   End With
End Select
End Function


Public Sub OverlayDrawVGrid_MM(ByVal hDC As Long)
'--------------------------------------------------------------------------
'grid is obtained by increasing for .Grid(MW/ET) from 0, and marking labels
'--------------------------------------------------------------------------
Dim Ind1 As Long, Ind2 As Long
Dim i As Long
Dim CurrVal As Double
Dim CurrValL As Long            'position in logical units
Dim CurrLbl As String
Dim CurrFmt As String
Dim OldPen As Long
Dim Res As Long
Dim pt As POINTAPI
Dim sz As Size
Dim lNewFont As Long, lOldFont As Long
Dim lfFnt As LOGFONT
OldPen = SelectObject(hDC, hOlyGridClrPen)
Res = SetROP2(hDC, R2_COPYPEN)
With MyMetricPrint
     'create new logical font based on metric specification and load it to the device context
     lOldFont = SelectObject(hDC, GetStockObject(SYSTEM_FONT))      'get selected font by tmp. selecting
                                                                    'SYSTEM_FONT in device context
     Res = GetObjectAPI(lOldFont, Len(lfFnt), lfFnt)                'get font structure
     Res = SelectObject(hDC, lOldFont)                              'select old font back to dev.context
     
     lfFnt.lfWidth = .FontLabelWidthL                               'use obtained font structure to
     lfFnt.lfHeight = .FontLabelHeightL                             'indirectly create new font
     lNewFont = CreateFontIndirect(lfFnt)                           'and select it to the device context
     lOldFont = SelectObject(hDC, lNewFont)
     
     Select Case .Orientation
     Case OrientMWHrzETVrt                      'vertical grid is based on MW
          Ind1 = CLng(.MinMW / .GridMW_R) - 1
          Ind2 = CLng(.MaxMW / .GridMW_R) + 1
          CurrFmt = GetNumFormat(.MWDecPlaces)
          For i = Ind1 To Ind2
              CurrVal = i * .GridMW_R
              If CurrVal >= (.MinMW - SMALL_VALUE) And CurrVal <= (.MaxMW + SMALL_VALUE) Then
                 'draw line with a bit of indication of a thick mark
                 CurrValL = CurrVal * .mmPerDa * HI_MM - .MinMWL
                 Res = MoveToEx(hDC, CurrValL, .ETRangeL, pt)
                 Res = LineTo(hDC, CurrValL, -.VertOffsetL / 4)     'just indication of thick mark
                 'draw label under the thick mark
                 CurrLbl = Format$(CurrVal, CurrFmt)
                 Res = GetTextExtentPoint32(hDC, CurrLbl, Len(CurrLbl), sz)
                 Res = TextOut(hDC, CurrValL - CLng(sz.cx / 2), -.VertOffsetL, CurrLbl, Len(CurrLbl))
              End If
          Next i
          CurrLbl = "NET"                  'label on the vertical axis is NET
          Res = GetTextExtentPoint32(hDC, CurrLbl, Len(CurrLbl), sz)
          Res = TextOut(hDC, .HorzOffsetL, .ETRangeL + .HorzOffsetL + CLng(sz.cy / 2), CurrLbl, Len(CurrLbl))
     Case OrientMWVrtETHrz                      'vertical grid is based on ET
          Ind1 = CLng(.MinET / .GridET_R) - 1
          Ind2 = CLng(.MaxET / .GridET_R) + 1
          CurrFmt = GetNumFormat(.ETDecPlaces)
          For i = Ind1 To Ind2
              CurrVal = i * .GridET_R
              If CurrVal >= (.MinET - SMALL_VALUE) And CurrVal <= (.MaxET + SMALL_VALUE) Then
                 'draw line
                 CurrValL = CurrVal * .mmPerET * HI_MM - .MinETL
                 Res = MoveToEx(hDC, CurrValL, .MWRangeL, pt)
                 Res = LineTo(hDC, CurrValL, -CLng(.VertOffsetL / 4))     'just indication of thick mark
                 'draw label under the thick mark
                 CurrLbl = Format$(CurrVal, CurrFmt)
                 Res = GetTextExtentPoint32(hDC, CurrLbl, Len(CurrLbl), sz)
                 Res = TextOut(hDC, CurrValL - CLng(sz.cx / 2), -.VertOffsetL, CurrLbl, Len(CurrLbl))
              End If
          Next i
          CurrLbl = "MW"                  'label on the vertical axis is MW
          Res = GetTextExtentPoint32(hDC, CurrLbl, Len(CurrLbl), sz)
          Res = TextOut(hDC, .HorzOffsetL, .MWRangeL + .HorzOffsetL + CLng(sz.cy / 2), CurrLbl, Len(CurrLbl))
     End Select
End With
Res = SelectObject(hDC, lOldFont)
Res = SelectObject(hDC, OldPen)
Res = DeleteObject(lNewFont)
End Sub


Public Sub OverlayDrawHGrid_MM(ByVal hDC As Long)
'--------------------------------------------------------------------------
'grid is obtained by increasing for .Grid(MW/ET) from 0, and marking labels
'--------------------------------------------------------------------------
Dim Ind1 As Long, Ind2 As Long
Dim i As Long
Dim CurrVal As Double
Dim CurrValL As Long            'position in logical units
Dim CurrLbl As String
Dim CurrFmt As String
Dim OldPen As Long
Dim Res As Long
Dim pt As POINTAPI
Dim sz As Size
Dim lNewFont As Long, lOldFont As Long
Dim lfFnt As LOGFONT
OldPen = SelectObject(hDC, hOlyGridClrPen)
Res = SetROP2(hDC, R2_COPYPEN)
With MyMetricPrint
     'create new logical font based on metric specification and load it to the device context
     lOldFont = SelectObject(hDC, GetStockObject(SYSTEM_FONT))      'get selected font by tmp. selecting
                                                                    'SYSTEM_FONT in device context
     Res = GetObjectAPI(lOldFont, Len(lfFnt), lfFnt)                'get font structure
     Res = SelectObject(hDC, lOldFont)                              'select old font back to dev.context
     
     lfFnt.lfWidth = .FontLabelWidthL                               'use obtained font structure to
     lfFnt.lfHeight = .FontLabelHeightL                             'indirectly create new font
     lNewFont = CreateFontIndirect(lfFnt)                           'and select it to the device context
     lOldFont = SelectObject(hDC, lNewFont)
     
     Select Case .Orientation
     Case OrientMWHrzETVrt                   'horizontal grid is based on ET
          Ind1 = CLng(.MinET / .GridET_R) - 1
          Ind2 = CLng(.MaxET / .GridET_R) + 1
          CurrFmt = GetNumFormat(.ETDecPlaces)
          For i = Ind1 To Ind2
              CurrVal = i * .GridET_R
              If CurrVal >= (.MinET - SMALL_VALUE) And CurrVal <= (.MaxET + SMALL_VALUE) Then
                 'draw line with a bit of indication of a thick mark
                 CurrValL = CurrVal * .mmPerET * HI_MM - .MinETL
                 Res = MoveToEx(hDC, .MWRangeL, CurrValL, pt)
                 Res = LineTo(hDC, -.HorzOffsetL / 4, CurrValL)     'just indication of thick mark
                 'draw label left of the thick mark
                 CurrLbl = Format$(CurrVal, CurrFmt)
                 Res = GetTextExtentPoint32(hDC, CurrLbl, Len(CurrLbl), sz)
                 Res = TextOut(hDC, -(.HorzOffsetL + sz.cx), CurrValL + CLng(sz.cy / 2), CurrLbl, Len(CurrLbl))
              End If
          Next i
          CurrLbl = "MW"                  'label on the horizontal axis is MW
          Res = GetTextExtentPoint32(hDC, CurrLbl, Len(CurrLbl), sz)
          Res = TextOut(hDC, .MWRangeL + .HorzOffsetL - sz.cx, -(2 * .VertOffsetL + sz.cy), CurrLbl, Len(CurrLbl))
     Case OrientMWVrtETHrz                   'horizontal grid is based on MW
          Ind1 = CLng(.MinMW / .GridMW_R) - 1
          Ind2 = CLng(.MaxMW / .GridMW_R) + 1
          CurrFmt = GetNumFormat(.MWDecPlaces)
          For i = Ind1 To Ind2
              CurrVal = i * .GridMW_R
              If CurrVal >= (.MinMW - SMALL_VALUE) And CurrVal <= (.MaxMW + SMALL_VALUE) Then
                 'draw line
                 CurrValL = CurrVal * .mmPerDa * HI_MM - .MinMWL
                 Res = MoveToEx(hDC, .ETRangeL, CurrValL, pt)
                 Res = LineTo(hDC, -.HorzOffsetL / 4, CurrValL)     'just indication of thick mark
                 'draw label under the thick mark
                 CurrLbl = Format$(CurrVal, CurrFmt)
                 Res = GetTextExtentPoint32(hDC, CurrLbl, Len(CurrLbl), sz)
                 Res = TextOut(hDC, -(.HorzOffsetL + sz.cx), CurrValL + CLng(sz.cy / 2), CurrLbl, Len(CurrLbl))
              End If
          Next i
          CurrLbl = "NET"                  'label on the horizontal axis is MW
          Res = GetTextExtentPoint32(hDC, CurrLbl, Len(CurrLbl), sz)
          Res = TextOut(hDC, .ETRangeL + .HorzOffsetL - sz.cx, -(2 * .VertOffsetL + sz.cy), CurrLbl, Len(CurrLbl))
     End Select
End With
Res = SelectObject(hDC, lOldFont)
Res = SelectObject(hDC, OldPen)
Res = DeleteObject(lNewFont)
End Sub


Private Function OverlayDraw_MM(ByVal hDC As Long) As Boolean
'---------------------------------------------------------------------------------------
'handles draw of individual overlays; since metric drawing uses the same structures as
'screen drawing we need to recalculate all overlays and after drawing restore old values
'---------------------------------------------------------------------------------------
Dim i As Long
Dim TmpCoo As OverlayCoo
Dim ZOrder() As Long
ScreenUpdateEnabled = False                        'disable screen update during this operation
If GetOlyZOrder(ZOrder()) Then
   Select Case MyMetricPrint.Orientation
   Case OrientMWVrtETHrz
     For i = 0 To OlyCnt - 1
       If Oly(ZOrder(i)).Visible Then
          TmpCoo = OlyCoo(ZOrder(i))               'save old overlay numbers
          Call CalcOverlayDataOne_MM(ZOrder(i))
          Select Case Oly(ZOrder(i)).Type
          Case olySolo
               Select Case Oly(ZOrder(i)).Shape
               Case olyBox
                    OlyDrawBox ZOrder(i), hDC
               Case olyBoxEmpty
                    OlyDrawEmptyBox ZOrder(i), hDC
               Case olySpot
                    OlyDrawSpot ZOrder(i), hDC
               Case olySpotEmpty
                    OlyDrawEmptySpot ZOrder(i), hDC
               Case olyStick
                    OlyDrawStickNETHorz_MM ZOrder(i), hDC, CLng(MyMetricPrint.StickWidthL / 2)
               Case olyTriangle
                    OlyDrawTriangle ZOrder(i), hDC
               Case olyTriangleEmpty
                    OlyDrawEmptyTriangle ZOrder(i), hDC
               Case olyTriStar
                    OlyDrawTriStar ZOrder(i), hDC
               End Select
          Case OlyUMC
               Select Case Oly(ZOrder(i)).Shape
               Case olyBox
                    OlyUMCDrawBox ZOrder(i), hDC
               Case olyBoxEmpty
                    OlyUMCDrawEmptyBox ZOrder(i), hDC
               Case olySpot
                    OlyUMCDrawSpot ZOrder(i), hDC
               Case olySpotEmpty
                    OlyUMCDrawEmptySpot ZOrder(i), hDC
               Case olyStick
                    OlyUMCDrawStickHorz_MM ZOrder(i), hDC, CLng(MyMetricPrint.StickWidthL / 2)
               Case olyTriangle
                    OlyUMCDrawTriangle ZOrder(i), hDC
               Case olyTriangleEmpty
                    OlyUMCDrawEmptyTriangle ZOrder(i), hDC
               Case olyTriStar
                    OlyUMCDrawTriStar ZOrder(i), hDC
               End Select
          End Select
          If Oly(ZOrder(i)).ShowText Then OlyDrawTextH_MM ZOrder(i), hDC, MyMetricPrint.FontHeightL, _
                                                MyMetricPrint.FontWidthL, MyMetricPrint.NameOffsetL
          OlyCoo(ZOrder(i)) = TmpCoo            'restore old OverlayCoo structure
       End If
     Next i
   Case OrientMWHrzETVrt
     For i = 0 To OlyCnt - 1
       If Oly(ZOrder(i)).Visible Then
          TmpCoo = OlyCoo(ZOrder(i))                            'save old overlay coo data
          Call CalcOverlayDataOne_MM(ZOrder(i))
          Select Case Oly(ZOrder(i)).Type
          Case olySolo
               Select Case Oly(ZOrder(i)).Shape
               Case olyBox
                    OlyDrawBoxNETVert ZOrder(i), hDC
               Case olyBoxEmpty
                    OlyDrawEmptyBoxNETVert ZOrder(i), hDC
               Case olySpot
                    OlyDrawSpotNETVert ZOrder(i), hDC
               Case olySpotEmpty
                    OlyDrawEmptySpotNETVert ZOrder(i), hDC
               Case olyStick
                    OlyDrawStickNETVert_MM ZOrder(i), hDC, CLng(MyMetricPrint.StickWidthL / 2)
               Case olyTriangle
                    OlyDrawTriangleNETVert ZOrder(i), hDC
               Case olyTriangleEmpty
                    OlyDrawEmptyTriangleNETVert ZOrder(i), hDC
               Case olyTriStar
                    OlyDrawTriStarNETVert ZOrder(i), hDC
               End Select
          Case OlyUMC
               Select Case Oly(ZOrder(i)).Shape
               Case olyBox
                    OlyUMCDrawBoxNETVert ZOrder(i), hDC
               Case olyBoxEmpty
                    OlyUMCDrawEmptyBoxNETVert ZOrder(i), hDC
               Case olySpot
                    OlyUMCDrawSpotNETVert ZOrder(i), hDC
               Case olySpotEmpty
                    OlyUMCDrawEmptySpotNETVert ZOrder(i), hDC
               Case olyStick
                    OlyUMCDrawStickVert_MM ZOrder(i), hDC, CLng(MyMetricPrint.StickWidthL / 2)
               Case olyTriangle
                    OlyUMCDrawTriangleNETVert ZOrder(i), hDC
               Case olyTriangleEmpty
                    OlyUMCDrawEmptyTriangleNETVert ZOrder(i), hDC
               Case olyTriStar
                    OlyUMCDrawTriStarNETVert ZOrder(i), hDC
               End Select
          End Select
          If Oly(ZOrder(i)).ShowText Then OlyDrawTextV_MM ZOrder(i), hDC, MyMetricPrint.FontHeightL, _
                                                MyMetricPrint.FontWidthL, MyMetricPrint.NameOffsetL
          OlyCoo(ZOrder(i)) = TmpCoo            'restore old OverlayCoo structure
       End If
     Next i
   End Select
End If
ScreenUpdateEnabled = True
End Function


Public Function ClippingRegionA_MM(hDC As Long) As Long
'----------------------------------------------------------------------------------
'returns clipping region for metric drawing
'----------------------------------------------------------------------------------
Dim Res As Long
Dim hCR As Long
Dim ptClipRect() As POINTAPI
On Error Resume Next
ReDim ptClipRect(1)
With MyMetricPrint
    Select Case .Orientation
    Case OrientMWHrzETVrt
         ptClipRect(0).x = 0:            ptClipRect(0).y = 0
         ptClipRect(1).x = .MWRangeL:    ptClipRect(1).y = .ETRangeL
    Case OrientMWVrtETHrz
         ptClipRect(0).x = 0:            ptClipRect(0).y = 0
         ptClipRect(1).x = .ETRangeL:    ptClipRect(1).y = .MWRangeL
    End Select
End With
Res = LPtoDP(hDC, ptClipRect(0), 2)
hCR = CreateRectRgn(ptClipRect(0).x, ptClipRect(0).y, ptClipRect(1).x, ptClipRect(1).y)
Res = SelectClipRgn(hDC, hCR)
ClippingRegionA_MM = hCR
End Function


