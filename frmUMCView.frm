VERSION 5.00
Begin VB.Form frmUMCView 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Unique Mass Classes Viewer"
   ClientHeight    =   1845
   ClientLeft      =   2775
   ClientTop       =   3720
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   Begin VIPER.LaSpotsRWC LaThisUMC 
      Height          =   1800
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3175
   End
End
Attribute VB_Name = "frmUMCView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'function implements Unique Mass Classes viewer for 2D displays
'NOTE: this form stays always on top while visible
'-------------------------------------------------------------------------
'created: 04/04/2003 nt
'last modified: 04/04/2003 nt
'-------------------------------------------------------------------------
Option Explicit

Public CallerID As Long             'index opf calling display

Public Event pvControlDone()        'need these two events to signal
Public Event pvUnload()             'calling form what's going on
'pvControlDone event occurs whenever control signals it is done with
'whatever it was busy with------------------------------------------


Private Sub Form_Load()
Dim Res As Long
Dim lTop As Long
Dim lLeft As Long
Dim lFlags As Long
On Error Resume Next

LaThisUMC.MyCooSys.BuildingCS = True
LaThisUMC.OwnerInd = CallerID
LaThisUMC.ForeColor = glForeColor
LaThisUMC.BackColor = glBackColor
LaThisUMC.CSColor = glCSColor
LaThisUMC.IsoColor = glIsoColor
LaThisUMC.MyCooSys.csIdealXYAspRat = 7.5
LaThisUMC.MyCooSys.BuildingCS = False

'this form should always stay on top
lTop = GetSystemMetrics(SM_CYMENU) \ 2
lLeft = GetSystemMetrics(SM_CXSCREEN) \ 2
lFlags = SWP_NOSIZE
Res = SetWindowPos(Me.hwnd, HWND_TOPMOST, lLeft, lTop, 0, 0, lFlags)
End Sub


Public Function Zoom_UMC(FirstScan As Long, LastScan As Long, _
                         MinMW As Double, MaxMW As Double) As Boolean
'----------------------------------------------------------------------
'zooms light member of pair to specified range
'----------------------------------------------------------------------
Dim CSCnt As Long, CSInd() As Long
Dim IsoCnt As Long, IsoInd() As Long
Dim UMCCnt As Long, UMCInd() As Long
Dim IsoMWField As Integer
On Error Resume Next
If FirstScan <= LastScan Then
   If MinMW <= MaxMW Then
      IsoMWField = GelData(CallerID).Preferences.IsoDataField
      CSCnt = GetWindowCS(CallerID, CSInd(), FirstScan, LastScan, MinMW, MaxMW)
      If CSCnt > 0 Then LaThisUMC.AddSpotsCS CSInd()
      IsoCnt = GetWindowIso(CallerID, IsoMWField, IsoInd, _
                          FirstScan, LastScan, MinMW, MaxMW)
      If IsoCnt > 0 Then LaThisUMC.AddSpotsIso IsoInd()
      UMCCnt = GetUMCList(CallerID, FirstScan, LastScan, MinMW, MaxMW, UMCInd())
      If UMCCnt > 0 Then LaThisUMC.AddBoxes UMCInd()
      LaThisUMC.SetRangeLimits FirstScan, LastScan, MinMW, MaxMW
      LaThisUMC.InitCoordinateSystem
   End If
End If
End Function


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then         'hide but don't unload
   Cancel = True
   Me.Hide
   RaiseEvent pvUnload
End If
End Sub


Private Sub Form_Resize()
If Me.WindowState <> vbMinimized Then
   LaThisUMC.width = Me.ScaleWidth
   LaThisUMC.Height = Me.ScaleHeight
End If
End Sub

Private Sub LaThisUMC_MenuClosed()
LaThisUMC.DrawRefresh
RaiseEvent pvControlDone
End Sub
