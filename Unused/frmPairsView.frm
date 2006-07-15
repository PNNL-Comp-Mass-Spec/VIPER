VERSION 5.00
Begin VB.Form frmPairsView 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pairs Viewer"
   ClientHeight    =   4080
   ClientLeft      =   2760
   ClientTop       =   3705
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   Begin VIPER.LaSpotsRWC LaHeavy 
      Height          =   1800
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3175
   End
   Begin VIPER.LaSpotsRWC LaLight 
      Height          =   1800
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Light"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Heavy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmPairsView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'function implements pairs viewer for 2D displays
'NOTE: this form stays always on top while visible
'-------------------------------------------------------------------------
'created: 04/02/2003 nt
'last modified: 04/03/2003 nt
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

LaHeavy.MyCooSys.BuildingCS = True
LaHeavy.OwnerInd = CallerID
LaHeavy.IsoColor = glCSColor
LaHeavy.IsoColor = glIsoColor
LaHeavy.BackColor = &HFF0C0C
LaHeavy.MyCooSys.BuildingCS = False

LaHeavy.MyCooSys.BuildingCS = True
LaLight.OwnerInd = CallerID
LaLight.CSColor = glCSColor
LaLight.IsoColor = glIsoColor
LaLight.BackColor = &HC0CFF
LaHeavy.MyCooSys.BuildingCS = False

'this form should always stay on top
lTop = GetSystemMetrics(SM_CYMENU) \ 2
lLeft = GetSystemMetrics(SM_CXSCREEN) \ 2
lFlags = SWP_NOSIZE
Res = SetWindowPos(Me.hwnd, HWND_TOPMOST, lLeft, lTop, 0, 0, lFlags)
End Sub


Public Function Zoom_Light(FirstScan As Long, LastScan As Long, _
                           MinMW As Double, MaxMW As Double) As Boolean
'----------------------------------------------------------------------
'zooms light member of pair to specified range
'----------------------------------------------------------------------
Dim CSCnt As Long, CSInd() As Long
Dim IsoCnt As Long, IsoInd() As Long
Dim UMCCnt As Long, UMCInd() As Long
Dim IsoMWField As Long
On Error Resume Next
If FirstScan <= LastScan Then
   If MinMW <= MaxMW Then
      IsoMWField = GelData(CallerID).Preferences.IsoDataField
      CSCnt = GetWindowCS(CallerID, CSInd(), FirstScan, LastScan, MinMW, MaxMW)
      If CSCnt > 0 Then LaLight.AddSpotsCS CSInd()
      IsoCnt = GetWindowIso(CallerID, IsoMWField, IsoInd, _
                          FirstScan, LastScan, MinMW, MaxMW)
      If IsoCnt > 0 Then LaLight.AddSpotsIso IsoInd()
      UMCCnt = GetUMCList(CallerID, FirstScan, LastScan, MinMW, MaxMW, UMCInd())
      If UMCCnt > 0 Then LaLight.AddBoxes UMCInd()
      LaLight.SetRangeLimits FirstScan, LastScan, MinMW, MaxMW
      LaLight.InitCoordinateSystem
   End If
End If
End Function


Public Function Zoom_Heavy(FirstScan As Long, LastScan As Long, _
                           MinMW As Double, MaxMW As Double) As Boolean
'----------------------------------------------------------------------
'zooms heavy member of pair to specified range
'----------------------------------------------------------------------
Dim CSCnt As Long, CSInd() As Long
Dim IsoCnt As Long, IsoInd() As Long
Dim UMCCnt As Long, UMCInd() As Long
Dim IsoMWField As Long
On Error Resume Next
If FirstScan <= LastScan Then
   If MinMW <= MaxMW Then
      IsoMWField = GelData(CallerID).Preferences.IsoDataField
      CSCnt = GetWindowCS(CallerID, CSInd(), FirstScan, LastScan, MinMW, MaxMW)
      If CSCnt > 0 Then LaHeavy.AddSpotsCS CSInd()
      IsoCnt = GetWindowIso(CallerID, IsoMWField, IsoInd, _
                          FirstScan, LastScan, MinMW, MaxMW)
      If IsoCnt > 0 Then LaHeavy.AddSpotsIso IsoInd()
      UMCCnt = GetUMCList(CallerID, FirstScan, LastScan, MinMW, MaxMW, UMCInd())
      If UMCCnt > 0 Then LaHeavy.AddBoxes UMCInd()
      LaHeavy.SetRangeLimits FirstScan, LastScan, MinMW, MaxMW
      LaHeavy.InitCoordinateSystem
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


Private Sub LaHeavy_MenuClosed()
LaHeavy.DrawRefresh
RaiseEvent pvControlDone
End Sub


Private Sub LaLight_MenuClosed()
LaLight.DrawRefresh
RaiseEvent pvControlDone
End Sub
