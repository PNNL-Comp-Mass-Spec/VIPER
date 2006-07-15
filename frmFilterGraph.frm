VERSION 5.00
Begin VB.Form frmFilterGraph 
   Caption         =   "Filter Data - Graph"
   ClientHeight    =   7065
   ClientLeft      =   2775
   ClientTop       =   4050
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   120
      ScaleHeight     =   2265
      ScaleWidth      =   1305
      TabIndex        =   1
      Top             =   240
      Width           =   1335
      Begin VB.CommandButton cmdClose 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   17
         Top             =   30
         Width           =   255
      End
      Begin VB.TextBox txtYRes 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   600
         TabIndex        =   14
         Text            =   "128"
         Top             =   1920
         Width           =   600
      End
      Begin VB.TextBox txtXRes 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   600
         TabIndex        =   12
         Text            =   "128"
         Top             =   1680
         Width           =   600
      End
      Begin VB.TextBox txtClrRes 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   600
         TabIndex        =   10
         Text            =   "256"
         Top             =   1440
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Options"
         Height          =   255
         Index           =   7
         Left            =   30
         TabIndex        =   16
         Top             =   30
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y Res."
         Height          =   195
         Index           =   6
         Left            =   30
         TabIndex        =   15
         Top             =   1920
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X Res."
         Height          =   195
         Index           =   5
         Left            =   30
         TabIndex        =   13
         Top             =   1680
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clr.Res."
         Height          =   195
         Index           =   4
         Left            =   30
         TabIndex        =   11
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Selection"
         Height          =   255
         Index           =   2
         Left            =   30
         TabIndex        =   9
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblSelClr 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fore Color"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   30
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblForeClr 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Base Color"
         Height          =   255
         Index           =   3
         Left            =   30
         TabIndex        =   5
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblBaseClr 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Back Color"
         Height          =   255
         Index           =   0
         Left            =   30
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblBackClr 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   255
      End
   End
   Begin VIPER.La2DGrid La2DGrid1 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   11668
   End
   Begin VB.Menu mnuF 
      Caption         =   "Function"
      Begin VB.Menu mnuFS 
         Caption         =   "Show"
         Begin VB.Menu mnuFShow 
            Caption         =   "Count"
            Index           =   0
         End
         Begin VB.Menu mnuFShow 
            Caption         =   "Abundance"
            Index           =   1
         End
         Begin VB.Menu mnuFShow 
            Caption         =   "Charge"
            Index           =   2
         End
         Begin VB.Menu mnuFShow 
            Caption         =   "Pairs Count"
            Index           =   3
         End
      End
      Begin VB.Menu mnuFSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFExcSel 
         Caption         =   "Exclude Selected"
      End
      Begin VB.Menu mnuFExcNotSel 
         Caption         =   "Exclude Not Selected"
      End
      Begin VB.Menu mnuFSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuFSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmFilterGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'created: 10/20/2002 nt
'last modified: 10/25/2002 nt
'-------------------------------------------------------
Option Explicit

Const SAY_COUNT = 0
Const SAY_ABUNDANCE = 1
Const SAY_CHARGE = 2
Const SAY_PAIRS_COUNT = 3

Dim CallerID As Long
Dim bLoading As Boolean

Dim FNMin As Long
Dim FNMax As Long
Dim ScanRange As Long

Dim ShowWhat As Long

Private Sub cmdClose_Click()
picOptions.Visible = False
End Sub

Private Sub Form_Activate()
If bLoading Then
   CallerID = Me.Tag
   bLoading = False
   GetScanRange CallerID, FNMin, FNMax, ScanRange
   La2DGrid1.MinX = CDbl(FNMin)
   La2DGrid1.MaxX = CDbl(FNMax)
   La2DGrid1.MinY = GelData(CallerID).MinMW
   La2DGrid1.MaxY = GelData(CallerID).MaxMW
End If
End Sub

Private Sub Form_Load()
bLoading = True
La2DGrid1.BackColor = lblBackClr.BackColor
La2DGrid1.ForeColor = lblForeClr.BackColor
La2DGrid1.ValueColor = lblBaseClr.BackColor
La2DGrid1.SelectionColor = lblSelClr.BackColor
La2DGrid1.COLORRES = txtClrRes.Text
La2DGrid1.XRes = txtXRes.Text
La2DGrid1.YRes = txtYRes.Text
End Sub


Private Sub Form_Resize()
La2DGrid1.width = Me.ScaleWidth
La2DGrid1.Height = Me.ScaleHeight
End Sub

Private Sub lblBackClr_DblClick()
Dim TmpClr As Long
On Error GoTo exit_lblBackClr_DblClick
TmpClr = lblBackClr.BackColor
Call GetColorAPIDlg(Me.hwnd, TmpClr)
lblBackClr.BackColor = TmpClr
La2DGrid1.BackColor = TmpClr
exit_lblBackClr_DblClick:
End Sub

Private Sub lblBaseClr_DblClick()
Dim TmpClr As Long
On Error GoTo exit_lblBaseClr_DblClick
TmpClr = lblBaseClr.BackColor
Call GetColorAPIDlg(Me.hwnd, TmpClr)
lblBaseClr.BackColor = TmpClr
La2DGrid1.ValueColor = TmpClr
exit_lblBaseClr_DblClick:
End Sub

Private Sub lblForeClr_DblClick()
Dim TmpClr As Long
On Error GoTo exit_lblForeClr_DblClick
TmpClr = lblForeClr.BackColor
Call GetColorAPIDlg(Me.hwnd, TmpClr)
lblForeClr.BackColor = TmpClr
La2DGrid1.ForeColor = TmpClr
exit_lblForeClr_DblClick:
End Sub

Private Sub lblSelClr_DblClick()
Dim TmpClr As Long
On Error GoTo exit_lblSelClr_DblClick
TmpClr = lblSelClr.BackColor
Call GetColorAPIDlg(Me.hwnd, TmpClr)
lblSelClr.BackColor = TmpClr
La2DGrid1.SelectionColor = TmpClr
exit_lblSelClr_DblClick:
End Sub


Private Sub FillData_Count()
Dim CurrMW As Double
Dim CurrScan As Long
Dim i As Long
With GelData(CallerID)
    For i = 1 To .CSLines
        CurrMW = .CSData(i).AverageMW
        CurrScan = .CSData(i).ScanNumber
        Call La2DGrid1.AddData(CurrScan, CurrMW, 1)
    Next i
    For i = 1 To .IsoLines
        CurrMW = GetIsoMass(.IsoData(i), .Preferences.IsoDataField)
        CurrScan = .IsoData(i).ScanNumber
        Call La2DGrid1.AddData(CurrScan, CurrMW, 1)
    Next i
End With
End Sub

Private Sub FillData_Abundance()
Dim CurrMW As Double
Dim CurrScan As Long
Dim CurrAbu As Double
Dim i As Long
With GelData(CallerID)
    For i = 1 To .CSLines
        CurrMW = .CSData(i).AverageMW
        CurrScan = .CSData(i).ScanNumber
        CurrAbu = CSng(.CSData(i).Abundance)
        Call La2DGrid1.AddData(CurrScan, CurrMW, CurrAbu)
    Next i
    For i = 1 To .IsoLines
        CurrMW = GetIsoMass(.IsoData(i), .Preferences.IsoDataField)
        CurrScan = .IsoData(i).ScanNumber
        CurrAbu = .IsoData(i).Abundance
        Call La2DGrid1.AddData(CurrScan, CurrMW, CurrAbu)
    Next i
End With
End Sub

Private Sub FillData_Charge()
Dim CurrMW As Double
Dim CurrScan As Long
Dim CurrCS As Double
Dim i As Long
With GelData(CallerID)
    For i = 1 To .CSLines
        CurrMW = .CSData(i).AverageMW
        CurrScan = .CSData(i).ScanNumber
        CurrCS = CSng(.CSData(i).Charge)
        Call La2DGrid1.AddData(CurrScan, CurrMW, CurrCS)
    Next i
    For i = 1 To .IsoLines
        CurrMW = GetIsoMass(.IsoData(i), .Preferences.IsoDataField)
        CurrScan = .IsoData(i).ScanNumber
        CurrCS = CSng(.IsoData(i).Charge)
        Call La2DGrid1.AddData(CurrScan, CurrMW, CurrCS)
    Next i
End With
End Sub

Private Sub FillData_Pairs_Count()
Dim CurrMWL As Double           'count once for both heavy and
Dim CurrMWH As Double           'light pair member
Dim CurrScan As Long
Dim RepInd As Long              'index of class representative
Dim IsoField As Integer
Dim i As Long
With GelP_D_L(CallerID)
  Select Case .DltLblType
  Case ptUMCDlt, ptUMCLbl, ptUMCDltLbl        'have to pull from GelUMC
     For i = 0 To .PCnt - 1
         CurrMWL = GelUMC(CallerID).UMCs(.Pairs(i).P1).ClassMW
         'use scan of class representative
         RepInd = GelUMC(CallerID).UMCs(.Pairs(i).P1).ClassRepInd
         Select Case GelUMC(CallerID).UMCs(.Pairs(i).P1).ClassRepType
         Case glCSType
            CurrScan = GelData(CallerID).CSData(RepInd).ScanNumber
         Case glIsoType
            CurrScan = GelData(CallerID).IsoData(RepInd).ScanNumber
         End Select
         Call La2DGrid1.AddData(CurrScan, CurrMWL, 1)
         RepInd = GelUMC(CallerID).UMCs(.Pairs(i).P2).ClassRepInd
         CurrMWH = GelUMC(CallerID).UMCs(.Pairs(i).P2).ClassMW
         Select Case GelUMC(CallerID).UMCs(.Pairs(i).P2).ClassRepType
         Case glCSType
             CurrScan = GelData(CallerID).CSData(RepInd).ScanNumber
         Case glIsoType
             CurrScan = GelData(CallerID).IsoData(RepInd).ScanNumber
         End Select
         Call La2DGrid1.AddData(CurrScan, CurrMWH, 1)
     Next i
  Case ptS_Dlt, ptS_Lbl, ptS_DltLbl           'have to pull from GelData(isotopic data only)
     IsoField = GelData(CallerID).Preferences.IsoDataField
     For i = 0 To .PCnt - 1
         CurrMWL = GetIsoMass(GelData(CallerID).IsoData(.Pairs(i).P1), IsoField)
         CurrMWH = GetIsoMass(GelData(CallerID).IsoData(.Pairs(i).P2), IsoField)
         CurrScan = GelData(CallerID).IsoData(.Pairs(i).P1).ScanNumber
         Call La2DGrid1.AddData(CurrScan, CurrMWL, 1)
         Call La2DGrid1.AddData(CurrScan, CurrMWH, 1)
     Next i
  End Select
End With
End Sub

Private Sub mnuFClose_Click()
Unload Me
End Sub

Private Sub mnuFExcNotSel_Click()
Dim i As Long
Dim CurrMW As Double
Dim CurrScan As Long
Dim ExcCnt As Long
On Error Resume Next
With GelData(CallerID)
    For i = 1 To .CSLines
        CurrMW = .CSData(i).AverageMW
        CurrScan = .CSData(i).ScanNumber
        If Not La2DGrid1.IsInSelection(CurrScan, CurrMW) Then
           ExcCnt = ExcCnt + 1
           GelDraw(CallerID).CSID(i) = -Abs(GelDraw(CallerID).CSID(i))
        End If
    Next i
    For i = 1 To .IsoLines
        CurrMW = GetIsoMass(.IsoData(i), .Preferences.IsoDataField)
        CurrScan = .IsoData(i).ScanNumber
        If Not La2DGrid1.IsInSelection(CurrScan, CurrMW) Then
           ExcCnt = ExcCnt + 1
           GelDraw(CallerID).IsoID(i) = -Abs(GelDraw(CallerID).IsoID(i))
        End If
    Next i
End With
If ExcCnt > 0 Then MsgBox "Number of excluded spots: " & ExcCnt, vbOKOnly, glFGTU
End Sub

Private Sub mnuFExcSel_Click()
Dim i As Long
Dim CurrMW As Double
Dim CurrScan As Long
Dim ExcCnt As Long
On Error Resume Next
With GelData(CallerID)
    For i = 1 To .CSLines
        CurrMW = .CSData(i).AverageMW
        CurrScan = .CSData(i).ScanNumber
        If La2DGrid1.IsInSelection(CurrScan, CurrMW) Then
           ExcCnt = ExcCnt + 1
           GelDraw(CallerID).CSID(i) = -Abs(GelDraw(CallerID).CSID(i))
        End If
    Next i
    For i = 1 To .IsoLines
        CurrMW = GetIsoMass(.IsoData(i), .Preferences.IsoDataField)
        CurrScan = .IsoData(i).ScanNumber
        If La2DGrid1.IsInSelection(CurrScan, CurrMW) Then
           ExcCnt = ExcCnt + 1
           GelDraw(CallerID).IsoID(i) = -Abs(GelDraw(CallerID).IsoID(i))
        End If
    Next i
End With
If ExcCnt > 0 Then MsgBox "Number of excluded spots: " & ExcCnt, vbOKOnly, glFGTU
End Sub

Private Sub mnuFOptions_Click()
If Not picOptions.Visible Then picOptions.Visible = True
End Sub

Private Sub mnuFShow_Click(Index As Integer)
ShowWhat = Index
Call RecalcGraph
End Sub

Private Sub txtClrRes_LostFocus()
On Error Resume Next
La2DGrid1.COLORRES = txtClrRes.Text
Call RecalcGraph
End Sub

Private Sub txtXRes_LostFocus()
On Error Resume Next
La2DGrid1.XRes = txtXRes.Text
Call RecalcGraph
End Sub

Private Sub txtYRes_LostFocus()
On Error Resume Next
La2DGrid1.YRes = txtYRes.Text
Call RecalcGraph
End Sub

Private Sub RecalcGraph()
La2DGrid1.InitData
Select Case ShowWhat
Case SAY_COUNT
     Call FillData_Count
Case SAY_ABUNDANCE
     Call FillData_Abundance
Case SAY_CHARGE
     Call FillData_Charge
Case SAY_PAIRS_COUNT
     If GelP_D_L(CallerID).PCnt > 0 Then
        Call FillData_Pairs_Count
     Else
        MsgBox "Pairs not found; make sure some of Delta-Label pairing function was applied.", vbOKOnly, glFGTU
     End If
End Select
Call UpdateGraph
End Sub

Public Sub UpdateGraph()
La2DGrid1.CalcColors ColorScaleLin
La2DGrid1.Refresh
End Sub

