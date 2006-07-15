VERSION 5.00
Begin VB.Form frmUMC 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Unique Molecular Mass Classes Definition"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbUMCMW 
      Height          =   315
      ItemData        =   "frmUMC.frx":0000
      Left            =   2640
      List            =   "frmUMC.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   2160
      Width           =   2775
   End
   Begin VB.ComboBox cmbUMCAbu 
      Height          =   315
      ItemData        =   "frmUMC.frx":0004
      Left            =   2640
      List            =   "frmUMC.frx":000B
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "&Report"
      Height          =   375
      Left            =   3720
      TabIndex        =   24
      ToolTipText     =   "Generates various statistics on current UMC"
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox txtHolePct 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5280
      TabIndex        =   22
      Text            =   "0"
      Top             =   3780
      Width           =   495
   End
   Begin VB.TextBox txtHoleSize 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5280
      TabIndex        =   21
      Text            =   "0"
      Top             =   3300
      Width           =   495
   End
   Begin VB.TextBox txtHoleNum 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5280
      TabIndex        =   20
      Text            =   "0"
      Top             =   2820
      Width           =   495
   End
   Begin VB.Frame fraTol 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Molecular Mass Tolerance"
      Height          =   1095
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   2295
      Begin VB.TextBox txtTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   15
         Text            =   "10"
         Top             =   520
         Width           =   735
      End
      Begin VB.OptionButton optTolType 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Dalton"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   14
         Top             =   666
         Width           =   855
      End
      Begin VB.OptionButton optTolType 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&ppm"
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   13
         Top             =   333
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Tolerance:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   280
         Width           =   735
      End
   End
   Begin VB.ComboBox cmbCountType 
      Height          =   315
      ItemData        =   "frmUMC.frx":001A
      Left            =   2640
      List            =   "frmUMC.frx":0021
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   720
      Width           =   2775
   End
   Begin VB.Frame fraMWField 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Molecular Mass Field"
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2295
      Begin VB.OptionButton optMWField 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&The Most Abundant"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   920
         Width           =   1815
      End
      Begin VB.OptionButton optMWField 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Monoisotopic"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optMWField 
         BackColor       =   &H00C0FFC0&
         Caption         =   "A&verage"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   280
         Width           =   1335
      End
   End
   Begin VB.Frame fraUMCScope 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Definition Scope"
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2295
      Begin VB.OptionButton optDefScope 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Current View"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optDefScope 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&All Data Points"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   280
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&UMC"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      ToolTipText     =   "Generates UM Classes and returns number of it"
      Top             =   4320
      Width           =   975
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   120
      X2              =   2400
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   120
      X2              =   2400
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   120
      X2              =   2400
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   2400
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Class Molecular Mass"
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   27
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Class Abundance"
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   25
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   4800
      Width           =   5655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Percentage of allowed scan holes in the Unique Mass Class:"
      Height          =   495
      Index           =   2
      Left            =   2640
      TabIndex        =   19
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum size of scan hole in the Unique Mass Class:"
      Height          =   495
      Index           =   1
      Left            =   2640
      TabIndex        =   18
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum number of scan holes in the Unique Mass Class:"
      Height          =   495
      Index           =   0
      Left            =   2640
      TabIndex        =   17
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Count Type"
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   11
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblGelName 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmUMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'unique mass class function
'breaks gel to unique mass classes
'-----------------------------------------
'last modified 11/21/2001 nt
'-----------------------------------------
Option Explicit
Dim CallerID As Long

Private Sub FillComboBoxes()
    With cmbCountType
        .Clear
        .AddItem "Favor Higher Intensity"
        .AddItem "Favor Better Fit"
        .AddItem "Minimize Count"
        .AddItem "Maximize Count"
        .AddItem "Unique AMT"
        .AddItem "FHI - Shrinking Box"
    End With
    
    With cmbUMCAbu
        .Clear
        .AddItem "Average of Class Abu."
        .AddItem "Sum of Class Abu."
        .AddItem "Abu. of Class Representative"
    End With
    
    With cmbUMCMW
        .Clear
        .AddItem "Class Average"
        .AddItem "Mol.Mass Of Class Representative"
        .AddItem "Class Median"
    End With
End Sub

Public Sub InitializeUMCSearch()
CallerID = Me.Tag
' MonroeMod
If CallerID >= 1 And CallerID <= UBound(GelBody) Then UMCDef = GelSearchDef(CallerID).UMCDef
lblGelName.Caption = GelBody(CallerID).Caption

' MonroeMod: This code used to be in Form_Load
With UMCDef
    txtTol.Text = .Tol
    cmbCountType.ListIndex = .UMCType
    cmbUMCAbu.ListIndex = .ClassAbu
    cmbUMCMW.ListIndex = .ClassMW
    optDefScope(.DefScope).value = True
    optMWField(.MWField - MW_FIELD_OFFSET).value = True
    Select Case .TolType
    Case glTOLERANCE_PPM
      optTolType(0).value = True
    Case glTOLERANCE_ABS
      optTolType(1).value = True
    End Select
    txtHoleNum.Text = .GapMaxCnt
    txtHoleSize.Text = .GapMaxSize
    txtHolePct.Text = Round(CLng(.GapMaxPct * 100), 2)
End With
End Sub

Private Sub cmbUMCAbu_Click()
UMCDef.ClassAbu = cmbUMCAbu.ListIndex
End Sub

Private Sub cmbUMCMW_Click()
UMCDef.ClassMW = cmbUMCMW.ListIndex
End Sub

Private Sub Form_Activate()
InitializeUMCSearch
End Sub

Private Sub Form_Load()
    ' MonroeMod: The code that was here has been moved to Form_Activate
    '            This was done so that the Statement: UMCDef = GelSearchDef(CallerID).UMCDef
    '             will be encountered before the controls are updated
    FillComboBoxes
End Sub

Private Sub cmbCountType_Click()
UMCDef.UMCType = cmbCountType.ListIndex
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim TtlCnt As Long
Dim Cnt As Long
On Error Resume Next
Me.MousePointer = vbHourglass
TtlCnt = GelData(CallerID).CSLines + GelData(CallerID).IsoLines
Cnt = UMCCount(CallerID, TtlCnt, Me)
Me.MousePointer = vbDefault
If Cnt >= 0 Then
   Status "(-: Unique Count Done :-)"
   MsgBox "Number of Unique Mass Classes: " & Cnt
   Status ""
    
    ' MonroeMod
    GelSearchDef(CallerID).UMCDef = UMCDef
    AddToAnalysisHistory CallerID, "Identified UMC's (" & UMC_SEARCH_MODE_SETTING_TEXT & ": UMC2002); UMC Count = " & Trim(Cnt) & "; Mass Tolerance = ±" & Trim(UMCDef.Tol) & " " & GetSearchToleranceUnitText(samtDef.TolType) & "; Max # scan holes = " & Trim(UMCDef.GapMaxCnt) & "; Max size of holes = " & Trim(UMCDef.GapMaxSize) & " scans; Allowed % of gaps = " & Format(UMCDef.GapMaxPct, "#00%")
Else
   Status ")-: Error counting Unique Mass Classes :-("
End If

' The following calls CalculateClasses, UpdateIonToUMCIndices, and InitDrawUMC
UpdateUMCStatArrays CallerID, False, Me

'if there is new UMC count, everything done with pairs
'has to be redone if pairs are UMC pairs
If GelP_D_L(CallerID).DltLblType > 2000 Then
   GelP_D_L(CallerID).SyncWithUMC = False
End If
End Sub

Private Sub cmdReport_Click()
Me.MousePointer = vbHourglass
Status "Generating UMC report..."
Call ReportUMC(CallerID, "UMC 2002")
Status ""
Me.MousePointer = vbDefault
End Sub

Private Sub optDefScope_Click(Index As Integer)
UMCDef.DefScope = Index
End Sub

Private Sub optTolType_Click(Index As Integer)
If Index = 0 Then
   UMCDef.TolType = glTOLERANCE_PPM
Else
   UMCDef.TolType = glTOLERANCE_ABS
End If
End Sub

Private Sub optMWField_Click(Index As Integer)
UMCDef.MWField = 6 + Index
End Sub

Private Sub txtHoleNum_LostFocus()
If IsNumeric(txtHoleNum.Text) Then
   UMCDef.GapMaxCnt = CLng(txtHoleNum.Text)
Else
   MsgBox "This argument should be integer value!", vbOKOnly
   txtHoleNum.SetFocus
End If
End Sub

Private Sub txtHolePct_LostFocus()
If IsNumeric(txtHolePct.Text) Then
   UMCDef.GapMaxPct = CDbl(txtHolePct.Text) / 100
Else
   MsgBox "This argument should be numeric value!", vbOKOnly
   txtHolePct.SetFocus
End If
End Sub

Private Sub txtHoleSize_LostFocus()
If IsNumeric(txtHoleSize.Text) Then
   UMCDef.GapMaxSize = CLng(txtHoleSize.Text)
Else
   MsgBox "This argument should be integer value!", vbOKOnly
   txtHoleSize.SetFocus
End If
End Sub

Private Sub txtTol_LostFocus()
If IsNumeric(txtTol.Text) Then
   UMCDef.Tol = txtTol.Text
Else
   MsgBox "Molecular Mass Tolerance should be numeric value!", vbOKOnly
   txtTol.SetFocus
End If
End Sub

Public Sub Status(ByVal StatusText As String)
lblStatus.Caption = StatusText
DoEvents
End Sub
