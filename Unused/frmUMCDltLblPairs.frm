VERSION 5.00
Begin VB.Form frmUMCDltLblPairs 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "UMC N14/N15 Cys-based Labeled Pairing"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5190
   Icon            =   "frmUMCDltLblPairs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkPairsExcludeAmbiguousKeepMostConfident 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ambiguous pairs exclusion keeps most confident pair"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   5280
      Value           =   1  'Checked
      Width           =   4485
   End
   Begin VB.CheckBox chkPairsRequireOverlapAtEdge 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Require pair-classes &overlap at UMC edges"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      ToolTipText     =   "If checked pair classes have to show at least once in the same scan"
      Top             =   4095
      Value           =   1  'Checked
      Width           =   3600
   End
   Begin VB.CheckBox chkPairsRequireOverlapAtApex 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Require pair-classes &overlap at UMC apexes"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      ToolTipText     =   "If checked pair classes have to show at least once in the same scan"
      Top             =   4440
      Width           =   3600
   End
   Begin VB.TextBox txtPairsScanTolEdge 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3840
      TabIndex        =   24
      Text            =   "5"
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox txtPairsScanTolApex 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3840
      TabIndex        =   26
      Text            =   "5"
      Top             =   4425
      Width           =   495
   End
   Begin VB.ComboBox cboAverageERsWeightingMode 
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   6240
      Width           =   2055
   End
   Begin VB.CheckBox chkAverageERsAllChargeStates 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Average ER's for all charge states"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   6285
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.CheckBox chkComputeERScanByScan 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Compute ER Scan by Scan"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   6540
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.CommandButton cmdAbortProcess 
      Caption         =   "Abort"
      Height          =   300
      Left            =   120
      TabIndex        =   38
      Top             =   120
      Width           =   1455
   End
   Begin VB.CheckBox chkRequireMatchingChargeStates 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Require matching charge states for pair"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   5760
      Value           =   1  'Checked
      Width           =   3255
   End
   Begin VB.CheckBox chkUseIdenticalChargeStatesForER 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Use identical charge states for expression ratio"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   6015
      Value           =   1  'Checked
      Width           =   4095
   End
   Begin VB.CheckBox chkAutoMinMaxDelta 
      BackColor       =   &H00C0FFC0&
      Caption         =   "C&alculate N14/N15 Min/Max Deltas from class molecular mass"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1080
      Width           =   2775
   End
   Begin VB.CommandButton cmdSetToO18 
      Caption         =   "Set to O16/O18"
      Height          =   300
      Left            =   2280
      TabIndex        =   9
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdSetToN15 
      Caption         =   "Set to N14/N15"
      Height          =   300
      Left            =   600
      TabIndex        =   8
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtLabel 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   960
      TabIndex        =   11
      Text            =   "442.2249697"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtMaxLbl 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3720
      TabIndex        =   17
      Text            =   "5"
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox txtMinLbl 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      TabIndex        =   15
      Text            =   "1"
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox txtHeavyLightDelta 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3720
      TabIndex        =   13
      Text            =   "8.05"
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox txtMaxLblDiff 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3720
      TabIndex        =   19
      Text            =   "1"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox txtERMax 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3600
      TabIndex        =   30
      Text            =   "5"
      Top             =   4905
      Width           =   855
   End
   Begin VB.TextBox txtERMin 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2160
      TabIndex        =   28
      Text            =   "-5"
      Top             =   4905
      Width           =   855
   End
   Begin VB.TextBox txtMinDelta 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Text            =   "1"
      Top             =   1545
      Width           =   855
   End
   Begin VB.TextBox txtMaxDelta 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3600
      TabIndex        =   7
      Text            =   "100"
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtDeltaTol 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1440
      TabIndex        =   21
      Text            =   "0.02"
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox txtDelta 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "0.9970356"
      Top             =   1140
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   5
      X1              =   120
      X2              =   4560
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label (Lt.):"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   10
      Top             =   2580
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Max Labels:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   9
      Left            =   2280
      TabIndex        =   16
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Min Labels:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   14
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Heavy/Light Delta:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   12
      Top             =   2580
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Max. difference in number of labels in Lt/Hv:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   18
      Top             =   3240
      Width           =   3375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   4
      X1              =   120
      X2              =   4560
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   3
      X1              =   120
      X2              =   4560
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ER Inclusion Range:"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   2
      X1              =   120
      X2              =   4560
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   1
      X1              =   120
      X2              =   4560
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   0
      X1              =   120
      X2              =   4560
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   29
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Min Deltas:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   4
      Top             =   1605
      Width           =   975
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Height          =   480
      Left            =   120
      TabIndex        =   37
      Top             =   7080
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Scan Tolerance:"
      Height          =   255
      Index           =   5
      Left            =   3600
      TabIndex        =   22
      Top             =   3800
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Max Deltas:"
      Height          =   255
      Index           =   4
      Left            =   2280
      TabIndex        =   6
      Top             =   1605
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pair Tolerance:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   20
      Top             =   3780
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Delta:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblSummary 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmUMCDltLblPairs.frx":030A
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Menu mnuFunction 
      Caption         =   "&Function"
      Begin VB.Menu mnuFFindPairs 
         Caption         =   "&Find Pairs"
      End
      Begin VB.Menu mnuFSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFMarkAmbPairs 
         Caption         =   "Mark &Ambiguous Pairs"
      End
      Begin VB.Menu mnuFMarkBadERPairs 
         Caption         =   "&Mark Pairs Out Of ER Range"
      End
      Begin VB.Menu mnuFResetExclusionFlags 
         Caption         =   "Reset Exclusion Flags for All Pairs"
      End
      Begin VB.Menu mnuFSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFClearAllPairs 
         Caption         =   "Clear All &Pairs"
      End
      Begin VB.Menu mnuFDelExcPairs 
         Caption         =   "Delete &Excluded Pairs"
      End
      Begin VB.Menu mnuFSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFDelER 
         Caption         =   "&Clear Pairs ER"
      End
      Begin VB.Menu mnuFERRecalculation 
         Caption         =   "&Recalculate ER"
      End
      Begin VB.Menu mnuFSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Report"
      Begin VB.Menu mnuRPairsAll 
         Caption         =   "Pairs &All"
      End
      Begin VB.Menu mnuRPairsIncluded 
         Caption         =   "Pairs I&ncluded Only"
      End
      Begin VB.Menu mnuRPairsExcluded 
         Caption         =   "Pairs &Excluded Only"
      End
      Begin VB.Menu mnuRERStat 
         Caption         =   "ER &Statistics (Text)"
      End
      Begin VB.Menu mnuRERStatGraph 
         Caption         =   "ER Statistics (Graph)"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmUMCDltLblPairs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'last modified 07/29/2002 nt
'-------------------------------------------------------------------------
'Lower and Upper limit of number of nitrogen atoms for each molecular mass
'estimated based on Gordon's analysis
'NCount=0.012 * MW(Da); lower boundary was taken NCount/2; upper count was
'estimated on 3*NCount/2 (although analysis was not very sophisticated it
'is safe to say that estimate is conservative)
'-------------------------------------------------------------------------

'THIS FUNCTION IS NOT FINISHED NOR TESTED
'In particular, Label and LabelDelta are not used

Option Explicit

Const MAXPAIRS As Long = 10000000

Const DEF_LBL_LT = 414.1936713
Const DEF_LBL_HY = 414.1936713
Const DEF_DELTA = 0.9970356
Const DEF_DELTA_TOL = 0.02
Const DEF_MIN_DELTAS = 1
Const DEF_MAX_DELTAS = 100
Const DEF_SCAN_TOL = 5
Const DEF_ER_MIN = -5
Const DEF_ER_MAX = 5

Dim CallerID As Long
Dim bLoading As Boolean

'ER statistic depends on type of ER calculation but it always uses 1000 bins
'for ratio                  nonequidistant nodes from 0 to 50
'for logarithmic ratio      equidistant nodes from -50 to 50
'for symmetric ratio         equidistant nodes from -50 to 50
Dim ERBin() As Double       'ER nodes
Dim ERBinAll() As Long      'bin count - all pairs
Dim ERBinInc() As Long      'bin count - included pairs
Dim ERBinExc() As Long      'bin count - excluded pairs
Dim ERAllS As ERStatHelper
Dim ERIncS As ERStatHelper
Dim ERExcS As ERStatHelper

Private mPairInfoChanged As Boolean
Private mAbortProcess As Boolean
'

Private Sub cboAverageERsWeightingMode_Click()
    glbPreferencesExpanded.PairSearchOptions.SearchDef.AverageERsWeightingMode = cboAverageERsWeightingMode.ListIndex
End Sub

Private Sub chkAutoMinMaxDelta_Click()
If chkAutoMinMaxDelta.value = vbChecked Then
   txtMinDelta.Enabled = False
   txtMaxDelta.Enabled = False
   glbPreferencesExpanded.PairSearchOptions.SearchDef.AutoCalculateDeltaMinMaxCount = True
Else
   txtMinDelta.Enabled = True
   txtMaxDelta.Enabled = True
   glbPreferencesExpanded.PairSearchOptions.SearchDef.AutoCalculateDeltaMinMaxCount = False
End If
End Sub

Private Sub chkAverageERsAllChargeStates_Click()
    glbPreferencesExpanded.PairSearchOptions.SearchDef.AverageERsAllChargeStates = cChkBox(chkAverageERsAllChargeStates)
    UpdateDynamicControls
End Sub

Private Sub chkComputeERScanByScan_Click()
    glbPreferencesExpanded.PairSearchOptions.SearchDef.ComputeERScanByScan = cChkBox(chkComputeERScanByScan)
End Sub


Private Sub chkPairsExcludeAmbiguousKeepMostConfident_Click()
    glbPreferencesExpanded.PairSearchOptions.KeepMostConfidentAmbiguous = cChkBox(chkPairsExcludeAmbiguousKeepMostConfident)
End Sub

Private Sub chkRequireMatchingChargeStates_Click()
    glbPreferencesExpanded.PairSearchOptions.SearchDef.RequireMatchingChargeStatesForPairMembers = cChkBox(chkRequireMatchingChargeStates)
    UpdateDynamicControls
End Sub

Private Sub chkPairsRequireOverlapAtApex_Click()
    glbPreferencesExpanded.PairSearchOptions.SearchDef.RequireUMCOverlapAtApex = cChkBox(chkPairsRequireOverlapAtApex)
End Sub

Private Sub chkPairsRequireOverlapAtEdge_Click()
    glbPreferencesExpanded.PairSearchOptions.SearchDef.RequireUMCOverlap = cChkBox(chkPairsRequireOverlapAtEdge)
End Sub

Private Sub chkUseIdenticalChargeStatesForER_Click()
    glbPreferencesExpanded.PairSearchOptions.SearchDef.UseIdenticalChargesForER = cChkBox(chkUseIdenticalChargeStatesForER)
    UpdateDynamicControls
End Sub

Private Sub cmdAbortProcess_Click()
    mAbortProcess = True
End Sub

Private Sub cmdSetToN15_Click()
txtDelta = glN14N15_DELTA
glbPreferencesExpanded.PairSearchOptions.SearchDef.DeltaMass = glN14N15_DELTA
End Sub

Private Sub cmdSetToO18_Click()
txtDelta = glO16O18_DELTA
glbPreferencesExpanded.PairSearchOptions.SearchDef.DeltaMass = glO16O18_DELTA
End Sub

Private Sub Form_Activate()
    InitializeForm
End Sub

Private Sub Form_Load()

'set defaults
bLoading = True

With cboAverageERsWeightingMode
    .Clear
    .AddItem "No weighting"
    .AddItem "Weight by Abu"
    .AddItem "Weight by Members"
    .ListIndex = aewAbundance
End With

If CallerID >= 1 And CallerID <= UBound(GelBody) Then
    glbPreferencesExpanded.PairSearchOptions.SearchDef = GelP_D_L(CallerID).SearchDef
End If

With glbPreferencesExpanded.PairSearchOptions.SearchDef
    txtDelta.Text = .DeltaMass
    
    SetCheckBox chkAutoMinMaxDelta, .AutoCalculateDeltaMinMaxCount
    txtMinDelta.Text = .DeltaCountMin
    txtMaxDelta.Text = .DeltaCountMax
    
    txtLabel = .LightLabelMass
    txtHeavyLightDelta = .HeavyLightMassDifference
    
    txtMinLbl = .LabelCountMin
    txtMaxLbl = .LabelCountMax
    txtMaxLblDiff = .MaxDifferenceInNumberOfLightHeavyLabels
    
    txtDeltaTol.Text = .DeltaMassTolerance
    txtPairsScanTolEdge.Text = .ScanTolerance
    txtPairsScanTolApex.Text = .ScanToleranceAtApex
    
    SetCheckBox chkPairsRequireOverlapAtEdge, .RequireUMCOverlap
    SetCheckBox chkPairsRequireOverlapAtApex, .RequireUMCOverlapAtApex
    
    txtERMin.Text = .ERInclusionMin
    txtERMax.Text = .ERInclusionMax
    
    SetCheckBox chkPairsExcludeAmbiguousKeepMostConfident, glbPreferencesExpanded.PairSearchOptions.KeepMostConfidentAmbiguous
    
    SetCheckBox chkRequireMatchingChargeStates, .RequireMatchingChargeStatesForPairMembers
    SetCheckBox chkUseIdenticalChargeStatesForER, .UseIdenticalChargesForER
    SetCheckBox chkComputeERScanByScan, .ComputeERScanByScan
    SetCheckBox chkAverageERsAllChargeStates, .AverageERsAllChargeStates
    
    cboAverageERsWeightingMode.ListIndex = .AverageERsWeightingMode
    UpdateDynamicControls
End With

mPairInfoChanged = False

MsgBox "Warning, the Delta/Label Pairs search algorithm is under development and has not been fully implemented.  In particular, the settings for Label mass and Heavy/Light Delta mass are not utilized in the search."

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mPairInfoChanged Then
        Me.MousePointer = vbHourglass
        UpdateStatus "Filling comparative display structures..."
        Call FillUMC_ERs(CallerID)
        GelBody(CallerID).ResetGraph True, False, GelBody(CallerID).fgDisplay
        Me.MousePointer = vbDefault
    End If
End Sub

Private Sub mnuFClearAllPairs_Click()
ClearAllPairs
End Sub

Private Sub mnuFClose_Click()
Unload Me
End Sub

Private Sub mnuFDelER_Click()
'-----------------------------------------------------
'this will resdimension ER structure and leave type as
'glER_NONE as warning that calculation is not done
'-----------------------------------------------------
InitDltLblPairsER CallerID
mPairInfoChanged = True
End Sub

Private Sub mnuFDelExcPairs_Click()
UpdateStatus "Deleting excluded pairs ...!"
Me.MousePointer = vbHourglass
UpdateStatus DeleteExcludedPairs(CallerID)
mPairInfoChanged = True
Me.MousePointer = vbDefault
End Sub

Private Sub mnuFERRecalculation_Click()
UpdateERCalculationOptions
CalcDltLblPairsER_UMC CallerID
mPairInfoChanged = True
UpdateStatus "Recalculated pair expression ratios; Pair count = " & GelP_D_L(CallerID).PCnt
End Sub

Private Sub mnuFFindPairs_Click()
'--------------------------------
'find all potential pairs
'--------------------------------
Dim eResponse As VbMsgBoxResult
Dim blnSuccess As Boolean
On Error GoTo exit_cmdFindPairs

If GelP_D_L(CallerID).PCnt > 0 Then
'something is already in pairs structure; give user chance to change their mind
   eResponse = MsgBox("Pairs structure already contains pairs! Selected procedure will clear all existing pairs! Continue?", vbOKCancel, glFGTU)
   If eResponse <> vbOK Then Exit Sub
End If

blnSuccess = FindPairsWrapper(True)

exit_cmdFindPairs:
Me.MousePointer = vbDefault
End Sub

Private Sub mnuFMarkAmbPairs_Click()
Dim strMessage As String
strMessage = PairsSearchMarkAmbiguous(Me, CallerID, True)
mPairInfoChanged = True
UpdateStatus strMessage
End Sub

Private Sub mnuFMarkBadERPairs_Click()
'-----------------------------------------------------------------
'mark pair as bad if expression ratio is out of ER inclusion range
'-----------------------------------------------------------------
MarkBadERPairs
End Sub

Private Sub mnuFResetExclusionFlags_Click()
UpdateStatus "Resetting pair exclusion flags...!"
Me.MousePointer = vbHourglass
UpdateStatus PairsResetExclusionFlag(CallerID)
mPairInfoChanged = True
Me.MousePointer = vbDefault
End Sub

Private Sub mnuFunction_Click()
Call PickParameters
End Sub

Private Sub mnuReport_Click()
Call PickParameters
End Sub

Private Sub mnuRERStat_Click()
ReportERStatistics
End Sub

Private Sub mnuRPairsAll_Click()
ReportPairs 0
End Sub

Private Sub mnuRPairsExcluded_Click()
ReportPairs glPAIR_Exc
End Sub

Private Sub mnuRPairsIncluded_Click()
ReportPairs glPAIR_Inc
End Sub

Private Sub txtDelta_LostFocus()
On Error GoTo err_Delta
glbPreferencesExpanded.PairSearchOptions.SearchDef.DeltaMass = CDbl(txtDelta.Text)
If glbPreferencesExpanded.PairSearchOptions.SearchDef.DeltaMass > 0 Then Exit Sub
err_Delta:
MsgBox "Delta should be positive number!", vbOKOnly, glFGTU
txtDelta.SetFocus
End Sub

Private Sub txtDeltaTol_LostFocus()
On Error GoTo err_DeltaTol
glbPreferencesExpanded.PairSearchOptions.SearchDef.DeltaMassTolerance = CDbl(txtDeltaTol.Text)
If glbPreferencesExpanded.PairSearchOptions.SearchDef.DeltaMassTolerance > 0 Then Exit Sub
err_DeltaTol:
MsgBox "Delta tolerance should be positive number!", vbOKOnly, glFGTU
txtDeltaTol.SetFocus
End Sub

Private Sub txtERMax_LostFocus()
On Error GoTo err_ERMax
glbPreferencesExpanded.PairSearchOptions.SearchDef.ERInclusionMax = CDbl(txtERMax.Text)
Exit Sub
err_ERMax:
MsgBox "Maximum of ER range should be a number!", vbOKOnly, glFGTU
txtERMax.SetFocus
End Sub

Private Sub txtERMin_LostFocus()
On Error GoTo err_ERMin
glbPreferencesExpanded.PairSearchOptions.SearchDef.ERInclusionMin = CDbl(txtERMin.Text)
Exit Sub
err_ERMin:
MsgBox "Minimum of ER range should be a number!", vbOKOnly, glFGTU
txtERMin.SetFocus
End Sub

Private Sub txtHeavyLightDelta_LostFocus()
On Error GoTo err_HeavyLabelDelta
glbPreferencesExpanded.PairSearchOptions.SearchDef.HeavyLightMassDifference = CDbl(txtHeavyLightDelta.Text)
If glbPreferencesExpanded.PairSearchOptions.SearchDef.HeavyLightMassDifference >= 0 Then Exit Sub
err_HeavyLabelDelta:
MsgBox "Label mass should be non-negative number!", vbOKOnly, glFGTU
txtHeavyLightDelta.SetFocus
End Sub

Private Sub txtLabel_LostFocus()
On Error GoTo err_Label
glbPreferencesExpanded.PairSearchOptions.SearchDef.LightLabelMass = CDbl(txtLabel.Text)
If glbPreferencesExpanded.PairSearchOptions.SearchDef.LightLabelMass >= 0 Then Exit Sub
err_Label:
MsgBox "Label mass should be non-negative number!", vbOKOnly, glFGTU
txtLabel.SetFocus
End Sub

Private Sub txtMaxDelta_LostFocus()
On Error GoTo err_MaxDelta
glbPreferencesExpanded.PairSearchOptions.SearchDef.DeltaCountMax = CLng(txtMaxDelta.Text)
If glbPreferencesExpanded.PairSearchOptions.SearchDef.DeltaCountMax > 0 Then Exit Sub
err_MaxDelta:
MsgBox "Maximum Delta should be positive integer!", vbOKOnly, glFGTU
txtMaxDelta.SetFocus
End Sub

Private Sub txtMaxLbl_LostFocus()
On Error GoTo err_MaxLbl
glbPreferencesExpanded.PairSearchOptions.SearchDef.LabelCountMax = CLng(txtMaxLbl.Text)
If glbPreferencesExpanded.PairSearchOptions.SearchDef.LabelCountMax > 0 Then Exit Sub
err_MaxLbl:
MsgBox "Maximum number of labels should be non-negative integer!", vbOKOnly, glFGTU
txtMaxLbl.SetFocus
End Sub

Private Sub txtMaxLblDiff_LostFocus()
On Error GoTo err_MaxLblDiff
glbPreferencesExpanded.PairSearchOptions.SearchDef.MaxDifferenceInNumberOfLightHeavyLabels = CLng(txtMaxLblDiff.Text)
If glbPreferencesExpanded.PairSearchOptions.SearchDef.MaxDifferenceInNumberOfLightHeavyLabels >= 0 Then Exit Sub
err_MaxLblDiff:
MsgBox "Maximum difference between number of light and heavy labels should be non-negative integer!", vbOKOnly, glFGTU
txtMaxLblDiff.SetFocus
End Sub

Private Sub txtMinDelta_LostFocus()
On Error GoTo err_MinDelta
glbPreferencesExpanded.PairSearchOptions.SearchDef.DeltaCountMin = CLng(txtMinDelta.Text)
If glbPreferencesExpanded.PairSearchOptions.SearchDef.DeltaCountMin > 0 Then Exit Sub
err_MinDelta:
MsgBox "Minimum Delta should be positive integer!", vbOKOnly, glFGTU
txtMinDelta.SetFocus
End Sub

Private Sub txtMinLbl_LostFocus()
On Error GoTo err_MinLbl
glbPreferencesExpanded.PairSearchOptions.SearchDef.LabelCountMin = CLng(txtMinLbl.Text)
If glbPreferencesExpanded.PairSearchOptions.SearchDef.LabelCountMin >= 0 Then Exit Sub
err_MinLbl:
MsgBox "Minimum number of labels should be non-negative integer!", vbOKOnly, glFGTU
txtMinLbl.SetFocus
End Sub

Private Sub txtPairsScanTolApex_LostFocus()
    On Error GoTo err_ScanTol
    glbPreferencesExpanded.PairSearchOptions.SearchDef.ScanToleranceAtApex = CLng(txtPairsScanTolApex.Text)
    If glbPreferencesExpanded.PairSearchOptions.SearchDef.ScanToleranceAtApex >= 0 Then Exit Sub
err_ScanTol:
    MsgBox "Scan tolerance should be non-negative integer!", vbOKOnly, glFGTU
    txtPairsScanTolApex.SetFocus
End Sub

Private Sub txtPairsScanTolEdge_LostFocus()
    On Error GoTo err_ScanTol
    glbPreferencesExpanded.PairSearchOptions.SearchDef.ScanTolerance = CLng(txtPairsScanTolEdge.Text)
    If glbPreferencesExpanded.PairSearchOptions.SearchDef.ScanTolerance >= 0 Then Exit Sub
err_ScanTol:
    MsgBox "Scan tolerance should be non-negative integer!", vbOKOnly, glFGTU
    txtPairsScanTolEdge.SetFocus
End Sub

Private Sub UpdateStatus(ByVal Status As String)
'-----------------------------------------------
'set status label; entertain user so it doesn't
'freak out before function finishes
'-----------------------------------------------
lblStatus.Caption = Status
DoEvents
End Sub

Private Sub ClearAllPairs()
    mPairInfoChanged = True
    DestroyDltLblPairs CallerID
End Sub

Private Sub CleanPairsERs()
'-------------------------------------------
'this function resets ER in Pairs structure;
'underlying gel does not change
'-------------------------------------------
Dim i As Long
With GelP_D_L(CallerID)
    For i = 0 To .PCnt - 1
        With .Pairs(i)
            .ER = ER_CALC_ERR
            .ERStDev = 0
            .ERChargeStateBasisCount = 0
            ReDim .ERChargesUsed(0)
            .ERMemberBasisCount = 0
        End With
    Next i
End With
End Sub

Public Function FindPairsWrapper(Optional blnShowMessages As Boolean = True) As Boolean
    ' Returns True if success, False if error or searching was cancelled prematurely

Me.MousePointer = vbHourglass

If GelP_D_L(CallerID).PCnt > 0 Then ClearAllPairs

UpdateStatus "Calculating UMC parameters ..."
If GelUMC(CallerID).UMCCnt <= 0 Then
   If blnShowMessages Then MsgBox "Make sure that gel was broken to Unique Mass Classes before applying this function!", vbOKOnly, glFGTU
   FindPairsWrapper = False
Else
   UpdateStatus "Finding pair classes ..."
   FindPairsWrapper = FindPairs(blnShowMessages)
End If

End Function

Private Function FindPairs(Optional blnShowMessages As Boolean = True) As Boolean
'-----------------------------------------------------
'actual pairing function; finds and put into structure
'all potential pairs based on numerical criteria
' Returns True if success, False if error or searching was cancelled prematurely
'-----------------------------------------------------
Dim i As Long, j As Long, k As Long
Dim clsMW As Double             ' Labeled member MW
Dim ClsMinDelta As Long
Dim ClsMaxDelta As Long
Dim ClsMidDelta As Long
Dim LClsMW As Double            ' Light member MW
Dim mwDiff As Double
Dim OverlapOK As Boolean
Dim strMessage As String
Dim blnSuccess As Boolean

Dim ScanMaxAbuLight As Double
Dim ScanMaxAbuHeavy As Double

On Error GoTo err_FindPairs

mPairInfoChanged = True

' Copy current settings to GelP_D_L(Ind)
GelP_D_L(CallerID).SearchDef = glbPreferencesExpanded.PairSearchOptions.SearchDef

mAbortProcess = False
cmdAbortProcess.Visible = True
lblSummary.Visible = False

If InitDltLblPairs(CallerID) Then     'this will reserve 40000 pairs to start with
   With GelP_D_L(CallerID)
     .DltLblType = ptUMCDlt
     .SyncWithUMC = True    'whatever happens we have tried
     .PCnt = 0
   
     For i = 0 To GelUMC(CallerID).UMCCnt - 1
       '--------------------------------------------------------------
       'create all pairs in which class i is the heavy member
       'since this is Delta calculation light member has to be
       'lighter than the heavy member
       '--------------------------------------------------------------
       clsMW = GelUMC(CallerID).UMCs(i).ClassMW
       If .SearchDef.AutoCalculateDeltaMinMaxCount Then         'calculate for this specific mass
          ClsMidDelta = CLng(0.012 * clsMW)
          ClsMinDelta = CLng(0.5 * ClsMidDelta)
          ClsMaxDelta = CLng(1.5 * ClsMidDelta)
       Else                            'use whatever is provided
          ClsMinDelta = .SearchDef.DeltaCountMin
          ClsMaxDelta = .SearchDef.DeltaCountMax
       End If
       UpdateStatus "Class: " & i & " MW: " & Format(clsMW, ".0000") & " Delta from " & ClsMinDelta & " to " & ClsMaxDelta
       If mAbortProcess Then Exit For
       For j = 0 To GelUMC(CallerID).UMCCnt - 1
         If i <> j Then
            'check is 'overlap' condition required and if yes do class i and j overlap at the edges?
            OverlapOK = True
            If .SearchDef.RequireUMCOverlap Then
                If ((GelUMC(CallerID).UMCs(j).MaxScan < GelUMC(CallerID).UMCs(i).MinScan) Or _
                   (GelUMC(CallerID).UMCs(i).MaxScan < GelUMC(CallerID).UMCs(j).MinScan)) Then
                        OverlapOK = False                       'no overlap
                End If
            End If
            
            If .SearchDef.RequireUMCOverlapAtApex And OverlapOK Then
                Select Case GelUMC(CallerID).UMCs(i).ClassRepType
                Case glCSType
                    ScanMaxAbuLight = GelData(CallerID).CSNum(GelUMC(CallerID).UMCs(i).ClassRepInd, csfScan)
                    ScanMaxAbuHeavy = GelData(CallerID).CSNum(GelUMC(CallerID).UMCs(j).ClassRepInd, csfScan)
                Case glIsoType
                    ScanMaxAbuLight = GelData(CallerID).IsoNum(GelUMC(CallerID).UMCs(i).ClassRepInd, isfScan)
                    ScanMaxAbuHeavy = GelData(CallerID).IsoNum(GelUMC(CallerID).UMCs(j).ClassRepInd, isfScan)
                Case Else
                    ' This shouldn't happen
                    Debug.Assert False
                    ScanMaxAbuLight = 0
                    ScanMaxAbuHeavy = .SearchDef.ScanToleranceAtApex + 1
                End Select
            
                If Abs(ScanMaxAbuLight - ScanMaxAbuHeavy) > .SearchDef.ScanToleranceAtApex Then
                    OverlapOK = False                       'no overlap
                End If
            
            End If
            
            If OverlapOK Then
              'check is it possible that this is a pair
              LClsMW = GelUMC(CallerID).UMCs(j).ClassMW
              If LClsMW < clsMW Then
                 For k = ClsMinDelta To ClsMaxDelta
                     mwDiff = clsMW - (LClsMW + k * .SearchDef.DeltaMass)
                     If Abs(mwDiff) <= .SearchDef.DeltaMassTolerance Then
                       ' See if pairs overlap within Scan Tolerance
                       If PairsOverlapAtEdgesWithinTol(CallerID, i, j, .SearchDef.ScanTolerance) Then
                         ' If necessary, see if pairs contain matching charge states
                         If ChargeStatesMatch(CallerID, i, j) Or Not .SearchDef.RequireMatchingChargeStatesForPairMembers Then
                            .PCnt = .PCnt + 1
                            With .Pairs(.PCnt - 1)
                                .P1 = j
                                .P2 = i
                                .P2DltCnt = k
                            End With
                          End If
                        End If
                     End If
                     If mwDiff < -.SearchDef.DeltaMassTolerance Then Exit For  'you can stop with class j
                 Next k
              End If
            End If
         End If
       Next j
     Next i
   End With
    
    If Not GelAnalysis(CallerID) Is Nothing Then
        GelAnalysis(CallerID).MD_Type = stPairsPEON14N15
    End If
       
    'MonroeMod
    With GelP_D_L(CallerID)
        strMessage = "Searched for Delta-Label pairs (using UMC's); Pair Count = " & Trim(.PCnt)
        With .SearchDef
            strMessage = strMessage & "; Delta = " & Trim(.DeltaMass) & " Da; Pair Tolerance = " & Trim(.DeltaMassTolerance)
            strMessage = strMessage & " Da; Label = " & Trim(.LightLabelMass) & " Da; Heavy/Light Delta = " & Trim(.HeavyLightMassDifference)
            strMessage = strMessage & " Da; Min Labels = " & Trim(.LabelCountMin) & "; Max Labels = " & Trim(.LabelCountMax)
            strMessage = strMessage & "; Max difference in number of labels = " & Trim(.MaxDifferenceInNumberOfLightHeavyLabels)
            strMessage = strMessage & "; Auto-calculated Min/Max Delta = " & CStr(.AutoCalculateDeltaMinMaxCount)
            strMessage = strMessage & "; Min Deltas = " & Trim(.DeltaCountMin) & "; Max Deltas = " & Trim(.DeltaCountMax)
            strMessage = strMessage & "; " & PairsSearchGenerateDescription(CallerID)
        End With
        AddToAnalysisHistory CallerID, strMessage
    End With
    
    'calculate expression ratios here (note that GelP_D_L().SearchDef was updated earlier in this function)
    CalcDltLblPairsER_UMC CallerID
    
    blnSuccess = True
Else
    strMessage = "Unable to reserve space for pairs structures!"
    If blnShowMessages Then
        MsgBox strMessage, vbOKOnly, glFGTU
    Else
        LogErrors -1, "frmUMCDltLblPairs.FindPairs", strMessage
    End If
    blnSuccess = False
End If

exit_FindPairs:
If GelP_D_L(CallerID).PCnt > 0 Then
  Call TrimDltLblPairs(CallerID)
Else
  DestroyDltLblPairs CallerID, False
End If
If blnShowMessages Then MsgBox "Total number of unique pairs (uncleaned): " & GelP_D_L(CallerID).PCnt

exit_Cleanup:
cmdAbortProcess.Visible = False
lblSummary.Visible = True

FindPairs = blnSuccess
Exit Function

err_FindPairs:
Select Case Err.Number
Case 9
'increase array size and resume
  If UBound(GelP_D_L(CallerID).Pairs) > MAXPAIRS Then
     strMessage = "Number of detected pairs too high!(max. number of pairs " & MAXPAIRS & ")"
     If blnShowMessages Then
        MsgBox strMessage, vbOKOnly
     Else
        AddToAnalysisHistory CallerID, strMessage
     End If
     blnSuccess = False
     Resume exit_FindPairs
  Else
     If AddDltLblPairs(CallerID, 10000) Then Resume
  End If
Case Else
  MsgBox "Error establishing delta-labeled pairs " & vbCrLf _
    & "Error: " & Err.Number & ", " & Err.Description, vbOKOnly, glFGTU
End Select

blnSuccess = False
Resume exit_Cleanup

End Function


Private Function GenerateERStat() As Boolean
'-----------------------------------------------------
'do actual ER statistics for all currently included
'and currently excluded pairs
'-----------------------------------------------------
Dim i As Long
Dim BinInd As Long
Dim Done As Boolean
On Error Resume Next

With GelP_D_L(CallerID)
  If .PCnt > 0 Then
     ReDim ERBin(1000)
     ReDim ERBinAll(1000)
     ReDim ERBinInc(1000)
     ReDim ERBinExc(1000)
     ERAllS.ERCnt = 0
     ERAllS.ERBadL = 0
     ERAllS.ERBadR = 0
     ERIncS.ERCnt = 0
     ERIncS.ERBadL = 0
     ERIncS.ERBadR = 0
     ERExcS.ERCnt = 0
     ERExcS.ERBadL = 0
     ERExcS.ERBadR = 0
     
     Select Case GelP_D_L(CallerID).SearchDef.ERCalcType
     Case ectER_RAT                       'cover range from 0 to 50
        ERBin(500) = 1
        For i = 1 To 500
            ERBin(500 + i) = 1 + i * 0.1
            ERBin(500 - i) = 1 / ERBin(500 + i)
        Next i
     Case ectER_LOG                       'cover range from -10 to 10
        ERBin(500) = 0
        For i = 1 To 500
            ERBin(500 - i) = -i * 0.02
            ERBin(500 + i) = i * 0.02
        Next i
     Case ectER_ALT                       'cover range from -50 to 50
        ERBin(500) = 0
        For i = 1 To 500
            ERBin(500 - i) = -i * 0.1
            ERBin(500 + i) = i * 0.1
        Next i
     End Select
        
     For i = 0 To .PCnt - 1
           'find bin for this expression ratio
           BinInd = -1
           Done = False
           Do Until Done
              If .Pairs(i).ER < ERBin(BinInd + 1) Then
                 Done = True
              Else
                 BinInd = BinInd + 1
                 If BinInd >= 1000 Then Done = True
              End If
           Loop
           'add counts
           ERAllS.ERCnt = ERAllS.ERCnt + 1
           If .Pairs(i).STATE = glPAIR_Inc Then
              ERIncS.ERCnt = ERIncS.ERCnt + 1
           ElseIf .Pairs(i).STATE = glPAIR_Exc Then
              ERExcS.ERCnt = ERExcS.ERCnt + 1
           End If
           Select Case BinInd
           Case Is < 0
                ERAllS.ERBadL = ERAllS.ERBadL + 1
                If .Pairs(i).STATE = glPAIR_Inc Then
                    ERIncS.ERBadL = ERIncS.ERBadL + 1
                ElseIf .Pairs(i).STATE = glPAIR_Exc Then
                    ERExcS.ERBadL = ERExcS.ERBadL + 1
                End If
           Case Is > 1000
                ERAllS.ERBadR = ERAllS.ERBadR + 1
                If .Pairs(i).STATE = glPAIR_Inc Then
                    ERIncS.ERBadR = ERIncS.ERBadR + 1
                ElseIf .Pairs(i).STATE = glPAIR_Exc Then
                    ERExcS.ERBadR = ERExcS.ERBadR + 1
                End If
           Case Else            'some of our cases
                ERBinAll(BinInd) = ERBinAll(BinInd) + 1
                If .Pairs(i).STATE = glPAIR_Inc Then
                   ERBinInc(BinInd) = ERBinInc(BinInd) + 1
                ElseIf .Pairs(i).STATE = glPAIR_Exc Then
                   ERBinExc(BinInd) = ERBinExc(BinInd) + 1
                End If
           End Select
       Next i
       GenerateERStat = True
    Else
        MsgBox "No pairs found! Find Pairs function should be used first!"
    End If
End With
End Function

Public Sub InitializeForm()
If bLoading Then
   bLoading = False
   CallerID = Me.Tag
End If
End Sub

Public Sub MarkBadERPairs()
Dim strMessage As String
strMessage = PairsSearchMarkBadER(glbPreferencesExpanded.PairSearchOptions.SearchDef.ERInclusionMin, glbPreferencesExpanded.PairSearchOptions.SearchDef.ERInclusionMax, CallerID, True)
mPairInfoChanged = True
UpdateStatus strMessage
End Sub

Private Sub PickParameters()
'--------------------------------------------------------
'click on the menu bar does not trigger LostFocus event;
'we have to pick parameters after menu is clicked to make
'sure we have most recent typed parameters
'--------------------------------------------------------
Call txtDelta_LostFocus
Call txtMinDelta_LostFocus
Call txtMaxDelta_LostFocus
Call txtLabel_LostFocus
Call txtHeavyLightDelta_LostFocus
Call txtMinLbl_LostFocus
Call txtMaxLbl_LostFocus
Call txtMaxLblDiff_LostFocus
Call txtDeltaTol_LostFocus
Call txtPairsScanTolEdge_LostFocus
Call txtPairsScanTolApex_LostFocus
Call txtERMin_LostFocus
Call txtERMax_LostFocus
End Sub

Public Sub ReportPairs(PState As Integer, Optional strFilePath As String = "")
' PState can be 0 for all pairs, 1 for Included only (aka glPAIR_Inc), or
'   -1 for Excluded only (aka glPAIR_Exc)
' If Len(strFilePath) = 0, then displays report using frmDataInfo;
'   otherwise, saves the report to strFilePath

UpdateStatus "Generating report ...!"
Me.MousePointer = vbHourglass
ReportDltLblPairsUMCWrapper CallerID, PState, strFilePath
Me.MousePointer = vbDefault
UpdateStatus ""

End Sub

Public Sub ReportERStatistics(Optional strFilePath As String = "")
' If Len(strFilePath) = 0, then displays report using frmDataInfo;
'   otherwise, saves the report to strFilePath

UpdateStatus "Generating report ...!"
Me.MousePointer = vbHourglass
If GenerateERStat Then
   ReportERStat CallerID, ERBin(), ERBinAll(), ERBinInc(), _
                ERBinExc(), ERAllS, ERIncS, ERExcS, strFilePath
End If
Me.MousePointer = vbDefault
UpdateStatus ""

End Sub

Private Sub UpdateDynamicControls()
    With glbPreferencesExpanded.PairSearchOptions.SearchDef
        chkAverageERsAllChargeStates.Enabled = (cChkBox(chkUseIdenticalChargeStatesForER) And cChkBox(chkRequireMatchingChargeStates))
        cboAverageERsWeightingMode.Enabled = (cChkBox(chkAverageERsAllChargeStates) And chkAverageERsAllChargeStates.Enabled)
    End With
End Sub

Private Sub UpdateERCalculationOptions()

    With glbPreferencesExpanded.PairSearchOptions.SearchDef
        GelP_D_L(CallerID).SearchDef.RequireMatchingChargeStatesForPairMembers = .RequireMatchingChargeStatesForPairMembers
        GelP_D_L(CallerID).SearchDef.UseIdenticalChargesForER = .UseIdenticalChargesForER
        GelP_D_L(CallerID).SearchDef.ComputeERScanByScan = .ComputeERScanByScan
        GelP_D_L(CallerID).SearchDef.AverageERsAllChargeStates = .AverageERsAllChargeStates
        GelP_D_L(CallerID).SearchDef.AverageERsWeightingMode = .AverageERsWeightingMode
    End With
     
End Sub

