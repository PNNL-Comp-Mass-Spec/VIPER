VERSION 5.00
Begin VB.Form frmUMCDltPairs 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "UMC Delta Pairing Analysis"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5445
   Icon            =   "frmUMCDltPairs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdResetToDefaults 
      Caption         =   "Set to Defaults"
      Height          =   300
      Index           =   1
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdFindPairs 
      Caption         =   "Find Pairs"
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame fraLabelOptions 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label Mass Options"
      Height          =   2150
      Left            =   5400
      TabIndex        =   17
      Top             =   480
      Width           =   5175
      Begin VB.CommandButton cmdSetToICAT 
         Caption         =   "Set to ICAT"
         Height          =   300
         Left            =   240
         TabIndex        =   28
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtMaxLblDiff 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3840
         TabIndex        =   27
         Text            =   "1"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtHeavyLightDelta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3840
         TabIndex        =   21
         Text            =   "8.05"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtMinLbl 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   23
         Text            =   "1"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtMaxLbl 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3840
         TabIndex        =   25
         Text            =   "5"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtLabel 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   19
         Text            =   "442.2249697"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Max. difference in number of labels in Lt/Hv:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Heavy/Light Delta:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   2400
         TabIndex        =   20
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Min Labels:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Max Labels:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   24
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label (Lt.):"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.Frame fraInclusionOptions 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Inclusion/Exclusion Options"
      Height          =   975
      Left            =   120
      TabIndex        =   37
      Top             =   4680
      Width           =   5175
      Begin VB.TextBox txtERMin 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2400
         TabIndex        =   39
         Text            =   "-5"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtERMax 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3720
         TabIndex        =   41
         Text            =   "5"
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chkPairsExcludeAmbiguousKeepMostConfident 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Ambiguous pairs exclusion keeps most confident pair"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   615
         Value           =   1  'Checked
         Width           =   4485
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "to"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   40
         Top             =   255
         Width           =   255
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ER Inclusion Range:"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   255
         Width           =   2175
      End
   End
   Begin VB.Frame fraGeneralOptions 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Pair Search and ER Calculation Options"
      Height          =   2535
      Left            =   120
      TabIndex        =   43
      Top             =   5760
      Width           =   5175
      Begin VB.CheckBox chkOutlierRemovalUsesSymmetricERs 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Use symmetric ERs"
         Height          =   300
         Left            =   360
         TabIndex        =   52
         Top             =   2160
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.TextBox txtRemoveOutlierERsMinimumDataPointCount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4320
         TabIndex        =   53
         Text            =   "3"
         Top             =   2000
         Width           =   615
      End
      Begin VB.CheckBox chkRemoveOutlierERsIterate 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Repeatedly remove outliers"
         Height          =   300
         Left            =   360
         TabIndex        =   51
         Top             =   1920
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox chkRemoveOutlierERs 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Remove outlier ER values using Grubb's test (95% conf.)"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1680
         Value           =   1  'Checked
         Width           =   4695
      End
      Begin VB.CheckBox chkIReportEREnabled 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Enable I-Report ER computation"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   1360
         Value           =   1  'Checked
         Width           =   3375
      End
      Begin VB.ComboBox cboAverageERsWeightingMode 
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox chkAverageERsAllChargeStates 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Average ER's for all charge states"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   765
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox chkComputeERScanByScan 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Compute ER Scan by Scan"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox chkRequireMatchingChargeStates 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Require matching charge states for pair"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Value           =   1  'Checked
         Width           =   3255
      End
      Begin VB.CheckBox chkUseIdenticalChargeStatesForER 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Use identical charge states for expression ratio"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   495
         Value           =   1  'Checked
         Width           =   4095
      End
      Begin VB.Label lblRemoveOutlierERsMinimumDataPointCount 
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum final data point count"
         Height          =   400
         Left            =   2880
         TabIndex        =   55
         Top             =   1940
         Width           =   1455
      End
   End
   Begin VB.Frame fraToleranceOptions 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Tolerance Options"
      Height          =   1935
      Left            =   120
      TabIndex        =   29
      Top             =   2640
      Width           =   5175
      Begin VB.TextBox txtPairTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   31
         Text            =   "0.02"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtPairsScanTolApex 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3840
         TabIndex        =   36
         Text            =   "15"
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtPairsScanTolEdge 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3840
         TabIndex        =   34
         Text            =   "15"
         Top             =   600
         Width           =   495
      End
      Begin VB.CheckBox chkPairsRequireOverlapAtApex 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Require pair-classes &overlap at UMC apexes"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         ToolTipText     =   "If checked pair classes have to show at least once in the same scan"
         Top             =   1560
         Value           =   1  'Checked
         Width           =   3600
      End
      Begin VB.CheckBox chkPairsRequireOverlapAtEdge 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Require pair-classes &overlap at UMC edges"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         ToolTipText     =   "If checked pair classes have to show at least once in the same scan"
         Top             =   615
         Value           =   1  'Checked
         Width           =   3600
      End
      Begin VB.Label lblPairsRequireOverlapAtEdge 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmUMCDltPairs.frx":030A
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   58
         Top             =   920
         Width           =   4815
      End
      Begin VB.Label lblUnits 
         BackStyle       =   0  'Transparent
         Caption         =   "Da"
         Height          =   255
         Left            =   2400
         TabIndex        =   57
         Top             =   250
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Scan Tolerance:"
         Height          =   255
         Index           =   5
         Left            =   3600
         TabIndex        =   33
         Top             =   315
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pair Tolerance:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   30
         Top             =   250
         Width           =   1335
      End
   End
   Begin VB.Frame fraDeltaOptions 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Delta Mass Options"
      Height          =   2145
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   5175
      Begin VB.TextBox txtDeltaStepSize 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         TabIndex        =   12
         Text            =   "1"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtDelta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Text            =   "0.9970356"
         Top             =   300
         Width           =   855
      End
      Begin VB.CheckBox chkAutoMinMaxDelta 
         BackColor       =   &H00C0FFC0&
         Caption         =   "C&alculate N14/N15 Min/Max Deltas from class molecular mass"
         Height          =   375
         Left            =   2280
         TabIndex        =   6
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtMinDelta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Text            =   "1"
         Top             =   705
         Width           =   855
      End
      Begin VB.TextBox txtMaxDelta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3720
         TabIndex        =   10
         Text            =   "100"
         Top             =   720
         Width           =   855
      End
      Begin VB.Frame fraControls 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   720
         Left            =   120
         TabIndex        =   13
         Top             =   1395
         Width           =   4935
         Begin VB.CommandButton cmdSetDeuterium 
            Caption         =   "Set to Deuterium"
            Height          =   300
            Left            =   1680
            TabIndex        =   56
            Top             =   320
            Width           =   1455
         End
         Begin VB.CommandButton cmdSetToC13 
            Caption         =   "Set to C12/C13"
            Height          =   300
            Left            =   1680
            TabIndex        =   15
            Top             =   0
            Width           =   1455
         End
         Begin VB.CommandButton cmdSetToO18 
            Caption         =   "Set to O16/O18"
            Height          =   300
            Left            =   3360
            TabIndex        =   16
            Top             =   0
            Width           =   1455
         End
         Begin VB.CommandButton cmdSetToN15 
            Caption         =   "Set to N14/N15"
            Height          =   300
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Delta count step size:"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Delta:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Min Deltas:"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   7
         Top             =   765
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Max Deltas:"
         Height          =   255
         Index           =   4
         Left            =   2400
         TabIndex        =   9
         Top             =   765
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdAbortProcess 
      Caption         =   "Abort"
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Height          =   840
      Left            =   120
      TabIndex        =   54
      Top             =   8280
      Width           =   5295
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
         Caption         =   "Exclude &Ambiguous Pairs (all pairs)"
      End
      Begin VB.Menu mnuFMarkAmbPairsHitsOnly 
         Caption         =   "Exclude Ambiguous Pairs (only those with hits)"
      End
      Begin VB.Menu mnuFMarkBadERPairs 
         Caption         =   "&Exclude Pairs Out Of ER Range"
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
         Caption         =   "&Delete Excluded Pairs"
      End
      Begin VB.Menu mnuFAutoClearPairsWhenFindingPairs 
         Caption         =   "Auto-clear existing pairs when finding new pairs"
         Checked         =   -1  'True
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
         Caption         =   "E&xit"
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
Attribute VB_Name = "frmUMCDltPairs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'last modified 07/29/2002 nt
'-------------------------------------------------------------------------
'This form is derived from frmN14N15NoIDUMC; it should work
'very similar but resulting pairs should be stored in GelP_D_L
'structure with type glPDL_N14_N15_UMC
'-------------------------------------------------------------------------
'Lower and Upper limit of number of nitrogen atoms for each molecular mass
'estimated based on Gordon's analysis
'NCount=0.012 * MW(Da); lower boundary was taken NCount/2; upper count was
'estimated on 3*NCount/2 (although analysis was not very sophisticated it
'is safe to say that estimate is conservative)
'-------------------------------------------------------------------------

'When mFormMode = pfmDelta, then searching for delta pairs
'When mFormMode = pfmLabel, then searching for labeled pairs
'   Note: If number of labels could differ for light and heavy pair members
'   then light pair member could be heavier than heavy pair member
'When mFormMode = pfmDeltaLabel, then searching for delta-label pairs

Option Explicit

Private Const MAXPAIRS As Long = 10000000

Public Enum pfmPairFormMode
    pfmDelta = 0
    pfmLabel = 1
    pfmDeltaLabel = 2
End Enum

Public CallerID As Long
Private mFormMode As pfmPairFormMode
Private bLoading As Boolean

'ER statistic depends on type of ER calculation but it always uses 1000 bins
'for ratio                  nonequidistant nodes from 0 to 50
'for logarithmic ratio      equidistant nodes from -50 to 50
'for symmetric ratio         equidistant nodes from -50 to 50
Private ERBin() As Double       'ER nodes
Private ERBinAll() As Long      'bin count - all pairs
Private ERBinInc() As Long      'bin count - included pairs
Private ERBinExc() As Long      'bin count - excluded pairs
Private ERAllS As ERStatHelper
Private ERIncS As ERStatHelper
Private ERExcS As ERStatHelper

Private mPairInfoChanged As Boolean
Private mAbortProcess As Boolean
'

Public Property Let FormMode(eNewFormMode As pfmPairFormMode)
    SetFormMode eNewFormMode
End Property

Public Property Get FormMode() As pfmPairFormMode
    FormMode = mFormMode
End Property

Public Property Let AutoClearPairsWhenFindingPairs(blnEnable As Boolean)
    mnuFAutoClearPairsWhenFindingPairs.Checked = blnEnable
End Property

Public Property Get AutoClearPairsWhenFindingPairs() As Boolean
    AutoClearPairsWhenFindingPairs = mnuFAutoClearPairsWhenFindingPairs.Checked
End Property

Private Sub ClearAllPairs()
    mPairInfoChanged = True
    DestroyDltLblPairs CallerID
End Sub

Private Sub EnableDisableScanByScanAndIReport(blnEnable As Boolean)
    If cChkBox(chkComputeERScanByScan) <> blnEnable Then
        SetCheckBox chkComputeERScanByScan, blnEnable
    End If
    If cChkBox(chkIReportEREnabled) <> blnEnable Then
        SetCheckBox chkIReportEREnabled, blnEnable
    End If
End Sub

' Unused Procedure (February 2005)
''Private Sub CleanPairsERs()
'''-------------------------------------------
'''this function resets ER in Pairs structure;
'''underlying gel does not change
'''-------------------------------------------
''Dim i As Long
''With GelP_D_L(CallerID)
''    For i = 0 To .PCnt - 1
''        With .Pairs(i)
''            .ER = ER_CALC_ERR
''            .ERStDev = 0
''            .ERChargeStateBasisCount = 0
''            ReDim .ERChargesUsed(0)
''            .ERMemberBasisCount = 0
''        End With
''    Next i
''End With
''End Sub

Private Sub ExcludeAmbiguousPairsWrapper(blnOnlyExaminePairsWithHits As Boolean)
    Dim strMessage As String
    
    If blnOnlyExaminePairsWithHits Then
        strMessage = PairsSearchMarkAmbiguousPairsWithHitsOnly(Me, CallerID)
    Else
        strMessage = PairsSearchMarkAmbiguous(Me, CallerID, True)
    End If
    
    mPairInfoChanged = True
    UpdateStatus strMessage
End Sub

Public Function FindPairsWrapper(Optional blnShowMessages As Boolean = True) As Boolean
    ' Returns True if success, False if error or searching was cancelled prematurely

Dim eResponse As VbMsgBoxResult
Dim blnSuccess As Boolean

On Error GoTo exit_cmdFindPairs

blnSuccess = True
If blnShowMessages Then
    If GelP_D_L(CallerID).PCnt > 0 And Me.AutoClearPairsWhenFindingPairs() Then
        ' Data is already in pairs structure; give user chance to change their mind
        eResponse = MsgBox("Pairs structure already contains pairs. Selected procedure will clear all existing pairs. Continue?", vbOKCancel, glFGTU)
        If eResponse <> vbOK Then
            blnSuccess = False
        End If
    End If
End If
    
If blnSuccess Then
    Me.MousePointer = vbHourglass
    blnSuccess = False
    
    If GelP_D_L(CallerID).PCnt > 0 And Me.AutoClearPairsWhenFindingPairs() Then
        ClearAllPairs
    End If
    
    UpdateStatus "Validating UMC status ..."
    If GelUMC(CallerID).UMCCnt <= 0 Then
       If blnShowMessages Then MsgBox "You must cluster the data into Unique Mass Classes before finding pairs. Please use menu item 'Steps->2. Find UMCs' in the main window to cluster the data into unique mass classes.", vbOKOnly, glFGTU
    Else
        UpdateStatus "Finding pair classes ..."
        blnSuccess = FindPairs(mFormMode, blnShowMessages)
    End If
End If

exit_cmdFindPairs:
Me.MousePointer = vbDefault
FindPairsWrapper = blnSuccess

End Function

Private Function FindPairs(ePairFormMode As pfmPairFormMode, Optional blnShowMessages As Boolean = True) As Boolean
'-----------------------------------------------------
'Delta pairing function; finds and put into structure all potential pairs
' pairs based on numerical criteria
'Returns True if success, False if error or searching was cancelled prematurely
'-----------------------------------------------------
Dim lngIndexLight As Long, lngIndexHeavy As Long
Dim ClsMinDelta As Long
Dim ClsMaxDelta As Long
Dim ClsStepDelta As Long
Dim ClsMidDelta As Long
Dim LClsMW As Double            ' Light member MW
Dim HClsMW As Double            ' Heavy member mw
Dim OverlapOK As Boolean

Dim strSearchMode As String, strMessage As String
Dim blnDeltaInfo As Boolean, blnLabelInfo As Boolean

Dim blnSuccess As Boolean
Dim blnAutoCalculateDeltaMinMax As Boolean

Dim ScanMaxAbuLight As Double
Dim ScanMaxAbuHeavy As Double

On Error GoTo err_FindPairs

Select Case ePairFormMode
Case pfmDelta, pfmLabel, pfmDeltaLabel
    ' All is fine
Case Else
    MsgBox "Unknown pair search mode (sub FindPairsWrapper)", vbExclamation + vbOKOnly, "Error"
    FindPairs = False
    Exit Function
End Select

mPairInfoChanged = True

' Copy current settings to GelP_D_L(Ind)
GelP_D_L(CallerID).SearchDef = glbPreferencesExpanded.PairSearchOptions.SearchDef

mAbortProcess = False

ShowHideControls True

If GelP_D_L(CallerID).PCnt = 0 Or Me.AutoClearPairsWhenFindingPairs() Then
    ' Initially reserve space for 40000 pairs
    blnSuccess = InitDltLblPairs(CallerID)
Else
    blnSuccess = True
End If

If blnSuccess Then
    With GelP_D_L(CallerID)
        Select Case ePairFormMode
        Case pfmDelta
            .DltLblType = ptUMCDlt
            .SearchDef.LightLabelMass = 0
        Case pfmLabel
            .DltLblType = ptUMCLbl
            .SearchDef.DeltaMass = 0
        Case pfmDeltaLabel
            .DltLblType = ptUMCDltLbl
        End Select
         
        .SyncWithUMC = True    'whatever happens we have tried
        
        blnAutoCalculateDeltaMinMax = .SearchDef.AutoCalculateDeltaMinMaxCount
        If Not blnAutoCalculateDeltaMinMax Then
           ClsMinDelta = GelP_D_L(CallerID).SearchDef.DeltaCountMin
           ClsMaxDelta = GelP_D_L(CallerID).SearchDef.DeltaCountMax
        End If
        ClsStepDelta = GelP_D_L(CallerID).SearchDef.DeltaStepSize
    End With
   
    ' Step through the UMC's, treating each lngIndexHeavy'th UMC as the heavy member of the pair
    For lngIndexHeavy = 0 To GelUMC(CallerID).UMCCnt - 1
        HClsMW = GelUMC(CallerID).UMCs(lngIndexHeavy).ClassMW
       
        ' The ClsMinDelta and ClsMaxDelta variables are not used for labeled pairs, but are calculated anyway
        If blnAutoCalculateDeltaMinMax Then         'calculate for this specific mass
           ClsMidDelta = CLng(0.012 * HClsMW)
           ClsMinDelta = CLng(0.5 * ClsMidDelta)
           ClsMaxDelta = CLng(1.5 * ClsMidDelta)
        End If
        UpdateStatus "Examining UMCs: " & Trim(lngIndexHeavy + 1) & " / " & Trim(GelUMC(CallerID).UMCCnt) & "; Pairs found: " & Trim(GelP_D_L(CallerID).PCnt)
        If mAbortProcess Then Exit For
       
        ' Step through the UMC's, treating each lngIndexLight'th UMC as the light member of the pair
        For lngIndexLight = 0 To GelUMC(CallerID).UMCCnt - 1
            If lngIndexHeavy <> lngIndexLight Then
                'check if 'overlap' condition is required and if yes do class lngIndexHeavy and lngIndexLight overlap at the edges?
                OverlapOK = True
                If GelP_D_L(CallerID).SearchDef.RequireUMCOverlap Then
                    If ((GelUMC(CallerID).UMCs(lngIndexLight).MaxScan < GelUMC(CallerID).UMCs(lngIndexHeavy).MinScan) Or _
                       (GelUMC(CallerID).UMCs(lngIndexHeavy).MaxScan < GelUMC(CallerID).UMCs(lngIndexLight).MinScan)) Then
                            OverlapOK = False                       'no overlap
                    End If
                End If
                
                If OverlapOK And GelP_D_L(CallerID).SearchDef.RequireUMCOverlapAtApex Then
                    Select Case GelUMC(CallerID).UMCs(lngIndexHeavy).ClassRepType
                    Case glCSType
                        ScanMaxAbuLight = GelData(CallerID).CSData(GelUMC(CallerID).UMCs(lngIndexLight).ClassRepInd).ScanNumber
                        ScanMaxAbuHeavy = GelData(CallerID).CSData(GelUMC(CallerID).UMCs(lngIndexHeavy).ClassRepInd).ScanNumber
                    Case glIsoType
                        ScanMaxAbuLight = GelData(CallerID).IsoData(GelUMC(CallerID).UMCs(lngIndexLight).ClassRepInd).ScanNumber
                        ScanMaxAbuHeavy = GelData(CallerID).IsoData(GelUMC(CallerID).UMCs(lngIndexHeavy).ClassRepInd).ScanNumber
                    Case Else
                        ' This shouldn't happen
                        Debug.Assert False
                        ScanMaxAbuLight = 0
                        ScanMaxAbuHeavy = GelP_D_L(CallerID).SearchDef.ScanToleranceAtApex + 1
                    End Select
                
                    ' Future: Make this more sophisticated by fitting a Gaussian curve to each of the UMC's
                    '         and comparing the alignment of the fitted curves
                    ' For now, just check the scan distance between the apexes of the two UMCs
                    If Abs(ScanMaxAbuLight - ScanMaxAbuHeavy) > GelP_D_L(CallerID).SearchDef.ScanToleranceAtApex Then
                        OverlapOK = False                       'no overlap
                    End If
                
                End If
             
                If OverlapOK Then
                    'check is it possible that this is a pair
                    LClsMW = GelUMC(CallerID).UMCs(lngIndexLight).ClassMW
                          
                    Select Case ePairFormMode
                    Case pfmDelta
                        FindPairsWorkDelta blnShowMessages, LClsMW, HClsMW, lngIndexLight, lngIndexHeavy, ClsMinDelta, ClsMaxDelta, ClsStepDelta
                    Case pfmLabel
                        FindPairsWorkLabeled blnShowMessages, LClsMW, HClsMW, lngIndexLight, lngIndexHeavy
                    Case pfmDeltaLabel
                        FindPairsWorkDeltaLabeled blnShowMessages, LClsMW, HClsMW, lngIndexLight, lngIndexHeavy, ClsMinDelta, ClsMaxDelta, ClsStepDelta
                    End Select
                
                End If
            End If
        Next lngIndexLight
    Next lngIndexHeavy
   
    Select Case ePairFormMode
    Case pfmDelta
        If Not GelAnalysis(CallerID) Is Nothing Then
            If GelP_D_L(CallerID).SearchDef.DeltaMass = glO16O18_DELTA Then
                GelAnalysis(CallerID).MD_Type = stPairsO16O18
            Else
                ' Use N14/N15 type by default
                GelAnalysis(CallerID).MD_Type = stPairsN14N15
            End If
        End If

        blnDeltaInfo = True
        blnLabelInfo = False
        strSearchMode = "Delta"
        
    Case pfmLabel
        If Not GelAnalysis(CallerID) Is Nothing Then
            GelAnalysis(CallerID).MD_Type = stPairsICAT
        End If
        
        blnDeltaInfo = False
        blnLabelInfo = True
        strSearchMode = "Labeled"
        
    Case pfmDeltaLabel
        If Not GelAnalysis(CallerID) Is Nothing Then
            GelAnalysis(CallerID).MD_Type = stPairsPEON14N15
        End If
        
        blnDeltaInfo = True
        blnLabelInfo = True
        strSearchMode = "Delta-Label"
    Case Else
        strSearchMode = "??"
    End Select
    
    'MonroeMod
    With GelP_D_L(CallerID)
        strMessage = "Searched for " & strSearchMode & " pairs (using UMC's); Pair Count = " & Trim(.PCnt)
        With .SearchDef
            If blnDeltaInfo Then
                strMessage = strMessage & "; Delta = " & Trim(.DeltaMass) & " Da"
                strMessage = strMessage & "; Auto-calculated Min/Max Delta = " & CStr(.AutoCalculateDeltaMinMaxCount)
                strMessage = strMessage & "; Min Deltas = " & Trim(.DeltaCountMin) & "; Max Deltas = " & Trim(.DeltaCountMax)
                strMessage = strMessage & "; Delta Step Size = " & Trim(.DeltaStepSize)
            End If
            
            If blnLabelInfo Then
                strMessage = strMessage & "; Label = " & Trim(.LightLabelMass) & " Da; Heavy/Light Delta = " & Trim(.HeavyLightMassDifference) & " Da"
                strMessage = strMessage & "; Min Labels = " & Trim(.LabelCountMin) & "; Max Labels = " & Trim(.LabelCountMax)
                strMessage = strMessage & "; Max difference in number of labels = " & Trim(.MaxDifferenceInNumberOfLightHeavyLabels)
            End If
            
            strMessage = strMessage & "; Pair Tolerance = " & Trim(.DeltaMassTolerance) & " Da"
        
            strMessage = strMessage & "; Scan Tolerance at Edges = " & Trim(.ScanTolerance) & "; Require Overlap at Edges = " & CStr(.RequireUMCOverlap)
            strMessage = strMessage & "; Scan Tolerance at Apex = " & Trim(.ScanToleranceAtApex) & "; Require Overlap at Apex = " & CStr(.RequireUMCOverlapAtApex)
            strMessage = strMessage & "; ER Inclusion Range = " & Trim(.ERInclusionMin) & " to " & Trim(.ERInclusionMax)
            strMessage = strMessage & "; Require Matching Charge States = " & CStr(.RequireMatchingChargeStatesForPairMembers)
            strMessage = strMessage & "; Use Identical Charges for ER = " & CStr(.UseIdenticalChargesForER)
            strMessage = strMessage & "; Compute ER Scan by Scan = " & CStr(.ComputeERScanByScan)
            strMessage = strMessage & "; Avg ER All Charge States = " & CStr(.AverageERsAllChargeStates)
            strMessage = strMessage & "; Avg ERs Weighting Mode = " & CStr(.AverageERsWeightingMode)
        
        End With
        
        AddToAnalysisHistory CallerID, strMessage
    End With
    
    'calculate expression ratios here (note that GelP_D_L().SearchDef was updated earlier in this function)
    CalcDltLblPairsER_UMC CallerID, strMessage
    
    blnSuccess = True
Else
    strMessage = "Unable to reserve space for pairs structures."
    If blnShowMessages Then
        MsgBox strMessage, vbOKOnly, glFGTU
    Else
        LogErrors -1, "frmUMCDltPairs.FindPairs", strMessage
    End If
    blnSuccess = False
End If

exit_FindPairs:
If GelP_D_L(CallerID).PCnt > 0 Then
    Call TrimDltLblPairs(CallerID)
    UpdateStatus strMessage
Else
    DestroyDltLblPairs CallerID, False
    UpdateStatus "No pairs were found"
End If

exit_Cleanup:
ShowHideControls False
FindPairs = blnSuccess
Exit Function

err_FindPairs:
If blnShowMessages Then
    MsgBox "Error establishing delta pairs " & vbCrLf & "Error: " & Err.Number & ", " & Err.Description, vbOKOnly, glFGTU
Else
    LogErrors Err.Number, "frmUMCDltPairs.FindPairs"
End If

blnSuccess = False
Resume exit_Cleanup

End Function

Private Sub FindPairsWorkDelta(blnShowMessages As Boolean, LClsMW As Double, HClsMW As Double, lngIndexLight As Long, lngIndexHeavy As Long, ClsMinDelta As Long, ClsMaxDelta As Long, ClsStepDelta As Long)
    '--------------------------------------------------------------
    'create all pairs in which class i is the heavy member
    'since this is Delta calculation light member has to be
    'lighter than the heavy member
    '--------------------------------------------------------------

    Dim lngDeltaCnt As Long
    Dim MWDiff As Double
    Dim lngStepSize As Long
    
    If ClsStepDelta > 0 Then
        lngStepSize = ClsStepDelta
    Else
        lngStepSize = 1
    End If
    
    If LClsMW < HClsMW Then
        With GelP_D_L(CallerID)
            For lngDeltaCnt = ClsMinDelta To ClsMaxDelta Step lngStepSize
                MWDiff = HClsMW - (LClsMW + lngDeltaCnt * .SearchDef.DeltaMass)
                If Abs(MWDiff) <= .SearchDef.DeltaMassTolerance Then
                    FindPairsWorkValidatePair blnShowMessages, lngIndexLight, lngIndexHeavy, lngDeltaCnt, 0, 0
                End If
                If MWDiff < -.SearchDef.DeltaMassTolerance Or mAbortProcess Then
                    Exit For
                End If
            Next lngDeltaCnt
        End With
    End If

End Sub

Private Sub FindPairsWorkLabeled(blnShowMessages As Boolean, LClsMW As Double, HClsMW As Double, lngIndexLight As Long, lngIndexHeavy As Long)
    '----------------------------------------------------------
    'this is a little wicked situation in which heavy and light
    'members do not have to have same number of labels attached
    'and that can cause light member to be heavier than heavy
    '----------------------------------------------------------

    Dim LblCntHvy As Long
    Dim LblCntLgt As Long
    Dim MWDiff As Double
    
    MWDiff = HClsMW - LClsMW
    With GelP_D_L(CallerID)
        If .SearchDef.MaxDifferenceInNumberOfLightHeavyLabels > 0 Then
            ' Label count could differ
            For LblCntHvy = .SearchDef.LabelCountMin To .SearchDef.LabelCountMax
                If (HClsMW - LblCntHvy * (.SearchDef.LightLabelMass + .SearchDef.HeavyLightMassDifference) > 0) Then   'don't consider impossible pairs
                    For LblCntLgt = .SearchDef.LabelCountMin To .SearchDef.LabelCountMax
                        If (LClsMW - LblCntLgt * .SearchDef.LightLabelMass > 0) Then         'don't consider impossible pairs
                            If Abs(LblCntHvy - LblCntLgt) <= .SearchDef.MaxDifferenceInNumberOfLightHeavyLabels And (LblCntHvy + LblCntLgt) > 0 Then
                                If Abs(MWDiff - ((LblCntHvy - LblCntLgt) * .SearchDef.LightLabelMass + LblCntHvy * .SearchDef.HeavyLightMassDifference)) <= .SearchDef.DeltaMassTolerance Then
                                    FindPairsWorkValidatePair blnShowMessages, lngIndexLight, lngIndexHeavy, 0, LblCntLgt, LblCntHvy
                                End If
                            End If
                        End If
                    Next LblCntLgt
                End If
                If mAbortProcess Then Exit For
            Next LblCntHvy
        ElseIf .SearchDef.MaxDifferenceInNumberOfLightHeavyLabels = 0 Then
            ' Label count has to be the same (LblCntHvy = LblCntLgt)
            For LblCntHvy = .SearchDef.LabelCountMin To .SearchDef.LabelCountMax
                If (LClsMW - LblCntHvy * .SearchDef.LightLabelMass > 0) Then       'don't consider impossible pairs
                    If (HClsMW - LblCntHvy * (.SearchDef.LightLabelMass + .SearchDef.HeavyLightMassDifference) > 0) Then
                        If Abs(MWDiff - LblCntHvy * .SearchDef.HeavyLightMassDifference) <= .SearchDef.DeltaMassTolerance Then
                            FindPairsWorkValidatePair blnShowMessages, lngIndexLight, lngIndexHeavy, 0, LblCntHvy, LblCntHvy
                        End If
                    End If
                End If
                If mAbortProcess Then Exit For
            Next LblCntHvy
        Else
            ' Invalid label count
            Debug.Assert False
        End If
    End With
    
End Sub

Private Sub FindPairsWorkDeltaLabeled(blnShowMessages As Boolean, LClsMW As Double, HClsMW As Double, lngIndexLight As Long, lngIndexHeavy As Long, ClsMinDelta As Long, ClsMaxDelta As Long, ClsStepDelta As Long)
    '--------------------------------------------------------------
    'create all pairs in which class i is the heavy member
    'since this is Delta calculation light member has to be
    'lighter than the heavy member
    '
    ' WARNING: This algorithm is not functional
    '--------------------------------------------------------------

    Dim lngDeltaCnt As Long
    Dim MWDiff As Double
    Dim lngStepSize As Long
    
    If ClsStepDelta > 0 Then
        lngStepSize = ClsStepDelta
    Else
        lngStepSize = 1
    End If
    
    If LClsMW < HClsMW Then
        With GelP_D_L(CallerID)
            For lngDeltaCnt = ClsMinDelta To ClsMaxDelta Step lngStepSize
                MWDiff = HClsMW - (LClsMW + lngDeltaCnt * .SearchDef.DeltaMass)
                If Abs(MWDiff) <= .SearchDef.DeltaMassTolerance Then
                    FindPairsWorkValidatePair blnShowMessages, lngIndexLight, lngIndexHeavy, lngDeltaCnt, 0, 0
                End If
                If MWDiff < -.SearchDef.DeltaMassTolerance Or mAbortProcess Then
                    Exit For
                End If
            Next lngDeltaCnt
        End With
    End If
End Sub

Private Sub FindPairsWorkValidatePair(blnShowMessages As Boolean, lngIndexLight As Long, lngIndexHeavy As Long, DeltaCnt As Long, LblCntLgt As Long, LblCntHvy As Long)
    Dim blnPairOK As Boolean
    
    With GelP_D_L(CallerID)
        If .SearchDef.RequireUMCOverlap Then
            ' See if pairs overlap at the edges within Scan Tolerance
            blnPairOK = PairsOverlapAtEdgesWithinTol(CallerID, lngIndexLight, lngIndexHeavy, .SearchDef.ScanTolerance)
        Else
            blnPairOK = True
        End If
        
        If blnPairOK Then
            If .SearchDef.RequireMatchingChargeStatesForPairMembers Then
                ' See if pairs contain matching charge states
                blnPairOK = ChargeStatesMatch(CallerID, lngIndexLight, lngIndexHeavy)
            End If
            
            If blnPairOK Then
                If ValidatePairArraySpace(blnShowMessages) Then
                    .PCnt = .PCnt + 1
                    With .Pairs(.PCnt - 1)
                        .P1 = lngIndexLight
                        .P2 = lngIndexHeavy
                        .P2DltCnt = DeltaCnt
                        .P1LblCnt = LblCntLgt
                        .P2LblCnt = LblCntHvy
                    End With
                Else
                    ' Memory management error
                    Debug.Assert False
                    mAbortProcess = True
                End If
            End If
        End If
    End With
    
End Sub

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
  If .PCnt >= 0 Then
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
  End If
  
  If .PCnt > 0 Then
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
        If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            ' Even though no pairs exist, pretend they do, so that the report text file will be created anyway
            GenerateERStat = True
        Else
            MsgBox "No pairs found. Find Pairs function should be used first."
        End If
    End If
End With
End Function

Public Sub InitializeForm()
    On Error GoTo InitializeFormErrorHandler
    
    If bLoading Then
        CallerID = Me.Tag
        
        If CallerID >= 1 And CallerID <= UBound(GelBody) Then
            glbPreferencesExpanded.PairSearchOptions.SearchDef = GelP_D_L(CallerID).SearchDef
        End If
       
        With glbPreferencesExpanded.PairSearchOptions
            SetCheckBox chkOutlierRemovalUsesSymmetricERs, .OutlierRemovalUsesSymmetricERs
        End With
        
        With glbPreferencesExpanded.PairSearchOptions.SearchDef
            txtDelta.Text = .DeltaMass
            
            SetCheckBox chkAutoMinMaxDelta, .AutoCalculateDeltaMinMaxCount
            txtMinDelta.Text = .DeltaCountMin
            txtMaxDelta.Text = .DeltaCountMax
            txtDeltaStepSize.Text = .DeltaStepSize
            txtDeltaStepSize.ToolTipText = "If this value is greater than 1, then allowable deltas must be the given step size away from the minimum delta count.  For example, if it is 4, and Min Deltas is 4, then the allowed delta counts are 4, 8, 12, etc."
            
            txtLabel = .LightLabelMass
            txtHeavyLightDelta = .HeavyLightMassDifference
            txtMinLbl = .LabelCountMin
            txtMaxLbl = .LabelCountMax
            txtMaxLblDiff = .MaxDifferenceInNumberOfLightHeavyLabels
            
            txtPairTol.Text = .DeltaMassTolerance
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
            
            SetCheckBox chkIReportEREnabled, .IReportEROptions.Enabled
            
        
            SetCheckBox chkRemoveOutlierERs, .RemoveOutlierERs
            SetCheckBox chkRemoveOutlierERsIterate, .RemoveOutlierERsIterate
            
            If .RemoveOutlierERsMinimumDataPointCount < 2 Then .RemoveOutlierERsMinimumDataPointCount = 2
            txtRemoveOutlierERsMinimumDataPointCount.Text = .RemoveOutlierERsMinimumDataPointCount
            
            UpdateDynamicControls
        End With
        
       bLoading = False
       
    End If
    Exit Sub

InitializeFormErrorHandler:
    LogErrors Err.Number, "frmUMCDltPairs->InitializeForm", Err.Description
    Resume Next
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
If mFormMode <> pfmLabel Then
    Call txtDelta_LostFocus
    Call txtMinDelta_LostFocus
    Call txtMaxDelta_LostFocus
    Call txtDeltaStepSize_LostFocus
End If

If mFormMode <> pfmDelta Then
    Call txtLabel_LostFocus
    Call txtHeavyLightDelta_LostFocus
    Call txtMinLbl_LostFocus
    Call txtMaxLbl_LostFocus
    Call txtMaxLblDiff_LostFocus
End If

Call txtPairTol_LostFocus
Call txtPairsScanTolEdge_LostFocus
Call txtPairsScanTolApex_LostFocus
Call txtERMin_LostFocus
Call txtERMax_LostFocus

Call txtRemoveOutlierERsMinimumDataPointCount_LostFocus
End Sub

Public Sub ReportPairs(PState As Integer, Optional strFilePath As String = "")
' PState can be 0 for all pairs, 1 for Included only (aka glPAIR_Inc), or
'   -1 for Excluded only (aka glPAIR_Exc)
' If Len(strFilePath) = 0, then displays report using frmDataInfo;
'   otherwise, saves the report to strFilePath

UpdateStatus "Generating report ..."
Me.MousePointer = vbHourglass
ReportDltLblPairsUMCWrapper CallerID, PState, strFilePath
Me.MousePointer = vbDefault
UpdateStatus ""

End Sub

Public Sub ReportERStatistics(Optional strFilePath As String = "")
' If Len(strFilePath) = 0, then displays report using frmDataInfo;
'   otherwise, saves the report to strFilePath

UpdateStatus "Generating report ..."
Me.MousePointer = vbHourglass
If GenerateERStat Then
   ReportERStat CallerID, ERBin(), ERBinAll(), ERBinInc(), _
                ERBinExc(), ERAllS, ERIncS, ERExcS, strFilePath
End If
Me.MousePointer = vbDefault
UpdateStatus ""

End Sub

Private Sub ResetToDefaults()
    ResetExpandedPreferences glbPreferencesExpanded, "PairSearchOptions"
    GelP_D_L(CallerID).SearchDef = glbPreferencesExpanded.PairSearchOptions.SearchDef
    
    bLoading = True
    InitializeForm
End Sub

Public Sub SetDeltaMass(dblDeltaMass As Double)
    txtDelta.Text = dblDeltaMass
    glbPreferencesExpanded.PairSearchOptions.SearchDef.DeltaMass = dblDeltaMass
End Sub

Private Sub SetFormMode(ByVal eNewFormMode As pfmPairFormMode)
    
    Const CONTROL_SPACING = 100
    
    Dim lngToleranceFormTop As Long
    
    mFormMode = eNewFormMode
    
    Select Case eNewFormMode
    Case pfmDelta, pfmLabel
        With fraLabelOptions
            .Left = fraDeltaOptions.Left
            .Top = fraDeltaOptions.Top
        End With
        
        If eNewFormMode = pfmDelta Then
            ' Delta pairs
            fraDeltaOptions.Visible = True
            fraLabelOptions.Visible = False
            lngToleranceFormTop = fraDeltaOptions.Top + fraDeltaOptions.Height + CONTROL_SPACING
            Me.Caption = "UMC Delta Pairing Analysis"
            
            If txtDelta = "0" Then cmdSetToO18_Click
        Else
            ' Label pairs
            fraDeltaOptions.Visible = False
            fraLabelOptions.Visible = True
            lngToleranceFormTop = fraLabelOptions.Top + fraLabelOptions.Height + CONTROL_SPACING
            Me.Caption = "UMC Labeled Pairing Analysis"
            
            If txtLabel = "0" Then cmdSetToICAT_Click
        End If
        
    Case pfmDeltaLabel
        With fraLabelOptions
            .Left = fraDeltaOptions.Left
            .Top = fraDeltaOptions.Top + fraDeltaOptions.Height + CONTROL_SPACING
        End With
        
        fraDeltaOptions.Visible = True
        fraLabelOptions.Visible = True
        lngToleranceFormTop = fraLabelOptions.Top + fraLabelOptions.Height + CONTROL_SPACING
        Me.Caption = "UMC N14/N15 Cys-based Labeled Pairing"
        
        MsgBox "Warning, the Delta/Label Pairs search algorithm is under development and has not been fully implemented.  In particular, the settings for Label mass and Heavy/Light Delta mass are not utilized in the search (MEM - July 2004).", vbInformation + vbOKOnly, "Warning"
        
    Case Else
        ' Unknown mode
        Debug.Assert False
        SetFormMode pfmDelta
    End Select
    
    With fraToleranceOptions
        .Left = fraDeltaOptions.Left
        .Top = lngToleranceFormTop
    End With
    
    With fraInclusionOptions
        .Left = fraDeltaOptions.Left
        .Top = fraToleranceOptions.Top + fraToleranceOptions.Height + CONTROL_SPACING
    End With
    
    With fraGeneralOptions
        .Left = fraDeltaOptions.Left
        .Top = fraInclusionOptions.Top + fraInclusionOptions.Height + CONTROL_SPACING
    End With
    
    With lblStatus
        .Left = fraDeltaOptions.Left
        .Top = fraGeneralOptions.Top + fraGeneralOptions.Height + CONTROL_SPACING
    End With
    
    Me.Height = lblStatus.Top + lblStatus.Height + 740
    Me.width = 5650
End Sub

Private Sub SetPairSearchDeltas(dblDeltaMass As Double, DeltaCountMin As Long, DeltaCountMax As Long, Optional DeltaStepSize As Long = 1)
    txtDelta = dblDeltaMass
    txtMinDelta = DeltaCountMin
    txtMaxDelta = DeltaCountMax
    txtDeltaStepSize = DeltaStepSize
    
    With glbPreferencesExpanded.PairSearchOptions.SearchDef
        .DeltaMass = dblDeltaMass
        .DeltaCountMin = DeltaCountMin
        .DeltaCountMax = DeltaCountMax
        .DeltaStepSize = DeltaStepSize
    End With

End Sub

Private Sub SetPairSearchLabel(dblLightLabelMass As Double, dblHeavyLightDelta As Double, LabelCountMin As Long, LabelCountMax As Long, MaxLblCountDiff As Long)
    txtLabel = dblLightLabelMass
    txtHeavyLightDelta = dblHeavyLightDelta
    txtMinLbl = LabelCountMin
    txtMaxLbl = LabelCountMax
    txtMaxLblDiff = MaxLblCountDiff
    
    With glbPreferencesExpanded.PairSearchOptions.SearchDef
        .LightLabelMass = dblLightLabelMass
        .HeavyLightMassDifference = dblHeavyLightDelta
        .LabelCountMin = LabelCountMin
        .LabelCountMax = LabelCountMax
        .MaxDifferenceInNumberOfLightHeavyLabels = MaxLblCountDiff
    End With
End Sub

Private Sub ShowHideControls(blnSearchingForPairs As Boolean)
    cmdAbortProcess.Visible = blnSearchingForPairs
    cmdFindPairs.Visible = Not blnSearchingForPairs
    fraControls.Visible = Not blnSearchingForPairs
    cmdSetToICAT.Visible = Not blnSearchingForPairs
End Sub

Private Sub UpdateDynamicControls()
    Dim blnEnableOutlierControls As Boolean
    
    With glbPreferencesExpanded.PairSearchOptions.SearchDef
        chkAverageERsAllChargeStates.Enabled = (.UseIdenticalChargesForER And .RequireMatchingChargeStatesForPairMembers)
        cboAverageERsWeightingMode.Enabled = (.AverageERsAllChargeStates And chkAverageERsAllChargeStates.Enabled)
    
        chkRemoveOutlierERs.Enabled = .ComputeERScanByScan
        
        blnEnableOutlierControls = .ComputeERScanByScan And .RemoveOutlierERs
        chkRemoveOutlierERsIterate.Enabled = blnEnableOutlierControls
        txtRemoveOutlierERsMinimumDataPointCount.Enabled = blnEnableOutlierControls
        chkOutlierRemovalUsesSymmetricERs.Enabled = blnEnableOutlierControls
        lblRemoveOutlierERsMinimumDataPointCount.Enabled = blnEnableOutlierControls
    End With
    
    On Error Resume Next
    If (GelData(CallerID).DataStatusBits And GEL_DATA_STATUS_BIT_IREPORT) = GEL_DATA_STATUS_BIT_IREPORT Then
        chkIReportEREnabled.Enabled = glbPreferencesExpanded.PairSearchOptions.SearchDef.ComputeERScanByScan
    Else
        chkIReportEREnabled.Enabled = False
    End If
    
End Sub

Private Sub UpdateERCalculationOptions()

    With glbPreferencesExpanded.PairSearchOptions.SearchDef
        GelP_D_L(CallerID).SearchDef.RequireMatchingChargeStatesForPairMembers = .RequireMatchingChargeStatesForPairMembers
        GelP_D_L(CallerID).SearchDef.UseIdenticalChargesForER = .UseIdenticalChargesForER
        GelP_D_L(CallerID).SearchDef.ComputeERScanByScan = .ComputeERScanByScan
        GelP_D_L(CallerID).SearchDef.AverageERsAllChargeStates = .AverageERsAllChargeStates
        GelP_D_L(CallerID).SearchDef.IReportEROptions.Enabled = .IReportEROptions.Enabled
        GelP_D_L(CallerID).SearchDef.RemoveOutlierERs = .RemoveOutlierERs
        GelP_D_L(CallerID).SearchDef.RemoveOutlierERsIterate = .RemoveOutlierERsIterate
        GelP_D_L(CallerID).SearchDef.RemoveOutlierERsMinimumDataPointCount = .RemoveOutlierERsMinimumDataPointCount
        GelP_D_L(CallerID).SearchDef.RemoveOutlierERsConfidenceLevel = .RemoveOutlierERsConfidenceLevel
    End With
     
End Sub

Private Sub UpdateStatus(ByVal Status As String)
'-----------------------------------------------
'set status label; entertain user so it doesn't
'freak out before function finishes
'-----------------------------------------------
lblStatus.Caption = Status
DoEvents
End Sub

Private Function ValidatePairArraySpace(blnShowMessages As Boolean) As Boolean
    Dim blnContinue As Boolean
    Dim strMessage As String
    
    If GelP_D_L(CallerID).PCnt + 1 > UBound(GelP_D_L(CallerID).Pairs) Then
        If UBound(GelP_D_L(CallerID).Pairs) > MAXPAIRS Then
           strMessage = "Number of detected pairs too high. (max. number of pairs " & MAXPAIRS & ")"
           If blnShowMessages Then
              MsgBox strMessage, vbOKOnly
           Else
              AddToAnalysisHistory CallerID, strMessage
           End If
           blnContinue = False
        Else
           blnContinue = AddDltLblPairs(CallerID, 10000)
        End If
    Else
        blnContinue = True
    End If
    
    ValidatePairArraySpace = blnContinue
    
End Function

Private Sub cboAverageERsWeightingMode_Click()
    glbPreferencesExpanded.PairSearchOptions.SearchDef.AverageERsWeightingMode = cboAverageERsWeightingMode.ListIndex
End Sub

Private Sub chkAutoMinMaxDelta_Click()
If chkAutoMinMaxDelta.Value = vbChecked Then
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
    UpdateDynamicControls
End Sub

Private Sub chkIReportEREnabled_Click()
    glbPreferencesExpanded.PairSearchOptions.SearchDef.IReportEROptions.Enabled = cChkBox(chkIReportEREnabled)
End Sub

Private Sub chkOutlierRemovalUsesSymmetricERs_Click()
    glbPreferencesExpanded.PairSearchOptions.OutlierRemovalUsesSymmetricERs = cChkBox(chkOutlierRemovalUsesSymmetricERs)
End Sub

Private Sub chkPairsExcludeAmbiguousKeepMostConfident_Click()
    glbPreferencesExpanded.PairSearchOptions.KeepMostConfidentAmbiguous = cChkBox(chkPairsExcludeAmbiguousKeepMostConfident)
End Sub

Private Sub chkPairsRequireOverlapAtApex_Click()
    glbPreferencesExpanded.PairSearchOptions.SearchDef.RequireUMCOverlapAtApex = cChkBox(chkPairsRequireOverlapAtApex)
End Sub

Private Sub chkPairsRequireOverlapAtEdge_Click()
    glbPreferencesExpanded.PairSearchOptions.SearchDef.RequireUMCOverlap = cChkBox(chkPairsRequireOverlapAtEdge)
End Sub

Private Sub chkRemoveOutlierERs_Click()
    glbPreferencesExpanded.PairSearchOptions.SearchDef.RemoveOutlierERs = cChkBox(chkRemoveOutlierERs)
    UpdateDynamicControls
End Sub

Private Sub chkRemoveOutlierERsIterate_Click()
    glbPreferencesExpanded.PairSearchOptions.SearchDef.RemoveOutlierERsIterate = cChkBox(chkRemoveOutlierERsIterate)
End Sub

Private Sub chkRequireMatchingChargeStates_Click()
    glbPreferencesExpanded.PairSearchOptions.SearchDef.RequireMatchingChargeStatesForPairMembers = cChkBox(chkRequireMatchingChargeStates)
    UpdateDynamicControls
End Sub

Private Sub chkUseIdenticalChargeStatesForER_Click()
    glbPreferencesExpanded.PairSearchOptions.SearchDef.UseIdenticalChargesForER = cChkBox(chkUseIdenticalChargeStatesForER)
    UpdateDynamicControls
End Sub

Private Sub cmdAbortProcess_Click()
    mAbortProcess = True
End Sub

Private Sub cmdFindPairs_Click()
    FindPairsWrapper
End Sub

Private Sub cmdResetToDefaults_Click(Index As Integer)
    ResetToDefaults
End Sub

Private Sub cmdSetDeuterium_Click()
    SetPairSearchDeltas glDeuterium_DELTA, 3, 15, 3
    EnableDisableScanByScanAndIReport False
End Sub

Private Sub cmdSetToC13_Click()
    SetPairSearchDeltas glC12C13_DELTA, 1, 100
    EnableDisableScanByScanAndIReport False
End Sub

Private Sub cmdSetToICAT_Click()
    SetPairSearchLabel glICAT_Light, glICAT_Delta, 1, 5, 1
    EnableDisableScanByScanAndIReport False
End Sub

Private Sub cmdSetToN15_Click()
    SetPairSearchDeltas glN14N15_DELTA, 1, 100
    EnableDisableScanByScanAndIReport False
End Sub

Private Sub cmdSetToO18_Click()
    SetPairSearchDeltas glO16O18_DELTA, 1, 1
    EnableDisableScanByScanAndIReport True
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

AutoClearPairsWhenFindingPairs = True

mPairInfoChanged = False
ShowHideControls False

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

Private Sub mnuFAutoClearPairsWhenFindingPairs_Click()
    Me.AutoClearPairsWhenFindingPairs = Not Me.AutoClearPairsWhenFindingPairs()
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
Dim strMessage As String

UpdateStatus "Deleting excluded pairs ..."
Me.MousePointer = vbHourglass

strMessage = DeleteExcludedPairs(CallerID)
UpdateStatus strMessage
AddToAnalysisHistory CallerID, strMessage

mPairInfoChanged = True
Me.MousePointer = vbDefault
End Sub

Private Sub mnuFERRecalculation_Click()
Dim strMessage As String

UpdateERCalculationOptions
CalcDltLblPairsER_UMC CallerID, strMessage
UpdateStatus strMessage

mPairInfoChanged = True

End Sub

Private Sub mnuFFindPairs_Click()
    FindPairsWrapper
End Sub

Private Sub mnuFMarkAmbPairs_Click()
    ExcludeAmbiguousPairsWrapper False
End Sub

Private Sub mnuFMarkAmbPairsHitsOnly_Click()
    ExcludeAmbiguousPairsWrapper True
End Sub

Private Sub mnuFMarkBadERPairs_Click()
'-----------------------------------------------------------------
'mark pair as bad if expression ratio is out of ER inclusion range
'-----------------------------------------------------------------
MarkBadERPairs
End Sub

Private Sub mnuFResetExclusionFlags_Click()
UpdateStatus "Resetting pair exclusion flags..."
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
MsgBox "Delta should be positive number.", vbOKOnly, glFGTU
txtDelta.SetFocus
End Sub

Private Sub txtDeltaStepSize_LostFocus()
On Error GoTo err_DeltaStep
glbPreferencesExpanded.PairSearchOptions.SearchDef.DeltaStepSize = CLng(txtDeltaStepSize.Text)
If glbPreferencesExpanded.PairSearchOptions.SearchDef.DeltaStepSize >= 0 Then Exit Sub
Exit Sub
err_DeltaStep:
MsgBox "Delta step size should be 0 or greater.", vbOKOnly, glFGTU
txtDeltaStepSize.SetFocus
End Sub

Private Sub txtHeavyLightDelta_LostFocus()
On Error GoTo err_Delta
glbPreferencesExpanded.PairSearchOptions.SearchDef.HeavyLightMassDifference = CDbl(txtHeavyLightDelta.Text)
If glbPreferencesExpanded.PairSearchOptions.SearchDef.HeavyLightMassDifference > 0 Then Exit Sub
err_Delta:
MsgBox "Delta should be positive number.", vbOKOnly, glFGTU
txtHeavyLightDelta.SetFocus
End Sub

Private Sub txtLabel_LostFocus()
On Error GoTo err_Label
glbPreferencesExpanded.PairSearchOptions.SearchDef.LightLabelMass = CDbl(txtLabel.Text)
If glbPreferencesExpanded.PairSearchOptions.SearchDef.LightLabelMass >= 0 Then Exit Sub
err_Label:
MsgBox "Label mass should be non-negative number.", vbOKOnly, glFGTU
txtLabel.SetFocus
End Sub

Private Sub txtMaxLbl_LostFocus()
On Error GoTo err_MaxLbl
glbPreferencesExpanded.PairSearchOptions.SearchDef.LabelCountMax = CLng(txtMaxLbl.Text)
If glbPreferencesExpanded.PairSearchOptions.SearchDef.LabelCountMax > 0 Then Exit Sub
err_MaxLbl:
MsgBox "Maximum number of labels should be non-negative integer.", vbOKOnly, glFGTU
txtMaxLbl.SetFocus
End Sub

Private Sub txtMaxLblDiff_LostFocus()
On Error GoTo err_MaxLblDiff
glbPreferencesExpanded.PairSearchOptions.SearchDef.MaxDifferenceInNumberOfLightHeavyLabels = CLng(txtMaxLblDiff.Text)
If glbPreferencesExpanded.PairSearchOptions.SearchDef.MaxDifferenceInNumberOfLightHeavyLabels >= 0 Then Exit Sub
err_MaxLblDiff:
MsgBox "Maximum difference between number of light and heavy labels should be non-negative integer.", vbOKOnly, glFGTU
txtMaxLblDiff.SetFocus
End Sub

Private Sub txtMinLbl_LostFocus()
On Error GoTo err_MinLbl
glbPreferencesExpanded.PairSearchOptions.SearchDef.LabelCountMin = CLng(txtMinLbl.Text)
If glbPreferencesExpanded.PairSearchOptions.SearchDef.LabelCountMin >= 0 Then Exit Sub
err_MinLbl:
MsgBox "Minimum number of labels should be non-negative integer.", vbOKOnly, glFGTU
txtMinLbl.SetFocus
End Sub

Private Sub txtPairTol_LostFocus()
On Error GoTo err_DeltaTol
glbPreferencesExpanded.PairSearchOptions.SearchDef.DeltaMassTolerance = CDbl(txtPairTol.Text)
If glbPreferencesExpanded.PairSearchOptions.SearchDef.DeltaMassTolerance > 0 Then Exit Sub
err_DeltaTol:
MsgBox "Delta tolerance should be positive number.", vbOKOnly, glFGTU
txtPairTol.SetFocus
End Sub

Private Sub txtERMax_LostFocus()
On Error GoTo err_ERMax
glbPreferencesExpanded.PairSearchOptions.SearchDef.ERInclusionMax = CDbl(txtERMax.Text)
Exit Sub
err_ERMax:
MsgBox "Maximum of ER range should be a number.", vbOKOnly, glFGTU
txtERMax.SetFocus
End Sub

Private Sub txtERMin_LostFocus()
On Error GoTo err_ERMin
glbPreferencesExpanded.PairSearchOptions.SearchDef.ERInclusionMin = CDbl(txtERMin.Text)
Exit Sub
err_ERMin:
MsgBox "Minimum of ER range should be a number.", vbOKOnly, glFGTU
txtERMin.SetFocus
End Sub

Private Sub txtMaxDelta_LostFocus()
On Error GoTo err_MaxDelta
glbPreferencesExpanded.PairSearchOptions.SearchDef.DeltaCountMax = CLng(txtMaxDelta.Text)
If glbPreferencesExpanded.PairSearchOptions.SearchDef.DeltaCountMax > 0 Then Exit Sub
err_MaxDelta:
MsgBox "Maximum Delta should be positive integer.", vbOKOnly, glFGTU
txtMaxDelta.SetFocus
End Sub

Private Sub txtMinDelta_LostFocus()
On Error GoTo err_MinDelta
glbPreferencesExpanded.PairSearchOptions.SearchDef.DeltaCountMin = CLng(txtMinDelta.Text)
If glbPreferencesExpanded.PairSearchOptions.SearchDef.DeltaCountMin > 0 Then Exit Sub
err_MinDelta:
MsgBox "Minimum Delta should be positive integer.", vbOKOnly, glFGTU
txtMinDelta.SetFocus
End Sub

Private Sub txtPairsScanTolApex_LostFocus()
    On Error GoTo err_ScanTol
    glbPreferencesExpanded.PairSearchOptions.SearchDef.ScanToleranceAtApex = CLng(txtPairsScanTolApex.Text)
    If glbPreferencesExpanded.PairSearchOptions.SearchDef.ScanToleranceAtApex >= 0 Then Exit Sub
err_ScanTol:
    MsgBox "Scan tolerance should be non-negative integer.", vbOKOnly, glFGTU
    txtPairsScanTolApex.SetFocus
End Sub

Private Sub txtPairsScanTolEdge_LostFocus()
    On Error GoTo err_ScanTol
    glbPreferencesExpanded.PairSearchOptions.SearchDef.ScanTolerance = CLng(txtPairsScanTolEdge.Text)
    If glbPreferencesExpanded.PairSearchOptions.SearchDef.ScanTolerance >= 0 Then Exit Sub
err_ScanTol:
    MsgBox "Scan tolerance should be non-negative integer.", vbOKOnly, glFGTU
    txtPairsScanTolEdge.SetFocus
End Sub

Private Sub txtRemoveOutlierERsMinimumDataPointCount_LostFocus()
    On Error GoTo err_RemoveOutliers
    With glbPreferencesExpanded.PairSearchOptions.SearchDef
        .RemoveOutlierERsMinimumDataPointCount = CLng(txtRemoveOutlierERsMinimumDataPointCount)
        If .RemoveOutlierERsMinimumDataPointCount < 2 Then
            .RemoveOutlierERsMinimumDataPointCount = 2
            txtRemoveOutlierERsMinimumDataPointCount.Text = Trim(.RemoveOutlierERsMinimumDataPointCount)
        End If
    End With
    If glbPreferencesExpanded.PairSearchOptions.SearchDef.RemoveOutlierERsMinimumDataPointCount >= 2 Then Exit Sub
err_RemoveOutliers:
    MsgBox "Minimum final number of data points should be at least 2", vbOKOnly, glFGTU
    txtRemoveOutlierERsMinimumDataPointCount.SetFocus
End Sub
