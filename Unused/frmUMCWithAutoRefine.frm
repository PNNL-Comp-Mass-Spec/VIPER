VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmUMCWithAutoRefine 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Unique Molecular Mass Classes Definition (NOTE: Obsolete Method)"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6900
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab tbsTabStrip 
      Height          =   4875
      Left            =   2520
      TabIndex        =   15
      Top             =   480
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   8599
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "UMC Definition"
      TabPicture(0)   =   "frmUMCWithAutoRefine.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cboChargeStateAbuType"
      Tab(0).Control(1)=   "chkUseMostAbuChargeStateStatsForClassStats"
      Tab(0).Control(2)=   "cmbCountType"
      Tab(0).Control(3)=   "txtHoleNum"
      Tab(0).Control(4)=   "txtHoleSize"
      Tab(0).Control(5)=   "txtHolePct"
      Tab(0).Control(6)=   "cmbUMCAbu"
      Tab(0).Control(7)=   "cmbUMCMW"
      Tab(0).Control(8)=   "lblChargeStateAbuType"
      Tab(0).Control(9)=   "Label1(0)"
      Tab(0).Control(10)=   "Label3(0)"
      Tab(0).Control(11)=   "Label3(1)"
      Tab(0).Control(12)=   "Label3(2)"
      Tab(0).Control(13)=   "Label1(1)"
      Tab(0).Control(14)=   "Label1(2)"
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Auto Refine Options"
      TabPicture(1)   =   "frmUMCWithAutoRefine.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblAutoRefineLengthLabel(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblAutoRefineLengthLabel(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblAutoRefineMinimumMemberCount"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtHiCnt"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "chkRemoveHiCnt"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtLoCnt"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "chkRemoveLoCnt"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtHiAbuPct"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "chkRemoveHiAbu"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtLoAbuPct"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "chkRemoveLoAbu"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "fraSplitUMCsOptions"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "chkRefineUMCLengthByScanRange"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtAutoRefineMinimumMemberCount"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      Begin VB.ComboBox cboChargeStateAbuType 
         Height          =   315
         Left            =   -74760
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2520
         Width           =   3135
      End
      Begin VB.CheckBox chkUseMostAbuChargeStateStatsForClassStats 
         Caption         =   "Use most abundant charge state group stats for class stats"
         Height          =   405
         Left            =   -74760
         TabIndex        =   24
         ToolTipText     =   "Make single-member classes from unconnected nodes"
         Top             =   2880
         Width           =   2535
      End
      Begin VB.TextBox txtAutoRefineMinimumMemberCount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3360
         TabIndex        =   59
         Text            =   "3"
         Top             =   2000
         Width           =   495
      End
      Begin VB.CheckBox chkRefineUMCLengthByScanRange 
         Caption         =   "Test UMC length using scan range"
         Height          =   375
         Left            =   120
         TabIndex        =   58
         ToolTipText     =   "If True, then considers scan range for the length tests; otherwise, considers member count"
         Top             =   1900
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.Frame fraSplitUMCsOptions 
         Caption         =   "Split UMC's Options"
         Height          =   2400
         Left            =   120
         TabIndex        =   40
         Top             =   2340
         Width           =   3800
         Begin VB.TextBox txtSplitUMCsPeakPickingMinimumWidth 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2400
            TabIndex        =   52
            Text            =   "4"
            Top             =   1980
            Width           =   495
         End
         Begin VB.TextBox txtSplitUMCsPeakDetectIntensityThresholdPercentageOfMax 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2400
            TabIndex        =   49
            Text            =   "15"
            Top             =   1620
            Width           =   495
         End
         Begin VB.TextBox txtSplitUMCsMaximumPeakCount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2400
            TabIndex        =   46
            Text            =   "6"
            Top             =   1140
            Width           =   495
         End
         Begin VB.TextBox txtSplitUMCsMinimumDifferenceInAvgPpmMass 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2400
            TabIndex        =   43
            Text            =   "4"
            Top             =   660
            Width           =   495
         End
         Begin VB.CheckBox chkSplitUMCsByExaminingAbundance 
            Caption         =   "Split UMC's by Examining Abundance"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label lblSplitUMCsPeakPickingMinimumWidth 
            Caption         =   "Peak picking minimum width"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   2010
            Width           =   2295
         End
         Begin VB.Label lblUnits 
            Caption         =   "scans"
            Height          =   255
            Index           =   5
            Left            =   3000
            TabIndex        =   53
            Top             =   2010
            Width           =   735
         End
         Begin VB.Label lblSplitUMCsPeakDetectIntensityThresholdPercentageOfMax 
            Caption         =   "Peak picking intensity threshold"
            Height          =   405
            Left            =   120
            TabIndex        =   48
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label lblUnits 
            Caption         =   "% of max"
            Height          =   255
            Index           =   4
            Left            =   3000
            TabIndex        =   50
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label lblSplitUMCsMaximumPeakCount 
            Caption         =   "Maximum peak count to split UMC"
            Height          =   405
            Left            =   120
            TabIndex        =   45
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label lblUnits 
            Caption         =   "peaks"
            Height          =   255
            Index           =   3
            Left            =   3000
            TabIndex        =   47
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label lblSplitUMCsMinimumDifferenceInAvgPpmMass 
            Caption         =   "Minimum difference in average mass"
            Height          =   405
            Left            =   120
            TabIndex        =   42
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label lblUnits 
            Caption         =   "ppm"
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   44
            Top             =   690
            Width           =   495
         End
      End
      Begin VB.CheckBox chkRemoveLoAbu 
         Caption         =   "Remove low intensity classes(%)"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox txtLoAbuPct 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   32
         Text            =   "30"
         Top             =   480
         Width           =   615
      End
      Begin VB.CheckBox chkRemoveHiAbu 
         Caption         =   "Remove high intensity classes(%)"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtHiAbuPct 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   57
         Text            =   "30"
         Top             =   840
         Width           =   615
      End
      Begin VB.CheckBox chkRemoveLoCnt 
         Caption         =   "Remove cls. with less than"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtLoCnt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2520
         TabIndex        =   35
         Text            =   "3"
         Top             =   1200
         Width           =   495
      End
      Begin VB.CheckBox chkRemoveHiCnt 
         Caption         =   "Remove cls. with more than"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txtHiCnt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2520
         TabIndex        =   38
         Text            =   "500"
         Top             =   1560
         Width           =   495
      End
      Begin VB.ComboBox cmbCountType 
         Height          =   315
         ItemData        =   "frmUMCWithAutoRefine.frx":0038
         Left            =   -74760
         List            =   "frmUMCWithAutoRefine.frx":003A
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtHoleNum 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -72120
         TabIndex        =   26
         Text            =   "0"
         Top             =   3420
         Width           =   495
      End
      Begin VB.TextBox txtHoleSize 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -72120
         TabIndex        =   28
         Text            =   "0"
         Top             =   3900
         Width           =   495
      End
      Begin VB.TextBox txtHolePct 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -72120
         TabIndex        =   30
         Text            =   "0"
         Top             =   4380
         Width           =   495
      End
      Begin VB.ComboBox cmbUMCAbu 
         Height          =   315
         ItemData        =   "frmUMCWithAutoRefine.frx":003C
         Left            =   -74760
         List            =   "frmUMCWithAutoRefine.frx":003E
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1240
         Width           =   2775
      End
      Begin VB.ComboBox cmbUMCMW 
         Height          =   315
         ItemData        =   "frmUMCWithAutoRefine.frx":0040
         Left            =   -74760
         List            =   "frmUMCWithAutoRefine.frx":0042
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label lblChargeStateAbuType 
         BackStyle       =   0  'Transparent
         Caption         =   "Most Abu Charge State Group Type"
         Height          =   255
         Left            =   -74760
         TabIndex        =   22
         Top             =   2280
         Width           =   3135
      End
      Begin VB.Label lblAutoRefineMinimumMemberCount 
         Caption         =   "Minimum member count:"
         Height          =   375
         Left            =   2160
         TabIndex        =   60
         Top             =   1905
         Width           =   1125
      End
      Begin VB.Label lblAutoRefineLengthLabel 
         Caption         =   "members"
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   36
         Top             =   1230
         Width           =   1000
      End
      Begin VB.Label lblAutoRefineLengthLabel 
         Caption         =   "members"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   39
         Top             =   1590
         Width           =   1000
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Count Type"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum number of scan holes in the Unique Mass Class:"
         Height          =   495
         Index           =   0
         Left            =   -74760
         TabIndex        =   25
         Top             =   3360
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum size of scan hole in the Unique Mass Class:"
         Height          =   495
         Index           =   1
         Left            =   -74760
         TabIndex        =   27
         Top             =   3840
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Percentage of allowed scan holes in the Unique Mass Class:"
         Height          =   495
         Index           =   2
         Left            =   -74760
         TabIndex        =   29
         Top             =   4320
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Class Abundance"
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   18
         Top             =   1000
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Class Molecular Mass"
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   20
         Top             =   1680
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "&Report"
      Height          =   375
      Left            =   3600
      TabIndex        =   55
      ToolTipText     =   "Generates various statistics on current UMC"
      Top             =   5520
      Width           =   975
   End
   Begin VB.Frame fraTol 
      Caption         =   "Molecular Mass Tolerance"
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   2295
      Begin VB.TextBox txtTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Text            =   "10"
         Top             =   520
         Width           =   735
      End
      Begin VB.OptionButton optTolType 
         Caption         =   "&Dalton"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   14
         Top             =   666
         Width           =   855
      End
      Begin VB.OptionButton optTolType 
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
         Caption         =   "Tolerance:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   280
         Width           =   735
      End
   End
   Begin VB.Frame fraMWField 
      Caption         =   "Molecular Mass Field"
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   2295
      Begin VB.OptionButton optMWField 
         Caption         =   "&The Most Abundant"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   920
         Width           =   1815
      End
      Begin VB.OptionButton optMWField 
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
      Caption         =   "Definition Scope"
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2295
      Begin VB.OptionButton optDefScope 
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
      Left            =   4680
      TabIndex        =   1
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&UMC"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      ToolTipText     =   "Generates UM Classes and returns number of it"
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdAbortProcessing 
      Caption         =   "Abort!"
      Height          =   375
      Left            =   2520
      TabIndex        =   56
      Top             =   5520
      Width           =   975
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   120
      X2              =   2400
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   120
      X2              =   2400
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   120
      X2              =   2400
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   2400
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   54
      Top             =   6000
      Width           =   6615
   End
   Begin VB.Label lblGelName 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "frmUMCWithAutoRefine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'unique mass class function
'breaks gel to unique mass classes
'-----------------------------------------
'last modified 3/6/2003 nt; mem
'-----------------------------------------
Option Explicit
Private CallerID As Long

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
        .AddItem "Median of Class Abundance"
        .AddItem "Max of Class Abu."
        .AddItem "Sum of Top X Members of Class"
    End With
    
    With cmbUMCMW
        .Clear
        .AddItem "Class Average"
        .AddItem "Mol.Mass Of Class Representative"
        .AddItem "Class Median"
        .AddItem "Average of Top X Members of Class"
        .AddItem "Median of Top X Members of Class"
    End With
    
    With cboChargeStateAbuType
        .Clear
        .AddItem "Highest Abu Sum"
        .AddItem "Most Abu Member"
        .AddItem "Most Members"
    End With
    
End Sub

Public Sub InitializeUMCSearch()
    ' MonroeMod: This code was in Form_Activate
    
On Error GoTo InitializeUMCSearchErrorHandler
    
    CallerID = Me.Tag
    ' MonroeMod
    If CallerID >= 1 And CallerID <= UBound(GelBody) Then UMCDef = GelSearchDef(CallerID).UMCDef
    lblGelName.Caption = CompactPathString(GelBody(CallerID).Caption, 65)
    
    With UMCDef
        txtTol.Text = .Tol
        If .UMCType = glUMC_TYPE_FROM_NET Then .UMCType = glUMC_TYPE_INTENSITY
        cmbCountType.ListIndex = .UMCType
        cmbUMCAbu.ListIndex = .ClassAbu
        cmbUMCMW.ListIndex = .ClassMW
        cboChargeStateAbuType.ListIndex = .ChargeStateStatsRepType
        SetCheckBox chkUseMostAbuChargeStateStatsForClassStats, .UMCClassStatsUseStatsFromMostAbuChargeState
        
        optDefScope(.DefScope).value = True
        optMWField(.MWField - MW_FIELD_OFFSET).value = True
        Select Case .TolType
        Case gltPPM
          optTolType(0).value = True
        Case gltABS
          optTolType(1).value = True
        Case Else
          Debug.Assert False
        End Select
        txtHoleNum.Text = .GapMaxCnt
        txtHoleSize.Text = .GapMaxSize
        txtHolePct.Text = CLng(.GapMaxPct * 100)
    End With
    
    With glbPreferencesExpanded.UMCAutoRefineOptions
        SetCheckBox chkRemoveLoCnt, .UMCAutoRefineRemoveCountLow
        SetCheckBox chkRemoveHiCnt, .UMCAutoRefineRemoveCountHigh
        txtLoCnt = .UMCAutoRefineMinLength
        txtHiCnt = .UMCAutoRefineMaxLength
        SetCheckBox chkRefineUMCLengthByScanRange, .TestLengthUsingScanRange
        txtAutoRefineMinimumMemberCount = .MinMemberCountWhenUsingScanRange
        UpdateDynamicControls
        
        SetCheckBox chkRemoveLoAbu, .UMCAutoRefineRemoveAbundanceLow
        SetCheckBox chkRemoveHiAbu, .UMCAutoRefineRemoveAbundanceHigh
        txtLoAbuPct = .UMCAutoRefinePctLowAbundance
        txtHiAbuPct = .UMCAutoRefinePctHighAbundance
        
        SetCheckBox chkSplitUMCsByExaminingAbundance, .SplitUMCsByAbundance
        With .SplitUMCOptions
            txtSplitUMCsMaximumPeakCount = Trim(.MaximumPeakCountToSplitUMC)
            txtSplitUMCsMinimumDifferenceInAvgPpmMass = Trim(.MinimumDifferenceInAveragePpmMassToSplit)
            txtSplitUMCsPeakDetectIntensityThresholdPercentageOfMax = Trim(.PeakDetectIntensityThresholdPercentageOfMaximum)
            txtSplitUMCsPeakPickingMinimumWidth = Trim(.PeakWidthPointsMinimum)
        End With
    End With

    Exit Sub

InitializeUMCSearchErrorHandler:
    Debug.Print "Error in InitializeUMCSearch: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "frmUMCWithAutoRefine->InitializeUMCSearch"
    Resume Next

End Sub

Public Function StartUMCSearch() As Boolean
    ' Returns True if success, False if error or aborted
    Dim TtlCnt As Long
    Dim Cnt As Long
    Dim blnUMCIndicesUpdated As Boolean
    
On Error GoTo UMCSearchErrorHandler
    
    Me.MousePointer = vbHourglass
    cmdOK.Visible = False
    glAbortUMCProcessing = False
    TtlCnt = GelData(CallerID).CSLines + GelData(CallerID).IsoLines
    Cnt = UMCCount(CallerID, TtlCnt, Me, True)
    Me.MousePointer = vbDefault
    If Cnt >= 0 And Not glAbortUMCProcessing Then
        AddToAnalysisHistory CallerID, ConstructUMCDefDescription(CallerID, AUTO_ANALYSIS_UMC2002, UMCDef, glbPreferencesExpanded.UMCAdvancedStatsOptions, False)
        
        ' Possibly Auto-Refine the UMC's
        blnUMCIndicesUpdated = AutoRefineUMCs(CallerID, Me)
        
       ' Note: we need to update GelSearchDef before calling SplitUMCsByAbundance
       GelSearchDef(CallerID).UMCDef = UMCDef
        
        ' Update the IonToUMC Indices
        If Not glAbortUMCProcessing Then
            If Not blnUMCIndicesUpdated Then
                ' The following calls CalculateClasses, UpdateIonToUMCIndices, and InitDrawUMC
                UpdateUMCStatArrays CallerID, False, Me
            End If
        
            If glbPreferencesExpanded.UMCAutoRefineOptions.SplitUMCsByAbundance Then
               SplitUMCsByAbundance CallerID, Me, False, True
            End If
        End If
        
        Status "Number of Unique Mass Classes: " & GelUMC(CallerID).UMCCnt
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "Number of Unique Mass Classes: " & GelUMC(CallerID).UMCCnt
        End If
        
    Else
       Status ")-: Error counting Unique Mass Classes :-("
    End If

    cmdOK.Visible = True
    glAbortUMCProcessing = False
    'if there is new UMC count everything done with pairs
    'has to be redone if pairs are UMC pairs
    With GelP_D_L(CallerID)
        If .DltLblType <> ptS_Dlt And .DltLblType <> ptS_Lbl And .DltLblType <> ptS_DltLbl Then
            .SyncWithUMC = False
        End If
    End With
    StartUMCSearch = True
    Exit Function
    
UMCSearchErrorHandler:
    Debug.Print "Error in frmUMCWithAutoRefine->StartUMCSearch: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "frmUMCWithAutoRefine->StartUMCSearch"
    cmdOK.Visible = True
    glAbortUMCProcessing = False
    Me.MousePointer = vbDefault
    StartUMCSearch = False
    
End Function

Public Sub Status(ByVal StatusText As String)
lblStatus.Caption = StatusText
DoEvents
End Sub

Private Sub UpdateDynamicControls()
    ' Update the UMC auto refine length labels
    If glbPreferencesExpanded.UMCAutoRefineOptions.TestLengthUsingScanRange Then
        chkRemoveLoCnt.Caption = "Remove classes less than"
        chkRemoveHiCnt.Caption = "Remove classes more than"
        lblAutoRefineLengthLabel(0) = "scans wide"
        lblAutoRefineLengthLabel(1) = "scans wide"
        lblAutoRefineMinimumMemberCount.Visible = True
        txtAutoRefineMinimumMemberCount.Visible = True
    Else
        chkRemoveLoCnt.Caption = "Remove cls. with less than"
        chkRemoveHiCnt.Caption = "Remove cls. with more than"
        lblAutoRefineLengthLabel(0) = "members"
        lblAutoRefineLengthLabel(1) = "members"
        lblAutoRefineMinimumMemberCount.Visible = False
        txtAutoRefineMinimumMemberCount.Visible = False
    End If
End Sub

Private Sub cboChargeStateAbuType_Click()
    UMCDef.ChargeStateStatsRepType = cboChargeStateAbuType.ListIndex
End Sub

Private Sub chkRefineUMCLengthByScanRange_Click()
    glbPreferencesExpanded.UMCAutoRefineOptions.TestLengthUsingScanRange = cChkBox(chkRefineUMCLengthByScanRange)
    UpdateDynamicControls
End Sub

Private Sub chkRemoveHiAbu_Click()
    glbPreferencesExpanded.UMCAutoRefineOptions.UMCAutoRefineRemoveAbundanceHigh = cChkBox(chkRemoveHiAbu)
End Sub

Private Sub chkRemoveHiCnt_Click()
    glbPreferencesExpanded.UMCAutoRefineOptions.UMCAutoRefineRemoveCountHigh = cChkBox(chkRemoveHiCnt)
End Sub

Private Sub chkRemoveLoAbu_Click()
    glbPreferencesExpanded.UMCAutoRefineOptions.UMCAutoRefineRemoveAbundanceLow = cChkBox(chkRemoveLoAbu)
End Sub

Private Sub chkRemoveLoCnt_Click()
    glbPreferencesExpanded.UMCAutoRefineOptions.UMCAutoRefineRemoveCountLow = cChkBox(chkRemoveLoCnt)
End Sub

Private Sub chkSplitUMCsByExaminingAbundance_Click()
    glbPreferencesExpanded.UMCAutoRefineOptions.SplitUMCsByAbundance = cChkBox(chkSplitUMCsByExaminingAbundance)
End Sub

Private Sub chkUseMostAbuChargeStateStatsForClassStats_Click()
    UMCDef.UMCClassStatsUseStatsFromMostAbuChargeState = cChkBox(chkUseMostAbuChargeStateStatsForClassStats)
End Sub

Private Sub cmbUMCAbu_Click()
UMCDef.ClassAbu = cmbUMCAbu.ListIndex
End Sub

Private Sub cmbUMCMW_Click()
UMCDef.ClassMW = cmbUMCMW.ListIndex
End Sub

Private Sub cmdAbortProcessing_Click()
    glAbortUMCProcessing = True
End Sub

Private Sub cmbCountType_Click()
UMCDef.UMCType = cmbCountType.ListIndex
End Sub

Private Sub cmdCancel_Click()
If Not cmdOK.Visible Then glAbortUMCProcessing = True
Unload Me
End Sub

Private Sub cmdOK_Click()
    StartUMCSearch
End Sub

Private Sub cmdReport_Click()
Me.MousePointer = vbHourglass
Status "Generating UMC report..."
Call ReportUMC(CallerID, "UMC 2003")
Status ""
Me.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    InitializeUMCSearch
End Sub

Private Sub Form_Load()
    ' MonroeMod: The code that was here has been moved to Form_Activate
    '            This was done so that the Statement: UMCDef = GelSearchDef(CallerID).UMCDef
    '             will be encountered before the controls are updated
    FillComboBoxes
    tbsTabStrip.Tab = 0
End Sub

Private Sub optDefScope_Click(Index As Integer)
UMCDef.DefScope = Index
End Sub

Private Sub optTolType_Click(Index As Integer)
If Index = 0 Then
   UMCDef.TolType = gltPPM
Else
   UMCDef.TolType = gltABS
End If
End Sub

Private Sub optMWField_Click(Index As Integer)
UMCDef.MWField = 6 + Index
End Sub


Private Sub txtAutoRefineMinimumMemberCount_LostFocus()
If IsNumeric(txtAutoRefineMinimumMemberCount.Text) Then
    glbPreferencesExpanded.UMCAutoRefineOptions.MinMemberCountWhenUsingScanRange = Abs(CLng(txtAutoRefineMinimumMemberCount.Text))
Else
   MsgBox "This argument should be non-negative integer!", vbOKOnly, glFGTU
   txtAutoRefineMinimumMemberCount.SetFocus
End If
End Sub

Private Sub txtHiAbuPct_LostFocus()
If IsNumeric(txtHiAbuPct.Text) Then
   glbPreferencesExpanded.UMCAutoRefineOptions.UMCAutoRefinePctHighAbundance = Abs(CDbl(txtHiAbuPct.Text))
Else
   MsgBox "This argument should be non-negative number!", vbOKOnly, glFGTU
   txtHiAbuPct.SetFocus
End If
End Sub

Private Sub txtHiCnt_LostFocus()
If IsNumeric(txtHiCnt.Text) Then
    glbPreferencesExpanded.UMCAutoRefineOptions.UMCAutoRefineMaxLength = Abs(CLng(txtHiCnt.Text))
Else
   MsgBox "This argument should be non-negative integer!", vbOKOnly, glFGTU
   txtHiCnt.SetFocus
End If
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

Private Sub txtLoAbuPct_LostFocus()
If IsNumeric(txtLoAbuPct.Text) Then
   glbPreferencesExpanded.UMCAutoRefineOptions.UMCAutoRefinePctLowAbundance = Abs(CDbl(txtLoAbuPct.Text))
Else
   MsgBox "This argument should be non-negative number!", vbOKOnly, glFGTU
   txtLoAbuPct.SetFocus
End If
End Sub

Private Sub txtLoCnt_LostFocus()
If IsNumeric(txtLoCnt.Text) Then
    glbPreferencesExpanded.UMCAutoRefineOptions.UMCAutoRefineMinLength = Abs(CLng(txtLoCnt.Text))
Else
   MsgBox "This argument should be non-negative integer!", vbOKOnly, glFGTU
   txtLoCnt.SetFocus
End If
End Sub

Private Sub txtSplitUMCsMaximumPeakCount_LostFocus()
    ValidateTextboxValueLng txtSplitUMCsMaximumPeakCount, 2, 100, 6
    glbPreferencesExpanded.UMCAutoRefineOptions.SplitUMCOptions.MaximumPeakCountToSplitUMC = CLngSafe(txtSplitUMCsMaximumPeakCount)
End Sub

Private Sub txtSplitUMCsMinimumDifferenceInAvgPpmMass_LostFocus()
    ValidateTextboxValueDbl txtSplitUMCsMinimumDifferenceInAvgPpmMass, 0, 10000#, 4
    glbPreferencesExpanded.UMCAutoRefineOptions.SplitUMCOptions.MinimumDifferenceInAveragePpmMassToSplit = CDblSafe(txtSplitUMCsMinimumDifferenceInAvgPpmMass)
End Sub

Private Sub txtSplitUMCsPeakDetectIntensityThresholdPercentageOfMax_LostFocus()
    ValidateTextboxValueLng txtSplitUMCsPeakDetectIntensityThresholdPercentageOfMax, 0, 100, 15
    glbPreferencesExpanded.UMCAutoRefineOptions.SplitUMCOptions.PeakDetectIntensityThresholdPercentageOfMaximum = CLngSafe(txtSplitUMCsPeakDetectIntensityThresholdPercentageOfMax)
End Sub

Private Sub txtSplitUMCsPeakPickingMinimumWidth_LostFocus()
    ValidateTextboxValueLng txtSplitUMCsPeakPickingMinimumWidth, 0, 1000, 4
    glbPreferencesExpanded.UMCAutoRefineOptions.SplitUMCOptions.PeakWidthPointsMinimum = CLngSafe(txtSplitUMCsPeakPickingMinimumWidth)
End Sub

Private Sub txtTol_LostFocus()
If IsNumeric(txtTol.Text) Then
   UMCDef.Tol = txtTol.Text
Else
   MsgBox "Molecular Mass Tolerance should be numeric value!", vbOKOnly
   txtTol.SetFocus
End If
End Sub
