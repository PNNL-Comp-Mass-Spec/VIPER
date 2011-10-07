VERSION 5.00
Begin VB.Form frmSearchMTPairs 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Search MT Tag Database For Potential N14/N15 Pairs"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6990
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   6990
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDBSearchMinimumPeptideProphetProbability 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      TabIndex        =   30
      Text            =   "0"
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox txtDBSearchMinimumHighDiscriminantScore 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      TabIndex        =   28
      Text            =   "0"
      Top             =   3900
      Width           =   615
   End
   Begin VB.TextBox txtDBSearchMinimumHighNormalizedScore 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      TabIndex        =   26
      Text            =   "0"
      Top             =   3600
      Width           =   615
   End
   Begin VB.Frame fraMods 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modifications"
      Height          =   1575
      Left            =   120
      TabIndex        =   31
      Top             =   4560
      Width           =   5655
      Begin VB.TextBox txtDecoySearchNETWobble 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4320
         TabIndex        =   45
         Text            =   "0.1"
         Top             =   1080
         Width           =   615
      End
      Begin VB.Frame fraOptionFrame 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   750
         Index           =   49
         Left            =   3170
         TabIndex        =   39
         Top             =   240
         Width           =   1920
         Begin VB.OptionButton optDBSearchModType 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fixed"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   41
            ToolTipText     =   "Changes the mass of all loaded AMTs, adding the value specified by the modification mass"
            Top             =   240
            Width           =   750
         End
         Begin VB.OptionButton optDBSearchModType 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Dynamic"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   42
            ToolTipText     =   $"frmSearchMTPairs.frx":0000
            Top             =   480
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optDBSearchModType 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Decoy"
            Height          =   255
            Index           =   2
            Left            =   1080
            TabIndex        =   43
            ToolTipText     =   $"frmSearchMTPairs.frx":0092
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblDescription 
            BackStyle       =   0  'Transparent
            Caption         =   "Mod Type:"
            Height          =   255
            Index           =   100
            Left            =   120
            TabIndex        =   40
            Top             =   0
            Width           =   900
         End
      End
      Begin VB.CheckBox chkAlkylation 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Alkylation"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         ToolTipText     =   "Check to add the alkylation mass correction below to all MT Tag masses (added to each cys residue)"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtAlkylationMWCorrection 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   360
         TabIndex        =   34
         Text            =   "57.0215"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtResidueToModifyMass 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         TabIndex        =   38
         Text            =   "0"
         Top             =   1080
         Width           =   855
      End
      Begin VB.ComboBox cboResidueToModify 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblDecoySearchNETWobble 
         BackStyle       =   0  'Transparent
         Caption         =   "Decoy NET Wobble"
         Height          =   375
         Left            =   3240
         TabIndex        =   44
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Alkylation mass:"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   720
         Width           =   1335
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   3120
         X2              =   3120
         Y1              =   240
         Y2              =   1440
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Mass (Da):"
         Height          =   255
         Left            =   1680
         TabIndex        =   37
         Top             =   840
         Width           =   1095
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   1560
         X2              =   1560
         Y1              =   240
         Y2              =   1440
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Residue to modify:"
         Height          =   255
         Left            =   1680
         TabIndex        =   35
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CheckBox chkUseUMCConglomerateMass 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Use UMC Conglomerate Mass"
      Height          =   495
      Left            =   5040
      TabIndex        =   15
      Top             =   2880
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox chkUpdateGelDataWithSearchResults 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Update data in current file with results of search"
      Height          =   615
      Left            =   5040
      TabIndex        =   14
      Top             =   2160
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.Frame fraNET 
      BackColor       =   &H00E0E0E0&
      Caption         =   "NET  Calculation"
      Height          =   1455
      Left            =   120
      TabIndex        =   17
      Top             =   2040
      Width           =   4815
      Begin VB.TextBox txtNETTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3600
         TabIndex        =   23
         Text            =   "0.1"
         Top             =   1040
         Width           =   615
      End
      Begin VB.OptionButton optNETorRT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pred. NET"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   19
         ToolTipText     =   "Use NET calculated only from Sequest ""first choice"" peptides"
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optNETorRT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Obs. NET"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   18
         ToolTipText     =   "Use NET calculated from all peptides of MT Tags"
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox txtNETFormula 
         Height          =   285
         Left            =   1560
         TabIndex        =   21
         Text            =   "(FN-MinFN)/(MaxFN-MinFN)"
         Top             =   660
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "T&olerance"
         Height          =   255
         Left            =   2760
         TabIndex        =   22
         ToolTipText     =   "Normalized Elution Time Tolerance (if blank NET will not be considered in search)"
         Top             =   1060
         Width           =   855
      End
      Begin VB.Label lblNETFormula 
         BackStyle       =   0  'Transparent
         Caption         =   "&Formula  F(FN, MinFN, MaxFN)"
         Height          =   255
         Left            =   1560
         TabIndex        =   20
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame fraMWTolerance 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Molecular Mass Tolerance"
      Height          =   1455
      Left            =   4320
      TabIndex        =   8
      Top             =   120
      Width           =   2535
      Begin VB.ComboBox cboSearchRegionShape 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   960
         Width           =   2295
      End
      Begin VB.OptionButton optTolType 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Dalton"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   12
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton optTolType 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&ppm"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtMWTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   160
         TabIndex        =   10
         Text            =   "10"
         Top             =   525
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tolerance"
         Height          =   255
         Left            =   165
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraMWField 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Molecular Mass Field"
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   2175
      Begin VB.OptionButton optMWField 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&The Most Abundant"
         Height          =   255
         Index           =   2
         Left            =   80
         TabIndex        =   7
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton optMWField 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Monoisotopic"
         Height          =   255
         Index           =   1
         Left            =   80
         TabIndex        =   6
         Top             =   540
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optMWField 
         BackColor       =   &H00E0E0E0&
         Caption         =   "A&verage"
         Height          =   255
         Index           =   0
         Left            =   80
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      ToolTipText     =   "Searches loaded MT Tags for matches with established potential pairs"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum Peptide Prophet Probability"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   29
      Top             =   4215
      Width           =   2865
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum PMT Discriminant Score"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   27
      Top             =   3920
      Width           =   2505
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum MT Tag XCorr"
      Height          =   255
      Index           =   134
      Left            =   120
      TabIndex        =   25
      Top             =   3620
      Width           =   2145
   End
   Begin VB.Label lblETType 
      BackStyle       =   0  'Transparent
      Caption         =   "Generic NET"
      Height          =   255
      Left            =   4200
      TabIndex        =   24
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2940
      Picture         =   "frmSearchMTPairs.frx":012F
      ToolTipText     =   "Double-click for short info on this procedure"
      Top             =   600
      Width           =   480
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1755
      Width           =   6735
   End
   Begin VB.Label lblMTStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MTDB Status:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Status of the MT Tag database"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Menu mnuP 
      Caption         =   "&Pairs"
      Begin VB.Menu mnuPSearch 
         Caption         =   "Search (Identify)"
      End
      Begin VB.Menu mnuPSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPExcludeUnidentified 
         Caption         =   "Exclude Unidentified Pairs"
      End
      Begin VB.Menu mnuPExcludeIdentified 
         Caption         =   "Exclude Identified"
      End
      Begin VB.Menu mnuPIncludeUnqIdentified 
         Caption         =   "Include Only Uniquely Identified"
      End
      Begin VB.Menu mnuPSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPExcludeAmbiguous 
         Caption         =   "Exclude Ambiguous Pairs (all pairs)"
      End
      Begin VB.Menu mnuPExcludeAmbiguousHitsOnly 
         Caption         =   "Exclude Ambiguous Pairs (only those with hits)"
      End
      Begin VB.Menu mnuPResetExclusionFlags 
         Caption         =   "Reset Exclusion Flags for All Pairs"
      End
      Begin VB.Menu mnuPDeleteExcluded 
         Caption         =   "Delete Excluded Pairs"
      End
      Begin VB.Menu mnuPSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPCalculateER 
         Caption         =   "Calculate ER"
      End
      Begin VB.Menu mnuPSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPExpLegacyDB 
         Caption         =   "Export Results To &Legacy DB"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuExpMTDB 
         Caption         =   "Export Results To &MT Tag DB"
      End
      Begin VB.Menu mnuPExportDetailedMemberInformation 
         Caption         =   "Export detailed member information for each LC-MS Feature"
      End
      Begin VB.Menu mnuPSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSyncPairsStructure 
         Caption         =   "Sync With ID Pairs Structure"
      End
      Begin VB.Menu mnuPSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Report"
      Begin VB.Menu mnuRAllPairsAndIDs 
         Caption         =   "All Pairs And Identifications"
      End
      Begin VB.Menu mnuRIdentified 
         Caption         =   "Identified Pairs Only"
      End
      Begin VB.Menu mnuRUnqIdentified 
         Caption         =   "Uniquely Identified Pairs"
      End
      Begin VB.Menu mnuRUnidentified 
         Caption         =   "Unidentified Pairs Only"
      End
      Begin VB.Menu mnuReportSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReportIncludeORFName 
         Caption         =   "Include Proteins (ORFs) in Report"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuMT 
      Caption         =   "&MT Tags"
      Begin VB.Menu mnuMTLoadMT 
         Caption         =   "Load MT Tag DB"
      End
      Begin VB.Menu mnuMTLoadLegacy 
         Caption         =   "Load Legacy MT DB"
      End
      Begin VB.Menu mnuMTSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMTStatus 
         Caption         =   "MT Tags Status"
      End
   End
   Begin VB.Menu mnuETHeader 
      Caption         =   "&Elution Time"
      Begin VB.Menu mnuET 
         Caption         =   "&Generic NET"
         Index           =   0
      End
      Begin VB.Menu mnuET 
         Caption         =   "&TIC Fit NET"
         Index           =   1
      End
      Begin VB.Menu mnuET 
         Caption         =   "G&ANET"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmSearchMTPairs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'search of MT tag database for pairs member of the isotopic
'labeled pairs
'assumption is that pairs are calculated for UMC and that
'search of the database can be performed with loose tolerance
'with additional criteria of matching number of N with delta
'in molecular mass for established pairs;
'------------------------------------------------------------
'NOTE: search is always done for light pair members
'------------------------------------------------------------
'created: 11/12/2001 nt
'last modified: 07/29/2002 nt
'------------------------------------------------------------
Option Explicit

Private Const NET_PRECISION = 5

Const MOD_TKN_NONE = "none"
Const MOD_TKN_PEO = "PEO"
Const MOD_TKN_ICAT_D0 = "ICAT_D0"
Const MOD_TKN_ICAT_D8 = "ICAT_D8"
Const MOD_TKN_ALK = "ALK"
Const MOD_TKN_N14 = "N14"
Const MOD_TKN_N15 = "N15"
Const MOD_TKN_RES_MOD = "RES_MOD"
Const MOD_TKN_MT_MOD = "MT_MOD"

Const MOD_TKN_PAIR_LIGHT = "N14N15_LIGHT"
Const MOD_TKN_PAIR_HEAVY = "N14N15_HEAVY"

Const MODS_FIXED = 0
Const MODS_DYNAMIC = 1
Const MODS_DECOY = 2

'if called with any positive number add that many points
Const MNG_RESET = 0
Const MNG_ERASE = -1
Const MNG_TRIM = -2
Const MNG_ADD_START_SIZE = -3

Const MNG_START_SIZE = 100

Const NET_WOBBLE_SEED = 1000

'in this case CallerID is a public property
Public CallerID As Long

Private bLoading As Boolean

Private OldSearchFlag As Long

'for faster search mass array will be sorted; therefore all other arrays
'has to be addressed indirectly (mMTNET(mMTInd(i))
Private mMTCnt                  'count of masses to search
Private mMTInd() As Long        'index(unique key)              ' 0-based array
Private mMTOrInd() As Long      'index of original MT tag (in AMT array)
Private mMTMWN14() As Double    'AMT mass to look for
Private mMTNET() As Double      'NET value
Private mMTMods() As String     'modification description

Private MWFastSearch As MWUtil

Private AlkMWCorrection As Double
Private mSearchRegionShape As srsSearchRegionShapeConstants

Private ClsCnt As Long              'this is not actually neccessary except
Private ClsStat() As Double         'to create nice reports
                                
'mUMCMatchStats contains all possible identifications for all pairs with scores
'as count of each identification hits within the UMC

Private mMatchStatsCount As Long                                'count of Pair-ID matches
Private mUMCMatchStats() As udtUMCMassTagMatchStats             ' 0-based array

' The following hold match stats for each individual Pair
Private mCurrIDCnt As Long
Private mCurrIDMatches() As udtUMCMassTagRawMatches         ' 0-based array


'The following arrays are parallel to the pairs arrays; it is used for
'easier classification between identified and nonidentified pairs
Private PCount As Long              'shortcut for number of pairs
Private PIDCnt() As Long            'count of OK identifications(unique) for pair
Private PIDInd1() As Long           'first index in ID arrays for pair (pointer into mUMCMatchStats)
Private PIDInd2() As Long           'last index in ID arrays for pair (pointer into mUMCMatchStats)


'Expression Evaluator variables for elution time calculation
Private MyExprEva As ExprEvaluator
Private VarVals() As Long
Private MinFN As Long
Private MaxFN As Long

'names of stored procedures that will write data
'to database tables retrieved from init. file
Private ExpAnalysisSPName As String             ' Stored procedure AddMatchMaking
''Private ExpPeakSPName As String                 ' Stored procedure AddFTICRPeak; Unused variable
Private ExpUmcSPName As String                  ' Stored procedure AddFTICRUmc
Private ExpUMCMemberSPName As String            ' Stored procedure AddFTICRUmcMember
Private ExpUmcMatchSPName As String             ' Stored procedure AddFTICRUmcMatch
Private ExpUMCCSStats As String                 ' Stored procedure AddFTICRUmcCSStats
Private ExpQuantitationDescription As String    ' Stored procedure AddQuantitationDescription

Private mKeyPressAbortProcess As Integer
Private mUsingDefaultGANET As Boolean
Private mMTMinimumHighNormalizedScore As Single
Private mMTMinimumHighDiscriminantScore As Single
Private mMTMinimumPeptideProphetProbability As Single

Private objMTDBNameLookupClass As mtdbMTNames
'

Public Property Get SearchRegionShape() As srsSearchRegionShapeConstants
    SearchRegionShape = mSearchRegionShape
End Property
Public Property Let SearchRegionShape(Value As srsSearchRegionShapeConstants)
    cboSearchRegionShape.ListIndex = Value
    mSearchRegionShape = Value
End Property

Private Function CheckNAtomsVsDeltaCount(MW As Double, DltCnt As Long, AMTIndex As Long) As Boolean
    ' Returns True if the number of N atoms is valid
    
    Dim blnValidHit As Boolean
    Dim lngN14N15CorrectionMass As Long
    
    'lngN14N15CorrectionMass = glN14_N15CorrMW
    lngN14N15CorrectionMass = 1500
        
    ' Allow for +/-1 error for masses over lngN14N15CorrectionMass Da
    ' Allow for +/-2 error for masses over 2*lngN14N15CorrectionMass Da
        
    blnValidHit = False
    If MW >= 2 * lngN14N15CorrectionMass Then
       If Abs(AMTData(AMTIndex).CNT_N - DltCnt) <= 2 Then
          blnValidHit = True
       End If
    ElseIf MW >= lngN14N15CorrectionMass Then
       If Abs(AMTData(AMTIndex).CNT_N - DltCnt) <= 1 Then
          blnValidHit = True
       End If
    ElseIf MW >= 1000 Then
       If AMTData(AMTIndex).CNT_N = DltCnt Or AMTData(AMTIndex).CNT_N = DltCnt + 1 Then
          blnValidHit = True
       End If
    Else
       If AMTData(AMTIndex).CNT_N = DltCnt Then
          blnValidHit = True
       End If
    End If

    CheckNAtomsVsDeltaCount = blnValidHit
End Function

Private Sub CheckNETEquationStatus()
    If GelData(CallerID).CustomNETsDefined Then
        mUsingDefaultGANET = True
    Else
        If Not GelAnalysis(CallerID) Is Nothing Then
            If txtNETFormula.Text = ConstructNETFormula(GelAnalysis(CallerID).GANET_Slope, GelAnalysis(CallerID).GANET_Intercept) _
               And InStr(UCase(txtNETFormula), "MINFN") = 0 Then
                mUsingDefaultGANET = True
            Else
                mUsingDefaultGANET = False
            End If
        Else
            mUsingDefaultGANET = False
        End If
    End If
End Sub

Private Function ConvertScanToNET(lngScanNumber As Long) As Double

    If mUsingDefaultGANET Then
        ConvertScanToNET = ScanToGANET(CallerID, lngScanNumber)
    Else
        ConvertScanToNET = Elution(lngScanNumber, MinFN, MaxFN)
    End If

End Function

Public Sub DeleteExcludedPairsWrapper()
    '--------------------------------------------
    'removes excluded pairs from the structure
    '--------------------------------------------
    
    Dim strMessage As String
    
    UpdateStatus "Deleting excluded pairs ..."
    Me.MousePointer = vbHourglass
    
    strMessage = DeleteExcludedPairs(CallerID)
    UpdateStatus strMessage
    
    AddToAnalysisHistory CallerID, strMessage
    
    ' Must reset PCount to 0 now, since possibly deleted some pairs
    PCount = 0
    ReDim PIDCnt(0)
    ReDim PIDInd1(0)
    ReDim PIDInd2(0)
    
    Me.MousePointer = vbDefault
End Sub

Private Sub DestroyIDStructures()
On Error Resume Next
mMatchStatsCount = 0
Erase mUMCMatchStats
Call ManageCurrID(MNG_ERASE)
End Sub

Private Sub DestroySearchStructures()
On Error Resume Next
mMTCnt = 0
Erase mMTInd
Erase mMTOrInd
Erase mMTMWN14
Erase mMTNET
Erase mMTMods
Set MWFastSearch = Nothing
End Sub

Private Function Elution(FN As Long, MinFN As Long, MaxFN As Long) As Double
'---------------------------------------------------
'this function does not care are we using NET or RT
'---------------------------------------------------
VarVals(1) = FN
VarVals(2) = MinFN
VarVals(3) = MaxFN
Elution = MyExprEva.ExprVal(VarVals())
End Function

Private Sub EnableDisableControls()
    If optDBSearchModType(2).Value = True Then
        txtDecoySearchNETWobble.Enabled = True
    Else
        txtDecoySearchNETWobble.Enabled = False
    End If
End Sub

Public Sub ExcludeAmbiguousPairsWrapper(blnOnlyExaminePairsWithHits As Boolean)
    '---------------------------------------------------
    'mark as excluded all ambiguous pairs
    'to increase the number of unambiguous pairs, this
    'procedure should be applied at the end, after all
    'other filtering
    '---------------------------------------------------
    
    Dim strMessage As String
    
    If blnOnlyExaminePairsWithHits Then
        strMessage = PairsSearchMarkAmbiguousPairsWithHitsOnly(Me, CallerID)
    Else
        strMessage = PairsSearchMarkAmbiguous(Me, CallerID, True)
    End If
    
    UpdateStatus strMessage
End Sub

Public Function ExportMTDBbyUMC(Optional blnUpdateGANETForAnalysisInDB As Boolean = True, Optional ByVal blnExportUMCMembers As Boolean = False, Optional strIniFileName As String = "", Optional ByRef lngErrorNumber As Long, Optional ByRef lngMDID As Long, Optional ByVal blnExportExcludedPairs As Boolean = True) As String
'--------------------------------------------------------------------------------
' This function exports data to both T_FTICR_Peak_Results and T_FTICR_UMC_Results (plus T_FTICR_UMC_ResultDetails)
' Optionally returns the error number in lngErrorNumber
' Optionally returns the MD_ID value in lngMDID
'--------------------------------------------------------------------------------
    
    Dim strStatus As String
    Dim blnAddQuantitationEntry As Boolean
    
    lngMDID = -1
    cmdSearch.Visible = False
    UpdateStatus "Exporting ..."
    Me.MousePointer = vbHourglass
        
    If glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        With glbPreferencesExpanded.AutoAnalysisOptions
            blnAddQuantitationEntry = .AddQuantitationDescriptionEntry
        End With
    End If
    
    
    '' Legacy: strStatus = ExportIDPairsToPeakResultsTable(lngMDID, blnUpdateGANETForAnalysisInDB, lngErrorNumber, False, blnExportExcludedPairs)
    
    ' Note: The following function call will create a new entry in T_Match_Making_Description
    TraceLog 5, "frmSearchMTPairs->ExportMTDBbyUMC", "Call ExportIDPairsToUMCResultsTable"
    strStatus = strStatus & vbCrLf & ExportIDPairsToUMCResultsTable(lngMDID, True, blnUpdateGANETForAnalysisInDB, blnExportUMCMembers, lngErrorNumber, blnAddQuantitationEntry, strIniFileName, blnExportExcludedPairs)
    
    Me.MousePointer = vbDefault
    UpdateStatus strStatus
    cmdSearch.Visible = True
    
    ExportMTDBbyUMC = strStatus
    
End Function


Private Function ExportIDPairsToUMCResultsTable(ByRef lngMDID As Long, Optional blnCreateNewEntryInMMDTable As Boolean = False, Optional blnUpdateGANETForAnalysisInDB As Boolean = True, Optional ByVal blnExportUMCMembers As Boolean = False, Optional ByRef lngErrorNumber As Long, Optional ByVal blnAddQuantitationDescriptionEntry As Boolean = True, Optional ByVal strIniFileName As String = "", Optional ByVal blnExportExcludedPairs As Boolean = True) As String
'---------------------------------------------------
'This function will export data to the T_FTICR_UMC_Results table and the T_FTICR_UMC_ResultDetails tables
'
'It will create a new entry in the T_Match_Making_Description if blnCreateNewEntryInMMDTable = True
'If blnAddQuantitationDescriptionEntry = True, then calls ExportMTDBAddQuantitationDescriptionEntry
'  to create a new entry in T_Quantitation_Description and T_Quantitation_MDIDs
'
'Returns a status message
'lngErrorNumber will contain the error number, if an error occurs
'---------------------------------------------------
Dim lngPairInd As Long
Dim lngMassTagIndexPointer As Long, lngMassTagIndexOriginal As Long
Dim IndL As Long, IndH As Long

Dim MassTagExpCnt As Long
Dim strCaptionSaved As String
Dim strExportStatus As String
Dim strMassMods As String

Dim udtPairMatchStats As udtPairMatchStatsType

'ADO objects for stored procedure adding Match Making row
Dim cnNew As New ADODB.Connection
Dim sngDBSchemaVersion As Single

'ADO objects for stored procedure that adds FTICR UMC rows
Dim cmdPutNewUMC As New ADODB.Command
Dim udtPutUMCParams As udtPutUMCParamsListType
    
'ADO objects for stored procedure that adds FTICR UMC member rows
Dim cmdPutNewUMCMember As New ADODB.Command
Dim udtPutUMCMemberParams As udtPutUMCMemberParamsListType
    
'ADO objects for stored procedure adding UMC UMC Details
Dim cmdPutNewUMCMatch As New ADODB.Command
Dim udtPutUMCMatchParams As udtPutUMCMatchParamsListType

'ADO objects for stored procedure adding FTICR UMC CS Stats
Dim cmdPutNewUMCCSStats As New ADODB.Command
Dim udtPutUMCCSStatsParams As udtPutUMCCSStatsParamsListType

Dim strSearchDescription As String

On Error GoTo err_ExportMTDBbyUMC

strCaptionSaved = Me.Caption

' Connect to the database
TraceLog 4, "frmSearchMTPairs->ExportIDPairsToUMCResultsTable", "Connect to database"
Me.Caption = "Connecting to the database"
If Not EstablishConnection(cnNew, GelAnalysis(CallerID).MTDB.cn.ConnectionString, False) Then
    Debug.Assert False
    lngErrorNumber = -1
    Me.Caption = strCaptionSaved
    ExportIDPairsToUMCResultsTable = "Error: Unable to establish a connection to the database"
    Exit Function
End If

' Lookup the DB Schema Version
sngDBSchemaVersion = LookupDBSchemaVersion(cnNew)

If blnExportUMCMembers < 2 Then
    ' Force UMC Member export to be false
    blnExportUMCMembers = False
End If

If blnCreateNewEntryInMMDTable Then
    TraceLog 5, "frmSearchMTPairs->ExportIDPairsToUMCResultsTable", "Call AddEntryToMatchMakingDescriptionTable"
    'first write new analysis in T_Match_Making_Description table
    lngErrorNumber = AddEntryToMatchMakingDescriptionTableEx(cnNew, _
                                                             lngMDID, _
                                                             ExpAnalysisSPName, _
                                                             CallerID, _
                                                             mMatchStatsCount, _
                                                             GelData(CallerID).CustomNETsDefined, _
                                                             True, _
                                                             strIniFileName, _
                                                             mMTCnt, _
                                                             False, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
Else
    lngErrorNumber = 0
End If

If lngErrorNumber <> 0 Then
    Debug.Assert False
    GoTo err_Cleanup
End If

If blnCreateNewEntryInMMDTable Or mMatchStatsCount > 0 Then
    ' MonroeMod
    strSearchDescription = "N14/N15 Identification Pairs results (Conglomerate LC-MS Feature Mass)"
    
    AddToAnalysisHistory CallerID, "Exported " & strSearchDescription & " to LC-MS Feature Results table in database (" & ExtractDBNameFromConnectionString(GelAnalysis(CallerID).MTDB.cn.ConnectionString) & "); MMD_ID = " & lngMDID
    If blnCreateNewEntryInMMDTable Then
        AddToAnalysisHistory CallerID, "Export to MMD table details: Reference Job = " & GelAnalysis(CallerID).MD_Reference_Job & "; MD_File = " & GelAnalysis(CallerID).MD_file
    End If
End If

'nothing to export
If PCount <= 0 Then
    cnNew.Close
    Me.Caption = strCaptionSaved
    Exit Function
End If

' Initialize cmdPutNewUMC and all of the params in udtPutUMCParams
TraceLog 3, "frmSearchMTPairs->ExportIDPairsToUMCResultsTable", "Call ExportMTDBInitializePutNewUMCParams"
ExportMTDBInitializePutNewUMCParams cnNew, cmdPutNewUMC, udtPutUMCParams, lngMDID, ExpUmcSPName

' Initialize the variables for accessing the AddFTICRUmcMember SP
TraceLog 3, "frmSearchMTPairs->ExportIDPairsToUMCResultsTable", "Call ExportMTDBInitializePutNewUMCParams"
ExportMTDBInitializePutNewUMCMemberParams cnNew, cmdPutNewUMCMember, udtPutUMCMemberParams, ExpUMCMemberSPName

' Initialize the variables for accessing the AddFTICRUmcMatch SP
TraceLog 3, "frmSearchMTPairs->ExportIDPairsToUMCResultsTable", "Call ExportMTDBInitializePutUMCMatchParams"
ExportMTDBInitializePutUMCMatchParams cnNew, cmdPutNewUMCMatch, udtPutUMCMatchParams, ExpUmcMatchSPName

' Initialize the variables for accessing the AddFTICRUmcCSStats SP
ExportMTDBInitializePutUMCCSStatsParams cnNew, cmdPutNewUMCCSStats, udtPutUMCCSStatsParams, ExpUMCCSStats

Me.Caption = "Exporting LC-MS Features to DB: 0 / " & Trim(PCount)

'now export data
MassTagExpCnt = 0

For lngPairInd = 0 To PCount - 1
    If lngPairInd Mod 25 = 0 Then
        Me.Caption = "Exporting pairs to DB: " & Trim(lngPairInd) & " / " & Trim(PCount)
        DoEvents
        TraceLog 3, "frmSearchMTPairs->ExportIDPairsToUMCResultsTable", "Examining lngPairInd = " & Trim(lngPairInd)
    End If
    
    If PIDCnt(lngPairInd) > 0 And _
       (blnExportExcludedPairs Or GelP_D_L(CallerID).Pairs(lngPairInd).STATE <> glPAIR_Exc) Then      'this pair is identified
        IndL = GelP_D_L(CallerID).Pairs(lngPairInd).p1
        IndH = GelP_D_L(CallerID).Pairs(lngPairInd).p2
       
        ' First, add a new row to T_FTICR_UMC_Results for the light member of the pair
        With GelP_D_L(CallerID).Pairs(lngPairInd)
            udtPairMatchStats.PairIndex = lngPairInd
            udtPairMatchStats.ExpressionRatio = .ER
            udtPairMatchStats.ExpressionRatioStDev = .ERStDev
            udtPairMatchStats.ExpressionRatioChargeStateBasisCount = .ERChargeStateBasisCount
            udtPairMatchStats.ExpressionRatioMemberBasisCount = .ERMemberBasisCount
        End With
        
        ExportMTDBAddUMCResultRow cmdPutNewUMC, udtPutUMCParams, cmdPutNewUMCMember, udtPutUMCMemberParams, cmdPutNewUMCCSStats, udtPutUMCCSStatsParams, blnExportUMCMembers, CallerID, IndL, PIDCnt(lngPairInd), ClsStat(), udtPairMatchStats, FPR_Type_N14_N15_L, 0, GelUMC(CallerID).UMCs(IndL).DriftTime
        
        ' Write the match results for this UMC
        udtPutUMCMatchParams.UMCResultsID.Value = FixNullLng(udtPutUMCParams.UMCResultsIDReturn.Value)

        For lngMassTagIndexPointer = PIDInd1(lngPairInd) To PIDInd2(lngPairInd)
            lngMassTagIndexOriginal = mMTOrInd(mMTInd(mUMCMatchStats(lngMassTagIndexPointer).IDIndex))
            
            udtPutUMCMatchParams.MassTagID.Value = AMTData(lngMassTagIndexOriginal).ID
            udtPutUMCMatchParams.MatchingMemberCount.Value = mUMCMatchStats(lngMassTagIndexPointer).MemberHitCount
            udtPutUMCMatchParams.MatchScore.Value = mUMCMatchStats(lngMassTagIndexPointer).StacOrSLiC
            udtPutUMCMatchParams.DelMatchScore.Value = mUMCMatchStats(lngMassTagIndexPointer).DelScore
            udtPutUMCMatchParams.UniquenessProbability.Value = CSqlReal(mUMCMatchStats(lngMassTagIndexPointer).UniquenessProbability)
            
            strMassMods = MOD_TKN_PAIR_LIGHT
            If Len(mMTMods(mMTInd(mUMCMatchStats(lngMassTagIndexPointer).IDIndex))) > 0 Then
                strMassMods = strMassMods & " " & Trim(mMTMods(mMTInd(mUMCMatchStats(lngMassTagIndexPointer).IDIndex)))
                udtPutUMCMatchParams.MassTagModMass.Value = CSqlReal(mMTMWN14(mUMCMatchStats(lngMassTagIndexPointer).IDIndex) - AMTData(lngMassTagIndexOriginal).MW)
            Else
                udtPutUMCMatchParams.MassTagModMass.Value = 0
            End If
            If Len(strMassMods) > PUT_UMC_MATCH_MAX_MODSTRING_LENGTH Then strMassMods = Left(strMassMods, PUT_UMC_MATCH_MAX_MODSTRING_LENGTH)
            udtPutUMCMatchParams.MassTagMods.Value = strMassMods
            
            cmdPutNewUMCMatch.Execute
            MassTagExpCnt = MassTagExpCnt + 1
        Next lngMassTagIndexPointer
        
        ' Second, add a new row to T_FTICR_UMC_Results for the heavy member of the pair
        ' Note that we do not record any MT tag hits for the heavy member of the pair
        ExportMTDBAddUMCResultRow cmdPutNewUMC, udtPutUMCParams, cmdPutNewUMCMember, udtPutUMCMemberParams, cmdPutNewUMCCSStats, udtPutUMCCSStatsParams, blnExportUMCMembers, CallerID, IndH, 0, ClsStat(), udtPairMatchStats, FPR_Type_N14_N15_H, 0, GelUMC(CallerID).UMCs(IndH).DriftTime
        
    End If
Next lngPairInd

' MonroeMod
AddToAnalysisHistory CallerID, "Export to LC-MS Feature Results table details: MT tags Match Count = " & MassTagExpCnt

Me.Caption = strCaptionSaved

strExportStatus = MassTagExpCnt & " associations between MT tags and LC-MS Features exported."
Set cmdPutNewUMC.ActiveConnection = Nothing
Set cmdPutNewUMCMatch.ActiveConnection = Nothing
cnNew.Close

If blnUpdateGANETForAnalysisInDB Then
    ' Export the the GANET Slope, Intercept, and Fit to the database
    TraceLog 5, "frmSearchMTPairs->ExportIDPairsToUMCResultsTable", "Call ExportGANETtoMTDB"
    With GelAnalysis(CallerID)
        strExportStatus = strExportStatus & vbCrLf & ExportGANETtoMTDB(CallerID, .GANET_Slope, .GANET_Intercept, .GANET_Fit)
    End With
End If

If blnAddQuantitationDescriptionEntry Then
    TraceLog 5, "frmSearchMTPairs->ExportIDPairsToUMCResultsTable", "Call ExportMTDBAddQuantitationDescriptionEntry"
    If lngErrorNumber = 0 And lngMDID >= 0 And (MassTagExpCnt > 0) Then
        ExportMTDBAddQuantitationDescriptionEntry Me, CallerID, ExpQuantitationDescription, lngMDID, lngErrorNumber, strIniFileName, 1, 1, 1, Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled
    End If
End If

ExportIDPairsToUMCResultsTable = strExportStatus
lngErrorNumber = 0
Exit Function

err_ExportMTDBbyUMC:
TraceLog 5, "frmSearchMTPairs->ExportIDPairsToUMCResultsTable", "Error occurred: " & Err.Description
Debug.Assert False

LogErrors Err.Number, "ExportIDPairsToUMCResultsTable (Job " & GelAnalysis(CallerID).MD_Reference_Job & ", MD_ID " & lngMDID & ")"
If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
    MsgBox "Error exporting matches to the LC-MS Feature results table: " & Err.Description, vbExclamation + vbOKOnly, glFGTU
Else
    AddToAnalysisHistory CallerID, "Error exporting to LC-MS Feature Results table (occurred at " & lngPairInd & "/" & PCount & "; MDID is " & lngMDID & "): " & Err.Description
End If

err_Cleanup:
On Error Resume Next
If Not cnNew Is Nothing Then cnNew.Close
Me.Caption = strCaptionSaved

If Err.Number <> 0 Then lngErrorNumber = Err.Number
ExportIDPairsToUMCResultsTable = "Error: " & lngErrorNumber & vbCrLf & Err.Description

End Function

Private Function GetDBSearchModeType() As Byte
    If optDBSearchModType(MODS_DECOY).Value Then
        GetDBSearchModeType = 2
    ElseIf optDBSearchModType(MODS_DYNAMIC).Value Then
        GetDBSearchModeType = 1
    Else
        ' Assume mode MODS_FIXED mods
        GetDBSearchModeType = 0
    End If
End Function

Private Function GetWobbledNET(ByVal dblNET As Double, ByVal dblNETWobbleDistance As Double) As Double
    If Rnd() < 0.5 Then
        GetWobbledNET = dblNET - dblNETWobbleDistance
    Else
        GetWobbledNET = dblNET + dblNETWobbleDistance
    End If
End Function

Private Function InitExprEvaluator(ByVal sExpr As String) As Boolean
'-------------------------------------------------------------------
'initializes expression evaluator for elution time
'-------------------------------------------------------------------
On Error Resume Next
Set MyExprEva = New ExprEvaluator
With MyExprEva
    .Vars.add 1, "FN"
    .Vars.add 2, "MinFN"
    .Vars.add 3, "MaxFN"
    .Expr = sExpr
    InitExprEvaluator = .IsExprValid
    ReDim VarVals(1 To 3)
End With
End Function

Private Function InitializeORFInfo(blnForceDataReload As Boolean) As Boolean
    ' Initializes objMTDBNameLookupClass
    ' Returns True if success, False if failure
    ' If the class has already been initialized, then does nothing, unless blnForceDataReload = True
    
    Dim blnSuccess As Boolean
    
    If Not objMTDBNameLookupClass Is Nothing Then
        If Not blnForceDataReload Then
            If objMTDBNameLookupClass.DataStatus = dsLoaded Then
                InitializeORFInfo = True
                Exit Function
            End If
        End If
        
        objMTDBNameLookupClass.DeleteData
        Set objMTDBNameLookupClass = Nothing
    End If
    
    Set objMTDBNameLookupClass = New mtdbMTNames
    
    With objMTDBNameLookupClass
        'loading protein names
        UpdateStatus "Loading Protein info"
        
        If Not GelAnalysis(CallerID) Is Nothing Then
            If Len(GelAnalysis(CallerID).MTDB.cn.ConnectionString) > 0 And Not APP_BUILD_DISABLE_MTS Then
                Me.MousePointer = vbHourglass
                .DBConnectionString = GelAnalysis(CallerID).MTDB.cn.ConnectionString
                .RetrieveSQL = glbPreferencesExpanded.MTSConnectionInfo.sqlGetMTNames
                If .FillData(Me) Then
                   If .DataStatus = dsLoaded Then
                        blnSuccess = True
                    End If
                End If
                Me.MousePointer = vbDefault
            End If
        End If
    End With
    
    InitializeORFInfo = blnSuccess
End Function

Public Sub InitializeSearch()
'------------------------------------------------------------
'load MT tag database data if neccessary
'if CallerID is associated with MT tag database load that
'database if neccessary; if CallerID is not associated with
'MT tag database load legacy database
'------------------------------------------------------------
Dim eResponse As VbMsgBoxResult

On Error Resume Next
Me.MousePointer = vbHourglass
If bLoading Then
    If GelAnalysis(CallerID) Is Nothing Then
        If AMTCnt > 0 Then    'something is loaded
          If (Len(CurrMTDatabase) > 0 Or Len(CurrLegacyMTDatabase) > 0) And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                'MT tag data; we dont know is it appropriate; warn user
                WarnUserUnknownMassTags CallerID
           End If
           lblMTStatus.Caption = ConstructMTStatusText(True)
        
           ' Initialize the MT search object
           If Not CreateNewMTSearchObject() Then
                lblMTStatus.Caption = "Error creating search object."
            Else
               ' Error initializing MT search object
           End If
        
        Else                  'nothing is loaded
            If Len(GelData(CallerID).PathtoDatabase) > 0 And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                If APP_BUILD_DISABLE_MTS Then
                    eResponse = vbYes
                Else
                    eResponse = MsgBox("Current display is not associated with any MT tag database.  Do you want to load the MT tags from the defined legacy MT tag database?" & vbCrLf & GelData(CallerID).PathtoDatabase, vbQuestion + vbYesNoCancel + vbDefaultButton1, "Load Legacy MT tags")
                End If
            Else
                eResponse = vbNo
            End If
            
            If eResponse = vbYes Then
                LoadLegacyMassTags
            Else
                WarnUserNotConnectedToDB CallerID, True
                lblMTStatus.Caption = "No MT tags loaded"
            End If
        End If
    Else         'have to have MT tag database loaded
        Call LoadMTDB
    End If
    UpdateStatus "Generating LC-MS Feature statistics ..."
    ClsCnt = UMCStatistics1(CallerID, ClsStat())
    PCount = GelP_D_L(CallerID).PCnt
    UpdateStatus "Potential Pairs: " & PCount
   
    txtNETFormula.Enabled = Not GelData(CallerID).CustomNETsDefined
    lblNETFormula.Enabled = txtNETFormula.Enabled
    mnuETHeader.Enabled = txtNETFormula.Enabled
    
    Call mnuET_Click(etGANET)
   
    'memorize number of scans for Caller(to be used with elution)
    MinFN = GelData(CallerID).ScanInfo(1).ScanNumber
    MaxFN = GelData(CallerID).ScanInfo(UBound(GelData(CallerID).ScanInfo)).ScanNumber
    
    bLoading = False
End If
Me.MousePointer = vbDefault
End Sub

Private Function IsValidMatch(CurrMW As Double, AbsMWErr As Double, CurrScan As Long, lngMassTagIndexOriginal As Long, dblAMTMass As Double) As Boolean
    ' Checks if CurrMW is within tolerance of the given MT tag
    ' Also checks if the NET equivalent of CurrScan is within tolerance of the NET value for the given MT tag
    ' Returns True if both are within tolerance, false otherwise
    
    Dim InvalidMatch As Boolean
    
    ' If CurrMW is not within AbsMWErr of dblAMTMass then this match is inherited
    If Abs(CurrMW - dblAMTMass) > AbsMWErr Then
        InvalidMatch = True
    Else
        ' If CurrScan is not within .NETTol of mMTNET() then this match is inherited
        If samtDef.NETTol >= 0 Then
            If Abs(ConvertScanToNET(CurrScan) - AMTData(lngMassTagIndexOriginal).NET) > samtDef.NETTol Then
                InvalidMatch = True
            End If
        End If
    End If
    
    IsValidMatch = Not InvalidMatch
End Function

Private Sub LoadLegacyMassTags()

    '------------------------------------------------------------
    'load/reload MT tags
    '------------------------------------------------------------
    Dim eResponse As VbMsgBoxResult
    On Error Resume Next
    'ask user if it wants to replace legitimate MT tag DB with legacy DB
    If Not GelAnalysis(CallerID) Is Nothing And Not APP_BUILD_DISABLE_MTS Then
       eResponse = MsgBox("Current display is associated with MT tag database." & vbCrLf _
                    & "Are you sure you want to use a legacy database for search?", vbYesNoCancel, glFGTU)
       If eResponse <> vbYes Then Exit Sub
    End If
    Me.MousePointer = vbHourglass
    If Len(GelData(CallerID).PathtoDatabase) > 0 Then
       If ConnectToLegacyAMTDB(Me, CallerID, False, True, False) Then
          If CreateNewMTSearchObject() Then
             lblMTStatus.Caption = "Loaded; MT tag count: " & LongToStringWithCommas(AMTCnt)
          Else
             lblMTStatus.Caption = "Error creating search object."
          End If
       Else
          lblMTStatus.Caption = "Error loading MT tags."
       End If
    Else
        WarnUserInvalidLegacyDBPath
    End If
    Me.MousePointer = vbDefault

End Sub

Private Sub LoadMTDB(Optional blnForceReload As Boolean = False)
    Dim blnAMTsWereLoaded As Boolean, blnDBConnectionError As Boolean
    Dim strMessage As String
    
    Static blnWorking As Boolean
    
    If blnWorking Then Exit Sub
    blnWorking = True
    
    If ConfirmMassTagsAndInternalStdsLoaded(Me, CallerID, True, True, False, blnForceReload, 0, blnAMTsWereLoaded, blnDBConnectionError) Then
        lblMTStatus.Caption = ConstructMTStatusText(True)
    
        If Not CreateNewMTSearchObject() Then
           lblMTStatus.Caption = "Error creating search object."
        End If
    Else
        If blnDBConnectionError Then
            strMessage = "Error loading MT tags: database connection error."
        Else
            If Not GelAnalysis(CallerID) Is Nothing Then
                If Len(GelAnalysis(CallerID).MTDB.cn.ConnectionString) > 0 And Not APP_BUILD_DISABLE_MTS Then
                    strMessage = "Error loading MT tags: no valid MT tags were found (possibly missing NET values)"
                ElseIf Len(GelData(CallerID).PathtoDatabase) > 0 Then
                    strMessage = "Error loading MT tags from Legacy DB"
                Else
                    strMessage = "Error loading MT tags: MT tag database not defined"
                End If
            Else
                strMessage = "Error loading MT tags: MT tag database not defined"
            End If
        End If
    
        lblMTStatus.Caption = strMessage
    End If
    
    blnWorking = False
    
End Sub

Private Function ManageCurrID(ByVal ManageType As Long) As Boolean
On Error GoTo exit_ManageCurrID
Select Case ManageType
Case MNG_ERASE
     mCurrIDCnt = 0
     Erase mCurrIDMatches
Case MNG_TRIM
     If mCurrIDCnt > 0 Then
        ReDim Preserve mCurrIDMatches(mCurrIDCnt - 1)
     Else
        ManageCurrID = ManageCurrID(MNG_ERASE)
     End If
Case MNG_RESET
     mCurrIDCnt = 0
     ReDim mCurrIDMatches(MNG_START_SIZE)
Case MNG_ADD_START_SIZE
     ReDim Preserve mCurrIDMatches(mCurrIDCnt + MNG_START_SIZE)
Case Else
     If ManageType > 0 Then
        ReDim Preserve mCurrIDMatches(mCurrIDCnt + ManageType)
     End If
End Select
ManageCurrID = True
exit_ManageCurrID:
End Function

Private Sub PickParameters()
Call txtMWTol_LostFocus
Call txtNETTol_LostFocus
Call txtNETFormula_LostFocus
End Sub

Private Sub PopulateComboBoxes()
    Dim intIndex As Integer
    
On Error GoTo PopulateComboBoxesErrorHandler

    With cboResidueToModify
        .Clear
        .AddItem "Full MT"
        For intIndex = 0 To 25
            .AddItem Chr(vbKeyA + intIndex)
        Next intIndex
        .AddItem glPHOSPHORYLATION
        .ListIndex = 0
    End With
    
    With cboSearchRegionShape
        .Clear
        .AddItem "Elliptical search region"
        .AddItem "Rectangular search region"
        .ListIndex = srsSearchRegionShapeConstants.srsElliptical
    End With
    
    Exit Sub
    
PopulateComboBoxesErrorHandler:
    LogErrors Err.Number, "frmSearchMTPairs.frm->PopulateComboBoxes"
End Sub

Private Function PrepareMTArrays() As Boolean
'---------------------------------------------------------------
'prepares masses from loaded MT tags based on specified
'modifications; returns True if succesful, False on any error
'---------------------------------------------------------------
Dim I As Long, j As Long
Dim TmpCnt As Long
Dim CysCnt As Long                 'Cysteine count in peptide

Dim strResiduesToModify As String   ' One or more residues to modify (single letter amino acid symbols)
Dim dblResidueModMass As Double
Dim ResidueOccurrenceCount As Integer
Dim strResModToken As String
Dim blnAddMassTag As Boolean

Dim dblNETWobbleDistance As Double

On Error GoTo err_PrepareMTArrays

' Update GelSearchDef(CallerID).AMTSearchMassMods with the current settings
With GelSearchDef(CallerID).AMTSearchMassMods
    .PEO = False
    .ICATd0 = False
    .ICATd8 = False
    .Alkylation = cChkBox(chkAlkylation)
    .AlkylationMass = CDblSafe(txtAlkylationMWCorrection)
    If cboResidueToModify.ListIndex > 0 Then
        .ResidueToModify = cboResidueToModify
    Else
        .ResidueToModify = ""
    End If
    
    .ResidueMassModification = CDblSafe(txtResidueToModifyMass)
    txtResidueToModifyMass = Round(.ResidueMassModification, 5)
    
    strResiduesToModify = .ResidueToModify
    dblResidueModMass = .ResidueMassModification
    
    .N15InsteadOfN14 = False
        
    ' Superseded by .ModMode in August 2008
    '.DynamicMods = optDBSearchModType(MODS_DYNAMIC).Value
        
    .ModMode = GetDBSearchModeType()
End With

If IsNumeric(txtDBSearchMinimumHighNormalizedScore.Text) Then
    mMTMinimumHighNormalizedScore = CSngSafe(txtDBSearchMinimumHighNormalizedScore.Text)
Else
    mMTMinimumHighNormalizedScore = 0
End If
    
If IsNumeric(txtDBSearchMinimumHighDiscriminantScore.Text) Then
    mMTMinimumHighDiscriminantScore = CSngSafe(txtDBSearchMinimumHighDiscriminantScore.Text)
Else
    mMTMinimumHighDiscriminantScore = 0
End If

If IsNumeric(txtDBSearchMinimumPeptideProphetProbability.Text) Then
    mMTMinimumPeptideProphetProbability = CSngSafe(txtDBSearchMinimumPeptideProphetProbability.Text)
Else
    mMTMinimumPeptideProphetProbability = 0
End If

If mMTMinimumHighNormalizedScore > 0 Or mMTMinimumHighDiscriminantScore > 0 Or mMTMinimumPeptideProphetProbability > 0 Then
        If mMTMinimumHighDiscriminantScore > 0 Or mMTMinimumPeptideProphetProbability > 0 Then
        ' Make sure at least two of the loaded MT tags have score values >= mMTMinimumHighDiscriminantScore, also taking into account HighNormalizedScore
        ValidateMTMinimumDiscriminantAndPepProphet AMTData(), 1, AMTCnt, mMTMinimumHighDiscriminantScore, mMTMinimumPeptideProphetProbability, mMTMinimumHighNormalizedScore, 2
    Else
        ' Make sure at least two of the loaded MT tags have score values >= mMTMinimumHighNormalizedScore
        ValidateMTMinimimumHighNormalizedScore AMTData(), 1, AMTCnt, mMTMinimumHighNormalizedScore, 2
    End If
End If

If Not IsNumeric(txtDecoySearchNETWobble.Text) Then
    txtDecoySearchNETWobble.Text = 0.1
End If
dblNETWobbleDistance = CSngSafe(txtDecoySearchNETWobble.Text)

    If AMTCnt <= 0 Then
        mMTCnt = 0
    Else
   UpdateStatus "Preparing arrays for search..."
   
   If GelSearchDef(CallerID).AMTSearchMassMods.ModMode = 2 Then
       Randomize NET_WOBBLE_SEED
   End If
  
   'initially reserve space for AMTCnt peptides
   ReDim mMTInd(AMTCnt - 1)
   ReDim mMTOrInd(AMTCnt - 1)
   ReDim mMTMWN14(AMTCnt - 1)
   ReDim mMTNET(AMTCnt - 1)
   ReDim mMTMods(AMTCnt - 1)
   mMTCnt = 0
   For I = 1 To AMTCnt
        If mMTMinimumHighNormalizedScore > 0 Or mMTMinimumHighDiscriminantScore > 0 Or mMTMinimumPeptideProphetProbability > 0 Then
            If AMTData(I).HighNormalizedScore >= mMTMinimumHighNormalizedScore And _
               AMTData(I).HighDiscriminantScore >= mMTMinimumHighDiscriminantScore And _
               AMTData(I).PeptideProphetProbability >= mMTMinimumPeptideProphetProbability Then
                blnAddMassTag = True
            Else
                blnAddMassTag = False
            End If
        Else
            blnAddMassTag = True
        End If
        
        If blnAddMassTag Then
            mMTCnt = mMTCnt + 1
            mMTInd(mMTCnt - 1) = mMTCnt - 1
            mMTOrInd(mMTCnt - 1) = I             'index; not the ID
            mMTMWN14(mMTCnt - 1) = AMTData(I).MW
            Select Case samtDef.NETorRT
            Case glAMT_NET
                 mMTNET(mMTCnt - 1) = AMTData(I).NET
            Case glAMT_RT_or_PNET
                 mMTNET(mMTCnt - 1) = AMTData(I).PNET
            End Select
            mMTMods(mMTCnt - 1) = ""
        End If
   Next I
   
   If chkAlkylation.Value = vbChecked Then         'correct based on cys number for alkylation label
      UpdateStatus "Adding alkylated peptides..."
      TmpCnt = mMTCnt
      For I = 0 To TmpCnt - 1
          CysCnt = AMTData(mMTOrInd(I)).CNT_Cys
          If CysCnt > 0 Then
             If GelSearchDef(CallerID).AMTSearchMassMods.ModMode = 1 Or _
                GelSearchDef(CallerID).AMTSearchMassMods.ModMode = 2 Then
                
                ' Dynamic Mods
                For j = 1 To CysCnt
                    mMTCnt = mMTCnt + 1
                    mMTInd(mMTCnt - 1) = mMTCnt - 1
                    mMTOrInd(mMTCnt - 1) = mMTOrInd(I)
                    mMTMWN14(mMTCnt - 1) = mMTMWN14(I) + j * AlkMWCorrection
                    
                    If GelSearchDef(CallerID).AMTSearchMassMods.ModMode = 2 Then
                        mMTNET(mMTCnt - 1) = GetWobbledNET(mMTNET(I), dblNETWobbleDistance)
                    Else
                        mMTNET(mMTCnt - 1) = mMTNET(I)
                    End If
                    
                    mMTMods(mMTCnt - 1) = mMTMods(I) & " " & MOD_TKN_ALK & "/" & j
                Next j
             Else
                ' Static Mods
                ' Simply update the stats for this MT tag
                mMTMWN14(I) = mMTMWN14(I) + CysCnt * AlkMWCorrection
                mMTMods(I) = mMTMods(I) & " " & MOD_TKN_ALK & "/" & CysCnt
             End If
          End If
      Next I
   End If
   
       If dblResidueModMass <> 0 Or GelSearchDef(CallerID).AMTSearchMassMods.ModMode = 2 Then
      UpdateStatus "Adding modified residue mass peptides..."
      TmpCnt = mMTCnt
      For I = 0 To TmpCnt - 1
            
          If Len(strResiduesToModify) > 0 Then
            ResidueOccurrenceCount = LookupResidueOccurrence(mMTOrInd(I), strResiduesToModify)
            strResModToken = MOD_TKN_RES_MOD
          Else
            ' Add dblResidueModMass once to the entire MT tag
            ' Accomplish this by setting ResidueOccurrenceCount to 1
            ResidueOccurrenceCount = 1
            strResModToken = MOD_TKN_MT_MOD
          End If
          
          If ResidueOccurrenceCount > 0 Then
             If GelSearchDef(CallerID).AMTSearchMassMods.ModMode = 1 Or _
                GelSearchDef(CallerID).AMTSearchMassMods.ModMode = 2 Then
               
                ' Dynamic Mods
                For j = 1 To ResidueOccurrenceCount
                    mMTCnt = mMTCnt + 1
                    mMTInd(mMTCnt - 1) = mMTCnt - 1
                    mMTOrInd(mMTCnt - 1) = mMTOrInd(I)
                    mMTMWN14(mMTCnt - 1) = mMTMWN14(I) + j * dblResidueModMass
                    
                    If GelSearchDef(CallerID).AMTSearchMassMods.ModMode = 2 Then
                        mMTNET(mMTCnt - 1) = GetWobbledNET(mMTNET(I), dblNETWobbleDistance)
                    Else
                        mMTNET(mMTCnt - 1) = mMTNET(I)
                    End If
                    
                    mMTMods(mMTCnt - 1) = mMTMods(I) & " " & strResModToken & "/" & strResiduesToModify & j
                Next j
             Else
                ' Static Mods
                ' Simply update the stats for this MT tag
                mMTMWN14(I) = mMTMWN14(I) + ResidueOccurrenceCount * dblResidueModMass
                mMTMods(I) = mMTMods(I) & " " & strResModToken & "/" & strResiduesToModify & ResidueOccurrenceCount
             End If
          End If
      Next I
   End If
   
   If mMTCnt > 0 Then
      UpdateStatus "Preparing fast search structures..."
      ReDim Preserve mMTInd(mMTCnt - 1)
      ReDim Preserve mMTOrInd(mMTCnt - 1)
      ReDim Preserve mMTMWN14(mMTCnt - 1)
      ReDim Preserve mMTNET(mMTCnt - 1)
      ReDim Preserve mMTMods(mMTCnt - 1)
      If Not PrepareSearchN14() Then
         Debug.Assert False
         Call DestroySearchStructures
         Exit Function
      End If
   Else
      Call DestroySearchStructures
   End If
   
End If
PrepareMTArrays = True
Exit Function

err_PrepareMTArrays:
Select Case Err.Number
Case 9                      'add space in chunks of 10000
   ReDim Preserve mMTInd(mMTCnt + 10000)
   ReDim Preserve mMTOrInd(mMTCnt + 10000)
   ReDim Preserve mMTMWN14(mMTCnt + 10000)
   ReDim Preserve mMTNET(mMTCnt + 10000)
   ReDim Preserve mMTMods(mMTCnt + 10000)
   Resume
Case Else
   Debug.Assert False
   Call DestroySearchStructures
End Select

End Function

Private Function PrepareSearchN14() As Boolean
'---------------------------------------------------------------
'prepare search of N14 peptide (use loaded peptides masses)
'---------------------------------------------------------------
On Error Resume Next
If mMTCnt > 0 Then
   UpdateStatus "Preparing fast N14 search..."
   ' Dim qsd As New QSDouble
   ' Old: If qsd.QSAsc(mMTMWN14(), mMTInd()) Then
   If ShellSortDoubleWithParallelLong(mMTMWN14(), mMTInd(), 0, UBound(mMTMWN14)) Then
      Set MWFastSearch = New MWUtil
      If MWFastSearch.Fill(mMTMWN14()) Then PrepareSearchN14 = True
   End If
End If
End Function

Private Sub RecordSearchResultsInData()
    ' Step through mUMCMatchStats() and add the ID's for each UMC to all of the members of each UMC
    
    Dim lngIndex As Long, lngMemberIndex As Long
    Dim lngPairIndexOriginal As Long
    Dim lngUMCIndexOriginal As Long
    Dim lngMassTagIndexOriginal As Long                 'absolute index in AMT... arrays
    Dim lngIonIndexOriginal As Long
    Dim blnAddAMTRef As Boolean
    Dim lngIonCountUpdated As Long
    
    Dim AMTRef As String
    Dim dblAMTMass As Double
    Dim dblStacOrSLiC As Double
    Dim dblDelSLiC As Double
    Dim dblUPScore As Double
    
    Dim CurrMW As Double, AbsMWErr As Double
    Dim CurrScan As Long
     
    ' Need to remove any existing search results before adding these new ones
    RemoveAMT CallerID, glScope.glSc_All
    GelStatus(CallerID).Dirty = True
    AddToAnalysisHistory CallerID, "Deleted MT tag search results from ions"
    
    'always reinitialize statistics arrays
    InitAMTStat
    
    KeyPressAbortProcess = 0
    
    CheckNETEquationStatus
    
On Error GoTo RecordSearchResultsInDataErrorHandler

    With GelData(CallerID)
        For lngIndex = 0 To mMatchStatsCount - 1
            If lngIndex Mod 25 = 0 Then
                UpdateStatus "Storing results: " & LongToStringWithCommas(lngIndex) & " / " & LongToStringWithCommas(mMatchStatsCount)
                If KeyPressAbortProcess > 1 Then Exit For
            End If
            
            lngPairIndexOriginal = mUMCMatchStats(lngIndex).PairIndex
            lngUMCIndexOriginal = GelP_D_L(CallerID).Pairs(lngPairIndexOriginal).p1
            
            lngMassTagIndexOriginal = mMTOrInd(mMTInd(mUMCMatchStats(lngIndex).IDIndex))
            
            dblAMTMass = mMTMWN14(mUMCMatchStats(lngIndex).IDIndex)
            dblStacOrSLiC = mUMCMatchStats(lngIndex).StacOrSLiC
            dblDelSLiC = mUMCMatchStats(lngIndex).DelScore
            dblUPScore = mUMCMatchStats(lngIndex).UniquenessProbability
            
            For lngMemberIndex = 0 To GelUMC(CallerID).UMCs(lngUMCIndexOriginal).ClassCount - 1
                lngIonIndexOriginal = GelUMC(CallerID).UMCs(lngUMCIndexOriginal).ClassMInd(lngMemberIndex)
                blnAddAMTRef = False
                
                Select Case GelUMC(CallerID).UMCs(lngUMCIndexOriginal).ClassMType(lngMemberIndex)
                Case glCSType
                    CurrMW = .CSData(lngIonIndexOriginal).AverageMW
                    CurrScan = .CSData(lngIonIndexOriginal).ScanNumber
                    Select Case samtDef.TolType
                    Case gltPPM
                       AbsMWErr = CurrMW * samtDef.MWTol * glPPM
                    Case gltABS
                       AbsMWErr = samtDef.MWTol
                    Case Else
                       Debug.Assert False
                    End Select
                    
                    AMTRef = ConstructAMTReference(.CSData(lngIonIndexOriginal).AverageMW, ConvertScanToNET(.CSData(lngIonIndexOriginal).ScanNumber), 0, lngMassTagIndexOriginal, dblAMTMass, dblStacOrSLiC, dblDelSLiC, dblUPScore, False, False, 0)
                    If Len(.CSData(lngIonIndexOriginal).MTID) = 0 Then
                        blnAddAMTRef = True
                    ElseIf InStr(.CSData(lngIonIndexOriginal).MTID, AMTRef) <= 0 Then
                        blnAddAMTRef = True
                    End If

                    If blnAddAMTRef Then
                        lngIonCountUpdated = lngIonCountUpdated + 1
                        
                        If Not IsValidMatch(CurrMW, AbsMWErr, CurrScan, lngMassTagIndexOriginal, dblAMTMass) Then
                            AMTRef = Trim(AMTRef)
                            If Right(AMTRef, 1) = glARG_SEP Then
                                AMTRef = Left(AMTRef, Len(AMTRef) - 1)
                            End If
                            AMTRef = AMTRef & AMTMatchInheritedMark
                        End If
                        
                        InsertBefore .CSData(lngIonIndexOriginal).MTID, AMTRef
                    End If
                Case glIsoType
                    CurrMW = GetIsoMass(.IsoData(lngIonIndexOriginal), samtDef.MWField)
                    CurrScan = .IsoData(lngIonIndexOriginal).ScanNumber
                    Select Case samtDef.TolType
                    Case gltPPM
                       AbsMWErr = CurrMW * samtDef.MWTol * glPPM
                    Case gltABS
                       AbsMWErr = samtDef.MWTol
                    Case Else
                       Debug.Assert False
                    End Select

                    AMTRef = ConstructAMTReference(GetIsoMass(.IsoData(lngIonIndexOriginal), samtDef.MWField), ConvertScanToNET(.IsoData(lngIonIndexOriginal).ScanNumber), 0, lngMassTagIndexOriginal, dblAMTMass, dblStacOrSLiC, dblDelSLiC, 0, False, False, 0)
                    If Len(.IsoData(lngIonIndexOriginal).MTID) = 0 Then
                        blnAddAMTRef = True
                    ElseIf InStr(.IsoData(lngIonIndexOriginal).MTID, AMTRef) <= 0 Then
                        blnAddAMTRef = True
                    End If
                    
                    If blnAddAMTRef Then
                        lngIonCountUpdated = lngIonCountUpdated + 1
                        
                        If Not IsValidMatch(CurrMW, AbsMWErr, CurrScan, lngMassTagIndexOriginal, dblAMTMass) Then
                            AMTRef = Trim(AMTRef)
                            If Right(AMTRef, 1) = glARG_SEP Then
                                AMTRef = Left(AMTRef, Len(AMTRef) - 1)
                            End If
                            AMTRef = AMTRef & AMTMatchInheritedMark
                        End If
                        
                        InsertBefore .IsoData(lngIonIndexOriginal).MTID, AMTRef
                    End If
                End Select
            Next lngMemberIndex
        Next lngIndex
    End With
    
    If KeyPressAbortProcess <= 1 Then
        AddToAnalysisHistory CallerID, "Stored search results in ions; recorded all MT tag hits for each LC-MS Feature in all members of the UMC; total ions updated = " & Trim(lngIonCountUpdated)
    End If
    
    Exit Sub

RecordSearchResultsInDataErrorHandler:
    LogErrors Err.Number, "frmSearchMTPairs->RecordSearchResultsInData"
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "An error has occured while storing the search results in the data: " & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
End Sub

Private Sub SearchPairSingleMass(ByVal PairInd As Long)
'---------------------------------------------------------------
'finds all matching identifications for pair with index PairInd
'Search all class members of the light pair member for matching
'MT tags; criteria includes molecular mass, elution time and
'number of N atoms - compared with number of pair deltas
'NOTE: MWs over 2000(4000) Da allow for N count error of +/-1(2)
'---------------------------------------------------------------
Dim ClsInd1 As Long         'class index of light pair member
Dim DltCnt As Long          'established delta count for this pair

Dim lngIndex As Long
Dim lngMassTagIndexPointer As Long

Dim MWTolAbsBroad As Double     ' MWTol used to compute the MatchScore
Dim NETTolBroad As Double       ' NETTol used to compute the MatchScore

Dim MWTolAbsFinal As Double     ' Final MWErr required
Dim NETTolFinal As Double

Dim dblClassMass As Double

Dim blnUsingPrecomputedSLiCScores As Boolean
Dim blnFilterUsingFinalTolerances As Boolean

On Error GoTo err_SearchPairSingleMass

'couple of shortcut variables
ClsInd1 = GelP_D_L(CallerID).Pairs(PairInd).p1
DltCnt = GelP_D_L(CallerID).Pairs(PairInd).P2DltCnt

If ManageCurrID(MNG_RESET) Then
    ' Define the tolerances
    SearchAMTDefineTolerances CallerID, ClsInd1, samtDef, dblClassMass, MWTolAbsBroad, NETTolBroad, MWTolAbsFinal, NETTolFinal
    
    ' First search for the MT tags using broad tolerances
    SearchPairConglomerateMassAMT GelUMC(CallerID).UMCs(ClsInd1), DltCnt, MWTolAbsBroad, NETTolBroad
    
    ' Populate .IDIndexOriginal
    For lngIndex = 0 To mCurrIDCnt - 1
        lngMassTagIndexPointer = mMTInd(mCurrIDMatches(lngIndex).IDInd)
        mCurrIDMatches(lngIndex).IDIndexOriginal = mMTOrInd(lngMassTagIndexPointer)
    Next lngIndex
    

    blnUsingPrecomputedSLiCScores = False
    blnFilterUsingFinalTolerances = True
    
    ' Next compute the Match Scores
    SearchAMTComputeSLiCScores mCurrIDCnt, mCurrIDMatches, dblClassMass, MWTolAbsFinal, NETTolFinal, mSearchRegionShape, blnUsingPrecomputedSLiCScores, blnFilterUsingFinalTolerances
    
    '-----------------------------------------------------------------
    'all identifications for PairInd are collected; now order them in
    'unique identifications with scores and add it to all possible IDs
    '-----------------------------------------------------------------

    If mCurrIDCnt > 0 Then
        If ManageCurrID(MNG_TRIM) Then
            
            'memorize unique count with each pair
            PIDCnt(PairInd) = mCurrIDCnt
            
            'add unique identifications with scores to all ids, also
            'memorize first and last index of id block for current pair
            If mCurrIDCnt > 0 Then
               If UBound(mUMCMatchStats) < mMatchStatsCount + mCurrIDCnt Then    'add more room
                  'make sure it is enough to accomodate current batch
                  ReDim Preserve mUMCMatchStats(UBound(mUMCMatchStats) + mCurrIDCnt + 2000)
               End If
               
               PIDInd1(PairInd) = mMatchStatsCount            'first index
               'last index will remain -1 if no ids and PIDInd2(i)>=0
               'should always be checked when enumerating ids for pair
               For lngIndex = 0 To mCurrIDCnt - 1
                   mMatchStatsCount = mMatchStatsCount + 1
                   With mUMCMatchStats(mMatchStatsCount - 1)
                     .UMCIndex = -1
                     .PairIndex = PairInd
                     .IDIndex = mCurrIDMatches(lngIndex).IDInd
                     .MemberHitCount = mCurrIDMatches(lngIndex).MatchingMemberCount
                     .StacOrSLiC = mCurrIDMatches(lngIndex).StacOrSLiC
                     .DelScore = mCurrIDMatches(lngIndex).DelScore
                     .UniquenessProbability = 0
                     .FDRThreshold = 1
                     .MultiAMTHitCount = mCurrIDCnt
                     .IDIsInternalStd = False
                   End With
                   PIDInd2(PairInd) = mMatchStatsCount - 1    'last index
               Next lngIndex
            End If
            
        End If
    Else
        Call ManageCurrID(MNG_ERASE)
    End If

End If

Exit Sub

err_SearchPairSingleMass:
LogErrors Err.Number, "frmSearchMTPairs_SearchPairSingleMass"

End Sub

Private Sub SearchPairConglomerateMassAMT(ByRef udtTestUMC As udtUMCType, ByVal DltCnt As Long, ByVal dblMWTol As Double, ByVal dblNETTol As Double)

    ' Compare this Pair's UMC masses, NET, and charge with the MT tags,
    ' considering the delta count and the number of nitrogen atoms in the sequence

    Dim FastSearchMatchInd As Long
    Dim MatchInd1 As Long, MatchInd2 As Long
    Dim lngMemberIndex As Long
    
    Dim dblMassTagNET As Double
    Dim dblMassTagMass As Double
    
    Dim dblCurrMW As Double
    Dim dblNETDifference As Double
    
    Dim blnFirstMatchFound As Boolean
    
    MatchInd1 = 0
    MatchInd2 = -1
    If MWFastSearch.FindIndexRange(udtTestUMC.ClassMW, dblMWTol, MatchInd1, MatchInd2) Then
        If MatchInd1 <= MatchInd2 Then
            ' One or more MT tags is within dblMWTol of the median UMC mass
            
            With udtTestUMC
                
                For FastSearchMatchInd = MatchInd1 To MatchInd2
                    ' See if each MassTag is within the NET tolerance of any of the members of the class
                    ' Alternatively, if .UseUMCConglomerateNET = True, then use the NET value of the class representative
                    
                    ' In addition, for each MassTag within both NET and mass tolerance, increment PairScore
                    
                    ' Note that since we used UMCConglomerateMW in the call to FindIndexRange(), not all members
                    '  of the class will necessarily have a matching mass
                    
                    ' Additionally, it is possible that the conglomerate class mass will match a MT tag, but none
                    ' of the members will match.  An example of this is a UMC with two members, weighing 500.0 and 502.0 Da
                    ' The median mass is 501.0 Da.  If the AbsMWErr = 0.1, then the median will match, but none of the members
                    '  will match.  In this case, we'll record the match, but place a 0 in PairScore()
                    
                    dblMassTagMass = MWFastSearch.GetMWByIndex(FastSearchMatchInd)
                    dblMassTagNET = mMTNET(mMTInd(FastSearchMatchInd))
                    
                    blnFirstMatchFound = False
                    If glbPreferencesExpanded.UseUMCConglomerateNET Then
                        If SearchUMCTestNET(.ClassRepType, .ClassRepInd, dblMassTagNET, dblNETTol, dblNETDifference) Then
                            ' AMT Matches this LC-MS Feature's median mass and Class Rep NET
                            
                            ' See if match has correct number of N atoms
                            If CheckNAtomsVsDeltaCount(.ClassMW, DltCnt, mMTOrInd(mMTInd(FastSearchMatchInd))) Then
                                                
                                ' AMT Matches this LC-MS Feature's median mass, Class Rep NET, and has correct
                                '  number of N atoms; increment mCurrIDCnt
                                
                                If mCurrIDCnt > UBound(mCurrIDMatches) Then ManageCurrID (MNG_ADD_START_SIZE)
                                
                                mCurrIDMatches(mCurrIDCnt).IDInd = FastSearchMatchInd
                                mCurrIDMatches(mCurrIDCnt).MatchingMemberCount = 0
                                mCurrIDMatches(mCurrIDCnt).StacOrSLiC = -1    ' Set this to -1 for now
                                mCurrIDMatches(mCurrIDCnt).MassErr = .ClassMW - dblMassTagMass
                                mCurrIDMatches(mCurrIDCnt).NETErr = dblNETDifference
                                mCurrIDMatches(mCurrIDCnt).IDIsInternalStd = False
                                mCurrIDMatches(mCurrIDCnt).UniquenessProbability = 0
                                
                                mCurrIDCnt = mCurrIDCnt + 1
                                
                                blnFirstMatchFound = True
                            End If
                        End If
                    End If
            
                    If blnFirstMatchFound Or Not glbPreferencesExpanded.UseUMCConglomerateNET Then
                        For lngMemberIndex = 0 To .ClassCount - 1
                            If SearchUMCTestNET(CInt(.ClassMType(lngMemberIndex)), .ClassMInd(lngMemberIndex), dblMassTagNET, dblNETTol, dblNETDifference) Then
                            
                                Select Case .ClassMType(lngMemberIndex)
                                Case glCSType
                                    dblCurrMW = GelData(CallerID).CSData(.ClassMInd(lngMemberIndex)).AverageMW
                                Case glIsoType
                                    dblCurrMW = GetIsoMass(GelData(CallerID).IsoData(.ClassMInd(lngMemberIndex)), samtDef.MWField)
                                End Select
                                    
                                ' See if match has correct number of N atoms
                                If CheckNAtomsVsDeltaCount(dblCurrMW, DltCnt, mMTOrInd(mMTInd(FastSearchMatchInd))) Then
                                    If Not blnFirstMatchFound Then
                                        ' We haven't had a match for this index yet; add to mCurrIDMatches()
                                        
                                        If mCurrIDCnt > UBound(mCurrIDMatches) Then ManageCurrID (MNG_ADD_START_SIZE)
                                        
                                        mCurrIDMatches(mCurrIDCnt).IDInd = FastSearchMatchInd
                                        mCurrIDMatches(mCurrIDCnt).MatchingMemberCount = 0
                                        mCurrIDMatches(mCurrIDCnt).StacOrSLiC = -1    ' Set this to -1 for now
                                        mCurrIDMatches(mCurrIDCnt).MassErr = dblCurrMW - dblMassTagMass
                                        mCurrIDMatches(mCurrIDCnt).NETErr = dblNETDifference
                                        mCurrIDMatches(mCurrIDCnt).IDIsInternalStd = False
                                        mCurrIDMatches(mCurrIDCnt).UniquenessProbability = 0
                                        
                                        mCurrIDCnt = mCurrIDCnt + 1
                                        
                                        blnFirstMatchFound = True
                                    End If
                                    
                                    ' See if the member is within mass tolerance
                                    If Abs(dblMassTagMass - dblCurrMW) <= dblMWTol Then
                                        ' Yes, within both mass and NET tolerance; increment mCurrIDMatches().MatchingMemberCount
                                        mCurrIDMatches(mCurrIDCnt - 1).MatchingMemberCount = mCurrIDMatches(mCurrIDCnt - 1).MatchingMemberCount + 1
                                    End If
                                End If
                            End If
                        Next lngMemberIndex
                    End If
                Next FastSearchMatchInd
                
            End With
        End If
    End If

End Sub

Private Function SearchUMCTestNET(eMemberType As glDistType, lngMemberIndex As Long, dblAMTNET As Double, dblNETTol As Double, ByRef dblNETDifference As Double) As Boolean
    
    Dim lngScan As Long
    Dim blnNETMatch As Boolean
    
    Select Case eMemberType
    Case glCSType
        lngScan = GelData(CallerID).CSData(lngMemberIndex).ScanNumber
    Case glIsoType
        lngScan = GelData(CallerID).IsoData(lngMemberIndex).ScanNumber
    End Select
    
    blnNETMatch = False
    dblNETDifference = ConvertScanToNET(lngScan) - dblAMTNET
    If dblNETTol > 0 Then
        If Abs(dblNETDifference) <= dblNETTol Then
            blnNETMatch = True
        End If
    Else
        ' NETTol = 0; assume a match
        blnNETMatch = True
    End If

    SearchUMCTestNET = blnNETMatch
    
End Function

Public Sub SetAlkylationMWCorrection(ByVal dblMass As Double)
    txtAlkylationMWCorrection = dblMass
    AlkMWCorrection = dblMass
End Sub

Private Sub SetDBSearchModType(ByVal bytModMode As Byte)
    If bytModMode = 2 Then
        optDBSearchModType(MODS_DECOY).Value = True
    ElseIf bytModMode = 1 Then
        optDBSearchModType(MODS_DYNAMIC).Value = True
    Else
        ' Assumed fixed
        optDBSearchModType(MODS_FIXED).Value = True
    End If

    GelSearchDef(CallerID).AMTSearchMassMods.ModMode = GetDBSearchModeType()
    
End Sub


Public Sub SetMinimumHighDiscriminantScore(sngMinimumHighDiscriminantScore As Single)
    txtDBSearchMinimumHighDiscriminantScore = sngMinimumHighDiscriminantScore
End Sub

Public Sub SetMinimumHighNormalizedScore(sngMinimumHighNormalizedScore As Single)
    txtDBSearchMinimumHighNormalizedScore = sngMinimumHighNormalizedScore
End Sub

Public Sub SetMinimumPeptideProphetProbability(sngMinimumPeptideProphetProbability As Single)
    txtDBSearchMinimumPeptideProphetProbability = sngMinimumPeptideProphetProbability
End Sub

Private Sub ShowHidePNNLMenus()
    Dim blnVisible As Boolean
    blnVisible = Not APP_BUILD_DISABLE_MTS
    
    mnuExpMTDB.Visible = blnVisible
    mnuPExportDetailedMemberInformation.Visible = blnVisible

    mnuMTLoadMT.Visible = blnVisible
End Sub

Public Function ShowOrSavePairsAndIDs(Optional strOutputFilePath As String = "", Optional blnDisplayResults As Boolean = True, Optional blnIdentifiedPairsOnly As Boolean = False, Optional blnUniquelyIdentifiedOnly As Boolean = False, Optional blnUnidentifiedPairsOnly As Boolean = False, Optional blnShowExcludedPairs As Boolean, Optional ByVal blnIncludeORFInfo As Boolean = True) As Long
    '-------------------------------------
    ' Report pairs and identifications, or report only unidentified pairs if blnUnidentifiedOnly = True
    ' If strOutputFilePath = "", then saves the results to a temporary file and shows them to the user using frmDataInfo
    ' If strOutputFilePath is not blank, then saves the results to the file, but does not display them
    ' If blnIncludeORFInfo = True, then attempts to connect to the database and retrieve the ORF information for each MT tag
    '
    ' Returns 0 if no error, the error number if an error
    '-------------------------------------
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Dim strLineOut As String
    Dim fname As String
    Dim lngPairInd As Long
    Dim lngUMCIndexLight As Long, lngUMCIndexHeavy As Long
    Dim lngMatchIndex As Long
    Dim lngMassTagIndexPointer As Long
    Dim lngMassTagIndexOriginal As Long
    
    Dim objORFNameFastSearch As New FastSearchArrayLong
    Dim blnSuccess As Boolean
    
    Dim strPairInfo As String           ' pair part of line
    Dim strIDInfo As String             ' ID part of line
    Dim strSepChar As String
    
    Dim dblUMCMass As Double
    Dim dblMassErrorPPM As Double
    Dim dblGANETError As Double
    
    
    Dim lngLightScanClassRep As Long
    Dim dblLightNETClassRep As Double
    Dim dblLightDriftTimeClassRep As Double
    Dim dblLightClassRepAbundance As Double
    
    Dim lngHeavyScanClassRep As Long
    Dim dblHeavyNETClassRep As Double
    Dim dblHeavyDriftTimeClassRep As Double
    Dim dblHeavyClassRepAbundance As Double

    Dim blnReportAllPairs As Boolean
    Dim strReportHeader As String
    
    On Error GoTo ShowOrSavePairsAndIDsErrorHandler
    
    If PCount <= 0 And Len(strOutputFilePath) = 0 Then
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "No pairs found.", vbOKOnly, glFGTU
        End If
        Exit Function
    End If
    
    If blnIncludeORFInfo Then
        UpdateStatus "Sorting Protein lookup arrays"
        If MTtoORFMapCount = 0 Then
            blnIncludeORFInfo = InitializeORFInfo(False)
        Else
            ' We can use MTIDMap(), ORFIDMap(), and ORFRefNames() to get the ORF name
            blnSuccess = objORFNameFastSearch.Fill(MTIDMap())
            Debug.Assert blnSuccess
        End If
    End If
    
    blnReportAllPairs = Not (blnIdentifiedPairsOnly Or blnUniquelyIdentifiedOnly Or blnUnidentifiedPairsOnly)
    
    If blnReportAllPairs Then
        strReportHeader = "All pairs and identifications"
    Else
        If blnIdentifiedPairsOnly Or blnUniquelyIdentifiedOnly Then
            If mMatchStatsCount <= 0 And Len(strOutputFilePath) = 0 Then
                MsgBox "No identified pairs found.", vbOKOnly, glFGTU
                Exit Function
            End If
            If blnUniquelyIdentifiedOnly Then
                strReportHeader = "Uniquely identified pairs only"
            Else
                strReportHeader = "Identified pairs only"
            End If
        Else
            If blnUnidentifiedPairsOnly Then
                strReportHeader = "Unidentified pairs only"
            End If
        End If
    End If
    
    UpdateStatus "Preparing results: 0 / " & Trim(PCount)
    mKeyPressAbortProcess = 0
    Me.MousePointer = vbHourglass
    
    'temporary file for results output
    fname = GetTempFolder() & RawDataTmpFile
    If Len(strOutputFilePath) > 0 Then fname = strOutputFilePath
    Set ts = fso.OpenTextFile(fname, ForWriting, True)
    
    strSepChar = LookupDefaultSeparationCharacter()
    
    If Len(strOutputFilePath) = 0 Then
        ts.WriteLine "Generated by: " & GetMyNameVersion() & " on " & Now()
        ts.WriteLine "Gel File: " & GelBody(CallerID).Caption
        ts.WriteLine "Reporting identification for N14/N15 UMC pairs"
    End If
    
    ts.WriteLine strReportHeader
    
    If Len(strOutputFilePath) = 0 Then
        ts.WriteLine
        ts.WriteLine "Total data points: " & GelData(CallerID).DataLines
        ts.WriteLine "Total N14/N15 pairs: " & GelP_D_L(CallerID).PCnt
        ts.WriteLine "Total MT tags: " & AMTCnt
        ts.WriteLine
    End If
                        
    strLineOut = "Pair Index" & strSepChar & "UMC Light Ind" & strSepChar & "Light MW" & strSepChar & "Light Abu" & strSepChar
    strLineOut = strLineOut & "Light ScanStart" & strSepChar & "Light ScanEnd" & strSepChar & "Light MemberCount" & strSepChar & "Light Drift Time" & strSepChar
    strLineOut = strLineOut & "Light ScanClassRep" & strSepChar & "Light NETClassRep" & strSepChar & "Light AbundanceClassRep" & strSepChar
    strLineOut = strLineOut & "UMC Heavy Ind" & strSepChar & "Heavy MW" & strSepChar & "Heavy Abu" & strSepChar & "Delta Count" & strSepChar
    strLineOut = strLineOut & "Heavy ScanStart" & strSepChar & "Heavy ScanEnd" & strSepChar & "Heavy MemberCount" & strSepChar & "Heavy Drift Time" & strSepChar
    strLineOut = strLineOut & "Heavy ScanClassRep" & strSepChar & "Heavy NETClassRep" & strSepChar & "Heavy AbundanceClassRep" & strSepChar
    strLineOut = strLineOut & "ExpressionRatio" & strSepChar & "ExpressionRatio StDev" & strSepChar & "ER Charge State Basis Count" & strSepChar & "ER Member Basis Count" & strSepChar
    strLineOut = strLineOut & "MassTagID" & strSepChar & "MassTagMonoMW" & strSepChar & "MassTagMods" & strSepChar & "MemberCountMatchingMassTag" & strSepChar & "MassErrorPPM" & strSepChar & "NETError" & strSepChar
    strLineOut = strLineOut & "SLiC Score" & strSepChar & "Del_SLiC" & strSepChar
    strLineOut = strLineOut & "PeptideProphetProbability" & strSepChar & "Peptide" & strSepChar & "Peptide_N_Count"
    If blnIncludeORFInfo Then strLineOut = strLineOut & strSepChar & "MultiORFCount" & strSepChar & "ORFName"
    ts.WriteLine strLineOut
    
    For lngPairInd = 0 To PCount - 1
        If blnReportAllPairs Or _
           blnIdentifiedPairsOnly And PIDCnt(lngPairInd) > 0 Or _
           blnUniquelyIdentifiedOnly And PIDCnt(lngPairInd) = 1 Or _
           blnUnidentifiedPairsOnly And PIDCnt(lngPairInd) = 0 Then
        
            If blnShowExcludedPairs Or (GelP_D_L(CallerID).Pairs(lngPairInd).STATE <> glPAIR_Exc) Then
                'extract pairs information
                With GelP_D_L(CallerID).Pairs(lngPairInd)
                    lngUMCIndexLight = .p1
                    lngUMCIndexHeavy = .p2
                End With
                    
                strPairInfo = Trim(lngPairInd) & strSepChar

                GetUMCClassRepScanAndNET CallerID, lngUMCIndexLight, lngLightScanClassRep, dblLightNETClassRep, dblLightDriftTimeClassRep, dblLightClassRepAbundance
                GetUMCClassRepScanAndNET CallerID, lngUMCIndexHeavy, lngHeavyScanClassRep, dblHeavyNETClassRep, dblHeavyDriftTimeClassRep, dblHeavyClassRepAbundance
                
                With GelUMC(CallerID)
                    ' Light Member
                    With .UMCs(lngUMCIndexLight)
                        strPairInfo = strPairInfo & _
                                        lngUMCIndexLight & strSepChar & _
                                        .ClassMW & strSepChar & _
                                        .ClassAbundance & strSepChar & _
                                        .MinScan & strSepChar & _
                                        .MaxScan & strSepChar & _
                                        .ClassCount & strSepChar & _
                                        dblLightDriftTimeClassRep & strSepChar & _
                                        lngLightScanClassRep & strSepChar & _
                                        Format(dblLightNETClassRep, "0.0000") & strSepChar & _
                                        dblLightClassRepAbundance & strSepChar

                        dblUMCMass = .ClassMW
                    End With
                    
                    ' Heavy Member
                    With .UMCs(lngUMCIndexHeavy)
                        strPairInfo = strPairInfo & _
                                        lngUMCIndexHeavy & strSepChar & _
                                        .ClassMW & strSepChar & _
                                        .ClassAbundance & strSepChar & _
                                        GelP_D_L(CallerID).Pairs(lngPairInd).P2DltCnt & strSepChar & _
                                        .MinScan & strSepChar & _
                                        .MaxScan & strSepChar & _
                                        .ClassCount & strSepChar & _
                                        dblHeavyDriftTimeClassRep & strSepChar & _
                                        lngHeavyScanClassRep & strSepChar & _
                                        Format(dblHeavyNETClassRep, "0.0000") & strSepChar & _
                                        dblHeavyClassRepAbundance & strSepChar

                    End With
                
                End With
                
                With GelP_D_L(CallerID).Pairs(lngPairInd)
                    strPairInfo = strPairInfo & .ER & strSepChar
                    strPairInfo = strPairInfo & .ERStDev & strSepChar
                    strPairInfo = strPairInfo & .ERChargeStateBasisCount & strSepChar
                    strPairInfo = strPairInfo & .ERMemberBasisCount
                End With
                
                If PIDCnt(lngPairInd) < 0 Then          'error during pair identification
                   ts.WriteLine strPairInfo & strSepChar & "Error during identification"
                ElseIf PIDCnt(lngPairInd) = 0 Then      'no id for this pair
                   ts.WriteLine strPairInfo & strSepChar & "Unidentified"
                Else                                    'identified
                   For lngMatchIndex = PIDInd1(lngPairInd) To PIDInd2(lngPairInd)
                        lngMassTagIndexPointer = mMTInd(mUMCMatchStats(lngMatchIndex).IDIndex)
                        lngMassTagIndexOriginal = mMTOrInd(lngMassTagIndexPointer)
                        
                        strIDInfo = strSepChar & Trim(AMTData(lngMassTagIndexOriginal).ID) & strSepChar
                        strIDInfo = strIDInfo & mMTMWN14(mUMCMatchStats(lngMatchIndex).IDIndex) & strSepChar
                    
                        strIDInfo = strIDInfo & MOD_TKN_PAIR_LIGHT
                        If Len(mMTMods(lngMassTagIndexPointer)) > 0 Then
                            strIDInfo = strIDInfo & " " & mMTMods(lngMassTagIndexPointer)
                        End If
                    
                        dblMassErrorPPM = MassToPPM(dblUMCMass - mMTMWN14(mUMCMatchStats(lngMatchIndex).IDIndex), dblUMCMass)
                        dblGANETError = dblLightNETClassRep - mMTNET(lngMassTagIndexPointer)
                       
                        strIDInfo = strIDInfo & strSepChar & mUMCMatchStats(lngMatchIndex).MemberHitCount & strSepChar & Round(dblMassErrorPPM, 4) & strSepChar & Round(dblGANETError, NET_PRECISION)
                        strIDInfo = strIDInfo & strSepChar & Round(mUMCMatchStats(lngMatchIndex).StacOrSLiC, 4)
                        strIDInfo = strIDInfo & strSepChar & Round(mUMCMatchStats(lngMatchIndex).DelScore, 4)
                        strIDInfo = strIDInfo & strSepChar & Round(AMTData(lngMassTagIndexOriginal).PeptideProphetProbability, 5)
                        strIDInfo = strIDInfo & strSepChar & AMTData(lngMassTagIndexOriginal).Sequence
                        strIDInfo = strIDInfo & strSepChar & AMTData(lngMassTagIndexOriginal).CNT_N
                       
                        If Not blnIncludeORFInfo Then
                            ts.WriteLine strPairInfo & strIDInfo
                        Else
                            WriteORFResults ts, strPairInfo & strIDInfo, AMTData(lngMassTagIndexOriginal).ID, objORFNameFastSearch, strSepChar
                        End If
                   Next lngMatchIndex
                End If
            End If
        End If
    
        If lngPairInd Mod 25 = 0 Then
            UpdateStatus "Preparing results: " & Trim(lngPairInd) & " / " & Trim(PCount)
            If mKeyPressAbortProcess > 1 Then Exit For
        End If
    
    Next lngPairInd
    ts.Close
    
    If Len(strOutputFilePath) > 0 Then
        AddToAnalysisHistory CallerID, "Saved search results to disk: " & strOutputFilePath
    End If
    
    Me.MousePointer = vbDefault
    UpdateStatus ""
    If blnDisplayResults Then
         frmDataInfo.Tag = "N14_N15"
         frmDataInfo.Show vbModal
    End If
    
    Set ts = Nothing
    Set fso = Nothing
    Exit Function

ShowOrSavePairsAndIDsErrorHandler:
    Debug.Assert False
    ShowOrSavePairsAndIDs = Err.Number
    LogErrors Err.Number, "frmSearchMTPairs.ShowOrSavePairsAndIDs"
    Set fso = Nothing

End Function

Private Sub StartExportResultsToDB()
    '---------------------------------------------------------
    'triggers export of identified pairs to MT tag database
    'also gives user a chance to change their mind
    '---------------------------------------------------------
    Dim eResponse As VbMsgBoxResult
    Dim strUMCSearchMode As String
    Dim strStatus As String
    
On Error GoTo ExportResultsToDBErrorHandler
    
        If mMatchStatsCount = 0 And Not glbPreferencesExpanded.AutoAnalysisOptions.ExportUMCsWithNoMatches Then
            MsgBox "Search results not found in memory.", vbInformation + vbOKOnly, "Nothing to Export"
        Else
            eResponse = MsgBox("Proceed with exporting of the search results to the database (" & Trim(mMatchStatsCount) & " identified pairs)?  This is an advanced feature that should normally only be performed during VIPER Automated PRISM Analysis Mode.  If you continue, you will be prompted for a password.", vbQuestion + vbYesNo + vbDefaultButton1, "Export Results")
            If eResponse = vbYes Then
                If QueryUserForExportToDBPassword(, False) Then
                    ' Update the text in MD_Parameters
                    strUMCSearchMode = FindSettingInAnalysisHistory(CallerID, UMC_SEARCH_MODE_SETTING_TEXT, , True, ":", ";")
                    If Right(strUMCSearchMode, 1) = ")" Then strUMCSearchMode = Left(strUMCSearchMode, Len(strUMCSearchMode) - 1)
                    
                    GelAnalysis(CallerID).MD_Parameters = ConstructAnalysisParametersText(CallerID, strUMCSearchMode, AUTO_SEARCH_PAIRS_N14N15_CONGLOMERATEMASS)
                    
                    strStatus = ExportMTDBbyUMC(True, mnuPExportDetailedMemberInformation.Checked)
                    MsgBox strStatus, vbInformation + vbOKOnly, glFGTU
                Else
                    MsgBox "Invalid password, export aborted.", vbExclamation Or vbOKOnly, "Invalid"
                End If
            End If
        End If
    
    Exit Sub
    
ExportResultsToDBErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmSearchMTPairs.StartExportResultsToDB"
    Resume Next

End Sub

Public Function StartSearchPaired(Optional blnShowMessages As Boolean = True, Optional strStatusMessage As String = "") As Long
'--------------------------------------------------------
'search pairs in GelP_D_L structure for MT tags matches
'NOTE: possible errors are handled in this procedure
' Returns the number of hits; If an error occurs, will return:
'  -1 = Error searching Database
'  -2 = Error in NET calculation formula
'  -3 = No pairs found
'  -4 = Incorrect pairs type (must have .DltLblType = ptUMCDlt for this function)
'  -5 = Pairs are not synchronized with the LC-MS Features (.SyncWithUMC = False)
'  -6 = Error preparing MT tag search arrays
' Additionally, returns a status message in strStatusMessage
'--------------------------------------------------------
Dim HitsCnt As Long
Dim eResponse As VbMsgBoxResult
Dim I As Long
Dim blnUserNotifiedOfError As Boolean
Dim strSearchDescription As String

On Error GoTo StartSearchPairedErrorHandler

If AMTCnt <= 0 Then
    strStatusMessage = "No MT tags found."
    If blnShowMessages Then MsgBox strStatusMessage, vbOKOnly, glFGTU
    Exit Function
End If

If mwutSearch Is Nothing Then
    strStatusMessage = "Search object not found."
    If blnShowMessages Then MsgBox strStatusMessage, vbOKOnly, glFGTU
    Exit Function
End If

If mMatchStatsCount > 0 And blnShowMessages Then    'something already identified
   eResponse = MsgBox("Pairs identification found. If you continue current findings will be lost. Continue?", vbOKCancel, glFGTU)
   If eResponse <> vbOK Then Exit Function
End If

mKeyPressAbortProcess = 0
cmdSearch.Visible = False

GelData(CallerID).MostRecentSearchUsedSTAC = False

If mMatchStatsCount > 0 Then    'something already identified
   Call DestroyIDStructures
End If

' Unused variable (August 2003)
''mark that structure of identified pairs is not synchronized from this moment
'GelIDP(CallerID).SyncWithDltLblPairs = False

mSearchRegionShape = cboSearchRegionShape.ListIndex

'number of pairs might change so better check every time
PCount = GelP_D_L(CallerID).PCnt
samtDef.Formula = Trim(txtNETFormula.Text)

CheckNETEquationStatus
If PCount > 0 Then
  If PrepareMTArrays() Then
   ' Make sure pairs are the correct search type
   If GelP_D_L(CallerID).DltLblType = ptUMCDlt Then
      If GelP_D_L(CallerID).SyncWithUMC Then
         Me.MousePointer = vbHourglass
         UpdateStatus "Searching: 0 / " & PCount
         
         'reserve space for identifications per pair counts
         'These arrays are parallel to the GelP_D_L().Pairs arrays
         ReDim PIDCnt(PCount - 1)
         ReDim PIDInd1(PCount - 1)
         ReDim PIDInd2(PCount - 1)
         'set last index to -1 so that we know when there was
         'no identification if it doesn't change
         For I = 0 To PCount - 1
             PIDInd2(I) = -1
         Next I
         
         mMatchStatsCount = 0
         'reserve initial space for 10000 identifications
         ReDim mUMCMatchStats(10000)
         
         'do identification pair by pair
         For I = 0 To PCount - 1
             'do not try if pair already excluded
             If GelP_D_L(CallerID).Pairs(I).STATE <> glPAIR_Exc Then
                If I Mod 50 = 0 Then
                    UpdateStatus "Searching: " & Trim(I) & " / " & Trim(PCount)
                End If
                If mKeyPressAbortProcess > 1 Then Exit For
                SearchPairSingleMass (I)
             End If
         Next I
         
         'truncate results
         If mMatchStatsCount > 0 Then
            HitsCnt = mMatchStatsCount
            ReDim Preserve mUMCMatchStats(HitsCnt - 1)
         Else
            HitsCnt = 0
            Erase mUMCMatchStats
         End If
         Me.MousePointer = vbDefault
         If mKeyPressAbortProcess > 1 Then
            UpdateStatus "Search aborted."
         Else
            If chkUpdateGelDataWithSearchResults Then
                ' Store the search results in the gel data
                If mMatchStatsCount > 0 Then RecordSearchResultsInData
            End If
            UpdateStatus "Paired LC-MS Features - MT tag ID Cnt: " & mMatchStatsCount
         End If
         GelStatus(CallerID).Dirty = True
      Else
         HitsCnt = -5               'pairs should be recalculated
      End If
   Else                             'pairs are not correct type
      HitsCnt = -4
   End If
  Else
     HitsCnt = -6                   'error preparing MT search arrays
  End If
Else                                'no pairs found
   HitsCnt = -3
End If
Select Case HitsCnt
Case -1
    strStatusMessage = "Error searching MT database."
    If blnShowMessages Then MsgBox strStatusMessage, vbOKOnly, glFGTU
Case -2
    strStatusMessage = "Error in NET calculation formula."
    If blnShowMessages Then MsgBox strStatusMessage, vbOKOnly, glFGTU
    txtNETFormula.SetFocus
Case -3
    strStatusMessage = "No pairs found. Make sure that one of the LC-MS Feature pairing functions is applied first."
    If blnShowMessages Then MsgBox strStatusMessage, vbOKOnly, glFGTU
Case -4
    strStatusMessage = "Incorrect pairs type.  Should be LC-MS Feature Delta Pairs (e.g. N14/N15 pairs)"
    If blnShowMessages Then MsgBox strStatusMessage, vbOKOnly, glFGTU
Case -5
    strStatusMessage = "Pairs need to be recalculated. Close dialog and recalculate pairs, then return to this dialog."
    If blnShowMessages Then MsgBox strStatusMessage, vbOKOnly, glFGTU
Case -6
    strStatusMessage = "Error occurred while preparing the MT tag search arrays"
    If blnShowMessages Then MsgBox strStatusMessage, vbOKOnly, glFGTU
Case Else
    ' Success (no error)
    strStatusMessage = "MT tag hits: " & HitsCnt & " (non-unique)"
    If blnShowMessages Then MsgBox strStatusMessage, vbOKOnly, glFGTU

    If Not GelAnalysis(CallerID) Is Nothing Then
        If GelAnalysis(CallerID).MD_Type = stNotDefined Or GelAnalysis(CallerID).MD_Type = stStandardIndividual Then
            ' Only update MD_Type if it is currently stStandardIndividual
            GelAnalysis(CallerID).MD_Type = stPairsN14N15
        End If
    End If

    'MonroeMod
    GelSearchDef(CallerID).AMTSearchOnPairs = samtDef
    strSearchDescription = "Searched N14/N15 pairs for MT tags (Conglomerate LC-MS Feature Mass)"
    
    AddToAnalysisHistory CallerID, GetMassTagSearchSummaryText(strSearchDescription, HitsCnt, mMTMinimumHighNormalizedScore, mMTMinimumHighDiscriminantScore, mMTMinimumPeptideProphetProbability, samtDef, True, GelData(CallerID).CustomNETsDefined)
End Select

If Not blnShowMessages And HitsCnt < 0 Then
    If Len(strStatusMessage) = 0 Then
        strStatusMessage = "Unknown error when searching pairs against the MT tag database (HitsCnt = " & Trim(HitsCnt) & ")"
    End If
    AddToAnalysisHistory CallerID, strStatusMessage
End If

cmdSearch.Visible = True

StartSearchPaired = HitsCnt
Exit Function

StartSearchPairedErrorHandler:
    Debug.Assert False
    If blnShowMessages Then
        If Not blnUserNotifiedOfError Then
            MsgBox "Error in frmSearchMTPairs.StartSearchPaired: " & Err.Description
            blnUserNotifiedOfError = True
        End If
    Else
        LogErrors Err.Number, "frmSearchMTPairs.StartSearchPaired"
    End If
    Resume Next
    
End Function

Private Sub UpdateStatus(ByVal Msg As String)
lblStatus.Caption = Msg
DoEvents
End Sub

Private Sub WriteORFResults(ts As TextStream, strLineOutPrefix As String, lngMassTagID As Long, objORFNameFastSearch As FastSearchArrayLong, Optional strSepChar As String = glARG_SEP)
    
    Dim ORFNames() As String            ' 0-based array
    Dim lngORFNamesCount As Long
    Dim lngORFNameIndex As Long

    If MTtoORFMapCount = 0 Then
        lngORFNamesCount = LookupORFNamesForMTIDusingMTDBNamer(objMTDBNameLookupClass, lngMassTagID, ORFNames())
    Else
        lngORFNamesCount = LookupORFNamesForMTIDusingMTtoORFMapOptimized(lngMassTagID, ORFNames(), objORFNameFastSearch)
    End If
    
    If lngORFNamesCount > 0 Then
        For lngORFNameIndex = 0 To lngORFNamesCount - 1
            ts.WriteLine strLineOutPrefix & strSepChar & lngORFNamesCount & strSepChar & ORFNames(lngORFNameIndex)
        Next lngORFNameIndex
    Else
        ts.WriteLine strLineOutPrefix & strSepChar & "0" & strSepChar & "UnknownORF"
    End If

End Sub

Private Sub cboResidueToModify_Click()
    If cboResidueToModify.List(cboResidueToModify.ListIndex) = glPHOSPHORYLATION Then
        txtResidueToModifyMass = Trim(glPHOSPHORYLATION_Mass)
    Else
        ' For safety reasons, reset txtResidueToModifyMass to "0"
        txtResidueToModifyMass = "0"
    End If
End Sub

Private Sub cmdCancel_Click()
    mKeyPressAbortProcess = 2
    KeyPressAbortProcess = 2
End Sub

Private Sub cmdSearch_Click()
    StartSearchPaired
End Sub

Private Sub Form_Activate()
    InitializeSearch
End Sub

Private Sub Form_Load()
'----------------------------------------------------
'load search settings and initializes controls
'----------------------------------------------------

Dim intIndex As Integer

On Error GoTo FormLoadErrorHandler

bLoading = True
If IsWinLoaded(TrackerCaption) Then Unload frmTracker
' MonroeMod
If CallerID >= 1 And CallerID <= UBound(GelBody) Then samtDef = GelSearchDef(CallerID).AMTSearchOnPairs

ShowHidePNNLMenus

'set current Search Definition values
With samtDef
    txtMWTol.Text = .MWTol
    optMWField(.MWField - MW_FIELD_OFFSET).Value = True
    optNETorRT(.NETorRT).Value = True
    Select Case .TolType
    Case gltPPM
      optTolType(0).Value = True
    Case gltABS
      optTolType(1).Value = True
    Case Else
      Debug.Assert False
    End Select
    'save old value and set search on "search all"
    OldSearchFlag = .SearchFlag
    .SearchFlag = 0         'search all
    'NETTol is used both for NET and RT
    If .NETTol >= 0 Then
       txtNETTol.Text = .NETTol
       txtNETTol_Validate False
    Else
       txtNETTol.Text = ""
    End If
End With

With GelSearchDef(CallerID).AMTSearchMassMods
    SetCheckBox chkAlkylation, .Alkylation
    txtAlkylationMWCorrection = .AlkylationMass
    
    PopulateComboBoxes
    
    cboResidueToModify.ListIndex = 0
    If Len(.ResidueToModify) >= 1 Then
        For intIndex = 0 To cboResidueToModify.ListCount - 1
            If UCase(cboResidueToModify.List(intIndex)) = UCase(.ResidueToModify) Then
                cboResidueToModify.ListIndex = intIndex
                Exit For
            End If
        Next intIndex
    End If
    txtResidueToModifyMass = Round(.ResidueMassModification, 5)
    
    SetAlkylationMWCorrection .AlkylationMass
    SetDBSearchModType .ModMode
End With

With glbPreferencesExpanded.MTSConnectionInfo
    ExpAnalysisSPName = .spPutAnalysis
    'ExpPeakSPName = .spPutPeak
    ExpUmcSPName = .spPutUMC
    ExpUMCMemberSPName = .spPutUMCMember
    ExpUmcMatchSPName = .spPutUMCMatch
    ExpUMCCSStats = .spPutUMCCSStats
    ExpQuantitationDescription = .spAddQuantitationDescription
End With

If Len(ExpUmcSPName) = 0 Then
    ExpUmcSPName = "AddFTICRUmc"
End If
Debug.Assert ExpUmcSPName = "AddFTICRUmc"

If Len(ExpUmcMatchSPName) = 0 Then
    ExpUmcMatchSPName = "AddFTICRUmcMatch"
End If
Debug.Assert ExpUmcMatchSPName = "AddFTICRUmcMatch"

If Len(ExpUMCCSStats) = 0 Then
    ExpUMCCSStats = "AddFTICRUmcCSStats"
End If
Debug.Assert ExpUMCCSStats = "AddFTICRUmcCSStats"

If Len(ExpQuantitationDescription) = 0 Then
    ExpQuantitationDescription = "AddQuantitationDescription"
End If
Debug.Assert ExpQuantitationDescription = "AddQuantitationDescription"

If Len(ExpAnalysisSPName) = 0 Then
    ExpAnalysisSPName = "AddMatchMaking"
End If
Debug.Assert ExpAnalysisSPName = "AddMatchMaking"

' September 2004: Unused Variable
''If Len(ExpPeakSPName) = 0 Then
''    ExpPeakSPName = "AddFTICRPeak"
''End If
''Debug.Assert ExpPeakSPName = "AddFTICRPeak"

' Possibly add a checkmark to the mnuFReportIncludeORFs menu
mnuReportIncludeORFName.Checked = glbPreferencesExpanded.AutoAnalysisOptions.IncludeORFNameInTextFileOutput
Exit Sub

FormLoadErrorHandler:
LogErrors Err.Number, "frmSearchMTPairs.Form_Load"
Resume Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
' Restore .SearchFlag using the saved value
samtDef.SearchFlag = OldSearchFlag
End Sub

Private Sub Image1_DblClick()
'-------------------------------------------------------------------
'displays short information about algorithm behind this function
'-------------------------------------------------------------------
Dim tmp As String
tmp = "MT tag DB search for pair members. Pairs are established" & vbCrLf
tmp = tmp & "on unique mass classes and mass delta between heavy and" & vbCrLf
tmp = tmp & "light members determines number of N atoms in underlying" & vbCrLf
tmp = tmp & "peptide. Idea is to search MT tag DB with loose (25ppm)" & vbCrLf
tmp = tmp & "tolerance and select as possible identification those" & vbCrLf
tmp = tmp & "with matching numbers." & vbCrLf
tmp = tmp & "NOTE: Masses over 2500 Da allow N count error of +/-1 N" & vbCrLf
MsgBox tmp, vbOKOnly, glFGTU
End Sub

Private Sub mnuET_Click(Index As Integer)
    Dim I As Long
    Dim intIndexToUse As Integer
    
    If GelData(CallerID).CustomNETsDefined Then
        ' Do not update anything
        Exit Sub
    End If

On Error Resume Next
    If GelAnalysis(CallerID) Is Nothing Then
        intIndexToUse = etGenericNET
    Else
        intIndexToUse = Index
    End If
    
    Select Case intIndexToUse
    Case etGenericNET
        If Index <> etGenericNET Then
            txtNETFormula.Text = GelUMCNETAdjDef(CallerID).NETFormula
        Else
            txtNETFormula.Text = ConstructNETFormulaWithDefaults(GelUMCNETAdjDef(CallerID))
        End If
    Case etTICFitNET
      With GelAnalysis(CallerID)
        If .NET_Slope <> 0 Then
            txtNETFormula.Text = ConstructNETFormula(.NET_Slope, .NET_Intercept)
        Else
            txtNETFormula.Text = ConstructNETFormulaWithDefaults(GelUMCNETAdjDef(CallerID))
        End If
      End With
      If Err Then
         MsgBox "Make sure display is loaded as analysis. Use New Analysis command from the File menu.", vbOKOnly, glFGTU
         Exit Sub
      End If
    Case etGANET
      With GelAnalysis(CallerID)
        If .GANET_Slope <> 0 Then
           txtNETFormula.Text = ConstructNETFormula(.GANET_Slope, .GANET_Intercept)
        Else
           txtNETFormula.Text = ConstructNETFormulaWithDefaults(GelUMCNETAdjDef(CallerID))
        End If
      End With
      If Err Then
         MsgBox "Make sure display is loaded as analysis. Use New Analysis command from the File menu.", vbOKOnly, glFGTU
         Exit Sub
      End If
    End Select
    For I = mnuET.LBound To mnuET.UBound
        If I = Index Then
           mnuET(I).Checked = True
           lblETType.Caption = "ET Type: " & mnuET(I).Caption
        Else
           mnuET(I).Checked = False
        End If
    Next I
    Call txtNETFormula_LostFocus        'make sure expression evaluator is
                                        'initialized for this formula
End Sub

Private Sub mnuETHeader_Click()
Call PickParameters
End Sub

Private Sub mnuExpMTDB_Click()
    StartExportResultsToDB
End Sub

Private Sub mnuMT_Click()
Call PickParameters
End Sub

Private Sub mnuMTLoadLegacy_Click()
    LoadLegacyMassTags
End Sub

Private Sub mnuMTLoadMT_Click()
'------------------------------------------------------------
'load/reload MT tags
'------------------------------------------------------------
If Not GelAnalysis(CallerID) Is Nothing Then
   Call LoadMTDB(True)
Else
   WarnUserNotConnectedToDB CallerID, True
   lblMTStatus.Caption = "No MT tags loaded"
End If
End Sub

Private Sub mnuMTStatus_Click()
'----------------------------------------------
'displays short MT tags statistics, it might
'help with determining problems with MT tags
'----------------------------------------------
Me.MousePointer = vbHourglass
MsgBox CheckMassTags(), vbOKOnly
Me.MousePointer = vbDefault
End Sub

Private Sub mnuP_Click()
Call PickParameters
End Sub

Private Sub mnuPCalculateER_Click()
'------------------------------------
'recalculate ER numbers for all pairs
'------------------------------------
Dim strMessage As String

Dim objDltLblPairsUMC As New clsDltLblPairsUMC
objDltLblPairsUMC.CalcDltLblPairsER_UMC CallerID, strMessage

UpdateStatus strMessage
End Sub

Private Sub mnuPClose_Click()
Unload Me
End Sub

Private Sub mnuPDeleteExcluded_Click()
    Me.DeleteExcludedPairsWrapper
End Sub

Private Sub mnuPExcludeAmbiguous_Click()
    Me.ExcludeAmbiguousPairsWrapper False
End Sub

Private Sub mnuPExcludeAmbiguousHitsOnly_Click()
   Me.ExcludeAmbiguousPairsWrapper True
End Sub

Private Sub mnuPExcludeIdentified_Click()
'----------------------------------------
'exclude all identified pairs
'----------------------------------------
Dim I As Long
On Error Resume Next
With GelP_D_L(CallerID)
    For I = 0 To .PCnt - 1
        If PIDCnt(I) > 0 Then .Pairs(I).STATE = glPAIR_Exc
    Next I
End With
End Sub

Private Sub mnuPExcludeUnidentified_Click()
'------------------------------------------
'exclude all identified pairs
'------------------------------------------
Dim I As Long
On Error Resume Next
With GelP_D_L(CallerID)
    For I = 0 To .PCnt - 1
        If PIDCnt(I) <= 0 Then .Pairs(I).STATE = glPAIR_Exc
    Next I
End With
End Sub

Private Sub mnuPExportDetailedMemberInformation_Click()
    mnuPExportDetailedMemberInformation.Checked = Not mnuPExportDetailedMemberInformation.Checked
End Sub

Private Sub mnuPIncludeUnqIdentified_Click()
'---------------------------------------------------
'exclude everything that is not uniquelly identified
'---------------------------------------------------
Dim I As Long
On Error Resume Next
With GelP_D_L(CallerID)
    For I = 0 To .PCnt - 1
        If PIDCnt(I) = 1 Then
           .Pairs(I).STATE = glPAIR_Inc
        Else
           .Pairs(I).STATE = glPAIR_Exc
        End If
    Next I
End With
End Sub

Private Sub mnuPSearch_Click()
    StartSearchPaired
End Sub

Private Sub mnuRAllPairsAndIDs_Click()
    ShowOrSavePairsAndIDs "", True
End Sub

Private Sub mnuReport_Click()
Call PickParameters
End Sub

Private Sub mnuReportIncludeORFName_Click()
    mnuReportIncludeORFName.Checked = Not mnuReportIncludeORFName.Checked
    glbPreferencesExpanded.AutoAnalysisOptions.IncludeORFNameInTextFileOutput = mnuReportIncludeORFName.Checked
End Sub

Private Sub mnuRIdentified_Click()
    ShowOrSavePairsAndIDs "", True, True, False, False
End Sub

Private Sub mnuRUnidentified_Click()
    ShowOrSavePairsAndIDs "", True, False, False, True
End Sub

Private Sub mnuRUnqIdentified_Click()
    ShowOrSavePairsAndIDs "", True, False, True, False
End Sub

Private Sub optMWField_Click(Index As Integer)
samtDef.MWField = 6 + Index
End Sub

Private Sub optNETorRT_Click(Index As Integer)
samtDef.NETorRT = Index
End Sub

Private Sub optTolType_Click(Index As Integer)
If Index = 0 Then
   samtDef.TolType = gltPPM
Else
   samtDef.TolType = gltABS
End If
End Sub


Private Sub txtAlkylationMWCorrection_LostFocus()
If IsNumeric(txtAlkylationMWCorrection.Text) Then
   AlkMWCorrection = CDbl(txtAlkylationMWCorrection.Text)
Else
   txtAlkylationMWCorrection.Text = glALKYLATION
   AlkMWCorrection = glALKYLATION
End If
End Sub

Private Sub txtDBSearchMinimumHighDiscriminantScore_LostFocus()
    ValidateTextboxValueDbl txtDBSearchMinimumHighDiscriminantScore, 0, 1, 0
End Sub

Private Sub txtDBSearchMinimumHighNormalizedScore_LostFocus()
    ValidateTextboxValueDbl txtDBSearchMinimumHighNormalizedScore, 0, 100000, 0
End Sub

Private Sub txtDBSearchMinimumPeptideProphetProbability_LostFocus()
    ValidateTextboxValueDbl txtDBSearchMinimumPeptideProphetProbability, 0, 1, 0
End Sub

Private Sub txtMWTol_LostFocus()
If IsNumeric(txtMWTol.Text) Then
   samtDef.MWTol = CDbl(txtMWTol.Text)
Else
   MsgBox "Molecular Mass Tolerance should be numeric value.", vbOKOnly
   txtMWTol.SetFocus
End If
End Sub

Private Sub txtNETFormula_LostFocus()
'------------------------------------------------
'initialize new expression evaluator
'------------------------------------------------
If Not GelData(CallerID).CustomNETsDefined Then
    If Not InitExprEvaluator(txtNETFormula.Text) Then
       MsgBox "Error in elution calculation formula.", vbOKOnly, glFGTU
       txtNETFormula.SetFocus
    Else
       samtDef.Formula = txtNETFormula.Text
    End If
End If
End Sub

Private Sub txtNETTol_LostFocus()
If IsNumeric(txtNETTol.Text) Then
   samtDef.NETTol = CDbl(txtNETTol.Text)
Else
   If Len(Trim(txtNETTol.Text)) > 0 Then
      MsgBox "NET Tolerance should be number between 0 and 1.", vbOKOnly
      txtNETTol.SetFocus
   Else
      samtDef.NETTol = -1   'do not consider NET when searching
   End If
End If
End Sub

Private Sub txtNETTol_Validate(Cancel As Boolean)
    TextBoxLimitNumberLength txtNETTol, 12
End Sub

Private Sub txtResidueToModifyMass_LostFocus()
    ValidateTextboxValueDbl txtResidueToModifyMass, -10000, 10000, 0
End Sub

