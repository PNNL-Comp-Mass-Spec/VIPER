VERSION 5.00
Begin VB.Form frmSearchMTPairs_PEO 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search MT Tag DB For N14/N15 Labeled Pairs"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2640
      TabIndex        =   30
      Top             =   1200
      Width           =   495
      Begin VB.Image Image1 
         Height          =   480
         Left            =   0
         Picture         =   "frmSearchMTPairs_PEO.frx":0000
         ToolTipText     =   "Double-click for short info on this procedure"
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Frame fraLabelSettings 
      BackColor       =   &H0080C0FF&
      Caption         =   "Label Settings"
      Height          =   1335
      Left            =   3840
      TabIndex        =   20
      Top             =   1440
      Width           =   1935
      Begin VB.TextBox txtMaxLbls 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   26
         Text            =   "5"
         Top             =   900
         Width           =   495
      End
      Begin VB.TextBox txtMinLbls 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   24
         Text            =   "1"
         Top             =   570
         Width           =   495
      End
      Begin VB.TextBox txtLblMass 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   22
         Text            =   "414.1937"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max Labels:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Min Labels:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   630
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mol.Mass:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.Frame fraMTassumption 
      BackColor       =   &H0080C0FF&
      Caption         =   "MT Tags"
      Height          =   1215
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   1455
      Begin VB.OptionButton optAssumeLabels 
         BackColor       =   &H0080C0FF&
         Caption         =   "Not Labeled"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "MT Tags are not labeled (mass represent mass of the clean sequence)"
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optAssumeLabels 
         BackColor       =   &H0080C0FF&
         Caption         =   "Labeled"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "MT Tags are labeled (mass modified)"
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Frame fraNET 
      BackColor       =   &H0080C0FF&
      Caption         =   "NET  Calculation"
      Height          =   1575
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   3615
      Begin VB.CheckBox chkUseUMCConglomerateNET 
         BackColor       =   &H0080C0FF&
         Caption         =   "Use Class NET for LC-MS Features"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   $"frmSearchMTPairs_PEO.frx":030A
         Top             =   1200
         Width           =   3165
      End
      Begin VB.OptionButton optNETorRT 
         BackColor       =   &H0080C0FF&
         Caption         =   "Pred. NET"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   16
         ToolTipText     =   "Use NET calculated only from Sequest ""first choice"" peptides"
         Top             =   900
         Width           =   1215
      End
      Begin VB.OptionButton optNETorRT 
         BackColor       =   &H0080C0FF&
         Caption         =   "Obs. NET"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Use NET calculated from all peptides of MT Tags"
         Top             =   900
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox txtNETFormula 
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Text            =   "(FN-MinFN)/(MaxFN-MinFN)"
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox txtNETTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   18
         Text            =   "0.15"
         Top             =   870
         Width           =   495
      End
      Begin VB.Label lblNETFormula 
         BackStyle       =   0  'Transparent
         Caption         =   "&Formula  F(FN, MinFN, MaxFN)"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T&ol."
         Height          =   195
         Left            =   2640
         TabIndex        =   17
         ToolTipText     =   "Normalized Elution Time Tolerance (if blank NET will not be considered in search)"
         Top             =   900
         Width           =   270
      End
   End
   Begin VB.Frame fraMWTolerance 
      BackColor       =   &H0080C0FF&
      Caption         =   "Mol.Mass Tolerance"
      Height          =   1215
      Left            =   3840
      TabIndex        =   7
      Top             =   120
      Width           =   1935
      Begin VB.OptionButton optTolType 
         BackColor       =   &H0080C0FF&
         Caption         =   "&Dalton"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   11
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optTolType 
         BackColor       =   &H0080C0FF&
         Caption         =   "&ppm"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtMWTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   160
         TabIndex        =   9
         Text            =   "10"
         Top             =   640
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tolerance"
         Height          =   255
         Left            =   160
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame fraMWField 
      BackColor       =   &H0080C0FF&
      Caption         =   "Molecular Mass Field"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      Begin VB.OptionButton optMWField 
         BackColor       =   &H0080C0FF&
         Caption         =   "&The Most Abundant"
         Height          =   255
         Index           =   2
         Left            =   80
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton optMWField 
         BackColor       =   &H0080C0FF&
         Caption         =   "&Monoisotopic"
         Height          =   255
         Index           =   1
         Left            =   80
         TabIndex        =   2
         Top             =   540
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optMWField 
         BackColor       =   &H0080C0FF&
         Caption         =   "A&verage"
         Height          =   255
         Index           =   0
         Left            =   80
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Label lblETType 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3480
      TabIndex        =   31
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   3360
      Width           =   5655
   End
   Begin VB.Label lblMTStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      Height          =   255
      Left            =   1200
      TabIndex        =   28
      Top             =   3120
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MTDB Status:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   27
      ToolTipText     =   "Status of the MT Tag database"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Menu mnuP 
      Caption         =   "&Pairs"
      Begin VB.Menu mnuPSearch 
         Caption         =   "Search(Identify)"
      End
      Begin VB.Menu mnuPSep17 
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
         Caption         =   "Exclude Ambiguous Pairs"
      End
      Begin VB.Menu mnuPDeleteExcluded 
         Caption         =   "Delete Excluded Pairs"
      End
      Begin VB.Menu mnuPCalculateER 
         Caption         =   "Calculate ER"
      End
      Begin VB.Menu mnuPSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPExpLegacyDB 
         Caption         =   "Export Results To &Legacy DB"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuExpMTDB 
         Caption         =   "Export Results To &MT Tag DB"
      End
      Begin VB.Menu mnuPSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSyncPairsStructure 
         Caption         =   "Sync With ID Pairs Structure"
      End
      Begin VB.Menu mnuPSep3 
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
         Caption         =   "&Generic ET"
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
Attribute VB_Name = "frmSearchMTPairs_PEO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'search of MT tag database for pairs member of the isotopic
'labeled pairs labeled with PEO label;
'assumption is that pairs are calculated for UMC and that
'search of the database can be performed with loose tolerance
'with additional criteria of matching number of N with delta
'in molecular mass for established pairs;
'------------------------------------------------------------
'NOTE: pairs coming to this function are N14/N15 delta pairs
'in which PEO label does not play any role. PEO label is used
'just as an marker on Cys- containing peptides. From pairing
'information we do not know number of PEO so we have to assume
'any reasonable number.
'After potential identification is detected number of N atoms
'is checked with pair delta and Cys count has to be at least
'as assumed.
'NOTE: assumption has to be made also about loaded MT tags
'sequences; if they are modified we assume they are PEO
'modified in which case we count on user intelligence and
'search only light pairs; in that case number of cysteins also
'has to be exact match
'If MT tags are not modified then we need to substract
'number of assumed labels from the pair light member mass
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'NOTE: this approach makes it impossible to search both
'modified and non-modified MT tags
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'NOTE: this function is derived from N14/N15 search function
'------------------------------------------------------------
'created: 04/18/2002 nt
'last modified: 07/29/2002 nt
'------------------------------------------------------------
Option Explicit

Const MIN_MOL_MASS = 100  'minimum of reasonable molecular mass

'in this case CallerID is a public property
Public CallerID As Long

Dim bLoading As Boolean

Dim OldSearchFlag As Long

Dim bMTMod As Long

'label settings
Dim LblMW As Double             'should be PEO but it is allowed to change
Dim LblMin As Long              'min number of labels allowed
Dim LblMax As Long              'maximum number labels allowed

'following arrays are parallel to the pairs arrays; it is used for
'easier classification between identified and nonidentified pairs
Dim PCount As Long              'shortcut for number of pairs
Dim PIDCnt() As Long            'count of OK identifications(unique) for pair
Dim PIDInd1() As Long           'first index in ID arrays for pair
Dim PIDInd2() As Long           'last index in ID arrays for pair

Dim ClsCnt As Long              'this is not actually neccessary except
Dim ClsStat() As Double         'to create nice reports
                                
'next 3 arrays contain all possible identifications for all pairs
Dim mgCnt As Long               'count of pair-id matches
Dim mgPInd() As Long            'index of identified pair
Dim mgIDInd() As Long           'index of identification
Dim mgScore() As Double         'score for each identification
'NOTE: mg stands for "mare nostrum", which is the latin name for the Mediterranean Sea

'following variables are used for results output
Dim fso As New FileSystemObject
Dim ts As TextStream
Dim fname As String

'Expression Evaluator variables for elution time calculation
Dim MyExprEva As ExprEvaluator
Dim VarVals() As Long
Dim MinFN As Long
Dim MaxFN As Long

'names of stored procedures that will write data
'to database tables retrieved from init. file
Dim ExpAnalysisSPName As String             ' Stored procedure AddMatchMaking
''Dim ExpPeakSPName As String               ' Stored procedure AddFTICRPeak; Unused variable

Private mUsingDefaultGANET As Boolean

Private Sub chkUseUMCConglomerateNET_Click()
    glbPreferencesExpanded.UseUMCConglomerateNET = cChkBox(chkUseUMCConglomerateNET)
End Sub

Private Sub Form_Activate()
'------------------------------------------------------------
'load MT tag database data if neccessary
'if CallerID is associated with MT tag database load that
'database if neccessary; if CallerID is not associated with
'MT tag database load legacy database
'------------------------------------------------------------
On Error Resume Next
Me.MousePointer = vbHourglass
If bLoading Then
   txtNETFormula.Text = samtDef.Formula
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
         End If
      
      Else                  'nothing is loaded
         WarnUserNotConnectedToDB CallerID, True
         lblMTStatus.Caption = "No MT tags loaded"
      End If
   Else         'have to have MT tag database loaded
      Call LoadMTDB
      With GelAnalysis(CallerID)
         If .NET_Slope <> 0 Then
            txtNETFormula.Text = ConstructNETFormula(.NET_Slope, .NET_Intercept)
         End If
      End With
   End If
   UpdateStatus "Generating LC-MS Feature statistic ..."
   ClsCnt = UMCStatistics1(CallerID, ClsStat())
   PCount = GelP_D_L(CallerID).PCnt
   UpdateStatus "Potential Pairs: " & PCount
   
    txtNETFormula.Enabled = Not GelData(CallerID).CustomNETsDefined
    lblNETFormula.Enabled = txtNETFormula.Enabled
    mnuETHeader.Enabled = txtNETFormula.Enabled
    
    If Not GelData(CallerID).CustomNETsDefined Then
        Call mnuET_Click(etGANET)
    End If
   
   'check elution calculation formula
   'If Not InitExprEvaluator(txtNETFormula.Text) Then
   '   MsgBox "Error in elution evaluation formula.", vbOKOnly, glFGTU
   '   txtNETFormula.SetFocus
   'End If
   'memorize number of scans for Caller(to be used with elution)
   MinFN = GelData(CallerID).ScanInfo(1).ScanNumber
   MaxFN = GelData(CallerID).ScanInfo(UBound(GelData(CallerID).ScanInfo)).ScanNumber
   bLoading = False
End If
Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
'----------------------------------------------------
'load search settings and initializes controls
'----------------------------------------------------
On Error Resume Next
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
    End Select
    'save old value and set search on "search all"
    OldSearchFlag = .SearchFlag
    .SearchFlag = 0         'search all
    'NETTol is used both for NET and RT
    If .NETTol >= 0 Then
       txtNETTol.Text = .NETTol
    Else
       txtNETTol.Text = ""
    End If
End With
bMTMod = 0      'not modified by default

LblMW = CDbl(txtLblMass.Text)
LblMin = CLng(txtMinLbls.Text)
LblMax = CLng(txtMaxLbls.Text)

'temporary file for results output
fname = GetTempFolder() & RawDataTmpFile

ExpAnalysisSPName = glbPreferencesExpanded.MTSConnectionInfo.spPutAnalysis
'ExpPeakSPName = glbPreferencesExpanded.MTSConnectionInfo.spPutPeak

If Len(ExpAnalysisSPName) <= 0 Then
   mnuExpMTDB.Enabled = False
   UpdateStatus "Missing names of export functions."
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'DestroyAMTStat
samtDef.SearchFlag = OldSearchFlag
Set fso = Nothing
End Sub

Private Sub Image1_DblClick()
'-------------------------------------------------------------------
'displays short information about algorithm behind this function
'-------------------------------------------------------------------
Dim tmp As String
tmp = "MT tag DB search for pair members. Pairs are established" & vbCrLf
tmp = tmp & "on unique mass classes and mass delta between heavy and" & vbCrLf
tmp = tmp & "light members determines number of N atoms in underlying" & vbCrLf
tmp = tmp & "peptide. Idea is to search MT tag DB with loose(25ppm)" & vbCrLf
tmp = tmp & "tolerance and select as possible identification those" & vbCrLf
tmp = tmp & "with matching numbers." & vbCrLf
tmp = tmp & "NOTE: Masses over 2500 Da allow N count error of +/-1 N" & vbCrLf
MsgBox tmp, vbOKOnly, glFGTU
End Sub


Private Sub mnuET_Click(Index As Integer)
    Dim i As Long
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
    For i = mnuET.LBound To mnuET.UBound
        If i = Index Then
           mnuET(i).Checked = True
           lblETType.Caption = "ET Type: " & mnuET(i).Caption
        Else
           mnuET(i).Checked = False
        End If
    Next i
    Call txtNETFormula_LostFocus        'make sure expression evaluator is
                                        'initialized for this formula
End Sub

Private Sub mnuETHeader_Click()
Call PickParameters
End Sub

Private Sub mnuExpMTDB_Click()
'---------------------------------------------------------
'triggers export of identified pairs to MT tag database
'also gives user a chance to change their mind
'---------------------------------------------------------
Dim Res As Long
Dim WhatHappened As String
On Error Resume Next
If GelAnalysis(CallerID).Job > 0 Then
   If mgCnt > 0 Then
   
        MsgBox "Database export is no longer supported using this window.", vbInformation + vbOKOnly, "Error"
    
        ' September 2004: Unsupported code
''      Res = MsgBox(mgCnt & " identified pairs will be exported to MT tag database.", vbOKCancel, glFGTU)
''      If Res <> vbOK Then Exit Sub
''      UpdateStatus "Generating report ..."
''      Me.MousePointer = vbHourglass
''      WhatHappened = ExportIDPairsToMTDB
''      MsgBox WhatHappened, vbOKOnly, glFGTU
''      Me.MousePointer = vbDefault
''      UpdateStatus ""
   Else
      MsgBox "No identified pairs found.", vbOKOnly, glFGTU
   End If
Else        'this does not come from the database; can not export
   MsgBox "Current display was not loaded from the PRISM; therefore it can not be exported there. Sorry.", vbOKOnly, glFGTU
End If
End Sub

Private Sub mnuMT_Click()
Call PickParameters
End Sub

Private Sub mnuMTLoadLegacy_Click()
'------------------------------------------------------------
'load/reload MT tags
'------------------------------------------------------------
Dim Respond As Long
On Error Resume Next
'ask user if it wants to replace legitimate MT tag DB with legacy DB
If Not GelAnalysis(CallerID) Is Nothing And Not APP_BUILD_DISABLE_MTS Then
   Respond = MsgBox("Current display is associated with MT tag database." & vbCrLf _
                & "Are you sure you want to use a legacy database for search?", vbYesNoCancel, glFGTU)
   If Respond <> vbYes Then Exit Sub
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
'--------------------------------------------
'removes excluded pairs from the structure
'--------------------------------------------
UpdateStatus "Deleting excluded pairs ..."
Me.MousePointer = vbHourglass
UpdateStatus DeleteExcludedPairs(CallerID)
Me.MousePointer = vbDefault
End Sub

Private Sub mnuPExcludeAmbiguous_Click()
'---------------------------------------------------
'mark as excluded all ambiguous pairs
'to increase the number of unambiguous pairs, this
'procedure should be applied at the end, after all
'other filtering
'---------------------------------------------------
Dim strMessage As String
strMessage = PairsSearchMarkAmbiguous(Me, CallerID, True)
UpdateStatus strMessage
End Sub

Private Sub mnuPExcludeIdentified_Click()
'----------------------------------------
'exclude all identified pairs
'----------------------------------------
Dim i As Long
On Error Resume Next
With GelP_D_L(CallerID)
    For i = 0 To .PCnt - 1
        If PIDCnt(i) > 0 Then .Pairs(i).STATE = glPAIR_Exc
    Next i
End With
End Sub

Private Sub mnuPExcludeUnidentified_Click()
'------------------------------------------
'exclude all identified pairs
'------------------------------------------
Dim i As Long
On Error Resume Next
With GelP_D_L(CallerID)
    For i = 0 To .PCnt - 1
        If PIDCnt(i) <= 0 Then .Pairs(i).STATE = glPAIR_Exc
    Next i
End With
End Sub

Private Sub mnuPIncludeUnqIdentified_Click()
'---------------------------------------------------
'exclude everything that is not uniquelly identified
'---------------------------------------------------
Dim i As Long
On Error Resume Next
With GelP_D_L(CallerID)
    For i = 0 To .PCnt - 1
        If PIDCnt(i) = 1 Then
           .Pairs(i).STATE = glPAIR_Inc
        Else
           .Pairs(i).STATE = glPAIR_Exc
        End If
    Next i
End With
End Sub

Private Sub mnuPSearch_Click()
'--------------------------------------------------------
'search pairs in GelP_D_L structure for MT tags matches
'NOTE: possible errors are handled in this procedure
'--------------------------------------------------------
Dim HitsCnt As Long
Dim Respond As Long
Dim i As Long
On Error Resume Next

If AMTCnt <= 0 Then
   MsgBox "No MT tags found.", vbOKOnly, glFGTU
   Exit Sub
End If

If mwutSearch Is Nothing Then
   MsgBox "Search object not found.", vbOKOnly, glFGTU
   Exit Sub
End If

If mgCnt > 0 Then    'something already identified
   Respond = MsgBox("Pairs identification found. If you continue current findings will be lost. Continue?", vbOKCancel, glFGTU)
   If Respond <> vbOK Then Exit Sub
End If

' Unused variable (August 2003)
''mark that structure of identified pairs is not synchronized from this moment
'GelIDP(CallerID).SyncWithDltLblPairs = False

'number of pairs might change so better check every time
PCount = GelP_D_L(CallerID).PCnt
samtDef.Formula = Trim$(txtNETFormula.Text)
If PCount > 0 Then
  With GelP_D_L(CallerID)
    If .DltLblType = ptUMCDlt Then          'this can only work on delta pairs
      If .SyncWithUMC Then
        Me.MousePointer = vbHourglass
        'reserve space for identifications per pair counts
        ReDim PIDCnt(PCount - 1)
        ReDim PIDInd1(PCount - 1)
        ReDim PIDInd2(PCount - 1)
        'set last index to -1 so that we know when there was
        'no identification if it doesn't change
        For i = 0 To PCount - 1
            PIDInd2(i) = -1
        Next i
        mgCnt = 0
        'reserve initial space for 10000 identifications
        ReDim mgPInd(10000)
        ReDim mgIDInd(10000)
        ReDim mgScore(10000)
        'do identification pair by pair
        If bMTMod Then
           For i = 0 To PCount - 1
             If .Pairs(i).STATE <> glPAIR_Exc Then       'skip excluded pairs
                UpdateStatus "Identifying pair " & (i + 1) & "/" & PCount
                DoThePair_Mod (i)
             End If
           Next i
        Else
           For i = 0 To PCount - 1
             If .Pairs(i).STATE <> glPAIR_Exc Then       'skip excluded pairs
                UpdateStatus "Identifying pair " & (i + 1) & "/" & PCount
                DoThePair_NotMod (i)
             End If
           Next i
        End If
        'truncate results
        If mgCnt > 0 Then
           HitsCnt = mgCnt
           ReDim Preserve mgPInd(HitsCnt - 1)
           ReDim Preserve mgIDInd(HitsCnt - 1)
           ReDim Preserve mgScore(HitsCnt - 1)
        Else
           HitsCnt = 0
           Erase mgPInd
           Erase mgIDInd
           Erase mgScore
        End If
        Me.MousePointer = vbDefault
        UpdateStatus ""
        GelStatus(CallerID).Dirty = True
      Else
        HitsCnt = -5                    'pairs should be recalculated
      End If
    Else                                'pairs are not correct type
      HitsCnt = -4
    End If
  End With
Else                                    'no pairs found
   HitsCnt = -3
End If
Select Case HitsCnt
Case -1
   MsgBox "Error searching MT database.", vbOKOnly, glFGTU
Case -2
   MsgBox "Error in NET calculation formula.", vbOKOnly, glFGTU
   txtNETFormula.SetFocus
Case -3
   MsgBox "No pairs found. Make sure that one of LC-MS Feature pairing functions is applied first.", vbOKOnly, glFGTU
Case -4
   MsgBox "Incorrect pairs type.", vbOKOnly, glFGTU
Case -5
   MsgBox "Pairs need to be recalculated. Close dialog and recalculate pairs, then return to this dialog.", vbOKOnly, glFGTU
Case Else
   MsgBox "MT tag hits: " & HitsCnt & " (non-unique)", vbOKOnly, glFGTU
    
    If Not GelAnalysis(CallerID) Is Nothing Then
        If GelAnalysis(CallerID).MD_Type = stNotDefined Or GelAnalysis(CallerID).MD_Type = stStandardIndividual Then
            ' Only update MD_Type if it is currently stStandardIndividual
            GelAnalysis(CallerID).MD_Type = stPairsPEO
        End If
    End If

    'MonroeMod
    GelSearchDef(CallerID).AMTSearchOnPairs = samtDef
    AddToAnalysisHistory CallerID, GetMassTagSearchSummaryText("Searched PEO pairs for MT tags", HitsCnt, 0, 0, 0, samtDef, True, GelData(CallerID).CustomNETsDefined)
End Select
End Sub

Private Sub mnuRAllPairsAndIDs_Click()
'-------------------------------------
'report pairs and identifications
'-------------------------------------
Dim i As Long
Dim Respond As Long
If mgCnt <= 0 Then
   Respond = MsgBox("No identification found. Continue with generating report?", vbYesNo, glFGTU)
   If Respond <> vbYes Then Exit Sub
End If
UpdateStatus "Generating report ..."
Me.MousePointer = vbHourglass
Set ts = fso.OpenTextFile(fname, ForWriting, True)
WriteReportHeader "All pairs and identifications"
For i = 0 To PCount - 1
    ReportPair i
Next i
ts.Close
Set ts = Nothing
Me.MousePointer = vbDefault
UpdateStatus ""
frmDataInfo.Tag = "N14_N15"
frmDataInfo.Show vbModal
End Sub

Private Sub mnuReport_Click()
Call PickParameters
End Sub

Private Sub mnuRIdentified_Click()
'-------------------------------------
'report identified pairs only
'-------------------------------------
Dim i As Long
On Error Resume Next
If mgCnt > 0 Then
   UpdateStatus "Generating report ..."
   Me.MousePointer = vbHourglass
   Set ts = fso.OpenTextFile(fname, ForWriting, True)
   WriteReportHeader "Identified pairs only"
   For i = 0 To PCount - 1
       If PIDCnt(i) > 0 Then ReportPair i
   Next i
   ts.Close
   Set ts = Nothing
   Me.MousePointer = vbDefault
   UpdateStatus ""
   frmDataInfo.Tag = "N14_N15"
   frmDataInfo.Show vbModal
Else
   MsgBox "No identified pairs found.", vbOKOnly, glFGTU
End If
End Sub

Private Sub mnuRUnidentified_Click()
'-------------------------------------
'report unidentified pairs only
'-------------------------------------
Dim i As Long
On Error Resume Next
If PCount > 0 Then
   UpdateStatus "Generating report ..."
   Me.MousePointer = vbHourglass
   Set ts = fso.OpenTextFile(fname, ForWriting, True)
   WriteReportHeader "Unidentified pairs only"
   For i = 0 To PCount - 1
       If PIDCnt(i) <= 0 Then ReportPair i
   Next i
   ts.Close
   Set ts = Nothing
   Me.MousePointer = vbDefault
   UpdateStatus ""
   frmDataInfo.Tag = "N14_N15"
   frmDataInfo.Show vbModal
Else
   MsgBox "No pairs found.", vbOKOnly, glFGTU
End If
End Sub

Private Sub mnuRUnqIdentified_Click()
'-------------------------------------
'report uniquely identified pairs only
'-------------------------------------
Dim i As Long
On Error Resume Next
If mgCnt > 0 Then
   UpdateStatus "Generating report ..."
   Me.MousePointer = vbHourglass
   Set ts = fso.OpenTextFile(fname, ForWriting, True)
   WriteReportHeader "Uniquely identified pairs only"
   For i = 0 To PCount - 1
       If PIDCnt(i) = 1 Then ReportPair i
   Next i
   ts.Close
   Set ts = Nothing
   Me.MousePointer = vbDefault
   UpdateStatus ""
   frmDataInfo.Tag = "N14_N15"
   frmDataInfo.Show vbModal
Else
   MsgBox "No identified pairs found.", vbOKOnly, glFGTU
End If
End Sub

' Unused Function (August 2003)
' This function was used to populate the GelIDP() udt; however, that udt is no longer used
'Private Sub mnuSyncPairsStructure_Click()
''store identified pairs information to memory structure
''so that it can be comunicated to other functions
'Dim Res As Long
'Dim i As Long
'On Error Resume Next
'If mgCnt > 0 Then
'   With GelIDP(CallerID)
'       If .Cnt > 0 Then
'          Res = MsgBox("Some pairs found in identified pairs structure. Continuing will overwrite this information with the most recent pairs identifications. Continue?", vbYesNoCancel, glFGTU)
'          If Res <> vbYes Then Exit Sub
'       End If
'       ReDim .PInd(mgCnt - 1)
'       ReDim .PIDInd(mgCnt - 1)
'       .Cnt = mgCnt
'       For i = 0 To mgCnt - 1
'           .PInd(i) = mgPInd(i)
'           .PIDInd(i) = mgIDInd(i)
'       Next i
'       .SyncWithDltLblPairs = True
'   End With
'Else
'   Res = MsgBox("There are no ID pairs. Do you want to clear identified pairs structure?", vbYesNoCancel, glFGTU)
'   If Res = vbYes Then
'      With GelIDP(CallerID)
'          .Cnt = 0
'          Erase .PInd
'          Erase .PIDInd
'          .SyncWithDltLblPairs = True
'      End With
'   End If
'End If
'End Sub

Private Sub optAssumeLabels_Click(Index As Integer)
bMTMod = Index
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

Private Sub txtLblMass_LostFocus()
Dim tmp As String
tmp = txtLblMass.Text
If IsNumeric(tmp) Then
   If tmp >= 0 Then
      LblMW = CDbl(tmp)
      Exit Sub
   End If
End If
MsgBox "This argument should be non-negative number.", vbOKOnly, glFGTU
txtLblMass.SetFocus
End Sub

Private Sub txtMaxLbls_LostFocus()
Dim tmp As String
tmp = txtMaxLbls.Text
If IsNumeric(tmp) Then
   If tmp > 0 Then
      LblMax = CLng(tmp)
      Exit Sub
   End If
End If
MsgBox "This argument should be positive integer.", vbOKOnly, glFGTU
txtMaxLbls.SetFocus
End Sub

Private Sub txtMinLbls_LostFocus()
Dim tmp As String
tmp = txtMinLbls.Text
If IsNumeric(tmp) Then
   If tmp >= 0 Then
      LblMin = CLng(tmp)
      Exit Sub
   End If
End If
MsgBox "This argument should be non-negative integer.", vbOKOnly, glFGTU
txtMinLbls.SetFocus
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

Private Sub LoadMTDB(Optional blnForceReload As Boolean = False)
    Dim blnAMTsWereLoaded As Boolean, blnDBConnectionError As Boolean

    If ConfirmMassTagsAndInternalStdsLoaded(Me, CallerID, True, True, False, blnForceReload, 0, blnAMTsWereLoaded, blnDBConnectionError) Then
        lblMTStatus.Caption = ConstructMTStatusText(True)
    
        If Not CreateNewMTSearchObject() Then
           lblMTStatus.Caption = "Error creating search object."
        End If
    Else
        If blnDBConnectionError Then
            lblMTStatus.Caption = "Error loading MT tags: database connection error."
        Else
            lblMTStatus.Caption = "Error loading MT tags: no valid MT tags were found (possibly missing NET values)"
        End If
    End If
    
End Sub

Private Sub ShowHidePNNLMenus()
    Dim blnVisible As Boolean
    blnVisible = Not APP_BUILD_DISABLE_MTS
    
    mnuExpMTDB.Visible = blnVisible
    mnuMTLoadMT.Visible = blnVisible
End Sub

Private Sub UpdateStatus(ByVal Msg As String)
lblStatus.Caption = Msg
DoEvents
End Sub

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

Private Sub DoThePair_Mod(ByVal PairInd As Long)
'---------------------------------------------------------------
'finds all matching identifications for pair with index PairInd
'Search all class members of the light pair member for matching
'MT tags; criteria includes molecular mass, elution time and
'number of N atoms and Cys - compared with number of pair deltas
'NOTE: Assumption is that MT tags have modified molecular mass
'      In this case same peptide modified by 2 and 3 PEO labels
'      represent different MT tag; therefore match in this case
'      for number of Cys has to be exact
'NOTE: MWs over 2500(5000) Da allow for N count error of +/-1(2)
'---------------------------------------------------------------
Dim ClsInd1 As Long         'class index of light pair member
Dim DltCnt As Long          'established delta count for this pair
Dim MW As Double            'mol. mass of current distribution
Dim Scan As Long            'scan of current distribution
Dim ET As Double
Dim MWAbsErr As Double      'absolute value of error allowed
Dim i As Long, j As Long
Dim TmpCnt As Long
Dim Hits() As Long

'temporary arrays to deal with identifications of current pair
'non-unique identifications are first collected and then uniquely
'sorted and scored in separate procedure so that arrays PairIDInd
'and PairScore could be redimensioned during this process
Dim PairIDCnt As Long       'count of identifications for current pair
Dim PairIDInd() As Long     'MT tags indices of identifications
Dim PairScore() As Double   'score for each identification

On Error GoTo err_DoThePair_Mod
'couple of shortcut variables
ClsInd1 = GelP_D_L(CallerID).Pairs(PairInd).P1
DltCnt = GelP_D_L(CallerID).Pairs(PairInd).P2DltCnt
PairIDCnt = 0
ReDim PairIDInd(100)     'should be more than enough; do not allow for
                         'more than 1000 identifications per pair
CheckNETEquationStatus

With GelUMC(CallerID).UMCs(ClsInd1)   'search all members of
   For i = 0 To .ClassCount - 1       'light pair member class
     Select Case .ClassMType(i)
     Case glCSType
          MW = GelData(CallerID).CSData(.ClassMInd(i)).AverageMW
          Scan = GelData(CallerID).CSData(.ClassMInd(i)).ScanNumber
     Case glIsoType
          MW = GetIsoMass(GelData(CallerID).IsoData(.ClassMInd(i)), samtDef.MWField)
          Scan = GelData(CallerID).IsoData(.ClassMInd(i)).ScanNumber
     End Select
     ET = ConvertScanToNET(Scan)
     'calculate absolute molecular mass tolerance
     Select Case samtDef.TolType
     Case gltPPM
        MWAbsErr = MW * samtDef.MWTol * glPPM
     Case gltABS
        MWAbsErr = samtDef.MWTol
     End Select
     'functions GetMTHits1/2 filter MT tags on mass and elution
     Select Case samtDef.NETorRT
     Case glAMT_NET
        TmpCnt = GetMTHits1(MW, MWAbsErr, ET, samtDef.NETTol, Hits())
     Case glAMT_RT_or_PNET
        TmpCnt = GetMTHits2(MW, MWAbsErr, ET, samtDef.NETTol, Hits())
     End Select
     If TmpCnt > 0 Then
        For j = 0 To TmpCnt - 1
            'accept only identifications with correct number of N atoms
            'allow for +/-1(2) error for masses over 2500(5000) Da
            If MW > 5000 Then
               If Abs(AMTData(Hits(j)).CNT_N - DltCnt) <= 2 Then
                  PairIDCnt = PairIDCnt + 1
                  PairIDInd(PairIDCnt - 1) = Hits(j)
               End If
            ElseIf MW > glN14_N15CorrMW Then
               If Abs(AMTData(Hits(j)).CNT_N - DltCnt) <= 1 Then
                  PairIDCnt = PairIDCnt + 1
                  PairIDInd(PairIDCnt - 1) = Hits(j)
               End If
            Else
               If AMTData(Hits(j)).CNT_N = DltCnt Then
                  PairIDCnt = PairIDCnt + 1
                  PairIDInd(PairIDCnt - 1) = Hits(j)
               End If
            End If
        Next j
     End If
   Next i
End With
'-----------------------------------------------------------------
'all identifications for PairInd are collected; now order them in
'unique identifications with scores and add it to all possible IDs
'-----------------------------------------------------------------
mgScores:
If PairIDCnt > 0 Then
   ReDim Preserve PairIDInd(PairIDCnt - 1)      'trim the array with ID
   ScorePairIDs PairIDCnt, PairIDInd(), PairScore()
   'memorize unique count with each pair
   PIDCnt(PairInd) = PairIDCnt
   'add unique identifications with scores to all ids, also
   'memorize first and last index of id block for current pair
   If PairIDCnt > 0 Then
      If UBound(mgPInd) < mgCnt + PairIDCnt Then    'add more room
         'make sure it is enough to accomodate current batch
         ReDim Preserve mgPInd(UBound(mgPInd) + PairIDCnt + 2000)
         ReDim Preserve mgIDInd(UBound(mgPInd))
         ReDim Preserve mgScore(UBound(mgPInd))
      End If
      PIDInd1(PairInd) = mgCnt            'first index
      'last index will remain -1 if no ids and PIDInd2(i)>=0
      'should always be checked when enumerating ids for pair
      For i = 0 To PairIDCnt - 1
          mgCnt = mgCnt + 1
          mgPInd(mgCnt - 1) = PairInd
          mgIDInd(mgCnt - 1) = PairIDInd(i)
          mgScore(mgCnt - 1) = PairScore(i)
          PIDInd2(PairInd) = mgCnt - 1    'last index
      Next i
   End If
End If
Exit Sub

err_DoThePair_Mod:
Select Case Err.Number
Case 9  'make more room in arrays and resume; do not allow more than
    If PairIDCnt > 1000 Then        '1000 identifications per pair
       Resume mgScores
    Else
       ReDim Preserve PairIDInd(PairIDCnt + 100)
       Resume
    End If
Case Else
    LogErrors Err.Number, "frmSearchMTPairs_PEO_DoThePair_Mod"
End Select
End Sub


Private Sub DoThePair_NotMod(ByVal PairInd As Long)
'---------------------------------------------------------------
'finds all matching identifications for pair with index PairInd
'Search all class members of the light pair member for matching
'MT tags; criteria includes molecular mass, elution time and
'number of N atoms and Cys - compared with number of pair deltas
'NOTE: Assumption is that MT tags represent not modified
'peptides; therefore we need to correct peak mass for assumed
'number of labels and see if it matches any MT tag
'NOTE: MWs over 2500(5000) Da allow for N count error of +/-1(2)
'---------------------------------------------------------------
Dim ClsInd1 As Long         'class index of light pair member
Dim DltCnt As Long          'established delta count for this pair
Dim MW As Double            'mol. mass of current distribution
Dim MW1 As Double           'assumed non modified mass
Dim Scan As Long            'scan of current distribution
Dim ET As Double
Dim MWAbsErr As Double      'absolute value of error allowed
Dim i As Long, j As Long, k As Long
Dim TmpCnt As Long
Dim Hits() As Long

'temporary arrays to deal with identifications of current pair
'non-unique identifications are first collected and then uniquely
'sorted and scored in separate procedure so that arrays PairIDInd
'and PairScore could be redimensioned during this process
Dim PairIDCnt As Long       'count of identifications for current pair
Dim PairIDInd() As Long     'MT tags indices of identifications
Dim PairScore() As Double   'score for each identification

Dim PairIDOK As Boolean
On Error GoTo err_DoThePair_Mod
'couple of shortcut variables
ClsInd1 = GelP_D_L(CallerID).Pairs(PairInd).P1
DltCnt = GelP_D_L(CallerID).Pairs(PairInd).P2DltCnt
PairIDCnt = 0
ReDim PairIDInd(100)     'should be more than enough; do not allow for
                         'more than 1000 identifications per pair
CheckNETEquationStatus

With GelUMC(CallerID).UMCs(ClsInd1)   'search all members of
   For i = 0 To .ClassCount - 1       'light pair member class
     Select Case .ClassMType(i)
     Case glCSType
          MW = GelData(CallerID).CSData(.ClassMInd(i)).AverageMW
          Scan = GelData(CallerID).CSData(.ClassMInd(i)).ScanNumber
     Case glIsoType
          MW = GetIsoMass(GelData(CallerID).IsoData(.ClassMInd(i)), samtDef.MWField)
          Scan = GelData(CallerID).IsoData(.ClassMInd(i)).ScanNumber
     End Select
     ET = ConvertScanToNET(Scan)
     'assume any reasonable number of labels
     For k = LblMin To LblMax
        MW1 = MW - k * LblMW        'calculate peptide MW from assumption that
                                    'it is labeled with k labels
        If MW1 >= MIN_MOL_MASS Then
          'calculate absolute molecular mass tolerance
          Select Case samtDef.TolType
          Case gltPPM
            MWAbsErr = MW1 * samtDef.MWTol * glPPM
          Case gltABS
            MWAbsErr = samtDef.MWTol
          End Select
          'functions GetMTHits1/2 filter MT tags on mass and elution
          Select Case samtDef.NETorRT
          Case glAMT_NET
            TmpCnt = GetMTHits1(MW1, MWAbsErr, ET, samtDef.NETTol, Hits())
          Case glAMT_RT_or_PNET
            TmpCnt = GetMTHits2(MW1, MWAbsErr, ET, samtDef.NETTol, Hits())
          End Select
          If TmpCnt > 0 Then        'we have potential identifications
          'still need to test on N & Cys counts
             For j = 0 To TmpCnt - 1
                 PairIDOK = True                'be optimistic by default
                 If MW1 > 2 * glN14_N15CorrMW Then         'N count test
                    If Abs(AMTData(Hits(j)).CNT_N - DltCnt) > 2 Then PairIDOK = False
                 ElseIf MW1 > glN14_N15CorrMW Then
                    If Abs(AMTData(Hits(j)).CNT_N - DltCnt) > 1 Then PairIDOK = False
                 Else
                    If AMTData(Hits(j)).CNT_N <> DltCnt Then PairIDOK = False
                 End If
                 If k > AMTData(Hits(j)).CNT_Cys Then PairIDOK = False   'Cys count test
                 If PairIDOK Then
                    PairIDCnt = PairIDCnt + 1
                    PairIDInd(PairIDCnt - 1) = Hits(j)
                 End If
             Next j
          End If
        End If
     Next k
   Next i
End With
'-----------------------------------------------------------------
'all identifications for PairInd are collected; now order them in
'unique identifications with scores and add it to all possible IDs
'-----------------------------------------------------------------
mgScores:
If PairIDCnt > 0 Then
   ReDim Preserve PairIDInd(PairIDCnt - 1)      'trim the array with ID
   ScorePairIDs PairIDCnt, PairIDInd(), PairScore()
   'memorize unique count with each pair
   PIDCnt(PairInd) = PairIDCnt
   'add unique identifications with scores to all ids, also
   'memorize first and last index of id block for current pair
   If PairIDCnt > 0 Then
      If UBound(mgPInd) < mgCnt + PairIDCnt Then    'add more room
         'make sure it is enough to accomodate current batch
         ReDim Preserve mgPInd(UBound(mgPInd) + PairIDCnt + 2000)
         ReDim Preserve mgIDInd(UBound(mgPInd))
         ReDim Preserve mgScore(UBound(mgPInd))
      End If
      PIDInd1(PairInd) = mgCnt            'first index
      'last index will remain -1 if no ids and PIDInd2(i)>=0
      'should always be checked when enumerating ids for pair
      For i = 0 To PairIDCnt - 1
          mgCnt = mgCnt + 1
          mgPInd(mgCnt - 1) = PairInd
          mgIDInd(mgCnt - 1) = PairIDInd(i)
          mgScore(mgCnt - 1) = PairScore(i)
          PIDInd2(PairInd) = mgCnt - 1    'last index
      Next i
   End If
End If
Exit Sub

err_DoThePair_Mod:
Select Case Err.Number
Case 9  'make more room in arrays and resume; do not allow more than
    If PairIDCnt > 1000 Then        '1000 identifications per pair
       Resume mgScores
    Else
       ReDim Preserve PairIDInd(PairIDCnt + 100)
       Resume
    End If
Case Else
    LogErrors Err.Number, "frmSearchMTPairs_PEO_DoThePair_Mod"
End Select
End Sub



Private Sub ScorePairIDs(ByRef PCnt As Long, _
                         ByRef PID() As Long, _
                         ByRef PScores() As Double)
'--------------------------------------------------------
'does unique count of identifications found in array PID
'after this procedure PID array will contain only unique
'identifications
'NOTE: score is just unique count of each identification
'NOTE: results are returned in the same container arrays
'NOTE: this procedure is called only for PCnt>0
'--------------------------------------------------------
Dim UnqCnt As Long
Dim TmpID() As Long
Dim CurrID As Long
Dim i As Long, j As Long
On Error Resume Next

TmpID = PID     'copy ID data to temporary array
UnqCnt = 0
'zero scores and unique identifications
ReDim PID(PCnt - 1)
ReDim PScores(PCnt - 1)

For i = 0 To PCnt - 1
    CurrID = TmpID(i)
    For j = 0 To UnqCnt - 1
        If CurrID = PID(j) Then
           PScores(j) = PScores(j) + 1
           Exit For
        End If
    Next j
    If j > UnqCnt - 1 Then          'CurrID not found among unique
       UnqCnt = UnqCnt + 1          'IDs - add it
       PID(UnqCnt - 1) = CurrID
       PScores(UnqCnt - 1) = 1
    End If
Next i
'truncate the unique counts
If UnqCnt > 0 Then
   PCnt = UnqCnt
   ReDim Preserve PID(UnqCnt - 1)
   ReDim Preserve PScores(UnqCnt - 1)
Else                            'should not happen but...
   PCnt = -1
   Erase PID
   Erase PScores
End If
End Sub



Private Sub WriteReportHeader(ByVal TypeDescription As String)
'--------------------------------------------------------------------
'write report header block
'--------------------------------------------------------------------
ts.WriteLine "Generated by: " & GetMyNameVersion() & " on " & Now()
ts.WriteLine "Gel File: " & GelBody(CallerID).Caption
ts.WriteLine "Reporting identification for PEO labeled N14/N15 LC-MS Feature pairs"
ts.WriteLine TypeDescription
ts.WriteLine
ts.WriteLine "Total data points: " & GelData(CallerID).DataLines
ts.WriteLine "Total N14/N15 pairs: " & GelP_D_L(CallerID).PCnt
ts.WriteLine "Total MT tags: " & AMTCnt
ts.WriteLine
ts.WriteLine "UMC L Ind" & glARG_SEP & "L MW" & glARG_SEP & "L Abu" _
        & glARG_SEP & "L FN1" & glARG_SEP & "L FN2" & glARG_SEP _
           & "UMC H Ind" & glARG_SEP & "H MW" & glARG_SEP & "H Abu" _
        & glARG_SEP & "Dlt.Cnt" & glARG_SEP & "H FN1" & glARG_SEP _
        & "H FN2" & glARG_SEP & "ER" & glARG_SEP & "ER_StDev" & glARG_SEP & "ER_ChargeStateBasisCount" & glARG_SEP & "ER_MemberCount" & glARG_SEP & "ID" & glARG_SEP _
        & "ID MW" & glARG_SEP & "ID Score"
End Sub

Private Sub ReportPair(ByVal PairInd As Long)
'----------------------------------------------------------------
'writes lines of report(all identifications) for PairInd pair
'if there is no identification for this pair just write pair info
'----------------------------------------------------------------
Dim i As Long
Dim SP As String                    'pair part of line
Dim sID As String                   'ID part of line
On Error Resume Next
'extract pairs information
With GelP_D_L(CallerID).Pairs(PairInd)
  SP = .P1 & glARG_SEP & ClsStat(.P1, ustClassMW) _
        & glARG_SEP & ClsStat(.P1, ustClassIntensity) & glARG_SEP _
        & ClsStat(.P1, ustScanStart) & glARG_SEP & ClsStat(.P1, ustScanEnd) _
        & glARG_SEP & .P2 & glARG_SEP & ClsStat(.P2, ustClassMW) _
        & glARG_SEP & ClsStat(.P2, ustClassIntensity) & glARG_SEP _
        & .P2DltCnt & glARG_SEP & ClsStat(.P2, ustScanStart) _
        & glARG_SEP & ClsStat(.P2, ustScanEnd) & glARG_SEP _
        & .ER & glARG_SEP & .ERStDev & glARG_SEP & .ERChargeStateBasisCount & glARG_SEP & .ERMemberBasisCount
End With
If PIDCnt(PairInd) < 0 Then         'error during pair identification
   ts.WriteLine SP & glARG_SEP & "Error during identification"
ElseIf PIDCnt(PairInd) = 0 Then     'no id for this pair
   ts.WriteLine SP & glARG_SEP & "Unidentified"
Else                                'identified
   For i = PIDInd1(PairInd) To PIDInd2(PairInd)
       sID = glARG_SEP & Trim(AMTData(mgIDInd(i)).ID) & glARG_SEP _
            & AMTData(mgIDInd(i)).MW & glARG_SEP & mgScore(i)
       ts.WriteLine SP & sID
   Next i
End If
End Sub


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


Private Function Elution(FN As Long, MinFN As Long, MaxFN As Long)
'---------------------------------------------------
'this function does not care are we using NET or RT
'---------------------------------------------------
VarVals(1) = FN
VarVals(2) = MinFN
VarVals(3) = MaxFN
Elution = MyExprEva.ExprVal(VarVals())
End Function

' September 2004: Unused Function
''Public Function ExportIDPairsToMTDB(Optional ByRef lngErrorNumber As Long, Optional ByRef lngMDID As Long) As String
'''---------------------------------------------------
'''this is simple but long procedure of exporting data
'''results to Organism MT tag database associated
'''lngErrorNumber will contain the error number, if an error occurs
'''---------------------------------------------------
''Dim i As Long, j As Long, k As Long
''Dim ExpCnt As Long
'''ADO objects for stored procedure adding Match Making row
''Dim cnNew As New ADODB.Connection
'''ADO objects for stored procedure that adds FTICR peak rows
''Dim cmdPutNewPeak As New ADODB.Command
''Dim prmMMDID As New ADODB.Parameter
''Dim prmFTICRID As New ADODB.Parameter
''Dim prmFTICRType As New ADODB.Parameter
''Dim prmScanNumber As New ADODB.Parameter
''Dim prmChargeState As New ADODB.Parameter
''Dim prmMonoisotopicMass As New ADODB.Parameter
''Dim prmAbundance As New ADODB.Parameter
''Dim prmFit As New ADODB.Parameter
''Dim prmExpressionRatio As New ADODB.Parameter
''Dim prmLckID As New ADODB.Parameter
''Dim prmFreqShift As New ADODB.Parameter
''Dim prmMassCorrection As New ADODB.Parameter
''Dim prmMassTagID As New ADODB.Parameter
''Dim prmResType As New ADODB.Parameter
''Dim prmHitsCount As New ADODB.Parameter
''Dim prmUMCInd As New ADODB.Parameter
''Dim prmUMCFirstScan As New ADODB.Parameter
''Dim prmUMCLastScan As New ADODB.Parameter
''Dim prmUMCCount As New ADODB.Parameter
''Dim prmUMCAbundance As New ADODB.Parameter
''Dim prmUMCBestFit As New ADODB.Parameter
''Dim prmUMCAvgMW As New ADODB.Parameter
''Dim prmPairInd As New ADODB.Parameter
''
''Dim IndL As Long        'index in UMC of light pair member
''Dim IndH As Long        'index in UMC of heavy pair member
''Dim nInd As Long        'current numeric index - this is used only as a shortcut
''
''Dim UMCStat2() As Double        'UMC Statistics(needed for export function)
''Dim UMCCnt2 As Long
''
''On Error GoTo err_ExportMTDB
''
''UpdateStatus "Calculating statistics for UMC ..."
''UMCCnt2 = UMCStatistics2(CallerID, UMCStat2)
''If UMCCnt2 <= 0 Then
''   ExportIDPairsToMTDB = "Error calculating statistics for UMC. Export aborted."
''   Exit Function
''End If
''
''UpdateStatus "Exporting ..."
''If Not EstablishConnection(cnNew, GelAnalysis(CallerID).MTDB.cn.ConnectionString, False) Then
''    Debug.Assert False
''    lngErrorNumber = -1
''    ExportIDPairsToMTDB = "Error: Unable to establish a connection to the database"
''    Exit Function
''End If
''
'''first write new analysis in T_Match_Making_Description table
''AddEntryToMatchMakingDescriptionTable cnNew, lngMDID, ExpAnalysisSPName, CallerID, mgCnt, GelData(CallerID).CustomNETsDefined, False, strIniFileName
''
''' MonroeMod
''AddToAnalysisHistory CallerID, "Exported PEO Identification Pairs results to database (" & ExtractDBNameFromConnectionString(GelAnalysis(CallerID).MTDB.cn.ConnectionString) & "); MMD_ID = " & lngMDID
''AddToAnalysisHistory CallerID, "Export to MMD table details: Reference Job = " & GelAnalysis(CallerID).MD_Reference_Job & "; MD_File = " & GelAnalysis(CallerID).MD_file
''
''' Initialize the SP
''InitializeSPCommand cmdPutNewPeak, cnNew, ExpPeakSPName
''
''Set prmMMDID = cmdPutNewPeak.CreateParameter("MMDID", adInteger, adParamInput, , lngMDID)
''cmdPutNewPeak.Parameters.Append prmMMDID
''Set prmFTICRID = cmdPutNewPeak.CreateParameter("FTICRID", adVarChar, adParamInput, 50, Null)
''cmdPutNewPeak.Parameters.Append prmFTICRID
''Set prmFTICRType = cmdPutNewPeak.CreateParameter("FTICRType", adTinyInt, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmFTICRType
''Set prmScanNumber = cmdPutNewPeak.CreateParameter("ScanNumber", adInteger, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmScanNumber
''Set prmChargeState = cmdPutNewPeak.CreateParameter("ChargeState", adSmallInt, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmChargeState
''Set prmMonoisotopicMass = cmdPutNewPeak.CreateParameter("MonoisotopicMass", adDouble, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmMonoisotopicMass
''Set prmAbundance = cmdPutNewPeak.CreateParameter("Abundance", adDouble, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmAbundance
''Set prmFit = cmdPutNewPeak.CreateParameter("Fit", adDouble, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmFit
''Set prmExpressionRatio = cmdPutNewPeak.CreateParameter("ExpressionRatio", adDouble, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmExpressionRatio
''Set prmLckID = cmdPutNewPeak.CreateParameter("LckID", adInteger, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmLckID
''Set prmFreqShift = cmdPutNewPeak.CreateParameter("FreqShift", adDouble, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmFreqShift
''Set prmMassCorrection = cmdPutNewPeak.CreateParameter("MassCorrection", adDouble, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmMassCorrection
''Set prmMassTagID = cmdPutNewPeak.CreateParameter("MassTagID", adInteger, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmMassTagID
''Set prmResType = cmdPutNewPeak.CreateParameter("Type", adInteger, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmResType
''Set prmHitsCount = cmdPutNewPeak.CreateParameter("HitCount", adInteger, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmHitsCount
''Set prmUMCInd = cmdPutNewPeak.CreateParameter("UMCInd", adInteger, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmUMCInd
''Set prmUMCFirstScan = cmdPutNewPeak.CreateParameter("UMCFirstScan", adInteger, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmUMCFirstScan
''Set prmUMCLastScan = cmdPutNewPeak.CreateParameter("UMCLastScan", adInteger, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmUMCLastScan
''Set prmUMCCount = cmdPutNewPeak.CreateParameter("UMCCount", adInteger, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmUMCCount
''Set prmUMCAbundance = cmdPutNewPeak.CreateParameter("UMCAbundance", adDouble, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmUMCAbundance
''Set prmUMCBestFit = cmdPutNewPeak.CreateParameter("UMCBestFit", adDouble, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmUMCBestFit
''Set prmUMCAvgMW = cmdPutNewPeak.CreateParameter("UMCAvgMW", adDouble, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmUMCAvgMW
''Set prmPairInd = cmdPutNewPeak.CreateParameter("PairInd", adInteger, adParamInput, , 0)
''cmdPutNewPeak.Parameters.Append prmPairInd
''
'''now export data; there are some pairs if we are here
''ExpCnt = 0
''For i = 0 To PCount - 1
''    If PIDCnt(i) > 0 Then       'this pair is identified
''       IndL = GelP_D_L(CallerID).Pairs(i).P1
''       IndH = GelP_D_L(CallerID).Pairs(i).P2
''       'common values for all peaks
''       prmExpressionRatio.value = GelP_D_L(CallerID).Pairs(i).ER
''
''       prmPairInd.value = i
''       prmHitsCount.value = PIDCnt(i)   'number of different identifications
''
''       'report peaks from class IndL as light and ...
''       prmUMCInd.value = IndL
''       prmResType.value = FPR_Type_N14_N15_L
''       prmUMCFirstScan.value = UMCStat2(IndL, 2)
''       prmUMCLastScan.value = UMCStat2(IndL, 3)
''       prmUMCCount.value = UMCStat2(IndL, 8)
''       prmUMCAbundance.value = UMCStat2(IndL, 4)
''       prmUMCBestFit.value = UMCStat2(IndL, 6)
''       prmUMCAvgMW.value = UMCStat2(IndL, 1)
''
''       With GelUMC(CallerID).UMCs(IndL)
''           For j = 0 To .ClassCount - 1
''               nInd = .ClassMInd(j)
''               prmFTICRID.value = nInd
''               prmFTICRType.value = .ClassMType(j)
''               Select Case .ClassMType(j)
''               Case glCSType
''                 With GelData(CallerID)
''                    prmScanNumber.value = .CSData(nInd).ScanNumber
''                    prmChargeState.value = .CSData(nInd).Charge
''                    prmMonoisotopicMass.value = .CSData(nInd).AverageMW
''                    prmAbundance.value = .CSData(nInd).Abundance
''                    prmFit.value = .CSData(nInd).MassStDev     'standard deviation
''                    With GelLM(CallerID)
''                      If .CSCnt > 0 Then
''                         prmLckID.value = .CSLckID(nInd)
''                         prmFreqShift.value = .CSFreqShift(nInd)
''                         prmMassCorrection.value = .CSMassCorrection(nInd)
''                      End If
''                    End With
''                 End With
''               Case glIsoType
''                 With GelData(CallerID)
''                    prmScanNumber.value = .IsoData(nInd).ScanNumber
''                    prmChargeState.value = .IsoData(nInd).Charge
''                    prmMonoisotopicMass.value = .IsoData(nInd).MonoisotopicMW
''                    prmAbundance.value = .IsoData(nInd).Abundance
''                    prmFit.value = .IsoData(nInd).Fit
''                    With GelLM(CallerID)
''                      If .IsoCnt > 0 Then
''                        prmLckID.value = .IsoLckID(nInd)
''                        prmFreqShift.value = .IsoFreqShift(nInd)
''                        prmMassCorrection.value = .IsoMassCorrection(nInd)
''                      End If
''                    End With
''                 End With
''               End Select
''
''               'export all identifications
''               For k = PIDInd1(i) To PIDInd2(i)
''                   prmMassTagID = AMTData(mgIDInd(k)).ID     'MT tag ID
''                   cmdPutNewPeak.Execute
''                   ExpCnt = ExpCnt + 1
''               Next k
''           Next j
''       End With
''
''       '...all from IndH as heavy members
''       prmUMCInd.value = IndH
''       prmResType.value = FPR_Type_N14_N15_H
''       prmUMCFirstScan.value = UMCStat2(IndH, 2)
''       prmUMCLastScan.value = UMCStat2(IndH, 3)
''       prmUMCCount.value = UMCStat2(IndH, 8)
''       prmUMCAbundance.value = UMCStat2(IndH, 4)
''       prmUMCBestFit.value = UMCStat2(IndH, 6)
''       prmUMCAvgMW.value = UMCStat2(IndH, 1)
''
''       With GelUMC(CallerID).UMCs(IndH)
''           For j = 0 To .ClassCount - 1
''               nInd = .ClassMInd(j)
''               prmFTICRID.value = nInd
''               prmFTICRType.value = .ClassMType(j)
''               Select Case .ClassMType(j)
''               Case glCSType
''                 With GelData(CallerID)
''                    prmScanNumber.value = .CSData(nInd).ScanNumber
''                    prmChargeState.value = .CSData(nInd).Charge
''                    prmMonoisotopicMass.value = .CSData(nInd).AverageMW
''                    prmAbundance.value = .CSData(nInd).Abundance
''                    prmFit.value = .CSData(nInd).MassStDev     'standard deviation
''                    With GelLM(CallerID)
''                      If .CSCnt > 0 Then
''                         prmLckID.value = .CSLckID(nInd)
''                         prmFreqShift.value = .CSFreqShift(nInd)
''                         prmMassCorrection.value = .CSMassCorrection(nInd)
''                      End If
''                    End With
''                 End With
''               Case glIsoType
''                 With GelData(CallerID)
''                    prmScanNumber.value = .IsoData(nInd).ScanNumber
''                    prmChargeState.value = .IsoData(nInd).Charge
''                    prmMonoisotopicMass.value = .IsoData(nInd).MonoisotopicMW
''                    prmAbundance.value = .IsoData(nInd).Abundance
''                    prmFit.value = .IsoData(nInd).Fit
''                    With GelLM(CallerID)
''                      If .IsoCnt > 0 Then
''                        prmLckID.value = .IsoLckID(nInd)
''                        prmFreqShift.value = .IsoFreqShift(nInd)
''                        prmMassCorrection.value = .IsoMassCorrection(nInd)
''                      End If
''                    End With
''                 End With
''               End Select
''
''               'export all identifications
''               For k = PIDInd1(i) To PIDInd2(i)
''                   prmMassTagID = AMTData(mgIDInd(k)).ID     'MT tag ID
''                   cmdPutNewPeak.Execute
''                   ExpCnt = ExpCnt + 1
''               Next k
''           Next j
''       End With
''    End If
''Next i
''
''' MonroeMod
''AddToAnalysisHistory CallerID, "Export to Peak Results table details: Pairs Match Count = " & ExpCnt
''
''UpdateStatus "Export done."
''ExportIDPairsToMTDB = ExpCnt & " associations between MT tags and FTICR peaks exported."
''Set cmdPutNewPeak.ActiveConnection = Nothing
''cnNew.Close
''lngErrorNumber = 0
''Exit Function
''
''err_ExportMTDB:
''ExportIDPairsToMTDB = "Error: " & Err.Number & vbCrLf & Err.Description
''lngErrorNumber = Err.Number
''If Not cnNew Is Nothing Then cnNew.Close
''End Function

Private Sub PickParameters()
Call txtLblMass_LostFocus
Call txtMaxLbls_LostFocus
Call txtMinLbls_LostFocus
Call txtMWTol_LostFocus
Call txtNETTol_LostFocus
Call txtNETFormula_LostFocus
End Sub
