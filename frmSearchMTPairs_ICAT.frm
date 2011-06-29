VERSION 5.00
Begin VB.Form frmSearchMTPairs_ICAT 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search MT Tag Database For Potential ICAT Pairs"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraMTassumption 
      BackColor       =   &H00FFFFC0&
      Caption         =   "MT Tags"
      Height          =   1215
      Left            =   2160
      TabIndex        =   17
      Top             =   480
      Width           =   1455
      Begin VB.OptionButton optAssumeLabels 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Labeled"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "MT Tags are labeled (mass modified)"
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optAssumeLabels 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Not Labeled"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "MT Tags are not labeled (mass represents mass of the clean sequence)"
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame fraNET 
      BackColor       =   &H00FFFFC0&
      Caption         =   "NET  Calculation"
      Height          =   1455
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   4695
      Begin VB.CheckBox chkUseUMCConglomerateNET 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Use Class NET for UMC's"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         ToolTipText     =   $"frmSearchMTPairs_ICAT.frx":0000
         Top             =   1095
         Width           =   2205
      End
      Begin VB.TextBox txtNETTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3600
         TabIndex        =   21
         Text            =   "0.1"
         Top             =   1080
         Width           =   615
      End
      Begin VB.OptionButton optNETorRT 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Pred. NET"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   15
         ToolTipText     =   "Use NET calculated only from Sequest ""first choice"" peptides"
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optNETorRT 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Obs. NET"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   14
         ToolTipText     =   "Use NET calculated from all peptides of MT Tags"
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox txtNETFormula 
         Height          =   285
         Left            =   1560
         TabIndex        =   13
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
         Top             =   1095
         Width           =   855
      End
      Begin VB.Label lblNETFormula 
         BackStyle       =   0  'Transparent
         Caption         =   "&Formula  F(FN, MinFN, MaxFN)"
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame fraMWTolerance 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Mol. Mass Tolerance"
      Height          =   1215
      Left            =   3720
      TabIndex        =   6
      Top             =   480
      Width           =   2055
      Begin VB.OptionButton optTolType 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Dalton"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optTolType 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&ppm"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtMWTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Text            =   "10"
         Top             =   640
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tolerance"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame fraMWField 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Mol. Mass Field"
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1935
      Begin VB.OptionButton optMWField 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&The Most Abundant"
         Height          =   255
         Index           =   2
         Left            =   80
         TabIndex        =   5
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton optMWField 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Monoisotopic"
         Height          =   255
         Index           =   1
         Left            =   80
         TabIndex        =   4
         Top             =   540
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optMWField 
         BackColor       =   &H00FFFFC0&
         Caption         =   "A&verage"
         Height          =   255
         Index           =   0
         Left            =   80
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Label lblETType 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3720
      TabIndex        =   20
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5280
      Picture         =   "frmSearchMTPairs_ICAT.frx":0097
      ToolTipText     =   "Double-click for short info on this procedure"
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3360
      Width           =   3495
   End
   Begin VB.Label lblMTStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MTDB Status:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   0
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
Attribute VB_Name = "frmSearchMTPairs_ICAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'search of MT tag database for pairs member of the isotopic
'Mod pairs
'assumption is that pairs are calculated for UMC and that
'search of the database can be performed with loose tolerance
'with additional criteria of matching number of Cys with delta
'in molecular mass for established pairs;
'------------------------------------------------------------
'NOTE: search is always done for light pair members
'------------------------------------------------------------
'created: 11/12/2001 nt
'last modified: 07/29/2002 nt
'------------------------------------------------------------
'Asumption have to be made whether MT tags are Mod
'(modified in mass) or not. Search will produce different
'results with different choice of this option!!!!!!!!!!!!!!!
'If assumption is that MT tags are modified then search is
'simple search of database elements. If assumption is that
'MT tags are not modified(mass of the clean sequence) then
'pairs has to be modifed for search to produce valid results.
'------------------------------------------------------------
'If MT tags are modified then look for the light member of
'the pair and compare pair delta with number of cysteines in
'MT tag - if it is less or equal declare hit;
'------------------------------------------------------------
'Other thing in work here is that this procedure can handle
'different type of pairing searches
'a) pairs can be Delta pairs in which case Labels have value
'   0 but number of labels is determined with number of deltas
'   if this is the case number of labels has to be the same in
'   heavy and light member
'b) Mod pairs in which case delta has value 0 and number
'   of labels can be different in heavy and light member
'Procedures to determine match with database differs based on
'pairs type
'------------------------------------------------------------
Option Explicit

'in this case CallerID is a public property
Public CallerID As Long

Dim bLoading As Boolean

Dim OldSearchFlag As Long

Dim bMTMod As Long

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
''Dim ExpPeakSPName As String                 ' Stored procedure AddFTICRPeak; Unused variable

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
bMTMod = 0           'MT tags not Mod

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
tmp = tmp & "light members determines number of ICAT labels in underlying" & vbCrLf
tmp = tmp & "peptide. Idea is to search MT tag DB with loose(25ppm)" & vbCrLf
tmp = tmp & "tolerance and select as possible identification those" & vbCrLf
tmp = tmp & "with matching numbers." & vbCrLf
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
'---------------------------------------------------------
'triggers export of identified pairs to MT tag database
'also gives user a chance to change their mind
'---------------------------------------------------------
Dim Res As Long
Dim WhatHappened As String
If mgCnt > 0 Then
     MsgBox "Database export is no longer supported using this window.", vbInformation + vbOKOnly, "Error"
    
     ' September 2004: Unsupported code
''   Res = MsgBox(mgCnt & " identified pairs will be exported to MT tag database.", vbOKCancel, glFGTU)
''   If Res <> vbOK Then Exit Sub
''   UpdateStatus "Exporting ...."
''   Me.MousePointer = vbHourglass
''   WhatHappened = ExportIDPairsToMTDB
''   MsgBox WhatHappened, vbOKOnly, glFGTU
''   Me.MousePointer = vbDefault
''   UpdateStatus ""
Else
   MsgBox "No identified pairs found.", vbOKOnly, glFGTU
End If
End Sub

Private Sub mnuMT_Click()
Call PickParameters
End Sub

Private Sub mnuMTLoadLegacy_Click()
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

Private Sub mnuMTLoadMT_Click()
'---------------------------------------------
'load/reload MT tags
'---------------------------------------------
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
'--------------------------------------------------------
'search pairs in GelP_D_L structure for MT tags matches
'NOTE: error handling for search is in this procedure
'--------------------------------------------------------
Dim HitsCnt As Long
Dim eResponse As VbMsgBoxResult
Dim I As Long
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
   eResponse = MsgBox("Pairs identification found. If you continue current findings will be lost. Continue?", vbOKCancel, glFGTU)
   If eResponse <> vbOK Then Exit Sub
End If

GelData(CallerID).MostRecentSearchUsedSTAC = False

'number of pairs might change so better check every time
PCount = GelP_D_L(CallerID).PCnt
samtDef.Formula = Trim$(txtNETFormula.Text)
If PCount > 0 Then
   With GelP_D_L(CallerID)
      If ((.DltLblType = ptUMCDlt) Or (.DltLblType = ptUMCLbl)) Then
        If .SyncWithUMC Then
           Me.MousePointer = vbHourglass
           'reserve space for identifications per pair counts
           ReDim PIDCnt(PCount - 1)
           ReDim PIDInd1(PCount - 1)
           ReDim PIDInd2(PCount - 1)
           'set last index to -1 as indication if there were no ID
           For I = 0 To PCount - 1
               PIDInd2(I) = -1
           Next I
           mgCnt = 0
           'reserve initial space for 10000 identifications
           ReDim mgPInd(10000)
           ReDim mgIDInd(10000)
           ReDim mgScore(10000)
           'do identification pair by pair
           For I = 0 To PCount - 1
             'do not try if pair already excluded
             If .Pairs(I).STATE <> glPAIR_Exc Then
                UpdateStatus "Identifying pair " & (I + 1) & "/" & PCount
                'processing differs based on MT tags type(assumption)
                'and pairs type(delta or label pairs)
'For now only Dlt_Mod and Dlt_NotMod procedures work
                Select Case .DltLblType
                Case ptUMCDlt
                    'eliminate impossible pairs
                    If .Pairs(I).P2DltCnt * glICAT_Light >= GelUMC(CallerID).UMCs(I).ClassMW Then
                       .Pairs(I).STATE = glPAIR_Exc
                    Else
                       If bMTMod Then
                          DoThePair_Dlt_Mod (I)
                       Else
                          DoThePair_Dlt_NotMod (I)
                       End If
                    End If
                Case ptUMCLbl
                    If bMTMod Then
                       DoThePair_Lbl_Mod (I)
                    Else
                       DoThePair_Lbl_NotMod (I)
                    End If
                End Select
             End If
           Next I
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
           HitsCnt = -5               'pairs should be recalculated
        End If
      Else
        HitsCnt = -4                  'pairs type does not match
      End If
   End With
Else                                'no pairs found
   HitsCnt = -3
End If
Select Case HitsCnt
Case -1
   MsgBox "Error searching MT database.", vbOKOnly, glFGTU
Case -2
   MsgBox "Error in NET calculation formula.", vbOKOnly, glFGTU
   txtNETFormula.SetFocus
Case -3
   MsgBox "No pairs found. Make sure that one of the LC-MS Feature pairing functions is applied first.", vbOKOnly, glFGTU
Case -4
   MsgBox "Incorrect pairs type.", vbOKOnly, glFGTU
Case -5
   MsgBox "Pairs need to be recalculated. Close dialog and recalculate pairs, then return to this dialog.", vbOKOnly, glFGTU
Case Else
   MsgBox "MT tag hits: " & HitsCnt & " (non-unique)", vbOKOnly, glFGTU
    
    If GelAnalysis(CallerID).MD_Type = stNotDefined Or GelAnalysis(CallerID).MD_Type = stStandardIndividual Then
        ' Only update MD_Type if it is currently stStandardIndividual
        GelAnalysis(CallerID).MD_Type = stPairsICAT
    End If
    
    'MonroeMod
    GelSearchDef(CallerID).AMTSearchOnPairs = samtDef
    AddToAnalysisHistory CallerID, GetMassTagSearchSummaryText("Searched ICAT pairs for MT tags", HitsCnt, 0, 0, 0, samtDef, True, GelData(CallerID).CustomNETsDefined)
End Select

End Sub

Private Sub mnuRAllPairsAndIDs_Click()
'-------------------------------------
'report pairs and identifications
'-------------------------------------
Dim I As Long
Dim eResponse As VbMsgBoxResult
If mgCnt <= 0 Then
   eResponse = MsgBox("No identification found. Continue with generating report?", vbYesNo, glFGTU)
   If eResponse <> vbYes Then Exit Sub
End If
UpdateStatus "Generating report ...."
Me.MousePointer = vbHourglass
Set ts = fso.OpenTextFile(fname, ForWriting, True)
WriteReportHeader "All pairs and identifications"
For I = 0 To PCount - 1
    ReportPair I
Next I
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
Dim I As Long
On Error Resume Next
If mgCnt > 0 Then
   UpdateStatus "Generating report ..."
   Me.MousePointer = vbHourglass
   Set ts = fso.OpenTextFile(fname, ForWriting, True)
   WriteReportHeader "Identified pairs only"
   For I = 0 To PCount - 1
       If PIDCnt(I) > 0 Then ReportPair I
   Next I
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
Dim I As Long
On Error Resume Next
If PCount > 0 Then
   UpdateStatus "Generating report ..."
   Me.MousePointer = vbHourglass
   Set ts = fso.OpenTextFile(fname, ForWriting, True)
   WriteReportHeader "Unidentified pairs only"
   For I = 0 To PCount - 1
       If PIDCnt(I) <= 0 Then ReportPair I
   Next I
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
Dim I As Long
On Error Resume Next
If mgCnt > 0 Then
   UpdateStatus "Generating report ..."
   Me.MousePointer = vbHourglass
   Set ts = fso.OpenTextFile(fname, ForWriting, True)
   WriteReportHeader "Uniquely identified pairs only"
   For I = 0 To PCount - 1
       If PIDCnt(I) = 1 Then ReportPair I
   Next I
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
Dim I As Long, j As Long
On Error Resume Next

TmpID = PID     'copy ID data to temporary array
UnqCnt = 0
'zero scores and unique identifications
ReDim PID(PCnt - 1)
ReDim PScores(PCnt - 1)

For I = 0 To PCnt - 1
    CurrID = TmpID(I)
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
Next I
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
ts.WriteLine "Reporting identification for ICAT LC-MS Feature pairs"
ts.WriteLine TypeDescription
ts.WriteLine
ts.WriteLine "Total data points: " & GelData(CallerID).DataLines
ts.WriteLine "Total ICAT pairs: " & GelP_D_L(CallerID).PCnt
ts.WriteLine "Total MT tags: " & AMTCnt
ts.WriteLine
ts.WriteLine "UMC L Ind" & glARG_SEP & "L MW" & glARG_SEP & "L Abu" _
        & glARG_SEP & "L FN1" & glARG_SEP & "L FN2" & glARG_SEP _
           & "UMC H Ind" & glARG_SEP & "H MW" & glARG_SEP & "H Abu" _
        & glARG_SEP & "Dlt.Cnt" & glARG_SEP & "H FN1" & glARG_SEP _
        & "H FN2" & glARG_SEP & "ER" & glARG_SEP & "ID" & glARG_SEP _
        & "ID MW" & glARG_SEP & "ID Score"
End Sub

Private Sub ReportPair(ByVal PairInd As Long)
'----------------------------------------------------------------
'writes lines of report(all identifications) for PairInd pair
'if there is no identification for this pair just write pair info
'----------------------------------------------------------------
Dim I As Long
Dim SP As String                    'pair part of line
Dim sID As String                   'ID part of line
'extract pairs information
With GelP_D_L(CallerID).Pairs(PairInd)
  SP = .p1 & glARG_SEP & ClsStat(.p1, ustClassMW) _
        & glARG_SEP & ClsStat(.p1, ustClassIntensity) & glARG_SEP _
        & ClsStat(.p1, ustScanStart) & glARG_SEP & ClsStat(.p1, ustScanEnd) _
        & glARG_SEP & .p2 & glARG_SEP & ClsStat(.p2, ustClassMW) _
        & glARG_SEP & ClsStat(.p2, ustClassIntensity) & glARG_SEP _
        & .P2DltCnt & glARG_SEP & ClsStat(.p2, ustScanStart) _
        & glARG_SEP & ClsStat(.p2, ustScanEnd) & glARG_SEP _
        & .ER & glARG_SEP & .ERStDev & glARG_SEP & .ERChargeStateBasisCount & glARG_SEP & .ERMemberBasisCount & glARG_SEP
End With
If PIDCnt(PairInd) < 0 Then         'error during pair identification
   ts.WriteLine SP & glARG_SEP & "Error during identification"
ElseIf PIDCnt(PairInd) = 0 Then     'no id for this pair
   ts.WriteLine SP & glARG_SEP & "Unidentified"
Else                                'identified
   For I = PIDInd1(PairInd) To PIDInd2(PairInd)
       sID = glARG_SEP & Trim(AMTData(mgIDInd(I)).ID) & glARG_SEP _
            & AMTData(mgIDInd(I)).MW & glARG_SEP & mgScore(I)
       ts.WriteLine SP & sID
   Next I
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


Private Sub DoThePair_Dlt_Mod(ByVal PairInd As Long)
'---------------------------------------------------------------
'finds all matching identifications for pair with index PairInd
'Search all class members of the light pair member for matching
'MT tags; criteria includes molecular mass, elution time and
'number of Cys in peptide - compared with number of pair deltas
'Assumption is that MT tags have already modified peptide mass
'That also means that MT tags with same clean sequences and 2
'and 3 ICAT labels differs among themselves; this is one of more
'?nable decisions using mods as criteria of identification
'In this case everything is included in search!!!!!!!!!!!!!!!!!!
'---------------------------------------------------------------
Dim ClsInd1 As Long         'class index of light pair member
Dim DltCnt As Long          'established delta count for this pair
Dim MW As Double            'mol. mass of current distribution
Dim Scan As Long            'scan of current distribution
Dim ET As Double
Dim MWAbsErr As Double      'absolute value of error allowed
Dim I As Long, j As Long
Dim TmpCnt As Long
Dim Hits() As Long

'temporary arrays to deal with identifications of current pair
'non-unique identifications are first collected and then uniquely
'sorted and scored in separate procedure so that arrays PairIDInd
'and PairScore could be redimensioned during this process
Dim PairIDCnt As Long       'count of identifications for current pair
Dim PairIDInd() As Long     'MT tags indices of identifications
Dim PairScore() As Double   'score for each identification

On Error GoTo err_DoThePair_Dlt_Mod
'couple of shortcut variables
ClsInd1 = GelP_D_L(CallerID).Pairs(PairInd).p1
DltCnt = GelP_D_L(CallerID).Pairs(PairInd).P2DltCnt
PairIDCnt = 0
ReDim PairIDInd(100)     'should be more than enough; do not allow for
                         'more than 1000 identifications per pair
CheckNETEquationStatus

With GelUMC(CallerID).UMCs(ClsInd1)   'search all members of
   For I = 0 To .ClassCount - 1       'light pair member class
     Select Case .ClassMType(I)
     Case glCSType
          MW = GelData(CallerID).CSData(.ClassMInd(I)).AverageMW
          Scan = GelData(CallerID).CSData(.ClassMInd(I)).ScanNumber
     Case glIsoType
          MW = GetIsoMass(GelData(CallerID).IsoData(.ClassMInd(I)), samtDef.MWField)
          Scan = GelData(CallerID).IsoData(.ClassMInd(I)).ScanNumber
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
            'accept only IDs with possible number of deltas (# of Dlts <= # of Cys)
            ' ToDo for Weijun (Dec 2003): If DltCnt <= AMTData(Hits(j)).CNT_Cys and DltCnt = "# Modified Cysteines" Then
            If AMTData(Hits(j)).CNT_Cys >= DltCnt Then
               PairIDCnt = PairIDCnt + 1
               PairIDInd(PairIDCnt - 1) = Hits(j)
            End If
        Next j
     End If
   Next I
End With
'-----------------------------------------------------------------
'all identifications for PairInd are collected; now order them in
'unique identifications with scores and add it to all possible IDs
'-----------------------------------------------------------------
mgScores_Dlt_Mod:
AddPairIDs PairInd, PairIDCnt, PairIDInd(), PairScore()
Exit Sub

err_DoThePair_Dlt_Mod:
Select Case Err.Number
Case 9  'make more room in arrays and resume; do not allow more than
    If PairIDCnt > 1000 Then        '1000 identifications per pair
       Resume mgScores_Dlt_Mod
    Else
       ReDim Preserve PairIDInd(PairIDCnt + 100)
       Resume
    End If
Case Else
    LogErrors Err.Number, "frmSearchMTPairs_DoThePair_Dlt_Mod"
End Select
End Sub

Private Sub DoThePair_Dlt_NotMod(ByVal PairInd As Long)
'---------------------------------------------------------------
'see notes in DoThePair_Dlt_Mod - difference is that assumption
'in this procedures is that MT tags are not modified so light
'member mass has to be modified before search
'---------------------------------------------------------------
Dim ClsInd1 As Long         'class index of light pair member
Dim DltCnt As Long          'established delta count for this pair
Dim MW As Double            'mol. mass of current distribution
Dim Scan As Long            'scan of current distribution
Dim ET As Double
Dim MWAbsErr As Double      'absolute value of error allowed
Dim I As Long, j As Long
Dim TmpCnt As Long
Dim Hits() As Long

'temporary arrays to deal with identifications of current pair
'non-unique identifications are first collected and then uniquely
'sorted and scored in separate procedure so that arrays PairIDInd
'and PairScore could be redimensioned during this process
Dim PairIDCnt As Long       'count of identifications for current pair
Dim PairIDInd() As Long     'MT tags indices of identifications
Dim PairScore() As Double   'score for each identification

On Error GoTo err_DoThePair_Dlt_NotMod
'couple of shortcut variables
ClsInd1 = GelP_D_L(CallerID).Pairs(PairInd).p1
DltCnt = GelP_D_L(CallerID).Pairs(PairInd).P2DltCnt
PairIDCnt = 0
ReDim PairIDInd(100)     'should be more than enough; do not allow for
                         'more than 1000 identifications per pair
CheckNETEquationStatus

With GelUMC(CallerID).UMCs(ClsInd1)   'search all members of
   For I = 0 To .ClassCount - 1       'light pair member class
     Select Case .ClassMType(I)
     Case glCSType
          MW = GelData(CallerID).CSData(.ClassMInd(I)).AverageMW
          Scan = GelData(CallerID).CSData(.ClassMInd(I)).ScanNumber
     Case glIsoType
          MW = GetIsoMass(GelData(CallerID).IsoData(.ClassMInd(I)), samtDef.MWField)
          Scan = GelData(CallerID).IsoData(.ClassMInd(I)).ScanNumber
     End Select
     'modify mass since MT tags don't have modified mass
     MW = MW - DltCnt * glICAT_Light
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
        'accept only IDs with possible number of deltas (# of Dlts <= # of Cys)
            If AMTData(Hits(j)).CNT_Cys >= DltCnt Then
               PairIDCnt = PairIDCnt + 1
               PairIDInd(PairIDCnt - 1) = Hits(j)
            End If
        Next j
     End If
   Next I
End With
'-----------------------------------------------------------------
'all identifications for PairInd are collected; now order them in
'unique identifications with scores and add it to all possible IDs
'-----------------------------------------------------------------
mgScores_Dlt_NotMod:
AddPairIDs PairInd, PairIDCnt, PairIDInd(), PairScore()
Exit Sub

err_DoThePair_Dlt_NotMod:
Select Case Err.Number
Case 9  'make more room in arrays and resume; do not allow more than
    If PairIDCnt > 1000 Then        '1000 identifications per pair
       Resume mgScores_Dlt_NotMod
    Else
       ReDim Preserve PairIDInd(PairIDCnt + 100)
       Resume
    End If
Case Else
    LogErrors Err.Number, "frmSearchMTPairs_DoThePair_Dlt_NotMod"
End Select
End Sub


Private Sub DoThePair_Lbl_Mod(ByVal PairInd As Long)
'---------------------------------------------------
'same notes as in DoThePair_Dlt_Mod except:
'number of labels in light and heavy must be OK (it
'does not have to be the same though)
'---------------------------------------------------
Dim ClsInd1 As Long
Dim LLblCnt As Long
Dim HLblCnt As Long
Dim MW As Double
Dim Scan As Long
Dim ET As Double
Dim MWAbsErr As Double
Dim I As Long, j As Long
Dim TmpCnt As Long
Dim Hits() As Long

Dim PairIDCnt As Long
Dim PairIDInd() As Long
Dim PairScore() As Double

On Error GoTo err_DoThePair_Lbl_Mod
With GelP_D_L(CallerID).Pairs(PairInd)
   ClsInd1 = .p1
   LLblCnt = .P1LblCnt
   HLblCnt = .P2LblCnt
End With
PairIDCnt = 0
ReDim PairIDInd(100)
CheckNETEquationStatus

With GelUMC(CallerID).UMCs(ClsInd1)
   For I = 0 To .ClassCount - 1
     Select Case .ClassMType(I)
     Case glCSType
          MW = GelData(CallerID).CSData(.ClassMInd(I)).AverageMW
          Scan = GelData(CallerID).CSData(.ClassMInd(I)).ScanNumber
     Case glIsoType
          MW = GetIsoMass(GelData(CallerID).IsoData(.ClassMInd(I)), samtDef.MWField)
          Scan = GelData(CallerID).IsoData(.ClassMInd(I)).ScanNumber
     End Select
     ET = ConvertScanToNET(Scan)
     Select Case samtDef.TolType
     Case gltPPM
        MWAbsErr = MW * samtDef.MWTol * glPPM
     Case gltABS
        MWAbsErr = samtDef.MWTol
     End Select
     Select Case samtDef.NETorRT
     Case glAMT_NET
        TmpCnt = GetMTHits1(MW, MWAbsErr, ET, samtDef.NETTol, Hits())
     Case glAMT_RT_or_PNET
        TmpCnt = GetMTHits2(MW, MWAbsErr, ET, samtDef.NETTol, Hits())
     End Select
     If TmpCnt > 0 Then
        For j = 0 To TmpCnt - 1
           If ((AMTData(Hits(j)).CNT_Cys >= LLblCnt) And (AMTData(Hits(j)).CNT_Cys >= HLblCnt)) Then
              PairIDCnt = PairIDCnt + 1
              PairIDInd(PairIDCnt - 1) = Hits(j)
           End If
        Next j
     End If
   Next I
End With

mgScores_Lbl_Mod:
AddPairIDs PairInd, PairIDCnt, PairIDInd(), PairScore()
Exit Sub

err_DoThePair_Lbl_Mod:
Select Case Err.Number
Case 9
    If PairIDCnt > 1000 Then
       Resume mgScores_Lbl_Mod
    Else
       ReDim Preserve PairIDInd(PairIDCnt + 100)
       Resume
    End If
Case Else
    LogErrors Err.Number, "frmSearchMTPairs_DoThePair_Lbl_Mod"
End Select
End Sub

Private Sub DoThePair_Lbl_NotMod(ByVal PairInd As Long)
'---------------------------------------------------------------
'see notes in DoThePair_Dlt_NotMod
'---------------------------------------------------------------
Dim ClsInd1 As Long
Dim LLblCnt As Long
Dim HLblCnt As Long
Dim MW As Double
Dim Scan As Long
Dim ET As Double
Dim MWAbsErr As Double
Dim I As Long, j As Long
Dim TmpCnt As Long
Dim Hits() As Long

Dim PairIDCnt As Long
Dim PairIDInd() As Long
Dim PairScore() As Double

On Error GoTo err_DoThePair_Lbl_NotMod
With GelP_D_L(CallerID).Pairs(PairInd)
   ClsInd1 = .p1
   LLblCnt = .P1LblCnt
   HLblCnt = .P2LblCnt
End With
PairIDCnt = 0
ReDim PairIDInd(100)
CheckNETEquationStatus

With GelUMC(CallerID).UMCs(ClsInd1)
   For I = 0 To .ClassCount - 1
     Select Case .ClassMType(I)
     Case glCSType
          MW = GelData(CallerID).CSData(.ClassMInd(I)).AverageMW
          Scan = GelData(CallerID).CSData(.ClassMInd(I)).ScanNumber
     Case glIsoType
          MW = GetIsoMass(GelData(CallerID).IsoData(.ClassMInd(I)), samtDef.MWField)
          Scan = GelData(CallerID).IsoData(.ClassMInd(I)).ScanNumber
     End Select
     MW = MW - LLblCnt * glICAT_Light       'check only light mass
     ET = ConvertScanToNET(Scan)
     Select Case samtDef.TolType
     Case gltPPM
        MWAbsErr = MW * samtDef.MWTol * glPPM
     Case gltABS
        MWAbsErr = samtDef.MWTol
     End Select
     Select Case samtDef.NETorRT
     Case glAMT_NET
        TmpCnt = GetMTHits1(MW, MWAbsErr, ET, samtDef.NETTol, Hits())
     Case glAMT_RT_or_PNET
        TmpCnt = GetMTHits2(MW, MWAbsErr, ET, samtDef.NETTol, Hits())
     End Select
     If TmpCnt > 0 Then
        For j = 0 To TmpCnt - 1
            If ((AMTData(Hits(j)).CNT_Cys >= LLblCnt) And (AMTData(Hits(j)).CNT_Cys >= HLblCnt)) Then
               PairIDCnt = PairIDCnt + 1
               PairIDInd(PairIDCnt - 1) = Hits(j)
            End If
        Next j
     End If
   Next I
End With

mgScores_Lbl_NotMod:
AddPairIDs PairInd, PairIDCnt, PairIDInd(), PairScore()
Exit Sub

err_DoThePair_Lbl_NotMod:
Select Case Err.Number
Case 9
    If PairIDCnt > 1000 Then
       Resume mgScores_Lbl_NotMod
    Else
       ReDim Preserve PairIDInd(PairIDCnt + 100)
       Resume
    End If
Case Else
    LogErrors Err.Number, "frmSearchMTPairs_DoThePair_Lbl_NotMod"
End Select
End Sub


Private Sub AddPairIDs(PairInd As Long, tmpIdCnt As Long, _
                       tmpIdInd() As Long, tmpScore() As Double)
'-------------------------------------------------------------------------
'add pair identifications to arrays of all identifications
'this procedure is isolated to make code shorter
'-------------------------------------------------------------------------
Dim I As Long
On Error Resume Next
If tmpIdCnt > 0 Then
   ReDim Preserve tmpIdInd(tmpIdCnt - 1)      'trim the array with ID
   ScorePairIDs tmpIdCnt, tmpIdInd(), tmpScore()
   'memorize unique count with each pair
   PIDCnt(PairInd) = tmpIdCnt
   'add unique identifications with scores to all ids, also
   'memorize first and last index of id block for current pair
   If tmpIdCnt > 0 Then
      If UBound(mgPInd) < mgCnt + tmpIdCnt Then    'add more room
         'make sure it is enough to accomodate current batch
         ReDim Preserve mgPInd(UBound(mgPInd) + tmpIdCnt + 2000)
         ReDim Preserve mgIDInd(UBound(mgPInd))
         ReDim Preserve mgScore(UBound(mgPInd))
      End If
      PIDInd1(PairInd) = mgCnt            'first index
      'last index will remain -1 if no ids and PIDInd2(i)>=0
      'should always be checked when enumerating ids for pair
      For I = 0 To tmpIdCnt - 1
          mgCnt = mgCnt + 1
          mgPInd(mgCnt - 1) = PairInd
          mgIDInd(mgCnt - 1) = tmpIdInd(I)
          mgScore(mgCnt - 1) = tmpScore(I)
          PIDInd2(PairInd) = mgCnt - 1    'last index
      Next I
   End If
End If
End Sub

Private Sub PickParameters()
Call txtMWTol_LostFocus
Call txtNETTol_LostFocus
Call txtNETFormula_LostFocus
End Sub
