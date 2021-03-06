VERSION 5.00
Begin VB.Form frmUMCIdentification 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Search On LC-MS Feature"
   ClientHeight    =   4980
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4440
   Icon            =   "frmUMCIdentification.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "MW Tolerance"
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   2055
      Begin VB.TextBox txtMWTol 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Text            =   "1"
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton optMWTolType 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Da"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   12
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton optMWTolType 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&ppm"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   11
         Top             =   300
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Isotopic MW Field"
      Height          =   1455
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   2055
      Begin VB.OptionButton optMWField 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&The Most Abundant"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   1020
         Width           =   1815
      End
      Begin VB.OptionButton optMWField 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Monoisotopic"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   660
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optMWField 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Average"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.Frame fraExportDestination 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Export Destination"
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   4215
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse"
         Height          =   375
         Left            =   3240
         TabIndex        =   7
         Top             =   780
         Width           =   855
      End
      Begin VB.TextBox txtExportDestination 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   3975
      End
      Begin VB.OptionButton optExportDestination 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Text file"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optExportDestination 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Organism MT Tag Database"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Exp&ort"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "S&earch"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblMTCount 
      BackStyle       =   0  'Transparent
      Caption         =   "MT Tags count: 0"
      Height          =   495
      Left            =   2280
      TabIndex        =   17
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "mama mia, casa mia, tutti quanti, Etta zia"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4680
      Width           =   4215
   End
End
Attribute VB_Name = "frmUMCIdentification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Search of MT tag database following UMC structures
'This search is done only on MW (elution is not considered)
'and is coded for use in new MT tag DB promotion process
'created: 12/14/2001 nt
'last modified: 02/15/2002 nt
'----------------------------------------------------------
Option Explicit

Const EXPORT_MTDB = 0
Const EXPORT_TEXT = 1

Public CallerID As Long

'names of stored procedures that will write data
'to database tables retrieved from init. file
Dim ExpAnalysisSPName As String             ' Stored procedure AddMatchMaking
''Dim ExpPeakSPName As String                 ' Stored procedure AddFTICRPeak; Unused variable

Dim ExportDestination As Long

Dim fs As New FileSystemObject

Dim UMCStat2() As Double        'UMC Statistics

'search results can be exported to text file or to database
Dim SRCount As Long             'count of peaks for export
Dim SRInd() As Long             'index in data arrays
Dim SRType() As Long            'charge state or isotopic
Dim SRMassTagID() As String     'index of MT tag identification
Dim SRHitsCount() As Long       'number of hits by peak
Dim SRUMCInd() As Long          'class index

Dim mwutSearch As New MWUtil    'fast find object

Private Sub cmdBrowse_Click()
'--------------------------------------------
'displays Save dialog to specify name of file
'--------------------------------------------
Dim NewFName As String
On Error Resume Next
NewFName = FileSaveProc(Me.hwnd, Trim$(txtExportDestination.Text), fstFileSaveTypeConstants.fstTxt)
If Len(NewFName) > 0 Then txtExportDestination.Text = NewFName
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdExport_Click()
Dim lngErrorCode As Long

Me.MousePointer = vbHourglass
Select Case ExportDestination
Case EXPORT_MTDB
    UpdateStatus "Sending results to MT tag DB..."
    
    MsgBox "Database export is no longer supported using this window.", vbInformation + vbOKOnly, "Error"
    
' September 2004: Unsupported code
''    eResponse = MsgBox("Automatically re-generate the match making parameters description before exporting?", vbQuestion + vbYesNoCancel + vbDefaultButton1, "Re-generate parameters")
''
''    If eResponse = vbYes Then
''        ' Update the text in MD_Parameters
''        strUMCSearchMode = FindSettingInAnalysisHistory(CallerID, UMC_SEARCH_MODE_SETTING_TEXT, , True, ":", ";")
''        If Right(strUMCSearchMode, 1) = ")" Then strUMCSearchMode = Left(strUMCSearchMode, Len(strUMCSearchMode) - 1)
''        GelAnalysis(CallerID).MD_Parameters = ConstructAnalysisParametersText(CallerID, strUMCSearchMode, AUTO_SEARCH_UMC_MTDB)
''    End If
''
''    If eResponse <> vbCancel Then
''        MsgBox ExportMTDB(), vbOKOnly, glFGTU
''    Else
''        MsgBox "Export aborted.", vbOKOnly, glFGTU
''    End If
Case EXPORT_TEXT
    UpdateStatus "Sending results to text file..."
    lngErrorCode = ExportText()
    If lngErrorCode = 0 Then
        MsgBox "Results successfully exported.", vbOKOnly, glFGTU
    Else
        MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, vbOKOnly, glFGTU
    End If
End Select
Me.MousePointer = vbDefault
UpdateStatus ""
End Sub

Private Sub cmdSearch_Click()
StartSearch
End Sub

Private Sub Form_Load()
'present itself so that user does not freak out
Me.Visible = True
DoEvents
' MonroeMod
If CallerID >= 1 And CallerID <= UBound(GelBody) Then samtDef = GelSearchDef(CallerID).AMTSearchOnUMCs
'load default search definition(same as for other searches)
With samtDef
    optMWField(.MWField - MW_FIELD_OFFSET).Value = True
    Select Case .TolType
    Case gltPPM
        optMWTolType(0).Value = True
    Case gltABS
        optMWTolType(1).Value = True
    Case Else
        Debug.Assert False
    End Select
    txtMWTol.Text = .MWTol
End With
'load default settings
UpdateStatus ""
txtExportDestination.Text = App.Path & "\" & "UMCIDResults.txt"
If GelAnalysis(CallerID) Is Nothing Then
   optExportDestination(EXPORT_TEXT).Value = True
   optExportDestination(EXPORT_MTDB).Enabled = False
   'cmdSelectMassTags.Enabled = False
   cmdSearch.Enabled = False
   Exit Sub
End If

'retrieve names of stored procedures
ExpAnalysisSPName = glbPreferencesExpanded.MTSConnectionInfo.spPutAnalysis
'ExpPeakSPName = glbPreferencesExpanded.MTSConnectionInfo.spPutPeak

If Len(ExpAnalysisSPName) <= 0 Then
   MsgBox "Names of stored procedures neccessary to export data not found in the initialization file section associated with this analysis. Ask for help, dude.", vbOKOnly, glFGTU
   optExportDestination(EXPORT_TEXT).Value = True
   optExportDestination(EXPORT_MTDB).Enabled = False
End If
If optExportDestination(EXPORT_TEXT).Value Then
   ExportDestination = EXPORT_TEXT
Else
   ExportDestination = EXPORT_MTDB
End If
'load MT tags if neccessary(if not loaded already)
Me.MousePointer = vbHourglass
LoadMTDB
Me.MousePointer = vbDefault
Me.Visible = False
End Sub


Public Function ExportText() As Long
'---------------------------------------------------
'this is simple but long procedure of exporting data
'to text file
'
' Returns 0 if no error, the error number if an error
'---------------------------------------------------
Dim ts As TextStream
Dim sLine As String
Dim fname As String
Dim ShowExported As Boolean
Dim I As Long
Dim strSepChar As String

On Error GoTo err_ExportText

strSepChar = LookupDefaultSeparationCharacter()

fname = Trim$(txtExportDestination.Text)
If Len(fname) > 0 Then
   ShowExported = False
Else        'write and display to temporary file
   fname = GetTempFolder() & RawDataTmpFile
   ShowExported = True
End If
Set ts = fs.OpenTextFile(fname, ForWriting, True)
ts.WriteLine "Generated by: " & GetMyNameVersion() & " on " & Now()
ts.WriteLine "MT tags search results for " & GelBody(CallerID).Caption
sLine = "MMA Tolerance: " & samtDef.MWTol
sLine = sLine & " " & GetSearchToleranceUnitText(CInt(samtDef.TolType))
ts.WriteLine sLine
ts.WriteLine
'now export data
sLine = "Index" & strSepChar & "Type" & strSepChar & "Scan Number" & strSepChar & "Charge State" & strSepChar & "Molecular Mass" & strSepChar & "Abundance" & strSepChar & "" _
        & "Fit" & strSepChar & "ER" & strSepChar & "LockerID" & strSepChar & "Freq.Shift" & strSepChar & "Mass Correction" & strSepChar & "Hits Count" & strSepChar & "" _
        & "MT tag ID" & strSepChar & "UMC Ind" & strSepChar & "UMC First Scan" & strSepChar & "UMC Last Scan" & strSepChar & "" _
        & "UMC Count" & strSepChar & "UMC Abundance" & strSepChar & "UMC Best Fit" & strSepChar & "UMC Avg MW"
ts.WriteLine sLine
'here comes actual export----
With GelData(CallerID)
    For I = 0 To SRCount - 1
        sLine = SRInd(I) & strSepChar & SRType(I) & strSepChar
        Select Case SRType(I)
        Case glCSType
            sLine = sLine & .CSData(SRInd(I)).ScanNumber & strSepChar _
                & .CSData(SRInd(I)).Charge & strSepChar _
                & .CSData(SRInd(I)).AverageMW & strSepChar _
                & .CSData(SRInd(I)).Abundance & strSepChar _
                & .CSData(SRInd(I)).MassStDev & strSepChar
''            If IsNumeric(.CSVar(SRInd(i), csvfMTDDRatio)) Then
''                sLine = sLine & CDbl(.CSVar(SRInd(i), csvfMTDDRatio))
''            End If
            sLine = sLine & strSepChar
            With GelLM(CallerID)
              If .CSCnt > 0 Then
                sLine = sLine & .CSLckID(SRInd(I)) & strSepChar _
                    & .CSFreqShift(SRInd(I)) & strSepChar _
                    & .CSMassCorrection(SRInd(I)) & strSepChar
              Else
                sLine = sLine & strSepChar & strSepChar & strSepChar
              End If
            End With
        Case glIsoType
            sLine = sLine & .IsoData(SRInd(I)).ScanNumber & strSepChar _
                & .IsoData(SRInd(I)).Charge & strSepChar _
                & .IsoData(SRInd(I)).MonoisotopicMW & strSepChar _
                & .IsoData(SRInd(I)).Abundance & strSepChar _
                & .IsoData(SRInd(I)).Fit & strSepChar
''            If IsNumeric(.IsoVar(SRInd(i), isvfMTDDRatio)) Then
''              sLine = sLine & CDbl(.IsoVar(SRInd(i), isvfMTDDRatio))
''            End If
            sLine = sLine & strSepChar
            With GelLM(CallerID)
              If .IsoCnt > 0 Then
                sLine = sLine & .IsoLckID(SRInd(I)) & strSepChar _
                    & .IsoFreqShift(SRInd(I)) & strSepChar _
                    & .IsoMassCorrection(SRInd(I)) & strSepChar
              Else
                sLine = sLine & strSepChar & strSepChar & strSepChar
              End If
            End With
        End Select
        sLine = sLine & SRHitsCount(I) & strSepChar & SRMassTagID(I) & strSepChar _
            & SRUMCInd(I) & strSepChar & UMCStat2(SRUMCInd(I), 2) & strSepChar _
            & UMCStat2(SRUMCInd(I), 3) & strSepChar & UMCStat2(SRUMCInd(I), 8) _
            & strSepChar & UMCStat2(SRUMCInd(I), 4) & strSepChar _
            & UMCStat2(SRUMCInd(I), 6) & strSepChar & UMCStat2(SRUMCInd(I), 1)
        ts.WriteLine sLine
    Next I
End With

ts.Close
Set ts = Nothing
If Len(fname) > 0 Then
    AddToAnalysisHistory CallerID, "Saved search results to disk: " & fname
End If
If ShowExported Then
   frmDataInfo.Tag = "EXP"
   frmDataInfo.Show vbModal
End If
ExportText = 0
Exit Function

err_ExportText:
ExportText = Err.Number
End Function

Public Sub UpdateStatus(ByVal Msg As String)
lblStatus.Caption = Msg
DoEvents
End Sub

Public Function StartSearch(Optional blnShowMessages As Boolean = True) As Long
' Returns the number of hits
Dim Res As Long
Dim I As Long, j As Long, k As Long
Dim CurrInd As Long
Dim CurrType As Long
Dim CurrMW As Double
Dim HitsCount As Long
Dim HitsID() As String
On Error GoTo err_cmdSearch_Click

Me.MousePointer = vbHourglass
UpdateStatus "Preparing fast searching object ..."
Set mwutSearch = New MWUtil
If Not FillMWSearchObject(mwutSearch) Then GoTo err_cmdSearch_Click
'get statistics for unique mass classes
UpdateStatus "Calculating statistic for Unique Mass Classes ..."
Res = UMCStatistics2(CallerID, UMCStat2())
If Res > 0 Then
   'initially reserve space for 20000 identifications
   'add as needed in chunks of 5000
   If InitSRArrays(20000) Then
      With GelUMC(CallerID)
        'do class by class
        For I = 0 To .UMCCnt - 1
            ' MonroeMod
            If I Mod 25 = 0 Then UpdateStatus "Searching peaks from class: " & I & "/" & .UMCCnt
            'and class member by class member
            For j = 0 To .UMCs(I).ClassCount - 1
                CurrInd = .UMCs(I).ClassMInd(j)
                CurrType = .UMCs(I).ClassMType(j)
                Select Case CurrType
                Case glCSType
                     CurrMW = GelData(CallerID).CSData(CurrInd).AverageMW
                Case glIsoType
                     CurrMW = GetIsoMass(GelData(CallerID).IsoData(CurrInd), samtDef.MWField)
                End Select
                HitsCount = GetMT_ID_ForMW(CurrMW, HitsID())
                If HitsCount > 0 Then
                   For k = 0 To HitsCount - 1
                       SRCount = SRCount + 1
                       SRInd(SRCount - 1) = CurrInd
                       SRType(SRCount - 1) = CurrType
                       SRUMCInd(SRCount - 1) = I
                       SRMassTagID(SRCount - 1) = HitsID(k)
                       SRHitsCount(SRCount - 1) = HitsCount
                   Next k
                End If
            Next j
        Next I
        UpdateStatus "Searching peaks from class: " & .UMCCnt & "/" & .UMCCnt
      End With
   Else
      GoTo err_cmdSearch_Click
   End If
Else
   GoTo err_cmdSearch_Click
End If
If blnShowMessages Then MsgBox "Number of MT tags hits: " & SRCount, vbOKOnly, glFGTU

GelSearchDef(CallerID).AMTSearchOnUMCs = samtDef
AddToAnalysisHistory CallerID, GetMassTagSearchSummaryText("Searched LC-MS Features for MT tags (searched ion by ion, only examining ions belonging to LC-MS Features)", SRCount, 0, 0, 0, samtDef, False, False) & " Note: NET Information was not used in the search"

exit_cmdSearch_Click:
TrimSRArrays
UpdateStatus ""
Me.MousePointer = vbDefault
StartSearch = SRCount
Exit Function

err_cmdSearch_Click:
Select Case Err.Number
Case 9      'add more space to search results and continue
    If IncSRArrays(5000) Then
       Resume
    Else
       If blnShowMessages Then MsgBox "Processing aborted, could not store all results. Some results maybe preserved.", vbOKOnly, glFGTU
       Resume exit_cmdSearch_Click
    End If
Case Else
    If blnShowMessages Then MsgBox "Error producing necessary parameters of search. Mission impossible.", vbOKOnly, glFGTU
    Resume exit_cmdSearch_Click
End Select
End Function
'---------------------------------------------------------------
'Results arrays management functions
'---------------------------------------------------------------
Private Function InitSRArrays(ByVal InitSize As Long) As Boolean
On Error GoTo err_InitSRArrays
SRCount = 0
ReDim SRInd(InitSize)
ReDim SRType(InitSize)
ReDim SRMassTagID(InitSize)
ReDim SRUMCInd(InitSize)
ReDim SRHitsCount(InitSize)
InitSRArrays = True
Exit Function

err_InitSRArrays:
DestroySRArrays
End Function

Private Sub DestroySRArrays()
On Error Resume Next
SRCount = 0
Erase SRInd
Erase SRType
Erase SRMassTagID
Erase SRUMCInd
Erase SRHitsCount
End Sub

Private Sub TrimSRArrays()
On Error Resume Next
If SRCount > 0 Then
   ReDim Preserve SRInd(SRCount - 1)
   ReDim Preserve SRType(SRCount - 1)
   ReDim Preserve SRMassTagID(SRCount - 1)
   ReDim Preserve SRUMCInd(SRCount - 1)
   ReDim Preserve SRHitsCount(SRCount - 1)
Else
   DestroySRArrays
End If
End Sub

Private Function IncSRArrays(ByVal IncRate As Long) As Boolean
On Error Resume Next
If IncRate > 0 Then
   ReDim Preserve SRInd(SRCount + IncRate)
   ReDim Preserve SRType(SRCount + IncRate)
   ReDim Preserve SRMassTagID(SRCount + IncRate)
   ReDim Preserve SRUMCInd(SRCount + IncRate)
   ReDim Preserve SRHitsCount(SRCount + IncRate)
End If
IncSRArrays = True
Exit Function

err_IncSRArrays:
'try to recover without loosing previous data
TrimSRArrays
End Function

Private Sub optExportDestination_Click(Index As Integer)
ExportDestination = Index
End Sub

'end results management functions--------------------------------

Private Sub optMWTolType_Click(Index As Integer)
Select Case Index
Case 0
     samtDef.TolType = gltPPM
Case 1
     samtDef.TolType = gltABS
Case Else
     Debug.Assert False
End Select
End Sub

Private Sub txtMWTol_LostFocus()
If IsNumeric(txtMWTol.Text) Then
   samtDef.MWTol = CDbl(txtMWTol.Text)
Else
   MsgBox "MW tolerance should be positive number.", vbOKOnly, glFGTU
   txtMWTol.SetFocus
End If
End Sub

Private Sub optMWField_Click(Index As Integer)
samtDef.MWField = Index + 6
End Sub

Private Sub LoadMTDB(Optional blnForceReload As Boolean = False)
    Dim blnAMTsWereLoaded As Boolean, blnDBConnectionError As Boolean

    If ConfirmMassTagsAndInternalStdsLoaded(Me, CallerID, True, True, False, blnForceReload, 0, blnAMTsWereLoaded, blnDBConnectionError) Then
        lblMTCount.Caption = ConstructMTStatusText(True)
        cmdSearch.Enabled = True
    Else
        If blnDBConnectionError Then
            lblMTCount.Caption = "Error loading MT tags: database connection error."
        Else
            lblMTCount.Caption = "Error loading MT tags: no valid MT tags were found (possibly missing NET values)"
        End If
        cmdSearch.Enabled = True
    End If

End Sub

Private Function GetMT_ID_ForMW(ByVal MW As Double, _
                                ID() As String) As Long
'----------------------------------------------------------
'returns array of MT tags IDs matching MW
'this function is called from frmUMCIdentification
'----------------------------------------------------------
Dim FirstInd As Long
Dim LastInd As Long
Dim AbsTol As Double
Dim I As Long
On Error Resume Next

GetMT_ID_ForMW = 0
Select Case samtDef.TolType
Case gltPPM
    AbsTol = MW * samtDef.MWTol * glPPM
Case gltABS
    AbsTol = samtDef.MWTol
Case Else
    Debug.Assert False
End Select
If mwutSearch.FindIndexRange(MW, AbsTol, FirstInd, LastInd) Then
   If FirstInd <= LastInd And LastInd > 0 Then
      ReDim ID(LastInd - FirstInd)
      For I = FirstInd To LastInd
          ID(I - FirstInd) = AMTData(I).ID
      Next I
      GetMT_ID_ForMW = LastInd - FirstInd + 1
   End If
End If
End Function

