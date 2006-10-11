VERSION 5.00
Begin VB.Form frmUMCIdentification 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Search On UMC"
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
Dim i As Long
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
    For i = 0 To SRCount - 1
        sLine = SRInd(i) & strSepChar & SRType(i) & strSepChar
        Select Case SRType(i)
        Case glCSType
            sLine = sLine & .CSData(SRInd(i)).ScanNumber & strSepChar _
                & .CSData(SRInd(i)).Charge & strSepChar _
                & .CSData(SRInd(i)).AverageMW & strSepChar _
                & .CSData(SRInd(i)).Abundance & strSepChar _
                & .CSData(SRInd(i)).MassStDev & strSepChar
''            If IsNumeric(.CSVar(SRInd(i), csvfMTDDRatio)) Then
''                sLine = sLine & CDbl(.CSVar(SRInd(i), csvfMTDDRatio))
''            End If
            sLine = sLine & strSepChar
            With GelLM(CallerID)
              If .CSCnt > 0 Then
                sLine = sLine & .CSLckID(SRInd(i)) & strSepChar _
                    & .CSFreqShift(SRInd(i)) & strSepChar _
                    & .CSMassCorrection(SRInd(i)) & strSepChar
              Else
                sLine = sLine & strSepChar & strSepChar & strSepChar
              End If
            End With
        Case glIsoType
            sLine = sLine & .IsoData(SRInd(i)).ScanNumber & strSepChar _
                & .IsoData(SRInd(i)).Charge & strSepChar _
                & .IsoData(SRInd(i)).MonoisotopicMW & strSepChar _
                & .IsoData(SRInd(i)).Abundance & strSepChar _
                & .IsoData(SRInd(i)).Fit & strSepChar
''            If IsNumeric(.IsoVar(SRInd(i), isvfMTDDRatio)) Then
''              sLine = sLine & CDbl(.IsoVar(SRInd(i), isvfMTDDRatio))
''            End If
            sLine = sLine & strSepChar
            With GelLM(CallerID)
              If .IsoCnt > 0 Then
                sLine = sLine & .IsoLckID(SRInd(i)) & strSepChar _
                    & .IsoFreqShift(SRInd(i)) & strSepChar _
                    & .IsoMassCorrection(SRInd(i)) & strSepChar
              Else
                sLine = sLine & strSepChar & strSepChar & strSepChar
              End If
            End With
        End Select
        sLine = sLine & SRHitsCount(i) & strSepChar & SRMassTagID(i) & strSepChar _
            & SRUMCInd(i) & strSepChar & UMCStat2(SRUMCInd(i), 2) & strSepChar _
            & UMCStat2(SRUMCInd(i), 3) & strSepChar & UMCStat2(SRUMCInd(i), 8) _
            & strSepChar & UMCStat2(SRUMCInd(i), 4) & strSepChar _
            & UMCStat2(SRUMCInd(i), 6) & strSepChar & UMCStat2(SRUMCInd(i), 1)
        ts.WriteLine sLine
    Next i
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

' September 2004: Unused Function
''Public Function ExportMTDB(Optional ByRef lngErrorNumber As Long, Optional ByRef lngMDID As Long) As String
'''---------------------------------------------------
'''this is simple but long procedure of exporting data
'''results to Organism MT tag database associated
'''lngErrorNumber will contain the error number, if an error occurs
'''---------------------------------------------------
''Dim i As Long
''Dim ExpCnt As Long
''Dim strCaptionSaved As String
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
''On Error GoTo err_ExportMTDB
''
''strCaptionSaved = Me.Caption
''
''Me.Caption = "Connecting to database"
''If Not EstablishConnection(cnNew, GelAnalysis(CallerID).MTDB.cn.ConnectionString, False) Then
''    Debug.Assert False
''    lngErrorNumber = -1
''    ExportMTDB = "Error: Unable to establish a connection to the database"
''    Exit Function
''End If
''
'''first write new analysis in T_Match_Making_Description table
''AddEntryToMatchMakingDescriptionTable cnNew, lngMDID, ExpAnalysisSPName, CallerID, SRCount, GelData(CallerID).CustomNETsDefined, False, strIniFileName
''
''' MonroeMod
''AddToAnalysisHistory CallerID, "Exported UMC Identification results to database (" & ExtractDBNameFromConnectionString(GelAnalysis(CallerID).MTDB.cn.ConnectionString) & "); MMD_ID = " & lngMDID
''AddToAnalysisHistory CallerID, "Export to MMD table details: Reference Job = " & GelAnalysis(CallerID).MD_Reference_Job & "; MD_File = " & GelAnalysis(CallerID).MD_file
''
'''nothing to export
''If SRCount <= 0 Then Exit Function
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
''Set prmResType = cmdPutNewPeak.CreateParameter("Type", adInteger, adParamInput, , FPR_Type_Standard)
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
''Set prmPairInd = cmdPutNewPeak.CreateParameter("PairInd", adInteger, adParamInput, , -1)
''cmdPutNewPeak.Parameters.Append prmPairInd
''
'''now export data
''ExpCnt = 0
''With GelData(CallerID)
''    For i = 0 To SRCount - 1
''        If i Mod 25 = 0 Then
''            Me.Caption = "Exporting to DB: " & Trim(i) & " / " & Trim(SRCount)
''            DoEvents
''        End If
''
''        prmFTICRID.value = SRInd(i)
''        prmFTICRType.value = SRType(i)
''        Select Case SRType(i)
''        Case glCSType
''            prmScanNumber.value = .CSData(SRInd(i)).ScanNumber
''            prmChargeState.value = .CSData(SRInd(i)).Charge
''            prmMonoisotopicMass.value = .CSData(SRInd(i)).AverageMW
''            prmAbundance.value = .CSData(SRInd(i)).Abundance
''            prmFit.value = .CSData(SRInd(i)).MassStDev     'standard deviation
''            prmExpressionRatio.value = LookupExpressionRatioValue(CallerID, SRInd(i), False)
''            With GelLM(CallerID)
''              If .CSCnt > 0 Then
''                prmLckID.value = .CSLckID(SRInd(i))
''                prmFreqShift.value = .CSFreqShift(SRInd(i))
''                prmMassCorrection.value = .CSMassCorrection(SRInd(i))
''              End If
''            End With
''        Case glIsoType
''            prmScanNumber.value = .IsoData(SRInd(i)).ScanNumber
''            prmChargeState.value = .IsoData(SRInd(i)).Charge
''            prmMonoisotopicMass.value = .IsoData(SRInd(i)).MonoisotopicMW
''            prmAbundance.value = .IsoData(SRInd(i)).Abundance
''            prmFit.value = .IsoData(SRInd(i)).Fit
''            prmExpressionRatio.value = LookupExpressionRatioValue(CallerID, SRInd(i), True)
''            With GelLM(CallerID)
''              If .IsoCnt > 0 Then
''                prmLckID.value = .IsoLckID(SRInd(i))
''                prmFreqShift.value = .IsoFreqShift(SRInd(i))
''                prmMassCorrection.value = .IsoMassCorrection(SRInd(i))
''              End If
''            End With
''        End Select
''        prmHitsCount.value = SRHitsCount(i)
''        prmMassTagID.value = SRMassTagID(i)
''        prmUMCInd.value = SRUMCInd(i)
''        prmUMCFirstScan.value = UMCStat2(SRUMCInd(i), 2)
''        prmUMCLastScan.value = UMCStat2(SRUMCInd(i), 3)
''        prmUMCCount.value = UMCStat2(SRUMCInd(i), 8)
''        prmUMCAbundance.value = UMCStat2(SRUMCInd(i), 4)
''        prmUMCBestFit.value = UMCStat2(SRUMCInd(i), 6)
''        prmUMCAvgMW.value = UMCStat2(SRUMCInd(i), 1)
''
''        cmdPutNewPeak.Execute
''        ExpCnt = ExpCnt + 1
''    Next i
''End With
''
''' MonroeMod
''AddToAnalysisHistory CallerID, "Export to Peak Results table details: UMC Peaks Match Count = " & ExpCnt
''
''ExportMTDB = ExpCnt & " associations between MT tags and FTICR peaks exported."
''Set cmdPutNewPeak.ActiveConnection = Nothing
''cnNew.Close
''Me.Caption = strCaptionSaved
''lngErrorNumber = 0
''Exit Function
''
''err_ExportMTDB:
''ExportMTDB = "Error: " & Err.Number & vbCrLf & Err.Description
''lngErrorNumber = Err.Number
''If Not cnNew Is Nothing Then cnNew.Close
''Me.Caption = strCaptionSaved
''End Function


Public Sub UpdateStatus(ByVal Msg As String)
lblStatus.Caption = Msg
DoEvents
End Sub

Public Function StartSearch(Optional blnShowMessages As Boolean = True) As Long
' Returns the number of hits
Dim Res As Long
Dim i As Long, j As Long, k As Long
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
        For i = 0 To .UMCCnt - 1
            ' MonroeMod
            If i Mod 25 = 0 Then UpdateStatus "Searching peaks from class: " & i & "/" & .UMCCnt
            'and class member by class member
            For j = 0 To .UMCs(i).ClassCount - 1
                CurrInd = .UMCs(i).ClassMInd(j)
                CurrType = .UMCs(i).ClassMType(j)
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
                       SRUMCInd(SRCount - 1) = i
                       SRMassTagID(SRCount - 1) = HitsID(k)
                       SRHitsCount(SRCount - 1) = HitsCount
                   Next k
                End If
            Next j
        Next i
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
AddToAnalysisHistory CallerID, GetMassTagSearchSummaryText("Searched UMC's for MT tags (searched ion by ion, only examining ions belonging to UMC's)", SRCount, 0, 0, 0, samtDef, False, False) & " Note: NET Information was not used in the search"

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

    If ConfirmMassTagsAndInternalStdsLoaded(Me, CallerID, True, 0, blnForceReload, True, blnAMTsWereLoaded, blnDBConnectionError) Then
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
Dim i As Long
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
      For i = FirstInd To LastInd
          ID(i - FirstInd) = AMTData(i).ID
      Next i
      GetMT_ID_ForMW = LastInd - FirstInd + 1
   End If
End If
End Function

