VERSION 5.00
Begin VB.Form frmExportResults 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export Identification Results"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmExportResults.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Destination"
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4455
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse"
         Height          =   315
         Left            =   3120
         TabIndex        =   6
         Top             =   666
         Width           =   1095
      End
      Begin VB.TextBox txtFileName 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   3975
      End
      Begin VB.OptionButton optExportWhere 
         Caption         =   "Text File"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.OptionButton optExportWhere 
         Caption         =   "MTDB -- Not Enabled here; use corresponding Export to MTDB function on search forms"
         Enabled         =   0   'False
         Height          =   400
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   200
         Width           =   3975
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   315
      Left            =   3240
      TabIndex        =   1
      Top             =   2440
      Width           =   1095
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   2440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "This function exports all currently identified distributions (by ion) to a text file"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmExportResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Function exports results of current identification (currently
'identified peaks) to the text file or to the Organism MT tags
'database; format of data is format of FTICR peaks table in db
'---------------------------------------------------------------
'created: 07/06/2001 nt
'last modified: 11/07/2001 nt
'---------------------------------------------------------------
Option Explicit

Const EXPORT_MTDB = 0
Const EXPORT_TEXT = 1

Public CallerID As Long

'names of stored procedures that will write data
'to database tables retrieved from init. file
Dim ExpAnalysisSPName As String             ' Stored procedure AddMatchMaking
'' Dim ExpPeakSPName As String                 ' Stored procedure AddFTICRPeak; September 2004: Unused variable

Dim ExportWhere As Long

Dim fs As New FileSystemObject

Private Sub cmdBrowse_Click()
'--------------------------------------------
'displays Save dialog to specify name of file
'--------------------------------------------
Dim NewFName As String
On Error Resume Next
'NewFName = FileSaveProc(Me.hWnd, Trim$(txtFileName.Text), fstFileSaveTypeConstants.fstTxt)
NewFName = SelectFile(Me.hwnd, "Export to File", "", True, StripFullPath(Trim(txtFileName.Text)), , 2)
If Len(NewFName) > 0 Then txtFileName.Text = NewFName
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdExport_Click()
'-----------------------------------------------------
'do some simple tests and call actual export functions
'-----------------------------------------------------
Me.MousePointer = vbHourglass
Select Case ExportWhere
Case EXPORT_MTDB
'    MsgBox ExportMTDB(), vbOKOnly
Case EXPORT_TEXT
    MsgBox ExportText(), vbOKOnly
End Select
Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
txtFileName.Text = App.Path & "\" & "IDResults.txt"
If GelAnalysis(CallerID) Is Nothing Then
   optExportWhere(EXPORT_TEXT).Value = True
   optExportWhere(EXPORT_MTDB).Enabled = False
   Exit Sub
End If
'retrieve names of stored procedures
ExpAnalysisSPName = glbPreferencesExpanded.MTSConnectionInfo.spPutAnalysis
''ExpPeakSPName = glbPreferencesExpanded.MTSConnectionInfo.spPutPeak

If Len(ExpAnalysisSPName) <= 0 Then         ' Or Len(ExpPeakSPName) <= 0 Then
   MsgBox "Names of stored procedures neccessary to export data not found in the initialization file section associated with this analysis. Ask for help, dude..", vbOKOnly
   optExportWhere(EXPORT_TEXT).Value = True
   optExportWhere(EXPORT_MTDB).Enabled = False
End If
If optExportWhere(EXPORT_TEXT).Value Then
   ExportWhere = EXPORT_TEXT
Else
   ExportWhere = EXPORT_MTDB
End If
End Sub

Private Sub optExportWhere_Click(Index As Integer)
MsgBox "Please use the Export to MT tag database option available on each of the search forms to export data to the database", vbInformation + vbOKOnly, glFGTU
''ExportWhere = Index
End Sub

Public Function ExportText() As String
'---------------------------------------------------
'this is simple but long procedure of exporting data
'to text file
'---------------------------------------------------
Dim ts As TextStream
Dim tmp As String
Dim sLine As String
Dim fname As String
Dim Mass_Tag_ID As Long
Dim AMTRefs() As String
Dim AMTRefsCnt As Long
Dim ShowExported As Boolean
Dim i As Long, j As Long

' MonroeMod: New variable
Dim lngExportCount As Long

On Error GoTo err_ExportText

fname = Trim$(txtFileName.Text)
If Len(fname) > 0 Then
   ShowExported = False
Else        'write and display to temporary file
   fname = GetTempFolder() & RawDataTmpFile
   ShowExported = True
End If
Set ts = fs.OpenTextFile(fname, ForWriting, True)
'now export data
sLine = "Index;Type;Scan;CS;MonoMW_Exp;Abundance;Fit;ER;LockerID;Freq_Shift;Mass_Correction;Hits_Count;Mass_Tag_ID; Mono_MW_DB"
ts.WriteLine sLine
With GelData(CallerID)
  sLine = "Charge State Data Block"
  ts.WriteLine sLine
  For i = 1 To .CSLines
      If Not IsNull(.CSData(i).MTID) Then
         If IsAMTReferenced(.CSData(i).MTID) Then
            AMTRefsCnt = GetAMTRefFromString2(.CSData(i).MTID, AMTRefs())
            If AMTRefsCnt > 0 Then
            'for Charge State standard deviation is used on place of Fit
                tmp = tmp & i & glARG_SEP & glCSType & glARG_SEP & .CSData(i).ScanNumber _
                  & glARG_SEP & .CSData(i).Charge & glARG_SEP & .CSData(i).AverageMW & glARG_SEP _
                  & .CSData(i).Abundance & glARG_SEP & .CSData(i).MassStDev & glARG_SEP
                tmp = tmp & LookupExpressionRatioValue(CallerID, i, False)
                If GelLM(CallerID).CSCnt > 0 Then   'we have mass correction
                   tmp = tmp & glARG_SEP & GelLM(CallerID).CSLckID(i) & glARG_SEP _
                    & GelLM(CallerID).CSFreqShift(i) & glARG_SEP _
                    & GelLM(CallerID).CSMassCorrection(i)
                Else
                   tmp = tmp & glARG_SEP & glARG_SEP & glARG_SEP
                End If
                tmp = tmp & glARG_SEP & AMTRefsCnt
                For j = 1 To AMTRefsCnt         'extract MT tag ID
                    Mass_Tag_ID = CLng(GetIDFromString(AMTRefs(j), AMTMark, AMTIDEnd))
                    If Not Err Then
                       sLine = tmp & glARG_SEP & Mass_Tag_ID
                       ts.WriteLine sLine
                       ' MonroeMod: Counting number of entries exported
                       lngExportCount = lngExportCount + 1
                    End If
                Next j
            End If
         End If
      End If
  Next i
  sLine = "Isotopic Data Block"
  ts.WriteLine sLine
  For i = 1 To .IsoLines
      If Not IsNull(.IsoData(i).MTID) Then
         If IsAMTReferenced(.IsoData(i).MTID) Then
            AMTRefsCnt = GetAMTRefFromString2(.IsoData(i).MTID, AMTRefs())
            If AMTRefsCnt > 0 Then
                tmp = i & glARG_SEP & glIsoType & glARG_SEP & .IsoData(i).ScanNumber _
                  & glARG_SEP & .IsoData(i).Charge & glARG_SEP & .IsoData(i).MonoisotopicMW _
                  & glARG_SEP & .IsoData(i).Abundance & glARG_SEP & .IsoData(i).Fit & glARG_SEP
                tmp = tmp & LookupExpressionRatioValue(CallerID, i, True)
                If GelLM(CallerID).IsoCnt > 0 Then
                   tmp = tmp & glARG_SEP & GelLM(CallerID).IsoLckID(i) & glARG_SEP _
                         & GelLM(CallerID).IsoFreqShift(i) & glARG_SEP _
                         & GelLM(CallerID).IsoMassCorrection(i)
                Else
                   tmp = tmp & glARG_SEP & glARG_SEP & glARG_SEP
                End If
                tmp = tmp & glARG_SEP & AMTRefsCnt
                For j = 1 To AMTRefsCnt         'extract MT tag ID
                    Mass_Tag_ID = CLng(GetIDFromString(AMTRefs(j), AMTMark, AMTIDEnd))
                    If Not Err Then
                       sLine = tmp & glARG_SEP & Mass_Tag_ID
                       ts.WriteLine sLine
                       ' MonroeMod: Counting number of entries exported
                       lngExportCount = lngExportCount + 1
                    End If
                Next j
            End If
         End If
      End If
  Next i
End With
ts.Close
Set ts = Nothing
If ShowExported Then
   frmDataInfo.Tag = "EXP"
   frmDataInfo.Show vbModal
Else
    ' MonroeMod
    AddToAnalysisHistory CallerID, "Exported " & lngExportCount & " search results to text file: " & fname
End If
ExportText = "Results successfully exported."
Exit Function

err_ExportText:
ExportText = "Error: " & Err.Number & vbCrLf & Err.Description
End Function

' May 2003: This function is unused
''Public Function ExportMTDB(Optional ByRef lngErrorNumber As Long, Optional ByRef lngMDID As Long) As String
'''---------------------------------------------------
'''this is simple but long procedure of exporting data
'''results to Organism MT tag database associated
'''lngErrorNumber will contain the error number, if an error occurs
'''---------------------------------------------------
''Dim Mass_Tag_ID As Long
''Dim AMTRefs() As String
''Dim AMTRefsCnt As Long
''Dim i As Long, j As Long
''Dim ExpCnt As Long
''' MonroeMod: New Variable
''Dim lngChargeStateMatchCount As Long
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
''Dim prmState As New ADODB.Parameter
''Dim prmLckID As New ADODB.Parameter
''Dim prmFreqShift As New ADODB.Parameter
''Dim prmMassCorrection As New ADODB.Parameter
''Dim prmMassTagID As New ADODB.Parameter
''Dim prmType As New ADODB.Parameter
''Dim prmResType As New ADODB.Parameter
''Dim prmHitsCount As New ADODB.Parameter
'''Dim prmUMCInd As New ADODB.Parameter                   'this has to be finished if it is going to be used
'''Dim prmUMCFirstScan As New ADODB.Parameter
'''Dim prmUMCLastScan As New ADODB.Parameter
'''Dim prmUMCCount As New ADODB.Parameter
'''Dim prmUMCAbundance As New ADODB.Parameter
'''Dim prmUMCBestFit As New ADODB.Parameter
'''Dim prmUMCAvgMW As New ADODB.Parameter
'''Dim prmPairInd As New ADODB.Parameter
''On Error GoTo err_ExportMTDB
''
''strCaptionSaved = Me.Caption
''
''Me.Caption = "Connecting to the database"
''If Not EstablishConnection(cnNew, GelAnalysis(CallerID).MTDB.cn.ConnectionString, False) Then
''    lngErrorNumber = -1
''    Me.Caption = strCaptionSaved
''    ExportMTDB = "Error: Unable to establish a connection to the database"
''    Exit Function
''End If
''
'''first write new analysis in T_Match_Description table
''AddEntryToMatchMakingDescriptionTable cnNew, lngMDID, ExpAnalysisSPName, CallerID, 0, GelData(CallerID).CustomNETsDefined, False, strIniFileName
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
''
'''now export data
''ExpCnt = 0
''With GelData(CallerID)
'''first charge state and ...
''  For i = 1 To .CSLines
''      If i Mod 100 = 0 Then
''         Me.Caption = "Exporting results to the database: " & Trim(i) & " / " & Trim(.CSLines + .IsoLines)
''         DoEvents
''      End If
''      If Not IsNull(.CSData(i).mtid) Then
''         If IsAMTReferenced(.CSData(i).mtid) Then
''            AMTRefsCnt = GetAMTRefFromString2(.CSData(i).mtid, AMTRefs())
''            If AMTRefsCnt > 0 Then
''                prmFTICRID.value = i
''                prmFTICRType.value = glCSType
''                prmScanNumber.value = .CSData(i).ScanNumber
''                prmChargeState.value = .CSData(i).Charge
''                prmMonoisotopicMass.value = .CSData(i).AverageMW
''                prmAbundance.value = .CSData(i).Abundance
''                prmFit.value = .CSData(i).MassStDev     'standard deviation
''                prmExpressionRatio.value = LookupExpressionRatioValue(CallerID, i, False)
''                If GelLM(CallerID).CSCnt > 0 Then
''                   prmLckID.value = GelLM(CallerID).CSLckID(i)
''                   prmFreqShift.value = GelLM(CallerID).CSFreqShift(i)
''                   prmMassCorrection.value = GelLM(CallerID).CSMassCorrection(i)
''                End If
''                prmHitsCount.value = AMTRefsCnt
''                For j = 1 To AMTRefsCnt         'extract MT tag ID
''                    Mass_Tag_ID = CLng(GetIDFromString(AMTRefs(j), AMTMark, AMTIDEnd))
''                    prmMassTagID.value = Mass_Tag_ID
''                    cmdPutNewPeak.Execute
''                    ExpCnt = ExpCnt + 1
''                Next j
''            End If
''         End If
''      End If
''  Next i
''
''' MonroeMod
''lngChargeStateMatchCount = ExpCnt
''AddToAnalysisHistory CallerID, "Export to Peak Results table details: Charge State Match Count = " & lngChargeStateMatchCount
''
'''... then isotopic
''  For i = 1 To .IsoLines
''      If i Mod 100 = 0 Then
''         Me.Caption = "Exporting results to the database: " & Trim(i + .CSLines) & " / " & Trim(.CSLines + .IsoLines)
''         DoEvents
''      End If
''      If Not IsNull(.IsoData(i).MTID) Then
''         If IsAMTReferenced(.IsoData(i).MTID) Then
''            AMTRefsCnt = GetAMTRefFromString2(.IsoData(i).MTID, AMTRefs())
''            If AMTRefsCnt > 0 Then
''                prmFTICRID.value = i
''                prmFTICRType.value = glIsoType
''                prmScanNumber.value = .IsoData(i).ScanNumber
''                prmChargeState.value = .IsoData(i).Charge
''                prmMonoisotopicMass.value = .IsoData(i).MonoisotopicMW
''                prmAbundance.value =.IsoData(i).Abundance
''                prmFit.value = .IsoData(i).Fit
''                prmExpressionRatio.value = LookupExpressionRatioValue(CallerID, i, True)
''                If GelLM(CallerID).IsoCnt > 0 Then
''                   prmLckID.value = GelLM(CallerID).IsoLckID(i)
''                   prmFreqShift.value = GelLM(CallerID).IsoFreqShift(i)
''                   prmMassCorrection.value = GelLM(CallerID).IsoMassCorrection(i)
''                End If
''                prmHitsCount.value = AMTRefsCnt
''                For j = 1 To AMTRefsCnt         'extract MT tag ID
''                    Mass_Tag_ID = CLng(GetIDFromString(AMTRefs(j), AMTMark, AMTIDEnd))
''                    prmMassTagID.value = Mass_Tag_ID
''                    cmdPutNewPeak.Execute
''                    ExpCnt = ExpCnt + 1
''                Next j
''            End If
''         End If
''      End If
''  Next i
''
''End With
''
''' MonroeMod
''AddToAnalysisHistory CallerID, "Export to Peak Results table details: Isotopic Peak Match Count = " & ExpCnt - lngChargeStateMatchCount
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
