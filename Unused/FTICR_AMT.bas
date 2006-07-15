Attribute VB_Name = "Module15"
'FTICR_AMT database procedures
'last modified 08/14/2000 nt
'-------------------------------------------------------------
Option Explicit

Public Const glFTICR_AMT_TBL_FTICR_AMT = "FTICR_AMT"
Public Const glFTICR_AMT_TBL_FTSOURCES = "FTSources"

Public Function ConnectToFTICR_AMT(dbFTICR_AMT As Database, _
                                   ByVal sDBName As String, _
                                   Optional AskUser As Boolean) As Boolean
Static UserAlreadyAsked As Boolean
Dim eResponse As VbMsgBoxResult
Dim TblCnt As Long
On Error GoTo err_ConnectToFTICR_AMT

If Not IsMissing(AskUser) Then UserAlreadyAsked = Not AskUser

If Not UserAlreadyAsked Then
   eResponse = MsgBox("2DGel will access FTICR_AMT database " & sDBName _
             & ". Multiple versions of the that database might exist " _
             & "so please make sure that listed file is The FTICR_AMT " _
             & "database. To specify different database open Options dialog.", vbOKCancel)
   If eResponse <> vbOK Then Exit Function
   UserAlreadyAsked = True
End If
Set dbFTICR_AMT = DBEngine.Workspaces(0).OpenDatabase(sDBName, True, False)
'check for two essential tables
On Error Resume Next
TblCnt = dbFTICR_AMT.TableDefs(glFTICR_AMT_TBL_FTICR_AMT).RecordCount
If Err Then
   MsgBox "Error accessing FTICR_AMT table. Connection with FTICR_AMT database will be closed.", vbOKOnly
   GoTo err_ConnectToFTICR_AMT
End If
TblCnt = dbFTICR_AMT.TableDefs(glFTICR_AMT_TBL_FTSOURCES).RecordCount
If Err Then
   MsgBox "Error accessing FTSources table. Connection with FTICR_AMT database will be closed.", vbOKOnly
   GoTo err_ConnectToFTICR_AMT
End If
ConnectToFTICR_AMT = True
Exit Function

err_ConnectToFTICR_AMT:
Set dbFTICR_AMT = Nothing
ConnectToFTICR_AMT = False
End Function

Public Function F_ACheckTheSource(ByRef dbFTICR_AMT As Database, _
                                  ByVal GelName As String) As Long
'returns FTSFileID field value from the FTSources table if
'gel already in database or -1 if not
Dim SourceSQL As String
Dim rsSource As Recordset
On Error GoTo err_F_ACheckTheSource

SourceSQL = "SELECT * FROM [" & glFTICR_AMT_TBL_FTSOURCES & "] WHERE [" _
            & glFTICR_AMT_TBL_FTSOURCES & "].FTSFileName = '" & GelName & "';"
Set rsSource = dbFTICR_AMT.OpenRecordset(SourceSQL, dbOpenSnapshot)
rsSource.MoveLast      'this will trigger 3021 error if not found
F_ACheckTheSource = rsSource.Fields("FTSFileID").value

exit_CheckTheSource:
Set rsSource = Nothing
Exit Function

err_F_ACheckTheSource:
F_ACheckTheSource = -1
GoTo exit_CheckTheSource
End Function


Public Function F_AAddSource(ByRef dbFTICR_AMT As Database, _
                             ByVal GelInd As Long) As Long
'returns FTSFileID for new source if successful, -1 on error
Dim rsFTS As Recordset
On Error GoTo exit_F_AAddSource
Set rsFTS = dbFTICR_AMT.OpenRecordset(glFTICR_AMT_TBL_FTSOURCES, dbOpenTable)
With rsFTS
    .AddNew
    .Fields("FTSFileName").value = GelBody(GelInd).Caption
    .Fields("FTSMS_MSSearch").value = Now()
    .Fields("FTSFirstFN").value = GelData(GelInd).ScanInfo(1).ScanNumber
    .Fields("FTSLastFN").value = GelData(GelInd).ScanInfo(UBound(GelData(GelInd).ScanInfo)).ScanNumber
    .Fields("FTSComment").value = GelData(GelInd).Comment
    .Update
    .Bookmark = .LastModified   'make new record current(neccessary!!!)
    F_AAddSource = .Fields("FTSFileID").value
End With

exit_F_AAddSource:
If Err Then
   MsgBox "Error: " & Err.Number & " - " & Err.Description, vbOKOnly, "F_AAddSource"
   F_AAddSource = -1
End If
Set rsFTS = Nothing
End Function

Public Function F_AUpdateSource(ByRef dbFTICR_AMT As Database, _
                                ByVal FTSId As Integer, _
                                ByVal GelInd As Long) As Boolean
Dim rsFTS As Recordset
On Error GoTo exit_F_AUpdateSource
Set rsFTS = dbFTICR_AMT.OpenRecordset(glFTICR_AMT_TBL_FTSOURCES, dbOpenTable)
With rsFTS
    .Index = "PrimaryKey"
    .Seek "=", FTSId
    If .NoMatch Then
       F_AUpdateSource = False
       GoTo exit_F_AUpdateSource
    End If
    .Edit
    .Fields("FTSMS_MSSearch").value = Now()
    .Fields("FTSFirstFN").value = GelData(GelInd).ScanInfo(1).ScanNumber
    .Fields("FTSLastFN").value = GelData(GelInd).ScanInfo(UBound(GelData(GelInd).ScanInfo)).ScanNumber
    .Fields("FTSComment").value = GelData(GelInd).Comment
    .Update
End With
F_AUpdateSource = True

exit_F_AUpdateSource:
If Err Then
   MsgBox "Error: " & Err.Number & " - " & Err.Description, vbOKOnly, "F_AUpdateSource"
   F_AUpdateSource = False
End If
Set rsFTS = Nothing
End Function

Public Function F_ADeleteSourceRecords(ByRef dbFTICR_AMT As Database, _
                                       ByVal FTSId As Integer) As Boolean
Dim DelSQL As String
DelSQL = "DELETE * FROM [" & glFTICR_AMT_TBL_FTICR_AMT & "] WHERE [" _
         & glFTICR_AMT_TBL_FTICR_AMT & "].F_AFTSID = " & FTSId & ";"
dbFTICR_AMT.Execute DelSQL
F_ADeleteSourceRecords = True

exit_F_AUpdateSource:
If Err Then
   MsgBox "Error: " & Err.Number & " - " & Err.Description, vbOKOnly, "F_AUpdateSource"
   F_ADeleteSourceRecords = False
End If
End Function

' Unused Function (May 2003)
'''Public Function F_AGetSourceDesc(ByRef dbFTICR_AMT As Database, _
'''                                 ByVal FTSId As Integer) As String
''''returns record (formatted) of the specified source, zero-
''''length string if source was not found or error occured
'''Dim rsFTS As Recordset
'''Dim TmpDesc As String
'''On Error GoTo err_F_AGetSourceDesc
'''Set rsFTS = dbFTICR_AMT.OpenRecordset(glFTICR_AMT_TBL_FTSOURCES, dbOpenTable)
'''With rsFTS
'''    .Index = "PrimaryKey"
'''    .Seek "=", FTSId
'''    If .NoMatch Then GoTo exit_F_AGetSourceDesc
'''    TmpDesc = "Last MS/MS search: " & .Fields("FTSMS_MSSearch").value & vbCrLf
'''    TmpDesc = TmpDesc & .Fields("FTSComment").value
'''End With
'''F_AGetSourceDesc = TmpDesc
'''
'''exit_F_AGetSourceDesc:
'''Set rsFTS = Nothing
'''Exit Function
'''
'''err_F_AGetSourceDesc:
'''MsgBox "Error: " & Err.Number & " - " & Err.Description, vbOKOnly, "F_AUpdateSource"
'''F_AGetSourceDesc = ""
'''GoTo exit_F_AGetSourceDesc
'''End Function

Public Function F_AOpenFTICR_AMTTbl(dbFTICR_AMT As Database) As Recordset
On Error Resume Next
Set F_AOpenFTICR_AMTTbl = dbFTICR_AMT.OpenRecordset(glFTICR_AMT_TBL_FTICR_AMT, dbOpenTable)
If Err Then
   MsgBox "Error: " & Err.Number & " - " & Err.Description, vbOKOnly, "F_AOpenFTICR_AMTTbl"
   Set F_AOpenFTICR_AMTTbl = Nothing
End If
End Function

Public Sub F_AExportIsoAMTHits(ByVal Ind As Long, _
                               ByVal Scope As Integer)
'exports Isotopic data from GelData(Ind) matched
'with AMT databases to the FTICR_AMT database (monoisotopic masses)
Dim dbFTICR_AMT As Database
Dim rsFTICR_AMT As Recordset
Dim CallerFTSID As Long
Dim i As Long, j As Long
Dim AMTCnt As Long
Dim AMTs() As String
Dim RecCnt As Long
Dim eResponse As VbMsgBoxResult
On Error Resume Next

If Len(sFTICR_AMTPath) <= 0 Then
   MsgBox "Path to FTICR_AMT database not specified. Open Options dialog and select path to the database.", vbOKOnly
   Exit Sub
End If

If Not ConnectToFTICR_AMT(dbFTICR_AMT, sFTICR_AMTPath, True) Then
   MsgBox "Error connecting to FTICR_AMT database. It might be in use by another user, or removed from the path specified on Options dialog.", vbOKOnly
   Exit Sub
End If

CallerFTSID = F_ACheckTheSource(dbFTICR_AMT, GelBody(Ind).Caption)
If CallerFTSID < 0 Then             'new source
   CallerFTSID = F_AAddSource(dbFTICR_AMT, Ind)
   If CallerFTSID < 0 Then       'error while adding new source
      MsgBox "Error trying to add a source to the FTICR_AMT database. " _
         & "Data from gel can not be saved in the database. ", vbOKOnly
      Exit Sub
   End If
Else                             'source already exists
   eResponse = MsgBox("Data from current gel already found in FTICR_AMT database." _
      & " Delete found FTICR_AMT records related with current gel? If you choose Yes information " _
      & " could be lost because those records could come from MS/MS analysis)?", vbYesNoCancel)
   Select Case eResponse
   Case vbYes
       'delete all records in FTICR_AMT table related to CallerFTSID
       F_ADeleteSourceRecords dbFTICR_AMT, CallerFTSID
   Case vbNo        'do not delete but append new records
   Case vbCancel    'cancel the whole procedure
       Exit Sub
   End Select
   'update record in FTSources table
   F_AUpdateSource dbFTICR_AMT, CallerFTSID, Ind
End If

Set rsFTICR_AMT = F_AOpenFTICR_AMTTbl(dbFTICR_AMT)
If Not rsFTICR_AMT Is Nothing Then
  RecCnt = 0
  With GelData(Ind)
    If .IsoLines > 0 Then
       For i = 1 To .IsoLines
         Select Case Scope
         Case glScope.glSc_All
           AMTCnt = 0
           AMTCnt = GetAMTRefFromString2(.IsoData(i).MTID, AMTs())
           If AMTCnt > 0 Then
              For j = 1 To AMTCnt
                rsFTICR_AMT.AddNew
                rsFTICR_AMT.Fields("F_AFTSID").value = CallerFTSID   'caller gel source ID
                rsFTICR_AMT.Fields("F_AMTID").value = GetIDFromString(AMTs(j), AMTMark, AMTIDEnd)
                rsFTICR_AMT.Fields("F_AMW").value = .IsoData(i).MonoisotopicMW            'gel monoisotopic MW
                rsFTICR_AMT.Fields("F_AFN").value = .IsoData(i).ScanNumber            'scan number
                rsFTICR_AMT.Fields("F_AInt").value = .IsoData(i).Abundance          'intensity
                rsFTICR_AMT.Fields("F_AIndex").value = i                     'index in Isotopic data array
                rsFTICR_AMT.Fields("F_AMS_MSData").value = "NA"              'results from MS_MS search
                If GelDraw(Ind).IsoER(i) >= 0 Then
                   rsFTICR_AMT.Fields("F_AER").value = GelDraw(Ind).IsoER(i)    'expression ratio
                Else
                   rsFTICR_AMT.Fields("F_AER").value = Null
                End If
                rsFTICR_AMT.Update
                RecCnt = RecCnt + 1
              Next j
           End If
         Case glScope.glSc_Current
           If GelDraw(Ind).IsoID(i) > 0 And GelDraw(Ind).IsoR(i) > 0 Then
             AMTCnt = 0
             AMTCnt = GetAMTRefFromString2(.IsoData(i).MTID, AMTs())
             If AMTCnt > 0 Then
                For j = 1 To AMTCnt
                  rsFTICR_AMT.AddNew
                  rsFTICR_AMT.Fields("F_AFTSID").value = CallerFTSID
                  rsFTICR_AMT.Fields("F_AMTID").value = GetIDFromString(AMTs(j), AMTMark, AMTIDEnd)
                  rsFTICR_AMT.Fields("F_AMW").value = .IsoData(i).MonoisotopicMW
                  rsFTICR_AMT.Fields("F_AFN").value = .IsoData(i).ScanNumber
                  rsFTICR_AMT.Fields("F_AInt").value = .IsoData(i).Abundance  'intensity
                  rsFTICR_AMT.Fields("F_AIndex").value = i
                  rsFTICR_AMT.Fields("F_AMS_MSData").value = "NA"
                  If GelDraw(Ind).IsoER(i) >= 0 Then
                     rsFTICR_AMT.Fields("F_AER").value = GelDraw(Ind).IsoER(i)    'expression ratio
                  Else
                     rsFTICR_AMT.Fields("F_AER").value = Null
                  End If
                  rsFTICR_AMT.Update
                  RecCnt = RecCnt + 1
                Next j
             End If
           End If
         End Select
       Next i
    End If
  End With
  MsgBox RecCnt & " records appended to the FTICR_AMT database(table).", vbOKOnly
  Set rsFTICR_AMT = Nothing
  Set dbFTICR_AMT = Nothing
Else
  MsgBox "Operation failed( couldn't access FTICR_AMT table).", vbOKOnly
  Set dbFTICR_AMT = Nothing
End If
End Sub

Public Sub F_AGetSourceScan(ByRef dbFTICR_AMT As Database, _
                            ByVal FTSId As Integer, _
                            ByRef FNMin As Integer, _
                            ByRef FNMax As Integer)
'returns scan range of the specified source, or FNMin =1, FNMax=-1
Dim rsFTS As Recordset
On Error GoTo err_F_AGetSourceScan
FNMin = 1
FNMax = -1
Set rsFTS = dbFTICR_AMT.OpenRecordset(glFTICR_AMT_TBL_FTSOURCES, dbOpenTable)
With rsFTS
    .Index = "PrimaryKey"
    .Seek "=", FTSId
    If .NoMatch Then GoTo exit_F_AGetSourceScan
    FNMin = .Fields("FTSFirstFN").value
    FNMax = .Fields("FTSLastFN").value
End With

exit_F_AGetSourceScan:
Set rsFTS = Nothing
Exit Sub

err_F_AGetSourceScan:
MsgBox "Error: " & Err.Number & " - " & Err.Description, vbOKOnly, "F_AGetSourceScan"
GoTo exit_F_AGetSourceScan
End Sub


Public Sub F_AExport(ByVal Ind As Long, _
                     ByVal Scope As Integer)
'exports data from GelData(Ind) to the FTICR_AMT database (monoisotopic masses)
Dim dbFTICR_AMT As Database
Dim rsFTICR_AMT As Recordset
Dim CallerFTSID As Long
Dim i As Long, j As Long
Dim AMTCnt As Long
Dim AMTs() As String
Dim RecCnt As Long
Dim eResponse As VbMsgBoxResult
On Error Resume Next

If Len(sFTICR_AMTPath) <= 0 Then
   MsgBox "Path to FTICR_AMT database not specified. Open Options dialog and select path to the database.", vbOKOnly
   Exit Sub
End If

If Not ConnectToFTICR_AMT(dbFTICR_AMT, sFTICR_AMTPath, True) Then
   MsgBox "Error connecting to FTICR_AMT database. It might be in use by another user, or removed from the path specified on Options dialog.", vbOKOnly
   Exit Sub
End If

CallerFTSID = F_ACheckTheSource(dbFTICR_AMT, GelBody(Ind).Caption)
If CallerFTSID < 0 Then             'new source
   CallerFTSID = F_AAddSource(dbFTICR_AMT, Ind)
   If CallerFTSID < 0 Then       'error while adding new source
      MsgBox "Error trying to add a source to the FTICR_AMT database. " _
         & "Data from gel can not be saved in the database. ", vbOKOnly
      Exit Sub
   End If
Else                             'source already exists
   eResponse = MsgBox("Data from current gel already found in FTICR_AMT database." _
      & " Delete found FTICR_AMT records related with current gel? If you choose Yes information " _
      & " could be lost because those records could come from MS/MS analysis)?", vbYesNoCancel)
   Select Case eResponse
   Case vbYes
       'delete all records in FTICR_AMT table related to CallerFTSID
       F_ADeleteSourceRecords dbFTICR_AMT, CallerFTSID
   Case vbNo        'do not delete but append new records
   Case vbCancel    'cancel the whole procedure
       Exit Sub
   End Select
   'update record in FTSources table
   F_AUpdateSource dbFTICR_AMT, CallerFTSID, Ind
End If

Set rsFTICR_AMT = F_AOpenFTICR_AMTTbl(dbFTICR_AMT)
If Not rsFTICR_AMT Is Nothing Then
  RecCnt = 0
  With GelData(Ind)
    If .CSLines > 0 Then
    End If
    If .IsoLines > 0 Then
       For i = 1 To .IsoLines
         Select Case Scope
         Case glScope.glSc_All
           AMTCnt = 0
           AMTCnt = GetAMTRefFromString2(.IsoData(i).MTID, AMTs())
           If AMTCnt > 0 Then
              For j = 1 To AMTCnt
                rsFTICR_AMT.AddNew
                rsFTICR_AMT.Fields("F_AFTSID").value = CallerFTSID   'caller gel source ID
                rsFTICR_AMT.Fields("F_AMTID").value = GetIDFromString(AMTs(j), AMTMark, AMTIDEnd)
                rsFTICR_AMT.Fields("F_AMW").value = .IsoData(i).MonoisotopicMW            'gel monoisotopic MW
                rsFTICR_AMT.Fields("F_AFN").value = .IsoData(i).ScanNumber            'scan number
                rsFTICR_AMT.Fields("F_AInt").value = .IsoData(i).Abundance          'intensity
                rsFTICR_AMT.Fields("F_AIndex").value = i                     'index in Isotopic data array
                rsFTICR_AMT.Fields("F_AMS_MSData").value = "NA"              'results from MS_MS search
                If GelDraw(Ind).IsoER(i) >= 0 Then
                   rsFTICR_AMT.Fields("F_AER").value = GelDraw(Ind).IsoER(i)    'expression ratio
                Else
                   rsFTICR_AMT.Fields("F_AER").value = Null
                End If
                rsFTICR_AMT.Update
                RecCnt = RecCnt + 1
              Next j
           Else
              rsFTICR_AMT.AddNew
              rsFTICR_AMT.Fields("F_AFTSID").value = CallerFTSID   'caller gel source ID
              rsFTICR_AMT.Fields("F_AMTID").value = Null
              rsFTICR_AMT.Fields("F_AMW").value = .IsoData(i).MonoisotopicMW            'gel monoisotopic MW
              rsFTICR_AMT.Fields("F_AFN").value = .IsoData(i).ScanNumber            'scan number
              rsFTICR_AMT.Fields("F_AInt").value = .IsoData(i).Abundance          'intensity
              rsFTICR_AMT.Fields("F_AIndex").value = i                     'index in Isotopic data array
              rsFTICR_AMT.Fields("F_AMS_MSData").value = "NA"              'results from MS_MS search
              If GelDraw(Ind).IsoER(i) >= 0 Then
                 rsFTICR_AMT.Fields("F_AER").value = GelDraw(Ind).IsoER(i)    'expression ratio
              Else
                 rsFTICR_AMT.Fields("F_AER").value = Null
              End If
              rsFTICR_AMT.Update
              RecCnt = RecCnt + 1
           End If
         Case glScope.glSc_Current
           If GelDraw(Ind).IsoID(i) > 0 And GelDraw(Ind).IsoR(i) > 0 Then
             AMTCnt = 0
             AMTCnt = GetAMTRefFromString2(.IsoData(i).MTID, AMTs())
             If AMTCnt > 0 Then
                For j = 1 To AMTCnt
                  rsFTICR_AMT.AddNew
                  rsFTICR_AMT.Fields("F_AFTSID").value = CallerFTSID
                  rsFTICR_AMT.Fields("F_AMTID").value = GetIDFromString(AMTs(j), AMTMark, AMTIDEnd)
                  rsFTICR_AMT.Fields("F_AMW").value = .IsoData(i).MonoisotopicMW
                  rsFTICR_AMT.Fields("F_AFN").value = .IsoData(i).ScanNumber
                  rsFTICR_AMT.Fields("F_AInt").value = .IsoData(i).Abundance  'intensity
                  rsFTICR_AMT.Fields("F_AIndex").value = i
                  rsFTICR_AMT.Fields("F_AMS_MSData").value = "NA"
                  If GelDraw(Ind).IsoER(i) >= 0 Then
                     rsFTICR_AMT.Fields("F_AER").value = GelDraw(Ind).IsoER(i)    'expression ratio
                  Else
                     rsFTICR_AMT.Fields("F_AER").value = Null
                  End If
                  rsFTICR_AMT.Update
                  RecCnt = RecCnt + 1
                Next j
             Else
                rsFTICR_AMT.AddNew
                rsFTICR_AMT.Fields("F_AFTSID").value = CallerFTSID
                rsFTICR_AMT.Fields("F_AMTID").value = Null
                rsFTICR_AMT.Fields("F_AMW").value = .IsoData(i).MonoisotopicMW
                rsFTICR_AMT.Fields("F_AFN").value = .IsoData(i).ScanNumber
                rsFTICR_AMT.Fields("F_AInt").value = .IsoData(i).Abundance  'intensity
                rsFTICR_AMT.Fields("F_AIndex").value = i
                rsFTICR_AMT.Fields("F_AMS_MSData").value = "NA"
                If GelDraw(Ind).IsoER(i) >= 0 Then
                   rsFTICR_AMT.Fields("F_AER").value = GelDraw(Ind).IsoER(i)    'expression ratio
                Else
                   rsFTICR_AMT.Fields("F_AER").value = Null
                End If
                rsFTICR_AMT.Update
                RecCnt = RecCnt + 1
             End If
           End If
         End Select
       Next i
    End If
  End With
  MsgBox RecCnt & " records appended to the FTICR_AMT database(table).", vbOKOnly
  Set rsFTICR_AMT = Nothing
  Set dbFTICR_AMT = Nothing
Else
  MsgBox "Operation failed( couldn't access FTICR_AMT table).", vbOKOnly
  Set dbFTICR_AMT = Nothing
End If
End Sub


