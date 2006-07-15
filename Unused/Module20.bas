Attribute VB_Name = "Module20"
Option Explicit


Public Sub LoadMassTags(ByVal Ind As Long)
'-------------------------------------------------
'executes command that retrieves mass tags from
'Organism Mass Tags database
'-------------------------------------------------
Dim cnNew As New ADODB.Connection
Dim rsMassTags As New ADODB.Recordset
Dim cmdGetMassTags As New ADODB.Command
Dim prmOrgDBName As New ADODB.Parameter     'organism database name
Dim prmRetVal As New ADODB.Parameter        'return value from stored procedure
Dim ErrCnt As Long                          'list only first 10 errors
On Error GoTo err_LoadMassTags
'reserve space for 1000 lockers; increase in chunks of 200 after that
ReDim AMTID(100000)
ReDim AMTFlag(100000)
ReDim AMTMW(100000)
ReDim AMTNET(100000)
ReDim AMTRT(100000)
ReDim AMTCNT_N(100000)
ReDim AMTCNT_Cys(100000)

Screen.MousePointer = vbHourglass
AMTCnt = 0
cnNew.ConnectionString = GelAnalysis(Ind).MTDB.cn.ConnectionString
cnNew.Open
'create and tune command object to retrieve lockers
Set cmdGetMassTags.ActiveConnection = cnNew
cmdGetMassTags.CommandText = GetLockersCommand
cmdGetMassTags.CommandType = adCmdStoredProc
'procedure takes one parameter and returns error number or 0 if OK
Set prmOrgDBName = cmdGetMassTags.CreateParameter("OrgDBName", adVarChar, adParamInput, 50, GelAnalysis(Ind).Organism_DB_Name)
cmdGetMassTags.Parameters.Append prmOrgDBName
Set rsMassTags = cmdGetMassTags.Execute
With rsMassTags
    'load lockers data
    Do Until .EOF
       LckCnt = LckCnt + 1
       LckID(LckCnt - 1) = .Fields(0).value
       LckSeq(LckCnt - 1) = .Fields(1).value
       LckName(LckCnt - 1) = .Fields(2).value
       LckMW(LckCnt - 1) = .Fields(3).value
       LckNET(LckCnt - 1) = .Fields(4).value
       LckRET(LckCnt - 1) = .Fields(5).value
       LckOET(LckCnt - 1) = .Fields(6).value
       LckScore(LckCnt - 1) = .Fields(7).value
       .MoveNext
    Loop
End With
rsMassTags.Close
lblStatus.Caption = "Number of loaded lockers: " & LckCnt

'clean things and exit
exit_cmdLoadMassTags:
On Error Resume Next
Set cmdGetMassTags.ActiveConnection = Nothing
cnNew.Close
If LckCnt > 0 Then
   If LckCnt - 1 > UBound(LckID) Then
      ReDim Preserve LckID(LckCnt - 1)
      ReDim Preserve LckSeq(LckCnt - 1)
      ReDim Preserve LckName(LckCnt - 1)
      ReDim Preserve LckMW(LckCnt - 1)
      ReDim Preserve LckScore(LckCnt - 1)
      ReDim LckET(LckCnt - 1, 2)
   End If
   For i = 0 To LckCnt - 1
      LckCnt(i, NET_COL) = LckNET(i)
      LckCnt(i, RET_COL) = LckRET(i)
      LckCnt(i, OET_COL) = LckOET(i)
   Next i
   'enable commands for listing lockers and search
   cmdListLockers.Enabled = True
   'load data from gel and if successful enable search functions
   lblStatus.Caption = "Loading data from current file!"
   DoEvents
   If FillArrays() Then
      cmdSearch.Enabled = True
      lblStatus.Caption = MWCnt & " data distributions loaded!"
   Else
      lblStatus.Caption = "Data distributions missing. Locking masses is mission impossible!"
   End If
Else
   Erase LckID
   Erase LckSeq
   Erase LckName
   Erase LckMW
   Erase LckScore
End If
Screen.MousePointer = vbDefault
Exit Sub

err_cmdLoadMassTags:
Select Case Err.Number
Case 9                       'need more room for lockers
    ReDim Preserve LckID(LckCnt + 200)
    ReDim Preserve LckSeq(LckCnt + 200)
    ReDim Preserve LckName(LckCnt + 200)
    ReDim Preserve LckMW(LckCnt + 200)
    ReDim Preserve LckNET(LckCnt + 200)
    ReDim Preserve LckRET(LckCnt + 200)
    ReDim Preserve LckOET(LckCnt + 200)
    ReDim Preserve LckScore(LckCnt + 200)
    Resume Next
Case 13, 94                  'Type Mismatch or Invalid Use of Null
    Resume Next
Case Else
    ErrCnt = ErrCnt + 1
    lblStatus.Caption = "Error retrieving lockers information from the database!"
    If ErrCnt < 10 Then
       LogErrors Err.Number, "frmMTLockMass.cmdLoadLockers", Err.Description
       Resume Next
    End If
End Select
GoTo exit_cmdLoadMassTags
AMTCnt = -1
End Sub

