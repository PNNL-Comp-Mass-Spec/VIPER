Attribute VB_Name = "MTAMIsc"
Option Explicit

Public Type udtMassTagsAccessMTDBInfoType
    Name As String
    Description As String
    CnStr As String              ' Connection String
    DBState As String            ' Production, Pre-production, Frozen, Unused, Deleted
    DBSchemaVersion As Single
    Server As String
End Type

Public Sub AddUpdateNameValueEntry(ByRef objCol As Collection, ByVal PairName As String, ByVal NewValue As String)
'-------------------------------------------------------------------------
'modifies value of name value pair; if pair does not exist adds it
'-------------------------------------------------------------------------
Dim nv As NameValue
On Error Resume Next
objCol.Item(PairName).Value = NewValue
If Err Then
   Set nv = New NameValue
   nv.Name = PairName
   nv.Value = NewValue
   objCol.add nv, nv.Name
End If
End Sub

Public Function GetDBSchemaVersion(ByVal ConnStr As String, ByVal spName As String) As Single
'---------------------------------------------------------------
'Looks up DB schema version
'---------------------------------------------------------------

    Dim NewCn As ADODB.Connection
    Dim cmdSPCommand As New ADODB.Command
    Dim prmDBSchemaVersion As ADODB.Parameter


    If Len(spName) <= 0 Then
        spName = "GetDBSchemaVersion"
    End If

    Set NewCn = New ADODB.Connection
    NewCn.Open ConnStr

On Error GoTo GetDBSchemaVersionErrorHandler
    
    Set cmdSPCommand.ActiveConnection = NewCn
    With cmdSPCommand
        .CommandText = spName
        .CommandType = adCmdStoredProc
        .CommandTimeout = 30            ' Timeout in seconds
    End With
    
    
    Set prmDBSchemaVersion = cmdSPCommand.CreateParameter("DBSchemaVersion", adSingle, adParamOutput, , 0)
    cmdSPCommand.Parameters.Append prmDBSchemaVersion
    
    cmdSPCommand.Execute
    
    GetDBSchemaVersion = prmDBSchemaVersion.Value
    Exit Function

GetDBSchemaVersionErrorHandler:
    Debug.Assert False
    GetDBSchemaVersion = 1

End Function

Public Function GetMTSMasterDirectoryData(ByVal InitFileName As String, _
                                          ByRef udtMTDBInfo() As udtMassTagsAccessMTDBInfoType) As Long
    '-----------------------------------------------------------------------
    'retrieves names and connection string for Mass Tag DB from catalog
    '(directory)database (MTS_Master) and returns 0 if OK; error number if not
    '-----------------------------------------------------------------------
    Dim MTSMasterSec() As String
    Dim SecCnt As Long
    Dim FldNames As String
    Dim FldDescs As String
    Dim FldState As String
    Dim FldServer As String
    Dim FldDBSchemaVersion As String
    
    Dim NewCn As ADODB.Connection
    Dim NewRs As ADODB.Recordset
    
    Dim strSPName As String
    Dim cmdSPGetAllMassTagDBs As New ADODB.Command
    Dim prmIncludeUnused As ADODB.Parameter
    
    Dim Res As Long
    Dim MTDBsCnt As Long
    Dim MyInit As New InitFile
    Dim strMTSMasterConnStr As String
    
    On Error GoTo err_GetMTSMasterDirectoryData
    
    SecCnt = MyInit.GetSection(InitFileName, MyGl.SECTION_MTS_Master_DB, MTSMasterSec())
    Set MyInit = Nothing
    If SecCnt <= 0 Then
       ' Missing section; use the default values
       ReDim MTSMasterSec(6)
       MTSMasterSec(MTSMasterConnStr) = MyGl.DEFAULT_MTS_MASTER_CONN_STRING
       MTSMasterSec(MTSMasterSPRetrieve) = "GetAllMassTagDatabases"
       MTSMasterSec(MTSMasterNameFld) = "Name"
       MTSMasterSec(MTSMasterDescFld) = "Description"
       MTSMasterSec(MTSMasterStateFld) = "State"
       MTSMasterSec(MTSMasterServerNameFld) = "Server Name"
       SecCnt = 6
    End If
    
    If SecCnt >= 6 Then
    
        strMTSMasterConnStr = MTSMasterSec(MTSMasterConnStr)
        strSPName = MTSMasterSec(MTSMasterSPRetrieve)
        
        ' Connect to the server
        Set NewCn = New ADODB.Connection
        NewCn.ConnectionTimeout = 30
        NewCn.Open strMTSMasterConnStr
    
        ' Define the stored procedure command
        Set cmdSPGetAllMassTagDBs.ActiveConnection = NewCn
        With cmdSPGetAllMassTagDBs
            .CommandText = strSPName
            .CommandType = adCmdStoredProc
            .CommandTimeout = 30
        End With
        
        If LCase(strSPName) = "getallmasstagdatabases" Then
            ' Define parameter prmIncludeUnused
            Set prmIncludeUnused = cmdSPGetAllMassTagDBs.CreateParameter("IncludeUnused", adTinyInt, adParamInput, , 1)
            cmdSPGetAllMassTagDBs.Parameters.Append prmIncludeUnused
        End If
        
        ' Call the SP
        Set NewRs = cmdSPGetAllMassTagDBs.Execute()
        
        FldNames = MTSMasterSec(MTSMasterNameFld)
        FldDescs = MTSMasterSec(MTSMasterDescFld)
        FldState = MTSMasterSec(MTSMasterStateFld)
        FldServer = MTSMasterSec(MTSMasterServerNameFld)
        FldDBSchemaVersion = "DB Schema Version"
    
        ReDim udtMTDBInfo(512)    'should be plenty, but will expand by 100 if we run out of room
        
        MTDBsCnt = 0
        Do Until NewRs.EOF
            With udtMTDBInfo(MTDBsCnt)
                .Name = NewRs.Fields(FldNames).Value
                .Description = NewRs.Fields(FldDescs).Value
                .DBState = NewRs.Fields(FldState).Value
                .Server = NewRs.Fields(FldServer).Value
                .CnStr = ConstructConnectionString(.Server, .Name, strMTSMasterConnStr)
                .DBSchemaVersion = NewRs.Fields(FldDBSchemaVersion)
            End With
            MTDBsCnt = MTDBsCnt + 1
            NewRs.MoveNext
        Loop
    Else
       MsgBox "Missing information in " & MyGl.SECTION_MTS_Master_DB & " section of initialization file!", vbOKOnly
    End If
    
    If MTDBsCnt > 0 Then
       ReDim Preserve udtMTDBInfo(MTDBsCnt - 1)
    Else
       Erase udtMTDBInfo
    End If
    
exit_GetMTSMasterDirectoryData:
    Set NewRs = Nothing
    If NewCn.STATE <> adStateClosed Then NewCn.Close
    Set NewCn = Nothing
    Exit Function
    
err_GetMTSMasterDirectoryData:
    Select Case Err.Number
    Case 9           'subscript out of range
       ReDim Preserve udtMTDBInfo(MTDBsCnt + 100)
       Resume
    Case 13, 94      'type mismatch, invalid use of null
       Resume Next
    Case Else        'something else
       Debug.Assert False
       GetMTSMasterDirectoryData = Err.Number
       GoTo exit_GetMTSMasterDirectoryData
    End Select
End Function

Public Sub SortMTDBNameList(ByRef MTDBInfo() As udtMassTagsAccessMTDBInfoType, ByRef MTDBNameListPointers() As Long, ByVal lngLowIndex As Long, ByVal lngHighIndex As Long, ByVal blnShowFrozenDBs As Boolean, ByVal blnShowUnusedDBs As Boolean)
    Dim lngCount As Long
    Dim lngIncrement As Long
    Dim lngIndex As Long
    Dim lngIndexCompare As Long
    Dim lngCompareValPointer As Long
    Dim blnAddDB As Boolean
    
    lngCount = lngHighIndex - lngLowIndex + 1
    
    If lngCount <= 0 Then
        ReDim MTDBNameListPointers(0)
        Exit Sub
    End If
    
    ' Populate MTDBNameListPointers()
    ReDim MTDBNameListPointers(lngLowIndex To lngHighIndex)
    
    lngCount = 0
    For lngIndex = lngLowIndex To lngHighIndex
        If blnShowFrozenDBs And blnShowUnusedDBs Then
            blnAddDB = True
        Else
            Select Case UCase(MTDBInfo(lngIndex).DBState)
            Case "DELETED"
                blnAddDB = False
            Case "FROZEN"
                blnAddDB = blnShowFrozenDBs
            Case "UNUSED"
                blnAddDB = blnShowUnusedDBs
            Case Else
                If Left(MTDBInfo(lngIndex).DBState, 5) = "MOVED" Then
                    blnAddDB = False
                Else
                    blnAddDB = True
                End If
                
            End Select
        End If
        
        If blnAddDB Then
            ' Include MTDB
            MTDBNameListPointers(lngCount) = lngIndex
            lngCount = lngCount + 1
        End If
    Next lngIndex
    
    lngHighIndex = lngLowIndex + lngCount - 1
    If lngCount <= 0 Or lngHighIndex < lngLowIndex Then
        ReDim MTDBNameListPointers(0)
        Exit Sub
    End If
    
    If Not (blnShowFrozenDBs And blnShowUnusedDBs) Then
        ReDim Preserve MTDBNameListPointers(lngLowIndex To lngHighIndex)
    End If
    
    ' Sort using a Shell Sort
    ' Don't actually change MTDBNameList()
    ' Instead, change MTDBNameListPointers()

    ' compute largest increment
    lngIncrement = 1
    If (lngCount < 14) Then
        lngIncrement = 1
    Else
        Do While lngIncrement < lngCount
            lngIncrement = 3 * lngIncrement + 1
        Loop
        lngIncrement = lngIncrement \ 3
        lngIncrement = lngIncrement \ 3
    End If

    Do While lngIncrement > 0
        ' sort by insertion in increments of lngIncrement
        For lngIndex = lngLowIndex + lngIncrement To lngHighIndex
            lngCompareValPointer = MTDBNameListPointers(lngIndex)
            For lngIndexCompare = lngIndex - lngIncrement To lngLowIndex Step -lngIncrement
                ' Use <= to sort ascending; Use > to sort descending
                If MTDBInfo(MTDBNameListPointers(lngIndexCompare)).Name <= MTDBInfo(lngCompareValPointer).Name Then Exit For
                MTDBNameListPointers(lngIndexCompare + lngIncrement) = MTDBNameListPointers(lngIndexCompare)
            Next lngIndexCompare
            MTDBNameListPointers(lngIndexCompare + lngIncrement) = lngCompareValPointer
        Next lngIndex
        lngIncrement = lngIncrement \ 3
    Loop
    
    
End Sub

Public Sub VerifyNameDefined(ByRef objCol As Collection, ByVal PairName As String, ByVal ValueIfMissing As String)
    '-------------------------------------------------------------------------
    'Makes sure name/value pair exists in DBStuff; adds it if missing
    '-------------------------------------------------------------------------
    Dim nv As NameValue
    Dim strValue As String
    On Error Resume Next
    strValue = objCol.Item(PairName).Value
    If Err Then
       Set nv = New NameValue
       nv.Name = PairName
       nv.Value = ValueIfMissing
       objCol.add nv, nv.Name
    End If
End Sub



