Attribute VB_Name = "Module1"
'----------------------------------------------------------
'Although some of procedure listed here could be used in
'other components I will not encapsulate for purity reasons
'----------------------------------------------------------
'created: 06/13/2001 nt
'last modified: 12/18/2001 nt
'----------------------------------------------------------
Option Explicit

Public Const FolderSep = "\"
Public Const MyName = "Mass Tags Access"

Const LogFile = "MassTagsAccess.log"
Dim LFName As String

Public MyGl As INFTAXGlobals      'globals

'couple of APIs; ShellExecute is used to launch default browser
'GetDesktopWindow is in fact not important
Public Const SW_SHOWNORMAL = 1

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Type udtMTDBInfoType
    Name As String
    Description As String
    CnStr As String              ' Connection String
    DBState As String            ' Production, Pre-production, Frozen, Unused, Deleted
    DBSchemaVersion As Single
    Server As String
End Type


Public Sub CopyPairsCollection(SourceCol As Collection, _
                               TargetCol As Collection)
'--------------------------------------------------------
'copies items from Source collection to Target collection
'--------------------------------------------------------
Dim i As Long
Dim nv As NameValue
On Error Resume Next
For i = 1 To SourceCol.Count
    Set nv = SourceCol.Item(i)
    TargetCol.Add nv, nv.Name
Next i
End Sub

Public Sub EmptyCollection(SourceCol As Collection)
'--------------------------------------------------
'removes all items from the collection
'--------------------------------------------------
Dim i As Long
On Error Resume Next
For i = 1 To SourceCol.Count
    SourceCol.Remove 1
Next i
End Sub

Public Function CleanComments(ByRef StrArr() As String) As Long
'--------------------------------------------------------------
'removes comments from StrArr array (if after removing comments
'nothing is left array member is deleted; returns number of
'"good lines" or faxa_INIT_FILE_ANY_ERROR on any error
'comment can take full line or be at the line end in which case
'line does not have to be deleted, just cleaned
'--------------------------------------------------------------
On Error GoTo err_CleanComments
Dim i As Long
Dim OKCount As Long
Dim CommentPos As Long
OKCount = 0
For i = 0 To UBound(StrArr)    'delete comments from each line
    CommentPos = InStr(1, StrArr(i), MyGl.INIT_Comment)
    If CommentPos > 0 Then StrArr(i) = Trim$(Left$(StrArr(i), CommentPos - 1))
    If Len(StrArr(i)) > 0 Then
       OKCount = OKCount + 1
       StrArr(OKCount - 1) = StrArr(i)
    End If
Next i
If OKCount > 0 Then
   If OKCount - 1 < UBound(StrArr) Then ReDim Preserve StrArr(OKCount - 1)
Else
   Erase StrArr
End If
CleanComments = OKCount
Exit Function

err_CleanComments:
CleanComments = faxa_INIT_FILE_ANY_ERROR
End Function

Public Function ConstructConnectionString(strServerName As String, strDBName As String, strModelConnectionString As String) As String

    ' Typical ADODB connection string format:
    '  Provider=sqloledb;Data Source=pogo;Initial Catalog=MT_Deinococcus_P20;User ID=mtuser;Password=mt4fun
    ' Typical .NET connection string format:
    '  Server=pogo;database=MT_Main;uid=mtuser;Password=mt4fun

    Dim strConnStrParts() As String
    Dim strParameterName As String
    Dim strNewConnStr As String
    
    Dim intIndex As Integer
    Dim intCharIndex As Integer
    
    strConnStrParts = Split(strModelConnectionString, ";")
    strNewConnStr = ""
    
    For intIndex = 0 To UBound(strConnStrParts)
        intCharIndex = InStr(strConnStrParts(intIndex), "=")
        If intCharIndex > 0 Then
            strParameterName = Left(strConnStrParts(intIndex), intCharIndex - 1)
            Select Case Trim(LCase(strParameterName))
            Case "data source", "server"
                ' Server name
                strConnStrParts(intIndex) = strParameterName & "=" & strServerName
            Case "initial catalog", "database"
                ' DB name
                strConnStrParts(intIndex) = strParameterName & "=" & strDBName
            Case Else
                ' Ignore this entry
            End Select
        End If
        
        If Len(strNewConnStr) > 0 Then
            strNewConnStr = strNewConnStr & ";"
        End If
        strNewConnStr = strNewConnStr & strConnStrParts(intIndex)
        
    Next intIndex
    
    ConstructConnectionString = strNewConnStr

End Function

Public Function GetNamesValues(ByVal sText As String, _
                               ByRef Names() As String, _
                               ByRef Values() As String) As Long
'----------------------------------------------------------------------
'resolves sText in lines and then in names and values
'sText in arrays of Names and Values separated in lines; returns number
'of it/-1 on error; if no "=" is found value is considered to be "None"
'----------------------------------------------------------------------
Dim Lns() As String
Dim LnsCnt As Long
Dim ValPos As Long
Dim i As Long
On Error GoTo err_GetNamesValues

Lns = Split(sText, vbCrLf)
LnsCnt = UBound(Lns) + 1
If LnsCnt > 0 Then
   ReDim Names(LnsCnt - 1)
   ReDim Values(LnsCnt - 1)
   For i = 0 To LnsCnt - 1
       ValPos = InStr(1, Lns(i), MyGl.INIT_Value)
       If ValPos > 0 Then
          Names(i) = Trim(Left$(Lns(i), ValPos - 1))
          Values(i) = Trim$(Right$(Lns(i), Len(Lns(i)) - ValPos))
       Else      'everything is name; value is ""
          Names(i) = Trim(Lns(i))
          Values(i) = ""
       End If
   Next i
End If
GetNamesValues = LnsCnt
Exit Function

err_GetNamesValues:
GetNamesValues = faxa_ANY_ERROR
End Function

Public Sub AddFolderSeparator(SP As String)
'--------------------------------------------------------
'adds folder separator if it's not at the end of the SP
'except if sP is empty string in which case it returns ""
'--------------------------------------------------------
Dim sCoolString
sCoolString = Trim$(SP)
If Len(sCoolString) > 0 Then
   If Right$(sCoolString, 1) <> "\" Then
      SP = sCoolString & "\"
   Else
      SP = sCoolString
   End If
Else
   SP = ""
End If
End Sub

Public Sub Main()
Set MyGl = New INFTAXGlobals
LFName = App.Path & FolderSep & LogFile
End Sub

''
''Public Function GetMTDirectoryConnectionString(ByVal InitFileName As String) As String
''
''Dim MTDirSec() As String
''Dim SecCnt As Long
''Dim MyInit As New InitFile
''Dim strConnectionString As String
''
''On Error GoTo err_GetMTDirectoryConnectionString
''
''SecCnt = MyInit.GetSection(InitFileName, MyGl.SECTION_MT_Directory, MTDirSec())
''Set MyInit = Nothing
''If SecCnt >= 1 Then
''   strConnectionString = MTDirSec(MTDirConnStr)
''Else
''   MsgBox "Missing information in " & MyGl.SECTION_MT_Directory & " section of initialization file: " & InitFileName, vbOKOnly
''   strConnectionString = ""
''End If
''
''GetMTDirectoryConnectionString = strConnectionString
''Exit Function
''
''err_GetMTDirectoryConnectionString:
''   GetMTDirectoryConnectionString = ""
''
''End Function

''Public Function GetMTDirectoryData(ByVal InitFileName As String, _
''                                   ByRef udtMTDBInfo() As udtMTDBInfoType) As Long
'''-----------------------------------------------------------------------
'''retrieves names and connection string for Mass Tag db from catalog
'''(directory)database (MT_Main) and returns 0 if OK; error number if not
'''-----------------------------------------------------------------------
''Dim MTDirSec() As String
''Dim SecCnt As Long
''Dim FldCnStr As String
''Dim FldNames As String
''Dim FldDescs As String
''Dim FldState As String
''
''Dim NewCn As adodb.Connection
''Dim NewRs As adodb.Recordset
''Dim Res As Long
''Dim MTDBsCnt As Long
''Dim MyInit As New InitFile
''On Error GoTo err_GetMTDirectoryData
''
''SecCnt = MyInit.GetSection(InitFileName, MyGl.SECTION_MT_Directory, MTDirSec())
''Set MyInit = Nothing
''If SecCnt >= 5 Then
''   Set NewCn = New adodb.Connection
''   NewCn.ConnectionTimeout = 30
''   NewCn.Open MTDirSec(MTDirConnStr)
''
''   Set NewRs = New adodb.Recordset
''   NewRs.CursorLocation = adUseClient
''   NewRs.Open MTDirSec(MTDirRetrieve), NewCn, adOpenStatic, adLockReadOnly
''
''   FldNames = MTDirSec(MTDirNameFld)
''   FldDescs = MTDirSec(MTDirDescFld)
''   FldCnStr = MTDirSec(MTDirCnStrFld)
''
''   If SecCnt >= 6 Then
''        FldState = MTDirSec(MTDirStateFld)
''   Else
''        FldState = ""
''   End If
''
''   ReDim udtMTDBInfo(257)    'should be plenty, but will expand by 100 if we run out of room
''
''   Do Until NewRs.EOF
''      MTDBsCnt = MTDBsCnt + 1
''      With udtMTDBInfo(MTDBsCnt - 1)
''        .Name = NewRs.Fields(FldNames).Value
''        .Description = NewRs.Fields(FldDescs).Value
''        .CnStr = NewRs.Fields(FldCnStr).Value
''
''        If Len(FldState) > 0 Then
''            .DBState = NewRs.Fields(FldState).Value
''        Else
''            .DBState = "Unknown"
''        End If
''
''        .DBSchemaVersion = 0                ' We'll update this the first time we connect to the database
''      End With
''      NewRs.MoveNext
''   Loop
''Else
''   MsgBox "Missing information in " & MyGl.SECTION_MT_Directory & " section of initialization file!", vbOKOnly
''End If
''
''If MTDBsCnt > 0 Then
''   ReDim Preserve udtMTDBInfo(MTDBsCnt - 1)
''Else
''   Erase udtMTDBInfo
''End If
''
''exit_GetMTDirectoryData:
''NewRs.ActiveConnection = Nothing
''Set NewRs = Nothing
''If NewCn.State <> adStateClosed Then NewCn.Close
''Set NewCn = Nothing
''Exit Function
''
''err_GetMTDirectoryData:
''Select Case Err.Number
''Case 9           'subscript out of range
''   ReDim Preserve udtMTDBInfo(MTDBsCnt + 100)
''   Resume
''Case 13, 94      'type mismatch, invalid use of null
''   Resume Next
''Case Else        'something else
''   GetMTDirectoryData = Err.Number
''   GoTo exit_GetMTDirectoryData
''End Select
''End Function

Public Function GetMTSMasterConnectionString(ByVal InitFileName As String) As String

Dim MTSMasterSec() As String
Dim SecCnt As Long
Dim MyInit As New InitFile
Dim strConnectionString As String

On Error GoTo err_GetMTSMasterConnectionString

SecCnt = MyInit.GetSection(InitFileName, MyGl.SECTION_MTS_Master_DB, MTSMasterSec())
Set MyInit = Nothing
If SecCnt >= 1 Then
   strConnectionString = MTSMasterSec(MTSMasterConnStr)
Else
    ' Missing section, use the default value
   strConnectionString = MyGl.DEFAULT_MTS_MASTER_CONN_STRING
End If

GetMTSMasterConnectionString = strConnectionString
Exit Function

err_GetMTSMasterConnectionString:
   GetMTSMasterConnectionString = ""

End Function

Public Function GetMTSMasterDirectoryData(ByVal InitFileName As String, _
                                          ByRef udtMTDBInfo() As udtMTDBInfoType) As Long
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

Dim NewCn As adodb.Connection
Dim NewRs As adodb.Recordset

Dim strSPName As String
Dim cmdSPGetAllMassTagDBs As New adodb.Command
Dim prmIncludeUnused As adodb.Parameter

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
    Set NewCn = New adodb.Connection
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
NewRs.ActiveConnection = Nothing
Set NewRs = Nothing
If NewCn.State <> adStateClosed Then NewCn.Close
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
   GetMTSMasterDirectoryData = Err.Number
   GoTo exit_GetMTSMasterDirectoryData
End Select
End Function

Public Function GetDBSchemaVersion(ByVal ConnStr As String, ByVal SPName As String) As Single
'---------------------------------------------------------------
'Looks up DB schema version
'---------------------------------------------------------------

    Dim NewCn As adodb.Connection
    Dim cmdSPCommand As New adodb.Command
    Dim prmDBSchemaVersion As adodb.Parameter


    If Len(SPName) <= 0 Then
        SPName = "GetDBSchemaVersion"
    End If

    Set NewCn = New adodb.Connection
    NewCn.Open ConnStr

On Error GoTo GetDBSchemaVersionErrorHandler
    
    Set cmdSPCommand.ActiveConnection = NewCn
    With cmdSPCommand
        .CommandText = SPName
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

Public Function GetMTSubsets(ByVal ConnStr As String, _
                             ByVal MTSSSQL As String, _
                             ByRef MTSSID() As Long, _
                             ByRef MTSSName() As String, _
                             ByRef MTSSDesc() As String) As Long
'---------------------------------------------------------------
'retrieves information about existing Mass Tag database subsets
'connection string for database to search and SQL to retrieve
'subsets list are input parameters
'Assumption is that first returned column will be of type integer
'and will contain ID, second column should list subsets names
'and third subset description
'returns 0 if OK; error number if not
'---------------------------------------------------------------
Dim NewCn As adodb.Connection
Dim NewRs As adodb.Recordset
Dim Res As Long
Dim MTSSCnt As Long
On Error GoTo err_GetMTSubsets

ReDim MTSSID(141)    'should be plenty
ReDim MTSSName(141)
ReDim MTSSDesc(141)

If Len(MTSSSQL) > 0 Then
   Set NewCn = New adodb.Connection
   NewCn.Open ConnStr

   Set NewRs = New adodb.Recordset
   NewRs.CursorLocation = adUseClient
   NewRs.Open MTSSSQL, NewCn, adOpenStatic, adLockReadOnly
      
   Do Until NewRs.EOF
      MTSSCnt = MTSSCnt + 1
      MTSSID(MTSSCnt - 1) = NewRs.Fields(0).Value
      MTSSName(MTSSCnt - 1) = NewRs.Fields(1).Value
      MTSSDesc(MTSSCnt - 1) = NewRs.Fields(2).Value
      NewRs.MoveNext
   Loop
End If

If MTSSCnt > 0 Then
   ReDim Preserve MTSSID(MTSSCnt - 1)
   ReDim Preserve MTSSName(MTSSCnt - 1)
   ReDim Preserve MTSSDesc(MTSSCnt - 1)
Else
   Erase MTSSID
   Erase MTSSName
   Erase MTSSDesc
End If

exit_GetMTSubsets:
NewRs.ActiveConnection = Nothing
Set NewRs = Nothing
If NewCn.State <> adStateClosed Then NewCn.Close
Set NewCn = Nothing
Exit Function

err_GetMTSubsets:
Select Case Err.Number
Case 9           'subscript out of range
   ReDim Preserve MTSSID(MTSSCnt + 100)
   ReDim Preserve MTSSName(MTSSCnt + 100)
   ReDim Preserve MTSSDesc(MTSSCnt + 100)
   Resume
Case 13, 94      'type mismatch, invalid use of null
   Resume Next
Case Else        'something else
   GetMTSubsets = Err.Number
   GoTo exit_GetMTSubsets
End Select
End Function


Public Function GetGlobMods(ByVal ConnStr As String, _
                            ByVal GlobModViewName As String, _
                            ByRef GlobModID() As Long, _
                            ByRef GlobModName() As String, _
                            ByRef GlobModDesc() As String) As Long
'------------------------------------------------------------------
'retrieves information about global modifications associated with
'current mass tag database
'Assumption is that 1st column will be of type long and will contain
'ID, second column should list mods names third mods description
'returns 0 if OK; error number if not
'NOTE: this still works although Dynamic and Static modifications
'are now separated
'------------------------------------------------------------------
Dim NewCn As adodb.Connection
Dim NewRs As adodb.Recordset
Dim Res As Long
Dim GlobModCnt As Long
On Error GoTo err_GetGlobMods

ReDim GlobModID(141)    'should be plenty
ReDim GlobModName(141)
ReDim GlobModDesc(141)

If Len(GlobModViewName) > 0 Then
   Set NewCn = New adodb.Connection
   NewCn.Open ConnStr

   Set NewRs = New adodb.Recordset
   NewRs.CursorLocation = adUseClient
   NewRs.Open GlobModViewName, NewCn, adOpenStatic, adLockReadOnly
      
   Do Until NewRs.EOF
      GlobModCnt = GlobModCnt + 1
      GlobModID(GlobModCnt - 1) = NewRs.Fields(0).Value
      GlobModName(GlobModCnt - 1) = Trim$(NewRs.Fields(1).Value)
      GlobModDesc(GlobModCnt - 1) = NewRs.Fields(2).Value
      NewRs.MoveNext
   Loop
End If

If GlobModCnt > 0 Then
   ReDim Preserve GlobModID(GlobModCnt - 1)
   ReDim Preserve GlobModName(GlobModCnt - 1)
   ReDim Preserve GlobModDesc(GlobModCnt - 1)
Else
   Erase GlobModID
   Erase GlobModName
   Erase GlobModDesc
End If

exit_GetGlobMods:
NewRs.ActiveConnection = Nothing
Set NewRs = Nothing
If NewCn.State <> adStateClosed Then NewCn.Close
Set NewCn = Nothing
Exit Function

err_GetGlobMods:
Select Case Err.Number
Case 9           'subscript out of range
   ReDim Preserve GlobModID(GlobModCnt + 100)
   ReDim Preserve GlobModName(GlobModCnt + 100)
   ReDim Preserve GlobModDesc(GlobModCnt + 100)
   Resume
Case 13, 94      'type mismatch, invalid use of null
   Resume Next
Case Else        'something else
   Erase GlobModID
   Erase GlobModName
   Erase GlobModDesc
   GetGlobMods = Err.Number
   GoTo exit_GetGlobMods
End Select
End Function

Public Function GetInternalStandardNames(ByVal ConnStr As String, _
                                         ByVal InternalStdsViewName As String, _
                                         ByRef intInternalStandardCount As Integer, _
                                         ByRef strInternalStandardNames() As String) As Long
    
    '------------------------------------------------------------------
    ' Retrieves information about Internal Standards defined
    '  Assumption is that 1st column will be of type long and will contain
    '  ID, second column should list internal standard name, and third the description
    '  Populates strInternalStandardNames() with the names
    '
    ' Returns 0 if OK; error number if not
    '------------------------------------------------------------------
    
    Dim NewCn As adodb.Connection
    Dim NewRs As adodb.Recordset
    Dim Res As Long
    
On Error GoTo GetInternalStandardNamesErrorHandler
    
    ' Initially reserve space for 100 internal standards
    intInternalStandardCount = 0
    ReDim strInternalStandardNames(100)
    
    If Len(InternalStdsViewName) > 0 Then
       Set NewCn = New adodb.Connection
       NewCn.Open ConnStr
    
       Set NewRs = New adodb.Recordset
       NewRs.CursorLocation = adUseClient
       NewRs.Open InternalStdsViewName, NewCn, adOpenStatic, adLockReadOnly
          
       Do Until NewRs.EOF
          strInternalStandardNames(intInternalStandardCount) = Trim(NewRs.Fields(1).Value)
          intInternalStandardCount = intInternalStandardCount + 1
          NewRs.MoveNext
       Loop
    End If
    
    If intInternalStandardCount > 0 Then
       ReDim Preserve strInternalStandardNames(intInternalStandardCount - 1)
    Else
       Erase strInternalStandardNames
    End If
    
GetInternalStandardNamesExit:
    NewRs.ActiveConnection = Nothing
    Set NewRs = Nothing
    If NewCn.State <> adStateClosed Then NewCn.Close
    Set NewCn = Nothing
    Exit Function
    
GetInternalStandardNamesErrorHandler:
    Select Case Err.Number
    Case 9           'subscript out of range
       ReDim Preserve strInternalStandardNames(intInternalStandardCount + 100)
       Resume
    Case 13, 94      'type mismatch, invalid use of null
       Resume Next
    Case Else        'something else
       Erase strInternalStandardNames
       GetInternalStandardNames = Err.Number
       Resume GetInternalStandardNamesExit
    End Select

End Function

Public Sub LogMessages(ByVal sMsg As String)
Dim fsobj As New FileSystemObject
Dim ts As TextStream
On Error Resume Next

Set ts = fsobj.OpenTextFile(LFName, ForAppending, True)
ts.WriteLine sMsg
ts.Close
Set ts = Nothing
Set fsobj = Nothing
End Sub


Public Function GetLockerTypes(ByVal ConnStr As String, _
                               ByVal LT_SQL As String, _
                               ByRef LT_ID() As Long, _
                               ByRef LT_Name() As String) As Long
'----------------------------------------------------------------
'retrieves information about existing locker types
'connection string for database to search and SQL to retrieve
'subsets list are input parameters
'Assumtpion is that first returned column will be of type integer
'and will contain ID, second column should list locker type names
'returns 0 if OK; error number if not
'----------------------------------------------------------------
Dim NewCn As adodb.Connection
Dim NewRs As adodb.Recordset
Dim Res As Long
Dim LT_Cnt As Long
On Error GoTo err_GetLockerTypes

ReDim LT_ID(141)    'should be plenty
ReDim LT_Name(141)

If Len(LT_SQL) > 0 Then
   Set NewCn = New adodb.Connection
   NewCn.Open ConnStr

   Set NewRs = New adodb.Recordset
   NewRs.CursorLocation = adUseClient
   NewRs.Open LT_SQL, NewCn, adOpenStatic, adLockReadOnly
      
   Do Until NewRs.EOF
      LT_Cnt = LT_Cnt + 1
      LT_ID(LT_Cnt - 1) = NewRs.Fields(0).Value
      LT_Name(LT_Cnt - 1) = NewRs.Fields(1).Value
      NewRs.MoveNext
   Loop
End If

If LT_Cnt > 0 Then
   ReDim Preserve LT_ID(LT_Cnt - 1)
   ReDim Preserve LT_Name(LT_Cnt - 1)
Else
   Erase LT_ID
   Erase LT_Name
End If

exit_GetLockerTypes:
NewRs.ActiveConnection = Nothing
Set NewRs = Nothing
If NewCn.State <> adStateClosed Then NewCn.Close
Set NewCn = Nothing
Exit Function

err_GetLockerTypes:
Select Case Err.Number
Case 9           'subscript out of range
   ReDim Preserve LT_ID(LT_Cnt + 100)
   ReDim Preserve LT_Name(LT_Cnt + 100)
   Resume
Case 13, 94      'type mismatch, invalid use of null
   Resume Next
Case Else        'something else
   GetLockerTypes = Err.Number
   GoTo exit_GetLockerTypes
End Select
End Function



Public Sub RunShellExecute(ByVal sTopic As String, _
                           ByVal sFile As Variant, _
                           ByVal sParams As Variant, _
                           ByVal sDirectory As Variant, _
                           ByVal nShowCmd As Long)
'-------------------------------------------------------
'runs ShellExecute with specified parameters; sends all
'error messages to desktop window
'-------------------------------------------------------
Dim hWnd As Long
Dim Res As Long
On Error Resume Next
hWnd = GetDesktopWindow()
Res = ShellExecute(hWnd, sTopic, sFile, sParams, sDirectory, nShowCmd)
End Sub

Public Sub SortMTDBNameList(ByRef MTDBInfo() As udtMTDBInfoType, ByRef MTDBNameListPointers() As Long, ByVal lngLowIndex As Long, ByVal lngHighIndex As Long, ByVal blnShowFrozenDBs As Boolean, ByVal blnShowUnusedDBs As Boolean)
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

Public Sub TestAnalysisInitiator()
    Const INIT_FILE_NAME = "FAXA.init"
    Dim objNewAnalysis As New AnalysisInitiator
    
    objNewAnalysis.GetNewAnalysisDialog (App.Path & "\" & INIT_FILE_NAME)
    
    Set objNewAnalysis = Nothing
    
End Sub

Public Sub TestDummyAnalysisInitiator()
    Const INIT_FILE_NAME = "FAXA.init"
    Dim objNewAnalysis As New DummyAnalysisInitiator
    
    objNewAnalysis.GetNewAnalysisDialog (App.Path & "\" & INIT_FILE_NAME)
    
    Set objNewAnalysis = Nothing
    
End Sub

Public Sub TestMTDBInfoRetriever()
    Dim objMTDB As New MTDBInfoRetriever
    
    objMTDB.InitFilePath = App.Path & "\faxa.init"
    objMTDB.GetMTDBSchema
    
    With objMTDB.fAnalysis.MTDB
        Debug.Print .DBStuff.Count
    End With
    
    Set objMTDB = Nothing
    
End Sub
