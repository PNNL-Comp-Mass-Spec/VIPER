VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mtdbMTNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'this class serves as implementation of MT tag names
'it loads table T_Mass_Tag_to_ORF_Map and allows for fast search of it
'Class does not raise any error; rather client should inspect state of
'DataStatus property

'Properties mSQL and mDBConnString should be set and FillData method
'should be invoked. If it fails client should inspect value of DataStatus
'property and if in state dsErrorTryAgain advise user to try ot again


'Methods GetNamesForMTID and GetNamesForRefID will fill array with peptide
'names. This methods should be invoked only if FillData succeseds

'created: 11/25/2002 nt
'last modified: 11/26/2002 nt
'-------------------------------------------------------------------------
Option Explicit

Public Enum mtdbDataStatus
    dsEmpty = 0
    dsError = 1
    dsLoaded = 2
    dsErrorTryAgain = 3
End Enum

Public Enum ArrayManagement
    amErase = 0
    amReDim = 1
    amReDimPreserve = 2
    amTruncate = 3
End Enum

Const ERR_TRY_AGAIN = 34

Dim mSQL As String
Dim mDBConnString As String

Dim mCnt As Long                        'count of names
Dim mDataStatus As mtdbDataStatus

Dim mMTID() As Long         'MT tag ids   (points to T_Mass_Tags table)
Dim mMTName() As String     'MT tag names
Dim mRefID() As Long        'Ref ID   (points to T_ORF_Reference table)

Dim mOrigInd() As Long      'original index
Dim mMTIDInd() As Long      'index for order on MT tags ID
Dim mRefIDInd() As Long     'index for order on reference IDs

'make following search variables class-level to simplify search
Dim mCurrMTID As Long
Dim mCurrRefID As Long

Public Property Get RetrieveSQL() As String
RetrieveSQL = mSQL
End Property

Public Property Let RetrieveSQL(ByVal NewSQL As String)
mSQL = NewSQL
End Property

Public Property Get DBConnectionString() As String
DBConnectionString = mDBConnString
End Property

Public Property Let DBConnectionString(ByVal NewConnectionString As String)
mDBConnString = NewConnectionString
End Property

Public Property Get DataStatus() As String
DataStatus = mDataStatus
End Property


Public Function DeleteData() As Boolean
'------------------------------------------------------------
'erases data and destroys data structures; returns True if OK
'------------------------------------------------------------
On Error Resume Next
mCnt = 0
mDataStatus = dsEmpty
ManageArrays amErase, 0
End Function


Public Function FillData(Optional frmThisForm As VB.Form) As Boolean
Dim cnNew As New ADODB.Connection
Dim rsMTName As New ADODB.Recordset
Dim cmdGetMTName As New ADODB.Command
Dim blnShowProgress As Boolean
Dim strCaptionSaved As String

On Error GoTo err_FillData

If Not frmThisForm Is Nothing Then
    blnShowProgress = True
    strCaptionSaved = frmThisForm.Caption
    frmThisForm.Caption = "Loading Protein names: Connecting to " & ExtractDBNameFromConnectionString(mDBConnString)
End If

If Not ManageArrays(amReDim, 75000) Then GoTo err_FillData

If Not EstablishConnection(cnNew, mDBConnString, False) Then
    Debug.Assert False
    FillData = False
    Exit Function
End If

'create and tune command object to retrieve MT tag names
Set cmdGetMTName.ActiveConnection = cnNew
cmdGetMTName.CommandText = mSQL
cmdGetMTName.CommandType = adCmdText
cmdGetMTName.CommandTimeout = glbPreferencesExpanded.AutoAnalysisOptions.DBConnectionTimeoutSeconds
Set rsMTName = cmdGetMTName.Execute
With rsMTName
    .MoveFirst
    Do Until .EOF
       If blnShowProgress Then
          If mCnt Mod 1000 = 0 Then frmThisForm.Caption = "Loading Protein names for MT tags: " & Trim(mCnt) & " loaded"
       End If
       mCnt = mCnt + 1
       mMTID(mCnt - 1) = FixNullLng(.Fields(0).Value)
       mMTName(mCnt - 1) = FixNull(.Fields(1).Value)
       mRefID(mCnt - 1) = FixNullLng(.Fields(2).Value)
       mOrigInd(mCnt - 1) = mCnt - 1
       mMTIDInd(mCnt - 1) = mCnt - 1
       mRefIDInd(mCnt - 1) = mCnt - 1
       .MoveNext
    Loop
    .Close
End With
FillData = True

If blnShowProgress Then frmThisForm.Caption = strCaptionSaved

exit_FillData:
Set cmdGetMTName.ActiveConnection = Nothing
cnNew.Close
If ManageArrays(amTruncate, 0) Then
   If CreateMTIDIndex() And CreateRefIDIndex() Then
      mDataStatus = dsLoaded
   Else
      ManageArrays amErase, 0
      mDataStatus = dsError
   End If
Else
   ManageArrays amErase, 0
End If
Exit Function

err_FillData:
Select Case Err.Number
Case 9                              'add more space
    If ManageArrays(amReDimPreserve, 10000) Then
       Resume
    Else
       ManageArrays amErase, 0
       mDataStatus = dsError
    End If
Case 13, 94                  'Type Mismatch or Invalid Use of Null
    Resume Next              'just ignore it
Case 3265, 3704              'two errors I have encountered
    '2nd attempt will probably work so let user know it should try again
    mDataStatus = dsErrorTryAgain
Case Else
    mDataStatus = dsError
End Select
Resume exit_FillData
End Function

Public Function GetORFCount() As Long
    On Error Resume Next
    GetORFCount = UBound(mMTID())
End Function

Private Function ManageArrays(ByVal ManageType As ArrayManagement, ByVal Size As Long) As Boolean
On Error GoTo exit_ManageArrays
Select Case ManageType
Case amErase
        Erase mMTID
        Erase mMTName
        Erase mRefID
        Erase mOrigInd
        Erase mMTIDInd
        Erase mRefIDInd
Case amReDim
        ReDim mMTID(Size - 1)
        ReDim mMTName(Size - 1)
        ReDim mRefID(Size - 1)
        ReDim mOrigInd(Size - 1)
        ReDim mMTIDInd(Size - 1)
        ReDim mRefIDInd(Size - 1)
Case amReDimPreserve
        ReDim Preserve mMTID(mCnt + Size - 1)
        ReDim Preserve mMTName(mCnt + Size - 1)
        ReDim Preserve mRefID(mCnt + Size - 1)
        ReDim Preserve mOrigInd(mCnt + Size - 1)
        ReDim Preserve mMTIDInd(mCnt + Size - 1)
        ReDim Preserve mRefIDInd(mCnt + Size - 1)
Case amTruncate
        ReDim Preserve mMTID(mCnt - 1)
        ReDim Preserve mMTName(mCnt - 1)
        ReDim Preserve mRefID(mCnt - 1)
        ReDim Preserve mOrigInd(mCnt - 1)
        ReDim Preserve mMTIDInd(mCnt - 1)
        ReDim Preserve mRefIDInd(mCnt - 1)
End Select
ManageArrays = True
exit_ManageArrays:
End Function


Private Function CreateMTIDIndex() As Boolean
Dim QSL As New QSLong
On Error Resume Next
CreateMTIDIndex = QSL.QSAsc(mMTID(), mMTIDInd())
End Function


Private Function CreateRefIDIndex() As Boolean
Dim QSL As New QSLong
On Error Resume Next
CreateRefIDIndex = QSL.QSAsc(mRefID(), mRefIDInd())
End Function


Public Function GetNamesForMTID(MTID As Long, MTNames() As String) As Long
'-------------------------------------------------------------------------
'fills names for MT tag ID; returns number of it; -1 on any error
'-------------------------------------------------------------------------
Dim FirstInd As Long
Dim LastInd As Long
Dim i As Long
On Error GoTo err_GetNamesForMTID
mCurrMTID = MTID
FirstInd = 0
LastInd = mCnt - 1
FindMTIDRange FirstInd, LastInd
If FirstInd > LastInd Then
   Erase MTNames
   GetNamesForMTID = 0
Else
   ReDim MTNames(LastInd - FirstInd)
   For i = FirstInd To LastInd
       MTNames(i - FirstInd) = mMTName(mMTIDInd(i))
   Next i
   GetNamesForMTID = UBound(MTNames) + 1
End If
Exit Function

err_GetNamesForMTID:
GetNamesForMTID = -1
Erase MTNames
End Function


Public Function GetNamesForRefID(RefID As Long, MTNames() As String) As Long
'---------------------------------------------------------------------------
'fills names for reference ID; returns number of it; -1 on any error
'---------------------------------------------------------------------------
Dim FirstInd As Long
Dim LastInd As Long
Dim i As Long
On Error GoTo err_GetNamesForRefID
mCurrRefID = RefID
FirstInd = 0
LastInd = mCnt - 1
FindRefIDRange FirstInd, LastInd
If FirstInd > LastInd Then
   Erase MTNames
   GetNamesForRefID = 0
Else
   ReDim MTNames(LastInd - FirstInd + 1)
   For i = FirstInd To LastInd
       MTNames(i - FirstInd) = mMTName(mRefIDInd(i))
   Next i
   GetNamesForRefID = UBound(MTNames) + 1
End If
Exit Function

err_GetNamesForRefID:
GetNamesForRefID = -1
Erase MTNames
End Function


Private Sub FindMTIDRange(MinInd As Long, MaxInd As Long)
'-------------------------------------------------------------------------------
'recursive procedure to find all elements of array mMTID() that match mCurrMTID
'Note that array mMTID must be in ascending order for this to work
'-------------------------------------------------------------------------------
Dim MidInd As Long
Dim LeftDone As Boolean
Dim RightDone As Boolean
Dim LInd As Long
Dim RInd As Long

MidInd = (MinInd + MaxInd) \ 2
If MidInd = MinInd Then  'Min and Max next to each other
   'this will produce MinInd>MaxInd if none of them match mCurrMTID
   If mCurrMTID <> mMTID(MinInd) Then MinInd = MaxInd
   If mCurrMTID <> mMTID(MaxInd) Then MaxInd = MidInd
   Exit Sub         'we are done
End If
If mMTID(MidInd) > mCurrMTID Then        'we are out of range on right
   MaxInd = MidInd
   FindMTIDRange MinInd, MaxInd
ElseIf mMTID(MidInd) < mCurrMTID Then    'we are out of range on left
   MinInd = MidInd
   FindMTIDRange MinInd, MaxInd
Else                                        'we are inside the range; go
   LInd = MidInd                            'figure out borders
   Do Until LeftDone
      LInd = LInd - 1
      If LInd < MinInd Then
         LeftDone = True
      Else
         If mCurrMTID > mMTID(LInd) Then LeftDone = True
      End If
   Loop
   RInd = MidInd
   Do Until RightDone
      RInd = RInd + 1
      If RInd > MaxInd Then
         RightDone = True
      Else
         If mCurrMTID < mMTID(RInd) Then RightDone = True
      End If
   Loop
   MinInd = LInd + 1
   MaxInd = RInd - 1
End If
End Sub


Private Sub FindRefIDRange(MinInd As Long, MaxInd As Long)
'--------------------------------------------------------------------------------
'recursive procedure to find all elements of array mRefID() that match mCurrRefID
'Note that array mRefID must be in ascending order for this to work
'--------------------------------------------------------------------------------
Dim MidInd As Long
Dim LeftDone As Boolean
Dim RightDone As Boolean
Dim LInd As Long
Dim RInd As Long

MidInd = (MinInd + MaxInd) \ 2
If MidInd = MinInd Then  'Min and Max next to each other
   'this will produce MinInd>MaxInd if none of them match mCurrRefID
   If mCurrRefID <> mRefID(MinInd) Then MinInd = MaxInd
   If mCurrRefID <> mRefID(MaxInd) Then MaxInd = MidInd
   Exit Sub         'we are done
End If
If mRefID(MidInd) > mCurrRefID Then        'we are out of range on right
   MaxInd = MidInd
   FindRefIDRange MinInd, MaxInd
ElseIf mRefID(MidInd) < mCurrRefID Then    'we are out of range on left
   MinInd = MidInd
   FindRefIDRange MinInd, MaxInd
Else                                        'we are inside the range; go
   LInd = MidInd                            'figure out borders
   Do Until LeftDone
      LInd = LInd - 1
      If LInd < MinInd Then
         LeftDone = True
      Else
         If mCurrRefID > mRefID(LInd) Then LeftDone = True
      End If
   Loop
   RInd = MidInd
   Do Until RightDone
      RInd = RInd + 1
      If RInd > MaxInd Then
         RightDone = True
      Else
         If mCurrRefID < mRefID(RInd) Then RightDone = True
      End If
   Loop
   MinInd = LInd + 1
   MaxInd = RInd - 1
End If
End Sub

