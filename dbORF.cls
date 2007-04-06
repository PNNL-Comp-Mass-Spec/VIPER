VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dbORF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'encapsulation of ORF information used to search for intact proteins
'created: 10/04/2002 nt
'last modified: 10/04/2002 nt
'-------------------------------------------------------------------
Option Explicit

Public Enum ORFObjectType
    ORFObjNone = -1
    ORFObjORFDB = 0
    ORFObjFASTA = 1
    ORFObjLegacy = 2
End Enum


Dim mORFObjType As ORFObjectType
Dim mORFObjConnStr As String
Dim mORFObjspName As String
Dim mORFObjFASTAFileName As String

Const MY_NAME = "ORF Object"

Dim mORFCnt As Long                     'count
Dim mORFID() As Long                    'database ID
Dim mORFReference() As String           'reference
Dim mORFDescription() As String         'description
Dim mORFMW() As Double                  'MW
Dim mORFMF() As String                  'molecular formula C2453 H5678 N234 O1000 S21
Dim mORFAA() As String                  'amino acid composition A24 E35 etc.


Public Function LoadORFsFromFASTA(FileName As String) As Boolean
'--------------------------------------------------------------------------------
'loads ORFs from FASTA file and fills data structures; returns True if successful
'--------------------------------------------------------------------------------

End Function


Public Function LoadORFsFromORFDB(DBConnStr As String, spGetORFs As String) As Boolean
'-------------------------------------------------------------------------------------
'loads Proteins (ORFs) from Protein database and fills data structures; returns True if successful
'IniFileName is neccessary to retrieve name of stored procedure to retrieve Proteins;
'-------------------------------------------------------------------------------------
Dim cnNew As New ADODB.Connection
Dim sCommand As String
Dim rsORFs As New ADODB.Recordset
Dim cmdGetORFs As New ADODB.Command
'reserve space for 10000 Proteins; increase in chunks of 2000 after that
ReDim mORFID(10000)
ReDim mORFReference(10000)
ReDim mORFDescription(10000)
ReDim mORFMW(10000)
ReDim mORFMW(10000)
ReDim mORFAA(10000)
On Error GoTo err_LoadORFSFromORFDB
If Len(spGetORFs) <= 0 Then Exit Function
Screen.MousePointer = vbHourglass
mORFObjConnStr = DBConnStr
mORFObjspName = spGetORFs
mORFCnt = 0

If Not EstablishConnection(cnNew, DBConnStr, False) Then
    Debug.Assert False
    LoadORFsFromORFDB = False
    Exit Function
End If

'create and tune command object to retrieve MT tags
' Initialize the SP
InitializeSPCommand cmdGetORFs, cnNew, spGetORFs

'procedure returns error number or 0 if OK
Set rsORFs = cmdGetORFs.Execute
With rsORFs
    Do Until .EOF
       mORFCnt = mORFCnt + 1
       mORFID(mORFCnt - 1) = .Fields(0).Value
       mORFReference(mORFCnt - 1) = .Fields(1).Value
       mORFDescription(mORFCnt - 1) = .Fields(2).Value
       mORFMW(mORFCnt - 1) = .Fields(3).Value
       mORFMF(mORFCnt - 1) = .Fields(4).Value
       mORFAA(mORFCnt - 1) = .Fields(5).Value
       .MoveNext
    Loop
End With
rsORFs.Close
'clean things and exit
exit_LoadORFsFromORFDB:
On Error Resume Next
Set cmdGetORFs.ActiveConnection = Nothing
cnNew.Close
If mORFCnt > 0 Then
   If mORFCnt - 1 < UBound(AMTData) Then
      ReDim Preserve mORFID(mORFCnt - 1)
      ReDim Preserve mORFReference(mORFCnt - 1)
      ReDim Preserve mORFDescription(mORFCnt - 1)
      ReDim Preserve mORFMW(mORFCnt - 1)
      ReDim Preserve mORFMF(mORFCnt - 1)
      ReDim Preserve mORFAA(mORFCnt - 1)
   End If
Else
   Erase mORFID
   Erase mORFReference
   Erase mORFDescription
   Erase mORFMW
   Erase mORFMF
   Erase mORFAA
End If
Screen.MousePointer = vbDefault
LoadORFsFromORFDB = True
mORFObjType = ORFObjORFDB
Exit Function

err_LoadORFSFromORFDB:
Select Case Err.Number
Case 9                       'need more room for MT tags
    ReDim Preserve mORFID(mORFCnt + 2000)
    ReDim Preserve mORFReference(mORFCnt + 2000)
    ReDim Preserve mORFDescription(mORFCnt + 2000)
    ReDim Preserve mORFMW(mORFCnt + 2000)
    ReDim Preserve mORFMF(mORFCnt + 2000)
    ReDim Preserve mORFAA(mORFCnt + 2000)
    Resume
Case 13, 94                  'Type Mismatch or Invalid Use of Null
    Resume Next              'just ignore it
Case 3265, 3704              'two errors I have encountered
    '2nd attempt will probably work so let user know it should try again
    MsgBox "Error loading MT tags from the database. Error could " _
         & "have been caused by network/server issues(timeout) so you " _
         & "might try loading again.", vbOKOnly, MY_NAME
Case Else                    'try to ignore
    Resume Next
End Select
mORFCnt = -1
mORFObjType = ORFObjNone
GoTo exit_LoadORFsFromORFDB
End Function


Public Function LoadORFsFromLegacyDB() As Boolean
'---------------------------------------------------------------------------------------
'loads ORFs from legacy AMT database and fills data structures; returns True if successful
'IniFileName is neccessary to retrieve name of stored procedure to retrieve Proteins;
'---------------------------------------------------------------------------------------

End Function


Public Function ORFSearchMW(ByVal MW As Double, ByVal MWTol As Double, ResID() As Long) As Long
'----------------------------------------------------------------------------------------------
'fills array ResID with IDs of Proteins matching requirements and returns count; -1 on any error
'----------------------------------------------------------------------------------------------
On Error GoTo err_ORFSearchMW


Exit Function

err_ORFSearchMW:
ORFSearchMW = -1
End Function

Public Function ORFDescriptionByID(ByVal ORFID As Long) As String

End Function

Public Function ORFDescriptionByRef(ByVal ORFReference As String) As String

End Function


Private Sub Class_Initialize()
mORFObjType = ORFObjNone
End Sub


'public properties

Public Property Get ORFCount() As Long
ORFCount = mORFCnt
End Property

Public Property Get ORFObjType() As ORFObjectType
ORFObjType = mORFObjType
End Property

Public Property Let ORFObjType(ByVal NewORFObjType As ORFObjectType)
mORFObjType = NewORFObjType
End Property

Public Property Get ORFObjConnStr() As String
ORFObjConnStr = mORFObjConnStr
End Property

Public Property Let ORFObjConnStr(ByVal NewConnStr As String)
mORFObjConnStr = NewConnStr
End Property

Public Property Get ORFObjspName() As String
ORFObjspName = mORFObjspName
End Property

Public Property Let ORFObjspName(ByVal NewORFObjspName As String)
mORFObjspName = NewORFObjspName
End Property

Public Property Get ORFObjFASTAFileName() As String
ORFObjFASTAFileName = mORFObjFASTAFileName
End Property

Public Property Let ORFObjFASTAFileName(ByVal NewFASTAFileName As String)
mORFObjFASTAFileName = NewFASTAFileName
End Property
