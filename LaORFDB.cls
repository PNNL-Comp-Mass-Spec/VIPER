VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LaORFDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'similar to LaMTDB - content of this class should be ORF
'database information; however it should be made automatically
'for each instance of LaMTDB
'created: 07/25/2002 nt
'last modified: 07/25/2002 nt
'-------------------------------------------------------------
Option Explicit

Const MTDBStartLineToken = "LaORFDB"

Public WithEvents cn As adodb.Connection
Attribute cn.VB_VarHelpID = -1

Public DBStuff As New Collection
Public DBStatus As Long

Private Sub Class_Initialize()
'--------------------------------------
'create new ADO connection
'--------------------------------------
Set cn = New adodb.Connection
DBStatus = dbsNotReady
End Sub

Private Sub Class_Terminate()
'--------------------------------------
'close connection if open
'--------------------------------------
On Error GoTo err_Terminate
If (cn.STATE And adStateClosed) = 0 Then cn.Close
Exit Sub

err_Terminate:
Resume Next
End Sub

Public Function SetDBStuff(NewNames() As String, _
                           NewValues() As String)
'-------------------------------------------------
'destroys old DBStuff collection and creates new
'from referenced arrays
'-------------------------------------------------
Dim NewCnt As Long, I As Long
Dim nv As NameValue
On Error Resume Next
NewCnt = UBound(NewNames) + 1
If NewCnt > 0 Then
   DestroyDBStuff
   Set DBStuff = New Collection
   For I = 1 To NewCnt
       If Len(NewNames(I)) > 0 Then
          Set nv = New NameValue
          nv.Name = NewNames(I - 1)
          nv.Value = NewValues(I - 1)
          DBStuff.add nv, NewNames(I - 1)
       End If
   Next I
End If
End Function


Public Function DestroyDBStuff()
'-------------------------------
'destroys DBStuff collection
'-------------------------------
Dim I As Long
On Error Resume Next
For I = 1 To DBStuff.Count
    DBStuff.Remove 1
Next I
End Function


Public Function WriteDBStuffToFile(ByVal FileName As String) As Boolean
'----------------------------------------------------------------------
'appends content of the DBStuff collection to the file FileName
'----------------------------------------------------------------------
Dim I As Long
Dim fso As New FileSystemObject
Dim ts As TextStream
On Error GoTo err_WriteDBBlaBla
Set ts = fso.OpenTextFile(FileName, ForAppending, False)
ts.WriteLine MTDBStartLineToken & " - Almissa Gloria Victoria Fili Est"
For I = 1 To DBStuff.Count
    ts.WriteLine DBStuff(I).Name & "=" & DBStuff(I).Value
Next I
ts.Close
WriteDBStuffToFile = True
Exit Function

err_WriteDBBlaBla:
End Function


Public Function ReadDBStuffFromFile(ByVal FileName As String) As Boolean
'-----------------------------------------------------------------------
'fills DBStuff collection from specified file; reading starts after
'line containing token is found and everything beyound is considered
'to be stuff that needs to be loaded
'-----------------------------------------------------------------------
Dim I As Long
Dim fso As New FileSystemObject
Dim ts As TextStream
On Error GoTo err_ReadDBBlaBla
'not implemented/used for now
ReadDBStuffFromFile = True
Exit Function

err_ReadDBBlaBla:
End Function


Public Sub EditAddName(ByVal PairName As String, ByVal NewValue As String)
'-------------------------------------------------------------------------
'modifies value of name value pair; if pair does not exist adds it
'-------------------------------------------------------------------------
Dim nv As NameValue
On Error Resume Next
DBStuff.Item(PairName).Value = NewValue
If Err Then
   Set nv = New NameValue
   nv.Name = PairName
   nv.Value = NewValue
   DBStuff.add nv, nv.Name
End If
End Sub


Public Function GetDBStuff() As String
'----------------------------------------------------
'displays all name values found in DBStuff collection
'----------------------------------------------------
Dim nv As NameValue
Dim tmp As String
On Error Resume Next
For Each nv In DBStuff
   tmp = tmp & nv.NameValueStr & vbCrLf
Next
GetDBStuff = tmp
End Function

