VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LaMTDB"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------
'database class; includes connection and access data
'-----------------------------------------------------------
'last modified: 02/13/2002 nt
'-----------------------------------------------------------
Option Explicit

Const MTDBStartLineToken = "LaMTDB"
'Except  actual  to the database this class provides information
'storage for the mass tag database

'DBStuff collection contains access strings to database views, tables,
'stored procedures; selection criteria for mass tags, to load;
'NOTE: this class does not have to have any knowledge about it's content

Public WithEvents cn As adodb.Connection
Attribute cn.VB_VarHelpID = -1

Dim WithEvents MTSelectionDialog As frmMassTagsSelection
Attribute MTSelectionDialog.VB_VarHelpID = -1

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
If (cn.State And adStateClosed) = 0 Then cn.Close
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
Dim NewCnt As Long, i As Long
Dim nv As NameValue
On Error Resume Next
NewCnt = UBound(NewNames) + 1
If NewCnt > 0 Then
   DestroyDBStuff
   Set DBStuff = New Collection
   For i = 1 To NewCnt
       If Len(NewNames(i)) > 0 Then
          Set nv = New NameValue
          nv.Name = NewNames(i - 1)
          nv.Value = NewValues(i - 1)
          DBStuff.Add nv, NewNames(i - 1)
       End If
   Next i
End If
End Function

Public Function DestroyDBStuff()
'-------------------------------
'destroys DBStuff collection
'-------------------------------
Dim i As Long
On Error Resume Next
For i = 1 To DBStuff.Count
    DBStuff.Remove 1
Next i
End Function

Public Function WriteDBStuffToFile(ByVal FileName As String) As Boolean
'----------------------------------------------------------------------
'appends content of the DBStuff collection to the file FileName
'----------------------------------------------------------------------
Dim i As Long
Dim fso As New FileSystemObject
Dim ts As TextStream
On Error GoTo err_WriteDBBlaBla
Set ts = fso.OpenTextFile(FileName, ForAppending, False)
ts.WriteLine MTDBStartLineToken & " - Almissa Gloria Victoria Fili Est"
For i = 1 To DBStuff.Count
    ts.WriteLine DBStuff(i).Name & "=" & DBStuff(i).Value
Next i
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
Dim i As Long
Dim fso As New FileSystemObject
Dim ts As TextStream
On Error GoTo err_ReadDBBlaBla
ts.Close
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
   DBStuff.Add nv, nv.Name
End If
End Sub

''
''Private Sub LckSelectionDialog_DialogClosed()
'''---------------------------------------------------------
'''pick up new properties if not canceled and destroy dialog
'''---------------------------------------------------------
''Dim nv As NameValue
''On Error Resume Next
''If Not LckSelectionDialog.MyCancel Then
''   DestroyDBStuff
''   For Each nv In LckSelectionDialog.MyStuff
''       DBStuff.Add nv, nv.Name
''   Next
''End If
''Unload LckSelectionDialog
''Set LckSelectionDialog = Nothing
''End Sub

Private Sub MTSelectionDialog_DialogClosed()
'---------------------------------------------------------
'pick up new properties if not canceled and destroy dialog
'---------------------------------------------------------
Dim nv As NameValue
On Error Resume Next
If Not MTSelectionDialog.MyCancel Then
   DestroyDBStuff
   For Each nv In MTSelectionDialog.MyStuff
       DBStuff.Add nv, nv.Name
   Next
End If
Unload MTSelectionDialog
Set MTSelectionDialog = Nothing
End Sub


Public Sub SelectMassTags(InitFileName As String)
'-----------------------------------------
'opens interface to mass tags selection
'-----------------------------------------
Dim nv As NameValue
On Error Resume Next
Set MTSelectionDialog = New frmMassTagsSelection
MTSelectionDialog.MyConnString = cn.ConnectionString
MTSelectionDialog.MTSMasterConnString = GetMTSMasterConnectionString(InitFileName)

Set MTSelectionDialog.MyStuff = New Collection
For Each nv In DBStuff
    MTSelectionDialog.MyStuff.Add nv, nv.Name
Next
MTSelectionDialog.Show
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


Public Sub SelectLockers()
    ' No longer used
End Sub

