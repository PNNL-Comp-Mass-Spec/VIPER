VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MTDBInfoRetriever"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private MyInit As InitFile         'object that deals with initialization file

Private mInitFilePath As String       'full path to init file
Public fAnalysis As FTICRAnalysis    'FTICR Analysis object

Public Sub GetMTDBSchema()
'-----------------------------------------------------
'retrieves schema information from initialization file
'this is same for all Mass Tags databases
'also enables user to select subset of all mass tags
'to work with
'-----------------------------------------------------
Dim Arg() As String
Dim ArgCnt As Long
Dim MTName As String, MTValue As String
Dim MTValuePos As Long
Dim nv As NameValue
Dim i As Long
On Error Resume Next

ArgCnt = MyInit.GetSection(mInitFilePath, MyGl.SECTION_MTDB_Schema, Arg())
If ArgCnt > 0 Then
   For i = 0 To ArgCnt - 1
       MTValuePos = InStr(1, Arg(i), MyGl.INIT_Value)
       If MTValuePos > 0 Then
          MTName = Trim(Left$(Arg(i), MTValuePos - 1))
          MTValue = Trim$(Right$(Arg(i), Len(Arg(i)) - MTValuePos))
       Else     'everything is a name
          MTName = Trim(Arg(i))
          MTValue = ""
       End If
       If Len(MTName) > 0 Then
          Set nv = New NameValue
          nv.Name = MTName
          nv.Value = MTValue
          fAnalysis.MTDB.DBStuff.Add nv, nv.Name
       End If   'do nothing if name is missing
   Next i
End If
End Sub

Public Property Get InitFilePath() As String
    InitFilePath = mInitFilePath
End Property

Public Property Let InitFilePath(strInitFilePath As String)
    mInitFilePath = strInitFilePath
End Property

Private Sub Class_Initialize()
    Set MyInit = New InitFile
    Set fAnalysis = New FTICRAnalysis
End Sub

Private Sub Class_Terminate()
    Set MyInit = Nothing
    Set fAnalysis = Nothing

End Sub
