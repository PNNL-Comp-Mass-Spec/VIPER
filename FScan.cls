VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'FScan class; used as container of elements of one scan
'NOTE: this is working class; it will contain what you
'put in it
'created: 03/28/2002 nt
'last modified: 03/28/2002 nt
'-----------------------------------------------------
Option Explicit


Public Number As String     'has to be string so that it can be
                            'used as key in a collection
Dim m_Cnt As Long
Dim m_Ind() As Long
Dim m_Type() As Byte
Dim m_Misc() As Double

Public Function AddScanMember(ByVal DataInd As Long, _
                              ByVal DataType As Byte, _
                              Optional DataMisc As Double) As Boolean
'--------------------------------------------------------------------
'adds data point(index+type+misc(optional)) to structure
'--------------------------------------------------------------------
On Error GoTo exit_AddScanMember
'do not accept duplicates
If IsHere(DataInd, DataType) Then Exit Function
m_Cnt = m_Cnt + 1
ReDim Preserve m_Ind(m_Cnt - 1)
ReDim Preserve m_Type(m_Cnt - 1)
ReDim Preserve m_Misc(m_Cnt - 1)
m_Ind(m_Cnt - 1) = DataInd
m_Type(m_Cnt - 1) = DataType
If Not IsMissing(DataMisc) Then m_Misc(m_Cnt - 1) = DataMisc
AddScanMember = True
exit_AddScanMember:
End Function

Public Property Get Count() As Long
Count = m_Cnt
End Property


Public Function GetScanMembers(DataInd() As Long, _
                        DataType() As Byte) As Long
'-----------------------------------------------------
'returns number of scan members, and fills appropriate
'arrays with their indices and types
'-----------------------------------------------------
On Error Resume Next
If m_Cnt > 0 Then
   DataInd = m_Ind
   DataType = m_Type
Else
   Erase DataInd
   Erase DataType
End If
GetScanMembers = m_Cnt
End Function

Public Function GetMiscData(DataMisc() As Double) As Long
'--------------------------------------------------------
'returns number of scan members, and fills appropriate
'arrays with their indices and types
'--------------------------------------------------------
On Error Resume Next
If m_Cnt > 0 Then
   DataMisc = m_Misc
Else
   Erase DataMisc
End If
GetMiscData = m_Cnt
End Function


Private Function IsHere(ByVal DataInd As Long, _
                        ByVal DataType As Byte) As Boolean
'---------------------------------------------------------
'returns True if pair DataInd, DataType is found
'---------------------------------------------------------
Dim i As Long
For i = 0 To m_Cnt - 1
    If m_Ind(i) = DataInd Then
       If m_Type(i) = DataType Then
          IsHere = True
          Exit Function
       End If
    End If
Next i
End Function
