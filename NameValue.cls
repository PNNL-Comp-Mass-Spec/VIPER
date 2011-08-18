VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NameValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------
'name/value pair class used to create generic
'properties of other objects
'created: 05/17/2001 nt
'last modified: 05/17/2001 nt
'---------------------------------------------------
Public Name As String
Public Value As String

Public Function NameValueStr() As String
'---------------------------------------------
'returns Name, Value pair as string Name=Value
'---------------------------------------------
Dim tmp As String
On Error Resume Next
tmp = Name & "="
NameValueStr = tmp & Value
End Function
