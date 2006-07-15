VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'this class implements z-order
'array mInd contains indexes of real structures
'z-order is 0-based
'created: 09/03/2002 nt
'last modified: 09/04/2002 nt
'--------------------------------------------------------------
Option Explicit

Public Enum ErrZOrder
    z_err_NotFound = 1
    z_err_TooManyInZOrder = 2
    z_err_Other = 3
End Enum

Const MAX_Z_ORDER_COUNT = 24            'maximum number of structures in z-order

Dim mZCnt As Long                       'count of structures in z-order
Dim mInd() As Long

Public Function Clear()
mZCnt = 0
Erase mInd
End Function

Public Function ChangeZOrder(ByVal OldZPos As Long, ByVal NewZPos As Long) As Boolean
'------------------------------------------------------------------------------------
'sets index from z-position OldZPos to NewZPos position
'------------------------------------------------------------------------------------
Dim tmp As Long
Dim i As Long
On Error GoTo exit_ChangeZOrder
If NewZPos > OldZPos Then
   tmp = mInd(OldZPos)
   For i = OldZPos + 1 To NewZPos
       mInd(i - 1) = mInd(i)
   Next i
   mInd(NewZPos) = tmp
ElseIf NewZPos < OldZPos Then
   tmp = mInd(OldZPos)
   For i = OldZPos - 1 To NewZPos Step -1
       mInd(i + 1) = mInd(i)
   Next i
   mInd(NewZPos) = tmp
End If
ChangeZOrder = True
exit_ChangeZOrder:
End Function

Public Property Get ZCount() As Long
ZCount = mZCnt
End Property


Public Function AddInZOrder(NewInd As Long, ZPosition As Long) As Boolean
'------------------------------------------------------------------------
'adds new member in z-order and puts it on a specified position; returns
'True if succesful
'------------------------------------------------------------------------
If mZCnt < MAX_Z_ORDER_COUNT Then
   mZCnt = mZCnt + 1
   ReDim Preserve mInd(mZCnt - 1)
   mInd(mZCnt - 1) = NewInd
   If ZPosition < mZCnt - 1 Then
      If ChangeZOrder(mZCnt - 1, ZPosition) Then
         AddInZOrder = True
      Else
         On Error GoTo 0
         Err.Raise z_err_Other + vbObjectError, "Z-Order", "Error in z-order structure."
      End If
   End If
   Exit Function
End If
err_AddInZOrder:
On Error GoTo 0
Err.Raise z_err_TooManyInZOrder + vbObjectError, "Z-Order", "Too many values in z-order."
End Function


Public Function AddInZOrderAll(NewInd() As Long) As Boolean
'----------------------------------------------------------
'adds the whole array to z-order; returns True if succesful
'----------------------------------------------------------
Dim TmpCnt As Long
Dim i As Long
On Error GoTo err_AddInZOrderAll
TmpCnt = UBound(NewInd) - LBound(NewInd) + 1
If TmpCnt > MAX_Z_ORDER_COUNT Then GoTo err_AddInZOrderAll
Call Clear
ReDim mInd(TmpCnt - 1)
For i = 0 To TmpCnt - 1
    mInd(i) = NewInd(LBound(NewInd) + i)
Next i
mZCnt = UBound(mInd) + 1
AddInZOrderAll = True
Exit Function
err_AddInZOrderAll:
On Error GoTo 0
Err.Raise z_err_TooManyInZOrder + vbObjectError, "Z-Order", "Too many values in z-order."
End Function


Public Function IndOnZOrderPosition(ZPosition As Long) As Long
'-------------------------------------------------------------
'returns index on ZPosition
'-------------------------------------------------------------
Dim i As Long
On Error GoTo err_IndOnZOrderPosition
IndOnZOrderPosition = mInd(ZPosition)
Exit Function
err_IndOnZOrderPosition:
On Error GoTo 0
Err.Raise z_err_NotFound + vbObjectError, "Z-Order", "Index not found in z-order."
End Function


Public Function GetZOrder(IndFromZOrder As Long) As Long
'--------------------------------------------------------------------
'returns z-order position of specified Ind; -1 if not found or error
'--------------------------------------------------------------------
Dim i As Long
On Error GoTo err_GetZOrder
For i = 0 To mZCnt - 1
    If mInd(i) = IndFromZOrder Then
       GetZOrder = i
       Exit Function
    End If
Next i
err_GetZOrder:
GetZOrder = -1
End Function

Public Function RemoveFromZOrder(ZPosition As Long) As Boolean
'-------------------------------------------------------------
'removes specified position from z-order
'-------------------------------------------------------------
Dim i As Long
On Error Resume Next
For i = ZPosition To mZCnt - 2
    mInd(i) = mInd(i + 1)
Next i
mZCnt = mZCnt - 1
If mZCnt > 0 Then
   ReDim Preserve mInd(mZCnt - 1)
Else
   Erase mInd
End If
RemoveFromZOrder = True
End Function
