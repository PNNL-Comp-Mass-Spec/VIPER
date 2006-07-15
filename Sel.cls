VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Selection object - contains arrays with indexes
'in CSData and IsoData arrays and supporting functions;

'last modified 09/15/2000 nt
Option Explicit

'arrays are ordered for faster utilities
'there should be no duplicates in these arrays
Dim mCSSel() As Long
Dim mISSel() As Long

Dim mCSSelCnt As Long
Dim mISSelCnt As Long

Public Event ChangeCSSel()   'raised whenever CS selection changes
Public Event ChangeIsoSel()   'raised whenever Iso selection changes

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Source As Any, ByVal ByteCount As Long)

Public Sub Clear()
Erase mCSSel
Erase mISSel
mCSSelCnt = 0
mISSelCnt = 0
RaiseEvent ChangeCSSel
RaiseEvent ChangeIsoSel
End Sub

Public Property Get CSSelCnt() As Long
'read-only property - count of CS selection
CSSelCnt = mCSSelCnt
End Property

Public Property Get IsoSelCnt() As Long
'read-only property - count of Iso selection
IsoSelCnt = mISSelCnt
End Property


Public Sub AddToCSSelection(ByVal CSIndex As Long)
Dim tmp() As Long
Dim TmpArrSize As Long
Dim InsPos As Long
'if this index is already selected do nothing
If FindCSInd(CSIndex) >= 0 Then Exit Sub
If mCSSelCnt > 0 Then                   'at least one already in the array
   InsPos = FindPlaceCS(CSIndex)
   Select Case InsPos
   Case 0          'insert at the begining
     ReDim tmp(1 To mCSSelCnt)
     TmpArrSize = Len(mCSSel(1)) * mCSSelCnt
     CopyMemory tmp(1), mCSSel(1), TmpArrSize
     mCSSelCnt = mCSSelCnt + 1
     ReDim mCSSel(1 To mCSSelCnt)
     CopyMemory mCSSel(2), tmp(1), TmpArrSize
     mCSSel(1) = CSIndex
   Case mCSSelCnt + 1 'insert at the end
     mCSSelCnt = mCSSelCnt + 1
     ReDim Preserve mCSSel(1 To mCSSelCnt)
     mCSSel(mCSSelCnt) = CSIndex
   Case Else       'somwhere in the middle
     ReDim tmp(1 To mCSSelCnt - InsPos + 1)
     TmpArrSize = (mCSSelCnt - InsPos + 1) * Len(mCSSel(1))
     CopyMemory tmp(1), mCSSel(InsPos), TmpArrSize
     mCSSelCnt = mCSSelCnt + 1
     ReDim Preserve mCSSel(1 To mCSSelCnt)
     mCSSel(InsPos) = CSIndex
     CopyMemory mCSSel(InsPos + 1), tmp(1), TmpArrSize
   End Select
Else                                    'first element
   ReDim mCSSel(1 To 1)
   mCSSel(1) = CSIndex
   mCSSelCnt = 1
End If
RaiseEvent ChangeCSSel
End Sub

Public Sub AddToIsoSelection(ByVal IsoIndex As Long)
Dim tmp() As Long
Dim TmpArrSize As Long
Dim InsPos As Long
'if this index is already selected do nothing
If FindIsoInd(IsoIndex) >= 0 Then Exit Sub
If mISSelCnt > 0 Then                   'at least one already in the array
   InsPos = FindPlaceIS(IsoIndex)
   Select Case InsPos
   Case 0          'insert at the begining
     ReDim tmp(1 To mISSelCnt)
     TmpArrSize = Len(mISSel(1)) * mISSelCnt
     CopyMemory tmp(1), mISSel(1), TmpArrSize
     mISSelCnt = mISSelCnt + 1
     ReDim mISSel(1 To mISSelCnt)
     CopyMemory mISSel(2), tmp(1), TmpArrSize
     mISSel(1) = IsoIndex
   Case mISSelCnt + 1 'insert at the end
     mISSelCnt = mISSelCnt + 1
     ReDim Preserve mISSel(1 To mISSelCnt)
     mISSel(mISSelCnt) = IsoIndex
   Case Else       'somwhere in the middle
     ReDim tmp(1 To mISSelCnt - InsPos + 1)
     TmpArrSize = (mISSelCnt - InsPos + 1) * Len(mISSel(1))
     CopyMemory tmp(1), mISSel(InsPos), TmpArrSize
     mISSelCnt = mISSelCnt + 1
     ReDim Preserve mISSel(1 To mISSelCnt)
     mISSel(InsPos) = IsoIndex
     CopyMemory mISSel(InsPos + 1), tmp(1), TmpArrSize
   End Select
Else                                    'first element
   ReDim mISSel(1 To 1)
   mISSel(1) = IsoIndex
   mISSelCnt = 1
End If
RaiseEvent ChangeIsoSel
End Sub


Private Function FindPlaceCS(ByVal CS As Long) As Long
'returns position where new element of the CSSel
'array should be inserted
Dim SInd As Long
Dim EInd As Long
Dim mInd As Long
Dim Done As Boolean

SInd = 1
EInd = mCSSelCnt
If CS < mCSSel(1) Then
   FindPlaceCS = 0
ElseIf CS > mCSSel(mCSSelCnt) Then
   FindPlaceCS = mCSSelCnt + 1
Else
  Do Until Done
     mInd = (SInd + EInd) \ 2
     If mInd = SInd Then 'SInd and EInd next to each other
        If CS > mCSSel(mInd) Then mInd = mInd + 1
        Done = True
     Else
        If mCSSel(mInd) > CS Then
           EInd = mInd
        ElseIf mCSSel(mInd) < CS Then
           SInd = mInd
        Else
           Done = True
        End If
     End If
  Loop
  FindPlaceCS = mInd
End If
End Function

Private Function FindPlaceIS(ByVal Iso As Long) As Long
'returns position where new element of the IsoSel
'array should be inserted
Dim SInd As Long
Dim EInd As Long
Dim mInd As Long
Dim Done As Boolean

SInd = 1
EInd = mISSelCnt
If Iso < mISSel(1) Then
   FindPlaceIS = 0
ElseIf Iso > mISSel(mISSelCnt) Then
   FindPlaceIS = mISSelCnt + 1
Else
  Do Until Done
     mInd = (SInd + EInd) \ 2
     If mInd = SInd Then 'SInd and EInd next to each other
        If Iso > mISSel(mInd) Then mInd = mInd + 1
        Done = True
     Else
        If mISSel(mInd) > Iso Then
           EInd = mInd
        ElseIf mISSel(mInd) < Iso Then
           SInd = mInd
        Else
           Done = True
        End If
     End If
  Loop
  FindPlaceIS = mInd
End If
End Function


Public Property Get value(ByVal Ind As Long, _
                          ByVal SelType As Integer) As Long
'returns value in Ind place of CSSel or IsoSel arrays; -1 on error
On Error Resume Next
value = -1
Select Case SelType
Case 1
     value = mCSSel(Ind)
Case 2
     value = mISSel(Ind)
End Select
End Property

Public Sub DelFromCSSelection(ByVal CSInd As Long)
Dim aTmp() As Long
Dim TmpArrSize As Long
Dim DelPos As Long

If mCSSelCnt > 0 Then
   DelPos = FindCSInd(CSInd)
   If DelPos > 0 Then
      If mCSSelCnt > 1 Then
         Select Case DelPos
         Case 1            'delete first element
           ReDim aTmp(1 To mCSSelCnt - 1)
           TmpArrSize = (mCSSelCnt - 1) * Len(mCSSel(1))
           CopyMemory aTmp(1), mCSSel(2), TmpArrSize
           ReDim mCSSel(1 To mCSSelCnt - 1)
           CopyMemory mCSSel(1), aTmp(1), TmpArrSize
         Case mCSSelCnt    'delete last element
           ReDim aTmp(1 To mCSSelCnt - 1)
           TmpArrSize = (mCSSelCnt - 1) * Len(mCSSel(1))
           CopyMemory aTmp(1), mCSSel(1), TmpArrSize
           ReDim mCSSel(1 To mCSSelCnt - 1)
           CopyMemory mCSSel(1), aTmp(1), TmpArrSize
         Case Else         'delete element in the middle
           ReDim aTmp(1 To mCSSelCnt - DelPos)
           TmpArrSize = (mCSSelCnt - DelPos) * Len(mCSSel(1))
           CopyMemory aTmp(1), mCSSel(DelPos + 1), TmpArrSize
           ReDim Preserve mCSSel(1 To mCSSelCnt - 1)
           CopyMemory mCSSel(DelPos), aTmp(1), TmpArrSize
         End Select
      Else  'delete last element in array; just erase it
         Erase mCSSel
      End If
      mCSSelCnt = mCSSelCnt - 1
   End If
End If
RaiseEvent ChangeCSSel
End Sub

Public Sub DelFromIsoSelection(ByVal IsoInd As Long)
Dim aTmp() As Long
Dim TmpArrSize As Long
Dim DelPos As Long

If mISSelCnt > 0 Then
   DelPos = FindIsoInd(IsoInd)
   If DelPos > 0 Then
      If mISSelCnt > 1 Then
         Select Case DelPos
         Case 1            'delete first element
           ReDim aTmp(1 To mISSelCnt - 1)
           TmpArrSize = (mISSelCnt - 1) * Len(mISSel(1))
           CopyMemory aTmp(1), mISSel(2), TmpArrSize
           ReDim mISSel(1 To mISSelCnt - 1)
           CopyMemory mISSel(1), aTmp(1), TmpArrSize
         Case mISSelCnt    'delete last element
           ReDim aTmp(1 To mISSelCnt - 1)
           TmpArrSize = (mISSelCnt - 1) * Len(mISSel(1))
           CopyMemory aTmp(1), mISSel(1), TmpArrSize
           ReDim mISSel(1 To mISSelCnt - 1)
           CopyMemory mISSel(1), aTmp(1), TmpArrSize
         Case Else         'delete element in the middle
           ReDim aTmp(1 To mISSelCnt - DelPos)
           TmpArrSize = (mISSelCnt - DelPos) * Len(mISSel(1))
           CopyMemory aTmp(1), mISSel(DelPos + 1), TmpArrSize
           ReDim Preserve mISSel(1 To mISSelCnt - 1)
           CopyMemory mISSel(DelPos), aTmp(1), TmpArrSize
         End Select
      Else  'delete last element in array; just erase it
         Erase mCSSel
      End If
      mISSelCnt = mISSelCnt - 1
   End If
End If
RaiseEvent ChangeIsoSel
End Sub

Public Function FindCSInd(ByVal CS As Long) As Long
'returns index in mCSSel of CS if found; -1 if not
Dim SInd As Long
Dim EInd As Long
Dim mInd As Long
Dim Done As Boolean
On Error Resume Next
FindCSInd = -1
'if none selected no need to go further
If mCSSelCnt <= 0 Then Exit Function
SInd = 1
EInd = mCSSelCnt
'if out of range no need to go further
If ((CS < mCSSel(1)) Or (CS > mCSSel(mCSSelCnt))) Then Exit Function
'try to locate it
Do Until Done
   If CS = mCSSel(SInd) Then
      FindCSInd = SInd
      Done = True
   ElseIf CS = mCSSel(EInd) Then
      FindCSInd = EInd
      Done = True
   Else
      mInd = (SInd + EInd) \ 2
      If mInd = SInd Then 'SInd and EInd next to each other
         Done = True      'and it was already checked if any
      Else
         If mCSSel(mInd) = CS Then
            FindCSInd = mInd
            Done = True
         ElseIf mCSSel(mInd) > CS Then
            EInd = mInd
         Else
            SInd = mInd
         End If
      End If
   End If
Loop
End Function

Public Function FindIsoInd(ByVal Iso As Long) As Long
'returns index in mISSel of Iso if found; -1 if not
Dim SInd As Long
Dim EInd As Long
Dim mInd As Long
Dim Done As Boolean
On Error Resume Next
FindIsoInd = -1
'if no selection no need to go further
If mISSelCnt <= 0 Then Exit Function
SInd = 1
EInd = mISSelCnt
'if out of current range no need to go further
If ((Iso < mISSel(1)) Or (Iso > mISSel(mISSelCnt))) Then Exit Function
'try to find it
Do Until Done
   If Iso = mISSel(SInd) Then
      FindIsoInd = SInd
      Done = True
   ElseIf Iso = mISSel(EInd) Then
      FindIsoInd = EInd
      Done = True
   Else
      mInd = (SInd + EInd) \ 2
      If mInd = SInd Then 'SInd and EInd next to each other
         Done = True      'and it was already checked if any
      Else
         If mISSel(mInd) = Iso Then
            FindIsoInd = mInd
            Done = True
         ElseIf mISSel(mInd) > Iso Then
            EInd = mInd
         Else
            SInd = mInd
         End If
      End If
   End If
Loop
End Function

Public Sub GetCSSel(CSSel() As Long)
'copies array mCSSel to array CSSel
Dim ArrSize As Long
If mCSSelCnt > 0 Then
   ReDim CSSel(1 To mCSSelCnt)
   ArrSize = mCSSelCnt * Len(mCSSel(1))
   CopyMemory CSSel(1), mCSSel(1), ArrSize
End If
End Sub

Public Sub GetIsoSel(IsoSel() As Long)
'copies array mISSel to array IsoSel
Dim ArrSize As Long
If mISSelCnt > 0 Then
   ReDim IsoSel(1 To mISSelCnt)
   ArrSize = mISSelCnt * Len(mISSel(1))
   CopyMemory IsoSel(1), mISSel(1), ArrSize
End If
End Sub
