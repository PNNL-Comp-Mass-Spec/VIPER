VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QSDoubleWithDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Quick Sort Class for array of Doubles, with corresponding array of Doubles
'Modelled after QSLong.cls

Const qsSORTONE = 0
Const qsSORTPARALLEL = 1
Const qsNOSORT = -1

Dim A() As Double
Dim Ind() As Double

Private LInd As Long
Private HInd As Long
Private SortScope As Long
Private SortError As Boolean

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Source As Any, ByVal ByteCount As Long)

Public Function QSDesc(aL() As Double, IndArr() As Double) As Boolean
'sort is on array aL; while Index array is sorted parallel with it
Dim Cnt As Long
Dim ArrSize As Long     'ArrSize and IndArrSize will be the same for Long
Dim IndArrSize As Long  'and Double types but we might need other types classes
On Error GoTo exit_QSDesc

SortError = False
LInd = LBound(aL)
HInd = UBound(aL)
Cnt = HInd - LInd + 1
Select Case Cnt
Case 0
     MsgBox "No elements of array found.", vbOKOnly
     GoTo exit_QSDesc
Case 1
     'array is already sorted
Case Else
     ReDim A(LInd To HInd)
     ArrSize = Cnt * Len(aL(LInd))
     CopyMemory A(LInd), aL(LInd), ArrSize
     SortScope = SortType(IndArr)
     Select Case SortScope
     Case qsNOSORT
          GoTo exit_QSDesc
     Case qsSORTONE
          SortDescOne LInd, HInd
     Case qsSORTPARALLEL
          ReDim Ind(LInd To HInd)
          IndArrSize = Cnt * Len(IndArr(LInd))
          CopyMemory Ind(LInd), IndArr(LInd), IndArrSize
          SortDesc LInd, HInd
          CopyMemory IndArr(LInd), Ind(LInd), IndArrSize
     End Select
     CopyMemory aL(LInd), A(LInd), ArrSize
End Select
QSDesc = Not SortError
Exit Function

exit_QSDesc:
QSDesc = False
End Function

Public Function QSAsc(aL() As Double, IndArr() As Double) As Boolean
'sort is on array aL; while Index array is sorted parallel with it
Dim Cnt As Long
Dim ArrSize As Long
Dim IndArrSize As Long
On Error GoTo exit_QSAsc
SortError = False
LInd = LBound(aL)
HInd = UBound(aL)
Cnt = HInd - LInd + 1
Select Case Cnt
Case 0
     MsgBox "No elements of array found.", vbOKOnly
     GoTo exit_QSAsc
Case 1
     'array is already sorted
Case Else
     ReDim A(LInd To HInd)
     ArrSize = Cnt * Len(aL(LInd))
     CopyMemory A(LInd), aL(LInd), ArrSize
     SortScope = SortType(IndArr)
     Select Case SortScope
     Case qsNOSORT
          GoTo exit_QSAsc
     Case qsSORTONE
          SortAscOne LInd, HInd
     Case qsSORTPARALLEL
          ReDim Ind(LInd To HInd)
          IndArrSize = Cnt * Len(IndArr(LInd))
          CopyMemory Ind(LInd), IndArr(LInd), IndArrSize
          SortAsc LInd, HInd
          CopyMemory IndArr(LInd), Ind(LInd), IndArrSize
     End Select
     CopyMemory aL(LInd), A(LInd), ArrSize
End Select
QSAsc = Not SortError
Exit Function

exit_QSAsc:
QSAsc = False
End Function

Private Sub SortAsc(ByVal nLow As Long, ByVal nHigh As Long)
Dim I As Long, j As Long
Dim X As Double, Y As Double, dblSwap As Double
On Error GoTo err_SortAsc

I = nLow
j = nHigh
X = A((nLow + nHigh) / 2)
Do While I <= j
   Do While (A(I) < X And I < nHigh)
      I = I + 1
   Loop
   Do While (X < A(j) And j > nLow)
      j = j - 1
   Loop
   If I <= j Then    'swap them; both arrays
      Y = A(I)
      A(I) = A(j)
      A(j) = Y
      
      dblSwap = Ind(I)
      Ind(I) = Ind(j)
      Ind(j) = dblSwap
                  
      I = I + 1
      j = j - 1
   End If
Loop
If nLow < j Then SortAsc nLow, j   'recursions
If I < nHigh Then SortAsc I, nHigh
Exit Sub

err_SortAsc:
    MsgBox "Error: " & Err.Number & "; " & Err.Description, vbOKOnly
    SortError = True
End Sub

Private Sub SortAscOne(ByVal nLow As Long, ByVal nHigh As Long)
Dim I As Long, j As Long
Dim X As Double, Y As Double
On Error GoTo err_SortAscOne

I = nLow
j = nHigh
X = A((nLow + nHigh) / 2)
Do While I <= j
   Do While (A(I) < X And I < nHigh)
      I = I + 1
   Loop
   Do While (X < A(j) And j > nLow)
      j = j - 1
   Loop
   If I <= j Then    'swap them; both arrays
      Y = A(I)
      A(I) = A(j)
      A(j) = Y
                        
      I = I + 1
      j = j - 1
   End If
Loop
If nLow < j Then SortAscOne nLow, j   'recursions
If I < nHigh Then SortAscOne I, nHigh
Exit Sub

err_SortAscOne:
    MsgBox "Error: " & Err.Number & "; " & Err.Description, vbOKOnly
    SortError = True
End Sub


Private Sub SortDesc(ByVal nLow As Long, ByVal nHigh As Long)
'parallel sorting
Dim I As Long, j As Long
Dim X As Double, Y As Double, dblSwap As Double
On Error GoTo err_SortDesc

I = nLow
j = nHigh
X = A((nLow + nHigh) / 2)
Do While I <= j
   Do While (A(I) > X And I < nHigh)
      I = I + 1
   Loop
   Do While (X > A(j) And j > nLow)
      j = j - 1
   Loop
   If I <= j Then    'swap them - both arrays
      Y = A(I)
      A(I) = A(j)
      A(j) = Y
                  
      dblSwap = Ind(I)
      Ind(I) = Ind(j)
      Ind(j) = dblSwap
                  
      I = I + 1
      j = j - 1
   End If
Loop
If nLow < j Then SortDesc nLow, j   'recursions
If I < nHigh Then SortDesc I, nHigh
Exit Sub

err_SortDesc:
    MsgBox "Error: " & Err.Number & "; " & Err.Description, vbOKOnly
    SortError = True
End Sub

Private Sub SortDescOne(ByVal nLow As Long, ByVal nHigh As Long)
Dim I As Long, j As Long
Dim X As Double, Y As Double
On Error GoTo err_SortDescOne

I = nLow
j = nHigh
X = A((nLow + nHigh) / 2)
Do While I <= j
   Do While (A(I) > X And I < nHigh)
      I = I + 1
   Loop
   Do While (X > A(j) And j > nLow)
      j = j - 1
   Loop
   If I <= j Then    'swap them
      Y = A(I)
      A(I) = A(j)
      A(j) = Y
                  
      I = I + 1
      j = j - 1
   End If
Loop
If nLow < j Then SortDescOne nLow, j   'recursions
If I < nHigh Then SortDescOne I, nHigh
Exit Sub

err_SortDescOne:
    MsgBox "Error: " & Err.Number & "; " & Err.Description, vbOKOnly
    SortError = True
End Sub


Private Sub Class_Terminate()
On Error Resume Next
Erase A
If SortScope = qsSORTPARALLEL Then Erase Ind
End Sub

Private Function SortType(IndArr() As Double) As Integer
Dim IndArrL As Long
Dim IndArrH As Long
Dim Respond
On Error Resume Next
IndArrL = LBound(IndArr)
If Err Then
   SortType = qsSORTONE
   Exit Function
End If
IndArrH = UBound(IndArr)
If IndArrL <> LInd Or IndArrH <> HInd Then
   Respond = MsgBox("Arrays indexed differently. Can not perform parallel sort." _
           & " Do you want to sort only on one array?", vbYesNo)
   If Respond <> vbYes Then
      SortType = qsNOSORT
   Else
      SortType = qsSORTONE
   End If
Else
   SortType = qsSORTPARALLEL
End If
End Function


