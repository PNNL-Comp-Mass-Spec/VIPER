VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StatDoubles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'this function is used to calculate various statistics on an array
'of doubles; this is an overkill for some functions that not need sort
'Functions: Mean, Median, Quartile, Percentile, St.Dev.
'---------------------------------------------------------------------
'NOTE: array coming in fill is expected to be 0-based array
'---------------------------------------------------------------------
'created: 11/18/2002 nt
'last modified: 12/03/2002 nt
'extended to include minimum and maximum: 7/16/2003 by mem
'---------------------------------------------------------------------
Option Explicit

Const ERR_EMPTY = 27
Const ERR_BAD_VALUE = 28

Public Enum sdQuarts
    sdQuart0 = 0            'minimum element
    sdQuart1 = 1            '25 percentile
    sdQuart2 = 2            'median
    sdQuart3 = 3            '75 percentile
    sdQuart4 = 4            'maximum element
End Enum

Dim mDummyIndArray() As Long    'never gots initialized

Dim mCnt As Long                'count of array members
Dim mQSDbl() As Double          'sorted array of doubles            ' 0-based array

Public Function Median() As Double
'-----------------------------------------------------------------------
'NOTE:if the number of members is even then median is the average of two
'members in the middle
'-----------------------------------------------------------------------
Dim HalfInd As Long
If mCnt > 0 Then
   HalfInd = Int(mCnt / 2)
   If mCnt Mod 2 > 0 Then               'odd membership
      Median = mQSDbl(HalfInd)
   Else                                 'even membership
      Median = (mQSDbl(HalfInd - 1) + mQSDbl(HalfInd)) / 2
   End If
Else
   Err.Raise Number:=vbObjectError + ERR_EMPTY
End If
End Function


Public Function Mean() As Double
If mCnt > 0 Then
   Mean = Sum / mCnt
Else
   Err.Raise Number:=vbObjectError + ERR_EMPTY
End If
End Function

Public Function Maximum() As Double
Dim i As Long
Dim dblMaximum As Double
If mCnt > 0 Then
   dblMaximum = mQSDbl(0)
   For i = 1 To mCnt - 1
       If mQSDbl(i) > dblMaximum Then
          dblMaximum = mQSDbl(i)
       End If
   Next i
   Maximum = dblMaximum
Else
   Err.Raise Number:=vbObjectError + ERR_EMPTY
End If
End Function

Public Function Minimum() As Double
Dim i As Long
Dim dblMinimum As Double
If mCnt > 0 Then
   dblMinimum = mQSDbl(0)
   For i = 1 To mCnt - 1
       If mQSDbl(i) < dblMinimum Then
          dblMinimum = mQSDbl(i)
       End If
   Next i
   Minimum = dblMinimum
Else
   Err.Raise Number:=vbObjectError + ERR_EMPTY
End If
End Function

Public Function Mode() As Double
If mCnt > 0 Then
    ' For lists, the mode is the most common (frequent) value. A list can have more than one mode.
    ' This has not been coded
    Debug.Assert False
Else
   Err.Raise Number:=vbObjectError + ERR_EMPTY
End If
End Function

Public Function Sum() As Double
Dim i As Long
If mCnt > 0 Then
   For i = 0 To mCnt - 1
       Sum = Sum + mQSDbl(i)
   Next i
Else
   Err.Raise Number:=vbObjectError + ERR_EMPTY
End If
End Function


Public Function Sum2() As Double
Dim i As Long
If mCnt > 0 Then
   For i = 0 To mCnt - 1
       Sum2 = Sum2 + mQSDbl(i) * mQSDbl(i)
   Next i
Else
   Err.Raise Number:=vbObjectError + ERR_EMPTY
End If
End Function


Public Function Fill(DblArr() As Double) As Boolean
'-------------------------------------------------------------------------
'returns True if array is successfully sorted and has at least one element
'-------------------------------------------------------------------------
Dim qsDbl As New QSDouble
On Error Resume Next
mCnt = -1
mQSDbl = DblArr
If qsDbl.QSAsc(mQSDbl(), mDummyIndArray()) Then
   mCnt = UBound(mQSDbl) + 1
   If mCnt > 0 Then Fill = True
Else
   mCnt = -1
End If
End Function

Public Property Get Count() As Long
Count = mCnt
End Property


Public Function Quartile(Quart As sdQuarts) As Double
Select Case Quart
Case sdQuart0
     If mCnt > 0 Then
        Quartile = mQSDbl(0)
     Else
        Err.Raise Number:=vbObjectError + ERR_EMPTY
     End If
Case sdQuart1
     Quartile = Percentile(0.25)
Case sdQuart2
     Quartile = Median
Case sdQuart3
     Quartile = Percentile(0.75)
Case sdQuart4
     If mCnt > 0 Then
        Quartile = mQSDbl(mCnt - 1)
     Else
        Err.Raise Number:=vbObjectError + ERR_EMPTY
     End If
End Select
End Function

Public Function Percentile(Pct As Double) As Double
'-----------------------------------------------------------------------
'NOTE: we can probably do better interpolation but for practical purpose
'this is good enough
'-----------------------------------------------------------------------
Dim PctInd As Long
If (0 < Pct) And (Pct < 1) Then
   If mCnt > 0 Then
      PctInd = CLng(mCnt * Pct)
      If PctInd < mCnt * Pct Then
         If PctInd < mCnt - 1 Then
            Percentile = (mQSDbl(PctInd) + mQSDbl(PctInd + 1)) / 2
         Else
            Percentile = mQSDbl(PctInd)
         End If
      Else
         Percentile = mQSDbl(PctInd)
      End If
   Else
      Err.Raise Number:=vbObjectError + ERR_EMPTY
   End If
Else
   Err.Raise Number:=vbObjectError + ERR_BAD_VALUE
End If
End Function

Public Function StDev() As Double
'-----------------------------------------------------------------------
'returns standard deviation(nonbiased) of array of numbers(doubles);
'-1 if not applicable or on any error
'-----------------------------------------------------------------------
Dim SumOne As Double
Dim SumTwo As Double
On Error Resume Next
If mCnt > 1 Then
   SumOne = Sum()
   SumTwo = Sum2()
   StDev = Sqr(SumTwo / (mCnt - 1) - (SumOne * SumOne) / (mCnt * (mCnt - 1)))
ElseIf mCnt = 1 Then
   StDev = 0
Else
   StDev = -1
End If
End Function
