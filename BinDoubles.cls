VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BinDoubles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'this class could be used to calculate bin frequencies and averages
'on array of doubles
'since entry array is sorted it is also handy to get other data
'---------------------------------------------------------------------
'NOTE: array coming in fill is expected to be 0-based array
'---------------------------------------------------------------------
'created: 11/18/2002 nt
'last modified: 02/03/2003 nt
'---------------------------------------------------------------------
Option Explicit

Const ERR_EMPTY = 11
Const ERR_BAD_VALUE = 12

Dim mDummyIndArray() As Long    'never gots initialized

Dim mCnt As Long                'count of array members
Dim mQSDbl() As Double          'sorted array of doubles

Public MinValue As Double
Public MaxValue As Double
Public BinWidth As Double
Public NumFormat As String      'format in which some things will be displayed


Dim mBinCnt As Long             'count of bins

'bin 0 contains values less than MinValue
'bin 1 contains values MinValue<=Value<MinValue+BinWidth
' ...
'bin i contains values MinValue+(i-1)*BinWidth<=Value<MinValue+i*BinWidth
' ...
'bin mBinCnt-1 contains values MaxValue<=Value

Dim mBinInd() As Long           'index of bin; since we will order bins
                                'by number of hits this is essential to
                                'retrieve information about bin
Dim mBinFreq() As Long          'count of values belonging to the bin
Dim mBinSum() As Double         'sum of all values belonging to the bin


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
End If
End Function


Public Property Get Count() As Long
Count = mCnt
End Property


Public Property Get BinsCount() As Long
BinsCount = mBinCnt
End Property


Public Sub CalculateBins()
If CreateBins Then CalculateFrequencies
End Sub

Private Function CreateBins() As Boolean
'----------------------------------------------------------
'create bins array, based on min-max value and type of bins
'----------------------------------------------------------
Dim i As Long
On Error GoTo err_CreateBins
mBinCnt = (MaxValue - MinValue) / BinWidth + 2
ReDim mBinSum(mBinCnt - 1)
ReDim mBinFreq(mBinCnt - 1)
ReDim mBinInd(mBinCnt - 1)
For i = 0 To mBinCnt - 1
    mBinInd(i) = i
Next i
CreateBins = True
Exit Function
err_CreateBins:
Err.Raise vbObjectError + ERR_BAD_VALUE, "BinDoubles"
End Function


Private Function CalculateFrequencies()
Dim i As Long
Dim CurrBin As Long
Dim lQS As New QSLong                   'quick sorter object
On Error GoTo err_CalculateFrequencies

For i = 0 To mCnt - 1
    If mQSDbl(i) < MinValue Then
       CurrBin = 0
    ElseIf mQSDbl(i) >= MaxValue Then
       CurrBin = mBinCnt - 1
    Else
       CurrBin = Int((mQSDbl(i) - MinValue) / BinWidth) + 1
    End If
    mBinFreq(CurrBin) = mBinFreq(CurrBin) + 1
    mBinSum(CurrBin) = mBinSum(CurrBin) + mQSDbl(i)
Next i
'now order bins based on highest scores
CalculateFrequencies = lQS.QSDesc(mBinFreq, mBinInd)
Set lQS = Nothing
Exit Function

err_CalculateFrequencies:
End Function


Public Function GetDataMinValue() As Double
On Error GoTo err_GetDataMinValue
GetDataMinValue = mQSDbl(0)
Exit Function
err_GetDataMinValue:
Err.Raise vbObjectError + ERR_EMPTY, "BinDoubles"
End Function


Public Function GetDataMaxValue() As Double
On Error GoTo err_GetDataMaxValue
GetDataMaxValue = mQSDbl(mCnt - 1)
Exit Function
err_GetDataMaxValue:
Err.Raise vbObjectError + ERR_EMPTY, "BinDoubles"
End Function


Public Function GetBinCount(ByVal FreqOrderInd As Long) As Double
'------------------------------------------------------------------
'returns percent of total count for bin with index FreqOrderInd
'------------------------------------------------------------------
On Error GoTo err_GetBinCount
GetBinCount = mBinFreq(FreqOrderInd)
Exit Function
err_GetBinCount:
Err.Raise vbObjectError + ERR_EMPTY, "BinDoubles"
End Function


Public Function GetBinPercent(ByVal FreqOrderInd As Long) As Double
'------------------------------------------------------------------
'returns percent of total count for bin with index FreqOrderInd
'------------------------------------------------------------------
On Error GoTo err_GetBinPercent
GetBinPercent = mBinFreq(FreqOrderInd) / mCnt
Exit Function
err_GetBinPercent:
Err.Raise vbObjectError + ERR_EMPTY, "BinDoubles"
End Function


Public Function GetBinAverage(ByVal FreqOrderInd As Long) As Double
'------------------------------------------------------------------
'returns average value  for bin with index FreqOrderInd
'------------------------------------------------------------------
On Error GoTo err_GetBinAverage
GetBinAverage = mBinSum(mBinInd(FreqOrderInd)) / mBinFreq(FreqOrderInd)
Exit Function
err_GetBinAverage:
Err.Raise vbObjectError + ERR_EMPTY, "BinDoubles"
End Function


Public Function GetBinRange(ByVal FreqOrderInd As Long) As String
'------------------------------------------------------------------
'returns average value  for bin with index FreqOrderInd
'------------------------------------------------------------------
On Error GoTo err_GetBinRange
Select Case mBinInd(FreqOrderInd)
Case 0
     GetBinRange = "<-," & Format$(MinValue, NumFormat) & ">"
Case mBinCnt - 1
     GetBinRange = "[" & Format$(MaxValue, NumFormat) & ",+>"
Case Else
     GetBinRange = "[" & Format$(MinValue + (mBinInd(FreqOrderInd) - 1) * BinWidth, NumFormat) _
            & "," & Format$(MinValue + mBinInd(FreqOrderInd) * BinWidth, NumFormat) & ">"
End Select
Exit Function
err_GetBinRange:
Err.Raise vbObjectError + ERR_EMPTY, "BinDoubles"
End Function

Private Sub Class_Initialize()
NumFormat = "0.0000"                'just default numeric format
End Sub
