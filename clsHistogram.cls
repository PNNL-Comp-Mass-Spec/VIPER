VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHistogram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This class can be used to bin data into a histogram
' Written by Matthew Monroe
' Started February 14, 2005

Private mBinRangeMaximum As Single

Private mStartBin As Single                          ' Guaranteed to be <= 0
Private mStartBinDigitsAfterDecimal As Integer       ' Number of digits after decimal to round the start bin to
    
Private mBinCountMaximum As Long

Private mBinSize As Single
Private mDefaultBinSize As Single

Private mRequireNegativeStartBin As Boolean
Private mShowMessages As Boolean
Private mErrorNumber As Long

Public Property Get BinCountMaximum() As Long
     BinCountMaximum = mBinCountMaximum
End Property

Public Property Get BinSize() As Single
    BinSize = mBinSize
End Property
Public Property Let BinSize(sngValue As Single)
    mBinSize = sngValue
End Property

Public Property Get BinRangeMaximum() As Single
    BinRangeMaximum = mBinRangeMaximum
End Property

Public Property Get DefaultBinSize() As Single
    DefaultBinSize = mDefaultBinSize
End Property
Public Property Let DefaultBinSize(sngValue As Single)
    mDefaultBinSize = sngValue
End Property

Public Property Get ErrorNumber() As Long
    ErrorNumber = mErrorNumber
End Property

Public Property Get RequireNegativeStartBin() As Boolean
    RequireNegativeStartBin = mRequireNegativeStartBin
End Property
Public Property Let RequireNegativeStartBin(blnValue As Boolean)
    mRequireNegativeStartBin = blnValue
End Property

Public Property Get ShowMessages() As Boolean
    ShowMessages = mShowMessages
End Property
Public Property Let ShowMessages(blnValue As Boolean)
    mShowMessages = blnValue
End Property

Public Property Get StartBin() As Single
    StartBin = mStartBin
End Property
Public Property Let StartBin(sngValue As Single)
    mStartBin = sngValue
End Property

Public Property Get StartBinDigitsAfterDecimal() As Single
    StartBinDigitsAfterDecimal = mStartBinDigitsAfterDecimal
End Property
Public Property Let StartBinDigitsAfterDecimal(sngValue As Single)
    mStartBinDigitsAfterDecimal = sngValue
End Property

Public Function BinData(ByRef sngDataToBin() As Single, ByRef lngDataCount As Long, ByRef BinnedData() As Long, ByRef BinnedDataCount As Long) As Boolean
    ' sngDataToBin() should be a 0-based array
    ' Note that BinnedData ranges from index 0 to index BinnedDataCount

    Dim sngBinRangeMinimum As Single
    Dim sngBinRangeMaximum As Single
    Dim blnSuccess As Boolean
    
    Dim lngIndex As Long, lngBinNumber As Long

On Error GoTo BinErrorsErrorHandler

    mErrorNumber = 0

    If lngDataCount <= 0 Then
        BinnedDataCount = 0
        InitializeBinnedData BinnedDataCount, BinnedData, True
        BinData = False
        Exit Function
    End If

    ' Bin the data in sngDataToBin
    ' Need to know the minimum and maximum data values
    sngBinRangeMinimum = sngDataToBin(0)
    sngBinRangeMaximum = sngBinRangeMinimum
    For lngIndex = 1 To lngDataCount - 1
        If sngDataToBin(lngIndex) > sngBinRangeMaximum Then sngBinRangeMaximum = sngDataToBin(lngIndex)
        If sngDataToBin(lngIndex) < sngBinRangeMinimum Then sngBinRangeMinimum = sngDataToBin(lngIndex)
    Next lngIndex

    ' Call ComputeBinDetails to round the bin ranges to the appropriate values and update the bin size, if necessary
    blnSuccess = ComputeBinDetails(sngBinRangeMinimum, sngBinRangeMaximum, mBinSize, BinnedDataCount, RequireNegativeStartBin, StartBinDigitsAfterDecimal)
    
    If Not blnSuccess Then
        BinnedDataCount = 0
        InitializeBinnedData BinnedDataCount, BinnedData, True
        BinData = False
        Exit Function
    End If

    On Error GoTo BinErrorsErrorHandler

    mStartBin = sngBinRangeMinimum
    mBinRangeMaximum = sngBinRangeMaximum
    
    InitializeBinnedData BinnedDataCount, BinnedData, False

    For lngIndex = 0 To lngDataCount - 1
        lngBinNumber = DataPointToBinNumber(sngDataToBin(lngIndex), mStartBin, mBinSize)

        If lngBinNumber >= 0 And lngBinNumber <= BinnedDataCount Then
            BinnedData(lngBinNumber) = BinnedData(lngBinNumber) + 1
        Else
            Debug.Assert False
        End If
    Next lngIndex

    ' Now that the data has been binned, determine the maximum bin value
    mBinCountMaximum = 0
    For lngIndex = 0 To BinnedDataCount
        If BinnedData(lngIndex) > mBinCountMaximum Then
            mBinCountMaximum = BinnedData(lngIndex)
        End If
    Next lngIndex
    
    BinData = True
    Exit Function

BinErrorsErrorHandler:
    If mShowMessages Then
        MsgBox "Unexpected error in BinErrors()" & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    mErrorNumber = Err.Number
    BinnedDataCount = 0
    InitializeBinnedData BinnedDataCount, BinnedData, True
    BinData = False

End Function

Public Function ComputeBinDetails(ByRef sngBinRangeMinimum As Single, ByRef sngBinRangeMaximum As Single, ByRef sngBinSize As Single, _
                                  ByRef lngBinCountComputed As Long, ByVal blnRequireNegativeStartBin As Boolean, ByVal intStartBinDigitsAfterDecimal As Integer) As Boolean

    ' The user must provide sngBinRangeMinimum, sngBinRangeMaximum, and sngBinSize
    ' This function will round the ranges and bin size to appropriate values, then compute
    '  the number of bins to reserve given the values
    
    Dim lngMultiplier As Long
    Dim intDecrementCount As Integer
    Dim sngDataRangeLastVal As Single
    
    lngBinCountComputed = 0
    
    If IsInfinity(sngBinRangeMinimum) Or IsInfinity(sngBinRangeMaximum) Then
        ' Minimum and/or maximum are infinity; cannot make a histogram
        ComputeBinDetails = False
        Exit Function
    End If
    
    ' sngBinRangeMinimum should be less than or requal to 0 if blnRequireNegativeStartBin is True
    If sngBinRangeMinimum > 0 And blnRequireNegativeStartBin Then
        sngBinRangeMinimum = 0
    End If

    ' Round sngBinRangeMinimum and sngBinRangeMaximum down/up, depending on StartBinDigitsAfterDecimal

    If intStartBinDigitsAfterDecimal < 0 Then intStartBinDigitsAfterDecimal = 0
    If intStartBinDigitsAfterDecimal > 7 Then intStartBinDigitsAfterDecimal = 7

    lngMultiplier = 10 ^ mStartBinDigitsAfterDecimal

    ' Use Int() to round down
    If blnRequireNegativeStartBin Then
        sngBinRangeMinimum = (-1 * (Int(Abs(sngBinRangeMinimum * lngMultiplier)) + 1)) / lngMultiplier
    Else
        sngBinRangeMinimum = (Int(sngBinRangeMinimum * lngMultiplier)) / lngMultiplier
    End If
    
    sngBinRangeMaximum = (Int(sngBinRangeMaximum * lngMultiplier) + 1) / lngMultiplier

    If IsInfinity(sngBinSize) Then
        ' sngBinSize is Infinity; set to the default bin size
        sngBinSize = mDefaultBinSize
    End If
    
    If sngBinSize <= 0 Then sngBinSize = mDefaultBinSize
    If sngBinSize > 1 Then
        sngBinSize = Round(sngBinSize, 0)
    Else
        ' Round to 2 significant figures
        sngBinSize = CSng(Format(sngBinSize, "0.0E+00"))
        If sngBinSize <= 0 Then sngBinSize = mDefaultBinSize
    End If

    ' Make sure sngBinRangeMinimum and sngBinRangeMaximum are evenly divisible by sngBinSize
    ' We'll do this by decrementing/incrementing sngBinRangeMinimum in increments of sngBinSize/10
    '  until we pass or equal sngDataRangeLastVal

    intDecrementCount = 0
    sngDataRangeLastVal = sngBinRangeMinimum - sngBinSize
    Do While Round(sngBinRangeMinimum / sngBinSize, 0) <> Round(sngBinRangeMinimum / sngBinSize, mStartBinDigitsAfterDecimal + 1) And sngBinRangeMinimum > sngDataRangeLastVal And intDecrementCount <= 15
        sngBinRangeMinimum = Round(sngBinRangeMinimum - sngBinSize / 10#, mStartBinDigitsAfterDecimal + 2)
        intDecrementCount = intDecrementCount + 1
    Loop
    If sngBinRangeMinimum < sngDataRangeLastVal Then
        ' Acceptable answer not found
        Debug.Assert False
        sngBinRangeMinimum = sngDataRangeLastVal
    End If

    intDecrementCount = 0
    sngDataRangeLastVal = sngBinRangeMaximum + sngBinSize
    Do While Round(sngBinRangeMaximum / sngBinSize, 0) <> Round(sngBinRangeMaximum / sngBinSize, mStartBinDigitsAfterDecimal + 1) And sngBinRangeMaximum < sngDataRangeLastVal And intDecrementCount <= 15
        sngBinRangeMaximum = Round(sngBinRangeMaximum + sngBinSize / 10#, mStartBinDigitsAfterDecimal + 2)
        intDecrementCount = intDecrementCount + 1
    Loop
    If sngBinRangeMaximum > sngDataRangeLastVal Then
        sngBinRangeMaximum = sngDataRangeLastVal
    End If
    

    ' Calculate lngErrorsBinned; do not allow more than 1,000,000 bins
    Do
        On Error GoTo HandleBinCountOverload
        lngBinCountComputed = (sngBinRangeMaximum - sngBinRangeMinimum) / sngBinSize
        If lngBinCountComputed > 1000000 Then
            sngBinSize = sngBinSize * 10
        End If
    Loop While lngBinCountComputed > 1000000

    ComputeBinDetails = True

    Exit Function

HandleBinCountOverload:
    If Err.Number = 6 Then
        ' Set .lngBinCountComputed to a huge number (but still a valid Long) so that the BinSize gets incremented
        lngBinCountComputed = 1000000000#
        Resume Next
    Else
        Debug.Assert False
        lngBinCountComputed = 0
        ComputeBinDetails = False
        Exit Function
    End If

End Function

Public Function DataPointToBinNumber(ThisValue As Single, StartBin As Single, BinSize As Single) As Long
    Dim WorkingValue As Single

    ' First subtract StartBin from ThisValue
    ' For example, if StartBin is 500 and ThisValue is 500.28, then WorkingValue = 0.28
    ' Or, if StartBin is 500 and ThisValue is 530.83, then WorkingValue = 30.83
    WorkingValue = ThisValue - StartBin

    ' Now, dividing WorkingValue by BinSize and rounding to the nearest integer
    '  actually gives the bin
    ' For example, given WorkingValue = 0.28 and BinSize = 0.1, Bin = CLng(2.8) = 3
    ' Or, given WorkingValue = 30.83 and BinSize = 0.1, Bin = CLng(308.3) = 308
    DataPointToBinNumber = CLng(WorkingValue / BinSize)

End Function

Private Function IsInfinity(ByVal dblValue As Double) As Boolean
    Dim blnInfinity As Boolean
    
    On Error Resume Next
    blnInfinity = False
    If dblValue > 1.79769313486231E+308 Then
        blnInfinity = True
    ElseIf dblValue < -1.79769313486231E+308 Then
        blnInfinity = True
    ElseIf Err.Number <> 0 Then
        blnInfinity = True
    End If
    
    IsInfinity = blnInfinity
End Function

Private Sub InitializeBinnedData(ByVal BinnedDataCount As Long, ByRef BinnedData() As Long, blnResetDataValueRange As Boolean)

    If BinnedDataCount > 0 Then
        ReDim BinnedData(0 To BinnedDataCount)
    Else
        ReDim BinnedData(0)
    End If

    If blnResetDataValueRange Then
        mBinRangeMaximum = 0
        mBinCountMaximum = 0
    End If

End Sub

Private Sub Class_Initialize()
    mStartBin = 0
    mStartBinDigitsAfterDecimal = 1

    mBinSize = 0.5
    mDefaultBinSize = 0.5

    mRequireNegativeStartBin = False
    mShowMessages = False
    mErrorNumber = 0

End Sub

