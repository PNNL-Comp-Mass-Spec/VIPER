VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScanNumberTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This class can be used to track the average scan increment present for data in an input file

Private mCurrentAutoScanNumber As Single              ' This is a single precision value and not a long since mAutoNumberIncrement can be a decimal; when storing the scan we'll round to the nearest integer

Private mAutoNumberIncrement As Single                ' The value to increment each scan by when auto-numbering; initially populated with CInt(mAverageScanIncrement) but the user can override this
Private mAverageScanIncrement As Single               ' The average scan increment; updated continually if not auto-numbering
Private mDataCountForScanIncrementAverage As Long     ' The number of data points used to compute the average scan gap

Public Property Get AutoNumberIncrement() As Single
    AutoNumberIncrement = mAutoNumberIncrement
End Property
Public Property Let AutoNumberIncrement(ByVal Value As Single)
    mAutoNumberIncrement = Value
    If mAutoNumberIncrement < 1 Then
        mAutoNumberIncrement = 1
    End If
End Property

Public Property Get AverageScanIncrement() As Single
    AverageScanIncrement = mAverageScanIncrement
End Property

Public Property Get CurrentAutoScanNumber() As Single
    CurrentAutoScanNumber = mCurrentAutoScanNumber
End Property
Public Property Let CurrentAutoScanNumber(ByVal Value As Single)
    mCurrentAutoScanNumber = Value
    If mCurrentAutoScanNumber < 0 Then
        mCurrentAutoScanNumber = 0
    End If
End Property

Private Sub Class_Initialize()
    Me.Reset
End Sub

Public Sub AddScanNumberAndUpdateAverageIncrement(ByVal lngScanNumber As Long)
    Dim lngNewIncrement As Long

On Error GoTo AddScanNumberAndUpdateAverageIncrementErrorHandler
 
    If mCurrentAutoScanNumber = 0 Then
        ' First call to AddScanNumberAndUpdateAverageIncrement since last Reset
        ' Simply udpate mCurrentAutoScanNumber
        mCurrentAutoScanNumber = lngScanNumber
    ElseIf lngScanNumber > mCurrentAutoScanNumber Then
        ' Update the average since the new scan number is greater than the previous one
    
        lngNewIncrement = CLng(lngScanNumber - mCurrentAutoScanNumber)
        If mDataCountForScanIncrementAverage = 0 Then
            mAverageScanIncrement = lngNewIncrement
        Else
            mAverageScanIncrement = (mAverageScanIncrement * mDataCountForScanIncrementAverage + lngNewIncrement) / (mDataCountForScanIncrementAverage + 1)
        End If
        mDataCountForScanIncrementAverage = mDataCountForScanIncrementAverage + 1
    
        mCurrentAutoScanNumber = lngScanNumber
    End If
    
    Exit Sub

AddScanNumberAndUpdateAverageIncrementErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "AddScanNumberAndUpdateAverageIncrement"
End Sub

Public Function GetNextAutoNumberedScan() As Long
    mCurrentAutoScanNumber = mCurrentAutoScanNumber + mAutoNumberIncrement
    GetNextAutoNumberedScan = CInt(Round(mCurrentAutoScanNumber, 0))
End Function

Public Sub Reset()
    mCurrentAutoScanNumber = 0

    mAutoNumberIncrement = 1
    mAverageScanIncrement = 1
    mDataCountForScanIncrementAverage = 0
    
End Sub

Public Sub SetAutoNumberIncrementToCurrentAverage()
    Me.AutoNumberIncrement = Me.AverageScanIncrement
End Sub
