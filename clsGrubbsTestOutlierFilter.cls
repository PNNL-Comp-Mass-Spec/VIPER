VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrubbsTestOutlierFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class can be used to remove outliers from a list of numbers (doubles)
' It uses Grubb's test to determine whether or not each number in the list
'  is far enough away from the mean to be thrown out
'
' Utilizes classes QSDouble and StatDoubles
'
'created: 08/13/2004 mem

Option Explicit

Public Enum eclConfidenceLevelConstants
    ecl95Pct = 0
    ecl97Pct = 1
    ecl99Pct = 2
End Enum

Private mConfidenceLevel As eclConfidenceLevelConstants
Private mMinFinalValueCount As Long
Private mIterate As Boolean

Public Property Let ConfidenceLevel(ByVal eConfidenceLevel As eclConfidenceLevelConstants)
    mConfidenceLevel = eConfidenceLevel
End Property
Public Property Get ConfidenceLevel() As eclConfidenceLevelConstants
    ConfidenceLevel = mConfidenceLevel
End Property

Public Property Let MinFinalValueCount(ByVal lngCount As Long)
    If lngCount < 2 Then lngCount = 2
    mMinFinalValueCount = lngCount
End Property
Public Property Get MinFinalValueCount() As Long
    MinFinalValueCount = mMinFinalValueCount
End Property

Public Property Let RemoveMultipleValues(ByVal blnEnabled As Boolean)
   mIterate = blnEnabled
End Property
Public Property Get RemoveMultipleValues() As Boolean
   RemoveMultipleValues = mIterate
End Property

Public Function RemoveOutliers(ByRef dblValues() As Double, ByRef lngIndexPointers() As Long, ByRef lngValueCountRemovedOut As Long) As Boolean
    ' Removes outliers from dblValues() using Grubb's test and the given confidence level
    ' lngIndexPointers() is an array of longs that is parallel to dblValues(), and will be
    '  kept in sync with any changes made to dblValues

    ' If intMaxIterations > 1, then will repeatedly remove the outliers, until no outliers
    '  remain or the number of values falls below intMinValues
    '
    ' Returns True if success (even if no values removed) and false if an error or dblValuesSorted doesn't contain any data
    ' Returns the number of values removed in lngValueCountRemovedOut
    
    Dim qsDbl As New QSDouble

    Dim lngDataCount As Long                ' Number of data points
    Dim dblValuesSorted() As Double         ' Sorted array of doubles            ' 0-based array
    Dim lngIndexPointersSorted() As Long
    
    Dim blnSuccess As Boolean
    Dim blnValuesRemoved As Boolean
    
    On Error GoTo RemoveOutliersErrorHandler
    
    lngDataCount = 0
    dblValuesSorted = dblValues
    lngIndexPointersSorted = lngIndexPointers
    
    If qsDbl.QSAsc(dblValuesSorted(), lngIndexPointersSorted()) Then
        lngDataCount = UBound(dblValuesSorted) + 1
        If lngDataCount > 0 Then
            blnSuccess = True
        Else
            blnSuccess = False
        End If
    Else
       blnSuccess = False
    End If
    
    
    If lngDataCount <= 0 Or Not blnSuccess Then
        blnSuccess = False
    Else
    
        ' Copy the data from dblValuesSorted back to dblValues
        dblValues = dblValuesSorted
        lngIndexPointers = lngIndexPointersSorted
        
        lngValueCountRemovedOut = 0
        Do
            blnValuesRemoved = RemoveOutliersWork(dblValues, lngIndexPointers, mConfidenceLevel)
            
            If blnValuesRemoved Then
                lngValueCountRemovedOut = lngValueCountRemovedOut + 1
            End If
            
        Loop While blnValuesRemoved And mIterate And UBound(dblValues) + 1 > mMinFinalValueCount
        blnSuccess = True
        
    End If
    
    RemoveOutliers = blnSuccess
    Exit Function

RemoveOutliersErrorHandler:
    Debug.Assert False
    RemoveOutliers = False

End Function

Private Function RemoveOutliersWork(ByRef dblValues() As Double, ByRef lngIndexPointers() As Long, ByVal eclConfidenceLevel As eclConfidenceLevelConstants) As Boolean
    ' Removes, at most, one outlier from dblValues (and from the corresponding position in lngIndexPointers)
    ' Returns True if an outlier is removed, and false if not
    ' Returns false if an error occurs
    '
    ' NOTE: This function assumes that dblValues() is sorted ascending, and thus only tests
    '       the first and last value in the list as an outlier (removing only one of them, if appropriate)
    
    Dim objStatDoubles As New StatDoubles
    Dim blnValueRemoved As Boolean
    
    Dim lngCount As Long
    
    Dim lngIndex As Long, lngIndex2 As Long
    Dim lngTargetIndex As Long
    Dim dblTargetDistance As Double
    Dim dblCompareDistance As Double
    
    Dim dblMean As Double
    Dim dblStDev As Double
    Dim dblZScore As Double
    Dim dblPValue As Double
    
    lngCount = UBound(dblValues) + 1
    If lngCount < 3 Then
        ' Cannot remove an outlier from fewer than 3 values
        RemoveOutliersWork = False
        Exit Function
    End If
    
    blnValueRemoved = False
    If objStatDoubles.Fill(dblValues) Then
        dblMean = objStatDoubles.Mean
        dblStDev = objStatDoubles.StDev
        
        If dblStDev > 0 Then
            ' Find the value furthest away from the mean
            ' Since dblValues() is sorted, it can only be the first or last value
            
            lngTargetIndex = 0
            dblTargetDistance = Abs(dblValues(0) - dblMean)
            
            dblCompareDistance = Abs(dblValues(lngCount - 1) - dblMean)
            If dblCompareDistance > dblTargetDistance Then
                dblTargetDistance = dblCompareDistance
                lngTargetIndex = lngCount - 1
            End If
            
            ' Compute the z-score for lngTargetIndex
            dblZScore = dblTargetDistance / dblStDev
            
            ' Compute the p-value, based on eclConfidenceLevel
            Select Case eclConfidenceLevel
            Case ecl95Pct
                ' Estimate the P value at the 95%'ile using a formula provided by
                '  Robin Edwards <robin.edwards@argonet.co.uk>
                dblPValue = (3.6996 * lngCount + 145.9 - 186.7 / lngCount) / (lngCount + 59.5 + 58.5 / lngCount)
            Case ecl97Pct
                dblPValue = Lookup97PctPValue(lngCount)
            Case Else
                ' Includes ecl99pct
                ' Estimate the P value at the 99%'ile using a formula provided by
                '  Robin Edwards <robin.edwards@argonet.co.uk>
                dblPValue = (4.1068 * lngCount + 273.6 - 328.5 / lngCount) / (lngCount + 88.7 + 185 / lngCount)
            End Select
            
            If dblZScore > dblPValue Then
                 ' Remove the value
                 ' Copy the data in place, skipping the outlier value
                 
                 lngIndex2 = 0
                 For lngIndex = 0 To lngCount - 1
                        If lngIndex <> lngTargetIndex Then
                            dblValues(lngIndex2) = dblValues(lngIndex)
                            lngIndexPointers(lngIndex2) = lngIndexPointers(lngIndex)
                            lngIndex2 = lngIndex2 + 1
                        End If
                 Next lngIndex
                 
                 ReDim Preserve dblValues(lngCount - 2)
                 ReDim Preserve lngIndexPointers(lngCount - 2)
                 blnValueRemoved = True
            End If
        End If
    End If
    
    RemoveOutliersWork = blnValueRemoved
    Exit Function
    
RemoveOutliersErrorHandler:
    Debug.Assert False
    RemoveOutliersWork = False
   
End Function

Private Function Lookup97PctPValue(lngCount As Long) As Double

    If lngCount <= 3 Then
        Lookup97PctPValue = 1.15
    ElseIf lngCount = 4 Then
        Lookup97PctPValue = 1.48
    ElseIf lngCount = 5 Then
        Lookup97PctPValue = 1.71
    ElseIf lngCount = 6 Then
        Lookup97PctPValue = 1.89
    ElseIf lngCount = 7 Then
        Lookup97PctPValue = 2.02
    ElseIf lngCount = 8 Then
        Lookup97PctPValue = 2.13
    ElseIf lngCount = 9 Then
        Lookup97PctPValue = 2.21
    ElseIf lngCount = 10 Then
        Lookup97PctPValue = 2.29
    ElseIf lngCount = 11 Then
        Lookup97PctPValue = 2.34
    ElseIf lngCount = 12 Then
        Lookup97PctPValue = 2.41
    ElseIf lngCount = 13 Then
        Lookup97PctPValue = 2.46
    ElseIf lngCount = 14 Then
        Lookup97PctPValue = 2.51
    ElseIf lngCount = 15 Then
        Lookup97PctPValue = 2.55
    ElseIf lngCount = 16 Then
        Lookup97PctPValue = 2.59
    ElseIf lngCount = 17 Then
        Lookup97PctPValue = 2.62
    ElseIf lngCount = 18 Then
        Lookup97PctPValue = 2.65
    ElseIf lngCount = 19 Then
        Lookup97PctPValue = 2.68
    ElseIf lngCount = 20 Then
        Lookup97PctPValue = 2.71
    ElseIf lngCount = 21 Then
        Lookup97PctPValue = 2.73
    ElseIf lngCount = 22 Then
        Lookup97PctPValue = 2.76
    ElseIf lngCount = 23 Then
        Lookup97PctPValue = 2.78
    ElseIf lngCount = 24 Then
        Lookup97PctPValue = 2.8
    ElseIf lngCount = 25 Then
        Lookup97PctPValue = 2.82
    ElseIf lngCount = 26 Then
        Lookup97PctPValue = 2.84
    ElseIf lngCount = 27 Then
        Lookup97PctPValue = 2.86
    ElseIf lngCount = 28 Then
        Lookup97PctPValue = 2.88
    ElseIf lngCount = 29 Then
        Lookup97PctPValue = 2.89
    ElseIf lngCount = 30 Then
        Lookup97PctPValue = 2.91
    ElseIf lngCount = 31 Then
        Lookup97PctPValue = 2.92
    ElseIf lngCount = 32 Then
        Lookup97PctPValue = 2.94
    ElseIf lngCount = 33 Then
        Lookup97PctPValue = 2.95
    ElseIf lngCount = 34 Then
        Lookup97PctPValue = 2.97
    ElseIf lngCount = 35 Then
        Lookup97PctPValue = 2.98
    ElseIf lngCount = 36 Then
        Lookup97PctPValue = 2.99
    ElseIf lngCount = 37 Then
        Lookup97PctPValue = 3
    ElseIf lngCount = 38 Then
        Lookup97PctPValue = 3.01
    ElseIf lngCount = 39 Then
        Lookup97PctPValue = 3.03
    ElseIf lngCount = 40 Then
        Lookup97PctPValue = 3.04
    ElseIf lngCount <= 50 Then
        Lookup97PctPValue = 3.13
    ElseIf lngCount <= 60 Then
        Lookup97PctPValue = 3.2
    ElseIf lngCount <= 70 Then
        Lookup97PctPValue = 3.26
    ElseIf lngCount <= 80 Then
        Lookup97PctPValue = 3.31
    ElseIf lngCount <= 90 Then
        Lookup97PctPValue = 3.35
    ElseIf lngCount <= 100 Then
        Lookup97PctPValue = 3.38
    ElseIf lngCount <= 110 Then
        Lookup97PctPValue = 3.42
    ElseIf lngCount <= 120 Then
        Lookup97PctPValue = 3.44
    ElseIf lngCount <= 130 Then
        Lookup97PctPValue = 3.47
    ElseIf lngCount <= 140 Then
        Lookup97PctPValue = 3.49
    Else
        Lookup97PctPValue = 3.5
    End If
    
End Function


Private Sub Class_Initialize()
    mConfidenceLevel = ecl95Pct
    mMinFinalValueCount = 3
    mIterate = False
End Sub
