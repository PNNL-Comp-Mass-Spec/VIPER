VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FastSearchArrayLong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' FastSearchArrayLong class
' This class can be used to find the indices of the values in an array that
'  match the search value
' The search array need not be sorted
' This class uses the QSLong class

Private mSearchValsSorted() As Long           ' Array of values to be searched (sorted)
Private mOriginalIndices() As Long      ' Original indices of the values (sorted parallel with mSearchVals)

Private mLoIndex As Long
Private mHiIndex As Long

Private Function BinarySearchLngLocal(lngArrayToSearch() As Long, lngItemToSearchFor As Long) As Long
    ' Looks through lngArrayToSearch() for lngItemToSearchFor, returning
    '  the index of the item if found, or -1 if not found
    ' Assumes lngArrayToSearch() is already sorted
    
    Dim lngMidIndex As Long
    Dim lngFirstIndex As Long, lngLastIndex As Long
    
On Error GoTo BinarySearchLngErrorHandler

    lngFirstIndex = LBound(lngArrayToSearch())
    lngLastIndex = UBound(lngArrayToSearch())
    
    lngMidIndex = (lngFirstIndex + lngLastIndex) \ 2            ' Note: Using Integer division
    If lngMidIndex < lngFirstIndex Then lngMidIndex = lngFirstIndex
    
    Do While lngFirstIndex <= lngLastIndex And lngArrayToSearch(lngMidIndex) <> lngItemToSearchFor
        If lngItemToSearchFor < lngArrayToSearch(lngMidIndex) Then
            ' Search the lower half
            lngLastIndex = lngMidIndex - 1
        ElseIf lngItemToSearchFor > lngArrayToSearch(lngMidIndex) Then
            ' Search the upper half
            lngFirstIndex = lngMidIndex + 1
        End If
        ' Compute the new mid point
        lngMidIndex = (lngFirstIndex + lngLastIndex) \ 2
        If lngMidIndex < lngFirstIndex Then Exit Do
    Loop
    
    If lngMidIndex >= lngFirstIndex And lngMidIndex <= lngLastIndex Then
        If lngArrayToSearch(lngMidIndex) = lngItemToSearchFor Then
            BinarySearchLngLocal = lngMidIndex
        Else
            BinarySearchLngLocal = -1
        End If
    Else
        BinarySearchLngLocal = -1
    End If
    Exit Function
    
BinarySearchLngErrorHandler:
    Debug.Assert False
    BinarySearchLngLocal = -1
    Exit Function
End Function

Public Function Fill(lngSearchVals() As Long) As Boolean
    ' lngSearchVals() need not be sorted
    ' Returns True if success, false if failure
    
    Dim lngDataPoints As Long
    Dim lngPointerIndex As Long
    Dim objQSLong As QSLong
    
On Error GoTo err_Fill
    mLoIndex = LBound(lngSearchVals)
    mHiIndex = UBound(lngSearchVals)
    lngDataPoints = mHiIndex - mLoIndex + 1
    
    If lngDataPoints > 0 Then
        ReDim mSearchValsSorted(mLoIndex To mHiIndex)
        
        mSearchValsSorted() = lngSearchVals()
        
        ' Sort mSearchValsSorted()
        ' First construct a pointer array
        ReDim mOriginalIndices(mLoIndex To mHiIndex)
        For lngPointerIndex = mLoIndex To mHiIndex
            mOriginalIndices(lngPointerIndex) = lngPointerIndex
        Next lngPointerIndex
        
        Set objQSLong = New QSLong
        If Not objQSLong.QSAsc(mSearchValsSorted(), mOriginalIndices()) Then
            ' Failure with QSort
            Debug.Assert False
        End If
        Set objQSLong = Nothing
        
        Fill = True
    Else
        Fill = False
    End If
    
    Exit Function

err_Fill:
    Fill = False
End Function

Private Sub Class_Terminate()
    On Error Resume Next
    Erase mSearchValsSorted()
    Erase mOriginalIndices()
End Sub

Public Function FindMatchingIndices(ByVal lngValueToFind As Long, ByRef lngMatchingIndices() As Long, ByRef lngMatchCount As Long) As Boolean
    ' Searches mSearchValsSorted() for lngValueToFind
    ' Populates lngMatchingIndices() with the indices of the match
    '  (returning the index in the original array)
    ' Returns True if match found, False if error or no match

    Dim lngPointerIndex As Long, lngFirstIndex As Long, lngLastIndex As Long
    
On Error GoTo FindMatchingIndicesErrorHandler
    
    ' Initialize the search by setting lngFirstIndex and lngLastIndex to the range of mSearchValsSorted()
    lngFirstIndex = mLoIndex
    lngLastIndex = mHiIndex
    
    FindRange lngFirstIndex, lngLastIndex, lngValueToFind
    
    If lngFirstIndex > lngLastIndex Then
        ' If lngFirstIndex is > lngLastIndex then the search failed
        ReDim lngMatchingIndices(0)
        lngMatchCount = 0
        FindMatchingIndices = False
    Else
        ' If lngFirstIndex is <= lngLastIndex then the search found a match
        ReDim lngMatchingIndices(0 To lngLastIndex - lngFirstIndex)
        lngMatchCount = lngLastIndex - lngFirstIndex + 1
        
        For lngPointerIndex = lngFirstIndex To lngLastIndex
            lngMatchingIndices(lngPointerIndex - lngFirstIndex) = mOriginalIndices(lngPointerIndex)
        Next lngPointerIndex
        
        FindMatchingIndices = True
    End If

    Exit Function
    
FindMatchingIndicesErrorHandler:
Debug.Assert False
End Function

Private Sub FindRange(ByRef lngMinIndex As Long, ByRef lngMaxIndex As Long, lngValueToFind As Long)
    ' Procedure to find all indices in mSearchValsSorted() containing lngValueToFind
    ' The method for finding the values uses a binary search;
    '  thus mSearchValsSorted() must be sorted prior to calling this function
    
    Dim lngMatchIndex As Long
    
    ' Start by using a binary search to find the first match in mSearchValsSorted() to lngValueToFind
    lngMatchIndex = BinarySearchLngLocal(mSearchValsSorted(), lngValueToFind)
    If lngMatchIndex < 0 Then
        lngMinIndex = lngMaxIndex + 1
    Else
        ' Find the first and last indices in mSearchValsSorted() matching lngValueToFind
        ' Step backward while things match
        lngMinIndex = lngMatchIndex
        If lngMinIndex > mLoIndex Then
            Do While mSearchValsSorted(lngMinIndex - 1) = lngValueToFind
                lngMinIndex = lngMinIndex - 1
                If lngMinIndex = mLoIndex Then Exit Do
            Loop
        End If

        ' Step forward while things match
        lngMaxIndex = lngMatchIndex
        If lngMaxIndex < mHiIndex Then
            Do While mSearchValsSorted(lngMaxIndex + 1) = lngValueToFind
                lngMaxIndex = lngMaxIndex + 1
                If lngMaxIndex = mHiIndex Then Exit Do
            Loop
        End If
    End If
    
End Sub

Public Function GetValueByIndex(lngIndex As Long) As Long
    Dim lngPointerIndex As Long
    
    lngPointerIndex = BinarySearchLngLocal(mOriginalIndices(), lngIndex)
    
    If lngPointerIndex >= mLoIndex And lngPointerIndex <= mHiIndex Then
        GetValueByIndex = mSearchValsSorted(lngPointerIndex)
    Else
        GetValueByIndex = 0
    End If
End Function

