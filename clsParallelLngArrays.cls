VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParallelLngArrays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const INITIAL_LISTSIZE As Long = 10000

Private mCount As Long
Private mKeys() As Long
Private mValues() As Long

' These two variables are used in the Exists() function to
' quickly check if a key doesn't exist when the data isn't sorted
Private mKeyValueMin As Long
Private mKeyValueMax As Long

Private mIsSorted As Boolean
Private mPreventDuplicateKeys As Boolean

' ---------------
' Properties
' ---------------
Property Get PreventDuplicateKeys() As Boolean
    PreventDuplicateKeys = mPreventDuplicateKeys
End Property
Property Let PreventDuplicateKeys(Value As Boolean)
    mPreventDuplicateKeys = Value
End Property

Property Get Count() As Long
    Count = mCount
End Property

Property Get IsSorted() As Boolean
    IsSorted = mIsSorted
End Property

Property Get Item(ByVal Key As Long) As Long
    Dim lngItem As Long
    Dim blnMatchFound As Boolean
    
    lngItem = GetItemForKey(Key, blnMatchFound)
    
    If blnMatchFound Then
        Item = lngItem
    Else
        Err.Raise 1001, , "Key not found"
    End If
End Function

Property Get KeyByIndex(ByVal lngIndex As Long) As Long
    If lngIndex >= 0 And lngIndex < mCount Then
        KeyByIndex = mKeys(lngIndex)
    Else
        Err.Raise 1002, , "Invalid index"
    End If
End Property

Property Get ValueByIndex(ByVal lngIndex As Long) As Long
    If lngIndex >= 0 And lngIndex < mCount Then
        ValueByIndex = mValues(lngIndex)
    Else
        Err.Raise 1002, , "Invalid index"
    End If
End Property

' ---------------
' Constructor
' ---------------
Private Sub Class_Initialize()
    Me.PreventDuplicateKeys = False
    Me.ClearAndResetSize
End Sub

' ---------------
' Functions
' ---------------

' Add a new Key / Value pair
' If PreventDuplicateKeys = True, then will make sure duplicate keys are not added
' Return True if successfully added, false if a duplicate exists and thus not added
Public Function Add(ByVal Key As Long, ByVal Value As Long) As Boolean
     
    If mCount >= UBound(mKeys) Then
        SetSize UBound(mKeys) * 2
    End If
    
    If mPreventDuplicateKeys Then
        ' Make sure this key doesn't yet exist
        If Me.Exists(Key) Then
            Add = False
            Exit Function
        End If
    End If
    
    mKeys(mCount) = Key
    mValues(mCount) = Value
    
    mCount = mCount + 1
    mIsSorted = False
    
    ' Update the min and max key value tracking variables
    If mCount = 1 Then
        mKeyValueMin = Key
        mKeyValueMax = Key
    Else
        If Key < mKeyValueMin Then
            mKeyValueMin = Key
        End If

        If Key > mKeyValueMax Then
            mKeyValueMax = Key
        End If
    End If
    
    Add = True
     
End Function

' Clear the parallel arrays (doesn't actually release the memory, however)
Public Sub Clear()
    mCount = 0
    mIsSorted = False
    mKeyValueMin = 0
    mKeyValueMax = 0
End Sub

' Clear the parallel arrays and reset the reserved memory to INITIAL_LISTSIZE
Public Sub ClearAndResetSize()
    Me.Clear
    
    ReDim mKeys(INITIAL_LISTSIZE)
    ReDim mValues(INITIAL_LISTSIZE)
End Sub

' Return True if Key exists
Public Function Exists(ByVal Key As Long) As Boolean
    Dim lngIndexMatch As Long
     
    lngIndexMatch = GetIndexForKey(Key)
    If lngIndexMatch >= 0 Then
        Exists = True
    Else
        Exists = False
    End If
    
End Function

' Determines the index for the specified Key
Private Function GetIndexForKey(ByVal Key As Long) As Long
    Dim lngIndex As Long
    Dim lngIndexMatch As Long
    
    lngIndexMatch = -1
    
    If mCount > 0 Then
        If mIsSorted Then
            ' Use a binary search to look for the key
            lngIndexMatch = BinarySearchLng(mKeys(), Key, 0, mCount - 1)
        Else
            ' First see if Key is within the range mKeyValueMin and mKeyValueMax
            ' If it isn't, then the key definitely doesn't exist in mKeys()
            
            If Key >= mKeyValueMin And Key <= mKeyValueMax Then
                ' Use a brute-force search to look for the key
                For lngIndex = 0 To mCount - 1
                    If mKeys(lngIndex) = Key Then
                        lngIndexMatch = lngIndex
                        Exit For
                    End If
                Next lngIndex
            End If
        End If
    End If

    GetIndexForKey = lngIndexMatch
End Function

' Return the value associated with the given key
Public Function GetItemForKey(ByVal Key As Long, ByRef blnMatchFound As Boolean) As Long
    Dim lngIndexMatch As Long
   
    lngIndexMatch = GetIndexForKey(Key)
    
    If lngIndexMatch >= 0 Then
        blnMatchFound = True
        GetItemForKey = mValues(lngIndexMatch)
    Else
        blnMatchFound = False
        GetItemForKey = 0
    End If
End Function

' Removes the specified key
Public Sub Remove(ByVal Key As Long)
    Dim lngIndexMatch As Long
    Dim lngIndex As Long
   
    lngIndexMatch = GetIndexForKey(Key)
    
    If lngIndexMatch >= 0 Then
        ' Remove this key by shuffling all of the data
        For lngIndex = lngIndexMatch To mCount - 2
            mKeys(lngIndex) = mKeys(lngIndex + 1)
            mValues(lngIndex) = mValues(lngIndex + 1)
        Next lngIndex
        
        ' Decrement the count
        ' Note: we can leave mSorted unchanged since, if the data started out sorted, it is still sorted
        mCount = mCount - 1
    End If
    
End Sub

' Removes all keys; simply calls Me.Clear()
Public Sub RemoveAll()
    Me.Clear
End Sub

' Call this sub after adding all of the items using .Add()
' This will greatly speedup calls to .Exists() and .GetItemForKey()
Public Sub SortNow()
    If mCount > 1 Then
        ShellSortLongWithParallelLong mKeys, mValues, 0, mCount - 1
    End If
    
    mIsSorted = True
End Sub

' Can be used to expand (but not contract) the length of the parallel arrays
Public Sub SetSize(ByVal MaxExpectedRows As Long)
    If MaxExpectedRows > mCount Then
        ReDim Preserve mKeys(MaxExpectedRows)
        ReDim Preserve mValues(MaxExpectedRows)
    End If
End Sub

' Updates the value for the given Key
' Returns True if the key was found; False if not
Public Function Update(ByVal Key As Long, ByVal Value As Long) As Boolean
    Dim lngIndexMatch As Long
   
    lngIndexMatch = GetIndexForKey(Key)
    
    If lngIndexMatch >= 0 Then
        mValues(lngIndexMatch) = Value
        Update = True
    Else
        Update = False
    End If
End Function

' ---------------
' Search and Sort functions
' ---------------
Private Function BinarySearchLng(ByRef lngArrayToSearch() As Long, ByVal lngItemToSearchFor As Long, Optional ByVal lngFirstIndex As Long = -1, Optional ByVal lngLastIndex As Long = -1) As Long
    ' Looks through lngArrayToSearch() for lngItemToSearchFor, returning
    '  the index of the item if found, or -1 if not found
    ' Assumes lngArrayToSearch() is already sorted
    ' If lngFirstIndex < 0 or lngLastIndex is < 0 then uses LBound() and UBound() to determine the array range
    
    Dim lngMidIndex As Long
    
On Error GoTo BinarySearchLngErrorHandler

    If lngFirstIndex < 0 Or lngLastIndex < 0 Then
        lngFirstIndex = LBound(lngArrayToSearch())
        lngLastIndex = UBound(lngArrayToSearch())
    End If
    
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
            BinarySearchLng = lngMidIndex
        Else
            BinarySearchLng = -1
        End If
    Else
        BinarySearchLng = -1
    End If
    Exit Function
    
BinarySearchLngErrorHandler:
    Debug.Assert False
    BinarySearchLng = -1
    Exit Function
End Function

Private Sub ShellSortLongWithParallelLong(ByRef lngArray() As Long, ByRef lngParallelArray() As Long, ByVal lngLowIndex As Long, ByVal lngHighIndex As Long)
    Dim lngCount As Long
    Dim lngIncrement As Long
    Dim lngIndex As Long
    Dim lngIndexCompare As Long
    Dim lngCompareVal As Long
    Dim lngSwap As Long

On Error GoTo ShellSortErrorHandler

' sort array[lngLowIndex..lngHighIndex]

    ' compute largest increment
    lngCount = lngHighIndex - lngLowIndex + 1
    lngIncrement = 1
    If (lngCount < 14) Then
        lngIncrement = 1
    Else
        Do While lngIncrement < lngCount
            lngIncrement = 3 * lngIncrement + 1
        Loop
        lngIncrement = lngIncrement \ 3
        lngIncrement = lngIncrement \ 3
    End If

    Do While lngIncrement > 0
        ' sort by insertion in increments of lngIncrement
        For lngIndex = lngLowIndex + lngIncrement To lngHighIndex
            lngCompareVal = lngArray(lngIndex)
            lngSwap = lngParallelArray(lngIndex)
            For lngIndexCompare = lngIndex - lngIncrement To lngLowIndex Step -lngIncrement
                ' Use <= to sort ascending; Use > to sort descending
                If lngArray(lngIndexCompare) <= lngCompareVal Then Exit For
                lngArray(lngIndexCompare + lngIncrement) = lngArray(lngIndexCompare)
                lngParallelArray(lngIndexCompare + lngIncrement) = lngParallelArray(lngIndexCompare)
            Next lngIndexCompare
            lngArray(lngIndexCompare + lngIncrement) = lngCompareVal
            lngParallelArray(lngIndexCompare + lngIncrement) = lngSwap
        Next lngIndex
        lngIncrement = lngIncrement \ 3
    Loop
    
    Exit Sub

ShellSortErrorHandler:
    Debug.Assert False
End Sub
