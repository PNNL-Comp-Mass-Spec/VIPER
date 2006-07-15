Attribute VB_Name = "ListViewRoutines"
Option Explicit

' Written by Matthew Monroe
'
' First written in Richland, WA in 2003
'
' Last Modified:    August 17, 2003
' Version:          1.04

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SETREDRAW = &HB

' Maximum number of listviews to keep track of for the FindMatchingIDs history
Private Const MAX_LISTVIEW_COUNT = 10

Public Enum sfSortFormatConstants
    sfText = 0
    sfNumeric
End Enum

Public Type udtColumnSortFormatType
    SortColumnIndexSaved As Long
    ColumnCount As Long
    ColumnSortOrder() As sfSortFormatConstants      ' 0-based array
    SortKeyColumnIndex As Long
End Type

Private Type udtFindTextHistoryType
    SearchText As String        ' The last text searched for
    MatchIndex As Long          ' Index of the last match
End Type

Private mFindTextIndexHistory(MAX_LISTVIEW_COUNT) As udtFindTextHistoryType

Public Function ListViewConstructSortKeyFormatString(lngMaxValue As Long) As String
    ' Examines lngMaxValue to determine the number of digits in the number
    ' Returns a string of zeroes equal to the number of digits in the number
    
    If lngMaxValue > 1 Then
        ListViewConstructSortKeyFormatString = String(RoundToNearest(Log(lngMaxValue) / Log(10), 1, False) + 1, "0")
    Else
        ListViewConstructSortKeyFormatString = "0"
    End If

End Function

Public Function ListViewFindText(ByVal strSearchText As String, lvwThisListView As MSComctlLib.ListView, eListViewID As Integer, lngMaxColumnIndex As Long) As Long
    ' Looks for strSearchText in lvwThisListView
    ' Keeps track of the last matching row so that user can sequentially match the same text string
    ' Returns the row index of the match
    ' Returns -1 if the match is not found or if eListViewID is greater than MAX_LISTVIEW_COUNT
    
    Dim lngListItemCount As Long
    Dim lngIndex As Long, lngSubIndex As Long, lngIndexAtSearchStart As Long
    Dim lngIndexOfMatch As Long
    Dim strSearchTextCapitalized As String
    
    lngListItemCount = lvwThisListView.ListItems.Count
    
    If lngListItemCount = 0 Or eListViewID > MAX_LISTVIEW_COUNT Then
        ListViewFindText = -1
        Exit Function
    End If
    
    strSearchTextCapitalized = UCase(strSearchText)
    
    lngIndexOfMatch = -1
    
    With mFindTextIndexHistory(eListViewID)
        If .SearchText = strSearchText Then
            ' User already searched for this text before, start searching at .MatchIndex + 1
            lngIndexAtSearchStart = .MatchIndex + 1
            If lngIndexAtSearchStart > lngListItemCount Then lngIndexAtSearchStart = 1
        Else
            ' New text search, start searching at .MatchIndex
            .SearchText = strSearchText
            lngIndexAtSearchStart = .MatchIndex + 1
        End If
    End With
    
    lngIndex = lngIndexAtSearchStart
    Do
        For lngSubIndex = 1 To lngMaxColumnIndex
            If InStr(UCase(lvwThisListView.ListItems(lngIndex).SubItems(lngSubIndex)), strSearchTextCapitalized) Then
                ' Match found
                
                ' Highlight the item and make sure its visible
                ListViewHighlightItem lvwThisListView, lngIndex
                
                lngIndexOfMatch = lngIndex
                
                Exit Do
            End If
        Next lngSubIndex
        
        lngIndex = lngIndex + 1
        If lngIndex > lngListItemCount Then lngIndex = 1
    Loop While lngIndex <> lngIndexAtSearchStart
    
    ListViewFindText = lngIndexOfMatch
End Function

Public Function ListViewGetItemIndex(lvwThisListView As MSComctlLib.ListView, lstListItem As MSComctlLib.ListItem) As Long
    ' Determines the Index number of lstListItem in lvwThisListView
    ' Returns -1 if not found
    
    Dim lngIndex As Long, lngMatchingIndex As Long
    
    If lstListItem Is Nothing Then
        ListViewGetItemIndex = -1
        Exit Function
    End If
    
    lngMatchingIndex = -1
    With lvwThisListView
        For lngIndex = 1 To .ListItems.Count
            If .ListItems(lngIndex) = lstListItem Then
                lngMatchingIndex = lngIndex
                Exit For
            End If
        Next lngIndex
    End With
    
    ListViewGetItemIndex = lngMatchingIndex
    
End Function

Public Function ListViewGetSearchHistoryText(intListViewID As Integer) As String
    If intListViewID >= 0 And intListViewID <= MAX_LISTVIEW_COUNT Then
        ListViewGetSearchHistoryText = mFindTextIndexHistory(intListViewID).SearchText
    Else
        ListViewGetSearchHistoryText = ""
    End If
End Function

Private Sub ListViewSortNumerically(ByRef lvwThisListView As MSComctlLib.ListView, ByVal lngColumnIndex As Long, ByVal lngSortKeyColumnIndex As Long)
    Dim lngIndex As Long
    Dim lngTestValue As Long, lngMaximumValue As Long
    Dim strItemEntry As String
    Dim strFormatString As String, strSortKeyEntry As String

    With lvwThisListView
        ' Sort numerically
        ' Need to Construct the values for the SortKey Column
        
        ' First, determine the maximum value in the column
        For lngIndex = 1 To .ListItems.Count
            With .ListItems(lngIndex)
                If lngColumnIndex = 0 Then
                    strItemEntry = .Text
                Else
                    strItemEntry = .SubItems(lngColumnIndex)
                End If
                
                lngTestValue = Val(strItemEntry)
                If lngTestValue > lngMaximumValue Then lngMaximumValue = lngTestValue
            End With
        Next lngIndex
        
        ' Construct the format string based on the maximum value
        strFormatString = ListViewConstructSortKeyFormatString(lngMaximumValue)
        
        ' strFormatString does not contain format symbols for decimal values; need to add some
        strFormatString = strFormatString & ".0000000#"
        
        ' Assign the values to the SortKey column
        For lngIndex = 1 To .ListItems.Count
            With .ListItems(lngIndex)
                If lngColumnIndex = 0 Then
                    strItemEntry = .Text
                Else
                    strItemEntry = .SubItems(lngColumnIndex)
                End If
                
                strSortKeyEntry = Format(strItemEntry, strFormatString)
                .SubItems(lngSortKeyColumnIndex) = strSortKeyEntry
            End With
        Next lngIndex
    End With

End Sub

Public Sub ListViewHighlightItem(lvwThisListView As MSComctlLib.ListView, lngIndexToHighlight As Long)
    Dim lngIndex As Long
    
    With lvwThisListView
        For lngIndex = 1 To .ListItems.Count
            If lngIndex = lngIndexToHighlight Then
                .ListItems(lngIndex).Selected = True
                .ListItems(lngIndex).EnsureVisible
            Else
                .ListItems(lngIndex).Selected = False
            End If
        Next lngIndex
    End With
End Sub

Public Sub ListViewInvertSelection(lvwThisListView As MSComctlLib.ListView)
    
    Dim lstListItem As MSComctlLib.ListItem
    For Each lstListItem In lvwThisListView.ListItems
        lstListItem.Selected = Not lstListItem.Selected
    Next lstListItem
    
End Sub

Public Sub ListViewSelectAllItems(lvwThisListView As MSComctlLib.ListView)
    
    Dim lstListItem As MSComctlLib.ListItem
    For Each lstListItem In lvwThisListView.ListItems
        lstListItem.Selected = True
    Next lstListItem
    
End Sub

Public Sub ListViewSetFeatures(lvwThisListView As MSComctlLib.ListView, blnIncludeCheckBoxes As Boolean)
    With lvwThisListView
        .AllowColumnReorder = True
    
        ' Use report style
        .View = lvwReport
        .LabelEdit = lvwManual
        .Checkboxes = blnIncludeCheckBoxes
        .GridLines = True
        .FullRowSelect = True
        
        .HideSelection = False
    End With
    
End Sub

Public Sub ListViewShowHideForUpdating(ByRef lvwThisListView As MSComctlLib.ListView, blnHideListView As Boolean)

    ' The following code instructs the ListView control not to proceed on WM_Paint
    ' This greatly increases the speed with which hundreds or thousands of items can be added to a listview
    ' If this command isn't used before adding all of the items, then, regardless of whether the
    '  controls is visible or not, the rate of adding items slows down after a certain amount of data gets added to the listview
    ' The use of this command is not needed for listviews with <100 items
    
    SendMessage lvwThisListView.hwnd, WM_SETREDRAW, Not blnHideListView, 0&
        
    ' Hide the list to increase update speed
    lvwThisListView.Visible = Not blnHideListView
    
    If Not blnHideListView Then
        lvwThisListView.Refresh
    End If

End Sub

Public Sub ListViewSort(ByRef lvwThisListView As MSComctlLib.ListView, ByVal lngColumnIndex As Long, ByRef udtColumnSortFormat As udtColumnSortFormatType)
    
    With lvwThisListView
        If udtColumnSortFormat.SortColumnIndexSaved <> lngColumnIndex Then
            
            Debug.Assert lngColumnIndex < udtColumnSortFormat.ColumnCount
            
            ' Temporarily disable sorting
            .Sorted = False
            
            If udtColumnSortFormat.ColumnSortOrder(lngColumnIndex) = sfText Then
                .SortKey = lngColumnIndex
            Else
                ListViewSortNumerically lvwThisListView, lngColumnIndex, udtColumnSortFormat.SortKeyColumnIndex
                
                ' Assign the index of the SortKey column to .SortKey
                .SortKey = udtColumnSortFormat.SortKeyColumnIndex
            End If
            .SortOrder = lvwAscending
            udtColumnSortFormat.SortColumnIndexSaved = lngColumnIndex
        Else
            ' If the column is already selected then change the
            ' sortorder to be the opposite of what is currently being used
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        End If
        
        'Set the sorted property to use the new sortkey
        'and sort the contents
        .Sorted = True
    End With

End Sub

Public Sub ListViewUpdateRecentIndexHistory(lvwThisListView As MSComctlLib.ListView, ByVal Item As MSComctlLib.ListItem, intListViewID As Integer)
    
    If intListViewID >= 0 And intListViewID <= MAX_LISTVIEW_COUNT Then
        mFindTextIndexHistory(intListViewID).MatchIndex = ListViewGetItemIndex(lvwThisListView, Item)
    End If
    
End Sub

