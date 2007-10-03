VERSION 5.00
Begin VB.Form frmPairBrowser 
   Caption         =   "Pair Browser"
   ClientHeight    =   9285
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   ScaleHeight     =   9285
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraPlotOptions 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Options"
      Height          =   1905
      Left            =   120
      TabIndex        =   14
      Top             =   3600
      Width           =   5175
      Begin VB.ComboBox cboPointShapeHeavy 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox chkShowPointSymbols 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Point Symbols"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.TextBox txtGraphLineWidth 
         Height          =   285
         Left            =   1440
         TabIndex        =   19
         Text            =   "1"
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtGraphPointSize 
         Height          =   285
         Left            =   1440
         TabIndex        =   18
         Text            =   "3"
         Top             =   1080
         Width           =   735
      End
      Begin VB.CheckBox chkShowGridlines 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Gridlines"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkDrawLinesBetweenPoints 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Connect Points with Line"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.ComboBox cboPointShapeLight 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lblPointColorSelectionHeavy 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4560
         TabIndex        =   27
         ToolTipText     =   "Double click to change"
         Top             =   615
         Width           =   375
      End
      Begin VB.Label lblPointInfoLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Heavy Member Point Info"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   26
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblGraphLineWidth 
         BackStyle       =   0  'Transparent
         Caption         =   "Line Width"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1470
         Width           =   975
      End
      Begin VB.Label lblGraphPointSize 
         BackStyle       =   0  'Transparent
         Caption         =   "Point Size"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1110
         Width           =   855
      End
      Begin VB.Label lblPointInfoLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Light Member Point Info"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   22
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label lblPointColorSelectionLight 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4560
         TabIndex        =   21
         ToolTipText     =   "Double click to change"
         Top             =   1335
         Width           =   375
      End
   End
   Begin VIPER.ctlStatusBar ctlStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   8760
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   450
   End
   Begin VB.Frame fraPairInfo 
      Caption         =   "Info on Selected Pair"
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   6855
      Begin VB.Label lblPairInfo2 
         Caption         =   "Info"
         Height          =   615
         Left            =   3600
         TabIndex        =   28
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblPairInfo1 
         Caption         =   "Info"
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.ListBox lstPairs 
      Height          =   1815
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   480
      Width           =   3495
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   2100
      Left            =   3720
      TabIndex        =   1
      Top             =   360
      Width           =   3255
      Begin VB.CheckBox chkSortDescending 
         Caption         =   "Sort Descending"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   3015
      End
      Begin VB.CheckBox chkFilterPairsOnMTHits 
         Caption         =   "Only show pairs with MT Tag hits"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   3015
      End
      Begin VB.ComboBox cboPairSortOrder 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtMassRange 
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Text            =   "0.02"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox cboMassRangeUnits 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1305
         Width           =   975
      End
      Begin VB.TextBox txtScanRange 
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Text            =   "50"
         Top             =   1680
         Width           =   855
      End
      Begin VB.ComboBox cboScanRangeUnits 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1665
         Width           =   975
      End
      Begin VB.Label lblMassRange 
         Caption         =   "Mass range"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblScanRange 
         Caption         =   "Scan range"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VIPER.ctlSpectraPlotter ctlPairsPlot 
      Height          =   5055
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   8916
   End
   Begin VB.Label lblPairsList 
      Caption         =   "Pairs List"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileFindPairsDeltaUMC 
         Caption         =   "&Open Delta (UMC) Find Pairs Window"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSaveChanges 
         Caption         =   "&Save Changes (Delete Pairs)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndoDelete 
         Caption         =   "&Undo last pair deletion"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDeleteSelectedPair 
         Caption         =   "&Delete Selected Pair"
      End
      Begin VB.Menu mnuEditUndeletedSelectedPair 
         Caption         =   "&Include Selected Pair"
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCopyData 
         Caption         =   "&Copy Data"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopyDataLineUpByScan 
         Caption         =   "Copy Data &Lined up by Scan"
      End
      Begin VB.Menu mnuEditSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyChart 
         Caption         =   "Copy Plot as &BMP"
         Index           =   0
      End
      Begin VB.Menu mnuCopyChart 
         Caption         =   "Copy Plot as &WMF"
         Index           =   1
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCopyChart 
         Caption         =   "Copy Plot as &EMF"
         Index           =   2
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsPlotOptions 
         Caption         =   "Plot &Options"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuOptionsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsAutoZoom2D 
         Caption         =   "&Auto zoom 2D plot"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptionsHighlightPairMembers 
         Caption         =   "&Highlight pair members"
         Checked         =   -1  'True
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuOptionsAutoZoomFixedDimensions 
         Caption         =   "Absolute auto zoom &dimensions"
      End
      Begin VB.Menu mnuOptionsSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsPlotAllChargeStates 
         Caption         =   "&Plot all charge states"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuOptionsSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsKeepWindowOnTop 
         Caption         =   "&Keep Window on Top"
         Shortcut        =   ^K
      End
   End
End
Attribute VB_Name = "frmPairBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DELETE_PAIR_INDICATOR = " -- Delete"

Public CallerIDNew As Long
Private CallerIDLoaded As Long          ' 0 if no data is loaded

Private Enum epsPairSortOrderConstants
    epsPairIndex = 0
    epsMass = 1
    epsTime = 2
    epsAbundance = 3
    epsER = 4
End Enum

Private Enum ccmCopyChartMode
    ccmBMP = 0
    ccmWMF = 1
    ccmEMF = 2
End Enum

Private mPairsCount As Long
Private mPairValid() As Boolean                 ' 0-based array; Dereference into GelP_D_L() using mPairInfoOrignalIndex()
Private mPairInfoSortedPointerArray() As Long   ' 0-based array; pointer into mPairValid
Private mPairInfoOrignalIndex() As Long         ' 0-based array; Original index of pair in GelP_D_L(); needed for option to filter to only include pairs with hits
Private mPairsAreFilteredForHitsOnly As Boolean

Private mDeletedPairsStackCount As Long
Private mDeletedPairsStack() As Long            ' 0-based array; Indices of deleted pairs; pointers into mPairValid(); used with the undo command

' The following are 1-based arrays, for compatibility with the plot control
Private mDataPointCountLight As Long
Private mCurrentXDataLight() As Double          ' 1-based; Actually simply holding integer scan numbers, but must be type double to populate the chart
Private mCurrentYDataLight() As Double          ' 1-based

Private mDataPointCountHeavy As Long
Private mCurrentXDataHeavy() As Double          ' 1-based; Actually simply holding integer scan numbers but must be type double to populate the chart
Private mCurrentYDataHeavy() As Double          ' 1-based

Private mWindowStayOnTopEnabled As Boolean

Private mUpdatingControls As Boolean
Private mFormInitialized As Boolean

Public Sub AutoUpdatePlot(Optional blnForceUpdate As Boolean = False)
    PopulateFormWithData blnForceUpdate
End Sub

Private Sub AutoZoom2DPlot(lngPairIndexOriginal As Long)
    ' Note: lngPairIndexOriginal should be looked up from mPairInfoOrignalIndex
    
    Dim lngUMCIndexLight As Long, lngUMCIndexHeavy As Long

On Error GoTo AutoZoom2DPlotErrorHandler

    If lngPairIndexOriginal < 0 Or lngPairIndexOriginal >= GelP_D_L(CallerIDLoaded).PCnt Then
        Exit Sub
    End If

    ' Determine the UMC indices for this pair's members
    With GelP_D_L(CallerIDLoaded).Pairs(lngPairIndexOriginal)
        lngUMCIndexLight = .P1
        lngUMCIndexHeavy = .P2
    End With
    
    BrowseFeaturesZoom2DPlot glbPreferencesExpanded.PairBrowserPlottingOptions, CallerIDLoaded, lngUMCIndexLight, lngUMCIndexHeavy
    
    Exit Sub
    
AutoZoom2DPlotErrorHandler:
    Debug.Assert False
    Me.MousePointer = vbDefault
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error auto zooming: " & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    LogErrors Err.Number, "frmPairBrowser->AutoZoom2DPlot", Err.Description, CallerIDLoaded

End Sub

Private Sub CustomizeMenus()
    mnuEditDeleteSelectedPair.Caption = mnuEditDeleteSelectedPair.Caption & vbTab & "Delete Key"
    mnuEditUndeletedSelectedPair.Caption = mnuEditUndeletedSelectedPair.Caption & vbTab & "Insert Key"
End Sub

Private Function DeleteMarkedPairs(Optional blnInformIfNoneToDelete As Boolean = True) As Boolean
    ' Returns true if the pairs were deleted (or no deleted pairs exist)
    
    Dim eResponse As VbMsgBoxResult
    Dim lngPairCountToDelete As Long
    Dim lngPairIndex As Long
    
    Dim blnOriginalPairEntryExamined() As Boolean
    
    Dim lngNewPairCount As Long
    Dim udtNewPairList() As udtIsoPairsDetailsType
    
    Dim blnSuccess As Boolean

On Error GoTo DeleteMarkedPairsErrorHandler

    For lngPairIndex = 0 To mPairsCount - 1
        If Not mPairValid(lngPairIndex) Then
            lngPairCountToDelete = lngPairCountToDelete + 1
        End If
    Next lngPairIndex
    
    blnSuccess = True
    If lngPairCountToDelete > 0 And CallerIDLoaded > 0 Then
        eResponse = MsgBox("You have marked " & Trim(mDeletedPairsStackCount) & " pairs for deletion (" & Trim(GelP_D_L(CallerIDLoaded).PCnt - lngPairCountToDelete) & " will remain).  Choose Yes to proceed with deletion.", vbQuestion + vbYesNoCancel, "Delete pairs")
        
        If eResponse = vbYes Then
            ' Delete marked pairs
            ' Cannot copy in place since must use mPairInfoOrignalIndex() pointer array
            With GelP_D_L(CallerIDLoaded)
            
                lngNewPairCount = 0
                ReDim udtNewPairList(.PCnt - 1)
                ReDim blnOriginalPairEntryExamined(.PCnt - 1)

                For lngPairIndex = 0 To mPairsCount - 1
                    blnOriginalPairEntryExamined(mPairInfoOrignalIndex(lngPairIndex)) = True
                    If mPairValid(lngPairIndex) Then
                        udtNewPairList(lngNewPairCount) = .Pairs(mPairInfoOrignalIndex(lngPairIndex))
                        lngNewPairCount = lngNewPairCount + 1
                    End If
                Next lngPairIndex
                
                ' If mPairsAreFilteredForHitsOnly is True, then also need to copy the pairs that do not have any MT tag hits
                If mPairsAreFilteredForHitsOnly Then
                    For lngPairIndex = 0 To .PCnt - 1
                        If Not blnOriginalPairEntryExamined(lngPairIndex) Then
                            Debug.Assert Not IsAMTReferencedByUMC(GelUMC(CallerIDLoaded).UMCs(.Pairs(lngPairIndex).P1), CallerIDLoaded)
                            
                            udtNewPairList(lngNewPairCount) = .Pairs(lngPairIndex)
                            blnOriginalPairEntryExamined(lngPairIndex) = True
                            lngNewPairCount = lngNewPairCount + 1
                        End If
                    Next lngPairIndex
                End If
                
                .PCnt = lngNewPairCount
                If .PCnt > 0 Then
                    ReDim Preserve udtNewPairList(.PCnt - 1)
                Else
                    ReDim Preserve udtNewPairList(0)
                End If
                .Pairs = udtNewPairList
                
            End With
        Else
            blnSuccess = False
        End If
    Else
        If blnInformIfNoneToDelete And CallerIDLoaded > 0 Then
            MsgBox "No pairs have been marked for deletion.", vbInformation + vbOKOnly, "Nothing to do"
        End If
    End If
    
    DeleteMarkedPairs = blnSuccess
    Exit Function
    
DeleteMarkedPairsErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error deleting marked pairs: " & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    
    LogErrors Err.Number, "DeleteMarkedPairs", Err.Description, CallerIDLoaded

End Function

Private Sub DeleteSelectedPair()
    
    If lstPairs.ListIndex < 0 Or CallerIDLoaded <= 0 Then
        ' Nothing selected
    Else
        Me.MousePointer = vbHourglass
    
        If BrowseFeaturesDeleteSelected(lstPairs, mPairInfoSortedPointerArray(), mPairValid(), mDeletedPairsStackCount, mDeletedPairsStack()) Then
            UpdateListboxCaptionsSelected
        End If
    
        Me.MousePointer = vbDefault
    End If

End Sub

Private Sub DisplayPairInfoSelectedItem(Optional blnSkipAutoZoom As Boolean = False)
    Dim strDescription As String
    Dim strDescriptionAddnl As String
    Dim lngPairIndexDereferenced As Long
    Dim lngPairIndexOriginal As Long
    
    If mUpdatingControls Then Exit Sub
    
    If lstPairs.ListIndex < 0 Or CallerIDLoaded <= 0 Then
        strDescription = "Pair not selected"
        strDescriptionAddnl = ""
    Else
        lngPairIndexDereferenced = mPairInfoSortedPointerArray(lstPairs.ListIndex)
        lngPairIndexOriginal = mPairInfoOrignalIndex(lngPairIndexDereferenced)
        strDescription = GeneratePairDescription(lngPairIndexDereferenced, True, strDescriptionAddnl)
        
        UpdatePlotForPair lngPairIndexOriginal
        
        If glbPreferencesExpanded.PairBrowserPlottingOptions.AutoZoom2DPlot And Not blnSkipAutoZoom Then
            AutoZoom2DPlot lngPairIndexOriginal
        End If
    End If

    lblPairInfo1 = strDescription
    lblPairInfo2 = strDescriptionAddnl
End Sub

Private Sub DisplayPairsPopulateListbox(blnResortedData As Boolean)

    Dim lngIndex As Long
    Dim lngIndexSaved As Long
    Dim lngTopIndexSaved As Long
    
    Dim intCompareLen As Integer
    
    Dim strCaptionSaved As String
    Dim strCaption As String
    
    lngIndexSaved = lstPairs.ListIndex
    lngTopIndexSaved = lstPairs.TopIndex
    
    If blnResortedData And lngIndexSaved >= 0 Then
        strCaptionSaved = lstPairs.List(lngIndexSaved)
        If Right(strCaptionSaved, Len(DELETE_PAIR_INDICATOR)) = DELETE_PAIR_INDICATOR Then
            strCaptionSaved = Left(strCaptionSaved, Len(strCaptionSaved) - Len(DELETE_PAIR_INDICATOR))
        End If
        intCompareLen = Len(strCaptionSaved)
    End If
    
    lstPairs.Clear
    
    If blnResortedData Then
        For lngIndex = 0 To mPairsCount - 1
            strCaption = GeneratePairDescription(mPairInfoSortedPointerArray(lngIndex), False)
            lstPairs.AddItem strCaption
            
            If intCompareLen > 0 Then
                If Left(strCaption, intCompareLen) = strCaptionSaved Then
                    lngIndexSaved = lngIndex
                End If
            End If
        Next lngIndex
    Else
        For lngIndex = 0 To mPairsCount - 1
            lstPairs.AddItem mPairInfoOrignalIndex(mPairInfoSortedPointerArray(lngIndex), False)
        Next lngIndex
    End If
    
    If Not blnResortedData Then
        If lngIndexSaved >= lngTopIndexSaved Then
            lstPairs.TopIndex = lngTopIndexSaved
        End If
    End If
    
    If lngIndexSaved < 0 Then
        If lstPairs.ListCount > 0 Then lstPairs.ListIndex = 0
    ElseIf lngIndexSaved < lstPairs.ListCount Then
        lstPairs.ListIndex = lngIndexSaved
    Else
        lstPairs.ListIndex = lstPairs.ListCount - 1
    End If
        
    If blnResortedData Then
        If lstPairs.ListIndex > 0 Then
            lstPairs.TopIndex = lstPairs.ListIndex - 1
        End If
    End If

End Sub

Private Function ExportPlotDataToClipboardOrFile(blnLineUpByScan As Boolean, Optional strFilePath As String = "", Optional blnShowMessages As Boolean = True) As Long
    ' Returns 0 if success, the error code if an error

    Dim lngCombinedDataCount As Long
    Dim lngCombinedScanData() As Long           ' 0-based, 1D array
    Dim dblCombinedAbuData() As Double          ' 0-based, 2D array
    
    Dim strData() As String                     ' 0-based array
    Dim strTextToCopy As String
    
    Dim lngIndex As Long, lngTargetIndex As Long
    Dim lngScanMin As Long, lngScanMax As Long
    
    Dim OutFileNum As Integer
    
    If mDataPointCountLight = 0 And mDataPointCountHeavy = 0 Then
        If blnShowMessages And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "No data found to copy", vbInformation + vbOKOnly, "No data"
        End If
        ExportPlotDataToClipboardOrFile = -1
        Exit Function
    End If
    
On Error GoTo ExportPlotDataToClipboardOrFileErrorHandler

    Me.MousePointer = vbHourglass

    If blnLineUpByScan Then
        ' Examine mCurrentXDataLight and mCurrentXDataHeavy to determine the scan range
        ' copy the data into the output array, then populate strData
        
        ' Determine the minimum and maximum scan numbers
        ' We can assume mCurrentXDataLight and mCurrentXDataHeavy are sorted ascending
        lngScanMin = mCurrentXDataLight(1)
        If lngScanMin > mCurrentXDataHeavy(1) Then
            lngScanMin = mCurrentXDataHeavy(1)
        End If
        
        lngScanMax = mCurrentXDataLight(mDataPointCountLight)
        If lngScanMax < mCurrentXDataHeavy(mDataPointCountHeavy) Then
            lngScanMax = mCurrentXDataHeavy(mDataPointCountHeavy)
        End If
        
        lngCombinedDataCount = lngScanMax - lngScanMin + 1
        If lngCombinedDataCount < 1 Then lngCombinedDataCount = 1
        ReDim lngCombinedScanData(lngCombinedDataCount - 1)
        ReDim dblCombinedAbuData(lngCombinedDataCount - 1, 1)
        
        ' Populate lngCombinedScanData
        For lngIndex = 0 To lngCombinedDataCount - 1
            lngCombinedScanData(lngIndex) = lngIndex + lngScanMin
        Next lngIndex
        
        ' Populate dblCombinedAbuData
        For lngIndex = 1 To mDataPointCountLight
            lngTargetIndex = Round(mCurrentXDataLight(lngIndex), 0) - lngScanMin
            If lngTargetIndex >= 0 And lngTargetIndex < lngCombinedDataCount Then
                If Round(mCurrentXDataLight(lngIndex), 0) = lngCombinedScanData(lngTargetIndex) Then
                    dblCombinedAbuData(lngTargetIndex, 0) = mCurrentYDataLight(lngIndex)
                Else
                    ' This shouldn't happen
                    Debug.Assert False
                End If
            Else
                ' This shouldn't happen
                Debug.Assert False
            End If
        Next lngIndex
        
        For lngIndex = 1 To mDataPointCountHeavy
            lngTargetIndex = Round(mCurrentXDataHeavy(lngIndex), 0) - lngScanMin
            If lngTargetIndex >= 0 And lngTargetIndex < lngCombinedDataCount Then
                If Round(mCurrentXDataHeavy(lngIndex), 0) = lngCombinedScanData(lngTargetIndex) Then
                    dblCombinedAbuData(lngTargetIndex, 1) = mCurrentYDataHeavy(lngIndex)
                Else
                    ' This shouldn't happen
                    Debug.Assert False
                End If
            Else
                ' This shouldn't happen
                Debug.Assert False
            End If
        Next lngIndex
        
        
        ' Header row is strData(0), data starts at strData(1)
        ReDim strData(0 To lngCombinedDataCount)
        
        ' Fill strData()
        ' Define the header row
        strData(0) = "Scan Number" & vbTab & "Light Abu" & vbTab & "Heavy Abu"
        
        For lngIndex = 0 To lngCombinedDataCount - 1
            strData(lngIndex + 1) = lngCombinedScanData(lngIndex) & vbTab & dblCombinedAbuData(lngIndex, 0) & vbTab & dblCombinedAbuData(lngIndex, 1)
        Next lngIndex
        lngCombinedDataCount = lngCombinedDataCount + 1
        
    Else
        ' Not lining up by scan
        
        lngCombinedDataCount = mDataPointCountLight
        If mDataPointCountHeavy > lngCombinedDataCount Then
            lngCombinedDataCount = mDataPointCountHeavy
        End If
        
        ' Header row is strData(0), data starts at strData(1)
        ReDim strData(0 To lngCombinedDataCount)
        
        ' Fill strData()
        ' Define the header row
        strData(0) = "Scan Number Light" & vbTab & "Light Abu" & vbTab & "Scan Number Heavy" & vbTab & "Heavy Abu"
        
        For lngIndex = 1 To lngCombinedDataCount
            If lngIndex <= mDataPointCountLight Then
                strData(lngIndex) = Round(mCurrentXDataLight(lngIndex), 0) & vbTab & mCurrentYDataLight(lngIndex) & vbTab
            Else
                strData(lngIndex) = vbTab & vbTab
            End If
            
            If lngIndex <= mDataPointCountHeavy Then
                strData(lngIndex) = strData(lngIndex) & Round(mCurrentXDataHeavy(lngIndex), 0) & vbTab & mCurrentYDataHeavy(lngIndex)
            Else
                strData(lngIndex) = strData(lngIndex) & vbTab
            End If
        Next lngIndex
        lngCombinedDataCount = lngCombinedDataCount + 1
    End If
    
    If Len(strFilePath) > 0 Then
        OutFileNum = FreeFile()
        Open strFilePath For Output As #OutFileNum
        
        For lngIndex = 0 To lngCombinedDataCount - 1
            Print #OutFileNum, strData(lngIndex)
        Next lngIndex
        
        Close #OutFileNum
    Else
        strTextToCopy = FlattenStringArray(strData(), lngCombinedDataCount, vbCrLf, False)
        Clipboard.Clear
        Clipboard.SetText strTextToCopy, vbCFText
    End If
    
    Me.MousePointer = vbDefault
    
    ExportPlotDataToClipboardOrFile = 0
    Exit Function

ExportPlotDataToClipboardOrFileErrorHandler:
    If blnShowMessages And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error exporting plot data: " & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    
    ExportPlotDataToClipboardOrFile = Err.Number
    
End Function

Private Sub FilterPairsOnMTHits(blnEnableFilter As Boolean)

    Static blnUpdating As Boolean
    
    If blnUpdating Then Exit Sub
    blnUpdating = True
    
    If Not DeleteMarkedPairs(False) Then
        ' Unsaved changes exist, and the user cancelled saving changes
        ' Do not enable the filter
        SetCheckBox chkFilterPairsOnMTHits, Not blnEnableFilter
        blnUpdating = False
        Exit Sub
    End If
    
    SetCheckBox chkFilterPairsOnMTHits, blnEnableFilter
    PopulateFormWithData True
    
    blnUpdating = False
End Sub

Private Function GeneratePairDescription(lngPairIndexDereferenced As Long, blnVerbose As Boolean, Optional ByRef strDescriptionAddnl As String) As String
    ' Note: lngPairIndexDereferenced should point into mPairValid
    
    Dim strDescription As String
    Dim lngScanMin As Long, lngScanMax As Long
    Dim dblMassMin As Double, dblMassMax As Double
    Dim udtUMCLight As udtUMCType
    Dim udtUMCHeavy As udtUMCType
    Dim lngPairIndexOriginal
    
    Dim strDBMatchList As String
    
    Dim lngIndex As Long
    Dim lngAMTIDCnt As Long
    Dim strAMTIDs() As String
    
On Error GoTo GeneratePairDescriptionErrorHandler

    lngPairIndexOriginal = mPairInfoOrignalIndex(lngPairIndexDereferenced)
    
    With GelUMC(CallerIDLoaded).UMCs(GelP_D_L(CallerIDLoaded).Pairs(lngPairIndexOriginal).P1)
        strDBMatchList = FixNull(GelData(CallerIDLoaded).IsoData(.ClassRepInd).MTID)
        
        ' Extract just the MTID's from this list
        lngAMTIDCnt = GetAMTRefFromString1(strDBMatchList, strAMTIDs())
    End With
    
    With GelP_D_L(CallerIDLoaded).Pairs(lngPairIndexOriginal)
    
        udtUMCLight = GelUMC(CallerIDLoaded).UMCs(.P1)
        udtUMCHeavy = GelUMC(CallerIDLoaded).UMCs(.P2)
    
        If Not blnVerbose Then
            strDescription = "#" & Trim(lngPairIndexOriginal) & ", "
            strDescription = strDescription & Round(udtUMCLight.ClassMW, 1) & " Da, "
            strDescription = strDescription & "ER = " & Round(.ER, 2)
            
            If Not mPairValid(lngPairIndexDereferenced) Then
                strDescription = strDescription & DELETE_PAIR_INDICATOR
            End If
            
            strDescriptionAddnl = ""
        Else
            strDescription = ""
            strDescription = strDescription & "LC-MS Features " & .P1 & " and " & .P2 & vbCrLf
            strDescription = strDescription & "Abundances " & DoubleToStringScientific(udtUMCLight.ClassAbundance, 3)
            strDescription = strDescription & " and " & DoubleToStringScientific(udtUMCHeavy.ClassAbundance, 3) & vbCrLf
            
            ' Append the AMT matches, if any
            If lngAMTIDCnt > 0 Then
                strDescription = strDescription & "MTIDs: "
                For lngIndex = 1 To lngAMTIDCnt
                    strDescription = strDescription & strAMTIDs(lngIndex)
                    If lngIndex < lngAMTIDCnt Then
                        strDescription = strDescription & "; "
                    End If
                Next lngIndex
                strDescription = strDescription & vbCrLf
            End If

            ' Construct the additional description
            BrowseFeaturesLookupScanAndMassLimits udtUMCLight, udtUMCHeavy, lngScanMin, lngScanMax, dblMassMin, dblMassMax
            
            strDescriptionAddnl = ""
            strDescriptionAddnl = strDescriptionAddnl & "Scan range " & lngScanMin & " to " & lngScanMax & vbCrLf
            strDescriptionAddnl = strDescriptionAddnl & "Mass range " & Round(dblMassMin, 4) & " to " & Round(dblMassMax, 4)
        End If
    End With
    
    GeneratePairDescription = strDescription
    Exit Function

GeneratePairDescriptionErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmPairBrowser->GeneratePairDescription", Err.Description, CallerIDLoaded
    If Len(strDescription) = 0 Then
        strDescription = "Error forming description"
    Else
        strDescription = strDescription & "; Error encountered"
    End If
    
End Function

Public Sub InitializeForm()
    mUpdatingControls = True
    
On Error GoTo InitializeFormErrorHandler

    CustomizeMenus
    
    ' Update the controls with the values in .PairBrowserPlottingOptions
    With glbPreferencesExpanded.PairBrowserPlottingOptions
        cboPairSortOrder.ListIndex = .SortOrder
        SetCheckBox chkSortDescending, .SortDescending
        
        mnuOptionsAutoZoom2D.Checked = .AutoZoom2DPlot
        mnuOptionsHighlightPairMembers.Checked = .HighlightMembers
        mnuOptionsPlotAllChargeStates.Checked = .PlotAllChargeStates
        
        SetFixedDimensionsForAutoZoom .FixedDimensionsForAutoZoom
        cboMassRangeUnits.ListIndex = .MassRangeUnits
        txtMassRange = .MassRangeZoom
                
        cboScanRangeUnits.ListIndex = .ScanRangeUnits
        txtScanRange = .ScanRangeZoom
        
        With .Graph2DOptions
            SetCheckBox chkShowPointSymbols, .ShowPointSymbols
            SetCheckBox chkDrawLinesBetweenPoints, .DrawLinesBetweenPoints
            SetCheckBox chkShowGridlines, .ShowGridLines
            
            txtGraphPointSize = Trim(.PointSizePixels)
            
            If .PointShape < 1 Or .PointShape > OlectraChart2D.ShapeConstants.oc2dShapeSquare Then
                .PointShape = OlectraChart2D.ShapeConstants.oc2dShapeDot
            End If
            cboPointShapeLight.ListIndex = .PointShape - 1
            lblPointColorSelectionLight.BackColor = .PointAndLineColor
            
            txtGraphLineWidth = Trim(.LineWidthPixels)
        End With
    
        If .PointShapeHeavy < 1 Or .PointShapeHeavy > OlectraChart2D.ShapeConstants.oc2dShapeSquare Then
            .PointShapeHeavy = OlectraChart2D.ShapeConstants.oc2dShapeDot
        End If
        cboPointShapeHeavy.ListIndex = .PointShapeHeavy - 1
        lblPointColorSelectionHeavy.BackColor = .PointAndLineColorHeavy
    
        ToggleWindowStayOnTop .KeepWindowOnTop
    End With
    mUpdatingControls = False
    
    PopulateFormWithData True
    
    mFormInitialized = True
    Exit Sub

InitializeFormErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmPairBrowser->InitializeForm", Err.Description, CallerIDLoaded
    Resume Next
    
End Sub

Private Sub InitializePlot()
    
    Dim dblBlankDataX(1 To 1) As Double
    Dim dblBlankDataY(1 To 1) As Double
    
    With ctlPairsPlot
        .PopulateSymbolStyleComboBox cboPointShapeLight
        .PopulateSymbolStyleComboBox cboPointShapeHeavy
        
        .EnableDisableDelayUpdating True
        .SetCurrentGroup 2
        .SetSeriesCount 0

        .SetCurrentGroup 1
        .SetSeriesCount 2
        .SetCurrentSeries 1

        .SetSeriesDataPointCount 1, 1
        .SetDataX 1, dblBlankDataX()
        .SetDataY 1, dblBlankDataY()

        .SetSeriesDataPointCount 2, 1
        .SetDataX 2, dblBlankDataX()
        .SetDataY 2, dblBlankDataY()

        .EnableDisableDelayUpdating False
    End With
End Sub

Private Sub PopulateComboBoxes()
    
    mUpdatingControls = True
    With cboMassRangeUnits
        .Clear
        .AddItem "Da"
        .AddItem "ppm"
        .ListIndex = mruDa
    End With
    
    With cboScanRangeUnits
        .Clear
        .AddItem "Scan #"
        .AddItem "NET"
        .ListIndex = sruScan
    End With
    
    With cboPairSortOrder
        .Clear
        .AddItem "Sort by Pair Index"
        .AddItem "Sort by Mass"
        .AddItem "Sort by Time"
        .AddItem "Sort by Abundance"
        .AddItem "Sort by ER"
        .ListIndex = epsER
    End With
    mUpdatingControls = False
    
End Sub

Private Sub PopulateFormWithData(Optional blnForcePopulation As Boolean = False)
    
    Dim blnCallerIDValid As Boolean
    Dim lngPairIndex As Long
    Dim blnAddPair As Boolean
    
On Error GoTo PopulateControlsErrorHandler

    blnCallerIDValid = False
    If CallerIDNew <> CallerIDLoaded Or blnForcePopulation Then
        
        CallerIDLoaded = CallerIDNew
        
        Me.Caption = "Pairs Browser: " & GelBody(CallerIDLoaded).Caption
        
        If CallerIDLoaded > UBound(GelP_D_L) Then
            CallerIDLoaded = UBound(GelP_D_L)
        End If
        
        With GelP_D_L(CallerIDLoaded)
            blnCallerIDValid = True
            
            mPairsAreFilteredForHitsOnly = cChkBox(chkFilterPairsOnMTHits)
            
            If .PCnt <= 0 Then
                ReDim mPairValid(0)
                mPairsCount = 0
            Else
                ReDim mPairValid(.PCnt - 1)
                ReDim mPairInfoOrignalIndex(.PCnt - 1)
            
                If mPairsAreFilteredForHitsOnly Then
                    mPairsCount = 0
                    For lngPairIndex = 0 To .PCnt - 1
                        
                        With .Pairs(lngPairIndex)
                            If IsAMTReferencedByUMC(GelUMC(CallerIDLoaded).UMCs(.P1), CallerIDLoaded) Then
                                blnAddPair = True
                            Else
                                blnAddPair = False
                            End If
                            
                            If blnAddPair Then
                                mPairValid(mPairsCount) = True
                                mPairInfoOrignalIndex(mPairsCount) = lngPairIndex
                                mPairsCount = mPairsCount + 1
                            End If
                        End With
                    Next lngPairIndex
                    
                Else
                    mPairsCount = .PCnt
                    For lngPairIndex = 0 To .PCnt - 1
                        mPairValid(lngPairIndex) = True
                        mPairInfoOrignalIndex(lngPairIndex) = lngPairIndex
                    Next lngPairIndex
                End If
            End If
            
            mDeletedPairsStackCount = 0
            ReDim mDeletedPairsStack(0)
        End With
        
        SortAndDisplayPairs
    End If

Exit Sub

PopulateControlsErrorHandler:
    If Not blnCallerIDValid Then
        CallerIDLoaded = 0
    Else
        Debug.Assert False
    End If
    
End Sub

Private Sub PositionControls()
    Dim lngDesiredValue As Long
    
    With ctlStatus
        lngDesiredValue = Me.ScaleHeight - .Height - 30
        If lngDesiredValue < 2040 Then
            lngDesiredValue = 2040
        End If
        .Top = lngDesiredValue
        .Left = 120
        lngDesiredValue = Me.ScaleWidth - .Left - 30
        If lngDesiredValue < 1020 Then lngDesiredValue = 1020
        .width = lngDesiredValue
    End With
    
    With lstPairs
        lngDesiredValue = Me.ScaleHeight * 0.2
        If lngDesiredValue < 2100 Then lngDesiredValue = 2100
        .Height = lngDesiredValue
        
        lngDesiredValue = .Top + .Height + 60
        If lngDesiredValue < fraOptions.Top + fraOptions.Height Then
            lngDesiredValue = fraOptions.Top + fraOptions.Height
        End If
        fraPairInfo.Top = lngDesiredValue
    End With
    
    With ctlPairsPlot
        .Top = fraPairInfo.Top + fraPairInfo.Height + 60
        lngDesiredValue = Me.ScaleWidth - .Left - 120
        If lngDesiredValue < 2040 Then lngDesiredValue = 2040
        .width = lngDesiredValue
        
        lngDesiredValue = ctlStatus.Top - .Top - 60
        If lngDesiredValue < 2040 Then lngDesiredValue = 2040
        .Height = lngDesiredValue
        
        fraPlotOptions.Top = .Top
        fraPlotOptions.Left = .Left
    End With
    
End Sub

Private Function QueryUnloadForm() As Boolean
    ' Returns True if safe to unload; false otherwise
    Dim blnDeleted As Boolean
    Dim eResponse As VbMsgBoxResult
    
    blnDeleted = DeleteMarkedPairs(False)
    If Not blnDeleted Then
        eResponse = MsgBox("Close window and lose all changes?", vbQuestion + vbYesNoCancel, "Close window")
    Else
        eResponse = vbYes
    End If
    
    If eResponse = vbYes Then
        QueryUnloadForm = True
    Else
        QueryUnloadForm = False
    End If
End Function

Private Sub ShowHideOptions(Optional blnForceHide As Boolean)
    If blnForceHide Then
        fraPlotOptions.Visible = False
    Else
        fraPlotOptions.Visible = Not fraPlotOptions.Visible
    End If
    
    mnuOptionsPlotOptions.Checked = fraPlotOptions.Visible
End Sub

Private Sub SortAndDisplayPairs()
    Dim dblDataToSort() As Double
    Dim lngIndex As Long
    Dim lngUMCIndexOriginal As Long
    
    Dim objQSDouble As QSDouble
    Dim blnPerformSort As Boolean
    Dim blnSuccess As Boolean
    
On Error GoTo SortAndDisplayPairsErrorHandler

    If mPairsCount > 0 Then
        
        Me.MousePointer = vbHourglass
        DoEvents
        
        ReDim dblDataToSort(mPairsCount - 1)
        ReDim mPairInfoSortedPointerArray(mPairsCount - 1)
        For lngIndex = 0 To mPairsCount - 1
            mPairInfoSortedPointerArray(lngIndex) = lngIndex
        Next lngIndex
        
        With GelP_D_L(CallerIDLoaded)
            blnPerformSort = True
            Select Case cboPairSortOrder.ListIndex
            Case epsMass
                For lngIndex = 0 To mPairsCount - 1
                    dblDataToSort(lngIndex) = GelUMC(CallerIDLoaded).UMCs(.Pairs(mPairInfoOrignalIndex(lngIndex)).P1).ClassMW
                Next lngIndex
            Case epsTime
                For lngIndex = 0 To mPairsCount - 1
                    lngUMCIndexOriginal = .Pairs(mPairInfoOrignalIndex(lngIndex)).P1
                    With GelUMC(CallerIDLoaded).UMCs(lngUMCIndexOriginal)
                        Select Case .ClassRepType
                        Case gldtCS
                            dblDataToSort(lngIndex) = GelData(CallerIDLoaded).CSData(.ClassRepInd).ScanNumber
                        Case gldtIS
                            dblDataToSort(lngIndex) = GelData(CallerIDLoaded).IsoData(.ClassRepInd).ScanNumber
                        End Select
                    End With
                Next lngIndex
            Case epsAbundance
                For lngIndex = 0 To mPairsCount - 1
                    ' Sum of light and heavy member abundances
                    dblDataToSort(lngIndex) = GelUMC(CallerIDLoaded).UMCs(.Pairs(mPairInfoOrignalIndex(lngIndex)).P1).ClassAbundance
                    dblDataToSort(lngIndex) = dblDataToSort(lngIndex) + GelUMC(CallerIDLoaded).UMCs(.Pairs(mPairInfoOrignalIndex(lngIndex)).P2).ClassAbundance
                Next lngIndex
            Case epsER
                For lngIndex = 0 To mPairsCount - 1
                    dblDataToSort(lngIndex) = .Pairs(mPairInfoOrignalIndex(lngIndex)).ER
                Next lngIndex
            Case Else
                ' Includes epsPairIndex
                ' Nothing to sort
                blnPerformSort = False
            End Select
        End With
        
        If blnPerformSort Then
            Set objQSDouble = New QSDouble
            If glbPreferencesExpanded.PairBrowserPlottingOptions.SortDescending Then
                blnSuccess = objQSDouble.QSDesc(dblDataToSort, mPairInfoSortedPointerArray)
            Else
                blnSuccess = objQSDouble.QSAsc(dblDataToSort, mPairInfoSortedPointerArray)
            End If
            
            If Not blnSuccess Then
                ' Error performing sort
                Debug.Assert False
                MsgBox "Error sorting pairs: " & Err.Description, vbExclamation + vbOKOnly, "Error"
                LogErrors Err.Number, "frmPairBrowser->SortAndDisplayPairs", Err.Description, CallerIDLoaded
            End If
        End If
        
        DisplayPairsPopulateListbox True
    
    Else
        lstPairs.Clear
    End If
   
    Me.MousePointer = vbDefault
    Exit Sub

SortAndDisplayPairsErrorHandler:
    Me.MousePointer = vbDefault
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error sorting pairs and populating list: " & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    LogErrors Err.Number, "SortAndDisplayPairs", Err.Description, CallerIDLoaded

End Sub

Private Sub SelectCustomColor(lblThisLabel As Label)
    Dim lngTemporaryColor As Long
    
    lngTemporaryColor = lblThisLabel.BackColor
    Call GetColorAPIDlg(Me.hwnd, lngTemporaryColor)
    If lngTemporaryColor >= 0 Then
        lblThisLabel.BackColor = lngTemporaryColor
    End If
End Sub

Private Sub SetFixedDimensionsForAutoZoom(blnEnabled As Boolean)
    Dim dblDeltaMW As Double
    
    With glbPreferencesExpanded.PairBrowserPlottingOptions
        .FixedDimensionsForAutoZoom = blnEnabled
        mnuOptionsAutoZoomFixedDimensions.Checked = blnEnabled
    End With
    
    If blnEnabled Then
        lblMassRange = "Mass range"
        lblScanRange = "Scan range"
        
        If mFormInitialized Then
            dblDeltaMW = 5
            If CallerIDLoaded > 0 Then
                On Error Resume Next
                With GelP_D_L(CallerIDLoaded).SearchDef
                    If GelP_D_L(CallerIDLoaded).SearchDef.DeltaCountMax > 1 Then
                        dblDeltaMW = Int(.DeltaCountMax * .DeltaMass) + 1
                    Else
                        dblDeltaMW = Int(.DeltaMass) + 1
                    End If
                End With
            End If
            
            If dblDeltaMW < 4 Then dblDeltaMW = 4
            
            glbPreferencesExpanded.PairBrowserPlottingOptions.MassRangeZoom = dblDeltaMW
            txtMassRange = dblDeltaMW
            cboMassRangeUnits.ListIndex = mruDa
        End If
    Else
        lblMassRange = "Mass edge"
        lblScanRange = "Scan edge"
    
        If mFormInitialized Then
            dblDeltaMW = 50
            glbPreferencesExpanded.PairBrowserPlottingOptions.MassRangeZoom = dblDeltaMW
            
            txtMassRange = Trim(dblDeltaMW)
            cboMassRangeUnits.ListIndex = mruPpm
        End If
    End If
    
End Sub

Private Sub ToggleWindowStayOnTop(blnEnableStayOnTop As Boolean)
    
    mnuOptionsKeepWindowOnTop.Checked = blnEnableStayOnTop
    glbPreferencesExpanded.PairBrowserPlottingOptions.KeepWindowOnTop = blnEnableStayOnTop
    
    If mWindowStayOnTopEnabled = blnEnableStayOnTop Then Exit Sub
    
    Me.ScaleMode = vbTwips
    
    WindowStayOnTop Me.hwnd, blnEnableStayOnTop, Me.ScaleX(Me.Left, vbTwips, vbPixels), Me.ScaleY(Me.Top, vbTwips, vbPixels), Me.ScaleX(Me.width, vbTwips, vbPixels), Me.ScaleY(Me.Height, vbTwips, vbPixels)
    
    mWindowStayOnTopEnabled = blnEnableStayOnTop

End Sub

Private Sub UnDeleteSelectedPair()
    Dim blnUpdateListBox As Boolean

    If lstPairs.ListIndex < 0 Or CallerIDLoaded <= 0 Then
        ' Nothing selected
    Else
        Me.MousePointer = vbHourglass
        
        blnUpdateListBox = BrowseFeaturesUndeleteSelected(lstPairs, mPairInfoSortedPointerArray(), mPairValid(), mDeletedPairsStackCount, mDeletedPairsStack())
        
        If blnUpdateListBox Then
            UpdateListboxCaptionsSelected
        End If
    
    End If
    
    Me.MousePointer = vbDefault

End Sub

Private Sub UpdateListboxCaptionsSelected()
    Dim lngIndex As Long
    
On Error GoTo UpdateListboxCaptionsSelectedErrorHandler

    ' Update the caption for each item selected in lstPairs
    For lngIndex = 0 To lstPairs.ListCount - 1
        If lstPairs.Selected(lngIndex) Then
            lstPairs.List(lngIndex) = GeneratePairDescription(mPairInfoSortedPointerArray(lngIndex), False)
        End If
    Next lngIndex

Exit Sub

UpdateListboxCaptionsSelectedErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmPairBrowser->UpdateListboxCaptionsSelected", Err.Description, CallerIDLoaded
    DisplayPairsPopulateListbox True
    
End Sub

Private Sub UpdatePlotForPair(lngPairIndexOriginal As Long)
    ' Note: lngPairIndexOriginal should be looked up from mPairInfoOrignalIndex
    
    Dim lngUMCIndexLight As Long
    Dim lngUMCIndexHeavy As Long
    
    Dim intChargesUsedCount As Integer
    Dim intChargesUsed() As Integer
    Dim intChargeIndex As Integer
    
    Dim strTitle As String
    Dim udtGraphOptions As udtGraph2DOptionsType
    Dim lngLightColor As Long
    Dim lngHeavyColor As Long
    
    Dim blnUseMaxValueEachScan As Boolean
    
On Error GoTo UpdatePlotForPairErrorHandler

    ' Look up the charges used to compute this Expression Ratio
    With GelP_D_L(CallerIDLoaded)
        ' Convention from Pairs.bas->CalcDltLblPairsERScanByScan is the following:
        blnUseMaxValueEachScan = .SearchDef.UseIdenticalChargesForER
        
        With .Pairs(lngPairIndexOriginal)
            lngUMCIndexLight = .P1
            lngUMCIndexHeavy = .P2
        
            If .ERChargeStateBasisCount > 0 Then
                ' Copy the .ERChargesUsed() array to intChargesUsed()
                intChargesUsed = .ERChargesUsed
            Else
                ReDim intChargesUsed(0)
                intChargesUsed(0) = 0
            End If
        End With
    End With
    
    BrowseFeaturesPopulateUMCPlotData CallerIDLoaded, glbPreferencesExpanded.PairBrowserPlottingOptions.PlotAllChargeStates, intChargesUsed(), lngUMCIndexLight, mDataPointCountLight, mCurrentXDataLight(), mCurrentYDataLight(), blnUseMaxValueEachScan
    BrowseFeaturesPopulateUMCPlotData CallerIDLoaded, glbPreferencesExpanded.PairBrowserPlottingOptions.PlotAllChargeStates, intChargesUsed(), lngUMCIndexHeavy, mDataPointCountHeavy, mCurrentXDataHeavy(), mCurrentYDataHeavy(), blnUseMaxValueEachScan
    
    With ctlPairsPlot
        .EnableDisableDelayUpdating True
        
        strTitle = "Pair #" & Trim(lngPairIndexOriginal)
        
        intChargesUsedCount = UBound(intChargesUsed()) + 1
        If intChargesUsed(0) = 0 Then
            strTitle = strTitle & "; Charges used = All"
        Else
            If intChargesUsedCount = 1 Then
                strTitle = strTitle & "; Charge used = " & Trim(intChargesUsed(0))
            Else
                strTitle = strTitle & "; Charges used = "
            
                For intChargeIndex = 0 To intChargesUsedCount - 1
                    strTitle = strTitle & intChargesUsed(intChargeIndex)
                    If intChargeIndex < intChargesUsedCount - 1 Then
                        strTitle = strTitle & ", "
                    End If
                Next intChargeIndex
            End If
            
            If glbPreferencesExpanded.PairBrowserPlottingOptions.PlotAllChargeStates Then
                strTitle = strTitle & " (All)"
            End If
        End If
        .SetLabelGraphTitle strTitle
        
        ' Plot formatting
        .SetChartType oc2dTypePlot, 1
        .SetCurrentGroup 1
        .SetCurrentSeries 1
        
        ' Copying to local variable to make code cleaner
        udtGraphOptions = glbPreferencesExpanded.PairBrowserPlottingOptions.Graph2DOptions
        lngLightColor = udtGraphOptions.PointAndLineColor
        lngHeavyColor = glbPreferencesExpanded.PairBrowserPlottingOptions.PointAndLineColorHeavy
        
        .SetCurrentSeries 1
        If udtGraphOptions.ShowPointSymbols Then
            .SetStyleDataSymbol lngLightColor, val(udtGraphOptions.PointShape), udtGraphOptions.PointSizePixels
        Else
            .SetStyleDataSymbol lngLightColor, OlectraChart2D.ShapeConstants.oc2dShapeNone, 5
        End If

        If udtGraphOptions.DrawLinesBetweenPoints Then
            .SetStyleDataLine lngLightColor, oc2dLineSolid, udtGraphOptions.LineWidthPixels
        Else
            .SetStyleDataLine lngLightColor, oc2dLineNone, 1
        End If

        .SetStyleDataFill lngLightColor, oc2dFillSolid
        
        .SetCurrentSeries 2
        If udtGraphOptions.ShowPointSymbols Then
            .SetStyleDataSymbol lngHeavyColor, val(glbPreferencesExpanded.PairBrowserPlottingOptions.PointShapeHeavy), udtGraphOptions.PointSizePixels
        Else
            .SetStyleDataSymbol lngHeavyColor, OlectraChart2D.ShapeConstants.oc2dShapeNone, 5
        End If

        If udtGraphOptions.DrawLinesBetweenPoints Then
            .SetStyleDataLine lngHeavyColor, oc2dLineSolid, udtGraphOptions.LineWidthPixels
        Else
            .SetStyleDataLine lngHeavyColor, oc2dLineNone, 1
        End If

        .SetStyleDataFill lngHeavyColor, oc2dFillSolid
        
        .SetXAxisAnnotationMethod oc2dAnnotateValues
        .SetXAxisAnnotationPlacement oc2dAnnotateAuto
        
        .SetYAxisAnnotationMethod oc2dAnnotateValues
        .SetYAxisAnnotationPlacement oc2dAnnotateAuto

        .SetXAxisLabelFont 10
        .SetYAxisLabelFont 10
        
        .SetXAxisLabelFormatNumber 0
        .SetYAxisLabelFormatScientific 1, False

        If udtGraphOptions.ShowGridLines Then
            .SetYAxisGridlines oc2dLineDotted
        Else
            .SetYAxisGridlines oc2dLineNone
        End If
        
        ' Populate the data
        .SetCurrentSeries 1
        .SetSeriesDataPointCount 1, mDataPointCountLight
        .SetDataX 1, mCurrentXDataLight()
        .SetDataY 1, mCurrentYDataLight()

        .SetCurrentSeries 2
        .SetSeriesDataPointCount 2, mDataPointCountHeavy
        .SetDataX 2, mCurrentXDataHeavy()
        .SetDataY 2, mCurrentYDataHeavy()

        ' Set the Tick Spacing the default
        .SetXAxisTickSpacing 1, True

        .EnableDisableDelayUpdating False
    
    End With
    
    Exit Sub

UpdatePlotForPairErrorHandler:
    Debug.Assert False
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error updating plot: " & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    LogErrors Err.Number, "frmPairBrowser->UpdatePlotForPair", Err.Description, CallerIDLoaded
        
End Sub

Private Sub UpdateStatus(ByVal strMessage As String, Optional blnAppendOnOffText As Boolean = False, Optional blnOptionOn As Boolean = False)
    
    If blnAppendOnOffText Then
        If blnOptionOn Then
            strMessage = strMessage & "On"
        Else
            strMessage = strMessage & "Off"
        End If
    End If
    
    ctlStatus.AddMessageText strMessage
End Sub

' Unused Function (February 2005)
''Private Sub PopulateUMCPlotData(intChargesUsed() As Integer, ByVal lngUMCIndex As Long, ByRef lngDataPointCount As Long, ByRef dblScanData() As Double, ByRef dblAbuData() As Double, ByVal blnUseMaxValueEachScan As Boolean)
''    ' Note that the arrays are 1-based, for compatibility reasons with the plot control
''
''    Dim lngScanMin As Long
''    Dim lngScanIndex As Long
''    Dim lngScanCountNew As Long
''
''    Dim lngGapSize As Long
''    Dim lngMaxGapSize As Long
''    Dim lngScanIndexCompare As Long
''
''    Dim blnCopyDataPoint As Boolean
''    Dim blnCopyingGapData As Boolean
''
''    Dim intChargeIndex As Integer
''
''    With GelUMC(CallerIDLoaded).UMCs(lngUMCIndex)
''        lngScanMin = .MinScan
''        lngDataPointCount = .MaxScan - lngScanMin + 1
''    End With
''
''    ReDim dblScanData(1 To lngDataPointCount)
''    ReDim dblAbuData(1 To lngDataPointCount)
''
''    For lngScanIndex = 1 To lngDataPointCount
''        dblScanData(lngScanIndex) = lngScanIndex + lngScanMin - 1
''    Next lngScanIndex
''
''
''    If glbPreferencesExpanded.PairBrowserPlottingOptions.PlotAllChargeStates Or intChargesUsed(0) = 0 Then
''        ' Sum all charge states
''        PopulateUMCAbuDataWork dblAbuData(), CallerIDLoaded, lngUMCIndex, 0, lngScanMin, False
''    Else
''        ' Only use the charge states listed in intChargesUsed
''        For intChargeIndex = 0 To UBound(intChargesUsed())
''            PopulateUMCAbuDataWork dblAbuData(), CallerIDLoaded, lngUMCIndex, intChargesUsed(intChargeIndex), lngScanMin, blnUseMaxValueEachScan
''        Next intChargeIndex
''    End If
''
''    ' Remove the points with an abundance of 0, provided the gap size is less than
''    lngMaxGapSize = GelUMC(CallerIDLoaded).def.InterpolateMaxGapSize
''
''    lngScanCountNew = 0
''    blnCopyingGapData = False
''    For lngScanIndex = 1 To lngDataPointCount
''        If dblAbuData(lngScanIndex) = 0 Then
''            If blnCopyingGapData Then
''                blnCopyDataPoint = True
''            Else
''                If lngScanIndex = 1 Or lngScanIndex = lngDataPointCount Then
''                    blnCopyDataPoint = True
''                Else
''                    lngGapSize = lngDataPointCount - lngScanIndex + 1
''                    blnCopyDataPoint = False
''                    For lngScanIndexCompare = lngScanIndex + 1 To lngDataPointCount
''                        If dblAbuData(lngScanIndexCompare) > 0 Then
''                            ' Find the gap distance to the next scan with data
''                            lngGapSize = dblScanData(lngScanIndexCompare) - dblScanData(lngScanIndex)
''                            Exit For
''                        End If
''                    Next lngScanIndexCompare
''
''                    ' This should possibly be: If lngGapSize > lngMaxGapSize Then
''                    If lngGapSize > lngMaxGapSize Then
''                        blnCopyingGapData = True
''                        blnCopyDataPoint = True
''                    End If
''                End If
''            End If
''        Else
''            blnCopyDataPoint = True
''            blnCopyingGapData = False
''        End If
''
''        If blnCopyDataPoint Then
''            lngScanCountNew = lngScanCountNew + 1
''            dblScanData(lngScanCountNew) = dblScanData(lngScanIndex)
''            dblAbuData(lngScanCountNew) = dblAbuData(lngScanIndex)
''        End If
''    Next lngScanIndex
''
''    If lngScanCountNew <= 0 Then
''        ReDim dblScanData(1 To 1)
''        ReDim dblAbuData(1 To 1)
''        lngScanCountNew = 0
''    Else
''        ' Make sure there is a zero at the beginning of the array
''        If dblAbuData(1) <> 0 Then
''            lngScanCountNew = lngScanCountNew + 1
''            If lngScanCountNew > lngDataPointCount Then
''                ReDim Preserve dblScanData(1 To lngScanCountNew)
''                ReDim Preserve dblAbuData(1 To lngScanCountNew)
''            End If
''
''            For lngScanIndex = lngScanCountNew To 2 Step -1
''                dblAbuData(lngScanIndex) = dblAbuData(lngScanIndex - 1)
''                dblScanData(lngScanIndex) = dblScanData(lngScanIndex - 1)
''            Next lngScanIndex
''
''            dblAbuData(1) = 0
''            dblScanData(1) = dblScanData(2) - 1
''        End If
''
''        ' Make sure there is a zero at the end of the array
''        If dblAbuData(lngScanCountNew) <> 0 Then
''            lngScanCountNew = lngScanCountNew + 1
''            If lngScanCountNew > lngDataPointCount Then
''                ReDim Preserve dblScanData(1 To lngScanCountNew)
''                ReDim Preserve dblAbuData(1 To lngScanCountNew)
''            End If
''
''            dblAbuData(lngScanCountNew) = 0
''            dblScanData(lngScanCountNew) = dblScanData(lngScanCountNew - 1) + 1
''        End If
''
''        If lngScanCountNew < lngDataPointCount Then
''            ReDim Preserve dblScanData(1 To lngScanCountNew)
''            ReDim Preserve dblAbuData(1 To lngScanCountNew)
''        End If
''        lngDataPointCount = lngScanCountNew
''    End If
''
''End Sub
''
' Unused Function (February 2005)
''Private Sub PopulateUMCAbuDataWork(ByRef dblAbundance() As Double, ByVal lngGelIndex As Long, ByVal UMCIndex As Long, ByVal intTargetCharge As Integer, ByVal lngScanNumberStart As Long, blnUseMaxValueEachScan As Boolean)
''    ' Note: The algorithms in this function are the same as those in
''    '       Pairs.bas->CalcDltLblPairsScanByScanPopulate
''    '
''    ' Note that the dblAbundance() array is 1-based, for compatibility reasons with the plot control
''
''    Dim lngMemberIndex As Long
''    Dim lngScan As Long, lngScanIndex As Long
''    Dim intCharge As Integer
''    Dim dblAbu As Double
''
''    With GelUMC(lngGelIndex).UMCs(UMCIndex)
''        For lngMemberIndex = 0 To .ClassCount - 1
''            Select Case .ClassMType(lngMemberIndex)
''            Case gldtCS
''                 lngScan = GelData(lngGelIndex).CSData(.ClassMInd(lngMemberIndex)).ScanNumber
''                 intCharge = GelData(lngGelIndex).CSData(.ClassMInd(lngMemberIndex)).Charge
''                 dblAbu = GelData(lngGelIndex).CSData(.ClassMInd(lngMemberIndex)).Abundance
''            Case gldtIS
''                 lngScan = GelData(lngGelIndex).IsoData(.ClassMInd(lngMemberIndex)).ScanNumber
''                 intCharge = GelData(lngGelIndex).IsoData(.ClassMInd(lngMemberIndex)).Charge
''                 dblAbu = GelData(lngGelIndex).IsoData(.ClassMInd(lngMemberIndex)).Abundance
''            End Select
''
''            If intTargetCharge <= 0 Or intCharge = intTargetCharge Then
''                ' Note: Must add 1 due to 1-based array
''                lngScanIndex = lngScan - lngScanNumberStart + 1
''                If lngScanIndex < 1 Then
''                    ' This shouldn't happen
''                    Debug.Assert False
''                Else
''                    If blnUseMaxValueEachScan Then
''                        If dblAbu > dblAbundance(lngScanIndex) Then
''                            dblAbundance(lngScanIndex) = dblAbu
''                        End If
''                    Else
''                        dblAbundance(lngScanIndex) = dblAbundance(lngScanIndex) + dblAbu
''                    End If
''                End If
''            End If
''        Next lngMemberIndex
''    End With
''
''End Sub

Private Sub UndoPairDeletion()
    Dim lngIndex As Long
    Dim lngDereferencedIndex As Long
    Dim blnMatchFound As Boolean
    
    If mDeletedPairsStackCount > 0 Then
        
        lngDereferencedIndex = mDeletedPairsStack(mDeletedPairsStackCount - 1)
        
        ' Find lngDereferencedIndex in mPairInfoSortedPointerArray()
        For lngIndex = 0 To mPairsCount - 1
            If mPairInfoSortedPointerArray(lngIndex) = lngDereferencedIndex Then
                lstPairs.Selected(lngIndex) = True
            Else
                lstPairs.Selected(lngIndex) = False
            End If
        Next lngIndex
        
        If blnMatchFound Then
            UnDeleteSelectedPair
        Else
            mPairValid(mDeletedPairsStack(mDeletedPairsStackCount - 1)) = True
            mDeletedPairsStackCount = mDeletedPairsStackCount - 1
            
            UpdateListboxCaptionsSelected
            
        End If
    End If
    
End Sub

Private Sub cboMassRangeUnits_Click()
    If mFormInitialized Then glbPreferencesExpanded.PairBrowserPlottingOptions.MassRangeUnits = cboMassRangeUnits.ListIndex
End Sub

Private Sub cboPairSortOrder_Click()
    If mFormInitialized Then
        glbPreferencesExpanded.PairBrowserPlottingOptions.SortOrder = cboPairSortOrder.ListIndex
        SortAndDisplayPairs
    End If
End Sub

Private Sub cboPointShapeHeavy_Click()
    If mFormInitialized Then
        glbPreferencesExpanded.PairBrowserPlottingOptions.PointShapeHeavy = cboPointShapeHeavy.ListIndex + 1
        DisplayPairInfoSelectedItem True
    End If
End Sub

Private Sub cboPointShapeLight_Click()
    If mFormInitialized Then
        glbPreferencesExpanded.PairBrowserPlottingOptions.Graph2DOptions.PointShape = cboPointShapeLight.ListIndex + 1
        DisplayPairInfoSelectedItem True
    End If
End Sub

Private Sub cboScanRangeUnits_Click()
    If mFormInitialized Then glbPreferencesExpanded.PairBrowserPlottingOptions.ScanRangeUnits = cboScanRangeUnits.ListIndex
End Sub

Private Sub chkDrawLinesBetweenPoints_Click()
    If mFormInitialized Then
        glbPreferencesExpanded.PairBrowserPlottingOptions.Graph2DOptions.DrawLinesBetweenPoints = cChkBox(chkDrawLinesBetweenPoints)
        DisplayPairInfoSelectedItem True
    End If
End Sub

Private Sub chkFilterPairsOnMTHits_Click()
    FilterPairsOnMTHits cChkBox(chkFilterPairsOnMTHits)
End Sub

Private Sub chkShowGridlines_Click()
    If mFormInitialized Then
        glbPreferencesExpanded.PairBrowserPlottingOptions.Graph2DOptions.ShowGridLines = cChkBox(chkShowGridlines)
        DisplayPairInfoSelectedItem True
    End If
End Sub

Private Sub chkShowPointSymbols_Click()
    If mFormInitialized Then
        glbPreferencesExpanded.PairBrowserPlottingOptions.Graph2DOptions.ShowPointSymbols = cChkBox(chkShowPointSymbols)
        DisplayPairInfoSelectedItem True
    End If
End Sub

Private Sub chkSortDescending_Click()
    If mFormInitialized Then
        glbPreferencesExpanded.PairBrowserPlottingOptions.SortDescending = cChkBox(chkSortDescending)
        SortAndDisplayPairs
    End If
End Sub

Private Sub Form_Activate()
    If Not mFormInitialized Then InitializeForm
End Sub

Private Sub Form_Load()
    
    SizeAndCenterWindow Me, cWindowBottomRight, 7300, 10000, True
    
    Me.ScaleMode = vbTwips
    mFormInitialized = False
    
    InitializePlot
    
    PositionControls
    
    ShowHideOptions True
    
    PopulateComboBoxes
    
    InitializeForm

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        If Not QueryUnloadForm() Then
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_Resize()
    PositionControls
End Sub

Private Sub lblPointColorSelectionHeavy_Click()
    SelectCustomColor lblPointColorSelectionHeavy
    glbPreferencesExpanded.PairBrowserPlottingOptions.PointAndLineColorHeavy = lblPointColorSelectionHeavy.BackColor
    DisplayPairInfoSelectedItem True
End Sub

Private Sub lblPointColorSelectionLight_Click()
    SelectCustomColor lblPointColorSelectionLight
    glbPreferencesExpanded.PairBrowserPlottingOptions.Graph2DOptions.PointAndLineColor = lblPointColorSelectionLight.BackColor
    DisplayPairInfoSelectedItem True
End Sub

Private Sub lstPairs_Click()
    DisplayPairInfoSelectedItem False
End Sub

Private Sub lstPairs_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Shift = 0 Then
        If KeyCode = 45 Then
            ' Insert key pressed
            UnDeleteSelectedPair
        ElseIf KeyCode = 46 Then
            ' Delete key pressed
            DeleteSelectedPair
        End If
    End If
End Sub

Private Sub mnuEdit_Click()
    If lstPairs.ListCount > 0 Then
        mnuEditDeleteSelectedPair.Enabled = True
        mnuEditUndeletedSelectedPair.Enabled = True
    Else
        mnuEditDeleteSelectedPair.Enabled = False
        mnuEditUndeletedSelectedPair.Enabled = False
    End If
    
    If mDeletedPairsStackCount > 0 Then
        mnuEditUndoDelete.Enabled = True
    Else
        mnuEditUndoDelete.Enabled = False
    End If
    
End Sub

Private Sub mnuCopyChart_Click(Index As Integer)
    Select Case Index
    Case ccmWMF
        ctlPairsPlot.CopyToClipboard oc2dFormatMetafile
    Case ccmEMF
        ctlPairsPlot.CopyToClipboard oc2dFormatEnhMetafile
    Case Else
        ' Includes ccmBMP
        ctlPairsPlot.CopyToClipboard oc2dFormatBitmap
    End Select
End Sub

Private Sub mnuEditCopyData_Click()
    ExportPlotDataToClipboardOrFile False, "", True
End Sub

Private Sub mnuEditCopyDataLineUpByScan_Click()
    ExportPlotDataToClipboardOrFile True, "", True
End Sub

Private Sub mnuEditDeleteSelectedPair_Click()
    DeleteSelectedPair
End Sub

Private Sub mnuEditUndeletedSelectedPair_Click()
    UnDeleteSelectedPair
End Sub

Private Sub mnuEditUndoDelete_Click()
    UndoPairDeletion
End Sub

Private Sub mnuExit_Click()
    If QueryUnloadForm() Then
        Unload Me
    End If
End Sub

Private Sub mnuFileFindPairsDeltaUMC_Click()
    On Error Resume Next
    If IsLoaded("frmUMCDltPairs") Then
        MsgBox "The LC-MS Feature Delta Pairing Analysis window is already open.", vbInformation + vbOKOnly, "Cannot Open Window"
    Else
        frmUMCDltPairs.Tag = CallerIDLoaded
        frmUMCDltPairs.FormMode = pfmDelta
        frmUMCDltPairs.Show vbModal
        PopulateFormWithData True
    End If
    
End Sub

Private Sub mnuFileSaveChanges_Click()
    If DeleteMarkedPairs(True) Then
        PopulateFormWithData True
    End If
End Sub

Private Sub mnuOptionsAutoZoom2D_Click()
    With glbPreferencesExpanded.PairBrowserPlottingOptions
        .AutoZoom2DPlot = Not .AutoZoom2DPlot
        mnuOptionsAutoZoom2D.Checked = .AutoZoom2DPlot
    
        UpdateStatus "Auto zoom 2D plot now ", True, .AutoZoom2DPlot
        
        ToggleWindowStayOnTop .AutoZoom2DPlot
    End With
End Sub

Private Sub mnuOptionsAutoZoomFixedDimensions_Click()
    SetFixedDimensionsForAutoZoom Not glbPreferencesExpanded.PairBrowserPlottingOptions.FixedDimensionsForAutoZoom
    With glbPreferencesExpanded.PairBrowserPlottingOptions
        UpdateStatus "Absolute auto-zoom dimensions now ", True, .FixedDimensionsForAutoZoom
    End With
End Sub

Private Sub mnuOptionsHighlightPairMembers_Click()
    With glbPreferencesExpanded.PairBrowserPlottingOptions
        .HighlightMembers = Not .HighlightMembers
        mnuOptionsHighlightPairMembers.Checked = .HighlightMembers
    
        UpdateStatus "Highlight pair members now ", True, .HighlightMembers
    End With
End Sub

Private Sub mnuOptionsKeepWindowOnTop_Click()
    ToggleWindowStayOnTop Not mWindowStayOnTopEnabled
End Sub

Private Sub mnuOptionsPlotAllChargeStates_Click()
    With glbPreferencesExpanded.PairBrowserPlottingOptions
        .PlotAllChargeStates = Not .PlotAllChargeStates
        mnuOptionsPlotAllChargeStates.Checked = .PlotAllChargeStates
    
        UpdateStatus "Plot all charge states now ", True, .PlotAllChargeStates
    End With
    DisplayPairInfoSelectedItem True
End Sub

Private Sub mnuOptionsPlotOptions_Click()
    ShowHideOptions False
End Sub

Private Sub txtGraphLineWidth_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtGraphLineWidth, KeyAscii, True, False
End Sub

Private Sub txtGraphLineWidth_LostFocus()
    ValidateTextboxValueLng txtGraphLineWidth, 1, 20, 3
    glbPreferencesExpanded.PairBrowserPlottingOptions.Graph2DOptions.LineWidthPixels = CLngSafe(txtGraphLineWidth)
    DisplayPairInfoSelectedItem True
End Sub

Private Sub txtGraphPointSize_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtGraphPointSize, KeyAscii, True, False
End Sub

Private Sub txtGraphPointSize_LostFocus()
    ValidateTextboxValueLng txtGraphPointSize, 1, 20, 2
    glbPreferencesExpanded.PairBrowserPlottingOptions.Graph2DOptions.PointSizePixels = CLngSafe(txtGraphPointSize)
    DisplayPairInfoSelectedItem True
End Sub

Private Sub txtMassRange_KeyPress(KeyAscii As Integer)
   TextBoxKeyPressHandler txtMassRange, KeyAscii, True, True, False
End Sub

Private Sub txtMassRange_LostFocus()
    If mFormInitialized Then
        If Not IsNumeric(txtMassRange) Then
            If glbPreferencesExpanded.PairBrowserPlottingOptions.MassRangeUnits = mruDa Then
                txtMassRange = "5"
            Else
                txtMassRange = "50"
            End If
        End If
        glbPreferencesExpanded.PairBrowserPlottingOptions.MassRangeZoom = val(txtMassRange)
    End If
End Sub

Private Sub txtScanRange_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtScanRange, KeyAscii, True, True, False
End Sub

Private Sub txtScanRange_LostFocus()
    If mFormInitialized Then
        If Not IsNumeric(txtScanRange) Then
            If glbPreferencesExpanded.PairBrowserPlottingOptions.ScanRangeUnits = sruNet Then
                txtScanRange = "0.1"
            Else
                txtScanRange = "50"
            End If
        End If
        glbPreferencesExpanded.PairBrowserPlottingOptions.ScanRangeZoom = val(txtScanRange)
    End If
End Sub
