VERSION 5.00
Begin VB.Form frmUMCBrowser 
   Caption         =   "LC-MS Feature Browser"
   ClientHeight    =   9090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   9090
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraPlotOptions 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Options"
      Height          =   1905
      Left            =   120
      TabIndex        =   14
      Top             =   3600
      Width           =   5175
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
      Begin VB.ComboBox cboPointShape 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1320
         Width           =   1695
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
         Caption         =   "Point Info"
         Height          =   255
         Left            =   2640
         TabIndex        =   22
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label lblPointColorSelection 
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
   Begin VB.Frame fraUMCInfo 
      Caption         =   "Info on Selected Feature"
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   6855
      Begin VB.Label lblUMCInfo2 
         Caption         =   "Info"
         Height          =   615
         Left            =   3600
         TabIndex        =   27
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblUMCInfo1 
         Caption         =   "Info"
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.ListBox lstUMCs 
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
         TabIndex        =   26
         Top             =   600
         Width           =   3015
      End
      Begin VB.CheckBox chkFilterUMCsOnMTHits 
         Caption         =   "Only show features with MT Tag hits"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   940
         Width           =   3015
      End
      Begin VB.ComboBox cboUMCSortOrder 
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
         Top             =   1320
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
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblMassRange 
         Caption         =   "Mass range"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1340
         Width           =   1095
      End
      Begin VB.Label lblScanRange 
         Caption         =   "Scan range"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1700
         Width           =   1095
      End
   End
   Begin VIPER.ctlSpectraPlotter ctlAbundancePlot 
      Height          =   5055
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   8916
   End
   Begin VB.Label lblUMCList 
      Caption         =   "Feature List (UMCs)"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileFindUMCsIonNetworks 
         Caption         =   "&Open Find UMC Ion Networks Window (better)"
      End
      Begin VB.Menu mnuFileFindUMCs2003 
         Caption         =   "Open Find UMC 2003 Window (faster, but less accurate)"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSaveChanges 
         Caption         =   "&Save Changes (Delete LC-MS Features)"
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
         Caption         =   "&Undo last LC-MS Feature deletion"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDeleteSelectedUMC 
         Caption         =   "&Delete Selected LC-MS Feature"
      End
      Begin VB.Menu mnuEditUndeletedSelectedUMC 
         Caption         =   "&Include Selected LC-MS Feature"
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
      Begin VB.Menu mnuOptionsHighlightUMCMembers 
         Caption         =   "&Highlight LC-MS Feature members"
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
Attribute VB_Name = "frmUMCBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DELETE_UMC_INDICATOR = " -- Delete"

Public CallerIDNew As Long
Private CallerIDLoaded As Long          ' 0 if no data is loaded

Private Enum eusUMCSortOrderConstants
    eusUMCIndex = 0
    eusMass = 1
    eusTime = 2
    eusAbundance = 3
    eusMemberCount = 4
    eusMOverZ = 5
    eusCharge = 6
End Enum

Private Enum ccmCopyChartMode
    ccmBMP = 0
    ccmWMF = 1
    ccmEMF = 2
End Enum

Private mUMCsCount As Long
Private mUMCValid() As Boolean                 ' 0-based array; Dereference into GelUMC() using mUMCInfoOriginalIndex()
Private mUMCInfoSortedPointerArray() As Long   ' 0-based array; pointer into mUMCValid
Private mUMCInfoOriginalIndex() As Long        ' 0-based array; Original index of UMC in GelUMC(); needed for option to filter to only include LC-MS Features with hits

Private mDeletedUMCsStackCount As Long
Private mDeletedUMCsStack() As Long            ' 0-based array; Indices of deleted LC-MS Features; pointers into mUMCValid(); used with the undo command

' The following are 1-based arrays, for compatibility with the plot control
Private mDataPointCount As Long
Private mCurrentXData() As Double          ' 1-based; Actually simply holding integer scan numbers, but must be type double to populate the chart
Private mCurrentYData() As Double          ' 1-based

Private mWindowStayOnTopEnabled As Boolean

Private mUpdatingControls As Boolean
Private mFormInitialized As Boolean

Public Sub AutoUpdatePlot(Optional blnForceUpdate As Boolean = False)
    PopulateFormWithData blnForceUpdate
End Sub

Private Sub AutoZoom2DPlot(lngUMCIndexOriginal As Long)
    ' Note: lngUMCIndexOriginal should be looked up from mUMCInfoOriginalIndex
    
On Error GoTo AutoZoom2DPlotErrorHandler

    If lngUMCIndexOriginal < 0 Or lngUMCIndexOriginal >= GelUMC(CallerIDLoaded).UMCCnt Then
        Exit Sub
    End If

    BrowseFeaturesZoomAndHighlight2DPlot glbPreferencesExpanded.UMCBrowserPlottingOptions, CallerIDLoaded, lngUMCIndexOriginal
    Exit Sub
    
AutoZoom2DPlotErrorHandler:
    Debug.Assert False
    Me.MousePointer = vbDefault
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error auto zooming: " & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    LogErrors Err.Number, "frmUMCBrowser->AutoZoom2DPlot", Err.Description, CallerIDLoaded

End Sub

Private Sub CustomizeMenus()
    mnuEditDeleteSelectedUMC.Caption = mnuEditDeleteSelectedUMC.Caption & vbTab & "Delete Key"
    mnuEditUndeletedSelectedUMC.Caption = mnuEditUndeletedSelectedUMC.Caption & vbTab & "Insert Key"
End Sub

Private Function DeleteMarkedUMCs(Optional blnInformIfNoneToDelete As Boolean = True) As Boolean
    ' Returns true if the LC-MS Features were deleted (or no deleted LC-MS Features exist)
    
    Dim eResponse As VbMsgBoxResult
    Dim lngUMCCountToDelete As Long
    Dim lngIndex As Long
    
    Dim lngNewUMCCount As Long
    Dim udtNewUMCList() As udtUMCType
    
    Dim blnSuccess As Boolean
    
    For lngIndex = 0 To mUMCsCount - 1
        If Not mUMCValid(lngIndex) Then
            lngUMCCountToDelete = lngUMCCountToDelete + 1
        End If
    Next lngIndex
    
    blnSuccess = True
    If lngUMCCountToDelete > 0 And CallerIDLoaded > 0 Then
        eResponse = MsgBox("You have marked " & Trim(mDeletedUMCsStackCount) & " LC-MS Features for deletion (" & Trim(mUMCsCount - lngUMCCountToDelete) & " will remain).  Choose Yes to proceed with deletion.", vbQuestion + vbYesNoCancel, "Delete LC-MS Features")
        
        If eResponse = vbYes Then
            ' Delete marked LC-MS Features
            ' Cannot copy in place since must use mUMCInfoOriginalIndex() pointer array
            With GelUMC(CallerIDLoaded)
            
                lngNewUMCCount = 0
                ReDim udtNewUMCList(.UMCCnt - 1)

                For lngIndex = 0 To mUMCsCount - 1
                    If mUMCValid(lngIndex) Then
                        udtNewUMCList(lngNewUMCCount) = .UMCs(mUMCInfoOriginalIndex(lngIndex))
                        lngNewUMCCount = lngNewUMCCount + 1
                    End If
                Next lngIndex
                
                .UMCCnt = lngNewUMCCount
                If .UMCCnt > 0 Then
                    ReDim Preserve udtNewUMCList(.UMCCnt - 1)
                Else
                    ReDim Preserve udtNewUMCList(0)
                End If
                .UMCs = udtNewUMCList
                
            End With
        Else
            blnSuccess = False
        End If
    Else
        If blnInformIfNoneToDelete And CallerIDLoaded > 0 Then
            MsgBox "No LC-MS Features have been marked for deletion.", vbInformation + vbOKOnly, "Nothing to do"
        End If
    End If
    
    DeleteMarkedUMCs = blnSuccess
End Function

Private Sub DeleteSelectedUMC()

    If lstUMCs.ListIndex < 0 Or CallerIDLoaded <= 0 Then
        ' Nothing selected
    Else
        Me.MousePointer = vbHourglass
    
        If BrowseFeaturesDeleteSelected(lstUMCs, mUMCInfoSortedPointerArray(), mUMCValid(), mDeletedUMCsStackCount, mDeletedUMCsStack()) Then
            UpdateListboxCaptionsSelected
        End If
    
        Me.MousePointer = vbDefault
    End If
    
End Sub

Private Sub DisplayUMCInfoSelectedItem(Optional blnSkipAutoZoom As Boolean = False)
    Dim strDescription As String
    Dim strDescriptionAddnl As String
    Dim lngUMCIndexDereferenced As Long
    Dim lngUMCIndexOriginal As Long
    
    If mUpdatingControls Then Exit Sub
    
    If lstUMCs.ListIndex < 0 Or CallerIDLoaded <= 0 Then
        strDescription = "LC-MS Feature not selected"
        strDescriptionAddnl = ""
    Else
        lngUMCIndexDereferenced = mUMCInfoSortedPointerArray(lstUMCs.ListIndex)
        lngUMCIndexOriginal = mUMCInfoOriginalIndex(lngUMCIndexDereferenced)
        strDescription = GenerateUMCDescription(lngUMCIndexDereferenced, True, strDescriptionAddnl)
        
        UpdatePlotForUMC lngUMCIndexOriginal
        
        If (glbPreferencesExpanded.UMCBrowserPlottingOptions.AutoZoom2DPlot Or glbPreferencesExpanded.UMCBrowserPlottingOptions.HighlightMembers) And Not blnSkipAutoZoom Then
            AutoZoom2DPlot lngUMCIndexOriginal
        End If
    End If

    lblUMCInfo1 = strDescription
    lblUMCInfo2 = strDescriptionAddnl

End Sub

Private Sub DisplayUMCsPopulateListbox(blnResortedData As Boolean)

    Dim lngIndex As Long
    Dim lngIndexSaved As Long
    Dim lngTopIndexSaved As Long
    
    Dim intCompareLen As Integer
    
    Dim strCaptionSaved As String
    Dim strCaption As String
    
    lngIndexSaved = lstUMCs.ListIndex
    lngTopIndexSaved = lstUMCs.TopIndex
    
    If blnResortedData And lngIndexSaved >= 0 Then
        strCaptionSaved = lstUMCs.List(lngIndexSaved)
        If Right(strCaptionSaved, Len(DELETE_UMC_INDICATOR)) = DELETE_UMC_INDICATOR Then
            strCaptionSaved = Left(strCaptionSaved, Len(strCaptionSaved) - Len(DELETE_UMC_INDICATOR))
        End If
        intCompareLen = Len(strCaptionSaved)
    End If
    
    lstUMCs.Clear
    
    If blnResortedData Then
        For lngIndex = 0 To mUMCsCount - 1
            strCaption = GenerateUMCDescription(mUMCInfoSortedPointerArray(lngIndex), False)
            lstUMCs.AddItem strCaption
            
            If intCompareLen > 0 Then
                If Left(strCaption, intCompareLen) = strCaptionSaved Then
                    lngIndexSaved = lngIndex
                End If
            End If
        Next lngIndex
    Else
        For lngIndex = 0 To mUMCsCount - 1
            lstUMCs.AddItem mUMCInfoOriginalIndex(mUMCInfoSortedPointerArray(lngIndex), False)
        Next lngIndex
    End If
    
    If Not blnResortedData Then
        If lngIndexSaved >= lngTopIndexSaved Then
            lstUMCs.TopIndex = lngTopIndexSaved
        End If
    End If
    
    If lngIndexSaved < 0 Then
        If lstUMCs.ListCount > 0 Then lstUMCs.ListIndex = 0
    ElseIf lngIndexSaved < lstUMCs.ListCount Then
        lstUMCs.ListIndex = lngIndexSaved
    Else
        lstUMCs.ListIndex = lstUMCs.ListCount - 1
    End If
        
    If blnResortedData Then
        If lstUMCs.ListIndex > 0 Then
            lstUMCs.TopIndex = lstUMCs.ListIndex - 1
        End If
    End If

End Sub

Private Function ExportPlotDataToClipboardOrFile(Optional strFilePath As String = "", Optional blnShowMessages As Boolean = True) As Long
    ' Returns 0 if success, the error code if an error

    Dim strData() As String                     ' 0-based array
    Dim strTextToCopy As String
    
    Dim lngIndex As Long
    
    Dim OutFileNum As Integer
    
    If mDataPointCount = 0 Then
        If blnShowMessages And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "No data found to copy", vbInformation + vbOKOnly, "No data"
        End If
        ExportPlotDataToClipboardOrFile = -1
        Exit Function
    End If
    
On Error GoTo ExportPlotDataToClipboardOrFileErrorHandler

    Me.MousePointer = vbHourglass

    ' Header row is strData(0), data starts at strData(1)
    ReDim strData(0 To mDataPointCount)
    
    ' Fill strData()
    ' Define the header row
    strData(0) = "Scan Number" & vbTab & "Abu"
    
    For lngIndex = 1 To mDataPointCount
        strData(lngIndex) = Round(mCurrentXData(lngIndex), 0) & vbTab & mCurrentYData(lngIndex)
    Next lngIndex
    
    If Len(strFilePath) > 0 Then
        OutFileNum = FreeFile()
        Open strFilePath For Output As #OutFileNum
        
        For lngIndex = 0 To mDataPointCount
            Print #OutFileNum, strData(lngIndex)
        Next lngIndex
        
        Close #OutFileNum
    Else
        strTextToCopy = FlattenStringArray(strData(), mDataPointCount + 1, vbCrLf, False)
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

Private Sub FilterUMCsOnMTHits(blnEnableFilter As Boolean)

    Static blnUpdating As Boolean
    
    If blnUpdating Then Exit Sub
    blnUpdating = True
    
    If Not DeleteMarkedUMCs(False) Then
        ' Unsaved changes exist, and the user cancelled saving changes
        ' Do not enable the filter
        SetCheckBox chkFilterUMCsOnMTHits, Not blnEnableFilter
        blnUpdating = False
        Exit Sub
    End If
    
    SetCheckBox chkFilterUMCsOnMTHits, blnEnableFilter
    PopulateFormWithData True
    
    blnUpdating = False
End Sub

Private Function GenerateUMCDescription(lngUMCIndexDereferenced As Long, blnVerbose As Boolean, Optional ByRef strDescriptionAddnl As String) As String
    ' Note: lngUMCIndexDereferenced should point into mUMCValid
    
    Dim strDescription As String
    Dim lngScanMin As Long, lngScanMax As Long
    Dim dblMassMin As Double, dblMassMax As Double
    Dim udtUMC As udtUMCType
    
    Dim lngUMCIndexOriginal
    Dim lngScanNumber As Long
    
    Dim strDBMatchList As String
    
    Dim lngIndex As Long
    Dim lngAMTIDCnt As Long
    Dim strAMTIDs() As String
    
On Error GoTo GenerateUMCDescriptionErrorHandler

    lngUMCIndexOriginal = mUMCInfoOriginalIndex(lngUMCIndexDereferenced)
    
    With GelUMC(CallerIDLoaded).UMCs(lngUMCIndexOriginal)
        strDBMatchList = FixNull(GelData(CallerIDLoaded).IsoData(.ClassRepInd).MTID)
        
        ' Extract just the MTID's from this list
        lngAMTIDCnt = GetAMTRefFromString1(strDBMatchList, strAMTIDs())
    End With
    
    With GelUMC(CallerIDLoaded)
    
        udtUMC = .UMCs(lngUMCIndexOriginal)
    
        With .UMCs(lngUMCIndexOriginal)
            If Not blnVerbose Then
                strDescription = "#" & Trim(lngUMCIndexOriginal) & ", "
                strDescription = strDescription & Round(.ClassMW, 1) & " Da, "
                
                Select Case .ClassRepType
                Case gldtCS
                    lngScanNumber = GelData(CallerIDLoaded).CSData(.ClassRepInd).ScanNumber
                Case gldtIS
                    lngScanNumber = GelData(CallerIDLoaded).IsoData(.ClassRepInd).ScanNumber
                End Select
                
                strDescription = strDescription & Round(ScanToGANET(CallerIDLoaded, lngScanNumber), 3) & " NET, "
    
                If Not mUMCValid(lngUMCIndexDereferenced) Then
                    strDescription = strDescription & DELETE_UMC_INDICATOR
                End If
                
                strDescriptionAddnl = ""
            Else
                strDescription = ""
                strDescription = strDescription & "Abundance " & DoubleToStringScientific(.ClassAbundance, 3) & vbCrLf
                
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
                BrowseFeaturesLookupScanAndMassLimits udtUMC, udtUMC, lngScanMin, lngScanMax, dblMassMin, dblMassMax
                
                strDescriptionAddnl = ""
                strDescriptionAddnl = strDescriptionAddnl & "Scan range " & lngScanMin & " to " & lngScanMax & vbCrLf
                strDescriptionAddnl = strDescriptionAddnl & "Mass range " & Round(dblMassMin, 4) & " to " & Round(dblMassMax, 4)
            End If
        End With
    End With
    
    GenerateUMCDescription = strDescription
    Exit Function

GenerateUMCDescriptionErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmUMCBrowser->GenerateUMCDescription", Err.Description, CallerIDLoaded
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
    
    ' Update the controls with the values in .UMCBrowserPlottingOptions
    With glbPreferencesExpanded.UMCBrowserPlottingOptions
        cboUMCSortOrder.ListIndex = .SortOrder
        SetCheckBox chkSortDescending, .SortDescending
        
        mnuOptionsAutoZoom2D.Checked = .AutoZoom2DPlot
        mnuOptionsHighlightUMCMembers.Checked = .HighlightMembers
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
            cboPointShape.ListIndex = .PointShape - 1
            lblPointColorSelection.BackColor = .PointAndLineColor
            
            txtGraphLineWidth = Trim(.LineWidthPixels)
        End With
        
        ToggleWindowStayOnTop .KeepWindowOnTop
    End With
    mUpdatingControls = False
    
    PopulateFormWithData True
    
    mFormInitialized = True
    Exit Sub

InitializeFormErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmUMCBrowser->InitializeForm", Err.Description, CallerIDLoaded
    Resume Next
    
End Sub

Private Sub InitializePlot()
    
    Dim dblBlankDataX(1 To 1) As Double
    Dim dblBlankDataY(1 To 1) As Double
    
    With ctlAbundancePlot
        .PopulateSymbolStyleComboBox cboPointShape
        
        .EnableDisableDelayUpdating True
        .SetCurrentGroup 2
        .SetSeriesCount 0

        .SetCurrentGroup 1
        .SetSeriesCount 1
        .SetCurrentSeries 1

        .SetSeriesDataPointCount 1, 1
        .SetDataX 1, dblBlankDataX()
        .SetDataY 1, dblBlankDataY()

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
    
    With cboUMCSortOrder
        .Clear
        .AddItem "Sort by LC-MS Feature Index", eusUMCSortOrderConstants.eusUMCIndex
        .AddItem "Sort by Mass", eusUMCSortOrderConstants.eusMass
        .AddItem "Sort by Time", eusUMCSortOrderConstants.eusTime
        .AddItem "Sort by Abundance", eusUMCSortOrderConstants.eusAbundance
        .AddItem "Sort by Member Count", eusUMCSortOrderConstants.eusMemberCount
        .AddItem "Sort by m/z", eusUMCSortOrderConstants.eusMOverZ
        .AddItem "Sort by Charge", eusUMCSortOrderConstants.eusCharge
        .ListIndex = eusAbundance
    End With
    mUpdatingControls = False
    
End Sub

Private Sub PopulateFormWithData(Optional blnForcePopulation As Boolean = False)
    
    Dim blnCallerIDValid As Boolean
    Dim lngUMCIndex As Long
    Dim blnAddUMC As Boolean
    
On Error GoTo PopulateControlsErrorHandler

    blnCallerIDValid = False
    If CallerIDNew <> CallerIDLoaded Or blnForcePopulation Then
        
        CallerIDLoaded = CallerIDNew
        
        Me.Caption = "LC-MS Feature Browser: " & GelBody(CallerIDLoaded).Caption
        
        If CallerIDLoaded > UBound(GelUMC) Then
            CallerIDLoaded = UBound(GelUMC)
        End If
        
        With GelUMC(CallerIDLoaded)
            blnCallerIDValid = True
            
            If .UMCCnt <= 0 Then
                ReDim mUMCValid(0)
                mUMCsCount = 0
            Else
                ReDim mUMCValid(.UMCCnt - 1)
                ReDim mUMCInfoOriginalIndex(.UMCCnt - 1)
            
                If cChkBox(chkFilterUMCsOnMTHits) Then
                    mUMCsCount = 0
                    For lngUMCIndex = 0 To .UMCCnt - 1
                        
                        If IsAMTReferencedByUMC(.UMCs(lngUMCIndex), CallerIDLoaded) Then
                            blnAddUMC = True
                        Else
                            blnAddUMC = False
                        End If
                            
                        If blnAddUMC Then
                            mUMCValid(mUMCsCount) = True
                            mUMCInfoOriginalIndex(mUMCsCount) = lngUMCIndex
                            mUMCsCount = mUMCsCount + 1
                        End If
                        
                    Next lngUMCIndex
                    
                Else
                    mUMCsCount = .UMCCnt
                    For lngUMCIndex = 0 To .UMCCnt - 1
                        mUMCValid(lngUMCIndex) = True
                        mUMCInfoOriginalIndex(lngUMCIndex) = lngUMCIndex
                    Next lngUMCIndex
                End If
            End If
            
            mDeletedUMCsStackCount = 0
            ReDim mDeletedUMCsStack(0)
        End With
        
        SortAndDisplayUMCs
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
    
    With lstUMCs
        lngDesiredValue = Me.ScaleHeight * 0.2
        If lngDesiredValue < 2100 Then lngDesiredValue = 2100
        .Height = lngDesiredValue
        
        lngDesiredValue = .Top + .Height + 60
        If lngDesiredValue < fraOptions.Top + fraOptions.Height Then
            lngDesiredValue = fraOptions.Top + fraOptions.Height
        End If
        fraUMCInfo.Top = lngDesiredValue
    End With
    
    With ctlAbundancePlot
        .Top = fraUMCInfo.Top + fraUMCInfo.Height + 60
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
    
    blnDeleted = DeleteMarkedUMCs(False)
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

Private Sub SortAndDisplayUMCs()
    Dim dblDataToSort() As Double
    Dim lngIndex As Long
    
    Dim objQSDouble As QSDouble
    Dim blnPerformSort As Boolean
    Dim blnSuccess As Boolean
    
On Error GoTo SortAndDisplayUMCsErrorHandler

    If mUMCsCount > 0 Then
        
        Me.MousePointer = vbHourglass
        DoEvents
        
        ReDim dblDataToSort(mUMCsCount - 1)
        ReDim mUMCInfoSortedPointerArray(mUMCsCount - 1)
        For lngIndex = 0 To mUMCsCount - 1
            mUMCInfoSortedPointerArray(lngIndex) = lngIndex
        Next lngIndex
        
        With GelUMC(CallerIDLoaded)
            blnPerformSort = True
            Select Case cboUMCSortOrder.ListIndex
            Case eusMass
                For lngIndex = 0 To mUMCsCount - 1
                    dblDataToSort(lngIndex) = .UMCs(mUMCInfoOriginalIndex(lngIndex)).ClassMW
                Next lngIndex
            Case eusTime
                For lngIndex = 0 To mUMCsCount - 1
                    With .UMCs(mUMCInfoOriginalIndex(lngIndex))
                        Select Case .ClassRepType
                        Case gldtCS
                            dblDataToSort(lngIndex) = GelData(CallerIDLoaded).CSData(.ClassRepInd).ScanNumber
                        Case gldtIS
                            dblDataToSort(lngIndex) = GelData(CallerIDLoaded).IsoData(.ClassRepInd).ScanNumber
                        End Select
                    End With
                Next lngIndex
            Case eusAbundance
                For lngIndex = 0 To mUMCsCount - 1
                    dblDataToSort(lngIndex) = .UMCs(mUMCInfoOriginalIndex(lngIndex)).ClassAbundance
                Next lngIndex
            Case eusMemberCount
                For lngIndex = 0 To mUMCsCount - 1
                    dblDataToSort(lngIndex) = .UMCs(mUMCInfoOriginalIndex(lngIndex)).ClassCount
                Next lngIndex
            Case eusMOverZ
                ' Use class rep m/z
                For lngIndex = 0 To mUMCsCount - 1
                    With .UMCs(mUMCInfoOriginalIndex(lngIndex))
                        Select Case .ClassRepType
                        Case gldtCS
                            dblDataToSort(lngIndex) = GelData(CallerIDLoaded).CSData(.ClassRepInd).AverageMW
                        Case gldtIS
                            dblDataToSort(lngIndex) = GelData(CallerIDLoaded).IsoData(.ClassRepInd).MZ
                        End Select
                    End With
                Next lngIndex
            Case eusCharge
                ' Use class rep charge, unless it's 0, then use charge of class rep
                For lngIndex = 0 To mUMCsCount - 1
                    With .UMCs(mUMCInfoOriginalIndex(lngIndex))
                        dblDataToSort(lngIndex) = .ChargeStateBasedStats(.ChargeStateStatsRepInd).Charge
                        If dblDataToSort(lngIndex) = 0 Then
                            Select Case .ClassRepType
                            Case gldtCS
                                dblDataToSort(lngIndex) = GelData(CallerIDLoaded).CSData(.ClassRepInd).ChargeCount
                            Case gldtIS
                                dblDataToSort(lngIndex) = GelData(CallerIDLoaded).IsoData(.ClassRepInd).Charge
                            End Select
                        End If
                        
                        On Error Resume Next
                        ' Append the abundance
                        dblDataToSort(lngIndex) = dblDataToSort(lngIndex) + (Log(.ClassAbundance) / Log(10)) / 100
                        On Error GoTo SortAndDisplayUMCsErrorHandler
                        
                    End With
                Next lngIndex
            Case Else
                ' Includes eusUMCIndex
                ' Nothing to sort
                blnPerformSort = False
            End Select
        End With
        
        If blnPerformSort Then
            Set objQSDouble = New QSDouble
            If glbPreferencesExpanded.UMCBrowserPlottingOptions.SortDescending Then
                blnSuccess = objQSDouble.QSDesc(dblDataToSort, mUMCInfoSortedPointerArray)
            Else
                blnSuccess = objQSDouble.QSAsc(dblDataToSort, mUMCInfoSortedPointerArray)
            End If
            
            If Not blnSuccess Then
                ' Error performing sort
                Debug.Assert False
                MsgBox "Error sorting LC-MS Features: " & Err.Description, vbExclamation + vbOKOnly, "Error"
                LogErrors Err.Number, "frmUMCBrowser->SortAndDisplayUMCs", Err.Description, CallerIDLoaded
            End If
        End If
        
        DisplayUMCsPopulateListbox True
    
    Else
        lstUMCs.Clear
    End If
   
    Me.MousePointer = vbDefault
    Exit Sub

SortAndDisplayUMCsErrorHandler:
    Me.MousePointer = vbDefault
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error sorting LC-MS Features and populating list: " & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    LogErrors Err.Number, "SortAndDisplayUMCs", Err.Description, CallerIDLoaded

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
    
    With glbPreferencesExpanded.UMCBrowserPlottingOptions
        .FixedDimensionsForAutoZoom = blnEnabled
        mnuOptionsAutoZoomFixedDimensions.Checked = blnEnabled
    End With
    
    If blnEnabled Then
        lblMassRange = "Mass range"
        lblScanRange = "Scan range"
        
        If mFormInitialized Then
            dblDeltaMW = 50
            If CallerIDLoaded > 0 Then
                On Error Resume Next
                With GelUMC(CallerIDLoaded)
                    Select Case .def.TolType
                    Case gltPPM
                        dblDeltaMW = .def.Tol
                    Case gltABS
                        dblDeltaMW = .def.Tol / glPPM / 1000
                    End Select
                    dblDeltaMW = dblDeltaMW * 4
                End With
            End If
            
            If dblDeltaMW < 10 Then dblDeltaMW = 10
            
            glbPreferencesExpanded.UMCBrowserPlottingOptions.MassRangeZoom = dblDeltaMW
            txtMassRange = dblDeltaMW
            cboMassRangeUnits.ListIndex = mruPpm
        End If
    Else
        lblMassRange = "Mass edge"
        lblScanRange = "Scan edge"
    
        If mFormInitialized Then
            dblDeltaMW = 20
            glbPreferencesExpanded.UMCBrowserPlottingOptions.MassRangeZoom = dblDeltaMW
            
            txtMassRange = Trim(dblDeltaMW)
            cboMassRangeUnits.ListIndex = mruPpm
        End If
    End If
    
End Sub

Private Sub ToggleWindowStayOnTop(blnEnableStayOnTop As Boolean)
    
    mnuOptionsKeepWindowOnTop.Checked = blnEnableStayOnTop
    glbPreferencesExpanded.UMCBrowserPlottingOptions.KeepWindowOnTop = blnEnableStayOnTop
    
    If mWindowStayOnTopEnabled = blnEnableStayOnTop Then Exit Sub
    
    Me.ScaleMode = vbTwips
    
    WindowStayOnTop Me.hwnd, blnEnableStayOnTop, Me.ScaleX(Me.Left, vbTwips, vbPixels), Me.ScaleY(Me.Top, vbTwips, vbPixels), Me.ScaleX(Me.width, vbTwips, vbPixels), Me.ScaleY(Me.Height, vbTwips, vbPixels)
    
    mWindowStayOnTopEnabled = blnEnableStayOnTop

End Sub

Private Sub UnDeleteSelectedUMC()
    Dim blnUpdateListBox As Boolean

    If lstUMCs.ListIndex < 0 Or CallerIDLoaded <= 0 Then
        ' Nothing selected
    Else
        Me.MousePointer = vbHourglass
        
        blnUpdateListBox = BrowseFeaturesUndeleteSelected(lstUMCs, mUMCInfoSortedPointerArray(), mUMCValid(), mDeletedUMCsStackCount, mDeletedUMCsStack())
        
        If blnUpdateListBox Then
            UpdateListboxCaptionsSelected
        End If
    
    End If
    
    Me.MousePointer = vbDefault

End Sub

Private Sub UpdateListboxCaptionsSelected()
    Dim lngIndex As Long
    
On Error GoTo UpdateListboxCaptionsSelectedErrorHandler

    ' Update the caption for each item selected in lstUMCs
    For lngIndex = 0 To lstUMCs.ListCount - 1
        If lstUMCs.Selected(lngIndex) Then
            lstUMCs.List(lngIndex) = GenerateUMCDescription(mUMCInfoSortedPointerArray(lngIndex), False)
        End If
    Next lngIndex

Exit Sub

UpdateListboxCaptionsSelectedErrorHandler:
    Debug.Assert False
    LogErrors Err.Number, "frmUMCBrowser->UpdateListboxCaptionsSelected", Err.Description, CallerIDLoaded
    DisplayUMCsPopulateListbox True
    
End Sub

Private Sub UpdatePlotForUMC(lngUMCIndexOriginal As Long)
    ' Note: lngUMCIndexOriginal should be looked up from mUMCInfoOriginalIndex
    
    Dim intChargesUsedCount As Integer
    Dim intChargesUsed() As Integer
    Dim intChargeIndex As Integer
    
    Dim strTitle As String
    Dim udtGraphOptions As udtGraph2DOptionsType
    
    Dim blnUseMaxValueEachScan As Boolean
    
On Error GoTo UpdatePlotForUMCErrorHandler

    ' This is False for LC-MS Features
    blnUseMaxValueEachScan = False
    
    ' Look up the charges used to compute this LC-MS Feature's abundance
    With GelUMC(CallerIDLoaded)
        With .UMCs(lngUMCIndexOriginal)
            If .ChargeStateCount > 0 Then
                If glbPreferencesExpanded.UMCBrowserPlottingOptions.PlotAllChargeStates Then
                    ' Copy from .ChargeStateBasedStats().Charge to intChargesUsed()
                    ReDim intChargesUsed(.ChargeStateCount - 1)
                    For intChargeIndex = 0 To .ChargeStateCount - 1
                        intChargesUsed(intChargeIndex) = .ChargeStateBasedStats(intChargeIndex).Charge
                    Next intChargeIndex
                Else
                    ReDim intChargesUsed(0)
                    intChargesUsed(0) = .ChargeStateBasedStats(.ChargeStateStatsRepInd).Charge
                End If
            Else
                ReDim intChargesUsed(0)
                intChargesUsed(0) = 0
            End If
        End With
    End With
    
    BrowseFeaturesPopulateUMCPlotData CallerIDLoaded, glbPreferencesExpanded.UMCBrowserPlottingOptions.PlotAllChargeStates, intChargesUsed(), lngUMCIndexOriginal, mDataPointCount, mCurrentXData(), mCurrentYData(), blnUseMaxValueEachScan
    
    With ctlAbundancePlot
        .EnableDisableDelayUpdating True
        
        strTitle = "LC-MS Feature #" & Trim(lngUMCIndexOriginal)
        
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
            
            If glbPreferencesExpanded.UMCBrowserPlottingOptions.PlotAllChargeStates Then
                strTitle = strTitle & " (All)"
            End If
        End If
        .SetLabelGraphTitle strTitle
        
        ' Plot formatting
        .SetChartType oc2dTypePlot, 1
        .SetCurrentGroup 1
        .SetCurrentSeries 1
        
        ' Copying to local variable to make code cleaner
        udtGraphOptions = glbPreferencesExpanded.UMCBrowserPlottingOptions.Graph2DOptions
        
        .SetCurrentSeries 1
        If udtGraphOptions.ShowPointSymbols Then
            .SetStyleDataSymbol udtGraphOptions.PointAndLineColor, val(udtGraphOptions.PointShape), udtGraphOptions.PointSizePixels
        Else
            .SetStyleDataSymbol udtGraphOptions.PointAndLineColor, OlectraChart2D.ShapeConstants.oc2dShapeNone, 5
        End If

        If udtGraphOptions.DrawLinesBetweenPoints Then
            .SetStyleDataLine udtGraphOptions.PointAndLineColor, oc2dLineSolid, udtGraphOptions.LineWidthPixels
        Else
            .SetStyleDataLine udtGraphOptions.PointAndLineColor, oc2dLineNone, 1
        End If

        .SetStyleDataFill udtGraphOptions.PointAndLineColor, oc2dFillSolid
        
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
        .SetSeriesDataPointCount 1, mDataPointCount
        .SetDataX 1, mCurrentXData()
        .SetDataY 1, mCurrentYData()

        ' Set the Tick Spacing the default
        .SetXAxisTickSpacing 1, True

        .EnableDisableDelayUpdating False
    
    End With
    
    Exit Sub

UpdatePlotForUMCErrorHandler:
    Debug.Assert False
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error updating plot: " & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    LogErrors Err.Number, "frmUMCBrowser->UpdatePlotForUMC", Err.Description, CallerIDLoaded
        
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

Private Sub UndoUMCDeletion()
    Dim lngIndex As Long
    Dim lngDereferencedIndex As Long
    Dim blnMatchFound As Boolean
    
    If mDeletedUMCsStackCount > 0 Then
        
        lngDereferencedIndex = mDeletedUMCsStack(mDeletedUMCsStackCount - 1)
        
        ' Find lngDereferencedIndex in mUMCInfoSortedPointerArray()
        For lngIndex = 0 To mUMCsCount - 1
            If mUMCInfoSortedPointerArray(lngIndex) = lngDereferencedIndex Then
                lstUMCs.Selected(lngIndex) = True
            Else
                lstUMCs.Selected(lngIndex) = False
            End If
        Next lngIndex
        
        If blnMatchFound Then
            UnDeleteSelectedUMC
        Else
            mUMCValid(mDeletedUMCsStack(mDeletedUMCsStackCount - 1)) = True
            mDeletedUMCsStackCount = mDeletedUMCsStackCount - 1
            
            UpdateListboxCaptionsSelected
            
        End If
    End If

End Sub

Private Sub cboMassRangeUnits_Click()
    If mFormInitialized Then glbPreferencesExpanded.UMCBrowserPlottingOptions.MassRangeUnits = cboMassRangeUnits.ListIndex
End Sub

Private Sub cboUMCSortOrder_Click()
    If mFormInitialized Then
        glbPreferencesExpanded.UMCBrowserPlottingOptions.SortOrder = cboUMCSortOrder.ListIndex
        SortAndDisplayUMCs
    End If
End Sub

Private Sub cboPointShape_Click()
    If mFormInitialized Then
        glbPreferencesExpanded.UMCBrowserPlottingOptions.Graph2DOptions.PointShape = cboPointShape.ListIndex + 1
        DisplayUMCInfoSelectedItem True
    End If
End Sub

Private Sub cboScanRangeUnits_Click()
    If mFormInitialized Then glbPreferencesExpanded.UMCBrowserPlottingOptions.ScanRangeUnits = cboScanRangeUnits.ListIndex
End Sub

Private Sub chkDrawLinesBetweenPoints_Click()
    If mFormInitialized Then
        glbPreferencesExpanded.UMCBrowserPlottingOptions.Graph2DOptions.DrawLinesBetweenPoints = cChkBox(chkDrawLinesBetweenPoints)
        DisplayUMCInfoSelectedItem True
    End If
End Sub

Private Sub chkFilterUMCsOnMTHits_Click()
    FilterUMCsOnMTHits cChkBox(chkFilterUMCsOnMTHits)
End Sub

Private Sub chkShowGridlines_Click()
    If mFormInitialized Then
        glbPreferencesExpanded.UMCBrowserPlottingOptions.Graph2DOptions.ShowGridLines = cChkBox(chkShowGridlines)
        DisplayUMCInfoSelectedItem True
    End If
End Sub

Private Sub chkShowPointSymbols_Click()
    If mFormInitialized Then
        glbPreferencesExpanded.UMCBrowserPlottingOptions.Graph2DOptions.ShowPointSymbols = cChkBox(chkShowPointSymbols)
        DisplayUMCInfoSelectedItem True
    End If
End Sub

Private Sub chkSortDescending_Click()
    If mFormInitialized Then
        glbPreferencesExpanded.UMCBrowserPlottingOptions.SortDescending = cChkBox(chkSortDescending)
        SortAndDisplayUMCs
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

Private Sub lblPointColorSelection_Click()
    SelectCustomColor lblPointColorSelection
    glbPreferencesExpanded.UMCBrowserPlottingOptions.Graph2DOptions.PointAndLineColor = lblPointColorSelection.BackColor
    DisplayUMCInfoSelectedItem True
End Sub

Private Sub lstUMCs_Click()
    DisplayUMCInfoSelectedItem False
End Sub

Private Sub lstUMCs_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Shift = 0 Then
        If KeyCode = 45 Then
            ' Insert key pressed
            UnDeleteSelectedUMC
        ElseIf KeyCode = 46 Then
            ' Delete key pressed
            DeleteSelectedUMC
        End If
    End If
End Sub

Private Sub mnuEdit_Click()
    If lstUMCs.ListCount > 0 Then
        mnuEditDeleteSelectedUMC.Enabled = True
        mnuEditUndeletedSelectedUMC.Enabled = True
    Else
        mnuEditDeleteSelectedUMC.Enabled = False
        mnuEditUndeletedSelectedUMC.Enabled = False
    End If
    
    If mDeletedUMCsStackCount > 0 Then
        mnuEditUndoDelete.Enabled = True
    Else
        mnuEditUndoDelete.Enabled = False
    End If
    
End Sub

Private Sub mnuCopyChart_Click(Index As Integer)
    Select Case Index
    Case ccmWMF
        ctlAbundancePlot.CopyToClipboard oc2dFormatMetafile
    Case ccmEMF
        ctlAbundancePlot.CopyToClipboard oc2dFormatEnhMetafile
    Case Else
        ' Includes ccmBMP
        ctlAbundancePlot.CopyToClipboard oc2dFormatBitmap
    End Select
End Sub

Private Sub mnuEditCopyData_Click()
    ExportPlotDataToClipboardOrFile "", True
End Sub

Private Sub mnuEditCopyDataLineUpByScan_Click()
    ExportPlotDataToClipboardOrFile "", True
End Sub

Private Sub mnuEditDeleteSelectedUMC_Click()
    DeleteSelectedUMC
End Sub

Private Sub mnuEditUndeletedSelectedUMC_Click()
    UnDeleteSelectedUMC
End Sub

Private Sub mnuEditUndoDelete_Click()
    UndoUMCDeletion
End Sub

Private Sub mnuExit_Click()
    If QueryUnloadForm() Then
        Unload Me
    End If
End Sub

Private Sub mnuFileFindUMCs2003_Click()
    On Error Resume Next
    If IsLoaded("frmUMCSimple") Then
        MsgBox "The LC-MS Feature (UMC) Definition window is already open.", vbInformation + vbOKOnly, "Cannot Open Window"
    Else
        frmUMCSimple.Tag = CallerIDLoaded
        frmUMCSimple.Show vbModal
        If GelUMCDraw(CallerIDLoaded).Visible Then
            With GelBody(CallerIDLoaded)
                .RequestRefreshPlot
                .csMyCooSys.CoordinateDraw
                .picGraph.Refresh
                .UpdateTICPlotAndFeatureBrowsersIfNeeded True
            End With
        End If
        PopulateFormWithData True
    End If
End Sub

Private Sub mnuFileFindUMCsIonNetworks_Click()
    On Error Resume Next
    If IsLoaded("frmUMCIonNet") Then
        MsgBox "The LC-MS Feature (UMC) Ion Networks window is already open.", vbInformation + vbOKOnly, "Cannot Open Window"
    Else
        frmUMCIonNet.Tag = CallerIDLoaded
        frmUMCIonNet.Show vbModal
        If GelUMCDraw(CallerIDLoaded).Visible Then
            With GelBody(CallerIDLoaded)
                .RequestRefreshPlot
                .csMyCooSys.CoordinateDraw
                .picGraph.Refresh
                .UpdateTICPlotAndFeatureBrowsersIfNeeded True
            End With
        End If
        PopulateFormWithData True
    End If
End Sub

Private Sub mnuFileSaveChanges_Click()
    If DeleteMarkedUMCs(True) Then
        PopulateFormWithData True
    End If
End Sub

Private Sub mnuOptionsAutoZoom2D_Click()
    With glbPreferencesExpanded.UMCBrowserPlottingOptions
        .AutoZoom2DPlot = Not .AutoZoom2DPlot
        mnuOptionsAutoZoom2D.Checked = .AutoZoom2DPlot
    
        UpdateStatus "Auto zoom 2D plot now ", True, .AutoZoom2DPlot
        
        ToggleWindowStayOnTop .AutoZoom2DPlot
    End With
End Sub

Private Sub mnuOptionsAutoZoomFixedDimensions_Click()
    SetFixedDimensionsForAutoZoom Not glbPreferencesExpanded.UMCBrowserPlottingOptions.FixedDimensionsForAutoZoom
    With glbPreferencesExpanded.UMCBrowserPlottingOptions
        UpdateStatus "Absolute auto-zoom dimensions now ", True, .FixedDimensionsForAutoZoom
    End With
End Sub

Private Sub mnuOptionsHighlightUMCMembers_Click()
    With glbPreferencesExpanded.UMCBrowserPlottingOptions
        .HighlightMembers = Not .HighlightMembers
        mnuOptionsHighlightUMCMembers.Checked = .HighlightMembers
    
        UpdateStatus "Highlight LC-MS Feature members now ", True, .HighlightMembers
    End With
End Sub

Private Sub mnuOptionsKeepWindowOnTop_Click()
    ToggleWindowStayOnTop Not mWindowStayOnTopEnabled
End Sub

Private Sub mnuOptionsPlotAllChargeStates_Click()
    With glbPreferencesExpanded.UMCBrowserPlottingOptions
        .PlotAllChargeStates = Not .PlotAllChargeStates
        mnuOptionsPlotAllChargeStates.Checked = .PlotAllChargeStates
    
        UpdateStatus "Plot all charge states now ", True, .PlotAllChargeStates
    End With
    DisplayUMCInfoSelectedItem True
End Sub

Private Sub mnuOptionsPlotOptions_Click()
    ShowHideOptions False
End Sub

Private Sub txtGraphLineWidth_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtGraphLineWidth, KeyAscii, True, False
End Sub

Private Sub txtGraphLineWidth_LostFocus()
    ValidateTextboxValueLng txtGraphLineWidth, 1, 20, 3
    glbPreferencesExpanded.UMCBrowserPlottingOptions.Graph2DOptions.LineWidthPixels = CLngSafe(txtGraphLineWidth)
    DisplayUMCInfoSelectedItem True
End Sub

Private Sub txtGraphPointSize_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtGraphPointSize, KeyAscii, True, False
End Sub

Private Sub txtGraphPointSize_LostFocus()
    ValidateTextboxValueLng txtGraphPointSize, 1, 20, 2
    glbPreferencesExpanded.UMCBrowserPlottingOptions.Graph2DOptions.PointSizePixels = CLngSafe(txtGraphPointSize)
    DisplayUMCInfoSelectedItem True
End Sub

Private Sub txtMassRange_KeyPress(KeyAscii As Integer)
   TextBoxKeyPressHandler txtMassRange, KeyAscii, True, True, False
End Sub

Private Sub txtMassRange_LostFocus()
    If mFormInitialized Then
        If Not IsNumeric(txtMassRange) Then
            If glbPreferencesExpanded.UMCBrowserPlottingOptions.MassRangeUnits = mruDa Then
                txtMassRange = "5"
            Else
                txtMassRange = "50"
            End If
        End If
        glbPreferencesExpanded.UMCBrowserPlottingOptions.MassRangeZoom = val(txtMassRange)
    End If
End Sub

Private Sub txtScanRange_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtScanRange, KeyAscii, True, True, False
End Sub

Private Sub txtScanRange_LostFocus()
    If mFormInitialized Then
        If Not IsNumeric(txtScanRange) Then
            If glbPreferencesExpanded.UMCBrowserPlottingOptions.ScanRangeUnits = sruNet Then
                txtScanRange = "0.1"
            Else
                txtScanRange = "50"
            End If
        End If
        glbPreferencesExpanded.UMCBrowserPlottingOptions.ScanRangeZoom = val(txtScanRange)
    End If
End Sub
