VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmORFViewer 
   Caption         =   "ORF Viewer"
   ClientHeight    =   6615
   ClientLeft      =   165
   ClientTop       =   840
   ClientWidth     =   14925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   14925
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraViewOptions 
      BackColor       =   &H00FFFFFF&
      Height          =   705
      Left            =   840
      TabIndex        =   10
      Top             =   5280
      Width           =   13215
      Begin VB.ComboBox cboDataDisplayMode 
         Height          =   315
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   120
         Width           =   2655
      End
      Begin VB.TextBox txtMassRange 
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Text            =   "25"
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox txtNETRange 
         Height          =   285
         Left            =   4320
         TabIndex        =   11
         Text            =   "0.2"
         Top             =   120
         Width           =   735
      End
      Begin MSComctlLib.ListView lvwColorKey 
         Height          =   495
         Left            =   8160
         TabIndex        =   18
         Top             =   120
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   873
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Status"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   420
         Width           =   7935
      End
      Begin VB.Label lblMassRange 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mass Range (ppm)"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   150
         Width           =   1455
      End
      Begin VB.Label lblNETRange 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NET Range (0 to 1)"
         Height          =   255
         Left            =   2760
         TabIndex        =   13
         Top             =   150
         Width           =   1455
      End
   End
   Begin VB.ComboBox cboItemCountSourceGel 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   9
      ToolTipText     =   $"frmORFViewer.frx":0000
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton cmdHorizontalDivider 
      Height          =   135
      Left            =   0
      MaskColor       =   &H8000000F&
      MousePointer    =   7  'Size N S
      TabIndex        =   8
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton cmdVerticalDivider 
      Height          =   5115
      Left            =   2160
      MaskColor       =   &H8000000F&
      MousePointer    =   9  'Size W E
      TabIndex        =   7
      Top             =   120
      Width           =   150
   End
   Begin MSComctlLib.ListView lvwORFs 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3836
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame fraORFPicsClippingRegion 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5055
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   7695
      Begin VB.Frame fraORFPicsContainer 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4335
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   7215
         Begin VIPER.LaSpots ctlORF 
            Height          =   1455
            Index           =   0
            Left            =   840
            TabIndex        =   17
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   2566
         End
      End
   End
   Begin VB.VScrollBar VScroll 
      Height          =   5055
      Left            =   10200
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton cmdRollUpExpand 
      Height          =   255
      Left            =   360
      Picture         =   "frmORFViewer.frx":008A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Expand the ORF List"
      Top             =   5280
      Width           =   240
   End
   Begin VB.CommandButton cmdRollUpShrink 
      Height          =   255
      Left            =   120
      Picture         =   "frmORFViewer.frx":03CC
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Shrink the ORF List"
      Top             =   5280
      Width           =   240
   End
   Begin MSComctlLib.ListView lvwMassTags 
      Height          =   2295
      Left            =   0
      TabIndex        =   6
      Top             =   2520
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   4048
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoadNewORFViewerWindow 
         Caption         =   "Show &New ORF Viewer Window"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefreshORFList 
         Caption         =   "&Refresh ORF List and Source Data"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuReloadORFsFromMTDB 
         Caption         =   "Reload ORF's from Mass Tag &DB"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy Selected Row(s)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnuInvertSelection 
         Caption         =   "&Invert Selection"
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFindText 
         Caption         =   "&Find Text in ORF List"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFindORFContaingMassTag 
         Caption         =   "Find ORF containing &Mass Tag"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuORFHistoryMovePrevious 
         Caption         =   "&Previous ORF in History"
      End
      Begin VB.Menu mnuORFHistoryMoveNext 
         Caption         =   "&Next ORF in History"
      End
      Begin VB.Menu mnuORFHistoryList 
         Caption         =   "ORF History"
         Begin VB.Menu mnuORFHistoryListItem 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnuViewSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHighlightMassTagsForSelectedORF 
         Caption         =   "&Highlight Mass Tags for Selected ORF in Parent"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuZoomParentGel 
         Caption         =   "&Zoom to region of selected mass tag in Parent"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuViewSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetsGelsAndOptions 
         Caption         =   "Set Included Analyses and Options"
         Shortcut        =   ^O
      End
   End
End
Attribute VB_Name = "frmORFViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Routines to load ORF information from a mass tag database and display
' zoomed in areas of a 2D plot (or overlay of plots), displaying only those regions
' where a peptide for a given ORF was seen or was expected to be seen
'
' Written by Matthew Monroe, PNNL
' Started December 31, 2002

Private Const MIN_MSEC_BETWEEN_UPDATES = 250
Private Const ORF_DIM_CHUNK = 100
Private Const ORF_HISTORY_COUNT_MAX = 30

Private Const RULE_MATCH_YES = "Yes"
Private Const RULE_MATCH_NO = "No "     ' I want a space here, but not in Rule_Match_Yes

' Set the following to 0 to hide the SortKey Columns
Private Const SORTKEY_COL_WIDTH = 0

Private Const LISTVIEW_COUNT = 3
Private Enum lviORFFormListViewIDConstants
    lviORFs = 0
    lviMassTags = 1
    lviColorKey = 2
End Enum

Private Const GRID_COL_COUNT_ORFS = 11
Private Enum lvoORFsColumnConstants
    lvoID = 0
    lvoReference
    lvoMassTagHitsViaIons
    lvoORFIntensityViaIons
    lvoMassTagHitsViaUMCs
    lvoORFIntensityViaUMCs
    lvoMassTags
    lvoTryptics
    lvoDescription
    lvoMass
    lvoSortkey
End Enum

Private Const GRID_COL_COUNT_MASS_TAGS = 12
Private Enum lvmMassTagsColumnConstants
    lvmID = 0
    lvmMTID
    lvmIonHits
    lvmIonHitSum
    lvmUMCHits
    lvmUMCHitSum
    lvmMass
    lvmNET
    lvmResidueCount
    lvmTrypticStatus
    lvmSequence
    lvmSortKey
End Enum

Private Const GRID_COL_COUNT_COLOR_KEY = 2
Private Enum lvkColorKeyColumnConstants
    lvkName = 0
    lvkUMCColor
End Enum

Private Type udtDataToPopulateType
    Count As Long
    Labels() As String              ' 0-based array, as are all of the other arrays in this UDT
    NET() As Double
    NETExtentNeg() As Double        ' NET Extent in the negative direction; e.g. if NET = 0.5 and NETExtentNeg = 0.1 and NETExtentPos = 0.2 then the NET range is 0.4 to 0.7
    NETExtentPos() As Double
    Mass() As Double
    MassExtentNeg() As Double       ' Mass Extent for AMT's in the negative direction; e.g. if AMT mass = 1000 Da and AMT error is 15 ppm = 0.015 Da, then MassExtentNeg is 0.0075 Da and MassExtentPos is 0.0075 Da, meaning AMT ranges from 999.985 to 1000.015
    MassExtentPos() As Double
    Intensity() As Double
End Type

Private Type udtMassTagIntensitySumType
    Count As Long
    Sum As Double
    MassTagRefID As Long
End Type

Private Type udtScanRangeType
    ScanNumberStart As Long
    ScanNumberEnd As Long
    NETStart As Double
    NETEnd As Double
    NETClassRepresentative As Double
End Type

' Note: Using 1-based arrays for .IonMass() and .IonPointer() to stay consistent with .IsoData() in GelData()
'       Using 0-based arrays for the .UMC arrays to stay consistent with .UMCs() in GelUMC
Private Type udtORFViewerDataIndicesType
    IonCount As Long
    IonMass() As Double             ' 1-based array.  Sorted, ascending
    IonMassPointer() As Long
    
    IonNET() As Double              ' 1-based array.  Not sorted.  However, filled after sorting IonMass().  So, the NET value for the ion in IonMass(0) can be found in IonNet(0)
    
    UMCCount As Long
    UMCMassMin() As Double          ' 0-based array.  Sorted, ascending
    UMCMassMinPointer() As Long
    
    UMCMassMax() As Double          ' 0-based array.  Sorted, ascending
    UMCMassMaxPointer() As Long
    
    UMCScanRange() As udtScanRangeType   ' 0-based array.  Not sorted
    
End Type

Private Type udtORFViewerDataIndicesListType
    GelCount As Integer
    Gels() As udtORFViewerDataIndicesType           ' 1-based array; .Gels(1) corresponds to data in GelData(1)
End Type

' The following index array can be used to quickly identify any mass tag by mass, or any mass tags within a mass range
' The MassTagMass() array contains the masses of all of the mass tags, and is sorted ascending
' The MassTagMassPointer() array gets sorted along with the MassTagMass() array, to allow for dereferencing
' The MassTagLookupInfo() array does not get sorted, but is needed to determine the origin of each mass
'   tag in MassTagMass() and MassTagMassPointer()
Private Type udtORFViewerMassTagDataIndexDataLookupType
    GelIndex As Long            ' Pointer to be used to dereference x in GelORFData(x)
    ORFIndex As Long            ' Pointer to be used to dereference x in GelORFData(GelIndex).Orfs(x)
    MassTagIndex As Long        ' Pointer to be used to dereference x in GelORFMassTags(GelIndex).Orfs(ORFIndex).MassTags(x)
End Type

Private Type udtORFViewerMassTagMassIndexType
    MassTagCount As Long
    MassTagMass() As Double             ' 0-based array.  Sorted, ascending
    MassTagMassPointer() As Long
    
    MassTagLookupInfo() As udtORFViewerMassTagDataIndexDataLookupType
End Type

Private Type udtHistoryItem
    ORFGroupArrayIndex As Long
    ORFDescription As String
End Type

Private Type udtHistory
    History(ORF_HISTORY_COUNT_MAX) As udtHistoryItem
    HistoryCount As Integer
    CurrentHistoryIndex As Integer
End Type

' mGelDataIndices contains arrays holding the Gel Ions and Gel UMC's,
'  sorted by ascending mass
' This greatly speeds the process of finding the data within view
Private mGelDataIndices As udtORFViewerDataIndicesListType

' mORFViewerGroupList contains the list of ORF's to display
' It is actually composed of pointers to the GelORFData() arrays of the gels
'  associated with this instance of frmORFViewer
Private mORFViewerGroupList As udtORFViewerGroupListType

' The following contains a list of all of the masses for all of the mass tags in the
'  ORFs in mORFViewerGroupList, sorted by mass
Private mMassTagDataIndex As udtORFViewerMassTagMassIndexType

' mGelDisplayListAndOptions contains a list of the Gels whose ORF's are to be included
'  in mORFViewerGroupList, and are thus available for browsing in the ORF Viewer
Private mGelDisplayListAndOptions As udtORFViewerGelListType

' mColumnSortFormats() is used to determine how to sort a column in a ListView
Private mColumnSortFormats(LISTVIEW_COUNT) As udtColumnSortFormatType

' mORFHistory() contains a list of the ORFGroupArray index values that have been recently viewed
Private mORFHistory As udtHistory
Private mORFHistoryMenuItemsLoadedCount As Integer

Private mFormID As Long                  ' Unique ID value for this form so that the ORFViewerLoaderClass knows which form to kill/hide when the user closes it; also used when updating options or when user right clicks on a Mass Tag in a ctlORF control
Private mORFPicsLoadedCount As Integer
Private mORFPicsUseCount As Integer
Private mORFPicsRowCount As Integer
Private mORFPicsMaxLoadedCount As Integer       ' Defaults to 100, if more required, then queries user about increasing in chunks of 50

Private mPicWidth As Long, mPicHeight As Long
Private mPicSpacing As Long
Private mXAxisFormatLabel As String
Private mYAxisFormatLabel As String

Private mUserNotifiedOfIonMatchCountError As Boolean
Private mPopulatingSourceGelCombo As Boolean
Private mFormLoaded As Boolean

Private mMassTagListViewORFIndex As Long
Private mMassTagListViewMaxMatchCount As Long

Private mListViewsExpanded As Boolean
Private mSavedDividerXLoc As Long
Private mKeyPressAbortORFListPopulate As Integer

Private mMovingHorizDivider As Boolean, mMovingVerticalDivider As Boolean
Private mDividerMinX As Long, mDividerMaxX As Long
Private mDividerMinY As Long, mDividerMaxY As Long
Private mDividerXLoc As Long
Private mDividerYLoc As Long

Private mMTCountWithHitsForThisORF As Long

Private Sub CheckMoveDividerBars(ByVal xPos As Single, ByVal yPos As Single, blnDragDrop As Boolean, Optional blnOverListView As Boolean, Optional lvwThisListView As MSComctlLib.ListView)
        
    If mMovingHorizDivider Or mMovingVerticalDivider Then
        If blnOverListView Then
            xPos = xPos + lvwThisListView.Left
            yPos = yPos + lvwThisListView.Top
        End If
        
        If yPos < mDividerMinY Then yPos = mDividerMinY
        If yPos > mDividerMaxY Then yPos = mDividerMaxY
    
        If xPos < mDividerMinX Then xPos = mDividerMinX
        If xPos > mDividerMaxX Then xPos = mDividerMaxX
    
        If mMovingHorizDivider Then
            mDividerYLoc = yPos
        ElseIf mMovingVerticalDivider Then
            mDividerXLoc = xPos
        End If
        
        PositionControls False
        
        If blnDragDrop Then
            mMovingHorizDivider = False
            mMovingVerticalDivider = False
        End If
    End If
    
    
End Sub

Private Sub CopySelectedItems(lvwThisListView As MSComctlLib.ListView, lngMaxColumnIndexToCopy As Long, eListViewID As lviORFFormListViewIDConstants)
    Dim lstListItem As MSComctlLib.ListItem
    Dim strOutput As String, strItemText As String
    Dim lngItemIndex As Long
    Dim lngListItemCount As Long
    
    If lvwThisListView.ListItems.Count < 1 Then Exit Sub
    
    strItemText = ""
    For lngItemIndex = 0 To lngMaxColumnIndexToCopy
        strItemText = strItemText & LookupColumnTitle(lngItemIndex, eListViewID)
        If lngItemIndex < lngMaxColumnIndexToCopy Then strItemText = strItemText & vbTab
    Next lngItemIndex
    strOutput = strOutput & strItemText & vbCrLf
    
    Me.MousePointer = vbHourglass
    frmProgress.InitializeForm "Copying selected items", 0, lvwThisListView.ListItems.Count, True, False, True, MDIForm1
    
    For Each lstListItem In lvwThisListView.ListItems
        With lstListItem
            If .Selected Then
                strItemText = .Text & vbTab
                For lngItemIndex = 1 To lngMaxColumnIndexToCopy
                    strItemText = strItemText & .SubItems(lngItemIndex)
                    If lngItemIndex < lngMaxColumnIndexToCopy Then strItemText = strItemText & vbTab
                Next lngItemIndex
                strOutput = strOutput & strItemText & vbCrLf
            End If
        End With
        lngListItemCount = lngListItemCount + 1
        
        If lngListItemCount Mod 50 = 0 Then
            frmProgress.UpdateProgressBar lngListItemCount
        End If
    Next lstListItem
    
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText strOutput, vbCFText
    
    frmProgress.HideForm
    Me.MousePointer = vbDefault
End Sub

Private Sub ShrinkExpandListViews(blnExpandListViews As Boolean)
    mListViewsExpanded = blnExpandListViews
    
    If blnExpandListViews Then
        mSavedDividerXLoc = mDividerXLoc
        mDividerXLoc = mDividerMaxX
    Else
        mDividerXLoc = mSavedDividerXLoc
    End If
    
    cmdRollUpExpand.Visible = Not mListViewsExpanded
    cmdRollUpShrink.Visible = mListViewsExpanded
    
    PositionControls False, True
End Sub


' Unused Function (May 2003)
'''Private Function CountRandomHitsForAllORFs()
'''    Dim lngORFGroupIndex As Long
'''    Dim lngDataCountStats() As Long
'''
'''    Dim lngGelIndex As Long
'''    Dim lngORFIndex As Long
'''
'''    Dim intOutFileNum As Integer
'''    Dim strOutFilePath As String
'''
'''    ReDim lngDataCountStats(mORFViewerGroupList.ORFCount, 2)
'''
'''    frmProgress.InitializeForm "Finding hit counts for all ORF's", 0, mORFViewerGroupList.ORFCount
'''
'''    For lngORFGroupIndex = 0 To mORFViewerGroupList.ORFCount - 1
'''        If Not frmProgress.Visible Then
'''            frmProgress.InitializeForm "Finding hit counts for all ORF's", 0, mORFViewerGroupList.ORFCount
'''        End If
'''
'''        PopulateMassTagsListView lngORFGroupIndex
'''
'''        lngGelIndex = mORFViewerGroupList.Orfs(lngORFGroupIndex).Items(0).GelIndex
'''        lngORFIndex = mORFViewerGroupList.Orfs(lngORFGroupIndex).Items(0).ORFIndex
'''
'''        lngDataCountStats(lngORFGroupIndex, 0) = GelORFMassTags(lngGelIndex).Orfs(lngORFIndex).MassTagCount
'''        lngDataCountStats(lngORFGroupIndex, 1) = mMTCountWithHitsForThisORF
'''
'''        If lngORFGroupIndex Mod 10 = 0 Then
'''            frmProgress.UpdateProgressBar lngORFGroupIndex
'''        End If
'''    Next lngORFGroupIndex
'''
'''    intOutFileNum = FreeFile()
'''
'''    strOutFilePath = GetTemporaryDir() & "MTHitCountsByORF.txt"
'''
'''    Open strOutFilePath For Output As #intOutFileNum
'''
'''    Print #intOutFileNum, "ORF Index " & vbTab & "Mass Tag Count" & vbTab & "Mass Tags with >0 Hits"
'''    For lngORFGroupIndex = 0 To mORFViewerGroupList.ORFCount - 1
'''        Print #intOutFileNum, lngORFGroupIndex & vbTab & lngDataCountStats(lngORFGroupIndex, 0) & vbTab & lngDataCountStats(lngORFGroupIndex, 1)
'''    Next lngORFGroupIndex
'''
'''    Close intOutFileNum
'''
'''    frmProgress.HideForm
'''
'''End Function

Private Function FindIonsInRange(ByRef objMWSearch As MWUtil, ByRef udtDataToPopulate As udtDataToPopulateType, ByVal lngGelIndex As Long, ByVal dblAMTNet As Double, ByVal dblNETRangeHalfWindow As Double, ByVal dblAMTMass As Double, ByVal dblMassRangeHalfWindow As Double, ByVal lngMassTagID As Long, ByRef lngAMTMatchCount As Long) As Long
    ' Examines the ions for the desired Gel to determine those
    '   values that are within range, returning the data in udtDataToPopulate()
    ' Function returns the number of data points in udtDataToPopulate()
    
    Const DATA_DIM_CHUNK = 10
    
    Dim lngFirstIndex As Long, lngLastIndex As Long
    
    Dim lngPointerIndex As Long
    Dim lngIonIndex As Long
    Dim lngDataCount As Long, lngDataCountDimmed As Long
    Dim dblIonNET As Double
    
    Dim strDescription As String
    Dim blnAMTFound As Boolean
    
    lngAMTMatchCount = 0
    
    lngDataCountDimmed = DATA_DIM_CHUNK
    InitializeDataToPopulate udtDataToPopulate, lngDataCountDimmed, False
    
    lngDataCount = 0
    
    If mGelDataIndices.Gels(lngGelIndex).IonCount = 0 Then
        udtDataToPopulate.Count = 0
        FindIonsInRange = 0
        Exit Function
    End If
    
    With udtDataToPopulate
        
        ' Fill objMWSearch with the sorted .IonMass() values
        If Not objMWSearch.Fill(mGelDataIndices.Gels(lngGelIndex).IonMass()) Then
            ' Failure initializing objMWSearch
            Debug.Assert False
        End If

        If objMWSearch.FindIndexRange(dblAMTMass, dblMassRangeHalfWindow, lngFirstIndex, lngLastIndex) Then
            For lngPointerIndex = lngFirstIndex To lngLastIndex
                
                lngIonIndex = mGelDataIndices.Gels(lngGelIndex).IonMassPointer(lngPointerIndex)
                
                '' The following assertion can be used to test the index arrays
                '' However, enabling it will decrease execution time in the IDE
                'Debug.Assert mGelDataIndices.Gels(lngGelIndex).IonMass(lngPointerIndex) = GetIsoMass(GelData(lngGelIndex).IsoData(lngIonIndex), GelData(lngGelIndex).Preferences.IsoDataField)
                
                ' Mass is within tolerance; now check NET
                dblIonNET = mGelDataIndices.Gels(lngGelIndex).IonNET(lngPointerIndex)

                '' The following assertion can be used to test the index arrays
                '' However, enabling it will decrease execution time in the IDE
                'Debug.Assert dblIonNET = ScanToGANET(lngGelIndex, GelData(lngGelIndex).IsoData(lngIonIndex).ScanNumber)

                If WithinToleranceDbl(dblIonNET, dblAMTNet, dblNETRangeHalfWindow) Then
                    ' Include ion

                    ' Examine strDescription to see if AMT:lngMasssTagID is present
                    ' If it is, move it to the front of the string
                    With GelData(lngGelIndex)
                        If IsNull(.IsoData(lngIonIndex).MTID) Then
                            strDescription = ""
                        Else
                            strDescription = .IsoData(lngIonIndex).MTID
                        End If
                        blnAMTFound = FindMassTagIDInDescription(strDescription, lngMassTagID)

                        If blnAMTFound Then
                            lngAMTMatchCount = lngAMTMatchCount + 1
                        End If
                    End With

                    With udtDataToPopulate
                        .Labels(lngDataCount) = ORF_VIEWER_ION_STRING & Trim(lngIonIndex) & ORF_VIEWER_ID_DELIMETER & " " & strDescription
                        .NET(lngDataCount) = dblIonNET
                        .NETExtentNeg(lngDataCount) = 0
                        .NETExtentPos(lngDataCount) = 0
                        
                        .Mass(lngDataCount) = mGelDataIndices.Gels(lngGelIndex).IonMass(lngPointerIndex)
                        .MassExtentNeg(lngDataCount) = 0
                        .MassExtentPos(lngDataCount) = 0
                        
                        .Intensity(lngDataCount) = GelData(lngGelIndex).IsoData(lngIonIndex).Abundance

                        ' Need to scale the intensity if plotting both Ions and UMC's
                        If mGelDisplayListAndOptions.DisplayOptions.DataDisplayMode = ddmIonsAndUMCs Then
                            .Intensity(lngDataCount) = .Intensity(lngDataCount) * mGelDisplayListAndOptions.DisplayOptions.IonToUMCPlottingIntensityRatio
                        End If
                        
                        lngDataCount = lngDataCount + 1
                        If lngDataCount >= lngDataCountDimmed Then
                            lngDataCountDimmed = lngDataCountDimmed + DATA_DIM_CHUNK
                            InitializeDataToPopulate udtDataToPopulate, lngDataCountDimmed, True
                        End If
                    End With
                End If
            Next lngPointerIndex
        End If
    End With
    
    udtDataToPopulate.Count = lngDataCount
    FindIonsInRange = lngDataCount

End Function

Private Function FindUMCsInRange(ByRef objMWSearch As MWUtil, ByRef udtDataToPopulate As udtDataToPopulateType, ByVal lngGelIndex As Long, ByVal dblAMTNet As Double, ByVal dblNETRangeHalfWindow As Double, ByVal dblAMTMass As Double, ByVal dblMassRangeHalfWindow As Double, ByRef lngAMTMatchCount As Long, ByVal blnUseClassRepresentativeNET As Boolean) As Long
    ' Examines the UMC's for the desired Gel to determine those
    '   values that are within range, returning the data in udtDataToPopulate()
    ' Function returns the number of data points in udtDataToPopulate()
    
    Const DATA_DIM_CHUNK = 10
    
    Static LastErrorWarnTime As Date
    Dim lngFirstIndex As Long, lngLastIndex As Long
    Dim lngFirstIndexMassMax As Long, lngLastIndexMassMax As Long
    
    Dim lngPointerIndex As Long, lngCompareIndex As Long
    Dim lngNewUMCIndex As Long
    Dim blnMatched As Boolean
    
    Dim lngUMCIndex As Long
    Dim lngDataCount As Long, lngDataCountDimmed As Long
    Dim dblUMCNetStart As Double, dblUMCNetEnd As Double
    Dim dblUMCNet As Double
    Dim dblNETExtentNeg As Double, dblNETExtentPos As Double
        
    Dim lngMemberIndex As Long
    
    Dim blnMatchFoundViaMassMin As Boolean, blnMatchFoundViaMassMax As Boolean
    
    Dim UMCMassMatchIndices() As Long
    Dim UMCMassMatchCount As Long
    
    Dim strDescription As String
    
On Error GoTo FindUMCsInRangeErrorHandler

    lngAMTMatchCount = 0
    
    lngDataCountDimmed = DATA_DIM_CHUNK
    InitializeDataToPopulate udtDataToPopulate, lngDataCountDimmed, False
    
    lngDataCount = 0
    
    If mGelDataIndices.Gels(lngGelIndex).UMCCount = 0 Then
        udtDataToPopulate.Count = 0
        FindUMCsInRange = 0
        Exit Function
    End If
    
    With udtDataToPopulate
        
        ' Fill objMWSearch with the sorted .UMCMassMin() values
        If Not objMWSearch.Fill(mGelDataIndices.Gels(lngGelIndex).UMCMassMin()) Then
            ' Failure initializing objMWSearch
            Debug.Assert False
        End If
                
        ' Find the UMC's within range
        blnMatchFoundViaMassMin = objMWSearch.FindIndexRange(dblAMTMass, dblMassRangeHalfWindow, lngFirstIndex, lngLastIndex)
        
        ' Now fill objMWSearch with the sorted .UMCMassMax() values and repeat the search
        If Not objMWSearch.Fill(mGelDataIndices.Gels(lngGelIndex).UMCMassMax()) Then
            ' Failure initializing objMWSearch
            Debug.Assert False
        End If
        
        blnMatchFoundViaMassMax = objMWSearch.FindIndexRange(dblAMTMass, dblMassRangeHalfWindow, lngFirstIndexMassMax, lngLastIndexMassMax)
        
        If blnMatchFoundViaMassMin Or blnMatchFoundViaMassMax Then
            ' Merge the results of the two UMC searches
                    
            UMCMassMatchCount = 0
            ReDim UMCMassMatchIndices(0)
            
            If blnMatchFoundViaMassMin Then
                UMCMassMatchCount = lngLastIndex - lngFirstIndex + 1
                ReDim UMCMassMatchIndices(UMCMassMatchCount)
                For lngPointerIndex = lngFirstIndex To lngLastIndex
                    UMCMassMatchIndices(lngPointerIndex - lngFirstIndex) = mGelDataIndices.Gels(lngGelIndex).UMCMassMinPointer(lngPointerIndex)
                Next lngPointerIndex
            End If
            
            If blnMatchFoundViaMassMax Then
                ReDim Preserve UMCMassMatchIndices(UMCMassMatchCount + lngLastIndexMassMax - lngFirstIndexMassMax + 1)
                For lngPointerIndex = lngFirstIndexMassMax To lngLastIndexMassMax
                    lngNewUMCIndex = mGelDataIndices.Gels(lngGelIndex).UMCMassMaxPointer(lngPointerIndex)
                    blnMatched = False
                    For lngCompareIndex = 0 To UMCMassMatchCount - 1
                        If UMCMassMatchIndices(lngCompareIndex) = lngNewUMCIndex Then
                            blnMatched = True
                            Exit For
                        End If
                    Next lngCompareIndex
                    
                    If Not blnMatched Then
                        UMCMassMatchIndices(UMCMassMatchCount) = lngNewUMCIndex
                        UMCMassMatchCount = UMCMassMatchCount + 1
                    End If
                Next lngPointerIndex
            End If
            
            ' Step through the UMC's that were found to be in the mass range
            ' See if within the NET range
            For lngPointerIndex = 0 To UMCMassMatchCount - 1
                lngUMCIndex = UMCMassMatchIndices(lngPointerIndex)
                
'''                Debug.Assert mGelDataIndices.Gels(lngGelIndex).UMCMassMin(lngPointerIndex) = GelUMC(lngGelIndex).UMCs(lngUMCIndex).ClassMW
'''                Debug.Assert mGelDataIndices.Gels(lngGelIndex).UMCMassMax(lngPointerIndex) = GelUMC(lngGelIndex).UMCs(lngUMCIndex).ClassMW
                
                ' Mass is within tolerance; now check NET
                dblUMCNetStart = mGelDataIndices.Gels(lngGelIndex).UMCScanRange(lngUMCIndex).NETStart
                dblUMCNetEnd = mGelDataIndices.Gels(lngGelIndex).UMCScanRange(lngUMCIndex).NETEnd
                
                If WithinToleranceDbl(dblUMCNetStart, dblAMTNet, dblNETRangeHalfWindow) Or _
                   WithinToleranceDbl(dblUMCNetEnd, dblAMTNet, dblNETRangeHalfWindow) Then
                    ' Include UMC

                    ' List the ions that are members of this UMC, separating them by commas
                    ' Add ** if the ion is the class representative
                    With GelUMC(lngGelIndex).UMCs(lngUMCIndex)
                        If .ClassCount = 1 Then
                            strDescription = Trim(.ClassCount) & " Ion" & ORF_VIEWER_UMC_ION_LIST_START_STRING
                        Else
                            strDescription = Trim(.ClassCount) & " Ions" & ORF_VIEWER_UMC_ION_LIST_START_STRING
                        End If
                        
                        For lngMemberIndex = 0 To .ClassCount - 1
                            strDescription = strDescription & .ClassMInd(lngMemberIndex)
                            If .ClassRepInd = .ClassMInd(lngMemberIndex) Then
                                strDescription = strDescription & ORF_VIEWER_UMC_REPRESENTATIVE_MEMBER_INDICATOR
                            End If
                            If lngMemberIndex < .ClassCount - 1 Then
                                strDescription = strDescription & ORF_VIEWER_UMC_ION_LIST_DELIMETER
                            End If
                        Next lngMemberIndex
                        
                    End With
                    
                    If blnUseClassRepresentativeNET Then
                        ' Could use the NET of the Class representative
                        dblUMCNet = mGelDataIndices.Gels(lngGelIndex).UMCScanRange(lngUMCIndex).NETClassRepresentative
                    Else
                        ' Or could use the average NET value of the class
                        dblUMCNet = (dblUMCNetStart + dblUMCNetEnd) / 2
                    End If
                    dblNETExtentNeg = dblUMCNet - dblUMCNetStart
                    dblNETExtentPos = dblUMCNetEnd - dblUMCNet
                    
                    With udtDataToPopulate
                        .Labels(lngDataCount) = ORF_VIEWER_UMC_STRING & Trim(lngUMCIndex) & ORF_VIEWER_ID_DELIMETER & " " & strDescription
                        .NET(lngDataCount) = dblUMCNet
                        .NETExtentNeg(lngDataCount) = dblNETExtentNeg
                        .NETExtentPos(lngDataCount) = dblNETExtentPos
                        
                        .Mass(lngDataCount) = GelUMC(lngGelIndex).UMCs(lngUMCIndex).ClassMW
                        .MassExtentNeg(lngDataCount) = 0
                        .MassExtentPos(lngDataCount) = 0
                        
                        .Intensity(lngDataCount) = GelUMC(lngGelIndex).UMCs(lngUMCIndex).ClassAbundance

                        lngDataCount = lngDataCount + 1
                        If lngDataCount >= lngDataCountDimmed Then
                            lngDataCountDimmed = lngDataCountDimmed + DATA_DIM_CHUNK
                            InitializeDataToPopulate udtDataToPopulate, lngDataCountDimmed, True
                        End If
                    End With
                End If
            Next lngPointerIndex
        End If
    End With
    
    udtDataToPopulate.Count = lngDataCount
    FindUMCsInRange = lngDataCount
    Exit Function
    
FindUMCsInRangeErrorHandler:
    If Now() > LastErrorWarnTime + 1# / 60# / 24# Then
        MsgBox "Error while finding the UMC's matching the current ORF.  You may need to re-initialize the match arrays using File->Refresh ORF List and Source Data. This error will occur if the master UMC list has changed since the last time the ORF viewer was used.", vbExclamation + vbOKOnly, "Error"
        LastErrorWarnTime = Now()
    End If

End Function

Private Function FindMassTagsInRange(ByRef objMWSearch As MWUtil, ByRef udtDataToPopulate As udtDataToPopulateType, ByVal dblAMTNet As Double, ByVal dblNETRangeHalfWindow As Double, ByVal dblAMTMass As Double, ByVal dblMassRangeHalfWindow As Double, ByVal dblMassTagNETError As Double, ByVal dblMassTagErrorPPM As Double, ByVal lngCentralMassTagID As Long, ByVal lngCentralMassTagORFIndex As Long) As Long
    ' Examines the mass tags in mMassTagDataIndex.MassTagMass() to find those within
    '  the given search range, returning them in udtDataToPopulate
    ' Function returns the number of data points in udtDataToPopulate()
    ' Note: lngCentralMassTagID is the mass tag whose value should be in the center of the search range
    '       I'm sending it to this function so that we can make sure the mass tag is found
        
    Const DATA_DIM_CHUNK = 5
    
    Dim lngFirstIndex As Long, lngLastIndex As Long
    Dim lngPointerIndex As Long
    Dim lngMassTagLookupArrayIndex As Long
    Dim lngGelIndex As Long
    Dim lngORFIndex As Long
    Dim lngMassTagIndex As Long
    Dim lngDataCount As Long, lngDataCountDimmed As Long
    Dim blnCentralMassTagFound As Boolean
    Dim blnProceed As Boolean
    
    Dim strDescription As String
    
    lngDataCountDimmed = DATA_DIM_CHUNK
    InitializeDataToPopulate udtDataToPopulate, lngDataCountDimmed, False
    
    With udtDataToPopulate
        
        ' Fill objMWSearch with the sorted .MassTagMass() values
        If Not objMWSearch.Fill(mMassTagDataIndex.MassTagMass()) Then
            ' Failure initializing objMWSearch
            Debug.Assert False
        End If

        lngDataCount = 0
        If objMWSearch.FindIndexRange(dblAMTMass, dblMassRangeHalfWindow, lngFirstIndex, lngLastIndex) Then
            For lngPointerIndex = lngFirstIndex To lngLastIndex
                
                lngMassTagLookupArrayIndex = mMassTagDataIndex.MassTagMassPointer(lngPointerIndex)
                
                With mMassTagDataIndex.MassTagLookupInfo(lngMassTagLookupArrayIndex)
                    lngGelIndex = .GelIndex
                    lngORFIndex = .ORFIndex
                    lngMassTagIndex = .MassTagIndex
                End With
                
                ' Mass is within tolerance; now check NET
                With GelORFMassTags(lngGelIndex).Orfs(lngORFIndex).MassTags(lngMassTagIndex)
        
                    ' Don't include the mass tag if it is a Theoretical mass tag that doesn't
                    '  belong to the same ORF as the central mass tag
                    blnProceed = True
                    If .MassTagRefID <= -10000 Then
                        If .MassTagRefID <> lngCentralMassTagID Then
                            If lngORFIndex <> lngCentralMassTagORFIndex Then
                                blnProceed = False
                            End If
                        End If
                    End If
                    
                    If blnProceed Then
                        If .MassTagRefID = lngCentralMassTagID Then
                            blnCentralMassTagFound = True
                        End If
                    
                        If WithinToleranceDbl(.GANET, dblAMTNet, dblNETRangeHalfWindow) Then
                            ' Include mass tag
                            
                            strDescription = "MW=" & Round(.Mass, 3) & " Da, NET=" & Round(.GANET, 2)
                            
                            udtDataToPopulate.Labels(lngDataCount) = ORF_VIEWER_MASS_TAG_STRING & Trim(.MassTagRefID) & ORF_VIEWER_ID_DELIMETER & " " & strDescription
                            udtDataToPopulate.NET(lngDataCount) = .GANET
                            udtDataToPopulate.NETExtentNeg(lngDataCount) = dblMassTagNETError / 2
                            udtDataToPopulate.NETExtentPos(lngDataCount) = dblMassTagNETError / 2
                            udtDataToPopulate.Mass(lngDataCount) = .Mass
                            udtDataToPopulate.MassExtentNeg(lngDataCount) = PPMToMass(dblMassTagErrorPPM, .Mass) / 2
                            udtDataToPopulate.MassExtentPos(lngDataCount) = udtDataToPopulate.MassExtentNeg(lngDataCount)
                            udtDataToPopulate.Intensity(lngDataCount) = 0
                            
                            lngDataCount = lngDataCount + 1
                            If lngDataCount >= lngDataCountDimmed Then
                                lngDataCountDimmed = lngDataCountDimmed + DATA_DIM_CHUNK
                                InitializeDataToPopulate udtDataToPopulate, lngDataCountDimmed, True
                            End If
                        End If
                    End If
                End With
            Next lngPointerIndex
        End If
        
'        Debug.Assert blnCentralMassTagFound
            
        ' Make sure the "central" mass tag is the last one in udtDataToPopulate, to assure it gets drawn on top
        
    End With
    
    
    FindMassTagsInRange = lngDataCount

End Function

Private Sub HighlightMassTagsForSelectedORF()
    Dim lngSourceGelIndex As Long
    Dim lngORFDataIndexSelected As Long
    Dim lngIonMatchIndex  As Long
    
    lngORFDataIndexSelected = GetSelectedORFDataIndex()
    If lngORFDataIndexSelected < 0 Then Exit Sub
    
    ' Determine the desired parent gel index
    lngSourceGelIndex = GetDesiredParentGelIndex()
    If lngSourceGelIndex < 1 Then Exit Sub
    
    ' Clear any currently selected spots
    With GelBody(lngSourceGelIndex)
        .fgSelProtected = False
        .mnu2lsLockSelection.Checked = False
        .GelSel.Clear
    End With
    
    With GelORFData(lngSourceGelIndex).Orfs(lngORFDataIndexSelected)
        For lngIonMatchIndex = 0 To .IonMatchCount - 1
            GelBody(lngSourceGelIndex).GelSel.AddToIsoSelection .IonMatches(lngIonMatchIndex).IonDataIndex
        Next lngIonMatchIndex
    End With
    
    With GelBody(lngSourceGelIndex)
        .fgSelProtected = True
        .mnu2lsLockSelection.Checked = True
    End With

    GelDrawScreen lngSourceGelIndex
End Sub

Private Sub InitializeDataToPopulate(udtDataToPopulate As udtDataToPopulateType, lngMaxArrayIndex As Long, Optional blnPreserveData As Boolean = False)
    With udtDataToPopulate
        If blnPreserveData Then
            ReDim Preserve .Labels(0 To lngMaxArrayIndex)
            ReDim Preserve .NET(0 To lngMaxArrayIndex)
            ReDim Preserve .NETExtentNeg(0 To lngMaxArrayIndex)
            ReDim Preserve .NETExtentPos(0 To lngMaxArrayIndex)
            ReDim Preserve .Mass(0 To lngMaxArrayIndex)
            ReDim Preserve .MassExtentNeg(0 To lngMaxArrayIndex)
            ReDim Preserve .MassExtentPos(0 To lngMaxArrayIndex)
            ReDim Preserve .Intensity(0 To lngMaxArrayIndex)
        Else
            ReDim .Labels(0 To lngMaxArrayIndex)
            ReDim .NET(0 To lngMaxArrayIndex)
            ReDim .NETExtentNeg(0 To lngMaxArrayIndex)
            ReDim .NETExtentPos(0 To lngMaxArrayIndex)
            ReDim .Mass(0 To lngMaxArrayIndex)
            ReDim .MassExtentNeg(0 To lngMaxArrayIndex)
            ReDim .MassExtentPos(0 To lngMaxArrayIndex)
            ReDim .Intensity(0 To lngMaxArrayIndex)
        End If
    End With

End Sub

Private Function FindMassTagIDInDescription(ByRef strDescription As String, ByVal lngMassTagID As Long, Optional blnMoveToFrontOfStringIfFound As Boolean = True) As Boolean
    ' Looks for the text "AMT:lngMassTagID" in strDescription, returning True if found
    ' If blnMoveToFrontOfStringIfFound = True, then moves the text, along with any other text following,
    '   up to the the next semicolon, to the front of strDescription
    Dim intCharLocStart As Long, intCharLocEnd As Long
    Dim strMatchText As String, strLeftPortion As String, strRightPortion As String
    
    ' AMTMark = "AMT:"
    intCharLocStart = InStr(UCase(strDescription), AMTMark & Trim(lngMassTagID))
    
    If intCharLocStart > 0 Then
        If blnMoveToFrontOfStringIfFound Then
            intCharLocEnd = InStr(Mid(strDescription, intCharLocStart), ";")
            If intCharLocEnd > 0 Then
                intCharLocEnd = intCharLocEnd + intCharLocStart - 1
                strRightPortion = Mid(strDescription, intCharLocEnd + 1)
            Else
                intCharLocEnd = Len(strDescription)
            End If
            
            strMatchText = Mid(strDescription, intCharLocStart, intCharLocEnd - intCharLocStart + 1)
            If intCharLocStart > 1 Then
                strLeftPortion = Left(strDescription, intCharLocStart - 1)
            End If
            
            strDescription = strMatchText & strLeftPortion & strRightPortion
        End If
        FindMassTagIDInDescription = True
    Else
        FindMassTagIDInDescription = False
    End If
    
End Function

Public Function FindORFContainingMassTag(blnQueryUserForMassTagID As Boolean, Optional lngMassTagID As Long) As Boolean
    ' Looks for the ORF containing MassTagID, querying the user if blnQueryUserForMassTagID = true
    ' Returns True if found, False if not found
    Dim lngORFListViewItemIndex As Long, lngGelIndex As Long, lngORFIndex As Long
    Dim lngItemIndex As Long
    Dim lngORFGroupArrayIndex As Long
    Dim lngMassTagIndex As Long
    Dim strMassTagID As String
    Dim blnMatchFound As Boolean
    
    If blnQueryUserForMassTagID Then
        strMassTagID = InputBox("Enter the Mass Tag ID to search for", "Find Mass Tag ID")
        If Not IsNumeric(strMassTagID) Then
            FindORFContainingMassTag = False
            Exit Function
        End If
        lngMassTagID = CLngSafe(strMassTagID)
    End If
    
    For lngORFListViewItemIndex = 1 To lvwORFs.ListItems.Count
        lngORFGroupArrayIndex = CLngSafe(lvwORFs.ListItems(lngORFListViewItemIndex).Text)
        With mORFViewerGroupList.Orfs(lngORFGroupArrayIndex)
            For lngItemIndex = 0 To .ItemCount - 1
                lngGelIndex = .Items(lngItemIndex).GelIndex
                lngORFIndex = .Items(lngItemIndex).ORFIndex
                
                For lngMassTagIndex = 0 To GelORFMassTags(lngGelIndex).Orfs(lngORFIndex).MassTagCount - 1
                        If GelORFMassTags(lngGelIndex).Orfs(lngORFIndex).MassTags(lngMassTagIndex).MassTagRefID = lngMassTagID Then
                            ' Match Found
                            blnMatchFound = True
                            Exit For
                        End If
                Next lngMassTagIndex
                If blnMatchFound Then Exit For
            Next lngItemIndex
        End With
        
        If blnMatchFound Then
            ' Do not show the ORF if it is the current one
            If lngORFGroupArrayIndex <> mMassTagListViewORFIndex Then
                ' Find the item in lvwORFs with ID = lngORFGroupArrayIndex
                SelectORFByID lngORFGroupArrayIndex
                
                UpdateStatus "Target mass tag found: " & lngMassTagID
            End If
            
            Exit For
        End If
    Next
    
    FindORFContainingMassTag = blnMatchFound
End Function

Private Sub FindTextInListViewWrapper(eListViewID As lviORFFormListViewIDConstants)
    Dim strSearchText As String
    Dim lngIndexOfMatch As Long
    
    strSearchText = InputBox("Enter the text to search for.", "Find", ListViewGetSearchHistoryText(CInt(eListViewID)))
    
    If Len(strSearchText) = 0 Then Exit Sub

On Error Resume Next
    
    Select Case eListViewID
    Case lviORFs
        lngIndexOfMatch = ListViewFindText(strSearchText, lvwORFs, CInt(eListViewID), lvoSortkey - 1)
        If lngIndexOfMatch >= 0 Then HandleListViewItemClick lvwORFs.ListItems(lngIndexOfMatch), eListViewID
    Case Else
        Debug.Assert False
    End Select
    
End Sub

Private Function AllGelsForItemCount(ByRef lngSourceGelIndex As Long) As Boolean
    ' Examines cboItemCountSourceGel
    ' If "All Gels" is selected, then returns True
    ' Otherwise, returns false and returns the selecte source Gel index, using lngSourceGelIndex
    
    Dim blnUseAllGels As Boolean
    
    Dim strListItem As String
    Dim intColonLoc As Integer
    
    If cboItemCountSourceGel.ListIndex <= 0 Then
        blnUseAllGels = True
    Else
        blnUseAllGels = False
                
        strListItem = cboItemCountSourceGel.List(cboItemCountSourceGel.ListIndex)
        intColonLoc = InStr(strListItem, ":")
        If intColonLoc > 0 Then
            lngSourceGelIndex = CLngSafe(Left(strListItem, intColonLoc - 1))
        Else
            ' This is unexpected
            Debug.Assert False
        End If
            
        If lngSourceGelIndex < 1 Then
            ' This is unexpected
            Debug.Assert False
            lngSourceGelIndex = 1
        End If
    End If
    
    AllGelsForItemCount = blnUseAllGels
End Function

Private Function GetDesiredMatchCount(ByVal lngORFGroupArrayIndex As Long, ByVal blnReturnUMCCount As Boolean, ByRef dblReturnIntensity As Double, ByVal blnUseAllGels As Boolean, Optional lngSourceGelIndex As Long = -1, Optional blnFilterByMassTagID As Boolean = False, Optional lngMassTagIDFilter As Long = 0) As Long
    ' Returns number of ion matches or UMC matches in the given ORF
    ' If blnFilterByMassTagID = True, then only returns those matches for a given mass tag
    ' Further, If blnFilterByMassTagID = True, then returns a count of the number of mass tags containing an Ion Hit or a UMC Hit, not the total number of Ions or UMC's that hit mass tags for the ORF
    ' In addition, returns either the representative ORF intensity in dblReturnIntensity, or
    '  the sum of the ions for a single mass tag
    
    Static LastErrorWarnTime As Date
    Dim lngItemIndex As Long
    Dim lngGelIndex As Long, lngORFIndex As Long
    Dim lngIonMatchIndex As Long
    Dim lngUMCMatchIndex As Long
    Dim lngMassTagIndex As Long
    Dim blnIntensityAverageInitialized As Boolean
    Dim blnProceed As Boolean
    
    Dim lngIonMatchCount As Long
    
    Dim dblFullORFSum As Double
    Dim lngFullORFMTCountToAverage As Long      ' Count of the mass tags to use when computing the ORF average
    Dim lngFullORFMTCount As Long               ' Count of the mass tags that have a hit (regardless of whether it is used in the average)
    
    Dim MTIntensityAverageCount As Long
    Dim MTIntensityAverage() As udtMassTagIntensitySumType
    Dim MTIntensitySums() As udtMassTagIntensitySumType
    Dim dblMaximumMassTagSum As Double
    Dim dbl50PctOfMaximumMassTagSum As Double
    
On Error GoTo GetDesiredMatchCountErrorHandler

    If lngSourceGelIndex = -1 Then blnUseAllGels = True
    
    ReDim MTIntensityAverage(0)
    
    lngIonMatchCount = 0
    dblMaximumMassTagSum = 0
    blnIntensityAverageInitialized = False
    For lngItemIndex = 0 To mORFViewerGroupList.Orfs(lngORFGroupArrayIndex).ItemCount - 1
        lngGelIndex = mORFViewerGroupList.Orfs(lngORFGroupArrayIndex).Items(lngItemIndex).GelIndex
        lngORFIndex = mORFViewerGroupList.Orfs(lngORFGroupArrayIndex).Items(lngItemIndex).ORFIndex
            
        If lngGelIndex = lngSourceGelIndex Or blnUseAllGels Then
            
            ReDim MTIntensitySums(GelORFMassTags(lngGelIndex).Orfs(lngORFIndex).MassTagCount)
            
            With GelORFData(lngGelIndex).Orfs(lngORFIndex)
                If blnReturnUMCCount Then
                    For lngUMCMatchIndex = 0 To .UMCMatchCount - 1
                        lngMassTagIndex = .UMCMatches(lngUMCMatchIndex).MassTagIndex
                        
                        MTIntensitySums(lngMassTagIndex).MassTagRefID = GelORFMassTags(lngGelIndex).Orfs(lngORFIndex).MassTags(lngMassTagIndex).MassTagRefID
                        If Not blnFilterByMassTagID Or MTIntensitySums(lngMassTagIndex).MassTagRefID = lngMassTagIDFilter Then
                            MTIntensitySums(lngMassTagIndex).Count = MTIntensitySums(lngMassTagIndex).Count + 1
                            MTIntensitySums(lngMassTagIndex).Sum = MTIntensitySums(lngMassTagIndex).Sum + GelUMC(lngGelIndex).UMCs(.UMCMatches(lngUMCMatchIndex).UMCDataIndex).ClassAbundance
                            
                        End If
                    Next lngUMCMatchIndex
                Else
                    For lngIonMatchIndex = 0 To .IonMatchCount - 1
                        lngMassTagIndex = .IonMatches(lngIonMatchIndex).MassTagIndex
                        
                        MTIntensitySums(lngMassTagIndex).MassTagRefID = GelORFMassTags(lngGelIndex).Orfs(lngORFIndex).MassTags(lngMassTagIndex).MassTagRefID
                        If Not blnFilterByMassTagID Or MTIntensitySums(lngMassTagIndex).MassTagRefID = lngMassTagIDFilter Then
                            MTIntensitySums(lngMassTagIndex).Count = MTIntensitySums(lngMassTagIndex).Count + 1
                            MTIntensitySums(lngMassTagIndex).Sum = MTIntensitySums(lngMassTagIndex).Sum + GelData(lngGelIndex).IsoData(.IonMatches(lngIonMatchIndex).IonDataIndex).Abundance
                        End If
                    Next lngIonMatchIndex
                End If
            End With
        
            ' This will always be true for lngItemIndex = 0
            If UBound(MTIntensityAverage()) < GelORFMassTags(lngGelIndex).Orfs(lngORFIndex).MassTagCount Then
                MTIntensityAverageCount = GelORFMassTags(lngGelIndex).Orfs(lngORFIndex).MassTagCount
                ReDim Preserve MTIntensityAverage(MTIntensityAverageCount)
            End If
            
            For lngMassTagIndex = 0 To GelORFMassTags(lngGelIndex).Orfs(lngORFIndex).MassTagCount - 1
                
                With MTIntensitySums(lngMassTagIndex)
                    ' Store the sum of the ion intensities for this mass tag
                    ' Note that blnIntensityAverageInitialized is False for Item 0, but true for subsequent items
                    If .Count > 0 Then
                        blnProceed = True
                        If MTIntensityAverage(lngMassTagIndex).MassTagRefID > 0 Then
                            ' Make sure the Mass Tag RefID's match
                            If MTIntensityAverage(lngMassTagIndex).MassTagRefID <> .MassTagRefID Then
                                blnProceed = False
                            End If
                        Else
                            MTIntensityAverage(lngMassTagIndex).MassTagRefID = .MassTagRefID
                        End If
                    
                        If blnProceed Then
                            ' Increment the count
                            lngIonMatchCount = lngIonMatchCount + .Count
                        
                            MTIntensityAverage(lngMassTagIndex).Count = MTIntensityAverage(lngMassTagIndex).Count + 1
                            MTIntensityAverage(lngMassTagIndex).Sum = MTIntensityAverage(lngMassTagIndex).Sum + .Sum
                            
                            If MTIntensityAverage(lngMassTagIndex).Sum > dblMaximumMassTagSum Then
                                dblMaximumMassTagSum = MTIntensityAverage(lngMassTagIndex).Sum
                            End If
                        Else
                            ' The Mass tags for the given ORF in the gel for this item do not match up
                            '  with the mass tags for the given ORF in the previous items
                            ' Since this means the user isn't comparing comparable gels, for speed purposes, I will
                            '  not try to fix this, and will instead warn the user
                            If Not mUserNotifiedOfIonMatchCountError Then
                                MsgBox "The mass tags for the selected ORF are not the same for all of the displayed gels.  Ion counts will be incorrect.  You can fix this problem by making sure all of the displayed gels reference the same mass tag database (use 'Edit->Display Parameters and Paths'), and by choosing Reload ORF's from Mass Tag database to update the ORF's.", vbExclamation + vbOKOnly, "Error"
                            End If
                            mUserNotifiedOfIonMatchCountError = True
                            Debug.Assert False
                        
                        End If
                    End If
                End With
            Next lngMassTagIndex
        
            blnIntensityAverageInitialized = True
        End If
    
    Next lngItemIndex

    If blnFilterByMassTagID Then
        ' Return the summed intensity for this mass tag
        ' Note that MTIntensityAverage(lngMassTagIndex).MassTagRefID is initialized to 0
        '   and is not set unless a match is found to it
        dblReturnIntensity = 0
        For lngMassTagIndex = 0 To MTIntensityAverageCount - 1
            If MTIntensityAverage(lngMassTagIndex).MassTagRefID = lngMassTagIDFilter Then
                dblReturnIntensity = MTIntensityAverage(lngMassTagIndex).Sum
            End If
        Next lngMassTagIndex
        
        ' Return the number of mass tags containing hits
        GetDesiredMatchCount = lngIonMatchCount
    Else
        ' Compute the average intensity for the entire ORF
        ' To find the intensity for the ORF, we take the average of the mass tag ion sums
        ' If blnOnlyUseTop50PctForAveraging = True, then compute the average using only those
        '   mass tags whose summed intensity is at least 50% of the highest mass tag summed intensity
        dbl50PctOfMaximumMassTagSum = dblMaximumMassTagSum / 2
        For lngMassTagIndex = 0 To MTIntensityAverageCount - 1
            If MTIntensityAverage(lngMassTagIndex).Count > 0 Then
                
                ' Increment the mass tag hit count
                lngFullORFMTCount = lngFullORFMTCount + 1
                
                If Not mGelDisplayListAndOptions.DisplayOptions.OnlyUseTop50PctForAveraging Or _
                   MTIntensityAverage(lngMassTagIndex).Sum >= dbl50PctOfMaximumMassTagSum Then
                    ' Only increment the count ot average (and the Sum) if mass tag meets the above criteria
                    lngFullORFMTCountToAverage = lngFullORFMTCountToAverage + 1
                    dblFullORFSum = dblFullORFSum + MTIntensityAverage(lngMassTagIndex).Sum
                End If
            End If
        Next lngMassTagIndex
        
        If lngFullORFMTCountToAverage > 0 Then
            dblReturnIntensity = dblFullORFSum / lngFullORFMTCountToAverage
        Else
            dblReturnIntensity = 0
        End If
        
        ' Return the number of hits for this mass tag
        GetDesiredMatchCount = lngFullORFMTCount
    End If
    
    
    Exit Function
    
GetDesiredMatchCountErrorHandler:
    If Now() > LastErrorWarnTime + 1# / 60# / 24# Then
        If blnReturnUMCCount Then
            MsgBox "Error while determining the number of UMC's matching the current ORF.  You may need to re-initialize the match arrays using File->Refresh ORF List and Source Data. This error will occur if the master UMC list has changed since the last time the ORF viewer was used.", vbExclamation + vbOKOnly, "Error"
        Else
            MsgBox "Error while determining the number of ion's matching the current ORF.  You may need to re-initialize the match arrays using File->Refresh ORF List and Source Data.", vbExclamation + vbOKOnly, "Error"
        End If
        LastErrorWarnTime = Now()
    End If
    
End Function

Private Function GetDesiredParentGelIndex() As Long
    ' Examines cboItemCountSourceGel
    ' Returns the desired Gel Index if one is selected
    ' Otherwise, returns the lowest gel index displayed in the ORF viewer
    ' Returns -1 if no gels selected in the ORF viewer
    
    Dim lngSourceGelIndex As Long
    Dim lngGelIndex As Long
    
    ' Determine the desired Parent Gel index
    If AllGelsForItemCount(lngSourceGelIndex) = True Then
        ' User has "All Gels for Item Count" selected
        ' Use the first gel in the list
        For lngGelIndex = 1 To mGelDisplayListAndOptions.GelCount
            If mGelDisplayListAndOptions.Gels(lngGelIndex).IncludeGel Then
                lngSourceGelIndex = lngGelIndex
            End If
        Next lngGelIndex
    End If
    
    If lngSourceGelIndex < 1 Then lngSourceGelIndex = -1
    GetDesiredParentGelIndex = lngSourceGelIndex
    
End Function

Private Function GetMassTagItemHitCount(lngListIndex As Long) As Long
    
    ' Only use Ion hit counts if cboDataDisplayMode is ddmIonsOnly
    ' Otherwise, use UMC hit counts
    With lvwMassTags.ListItems(lngListIndex)
        If cboDataDisplayMode.ListIndex = ddmIonsOnly Then
            GetMassTagItemHitCount = CLngSafe(.SubItems(lvmIonHits))
        Else
            GetMassTagItemHitCount = CLngSafe(.SubItems(lvmUMCHits))
        End If
    End With
    
End Function

Private Sub GetNETSlopeAndIntercept(ByVal eNETAdjustmentType As natNETTypeConstants, ByVal lngGelIndex As Long, ByRef dblNETSlope As Double, ByRef dblNETIntercept As Double)
    Dim lngFirstScan As Long, lngLastScan As Long, lngScanRange As Long
    Dim eNETAdjustmentTypeToUse As natNETTypeConstants
    
    If lngGelIndex < 0 Or lngGelIndex > UBound(GelBody()) Then
        dblNETSlope = 1
        dblNETIntercept = 0
        Exit Sub
    End If
    
    If Not GelAnalysis(lngGelIndex) Is Nothing Then
        eNETAdjustmentTypeToUse = eNETAdjustmentType
    Else
        ' GelAnalysis() has not yet been initialized
        eNETAdjustmentTypeToUse = natGeneric
    End If
    
    Select Case eNETAdjustmentTypeToUse
    Case natTICNET
        dblNETSlope = GelAnalysis(lngGelIndex).NET_Slope
        dblNETIntercept = GelAnalysis(lngGelIndex).NET_Intercept
    Case natGANET
        dblNETSlope = GelAnalysis(lngGelIndex).GANET_Slope
        dblNETIntercept = GelAnalysis(lngGelIndex).GANET_Intercept
    Case Else
        ' Includes natGeneric
        GetScanRange lngGelIndex, lngFirstScan, lngLastScan, lngScanRange
        If lngLastScan > lngFirstScan Then
            dblNETSlope = 1 / (lngLastScan - lngFirstScan)
            dblNETIntercept = -lngFirstScan / (lngLastScan - lngFirstScan)
        Else
            dblNETSlope = 1
            dblNETIntercept = 0
        End If
    End Select
    
End Sub

Private Function GetSelectedMassTagDataIndex() As Long
    ' Determines the index of the selected item
    ' Returns -1 if no item selected
    
    Dim lstListItem As MSComctlLib.ListItem
    Set lstListItem = lvwMassTags.SelectedItem

    If lstListItem Is Nothing Then
        GetSelectedMassTagDataIndex = -1
    Else
        GetSelectedMassTagDataIndex = val(lstListItem.Text)
    End If
End Function

Private Function GetSelectedORFDataIndex() As Long
    ' Determines the index of the selected item
    ' Returns -1 if no item selected
    
    Dim lstListItem As MSComctlLib.ListItem
    Set lstListItem = lvwORFs.SelectedItem

    If lstListItem Is Nothing Then
        GetSelectedORFDataIndex = -1
    Else
        GetSelectedORFDataIndex = val(lstListItem.Text)
    End If
End Function

Private Sub HandleListViewColumnClick(ByRef lvwThisListView As MSComctlLib.ListView, ByRef ColumnHeader As MSComctlLib.ColumnHeader, ByRef udtColumnSortFormat As udtColumnSortFormatType, ByVal eListViewID As lviORFFormListViewIDConstants)
    
    Dim lngColumnIndex As Long
    
    ' Need to subtract 1 from ColumnHeader.Index to make 0-based
    lngColumnIndex = ColumnHeader.Index - 1
    
    SortListViewWrapper lvwThisListView, lngColumnIndex, udtColumnSortFormat, eListViewID

End Sub

Private Sub HandleListViewItemClick(ByVal Item As MSComctlLib.ListItem, ByVal eListViewID As lviORFFormListViewIDConstants)
    Dim lngORFGroupArrayIndex As Long
    
    If Not lvwORFs.SelectedItem Is Nothing Then
        If Item <> lvwORFs.SelectedItem Then
            Exit Sub
        End If
    End If
    
    Select Case eListViewID
    Case lviORFs
        lngORFGroupArrayIndex = CLng(Item.Text)
        PopulateMassTagsListView lngORFGroupArrayIndex
                
        ' Make sure the scroll bar is at the top
        HandleVerticalScroll True
        
        ' Add this ORF to the navigation history
        ORFHistoryAdd lngORFGroupArrayIndex
        
        ' Update the recent index history (used with the Find Text in ListView feature)
        ListViewUpdateRecentIndexHistory lvwORFs, Item, CInt(eListViewID)
    Case Else
        
    End Select
End Sub

Private Sub HandleVerticalScroll(Optional blnScrollToTop As Boolean = False)
    Dim lngDesiredTopRowIndex As Long
    Static blnScrolling As Boolean
    
    If blnScrolling Then Exit Sub
    blnScrolling = True
    
    If blnScrollToTop Then VScroll.value = 0
    
    lngDesiredTopRowIndex = VScroll.value
    
    Debug.Assert lngDesiredTopRowIndex < mORFPicsRowCount
    
    ' Simulate scrolling of the pictures by moving fraORFPicsContainer up or down in fraORFPicsClippingRegion
    
    fraORFPicsContainer.Top = lngDesiredTopRowIndex * (mPicHeight + mPicSpacing) * -1
    
    blnScrolling = False
    
End Sub

Private Sub InitializeListViews()
    
    Const COL_WIDTH_NARROW = 750
    Const COL_WIDTH_NORMAL = 1000
    Const COL_WIDTH_WIDE = 4000
    
    Dim lngIndex As Long
    
    ListViewSetFeatures lvwORFs, False
    ListViewSetFeatures lvwMassTags, False
    ListViewSetFeatures lvwColorKey, False

    With lvwORFs
        .ColumnHeaders.add , , "#", COL_WIDTH_NARROW
        For lngIndex = 1 To GRID_COL_COUNT_ORFS - 1
            .ColumnHeaders.add , , LookupColumnTitle(lngIndex, lviORFs), COL_WIDTH_NARROW
        Next lngIndex

        ' Adjust some of the column widths
        .ColumnHeaders(lvoReference + 1).width = COL_WIDTH_NORMAL
        .ColumnHeaders(lvoMassTags + 1).width = COL_WIDTH_NORMAL
        .ColumnHeaders(lvoDescription + 1).width = COL_WIDTH_WIDE
        .ColumnHeaders(lvoMass + 1).width = COL_WIDTH_NORMAL
        
        .ColumnHeaders(lvoSortkey + 1).width = SORTKEY_COL_WIDTH
        
        Debug.Assert .ColumnHeaders.Count = GRID_COL_COUNT_ORFS
    End With

    With lvwMassTags
        .ColumnHeaders.add , , "#", COL_WIDTH_NARROW
        For lngIndex = 1 To GRID_COL_COUNT_MASS_TAGS - 1
            .ColumnHeaders.add , , LookupColumnTitle(lngIndex, lviMassTags), COL_WIDTH_NARROW
        Next lngIndex

        ' Adjust some of the column widths
        .ColumnHeaders(lvmID + 1).width = COL_WIDTH_NARROW
        .ColumnHeaders(lvmMass + 1).width = COL_WIDTH_NORMAL
        .ColumnHeaders(lvmTrypticStatus + 1).width = COL_WIDTH_NORMAL * 2
        .ColumnHeaders(lvmSequence + 1).width = COL_WIDTH_WIDE
        
        .ColumnHeaders(lvmSortKey + 1).width = SORTKEY_COL_WIDTH
        
        Debug.Assert .ColumnHeaders.Count = GRID_COL_COUNT_MASS_TAGS
    End With

    With lvwColorKey
        ' No column headers
        .ColumnHeaders.add , , LookupColumnTitle(0, lviColorKey), COL_WIDTH_WIDE
        For lngIndex = 1 To GRID_COL_COUNT_COLOR_KEY - 1
            .ColumnHeaders.add , , LookupColumnTitle(lngIndex, lviColorKey), COL_WIDTH_NORMAL
        Next lngIndex
    
        .HideSelection = True
        .HideColumnHeaders = True
        .MultiSelect = False
        
        Debug.Assert .ColumnHeaders.Count = GRID_COL_COUNT_COLOR_KEY
    End With
    
    ' Define the column sort formats
    ' Default sort format is text (sfText=0)
    With mColumnSortFormats(lviORFs)
        .ColumnCount = GRID_COL_COUNT_ORFS
        .SortColumnIndexSaved = -1
        .SortKeyColumnIndex = lvoSortkey
        ReDim .ColumnSortOrder(.ColumnCount)
        
        ' Set all to Numeric, except lvoReference and lvoDescription
        For lngIndex = 0 To lvoSortkey - 1
            .ColumnSortOrder(lngIndex) = sfNumeric
        Next lngIndex
        .ColumnSortOrder(lvoReference) = sfText
        .ColumnSortOrder(lvoDescription) = sfText
    End With
    
    ' Define the column sort formats
    ' Default sort format is text (sfText=0)
    With mColumnSortFormats(lviMassTags)
        .ColumnCount = GRID_COL_COUNT_MASS_TAGS
        .SortColumnIndexSaved = -1
        .SortKeyColumnIndex = lvmSortKey
        ReDim .ColumnSortOrder(.ColumnCount)
        
        ' Set all to Numeric, except lvmSequence and lvmTrypticStatus
        For lngIndex = 0 To lvmSortKey - 1
            .ColumnSortOrder(lngIndex) = sfNumeric
        Next lngIndex
        .ColumnSortOrder(lvmSequence) = sfText
        .ColumnSortOrder(lvmTrypticStatus) = sfText
    End With
    
    PopulateColorKeyListView
End Sub

Private Sub InitializeLocalVariables()
    
    ' Controls
    cmdVerticalDivider.ZOrder
    
    mDividerXLoc = lvwORFs.Left + lvwORFs.width + 60
    mDividerYLoc = Me.ScaleHeight / 2
    
    ' Arrays
    ReDim mORFViewerGroupList.Orfs(0)
    mORFViewerGroupList.ORFCount = 0
    mORFHistoryMenuItemsLoadedCount = 1

    ReDim mGelDisplayListAndOptions.Gels(0)
    mGelDisplayListAndOptions.GelCount = 0
    
    mMassTagListViewORFIndex = -1
    
    ORFHistoryClear

    ' mGelDisplayListAndOptions
    With mGelDisplayListAndOptions
        With .DisplayOptions
            .DataDisplayMode = ddmIonsAndUMCs
            
            .UseClassRepresentativeNET = True ' change to true to get asymmetric triangles
            
            .PicturePixelHeight = DEFAULT_ORF_PICTURE_HEIGHT
            .PicturePixelWidth = DEFAULT_ORF_PICTURE_WIDTH
            .PicturePixelSpacing = DEFAULT_ORF_PICTURE_SPACING_PIXELS
            .SwapPlottingAxes = True
            
            .MinSpotSizePixels = DEFAULT_ORF_MIN_SPOT_SIZE_PIXELS
            .MaxSpotSizePixels = DEFAULT_ORF_MAX_SPOT_SIZE_PIXELS
            
            Select Case samtDef.TolType
            Case gltPPM
                .MassTagMassErrorPPM = samtDef.MWTol * 2
            Case gltABS
                .MassTagMassErrorPPM = samtDef.MWTol / 1000 / glPPM * 2
            Case Else
                Debug.Assert False
            End Select
            ValidateValueDbl .MassTagMassErrorPPM, 0, 10000, DEFAULT_ORF_MASS_TAG_MASS_ERROR_PPM
            
            .MassTagNETError = samtDef.NETTol * 2
            ValidateValueDbl .MassTagNETError, 0, 2, DEFAULT_ORF_MASS_TAG_NET_ERROR
            
            ' Set the Display ranges to values 20% larger than the MassTag sizes
            .MassDisplayRangePPM = .MassTagMassErrorPPM * 1.2
            txtMassRange = Trim(.MassDisplayRangePPM)
            
            .NETDisplayRange = .MassTagNETError * 1.2
            txtNETRange = Trim(.NETDisplayRange)
            
            ' Default mass tag color and shape
            .MassTagSpotColor = RGB(0, 192, 192)
            .MassTagSpotShape = sEmptyRectangle
            
            .LogarithmicIntensityPlotting = True
            .IntensityScalar = DEFAULT_ORF_LISTVIEW_INTENSITY_SCALAR
            .IonToUMCPlottingIntensityRatio = DEFAULT_ORF_PICTURE_ION_TO_UMC_INTENSITY_SCALING_RATIO
            
            .CleavageRuleID = 1
            
            .ShowPosition = True
            .ShowGridLines = True
            .ShowTickMarkLabels = False
            
            .LoadPMTs = False
            .ShowNonTrypticMassTagsWithoutIonHits = False
            .HideEmptyMassTagPictures = True
            .IncludeUnobservedTrypticMassTags = True
            
            .OnlyUseTop50PctForAveraging = True
        End With
    End With
    
    InitializeGelDisplayOptions mGelDisplayListAndOptions, 0, True
    
    ' Note: The following will also call UpdateAxisLabelFormattingStrings()
    UpdatePictureSizeAndSpacing

    ' Menu captions
    mnuORFHistoryMovePrevious.Caption = "Previous ORF in History" & vbTab & "Alt+Left Arrow"
    mnuORFHistoryMoveNext.Caption = "Next ORF in History" & vbTab & "Alt+Right Arrow"
End Sub

Private Sub ListViewKeyHandler(lvwThisListView As MSComctlLib.ListView, eListViewID As lviORFFormListViewIDConstants, KeyCode As Integer, Shift As Integer, Optional lngMaxColumnIndex As Long = -1)
    
    If Shift = 2 Then
        If KeyCode = vbKeyA Then
            ListViewSelectAllItems lvwThisListView
        ElseIf KeyCode = vbKeyC Then
            If lngMaxColumnIndex < 0 Then
                lngMaxColumnIndex = lvwThisListView.ColumnHeaders.Count - 1
            End If
            CopySelectedItems lvwThisListView, lngMaxColumnIndex, eListViewID
        End If
    End If

End Sub

' This sub is no longer needed, since GANET computation can now be done using the GANET Class (objGANET)
'''Private Sub LoadGANETValues()
'''    ' Load GANET values from a file into memory
'''    ' Sort by sequence
'''    ' For each of the theoretical mass tags in memory, search for the sequence in the GANET array
'''    ' If found, update the NET value
'''
'''    Const SEQUENCE_DIM_CHUNK = 1000
'''    Dim strFilter As String
'''    Dim strOpenFileName As String
'''    Dim InFileNum As Integer
'''    Dim lngSequenceIndex As Long
'''    Dim lngCharLoc As Long
'''
'''    Dim lngSequenceCount As Long, lngSequenceDimCount As Long
'''    Dim strSequences() As String
'''    Dim sngGANET() As Single
'''
'''    Dim strMassTagSequence As String
'''    Dim lngGelIndex As Long
'''    Dim lngORFIndex As Long, lngMassTagIndex As Long
'''    Dim lngMatchIndex As Long
'''    Dim lngHitCount As Long
'''    Dim strHitCountSummary As String
'''
'''    strFilter = "Text Files (*.txt)|*.txt|All files(*.*)|*.*"
'''    strOpenFileName = SelectFile(Me.hwnd, "Open File", "", False, "", strFilter, 1)
'''
'''    If Len(strOpenFileName) > 0 Then
'''
'''        frmProgress.InitializeForm "Loading GANET values", 0, 3 + UBound(GelORFMassTags()), False, True, True, MDIForm1
'''        frmProgress.InitializeSubtask "Reading file", 0, 1
'''
'''        lngSequenceCount = 0
'''        lngSequenceDimCount = SEQUENCE_DIM_CHUNK
'''        ReDim strSequences(lngSequenceDimCount)
'''
'''        InFileNum = FreeFile()
'''        Open strOpenFileName For Input As #InFileNum
'''
'''        Do While Not EOF(InFileNum)
'''            Line Input #InFileNum, strSequences(lngSequenceCount)
'''
'''            lngSequenceCount = lngSequenceCount + 1
'''            If lngSequenceCount >= lngSequenceDimCount Then
'''                lngSequenceDimCount = lngSequenceDimCount + SEQUENCE_DIM_CHUNK
'''                ReDim Preserve strSequences(lngSequenceDimCount)
'''            End If
'''        Loop
'''
'''        Close #InFileNum
'''
'''        frmProgress.UpdateProgressBar 1
'''        frmProgress.InitializeSubtask "Sorting", 0, 1
'''
'''        QuickSortString strSequences(), 0, lngSequenceCount - 1
'''
'''        ' Extract out the GANET values from strSequences
'''        ReDim sngGANET(lngSequenceCount)
'''
'''        frmProgress.UpdateProgressBar 2
'''        frmProgress.InitializeSubtask "Splitting GANET values from sequences", 0, 1
'''
'''        For lngSequenceIndex = 0 To lngSequenceCount - 1
'''            lngCharLoc = InStr(strSequences(lngSequenceIndex), vbTab)
'''            If lngCharLoc > 0 Then
'''                sngGANET(lngSequenceIndex) = Mid(strSequences(lngSequenceIndex), lngCharLoc + 1)
'''                strSequences(lngSequenceIndex) = Left(strSequences(lngSequenceIndex), lngCharLoc - 1)
'''            End If
'''        Next lngSequenceIndex
'''
'''        For lngGelIndex = 1 To UBound(GelORFMassTags())
'''            lngHitCount = 0
'''            With GelORFMassTags(lngGelIndex)
'''                frmProgress.UpdateProgressBar 2 + lngGelIndex
'''                frmProgress.InitializeSubtask "Finding matching theoretical mass tags", 0, .ORFCount
'''
'''                For lngORFIndex = 0 To .ORFCount - 1
'''                    With .Orfs(lngORFIndex)
'''                        For lngMassTagIndex = 0 To .MassTagCount - 1
'''                            With .MassTags(lngMassTagIndex)
'''                                If .IsTheoretical Then
'''                                    strMassTagSequence = GetSequencePortion(GelORFData(lngGelIndex).Orfs(lngORFIndex).Sequence, .Location.ResidueStart, .Location.ResidueEnd, False)
'''                                    lngMatchIndex = BinarySearchStr(strSequences(), strMassTagSequence, 0, lngSequenceCount - 1)
'''
'''                                    If lngMatchIndex >= 0 Then
'''                                        .GANET = sngGANET(lngMatchIndex)
'''                                        lngHitCount = lngHitCount + 1
'''                                    End If
'''
'''                                End If
'''                            End With
'''                        Next lngMassTagIndex
'''                    End With
'''
'''                    If lngORFIndex Mod 50 = 0 Then
'''                        frmProgress.UpdateSubtaskProgressBar lngORFIndex
'''                    End If
'''                Next lngORFIndex
'''
'''            End With
'''
'''            GelStatus(lngGelIndex).Dirty = True
'''
'''            strHitCountSummary = strHitCountSummary & vbCrLf & "GelIndex " & lngGelIndex & ": Updated " & lngHitCount & " mass tags with new GANET values."
'''
'''        Next lngGelIndex
'''
'''        MsgBox "Done.  " & strHitCountSummary
'''
'''        frmProgress.HideForm
'''
'''        PopulateORFGroupList True, True
'''    End If
'''
'''End Sub

Private Function LookupColumnTitle(lngColumnIndex As Long, eListViewID As lviORFFormListViewIDConstants) As String
    Select Case eListViewID
    Case lviORFs
        Select Case lngColumnIndex
        Case lvoID: LookupColumnTitle = "#"
        Case lvoReference: LookupColumnTitle = "Reference"
        Case lvoMassTagHitsViaIons: LookupColumnTitle = "Mass Tag Hits via Peaks"
        Case lvoMassTagHitsViaUMCs: LookupColumnTitle = "Mass Tag Hits via UMCs"
        Case lvoORFIntensityViaIons: LookupColumnTitle = "Orf Intensity via Peaks"
        Case lvoORFIntensityViaUMCs: LookupColumnTitle = "Orf Intensity via UMCs"
        Case lvoMassTags: LookupColumnTitle = "Total Mass Tags"
        Case lvoTryptics: LookupColumnTitle = "Tryptics"
        Case lvoDescription: LookupColumnTitle = "Description"
        Case lvoMass: LookupColumnTitle = "Mass"
        Case lvoSortkey: LookupColumnTitle = "SortKey"
        Case Else
            Debug.Assert False
            LookupColumnTitle = "??"
        End Select
    Case lviMassTags
        Select Case lngColumnIndex
        Case lvmID: LookupColumnTitle = "#"
        Case lvmMTID: LookupColumnTitle = "MTID"
        Case lvmIonHits: LookupColumnTitle = "Peak Hits"
        Case lvmIonHitSum: LookupColumnTitle = "Peak Hit Intensity Sum"
        Case lvmUMCHits: LookupColumnTitle = "UMC Hits"
        Case lvmUMCHitSum: LookupColumnTitle = "UMC Hit Intensity Sum"
        Case lvmMass: LookupColumnTitle = "Mass"
        Case lvmNET: LookupColumnTitle = "NET"
        Case lvmResidueCount: LookupColumnTitle = "Res Cnt"
        Case lvmSequence: LookupColumnTitle = "Sequence"
        Case lvmTrypticStatus: LookupColumnTitle = "Rule Match?"
        Case lvmSortKey: LookupColumnTitle = "SortKey"
        Case Else
            Debug.Assert False
            LookupColumnTitle = "??"
        End Select
    Case lviColorKey
        Select Case lngColumnIndex
        Case lvkName:  LookupColumnTitle = "File Name"
        Case lvkUMCColor: LookupColumnTitle = "UMC Color"
        Case Else
            Debug.Assert False
            LookupColumnTitle = "??"
        End Select
    Case Else
        Debug.Assert False
        LookupColumnTitle = "??"
    End Select
    
End Function

Private Sub ORFHistoryAdd(lngORFGroupArrayIndex As Long)
    ' Add the item after location .CurrentHistoryIndex in ORFHistory
    ' However, do not add the item if the item in .CurrentHistoryIndex matches the item to add
    ' Remove any items after the item (if added)

    Dim blnAddItem As Boolean
    Dim intHistoryIndex As Integer
    
    With mORFHistory
        With .History(.CurrentHistoryIndex)
            If .ORFGroupArrayIndex <> lngORFGroupArrayIndex Or .ORFDescription <> mORFViewerGroupList.Orfs(lngORFGroupArrayIndex).Reference Then
                ' Add item to history, remove items after
                blnAddItem = True
            End If
        End With
        
        If blnAddItem Then
            If .CurrentHistoryIndex = ORF_HISTORY_COUNT_MAX - 1 Then
                ' Max history count reached; need to shift entries in the history down by one index
                For intHistoryIndex = 0 To .CurrentHistoryIndex - 1
                    .History(intHistoryIndex) = .History(intHistoryIndex + 1)
                Next intHistoryIndex
            Else
                If .HistoryCount > 0 Then
                    .CurrentHistoryIndex = .CurrentHistoryIndex + 1
                Else
                    .CurrentHistoryIndex = 0
                End If
            End If
            
            With .History(.CurrentHistoryIndex)
                .ORFGroupArrayIndex = lngORFGroupArrayIndex
                .ORFDescription = mORFViewerGroupList.Orfs(lngORFGroupArrayIndex).Reference
            End With
            
            ' Clear the remaining history entries
            ' This isn't truly necessary, but helps keep the history clean
            For intHistoryIndex = .CurrentHistoryIndex + 1 To ORF_HISTORY_COUNT_MAX - 1
                .History(intHistoryIndex).ORFGroupArrayIndex = 0
                .History(intHistoryIndex).ORFDescription = ""
            Next intHistoryIndex
            
            ' Update .HistoryCount
            .HistoryCount = .CurrentHistoryIndex + 1
        End If
    End With
                
    UpdateORFHistoryMenuList

End Sub

Private Sub ORFHistoryClear()

    With mORFHistory
        .HistoryCount = 0
        Erase .History      ' Note, since array Dim'd at design time, using Erase will clear it, but not release the memory
        .CurrentHistoryIndex = 0
    End With
    
    UpdateORFHistoryMenuList
End Sub

Private Sub ORFHistoryNavigate(blnMoveForward As Boolean)
    ' Move forward or backward in the history
    
    ' Note: No need to check if the value to jump to is valid here, since the
    '       ORFHistoryJump function checks for this
    If blnMoveForward Then
        ORFHistoryJump mORFHistory.CurrentHistoryIndex + 1
    Else
        ORFHistoryJump mORFHistory.CurrentHistoryIndex - 1
    End If
End Sub

Private Sub ORFHistoryJump(intHistoryIndex As Integer)
    ' Jump to the item with intHistoryIndex in the OrfHistory
    
    With mORFHistory
        ' Make sure intHistoryIndex is valid
        If intHistoryIndex < .HistoryCount And intHistoryIndex >= 0 Then
            .CurrentHistoryIndex = intHistoryIndex
            
            ' The following call will cause ORFHistoryAdd() to be called, which
            '  in turn will call UpdateORFHistoryMenuList()
            SelectORFByID .History(.CurrentHistoryIndex).ORFGroupArrayIndex
        Else
            ' Update the ORF History Menu list to hopefully disable the Forward/Previous menu
            ' item that resulted in the invalid call to this function
            UpdateORFHistoryMenuList
        End If
    End With
End Sub

Private Sub PopulateColorKeyListView()
    
    Const MINIMUM_VIEW_OPTIONS_FRAME_HEIGHT = 700
    Dim lstNewItem As MSComctlLib.ListItem
    Dim lngGelIndex As Long
    Dim lngRowHeight As Long
    
    With lvwColorKey
    
        .ListItems.Clear
        For lngGelIndex = 1 To mGelDisplayListAndOptions.GelCount
            If mGelDisplayListAndOptions.Gels(lngGelIndex).IncludeGel Then
                Set lstNewItem = .ListItems.add(, , CompactPathString(mGelDisplayListAndOptions.Gels(lngGelIndex).GelFileName, 75))
                lstNewItem.ForeColor = mGelDisplayListAndOptions.Gels(lngGelIndex).IonSpotColor
                lstNewItem.Bold = True
                
                lstNewItem.SubItems(lvkUMCColor) = "UMC"
                lstNewItem.ListSubItems(lvkUMCColor).ForeColor = mGelDisplayListAndOptions.Gels(lngGelIndex).UMCSpotColor
                lstNewItem.ListSubItems(lvkUMCColor).Bold = True
                
                lngRowHeight = lstNewItem.Height
            End If
        Next lngGelIndex
        
        
        .Refresh
    End With

    ' Set the size of fraViewOptions based on the number of included gels
    If lvwColorKey.ListItems.Count <= 2 Then
        fraViewOptions.Height = MINIMUM_VIEW_OPTIONS_FRAME_HEIGHT
    Else
        fraViewOptions.Height = MINIMUM_VIEW_OPTIONS_FRAME_HEIGHT + (lngRowHeight + 20) * (lvwColorKey.ListItems.Count - 2)
    End If

    PositionControls
    
End Sub

Private Sub PopulateComboBoxes()
    
    With cboDataDisplayMode
        .Clear
        .AddItem "Display Ions"
        .AddItem "Display UMCs"
        .AddItem "Display Ions and UMCs"
        .ListIndex = ddmIonsAndUMCs
    End With
End Sub

Public Sub PopulateMassTagsListView(lngORFGroupArrayIndex As Long)
    
    Dim lngGelIndex As Long
    Dim lngORFIndex As Long
    Dim lngMassTagIndex As Long
    Dim lngSourceGelIndex As Long
    Dim blnUseAllGelsForItemCount As Boolean
    Dim blnShowMassTag As Boolean
    Dim lngIonMatchCount As Long, lngIonMatchSum As Long
    Dim lngUMCMatchCount As Long, lngUMCMatchSum As Long
    
    Dim dblMassTagIntensitySum As Double
    Dim strSequence As String, strSequenceWithFormatting As String
    Dim blnMatchesRule As Boolean
    Dim intRuleMatchCount As Integer, strTrypticStatus As String
    Dim eDataDisplayMode As ddmDataDisplayModeConstants
    
    Dim lstNewItem As MSComctlLib.ListItem
    
    If lngORFGroupArrayIndex < 0 Or lngORFGroupArrayIndex >= mORFViewerGroupList.ORFCount Then
        ' This shouldn't happen
        Debug.Assert False
        mMassTagListViewORFIndex = -1
        Exit Sub
    End If
    
    ' Determine the Source Gel to be used for item counts
    blnUseAllGelsForItemCount = AllGelsForItemCount(lngSourceGelIndex)
    eDataDisplayMode = mGelDisplayListAndOptions.DisplayOptions.DataDisplayMode
    
    lvwMassTags.Visible = False
    Me.MousePointer = vbHourglass
    DoEvents
    
    lvwMassTags.ListItems.Clear
    mMassTagListViewMaxMatchCount = 0
    mMassTagListViewORFIndex = lngORFGroupArrayIndex
    
    If mORFViewerGroupList.Orfs(lngORFGroupArrayIndex).ItemCount > 0 Then
        
        lngGelIndex = mORFViewerGroupList.Orfs(lngORFGroupArrayIndex).Items(0).GelIndex
        lngORFIndex = mORFViewerGroupList.Orfs(lngORFGroupArrayIndex).Items(0).ORFIndex
        
        With GelORFMassTags(lngGelIndex).Orfs(lngORFIndex)
        
            For lngMassTagIndex = 0 To .MassTagCount - 1
            
                With .MassTags(lngMassTagIndex)
                    
                    ' If .IncludeUnobservedTrypticMassTags = False, then do not show Theoretical mass tags
                    If mGelDisplayListAndOptions.DisplayOptions.IncludeUnobservedTrypticMassTags Or Not .IsTheoretical Then
                        ' Determine if the ion is tryptic and how many ion hits it has
                        ' First need to get the sequence, formatted with prefix and suffix residues
                        strSequence = GetSequencePortion(GelORFData(lngGelIndex).Orfs(lngORFIndex).Sequence, .Location.ResidueStart, .Location.ResidueEnd, False)
                        strSequenceWithFormatting = GetSequencePortion(GelORFData(lngGelIndex).Orfs(lngORFIndex).Sequence, .Location.ResidueStart, .Location.ResidueEnd, True)
                        
                        ' See if the sequence is tryptic (or Stryptic, or whatever our current rule residues are)
                        blnMatchesRule = CheckSequenceAgainstCleavageRuleWrapper(strSequenceWithFormatting, mGelDisplayListAndOptions.DisplayOptions.CleavageRuleID, intRuleMatchCount)
                        
                        If blnMatchesRule Then
                            strTrypticStatus = RULE_MATCH_YES
                        Else
                            strTrypticStatus = RULE_MATCH_NO
                        End If
                        
                        If intRuleMatchCount = 2 Then
                            strTrypticStatus = strTrypticStatus & " (Full)"
                        ElseIf intRuleMatchCount = 1 Then
                            strTrypticStatus = strTrypticStatus & " (Partial)"
                        End If
                        
                        lngIonMatchCount = GetDesiredMatchCount(lngORFGroupArrayIndex, False, dblMassTagIntensitySum, blnUseAllGelsForItemCount, lngSourceGelIndex, True, .MassTagRefID)
                        If lngIonMatchCount > mMassTagListViewMaxMatchCount Then mMassTagListViewMaxMatchCount = lngIonMatchCount
                        lngIonMatchSum = Round(dblMassTagIntensitySum / mGelDisplayListAndOptions.DisplayOptions.IntensityScalar, 0)
                        
                        lngUMCMatchCount = GetDesiredMatchCount(lngORFGroupArrayIndex, True, dblMassTagIntensitySum, blnUseAllGelsForItemCount, lngSourceGelIndex, True, .MassTagRefID)
                        If lngUMCMatchCount > mMassTagListViewMaxMatchCount Then mMassTagListViewMaxMatchCount = lngUMCMatchCount
                        lngUMCMatchSum = Round(dblMassTagIntensitySum / mGelDisplayListAndOptions.DisplayOptions.IntensityScalar, 0)
                        
                        If mGelDisplayListAndOptions.DisplayOptions.ShowNonTrypticMassTagsWithoutIonHits Then
                            ' Show all mass tags
                            blnShowMassTag = True
                        Else
                            ' Only show full tryptic mass tags, or those mass tags with ion or UMC matches
                            If blnMatchesRule Or (eDataDisplayMode = ddmIonsOnly And lngIonMatchCount > 0) Or _
                                                 (eDataDisplayMode <> ddmIonsOnly And lngUMCMatchCount > 0) Then
                                blnShowMassTag = True
                            Else
                                blnShowMassTag = False
                            End If
                        End If
                        
                        If blnShowMassTag Then
                            Set lstNewItem = lvwMassTags.ListItems.add(, , lngMassTagIndex)    ' Internal ID # in GelORFMassTags().ORFS().MassTags()
                            
                            lstNewItem.SubItems(lvmMTID) = .MassTagRefID
                                                              
                            lstNewItem.SubItems(lvmIonHits) = lngIonMatchCount
                            lstNewItem.SubItems(lvmIonHitSum) = lngIonMatchSum
                            
                            lstNewItem.SubItems(lvmUMCHits) = lngUMCMatchCount
                            lstNewItem.SubItems(lvmUMCHitSum) = lngUMCMatchSum
                            
                            lstNewItem.SubItems(lvmMass) = Round(.Mass, 4)
                            
                            lstNewItem.SubItems(lvmNET) = Round(.GANET, 2)
                            
                            
                            lstNewItem.SubItems(lvmResidueCount) = Len(strSequence)
                            
                            lstNewItem.SubItems(lvmSequence) = strSequenceWithFormatting
                            
                            lstNewItem.SubItems(lvmTrypticStatus) = strTrypticStatus
                        End If
                    End If
                End With
            Next lngMassTagIndex
        End With
        
        ' Sort the listview by TrypticStatus
        ' This will automatically call UpdateORFPics
        ' It is important that we set mMassTagListViewORFIndex before calling SortListViewWrapper (this is done above)
        ' In addition, must make sure .SortColumnIndexSaved is not equal to lvmTrypticStatus
        mColumnSortFormats(lviMassTags).SortColumnIndexSaved = -1
        SortListViewWrapper lvwMassTags, lvmTrypticStatus, mColumnSortFormats(lviMassTags), lviMassTags
        
        If lvwMassTags.ListItems.Count = 0 Then
            UpdateStatusORFDisplayCount -1
        Else
            ' Highlight the first entry in the list
            ListViewHighlightItem lvwMassTags, 1
        End If
        
    Else
        UpdateORFPics
    End If

    lvwMassTags.Visible = True
    
    Me.MousePointer = vbDefault
End Sub

Private Sub PopulateMassTagsListViewCustomSortColumn()
    Dim lngIndex As Long
    Dim lngHitCount As Long
    Dim strFormatString As String
    Dim strSortKey As String, strPrimaryKey As String
    Dim lngSubKeyValue As Long
    Dim lngMassTagIndex As Long, lngStartResidueLoc As Long
    
    ' User clicked on the TrypticStatus column in lvwMassTags
    ' Need to construct a custom string for the SortKey column
    ' Sort in the following order:
    ' Tryptic Mass tags with >0 ion (or UMC) matches        strPrimaryKey = "0"
    '   Sub-sort by decreasing match count
    ' Non-tryptic mass tags with >0 ion (or UMC) matches    strPrimaryKey = "1"
    '   Sub-sort by decreasing match count
    ' Tryptic mass tags without matches                     strPrimaryKey = "2"
    '   Sub-sort by location in the ORF
    ' Non-tryptic mass tags without matches                 strPrimaryKey = "3"
    '   Sub-sort by location in the ORF
    
    ' In order to sort by decreasing match count, need to know the highest match count
    ' I have already determined this in PopulateMassTagsListView() and stored it in mMassTagListViewMaxMatchCount
    
    ' The SortKey that is constructed will be of the form 0.08 where the first digit
    ' is 0, 1, 2, or 3 (see above), and the remaining digits are rank by hit count (highest hit count is 00, second highest is 01, etc)
    
    ' Set the format string to "000000" since a protein could contain thousands of residues
    strFormatString = String(6, "0")
    
    lvwMassTags.Sorted = False
    
    For lngIndex = 1 To lvwMassTags.ListItems.Count
        lngMassTagIndex = CLngSafe(lvwMassTags.ListItems(lngIndex).Text)
        
        ' Attempt to determine the start residue of this mass tag
        ' Although this shouldn't give an error, I'may want to enable On Error Resume Next handling, just in case
        With mORFViewerGroupList.Orfs(mMassTagListViewORFIndex).Items(0)
            lngStartResidueLoc = GelORFMassTags(.GelIndex).Orfs(.ORFIndex).MassTags(lngMassTagIndex).Location.ResidueStart
        End With
        
        With lvwMassTags.ListItems(lngIndex)
            lngHitCount = GetMassTagItemHitCount(lngIndex)
            
            ' Construct a string for the SortKey column
            If Left(.SubItems(lvmTrypticStatus), Len(RULE_MATCH_YES)) = RULE_MATCH_YES Then
                ' Matched cleavage rule
                If lngHitCount > 0 Then
                    strPrimaryKey = "0"
                Else
                    strPrimaryKey = "2"
                End If
                lngSubKeyValue = mMassTagListViewMaxMatchCount - lngHitCount
            Else
                ' Did not match cleavage rule
                If lngHitCount > 0 Then
                    strPrimaryKey = "1"
                Else
                    strPrimaryKey = "3"
                End If
                lngSubKeyValue = 0
            End If
            
            strSortKey = strPrimaryKey & "." & Format(lngSubKeyValue, strFormatString) & "." & Format(lngStartResidueLoc, strFormatString)
            
            .SubItems(lvmSortKey) = strSortKey
        End With
    Next lngIndex

    lvwMassTags.Sorted = False

End Sub

Public Sub PopulateORFGroupList(Optional blnShowORFViewerOptionsFormIfNoGelsSelected As Boolean = True, Optional blnForceIonMatchRecount As Boolean = False)
    
    Dim lngGelIndex As Long
    Dim lngORFIndex As Long
    Dim lngORFGroupArrayIndex As Long, lngORFGroupArrayIndexPrevious As Long
    Dim intIncludedGelCount As Integer
    Dim lngORFDimCount As Long
    
    Dim lngORFRefID As Long
    Dim strReference As String
    Dim blnMatchFound As Boolean
    
    mORFViewerGroupList.ORFCount = 0
    lngORFDimCount = ORF_DIM_CHUNK
    ReDim mORFViewerGroupList.Orfs(lngORFDimCount)
    
    VerifyORFsLoaded False, False, blnForceIonMatchRecount
    
    ' Reset the flag that warns the user when ion count errors exist
    mUserNotifiedOfIonMatchCountError = False
    
    If mGelDisplayListAndOptions.GelCount > 0 Then
        
        ' Populate mORFViewerGroupList() using the ORFs from the gels
        '  in mGelDisplayListAndOptions with .IncludeGel = True
        
        ' Determine the number of gels with .IncludeGel = True
        intIncludedGelCount = GetNumberOfIncludedGels(mGelDisplayListAndOptions)
        If intIncludedGelCount > 0 Then
            frmProgress.InitializeForm "Populating ORF Groups list", 0, CLng(intIncludedGelCount), True, (intIncludedGelCount <> 1), True, MDIForm1
        End If
        
        lngORFGroupArrayIndex = 0
        lngORFGroupArrayIndexPrevious = 0
        intIncludedGelCount = 0
        For lngGelIndex = 1 To mGelDisplayListAndOptions.GelCount
            If mGelDisplayListAndOptions.Gels(lngGelIndex).IncludeGel Then
                frmProgress.InitializeSubtask "", 0, GelORFData(lngGelIndex).ORFCount
                
                intIncludedGelCount = intIncludedGelCount + 1
                
                With mORFViewerGroupList
                    For lngORFIndex = 0 To GelORFData(lngGelIndex).ORFCount - 1
                        If lngORFIndex Mod 100 = 0 Then
                            frmProgress.UpdateSubtaskProgressBar lngORFIndex
                        End If
                        
                        lngORFRefID = GelORFData(lngGelIndex).Orfs(lngORFIndex).RefID
                        strReference = GelORFData(lngGelIndex).Orfs(lngORFIndex).Reference
                        
                        blnMatchFound = False
                        If intIncludedGelCount = 1 Then
                            ' Copy ORFs from first selected gel into mORFViewerGroupList
                            ' No need to look for existing entry in list since it won't exist
                        Else
                            ' Find the record in mORFViewerGroupList.Orfs() with lngORFRefID
                            Do
                                If .Orfs(lngORFGroupArrayIndex).RefID = lngORFRefID Then
                                    ' ID's match; make sure the string matches too
                                    If .Orfs(lngORFGroupArrayIndex).Reference = strReference Then
                                        blnMatchFound = True
                                        lngORFGroupArrayIndexPrevious = lngORFGroupArrayIndex
                                        Exit Do
                                    End If
                                End If
                            
                                lngORFGroupArrayIndex = lngORFGroupArrayIndex + 1
                                If lngORFGroupArrayIndex >= .ORFCount Then
                                    lngORFGroupArrayIndex = 0
                                End If
                                If lngORFGroupArrayIndex = lngORFGroupArrayIndexPrevious Then
                                    ' Search looped without finding a match
                                    Exit Do
                                End If
                            Loop
                        End If
            
                        If Not blnMatchFound Then
                            ' Match not found; add a new entry for the ORF
                            lngORFGroupArrayIndex = .ORFCount
                            lngORFGroupArrayIndexPrevious = lngORFGroupArrayIndex
                            With .Orfs(lngORFGroupArrayIndex)
                                .Reference = strReference
                                .RefID = lngORFRefID
                                .ItemCount = 0
                                ReDim .Items(0)
                            End With
                            blnMatchFound = True
                            .ORFCount = .ORFCount + 1
                            If .ORFCount >= lngORFDimCount Then
                                lngORFDimCount = lngORFDimCount + ORF_DIM_CHUNK
                                ReDim Preserve mORFViewerGroupList.Orfs(lngORFDimCount)
                            End If
                        End If
                        
                        Debug.Assert blnMatchFound
                        
                        With .Orfs(lngORFGroupArrayIndex)
                            .ItemCount = .ItemCount + 1
                            ReDim Preserve .Items(.ItemCount)
                            .Items(.ItemCount - 1).GelIndex = lngGelIndex
                            .Items(.ItemCount - 1).ORFIndex = lngORFIndex
                        End With
                        
                    Next lngORFIndex
                End With
            
                frmProgress.UpdateProgressBar intIncludedGelCount
            
            End If
        Next lngGelIndex
        
        If intIncludedGelCount = 0 Then
            If blnShowORFViewerOptionsFormIfNoGelsSelected Then
                ShowORFViewerOptions
            End If
        Else
            ' Need to update the mass indexing arrays
            UpdateMassIndexingArrays
        End If
    End If
    
    PopulateORFListView

    frmProgress.HideForm
    
End Sub

Public Sub PopulateORFListView()
    ' Fill lvwORFs with ORFs in mORFViewerGroupList
    
    Dim lngPointerArray() As Long
    Dim lngPointerArrayCount As Long
    Dim lngPointerIndex As Long
    Dim lngORFGroupArrayIndex As Long
    Dim lngZerothItemGelIndex As Long
    Dim lngORFArrayIndexInGel As Long
    
    Dim lngSourceGelIndex As Long
    Dim blnUseAllGelsForItemCount As Boolean
    Dim dblORFIntensity As Double
    
    Dim lstNewItem As MSComctlLib.ListItem
    
    Dim lngORFDataIndexSelected As Long
    
    lvwORFs.ListItems.Clear
    If mORFViewerGroupList.ORFCount < 1 Then
        UpdateStatusORFDisplayCount -1
        lvwMassTags.ListItems.Clear
        UpdateORFPics
        Exit Sub
    End If
    
    ' Hide the ListView and set WM_SETREDRAW to False
    ListViewShowHideForUpdating lvwORFs, True
    lvwORFs.Visible = False
    
    Me.MousePointer = vbHourglass
    DoEvents
    
    ' Use a PointerArray to allow filtering out of which ORFs to display
    
    ' Initialize the PointerArray
    ReDim lngPointerArray(mORFViewerGroupList.ORFCount)
    lngPointerArrayCount = 0
    
    ' For now, display all of the ORFs
    ' Thus, fill lngPointerArray() accordingly
    lngPointerArrayCount = mORFViewerGroupList.ORFCount
    For lngPointerIndex = 0 To lngPointerArrayCount - 1
        lngPointerArray(lngPointerIndex) = lngPointerIndex
    Next lngPointerIndex
        
    ' Determine the Source Gel to be used for item counts
    blnUseAllGelsForItemCount = AllGelsForItemCount(lngSourceGelIndex)
    
    If mGelDisplayListAndOptions.DisplayOptions.IntensityScalar < 1 Then
        mGelDisplayListAndOptions.DisplayOptions.IntensityScalar = 1
    End If
    
    ' Display the data
    UpdateStatus "Populating list:"
    mKeyPressAbortORFListPopulate = 1
    For lngPointerIndex = 0 To lngPointerArrayCount - 1
        lngORFGroupArrayIndex = lngPointerArray(lngPointerIndex)        ' Not necessarily the same as .Orfs().ORFID
                
        If mORFViewerGroupList.Orfs(lngORFGroupArrayIndex).ItemCount > 0 Then
            lngZerothItemGelIndex = mORFViewerGroupList.Orfs(lngORFGroupArrayIndex).Items(0).GelIndex
            lngORFArrayIndexInGel = mORFViewerGroupList.Orfs(lngORFGroupArrayIndex).Items(0).ORFIndex
        
            If lngZerothItemGelIndex >= 1 Then
                With GelORFData(lngZerothItemGelIndex).Orfs(lngORFArrayIndexInGel)
                    Set lstNewItem = lvwORFs.ListItems.add(, , lngORFGroupArrayIndex)    ' Internal ID # in mORFViewerGroupList
                    
                    lstNewItem.SubItems(lvoReference) = .Reference
                    lstNewItem.SubItems(lvoMassTagHitsViaIons) = GetDesiredMatchCount(lngORFGroupArrayIndex, False, dblORFIntensity, blnUseAllGelsForItemCount, lngSourceGelIndex, False)
                    lstNewItem.SubItems(lvoORFIntensityViaIons) = Round(dblORFIntensity / mGelDisplayListAndOptions.DisplayOptions.IntensityScalar, 0)

                    lstNewItem.SubItems(lvoMassTagHitsViaUMCs) = GetDesiredMatchCount(lngORFGroupArrayIndex, True, dblORFIntensity, blnUseAllGelsForItemCount, lngSourceGelIndex, False)
                    lstNewItem.SubItems(lvoORFIntensityViaUMCs) = Round(dblORFIntensity / mGelDisplayListAndOptions.DisplayOptions.IntensityScalar, 0)

                    lstNewItem.SubItems(lvoMassTags) = GelORFMassTags(lngZerothItemGelIndex).Orfs(lngORFArrayIndexInGel).MassTagCount
                    lstNewItem.SubItems(lvoTryptics) = .TrypticFragmentCount
                    lstNewItem.SubItems(lvoDescription) = .Description
                    lstNewItem.SubItems(lvoMass) = Round(.MassMonoisotopic, 1)
                End With
            Else
                ' This is unexpected
                ' lngZerothItemGelIndex should be 1 or greater
                Debug.Assert False
            End If
            
            ' Update lblStatus since updating the ListView takes awhile
            If lngPointerIndex Mod 50 = 0 Then
                UpdateStatus "Populating list: " & lngPointerIndex & " / " & lngPointerArrayCount
                If mKeyPressAbortORFListPopulate > 1 Then Exit For
            End If
        End If
    Next lngPointerIndex
    mKeyPressAbortORFListPopulate = 0
    
    ' Resort lvwORFs
    ' Need to change udtColumnSortFormat.SortColumnIndexSaved to not = lvoReference
    mColumnSortFormats(lviORFs).SortColumnIndexSaved = -1
    ListViewSort lvwORFs, lvoReference, mColumnSortFormats(lviORFs)

    ListViewShowHideForUpdating lvwORFs, False
    On Error Resume Next
    lvwORFs.SetFocus
    On Error GoTo 0
    
    ' lngORFDataIndexSelected contains the OrfGroupArray index of the selected item in lvwORFS
    lngORFDataIndexSelected = GetSelectedORFDataIndex()
    If lngORFDataIndexSelected >= 0 Then
        PopulateMassTagsListView lngORFDataIndexSelected
        ORFHistoryAdd lngORFDataIndexSelected
        
        UpdateStatusORFDisplayCount -1
    Else
        UpdateORFPics
    End If
    
    Me.MousePointer = vbDefault

End Sub

Private Sub PopulateSourceGelComboBox()
    Dim strSourceGelComboSaved As String
    Dim lngGelIndex As Long
    
    mPopulatingSourceGelCombo = True
    strSourceGelComboSaved = cboItemCountSourceGel.List(cboItemCountSourceGel.ListIndex)
    
    ' Populate cboItemCountSourceGel
    cboItemCountSourceGel.Clear
    cboItemCountSourceGel.AddItem "Counts use All Gels"
    If strSourceGelComboSaved = "Counts use All Gels" Then
        cboItemCountSourceGel.ListIndex = 0
    End If
    
    With mGelDisplayListAndOptions
        For lngGelIndex = 1 To .GelCount
            If .Gels(lngGelIndex).IncludeGel Then
                cboItemCountSourceGel.AddItem Trim(lngGelIndex) & ": " & .Gels(lngGelIndex).GelFileName
                
                If cboItemCountSourceGel.List(cboItemCountSourceGel.ListCount - 1) = strSourceGelComboSaved Then
                    cboItemCountSourceGel.ListIndex = cboItemCountSourceGel.ListCount - 1
                End If
            End If
        Next lngGelIndex
    End With
    
    If cboItemCountSourceGel.ListIndex < 0 Then
        cboItemCountSourceGel.ListIndex = 0
    End If
    
    mPopulatingSourceGelCombo = False

End Sub

Private Sub PositionControls(Optional blnFormResized As Boolean = False, Optional blnForceORFPicRearrange As Boolean = False)
    ' Arranges the controls on the form
    
    Dim lngDesiredValue As Long, lngCompareValue As Long
    Dim lngMaxXSaved As Long, lngMaxYSaved As Long
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If blnFormResized Then
        lngMaxXSaved = mDividerMaxX
        lngMaxYSaved = mDividerMaxY
        
        ' Need to update the divider limits
        UpdateDividerLimits
        
        If lngMaxXSaved = 0 Or lngMaxYSaved = 0 Then
            lngMaxXSaved = mDividerMaxX
            lngMaxYSaved = mDividerMaxY
        End If
        
        ' Update the locations
        mDividerXLoc = mDividerXLoc / (lngMaxXSaved - mDividerMinX) * (mDividerMaxX - mDividerMinX)
        mDividerYLoc = mDividerYLoc / (lngMaxYSaved - mDividerMinY) * (mDividerMaxY - mDividerMinY)
        
        If mDividerXLoc > mDividerMaxX Then mDividerXLoc = mDividerMaxX
        If mDividerYLoc > mDividerMaxY Then mDividerYLoc = mDividerMaxY
    End If
    
    With cmdVerticalDivider
        .Top = 60
        lngDesiredValue = Me.ScaleHeight - .Top - fraViewOptions.Height - 60
        If lngDesiredValue < 0 Then lngDesiredValue = 0
        .Height = lngDesiredValue
        .width = 150
        .Left = mDividerXLoc
    End With
    
    cmdRollUpShrink.Top = cmdVerticalDivider.Top + cmdVerticalDivider.Height + 60
    cmdRollUpShrink.Left = 60
    
    cmdRollUpExpand.Top = cmdRollUpShrink.Top
    cmdRollUpExpand.Left = cmdRollUpShrink.Left
    
    With fraViewOptions
        .Top = cmdRollUpShrink.Top
        .Left = cmdRollUpShrink.Left + cmdRollUpShrink.width + 120
        lngDesiredValue = Me.ScaleWidth - .Left - 60
        If lngDesiredValue < 0 Then lngDesiredValue = 0
        .width = lngDesiredValue
    End With
    
    With lvwColorKey
        .Top = 110
        .Left = cboDataDisplayMode.Left + cboDataDisplayMode.width + 120
        lngDesiredValue = fraViewOptions.width - .Left - 120
        If lngDesiredValue < 1000 Then lngDesiredValue = 1000
        .width = lngDesiredValue
        .Height = fraViewOptions.Height - .Top - 50
        
        If .ColumnHeaders.Count >= 2 Then
            .ColumnHeaders(lvkUMCColor + 1).width = 700
            .ColumnHeaders(lvkName + 1).width = .width - .ColumnHeaders(lvkUMCColor + 1).width - 90
        End If
    End With
    
    With cmdHorizontalDivider
        .Left = cmdRollUpShrink.Left
        lngDesiredValue = cmdVerticalDivider.Left - 30
        If lngDesiredValue < 0 Then lngDesiredValue = 0
        .width = lngDesiredValue
        .Top = mDividerYLoc
        .Height = 150
    End With
    
    With lvwORFs
        .Top = cmdVerticalDivider.Top
        .Left = 60
        lngDesiredValue = cmdHorizontalDivider.Top - .Top - 60
        If lngDesiredValue < 0 Then lngDesiredValue = 0
        .Height = lngDesiredValue
        lngDesiredValue = cmdVerticalDivider.Left - lvwORFs.Left - 60
        If lngDesiredValue < 0 Then lngDesiredValue = 0
        lvwORFs.width = lngDesiredValue
        .ZOrder
    End With

    With lvwMassTags
        .Left = lvwORFs.Left
        .Top = cmdHorizontalDivider.Top + cmdHorizontalDivider.Height + 60
        lngDesiredValue = (cmdVerticalDivider.Top + cmdVerticalDivider.Height) - .Top - cboItemCountSourceGel.Height - 60
        If lngDesiredValue < 0 Then lngDesiredValue = 0
        .Height = lngDesiredValue
        .width = lvwORFs.width
        .ZOrder
    End With
    
    With cboItemCountSourceGel
        .Left = lvwORFs.Left
        .Top = lvwMassTags.Top + lvwMassTags.Height + 60
        .width = lvwORFs.width
    End With
    
    With VScroll
        .Top = lvwORFs.Top
        lngDesiredValue = Me.ScaleWidth - .width - 30
        lngCompareValue = lvwORFs.Left + lvwORFs.width + 120
        If lngDesiredValue < lngCompareValue Then
            lngDesiredValue = lngCompareValue
        End If
        .Left = lngDesiredValue
        .Height = (lvwMassTags.Top + lvwMassTags.Height) - lvwORFs.Top
        
        ' Make sure VScroll is on top
        .ZOrder 0
    End With
    
    fraORFPicsClippingRegion.BackColor = vbWhite
    fraORFPicsContainer.BackColor = vbWhite
    
    PositionPictures blnForceORFPicRearrange
End Sub

Private Sub PositionPictures(Optional blnForceORFPicRearrange As Boolean = False)
    ' Will arrange the loaded (and visible) picORFs, provided at least MIN_MSEC_BETWEEN_UPDATES msec
    '   has elapsed since the last call to this function (or, if blnForceORFPicRearrange is True)
    
    Static lngTickCountLastUpdate As Long
    
    Dim lngCurrentTickCount As Long
    Dim blnArrangeORFPics As Boolean
    Dim lngIndex As Long
    Dim lngDesiredValue As Long
    
    Dim lngPicOriginTop As Long, lngPicOriginLeft As Long
    Dim lngColumnIndex As Long              ' Column 0 is the left-most column
    Dim lngRowIndex As Long                 ' Row 0 is the top row
    Dim lngVisibleRows As Long, lngCurrentVisibleTopRowIndex As Long
    
    Dim lngPicWidthWithSpacing As Long
    Dim lngColumnsPerRow As Long, lngRequiredRows As Long
    
    ' Exam the system clock Tick count to determine if it's OK to arrange the pictures
    lngCurrentTickCount = GetTickCount()     ' Note that GetTickCount returns a negative number after 24 days of computer Uptime and resets to 0 after 48 days
    
    If (lngCurrentTickCount - lngTickCountLastUpdate) >= MIN_MSEC_BETWEEN_UPDATES Or _
       blnForceORFPicRearrange Then
        lngTickCountLastUpdate = lngCurrentTickCount
        blnArrangeORFPics = True
    End If
    
    If blnArrangeORFPics Then
        
        With fraORFPicsClippingRegion
            .Left = lvwORFs.Left + lvwORFs.width + cmdVerticalDivider.width + 120
            .Top = VScroll.Top
            lngDesiredValue = VScroll.Left - .Left - 120
            If lngDesiredValue < 0 Then lngDesiredValue = 0
            .width = lngDesiredValue
            .Height = VScroll.Height
        End With
        
        lngPicOriginTop = 0
        lngPicOriginLeft = 0
        
        lngPicWidthWithSpacing = mPicWidth + mPicSpacing
        
        With fraORFPicsContainer
            .Left = 0
            lngDesiredValue = fraORFPicsClippingRegion.width
            If lngDesiredValue < lngPicWidthWithSpacing Then lngDesiredValue = lngPicWidthWithSpacing
            .width = lngDesiredValue
            
            ' Note: Use CSng() to guarantee a floating point result
            lngColumnsPerRow = RoundToNearest(.width / CSng(lngPicWidthWithSpacing), 1, False)
        
            lngRequiredRows = RoundToNearest(mORFPicsUseCount / CSng(lngColumnsPerRow), 1, True)
            If lngRequiredRows < 1 Then lngRequiredRows = 1
            
            .Height = lngRequiredRows * (mPicHeight + mPicSpacing)
        End With
        
        lngColumnIndex = 0
        lngRowIndex = 0
        
        For lngIndex = 0 To mORFPicsUseCount - 1
            With ctlORF(lngIndex)
                .Height = mPicHeight
                .width = mPicWidth
                
                .Top = lngPicOriginTop + lngRowIndex * (mPicHeight + mPicSpacing)
                .Left = lngPicOriginLeft + lngColumnIndex * (mPicWidth + mPicSpacing)
                
                lngColumnIndex = lngColumnIndex + 1
                If lngColumnIndex >= lngColumnsPerRow Then
                    ' I put the following check in to avoid incrementing lngRowIndex when the last
                    ' picture to display is in the right-most column
                    If lngIndex < mORFPicsUseCount - 1 Then
                        lngRowIndex = lngRowIndex + 1
                        lngColumnIndex = 0
                    End If
                End If
            End With
        Next lngIndex
                
        Debug.Assert lngRowIndex + 1 = lngRequiredRows
        
        ' Adjust VSCroll.Value if needed
        lngCurrentVisibleTopRowIndex = VScroll.value
        If lngCurrentVisibleTopRowIndex > lngRowIndex Then
            lngCurrentVisibleTopRowIndex = lngRowIndex
        End If
        
        mORFPicsRowCount = lngRowIndex + 1
        With VScroll
            .Min = 0
            .Max = lngRowIndex
            .SmallChange = 1
            
            ' Determine the number of visible rows
            lngVisibleRows = Int(fraORFPicsClippingRegion.Height / CSng((mPicHeight + mPicSpacing)))
            
            ' Set the LargeChange to the number of visible rows
            If lngVisibleRows < 1 Then lngVisibleRows = 1
            .LargeChange = lngVisibleRows
            
            .value = lngCurrentVisibleTopRowIndex
        End With
    End If

End Sub
                
Private Sub SelectORFByID(lngORFGroupArrayIndex As Long)
    Dim lngItemIndex As Long
    Dim strOrfID As String
    
    strOrfID = Trim(lngORFGroupArrayIndex)
    For lngItemIndex = 1 To lvwORFs.ListItems.Count
        If lvwORFs.ListItems(lngItemIndex).Text = strOrfID Then
            ListViewHighlightItem lvwORFs, lngItemIndex
            HandleListViewItemClick lvwORFs.ListItems(lngItemIndex), lviORFs
        End If
    Next lngItemIndex

End Sub

Public Sub SetFormID(lngNewFormID As Long)
    ' WARNING: This sub should only be called by the ORFViewerLoaderClass
    
    mFormID = lngNewFormID
End Sub

Public Sub SetPictureDimensionsAndSpacing(Optional lngPicSpacing As Long = 15, Optional lngPicHeight As Long = 2000, Optional lngPicWidth As Long = 2000)
    
    If lngPicSpacing < 1 Then lngPicSpacing = 1
    mPicSpacing = lngPicSpacing
    
    If lngPicHeight < 1 Then lngPicHeight = 1
    mPicHeight = lngPicHeight
    
    If lngPicWidth < 1 Then lngPicWidth = 1
    mPicWidth = lngPicWidth
    
    PositionControls
End Sub

Private Sub SetVisiblePictureCount(ByVal intDesiredTotalPictures As Integer, Optional ByVal blnPositionPictures As Boolean = True)
    ' Shows/hides pictures so that the desired number are shown
    ' If intDesiredTotalPictures is greater than mORFPicsLoadedCount
    '  then loads new pictures as needed
    
    Dim intIndex As Integer
    Dim eResponse As VbMsgBoxResult
    Dim blnShowProgess As Boolean
    
    If intDesiredTotalPictures > mORFPicsLoadedCount Then
        Do While intDesiredTotalPictures > mORFPicsMaxLoadedCount
            eResponse = MsgBox(Trim(intDesiredTotalPictures) & " ORF pictures are required, but the currently defined maximum is " & Trim(mORFPicsMaxLoadedCount) & "; the current maximum can be increased by 50 pictures, but the processing time required to actually display all of the pictures may be quite long.  Actually increase the maximum by 50?", vbQuestion + vbYesNo + vbDefaultButton2, "Increase ORF Picture Count")
            If eResponse = vbYes Then
                mORFPicsMaxLoadedCount = mORFPicsMaxLoadedCount + 50
            Else
                If intDesiredTotalPictures > mORFPicsMaxLoadedCount Then
                    intDesiredTotalPictures = mORFPicsMaxLoadedCount
                End If
            End If
        Loop
        
        If intDesiredTotalPictures - mORFPicsLoadedCount > 10 Then blnShowProgess = True
            
        If blnShowProgess Then frmProgress.InitializeForm "Loading plot controls", 0, CLng(intDesiredTotalPictures - mORFPicsLoadedCount), True, False, False, MDIForm1
        
        For intIndex = mORFPicsLoadedCount To intDesiredTotalPictures - 1
            Load ctlORF(intIndex)
            
            If blnShowProgess Then frmProgress.UpdateProgressBar CLng(intIndex - mORFPicsLoadedCount)
        Next intIndex
        
        If blnShowProgess Then frmProgress.HideForm
        
        mORFPicsLoadedCount = intDesiredTotalPictures
    End If
    
    ' Show/hide the correct pictures
    For intIndex = 0 To mORFPicsLoadedCount - 1
        ctlORF(intIndex).Visible = (intIndex < intDesiredTotalPictures)
    Next intIndex
    
    mORFPicsUseCount = intDesiredTotalPictures
    
    If blnPositionPictures Then
        PositionControls False, True
        HandleVerticalScroll True
    End If
    
End Sub

Private Sub ShowORFViewerOptions()
    ORFViewerOptionsInitializeForm mGelDisplayListAndOptions, mFormID
End Sub

Private Sub UpdateMassIndexingArrays()

    Const MASS_TAG_DIM_CHUNK = 500
    
    Dim objQSDouble As New QSDouble
    
    Dim lngGelIndex As Long
    Dim intIncludedGelCount As Integer
    Dim lngPointerIndex As Long, lngUMCIndex As Long
    Dim lngUMCMemberIndex As Long
    Dim lngDataIndex As Long
    
    Dim lngORFGroupArrayIndex As Long, lngORFIndex As Long
    Dim lngMassTagIndex As Long
    
    Dim lngMassTagDimCount As Long
    Dim intIsoMWField As Integer, intUMCMWField As Integer
    Dim lngMinScanNumber As Long, lngMaxScanNumber As Long
    Dim lngClassRepresentativeScanNumber As Long
    
    Dim dblIonMass As Double, dblIonMassMin As Double, dblIonMassMax As Double
    
    mGelDataIndices.GelCount = mGelDisplayListAndOptions.GelCount
    ReDim mGelDataIndices.Gels(mGelDataIndices.GelCount)
        
    intIncludedGelCount = GetNumberOfIncludedGels(mGelDisplayListAndOptions)
    If intIncludedGelCount = 0 Then Exit Sub
    
    ' Note: Need to add 1 to intIncludedGelCount due to the Mass Tag Index Array step following the
    '         lngGelIndex For-Next loop
    frmProgress.InitializeForm "Sorting mass index arrays", 0, CLng(intIncludedGelCount) + 1, True, True, True, MDIForm1
    
On Error GoTo UpdateMassIndexingArraysErrorHandler

    intIncludedGelCount = 0
    For lngGelIndex = 1 To mGelDisplayListAndOptions.GelCount
        If mGelDisplayListAndOptions.Gels(lngGelIndex).IncludeGel Then
            frmProgress.InitializeSubtask "Examining Ions", 0, 3
            
            If Not GelData(lngGelIndex).CustomNETsDefined Then
                Dim dblNETSlope As Double, dblNETIntercept As Double
                
                ' Determine NETSlope and NETIntercept for this gel
                GetNETSlopeAndIntercept mGelDisplayListAndOptions.Gels(lngGelIndex).NETAdjustmentType, lngGelIndex, dblNETSlope, dblNETIntercept
            
                If dblNETSlope = 0 And dblNETIntercept = 0 Then
                    GetNETSlopeAndIntercept natGeneric, lngGelIndex, dblNETSlope, dblNETIntercept
                    MsgBox "Warning, the NET slope and NET intercept for gel " & GelBody(lngGelIndex).Caption & " are both 0.  This will lead to incorrect NET values.  Using a generic slope and intercept instead.", vbInformation + vbOKOnly, "Warning"
                ElseIf dblNETSlope = 1 And dblNETIntercept = 0 And mGelDisplayListAndOptions.Gels(lngGelIndex).NETAdjustmentType <> natGeneric Then
                    MsgBox "Warning, the NET slope and NET intercept for gel " & GelBody(lngGelIndex).Caption & " are 1 and 0, but you have selected usage of TIC NET or GANET, which never have values of 1 and 0.  There is probably a problem with the NET values for this gel.", vbInformation + vbOKOnly, "Warning"
                End If
            End If
            
            
        
' Step 1: Fill and sort the Ion Mass and NET Index Arrays

            With mGelDataIndices.Gels(lngGelIndex)
            
                ' Record the desired molecular weight field type
                intIsoMWField = GelData(lngGelIndex).Preferences.IsoDataField

                ' Use GetISScope() to fill .IonMass() and .IonMassPointer() with either
                '   all of the data, or just the data in the current scope
                ' Note that the GetISScope() function will ReDim lngPointer()
                
                If mGelDisplayListAndOptions.Gels(lngGelIndex).VisibleScopeOnly Then
                    ' Only show data visible in the current scope
                    ' This allows only filtered data to be visible
                    ' Need to make sure GelBody() window is fully zoomed out
                    ZoomOutGelBody lngGelIndex
                    
                    ' Retrieve an array of the ion indices of the ions currently "In Scope"
                    .IonCount = GetISScope(lngGelIndex, .IonMassPointer(), glScope.glSc_Current)
                Else
                    ' Use all of the data points
                    .IonCount = GetISScope(lngGelIndex, .IonMassPointer(), glScope.glSc_All)
                End If
            
                ' Copy the ion masses into .IonMass()
                ' In addition, precompute the NET values for all of the ions
                ReDim .IonMass(.IonCount)
                For lngPointerIndex = 1 To .IonCount
                    .IonMass(lngPointerIndex) = GetIsoMass(GelData(lngGelIndex).IsoData(.IonMassPointer(lngPointerIndex)), intIsoMWField)
                Next lngPointerIndex
                
                ' Next sort .IonMass() ascending
                ' The order of the masses in .IonMass() will be changed, and
                '   the .IonMassPointer() array will be updated to allow dereferencing
                If Not objQSDouble.QSAsc(.IonMass(), .IonMassPointer()) Then
                    ' Failure with QSort
                    Debug.Assert False
                End If
                
                ' Now that .IonMass() has been sorted, the following relations hold
                ' item .IonMass(1000) is the 1000th heaviest ion mass
                '  .IonMassPointer(1000) contains the index needed to lookup information in the .IsoData()
                '  or in the .IonNET() arrays
                
                ' We can now precompute the NET values for all of the ions
                ' We couldn't do this above because objQSDouble.QSAsc() sorts .IonMass() and .IonMassPointer(),
                '  but has no effect on .IonNET(), so we wouldn't be able to find the correct .IonNET value for the
                '  data point in .IonMass(x)
                ReDim .IonNET(.IonCount)
                For lngPointerIndex = 1 To .IonCount
                    .IonNET(lngPointerIndex) = ScanToGANET(lngGelIndex, GelData(lngGelIndex).IsoData(.IonMassPointer(lngPointerIndex)).ScanNumber)
                Next lngPointerIndex
                
            End With
            
' Step 2: Fill the UMC Index Arrays
            
            ' Now determine the minimum and maximum mass of each UMC
            ' In addition, record the StartScan and EndScan of the UMC (obtained using UMCStatistics2)
            ' Pre-determination of the start and end scans for the class will speed the selection process
                
            frmProgress.UpdateCurrentSubTask "Examining UMC's"
            frmProgress.UpdateSubtaskProgressBar 1
            
            With mGelDataIndices.Gels(lngGelIndex)
                
                .UMCCount = GelUMC(lngGelIndex).UMCCnt
                                    
                ReDim .UMCMassMin(.UMCCount)
                ReDim .UMCMassMinPointer(.UMCCount)
                
                ReDim .UMCMassMax(.UMCCount)
                ReDim .UMCMassMaxPointer(.UMCCount)
                
                ReDim .UMCScanRange(.UMCCount)
                
            End With
            
            intUMCMWField = GelUMC(lngGelIndex).def.MWField
            
            ' Note: Unlike the IsoData() array in GelData(), the UMC arrays are 0-based
            ' Further, we will use all of the UMC's when populating these arrays
            For lngUMCIndex = 0 To GelUMC(lngGelIndex).UMCCnt - 1
                If GelUMC(lngGelIndex).UMCs(lngUMCIndex).ClassCount = 0 Then
                    With mGelDataIndices.Gels(lngGelIndex).UMCScanRange(lngUMCIndex)
                        .ScanNumberStart = 0
                        .ScanNumberEnd = 0
                        .NETStart = 0
                        .NETEnd = 0
                        .NETClassRepresentative = 0
                    End With
                Else
                    With GelUMC(lngGelIndex).UMCs(lngUMCIndex)
                    
                        dblIonMassMin = 1E+300
                        dblIonMassMax = -1E+300
                        
                        ' Step through the class members to find the minimum and maximum mass
                        For lngUMCMemberIndex = 0 To .ClassCount - 1
                            lngDataIndex = .ClassMInd(lngUMCMemberIndex)
                            If lngDataIndex <= GelData(lngGelIndex).IsoLines Then
                                Select Case .ClassMType(lngUMCMemberIndex)
                                Case glCSType
                                    dblIonMass = GelData(lngGelIndex).CSData(lngDataIndex).AverageMW
                                Case glIsoType
                                    
                                    dblIonMass = GetIsoMass(GelData(lngGelIndex).IsoData(lngDataIndex), intUMCMWField)
                                End Select
                            End If
                            
                            If dblIonMass < dblIonMassMin Then
                                dblIonMassMin = dblIonMass
                            End If
                            
                            If dblIonMass > dblIonMassMax Then
                                dblIonMassMax = dblIonMass
                            End If
                            
                        Next lngUMCMemberIndex
    
                        mGelDataIndices.Gels(lngGelIndex).UMCMassMin(lngUMCIndex) = dblIonMassMin
                        mGelDataIndices.Gels(lngGelIndex).UMCMassMax(lngUMCIndex) = dblIonMassMax
                        
                        mGelDataIndices.Gels(lngGelIndex).UMCMassMinPointer(lngUMCIndex) = lngUMCIndex
                        mGelDataIndices.Gels(lngGelIndex).UMCMassMaxPointer(lngUMCIndex) = lngUMCIndex
                    
                        ' Determine the first and last scan number
                        ' Class members are ordered on scan numbers
                        ' First scan number
                        lngDataIndex = .ClassMInd(0)
                        If lngDataIndex >= 1 And lngDataIndex <= GelData(lngGelIndex).IsoLines Then
                            Select Case .ClassMType(0)
                            Case gldtCS
                                 lngMinScanNumber = GelData(lngGelIndex).CSData(lngDataIndex).ScanNumber
                            Case gldtIS
                                 lngMinScanNumber = GelData(lngGelIndex).IsoData(lngDataIndex).ScanNumber
                            End Select
                        Else
                            lngMinScanNumber = 1
                        End If
                        
                        ' Last scan number
                        lngDataIndex = .ClassMInd(.ClassCount - 1)
                        If lngDataIndex <= GelData(lngGelIndex).IsoLines Then
                            Select Case .ClassMType(.ClassCount - 1)
                            Case gldtCS
                                lngMaxScanNumber = GelData(lngGelIndex).CSData(lngDataIndex).ScanNumber
                            Case gldtIS
                                lngMaxScanNumber = GelData(lngGelIndex).IsoData(lngDataIndex).ScanNumber
                            End Select
                        Else
                            lngMaxScanNumber = lngMinScanNumber
                        End If
                    
                        mGelDataIndices.Gels(lngGelIndex).UMCScanRange(lngUMCIndex).ScanNumberStart = lngMinScanNumber
                        mGelDataIndices.Gels(lngGelIndex).UMCScanRange(lngUMCIndex).ScanNumberEnd = lngMaxScanNumber
                        
                        ' Determine the ScanNumber of the "Class Representative"
                        If .ClassRepInd <= GelData(lngGelIndex).IsoLines Then
                            Select Case .ClassRepType
                            Case gldtCS
                                lngClassRepresentativeScanNumber = GelData(lngGelIndex).CSData(.ClassRepInd).ScanNumber
                            Case gldtIS
                                lngClassRepresentativeScanNumber = GelData(lngGelIndex).IsoData(.ClassRepInd).ScanNumber
                            End Select
                        Else
                            lngClassRepresentativeScanNumber = lngMinScanNumber
                        End If
                    
                    End With
                
                    With mGelDataIndices.Gels(lngGelIndex).UMCScanRange(lngUMCIndex)
                        ' Compute the NET Range of the UMC, based on  the scan number range
                        .NETStart = ScanToGANET(lngGelIndex, .ScanNumberStart)
                        .NETEnd = ScanToGANET(lngGelIndex, .ScanNumberEnd)
                        If lngClassRepresentativeScanNumber > 0 Then
                            .NETClassRepresentative = ScanToGANET(lngGelIndex, lngClassRepresentativeScanNumber)
                        End If
                    End With
                End If
            Next lngUMCIndex
            
' Step 3: Sort the UMC Index Arrays
            
            frmProgress.UpdateCurrentSubTask "Examining UMC's"
            frmProgress.UpdateSubtaskProgressBar 2
            
            With mGelDataIndices.Gels(lngGelIndex)
            
                If Not objQSDouble.QSAsc(.UMCMassMin(), .UMCMassMinPointer()) Then
                    ' Failure with QSort
                    Debug.Assert False
                End If
            
                If Not objQSDouble.QSAsc(.UMCMassMax(), .UMCMassMaxPointer()) Then
                    ' Failure with QSort
                    Debug.Assert False
                End If
            End With
            
            intIncludedGelCount = intIncludedGelCount + 1
            frmProgress.UpdateProgressBar intIncludedGelCount
            If KeyPressAbortProcess > 1 Then
                AddToAnalysisHistory lngGelIndex, "User prematurely aborted update of mass indexing arrays in ORF viewer"
                Exit For
            End If
        End If
    Next lngGelIndex
    
    
' Final Step: Fill the MassTagMass() array with masses of all of the mass tags

    frmProgress.InitializeSubtask "Examining Mass Tags for all ORFs", 0, mORFViewerGroupList.ORFCount

    With mMassTagDataIndex
        .MassTagCount = 0
        lngMassTagDimCount = MASS_TAG_DIM_CHUNK
        
        ReDim .MassTagMass(MASS_TAG_DIM_CHUNK)
        ReDim .MassTagMassPointer(MASS_TAG_DIM_CHUNK)
        ReDim .MassTagLookupInfo(MASS_TAG_DIM_CHUNK)
    
    End With
    
    For lngORFGroupArrayIndex = 0 To mORFViewerGroupList.ORFCount - 1
        ' Examine the mass tags in Item 0 of each ORF
        ' No need to examine the other items, since they have equivalent Mass Tags
        If mORFViewerGroupList.Orfs(lngORFGroupArrayIndex).ItemCount > 0 Then
            lngGelIndex = mORFViewerGroupList.Orfs(lngORFGroupArrayIndex).Items(0).GelIndex
            lngORFIndex = mORFViewerGroupList.Orfs(lngORFGroupArrayIndex).Items(0).ORFIndex
            
            With GelORFMassTags(lngGelIndex).Orfs(lngORFIndex)
                
                For lngMassTagIndex = 0 To .MassTagCount - 1
                    
                    ' Add an entry to mMassTagDataIndex
                    ' I realize I could move this into a subroutine to make it prettier,
                    '  but the execution speed would decrease
                    
                    mMassTagDataIndex.MassTagMass(mMassTagDataIndex.MassTagCount) = .MassTags(lngMassTagIndex).Mass
                    
                    mMassTagDataIndex.MassTagMassPointer(mMassTagDataIndex.MassTagCount) = mMassTagDataIndex.MassTagCount
                    
                    mMassTagDataIndex.MassTagLookupInfo(mMassTagDataIndex.MassTagCount).GelIndex = lngGelIndex
                    mMassTagDataIndex.MassTagLookupInfo(mMassTagDataIndex.MassTagCount).ORFIndex = lngORFIndex
                    mMassTagDataIndex.MassTagLookupInfo(mMassTagDataIndex.MassTagCount).MassTagIndex = lngMassTagIndex
                    
                    mMassTagDataIndex.MassTagCount = mMassTagDataIndex.MassTagCount + 1
                    If mMassTagDataIndex.MassTagCount >= lngMassTagDimCount Then
                        lngMassTagDimCount = lngMassTagDimCount + MASS_TAG_DIM_CHUNK
                        ReDim Preserve mMassTagDataIndex.MassTagMass(lngMassTagDimCount)
                        ReDim Preserve mMassTagDataIndex.MassTagMassPointer(lngMassTagDimCount)
                        ReDim Preserve mMassTagDataIndex.MassTagLookupInfo(lngMassTagDimCount)
                    End If
                Next lngMassTagIndex
            End With
        End If
        
        If lngORFGroupArrayIndex Mod 50 = 0 Then
            frmProgress.UpdateSubtaskProgressBar lngORFGroupArrayIndex
            If KeyPressAbortProcess > 1 Then
                AddToAnalysisHistory lngGelIndex, "User prematurely aborted update of mass indexing arrays in ORF Viewer"
                Exit For
            End If
        End If
    Next lngORFGroupArrayIndex
    
    If mMassTagDataIndex.MassTagCount > 0 Then
        
        ' Sort the .MassTagMass() array
        ' First need to ReDim the arrays to remove unused memory space
        lngMassTagDimCount = mMassTagDataIndex.MassTagCount
        ReDim Preserve mMassTagDataIndex.MassTagMass(0 To lngMassTagDimCount - 1)
        ReDim Preserve mMassTagDataIndex.MassTagMassPointer(0 To lngMassTagDimCount - 1)
        ReDim Preserve mMassTagDataIndex.MassTagLookupInfo(0 To lngMassTagDimCount - 1)
        
        If Not objQSDouble.QSAsc(mMassTagDataIndex.MassTagMass(), mMassTagDataIndex.MassTagMassPointer()) Then
            ' Failure with QSort
            Debug.Assert False
        End If
    
    End If
    
    ' Update the progress bar once more to show 100% completed
    frmProgress.UpdateProgressBar intIncludedGelCount + 1
    
    ' Make sure objMwtWin is set to the correct mass type
    If gMwtWinLoaded Then
        If intIsoMWField = 6 Then
            ' Use average masses
            objMwtWin.SetElementMode emAverageMass
        Else
            ' intIsoMWField = 7 (isotopic) or 8 (most abundant)
            ' Use isotopic
            objMwtWin.SetElementMode emIsotopicMass
        End If
    End If
    
    Set objQSDouble = Nothing
    
    frmProgress.HideForm
    
    Exit Sub
    
UpdateMassIndexingArraysErrorHandler:
    If Err.Number = -2147024770 Or Err.Number = 429 Then
        MsgBox "Error connecting to MwtWinDll.Dll; you probably need to re-install this application or the Molecular Weight Calculator to properly register the DLL", vbExclamation + vbOKOnly, "Error"
    Else
        MsgBox "Error in function UpdateMassIndexingArrays: " & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    frmProgress.HideForm
    
End Sub

Private Sub UpdateORFHistoryMenuList()
    Dim intHistoryIndex As Integer
    Dim intMenuIndex As Integer
    
    ' Make sure we have enough menu items
    With mORFHistory
        Do While mORFHistoryMenuItemsLoadedCount < .HistoryCount
            Load mnuORFHistoryListItem(mORFHistoryMenuItemsLoadedCount)
            mORFHistoryMenuItemsLoadedCount = mORFHistoryMenuItemsLoadedCount + 1
        Loop
        
        For intHistoryIndex = 0 To .HistoryCount - 1
            mnuORFHistoryListItem(intHistoryIndex).Caption = Trim(.History(intHistoryIndex).ORFGroupArrayIndex) & vbTab & .History(intHistoryIndex).ORFDescription
            If .CurrentHistoryIndex = intHistoryIndex Then
                mnuORFHistoryListItem(intHistoryIndex).Checked = True
            Else
                mnuORFHistoryListItem(intHistoryIndex).Checked = False
            End If
            mnuORFHistoryListItem(intHistoryIndex).Visible = True
        Next intHistoryIndex
    
        ' Enable/Disable the Previous/Next menu items as needed
        If .CurrentHistoryIndex = 0 Then
            mnuORFHistoryMovePrevious.Enabled = False
        Else
            mnuORFHistoryMovePrevious.Enabled = True
        End If
        
        If .CurrentHistoryIndex = .HistoryCount - 1 Or .HistoryCount = 0 Then
            mnuORFHistoryMoveNext.Enabled = False
        Else
            mnuORFHistoryMoveNext.Enabled = True
        End If
        
        ' Hide any remaining menus
        ' VB requires that at least one submenu be visible at a given time
        ' Therefore, the following sometimes produces an error
        ' Thus, we'll enable On Error Resume Next handling
        On Error Resume Next
        For intMenuIndex = .HistoryCount To mORFHistoryMenuItemsLoadedCount - 1
            mnuORFHistoryListItem(intMenuIndex).Caption = ""
            mnuORFHistoryListItem(intMenuIndex).Checked = False
            mnuORFHistoryListItem(intMenuIndex).Visible = False
        Next intMenuIndex
    
    End With
End Sub

Private Sub UpdateORFPics(Optional lngOrfDataIndexOverride As Long = -1)
    ' Display data for selected ORF in ctlORF()'s
    
    Dim objMWSearch As MWUtil
    
    Dim lngORFDataIndexSelected As Long
    Dim lngGelIndexForORF As Long, lngORFIndex As Long
    Dim lngListIndex As Long
    Dim lngMassTagIndex As Long, lngMassTagCount As Long
    Dim lngGelIndexForIon As Long
    Dim intSeriesCount As Integer
    Dim intPictureIndex As Integer
    Dim intHiddentMTPicCount As Integer
    Dim lngMaxSpotSizePixels As Long, lngMinSpotSizePixels As Long
    
    Dim blnProceed As Boolean
    Dim dblAMTNet As Double, dblAMTMass As Double
    Dim dblNETRange As Double, dblNETRangeHalfWindow
    Dim dblMassRangePPM As Double, dblMassRange As Double, dblMassRangeHalfWindow As Double
    Dim lngMassTagID As Long, lngAMTMatchCount As Long, lngMassTagHitCount As Long
    Dim strRuleDescription As String, strMatches As String
    Dim eDataDisplayMode As ddmDataDisplayModeConstants
    Dim blnUseClassRepresentativeNET As Boolean
    Dim blnIonsOrUMCsPresent As Boolean
    Dim blnShowProgress As Boolean
    
    Dim udtData As udtDataToPopulateType
    Dim lngDataCount As Long
    
    Dim objQSDouble As New QSDouble
    
    If lngOrfDataIndexOverride < 0 Then
        lngORFDataIndexSelected = GetSelectedORFDataIndex()
    Else
        lngORFDataIndexSelected = lngOrfDataIndexOverride
    End If
    
    mMTCountWithHitsForThisORF = 0
    intHiddentMTPicCount = 0
    
    If lngORFDataIndexSelected < 0 Then
        SetVisiblePictureCount 0
    Else
        If lngORFDataIndexSelected >= mORFViewerGroupList.ORFCount Then
            ' This should not happen
            Debug.Assert False
        Else
            With mORFViewerGroupList.Orfs(lngORFDataIndexSelected)
                ' mORFViewerGroupList only contains pointers to a GelIndex and an ORFIndex
                ' Need to find these so we can reference the correct Mass Tag in GelORFMassTags()
                If .ItemCount >= 1 Then
                    lngGelIndexForORF = .Items(0).GelIndex
                    lngORFIndex = .Items(0).ORFIndex
                    blnProceed = True
                Else
                    SetVisiblePictureCount 0
                End If
            End With
            
            If blnProceed Then
                
                With mGelDisplayListAndOptions.DisplayOptions
                    
                    dblNETRange = .NETDisplayRange
                    If dblNETRange <= 0 Then dblNETRange = 0.2
                    
                    dblMassRangePPM = .MassDisplayRangePPM
                    If dblMassRangePPM <= 0 Then dblMassRangePPM = 25
                    
                    eDataDisplayMode = .DataDisplayMode
                    blnUseClassRepresentativeNET = .UseClassRepresentativeNET
                    
                    lngMaxSpotSizePixels = .MaxSpotSizePixels
                    lngMinSpotSizePixels = .MinSpotSizePixels
                    
                    If .IonToUMCPlottingIntensityRatio = 0 Then .IonToUMCPlottingIntensityRatio = 1
                    
                    Select Case .CleavageRuleID
                    Case 1: strRuleDescription = "Tryptic"
                    Case 2: strRuleDescription = "Stryptic"
                    Case Else: strRuleDescription = "Cleavage match"
                    End Select
                End With

                ' Could determine lngMassTagCount using this:
                ' lngMassTagCount = GelORFMassTags(lngGelIndexForORF).Orfs(lngORFIndex).MassTagCount
                
                ' However, we only want to show pictures for the ions in lvwMassTags
                lngMassTagCount = lvwMassTags.ListItems.Count
                
                SetVisiblePictureCount CInt(lngMassTagCount)

                If lngMassTagCount > 0 Then
                    
                    ' Now draw the pictures, drawing them in the same order as the mass tags in lvwMassTags
                    ' Do not draw the picture if it isn't in lvwMassTags (i.e. in case the user has filtered the mass tags)
                    
                    Set objMWSearch = New MWUtil
                    
                    If lvwMassTags.ListItems.Count > 50 Then
                        blnShowProgress = True
                        frmProgress.InitializeForm "Drawing plots", 0, lvwMassTags.ListItems.Count, True, False, False, MDIForm1
                    End If
                    
                    intPictureIndex = 0
                    For lngListIndex = 1 To lvwMassTags.ListItems.Count
                        If lngListIndex > mORFPicsLoadedCount Then
                            ' The user has limited the number of ORF pictures that will be displayed (in sub SetVisiblePictureCount)
                            ' Thus, do not display this mass tag
                        Else
                            lngMassTagIndex = CLngSafe(lvwMassTags.ListItems(lngListIndex).Text)
                            
                            blnIonsOrUMCsPresent = False
                            If lngMassTagIndex < GelORFMassTags(lngGelIndexForORF).Orfs(lngORFIndex).MassTagCount Then
                                With GelORFMassTags(lngGelIndexForORF).Orfs(lngORFIndex).MassTags(lngMassTagIndex)
                                    dblAMTNet = .GANET
                                    dblAMTMass = .Mass
                    
                                    If dblAMTMass < 0 Then dblAMTMass = 0
                                    lngMassTagID = .MassTagRefID
                                End With
                                
                                dblMassRange = PPMToMass(dblMassRangePPM, dblAMTMass)
                                dblMassRangeHalfWindow = dblMassRange / 2
                                dblNETRangeHalfWindow = dblNETRange / 2
                                
                                ' Initialize ctlORF for this Mass Tag
                                With ctlORF(intPictureIndex)
                                    .ClearGraphAndData
    
                                    .SetFixedWindow dblAMTNet - dblNETRange / 2, dblAMTNet + dblNETRange / 2, dblAMTMass - dblMassRange / 2, dblAMTMass + dblMassRange / 2
                                    .HNumFmt = mXAxisFormatLabel
                                    .VNumFmt = mYAxisFormatLabel
                                    .SwapAxes = mGelDisplayListAndOptions.DisplayOptions.SwapPlottingAxes
    
                                    .FontWidth = 320
                                    .FontHeight = 800
    
                                    If .SwapAxes Then
                                        .HLabel = "Mass"
                                        .VLabel = "NET"
                                    Else
                                        .HLabel = "NET"
                                        .VLabel = "Mass"
                                    End If
    
                                    .SetMaxSpotSize lngMaxSpotSizePixels, True
                                    .SetMinSpotSize lngMinSpotSizePixels, True
    
                                    .IntensityLogScale = mGelDisplayListAndOptions.DisplayOptions.LogarithmicIntensityPlotting
    
                                    .ShowPosition = mGelDisplayListAndOptions.DisplayOptions.ShowPosition
                                    .ShowTickMarkLabels = mGelDisplayListAndOptions.DisplayOptions.ShowTickMarkLabels
                                    .ShowGridLines = mGelDisplayListAndOptions.DisplayOptions.ShowGridLines
    
                                    lngMassTagHitCount = GetMassTagItemHitCount(lngListIndex)
    
                                    If lngMassTagHitCount = 1 Then
                                        strMatches = " match"
                                    Else
                                        strMatches = " matches"
                                    End If
                                    ' .PlotTitleBottom = "MTID:" & lngMassTagID & ", " & lngMassTagHitCount & strMatches & "; " & strRuleDescription & ": " & lvwMassTags.ListItems(lngListIndex).SubItems(lvmTrypticStatus)
                                    .PlotTitleBottom = "MTID:" & lngMassTagID
    
                                    .CallingFormID = mFormID
                                    .Visible = True
                                End With
                                
                                ' Find the AMT's within the view window and add as series 0
                                ' This will include the mass tag at lngMassTagIndex, in addition to any other AMT's within the view range
                                With mGelDisplayListAndOptions.DisplayOptions
                                    intSeriesCount = 1
                                    lngDataCount = FindMassTagsInRange(objMWSearch, udtData, dblAMTNet, dblNETRangeHalfWindow, dblAMTMass, dblMassRangeHalfWindow, .MassTagNETError, .MassTagMassErrorPPM, lngMassTagID, lngORFIndex)
    
                                    ctlORF(intPictureIndex).AddSpotsManyWithExtents udtData.Labels(), udtData.NET(), udtData.NETExtentNeg(), udtData.NETExtentPos(), udtData.Mass(), udtData.MassExtentNeg(), udtData.MassExtentPos(), udtData.Intensity(), True, False, lngDataCount, intSeriesCount - 1, CInt(.MassTagSpotShape), .MassTagSpotColor, .MassTagSpotColor
                                End With
                                
                                ' Determine the Ions to show by examining each of the Gels in mGelDisplayListAndOptions
                                ' If the Gel is included, then find the data points within
                                '  the AMT's view range and add to the graph
                                For lngGelIndexForIon = 1 To mGelDisplayListAndOptions.GelCount
                                    If mGelDisplayListAndOptions.Gels(lngGelIndexForIon).IncludeGel Then
                                        ' Gel is included, find the ions within the view range
                                        If eDataDisplayMode = ddmIonsOnly Or eDataDisplayMode = ddmIonsAndUMCs Then
                                            intSeriesCount = intSeriesCount + 1
                                            lngDataCount = FindIonsInRange(objMWSearch, udtData, lngGelIndexForIon, dblAMTNet, dblNETRangeHalfWindow, dblAMTMass, dblMassRangeHalfWindow, lngMassTagID, lngAMTMatchCount)
                                            
                                            If lngDataCount > 0 Then blnIonsOrUMCsPresent = True
                                            
                                            ' Sort the data by increasing intensity so that it plots nicely
                                            SortDataToPlot udtData, objQSDouble
    
                                            ' Populate the graph ions (without extents)
                                            With mGelDisplayListAndOptions.Gels(lngGelIndexForIon)
                                                ctlORF(intPictureIndex).AddSpotsManyWithExtents udtData.Labels(), udtData.NET(), udtData.NETExtentNeg(), udtData.NETExtentPos(), udtData.Mass(), udtData.MassExtentNeg(), udtData.MassExtentPos(), udtData.Intensity(), False, True, lngDataCount, intSeriesCount - 1, CInt(.IonSpotShape), .IonSpotColor, .IonSpotColorSelected
                                            End With
                                        End If
                                        
                                        If eDataDisplayMode = ddmUMCsOnly Or eDataDisplayMode = ddmIonsAndUMCs Then
                                            intSeriesCount = intSeriesCount + 1
                                            lngDataCount = FindUMCsInRange(objMWSearch, udtData, lngGelIndexForIon, dblAMTNet, dblNETRangeHalfWindow, dblAMTMass, dblMassRangeHalfWindow, lngAMTMatchCount, blnUseClassRepresentativeNET)
                                            If lngDataCount > 0 Then blnIonsOrUMCsPresent = True
                                            
                                            ' Sort the data by increasing intensity so that it plots nicely
                                            SortDataToPlot udtData, objQSDouble
    
                                            ' Populate the graph with UMC's (using extents)
                                            With mGelDisplayListAndOptions.Gels(lngGelIndexForIon)
                                                ctlORF(intPictureIndex).AddSpotsManyWithExtents udtData.Labels(), udtData.NET(), udtData.NETExtentNeg(), udtData.NETExtentPos(), udtData.Mass(), udtData.MassExtentNeg(), udtData.MassExtentPos(), udtData.Intensity(), True, True, lngDataCount, intSeriesCount - 1, CInt(.UMCSpotShape), .UMCSpotColor, .UMCSpotColorSelected
                                            End With
                                            
                                        End If
                                        
                                    End If
                                Next lngGelIndexForIon
                                
                            Else
                                ' This shouldn't happen
                                Debug.Assert False
                                ctlORF(intPictureIndex).ClearGraphAndData
                            End If
                            
                            If blnIonsOrUMCsPresent Or Not mGelDisplayListAndOptions.DisplayOptions.HideEmptyMassTagPictures Then
                                intPictureIndex = intPictureIndex + 1
                                mMTCountWithHitsForThisORF = mMTCountWithHitsForThisORF + 1
                            Else
                                ctlORF(intPictureIndex).ClearGraphAndData
                                ctlORF(intPictureIndex).Visible = False
                                intHiddentMTPicCount = intHiddentMTPicCount + 1
                            End If
                        End If
                        
                        If blnShowProgress Then frmProgress.UpdateProgressBar lngListIndex
                    Next lngListIndex
                    
                    If blnShowProgress Then frmProgress.HideForm
                    
                    UpdateStatusORFDisplayCount lvwMassTags.ListItems.Count, intHiddentMTPicCount
                    
                    If mGelDisplayListAndOptions.DisplayOptions.HideEmptyMassTagPictures Then
                        SetVisiblePictureCount intPictureIndex
                    End If
                    
                    Set objMWSearch = Nothing
                End If
            End If
        End If
    End If
    
    Set objQSDouble = Nothing
End Sub

Private Sub UpdateDividerLimits()
    
    mDividerMinY = 120
    mDividerMaxY = fraViewOptions.Top - 360
    
    If mDividerMaxY <= mDividerMinY Then mDividerMaxY = mDividerMinY + 1

    mDividerMinX = 120
    If Me.WindowState <> vbMinimized Then
        mDividerMaxX = Me.ScaleWidth - 500
    End If
    
    If mDividerMaxX <= mDividerMinX Then mDividerMaxX = mDividerMinY + 1
    
End Sub

Private Sub UpdateCurrentDisplayOptions()
    
    If Not mFormLoaded Then Exit Sub
    
    With mGelDisplayListAndOptions.DisplayOptions
        .MassDisplayRangePPM = txtMassRange
        ValidateValueDbl .MassDisplayRangePPM, 0.000001, 20000, DEFAULT_ORF_MASS_RANGE_PPM
        
        .NETDisplayRange = txtNETRange
        ValidateValueDbl .NETDisplayRange, 0, 2, DEFAULT_ORF_NET_RANGE
        
        .DataDisplayMode = cboDataDisplayMode.ListIndex
            
    End With
    
    UpdateSavedGelListAndOptions mGelDisplayListAndOptions
    
    UpdateORFPics

End Sub

Public Function UpdateGelDisplayList(Optional lngGelIndexAssureIncluded As Long = -1) As Boolean
    ' Updates mGelDisplayListAndOptions, returning true if the ORFGroupList was populated
    ' if lngGelIndexAssureIncluded >=0 then sets .IncludeGel to True for gel with index lngGelIndexAssureIncluded
    
    Dim blnUpdateORFGroupList As Boolean
            
    blnUpdateORFGroupList = InitializeGelDisplayListAndOptions(mGelDisplayListAndOptions, lngGelIndexAssureIncluded)
    
    PopulateSourceGelComboBox
    PopulateColorKeyListView
        
    ' Update GelORFViewerSavedGelListAndOptions() with the latest GelList and options for all included gels
    UpdateSavedGelListAndOptions mGelDisplayListAndOptions
    
    If blnUpdateORFGroupList Then PopulateORFGroupList
        
    UpdateGelDisplayList = blnUpdateORFGroupList
    
End Function

Public Sub UpdateGelDisplayListAndOptions(blnPopulateORFGroupList As Boolean)
    ' This sub should only be called from the ORFViewerLoaderClass
    
    Dim blnORFGroupListUpdated As Boolean
    Dim lngORFDataIndexSelected As Long
    
    mGelDisplayListAndOptions = gOrfViewerOptionsCurrentGelList
    
    blnORFGroupListUpdated = UpdateGelDisplayList()
    
    If Not blnORFGroupListUpdated Then
        If blnPopulateORFGroupList Then
            PopulateORFGroupList
        Else
            lngORFDataIndexSelected = GetSelectedORFDataIndex()
            If lngORFDataIndexSelected >= 0 Then
                PopulateMassTagsListView lngORFDataIndexSelected
            End If
        End If
    End If
    
    UpdatePictureSizeAndSpacing
    
End Sub

Private Sub SortDataToPlot(ByRef udtData As udtDataToPopulateType, objQSDouble As QSDouble)
    Dim lngPointerArray() As Long, lngPointerIndex As Long
    
    Dim udtOldData As udtDataToPopulateType
    
    If udtData.Count > 1 Then
        ' Copy data from udtData to udtOldData
        udtOldData = udtData
        
        With udtOldData
        
            ' Need to ReDim the data so that we can use the QSDouble function to sort by increasing intensity
            InitializeDataToPopulate udtOldData, udtOldData.Count - 1, True

            ' Construct a pointer array
            ReDim lngPointerArray(0 To udtOldData.Count - 1)
            
            For lngPointerIndex = 0 To .Count - 1
                lngPointerArray(lngPointerIndex) = lngPointerIndex
            Next lngPointerIndex
            
            ' Sort the intensity array
            objQSDouble.QSDesc .Intensity(), lngPointerArray()
            
            ' Copy old data into udtData
            udtData.Count = .Count
            For lngPointerIndex = 0 To .Count - 1
                udtData.Labels(lngPointerIndex) = .Labels(lngPointerArray(lngPointerIndex))
                udtData.NET(lngPointerIndex) = .NET(lngPointerArray(lngPointerIndex))
                udtData.NETExtentNeg(lngPointerIndex) = .NETExtentNeg(lngPointerArray(lngPointerIndex))
                udtData.NETExtentPos(lngPointerIndex) = .NETExtentPos(lngPointerArray(lngPointerIndex))
                udtData.Mass(lngPointerIndex) = .Mass(lngPointerArray(lngPointerIndex))
                udtData.MassExtentNeg(lngPointerIndex) = .MassExtentNeg(lngPointerArray(lngPointerIndex))
                udtData.MassExtentPos(lngPointerIndex) = .MassExtentPos(lngPointerArray(lngPointerIndex))
                udtData.Intensity(lngPointerIndex) = .Intensity(lngPointerIndex)
            Next lngPointerIndex
            
        End With
        
        
    End If
End Sub

Private Sub SortListViewWrapper(ByRef lvwThisListView As MSComctlLib.ListView, ByVal lngColumnIndex As Long, ByRef udtColumnSortFormat As udtColumnSortFormatType, ByVal eListViewID As lviORFFormListViewIDConstants)
    ' Wrapper function to call the SortListView Function
    ' Calls PopulateMassTagsListViewCustomSort if called from lvwMassTags and user clicked on lvmTrypticStatus
    
    If eListViewID = lviMassTags And lngColumnIndex = lvmTrypticStatus Then
        PopulateMassTagsListViewCustomSortColumn
        lngColumnIndex = lvmSortKey
    End If

    ListViewSort lvwThisListView, lngColumnIndex, udtColumnSortFormat

    If eListViewID = lviMassTags Then
        UpdateORFPics mMassTagListViewORFIndex
    End If
    
End Sub

Private Sub UpdateAxisLabelFormattingStrings(dblMassRange As Double, dblNETRange As Double)
    ' Updates mXAxisFormatLabel and mYAxisFormatLabel based on the Mass and NET ranges
    ' For a mass range of 0.5 Da, divides by 10 to give 0.05 Da, and generates a
    '   format string of "0.00", thus showing 2 digits of precision
    
    If mGelDisplayListAndOptions.DisplayOptions.SwapPlottingAxes Then
        ' Mass
        mXAxisFormatLabel = ConstructFormatString(dblMassRange / 10)
        
        ' NET
        mYAxisFormatLabel = ConstructFormatString(dblNETRange / 10)
    Else
        ' NET
        mXAxisFormatLabel = ConstructFormatString(dblMassRange / 10)
        
        ' Mass
        mYAxisFormatLabel = ConstructFormatString(dblNETRange / 10)
    End If
    
End Sub

Private Sub UpdatePictureSizeAndSpacing()

    Dim lngNewPicHeight As Long, lngNewPicWidth As Long, lngNewPicSpacing As Long

    ' See if we need to update the picture size or spacing
    With mGelDisplayListAndOptions.DisplayOptions
        lngNewPicHeight = .PicturePixelHeight * Screen.TwipsPerPixelY
        lngNewPicWidth = .PicturePixelWidth * Screen.TwipsPerPixelX
        lngNewPicSpacing = .PicturePixelSpacing
        
        If lngNewPicHeight <> mPicHeight Or lngNewPicWidth <> mPicWidth Or lngNewPicSpacing <> mPicSpacing Then
            SetPictureDimensionsAndSpacing lngNewPicSpacing, lngNewPicHeight, lngNewPicWidth
            PositionControls
            UpdateORFPics
        End If
        
        ' Update the axis label formatting strings based on .MassRange and .NETRange
        ' Need to convert MassRange from PPM to daltons
        ' We'll simply use 1500 Da as a general mass to get an estimate of the mass range
        
        UpdateAxisLabelFormattingStrings PPMToMass(.MassDisplayRangePPM, 1500), .NETDisplayRange
        
    End With

End Sub

Private Sub UpdateStatus(ByVal strNewStatus As String)
    lblStatus = strNewStatus
    DoEvents
End Sub

Private Sub UpdateStatusORFDisplayCount(Optional intMTCountTotal As Integer = 0, Optional intHiddentMTPicCount As Integer = 0)
    ' Note: Call this sub with intMTCountTotal set to -1 to simply display the number of ORF's displayed
    Dim strStatus As String
    Dim strTags As String
    
    strStatus = Trim(Str(lvwORFs.ListItems.Count)) & " ORF's displayed"
    
    If intMTCountTotal = intHiddentMTPicCount And intMTCountTotal > 0 Then
        strStatus = strStatus & "; all " & Trim(Str(intMTCountTotal)) & " mass tags for selected ORF are hidden (no peaks in range)"
    Else
        If intMTCountTotal >= 0 Then
            If intMTCountTotal = 1 Then strTags = "tag" Else strTags = "tags"
            strStatus = strStatus & "; " & Trim(Str(intMTCountTotal)) & " mass " & strTags & " for selected ORF"
        End If
        
        If intHiddentMTPicCount > 0 Then
            If intHiddentMTPicCount = 1 Then strTags = "tag" Else strTags = "tags"
            strStatus = strStatus & " (" & Trim(Str(intHiddentMTPicCount)) & " mass " & strTags & " hidden since no peaks in range)"
        End If
    End If
    
    
    UpdateStatus strStatus
End Sub

Private Function VerifyORFsLoaded(blnForceReload As Boolean, Optional blnConfirmReloadIfExistingORFs As Boolean = True, Optional blnForceIonMatchRecount As Boolean = False, Optional blnForceORFGroupListUpdate As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------
    ' Check if necessary to load ORFs from Mass Tags database; optionally, force a reload
    ' Assumes GelAnalysis().MTDB.cn.ConnectionString contains the
    '  correct connection string to the database
    '
    ' Returns True if ORFs are present or were succesfully loaded
    ' Returns False if a failure, notifying user of failure if blnInformUserOnFail is True
    '------------------------------------------------------------------------------------
    
    Const ORF_DIM_CHUNK = 100
    
    Dim eResponse As VbMsgBoxResult
    Dim lngLoadedORFCount As Long
    Dim lngGelIndex As Long
    Dim lngGelIncludeCount As Long
    Dim blnContinueWithLoad As Boolean, blnSuccessfulLoad As Boolean, blnOnlyLoadTheoretical As Boolean
    Dim strMTDBConnectionString As String
    Dim blnNeedToPopulateORFListView As Boolean
    Dim blnCopiedValuesFromOtherGel As Boolean
    Dim intCopiedValuesCount As Integer
    
    If mGelDisplayListAndOptions.GelCount <= 0 Then
        MsgBox "No Gels found in memory.  Unable to load ORF's for them"
        blnSuccessfulLoad = False
    Else
        For lngGelIndex = 1 To mGelDisplayListAndOptions.GelCount
            If mGelDisplayListAndOptions.Gels(lngGelIndex).IncludeGel Then
                
                lngGelIncludeCount = lngGelIncludeCount + 1
                If Not GelAnalysis(lngGelIndex) Is Nothing Then
                   
                    strMTDBConnectionString = GelAnalysis(lngGelIndex).MTDB.cn.ConnectionString
            
                    lngLoadedORFCount = GelORFData(lngGelIndex).ORFCount
                    blnContinueWithLoad = False
                    If blnForceReload Or lngLoadedORFCount <= 0 Or _
                       (mGelDisplayListAndOptions.DisplayOptions.LoadPMTs And Not GelORFMassTags(lngGelIndex).Definition.IncludesPMTs) Then
                        blnContinueWithLoad = True
                        blnOnlyLoadTheoretical = False
                    ElseIf (mGelDisplayListAndOptions.DisplayOptions.IncludeUnobservedTrypticMassTags And Not GelORFMassTags(lngGelIndex).Definition.IncludesTheoreticalTrypticMassTags) Then
                        blnContinueWithLoad = True
                        blnOnlyLoadTheoretical = True
                    End If
                    
                    If blnContinueWithLoad Then
                        If lngLoadedORFCount > 0 And blnConfirmReloadIfExistingORFs Then
                            eResponse = MsgBox(Trim(Str(lngLoadedORFCount)) & " ORF's are already in memory for file " & StripFullPath(GelBody(lngGelIndex).Caption) & "; Re-load ORF's from the mass tag database?", vbQuestion + vbYesNoCancel + vbDefaultButton1, "Existing data")
                            blnContinueWithLoad = (eResponse = vbYes)
                        End If
                    End If
                    
                    If blnContinueWithLoad Then
                        UpdateStatus "Loading Protein data ..."
                        
                        If Not blnOnlyLoadTheoretical Then
                            intCopiedValuesCount = 0
                            
                            LoadORFsFromMTDB strMTDBConnectionString, lngGelIndex, blnCopiedValuesFromOtherGel
                            If blnCopiedValuesFromOtherGel Then intCopiedValuesCount = intCopiedValuesCount + 1
                            
                            LoadMassTagsForORFSFromMTDB strMTDBConnectionString, lngGelIndex, mGelDisplayListAndOptions.DisplayOptions.LoadPMTs, blnCopiedValuesFromOtherGel
                            If blnCopiedValuesFromOtherGel Then intCopiedValuesCount = intCopiedValuesCount + 1
                            
                            If intCopiedValuesCount < 2 Or Not GelORFData(lngGelIndex).Definition.DataParsedCompletely Or Not GelORFMassTags(lngGelIndex).Definition.DataParsedCompletely Then
                                ' Compute the number of tryptic peptides in each ORF
                                ' In addition, compute the start and stop residues for each tryptic mass tag
                                ' Can skip this step if the ORFs and mass tags were copied from an existing gel
                                UpdateORFStatistics lngGelIndex
                            End If
                        End If
                                            
                        ' If IncludeUnobservedTrypticMassTags = True, then examine the mass tags and
                        '  add missing tryptic mass tags to GelORFMassTags()
                        '  The new mass tags will have a theoretical NET value
                        If mGelDisplayListAndOptions.DisplayOptions.IncludeUnobservedTrypticMassTags Then
                            
                            If intCopiedValuesCount >= 2 And GelORFMassTags(lngGelIndex).Definition.IncludesTheoreticalTrypticMassTags Then
                                ' No need to re-compute theoretical tryptic mass tags since the mass tags were copied from
                                ' another gel, and that gel included theoretical tryptic mass tags
                            Else
                                ComputeTheoreticalTrypticMassTags GelORFData(lngGelIndex), GelORFMassTags(lngGelIndex), lngGelIndex
                                mGelDisplayListAndOptions.DisplayOptions.IncludeUnobservedTrypticMassTags = GelORFMassTags(lngGelIndex).Definition.IncludesTheoreticalTrypticMassTags
                            End If
                        End If
                    
                        ' Examine the Ions and UMC's for this gel and record the AMT matches in GelORFMassTags()
                        RecordIonMatchesInORFMassTags lngGelIndex
                    
                        GelStatus(lngGelIndex).Dirty = True
                    Else
                        ' May need to reserve memory for GelORFData(lngGelIndex) and GelORFMassTags(lngGelIndex)
                        If GelORFData(lngGelIndex).ORFCount = 0 Then
                            ReDim GelORFData(lngGelIndex).Orfs(ORF_DIM_CHUNK)
                            
                            ' Must re-initialize GelOrfMassTags() too; set .ORFCount to 0 to trigger this below
                            GelORFMassTags(lngGelIndex).ORFCount = 0
                        End If
                        
                        If GelORFMassTags(lngGelIndex).ORFCount = 0 Then
                            GelORFMassTags(lngGelIndex).ORFCount = GelORFData(lngGelIndex).ORFCount
                            ReDim GelORFMassTags(lngGelIndex).Orfs(UBound(GelORFData(lngGelIndex).Orfs()))
                        End If
                        
                        If blnForceIonMatchRecount Or GelUMC(lngGelIndex).UMCCnt <> CLngSafe(LookupValueInStringByKey(GelORFData(lngGelIndex).Definition.OtherInfo, UMC_COUNT_LAST_RECORD_ION_MATCH_CALL)) Then
                            RecordIonMatchesInORFMassTags lngGelIndex
                            GelStatus(lngGelIndex).Dirty = True
                        End If
                    End If
                    
                    blnNeedToPopulateORFListView = True
                    blnSuccessfulLoad = True
                    
                    UpdateStatus "Data loaded"
                Else
                    MsgBox "The connection to the Mass Tag database has not yet been defined for gel " & StripFullPath(GelBody(lngGelIndex).Caption) & vbCrLf & "Please define it using Edit -> Select/Modify Database Connection on the main plot window.  After this, choose Reload ORFs from the File menu of this window."
                    blnSuccessfulLoad = False
                End If
            End If
        Next lngGelIndex
    End If

    If blnNeedToPopulateORFListView Then
        If blnForceORFGroupListUpdate Then
            PopulateORFGroupList False, False
        End If
        
        PopulateORFListView
    Else
        If (lngGelIncludeCount = 0 Or mORFViewerGroupList.ORFCount = 0) And lvwORFs.ListItems.Count > 0 Then
            PopulateORFListView
        End If
    End If

    VerifyORFsLoaded = blnSuccessfulLoad

End Function

Private Sub ZoomOutGelBody(lngGelIndex As Long)
    If GelBody(lngGelIndex).csMyCooSys.csZmLvl > 1 Then
        GelBody(lngGelIndex).csMyCooSys.ZoomOut
    End If

End Sub

Private Sub ZoomParentGelToRegionOfSelectedMassTag()
    Const ZOOM_SUBLEVEL_COUNT = 3
    
    Dim lngZoomLevelScanNum(ZOOM_SUBLEVEL_COUNT, 2) As Long
    Dim dblZoomLevelMass(ZOOM_SUBLEVEL_COUNT, 2) As Double
    Dim intZoomIndex As Integer
    
    Dim lngSourceGelIndex As Long
    Dim lngSelectedMassTagIndex As Long
    Dim dblNETRangeHalf As Double, dblMassRangeDaltons As Double
    
    Dim lngZoomedScanNumberMin As Long, lngZoomedScanNumberMax As Long
    Dim dblZoomedMassMin As Double, dblZoomedMassMax As Double
    
    Dim lngGelScanNumberMin As Long, lngGelScanNumberMax As Long
    Dim dblGelMassMin As Double, dblGelMassMax As Double
    
    Dim lngScanRangeIncrementDown As Long, lngScanRangeIncrementUp As Long
    Dim dblMassRangeIncrementDown As Double, dblMassRangeIncrementUp As Double
    
    
    ' Find the first selected mass tag and verify that mMassTagListViewORFIndex is valid
    lngSelectedMassTagIndex = GetSelectedMassTagDataIndex()
    If lngSelectedMassTagIndex < 0 Or mMassTagListViewORFIndex < 0 Then Exit Sub
    
    ' Determine the desired parent gel index
    lngSourceGelIndex = GetDesiredParentGelIndex()
    If lngSourceGelIndex < 1 Then Exit Sub
    
    ' Determine the NET value and mass of the selected mass tag
    With GelORFMassTags(lngSourceGelIndex).Orfs(mMassTagListViewORFIndex).MassTags(lngSelectedMassTagIndex)
        
        ' Compute the range values, divided by 2
        dblNETRangeHalf = mGelDisplayListAndOptions.DisplayOptions.NETDisplayRange / 2
        dblMassRangeDaltons = PPMToMass(mGelDisplayListAndOptions.DisplayOptions.MassDisplayRangePPM, .Mass) / 2
    
        ' Compute the scan number range for the OrfPics
        lngZoomedScanNumberMin = GANETToScan(lngSourceGelIndex, .GANET - dblNETRangeHalf)
        lngZoomedScanNumberMax = GANETToScan(lngSourceGelIndex, .GANET + dblNETRangeHalf)
        
        dblZoomedMassMin = .Mass - dblMassRangeDaltons
        dblZoomedMassMax = .Mass + dblMassRangeDaltons
        
        If Abs(lngZoomedScanNumberMax - lngZoomedScanNumberMin) < 2 Then
            lngZoomedScanNumberMin = lngZoomedScanNumberMin - 1
            lngZoomedScanNumberMax = lngZoomedScanNumberMax + 1
        End If
    End With
    
    If lngZoomedScanNumberMin > 0 And dblZoomedMassMin > 0 Then

        ' Zoom in 4 levels to allow the user the chance to zoom out
        ' First zoom out full
        ZoomOutGelBody lngSourceGelIndex
        
        ' Determine the current scan number range
        GetScanRange lngSourceGelIndex, lngGelScanNumberMin, lngGelScanNumberMax, 0
        
        lngScanRangeIncrementDown = (lngZoomedScanNumberMin - lngGelScanNumberMin) / ZOOM_SUBLEVEL_COUNT
        lngScanRangeIncrementUp = (lngGelScanNumberMax - lngZoomedScanNumberMax) / ZOOM_SUBLEVEL_COUNT
        
        GetMassRangeCurrent lngSourceGelIndex, dblGelMassMin, dblGelMassMax
        
        dblMassRangeIncrementDown = (dblZoomedMassMin - dblGelMassMin) / ZOOM_SUBLEVEL_COUNT
        dblMassRangeIncrementUp = (dblGelMassMax - dblZoomedMassMax) / ZOOM_SUBLEVEL_COUNT
        
        lngZoomLevelScanNum(0, 0) = lngZoomedScanNumberMin
        lngZoomLevelScanNum(0, 1) = lngZoomedScanNumberMax
        
        dblZoomLevelMass(0, 0) = dblZoomedMassMin
        dblZoomLevelMass(0, 1) = dblZoomedMassMax
        
        For intZoomIndex = 1 To ZOOM_SUBLEVEL_COUNT - 1
            lngZoomLevelScanNum(intZoomIndex, 0) = lngZoomLevelScanNum(intZoomIndex - 1, 0) - lngScanRangeIncrementDown
            lngZoomLevelScanNum(intZoomIndex, 1) = lngZoomLevelScanNum(intZoomIndex - 1, 1) + lngScanRangeIncrementUp
            
            dblZoomLevelMass(intZoomIndex, 0) = dblZoomLevelMass(intZoomIndex - 1, 0) - dblMassRangeIncrementDown
            dblZoomLevelMass(intZoomIndex, 1) = dblZoomLevelMass(intZoomIndex - 1, 1) + dblMassRangeIncrementUp
        Next intZoomIndex
        
        For intZoomIndex = ZOOM_SUBLEVEL_COUNT - 1 To 0 Step -1
            ZoomGelToDimensions lngSourceGelIndex, CSng(lngZoomLevelScanNum(intZoomIndex, 0)), dblZoomLevelMass(intZoomIndex, 0), CSng(lngZoomLevelScanNum(intZoomIndex, 1)), dblZoomLevelMass(intZoomIndex, 1)
        Next intZoomIndex
    Else
        MsgBox "An invalid scan range has been computed for the given mass tag.  Zoom aborted.", vbInformation + vbOKOnly, "Unable to Zoom"
    End If
    
    
End Sub

Private Sub cboDataDisplayMode_Click()
    UpdateCurrentDisplayOptions
End Sub

Private Sub cboItemCountSourceGel_Click()
    
    If mPopulatingSourceGelCombo Then Exit Sub
    
    PopulateORFListView
    
    UpdateSavedGelListAndOptions mGelDisplayListAndOptions
End Sub

Private Sub cmdHorizontalDivider_DragDrop(Source As Control, x As Single, y As Single)
    fraORFPicsContainer.Visible = True
    PositionControls False, True
End Sub

Private Sub cmdHorizontalDivider_DragOver(Source As Control, x As Single, y As Single, STATE As Integer)
    CheckMoveDividerBars x + cmdHorizontalDivider.Left, y + cmdHorizontalDivider.Top, False
End Sub

Private Sub cmdHorizontalDivider_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mMovingHorizDivider = True
    mMovingVerticalDivider = False
    cmdHorizontalDivider.Drag vbBeginDrag
End Sub

Private Sub cmdRollUpShrink_Click()
    ShrinkExpandListViews False
End Sub

Private Sub cmdRollUpExpand_Click()
    ShrinkExpandListViews True
End Sub

Private Sub cmdVerticalDivider_DragDrop(Source As Control, x As Single, y As Single)
    fraORFPicsContainer.Visible = True
    PositionControls False, True
End Sub

Private Sub cmdVerticalDivider_DragOver(Source As Control, x As Single, y As Single, STATE As Integer)
    CheckMoveDividerBars x + cmdVerticalDivider.Left, y + cmdVerticalDivider.Top, False
End Sub

Private Sub cmdVerticalDivider_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mMovingHorizDivider = False
    mMovingVerticalDivider = True
    fraORFPicsContainer.Visible = False
    DoEvents
    cmdVerticalDivider.Drag vbBeginDrag
End Sub

Private Sub Form_Activate()
    UpdateGelDisplayList
End Sub

Private Sub Form_DragDrop(Source As Control, x As Single, y As Single)
    CheckMoveDividerBars x, y, True
End Sub

Private Sub Form_DragOver(Source As Control, x As Single, y As Single, STATE As Integer)
    CheckMoveDividerBars x, y, False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Form-wide key handler
    
    Dim lngNewVScrollValue As Long
    
    If Shift = vbAltMask Then
        ' Check for Alt+Left Arrow and Alt+Right Arrow
        If KeyCode = 37 Then
            ORFHistoryNavigate False
        ElseIf KeyCode = 39 Then
            ORFHistoryNavigate True
        End If
    ElseIf mKeyPressAbortORFListPopulate > 0 Then
        ' check for Esc key when populating the ORF listview
        If KeyCode = vbKeyEscape Then mKeyPressAbortORFListPopulate = 2
    Else
        ' Look for up or down arrow, or page-up or page-down
        lngNewVScrollValue = -1
        Select Case KeyCode
        Case vbKeyPageUp
            lngNewVScrollValue = VScroll.value - VScroll.LargeChange
        Case vbKeyPageDown
            lngNewVScrollValue = VScroll.value + VScroll.LargeChange
        Case vbKeyUp
            lngNewVScrollValue = VScroll.value - 1
        Case vbKeyDown
            lngNewVScrollValue = VScroll.value + 1
        Case vbKeyHome
            lngNewVScrollValue = 0
        Case vbKeyEnd
            lngNewVScrollValue = VScroll.Max
        End Select
        
        If lngNewVScrollValue >= 0 Then
            If lngNewVScrollValue < VScroll.Min Then
                lngNewVScrollValue = VScroll.Min
            ElseIf lngNewVScrollValue > VScroll.Max Then
                lngNewVScrollValue = VScroll.Max
            End If
            VScroll.value = lngNewVScrollValue
        End If
    End If
    
End Sub

Private Sub Form_Load()
    ' Change the background color to white
    Me.BackColor = &HFFFFFF
    
    'fraORFPicsContainer.BackColor = Me.BackColor
    'fraORFPicsClippingRegion.BackColor = Me.BackColor
    
    ' We start out with one loaded ctlORF (though SetVisiblePictureCount below is used to hide it)
    mORFPicsLoadedCount = 1
    mORFPicsRowCount = 1
    mORFPicsMaxLoadedCount = 100
    
    cmdRollUpExpand.Visible = True
    cmdRollUpShrink.Visible = False
    mListViewsExpanded = False
    
    PopulateComboBoxes
    
    InitializeLocalVariables
    InitializeListViews
    
    mFormLoaded = True
    
    ' Make sure fraORFPicsContainer.Top is 0 to start with
    fraORFPicsContainer.Top = 0
    
    SetVisiblePictureCount 0, False

    InitializeGANET
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        ORFViewerLoader.HideORFViewerForm Me, mFormID
    End If
    
End Sub

Private Sub Form_Resize()
    PositionControls True, False
End Sub

Private Sub fraORFPicsClippingRegion_DragDrop(Source As Control, x As Single, y As Single)
    CheckMoveDividerBars fraORFPicsClippingRegion.Left + x, fraORFPicsClippingRegion.Top + y, True
End Sub

Private Sub fraORFPicsClippingRegion_DragOver(Source As Control, x As Single, y As Single, STATE As Integer)
    CheckMoveDividerBars fraORFPicsClippingRegion.Left + x, fraORFPicsClippingRegion.Top + y, False
End Sub

Private Sub fraORFPicsContainer_DragDrop(Source As Control, x As Single, y As Single)
    CheckMoveDividerBars fraORFPicsContainer.Left + x, fraORFPicsContainer.Top + y, True
End Sub

Private Sub fraORFPicsContainer_DragOver(Source As Control, x As Single, y As Single, STATE As Integer)
    CheckMoveDividerBars fraORFPicsContainer.Left + x, fraORFPicsContainer.Top + y, False
End Sub

Private Sub lvwColorKey_Click()
    Dim lngListIndex As Long
    
    ' Deselect item if user clicks on it
    For lngListIndex = 1 To lvwColorKey.ListItems.Count
        If lvwColorKey.ListItems(lngListIndex).Selected = True Then
            lvwColorKey.ListItems(lngListIndex).Selected = False
        End If
    Next lngListIndex
    
End Sub

Private Sub lvwMassTags_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ' Note: This will eventually call UpdateORFPics automatically
    HandleListViewColumnClick lvwMassTags, ColumnHeader, mColumnSortFormats(lviMassTags), lviMassTags
End Sub

Private Sub lvwMassTags_DragDrop(Source As Control, x As Single, y As Single)
    CheckMoveDividerBars x, y, True, True, lvwMassTags
End Sub

Private Sub lvwMassTags_DragOver(Source As Control, x As Single, y As Single, STATE As Integer)
    CheckMoveDividerBars x, y, False, True, lvwMassTags
End Sub

Private Sub lvwMassTags_KeyDown(KeyCode As Integer, Shift As Integer)
    ListViewKeyHandler lvwMassTags, lviMassTags, KeyCode, Shift, lvoSortkey - 1
End Sub

Private Sub lvwORFs_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    HandleListViewColumnClick lvwORFs, ColumnHeader, mColumnSortFormats(lviORFs), lviORFs
End Sub

Private Sub lvwORFs_DragDrop(Source As Control, x As Single, y As Single)
    CheckMoveDividerBars x, y, True, True, lvwORFs
End Sub

Private Sub lvwORFs_DragOver(Source As Control, x As Single, y As Single, STATE As Integer)
    CheckMoveDividerBars x, y, False, True, lvwORFs
End Sub

Private Sub lvwORFs_ItemClick(ByVal Item As MSComctlLib.ListItem)
    HandleListViewItemClick Item, lviORFs
End Sub

Private Sub lvwORFs_KeyDown(KeyCode As Integer, Shift As Integer)
    ListViewKeyHandler lvwORFs, lviORFs, KeyCode, Shift, lvoSortkey - 1
End Sub

Private Sub mnuClose_Click()
    ORFViewerLoader.HideORFViewerForm Me, mFormID
End Sub

Private Sub mnuCopy_Click()
    CopySelectedItems lvwORFs, lvoSortkey - 1, lviORFs
End Sub

Private Sub mnuFindORFContaingMassTag_Click()
    FindORFContainingMassTag True
End Sub

Private Sub mnuFindText_Click()
    FindTextInListViewWrapper lviORFs
End Sub

Private Sub mnuHighlightMassTagsForSelectedORF_Click()
    HighlightMassTagsForSelectedORF
End Sub

Private Sub mnuLoadNewORFViewerWindow_Click()
    ORFViewerLoader.ShowORFViewerForm "", -1, True
End Sub

Private Sub mnuORFHistoryListItem_Click(Index As Integer)
    ORFHistoryJump Index
End Sub

Private Sub mnuORFHistoryMoveNext_Click()
    ORFHistoryNavigate True
End Sub

Private Sub mnuORFHistoryMovePrevious_Click()
    ORFHistoryNavigate False
End Sub

Private Sub mnuRefreshORFList_Click()
    PopulateORFGroupList True, True
End Sub

Private Sub mnuReloadORFsFromMTDB_Click()
    VerifyORFsLoaded True, True, True, True
    UpdateGelDisplayList
End Sub

Private Sub mnuSelectAll_Click()
    ListViewSelectAllItems lvwORFs
End Sub

Private Sub mnuSetsGelsAndOptions_Click()
    ShowORFViewerOptions
End Sub

Private Sub mnuZoomParentGel_Click()
    ZoomParentGelToRegionOfSelectedMassTag
End Sub

Private Sub txtMassRange_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtMassRange_Validate (False)
    TextBoxKeyPressHandler txtMassRange, KeyAscii, True, True, False
End Sub

Private Sub txtMassRange_Validate(Cancel As Boolean)
    ValidateTextboxValueDbl txtMassRange, 0.000001, 20000, DEFAULT_ORF_MASS_RANGE_PPM
    UpdateCurrentDisplayOptions
End Sub

Private Sub txtNETRange_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtNETRange_Validate (False)
    TextBoxKeyPressHandler txtMassRange, KeyAscii, True, True, False
End Sub

Private Sub txtNETRange_Validate(Cancel As Boolean)
    ValidateTextboxValueDbl txtNETRange, 0.000001, 2, DEFAULT_ORF_NET_RANGE
    UpdateCurrentDisplayOptions
End Sub

Private Sub VSCroll_Change()
    HandleVerticalScroll
End Sub

