VERSION 5.00
Begin VB.Form frmExcludeMassRange 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filter By Mass Range"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRestoreDefaultFilter 
      Caption         =   "Restore Default Filter Ion Visibility"
      Height          =   495
      Left            =   6600
      TabIndex        =   37
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Frame fraAutoPopulateSearchScope 
      Caption         =   "Search Scope"
      Height          =   975
      Left            =   6240
      TabIndex        =   30
      Top             =   3360
      Width           =   2295
      Begin VB.OptionButton optAutoPopulateSearchScope 
         Caption         =   "&Current View"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   32
         Top             =   600
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optAutoPopulateSearchScope 
         Caption         =   "&All Data Points"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   31
         Top             =   280
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdInclude 
      Caption         =   "&Include Listed Ions"
      Height          =   375
      Left            =   6600
      TabIndex        =   34
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton cmdExclude 
      Caption         =   "&Exclude Listed Ions"
      Height          =   375
      Left            =   6600
      TabIndex        =   33
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox txtDefaultMassSliceWidthPPM 
      Height          =   285
      Left            =   3600
      TabIndex        =   5
      Text            =   "2"
      Top             =   2880
      Width           =   615
   End
   Begin VB.Frame fraAutoPopulate 
      Caption         =   "Auto-Populate Exclusion List"
      Height          =   3255
      Left            =   240
      TabIndex        =   7
      Top             =   3360
      Width           =   5775
      Begin VB.CommandButton cmdAutoPopulate 
         Caption         =   "&Auto-Populate"
         Height          =   375
         Left            =   4320
         TabIndex        =   28
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton cmdAutoPopulateCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4320
         TabIndex        =   29
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtAutoPopulateLimitMassRangeEnd 
         Height          =   285
         Left            =   3240
         TabIndex        =   27
         Text            =   "6000"
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox txtAutoPopulateLimitMassRangeStart 
         Height          =   285
         Left            =   3240
         TabIndex        =   25
         Text            =   "100"
         Top             =   2400
         Width           =   615
      End
      Begin VB.CheckBox chkAutoPopulateLimitMassRange 
         Caption         =   "Limit Mass Range"
         Height          =   255
         Left            =   2280
         TabIndex        =   23
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CheckBox chkAutoPopulateRequireIdenticalCharge 
         Caption         =   "Require Identical Charge for matching ions"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin VB.TextBox txtAutoPopulateMassSliceWidthPPM 
         Height          =   285
         Left            =   3360
         TabIndex        =   15
         Text            =   "2"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtAutoPopulateNeighborCountPercentageThreshold 
         Height          =   285
         Left            =   3360
         TabIndex        =   12
         Text            =   "50"
         ToolTipText     =   "Defines the minimum scan occurrence percentage for the neighbor slice for it to be added to the search slice"
         Top             =   780
         Width           =   615
      End
      Begin VB.TextBox txtAutoPopulateLimitScanRangeEnd 
         Height          =   285
         Left            =   1080
         TabIndex        =   22
         Text            =   "5000"
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox txtAutoPopulateLimitScanRangeStart 
         Height          =   285
         Left            =   1080
         TabIndex        =   20
         Text            =   "1"
         Top             =   2400
         Width           =   615
      End
      Begin VB.CheckBox chkAutoPopulateLimitScanRange 
         Caption         =   "Limit Scan Range"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtAutoPopulateCountPercentageThresholdForExclusion 
         Height          =   285
         Left            =   3360
         TabIndex        =   9
         Text            =   "50"
         ToolTipText     =   "Defines the mininum scan occurrence percentage to include in the exclusion list"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblAutoPopulateLimitMassRangeEnd 
         Caption         =   "End Mass"
         Height          =   255
         Left            =   2280
         TabIndex        =   26
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label lblAutoPopulateLimitMassRangeStart 
         Caption         =   "Start Mass"
         Height          =   255
         Left            =   2280
         TabIndex        =   24
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label lblAutoPopulateMassSliceWidthPPMUnits 
         Caption         =   "ppm"
         Height          =   255
         Left            =   4080
         TabIndex        =   16
         Top             =   1230
         Width           =   495
      End
      Begin VB.Label lblAutoPopulateMassSliceWidthPPM 
         Caption         =   "Mass slice width (± search mass)"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label lblAutoPopulateNeighborCountPercentageThresholdUnits 
         Caption         =   "% of scan range"
         Height          =   255
         Left            =   4080
         TabIndex        =   13
         Top             =   810
         Width           =   1455
      End
      Begin VB.Label lblAutoPopulateNeighborCountPercentageThreshold 
         Caption         =   "Neighbor slice scan threshold for addition to search slice"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label lblAutoPopulateLimitScanRangeEnd 
         Caption         =   "End Scan"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label lblAutoPopulateLimitScanRangeStart 
         Caption         =   "Start Scan"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label lblAutoPopulateCountPercentageThresholdForExclusionUnits 
         Caption         =   "% of scan range"
         Height          =   255
         Left            =   4080
         TabIndex        =   10
         Top             =   390
         Width           =   1215
      End
      Begin VB.Label lblAutoPopulateCountPercentageThresholdForExclusion 
         Caption         =   "Search slice threshold for exclusion"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.TextBox txtExclusionList 
      Height          =   1935
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   840
      Width           =   4215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6960
      TabIndex        =   35
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   240
      TabIndex        =   36
      Top             =   6720
      Width           =   5775
   End
   Begin VB.Label lblExclusionListHeader 
      Caption         =   "Monoisotopic mass, ppm tol, charge, scan start, scan end"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label lblDefaultMassSliceWidthPPMUnits 
      Caption         =   "ppm"
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lblDefaultMassSliceWidthPPM 
      Caption         =   "Default mass slice tolerance (± search mass)"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Label lblDirections 
      Caption         =   "Directions"
      Height          =   2415
      Left            =   4920
      TabIndex        =   3
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label lblExclusionList 
      Caption         =   "Exclusion List"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmExcludeMassRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type udtAutoPopulateStatsType
    SearchMass As Double            ' Could be the central search ion mass, or a neighbor search ion mass
    AbsMassTol As Double            ' half window tolerance, +- dblSearchIonMass
    Charge As Integer
    
    LimitScanRange As Boolean
    ScanRangeStart As Long
    ScanRangeEnd As Long
    ScanRangeCount As Long
    
    LimitMassRange As Boolean
    MassStart As Double
    MassEnd As Double
    
    PercentScansInUseThreshold As Single
    RequireIdenticalCharge As Boolean
End Type

' Indexing schema modelled after that used in frmUMCSimple, written by Nikola Tolic in Spring 2003
' All of the arrays are 0-based

Private CSCnt As Long               'count of CS data points included in count
Private ISCnt As Long               'count of IS data points included in count

' The O_ arrays contain the data for the ions to be searched
' We could get by with just the O_Index() and O_Type() arrays, but then we'd have to
'  continually be looking up values in GelData().CSLines and GelData().IsoLines
' Copying the values to these arrays speeds up the search
'
Private O_Cnt As Long               'total number of ions to search
Private O_Index() As Long           'index in CS/Iso arrays
Private O_Type() As glDistType      'type of data(CS/Iso)
Private O_MW() As Double            'mass array
Private O_Charge() As Double        'charge
Private O_Order() As Double         'Abundance
Private O_Scan() As Long            'scan number

Private S_MW() As Double            'sorted mass array used for fast search

Private IndMW() As Long             'index on mass
Private IndOrder() As Long          'index on order (ordered by decreasing abundance)

Private IsMatched() As Boolean      'is already matched

Private MWRangeFinder As MWUtil     'fast search of mass range

Private CallerID As Long
Private mCancelOperation As Boolean
'

Private Sub AddToExclusionList(udtNewEntry As udtExclusionIonType, Optional blnAllowMergingWithExistingEntries As Boolean = True, Optional blnRePopulateExclusionListArray As Boolean = True, Optional blnUpdateExclusionListTextbox As Boolean = False)

    Dim lngindex As Long
    Dim dblAbsTolerance As Double, dblAbsToleranceNewEntry As Double
    
    Dim dblExcludeIonStartMass As Double, dblExcludeIonEndMass As Double
    Dim dblNewEntryStartMass As Double, dblNewEntryEndMass As Double
    Dim blnEntriesMerged As Boolean
    
    If blnRePopulateExclusionListArray Then
        FillExclusionListFromTextbox False
    End If

    ' Look through the exclusion list to see if any of the current entries overlaps with the
    '   new entry by mass
    ' If they do, and if they have the same charge and scan numbers, then combine them
    
    If blnAllowMergingWithExistingEntries And glbPreferencesExpanded.NoiseRemovalOptions.ExclusionListCount > 0 Then
        With udtNewEntry
            dblAbsToleranceNewEntry = PPMToMass(.TolerancePPM, .IonMass)
            dblNewEntryStartMass = .IonMass - dblAbsToleranceNewEntry
            dblNewEntryEndMass = .IonMass + dblAbsToleranceNewEntry
        End With
        
        With glbPreferencesExpanded.NoiseRemovalOptions
            For lngindex = 0 To .ExclusionListCount - 1
                With .ExclusionList(lngindex)
                    dblAbsTolerance = PPMToMass(.TolerancePPM, .IonMass)
                    
                    dblExcludeIonStartMass = .IonMass - dblAbsTolerance
                    dblExcludeIonEndMass = .IonMass + dblAbsTolerance
                    
                    If dblNewEntryEndMass >= dblExcludeIonStartMass And dblNewEntryEndMass <= dblExcludeIonEndMass Or _
                       dblNewEntryStartMass >= dblExcludeIonStartMass And dblNewEntryStartMass <= dblExcludeIonEndMass Then
                       
                        ' Yes, they do overlap; merge them together if their charges match and the scan limits match
                        If .Charge = udtNewEntry.Charge And .LimitScanRange = udtNewEntry.LimitScanRange Then
                            If Not .LimitScanRange Or (.ScanStart = udtNewEntry.ScanStart And .ScanEnd = udtNewEntry.ScanEnd) Then
                                ' Yes, all agree
                                ' Merge together
                                
                                If dblNewEntryStartMass < dblExcludeIonStartMass Then dblExcludeIonStartMass = dblNewEntryStartMass
                                If dblNewEntryEndMass > dblExcludeIonEndMass Then dblExcludeIonEndMass = dblNewEntryEndMass
                                
                                ' Average the start and end masses to get the central mass
                                .IonMass = Round((dblExcludeIonStartMass + dblExcludeIonEndMass) / 2, 6)
                                
                                ' The new PPM tolerance is 50% of the mass range, converted to ppm
                                .TolerancePPM = Round(MassToPPM((dblExcludeIonEndMass - dblExcludeIonStartMass) / 2, .IonMass), 2)
                                
                                blnEntriesMerged = True
                            End If
                        End If
                    End If
                End With
                If blnEntriesMerged Then Exit For
            Next lngindex
        End With
    End If
    
    If Not blnEntriesMerged Then
        ' Add an entry to the exclusion list
        With glbPreferencesExpanded.NoiseRemovalOptions
            .ExclusionListCount = .ExclusionListCount + 1
            ReDim Preserve .ExclusionList(0 To .ExclusionListCount - 1)
        
            .ExclusionList(.ExclusionListCount - 1) = udtNewEntry
        End With
    End If
    
    If blnUpdateExclusionListTextbox Then
        FillExclusionTextBox
    End If
End Sub

Private Function AutoPopulateFindPercentScansInUse(udtAutoPopulateStats As udtAutoPopulateStatsType, Optional ByRef MWRangeMinInd As Long, Optional ByRef MWRangeMaxInd As Long, Optional ByVal blnMarkIonsIfOverThreshold As Boolean = True, Optional ByVal blnRigorousComputation As Boolean = False) As Boolean
    ' Returns True if the Search Mass is found to have a PercentScansInUse value >= sngPercentScansInUseThreshold
    ' Returns False otherwise
    ' Returns the percent scans in use in the ByRef variable sngPercentScansInUse (a number between 0.0 and 100.0)
    ' If blnRigorousComputation = False, then sngPercentScansInUse could contain the maximum possible
    '   percent scans in use; use blnRigorousComputation = True to obtain the accurate value (which will
    '   always be at most equal to the approximate value)
    
    Dim sngPercentScansInUse As Single
    
    Dim ScanUsed() As Boolean
    Dim lngScansInUseCount As Long
    
    Dim blnUseMatchIon As Boolean
    Dim lngindex As Long
    
    MWRangeMinInd = 0
    MWRangeMaxInd = O_Cnt - 1
    With udtAutoPopulateStats
        If MWRangeFinder.FindIndexRange(.SearchMass, .AbsMassTol, MWRangeMinInd, MWRangeMaxInd) Then
            If MWRangeMaxInd >= MWRangeMinInd Then
            
                ' At most, the scans in use count will be MWRangeMaxInd - MWRangeMinInd + 1
                ' Make an initial check to see if this mass slice might pass the scan count threshold
                sngPercentScansInUse = (MWRangeMaxInd - MWRangeMinInd + 1) / CSng(.ScanRangeCount) * 100
                
                If sngPercentScansInUse >= .PercentScansInUseThreshold Or blnRigorousComputation Then
                
                    ReDim ScanUsed(.ScanRangeStart To .ScanRangeEnd)
                    lngScansInUseCount = 0
                
                    ' Step through the matching data
                    ' Set ScanUsed() to True for the scan of each matching data point
                    ' However, do not change ScanUsed() if limiting to identical charge
                    '   or if limiting scan range and out of range
                    blnUseMatchIon = True
                    For lngindex = MWRangeMinInd To MWRangeMaxInd
                        ' Possibly check for matching charge
                        If .RequireIdenticalCharge Then
                            blnUseMatchIon = (O_Charge(IndMW(lngindex)) = .Charge)
                        End If
                        
                        ' Possibly check for scan within range
                        If .LimitScanRange Then
                            If blnUseMatchIon Then
                                If O_Scan(IndMW(lngindex)) < .ScanRangeStart Or _
                                   O_Scan(IndMW(lngindex)) > .ScanRangeEnd Then
                                    blnUseMatchIon = False
                                End If
                            End If
                        End If
                        
                        If blnUseMatchIon Then
                            If Not ScanUsed(O_Scan(IndMW(lngindex))) Then
                                ScanUsed(O_Scan(IndMW(lngindex))) = True
                                lngScansInUseCount = lngScansInUseCount + 1
                            End If
                        End If
                    Next lngindex
                    
                    ' We can now compute the rigorous percent scans in use
                    sngPercentScansInUse = lngScansInUseCount / CSng(.ScanRangeCount) * 100
                
                End If
            Else
                sngPercentScansInUse = 0
            End If
        Else
            ' This shouldn't happen
            Debug.Assert False
            sngPercentScansInUse = 0
        End If
        
        If sngPercentScansInUse >= .PercentScansInUseThreshold Then
            If blnMarkIonsIfOverThreshold Then
                ' Mark the ions in this slice as matched
                For lngindex = MWRangeMinInd To MWRangeMaxInd
                    IsMatched(IndMW(lngindex)) = True
                Next lngindex
            End If
            
            AutoPopulateFindPercentScansInUse = True
        Else
            AutoPopulateFindPercentScansInUse = False
        End If
    End With
    
End Function

Public Function AutoPopulateStart(Optional blnClearIonExclusionList As Boolean = False) As Long
    ' Returns the number of ions added to txtExclusionList
    ' Returns -1 if cancelled prior to the Do-Loop or if an error
    
    Dim eResponse As VbMsgBoxResult
    Dim blnSuccess As Boolean
    Dim strErrorMessage As String
    Dim strHistoryText As String
    
    Dim lngIonsAddedCount As Long
    
On Error GoTo AutoPopulateStartErrorHandler

    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        ' Query user about clearing the exclusion list, but only if the
        '  list isn't empty and if not auto analyzing
        If Len(Trim(txtExclusionList.Text)) > 0 Then
            eResponse = MsgBox("Clear the exclusion list before auto-populating?", vbQuestion + vbYesNoCancel + vbDefaultButton1, "Clear List")
            If eResponse = vbYes Then
                txtExclusionList.Text = ""
            ElseIf eResponse = vbCancel Then
                AutoPopulateStart = -1
                Exit Function
            End If
        End If
    Else
        If blnClearIonExclusionList Then
            txtExclusionList.Text = ""
        End If
    End If
    
    mCancelOperation = False
    Me.MousePointer = vbHourglass
    cmdAutoPopulate.Visible = False
    txtExclusionList.Locked = True
    
    ' Initialize the search tolerances
    UpdateCurrentNoiseRemovalOptions
    
    blnSuccess = True
    With glbPreferencesExpanded.NoiseRemovalOptions
        If .LimitMassRange Then
            If .MassEnd < .MassStart Then
                blnSuccess = False
                lngIonsAddedCount = -1
                strErrorMessage = "End mass value is less than the start mass value"
            End If
        End If
        
        If .LimitScanRange Then
            If .ScanEnd < .ScanStart Then
                blnSuccess = False
                lngIonsAddedCount = -1
                strErrorMessage = "End scan value is less than the start scan value"
            End If
        End If
    End With
    
    If blnSuccess Then
        If InitializeSearchIndices(glbPreferencesExpanded.NoiseRemovalOptions.SearchScope) And Not mCancelOperation Then
            If Not CreateIndMW() Or mCancelOperation Then blnSuccess = False
            
            If blnSuccess Then
                If Not CreateIndOrder() Or mCancelOperation Then blnSuccess = False
            End If
            
            If blnSuccess Then
                lngIonsAddedCount = AutoPopulateWork(glbPreferencesExpanded.NoiseRemovalOptions)
            Else
                lngIonsAddedCount = -1
            End If
        Else
            blnSuccess = False
            lngIonsAddedCount = -1
        End If
    End If
    
    If blnSuccess Then
        Status "Auto-populate success. Added " & Trim(lngIonsAddedCount) & " ions"
        
        With glbPreferencesExpanded.NoiseRemovalOptions
            strHistoryText = "Auto-populated the mass exclusion search list: Entries added = " & Trim(lngIonsAddedCount)
            
            If .SearchScope = glScope.glSc_All Then
                strHistoryText = strHistoryText & "; Search Scope = All"
            Else
                strHistoryText = strHistoryText & "; Search Scope = Current View"
            End If
        
            strHistoryText = strHistoryText & "; Percentage Threshold To Exclude Slice = " & Trim(.PercentageThresholdToExcludeSlice)
            strHistoryText = strHistoryText & "; Percentage Threshold To Add Neighbor To Search Slice = " & Trim(.PercentageThresholdToAddNeighborToSearchSlice)
            strHistoryText = strHistoryText & "; Search Tolerance PPM = " & Trim(.SearchTolerancePPMAutoRemoval)
            strHistoryText = strHistoryText & "; Require Identical Charge = " & Trim(.RequireIdenticalCharge)
            
            strHistoryText = strHistoryText & "; Limit Scan Range = " & Trim(.LimitScanRange)
            If .LimitScanRange Then
                strHistoryText = strHistoryText & "; Scan Start = " & Trim(.ScanStart)
                strHistoryText = strHistoryText & "; Scan End = " & Trim(.ScanEnd)
            End If
            
            strHistoryText = strHistoryText & "; Limit Mass Range = " & Trim(.LimitMassRange)
            If .LimitMassRange Then
                strHistoryText = strHistoryText & "; Mass Start = " & Trim(.MassStart)
                strHistoryText = strHistoryText & "; Mass End = " & Trim(.MassEnd)
            End If
            
        End With
        
        AddToAnalysisHistory CallerID, strHistoryText
    Else
        If Len(strErrorMessage) > 0 Then strErrorMessage = ": " & strErrorMessage
        strErrorMessage = "Auto-populate error or cancelled" & strErrorMessage
        Status strErrorMessage
    End If

AutoPopulateExitFunction:

    DestroyStructuresLocal
    
    txtExclusionList.Locked = False
    cmdAutoPopulate.Visible = True
    Me.MousePointer = vbDefault
    AutoPopulateStart = lngIonsAddedCount
    
    Exit Function

AutoPopulateStartErrorHandler:
    Debug.Print "Error in AutoPopulateStart Function: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "frmExcludeMassRange->AutoPopulateStart"
    lngIonsAddedCount = -1
    Status "Auto-populate error"
    Resume AutoPopulateExitFunction

End Function

Private Function AutoPopulateWork(udtNoiseRemovalOptions As udtNoiseRemovalOptionsType) As Long
    ' Returns the number of ions added to txtExclusionList
    ' Returns -1 if an error
    
    ' Step through ions
    ' If given ion is not excluded, then compute mass half window based on dblSearchTolerancePPM
    ' Perform mass search
    ' Compute match count scan percentage
    ' If percentage is > threshold, then change search mass to given ion mass plus
    '   half mass window, search for number of matching ions to give neighbor slice
    '   ion coun
    ' Repeat neighbor slice search, but use given ion mass minus half mass window
    ' Add to exclusion list if the neighbor ion counts are < lngNeighborSliceMaxIonCount
    
    
    Dim udtAutoPopulateStats As udtAutoPopulateStatsType
    Dim udtNewExclusionEntry As udtExclusionIonType
    
    Dim dblSearchIonMass As Double
    Dim dblSearchIonAbsMassTol As Double
    
    Dim dblExcludeIonStartMass As Double, dblExcludeIonEndMass
    
    Dim lngCurrentIndex As Long
    Dim lngSearchIonScan As Long
    
    Dim intSearchIonCharge As Integer
    
    Dim blnDone As Boolean, blnFoundNext As Boolean
    Dim blnSkipSearchIon As Boolean
    Dim blnNeighborSliceIncluded As Boolean
    
    Dim lngIonsAddedCount As Long
    
    On Error GoTo AutoPopulateWorkErrorHandler
    
    ' Initialize udtAutoPopulateStats
    With udtAutoPopulateStats
        ' Determine the Scan Range
        ' First, fill .ScanRangeStart and .ScanRangeEnd with the current scan range for the given gel
        ' If the scope is All, then the full range is used
        ' If the scope is Current, then only the current range is used
        If udtNoiseRemovalOptions.SearchScope = glScope.glSc_All Then
            GetScanRange CallerID, .ScanRangeStart, .ScanRangeEnd, .ScanRangeCount
        Else
            Debug.Assert udtNoiseRemovalOptions.SearchScope = glScope.glSc_Current
            .ScanRangeStart = GelBody(CallerID).csMyCooSys.CurrRXMin
            .ScanRangeEnd = GelBody(CallerID).csMyCooSys.CurrRXMax
        End If
        
        ' If the user has enabled Limit Scan Range, then may need to decrease
        '   udtAutoPopulateStats.scanRangeStart and udtAutoPopulateStats.ScanRangeEnd;
        '   however, we will not increase them
        .LimitScanRange = udtNoiseRemovalOptions.LimitScanRange
        If .LimitScanRange Then
            If udtNoiseRemovalOptions.ScanStart > .ScanRangeStart Then .ScanRangeStart = udtNoiseRemovalOptions.ScanStart
            If udtNoiseRemovalOptions.ScanEnd < .ScanRangeEnd Then .ScanRangeEnd = udtNoiseRemovalOptions.ScanEnd
            
            If .ScanRangeEnd < .ScanRangeStart Then .ScanRangeEnd = .ScanRangeStart
        End If
        
        If udtNoiseRemovalOptions.SearchScope = glScope.glSc_Current Then
            ' We are effectively limiting the scan range, so set this to true
            .LimitScanRange = True
        End If
        
        ' Compute the scan range count, given the start and end scan numbers
        .ScanRangeCount = .ScanRangeEnd - .ScanRangeStart + 1
        If .ScanRangeCount < 1 Then
            ' This shouldn't happen
            Debug.Assert False
            lngIonsAddedCount = -1
            blnDone = True
        End If
        
        ' Copy the other options from udtNoiseRemovalOptions
        .LimitMassRange = udtNoiseRemovalOptions.LimitMassRange
        .MassStart = udtNoiseRemovalOptions.MassStart
        .MassEnd = udtNoiseRemovalOptions.MassEnd
        
        .RequireIdenticalCharge = udtNoiseRemovalOptions.RequireIdenticalCharge
    End With
    
    ' Make sure .ExclusionList() is up-to-date
    FillExclusionListFromTextbox False

    lngCurrentIndex = -1
    lngIonsAddedCount = 0
    
    Do Until blnDone
        blnFoundNext = False
        Do Until blnFoundNext
            lngCurrentIndex = lngCurrentIndex + 1
            If lngCurrentIndex > O_Cnt - 1 Then     'all data has been used
                blnDone = True
                Exit Do
            Else
                ' no need to search using already matched (excluded) data
                If Not IsMatched(IndOrder(lngCurrentIndex)) Then blnFoundNext = True
            End If
        Loop
        
        If blnFoundNext Then      ' new ion found to use as the search mass
            
            If lngCurrentIndex Mod 50 = 0 Then
                Status "Populating exclusion list: " & Trim(Format(lngCurrentIndex / O_Cnt * 100, "0.0")) & "% done"
            End If
        
            dblSearchIonMass = O_MW(IndOrder(lngCurrentIndex))
            lngSearchIonScan = O_Scan(IndOrder(lngCurrentIndex))
            intSearchIonCharge = O_Charge(IndOrder(lngCurrentIndex))
            
            blnSkipSearchIon = False
            With udtAutoPopulateStats
                If .LimitMassRange Then
                    ' Skip this search ion if its mass is out of range
                    If dblSearchIonMass < .MassStart Or dblSearchIonMass > .MassEnd Then
                        blnSkipSearchIon = True
                    End If
                End If
            
                If Not blnSkipSearchIon And .LimitScanRange Then
                    ' Skip this search ion if its scan is out of range
                    If lngSearchIonScan < .ScanRangeStart Or lngSearchIonScan > .ScanRangeEnd Then
                        blnSkipSearchIon = True
                    End If
                End If
            End With
            
            If Not blnSkipSearchIon Then
                ' Find all data points within tolerance of the search ion
                dblSearchIonAbsMassTol = PPMToMass(udtNoiseRemovalOptions.SearchTolerancePPMAutoRemoval, dblSearchIonMass)
                
                With udtAutoPopulateStats
                    .SearchMass = dblSearchIonMass
                    .AbsMassTol = dblSearchIonAbsMassTol
                    .Charge = intSearchIonCharge
                    .PercentScansInUseThreshold = udtNoiseRemovalOptions.PercentageThresholdToExcludeSlice
                End With
                
                If AutoPopulateFindPercentScansInUse(udtAutoPopulateStats) Then
                    ' This slice has a percent scans in use value over the threshold

                    dblExcludeIonStartMass = udtAutoPopulateStats.SearchMass - udtAutoPopulateStats.AbsMassTol
                    dblExcludeIonEndMass = udtAutoPopulateStats.SearchMass + udtAutoPopulateStats.AbsMassTol
                    
                    ' Check the neighbor slices and possibly add them to the search slice
                    blnNeighborSliceIncluded = False
                    
                    ' Negative neighbor
                    With udtAutoPopulateStats
                        .SearchMass = dblSearchIonMass - dblSearchIonAbsMassTol
                        .PercentScansInUseThreshold = udtNoiseRemovalOptions.PercentageThresholdToAddNeighborToSearchSlice
                    End With
                
                    If AutoPopulateFindPercentScansInUse(udtAutoPopulateStats) Then
                        ' Adjust dblExcludeIonStartMass down
                        dblExcludeIonStartMass = udtAutoPopulateStats.SearchMass - udtAutoPopulateStats.AbsMassTol
                        blnNeighborSliceIncluded = True
                    End If
                    
                    ' Positive neighbor
                    With udtAutoPopulateStats
                        .SearchMass = dblSearchIonMass + dblSearchIonAbsMassTol
                        .PercentScansInUseThreshold = udtNoiseRemovalOptions.PercentageThresholdToAddNeighborToSearchSlice
                    End With
                
                    If AutoPopulateFindPercentScansInUse(udtAutoPopulateStats) Then
                        ' Adjust dblExcludeIonEndMass up
                        dblExcludeIonEndMass = udtAutoPopulateStats.SearchMass + udtAutoPopulateStats.AbsMassTol
                        blnNeighborSliceIncluded = True
                    End If
                    
                    With udtNewExclusionEntry
                        If blnNeighborSliceIncluded Then
                            ' Average the start and end masses to get the central mass
                            .IonMass = Round((dblExcludeIonStartMass + dblExcludeIonEndMass) / 2, 6)
                            
                            ' The new PPM tolerance is 50% of the mass range, converted to ppm
                            .TolerancePPM = Round(MassToPPM((dblExcludeIonEndMass - dblExcludeIonStartMass) / 2, .IonMass), 2)
                        Else
                            ' Neighbor slices were not included, so simply use the search ion specs
                            .IonMass = dblSearchIonMass
                            .TolerancePPM = udtNoiseRemovalOptions.SearchTolerancePPMAutoRemoval
                        End If
                        
                        If udtAutoPopulateStats.RequireIdenticalCharge Then
                            .Charge = intSearchIonCharge
                        Else
                            .Charge = 0
                        End If
                        
                        If udtAutoPopulateStats.LimitScanRange Then
                            .LimitScanRange = True
                            .ScanStart = udtAutoPopulateStats.ScanRangeStart
                            .ScanEnd = udtAutoPopulateStats.ScanRangeEnd
                        Else
                            .LimitScanRange = False
                            .ScanStart = 0
                            .ScanEnd = 0
                        End If
                    End With
                    
                    AddToExclusionList udtNewExclusionEntry, True, False, True
                    
                    lngIonsAddedCount = lngIonsAddedCount + 1
                
                End If
            End If
        End If
    
        If mCancelOperation Then Exit Do
    Loop
    
    AutoPopulateWork = lngIonsAddedCount
    Exit Function
    
AutoPopulateWorkErrorHandler:
    Debug.Assert False
    Status "Error creating auto-populating the exclusion list"
    AutoPopulateWork = -1
    
End Function

Private Sub EnableDisableControls()
    Dim blnEnable As Boolean
    
    blnEnable = cChkBox(chkAutoPopulateLimitMassRange)
    txtAutoPopulateLimitMassRangeStart.Enabled = blnEnable
    txtAutoPopulateLimitMassRangeEnd.Enabled = blnEnable
    
    blnEnable = cChkBox(chkAutoPopulateLimitScanRange)
    txtAutoPopulateLimitScanRangeStart.Enabled = blnEnable
    txtAutoPopulateLimitScanRangeEnd.Enabled = blnEnable

End Sub

Private Sub DestroyStructuresLocal()
    On Error Resume Next
    O_Cnt = 0
    CSCnt = 0
    ISCnt = 0
    Erase O_Index
    Erase O_Type
    Erase O_MW
    Erase O_Charge
    Erase O_Order
    Erase O_Scan
    Erase S_MW
    Erase IndMW
    Erase IndOrder
    Erase IsMatched
    Set MWRangeFinder = Nothing
End Sub

Private Function CreateIndMW() As Boolean
    '--------------------------------------------------------------
    'creates index on molecular mass; sorts its members and creates
    'fast search object; returns True if successful
    '--------------------------------------------------------------
    Dim qsDbl As New QSDouble
    On Error GoTo err_CreateIndMW
    Status "Creating MW index"
    S_MW() = O_MW()                                             'array assignment
    CreateIndMW = qsDbl.QSAsc(S_MW, IndMW)
    Set MWRangeFinder = New MWUtil
    If Not MWRangeFinder.Fill(S_MW) Then GoTo err_CreateIndMW
    
exit_CreateIndMW:
    Set qsDbl = Nothing
    Exit Function
    
err_CreateIndMW:
    Erase IndMW
    Erase S_MW
    Resume exit_CreateIndMW
    Status "Error creating MW index"
End Function

Private Function CreateIndOrder() As Boolean
    '--------------------------------------------------------------
    'creates index on order by decreasing abundance and returns True if successful;
    '--------------------------------------------------------------
    Dim TmpOrder() As Double
    Dim qsDbl As New QSDouble
    On Error GoTo err_CreateIndOrder
    Status "Creating order index"
    TmpOrder() = O_Order()                                      'array assignment
    CreateIndOrder = qsDbl.QSDesc(TmpOrder, IndOrder)           'higher intensity is better
    
exit_CreateIndOrder:
    Set qsDbl = Nothing
    Exit Function
    
err_CreateIndOrder:
    Erase IndOrder
    Resume exit_CreateIndOrder
    Status "Error creating order index"
End Function

Private Sub FillExclusionListFromTextbox(Optional blnUpdateTextBox As Boolean = True, Optional ByVal strDelimeterList As String = ",;", Optional ByVal blnUseSpaceDelimeter As Boolean = True, Optional ByVal blnUseTabDelimeter As Boolean = True)

    Const MAX_LIST_COUNT = 1000000
    Const MAX_PARSE_COUNT = 5
    
    Dim strListItems() As String        ' 0-based array
    Dim lngItemCount As Long
    
    Dim strExclusionList As String
    
    Dim lngindex As Long
    Dim dblParsedVals() As Double       ' 0-based array
    Dim intParseCount As Integer
    
    Dim dblDefaultTolerancePPM As Double
    
    dblDefaultTolerancePPM = CDblSafe(txtDefaultMassSliceWidthPPM)
    
    ' Clear .ExclusionList
    With glbPreferencesExpanded.NoiseRemovalOptions
        .ExclusionListCount = 0
        ReDim .ExclusionList(0)
    End With
    
    strExclusionList = txtExclusionList.Text
    If Len(strExclusionList) = 0 Then
        lngItemCount = 0
    Else
        ' The following will populate strListItems() with the each row in the exclusion list
        ParseAndSortList strExclusionList, strListItems(), lngItemCount, "", False, True, True, True, True, MAX_LIST_COUNT
    End If
        
    If blnUseSpaceDelimeter Then strDelimeterList = strDelimeterList & " "
    If blnUseTabDelimeter Then strDelimeterList = strDelimeterList & vbTab
    
    ' Parse each item in strListItems() to find the specific exclusion parameters
    For lngindex = 0 To lngItemCount - 1
        intParseCount = ParseStringValuesDbl(strListItems(lngindex), dblParsedVals(), MAX_PARSE_COUNT, strDelimeterList, "", False, True, False)
    
        If intParseCount >= 1 Then
            With glbPreferencesExpanded.NoiseRemovalOptions
                .ExclusionListCount = .ExclusionListCount + 1
                ReDim Preserve .ExclusionList(0 To .ExclusionListCount - 1)
                With .ExclusionList(.ExclusionListCount - 1)
                    .IonMass = dblParsedVals(0)
                    
                    If intParseCount >= 2 Then
                        .TolerancePPM = dblParsedVals(1)
                    Else
                        .TolerancePPM = dblDefaultTolerancePPM
                    End If
                    If intParseCount >= 3 Then
                        .Charge = dblParsedVals(2)
                    Else
                        .Charge = 0
                    End If
                    If intParseCount >= 5 Then
                        .LimitScanRange = True
                        .ScanStart = dblParsedVals(3)
                        .ScanEnd = dblParsedVals(4)
                    Else
                        .LimitScanRange = False
                    End If
                End With
            End With
        End If
    Next lngindex
    
    If blnUpdateTextBox Then
        FillExclusionTextBox
    End If

End Sub

Private Sub FillExclusionTextBox()
    ' Fills the exclusion textbox using the data in .ExclusionList()
    ' Format is: Monoisotopic mass, ppm width, charge, scan start, scan end
    
    Dim strExclusionList As String
    Dim dblExclusionMasses() As Double, lngExclusionMassPointers() As Long
    Dim blnSuccess As Boolean
    Dim qsDbl As New QSDouble
    
    Dim lngindex As Long
    
    ' Fill in order of increasing mass
    strExclusionList = ""
    With glbPreferencesExpanded.NoiseRemovalOptions
        If .ExclusionListCount > 0 Then
            If .ExclusionListCount > 1 Then
                ReDim dblExclusionMasses(0 To .ExclusionListCount - 1)
                ReDim lngExclusionMassPointers(0 To .ExclusionListCount - 1)
                For lngindex = 0 To .ExclusionListCount - 1
                    dblExclusionMasses(lngindex) = .ExclusionList(lngindex).IonMass
                    lngExclusionMassPointers(lngindex) = lngindex
                Next lngindex
                
                blnSuccess = qsDbl.QSAsc(dblExclusionMasses(), lngExclusionMassPointers())
                Debug.Assert blnSuccess
            Else
                ReDim lngExclusionMassPointers(0)
                lngExclusionMassPointers(0) = 0
            End If
            
            For lngindex = 0 To .ExclusionListCount - 1
                With .ExclusionList(lngExclusionMassPointers(lngindex))
                    strExclusionList = strExclusionList & Trim(.IonMass) & ", " & Trim(.TolerancePPM) & ", " & Trim(.Charge)
                    If .LimitScanRange Then
                        strExclusionList = strExclusionList & ", " & Trim(.ScanStart) & ", " & Trim(.ScanEnd)
                    End If
                    strExclusionList = strExclusionList & vbCrLf
                End With
            Next lngindex
        End If
    End With
    
    txtExclusionList.Text = strExclusionList
    DoEvents
    
    Set qsDbl = Nothing
End Sub

Public Function IncludeExcludeIons(blnExcludeIons As Boolean) As Long
    ' Includes or excludes the ions in the data matching the ions in txtExclusionList
    ' Returns the number of ions matched, and thus included or excluded
    ' Returns -1 if an error
    ' Returns 0 if txtExclusionList is empty
    
    Dim lngIonMatchCount As Long
    Dim blnSuccess As Boolean
    Dim eScopeAtStart As glScope
    
    Dim lngExclusionIndex As Long, lngindex As Long
    Dim lngMultiplier As Long           ' -1 for exclude, 1 for include
    
    Dim dblAbsMassTolerance As Double
    Dim MWRangeMinInd As Long, MWRangeMaxInd As Long
    
On Error GoTo IncludeExcludeIonsErrorHandler
    
    If blnExcludeIons Then
        lngMultiplier = -1
    Else
        lngMultiplier = 1
    End If
    
    mCancelOperation = False
    cmdAutoPopulate.Visible = False
    Me.MousePointer = vbHourglass

    FillExclusionListFromTextbox False
    
    lngIonMatchCount = 0
    If glbPreferencesExpanded.NoiseRemovalOptions.ExclusionListCount > 0 Then
        
        blnSuccess = True
        UpdateCurrentNoiseRemovalOptions
        
        With glbPreferencesExpanded.NoiseRemovalOptions
            ' Must have current scope be All when Including ions (i.,e. whe blnExcludeIons = False)
            eScopeAtStart = .SearchScope
            If Not blnExcludeIons And .SearchScope = glScope.glSc_Current Then
                .SearchScope = glScope.glSc_All
            End If
        End With
        
        If InitializeSearchIndices(glbPreferencesExpanded.NoiseRemovalOptions.SearchScope) And Not mCancelOperation Then
            If Not CreateIndMW() Or mCancelOperation Then blnSuccess = False
            
            If blnSuccess Then
                ' Find the ions in the data matching each search specification in .ExclusionList()
                
                With glbPreferencesExpanded.NoiseRemovalOptions
                    For lngExclusionIndex = 0 To .ExclusionListCount - 1
                        Status "Working: " & Trim(lngExclusionIndex) & " of " & Trim(.ExclusionListCount)
                        
                        With .ExclusionList(lngExclusionIndex)
                            dblAbsMassTolerance = PPMToMass(.TolerancePPM, .IonMass)
                            
                            If MWRangeFinder.FindIndexRange(.IonMass, dblAbsMassTolerance, MWRangeMinInd, MWRangeMaxInd) Then
                                ' Mark the ions in the given range as "Matched"
                                For lngindex = MWRangeMinInd To MWRangeMaxInd
                                    If Not .LimitScanRange Or (O_Scan(IndMW(lngindex)) >= .ScanStart And O_Scan(IndMW(lngindex)) <= .ScanEnd) Then
                                        IsMatched(IndMW(lngindex)) = True
                                    End If
                                Next lngindex
                            End If
                            
                        End With
                    Next lngExclusionIndex
                End With
                
                ' Exclude data by setting .CSID or .IsoID to a negative value for data with IsMatched() = True
                ' Include data by setting .CSID or .IsoID to a positive value for data with IsMatched() = true
                For lngindex = 0 To O_Cnt - 1
                    If IsMatched(lngindex) Then
                        If O_Type(lngindex) = gldtCS Then
                            GelDraw(CallerID).CSID(O_Index(lngindex)) = lngMultiplier * Abs(GelDraw(CallerID).CSID(O_Index(lngindex)))
                        Else
                            Debug.Assert O_Type(lngindex) = gldtIS
                            GelDraw(CallerID).IsoID(O_Index(lngindex)) = lngMultiplier * Abs(GelDraw(CallerID).IsoID(O_Index(lngindex)))
                        End If
                        lngIonMatchCount = lngIonMatchCount + 1
                    End If
                Next lngindex
            
                GelBody(CallerID).RequestRefreshPlot
                
            End If
        Else
            blnSuccess = False
            lngIonMatchCount = -1
        End If
        
        If blnSuccess Then
            If blnExcludeIons Then
                Status "Success. Explicitly excluded " & Trim(lngIonMatchCount) & " ions"
                AddToAnalysisHistory CallerID, "Excluded mass ranges (noise streaks): Ion removal count = " & Trim(lngIonMatchCount) & "; Search ion count = " & Trim(glbPreferencesExpanded.NoiseRemovalOptions.ExclusionListCount)
            Else
                Status "Success. Explicitly included " & Trim(lngIonMatchCount) & " ions"
                AddToAnalysisHistory CallerID, "Included mass ranges: Ion inclusion count = " & Trim(lngIonMatchCount) & "; Search ion count = " & Trim(glbPreferencesExpanded.NoiseRemovalOptions.ExclusionListCount)
            End If
        Else
            Status "Operation cancelled (or an error has occurred)"
        End If
        
        ' Restore the search scope
        glbPreferencesExpanded.NoiseRemovalOptions.SearchScope = eScopeAtStart
    End If

IncludeExcludeIonsExitFunction:

    DestroyStructuresLocal
    
    cmdAutoPopulate.Visible = True
    Me.MousePointer = vbDefault
    IncludeExcludeIons = lngIonMatchCount
    Exit Function

IncludeExcludeIonsErrorHandler:
    Debug.Print "Error in IncludeExcludeIons Function: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "frmExcludeMassRange->IncludeExcludeIons"
    lngIonMatchCount = -1
    Status "Include/Exclude ions error"
    Resume IncludeExcludeIonsExitFunction
    
End Function

Public Sub InitializeForm()

    Dim strDirections As String

    If Len(Me.Tag) > 0 Then
        If IsNumeric(Me.Tag) Then
            ' Use SetCallerID() function to set the CallerID rather than via .Tag
            Debug.Assert False
            CallerID = val(Me.Tag)
         End If
    End If
    
    strDirections = ""
    strDirections = strDirections & "Enter a list of masses to exclude, one mass per line.  Use the de-isotoped, monoisotopic mass (not the m/z value or the MH+ mass).  "
    strDirections = strDirections & "You may also include the ppm tolerance (mass slice half-width), the charge to match, and the start and end scan number to limit to.  "
    strDirections = strDirections & "A charge state of 0 will match all charges.  Separate the values using commas or semicolons."
    lblDirections.Caption = strDirections

    With glbPreferencesExpanded.NoiseRemovalOptions
        txtDefaultMassSliceWidthPPM = Trim(.SearchTolerancePPMDefault)
        txtAutoPopulateMassSliceWidthPPM = Trim(.SearchTolerancePPMAutoRemoval)
        
        txtAutoPopulateCountPercentageThresholdForExclusion = Trim(.PercentageThresholdToExcludeSlice)
        txtAutoPopulateNeighborCountPercentageThreshold = Trim(.PercentageThresholdToAddNeighborToSearchSlice)
        
        SetCheckBox chkAutoPopulateLimitMassRange, .LimitMassRange
        txtAutoPopulateLimitMassRangeStart = Trim(.MassStart)
        txtAutoPopulateLimitMassRangeEnd = Trim(.MassEnd)
        
        SetCheckBox chkAutoPopulateLimitScanRange, .LimitScanRange
        txtAutoPopulateLimitScanRangeStart = Trim(.ScanStart)
        txtAutoPopulateLimitScanRangeEnd = Trim(.ScanEnd)
        
        If .SearchScope = glScope.glSc_All Then
            optAutoPopulateSearchScope(glScope.glSc_All).Value = True
        Else
            optAutoPopulateSearchScope(glScope.glSc_Current).Value = True
        End If
        
        SetCheckBox chkAutoPopulateRequireIdenticalCharge, .RequireIdenticalCharge
    End With
    
End Sub

Private Function InitializeSearchIndices(eSearchScope As glScope) As Boolean
    ' Initializes the search indices
    ' Returns True if success, or False if failure or no data in scope
    ' intSearchScope should be 0 to search all the data, or 1 to search just the current view
    
    ' This sub modelled after that in frmUMCSimple, written by Nikola Tolic in Spring 2003
    
    Dim MaxCnt As Long
    Dim CSInd() As Long
    Dim ISInd() As Long
    Dim i As Long
    
    On Error GoTo InitializeSearchIndicesErrorHandler
    
    Status "Loading arrays"
    
    MaxCnt = GelData(CallerID).CSLines + GelData(CallerID).IsoLines
    If MaxCnt > 0 Then
       ReDim O_Index(MaxCnt - 1)
       ReDim O_Type(MaxCnt - 1)
       ReDim O_MW(MaxCnt - 1)
       ReDim O_Charge(MaxCnt - 1)
       ReDim O_Order(MaxCnt - 1)
       ReDim O_Scan(MaxCnt - 1)
       O_Cnt = 0
       With GelData(CallerID)
         CSCnt = GetCSScope(CallerID, CSInd(), eSearchScope)
         If CSCnt > 0 Then
            For i = 1 To CSCnt
                O_Cnt = O_Cnt + 1
                O_Index(O_Cnt - 1) = CSInd(i)
                O_Type(O_Cnt - 1) = gldtCS
                O_MW(O_Cnt - 1) = .CSData(CSInd(i)).AverageMW
                O_Charge(O_Cnt - 1) = .CSData(CSInd(i)).Charge
                O_Scan(O_Cnt - 1) = .CSData(CSInd(i)).ScanNumber
                O_Order(O_Cnt - 1) = .CSData(CSInd(i)).Abundance
            Next i
         End If
         ISCnt = GetISScope(CallerID, ISInd(), eSearchScope)
         If ISCnt > 0 Then
            For i = 1 To ISCnt
                O_Cnt = O_Cnt + 1
                O_Index(O_Cnt - 1) = ISInd(i)
                O_Type(O_Cnt - 1) = gldtIS
                O_MW(O_Cnt - 1) = GetIsoMass(.IsoData(ISInd(i)), UMCDef.MWField)
                O_Charge(O_Cnt - 1) = .IsoData(ISInd(i)).Charge
                O_Scan(O_Cnt - 1) = .IsoData(ISInd(i)).ScanNumber
                O_Order(O_Cnt - 1) = .IsoData(ISInd(i)).Abundance
            Next i
         End If
       End With
    End If
    If O_Cnt <= 0 Then Status "No data found in scope"
    
InitializeSearchIndicesErrorHandlerExit:

    If O_Cnt > 0 Then
       ReDim Preserve O_Index(O_Cnt - 1)
       ReDim Preserve O_Type(O_Cnt - 1)
       ReDim Preserve O_MW(O_Cnt - 1)
       ReDim Preserve O_Charge(O_Cnt - 1)
       ReDim Preserve O_Order(O_Cnt - 1)
       ReDim Preserve O_Scan(O_Cnt - 1)
       'initialize index arrays
       ReDim IndMW(O_Cnt - 1)
       ReDim IndOrder(O_Cnt - 1)
       For i = 0 To O_Cnt - 1
           IndMW(i) = i
           IndOrder(i) = i
       Next i
       ReDim IsMatched(O_Cnt - 1)
       InitializeSearchIndices = True
    Else
       Erase O_Index
       Erase O_Type
       Erase O_MW
       Erase O_Charge
       Erase O_Order
       Erase O_Scan
       InitializeSearchIndices = False
    End If
    
    Exit Function
    
InitializeSearchIndicesErrorHandler:
    Debug.Assert False
    O_Cnt = 0               'this will cause everything to be cleared
    Resume InitializeSearchIndicesErrorHandlerExit
    Status "Error loading arrays"
    
End Function

Public Sub RestoreDefaultFilterPoints()
    
    frmFilter.Tag = CallerID
    frmFilter.InitializeControls True

End Sub

Public Sub SetCallerID(ByVal lngGelIndex As Long)
    CallerID = lngGelIndex
End Sub

Private Sub Status(ByVal Msg As String)
    lblStatus.Caption = Msg
    DoEvents
End Sub

Private Sub UpdateCurrentNoiseRemovalOptions()

On Error GoTo UpdateCurrentNoiseRemovalOptionsErrorHandler

    With glbPreferencesExpanded.NoiseRemovalOptions
        .SearchTolerancePPMAutoRemoval = CDblSafe(txtAutoPopulateMassSliceWidthPPM)
        .SearchTolerancePPMDefault = CDblSafe(txtDefaultMassSliceWidthPPM)
        
        .PercentageThresholdToExcludeSlice = CSngSafe(txtAutoPopulateCountPercentageThresholdForExclusion)
        .PercentageThresholdToAddNeighborToSearchSlice = CSngSafe(txtAutoPopulateNeighborCountPercentageThreshold)
        
        .LimitMassRange = cChkBox(chkAutoPopulateLimitMassRange)
        .MassStart = CDblSafe(txtAutoPopulateLimitMassRangeStart)
        .MassEnd = CDblSafe(txtAutoPopulateLimitMassRangeEnd)
    
        .LimitScanRange = cChkBox(chkAutoPopulateLimitScanRange)
        .ScanStart = CLngSafe(txtAutoPopulateLimitScanRangeStart)
        .ScanEnd = CLngSafe(txtAutoPopulateLimitScanRangeEnd)
    
        If optAutoPopulateSearchScope(0).Value = True Then
            .SearchScope = glScope.glSc_All
        Else
            .SearchScope = glScope.glSc_Current
        End If
        
        .RequireIdenticalCharge = cChkBox(chkAutoPopulateRequireIdenticalCharge)
    End With

    Exit Sub

UpdateCurrentNoiseRemovalOptionsErrorHandler:
    Debug.Print "Error in UpdateCurrentNoiseRemovalOptions Sub: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "frmExcludeMassRange->UpdateCurrentNoiseRemovalOptions"
    Resume Next
    
End Sub

Private Sub chkAutoPopulateLimitMassRange_Click()
    EnableDisableControls
End Sub

Private Sub chkAutoPopulateLimitScanRange_Click()
    EnableDisableControls
End Sub

Private Sub cmdAutoPopulate_Click()
    AutoPopulateStart
End Sub

Private Sub cmdAutoPopulateCancel_Click()
    mCancelOperation = True
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdExclude_Click()
    IncludeExcludeIons True
End Sub

Private Sub cmdInclude_Click()
    IncludeExcludeIons False
End Sub

Private Sub cmdRestoreDefaultFilter_Click()
    RestoreDefaultFilterPoints
End Sub

Private Sub Form_Activate()
    InitializeForm
End Sub

Private Sub Form_Load()
    
    SizeAndCenterWindow Me, cWindowUpperThird, 9050, 7550, False
    
    EnableDisableControls
    
End Sub

Private Sub txtAutoPopulateCountPercentageThresholdForExclusion_GotFocus()
    TextBoxGotFocusHandler txtAutoPopulateCountPercentageThresholdForExclusion, False
End Sub

Private Sub txtAutoPopulateCountPercentageThresholdForExclusion_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtAutoPopulateCountPercentageThresholdForExclusion, KeyAscii, True, True, False
End Sub

Private Sub txtAutoPopulateCountPercentageThresholdForExclusion_LostFocus()
    ValidateTextboxValueDbl txtAutoPopulateCountPercentageThresholdForExclusion, 0, 100, 50
End Sub


Private Sub txtAutoPopulateLimitMassRangeEnd_GotFocus()
    TextBoxGotFocusHandler txtAutoPopulateLimitMassRangeEnd, False
End Sub

Private Sub txtAutoPopulateLimitMassRangeEnd_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtAutoPopulateLimitMassRangeEnd, KeyAscii, True, True, False
End Sub

Private Sub txtAutoPopulateLimitMassRangeEnd_LostFocus()
    ValidateTextboxValueDbl txtAutoPopulateLimitMassRangeEnd, 0, 1000000, 6000
End Sub

Private Sub txtAutoPopulateLimitMassRangeStart_GotFocus()
    TextBoxGotFocusHandler txtAutoPopulateLimitMassRangeStart, False
End Sub

Private Sub txtAutoPopulateLimitMassRangeStart_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtAutoPopulateLimitMassRangeStart, KeyAscii, True, True, False
End Sub

Private Sub txtAutoPopulateLimitMassRangeStart_LostFocus()
    ValidateTextboxValueLng txtAutoPopulateLimitMassRangeStart, 0, 1000000, 100
End Sub


Private Sub txtAutoPopulateLimitScanRangeEnd_GotFocus()
    TextBoxGotFocusHandler txtAutoPopulateLimitScanRangeEnd, False
End Sub

Private Sub txtAutoPopulateLimitScanRangeEnd_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtAutoPopulateLimitScanRangeEnd, KeyAscii, True, False
End Sub

Private Sub txtAutoPopulateLimitScanRangeEnd_LostFocus()
    ValidateTextboxValueLng txtAutoPopulateLimitScanRangeEnd, 0, 1000000, 5000
End Sub

Private Sub txtAutoPopulateLimitScanRangeStart_GotFocus()
    TextBoxGotFocusHandler txtAutoPopulateLimitScanRangeStart, False
End Sub

Private Sub txtAutoPopulateLimitScanRangeStart_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtAutoPopulateLimitScanRangeStart, KeyAscii, True, False
End Sub

Private Sub txtAutoPopulateLimitScanRangeStart_LostFocus()
    ValidateTextboxValueLng txtAutoPopulateLimitScanRangeStart, 0, 1000000, 1
End Sub


Private Sub txtAutoPopulateMassSliceWidthPPM_GotFocus()
    TextBoxGotFocusHandler txtAutoPopulateMassSliceWidthPPM, False
End Sub

Private Sub txtAutoPopulateMassSliceWidthPPM_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtAutoPopulateMassSliceWidthPPM, KeyAscii, True, True, False
End Sub

Private Sub txtAutoPopulateMassSliceWidthPPM_LostFocus()
    ValidateTextboxValueDbl txtAutoPopulateMassSliceWidthPPM, 0, 1E+300, 2
End Sub

Private Sub txtAutoPopulateNeighborCountPercentageThreshold_GotFocus()
    TextBoxGotFocusHandler txtAutoPopulateNeighborCountPercentageThreshold, False
End Sub

Private Sub txtAutoPopulateNeighborCountPercentageThreshold_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtAutoPopulateNeighborCountPercentageThreshold, KeyAscii, True, False
End Sub

Private Sub txtAutoPopulateNeighborCountPercentageThreshold_LostFocus()
    ValidateTextboxValueDbl txtAutoPopulateNeighborCountPercentageThreshold, 0, 100, 50
End Sub

Private Sub txtDefaultMassSliceWidthPPM_Change()
    If IsNumeric(txtDefaultMassSliceWidthPPM) Then
        glbPreferencesExpanded.NoiseRemovalOptions.SearchTolerancePPMDefault = CDblSafe(txtDefaultMassSliceWidthPPM)
    End If
End Sub

Private Sub txtDefaultMassSliceWidthPPM_GotFocus()
    TextBoxGotFocusHandler txtDefaultMassSliceWidthPPM, False
End Sub

Private Sub txtDefaultMassSliceWidthPPM_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtDefaultMassSliceWidthPPM, KeyAscii, True, False
End Sub

Private Sub txtDefaultMassSliceWidthPPM_LostFocus()
    ValidateTextboxValueDbl txtDefaultMassSliceWidthPPM, 0, 1E+300, 2
End Sub

Private Sub txtExclusionList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then
        ' Allow the comma key
    Else
        TextBoxKeyPressHandler txtExclusionList, KeyAscii, True, True, False, False, False, False, False, False, True, False, True
    End If
End Sub
