VERSION 5.00
Begin VB.Form frmExcludePolygonRegion 
   Caption         =   "Filter By Polygon Region"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   10305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   10305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCreateExamplePolygon 
      Caption         =   "Create example polygon"
      Height          =   495
      Left            =   6360
      TabIndex        =   12
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdRestoreDefaultFilter 
      Caption         =   "Restore Default Filter Ion Visibility"
      Height          =   495
      Left            =   3720
      TabIndex        =   11
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Frame fraAutoPopulateSearchScope 
      Caption         =   "Search Scope"
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   1815
      Begin VB.OptionButton optAutoPopulateSearchScope 
         Caption         =   "&Current View"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optAutoPopulateSearchScope 
         Caption         =   "&All Data Points"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   280
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdInclude 
      Caption         =   "&Include Ions within Polygon"
      Height          =   495
      Left            =   2040
      TabIndex        =   8
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdExclude 
      Caption         =   "&Exclude Ions within Polygon"
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox txtPolygonVertices 
      Height          =   1335
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1080
      Width           =   4215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   3720
      TabIndex        =   9
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lblStatus 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   9975
   End
   Begin VB.Label lblExclusionListHeader 
      Caption         =   "List of X,Y coordinates (Scan, Monoisotopic mass) that define the vertices of the polygon region"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label lblDirections 
      Caption         =   "Directions"
      Height          =   2055
      Left            =   4680
      TabIndex        =   3
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label lblPolygonVertices 
      Caption         =   "Polygon Vertices"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmExcludePolygonRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private IsMatched() As Boolean      'is already matched

Private MWRangeFinder As MWUtil     'fast search of mass range

Private CallerID As Long
Private mCancelOperation As Boolean
'

Private Sub CreateExamplePolygon()
    Dim eResponse As VbMsgBoxResult
    Dim strPolygonDef As String
    
    If Len(txtPolygonVertices.Text) > 0 Then
        eResponse = MsgBox("Existing polygon vertices will be replaced.  Are you sure?", vbYesNoCancel + vbDefaultButton3 + vbQuestion, "Create Example Polygon")
        If eResponse <> vbYes Then Exit Sub
    End If
    
    strPolygonDef = "8000, 300" & vbCrLf & _
                    "20000, 300" & vbCrLf & _
                    "20000, 2750" & vbCrLf & _
                    "8000, 500"

    txtPolygonVertices.Text = strPolygonDef
    
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
    Erase IsMatched
    Set MWRangeFinder = Nothing
End Sub

Private Sub DefinePolygonsFromTextbox(Optional blnUpdateTextBox As Boolean = True)

    Dim strPolygonDefs As String
    Dim strWarnings As String
    
    PolygonTextToDefs txtPolygonVertices.Text
    
    If blnUpdateTextBox Then
        strPolygonDefs = PolygonDefsToText(glbPreferencesExpanded.NoiseRemovalOptions.ExclusionPolygonCount, glbPreferencesExpanded.NoiseRemovalOptions.ExclusionPolygonList, True, strWarnings)
        txtPolygonVertices.Text = strPolygonDefs
    End If

End Sub

Public Sub DefinePolygonRegions(ByVal strPolygonDefs As String)
    ' Note: Use PolygonDefsToText (in MonroeLaVRoutines) to convert from
    ' udtPolygonList() As udtExclusionPolygonType to a text-based list of vertex definitions
    
    txtPolygonVertices.Text = strPolygonDefs
 
End Sub

Public Function IncludeExcludeIons(blnExcludeIons As Boolean) As Long
    ' Includes or excludes the ions based on whether or not they are within the polygons defined by txtPolygonVertices
    ' Returns the number of ions matched, and thus included or excluded
    ' Returns -1 if an error
    ' Returns 0 if txtPolygonVertices is empty
    
    Dim lngIonMatchCount As Long
    Dim blnSuccess As Boolean
    Dim eScopeAtStart As glScope
    
    Dim lngExclusionIndex As Long, lngIndex As Long
    Dim lngMultiplier As Long           ' -1 for exclude, 1 for include
    
    Dim lngPolygonIndex As Long
    Dim lngIonIndex As Long
    
    Dim lngXVertexPoints() As Long
    Dim lngYVertextPoints() As Long
    
    Dim sngPercentComplete As Single
    
    Dim strMessage As String
    Dim strWarnings As String
    
    Dim objPointInPolygon As clsPointInPolygon
    
On Error GoTo IncludeExcludeIonsErrorHandler
    
    If blnExcludeIons Then
        lngMultiplier = -1
    Else
        lngMultiplier = 1
    End If
    
    mCancelOperation = False
    Me.MousePointer = vbHourglass

    DefinePolygonsFromTextbox False
    
    lngIonMatchCount = 0
    If glbPreferencesExpanded.NoiseRemovalOptions.ExclusionPolygonCount > 0 Then
        
        blnSuccess = True
        UpdateCurrentNoiseRemovalOptions
        
        With glbPreferencesExpanded.NoiseRemovalOptions
            ' Must have current scope be All when Including ions (i.e. when blnExcludeIons = False)
            eScopeAtStart = .SearchScope
            If Not blnExcludeIons And .SearchScope = glScope.glSc_Current Then
                .SearchScope = glScope.glSc_All
            End If
        End With
        
        Set objPointInPolygon = New clsPointInPolygon
        
        If InitializeSearchIndices(glbPreferencesExpanded.NoiseRemovalOptions.SearchScope) And Not mCancelOperation Then
            ' Find the ions in the data matching each search specification in .ExclusionList()
            
        
            For lngPolygonIndex = 0 To glbPreferencesExpanded.NoiseRemovalOptions.ExclusionPolygonCount - 1
                
                With glbPreferencesExpanded.NoiseRemovalOptions.ExclusionPolygonList(lngPolygonIndex)
                    If Not ParseVertexList(.VertexCount, .VertexList, lngXVertexPoints, lngYVertextPoints) Then
                        If strWarnings <> "" Then strWarnings = strWarnings & vbCrLf
                        strWarnings = "Invalid vertices defined for polygon index " & lngPolygonIndex
                    Else
                        If Not objPointInPolygon.SetPolygon(.VertexCount, lngXVertexPoints, lngYVertextPoints) Then
                            If strWarnings <> "" Then strWarnings = strWarnings & vbCrLf
                            strWarnings = objPointInPolygon.ErrorMessage
                        Else
                       
                        For lngIonIndex = 0 To O_Cnt - 1
                            If objPointInPolygon.TestPointInPolygonDbl(O_Scan(lngIonIndex), O_MW(lngIonIndex)) Then
                                IsMatched(lngIonIndex) = True
                            End If
                            
                            If lngIonIndex Mod 1000 = 0 Then
                                sngPercentComplete = (lngPolygonIndex / glbPreferencesExpanded.NoiseRemovalOptions.ExclusionPolygonCount + _
                                                      lngIonIndex / CDbl(O_Cnt) / glbPreferencesExpanded.NoiseRemovalOptions.ExclusionPolygonCount) * 100
                                Status "Working: " & Format(sngPercentComplete, "0.0") & "% complete"
                            End If
                        Next lngIonIndex
                    
                            
                        End If
                    
                    End If
                End With
                
            Next lngPolygonIndex
        
            strMessage = "Working: " & Format(100, "0.0") & "% complete"
            
            If strWarnings <> "" Then
                strMessage = strMessage & "; " & strWarnings
            End If
                        
            Status strMessage

            
            ' Exclude data by setting .CSID or .IsoID to a negative value for data with IsMatched() = True
            ' Include data by setting .CSID or .IsoID to a positive value for data with IsMatched() = true
            For lngIndex = 0 To O_Cnt - 1
                If IsMatched(lngIndex) Then
                    If O_Type(lngIndex) = gldtCS Then
                        GelDraw(CallerID).CSID(O_Index(lngIndex)) = lngMultiplier * Abs(GelDraw(CallerID).CSID(O_Index(lngIndex)))
                    Else
                        Debug.Assert O_Type(lngIndex) = gldtIS
                        GelDraw(CallerID).IsoID(O_Index(lngIndex)) = lngMultiplier * Abs(GelDraw(CallerID).IsoID(O_Index(lngIndex)))
                    End If
                    lngIonMatchCount = lngIonMatchCount + 1
                End If
            Next lngIndex
        
            GelBody(CallerID).RequestRefreshPlot
            
            
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
    
    Me.MousePointer = vbDefault
    IncludeExcludeIons = lngIonMatchCount
    Exit Function

IncludeExcludeIonsErrorHandler:
    Debug.Print "Error in IncludeExcludeIons Function: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "frmExcludePolygonRegion->IncludeExcludeIons"
    lngIonMatchCount = -1
    Status "Include/Exclude ions error"
    Resume IncludeExcludeIonsExitFunction
    
End Function

Public Sub InitializeForm()

    Dim strDirections As String
    Dim strPolygonDefs As String
    Dim strWarnings As String

    If Len(Me.Tag) > 0 Then
        If IsNumeric(Me.Tag) Then
            ' Use SetCallerID() function to set the CallerID rather than via .Tag
            Debug.Assert False
            CallerID = val(Me.Tag)
         End If
    End If
    
    strDirections = ""
    strDirections = strDirections & "Enter the coordinates of each vertex of a polygon that defines the region of data you wish to process. "
    strDirections = strDirections & "Each vertex is defined by a scan number and a monoisotopic mass, separated by a comma. "
    strDirections = strDirections & "To enter multiple polygons, list the coordinates for each polygon on a single line, separating the X,Y pairs with semicolons.  Next, define additional polygons on subsequent lines."
    strDirections = strDirections & "If the first line does not contain a semicolon, then we will assume that you are entering just one polygon region, with one X,Y pair per line."
    strDirections = strDirections & vbCrLf
    strDirections = strDirections & "For example, to exclude a triangular region in the bottom right corner of a separation with 20000 scans, "
    strDirections = strDirections & "enter the coordinates: 10000,250; 20000,250; 20000,2000"
    lblDirections.Caption = strDirections

    With glbPreferencesExpanded.NoiseRemovalOptions
        strPolygonDefs = PolygonDefsToText(.ExclusionPolygonCount, .ExclusionPolygonList, True, strWarnings)
        Me.DefinePolygonRegions strPolygonDefs
        
        If Len(strWarnings) > 0 Then
            lblStatus.Caption = strWarnings
        Else
            lblStatus.Caption = ""
        End If
        
        If .SearchScope = glScope.glSc_All Then
            optAutoPopulateSearchScope(glScope.glSc_All).Value = True
        Else
            optAutoPopulateSearchScope(glScope.glSc_Current).Value = True
        End If
        
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

Private Function ParseVertexList(ByVal intVertexCount As Integer, ByRef strVertexList() As String, ByRef lngXVertexPoints() As Long, ByRef lngYVertexPoints() As Long) As Boolean
    ' strVertexList() should contain a list of comma separated values
    ' The list will be parsed, and the values stored in lngXVertexPoints() and lngYVertexPoints()
    ' If any of the entries is invalid, then returns false
    
    Dim intIndex As Integer
    Dim intCommaLoc As Integer
    
    Dim strValue As String
    Dim lngX As Long
    Dim lngY As Long
    
    Dim blnSuccess As Boolean
    
    ReDim lngXVertexPoints(intVertexCount - 1)
    ReDim lngYVertexPoints(intVertexCount - 1)
    
On Error GoTo ParseVertexListErrorHandler

    blnSuccess = True
    For intIndex = 0 To intVertexCount - 1
        
        intCommaLoc = InStr(strVertexList(intIndex), ",")
        
        strValue = Left(strVertexList(intIndex), intCommaLoc - 1)
        If IsNumeric(strValue) Then
            lngX = CLng(strValue)
            
            strValue = Mid(strVertexList(intIndex), intCommaLoc + 1)
            If IsNumeric(strValue) Then
                lngY = CLng(strValue)
                
                lngXVertexPoints(intIndex) = lngX
                lngYVertexPoints(intIndex) = lngY
            Else
                blnSuccess = False
            End If
        Else
            blnSuccess = False
        End If
        
        If Not blnSuccess Then Exit For
    Next intIndex

    ParseVertexList = blnSuccess
    Exit Function

ParseVertexListErrorHandler:
    Debug.Print "Error in ParseVertexList: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "frmExcludePolygonRegion->ParseVertexList"

    ParseVertexList = False
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
    
        If optAutoPopulateSearchScope(0).Value = True Then
            .SearchScope = glScope.glSc_All
        Else
            .SearchScope = glScope.glSc_Current
        End If
        
    End With

    Exit Sub

UpdateCurrentNoiseRemovalOptionsErrorHandler:
    Debug.Print "Error in UpdateCurrentNoiseRemovalOptions Sub: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "frmExcludePolygonRegion->UpdateCurrentNoiseRemovalOptions"
    Resume Next
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCreateExamplePolygon_Click()
    CreateExamplePolygon
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
    
    SizeAndCenterWindow Me, cWindowUpperThird, 10400, 4900, False
   
End Sub

Private Sub txtPolygonVertices_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then
        ' Allow commas
    ElseIf KeyAscii = 13 Or KeyAscii = 10 Then
        ' Allow carriage return
    ElseIf KeyAscii = 59 Then
        ' Allow semicolons
    Else
        TextBoxKeyPressHandler txtPolygonVertices, KeyAscii, True, True, False, False, False, False, False, False, True, False, True
    End If
End Sub
