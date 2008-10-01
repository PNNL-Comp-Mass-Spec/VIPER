VERSION 5.00
Object = "{92D71E90-25A8-11CF-A640-9986B64D9618}#5.0#0"; "olch2x32.ocx"
Begin VB.UserControl ctlSpectraPlotter 
   ClientHeight    =   5385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9465
   KeyPreview      =   -1  'True
   ScaleHeight     =   5385
   ScaleWidth      =   9465
   Begin VB.Timer tmrAutoScale 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   8760
      Top             =   4680
   End
   Begin VB.Frame fraControls 
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   4440
      Width           =   8055
      Begin VB.CheckBox chkAutoScaleY 
         Caption         =   "Auto Scale &Y Axis"
         Height          =   255
         Left            =   3840
         TabIndex        =   5
         Top             =   240
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.Frame fraMouseAction 
         Caption         =   "Mouse Action"
         Height          =   645
         Left            =   1080
         TabIndex        =   3
         Top             =   0
         Width           =   2700
         Begin VB.ComboBox cboMouseAction 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.CommandButton cmdZoomOut 
         Caption         =   "&Zoom Out (Ctrl+A)"
         Height          =   540
         Left            =   75
         TabIndex        =   2
         Top             =   120
         Width           =   945
      End
      Begin VB.Label lblLocation 
         Caption         =   "Position"
         Height          =   375
         Left            =   5760
         TabIndex        =   6
         Top             =   120
         Width           =   2175
      End
   End
   Begin OlectraChart2D.Chart2D Chart2D 
      Height          =   4425
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9465
      _Version        =   327680
      _Revision       =   4
      _ExtentX        =   16695
      _ExtentY        =   7805
      _StockProps     =   0
      ControlProperties=   "ctlSpectraPlotter.ctx":0000
   End
End
Attribute VB_Name = "ctlSpectraPlotter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' This file contains constants used to assist in programming
' with the Olectra Chart controls.

' HugeValue is returned in some API calls when the control
' can't determine an appropriate value.
'
Private Const ocHugeValue As Double = 1E+308

' Windows messages, taken out of the WINAPI.TXT file
'
' Constants for dealing with mouse events
Private Const WM_MOUSEFIRST = &H200
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MOUSELAST = &H209

' Flags set when one of the mouse events is triggered
Private Const MK_LBUTTON = &H1
Private Const MK_MBUTTON = &H10
Private Const MK_RBUTTON = &H2

' Keyboard events for when a key is pressed/released
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101

' Flags set when a keyboard event is triggered
Private Const MK_ALT = &H20
Private Const MK_CONTROL = &H8
Private Const MK_SHIFT = &H4

' The Virtual Key codes
Private Const VK_ESCAPE = &H1B  'The <Esc> key (ASCII Character 27)
Private Const VK_SHIFT = &H10   'The <Shift> key
Private Const VK_CONTROL = &H11 'The <Ctrl> key

' UDT's

Private Type udtPlotRange
    XAxisMin As Double
    XAxisMax As Double
    YAxisMin As Double
    YAxisMax As Double
End Type

' Constants
Private Const ZOOM_HISTORY_MAX = 25
Private Const MAX_SERIES_COUNT = 3
Private Const MAX_GROUPS = 2            ' This is a fixed value for Olectra Chart

' Stack to keep track of the last 20 plot ranges displayed to allow for undoing
' 0 based array, ranging from 0 to ZOOM_HISTORY_MAX - 1
Private ZoomHistory(ZOOM_HISTORY_MAX, MAX_SERIES_COUNT) As udtPlotRange

' Private variables
Private mCurrentAction As ActionConstants
Private mTemporaryActionMode As Boolean
Private mZoomChanged As Boolean

Private mShowLocationDataPoint As Boolean
Private mShowLocationEverywhere As Boolean

Private mCurrentGroup As Integer
Private mCurrentSeries As Integer
Private mSeriesCount As Integer

Private mMasterOverrideOnIsBatch As Boolean

Public Sub AutoScaleXNow()
    With Chart2D
        .IsBatched = True
        With .ChartArea.Axes("X")
            .Min.IsDefault = True
            .Max.IsDefault = True
        End With
        If Not mMasterOverrideOnIsBatch Then .IsBatched = False
    End With
End Sub

Public Sub AutoScaleYNow()
    Dim lngIndex As Long, lngNumPoints As Long
    Dim lngLeftMostPoint As Long, lngRightMostPoint As Long
    Dim dblDataXVal As Double
    Dim dblDataYVal As Double
    Dim dblXStart As Double, dblXEnd As Double
    Dim dblZeroOrMinVisibleValue As Double, dblMaxVisibleValue As Double
    
    dblXStart = Chart2D.ChartArea.Axes("x").Min
    dblXEnd = Chart2D.ChartArea.Axes("x").Max
    
    ' This is inefficient, but I don't know a better way
    dblZeroOrMinVisibleValue = 0
    With Chart2D.ChartGroups(mCurrentGroup).Data
        lngLeftMostPoint = -1
        lngRightMostPoint = -1
        dblMaxVisibleValue = -1E+308
        lngNumPoints = .NumPoints(mCurrentSeries)
        For lngIndex = 1 To lngNumPoints
            dblDataXVal = .x(mCurrentSeries, lngIndex)
            dblDataYVal = .y(mCurrentSeries, lngIndex)
            If dblDataXVal <= dblXEnd Then
                If dblDataXVal >= dblXStart Then
                    If dblDataYVal > dblMaxVisibleValue Then dblMaxVisibleValue = dblDataYVal
                    If dblDataYVal < dblZeroOrMinVisibleValue Then dblZeroOrMinVisibleValue = dblDataYVal
                End If
            Else
                Exit For
            End If
        Next lngIndex
    End With
    
    With Chart2D
        .IsBatched = True
        With .ChartArea.Axes("Y")
            .Min.Value = dblZeroOrMinVisibleValue
            .Max.Value = dblMaxVisibleValue
        End With
        If Not mMasterOverrideOnIsBatch Then .IsBatched = False
    End With

End Sub

Private Sub ClearEvents()
    'Clear every event out of the ActionMaps.
    
    Chart2D.ActionMaps.RemoveAll
    mCurrentAction = oc2dActionNone

End Sub

Private Sub ComboBoxSetAction()
    Select Case cboMouseAction.ListIndex
    Case 1
        EnableActionMove
    Case 2
        EnableActionScale
    Case Else
        ' Includes case 0
        EnableActionZoom
    End Select
End Sub

Public Sub CopyToClipboard(eCopyFormat As OlectraChart2D.FormatConstants)
    Chart2D.CopyToClipboard eCopyFormat
End Sub

Public Sub EnableActionMove()
    'Clear any previous ActionMaps, and construct the new ones.
    'Make the rotation constraints inaccessible
    
    ClearEvents
    
    With Chart2D.ActionMaps
        .add WM_LBUTTONDOWN, 0, 0, oc2dActionModifyStart
        .add WM_MOUSEMOVE, MK_LBUTTON, 0, oc2dActionTranslate
        .add WM_LBUTTONUP, 0, 0, oc2dActionModifyEnd
    End With

    mCurrentAction = oc2dActionTranslate
    
End Sub

Public Sub EnableActionNone()
    'Clear any previous ActionMaps.
    'Make the rotation constraints inaccessible
    
    ClearEvents
    mCurrentAction = oc2dActionNone

End Sub

Public Sub EnableActionRotate()
    'Clear any previous ActionMaps, and construct the new ones.
    'Make the rotation constraints accessible
    
    ClearEvents
    
    With Chart2D.ActionMaps
        .add WM_LBUTTONDOWN, 0, 0, oc2dActionModifyStart
        .add WM_MOUSEMOVE, MK_LBUTTON, 0, oc2dActionRotate
        .add WM_LBUTTONUP, 0, 0, oc2dActionModifyEnd
    End With
    
    mCurrentAction = oc2dActionRotate
    
End Sub

Public Sub EnableActionScale()
    'Clear any previous ActionMaps, and construct the new ones.
    'Make the rotation constraints inaccessible
        
    ClearEvents
    
    With Chart2D.ActionMaps
        .add WM_LBUTTONDOWN, 0, 0, oc2dActionModifyStart
        .add WM_MOUSEMOVE, MK_LBUTTON, 0, oc2dActionScale
        .add WM_LBUTTONUP, 0, 0, oc2dActionModifyEnd
    End With
    
    mCurrentAction = oc2dActionScale

End Sub

Public Sub EnableActionZoom()
    'Clear any previous ActionMaps, and construct the new ones.
    'Make the rotation constraints inaccessible
    
    ClearEvents
    
    With Chart2D.ActionMaps
        .add WM_LBUTTONDOWN, 0, 0, oc2dActionZoomStart
        .add WM_MOUSEMOVE, MK_LBUTTON, 0, oc2dActionZoomUpdate
        .add WM_LBUTTONUP, 0, 0, oc2dActionZoomEnd
        .add WM_KEYDOWN, MK_LBUTTON, VK_ESCAPE, oc2dActionZoomCancel
    End With
    
    mCurrentAction = oc2dActionZoomStart
    
    SetZoomMode False

End Sub

Public Sub EnableDisableDelayUpdating(blnEnableOverride As Boolean)
    mMasterOverrideOnIsBatch = blnEnableOverride
    
    If Not mMasterOverrideOnIsBatch Then
        Chart2D.IsBatched = False
    Else
        Chart2D.IsBatched = True
    End If
End Sub

Public Sub EnableDisablePropertyPageView(blnAllowView As Boolean)
    ' Control whether the user can bring up the property pages at run-time
    Chart2D.AllowUserChanges = blnAllowView
End Sub

Public Function GetCurrentGroupNumber() As Integer
    GetCurrentGroupNumber = mCurrentGroup
End Function

Public Function GetCurrentSeriesNumber() As Integer
    GetCurrentSeriesNumber = mCurrentSeries
End Function

Public Function GetXAxisRangeMin() As Double
    GetXAxisRangeMin = Chart2D.ChartArea.Axes("X").Min
End Function

Public Function GetXAxisRangeMax() As Double
    GetXAxisRangeMax = Chart2D.ChartArea.Axes("X").Max
End Function

Public Function GetYAxisRangeMin() As Double
    GetYAxisRangeMin = Chart2D.ChartArea.Axes("Y").Min
End Function

Public Function GetYAxisRangeMax() As Double
    GetYAxisRangeMax = Chart2D.ChartArea.Axes("Y").Max
End Function

Public Function GetSeriesCount() As Integer
    GetSeriesCount = mSeriesCount
End Function

Private Sub InitializeData()
    Const NUM_PNTS = 9
    Dim XData(1 To NUM_PNTS) As Double
    Dim YData(1 To NUM_PNTS) As Double
    
    ' Default data -- represents a doubly charged peak with Cl and Br atoms (C4H8BrCl2)
    XData(1) = 205
    XData(2) = 205.5
    XData(3) = 206
    XData(4) = 206.5
    XData(5) = 207
    XData(6) = 207.5
    XData(7) = 208
    XData(8) = 208.5
    XData(9) = 209

    YData(1) = 62.01
    YData(2) = 2.76
    YData(3) = 100
    YData(4) = 4.44
    YData(5) = 44.97
    YData(6) = 2
    YData(7) = 6.2
    YData(8) = 0.27
    YData(9) = 0
    
    ' Set current Series as 1
    SetCurrentGroup 1
    SetSeriesCount 1
    SetCurrentSeries 1
    
    SetSeriesDataPointCount 1, NUM_PNTS
    SetDataX 1, XData(), False
    SetDataY 1, YData(), False
    
    
End Sub

Private Sub InitializeGraphControl()
    
    ClearEvents
    
    ' Set the member variables
    mCurrentGroup = 1
    mCurrentSeries = 1
    mSeriesCount = 1
    
    ' Populate the combo box
    cboMouseAction.AddItem "Zoom"
    cboMouseAction.AddItem "Move (space+drag)"
    cboMouseAction.AddItem "Scale"
    cboMouseAction.ListIndex = 0
     
    ' Disallow user changing graph properties using Right-Click
    EnableDisablePropertyPageView False
    
    ' Enable showing location of nearest data point
    SetLocationDisplay True, False
    
    ' Enable batch chart updating
    With Chart2D
        .IsBatched = True
        .Legend.IsShowing = False
    
        SetChartType oc2dTypeBar
        
        .ChartArea.Axes("x").AnnotationMethod = oc2dAnnotatePointLabels
        SetBarWidth 100
    End With
    
    InitializeData
    
    ' Make sure graph is fully zoomed out
    ZoomOutFull
    
    ' Initialize the zoom history stack
    
    With Chart2D
        With .ChartArea
            ZoomHistoryPush .Axes("x").Min, .Axes("x").Max, .Axes("y").Min, .Axes("y").Max
        End With
        
        ' 3D view options
'        .ChartArea.View3D.Depth = 10
'        .ChartArea.View3D.Elevation = 10
'        .ChartArea.View3D.Rotation = 10
'
'        .Interior.BackgroundColor = RGB(&HFA, &HFA, &HD2)                       'LightGoldenRodYellow
'        .Interior.ForegroundColor = RGB(&H0, &H0, &H0)                          'Black
'        .ChartArea.PlotArea.Interior.BackgroundColor = RGB(&HD2, &HB4, &H8C)    'Tan
        
        ' Return chart to immediate-update mode
        .IsBatched = False
    End With
    
    UserControl.width = 500
    UserControl.Height = 300
End Sub

Private Sub LogEventToConsole(Msg As Integer, modf As Integer, key As Integer, action As Integer)
    'Debugging output for the ActionMaps.  Uncomment in chart_Click to view these.
    
    Dim Output As String
    
    Select Case Msg
        Case WM_MOUSEMOVE
            Output = "MouseMove"
        Case WM_LBUTTONDOWN
            Output = "LButtonDown"
        Case WM_LBUTTONUP
            Output = "LButtonUp"
        Case WM_LBUTTONDBLCLK
            Output = "LButtonDbl"
        Case WM_RBUTTONDOWN
            Output = "RButtonDown"
        Case WM_RBUTTONUP
            Output = "RButtonUp"
        Case WM_RBUTTONDBLCLK
            Output = "RButtonDbl"
        Case WM_MBUTTONDOWN
            Output = "MButtonDown"
        Case WM_MBUTTONUP
            Output = "MButtonUp"
        Case WM_MBUTTONDBLCLK
            Output = "MButtonDbl"
        Case WM_KEYDOWN
            Output = "KeyDown"
        Case WM_KEYUP
            Output = "KeyUp"
    End Select
    
    Output = Output & ", "
    If (modf And MK_ALT) = MK_ALT Then Output = Output & "ALT+"
    If (modf And MK_MBUTTON) = MK_MBUTTON Then Output = Output & "MBUTTON+"
    If (modf And MK_CONTROL) = MK_CONTROL Then Output = Output & "CONTROL+"
    If (modf And MK_SHIFT) = MK_SHIFT Then Output = Output & "SHIFT+"
    If (modf And MK_RBUTTON) = MK_RBUTTON Then Output = Output & "RBUTTON+"
    If (modf And MK_LBUTTON) = MK_LBUTTON Then Output = Output & "LBUTTON+"
        
    Output = Output & ", "
    
    If key <> 0 Then Output = Output & Chr(key)
    
    Output = Output & " = "
    
    Select Case action
        Case oc2dActionNone
            Output = Output & "None"
        Case oc2dActionModifyStart
            Output = Output & "ModifyStart"
        Case oc2dActionModifyEnd
            Output = Output & "ModifyEnd"
        Case oc2dActionRotate
            Output = Output & "Rotate"
        Case oc2dActionScale
            Output = Output & "Scale"
        Case oc2dActionTranslate
            Output = Output & "Translate"
        Case oc2dActionZoomStart
            Output = Output & "ZoomStart"
        Case oc2dActionZoomUpdate
            Output = Output & "ZoomUpdate"
        Case oc2dActionZoomEnd
            Output = Output & "ZoomEnd"
        Case oc2dActionZoomCancel
            Output = Output & "ZoomCancel"
        Case oc2dActionProperties
            Output = Output & "Properties Page"
        Case oc2dActionReset
            Output = Output & "Reset"
    End Select
    
    Debug.Print Output

End Sub

Public Sub PopulateSymbolStyleComboBox(ByRef cboThisComboBox As ComboBox)
    
    ' Note that the symbol styles (olectrachart2d.ShapeConstants) start with 1 and not 0
    With cboThisComboBox
        .AddItem "None"
        .AddItem "Dot"
        .AddItem "Box"
        .AddItem "Triangle"
        .AddItem "Diamond"
        .AddItem "Star"
        .AddItem "Vertical Line"
        .AddItem "Horizontal line"
        .AddItem "Cross"
        .AddItem "Circle"
        .AddItem "Square"
        .ListIndex = 0
    End With
End Sub

Private Sub PositionControls()
    Dim lngDesiredVal As Long
    
    lngDesiredVal = ScaleWidth - 10
    If lngDesiredVal < 5 Then lngDesiredVal = 5
    Chart2D.width = lngDesiredVal
    
    lngDesiredVal = ScaleHeight - 120 - fraControls.Height
    If lngDesiredVal < 5 Then lngDesiredVal = 5
    Chart2D.Height = lngDesiredVal

    lngDesiredVal = ScaleHeight - fraControls.Height
    If lngDesiredVal < 0 Then lngDesiredVal = 0
    fraControls.Top = lngDesiredVal

    fraControls.Left = 120
End Sub

Public Sub SetLocationDisplay(blnShowLocNearestDataPoint As Boolean, blnShowLocAtAllPositions As Boolean)
    mShowLocationDataPoint = blnShowLocNearestDataPoint
    mShowLocationEverywhere = blnShowLocAtAllPositions

End Sub

Private Sub ReportCursorLocation(x As Single, y As Single)
    ' Update value of nearest data point in lblLocation
    
    Dim Distance As Long, Point As Long, series As Long
    Dim XVal As Double, YVAl As Double
    
    Dim Region As OlectraChart2D.RegionConstants
    Dim XPixel As Long, YPixel As Long
    Dim strPointCoords As String, strLocationCoords As String, strOutput As String
    Const DECIMAL_PLACES = 3
    
    XPixel = x / Screen.TwipsPerPixelX
    YPixel = y / Screen.TwipsPerPixelY
    
    With Chart2D.ChartGroups(mCurrentGroup)
        Region = .CoordToDataIndex(XPixel, YPixel, oc2dFocusXY, series, Point, Distance)
        
        'strLocationDescription = Distance & " units away from point " & Point & " on series " & series
        
        If Point > 0 Then
            strPointCoords = FormatNumber(.Data.x(series, Point), DECIMAL_PLACES) & ", " & FormatNumber(.Data.y(series, Point), DECIMAL_PLACES)
        Else
            strPointCoords = ""
        End If
        
        Region = .CoordToDataCoord(XPixel, YPixel, XVal, YVAl)
        
        If XVal < ocHugeValue And YVAl < ocHugeValue Then
            strLocationCoords = "Point " & strPointCoords & vbCrLf & "(" & FormatNumber(XVal, DECIMAL_PLACES) & ", " & FormatNumber(YVAl, DECIMAL_PLACES) & ")"
        Else
            strLocationCoords = ""
        End If
    End With
    
    If mShowLocationDataPoint Then
        strOutput = strPointCoords
    End If
    
    If mShowLocationEverywhere Then
        If Len(strOutput) > 0 Then strOutput = strOutput & vbCrLf
        strOutput = strOutput & strLocationCoords
    End If

    lblLocation = strOutput
End Sub

Public Sub ResetGraph()
    Chart2D.CallAction oc2dActionReset, 0, 0
End Sub

Public Function SaveChartPictureToFile(blnSaveAsPNG As Boolean, strFilePath As String) As Boolean
    ' If blnSaveAsPNG = True, then saves a PNG file
    ' If blnSaveAsPNG = False, then saves a JPG file
    
    ' Returns True if success, False if failure

On Error GoTo SaveChartPictureToFileErrorHandler
        
    If blnSaveAsPNG Then
        Chart2D.SaveImageAsPng strFilePath, False
    Else
        Chart2D.SaveImageAsJpeg strFilePath, 90, False, True, False
    End If
    SaveChartPictureToFile = True
    
    Exit Function

SaveChartPictureToFileErrorHandler:
    SaveChartPictureToFile = False
    
End Function

Public Sub SetBarWidth(lngWidthPercentage As Long)
    Chart2D.ChartArea.Bar.ClusterWidth = lngWidthPercentage
End Sub

Public Sub SetChartType(NewChartType As OlectraChart2D.ChartTypeConstants, Optional GroupNumber As Integer = 1)
    ' Usually oc2dTypeBar or oc2dTypePlot
    
    If GroupNumber = 1 Or GroupNumber = 2 Then
        Chart2D.ChartGroups(mCurrentGroup).ChartType = NewChartType
    End If

End Sub

Public Sub SetDepth(intNewDepthVal As Integer)
    
    If intNewDepthVal > 100 Then
        intNewDepthVal = 100
    ElseIf intNewDepthVal < 0 Then
        intNewDepthVal = 0
    End If
    
    Chart2D.ChartArea.View3D.Depth = intNewDepthVal

End Sub

Public Sub SetCurrentGroup(intCurrentGroup As Integer)
    If intCurrentGroup < 1 Then intCurrentGroup = 1
    If intCurrentGroup > MAX_GROUPS Then intCurrentGroup = MAX_GROUPS
    mCurrentGroup = intCurrentGroup
End Sub

Public Sub SetCurrentSeries(intCurrentSeries As Integer)
    If intCurrentSeries < 1 Then intCurrentSeries = 1
    If intCurrentSeries > MAX_SERIES_COUNT Then intCurrentSeries = MAX_SERIES_COUNT
    mCurrentSeries = intCurrentSeries
End Sub

Public Sub SetLabelXAxis(strLabel As String, Optional lngFontSize As Long = 10)
    With Chart2D.ChartArea.Axes("X")
        .Title = strLabel
        .Font.Size = lngFontSize
    End With
End Sub

Public Sub SetLabelYAxis(strLabel As String, Optional lngFontSize As Long = 10)
    With Chart2D.ChartArea.Axes("Y")
        .Title = strLabel
        .Font.Size = lngFontSize
    End With
End Sub

Public Sub SetLabelGraphTitle(strLabel As String, Optional lngFontSize As Long = 10)
    With Chart2D.Header
        .Text = strLabel
        .Border = oc2dBorderNone
        .Font.Size = lngFontSize
    End With
End Sub

Public Sub SetSeriesCount(intSeriesCount As Integer)
    If intSeriesCount < 0 Then intSeriesCount = 0
    If intSeriesCount > MAX_SERIES_COUNT Then intSeriesCount = MAX_SERIES_COUNT
    
    mSeriesCount = intSeriesCount
    Chart2D.ChartGroups(mCurrentGroup).Data.NumSeries = mSeriesCount
End Sub

Public Sub SetSeriesDataPointCount(mSeriesNumber As Integer, lngNumberOfPoints As Long)
    If mSeriesNumber < 1 Or mSeriesNumber > MAX_SERIES_COUNT Then mSeriesNumber = 1
    Chart2D.ChartGroups(mCurrentGroup).Data.NumPoints(mSeriesNumber) = lngNumberOfPoints
End Sub

Private Sub AssureOneBasedArray(ArrayIn() As Double, ArrayOut() As Double)
    Dim lngUpperIndex As Long, lngIndex As Long
    
    lngUpperIndex = UBound(ArrayIn)
    
    If LBound(ArrayIn) = 0 Then
        ReDim ArrayOut(1 To lngUpperIndex + 1)
        For lngIndex = 1 To lngUpperIndex + 1
            ArrayOut(lngIndex) = ArrayIn(lngIndex - 1)
        Next lngIndex
    Else
        ReDim ArrayOut(1 To lngUpperIndex)
        For lngIndex = 1 To lngUpperIndex
            ArrayOut(lngIndex) = ArrayIn(lngIndex)
        Next lngIndex
    End If
End Sub

Public Sub SetDataX(mSeriesNumber As Integer, XDataOneBased1DArray() As Double, Optional blnEnableAutoScaleTimer As Boolean = True)
    
On Error GoTo SetDataXErrorHandler
    
    If mSeriesNumber < 1 Or mSeriesNumber > MAX_SERIES_COUNT Then mSeriesNumber = 1
    If mCurrentGroup < 1 Or mCurrentGroup > MAX_GROUPS Then mCurrentGroup = 1
    
'    ' Olectra Chart requires that the data Array() be 1 based and not 0 based
'    ' Thus, make sure data is in correct format
'    Dim ArrayToWrite() As Double
'    AssureOneBasedArray XDataOneBased1DArray(), ArrayToWrite()
'    Chart2D.ChartGroups(mCurrentGroup).Data.CopyXVectorIn mSeriesNumber, ArrayToWrite()
     
    ' Must set the Layout to oc2dDataGeneral to allow series to have different X-axis values
    Chart2D.ChartGroups(mCurrentGroup).Data.Layout = oc2dDataGeneral
    Chart2D.ChartGroups(mCurrentGroup).Data.CopyXVectorIn mSeriesNumber, XDataOneBased1DArray()
     
    If blnEnableAutoScaleTimer Then
       tmrAutoScale.Enabled = True
       tmrAutoScale.Interval = 50
    End If

    Exit Sub

SetDataXErrorHandler:
    Debug.Print "Error in SetDataX: & "; Err.Description
    Debug.Assert False
    Resume Next
End Sub

Public Sub SetDataY(mSeriesNumber As Integer, YDataOneBased1DArray() As Double, Optional blnEnableAutoScaleTimer As Boolean = True)
    
On Error GoTo SetDataYErrorHandler
    
    If mSeriesNumber < 1 Or mSeriesNumber > MAX_SERIES_COUNT Then mSeriesNumber = 1
    If mCurrentGroup < 1 Or mCurrentGroup > MAX_GROUPS Then mCurrentGroup = 1
    
    ' Olectra Chart requires that the data Array() 1 based and not 0 based
    ' Thus, make sure data is in correct format
    Chart2D.ChartGroups(mCurrentGroup).Data.CopyYVectorIn mSeriesNumber, YDataOneBased1DArray()

    If blnEnableAutoScaleTimer Then
       tmrAutoScale.Enabled = True
       tmrAutoScale.Interval = 50
    End If
    
    Exit Sub
    
SetDataYErrorHandler:
    Debug.Print "Error in SetDataY: & "; Err.Description
    Debug.Assert False
    Resume Next

End Sub

Public Sub SetStyleDataFill(lngColor As Long, ePattern As OlectraChart2D.FillPatternConstants)
    If mCurrentGroup < 1 Or mCurrentGroup > 2 Then mCurrentGroup = 1
    With Chart2D.ChartGroups(mCurrentGroup).Styles(mCurrentSeries).Fill
        .Color = lngColor
        .Pattern = ePattern
    End With
End Sub

Public Sub SetStyleDataLine(lngLineColor As Long, eLinePattern As OlectraChart2D.LinePatternConstants, lngWidth As Long)
    If mCurrentGroup < 1 Or mCurrentGroup > 2 Then mCurrentGroup = 1
    With Chart2D.ChartGroups(mCurrentGroup).Styles(mCurrentSeries).Line
        .Color = lngLineColor
        .Pattern = eLinePattern
        .width = lngWidth
    End With
End Sub


Public Sub SetStyleDataSymbol(lngSymbolColor As Long, eSymbolShape As OlectraChart2D.ShapeConstants, lngSize As Long)
    If mCurrentGroup < 1 Or mCurrentGroup > 2 Then mCurrentGroup = 1
    With Chart2D.ChartGroups(mCurrentGroup).Styles(mCurrentSeries).Symbol
        .Color = lngSymbolColor
        .Shape = eSymbolShape
        .Size = lngSize
    End With
End Sub

Public Sub SetXAxisAnnotationMethod(eAnnotationmethod As OlectraChart2D.AnnotationMethodConstants)
    Chart2D.ChartArea.Axes("X").AnnotationMethod = eAnnotationmethod
End Sub

Public Sub SetYAxisAnnotationMethod(eAnnotationmethod As OlectraChart2D.AnnotationMethodConstants)
    Chart2D.ChartArea.Axes("Y").AnnotationMethod = eAnnotationmethod
End Sub

Public Sub SetXAxisAnnotationPlacement(eAnnotationPlacement As OlectraChart2D.AnnotationPlacementConstants)
    Chart2D.ChartArea.Axes("X").AnnotationPlacement = eAnnotationPlacement
End Sub

Public Sub SetYAxisAnnotationPlacement(eAnnotationPlacement As OlectraChart2D.AnnotationPlacementConstants)
    Chart2D.ChartArea.Axes("Y").AnnotationPlacement = eAnnotationPlacement
End Sub

Public Sub SetYAxisOriginVsXAxis(dblOriginLocation As Double)
    With Chart2D
        .IsBatched = True
        With .ChartArea.Axes("X")
            .Origin = dblOriginLocation
            .AnnotationPlacement = oc2dAnnotateAuto
        End With
        If Not mMasterOverrideOnIsBatch Then .IsBatched = False
    End With

End Sub

Public Sub SetXAxisLabelFont(intFontSize As Integer, Optional strFontName As String = "Arial", Optional blnBold As Boolean = False)
    With Chart2D.ChartArea.Axes("X").Font
        .Size = intFontSize
        .Name = strFontName
        .Bold = blnBold
    End With
End Sub

Public Sub SetYAxisLabelFont(intFontSize As Integer, Optional strFontName As String = "Arial", Optional blnBold As Boolean = False)
    With Chart2D.ChartArea.Axes("Y").Font
        .Size = intFontSize
        .Name = strFontName
        .Bold = blnBold
    End With
End Sub

Public Sub SetXAxisLabelFormatNumber(intDecimalPlaces As Integer, Optional blnUseCommaSeparator As Boolean = False)
    With Chart2D.ChartArea.Axes("X").LabelFormat
        .Category = oc2dCategoryNumber
        With .Number
            .DecimalPlaces = intDecimalPlaces
            .UseSeparators = blnUseCommaSeparator
        End With
    End With
End Sub

Public Sub SetYAxisLabelFormatNumber(intDecimalPlaces As Integer, Optional blnUseCommaSeparator As Boolean = False)
    With Chart2D.ChartArea.Axes("Y").LabelFormat
        .Category = oc2dCategoryNumber
        With .Number
            .DecimalPlaces = intDecimalPlaces
            .UseSeparators = blnUseCommaSeparator
        End With
    End With
End Sub

Public Sub SetXAxisLabelFormatScientific(intDecimalPlaces As Integer, Optional blnUseSmallExponent As Boolean = False)
    With Chart2D.ChartArea.Axes("X").LabelFormat
        .Category = oc2dCategoryScientific
        With .Scientific
            .DecimalPlaces = intDecimalPlaces
            .UseSmallExponent = blnUseSmallExponent
        End With
    End With
End Sub

Public Sub SetYAxisLabelFormatScientific(intDecimalPlaces As Integer, Optional blnUseSmallExponent As Boolean = False)
    With Chart2D.ChartArea.Axes("Y").LabelFormat
        .Category = oc2dCategoryScientific
        With .Scientific
            .DecimalPlaces = intDecimalPlaces
            .UseSmallExponent = blnUseSmallExponent
        End With
    End With
End Sub

Public Sub SetXAxisGridlines(eGridPattern As OlectraChart2D.LinePatternConstants)
    With Chart2D.ChartArea.Axes("X").MajorGrid
        .Style.Pattern = eGridPattern
        .Style.width = 1
        .Spacing.IsDefault = True
    End With
End Sub

Public Sub SetYAxisGridlines(eGridPattern As OlectraChart2D.LinePatternConstants)
    With Chart2D.ChartArea.Axes("Y").MajorGrid
        .Style.Pattern = eGridPattern
        .Style.width = 1
        .Spacing.IsDefault = True
    End With
End Sub

Public Sub SetXAxisTickSpacing(dblSpacing As Double, blnResetToDefault As Boolean)
    With Chart2D
        .IsBatched = True
        With .ChartArea.Axes("X")
            If blnResetToDefault Then
                .TickSpacing.IsDefault = True
            Else
                .TickSpacing.IsDefault = False
                .TickSpacing.Value = dblSpacing
            End If
        End With
        
        If Not mMasterOverrideOnIsBatch Then .IsBatched = False
    End With
End Sub

Public Sub SetYAxisTickSpacing(dblSpacing As Double, blnResetToDefault As Boolean)
    With Chart2D
        .IsBatched = True
        With .ChartArea.Axes("Y")
            If blnResetToDefault Then
                .TickSpacing.IsDefault = True
            Else
                .TickSpacing.IsDefault = False
                .TickSpacing.Value = dblSpacing
            End If
        End With
        
        If Not mMasterOverrideOnIsBatch Then .IsBatched = False
    End With
End Sub

Public Sub SetXRange(dblRangeMin As Double, dblRangeMax As Double)
            
    With Chart2D
        .IsBatched = True
        With .ChartArea.Axes("X")
            .Min.IsDefault = False
            .Max.IsDefault = False
            
            .Min.Value = dblRangeMin
            .Max.Value = dblRangeMax
        End With
        If Not mMasterOverrideOnIsBatch Then .IsBatched = False
    End With

End Sub

Public Sub SetYRange(dblRangeMin As Double, dblRangeMax As Double)
            
    With Chart2D
        .IsBatched = True
        With .ChartArea.Axes("Y")
            .Min.IsDefault = False
            .Max.IsDefault = False
            
            .Min.Value = dblRangeMin
            .Max.Value = dblRangeMax
        End With
        If Not mMasterOverrideOnIsBatch Then .IsBatched = False
    End With

End Sub

Public Sub SetZoomMode(blnGraphical As Boolean)
    ' if blnGraphical = True then graphically zooms, otherwise, performs regular axis-based zoom
                
    If blnGraphical Then
        Chart2D.ActionMaps.Remove WM_LBUTTONUP, 0, 0       'oc2dActionAxisBound
        Chart2D.ActionMaps.add WM_LBUTTONUP, 0, 0, oc2dActionZoomEnd
    Else
        Chart2D.ActionMaps.Remove WM_LBUTTONUP, 0, 0       'oc2dActionZoomEnd
        Chart2D.ActionMaps.add WM_LBUTTONUP, 0, 0, oc2dActionZoomAxisEnd
    End If

End Sub

Private Sub ZoomHistoryPush(XAxisMin As Double, XAxisMax As Double, YAxisMin As Double, YAxisMax As Double)
    ' Push the new zoom limits onto the history stack
    
    Dim intIndex As Integer
    Dim blnAbortHistoryUpdate As Boolean
    
    ' Make sure new values do not match the most recent values in the history
    ' If they do, there's no point in adding this to the history
    With ZoomHistory(0, mCurrentSeries)
        If .XAxisMin = XAxisMin And .XAxisMax = XAxisMax And .YAxisMin = YAxisMin And .YAxisMax = YAxisMax Then
            blnAbortHistoryUpdate = True
        End If
    End With
    
    If Not blnAbortHistoryUpdate Then
        ' Shift values on stack up by one position
        For intIndex = ZOOM_HISTORY_MAX - 1 To 1 Step -1
            ZoomHistory(intIndex, mCurrentSeries) = ZoomHistory(intIndex - 1, mCurrentSeries)
        Next intIndex
        
        ' Add the new zoom settings to the stack
        With ZoomHistory(0, mCurrentSeries)
            .XAxisMin = XAxisMin
            .XAxisMax = XAxisMax
            .YAxisMin = YAxisMin
            .YAxisMax = YAxisMax
        End With
               
    End If
    
End Sub

Public Sub ZoomOutFull()
    Chart2D.IsBatched = True
    With Chart2D.ChartArea
        .Axes("x").Min.IsDefault = True
        .Axes("x").Max.IsDefault = True
        .Axes("y").Min.IsDefault = True
        .Axes("y").Max.IsDefault = True
    End With
    
    Chart2D.IsBatched = False

End Sub

Private Function ZoomHistoryValueValid(ThisZoomHistoryValue As udtPlotRange) As Boolean
    With ThisZoomHistoryValue
        If .XAxisMin = 0 And .XAxisMax = 0 And .YAxisMin = 0 And .YAxisMax = 0 Then
            ZoomHistoryValueValid = False
        Else
            ZoomHistoryValueValid = True
        End If
    End With
End Function

Public Sub ZoomToPreviousHistoryValue()
    Dim intIndex As Integer

    ' First see if the previous value has valid data

    If ZoomHistoryValueValid(ZoomHistory(1, mCurrentSeries)) Then
        Chart2D.IsBatched = True
        With Chart2D.ChartArea.Axes("X")
            .Min = ZoomHistory(1, mCurrentSeries).XAxisMin
            .Max = ZoomHistory(1, mCurrentSeries).XAxisMax
        End With
        
        With Chart2D.ChartArea.Axes("Y")
            .Min = ZoomHistory(1, mCurrentSeries).YAxisMin
            .Max = ZoomHistory(1, mCurrentSeries).YAxisMax
        End With
        
        Chart2D.IsBatched = False
    
        ' Shift the values in the history
        For intIndex = 0 To ZOOM_HISTORY_MAX - 2
            ZoomHistory(intIndex, mCurrentSeries) = ZoomHistory(intIndex + 1, mCurrentSeries)
        Next intIndex
        
    End If
    
    
End Sub

Public Sub SetGraphRange(XAxisMin As Double, XAxisMax As Double, YAxisMin As Double, YAxisMax As Double)
    
    Chart2D.IsBatched = True
    With Chart2D.ChartArea
        .Axes("X").Min = XAxisMin
        .Axes("X").Max = XAxisMax
        .Axes("Y").Min = YAxisMin
        .Axes("Y").Max = YAxisMax
    End With
    If Not mMasterOverrideOnIsBatch Then Chart2D.IsBatched = False

End Sub
Private Sub cboMouseAction_Click()
    ComboBoxSetAction
End Sub

Private Sub chart2d_Click()
'    'Uncomment these lines to debug the action maps.
'
'    Dim objEvent As ActionMap
'
'    For Each objEvent In chart2d.ActionMaps
'        LogEventToConsole objEvent.Message, objEvent.Modifier, objEvent.Keycode, objEvent.action
'    Next
End Sub

Private Sub Chart2D_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        ' Zoom to previous zoom settings in history
        ZoomToPreviousHistoryValue
    ElseIf Button = vbLeftButton Then
        
    End If
End Sub

Private Sub Chart2d_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ReportCursorLocation x, y
End Sub

Private Sub Chart2D_ZoomAxisModify(XAxisMin As Double, XAxisMax As Double, YAxisMin As Double, YAxisMax As Double, Y2AxisMin As Double, Y2AxisMax As Double, IsOK As Boolean)
    
'    If (mCurrentAction = oc2dActionZoomStart) And IsOK Then
        ' Graph zoom has just changed
        ' Update the zoom history stack and auto-scale Y if necessary
        ZoomHistoryPush XAxisMin, XAxisMax, YAxisMin, YAxisMax
        mZoomChanged = True
'    End If

End Sub

Private Sub cmdZoomOut_Click()
    'ResetGraph
    ZoomOutFull
    
    Chart2D.SetFocus
End Sub

Private Sub tmrAutoScale_Timer()
    If chkAutoScaleY = vbChecked And mZoomChanged Then
        AutoScaleYNow
        mZoomChanged = False
    End If
End Sub

Private Sub UserControl_Initialize()
    InitializeGraphControl
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
            
' Note: KeyPreview must be enabled on this form for this code to execute

'    Debug.Print "Keycode = " & KeyCode & " and shift = " & Shift
    If Shift = 2 Then
        ' Ctrl pressed
        If KeyCode = 65 Then
            ' Ctrl+A pressed; Reset graph
            ResetGraph
        End If
    ElseIf KeyCode = vbKeySpace Then
        mTemporaryActionMode = True
        EnableActionMove
    End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If mTemporaryActionMode Then
        ComboBoxSetAction
        mTemporaryActionMode = False
    End If
End Sub

Private Sub UserControl_Resize()
    PositionControls
End Sub
