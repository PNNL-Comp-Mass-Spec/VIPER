VERSION 5.00
Begin VB.UserControl LaSpots 
   ClientHeight    =   3570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5520
   ScaleHeight     =   3570
   ScaleWidth      =   5520
   ToolboxBitmap   =   "LaSpots.ctx":0000
   Begin VB.PictureBox picD 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   205
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
   Begin VB.Menu mnuF 
      Caption         =   "Function"
      Visible         =   0   'False
      Begin VB.Menu mnuFSpotsNearCursor 
         Caption         =   "Spots near Cursor"
         Begin VB.Menu mnuSpotInfo 
            Caption         =   "Spot 1"
            Index           =   0
         End
      End
      Begin VB.Menu mnuFInfoOnSpotsNearCursor 
         Caption         =   "Show Info on Spots Near Cursor"
      End
      Begin VB.Menu mnuFInfoOnSelection 
         Caption         =   "Show Info on Selected Spots"
      End
      Begin VB.Menu mnuFSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFSelectAll 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnuFClearSelection 
         Caption         =   "Cl&ear Selection"
      End
      Begin VB.Menu mnuFSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFViewCoo 
         Caption         =   "View Coordinates"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFSSzIncrease 
         Caption         =   "Increase Spot Size"
      End
      Begin VB.Menu mnuFSSzDecrease 
         Caption         =   "Decrease Spot Size"
      End
      Begin VB.Menu mnuFSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDFCopy 
         Caption         =   "&Copy Picture"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "LaSpots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'function spot selection user control
'--------------------------------------------------------
'Originally created in 2002 by Nikola Tolic
'Functionality greatly expanded in January 2003 by Matthew Monroe
'--------------------------------------------------------
Option Explicit

Private Const MSG_TITLE = "2D Displays - Spots"

Private Const MAXDATACOUNT = 1000000              'maximum data points in this display
Private Const MAX_SPOT_MATCHES_TO_TRACK = 1000

Private Const mSpotsShapeCount = 6
Public Enum sSpotsShape
    sCircle = 0
    sRectangle = 1
    sRoundRectangle = 2
    sStar = 3
    sEmptyRectangle = 4
    sTriangleWithExtents = 5
End Enum

Public Enum sSpotsView
    sFixedWindow = 0
    sVariableWindow = 1
End Enum

Private Type udtDataPointType
    Description As String
    x As Double                       ' data that has to be drawn
    y As Double
    ExtentInXNeg As Double            ' Means the data actually has a width in X, toward the negative direction
    ExtentInXPos As Double            ' Means the data actually has a width in X, toward the positive direction
    ExtentInYNeg As Double            ' Similar to ExtentInXNeg
    ExtentInYPos As Double            ' Similar to ExtentInXPos
    Intensity As Double               ' Data Intensity
    Selected As Boolean               ' Keeps track of whether a data point is selected or not
End Type

Private Type udtDataPointScaledType
    x As Long
    y As Long
    ExtentInXNeg As Long
    ExtentInXPos As Long
    ExtentInYNeg As Long
    ExtentInYPos As Long
    IntensityHalf As Long       ' Scaled intensity value, divided by 2
    
    ' The actual location of the shape corner's; for Triangles, pretend it's a square
    ' (CornerXNeg,CornerYNeg) defines the X,Y coordinate of the lower-left corner of the shape
    ' (CornerXPos,CornerYPos) defines the X,Y coordinate of the upper-right corner of the shape
    ' These values are filled in FillExtents
    CornerXNeg As Long
    CornerXPos As Long
    CornerYNeg As Long
    CornerYPos As Long
End Type

Private Type udtDataSeriesDataType
    DataCount As Long
    Data() As udtDataPointType                       ' 0-based array
    
    UseExtents As Boolean
    UseIntensity As Boolean
    
    IonSpotShape As sSpotsShape
    RegColor As Long             'regular spot color
    SelColor As Long             'selection spot color

    ' Data scaled to the units of drawing
    mgDataScaled() As udtDataPointScaledType        ' 0-based array
End Type

' Coordinate dimensions
Private Const mLDfX0 = 0
Private Const mLDfY0 = 0
Private mLDfXE As Long
Private mLDfYE As Long

'public properties
Public SwapAxes As Boolean
Public FontWidth As Long
Public FontHeight As Long
Public PlotTitleBottom As String
Public VLabel As String     'label on vertical axis
Public VNumFmt As String    'numerical format for vertical axis
Public HLabel As String     'label on horizontal axis
Public HNumFmt As String    'numerical format for horizontal axis
Public CenterTicksAroundPlotCenter As Boolean       ' When true, the ticks are evenly spaced from the center of the plot
Public ShowTickMarkLabels As Boolean
Public ShowGridLines As Boolean
Public ShowPosition As Boolean
Public IntensityLogScale As Boolean     ' When true, plots the base 10 logs of the intensities, rather than raw intensity
Public CallingFormID As Long            ' Use by the ORF Viewer Forms

Private mViewWindow As sSpotsView       'window can be fixed or variable depending on
                                        'loaded data
                                                                       
'fixed coordinates (based on Series 0)
Private mFixedMinX As Double
Private mFixedMaxX As Double
Private mFixedMinY As Double
Private mFixedMaxY As Double

'current data coordinates (based on Series 0)
Private mCurrMinX As Double
Private mCurrMaxX As Double
Private mCurrMinY As Double
Private mCurrMaxY As Double

Private mCurrMaximumIntensity As Double      ' Only applicable if UseIntensity = True
Private mCurrMinimumIntensity As Double

'working coordinates (based on Series 0)
Private mMinX As Double            'minimum of data range X
Private mMaxX As Double            'maximum of data range X
Private mMinY As Double            'minimum of data range Y
Private mMaxY As Double            'maximum of data range Y

' Labels and tick marks
Private mAutoComputeXAxisTickCount As Boolean
Private mAutoComputeYAxisTickCount As Boolean
Private mXAxisTickCount As Long
Private mYAxisTickCount As Long

' Variables to hold the data
Private mDataSeries() As udtDataSeriesDataType          ' 0-based array
Private mSeriesCount As Integer

Private mMaxSpotSize As Long
Private mHalfSize As Long
Private mQuarterSize As Long
Private mMinSpotSize As Long
Private mHalfMinSpotSize As Long

Private mBorderSize As Long

Private mScaleX As Double         'scales used to draw
Private mScaleY As Double         'on logical window

Private mHotSpotSeriesIndex As Integer  ' data series of hot spot; -1 if none
Private mHotSpotDataIndex As Long       'index of hot spot; -1 if none

Private mCurrX As String          'current coordinates (formatted); based on Series 0
Private mCurrY As String

Private mDefaultSpotShape As sSpotsShape
Private mDefaultRegColor As Long
Private mDefaultSelColor As Long

'coordinate system-viewport coordinates
Private VPX0 As Long
Private VPY0 As Long
Private VPXE As Long
Private VPYE As Long

Private hRegBrush As Long           ' Regular brush
Private hSelBrush As Long           ' Selected brush
Private hRegPen As Long
Private hSelPen As Long
Private hBlackPen As Long
Private hWhitePen As Long            ' Back color pen
Private hGridDashPen As Long        ' Pen for drawing dashed grid
Private paPoints() As POINTAPI      ' Used to calculate device to logical coordinate conversion

Private mSpotMatchCount As Integer
Private mSpotMatches() As Long   ' 2-dimensional array, 0-based in each dimension, recording the series number and spot number of the spots that the mouse is currently over
Private mSpotInfoMenuCountLoaded As Long

Public Function AddSpotsMany(SpotID() As String, SpotX() As Double, SpotY() As Double, SpotIntensity() As Double, Optional lngDataCount As Long = -1, Optional intSeriesIndex As Integer = 0, Optional eSpotShape As sSpotsShape = -1, Optional lngRegularSpotColor As Long = -1, Optional lngSelectedSpotColor As Long = -1) As Boolean
    ' See AddSpotsManyWithExtents() for full explanation
    ' Setting blnYIsIntensity to False will cause SpotIntensity() to be ignored
    
    ' Need to construct an array with Extents of 0, then call AddSpotsManyWithXExtents()
    
    Dim EmptyExtentArray() As Double
    
    Dim lngBaseIndex As Long, lngTopIndex As Long
    
    lngBaseIndex = LBound(SpotID)
    lngTopIndex = UBound(SpotID)
    
    ReDim EmptyExtentArray(lngBaseIndex To lngTopIndex)
    
    AddSpotsMany = AddSpotsManyWithExtents(SpotID(), SpotX(), EmptyExtentArray(), EmptyExtentArray(), SpotY(), EmptyExtentArray(), EmptyExtentArray(), SpotIntensity(), False, True, lngDataCount, intSeriesIndex, eSpotShape, lngRegularSpotColor, lngSelectedSpotColor)
    
End Function

Public Function AddSpotsManyWithExtents(ByRef SpotID() As String, ByRef SpotX() As Double, ByRef ExtentInXNeg() As Double, ByRef ExtentInXPos() As Double, ByRef SpotY() As Double, ByRef ExtentInYNeg() As Double, ByRef ExtentInYPos() As Double, Intensity() As Double, Optional ByVal blnUseExtents As Boolean = False, Optional ByVal blnUseIntensity As Boolean = True, Optional ByVal lngDataCount = -1, Optional ByRef intSeriesIndex As Integer = 0, Optional eSpotShape As sSpotsShape = -1, Optional lngRegularSpotColor As Long = -1, Optional lngSelectedSpotColor As Long = -1) As Boolean
    '-----------------------------------------------------------------------------------------------
    ' Adds data to given series
    ' If lngDataCount = -1, then the number of data points is computed by examining the LBound and UBound of SpotX()
    ' Will update intSeriesIndex to the correct series value if necessary
    ' Returns True if Success
    '
    ' If blnUseExtents = True, then  ExtentInXNeg() and ExtentInXPos() records how wide each spot is,
    '   centered around SpotX(); e.g. if SpotX = 0.5 and ExtentInXNeg = 0.1 and ExtentInXPos = 0.2, then SpotX ranges from 0.4 to 0.7, centered at 0.5
    ' If blnUseExtents = False, then ExtentInXNeg() and ExtentInXPos() are ignored
    ' ExtentInYNeg() and ExtentInYPos() record the extents in Y, similar to above
    ' Intensity() can be used to send spot intensity to the plot
    ' If blnUseIntensity = True, then the intensity is used, and ExtentInYNeg() and ExtentInYPos() are ignored
    ' It is possible to have ExtentInX data and Intensity data by setting blnUseExtents = True and blnUseIntensity = True
    '-----------------------------------------------------------------------------------------------
    
    Dim lngDataCountNew As Long, lngIndexTop As Long
    Dim lngDataIndex As Long, lngFirstSpotIndex As Long, lngDataPointsInInputArrays As Long
    Dim lngCurrentDataIndex As Long
    Dim lngTotalDataCount As Long
    Dim intIndex As Integer
    Dim eResponse As VbMsgBoxResult
    Dim dblMaximumIntensity As Double
    Dim dblMinimumIntensity As Double
    
On Error GoTo AddSpotsManyWithXExtentsErrorHandler

    lngFirstSpotIndex = LBound(SpotX)
    lngDataPointsInInputArrays = UBound(SpotX) - lngFirstSpotIndex + 1
    
    If lngDataCount < 0 Then
        lngDataCountNew = lngDataPointsInInputArrays
    Else
        lngDataCountNew = lngDataCount
        If lngDataCountNew > lngDataPointsInInputArrays Then
            lngDataCountNew = lngDataPointsInInputArrays
        End If
    End If
    
    If lngDataCountNew > MAXDATACOUNT Then
       lngDataCountNew = MAXDATACOUNT
       eResponse = MsgBox("Too many spots; only first " & MAXDATACOUNT & " will be presented.  Continue?", vbYesNoCancel + vbDefaultButton1, MSG_TITLE)
       If eResponse <> vbYes Then Exit Function
    End If
    
    If intSeriesIndex < 0 Then intSeriesIndex = 0
    If intSeriesIndex >= mSeriesCount Then
        intSeriesIndex = mSeriesCount
        mSeriesCount = mSeriesCount + 1
        ReDim Preserve mDataSeries(0 To mSeriesCount - 1)
    End If
    
    With mDataSeries(intSeriesIndex)
        lngIndexTop = lngDataCountNew - 1
        ' Need to fix lngIndexTop when lngDataCountNew <= 0
        If lngIndexTop < 0 Then lngIndexTop = 0
        
        ReDim .Data(lngIndexTop)            ' 0-based array
        ReDim .mgDataScaled(lngIndexTop)    ' 0-based array
        
        If lngDataCountNew > 0 Then
            .DataCount = lngDataCountNew
            For lngDataIndex = lngFirstSpotIndex To lngIndexTop
                lngCurrentDataIndex = lngDataIndex - lngFirstSpotIndex
                
                With .Data(lngDataIndex)
                    .Description = SpotID(lngDataIndex)
                    .x = SpotX(lngDataIndex)
                    .ExtentInXNeg = ExtentInXNeg(lngDataIndex)
                    .ExtentInXPos = ExtentInXPos(lngDataIndex)
                    .y = SpotY(lngDataIndex)
                    .ExtentInYNeg = ExtentInYNeg(lngDataIndex)
                    .ExtentInYPos = ExtentInYPos(lngDataIndex)
                    .Intensity = Intensity(lngDataIndex)
                End With
                
            Next lngDataIndex
        End If
        
        .UseIntensity = blnUseIntensity
        .UseExtents = blnUseExtents
        
        .IonSpotShape = mDefaultSpotShape
        .RegColor = mDefaultRegColor
        .SelColor = mDefaultSelColor
    End With
    
    ' Find the minima and maxima of all of the loaded data (in all series)
    mCurrMinX = 1E+308:     mCurrMaxX = -1E+308
    mCurrMinY = 1E+308:     mCurrMaxY = -1E+308
    
    ' Will also determine the maximum and minimum intensity, but only if .UseIntensity = True
    mCurrMaximumIntensity = -1E+308
    mCurrMinimumIntensity = 1E+308
    
    ' Finally, determine the total data count
    lngTotalDataCount = 0
    
    For intIndex = 0 To mSeriesCount - 1
        With mDataSeries(intIndex)
            dblMaximumIntensity = -1E+308
            dblMinimumIntensity = 1E+308
            lngTotalDataCount = lngTotalDataCount + .DataCount
            
            For lngDataIndex = 0 To .DataCount - 1
                With .Data(lngDataIndex)
                    If .x < mCurrMinX Then mCurrMinX = .x
                    If .x > mCurrMaxX Then mCurrMaxX = .x
                    If .y < mCurrMinY Then mCurrMinY = .y
                    If .y > mCurrMaxY Then mCurrMaxY = .y
                    
                    If mDataSeries(intIndex).UseIntensity Then
                        If .Intensity > dblMaximumIntensity Then dblMaximumIntensity = .Intensity
                        If .Intensity < dblMinimumIntensity Then dblMinimumIntensity = .Intensity
                    End If
                End With
            Next lngDataIndex
            
            If .UseIntensity Then
                ' Only update mCurrMaximumIntensity if this data series has intensity data
                If dblMaximumIntensity > mCurrMaximumIntensity Then mCurrMaximumIntensity = dblMaximumIntensity
                If dblMinimumIntensity < mCurrMinimumIntensity Then mCurrMinimumIntensity = dblMinimumIntensity
            End If
        End With
    Next intIndex
    
    mSpotMatchCount = 0
    ReDim mSpotMatches(lngTotalDataCount, 2)
    
    If mCurrMaximumIntensity < -1E+307 Then mCurrMaximumIntensity = 0
    If mCurrMinimumIntensity > 1E+307 Then mCurrMinimumIntensity = 0
    
    If mSeriesCount = 1 Or mCurrMinimumIntensity = mCurrMaximumIntensity Then
        mCurrMinimumIntensity = 0
    End If
    
    mHotSpotSeriesIndex = -1
    mHotSpotDataIndex = -1
    
    picD.Cls
    
    EstablishCoordinateSystem
    
    SetSpotShapesAndColors eSpotShape, lngRegularSpotColor, lngSelectedSpotColor, intSeriesIndex
    
    Call RefreshPlot
    
    AddSpotsManyWithExtents = True
    Exit Function
    
AddSpotsManyWithXExtentsErrorHandler:
    MsgBox "Error in LaSpots->AddSpotsManyWithXExtents: " & Err.Description
End Function

Private Sub CalculateSpotsPoint(DataPoint As Double, MinDataRange As Double, ByRef DataPointScaled As Long, Scalar As Double)
    DataPointScaled = CLng((DataPoint - MinDataRange) * Scalar)
End Sub

Private Sub CalculateSpotsExtent(blnUseExtents As Boolean, DataExtentNeg As Double, DataExtentPos As Double, ByRef DataExtentScaledNeg As Long, ByRef DataExtentScaledPos As Long, Scalar As Double)
    Dim lngMinExtent As Long
    
    If blnUseExtents Then
        DataExtentScaledNeg = CLng(DataExtentNeg * Scalar)
        DataExtentScaledPos = CLng(DataExtentPos * Scalar)
    Else
        DataExtentScaledNeg = 0
        DataExtentScaledPos = 0
    End If
    
    lngMinExtent = mLDfXE / 75 / 2

    ' Use mLDfXE / 75 / 2 as a lower limit for the extent to guarantee we see something, even if all of the extents are 0
    If DataExtentScaledNeg < mHalfMinSpotSize Then DataExtentScaledNeg = mHalfMinSpotSize
    If DataExtentScaledPos < mHalfMinSpotSize Then DataExtentScaledPos = mHalfMinSpotSize

End Sub

Private Sub CalculateSpots(intSeriesIndex As Integer)
    
    Dim lngIndex As Long
    Dim dblIntensity As Double
    
    If SwapAxes Then
        ' Swapping axes, and computing the coordinates swapped
        mScaleX = mLDfXE / (mMaxY - mMinY)
        mScaleY = mLDfYE / (mMaxX - mMinX)
    Else
        mScaleX = mLDfXE / (mMaxX - mMinX)
        mScaleY = mLDfYE / (mMaxY - mMinY)
    End If
    
    
    With mDataSeries(intSeriesIndex)

        For lngIndex = 0 To .DataCount - 1
            ' Data Points
            
            If SwapAxes Then
                ' Swapping axes, and computing the coordinates swapped
                CalculateSpotsPoint .Data(lngIndex).y, mMinY, .mgDataScaled(lngIndex).x, mScaleX
                CalculateSpotsPoint .Data(lngIndex).x, mMinX, .mgDataScaled(lngIndex).y, mScaleY
            Else
                CalculateSpotsPoint .Data(lngIndex).x, mMinX, .mgDataScaled(lngIndex).x, mScaleX
                CalculateSpotsPoint .Data(lngIndex).y, mMinY, .mgDataScaled(lngIndex).y, mScaleY
            End If
            
            ' Extents
            If SwapAxes Then
                ' Swapping axes, and computing the coordinates swapped
                CalculateSpotsExtent .UseExtents, .Data(lngIndex).ExtentInYNeg, .Data(lngIndex).ExtentInYPos, .mgDataScaled(lngIndex).ExtentInXNeg, .mgDataScaled(lngIndex).ExtentInXPos, mScaleX
                CalculateSpotsExtent .UseExtents, .Data(lngIndex).ExtentInXNeg, .Data(lngIndex).ExtentInXPos, .mgDataScaled(lngIndex).ExtentInYNeg, .mgDataScaled(lngIndex).ExtentInYPos, mScaleY
            Else
                CalculateSpotsExtent .UseExtents, .Data(lngIndex).ExtentInXNeg, .Data(lngIndex).ExtentInXPos, .mgDataScaled(lngIndex).ExtentInXNeg, .mgDataScaled(lngIndex).ExtentInXPos, mScaleX
                CalculateSpotsExtent .UseExtents, .Data(lngIndex).ExtentInYNeg, .Data(lngIndex).ExtentInYPos, .mgDataScaled(lngIndex).ExtentInYNeg, .mgDataScaled(lngIndex).ExtentInYPos, mScaleY
            End If
            
            If .UseIntensity And mCurrMaximumIntensity > 0 Then
                ' Scale intensity to 0 to 1, then multiply by mMaxSpotSize
                dblIntensity = (.Data(lngIndex).Intensity - mCurrMinimumIntensity) / (mCurrMaximumIntensity - mCurrMinimumIntensity)
                If IntensityLogScale Then
                    ' Take the base 10 log of normalized intensity, though need to take times 10 before taking log
                    '  since Log values of numbers below 1 start to grow again
                    dblIntensity = dblIntensity * 10
                    If dblIntensity > 1 Then
                        ' dblIntensity contains an intensity between 1 and 10
                        ' Taking the log will give a number between 0 and 1
                        dblIntensity = Log10(dblIntensity)
                    Else
                        dblIntensity = 0
                    End If
                End If
                
                ' Now scale the intensity
                dblIntensity = dblIntensity * mMaxSpotSize
                
                ' Make sure dblIntensity is at least mHalfMinSpotSize
                If dblIntensity < mHalfMinSpotSize / 2 Then
                    dblIntensity = mHalfMinSpotSize / 2
                End If
                            
                .mgDataScaled(lngIndex).IntensityHalf = CLng(dblIntensity / 2)
            End If
            
        Next lngIndex
    End With
    
End Sub

Public Sub ClearGraphAndData()
    '---------------------------------------------------
    'erase arrays and clean picture; declare object free
    '---------------------------------------------------
    On Error Resume Next
    
    Erase mDataSeries
    ReDim mDataSeries(0 To 0)
    
    With mDataSeries(0)
        .IonSpotShape = mDefaultSpotShape
        .RegColor = mDefaultRegColor
        .SelColor = mDefaultSelColor
    End With
    
    mSeriesCount = 0
    mCurrMinX = 1E+308:     mCurrMaxX = -1E+308
    mCurrMinY = 1E+308:     mCurrMaxY = -1E+308
    mHotSpotSeriesIndex = -1
    mHotSpotDataIndex = -1
    
    picD.Cls
End Sub

Private Sub ComputeOptimalAxisTickCounts()
    ' If mAutoComputeXAxisTickCount = True, then determines a good tick count
    '  value based on the current chart size
    ' Same for mAutoComputeYAxisTickCount
    
    If mAutoComputeXAxisTickCount Then
        mXAxisTickCount = ComputeOptimalTickCount(picD.width)
        ' Need to limit the number of x-axis ticks to avoid overlapping labels
        If mXAxisTickCount > 7 Then mXAxisTickCount = 7
    End If
    
    If mAutoComputeYAxisTickCount Then
        mYAxisTickCount = ComputeOptimalTickCount(picD.Height)
        If mYAxisTickCount > 11 Then mYAxisTickCount = 11
    End If
End Sub

Private Function ComputeOptimalTickCount(lngPictureSize As Long) As Long
    ' Examines the size of the dimension to determine a good number of ticks to have
    ' A good number is 2, 3, 5, 6, or 11; 20, 30, 50, 60, or 110; etc.
    
    Dim lngSuggestedTickCount As Long
    Dim intMagnitudeAdjustmentCount As Integer, intIndex As Integer
    
    lngSuggestedTickCount = Int(lngPictureSize / 500)
    Do While lngSuggestedTickCount > 11
        intMagnitudeAdjustmentCount = intMagnitudeAdjustmentCount + 1
        lngSuggestedTickCount = CLng(lngSuggestedTickCount / 10#)
    Loop
    
    Select Case lngSuggestedTickCount
    Case 0 To 2:    lngSuggestedTickCount = 2
    Case 3 To 5:    lngSuggestedTickCount = 3           ' 3 tick marks total could give: 0.0, 0.5, and 1.0 (for example)
    Case 5 To 11:    lngSuggestedTickCount = 5           ' 5 tick marks total could give: 0.0, 0.25, 0.5, 0.75, and 1.0 (for example)
    End Select
    
    For intIndex = 1 To intMagnitudeAdjustmentCount
        lngSuggestedTickCount = lngSuggestedTickCount * 10
    Next intIndex
    
    ComputeOptimalTickCount = lngSuggestedTickCount
    
End Function

Public Sub CopyFD()
    Dim lngResult As Long
    Dim intSeriesIndex As Integer
    
    Dim OldDC As Long
    Dim hRefDC As Long
    Dim emfDC As Long       'metafile device context
    Dim emfHandle As Long   'metafile handle
    
    Dim iWidthMM As Long
    Dim iHeightMM As Long
    Dim iWidthPels As Long
    Dim iHeightPels As Long
    Dim iMMPerPelX As Double
    Dim iMMPerPelY As Double
    Dim rcRef As Rect           'reference rectangle
    On Error Resume Next
    
    
    hRefDC = picD.hDC
    iWidthMM = GetDeviceCaps(hRefDC, HORZSIZE)
    iHeightMM = GetDeviceCaps(hRefDC, VERTSIZE)
    iWidthPels = GetDeviceCaps(hRefDC, HORZRES)
    iHeightPels = GetDeviceCaps(hRefDC, VERTRES)
    
    iMMPerPelX = (iWidthMM * 100) / iWidthPels
    iMMPerPelY = (iHeightMM * 100) / iHeightPels
    
    rcRef.Top = 0
    rcRef.Left = 0
    rcRef.Bottom = picD.ScaleHeight
    rcRef.Right = picD.ScaleWidth
    'convert to himetric units
    rcRef.Left = rcRef.Left * iMMPerPelX
    rcRef.Top = rcRef.Top * iMMPerPelY
    rcRef.Right = rcRef.Right * iMMPerPelX
    rcRef.Bottom = rcRef.Bottom * iMMPerPelY
    
    emfDC = CreateEnhMetaFile(hRefDC, vbNullString, rcRef, vbNullString)
    OldDC = SaveDC(emfDC)
    
    DGCooSys emfDC, picD.ScaleWidth, picD.ScaleHeight
    DGDrawCooSys emfDC
    DrawLabelsAndTickMarks emfDC
    
    For intSeriesIndex = 0 To mSeriesCount - 1
        DrawSpots emfDC, intSeriesIndex
        DrawSpots emfDC, intSeriesIndex, True
    Next intSeriesIndex
    
    If ShowPosition Then WriteCoordinates emfDC, picD.ScaleWidth, picD.ScaleHeight
    lngResult = RestoreDC(emfDC, OldDC)
    emfHandle = CloseEnhMetaFile(emfDC)
    
    lngResult = OpenClipboard(picD.hwnd)
    lngResult = EmptyClipboard()
    lngResult = SetClipboardData(CF_ENHMETAFILE, emfHandle)
    lngResult = CloseClipboard
End Sub

Private Sub DestroyDrawingObjects()
    On Error Resume Next
    If hRegBrush <> 0 Then DeleteObject (hRegBrush)
    If hSelBrush <> 0 Then DeleteObject (hSelBrush)
    
    If hRegPen <> 0 Then DeleteObject (hRegPen)
    If hSelPen <> 0 Then DeleteObject (hSelPen)
    If hBlackPen <> 0 Then DeleteObject (hBlackPen)
    
    If hWhitePen <> 0 Then DeleteObject (hWhitePen)
    If hGridDashPen <> 0 Then DeleteObject (hGridDashPen)
End Sub

Private Sub DeviceLogicalCoordsConversion(ByVal ConversionType As Integer, ByVal NumOfPoints As Integer)
    Dim OldDC As Long, lngResult As Long
    
    OldDC = SaveDC(picD.hDC)
    EstablishCoordinateSystem False
    Select Case ConversionType
    Case 0
        ' Convert from device coordinates (pixels) to logical coordinates
        lngResult = DPtoLP(picD.hDC, paPoints(0), NumOfPoints)
    Case 1
        ' Convert from logical coordinates to device coordinates (pixels)
        lngResult = LPtoDP(picD.hDC, paPoints(0), NumOfPoints)
    End Select
    
    lngResult = RestoreDC(picD.hDC, OldDC)
End Sub

Private Sub DGCooSys(ByVal hDC As Long, ByVal dcWidth As Long, ByVal dcHeight As Long)
    '----------------------------------------------------------------------------
    ' Establishes the size of the coordinate system on the device context
    '----------------------------------------------------------------------------
    Dim lngResult As Long
    Dim ptPoint As POINTAPI
    Dim szSize As Size
    
    On Error Resume Next
    
    Const PlotTitleTextHeight = 13
    Const PositionTextHeight = 13
    Const TagIDTextHeight = 13
    Dim XTickMarkLabelsTextHeight As Long
    Dim YTickMarkLabelsTextWidth As Long
    
    XTickMarkLabelsTextHeight = 15
    If SwapAxes Then
        YTickMarkLabelsTextWidth = 7 * Len(Format((mCurrMinX + mCurrMaxX) / 2, VNumFmt))
    Else
        YTickMarkLabelsTextWidth = 6 * Len(Format((mCurrMinY + mCurrMaxY) / 2, VNumFmt))
    End If
    
    ' The following will fill the window, leaving a little space at the top for the TagID
    ' Note that the drawing origin is in the lower left of the picture, with the X axis
    '  extending toward the right and the Y axis extending upward; this is why
    '  VPY0 = dcHeight and VPYE is negative
    VPX0 = 0
    VPXE = dcWidth
    VPY0 = dcHeight
    VPYE = -dcHeight + TagIDTextHeight      ' Negative extent will result in drawing upward
    
    If Len(PlotTitleBottom) > 0 Then
        ' Shift the bottom up a little, and adjust the extent accordingly
        VPY0 = VPY0 - PlotTitleTextHeight
        VPYE = VPYE + PlotTitleTextHeight
    End If
    
    If ShowPosition Then
        ' Shrink the Y extent some more
        VPYE = VPYE + PositionTextHeight
    End If
    
    If ShowTickMarkLabels Then
        ' Move the x axis in a little and the bottom up some, adjusting extents as needed
        VPX0 = VPX0 + YTickMarkLabelsTextWidth
        VPXE = VPXE - YTickMarkLabelsTextWidth * 1.75
        
        VPY0 = VPY0 - XTickMarkLabelsTextHeight
        VPYE = VPYE + XTickMarkLabelsTextHeight
    End If
    
    lngResult = SetMapMode(hDC, MM_ANISOTROPIC)
    'logical window
    lngResult = SetWindowOrgEx(hDC, mLDfX0, mLDfY0, ptPoint)
    lngResult = SetWindowExtEx(hDC, mLDfXE, mLDfYE, szSize)
    
    'viewport
    lngResult = SetViewportOrgEx(hDC, VPX0, VPY0, ptPoint)
    lngResult = SetViewportExtEx(hDC, VPXE, VPYE, szSize)
    
    'set working window coordinates
    Select Case mViewWindow
    Case sFixedWindow
         mMinX = mFixedMinX:        mMaxX = mFixedMaxX
         mMinY = mFixedMinY:        mMaxY = mFixedMaxY
    Case sVariableWindow
         mMinX = mCurrMinX:        mMaxX = mCurrMaxX
         mMinY = mCurrMinY:        mMaxY = mCurrMaxY
    End Select
End Sub

Private Sub DGDrawCooSys(ByVal hDC As Long)
    '-----------------------------------------
    'draws the coordinate system on device context (i.e. draws the X and Y axes)
    '-----------------------------------------
    Dim ptPoint As POINTAPI
    Dim lngResult As Long
    
    On Error Resume Next
    'horizontal
    lngResult = MoveToEx(hDC, LDfX0, LDfY0, ptPoint)
    lngResult = LineTo(hDC, LDfX0 + mLDfXE, LDfY0)
    
    'vertical
    lngResult = MoveToEx(hDC, LDfX0, LDfY0, ptPoint)
    lngResult = LineTo(hDC, LDfX0, LDfY0 + mLDfYE)
End Sub

Private Sub DisplaySpotInfoOnSelected()
    Dim strFilePath As String
    Dim intFileNum As Integer
    Dim intSeriesIndex  As Integer
    Dim lngDataIndex As Long
    Dim lngMatchCount As Long
    
On Error GoTo DisplayInfoError

    intFileNum = FreeFile()
    strFilePath = GetTempFolder() & RawDataTmpFile
    
    Open strFilePath For Output As intFileNum
    Print #intFileNum, "Info on selected spots" & vbCrLf
    
    For intSeriesIndex = 0 To mSeriesCount - 1
        With mDataSeries(intSeriesIndex)
            For lngDataIndex = 0 To .DataCount - 1
                With .Data(lngDataIndex)
                    If .Selected Then
                        Print #intFileNum, "Series " & Trim(intSeriesIndex) & ", point " & Trim(lngDataIndex) & ", XLoc " & Format(.x, "0.0000") & ", YLoc " & Format(.y, "0.0000") & ", Intensity " & Round(.Intensity, 0) & ", " & .Description
                        lngMatchCount = lngMatchCount + 1
                    End If
                End With
            Next lngDataIndex
        End With
    Next intSeriesIndex
    
    If lngMatchCount = 0 Then Print #intFileNum, "No spots are selected."
    
    Close intFileNum
    DoEvents
    frmDataInfo.Tag = "Misc"
    DoEvents
    frmDataInfo.Show vbModal

    Exit Sub

DisplayInfoError:
    Close intFileNum
    MsgBox "Error writing output file (" & strFilePath & ") with the info to be displayed:" & Err.Description, vbExclamation + vbOKOnly, "Error"
    
End Sub

Private Sub DisplaySpotInfoOnNearbySpots()
    Dim strFilePath As String
    Dim intFileNum As Integer
    Dim intSpotMatchIndex As Integer
    Dim intSeriesIndex As Integer
    Dim lngDataIndex As Long
    
On Error GoTo DisplayInfoError

    intFileNum = FreeFile()
    strFilePath = GetTempFolder() & RawDataTmpFile
    
    Open strFilePath For Output As intFileNum
    Print #intFileNum, "Info on nearby spots" & vbCrLf
    
    If mSpotMatchCount = 0 Then
        Print #intFileNum, "No spots are near the cursor."
    Else
        
        For intSpotMatchIndex = 0 To mSpotMatchCount - 1
            intSeriesIndex = mSpotMatches(intSpotMatchIndex, 0)
            lngDataIndex = mSpotMatches(intSpotMatchIndex, 1)
            With mDataSeries(intSeriesIndex).Data(lngDataIndex)
                Print #intFileNum, "Series " & Trim(intSeriesIndex) & ", point " & Trim(lngDataIndex) & ", XLoc " & Format(.x, "0.0000") & ", YLoc " & Format(.y, "0.0000") & ", Intensity " & Round(.Intensity, 0) & ", " & .Description
            End With
        Next intSpotMatchIndex
    End If
    
    Close intFileNum
    DoEvents
    frmDataInfo.Tag = "Misc"
    DoEvents
    frmDataInfo.Show vbModal

    Exit Sub

DisplayInfoError:
    Close intFileNum
    MsgBox "Error writing output file (" & strFilePath & ") with the info to be displayed:" & Err.Description, vbExclamation + vbOKOnly, "Error"
    
End Sub

Private Sub Draw(intSeriesIndex As Integer)
    '-----------------------------------------------------
    'draws spots on picture device context
    '-----------------------------------------------------
    
    If mDataSeries(intSeriesIndex).DataCount > 0 Then
        ' Load the correct brush colors
        LoadBrushColors intSeriesIndex
        
        ' Draw the spots
        DrawSpots picD.hDC, intSeriesIndex
        
        ' Draw the selected spots
        ' Can use the DrawSpots function again, and simply draw over the spots
        '  that were already drawn, but use a different color
        DrawSpots picD.hDC, intSeriesIndex, True
    End If
End Sub

Private Sub DrawLabelsAndTickMarks(ByVal hDC As Long)
    Dim lfLogFont As LOGFONT
    Dim lngOldFont As Long
    Dim lngNewFont As Long
    Dim lngFont As Long
    Dim lngResult As Long
    Dim szHLabel As Size
    
    On Error Resume Next
    ' Get the font from the picture box control (Arial Narrow)
    lngFont = SelectObject(hDC, GetStockObject(SYSTEM_FONT))
    lngResult = GetObjectAPI(lngFont, Len(lfLogFont), lfLogFont)
    lngResult = SelectObject(hDC, lngFont)
    
    'create new logical font, using globally defined FontWidth and FontHeight
    lfLogFont.lfWidth = FontWidth
    lfLogFont.lfHeight = FontHeight
    lngNewFont = CreateFontIndirect(lfLogFont)
    
    'select newly created logical font to DC
    lngOldFont = SelectObject(hDC, lngNewFont)
    
    If ShowTickMarkLabels Then
        ' Draw coordinate axes labels (i.e. the name of the X axis and the name of the Y axis)
        lngResult = GetTextExtentPoint32(hDC, VLabel, Len(VLabel), szHLabel)
        
        lngResult = TextOut(hDC, mLDfXE - szHLabel.cx, -szHLabel.cy, HLabel, Len(HLabel))
        lngResult = TextOut(hDC, 50, mLDfYE + szHLabel.cy, VLabel, Len(VLabel))
    End If
    
    ' Draw tick marks on each of the axes
    WriteHMarkers hDC, mXAxisTickCount
    WriteVMarkers hDC, mYAxisTickCount
    
    'restore old font
    lngResult = SelectObject(hDC, lngOldFont)
    DeleteObject (lngNewFont)
End Sub

Private Sub DrawPlotTitle(ByVal hDC As Long)
    Dim lfLogFont As LOGFONT
    Dim lngOldFont As Long
    Dim lngNewFont As Long
    Dim lngFont As Long
    
    Dim szPlotTitle As Size
    Dim OldPen As Long
    Dim lngResult As Long
    
    Dim lngX1 As Long, lngX2 As Long
    Dim lngY1 As Long, lngY2 As Long
    
    On Error Resume Next
    ' Get the font from the picture box control (Arial Narrow)
    lngFont = SelectObject(hDC, GetStockObject(SYSTEM_FONT))
    lngResult = GetObjectAPI(lngFont, Len(lfLogFont), lfLogFont)
    lngResult = SelectObject(hDC, lngFont)
    
    'create new logical font, using globally defined FontWidth and FontHeight
    lfLogFont.lfWidth = FontWidth
    lfLogFont.lfHeight = FontHeight
    lngNewFont = CreateFontIndirect(lfLogFont)
    
    'select newly created logical font to DC
    lngOldFont = SelectObject(hDC, lngNewFont)
    
    If Len(PlotTitleBottom) > 0 Then
        
        lngResult = GetTextExtentPoint32(hDC, PlotTitleBottom, Len(PlotTitleBottom), szPlotTitle)
        
        ' Position PlotTitleBottom in the bottom left corner
        
        If ShowTickMarkLabels Then
            ' Found that -1600 x -1100 works well for a plot of size 200x200
            ' Or that -1300 x -800 works well for a plot of size 600x600
            'lngX1 = -1600
            'lngY1 = -1100
            lngX1 = -1400
            lngY1 = -800
        Else
            lngX1 = 0
            lngY1 = 0
        End If
        lngX2 = mLDfXE
        lngY2 = lngY1 - szPlotTitle.cy
        
        ' Need to clear whatever was written before; just draw regular rectangle
        OldPen = SelectObject(hDC, hWhitePen)             'invisible pen
        If ShowTickMarkLabels Then
            Rectangle hDC, lngX1, lngY1, lngX1 + szPlotTitle.cx, lngY2
        Else
            Rectangle hDC, lngX1, lngY1, lngX2, lngY2
        End If
        lngResult = SelectObject(hDC, OldPen)
        
        ' Display the plot title in the bottom left corner
        lngResult = TextOut(hDC, lngX1, lngY1, PlotTitleBottom, Len(PlotTitleBottom))
    End If
    
    'restore old font
    lngResult = SelectObject(hDC, lngOldFont)
    DeleteObject (lngNewFont)
End Sub


Public Sub DrawSpots(ByVal hDC As Long, intSeriesIndex As Integer, Optional blnDrawSelectedOnly As Boolean = False)
    '--------------------------------------------
    'draws spots to device context
    '--------------------------------------------
    Dim OldBrush As Long
    Dim OldPen As Long
    Dim lngResult As Long
    Dim lngDataPointIndex As Long
    Dim blnDrawAllSpots As Boolean
    Dim blnUseExtents As Boolean, blnUseIntensity As Boolean
    
    Dim ThisStar() As Long
    Dim ptStarPoints(7) As POINTAPI
    
    Dim intStarIndex As Long
    
    blnDrawAllSpots = Not blnDrawSelectedOnly
    
    On Error Resume Next
    
    With mDataSeries(intSeriesIndex)
        blnUseExtents = .UseExtents
        blnUseIntensity = .UseIntensity
        
        ' If .IonSpotShape = sEmptyRectangle Or .IonSpotShape = sTriangleWithExtents Then
        If .IonSpotShape = sEmptyRectangle Then
            ' Empty shapes are empty in the middle, but colored on the edge
            If blnDrawSelectedOnly Then
                OldBrush = SelectObject(hDC, GetStockObject(NULL_BRUSH))
                OldPen = SelectObject(hDC, hSelPen)
            Else
                OldBrush = SelectObject(hDC, GetStockObject(NULL_BRUSH))
                OldPen = SelectObject(hDC, hRegPen)
            End If
        Else
            ' Solid shapes use a black pen, but have a colored interior
            If blnDrawSelectedOnly Then
                OldBrush = SelectObject(hDC, hSelBrush)
            Else
                OldBrush = SelectObject(hDC, hRegBrush)
            End If
            OldPen = SelectObject(hDC, hBlackPen)
            
            If Not blnDrawSelectedOnly Then
                ' When drawing all spots, Use Mask_Pen so overlapped spots are shaded
                lngResult = SetROP2(hDC, R2_MASKPEN)
            End If
        End If
        
        Select Case .IonSpotShape
        Case sCircle
            For lngDataPointIndex = 0 To .DataCount - 1
                If blnDrawAllSpots Or .Data(lngDataPointIndex).Selected Then
                    FillExtents .mgDataScaled(lngDataPointIndex), blnUseExtents, blnUseIntensity
                    With .mgDataScaled(lngDataPointIndex)
                        DrawEllipse hDC, .CornerXNeg, .CornerYNeg, .CornerXPos, .CornerYPos
                    End With
                End If
            Next lngDataPointIndex
            
        Case sRectangle, sEmptyRectangle
            For lngDataPointIndex = 0 To .DataCount - 1
                If blnDrawAllSpots Or .Data(lngDataPointIndex).Selected Then
                    FillExtents .mgDataScaled(lngDataPointIndex), blnUseExtents, blnUseIntensity
                    With .mgDataScaled(lngDataPointIndex)
                        DrawRectangle hDC, .CornerXNeg, .CornerYNeg, .CornerXPos, .CornerYPos
                    End With
                End If
            Next lngDataPointIndex
            
        Case sRoundRectangle
            For lngDataPointIndex = 0 To .DataCount - 1
                If blnDrawAllSpots Or .Data(lngDataPointIndex).Selected Then
                    FillExtents .mgDataScaled(lngDataPointIndex), blnUseExtents, blnUseIntensity
                    With .mgDataScaled(lngDataPointIndex)
                        DrawRoundRectangle hDC, .CornerXNeg, .CornerYNeg, .CornerXPos, .CornerYPos
                    End With
                End If
            Next lngDataPointIndex
            
        Case sTriangleWithExtents
            For lngDataPointIndex = 0 To .DataCount - 1
                If blnDrawAllSpots Or .Data(lngDataPointIndex).Selected Then
                    FillExtents .mgDataScaled(lngDataPointIndex), blnUseExtents, blnUseIntensity
                    DrawTriangle hDC, intSeriesIndex, lngDataPointIndex
                End If
            Next lngDataPointIndex
            
        Case sStar
            ' Note: The stars are not programmed to handle extents or intensity
            For lngDataPointIndex = 0 To .DataCount - 1
                If blnDrawAllSpots Or .Data(lngDataPointIndex).Selected Then
                    GetStarPoints .mgDataScaled(lngDataPointIndex).x, .mgDataScaled(lngDataPointIndex).y, ThisStar()
                    For intStarIndex = 0 To 7
                        ptStarPoints(intStarIndex).x = ThisStar(intStarIndex, 0)
                        ptStarPoints(intStarIndex).y = ThisStar(intStarIndex, 1)
                    Next intStarIndex
                    Polygon hDC, ptStarPoints(0), 8
                End If
            Next lngDataPointIndex
        End Select
       
    End With
    
    Call SetROP2(hDC, R2_COPYPEN)
    Call SelectObject(hDC, OldPen)
    Call SelectObject(hDC, OldBrush)
End Sub

Private Sub DrawEllipse(hDC As Long, CornerXNeg As Long, CornerYNeg As Long, CornerXPos As Long, CornerYPos As Long)
    Ellipse hDC, CornerXNeg, CornerYNeg, CornerXPos, CornerYPos
End Sub

Private Sub DrawRectangle(hDC As Long, CornerXNeg As Long, CornerYNeg As Long, CornerXPos As Long, CornerYPos As Long)
    Rectangle hDC, CornerXNeg, CornerYNeg, CornerXPos, CornerYPos
End Sub

Private Sub DrawRoundRectangle(hDC As Long, CornerXNeg As Long, CornerYNeg As Long, CornerXPos As Long, CornerYPos As Long)
    ' The last two parameters define the degree of rounding for the rectangle
    RoundRect hDC, CornerXNeg, CornerYNeg, CornerXPos, CornerYPos, CornerXPos / 25, CornerXPos / 25
End Sub

Private Sub DrawTriangle(hDC As Long, intSeriesIndex As Integer, lngDataPointIndex As Long)
    Dim ptTrianglePoints(0 To 2) As POINTAPI
    
    GetTrianglePointsLaSpots intSeriesIndex, lngDataPointIndex, ptTrianglePoints()
    
    Polygon hDC, ptTrianglePoints(0), 3

End Sub

Private Sub FillExtents(ByRef mgDataScaled As udtDataPointScaledType, blnUseExtents As Boolean, blnUseIntensity As Boolean)
    
    Dim lngExtentInXNeg As Long, lngExtentInXPos As Long
    Dim lngExtentInYNeg As Long, lngExtentInYPos As Long
    
    With mgDataScaled
        If blnUseIntensity Then
            If SwapAxes Then
                lngExtentInXNeg = .IntensityHalf
                lngExtentInXPos = .IntensityHalf
                
                If blnUseExtents Then
                    ' Use extents in Y, but record Intensity in X Extents (done above)
                    lngExtentInYNeg = .ExtentInYNeg
                    lngExtentInYPos = .ExtentInYPos
                Else
                    ' Use intensity in both X and Y to give a symmetric shape that grows with intensity
                    lngExtentInYNeg = lngExtentInXNeg
                    lngExtentInYPos = lngExtentInXPos
                End If
            Else
                lngExtentInYNeg = .IntensityHalf
                lngExtentInYPos = .IntensityHalf
                
                If blnUseExtents Then
                    ' Use extents in X, but record Intensity in Y Extents (done above)
                    lngExtentInXNeg = .ExtentInXNeg
                    lngExtentInXPos = .ExtentInXPos
                Else
                    ' Use intensity in both X and Y to give a symmetric shape that grows with intensity
                    lngExtentInXNeg = lngExtentInYNeg
                    lngExtentInXPos = lngExtentInYPos
                End If
            End If
        Else
            If blnUseExtents Then
                lngExtentInXNeg = .ExtentInXNeg
                lngExtentInXPos = .ExtentInXPos
                
                lngExtentInYNeg = .ExtentInYNeg
                lngExtentInYPos = .ExtentInYPos
            Else
                ' No extents and no intensity
                ' All spots are the same size
                lngExtentInXNeg = mHalfSize
                lngExtentInXPos = mHalfSize
                
                lngExtentInYNeg = mHalfSize
                lngExtentInYPos = mHalfSize
            End If
        End If
        
        .CornerXNeg = .x - lngExtentInXNeg
        .CornerXPos = .x + lngExtentInXPos
        .CornerYNeg = .y - lngExtentInYNeg
        .CornerYPos = .y + lngExtentInYPos
        
        ' Make sure the extents are at least as big as mMinSpotSize
        If .CornerXPos - .CornerXNeg < mMinSpotSize Then
            .CornerXNeg = .CornerXNeg - mHalfMinSpotSize
            .CornerXPos = .CornerXPos + mHalfMinSpotSize
        End If
        
        If .CornerYPos - .CornerYNeg < mMinSpotSize Then
            .CornerYNeg = .CornerYNeg - mHalfMinSpotSize
            .CornerYPos = .CornerYPos + mHalfMinSpotSize
        End If
        
    End With

End Sub

Private Sub EstablishCoordinateSystem(Optional blnSaveReleaseDeviceContext As Boolean = True)

    Dim OldDC As Long
    Dim lngResult As Long
    
    ' Establish coordinate system
    If blnSaveReleaseDeviceContext Then OldDC = SaveDC(picD.hDC)
    
    DGCooSys picD.hDC, picD.ScaleWidth, picD.ScaleHeight
    
    If blnSaveReleaseDeviceContext Then lngResult = RestoreDC(picD.hDC, OldDC)

End Sub

Public Function GetSelection(SelIndices() As Long, SelDescription() As String, Optional intSeriesIndex As Integer = 0) As Long
    '----------------------------------------------------
    'fills array Sel with ids of currently selected spots
    'returns number of spots; -1 on any error
    '----------------------------------------------------
    Dim SelCnt As Long
    Dim i As Long
    
    On Error GoTo err_GetSelection
        
    If intSeriesIndex < 0 Or intSeriesIndex >= mSeriesCount Then
        intSeriesIndex = 0
    End If
    
    With mDataSeries(intSeriesIndex)
        If .DataCount > 0 Then
            ReDim SelIndices(.DataCount - 1)
            ReDim SelDescription(.DataCount - 1)
            For i = 0 To .DataCount - 1
                If .Data(i).Selected Then
                   SelCnt = SelCnt + 1
                   SelIndices(SelCnt - 1) = i
                   SelDescription(SelCnt - 1) = .Data(i).Description
                End If
            Next i
        End If
        
        If SelCnt > 0 Then
            If SelCnt < .DataCount - 1 Then
                ReDim Preserve SelIndices(SelCnt - 1)
                ReDim Preserve SelDescription(SelCnt - 1)
            End If
        Else
           Erase SelIndices
           Erase SelDescription
        End If
    End With
    
    GetSelection = SelCnt
    Exit Function
    
err_GetSelection:
    MsgBox "Error in LaSpots->GetSelection: " & Err.Description
    
    Erase SelIndices
    Erase SelDescription
    GetSelection = -1
End Function

Private Sub GetStarPoints(ByVal x As Long, ByVal y As Long, StarPoints() As Long)
    Dim i As Long
    
    On Error Resume Next
    ReDim StarPoints(7, 1)
    For i = 0 To 7
        Select Case i
        Case 0, 4
            StarPoints(i, 0) = x
        Case 1, 3
            StarPoints(i, 0) = x - mHalfSize \ 4
        Case 5, 7
            StarPoints(i, 0) = x + mHalfSize \ 4
        Case 2
            StarPoints(i, 0) = x - mHalfSize
        Case 6
            StarPoints(i, 0) = x + mHalfSize
        End Select
        Select Case i
        Case 2, 6
            StarPoints(i, 1) = y
        Case 1, 7
            StarPoints(i, 1) = y - CLng(mHalfSize / 4)
        Case 3, 5
            StarPoints(i, 1) = y + CLng(mHalfSize / 4)
        Case 0
            StarPoints(i, 1) = y - CLng(mHalfSize)
        Case 4
            StarPoints(i, 1) = y + CLng(mHalfSize)
        End Select
    Next i
End Sub

Private Sub GetTrianglePointsLaSpots(intSeriesIndex As Integer, lngDataPointIndex As Long, ptTrianglePoints() As POINTAPI, Optional blnCenterTriangleOnSpotCenter As Boolean = False)
    
    With mDataSeries(intSeriesIndex)
        With .mgDataScaled(lngDataPointIndex)
            ' Point 0 and Point 1 define the base of the triangle
            ' Point 2 is the apex
            
            If SwapAxes Then
                ' Data has already been calculated swapped, so we
                ' need to draw the triangle differently
                '   X = Mass
                '   Y = NET
                
                If blnCenterTriangleOnSpotCenter Then
                    ' If we wish to center the triangle around X,Y, then use
                    
                    ptTrianglePoints(0).x = .CornerXNeg
                    ptTrianglePoints(0).y = .CornerYNeg
                    
                    ptTrianglePoints(1).x = .CornerXNeg
                    ptTrianglePoints(1).y = .CornerYPos
                    
                    ptTrianglePoints(2).x = .CornerXPos      ' .CornerXPos is based on .IntensityHalf()
                    ptTrianglePoints(2).y = .y
                Else
                    ' Otherwise, place the base of the triangle at .X
                    ' Since .CornerXPos is based on .IntensityHalf(), must subtract
                    '  (.CornerXPos - .CornerXNeg) and add to .X to get the correct X location for the apex of the triangle
                    ' This allows for asymmetric triangles
                    ptTrianglePoints(0).x = .x
                    ptTrianglePoints(0).y = .CornerYNeg
                    
                    ptTrianglePoints(1).x = .x
                    ptTrianglePoints(1).y = .CornerYPos
                    
                    ptTrianglePoints(2).x = .x + (.CornerXPos - .CornerXNeg)
                    ptTrianglePoints(2).y = .y
                End If
            Else
                ' Data was calculated normal, and we'll swap the points just prior to drawing the triangle
                '   X = NET
                '   Y = Mass

                If blnCenterTriangleOnSpotCenter Then
                    ' If we wish to center the triangle around X,Y, then use
                    ptTrianglePoints(0).x = .CornerXNeg
                    ptTrianglePoints(0).y = .CornerYNeg
                    
                    ptTrianglePoints(1).x = .CornerXPos
                    ptTrianglePoints(1).y = .CornerYNeg
                    
                    ptTrianglePoints(2).x = .x
                    ptTrianglePoints(2).y = .CornerYPos     ' .CornerYPos is based on .IntensityHalf()
                Else
                    ' Otherwise, place the base of the triangle at .Y
                    ' Since .CornerYPos is based on .IntensityHalf(), must subtract
                    '  (.CornerYPos - .CornerYNeg) and add to .Y to get the correct Y location for the apex of the triangle
                    ' This allows for asymmetric triangles
                    ptTrianglePoints(0).x = .CornerXNeg
                    ptTrianglePoints(0).y = .y
                    
                    ptTrianglePoints(1).x = .CornerXPos
                    ptTrianglePoints(1).y = .y
                    
                    ptTrianglePoints(2).x = .x
                    ptTrianglePoints(2).y = .y + (.CornerYPos - .CornerYNeg)
                End If
            End If
        End With
    End With
    
End Sub

Private Function GetTotalDataCount() As Long
    Dim intSeriesIndex As Integer
    Dim lngDataCount As Long
    
    lngDataCount = 0
    For intSeriesIndex = 0 To mSeriesCount - 1
        lngDataCount = lngDataCount + mDataSeries(intSeriesIndex).DataCount
    Next intSeriesIndex
    
    GetTotalDataCount = lngDataCount
End Function

Private Sub HandleSpotInfoClick(intMenuIndex As Integer)
    ' Examine the menu caption
    ' It it begins with "Ion", then select the ion
    ' If it begins with "UMC", then select the ions belonging to the UMC (if one of the series contains ions), otherwise, select the UMC
    ' If it begins with "Mass Tag" then, then call the ?? Procedure in OrfViewerRoutines.Bas to instruct the
    '  ORF viewer containing this picture to jump to the ORF containing the clicked MT tag
    
    Dim strCaption As String, strUMCIons As String
    Dim lngSpotSeriesIndex As Long, lngSpotDataIndex As Long
    Dim lngColonLoc As Long, lngEqualLoc As Long
    Dim lngClassMemberIndicatorLoc As Long
    Dim strID As String
    Dim varSplitArray As Variant
    Dim intIonIndex As Integer
    Dim strIonDescription As String
    Dim lngIonDescriptionLength As Long

    Dim intSeriesIndex As Integer, lngDataIndex As Long
    
On Error GoTo HandleSpotInfoClickErrorHandler

    strCaption = mnuSpotInfo(intMenuIndex).Caption
    
    ' Deselect all spots
    SelectDeselectAll False
    
    lngSpotSeriesIndex = mSpotMatches(intMenuIndex, 0)
    lngSpotDataIndex = mSpotMatches(intMenuIndex, 1)
    
'' Code that was used by the ORFViewer; No longer supported (March 2006)
''    If Left(strCaption, Len(ORF_VIEWER_MASS_TAG_STRING)) = ORF_VIEWER_MASS_TAG_STRING Then
''        ' Determine the MT tag ID for the selected MT tag
''        lngColonLoc = InStr(strCaption, ORF_VIEWER_ID_DELIMETER)
''        If lngColonLoc > 0 Then
''            strID = Trim(Mid(strCaption, Len(ORF_VIEWER_MASS_TAG_STRING), lngColonLoc - Len(ORF_VIEWER_MASS_TAG_STRING)))
''            If IsNumeric(strID) Then
''                ORFViewerLoader.FindORFContainingMassTagInCallingForm CallingFormID, val(strID)
''            End If
''        End If
''    ElseIf Left(strCaption, Len(ORF_VIEWER_UMC_STRING)) = ORF_VIEWER_UMC_STRING Then
''        lngEqualLoc = InStr(strCaption, ORF_VIEWER_UMC_ION_LIST_START_STRING)
''        If lngEqualLoc > 0 Then
''            strUMCIons = Mid(strCaption, lngEqualLoc + Len(lngEqualLoc) - 1)
''
''            varSplitArray = Split(strUMCIons, ORF_VIEWER_UMC_ION_LIST_DELIMETER)
''
''            For intIonIndex = 0 To UBound(varSplitArray)
''                lngClassMemberIndicatorLoc = InStr(varSplitArray(intIonIndex), ORF_VIEWER_UMC_REPRESENTATIVE_MEMBER_INDICATOR)
''                If lngClassMemberIndicatorLoc > 0 Then
''                    varSplitArray(intIonIndex) = Left(varSplitArray(intIonIndex), lngClassMemberIndicatorLoc - 1)
''                End If
''                If IsNumeric(varSplitArray(intIonIndex)) Then
''                    strIonDescription = ORF_VIEWER_ION_STRING & Trim(varSplitArray(intIonIndex))
''                    lngIonDescriptionLength = Len(strIonDescription)
''
''                    For intSeriesIndex = 0 To mSeriesCount - 1
''                        With mDataSeries(intSeriesIndex)
''                            strIonDescription = ORF_VIEWER_ION_STRING & Trim(varSplitArray(intIonIndex))
''                            lngIonDescriptionLength = Len(strIonDescription)
''                            For lngDataIndex = 0 To .DataCount - 1
''                                If Left(.Data(lngDataIndex).Description, lngIonDescriptionLength) = strIonDescription Then
''                                    .Data(lngDataIndex).Selected = True
''                                End If
''                            Next lngDataIndex
''                        End With
''                    Next intSeriesIndex
''                End If
''            Next intIonIndex
''            RefreshPlot
''        End If
''    Else
        ' Includes single ions (which start with ORF_VIEWER_ION_STRING)
        ' Simply select the spot
        mDataSeries(lngSpotSeriesIndex).Data(lngSpotDataIndex).Selected = True
        RefreshPlot
''    End If
    
    Exit Sub

HandleSpotInfoClickErrorHandler:
    MsgBox "An error occured in LaSpots|HandleSpotInfoClick: " & Err.Description

End Sub
Private Sub LoadBrushColors(intSeriesIndex As Integer)
    On Error Resume Next
    
    ' May or may not need to destroy before loading new
    DestroyDrawingObjects
    
    With mDataSeries(intSeriesIndex)
        hRegBrush = CreateSolidBrush(.RegColor)
        hSelBrush = CreateSolidBrush(.SelColor)
        
        hRegPen = CreatePen(PS_SOLID, mBorderSize, .RegColor)
        hSelPen = CreatePen(PS_SOLID, mBorderSize, .SelColor)
        
        hBlackPen = CreatePen(PS_SOLID, mBorderSize, picD.ForeColor)
        hWhitePen = CreatePen(PS_SOLID, 1, picD.BackColor)
        hGridDashPen = CreatePen(PS_DOT, 1, picD.ForeColor)
    End With
End Sub

Private Sub SetLogicalCoordinates()
    Const LogicalCoordScaleWithPicSize = True
    
    If LogicalCoordScaleWithPicSize Then
        mLDfXE = picD.width * 3
        mLDfYE = picD.Height * 3
        
        If mLDfXE < 100 Then mLDfXE = 100
        If mLDfYE < 100 Then mLDfYE = 100
    Else
        mLDfXE = 10000
        mLDfYE = 10000
    End If
    
    mMaxSpotSize = 500 * LogicalPointsPerPixel(True)
    mHalfSize = mMaxSpotSize / 2
    mQuarterSize = mMaxSpotSize / 4
    
    mMinSpotSize = 1 * LogicalPointsPerPixel(True)
    mHalfMinSpotSize = mMinSpotSize / 2

    mBorderSize = mMinSpotSize

End Sub
Public Sub SetSpotShapesAndColors(Optional eSpotShape As sSpotsShape = -1, Optional lngRegularSpotColor As Long = -1, Optional lngSelectedSpotColor As Long = -1, Optional intSeriesIndex As Integer = -1, Optional blnMakeDefault As Boolean = False, Optional lngBackgroundColor As Long = -1)
    Dim intIndex As Integer
    Dim intIndexStart As Integer, intIndexEnd As Integer
    
    If mSeriesCount = 0 Then Exit Sub
    
    If intSeriesIndex < 0 Then
        intIndexStart = 0
        intIndexEnd = mSeriesCount - 1
    ElseIf intSeriesIndex >= mSeriesCount Then
        intIndexStart = mSeriesCount - 1
        intIndexEnd = mSeriesCount - 1
    Else
        intIndexStart = intSeriesIndex
        intIndexEnd = intSeriesIndex
    End If
    
    For intIndex = intIndexStart To intIndexEnd
        With mDataSeries(intIndex)
            If eSpotShape >= 0 Then
                .IonSpotShape = eSpotShape
            Else
                .IonSpotShape = mDefaultSpotShape
            End If
            
            If lngRegularSpotColor >= 0 Then
                .RegColor = lngRegularSpotColor
            Else
                .RegColor = mDefaultRegColor
            End If
            
            If lngSelectedSpotColor >= 0 Then
                .SelColor = lngSelectedSpotColor
            Else
                .SelColor = mDefaultSelColor
            End If
            
        End With
    Next intIndex
    
    If blnMakeDefault Then
        If eSpotShape >= 0 Then mDefaultSpotShape = eSpotShape
        If lngRegularSpotColor >= 0 Then mDefaultRegColor = lngRegularSpotColor
        If lngSelectedSpotColor >= 0 Then mDefaultSelColor = lngSelectedSpotColor
    End If
    
    If lngBackgroundColor >= 0 Then
        picD.BackColor = lngBackgroundColor
    End If
End Sub

Public Sub SetBorderSize(ByVal lngNewBorderSize As Long, Optional ByVal blnConvertFromPixelsToLogicalCoordinates As Boolean = True)
    
    If blnConvertFromPixelsToLogicalCoordinates Then
        lngNewBorderSize = lngNewBorderSize * LogicalPointsPerPixel(False)
    End If
    
    mBorderSize = lngNewBorderSize
End Sub

Public Sub SetMaxSpotSize(ByVal lngNewSpotSize As Long, Optional ByVal blnConvertFromPixelsToLogicalCoordinates As Boolean = True)
    
    If blnConvertFromPixelsToLogicalCoordinates Then
        lngNewSpotSize = lngNewSpotSize * LogicalPointsPerPixel(False)
    End If
    
    mMaxSpotSize = lngNewSpotSize
    mHalfSize = mMaxSpotSize / 2
    mQuarterSize = mMaxSpotSize / 4
End Sub

Public Sub SetMinSpotSize(ByVal lngNewSpotSize As Long, Optional ByVal blnConvertFromPixelsToLogicalCoordinates As Boolean = True)
    
    If blnConvertFromPixelsToLogicalCoordinates Then
        lngNewSpotSize = lngNewSpotSize * LogicalPointsPerPixel(False)
    End If
    
    mMinSpotSize = lngNewSpotSize
    mHalfMinSpotSize = mMinSpotSize / 2
    
End Sub

Public Sub SetFixedWindow(ByVal XMin As Double, ByVal XMax As Double, _
                          ByVal YMin As Double, ByVal YMax As Double)
    '----------------------------------------------------------------------
    'sets fixed window coordinates to be used for coordinate system
    '----------------------------------------------------------------------
    mFixedMinX = XMin:  mFixedMaxX = XMax
    mFixedMinY = YMin:  mFixedMaxY = YMax
End Sub

Private Function LogicalPointsPerPixel(blnReturnXDimension As Boolean) As Double
    
    ' Fill paPoints(0) with a coordinate of 1,1, then call DeviceLogicalCoordsConversion
    '  and read the new coordinates
    
    paPoints(0).x = 1
    paPoints(0).y = 1
    
    If blnReturnXDimension Then
        If picD.width > 0 Then LogicalPointsPerPixel = (mLDfXE - LDfX0) / picD.width
    Else
        If picD.Height > 0 Then LogicalPointsPerPixel = (mLDfYE - LDfY0) / picD.Height
    End If
'
'    DeviceLogicalCoordsConversion 1, 1
'
'    If blnReturnXDimension Then
'        LogicalPointsPerPixel = paPoints(0).x
'    Else
'        LogicalPointsPerPixel = paPoints(0).y
'    End If
    
    If LogicalPointsPerPixel <= 0 Then LogicalPointsPerPixel = 1
    
End Function

Private Sub PositionControls()
    picD.Left = 0
    picD.Top = 0
    picD.width = UserControl.ScaleWidth
    picD.Height = UserControl.ScaleHeight
    
    ComputeOptimalAxisTickCounts
End Sub

Private Sub SelectDeselectAll(blnSelectAll As Boolean, Optional blnRefreshPlot As Boolean = True)
    Dim intSeriesIndex As Integer
    Dim lngDataPointIndex As Long
    
    For intSeriesIndex = 0 To mSeriesCount - 1
        With mDataSeries(intSeriesIndex)
            For lngDataPointIndex = 0 To .DataCount - 1
                .Data(lngDataPointIndex).Selected = blnSelectAll
            Next lngDataPointIndex
        End With
    Next intSeriesIndex
    
    If blnRefreshPlot Then RefreshPlot
End Sub

Public Sub RefreshPlot()
    '-------------------------------------------
    'rebuild coordinate system and redraws spots
    '-------------------------------------------
    
    Dim OldDC As Long
    Dim lngResult As Long
    Dim intSeriesIndex As Integer
    
    picD.Cls
    
    OldDC = SaveDC(picD.hDC)
    EstablishCoordinateSystem False
    
    DGDrawCooSys picD.hDC
    
    DrawLabelsAndTickMarks picD.hDC

    ' Need to define a clipping region before drawing the spots
    ' This isn't working
    
'    Dim ptOrg As POINTAPI
'    Dim szExt As Size
    Dim lngReturn As Long
    Dim hCR As Long
'    Dim lpRect As Rect
    
'    lngReturn = GetViewportOrgEx(hDC, ptOrg)
'    lngReturn = GetViewportExtEx(hDC, szExt)
    hCR = CreateRectRgn(VPX0, VPY0, VPXE, VPYE)
    'hCR = CreateRectRgn(LDfX0, LDfY0, LDfXE, LDfYE)
    
'    lngReturn = GetRgnBox(hCR, lpRect)
    
    lngReturn = SelectClipRgn(hDC, hCR)

    For intSeriesIndex = 0 To mSeriesCount - 1
        Call CalculateSpots(intSeriesIndex)
        Call Draw(intSeriesIndex)
    Next intSeriesIndex
    
    ' Since I can't get the clipping region to work,
    ' I'll add the plot title after drawing the spots
    DrawPlotTitle picD.hDC
    
    lngResult = RestoreDC(picD.hDC, OldDC)
    lngReturn = DeleteObject(hCR)
End Sub

Public Sub SetAxisLabelPrecision(intXAxisDigitsAfterDecimal As Integer, intYAxisDigitsAfterDecimal As Integer)
    ' Changes HNumFmt and VNumFmt to be of the form "0" or "0.00", etc.
    ' Note that HNumFmt and VNumFmt are Public, so they can be set to fancy formatting strings
    
    If intXAxisDigitsAfterDecimal < 0 Then intXAxisDigitsAfterDecimal = 0
    If intYAxisDigitsAfterDecimal < 0 Then intYAxisDigitsAfterDecimal = 0
    
    If intXAxisDigitsAfterDecimal > 15 Then intXAxisDigitsAfterDecimal = 15
    If intYAxisDigitsAfterDecimal > 15 Then intYAxisDigitsAfterDecimal = 15
    
    HNumFmt = "0"
    If intXAxisDigitsAfterDecimal > 0 Then
        HNumFmt = HNumFmt & "." & String(intXAxisDigitsAfterDecimal, "0")
    End If

    VNumFmt = "0"
    If intXAxisDigitsAfterDecimal > 0 Then
        VNumFmt = VNumFmt & "." & String(intYAxisDigitsAfterDecimal, "0")
    End If

End Sub


Public Sub SetAxisTickMarkCount(Optional ByRef lngXAxisTickCount = -1, Optional ByRef lngYAxisTickCount = -1)
    '------------------------------------------------------------------------
    ' Sets the number of tick marks to display on each axis
    ' If a value of -1 is supplied, computes the number of tick marks based on the current graph size
    ' Variables are passed ByRef so that user can see the value that was actually set
    '------------------------------------------------------------------------
    
    If lngXAxisTickCount <= 1 Then
        mAutoComputeXAxisTickCount = True
    Else
        mXAxisTickCount = lngXAxisTickCount
    End If
    
    If lngYAxisTickCount <= 1 Then
        mAutoComputeYAxisTickCount = True
    Else
        mYAxisTickCount = lngYAxisTickCount
    End If
    
    ComputeOptimalAxisTickCounts
        
End Sub

Private Sub SwapLong(ByRef FirstValue As Long, ByRef SecondValue As Long)
    Dim lngTemp As Long
    lngTemp = FirstValue
    FirstValue = SecondValue
    SecondValue = lngTemp
End Sub

Public Function ToggleSpotSelection(ByVal lngDataPointIndex As Long, Optional intSeriesIndex As Integer = 0) As Boolean
    '------------------------------------------------------------------------
    'toggles spot selection and returns current state of selection for a spot
    '------------------------------------------------------------------------
    On Error Resume Next
    
    If intSeriesIndex < 0 Or intSeriesIndex >= mSeriesCount Then
        intSeriesIndex = 0
    End If
    
    With mDataSeries(intSeriesIndex)
        If lngDataPointIndex >= 0 And lngDataPointIndex < .DataCount Then
            .Data(lngDataPointIndex).Selected = Not .Data(lngDataPointIndex).Selected
        End If
    End With
    
    RefreshPlot
End Function

Private Sub TrackHotSpot(lx As Long, ly As Long)
    Dim intSeriesIndex As Integer
    Dim lngDataIndex As Long
    Dim blnMouseIsOverSpot As Boolean
    Dim blnTrackMouseOverEdge As Boolean
    
    On Error Resume Next
    If SwapAxes Then
        ' Compute current X and Y using the following
        mCurrX = Format((lx / mScaleX) + mMinY, "0.0000")        'current coordinates
        mCurrY = Format((ly / mScaleY) + mMinX, "0.0000")
    Else
        ' Compute current X and Y using the following
        mCurrX = Format((lx / mScaleX) + mMinX, "0.0000")            'current coordinates
        mCurrY = Format((ly / mScaleY) + mMinY, "0.0000")
    End If
    
    mSpotMatchCount = 0
    mHotSpotSeriesIndex = -1
    mHotSpotDataIndex = -1
    For intSeriesIndex = 0 To mSeriesCount - 1
        With mDataSeries(intSeriesIndex)
            If .UseExtents And Not .UseIntensity Then
                blnTrackMouseOverEdge = True
            Else
                blnTrackMouseOverEdge = False
            End If
            
            For lngDataIndex = 0 To .DataCount - 1
                blnMouseIsOverSpot = False
                
                With .mgDataScaled(lngDataIndex)
                    
                    If blnTrackMouseOverEdge Then
                        ' The following will tell you if the mouse is over a shape, but not if it's over the "Middle" of the shape
                                                
                        ' Determine if the mouse is over one of the edges of the spot
                        ' Will use the .CornerX and .CornerY information to determine this
                        ' Treating all spots as rectangles, even if they're not
                        
                        If (WithinToleranceLng(lx, .CornerXNeg, mQuarterSize) Or WithinToleranceLng(lx, .CornerXPos, mQuarterSize)) And ly >= .CornerYNeg And ly <= .CornerYPos Then
                            ' Mouse is over the left or right edge
                            blnMouseIsOverSpot = True
                        ElseIf (WithinToleranceLng(ly, .CornerYNeg, mQuarterSize) Or WithinToleranceLng(ly, .CornerYPos, mQuarterSize)) And lx >= .CornerXNeg And lx <= .CornerXPos Then
                            ' Mouse is over top or bottom edge
                            blnMouseIsOverSpot = True
                        End If
                    Else
                        ' See if mouse is over the middle of the shape
                        If Abs(lx - .x) <= mQuarterSize Then
                            If Abs(ly - .y) <= mQuarterSize Then
                                blnMouseIsOverSpot = True
                            End If
                        End If
                    End If
                End With
                
                If blnMouseIsOverSpot Then
                    If mSpotMatchCount < MAX_SPOT_MATCHES_TO_TRACK Then
                        mSpotMatches(mSpotMatchCount, 0) = intSeriesIndex
                        mSpotMatches(mSpotMatchCount, 1) = lngDataIndex
                        mSpotMatchCount = mSpotMatchCount + 1
                    End If
                    
                    mHotSpotSeriesIndex = intSeriesIndex
                    mHotSpotDataIndex = lngDataIndex
                    If SwapAxes Then
                        ' X and Y are reversed
                        mCurrX = Format(.Data(lngDataIndex).y, "0.0000")     'change current coordinates to
                        mCurrY = Format(.Data(lngDataIndex).x, "0.0000")     'exact data location if hot spot
                    Else
                        mCurrX = Format(.Data(lngDataIndex).x, "0.0000")     'change current coordinates to
                        mCurrY = Format(.Data(lngDataIndex).y, "0.0000")     'exact data location if hot spot
                    End If
                End If
            Next lngDataIndex
        End With
        
    Next intSeriesIndex
    
End Sub

Private Sub UpdateDynamicSpotsNearCursorMenu()
    ' Update the menu items to list the descriptions of the spots near the cursor
    Dim intSpotMatchIndex As Integer, lngMenuIndex As Integer
    Dim intSpotCountToDisplay As Integer
    
    intSpotCountToDisplay = mSpotMatchCount
    If intSpotCountToDisplay > 15 Then intSpotCountToDisplay = 15
    
    Do While mSpotInfoMenuCountLoaded < intSpotCountToDisplay
        Load mnuSpotInfo(mSpotInfoMenuCountLoaded)
        mSpotInfoMenuCountLoaded = mSpotInfoMenuCountLoaded + 1
    Loop
    
    For intSpotMatchIndex = 0 To intSpotCountToDisplay - 1
        With mDataSeries(mSpotMatches(intSpotMatchIndex, 0)).Data(mSpotMatches(intSpotMatchIndex, 1))
            If Len(.Description) < 50 Then
                mnuSpotInfo(intSpotMatchIndex).Caption = .Description
            Else
                mnuSpotInfo(intSpotMatchIndex).Caption = Left(.Description, 100) & " ..."
            End If
            mnuSpotInfo(intSpotMatchIndex).Visible = True
        End With
    Next intSpotMatchIndex
    
    ' Hide any remaining menus
    ' VB requires that at least one submenu be visible at a given time
    ' Therefore, the following sometimes produces an error
    ' Thus, we'll enable On Error Resume Next handling
    On Error Resume Next
    For lngMenuIndex = intSpotCountToDisplay To mSpotInfoMenuCountLoaded - 1
        mnuSpotInfo(lngMenuIndex).Caption = ""
        mnuSpotInfo(lngMenuIndex).Visible = False
    Next lngMenuIndex
End Sub

Private Function WithinToleranceLng(ThisNumber As Long, CompareNumber As Long, ThisTolerance As Long) As Boolean
    If ThisNumber <= CompareNumber + ThisTolerance And ThisNumber >= CompareNumber - ThisTolerance Then
        WithinToleranceLng = True
    Else
        WithinToleranceLng = False
    End If
End Function

Private Sub WriteCoordinates(hDC As Long, w As Long, h As Long)
    Dim OldDC As Long
    Dim lfLogFont As LOGFONT
    Dim lngOldFont As Long
    Dim lngNewFont As Long
    Dim strLabelID As String, strLabelPos As String
    Dim lngFont As Long
    Dim szLabelID As Size, szLabelPos As Size
    Dim OldPen As Long
    Dim lngResult As Long
    
    Dim lngX1 As Long, lngX2 As Long
    Dim lngY1 As Long, lngY2 As Long
    
    On Error Resume Next
    OldDC = SaveDC(hDC)
    DGCooSys hDC, w, h
    
    ' Get the font from the picture box control (Arial Narrow)
    lngFont = SelectObject(hDC, GetStockObject(SYSTEM_FONT))
    lngResult = GetObjectAPI(lngFont, Len(lfLogFont), lfLogFont)
    lngResult = SelectObject(hDC, lngFont)
    
    'create new logical font
    ' The following defines the font size
    lfLogFont.lfWidth = FontWidth
    lfLogFont.lfHeight = FontHeight
    lngNewFont = CreateFontIndirect(lfLogFont)
    
    'select newly created logical font to DC
    lngOldFont = SelectObject(hDC, lngNewFont)
    
    ' Display the Data point name, if it exists
    If mHotSpotDataIndex >= 0 Then
        strLabelID = mDataSeries(mHotSpotSeriesIndex).Data(mHotSpotDataIndex).Description
    End If
    
    If Len(strLabelID) = 0 Then
        strLabelID = " "
    End If
    
    lngResult = GetTextExtentPoint32(hDC, strLabelID, Len(strLabelID), szLabelID)
    
    ' Position strLabelID in the upper left
    If ShowTickMarkLabels Then
        lngX1 = -1300
        lngX2 = mLDfXE
    Else
        lngX1 = 0
        lngX2 = mLDfXE
    End If
    lngY1 = mLDfYE + FontHeight * 2
    lngY2 = lngY1 - szLabelID.cy
    
    ' Clear whatever was written before; just draw regular rectangle
    ' Need to take lngX2 * 2 since not always overwriting the stuff I want to overwrite
    OldPen = SelectObject(hDC, hWhitePen)             'invisible pen
    Rectangle hDC, lngX1, lngY1, lngX2 * 2, lngY2
    lngResult = SelectObject(hDC, OldPen)
    
    ' Display strLabelID
    lngResult = TextOut(hDC, lngX1, lngY1, strLabelID, Len(strLabelID))
    
    If ShowPosition Then
        ' Diplay the current coordinates
        strLabelPos = "Coordinates: " & mCurrX & ",  " & mCurrY
        lngResult = GetTextExtentPoint32(hDC, strLabelID, Len(strLabelPos), szLabelPos)
        
        ' Position strLabelID on the left, below the ID
        ' Use the same lngX1 and lngX2 values as above
        lngY1 = mLDfYE + FontHeight * 2 - szLabelID.cy
        lngY2 = lngY1 - szLabelPos.cy
        
        'clear whatever was written before; just draw regular rectangle
        OldPen = SelectObject(hDC, hWhitePen)             'invisible pen
        Rectangle hDC, lngX1, lngY1, lngX2, lngY2
        lngResult = SelectObject(hDC, OldPen)
    
        ' Display strLabelID
        lngResult = TextOut(hDC, lngX1, lngY1, strLabelPos, Len(strLabelPos))
        
    End If
    
    'restore old font and RestoreDC
    lngResult = SelectObject(hDC, lngOldFont)
    DeleteObject (lngNewFont)
    lngResult = RestoreDC(hDC, OldDC)
End Sub

Private Sub WriteHMarkers(ByVal hDC As Long, Optional lngMarkerCount As Long = 5)
    Dim HDelta As Double
    Dim LDelta As Long
    Dim strNewLabel As String
    Dim szLbl As Size
    Dim ptPoint As POINTAPI
    Dim lngResult As Long, OldPen As Long
    Dim lngLabelIndex As Long

    If lngMarkerCount < 2 Then lngMarkerCount = 2
    LDelta = mLDfXE / (lngMarkerCount - 1)
    
    If SwapAxes Then
        HDelta = (mMaxY - mMinY) / (lngMarkerCount - 1)
    Else
        HDelta = (mMaxX - mMinX) / (lngMarkerCount - 1)
    End If
    
    For lngLabelIndex = 0 To lngMarkerCount - 1
        If ShowTickMarkLabels Then
            If SwapAxes Then
                strNewLabel = Format(mMinY + lngLabelIndex * HDelta, HNumFmt)
            Else
                strNewLabel = Format(mMinX + lngLabelIndex * HDelta, HNumFmt)
            End If
            lngResult = GetTextExtentPoint32(hDC, strNewLabel, Len(strNewLabel), szLbl)
            lngResult = TextOut(hDC, lngLabelIndex * LDelta - szLbl.cy \ 2 - 500, -200, strNewLabel, Len(strNewLabel))
        End If
        
        ' Draw the tick mark
        lngResult = MoveToEx(hDC, lngLabelIndex * LDelta, -50, ptPoint)
        lngResult = LineTo(hDC, lngLabelIndex * LDelta, 100)
        
        If ShowGridLines Then
            ' Draw the grid lines
            OldPen = SelectObject(hDC, hGridDashPen)
            lngResult = MoveToEx(hDC, lngLabelIndex * LDelta, 50, ptPoint)
            lngResult = LineTo(hDC, lngLabelIndex * LDelta, mLDfYE)
            lngResult = SelectObject(hDC, OldPen)
        End If
    Next lngLabelIndex
End Sub

Private Sub WriteVMarkers(ByVal hDC As Long, Optional lngMarkerCount As Long = 5)
    Dim VDelta As Double
    Dim LDelta As Long
    Dim strNewLabel As String
    Dim szLbl As Size
    Dim ptPoint As POINTAPI
    Dim lngResult As Long
    Dim OldPen As Long
    Dim lngLabelIndex As Long
    
    If lngMarkerCount < 2 Then lngMarkerCount = 2
    LDelta = mLDfYE / (lngMarkerCount - 1)
    
    If SwapAxes Then
        VDelta = (mMaxX - mMinX) / (lngMarkerCount - 1)
    Else
        VDelta = (mMaxY - mMinY) / (lngMarkerCount - 1)
    End If
    
    For lngLabelIndex = 0 To lngMarkerCount - 1
        If ShowTickMarkLabels Then
            If SwapAxes Then
                strNewLabel = Format(mMinX + lngLabelIndex * VDelta, VNumFmt)
            Else
                strNewLabel = Format(mMinY + lngLabelIndex * VDelta, VNumFmt)
            End If
            lngResult = GetTextExtentPoint32(hDC, strNewLabel, Len(strNewLabel), szLbl)
            lngResult = TextOut(hDC, -szLbl.cx - 100, lngLabelIndex * LDelta + 400, strNewLabel, Len(strNewLabel))
        End If
        
        ' Draw the tick mark
        lngResult = MoveToEx(hDC, -50, lngLabelIndex * LDelta, ptPoint)
        lngResult = LineTo(hDC, 50, lngLabelIndex * LDelta)
        
        If ShowGridLines Then
            ' Draw the grid lines
            OldPen = SelectObject(hDC, hGridDashPen)
            lngResult = MoveToEx(hDC, -50, lngLabelIndex * LDelta, ptPoint)
            lngResult = LineTo(hDC, mLDfXE, lngLabelIndex * LDelta)
            lngResult = SelectObject(hDC, OldPen)
        End If
    Next lngLabelIndex
    
End Sub


Public Function WriteSToFile(ByVal FileName As String) As Boolean
    '---------------------------------------------------------------
    'write data to text semicolon delimited file
    'returns True if successful, False if not or any error occurs
    'NOTE: data is appended to a file; create file if not found
    '---------------------------------------------------------------
    Dim intSeriesIndex As Integer
    Dim lngDataPointIndex As Long
    Dim hfile As Integer
    
On Error GoTo exit_WriteSToFile
    
    hfile = FreeFile
    Open FileName For Append As hfile
    Print #hfile, "Presenting - " & HLabel & ", " & VLabel
    Print #hfile, "Range - [" & mMinX & ", " & mMaxX & "]X[" & mMinY & ", " & mMaxY & "]"
    Print #hfile, vbCrLf
    Print #hfile, "ID;X;Y"
    
    For intSeriesIndex = 0 To mSeriesCount - 1
        With mDataSeries(intSeriesIndex)
            For lngDataPointIndex = 0 To .DataCount - 1
                With .Data(lngDataPointIndex)
                    Print #hfile, .Description & ";" & .x & ";" & .y
                End With
            Next lngDataPointIndex
        End With
    Next intSeriesIndex
    
    Close hfile
    WriteSToFile = True
    Exit Function
    
exit_WriteSToFile:
    MsgBox "Error in LaSpots->AddSpotsManyWithXExtents: " & Err.Description
End Function

' Property Let/Get statements

Public Property Get BorderSize() As Long
    BorderSize = mBorderSize
End Property

Public Property Let BorderSize(ByVal NewBorderSize As Long)
    mBorderSize = NewBorderSize
End Property

Public Property Get MaxSpotSizeLogicalCoords() As Long
    MaxSpotSizeLogicalCoords = mMaxSpotSize
End Property

Public Property Let MaxSpotSizeLogicalCoords(ByVal NewSize As Long)
    mMaxSpotSize = NewSize
    mHalfSize = mMaxSpotSize / 2
    mQuarterSize = mMaxSpotSize / 4
End Property

Public Property Get ViewWindow() As sSpotsView
    ViewWindow = mViewWindow
End Property

Public Property Let ViewWindow(ByVal NewViewWindow As sSpotsView)
    mViewWindow = NewViewWindow
    RefreshPlot
End Property

Private Sub mnuF_Click()
    UpdateDynamicSpotsNearCursorMenu
End Sub

Private Sub mnuFClearSelection_Click()
    SelectDeselectAll False
End Sub

Private Sub mnuDFCopy_Click()
    Call CopyFD
End Sub

Private Sub mnuFInfoOnSelection_Click()
    DisplaySpotInfoOnSelected
End Sub

Private Sub mnuFInfoOnSpotsNearCursor_Click()
    ' Display information about the spots near the cursor
    DisplaySpotInfoOnNearbySpots
End Sub

Private Sub mnuFSelectAll_Click()
    SelectDeselectAll True
End Sub

Private Sub mnuFSSzDecrease_Click()
    If mMaxSpotSize > 10 Then
        mMaxSpotSize = mMaxSpotSize - mMaxSpotSize \ 3
        mHalfSize = mMaxSpotSize \ 2
        mQuarterSize = mMaxSpotSize \ 4
        RefreshPlot
    End If
End Sub

Private Sub mnuFSSzIncrease_Click()
    mMaxSpotSize = mMaxSpotSize + mMaxSpotSize \ 3
    mHalfSize = mMaxSpotSize \ 2
    mQuarterSize = mMaxSpotSize \ 4
    RefreshPlot
End Sub

Private Sub mnuFViewCoo_Click()
    'toggle visibility of coordinates
    ShowPosition = Not ShowPosition
    mnuFViewCoo.Checked = ShowPosition
    If Not ShowPosition Then WriteCoordinates picD.hDC, picD.ScaleWidth, picD.ScaleHeight  'this will clear coordinates
End Sub

Private Sub mnuSpotInfo_Click(Index As Integer)
    HandleSpotInfoClick Index
End Sub

Private Sub picD_Click()
    On Error Resume Next
    If mHotSpotDataIndex >= 0 Then
        With mDataSeries(mHotSpotSeriesIndex)
            .Data(mHotSpotDataIndex).Selected = Not .Data(mHotSpotDataIndex).Selected
        End With
        RefreshPlot
    End If
End Sub

Private Sub picD_KeyDown(KeyCode As Integer, Shift As Integer)
    '-------------------------------------------------------------
    'copy as metafile on system clipboard on Ctrl+C combination
    '-------------------------------------------------------------
    Dim CtrlDown As Boolean
    CtrlDown = (Shift And vbCtrlMask) > 0
    If KeyCode = vbKeyC And CtrlDown Then CopyFD
End Sub

Private Sub picD_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then UserControl.PopupMenu mnuF
End Sub

Private Sub picD_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    paPoints(0).x = x
    paPoints(0).y = y
    DeviceLogicalCoordsConversion 0, 1
    TrackHotSpot paPoints(0).x, paPoints(0).y
    If ShowPosition Then WriteCoordinates picD.hDC, picD.ScaleWidth, picD.ScaleHeight
End Sub

Private Sub picD_Paint()
    RefreshPlot
End Sub

Private Sub picD_Resize()
    SetLogicalCoordinates
End Sub

Private Sub UserControl_Initialize()
    
    ReDim paPoints(0)

    FontWidth = 350
    FontHeight = 800
    VLabel = "?"
    HLabel = "?"
    HNumFmt = "0.00"
    VNumFmt = "0"
    ShowTickMarkLabels = False
    ShowGridLines = True
    CenterTicksAroundPlotCenter = True
    
    IntensityLogScale = True
    
    SetLogicalCoordinates
    
    Call ClearGraphAndData

    mDefaultSpotShape = sCircle
    mDefaultRegColor = vbGreen
    mDefaultSelColor = vbRed
    
    ShowPosition = mnuFViewCoo.Checked
    mViewWindow = sFixedWindow
    mFixedMinX = 0
    mFixedMaxX = 1
    mFixedMinY = 0
    mFixedMaxY = 10000
    
    SetAxisTickMarkCount -1, -1
    PositionControls
    
    LoadBrushColors (0)
    
    mSpotMatchCount = 0
    ReDim mSpotMatches(1, 1)
    mSpotInfoMenuCountLoaded = 1
    
End Sub

Private Sub UserControl_Resize()
    PositionControls
End Sub

Private Sub UserControl_Terminate()
    DestroyDrawingObjects
End Sub
