Attribute VB_Name = "Module21"
'created as an unified version of overlay functions
'it covers both overlays of individual spots and unique mass classes
'created: 12/20/2002 nt
'last modified: 12/27/2002 nt
'-------------------------------------------------------------------
Option Explicit

Public Const OlyCallerID = -1
'''Public Const OlyUMCCallerID = -2

'
Public Const OrientMWVrtETHrz = 0
Public Const OrientMWHrzETVrt = 1

Public Type OverlayOptions
    DefType As OverlayType
    DefShape As OverlayShape
    DefColor As Long
    DefVisible As Boolean
    DefMinSize As Single                    'percentage of logical resolution
    DefMaxSize As Single                    'to be used as min & max sizes
    DefFontWidth As Single                  'percentage of logical window to be used for
    DefFontHeight As Single                 'font width and height
    DefTextHeight As Single                 'height of the text displayed with spots
                                            'compared with the DefTextHeight
    DefStickWidth As Double
    DefMinNET As Double
    DefMaxNET As Double
    DefNETAdjustment As Long                'type of NET adjustment
    DefNETAdjustmentDisplayInd As Long      'index of display on which this adjustment is based on
    DefNETTol As Double                     'NET tolerance is used only with sticks in which case
                                            'represent accuratly on the drawing
    DefUniformSize As Boolean
    DefBoxSizeAsSpotSize As Boolean
    DefWithID As Boolean
    DefCurrScopeVisible As Boolean          'if True only current scope of created
                                            'overlay will be visible
    BackColor As Long
    ForeColor As Long
    Orientation As Long
    GRID As LaAutoGrid
End Type

Public Type OverlayStructure
    Name As String
    Comment As String
    DisplayInd As Long
    DisplayCaption As String
    Type As OverlayType
    Shape As OverlayShape
    Color As Long
    UniformSize As Boolean
    BoxSizeAsSpotSize As Boolean
    TextHeightPct As Double          'height of the text relative with coordinate axes text
                                     'this way we can draw different display with different
                                     'caption sizes
    ZOrder As Long
    NETAdjustment As Long
    NETDisplayInd As Long            'index of display to adjust with(if applicable)
    NETSlope As Double
    NETIntercept As Double
    NETFit As Double
    NETTol As Double            'start net tolerance(use only with stick)
    minNET As Double            'coordinates necesary to initialize coordinate system
    maxNET As Double
    MinMW As Double
    MaxMW As Double
    MinAbu As Double
    MaxAbu As Double
    Visible As Boolean
    ShowText As Boolean        'False by default
End Type


'---------------------------------------------------------------------------
'for unique mass classes being overlaid coordinates are used as follows
'X, Y are the coordinates of the class representative; XL,YL are coordinates
'of the first point in the class and XU, YU are the coordinates of the
'last spot in the UMC
'---------------------------------------------------------------------------
Public Type OverlayCoo
    DataCnt As Long
    CSCnt As Long
    IsoCnt As Long
    x() As Single
    XL() As Single            'adjustment before    - box only
    XU() As Single            'adjustment after     - box only
    y() As Single
    YL() As Single            'adjustment lower     - box only
    YU() As Single            'adjustment upper     - box only
    R() As Single             '                     - spots
    Visible() As Long
    OutOfScope() As Boolean
    Text() As String          'overlay string that can be displayed
End Type

Public Type OverlayAdjustment
    NETL() As Single
    NETU() As Single
    MW_ppmL() As Single             'not used for now
    MW_ppmU() As Single
End Type

Public Type OverlayJiggyOptions
    UseMWConstraint As Boolean
    MWTol As Double
    UseNetConstraint As Boolean
    NETTol As Double
    UseAbuConstraint As Boolean
    AbuTol As Double
    JiggyScope As Long
    JiggyType As Long
    BaseDisplayInd As Long                  'index of base display (in Oly); -1 if none
End Type

Public OlyCnt As Long                       'count of overlaid displays
Public Oly() As OverlayStructure
Public OlyCoo() As OverlayCoo
Public OlyAdj() As OverlayAdjustment        'used only with sticks
Public OlyOptions As OverlayOptions
Public OlyJiggyOptions As OverlayJiggyOptions

' Unused Function (May 2003)
'''Public Function IsOverlaid(ByVal Ind As Long) As Boolean
''''-------------------------------------------------------
''''return True if display Ind is already overlaid
''''-------------------------------------------------------
'''Dim i As Long
'''On Error GoTo exit_IsOverlaid
'''For i = 0 To OlyCnt - 1
'''    If Oly(i).DisplayInd Then
'''       IsOverlaid = True
'''       GoTo exit_IsOverlaid
'''    End If
'''Next i
'''exit_IsOverlaid:
'''End Function

Public Function AddDisplayToOverlay(ByVal Ind As Long) As Boolean
'----------------------------------------------------------------
'creates overlay based on display Ind; returns True if succesful
'----------------------------------------------------------------
On Error GoTo exit_AddDisplayToOverlay
'check if UMC should be overlaid that LC-MS Features exist for specified display
If OlyOptions.DefType = OlyUMC Then
   If GelUMC(Ind).UMCCnt <= 0 Then
      MsgBox "No unique mass classes found for selected display.", vbOKOnly, glFGTU
      Exit Function
   End If
End If
OlyCnt = OlyCnt + 1
ReDim Preserve Oly(OlyCnt - 1)
With Oly(OlyCnt - 1)
    .DisplayInd = Ind
    .DisplayCaption = GelBody(Ind).Caption
    .ZOrder = OlyCnt - 1               'put new overlay at the bottom
    .Name = "Overlaid " & Ind
    .Color = OlyOptions.DefColor
    .Type = OlyOptions.DefType
    .Shape = OlyOptions.DefShape
    .NETAdjustment = OlyOptions.DefNETAdjustment
    .NETDisplayInd = OlyOptions.DefNETAdjustmentDisplayInd
    .NETTol = OlyOptions.DefNETTol
    .UniformSize = OlyOptions.DefUniformSize
    .BoxSizeAsSpotSize = OlyOptions.DefBoxSizeAsSpotSize
    .TextHeightPct = OlyOptions.DefTextHeight
    .Visible = OlyOptions.DefVisible
    If OlyCnt = 1 Then .Comment = "Generated by: " & GetMyNameVersion() & " on " & Now() & vbCrLf
    .Comment = .Comment & "Overlaid: " & .Name & " - Overlaid display: " & .DisplayCaption & vbCrLf
    .MinMW = GelData(Ind).MinMW
    .MaxMW = GelData(Ind).MaxMW
    .MinAbu = GelData(Ind).MinAbu
    .MaxAbu = GelData(Ind).MaxAbu
    If Not GetOverlayNETAdjustment(OlyCnt - 1) Then
       MsgBox "Selected display can not be overlaid with specified adjustment. Make sure that adjustments are precalculated for database overlaid.", vbOKOnly, glFGTU
       OlyCnt = OlyCnt - 1
       If OlyCnt > 0 Then
          ReDim Preserve Oly(OlyCnt - 1)
       Else
          Erase Oly
       End If
       Exit Function
    End If
    Call AddEditOlyClr(OlyCnt - 1, .Color)
End With
'initialize coordinate system
ReDim Preserve OlyCoo(OlyCnt - 1)
ReDim Preserve OlyAdj(OlyCnt - 1)
AddDisplayToOverlay = InitOverlayCooAdj(OlyCnt - 1, Ind)
exit_AddDisplayToOverlay:
End Function
'
'
Public Function RemoveZOrderPositionFromOverlay(ByVal ZPosition As Long) As Boolean
'----------------------------------------------------------------------------------
'removes display on ZPosition in the overlay; returns True if successful
'----------------------------------------------------------------------------------
Dim i As Long
Dim ZPositionFound As Boolean
On Error Resume Next
For i = 0 To OlyCnt - 1
    If Oly(i).ZOrder > ZPosition Then Oly(i).ZOrder = Oly(i).ZOrder - 1
    If ZPositionFound Then
       Oly(i - 1) = Oly(i)
       OlyCoo(i - 1) = OlyCoo(i - 1)
       OlyAdj(i - 1) = OlyAdj(i - 1)
    Else
       If Oly(i).ZOrder = ZPosition Then
          ZPositionFound = True
          Call RemoveOlyClr(i - 1)
       End If
    End If
Next i
If ZPositionFound Then
   OlyCnt = OlyCnt - 1
   ReDim Preserve Oly(OlyCnt - 1)
   ReDim Preserve OlyCoo(OlyCnt - 1)
   ReDim Preserve OlyAdj(OlyCnt - 1)
   RemoveZOrderPositionFromOverlay = True
End If
End Function
'
' Unused Function (May 2003)
'''Public Function RemoveDisplayAllFromOverlay(ByVal Ind As Long) As Boolean
''''--------------------------------------------------------------------------
''''removes all displays with index Ind from overlay; returns True if OK; this
''''is neccessary - more than one instance of a display could exist in overlay
''''--------------------------------------------------------------------------
'''Dim i As Long
'''Dim Done As Boolean
'''Do Until Done
'''   For i = 0 To OlyCnt - 1
'''       If Oly(i).DisplayInd = Ind Then
'''          RemoveZOrderPositionFromOverlay Oly(i).ZOrder
'''          Exit For
'''       End If
'''   Next i
'''   If i >= OlyCnt Then Done = True
'''Loop
'''End Function
'
'
Public Function GetOlyIndFromZOrder(ZPosition As Long) As Long
'---------------------------------------------------------------
'returns index in array Oly for specified ZPosition; -1 on error
'---------------------------------------------------------------
Dim i As Long
On Error Resume Next
GetOlyIndFromZOrder = -1
For i = 0 To OlyCnt - 1
    If Oly(i).ZOrder = ZPosition Then
       GetOlyIndFromZOrder = i
       Exit Function
    End If
Next i
End Function
'
Private Function InitOverlayCooAdj(ByVal OlyInd As Long, DisplayInd As Long) As Boolean
'------------------------------------------------------------------------------------
'initializes overlay coordinates for display DisplayInd as OlyInd overlaid structure
'and adjustment arrays if we are gonna work some freakin' scary stuff with jiggy
'------------------------------------------------------------------------------------
On Error GoTo exit_InitOverlayCooAdj
With OlyCoo(OlyInd)
    Select Case Oly(OlyInd).Type
    Case olySolo
         .CSCnt = GelData(DisplayInd).CSLines
         .IsoCnt = GelData(DisplayInd).IsoLines
         .DataCnt = .CSCnt + .IsoCnt
    Case OlyUMC
         .DataCnt = GelUMC(DisplayInd).UMCCnt
         .CSCnt = 0
         .IsoCnt = 0
    End Select
    ReDim .x(.DataCnt - 1)
    ReDim .XL(.DataCnt - 1)
    ReDim .XU(.DataCnt - 1)
    ReDim .y(.DataCnt - 1)
    ReDim .YL(.DataCnt - 1)
    ReDim .YU(.DataCnt - 1)
    ReDim .R(.DataCnt - 1)
    ReDim .Visible(.DataCnt - 1)
    ReDim .OutOfScope(.DataCnt - 1)
    ReDim .Text(.DataCnt - 1)
End With
With OlyAdj(OlyInd)
    ReDim .NETL(OlyCoo(OlyInd).DataCnt - 1)
    ReDim .NETU(OlyCoo(OlyInd).DataCnt - 1)
    'ReDim .MW_ppmL(OlyCoo(OlyInd).DataCnt - 1)
    'ReDim .MW_ppmU(OlyCoo(OlyInd).DataCnt - 1)
End With
Call InitOlyAdjustmentFixSym(OlyInd)
Call InitOlyVisibility(OlyInd)
InitOverlayCooAdj = True
exit_InitOverlayCooAdj:
End Function
'
'
Private Function InitOlyAdjustmentFixSym(ByVal OlyInd As Long) As Boolean
'-----------------------------------------------------------------------
'initialize overlay adjustment as a fixed value for all spots
'-----------------------------------------------------------------------
Dim i As Long
With OlyAdj(OlyInd)
    ReDim .NETL(OlyCoo(OlyInd).DataCnt - 1)
    ReDim .NETU(OlyCoo(OlyInd).DataCnt - 1)
    'ReDim .MW_ppmL(OlyCoo(OlyInd).DataCnt - 1)
    'ReDim .MW_ppmU(OlyCoo(OlyInd).DataCnt - 1)
    For i = 0 To OlyCoo(OlyInd).DataCnt - 1
        .NETL(i) = Oly(OlyInd).NETTol
        .NETU(i) = Oly(OlyInd).NETTol
    Next i
End With
End Function
'
'
Public Function GetOlyZOrder(ZOrder() As Long) As Boolean
'--------------------------------------------------------
'fills Oly indexes in z-order; returns True if OK
'--------------------------------------------------------
Dim i As Long
On Error GoTo exit_GetOlyZOrder
If OlyCnt > 0 Then
   ReDim ZOrder(OlyCnt - 1)
   For i = 0 To OlyCnt - 1
       ZOrder(Oly(i).ZOrder) = i
   Next i
End If
GetOlyZOrder = True
exit_GetOlyZOrder:
End Function
'
'
Public Sub GetOverlayLimits(dMinNET As Double, dMinMW As Double, dMinAbu As Double, _
                            dMaxNET As Double, dMaxMW As Double, dMaxAbu As Double)
'------------------------------------------------------------------------------------
'retrieve limits of overlaid displays so that all fit in - use the largest box
'------------------------------------------------------------------------------------
Dim i As Long
dMinNET = glHugeOverExp:     dMinMW = glHugeOverExp:      dMinAbu = glHugeOverExp
dMaxNET = -glHugeOverExp:    dMaxMW = -glHugeOverExp:     dMaxAbu = -glHugeOverExp
For i = 0 To OlyCnt - 1
    With Oly(i)
        If .minNET < dMinNET Then dMinNET = .minNET
        If .MinMW < dMinMW Then dMinMW = .MinMW
        If .MinAbu < dMinAbu Then dMinAbu = .MinAbu
        If .maxNET > dMaxNET Then dMaxNET = .maxNET
        If .MaxMW > dMaxMW Then dMaxMW = .MaxMW
        If .MaxAbu > dMaxAbu Then dMaxAbu = .MaxAbu
    End With
Next i
End Sub
'
'
Private Function GetOverlayNETAdjustment(ByVal OlyInd As Long) As Boolean
'------------------------------------------------------------------------------------
'overlay adjustment on NET; this function is called only when constructing overlays
'------------------------------------------------------------------------------------
Dim FirstScan As Long, LastScan As Long, Range As Long
On Error GoTo exit_GetOverlayNETAdjustment
With Oly(OlyInd)
    GetScanRange .DisplayInd, FirstScan, LastScan, Range
    Select Case .NETAdjustment
    Case olyNETFromMinMax
         .NETSlope = (OlyOptions.DefMaxNET - OlyOptions.DefMinNET) / (LastScan - FirstScan)
         .NETIntercept = OlyOptions.DefMaxNET - .NETSlope * LastScan
         .NETFit = 0
    Case olyNETDB_TIC
         .NETSlope = GelAnalysis(.DisplayInd).NET_Slope
         .NETIntercept = GelAnalysis(.DisplayInd).NET_Intercept
         .NETFit = GelAnalysis(.DisplayInd).NET_TICFit
    Case olyNETDB_GANET
         .NETSlope = GelAnalysis(.DisplayInd).GANET_Slope
         .NETIntercept = GelAnalysis(.DisplayInd).GANET_Intercept
         .NETFit = GelAnalysis(.DisplayInd).GANET_Fit
    Case olyNETDisplay
         'have to calculate
    End Select
    'calculate also minimum and maximum
    .minNET = .NETSlope * FirstScan + .NETIntercept
    .maxNET = .NETSlope * LastScan + .NETIntercept
End With
GetOverlayNETAdjustment = True
exit_GetOverlayNETAdjustment:
End Function
'
'
Private Function InitOlyVisibility(ByVal OlyInd As Long) As Boolean
'--------------------------------------------------------------------------
'initializes visibility for OlyInd of DisplayInd; returns True if succesful
'for unique mass class is vissible if it's class representative is visible
'--------------------------------------------------------------------------
Dim i As Long, TmpCnt As Long
Dim ClsRepInd As Long, ClsRepType As Long
On Error GoTo exit_InitOlyVisibility
If OlyOptions.DefCurrScopeVisible Then
   Select Case Oly(OlyInd).Type
   Case olySolo
      With GelDraw(Oly(OlyInd).DisplayInd)
          If (.CSCount > 0 And .CSVisible) Then
             For i = 1 To .CSCount
                 TmpCnt = TmpCnt + 1
                 If .CSID(i) > 0 And .CSR(i) > 0 Then
                    OlyCoo(OlyInd).Visible(TmpCnt - 1) = True
                 Else
                    OlyCoo(OlyInd).Visible(TmpCnt - 1) = False
                 End If
             Next i
          End If
          If (.IsoCount > 0 And .IsoVisible) Then
             For i = 1 To .IsoCount
                 TmpCnt = TmpCnt + 1
                 If .IsoID(i) > 0 And .IsoR(i) > 0 Then
                    OlyCoo(OlyInd).Visible(TmpCnt - 1) = True
                 Else
                    OlyCoo(OlyInd).Visible(TmpCnt - 1) = False
                 End If
             Next i
         End If
      End With
   Case OlyUMC
      With GelDraw(Oly(OlyInd).DisplayInd)
         For i = 0 To GelUMC(Oly(OlyInd).DisplayInd).UMCCnt - 1
             ClsRepInd = GelUMC(Oly(OlyInd).DisplayInd).UMCs(i).ClassRepInd
             ClsRepType = GelUMC(Oly(OlyInd).DisplayInd).UMCs(i).ClassRepType
             Select Case ClsRepType
             Case glCSType
                  If .CSVisible Then
                     If .CSID(ClsRepInd) > 0 And .CSR(ClsRepInd) > 0 Then
                        OlyCoo(OlyInd).Visible(i) = True
                     Else
                        OlyCoo(OlyInd).Visible(i) = False
                     End If
                  End If
             Case glIsoType
                  If .IsoVisible Then
                     If .IsoID(ClsRepInd) > 0 And .IsoR(ClsRepInd) > 0 Then
                        OlyCoo(OlyInd).Visible(i) = True
                     Else
                        OlyCoo(OlyInd).Visible(i) = False
                     End If
                  End If
             End Select
         Next i
      End With
   End Select
Else
   For i = 0 To OlyCoo(OlyInd).DataCnt - 1
       OlyCoo(OlyInd).Visible(i) = True
   Next i
End If
InitOlyVisibility = True
exit_InitOlyVisibility:
End Function
'
'
' Unused Function (May 2003)
'''Public Function GetOlyTextFromID(ByVal OlyInd As Long) As Boolean
'''Dim i As Long
'''Dim TmpCnt As Long
'''Dim FullText As String
'''Dim DisplayInd As Long
'''On Error Resume Next
'''DisplayInd = Oly(OlyInd).DisplayInd
'''With OlyCoo(OlyInd)
'''    For i = 0 To .CSCnt - 1
'''        TmpCnt = TmpCnt + 1
'''        FullText = ""
'''        FullText = GelData(DisplayInd).CSData(i + 1).MTID
'''        If Len(FullText) > 20 Then
'''           .Text(TmpCnt - 1) = Left$(FullText, 20)
'''        Else
'''           .Text(TmpCnt - 1) = FullText
'''        End If
'''    Next i
'''    For i = 0 To .IsoCnt - 1
'''        TmpCnt = TmpCnt + 1
'''        FullText = ""
'''        FullText = GelData(DisplayInd).IsoData(i + 1).MTID
'''        If Len(FullText) > 20 Then
'''           .Text(TmpCnt - 1) = Left$(FullText, 20)
'''        Else
'''           .Text(TmpCnt - 1) = FullText
'''        End If
'''    Next i
'''End With
'''End Function

