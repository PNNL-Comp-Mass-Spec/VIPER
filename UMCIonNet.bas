Attribute VB_Name = "UMCIonNet"
'created: 04/10/2003 nt
'last modified: 04/10/2003 nt
'--------------------------------------------------------------------
Option Explicit

' Unused Constants
'''Public Const MASSMANIA_UMC_MW = 0
'''Public Const MASSMANIA_UMC_MW_RANGE = 1
'''Public Const MASSMANIA_UMC_MW_STD = 2
'''Public Const MASSMANIA_UMC_MW_DIFF = 3


' Unused Constants
'''Public Const REC_PLOT_BW = 0
'''Public Const REC_PLOT_COLOR = 1
'''Public Const REC_PLOT_DRAW_MAP = 0
'''Public Const REC_PLOT_DRAW_DIRECT = 1

Public Enum uindUMCIonNetDimConstants
    uindMonoMW = 0
    uindAvgMW = 1
    uindTmaMW = 2
    uindScan = 3
    uindFit = 4
    uindMZ = 5
    uindGenericNET = 6
    uindChargeState = 7
    uindLogAbundance = 8
    uindIMSDriftTime = 9
End Enum

Public Const DATA_UNITS_MASS_DA = 0
Public Const DATA_UNITS_MASS_PPM = 1

Public Const UMC_IONNET_PPM_CONVERSION_MASS = 2000

' Unused Constants
'''Public Const DATA_UMC_MW = 9            'used for recurrence plots
'''Public Const DATA_UMC_ABU = 10
'''Public Const DATA_UMC_ABU_LOG = 11
'''Public Const DATA_UMC_MW_STDEV = 12
'''Public Const DATA_UMC_MW_RANGE = 13
'''Public Const DATA_UMC_SCAN_RANGE = 14

'x=(x1,....,xn), y=(y1,...,yn) two points in n-dimensional space
'Euclidean distance is defined as d = SQRT(SUM i=1 to n (xi-yi)^2)
'Honduras distance is defined as d = SUM i=1 to n ABS(xi-yi)
'Infinity metric is defined as d= MAX(ABS(xi-yi):i=1 to n)
Public Const METRIC_EUCLIDEAN = 0
Public Const METRIC_HONDURAS = 1        'some call it Manhattan, some Trta-Mrta
Public Const METRIC_INFINITY = 2

Public Const Net_SPIDER_66 = 66

Public Const Net_CT_None = 0            'no constraint
Public Const Net_CT_LT = 1              'use only if difference Less Then
Public Const Net_CT_GT = 2              'use only if difference Greater Then
Public Const Net_CT_EQ = 3              'use only if difference equal


'class representative has to be determined even if it does not directly
'relates to how the class was created; therefore the quality of class
'representative setting is not preserved for LC-MS Features created from Nets; it
'is used just for practical purposes
Public Const UMCFROMNet_REP_ABU = 0
Public Const UMCFROMNet_REP_FIT = 1
Public Const UMCFROMNet_REP_FST_SCAN = 2
Public Const UMCFROMNet_REP_LST_SCAN = 3
Public Const UMCFROMNet_REP_MED_SCAN = 4

Public Type MetricDataDef
    Use As Boolean
    DataType As Long
    WeightFactor As Double
    ConstraintType As Long
    ConstraintValue As Double             'additional constraint on each variable
    ConstraintUnits As Long               ' DATA_UNITS_MASS_DA = 0 = Da, DATA_UNITS_MASS_PPM = 1 = ppm; Only applies to Mass-based DataTypes
End Type

Public Type UMCIonNetDefinition
    NetDim As Long                      'dimension of net space
    NetActualDim As Long                'actual dimension
    MetricType As Long                  'type of metric
    NETType As Long                     'type of net
    TooDistant As Double
    MetricData() As MetricDataDef
End Type

' Unused UDT
'''Public Type RecPlotDef
'''    MetricType As Long                  'type of metric
'''    DataType As Long                    'data to use in plot
'''    Dimension As Long
'''    Delay As Long
'''    PlotType As Long                    'black & white or colored
'''    MinClr As Long                      'transition colors used
'''    MidClr As Long                      'with colored plot
'''    MaxClr As Long
'''    ClrResolution As Long               'color resolution for color plot
'''    ThresholdBW As Double               'threshold used with black and white plot
'''    DrawMode As Long
'''End Type

' Unused UDT
'''Public Type MassManiaDefinition
'''    UMCManiaType As Long
'''    UMCScanTolerance As Long
'''    UMCRequireOverlap As Boolean
'''
'''    SoloManiaType As Long
'''    SoloScope As glScope
'''    SoloScanTolerance As Long
'''End Type

Public UMCIonNetDef As UMCIonNetDefinition           ' Default UMCIonNetDefinition

' Unused variables
'''Public MyRecPlotDef As RecPlotDef
'''Public MyMassMania As MassManiaDefinition

Public Function FindUMCClassRepIndex(ByVal lngGelIndex As Long, ByVal lngUMCIndex As Long, ByVal intUMCRepresentativeType As Integer, Optional ByRef eClassRepIonType As glDistType = gldtIS) As Long
    ' Determine the index of the ion in the given UMC that should be the class representative
    ' Returns the class rep index
    ' Note: this is the index of the ion in the UMC and NOT the index of the ion in GelData()
    ' Assumes that the members of the UMC are ordered by Scan
    '
    ' Note that this Function is very similar to UC.Bas->CalculateClassesFindRepIndex
    
    Dim i As Long
    Dim BestInd As Long
    Dim BestValue As Double
    Dim CurrValue As Double
    
On Error GoTo FindUMCClassRepIndexErrorHandler

    With GelUMC(lngGelIndex).UMCs(lngUMCIndex)
        Select Case intUMCRepresentativeType
        Case UMCFROMNet_REP_ABU
             BestInd = -1
             BestValue = -glHugeDouble
             For i = 0 To .ClassCount - 1
                 Select Case .ClassMType(i)
                 Case gldtCS
                    CurrValue = GelData(lngGelIndex).CSData(.ClassMInd(i)).Abundance
                 Case gldtIS
                    CurrValue = GelData(lngGelIndex).IsoData(.ClassMInd(i)).Abundance
                 Case Else
                    ' This shouldn't happen
                    Debug.Assert False
                    CurrValue = 0
                 End Select
                 If CurrValue > BestValue Then
                    BestValue = CurrValue
                    BestInd = i
                    eClassRepIonType = .ClassMType(i)
                 End If
             Next i
        Case UMCFROMNet_REP_FIT
             BestInd = -1
             BestValue = glHugeDouble
             For i = 0 To .ClassCount - 1
                 Select Case .ClassMType(i)
                 Case gldtCS
                    ' Isotopic fit is not defined for charge state data; use standard deviation instead
                    CurrValue = GelData(lngGelIndex).CSData(.ClassMInd(i)).MassStDev
                 Case gldtIS
                    CurrValue = GelData(lngGelIndex).IsoData(.ClassMInd(i)).Fit
                 Case Else
                    ' This shouldn't happen
                    Debug.Assert False
                    CurrValue = 0
                 End Select
                 If CurrValue < BestValue Then
                    BestValue = CurrValue
                    BestInd = i
                    eClassRepIonType = .ClassMType(i)
                 End If
             Next i
        Case UMCFROMNet_REP_FST_SCAN
             BestInd = 0
             eClassRepIonType = gldtIS
        Case UMCFROMNet_REP_LST_SCAN
             BestInd = .ClassCount - 1
             eClassRepIonType = gldtIS
        Case UMCFROMNet_REP_MED_SCAN
             BestInd = CLng(.ClassCount / 2)
             eClassRepIonType = gldtIS
        End Select
    End With
    
    If BestInd < 0 Then
        ' This shouldn't happen
        Debug.Assert False
        BestInd = 0
    End If
    
    FindUMCClassRepIndex = BestInd
    
    Exit Function

FindUMCClassRepIndexErrorHandler:
    Debug.Print "Error in FindUMCClassRepIndex: " & Err.Description
    Debug.Assert False
    LogErrors Err.Number, "frmUMCIonNet->FindUMCClassRepIndex"
    FindUMCClassRepIndex = 0
    
End Function

Public Function GetUMCIonNetInfo(Ind As Long) As String
'----------------------------------------------------------
'returns short info on IonNet for 2D display with index Ind
'----------------------------------------------------------
On Error Resume Next
Dim Info As String
Info = "Number of connections: " & GelUMCIon(Ind).NETCount & vbCrLf
If GelUMCIon(Ind).NETCount > 0 Then
   Info = Info & "Shortest connection: " & Format$(GelUMCIon(Ind).MinDist, "0.00E-00") & " = " & Format$(GelUMCIon(Ind).MinDist, "0.000") & vbCrLf
   Info = Info & "Longest connection: " & Format$(GelUMCIon(Ind).MaxDist, "0.00E-00") & " = " & Format$(GelUMCIon(Ind).MaxDist, "0.000")
End If
GetUMCIonNetInfo = Info
End Function

Public Sub LookupUMCIonNetMassTolerances(ByRef dblTolPPM As Double, ByRef eTolTypeReturn As glMassToleranceConstants, udtUMCIonNetDef As UMCIonNetDefinition, Optional ByRef eTopTypeActual As glMassToleranceConstants, Optional ByVal dblPPMConversionMass As Double = 2000)
    ' Examines udtUMCIonNetDef to determine the maximum Net_CT_LT tolerance set for any of the mass-based fields
    ' If a Da tolerance is set, converts to ppm (using a mass of dblPPMConversionMass = 2000) and returns the tolerance
    ' This is used during default naming of .Ini files using frmEditAnalysisSettings
    ' Additionally, this is used to update GelUMC().def in frmUMCIonNet.FormClassesFromNets
    
    Dim intIndex As Integer
    Dim dblTestTol As Double
    
    dblTolPPM = -1
    eTolTypeReturn = gltPPM
    eTopTypeActual = gltPPM
    
    With udtUMCIonNetDef
        ' Find the Mass field, if any
        For intIndex = 0 To .NetDim - 1
            With .MetricData(intIndex)
                If .Use Then
                    If .DataType = uindUMCIonNetDimConstants.uindMonoMW Or .DataType = uindUMCIonNetDimConstants.uindAvgMW Or .DataType = uindUMCIonNetDimConstants.uindTmaMW Then
                        If .ConstraintType = Net_CT_LT Then
                            dblTestTol = .ConstraintValue
                            If .ConstraintUnits = DATA_UNITS_MASS_DA Then
                                eTopTypeActual = gltABS
                                ' Need to convert dblTolPPM to ppm, since it is stored in the database in ppm
                                dblTestTol = MassToPPM(dblTestTol, dblPPMConversionMass)
                            End If
                            
                            If dblTestTol > dblTolPPM Then
                                dblTolPPM = dblTestTol
                            End If
                        End If
                    End If
                End If
            End With
        Next intIndex
    End With
End Sub

' The GetNetDefinition() function has been moved to frmUMCIonNet, now called GetUMCIsoDefinitionText
'''Public Function GetNetDefinition(Ind As Long) As String

