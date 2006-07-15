VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmErrorDistribution3DFromFile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Error Distribution From File"
   ClientHeight    =   3960
   ClientLeft      =   5625
   ClientTop       =   5970
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkUseDiscreteLabels 
      Caption         =   "Use discrete axis labels"
      Height          =   255
      Left            =   3240
      TabIndex        =   28
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton cmdReloadDef 
      Caption         =   "Reload Default Values"
      Height          =   375
      Left            =   720
      TabIndex        =   27
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdGraphCum 
      Caption         =   "Graph Cum. Data"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6840
      TabIndex        =   24
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdGraphNorm 
      Caption         =   "Graph Data"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6840
      TabIndex        =   23
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdViewReport 
      Caption         =   "View Report"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6840
      TabIndex        =   22
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   6840
      TabIndex        =   21
      Top             =   3240
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog fileSave 
      Left            =   480
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Text Files (*.txt)|*.txt"
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Report"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6840
      TabIndex        =   9
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox filePath 
      Height          =   315
      Left            =   3240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   600
      Width           =   5055
   End
   Begin VB.CommandButton cmdLoadFile 
      Caption         =   "Load"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog loadFile 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Text Files (*.txt)|*.txt"
   End
   Begin VB.Frame frameNET 
      Height          =   1335
      Left            =   240
      TabIndex        =   11
      Top             =   360
      Width           =   2775
      Begin VB.TextBox inputNETWidth 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox inputMaxNET 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox inputMinNET 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lableNETWidth 
         Caption         =   "NET Error Width"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label labelMaxNET 
         Caption         =   "Max NET Error"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label labelMinNET 
         Caption         =   "Min NET Error"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame frameMW 
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   2775
      Begin VB.TextBox inputMWWidth 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox inputMaxMW 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox inputMinMW 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label labelMWWidth 
         Caption         =   "MW Error Width"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label labelMinMW 
         Caption         =   "Min MW Error"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.Label labelMaxMW 
         Caption         =   "Max MW Error"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Label labelMWbinCnt 
      Caption         =   "MW Bin Count: "
      Height          =   255
      Left            =   3240
      TabIndex        =   26
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Label labelNETbinCnt 
      Caption         =   "NET Bin Count: "
      Height          =   255
      Left            =   3240
      TabIndex        =   25
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label labelInitWidth 
      Caption         =   "Note: Initial MW and NET error widths from loading are based on a 15 column/row square matrix."
      Height          =   615
      Left            =   3240
      TabIndex        =   20
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label labelCounter 
      Caption         =   "Lines of Data: "
      Height          =   255
      Left            =   3240
      TabIndex        =   19
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label labelFilePath 
      Caption         =   "File Path:"
      Height          =   255
      Left            =   3240
      TabIndex        =   18
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frmErrorDistribution3DFromFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DEFAULT_BIN_COUNT = 15
Private Const DEFAULT_NET_BIN_WIDTH = 0.005
Private Const DEFAULT_MASS_BIN_PPM = 1

Private Enum gmGraphModeConstants
    gmNormal = 0
    gmCumulative = 1
End Enum

Dim frmRepOptions As New frmErrorDistribution3DReportOptions
Dim frmGraphChart As New frmChart3D

Dim DataCnt As Long
Dim DataX() As Double       ' GANET errors
Dim DataY() As Double       ' Mass errors (ppm)
Dim matrix() As Long        ' 2D matrix of GANET and Mass errors
Dim cumMatrix() As Long

Dim minNET As Double
Dim maxNET As Double
Dim NETWidth As Double
Dim NETbinCnt As Integer
Dim MinMW As Double
Dim MaxMW As Double
Dim MWWidth As Double
Dim MWbinCnt As Integer
'default value holders
Dim DefMinNET As Double
Dim DefMaxNET As Double
Dim defNETWidth As Double
Dim defMinMW As Double
Dim defmaxmw As Double
Dim defMWWidth As Double
'declarations for View Report cmd
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32" _
    Alias "ShellExecuteA" _
   (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWDEFAULT As Long = 10
Private Const SE_ERR_NOASSOC As Long = 31

Private mGraphTitle As String

Public Sub CalculateMatrices(frmCallingForm As VB.Form, Optional lngBinCountWarn As Long = 100, Optional sngMassBinWidth As Single = -1, Optional sngGANETBinWidth As Single = -1)
'Calculate matrix and cumulative matrix based on data from DataX and DataY arrays
Dim CurrX As Double
Dim CurrY As Double
Dim i As Long
Dim j As Long
Dim k As Long
Dim L As Long
Dim CurrVal As Double
Dim MaxVal As Double
Dim eResponse As VbMsgBoxResult
Dim strCaptionSaved As String

If inputMinMW.Text = "" Or inputMaxMW.Text = "" Or inputMWWidth.Text = "" Or inputMinNET.Text = "" Or inputMaxNET.Text = "" Or inputNETWidth.Text = "" Then
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "You must fill in a value for every parameter.", vbOKOnly + vbExclamation, "Error."
    End If
Else
    
    labelNETbinCnt.Caption = "NET Bin Count: "
    labelMWbinCnt.Caption = "MW Bin Count: "
    
    MinMW = inputMinMW.Text
    MaxMW = inputMaxMW.Text
    If sngMassBinWidth > 0 Then
        MWWidth = sngMassBinWidth
    Else
        MWWidth = inputMWWidth.Text
    End If
    If MWWidth = 0 Then MWWidth = DEFAULT_MASS_BIN_PPM
    MWbinCnt = ((MaxMW - MinMW) / MWWidth) + 2
    
    minNET = inputMinNET.Text
    maxNET = inputMaxNET.Text
    If sngGANETBinWidth > 0 Then
        NETWidth = sngGANETBinWidth
    Else
        NETWidth = inputNETWidth.Text
    End If
    If NETWidth = 0 Then NETWidth = DEFAULT_NET_BIN_WIDTH
    NETbinCnt = ((maxNET - minNET) / NETWidth) + 2
    
    If (NETbinCnt > lngBinCountWarn Or MWbinCnt > lngBinCountWarn) And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        eResponse = MsgBox("NET Bin Count is " & NETbinCnt & " and MW Bin Count is " & MWbinCnt & vbCrLf & "Are you sure you want this many bins?", vbYesNo, "High Bin Count.")
        If eResponse = vbNo Then  'If they answer No
            cmdReport.Enabled = False
            cmdViewReport.Enabled = False
            cmdGraphNorm.Enabled = False
            cmdGraphCum.Enabled = False
            Exit Sub
        End If
    End If
    
    strCaptionSaved = frmCallingForm.Caption
    frmCallingForm.Caption = "Parsing data: 0 / " & DataCnt
    
    MousePointer = vbHourglass
    ReDim matrix(NETbinCnt - 1, MWbinCnt - 1) As Long
    'first do normal count
    For i = 0 To DataCnt - 1
        If DataX(i) < minNET Then
           CurrX = 0
        ElseIf DataX(i) > maxNET Then
           CurrX = NETbinCnt - 1
        Else
           CurrX = Int((DataX(i) - minNET) / NETWidth) + 1
        End If
        
        If DataY(i) < MinMW Then
           CurrY = 0
        ElseIf DataY(i) > MaxMW Then
           CurrY = MWbinCnt - 1
        Else
           CurrY = Int((DataY(i) - MinMW) / MWWidth) + 1
        End If
        
        matrix(CurrX, CurrY) = matrix(CurrX, CurrY) + 1
             
        If DataCnt Mod 50 = 0 Then
            frmCallingForm.Caption = "Parsing data: " & Trim(i) & " / " & DataCnt
            DoEvents
        End If
    Next i
    
    frmCallingForm.Caption = "Constructing matrix: 0 / " & NETbinCnt - 2
    DoEvents
    
    'do not count use border cells in this calculation since it can be misleading
    ReDim cumMatrix(NETbinCnt - 1, MWbinCnt - 1) As Long
    For i = 1 To NETbinCnt - 2
        For j = 1 To MWbinCnt - 2
            cumMatrix(i, j) = matrix(i, j)
        Next j
    Next i
    'make funnel
    MaxVal = 0
    For i = 1 To NETbinCnt - 2
        For j = 1 To MWbinCnt - 2
            CurrVal = matrix(i, j)
            For k = 1 To NETbinCnt - 2
                For L = 1 To MWbinCnt - 2
                    If k <> i Or L <> j Then
                       If matrix(k, L) < CurrVal Then
                          cumMatrix(k, L) = cumMatrix(k, L) + CurrVal
                          If cumMatrix(k, L) > MaxVal Then MaxVal = cumMatrix(k, L)
''                              Debug.Print cumMatrix(k, l) & " "
                       End If
                    End If
                Next L
            Next k
        Next j
''            Debug.Print vbCrLf
        frmCallingForm.Caption = "Constructing matrix: " & Trim(i) & " / " & NETbinCnt - 2
        DoEvents
    Next i
    'set borders to maximum value
    For i = 0 To NETbinCnt - 1
        cumMatrix(i, 0) = MaxVal
        cumMatrix(i, MWbinCnt - 1) = MaxVal
    Next i
    For j = 0 To MWbinCnt - 1
        cumMatrix(0, j) = MaxVal
        cumMatrix(NETbinCnt - 1, j) = MaxVal
    Next j
    
    cmdReport.Enabled = True
    cmdViewReport.Enabled = False
    cmdGraphNorm.Enabled = True
    cmdGraphCum.Enabled = True
    labelNETbinCnt.Caption = labelNETbinCnt.Caption & NETbinCnt
    labelMWbinCnt.Caption = labelMWbinCnt.Caption & MWbinCnt
    MousePointer = vbDefault
    
    frmCallingForm.Caption = strCaptionSaved
    
End If
Exit Sub
err_zero:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Can't specify NET Error Width or MW Error Width as zero.", vbOKOnly + vbExclamation, "Error."
    End If
End Sub

Public Sub SaveGraphPicture(blnSaveAsPNG As Boolean, Optional strFilePath As String = "", Optional lngWidthTwips As Long = -1, Optional lngHeightTwips As Long = -1)
    If lngWidthTwips > 0 And lngHeightTwips > 0 Then
        With frmGraphChart
            .ScaleMode = vbTwips
            .width = lngWidthTwips
            .Height = lngHeightTwips
        End With
    End If
    
    frmGraphChart.SaveGraphPicture blnSaveAsPNG, strFilePath
End Sub

Public Sub SetGraphTitle(strNewTitle As String)
    mGraphTitle = strNewTitle
End Sub

Public Sub cmdCalculate_Click()
CalculateMatrices Me
End Sub

Private Sub cmdClose_Click()
UnloadMyself
End Sub

Private Sub cmdGraphCum_Click()
Graph3DCumulativeData Me
End Sub

Private Sub cmdGraphNorm_Click()
Graph3DNormalData Me
End Sub

Public Function UpdateGraph3D(eGraphType As Integer, frmCallingForm As VB.Form) As Long
'Function used to create graph
'Returns 0 if success, the error code if an error

Dim i As Long
Dim j As Long
Dim xSl As Double
Dim ySl As Double
Dim strCaptionSaved As String
Dim blnUseDiscreteLabels As Boolean

On Error GoTo UpdateGraph3DErrorHandler

strCaptionSaved = frmCallingForm.Caption
MousePointer = vbHourglass

frmCallingForm.Caption = "Configuring graph"

blnUseDiscreteLabels = cChkBox(chkUseDiscreteLabels)

'Batch everything for better performance
frmGraphChart.Chart3D1.IsBatched = True

'Reset axis values for graph
frmGraphChart.Chart3D1.ChartGroups(1).RowLabels.RemoveAll
frmGraphChart.Chart3D1.ChartGroups(1).ColumnLabels.RemoveAll

With frmGraphChart.Chart3D1
    .AllowUserChanges = False
    
    .Header.Text = mGraphTitle
    
    'Set the Maximum number of Rows and Columns to Plot
    .ChartGroups(1).ElevationData.RowCount = NETbinCnt
    .ChartGroups(1).ElevationData.ColumnCount = MWbinCnt

    'Apply the various display options
    .ChartGroups(1).Elevation.IsMeshed = True
    .ChartGroups(1).Elevation.IsShaded = True
    .ChartGroups(1).Contour.IsContoured = False
    .ChartGroups(1).Contour.IsZoned = True
    .ChartArea.Surface.IsSolid = False
    .ChartGroups(1).Elevation.IsTransparent = False
    .ChartArea.Surface.IsRowMeshShowing = True
    .ChartArea.Surface.IsColumnMeshShowing = True
    .Legend = False
    .IsDoubleBuffered = True
    .ChartArea.PlotCube.Floor.Projection.IsContoured = False
    .ChartArea.PlotCube.Floor.Projection.IsZoned = False
    .ChartArea.PlotCube.Ceiling.Projection.IsContoured = False
    .ChartArea.PlotCube.Ceiling.Projection.IsZoned = False
    .ChartArea.Axes("X").AnnotationMethod = oc3dAnnotateDataLabels
    .ChartArea.Axes("Y").AnnotationMethod = oc3dAnnotateDataLabels
    .ChartLabels.add
    .ChartLabels(1).StrokeFont.Type = oc3dRomanSimplex
    .ChartLabels(1).StrokeFont.Size = glbPreferencesExpanded.ErrorPlottingOptions.Graph3DOptions.AnnotationFontSize

    'Set the scale for the X-Axis and Y-Axis
    If NETbinCnt < MWbinCnt Then
        xSl = NETbinCnt
        ySl = MWbinCnt
        ySl = (ySl / xSl) * 100
        xSl = 100
    Else
        xSl = NETbinCnt
        ySl = MWbinCnt
        xSl = (xSl / ySl) * 100
        ySl = 100
    End If
    .ChartArea.PlotCube.XScale = xSl
    .ChartArea.PlotCube.YScale = ySl

    'Set the viewing angles for the 3 Axes
    .ChartArea.View3D.XRotation = glbPreferencesExpanded.ErrorPlottingOptions.Graph3DOptions.Elevation
    .ChartArea.View3D.YRotation = glbPreferencesExpanded.ErrorPlottingOptions.Graph3DOptions.YRotation
    .ChartArea.View3D.ZRotation = glbPreferencesExpanded.ErrorPlottingOptions.Graph3DOptions.ZRotation

    'Apply the grid lines in the various directions
    With .ChartArea.Axes("X")
        .Title.Text = "NET Error"
        .MajorGrid.IsOnXYPlane = True
        .MajorGrid.IsOnXZPlane = True
    End With
    With .ChartArea.Axes("Y")
        .Title.Text = "Mass Error (ppm)"
        .MajorGrid.IsOnXYPlane = True
        .MajorGrid.IsOnYZPlane = True
    End With
    With .ChartArea.Axes("Z")
        .MajorGrid.IsOnXZPlane = True
        .MajorGrid.IsOnYZPlane = True
    End With

    With .ChartArea
        If blnUseDiscreteLabels Then
            .Axes("X").AnnotationMethod = oc3dAnnotateDataLabels
            .Axes("Y").AnnotationMethod = oc3dAnnotateDataLabels
            .Axes("Z").AnnotationMethod = oc3dAnnotateDataLabels
        Else
            .Axes("X").AnnotationMethod = oc3dAnnotateValues
            .Axes("Y").AnnotationMethod = oc3dAnnotateValues
            .Axes("Z").AnnotationMethod = oc3dAnnotateValues
        End If
    End With
    

    'Set the number of Contour Levels
    .ChartGroups(1).Contour.Levels.NumLevels = glbPreferencesExpanded.ErrorPlottingOptions.Graph3DOptions.ContourLevelsCount

    'Set the Perspective of how the graph looks
    .ChartArea.View3D.Perspective = glbPreferencesExpanded.ErrorPlottingOptions.Graph3DOptions.Perspective
End With

With frmGraphChart.Chart3D1.ChartGroups(1).ElevationData

    'Set the Origin
    .RowOrigin = minNET
    .ColumnOrigin = MinMW

    'Space between the Rows and Columns
    .RowDelta(1) = NETWidth
    .ColumnDelta(1) = MWWidth
    
    frmCallingForm.Caption = "Populating Graph: 0 / " & Trim(NETbinCnt)
            
    'Graph using a normal matrix
    If eGraphType = gmNormal Then
        For i = 0 To NETbinCnt - 1
            If blnUseDiscreteLabels Then
                If i = NETbinCnt - 1 Then
                    frmGraphChart.Chart3D1.ChartGroups(1).RowLabels.add ("+max")
                Else
                    frmGraphChart.Chart3D1.ChartGroups(1).RowLabels.add (Format(minNET + (i * .RowDelta(1)), "0.000000"))
                End If
            End If
            For j = 0 To MWbinCnt - 1
                If blnUseDiscreteLabels Then
                    If j = 0 Then
                        frmGraphChart.Chart3D1.ChartGroups(1).ColumnLabels.add ("+max")
                    Else
                        frmGraphChart.Chart3D1.ChartGroups(1).ColumnLabels.add (Format(.ColumnOrigin + ((MWbinCnt - 1 - j) * .ColumnDelta(1)), "0.000000"))
                    End If
                End If
                .Value(i + 1, j + 1) = matrix(i, MWbinCnt - 1 - j)
            Next j
            frmCallingForm.Caption = "Populating Graph: " & Trim(i) & " / " & Trim(NETbinCnt)
        Next i
    'Graph using a cumulative matrix
    ElseIf eGraphType = gmCumulative Then
        For i = 0 To NETbinCnt - 1
            If blnUseDiscreteLabels Then
                If i = NETbinCnt - 1 Then
                    frmGraphChart.Chart3D1.ChartGroups(1).RowLabels.add ("+max")
                Else
                    frmGraphChart.Chart3D1.ChartGroups(1).RowLabels.add (Format(minNET + (i * .RowDelta(1)), "0.000000"))
                End If
            End If
            For j = 0 To MWbinCnt - 1
                If blnUseDiscreteLabels Then
                    If j = 0 Then
                        frmGraphChart.Chart3D1.ChartGroups(1).ColumnLabels.add ("+max")
                    Else
                        frmGraphChart.Chart3D1.ChartGroups(1).ColumnLabels.add (Format(.ColumnOrigin + ((MWbinCnt - 1 - j) * .ColumnDelta(1)), "0.000000"))
                    End If
                End If
                .Value(i + 1, j + 1) = cumMatrix(i, MWbinCnt - 1 - j)
            Next j
            frmCallingForm.Caption = "Populating Graph: " & Trim(i) & " / " & Trim(NETbinCnt)
        Next i
    End If

End With

'Exit out of batch mode to display graph
frmGraphChart.Chart3D1.IsBatched = False
MousePointer = vbDefault
frmCallingForm.Caption = strCaptionSaved
            
UpdateGraph3D = 0
Exit Function

UpdateGraph3DErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error in UpdateGraph3D:" & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    UpdateGraph3D = Err.Number
    
End Function

Public Sub cmdLoadFile_Click()

On Error GoTo localerror

loadFile.flags = &H4 Or &H8 Or &H1 Or &H800
loadFile.ShowOpen
ReadFile (loadFile.FileName)
SetGraphTitle loadFile.FileName

Exit Sub

localerror:

End Sub

Private Function Graph3DNormalData(frmCallingForm As VB.Form, Optional blnShowForm As Boolean = True) As Long
    ' Returns 0 if success, the error code if an error

Graph3DNormalData = UpdateGraph3D(gmNormal, frmCallingForm)
frmGraphChart.Caption = "Mass Error vs. NET Error"
If blnShowForm Then frmGraphChart.Show vbModal

End Function

Private Function Graph3DCumulativeData(frmCallingForm As VB.Form, Optional blnShowForm As Boolean = True) As Long
    ' Returns 0 if success, the error code if an error

Graph3DCumulativeData = UpdateGraph3D(gmCumulative, frmCallingForm)
frmGraphChart.Caption = "Mass Error vs. NET Error"
If blnShowForm Then frmGraphChart.Show vbModal

End Function

Public Function InitializeDataUsingArrays(sngMassErrors() As Single, sngGANETErrors() As Single, lngDataCount As Long, sngMassBinWidth As Single, sngGANETBinWidth As Single, lngBinCountWarn As Long, strTitle As String, blnShowCumulativeData As Boolean, frmCallingForm As VB.Form, Optional blnShowForm As Boolean = True) As Long
    ' The input arrays should be zero-based
    ' Returns 0 if success, the error code if an error
    
    Dim lngIndex As Long
    
On Error GoTo InitializeDataUsingArraysErrorHandler

    If lngDataCount <= 1 Then
        ReDim DataX(0)
        ReDim DataY(0)
        DataCnt = 0
        Exit Function
    End If
    
    DataCnt = lngDataCount
    ReDim DataX(0 To DataCnt - 1)
    ReDim DataY(0 To DataCnt - 1)
    
    minNET = 1E+308
    maxNET = -1E+308
    MinMW = 1E+308
    MaxMW = -1E+308
    
    For lngIndex = 0 To DataCnt - 1
        DataX(lngIndex) = sngGANETErrors(lngIndex)
        DataY(lngIndex) = sngMassErrors(lngIndex)
    
        If sngGANETErrors(lngIndex) < minNET Then minNET = sngGANETErrors(lngIndex)
        If sngGANETErrors(lngIndex) > maxNET Then maxNET = sngGANETErrors(lngIndex)
        
        If sngMassErrors(lngIndex) < MinMW Then MinMW = sngMassErrors(lngIndex)
        If sngMassErrors(lngIndex) > MaxMW Then MaxMW = sngMassErrors(lngIndex)
    
    Next lngIndex
    
    'Display all the values found on the main form
    UpdateDisplayedValues
    DoEvents
    
    cmdCalculate.Enabled = True
    cmdReport.Enabled = False
    cmdViewReport.Enabled = False
    cmdGraphNorm.Enabled = False
    cmdGraphCum.Enabled = False

    inputMWWidth.Text = sngMassBinWidth
    inputNETWidth.Text = sngGANETBinWidth
    
    CalculateMatrices frmCallingForm, lngBinCountWarn, sngMassBinWidth, sngGANETBinWidth
    
    SetGraphTitle strTitle
    
    If blnShowCumulativeData Then
        InitializeDataUsingArrays = Graph3DCumulativeData(frmCallingForm, blnShowForm)
    Else
        InitializeDataUsingArrays = Graph3DNormalData(frmCallingForm, blnShowForm)
    End If
    
    InitializeDataUsingArrays = 0
    Exit Function
    
InitializeDataUsingArraysErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "An error has occured in InitializeDataUsingArrays:" & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Error"
    End If
    InitializeDataUsingArrays = Err.Number
    
End Function

Private Sub UpdateDisplayedValues()
    Dim lngBinCount As Long
    Dim lngErrorOccurrenceCount As Long
    
On Error GoTo UpdateDisplayedValuesErrorHandler

    lngBinCount = DEFAULT_BIN_COUNT
    If lngBinCount <= 0 Then lngBinCount = 15
    
    'Retain original values for Restore Default Values cmd
    DefMinNET = minNET
    DefMaxNET = maxNET
    
    ' April 2003:
    ' When this software gets compiled to a .Exe, the following call is
    '   generating an overflow, but only during AutoAnalysis; I have no idea why
    ' The error does not affect AutoAnalysis since defNETWidth and inputNETWidth.Text
    '   are not used
    defNETWidth = CDbl((maxNET - minNET)) / CDbl(lngBinCount)
    If defNETWidth <= 0 Then defNETWidth = DEFAULT_NET_BIN_WIDTH
    
    defMinMW = MinMW
    defmaxmw = MaxMW
    
    defMWWidth = CDbl((MaxMW - MinMW)) / CDbl(lngBinCount)
    If defMWWidth <= 0 Then defMWWidth = DEFAULT_MASS_BIN_PPM
    
    inputMinNET.Text = DefMinNET
    inputMaxNET.Text = DefMaxNET
    inputNETWidth.Text = defNETWidth
    inputMinMW.Text = defMinMW
    inputMaxMW.Text = defmaxmw
    inputMWWidth.Text = defMWWidth

    labelCounter.Caption = "Lines of Data: " & DataCnt
    labelNETbinCnt.Caption = "NET Bin Count: "
    labelMWbinCnt.Caption = "MW Bin Count: "

    If lngErrorOccurrenceCount > 0 Then
        Debug.Print "UpdateDisplayedValues: ErrorCount = " & lngErrorOccurrenceCount
    End If
    
    Exit Sub

UpdateDisplayedValuesErrorHandler:
    lngErrorOccurrenceCount = lngErrorOccurrenceCount + 1
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error while updating the displayed values in frmErrorDistribution3DFromFile.UpdateDisplayedValuesErrorHandler(): " & vbCrLf & Err.Description, vbExclamation + vbOKOnly, Error
    End If
    Resume Next
End Sub

Public Sub UnloadMyself()
    Unload Me
End Sub

Private Function ReadFile(FileName As String)
'Read the loaded file and populate DataX and DataY arrays
Dim fs As New FileSystemObject
Dim ts As TextStream
Dim Ln As String
Dim LnParts() As String
Dim MWvalue As Double
Dim NETvalue As Double

On Error GoTo err_BadFile
MousePointer = vbHourglass

Set ts = fs.OpenTextFile(FileName, ForReading)

DataCnt = 0
ReDim DataX(10000)
ReDim DataY(10000)

minNET = 1E+308
maxNET = -1E+308
MinMW = 1E+308
MaxMW = -1E+308

Do While Not ts.AtEndOfStream
    Ln = ts.ReadLine
    LnParts() = Split(Ln, vbTab)
    If UBound(LnParts()) >= 1 Then
        If IsNumeric(LnParts(0)) And IsNumeric(LnParts(1)) Then
            NETvalue = CDbl(LnParts(0))
            MWvalue = CDbl(LnParts(1))
    
            On Error GoTo err_ReadFile
            
            If NETvalue < minNET Then minNET = NETvalue
            If NETvalue > maxNET Then maxNET = NETvalue
            
            If MWvalue < MinMW Then MinMW = MWvalue
            If MWvalue > MaxMW Then MaxMW = MWvalue
        
            DataCnt = DataCnt + 1
            DataX(DataCnt - 1) = NETvalue
            DataY(DataCnt - 1) = MWvalue
        End If
    End If
Loop

If DataCnt = 0 Then
    MsgBox "No valid data found in the input file.  The data should be NET error and Mass Error pairs, one pair per line, separated by a tab.", vbExclamation + vbOKCancel, "Error Loading File"
Else
    
    ReDim Preserve DataX(DataCnt - 1)
    ReDim Preserve DataY(DataCnt - 1)
    
    UpdateDisplayedValues
    
    'Enable/Disable appropriate buttons
    filePath.Text = loadFile.FileName
    cmdCalculate.Enabled = True
    cmdReport.Enabled = False
    cmdViewReport.Enabled = False
    cmdGraphNorm.Enabled = False
    cmdGraphCum.Enabled = False
End If

MousePointer = vbDefault

Exit Function

err_ReadFile:
Select Case Err.Number
Case 9
    ReDim Preserve DataX(DataCnt + 5000)
    ReDim Preserve DataY(DataCnt + 5000)
    Resume
Case Else
End Select

err_BadFile:
MsgBox "An error occured while loading.  Please make sure your file's data is correctly formatted.", vbOKOnly + vbExclamation, "Error Loading File."
MousePointer = vbDefault

End Function

Public Function Report()
'Function used to create text file report based on criteria chosen from repOptions form
Dim fs As New FileSystemObject
Dim ts As TextStream
Dim i As Long
Dim j As Long
Dim k As Long
Dim L As Long
Dim m As Long
Dim sLN As String
Dim firstLN As String
On Error GoTo localerror

fileSave.flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt
fileSave.ShowSave
Set ts = fs.CreateTextFile(fileSave.FileName)

MousePointer = vbHourglass

If frmRepOptions.chkHeader Then
    ts.WriteLine ("First row contains NET Error values.  First column contains MW Error values.")
    ts.WriteLine ("Number of Data Points: " & DataCnt)
    ts.WriteLine ("NET Width: " & NETWidth)
    ts.WriteLine ("MW Width: " & MWWidth)
    ts.WriteLine ("")
End If

If frmRepOptions.inputComments.Text <> "" Then
    ts.WriteLine ("Comments: " & frmRepOptions.inputComments.Text)
    ts.WriteLine ("")
End If

If frmRepOptions.chkCountsTable Then
    firstLN = ""
    sLN = ""
    For k = 0 To NETbinCnt - 1
        If k = NETbinCnt - 1 Then
            firstLN = firstLN & vbTab & ""
        Else
            firstLN = firstLN & vbTab & Format(minNET + k * NETWidth, "0.0000")
        End If
    Next k
    ts.WriteLine (firstLN)

    For j = 0 To MWbinCnt - 1
        If j = MWbinCnt - 1 Then
            sLN = ""
        Else
            sLN = Format(MinMW + j * MWWidth, "0.0000")
        End If
        For i = 0 To NETbinCnt - 1
                sLN = sLN & vbTab & matrix(i, j)
        Next i
        ts.WriteLine (sLN)
    Next j
    
    ts.WriteLine ("")
End If

If frmRepOptions.chkCountsCol Then
    For L = 0 To NETbinCnt - 1
        For m = 0 To MWbinCnt - 1
            If L = NETbinCnt - 1 Then
                sLN = "+max" & vbTab & Format(minNET + m * NETWidth, "0.0000") & vbTab & matrix(L, m)
            ElseIf m = MWbinCnt - 1 Then
                sLN = Format(MinMW + L * MWWidth, "0.0000") & vbTab & "+max" & vbTab & matrix(L, m)
            Else
                sLN = Format(MinMW + L * MWWidth, "0.0000") & vbTab & Format(minNET + m * NETWidth, "0.0000") & vbTab & matrix(L, m)
            End If
            ts.WriteLine (sLN)
        Next m
    Next L
    
    ts.WriteLine ("")
End If

If frmRepOptions.chkCumCountsTable Then
    firstLN = ""
    sLN = ""
    For k = 0 To NETbinCnt - 1
        If k = NETbinCnt - 1 Then
            firstLN = firstLN & vbTab & ""
        Else
            firstLN = firstLN & vbTab & Format(minNET + k * NETWidth, "0.0000")
        End If
    Next k
    ts.WriteLine (firstLN)

    For j = 0 To MWbinCnt - 1
        If j = MWbinCnt - 1 Then
            sLN = ""
        Else
            sLN = Format(MinMW + j * MWWidth, "0.0000")
        End If
        For i = 0 To NETbinCnt - 1
                sLN = sLN & vbTab & cumMatrix(i, j)
        Next i
        ts.WriteLine (sLN)
    Next j
    
    ts.WriteLine ("")

End If

If frmRepOptions.chkCumCountsCol Then
    For L = 0 To NETbinCnt - 1
        For m = 0 To MWbinCnt - 1
            If L = NETbinCnt - 1 Then
                sLN = "+max" & vbTab & Format(minNET + m * NETWidth, "0.0000") & vbTab & cumMatrix(L, m)
            ElseIf m = MWbinCnt - 1 Then
                sLN = Format(MinMW + L * MWWidth, "0.0000") & vbTab & "+max" & vbTab & cumMatrix(L, m)
            Else
                sLN = Format(MinMW + L * MWWidth, "0.0000") & vbTab & Format(minNET + m * NETWidth, "0.0000") & vbTab & cumMatrix(L, m)
            End If
            ts.WriteLine (sLN)
        Next m
    Next L
    
    ts.WriteLine ("")

End If

ts.Close

cmdViewReport.Enabled = True
MousePointer = vbDefault
Beep

Exit Function

localerror:

Call cmdReport_Click

End Function

Private Sub cmdReloadDef_Click()
'Reload original min/max/width values to main form
inputMinNET.Text = DefMinNET
inputMaxNET.Text = DefMaxNET
inputNETWidth.Text = defNETWidth
inputMinMW.Text = defMinMW
inputMaxMW.Text = defmaxmw
inputMWWidth.Text = defMWWidth

End Sub

Private Sub cmdReport_Click()

    frmRepOptions.Show vbModal
    DoEvents
    If Not frmRepOptions.roCancel Then Report

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmRepOptions
Unload frmGraphChart
End Sub

Public Sub RunShellExecute(sTopic As String, _
                           sFile As Variant, _
                           sParams As Variant, _
                           sDirectory As Variant, _
                           nShowCmd As Long)

   Dim hWndDesk As Long
   Dim success As Long
  
  'the desktop will be the
  'default for error messages
   hWndDesk = GetDesktopWindow()
  
  'execute the passed operation
   success = ShellExecute(hWndDesk, sTopic, sFile, sParams, sDirectory, nShowCmd)

  'This is optional. Uncomment the three lines
  'below to have the "Open With.." dialog appear
  'when the ShellExecute API call fails
  'If success = SE_ERR_NOASSOC Then
  '   Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sFile, vbNormalFocus)
  'End If
   
End Sub


Private Sub cmdViewReport_Click()
'Open up window's default text editor to view report previously created
   Dim sTopic As String
   Dim sFile As String
   Dim sParams As Variant
   Dim sDirectory As Variant
    
   MousePointer = vbHourglass
    
    sTopic = "Open"
    sFile = fileSave.FileName
    sParams = 0&
    sDirectory = 0&

'      Case 4: 'open autoexec.bat with notepad
'               sTopic = "Open"
'               sFile = "C:\win\notepad.exe"
'               sParams = "C:\Autoexec.bat"
'               sDirectory = 0&

   Call RunShellExecute(sTopic, sFile, sParams, sDirectory, SW_SHOWNORMAL)
    
    MousePointer = vbDefault
End Sub

