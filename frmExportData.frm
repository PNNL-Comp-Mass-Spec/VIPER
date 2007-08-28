VERSION 5.00
Begin VB.Form frmExportData 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export Data"
   ClientHeight    =   6030
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraExpType 
      Caption         =   "Export Type"
      Height          =   1215
      Left            =   3600
      TabIndex        =   18
      Top             =   3480
      Width           =   1335
      Begin VB.OptionButton optExpType 
         Caption         =   "UMC"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   20
         ToolTipText     =   "Export unique mass classes"
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optExpType 
         Caption         =   "Data"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   19
         ToolTipText     =   "Export data peaks"
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame fraMWField 
      Caption         =   "Molecular Mass Type"
      Height          =   1335
      Left            =   2400
      TabIndex        =   14
      Top             =   2040
      Width           =   2535
      Begin VB.OptionButton optMWType 
         Caption         =   "The most abundant"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   900
         Width           =   2055
      End
      Begin VB.OptionButton optMWType 
         Caption         =   "Monoisotopic"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton optMWType 
         Caption         =   "Average"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.Frame fraNETCalculation 
      Caption         =   "NET Calculation"
      Height          =   1215
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Width           =   3255
      Begin VB.TextBox txtNETFormula 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Text            =   "(FN-MinFN)/(MaxFN-MinFN)"
         Top             =   720
         Width           =   3015
      End
      Begin VB.OptionButton optNETCalculation 
         Caption         =   "GANET"
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton optNETCalculation 
         Caption         =   "TIC"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton optNETCalculation 
         Caption         =   "Generic"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame fraScope 
      Caption         =   "Export Scope"
      Height          =   1335
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   2055
      Begin VB.OptionButton optExportScope 
         Caption         =   "Current view"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton optExportScope 
         Caption         =   "All data points"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox txtOutFile 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   4920
      Width           =   4695
   End
   Begin VB.ComboBox cmbDisplays 
      Height          =   315
      ItemData        =   "frmExportData.frx":0000
      Left            =   240
      List            =   "frmExportData.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1560
      Width           =   4695
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label lblNote 
      Caption         =   $"frmExportData.frx":0004
      Height          =   1095
      Left            =   240
      TabIndex        =   21
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label lblInfo 
      Caption         =   "Select display to export"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   3735
   End
End
Attribute VB_Name = "frmExportData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this function exports gel data in a form applicable for drawing
'created: 11/26/2002 nt
'last modified: 12/26/2002 nt
'-----------------------------------------------------------------------------
'Exported file is plain semi-colon delimited text file with header information
'File format
'Filename:
'First scan:
'Last scan:
'NET formula:
'Isotopic mass:
'NET; MW; Abundance; Text
'-----------------------------------------------------------------------------
Option Explicit

Const EXP_DATA = 0
Const EXP_UMC = 1

Dim DisplayInd As Long

Dim ExpScope As Long
Dim ExpNETType As Long
Dim ExpMWField As Integer
Dim ExpType As Long

Dim OutFileName As String

Dim fso As New FileSystemObject

Dim NETExprEva As ExprEvaluator     'expression evaluator for NET
Dim VarVals() As Long               'variable for expression evaluator

Dim FirstScan As Long               'scan range of selected display
Dim LastScan As Long
Dim ScanRange As Long

Private Sub cmbDisplays_Click()
DisplayInd = cmbDisplays.ItemData(cmbDisplays.ListIndex)
'set NET calculation formula on Generic
optNETCalculation(etGenericNET).Value = True

EnableDisableControls
End Sub

Private Sub cmdBrowse_Click()
OutFileName = SaveFileAPIDlg(Me.hwnd, "Text file (*.txt)" & Chr(0) & "*.txt" _
            & Chr(0), 1, "DataFile.txt", "Save Text File")
txtOutFile = OutFileName
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdExport_Click()

If ExpType = EXP_UMC Then
    MsgBox "This window cannot be used to export LC-MS Features to a file.  Please use Edit->Copy LC-MS Features in View->Copy to File.  Be sure the view is completely zoomed out and that all of the LC-MS Features are shown (using Tools->Show LC-MS Feature Points).", vbOKOnly, glFGTU
    Exit Sub
End If

If DisplayInd >= 0 Then
   If Len(OutFileName) > 0 Then
      If fso.FileExists(OutFileName) Then
         If MsgBox("File already exists. Overwrite?", vbYesNo, glFGTU) <> vbYes Then Exit Sub
      End If
      Select Case ExpType
      Case EXP_DATA
            Me.MousePointer = vbHourglass
            If GelData(DisplayInd).CustomNETsDefined Then
                Call ExportData
            Else
                If InitExprEvaluator(txtNETFormula.Text) Then
                    Call ExportData
                Else
                    MsgBox "Error in NET calculation formula.", vbOKOnly
                    txtNETFormula.SetFocus
                End If
            End If
            Me.MousePointer = vbDefault
           
'      Case EXP_UMC
'           If GelUMC(DisplayInd).UMCCnt > 0 Then
'              Me.MousePointer = vbHourglass
'              Call ExportUMC
'              Me.MousePointer = vbDefault
'           Else
'              MsgBox "Unique mass classes for selected display not found.", vbOKOnly, glFGTU
'           End If
      End Select
   Else
      MsgBox "Enter name of output file.", vbOKOnly, glFGTU
   End If
Else
   MsgBox "Select display to export.", vbOKOnly, glFGTU
End If
End Sub

Private Sub Form_Load()
Dim i As Long
For i = 1 To UBound(GelData)
    If Not GelStatus(i).Deleted Then
       cmbDisplays.AddItem GelBody(i).Caption
       cmbDisplays.ItemData(cmbDisplays.NewIndex) = i
    End If
Next i
ExpType = EXP_DATA
ExpScope = glScope.glSc_All
ExpMWField = mftMWMono
ExpNETType = etGenericNET
End Sub

Private Sub optExportScope_Click(Index As Integer)
ExpScope = Index
End Sub

Private Sub optExpType_Click(Index As Integer)
ExpType = Index
End Sub

Private Sub optMWType_Click(Index As Integer)
ExpMWField = Index + 6
End Sub

Private Sub optNETCalculation_Click(Index As Integer)
On Error Resume Next
ExpNETType = Index
Select Case Index
Case etGenericNET
  txtNETFormula.Text = ConstructNETFormula(0, 0, True)
Case etTICFitNET
  With GelAnalysis(DisplayInd)
     If .NET_Intercept <> 0 Or .NET_Slope <> 0 Then _
        txtNETFormula.Text = ConstructNETFormula(.NET_Slope, .NET_Intercept)
  End With
  If Err Then
     MsgBox "Make sure display is loaded as analysis. Use New Analysis command from the File menu.", vbOKOnly, glFGTU
     Exit Sub
  End If
Case etGANET
  With GelAnalysis(DisplayInd)
    If .GANET_Intercept <> 0 Or .GANET_Slope <> 0 Then _
       txtNETFormula.Text = ConstructNETFormula(.GANET_Slope, .GANET_Intercept)
  End With
  If Err Then
     MsgBox "Make sure display is loaded as analysis. Use New Analysis command from the File menu.", vbOKOnly, glFGTU
     Exit Sub
  End If
End Select
txtNETFormula.SetFocus
End Sub

Private Sub EnableDisableControls()
    If DisplayInd > 0 Then
        fraNETCalculation.Enabled = Not GelData(DisplayInd).CustomNETsDefined
        
        txtNETFormula.Enabled = fraNETCalculation.Enabled
        optNETCalculation(0).Enabled = fraNETCalculation.Enabled
        optNETCalculation(1).Enabled = fraNETCalculation.Enabled
        optNETCalculation(2).Enabled = fraNETCalculation.Enabled
    End If
End Sub

Public Sub ExportData()
Dim TmpCnt As Long
Dim ts As TextStream
Dim sL As String
Dim CurrNET As Double
Dim ScopeInd() As Long
Dim i As Long, k As Long
Dim strLabelSaved As String

On Error Resume Next
strLabelSaved = lblInfo.Caption
lblInfo.Caption = "Opening " & OutFileName
DoEvents

Set ts = fso.OpenTextFile(OutFileName, ForWriting, True)
'write file header in a format specified above
ts.WriteLine "FileName=" & GelBody(DisplayInd).Caption
ts.WriteLine "First scan=" & FirstScan
ts.WriteLine "Last scan=" & LastScan
ts.WriteLine "NET formula=" & txtNETFormula.Text
Select Case ExpMWField
Case mftMWAvg
    ts.WriteLine "Isotopic mass=Average"
Case mftMWMono
    ts.WriteLine "Isotopic mass=Monoisotopic"
Case mftMWTMA
    ts.WriteLine "Isotopic mass=The Most Abundant"
End Select
ts.WriteLine "NET; MW; Abundance; Text"
Select Case ExpScope
Case glScope.glSc_All
  With GelData(DisplayInd)
    DoEvents
    For k = 1 To .CSLines
      If GelData(DisplayInd).CustomNETsDefined Then
        CurrNET = ScanToGANET(DisplayInd, .CSData(k).ScanNumber)
      Else
        VarVals(1) = .CSData(k).ScanNumber
        CurrNET = NETExprEva.ExprVal(VarVals())
      End If
      
      sL = CurrNET & glARG_SEP & .CSData(k).AverageMW & glARG_SEP & .CSData(k).Abundance & glARG_SEP
      If Len(.CSData(k).MTID) > 0 Then sL = sL & .CSData(k).MTID
      ts.WriteLine sL
      If k Mod 500 = 0 Or k = 1 Then
        lblInfo.Caption = "Writing Charge State data: " & Trim(k) & " / " & .CSLines
        DoEvents
      End If
    Next k
    'and now Isotopic
    For k = 1 To .IsoLines
      If GelData(DisplayInd).CustomNETsDefined Then
        CurrNET = ScanToGANET(DisplayInd, .IsoData(k).ScanNumber)
      Else
        VarVals(1) = .IsoData(k).ScanNumber
        CurrNET = NETExprEva.ExprVal(VarVals())
      End If
      sL = CurrNET & glARG_SEP & GetIsoMass(.IsoData(k), ExpMWField) & glARG_SEP & .IsoData(k).Abundance & glARG_SEP
      If Len(.IsoData(k).MTID) > 0 Then sL = sL & .IsoData(k).MTID
      ts.WriteLine sL
      If k Mod 500 = 0 Or k = 1 Then
        lblInfo.Caption = "Writing Isotopic data: " & Trim(k) & " / " & .IsoLines
        DoEvents
      End If
    Next k
  End With
Case glScope.glSc_Current
  With GelData(DisplayInd)
    TmpCnt = GetScopeCS(DisplayInd, ScopeInd())
    If TmpCnt > 0 Then
       For i = 1 To TmpCnt
         k = ScopeInd(i)
         If GelData(DisplayInd).CustomNETsDefined Then
            CurrNET = ScanToGANET(DisplayInd, .CSData(k).ScanNumber)
         Else
            VarVals(1) = .CSData(k).ScanNumber
            CurrNET = NETExprEva.ExprVal(VarVals())
         End If
         sL = CurrNET & glARG_SEP & .CSData(k).AverageMW & glARG_SEP & .CSData(k).Abundance & glARG_SEP
         If Len(.CSData(k).MTID) > 0 Then sL = sL & .CSData(k).MTID
         ts.WriteLine sL
         If i Mod 500 = 0 Or i = 1 Then
           lblInfo.Caption = "Writing Charge State data: " & Trim(i) & " / " & Trim(TmpCnt)
           DoEvents
         End If
       Next i
    End If
    'and now Isotopic
    TmpCnt = GetScopeIso(DisplayInd, ScopeInd())
    If TmpCnt > 0 Then
       For i = 1 To TmpCnt
         k = ScopeInd(i)
         If GelData(DisplayInd).CustomNETsDefined Then
            CurrNET = ScanToGANET(DisplayInd, .IsoData(k).ScanNumber)
         Else
            VarVals(1) = .IsoData(k).ScanNumber
            CurrNET = NETExprEva.ExprVal(VarVals())
         End If
         sL = CurrNET & glARG_SEP & GetIsoMass(.IsoData(k), ExpMWField) & glARG_SEP & .IsoData(k).Abundance & glARG_SEP
         If Len(.IsoData(k).MTID) > 0 Then sL = sL & .IsoData(k).MTID
         ts.WriteLine sL
         If i Mod 500 = 0 Or i = 1 Then
           lblInfo.Caption = "Writing Isotopic data: " & Trim(i) & " / " & Trim(TmpCnt)
           DoEvents
         End If
       Next i
    End If
  End With
End Select
ts.Close
Set ts = Nothing
lblInfo.Caption = strLabelSaved
If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
    MsgBox "Data has been saved to: " & OutFileName
End If
End Sub

Private Sub ExportUMC()
MsgBox "This function is not implemented yet.", vbOKOnly, glFGTU
End Sub


Private Function InitExprEvaluator(ByVal sExpr As String) As Boolean
'-------------------------------------------------------------------
'initializes expression evaluator for elution time
'-------------------------------------------------------------------
On Error Resume Next
Set NETExprEva = New ExprEvaluator
NETExprEva.Vars.add 1, "FN"
NETExprEva.Vars.add 2, "MinFN"
NETExprEva.Vars.add 3, "MaxFN"
NETExprEva.Expr = sExpr
InitExprEvaluator = NETExprEva.IsExprValid
GetScanRange DisplayInd, FirstScan, LastScan, ScanRange
ReDim VarVals(1 To 3)
VarVals(2) = FirstScan
VarVals(3) = LastScan
End Function

Private Sub txtOutFile_LostFocus()
OutFileName = txtOutFile.Text
End Sub
