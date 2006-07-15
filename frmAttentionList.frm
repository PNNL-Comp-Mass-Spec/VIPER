VERSION 5.00
Begin VB.Form frmAttentionList 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Build ER Attention List"
   ClientHeight    =   3345
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6375
   Icon            =   "frmAttentionList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCurrCriteria 
      Enabled         =   0   'False
      Height          =   1575
      Left            =   2280
      MultiLine       =   -1  'True
      TabIndex        =   27
      Top             =   600
      Width           =   1815
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00FFFFFF&
      Caption         =   "List Options"
      Height          =   3135
      Left            =   4200
      TabIndex        =   18
      Top             =   120
      Width           =   2055
      Begin VB.TextBox txtMOverZTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   31
         Text            =   "0.5"
         Top             =   2760
         Width           =   375
      End
      Begin VB.TextBox txtMaxScanTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   29
         Text            =   "3"
         Top             =   2490
         Width           =   375
      End
      Begin VB.OptionButton optListWhat 
         BackColor       =   &H00FFFFFF&
         Caption         =   "List light && heavy"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   25
         Top             =   2190
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton optListWhat 
         BackColor       =   &H00FFFFFF&
         Caption         =   "List heavy member"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   1910
         Width           =   1695
      End
      Begin VB.OptionButton optListWhat 
         BackColor       =   &H00FFFFFF&
         Caption         =   "List light member"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   1630
         Width           =   1575
      End
      Begin VB.ComboBox cmbOrder 
         Height          =   315
         ItemData        =   "frmAttentionList.frx":0442
         Left            =   120
         List            =   "frmAttentionList.frx":0444
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   870
         Width           =   1815
      End
      Begin VB.TextBox txtMaxPerScan 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   20
         Text            =   "5"
         Top             =   280
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "m/z Tolerance"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   30
         Top             =   2805
         Width           =   1035
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Max scan tolerance"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   28
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Include in list"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   1350
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Order by"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   630
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Max items per scan"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   300
         Width           =   1455
      End
   End
   Begin VB.Frame fraPairsInc 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pairs Inclusion"
      Height          =   975
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   2055
      Begin VB.OptionButton optPairsInc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use only pairs marked as included"
         Height          =   315
         Index           =   1
         Left            =   60
         TabIndex        =   17
         Top             =   540
         Width           =   1900
      End
      Begin VB.OptionButton optPairsInc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use all pairs"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   16
         Top             =   270
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame fraERSelection 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ER Inclusion"
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2055
      Begin VB.TextBox txtERMax 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox txtERMin 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   480
         TabIndex        =   13
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox txtERMinR 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   12
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtERMaxL 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   11
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton optER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ER between"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton optER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ER >="
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ER <="
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Make attention list for ER"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1785
      End
   End
   Begin VB.CommandButton cmdBuild 
      Caption         =   "&Build"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      ToolTipText     =   "Build and show list"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdClearList 
      Caption         =   "Cl&ear"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      ToolTipText     =   "Clears attention list"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdAddToList 
      Caption         =   "&Add"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      ToolTipText     =   "Add items to list based on current selection"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   2880
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   7
      X1              =   3600
      X2              =   4000
      Y1              =   3140
      Y2              =   3140
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000A&
      Index           =   6
      X1              =   3600
      X2              =   4000
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000A&
      Index           =   5
      X1              =   3600
      X2              =   4000
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   4
      X1              =   3600
      X2              =   4000
      Y1              =   2980
      Y2              =   2980
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   3
      X1              =   2360
      X2              =   2760
      Y1              =   3140
      Y2              =   3140
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000A&
      Index           =   2
      X1              =   2360
      X2              =   2760
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000A&
      Index           =   1
      X1              =   2360
      X2              =   2760
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   0
      X1              =   2360
      X2              =   2760
      Y1              =   2980
      Y2              =   2980
   End
   Begin VB.Label lblPairTypes 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   2280
      Width           =   1815
   End
End
Attribute VB_Name = "frmAttentionList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'builds attention list based on expression ratios in
'GelP_D_L structure
'NOTE: if this form opens there are some pairs for sure!
'attention list is built as an array that can be written
'to ASCII file as needed
'Format of file is
'MOverZ, MOverZ Delta, Scan Number, Scan Number Delta
'NOTE: ONLY UMC CASE IMPLEMENTED
'-------------------------------------------------------
'created: 03/11/2002 nt
'last modified: 04/01/2002 nt
'-------------------------------------------------------
Option Explicit

'managing attention list
Const MNG_AL_DESTROY = 0
Const MNG_AL_REDIM = 1
Const MNG_AL_REDIM_PRESERVE = 2

'type of inclusions
Const ER_LE = 0
Const ER_GE = 1
Const ER_BETWEEN = 2

Const P_ALL = 0
Const P_INC = 1

Const LIST_L = 0
Const LIST_H = 1
Const LIST_LH = 2

Const ORDER_ABU_DESC = 0
Const ORDER_ABU_ASC = 1
Const ORDER_FIT_DESC = 2
Const ORDER_FIT_ASC = 3

Dim CallerID As Long
Dim bLoading As Long

'parameters of attention list creation
Dim ProcessType As Long     '0 for UMC, 1 for Solo,
                            '-1 for not possible

Dim PairsIncType As Long    'which pairs to include - all(1) or just
                            'those marked as included(0)

Dim ERIncType As Long       'what kind of inclusion range to use
Dim ERMaxL As Double        'used with option ER_LE
Dim ERMinR As Double        'used with option ER_GE
Dim ERMin As Double         'used with option ER_BETWEEN
Dim ERMax As Double         'used with option ER_BETWEEN


Dim Order As Long           'order in which items are ordered within one scan
'Order is important when deciding what to include in list if not everything
Dim MaxPerScan As Long      'maximum number of items in list per scan
Dim ListWhat As Long        '0 light only, 1 heavy only, 2 both
Dim MaxScanTol As Long      'max for scan tolerance
Dim MOverZTol As Double     'm/z tolerance; for now it is constant for each item
                            'in attention list

Dim alCnt As Long               'attention list count
Dim alMOverZ() As Double
Dim alMOverZTol() As Double
Dim alScan() As Long
Dim alScanTol() As Long

Dim pIncInd() As Long       'array parallel to GelP_D_L arrays with 1
                            'on position i if pair i should be included
                            'in creation of list
                            
Dim pc As Long              'total count of pairs in GelP_D_L (just shortcut)

Dim Scans As Collection     'scans and scan members

Dim fso As New FileSystemObject

Private Sub cmbOrder_Click()
Order = cmbOrder.ListIndex
End Sub

Private Sub cmdAddToList_Click()
'----------------------------------------------------
'adds items to list of included pairs; this procedure
'does not actually build attention list
'----------------------------------------------------
Dim i As Long
Dim NewCriteria As String
With GelP_D_L(CallerID)
    Select Case ERIncType
    Case ER_LE
        If PairsIncType = 0 Then        'all pairs
           For i = 0 To .PCnt - 1
             If (.Pairs(i).ER <= ERMaxL) Then pIncInd(i) = 1
           Next i
           NewCriteria = "All"
        Else                            'included only
           For i = 0 To .PCnt - 1
             If .Pairs(i).STATE = glPAIR_Inc Then
               If (.Pairs(i).ER <= ERMaxL) Then pIncInd(i) = 1
             End If
           Next i
           NewCriteria = "Inc"
        End If
        NewCriteria = "ER<=" & ERMaxL & "-" & NewCriteria
    Case ER_GE
        If PairsIncType = 0 Then        'all pairs
           NewCriteria = "All"
           For i = 0 To .PCnt - 1
             If (.Pairs(i).ER >= ERMinR) Then pIncInd(i) = 1
           Next i
        Else                            'included only
           NewCriteria = "Inc"
           For i = 0 To .PCnt - 1
             If .Pairs(i).STATE = glPAIR_Inc Then
               If (.Pairs(i).ER >= ERMinR) Then pIncInd(i) = 1
             End If
           Next i
        End If
        NewCriteria = ERMinR & "<=ER" & "-" & NewCriteria
    Case ER_BETWEEN
        If PairsIncType = 0 Then        'all pairs
           NewCriteria = "All"
           For i = 0 To .PCnt - 1
             If ((.Pairs(i).ER >= ERMin) And _
                 (.Pairs(i).ER <= ERMax)) Then pIncInd(i) = 1
           Next i
        Else                            'included only
           NewCriteria = "Inc"
           For i = 0 To .PCnt - 1
             If .Pairs(i).STATE = glPAIR_Inc Then
               If ((.Pairs(i).ER >= ERMin) And _
                  (.Pairs(i).ER <= ERMax)) Then pIncInd(i) = 1
             End If
           Next i
        End If
        NewCriteria = ERMin & "<=ER<=" & ERMax & "-" & NewCriteria
    End Select
    txtCurrCriteria.Text = txtCurrCriteria.Text & NewCriteria & vbCrLf
    UpdateStatus "Inc.pairs: " & GetIncPairsCnt() & "/" & pc
End With
End Sub

Private Sub cmdBuild_Click()
'-------------------------------------------------------
'build and display attention list
'-------------------------------------------------------
Dim fname As String
Dim ts As TextStream
Dim i As Long


UpdateStatus "Preparing structures..."
Select Case ProcessType
Case 0
     Call Prepare_UMC
Case 1
     Call Prepare_Solo
Case Else
     MsgBox "Unknown type of pairs. Mission aborted.", vbOKOnly, glFGTU
     Exit Sub
End Select


Me.MousePointer = vbHourglass
UpdateStatus "Creating attention list..."
Call CreateAL
Me.MousePointer = vbDefault

If alCnt > 0 Then
   UpdateStatus "Generating report ..."
   Me.MousePointer = vbHourglass
   fname = GetTempFolder() & RawDataTmpFile
   Set ts = fso.OpenTextFile(fname, ForWriting, True)
   For i = 0 To alCnt - 1
       ts.WriteLine alMOverZ(i) & "," & alMOverZTol(i) & "," _
                    & alScan(i) & "," & alScanTol(i)
   Next i
   ts.Close
   Set ts = Nothing
   Me.MousePointer = vbDefault
   UpdateStatus "Att.List: " & alCnt
   frmDataInfo.Tag = "AL"
   frmDataInfo.Show vbModal
Else
   UpdateStatus "Number of pairs: " & pc
   MsgBox "Attention list is empty.", vbOKOnly, glFGTU
End If
End Sub

Private Sub cmdClearList_Click()
'----------------------------------------------------------
'deletes complete attention list and list of included pairs
'give user chance to change her/his mind
'----------------------------------------------------------
Dim Res As Long
On Error Resume Next
If alCnt > 0 Then
   Res = MsgBox("Attention list contains " & alCnt & " items. Delete?", vbYesNo, glFGTU)
   If Res <> vbYes Then Exit Sub
End If
ManageAL MNG_AL_DESTROY, 0
ReDim pIncInd(pc - 1)
txtCurrCriteria.Text = ""
UpdateStatus "Number of pairs: " & pc
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Activate()
'------------------------------------------------------------
'if loading activation set default settings
'------------------------------------------------------------
If bLoading Then
   CallerID = Me.Tag
   pc = GelP_D_L(CallerID).PCnt         'just to have it as a shortcut
   UpdateStatus "Number of pairs: " & pc
   Select Case GelP_D_L(CallerID).DltLblType
   Case ptNone
      lblPairTypes.Caption = "None"
      ProcessType = -1
   Case ptUMCDlt
      lblPairTypes.Caption = "UMC-Dlt"
      ProcessType = 0
   Case ptUMCLbl
      lblPairTypes.Caption = "UMC-Lbl"
      ProcessType = 0
   Case ptUMCDltLbl
      lblPairTypes.Caption = "UMC-DltLbl"
      ProcessType = 0
   Case ptS_Dlt
      lblPairTypes.Caption = "Solo-Dlt"
      ProcessType = 1
   Case ptS_Lbl
      lblPairTypes.Caption = "Solo-Lbl"
      ProcessType = 1
   Case ptS_DltLbl
      lblPairTypes.Caption = "Solo-DltLbl"
      ProcessType = 1
   Case Else
      lblPairTypes.Caption = "Unknown"
      ProcessType = -1
   End Select
   
   ERIncType = ER_GE
   PairsIncType = P_ALL
   MaxPerScan = txtMaxPerScan.Text
   ListWhat = LIST_LH
   MaxScanTol = txtMaxScanTol.Text
   
   Select Case GelP_D_L(CallerID).SearchDef.ERCalcType
   Case ectER_LOG, ectER_ALT      'symmetric ranges
      ERMaxL = -5
      ERMinR = 5
      ERMin = -5
      ERMax = 5
   Case Else                         'ratio-symmetric ranges
      ERMaxL = 0.5
      ERMinR = 2
      ERMin = 0.5
      ERMax = 2
   End Select
   txtERMaxL.Text = ERMaxL
   txtERMinR.Text = ERMinR
   txtERMin.Text = ERMin
   txtERMax.Text = ERMax
   
   Order = ORDER_ABU_DESC
   cmbOrder.ListIndex = Order
   
   MOverZTol = txtMOverZTol.Text
      
   ReDim pIncInd(pc - 1)
   bLoading = False
End If
End Sub

Private Sub UpdateStatus(ByVal Msg As String)
lblStatus.Caption = Msg
DoEvents
End Sub

Private Sub Form_Load()
bLoading = True

With cmbOrder
    .Clear
    .AddItem "Intensity Desc."
    .AddItem "Intensity Asc."
    .AddItem "Fit Desc."
    .AddItem "Fit Asc."
End With

End Sub

Private Sub optER_Click(Index As Integer)
ERIncType = Index
End Sub

Private Sub optListWhat_Click(Index As Integer)
ListWhat = Index
End Sub

Private Sub optPairsInc_Click(Index As Integer)
PairsIncType = Index
End Sub

Private Sub txtERMax_LostFocus()
Dim tmp As String
tmp = Trim$(txtERMax.Text)
If IsNumeric(tmp) Then
   ERMax = CDbl(tmp)
Else
   If Len(tmp) > 0 Then
      MsgBox "This argument should be numeric.", vbOKOnly, glFGTU
      txtERMax.SetFocus
   Else
      ERMax = glHugeDouble
      txtERMax.Text = ERMax
   End If
End If
End Sub


Private Sub txtERMaxL_LostFocus()
Dim tmp As String
tmp = Trim$(txtERMaxL.Text)
If IsNumeric(tmp) Then
   ERMaxL = CDbl(tmp)
Else
   If Len(tmp) > 0 Then
      MsgBox "This argument should be numeric.", vbOKOnly, glFGTU
      txtERMaxL.SetFocus
   Else
      ERMaxL = glHugeDouble
      txtERMaxL.Text = ERMaxL
   End If
End If
End Sub

Private Sub txtERMin_LostFocus()
Dim tmp As String
tmp = Trim$(txtERMin.Text)
If IsNumeric(tmp) Then
   ERMin = CDbl(tmp)
Else
   If Len(tmp) > 0 Then
      MsgBox "This argument should be numeric.", vbOKOnly, glFGTU
      txtERMin.SetFocus
   Else
      ERMin = -glHugeDouble
      txtERMin.Text = ERMin
   End If
End If
End Sub

Private Sub txtERMinR_LostFocus()
Dim tmp As String
tmp = Trim$(txtERMinR.Text)
If IsNumeric(tmp) Then
   ERMinR = CDbl(tmp)
Else
   If Len(tmp) > 0 Then
      MsgBox "This argument should be numeric.", vbOKOnly, glFGTU
      txtERMinR.SetFocus
   Else
      ERMinR = -glHugeDouble
      txtERMinR.Text = ERMinR
   End If
End If
End Sub

Private Function GetIncPairsCnt() As Long
'------------------------------------------------
'returnes number of pairs included in AL creation
'------------------------------------------------
Dim i As Long
Dim Cnt As Long
For i = 0 To pc - 1
    Cnt = Cnt + pIncInd(i)
Next i
GetIncPairsCnt = Cnt
End Function


Private Sub Prepare_UMC()
'----------------------------------------------------------
'prepares structure for AL creation from currently selected
'pairs - pairs are UMC-based
'----------------------------------------------------------
Dim i As Long, j As Long
Dim LInd As Long
Dim HInd As Long
Dim CurrInd As Long
Dim CurrType As Long
Dim CurrScan As String
Dim ScanTol As Double       'scan tolerance for each class member
                            'is half of class length(at least 1)
                            
If PrepareScans() Then
   With GelP_D_L(CallerID)
     For i = 0 To pc - 1
       LInd = .Pairs(i).P1
       HInd = .Pairs(i).P2
       If pIncInd(i) = 1 Then
          With GelUMC(CallerID)
            'add all elements of paired UMC classes to Scans collection
            If ListWhat = 0 Or ListWhat = 2 Then      'list light member members
               ScanTol = CLng(.UMCs(LInd).ClassCount / 2)
               If ScanTol = 0 Then ScanTol = 1
               For j = 0 To .UMCs(LInd).ClassCount - 1
                   CurrInd = .UMCs(LInd).ClassMInd(j)
                   CurrType = .UMCs(LInd).ClassMType(j)
                   Select Case CurrType
                   Case glCSType
                     CurrScan = CStr(GelData(CallerID).CSData(CurrInd).ScanNumber)
                   Case glIsoType
                     CurrScan = CStr(GelData(CallerID).IsoData(CurrInd).ScanNumber)
                   End Select
                   Call Scans(CurrScan).AddScanMember(CurrInd, CurrType, ScanTol)
               Next j
            End If
            If ListWhat = 1 Or ListWhat = 2 Then      'list heavy member members
               ScanTol = CLng(.UMCs(HInd).ClassCount / 2)
               If ScanTol = 0 Then ScanTol = 1
               For j = 0 To .UMCs(HInd).ClassCount - 1
                   CurrInd = .UMCs(HInd).ClassMInd(j)
                   CurrType = .UMCs(HInd).ClassMType(j)
                   Select Case CurrType
                   Case glCSType
                     CurrScan = CStr(GelData(CallerID).CSData(CurrInd).ScanNumber)
                   Case glIsoType
                     CurrScan = CStr(GelData(CallerID).IsoData(CurrInd).ScanNumber)
                   End Select
                   Call Scans(CurrScan).AddScanMember(CurrInd, CurrType, ScanTol)
               Next j
            End If
          End With
       End If
     Next i
   End With
Else
   MsgBox "Error initializing attention list", vbOKOnly, glFGTU
End If
End Sub

Private Sub Prepare_Solo()
'----------------------------------------------------------
'prepares structure for AL creation from currently selected
'pairs - pairs are Solo-based
'GelP_D_L structure allows only for Isotopic peaks to be
'considered when determining pairs - therefore...
'----------------------------------------------------------
Dim i As Long
Dim LInd As Long
Dim HInd As Long
Dim CurrScan As String
Dim ScanTol As Double       'scan tolerance for each class member

If PrepareScans() Then
   ScanTol = 1              'always 1 for individual pairs
   With GelP_D_L(CallerID)
     For i = 0 To pc - 1
       LInd = .Pairs(i).P1
       HInd = .Pairs(i).P2
       If pIncInd(i) = 1 Then
          If ListWhat = 0 Or ListWhat = 2 Then      'list light members
             CurrScan = CStr(GelData(CallerID).IsoData(LInd).ScanNumber)
             Call Scans(CurrScan).AddScanMember(LInd, glIsoType, ScanTol)
          End If
          If ListWhat = 1 Or ListWhat = 2 Then      'list heavy member members
             CurrScan = CStr(GelData(CallerID).IsoData(HInd).ScanNumber)
             Call Scans(CurrScan).AddScanMember(HInd, glIsoType, ScanTol)
          End If
       End If
     Next i
   End With
Else
   MsgBox "Error initializing attention list", vbOKOnly, glFGTU
End If
End Sub

Private Sub txtMaxPerScan_LostFocus()
Dim tmp As String
tmp = txtMaxPerScan.Text
If IsNumeric(tmp) Then
   MaxPerScan = CLng(Abs(tmp))
   txtMaxPerScan.Text = MaxPerScan
Else
   MsgBox "This argument should be positive integer.", vbOKOnly, glFGTU
   txtMaxPerScan.SetFocus
End If
End Sub

Private Sub txtMaxScanTol_LostFocus()
Dim tmp As String
tmp = txtMaxScanTol.Text
If IsNumeric(tmp) Then
   MaxScanTol = CLng(Abs(tmp))
   txtMaxScanTol.Text = MaxScanTol
Else
   MsgBox "This argument should be positive integer.", vbOKOnly, glFGTU
   txtMaxScanTol.SetFocus
End If
End Sub

Private Sub txtMOverZTol_LostFocus()
Dim tmp As String
tmp = txtMOverZTol.Text
If IsNumeric(tmp) Then
   MOverZTol = CDbl(Abs(tmp))
   txtMOverZTol.Text = MOverZTol
Else
   MsgBox "This argument should be positive integer.", vbOKOnly, glFGTU
   txtMOverZTol.SetFocus
End If
End Sub

Private Function PrepareScans() As Boolean
'-----------------------------------------
'prepares scans array
'-----------------------------------------
Dim i As Long
Dim sc As FScan
On Error GoTo exit_PrepareScans
Set Scans = New Collection
With GelData(CallerID)
  For i = 1 To UBound(.ScanInfo)
     Set sc = New FScan
     sc.Number = CStr(.ScanInfo(i).ScanNumber)
     Scans.add sc, sc.Number
  Next i
End With
PrepareScans = True
exit_PrepareScans:
End Function


Private Sub CreateAL()
'-----------------------------------------------------------------------
'creates attention list based on user specifications peaks that will
'participate in attention list buildup are stored in Scans colection
'Note that at this point we do not care what kind of pairs we have to do
'-----------------------------------------------------------------------
Dim DInd() As Long
Dim DType() As Byte
Dim dMisc() As Double       'contains scan tolerance for each peak

'following two arrays are used to sort peaks so that we can select
'top ItemsCnt of them
Dim SortInd() As Long       'index of scan members
Dim SortWhat() As Double    'abundance or fit

Dim sorter As QSDouble

Dim sc As FScan
Dim ItemsCnt As Long
Dim i As Long
On Error Resume Next

ManageAL MNG_AL_REDIM, 0        'initialize list
For Each sc In Scans
    If sc.Count > 0 Then        'retrieve references for this scan
       Call sc.GetScanMembers(DInd(), DType())
       Call sc.GetMiscData(dMisc)
       'sort them even if not neccessary(when all go to AL)
       ReDim SortInd(sc.Count - 1)
       ReDim SortWhat(sc.Count - 1)
       Select Case Order
       Case ORDER_ABU_DESC, ORDER_ABU_ASC       'abundance criteria
            For i = 0 To sc.Count - 1
                SortInd(i) = i
                Select Case DType(i)
                Case glCSType
                  SortWhat(i) = GelData(CallerID).CSData(DInd(i)).Abundance
                Case glIsoType
                  SortWhat(i) = GelData(CallerID).IsoData(DInd(i)).Abundance
                End Select
            Next i
       Case ORDER_FIT_DESC, ORDER_FIT_ASC       'fit criteria
            For i = 0 To sc.Count - 1
                SortInd(i) = i
                Select Case DType(i)
                Case glCSType
                  SortWhat(i) = GelData(CallerID).CSData(DInd(i)).MassStDev
                Case glIsoType
                  SortWhat(i) = GelData(CallerID).IsoData(DInd(i)).Fit
                End Select
            Next i
       End Select
       
       Set sorter = New QSDouble
       If sc.Count > 1 Then
          Select Case Order
          Case ORDER_ABU_ASC, ORDER_FIT_ASC         'order ascending
               Call sorter.QSAsc(SortWhat(), SortInd())
          Case ORDER_ABU_DESC, ORDER_FIT_DESC       'order descending
               Call sorter.QSDesc(SortWhat(), SortInd())
          End Select
       End If
       
       If sc.Count > MaxPerScan Then
          ItemsCnt = MaxPerScan
       Else
          ItemsCnt = sc.Count
       End If
       If ItemsCnt > 0 Then
          ManageAL MNG_AL_REDIM_PRESERVE, UBound(alMOverZ) + ItemsCnt + 1
          With GelData(CallerID)
            For i = 0 To ItemsCnt - 1
                alCnt = alCnt + 1
                Select Case DType(SortInd(i))
                Case glCSType
                    alMOverZ(alCnt - 1) = .CSData(DInd(SortInd(i))).AverageMW / .CSData(DInd(SortInd(i))).Charge
                Case glIsoType
                    alMOverZ(alCnt - 1) = .IsoData(DInd(SortInd(i))).MZ
                End Select
                alMOverZTol(alCnt - 1) = MOverZTol
                alScan(alCnt - 1) = CLng(sc.Number)
                If dMisc(SortInd(i)) <= MaxScanTol Then
                   alScanTol(alCnt - 1) = dMisc(SortInd(i))
                Else
                   alScanTol(alCnt - 1) = MaxScanTol
                End If
            Next i
          End With
       End If
    End If
Next
If alCnt > 0 Then
   ManageAL MNG_AL_REDIM_PRESERVE, alCnt
Else
   ManageAL MNG_AL_DESTROY, 0
End If
End Sub


Private Sub ManageAL(ByVal Mode As Long, ByVal Count As Long)
'-----------------------------------------------------------
'manages attention list memory requirements
'-----------------------------------------------------------
On Error Resume Next
Select Case Mode
Case MNG_AL_DESTROY         'destroy list
    alCnt = 0
    Erase alMOverZ
    Erase alMOverZTol
    Erase alScan
    Erase alScanTol
Case MNG_AL_REDIM
    If Count <= 0 Then Count = 100
    ReDim alMOverZ(Count)
    ReDim alMOverZTol(Count)
    ReDim alScan(Count)
    ReDim alScanTol(Count)
Case MNG_AL_REDIM_PRESERVE
    If Count > 0 Then
       ReDim Preserve alMOverZ(Count - 1)
       ReDim Preserve alMOverZTol(Count - 1)
       ReDim Preserve alScan(Count - 1)
       ReDim Preserve alScanTol(Count - 1)
    End If
End Select
End Sub
