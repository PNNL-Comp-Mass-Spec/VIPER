VERSION 5.00
Begin VB.Form frmVisPairs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delta/Label Pairs Viewer/Editor"
   ClientHeight    =   8235
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   12585
   Icon            =   "frmVisPairs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   12585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraEdit 
      Caption         =   "Editor Functions"
      Height          =   7695
      Left            =   9240
      TabIndex        =   1
      Top             =   120
      Width           =   3300
      Begin VB.CommandButton cmdResetClasses 
         Caption         =   "&Reset"
         Enabled         =   0   'False
         Height          =   330
         Left            =   720
         TabIndex        =   44
         ToolTipText     =   "Reset classes to the state before this form was loaded!"
         Top             =   7200
         Width           =   855
      End
      Begin VB.CheckBox chkMultiMemberGroups 
         Caption         =   "List only multi-member groups"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   3480
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.Frame fraPairsArranger 
         Caption         =   "Pairs Arranger"
         Height          =   3135
         Left            =   120
         TabIndex        =   26
         Top             =   3960
         Width           =   3015
         Begin VB.CheckBox chkShowPairs 
            Caption         =   "Show Pairs"
            Height          =   195
            Left            =   240
            TabIndex        =   56
            Top             =   2800
            Width           =   1455
         End
         Begin VB.CommandButton cmdArrangeGroup 
            Caption         =   "&Res. Grp."
            Enabled         =   0   'False
            Height          =   325
            Left            =   1080
            TabIndex        =   45
            ToolTipText     =   "Resolve pairs belonging to the same group in currently listed groups"
            Top             =   2400
            Width           =   855
         End
         Begin VB.CommandButton cmdAddRemovePair 
            Caption         =   "Rem. Pair"
            Enabled         =   0   'False
            Height          =   325
            Left            =   1080
            TabIndex        =   43
            ToolTipText     =   "Remove class from the list of classes to merge"
            Top             =   2040
            Width           =   855
         End
         Begin VB.CommandButton cmdArrAddGroup 
            Caption         =   "Add &Grp."
            Enabled         =   0   'False
            Height          =   325
            Left            =   2040
            TabIndex        =   42
            Top             =   2040
            Width           =   855
         End
         Begin VB.CommandButton cmdArrangeAll 
            Caption         =   "Res. All"
            Enabled         =   0   'False
            Height          =   325
            Left            =   2040
            TabIndex        =   40
            ToolTipText     =   "Resolves all pairs belonging to the same group"
            Top             =   2400
            Width           =   855
         End
         Begin VB.CommandButton cmdArrange 
            Caption         =   "Resolve"
            Enabled         =   0   'False
            Height          =   325
            Left            =   120
            TabIndex        =   31
            ToolTipText     =   "Resolve listed pairs; updates the whole structure"
            Top             =   2400
            Width           =   855
         End
         Begin VB.CommandButton cmdArrClear 
            Caption         =   "&Clear"
            Enabled         =   0   'False
            Height          =   325
            Left            =   2040
            TabIndex        =   30
            Top             =   2760
            Width           =   855
         End
         Begin VB.CommandButton cmdArrAddPair 
            Caption         =   "&Add Pair"
            Enabled         =   0   'False
            Height          =   325
            Left            =   120
            TabIndex        =   29
            ToolTipText     =   "Add currently selected class to the list of classes to merge"
            Top             =   2040
            Width           =   855
         End
         Begin VB.ListBox lstPairsToArr 
            Height          =   1425
            Left            =   120
            TabIndex        =   27
            Top             =   540
            Width           =   2775
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Pairs to arrange"
            Height          =   195
            Left            =   1410
            TabIndex        =   28
            Top             =   300
            Width           =   1110
         End
      End
      Begin VB.ComboBox cmbLstGroups 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmVisPairs.frx":030A
         Left            =   120
         List            =   "frmVisPairs.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2760
         Width           =   3015
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   330
         Left            =   1800
         TabIndex        =   8
         Top             =   7200
         Width           =   855
      End
      Begin VB.Frame fraFunction1 
         Caption         =   "Grouping Definition"
         Height          =   2175
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3015
         Begin VB.TextBox txtf1MWTol 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            TabIndex        =   35
            Text            =   "0.02"
            Top             =   1620
            Width           =   495
         End
         Begin VB.ComboBox cmbGroupFunction 
            Height          =   315
            ItemData        =   "frmVisPairs.frx":030E
            Left            =   120
            List            =   "frmVisPairs.frx":0310
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   540
            Width           =   2775
         End
         Begin VB.TextBox txtf1MWDist 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            TabIndex        =   5
            Text            =   "50"
            Top             =   1020
            Width           =   495
         End
         Begin VB.TextBox txtf1ScanDist 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            TabIndex        =   4
            Text            =   "5"
            Top             =   1320
            Width           =   495
         End
         Begin VB.CommandButton cmdf1Group 
            Caption         =   "Group"
            Height          =   325
            Left            =   2280
            TabIndex        =   3
            Top             =   1305
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "MW tolerance"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   34
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Grouping Function"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   32
            Top             =   300
            Width           =   2655
         End
         Begin VB.Label Label2 
            Caption         =   "MW distance"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Max. scan separation"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   1380
            Width           =   1575
         End
      End
      Begin VB.Label lblGroupsCount 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   1200
         TabIndex        =   39
         Top             =   3160
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Groups count: "
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   38
         Top             =   3160
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Groups Selection"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   37
         Top             =   2520
         Width           =   1335
      End
   End
   Begin VB.PictureBox picContainer 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   8175
      Left            =   0
      ScaleHeight     =   8115
      ScaleWidth      =   9075
      TabIndex        =   10
      Top             =   0
      Width           =   9135
      Begin VB.PictureBox picDummy1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   7215
         Index           =   1
         Left            =   8880
         ScaleHeight     =   7215
         ScaleWidth      =   255
         TabIndex        =   55
         Top             =   960
         Width           =   255
      End
      Begin VB.PictureBox picDummy1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   7335
         Index           =   0
         Left            =   3480
         ScaleHeight     =   7335
         ScaleWidth      =   285
         TabIndex        =   36
         Top             =   1080
         Width           =   280
      End
      Begin VB.ListBox lstPairs 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   2550
         Left            =   3840
         TabIndex        =   24
         Top             =   1200
         Width           =   5175
      End
      Begin VB.ListBox lstGroups 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   6855
         Left            =   0
         TabIndex        =   22
         Top             =   1200
         Width           =   3735
      End
      Begin VB.ListBox lstPeaksL 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   1710
         Left            =   3840
         TabIndex        =   53
         Top             =   4200
         Width           =   5175
      End
      Begin VB.ListBox lstPeaksH 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   1710
         Left            =   3840
         TabIndex        =   54
         Top             =   6240
         Width           =   5175
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pair details - heavy member"
         ForeColor       =   &H00C0FFC0&
         Height          =   225
         Index           =   3
         Left            =   4560
         TabIndex        =   52
         Top             =   6000
         Width           =   1965
      End
      Begin VB.Label lblPairsType 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5760
         TabIndex        =   51
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pairs type:"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   7
         Left            =   4680
         TabIndex        =   50
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblCurrPairsCount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6240
         TabIndex        =   49
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblOrigPairsCount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6240
         TabIndex        =   48
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Current pairs count:"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   6
         Left            =   4680
         TabIndex        =   47
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Original pairs count:"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   5
         Left            =   4680
         TabIndex        =   46
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pairs belonging to a group"
         ForeColor       =   &H00C0FFC0&
         Height          =   225
         Index           =   1
         Left            =   4560
         TabIndex        =   25
         Top             =   960
         Width           =   1875
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Groups of pairs"
         ForeColor       =   &H00C0FFC0&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total number of peaks:"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Original UMC count:"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Current UMC count:"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblPeaksCnt 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblOrigUMCCnt 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   17
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblCurrUMCCnt 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   16
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Original UMC ratio:"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   15
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Current UMC ratio:"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   14
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblOrigUMCRatio 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3840
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblCurrUMCRatio 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pair details - light member"
         ForeColor       =   &H00C0FFC0&
         Height          =   225
         Index           =   2
         Left            =   4560
         TabIndex        =   11
         Top             =   3960
         Width           =   1815
      End
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Status: Woozup"
      Height          =   255
      Left            =   9240
      TabIndex        =   0
      Top             =   7920
      Width           =   3255
   End
End
Attribute VB_Name = "frmVisPairs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Delta/Label pairs editor; as a result of pairs editing unique
'mass classes might also be modified; this form deals with
'temporary structures and changes are accepted together
'created: 04/08/2002 nt
'last modified: 04/03/2003 nt
'-------------------------------------------------------------
Option Explicit

Const MAX_GROUPS_IN_LIST = 250

Const F_MEMBERS_SHARING = 0
Const F_MEMBERS_SHARING_GT_MW = 1
Const F_MEMBERS_SHARING_LT_MW = 2

Const REMOVE_UMC_MARK = -1

Const P_LIGHT = 0       'light pair member
Const P_HEAVY = 1       'heavy pair member

'used when calling UpdateLabels function
Const LBL_ERR = -1
Const LBL_ZLS = 0

Dim bLoading As Long
Dim CallerID As Long

Dim Tmp As UMCListType              'all work on this form is done with temporary UMC
Dim TmpInc() As Long        'array parallel with UMCs in Tmp used to determine
                            'what will be included in newly defined UMC
                            
Dim tmpPairs As IsoPairsDltLblType
                        
Dim UMCStat() As Double     'precalculated classes statistics used to easily
                            'access class properties and display descriptions
Dim UMCDisplay() As String  'class descriptions used to list classes
Dim UMCMW() As Double       'classes will be sorted on molecular masses
Dim UMCInd() As Long        'original classes sort order
Dim UMCCntL() As Long       'count of UMC as light member of a pair
Dim UMCCntH() As Long       'count of UMC as heavy member of a pair

Dim UMCPairs() As GR        'contains pairs in which UMC is a pair member
                            
Dim P_Display() As String   'pairs description that will be used on list
Dim P_Ind() As Long         'original pairs sort order
Dim P_MW() As Double        'pairs will be sorted on molecular masses for
                            'display purposes
                                                        
'results of analysis
Dim GrRes As GR2           'result is a group of group of classes

'if number of results is more than MAX_GROUPS_IN_LIST results are
'split in groups which then can be selected/displayed from combo box
Dim LstGroupsCnt As Long
Dim LstGroupsInd1() As Long             'first index of groups belonging to list
Dim LstGroupsInd2() As Long             'last index of groups belonging to list
Dim LstGroupsDisplay() As String        'display name of list portion

Dim CurrLstGroupInd As Long             'index of selected item in the groups combo box
Dim CurrGroupInd As Long                'index of selected group in Res structure
'NOTE: CurrGroupInd=CurrLstGroupInd*MAX_GROUPS_IN_LIST+CurrLstInd
Dim CurrPairInd As Long                 'index of selected pair
Dim CurrPeakIndL As Long                'index of selected peak in light member
Dim CurrPeakIndH As Long                'index of selected peak in heavy member


'parameters of grouping
'NOTE: not all parameters are used with all types of grouping, and even same
'parameter could have different interpretation with different functions
Dim f1Type As Long          'type of grouping
Dim f1MWDist As Double      'mw distance(measurement unit depends on type)
Dim f1MWTol As Double       'molecular mass tolerance when needed
Dim f1ScanDist As Long      'scan distance (scan distance 0 means classes must overlap)

Dim MultiGroupsOnly As Boolean      'if True only groups with multiple membership
                                    'will be displayed
                                    
Dim NeedToSave As Boolean       'if this flag is set ask user does it want to save

Dim WithEvents MyPairViewer As frmPairsView
Attribute MyPairViewer.VB_VarHelpID = -1

Private Sub FillComboBoxes()
    With cmbGroupFunction
        .Clear
        .AddItem "Pairs sharing UMC"
        .AddItem "Members sharing pairs > MW"
        .AddItem "Members sharing pairs < MW"
    End With
    
End Sub

Private Sub chkMultiMemberGroups_Click()
MultiGroupsOnly = (chkMultiMemberGroups.value = vbChecked)
End Sub

Private Sub chkShowPairs_Click()
On Error GoTo ChkShowPairsErrorHandler
If chkShowPairs.value = vbChecked Then
   MyPairViewer.Show
Else
   MyPairViewer.Hide
End If
Exit Sub
ChkShowPairsErrorHandler:
MsgBox "Error: Cannot show the PairView from this form: Contact Programmer!"

End Sub

Private Sub cmbGroupFunction_Click()
f1Type = cmbGroupFunction.ListIndex
'Select Case f1Type                  'special settings if neccesary
'Case F_MW_SCAN_PROXIMITY            'has sense only for multi member groups
'     chkMultiMemberGroups.value = vbChecked
'Case F_MW_SCAN_EXACT_DISTANCE       'has sense only for multi member groups
'     chkMultiMemberGroups.value = vbChecked
'Case F_MW_EQUIVALENCY
'Case F_MEMBERS_SHARING_EQUIVALENCY
'End Select
End Sub

Private Sub cmbLstGroups_Click()
Dim i As Long
On Error Resume Next
CurrLstGroupInd = cmbLstGroups.ListIndex
lstGroups.Clear
lstPairs.Clear
lstPeaksL.Clear
For i = LstGroupsInd1(CurrLstGroupInd) To LstGroupsInd2(CurrLstGroupInd)
    lstGroups.AddItem GrRes.Members(i).Description
Next i
End Sub


Private Sub cmdClose_Click()
Unload Me
End Sub


Private Sub cmdf1Group_Click()
Dim Grouped As Boolean
UpdateStatus "Grouping classes..."
ClearGroupsAndLists
Select Case f1Type
Case F_MEMBERS_SHARING
     Grouped = GroupByMemberSharingUMC()
Case F_MEMBERS_SHARING_LT_MW
     MsgBox "This option is not implemented yet!", vbOKOnly, glFGTU
     Grouped = True
Case F_MEMBERS_SHARING_GT_MW
     MsgBox "This option is not implemented yet!", vbOKOnly, glFGTU
     Grouped = True
End Select

If Grouped Then
   UpdateStatus ""
Else
   UpdateStatus "Error grouping classes!"
End If
End Sub





Private Sub cmdResetClasses_Click()
'---------------------------------------------------------------
'reset classes to what they were before we started this function
'allow user to change mind if by excident
'---------------------------------------------------------------
Dim Res As Long
If NeedToSave Then
   Res = MsgBox("Changes made to the unique mass classes will be lost! Continue?", vbOKCancel, glFGTU)
   If Res = vbOK Then
      UpdateStatus "Reseting..."
      Erase Tmp.UMCs
      Tmp = GelUMC(CallerID)
      Call ClearGroupsAndLists
      Call ResetUMC
      NeedToSave = False                  'indicate that nothing changed
      UpdateStatus ""
   End If
End If
End Sub

Private Sub Form_Activate()
On Error GoTo err_Activate
If bLoading Then
   UpdateStatus "Loading..."
   CallerID = Me.Tag
   Tmp = GelUMC(CallerID)
   tmpPairs = GelP_D_L(CallerID)
        
   lblPeaksCnt.Caption = GelData(CallerID).CSLines + GelData(CallerID).IsoLines
   lblOrigPairsCount.Caption = GelP_D_L(CallerID).PCnt
   lblCurrPairsCount.Caption = tmpPairs.PCnt
      
   Select Case tmpPairs.DltLblType
   Case ptUMCDlt
        lblPairsType.Caption = "UMC Delta"
   Case ptUMCLbl
        lblPairsType.Caption = "UMC Label"
   Case ptUMCDltLbl
        lblPairsType.Caption = "UMC Delta Label"
   End Select
   lblOrigUMCCnt.Caption = GelUMC(CallerID).UMCCnt
   lblOrigUMCRatio.Caption = Format$(lblOrigUMCCnt.Caption / lblPeaksCnt.Caption, "0.00")
   lblCurrUMCCnt.Caption = Tmp.UMCCnt
   lblCurrUMCRatio.Caption = Format$(lblCurrUMCCnt.Caption / lblPeaksCnt.Caption, "0.00")
   
   UpdateStatus "Preparing analysis..."
   If Not ResetUMC() Then GoTo err_Activate
   If Not tmpPairs.SyncWithUMC Then
      UpdateStatus "Pairs and UMCs not synchronized!"
   Else
      UpdateStatus ""
   End If
   
   Set MyPairViewer = New frmPairsView
   MyPairViewer.CallerID = CallerID
   bLoading = False
End If
Exit Sub

err_Activate:
UpdateStatus "Error preparing analysis..."
End Sub

Private Sub Form_Load()
'Me.WindowState = vbMaximized
FillComboBoxes
Me.Move 100, 100
If IsWinLoaded(TrackerCaption) Then frmTracker.Visible = False
DoEvents
bLoading = True
f1MWDist = txtf1MWDist.Text
f1ScanDist = txtf1ScanDist.Text
f1MWTol = txtf1MWTol.Text
MultiGroupsOnly = chkMultiMemberGroups.value
CurrGroupInd = -1
CurrPairInd = -1
CurrPeakIndL = -1
CurrPeakIndH = -1
End Sub

Private Function GetIncCount() As Long
'---------------------------------------------
'included distributions are all with TmpInc>=0
'---------------------------------------------
Dim i As Long
Dim Cnt As Long
On Error Resume Next
For i = 0 To Tmp.UMCCnt - 1
    If TmpInc(i) >= 0 Then Cnt = Cnt + 1
Next i
GetIncCount = Cnt
End Function

Private Sub UpdateStatus(ByVal Msg As String)
lblStatus.Caption = Msg
DoEvents
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Res As Long
If NeedToSave Then
   Res = MsgBox("Do you want to save the changes?", vbYesNoCancel, glFGTU)
   Select Case Res
   Case vbYes                               'if Yes save and unload
      UpdateStatus "Saving ..."
      GelUMC(CallerID) = Tmp
   Case vbCancel                            'Cancel unload
      Cancel = True
   End Select
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
NeedToSave = False
Erase Tmp.UMCs
Tmp.UMCCnt = 0
Unload MyPairViewer
Set MyPairViewer = Nothing
'If IsWinLoaded(TrackerCaption) Then frmTracker.Visible = True
End Sub

Private Sub lstGroups_Click()
Dim CurrLstInd As Long
CurrLstInd = lstGroups.ListIndex
If CurrLstInd >= 0 Then
   CurrGroupInd = CurrLstGroupInd * MAX_GROUPS_IN_LIST + CurrLstInd
   ListPairsForGroup CurrGroupInd
End If
CurrPairInd = -1
CurrPeakIndL = -1
CurrPeakIndH = -1
End Sub

Private Sub lstPairs_Click()
Dim FirstScan As Long, LastScan As Long
Dim MinMW As Double, MaxMW As Double
Dim CurrLstInd As Long
On Error GoTo LstPairsErrorHandler
CurrLstInd = lstPairs.ListIndex
If CurrLstInd >= 0 Then
   CurrPairInd = GrRes.Members(CurrGroupInd).Members(CurrLstInd)
   If CurrPairInd >= 0 Then
     With tmpPairs.Pairs(CurrPairInd)
      ListPeaksForClass .P1, P_LIGHT
      ListPeaksForClass .P2, P_HEAVY
      If MyPairViewer.Visible Then   'zoom to proper region with viewer
         FirstScan = Tmp.UMCs(.P1).MinScan - 2
         LastScan = Tmp.UMCs(.P1).MaxScan + 2
         MinMW = Tmp.UMCs(.P1).ClassMW - 1
         MaxMW = Tmp.UMCs(.P1).ClassMW + 1
         MyPairViewer.Zoom_Light FirstScan, LastScan, MinMW, MaxMW
         FirstScan = Tmp.UMCs(.P2).MinScan - 2
         LastScan = Tmp.UMCs(.P2).MaxScan + 2
         MinMW = Tmp.UMCs(.P2).ClassMW - 1
         MaxMW = Tmp.UMCs(.P2).ClassMW + 1
         MyPairViewer.Zoom_Heavy FirstScan, LastScan, MinMW, MaxMW
      End If
     End With
   End If
End If
CurrPeakIndL = -1
CurrPeakIndH = -1
Exit Sub

LstPairsErrorHandler:
Debug.Print "Error occurred in frmVisPairs->lstPairs_Click: " & Err.Description
Debug.Assert False
LogErrors Err.Number, "frmVisPairs->lstPairs_Click"
Resume Next
End Sub

Private Sub MyPairViewer_pvControlDone()
Me.SetFocus
End Sub

Private Sub MyPairViewer_pvUnload()
chkShowPairs.value = vbUnchecked
End Sub

Private Sub txtf1MWDist_LostFocus()
Dim Tmp As String
Tmp = Trim$(txtf1MWDist.Text)
If IsNumeric(Tmp) Then
   If Tmp > 0 Then
      f1MWDist = CDbl(Tmp)
      Exit Sub
   End If
End If
'at this point something is wrong
MsgBox "This argument should be positive number!", vbOKOnly, glFGTU
txtf1MWDist.SetFocus
End Sub

Private Sub txtf1MWTol_LostFocus()
Dim Tmp As String
Tmp = Trim$(txtf1MWTol.Text)
If IsNumeric(Tmp) Then
   If Tmp > 0 Then
      f1MWTol = CDbl(Tmp)
      Exit Sub
   End If
End If
'at this point something is wrong
MsgBox "This argument should be positive number!", vbOKOnly, glFGTU
txtf1MWTol.SetFocus
End Sub

Private Sub txtf1ScanDist_LostFocus()
Dim Tmp As String
Tmp = Trim$(txtf1ScanDist.Text)
If IsNumeric(Tmp) Then
   If Tmp >= 0 Then
      f1ScanDist = CLng(Tmp)
      Exit Sub
   End If
End If
'at this point something is wrong
MsgBox "This argument should be non-negative integer!", vbOKOnly, glFGTU
txtf1ScanDist.SetFocus
End Sub


Private Sub ResolveResults()
'------------------------------------------------------------------------
'resolve results obtained from grouping functions in a user friendly form
'eliminates groups with less than 2 class-members if required
'------------------------------------------------------------------------
Dim i As Long
Dim NewResCnt As Long
cmbLstGroups.Clear
cmbLstGroups.Enabled = False
With GrRes
    If MultiGroupsOnly Then
       If GetOneClassGroupsCount() > 0 Then
          NewResCnt = 0
          For i = 0 To .Count - 1
              If .Members(i).Count > 1 Then
                 NewResCnt = NewResCnt + 1
                 .Members(NewResCnt - 1) = .Members(i)
              End If
          Next i
          .Count = NewResCnt
          If NewResCnt > 0 Then
             ReDim Preserve .Members(NewResCnt - 1)
          Else
             Erase .Members
          End If
       End If
    End If
    If .Count > 0 Then
       If .Count > MAX_GROUPS_IN_LIST Then
          If .Count Mod MAX_GROUPS_IN_LIST < 10 Then            'don't allow last list
            LstGroupsCnt = Int(.Count / MAX_GROUPS_IN_LIST)     'to be too small
          Else
            LstGroupsCnt = Int(.Count / MAX_GROUPS_IN_LIST) + 1
          End If
          ReDim LstGroupsInd1(LstGroupsCnt - 1)
          ReDim LstGroupsInd2(LstGroupsCnt - 1)
          ReDim LstGroupsDisplay(LstGroupsCnt - 1)
          For i = 0 To LstGroupsCnt - 1
              LstGroupsInd1(i) = i * MAX_GROUPS_IN_LIST
              LstGroupsInd2(i) = (i + 1) * MAX_GROUPS_IN_LIST - 1
          Next i
          LstGroupsInd2(LstGroupsCnt - 1) = .Count - 1          'make sure last index is really last
          'create names for each part of the list and fill the combo box
          For i = 0 To LstGroupsCnt - 1
              LstGroupsDisplay(i) = "Groups " & LstGroupsInd1(i) & " - " & LstGroupsInd2(i)
              cmbLstGroups.AddItem LstGroupsDisplay(i)
          Next i
          cmbLstGroups.Enabled = True
          'present first group of groups
          cmbLstGroups.ListIndex = 0
       Else
          For i = 0 To .Count - 1
              lstGroups.AddItem .Members(i).Description
          Next i
          CurrLstGroupInd = 0           'important for calculation of absolute index
       End If
       lblGroupsCount.Caption = .Count
    Else
       lblGroupsCount.Caption = "0"
    End If
End With
End Sub



Public Sub ListPairsForGroup(ByVal GroupInd As Long)
'----------------------------------------------------------
'fills list box with pairs descriptions for specified group
'----------------------------------------------------------
Dim i As Long
On Error Resume Next
lstPairs.Clear
lstPeaksL.Clear
lstPeaksH.Clear
If GroupInd >= 0 Then
   With GrRes.Members(GroupInd)
       For i = 0 To .Count
           lstPairs.AddItem P_Display(.Members(i))
       Next i
   End With
End If
End Sub


Private Function PrepareClasses() As Boolean
'------------------------------------------------------------
'calculates UMC statistics and prepares classes display names
'NOTE: this function is from VisUMC function; however here
'class display names are shorter since we need to list pairs
'------------------------------------------------------------
'column 0 - class index in .UMCs
'column 1 - class first scan number
'column 2 - class last scan number
Dim i As Long, j As Long
On Error GoTo err_PrepareClasses

With Tmp
   ReDim UMCStat(.UMCCnt - 1, 2)
   ReDim UMCDisplay(.UMCCnt - 1)
   For i = 0 To .UMCCnt - 1
      With .UMCs(i)
          UMCStat(i, 0) = i
          If .ClassCount > 0 Then
             'class members are ordered on scan numbers
             Select Case .ClassMType(0)                 'first scan number
             Case gldtCS
               UMCStat(i, 1) = GelData(CallerID).CSNum(.ClassMInd(0), csfScan)
             Case gldtIS
               UMCStat(i, 1) = GelData(CallerID).IsoNum(.ClassMInd(0), isfScan)
             End Select
             Select Case .ClassMType(.ClassCount - 1)   'last scan number
             Case gldtCS
               UMCStat(i, 2) = GelData(CallerID).CSNum(.ClassMInd(.ClassCount - 1), csfScan)
             Case gldtIS
               UMCStat(i, 2) = GelData(CallerID).IsoNum(.ClassMInd(.ClassCount - 1), isfScan)
             End Select
          Else     'this should not happen
             For j = 0 To 2
                 UMCStat(i, j) = -1
             Next j
          End If
          UMCDisplay(i) = "UMC " & i & "; MW~" & Format$(.ClassMW, "0.00") & "; Cnt " & .ClassCount
      End With
   Next i
End With
PrepareClasses = True
Exit Function

err_PrepareClasses:
End Function



Private Sub ListPeaksForClass(ByVal ClassInd As Long, _
                              ByVal MemberType As Long)
'---------------------------------------------------------------------
'fills list with desc. of class members for light or heavy pair member
'desc. contains type CS or IS index in Num arrays, scan #, charge st.,
'fit, abundance
'---------------------------------------------------------------------
Dim i As Long
Dim DataInd As Long
Dim Desc As String
Select Case MemberType
Case P_LIGHT
     lstPeaksL.Clear
Case P_HEAVY
     lstPeaksH.Clear
End Select
If ClassInd >= 0 Then
  With GelData(CallerID)
    For i = 0 To Tmp.UMCs(ClassInd).ClassCount - 1
      DataInd = Tmp.UMCs(ClassInd).ClassMInd(i)
      Select Case Tmp.UMCs(ClassInd).ClassMType(i)
      Case glCSType
        Desc = "CS " & DataInd & "; Scan " & .CSNum(DataInd, csfScan) & "; " _
               & Format$(.CSNum(DataInd, csfMW), "0.00") & "; " _
               & Format$(.CSNum(DataInd, csfAbu), "Scientific") _
               & "; CS" & .CSNum(DataInd, csfFirstCS) & "; " _
               & "; Fit NA"

      Case glIsoType
        Desc = "IS " & DataInd & "; Scan " & .IsoNum(DataInd, isfScan) & "; " _
               & Format$(.IsoNum(DataInd, Tmp.def.MWField), "0.00") & "; " _
               & Format$(.IsoNum(DataInd, isfAbu), "Scientific") _
               & "; CS " & .IsoNum(DataInd, isfCS) & "; Fit " _
               & Format$(.IsoNum(DataInd, isfFit), "0.00")
      End Select
      Select Case MemberType
      Case P_LIGHT
           lstPeaksL.AddItem Desc
      Case P_HEAVY
           lstPeaksH.AddItem Desc
      End Select
    Next i
  End With
End If
End Sub

Private Function MergeClasses(ByVal UMCInd1 As Long, _
                              ByVal UMCInd2 As Long) As Long
'-----------------------------------------------------------------
'merge elements of classes UMCInd1 and UMCInd2 in one class
'resulting class has lower index of UMCInd1 and UMCInd2 and class
'representative from that class - that way preferences in class
'build is preserved; class with higher index is not removed - its
'index in TmpInc is marked with -1; removal of all marked classes
'is done with function RemoveClasses
'Function returns index of MIP class; -1 on any error
'NOTE: If called with UMCInd1=UMCInd2 function will mark class for
'removal without doing anything else
'-----------------------------------------------------------------
Dim MIPClassInd As Long
Dim LIPClassInd As Long
'peaks that has to be added to More Important Class
Dim PeakCnt As Long
Dim PeakType() As Long
Dim PeakInd() As Long
Dim i As Long
On Error GoTo err_MergeClasses
If UMCInd1 = UMCInd2 Then
   TmpInc(UMCInd1) = REMOVE_UMC_MARK
Else
   If UMCInd1 < UMCInd2 Then
      MIPClassInd = UMCInd1
      LIPClassInd = UMCInd2
   Else
      MIPClassInd = UMCInd2
      LIPClassInd = UMCInd1
   End If
   With GelUMC(CallerID).UMCs(LIPClassInd)
       'redimension to highest number of peaks that could be added
       ReDim PeakType(.ClassCount - 1)
       ReDim PeakInd(.ClassCount - 1)
       PeakCnt = 0
       For i = 0 To .ClassCount - 1
           If Not IsClassMember(.ClassMType(i), .ClassMInd(i), MIPClassInd) Then
              PeakCnt = PeakCnt + 1
              PeakType(PeakCnt - 1) = .ClassMType(i)
              PeakInd(PeakCnt - 1) = .ClassMInd(i)
           End If
       Next i
       If PeakCnt > 0 Then
          If PeakCnt < .ClassCount Then 'redimension if necessary
             ReDim Preserve PeakType(PeakCnt - 1)
             ReDim Preserve PeakInd(PeakCnt - 1)
          End If
          If Not AddPeakArrayToTheClass(PeakType(), PeakInd(), MIPClassInd) Then GoTo err_MergeClasses
       End If
   End With
   If Not RecalculateClass(MIPClassInd) Then GoTo err_MergeClasses
   TmpInc(LIPClassInd) = REMOVE_UMC_MARK
End If
MergeClasses = MIPClassInd
Exit Function

err_MergeClasses:
MergeClasses = -1
End Function



Private Function RemoveClasses() As Long
'---------------------------------------------------------------
'update unique mass classes structure by removing classes marked
'as REMOVE_UMC_MARK classes; returns number of removed classes
'or -1 on any error
'---------------------------------------------------------------
Dim i As Long
Dim Cnt As Long, RemoveCnt As Long
On Error GoTo err_RemoveClasses
With Tmp
  If GetIncCount() <> .UMCCnt Then          'don't go in this process
    For i = 0 To .UMCCnt - 1                'before first verifying there
        If TmpInc(i) < 0 Then               'is something to remove
           Erase .UMCs(i).ClassMInd
           Erase .UMCs(i).ClassMType
           .UMCs(i).ClassCount = 0
           RemoveCnt = RemoveCnt + 1
        Else
           Cnt = Cnt + 1
           .UMCs(Cnt - 1) = .UMCs(i)
        End If
    Next i
    If Cnt > 0 Then
        ReDim Preserve .UMCs(Cnt - 1)
    Else
        ReDim Preserve .UMCs(0)
    End If
    .UMCCnt = Cnt
  End If
End With
RemoveClasses = RemoveCnt
Exit Function

err_RemoveClasses:
Debug.Assert False
RemoveClasses = -1
End Function


Private Function ResetUMC() As Boolean
'-------------------------------------------------------------
'reset all arrays related with unique mass classes and returns
'True if succesful; this function should be called when form
'loads and after each change in unique mass classes structure
'Function also counts UMCs regarding belonging to the pairs
'-------------------------------------------------------------
On Error GoTo err_ResetUMC
Dim dQS As New QSDouble
Dim i As Long

UpdateLabels LBL_ZLS
If Not PrepareClasses() Then GoTo err_ResetUMC
If Not PreparePairs() Then GoTo err_ResetUMC
With Tmp
   ReDim TmpInc(.UMCCnt - 1)
   ReDim UMCInd(.UMCCnt - 1)
   ReDim UMCMW(.UMCCnt - 1)
   ReDim UMCCntL(.UMCCnt - 1)
   ReDim UMCCntH(.UMCCnt - 1)
   ReDim UMCPairs(.UMCCnt - 1)
   For i = 0 To .UMCCnt - 1
       UMCMW(i) = .UMCs(i).ClassMW
       UMCInd(i) = i
   Next i
   If Not dQS.QSAsc(UMCMW, UMCInd) Then GoTo err_ResetUMC
End With
'----------------------------------------------------------
'count all occurances of UMC as light and heavy pair member
'and list all pairs in which UMC participate as a member
'----------------------------------------------------------
With tmpPairs
   For i = 0 To .PCnt - 1
       UMCCntL(.Pairs(i).P1) = UMCCntL(.Pairs(i).P1)
       UMCCntH(.Pairs(i).P2) = UMCCntH(.Pairs(i).P2)
       UMCPairs(.Pairs(i).P1).Count = UMCPairs(.Pairs(i).P1).Count + 1
       ReDim Preserve UMCPairs(.Pairs(i).P1).Members(UMCPairs(.Pairs(i).P1).Count - 1)
       UMCPairs(.Pairs(i).P1).Members(UMCPairs(.Pairs(i).P1).Count - 1) = i
       UMCPairs(.Pairs(i).P2).Count = UMCPairs(.Pairs(i).P2).Count + 1
       ReDim Preserve UMCPairs(.Pairs(i).P2).Members(UMCPairs(.Pairs(i).P2).Count - 1)
       UMCPairs(.Pairs(i).P2).Members(UMCPairs(.Pairs(i).P2).Count - 1) = i
   Next i
End With
UpdateLabels 1
ResetUMC = True
Exit Function

err_ResetUMC:
UpdateLabels LBL_ERR
End Function

Private Function AddPeakToTheClass(ByVal PeakType As Long, _
                                   ByVal PeakInd As Long, _
                                   ByVal ClassInd As Long) As Long
'-------------------------------------------------------------------
'adds specified peak to the class and returns its index in the class
'or -1 in case of any error
'-------------------------------------------------------------------
On Error GoTo err_AddPeakToTheClass
With GelUMC(CallerID).UMCs(ClassInd)
     .ClassCount = .ClassCount + 1
     ReDim Preserve .ClassMInd(.ClassCount - 1)
     ReDim Preserve .ClassMType(.ClassCount - 1)
     .ClassMInd(.ClassCount - 1) = PeakInd
     .ClassMType(.ClassCount - 1) = PeakType
     AddPeakToTheClass = .ClassCount - 1
End With
Exit Function

err_AddPeakToTheClass:
AddPeakToTheClass = -1
End Function


Private Function AddPeakArrayToTheClass(NewPeakType() As Long, _
                                        NewPeakInd() As Long, _
                                        ByVal ClassInd As Long) As Boolean
'-------------------------------------------------------------------------
'adds specified peak array to the class and returns True if successful
'NOTE: has to be careful here since class members have to be ordered on
'      scan numbers
'-------------------------------------------------------------------------
Dim NewPeaksCnt As Long
Dim TTlPeaksCnt As Long
Dim TTlPeakInd() As Long
Dim TTLPeakType() As Long
Dim TTlFN() As Long             'scan number
Dim TTlOrd() As Long            'ordering array
Dim QSL As New QSLong           'ordering object
Dim i As Long, j As Long
On Error GoTo err_AddPeakArrayToTheClass
NewPeaksCnt = UBound(NewPeakInd) + 1
If NewPeaksCnt > 0 Then         'not error if nothing to add
   With GelUMC(CallerID).UMCs(ClassInd)
       'we have to put new and old together and order based on scan numbers
       TTlPeaksCnt = .ClassCount + NewPeaksCnt
       ReDim TTlPeakInd(TTlPeaksCnt - 1)
       ReDim TTLPeakType(TTlPeaksCnt - 1)
       ReDim TTlOrd(TTlPeaksCnt - 1)
       ReDim TTlFN(TTlPeaksCnt - 1)
       For i = 0 To .ClassCount - 1             'first old peaks
           TTlPeakInd(i) = .ClassMInd(i)
           TTLPeakType(i) = .ClassMType(i)
           TTlOrd(i) = i
           Select Case .ClassMType(i)
           Case glCSType
                TTlFN(i) = GelData(CallerID).CSNum(.ClassMInd(i), csfScan)
           Case glIsoType
                TTlFN(i) = GelData(CallerID).IsoNum(.ClassMInd(i), isfScan)
           End Select
       Next i
       For i = 0 To NewPeaksCnt - 1             'then new peaks
           j = .ClassCount + i
           TTlOrd(j) = j
           TTlPeakInd(j) = NewPeakInd(i)
           TTLPeakType(j) = NewPeakType(i)
           Select Case TTLPeakType(j)
           Case glCSType
                TTlFN(j) = GelData(CallerID).CSNum(TTlPeakInd(j), csfScan)
           Case glIsoType
                TTlFN(j) = GelData(CallerID).IsoNum(TTlPeakInd(j), isfScan)
           End Select
       Next i
       'now order them on scan numbers ascending
       If Not QSL.QSAsc(TTlFN(), TTlOrd()) Then GoTo err_AddPeakArrayToTheClass
       Set QSL = Nothing
       ReDim .ClassMInd(TTlPeaksCnt - 1)
       ReDim .ClassMType(TTlPeaksCnt - 1)
       For i = 0 To TTlPeaksCnt - 1
           .ClassMInd(i) = TTlPeakInd(TTlOrd(i))
           .ClassMType(i) = TTLPeakType(TTlOrd(i))
       Next i
       .ClassCount = TTlPeaksCnt
   End With
End If
AddPeakArrayToTheClass = True
Exit Function

err_AddPeakArrayToTheClass:
End Function



Private Function RemovePeakFromTheClass(ByVal PeakClassInd As Long, _
                                        ByVal ClassInd As Long) As Boolean
'----------------------------------------------------------------------------
'removes peak with class index from the class and return True if OK.
'If class becomes empty by removing this peak or peak is class representative
'class is marked for removal
'NOTE: Index here is peak index in the class
'----------------------------------------------------------------------------
Dim i As Long
On Error GoTo err_RemovePeakFromTheClass

With Tmp.UMCs(ClassInd)
    If .ClassCount > 1 Then
       For i = PeakClassInd To .ClassCount - 2
           .ClassMType(i) = .ClassMType(i + 1)
           .ClassMInd(i) = .ClassMInd(i + 1)
       Next i
       .ClassCount = .ClassCount - 1
       ReDim Preserve .ClassMInd(.ClassCount - 1)
       ReDim Preserve .ClassMType(.ClassCount - 1)
       'call recalculate class MW/Abu
    Else
       .ClassCount = 0
       Erase .ClassMType
       Erase .ClassMInd
       TmpInc(ClassInd) = REMOVE_UMC_MARK
    End If
End With

RemovePeakFromTheClass = True
Exit Function

err_RemovePeakFromTheClass:
End Function

Private Function RecalculateClass(ByVal ClassInd) As Boolean
'-----------------------------------------------------------
'recalculates class MW and abundance based on current class
'membership and definition
'-----------------------------------------------------------
Dim MWSum As Double
Dim AbuSum As Double
Dim i As Long
On Error GoTo err_RecalculateClass

'no need to waste time on excluded classes
If TmpInc(ClassInd) < 0 Then GoTo err_RecalculateClass
With Tmp.UMCs(ClassInd)
    'no need to recalculate for class representative since it can not change
    If (Tmp.def.ClassMW <> UMCClassMassConstants.UMCMassRep) Or (Tmp.def.ClassAbu <> UMCClassAbundanceConstants.UMCAbuRep) Then
        For i = 0 To .ClassCount - 1
            Select Case .ClassMType(i)
            Case glCSType
                 MWSum = MWSum + GelData(CallerID).CSNum(.ClassMInd(i), csfMW)
                 AbuSum = AbuSum + GelData(CallerID).CSNum(.ClassMInd(i), csfAbu)
            Case glIsoType
                 MWSum = MWSum + GelData(CallerID).IsoNum(.ClassMInd(i), Tmp.def.MWField)
                 AbuSum = AbuSum + GelData(CallerID).IsoNum(.ClassMInd(i), isfAbu)
            End Select
        Next i
        If Tmp.def.ClassMW = UMCClassMassConstants.UMCMassAvg Then .ClassMW = MWSum / .ClassCount
        Select Case Tmp.def.ClassAbu
        Case UMCClassAbundanceConstants.UMCAbuRep      'do nothing
        Case UMCClassAbundanceConstants.UMCAbuSum
             .ClassAbundance = AbuSum
        Case UMCClassAbundanceConstants.UMCAbuAvg
             .ClassAbundance = AbuSum / .ClassCount
        End Select
    End If
End With
RecalculateClass = True
Exit Function

err_RecalculateClass:
End Function

Private Function GetOneClassGroupsCount() As Long
'------------------------------------------------
'returns number of groups with only one class
'------------------------------------------------
Dim i As Long
Dim Cnt As Long
On Error Resume Next
With GrRes
     For i = 0 To .Count - 1
         If .Members(i).Count <= 1 Then Cnt = Cnt + 1
     Next i
End With
GetOneClassGroupsCount = Cnt
End Function

Private Function IsClassMember(ByVal PeakType As Long, _
                               ByVal PeakInd As Long, _
                               ByVal ClassInd As Long) As Boolean
'---------------------------------------------------------------
'returns True if data point with index PeakInd of PeakType is
'member of class ClassInd
'---------------------------------------------------------------
Dim i As Long
On Error GoTo exit_IsClassMember
With GelUMC(CallerID).UMCs(ClassInd)
    For i = 0 To .ClassCount - 1
        If .ClassMInd(i) = PeakInd Then
           If .ClassMType(i) = PeakType Then
              IsClassMember = True
              GoTo exit_IsClassMember
           End If
        End If
    Next i
End With

exit_IsClassMember:
End Function

Private Sub ClearGroupsAndLists()
'--------------------------------------------
'clears lists with groups, classes and peaks
'--------------------------------------------
Call DestroyGroups
lstGroups.Clear
lstPairs.Clear
lstPeaksL.Clear
lstPeaksH.Clear
CurrGroupInd = -1
CurrPairInd = -1
CurrPeakIndL = -1
CurrPeakIndH = -1
cmbLstGroups.Clear
End Sub


Private Sub UpdateLabels(ByVal lblType As Long)
'----------------------------------------------
'updates labels displaying current counts
'WILL HAVE TO BE UPDATED
'----------------------------------------------
On Error Resume Next
Select Case lblType
Case LBL_ERR
    lblGroupsCount.Caption = "Error"
    lblCurrUMCCnt.Caption = "Error"
    lblCurrUMCRatio.Caption = "Error"
Case LBL_ZLS
    lblGroupsCount.Caption = ""
    lblCurrUMCCnt.Caption = ""
    lblCurrUMCRatio.Caption = ""
Case Else
    lblGroupsCount.Caption = GrRes.Count
    lblCurrUMCCnt.Caption = Tmp.UMCCnt
    lblCurrUMCRatio.Caption = Format$(lblCurrUMCCnt.Caption / lblPeaksCnt.Caption, "0.00")
End Select
End Sub

Private Sub DestroyGroups()
'------------------------------------------------
'destroys groups structure
'------------------------------------------------
Dim i As Long
On Error Resume Next
With GrRes
     For i = 0 To .Count - 1
        Erase .Members(i).Members
        .Members(i).Count = 0
        .Members(i).Description = ""
     Next i
     Erase .Members
     .Count = 0
     .Description = ""
End With
End Sub


Private Function ScanCloseClasses(ByVal ClassInd1 As Long, _
                                  ByVal ClassInd2 As Long) As Boolean
'--------------------------------------------------------------------
'returns True if classes ClassInd1 and ClassInd2 are close in regard
'of scan numbers
'NOTE: close here means that distance between closest(in scan regard)
'points of two classes is not more than f1ScanDist
'NOTE: to understand logic of this function draw all possible cases
'      of arangement of two segments in a 1D space
'--------------------------------------------------------------------
Dim ClosestScanDist As Long
If (UMCStat(ClassInd1, 1) < UMCStat(ClassInd2, 1)) And (UMCStat(ClassInd1, 2) < UMCStat(ClassInd2, 1)) Then
    ClosestScanDist = UMCStat(ClassInd2, 1) - UMCStat(ClassInd1, 2)
ElseIf (UMCStat(ClassInd2, 1) < UMCStat(ClassInd1, 1)) And (UMCStat(ClassInd2, 2) < UMCStat(ClassInd1, 1)) Then
    ClosestScanDist = UMCStat(ClassInd1, 1) - UMCStat(ClassInd2, 2)
Else
    ClosestScanDist = 0
End If
ScanCloseClasses = (ClosestScanDist <= f1ScanDist)
End Function


Private Function PreparePairs() As Boolean
'------------------------------------------
'prepares display names for pairs and index
'pairs on mass of light pair member
'------------------------------------------
Dim i As Long

On Error GoTo err_PreparePairs

With tmpPairs
   ReDim P_Ind(.PCnt - 1)
   ReDim P_MW(.PCnt - 1)
   ReDim P_Display(.PCnt - 1)
   For i = 0 To .PCnt - 1
       P_Ind(i) = i
       P_MW(i) = Tmp.UMCs(.Pairs(i).P1).ClassMW
       P_Display(i) = UMCDisplay(.Pairs(i).P1) & "  *  " & UMCDisplay(.Pairs(i).P2)
   Next i
End With
PreparePairs = True
Exit Function

err_PreparePairs:
End Function

Private Function GroupByMemberSharingUMC() As Boolean
'------------------------------------------------------------------------------
'in this case we have everything prepared since pairs that share same UMC class
'will go to the same group and will be listed only if more than one
'NOTE: here we can expect significant overlap among groups
'      list is ordered on UMCs molecular masses
'------------------------------------------------------------------------------
Dim i As Long, j As Long
On Error GoTo err_GroupByMemberSharingUMC
With GrRes
    For i = 0 To UBound(UMCPairs)
        If UMCPairs(UMCInd(i)).Count > 1 Then
           .Count = .Count + 1
           ReDim Preserve .Members(.Count - 1)
           .Members(.Count - 1).Description = "Pairs sharing " & UMCDisplay(UMCInd(i))
           With .Members(.Count - 1)
                ReDim .Members(UMCPairs(UMCInd(i)).Count - 1)
                For j = 0 To UMCPairs(UMCInd(i)).Count - 1
                    .Count = .Count + 1
                    .Members(j) = UMCPairs(UMCInd(i)).Members(j)
                Next j
           End With
        End If
    Next i
End With

Call ResolveResults
GroupByMemberSharingUMC = True
Exit Function

err_GroupByMemberSharingUMC:
End Function






