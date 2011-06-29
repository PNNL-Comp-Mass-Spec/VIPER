VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data Filter"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   HelpContextID   =   113
   Icon            =   "frmFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdUseDefaultFilters 
      Caption         =   "Use &Defaults"
      Height          =   495
      Left            =   2280
      TabIndex        =   74
      ToolTipText     =   "Apply default filters (do not close the form)"
      Top             =   6000
      Width           =   855
   End
   Begin TabDlg.SSTab tbsFilters 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   9551
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tolerances"
      TabPicture(0)   =   "frmFilter.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraExcludeDuplicatesOrFit"
      Tab(0).Control(1)=   "fraTolerances"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Identity and Comparative Display"
      TabPicture(1)   =   "frmFilter.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraIdentityAndComparativeDisplay"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Mass Range and Scan #"
      TabPicture(2)   =   "frmFilter.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraMassRange"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Charge and Abundance"
      TabPicture(3)   =   "frmFilter.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraChargeAndAbundance"
      Tab(3).ControlCount=   1
      Begin VB.Frame fraExcludeDuplicatesOrFit 
         Caption         =   "Duplicates, Fit, and St. Dev."
         Height          =   2655
         Left            =   -74760
         TabIndex        =   2
         Top             =   840
         Width           =   4935
         Begin VB.TextBox txtStDevTolerance 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3960
            TabIndex        =   11
            Top             =   2040
            Width           =   735
         End
         Begin VB.CheckBox chkBadStDevElimination 
            Caption         =   "Ex&clude data with St.Dev worse than: (Charge State data only)"
            Height          =   375
            Left            =   240
            TabIndex        =   10
            Top             =   2040
            Width           =   3495
         End
         Begin VB.TextBox txtFitTolerance 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3960
            TabIndex        =   9
            Top             =   1455
            Width           =   735
         End
         Begin VB.CheckBox chkBadFitElimination 
            Caption         =   "Ex&clude data with calculated fit worse than: (Isotopic data only)"
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   1440
            Width           =   3495
         End
         Begin VB.TextBox txtDBFitTolerance 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3960
            TabIndex        =   7
            Top             =   915
            Width           =   735
         End
         Begin VB.CheckBox chkDBBadFitElimination 
            Caption         =   "E&xclude data with database match worse than:"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   960
            Width           =   3735
         End
         Begin VB.TextBox txtDupTolerance 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3960
            TabIndex        =   5
            Top             =   405
            Width           =   735
         End
         Begin VB.CheckBox chkDupElimination 
            Caption         =   "&Exclude duplicates from the Isotopic data"
            Height          =   375
            Left            =   240
            TabIndex        =   3
            ToolTipText     =   "Isotopic Data Only"
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label Label1 
            Caption         =   "Duplicate Tolerance"
            Height          =   375
            Left            =   2880
            TabIndex        =   4
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame fraChargeAndAbundance 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   -74760
         TabIndex        =   50
         Top             =   720
         Width           =   4695
         Begin VB.Frame fraAbuRange 
            Caption         =   "Abundance Range"
            Height          =   2055
            Left            =   0
            TabIndex        =   57
            Top             =   1320
            Width           =   4575
            Begin VB.CheckBox chkCSIsoSameRange 
               Caption         =   "Same range for Charge State and Isotopic data"
               Height          =   255
               Left            =   480
               TabIndex        =   68
               Top             =   1680
               Value           =   1  'Checked
               Width           =   3975
            End
            Begin VB.TextBox txtIsoMaxAbu 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3240
               TabIndex        =   65
               Top             =   960
               Width           =   1095
            End
            Begin VB.TextBox txtIsoMinAbu 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3240
               TabIndex        =   67
               Top             =   1320
               Width           =   1095
            End
            Begin VB.CheckBox chkIsoAbuRange 
               Caption         =   "Use only Isotopic data with abundance within range  "
               Height          =   615
               Left            =   240
               TabIndex        =   63
               Top             =   960
               Width           =   2175
            End
            Begin VB.TextBox txtCSMaxAbu 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3240
               TabIndex        =   60
               Top             =   240
               Width           =   1095
            End
            Begin VB.TextBox txtCSMinAbu 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3240
               TabIndex        =   62
               Top             =   600
               Width           =   1095
            End
            Begin VB.CheckBox chkCSAbuRange 
               Caption         =   "Use only Charge State data with abundance within range  "
               Height          =   615
               Left            =   240
               TabIndex        =   58
               Top             =   240
               Width           =   2175
            End
            Begin VB.Label Label2 
               Caption         =   "Max"
               Height          =   255
               Index           =   2
               Left            =   2520
               TabIndex        =   64
               Top             =   1005
               Width           =   495
            End
            Begin VB.Label Label2 
               Caption         =   "Min"
               Height          =   255
               Index           =   3
               Left            =   2520
               TabIndex        =   66
               Top             =   1365
               Width           =   495
            End
            Begin VB.Label Label2 
               Caption         =   "Max"
               Height          =   255
               Index           =   0
               Left            =   2520
               TabIndex        =   59
               Top             =   285
               Width           =   495
            End
            Begin VB.Label Label2 
               Caption         =   "Min"
               Height          =   255
               Index           =   1
               Left            =   2520
               TabIndex        =   61
               Top             =   645
               Width           =   495
            End
         End
         Begin VB.Frame fraCSFlt 
            Caption         =   "Charge State Filter (Isotopic Data Only)"
            Height          =   975
            Left            =   0
            TabIndex        =   51
            Top             =   120
            Width           =   4575
            Begin VB.CheckBox chkIsoUseCSRange 
               Caption         =   "Use only charge states within range"
               Height          =   375
               Left            =   240
               TabIndex        =   52
               Top             =   360
               Width           =   2295
            End
            Begin VB.TextBox txtIsoMinCS 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3480
               TabIndex        =   54
               Text            =   "0"
               Top             =   240
               Width           =   735
            End
            Begin VB.TextBox txtIsoMaxCS 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3480
               TabIndex        =   56
               Text            =   "0"
               Top             =   540
               Width           =   735
            End
            Begin VB.Label Label4 
               Caption         =   "First C.S."
               Height          =   255
               Index           =   0
               Left            =   2640
               TabIndex        =   53
               Top             =   300
               Width           =   735
            End
            Begin VB.Label Label4 
               Caption         =   "Last C.S."
               Height          =   255
               Index           =   1
               Left            =   2640
               TabIndex        =   55
               Top             =   600
               Width           =   735
            End
         End
      End
      Begin VB.Frame fraMassRange 
         BorderStyle     =   0  'None
         Height          =   4455
         Left            =   -74760
         TabIndex        =   28
         Top             =   720
         Width           =   4695
         Begin VB.Frame fraEvenOddScanNumber 
            Caption         =   "Even/Odd Scan Number Filtering"
            Height          =   855
            Left            =   0
            TabIndex        =   47
            Top             =   3600
            Width           =   4575
            Begin VB.ComboBox cboEvenOddScanNumber 
               Height          =   315
               Left            =   2400
               Style           =   2  'Dropdown List
               TabIndex        =   49
               Top             =   360
               Width           =   2055
            End
            Begin VB.Label lblEvenOddScanNumber 
               Caption         =   "Use this filter for DREAMS-based data files."
               Height          =   375
               Left            =   240
               TabIndex        =   48
               Top             =   360
               Width           =   2175
            End
         End
         Begin VB.Frame fraMZRange 
            Caption         =   "M/Z Range (isotopic data only)"
            Height          =   1215
            Left            =   0
            TabIndex        =   41
            Top             =   2280
            Width           =   4575
            Begin VB.CheckBox chkIsoMZRange 
               Caption         =   "Use only Isotopic data with m/z within range"
               Height          =   495
               Left            =   240
               TabIndex        =   42
               Top             =   405
               Width           =   2175
            End
            Begin VB.TextBox txtIsoMinMZ 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3240
               TabIndex        =   46
               Top             =   720
               Width           =   1095
            End
            Begin VB.TextBox txtIsoMaxMZ 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3240
               TabIndex        =   44
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label2 
               Caption         =   "Min"
               Height          =   255
               Index           =   10
               Left            =   2520
               TabIndex        =   45
               Top             =   765
               Width           =   495
            End
            Begin VB.Label Label2 
               Caption         =   "Max"
               Height          =   255
               Index           =   11
               Left            =   2520
               TabIndex        =   43
               Top             =   405
               Width           =   495
            End
         End
         Begin VB.Frame fraMWRange 
            Caption         =   "Molecular Mass Range (monoisotopic mass)"
            Height          =   2055
            Left            =   0
            TabIndex        =   29
            Top             =   120
            Width           =   4575
            Begin VB.CheckBox chkCSIsoSameRangeMW 
               Caption         =   "Same range for Charge State and Isotopic data"
               Height          =   255
               Left            =   480
               TabIndex        =   40
               Top             =   1680
               Value           =   1  'Checked
               Width           =   3975
            End
            Begin VB.TextBox txtIsoMaxMW 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3240
               TabIndex        =   37
               Top             =   960
               Width           =   1095
            End
            Begin VB.TextBox txtIsoMinMW 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3240
               TabIndex        =   39
               Top             =   1320
               Width           =   1095
            End
            Begin VB.CheckBox chkIsoMWRange 
               Caption         =   "Use only Isotopic data with molecular mass within range"
               Height          =   615
               Left            =   240
               TabIndex        =   35
               Top             =   960
               Width           =   2175
            End
            Begin VB.TextBox txtCSMaxMW 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3240
               TabIndex        =   32
               Top             =   240
               Width           =   1095
            End
            Begin VB.TextBox txtCSMinMW 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3240
               TabIndex        =   34
               Top             =   600
               Width           =   1095
            End
            Begin VB.CheckBox chkCSMWRange 
               Caption         =   "Use only Charge State data with molecular mass within range"
               Height          =   615
               Left            =   240
               TabIndex        =   30
               Top             =   240
               Width           =   2175
            End
            Begin VB.Label Label2 
               Caption         =   "Max"
               Height          =   255
               Index           =   4
               Left            =   2520
               TabIndex        =   36
               Top             =   1005
               Width           =   495
            End
            Begin VB.Label Label2 
               Caption         =   "Min"
               Height          =   255
               Index           =   5
               Left            =   2520
               TabIndex        =   38
               Top             =   1365
               Width           =   495
            End
            Begin VB.Label Label2 
               Caption         =   "Max"
               Height          =   255
               Index           =   6
               Left            =   2520
               TabIndex        =   31
               Top             =   285
               Width           =   495
            End
            Begin VB.Label Label2 
               Caption         =   "Min"
               Height          =   255
               Index           =   7
               Left            =   2520
               TabIndex        =   33
               Top             =   645
               Width           =   495
            End
         End
      End
      Begin VB.Frame fraIdentityAndComparativeDisplay 
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   4815
         Begin VB.Frame frIdentity 
            Caption         =   "&Identity"
            Height          =   615
            Left            =   0
            TabIndex        =   16
            Top             =   120
            Width           =   4695
            Begin VB.CheckBox chkIdentity 
               Caption         =   "Exclude ide&ntified data"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   17
               Top             =   240
               Width           =   2055
            End
            Begin VB.CheckBox chkIdentity 
               Caption         =   "Exclude &unidentified data"
               Height          =   255
               Index           =   1
               Left            =   2400
               TabIndex        =   18
               Top             =   240
               Width           =   2175
            End
         End
         Begin VB.Frame frDiffDisplay 
            Caption         =   "&Comparative Display"
            Height          =   1900
            Left            =   0
            TabIndex        =   19
            Top             =   840
            Width           =   4695
            Begin VB.TextBox txtERMin 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3040
               TabIndex        =   25
               Top             =   1400
               Width           =   615
            End
            Begin VB.CheckBox chkDiffDisplay 
               Caption         =   "E&xclude data with ER out of range"
               Height          =   255
               Index           =   4
               Left            =   240
               TabIndex        =   24
               Top             =   1420
               Width           =   2775
            End
            Begin VB.CheckBox chkDiffDisplay 
               Caption         =   "Exclude data &with Expression Ratio (ER)"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   20
               Top             =   260
               Width           =   3375
            End
            Begin VB.CheckBox chkDiffDisplay 
               Caption         =   "Exclude d&ata showing huge overexpression"
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   23
               Top             =   1120
               Width           =   3735
            End
            Begin VB.CheckBox chkDiffDisplay 
               Caption         =   "Exclude &data showing huge underexpression"
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   22
               Top             =   840
               Width           =   3735
            End
            Begin VB.CheckBox chkDiffDisplay 
               Caption         =   "Exclude data with&out ER"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   21
               Top             =   540
               Width           =   2295
            End
            Begin VB.TextBox txtERMax 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3900
               TabIndex        =   27
               Top             =   1400
               Width           =   615
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "-"
               Height          =   195
               Left            =   3740
               TabIndex        =   26
               Top             =   1440
               Width           =   45
            End
         End
      End
      Begin VB.Frame fraTolerances 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   -74760
         TabIndex        =   1
         Top             =   840
         Width           =   4935
         Begin VB.Frame frSecondGuess 
            Caption         =   "Second Guess (Isotopic data only)"
            Height          =   975
            Left            =   0
            TabIndex        =   12
            Top             =   2760
            Width           =   4455
            Begin VB.CheckBox chkSecGuessElimination 
               Caption         =   "Exclude &second guess"
               Height          =   255
               Index           =   0
               Left            =   360
               TabIndex        =   13
               ToolTipText     =   "Check this option to keep data with better fit from calculation"
               Top             =   280
               Width           =   3735
            End
            Begin VB.CheckBox chkSecGuessElimination 
               Caption         =   "Exclude less likely guess"
               Height          =   255
               Index           =   1
               Left            =   360
               TabIndex        =   14
               ToolTipText     =   "Check this option to keep data more likely (in comparisson with other data)"
               Top             =   600
               Width           =   3735
            End
         End
      End
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset to None"
      Height          =   495
      Left            =   240
      TabIndex        =   70
      ToolTipText     =   "Roll back changes without closing the form"
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmdApplyNow 
      Caption         =   "Apply &Now"
      Height          =   495
      Left            =   1320
      TabIndex        =   71
      ToolTipText     =   "Apply filters but do not close the form"
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   73
      ToolTipText     =   "Reject all changes and close the form"
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   72
      ToolTipText     =   "Apply filters and close the form"
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label lblFilterStatus 
      Caption         =   "0 data points visible (0 total points)"
      Height          =   255
      Left            =   240
      TabIndex        =   69
      Top             =   5640
      Width           =   5055
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is filter interface function for one 2D display
'NOTE: This form will unload if looses the focus by event caused outside
'it - the effect is the same as if user pressed the Cancel button
'-----------------------------------------------------------------------
'Last modified: 02/12/2003 nt
'-----------------------------------------------------------------------
Option Explicit

Private Enum eosEvenOddScanConstants
    eosAllScans = 0
    eosOddScans = 1
    eosEvenScans = 2
End Enum

Private Enum erfERFilterConstants
    erfExcludeWithExpression = 0
    erfExcludeWithoutExpression = 1
    erfExcludeHugeUnderExpression = 2
    erfExcludeHugeOverExpression = 3
    erfExcludeOutOfRange = 4
End Enum
    
Private mUpdatingSettings As Boolean    'used to control unwanted behaviour of CheckBox control
                                        'changing Value property triggers Click event

Private mWindowStayOnTopEnabled As Boolean

Private CallerID As Long
                            
Dim OldSettings(1 To 20, 2) As Variant

Private Sub ApplyFilterAndClose()
    If FilterChanged() Then GelStatus(CallerID).Dirty = True
    If FilterThis(True) Then
        GelBody(CallerID).picGraph.Refresh
        Unload Me
    End If
End Sub

Private Sub ApplyDefaultFilters()
    ' Reset to the default filters
    ResetExpandedPreferences glbPreferencesExpanded, "AutoAnalysisFilterPrefs", True
    
    ' Update the filters
    ApplyAutoAnalysisFilter glbPreferencesExpanded.AutoAnalysisFilterPrefs, CallerID, False
    
    ' Update the controls on this form
    DisplayCurrentSettings False
    
    ' Apply the filters
    ApplyFilterNow True
End Sub

Private Sub ApplyFilterNow(blnQueryIfNoVisiblePoints As Boolean)
    FilterThis blnQueryIfNoVisiblePoints
    GelBody(CallerID).picGraph.Refresh
End Sub

Private Sub CancelChanges()
    If FilterChanged() Then   'if something was changed
       RestoreOldSettings   'then restore old settings
       ApplyFilterNow False    'and refilter back changes
    End If
    Unload Me
End Sub

Private Sub DisplayCurrentSettings(blnQueryIfNoVisiblePoints As Boolean)
Dim I As Long
Dim tmp As Long
mUpdatingSettings = True
With GelData(CallerID)
     If .DataFilter(fltDupTolerance, 1) >= 0 Then
        txtDupTolerance.Text = .DataFilter(fltDupTolerance, 1)
     Else
        txtDupTolerance.Text = ""
     End If
     If .DataFilter(fltDBTolerance, 1) >= 0 Then
        txtDBFitTolerance.Text = .DataFilter(fltDBTolerance, 1)
     Else
        txtDBFitTolerance.Text = ""
     End If
     txtFitTolerance.Text = .DataFilter(fltIsoFit, 1)
     txtStDevTolerance.Text = .DataFilter(fltCSStDev, 1)
     
     ' Abundance
     txtCSMinAbu.Text = Format(.DataFilter(fltCSAbu, 1), "Scientific")
     txtCSMaxAbu.Text = Format(.DataFilter(fltCSAbu, 2), "Scientific")
     txtIsoMinAbu.Text = Format(.DataFilter(fltIsoAbu, 1), "Scientific")
     txtIsoMaxAbu.Text = Format(.DataFilter(fltIsoAbu, 2), "Scientific")
     
     ' Monosisotopic mass
     txtCSMinMW.Text = Format(.DataFilter(fltCSMW, 1), "0.00")
     txtCSMaxMW.Text = Format(.DataFilter(fltCSMW, 2), "0.00")
     txtIsoMinMW.Text = Format(.DataFilter(fltIsoMW, 1), "0.00")
     txtIsoMaxMW.Text = Format(.DataFilter(fltIsoMW, 2), "0.00")
     
     ' M/Z
     txtIsoMinMZ.Text = Format(.DataFilter(fltIsoMZ, 1), "0.00")
     txtIsoMaxMZ.Text = Format(.DataFilter(fltIsoMZ, 2), "0.00")
     
     ' Charge
     txtIsoMinCS.Text = .DataFilter(fltIsoCS, 1)
     txtIsoMaxCS.Text = .DataFilter(fltIsoCS, 2)
     
     If CBool(.DataFilter(fltDupTolerance, 0)) Then
        chkDupElimination.Value = vbChecked
        txtDupTolerance.Enabled = True
     Else
        chkDupElimination.Value = vbUnchecked
        txtDupTolerance.Enabled = False
     End If
     If CBool(.DataFilter(fltDBTolerance, 0)) Then
        chkDBBadFitElimination.Value = vbChecked
        txtDBFitTolerance.Enabled = True
     Else
        chkDBBadFitElimination.Value = vbUnchecked
        txtDBFitTolerance.Enabled = False
     End If
     If CBool(.DataFilter(fltIsoFit, 0)) Then
        chkBadFitElimination.Value = vbChecked
        txtFitTolerance.Enabled = True
     Else
        chkBadFitElimination.Value = vbUnchecked
        txtFitTolerance.Enabled = False
     End If
     If CBool(.DataFilter(fltCSStDev, 0)) Then
        chkBadStDevElimination.Value = vbChecked
        txtStDevTolerance.Enabled = True
     Else
        chkBadStDevElimination.Value = vbUnchecked
        txtStDevTolerance.Enabled = False
     End If
     
    SetCheckBox chkCSAbuRange, CBool(.DataFilter(fltCSAbu, 0))        'CS data ranges - abundance
    txtCSMinAbu.Enabled = CBool(.DataFilter(fltCSAbu, 0))
    txtCSMaxAbu.Enabled = CBool(.DataFilter(fltCSAbu, 0))
     
    SetCheckBox chkIsoAbuRange, CBool(.DataFilter(fltIsoAbu, 0))      'Isotopic data ranges - abundance
    txtIsoMinAbu.Enabled = CBool(.DataFilter(fltIsoAbu, 0))
    txtIsoMaxAbu.Enabled = CBool(.DataFilter(fltIsoAbu, 0))
     
    If ((chkCSAbuRange = chkIsoAbuRange) And (txtCSMinAbu.Text = txtIsoMinAbu.Text) And _
        (txtCSMaxAbu.Text = txtIsoMaxAbu.Text)) Then
        chkCSIsoSameRange.Value = vbChecked
    Else
        chkCSIsoSameRange.Value = vbUnchecked
    End If
     
    SetCheckBox chkCSMWRange, CBool(.DataFilter(fltCSMW, 0))        'CS data ranges - molecular mass values
    txtCSMinMW.Enabled = CBool(.DataFilter(fltCSMW, 0))
    txtCSMaxMW.Enabled = CBool(.DataFilter(fltCSMW, 0))
     
    SetCheckBox chkIsoMWRange, CBool(.DataFilter(fltIsoMW, 0))      'Isotopic data ranges - molecular mass values
    txtIsoMinMW.Enabled = CBool(.DataFilter(fltIsoMW, 0))
    txtIsoMaxMW.Enabled = CBool(.DataFilter(fltIsoMW, 0))
     
    If ((chkCSMWRange = chkIsoMWRange) And (txtCSMinMW.Text = txtIsoMinMW.Text) And _
        (txtCSMaxMW.Text = txtIsoMaxMW.Text)) Then
        chkCSIsoSameRangeMW.Value = vbChecked
    Else
        chkCSIsoSameRangeMW.Value = vbUnchecked
    End If
     
    SetCheckBox chkIsoMZRange, CBool(.DataFilter(fltIsoMZ, 0))      'Isotopic data ranges - m/z values
    txtIsoMinMZ.Enabled = CBool(.DataFilter(fltIsoMZ, 0))
    txtIsoMaxMZ.Enabled = CBool(.DataFilter(fltIsoMZ, 0))
     
    SetCheckBox chkIsoUseCSRange, CBool(.DataFilter(fltIsoCS, 0))        'Charge ranges for Isotopic data
    txtIsoMinCS.Enabled = CBool(.DataFilter(fltIsoCS, 0))
    txtIsoMaxCS.Enabled = CBool(.DataFilter(fltIsoCS, 0))
     
     If CBool(.DataFilter(fltCase2CloseResults, 0)) Then
        Select Case CInt(.DataFilter(fltCase2CloseResults, 1))
        Case 1
          chkSecGuessElimination(0).Value = vbChecked
          chkSecGuessElimination(1).Value = vbUnchecked
        Case 2
          chkSecGuessElimination(0).Value = vbUnchecked
          chkSecGuessElimination(1).Value = vbChecked
        Case Else
          chkSecGuessElimination(0).Value = vbUnchecked
          chkSecGuessElimination(1).Value = vbUnchecked
          .DataFilter(fltCase2CloseResults, 1) = 0
        End Select
     Else
        chkSecGuessElimination(0).Value = vbUnchecked
        chkSecGuessElimination(1).Value = vbUnchecked
        .DataFilter(fltCase2CloseResults, 1) = 0
     End If
     Select Case CInt(.DataFilter(fltID, 1))
     Case 0
        chkIdentity(0).Value = vbUnchecked
        chkIdentity(1).Value = vbUnchecked
     Case 1
        chkIdentity(0).Value = vbChecked
        chkIdentity(1).Value = vbUnchecked
     Case 2
        chkIdentity(0).Value = vbUnchecked
        chkIdentity(1).Value = vbChecked
     Case Else
        chkIdentity(0).Value = vbChecked
        chkIdentity(1).Value = vbChecked
     End Select
     If .DataFilter(fltAR, 0) Then
        tmp = Abs(.DataFilter(fltAR, 0))
        For I = 4 To 0 Step -1
            If tmp >= 2 ^ I Then
               chkDiffDisplay(I).Value = vbChecked
               tmp = tmp - 2 ^ I
            Else
               chkDiffDisplay(I).Value = vbUnchecked
            End If
        Next I
     Else
        For I = 0 To 4
            chkDiffDisplay(0).Value = vbUnchecked
        Next I
     End If
     'put ER range no matter are we using it or not
     If .DataFilter(fltAR, 1) < 0 Then
        txtERMin.Text = ""
     Else
        txtERMin.Text = .DataFilter(fltAR, 1)
     End If
     If .DataFilter(fltAR, 2) < 0 Then
        txtERMax.Text = ""
     Else
        txtERMax.Text = .DataFilter(fltAR, 2)
     End If
     
     ' Even/Odd Scan Number
     If CBool(.DataFilter(fltEvenOddScanNumber, 0)) Then
        If .DataFilter(fltEvenOddScanNumber, 1) = eosOddScans Then
            cboEvenOddScanNumber.ListIndex = eosOddScans
        Else
            cboEvenOddScanNumber.ListIndex = eosEvenScans
            .DataFilter(fltEvenOddScanNumber, 1) = eosEvenScans
        End If
     Else
        cboEvenOddScanNumber.ListIndex = eosAllScans
     End If
End With
UpdateFilterStatus blnQueryIfNoVisiblePoints
mUpdatingSettings = False
End Sub


Private Function FilterChanged() As Boolean
Dim I As Long, j As Long
On Error Resume Next
FilterChanged = False
For I = 1 To MAX_FILTER_COUNT
    For j = 0 To 2
        If OldSettings(I, j) <> GelData(CallerID).DataFilter(I, j) Then
           FilterChanged = True
           Exit Function
        End If
    Next j
Next I
End Function

Public Function FilterRange(ByVal I As Long) As Boolean
'------------------------------------------------------------------------------
'verification that filter parameters have sense; this function serves abundance
'molecular mass and charge state for isotopic data ranges
'------------------------------------------------------------------------------
Dim sMsg As String
On Error Resume Next
Select Case I
Case fltCSAbu
  sMsg = "Charge State Abundance Range"
Case fltIsoAbu
  sMsg = "Isotopic Abundance Range"
Case fltCSAbu
  sMsg = "Charge State Molecular Mass Range"
Case fltIsoAbu
  sMsg = "Isotopic Molecular Mass Range"
Case fltIsoCS
  sMsg = "Isotopic Charge State Range"
End Select
If (GelData(CallerID).DataFilter(I, 1) > GelData(CallerID).DataFilter(I, 2) Or _
   (GelData(CallerID).DataFilter(I, 2) <= 0)) Then
        If SayUserURDumb(sMsg) = vbYes Then
           FilterRange = True
        Else
           FilterRange = False
        End If
Else
    FilterRange = True
End If
End Function


Public Function FilterIdentity() As Boolean
On Error Resume Next
If GelData(CallerID).DataFilter(fltID, 1) < 3 Then
   FilterIdentity = True
Else
   If SayUserURDumb("Identity") = vbYes Then
      FilterIdentity = True
   Else
      FilterIdentity = False
   End If
End If
End Function

Public Function FilterDiffDisplay() As Boolean
With GelData(CallerID)
     Select Case .DataFilter(fltAR, 0)
     Case 0, 1, 2, 4, 5, 6, 8, 9, 10, 12, 13, 14
       FilterDiffDisplay = True
     Case 16, 18, 20, 22, 24, 26, 28, 30
       'check range boundaries
       If .DataFilter(fltAR, 1) < .DataFilter(fltAR, 2) Then
          FilterDiffDisplay = True
       ElseIf .DataFilter(fltAR, 1) > 0 And .DataFilter(fltAR, 2) < 0 Then
          FilterDiffDisplay = True
       Else
          If SayUserURDumb("Differential Display") = vbYes Then
             FilterDiffDisplay = True
          Else
             FilterDiffDisplay = False
          End If
       End If
     Case Else
       If SayUserURDumb("Differential Display") = vbYes Then
          FilterDiffDisplay = True
       Else
          FilterDiffDisplay = False
       End If
     End Select
End With
End Function

Public Function FilterIsotopicFit() As Boolean
Dim aBadFits As Variant
Dim BadFitsCount As Long
Dim j As Long
On Error Resume Next
With GelData(CallerID)
    If .DataFilter(fltIsoFit, 0) Then
       If .DataFilter(fltIsoFit, 1) > 0 Then
           aBadFits = GetBadFits(CallerID, CDbl(.DataFilter(fltIsoFit, 1)))
           If Not IsNull(aBadFits) Then
              BadFitsCount = UBound(aBadFits)
              With GelDraw(CallerID)
                For j = 1 To BadFitsCount
                    .IsoID(aBadFits(j)) = -Abs(.IsoID(aBadFits(j)))
                Next j
              End With
           End If
           FilterIsotopicFit = True
       Else
           If SayUserURDumb("Exclude data with calculated ...") = vbYes Then
              GelIsoExcludeAll (CallerID)
              FilterIsotopicFit = True
           Else
              FilterIsotopicFit = False
           End If
       End If
    Else
       FilterIsotopicFit = True
    End If
End With
End Function


Public Function FilterCSStDev() As Boolean
Dim aBadStDevs As Variant
Dim BadStDevsCount As Long
Dim j As Long
On Error Resume Next
With GelData(CallerID)
    If .DataFilter(fltCSStDev, 0) Then
       If .DataFilter(fltCSStDev, 1) > 0 Then
           aBadStDevs = GetBadStDevs(CallerID, CDbl(.DataFilter(fltCSStDev, 1)))
           If Not IsNull(aBadStDevs) Then
              BadStDevsCount = UBound(aBadStDevs)
              With GelDraw(CallerID)
                For j = 1 To BadStDevsCount
                    .CSID(aBadStDevs(j)) = -Abs(.CSID(aBadStDevs(j)))
                Next j
              End With
           End If
           FilterCSStDev = True
       Else
           If SayUserURDumb("Exclude data with St.Dev. ...") = vbYes Then
              GelCSExcludeAll (CallerID)
              FilterCSStDev = True
           Else
              FilterCSStDev = False
           End If
       End If
    Else
       FilterCSStDev = True
    End If
End With
End Function

Public Function FilterDuplicates() As Boolean
Dim aDuplicates As Variant
Dim DuplicatesCount As Long
Dim j As Long
On Error Resume Next
With GelData(CallerID)
    If .DataFilter(fltDupTolerance, 0) Then
       If .DataFilter(fltDupTolerance, 1) > 0 Then
          aDuplicates = GetDuplicates(CallerID, CDbl(.DataFilter(fltDupTolerance, 1)))
          If Not IsNull(aDuplicates) Then
             DuplicatesCount = UBound(aDuplicates)
             With GelDraw(CallerID)
               For j = 1 To DuplicatesCount
                 .IsoID(aDuplicates(j)) = -Abs(.IsoID(aDuplicates(j)))
               Next j
             End With
          End If
          FilterDuplicates = True
       Else
          If SayUserURDumb("Exclude duplicates ...") = vbYes Then
             GelIsoExcludeAll (CallerID)
             FilterDuplicates = True
          Else
             FilterDuplicates = False
          End If
       End If
    Else
       FilterDuplicates = True
    End If
End With
End Function

Public Function Filter2CloseGuesses() As Boolean
Dim aWorseGuesses As Variant
Dim WorseGuessesCount As Long
Dim I As Long
On Error Resume Next
With GelData(CallerID)
    If .DataFilter(fltCase2CloseResults, 0) Then
'we need to have Duplicate Tolerance at this point
        If Not IsNumeric(.DataFilter(fltDupTolerance, 1)) Then
            If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                MsgBox "Duplicate Tolerance has to be defined for this option." _
                & vbCrLf & "NOTE: It is not neccessary to apply Duplicate elimination, " _
                & "just define Duplicate Tolerance.", vbOKOnly, "2D Gel"
            End If
           Filter2CloseGuesses = False
        Else
           If .DataFilter(fltCase2CloseResults, 1) < 3 Then
              'aWorseGuesses = GetWorseGuess(CallerID, CInt(.DataFilter(fltCase2CloseResults, 1)), CDbl(.DataFilter(1, 1)))
              aWorseGuesses = GetWorseGuess(CallerID, CInt(.DataFilter(fltCase2CloseResults, 1)))
              If Not IsNull(aWorseGuesses) Then
                 WorseGuessesCount = UBound(aWorseGuesses)
                 With GelDraw(CallerID)
                    For I = 1 To WorseGuessesCount
                     .IsoID(aWorseGuesses(I)) = -Abs(.IsoID(aWorseGuesses(I)))
                    Next I
                 End With
              End If
              Filter2CloseGuesses = True
           Else
               If SayUserURDumb("Second Guess") = vbYes Then
                  GelIsoExcludeAll (CallerID)
                  Filter2CloseGuesses = True
               Else
                  Filter2CloseGuesses = False
               End If
           End If
        End If
    Else
       Filter2CloseGuesses = True
    End If
End With
End Function

Public Function FilterThis(blnQueryIfNoVisiblePoints As Boolean) As Boolean
'-----------------------------------------------------------
'apply individual filters one by one, if any fails reset all
'-----------------------------------------------------------

On Error GoTo FilterThisErrorHandler

If GelStatus(CallerID).Deleted Then
    Unload Me
    Exit Function
End If

FilterThis = False
GelCSIncludeAll (CallerID)
GelIsoIncludeAll (CallerID)
If chkCSAbuRange.Value = vbChecked Then
   If FilterRange(fltCSAbu) Then
      GelCSExcludeAbuRange (CallerID)
   Else
      ResetFiltersAndUpdateStatus
      Exit Function
   End If
End If
If chkIsoAbuRange.Value = vbChecked Then
   If FilterRange(fltIsoAbu) Then
      GelIsoExcludeAbuRange (CallerID)
   Else
      ResetFiltersAndUpdateStatus
      Exit Function
   End If
End If
If chkCSMWRange.Value = vbChecked Then
   If FilterRange(fltCSMW) Then
      GelCSExcludeMWRange (CallerID)
   Else
      ResetFiltersAndUpdateStatus
      Exit Function
   End If
End If
If chkIsoMWRange.Value = vbChecked Then
   If FilterRange(fltIsoMW) Then
      GelIsoExcludeMWRange (CallerID)
   Else
      ResetFiltersAndUpdateStatus
      Exit Function
   End If
End If
If chkIsoMZRange.Value = vbChecked Then
   If FilterRange(fltIsoMZ) Then
      GelIsoExcludeMZRange (CallerID)
   Else
      ResetFiltersAndUpdateStatus
      Exit Function
   End If
End If
If GelData(CallerID).DataFilter(fltEvenOddScanNumber, 0) Then
    GelExcludeEvenOddScans CallerID
End If
If chkIsoUseCSRange.Value = vbChecked Then
   If FilterRange(fltIsoCS) Then
      GelIsoExcludeCSRange (CallerID)
   Else
      ResetFiltersAndUpdateStatus
      Exit Function
   End If
End If
If GelData(CallerID).DataFilter(fltID, 0) Then
   If FilterIdentity Then
      GelCSExcludeIdentity CallerID
      GelIsoExcludeIdentity CallerID
   Else
      ResetFiltersAndUpdateStatus
      Exit Function
   End If
End If
If GelData(CallerID).DataFilter(fltAR, 0) Then
   If FilterDiffDisplay Then
      GelCSExcludeER CallerID
      GelIsoExcludeER CallerID
   Else
      ResetFiltersAndUpdateStatus
      Exit Function
   End If
End If
If Not FilterIsotopicFit Then
    ResetFiltersAndUpdateStatus
    Exit Function
End If
If Not FilterCSStDev Then
    ResetFiltersAndUpdateStatus
    Exit Function
End If
''If Not FilterDBFit Then
''    ResetFiltersAndUpdateStatus
''    Exit Function
''End If
If Not FilterDuplicates Then
    ResetFiltersAndUpdateStatus
    Exit Function
End If
If Not Filter2CloseGuesses Then
    ResetFiltersAndUpdateStatus
    Exit Function
End If
''Call FilterIsoCom               'this filter is not interfaced here
UpdateFilterStatus blnQueryIfNoVisiblePoints
FilterThis = True

Exit Function

FilterThisErrorHandler:
Debug.Print "Error in FilterThis (most likely one of its subroutines): " & Err.Description
Debug.Assert False
LogErrors Err.Number, "frmFilter->FilterThis (most likely one of the subroutines)"
FilterThis = False

End Function

Public Sub InitializeControls(Optional blnApplyFilterAndUnloadForm As Boolean = False)
On Error GoTo err_FilterActivate
CallerID = Me.Tag
If glTracking Then StopTracking
SaveOldSettings
DisplayCurrentSettings False
If blnApplyFilterAndUnloadForm Then
    ApplyFilterNow False
    ApplyFilterAndClose
    ' The following statement shouldn't have any effect since the form is unloaded in ApplyFilterAndClose
    ' However, we'll place it here just in case
    Unload Me
End If
Exit Sub

err_FilterActivate:
LogErrors Err.Number, "Filter-Activate"
If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
    MsgBox "Error loading filter for current display.", vbOKOnly, glFGTU
End If
Unload Me
End Sub

Private Sub PopulateComboBoxes()
    With cboEvenOddScanNumber
        .Clear
        .AddItem "Use All Scans"
        .AddItem "Use Odd Scans"
        .AddItem "Use Even Scans"
    End With
End Sub

Private Sub RestoreOldSettings()
Dim I As Long, j As Long
On Error Resume Next
For I = 1 To MAX_FILTER_COUNT
    For j = 0 To 2
        GelData(CallerID).DataFilter(I, j) = OldSettings(I, j)
    Next j
Next I
End Sub

Private Sub SaveOldSettings()
Dim I As Long, j As Long
On Error Resume Next
For I = 1 To MAX_FILTER_COUNT
    For j = 0 To 2
        OldSettings(I, j) = GelData(CallerID).DataFilter(I, j)
    Next j
Next I
End Sub

Private Function SayUserURDumb(sSection As String) As Integer
Dim sURDumb As String
If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
    sURDumb = "Filter specified in [" & sSection & "] section will exclude all points." _
        & vbCrLf & "Continue anyway?"
    SayUserURDumb = MsgBox(sURDumb, vbYesNo, glFGTU)
Else
    AddToAnalysisHistory CallerID, "Warning, Filter specified in [" & sSection & "] section will exclude all points."
    SayUserURDumb = vbYes
End If
End Function

Private Sub ToggleERDiffOption(ByVal erfCheckbox As erfERFilterConstants)
    
    With GelData(CallerID)
        If chkDiffDisplay(erfCheckbox).Value = vbChecked Then
           .DataFilter(fltAR, 0) = .DataFilter(fltAR, 0) + 2 ^ erfCheckbox
        Else
           .DataFilter(fltAR, 0) = .DataFilter(fltAR, 0) - 2 ^ erfCheckbox
        End If
    End With
    
    If chkDiffDisplay(erfERFilterConstants.erfExcludeWithExpression).Value = vbChecked And chkDiffDisplay(erfERFilterConstants.erfExcludeWithoutExpression).Value = vbChecked Then
        If erfCheckbox = erfERFilterConstants.erfExcludeWithExpression Then
            chkDiffDisplay(erfERFilterConstants.erfExcludeWithoutExpression).Value = vbUnchecked
        Else
            chkDiffDisplay(erfERFilterConstants.erfExcludeWithExpression).Value = vbUnchecked
        End If
    End If
    
    If chkDiffDisplay(erfERFilterConstants.erfExcludeWithExpression).Value = vbChecked And chkDiffDisplay(erfERFilterConstants.erfExcludeOutOfRange).Value = vbChecked Then
        chkDiffDisplay(erfERFilterConstants.erfExcludeOutOfRange).Value = vbUnchecked
    End If

End Sub

Private Sub ToggleExcludeIdentityOption(ByVal Index As Integer)

    With GelData(CallerID)
        If chkIdentity(Index).Value = vbChecked Then
           .DataFilter(fltID, 1) = .DataFilter(fltID, 1) + 2 ^ Index
        Else
           .DataFilter(fltID, 1) = .DataFilter(fltID, 1) - 2 ^ Index
        End If
        If .DataFilter(fltID, 1) > 0 Then
           .DataFilter(fltID, 0) = True
        Else
           .DataFilter(fltID, 0) = True
        End If
    End With
   
    ' Do not allow both Exclude Identified and Exclude Unidentifed to be checked
    If chkIdentity(0).Value = vbChecked And chkIdentity(1).Value = vbChecked Then
        If Index = 0 Then
            chkIdentity(1).Value = vbUnchecked
        Else
            chkIdentity(0).Value = vbUnchecked
        End If
    End If
End Sub

Private Sub ToggleWindowStayOnTop(blnEnableStayOnTop As Boolean)
    
    If mWindowStayOnTopEnabled = blnEnableStayOnTop Then Exit Sub
    
    Me.ScaleMode = vbTwips
    
    WindowStayOnTop Me.hwnd, blnEnableStayOnTop, Me.ScaleX(Me.Left, vbTwips, vbPixels), Me.ScaleY(Me.Top, vbTwips, vbPixels), Me.ScaleX(Me.width, vbTwips, vbPixels), Me.ScaleY(Me.Height, vbTwips, vbPixels)
    
    mWindowStayOnTopEnabled = blnEnableStayOnTop

End Sub

Private Sub UpdateFilterStatus(blnQueryIfNoVisiblePoints As Boolean)
    Dim I As Long
    Dim lngVisiblePoints As Long
    Dim lngTotalPoints As Long
    Dim strPercentVisible As String
    
    Dim eResponse As VbMsgBoxResult
    
On Error GoTo UpdateFilterStatusErrorHandler
    
    With GelDraw(CallerID)
        For I = 1 To .IsoCount
            If .IsoID(I) >= 0 Then lngVisiblePoints = lngVisiblePoints + 1
        Next I
        
        For I = 1 To .CSCount
            If .CSID(I) >= 0 Then lngVisiblePoints = lngVisiblePoints + 1
        Next I
        
        lngTotalPoints = .IsoCount + .CSCount
        
        If lngTotalPoints > 0 And lngVisiblePoints = 0 And blnQueryIfNoVisiblePoints Then
            If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
                ToggleWindowStayOnTop False
                eResponse = MsgBox("The current filter parameters exclude all " & Trim(lngTotalPoints) & " data points.  Continue using these filters?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "All Data Excluded")
                ToggleWindowStayOnTop True

                If eResponse <> vbYes Then
                    ResetFiltersAndUpdateStatus
                    Exit Sub
                End If
            End If
        End If
        frmFilter.SetFocus
        
        If lngVisiblePoints = lngTotalPoints Then
            lblFilterStatus = "All " & LongToStringWithCommas(lngTotalPoints) & " data points are visible"
        Else
            If lngVisiblePoints / lngTotalPoints >= 0.01 Then
                strPercentVisible = Format(lngVisiblePoints / lngTotalPoints, "##0.0%")
            Else
                strPercentVisible = Format(lngVisiblePoints / lngTotalPoints, "0.00%")
            End If
            lblFilterStatus = strPercentVisible & " visible (" & LongToStringWithCommas(lngVisiblePoints) & " out of " & LongToStringWithCommas(lngTotalPoints) & " data points)"
        End If
    End With

    Exit Sub
    
UpdateFilterStatusErrorHandler:
    lblFilterStatus = "0 data points total"

End Sub

Private Sub ResetFiltersAndUpdateStatus()
    ResetDataFilters CallerID, glPreferences     ' Reset to defaults
    If FilterChanged() Then   'if something was changed
        DisplayCurrentSettings False
        ApplyFilterNow False   'and refilter back changes
    Else
        UpdateFilterStatus False
    End If
End Sub

Private Sub Form_Activate()
    InitializeControls
End Sub

Private Sub Form_Load()
    SizeAndCenterWindow Me, cWindowBottomRight, 5900, 7025, False
    PopulateComboBoxes
    
    ToggleWindowStayOnTop True

    tbsFilters.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
If IsWinLoaded(TrackerCaption) Then glTracking = True
End Sub


Private Sub cboEvenOddScanNumber_Click()
If Not mUpdatingSettings Then
    GelData(CallerID).DataFilter(fltEvenOddScanNumber, 1) = cboEvenOddScanNumber.ListIndex
    If cboEvenOddScanNumber.ListIndex = eosAllScans Then
        GelData(CallerID).DataFilter(fltEvenOddScanNumber, 0) = False
    Else
        GelData(CallerID).DataFilter(fltEvenOddScanNumber, 0) = True
    End If
End If
End Sub

Private Sub chkBadStDevElimination_Click()
If Not mUpdatingSettings Then
   If chkBadStDevElimination.Value = vbChecked Then
      txtStDevTolerance.Enabled = True
      txtStDevTolerance.SetFocus
      GelData(CallerID).DataFilter(fltCSStDev, 0) = True
   Else
      txtStDevTolerance.Enabled = False
      GelData(CallerID).DataFilter(fltCSStDev, 0) = False
   End If
End If
End Sub

Private Sub chkCSAbuRange_Click()
If Not mUpdatingSettings Then
   If chkCSAbuRange.Value = vbChecked Then
      txtCSMinAbu.Enabled = True
      txtCSMaxAbu.Enabled = True
      txtCSMaxAbu.SetFocus
      GelData(CallerID).DataFilter(fltCSAbu, 0) = True
   Else
      txtCSMinAbu.Enabled = False
      txtCSMaxAbu.Enabled = False
      GelData(CallerID).DataFilter(fltCSAbu, 0) = False
   End If
   If chkCSIsoSameRange.Value = vbChecked Then chkIsoAbuRange.Value = chkCSAbuRange.Value
End If
End Sub

Private Sub chkBadFitElimination_Click()
If Not mUpdatingSettings Then
   If chkBadFitElimination.Value = vbChecked Then
      txtFitTolerance.Enabled = True
      txtFitTolerance.SetFocus
      GelData(CallerID).DataFilter(fltIsoFit, 0) = True
   Else
      txtFitTolerance.Enabled = False
      GelData(CallerID).DataFilter(fltIsoFit, 0) = False
   End If
End If
End Sub

Private Sub chkCSIsoSameRange_Click()
If Not mUpdatingSettings Then
   If chkCSIsoSameRange.Value = vbChecked Then
'take Iso parameters from the Charge State
      chkIsoAbuRange.Value = chkCSAbuRange.Value
      txtIsoMinAbu.Text = txtCSMinAbu.Text
      txtIsoMaxAbu.Text = txtCSMaxAbu.Text
      With GelData(CallerID)
           .DataFilter(fltIsoAbu, 0) = .DataFilter(fltCSAbu, 0)
           .DataFilter(fltIsoAbu, 1) = .DataFilter(fltCSAbu, 1)
           .DataFilter(fltIsoAbu, 2) = .DataFilter(fltCSAbu, 2)
      End With
   End If
End If
End Sub

Private Sub chkCSIsoSameRangeMW_Click()
If Not mUpdatingSettings Then
   If chkCSIsoSameRangeMW.Value = vbChecked Then
      'take Iso parameters from the Charge State
      chkIsoMWRange.Value = chkCSMWRange.Value
      txtIsoMinMW.Text = txtCSMinMW.Text
      txtIsoMaxMW.Text = txtCSMaxMW.Text
      With GelData(CallerID)
           .DataFilter(fltIsoMW, 0) = .DataFilter(fltCSMW, 0)
           .DataFilter(fltIsoMW, 1) = .DataFilter(fltCSMW, 1)
           .DataFilter(fltIsoMW, 2) = .DataFilter(fltCSMW, 2)
      End With
   End If
End If
End Sub

Private Sub chkCSMWRange_Click()
If Not mUpdatingSettings Then
   If chkCSMWRange.Value = vbChecked Then
      txtCSMinMW.Enabled = True
      txtCSMaxMW.Enabled = True
      txtCSMaxMW.SetFocus
      GelData(CallerID).DataFilter(fltCSMW, 0) = True
   Else
      txtCSMinMW.Enabled = False
      txtCSMaxMW.Enabled = False
      GelData(CallerID).DataFilter(fltCSMW, 0) = False
   End If
   If chkCSIsoSameRangeMW.Value = vbChecked Then chkIsoMWRange.Value = chkCSMWRange.Value
End If
End Sub

Private Sub chkDBBadFitElimination_Click()
If Not mUpdatingSettings Then
   If chkDBBadFitElimination.Value = vbChecked Then
      txtDBFitTolerance.Enabled = True
      txtDBFitTolerance.SetFocus
      GelData(CallerID).DataFilter(fltDBTolerance, 0) = True
   Else
      txtDBFitTolerance.Enabled = False
      GelData(CallerID).DataFilter(fltDBTolerance, 0) = False
   End If
End If
End Sub

Private Sub chkDiffDisplay_Click(Index As Integer)
    If Not mUpdatingSettings Then
        ToggleERDiffOption Index
    End If
End Sub

Private Sub chkDupElimination_Click()
If Not mUpdatingSettings Then
   If chkDupElimination.Value = vbChecked Then
      txtDupTolerance.Enabled = True
      txtDupTolerance.SetFocus
      GelData(CallerID).DataFilter(fltDupTolerance, 0) = True
   Else
      txtDupTolerance.Enabled = False
      GelData(CallerID).DataFilter(fltDupTolerance, 0) = False
   End If
End If
End Sub

Private Sub chkIdentity_Click(Index As Integer)
    If Not mUpdatingSettings Then
        ToggleExcludeIdentityOption Index
    End If
End Sub

Private Sub chkIsoAbuRange_Click()
If Not mUpdatingSettings Then
   If chkIsoAbuRange.Value = vbChecked Then
      txtIsoMinAbu.Enabled = True
      txtIsoMaxAbu.Enabled = True
      txtIsoMaxAbu.SetFocus
      GelData(CallerID).DataFilter(fltIsoAbu, 0) = True
   Else
      txtIsoMinAbu.Enabled = False
      txtIsoMaxAbu.Enabled = False
      GelData(CallerID).DataFilter(fltIsoAbu, 0) = False
   End If
   If chkCSIsoSameRange.Value = vbChecked Then chkCSAbuRange.Value = chkIsoAbuRange.Value
End If
End Sub

Private Sub chkIsoMWRange_Click()
If Not mUpdatingSettings Then
   If chkIsoMWRange.Value = vbChecked Then
      txtIsoMinMW.Enabled = True
      txtIsoMaxMW.Enabled = True
      txtIsoMaxMW.SetFocus
      GelData(CallerID).DataFilter(fltIsoMW, 0) = True
   Else
      txtIsoMinMW.Enabled = False
      txtIsoMaxMW.Enabled = False
      GelData(CallerID).DataFilter(fltIsoMW, 0) = False
   End If
   If chkCSIsoSameRangeMW.Value = vbChecked Then chkCSMWRange.Value = chkIsoMWRange.Value
End If
End Sub

Private Sub chkIsoMZRange_Click()
If Not mUpdatingSettings Then
   If chkIsoMZRange.Value = vbChecked Then
      txtIsoMinMZ.Enabled = True
      txtIsoMaxMZ.Enabled = True
      txtIsoMaxMZ.SetFocus
      GelData(CallerID).DataFilter(fltIsoMZ, 0) = True
   Else
      txtIsoMinMZ.Enabled = False
      txtIsoMaxMZ.Enabled = False
      GelData(CallerID).DataFilter(fltIsoMZ, 0) = False
   End If
End If
End Sub

Private Sub chkIsoUseCSRange_Click()
If Not mUpdatingSettings Then
   If chkIsoUseCSRange.Value = vbChecked Then
      txtIsoMinCS.Enabled = True
      txtIsoMaxCS.Enabled = True
      txtIsoMaxCS.SetFocus
      GelData(CallerID).DataFilter(fltIsoCS, 0) = True
   Else
      txtIsoMinCS.Enabled = False
      txtIsoMaxCS.Enabled = False
      GelData(CallerID).DataFilter(fltIsoCS, 0) = False
   End If
End If
End Sub

Private Sub chkSecGuessElimination_Click(Index As Integer)
If Not mUpdatingSettings Then
   With GelData(CallerID)
        If chkSecGuessElimination(Index).Value = vbChecked Then
           .DataFilter(fltCase2CloseResults, 1) = .DataFilter(fltCase2CloseResults, 1) + 2 ^ Index
        Else
           .DataFilter(fltCase2CloseResults, 1) = .DataFilter(fltCase2CloseResults, 1) - 2 ^ Index
        End If
        If .DataFilter(fltCase2CloseResults, 1) > 0 Then
           .DataFilter(fltCase2CloseResults, 0) = True
        Else
           .DataFilter(fltCase2CloseResults, 0) = False
        End If
   End With
End If
End Sub

Private Sub cmdApplyNow_Click()
ApplyFilterNow True
End Sub

Private Sub cmdCancel_Click()
    CancelChanges
End Sub

Private Sub cmdOK_Click()
    ApplyFilterAndClose
End Sub

Private Sub cmdReset_Click()
    ResetFiltersAndUpdateStatus
End Sub

Private Sub cmdUseDefaultFilters_Click()
    ApplyDefaultFilters
End Sub

Private Sub txtCSMaxAbu_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtCSMaxAbu, KeyAscii, True, True, True, False, True, False, False, False, False, True
End Sub

Private Sub txtCSmaxmw_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtCSMaxMW, KeyAscii, True, True
End Sub

Private Sub txtCSmaxmw_LostFocus()
On Error Resume Next
If Not IsNumeric(txtCSMaxMW.Text) Then
   txtCSMaxMW.Text = Format(OldSettings(fltCSMW, 2), "0.00")
Else
   GelData(CallerID).DataFilter(fltCSMW, 2) = CDbl(txtCSMaxMW.Text)
   txtCSMaxMW.Text = Format(GelData(CallerID).DataFilter(fltCSMW, 2), "0.00")
End If
If chkCSIsoSameRangeMW.Value = vbChecked Then
   txtIsoMaxMW.Text = txtCSMaxMW.Text
   GelData(CallerID).DataFilter(fltIsoMW, 2) = GelData(CallerID).DataFilter(fltCSMW, 2)
End If
End Sub

Private Sub txtCSMinAbu_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtCSMinAbu, KeyAscii, True, True, True, False, True, False, False, False, False, True
End Sub

Private Sub txtCSMinMW_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtCSMinMW, KeyAscii, True, True
End Sub

Private Sub txtCSMinMW_LostFocus()
On Error Resume Next
If Not IsNumeric(txtCSMinMW.Text) Then
   txtCSMinMW.Text = Format(OldSettings(fltCSMW, 1), "0.00")
Else
   GelData(CallerID).DataFilter(fltCSMW, 1) = CDbl(txtCSMinMW.Text)
   txtCSMinMW.Text = Format(GelData(CallerID).DataFilter(fltCSMW, 1), "0.00")
End If
If chkCSIsoSameRangeMW.Value = vbChecked Then
   txtIsoMinMW.Text = txtCSMinMW.Text
   GelData(CallerID).DataFilter(fltIsoMW, 1) = GelData(CallerID).DataFilter(fltCSMW, 1)
End If
End Sub

Private Sub txtDBFitTolerance_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtDBFitTolerance, KeyAscii, True, True
End Sub

Private Sub txtDBFitTolerance_LostFocus()
On Error Resume Next
If Not IsNumeric(txtDBFitTolerance.Text) And cChkBox(chkDBBadFitElimination) Then
   txtDBFitTolerance.Text = OldSettings(fltDBTolerance, 1)
Else
   GelData(CallerID).DataFilter(fltDBTolerance, 1) = CDbl(txtDBFitTolerance.Text)
End If
End Sub

Private Sub txtDupTolerance_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtDupTolerance, KeyAscii, True, True
End Sub

Private Sub txtDupTolerance_LostFocus()
On Error Resume Next
If Not IsNumeric(txtDupTolerance.Text) Then
   txtDupTolerance.Text = OldSettings(fltDupTolerance, 1)
Else
   GelData(CallerID).DataFilter(fltDupTolerance, 1) = CDbl(txtDupTolerance.Text)
End If
End Sub

Private Sub txtERMax_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtERMax, KeyAscii, True, True, True, False, True, False, False, False, False, True
End Sub

Private Sub txtERMax_LostFocus()
Dim TmpERMax As Double
On Error Resume Next
If Len(Trim$(txtERMax.Text)) > 0 Then
   If Not IsNumeric(txtERMax.Text) Then
      GoTo txtERMax_Invalid
   Else
      TmpERMax = CDbl(txtERMax.Text)
      If TmpERMax >= 0 Then
         GelData(CallerID).DataFilter(fltAR, 2) = TmpERMax
      Else
         GoTo txtERMax_Invalid
      End If
   End If
Else
   GelData(CallerID).DataFilter(fltAR, 2) = -1
End If
Exit Sub

txtERMax_Invalid:
If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
    MsgBox "This argument should be positive number or left blank.", vbOKOnly, glFGTU
End If
txtERMax.SetFocus
txtERMax.Text = OldSettings(fltAR, 2)
End Sub

Private Sub txtERMin_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtERMin, KeyAscii, True, True, True, False, True, False, False, False, False, True
End Sub

Private Sub txtERMin_LostFocus()
Dim TmpERMin As Double
On Error Resume Next
If Len(Trim$(txtERMin.Text)) > 0 Then
   If Not IsNumeric(txtERMin.Text) Then
      GoTo txtERMin_Invalid
   Else
      TmpERMin = CDbl(txtERMin.Text)
      If TmpERMin >= 0 Then
         GelData(CallerID).DataFilter(fltAR, 1) = TmpERMin
      Else
         GoTo txtERMin_Invalid
      End If
   End If
Else
   GelData(CallerID).DataFilter(fltAR, 1) = -1
End If
Exit Sub

txtERMin_Invalid:
If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
    MsgBox "This argument should be non-negative number or left blank.", vbOKOnly, glFGTU
End If
txtERMin.SetFocus
txtERMin.Text = OldSettings(fltAR, 1)
End Sub

Private Sub txtFitTolerance_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtFitTolerance, KeyAscii, True, True
End Sub

Private Sub txtFitTolerance_LostFocus()
On Error Resume Next
If Not IsNumeric(txtFitTolerance.Text) Then
   txtFitTolerance.Text = OldSettings(fltIsoFit, 1)
Else
   GelData(CallerID).DataFilter(fltIsoFit, 1) = Abs(CDbl(txtFitTolerance.Text))
End If
End Sub

Private Sub txtCSMaxAbu_LostFocus()
With GelData(CallerID)
     If Not IsNumeric(txtCSMaxAbu.Text) Then
        txtCSMaxAbu.Text = Format(OldSettings(fltCSAbu, 2), "Scientific")
     Else
        .DataFilter(fltCSAbu, 2) = CDbl(txtCSMaxAbu.Text)
        txtCSMaxAbu.Text = Format(.DataFilter(fltCSAbu, 2), "Scientific")
     End If
     If chkCSIsoSameRange.Value = vbChecked Then
        txtIsoMaxAbu.Text = txtCSMaxAbu.Text
        .DataFilter(fltIsoAbu, 2) = .DataFilter(fltCSAbu, 2)
     End If
End With
End Sub

Private Sub txtCSMinAbu_LostFocus()
With GelData(CallerID)
     If Not IsNumeric(txtCSMinAbu.Text) Then
        txtCSMinAbu.Text = Format(OldSettings(fltCSAbu, 1), "Scientific")
     Else
        .DataFilter(fltCSAbu, 1) = CDbl(txtCSMinAbu.Text)
        txtCSMinAbu.Text = Format(.DataFilter(fltCSAbu, 1), "Scientific")
     End If
     If chkCSIsoSameRange.Value = vbChecked Then
        txtIsoMinAbu.Text = txtCSMinAbu.Text
        .DataFilter(fltIsoAbu, 1) = .DataFilter(fltCSAbu, 1)
     End If
End With
End Sub

Private Sub txtIsoMaxAbu_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtIsoMaxAbu, KeyAscii, True, True, True, False, True, False, False, False, False, True
End Sub

Private Sub txtIsoMaxAbu_LostFocus()
With GelData(CallerID)
     If Not IsNumeric(txtIsoMaxAbu.Text) Then
        txtIsoMaxAbu.Text = Format(OldSettings(fltIsoAbu, 2), "Scientific")
     Else
        .DataFilter(fltIsoAbu, 2) = CDbl(txtIsoMaxAbu.Text)
        txtIsoMaxAbu.Text = Format(.DataFilter(fltIsoAbu, 2), "Scientific")
     End If
     If chkCSIsoSameRange.Value = vbChecked Then
        txtCSMaxAbu.Text = txtIsoMaxAbu.Text
        .DataFilter(fltCSAbu, 2) = .DataFilter(fltIsoAbu, 2)
     End If
End With
End Sub

Private Sub txtIsoMaxCS_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtIsoMaxCS, KeyAscii, True, False
End Sub

Private Sub txtIsoMaxCS_LostFocus()
On Error Resume Next
If Not IsNumeric(txtIsoMaxCS.Text) Then
   txtIsoMaxCS.Text = OldSettings(fltIsoCS, 2)
Else
   GelData(CallerID).DataFilter(fltIsoCS, 2) = CLng(txtIsoMaxCS.Text)
   txtIsoMaxCS.Text = GelData(CallerID).DataFilter(fltIsoCS, 2)
End If
End Sub

Private Sub txtIsomaxmw_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtIsoMaxMW, KeyAscii, True, True
End Sub

Private Sub txtIsomaxmw_LostFocus()
On Error Resume Next
If Not IsNumeric(txtIsoMaxMW.Text) Then
   txtIsoMaxMW.Text = Format(OldSettings(fltIsoMW, 2), "0.00")
Else
   GelData(CallerID).DataFilter(fltIsoMW, 2) = CDbl(txtIsoMaxMW.Text)
   txtIsoMaxMW.Text = Format(GelData(CallerID).DataFilter(fltIsoMW, 2), "0.00")
End If
If chkCSIsoSameRangeMW.Value = vbChecked Then
   txtCSMaxMW.Text = txtIsoMaxMW.Text
   GelData(CallerID).DataFilter(fltCSMW, 2) = GelData(CallerID).DataFilter(fltIsoMW, 2)
End If
End Sub

Private Sub txtIsoMaxMZ_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtIsoMaxMZ, KeyAscii, True, True
End Sub

Private Sub txtIsoMaxMZ_LostFocus()
On Error Resume Next
If Not IsNumeric(txtIsoMaxMZ.Text) Then
   txtIsoMaxMZ.Text = Format(OldSettings(fltIsoMZ, 2), "0.00")
Else
   GelData(CallerID).DataFilter(fltIsoMZ, 2) = CDbl(txtIsoMaxMZ.Text)
   txtIsoMaxMZ.Text = Format(GelData(CallerID).DataFilter(fltIsoMZ, 2), "0.00")
End If
End Sub

Private Sub txtIsoMinAbu_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtIsoMinAbu, KeyAscii, True, True, True, False, True, False, False, False, False, True
End Sub

Private Sub txtIsoMinAbu_LostFocus()
With GelData(CallerID)
     If Not IsNumeric(txtIsoMinAbu.Text) Then
        txtIsoMinAbu.Text = Format(OldSettings(fltIsoAbu, 1), "Scientific")
     Else
        .DataFilter(fltIsoAbu, 1) = CDbl(txtIsoMinAbu.Text)
        txtIsoMinAbu.Text = Format(.DataFilter(fltIsoAbu, 1), "Scientific")
     End If
     If chkCSIsoSameRange.Value = vbChecked Then
        txtCSMinAbu.Text = txtIsoMinAbu.Text
        .DataFilter(fltCSAbu, 1) = .DataFilter(fltIsoAbu, 1)
     End If
End With
End Sub

Private Sub txtIsoMinCS_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtIsoMinCS, KeyAscii, True, False
End Sub

Private Sub txtIsoMinCS_LostFocus()
On Error Resume Next
If Not IsNumeric(txtIsoMinCS.Text) Then
   txtIsoMinCS.Text = OldSettings(fltIsoCS, 1)
Else
   GelData(CallerID).DataFilter(fltIsoCS, 1) = CLng(txtIsoMinCS.Text)
   txtIsoMinCS.Text = GelData(CallerID).DataFilter(fltIsoCS, 1)
End If
End Sub

Private Sub txtIsoMinMW_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtIsoMinMW, KeyAscii, True, True
End Sub

Private Sub txtIsoMinMW_LostFocus()
On Error Resume Next
If Not IsNumeric(txtIsoMinMW.Text) Then
   txtIsoMinMW.Text = Format(OldSettings(fltIsoMW, 1), "0.00")
Else
   GelData(CallerID).DataFilter(fltIsoMW, 1) = CDbl(txtIsoMinMW.Text)
   txtIsoMinMW.Text = Format(GelData(CallerID).DataFilter(fltIsoMW, 1), "0.00")
End If
If chkCSIsoSameRangeMW.Value = vbChecked Then
   txtCSMinMW.Text = txtIsoMinMW.Text
   GelData(CallerID).DataFilter(fltCSMW, 1) = GelData(CallerID).DataFilter(fltIsoMW, 1)
End If
End Sub

Private Sub txtIsoMinMZ_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtIsoMinMZ, KeyAscii, True, True
End Sub

Private Sub txtIsoMinMZ_LostFocus()
On Error Resume Next
If Not IsNumeric(txtIsoMinMZ.Text) Then
   txtIsoMinMZ.Text = Format(OldSettings(fltIsoMZ, 1), "0.00")
Else
   GelData(CallerID).DataFilter(fltIsoMZ, 1) = CDbl(txtIsoMinMZ.Text)
   txtIsoMinMZ.Text = Format(GelData(CallerID).DataFilter(fltIsoMZ, 1), "0.00")
End If
End Sub

Private Sub txtStDevTolerance_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtStDevTolerance, KeyAscii, True, True
End Sub

Private Sub txtStDevTolerance_LostFocus()
On Error Resume Next
If Not IsNumeric(txtStDevTolerance.Text) Then
   txtStDevTolerance.Text = OldSettings(fltCSStDev, 1)
Else
   GelData(CallerID).DataFilter(fltCSStDev, 1) = Abs(CDbl(txtStDevTolerance.Text))
End If
End Sub

