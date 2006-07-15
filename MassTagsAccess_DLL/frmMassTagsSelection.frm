VERSION 5.00
Begin VB.Form frmMassTagsSelection 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mass Tags Selection"
   ClientHeight    =   7515
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6480
   Icon            =   "frmMassTagsSelection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraSettings3 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   120
      TabIndex        =   43
      Top             =   4080
      Width           =   6135
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   375
         Left            =   4560
         TabIndex        =   39
         Top             =   2760
         Width           =   975
      End
      Begin VB.ListBox lstIncList 
         Height          =   2010
         Left            =   2880
         TabIndex        =   37
         Top             =   480
         Width           =   3015
      End
      Begin VB.CommandButton cmdClearInclude 
         Caption         =   "Clear"
         Height          =   315
         Left            =   5040
         TabIndex        =   36
         ToolTipText     =   "Clears inclusion list"
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3360
         TabIndex        =   38
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton cmdInclude 
         Caption         =   "I&nclude >"
         Height          =   375
         Left            =   1680
         TabIndex        =   28
         Top             =   1440
         Width           =   975
      End
      Begin VB.Frame fraD 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Dynamic Modifications"
         Height          =   1215
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   2655
         Begin VB.OptionButton optDIsIsNotAny 
            BackColor       =   &H00C0E0FF&
            Caption         =   "I&s"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optDIsIsNotAny 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Is &Not"
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   23
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optDIsIsNotAny 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Any"
            Height          =   255
            Index           =   2
            Left            =   1800
            TabIndex        =   24
            Top             =   360
            Width           =   615
         End
         Begin VB.ComboBox cmbDMod 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   720
            Width           =   2415
         End
      End
      Begin VB.Frame fraS 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Static Modifications"
         Height          =   1215
         Left            =   0
         TabIndex        =   29
         Top             =   1920
         Width           =   2655
         Begin VB.OptionButton optSIsIsNotAny 
            BackColor       =   &H00C0E0FF&
            Caption         =   "I&s"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   30
            Top             =   360
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optSIsIsNotAny 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Is &Not"
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   31
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optSIsIsNotAny 
            BackColor       =   &H00C0E0FF&
            Caption         =   "&Any"
            Height          =   255
            Index           =   2
            Left            =   1800
            TabIndex        =   32
            Top             =   360
            Width           =   615
         End
         Begin VB.ComboBox cmbSMod 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   720
            Width           =   2415
         End
      End
      Begin VB.OptionButton optAndOr 
         Caption         =   "&AND"
         Height          =   375
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1440
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton optAndOr 
         Caption         =   "&OR"
         Height          =   375
         Index           =   1
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton cmdIncItemDelete 
         Caption         =   "&Del."
         Height          =   315
         Left            =   4080
         TabIndex        =   35
         ToolTipText     =   "Delete selected item from the list"
         Top             =   60
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inclusion List"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   34
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame fraSettings2 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   120
      TabIndex        =   40
      Top             =   1440
      Width           =   6135
      Begin VB.ComboBox cboInternalStandardExplicit 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox txtExperimentExclusionFilter 
         Height          =   285
         Left            =   2160
         TabIndex        =   18
         Top             =   1860
         Width           =   2175
      End
      Begin VB.TextBox txtExperimentInclusionFilter 
         Height          =   285
         Left            =   2160
         TabIndex        =   16
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox txtMinimumHighDiscriminantScore 
         Height          =   285
         Left            =   5520
         TabIndex        =   10
         Text            =   "0"
         Top             =   375
         Width           =   495
      End
      Begin VB.TextBox txtMinimumHighNormalizedScore 
         Height          =   285
         Left            =   5520
         TabIndex        =   8
         Text            =   "0"
         Top             =   75
         Width           =   495
      End
      Begin VB.ComboBox cboNETValueType 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox txtMinimumPMTQualityScore 
         Height          =   285
         Left            =   5520
         TabIndex        =   12
         Text            =   "0"
         Top             =   675
         Width           =   495
      End
      Begin VB.Label lblInternalStdExplicit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Explicit internal standard:"
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label lblExperimentExclusionFilter 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Experiment Exclusion Filter:"
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   1860
         Width           =   2175
      End
      Begin VB.Label lblExperimentInclusionFilter 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Experiment Inclusion Filter:"
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label lblMinimumHighDiscriminantScore 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Minimum high discriminant score (0 to load all mass tags, regardless of score):"
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   420
         Width           =   5535
      End
      Begin VB.Label lblMinimumHighNormalizedScore 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Minimum high normalized XCorr (0 to load all mass tags, regardless of score):"
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   120
         Width           =   5535
      End
      Begin VB.Label lblNETValueType 
         BackColor       =   &H00C0E0FF&
         Caption         =   "NET Values to use:"
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   1110
         Width           =   1575
      End
      Begin VB.Label lblMinimumPMTQualityScore 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Minimum PMT quality score (0 to load all mass tags, regardless of score):"
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   720
         Width           =   5535
      End
   End
   Begin VB.Frame fraSettings1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   120
      TabIndex        =   41
      Top             =   120
      Width           =   6135
      Begin VB.CheckBox chkLimitToPMTsFromDataset 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Limit to PMTs from dataset for given job"
         Height          =   255
         Left            =   2040
         TabIndex        =   1
         Top             =   0
         Width           =   4095
      End
      Begin VB.CheckBox chkUseAll 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Use all mass tags (belonging to MT subset)"
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   880
         Width           =   2175
      End
      Begin VB.CheckBox chkUseMTSubset 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Use subset"
         Height          =   255
         Left            =   2040
         TabIndex        =   4
         Top             =   330
         Width           =   1095
      End
      Begin VB.ComboBox cmbMTSubset 
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   300
         Width           =   2655
      End
      Begin VB.CheckBox chkConfirmedOnly 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Confirmed mass tags"
         Height          =   255
         Left            =   0
         TabIndex        =   0
         ToolTipText     =   "Use confirmed mass tags only"
         Top             =   0
         Width           =   1815
      End
      Begin VB.CheckBox chkAccurateOnly 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Accurate mass tags"
         Height          =   255
         Left            =   0
         TabIndex        =   2
         ToolTipText     =   "Use accurate mass tags only (AMT)"
         Top             =   300
         Width           =   1815
      End
      Begin VB.CheckBox chkLockersOnly 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Lockers mass tags"
         Height          =   255
         Left            =   0
         TabIndex        =   3
         ToolTipText     =   "Select only among lockers mass tags"
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblMTSubsetDescription 
         BackStyle       =   0  'Transparent
         Height          =   585
         Left            =   2040
         TabIndex        =   42
         Top             =   720
         Width           =   4005
      End
   End
End
Attribute VB_Name = "frmMassTagsSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this dialog is used to specify which mass tags should be used for search
'scope is determined with MTSubsetID and IncList
'MTSubsetID=-1      don't use subsets at all(use all mass tags)
'MTSubsetID=""      subsets were not specified yet
'IncList=-1         don't use inclusion list(use all mass tags)
'IncList=""         inclusion list was not specified yet
'new addition is selection of Confirmed/Accurate/Lockers mass tags only
'NetValueType=0          use GANET values from DB (nvtGANET)
'NetValueType=1          use PNet values from DB (nvtPNET)
'NetValueType=2          use theoretical NET values (computed using GANETClass.dll) (nvtTheoreticalNET)
'------------------------------------------------------------------------
'created: 12/07/2001 nt
'last modified: 12/21/2005 mem
'------------------------------------------------------------------------
Option Explicit

'names of database properties(Names in DBStuff list)
'containing subset and inclusion list for mass tags search
'added confirmed and AMT only selection
Const NAME_SUBSET As String = "MTSubset ID"                           ' Not used in DB Schema Version 2
Const NAME_INC_LIST As String = "Search Inclusion List"
Const NAME_CONFIRMED_ONLY As String = "Confirmed Only"
Const NAME_ACCURATE_ONLY As String = "Accurate Only"                  ' Not used in DB Schema Version 2
Const NAME_LOCKERS_ONLY As String = "Lockers Only"                    ' Not used in DB Schema Version 2
Const NAME_LIMIT_TO_PMTS_FROM_DATASET As String = "Limit to PMTs from Dataset"

Const NAME_MINIMUM_HIGH_NORMALIZED_SCORE As String = "Minimum High Normalized Score"
Const NAME_MINIMUM_HIGH_DISCRIMINANT_SCORE As String = "Minimum High Discriminant Score"
Const NAME_MINIMUM_PMT_QUALITY_SCORE As String = "MinimumPMTQualityScore"
Const NAME_EXPERIMENT_INCLUSION_FILTER As String = "Experiment Inclusion Filter"
Const NAME_EXPERIMENT_EXCLUSION_FILTER As String = "Experiment Exclusion Filter"
Const NAME_INTERNAL_STD_EXPLICIT As String = "Internal Standard Explicit"

Const NAME_NET_VALUE_TYPE As String = "NETValueType"

Const VAL_TRUE As String = "True"
Const VAL_FALSE As String = "False"

'name in DBStuff collection containing list
Const MT_SELECTION_LIST_NAME As String = "MassTagsSelectionList"

'name of pair containing global modifications list access query
'NOTE:View is used here so that link to Peptide database doesn't
'have to be managed in initialization file
'NOTE:Link to peptide database is replaced with link to MT_Main
'however view is still used instead of stored procedure so that
'we have "Diversity In Initialization File" :-)
Const INIT_Get_Global_Mod_D As String = "v_Get_Global_Mods_D"    ' Not used in DB Schema Version 2
Const INIT_Get_Global_Mod_S As String = "v_Get_Global_Mods_S"    ' Not used in DB Schema Version 2
Const INIT_Fill_MTSubsets As String = "sql_GET_Subsets_List"     ' Not used in DB Schema Version 2

Const INIT_Get_Global_Mass_Correction_Factors As String = "v_Get_Global_Mass_Correction_Factors"     ' Not used in DB Schema Version 1
Const INIT_Get_Internal_Standard_Names As String = "v_Get_Internal_Standard_Names"                   ' Not used in DB Schema Version 1

Const INIT_Get_DB_Schema_Version As String = "sp_GetDBSchemaVersion"


'following words are admissable in inclusion items
'only other words(phrases) allowed are coming from
'D/SGlobModName arrays
Const VC_Dynamic As String = "Dynamic"        ' Not used in DB Schema Version 2
Const VC_Static As String = "Static"          ' Not used in DB Schema Version 2
Const VC_And As String = "And"                ' Not used in DB Schema Version 2
Const VC_Or As String = "Or"                  ' Not used in DB Schema Version 2
Const VC_Not As String = "Not"
Const VC_Any As String = "Any"

Const l_Is As Integer = 0
Const l_IsNot As Integer = 1
Const l_Any As Integer = 2

Private Enum nvtNetValueTypeConstants
    nvtGANET = 0
    nvtPNET = 1
    nvtTheoreticalNET = 2
End Enum

Dim Op As String
Dim SIsIsNotAny As Long
Dim DIsIsNotAny As Long

Public MyStuff As Collection
Public MyConnString As String       'connection string for database
Public MyCancel As Boolean

Public MTSMasterConnString As String  ' Connection string to MTS_Master

Private DBSchemaVersion As Single

Dim MTSubsetID As String            'MT subset ID                       ' Not used in DB Schema Version 2
Dim IncList As String               'delimited inclusion list

Dim UseAllSubMT As Boolean      'control variables                      ' Not used in DB Schema Version 2
Dim UseMTSubset As Boolean
                
Dim MTSubCnt As Long                            ' Not used in DB Schema Version 2
Dim MTSubID() As Long                           ' Not used in DB Schema Version 2
Dim MTSubName() As String
Dim MTSubDesc() As String                       ' Not used in DB Schema Version 2

'dynamic modifications arrays (or simply modification list in DB Schema Version 2)
Dim DGlobModCnt As Long
Dim DGlobModID() As Long
Dim DGlobModName() As String
Dim DGlobModDesc() As String

'static modifications arrays
Dim SGlobModCnt As Long
Dim SGlobModID() As Long
Dim SGlobModName() As String
Dim SGlobModDesc() As String

Dim IncItemsCnt As Long         'count of included items
Dim IncItems() As String        'contains include items

Dim bLoading As Boolean         'True until the first activation of form

Public Event DialogClosed()     'public event raised when this dialog is closed


Private Sub InitDBSchemaVersion()
    ' Query the database to determine the schema version
    
    Dim spGetDBSchemaVersion As String
    
    On Error Resume Next
    spGetDBSchemaVersion = MyStuff.Item(INIT_Get_DB_Schema_Version).Value
    If Err Then
        spGetDBSchemaVersion = "GetDBSchemaVersion"
        Err.Clear
    End If
    
    DBSchemaVersion = GetDBSchemaVersion(MyConnString, spGetDBSchemaVersion)

    If DBSchemaVersion = 0 Then
        DBSchemaVersion = 1
    End If
    
End Sub

Private Sub InitMTSubsets()
'----------------------------------------------
'loads list of subsets for selected Mass Tag db
'to combo box and initializes selection of it
'----------------------------------------------
Dim Res As Long
Dim i As Long
On Error GoTo exit_NotEndOfTheWorld

cmbMTSubset.Clear
lblMTSubsetDescription.Caption = ""

If DBSchemaVersion < 2 Then
    Res = GetMTSubsets(MyConnString, MyStuff.Item(INIT_Fill_MTSubsets).Value, _
                       MTSubID(), MTSubName(), MTSubDesc())
    If Res = 0 Then
       MTSubCnt = UBound(MTSubID) + 1
       If MTSubCnt > 0 Then
          For i = 0 To MTSubCnt - 1
              cmbMTSubset.AddItem MTSubName(i), i
          Next i
          'resolve current MTSubset situation(see if subset specified already)
          If ResolveCurrSubSetting() Then Exit Sub  'everything fine with subsets
       End If
    End If
End If

exit_NotEndOfTheWorld:
UseMTSubset = False
chkUseMTSubset.Value = vbUnchecked
chkUseMTSubset.Enabled = False
cmbMTSubset.Enabled = False

ShowHideControls
   

End Sub

Private Sub InitInternalStandardExplicit()

    Dim intInternalStandardCount As Integer
    Dim strInternalStandardNames() As String
    Dim strViewName As String
    
    Dim intIndex As Integer
    Dim lngReturn As Integer
    
    On Error Resume Next
    strViewName = MyStuff.Item(INIT_Get_Internal_Standard_Names).Value
    If Err Then
        strViewName = "V_Internal_Standards"
        Err.Clear
    End If
  
On Error GoTo InitInternalStandardExplicitErrorHandler

    lngReturn = GetInternalStandardNames(MTSMasterConnString, strViewName, intInternalStandardCount, strInternalStandardNames)

    With cboInternalStandardExplicit
        .Clear
        .AddItem ""
        
        If lngReturn = 0 Then
            For intIndex = 0 To intInternalStandardCount - 1
                .AddItem strInternalStandardNames(intIndex)
            Next intIndex
        Else
            .AddItem "PepChromeA"
        End If
        .ListIndex = 0
    End With
    Exit Sub

InitInternalStandardExplicitErrorHandler:
    Debug.Assert False
    Resume Next
    
End Sub

Private Sub InitNETValueType()

    Const NET_VALUE_TYPE_DESC_GANET = "Average NET - from DB"
    Const NET_VALUE_TYPE_DESC_PNET = "PNET - from DB"
    Const NET_VALUE_TYPE_DESC_THEORETICAL_NET = "Theoretical NET"
   
    With cboNETValueType
       .Clear
       .AddItem NET_VALUE_TYPE_DESC_GANET, nvtGANET
       .AddItem NET_VALUE_TYPE_DESC_PNET, nvtPNET
       .AddItem NET_VALUE_TYPE_DESC_THEORETICAL_NET, nvtTheoreticalNET
       .ListIndex = nvtGANET
    End With

End Sub

Private Sub InitInclusionList()
'---------------------------------------------------------
'loads list of global modifications that will be used to
'fill list of possible choices to be included in MT search
'---------------------------------------------------------
Dim Res As Long
Dim i As Long
Dim strViewName As String

On Error GoTo exit_NotTheEndOfTheWorld1

cmbDMod.Clear
cmbSMod.Clear
lstIncList.Clear

If DBSchemaVersion < 2 Then
    fraD.Caption = "Dynamic Modifications"
    
    strViewName = MyStuff.Item(INIT_Get_Global_Mod_D).Value
    If Err Then
        strViewName = "v_Global_Mod_D"
        Err.Clear
    End If
    
    'retrieve dynamic modifications and fill list box
    Res = GetGlobMods(MyConnString, strViewName, _
                      DGlobModID(), DGlobModName(), DGlobModDesc())
Else
    fraD.Caption = "Mass Correction Factors"
    
    strViewName = MyStuff.Item(INIT_Get_Global_Mass_Correction_Factors).Value
    If Err Then
        strViewName = "V_DMS_Mass_Correction_Factors"
        Err.Clear
    End If
    
    'retrieve dynamic modifications and fill list box
    Res = GetGlobMods(MTSMasterConnString, strViewName, _
                      DGlobModID(), DGlobModName(), DGlobModDesc())
End If

If Res = 0 Then
   DGlobModCnt = UBound(DGlobModID) + 1
   If DGlobModCnt > 0 Then
      For i = 0 To DGlobModCnt - 1
          cmbDMod.AddItem DGlobModName(i), i
      Next i
      cmbDMod.ListIndex = 0
   End If
End If

If DBSchemaVersion < 2 Then
    strViewName = MyStuff.Item(INIT_Get_Global_Mod_S).Value
    If Err Then
        strViewName = "v_Global_Mod_S"
        Err.Clear
    End If
    
    'retrieve static modifications and fill list box
    Res = GetGlobMods(MyConnString, strViewName, _
                      SGlobModID(), SGlobModName(), SGlobModDesc())
    If Res = 0 Then
       SGlobModCnt = UBound(SGlobModID) + 1
       If SGlobModCnt > 0 Then
          For i = 0 To SGlobModCnt - 1
              cmbSMod.AddItem SGlobModName(i), i
          Next i
          cmbSMod.ListIndex = 0
       End If
    End If
End If

ShowHideControls

DIsIsNotAny = 0
SIsIsNotAny = 0

optDIsIsNotAny(DIsIsNotAny).Value = True
optSIsIsNotAny(SIsIsNotAny).Value = True

If ResolveCurrIncSetting() Then Exit Sub
exit_NotTheEndOfTheWorld1:
End Sub

Private Sub ShowHideControls()

    Dim blnShowlegacyControls As Boolean
    
    If DBSchemaVersion < 2 Then
        blnShowlegacyControls = True
    Else
        blnShowlegacyControls = False
    End If
    
    chkAccurateOnly.Visible = blnShowlegacyControls
    chkLockersOnly.Visible = blnShowlegacyControls
    chkLimitToPMTsFromDataset.Visible = Not blnShowlegacyControls
    
    cmbMTSubset.Visible = blnShowlegacyControls
    lblMTSubsetDescription.Visible = blnShowlegacyControls
    chkUseMTSubset.Visible = blnShowlegacyControls
    chkUseAll.Visible = blnShowlegacyControls

    fraS.Visible = blnShowlegacyControls

    txtMinimumHighDiscriminantScore.Enabled = Not blnShowlegacyControls
    lblMinimumHighDiscriminantScore.Enabled = Not blnShowlegacyControls

    txtExperimentInclusionFilter.Enabled = Not blnShowlegacyControls
    txtExperimentExclusionFilter.Enabled = Not blnShowlegacyControls
    cboInternalStandardExplicit.Enabled = Not blnShowlegacyControls
    
    optAndOr(0).Visible = blnShowlegacyControls
    optAndOr(1).Visible = blnShowlegacyControls

    If blnShowlegacyControls Then
        fraSettings1.Height = 1335
        fraSettings2.Height = 1500
        fraSettings3.Height = 3255
    Else
        fraSettings1.Height = 375
        fraSettings2.Height = 2535
        fraSettings3.Height = 2200
    End If
    
    fraSettings2.Top = fraSettings1.Top + fraSettings1.Height
    fraSettings3.Top = fraSettings2.Top + fraSettings2.Height + 100
    
    lstIncList.Height = fraSettings3.Height - lstIncList.Top - 510
    cmdCancel.Top = lstIncList.Top + lstIncList.Height + 90
    cmdOK.Top = cmdCancel.Top
    
    Me.Height = fraSettings3.Top + fraSettings3.Height + 650
End Sub

Private Sub chkUseAll_Click()
If chkUseAll.Value = vbChecked Then
   UseAllSubMT = True
Else
   UseAllSubMT = False
End If
EnableDisableInclusionListControls
End Sub

Private Sub chkUseMTSubset_Click()
If chkUseMTSubset.Value = vbChecked Then
   UseMTSubset = True
Else
   UseMTSubset = False
End If
End Sub

Private Sub cmbMTSubset_Click()
Dim Ind As Long
On Error Resume Next
Ind = cmbMTSubset.ListIndex
If Ind >= 0 Then
   MTSubsetID = CStr(MTSubID(Ind))
   lblMTSubsetDescription.Caption = MTSubDesc(Ind)
End If
End Sub

Private Sub cmdCancel_Click()
MyCancel = True
Me.Hide
RaiseEvent DialogClosed
End Sub

Private Sub cmdClearInclude_Click()
'-----------------------------------------
'deletes all items from the inclusion list
'-----------------------------------------
lstIncList.Clear
Erase IncItems
IncItemsCnt = 0
End Sub

Private Sub cmdIncItemDelete_Click()
With lstIncList
    If .ListIndex >= 0 Then
       DeleteIncItemInd .ListIndex
       .RemoveItem .ListIndex
    Else
       MsgBox "No list item selected!", vbOKOnly, App.Title
    End If
End With
End Sub

Private Sub cmdInclude_Click()
Dim lstIncItem As String    'inclusion list item that goes to user interface
Dim actIncItem As String    'actual inclusion list item
Dim DName As String
Dim DID As Long
Dim SName As String
Dim SID As Long
Dim lDPart As String        ' dynamic part that goes to user interface
Dim aDPart As String        ' actual dynamic part
Dim lSPart As String
Dim aSPart As String

On Error GoTo err_cmdInclude_Click

If DBSchemaVersion < 2 Then
    'resolve dynamic modifications selection
    DName = cmbDMod.Text
    DID = GetIDFromNameD(DName)
    Select Case DIsIsNotAny
    Case l_Is
         If DID >= 0 Then
            lDPart = VC_Dynamic & Chr$(32) & DName
            aDPart = VC_Dynamic & Chr$(32) & DID
         Else
            GoTo err_cmdInclude_Click
         End If
    Case l_IsNot
         If DID >= 0 Then
            lDPart = VC_Dynamic & Chr$(32) & VC_Not & Chr$(32) & DName
            aDPart = VC_Dynamic & Chr$(32) & VC_Not & Chr$(32) & DID
         Else
            GoTo err_cmdInclude_Click
         End If
    Case l_Any
         lDPart = VC_Dynamic & Chr$(32) & VC_Any
         aDPart = VC_Dynamic & Chr$(32) & VC_Any
    End Select
    'resolve static modifications selection
    SName = cmbSMod.Text
    SID = GetIDFromNameS(SName)
    Select Case SIsIsNotAny
    Case l_Is
         If SID >= 0 Then
            lSPart = VC_Static & Chr$(32) & SName
            aSPart = VC_Static & Chr$(32) & SID
         Else
            GoTo err_cmdInclude_Click
         End If
    Case l_IsNot
         If SID >= 0 Then
            lSPart = VC_Static & Chr$(32) & VC_Not & Chr$(32) & SName
            aSPart = VC_Static & Chr$(32) & VC_Not & Chr$(32) & SID
         Else
            GoTo err_cmdInclude_Click
         End If
    Case l_Any
         lSPart = VC_Static & Chr$(32) & VC_Any
         aSPart = VC_Static & Chr$(32) & VC_Any
    End Select
    
    
    actIncItem = aDPart & Chr$(32) & Op & Chr$(32) & aSPart
    lstIncItem = lDPart & Chr$(32) & Op & Chr$(32) & lSPart
Else

    'resolve dynamic modifications selection
    DName = cmbDMod.Text
    DID = GetIDFromNameD(DName)
    Select Case DIsIsNotAny
    Case l_Is
         If DID >= 0 Then
            lDPart = DName
            aDPart = DID
         Else
            GoTo err_cmdInclude_Click
         End If
    Case l_IsNot
         If DID >= 0 Then
            lDPart = VC_Not & Chr$(32) & DName
            aDPart = VC_Not & Chr$(32) & DID
         Else
            GoTo err_cmdInclude_Click
         End If
    Case l_Any
         lDPart = VC_Any
         aDPart = VC_Any
    End Select

    actIncItem = aDPart
    lstIncItem = lDPart

End If

IncItemsCnt = IncItemsCnt + 1
ReDim Preserve IncItems(IncItemsCnt - 1)
IncItems(IncItemsCnt - 1) = actIncItem
lstIncList.AddItem lstIncItem, IncItemsCnt - 1
Exit Sub

err_cmdInclude_Click:
MsgBox "Inclusion list could not accept selected item!", vbOKOnly, App.Title
End Sub

Private Sub cmdOK_Click()
'-----------------------------------------------
'accept settings
'-----------------------------------------------
Dim i As Long
Dim Res As Long
On Error Resume Next
IncList = ""

If Not UseAllSubMT Or DBSchemaVersion >= 2 Then
    If IncItemsCnt > 0 Then
       For i = 0 To IncItemsCnt - 1
           IncList = IncList & IncItems(i) & ";"
       Next i
       If Len(IncList) > 0 Then IncList = Left$(IncList, Len(IncList) - 1)
    Else
        If DBSchemaVersion < 2 Then
            MsgBox "The inclusion list is empty.  Thus, will use all mass tags matching the other criteria, regardless of their modification status."
            IncList = "-1"
        Else
            IncList = ""
        End If
    End If
End If

If DBSchemaVersion < 2 Then
    If UseMTSubset Then
       If IsNumeric(MTSubsetID) Then
          EditAddName NAME_SUBSET, MTSubsetID
       Else
          EditAddName NAME_SUBSET, "-1"
       End If
    Else
       EditAddName NAME_SUBSET, "-1"
    End If

    If UseAllSubMT Then
       EditAddName NAME_INC_LIST, "-1"
    Else
       EditAddName NAME_INC_LIST, IncList
    End If
Else
    EditAddName NAME_INC_LIST, IncList
End If

If chkConfirmedOnly.Value = vbChecked Then
   EditAddName NAME_CONFIRMED_ONLY, VAL_TRUE
Else
   EditAddName NAME_CONFIRMED_ONLY, VAL_FALSE
End If

If DBSchemaVersion < 2 Then
    If chkAccurateOnly.Value = vbChecked Then
       EditAddName NAME_ACCURATE_ONLY, VAL_TRUE
    Else
       EditAddName NAME_ACCURATE_ONLY, VAL_FALSE
    End If
    
    If chkLockersOnly.Value = vbChecked Then
       EditAddName NAME_LOCKERS_ONLY, VAL_TRUE
    Else
       EditAddName NAME_LOCKERS_ONLY, VAL_FALSE
    End If
Else
    If chkLimitToPMTsFromDataset.Value = vbChecked Then
        EditAddName NAME_LIMIT_TO_PMTS_FROM_DATASET, VAL_TRUE
    Else
        EditAddName NAME_LIMIT_TO_PMTS_FROM_DATASET, VAL_FALSE
    End If
End If

EditAddName NAME_MINIMUM_HIGH_NORMALIZED_SCORE, txtMinimumHighNormalizedScore
EditAddName NAME_MINIMUM_HIGH_DISCRIMINANT_SCORE, txtMinimumHighDiscriminantScore
EditAddName NAME_MINIMUM_PMT_QUALITY_SCORE, txtMinimumPMTQualityScore

EditAddName NAME_EXPERIMENT_INCLUSION_FILTER, txtExperimentInclusionFilter
EditAddName NAME_EXPERIMENT_EXCLUSION_FILTER, txtExperimentExclusionFilter

If cboInternalStandardExplicit.ListIndex >= 0 Then
    EditAddName NAME_INTERNAL_STD_EXPLICIT, cboInternalStandardExplicit.List(cboInternalStandardExplicit.ListIndex)
Else
    EditAddName NAME_INTERNAL_STD_EXPLICIT, ""
End If

EditAddName NAME_NET_VALUE_TYPE, cboNETValueType.ListIndex

Me.Hide
RaiseEvent DialogClosed
End Sub

Private Sub EnableDisableInclusionListControls()
    Dim intIndex As Integer
    Dim blnEnable As Boolean
    
    If DBSchemaVersion < 2 Then
        blnEnable = Not UseAllSubMT
        
        cmdInclude.Enabled = blnEnable
        cmbDMod.Enabled = blnEnable
        cmbSMod.Enabled = blnEnable
        
        cmdIncItemDelete.Enabled = blnEnable
        cmdClearInclude.Enabled = blnEnable
        
        For intIndex = 0 To 2
            optDIsIsNotAny(intIndex).Enabled = blnEnable
            optSIsIsNotAny(intIndex).Enabled = blnEnable
        Next intIndex
    Else
        cmdInclude.Enabled = True
        cmbDMod.Enabled = True
        
        cmdIncItemDelete.Enabled = True
        cmdClearInclude.Enabled = True
        
        For intIndex = 0 To 2
            optDIsIsNotAny(intIndex).Enabled = True
        Next intIndex
        
    End If
    
End Sub

Private Sub EditAddName(ByVal PairName As String, ByVal NewValue As String)
'-------------------------------------------------------------------------
'modifies value of name value pair; if pair does not exist adds it
'-------------------------------------------------------------------------
Dim nv As NameValue
On Error Resume Next
MyStuff.Item(PairName).Value = NewValue
If Err Then
   Set nv = New NameValue
   nv.Name = PairName
   nv.Value = NewValue
   MyStuff.Add nv, nv.Name
End If
End Sub

Private Function ResolveCurrSubSetting() As Boolean
'---------------------------------------------------------------
'resolves current subset settings; returns True if everything OK
'False if not and subseting discrimination should be ignored
'---------------------------------------------------------------
Dim lMTSubID As Long
Dim lMTSubIDInd As Long
Dim i As Long
On Error Resume Next
MTSubsetID = MyStuff.Item(NAME_SUBSET).Value
If Err Then
   MTSubsetID = ""     'assume there was no MT subset name/value in MyStuff
Else
   If IsNumeric(MTSubsetID) Then
      lMTSubID = CLng(MTSubsetID)
      If lMTSubID >= 0 Then         'actual subset; try to find it among subsets
         lMTSubIDInd = -1
         For i = 0 To MTSubCnt - 1
             If MTSubID(i) = lMTSubID Then
                lMTSubIDInd = i
                Exit For
             End If
         Next i
         If lMTSubIDInd >= 0 Then   'everything is cool
            UseMTSubset = True
            chkUseMTSubset.Value = vbChecked
            cmbMTSubset.ListIndex = lMTSubIDInd
         Else                       'subset not found; play it safe and say ignore subsets
            Exit Function           'will return false
         End If
      Else                          'don't discriminate on MT subsets
         UseMTSubset = False
         chkUseMTSubset.Value = vbUnchecked
      End If
   Else                'must be an error(or empty name which comes to same)
      MTSubsetID = ""
   End If
End If
ResolveCurrSubSetting = True
End Function

Private Function ResolveCurrIncSetting() As Boolean
'-----------------------------------------------------------------------
'resolves current inclusion list settings; returns True if everything OK
'False if not and inclusion list should be ignored(everything included)
'Bad items in the list are ignored but does not cause False return
'-----------------------------------------------------------------------
Dim CurrItem() As String
Dim CurrItemCnt As Long
Dim i As Long, j As Long
Dim OpPos As Long      'position of operator
Dim DIDInd As Long     'index of dynamic ID
Dim DName As String
Dim SName As String
Dim SIDInd As Long     'index of static ID
Dim DPart As String    'dynamic part for list
Dim SPart As String    'static part for list
Dim lstCurrItem As String
Dim bIsNot As Boolean

On Error Resume Next

IncList = MyStuff.Item(NAME_INC_LIST).Value
If Err Then         'assume this is the first access
   IncList = ""
Else
   If Len(IncList) > 0 Then
        If DBSchemaVersion < 2 Then
            
            If (InStr(UCase(IncList), "DYNAMIC") = 0 Or InStr(UCase(IncList), "STATIC") = 0) And IncList <> "-1" Then
                ' new style list; reset to blank
                IncList = ""
                lstIncList.Clear
            End If
            
            If IsNumeric(IncList) Then
               If CLng(IncList) = -1 Then             'use all mass tags regardless
                  chkUseAll.Value = vbChecked         'of global mods; this is enough
                  ResolveCurrIncSetting = True        'since this option can be marked
                  Exit Function                       'on two different ways
               End If
            End If
            IncItems = Split(IncList, ";")
            IncItemsCnt = UBound(IncItems) + 1
            
            For i = 0 To IncItemsCnt - 1
                DPart = ""
                DName = ""
                SPart = ""
                SName = ""
                CurrItem = Split(IncItems(i), Chr$(32))
                CurrItemCnt = UBound(CurrItem) + 1
                SIDInd = CurrItemCnt - 1              'static always last
                DIDInd = -1                           'dynamic one before operator
                For j = 0 To CurrItemCnt - 1
                    If ((CurrItem(j) = VC_And) Or (CurrItem(j) = VC_Or)) Then
                       OpPos = j
                       DIDInd = j - 1
                       Exit For
                    End If
                Next j
                'now explore is numeric and what exactly does it say
                'if it is not clear just keep whatever it says
                DName = GetNameFromIDD(CLng(CurrItem(DIDInd)))
                If Len(DName) > 0 Then
                   For j = 0 To DIDInd - 1
                       DPart = DPart & CurrItem(j) & Chr$(32)
                   Next j
                   DPart = DPart & DName
                Else          'leave whatever was there
                   For j = 0 To DIDInd
                       DPart = DPart & CurrItem(j) & Chr$(32)
                   Next j
                   DPart = Left$(DPart, Len(DPart) - 1)
                End If
                
                SName = GetNameFromIDS(CLng(CurrItem(SIDInd)))
                If Len(SName) > 0 Then
                   For j = OpPos + 1 To CurrItemCnt - 2
                       SPart = SPart & CurrItem(j) & Chr$(32)
                   Next j
                   SPart = SPart & SName
                Else          'leave whatever was there
                   For j = OpPos + 1 To CurrItemCnt - 1
                       SPart = SPart & CurrItem(j) & Chr$(32)
                   Next j
                End If
                'assemble list version and add to the list
                lstCurrItem = DPart & Chr$(32) & CurrItem(OpPos) & Chr$(32) & SPart
                lstIncList.AddItem lstCurrItem, i
            Next i
        Else
            ' In DB Schema Version 2, each item should be a single mod ID number, or Not Mod ID, or Any
            ' For example: 1014
            '          or: Not 1014
            '          or: Any
            '
            ' We need to parse this list to find the numbers and replace them with human-readable mod names
            
            If InStr(UCase(IncList), "DYNAMIC") > 0 Or InStr(UCase(IncList), "STATIC") > 0 Or IncList = "-1" Then
                ' Old style list; reset to blank
                IncList = ""
                lstIncList.Clear
            Else
                IncItems = Split(IncList, ";")
                IncItemsCnt = UBound(IncItems) + 1
                
                For i = 0 To IncItemsCnt - 1
                    DPart = ""
                    DName = ""
                    CurrItem = Split(IncItems(i), Chr$(32))
                    CurrItemCnt = UBound(CurrItem) + 1
                    DIDInd = -1
                    
                    If CurrItem(0) = VC_Not And CurrItemCnt > 1 Then
                        bIsNot = True
                        CurrItem(0) = CurrItem(1)
                    Else
                        bIsNot = False
                    End If
                    
                    If IsNumeric(CurrItem(0)) Then
                        DName = GetNameFromIDD(CLng(CurrItem(0)))
                        If Len(DName) > 0 Then
                            DPart = DName
                        Else
                            DPart = CurrItem(0)
                        End If
                    Else
                        ' Leave whatever was there
                        DPart = CurrItem(0)
                        For j = 1 To CurrItemCnt - 1
                            DPart = DPart & Chr$(32) & CurrItem(j)
                        Next j
                    End If
                    
                    'assemble list version and add to the list
                    If bIsNot Then
                        DPart = VC_Not & Chr$(32) & DPart
                    End If
                    
                    lstCurrItem = DPart
                    lstIncList.AddItem lstCurrItem, i
                Next i
            End If
        
        End If
   End If
End If
ResolveCurrIncSetting = True
End Function

Private Sub Form_Activate()
DoEvents
If bLoading Then
   Me.MousePointer = vbHourglass
   InitDBSchemaVersion
   InitMTSubsets
   InitNETValueType
   InitInternalStandardExplicit
   InitInclusionList
   InitConfirmedAccurateLockersOnly
   Me.MousePointer = vbDefault
   bLoading = False
   
   EnableDisableInclusionListControls
End If
End Sub

Private Sub Form_Load()
bLoading = True
'default settings
Op = VC_And
DIsIsNotAny = l_Any
SIsIsNotAny = l_Any
Me.Visible = True
Me.Show
End Sub

Private Sub Form_LostFocus()
'--------------------------------------------
'same effect as if user pressed cancel button
'--------------------------------------------
'Call cmdCancel_Click
End Sub

Private Sub optAndOr_Click(Index As Integer)
Select Case Index
Case 0
     Op = VC_And
Case 1
     Op = VC_Or
End Select
End Sub

Private Sub optDIsIsNotAny_Click(Index As Integer)
DIsIsNotAny = Index
End Sub

Private Sub optSIsIsNotAny_Click(Index As Integer)
SIsIsNotAny = Index
End Sub

Public Function GetIDFromNameD(ByVal DName As String) As Long
'------------------------------------------------------------
'returns ID of dynamic modification for specified name
'-1 on any error or if not found
'------------------------------------------------------------
Dim i As Long
On Error Resume Next
GetIDFromNameD = -1
For i = 0 To DGlobModCnt - 1
    If DGlobModName(i) = DName Then
       GetIDFromNameD = DGlobModID(i)
       Exit Function
    End If
Next i
End Function

Public Function GetNameFromIDD(ByVal DID As Long) As String
'------------------------------------------------------------
'returns name of dynamic modification for specified ID
'empty string if err or not found
'------------------------------------------------------------
Dim i As Long
On Error Resume Next
For i = 0 To DGlobModCnt - 1
    If DGlobModID(i) = DID Then
       GetNameFromIDD = DGlobModName(i)
       Exit Function
    End If
Next i
End Function

Public Function GetIDFromNameS(ByVal SName As String) As Long
'------------------------------------------------------------
'returns ID of static modification for specified name
'-1 on any error or if not found
'------------------------------------------------------------
Dim i As Long
On Error Resume Next
GetIDFromNameS = -1
For i = 0 To SGlobModCnt - 1
    If SGlobModName(i) = SName Then
       GetIDFromNameS = SGlobModID(i)
       Exit Function
    End If
Next i
End Function

Public Function GetNameFromIDS(ByVal SID As Long) As String
'------------------------------------------------------------
'returns name of static modification for specified ID
'empty string if err or not found
'------------------------------------------------------------
Dim i As Long
On Error Resume Next
For i = 0 To SGlobModCnt - 1
    If SGlobModID(i) = SID Then
       GetNameFromIDS = SGlobModName(i)
       Exit Function
    End If
Next i
End Function


Private Sub DeleteIncItemInd(ByVal Ind As Long)
'----------------------------------------------
'deletes inclusion list item on position Ind
'----------------------------------------------
Dim i As Long
On Error Resume Next
If Ind < 0 Or Ind > IncItemsCnt - 1 Then Exit Sub
If Ind < IncItemsCnt - 1 Then       'not the last element
   For i = Ind To IncItemsCnt - 1
       IncItems(i) = IncItems(i + 1)
   Next i
End If
IncItemsCnt = IncItemsCnt - 1
ReDim Preserve IncItems(IncItemsCnt - 1)
End Sub

Public Sub InitConfirmedAccurateLockersOnly()
Dim TmpVal As String
Dim intIndex As Integer
On Error Resume Next
TmpVal = ""
TmpVal = MyStuff.Item(NAME_CONFIRMED_ONLY).Value
If UCase(TmpVal) = UCase(VAL_TRUE) Then
   chkConfirmedOnly.Value = vbChecked
Else
   chkConfirmedOnly.Value = vbUnchecked
End If

If DBSchemaVersion < 2 Then
    TmpVal = ""
    TmpVal = MyStuff.Item(NAME_ACCURATE_ONLY).Value
    If UCase(TmpVal) = UCase(VAL_TRUE) Then
       chkAccurateOnly.Value = vbChecked
    Else
       chkAccurateOnly.Value = vbUnchecked
    End If
    TmpVal = ""
    TmpVal = MyStuff.Item(NAME_LOCKERS_ONLY).Value
    If UCase(TmpVal) = UCase(VAL_TRUE) Then
       chkLockersOnly.Value = vbChecked
    Else
       chkLockersOnly.Value = vbUnchecked
    End If
    
    chkLimitToPMTsFromDataset.Value = vbUnchecked
Else
    chkAccurateOnly.Value = vbUnchecked
    chkLockersOnly.Value = vbUnchecked

    TmpVal = ""
    TmpVal = MyStuff.Item(NAME_LIMIT_TO_PMTS_FROM_DATASET).Value
    If UCase(TmpVal) = UCase(VAL_TRUE) Then
       chkLimitToPMTsFromDataset.Value = vbChecked
    Else
       chkLimitToPMTsFromDataset.Value = vbUnchecked
    End If

End If

ShowHideControls

TmpVal = ""
TmpVal = MyStuff.Item(NAME_MINIMUM_HIGH_NORMALIZED_SCORE).Value
If IsNumeric(TmpVal) Then
    If Val(TmpVal) < 0 Or Val(TmpVal) > 10000 Then
        TmpVal = "0"
    End If
    txtMinimumHighNormalizedScore = TmpVal
Else
    txtMinimumHighNormalizedScore = "0"
End If

TmpVal = ""
TmpVal = MyStuff.Item(NAME_MINIMUM_HIGH_DISCRIMINANT_SCORE).Value
If IsNumeric(TmpVal) Then
    If Val(TmpVal) < 0 Or Val(TmpVal) > 10000 Then
        TmpVal = "0"
    End If
    txtMinimumHighDiscriminantScore = TmpVal
Else
    txtMinimumHighDiscriminantScore = "0"
End If

TmpVal = ""
TmpVal = MyStuff.Item(NAME_MINIMUM_PMT_QUALITY_SCORE).Value
If IsNumeric(TmpVal) Then
    If Val(TmpVal) < 0 Or Val(TmpVal) > 10000 Then
        TmpVal = "0"
    End If
    txtMinimumPMTQualityScore = TmpVal
Else
    txtMinimumPMTQualityScore = "0"
End If

TmpVal = ""
TmpVal = MyStuff.Item(NAME_EXPERIMENT_INCLUSION_FILTER).Value
txtExperimentInclusionFilter = TmpVal

TmpVal = ""
TmpVal = MyStuff.Item(NAME_EXPERIMENT_EXCLUSION_FILTER).Value
txtExperimentExclusionFilter = TmpVal

TmpVal = ""
TmpVal = MyStuff.Item(NAME_INTERNAL_STD_EXPLICIT).Value
cboInternalStandardExplicit.ListIndex = 0
For intIndex = 0 To cboInternalStandardExplicit.ListCount - 1
    If LCase(TmpVal) = LCase(cboInternalStandardExplicit.List(intIndex)) Then
        cboInternalStandardExplicit.ListIndex = intIndex
        Exit For
    End If
Next intIndex



TmpVal = ""
TmpVal = MyStuff.Item(NAME_NET_VALUE_TYPE).Value
If IsNumeric(TmpVal) Then
    If Val(TmpVal) < nvtGANET Or Val(TmpVal) > nvtTheoreticalNET Then
        TmpVal = CStr(nvtGANET)
    End If
    cboNETValueType.ListIndex = Val(TmpVal)
Else
    cboNETValueType.ListIndex = nvtGANET
End If

End Sub

Private Sub txtMinimumHighDiscriminantScore_LostFocus()
  If IsNumeric(txtMinimumHighDiscriminantScore) Then
        If Val(txtMinimumHighDiscriminantScore) < 0 Then
            txtMinimumHighDiscriminantScore = "0"
        ElseIf Val(txtMinimumHighDiscriminantScore) > 1 Then
            txtMinimumHighDiscriminantScore = ".999"
        End If
    Else
        txtMinimumHighDiscriminantScore = "0"
    End If
End Sub

Private Sub txtMinimumHighNormalizedScore_LostFocus()
    If IsNumeric(txtMinimumHighNormalizedScore) Then
        If Val(txtMinimumHighNormalizedScore) < 0 Then
            txtMinimumHighNormalizedScore = "0"
        ElseIf Val(txtMinimumHighNormalizedScore) > 1000 Then
            txtMinimumHighNormalizedScore = "4"
        End If
    Else
        txtMinimumHighNormalizedScore = "0"
    End If
End Sub

Private Sub txtMinimumPMTQualityScore_LostFocus()
    If IsNumeric(txtMinimumPMTQualityScore) Then
        If Val(txtMinimumPMTQualityScore) < 0 Or Val(txtMinimumPMTQualityScore) > 1000 Then
            txtMinimumHighNormalizedScore = "0"
        End If
    Else
        txtMinimumPMTQualityScore = "0"
    End If
End Sub
