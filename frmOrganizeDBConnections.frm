VERSION 5.00
Begin VB.Form frmOrganizeDBConnections 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select/Modify Database Connection"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   9930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraLegacyDB 
      Caption         =   "Path to Legacy DB (Access DB with MT Tags)"
      Height          =   615
      Left            =   120
      TabIndex        =   56
      Top             =   7200
      Width           =   9615
      Begin VB.CommandButton cmdBrowseForLegacyDB 
         Caption         =   "B&rowse"
         Height          =   375
         Left            =   8640
         TabIndex        =   58
         Top             =   160
         Width           =   855
      End
      Begin VB.TextBox txtLegacyDBPath 
         Height          =   285
         Left            =   120
         TabIndex        =   57
         Top             =   220
         Width           =   8415
      End
   End
   Begin VB.Frame fraGelFileDBInfo 
      Caption         =   "Database info for the current gel file"
      Height          =   1095
      Left            =   120
      TabIndex        =   38
      Top             =   6000
      Width           =   9615
      Begin VB.CommandButton cmdOverrideInfoSaveChanges 
         Caption         =   "Save Job Info Changes"
         Height          =   375
         Left            =   120
         TabIndex        =   60
         Top             =   600
         Width           =   2175
      End
      Begin VB.CheckBox chkOverrideJobInfo 
         Caption         =   "Override Job Info"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   280
         Width           =   2175
      End
      Begin VB.TextBox txtMDType 
         Height          =   315
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtSourceFileName 
         Height          =   315
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   600
         Width           =   5055
      End
      Begin VB.TextBox txtJobNumber 
         Height          =   315
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblMDType 
         Caption         =   "MD_Type:"
         Height          =   255
         Left            =   6240
         TabIndex        =   41
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label lblSourceFileName 
         Caption         =   "Source Filename:"
         Height          =   255
         Left            =   2880
         TabIndex        =   43
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label lblJobNumber 
         Caption         =   "Job number:"
         Height          =   255
         Left            =   2880
         TabIndex        =   39
         Top             =   300
         Width           =   1455
      End
   End
   Begin VB.Frame fraSelectingMassTags 
      Height          =   1815
      Left            =   240
      TabIndex        =   52
      Top             =   8280
      Width           =   4695
      Begin VB.CommandButton cmdSelectingMassTagsOK 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   1920
         TabIndex        =   54
         ToolTipText     =   "Select MT Tags to load for search"
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdSelectingMassTagsCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   720
         TabIndex        =   53
         ToolTipText     =   "Select MT Tags to load for search"
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblSelectingMassTags 
         Caption         =   $"frmOrganizeDBConnections.frx":0000
         Height          =   1095
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame fraControls 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   240
      TabIndex        =   45
      Top             =   3120
      Width           =   2655
      Begin VB.ComboBox cboSortBy 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   0
         Width           =   2565
      End
      Begin VB.CommandButton cmdLinkToArbitraryDatabase 
         Caption         =   "Link to &DB Not Listed Above"
         Height          =   375
         Left            =   0
         TabIndex        =   48
         ToolTipText     =   "Link to a MT Tag database that isn't listed above"
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   0
         TabIndex        =   50
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   1200
         TabIndex        =   51
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmdBreakDatabaseLink 
         Caption         =   "&Break Current DB Link"
         Height          =   375
         Left            =   0
         TabIndex        =   49
         ToolTipText     =   "Remove the current link to a MT Tag database"
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CommandButton cmdLinkToSelectedDatabase 
         Caption         =   "&Link to Selected DB"
         Height          =   375
         Left            =   0
         TabIndex        =   47
         ToolTipText     =   "Link with the selected MT Tag database"
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame fraCurrentConnectionDetails 
      Caption         =   "Details for the currently connected database"
      Height          =   3015
      Left            =   3000
      TabIndex        =   19
      Top             =   3000
      Width           =   6735
      Begin VB.CheckBox chkCurrentDBLimitToPMTsFromDataset 
         Caption         =   "Limit to MT tags from Dataset for Job"
         Height          =   375
         Left            =   600
         TabIndex        =   29
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtCurrentDBMinimumHighDiscriminantScore 
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "0"
         Top             =   2220
         Width           =   495
      End
      Begin VB.TextBox txtCurrentDBMinimumPMTQualityScore 
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "0"
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox txtCurrentDBMinimumHighNormalizedScore 
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "0"
         Top             =   1920
         Width           =   495
      End
      Begin VB.CheckBox chkCurrentDBLockersOnly 
         Caption         =   "Lockers Only"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1640
         Width           =   1455
      End
      Begin VB.CheckBox chkCurrentDBConfirmedOnly 
         Caption         =   "Confirmed Only"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1160
         Width           =   1455
      End
      Begin VB.CommandButton cmdSelectMassTags 
         Caption         =   "&Select MT Tags"
         Height          =   375
         Left            =   4920
         TabIndex        =   25
         ToolTipText     =   "Select the MT Tags to use"
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtCurrentDBAllowedModifications 
         Height          =   615
         Left            =   3360
         TabIndex        =   30
         Top             =   1200
         Width           =   3255
      End
      Begin VB.CheckBox chkCurrentDBAMTsOnly 
         Caption         =   "AMT's Only"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1400
         Width           =   1455
      End
      Begin VB.TextBox txtCurrentDBName 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label lblCurrentDBMinimumHighDiscriminantScore 
         Caption         =   "Minimum high discriminant score:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   2240
         Width           =   2655
      End
      Begin VB.Label lblCurrentDBMinimumPMTQualityScore 
         Caption         =   "Minimum PMT quality score:"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   2540
         Width           =   2655
      End
      Begin VB.Label lblCurrentDBNETValueType 
         Caption         =   "Avg Obs NET - from DB"
         Height          =   255
         Left            =   3600
         TabIndex        =   37
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label lblCurrentDBMinimumHighNormalizedScore 
         Caption         =   "Minimum high normalized XCorr:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1940
         Width           =   2655
      End
      Begin VB.Label lblCurrentDBInternalStdExplicitOrMTSubset 
         Caption         =   "Explicit Internal Standard:"
         Height          =   225
         Left            =   120
         TabIndex        =   24
         Top             =   900
         Width           =   4695
      End
      Begin VB.Label lblCurrentDBMassTagCount 
         Caption         =   "0"
         Height          =   255
         Left            =   3480
         TabIndex        =   23
         Top             =   660
         Width           =   735
      End
      Begin VB.Label lblCurrentDBMassTagCountLabel 
         Caption         =   "Count of selected MT Tags in current DB:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   660
         Width           =   3255
      End
      Begin VB.Label lblCurrentDBName 
         Caption         =   "Database Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   300
         Width           =   1455
      End
   End
   Begin VB.Frame fraSelectedConnectionDetails 
      Caption         =   "Details for the selected connection in the list at left"
      Height          =   2895
      Left            =   4200
      TabIndex        =   1
      Top             =   0
      Width           =   5535
      Begin VB.CheckBox chkSelectedDBLimitToPMTsFromDataset 
         Caption         =   "Limit to MT tags from Dataset for Job"
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtSelectedDBMinimumHighDiscriminantScore 
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0"
         Top             =   2220
         Width           =   495
      End
      Begin VB.TextBox txtSelectedDBMinimumPMTQualityScore 
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "0"
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox txtSelectedDBMinimumHighNormalizedScore 
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "0"
         Top             =   1920
         Width           =   495
      End
      Begin VB.CheckBox chkSelectedDBLockersOnly 
         Caption         =   "Lockers Only"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1640
         Width           =   1455
      End
      Begin VB.CheckBox chkSelectedDBConfirmedOnly 
         Caption         =   "Confirmed Only"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1160
         Width           =   1455
      End
      Begin VB.TextBox txtSelectedDBAllowedModifications 
         Height          =   615
         Left            =   2160
         TabIndex        =   11
         Top             =   1200
         Width           =   3255
      End
      Begin VB.CheckBox chkSelectedDBAMTsOnly 
         Caption         =   "AMT's Only"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1400
         Width           =   1455
      End
      Begin VB.TextBox txtSelectedDBName 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label lblSelectedDBMinimumHighDiscriminantScore 
         Caption         =   "Minimum high discriminant score:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2240
         Width           =   2415
      End
      Begin VB.Label lblSelectedDBMinimumPMTQualityScore 
         Caption         =   "Minimum PMT quality score:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2540
         Width           =   2415
      End
      Begin VB.Label lblSelectedDBNETValueType 
         Caption         =   "Avg Obs NET - from DB"
         Height          =   255
         Left            =   3480
         TabIndex        =   18
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label lblSelectedDBMinimumHighNormalizedScore 
         Caption         =   "Minimum high normalized XCorr:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1940
         Width           =   2415
      End
      Begin VB.Label lblSelectedDBInternalStdExplicitOrMTSubset 
         Caption         =   "Explicit Internal Standard:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   900
         Width           =   5175
      End
      Begin VB.Label lblSelectedDBMassTagCount 
         Caption         =   "0"
         Height          =   255
         Left            =   3600
         TabIndex        =   5
         Top             =   660
         Width           =   735
      End
      Begin VB.Label lblSelectedDBMassTagCountLabel 
         Caption         =   "Count of selected MT Tags in selected DB:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   3255
      End
      Begin VB.Label lblSelectedDBName 
         Caption         =   "Database Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   1455
      End
   End
   Begin VB.ListBox lstRecentDBConnections 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmOrganizeDBConnections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This form lists the recent database connections that have been established
' The information is retrieved from the ini file
' The SelectedDB controls show information about the database that the user has
'  selected in lstRecentDBConnections
'
' The CurrentDB controls show information about the currently established connection
'  for the data file with index mCallerID

Private Enum sbSortByConstants
    sbMostRecentlyUsed = 0
    sbAlphabetical = 1
End Enum

Private mCallerID As Long        ' Contains the index of the selected file when this form was called
Private mConnectionChanged As Boolean

Private mCurrentDBSettings As udtDBSettingsType

Private mDBIndexLookupArray() As Integer            ' 0-based array
Private mRecentDBSettings() As udtDBSettingsType    ' 0-based array
Private mRecentDBCount As Integer
Private mFormLoaded As Boolean

Private mUnloadForm As Boolean
Private mLegacyDBSaved As String

Private objSelectMassTags As FTICRAnalysis

Private WithEvents objMTConnectionSelector As DummyAnalysisInitiator
Attribute objMTConnectionSelector.VB_VarHelpID = -1

Private Sub BreakMTSLink()
    With mCurrentDBSettings
        .IsDeleted = True
        .ConnectionString = ""
        .DatabaseName = ""
        ' Leave unchanged: .DBSchemaVersion
        
        .AMTsOnly = False
        .ConfirmedOnly = False
        .LockersOnly = False
        .LimitToPMTsFromDataset = False
        
        .MinimumHighNormalizedScore = 0
        .MinimumHighDiscriminantScore = 0
        .MinimumPMTQualityScore = 0
        .ExperimentInclusionFilter = ""
        .ExperimentExclusionFilter = ""
        .InternalStandardExplicit = ""
        
        .NETValueType = nvtGANET
        .MassTagSubsetID = -1
        .ModificationList = "-1"
        .SelectedMassTagCount = 0
    End With
    
    mConnectionChanged = True
    DisplayCurrentDBSettings
End Sub

Private Function ConstructInternalStdOrMTSubsetDescription(sngDBSchemaVersion As Single, strInternalStandardExplicit As String, lngMassTagSubsetID As Long) As String
    Dim strText As String
    
    If sngDBSchemaVersion >= 2 Then
        strText = "Internal Standard Explicit: " & strInternalStandardExplicit
    Else
        If lngMassTagSubsetID = -1 Then
            strText = "Mass tag subset ID: "
        Else
            strText = "Mass tag subset ID: " & Trim(lngMassTagSubsetID)
        End If
    End If

    ConstructInternalStdOrMTSubsetDescription = strText
End Function

Private Sub DisplayCurrentDBSettings()
    Dim blnShowDBSchema1Controls As Boolean
    
    With mCurrentDBSettings
        If Not .IsDeleted Then
            txtCurrentDBName = .DatabaseName
            lblCurrentDBMassTagCount = .SelectedMassTagCount
            
            lblCurrentDBInternalStdExplicitOrMTSubset = ConstructInternalStdOrMTSubsetDescription(.DBSchemaVersion, .InternalStandardExplicit, .MassTagSubsetID)
            
            SetCheckBox chkCurrentDBConfirmedOnly, .ConfirmedOnly
            SetCheckBox chkCurrentDBAMTsOnly, .AMTsOnly
            SetCheckBox chkCurrentDBLockersOnly, .LockersOnly
            SetCheckBox chkCurrentDBLimitToPMTsFromDataset, .LimitToPMTsFromDataset
            
            If .DBSchemaVersion >= 2 Then
                blnShowDBSchema1Controls = False
            Else
                blnShowDBSchema1Controls = True
            End If
            
            chkCurrentDBAMTsOnly.Visible = blnShowDBSchema1Controls
            chkCurrentDBLockersOnly.Visible = blnShowDBSchema1Controls
            chkCurrentDBLimitToPMTsFromDataset.Visible = Not blnShowDBSchema1Controls
            
            txtCurrentDBMinimumHighNormalizedScore = .MinimumHighNormalizedScore
            txtCurrentDBMinimumHighDiscriminantScore = .MinimumHighDiscriminantScore
            txtCurrentDBMinimumPMTQualityScore = .MinimumPMTQualityScore
            
            lblCurrentDBNETValueType.Caption = LookupNETValueTypeDescription(CInt(.NETValueType))
            If .ModificationList = "-1" Then
                txtCurrentDBAllowedModifications = ""
            Else
                txtCurrentDBAllowedModifications = .ModificationList
            End If
        Else
            txtCurrentDBName = ""
            lblCurrentDBMassTagCount = ""
            lblCurrentDBInternalStdExplicitOrMTSubset = ""
            chkCurrentDBConfirmedOnly = vbUnchecked
            chkCurrentDBAMTsOnly = vbUnchecked
            chkCurrentDBLockersOnly = vbUnchecked
            chkCurrentDBLimitToPMTsFromDataset = vbUnchecked
            
            txtCurrentDBMinimumHighNormalizedScore = "0"
            txtCurrentDBMinimumHighDiscriminantScore = "0"
            txtCurrentDBMinimumPMTQualityScore = "0"
            
            lblCurrentDBNETValueType.Caption = LookupNETValueTypeDescription(nvtGANET)
            txtCurrentDBAllowedModifications = ""
        End If
    End With

End Sub

Private Sub DisplayInfoOnSelectedItem()
    Dim blnShowDBSchema1Controls As Boolean
    
    If lstRecentDBConnections.ListIndex >= 0 Then
        With mRecentDBSettings(GetIndexOfSelectedDBConnection())
            txtSelectedDBName = .DatabaseName
            lblSelectedDBMassTagCount = .SelectedMassTagCount
            
            lblSelectedDBInternalStdExplicitOrMTSubset = ConstructInternalStdOrMTSubsetDescription(.DBSchemaVersion, .InternalStandardExplicit, .MassTagSubsetID)
            
            SetCheckBox chkSelectedDBConfirmedOnly, .ConfirmedOnly
            SetCheckBox chkSelectedDBAMTsOnly, .AMTsOnly
            SetCheckBox chkSelectedDBLockersOnly, .LockersOnly
            SetCheckBox chkSelectedDBLimitToPMTsFromDataset, .LimitToPMTsFromDataset
            
            If .DBSchemaVersion >= 2 Then
                blnShowDBSchema1Controls = False
            Else
                blnShowDBSchema1Controls = True
            End If
            
            chkSelectedDBAMTsOnly.Visible = blnShowDBSchema1Controls
            chkSelectedDBLockersOnly.Visible = blnShowDBSchema1Controls
            chkSelectedDBLimitToPMTsFromDataset.Visible = Not blnShowDBSchema1Controls
            
            txtSelectedDBMinimumHighNormalizedScore = .MinimumHighNormalizedScore
            txtSelectedDBMinimumHighDiscriminantScore = .MinimumHighDiscriminantScore
            txtSelectedDBMinimumPMTQualityScore = .MinimumPMTQualityScore
            
            lblSelectedDBNETValueType.Caption = LookupNETValueTypeDescription(CInt(.NETValueType))
            If .ModificationList = "-1" Then
                txtSelectedDBAllowedModifications = ""
            Else
                txtSelectedDBAllowedModifications = .ModificationList
            End If
        End With
    Else
        txtSelectedDBName = ""
        lblSelectedDBMassTagCount = ""
        lblSelectedDBInternalStdExplicitOrMTSubset = ""
        chkSelectedDBAMTsOnly = vbUnchecked
        chkSelectedDBConfirmedOnly = vbUnchecked
        chkSelectedDBLockersOnly = vbUnchecked
        chkSelectedDBLimitToPMTsFromDataset = vbUnchecked
        
        txtSelectedDBMinimumHighNormalizedScore = "0"
        txtSelectedDBMinimumHighDiscriminantScore = "0"
        txtSelectedDBMinimumPMTQualityScore = "0"
        
        lblSelectedDBNETValueType.Caption = LookupNETValueTypeDescription(nvtGANET)
        txtSelectedDBAllowedModifications = ""
    End If

End Sub

Private Sub EnableDisableJobInfoOverride(blnEnableOverride As Boolean)

    txtJobNumber.Locked = Not blnEnableOverride
    txtMDType.Locked = Not blnEnableOverride
    txtSourceFileName.Locked = Not blnEnableOverride
    
    cmdOverrideInfoSaveChanges.Enabled = blnEnableOverride
    
End Sub

Private Function GetIndexOfSelectedDBConnection() As Integer
    If lstRecentDBConnections.ListIndex >= 0 And lstRecentDBConnections.ListIndex < mRecentDBCount Then
        GetIndexOfSelectedDBConnection = mDBIndexLookupArray(lstRecentDBConnections.ListIndex)
    Else
        GetIndexOfSelectedDBConnection = 0
    End If
End Function

Private Sub HandleMTConnectionSelectorDialogClose()
    '--------------------------------------------------
    'accept settings if new analysis is specified
    '--------------------------------------------------
        
    Dim objGelAnalysis As New FTICRAnalysis
    
    Me.MousePointer = vbHourglass

    On Error GoTo HandleMTConnectionSelectorDialogCloseErrorHandler

    If Not objMTConnectionSelector.NewAnalysis Is Nothing Then
       
        Set objGelAnalysis = objMTConnectionSelector.NewAnalysis
    
        FillDBSettingsUsingAnalysisObject mCurrentDBSettings, objGelAnalysis
    End If
    
HandleMTConnectionSelectorDialogCloseContinue:
    Set objMTConnectionSelector = Nothing
    
    cmdOK.Enabled = True
    
    ' Determine number of matching MT tags for the given settings
    mCurrentDBSettings.SelectedMassTagCount = GetMassTagMatchCount(mCurrentDBSettings, LookupCurrentJob(), Me)
    mCurrentDBSettings.IsDeleted = False
    
    ' By calling the following function, the new database settings will be stored
    '  in the ini file at position 0 (i.e. the position of the latest settings)
    ' After populating lstRecentDBConnections, we'll call LinkToSelectedDatabase to use
    '  the new settings
    IniFileUpdateRecentDatabaseConnectionInfo mCurrentDBSettings
    PopulateRecentDBList
    
    If lstRecentDBConnections.ListCount > 0 Then
        LinkToSelectedDatabase
    End If
    
    Me.MousePointer = vbDefault

    Exit Sub
    
HandleMTConnectionSelectorDialogCloseErrorHandler:
    LogErrors Err.Number, "frmOrganizeDBConnections.objMTConnectionSelector_DialogClosed", Err.Description, mCallerID
    MsgBox "Error initiating new dummy analysis.", vbOKOnly
    Resume HandleMTConnectionSelectorDialogCloseContinue

End Sub

Public Sub InitializeForm()
    ' Examine GelAnalysis(mCallerID) and fill mCurrentDBSettings
    
    fraSelectingMassTags.Visible = False
    
    mCallerID = Me.Tag
    
    Me.MousePointer = vbHourglass
    
    If mCallerID = 0 Then
        If UBound(GelBody()) > 0 Then
            mCallerID = 1
        Else
            MsgBox "Please load and plot actual experimental data, prior to defining the database connection"
            cmdCancel_Click
        End If
    ElseIf mCallerID < 1 Or mCallerID > UBound(GelBody()) Then
        MsgBox "Error, this function was called with an invalid CallerID value.  Aborting"
        cmdCancel_Click
    Else
        FillDBSettingsUsingAnalysisObject mCurrentDBSettings, GelAnalysis(mCallerID)
    End If

    Me.Visible = True
    
    If Not APP_BUILD_DISABLE_MTS Then
        mCurrentDBSettings.SelectedMassTagCount = GetMassTagMatchCount(mCurrentDBSettings, LookupCurrentJob(), Me)
    
        DisplayCurrentDBSettings
    
        PopulateRecentDBList
    End If

    If Not GelAnalysis(mCallerID) Is Nothing Then
        With GelAnalysis(mCallerID)
            txtJobNumber = Trim(.MD_Reference_Job)
            txtMDType = Trim(.MD_Type)
            txtSourceFileName = Trim(.MD_file)
        End With
    End If

    txtLegacyDBPath.Text = GelData(mCallerID).PathtoDatabase
    mLegacyDBSaved = GelData(mCallerID).PathtoDatabase
    
    Me.MousePointer = vbDefault

End Sub

Private Sub LinkToSelectedDatabase()
    Dim lngCharLoc As Long
    Dim strWork As String, strJobNumber As String
    Dim fso As New FileSystemObject
    
    If lstRecentDBConnections.ListIndex >= 0 Then
        mCurrentDBSettings = mRecentDBSettings(mDBIndexLookupArray(lstRecentDBConnections.ListIndex))
        mCurrentDBSettings.IsDeleted = False
        
        If Len(txtJobNumber) = 0 And Len(txtSourceFileName) = 0 Then
            ' Attempt to populate these two textboxes
                        
            With GelData(mCallerID)
                strWork = fso.GetFileName(GelBody(mCallerID).Caption)
                lngCharLoc = InStr(UCase(strWork), "JOB")
                If lngCharLoc > 0 Then
                    strWork = Mid(strWork, lngCharLoc + 3)
                    If Not IsNumeric(Left(strWork, 1)) Then
                        If IsNumeric(Mid(strWork, 2, 1)) Then
                            strWork = Mid(strWork, 2)
                        End If
                    End If
                    
                    strJobNumber = ""
                    Do While IsNumeric(Left(strWork, 1))
                        strJobNumber = strJobNumber & Left(strWork, 1)
                        If Len(strWork) <= 1 Then Exit Do
                        strWork = Mid(strWork, 2)
                    Loop
                    
                    txtJobNumber = strJobNumber
                Else
                    txtJobNumber = "0"
                End If
                
                ' Extract the PEK/CSV/mzXML/mzData FileName from .Filename (which actually holds the full path to the file)
                txtSourceFileName = fso.GetFileName(.FileName)
                
                txtMDType = "1"
            End With
        End If
        
        mConnectionChanged = True
        
        DisplayCurrentDBSettings
    End If

    Set fso = Nothing
End Sub

Private Function LookupCurrentJob() As Long
    On Error Resume Next
    
    If Not GelAnalysis(mCallerID) Is Nothing Then
        LookupCurrentJob = Trim(GelAnalysis(mCallerID).MD_Reference_Job)
    Else
        If IsNumeric(txtJobNumber) Then
            LookupCurrentJob = val(txtJobNumber)
        Else
            LookupCurrentJob = 0
        End If
    End If
End Function

Private Sub PopulateRecentDBList()
    
    ' Populate the list
    lstRecentDBConnections.Clear

    IniFileReadRecentDatabaseConnections mRecentDBSettings(), mRecentDBCount

    ReDim mDBIndexLookupArray(mRecentDBCount)
    
    ' Populate the list by calling SortRecentDBConnections
    SortRecentDBConnections
        
End Sub

Private Sub SaveJobInfoOverrideChanges()
    
    Dim blnValidUpdate As Boolean
    
    If Not GelAnalysis(mCallerID) Is Nothing Then
        blnValidUpdate = False
        
        With GelAnalysis(mCallerID)
            If IsNumeric(txtJobNumber) Then
                blnValidUpdate = True
                .MD_Reference_Job = val(txtJobNumber)
            End If
            
            If IsNumeric(txtMDType) Then
                blnValidUpdate = True
                .MD_Type = val(txtMDType)
            End If
            
            If blnValidUpdate Then
                txtSourceFileName = Trim(.MD_file)
                mConnectionChanged = True
            End If
        End With
        
        With GelAnalysis(mCallerID)
            txtJobNumber = Trim(.MD_Reference_Job)
            txtMDType = Trim(.MD_Type)
            txtSourceFileName = Trim(.MD_file)
        End With
    End If
    
    ' Determine number of matching MT tags for the given settings
    If mCurrentDBSettings.LimitToPMTsFromDataset Then
        mCurrentDBSettings.SelectedMassTagCount = GetMassTagMatchCount(mCurrentDBSettings, LookupCurrentJob(), Me)
        DisplayCurrentDBSettings
    End If
    
    chkOverrideJobInfo.Value = vbUnchecked
    
End Sub

Private Sub SaveNewSettings()
    
    Dim udtExistingAnalysisInfo As udtGelAnalysisInfoType
    Dim eResponse As VbMsgBoxResult
    
    Me.MousePointer = vbHourglass
    
    If mCurrentDBSettings.IsDeleted Then
        If Not GelAnalysis(mCallerID) Is Nothing Then
            AddToAnalysisHistory mCallerID, "Database definition removed (was " & ExtractDBNameFromConnectionString(GelAnalysis(mCallerID).MTDB.cn.ConnectionString) & ")"
        End If
        
        ClearGelAnalysisObject mCallerID, False
    Else
        If GelAnalysis(mCallerID) Is Nothing Then
            Set GelAnalysis(mCallerID) = New FTICRAnalysis
            udtExistingAnalysisInfo.ValidAnalysisDataPresent = False
        Else
            FillGelAnalysisInfo udtExistingAnalysisInfo, GelAnalysis(mCallerID)
        End If
        
        FillGelAnalysisObject GelAnalysis(mCallerID), mCurrentDBSettings.AnalysisInfo
        
        If udtExistingAnalysisInfo.ValidAnalysisDataPresent Then
        
            ' Update GelAnalysis() with the settings in udtAnalysisInfo
            ' However, do not update .MTDB or the DBStuff() collection since we want the settings in
            '  mCurrentDBSettings to take precedence
            FillGelAnalysisObject GelAnalysis(mCallerID), udtExistingAnalysisInfo, False, False
            
        End If
        
        If Not APP_BUILD_DISABLE_MTS Then
            IniFileUpdateRecentDatabaseConnectionInfo mCurrentDBSettings
            AddToAnalysisHistory mCallerID, "Database connection defined: " & mCurrentDBSettings.DatabaseName & "; " & mCurrentDBSettings.SelectedMassTagCount & " MT tags"
        End If
    End If
    
    ' Update the cached LegacyAMTDBPath value
    glbPreferencesExpanded.LegacyAMTDBPath = GelData(mCallerID).PathtoDatabase
    
    ' Update .MD_Reference_Job and .MD_File
    If Not GelAnalysis(mCallerID) Is Nothing Then
        With GelAnalysis(mCallerID)
            .MD_Reference_Job = CLngSafe(txtJobNumber)
            If .MD_Reference_Job < 0 Then .MD_Reference_Job = 0
            
            .MD_Type = CLngSafe(txtMDType)
            If .MD_Type = stNotDefined Then
                eResponse = MsgBox("MD_Type should normally not be 0; change it to a value of 1?", vbQuestion + vbYesNo + vbDefaultButton1, "Invalid MD_Type")
                If eResponse = vbYes Then .MD_Type = stStandardIndividual
            End If
            
            .MD_file = txtSourceFileName
        End With
    End If
    
    GelStatus(mCallerID).Dirty = True

    Me.MousePointer = vbDefault
End Sub

Private Sub SelectMassTagsForCurrentDB()
    
    If mCurrentDBSettings.IsDeleted Then
        MsgBox "Cannot select MT tags since not currently linked to a database", vbExclamation + vbOKOnly, "No connection"
        Exit Sub
    End If
    
    lblSelectingMassTags.Caption = "The MT tag selection window should now be visible.  When done selecting the MT tags, press OK on that Window.  Next, press the OK button below.  To cancel (or if you can't see the window), select Cancel."
    
    With fraSelectingMassTags
        .Left = 0
        .Top = 0
        .width = Me.ScaleWidth
        .Height = Me.ScaleHeight
        .Visible = True
        .ZOrder
    End With
    
    If objSelectMassTags Is Nothing Then
        Set objSelectMassTags = New FTICRAnalysis
    End If
        
    FillGelAnalysisObject objSelectMassTags, mCurrentDBSettings.AnalysisInfo
    
    ' Use the following to display the MT tags selection window
    ' Unfortunately, there is no way to wait for this to finish
    ' Thus the reason for fraSelectingMassTags above, which fills the window to cover the other controls,
    '  and requires the user to click OK or Cancel when done selecting MT tags
    
    objSelectMassTags.MTDB.SelectMassTags glInitFile
    
End Sub

Private Sub ShellSortDBIndexLookupArray(ByRef mDBIndexLookupArray() As Integer, ByRef mRecentDBSettings() As udtDBSettingsType, ByVal lngLowIndex As Integer, ByVal lngHighIndex As Integer)
    Dim intCount As Integer
    Dim intIncrement As Integer
    Dim intIndex As Integer
    Dim intIndexCompare As Integer
    Dim intPointerValSaved As Integer

On Error GoTo ShellSortDBIndexErrorHandler

    ' sort array[lngLowIndex..lngHighIndex]

    ' compute largest increment
    intCount = lngHighIndex - lngLowIndex + 1
    intIncrement = 1
    If (intCount < 14) Then
        intIncrement = 1
    Else
        Do While intIncrement < intCount
            intIncrement = 3 * intIncrement + 1
        Loop
        intIncrement = intIncrement \ 3
        intIncrement = intIncrement \ 3
    End If

    Do While intIncrement > 0
        ' sort by insertion in increments of intIncrement
        For intIndex = lngLowIndex + intIncrement To lngHighIndex
            intPointerValSaved = mDBIndexLookupArray(intIndex)
            For intIndexCompare = intIndex - intIncrement To lngLowIndex Step -intIncrement
                ' Use <= to sort ascending; Use > to sort descending
                If LCase(mRecentDBSettings(mDBIndexLookupArray(intIndexCompare)).DatabaseName) <= LCase(mRecentDBSettings(intPointerValSaved).DatabaseName) Then Exit For
                mDBIndexLookupArray(intIndexCompare + intIncrement) = mDBIndexLookupArray(intIndexCompare)
            Next intIndexCompare
            mDBIndexLookupArray(intIndexCompare + intIncrement) = intPointerValSaved
        Next intIndex
        intIncrement = intIncrement \ 3
    Loop
    
    Exit Sub
    
ShellSortDBIndexErrorHandler:
Debug.Assert False

End Sub

Private Sub ShowHidePNNLMenus()
    If APP_BUILD_DISABLE_MTS Then
        fraSelectedConnectionDetails.Visible = False
        fraCurrentConnectionDetails.Visible = False
        cmdSelectMassTags.Enabled = False
        lstRecentDBConnections.Visible = False
    
        cboSortBy.Visible = False
        
        cmdBreakDatabaseLink.Visible = False
        cmdLinkToSelectedDatabase.Visible = False
        cmdLinkToArbitraryDatabase.Visible = False
    
        fraControls.Top = 0
        fraGelFileDBInfo.Top = 60
        fraLegacyDB.Top = fraGelFileDBInfo.Top + fraGelFileDBInfo.Height + 120
       
        Me.Height = 3250
    End If
End Sub

Private Sub SortRecentDBConnections()
    Dim intIndex As Integer
    Dim intCurrentSelectionItemIndex As Integer
    
    If mRecentDBCount < 1 Then
        lstRecentDBConnections.Clear
        Exit Sub
    End If
    
    intCurrentSelectionItemIndex = GetIndexOfSelectedDBConnection()
    
    ' Initialize mDBIndexLookupArray()
    For intIndex = 0 To mRecentDBCount - 1
        mDBIndexLookupArray(intIndex) = intIndex
    Next intIndex
    
    ' If sorting alphabetically, then need to re-order mDBIndexLookupArray() by Database name
    If cboSortBy.ListIndex = sbAlphabetical Then
        ShellSortDBIndexLookupArray mDBIndexLookupArray(), mRecentDBSettings(), 0, mRecentDBCount - 1
    End If
    
    ' Populate the list
    lstRecentDBConnections.Clear
    
    For intIndex = 0 To mRecentDBCount - 1
        With mRecentDBSettings(mDBIndexLookupArray(intIndex))
            lstRecentDBConnections.AddItem .DatabaseName & ", " & .SelectedMassTagCount & " MT tags"
        End With
        If mDBIndexLookupArray(intIndex) = intCurrentSelectionItemIndex Then
            lstRecentDBConnections.ListIndex = intIndex
        End If
    Next intIndex
    
    If lstRecentDBConnections.ListIndex < 0 Then lstRecentDBConnections.ListIndex = 0
End Sub

Public Sub WaitUntilFormClose()

    mUnloadForm = False
    Do
        Sleep 50
        DoEvents
    Loop While Not mUnloadForm
    
End Sub

Private Sub cboSortBy_Click()
    If mFormLoaded Then SortRecentDBConnections
End Sub

Private Sub chkCurrentDBAMTsOnly_Click()
    ' Do not allow user to change this; it is for display only
    On Error Resume Next
    SetCheckBox chkCurrentDBAMTsOnly, mCurrentDBSettings.AMTsOnly
End Sub

Private Sub chkCurrentDBConfirmedOnly_Click()
    ' Do not allow user to change this; it is for display only
    On Error Resume Next
    SetCheckBox chkCurrentDBConfirmedOnly, mCurrentDBSettings.ConfirmedOnly
End Sub

Private Sub chkCurrentDBLimitToPMTsFromDataset_Click()
    ' Do not allow user to change this; it is for display only
    On Error Resume Next
    SetCheckBox chkCurrentDBLimitToPMTsFromDataset, mCurrentDBSettings.LimitToPMTsFromDataset
End Sub

Private Sub chkCurrentDBLockersOnly_Click()
    ' Do not allow user to change this; it is for display only
    On Error Resume Next
    SetCheckBox chkCurrentDBLockersOnly, mCurrentDBSettings.LockersOnly
End Sub

Private Sub chkOverrideJobInfo_Click()
    EnableDisableJobInfoOverride cChkBox(chkOverrideJobInfo)
End Sub

Private Sub chkSelectedDBAMTsOnly_Click()
    ' Do not allow user to change this; it is for display only
    On Error Resume Next
    SetCheckBox chkSelectedDBAMTsOnly, mRecentDBSettings(mDBIndexLookupArray(lstRecentDBConnections.ListIndex)).AMTsOnly
End Sub

Private Sub chkselectedDBConfirmedOnly_Click()
    ' Do not allow user to change this; it is for display only
    On Error Resume Next
    SetCheckBox chkSelectedDBConfirmedOnly, mRecentDBSettings(mDBIndexLookupArray(lstRecentDBConnections.ListIndex)).ConfirmedOnly
End Sub

Private Sub chkSelectedDBLimitToPMTsFromDataset_Click()
    ' Do not allow user to change this; it is for display only
    On Error Resume Next
    SetCheckBox chkSelectedDBLimitToPMTsFromDataset, mRecentDBSettings(mDBIndexLookupArray(lstRecentDBConnections.ListIndex)).LimitToPMTsFromDataset
End Sub

Private Sub chkselectedDBLockersOnly_Click()
    ' Do not allow user to change this; it is for display only
    On Error Resume Next
    SetCheckBox chkSelectedDBLockersOnly, mRecentDBSettings(mDBIndexLookupArray(lstRecentDBConnections.ListIndex)).LockersOnly
End Sub

Private Sub cmdBreakDatabaseLink_Click()
    BreakMTSLink
End Sub

Private Sub cmdBrowseForLegacyDB_Click()
    Dim strNewFilePath As String
    
    strNewFilePath = SelectLegacyMTDB(Me, txtLegacyDBPath.Text)
    If Len(strNewFilePath) > 0 Then
        txtLegacyDBPath.Text = strNewFilePath
    End If
End Sub

Private Sub cmdCancel_Click()

    HideMTConnectionClassForm objMTConnectionSelector
    
    GelData(mCallerID).PathtoDatabase = mLegacyDBSaved
    glbPreferencesExpanded.LegacyAMTDBPath = mLegacyDBSaved
    
    mUnloadForm = True
    
    Unload Me

End Sub

Private Sub cmdLinkToArbitraryDatabase_Click()
    On Error Resume Next
    
    cmdOK.Enabled = False
    
    Set objMTConnectionSelector = New DummyAnalysisInitiator
    objMTConnectionSelector.GetNewAnalysisDialog glInitFile

End Sub

Private Sub cmdLinkToSelectedDatabase_Click()
    LinkToSelectedDatabase
End Sub

Private Sub cmdOK_Click()
    If mConnectionChanged Then
        SaveNewSettings
    End If
    
    mUnloadForm = True
    
    Unload Me
End Sub

Private Sub cmdOverrideInfoSaveChanges_Click()
    SaveJobInfoOverrideChanges
End Sub

Private Sub cmdSelectingMassTagsCancel_Click()
    fraSelectingMassTags.Visible = False
    HideMTConnectionClassForm objMTConnectionSelector

End Sub

Private Sub cmdSelectingMassTagsOK_Click()
    fraSelectingMassTags.Visible = False

    FillDBSettingsUsingAnalysisObject mCurrentDBSettings, objSelectMassTags
    
    Me.MousePointer = vbHourglass
    cmdOK.Enabled = False
    cmdSelectMassTags.Enabled = False
    
    ' Determine number of matching MT tags for the given settings
    mCurrentDBSettings.SelectedMassTagCount = GetMassTagMatchCount(mCurrentDBSettings, LookupCurrentJob(), Me)
    
    DisplayCurrentDBSettings
    mConnectionChanged = True
    
    Me.MousePointer = vbDefault
    cmdOK.Enabled = True
    cmdSelectMassTags.Enabled = True
    
End Sub

Private Sub cmdSelectMassTags_Click()
    
    SelectMassTagsForCurrentDB
    
End Sub

Private Sub Form_Load()
    SizeAndCenterWindow Me, cWindowUpperThird, 9950, 8350, False
    
    mFormLoaded = False
    With cboSortBy
        .Clear
        .AddItem "Sort by Most Recent"
        .AddItem "Sort alphabetically"
        .ListIndex = 0
    End With
    
    chkSelectedDBLimitToPMTsFromDataset.Left = chkSelectedDBAMTsOnly.Left
    chkSelectedDBLimitToPMTsFromDataset.Visible = False
    
    chkCurrentDBLimitToPMTsFromDataset.Left = chkCurrentDBAMTsOnly.Left
    chkCurrentDBLimitToPMTsFromDataset.Visible = False
    
    chkOverrideJobInfo.Value = vbUnchecked
    EnableDisableJobInfoOverride False
    
    ShowHidePNNLMenus
    
    mFormLoaded = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mUnloadForm = True
    DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objSelectMassTags = Nothing
    Set objMTConnectionSelector = Nothing
End Sub

Private Sub lstRecentDBConnections_Click()
    DisplayInfoOnSelectedItem
End Sub

Private Sub lstRecentDBConnections_DblClick()
    DisplayInfoOnSelectedItem
    LinkToSelectedDatabase
End Sub

Private Sub objMTConnectionSelector_DialogClosed()
    HandleMTConnectionSelectorDialogClose
End Sub

Private Sub txtLegacyDBPath_Change()
    GelData(mCallerID).PathtoDatabase = txtLegacyDBPath.Text
End Sub

Private Sub txtCurrentDBAllowedModifications_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then KeyCode = 0
End Sub

Private Sub txtCurrentDBAllowedModifications_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then KeyAscii = 0
    TextBoxKeyPressHandler txtCurrentDBAllowedModifications, KeyAscii, False
End Sub

Private Sub txtCurrentDBName_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtCurrentDBName, KeyAscii, False
End Sub

Private Sub txtJobNumber_GotFocus()
    TextBoxGotFocusHandler txtJobNumber, False
End Sub

Private Sub txtJobNumber_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtJobNumber, KeyAscii, True, False
End Sub

Private Sub txtMDType_GotFocus()
    TextBoxGotFocusHandler txtMDType, False
End Sub

Private Sub txtMDType_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtMDType, KeyAscii, True, False
End Sub

Private Sub txtSelectedDBAllowedModifications_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then KeyCode = 0
End Sub

Private Sub txtSelectedDBAllowedModifications_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then KeyAscii = 0
    TextBoxKeyPressHandler txtSelectedDBAllowedModifications, KeyAscii, False
End Sub

Private Sub txtSelectedDBName_KeyPress(KeyAscii As Integer)
    TextBoxKeyPressHandler txtSelectedDBName, KeyAscii, False
End Sub
