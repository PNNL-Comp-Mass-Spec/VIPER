VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmParameters 
   Caption         =   "Edit Gel Parameters"
   ClientHeight    =   5715
   ClientLeft      =   2580
   ClientTop       =   1515
   ClientWidth     =   7755
   Icon            =   "frmParameters.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   7755
   Begin TabDlg.SSTab tbsParameters 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Calibration"
      TabPicture(0)   =   "frmParameters.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraCalibration"
      Tab(0).Control(1)=   "lblRecalibration(0)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Comments"
      TabPicture(1)   =   "frmParameters.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtComment"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "General"
      TabPicture(2)   =   "frmParameters.frx":0342
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lblMTDBAssociation"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label2(3)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label2(2)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label2(0)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label2(4)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label2(5)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label2(6)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label2(1)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "cmdSelectMassTags"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "cmdBreakMTLink"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "cmdDummyMTLink"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "txtMediaType"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "txtCertificate"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "txtDFilesPath"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "txtLegacyDBPath"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "cmdBrowseForLegacyDB"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "cmdBrowseForDataFileFolder"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "cboRawDataFileFormat"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "cmdBrowseForInputFile"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "txtInputFilePath"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "cmdBrowseForFinniganRawFilePath"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "txtFinniganRawFilePath"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).ControlCount=   22
      Begin VB.TextBox txtFinniganRawFilePath 
         Height          =   285
         Left            =   1680
         TabIndex        =   22
         Top             =   1560
         Width           =   4695
      End
      Begin VB.CommandButton cmdBrowseForFinniganRawFilePath 
         Caption         =   "B&rowse"
         Height          =   375
         Left            =   6480
         TabIndex        =   23
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtInputFilePath 
         Height          =   285
         Left            =   1680
         TabIndex        =   25
         Top             =   2160
         Width           =   4695
      End
      Begin VB.CommandButton cmdBrowseForInputFile 
         Caption         =   "B&rowse"
         Height          =   375
         Left            =   6480
         TabIndex        =   26
         Top             =   2160
         Width           =   855
      End
      Begin VB.ComboBox cboRawDataFileFormat 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CommandButton cmdBrowseForDataFileFolder 
         Caption         =   "B&rowse"
         Height          =   375
         Left            =   6480
         TabIndex        =   18
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdBrowseForLegacyDB 
         Caption         =   "B&rowse"
         Height          =   375
         Left            =   6480
         TabIndex        =   29
         Top             =   2610
         Width           =   855
      End
      Begin VB.TextBox txtLegacyDBPath 
         Height          =   285
         Left            =   1680
         TabIndex        =   28
         Top             =   2640
         Width           =   4695
      End
      Begin VB.Frame fraCalibration 
         Caption         =   "Calibration Equation"
         Height          =   1905
         Left            =   -74760
         TabIndex        =   1
         Top             =   600
         Width           =   5175
         Begin VB.ComboBox cmbCalEq 
            Height          =   315
            ItemData        =   "frmParameters.frx":035E
            Left            =   240
            List            =   "frmParameters.frx":0360
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   480
            Width           =   2895
         End
         Begin VB.TextBox txtCalEqPar 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   480
            TabIndex        =   5
            Top             =   1040
            Width           =   1455
         End
         Begin VB.TextBox txtCalEqPar 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   480
            TabIndex        =   9
            Top             =   1400
            Width           =   1455
         End
         Begin VB.TextBox txtCalEqPar 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   2400
            TabIndex        =   7
            Top             =   1040
            Width           =   1455
         End
         Begin VB.TextBox txtCalEqPar 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   2400
            TabIndex        =   11
            Top             =   1400
            Width           =   1455
         End
         Begin VB.CommandButton cmdCalClear 
            Caption         =   "Cl&ear"
            Height          =   375
            Left            =   3600
            TabIndex        =   3
            Top             =   420
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "A"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   4
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "B"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   8
            Top             =   1440
            Width           =   255
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "C"
            Height          =   255
            Index           =   2
            Left            =   2160
            TabIndex        =   6
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "D"
            Height          =   255
            Index           =   3
            Left            =   2160
            TabIndex        =   10
            Top             =   1440
            Width           =   255
         End
      End
      Begin VB.TextBox txtDFilesPath 
         Height          =   285
         Left            =   1680
         TabIndex        =   17
         Top             =   840
         Width           =   4695
      End
      Begin VB.TextBox txtCertificate 
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtMediaType 
         Height          =   285
         Left            =   1680
         TabIndex        =   31
         Top             =   3120
         Width           =   2295
      End
      Begin VB.CommandButton cmdDummyMTLink 
         Caption         =   "&Link With MT Database"
         Height          =   375
         Left            =   5280
         TabIndex        =   33
         ToolTipText     =   "Link with MT Tag database"
         Top             =   3420
         Width           =   1935
      End
      Begin VB.CommandButton cmdBreakMTLink 
         Caption         =   "&Break Link With MT DB"
         Height          =   375
         Left            =   5280
         TabIndex        =   34
         Top             =   3900
         Width           =   1935
      End
      Begin VB.CommandButton cmdSelectMassTags 
         Caption         =   "&Select MT Tags"
         Height          =   375
         Left            =   5280
         TabIndex        =   35
         ToolTipText     =   "Select MT Tags to load for search"
         Top             =   4380
         Width           =   1935
      End
      Begin VB.TextBox txtComment 
         Height          =   2775
         Left            =   -74640
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   720
         Width           =   5175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Path To Finnigan .Raw file"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   21
         Top             =   1605
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Path To Input File:"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   24
         Top             =   2205
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Raw data format:"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   19
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Path to Legacy DB:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   27
         Top             =   2685
         Width           =   1455
      End
      Begin VB.Label lblRecalibration 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmParameters.frx":0362
         Height          =   735
         Index           =   0
         Left            =   -74760
         TabIndex        =   12
         Top             =   2760
         Width           =   4215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Media Type:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   30
         Top             =   3165
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Path To Data Files:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   880
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Certificate:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   520
         Width           =   1455
      End
      Begin VB.Label lblMTDBAssociation 
         BackStyle       =   0  'Transparent
         Height          =   1095
         Left            =   240
         TabIndex        =   32
         Top             =   3600
         Width           =   4575
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   38
      Top             =   5175
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   37
      Top             =   5175
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "WARNING: Editing information on this dialog might change data or access to data files."
      Height          =   495
      Left            =   120
      TabIndex        =   36
      Top             =   5160
      Width           =   3255
   End
End
Attribute VB_Name = "frmParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'last modified 11/08/2001 nt
'---------------------------------------------------
Option Explicit

Const LEGACY_DB = "Legacy MT tag database"

Dim CallerID As Long

Dim OldComment As String
Dim OldCalEq As String
Dim OldCalArg(1 To 10) As Double
Dim OldPathToDataF As String
Dim OldRawDataFileFormat As rfcRawFileConstants
Dim OldInputFilePath As String

Dim OldPathToDB As String
Dim OldCertificate As String
Dim OldMediaType As String

' MonroeMod
Dim mDBInfoModified As Boolean

'used when dummy analysis is established so that
'it can be canceled if user selects Cancel button
Dim DummyMTLinkEstablished As Boolean

Dim WithEvents MyDummyInit As DummyAnalysisInitiator
Attribute MyDummyInit.VB_VarHelpID = -1

Private Sub cboRawDataFileFormat_Click()
    GelStatus(CallerID).SourceDataRawFileType = cboRawDataFileFormat.ListIndex
End Sub

Private Sub cmbCalEq_Click()
    GelData(CallerID).CalEquation = cmbCalEq.Text
End Sub

Private Sub cmdBreakMTLink_Click()
    BreakMTSLink
End Sub

Private Sub cmdBrowseForFinniganRawFilePath_Click()
    Dim PathToInputFile As String
    
    On Error Resume Next
    PathToInputFile = SelectFile(Me.hwnd, _
                          "Select Finnigan data file", "", False, "", _
                          "RAW Files (*.raw)|*.raw|" & _
                          "All Files (*.*)|*.*", _
                          1)
                          
    If Len(PathToInputFile) > 0 Then
        txtFinniganRawFilePath.Text = PathToInputFile
        txtFinniganRawFilePath_LostFocus
    End If

End Sub

Private Sub cmdBrowseForLegacyDB_Click()
    Dim strNewFilePath As String
    
    strNewFilePath = SelectLegacyMTDB(Me, txtLegacyDBPath.Text)
    If Len(strNewFilePath) > 0 Then
        txtLegacyDBPath.Text = strNewFilePath
    End If
End Sub

Private Sub cmdBrowseForDataFileFolder_Click()
    Dim strNewFolder As String
    
    strNewFolder = BrowseForFileOrFolder(Me.hwnd, txtDFilesPath, "Select folder with original (Zipped) raw data files", True)
    If Len(strNewFolder) > 0 Then
        txtDFilesPath = strNewFolder
        txtDFilesPath_LostFocus
    End If
End Sub

Private Sub cmdBrowseForInputFile_Click()
    Dim PathToInputFile As String
    
    On Error Resume Next
    PathToInputFile = SelectFile(Me.hwnd, _
                          "Select source .Pek, .CSV, .mzXML, or .mzData file", "", False, "", _
                          "All Files (*.*)|*.*|" & _
                          "PEK Files (*.pek)|*.pek|" & _
                          "CSV Files (*.csv)|*.csv|" & _
                          "mzXML Files (*.mzXML)|*.mzXml|" & _
                          "mzXML Files (*mzXML.xml)|*mzXML.xml|" & _
                          "mzData Files (*.mzData)|*.mzData|" & _
                          "mzData Files (*mzData.xml)|*mzData.xml", _
                          glbPreferencesExpanded.LastInputFileMode + 2)
                          
    If Len(PathToInputFile) > 0 Then
        txtInputFilePath.Text = PathToInputFile
        txtInputFilePath_LostFocus
        UpdatePreferredFileExtension PathToInputFile
    End If

End Sub

Private Sub cmdCalClear_Click()
Dim i As Long
With GelData(CallerID)
    .CalEquation = ""
    For i = 1 To 10
        .CalArg(i) = 0
    Next i
    FillCalEqCombo
    For i = 0 To 3
        txtCalEqPar(i).Text = 0
    Next i
End With
End Sub

Private Sub cmdCancel_Click()

' MonroeMod Begin
    HideMTConnectionClassForm MyDummyInit
' MonroeMod Finish

RestoreOldSettings
Unload Me
End Sub

Private Sub cmdDummyMTLink_Click()
On Error Resume Next

' MonroeMod
cmdOK.Enabled = False

Set MyDummyInit = New DummyAnalysisInitiator
MyDummyInit.GetNewAnalysisDialog glInitFile

End Sub

Private Sub cmdOK_Click()
If ParametersChange() Then GelStatus(CallerID).Dirty = True

' MonroeMod
If mDBInfoModified Then UpdateIniFileWithDBInfo

Unload Me
End Sub

Private Sub cmdSelectMassTags_Click()
If GelAnalysis(CallerID) Is Nothing Then
    WarnUserNotConnectedToDB CallerID, False
Else
   GelAnalysis(CallerID).MTDB.SelectMassTags glInitFile
   
   ' MonroeMod
   mDBInfoModified = True
End If
End Sub

Private Sub FillCalEqCombo()
    With cmbCalEq
        .Clear
        .AddItem CAL_EQUATION_1
        .AddItem CAL_EQUATION_2
        .AddItem CAL_EQUATION_3
        .AddItem CAL_EQUATION_4
        .AddItem CAL_EQUATION_5
    End With
End Sub

Private Sub Form_Activate()
    CallerID = Me.Tag

On Error GoTo FormActivateErrorHandler
    SaveOldSettings
    Settings
    
    'association with MT tags
    If GelAnalysis(CallerID) Is Nothing Then
       cmdDummyMTLink.Enabled = Not APP_BUILD_DISABLE_MTS
       lblMTDBAssociation.Caption = LEGACY_DB
    Else        'already linked; don't allow relink
       cmdDummyMTLink.Enabled = False
       lblMTDBAssociation.Caption = GelAnalysis(CallerID).MTDB.cn.ConnectionString
    End If
    
    txtLegacyDBPath.Text = GelData(CallerID).PathtoDatabase
    Exit Sub

FormActivateErrorHandler:
    Debug.Assert False
    Resume Next
End Sub

Private Sub Form_Load()
    PopulateComboBoxes

    tbsParameters.Tab = 0

    ShowHidePNNLMenus

    PositionControls
End Sub

Private Sub Form_Resize()
PositionControls
End Sub

Private Sub MyDummyInit_DialogClosed()
'--------------------------------------------------
'accept settings if new dummy analysis is specified
'--------------------------------------------------
On Error GoTo err_MyDummyInit
If Not MyDummyInit.NewAnalysis Is Nothing Then
   DummyMTLinkEstablished = True
   If GelAnalysis(CallerID) Is Nothing Then Set GelAnalysis(CallerID) = New FTICRAnalysis
   Set GelAnalysis(CallerID) = MyDummyInit.NewAnalysis
   
   'don't allow relink with some other database in this session
   cmdDummyMTLink.Enabled = False
   lblMTDBAssociation.Caption = GelAnalysis(CallerID).MTDB.cn.ConnectionString
End If

exit_MyDummyInit_DialogClosed:
Set MyDummyInit = Nothing

' MonroeMod Start
cmdOK.Enabled = True
mDBInfoModified = True
' MonroeMod Finish

Exit Sub

err_MyDummyInit:
LogErrors Err.Number, "MyDummyInit_DialogClosed"
MsgBox "Error initiating new dummy analysis.", vbOKOnly
Resume exit_MyDummyInit_DialogClosed:
End Sub

Private Sub BreakMTSLink()
    ClearGelAnalysisObject CallerID, False
    
    cmdDummyMTLink.Enabled = Not APP_BUILD_DISABLE_MTS
    lblMTDBAssociation.Caption = LEGACY_DB
End Sub

Private Sub SaveOldSettings()
    Dim ArrSize As Long
    With GelData(CallerID)
        OldCalEq = .CalEquation
        ArrSize = 10 * Len(.CalArg(1))
        CopyMemory OldCalArg(1), .CalArg(1), ArrSize
        OldComment = .Comment
        OldPathToDataF = .PathtoDataFiles
        OldRawDataFileFormat = GelStatus(CallerID).SourceDataRawFileType
        OldInputFilePath = .FileName
        OldCertificate = .Certificate
        OldMediaType = .MediaType
        OldPathToDB = .PathtoDatabase
    End With
End Sub

Private Sub RestoreOldSettings()
Dim ArrSize As Long
With GelData(CallerID)
    .Certificate = OldCertificate
    .CalEquation = OldCalEq
    ArrSize = 10 * Len(OldCalArg(1))
    CopyMemory .CalArg(1), OldCalArg(1), ArrSize
    .Comment = OldComment
    .PathtoDataFiles = OldPathToDataF
    GelStatus(CallerID).SourceDataRawFileType = OldRawDataFileFormat
    .FileName = OldInputFilePath
    .MediaType = OldMediaType
    .PathtoDatabase = OldPathToDB
End With
' MonroeMod
' Commented out the following since I don't want GelAnalysis(CallerID) to
'  be deallocated when the user clicks Cancel
'If DummyMTLinkEstablished Then
'   Set GelAnalysis(CallerID) = Nothing
'End If
End Sub

Private Sub Settings()
Dim i As Long
    With GelData(CallerID)
        Select Case UCase(.CalEquation)
        Case UCase(CAL_EQUATION_1)
             cmbCalEq.Text = CAL_EQUATION_1
        Case UCase(CAL_EQUATION_2)
             cmbCalEq.Text = CAL_EQUATION_2
        Case UCase(CAL_EQUATION_3)
             cmbCalEq.Text = CAL_EQUATION_3
        Case UCase(CAL_EQUATION_4)
             cmbCalEq.Text = CAL_EQUATION_4
        Case UCase(CAL_EQUATION_5)
             cmbCalEq.Text = CAL_EQUATION_5
        End Select
        For i = 1 To 4
          txtCalEqPar(i - 1).Text = .CalArg(i)
        Next i
        txtComment = .Comment
        txtDFilesPath = .PathtoDataFiles
        
        Select Case GelStatus(CallerID).SourceDataRawFileType
        Case rfcZippedSFolders
            cboRawDataFileFormat.ListIndex = rfcZippedSFolders
        Case rfcFinniganRaw
            cboRawDataFileFormat.ListIndex = rfcFinniganRaw
        Case Else
           ' Includes rfcUnknown
            cboRawDataFileFormat.ListIndex = rfcUnknown
        End Select
        
        txtInputFilePath = .FileName
        txtFinniganRawFilePath.Text = GelStatus(CallerID).FinniganRawFilePath
    
        txtLegacyDBPath = .PathtoDatabase
        
        txtCertificate = .Certificate
        txtMediaType = .MediaType
    End With
End Sub

Private Sub ShowHidePNNLMenus()
    Dim blnEnabled As Boolean
    blnEnabled = Not APP_BUILD_DISABLE_MTS
    
    cmdDummyMTLink.Enabled = blnEnabled
    cmdSelectMassTags.Enabled = blnEnabled
End Sub

' MonroeMod Start
Private Sub UpdateIniFileWithDBInfo()
    Dim udtDBSettings As udtDBSettingsType
    Dim lngCurrentJob As Long
    
    FillDBSettingsUsingAnalysisObject udtDBSettings, GelAnalysis(CallerID)
    
    If Not GelAnalysis(CallerID) Is Nothing Then
        lngCurrentJob = GelAnalysis(CallerID).Job
    Else
        lngCurrentJob = 0
    End If
    
    If Not udtDBSettings.IsDeleted Then
        ' Determine number of matching MT tags for the given settings
        udtDBSettings.SelectedMassTagCount = GetMassTagMatchCount(udtDBSettings, lngCurrentJob, Me)
    
        IniFileUpdateRecentDatabaseConnectionInfo udtDBSettings
    End If

End Sub
' MonroeMod Finish

Private Sub tbsParameters_Click(PreviousTab As Integer)
    PositionControls
End Sub

Private Sub txtFinniganRawFilePath_LostFocus()
    GelStatus(CallerID).FinniganRawFilePath = Trim$(txtFinniganRawFilePath.Text)
End Sub

Private Sub txtLegacyDBPath_LostFocus()
    GelData(CallerID).PathtoDatabase = Trim$(txtLegacyDBPath.Text)
End Sub

Private Sub txtCalEqPar_LostFocus(Index As Integer)
If IsNumeric(txtCalEqPar(Index).Text) Then
   GelData(CallerID).CalArg(Index + 1) = CDbl(txtCalEqPar(Index).Text)
Else
   If Len(Trim$(txtCalEqPar(Index).Text)) > 0 Then
      MsgBox "Equation parameters should be numbers dude.", vbOKOnly
      txtCalEqPar(Index).SetFocus
   Else
      txtCalEqPar(Index).Text = 0
   End If
End If
End Sub

Private Function ParametersChange() As Boolean
Dim i As Integer
ParametersChange = True
With GelData(CallerID)
    If .CalEquation <> OldCalEq Then Exit Function
    For i = 1 To 10
       If .CalArg(i) <> OldCalArg(i) Then Exit Function
    Next i
    If .Comment <> OldComment Then Exit Function
    If .PathtoDatabase <> OldPathToDB Then Exit Function
    
End With
ParametersChange = False
End Function

Private Sub PopulateComboBoxes()
    
    FillCalEqCombo
    
    With cboRawDataFileFormat
        .Clear
        .AddItem "Unknown"
        .AddItem "Zipped S-Folders"
        .AddItem "Finnigan .Raw file"
    End With

End Sub

Private Sub PositionControls()
    Dim lngDesiredValue As Long
    
    On Error Resume Next
    
    If Not Me.WindowState = vbMinimized Then
        lngDesiredValue = Me.width - tbsParameters.Left - 240
        If lngDesiredValue < 7450 Then lngDesiredValue = 7450
        
        tbsParameters.width = lngDesiredValue
        
        If txtDFilesPath.Left > 0 Then
            cmdBrowseForDataFileFolder.Left = tbsParameters.width - cmdBrowseForDataFileFolder.width - 120
            cmdBrowseForInputFile.Left = cmdBrowseForDataFileFolder.Left
            cmdBrowseForFinniganRawFilePath.Left = cmdBrowseForDataFileFolder.Left
            cmdBrowseForLegacyDB.Left = cmdBrowseForDataFileFolder.Left
            
            txtDFilesPath.width = cmdBrowseForDataFileFolder.Left - txtDFilesPath.Left - 120
            txtInputFilePath.width = txtDFilesPath.width
            txtFinniganRawFilePath.width = txtDFilesPath.width
            txtLegacyDBPath.width = txtDFilesPath.width
        End If
    End If

End Sub

Private Sub txtCertificate_LostFocus()
GelData(CallerID).Certificate = txtCertificate.Text
End Sub

Private Sub txtComment_LostFocus()
GelData(CallerID).Comment = txtComment.Text
End Sub

' No longer supported (March 2006)
''Private Sub txtDBPath_LostFocus()
''GelData(CallerID).PathtoDatabase = txtDBPath.Text
''End Sub

Private Sub txtDFilesPath_LostFocus()
GelData(CallerID).PathtoDataFiles = txtDFilesPath.Text
End Sub

Private Sub txtMediaType_LostFocus()
GelData(CallerID).MediaType = txtMediaType.Text
End Sub

Private Sub txtInputFilePath_LostFocus()
GelData(CallerID).FileName = txtInputFilePath.Text
End Sub
