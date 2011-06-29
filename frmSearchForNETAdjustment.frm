VERSION 5.00
Begin VB.Form frmSearchForNETAdjustment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search MT Tag Database For NET Adjustment"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4410
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraID 
      Caption         =   "Search (Identification) Options"
      Height          =   2655
      Left            =   120
      TabIndex        =   13
      Top             =   3120
      Width           =   4215
      Begin VB.TextBox txtNETTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3240
         TabIndex        =   24
         Text            =   "0.2"
         Top             =   1740
         Width           =   615
      End
      Begin VB.CheckBox chkUseNETForID 
         Caption         =   "Use NET criteria with tolerance"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1800
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.Frame fraMWTolerance 
         Caption         =   "MW Tolerance"
         Height          =   1215
         Left            =   2160
         TabIndex        =   18
         Top             =   360
         Width           =   1935
         Begin VB.TextBox txtMWTol 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   160
            TabIndex        =   21
            Text            =   "10"
            Top             =   640
            Width           =   735
         End
         Begin VB.OptionButton optTolType 
            Caption         =   "&ppm"
            Height          =   255
            Index           =   0
            Left            =   1020
            TabIndex        =   20
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optTolType 
            Caption         =   "&Dalton"
            Height          =   255
            Index           =   1
            Left            =   1020
            TabIndex        =   19
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Tolerance"
            Height          =   255
            Left            =   160
            TabIndex        =   22
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame fraMWField 
         Caption         =   "Molecular Mass Field"
         Height          =   1215
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1935
         Begin VB.OptionButton optMWField 
            Caption         =   "A&verage"
            Height          =   255
            Index           =   0
            Left            =   80
            TabIndex        =   17
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optMWField 
            Caption         =   "&Monoisotopic"
            Height          =   255
            Index           =   1
            Left            =   80
            TabIndex        =   16
            Top             =   540
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton optMWField 
            Caption         =   "&The Most Abundant"
            Height          =   255
            Index           =   2
            Left            =   80
            TabIndex        =   15
            Top             =   840
            Width           =   1815
         End
      End
      Begin VB.Label Label5 
         Caption         =   "(NET formula here is generic normalized elution: NET=(Scan-ScanMin)/(ScanMax-ScanMin))"
         Height          =   495
         Left            =   360
         TabIndex        =   25
         Top             =   2040
         Width           =   3495
      End
   End
   Begin VB.Frame fraNET 
      Caption         =   "NET  Adjustment Calculation"
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   4215
      Begin VB.TextBox txtMaxIDToUse 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3240
         TabIndex        =   11
         Text            =   "2000"
         Top             =   2220
         Width           =   615
      End
      Begin VB.CheckBox chkEliminateBadNETIDs 
         Caption         =   "Eliminate identifications with bad NET"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   3495
      End
      Begin VB.CheckBox chkAutoElimination 
         Caption         =   "Automatically select identifications(if too many)"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Width           =   3855
      End
      Begin VB.TextBox txtEliminateHowLong 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3100
         TabIndex        =   6
         Text            =   "20"
         Top             =   675
         Width           =   375
      End
      Begin VB.CheckBox chkEliminateLongIDs 
         Caption         =   "Eliminate identifications lasting over"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   2895
      End
      Begin VB.CheckBox chkEliminateDuplicates 
         Caption         =   "Use only first instance of each identification"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   "Number of identifications to use for adj. "
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2040
         Width           =   3735
      End
      Begin VB.Label Label3 
         Caption         =   "%"
         Height          =   255
         Index           =   0
         Left            =   3500
         TabIndex        =   7
         Top             =   720
         Width           =   375
      End
   End
   Begin VB.Label lblStatus 
      Caption         =   "quo vadis domine?"
      Height          =   425
      Left            =   120
      TabIndex        =   3
      Top             =   5880
      Width           =   4215
   End
   Begin VB.Label lblMTStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MTDB Status:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Status of the MT Tag database"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Menu mnuF 
      Caption         =   "&Function"
      Begin VB.Menu mnuFCalculate 
         Caption         =   "C&alculate"
      End
      Begin VB.Menu mnuFReport 
         Caption         =   "&Report"
      End
      Begin VB.Menu mnuFExportToMTDB 
         Caption         =   "&Export to MTDB"
      End
      Begin VB.Menu mnuFSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuMT 
      Caption         =   "&MT Tags"
      Begin VB.Menu mnuMTLoad 
         Caption         =   "&Load MT Tags"
      End
      Begin VB.Menu mnuLoadLegacy 
         Caption         =   "Load L&egacy MT Tags"
      End
      Begin VB.Menu mnuMTSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMTStatus 
         Caption         =   "&Status"
      End
   End
End
Attribute VB_Name = "frmSearchForNETAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------
'last modified: 09/26/2002 nt
'-----------------------------------------------------------------
'NOTE: in this case NET array contains Lars elution times and RT
'      array contains their standard deviation
'NOTE: Search is always on all data points
'NOTE: no data is changed based on this function
'-----------------------------------------------------------------
'This is how it works; find all matching MT tags then look to
'reduce that number by selecting pairs that will yield best match
'(eliminate all except first instance of each peptide; use higher
'intensity peaks). Look for best Slope and Intercept for IDs by
'least square method.
'-----------------------------------------------------------------
Const MAX_ID_CNT = 50000            'maximum number of IDs
Const MAX_ID_TO_USE = 2000

Const STATE_NOT_1ST_INSTANCE = 1024
Const STATE_TOO_LONG_ELUTION = 2048
Const STATE_BAD_NET = 4096
Const STATE_OUTSCORED = 8192

Dim GRID() As GR            'GRID array is parallel with AMT array
                            'it is indexed the same way and it's
                            'members are indexes in ID belonging to
                            'AMT (in other words have same ID as
                            'MT tags index in AMT arrays)

Dim IDCnt As Long           'identification count
Dim ID() As Long            'index in AMT array
Dim IDInd() As Long         'peak index in data arrays
Dim IDType() As Long        'type of peak
Dim IDState() As Long       'used to clean IDs from duplicates and bads
Dim IDScan() As Long        'this is redundant but will make life easier

Option Explicit

'in this case CallerID is a public property
Public CallerID As Long

Dim bLoading As Boolean

'load from definition but then adjust for this search
Dim ThisDef As SearchAMTDefinition

'options that would reduce number of identifications
Dim bUseFirstIDOnly As Boolean      'if True only first instance of each
                                    'identifications will be used
Dim bEliminateLongIDs As Boolean    'if True only ids that elute for less
                                    'than EliminateHowLong will be used
Dim EliminateHowLong As Double      'percentage of run
Dim bAutoElimination As Boolean     'if True and we have too many IDs use
                                    'scoring procedure to reduce numbers
Dim bEliminateBadNETs As Boolean    'eliminate identification with out of
                                    'range elution numbers
Dim MaxIDToUse As Long              'maximum number of identifications to use
Dim bUseNETForID As Boolean         'if True use NET criteria for search; if
                                    'False use only MW
                                    
Dim ScanMin As Long                 'first scan number
Dim ScanMax As Long                 'last scan number
Dim ScanRange As Long               'last-first+1

Dim AdjSlp As Double                'slope of GANET adjustments
Dim AdjInt As Double                'intercept of GANET adjustment
Dim AdjAvD As Double

Dim EditGANETSPName As String

                                                                        
Private Sub chkAutoElimination_Click()
bAutoElimination = (chkAutoElimination.Value = vbChecked)
End Sub

Private Sub chkEliminateBadNETIDs_Click()
bEliminateBadNETs = (chkEliminateBadNETIDs.Value = vbChecked)
End Sub

Private Sub chkEliminateDuplicates_Click()
bUseFirstIDOnly = (chkEliminateDuplicates.Value = vbChecked)
End Sub

Private Sub chkEliminateLongIDs_Click()
bEliminateLongIDs = (chkEliminateLongIDs.Value = vbChecked)
End Sub


Private Function SearchMassTagsMW() As Boolean
'----------------------------------------------------------
'searches MT tags for matching masses and returns True if
'OK, False if any error or user canceled the whole thing
'----------------------------------------------------------
Dim I As Long, j As Long
Dim eResponse As VbMsgBoxResult
Dim TmpCnt As Long
Dim Hits() As Long
Dim MW As Double, MWAbsErr As Double
Dim Scan As Long
On Error GoTo err_SearchMassTagsMW

UpdateStatus "Searching for MT tags ..."
ReDim ID(MAX_ID_CNT - 1)            'prepare for the worst case
ReDim IDInd(MAX_ID_CNT - 1)
ReDim IDType(MAX_ID_CNT - 1)
ReDim IDScan(MAX_ID_CNT - 1)
ReDim IDState(MAX_ID_CNT - 1)

With GelData(CallerID)
     Select Case ThisDef.TolType
     Case gltPPM
          For I = 1 To .CSLines
             MW = .CSData(I).AverageMW
             Scan = GelData(CallerID).CSData(I).ScanNumber
             MWAbsErr = MW * ThisDef.MWTol * glPPM
             Select Case ThisDef.NETorRT
             Case glAMT_NET
                TmpCnt = GetMTHits1(MW, MWAbsErr, -1, -1, Hits())
             Case glAMT_RT_or_PNET
                TmpCnt = GetMTHits2(MW, MWAbsErr, -1, -1, Hits())
             End Select
             If TmpCnt > 0 Then
                For j = 0 To TmpCnt - 1
                    IDType(IDCnt) = glCSType
                    IDInd(IDCnt) = I
                    ID(IDCnt) = Hits(j)
                    IDState(IDCnt) = 0
                    IDScan(IDCnt) = Scan
                    IDCnt = IDCnt + 1   'if reaches limit we will have correct
                Next j                  'results by doing increase at the end
             End If
          Next I
          For I = 1 To .IsoLines
             MW = GetIsoMass(.IsoData(I), ThisDef.MWField)
             MWAbsErr = MW * ThisDef.MWTol * glPPM
             Scan = GelData(CallerID).IsoData(I).ScanNumber
             Select Case ThisDef.NETorRT
             Case glAMT_NET
                TmpCnt = GetMTHits1(MW, MWAbsErr, -1, -1, Hits())
             Case glAMT_RT_or_PNET
                TmpCnt = GetMTHits2(MW, MWAbsErr, -1, -1, Hits())
             End Select
             If TmpCnt > 0 Then
                For j = 0 To TmpCnt - 1
                    IDType(IDCnt) = glIsoType
                    IDInd(IDCnt) = I
                    ID(IDCnt) = Hits(j)
                    IDState(IDCnt) = 0
                    IDScan(IDCnt) = Scan
                    IDCnt = IDCnt + 1
                Next j
             End If
          Next I
     Case gltABS
          MWAbsErr = ThisDef.MWTol
          For I = 1 To .CSLines
             MW = .CSData(I).AverageMW
             Scan = GelData(CallerID).CSData(I).ScanNumber
             Select Case ThisDef.NETorRT
             Case glAMT_NET
                TmpCnt = GetMTHits1(MW, MWAbsErr, -1, -1, Hits())
             Case glAMT_RT_or_PNET
                TmpCnt = GetMTHits2(MW, MWAbsErr, -1, -1, Hits())
             End Select
             If TmpCnt > 0 Then
                For j = 0 To TmpCnt - 1
                    IDType(IDCnt) = glCSType
                    IDInd(IDCnt) = I
                    ID(IDCnt) = Hits(j)
                    IDState(IDCnt) = 0
                    IDScan(IDCnt) = Scan
                    IDCnt = IDCnt + 1   'if reaches limit we will have correct
                Next j                  'results by doing increase at the end
             End If
          Next I
          For I = 1 To .IsoLines
             MW = GetIsoMass(.IsoData(I), ThisDef.MWField)
             Scan = GelData(CallerID).IsoData(I).ScanNumber
             Select Case ThisDef.NETorRT
             Case glAMT_NET
                TmpCnt = GetMTHits1(MW, MWAbsErr, -1, -1, Hits())
             Case glAMT_RT_or_PNET
                TmpCnt = GetMTHits2(MW, MWAbsErr, -1, -1, Hits())
             End Select
             If TmpCnt > 0 Then
                For j = 0 To TmpCnt - 1
                    IDType(IDCnt) = glIsoType
                    IDInd(IDCnt) = I
                    ID(IDCnt) = Hits(j)
                    IDState(IDCnt) = 0
                    IDScan(IDCnt) = Scan
                    IDCnt = IDCnt + 1
                Next j
             End If
          Next I
    Case Else
          Debug.Assert False
     End Select
End With

If IDCnt > 0 Then
   ReDim Preserve ID(IDCnt - 1)
   ReDim Preserve IDInd(IDCnt - 1)
   ReDim Preserve IDType(IDCnt - 1)
   ReDim Preserve IDScan(IDCnt - 1)
   ReDim Preserve IDState(IDCnt - 1)
Else
   Call ClearIDArrays
End If
UpdateStatus "Possible identifications: " & IDCnt
SearchMassTagsMW = True
Exit Function


err_SearchMassTagsMW:
Select Case Err.Number
Case 9      'too many identifications
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        eResponse = MsgBox("Too many possible identifications detected.  " _
                   & "To proceed with the first " & MAX_ID_CNT & _
                   " identifications select OK.", vbOKCancel, glFGTU)
    Else
        eResponse = vbOK
    End If
    
    Select Case eResponse
    Case vbOK
        UpdateStatus "Possible identifications: " & IDCnt
        SearchMassTagsMW = True
    Case Else
        UpdateStatus ""
        Call ClearIDArrays
    End Select
Case 7      'short on memory; try to recover by releasing arrays
    Call ClearIDArrays
    UpdateStatus ""
    MsgBox "System low on memory. Process aborted in recovery attempt.", vbOKOnly, glFGTU
Case Else
    UpdateStatus "Error searching for MT tags."
    LogErrors Err.Number, "frmSearchForNETAdjustment.SearchMassTagsMW"
End Select
End Function


Private Function SearchMassTagsMWNET() As Boolean
'----------------------------------------------------------
'searches MT tags for matching masses and returns True if
'OK, False if any error or user canceled the whole thing
'----------------------------------------------------------
Dim I As Long, j As Long
Dim eResponse As VbMsgBoxResult
Dim TmpCnt As Long
Dim Hits() As Long
Dim MW As Double, MWAbsErr As Double
Dim Scan As Long
On Error GoTo err_SearchMassTagsMWNET

UpdateStatus "Searching for MT tags ..."
ReDim ID(MAX_ID_CNT - 1)            'prepare for the worst case
ReDim IDInd(MAX_ID_CNT - 1)
ReDim IDType(MAX_ID_CNT - 1)
ReDim IDScan(MAX_ID_CNT - 1)
ReDim IDState(MAX_ID_CNT - 1)

With GelData(CallerID)
     Select Case ThisDef.TolType
     Case gltPPM
          For I = 1 To .CSLines
             MW = .CSData(I).AverageMW
             Scan = GelData(CallerID).CSData(I).ScanNumber
             MWAbsErr = MW * ThisDef.MWTol * glPPM
             Select Case ThisDef.NETorRT
             Case glAMT_NET
                TmpCnt = GetMTHits1(MW, MWAbsErr, GetNET(Scan), ThisDef.NETTol, Hits())
             Case glAMT_RT_or_PNET
                TmpCnt = GetMTHits2(MW, MWAbsErr, GetNET(Scan), ThisDef.NETTol, Hits())
             End Select
             If TmpCnt > 0 Then
                For j = 0 To TmpCnt - 1
                    IDType(IDCnt) = glCSType
                    IDInd(IDCnt) = I
                    ID(IDCnt) = Hits(j)
                    IDState(IDCnt) = 0
                    IDScan(IDCnt) = Scan
                    IDCnt = IDCnt + 1   'if reaches limit we will have correct
                Next j                  'results by doing increase at the end
             End If
          Next I
          For I = 1 To .IsoLines
             MW = GetIsoMass(.IsoData(I), ThisDef.MWField)
             MWAbsErr = MW * ThisDef.MWTol * glPPM
             Scan = GelData(CallerID).IsoData(I).ScanNumber
             Select Case ThisDef.NETorRT
             Case glAMT_NET
                TmpCnt = GetMTHits1(MW, MWAbsErr, GetNET(Scan), ThisDef.NETTol, Hits())
             Case glAMT_RT_or_PNET
                TmpCnt = GetMTHits2(MW, MWAbsErr, GetNET(Scan), ThisDef.NETTol, Hits())
             End Select
             If TmpCnt > 0 Then
                For j = 0 To TmpCnt - 1
                    IDType(IDCnt) = glIsoType
                    IDInd(IDCnt) = I
                    ID(IDCnt) = Hits(j)
                    IDState(IDCnt) = 0
                    IDScan(IDCnt) = Scan
                    IDCnt = IDCnt + 1
                Next j
             End If
          Next I
     Case gltABS
          MWAbsErr = ThisDef.MWTol
          For I = 1 To .CSLines
             MW = .CSData(I).AverageMW
             Scan = GelData(CallerID).CSData(I).ScanNumber
             Select Case ThisDef.NETorRT
             Case glAMT_NET
                TmpCnt = GetMTHits1(MW, MWAbsErr, GetNET(Scan), ThisDef.NETTol, Hits())
             Case glAMT_RT_or_PNET
                TmpCnt = GetMTHits2(MW, MWAbsErr, GetNET(Scan), ThisDef.NETTol, Hits())
             End Select
             If TmpCnt > 0 Then
                For j = 0 To TmpCnt - 1
                    IDType(IDCnt) = glCSType
                    IDInd(IDCnt) = I
                    ID(IDCnt) = Hits(j)
                    IDState(IDCnt) = 0
                    IDScan(IDCnt) = Scan
                    IDCnt = IDCnt + 1   'if reaches limit we will have correct
                Next j                  'results by doing increase at the end
             End If
          Next I
          For I = 1 To .IsoLines
             MW = GetIsoMass(.IsoData(I), ThisDef.MWField)
             Scan = GelData(CallerID).IsoData(I).ScanNumber
             Select Case ThisDef.NETorRT
             Case glAMT_NET
                TmpCnt = GetMTHits1(MW, MWAbsErr, GetNET(Scan), ThisDef.NETTol, Hits())
             Case glAMT_RT_or_PNET
                TmpCnt = GetMTHits2(MW, MWAbsErr, GetNET(Scan), ThisDef.NETTol, Hits())
             End Select
             If TmpCnt > 0 Then
                For j = 0 To TmpCnt - 1
                    IDType(IDCnt) = glIsoType
                    IDInd(IDCnt) = I
                    ID(IDCnt) = Hits(j)
                    IDState(IDCnt) = 0
                    IDScan(IDCnt) = Scan
                    IDCnt = IDCnt + 1
                Next j
             End If
          Next I
    Case Else
          Debug.Assert False
     End Select
End With

If IDCnt > 0 Then
   ReDim Preserve ID(IDCnt - 1)
   ReDim Preserve IDInd(IDCnt - 1)
   ReDim Preserve IDType(IDCnt - 1)
   ReDim Preserve IDScan(IDCnt - 1)
   ReDim Preserve IDState(IDCnt - 1)
Else
   Call ClearIDArrays
End If
UpdateStatus "Possible identifications: " & IDCnt
SearchMassTagsMWNET = True
Exit Function


err_SearchMassTagsMWNET:
Select Case Err.Number
Case 9      'too many identifications
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        eResponse = MsgBox("Too many possible identifications detected.  " _
                   & "To proceed with the first " & MAX_ID_CNT & _
                   " identifications select OK.", vbOKCancel, glFGTU)
    Else
        eResponse = vbOK
    End If

    Select Case eResponse
    Case vbOK
        UpdateStatus "Possible identifications: " & IDCnt
        SearchMassTagsMWNET = True
    Case Else
        UpdateStatus ""
        Call ClearIDArrays
    End Select
Case 7      'short on memory; try to recover by releasing arrays
    Call ClearIDArrays
    UpdateStatus ""
    MsgBox "System low on memory. Process aborted in recovery attempt.", vbOKOnly, glFGTU
Case Else
    UpdateStatus "Error searching for MT tags."
    LogErrors Err.Number, "frmSearchForNETAdjustment.SearchMassTagsMWNET"
End Select
End Function

Private Sub ShowHidePNNLMenus()
    Dim blnVisible As Boolean
    blnVisible = Not APP_BUILD_DISABLE_MTS
    
    mnuFExportToMTDB.Visible = blnVisible
    mnuMTLoad.Visible = blnVisible
End Sub

Private Sub chkUseNETForID_Click()
bUseNETForID = (chkUseNETForID.Value = vbChecked)
End Sub

Private Sub Form_Activate()
'------------------------------------------------------------
'load MT tag database data if neccessary
'if CallerID is associated with MT tag database load that
'database if neccessary; if CallerID is not associated with
'MT tag database load legacy database
'------------------------------------------------------------
On Error Resume Next
UpdateStatus ""
Me.MousePointer = vbHourglass
If bLoading Then
   If GelAnalysis(CallerID) Is Nothing Then
      If AMTCnt > 0 Then    'something is loaded
          If (Len(CurrMTDatabase) > 0 Or Len(CurrLegacyMTDatabase) > 0) And Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            'MT tag data; we dont know is it appropriate; warn user
            WarnUserUnknownMassTags CallerID
         End If
         lblMTStatus.Caption = ConstructMTStatusText(False)
      
         ' Initialize the MT search object
         If Not CreateNewMTSearchObject() Then
            lblMTStatus.Caption = "Error creating search object."
         End If
      
      Else                  'nothing is loaded
         WarnUserNotConnectedToDB CallerID, True
         lblMTStatus.Caption = "No MT tags loaded"
      End If
   Else         'have to have MT tag database loaded
      Call LoadMTDB
   End If
   GetScanRange CallerID, ScanMin, ScanMax, ScanRange
   If ScanRange < 2 Then
      MsgBox "Scan range for this display could lead to unpredictable results.", vbOKOnly, glFGTU
   End If
   bLoading = False
   EditGANETSPName = glbPreferencesExpanded.MTSConnectionInfo.spEditGANET
   If Len(EditGANETSPName) > 0 Then         'enable export if neccessary
      mnuFExportToMTDB.Enabled = True
   Else
      mnuFExportToMTDB.Enabled = False
   End If
End If
Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
bLoading = True

' MonroeMod
If CallerID >= 1 And CallerID <= UBound(GelBody) Then samtDef = GelSearchDef(CallerID).AMTSearchOnIons
ThisDef = samtDef
If IsWinLoaded(TrackerCaption) Then Unload frmTracker

ShowHidePNNLMenus

'set current Search Definition values
With ThisDef
    txtMWTol.Text = .MWTol
    .SearchScope = glScope.glSc_All
    optMWField(.MWField - MW_FIELD_OFFSET).Value = True
    Select Case .TolType
    Case gltPPM
         optTolType(0).Value = True
    Case gltABS
         optTolType(1).Value = True
    Case Else
         Debug.Assert False
    End Select
    .NETTol = 0.2
    txtNETTol.Text = .NETTol
    .NETorRT = glAMT_NET
    .SearchFlag = 0         'search all
    .MassTag = 0            'no MT tags assumptions here
    .MaxMassTags = 0
    .SkipReferenced = False
    .SaveNCnt = False
End With
bUseFirstIDOnly = (chkEliminateDuplicates.Value = vbChecked)
bEliminateLongIDs = (chkEliminateLongIDs.Value = vbChecked)
EliminateHowLong = CDbl(txtEliminateHowLong.Text)
bEliminateBadNETs = (chkEliminateBadNETIDs.Value = vbChecked)
bAutoElimination = (chkAutoElimination.Value = vbChecked)
bUseNETForID = (chkUseNETForID.Value = vbChecked)
MaxIDToUse = CLng(txtMaxIDToUse.Text)
DoEvents
End Sub

Private Sub mnuF_Click()
Call PickParameters
chkEliminateDuplicates.SetFocus     'probably not neccessary
End Sub

Private Sub mnuFCalculate_Click()
'--------------------------------------------
'first look for hits; then process them; if
'more then acceptable to work with clean them
'--------------------------------------------
Dim Res As Long
On Error Resume Next
Call ClearIDArrays
Call ClearTheGRID
If bUseNETForID Then
   If Not SearchMassTagsMWNET() Then Exit Sub
Else
   If Not SearchMassTagsMW() Then Exit Sub
End If
If Not FillTheGRID() Then Exit Sub
'do some prefiltering; maybe this will be enough to eliminate auto selection
If bUseFirstIDOnly Then MarkDuplicateIDInstances
If bEliminateLongIDs Then MarkLongIDs
If bEliminateBadNETs Then MarkIDsWithBadNET
Call CleanIdentifications
If IDCnt > MaxIDToUse Then
   If Not bAutoElimination Then     'ask
      Res = MsgBox("Number of identifications too high. Do you want application to select " & MaxIDToUse & " best looking IDs for you?", vbYesNoCancel, glFGTU)
      If Res <> vbYes Then
         MsgBox "To reduce number of identifications tighten your search criteria or allow application to choose set of IDs to use.", vbOKOnly, glFGTU
         Call ClearIDArrays
         Call ClearTheGRID
         Exit Sub
      End If
   End If
   If AutoSelectIDs() Then
      Call CleanIdentifications
   Else
      MsgBox "Auto selection process failed. Change settings and try again.", vbOKOnly, glFGTU
      Call ClearIDArrays
      Call ClearTheGRID
   End If
End If
Me.MousePointer = vbHourglass
Call CalculateSlopeIntercept
'don't automatically  send it to the database but fill memory
'structures so that we can use it to search with that function
If Not GelAnalysis(CallerID) Is Nothing Then
   GelAnalysis(CallerID).GANET_Fit = AdjAvD
   GelAnalysis(CallerID).GANET_Slope = AdjSlp
   GelAnalysis(CallerID).GANET_Intercept = AdjInt
End If
Me.MousePointer = vbDefault
MsgBox "Slope: " & AdjSlp & vbCrLf & "Intercept: " & AdjInt & vbCrLf & "Average Deviation: " & AdjAvD, vbOKOnly, glFGTU

AddToAnalysisHistory CallerID, "Calculated NET adjustment using individual spectra; Mass tolerance = ±" & Trim(ThisDef.MWTol) & " " & GetSearchToleranceUnitText(CInt(ThisDef.TolType)) & "; NET Tolerance = ±" & Trim(ThisDef.NETTol) & "; Peaks (ions) with DB hits = " & Trim(IDCnt) & "; Slope = " & Trim(AdjSlp) & "; Intercept = " & Trim(AdjInt) & "; Average Deviation = " & Trim(AdjAvD)
AddToAnalysisHistory CallerID, "Database for NET adjustment = " & ExtractDBNameFromConnectionString(GelAnalysis(CallerID).MTDB.cn.ConnectionString) & "; MT tag Count = " & AMTCnt

End Sub

Private Sub mnuFClose_Click()
Unload Me
End Sub

Private Sub mnuFExportToMTDB_Click()
    MsgBox ExportGANETtoMTDB(CallerID, AdjSlp, AdjInt, AdjAvD)
End Sub

Private Sub mnuFReport_Click()
Call ReportAdjustments
End Sub

Private Sub mnuLoadLegacy_Click()
'------------------------------------------------------------
'load/reload MT tags
'------------------------------------------------------------
Dim Respond As Long
On Error Resume Next
'ask user if it wants to replace legitimate MT tag DB with legacy DB
If Not GelAnalysis(CallerID) Is Nothing And Not APP_BUILD_DISABLE_MTS Then
   Respond = MsgBox("Current display is associated with MT tag database." & vbCrLf _
                & "Are you sure you want to use a legacy database for search?", vbYesNoCancel, glFGTU)
   If Respond <> vbYes Then Exit Sub
End If
Me.MousePointer = vbHourglass
If Len(GelData(CallerID).PathtoDatabase) > 0 Then
   If ConnectToLegacyAMTDB(Me, CallerID, False, True, False) Then
      If CreateNewMTSearchObject() Then
         lblMTStatus.Caption = "Loaded; MT tag count: " & LongToStringWithCommas(AMTCnt)
      Else
         lblMTStatus.Caption = "Error creating search object."
      End If
   Else
      lblMTStatus.Caption = "Error loading MT tags."
   End If
Else
    WarnUserInvalidLegacyDBPath
End If
Me.MousePointer = vbDefault
End Sub

Private Sub mnuMT_Click()
Call PickParameters
chkEliminateDuplicates.SetFocus    'probably not neccessary
End Sub

Private Sub mnuMTLoad_Click()
'------------------------------------------------------------
'load/reload MT tags
'------------------------------------------------------------
If Not GelAnalysis(CallerID) Is Nothing Then
   Call LoadMTDB(True)
Else
   WarnUserNotConnectedToDB CallerID, True
   lblMTStatus.Caption = "No MT tags loaded"
End If
End Sub

Private Sub mnuMTStatus_Click()
'----------------------------------------------
'displays short MT tags statistics, it might
'help with determining problems with MT tags
'----------------------------------------------
Me.MousePointer = vbHourglass
MsgBox CheckMassTags(), vbOKOnly
Me.MousePointer = vbDefault
End Sub

Private Sub optMWField_Click(Index As Integer)
ThisDef.MWField = 6 + Index
End Sub

Private Sub optTolType_Click(Index As Integer)
If Index = 0 Then
   ThisDef.TolType = gltPPM
Else
   ThisDef.TolType = gltABS
End If
End Sub


Private Sub txtEliminateHowLong_LostFocus()
Dim tmp As String
tmp = Trim$(txtEliminateHowLong.Text)
If IsNumeric(tmp) Then
   If tmp >= 0 And tmp <= 100 Then
      EliminateHowLong = CDbl(tmp)
      Exit Sub
   End If
End If
MsgBox "This argument should be a number between 0 and 100.", vbOKOnly, glFGTU
txtEliminateHowLong.SetFocus
End Sub



Private Sub txtMaxIDToUse_LostFocus()
Dim tmp As String
Dim Res As Long
tmp = Trim$(txtMaxIDToUse.Text)
If IsNumeric(tmp) Then
   If tmp > 0 Then
      If tmp <= 100 Then
         Res = MsgBox("Number of identifications is too low for reliable calculations(use at least 100). Accept anyway?", vbYesNo, glFGTU)
         If Res <> vbYes Then
            txtMaxIDToUse.SetFocus
            Exit Sub
         End If
      End If
      If tmp <= 2000 Then
         MaxIDToUse = CLng(tmp)
         Exit Sub
      End If
   End If
End If
MsgBox "Use number between 100 and " & MAX_ID_TO_USE & ".", vbOKOnly, glFGTU
txtMaxIDToUse.SetFocus
End Sub


Private Sub txtMWTol_LostFocus()
If IsNumeric(txtMWTol.Text) Then
   ThisDef.MWTol = CDbl(txtMWTol.Text)
Else
   MsgBox "Molecular Mass Tolerance should be numeric value.", vbOKOnly
   txtMWTol.SetFocus
End If
End Sub


Private Sub CleanIdentifications()
'-------------------------------------------------------------
'removes identifications that will not be used from the arrays
'-------------------------------------------------------------
Dim I As Long
Dim NewCnt As Long
On Error Resume Next
UpdateStatus "Restructuring data ..."
For I = 0 To IDCnt - 1
    If IDState(I) = 0 Then
       NewCnt = NewCnt + 1
       ID(NewCnt - 1) = ID(I)
       IDInd(NewCnt - 1) = IDInd(I)
       IDType(NewCnt - 1) = IDType(I)
       IDScan(NewCnt - 1) = IDScan(I)
    End If
Next I
If NewCnt > 0 Then
   ReDim Preserve ID(NewCnt - 1)
   ReDim Preserve IDInd(NewCnt - 1)
   ReDim Preserve IDType(NewCnt - 1)
   ReDim IDState(NewCnt - 1)            'after cleaning always put state to 0
   ReDim Preserve IDScan(NewCnt - 1)
Else
   Call ClearIDArrays
End If
If NewCnt < IDCnt Then
   IDCnt = NewCnt
   ClearTheGRID         'have to recalculate GRID data
   Call FillTheGRID
End If
UpdateStatus ""
End Sub

Public Sub ClearIDArrays()
Erase ID
Erase IDInd
Erase IDType
Erase IDState
Erase IDScan
IDCnt = 0
End Sub

Private Function AutoSelectIDs() As Boolean
'---------------------------------------------------------------------
'score IDs and mark for elimination those not in first MaxIDToUse
'this is used when we have too many IDs to work with
'---------------------------------------------------------------------
On Error GoTo err_AutoSelectIDs
'first try standard things; if not enough do scoring and select best
If Not bEliminateBadNETs Then
   Call MarkIDsWithBadNET
   Call CleanIdentifications
   If IDCnt <= MaxIDToUse Then
      AutoSelectIDs = True
      Exit Function
   End If
End If
If Not bUseFirstIDOnly Then
   Call MarkDuplicateIDInstances
   Call CleanIdentifications
   If IDCnt <= MaxIDToUse Then
      AutoSelectIDs = True
      Exit Function
   End If
End If
If Not bEliminateLongIDs Then
   Call MarkLongIDs
   Call CleanIdentifications
   If IDCnt <= MaxIDToUse Then
      AutoSelectIDs = True
      Exit Function
   End If
End If
If Not ScoreIDs() Then GoTo err_AutoSelectIDs
Call CleanIdentifications
AutoSelectIDs = True
Exit Function

err_AutoSelectIDs:
End Function

Private Sub ReportAdjustments()
'---------------------------------------------------
'report identifications on which adjustment is based
'and actual adjustment
'---------------------------------------------------
Dim fso As New FileSystemObject
Dim ts As TextStream
Dim fname As String
Dim I As Long
On Error Resume Next
UpdateStatus "Generating report ..."
fname = GetTempFolder() & RawDataTmpFile
Set ts = fso.OpenTextFile(fname, ForWriting, True)
ts.WriteLine "Generated by: " & GetMyNameVersion() & " on " & Now()
'print gel file name and pairs definitions as reference
ts.WriteLine "Gel File: " & GelBody(CallerID).Caption
ts.WriteLine "Reporting NET adjustment(based on MT tag NETs)"
ts.WriteLine
ts.WriteLine "Slope: " & AdjSlp
ts.WriteLine "Intercept: " & AdjInt
ts.WriteLine "Average Deviation: " & AdjAvD
ts.WriteLine
ts.WriteLine "ID" & glARG_SEP & "ID_NET" & glARG_SEP & "Scan"
For I = 0 To IDCnt - 1
    ts.WriteLine Trim(AMTData(ID(I)).ID) & glARG_SEP & AMTData(ID(I)).NET & glARG_SEP & IDScan(I)
Next I
ts.Close
Set fso = Nothing
UpdateStatus ""
DoEvents
frmDataInfo.Tag = "AdjNET"
frmDataInfo.Show vbModal
End Sub



Private Function FillTheGRID() As Boolean
'----------------------------------------
'fills GRID arrays with ID information
'----------------------------------------
Dim I As Long
Dim DummyInd() As Long      'dummy array(empty) will allow us to
                            'sort only on one array
Dim QSL As QSLong
On Error GoTo err_FillTheGRID
UpdateStatus "Loading data structures ..."
If IDCnt > 0 And AMTCnt > 0 Then
   ReDim GRID(AMTCnt)      'AMT arrays are 1-based
   For I = 0 To IDCnt - 1
       With GRID(ID(I))
           .Count = .Count + 1
           ReDim Preserve .Members(.Count - 1)
           .Members(.Count - 1) = I
       End With
   Next I
   'order members of each group on scan numbers
   For I = 0 To AMTCnt
       If GRID(I).Count > 1 Then
          Set QSL = New QSLong
          If Not QSL.QSAsc(GRID(I).Members, DummyInd) Then GoTo err_FillTheGRID
          Set QSL = Nothing
       End If
   Next I
   FillTheGRID = True
   UpdateStatus ""
Else
   UpdateStatus "Data not found."
End If
Exit Function

err_FillTheGRID:
Select Case Err.Number
Case 7
   Call ClearIDArrays
   Call ClearTheGRID
   UpdateStatus ""
   MsgBox "System low on memory. Process aborted in recovery attempt.", vbOKOnly, glFGTU
Case Else
   Call ClearIDArrays
   Call ClearTheGRID
   UpdateStatus "Error loading data structures."
End Select
End Function


Private Sub UpdateStatus(ByVal Msg As String)
lblStatus.Caption = Msg
DoEvents
End Sub

Private Sub ClearTheGRID()
'--------------------------------------------
'destroys GRID data structure
'--------------------------------------------
Dim I As Long
On Error Resume Next
For I = 0 To UBound(GRID)
    If GRID(I).Count > 0 Then Erase GRID(I).Members
Next I
Erase GRID
End Sub

Private Sub MarkDuplicateIDInstances()
'----------------------------------------------------
'sets state of all identifications not first instance
'(in scan order) to STATE_NOT_1ST_INSTANCE
'----------------------------------------------------
Dim I As Long, j As Long
On Error Resume Next
UpdateStatus "Eliminating duplicate IDs ..."
For I = 0 To IDCnt - 1
    With GRID(ID(I))
         If .Count > 1 Then
            For j = 1 To .Count - 1
                IDState(.Members(j)) = IDState(.Members(j)) + STATE_NOT_1ST_INSTANCE
            Next j
         End If
    End With
Next I
UpdateStatus ""
End Sub

Private Sub MarkLongIDs()
'-----------------------------------------------------
'sets state of all identifications spaning through the
'long range to STATE_TOO_LONG_ELUTION
'-----------------------------------------------------
Dim AllowedScanRange As Double
Dim I As Long, j As Long
On Error Resume Next
UpdateStatus "Eliminating IDs with long elution ..."
AllowedScanRange = EliminateHowLong * ScanRange / 100
For I = 0 To IDCnt - 1
  With GRID(ID(I))
    If .Count > 1 Then
       If (IDScan(.Members(.Count - 1)) - IDScan(.Members(0))) > AllowedScanRange Then
          For j = 0 To .Count - 1       'mark them all as too long
              IDState(.Members(j)) = IDState(.Members(j)) + STATE_TOO_LONG_ELUTION
          Next j
       End If
    End If
  End With
Next I
UpdateStatus ""
End Sub


Private Sub MarkIDsWithBadNET()
'------------------------------------------------------
'sets state of all identifications with bad NET numbers
'to STATE_BAD_NET
'------------------------------------------------------
Dim I As Long
On Error Resume Next
UpdateStatus "Eliminating IDs with bad elution ..."
For I = 0 To IDCnt - 1
    If AMTData(ID(I)).NET < 0 Or AMTData(ID(I)).NET > 1 Then
       IDState(I) = IDState(I) + STATE_BAD_NET
    End If
Next I
UpdateStatus ""
End Sub

Private Sub LoadMTDB(Optional blnForceReload As Boolean = False)
    Dim blnAMTsWereLoaded As Boolean, blnDBConnectionError As Boolean
    
    If ConfirmMassTagsAndInternalStdsLoaded(Me, CallerID, True, True, False, blnForceReload, 0, blnAMTsWereLoaded, blnDBConnectionError) Then
        lblMTStatus.Caption = ConstructMTStatusText(False)
    
        If Not CreateNewMTSearchObject() Then
           lblMTStatus.Caption = "Error creating search object."
        End If
    
    Else
        If blnDBConnectionError Then
            lblMTStatus.Caption = "Error loading MT tags: database connection error."
        Else
            lblMTStatus.Caption = "Error loading MT tags: no valid MT tags were found (possibly missing NET values)"
        End If
    End If
    
End Sub

Private Function ScoreIDs() As Boolean
'-------------------------------------------------------------------
'score ids and set states of IDs that don't make top MaxIDToUse
'to STATE_OUTSCORED
'Score=Log10(Peak Abundance) - Fit for Isotopic peaks and
'Log10(Peak Abundance) for Charge State peaks
'-------------------------------------------------------------------
Dim ScoreOrder() As Long
Dim Score() As Double
Dim qsd As New QSDouble
Dim I As Long
On Error GoTo err_ScoreIDs

ReDim ScoreOrder(IDCnt - 1)
ReDim Score(IDCnt - 1)
With GelData(CallerID)
    For I = 0 To IDCnt - 1
        ScoreOrder(I) = I
        Select Case IDType(I)
        Case glCSType
             Score(I) = (Log(.CSData(IDInd(I)).Abundance) / Log(10))
        Case glIsoType
             Score(I) = (Log(.IsoData(IDInd(I)).Abundance) / Log(10)) - .IsoData(IDInd(I)).Fit
        End Select
    Next I
End With
If Not qsd.QSDesc(Score(), ScoreOrder()) Then GoTo err_ScoreIDs
Set qsd = Nothing

If IDCnt > MaxIDToUse Then
   For I = 0 To MaxIDToUse - 1
       IDState(ScoreOrder(I)) = 0
   Next I
   For I = MaxIDToUse To IDCnt - 1
       IDState(ScoreOrder(I)) = IDState(ScoreOrder(I)) + STATE_OUTSCORED
   Next I
Else
   For I = 0 To IDCnt - 1
       IDState(ScoreOrder(I)) = 0
   Next I
End If
ScoreIDs = True
Exit Function

err_ScoreIDs:
LogErrors Err.Number, "frmSearchForNETAdjustment.ScoreIDs"
End Function

Private Sub txtNETTol_LostFocus()
Dim tmp As String
tmp = Trim$(txtNETTol.Text)
If IsNumeric(tmp) Then
   If tmp >= 0 And tmp <= 1 Then
      ThisDef.NETTol = CDbl(tmp)
      Exit Sub
   End If
End If
MsgBox "NET Tolerance should be a number between 0 and 1.", vbOKOnly, glFGTU
txtNETTol.SetFocus
End Sub

Private Function GetNET(ByVal ScanNumber) As Double
'-------------------------------------------------------
'returns generic NET for scan number-this should lead to
'solid results in conjuction with GANET approach
'-------------------------------------------------------
On Error Resume Next
GetNET = (ScanNumber - ScanMin) / ScanRange
End Function

Private Sub CalculateSlopeIntercept()
'-----------------------------------------------------
'least square method to lay best straight line through
'set of points (xi,yi)
'-----------------------------------------------------
Dim SumY As Double
Dim SumX As Double
Dim SumXY As Double
Dim SumXX As Double
Dim I As Long
UpdateStatus "Calculating Slope & Intercept"
SumY = 0
SumX = 0
SumXY = 0
SumXX = 0
' Loop through all the selected identifications
For I = 0 To IDCnt - 1
    SumX = SumX + IDScan(I)
    SumY = SumY + AMTData(ID(I)).NET
    SumXY = SumXY + IDScan(I) * AMTData(ID(I)).NET
    SumXX = SumXX + CDbl(IDScan(I)) ^ 2
Next I
AdjSlp = (IDCnt * SumXY - SumX * SumY) / (IDCnt * SumXX - SumX * SumX)
AdjInt = (SumY - AdjSlp * SumX) / IDCnt
Call CalculateAvgDev
UpdateStatus ""
Exit Sub


err_CalculateSlopeIntercept:
UpdateStatus "Error calculating slope and intercept."
AdjSlp = 0
AdjInt = 0
AdjAvD = -1
End Sub


Private Sub CalculateAvgDev()
Dim I As Long
Dim TtlDist As Double
On Error GoTo err_CalculateAvgDev
TtlDist = 0
For I = 0 To IDCnt - 1
    TtlDist = TtlDist + (AdjSlp * IDScan(I) + AdjInt - AMTData(ID(I)).NET) ^ 2
Next I
AdjAvD = TtlDist / IDCnt
Exit Sub

err_CalculateAvgDev:
AdjAvD = -1
End Sub

Private Sub PickParameters()
Call txtEliminateHowLong_LostFocus
Call txtMaxIDToUse_LostFocus
Call txtMWTol_LostFocus
Call txtNETTol_LostFocus
End Sub
