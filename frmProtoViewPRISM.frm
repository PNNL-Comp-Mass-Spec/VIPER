VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmProtoViewPRISM 
   Caption         =   "Protein Based Analysis - PRISM"
   ClientHeight    =   8670
   ClientLeft      =   2775
   ClientTop       =   4050
   ClientWidth     =   9555
   Icon            =   "frmProtoViewPRISM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   578
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   637
   StartUpPosition =   1  'CenterOwner
   Begin VIPER.LaSpots LaSpots1 
      Height          =   4095
      Left            =   120
      TabIndex        =   14
      Top             =   4440
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   7223
   End
   Begin VB.ListBox lstPeptides 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4050
      Left            =   5880
      TabIndex        =   7
      Top             =   4560
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Peptide Search/Significance"
      Height          =   615
      Left            =   6240
      TabIndex        =   8
      Top             =   0
      Width           =   3015
      Begin VB.TextBox txtPSNET 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2400
         TabIndex        =   12
         Text            =   "5"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtPSMMA 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Text            =   "5"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "NET(pct):"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   10
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "MW(ppm):"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   9000
      Top             =   2280
   End
   Begin VB.TextBox txtORFNavPos 
      Height          =   315
      Left            =   900
      TabIndex        =   6
      Text            =   "0"
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdORFNavigate 
      Caption         =   ">>"
      Height          =   315
      Index           =   3
      Left            =   2730
      TabIndex        =   5
      ToolTipText     =   "Last"
      Top             =   120
      Width           =   390
   End
   Begin VB.CommandButton cmdORFNavigate 
      Caption         =   ">"
      Height          =   315
      Index           =   2
      Left            =   2340
      TabIndex        =   4
      ToolTipText     =   "Next"
      Top             =   120
      Width           =   390
   End
   Begin VB.CommandButton cmdORFNavigate 
      Caption         =   "<"
      Height          =   315
      Index           =   1
      Left            =   510
      TabIndex        =   3
      ToolTipText     =   "Last"
      Top             =   120
      Width           =   390
   End
   Begin VB.CommandButton cmdORFNavigate 
      Caption         =   "<<"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "First"
      Top             =   120
      Width           =   390
   End
   Begin RichTextLib.RichTextBox rtbORFSeq 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   6588
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmProtoViewPRISM.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtResults 
      Height          =   3735
      Left            =   5880
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   60
      Width           =   2655
   End
   Begin VB.Menu mnuF 
      Caption         =   "&File"
      Begin VB.Menu mnuMTS 
         Caption         =   "MT Tags Selection"
         Begin VB.Menu mnuMTSOptions 
            Caption         =   "All ORF MT Tags"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu mnuMTSOptions 
            Caption         =   "Loaded MT Tags"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuMTSOptions 
            Caption         =   "Local Digestion"
            Enabled         =   0   'False
            Index           =   2
         End
      End
      Begin VB.Menu mnuFNET 
         Caption         =   "NET Calculation"
         Begin VB.Menu mnuFNETGeneric 
            Caption         =   "Generic NET"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuFNETTIC 
            Caption         =   "TIC NET"
         End
         Begin VB.Menu mnuFNETGANET 
            Caption         =   "GANET"
         End
      End
      Begin VB.Menu mnuFSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuV 
      Caption         =   "&View"
      Begin VB.Menu mnuVAll 
         Caption         =   "&All"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuVPeptides 
         Caption         =   "&Peptides"
      End
      Begin VB.Menu mnuVSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVFixedWindow 
         Caption         =   "Fixed View Window"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuVVariableWindow 
         Caption         =   "Variable View Window"
      End
      Begin VB.Menu mnuVSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVSeqOnly 
         Caption         =   "Sequences Only"
      End
   End
   Begin VB.Menu mnuR 
      Caption         =   "&Report"
   End
   Begin VB.Menu mnuT 
      Caption         =   "&Tools"
      Begin VB.Menu mnuTSearchAll 
         Caption         =   "Search All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuTSearchCurr 
         Caption         =   "Search Current View"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuTRndORFs 
         Caption         =   "Random ORFs"
      End
   End
End
Attribute VB_Name = "frmProtoViewPRISM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'NOTE: proteins are loaded from database; it would be relatively
'simple to extend this functionality to FASTA file  but hey ....
'NOTE: Peptide Significance(PS) is defined as product of number
'      of peptides pointing to ORF with number of peptides pointing
'      to different Proteins within defined mass/time range
'      In this instance peptide significance is defined on a set of
'      loaded MT tags
'created: 08/01/2002 nt
'last modified: 08/26/2002 nt
'------------------------------------------------------------------
Option Explicit

Const NAV_FIRST = 0
Const NAV_PREVIOUS = 1
Const NAV_NEXT = 2
Const NAV_LAST = 3
Const NAV_REC_NUMBER = 4
Const NAV_STAY = 21

Const OPT_PEPT_DB_ALL = 0
Const OPT_PEPT_DB_LOADED = 1
Const OPT_PEPT_DIGESTION = 2

Const NET_GENERIC = 0
Const NET_TIC = 1
Const NET_GANET = 2

Const MNG_ARRAY_REDIM = 0                       'redim on specified number
Const MNG_ARRAY_REDIM_PRESERVE = 1              'redim on specified number with preserve
Const MNG_ARRAY_REDIM_PRESERVE_ADD = 2          'add specified number with preserve
Const MNG_ARRAY_TRIM = 3                        'trim array on current count
Const MNG_ARRAY_ERASE = 4                       'destroy array

Dim CallerID As Long
Dim bLoading As Boolean

Dim CurrORFInd As Long          'index in RefID array

Dim PeptOption As Long          'option that determines which peptides are loaded

Dim PeptCnt As Long             'count of peptides for current protein
' MonroeMod: Changed to String
Dim PeptID() As String
Dim PeptMW() As Double
Dim PeptNET() As Double
' MonroeMod
Dim PeptIntensity() As Double
Dim PeptPS() As Double

Dim AMTIDCol As Collection      'used for search of loaded MT tags based on ID

'folowing variables are used for various views (all or peptides only)
Dim ListTopPosViewAll As Long
Dim ListTopPosPeptOnly As Long
Dim SpotsTopPosViewAll As Long
Dim SpotsTopPosPeptOnly As Long
Dim HViewAll As Long
Dim HViewPeptOnly As Long

'parameters used for peptide significance calculation and database search
Dim psMMA As Double             'mass measurement accuracy
Dim psNET As Double             'elution time tolerance
Dim NETType As Long
Dim NETSlope As Double
Dim NETIntercept As Double

'database search results; associations between peptides of current ORF
'and peaks of the 2D display
Dim CurrHitsCnt As Long
Dim CurrHitsPeakType() As Long
Dim CurrHitsPeakInd() As Long


Dim MWSearchObject As New MWUtil    'search object for fast MW search


Private Sub cmdORFNavigate_Click(Index As Integer)
Select Case Index
Case NAV_FIRST
     CurrORFInd = 1
Case NAV_PREVIOUS
     If CurrORFInd > 1 Then CurrORFInd = CurrORFInd - 1
Case NAV_NEXT
     If CurrORFInd < ORFCnt Then CurrORFInd = CurrORFInd + 1
Case NAV_LAST
     CurrORFInd = ORFCnt
Case Else                       'don't go anywhere
End Select
txtORFNavPos.Text = " ORF: " & CurrORFInd & "/" & ORFCnt
If mnuVSeqOnly.Checked Then
   rtbORFSeq.Text = GetORFSequence(CallerID, ORFID(CurrORFInd))
Else
   rtbORFSeq.Text = GetORFRecord(CallerID, ORFID(CurrORFInd))
End If
PeptCnt = 0
ReDim PeptID(1 To 1000)      'will be plenty for most proteins
ReDim PeptMW(1 To 1000)
ReDim PeptNET(1 To 1000)
ReDim PeptPS(1 To 1000)
lstPeptides.Clear
LaSpots1.ClearGraphAndData
Select Case PeptOption
Case OPT_PEPT_DB_ALL
     Call GetPeptidesDBAll
Case OPT_PEPT_DB_LOADED
     Call GetPeptidesDBLoaded
Case OPT_PEPT_DIGESTION
     Call GetPeptidesDigestion
End Select
End Sub

Private Sub Form_Activate()
'------------------------------------------------------------
'load MT tags and ORF databases data if neccessary
'------------------------------------------------------------
On Error Resume Next
CallerID = Me.Tag
Me.MousePointer = vbHourglass
If bLoading Then
   UpdateStatus "Loading ORF list ..."
   SetNETSlopeIntercept
   LoadORFs CallerID
   UpdateStatus "Number of loaded Proteins: " & ORFCnt
   UpdateStatus "Loading MT tag - Protein mappings ..."
   LoadMassTagToProteinMapping Me, CallerID, True
   UpdateStatus "Proteins: " & ORFCnt & " MT tags: " & AMTCnt & " Mappings: " & MTtoORFMapCount
   bLoading = False
   Call cmdORFNavigate_Click(0)
End If
Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim i As Long
bLoading = True
ListTopPosViewAll = lstPeptides.Top
ListTopPosPeptOnly = rtbORFSeq.Top
SpotsTopPosViewAll = LaSpots1.Top
SpotsTopPosPeptOnly = rtbORFSeq.Top - Abs(ListTopPosViewAll - SpotsTopPosViewAll)
HViewAll = Me.Height
HViewPeptOnly = Me.Height - (SpotsTopPosViewAll - SpotsTopPosPeptOnly) * Screen.TwipsPerPixelY
If AMTCnt > 0 Then
   PeptOption = OPT_PEPT_DB_LOADED         'select ORF peptides among loaded MT tags
   'add MT tags indexes in collection with ID as an key so that we can access them fast
   Set AMTIDCol = New Collection
   For i = 1 To AMTCnt
       AMTIDCol.add i, AMTData(i).ID
   Next i
Else                                       'nothing is loaded so that is not an option
   PeptOption = OPT_PEPT_DIGESTION
   mnuMTSOptions(OPT_PEPT_DIGESTION).Checked = True
End If
psMMA = CDbl(txtPSMMA.Text) * glPPM
psNET = CDbl(txtPSNET.Text) * glPCT
End Sub

Private Sub UpdateStatus(ByVal Msg As String)
lblStatus.Caption = Msg
DoEvents
End Sub

Private Sub lstPeptides_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
LaSpots1.ToggleSpotSelection lstPeptides.ListIndex
End Sub

Private Sub lstPeptides_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
LaSpots1.ToggleSpotSelection lstPeptides.ListIndex
End Sub

Private Sub mnuF_Click()
PickParameters
End Sub

Private Sub mnuFClose_Click()
Unload Me
End Sub

Private Sub mnuFNETGANET_Click()
NETType = NET_GANET
mnuFNETGeneric.Checked = False
mnuFNETTIC.Checked = False
mnuFNETGANET.Checked = True
End Sub

Private Sub mnuFNETGeneric_Click()
NETType = NET_GENERIC
mnuFNETGeneric.Checked = True
mnuFNETTIC.Checked = False
mnuFNETGANET.Checked = False
End Sub

Private Sub mnuFNETTIC_Click()
NETType = NET_TIC
mnuFNETGeneric.Checked = False
mnuFNETTIC.Checked = True
mnuFNETGANET.Checked = False
End Sub

Private Sub mnuMTSOptions_Click(Index As Integer)
Dim i As Long
PeptOption = Index
For i = mnuMTSOptions.LBound To mnuMTSOptions.UBound
    If i = PeptOption Then
       mnuMTSOptions(i).Checked = True
    Else
       mnuMTSOptions(i).Checked = False
    End If
Next i
End Sub

Private Sub mnuR_Click()
PickParameters
End Sub

Private Sub mnuT_Click()
PickParameters
End Sub

Private Sub mnuTRndORFs_Click()
If mnuTRndORFs.Checked Then
   mnuTRndORFs.Checked = False
   Timer1.Enabled = False
Else
   Randomize (CurrORFInd)
   mnuTRndORFs.Checked = True
   Timer1.Enabled = True
End If
End Sub

Private Sub mnuTSearchAll_Click()
Dim CurrMW As Double
Dim CurrNET As Double
Dim i As Long
If PrepareSearchObject() Then
   If ManageHitsArray(MNG_ARRAY_REDIM, 1000) Then
      With GelData(CallerID)
        For i = 1 To .CSLines
            CurrMW = .CSData(i).AverageMW
            CurrNET = CalcNET(.CSData(i).ScanNumber)
        Next i
        For i = 1 To .IsoLines
            CurrMW = GetIsoMass(.IsoData(i), .Preferences.IsoDataField)
            CurrNET = CalcNET(.IsoData(i).ScanNumber)
        Next i
      End With
   End If
Else
   If PeptCnt > 0 Then      'something is wrong
      txtResults.Text = "Error preparing sort objects."
   Else
      txtResults.Text = "No peptides for current ORF found."
   End If
End If
End Sub

Private Sub mnuV_Click()
PickParameters
End Sub

Private Sub mnuVAll_Click()
On Error Resume Next
mnuVAll.Checked = True
mnuVPeptides.Checked = False
lstPeptides.Top = ListTopPosViewAll
LaSpots1.Top = SpotsTopPosViewAll
Me.Height = HViewAll
End Sub

Private Sub mnuVFixedWindow_Click()
mnuVFixedWindow.Checked = True
mnuVVariableWindow.Checked = False
LaSpots1.ViewWindow = sFixedWindow
End Sub

Private Sub mnuVPeptides_Click()
On Error Resume Next
mnuVAll.Checked = False
mnuVPeptides.Checked = True
lstPeptides.Top = ListTopPosPeptOnly
LaSpots1.Top = SpotsTopPosPeptOnly
Me.Height = HViewPeptOnly
End Sub

Private Sub mnuVSeqOnly_Click()
mnuVSeqOnly.Checked = Not mnuVSeqOnly.Checked
Call cmdORFNavigate_Click(NAV_STAY)
End Sub

Private Sub mnuVVariableWindow_Click()
mnuVFixedWindow.Checked = False
mnuVVariableWindow.Checked = True
LaSpots1.ViewWindow = sVariableWindow
End Sub


Private Sub GetPeptidesDBAll()

End Sub

Private Sub GetPeptidesDBLoaded()
Dim i As Long
Dim TmpMW As Double
Dim IsPeptLoaded As Boolean
Dim ListLabel As String
On Error GoTo err_GetPeptidesDBLoaded

'find all peptides among loaded MT tags mapped with current ORF - they don't have to be
'among loaded MT tags - therefore call to retrieve it from AMTIDCol could fail (RTE - 5)
For i = 1 To MTtoORFMapCount
    If ORFIDMap(i) = ORFID(CurrORFInd) Then
       IsPeptLoaded = True                                   'assume peptide is loaded
       TmpMW = AMTData(AMTIDCol.Item(CStr(MTIDMap(i)))).MW   'if not this will fail
       If IsPeptLoaded Then
          PeptCnt = PeptCnt + 1
          PeptID(PeptCnt) = MTIDMap(i)
          PeptMW(PeptCnt) = TmpMW
          PeptNET(PeptCnt) = AMTData(AMTIDCol.Item(CStr(MTIDMap(i)))).NET
       End If
    End If
Next i
If PeptCnt > 0 Then
   If PeptCnt < UBound(PeptID) Then
      ReDim Preserve PeptID(1 To PeptCnt)
      ReDim Preserve PeptMW(1 To PeptCnt)
' MonroeMod
      ReDim Preserve PeptIntensity(1 To PeptCnt)
      ReDim Preserve PeptNET(1 To PeptCnt)
      ReDim Preserve PeptPS(1 To PeptCnt)
   End If
   'calculate peptide significances and make list entries
   Dim MWFmt As String
   Dim MWSpc As String
   For i = 1 To PeptCnt
       PeptPS(i) = CalcMTPS(i)
       MWFmt = Format$(PeptMW(i), "0.00")
       If Len(MWFmt) > 7 Then
          MWSpc = ""
       Else
          MWSpc = String(8 - Len(MWFmt), Chr$(32))
       End If
       ListLabel = Format$(i, "@@@@") & " " & Format$(PeptID(i), "@@@@@@@@") _
                   & " " & MWSpc & MWFmt & "Da " & Format$(PeptPS(i), "0.0000")
       lstPeptides.AddItem ListLabel
   Next i
   'update picture
' MonroeMod: Added PeptIntensity()
   If Not LaSpots1.AddSpotsMany(PeptID(), PeptNET(), PeptMW(), PeptIntensity()) Then
      UpdateStatus "Error presenting protein peptides pattern."
   End If
Else
   LaSpots1.ClearGraphAndData
   ' MonroeMod: Changed RefreshSpots to RefreshPlot
   LaSpots1.RefreshPlot
   Erase PeptID
   Erase PeptMW
   Erase PeptNET
   Erase PeptPS
End If
Exit Sub

err_GetPeptidesDBLoaded:
Select Case Err.Number
Case 5                                      'failed call to procedure(see note above)
   Err.Clear
   IsPeptLoaded = False
   Resume Next
Case 9                                      'add more room in peptide arrays and continue
   Err.Clear
   ReDim Preserve PeptID(1 To PeptCnt + 1000)
   ReDim Preserve PeptMW(1 To PeptCnt + 1000)
   ReDim Preserve PeptNET(1 To PeptCnt + 1000)
   ReDim Preserve PeptPS(1 To PeptCnt + 1000)
   Resume
End Select
End Sub

Private Sub GetPeptidesDigestion()

End Sub

Private Function LoadORFs(ByVal Ind As Long) As Boolean
    '--------------------------------------------------------------
    'executes command that retrieves list of Proteins (ORFs) from Organism
    'MT tag database; returns True if at least one ORF loaded.
    '--------------------------------------------------------------
    Dim cnNew As New ADODB.Connection
    Dim sCommand As String
    Dim rsORFs As New ADODB.Recordset
    Dim cmdGetORFs As New ADODB.Command     'no arguments with ORF list
    
    'reserve space for 50000 ORFs; increase in chunks of 2000 after that
    ReDim ORFID(1 To 50000)
    
    On Error Resume Next
    sCommand = glbPreferencesExpanded.MTSConnectionInfo.spGetORFIDs
    If Len(sCommand) <= 0 Then Exit Function
    
    On Error GoTo err_LoadORFs
    Screen.MousePointer = vbHourglass
    ORFCnt = 0
    
    If Not EstablishConnection(cnNew, GelAnalysis(Ind).MTDB.cn.ConnectionString, False) Then
        Debug.Assert False
        LoadORFs = False
        Exit Function
    End If
    
    'create and tune command object to retrieve MT tags
    ' Initialize the SP
    InitializeSPCommand cmdGetORFs, cnNew, sCommand
    
    'procedure returns error number or 0 if OK
    Set rsORFs = cmdGetORFs.Execute
    With rsORFs
        Do Until .EOF
           ORFCnt = ORFCnt + 1
           ORFID(ORFCnt) = .Fields(0).Value
           .MoveNext
        Loop
    End With
    rsORFs.Close
    
    'clean things and exit
exit_LoadORFs:
    On Error Resume Next
    Set cmdGetORFs.ActiveConnection = Nothing
    cnNew.Close
    If ORFCnt > 0 Then
       If ORFCnt < UBound(ORFID) Then ReDim Preserve ORFID(1 To ORFCnt)
    Else
       Erase ORFID
    End If
    Screen.MousePointer = vbDefault
    LoadORFs = (ORFCnt > 0)
    Exit Function

err_LoadORFs:
    Select Case Err.Number
    Case 9                       'need more room for MT tags
        ReDim Preserve ORFID(1 To ORFCnt + 2000)
        Resume
    Case 13, 94                  'Type Mismatch or Invalid Use of Null
        Resume Next              'just ignore it
    Case 3265, 3704              'two errors I have encountered
        '2nd attempt will probably work so let user know it should try again
        If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
            MsgBox "Error loading Proteins from the database. Error could " _
                 & "have been caused by network/server issues(timeout) so you " _
                 & "might try loading again with Refresh function.", vbOKOnly, glFGTU
        End If
    Case Else
        LogErrors Err.Number, "LoadORFs"
    End Select
    ORFCnt = -1
    GoTo exit_LoadORFs
End Function

Private Sub Timer1_Timer()
CurrORFInd = CLng(Rnd() * ORFCnt)
cmdORFNavigate_Click (NAV_STAY)
End Sub

Private Sub txtORFNavPos_GotFocus()
'select everything so that it is easier to type number
txtORFNavPos.SelStart = 1
txtORFNavPos.SelLength = Len(txtORFNavPos.Text)
End Sub

Private Sub txtORFNavPos_LostFocus()
'-----------------------------------------------------
'user can type record number to which it wants to jump
'-----------------------------------------------------
Dim TargetRecInd As Long
On Error Resume Next
If IsNumeric(txtORFNavPos.Text) Then
   TargetRecInd = CLng(txtORFNavPos.Text)
   If TargetRecInd < 1 Then TargetRecInd = 1
   If TargetRecInd > ORFCnt Then TargetRecInd = ORFCnt
   CurrORFInd = TargetRecInd
End If
Call cmdORFNavigate_Click(NAV_STAY)
End Sub

Private Function CalcMTPS(ByVal Ind As Long) As Double
'-------------------------------------------------------------
'calculates significance of peptide for protein identification
'-------------------------------------------------------------
Dim AbsPeptMMA As Double
Dim ClosePeptCnt As Long
Dim CloseHits() As Long
On Error GoTo err_CalcMTPS
AbsPeptMMA = PeptMW(Ind) * psMMA
ClosePeptCnt = GetMTHits1(PeptMW(Ind), AbsPeptMMA, PeptNET(Ind), psNET, CloseHits())
If ClosePeptCnt > 0 Then
   CalcMTPS = 1 / (ClosePeptCnt * PeptCnt)
   Exit Function
End If

err_CalcMTPS:
CalcMTPS = -1
End Function


Private Sub txtPSMMA_LostFocus()
On Error Resume Next
If IsNumeric(txtPSMMA.Text) Then
   psMMA = Abs(CDbl(txtPSMMA.Text)) * glPPM
   txtPSMMA.Text = psMMA / glPPM
Else
   MsgBox "This argument should be numeric.", vbOKOnly, glFGTU
   txtPSMMA.SetFocus
End If
End Sub


Private Sub txtPSNET_LostFocus()
On Error Resume Next
If IsNumeric(txtPSNET.Text) Then
   psNET = Abs(CDbl(txtPSNET.Text)) * glPCT
   txtPSNET.Text = psNET / glPCT
Else
   MsgBox "This argument should be numeric.", vbOKOnly, glFGTU
   txtPSNET.SetFocus
End If
End Sub

' Unused Function (May 2003)
'''Private Function SearchCurrentORF() As Long
''''search 2D display for selected peptides of current ORF; returns
''''number of hits; -1 on any error
'''End Function

Private Function ManageHitsArray(ByVal ManageType As Long, ByVal Count As Long) As Boolean
On Error GoTo exit_ManageHitsArray
Select Case ManageType
Case MNG_ARRAY_REDIM                        'redim on specified number
     CurrHitsCnt = 0
     ReDim CurrHitsPeptInd(Count - 1)
     ReDim CurrHitsPeakType(Count - 1)
     ReDim CurrHitsPeakInd(Count - 1)
Case MNG_ARRAY_REDIM_PRESERVE               'redim on specified number with preserve
     ReDim Preserve CurrHitsPeptInd(Count - 1)
     ReDim Preserve CurrHitsPeakType(Count - 1)
     ReDim Preserve CurrHitsPeakInd(Count - 1)
Case MNG_ARRAY_REDIM_PRESERVE_ADD           'add specified number with preserve
     ReDim Preserve CurrHitsPeptInd(CurrHitsCnt + Count - 1)
     ReDim Preserve CurrHitsPeakType(CurrHitsCnt + Count - 1)
     ReDim Preserve CurrHitsPeakInd(CurrHitsCnt + Count - 1)
Case MNG_ARRAY_TRIM                         'trim array on current count
     ReDim Preserve CurrHitsPeptInd(CurrHitsCnt - 1)
     ReDim Preserve CurrHitsPeakType(CurrHitsCnt - 1)
     ReDim Preserve CurrHitsPeakInd(CurrHitsCnt - 1)
Case MNG_ARRAY_ERASE                        'destroy array
     CurrHitsCnt = 0
     Erase CurrHitsPeptInd
     Erase CurrHitsPeakType
     Erase CurrHitsPeakInd
End Select
ManageHitsArray = True
exit_ManageHitsArray:
End Function

Public Function PrepareSearchObject() As Boolean
'----------------------------------------------------------------
'prepares fast search object for current set of peptides
'----------------------------------------------------------------
Dim TmpMW() As Double
Dim MWInd() As Long
Dim qsdSort As New QSDouble
Dim i As Long
On Error GoTo exit_PrepareSearchObject
Set MWSearchObject = Nothing
If PeptCnt > 0 Then
   ReDim TmpMW(PeptCnt - 1)
   ReDim MWInd(PeptCnt - 1)
   For i = 0 To PeptCnt - 1
       MWInd(i) = i
       TmpMW(i) = PeptMW(i)
   Next i
   If qsdSort.QSAsc(TmpMW(), MWInd()) Then
      If MWSearchObject.Fill(TmpMW()) Then PrepareSearchObject = True
   End If
End If
exit_PrepareSearchObject:
End Function

Private Sub PickParameters()
Call txtPSMMA_LostFocus
Call txtPSNET_LostFocus
End Sub

Private Function CalcNET(ByVal ScanNum As Long) As Double
CalcNET = NETSlope * ScanNum + NETIntercept
End Function

Private Sub SetNETSlopeIntercept()
Dim FirstScan As Long
Dim LastScan As Long
Dim ScanRange As Long
On Error GoTo err_SetNETSlopeIntercept
Select Case NETType
Case NET_GENERIC
   GetScanRange CallerID, FirstScan, LastScan, ScanRange
   NETSlope = 1 / (LastScan - FirstScan)
   NETIntercept = -FirstScan / (LastScan - FirstScan)
Case NET_TIC
   NETSlope = GelAnalysis(CallerID).NET_Slope
   NETIntercept = GelAnalysis(CallerID).NET_Intercept
Case NET_GANET
   NETSlope = GelAnalysis(CallerID).GANET_Slope
   NETIntercept = GelAnalysis(CallerID).GANET_Intercept
End Select

err_SetNETSlopeIntercept:
'if it fails it probably failed because it is not connected to DB
If NETType <> NET_GENERIC Then
   Call mnuFNETGeneric_Click
Else
   txtResults.Text = "Error selecting NET calculation formula."
End If
End Sub
