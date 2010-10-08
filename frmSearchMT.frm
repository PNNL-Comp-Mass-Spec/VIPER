VERSION 5.00
Begin VB.Form frmSearchMT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search MT Tag Database"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraMisc 
      Caption         =   "Miscellaneous"
      Height          =   1215
      Left            =   120
      TabIndex        =   32
      Top             =   3240
      Width           =   4695
      Begin VB.CheckBox chkSaveNCnt 
         Caption         =   "Store N-atoms count with MT Tag reference"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   720
         Width           =   3615
      End
      Begin VB.CheckBox chkSkipAMTReferenced 
         Caption         =   "S&kip records already MT referenced"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Value           =   1  'Checked
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdCheckMTDB 
      Caption         =   "C&heck MTDB"
      Height          =   375
      Left            =   1800
      TabIndex        =   31
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Frame fraMassTag 
      Caption         =   "MT Tag"
      Height          =   1215
      Left            =   4920
      TabIndex        =   27
      Top             =   3240
      Width           =   1455
      Begin VB.TextBox txtMTNum 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   600
         TabIndex        =   29
         Text            =   "5"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtMassTag 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   28
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Max."
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   780
         Width           =   345
      End
   End
   Begin VB.Frame fraNET 
      Caption         =   "Elution  Calculation"
      Height          =   1335
      Left            =   120
      TabIndex        =   19
      Top             =   1800
      Width           =   6255
      Begin VB.CommandButton cmdETType 
         Caption         =   "G&ANET"
         Height          =   315
         Index           =   2
         Left            =   3720
         TabIndex        =   37
         ToolTipText     =   "Sets GANET formula to be used with elution calculation"
         Top             =   900
         Width           =   855
      End
      Begin VB.CommandButton cmdETType 
         Caption         =   "&TIC Fit"
         Height          =   315
         Index           =   1
         Left            =   2640
         TabIndex        =   36
         ToolTipText     =   "Sets TIC Fit NET formula to be used with elution calculation"
         Top             =   900
         Width           =   855
      End
      Begin VB.CommandButton cmdETType 
         Caption         =   "&Generic"
         Height          =   315
         Index           =   0
         Left            =   1560
         TabIndex        =   35
         ToolTipText     =   "Sets generic NET formula to be used with elution calculation"
         Top             =   900
         Width           =   855
      End
      Begin VB.OptionButton optNETorRT 
         Caption         =   "First choice"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   25
         ToolTipText     =   "Use NET calculated only from Sequest ""first choice"" peptides"
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton optNETorRT 
         Caption         =   "All results"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   24
         ToolTipText     =   "Use NET calculated from all peptides of MT Tags"
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox txtNETFormula 
         Height          =   285
         Left            =   1560
         TabIndex        =   23
         Text            =   "(FN-MinFN)/(MaxFN-MinFN)"
         Top             =   540
         Width           =   3015
      End
      Begin VB.TextBox txtNETTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4800
         TabIndex        =   21
         Text            =   "0.1"
         Top             =   540
         Width           =   615
      End
      Begin VB.Label lblETType 
         Height          =   255
         Left            =   4800
         TabIndex        =   38
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblSec 
         Caption         =   "sec."
         Height          =   255
         Left            =   5520
         TabIndex        =   26
         ToolTipText     =   "Normalized Elution Time Tolerance (if blank NET will not be considered in search)"
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "&Formula  F(FN, MinFN, MaxFN)"
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Top             =   285
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "T&olerance"
         Height          =   255
         Left            =   4800
         TabIndex        =   20
         ToolTipText     =   "Normalized Elution Time Tolerance (if blank NET will not be considered in search)"
         Top             =   285
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdSearchResults 
      Caption         =   "S&earch Results"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      ToolTipText     =   "Display search statistic for MT database"
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "D&elete MT Ref."
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      ToolTipText     =   "Remove MT reference for current gel"
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Frame fraMWTolerance 
      Caption         =   "Molecular Mass Tolerance"
      Height          =   1215
      Left            =   4200
      TabIndex        =   14
      Top             =   480
      Width           =   2175
      Begin VB.OptionButton optTolType 
         Caption         =   "&Dalton"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   18
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optTolType 
         Caption         =   "&ppm"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   17
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtMWTol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   160
         TabIndex        =   16
         Text            =   "10"
         Top             =   640
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Tolerance"
         Height          =   255
         Left            =   160
         TabIndex        =   15
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame fraMWField 
      Caption         =   "Molecular Mass Field"
      Height          =   1215
      Left            =   2160
      TabIndex        =   10
      Top             =   480
      Width           =   1935
      Begin VB.OptionButton optMWField 
         Caption         =   "&The Most Abundant"
         Height          =   255
         Index           =   2
         Left            =   80
         TabIndex        =   13
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton optMWField 
         Caption         =   "&Monoisotopic"
         Height          =   255
         Index           =   1
         Left            =   80
         TabIndex        =   12
         Top             =   540
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optMWField 
         Caption         =   "A&verage"
         Height          =   255
         Index           =   0
         Left            =   80
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame fraSearchScope 
      Caption         =   "Search Scope"
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   1935
      Begin VB.OptionButton optSearchScope 
         Caption         =   "C&urrent View"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton optSearchScope 
         Caption         =   "&All Data Points"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh from DB"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      ToolTipText     =   "Refresh connection to the MT database"
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblAMTStatus 
      Caption         =   "Loading..."
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "MTDB Status:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Status of the MT Tag database"
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmSearchMT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'search of MT tag database; this is just a patch
'until new procedures are developed; it works OK.
'--------------------------------------------------
'last modified: 05/24/2002 nt
'--------------------------------------------------
Option Explicit

'in this case CallerID is a public property
Public CallerID As Long

Dim bLoading As Boolean

Dim OldSearchFlag As Long

Public Sub InitializeDBSearch(Optional blnForceReload As Boolean = False)
    Dim blnAMTsWereLoaded As Boolean, blnDBConnectionError As Boolean

    If bLoading Then
        If ConfirmMassTagsAndInternalStdsLoaded(Me, CallerID, True, True, False, blnForceReload, 0, blnAMTsWereLoaded, blnDBConnectionError) Then
            lblAMTStatus.Caption = ConstructMTStatusText(True)
            cmdSearch.Enabled = True
        Else
            If blnDBConnectionError Then
                lblAMTStatus.Caption = "Error loading MT tags: database connection error."
            Else
                lblAMTStatus.Caption = "Error loading MT tags: no valid MT tags were found (possibly missing NET values)"
            End If
            cmdSearch.Enabled = False
        End If
        
        ' MonroeMod: Only show the "Seconds" label if samtDef.NETorRT <> 0
        lblSec.Visible = CBool(samtDef.NETorRT)
        Call cmdETType_Click(etGANET)
        
        bLoading = False
    End If
End Sub

Public Function ShowOrSaveResults(Optional strOutputFilePath As String = "", Optional blnDisplayResults As Boolean = True) As Long
' If strOutputFilePath = "", then saves the results to a temporary file and shows them to the user using frmDataInfo
' If strOutputFilePath is not blank, then saves the results to the file, but does not display them
'
' Returns 0 if no error, the error number if an error
    
Dim FileNum As Integer
Dim FileName As String
Dim sLine As String
Dim AvgErrDa As String
Dim AvgErrppm As String
Dim AvgErrNET As String
Dim i As Long
Dim strSepChar As String

' MonroeMod
Dim strCaptionSaved As String
Dim lngHitCountTotal As Long

On Error GoTo exit_cmdSearchResults_Click

If AMTCnt > 0 Then  'global AMT count
       
    ' MonroeMod Begin
    strCaptionSaved = Me.Caption
    Me.Caption = "Preparing search results for display"
    
    strSepChar = LookupDefaultSeparationCharacter()
    
    For i = 1 To AMTCnt
       lngHitCountTotal = lngHitCountTotal + AMTHits(i)
    Next i
    ' MonroeMod Finish
   
    ' MonroeFix
   On Error Resume Next
    ReDim Preserve AMTHits(AMTCnt)
   On Error GoTo 0
   
   Me.MousePointer = vbHourglass
   FileNum = FreeFile()
   If Len(strOutputFilePath) > 0 Then
      FileName = strOutputFilePath
   Else
      FileName = GetTempFolder() & RawDataTmpFile
   End If
   Open FileName For Output As FileNum
   'print gel file name and Search definition as reference
   Print #FileNum, "Generated by: " & GetMyNameVersion() & " on " & Now()
   Print #FileNum, "Gel File: " & GelBody(CallerID).Caption
   Print #FileNum, "Total hits (non-unique): " & Trim(lngHitCountTotal)
   Print #FileNum, GetAMTSearchDefDesc()
   
   Select Case samtDef.NETorRT
   Case glAMT_NET
      sLine = "MT ID" & strSepChar & "MT MW" & strSepChar & "MT NET" & strSepChar _
           & "Hits" & strSepChar & "MT MW Avg Error(Da)" & strSepChar & _
           "MT MW Avg Error(ppm)" & strSepChar & "MT NET Avg Error" & _
           strSepChar & "MT Range Min" & strSepChar & "MT Range Max"
   Case glAMT_RT_or_PNET
      sLine = "MT ID" & strSepChar & "MT MW" & strSepChar & "MT RT" & strSepChar _
           & "Hits" & strSepChar & "MT MW Avg Error(Da)" & strSepChar & _
           "MT MW Avg Error(ppm)" & strSepChar & "MT RT Avg Error" & _
           strSepChar & "MT Range Min" & strSepChar & "MT Range Max"
   End Select
   Print #FileNum, sLine
   For i = 1 To AMTCnt
        ' MonroeMod Begin
        If i Mod AMTCnt / 30 = 0 Then
            Me.Caption = Me.Caption & "."
        End If
        ' MonroeMod Finish
        
       Select Case samtDef.NETorRT
       Case glAMT_NET
         sLine = Trim(AMTData(i).ID) & strSepChar & AMTData(i).MW & strSepChar & AMTData(i).NET & strSepChar & AMTHits(i)
       Case glAMT_RT_or_PNET
         sLine = Trim(AMTData(i).ID) & strSepChar & AMTData(i).MW & strSepChar & AMTData(i).PNET & strSepChar & AMTHits(i)
       End Select
       If AMTHits(i) > 0 Then
          AvgErrDa = Str(AMTMWErr(i) / AMTHits(i))
          AvgErrppm = Str(AvgErrDa / (AMTData(i).MW * glPPM))
          AvgErrNET = Str(AMTNETErr(i) / AMTHits(i))
          sLine = sLine & strSepChar & AvgErrDa & strSepChar & AvgErrppm & strSepChar _
              & AvgErrNET & strSepChar & AMTNETMin(i) & strSepChar & AMTNETMax(i)
       End If
       Print #FileNum, sLine
   Next i
   Close FileNum
   If Len(strOutputFilePath) > 0 Then
      AddToAnalysisHistory CallerID, "Saved search results to disk: " & strOutputFilePath
   End If
   DoEvents
   If blnDisplayResults Then
      Me.Caption = strCaptionSaved
      frmDataInfo.Tag = "AMT"
      DoEvents
      frmDataInfo.Show vbModal
   End If
Else
   If blnDisplayResults Then MsgBox "No loaded MT data found.", vbOKOnly
End If

exit_cmdSearchResults_Click:
On Error Resume Next
Close FileNum
Me.MousePointer = vbDefault
    
' MonroeMod
Me.Caption = strCaptionSaved
ShowOrSaveResults = Err.Number
End Function

Public Function StartSearch(Optional blnShowMessages As Boolean = True) As Long
' Returns the number of hits
Dim HitsCnt As Long
On Error Resume Next

'always reinitialize statistics arrays
InitAMTStat
GelData(CallerID).MostRecentSearchUsedSTAC = False
samtDef.Formula = Trim$(txtNETFormula.Text)
Me.MousePointer = vbHourglass
If samtDef.MassTag > 0 Then
   'HitsCnt = SearchAMTWithTag(CallerID, samtDef.Formula)
   HitsCnt = SearchAMTWithTag1(CallerID, samtDef.Formula)
Else
' MonroeMod
   HitsCnt = SearchAMT(CallerID, samtDef.Formula, Me)
End If
GelStatus(CallerID).Dirty = True

'MonroeMod
GelSearchDef(CallerID).AMTSearchOnIons = samtDef
AddToAnalysisHistory CallerID, GetMassTagSearchSummaryText("Searched ions for MT tags", HitsCnt, 0, 0, 0, samtDef, False, GelData(CallerID).CustomNETsDefined)

Me.MousePointer = vbDefault
Select Case HitsCnt
Case -1
   If blnShowMessages Then MsgBox "Error searching MT database.", vbOKOnly
Case -2
   If blnShowMessages Then MsgBox "Error in NET calculation formula.", vbOKOnly
   txtNETFormula.SetFocus
Case Else
   If blnShowMessages Then MsgBox "MT tag hits: " & HitsCnt & " (non-unique)", vbOKOnly
End Select
StartSearch = HitsCnt
End Function

Private Sub chkSaveNCnt_Click()
If chkSaveNCnt.Value = vbChecked Then
   samtDef.SaveNCnt = True
Else
   samtDef.SaveNCnt = False
End If
End Sub

Private Sub chkSkipAMTReferenced_Click()
If chkSkipAMTReferenced.Value = vbChecked Then
   samtDef.SkipReferenced = True
Else
   samtDef.SkipReferenced = False
End If
End Sub

Private Sub cmdCheckMTDB_Click()
'----------------------------------------------
'displays short MT tags statistics, it might
'help with determining problems with MT tags
'----------------------------------------------
Me.MousePointer = vbHourglass
MsgBox CheckMassTags(), vbOKOnly, glFGTU
Me.MousePointer = vbDefault
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdETType_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case etGenericNET
  txtNETFormula.Text = ConstructNETFormula(0, 0, True)
Case etTICFitNET
  With GelAnalysis(CallerID)
    If .NET_Intercept <> 0 Or .NET_Slope <> 0 Then _
       txtNETFormula.Text = ConstructNETFormula(.NET_Slope, .NET_Intercept)
  End With
  If Err Then
     MsgBox "Make sure display is loaded as analysis. Use New Analysis command from the File menu.", vbOKOnly, glFGTU
     Exit Sub
  End If
Case etGANET
  With GelAnalysis(CallerID)
    If .GANET_Intercept <> 0 Or .GANET_Slope <> 0 Then
       txtNETFormula.Text = ConstructNETFormula(.GANET_Slope, .GANET_Intercept)
    Else
        txtNETFormula.Text = ConstructNETFormula(0, 0, True)
    End If
  End With
  If Err Then
     MsgBox "Make sure display is loaded as analysis. Use New Analysis command from the File menu.", vbOKOnly, glFGTU
     Exit Sub
  End If
End Select
lblETType.Caption = "ET: " & cmdETType(Index).Caption
End Sub

Private Sub cmdRefresh_Click()
    Dim blnAMTsWereLoaded As Boolean, blnDBConnectionError As Boolean
    Dim blnForceReload As Boolean
    blnForceReload = True
    
    If ConfirmMassTagsAndInternalStdsLoaded(Me, CallerID, True, True, False, blnForceReload, 0, blnAMTsWereLoaded, blnDBConnectionError) Then
        lblAMTStatus.Caption = "Loaded; MT tag count: " & LongToStringWithCommas(AMTCnt)
        cmdSearch.Enabled = True
    Else
        If blnDBConnectionError Then
            lblAMTStatus.Caption = "Error loading MT tags: database connection error."
        Else
            lblAMTStatus.Caption = "Error loading MT tags: no valid MT tags were found (possibly missing NET values)"
        End If
        cmdSearch.Enabled = False
    End If
End Sub

Private Sub cmdRemove_Click()
Dim eResponse As VbMsgBoxResult
eResponse = MsgBox("Remove MT tag references from the current gel?", vbYesNo + vbDefaultButton1)
If eResponse <> vbYes Then Exit Sub
RemoveAMT CallerID, samtDef.SearchScope
GelStatus(CallerID).Dirty = True
'MonroeMod
AddToAnalysisHistory CallerID, "Deleted MT tag search results from ions"
End Sub

Private Sub cmdSearch_Click()
StartSearch
End Sub

Private Sub cmdSearchResults_Click()
ShowOrSaveResults ""
End Sub

Private Sub Form_Activate()
InitializeDBSearch
End Sub

Private Sub Form_Load()
bLoading = True
If IsWinLoaded(TrackerCaption) Then Unload frmTracker
' MonroeMod
If CallerID >= 1 And CallerID <= UBound(GelBody) Then samtDef = GelSearchDef(CallerID).AMTSearchOnIons
'set current Search Definition values
With samtDef
    txtMWTol.Text = .MWTol
    optSearchScope(.SearchScope).Value = True
    optMWField(.MWField - MW_FIELD_OFFSET).Value = True
    optNETorRT(.NETorRT).Value = True
    Select Case .TolType
    Case gltPPM
      optTolType(0).Value = True
    Case gltABS
      optTolType(1).Value = True
    Case Else
      Debug.Assert False
    End Select
    'save old value and set search on "search all"
    OldSearchFlag = .SearchFlag
    .SearchFlag = 0         'search all
    'NETTol is used both for NET and RT
    If .NETTol >= 0 Then
       txtNETTol.Text = .NETTol
    Else
       txtNETTol.Text = ""
    End If
    If .MassTag > 0 Then
       txtMassTag.Text = .MassTag
    Else
       txtMassTag.Text = ""
    End If
    txtMTNum.Text = .MaxMassTags
    If .SkipReferenced Then
       chkSkipAMTReferenced.Value = vbChecked
    Else
       chkSkipAMTReferenced.Value = vbUnchecked
    End If
    If .SaveNCnt Then
       chkSaveNCnt.Value = vbChecked
    Else
       chkSaveNCnt.Value = vbUnchecked
    End If
End With
DoEvents

End Sub


Private Sub Form_Unload(Cancel As Integer)
DestroyAMTStat
samtDef.SearchFlag = OldSearchFlag
End Sub

Private Sub optMWField_Click(Index As Integer)
samtDef.MWField = 6 + Index
End Sub

Private Sub optNETorRT_Click(Index As Integer)
samtDef.NETorRT = Index
lblSec.Visible = CBool(Index)
End Sub

Private Sub optSearchScope_Click(Index As Integer)
samtDef.SearchScope = Index
End Sub

Private Sub optTolType_Click(Index As Integer)
If Index = 0 Then
   samtDef.TolType = gltPPM
Else
   samtDef.TolType = gltABS
End If
End Sub


Private Sub txtMassTag_LostFocus()
Dim dTmp As Double
If IsNumeric(txtMassTag.Text) Then
   dTmp = CDbl(txtMassTag.Text)
   If dTmp > 0 Then
      samtDef.MassTag = dTmp
   Else
      MsgBox "MT tag should be positive number.", vbOKOnly
      txtMassTag.Text = ""
      txtMassTag.SetFocus
   End If
Else
   If Len(Trim(txtMassTag.Text)) > 0 Then
      MsgBox "MT tag should be positive number.", vbOKOnly
      txtMassTag.SetFocus
   Else
      samtDef.MassTag = -1
   End If
End If
End Sub

Private Sub txtMTNum_LostFocus()
On Error Resume Next
samtDef.MaxMassTags = CLng(Abs(txtMTNum.Text))
If Err Then
   MsgBox "Maximum number of MT tags should be positive integer.", vbOKOnly
   txtMTNum.SetFocus
End If
End Sub

Private Sub txtMWTol_LostFocus()
If IsNumeric(txtMWTol.Text) Then
   samtDef.MWTol = CDbl(txtMWTol.Text)
Else
   MsgBox "Molecular Mass Tolerance should be numeric value.", vbOKOnly
   txtMWTol.SetFocus
End If
End Sub

Private Sub txtNETTol_LostFocus()
If IsNumeric(txtNETTol.Text) Then
   samtDef.NETTol = CDbl(txtNETTol.Text)
Else
   If Len(Trim(txtNETTol.Text)) > 0 Then
      MsgBox "NET Tolerance should be number between 0 and 1.", vbOKOnly
      txtNETTol.SetFocus
   Else
      samtDef.NETTol = -1   'do not consider NET when searching
   End If
End If
End Sub

