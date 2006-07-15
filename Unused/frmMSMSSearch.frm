VERSION 5.00
Begin VB.Form frmMSMSSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ICR-2LS Services - MS/MS Search"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   Icon            =   "frmMSMSSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdStop 
      Caption         =   "Sto&p"
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      ToolTipText     =   "Stops current search (changes already made stay)"
      Top             =   1440
      Width           =   975
   End
   Begin VB.Frame fraResults 
      Caption         =   "Store Results In"
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   2895
      Begin VB.OptionButton optResults 
         Caption         =   "&Text File"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   2295
      End
      Begin VB.OptionButton optResults 
         Caption         =   "FTICR_AMT &Database"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.Frame fraFragments 
      Caption         =   "Fragments"
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2895
      Begin VB.OptionButton optFragments 
         Caption         =   "&Use Data Currently in Scope"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   2415
      End
      Begin VB.OptionButton optFragments 
         Caption         =   "Use &All Data"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.CheckBox chkUseAMTFoundOnly 
      Caption         =   "&Use only data points matched with some of AMT records"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Value           =   1  'Checked
      Width           =   4335
   End
   Begin VB.Label lblProgress 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Parents are always selected from data currently in scope"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmMSMSSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'last modified 06/29/2000 nt
Option Explicit

Dim CallerID As Long     'index of current gel in GelBody array
Dim CallerFTSID As Long     'index of current gel in the FTICR_AMT db
'parent and fragment arrays are 0 based
'Significant speed improvement could be achieved if parents
'are ordered by scan numbers so that we have to load fragments
'only once for each scan number
Dim PID() As Long
Dim PIDType() As Integer
Dim PMW() As Double         'parent masses
Dim PFN() As Integer        'this is not neccessary but we can save time
                            'on loading fragments if parent arrays are
                            'ordered on scan numbers
Dim PResults() As String    'results returned
Dim PCnt As Long            'parent masses count
Dim FMW() As Double         'fragments masses
Dim FIN() As Double         'fragments intensities

Dim dbFTICR_AMT As Database
Dim rsFTICR_AMT As Recordset

Dim hfile As Integer
Dim TextFile As String
Dim StopSearch As Boolean

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSearch_Click()
If optResults(0).value Then     'save to database
   SearchResToDB
Else                            'save it to text file
   SearchResToText
End If
End Sub

Private Sub cmdStop_Click()
StopSearch = True
End Sub

Private Sub Form_Activate()
Dim Resp
CallerID = Me.Tag
If InStr(1, GelBody(CallerID).Caption, "Untitled") > 0 Then
   Resp = MsgBox("Requested function associates data from current gel file with AMT records." _
        & " Keeping track is assured by gel files' names which should not be changed once assigned." _
        & " It appears that current gel file still has default name. It would be good idea to" _
        & " cancel current procedure and assign gel more meaningful name. Cancel?", vbYesNo)
   If Resp = vbYes Then
      Unload Me
      Exit Sub
   End If
End If
PCnt = 0
End Sub

Private Sub LoadParents()
Dim i As Long
Dim IsoF As Integer     'just shortcut for MW field for Isotopic data

IsoF = GelData(CallerID).Preferences.IsoDataField
PCnt = 0
With GelDraw(CallerID)
  If .CSCount > 0 Or .IsoCount > 0 Then
     ReDim PID(.CSCount + .IsoCount - 1)
     ReDim PIDType(.CSCount + .IsoCount - 1)
     ReDim PMW(.CSCount + .IsoCount - 1)
     ReDim PFN(.CSCount + .IsoCount - 1)
  Else
     Exit Sub
  End If
  If chkUseAMTFoundOnly.value = vbChecked Then
    If .CSCount > 0 Then
       For i = 1 To .CSCount
           If .CSID(i) > 0 And .CSR(i) > 0 Then
              If IsAMTReferenced(GelData(CallerID).CSData(i).MTID) Then
                 PCnt = PCnt + 1
                 PID(PCnt - 1) = i
                 PIDType(PCnt - 1) = glCSType
                 PMW(PCnt - 1) = GelData(CallerID).CSData(i).AverageMW
                 PFN(PCnt - 1) = GelData(CallerID).CSData(i).ScanNumber
              End If
           End If
       Next i
    End If
    If .IsoCount > 0 Then
       For i = 1 To .IsoCount
           If .IsoID(i) > 0 And .IsoR(i) > 0 Then
              If IsAMTReferenced(GelData(CallerID).IsoData(i).MTID) Then
                 PCnt = PCnt + 1
                 PID(PCnt - 1) = i
                 PIDType(PCnt - 1) = glIsoType
                 PMW(PCnt - 1) = GetIsoMass(GelData(CallerID).IsoData(i), IsoF)
                 PFN(PCnt - 1) = GelData(CallerID).IsoData(i).ScanNumber
              End If
           End If
       Next i
    End If
  Else
    If .CSCount > 0 Then
       For i = 1 To .CSCount
           If .CSID(i) > 0 And .CSR(i) > 0 Then
              PCnt = PCnt + 1
              PID(PCnt - 1) = i
              PIDType(PCnt - 1) = glCSType
              PMW(PCnt - 1) = GelData(CallerID).CSData(i).AverageMW
              PFN(PCnt - 1) = GelData(CallerID).CSData(i).ScanNumber
           End If
       Next i
    End If
    If .IsoCount > 0 Then
       For i = 1 To .IsoCount
           If .IsoID(i) > 0 And .IsoR(i) > 0 Then
              PCnt = PCnt + 1
              PID(PCnt - 1) = i
              PIDType(PCnt - 1) = glIsoType
              PMW(PCnt - 1) = GetIsoMass(GelData(CallerID).IsoData(i), IsoF)
              PFN(PCnt - 1) = GelData(CallerID).IsoData(i).ScanNumber
           End If
       Next i
    End If
  End If
End With
If PCnt > 0 Then
   ReDim Preserve PID(PCnt - 1)
   ReDim Preserve PIDType(PCnt - 1)
   ReDim Preserve PMW(PCnt - 1)
   ReDim Preserve PFN(PCnt - 1)
   SortParentsOnFN 0, PCnt - 1
Else
   Erase PID
   Erase PIDType
   Erase PMW
   Erase PFN
End If
End Sub

Private Function LoadFragments(ByVal FN As Integer) As Long
'returns number of fragments
Dim i As Long
Dim FrgCnt As Long
Dim IsoF As Integer
On Error GoTo exit_LoadFragments

FrgCnt = 0
With GelData(CallerID)
   If .CSLines > 0 Or .IsoLines > 0 Then
      ReDim FMW(.CSLines + .IsoLines - 1)
      ReDim FIN(.CSLines + .IsoLines - 1)
   Else
      GoTo exit_LoadFragments
   End If
   IsoF = .Preferences.IsoDataField
   If optFragments(0).value Then       'use all data
      If .CSLines > 0 Then
         For i = 1 To .CSLines
             If .CSData(i).ScanNumber = FN Then
                FrgCnt = FrgCnt + 1
                FMW(FrgCnt - 1) = .CSData(i).AverageMW
                FIN(FrgCnt - 1) = .CSData(i).Abundance
             End If
         Next i
      End If
      If .IsoLines > 0 Then
         For i = 1 To .IsoLines
             If .IsoData(i).ScanNumber = FN Then
                FrgCnt = FrgCnt + 1
                FMW(FrgCnt - 1) = GetIsoMass(.IsoData(i), IsoF)
                FIN(FrgCnt - 1) = .IsoData(i).Abundance
             End If
         Next i
      End If
   Else                              'use data currently in scope here it
      If .CSLines > 0 Then           'means to exclude only data filtered out;
         For i = 1 To .CSLines       'not data not visible because of zoom
             If ((.CSData(i).ScanNumber = FN) And (GelDraw(CallerID).CSID(i) > 0)) Then
                FrgCnt = FrgCnt + 1
                FMW(FrgCnt - 1) = .CSData(i).AverageMW
                FIN(FrgCnt - 1) = .CSData(i).Abundance
             End If
         Next i
      End If
      If .IsoLines > 0 Then
         For i = 1 To .IsoLines
             If ((.IsoData(i).ScanNumber = FN) And (GelDraw(CallerID).IsoID(i) > 0)) Then
                FrgCnt = FrgCnt + 1
                FMW(FrgCnt - 1) = GetIsoMass(.IsoData(i), IsoF)
                FIN(FrgCnt - 1) = .IsoData(i).Abundance
             End If
         Next i
      End If
   End If
End With

exit_LoadFragments:
If FrgCnt > 0 Then
   ReDim Preserve FMW(FrgCnt - 1)
   ReDim Preserve FIN(FrgCnt - 1)
Else
   Erase FMW
   Erase FIN
End If
LoadFragments = FrgCnt
End Function

Private Sub SearchResToText()
Dim i As Long
Dim FCnt As Long
Dim LastFN As Integer
Dim AParent() As Double 'ICR-2LS function expects array
Dim Ret As Integer
On Error GoTo exit_SearchResToText
TextFile = SaveFileAPIDlg(Me.hwnd, "Text file (*.txt)" & Chr(0) & "*.txt" _
            & Chr(0), 1, "MS_MSSearch.txt", "Save Results of MS/MS Search")
If Len(TextFile) <= 0 Then Exit Sub
hfile = FreeFile
Open TextFile For Output As #hfile
Me.MousePointer = vbHourglass
LoadParents
If PCnt > 0 Then
   ReDim PResults(PCnt - 1)
   ReDim AParent(0)
   'do search
   LastFN = -1
   For i = 0 To PCnt - 1
       DoEvents
       If StopSearch Then
          StopSearch = False
          GoTo exit_SearchResToText
       End If
       If PFN(i) <> LastFN Then
          FCnt = LoadFragments(PFN(i))
          LastFN = PFN(i)
       End If
       lblProgress.Caption = "Parent: " & (i + 1) & "/" & PCnt & " --- Fragments load: " & FCnt
       DoEvents
       If FCnt > 0 Then
          AParent(0) = PMW(i)
          Ret = objICR2LS.ParentMasses(AParent)
          Ret = objICR2LS.FragmentMasses(FMW, FIN)
          PResults(i) = objICR2LS.SearchDB
          WriteResToTextFile i
       End If
   Next i
   lblProgress.Caption = ""
Else
   MsgBox "No parent masses found.", vbOKOnly
End If

exit_SearchResToText:
On Error Resume Next
Me.MousePointer = vbDefault
Close #hfile
End Sub

Private Sub SearchResToDB()
Dim i As Long
Dim FCnt As Long
Dim LastFN As Integer
Dim AParent() As Double 'ICR-2LS function expects array
Dim Ret As Integer
Dim Resp
Dim SaveResults As Boolean

SaveResults = True
CallerFTSID = F_ACheckTheSource(dbFTICR_AMT, GelBody(CallerID).Caption)
If CallerFTSID < 0 Then             'new source
   CallerFTSID = F_AAddSource(dbFTICR_AMT, CallerID)
   If CallerFTSID < 0 Then       'error while adding new source
      Resp = MsgBox("Error trying to add a source to the FTICR_AMT database. " _
         & "Search results can not be saved in the database. " _
         & "Save results in text file? ", vbYesNoCancel, "MS/MS Search")
      Select Case Resp
      Case vbYes        'save it as text file
           SearchResToText
           Exit Sub
      Case vbNo         'do the search without saving the results
           SaveResults = False
      Case vbCancel     'cancel the search
           Exit Sub
      End Select
   End If
Else                             'source already exists
   Resp = MsgBox("Data from current gel already found in FTICR_AMT database." _
        & " Continue with MS/MS analysis(if you select Yes, old records" _
        & " related to this gel will be deleted and new one appended)?", vbYesNo)
   If Resp <> vbYes Then Exit Sub
   'delete all records in FTICR_AMT table related to CallerFTSID
   F_ADeleteSourceRecords dbFTICR_AMT, CallerFTSID
   'update record in FTSources table
   F_AUpdateSource dbFTICR_AMT, CallerFTSID, CallerID
End If
Me.MousePointer = vbHourglass
LoadParents
If PCnt > 0 Then
   ReDim PResults(PCnt - 1)
   ReDim AParent(0)
   'do search
   LastFN = -1
   For i = 0 To PCnt - 1
       DoEvents
       If StopSearch Then
          StopSearch = False
          GoTo exit_SearchResToDB
       End If
       If PFN(i) <> LastFN Then
          FCnt = LoadFragments(PFN(i))
          LastFN = PFN(i)
       End If
       lblProgress.Caption = "Parent: " & (i + 1) & "/" & PCnt & " --- Fragments load: " & FCnt
       DoEvents
       If FCnt > 0 Then
          AParent(0) = PMW(i)
          Ret = objICR2LS.ParentMasses(AParent)
          Ret = objICR2LS.FragmentMasses(FMW, FIN)
          PResults(i) = objICR2LS.SearchDB
          If SaveResults Then WriteResToDB i
       End If
   Next i
   lblProgress.Caption = ""
Else
   MsgBox "No parent masses found.", vbOKOnly
End If
exit_SearchResToDB:
Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
'try to connect to FTICR_AMT database
If Not ConnectToFTICR_AMT(dbFTICR_AMT, sFTICR_AMTPath, True) Then GoTo FTICR_AMTNotOK
Set rsFTICR_AMT = F_AOpenFTICR_AMTTbl(dbFTICR_AMT)
If Not rsFTICR_AMT Is Nothing Then Exit Sub

FTICR_AMTNotOK:
optResults(1).value = True
optResults(0).Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set rsFTICR_AMT = Nothing
Set dbFTICR_AMT = Nothing
End Sub

Private Function WriteResToDB(ByVal PInd As Long)
'PInd is index in Pxxx arrays
Dim aAMT() As String     'array of AMT IDs; usually with 1 element,
                         'but possibly more
Dim aAMTCnt As Long
Dim sIdentity As String
Dim iFN As Integer
Dim i As Long

With GelData(CallerID)
    sIdentity = .IsoData(PID(PInd)).MTID
    iFN = .IsoData(PID(PInd)).ScanNumber
End With

aAMTCnt = GetAMTRefFromString1(sIdentity, aAMT())
If InStr(1, PResults(PInd), "Score") > 0 And aAMTCnt > 0 Then
  For i = 1 To aAMTCnt
    With rsFTICR_AMT
      .AddNew
      .Fields("F_AFTSID").value = CallerFTSID         'caller gel source ID
      .Fields("F_AMTID").value = aAMT(i)             'AMT ID (from Identity)
      .Fields("F_AMW").value = PMW(PInd)              'gel monoisotopic MW
      .Fields("F_AFN").value = iFN                    'scan number
      .Fields("F_AIndex").value = PID(PInd)           'index in Isotopic file
      .Fields("F_AMS_MSData").value = PResults(PInd)  'results from MS_MS search
      .Update
    End With
  Next i
End If
End Function

Private Function WriteResToTextFile(ByVal PInd As Long)
'PInd is index in Pxxx arrays
Dim aAMT() As String     'array of AMT IDs; usually with 1 element,
                         'but possibly more
Dim sTmp As String
Dim aAMTCnt As Long
Dim sIdentity As String
Dim iFN As Integer
Dim i As Long

With GelData(CallerID)
    sIdentity = .IsoData(PID(PInd)).MTID
    iFN = .IsoData(PID(PInd)).ScanNumber
End With

aAMTCnt = GetAMTRefFromString1(sIdentity, aAMT())
If InStr(1, PResults(PInd), "Score") > 0 And aAMTCnt > 0 Then
  For i = 1 To aAMTCnt
    sTmp = sTmp & CallerFTSID & vbTab
    sTmp = sTmp & aAMT(i) & vbTab
    sTmp = sTmp & PMW(PInd) & vbTab
    sTmp = sTmp & iFN & vbTab
    sTmp = sTmp & PID(PInd) & vbCrLf
    sTmp = sTmp & PResults(PInd) & vbCrLf
    Print #hfile, sTmp
  Next i
End If
End Function

Private Sub SortParentsOnFN(ByVal nLow As Long, ByVal nHigh As Long)
Dim i As Long, j As Long
Dim x As Integer, y As Integer
Dim z As Long, w As Double

i = nLow
j = nHigh
x = PFN((nLow + nHigh) / 2)
Do While i <= j
   Do While (PFN(i) < x And i < nHigh)
      i = i + 1
   Loop
   Do While (x < PFN(j) And j > nLow)
      j = j - 1
   Loop
   If i <= j Then    'swap them; all arrays
      y = PFN(i)
      PFN(i) = PFN(j)
      PFN(j) = y
      
      y = PIDType(i)
      PIDType(i) = PIDType(j)
      PIDType(j) = y
      
      z = PID(i)
      PID(i) = PID(j)
      PID(j) = z
      
      w = PMW(i)
      PMW(i) = PMW(j)
      PMW(j) = w
                  
      i = i + 1
      j = j - 1
   End If
Loop
If nLow < j Then SortParentsOnFN nLow, j   'recursions
If i < nHigh Then SortParentsOnFN i, nHigh
End Sub
