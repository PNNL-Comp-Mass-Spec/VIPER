Attribute VB_Name = "Module17"
'last modified 08/14/2000 nt
'ORF file vizualisation procedures
'ORF gel is loaded from FTICR_AMT database
'ORF gel can also be saved as gel file
'all data from ORF table are loaded as Isotopic data
'Database property of GelData structure in case of ORF file
'is actual FTICR_AMT file and is always reconnected when ORF(saved) gel
'is loaded
'Certificate identifies ORF files
'ORF gels has to be in GelData structure so we can apply
'same procedures; still I would like to have separate class
'for visualization
'This is now extended on AMT and Source tables; there is line in
'comment field that distinguinish between the three
'----------------------------------------------------------------------
Option Explicit

Public Const glORF_TBL_NAME = "ORF"     'this is linked name; original name
                                        'is Molecule Database

Public Const ORFMark = "ORF:"

Public Const glORF_FLD_ID = "Molecule"
Public Const glORF_FLD_DESCRIPTION = "Notes"
Public Const glORF_FLD_MWMONO = "Monoisotopic mass"
Public Const glORF_FLD_MWAVG = "Average mass"
Public Const glORF_FLD_MWTMA = "Most abundant isotope"
Public Const glORF_FLD_PI = "PI"

Public Const glSRC_FLD_SCAN = "F_AFN"
Public Const glSRC_FLD_INT = "F_AInt"
Public Const glSRC_FLD_MWMONO = "F_AMW"
Public Const glSRC_FLD_ER = "F_AER"

Public Const glORF_UNI_INT = 10000000#  'uniform intensity

Public Function LoadNewDBGel(ByVal fname As String, ByVal Ind As Long) As Integer
Select Case GelStatus(Ind).DBGel
Case glDBGEL_ORF
     LoadNewDBGel = LoadNewORF(Ind)
     With GelData(Ind)
        .Comment = .Comment & vbCrLf & glCOMMENT_DBGEL & glCOMMENT_DBGEL_ORF _
        & glCOMMENT_DO_NOT_EDIT & vbCrLf
        .pICooSysEnabled = True
     End With
Case glDBGEL_AMT
     LoadNewDBGel = LoadNewAMT(Ind)
     With GelData(Ind)
        .Comment = .Comment & vbCrLf & glCOMMENT_DBGEL & glCOMMENT_DBGEL_AMT _
        & glCOMMENT_DO_NOT_EDIT & vbCrLf
        .pICooSysEnabled = True
     End With
Case Else
     LoadNewDBGel = LoadNewSource(Ind)
     With GelData(Ind)
        .Comment = .Comment & vbCrLf & glCOMMENT_DBGEL & GelStatus(Ind).DBGel _
        & glCOMMENT_DO_NOT_EDIT & vbCrLf
        .pICooSysEnabled = False
     End With
End Select
End Function

Private Function LoadNewORF(ByVal Ind As Long) As Integer
'Returns 0 if data successfuly loaded, -2 if data set is too large,
'- 3 if problems with file numbers, -4 if no data found,
'1 for any other error
Dim rsORF As Recordset
Dim MinMW As Double, MaxMW As Double
Dim ORFCnt As Long
Dim i As Long, k As Long
On Error GoTo err_LoadNewORF

LoadNewORF = 1
'open database, open ORF table
'OPENING DATABASE OS NOW SOMWHERE ELSE
Set rsORF = GelDB(Ind).OpenRecordset(glORF_TBL_NAME, dbOpenDynaset)
'populate recordset
rsORF.MoveLast
ORFCnt = rsORF.RecordCount
If ORFCnt > 0 Then
  With GelData(Ind)
     ReDim Preserve .ScanInfo(351)
     For i = 1 To 351
       With .ScanInfo(i)
          .ScanNumber = i
          .ScanFileName = i
          .ScanPI = (i - 1) / 25
       End With
     Next i
     .DataLines = ORFCnt
     .CSLines = 0
     .IsoLines = ORFCnt
     ReDim .IsoData(ORFCnt)
     'set index to match sort in ordinary gel files
     i = 0
     MinMW = glHugeOverExp
     MaxMW = 0
     .MinAbu = glORF_UNI_INT
     .MaxAbu = glORF_UNI_INT
     rsORF.MoveFirst
     Do Until (rsORF.EOF Or i > .IsoLines)
        i = i + 1
        .IsoData(i).ScanNumber = CInt(rsORF.Fields(glORF_FLD_PI).value * 25)
        'actual calculated pI is in column 2
        .IsoData(i).Charge = rsORF.Fields(glORF_FLD_PI).value
        .IsoData(i).Abundance = glORF_UNI_INT
        .IsoData(i).MZ = 0
        .IsoData(i).Fit = 0
        .IsoData(i).AverageMW = rsORF.Fields(glORF_FLD_MWAVG).value
        .IsoData(i).MonoisotopicMW = rsORF.Fields(glORF_FLD_MWMONO).value
        .IsoData(i).MostAbundantMW = rsORF.Fields(glORF_FLD_MWTMA).value
        
''        .IsoVar(i, isvfIsotopeLabel) = ""
''        .IsoVar(i, isvfMTDDRatio) = ""
        .IsoData(i).MTID = ORFMark & rsORF.Fields(glORF_FLD_ID).value & glARG_SEP _
                        & Chr$(32) & rsORF.Fields(glORF_FLD_DESCRIPTION).value
        FindMWExtremes .IsoData(i), MinMW, MaxMW, 0
        rsORF.MoveNext
     Loop
     .MinMW = MinMW
     .MaxMW = MaxMW
  End With
  LoadNewORF = 0
  Set rsORF = Nothing
Else
  LoadNewORF = -4
End If
Exit Function

err_LoadNewORF:
If Err.Number > 0 Then LogErrors Err.Number, "LoadNewORF"
End Function

Private Function LoadNewAMT(ByVal Ind As Long) As Integer

End Function

Private Function LoadNewSource(ByVal Ind As Long) As Integer
'Returns 0 if data successfuly loaded, -2 if data set is too large,
'- 3 if problems with file numbers, -4 if no data found,
'1 for any other error
Dim SourceID As Long
Dim SrcSQL As String
Dim rsSrc As Recordset
Dim MinFN As Integer, MaxFN As Integer
Dim MinMW As Double, MaxMW As Double
Dim MinAbu As Double, MaxAbu As Double
Dim SrcCnt As Long
Dim i As Long
On Error GoTo err_LoadNewSource

LoadNewSource = 1

SourceID = GelStatus(Ind).DBGel
If SourceID > 0 Then
   'retrieve Min and Max scan numbers from the FTSource table
   F_AGetSourceScan GelDB(Ind), SourceID, MinFN, MaxFN
   If MinFN <= MaxFN Then
      SrcSQL = "SELECT * FROM FTICR_AMT WHERE FTICR_AMT.F_AFTSID= " & SourceID & ";"
      Set rsSrc = GelDB(Ind).OpenRecordset(SrcSQL, dbOpenDynaset)
      'populate recordset
      rsSrc.MoveLast
      SrcCnt = rsSrc.RecordCount
      If SrcCnt > 0 Then
         With GelData(Ind)
            ReDim Preserve .ScanInfo(MaxFN - MinFN + 1)
            For i = MinFN To MaxFN
             With .ScanInfo(i - MinFN + 1)
                .ScanNumber = i
                .ScanFileName = i
             End With
            Next i
            .DataLines = SrcCnt
            .CSLines = 0
            .IsoLines = SrcCnt
            ReDim .IsoData(SrcCnt)
            'set index to match sort in ordinary gel files
            i = 0
            MinMW = glHugeOverExp
            MaxMW = 0
            MinAbu = glHugeOverExp
            MaxAbu = 0
            rsSrc.MoveFirst
            Do Until (rsSrc.EOF Or i > .IsoLines)
               i = i + 1
               .IsoData(i).ScanNumber = rsSrc.Fields(glSRC_FLD_SCAN).value
               .IsoData(i).Abundance = rsSrc.Fields(glSRC_FLD_INT).value
               'load Monoisotopic MW as all three MW fields
               .IsoData(i).MonoisotopicMW = rsSrc.Fields(glSRC_FLD_MWMONO).value
               .IsoData(i).AverageMW = .IsoData(i).MonoisotopicMW
               .IsoData(i).MostAbundantMW = .IsoData(i).MonoisotopicMW
               
               .IsoData(i).ExpressionRatio = rsSrc.Fields(glSRC_FLD_ER).value
               .IsoData(i).MTID = ""
               If .IsoData(i).MonoisotopicMW > MaxMW Then MaxMW = .IsoData(i).MonoisotopicMW
               If .IsoData(i).MonoisotopicMW < MinMW Then MinMW = .IsoData(i).MonoisotopicMW
               If .IsoData(i).Abundance > MaxAbu Then MaxAbu = .IsoData(i).Abundance
               If .IsoData(i).Abundance < MinAbu Then MinAbu = .IsoData(i).Abundance
               rsSrc.MoveNext
            Loop
            .MinMW = MinMW
            .MaxMW = MaxMW
            .MinAbu = MinAbu
            .MaxAbu = MaxAbu
          End With
          LoadNewSource = 0
          Set rsSrc = Nothing
       Else
          LoadNewSource = -4
       End If
   Else
     LoadNewSource = -3
   End If
End If
Exit Function

err_LoadNewSource:
If Err.Number > 0 Then LogErrors Err.Number, "LoadNewSource"
End Function

Public Function LoadSources(ByVal Ind As Long, _
                            ByRef Src() As String, _
                            ByRef SrcID() As Long) As Integer
'Loads sources from the FTSources table of GelDB(Ind)
'database and returns success code
Dim rsSrc As Recordset
Dim SrcCnt As Long
Dim i As Long
On Error GoTo err_LoadSources

Set rsSrc = GelDB(Ind).OpenRecordset(glFTICR_AMT_TBL_FTSOURCES, dbOpenTable)

SrcCnt = rsSrc.RecordCount
If SrcCnt > 0 Then
   ReDim Src(0 To SrcCnt - 1)
   ReDim SrcID(0 To SrcCnt - 1)
   With rsSrc
      .MoveFirst
      i = 0
      Do Until (.EOF Or i >= SrcCnt)
         Src(i) = GetFileNameOnly(.Fields("FTSFileName").value)
         SrcID(i) = .Fields("FTSFileID").value
         i = i + 1
         .MoveNext
      Loop
   End With
   LoadSources = 0
Else
   LoadSources = 2
End If

exit_LoadSources:
On Error Resume Next
rsSrc.Close
Exit Function

err_LoadSources:
If Err.Number = 3011 Then   'object not found
   LoadSources = 1
Else
   MsgBox "Error: " & Err.Number & " - " & Err.Description, vbOKOnly
   LoadSources = -1
End If
Err.Clear
GoTo exit_LoadSources
End Function
