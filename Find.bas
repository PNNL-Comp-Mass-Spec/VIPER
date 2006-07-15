Attribute VB_Name = "Module11"
'Procedures and declaration supporting Find function
'last modified 08/22/2000 nt
Option Explicit

Public Const glFIND_NOTHING = 0
Public Const glFIND_VALUE = 1
Public Const glFIND_LIST = 2
Public Const glFIND_RANGE = 3
Public Const glFIND_SELECTION = 4

Public Const glFIELD_MW = 0
Public Const glFIELD_ER = 1
Public Const glFIELD_MOVERZ = 2
Public Const glFIELD_CS = 3
Public Const glFIELD_ABU = 4
Public Const glFIELD_ID = 5
Public Const glFIELD_FIT = 6


Public Function FindNumeric(ByVal Ind As Long, _
                            ByRef v() As Double, _
                            ByVal SearchField As Integer, _
                            ByVal Tolerance As Double, _
                            ByVal ToleranceType As Integer, _
                            ByVal FindArgType As Integer, _
                            ByVal bZoom As Boolean, _
                            ByVal bSelect As Boolean) As Integer
'returns number of values found
'Ind - gel index; V - array with values(range) to find,
'Search field - determines what a we looking for
'Tolerance - tolerance value (ignored if range is looked),
'ToleranceType - ppm, percentage, absolute
'FindArgType - value, list of values, range
'bZoom - if True zoom in found, bSelect - if True select findings
Dim CurrMMA As Double
Dim dblValue As Double
Dim FindCnt As Long
Dim FoundCnt As Long
Dim MinMW As Double
Dim MaxMW As Double
Dim MinFN As Double
Dim MaxFN As Double
Dim IsoMWF As Integer     'Isotopic molecular mass field
Dim i As Long, j As Long

If FindArgType = glFIND_NOTHING Then Exit Function
FindCnt = UBound(v)
'track where findings are located even if there will be no zoom
MinMW = glHugeOverExp
MaxMW = 0
With GelData(Ind)
    If UBound(.ScanInfo) > 0 Then
       MinFN = .ScanInfo(UBound(.ScanInfo)).ScanNumber
       MaxFN = .ScanInfo(1).ScanNumber
    Else
       FindNumeric = -1
       Exit Function
    End If
    IsoMWF = .Preferences.IsoDataField
End With

'if Findings are to be selected clear old selection
If bSelect Then GelBody(Ind).GelSel.Clear
FoundCnt = 0
Select Case FindArgType
Case glFIND_VALUE, glFIND_LIST, glFIND_SELECTION
    With GelData(Ind)
      For i = 1 To FindCnt
        Select Case ToleranceType
        Case gltPct
           CurrMMA = v(i) * Tolerance * glPCT
        Case gltPPM
           CurrMMA = v(i) * Tolerance * glPPM
        Case Else
           CurrMMA = Tolerance
        End Select
        If .CSLines > 0 Then
           For j = 1 To .CSLines
             'count only visible spots
             If GelDraw(Ind).CSID(j) > 0 And GelDraw(Ind).CSR(j) > 0 Then
               dblValue = GetIsoOrCSDataByField(Ind, .CSData(j), SearchField, True)
               If Abs(dblValue - v(i)) <= CurrMMA Then
                  FoundCnt = FoundCnt + 1
                  If .CSData(j).AverageMW < MinMW Then MinMW = .CSData(j).AverageMW
                  If .CSData(j).AverageMW > MaxMW Then MaxMW = .CSData(j).AverageMW
                  If .CSData(j).ScanNumber < MinFN Then MinFN = .CSData(j).ScanNumber
                  If .CSData(j).ScanNumber > MaxFN Then MaxFN = .CSData(j).ScanNumber
                  If bSelect Then GelBody(Ind).GelSel.AddToCSSelection j
               End If
             End If
           Next j
        End If
        If .IsoLines > 0 Then
           For j = 1 To .IsoLines
             If GelDraw(Ind).IsoID(j) > 0 And GelDraw(Ind).IsoR(j) > 0 Then
               dblValue = GetIsoOrCSDataByField(Ind, .IsoData(j), SearchField, False)
               If Abs(dblValue - v(i)) <= CurrMMA Then
                  FoundCnt = FoundCnt + 1
                  If GetIsoMass(.IsoData(j), IsoMWF) < MinMW Then MinMW = GetIsoMass(.IsoData(j), IsoMWF)
                  If GetIsoMass(.IsoData(j), IsoMWF) > MaxMW Then MaxMW = GetIsoMass(.IsoData(j), IsoMWF)
                  If .IsoData(j).ScanNumber < MinFN Then MinFN = .IsoData(j).ScanNumber
                  If .IsoData(j).ScanNumber > MaxFN Then MaxFN = .IsoData(j).ScanNumber
                  If bSelect Then GelBody(Ind).GelSel.AddToIsoSelection j
               End If
             End If
           Next j
        End If
      Next i
      'move a bit from the edges
      If MinFN > .ScanInfo(1).ScanNumber Then MinFN = MinFN - 1
      If MaxFN < .ScanInfo(UBound(.ScanInfo)).ScanNumber Then MaxFN = MaxFN + 1
    End With
Case glFIND_RANGE
    With GelData(Ind)
        If .CSLines > 0 Then
           For j = 1 To .CSLines
             'count only visible spots
             If GelDraw(Ind).CSID(j) > 0 And GelDraw(Ind).CSR(j) > 0 Then
               dblValue = GetIsoOrCSDataByField(Ind, .CSData(j), SearchField, True)
               If dblValue >= v(1) And dblValue <= v(2) Then
                  FoundCnt = FoundCnt + 1
                  If .CSData(j).AverageMW < MinMW Then MinMW = .CSData(j).AverageMW
                  If .CSData(j).AverageMW > MaxMW Then MaxMW = .CSData(j).AverageMW
                  If .CSData(j).ScanNumber < MinFN Then MinFN = .CSData(j).ScanNumber
                  If .CSData(j).ScanNumber > MaxFN Then MaxFN = .CSData(j).ScanNumber
                  If bSelect Then GelBody(Ind).GelSel.AddToCSSelection j
               End If
             End If
           Next j
        End If
        If .IsoLines > 0 Then
           For j = 1 To .IsoLines
             If GelDraw(Ind).IsoID(j) > 0 And GelDraw(Ind).IsoR(j) > 0 Then
               dblValue = GetIsoOrCSDataByField(Ind, .IsoData(j), SearchField, False)
               If dblValue >= v(1) And dblValue <= v(2) Then
                  FoundCnt = FoundCnt + 1
                  If GetIsoMass(.IsoData(j), IsoMWF) < MinMW Then MinMW = GetIsoMass(.IsoData(j), IsoMWF)
                  If GetIsoMass(.IsoData(j), IsoMWF) > MaxMW Then MaxMW = GetIsoMass(.IsoData(j), IsoMWF)
                  If .IsoData(j).ScanNumber < MinFN Then MinFN = .IsoData(j).ScanNumber
                  If .IsoData(j).ScanNumber > MaxFN Then MaxFN = .IsoData(j).ScanNumber
                  If bSelect Then GelBody(Ind).GelSel.AddToIsoSelection j
               End If
             End If
           Next j
        End If
        'move a bit from the edges
        If MinFN > .ScanInfo(1).ScanNumber Then MinFN = MinFN - 1
        If MaxFN < .ScanInfo(UBound(.ScanInfo)).ScanNumber Then MaxFN = MaxFN + 1
    End With
End Select
FindNumeric = FoundCnt
If FoundCnt > 0 Then
   If bZoom Then
      GelBody(Ind).csMyCooSys.ZoomInR MinFN, MinMW, MaxFN, MaxMW
   Else
      GelBody(Ind).picGraph.Refresh
   End If
End If
End Function


Public Function FindMatchingIDs(ByVal Ind As Long, _
                         ByRef v() As String, _
                         ByVal SearchField As Integer, _
                         ByVal FindArgType As Integer, _
                         ByVal bZoom As Boolean, _
                         ByVal bSelect As Boolean) As Integer
'returns number of values found
'Ind - gel index; V - array with strings to find,
'Search field - determines what are we looking for
'bZoom - if True zoom in found, bSelect - if True select findings
'CURRENTLY SEARCH IS DONE AS IN CASE OF 'OR' LOGICAL OPERATOR
'IS BETWEEN ELEMENTS OF ARRAY V (IT IS ENOUGH TO FIND ONE OF STRINGS)
Dim FindCnt As Long
Dim FoundCnt As Long
Dim MinMW As Double
Dim MaxMW As Double
Dim MinFN As Double
Dim MaxFN As Double
Dim IsoMWF As Integer     'Isotopic molecular mass field
Dim CSFld As Integer      'Charge state field to search
Dim IsoFld As Integer      'Isotopic field to search
Dim i As Long, j As Long

' This function only handles searching for PMT matches in .IsoData().MTID or .CSData().MTID
Debug.Assert SearchField <> glFIELD_ID

If FindArgType = glFIND_NOTHING Then Exit Function
FindCnt = UBound(v)
'track where findings are located even if there will be no zoom
MinMW = glHugeOverExp
MaxMW = 0
With GelData(Ind)
    If UBound(.ScanInfo) > 0 Then
       MinFN = .ScanInfo(UBound(.ScanInfo)).ScanNumber
       MaxFN = .ScanInfo(1).ScanNumber
    Else
       FindMatchingIDs = -1
       Exit Function
    End If
    IsoMWF = .Preferences.IsoDataField
End With

'if Findings are to be selected clear old selection
If bSelect Then GelBody(Ind).GelSel.Clear
FoundCnt = 0
Select Case FindArgType
Case glFIND_VALUE, glFIND_LIST, glFIND_SELECTION
    With GelData(Ind)
      For i = 1 To FindCnt
        If .CSLines > 0 And CSFld >= 0 Then
           For j = 1 To .CSLines
             'count only visible spots
             If GelDraw(Ind).CSID(j) > 0 And _
                GelDraw(Ind).CSR(j) > 0 And _
                Len(.CSData(j).MTID) > 0 Then
                   If InStr(1, .CSData(j).MTID, v(i)) > 0 Then
                      FoundCnt = FoundCnt + 1
                      If .CSData(j).AverageMW < MinMW Then MinMW = .CSData(j).AverageMW
                      If .CSData(j).AverageMW > MaxMW Then MaxMW = .CSData(j).AverageMW
                      If .CSData(j).ScanNumber < MinFN Then MinFN = .CSData(j).ScanNumber
                      If .CSData(j).ScanNumber > MaxFN Then MaxFN = .CSData(j).ScanNumber
                      If bSelect Then GelBody(Ind).GelSel.AddToCSSelection j
                   End If
             End If
           Next j
        End If
        If .IsoLines > 0 And IsoFld >= 0 Then
           For j = 1 To .IsoLines
             If GelDraw(Ind).IsoID(j) > 0 And _
                GelDraw(Ind).IsoR(j) > 0 And _
                Len(.IsoData(j).MTID) > 0 Then
                   If InStr(1, .IsoData(j).MTID, v(i)) > 0 Then
                      FoundCnt = FoundCnt + 1
                      If GetIsoMass(.IsoData(j), IsoMWF) < MinMW Then MinMW = GetIsoMass(.IsoData(j), IsoMWF)
                      If GetIsoMass(.IsoData(j), IsoMWF) > MaxMW Then MaxMW = GetIsoMass(.IsoData(j), IsoMWF)
                      If .IsoData(j).ScanNumber < MinFN Then MinFN = .IsoData(j).ScanNumber
                      If .IsoData(j).ScanNumber > MaxFN Then MaxFN = .IsoData(j).ScanNumber
                      If bSelect Then GelBody(Ind).GelSel.AddToIsoSelection j
                   End If
             End If
           Next j
        End If
      Next i
      'move a bit from the edges
      If MinFN > .ScanInfo(1).ScanNumber Then MinFN = MinFN - 1
      If MaxFN < .ScanInfo(UBound(.ScanInfo)).ScanNumber Then MaxFN = MaxFN + 1
    End With
Case glFIND_RANGE
    MsgBox "Range of values not implemented for text searches.", vbOKOnly
End Select
FindMatchingIDs = FoundCnt
If FoundCnt > 0 Then
   If bZoom Then
      GelBody(Ind).csMyCooSys.ZoomInR MinFN, MinMW, MaxFN, MaxMW
   Else
      GelBody(Ind).picGraph.Refresh
   End If
End If
End Function


Public Function FindER(ByVal Ind As Long, _
                       ByRef v() As Double, _
                       ByVal Tolerance As Double, _
                       ByVal ToleranceType As Integer, _
                       ByVal FindArgType As Integer, _
                       ByVal bZoom As Boolean, _
                       ByVal bSelect As Boolean) As Integer
'returns number of values found
'Ind - gel index; V - array with values(range) to find,
'Tolerance - tolerance value (ignored if range is looked),
'ToleranceType - ppm, percentage, absolute
'FindArgType - value, list of values, range
'bZoom - if True zoom in found, bSelect - if True select findings
Dim CurrMMA As Double
Dim FindCnt As Long
Dim FoundCnt As Long
Dim MinMW As Double
Dim MaxMW As Double
Dim MinFN As Double
Dim MaxFN As Double
Dim IsoMWF As Integer     'Isotopic molecular mass field
Dim i As Long, j As Long

If FindArgType = glFIND_NOTHING Then Exit Function
FindCnt = UBound(v)
'track where findings are located even if there will be no zoom
MinMW = glHugeOverExp
MaxMW = 0
With GelData(Ind)
    If UBound(.ScanInfo) > 0 Then
       MinFN = .ScanInfo(UBound(.ScanInfo)).ScanNumber
       MaxFN = .ScanInfo(1).ScanNumber
    Else
       FindER = -1
       Exit Function
    End If
    IsoMWF = .Preferences.IsoDataField
End With
'if Findings are to be selected clear old selection
If bSelect Then GelBody(Ind).GelSel.Clear
FoundCnt = 0
Select Case FindArgType
Case glFIND_VALUE, glFIND_LIST, glFIND_SELECTION
    With GelDraw(Ind)
      For i = 1 To FindCnt
        Select Case ToleranceType
        Case gltPct
           CurrMMA = v(i) * Tolerance * glPCT
        Case gltPPM
           CurrMMA = v(i) * Tolerance * glPPM
        Case Else
           CurrMMA = Tolerance
        End Select
        If .CSCount > 0 Then
           For j = 1 To .CSCount
             'count only visible spots
             If .CSID(j) > 0 And .CSR(j) > 0 Then
                If Abs(.CSER(j) - v(i)) <= CurrMMA Then
                   FoundCnt = FoundCnt + 1
                   With GelData(Ind)
                     If .CSData(j).AverageMW < MinMW Then MinMW = .CSData(j).AverageMW
                     If .CSData(j).AverageMW > MaxMW Then MaxMW = .CSData(j).AverageMW
                     If .CSData(j).ScanNumber < MinFN Then MinFN = .CSData(j).ScanNumber
                     If .CSData(j).ScanNumber > MaxFN Then MaxFN = .CSData(j).ScanNumber
                   End With
                   If bSelect Then GelBody(Ind).GelSel.AddToCSSelection j
                End If
             End If
           Next j
        End If
        If .IsoCount > 0 Then
           For j = 1 To .IsoCount
             If .IsoID(j) > 0 And .IsoR(j) > 0 Then
                If Abs(.IsoER(j) - v(i)) <= CurrMMA Then
                   FoundCnt = FoundCnt + 1
                   With GelData(Ind)
                     If GetIsoMass(.IsoData(j), IsoMWF) < MinMW Then MinMW = GetIsoMass(.IsoData(j), IsoMWF)
                     If GetIsoMass(.IsoData(j), IsoMWF) > MaxMW Then MaxMW = GetIsoMass(.IsoData(j), IsoMWF)
                     If .IsoData(j).ScanNumber < MinFN Then MinFN = .IsoData(j).ScanNumber
                     If .IsoData(j).ScanNumber > MaxFN Then MaxFN = .IsoData(j).ScanNumber
                   End With
                   If bSelect Then GelBody(Ind).GelSel.AddToIsoSelection j
                End If
             End If
           Next j
        End If
      Next i
      'move a bit from the edges
      With GelData(Ind)
        If MinFN > .ScanInfo(1).ScanNumber Then MinFN = MinFN - 1
        If MaxFN < .ScanInfo(UBound(.ScanInfo)).ScanNumber Then MaxFN = MaxFN + 1
      End With
    End With
Case glFIND_RANGE
    With GelDraw(Ind)
        If .CSCount > 0 Then
           For j = 1 To .CSCount
             'count only visible spots
             If .CSID(j) > 0 And .CSR(j) > 0 Then
                If .CSER(j) >= v(1) And .CSER(j) <= v(2) Then
                   FoundCnt = FoundCnt + 1
                   With GelData(Ind)
                     If .CSData(j).AverageMW < MinMW Then MinMW = .CSData(j).AverageMW
                     If .CSData(j).AverageMW > MaxMW Then MaxMW = .CSData(j).AverageMW
                     If .CSData(j).ScanNumber < MinFN Then MinFN = .CSData(j).ScanNumber
                     If .CSData(j).ScanNumber > MaxFN Then MaxFN = .CSData(j).ScanNumber
                   End With
                   If bSelect Then GelBody(Ind).GelSel.AddToCSSelection j
                End If
             End If
           Next j
        End If
        If .IsoCount > 0 Then
           For j = 1 To .IsoCount
             If .IsoID(j) > 0 And .IsoR(j) > 0 Then
                If .IsoER(j) >= v(1) And .IsoER(j) <= v(2) Then
                   FoundCnt = FoundCnt + 1
                   With GelData(Ind)
                     If GetIsoMass(.IsoData(j), IsoMWF) < MinMW Then MinMW = GetIsoMass(.IsoData(j), IsoMWF)
                     If GetIsoMass(.IsoData(j), IsoMWF) > MaxMW Then MaxMW = GetIsoMass(.IsoData(j), IsoMWF)
                     If .IsoData(j).ScanNumber < MinFN Then MinFN = .IsoData(j).ScanNumber
                     If .IsoData(j).ScanNumber > MaxFN Then MaxFN = .IsoData(j).ScanNumber
                   End With
                   If bSelect Then GelBody(Ind).GelSel.AddToIsoSelection j
                End If
             End If
           Next j
        End If
        'move a bit from the edges
        With GelData(Ind)
          If MinFN > .ScanInfo(1).ScanNumber Then MinFN = MinFN - 1
          If MaxFN < .ScanInfo(UBound(.ScanInfo)).ScanNumber Then MaxFN = MaxFN + 1
        End With
    End With
End Select
FindER = FoundCnt
If FoundCnt > 0 Then
   If bZoom Then
      GelBody(Ind).csMyCooSys.ZoomInR MinFN, MinMW, MaxFN, MaxMW
   Else
      GelBody(Ind).picGraph.Refresh
   End If
End If
End Function

