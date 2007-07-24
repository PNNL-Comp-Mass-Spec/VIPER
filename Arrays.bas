Attribute VB_Name = "Module4"
'Module with array procedures
'Last Modified 04/15/2003 nt
'--------------------------------------------------
Option Explicit

Public Function GetBadFits(Ind As Long, Fit As Double) As Variant
'-------------------------------------------------------------------
'Ind Index of GelData, Fit is acceptable Fit limit; function returns
'variant array index of bad fits in Iso data arrays
'-------------------------------------------------------------------
Dim i As Long
Dim BFDCount As Long
Dim aTmp() As Variant
On Error GoTo err_GetBadFits
BFDCount = 0
With GelData(Ind)
    ReDim aTmp(.IsoLines)
    For i = 1 To .IsoLines
        If .IsoData(i).Fit > Fit Then
            BFDCount = BFDCount + 1
            aTmp(BFDCount) = i
        End If
    Next i
End With
If BFDCount > 0 Then
   ReDim Preserve aTmp(BFDCount)
   GetBadFits = aTmp
Else
   GetBadFits = Null
End If
Exit Function

err_GetBadFits:
GetBadFits = Null
LogErrors Err.Number, "GetBadFits"
End Function

Public Function GetBadStDevs(Ind As Long, StDev As Double) As Variant
'--------------------------------------------------------------------
'Ind Index of GelData, StDev is acceptable standard deviation limit;
'function returns variant array index of bad fits in Iso data arrays
'--------------------------------------------------------------------
Dim i As Long
Dim BadCount As Long
Dim aTmp() As Variant
On Error GoTo err_GetBadStDevs
BadCount = 0
With GelData(Ind)
    ReDim aTmp(.CSLines)
    For i = 1 To .CSLines
        If .CSData(i).MassStDev > StDev Then
            BadCount = BadCount + 1
            aTmp(BadCount) = i
        End If
    Next i
End With
If BadCount > 0 Then
   ReDim Preserve aTmp(BadCount)
   GetBadStDevs = aTmp
Else
   GetBadStDevs = Null
End If
Exit Function

err_GetBadStDevs:
GetBadStDevs = Null
LogErrors Err.Number, "GetBadStDevs"
End Function

Public Sub SortIsotopicData(ByVal Ind As Long)
'---------------------------------------------------------------------------------------
'sort .IsoData on the .IsoDataField inside one scan;
'this procedure is called only when new Display file is created (duplicates elimination)
'---------------------------------------------------------------------------------------
Dim i As Long
Dim nScanCount As Long
Dim nCurrentScan As Long
Dim IsoField As Integer
On Error GoTo err_SortIsotopicData

nCurrentScan = -1
nScanCount = 0
With GelData(Ind)
  If .IsoLines > 1 Then       'nothing to sort if not at least 2 points
     IsoField = .Preferences.IsoDataField
     For i = 1 To .IsoLines
        If nCurrentScan Mod 50 = 0 Then
            ' Yes, I want nCurrentScan Mod 50 but the progress value to be i
            frmProgress.UpdateSubtaskProgressBar i
            If KeyPressAbortProcess > 1 Then Exit For
        End If
        
        nScanCount = 1
        nCurrentScan = .IsoData(i).ScanNumber

        If i < .IsoLines Then
            ' Find the ending point for this scan
            Do While .IsoData(i + nScanCount).ScanNumber = nCurrentScan
                nScanCount = nScanCount + 1
                If i + nScanCount >= .IsoLines Then
                    nScanCount = .IsoLines - i + 1
                    Exit Do
                End If
            Loop
        Else
            nScanCount = 1
        End If
        
        If nScanCount > 1 Then
            ' Now sort the peaks in this scan
            ' Sort .IsoData() on the IsoField mass
            ShellSortIsoArray .IsoData(), i, i + nScanCount - 1, IsoField

            ' Bump up i as required
            i = i + nScanCount - 1
        End If
     Next i
  End If
End With
Exit Sub

err_SortIsotopicData:
Debug.Assert False
LogErrors Err.Number, "SortIsotopicData"
End Sub

Private Sub ShellSortIsoArray(ByRef IsoData() As udtIsotopicDataType, ByVal lngLowIndex As Long, ByVal lngHighIndex As Long, ByVal IsoField As Integer)
'-----------------------------------------------------------------------------------
' sort IsoData() on field IsoField
'-----------------------------------------------------------------------------------

    Dim lngCount As Long
    Dim lngIncrement As Long
    Dim lngIndex As Long
    Dim lngIndexCompare As Long
    Dim udtCompareVal As udtIsotopicDataType
    Dim dblCompareMass As Double

On Error GoTo ShellSortIsoArrayErrorHandler

    ' compute largest increment
    lngCount = lngHighIndex - lngLowIndex + 1
    lngIncrement = 1
    If (lngCount < 14) Then
        lngIncrement = 1
    Else
        Do While lngIncrement < lngCount
            lngIncrement = 3 * lngIncrement + 1
        Loop
        lngIncrement = lngIncrement \ 3
        lngIncrement = lngIncrement \ 3
    End If

    Do While lngIncrement > 0
        ' sort by insertion in increments of lngIncrement
        For lngIndex = lngLowIndex + lngIncrement To lngHighIndex
            udtCompareVal = IsoData(lngIndex)
            dblCompareMass = GetIsoMass(udtCompareVal, IsoField)
            For lngIndexCompare = lngIndex - lngIncrement To lngLowIndex Step -lngIncrement
                ' Use <= to sort ascending; Use > to sort descending
                If GetIsoMass(IsoData(lngIndexCompare), IsoField) <= dblCompareMass Then Exit For
                IsoData(lngIndexCompare + lngIncrement) = IsoData(lngIndexCompare)
            Next lngIndexCompare
            IsoData(lngIndexCompare + lngIncrement) = udtCompareVal
        Next lngIndex
        lngIncrement = lngIncrement \ 3
    Loop
    
    Exit Sub

ShellSortIsoArrayErrorHandler:
    Debug.Assert False
  
End Sub

Public Function GetWorseGuess(ByVal Ind As Long, ByVal Choice As Integer) As Variant
'-----------------------------------------------------------------------------------
'PEK file generated with ICR-2LS lists two guesses in the case of close fit for two
'charge states. The second best fit is marked with asterisk. This function returns
'worse results in variant array  depending of Choice
'Choice=1 return those with * in .IsoVar(i,1); Choice=2 return less likely of two
'based on other results in the same scan
'-----------------------------------------------------------------------------------
Dim i As Long, j As Long, k As Long
Dim bLastScan As Boolean
Dim ScanStart As Long
Dim ScanCount As Long
Dim AsteriskCount As Long
Dim OtherCount As Long
Dim CurrScan As Long
Dim WorseGuessesCount As Long
Dim IsoF As Integer 'just a shortcut
Dim aTmp() As Variant
On Error GoTo err_GetWorseGuess
bLastScan = False
WorseGuessesCount = 0
With GelData(Ind)
  ReDim aTmp(.IsoLines)
  IsoF = .Preferences.IsoDataField
  Select Case Choice
  Case 1 'eliminate second guess
''    For i = 1 To .IsoLines
''      If InStr(1, .IsoVar(i, isvfIsotopeLabel), "*") Then
''        WorseGuessesCount = WorseGuessesCount + 1
''        aTmp(WorseGuessesCount) = i
''      End If
''    Next i
  Case 2  'eliminate less likely
    If .IsoLines > 1 Then
      CurrScan = 0
      ScanCount = 0
      For i = 1 To .IsoLines
        If i = .IsoLines Then   'this is neccessary to take care of the last scan
           ScanCount = ScanCount + 1
           bLastScan = True
        End If
        If (.IsoData(i).ScanNumber = CurrScan) And (Not bLastScan) Then
           ScanCount = ScanCount + 1
        Else      'new scan - do and reset
           If ScanCount > 1 Then
              For j = ScanStart To ScanStart + ScanCount - 1
''                If InStr(1, .IsoVar(j, isvfIsotopeLabel), "*") Then
''                  AsteriskCount = CountWhat(Ind, CurrScan, GetIsoMass(.IsoData(j), IsoF), .Preferences.DupTolerance)
''                  'find row with same m/z (just one possible)
''                  For k = ScanStart To ScanStart + ScanCount - 1
''                    If (k <> j) And (.IsoData(k).MZ = .IsoData(j).MZ) Then Exit For
''                  Next k
''                  If k = ScanStart + ScanCount Then   'if same m/z not found
''                     OtherCount = 0                     'just ignore line
''                  Else
''                     OtherCount = CountWhat(Ind, CurrScan, GetIsoMass(.IsoData(k), IsoF), .Preferences.DupTolerance)
''                     WorseGuessesCount = WorseGuessesCount + 1
''                     If AsteriskCount > OtherCount Then
''                        aTmp(WorseGuessesCount) = k
''                     Else
''                        aTmp(WorseGuessesCount) = j
''                     End If
''                  End If
''                End If
              Next j
           End If
           ScanCount = 1
           CurrScan = .IsoData(i).ScanNumber
           ScanStart = i
        End If
      Next i
    End If
  End Select
End With
If WorseGuessesCount > 0 Then
    ReDim Preserve aTmp(WorseGuessesCount)
    GetWorseGuess = aTmp
Else
    GetWorseGuess = Null
End If
Exit Function

err_GetWorseGuess:
GetWorseGuess = Null
LogErrors Err.Number, "GetWorseGuess"
End Function

Public Function GetDuplicates(iInd As Long, nDT As Double) As Variant
'--------------------------------------------------------------------------
'iInd Index of GelData, nDT duplicate tolerance
'returns variant containing array of indexes of duplicates in .IsoData array
'(most abundant) in the multiplet is not returned as a duplicate
'--------------------------------------------------------------------------
Dim i As Long
Dim DupCount As Long
Dim ScanStart As Long
Dim ScanCount As Long
Dim CurrScan As Long
Dim bDone As Boolean
Dim j As Long
Dim k As Long
Dim IsoField As Integer
Dim aTmp() As Variant
On Error GoTo err_getduplicates
DupCount = 0
ScanCount = 0
CurrScan = 0
With GelData(iInd)
  If .IsoLines > 1 Then        'no sense otherwise
    ReDim aTmp(.IsoLines)
    IsoField = .Preferences.IsoDataField
    For i = 1 To .IsoLines
        If i = .IsoLines Then   'this is neccessary to take care of the last scan
           ScanCount = ScanCount + 1
           CurrScan = -1
        End If
        If .IsoData(i).ScanNumber = CurrScan Then
          ScanCount = ScanCount + 1
        Else                          'new scan; eliminate duplicates
          j = ScanStart              'in previous scan
          Do While j < ScanStart + ScanCount - 1
            k = 1
            bDone = False
            Do While (Not bDone) And (k < ScanStart + ScanCount - j)
              If Abs(GetIsoMass(.IsoData(j), IsoField) - GetIsoMass(.IsoData(j + k), IsoField)) >= nDT Then
                bDone = True
              Else
                If .IsoData(j).Abundance < .IsoData(j + k).Abundance Then
                  DupCount = DupCount + 1
                  aTmp(DupCount) = j
                  bDone = True
                Else
                  DupCount = DupCount + 1
                  aTmp(DupCount) = j + k
                  k = k + 1
                End If
              End If
            Loop
            j = j + k
          Loop
          ScanCount = 1             'and reset counter
          CurrScan = .IsoData(i).ScanNumber
          ScanStart = i
        End If
    Next i
  End If
End With
If DupCount > 0 Then
   ReDim Preserve aTmp(DupCount)
   GetDuplicates = aTmp
Else
   GetDuplicates = Null
End If
Exit Function

err_getduplicates:
GetDuplicates = Null
LogErrors Err.Number, "GetDuplicates"
End Function

' Unused function (February 2005)
''Public Function GetBadDBFits(iInd As Long, nDBFit As Double) As Variant
'''--------------------------------------------------------------------------
''' THIS FUNCTION IS UNUSED
'''
'''iInd Index of GelData, nDBFit is acceptable Fit limit
'''function returns variant containing array of indexes in .CSData and .IsoData
'''first row contains glCSType or glIsoType; second row index
'''rows and columns are here transposed for easier redimensioning
'''--------------------------------------------------------------------------
''Dim i As Long
''Dim IsoField As Integer
''Dim BDBFCount As Long
''Dim aTmp() As Variant
''On Error GoTo err_getbaddbfits
''BDBFCount = 0
''With GelData(iInd)
''    ReDim aTmp(1, .DataLines)
''    If .CSLines > 0 Then
''       For i = 1 To .CSLines
''         If Abs(.CSData(i).AverageMW - .CSData(i).IsotopicFitRatio) > nDBFit Then
''            BDBFCount = BDBFCount + 1
''            aTmp(0, BDBFCount) = glCSType
''            aTmp(1, BDBFCount) = i
''         End If
''       Next i
''    End If
''    IsoField = .Preferences.IsoDataField
''    If .IsoLines > 0 Then
''       For i = 1 To .IsoLines
''         If Abs(GetIsoMass(.IsoData(i), IsoField) - .IsoData(i).IsotopicFitRatio) > nDBFit Then
''            BDBFCount = BDBFCount + 1
''            aTmp(0, BDBFCount) = glIsoType
''            aTmp(1, BDBFCount) = i
''         End If
''       Next i
''    End If
''End With
''If BDBFCount > 0 Then
''   ReDim Preserve aTmp(1, BDBFCount)
''   GetBadDBFits = aTmp
''Else
''   GetBadDBFits = Null
''End If
''Exit Function
''
''err_getbaddbfits:
''GetBadDBFits = Null
''LogErrors Err.Number, "GetBadDBFits"
''End Function

Private Function CountWhat(iInd As Long, ScanNum As Double, MW As Double, DT As Double) As Long
'-------------------------------------------------------------------------------------------------
'This function is just to make code for GetWorseGuess shorter; it returns number of MWs
'same as the MW (same within DT -duplicate tolerance)inside the CS and ISO data arrays
'-------------------------------------------------------------------------------------------------
Dim MyCount As Long
Dim i As Long
Dim IsoField As Integer
MyCount = 0
With GelData(iInd)
   If .CSLines > 0 Then
      For i = 1 To .CSLines
          If .CSData(i).ScanNumber = ScanNum Then
             If Abs(.CSData(i).AverageMW - MW) < DT Then MyCount = MyCount + 1
          End If
      Next i
   End If
   If .IsoLines > 0 Then
      IsoField = .Preferences.IsoDataField
      For i = 1 To .IsoLines
          If .IsoData(i).ScanNumber = ScanNum Then
             If Abs(GetIsoMass(.IsoData(i), IsoField) - MW) < DT Then MyCount = MyCount + 1
          End If
      Next i
   End If
End With
CountWhat = MyCount
End Function


Public Function GetDFIndex(ByVal Ind As Long, ByVal FN As Long) As Long
'-------------------------------------------------------------------------
'retrieves index in DF arrays for specified scan number
'-------------------------------------------------------------------------
Dim i As Long
Dim MaxInd As Long
On Error GoTo exit_GetDFIndex
With GelData(Ind)
    MaxInd = UBound(.ScanInfo)
    If MaxInd > 0 Then
       For i = 1 To MaxInd
           If .ScanInfo(i).ScanNumber = FN Then
              GetDFIndex = i
              Exit Function
           End If
       Next i
    End If
End With
'if here scan number was not found
exit_GetDFIndex:
GetDFIndex = -1
End Function

' Unused Function (March 2003)
'''Public Function GetpIIndex(ByVal Ind As Long, ByVal pi) As Long
''''--------------------------------------------------------------
''''retrieves index in DF arrays for specified scan number
''''--------------------------------------------------------------
'''Dim i As Long
'''Dim MaxInd As Long
'''On Error GoTo err_GetpIIndex
'''With GelData(Ind)
'''    MaxInd = UBound(.DFPI)
'''    If MaxInd > 0 And IsNumeric(pi) Then
'''       For i = 1 To MaxInd
'''           If .DFPI(i) = pi Then
'''              GetpIIndex = i
'''              Exit Function
'''           End If
'''       Next i
'''    End If
'''End With
''''if here pI was not found
'''err_GetpIIndex:
'''GetpIIndex = -1
'''End Function


Public Function MinpIToFN(ByVal Ind As Long, ByVal spI As Double) As Long
'---------------------------------------------------------------------------
'returns FN for closest element of .ScanInfo().ScanPI <=  spI
'---------------------------------------------------------------------------
Dim MaxInd As Long
Dim CurrIndex As Long
Dim bDone As Boolean
On Error GoTo err_MinpIToFN
With GelData(Ind)
   MaxInd = UBound(.ScanInfo)
   If MaxInd > 0 Then
      CurrIndex = 0
      bDone = False
      Do While Not bDone
         CurrIndex = CurrIndex + 1
         bDone = (CurrIndex = MaxInd) Or (.ScanInfo(CurrIndex).ScanPI <= spI)
      Loop
      MinpIToFN = .ScanInfo(CurrIndex).ScanNumber
   End If
End With
Exit Function

err_MinpIToFN:
MinpIToFN = -1
End Function

Public Function MaxpIToFN(ByVal Ind As Long, ByVal spI As Double) As Long
'------------------------------------------------------------------------
'returns FN for closest element of .ScanInfo().ScanPI >=  spI
'------------------------------------------------------------------------
Dim MaxInd As Long
Dim CurrIndex As Long
Dim bDone As Boolean
On Error GoTo err_MaxpIToFN
With GelData(Ind)
   MaxInd = UBound(.ScanInfo)
   If MaxInd > 0 Then
      CurrIndex = MaxInd + 1
      bDone = False
      Do While Not bDone
         CurrIndex = CurrIndex - 1
         bDone = (CurrIndex = 1) Or (.ScanInfo(CurrIndex).ScanPI >= spI)
      Loop
      MaxpIToFN = .ScanInfo(CurrIndex).ScanNumber
   End If
End With
Exit Function

err_MaxpIToFN:
MaxpIToFN = -1
End Function

Public Sub GetMassRangeCurrent(ByVal Ind As Long, ByRef Min As Double, ByRef Max As Double)
'-------------------------------------------------------------------------------------------------------
'Returns the current mass range for gel
'-------------------------------------------------------------------------------------------------------
On Error GoTo err_GetMassRangeCurrent
With GelBody(Ind).csMyCooSys
    If .csYScale = glVAxisLin Then
       Min = .CurrRYMin
       Max = .CurrRYMax
    Else
       Min = 10 ^ .CurrRYMin
       Max = 10 ^ .CurrRYMax
    End If
End With
Exit Sub

err_GetMassRangeCurrent:
Min = 0
Max = 0
End Sub

Public Sub GetScanRangeCurrent(ByVal Ind As Long, ByRef Min As Long, ByRef Max As Long, Optional ByRef Using_pI_Mode As Boolean)
'-------------------------------------------------------------------------------------------------------
'Returns the current scan range for gel; if pI mode is enabled, then returns Min and Max pI
'-------------------------------------------------------------------------------------------------------
On Error GoTo err_GetScanRangeCurrent
With GelBody(Ind).csMyCooSys
     Select Case .csType
     Case glPICooSys
        Using_pI_Mode = True
        Min = GelData(Ind).ScanInfo(GetDFIndex(Ind, .CurrRXMax)).ScanPI
        Max = GelData(Ind).ScanInfo(GetDFIndex(Ind, .CurrRXMin)).ScanPI
     Case glFNCooSys, glNETCooSys
        Using_pI_Mode = False
        Min = .CurrRXMin
        Max = .CurrRXMax
     End Select
End With
Exit Sub

err_GetScanRangeCurrent:
Min = 0
Max = 0
End Sub

Public Sub GetScanRange(ByVal Ind As Long, ByRef ScanMin As Long, ByRef ScanMax As Long, ByRef ScanRange As Long, Optional ByRef ScanCount As Long = 0)
'-------------------------------------------------------------------------------------------------------
'determines scan range for gel (the maximum possible scan range)
'-------------------------------------------------------------------------------------------------------
On Error GoTo err_GetScanRange

ScanMin = 0
ScanMax = 0
ScanRange = 0

With GelData(Ind)
    ScanCount = UBound(.ScanInfo)
    ScanMin = .ScanInfo(1).ScanNumber
    ScanMax = .ScanInfo(UBound(.ScanInfo)).ScanNumber
    ScanRange = ScanMax - ScanMin + 1
End With
Exit Sub

err_GetScanRange:
Debug.Assert False
LogErrors Err.Number, "Arrays.bas->GetScanRange"

End Sub


Public Sub Avg_D(ByRef aVal() As Double, ByRef Cnt As Long, ByRef Res As Double)
'---------------------------------------------------------------------------------
'calculates average value of an array; Res returns average; 'Cnt returns number of
'values included; -1 on any error non-numeric values in an array are ignored
'---------------------------------------------------------------------------------
Dim SM As Double
Dim TmpCnt As Long
Dim MinInd As Long
Dim MaxInd As Long
Dim i As Long
On Error GoTo err_Avg_V

MinInd = LBound(aVal)
MaxInd = UBound(aVal)
SM = 0
TmpCnt = 0
For i = MinInd To MaxInd
    TmpCnt = TmpCnt + 1
    SM = SM + CDbl(aVal(i))
Next i
Cnt = TmpCnt
If TmpCnt > 0 Then Res = SM / TmpCnt
Exit Sub

err_Avg_V:
Cnt = -1
End Sub


'FUNCTIONS RELATED TO MAINTENANCE OF FIELDS IN CSData/IsoData ARRAYS

' Unused Function (March 2003)
'''Public Function RemoveRefFromIDField(ByVal Ind As Long, ByVal Scope As Integer, ByVal RefMark As String)
''''-------------------------------------------------------------------------------------------------------
''''removes all refs. marked with RefMark from the field CS/IsoVar(i,3); which is considered to be ID field
''''-------------------------------------------------------------------------------------------------------
'''Dim i As Long
'''On Error Resume Next
'''Select Case Scope
'''Case glScope.glSc_All
'''  With GelData(Ind)
'''      If .CSLines > 0 Then
'''         For i = 1 To .CSLines
'''             CleanRef .CSData(i).mtid, RefMark
'''         Next i
'''      End If
'''      If .IsoLines > 0 Then
'''         For i = 1 To .IsoLines
'''             CleanRef .IsoData(i).MTID, RefMark
'''         Next i
'''      End If
'''  End With
'''Case glScope.glSc_Current
'''  With GelData(Ind)
'''    If .CSLines > 0 Then
'''       For i = 1 To .CSLines
'''         If GelDraw(Ind).CSID(i) > 0 And GelDraw(Ind).CSR(i) > 0 Then
'''            CleanRef .CSData(i).mtid, RefMark
'''         End If
'''       Next i
'''    End If
'''    If .IsoLines > 0 Then
'''       For i = 1 To .IsoLines
'''         If GelDraw(Ind).IsoID(i) > 0 And GelDraw(Ind).IsoR(i) > 0 Then
'''            CleanRef .IsoData(i).MTID, RefMark
'''         End If
'''       Next i
'''    End If
'''  End With
'''End Select
'''End Function
'''
'''
'''Public Sub CleanRef(ByRef s As Variant, ByVal RefMark As String)
''''---------------------------------------------------------------
''''cleans string s from all references marked as RefMark
''''---------------------------------------------------------------
'''Dim Done As Boolean
'''Dim RefMarkPos As Long
'''Dim TerminatorPos As Long
'''On Error Resume Next
'''If IsNull(s) Then Exit Sub  'do not do anything if already null
'''Do Until Done
'''   RefMarkPos = InStr(1, s, RefMark)
'''   If RefMarkPos > 0 Then
'''      TerminatorPos = InStr(RefMarkPos, s, glARG_SEP)
'''      If TerminatorPos > 0 Then
'''         s = Trim$(Left$(s, RefMarkPos - 1) & Right$(s, Len(s) - TerminatorPos))
'''         If Len(s) <= 0 Then Done = True
'''      Else
'''         s = Trim$(Left$(s, RefMarkPos - 1))
'''         Done = True
'''      End If
'''   Else
'''      Done = True
'''   End If
'''Loop
'''End Sub

''Public Sub CleanVarField(ByVal Ind As Long, ByVal Field As Integer)
'''---------------------------------------------------------------------
'''cleans CS/IsoVar fields in specific gel (sets it to Null)
'''---------------------------------------------------------------------
''Dim i As Long
''With GelData(Ind)
''  If .CSLines > 0 Then
''     For i = 1 To .CSLines
''         If Not IsNull(.CSVar(i, Field)) Then .CSVar(i, Field) = Null
''     Next i
''  End If
''  If .IsoLines > 0 Then
''     For i = 1 To .IsoLines
''         If Not IsNull(.IsoVar(i, Field)) Then .IsoVar(i, Field) = Null
''     Next i
''  End If
''End With
''End Sub

' Unused Function (March 2003)
'''Public Sub RemoveStringFromVarField(ByVal Ind As Long, ByVal s As String, ByVal Field As Integer)
''''------------------------------------------------------------------------------------------------
''''removes string s from Field in CS/IsoVar arrays if found
''''------------------------------------------------------------------------------------------------
'''Dim i As Long
'''With GelData(Ind)
'''  If .CSLines > 0 Then
'''     For i = 1 To .CSLines
'''         RemoveSubstring .CSVar(i, Field), s
'''     Next i
'''  End If
'''  If .IsoLines > 0 Then
'''     For i = 1 To .IsoLines
'''         RemoveSubstring .IsoVar(i, Field), s
'''     Next i
'''  End If
'''End With
'''End Sub

Public Sub Remove1stSubstring(ByRef S As String, ByVal SubS As String)
'---------------------------------------------------------------------
'removes first occurence of the substring SubS from the string S
'---------------------------------------------------------------------
Dim SubSPos As Long
On Error GoTo err_RemoveSubString
If (Len(S) > 0) And (Len(SubS) > 0) Then
   SubSPos = InStr(1, S, SubS)
   If SubSPos > 0 Then
      S = Trim$(Left$(S, SubSPos - 1) & Right$(S, Len(S) - SubSPos - Len(SubS)))
   End If
End If
err_RemoveSubString:
End Sub

Private Sub RemoveSubstring(ByRef vStr As Variant, ByVal SubS As String)
'------------------------------------------------------------------------
'removes all substring SubS occuring in string vStr; trims also remaining
'argument separator ";" if on right side  (lm:04/02/2001;nt)
'------------------------------------------------------------------------
Dim StartPos As Integer
Dim sStr As String
Dim LStr As String          'left portion of the string
Dim RStr As String          'right portion of the string
Dim Done As Boolean
On Error GoTo err_DoNothing
sStr = CStr(vStr)
Do Until Done
   StartPos = InStr(1, sStr, SubS)
   If StartPos > 0 Then
      LStr = Left$(sStr, StartPos - 1)
      RStr = Right$(sStr, Len(sStr) - StartPos - Len(SubS) + 1)
      'trim argument separator and space if still there(in RStr)
      If Left$(RStr, 1) = glARG_SEP Then RStr = Right$(RStr, Len(RStr) - 1)
      sStr = LStr & Trim$(RStr)
   Else
      Done = True
   End If
Loop
vStr = sStr
err_DoNothing:
End Sub


Public Sub InsertBefore(ByRef BeforeWhat As Variant, ByVal InsertWhat As String)
'inserts InsertWhat before whatever is in BeforeWhat and separates it with ";"
Dim TInsertWhat As String
On Error Resume Next
TInsertWhat = Trim$(InsertWhat)
If Len(TInsertWhat) > 0 Then
   If Right$(TInsertWhat, 1) <> glARG_SEP Then
      BeforeWhat = TInsertWhat & glARG_SEP & Chr$(32) & BeforeWhat
   Else
      BeforeWhat = TInsertWhat & Chr$(32) & BeforeWhat
   End If
End If
End Sub


Public Sub RemoveBadMTs_Delta(ByVal Ind As Long)
'----------------------------------------------------------------------------------------------
'removes bad MT marks from gel based on Pairs numbers of deltas removes all MT reference with
'delta information that don't match independent delta reference(IDR); if there is more than one
'IDR all MT delta matches are included;if there is no IDR all MTs are included; if MT does not
'contain delta information it is included (cr:04/02/2001;nt)
'----------------------------------------------------------------------------------------------
Dim i As Long, j As Long, k As Long
Dim MTCnt As Long
Dim MTRef() As String
Dim sMTDlt As String
Dim sIDlt As String
Dim IDRCnt As Long
Dim IDRRef() As String
Dim IDlt() As Long
Dim HasDltMatch As Boolean
On Error Resume Next
With GelData(Ind)
  For i = 1 To .CSLines
    If IsAMTReferenced(.CSData(i).MTID) Then
       IDRCnt = GetTagRefFromString(PAIR_DLT_MARK, CStr(.CSData(i).MTID), IDRRef())
       If IDRCnt > 0 Then
          ReDim IDlt(IDRCnt - 1)
          For j = 0 To IDRCnt - 1
              sIDlt = GetIDFromString(IDRRef(j), PAIR_DLT_MARK)
              If IsNumeric(sIDlt) Then IDlt(j) = CLng(sIDlt)
          Next j
          MTCnt = GetTagRefFromString(AMTMark, CStr(.CSData(i).MTID), MTRef())
          For j = 0 To MTCnt - 1
            sMTDlt = GetIDFromString(MTRef(j), MTDltMark)
            HasDltMatch = False
            If IsNumeric(sMTDlt) Then
               For k = 0 To IDRCnt - 1
                 If CLng(sMTDlt) = IDlt(k) Then
                    HasDltMatch = True
                 End If
               Next k
            End If
            If Not HasDltMatch Then RemoveSubstring .CSData(i).MTID, MTRef(j)
          Next j
       End If
    End If
  Next i
  For i = 1 To .IsoLines
    If IsAMTReferenced(.IsoData(i).MTID) Then
       IDRCnt = GetTagRefFromString(PAIR_DLT_MARK, CStr(.IsoData(i).MTID), IDRRef())
       If IDRCnt > 0 Then
          ReDim IDlt(IDRCnt - 1)
          For j = 0 To IDRCnt - 1
              sIDlt = GetIDFromString(IDRRef(j), PAIR_DLT_MARK)
              If IsNumeric(sIDlt) Then IDlt(j) = CLng(sIDlt)
          Next j
          MTCnt = GetTagRefFromString(AMTMark, CStr(.IsoData(i).MTID), MTRef())
          For j = 0 To MTCnt - 1
            sMTDlt = GetIDFromString(MTRef(j), MTDltMark)
            HasDltMatch = False
            If IsNumeric(sMTDlt) Then
               For k = 0 To IDRCnt - 1
                 If CLng(sMTDlt) = IDlt(k) Then
                    HasDltMatch = True
                 End If
               Next k
            End If
            If Not HasDltMatch Then RemoveSubstring .IsoData(i).MTID, MTRef(j)
          Next j
       End If
    End If
  Next i
End With
GelStatus(Ind).Dirty = True
End Sub

Public Function GetTagRefFromString(ByVal Tag As String, ByVal S As String, _
                                    ByRef TagRef() As String) As Long
'------------------------------------------------------------------------------
'fills array TagRef with all references from string s tagged with Tag reference
'starts on first position after Tag and ends with first argument separator ";"
'after that (AMT:DR2541(MW:0.58ppm);)
'------------------------------------------------------------------------------
Dim TagRefCnt  As Long
Dim TagLen As Long
Dim AllRef() As String
Dim AllRefCnt As Long
Dim i As Long
On Error Resume Next

TagLen = Len(Tag)
If (Len(S) > 0 And TagLen > 0) Then             'pick all references
   AllRef = Split(S, glARG_SEP)
   AllRefCnt = UBound(AllRef) + 1
End If
If AllRefCnt > 0 Then
   ReDim TagRef(AllRefCnt - 1)
   For i = 0 To AllRefCnt - 1
       AllRef(i) = Trim$(AllRef(i))
       If Left$(AllRef(i), TagLen) = Tag Then   'select desired tag references
          TagRefCnt = TagRefCnt + 1
          TagRef(TagRefCnt - 1) = AllRef(i)
       End If
   Next i
End If
If TagRefCnt > 0 Then                           'trim array
   ReDim Preserve TagRef(TagRefCnt - 1)
Else
   Erase TagRef
End If
GetTagRefFromString = TagRefCnt
End Function


Public Sub RemoveBadMTs_DeltaMT(ByVal Ind As Long, ByVal DeltaType As Long)
'------------------------------------------------------------------
'removes bad MT marks from gel based on number of deltas
'DeltaType can be N14/N15 or ICAT
'cr:04/02/2001;nt
'------------------------------------------------------------------
Dim i As Long, j As Long
Dim MTCnt As Long
Dim MTRef() As String
Dim sMTDlt As String
Dim sMTID As String
Dim MTIDInd As Long

Dim lngAMTID() As Long
Dim objAMTIDFastSearch As FastSearchArrayLong

Dim lngIndex As Long
Dim lngMatchCount As Long
Dim lngMatchingIndices() As Long

On Error Resume Next
If AMTGeneration < dbgGeneration1000 Then
   MsgBox "MT tag database does not support this feature.", vbOKOnly, glFGTU
   Exit Sub
End If

' Construct the MT tag ID lookup arrays
' We need to copy the AMT ID's from AMTData() to lngAMTID() since AMTData().ID is a String array that actually simply holds numbers
If AMTCnt > 0 Then
    ReDim lngAMTID(1 To AMTCnt)
    For lngIndex = 1 To AMTCnt
        lngAMTID(lngIndex) = CLngSafe(AMTData(lngIndex).ID)
    Next lngIndex
Else
    ReDim lngAMTID(1 To 1)
End If

' Initialize objAMTIDFastSearch
Set objAMTIDFastSearch = New FastSearchArrayLong
If Not objAMTIDFastSearch.Fill(lngAMTID()) Then
    Exit Sub
End If

With GelData(Ind)
  For i = 1 To .CSLines
    If IsAMTReferenced(.CSData(i).MTID) Then
        MTCnt = GetTagRefFromString(AMTMark, CStr(.CSData(i).MTID), MTRef())
        For j = 0 To MTCnt - 1
            sMTDlt = GetIDFromString(MTRef(j), MTDltMark)
            sMTID = GetIDFromString(MTRef(j), AMTMark, AMTIDEnd)
         
            If IsNumeric(sMTID) And IsNumeric(sMTDlt) Then
                If objAMTIDFastSearch.FindMatchingIndices(CLng(sMTID), lngMatchingIndices(), lngMatchCount) Then
                    MTIDInd = lngMatchingIndices(0)
                    Select Case DeltaType
                    Case PAIR_ICAT
                        If AMTData(MTIDInd).CNT_Cys < CLng(sMTDlt) Then
                            RemoveSubstring .CSData(i).MTID, MTRef(j)
                        End If
                    Case PAIR_N14N15
                        If AMTData(MTIDInd).CNT_N < CLng(sMTDlt) Then
                            RemoveSubstring .CSData(i).MTID, MTRef(j)
                        End If
                    End Select
                End If
            End If
       Next j
    End If
  Next i
  For i = 1 To .IsoLines
    If IsAMTReferenced(.IsoData(i).MTID) Then
        MTCnt = GetTagRefFromString(AMTMark, CStr(.IsoData(i).MTID), MTRef())
        For j = 0 To MTCnt - 1
            sMTDlt = GetIDFromString(MTRef(j), MTDltMark)
            sMTID = GetIDFromString(MTRef(j), AMTMark, AMTIDEnd)
                        
            If IsNumeric(sMTID) And IsNumeric(sMTDlt) Then
                If objAMTIDFastSearch.FindMatchingIndices(CLng(sMTID), lngMatchingIndices(), lngMatchCount) Then
                    MTIDInd = lngMatchingIndices(0)
                    Select Case DeltaType
                    Case PAIR_ICAT
                        If AMTData(MTIDInd).CNT_Cys < CLng(sMTDlt) Then
                            RemoveSubstring .IsoData(i).MTID, MTRef(j)
                        End If
                    Case PAIR_N14N15
                        If AMTData(MTIDInd).CNT_N < CLng(sMTDlt) Then
                            RemoveSubstring .IsoData(i).MTID, MTRef(j)
                        End If
                    End Select
                End If
            End If
        Next j
    End If
  Next i
End With
GelStatus(Ind).Dirty = True
End Sub

Public Sub RemoveIDWithoutID(ByVal Ind As Long)
'-------------------------------------------------------
'clears all IDs that does not include Database reference
'-------------------------------------------------------
Dim i As Long
On Error Resume Next
With GelData(Ind)
  For i = 1 To .CSLines
    If Not IsAMTReferenced(.CSData(i).MTID) Then .CSData(i).MTID = ""
  Next i
  For i = 1 To .IsoLines
    If Not IsAMTReferenced(.IsoData(i).MTID) Then .IsoData(i).MTID = ""
  Next i
End With
GelStatus(Ind).Dirty = True
End Sub


Public Sub CleanIDData(ByVal Ind As Long)
'----------------------------------------
'clears all IDs
'----------------------------------------
Dim i As Long
On Error Resume Next
With GelData(Ind)
  For i = 1 To .CSLines
    .CSData(i).MTID = ""
  Next i
  For i = 1 To .IsoLines
    .IsoData(i).MTID = ""
  Next i
End With
GelStatus(Ind).Dirty = True
End Sub

Public Function Sort2LongArrays(Ind1() As Long, Ind2() As Long, SortInd() As Long) As Boolean
'---------------------------------------------------------------------------------------------
'sorts 2 arrays of longs on Ind1/Ind2; returns True if successful
'NOTE: sorting using QSLongIndOnly class should be slightly faster than QSLong class
'---------------------------------------------------------------------------------------------
Dim ArrCnt As Long
Dim MySort As New QSLongIndOnly           'use sort that does not change order of original arrays
Dim MyScanSort As QSLongIndOnly          'this will sort just temporary arrays
Dim Scan() As Long, ScanInd() As Long   'used to sort on second index
Dim i As Long, j As Long, tmp As Long
Dim CurrInd1 As Long
Dim ScanStartPos As Long, ScanEndPos As Long, ScanCnt As Long
On Error GoTo err_Sort2LongArrays
ArrCnt = UBound(SortInd) + 1
If ArrCnt > 0 Then
   If ArrCnt > 1 Then                                   'otherwise nothing to sort
      If MySort.QSAsc(Ind1(), SortInd()) Then           'sort fast on first index
         'now sort scans of same first index rows on second index
         CurrInd1 = -1:      ScanCnt = 0
         For i = 0 To ArrCnt - 1
             If Ind1(SortInd(i)) = CurrInd1 Then           'still in the same Scan
                ScanEndPos = i                            'index in SortInd array
                ScanCnt = ScanCnt + 1
             Else                                          'new Scan; sort previous if anything to sort
                If ScanCnt > 0 Then                       'nothing to sort otherwise
                   If ScanCnt > 1 Then                    'nothing to sort otherwise
                      If ScanCnt > 2 Then                 'sort
                         ReDim Scan(ScanCnt - 1):   ReDim ScanInd(ScanCnt - 1)
                         For j = ScanStartPos To ScanEndPos
                             Scan(j - ScanStartPos) = Ind2(SortInd(j))
                             ScanInd(j - ScanStartPos) = SortInd(j)
                         Next j
                         Set MyScanSort = New QSLongIndOnly
                         If MyScanSort.QSAsc(Scan(), ScanInd()) Then
                            For j = ScanStartPos To ScanEndPos
                                SortInd(j) = ScanInd(j - ScanStartPos)
                            Next j
                         End If
                         Set MyScanSort = Nothing
                      Else                                 'sort simple
                         If Ind2(SortInd(ScanStartPos)) > Ind2(SortInd(ScanEndPos)) Then
                            tmp = SortInd(ScanStartPos)
                            SortInd(ScanStartPos) = SortInd(ScanEndPos)
                            SortInd(ScanEndPos) = tmp
                         End If
                      End If
                   End If
                End If
                CurrInd1 = Ind1(SortInd(i))                'start new Scan
                ScanStartPos = i:   ScanEndPos = i:   ScanCnt = 1
             End If
         Next i
         'sort last Scan
         If ScanCnt > 0 Then                       'nothing to sort otherwise
            If ScanCnt > 1 Then                    'nothing to sort otherwise
               If ScanCnt > 2 Then                 'sort
                  ReDim Scan(ScanCnt - 1):   ReDim ScanInd(ScanCnt - 1)
                  For j = ScanStartPos To ScanEndPos
                      Scan(j - ScanStartPos) = Ind2(SortInd(j))
                      ScanInd(j - ScanStartPos) = SortInd(j)
                  Next j
                  Set MyScanSort = New QSLongIndOnly
                  If MyScanSort.QSAsc(Scan(), ScanInd()) Then
                     For j = ScanStartPos To ScanEndPos
                         SortInd(j) = ScanInd(j - ScanStartPos)
                     Next j
                  End If
                  Set MyScanSort = Nothing
               Else                                 'sort simple
                  If Ind2(SortInd(ScanStartPos)) > Ind2(SortInd(ScanEndPos)) Then
                     tmp = SortInd(ScanStartPos)
                     SortInd(ScanStartPos) = SortInd(ScanEndPos)
                     SortInd(ScanEndPos) = tmp
                  End If
               End If
            End If
         End If
         Sort2LongArrays = True
      End If
   End If
End If

exit_Sort2LongArrays:
Set MySort = Nothing
Exit Function

err_Sort2LongArrays:
LogErrors Err.Number, "Sort2LongArrays"
Resume exit_Sort2LongArrays
End Function
