VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScanGapPattern"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'implements scan gaps
'created: 03/17/2003 nt
'last modified: 03/24/2003 nt
'-------------------------------------------------------------------------
Option Explicit

Dim mMaxGapSize As Long
Dim mMaxGapCount As Long
Dim mBestScore As Long          'might need exposed best score of patterns

Dim mScanCnt As Long
Dim mScans() As Long            'scans array
Dim mInd() As Long              'index


Dim MFirstScan As Long
Dim mLastScan As Long
Dim mPatternCnt As Long
Dim mPatterns() As Long

Dim mScanZero As Long

'public properties
'-------------------------------------------------------------
Public Property Get MaxGapSize() As Long
MaxGapSize = mMaxGapSize
End Property

Public Property Let MaxGapSize(ByVal NewMaxGapSize As Long)
mMaxGapSize = NewMaxGapSize
End Property

Public Property Get MaxGapCount() As Long
MaxGapCount = mMaxGapCount
End Property

Public Property Let MaxGapCount(ByVal NewMaxGapCount As Long)
mMaxGapCount = NewMaxGapCount
End Property

Public Property Get BestScore() As Long
BestScore = mBestScore
End Property


'-------------------------------------------------------------

Public Function ProcessScanPattern(Scans() As Long, ScanZero As Long, _
                                   ResInd() As Long) As Long
'--------------------------------------------------------------------------------
'process scan patterns and returns number of accepted members
'Scans is a list of scans that has to be tested, ScanZero is scan number that
'has to be included; ResInd() is list of indexes in original Scans array that
'is accepted
'--------------------------------------------------------------------------------
Dim i As Long
Dim j As Long
Dim QSL As New QSLong          'sort object
Dim IncCnt As Long             'count of accepted members
Dim lngScanCorrection As Long
Dim lngNextScanNumber As Long

On Error Resume Next
mBestScore = 0                 'reset best score
mScanZero = ScanZero
mScanCnt = UBound(Scans) + 1
If mScanCnt > 0 Then
    ' Copy scans from Scans() to mScans()
    ReDim mScans(mScanCnt - 1)
    ReDim mInd(mScanCnt - 1)
    For i = 0 To mScanCnt - 1
         mScans(i) = Scans(i)
         mInd(i) = i
    Next i
   
    ' Sort on scan number
    If QSL.QSAsc(mScans(), mInd()) Then
        MFirstScan = mScans(0)                'this maybe can be optimized by limiting range
        mLastScan = mScans(mScanCnt - 1)      'to possible values of first and last scan
        mPatternCnt = mLastScan - MFirstScan + 1
        If mPatternCnt > 0 Then
           ReDim mPatterns(mPatternCnt - 1)
           For i = 0 To mScanCnt - 1
              'multiple members could set Patterns to 1; we don't care which
              'since all members from the same scan will be included or excluded
               mPatterns(mScans(i) - MFirstScan) = 1
           Next i
           'here explore patterns and see what should be included
           If ExplorePattern1() Then
              'then organize returned list of indexes;
              'NOTE that returned indexes will be ordered by scan number with
              'uncertainty for members coming from the same scan
              IncCnt = 0
              For i = 0 To mScanCnt - 1
                  If mPatterns(mScans(i) - MFirstScan) > 1 Then
                     IncCnt = IncCnt + 1
                     ResInd(IncCnt - 1) = mInd(i)       'return original index
                  End If
              Next i
              'trim results if neccessary
              If IncCnt > 0 Then
                 ReDim Preserve ResInd(IncCnt - 1)
              Else
                 Erase ResInd
              End If
              ProcessScanPattern = IncCnt               'return number of accepted members
           End If
        End If
    End If
End If
Set QSL = Nothing
End Function


Public Function ExplorePattern1() As Boolean
'---------------------------------------------------------------------------------------
'increases value for all pattern members that are acceptable with current gap definition
'this is alternative approach to pattern lookup; count all gaps left and right and all
'non-gaps left and right and maximize the count
'---------------------------------------------------------------------------------------
Dim ZeroInd As Long
Dim bDone As Boolean
Dim PatternGaps() As Long               'count of gaps up to position (starting from ZeroInd)
Dim PatternNonGaps() As Long            'count of non-gaps up to position(starting from ZeroInd)
Dim PatternSISH() As Long               'scoring array(number of non-gap scans for each position)
Dim GapsCnt As Long                     'current count of gaps(helper variable)
Dim NonGapsCnt As Long                  'current count of scans out of gaps(helper variable)
Dim IsGap As Boolean                    'helper variable to follow gap measurements
Dim CurrGapSize As Long                 'size of a current gap
Dim i As Long, j As Long
Dim GapsAllowedOnOpositeSide As Long    'used while scoring each position
Dim HiSISH_L As Long                    'highest score on left
Dim HiSISH_L_Start As Long              'start of best score on the left
Dim HiSISH_L_End As Long                'end of the best score on the the left
Dim HiSISH_R As Long                    'highest score on right
Dim HiSISH_R_Start As Long              'start of best score on the right
Dim HiSISH_R_End As Long                'end of the best score on the the right
Dim BestIndStart As Long
Dim BestIndEnd As Long
On Error GoTo exit_ExplorePattern1

ZeroInd = mScanZero - MFirstScan            'index in mPattern of mScanZero
ReDim PatternGaps(mPatternCnt - 1)          'gaps count
ReDim PatternNonGaps(mPatternCnt - 1)       'non-gaps scan count
ReDim PatternSISH(mPatternCnt - 1)          'scoring array that will detrmine class

'process gaps and non-gaps left of the ZeroInd
IsGap = False               'can not be gap at ZeroInd (ScanZero)
GapsCnt = 0
NonGapsCnt = 0
CurrGapSize = 0
bDone = False
i = ZeroInd
Do Until bDone
   If mPatterns(i) = 0 Then                 'gap scan
      If Not IsGap Then                     'new gap
         IsGap = True
         GapsCnt = GapsCnt + 1
      End If
      CurrGapSize = CurrGapSize + 1
      'no point at going any further if gap too big or too many gaps already found
      If ((GapsCnt > mMaxGapCount) Or (CurrGapSize > mMaxGapSize)) Then bDone = True
   Else                                     'non gap scan
      If IsGap Then                         'if it was gap before
         IsGap = False                      'it is not any more, therefore
         CurrGapSize = 0                    'reset measure of the size of current gap
      End If
      NonGapsCnt = NonGapsCnt + 1
   End If
   If Not bDone Then                        'everything beyond stop point stays with value 0
                                            'in PatternGaps and PatternNonGaps arrays
      PatternGaps(i) = GapsCnt
      PatternNonGaps(i) = NonGapsCnt
      If i > 0 Then
         i = i - 1                          'going left
      Else
         bDone = True
      End If
   End If
Loop

'process gaps and non-gaps right of the ZeroInd
IsGap = False               'can not be gap at ZeroInd (ScanZero)
GapsCnt = 0
NonGapsCnt = 0
CurrGapSize = 0
bDone = False
i = ZeroInd
Do Until bDone
   If mPatterns(i) = 0 Then                 'gap scan
      If Not IsGap Then                     'new gap
         IsGap = True
         GapsCnt = GapsCnt + 1
      End If
      CurrGapSize = CurrGapSize + 1
      'no point at going any further if gap too big or too many gaps already found
      If ((GapsCnt > mMaxGapCount) Or (CurrGapSize > mMaxGapSize)) Then bDone = True
   Else                                     'non gap scan
      If IsGap Then                         'if it was gap before
         IsGap = False                      'it is not any more, therefore
         CurrGapSize = 0                    'reset measure of the size of current gap
      End If
      NonGapsCnt = NonGapsCnt + 1
   End If
   If Not bDone Then                        'everything beyond stop point stays with value 0
                                            'in PatternGaps and PatternNonGaps arrays
      PatternGaps(i) = GapsCnt
      PatternNonGaps(i) = NonGapsCnt
      If i < mPatternCnt - 1 Then
         i = i + 1                          'going right
      Else
         bDone = True
      End If
   End If
Loop
'now score each individual scan in Pattern with number of non-gap scans
'to get the optimal members (those with highest scores)
'NOTE: this has to be done from both sides
HiSISH_L = 0
For i = 0 To ZeroInd
    If mPatterns(i) > 0 Then                       'no need to score gaps
       If PatternNonGaps(i) > 0 Then               'or already rejected
          GapsAllowedOnOpositeSide = mMaxGapCount - PatternGaps(i)
          'start from the end and find first member before ZeroInd that has
          'less of equal gaps listed than allowed on oposite side
          bDone = False
          j = mPatternCnt - 1
          Do Until bDone
             If j > ZeroInd Then
                If PatternNonGaps(j) > 0 Then         'use only pattern members that were not rejected
                   If PatternGaps(j) <= GapsAllowedOnOpositeSide Then
                      bDone = True
                   End If
                End If
             Else
                bDone = True
             End If
             If Not bDone Then j = j - 1
          Loop
          'now construct result for position i
          PatternSISH(i) = PatternNonGaps(i) + PatternNonGaps(j) - 1
          'and remember best scores range
          If PatternSISH(i) > HiSISH_L Then
             HiSISH_L = PatternSISH(i)
             HiSISH_L_Start = i
             HiSISH_L_End = j
          End If
       End If
    End If
Next i


HiSISH_R = 0
For i = mPatternCnt - 1 To ZeroInd Step -1
    If mPatterns(i) > 0 Then                       'no need to score gaps
       If PatternNonGaps(i) > 0 Then               'or already rejected
          GapsAllowedOnOpositeSide = mMaxGapCount - PatternGaps(i)
          'start from the begining and find first member before ZeroInd that has
          'less of equal gaps listed than allowed on oposite side
          bDone = False
          j = 0
          Do Until bDone
             If j < ZeroInd Then
                If PatternNonGaps(j) > 0 Then         'use only pattern members that were not rejected
                   If PatternGaps(j) <= GapsAllowedOnOpositeSide Then
                      bDone = True
                   End If
                End If
             Else
                bDone = True
             End If
             If Not bDone Then j = j + 1            'going right
          Loop
          'now construct result for position i
          PatternSISH(i) = PatternNonGaps(i) + PatternNonGaps(j) - 1
          'and remember best scores range
          If PatternSISH(i) > HiSISH_R Then
             HiSISH_R = PatternSISH(i)
             HiSISH_R_End = i                   'note reversal of Start End
             HiSISH_R_Start = j
          End If
       End If
    End If
Next i

If HiSISH_L > HiSISH_R Then
   BestIndStart = HiSISH_L_Start
   BestIndEnd = HiSISH_L_End
   mBestScore = HiSISH_L                    'keep the best score
ElseIf HiSISH_L < HiSISH_R Then
   BestIndStart = HiSISH_R_Start
   BestIndEnd = HiSISH_R_End
   mBestScore = HiSISH_R                    'keep the best score
Else                                        'same;
   BestIndStart = HiSISH_L_Start            'take start of left and
   BestIndEnd = HiSISH_R_End                'end of right as boundary
   mBestScore = HiSISH_L                    'keep the best score
End If
For i = BestIndStart To BestIndEnd          'mark range; that is it
    mPatterns(i) = mPatterns(i) + 1
Next i

ExplorePattern1 = True
exit_ExplorePattern1:
End Function

