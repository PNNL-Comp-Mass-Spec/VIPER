Attribute VB_Name = "Module9"
'Internet search related module
'last modified 06/13/2000 nt
Option Explicit
Option Compare Text
'private constants
Const sDBURL = "Protorado.MDB"          'name of database with URLs of the Internet databases
Const s2lsURLTbl = "ToolsSearch"        'name of table in Protorado
Const sIDURLTbl = "Identification"      'name of table in Protorado
Const StartTag = "ST="
Const EndTag = "EN="
Const AddTag = "+++"
Const OrigSep = "/"
Const DoNotUseCh = "!"

Public Const glARG_SEP = ";"
Public Const sMyHTML = "2DGel.htm"             'name of HTML file used to navigate the Net
Public Const NoHarvest = "Not found!"
Public Const NoSearchDone = "Search not performed on this database!"
Public Const ErrSearchNetSuccess = 0    'results of the SearchNet function
Public Const ErrSearchNetNotFound = 1   'identification for this database not found
Public Const ErrSearchNetNoHarvest = 2  'explicit Not Found for this database
Public Const ErrSearchNetOther = 3      'other errors

Const HTMLNoMatchesTag = "NO="
Const HTMLFoundTag = "OK="
Const HTMLExpLenTag = "EL="         'expected length of tags
Const HTMLMarkTag = "MR="           'mark each entry with this
Const HTMLRuleSep = "|"
Const HTMLMaxEntries = 100

Public iNetStatus As Integer
Public hBrowser As Long             'handle to the Browser instance
Public dbURL As Database
Public sURLDBName As String
Public a2lsURL()        '
Public aIDURL()

Public i2lsURLCnt As Integer
Public iIDURLCnt As Integer

'used to extract desired information from the HTML
Public aNFRules(1 To 100) As String
Public aOKRules(1 To 100) As String
Public iNFRulesCnt As Integer
Public iOKRulesCnt As Integer
Public MinEntryLen As Integer
Public MaxEntryLen As Integer
Public sEntryMark As String


Public Function DefaultBrowser() As Integer
'determines and launches system default internet browser
'path to the browser EXE file is saved in global variable sBrowserEXE
'returns True on success, False on failure
Dim Dummy As String
Dim ExecFile As String * 255
Dim res As Long
On Error GoTo err_DefaultBrowser

DefaultBrowser = netNoDefBrowser

ExecFile = Space(255)
res = FindExecutable(sHTMLFile, Dummy, ExecFile)
If res > 32 Then
   ExecFile = Trim(ExecFile)
   If Not IsEmpty(ExecFile) Then
      sBrowserEXE = ExecFile
      DefaultBrowser = netEnabled
      'res = ShellExecute(MDIForm1.hWnd, "open", ExecFile, sHTMLFile, Dummy, SW_SHOWNORMAL)
      'If res > 32 Then DefaultBrowser = netEnabled
   End If
End If
Exit Function

err_DefaultBrowser:
MsgBox "Error: " & Err.Number & " - " & Err.Description, vbOKOnly
End Function

Public Sub InitInternetSearch()
   iNetStatus = DefaultBrowser() + Get2lsURLs()
   Select Case iNetStatus
   Case netEnabled
   Case netSearchURLMissing
      MsgBox "Internet search disabled(error loading URL tables.)", vbOKOnly
   Case netIdURLMissing
      MsgBox "Internet identification disabled(error loading URL tables.)", vbOKOnly
   Case netNoDefBrowser
      MsgBox "Internet search disabled(couldn't start default browser.)", vbOKOnly
   Case Else
      MsgBox "Internet functions disabled(error loading URL tables.)", vbOKOnly
   End Select
End Sub

Public Function BrowseTheNet(ByVal sTarget As String) As Boolean
Dim res As Long, Dummy As Long
res = ShellExecute(MDIForm1.hWnd, "open", sBrowserEXE, sTarget, Dummy, SW_SHOWMAXIMIZED)
If res > 32 Then
   BrowseTheNet = True
Else
   BrowseTheNet = False
End If
End Function

Public Function Get2lsURLs() As Integer
Dim sDir As String
Dim res As Integer
On Error Resume Next

sDir = App.Path
sURLDBName = sDir & "\" & sDBURL
Set dbURL = DBEngine.Workspaces(0).OpenDatabase(sURLDBName)
If Err Then
   Get2lsURLs = netSearchURLMissing + netIdURLMissing
   Exit Function
End If

i2lsURLCnt = GetURLTables(s2lsURLTbl, a2lsURL())
If i2lsURLCnt = 0 Then res = netSearchURLMissing

iIDURLCnt = GetURLTables(sIDURLTbl, aIDURL())
If iIDURLCnt = 0 Then res = res + netIdURLMissing
Get2lsURLs = res

dbURL.Close
Set dbURL = Nothing
End Function

Public Function GetURLTables(ByVal sTbl As String, A()) As Integer
Dim rsTbl As Recordset
Dim idName As Index
Dim iFieldsCnt As Integer
Dim iRecordsCnt As Integer
Dim r As Integer, f As Integer
On Error GoTo err_GetURLTables

GetURLTables = 0
Set rsTbl = dbURL.OpenRecordset(sTbl, dbOpenTable)
rsTbl.Index = "Name"
rsTbl.MoveFirst
iFieldsCnt = rsTbl.Fields.Count
iRecordsCnt = rsTbl.RecordCount
If iRecordsCnt > 0 And iFieldsCnt > 0 Then
   ReDim A(1 To iRecordsCnt, 1 To iFieldsCnt)
   For r = 1 To iRecordsCnt
     For f = 0 To iFieldsCnt - 1
         A(r, f + 1) = rsTbl.Fields(f).Value
     Next f
     rsTbl.MoveNext
   Next r
   GetURLTables = iRecordsCnt
End If
err_GetURLTables:
Set rsTbl = Nothing
End Function

Public Function GetQueryT(ByVal sQuery, ByVal sRule) As String
'returns query to use in Internet identification process
'sQuery - query in the form presented on the frmNetFind
'sRule - instruction how to translate sQuery to specific Internet DB
Dim aOrQrParts(100) As String
Dim aTrQrParts(100) As String
Dim sRulePart As String
Dim iOrQrCnt As Integer
Dim iOrQrUsedCnt As Integer
Dim iPosStart As Integer, iPosSep  As Integer
Dim sTrTmp As String
Dim sTrSep As String
Dim sAddStart As String
Dim sAddEnd As String
Dim i As Integer

iOrQrCnt = 0
iPosSep = 0
Do While Len(sQuery) > 0
   iPosStart = 1
   iPosSep = InStr(iPosStart, sQuery, OrigSep)
   If iPosSep > iPosStart Then
      iOrQrCnt = iOrQrCnt + 1
      aOrQrParts(iOrQrCnt) = Left$(sQuery, iPosSep - 1)
      sQuery = Right$(sQuery, Len(sQuery) - iPosSep)
   Else
      If Len(sQuery) > 0 Then
         iOrQrCnt = iOrQrCnt + 1
         aOrQrParts(iOrQrCnt) = sQuery
         sQuery = ""
      End If
   End If
Loop

If iOrQrCnt > 0 And Len(sRule) > 3 Then
   
   For i = 1 To iOrQrCnt
       aTrQrParts(i) = ApplyRule(aOrQrParts(i), GetRulePart(Left$(aOrQrParts(i), 3), sRule))
   Next i
   
   sTrSep = GetRuleSeparator(sRule)
   sTrTmp = ""
   iOrQrUsedCnt = 0
   For i = 1 To iOrQrCnt
       If Len(aTrQrParts(i)) > 0 Then
          iOrQrUsedCnt = iOrQrUsedCnt + 1
          If iOrQrUsedCnt = 1 Then
             sTrTmp = aTrQrParts(i)
          Else
             sTrTmp = sTrTmp & sTrSep & aTrQrParts(i)
          End If
       End If
   Next i
   sAddStart = Trim(GetRulePart(StartTag, sRule))
   If Len(sAddStart) > 0 Then sAddStart = Right$(sAddStart, Len(sAddStart) - 3)
   sAddEnd = Trim(GetRulePart(EndTag, sRule))
   If Len(sAddEnd) > 0 Then sAddEnd = Right$(sAddEnd, Len(sAddEnd) - 3)
   GetQueryT = Trim(sAddStart & sTrTmp & sAddEnd)
Else
   GetQueryT = ""
End If
End Function

Public Function GetRuleSeparator(ByVal sRule As String) As String
'retrieves from the rule separator (connector) for various criteria
Dim i As Integer
i = InStr(1, sRule, ")")
GetRuleSeparator = Mid$(sRule, 2, i - 2)
End Function

Public Function GetRulePart(ByVal sTag As String, ByVal sRule As String) As String
'retrieves individual rules - sandwiched with ()
Dim iStartPos As Integer
Dim iEndPos As Integer
Dim sPartOfRule As String
If Len(sRule) > 0 Then
   iStartPos = 1
   iEndPos = 0
   Do While iStartPos < Len(sRule)
      iStartPos = InStr(iStartPos, sRule, "(")
      iEndPos = InStr(iStartPos + 1, sRule, ")")
      If iEndPos > iStartPos Then
         sPartOfRule = Mid$(sRule, iStartPos + 1, iEndPos - iStartPos - 1)
         iStartPos = iEndPos
         If InStr(1, sPartOfRule, sTag) > 0 Then
            GetRulePart = sPartOfRule
            Exit Function
         End If
      Else
         iStartPos = Len(sRule)
      End If
   Loop
Else
   GetRulePart = ""
End If
End Function

Public Function ApplyRule(ByVal sQ As String, ByVal sR As String) As String
'applies part of the rule to part of the query
Dim iOrArgCnt As Integer
Dim iRuArgCnt As Integer
Dim sRuArgSep As String
Dim sTmpRes As String
Dim iStartPos As Integer
Dim iEndPos As Integer
Dim iRCnt As Integer
Dim aRule(1 To 5) As String
Dim sOrArg As String
Dim nArgFactor As Double
Dim nArg As Double
Dim sArg As String
On Error GoTo err_ApplyRule

If Len(sQ) > 0 And Len(sR) > 10 And Mid$(sR, 5, 1) <> DoNotUseCh Then
   iStartPos = 1
   iEndPos = 0
   iRCnt = 0
   Do While iStartPos < Len(sR)
      iStartPos = InStr(iStartPos, sR, "<")
      iEndPos = InStr(iStartPos + 1, sR, ">")
      If iEndPos > iStartPos Then
         iRCnt = iRCnt + 1
         aRule(iRCnt) = Mid$(sR, iStartPos + 1, iEndPos - iStartPos - 1)
         iStartPos = iEndPos
      Else
         iStartPos = Len(sR)
      End If
   Loop
   If iRCnt <> 5 Then GoTo err_ApplyRule
   sTmpRes = aRule(1)
   iRuArgCnt = CInt(aRule(2))
   sRuArgSep = aRule(3)
   nArgFactor = 10 ^ CDbl(aRule(4))
   sOrArg = Right(sQ, Len(sQ) - 3)
   
   If Left(sR, 3) = AddTag Then     ' this is where we go to direct browsing, sQ is here ID for whatever
      If Len(aRule(5)) > 0 Then
         ApplyRule = sTmpRes & sOrArg & aRule(5)
      Else
         ApplyRule = sTmpRes & sOrArg
      End If
      Exit Function
   End If
   
   If Len(sOrArg) > 0 Then
      iOrArgCnt = 0
      iStartPos = 1
      iEndPos = 0
      Do While iEndPos < Len(sOrArg) And iOrArgCnt < iRuArgCnt
         iEndPos = InStr(iStartPos, sOrArg, glARG_SEP)
         If iEndPos > iStartPos Then
            If nArgFactor <> 1 Then
               nArg = CDbl(Mid$(sOrArg, iStartPos, iEndPos - iStartPos))
               sArg = Format(nArg * nArgFactor, "#,###,##0.0000")
            Else
               sArg = Mid$(sOrArg, iStartPos, iEndPos - iStartPos)
            End If
            If iOrArgCnt > 0 Then
               sTmpRes = sTmpRes & sRuArgSep & sArg
            Else
               sTmpRes = sTmpRes & sArg
            End If
            iOrArgCnt = iOrArgCnt + 1
            iStartPos = iEndPos + 1
         Else
            If Len(Right$(sOrArg, Len(sOrArg) - iStartPos + 1)) > 0 Then
               If nArgFactor <> 1 Then
                  nArg = CDbl(Right$(sOrArg, Len(sOrArg) - iStartPos + 1))
                  sArg = Format(nArg * nArgFactor, "#,###,##0.0000")
               Else
                  sArg = Right$(sOrArg, Len(sOrArg) - iStartPos + 1)
               End If
               If iOrArgCnt > 0 Then
                  sTmpRes = sTmpRes & sRuArgSep & sArg
               Else
                  sTmpRes = sTmpRes & sArg
               End If
            End If
            iEndPos = Len(sOrArg)
         End If
      Loop
   End If
   If Len(aRule(5)) > 0 Then
      ApplyRule = sTmpRes & aRule(5)
   Else
      ApplyRule = sTmpRes
   End If
Else
   ApplyRule = ""
End If
Exit Function

err_ApplyRule:
ApplyRule = ""
End Function


Public Function HarvestHTML(ByVal sHTML As String) As String
'returns desired information from the HTML file
Dim iStartPos As Long
Dim iEndPos As Long
Dim sHarvest As String
Dim i As Integer
On Error GoTo err_HarvestHTML

If Len(sHTML) > 0 Then
   If iNFRulesCnt = 0 And iOKRulesCnt = 0 Then GoTo err_HarvestHTML
   
   If iNFRulesCnt > 0 Then
   'if any of Not Found rules works terminate harvest
      For i = 1 To iNFRulesCnt
          If IsHTMLNF(sHTML, aNFRules(i)) Then
             If Len(sEntryMark) > 0 Then
                HarvestHTML = sEntryMark & NoHarvest
             Else
                HarvestHTML = NoHarvest
             End If
             Exit Function
          End If
      Next
   End If
   
   If iOKRulesCnt > 0 Then
      For i = 1 To iOKRulesCnt
          sHarvest = ParseHTML(sHTML, aOKRules(i))
          If Len(sHarvest) > 0 Then
             HarvestHTML = sHarvest
             Exit Function
          End If
      Next
   End If
End If
'if you came this far it is OK to return nothing...
'(no matter what will people say)
err_HarvestHTML:
HarvestHTML = ""
End Function


Private Function IsHTMLNF(ByVal sHTML As String, ByVal sNFRule As String) As Boolean
'returns True if "Not Found" rule is found in the sHTML
Dim sTerminator As String
On Error GoTo err_IsHTMLNF

IsHTMLNF = False
sTerminator = Right$(sNFRule, Len(sNFRule) - 3)
If InStr(1, sHTML, sTerminator) > 0 Then IsHTMLNF = True
err_IsHTMLNF:
End Function

Private Function ParseHTML(ByVal sHTML As String, ByVal sOKRule As String) As String
'parse sHTML based on instructions from the sOKRule and return
'found information
Dim sFrontTag As String
Dim sBackTag As String
Dim aFindOnceTag(1 To 100) As String
Dim lFOTCnt As Long
'strings used to specify start point for entry search
'have to be in order in which they appear in a HTML document
Dim aFindEntryTag(1 To 100) As String
Dim lFETCnt As Long
Dim sTerminator As String
Dim lStartPos As Long
Dim lEndPos As Long
Dim sPartOfRule As String
Dim PartOfRuleString As String
Dim l As Long, k As Long

Dim lCurrPos As Long
Dim lTmpPos As Long
Dim lTerminatorPos As Long

Dim aEntry(1000) As String
Dim lEntriesCnt As Long     'do not accept more than 200 entries
Dim sEntry As String
Dim sTmp As String
On Error GoTo exit_ParseHTML

sTmp = ""
lStartPos = 4
lEndPos = 0
lFOTCnt = 0
lFETCnt = 0
lEntriesCnt = 0
sTerminator = ""
sFrontTag = ""
sBackTag = ""
Do While lStartPos < Len(sOKRule)
   lEndPos = InStr(lStartPos, sOKRule, HTMLRuleSep)
   If lEndPos > lStartPos Then
      sPartOfRule = Mid$(sOKRule, lStartPos, lEndPos - lStartPos)
      PartOfRuleString = Right$(sPartOfRule, Len(sPartOfRule) - 2)
      Select Case Left(sPartOfRule, 1)
      Case "0"
           If lFOTCnt < 100 Then
              lFOTCnt = lFOTCnt + 1
              aFindOnceTag(lFOTCnt) = PartOfRuleString
           End If
      Case "1"
           If lFETCnt > 0 Then
              lFETCnt = lFETCnt + 1
              aFindEntryTag(lFETCnt) = PartOfRuleString
           End If
      Case "5"
           sFrontTag = PartOfRuleString
      Case "6"
           sBackTag = PartOfRuleString
      Case "7"
           sTerminator = PartOfRuleString
      Case Else
      End Select
      lStartPos = lEndPos + 1
   Else
      lStartPos = Len(sOKRule)
   End If
Loop
If Len(sFrontTag) = 0 Or Len(sBackTag) = 0 Then    'must have these two
   ParseHTML = ""
   Exit Function
End If
'go through tags that should be used only once
'stop only if first 0:XXX rule not found
lCurrPos = 1
If lFOTCnt > 0 Then
   For l = 1 To lFOTCnt
       lTmpPos = InStr(lCurrPos, sHTML, aFindOnceTag(l))
       If lTmpPos > 0 Then
          lCurrPos = lTmpPos
       Else     'first specified 0:xxx rule has to be found to proceede
          If l = 1 Then
             ParseHTML = ""
             Exit Function
          End If
       End If
   Next l
End If
'cut unneccessary part from the left
If lCurrPos > 1 Then
   sHTML = Right$(sHTML, Len(sHTML) - lCurrPos)
End If

lTerminatorPos = 0
If Len(sTerminator) > 0 Then
   lTerminatorPos = InStr(1, sHTML, sTerminator)
End If
If lTerminatorPos <= 0 Then
   lTerminatorPos = Len(sHTML)
End If

'finally do the search
lCurrPos = 1
Do While lCurrPos < lTerminatorPos
   If lFETCnt > 0 Then
      For l = 1 To lFETCnt  'have to find all of these
          lTmpPos = InStr(lCurrPos, sHTML, aFindEntryTag(l))
          If lTmpPos > 0 And lTmpPos < lTerminatorPos Then
             lCurrPos = lTmpPos + 1
          Else  'we are done
             Exit Do
          End If
      Next l
   End If
   lStartPos = InStr(lCurrPos, sHTML, sFrontTag)
   If lStartPos > 0 Then
      lEndPos = InStr(lStartPos, sHTML, sBackTag)
      If (lEndPos > 0) And (lEndPos < lTerminatorPos) Then
         'find last FrontTag in front of BackTag
         lTmpPos = lStartPos
         Do While lTmpPos > 0 And lTmpPos < lEndPos
            lStartPos = lTmpPos
            lTmpPos = InStr(lTmpPos + 1, sHTML, sFrontTag)
         Loop
         'extract what we were looking for
         lStartPos = lStartPos + Len(sFrontTag)
         sEntry = Trim$(Mid$(sHTML, lStartPos, lEndPos - lStartPos))
         If VerifyLen(MinEntryLen, MaxEntryLen, sEntry) Then
            If lEntriesCnt < HTMLMaxEntries Then
               lEntriesCnt = lEntriesCnt + 1
               If Len(sEntryMark) > 0 Then sEntry = sEntryMark & sEntry
               aEntry(lEntriesCnt) = sEntry
            Else
               Exit Do
            End If
         End If
         lCurrPos = lEndPos + 1
      Else
         Exit Do
      End If
   Else
      Exit Do
   End If
Loop

If lEntriesCnt > 0 Then
   For l = 1 To lEntriesCnt
       If l <> 1 Then aEntry(l) = "; " & aEntry(l)
       sTmp = sTmp & aEntry(l)
   Next l
End If

exit_ParseHTML:
ParseHTML = Trim(sTmp)
End Function

Private Function VerifyLen(ByVal minL As Integer, ByVal maxL As Integer, ByVal s As String) As Boolean
'returns True if length of the string s is in between minL, maxL
Dim sLen As Integer
VerifyLen = False
sLen = Len(s)
If sLen = 0 Then Exit Function
If minL = -1 Then minL = 0
If maxL = -1 Then maxL = 32767
If sLen >= minL And sLen <= maxL Then VerifyLen = True
End Function


Public Sub PrepareHTMLRule(ByVal sHTMLRule As String)
'prepare global variables from the sHTMLRule retrieved from the DB
Dim sPartOfRule As String
Dim iStartPos As Long
Dim iEndPos As Long
Dim iExpLenSepPos As Integer
On Error GoTo exit_PrepareHTMLRule
   
iNFRulesCnt = 0
iOKRulesCnt = 0
If Len(sHTMLRule) > 0 Then           'extract rules from the database
   iStartPos = 1
   iEndPos = 0
   MinEntryLen = -1
   MaxEntryLen = -1
   sEntryMark = ""
   Do While iStartPos < Len(sHTMLRule)
      iStartPos = InStr(iStartPos, sHTMLRule, "{")
      iEndPos = InStr(iStartPos + 1, sHTMLRule, "}")
      If iEndPos > iStartPos Then
         sPartOfRule = Mid$(sHTMLRule, iStartPos + 1, iEndPos - iStartPos - 1)
         iStartPos = iEndPos
         If Len(sPartOfRule) > 3 Then
            Select Case Left$(sPartOfRule, 3)
            Case HTMLNoMatchesTag
              iNFRulesCnt = iNFRulesCnt + 1
              aNFRules(iNFRulesCnt) = sPartOfRule
            Case HTMLFoundTag
              iOKRulesCnt = iOKRulesCnt + 1
              aOKRules(iOKRulesCnt) = sPartOfRule
            Case HTMLExpLenTag
              iExpLenSepPos = InStr(1, sPartOfRule, ";")
              If iExpLenSepPos > 0 Then
                 Select Case iExpLenSepPos
                 Case 4                   'no minimum
                   If IsNumeric(Right(sPartOfRule, Len(sPartOfRule) - 4)) Then
                      MaxEntryLen = CInt(Right(sPartOfRule, Len(sPartOfRule) - 4))
                   End If
                 Case Len(sPartOfRule)    'no maximum
                   If IsNumeric(Mid$(sPartOfRule, 4, Len(sPartOfRule) - 4)) Then
                      MinEntryLen = CInt(Mid$(sPartOfRule, 4, Len(sPartOfRule) - 4))
                   End If
                 Case Else                'min ad max specified
                   If IsNumeric(Mid$(sPartOfRule, 4, iExpLenSepPos - 4)) Then
                      MinEntryLen = CInt(Mid$(sPartOfRule, 4, iExpLenSepPos - 4))
                   End If
                   If IsNumeric(Right(sPartOfRule, Len(sPartOfRule) - iExpLenSepPos)) Then
                      MaxEntryLen = CInt(Right(sPartOfRule, Len(sPartOfRule) - iExpLenSepPos))
                   End If
                 End Select
              End If
            Case HTMLMarkTag
              sEntryMark = Right$(sPartOfRule, Len(sPartOfRule) - 3)
            End Select
         End If
      Else
         iStartPos = Len(sHTMLRule)
      End If
   Loop
End If
exit_PrepareHTMLRule:
End Sub

Private Function GetHTMLBlock(ByVal iInd As Integer, AID() As String, ByVal iCnt As Integer) As String
Dim i As Integer
Dim sLinkURL(1 To HTMLMaxEntries) As String
Dim sTmp As String
On Error GoTo err_GetHTMLBlock
sTmp = "<li><font face=arial,helvetica><b>" & a2lsURL(iInd, 2) & "</b><br>"
'iCnt is positive
For i = 1 To iCnt
  If InStr(1, AID(iInd, i), NoHarvest) > 0 Then
    sLinkURL(i) = AID(iInd, i)
  Else
    sLinkURL(i) = "<a href=" & a2lsURL(iInd, 3) _
    & GetQueryT("+++" & AID(iInd, i), a2lsURL(iInd, 4)) & ">" & AID(iInd, i) & "</a>"
  End If
Next i
For i = 1 To iCnt
    sTmp = sTmp & sLinkURL(i) & " "
Next i
sTmp = sTmp & "</font><br><br>"
GetHTMLBlock = sTmp
Exit Function

err_GetHTMLBlock:
sTmp = sTmp & "Could not create hyperlinks!"
End Function

Public Function SearchNet1(ByVal iURLInd, ByVal sIdentity) As Integer
Dim aIDs() As String
Dim IdsFound As Integer
Dim sURL As String
Dim sTmp
On Error GoTo exit_SearchNet1

SearchNet1 = ErrSearchNetOther
ReDim aIDs(1 To i2lsURLCnt, 1 To HTMLMaxEntries)
IdsFound = FindIdentities(iURLInd, aIDs(), sIdentity)
Select Case IdsFound
Case 0
     SearchNet1 = ErrSearchNetNotFound
Case 1
     If InStr(1, aIDs(iURLInd, 1), NoHarvest) Then
        SearchNet1 = ErrSearchNetNoHarvest
     Else
        sURL = a2lsURL(iURLInd, 3) & GetQueryT("+++" & aIDs(iURLInd, 1), a2lsURL(iURLInd, 4))
        If BrowseTheNet(sURL) Then SearchNet1 = ErrSearchNetSuccess
     End If
Case Else
     sTmp = GetHTMLStart()
     sTmp = sTmp & GetHTMLBlock(iURLInd, aIDs(), IdsFound)
     sTmp = sTmp & GetHTMLEnd()
     WriteHTMLFile sTmp
     If BrowseTheNet(sHTMLFile) Then SearchNet1 = ErrSearchNetSuccess
End Select

exit_SearchNet1:
End Function

Public Function SearchNetAll(ByVal sIdentity) As Integer
Dim i As Integer
Dim aIDs() As String
Dim TotalIdsFound As Long
Dim aIdsFound() As Long  'this is parallel array to aIDs - contains
                        'nubler of IDs for each database
Dim sTmp
Dim sURL As String
Dim NonEmptyIndex As Integer
On Error GoTo exit_SearchNetAll

SearchNetAll = ErrSearchNetOther
If i2lsURLCnt > 0 Then
   ReDim aIDs(1 To i2lsURLCnt, 1 To HTMLMaxEntries)
   ReDim aIdsFound(1 To i2lsURLCnt)
   TotalIdsFound = 0
   For i = 1 To i2lsURLCnt
       aIdsFound(i) = FindIdentities(i, aIDs(), sIdentity)
       'remember nonempty in case that only one Id is found
       If aIdsFound(i) > 0 Then
          If InStr(1, aIDs(i, 1), NoHarvest) > 0 Then
             If aIdsFound(i) > 1 Then
                TotalIdsFound = TotalIdsFound + aIdsFound(i) - 1
             End If
          Else
             TotalIdsFound = TotalIdsFound + aIdsFound(i)
             NonEmptyIndex = i
          End If
       End If
   Next i
   Select Case TotalIdsFound
   Case 0
        SearchNetAll = ErrSearchNetNotFound
   Case 1
        'go directly to that one
        sURL = a2lsURL(NonEmptyIndex, 3) & GetQueryT("+++" & aIDs(NonEmptyIndex, 1), a2lsURL(NonEmptyIndex, 4))
        If BrowseTheNet(sURL) Then SearchNetAll = ErrSearchNetSuccess
   Case Else
        sTmp = GetHTMLStart()
        For i = 1 To i2lsURLCnt
            If aIdsFound(i) > 0 Then
               sTmp = sTmp & GetHTMLBlock(i, aIDs(), aIdsFound(i))
            Else
               sTmp = sTmp & GetHTMLEmptyBlock(i)
            End If
        Next i
        sTmp = sTmp & GetHTMLEnd()
        WriteHTMLFile sTmp
        If BrowseTheNet(sHTMLFile) Then SearchNetAll = ErrSearchNetSuccess
   End Select
End If

exit_SearchNetAll:
End Function

Private Function GetHTMLStart() As String
Dim sTmp As String
sTmp = "<html><body><table cellpadding=4 width=100%><tr><td bgcolor="
sTmp = sTmp & "cc3333><font color=cccccc face=arial,helvetica size=+1><b>2D Gel Internet "
sTmp = sTmp & " Connection </b></font></td></tr></table><p><ul>"
GetHTMLStart = sTmp
End Function

Private Function GetHTMLEnd() As String
Dim sTmp As String
sTmp = "<table width=100%><tr><td width=75%><hr></td><td width=15%>" _
  & "<font face=arial,helvetica size=-2><center>Gels are for ever." _
  & "</center></font></td><td width=10%><hr></td></tr></table>"
'GetHTMLEnd = "</ul><hr width=100%></body></html>"
GetHTMLEnd = "</ul>" & sTmp & "</body></html>"
End Function

Private Function GetHTMLEmptyBlock(ByVal iInd As Integer) As String
Dim sTmp As String
sTmp = "<li><font face=arial,helvetica><b>" & a2lsURL(iInd, 2) & "</b><br>"
sTmp = sTmp & NoSearchDone & "<br><br>"
GetHTMLEmptyBlock = sTmp
End Function

Private Function FindIdentities(ByVal iR As Integer, aIDs() As String, ByVal sID As String) As Integer
Dim sDBMark As String
Dim iCnt As Integer
Dim iStartPos As Integer
Dim iEndPos As Integer
Dim i As Integer
Dim sCh As String
Dim sCurrItem As String
Dim bDone As Boolean

On Error GoTo exit_FindIdentities
FindIdentities = 0

'pick mark of the database
sDBMark = Trim$(Right$(a2lsURL(iR, 1), 3))
If Len(sDBMark) <> 3 Then GoTo exit_FindIdentities

iCnt = 0
iStartPos = 1
iEndPos = 0
Do While iStartPos <= Len(sID) And iCnt < HTMLMaxEntries
   iStartPos = InStr(iStartPos, sID, sDBMark)
   If iStartPos > 0 Then
'looking for glARG_SEP
      iEndPos = iStartPos + 3
      bDone = False
      Do While (iEndPos <= Len(sID)) And (Not bDone)
         sCh = Mid$(sID, iEndPos, 1)
         If sCh = glARG_SEP Then
            bDone = True
         Else
            iEndPos = iEndPos + 1
         End If
      Loop
      If iEndPos > iStartPos + 3 Then   'we have something
         sCurrItem = Trim$(Mid$(sID, iStartPos + 3, iEndPos - iStartPos - 3))
         If Len(sCurrItem) > 0 Then
            iCnt = iCnt + 1
            aIDs(iR, iCnt) = sCurrItem
         End If
      End If
      iStartPos = iEndPos + 1
   Else
      iStartPos = Len(sID) + 1
   End If
Loop
exit_FindIdentities:
FindIdentities = iCnt
End Function

Private Function WriteHTMLFile(ByVal sHTML) As Boolean
Dim FileNum As Integer
On Error GoTo exit_WriteHTMLFile

WriteHTMLFile = False
FileNum = FreeFile
Open sHTMLFile For Output As #FileNum
Print #FileNum, sHTML
Close #FileNum
WriteHTMLFile = True
exit_WriteHTMLFile:
End Function

