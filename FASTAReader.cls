VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FastaReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'class implementing enumeration of ORF records in FASTA file
'Kyle Littlefield and Nikola Tolic
'created: somewhere last week
'last modified: 08/12/2002
'---------------------------------------------------------------------------------
'Properties: FASTAFile              - name of FASTA file
'            RecordStartCharacter   - begining of FASTA record(default is ">")
'            RecordCount            - number of FASTA records
'            LastRecord             - index of last record something was read from
'Methods:    EnumerateFASTARecords  - goes through FASTA file and enumerates all
'                                     records with byte positions and lengths
'            GetFASTARecordDesc     - retrieves description from FASTA file
'            GetFASTARecordSeq      - retrieves sequence from FASTA file
'---------------------------------------------------------------------------------

Option Explicit
Private DS As Integer   '> - description start
Const LF = 10           'line feed
Const CR = 13           'carriage return

Private mFile As String
Private ORFCnt As Long
Private mLastRecord As Long

'''Private ORFID() As Long
Private ORFDescStart() As Long
Private ORFDescLen() As Long

Private ORFSeqStart() As Long
Private ORFSeqLen() As Long

Private Sub Class_Initialize()
    ORFCnt = 0
    Me.RecordStartCharacter = ">"
End Sub

Public Property Get FASTAFile() As String
    FASTAFile = mFile
End Property

Public Property Let FASTAFile(ByVal FileName As String)
    mFile = FileName
End Property

Public Function EnumerateFASTARecords() As Boolean
    Dim i As Long
    Dim bCurr As Byte
    Dim CurrLen As Long
    Dim hfile As Integer
    Dim ReadingSeq As Boolean
    On Error GoTo err_EnumerateFASTARecords
    ReDim ORFDescStart(1 To 10000)
    ReDim ORFDescLen(1 To 10000)
    ReDim ORFSeqStart(1 To 10000)
    ReDim ORFSeqLen(1 To 10000)
    hfile = FreeFile
    Open mFile For Binary Access Read As hfile
    ORFCnt = 0
    ReadingSeq = False
    i = 1
    Do Until EOF(hfile)
       Get #hfile, i, bCurr
       If bCurr = DS Then                                    'new FASTA record
          If i > 1 Then ORFSeqLen(ORFCnt) = CurrLen          'close previous record
          CurrLen = 1
          ORFCnt = ORFCnt + 1
          ORFDescStart(ORFCnt) = i                           'mark start of new record
          ReadingSeq = False
       Else
          If bCurr = LF Then                    'description ends on first LF
             If ReadingSeq Then
                CurrLen = CurrLen + 1
             Else                                            'start of sequence
                ORFDescLen(ORFCnt) = CurrLen                 'mark end of description
                ORFSeqStart(ORFCnt) = i + 1                  'mark start of sequence
                ReadingSeq = True
                CurrLen = 0
             End If
          Else
             CurrLen = CurrLen + 1
          End If
       End If
       i = i + 1
    Loop
    ORFSeqLen(ORFCnt) = CurrLen                              'end of las sequence
    If ORFCnt > 0 Then
       If ORFCnt < UBound(ORFSeqStart) Then
          ReDim Preserve ORFDescStart(1 To ORFCnt)
          ReDim Preserve ORFDescLen(1 To ORFCnt)
          ReDim Preserve ORFSeqStart(1 To ORFCnt)
          ReDim Preserve ORFSeqLen(1 To ORFCnt)
       End If
    Else
       Erase ORFDescStart
       Erase ORFDescLen
       Erase ORFSeqStart
       Erase ORFSeqLen
    End If
    Close hfile
    EnumerateFASTARecords = True
    Exit Function
    
err_EnumerateFASTARecords:
    MsgBox "Error reading file.", vbOKOnly
End Function

Public Property Get RecordCount() As Long
    RecordCount = ORFCnt
End Property

Public Property Get LastRecord() As Long
    LastRecord = mLastRecord
End Property

Public Function GetFASTARecordSeq(ByVal Index As Long) As String
    Dim hfile As Integer
    Dim bORF() As Byte
    On Error GoTo err_GetFASTARecordSeq
    ReDim bORF(ORFSeqLen(Index) - 1)
    mLastRecord = Index
    hfile = FreeFile
    Open mFile For Binary Access Read As hfile
         Get #hfile, ORFSeqStart(Index), bORF
    Close hfile
    GetFASTARecordSeq = CleanFASTARecord(bORF)
    Exit Function
    
err_GetFASTARecordSeq:
    GetFASTARecordSeq = "Error retrieving sequence with index: " & Index
End Function

Public Function GetFASTARecordDesc(ByVal Index As Long) As String
    Dim hfile As Integer
    Dim bORF() As Byte
    On Error GoTo err_GetFASTARecordDesc
    ReDim bORF(ORFDescLen(Index) - 1)
    mLastRecord = Index
    hfile = FreeFile
    Open mFile For Binary Access Read As hfile
        Get #hfile, ORFDescStart(Index), bORF
    Close hfile
    GetFASTARecordDesc = StrConv(bORF, vbUnicode)
    Exit Function
    
err_GetFASTARecordDesc:
    GetFASTARecordDesc = "Error retrieving sequence description with index: " & Index
End Function

Private Function CleanFASTARecord(ORF() As Byte) As String
    Dim i As Variant
    Dim bCleanedORF() As Byte
    Dim CleanedOrfIndex As Long
    
    ReDim bCleanedORF(1 To ORFSeqLen(mLastRecord))
    CleanedOrfIndex = 1
    For Each i In ORF
        If i > 32 Then
            bCleanedORF(CleanedOrfIndex) = i
            CleanedOrfIndex = CleanedOrfIndex + 1
        End If
    Next i
    ReDim Preserve bCleanedORF(1 To (CleanedOrfIndex - 1))
    CleanFASTARecord = StrConv(bCleanedORF, vbUnicode)
End Function

Public Property Let RecordStartCharacter(Char As String)
    Dim bchar() As Byte
    bchar = StrConv(Char, vbFromUnicode)
    DS = CInt(bchar(0))
End Property

Public Property Get RecordStartCharacter() As String
    RecordStartCharacter = StrConv(DS, vbUnicode)
End Property
