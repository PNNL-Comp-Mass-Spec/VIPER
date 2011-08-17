VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InitFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------
'initialization file class
'initialization file contains set of blocks of type

    '[Section]
    '|connection string| (if Section contains word Database)
    'name_1=value_1
    '...
    'name_n=value_n
    '[End]

'-----------------------------------------------------------
'created: 06/13/2001 nt
'last modified: 08/16/2001 nt
'-----------------------------------------------------------

Option Explicit

Const faxa_MAX_FILE_SIZE = 100000   'maximum size of editable file

Dim fs As New FileSystemObject

Public Function WriteSection(ByVal FName As String, _
                             ByVal SecName As String, _
                             ByRef NewValue As String) As Long
'--------------------------------------------------------------
'rewrites section in the file; if section does not exists; this
'procedure will write it at the end of file;
'returns 0 on success; error number on any error
'last modified:06/11/2001 nt
'--------------------------------------------------------------
Dim sFile As String         'whole Init file (not too large?)
Dim SecStartPos As Long
Dim SecEndPos As Long
Dim tsIn As TextStream
Dim tsOut As TextStream
On Error GoTo err_WriteSection

Set tsIn = fs.OpenTextFile(FName, ForReading, False)
sFile = tsIn.ReadAll
tsIn.Close
SecStartPos = InStr(1, sFile, SecName)
With MyGl
   If SecStartPos > 0 Then
      SecEndPos = InStr(SecStartPos + 1, sFile, .INIT_SectionEnd)
      If SecEndPos > 0 Then
         sFile = Left$(sFile, SecStartPos - 1) & SecName _
           & vbCrLf & NewValue & vbCrLf & .INIT_SectionEnd & vbCrLf _
           & Right$(sFile, Len(sFile) - SecEndPos - Len(.INIT_SectionEnd) + 1)
      Else     'no section end; put new value and correct illegality
         sFile = Left$(sFile, SecStartPos - 1) & SecName _
           & vbCrLf & NewValue & vbCrLf & .INIT_SectionEnd
      End If
   Else        'no such section; put it at the end
      sFile = sFile & vbCrLf & SecName & vbCrLf _
           & NewValue & vbCrLf & .INIT_SectionEnd
   End If
End With

Set tsOut = fs.CreateTextFile(FName, True)
tsOut.Write sFile
tsOut.Close
WriteSection = 0
Exit Function

err_WriteSection:
WriteSection = Err.Number
End Function


Public Function GetSection(ByVal FName As String, _
                           ByVal Section As String, _
                           ByRef Sec() As String) As Long
'------------------------------------------------------------
'retrieves first Section of file as array of strings(lines in
'file) and returns number of it; returns -1 if file not found
'or -3 on other error; -2 if section not found
'last modified:06/11/2001 nt
'------------------------------------------------------------
Dim SecLnCnt As Long
Dim BOS As Boolean
Dim EOS As Boolean
Dim sLine As String
Dim ts As TextStream
On Error GoTo err_GetSection

Set ts = fs.OpenTextFile(FName, ForReading, False)

ReDim Sec(100)      'it is not very likely that section will need more space
Do Until BOS
   If ts.AtEndOfStream Then
      BOS = True
      EOS = True
      GetSection = faxa_INIT_FILE_SECTION_NOT_FOUND
   Else
      sLine = ts.ReadLine
      If InStr(1, sLine, Section) > 0 Then BOS = True
   End If
Loop
Do Until EOS
   If ts.AtEndOfStream Then
      EOS = True
   Else
      sLine = ts.ReadLine
      sLine = Trim$(sLine)
      If Len(sLine) > 0 Then
         If InStr(1, sLine, MyGl.INIT_SectionEnd) > 0 Then
            EOS = True
         Else
            If Left$(sLine, 2) <> MyGl.INIT_Comment Then
               SecLnCnt = SecLnCnt + 1
               Sec(SecLnCnt - 1) = sLine     'this might trigger error #9
            End If
         End If
      End If
   End If
Loop
ts.Close
If SecLnCnt > 0 Then
   ReDim Preserve Sec(SecLnCnt - 1)
Else
   Erase Sec
End If
GetSection = SecLnCnt
Exit Function

err_GetSection:
Select Case Err.Number
Case 9      'subscript out of range; add more space and resume
    ReDim Preserve Sec(SecLnCnt + 100)
    Resume
Case 53     'Ini file not found
    GetSection = faxa_INIT_FILE_NOT_FOUND
Case Else
    GetSection = faxa_INIT_FILE_ANY_ERROR
End Select
End Function


Public Function GetResolvedSection(ByVal FName As String, _
                                   ByVal SectionName As String, _
                                   ByRef SecNames() As String, _
                                   ByRef SecValues() As String) As Long
'----------------------------------------------------------------------
'retrieves section of file as parallel array of strings; returns number
'of it/-1 on error; if no "=" is found value is considered to be "None"
'last modified:04/12/2001 nt
'----------------------------------------------------------------------
Dim Section() As String
Dim SecLnCnt As Long
Dim ValPos As Long
Dim i As Long
On Error GoTo err_GetResolvedSection

SecLnCnt = GetSection(FName, SectionName, Section())
If SecLnCnt > 0 Then
   SecLnCnt = CleanComments(Section())
   If SecLnCnt > 0 Then
      ReDim SecNames(SecLnCnt - 1)
      ReDim SecValues(SecLnCnt - 1)
      For i = 0 To SecLnCnt - 1
          ValPos = InStr(1, Section(i), MyGl.INIT_Value)
          If ValPos > 0 Then
             SecNames(i) = Trim(Left$(Section(i), ValPos - 1))
             SecValues(i) = Trim$(Right$(Section(i), Len(Section(i)) - ValPos))
          Else      'everything is name; value is "None"
             SecNames(i) = Trim(Section(i))
             SecValues(i) = MyGl.INIT_None
          End If
      Next i
   End If
End If
GetResolvedSection = SecLnCnt
Exit Function

err_GetResolvedSection:
GetResolvedSection = -1
End Function


Private Sub Class_Terminate()
Set fs = Nothing
End Sub


Public Function GetSectionNames(ByVal FName As String, _
                                ByRef SecNames() As String) As Long
'------------------------------------------------------------------
'retrieves names of all sections in FName; returns number of it
'------------------------------------------------------------------
Dim SecCnt As Long
Dim sLine As String
Dim ts As TextStream
On Error GoTo err_GetSectionNames

Set ts = fs.OpenTextFile(FName, ForReading, False)
ReDim SecNames(100)   'it is not very likely there will be more sections
SecCnt = 0
Do Until ts.AtEndOfStream
   sLine = Trim$(ts.ReadLine)
   If Left$(sLine, 1) = "[" Then
      If Right$(sLine, 1) = "]" Then
         If sLine <> MyGl.INIT_SectionEnd Then
            SecCnt = SecCnt + 1
            SecNames(SecCnt - 1) = sLine
         End If
      End If
   End If
Loop
ts.Close
If SecCnt > 0 Then
   ReDim Preserve SecNames(SecCnt - 1)
Else
   Erase SecNames
End If
GetSectionNames = SecCnt
Exit Function

err_GetSectionNames:
Select Case Err.Number
Case 9      'subscript out of range; add more space and resume
    ReDim Preserve SecNames(SecCnt + 100)
    Resume
Case 53     'Ini file not found
    GetSectionNames = faxa_INIT_FILE_NOT_FOUND
Case Else
    GetSectionNames = faxa_INIT_FILE_ANY_ERROR
End Select
End Function
