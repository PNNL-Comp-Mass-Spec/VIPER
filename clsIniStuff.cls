VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniStuff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------------------------------------
' clsIniStuff: easy wrapper for inifiles
'
' Usage: set the filename, and call the Write/Read functions
'        These functions are booleans, if they return false, an error has occured
'        Use LastErrorNumber for the errornumber, and LastErrorDescription to get
'        the last errors (note that these are API errors, not VB errors)
'        You can use this number and description to show the user why it didn't work
'
' DAC Mods:
'        Corrected error in Enum_SectionNames that caused return of blank item
'        at end of list; added comments to Enum_SectionNames heading; modified
'        return value to ensure VB would correctly interpret True and False;
'        modified ReadValue to accept a default value; modified ReadSection to account
'        for API call return differences between Win98 and Win2000
'-----------------------------------------------------------------------------------

' API declares
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

' last error
' format the errornumber to a description windooz uses
' note that you must use both FORMAT_MESSAGE constants (that's why I created the FORMAT_PARAM)
' if you don't use FORMAT_MESSAGE_IGNORE_INSERTS, your vb can crash if there's a %s in the errordesription
' (try it out with errornumber 129 if you don't believe me LOL)
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_PARAM = FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS

' private variables
Private m_FileName As String             ' ini filename
Private m_LastErrorNumber As Long        ' last errornumber if any occured
Private m_LastErrorDescription As String ' error description of the last occured error

' Default stuff....
Private Sub Class_Initialize()
    m_FileName = ""
    m_LastErrorDescription = ""
    m_LastErrorNumber = 0
End Sub

Private Sub Class_Terminate()
    m_FileName = ""
    m_LastErrorDescription = ""
    m_LastErrorNumber = 0
End Sub

' Ini filename
Public Property Get FileName() As String
    FileName = m_FileName
End Property

Public Property Let FileName(New_FileName As String)
    m_FileName = New_FileName
End Property

' Last Error Description
Public Property Get LastErrorDescription() As String
    LastErrorDescription = m_LastErrorDescription
End Property

'-----------------------------------------------------------------------------------
' ReadValue: reads the specified value from the inifile
' In:
'   Section name, Key name, and a string which will contain the Value on return
' Returns:
'   true if success, false if fails (get LastErrorDescription for more info.....)
'   NOTE that psValue will contain the stringvalue read from the ini on return!
'-----------------------------------------------------------------------------------
Public Function ReadValue(psSection As String, psKey As String, psDefault As String, psValue As String) As Boolean
    
    ' get out if we don't have a filename ....
    If Not HaveFileName Then ReadValue = False: Exit Function
    Dim lRet As Long, sRetValue As String * 255
    lRet = GetPrivateProfileString(psSection, psKey, psDefault, sRetValue, Len(sRetValue), m_FileName)
    If lRet = 0 Then
        GetTheLastError
    Else
        psValue = Left(sRetValue, InStr(1, sRetValue, vbNullChar) - 1)
    End If
'    ReadValue = CBool(lRet)
    ReadValue = SetReturnValue(lRet)
    
End Function

'-----------------------------------------------------------------------------------
' ReadSection: reads specifed section at once
' In:
'   section name, key array, value array
' returns:
'   true if success, false if fails (get LastErrorDescription for more info.....)
'   NOTE that both arrays must not be fixed dimension, else the redim will fail!!!
'   these 2 arrays will be redimmed as much as needed to contain all the keynames and values!
'-----------------------------------------------------------------------------------
Public Function ReadSection(psSection As String, psKeys() As String, psValues() As String) As Boolean
    
    If Not HaveFileName Then ReadSection = False: Exit Function
    ' return value
    Dim lRet As Long, sReturned As String * 32767 ' max chars allowed in Win95
    lRet = GetPrivateProfileSection(psSection, sReturned, Len(sReturned), m_FileName)
    If lRet = 0 Then
        GetTheLastError
    Else
        ReDim psKeys(0)
        ReDim psValues(0)
        Dim iNull As Integer, iStart As Integer, i As Integer, S As String
        i = 0: iStart = 1
        iNull = InStr(iStart, sReturned, vbNullChar)
        Do While iNull
            S = Mid(sReturned, iStart, iNull - iStart)
            If Left(S, 1) <> ";" Then   'Win98 includes comments in returned string, so ignore them
              ReDim Preserve psKeys(i)
              ReDim Preserve psValues(i)
              psKeys(i) = Left(S, InStr(1, S, "=") - 1)
              psValues(i) = Right(S, Len(S) - InStr(1, S, "="))
              i = i + 1
            End If
            iStart = iNull + 1
            iNull = InStr(iStart, sReturned, vbNullChar)
            ' lRet contains the numbers of chars copied to the buffer, so if iNull > lRet then we have it all...
            If iNull > lRet Then iNull = 0
        Loop
    End If
'    ReadSection = CBool(lRet)
    ReadSection = SetReturnValue(lRet)
    
End Function

'-----------------------------------------------------------------------------------
' WriteValue: writes the specified value to the ini file pointed to in the Filename property
' In:
'   Section name, Key name, and Value
' Returns:
'   true if success, false if fails (get LastErrorDescription for more info.....)
'-----------------------------------------------------------------------------------
Public Function WriteValue(psSection As String, psKey As String, psValue As String) As Boolean
    ' get out if we don't have a filename ....
    If Not HaveFileName Then WriteValue = False: Exit Function
    ' Return value
    Dim lRet As Long
    ' Try to write the value
    lRet = WritePrivateProfileString(psSection, psKey, psValue, m_FileName)
    ' if lRet = 0 an error has occured
    If lRet = 0 Then GetTheLastError
'    WriteValue = CBool(lRet)
    WriteValue = SetReturnValue(lRet)
    
End Function

'-----------------------------------------------------------------------------------
' WriteSection: Write a section at once.
' In:
'   Section name, array of keys, array of values
'   Both array must be of the same size, else nothing will be written
' Returns:
'   true if success, false if fails (get LastErrorDescription for more info.....)
'-----------------------------------------------------------------------------------
Public Function WriteSection(ByRef psSection As String, ByRef psKeys() As String, psValues() As String, ByVal intItemCount As Integer) As Boolean
    ' If intItemcount is <= 0, then writes out the entire psKeys() and psValues() array
    ' Otherwise, writes out the number of items specified by intItemCount
    
    Dim lngIndexStart As Long
    Dim lngIndexEnd As Long
    
    ' Abort if we don't have a filename ....
    If Not HaveFileName Then
        WriteSection = False
        Exit Function
    End If
    
    If intItemCount <= 0 Then
        If UBound(psKeys) <> UBound(psValues) Then
            WriteSection = False
            m_LastErrorDescription = "Key array and Value array must be of the same size"
            Exit Function
        End If
        lngIndexStart = LBound(psKeys)
        lngIndexEnd = UBound(psKeys)
    Else
        If UBound(psKeys) < intItemCount - 1 Or UBound(psValues) < intItemCount - 1 Then
            WriteSection = False
            m_LastErrorDescription = "Key array or Value array is not of length " & intItemCount
            Exit Function
        End If
        lngIndexStart = 0
        lngIndexEnd = intItemCount - 1
    End If
    
    ' return value
    Dim lRet As Long
    
    ' tempstring which will contain the value to write on this format: key=value+vbNullChar+key=value etc...
    Dim S As String, L As Long
    
    ' format the string to write
    For L = lngIndexStart To lngIndexEnd
        S = S & psKeys(L) & "=" & psValues(L) & vbNullChar
    Next
    
    ' write section
    lRet = WritePrivateProfileSection(psSection, S, m_FileName)
    If lRet = 0 Then GetTheLastError
    
'    WriteSection = CBool(lRet)
    WriteSection = SetReturnValue(lRet)
    
End Function

'-----------------------------------------------------------------------------------
' Enum_SectionNames: Get list of section names.
' In:
'   Value array
' Returns:
'   true if success, false if fails (get LastErrorDescription for more info.....)
'-----------------------------------------------------------------------------------
Public Function Enum_SectionNames(psSectionNames() As String) As Boolean
    ' get out if we don't have a filename ....
    If Not HaveFileName Then Enum_SectionNames = False: Exit Function
    ' return value
    Dim lRet As Long, sReturned As String * 32767 ' max chars allowed in Win95
    lRet = GetPrivateProfileSectionNames(sReturned, Len(sReturned), m_FileName)
    If lRet = 0 Then
        GetTheLastError
    Else
        Dim iNull As Integer, iStart As Integer, i As Integer
        i = 0: iStart = 1
        iNull = InStr(iStart, sReturned, vbNullChar)
        Do While iNull
            ReDim Preserve psSectionNames(i)
            psSectionNames(i) = Mid(sReturned, iStart, iNull - iStart)
            iStart = iNull + 1
            iNull = InStr(iStart, sReturned, vbNullChar)
            ' lRet contains the numbers of chars copied to the buffer, so if iNull > lRet then we have it all...
            If iNull > lRet Then iNull = 0
            i = i + 1
        Loop
    End If
'    Enum_SectionNames = CBool(lRet)
    Enum_SectionNames = SetReturnValue(lRet)
    
End Function


'-----------------------------------------------------------------------------------
' PRIVATE FUNCTIONS
'-----------------------------------------------------------------------------------
' check if we have a filename...
Private Function HaveFileName() As Boolean
    If m_FileName = "" Then
        m_LastErrorDescription = "No filename specified"
        HaveFileName = False
    Else
        HaveFileName = True
    End If
End Function

' set m_LastErrorDescription...
Private Sub GetTheLastError()
    m_LastErrorNumber = GetLastError
    If m_LastErrorNumber <> 0 Then ' call succesful, so the specified string was not found... tell it...
        ' and format it...
        Dim strError As String * 255
        Call FormatMessage(FORMAT_PARAM, 0, m_LastErrorNumber, 0, strError, Len(strError), 0)
        m_LastErrorDescription = Left(strError, InStr(1, strError, vbNullChar) - 1)
    Else
        m_LastErrorDescription = "An unknown error has occured"
    End If
End Sub

Private Function SetReturnValue(ByVal MyValue) As Boolean

  'Ensures return value is consistent with VB boolean logic
  If MyValue = 0 Then
    SetReturnValue = False
  Else
    SetReturnValue = True
  End If
  
End Function

