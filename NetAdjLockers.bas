Attribute VB_Name = "NetAdjLockers"
Option Explicit

' Internal Standards are either defined in a .Ini file or downloaded
'  from the database when the AMT's are downloaded.

Public Enum issmInternalStandardSearchModeConstants
    issmFindOnlyMassTags = 0
    issmFindWithMassTags = 1
    issmFindOnlyInternalStandards = 2
End Enum

' The following holds the Internal Standards; it is not saved with the .Gel file
Public Type udtInternalStandardEntryType
    SeqID As String         ' This is a string to allow text-based ID's; however, Internal Standards downloaded from the database will have numeric values
    Description As String
    PeptideSequence As String
    MonoisotopicMass As Double      ' Expected Monoisotopic Mass
    NET As Double                   ' Expected NET
    ChargeMinimum As Integer        ' Minimum expected charge; set to 0 to allow any charge
    ChargeMaximum As Integer
    ChargeMostAbundant As Integer
End Type

Public Type udtInternalStandardsType
    Count As Integer
    InternalStandards() As udtInternalStandardEntryType            ' 0-based array
    StandardsAreFromDB As Boolean
End Type

Public UMCInternalStandards As udtInternalStandardsType

Public Function LoadInternalStandards(ByRef frmCallingForm As VB.Form, ByVal lngGelIndex As Long, ByRef udtFilteringOptions As udtMTFilteringOptionsType) As Boolean
    '---------------------------------------------------------------------------
    ' Retrieves list of Internal Stanards from the database by calling the
    '  stored procedure specified by glbPreferencesExpanded.MTSConnectionInfo.spGetInternalStandards
    '  (typically GetInternalStandards)
    '---------------------------------------------------------------------------
    Dim cnNew As New ADODB.Connection
    Dim sCommand As String
    Dim rsInternalStandards As New ADODB.Recordset
    Dim cmdGetInternalStandards As New ADODB.Command
    
    Dim prmJob As ADODB.Parameter
    Dim prmInternalStdExplicit As ADODB.Parameter
    
    Dim strCaptionSaved As String
    Dim blnSuccess As Boolean
    
    ' Clear the current Internal Standards
    With UMCInternalStandards
        .Count = 0
        ReDim .InternalStandards(0)
    End With
    
    If GelAnalysis(lngGelIndex) Is Nothing Then
        Debug.Print "GelAnalysis object not defined; unable to continue"
        Debug.Assert False
        LoadInternalStandards = False
        Exit Function
    End If
        
    ' Save the form's caption
    strCaptionSaved = frmCallingForm.Caption
    
On Error Resume Next
    sCommand = glbPreferencesExpanded.MTSConnectionInfo.spGetInternalStandards
    If Len(sCommand) <= 0 Then
        Debug.Print "Invalid or missing spGetInternalStandards stored procedure name"
        sCommand = "GetInternalStandards"
    End If
    
On Error GoTo LoadInternalStandardPeptidesErrorHandler
    
    Screen.MousePointer = vbHourglass
    
    If Not EstablishConnection(cnNew, GelAnalysis(lngGelIndex).MTDB.cn.ConnectionString, False) Then
        Debug.Assert False
        LoadInternalStandards = False
        Exit Function
    End If
    
    ' Initialize the Stored Procedure
    Set cmdGetInternalStandards.ActiveConnection = cnNew
    With cmdGetInternalStandards
        .CommandText = sCommand
        .CommandType = adCmdStoredProc
        .CommandTimeout = glbPreferencesExpanded.AutoAnalysisOptions.DBConnectionTimeoutSeconds
    End With
    
    Set prmJob = cmdGetInternalStandards.CreateParameter("Job", adInteger, adParamInput, , udtFilteringOptions.CurrentJob)
    cmdGetInternalStandards.Parameters.Append prmJob
    
    Set prmInternalStdExplicit = cmdGetInternalStandards.CreateParameter("InternalStdExplicit", adVarChar, adParamInput, 255, udtFilteringOptions.InternalStandardExplicit)
    cmdGetInternalStandards.Parameters.Append prmInternalStdExplicit
    
    'procedure returns error number or 0 if OK
    Set rsInternalStandards = cmdGetInternalStandards.Execute
    
    ' Note: Expected fields are:
    ' 0 = GANET_Locker_ID
    ' 1 = Description,
    ' 2 = Peptide,
    ' 3 = Monoisotopic_Mass,
    ' 4 = NET
    ' 5 = Charge_Minimum,
    ' 6 = Charge_Maximum,
    ' 7 = Charge_Highest_Abu
    
    With UMCInternalStandards
        Do Until rsInternalStandards.EOF
            ReDim Preserve .InternalStandards(.Count)
            With .InternalStandards(.Count)
                .SeqID = rsInternalStandards.Fields(0)                          ' This should be a integer, but .SeqID is a string for flexibility reasons
                .Description = FixNull(rsInternalStandards.Fields(1))
                .PeptideSequence = FixNull(rsInternalStandards.Fields(2))
                .MonoisotopicMass = rsInternalStandards.Fields(3)               ' This should never be null
                .NET = rsInternalStandards.Fields(4)                            ' This should never be null
                .ChargeMinimum = rsInternalStandards.Fields(5)                  ' This should never be null
                .ChargeMaximum = rsInternalStandards.Fields(6)                  ' This should never be null
                .ChargeMostAbundant = rsInternalStandards.Fields(7)             ' This should never be null
            End With
            .Count = .Count + 1
           
           rsInternalStandards.MoveNext
           frmCallingForm.Caption = "Loading Internal Standards: " & .Count
        Loop
    End With
    
    rsInternalStandards.Close
    
    ' Update CurrMTFilteringOptions
    CurrMTFilteringOptions.InternalStandardExplicit = udtFilteringOptions.InternalStandardExplicit
    
    If UMCInternalStandards.Count > 0 Then
        UMCInternalStandards.StandardsAreFromDB = True
    Else
        PopulateDefaultInternalStds UMCInternalStandards
    End If
    blnSuccess = True
    
    'clean things and exit
LoadInternalStandardPeptidesCleanup:
    On Error Resume Next
    Set cmdGetInternalStandards.ActiveConnection = Nothing
    cnNew.Close
    
    ' Restore the caption on the calling form
    frmCallingForm.Caption = strCaptionSaved
    Screen.MousePointer = vbDefault

    LoadInternalStandards = blnSuccess
Exit Function

LoadInternalStandardPeptidesErrorHandler:
    Debug.Print "Error in LoadInternalStandards"
    Debug.Assert False
    Select Case Err.Number
    Case 13, 94                  'Type Mismatch or Invalid Use of Null
        Resume Next              'just ignore it
    Case -2147217900
        ' Stored procedure not found in database
        LogErrors Err.Number, "LoadInternalStandards"
    Case Else
        LogErrors Err.Number, "LoadInternalStandards"
    End Select
    
    blnSuccess = False
    Resume LoadInternalStandardPeptidesCleanup

End Function

