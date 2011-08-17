VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FTICRAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'-----------------------------------------------------------------
'FTICR Analysis class
'-----------------------------------------------------------------
'created: 06/12/2001 nt
'last modified: 05/23/2002 nt
'-----------------------------------------------------------------
Const t_FAD = "FTICR Analysis Description Table"

Public ProcessingType As Long       'mass locking or mass matching

'fields coming from T_FTICR_Analysis_Description table; after
'selection is made this contains data from Reference Job record
Const FLD_Job = "Job"
Const FLD_Dataset = "Dataset"
Const FLD_Dataset_ID = "Dataset_ID"
Const FLD_Experiment = "Experiment"
Const FLD_Organism = "Organism"
Const FLD_Instrument_Class = "Instrument_Class"
Const FLD_Analysis_Tool = "Analysis_Tool"
Const FLD_Parameter_File_Name = "Parameter_File_Name"
Const FLD_Settings_File_Name = "Settings_File_Name"
Const FLD_Organism_DB_Name = "Organism_DB_Name"
Const FLD_Vol_Client = "Vol_Client"
Const FLD_Vol_Server = "Vol_Server"
Const FLD_Storage_Path = "Storage_Path"
Const FLD_Dataset_Folder = "Dataset_Folder"
Const FLD_Results_Folder = "Results_Folder"
Const FLD_Created = "Created"
Const FLD_State = "State"
Const FLD_NET_TICFit = "NET_TICFit"
Const FLD_NET_Slope = "NET_Slope"
Const FLD_NET_Intercept = "NET_Intercept"
Const FLD_Total_Scans = "Total_Scans"
Const FLD_Duration = "Duration"
Const FLD_GANET_Fit = "GANET_Fit"
Const FLD_GANET_Slope = "GANET_Slope"
Const FLD_GANET_Intercept = "GANET_Intercept"

Public Job As Long
Public Dataset As String
Public Dataset_ID As Long
Public Experiment As String
Public Organism As String
Public Instrument_Class As String
Public Analysis_Tool As String
Public Parameter_File_Name As String
Public Settings_File_Name As String
Public Organism_DB_Name As String
Public Vol_Client As String
Public Vol_Server As String
Public Storage_Path As String
Public Dataset_Folder As String
Public Results_Folder As String
Public Created As String
Public State As Long
Public NET_TICFit As Double         ' Not used in DB Schema Version 2
Public NET_Slope As Double          ' Not used in DB Schema Version 2
Public NET_Intercept As Double      ' Not used in DB Schema Version 2
Public Total_Scans As Long
Public Duration As Long
Public GANET_Fit As Double
Public GANET_Slope As Double
Public GANET_Intercept As Double
Public DB_Schema_Version As Single

'fields that will go to T_Match_Making_Description table
Public MD_Reference_Job As Long
Public MD_File As String            'have to use for legacy
Public MD_Type As Long
Public MD_Parameters As String
Public MD_Date As String
Public MD_State As Long

'descriptive properties
Public Desc_DataFolder As String          'UNC of folder containing file
Public Desc_Type As String                'type of matching description

Public MTDB As LaMTDB                     'Organism Mass Tags database
Attribute MTDB.VB_VarHelpID = -1
Public ORFDB As LaORFDB                   'Organism ORF database

Public Parameters As New Collection       'parameters of analysis

Private Sub Class_Initialize()
Set MTDB = New LaMTDB
End Sub

Private Sub Class_Terminate()
Set MTDB = Nothing
End Sub

Public Function GetJobInfo() As String
'--------------------------------------------
'returns job info based on current properties
'--------------------------------------------
On Error Resume Next
Dim tmp As String

tmp = FLD_Job & "=" & Job & vbCrLf
tmp = tmp & FLD_Dataset & "=" & Dataset & vbCrLf
tmp = tmp & FLD_Dataset_ID & "=" & Dataset_ID & vbCrLf
tmp = tmp & FLD_Experiment & "=" & Experiment & vbCrLf
tmp = tmp & FLD_Organism & "=" & Organism & vbCrLf
tmp = tmp & FLD_Instrument_Class & "=" & Instrument_Class & vbCrLf
tmp = tmp & FLD_Analysis_Tool & "=" & Analysis_Tool & vbCrLf
tmp = tmp & FLD_Parameter_File_Name & "=" & Parameter_File_Name & vbCrLf
tmp = tmp & FLD_Settings_File_Name & "=" & Settings_File_Name & vbCrLf
tmp = tmp & FLD_Organism_DB_Name & "=" & Organism_DB_Name & vbCrLf
tmp = tmp & "Data folder=" & Desc_DataFolder & vbCrLf
tmp = tmp & FLD_Created & "=" & Created & vbCrLf
tmp = tmp & "State" & "=" & State & vbCrLf
If DB_Schema_Version < 2 Then
    tmp = tmp & FLD_NET_TICFit & "=" & NET_TICFit & vbCrLf
    tmp = tmp & FLD_NET_Slope & "=" & NET_Slope & vbCrLf
    tmp = tmp & FLD_NET_Intercept & "=" & NET_Intercept & vbCrLf
End If
tmp = tmp & FLD_GANET_Fit & "=" & GANET_Fit & vbCrLf
tmp = tmp & FLD_GANET_Slope & "=" & GANET_Slope & vbCrLf
tmp = tmp & FLD_GANET_Intercept & "=" & GANET_Intercept & vbCrLf
tmp = tmp & FLD_Total_Scans & "=" & Total_Scans & vbCrLf
tmp = tmp & FLD_Duration & "=" & Duration & vbCrLf
GetJobInfo = tmp
End Function


Public Function GetJobInfoTIC() As String
'----------------------------------------
'returns TIC alignment info
'----------------------------------------
On Error Resume Next
Dim tmp As String
tmp = FLD_Job & "=" & Job & vbCrLf
If DB_Schema_Version < 2 Then
    tmp = tmp & FLD_NET_TICFit & "=" & NET_TICFit & vbCrLf
    tmp = tmp & FLD_NET_Slope & "=" & NET_Slope & vbCrLf
    tmp = tmp & FLD_NET_Intercept & "=" & NET_Intercept & vbCrLf
Else
    tmp = tmp & "NET_TIC info is not defined for DB Schema Version " & Trim(DB_Schema_Version)
End If
GetJobInfoTIC = tmp
End Function


Public Function GetJobInfoGANET() As String
'----------------------------------------
'returns TIC alignment info
'----------------------------------------
On Error Resume Next
Dim tmp As String
tmp = FLD_Job & "=" & Job & vbCrLf
tmp = tmp & FLD_GANET_Fit & "=" & GANET_Fit & vbCrLf
tmp = tmp & FLD_GANET_Slope & "=" & GANET_Slope & vbCrLf
tmp = tmp & FLD_GANET_Intercept & "=" & GANET_Intercept & vbCrLf
GetJobInfoGANET = tmp
End Function


Public Function FillFADRecord() As Boolean
'----------------------------------------------------------------
'retrieves job record from the FTICR_Analysis_Description table
'----------------------------------------------------------------
Dim t_Name As String
Dim RetrieveSQL As String
Dim rsJob As adodb.Recordset
On Error GoTo err_FillFADRecord

t_Name = MTDB.DBStuff.Item(t_FAD).Value
If Len(t_Name) <= 0 Then GoTo err_FillFADRecord
RetrieveSQL = "SELECT * FROM " & t_Name & " WHERE " & FLD_Job & "=" & Job
MTDB.cn.Open
Set rsJob = New adodb.Recordset
Set rsJob.ActiveConnection = MTDB.cn
rsJob.CursorLocation = adUseClient
rsJob.CursorType = adOpenForwardOnly
Set rsJob = MTDB.cn.Execute(RetrieveSQL)

If rsJob.EOF Then
    ' No match
    Dataset = ""
    Dataset_ID = 0
    Experiment = ""
    Organism = ""
    Instrument_Class = ""
    Analysis_Tool = ""
    Parameter_File_Name = ""
    Settings_File_Name = ""
    Organism_DB_Name = ""
    Vol_Client = ""
    Vol_Server = ""
    Storage_Path = ""
    Dataset_Folder = ""
    Results_Folder = ""
    State = 0
        
    NET_TICFit = 0
    NET_Slope = 0
    NET_Intercept = 0
    
    GANET_Fit = 0
    GANET_Slope = 0
    GANET_Intercept = 0
    Total_Scans = 0
    Duration = 0
    
    Desc_DataFolder = ""
Else
    'following line is not neccessary but I want to trigger error if
    Job = rsJob.Fields(FLD_Job).Value     'no records is returned
    On Error Resume Next
    With rsJob
        Dataset = .Fields(FLD_Dataset).Value
        Dataset_ID = .Fields(FLD_Dataset_ID).Value
        Experiment = .Fields(FLD_Experiment).Value
        Organism = .Fields(FLD_Organism).Value
        Instrument_Class = .Fields(FLD_Instrument_Class).Value
        Analysis_Tool = .Fields(FLD_Analysis_Tool).Value
        Parameter_File_Name = .Fields(FLD_Parameter_File_Name).Value
        Settings_File_Name = .Fields(FLD_Settings_File_Name).Value
        Organism_DB_Name = .Fields(FLD_Organism_DB_Name).Value
        Vol_Client = .Fields(FLD_Vol_Client).Value
        Vol_Server = .Fields(FLD_Vol_Server).Value
        Storage_Path = .Fields(FLD_Storage_Path).Value
        Dataset_Folder = .Fields(FLD_Dataset_Folder).Value
        Results_Folder = .Fields(FLD_Results_Folder).Value
        State = .Fields(FLD_State).Value
            
        If DB_Schema_Version < 2 Then
            NET_TICFit = .Fields(FLD_NET_TICFit).Value
            NET_Slope = .Fields(FLD_NET_Slope).Value
            NET_Intercept = .Fields(FLD_NET_Intercept).Value
        End If
        
        GANET_Fit = .Fields(FLD_GANET_Fit).Value
        GANET_Slope = .Fields(FLD_GANET_Slope).Value
        GANET_Intercept = .Fields(FLD_GANET_Intercept).Value
        Total_Scans = .Fields(FLD_Total_Scans).Value
        Duration = .Fields(FLD_Duration).Value
    End With
    
    'edit fields that should have folder separator at the end
    AddFolderSeparator Vol_Client
    AddFolderSeparator Vol_Server
    AddFolderSeparator Storage_Path
    AddFolderSeparator Dataset_Folder
    AddFolderSeparator Results_Folder
    Desc_DataFolder = Vol_Client & Storage_Path & Dataset_Folder & Results_Folder

End If

On Error GoTo err_FillFADRecord
rsJob.Close
MTDB.cn.Close

FillFADRecord = True
Exit Function

err_FillFADRecord:
If Err.Number > 0 Then
   Resume Next
Else            'FTICR Analysis Description table entry not found
   MTDB.DBStatus = dbsErr
End If
End Function

Public Function GetParameters() As String
'------------------------------------------------
'list current parameters as a list of Name=Value
'------------------------------------------------
Dim i As Long
Dim nv As New NameValue
Dim tmp As String
On Error Resume Next
For i = 1 To Parameters.Count
    Set nv = Parameters(i)
    tmp = tmp & nv.NameValueStr & vbCrLf
Next i
GetParameters = tmp
End Function

Public Function DestroyParameters()
'-------------------------------
'destroys DBStuff collection
'-------------------------------
Dim i As Long
On Error Resume Next
For i = 1 To Parameters.Count
    Parameters.Remove 1
Next i
End Function

