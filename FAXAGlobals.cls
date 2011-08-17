VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FAXAGlobals"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'global declaration, procedures and functions
'for FAXA componenet
'created: 06/17/2001 nt
'last modified: 06/14/2001 nt
'----------------------------------------------------
Public Enum FAXA_Processing_Type
    fptMassLock = 0
    fptMassMatch = 1
End Enum


Public Enum FAXA_DB_Type
    dbtOrgMTDB = 1
    dbtDMS = 2
    dbtExport = 3
End Enum

Public Enum FAXA_DB_Status
    dbsLoaded = 0       'data from database loaded
    dbsReady = 1        'elements of connection found but not loaded
    dbsNotReady = 2     'elements of connection not found
    dbsEmpty = 3        'loaded but nothing found
    dbsErr = -1         'error-wrong format or any error after attempt to load data
End Enum

Public Enum FAXA_Errors
    faxa_ANY_ERROR = -1
    faxa_INIT_FILE_NOT_FOUND = -1
    faxa_INIT_FILE_SECTION_NOT_FOUND = -2
    faxa_INIT_FILE_ANY_ERROR = -3
End Enum

Public TheMessage As String         'used to comunicate with
Public Event IHaveMessage()         'other applications

'global constants simulation
Public Property Get INIT_Comment() As String
INIT_Comment = "//"
End Property

Public Property Get INIT_Value() As String
INIT_Value = "="
End Property

Public Property Get INIT_SectionEnd() As String
INIT_SectionEnd = "[END]"
End Property

Public Property Get INIT_None() As String
INIT_None = "None"
End Property

Public Property Get SECTION_MassTagDB() As String
SECTION_MassTagDB = "[Organism Mass Tag Database]"
End Property

Public Property Get SECTION_Parameters_Lock() As String
SECTION_Parameters_Lock = "[Parameters Lock]"
End Property

Public Property Get SECTION_Parameters_Match() As String
SECTION_Parameters_Match = "[Parameters Match]"
End Property


