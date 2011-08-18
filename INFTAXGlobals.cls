VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INFTAXGlobals"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------
'global declaration, procedures and functions
'for FAXA componenet
'----------------------------------------------------
'created: 06/17/2001 nt
'last modified: 02/13/2002 nt
'----------------------------------------------------
'Public Enum FAXA_MTDirectory
'    MTDirConnStr = 0
'    MTDirRetrieve = 1
'    MTDirCnStrFld = 2
'    MTDirDescFld = 3
'    MTDirNameFld = 4
'    MTDirStateFld = 5
'End Enum

Public Enum FAXA_MTSMasterDirectory
    MTSMasterConnStr = 0
    MTSMasterSPRetrieve = 1
    MTSMasterNameFld = 2
    MTSMasterDescFld = 3
    MTSMasterStateFld = 4
    MTSMasterServerNameFld = 5
End Enum

Public Enum FAXA_Processing_Type
    fptDummy = -1
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

Public Property Get DEFAULT_MTS_MASTER_CONN_STRING() As String
    DEFAULT_MTS_MASTER_CONN_STRING = "Provider=sqloledb;Data Source=pogo;Initial Catalog=MTS_Master;User ID=MTUser;Password=mt4fun"
End Property

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


Public Property Get SECTION_DMS() As String
SECTION_DMS = "[DMS Database]"
End Property

'
'Public Property Get SECTION_MT_Directory() As String
'SECTION_MT_Directory = "[MT Directory Database]"
'End Property

Public Property Get SECTION_MTS_Master_DB() As String
SECTION_MTS_Master_DB = "[MTS Master Database]"
End Property


Public Property Get SECTION_MTDB_Schema() As String
SECTION_MTDB_Schema = "[Org MT DB Schema]"
End Property


Public Property Get SECTION_Parameters_Lock() As String
SECTION_Parameters_Lock = "[Parameters Lock]"
End Property


Public Property Get SECTION_Parameters_Match_Standard() As String
SECTION_Parameters_Match_Standard = "[Parameters Match; Standard]"
End Property


Public Property Get SECTION_Parameters_Match_N14N15() As String
SECTION_Parameters_Match_N14N15 = "[Parameters Match; N14/N15]"
End Property


Public Property Get SECTION_Parameters_Match_ICAT() As String
SECTION_Parameters_Match_ICAT = "[Parameters Match; ICAT]"
End Property


Public Property Get SECTION_Parameters_Match_LysC12C13() As String
SECTION_Parameters_Match_LysC12C13 = "[Parameters Match; Lys-C12/C13]"
End Property


Public Property Get SECTION_Parameters_Match_PEO() As String
SECTION_Parameters_Match_PEO = "[Parameters Match; PEO]"
End Property


Public Property Get SECTION_Parameters_Match_PhIAT() As String
SECTION_Parameters_Match_PhIAT = "[Parameters Match; PhIAT]"
End Property


Public Property Get SECTION_Parameters_Match_Any() As String
'-----------------------------------------------------------
'generic parameters match contained in each section that
'describes match making
'-----------------------------------------------------------
SECTION_Parameters_Match_Any = "Parameters Match"
End Property


Public Property Get SECTION_URL() As String
SECTION_URL = "URL-DMS"
End Property


Public Property Get SECTION_Parameters_UMC() As String
SECTION_Parameters_UMC = "[Parameters UMC]"
End Property

Public Property Get NAME_SUBSET() As String
NAME_SUBSET = "MTSubset ID"
End Property

Public Property Get NAME_INC_LIST() As String
NAME_INC_LIST = "Search Inclusion List"
End Property

Public Property Get VALUE_TRUE() As String
VALUE_TRUE = "True"
End Property

Public Property Get VALUE_FALSE() As String
VALUE_FALSE = "False"
End Property



