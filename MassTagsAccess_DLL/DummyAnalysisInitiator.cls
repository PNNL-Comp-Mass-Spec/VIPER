VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DummyAnalysisInitiator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------
'initiator of dummy FTICR analysis
'this is neccessary to do Out-Of-System analysis
'for such analysis it is necessary to associate
'data file with Mass Tag database so it can use
'lockers and mass tags tables
'------------------------------------------------------
'created: 06/14/2001 nt
'last modified: 10/15/2001 nt
'------------------------------------------------------
Option Explicit

'actual analysis settings(no need for events here)
Public NewAnalysis As FTICRAnalysis
'client notification event
Public Event DialogClosed()

'actual analysis dialog
Dim WithEvents ADialog As frmNewDummyAnalysis
Attribute ADialog.VB_VarHelpID = -1


Private Sub ADialog_AnalysisDialogClose()
'------------------------------------------------------
'expose new analysis object and notify client
'------------------------------------------------------
On Error Resume Next
Set NewAnalysis = ADialog.fAnalysis     'accept new analysis
Unload ADialog                          'unload form
Set ADialog = Nothing
RaiseEvent DialogClosed                 'notify client
End Sub

Public Sub GetNewAnalysisDialog(ByVal InitFileName As String)
'------------------------------------------------------------
'open new analysis dialog/non modal
'------------------------------------------------------------
Set ADialog = New frmNewDummyAnalysis
ADialog.InitFileName = InitFileName
ADialog.Show 0
End Sub

