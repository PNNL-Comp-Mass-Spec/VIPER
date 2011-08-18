VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AnalysisInitiator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------
'initiator of analysis dialog - opens dialog and raises
'event when dialog is closed
'NOTE: dialog opens as non-modal; since this runs as an
'Active X EXE there are some issues with modal dialogs;
'however I don't like this non-modal solution!!!
'------------------------------------------------------
'created: 06/14/2001 nt
'last modified: 06/14/2001 nt
'------------------------------------------------------
Option Explicit

'actual analysis settings(no need for events here)
Public NewAnalysis As FTICRAnalysis
'client notification event
Public Event DialogClosed()

'actual analysis dialog
Dim WithEvents ADialog As frmNewAnalysis
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
Set ADialog = New frmNewAnalysis
ADialog.InitFileName = InitFileName
ADialog.Show 0
End Sub
