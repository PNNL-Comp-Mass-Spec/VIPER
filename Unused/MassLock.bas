Attribute VB_Name = "Module16"
'last modified 06/21/2000 nt
'-----------------------------------------------------
Option Explicit

Public Const glLM_PROPAGATE_NO = 0
'Public Const glLM_PROPAGATE_1 = 1
'Public Const glLM_PROPAGATE_2 = 2
'Public Const glLM_PROPAGATE_USER = 3

Public Const glLM_MULTI_INTENSITY = 0
'Public Const glLM_MULTI_FIT = 1
'Public Const glLM_MULTI_AMTFIT = 2
'Public Const glLM_MULTI_USER = 3

Public Const glLM_AMT_MULTI_FIT = 0
'Public Const glLM_AMT_MULTI_USER = 1

Public Const glLM_SAVE_NOT = 0
Public Const glLM_SAVE_ORIGINAL = 1
Public Const glLM_SAVE_NEW = 2

Public Const glMASS_LOCKER_MARK = ">L<"
'mass lock options structures

Public Type AMTLMDefinition
    lmScope As Integer
    lmIsoField As Integer           ' 6, 7, or 8
    lmPropagate As Integer
    lmMultiCandidates As Integer
    lmMultiAMTHits As Integer
    lmSaveResults As Integer
End Type

'similar like UMC & UC definition variables; this
'thing is alive for the duration of application
Public amtlmDef As AMTLMDefinition

' Unused Function (May 2003)
'''Public Function GetAMTLMDefDesc() As String
'''Dim sTmp As String
'''With amtlmDef
'''    Select Case .lmScope
'''    Case glScope.glSc_All
'''         sTmp = "AMT lock mass on all data points." & vbCrLf
'''    Case glScope.glSc_Current
'''         sTmp = "AMT lock mass on data points currently in scope." & vbCrLf
'''    End Select
'''    Select Case .lmIsoField
'''    Case 6
'''         sTmp = sTmp & "MW field(Isotopic): Average" & vbCrLf
'''    Case 7
'''         sTmp = sTmp & "MW field(Isotopic): Monoisotopic" & vbCrLf
'''    Case 8
'''         sTmp = sTmp & "MW field(Isotopic): The Most Abundant" & vbCrLf
'''    End Select
'''    Select Case .lmMultiCandidates
'''    Case glLM_MULTI_INTENSITY
'''         sTmp = sTmp & "In case of multiple lock mass candidates selected candidate with highest intensity." & vbCrLf
'''    Case glLM_MULTI_FIT
'''         sTmp = sTmp & "In case of multiple lock mass candidates selected candidate with best calculated fit." & vbCrLf
'''    Case glLM_MULTI_AMTFIT
'''         sTmp = sTmp & "In case of multiple lock mass candidates selected candidate with smallest AMT error." & vbCrLf
'''    Case glLM_MULTI_USER
'''         sTmp = sTmp & "In case of multiple lock mass candidates selected user's choice." & vbCrLf
'''    End Select
'''    Select Case .lmMultiAMTHits
'''    Case glLM_AMT_MULTI_FIT
'''         sTmp = sTmp & "In case of multiple AMT hits lock mass on AMT with smallest error." & vbCrLf
'''    Case glLM_AMT_MULTI_USER
'''         sTmp = sTmp & "In case of multiple AMT hits lock mass on AMT selected by user(except when best candidate was selected from smallest AMT error)." & vbCrLf
'''    End Select
'''    Select Case .lmPropagate
'''    Case glLM_PROPAGATE_NO
'''         sTmp = sTmp & "No propagation method applied."
'''    Case glLM_PROPAGATE_1
'''         sTmp = sTmp & "Propagation method 1 applied."
'''    Case glLM_PROPAGATE_2
'''         sTmp = sTmp & "Propagation method 2 applied."
'''    Case glLM_PROPAGATE_USER
'''    End Select
'''End With
'''GetAMTLMDefDesc = sTmp
'''End Function
