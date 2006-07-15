VERSION 5.00
Begin VB.Form frmNETFormulaEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NET Formula"
   ClientHeight    =   4395
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraControls 
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   3255
      Begin VB.CommandButton cmdDisableCustomNETs 
         Caption         =   "&Disable Custom NETs"
         Height          =   375
         Left            =   1080
         TabIndex        =   13
         Top             =   2085
         Width           =   2175
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   1605
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   1560
         TabIndex        =   11
         Top             =   1605
         Width           =   975
      End
      Begin VB.TextBox txtTICSlope 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Text            =   "0"
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox txtTICIntercept 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Text            =   "0"
         Top             =   465
         Width           =   1335
      End
      Begin VB.TextBox txtGANETSlope 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Text            =   "0"
         Top             =   825
         Width           =   1335
      End
      Begin VB.TextBox txtGANETIntercept 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Text            =   "0"
         Top             =   1185
         Width           =   1335
      End
      Begin VB.CommandButton cmdResetToGeneric 
         Caption         =   "&Reset gaNET to Generic"
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   2085
         Width           =   2175
      End
      Begin VB.CommandButton cmdResetRangeToZeroToOne 
         Caption         =   "&Reset gaNET to give NET 0 to 1"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   2565
         Width           =   2655
      End
      Begin VB.Label lblTICSlopeOrNetEquation 
         Caption         =   "TIC Slope"
         Height          =   255
         Left            =   150
         TabIndex        =   2
         Top             =   165
         Width           =   1215
      End
      Begin VB.Label lblValue 
         Caption         =   "TIC Intercept"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   4
         Top             =   525
         Width           =   1215
      End
      Begin VB.Label lblValue 
         Caption         =   "gaNET Slope"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   6
         Top             =   885
         Width           =   1215
      End
      Begin VB.Label lblValue 
         Caption         =   "gaNET Intercept"
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   8
         Top             =   1245
         Width           =   1215
      End
   End
   Begin VB.Label lblCustomNETMessage 
      Caption         =   $"frmNETFormulaEditor.frx":0000
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmNETFormulaEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'created: 01/31/2003 nt
'last modified: 01/31/2003 nt
'-----------------------------------------------------------
Option Explicit

Private CallerID As Long
Private bLoading As Boolean

Private mNeedToClearCustomNETs As Boolean

' The following is True if GelAnalysis(CallerID) is defined
Private mUseGelAnalysisObject As Boolean

Private Sub ApplyChanges(blnCloseForm As Boolean)
    Dim blnUpdated As Boolean
    Dim dblNewValue As Double
    
    If mNeedToClearCustomNETs Then
        CustomNETsClear CallerID
    End If
    
    If GelData(CallerID).CustomNETsDefined Then
        ' Nothing to update
    Else
        If mUseGelAnalysisObject Then
            With GelAnalysis(CallerID)
                If CheckForNewValue(.NET_Slope, txtTICSlope, dblNewValue) Then
                    .NET_Slope = dblNewValue
                    blnUpdated = True
                End If
                
                If CheckForNewValue(.NET_Intercept, txtTICIntercept, dblNewValue) Then
                    .NET_Intercept = dblNewValue
                    blnUpdated = True
                End If
                
                If CheckForNewValue(.GANET_Slope, txtGANETSlope, dblNewValue) Then
                    .GANET_Slope = dblNewValue
                    blnUpdated = True
                End If
                
                If CheckForNewValue(.GANET_Intercept, txtGANETIntercept, dblNewValue) Then
                    .GANET_Intercept = dblNewValue
                    blnUpdated = True
                End If
                
                If blnUpdated Then
                    AddToAnalysisHistory CallerID, "NET parameters updated; NET Slope = " & Trim(.NET_Slope) & "; NET Intercept = " & Trim(.NET_Intercept) & "; GANET Slope = " & Trim(.GANET_Slope) & "; GANET Intercept = " & Trim(.GANET_Intercept)
                End If
            End With
        Else
            With GelUMCNETAdjDef(CallerID)
                If .NETFormula <> txtTICSlope And Len(txtTICSlope) > 0 Then
                    .NETFormula = txtTICSlope
                    AddToAnalysisHistory CallerID, "NET parameters updated; New NET Formula = " & .NETFormula
                End If
            End With
        End If
    End If
    
    If blnCloseForm Then Unload Me

End Sub

Private Function CheckForNewValue(ByVal dblCurrentValue As Double, objTextbox As TextBox, ByRef dblNewValue As Double) As Boolean
    
    If IsNumeric(objTextbox) Then
        dblNewValue = CDbl(objTextbox)
        If dblCurrentValue <> dblNewValue Then
            CheckForNewValue = True
        Else
            CheckForNewValue = False
        End If
    Else
        CheckForNewValue = False
    End If
    
End Function

Private Sub DisplayNETParameterValue(txtThisTextBox As VB.TextBox, dblValue As Double)
    If dblValue = 0 Then
        txtThisTextBox.Text = "0"
    Else
        txtThisTextBox.Text = DoubleToStringScientific(dblValue, 5)
    End If
End Sub

Private Sub ResetRangeToZeroToOne()
    Dim lngMinFN As Long, lngMaxFN As Long, lngScanRangeTotal As Long
    
    GetScanRange CallerID, lngMinFN, lngMaxFN, lngScanRangeTotal
    
    If lngScanRangeTotal > 0 Then
        If mUseGelAnalysisObject Then
            txtGANETSlope = DoubleToStringScientific(1 / CDbl(lngScanRangeTotal), 7)
            txtGANETIntercept = "0"
        Else
            txtTICSlope = ConstructNETFormula(1 / CDbl(lngScanRangeTotal), 0)
        End If
    Else
        MsgBox "Unable to reset to Generic values since total number of scans is 0", vbExclamation + vbOKOnly, "Error"
    End If

End Sub

Private Sub ShowHideControls()

    Dim blnShowSlopeIntercept As Boolean
    
    blnShowSlopeIntercept = Not GelData(CallerID).CustomNETsDefined
    
    Me.ScaleMode = vbTwips

    If blnShowSlopeIntercept Then
        fraControls.Top = 120
    Else
        fraControls.Top = lblCustomNETMessage.Top + lblCustomNETMessage.Height + 120
    End If
    
    Me.Height = fraControls.Top + fraControls.Height + (Me.Height - Me.ScaleHeight) + 120
    lblCustomNETMessage.Visible = Not blnShowSlopeIntercept
    
    
    cmdDisableCustomNETs.Left = cmdResetToGeneric.Left
    cmdDisableCustomNETs.Top = cmdResetToGeneric.Top

    txtTICSlope.Visible = True
    txtTICIntercept.Visible = mUseGelAnalysisObject
    txtGANETSlope.Visible = mUseGelAnalysisObject
    txtGANETIntercept.Visible = mUseGelAnalysisObject
    
    txtTICSlope.Enabled = blnShowSlopeIntercept
    txtTICIntercept.Enabled = blnShowSlopeIntercept
    txtGANETSlope.Enabled = blnShowSlopeIntercept
    txtGANETIntercept.Enabled = blnShowSlopeIntercept
    
    lblValue(0).Visible = mUseGelAnalysisObject
    lblValue(1).Visible = mUseGelAnalysisObject
    lblValue(2).Visible = mUseGelAnalysisObject
    
    lblTICSlopeOrNetEquation.Enabled = blnShowSlopeIntercept
    lblValue(0).Enabled = blnShowSlopeIntercept
    lblValue(1).Enabled = blnShowSlopeIntercept
    lblValue(2).Enabled = blnShowSlopeIntercept
        
    If mUseGelAnalysisObject Then
        lblTICSlopeOrNetEquation.Caption = "TIC Slope"
        txtTICSlope.Top = lblTICSlopeOrNetEquation.Top - 40
        txtTICSlope.Left = txtTICIntercept.Left
        txtTICSlope.width = txtTICIntercept.width
        txtTICSlope.Alignment = vbRightJustify
    Else
        lblTICSlopeOrNetEquation.Caption = "NET Equation"
        txtTICSlope.Top = txtTICIntercept.Top
        txtTICSlope.Left = 120
        txtTICSlope.width = 2700
        txtTICSlope.Alignment = vbLeftJustify
    End If
    
    cmdResetRangeToZeroToOne.Visible = blnShowSlopeIntercept
    cmdResetToGeneric.Visible = blnShowSlopeIntercept
    cmdDisableCustomNETs.Visible = Not blnShowSlopeIntercept
   
End Sub

Private Sub cmdCancel_Click()
    If mNeedToClearCustomNETs Then
        CustomNETsValidateStatus CallerID
    End If
    Unload Me
End Sub

Private Sub cmdDisableCustomNETs_Click()
    GelData(CallerID).CustomNETsDefined = False
    mNeedToClearCustomNETs = True
    ShowHideControls
End Sub

Private Sub cmdOK_Click()
    ApplyChanges True
End Sub

Private Sub cmdResetRangeToZeroToOne_Click()
    ResetRangeToZeroToOne
End Sub

Private Sub cmdResetToGeneric_Click()

    On Error GoTo ResetToGenericErrorHandler

    If GelUMCNETAdjDef(CallerID).InitialSlope <> 0 Then
        If mUseGelAnalysisObject Then
            txtGANETSlope = GelUMCNETAdjDef(CallerID).InitialSlope
            txtGANETIntercept = GelUMCNETAdjDef(CallerID).InitialIntercept
        Else
            txtTICSlope = ConstructNETFormula(GelUMCNETAdjDef(CallerID).InitialSlope, GelUMCNETAdjDef(CallerID).InitialIntercept)
        End If
    Else
        ResetRangeToZeroToOne
    End If

Exit Sub

ResetToGenericErrorHandler:
    ResetRangeToZeroToOne
    
End Sub

Private Sub Form_Activate()
    If bLoading Then
        CallerID = Me.Tag
        mNeedToClearCustomNETs = False
        
        If Not GelAnalysis(CallerID) Is Nothing Then
            mUseGelAnalysisObject = True
        Else
            mUseGelAnalysisObject = False
        End If
        
        ShowHideControls
        
        If mUseGelAnalysisObject Then
            With GelAnalysis(CallerID)
                DisplayNETParameterValue txtTICSlope, .NET_Slope
                DisplayNETParameterValue txtTICIntercept, .NET_Intercept
                DisplayNETParameterValue txtGANETSlope, .GANET_Slope
                DisplayNETParameterValue txtGANETIntercept, .GANET_Intercept
            End With
        Else
            txtTICSlope = GelUMCNETAdjDef(CallerID).NETFormula
        End If
        bLoading = False
    End If
End Sub

Private Sub Form_Load()
    bLoading = True
End Sub

Private Sub txtGANETIntercept_LostFocus()
    If mUseGelAnalysisObject And Not IsNumeric(txtGANETIntercept.Text) Then
        DisplayNETParameterValue txtGANETIntercept, GelAnalysis(CallerID).GANET_Intercept
    End If
End Sub

Private Sub txtGANETSlope_LostFocus()
    If mUseGelAnalysisObject And Not IsNumeric(txtGANETSlope.Text) Then
        DisplayNETParameterValue txtGANETSlope, GelAnalysis(CallerID).GANET_Slope
    End If
End Sub

Private Sub txtTICIntercept_LostFocus()
    If mUseGelAnalysisObject And Not IsNumeric(txtTICIntercept.Text) Then
        DisplayNETParameterValue txtTICIntercept, GelAnalysis(CallerID).NET_Intercept
    End If
End Sub

Private Sub txtTICSlope_LostFocus()
    If mUseGelAnalysisObject And Not IsNumeric(txtTICSlope.Text) Then
        DisplayNETParameterValue txtTICSlope, GelAnalysis(CallerID).NET_Slope
End If
End Sub
