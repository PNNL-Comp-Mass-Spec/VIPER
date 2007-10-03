VERSION 5.00
Begin VB.Form frmZoomRegionList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zoom Region List"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShowPepChromStds 
      Caption         =   "&Pep Chrom Standards"
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   1680
      Width           =   1935
   End
   Begin VB.ComboBox cboScanRangeUnits 
      Height          =   315
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtScanRange 
      Height          =   285
      Left            =   5400
      TabIndex        =   7
      Text            =   "50"
      Top             =   3720
      Width           =   855
   End
   Begin VB.ComboBox cboMassRangeUnits 
      Height          =   315
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox txtMassRange 
      Height          =   285
      Left            =   5400
      TabIndex        =   4
      Text            =   "0.02"
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtZoomRanges 
      Height          =   2295
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1680
      Width           =   3615
   End
   Begin VB.CommandButton cmdZoom 
      Caption         =   "&Zoom To Selected"
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblScanRange 
      Caption         =   "Scan Range"
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label lblMassRange 
      Caption         =   "Mass Range"
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label lblDirections 
      Caption         =   "Directions"
      Height          =   1335
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "frmZoomRegionList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CallerID As Long

Private Sub InitializeControls()
    lblDirections = "Enter a list of mass values to zoom to, one mass per line.  Optionally, add a central scan number or a start and end scan number, also separated by commas.  The mass and scan range can be specified using the controls below.  Optionally, enter a start and end scan value to override the default scan range.  Place the cursor on the desired line and click Zoom to Selected to show that region.  For example, entering 750,400 will zoom to 750 Da and scan number 400, with the tolerances given below.  Entering 750,400,420 will zoom to 750 Da, and a scan range of 400 to 420."
    
    With cboMassRangeUnits
        .Clear
        .AddItem "Da"
        .AddItem "ppm"
        .ListIndex = mruDa
    End With
    
    With cboScanRangeUnits
        .Clear
        .AddItem "Scan #"
        .AddItem "NET"
        .ListIndex = sruScan
    End With
    
    ShowPepChromStandards
End Sub

Private Sub ShowPepChromStandards()

    Dim strStds As String
    
    strStds = ""
    strStds = strStds & "823.4664,775" & vbCrLf
    strStds = strStds & "966.5471,465" & vbCrLf
    strStds = strStds & "1124.6229,1430" & vbCrLf
    strStds = strStds & "1478.0061,2185" & vbCrLf
    strStds = strStds & "1522.037443,2505" & vbCrLf

    ' 823.4664,0.249
    ' 966.5471,0.169
    ' 1124.6229,0.460
    ' 1478.0061,0.796
    ' 1522.037443,0.870

    txtZoomRanges = strStds
End Sub

Private Sub ZoomGelHandler()

    Dim dblMassMin As Double, dblMassMax As Double
    Dim sngScanOrNETMin As Single, sngScanOrNETMax As Single
    
    Dim blnUsePpm As Boolean
    Dim blnUseNET As Boolean
    
    Dim strZoomList As String
    Dim intCursorIndex As Integer
    Dim intMatchIndex As Integer
    
    Dim intParseCount As Integer
    Dim dblParsedVals() As Double
    
    Dim dblMassHalfWidthUser As Double
    Dim dblMassHalfWidthDa As Double
    Dim dblCentralMass As Double
    
    Dim sngScanWidth As Single
    Dim sngScanHalfWidth As Single
    
On Error GoTo ZoomGelHandlerErrorHandler

    If Not IsNumeric(txtMassRange) Then
        txtMassRange = "0.01"
        cboMassRangeUnits.ListIndex = mruDa
    End If
    dblMassHalfWidthUser = CDbl(txtMassRange)
    
    If Not IsNumeric(txtScanRange) Then
        txtScanRange = "50"
        cboScanRangeUnits.ListIndex = sruScan
    End If
    sngScanWidth = CSng(txtScanRange)
    
    If cboMassRangeUnits.ListIndex = mruPpm Then
        blnUsePpm = True
    Else
        blnUsePpm = False
    End If
    
    If cboScanRangeUnits.ListIndex = sruNet Then
        blnUseNET = True
    Else
        blnUseNET = False
    End If
    
    ' Determine the current desired zoom range by examining txtZoomRanges
    strZoomList = txtZoomRanges.Text
    If Len(strZoomList) = 0 Then
        MsgBox "The zoom list is empty.", vbInformation + vbOKOnly, "Nothing to Do"
        Exit Sub
    End If
    
    intCursorIndex = txtZoomRanges.SelStart
    
    ' Extract the text from strZoomList that contains the line the cursor is located on
    ' Find the next vbCrLf after intCursorIndex
    If intCursorIndex = 0 Then intCursorIndex = 1
    intMatchIndex = InStr(intCursorIndex, strZoomList, vbCrLf)
    If intMatchIndex > 0 Then
        strZoomList = Left(strZoomList, intMatchIndex - 1)
    End If
    
    ' Find the previous vbCrLf
    ' Do this by reversing the string, then looking for StrReverse(vbCrLf)
    strZoomList = StrReverse(strZoomList)
    intMatchIndex = InStr(strZoomList, StrReverse(vbCrLf))
    
    If intMatchIndex > 0 Then
        strZoomList = Left(strZoomList, intMatchIndex - 1)
    End If
    strZoomList = StrReverse(strZoomList)
    
    If Len(strZoomList) = 0 Then
        MsgBox "The cursor is on an empty line.", vbInformation + vbOKOnly, "Nothing to Do"
        Exit Sub
    End If
    
    ' Split strZoomList on commas or tabs
    intParseCount = ParseStringValuesDbl(strZoomList, dblParsedVals(), 4, "," & vbTab, "", False, True, False)
    
    If intParseCount <= 0 Then
        MsgBox "The cursor is on an invalid line.  It should contain a comma separated list of values: Central Mass, Central Scan  or  Central Mass, Start Scan, End Scan", vbExclamation + vbOKOnly, "Invalid Line"
    Else
        dblCentralMass = dblParsedVals(0)
                
        If blnUsePpm Then
            dblMassHalfWidthUser = PPMToMass(dblMassHalfWidthUser, dblCentralMass)
        End If
        dblMassHalfWidthDa = Abs(dblMassHalfWidthUser / 2)
        
        If dblMassHalfWidthDa < 0.00001 Then dblMassHalfWidthDa = 0.00001
        
        dblMassMin = dblCentralMass - dblMassHalfWidthDa
        dblMassMax = dblCentralMass + dblMassHalfWidthDa
                
        If intParseCount >= 3 Then
            ' Assume Central Mass, Start Scan, and End Scan
            sngScanOrNETMin = dblParsedVals(1)
            sngScanOrNETMax = dblParsedVals(2)
        ElseIf intParseCount >= 2 Then
            sngScanHalfWidth = Abs(sngScanWidth / 2)
            
            If blnUseNET Then
                If sngScanHalfWidth < 0.00001 Then sngScanHalfWidth = 0.00001
            Else
                If sngScanHalfWidth < 2 Then sngScanHalfWidth = 2
            End If
            
            ' Assume Central Mass and Central Scan
            sngScanOrNETMin = dblParsedVals(1) - sngScanHalfWidth
            sngScanOrNETMax = dblParsedVals(1) + sngScanHalfWidth
            
            If Not blnUseNET Then
                If sngScanOrNETMin < 0 Then sngScanOrNETMin = 0
                If sngScanOrNETMax < sngScanOrNETMin + 1 Then sngScanOrNETMax = sngScanOrNETMin + 1
            End If
        Else
            sngScanOrNETMin = -1
            sngScanOrNETMax = -1
            blnUseNET = False
        End If
        
        ZoomGelToDimensionsScanOrNET dblMassMin, dblMassMax, sngScanOrNETMin, sngScanOrNETMax, blnUseNET, CallerID
    End If
    
    Exit Sub

ZoomGelHandlerErrorHandler:
    Debug.Assert False
    
    LogErrors Err.Number, "ZoomGelHandler", Err.Description
    
End Sub

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub cmdShowPepChromStds_Click()
    ShowPepChromStandards
End Sub

Private Sub cmdZoom_Click()
    ZoomGelHandler
End Sub

Private Sub Form_Load()
    InitializeControls
End Sub
