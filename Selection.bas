Attribute VB_Name = "Module10"
'module with procedures related to points selection
'last modified 09/10/2000 nt
Option Explicit

Public Const glSEL_DF_MW = 0
Public Const glSEL_DF_INTENSITY = 1
Public Const glSEL_DF_FIT = 2
Public Const glSEL_DF_ER = 3

Public Enum ssrfSelectionStatsResultFormatConstants
    ssrfMinimum = 0
    ssrfMaximum = 1
    ssrfRange = 2
    ssrfAverage = 3
    ssrfStDev = 4
End Enum

Private Const glRES_NA = "N/A"
Private Const glRES_ERR = "Error"

Public Sub ExcludeAllButSelection(ByVal Ind As Long)
'Excludes and clears all points except those selected from the visible graph
Dim i As Long, iSel As Long

'On Error Resume Next
With GelBody(Ind).GelSel
  If .CSSelCnt > 0 Then
    ' First exclude all of the points
    For i = 1 To GelDraw(Ind).CSCount
        GelDraw(Ind).CSID(i) = -Abs(GelDraw(Ind).CSID(i))
    Next i
    
    ' Now show the selected ones
    For i = 1 To .CSSelCnt
        iSel = .value(i, glCSType)
        GelDraw(Ind).CSID(iSel) = Abs(GelDraw(Ind).CSID(iSel))
    Next i
  End If
  
  If .IsoSelCnt > 0 Then
    ' First exclude all of the points
    For i = 1 To GelDraw(Ind).IsoCount
        GelDraw(Ind).IsoID(i) = -Abs(GelDraw(Ind).IsoID(i))
    Next i
    
    ' Now show the selected ones
    For i = 1 To .IsoSelCnt
        iSel = .value(i, glIsoType)
        GelDraw(Ind).IsoID(iSel) = Abs(GelDraw(Ind).IsoID(iSel))
    Next i
  End If
  .Clear
End With
End Sub

Public Sub ExcludeSelection(ByVal Ind As Long)
'exclude and clears selection from the visible graph
Dim i As Long, iSel As Long

On Error Resume Next
With GelBody(Ind).GelSel
  If .CSSelCnt > 0 Then
     For i = 1 To .CSSelCnt
         iSel = .value(i, glCSType)
         GelDraw(Ind).CSID(iSel) = -Abs(GelDraw(Ind).CSID(iSel))
     Next i
  End If
  If .IsoSelCnt > 0 Then
     For i = 1 To .IsoSelCnt
         iSel = .value(i, glIsoType)
         GelDraw(Ind).IsoID(iSel) = -Abs(GelDraw(Ind).IsoID(iSel))
     Next i
  End If
  .Clear
End With
End Sub

Public Function GetSelectionFieldNumeric(ByVal Ind As Long, _
                                         ByVal Field As Integer, _
                                         ByRef Values() As Double) As Long
'Ind - GelBody index
'Field - field we need to fill
'Values - array to be filled
'fills Values with numbers from GelData and returns total number of selected points
Dim i As Long, ID As Long
Dim SelCnt As Long
On Error Resume Next

SelCnt = 0
With GelBody(Ind).GelSel
   If .CSSelCnt > 0 Then
      ReDim Values(1 To .CSSelCnt)
      For i = 1 To .CSSelCnt
         ID = .value(i, glCSType)
         SelCnt = SelCnt + 1
         Values(SelCnt) = GetIsoOrCSDataByField(Ind, GelData(Ind).CSData(ID), Field, True)
      Next i
   End If
   If .IsoSelCnt > 0 Then
      If UBound(Values) > 0 Then
         ReDim Preserve Values(1 To UBound(Values) + .IsoSelCnt)
      Else
         ReDim Values(1 To .IsoSelCnt)
      End If
      If Err.Number = 9 Then    'subscript out of range
         Err.Clear
         ReDim Values(1 To .IsoSelCnt)
      End If
      For i = 1 To .IsoSelCnt
         ID = .value(i, glIsoType)
         SelCnt = SelCnt + 1
         Values(SelCnt) = GetIsoOrCSDataByField(Ind, GelData(Ind).IsoData(ID), Field, False)
      Next i
   End If
End With
If SelCnt > 0 Then
   ReDim Preserve Values(1 To SelCnt)
End If
GetSelectionFieldNumeric = SelCnt
End Function

Public Function GetSelectionFieldMatchingIDs(ByVal Ind As Long, _
                                      ByVal Field As Integer, _
                                      ByRef Values() As String) As Long
'Ind - GelBody index
'Field - field we need to fill
'Values - array to be filled
'fills Values with matching identifications from GelData.IsoData().MTID and returns total number of selected points
Dim i As Long, ID As Long
Dim SelCnt As Long
On Error Resume Next

' This function only handles searching for PMT matches in .IsoData().MTID or .CSData().MTID
Debug.Assert Field <> glFIELD_ID

SelCnt = 0
With GelBody(Ind).GelSel
   If .CSSelCnt > 0 Then
      ReDim Values(1 To .CSSelCnt)
      For i = 1 To .CSSelCnt
         ID = .value(i, glCSType)
         SelCnt = SelCnt + 1
         Values(SelCnt) = GelData(Ind).CSData(ID).MTID
      Next i
   End If
   If .IsoSelCnt > 0 Then
      If UBound(Values) > 0 Then
         ReDim Preserve Values(1 To UBound(Values) + .IsoSelCnt)
      Else
         ReDim Values(1 To .IsoSelCnt)
      End If
      If Err.Number = 9 Then    'subscript out of range
         Err.Clear
         ReDim Values(1 To .IsoSelCnt)
      End If
      For i = 1 To .IsoSelCnt
         ID = .value(i, glIsoType)
         SelCnt = SelCnt + 1
         Values(SelCnt) = GelData(Ind).IsoData(ID).MTID
      Next i
   End If
End With
If SelCnt > 0 Then
   ReDim Preserve Values(1 To SelCnt)
End If
GetSelectionFieldMatchingIDs = SelCnt
End Function

' Unused function (February 2005)
''Public Function GetSelectionFieldTextAsNumeric(ByVal Ind As Long, _
''                                               ByVal Field As Integer, _
''                                               ByRef Values() As Double) As Long
'''Ind - GelBody index
'''Field - field we need to fill
'''Values - array to be filled
'''fills Values with numbers from GelData and returns total number of selected points
''Dim CSFld As Integer
''Dim IsoFld As Integer
''Dim i As Long, ID As Long
''Dim SelCnt As Long
''On Error Resume Next
''
''GetFieldIndexes Ind, Field, CSFld, IsoFld
''SelCnt = 0
''With GelBody(Ind).GelSel
''   If .CSSelCnt > 0 Then
''      ReDim Values(1 To .CSSelCnt)
''      For i = 1 To .CSSelCnt
''         ID = .value(i, glCSType)
''         If IsNumeric(GelData(Ind).CSVar(ID, CSFld)) Then
''            SelCnt = SelCnt + 1
''            Values(SelCnt) = CDbl(GelData(Ind).CSVar(ID, CSFld))
''         End If
''      Next i
''   End If
''   If .IsoSelCnt > 0 Then
''      If UBound(Values) > 0 Then
''         ReDim Preserve Values(1 To UBound(Values) + .IsoSelCnt)
''      Else
''         ReDim Values(1 To .IsoSelCnt)
''      End If
''      If Err.Number = 9 Then    'subscript out of range
''         Err.Clear
''         ReDim Values(1 To .IsoSelCnt)
''      End If
''      For i = 1 To .IsoSelCnt
''         ID = .value(i, glIsoType)
''         If IsNumeric(GelData(Ind).IsoVar(ID, IsoFld)) Then
''            SelCnt = SelCnt + 1
''            Values(SelCnt) = CDbl(GelData(Ind).IsoVar(ID, IsoFld))
''         End If
''      Next i
''   End If
''End With
''If SelCnt > 0 Then
''   ReDim Preserve Values(1 To SelCnt)
''End If
''GetSelectionFieldTextAsNumeric = SelCnt
''End Function

Public Sub SelStatsCompute(ByVal Ind As Long, _
                            ByVal SearchField As Integer, _
                            ByVal CustFormat As String, _
                            ByVal eFormat As ssrfSelectionStatsResultFormatConstants, _
                            udtResults As GelRes)

Dim Sel() As Long
Dim SelCnt As Long

Dim dblCSValues() As Double
Dim lngCSValuesCountDimmed As Long

Dim dblIsoValues() As Double
Dim lngIsoValuesCountDimmed As Long

Dim CSCnt As Long
Dim IsoCnt As Long
Dim i As Long

Dim objCSStats As New StatDoubles
Dim objIsoStats As New StatDoubles
Dim blnSuccess As Boolean

On Error GoTo SelStatsComputeErrorHandler

CSCnt = 0
IsoCnt = 0

With GelBody(Ind).GelSel
    If .CSSelCnt > 0 Then
        .GetCSSel Sel()
        SelCnt = UBound(Sel)
        If SelCnt > 0 Then
            lngCSValuesCountDimmed = SelCnt
            ReDim dblCSValues(0 To lngCSValuesCountDimmed - 1)
            
            With GelData(Ind)
                For i = 1 To SelCnt
                    CSCnt = CSCnt + 1
                    dblCSValues(CSCnt - 1) = GetIsoOrCSDataByField(Ind, .CSData(Sel(i)), SearchField, True)
                Next i
            End With
            If CSCnt > 0 Then
                ReDim Preserve dblCSValues(0 To CSCnt - 1)
                blnSuccess = objCSStats.Fill(dblCSValues())
                Debug.Assert blnSuccess
                
                Select Case eFormat
                Case ssrfMinimum: udtResults.CSRes = Format(objCSStats.Minimum, CustFormat)
                Case ssrfMaximum: udtResults.CSRes = Format(objCSStats.Maximum, CustFormat)
                Case ssrfRange: udtResults.CSRes = Format(objCSStats.Maximum - objCSStats.Minimum, CustFormat)
                Case ssrfAverage: udtResults.CSRes = Format(objCSStats.Mean, CustFormat)
                Case ssrfStDev: udtResults.CSRes = Format(objCSStats.StDev, CustFormat)
                Case Else
                   Debug.Assert False
                   udtResults.CSRes = glRES_ERR
                End Select
            Else
                udtResults.CSRes = "0 found"
            End If
        Else
            udtResults.CSRes = glRES_ERR
        End If
    Else
       udtResults.CSRes = glRES_NA
    End If
    
    If .IsoSelCnt > 0 Then
       .GetIsoSel Sel()
       SelCnt = UBound(Sel)
       If SelCnt > 0 Then
          
            lngIsoValuesCountDimmed = SelCnt
            ReDim dblIsoValues(0 To lngIsoValuesCountDimmed - 1)
            
            With GelData(Ind)
                For i = 1 To SelCnt
                    IsoCnt = IsoCnt + 1
                    dblIsoValues(IsoCnt - 1) = GetIsoOrCSDataByField(Ind, .IsoData(Sel(i)), SearchField, False)
                Next i
            End With
            If IsoCnt > 0 Then
                ReDim Preserve dblIsoValues(0 To IsoCnt - 1)
                blnSuccess = objIsoStats.Fill(dblIsoValues())
                Debug.Assert blnSuccess
                
                Select Case eFormat
                Case ssrfMinimum: udtResults.IsoRes = Format(objIsoStats.Minimum, CustFormat)
                Case ssrfMaximum: udtResults.IsoRes = Format(objIsoStats.Maximum, CustFormat)
                Case ssrfRange: udtResults.IsoRes = Format(objIsoStats.Maximum - objIsoStats.Minimum, CustFormat)
                Case ssrfAverage: udtResults.IsoRes = Format(objIsoStats.Mean, CustFormat)
                Case ssrfStDev: udtResults.IsoRes = Format(objIsoStats.StDev, CustFormat)
                Case Else
                   Debug.Assert False
                   udtResults.IsoRes = glRES_ERR
                End Select
            Else
                udtResults.IsoRes = "0 found"
            End If
        Else
           udtResults.IsoRes = glRES_ERR
        End If
    Else
        udtResults.IsoRes = glRES_NA
    End If
End With

With udtResults
    If InStr(1, .CSRes & .IsoRes, glRES_ERR) > 0 Then
        'if error in any report error in final
        .AllRes = glRES_ERR
    Else
        If IsNumeric(.CSRes) Then
            If IsNumeric(.IsoRes) Then
                ' Both are numeric; need to handle things differently depending on eFormat
                Select Case eFormat
                Case ssrfMinimum
                    If CDbl(.CSRes) < CDbl(.IsoRes) Then
                        .AllRes = .CSRes
                    Else
                        .AllRes = .IsoRes
                    End If
                Case ssrfMaximum, ssrfRange
                    If CDbl(.CSRes) > CDbl(.IsoRes) Then
                        .AllRes = .CSRes
                    Else
                        .AllRes = .IsoRes
                    End If
                Case ssrfAverage, ssrfStDev
                    If CSCnt + IsoCnt > 0 Then
                        ' Need to combine the CS and ER data; we'll copy the IsoValues data to the CSValues array
                        ReDim Preserve dblCSValues(0 To CSCnt + IsoCnt - 1)
                        For i = 0 To IsoCnt - 1
                            dblCSValues(CSCnt + i) = dblIsoValues(i)
                        Next i
                        objCSStats.Fill dblCSValues()
                        Select Case eFormat
                        Case ssrfAverage
                            .AllRes = Format(objCSStats.Mean)
                        Case ssrfStDev
                            ' eFormat = ssrfStDev
                            .AllRes = Format(objCSStats.StDev)
                        Case Else
                            Debug.Assert False
                            udtResults.AllRes = glRES_ERR
                        End Select
                    Else
                        .AllRes = "0 found"
                    End If
                End Select
            
            Else
               .AllRes = .CSRes
            End If
        Else
           .AllRes = .IsoRes
        End If
    End If
End With

Set objCSStats = Nothing
Set objIsoStats = Nothing

Exit Sub

SelStatsComputeErrorHandler:
Debug.Assert False
Debug.Print "Error: " & Err.Description
Resume Next
End Sub
