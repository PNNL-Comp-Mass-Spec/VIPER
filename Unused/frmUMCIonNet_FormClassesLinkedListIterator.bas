Private Sub FormClassesFromNetsLinkedListIterator()

Dim CurrConnInd As Long
Dim bDone As Boolean
Dim i As Long
Dim lngTickCountLastUpdate As Long, lngNewTickCount As Long
Dim HUMCNetNextUnused() As Long
Dim HUMCNetPrevUnused() As Long
ReDim HUMCNetNextUnused(HUMCNetCnt)         ' The top index needs to be HUMCNetCnt since the last element will point to that location
ReDim HUMCNetPrevUnused(HUMCNetCnt)


    ' This for loop needs to go from 0 to HUMCNetCnt and not from 0 to HUMCNetCnt-1
    For i = 0 To HUMCNetCnt
        HUMCNetNextUnused(i) = i + 1
        HUMCNetPrevUnused(i) = i - 1
    Next i
    HUMCNetNextUnused(HUMCNetCnt) = HUMCNetCnt

      With GelUMCIon(CallerID)
         CurrConnInd = 0
         Do While CurrConnInd < HUMCNetCnt
            If HUMCNetUsed(CurrConnInd) = HUMCUsed Then         'already used; go next
               ' If this algorithm is fully optimized, then we shouldn't reach this
               CurrConnInd = CurrConnInd + 1
            Else                                       'new class; find the whole class
               MarkNetUsed CurrConnInd
               HUMCIsoUsed(.NetInd1(CurrConnInd)) = HUMCInUse:   HUMCIsoUsed(.NetInd2(CurrConnInd)) = HUMCInUse
               'first index  < last index (if this changes this function has to be revised)
               HUMCEquClsWk(0) = .NetInd1(CurrConnInd):          HUMCEquClsWk(1) = .NetInd2(CurrConnInd)
               HUMCEquClsCnt = 2                    'always start this type of classes with 2 points
               'build class; we have to go in both direction to discover full connection
               i = HUMCNetNextUnused(CurrConnInd)
               Do While i < HUMCNetCnt
'                   Debug.Assert HUMCNetUsed(i) = HUMCNotUsed
                      'condition in the following two If statements will not be True
                      'simultaneously but this way it will work even if they are
                      If HUMCIsoUsed(.NetInd1(i)) = HUMCInUse Then
                         MarkNetUsed i
                         If HUMCIsoUsed(.NetInd2(i)) = HUMCNotUsed Then     'add it to class if not already there
                            HUMCIsoUsed(.NetInd2(i)) = HUMCInUse
                            HUMCEquClsCnt = HUMCEquClsCnt + 1
                            HUMCEquClsWk(HUMCEquClsCnt - 1) = .NetInd2(i)
                         End If
                      End If
                      If HUMCIsoUsed(.NetInd2(i)) = HUMCInUse Then
                         MarkNetUsed i
                         If HUMCIsoUsed(.NetInd1(i)) = HUMCNotUsed Then     'add it to class if not already there
                            HUMCIsoUsed(.NetInd1(i)) = HUMCInUse
                            HUMCEquClsCnt = HUMCEquClsCnt + 1
                            HUMCEquClsWk(HUMCEquClsCnt - 1) = .NetInd1(i)
                         End If
                      End If
'                   Else
'                       ' If this algorithm is fully optimized, then we shouldn't reach this
'                       Debug.Assert False
'                   End If

                   i = HUMCNetNextUnused(i)
               Loop

               'need to go in another direction to pick up eventual skiping transitions
               i = HUMCNetPrevUnused(HUMCNetCnt)
               Do While i > CurrConnInd
'                   Debug.Assert HUMCNetUsed(i) = HUMCNotUsed
                      If HUMCIsoUsed(.NetInd1(i)) = HUMCInUse Then
                         MarkNetUsed i
                         If HUMCIsoUsed(.NetInd2(i)) = HUMCNotUsed Then     'add it to class if not already there
                            HUMCIsoUsed(.NetInd2(i)) = HUMCInUse
                            HUMCEquClsCnt = HUMCEquClsCnt + 1
                            HUMCEquClsWk(HUMCEquClsCnt - 1) = .NetInd2(i)
                         End If
                      End If
                      If HUMCIsoUsed(.NetInd2(i)) = HUMCInUse Then
                         MarkNetUsed i
                         If HUMCIsoUsed(.NetInd1(i)) = HUMCNotUsed Then     'add it to class if not already there
                            HUMCIsoUsed(.NetInd1(i)) = HUMCInUse
                            HUMCEquClsCnt = HUMCEquClsCnt + 1
                            HUMCEquClsWk(HUMCEquClsCnt - 1) = .NetInd1(i)
                         End If
                      End If
'                   Else
'                       ' If this algorithm is fully optimized, then we shouldn't reach this
'                       Debug.Assert False
'                   End If

                   i = HUMCNetPrevUnused(i)
               Loop

               'now pack findings to nice small array convenient to create classes
               ReDim HUMCEquCls(HUMCEquClsCnt - 1)
               For i = 0 To HUMCEquClsCnt - 1       'make sure not to use more than belongs to this class
                   HUMCEquCls(i) = HUMCEquClsWk(i)
                   HUMCIsoUsed(HUMCEquCls(i)) = HUMCUsed                'they are used now
               Next i
               'extract and add class to the structure
               Call BuildCurrentClass

            End If
            CurrConnInd = HUMCNetNextUnused(CurrConnInd)

            lngNewTickCount = GetTickCount()
            If lngNewTickCount - lngTickCountLastUpdate > 250 Then
                ' Only update 4 times per second
                ChangeStatus "Building Class: " & GelUMC(CallerID).UMCCnt & " (" & Format(CurrConnInd / HUMCNetCnt * 100, "0.00") & "% completed)"
                lngTickCountLastUpdate = lngNewTickCount
                If mAbortProcess Then Exit Do
            End If
         Loop
      End With

End Sub

Private Sub MarkNetUsed(i As Long)
    If HUMCNetUsed(i) = HUMCNotUsed Then
        HUMCNetUsed(i) = HUMCUsed
        If HUMCNetPrevUnused(i) < 0 Then
            HUMCNetNextUnused(0) = HUMCNetNextUnused(i)
        Else
            HUMCNetNextUnused(HUMCNetPrevUnused(i)) = HUMCNetNextUnused(i)
        End If
        HUMCNetPrevUnused(HUMCNetNextUnused(i)) = HUMCNetPrevUnused(i)
    End If
End Sub

