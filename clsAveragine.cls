VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAveragine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Code ported by Matthew Monroe from the Averagine.cpp file in
' the Decon2LS application (written by Deep Jaitly and available at http://ncrr.pnl.gov/software/)

Private Type udtAveragineElementInfoType
    Symbol As String        ' Single letter symbol
    Count As Double         ' Empirical formula count
    MonoMass As Double
    AvgMass As Double
End Type

' Tag mass values are currently ignored
Private mUseTagMass As Boolean
Private mTagMass As Double

' Note: recompute mAveragineMassAvg and mAveragineMassMono if changing mAveragineElementInfo
Private mAveragineElementCount As Integer
Private mAveragineElementInfo() As udtAveragineElementInfoType

Private mAveragineMassAvg As Double
Private mAveragineMassMono As Double

Private Sub AddAveragineElement(ByVal strSymbol As String, ByVal dblCount As Double, ByVal dblMonoMass As Double, ByVal dblAvgMass As Double)
    If mAveragineElementCount > UBound(mAveragineElementInfo) Then
        ReDim Preserve mAveragineElementInfo((UBound(mAveragineElementInfo) + 1) * 2 - 1)
    End If

    With mAveragineElementInfo(mAveragineElementCount)
        .Symbol = strSymbol
        .Count = dblCount
        .MonoMass = dblMonoMass
        .AvgMass = dblAvgMass
    End With
    mAveragineElementCount = mAveragineElementCount + 1
    
    UpdateAveragineMasses
    
End Sub

Private Sub ClearAveragineElements(ByVal intSpaceToReserve As Integer)
    mAveragineElementCount = 0
    
    If intSpaceToReserve < 1 Then intSpaceToReserve = 1
    ReDim mAveragineElementInfo(intSpaceToReserve - 1)
End Sub

Public Function GetAveragineFormulaForMass(ByVal dblMonoMassTarget As Double, _
            Optional ByRef lngCountC As Long, _
            Optional ByRef lngCountH As Long, _
            Optional ByRef lngCountN As Long, _
            Optional ByRef lngCountO As Long, _
            Optional ByRef lngCountS As Long, _
            Optional ByRef dblMonoMassComputed As Double, _
            Optional ByRef dblAvgMassComputed As Double) As String

    ' Uses dblMonoMassTarget to compute the average empirical formula given the values in mAveragineElementInfo
    ' The function returns the empirical formula
    ' The ByRef variables return the atom counts for C, H, N, O, and S, plus the
    '  computed monoisotopic and average masses
    
    Dim blnUsingTagMass As Boolean
    
    Dim intElementIndex As Integer
    Dim intHydrogenIndex As Integer
    Dim intHydrogenInsertPos As Integer
    
    Dim lngTotalAtomCount As Long
    Dim lngAtomCountComputedCurrent As Long
    
    Dim dblRemainderMass As Double
    Dim dblHydrogenMassMono As Double
    Dim dblHydrogenMassAvg As Double
    
    Dim strEmpiricalFormula As String
    Dim strFormulaNew As String
    
    
    blnUsingTagMass = False
''    If mUseTagMass And dblMonoMassTarget > mTagMass Then
''        ' Subtract out the tag mass
''        dblMonoMassTarget = dblMonoMassTarget - mTagMass
''        blnUsingTagMass = True
''    Else
''        blnUsingTagMass = False
''    End If

    intHydrogenIndex = -1
    intHydrogenInsertPos = -1
    
    dblMonoMassComputed = 0
    dblAvgMassComputed = 0
    strEmpiricalFormula = ""
    lngCountC = 0
    lngCountH = 0
    lngCountN = 0
    lngCountO = 0
    lngCountS = 0

    For intElementIndex = 0 To mAveragineElementCount - 1
        ' Note: Int() always rounds down, and is thus equivalent to Floor()
        lngAtomCountComputedCurrent = Int(dblMonoMassTarget / mAveragineMassAvg * mAveragineElementInfo(intElementIndex).Count + 0.5)
        
        Select Case UCase(mAveragineElementInfo(intElementIndex).Symbol)
        Case "C"
            lngCountC = lngAtomCountComputedCurrent
        Case "H"
            lngCountH = lngAtomCountComputedCurrent
            
            ' Hydrogen is used as the remainder mass, so we skip it in the formula calculation by assigning it to 0
            ' for now and later below assigning the remainder to it.
            intHydrogenIndex = intElementIndex
            lngAtomCountComputedCurrent = 0
            
            intHydrogenInsertPos = Len(strEmpiricalFormula) + 1
        Case "N"
            lngCountN = lngAtomCountComputedCurrent
        Case "O"
            lngCountO = lngAtomCountComputedCurrent
        Case "S"
            lngCountS = lngAtomCountComputedCurrent
        Case Else
        End Select
        
        If lngAtomCountComputedCurrent > 0 Then
            dblAvgMassComputed = dblAvgMassComputed + mAveragineElementInfo(intElementIndex).AvgMass * lngAtomCountComputedCurrent
            dblMonoMassComputed = dblMonoMassComputed + mAveragineElementInfo(intElementIndex).MonoMass * lngAtomCountComputedCurrent
        
            lngTotalAtomCount = lngTotalAtomCount + lngAtomCountComputedCurrent
            
            If Len(strEmpiricalFormula) > 0 Then strEmpiricalFormula = strEmpiricalFormula & " "
            strEmpiricalFormula = strEmpiricalFormula + mAveragineElementInfo(intElementIndex).Symbol & LTrim(RTrim(CStr(lngAtomCountComputedCurrent)))
        End If
        
    Next intElementIndex
    
    If blnUsingTagMass Then
        ' This code is not currently implemented
''        ' Tag mass is enabled, so add it to the formula.
''        int numElementsTag = mobj_tag_formula.mint_num_elements ;
''        double totalElementCountTag = mobj_tag_formula.mdbl_total_atom_count ;
''        for (int elementNum = 0 ; elementNum < numElementsTag ; elementNum++)
''        {
''            AtomicCount elementCount = mobj_tag_formula.mvect_elemental_composition[elementNum] ;
''            int numAtoms = (int) elementCount.mdbl_num_copies  ;
''            double elementAvgMass = mobj_element_isotope_abundance.mvect_elemental_isotopes[elementCount.mint_index].mdbl_average_mass ;
''            double elementMonoMass = mobj_element_isotope_abundance.mvect_elemental_isotopes[elementCount.mint_index].marr_isotope_mass[0] ;
''            averageMass += elementAvgMass * numAtoms ;
''            monoMass += elementMonoMass * numAtoms ;
''            totalAtomCount += numAtoms ;
''            elementCount.mdbl_num_copies = numAtoms ;
''            // now go through each elementCount. If already added, then just add to its count.
''            // otherwise, create new.
''            int elementNumAlreadyPushed = -1 ;
''            for (int elementNumPushed = 0 ; elementNumPushed < (int) empirical_formula.mvect_elemental_composition.size() ;
''                elementNumPushed++)
''            {
''                if (empirical_formula.mvect_elemental_composition[elementNumPushed].mint_index == elementCount.mint_index)
''                {
''                    elementNumAlreadyPushed = elementNumPushed ;
''                    break ;
''                }
''            }
''            if (elementNumAlreadyPushed == -1)
''                empirical_formula.mvect_elemental_composition.push_back(elementCount) ;
''            Else
''                empirical_formula.mvect_elemental_composition[elementNumAlreadyPushed].mdbl_num_copies += elementCount.mdbl_num_copies ;
''        }
    End If
    
    ' Now whatever's left over in mass, is assigned to hydrogen because it is not expected to cause
    ' much of a distortion in the isotope profile.
     
    dblRemainderMass = dblMonoMassTarget - dblAvgMassComputed
    If blnUsingTagMass Then
        dblRemainderMass = dblRemainderMass + mTagMass
    End If
    
    If intHydrogenIndex >= 0 Then
        dblHydrogenMassMono = mAveragineElementInfo(intHydrogenIndex).MonoMass
        dblHydrogenMassAvg = mAveragineElementInfo(intHydrogenIndex).AvgMass
    Else
        dblHydrogenMassMono = 1.00782
        dblHydrogenMassAvg = 1.00794
    End If

    lngCountH = Int(dblRemainderMass / dblHydrogenMassAvg + 0.5)

    If lngCountH > 0 Then
        dblAvgMassComputed = dblAvgMassComputed + dblHydrogenMassAvg * lngCountH
        dblMonoMassComputed = dblMonoMassComputed + dblHydrogenMassMono * lngCountH
    
        lngTotalAtomCount = lngTotalAtomCount + lngCountH
        
        ' Add "H" to the empirical formula, inserting at position intHydrogenInsertPos
        If intHydrogenInsertPos < 0 Then
            If Len(strEmpiricalFormula) > 0 Then strEmpiricalFormula = strEmpiricalFormula & " "
            strEmpiricalFormula = strEmpiricalFormula + "H" & LTrim(RTrim(CStr(lngCountH)))
        Else
            If intHydrogenInsertPos > 1 Then
                strFormulaNew = RTrim(Left(strEmpiricalFormula, intHydrogenInsertPos - 1)) & " "
            Else
                strFormulaNew = ""
            End If
            
            strFormulaNew = strFormulaNew + "H" & LTrim(RTrim(CStr(lngCountH)))
            
            If intHydrogenInsertPos < Len(strEmpiricalFormula) Then
                If Len(strFormulaNew) > 0 Then strFormulaNew = strFormulaNew & " "
                strFormulaNew = strFormulaNew + LTrim(Mid(strEmpiricalFormula, intHydrogenInsertPos))
            End If
            strEmpiricalFormula = CStr(strFormulaNew)
        End If
    Else
        ' lngCountH is 0; this is unexpected
        Debug.Assert False
    End If
    
    GetAveragineFormulaForMass = strEmpiricalFormula
End Function

Private Sub InitializeDefaults()
    SetProteinAveragineMasses
    
    mUseTagMass = False
    mTagMass = 0
End Sub

Public Sub SetDNAAveragineMasses()
    ClearAveragineElements 5
    
    AddAveragineElement "C", 3.9, 12, 12.0107
    AddAveragineElement "H", 4.9, 1.00782, 1.00794
    AddAveragineElement "N", 1.5, 14.00307, 14.00674
    AddAveragineElement "O", 2.4, 15.99492, 15.9994
    AddAveragineElement "P", 0.4, 30.97376, 30.97376

End Sub

Public Sub SetProteinAveragineMasses()
    ClearAveragineElements 5
    
    AddAveragineElement "C", 4.9384, 12, 12.0107
    AddAveragineElement "H", 7.7583, 1.00782, 1.00794
    AddAveragineElement "N", 1.3577, 14.00307, 14.00674
    AddAveragineElement "O", 1.4773, 15.99492, 15.9994
    AddAveragineElement "S", 0.0417, 31.97207, 32.066

End Sub

Public Sub SetRNAAveragineMasses()
    ClearAveragineElements 5
    
    AddAveragineElement "C", 3.8, 12, 12.0107
    AddAveragineElement "H", 4.7, 1.00782, 1.00794
    AddAveragineElement "N", 1.5, 14.00307, 14.00674
    AddAveragineElement "O", 2.8, 15.99492, 15.9994
    AddAveragineElement "P", 0.4, 30.97376, 30.97376

End Sub

Private Sub UpdateAveragineMasses()
    Dim intIndex As Integer
    
    mAveragineMassAvg = 0
    mAveragineMassMono = 0
    
    For intIndex = 0 To mAveragineElementCount - 1
        With mAveragineElementInfo(intIndex)
            mAveragineMassAvg = mAveragineMassAvg + .Count * .AvgMass
            mAveragineMassMono = mAveragineMassMono + .Count * .MonoMass
        End With
        
    Next intIndex
End Sub

Private Sub Class_Initialize()
    InitializeDefaults
End Sub
