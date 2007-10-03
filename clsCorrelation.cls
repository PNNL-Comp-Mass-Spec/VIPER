VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCorrelation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This class can be used to correlate two lists of numbers (typically mass spectra) to determine their similarity
' The lists of numbers must have the same number of values
'
' These functions were originally written in VB6 and required the use of a C DLL
' They were then ported to VB.NET in 2004
' Synched the VB 6 version with the VB.NET version in February 2005
'
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
' Started October 24, 2003
'
' Last modified February 14, 2005


Private Const MIN_NON_ZERO_ION_COUNT As Long = 3

Public Enum cmCorrelationMethodConstants
    Pearson = 0
    Spearman = 1
    Kendall = 2
End Enum

Private mCoefficients() As Double

Private mErrorMessage As String
Private objQSSingle As QSSingleWithSingle
'

Public Property Get ErrorMessage() As String
    ErrorMessage = mErrorMessage
End Property

Private Function BetaCF(ByVal A As Double, ByVal B As Double, ByVal x As Double) As Double

    Const MAXIT As Long = 100
    Const EPS As Double = 0.0000003
    Const FPMIN As Double = 1E-30

    Dim m As Long, m2 As Long
    Dim aa As Double, c As Double, d As Double, del As Double
    Dim h As Double, qab As Double, qam As Double, qap As Double

    qab = A + B
    qap = A + 1#
    qam = A - 1#
    c = 1#
    d = 1# - qab * x / qap
    If (Abs(d) < FPMIN) Then d = FPMIN
    d = 1# / d
    h = d
    For m = 1 To MAXIT
        m2 = 2 * m
        aa = m * (B - m) * x / ((qam + m2) * (A + m2))
        d = 1# + aa * d
        If (Abs(d) < FPMIN) Then d = FPMIN
        c = 1# + aa / c
        If (Abs(c) < FPMIN) Then c = FPMIN
        d = 1# / d
        h = h * d * c
        aa = -(A + m) * (qab + m) * x / ((A + m2) * (qap + m2))
        d = 1# + aa * d
        If (Abs(d) < FPMIN) Then d = FPMIN
        c = 1# + aa / c
        If (Abs(c) < FPMIN) Then c = FPMIN
        d = 1# / d
        del = d * c
        h = h * del
        If (Abs(del - 1#) < EPS) Then Exit For
    Next m

    If (m > MAXIT) Then
        mErrorMessage = "a or b too big, or MAXIT too small in clsCorrelation->BetaCF"
        Debug.Assert False
        BetaCF = 0
    Else
        BetaCF = h
    End If

End Function

Private Function BetaI(ByVal A As Double, ByVal B As Double, ByVal x As Double) As Double

    Dim bt As Double

    If (x < 0# Or x > 1#) Then
        mErrorMessage = "Bad x in routine clsCorrelation->BetaI; should be between 0 and 1"
    Else

        If (x = 0# Or x = 1#) Then
            bt = 0#
        Else
            bt = Exp(GammLn(A + B) - GammLn(A) - GammLn(B) + A * Log(x) + B * Log(1# - x))
        End If

        If (x < (A + 1#) / (A + B + 2#)) Then
            BetaI = bt * BetaCF(A, B, x) / A
        Else
            BetaI = 1# - bt * BetaCF(B, A, 1# - x) / B
        End If
    End If

End Function

Public Function Correlate(ByRef sngDataList1() As Single, ByRef sngDataList2() As Single, ByVal eCorrelationMethod As cmCorrelationMethodConstants) As Single
    ' Finds the correlation value between the two lists of data
    ' The lists must have the same number of data points, and are assumed to be zero-based arrays, ranging from 0 to lngDataCount-1
    ' If they have fewer than MIN_NON_ZERO_ION_COUNT non-zero values, then the correlation value returned will be 0
    '
    ' Returns correlation value (0 to 1)
    ' If an error, returns -1
    '
    ' Note: We're passing the Data Lists ByRef for performance reasons; they are not modified by this function

    Dim lngIndex As Long
    Dim lngDataCount As Long
    Dim lngNonZeroDataCount As Long

    Dim RValue As Single, ProbOfSignificance As Single, FishersZ As Single
    Dim sngCorrelationValueToReturn As Single
    Dim DiffInRanks As Single, ZD As Single, RS As Single, ProbRS As Single
    Dim KendallsTau As Single, z As Single

    ''        Dim sngDataList1Test() As Single = New Single() {1, 2, 2, 8, 9, 0, 0, 3, 9, 0, 5, 6}
    ''        Dim sngDataList2Test() As Single = New Single() {2, 3, 7, 7, 11, 1, 3, 2, 13, 0, 4, 10}

    On Error GoTo CorrelateErrorHandler

mErrorMessage = ""

    RValue = 0
    RS = 0
    KendallsTau = 0

    lngDataCount = UBound(sngDataList1) + 1
    If lngDataCount <> UBound(sngDataList2) + 1 Then
        Correlate = -1
        Exit Function
    End If

    ' Determine the number of non-zero data points in the two lists
    lngNonZeroDataCount = 0
    For lngIndex = 0 To lngDataCount - 1
        If sngDataList1(lngIndex) > 0 Then lngNonZeroDataCount = lngNonZeroDataCount + 1
    Next lngIndex
    If lngNonZeroDataCount < MIN_NON_ZERO_ION_COUNT Then
        Correlate = 0
        Exit Function
    End If

    lngNonZeroDataCount = 0
    For lngIndex = 0 To lngDataCount - 1
        If sngDataList2(lngIndex) > 0 Then
            lngNonZeroDataCount = lngNonZeroDataCount + 1
        End If
    Next lngIndex
    
    If lngNonZeroDataCount < MIN_NON_ZERO_ION_COUNT Then
        sngCorrelationValueToReturn = 0
    Else
        Select Case eCorrelationMethod
        Case cmCorrelationMethodConstants.Pearson
            CorrelPearson sngDataList1, sngDataList2, RValue, ProbOfSignificance, FishersZ
            sngCorrelationValueToReturn = RValue
        Case cmCorrelationMethodConstants.Spearman
            CorrelSpearman sngDataList1, sngDataList2, DiffInRanks, ZD, ProbOfSignificance, RS, ProbRS
            sngCorrelationValueToReturn = RS
        Case cmCorrelationMethodConstants.Kendall
            CorrelKendall sngDataList1, sngDataList2, KendallsTau, z, ProbOfSignificance
            sngCorrelationValueToReturn = KendallsTau
        End Select
    End If
    
    Correlate = sngCorrelationValueToReturn
Exit Function

CorrelateErrorHandler:
    Debug.Assert False
    Correlate = -1

End Function

Private Sub CorrelPearson(ByRef sngDataList1() As Single, ByRef sngDataList2() As Single, ByRef RValue As Single, ByRef ProbOfSignificance As Single, ByRef FishersZ As Single)
    ' Performs a Pearson correlation (aka linear correlation) of the two lists
    ' The lists must have the same number of data points in each and should be 0-based arrays
    '
    ' Code from Numerical Recipes in C

    ' Note: We're passing the Data Lists ByRef for performance reasons; they are not modified by this function

    '  TINY is used to "regularize" the unusual case of complete correlation
    Const TINY As Double = 1E-20

    ' Given two arrays x[1..n] and y[1..n], this routine computes their correlation coeffcient
    ' r (returned as r), the signicance level at which the null hypothesis of zero correlation is
    ' disproved (prob whose small value indicates a significant correlation), and Fisher's z (returned
    ' as z), whose value can be used in further statistical tests as described above.

    Dim n As Long
    Dim j As Long
    Dim yt As Double, xt As Double, t As Double, df As Double
    Dim syy As Double
    Dim sxy As Double
    Dim sxx As Double
    Dim ay As Double
    Dim ax As Double

    syy = 0
    sxy = 0
    sxx = 0
    ay = 0
    ax = 0

    RValue = 0
    ProbOfSignificance = 0
    FishersZ = 0

    n = UBound(sngDataList1) + 1
    If n <> UBound(sngDataList2) + 1 Then
        mErrorMessage = "sngDataList1 and sngDataList2 must be arrays of the same length"
        Debug.Assert False
        n = 0
    End If
    If n <= 0 Then Exit Sub

    ' Find the means
    For j = 0 To n - 1
        ax = ax + sngDataList1(j)
        ay = ay + sngDataList2(j)
    Next j
    ax = ax / n
    ay = ay / n

    ' Compute the correlation coefficient
    For j = 0 To n - 1
        xt = sngDataList1(j) - ax
        yt = sngDataList2(j) - ay
        sxx = sxx + xt * xt
        syy = syy + yt * yt
        sxy = sxy + xt * yt
    Next j

    RValue = CSng(sxy / (Sqr(sxx * syy) + TINY))

    ' Fisher's z transformation
    FishersZ = CSng(0.5 * Log((1# + RValue + TINY) / (1# - RValue + TINY)))
    df = n - 2

    t = RValue * Sqr(df / ((1# - RValue + TINY) * (1# + RValue + TINY)))

    ' Student's t probability
    ProbOfSignificance = CSng(BetaI(0.5 * df, 0.5, df / (df + t * t)))

End Sub

Private Sub CorrelKendall(ByRef sngDataList1() As Single, ByRef sngDataList2() As Single, ByRef KendallsTau As Single, ByRef z As Single, ByRef ProbOfSignificance As Single)
    ' Performs a Kendall correlation (aka linear correlation) of the two lists
    ' The lists must have the same number of data points in each and should be 0-based arrays
    '
    ' Code from Numerical Recipes in C

    ' Note: We're passing the Data Lists ByRef for performance reasons; they are not modified by this function

    ' Given data arrays data1[1..n] and data2[1..n], this program returns Kendall's tau as tau,
    ' its number of standard deviations from zero as z, and its two-sided significance level as prob.
    ' Small values of prob indicate a significant correlation (tau positive) or anticorrelation (tau
    ' negative).

    Dim n As Long
    Dim n2 As Long
    Dim n1 As Long
    Dim k As Long, j As Long
    Dim lngIS As Long

    Dim svar As Double, aa As Double, a2 As Double, a1 As Double

    KendallsTau = 0
    z = 0
    ProbOfSignificance = 0

    n = UBound(sngDataList1) + 1
    If n <> UBound(sngDataList2) + 1 Then
        mErrorMessage = "sngDataList1 and sngDataList2 must be arrays of the same length"
        n = 0
    End If
    If n <= 0 Then Exit Sub

    For j = 0 To n - 2
        For k = j + 1 To n - 1
            a1 = sngDataList1(j) - sngDataList1(k)
            a2 = sngDataList2(j) - sngDataList2(k)
            aa = a1 * a2
            If aa <> 0 Then
                n1 = n1 + 1
                n2 = n2 + 1
                If aa > 0 Then
                    lngIS = lngIS + 1
                Else
                    lngIS = lngIS - 1
                End If
            Else
                If a1 <> 0 Then n1 = n1 + 1
                If a2 <> 0 Then n2 = n2 + 1
            End If
        Next k
    Next j

    KendallsTau = CSng(lngIS / (Sqr(n1) * Sqr(n2)))

    svar = (4# * n + 10#) / (9# * n * (n - 1#))
    z = CSng(KendallsTau / Sqr(svar))
    ProbOfSignificance = CSng(ErfCC(Abs(z) / 1.4142136))

End Sub

Private Sub CorrelSpearman(ByRef sngDataList1() As Single, ByRef sngDataList2() As Single, ByRef DiffInRanks As Single, ByRef ZD As Single, ByRef ProbOfSignificance As Single, ByRef RS As Single, ByRef ProbRS As Single)
    ' Performs a Spearman correlation of the two lists
    ' The lists must have the same number of data points in each and should be 0-based arrays
    '
    ' Code from Numerical Recipes in C

    ' Note: sngDataList1 and sngDataList2 are re-ordered by this function
    '    Since VB6 won't let us pass them ByVal, we must copy them to working arrays before processing

    ' Given two data arrays, data1[0..n-1] and data2[0..n-1], this routine returns their sum-squared
    ' difference of ranks as D, the number of standard deviations by which D deviates from its null hypothesis
    ' expected value as zd, the two-sided significance level of this deviation as probd,
    ' Spearman's rank correlation rs as rs, and the two-sided significance level of its deviation from
    ' zero as probrs. The external routine CRank is used.  A small value of either probd or probrs indicates
    ' a significant correlation (rs positive) or anticorrelation (rs negative).

    Dim n As Long
    Dim j As Long

    Dim sg As Single, sf As Single
    Dim vard As Double, t As Double
    Dim fac As Double, en3n As Double
    Dim en As Double, df As Double, AvgD As Double
    Dim DiffInRanksWork As Double
    
    Dim sngDataListWork1() As Single, sngDataListWork2() As Single
    
    DiffInRanks = 0
    ZD = 0
    ProbOfSignificance = 0
    RS = 0
    ProbRS = 0

    n = UBound(sngDataList1) + 1
    If n <> UBound(sngDataList2) + 1 Then
        mErrorMessage = "sngDataList1 and sngDataList2 must be arrays of the same length"
        n = 0
    End If
    If n <= 0 Then Exit Sub

    ' Copy data from sngDataList to sngDataListWork
    ReDim sngDataListWork1(n - 1)
    ReDim sngDataListWork2(n - 1)
    For j = 0 To n - 1
        sngDataListWork1(j) = sngDataList1(j)
        sngDataListWork2(j) = sngDataList2(j)
    Next j
    
    ' Sort sngDataListWork1, sorting sngDataListWork2 parallel to it
    If objQSSingle Is Nothing Then Set objQSSingle = New QSSingleWithSingle
    objQSSingle.QSAsc sngDataListWork1(), sngDataListWork2()
    CRank n, sngDataListWork1, sf

    ' Sort sngDataListWork2, sorting sngDataListWork1 parallel to it
    objQSSingle.QSAsc sngDataListWork2(), sngDataListWork1()
    CRank n, sngDataListWork2, sg

    DiffInRanksWork = 0#
    For j = 0 To n - 1
        DiffInRanksWork = DiffInRanksWork + SquareNum(sngDataListWork1(j) - sngDataListWork2(j))
    Next j
    DiffInRanks = CSng(DiffInRanksWork)

    en = n

    en3n = en * en * en - en
    AvgD = en3n / 6# - (sf + sg) / 12#
    fac = (1# - sf / en3n) * (1# - sg / en3n)
    vard = ((en - 1#) * en * en * SquareNum(en + 1#) / 36#) * fac
    ZD = CSng((DiffInRanks - AvgD) / Sqr(vard))

    ProbOfSignificance = CSng(ErfCC(Abs(ZD) / 1.4142136))
    RS = CSng((1# - (6# / en3n) * (DiffInRanks + (sf + sg) / 12#)) / Sqr(fac))

    fac = (RS + 1#) * (1# - RS)

    If (fac > 0#) Then
        t = RS * Sqr((en - 2#) / fac)
        df = en - 2#
        ProbRS = CSng(BetaI(0.5 * df, 0.5, df / (df + t * t)))
    Else
        ProbRS = 0#
    End If

End Sub

Private Sub CRank(ByVal n As Long, ByRef w() As Single, ByRef S As Single)

    ' Given a zero-based sorted array w(0..n-1), replaces the elements by their rank (1 .. n), including midranking of ties,
    ' and returns as s the sum of f^3 - f, where f is the number of elements in each tie.

    Dim j As Long
    Dim ji As Long, jt As Long
    Dim t As Single, rank As Single

    S = 0
    j = 0
    Do While j < n - 1
        If w(j + 1) <> w(j) Then
            w(j) = j + 1            ' Rank = j + 1
            j = j + 1
        Else
            jt = j + 1
            Do While jt < n
                If w(jt) = w(j) Then
                    jt = jt + 1
                Else
                    Exit Do
                End If
            Loop
            rank = 0.5! * (j + jt - 1) + 1

            For ji = j To jt - 1
                w(ji) = rank
            Next ji

            t = jt - j
            S = S + t * t * t - t        ' t^3 - t
            j = jt
        End If
    Loop

    If j = n - 1 Then
        w(n - 1) = n
    End If

End Sub

Private Function ErfCC(ByVal x As Double) As Double

    Dim t As Double, z As Double, ans As Double
    Dim dblTemp As Double
    
    z = Abs(x)
    t = 1# / (1# + 0.5 * z)

    ' We need to perform the following in multiple steps since VB6 gripes that the full expression is too complex (more than 8 levels of computation)
    'ans = t * Exp(-z * z - 1.26551223 + t * (1.00002368 + t * (0.37409196 + t * (0.09678418 + t * (-0.18628806 + t * (0.27886807 + t * (-1.13520398 + t * (1.48851587 + t * (-0.82215223 + t * 0.17087277)))))))))
    
    dblTemp = 1.48851587 + t * (-0.82215223 + t * 0.17087277)
    dblTemp = -0.18628806 + t * (0.27886807 + t * (-1.13520398 + t * dblTemp))
    dblTemp = t * (0.37409196 + t * (0.09678418 + t * dblTemp))
    ans = t * Exp(-z * z - 1.26551223 + t * (1.00002368 + dblTemp))
    

    If x >= 0# Then
        ErfCC = ans
    Else
        ErfCC = 2# - ans
    End If

End Function

Private Function GammLn(ByVal xx As Double) As Double
    Dim x As Double, y As Double, tmp As Double, ser As Double
    Dim j As Long

    x = xx
    y = x

    tmp = x + 5.5
    tmp = tmp - (x + 0.5) * Log(tmp)
    ser = 1.00000000019001

    For j = 0 To 5
        y = y + 1
        ser = ser + mCoefficients(j) / y
    Next j

    GammLn = -tmp + Log(2.50662827463101 * ser / x)

End Function

Private Function SquareNum(ByVal dblNum As Double) As Double

    If dblNum = 0 Then
        SquareNum = 0
    Else
        SquareNum = dblNum * dblNum
    End If

End Function

Private Sub Class_Initialize()
    ReDim mCoefficients(0 To 5)
    
    mCoefficients(0) = 76.1800917294715
    mCoefficients(1) = -86.5053203294168
    mCoefficients(2) = 24.0140982408309
    mCoefficients(3) = -1.23173957245015
    mCoefficients(4) = 1.20865097386618E-03
    mCoefficients(5) = -5.395239384953E-06

End Sub
