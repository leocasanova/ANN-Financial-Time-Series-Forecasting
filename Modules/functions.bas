Attribute VB_Name = "functions"
Option Explicit
'----------------------
Function GetDim(ByVal v As Variant)
    '------------------
    On Error GoTo fin:
    '------------------
    Dim n As Long
    Dim k As Long
    '------------------
    n = 0
    '------------------
    While True
        k = UBound(v, n + 1)
        n = n + 1
    Wend
    '------------------
fin:
    '------------------
    GetDim = n
    '------------------
End Function
'----------------------
Function GetNbElement(ByVal v As Variant)
    '------------------
    Dim n As Long
    Dim vTmp As Variant
    '------------------
    If VBA.Information.IsArray(v) Then
        n = 0
        For Each vTmp In v
            n = n + 1
        Next vTmp
    Else
        n = 1
    End If
    '------------------
    GetNbElement = n
    '------------------
End Function
'----------------------
Function Sigmoid(ByVal x As Double)
    Dim un As Double
    un = VBA.Conversion.CDbl(1)
    If x <= -300 Then
        Sigmoid = 0#
    ElseIf x >= 300 Then
        Sigmoid = un
    Else
        Sigmoid = (un + VBA.Math.Exp(-x)) ^ (-un)
    End If
End Function
'----------------------
Function d_Sigmoid(ByVal x As Double)
    d_Sigmoid = Sigmoid(x) * (VBA.Conversion.CDbl(1) - Sigmoid(x))
End Function
'----------------------
Function SplitVariant(ByVal v As Variant)
    Dim vRes As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim n As Long
    Dim m As Long
    If GetDim(v) = 1 Then
        If UBound(v, 1) - LBound(v, 1) <= 1 Then
            ReDim vRes(0 To 0, 0 To UBound(v, 1) - LBound(v, 1))
            j = 0
            For i = LBound(v, 1) To UBound(v, 1)
                vRes(0, j) = v(i)
                j = j + 1
            Next i
        Else
            n = UBound(v, 1) - LBound(v, 1) + 1
            m = Excel.WorksheetFunction.Fact(n) / (Excel.WorksheetFunction.Fact(n - 2) * 2)
            ReDim vRes(0 To m - 1, 0 To 1)
            For i = LBound(v, 1) To UBound(v, 1)
                For j = i + 1 To UBound(v, 1)
                    If i <> j Then
                        vRes(k, 0) = v(i)
                        vRes(k, 1) = v(j)
                        k = k + 1
                    End If
                Next j
            Next i
        End If
    Else
        vRes = ""
    End If
    SplitVariant = vRes
End Function
'----------------------
Sub tester()
    Dim v As Variant
    Dim i As Long
    Dim w As Variant
    
    ReDim v(0 To 3)
    
    For i = 0 To UBound(v, 1)
        v(i) = VBA.Math.Rnd()
    Next i
    w = SplitVariant(v)
    
End Sub
'----------------------
Function GetCombi(ByVal n As Long)
    GetCombi = Excel.WorksheetFunction.Fact(n) / (Excel.WorksheetFunction.Fact(n - 2) * 2)
End Function
'----------------------
