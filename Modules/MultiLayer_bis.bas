Attribute VB_Name = "MultiLayer_bis"
Option Explicit
Private m_ As MultiPerceptron
'-------------------------
Sub Learn_bis()
    '---------------------
    Dim i As Long
    Dim j As Long
    Dim p As Long
    Dim n As Long
    Dim x As Variant
    Dim s As Double
    Dim k As Long
    Dim l As Long
    '---------------------
    Dim f As Worksheet
    Dim r As Range
    '---------------------
    n = Excel.ThisWorkbook.Sheets("MultiLayer bis").Range("N").Value
    p = Excel.ThisWorkbook.Sheets("MultiLayer bis").Range("P").Value
    '---------------------
    Set m_ = Nothing
    Set m_ = New MultiPerceptron
    '---------------------
    Set f = Excel.ThisWorkbook.Sheets("MultiLayer bis")
    Set r = f.Range("A1")
    '---------------------
    r.Offset(0, 9).EntireColumn.ClearContents
    r.Offset(0, 10).EntireColumn.ClearContents
    r.Offset(0, 11).EntireColumn.ClearContents
    r.Offset(0, 9).Value = "Prediction"
    r.Offset(0, 10).Value = "Sens prediction"
    r.Offset(0, 11).Value = "Prédiction OK"
    '---------------------
    If f.Range("weight").Value = "" Then
        Call m_.init(4 * n, p)
    Else
        If VBA.Information.IsNumeric(f.Range("weight").Value) Then
            Call m_.init(4 * n, p, f.Range("weight").Value)
        Else
            Call m_.init(4 * n, p)
        End If
    End If
    Call m_.SetLearningRate(f.Range("rate").Value)
    Call m_.SetFactor(f.Range("factor").Value)
    '---------------------
    ReDim x(1 To 4 * n)
    '---------------------
    Excel.Application.ScreenUpdating = False
    '---------------------
    For i = 2 + n To 50000
        If r.Offset(i, 0).Value = "" Then GoTo fin:
        
        l = 1
        For k = 1 To 4
            For j = 1 To n
                x(l) = r.Offset(i - j, 2 * k).Value
                l = l + 1
            Next j
        Next k
        Call m_.ComputeZ(x)
        r.Offset(i, 9).Value = m_.Output(x)
        If r.Offset(i, 9).Value > 0.5 Then
            r.Offset(i, 10).Value = 1
        Else
            r.Offset(i, 10).Value = -1
        End If
        
        r.Offset(i, 11).Formula = "=" & r.Offset(i, 10).Address & "*" & r.Offset(i, 2).Address & ">0"
        If r.Offset(i, 2).Value <= 0 Then s = 0# Else s = 1#
        Call m_.UpdateWeight(x, s, f.Range("method").Value)
    Next i
fin:
End Sub
'-------------------------

