Attribute VB_Name = "MultiLayer"
Option Explicit
Private m_ As MultiPerceptron
'-------------------------
Sub Learn()
    '---------------------
    Dim i As Long
    Dim j As Long
    Dim p As Long
    Dim n As Long
    Dim x As Variant
    Dim s As Double
    '---------------------
    Dim f As Worksheet
    Dim r As Range
    '---------------------
    n = Excel.ThisWorkbook.Sheets("MultiLayer").Range("N_Layer").Value
    p = Excel.ThisWorkbook.Sheets("MultiLayer").Range("P_Layer").Value
    '---------------------
    Set m_ = Nothing
    Set m_ = New MultiPerceptron
    '---------------------
    Set f = Excel.ThisWorkbook.Sheets("MultiLayer")
    Set r = f.Range("A1")
    '---------------------
    r.Offset(0, 3).EntireColumn.ClearContents
    r.Offset(0, 4).EntireColumn.ClearContents
    r.Offset(0, 5).EntireColumn.ClearContents
    r.Offset(0, 3).Value = "Prediction"
    r.Offset(0, 4).Value = "Sens prediction"
    r.Offset(0, 5).Value = "Prédiction OK"
    '---------------------
    Call m_.init(n, p)
    Call m_.SetLearningRate(f.Range("LearningRate").Value)
    '---------------------
    ReDim x(1 To n)
    '---------------------
    Excel.Application.ScreenUpdating = False
    '---------------------
    For i = 2 + n To 50000
        If r.Offset(i, 0).Value = "" Then GoTo fin:
        For j = 1 To n
            x(j) = r.Offset(i - j, 2).Value
        Next j
        Call m_.ComputeZ(x)
        r.Offset(i, 3).Value = m_.Output(x)
        If r.Offset(i, 3).Value > 0.5 Then
            r.Offset(i, 4).Value = 1
        Else
            r.Offset(i, 4).Value = -1
        End If
        
        r.Offset(i, 5).Formula = "=" & r.Offset(i, 2).Address & "*" & r.Offset(i, 4).Address & ">0"
        If r.Offset(i, 2).Value <= 0 Then s = 0# Else s = 1#
        Call m_.UpdateWeight(x, s, f.Range("method").Value)
    Next i
fin:
End Sub
'-------------------------
