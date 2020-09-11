Attribute VB_Name = "SimplePerceptron"
Option Explicit
Public p_ As Perceptron
'---------------------------------------
Function GetOutput(ByVal x As Variant)
    '-----------------------------------
    If p_ Is Nothing Then Set p_ = New Perceptron
    '-----------------------------------
    GetOutput = p_.Output(x)
    '-----------------------------------
End Function
'---------------------------------------
Function UpdateWeigth(ByVal x As Variant, ByVal s As Double)
    '-----------------------------------
    If p_ Is Nothing Then Set p_ = New Perceptron
    '-----------------------------------
    UpdateWeigth = p_.UpdateWeight(x, s)
    '-----------------------------------
End Function
'---------------------------------------
Sub LearnAndPredict()
    '-----------------------------------
    Dim i As Long
    Dim f As Worksheet
    Dim r As Range
    Dim x As Variant
    Dim w As Variant
    Dim dVal As Double
    '-----------------------------------
    Set f = Excel.ThisWorkbook.Sheets("SP")
    Set r = f.Range("A1")
    '-----------------------------------
    If Not p_ Is Nothing Then
        Set p_ = Nothing
    End If
    '-----------------------------------
    Set p_ = New Perceptron
    Call p_.SetLearningRate(f.Range("Learning_Rate").Value)
    '-----------------------------------
    For i = 4 To 50000
        If r.Offset(i, 0).Value = "" Then GoTo fin:
        x = Array(r.Offset(i - 1, 2).Value, r.Offset(i - 2, 2).Value)
        ' Affichage la prédiction
        r.Offset(i, 3).Value = p_.Output(x)
        r.Offset(i, 4).Value = VBA.Math.Sgn(r.Offset(i, 3).Value - 0.5)
        
        ' Apprentissage
        If r.Offset(i, 2).Value > 0# Then dVal = 1# Else dVal = 0#
        Call p_.UpdateWeight(x, dVal)
        
        ' Affichage des poids
        w = p_.GetWeigth()
        r.Offset(i, 5).Value = w(0)
        r.Offset(i, 6).Value = w(1)
    Next i
fin:
    '-----------------------------------
End Sub
