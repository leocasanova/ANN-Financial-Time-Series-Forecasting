Attribute VB_Name = "Compare"
Option Explicit

Sub CompoundResultats()
    Dim f As Worksheet
    Dim r As Range
    Dim i As Long
    Set f = Excel.ThisWorkbook.Sheets("Calculs")
    Set r = f.Range("refParams")
    
    For i = 1 To 500
        If r.Offset(i, 0).Value = "" Then GoTo fin:
        Call locCompound(r.Offset(i, 0).Value, r.Offset(i, 1).Value, r.Offset(i, 2).Value, r.Offset(i, 3).Value, f.Range("StartCalculationDt").Value, f.Range("EndCalculationDt").Value, r.Offset(i, 4))
    Next i
fin:
End Sub

Private Sub locCompound(ByVal sName As String, ByVal sDates As String, ByVal sReturns As String, ByVal sPredictions As String, ByVal lDateStart As Long, ByVal lDateEnd As Long, ByVal r As Range)
    Dim f As Worksheet
    Dim rD As Range
    Dim rR As Range
    Dim rP As Range
    Dim i As Long
    Dim dPnL As Double
    Dim dSuccessRate As Double
    Dim dTotal As Double
    
    On Error GoTo fin:
    Set f = Excel.ThisWorkbook.Sheets(sName)
    Set rD = f.Range(sDates)
    Set rR = f.Range(sReturns)
    Set rP = f.Range(sPredictions)
    
    dPnL = 100
    
    For i = 1 To 50000
        If rD.Offset(i, 0).Value = "" Then GoTo fin:
        If rD.Offset(i, 0).Value >= lDateStart And rD.Offset(i, 0).Value <= lDateEnd Then
            If rP.Offset(i, 0).Value * rR.Offset(i, 0).Value > 0 Then dSuccessRate = dSuccessRate + 1
            dTotal = dTotal + 1
            If rP.Offset(i, 0).Value > 0 Then
                dPnL = dPnL * (1# + rR.Offset(i, 0).Value)
            Else
                dPnL = dPnL * (1# - rR.Offset(i, 0).Value)
            End If
        End If
    Next i
fin:
    r.Value = dPnL
    r.Offset(0, 1).Value = dSuccessRate / dTotal
End Sub


