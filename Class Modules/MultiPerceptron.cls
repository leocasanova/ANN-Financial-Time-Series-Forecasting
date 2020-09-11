VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MultiPerceptron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private n_ As Long
Private p_ As Long
Private w_ As Variant
Private z_ As Variant
Private e_ As Variant

Private wp_ As Variant
Private dW_ As Variant
Private dWp_ As Variant
Private learning_ As Double
Private lambda_ As Double
Private factor_ As Double
'-------------------------------------------
Sub SetLearningRate(ByVal dRate As Double)
    learning_ = dRate
End Sub
'-------------------------------------------
Sub init(ByVal n As Long, ByVal p As Long, Optional ByVal dWeight = -1)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    n_ = n
    p_ = p
    ReDim wp_(1 To n_)
    ReDim w_(1 To n_, 1 To n_, 0 To p_)
    For k = 0 To p_
        For i = 1 To n_
            If dWeight < 0 Then
                wp_(i) = VBA.Math.Rnd
            Else
                wp_(i) = dWeight
            End If
            For j = 1 To n_
                If dWeight < 0 Then
                    w_(i, j, k) = VBA.Math.Rnd
                Else
                    w_(i, j, k) = dWeight
                End If
            Next j
        Next i
    Next k
    
    lambda_ = VBA.Math.Rnd()
End Sub
'-------------------------------------------
Sub ComputeZ(ByVal x As Variant)
    ReDim z_(1 To n_, 0 To p_)
    Dim i As Long
    Dim k As Long
    Dim j As Long
    
    Dim dZ As Double
    For i = 1 To n_
        z_(i, 0) = x(LBound(x) + i - 1)
    Next i
    For k = 1 To p_
        For i = 1 To n_
            dZ = 0
            For j = 1 To n_
                dZ = dZ + w_(j, i, k - 1) * z_(j, k - 1)
            Next j
            z_(i, k) = functions.Sigmoid(dZ)
        Next i
    Next k
End Sub
'-------------------------------------------
Sub ComputeError(ByVal x As Variant, ByVal s As Double)
    Dim y As Double
    Dim z As Double
    Dim e As Double
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim dE As Double
    y = Output(x)
    e = (y - s) * y * (VBA.Conversion.CDbl(1) - y)
    ReDim e_(1 To n_, 1 To p_ + 1)
    
    For i = 1 To n_
        e_(i, p_ + 1) = e
    Next i
    For i = 1 To n_
        dE = 0
        For j = 1 To n_
            dE = dE + wp_(j) * e_(j, p_ + 1)
        Next j
        e_(i, p_) = functions.d_Sigmoid(z_(i, p_)) * dE
    Next i
    
    For k = p_ - 1 To 1 Step -1
        For i = 1 To n_
            dE = 0#
            For j = 1 To n_
                dE = dE + w_(i, j, k + 1) * e_(j, k + 1)
            Next j
            e_(i, k) = functions.d_Sigmoid(z_(i, k)) * dE
        Next i
    Next k
End Sub
'-------------------------------------------
Function Output(ByVal x As Variant) As Double
    Dim dX As Double
    Call ComputeZ(x)
    Dim i As Long
    
    For i = 1 To n_
        dX = dX + wp_(i) * z_(i, p_)
    Next i
    
    Output = functions.Sigmoid(dX)
End Function
'-------------------------------------------
Sub UpdateWeight(ByVal x As Variant, ByVal s As Double, Optional ByVal sMethod = "backpropagation")
    Select Case VBA.Strings.LCase(sMethod)
        Case "backpropagation":
            Call UpdateWeight_backpropagation(x, s)
        Case "momentum":
            Call UpdateWeight_momentum(x, s)
        Case "learning":
            Call UpdateWeight_learning(x, s)
    End Select
    learning_ = learning_ * factor_
End Sub
'-------------------------------------------
Sub UpdateWeight_momentum(ByVal x As Variant, ByVal s As Double)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim z As Double
    Dim e As Double
    
    If Not VBA.Information.IsArray(dW_) Then
        ReDim dWp_(1 To n_)
        ReDim dW_(1 To n_, 1 To n_, 0 To p_)
    End If
    
    Call ComputeError(x, s)
    
    For i = 1 To n_
        dWp_(i) = -lambda_ * learning_ * z_(i, p_) * e_(i, p_ + 1) + (VBA.Conversion.CDbl(1) - lambda_) * dWp_(i)
        wp_(i) = wp_(i) + dWp_(i)
    Next i
    
    For k = p_ To 1 Step -1
        For i = 1 To n_
            For j = 1 To n_
                dW_(i, j, k) = -lambda_ * learning_ * e_(i, k) * z_(i, k) + (VBA.Conversion.CDbl(1) - lambda_) * dW_(i, j, k)
                w_(i, j, k) = w_(i, j, k) + dW_(i, j, k)
            Next j
        Next i
    Next k
End Sub
'-------------------------------------------
Private Sub UpdateWeight_learning(ByVal x As Variant, ByVal s As Double)
    learning_ = VBA.Math.Rnd
    Call UpdateWeight_backpropagation(x, s)
End Sub
'-------------------------------------------
Private Sub UpdateWeight_backpropagation(ByVal x As Variant, ByVal s As Double)
    Dim y As Double
    Dim e As Double
    Dim i As Long
    
    y = Output(x)
    e = (y - s) * y * (VBA.Conversion.CDbl(1) - y)
    
    For i = 1 To n_
        wp_(i) = wp_(i) - learning_ * z_(i, p_) * e
    Next i
    Call ComputeError(x, s)
    Call locUpdateWeight_backpropagation(x, s)
End Sub
'-------------------------------------------
Private Sub locUpdateWeight_backpropagation(ByVal x As Variant, ByVal s As Double)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    For k = p_ To 1 Step -1
        For i = 1 To n_
            For j = 1 To n_
                w_(i, j, k) = w_(i, j, k) - learning_ * e_(i, k) * z_(i, k)
            Next j
        Next i
    Next k
End Sub
'-------------------------------------------
Private Sub Class_Initialize()
    learning_ = 0.5
    factor_ = 1#
End Sub
'-------------------------------------------
Sub SetFactor(ByVal dFactor As Double)
    factor_ = dFactor
End Sub
'-------------------------------------------
