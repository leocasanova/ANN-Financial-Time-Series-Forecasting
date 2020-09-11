VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Perceptron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'----------------------
Private w_ As Variant
Private learning_ As Double
'----------------------
Private Sub Class_Initialize()
    '------------------
    ReDim w_(0 To 1)
    '------------------
    w_(0) = VBA.Math.Rnd()
    w_(1) = VBA.Math.Rnd()
    learning_ = 0.5
    '------------------
End Sub
'----------------------
Function UpdateWeight(ByVal x As Variant, ByVal s As Double)
    '------------------
    Dim i As Long
    Dim v As Variant
    Dim y As Double
    Dim dRes As Double
    Dim e As Double
    '------------------
    If functions.GetDim(x) = 1 And UBound(x, 1) - LBound(x, 1) = 1 Then
        '--------------
        y = Output(x)
        e = (y - s) * y * (1# - y)
        w_(0) = w_(0) - learning_ * x(LBound(x, 1)) * e
        w_(1) = w_(1) - learning_ * x(UBound(x, 1)) * e
        '--------------
    End If
    '------------------
    UpdateWeight = w_
End Function
'----------------------
Function Output(ByVal x As Variant)
    '------------------
    Dim i As Long
    Dim dX As Double
    '------------------
    If functions.GetDim(x) = 1 And UBound(x, 1) - LBound(x, 1) = 1 Then
        dX = x(LBound(x, 1)) * w_(0) + x(UBound(x, 1)) * w_(1)
    Else
        dX = -10000000000#
    End If
    '------------------
    Output = Sigmoid(dX)
    '------------------
End Function
'----------------------
Function GetWeigth()
    GetWeigth = w_
End Function
'----------------------
Sub SetLearningRate(ByVal dRate As Double)
    learning_ = dRate
End Sub
'----------------------
Sub SetWeight(ByVal w As Variant)
    '------------------
    Dim v As Variant
    Dim i As Long
    '------------------
    If functions.GetNbElement(w) = 2 Then
        '--------------
        i = 0
        '--------------
        For Each v In w
            '----------
            w_(i) = v
            i = i + 1
            '----------
        Next v
        '--------------
    End If
End Sub
'----------------------
Function UpdateWeightM(ByVal x As Variant, ByVal s As Double, ByVal e As Double)
    '------------------
    Dim i As Long
    Dim v As Variant
    Dim y As Double
    Dim dRes As Double
    '------------------
    If functions.GetDim(x) = 1 And functions.GetNbElement(x) = 2 Then
        y = Output(x)
        i = 0
        '--------------
        For Each v In x
            '----------
            w_(i) = w_(i) - learning_ * v * e
            i = i + 1
            '----------
        Next v
    End If
    '------------------
End Function
'----------------------
Private Sub Class_Terminate()

End Sub
