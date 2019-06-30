'插点调整
Public Sub InsertV(ByVal Vnumber As Integer)
    Dim i As Integer, dt As Double
    Dim ET As Double
    Dim V1X As xy
    Dim t() As Double, P() As xy
    '
    '(1.1) 求数据集在第一主成分上的投影P
    ReDim t(1 To UBound(D)): ReDim P(1 To UBound(D))
    ET = 0
    For i = 1 To Nd
       '投影指标
       t(i) = D(i).x * ta + D(i).y * tb
       ET = ET + t(i)
       '投影
       P(i).x = t(i) * ta
       P(i).y = t(i) * tb
    Next i
    tmax = t(1): tmin = t(1)
    For i = 2 To Nd
        If tmax < t(i) Then tmax = t(i)
        If tmin > t(i) Then tmin = t(i)
    Next i
    '
    V1X.x = (ET / Nd) * ta
    V1X.y = (ET / Nd) * tb
    '
    ReDim V(1 To Vnumber)
    V(1).x = tmin * ta: V(1).y = tmin * tb
    dt = (tmax - tmin) / (Vnumber - 1)
    For i = 2 To Vnumber - 1
        V(i).x = ta * (dt * (i - 1) + tmin)
        V(i).y = tb * (dt * (i - 1) + tmin)
    Next i
    V(Vnumber).x = tmax * ta: V(Vnumber).y = tmax * tb
    
End Sub

