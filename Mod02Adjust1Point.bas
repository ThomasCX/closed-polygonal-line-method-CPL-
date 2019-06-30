Public Sub Adjust1Point(ByVal m As Integer, ByVal dv As Double)
    Dim Vtemp0 As xy, Vtemp1 As xy, GVtemp As Double, n As Integer
    Dim d1 As Double, d2 As Double
    Dim i As Integer, j As Integer
    Dim DistanceofDtoVSZtemp As Double  
    Dim bz As Boolean
    Dim Xjiaxs As Double, Xjianxs As Double, Yjiaxs As Double, Yjianxs As Double 
    Xjiaxs = 1#: Xjianxs = 1#: Yjiaxs = 1#: Yjianxs = 1#                         
    
    If V(m).X >= 0.99 Then Xjiaxs = 0
    If V(m).X <= -0.99 Then Xjianxs = 0
    
    If V(m).Y >= 0.99 Then Yjiaxs = 0
    If V(m).Y <= -0.99 Then Yjianxs = 0
    
    Vtemp0.X = V(m).X: Vtemp0.Y = V(m).Y: Vtemp1.X = V(m).X: Vtemp1.Y = V(m).Y
    GVtemp = GV(m)
    DistanceofDtoVSZtemp = DistanceofDtoVSZ
    i = 0
    i = i + 1
    V(m).X = Vtemp0.X: V(m).Y = Vtemp0.Y - Yjianxs * dv
    Call Adjust1PointSub1(i, m, DistanceofDtoVSZtemp, Vtemp1)   '调整
    
    i = i + 1
    V(m).X = Vtemp0.X: V(m).Y = Vtemp0.Y + Yjiaxs * dv
     Call Adjust1PointSub1(i, m, DistanceofDtoVSZtemp, Vtemp1)   '调整

    i = i + 1
    V(m).X = Vtemp0.X - Xjianxs * dv: V(m).Y = Vtemp0.Y
     Call Adjust1PointSub1(i, m, DistanceofDtoVSZtemp, Vtemp1)   '调整
     
     i = i + 1
    V(m).X = Vtemp0.X + Xjiaxs * dv: V(m).Y = Vtemp0.Y
    Call Adjust1PointSub1(i, m, DistanceofDtoVSZtemp, Vtemp1)   '调整
    
    i = i + 1
    V(m).X = Vtemp0.X + Xjiaxs * dv: V(m).Y = Vtemp0.Y + Yjiaxs * dv
     Call Adjust1PointSub1(i, m, DistanceofDtoVSZtemp, Vtemp1)   '调整
     
     i = i + 1
    V(m).X = Vtemp0.X + Xjiaxs * dv: V(m).Y = Vtemp0.Y - Yjianxs * dv
     Call Adjust1PointSub1(i, m, DistanceofDtoVSZtemp, Vtemp1)   '调整
     
    i = i + 1
    V(m).X = Vtemp0.X - Xjianxs * dv: V(m).Y = Vtemp0.Y + Yjiaxs * dv
     Call Adjust1PointSub1(i, m, DistanceofDtoVSZtemp, Vtemp1)   '调整
    '
    i = i + 1
    V(m).X = Vtemp0.X - Xjianxs * dv: V(m).Y = Vtemp0.Y - Yjianxs * dv
    Call Adjust1PointSub1(i, m, DistanceofDtoVSZtemp, Vtemp1)   '调整
    
    d1 = MoveDirectionDistance(1): j = 1               '求出最小的
    For i = 2 To 8
        If MoveDirectionDistance(i) < d1 Then                '距离函数变小
           d1 = MoveDirectionDistance(i): j = i
        End If
    Next i
    
    If (DistanceofDtoVSZtemp - MoveDirectionDistance(j) > 0.002) Then            
          V(m).X = MoveDirectionV(j).X: V(m).Y = MoveDirectionV(j).Y   
          Call SegmentExpression(V, tmin)        
          Call DataProject(D(), V, uxy, tsx)     
         
         For i = LBound(V) To UBound(V)                  
            bz = False
            If (Pi(i) - 1) >= 0.01 Then bz = True: Exit For  
         Next i
         
         If (bz = True) Then
                V(m).X = Vtemp0.X: V(m).Y = Vtemp0.Y
         End If
    Else
         V(m).X = Vtemp0.X: V(m).Y = Vtemp0.Y
    End If
    Call SegmentExpression(V, tmin)        
    Call DataProject(D(), V, uxy, tsx)     
Adjust1Point_Exit:

End Sub


Public Sub Adjust1PointSub1(ByVal i As Integer, ByVal m As Integer, ByVal DistanceofDtoVSZtemp As Double, ByRef Vtemp1 As xy)
     Call SegmentExpression(V, tmin)        
     Call DataProject(D(), V, uxy, tsx)     
     MoveDirectionDistance(i) = DistanceofDtoVSZ 
     MoveDirectionV(i).X = V(m).X: MoveDirectionV(i).Y = V(m).Y 
End Sub



