Public Sub SegmentExpression(ByRef Vex() As xy, ByVal tmin As Double)
   '
   Dim i As Integer, j As Integer, k As Integer
   Dim t1 As Double, d1 As Double
   Dim Vpoint As xy
   '
   ReDim tsx(1 To UBound(Vex))        '������߶�ͶӰָ��
   ReDim uxy(1 To UBound(Vex) - 1)    
   tsx(1) = tmin                      
   For i = 1 To UBound(Vex) - 1
       '�����ŷ�Ͼ���
       d1 = Sqr((Vex(i + 1).X - Vex(i).X) * (Vex(i + 1).X - Vex(i).X) + (Vex(i + 1).Y - Vex(i).Y) * (Vex(i + 1).Y - Vex(i).Y))
       uxy(i).X = (Vex(i + 1).X - Vex(i).X) / d1
       uxy(i).Y = (Vex(i + 1).Y - Vex(i).Y) / d1
   Next i
  
   For i = 1 To UBound(Vex) - 1
       d1 = Sqr((Vex(i + 1).X - Vex(i).X) * (Vex(i + 1).X - Vex(i).X) + (Vex(i + 1).Y - Vex(i).Y) * (Vex(i + 1).Y - Vex(i).Y))
       tsx(i + 1) = d1 + tsx(i)  'tsxͶӰָ��
       
   Next i
End Sub

Public Sub DataProject(ByRef D() As xy, ByRef V() As xy, ByRef uxy() As xy, ByRef tsx() As Double)
   Dim i As Integer, j As Integer, n As Integer, k As Integer, f As Integer
   Dim t1 As Double, Drtsx As Double, d1 As Double, namnapp As Double, namnap As Double
   Dim ProjectPoint As xy
 
   namnapp = 0.13           'the optimalsolution of �� p
   beta= 0.3                   '��
   
   
   ReDim DtoVS(1 To UBound(D))             
   ReDim DistanceofDtoVS(1 To UBound(D))   
   For j = 1 To UBound(D)  '�����ݵ�ѭ��(��ʼ)
        DtoVS(j) = 0                    
        DistanceofDtoVS(j) = 1000       
        For i = 1 To UBound(V) - 1      '���߶�ѭ��(��ʼ)
             t1 = (D(j).X - V(i).X) * uxy(i).X + (D(j).Y - V(i).Y) * uxy(i).Y + tsx(i)  
             If t1 <= tsx(i) Then    
                Drtsx = (D(j).X - V(i).X) * (D(j).X - V(i).X) + (D(j).Y - V(i).Y) * (D(j).Y - V(i).Y)
                If Drtsx <= DistanceofDtoVS(j) Then DtoVS(j) = i: DistanceofDtoVS(j) = Drtsx 
             Else
                If t1 >= tsx(i + 1) Then   
                  Drtsx = (D(j).X - V(i + 1).X) * (D(j).X - V(i + 1).X) + (D(j).Y - V(i + 1).Y) * (D(j).Y - V(i + 1).Y)
                  If Drtsx <= DistanceofDtoVS(j) Then DtoVS(j) = i + 1: DistanceofDtoVS(j) = Drtsx 
                Else                      
                  ProjectPoint.X = V(i).X + (t1 - tsx(i)) * uxy(i).X
                  ProjectPoint.Y = V(i).Y + (t1 - tsx(i)) * uxy(i).Y
                  Drtsx = (D(j).X - ProjectPoint.X) * (D(j).X - ProjectPoint.X) + (D(j).Y - ProjectPoint.Y) * (D(j).Y - ProjectPoint.Y)
                  If Drtsx <= DistanceofDtoVS(j) Then DtoVS(j) = 20000 + i: DistanceofDtoVS(j) = Drtsx 
                End If
             End If
         Next i                         '���߶�ѭ��(����)
    Next j                '�����ݵ�ѭ��(����)
     k = UBound(uxy)             
     ReDim Cgm(1 To k)          
     ReDim VV(1 To k + 1)        
     ReDim u2(1 To k)            
     ReDim Pi(1 To k + 1)       
     ReDim PV(1 To k + 1)        
     ReDim DairTa(1 To k + 1)    
     ReDim GV(1 To k + 1)        
     For i = 1 To k                
         Cgm(i) = 0
         For j = 1 To UBound(D)   
             If ((DtoVS(j) - 20000)) = i Then Cgm(i) = Cgm(i) + DistanceofDtoVS(j)
         Next j
     Next i
     For i = 1 To k + 1            '���߶�ѭ��
         VV(i) = 0
         For j = 1 To UBound(D)    '������ѭ��
             If DtoVS(j) = i Then VV(i) = VV(i) + DistanceofDtoVS(j)
         Next j
     Next i
     '�����߶γ���ƽ��
     For i = 1 To k             '���߶�Լ��ѭ��
         u2(i) = (V(i + 1).X - V(i).X) * (V(i + 1).X - V(i).X) + (V(i + 1).Y - V(i).Y) * (V(i + 1).Y - V(i).Y)
     Next i
     Pi(1) = 0: Pi(k + 1) = 0
     For i = 2 To k
         '
         d1 = (V(i - 1).X - V(i).X) * (V(i + 1).X - V(i).X) + (V(i - 1).Y - V(i).Y) * (V(i + 1).Y - V(i).Y)
         
         t1 = Sqr((V(i - 1).X - V(i).X) * (V(i - 1).X - V(i).X) + (V(i - 1).Y - V(i).Y) * (V(i - 1).Y - V(i).Y))
         t1 = t1 * Sqr((V(i + 1).X - V(i).X) * (V(i + 1).X - V(i).X) + (V(i + 1).Y - V(i).Y) * (V(i + 1).Y - V(i).Y))
         Pi(i) = 1 + d1 / t1    'ȡr=1
     Next i
     For i = 1 To k + 1 '�Զ���ѭ��
         If i = 1 Then PV(i) = u2(1) + Pi(2)
         If i = 2 Then PV(i) = u2(1) + Pi(2) + Pi(3)
         If (i > 2) And (i < k) Then PV(i) = Pi(i - 1) + Pi(i) + Pi(i + 1)
         If i = k Then PV(i) = Pi(i - 1) + Pi(i) + u2(i)
         If i = k + 1 Then PV(i) = Pi(i - 1) + u2(i - 1)
         PV(i) = PV(i) / (k + 1)
     Next i
 
      n = UBound(D)
     For i = 1 To k + 1 '�Զ���ѭ��
         If i = 1 Then DairTa(i) = VV(i) + Cgm(i)                                 'i=1
         If (i > 1) And (i < k + 1) Then DairTa(i) = Cgm(i - 1) + VV(i) + Cgm(i)  '1<i<k+1
         If i = k + 1 Then DairTa(i) = Cgm(i - 1) + VV(i)                         'i=k+1
         DairTa(i) = DairTa(i) / n
     Next i
     '���㶥��ľ���Լ��+�Ƕȳͷ�
     d1 = 0
     For i = 1 To n: d1 = d1 + DistanceofDtoVS(i): Next i 
     DistanceofDtoVSZ = d1  
     namnap = namnapp * k * (1 / ((n) ^ (1 / 3))) * Sqr(d1)
     For i = 1 To k + 1 '�Զ���ѭ��
         GV(i) = DairTa(i) + PV(i) * namnap
     Next i
     If k > beta * (1 / ((n) ^ (1 / 3))) * Sqr(d1) Then
          f = f + 1
     End If
     
End Sub
