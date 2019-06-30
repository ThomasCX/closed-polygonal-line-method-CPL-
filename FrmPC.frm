'-------------------------------------------------------------------
Private Sub Cmdstart_ClIck()           '“开始”
   '定义临时变量
   Dim i As Integer, j As Integer
   Dim Vpoint As xy, fPCA As xy
   Dim t1 As Double
   
   Call Drawcoordinate(PicC_Qc, vbWhite, vbWhite, vbWhite)  
   Call OpenTextFile(DataFileName) 
    
   '画数据点[D().x,D().y]
    Call DrawDataPoint(D)
   
'   '第一主成分
       tmax = 0.99999
       tmin = 0.00001
       ReDim V(1 To 5)
       V(1).X = -0.1: V(1).Y = -0.1
       V(2).X = -0.1: V(2).Y = 0.1
       V(3).X = 0.1: V(3).Y = 0.1
       V(4).X = 0.1: V(4).Y = -0.1
       V(5).X = -0.1: V(5).Y = -0.1
   '
  
   For i = LBound(V) To UBound(V)
       Call DrawData(PicC_Qc, V(i), vbRed, "DrawCircle", 30)     '在图片框中,画点,颜色,形状,大小
   Next i
    Call SegmentExpression(V, tmin)   
     For i = 1 To UBound(uxy)
          For t1 = tsx(i) To tsx(i + 1) Step 0.002
             Vpoint.X = V(i).X + (t1 - tsx(i)) * uxy(i).X: Vpoint.Y = V(i).Y + (t1 - tsx(i)) * uxy(i).Y
             Call DrawData(PicC_Qc, Vpoint, vbRed, "DrawForkX", 2)     
          Next t1
    Next i
    Cmdstart.Enabled = False           
    CmdAdjust.Enabled = True           
    CmdInsert1V.Enabled = True         
    CmdProjectAndSave.Enabled = True   
End Sub


Private Sub CmdAdjust_Click()             
    Dim i As Integer, j As Integer
    Dim CVcmp As Double, ux As Double, uy As Double
    Dim tt As Double, t1 As Double, Vpoint As xy
    Dim Kad As Integer
    Dim AdjustNum As Integer
    '
    Dim D1V As Double
    Dim DLS1 As Double
    Dim Vtemp As xy                      '顶点
    Dim tmintemp As Double
    Dim k1 As Double
    Dim m As Integer
    '
    Call SegmentExpression(V, tmin)        
    Call DataProject(D(), V, uxy, tsx)     
    For AdjustNum = 1 To 50
        DLS1 = DistanceofDtoVSZ
        For Kad = 1 To UBound(V) - 1 Step 1
             DoEvents
             
             Call Adjust1Point(Kad, 0.02)
             '
             
             TxtPara(1).Text = Kad: TxtPara(2).Text = PV(Kad)
             If Kad = 1 Then
                TxtPara(3).Text = Pi(Kad + 1) - 1
                Else
                  If Kad = UBound(V) Then
                     TxtPara(3).Text = Pi(Kad - 1) - 1
                  Else
                     TxtPara(3).Text = Pi(Kad) - 1
                  End If
             End If
        Next Kad
        V(UBound(V)).X = V(1).X: V(UBound(V)).Y = V(1).Y
        
         D1V = DistanceofDtoVSZ
         tmintemp = tmin
         Vtemp.X = V(1).X: Vtemp.Y = V(1).Y
         tmin = tmin + 0.01
         tsx(1) = tmin
         V(1).X = V(2).X - (tsx(2) - tsx(1)) * uxy(1).X
         V(1).Y = V(2).Y - (tsx(2) - tsx(1)) * uxy(1).Y
         Call SegmentExpression(V, tmin)        
         Call DataProject(D(), V, uxy, tsx)     
         If D1V < DistanceofDtoVSZ Then   
              V(1).X = Vtemp.X: V(1).Y = Vtemp.Y
              tmin = tmin + 0.01
              tsx(1) = tmin
              Call SegmentExpression(V, tmin)
              Call DataProject(D(), V, uxy, tsx)
          End If
          
          V(UBound(V)).X = V(1).X: V(UBound(V)).Y = V(1).Y            
        
            PicC_Qc.Cls
            Call DrawDataPoint(D)
             For i = 1 To UBound(uxy)
                   For t1 = tsx(i) To tsx(i + 1) Step 0.002
                      Vpoint.X = V(i).X + (t1 - tsx(i)) * uxy(i).X: Vpoint.Y = V(i).Y + (t1 - tsx(i)) * uxy(i).Y
                      Call DrawData(PicC_Qc, Vpoint, vbBlue, "DrawForkX", 2)    
                   Next t1
             Next i
             If Abs(DLS1 - DistanceofDtoVSZ) < 0.002 Then Exit For   //maximum distance 0.002
    Next AdjustNum  
End Sub

'画数据点[D().x,D().y]
Private Sub DrawDataPoint(ByRef DataPonit() As xy)
   Dim i As Integer
   For i = 1 To UBound(D)                     '对所有数据点循环
      'Call DrawData(PicC_Qc, DataPonit(i), vbBlack, "DrawCircle", 10)
      Call DrawData(PicC_Qc, DataPonit(i), vbRed, "DrawCircle", 10)      '设置所画D(i)点参数
   Next i
End Sub
