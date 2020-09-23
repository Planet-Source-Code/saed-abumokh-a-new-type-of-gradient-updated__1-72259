Attribute VB_Name = "mdlGradientFunctions"
Public Const TriCenter = 5
Public Enum RGBEnum
    r = 0
    g = 1
    b = 2
End Enum

Enum GradientRectDirection
    Horizontal = GRADIENT_FILL_RECT_H
    Vertical = GRADIENT_FILL_RECT_V
End Enum

Public Type PointColor
    X As Long
    Y As Long
    Color As Long
End Type


Public Function LongToSignedShort(dwUnsigned As Long) As Integer
    
   If dwUnsigned < 32768 Then
      LongToSignedShort = CInt(dwUnsigned)
   Else
      LongToSignedShort = CInt(dwUnsigned - &H10000)
   End If
    
End Function


Public Function RedColor(Color As Long) As Long
    RedColor = LongToSignedShort((Color And &HFF&) * 256)
End Function

Public Function GreenColor(Color As Long) As Long
    GreenColor = LongToSignedShort(((Color And &HFF00&) \ &H100&) * 256)
End Function

Public Function BlueColor(Color As Long) As Long
    BlueColor = LongToSignedShort(((Color And &HFF0000) \ &H10000) * 256)
End Function

Public Sub GradientRectHV(hDC As Long, PointColor1 As PointColor, PointColor2 As PointColor, Direction As GradientRectDirection)
    Dim RVetrex(0 To 1) As TRIVERTEX
    Dim GradRect As GRADIENT_RECT
    
    With RVetrex(0)
        .X = PointColor1.X
        .Y = PointColor1.Y
        .Alpha = 0
        .Red = RedColor(PointColor1.Color)
        .Green = GreenColor(PointColor1.Color)
        .Blue = BlueColor(PointColor1.Color)
    End With
    
    With RVetrex(1)
        .X = PointColor2.X
        .Y = PointColor2.Y
        .Alpha = 0
        .Red = RedColor(PointColor2.Color)
        .Green = GreenColor(PointColor2.Color)
        .Blue = BlueColor(PointColor2.Color)
    End With
    
    GradRect.UpperLeft = 0
    GradRect.LowerRight = 1
    
    GradientFill hDC, RVetrex(0), 2, GradRect, 1, Direction
End Sub

Public Sub GradientTriangle(hDC As Long, PointColor1 As PointColor, PointColor2 As PointColor, PointColor3 As PointColor)
    Dim TVetrex(0 To 2) As TRIVERTEX
    Dim GradTriangle As GRADIENT_TRIANGLE
    
    With TVetrex(0)
        .X = PointColor1.X
        .Y = PointColor1.Y
        .Alpha = 32767
        .Red = RedColor(PointColor1.Color)
        .Green = GreenColor(PointColor1.Color)
        .Blue = BlueColor(PointColor1.Color)
    End With
    
    With TVetrex(1)
        .X = PointColor2.X
        .Y = PointColor2.Y
        .Alpha = 0
        .Red = RedColor(PointColor2.Color)
        .Green = GreenColor(PointColor2.Color)
        .Blue = BlueColor(PointColor2.Color)
    End With
    
    With TVetrex(2)
        .X = PointColor3.X
        .Y = PointColor3.Y
        .Alpha = 0
        .Red = RedColor(PointColor3.Color)
        .Green = GreenColor(PointColor3.Color)
        .Blue = BlueColor(PointColor3.Color)
    End With
    
    With GradTriangle
        .Vertex1 = 0
        .Vertex2 = 1
        .Vertex3 = 2
    End With
    
    GradientFill hDC, TVetrex(0), 3, GradTriangle, 1, GRADIENT_FILL_TRIANGLE
End Sub

Private Sub GradientLineH(hDC As Long, X1 As Long, X2 As Long, Y As Long, Color1 As Long, Color2 As Long)

    Dim pClr1 As PointColor
    Dim pClr2 As PointColor
    
    pClr1.X = X1
    pClr1.Y = Y
    pClr1.Color = Color1
    
    pClr2.X = X2
    pClr2.Y = Y + 1
    pClr2.Color = Color2
    
    GradientRectHV hDC, pClr1, pClr2, Horizontal
End Sub

Public Sub GradientRect4Corners(hDC As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Color1 As Long, Color2 As Long, Color3 As Long, Color4 As Long)
    On Error Resume Next
    Dim gCount As Long
    Dim pColor1 As PointColor, pColor2 As PointColor
    
    For gCount = Y1 To Y2
        pColor1.X = X1
        pColor1.Y = gCount
        pColor1.Color = GetColorLevel(Color1, Color3, 256 / (Y2 - Y1) * (gCount - Y1))
        
        pColor2.X = X2
        pColor2.Y = gCount
        pColor2.Color = GetColorLevel(Color2, Color4, 256 / (Y2 - Y1) * (gCount - Y1))
        
        GradientLineH hDC, pColor1.X, pColor2.X, gCount, pColor1.Color, pColor2.Color
    Next

End Sub



Public Sub GradientPolyFromCenter(hDC As Long, Xs() As Long, Ys() As Long, Colors() As Long, PointColorCount As Integer, CenterPointColor As Long, AutoCenterPointColor As Boolean, CenterPointX As Single, AutoCenterPointX As Boolean, CenterPointY As Single, AutoCenterPointY As Boolean)
    
    Dim ptColor() As PointColor
    Dim AvgX
    Dim AvgY
    Dim AvgColor
    Dim nCount
    Dim PointColorCenter As PointColor
    Dim Rs() As Long, Bs() As Long, Gs() As Long
    Dim AvgR, AvgG, AvgB


    ReDim ptColor(PointColorCount)
    ReDim Rs(PointColorCount)
    ReDim Bs(PointColorCount)
    ReDim Gs(PointColorCount)



    For nCount = 0 To PointColorCount

        ptColor(nCount).X = Xs(nCount)
        ptColor(nCount).Y = Ys(nCount)
        ptColor(nCount).Color = Colors(nCount)
        
        If AutoCenterPointColor = True Then
            Rs(nCount) = FindRed(Colors(nCount))
            Gs(nCount) = FindGreen(Colors(nCount))
            Bs(nCount) = FindBlue(Colors(nCount))
            
            AvgR = Average(Rs, PointColorCount)
            AvgG = Average(Gs, PointColorCount)
            AvgB = Average(Bs, PointColorCount)
            AvgColors = RGB(AvgR, AvgG, AvgB)
            PointColorCenter.Color = AvgColors
        Else
            PointColorCenter.Color = CenterPointColor
        End If
    Next
    
    If AutoCenterPointX = False Then
        PointColorCenter.X = CenterPointX
    Else
        AvgX = Average(Xs, PointColorCount)
        PointColorCenter.X = AvgX
    End If
    
    If AutoCenterPointY = False Then
        PointColorCenter.Y = CenterPointY
    Else
        AvgY = Average(Ys, PointColorCount)
        PointColorCenter.Y = AvgY
    End If



    For nCount = 0 To PointColorCount - 1
        GradientTriangle hDC, ptColor(nCount), ptColor(nCount + 1), PointColorCenter
    Next
    GradientTriangle hDC, ptColor(nCount), ptColor(1), PointColorCenter

End Sub

