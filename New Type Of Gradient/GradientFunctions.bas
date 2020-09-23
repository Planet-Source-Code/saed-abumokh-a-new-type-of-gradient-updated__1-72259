Attribute VB_Name = "GradientFunctions"

Public Type TRIVERTEX
   X As Long
   Y As Long
   Red As Integer
   Green As Integer
   Blue As Integer
   Alpha As Integer
End Type

Public Type GRADIENT_TRIANGLE
   Vertex1 As Long
   Vertex2 As Long
   Vertex3 As Long
End Type
   
Public Type GRADIENT_RECT
   UpperLeft As Long
   LowerRight As Long
End Type
   
   
Public Enum GradientRectDirection
    Horizontal = &H0
    Vertical = &H1
End Enum

Public Type PointColor
    X As Long
    Y As Long
    Color As Long
End Type

   
'_______________________________________________________



Public Declare Function GradientFill Lib "msimg32" (ByVal hDC As Long, pVertex As Any, ByVal dwNumVertex As Long, pMesh As Any, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long

Private Enum GRADIENT_FILL_MODES
      GRADIENT_FILL_RECT_H = &H0
      GRADIENT_FILL_RECT_V = &H1
      GRADIENT_FILL_TRIANGLE = &H2
End Enum
Private Enum RGBEnum
    r = 0
    g = 1
    b = 2
End Enum















Public Function FindRed(Color As Long) As Long
    Dim Red, Green, Blue
    Blue = Color \ 65536
    Green = (Color - (Blue * 65536)) \ 256
    Red = Color - ((Blue * 65536) + (Green * 256))
    FindRed = Red
End Function
Public Function FindGreen(Color As Long) As Long
    Dim Green, Blue
    Blue = Color \ 65536
    Green = (Color - (Blue * 65536)) \ 256
    FindGreen = Green
End Function
Public Function FindBlue(Color As Long) As Long
    Dim Blue
    Blue = Color \ 65536
    FindBlue = Blue
End Function

Public Function BlendColors(Color1 As Long, Color2 As Long, BlendValue As Byte) As Long
    On Error Resume Next
    Dim R1, G1, B1
    Dim R2, G2, B2
    Dim stepRed, stepGreen, stepBlue
    Dim lvlRed, lvlGreen, lvlBlue
'______________________________________________________________
    R1 = FindRed(Color1)
    
    G1 = FindGreen(Color1)
    B1 = FindBlue(Color1)
    
    R2 = FindRed(Color2)
    G2 = FindGreen(Color2)
    B2 = FindBlue(Color2)
'______________________________________________________________
    
    stepRed = Abs(R1 - R2) / 256
    stepGreen = Abs(G1 - G2) / 256
    stepBlue = Abs(B1 - B2) / 256
'______________________________________________________________
    
    If R1 > R2 Then
        lvlRed = R1 - (ColorLevel * stepRed)
    ElseIf R2 > R1 Then
        lvlRed = (ColorLevel * stepRed) + R1
    ElseIf R1 = R2 Then
        lvlRed = R1
    End If
    
    If G1 > G2 Then
        lvlGreen = G1 - (ColorLevel * stepGreen)
    ElseIf G2 > G1 Then
        lvlGreen = (ColorLevel * stepGreen) + G1
    ElseIf G1 = G2 Then
        lvlGreen = G1
    End If
    
    
    If B1 > B2 Then
        lvlBlue = B1 - (ColorLevel * stepBlue)
    ElseIf B2 > B1 Then
        lvlBlue = (ColorLevel * stepBlue) + B1
    ElseIf B1 = B2 Then
        lvlBlue = B1
    End If
'______________________________________________________________
    BlendColors = RGB(lvlRed, lvlGreen, lvlBlue)
    
End Function






















Private Function LongToSignedShort(dwUnsigned As Long) As Integer
    
   If dwUnsigned < 32768 Then
      LongToSignedShort = CInt(dwUnsigned)
   Else
      LongToSignedShort = CInt(dwUnsigned - &H10000)
   End If
    
End Function

Private Function RedColor(Color As Long) As Long
    RedColor = LongToSignedShort((Color And &HFF&) * 256)
End Function

Private Function GreenColor(Color As Long) As Long
    GreenColor = LongToSignedShort(((Color And &HFF00&) \ &H100&) * 256)
End Function

Private Function BlueColor(Color As Long) As Long
    BlueColor = LongToSignedShort(((Color And &HFF0000) \ &H10000) * 256)
End Function

Public Sub GradientRectHV(hDC As Long, PointColor1 As PointColor, PointColor2 As PointColor, Direction)
    Dim RVetrex(0 To 1) As TRIVERTEX
    Dim GradRect As GRADIENT_RECT
    
    With PointColor1
        RVetrex(0) = InputTRIVERTEX(.X, .Y, RedColor(.Color), GreenColor(.Color), BlueColor(.Color), 0)
    End With
    
    With PointColor2
        RVetrex(1) = InputTRIVERTEX(.X, .Y, RedColor(.Color), GreenColor(.Color), BlueColor(.Color), 0)
    End With
    
    GradRect.UpperLeft = 0
    GradRect.LowerRight = 1
    
    GradientFill hDC, RVetrex(0), 2, GradRect, 1, Direction
End Sub

Public Sub GradientTriangle(hDC As Long, PointColor1 As PointColor, PointColor2 As PointColor, PointColor3 As PointColor)
    Dim TVetrex(0 To 2) As TRIVERTEX
    Dim GradTriangle As GRADIENT_TRIANGLE
    
    With PointColor1
            TVetrex(0) = InputTRIVERTEX(.X, .Y, RedColor(.Color), GreenColor(.Color), BlueColor(.Color), 0)
    End With
    
    With PointColor2
            TVetrex(1) = InputTRIVERTEX(.X, .Y, RedColor(.Color), GreenColor(.Color), BlueColor(.Color), 0)
    End With
    
    With PointColor3
            TVetrex(2) = InputTRIVERTEX(.X, .Y, RedColor(.Color), GreenColor(.Color), BlueColor(.Color), 0)
    End With
    
    With GradTriangle
        .Vertex1 = 0: .Vertex2 = 1: .Vertex3 = 2
    End With
    
    GradientFill hDC, TVetrex(0), 3, GradTriangle, 1, GRADIENT_FILL_TRIANGLE
End Sub

Private Sub GradientLineH(hDC As Long, X1 As Long, X2 As Long, Y As Long, Color1 As Long, Color2 As Long)

    Dim pClr1 As PointColor
    Dim pClr2 As PointColor
    
    pClr1 = InputPointColor(X1, Y + 0, Color1)
    
    pClr2 = InputPointColor(X2, Y + 1, Color2)
    
    GradientRectHV hDC, pClr1, pClr2, Horizontal
End Sub

Public Sub GradientRect4Corners(hDC As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Color1 As Long, Color2 As Long, Color3 As Long, Color4 As Long)
    On Error Resume Next
    Dim gCount As Long
    Dim pColor1 As PointColor, pColor2 As PointColor
    
    For gCount = Y1 To Y2
        pColor1.X = X1
        pColor1.Y = gCount
        pColor1.Color = BlendColors(Color1, Color3, 256 / (Y2 - Y1) * (gCount - Y1))
        
        pColor2.X = X2
        pColor2.Y = gCount
        pColor2.Color = BlendColors(Color2, Color4, 256 / (Y2 - Y1) * (gCount - Y1))
        
        GradientLineH hDC, pColor1.X, pColor2.X, gCount, pColor1.Color, pColor2.Color
    Next

End Sub

Private Function InputTRIVERTEX(X As Long, Y As Long, Red As Integer, Green As Integer, Blue As Integer, Alpha As Integer) As TRIVERTEX

    With InputTRIVERTEX
        .X = CLng(X)
        .Y = CLng(Y)
        .Red = CLng(Red)
        .Green = CLng(Green)
        .Blue = CLng(Blue)
        .Alpha = CLng(Alpha)
    End With

End Function

Public Function InputPointColor(X, Y, Color) As PointColor

    With InputPointColor
        .X = CLng(X): .Y = CLng(Y): .Color = CLng(Color)
    End With
    
End Function


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


