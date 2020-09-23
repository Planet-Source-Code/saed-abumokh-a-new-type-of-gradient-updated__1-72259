Attribute VB_Name = "mdlColor"
Function FindRed(Color) As Long
    Dim Red, Green, Blue
    Blue = Color \ 65536
    Green = (Color - (Blue * 65536)) \ 256
    Red = Color - ((Blue * 65536) + (Green * 256))
    FindRed = Red
End Function
Function FindGreen(Color) As Long
    Dim Green, Blue
    Blue = Color \ 65536
    Green = (Color - (Blue * 65536)) \ 256
    FindGreen = Green
End Function
Function FindBlue(Color) As Long
    Dim Blue
    Blue = Color \ 65536
    FindBlue = Blue
End Function

Function GetColorLevel(Color1 As Long, Color2 As Long, ColorLevel As Byte) As Long
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
    
    stepRed = Abs(R1 - R2) / 255
    stepGreen = Abs(G1 - G2) / 255
    stepBlue = Abs(B1 - B2) / 255
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
    GetColorLevel = RGB(lvlRed, lvlGreen, lvlBlue)
    
End Function




