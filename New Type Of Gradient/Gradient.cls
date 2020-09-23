VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Gradient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Sub Rectangle(ByVal hDC As Long, X1, Y1, X2, Y2, FromColor, ToColor, ByVal IsVertical As Boolean)
    GradientRectHV IsVertical, _
                   InputPointColor(X1, Y1, FromColor), _
                   InputPointColor(X2, Y2, ToColor), _
                   Direction
End Sub

Public Sub Triangle(ByVal hDC As Long, X1, Y1, Color1, X2, Y2, Color2, X3, Y3, Color3)
    GradientTriangle hDC, _
                     InputPointColor(X1, Y1, Color1), _
                     InputPointColor(X2, Y2, Color2), _
                     InputPointColor(X3, Y3, Color3)
End Sub

Public Sub Rectangle4Colors(ByVal hDC As Long, X1, Y1, X2, Y2, Color1, Color2, Color3, Color4)
    GradientRect4Corners hDC, CLng(X1), CLng(Y1), CLng(X2), CLng(Y2), CLng(Color1), CLng(Color2), CLng(Color3), CLng(Color4)
End Sub

Public Sub PolygonFromCenter(hDC As Long, Xs() As Long, Ys() As Long, Colors() As Long, PointColorCount As Integer, CenterPointColor As Long, AutoCenterPointColor As Boolean, CenterPointX As Single, AutoCenterPointX As Boolean, CenterPointY As Single, AutoCenterPointY As Boolean)
    GradientPoly hDC, Xs(), Ys(), Colors(), PointColorCount, CenterPointColor, AutoCenterPointColor, CenterPointX, AutoCenterPointX, CenterPointY, AutoCenterPointY
End Sub

Private Function InputPointColor(X, Y, Color) As PointColor

    With InputPointColor
        .X = CLng(X): .Y = CLng(Y): .Color = CLng(Color)
    End With
    
End Sub



