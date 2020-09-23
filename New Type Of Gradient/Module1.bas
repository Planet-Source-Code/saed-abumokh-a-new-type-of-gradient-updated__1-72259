Attribute VB_Name = "mdlAPIGradientFunctions"
Public Const GRADIENT_FILL_RECT_H  As Long = &H0
Public Const GRADIENT_FILL_RECT_V  As Long = &H1
Public Const GRADIENT_FILL_TRIANGLE As Long = &H2

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
   
   
   
'_______________________________________________________



Public Declare Function GradientFill Lib "msimg32" (ByVal hDC As Long, pVertex As Any, ByVal dwNumVertex As Long, pMesh As Any, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long




