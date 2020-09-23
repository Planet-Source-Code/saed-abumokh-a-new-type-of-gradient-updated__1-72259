VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Gradient Background Builder ( Draw a Rectangle &  Choose Four Colours )"
   ClientHeight    =   6510
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9120
   DrawWidth       =   2
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   434
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   608
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   0
      ScaleHeight     =   137
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   233
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6720
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   3960
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim pc(0 To 3) As PointColor
Dim CurrX As Integer, CurrY As Integer

Private Sub Form_Load()
i = 0
Me.Show
Me.AutoRedraw = True
Picture1.AutoRedraw = True
CommonDialog1.Flags = cdlCCRGBInit
    For i = 0 To 3
        pc(i).Color = 0
    Next
    
    Dim strMsg As String
    strMsg = "You will see a new type of gradient you have never seen it before!" & vbCrLf & _
    "Draw a rectangle and choose four colors and you will see the result" & vbCrLf & _
    "I have wrote extra gradient functions and i didnt use them in this project" & vbCrLf & _
    "See 'GradientFunctions' module, you can draw also triangular gradient and polygonal gradient!"
    MsgBox strMsg, , "Saed AbuMokh"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CurrX = X
    CurrY = Y
    Shape1.Visible = True
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button > 0 Then
        Shape1.Move CurrX, CurrY, X - CurrX, Y - CurrY
        Picture1.Move Picture1.Left, Picture1.Top, Shape1.Width, Shape1.Height
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim g As New Gradient
    If (CurrX = X) And (CurrY = Y) Then Exit Sub
    
    For i = 0 To 3
        CommonDialog1.DialogTitle = "Choose color for Corner " ' & i + 1
        CommonDialog1.ShowColor
        pc(i).Color = CommonDialog1.Color
        
        '***************************
        g.Rectangle4Colors Picture1.hdc, 0, 0, Picture1.Width, Picture1.Height, pc(0).Color, pc(1).Color, pc(2).Color, pc(3).Color
        '***************************
        
        Picture1.Picture = Picture1.Image
        Me.PaintPicture Picture1.Picture, Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height, , , , , vbSrcCopy
        Shape1.Visible = False
    Next
    
    If i > 3 Then
        i = 0
    End If
    
End Sub
