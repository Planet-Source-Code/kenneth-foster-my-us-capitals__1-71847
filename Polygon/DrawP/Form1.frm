VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   FillColor       =   &H00808080&
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   390
      Left            =   1320
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   1
      Top             =   1185
      Width           =   420
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   1185
      Width           =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Private Type POINTAPI
    x As Long
    y As Long
 End Type

Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Position() As POINTAPI
Private Polyctr As Long
Dim chg As Boolean

Private Sub Command1_Click()
   chg = Not chg
   If chg = False Then
      pic1.FillColor = vbGreen
      DrawCheck pic1
   Else
      pic1.FillColor = vbRed
      DrawX pic1
   End If
End Sub

Private Sub Draw()
Dim i As Integer

    Polyctr = Polyctr + 1
    ReDim Preserve Position(Polyctr)
    
    For i = 1 To Polyctr
      pic1.Line (Position(i - 1).x, Position(i - 1).y)-(Position(i).x, Position(i).y)
    Next i
    
    pic1.Cls
    pic1.AutoRedraw = True
    Polygon pic1.hdc, Position(0), Polyctr
    pic1.AutoRedraw = False
    pic1.Refresh
End Sub

Private Sub DrawX(pic As Object)
pic.Picture = LoadPicture()
Polyctr = 12
ReDim Position(12)
Position(0).x = 4
Position(0).y = 10
Position(1).x = 9
Position(1).y = 3
Position(2).x = 13
Position(2).y = 10
Position(3).x = 19
Position(3).y = 3
Position(4).x = 24
Position(4).y = 8
Position(5).x = 16
Position(5).y = 12
Position(6).x = 22
Position(6).y = 17
Position(7).x = 17
Position(7).y = 21
Position(8).x = 14
Position(8).y = 14
Position(9).x = 9
Position(9).y = 20
Position(10).x = 4
Position(10).y = 17
Position(11).x = 11
Position(11).y = 12
Position(12).x = 4
Position(12).y = 9
Draw
   
End Sub

Private Sub DrawCheck(pic As Object)
pic.Picture = LoadPicture()
Polyctr = 6
ReDim Position(6)
Position(0).x = 3
Position(0).y = 15
Position(1).x = 6
Position(1).y = 11
Position(2).x = 10
Position(2).y = 17
Position(3).x = 15
Position(3).y = 3
Position(4).x = 19
Position(4).y = 8
Position(5).x = 10
Position(5).y = 21
Position(6).x = 3
Position(6).y = 15
Draw
End Sub

Private Sub Form_Load()

End Sub
