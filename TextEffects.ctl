VERSION 5.00
Begin VB.UserControl TextEffects 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   1350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4740
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaskColor       =   &H00C0FFFF&
   ScaleHeight     =   1350
   ScaleWidth      =   4740
End
Attribute VB_Name = "TextEffects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Const m_def_Text = "Text Effects"
Const m_def_TextStyle = 0
Const m_def_TextBorderColor = vbBlack
Const m_def_TextColor = vbBlue
Const m_def_TextSz = 14

Dim m_TextBorderColor As OLE_COLOR
Dim m_TextColor As OLE_COLOR
Dim m_Text As String
Dim m_TextStyle As eStyle
Dim m_TextSz As Integer

Public Enum eStyle
   Embossed = 0
   Engraved = 1
   Outline = 2
   ThreeD_Text = 3
   Shadow = 4
   Horz_Grad_OutLine = 5
   Horz_Grad_Embossed = 6
   Vert_Grad_Engraved = 7
   Horz_Grad_Engraved = 8
   Vert_Grad_OutLine = 9
   Vert_Grad_Shadow = 10
   Horz_Grad_Shadow = 11
End Enum
'------------------------------------------------
'Constant Type's Used with DrawText

Private Const DT_TOP = &H0
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_BOTTOM = &H8
Private Const DT_MULTILINE = (&H1)
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10

'API Functions

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function TextOutA Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function InflateRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawTextW Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal lpRect As Long) As Long
Private Declare Function DrawTextA Lib "user32.dll" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal lpRect As Long) As Long

Private Type RECT
    Left  As Long
    Top   As Long
    Right As Long
    Bottom As Long
End Type

'Needed For The API CreateFontIndirect Function Calls
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * 32
End Type

'Private Variable For The API SetPixel Function
Dim x, y  As Single
'Private Variable For The API DrawText Function
Dim r     As RECT

Private Sub UserControl_Initialize()
   
 With UserControl
        .ScaleMode = vbPixels    'For API Functions
        .Width = 5500
        .Height = 7000
        .FontName = "Times New Roman"
        .FontSize = 14
        .FontBold = True
    End With
    m_TextColor = m_def_TextColor
    m_TextBorderColor = m_def_TextBorderColor
    m_TextStyle = m_def_TextStyle
End Sub

Private Sub UserControl_InitProperties()
Text = Extender.name
TextStyle = m_TextStyle
TextColor = m_TextColor
TextBorderColor = m_TextBorderColor
TextSize = m_def_TextSz
End Sub

Private Sub Effect1()
'Embossed_Effect
    With UserControl
      .Cls
        With r
            .Top = 0
            .Left = 0
            .Bottom = ScaleHeight
            .Right = ScaleWidth
        End With
   
        SetTextColor .hdc, TextBorderColor
        DrawTextA .hdc, Text, Len(Text), r, DT_LEFT Or DT_TOP

        OffsetRect r, 1, 1
        SetTextColor .hdc, &H808080
        DrawTextW .hdc, Text, Len(Text), r, DT_LEFT Or DT_TOP

        InflateRect r, 0, 0
        SetTextColor .hdc, TextColor
        DrawTextW .hdc, Text, Len(Text), r, DT_LEFT Or DT_TOP

    End With
End Sub
Private Sub Effect2()
'Engraved_text
    With UserControl
        .Cls
        With r
            .Top = 2
            .Left = 2
            .Bottom = ScaleHeight
            .Right = ScaleWidth
        End With

        SetTextColor .hdc, TextBorderColor
        DrawTextA .hdc, Text, Len(Text), r, DT_LEFT Or DT_TOP

        InflateRect r, 2, 1
        SetTextColor .hdc, &H808080
        DrawTextW .hdc, Text, Len(Text), r, DT_LEFT Or DT_TOP

        OffsetRect r, 1, 0
        SetTextColor .hdc, TextColor
        DrawTextW .hdc, Text, Len(Text), r, DT_LEFT Or DT_TOP

    End With
End Sub

Private Sub Effect3()
'OutLine Text Efeect
   
    With UserControl
        .Cls
        SetRect r, 1, 0, .ScaleWidth, .ScaleHeight
        For x = -1 To 1
            For y = -1 To 1
                InflateRect r, x, y
                SetTextColor .hdc, TextBorderColor
                DrawTextW .hdc, Text, Len(Text), r, DT_LEFT Or DT_TOP

                OffsetRect r, x, y
                SetTextColor .hdc, TextColor
                DrawTextA .hdc, Text, Len(Text), r, DT_LEFT Or DT_TOP
            Next y
        Next x
    End With
        
End Sub

Private Sub Effect4()
'3D Text Efeect
   Dim i As Integer
   Dim Xx As Single
   Dim Yy As Single
   Dim Red As Integer
   Dim Grn As Integer
   Dim Blu As Integer
   Dim RChange As Integer
   Dim GChange As Integer
   Dim BChange As Integer
   Dim SRed As Integer
   Dim SGreen As Integer
   Dim SBlue As Integer
   Dim ERed As Integer
   Dim EGreen As Integer
   Dim EBlue As Integer
   
   On Error Resume Next
   UserControl.Cls
   
   SRed = TextColor Mod 256
   SGreen = (TextColor And vbGreen) / 256
   SBlue = (TextColor And vbBlue) / 65536
   ERed = TextBorderColor Mod 256
   EGreen = (TextBorderColor And vbGreen) / 256
   EBlue = (TextBorderColor And vbBlue) / 65536
  
     For i = 0 To 254
         Xx = Xx + 0.06
         Yy = Yy + 0.06
         UserControl.CurrentX = Xx
         UserControl.CurrentY = Yy
         RChange = RChange + (ERed - SRed) / 255                'start of gradient colors
         GChange = GChange + (EGreen - SGreen) / 255
         BChange = BChange + (EBlue - SBlue) / 255
         Red = SRed + RChange
         Grn = SGreen + GChange
         Blu = SBlue + BChange
         UserControl.ForeColor = RGB(Red, Grn, Blu)                   'set text color
      
      If i >= 220 And i <= 249 Then UserControl.ForeColor = vbBlack
      If i >= 250 Then
          UserControl.ForeColor = TextColor      'highlights start text
          UserControl.CurrentX = Xx - 1.5
          UserControl.CurrentY = Yy - 1.5
      End If
      UserControl.Print Text
      Next i
      UserControl.MaskPicture = UserControl.Image
     
End Sub

Private Sub Effect5()
'Shadow Text Effect
    With UserControl
        .Cls
        SetRect r, 10, 1, .ScaleWidth, .ScaleHeight
        For x = 0 To 3
            For y = 0 To 3
                x = x + 2
                y = y + 2
                InflateRect r, x, y
                SetTextColor .hdc, &HC0C0C0
                DrawTextW .hdc, Text, Len(Text), r, DT_LEFT Or DT_TOP
            Next y
        Next x
        SetTextColor .hdc, TextColor
        DrawTextA .hdc, Text, Len(Text), r, DT_LEFT Or DT_TOP
    End With
End Sub
      
Private Sub Effect6()
   'Vertical Gradient Text+Embossed Text Effect
    With UserControl
        .Cls
        SetRect r, 3, 0, .ScaleWidth, .ScaleHeight
        For x = -1 To 1
            For y = -1 To 1
                InflateRect r, x, y
                SetTextColor .hdc, TextBorderColor
                DrawTextW .hdc, Text, Len(Text), r, DT_LEFT Or DT_TOP

                OffsetRect r, x, y
                SetTextColor .hdc, vbBlack
                DrawTextA .hdc, Text, Len(Text), r, DT_LEFT Or DT_TOP
            Next y
        Next x

        For x = 0 To .ScaleWidth
            For y = 0 To .ScaleHeight
                If GetPixel(.hdc, x, y) = &H0 Then
                    SetPixel .hdc, x, y, RGB(255 - x, 0, x)
                End If
            Next y
        Next x

    End With
End Sub

Private Sub Effect7()
   'Horizontal Gradient Text+Embossed Text Effect
    With UserControl
        .Cls
        SetRect r, 0, 0, .ScaleWidth, .ScaleHeight

        SetTextColor .hdc, TextBorderColor
        DrawTextA .hdc, Text, Len(Text), r, DT_LEFT Or DT_TOP

        OffsetRect r, 2, 2
        SetTextColor .hdc, &H808080
        DrawTextW .hdc, Text, Len(Text), r, DT_LEFT Or DT_TOP

        InflateRect r, 1, 1
        SetTextColor .hdc, vbBlack
        DrawTextW .hdc, Text, Len(Text), r, DT_LEFT Or DT_TOP

        For x = 0 To .ScaleWidth
            For y = 0 To .ScaleHeight
                If GetPixel(.hdc, x, y) = &H0 Then
                    SetPixel .hdc, x, y, RGB(x, x, 2 * x)

                End If
            Next y
        Next x
    End With
End Sub

Private Sub Effect8()
   'Vertical Gradient Text+Engraved Text Effect
   With UserControl
        .Cls
        SetRect r, 0, 0, .ScaleWidth, .ScaleHeight

        SetTextColor .hdc, TextBorderColor
        DrawTextA .hdc, Text, Len(Text), r, DT_LEFT Or DT_TOP

        InflateRect r, 2, 1
        SetTextColor .hdc, &H808080
        DrawTextW .hdc, Text, Len(Text), r, DT_LEFT Or DT_TOP

        OffsetRect r, 1, 0
        SetTextColor .hdc, vbBlack
        DrawTextW .hdc, Text, Len(Text), r, DT_LEFT Or DT_TOP

        For x = 0 To .ScaleWidth
            For y = 0 To .ScaleHeight
                If GetPixel(.hdc, x, y) = &H0 Then
                    SetPixel .hdc, x, y, RGB(12 * y, 0, 0)
                End If
            Next y
        Next x
    End With
End Sub

Private Sub Effect9()
   'Horizontal Gradient Text+Engraved Text Effect
   With UserControl
        .Cls
        SetRect r, 0, 0, .ScaleWidth, .ScaleHeight

        SetTextColor .hdc, TextBorderColor
        DrawTextA .hdc, Text, Len(Text), r, DT_LEFT Or DT_TOP

        InflateRect r, 2, 1
        SetTextColor .hdc, &H808080
        DrawTextW .hdc, Text, Len(Text), r, DT_LEFT Or DT_TOP

        OffsetRect r, 1, 0
        SetTextColor .hdc, vbBlack
        DrawTextW .hdc, Text, Len(Text), r, DT_LEFT Or DT_TOP

        For x = 0 To .ScaleWidth
            For y = 0 To .ScaleHeight
                If GetPixel(.hdc, x, y) = &H0 Then
                    SetPixel .hdc, x, y, RGB(x, x, x)
                End If
            Next y
        Next x
    End With
End Sub

Private Sub Effect10()
   'Vertical Gradient Text+OutLineText Effect
    With UserControl
        .Cls
        SetRect r, 0, 0, .ScaleWidth, .ScaleHeight

        For x = -1 To 1
            For y = -1 To 1
                InflateRect r, x, y
                SetTextColor .hdc, TextBorderColor
                DrawTextW .hdc, Text, Len(Text), r, DT_LEFT Or DT_TOP

                OffsetRect r, x, y
                SetTextColor .hdc, vbBlack
                DrawTextA .hdc, Text, Len(Text), r, DT_LEFT Or DT_TOP
            Next y
        Next x

        For x = 0 To .ScaleWidth
            For y = 0 To .ScaleHeight
                If GetPixel(.hdc, x, y) = &H0 Then
                    SetPixel .hdc, x, y, RGB(y, 0, 10 * y)
                End If
            Next y
        Next x

    End With
End Sub

Private Sub Effect11()
   'Vertical Gradient Text+Shadow Text Effect
    With UserControl
        .Cls
        SetRect r, 6, 3, .ScaleWidth, .ScaleHeight
        For x = 0 To 3
            For y = 0 To 3
                x = x + 2
                y = y + 2
                InflateRect r, x, y
                SetTextColor .hdc, &H808080
                DrawTextW .hdc, Text, Len(Text), r, DT_LEFT Or DT_TOP
            Next y
        Next x
        SetTextColor .hdc, vbBlack
        DrawTextA .hdc, Text, Len(Text), r, DT_LEFT Or DT_TOP

        For x = 0 To .ScaleWidth
            For y = 0 To .ScaleHeight
                If GetPixel(.hdc, x, y) = &H0 Then
                    SetPixel .hdc, x, y, RGB(15 * y, 30, 255)
                End If
            Next y
        Next x
    End With
End Sub

Private Sub Effect12()
   'Horizontal Gradient Text+Shadow Text Effect
    With UserControl
        .Cls
        SetRect r, 6, 3, .ScaleWidth, .ScaleHeight
        For x = 0 To 3
            For y = 0 To 3
                x = x + 2
                y = y + 2
                InflateRect r, x, y
                SetTextColor .hdc, &H808080
                DrawTextW .hdc, Text, Len(Text), r, DT_LEFT Or DT_TOP
            Next y
        Next x
        SetTextColor .hdc, vbBlack
        DrawTextA .hdc, Text, Len(Text), r, DT_LEFT Or DT_TOP

        For x = 0 To .ScaleWidth
            For y = 0 To .ScaleHeight
                If GetPixel(.hdc, x, y) = &H0 Then
                    SetPixel .hdc, x, y, RGB(255 - x, 200 - x, x)
                End If
            Next y
        Next x
    End With
End Sub

Public Property Get TextSize()
    TextSize = m_TextSz
End Property

Public Property Let TextSize(NewTextSz)
   m_TextSz = NewTextSz
   UserControl.FontSize = m_TextSz
   PropertyChanged "TextSize"
   SelEffect
End Property
Public Property Get TextStyle() As eStyle
   TextStyle = m_TextStyle
End Property

Public Property Let TextStyle(NewTextStyle As eStyle)
   m_TextStyle = NewTextStyle
   PropertyChanged "TextStyle"
   SelEffect
End Property

Public Property Get Text() As String
   Text = m_Text
End Property

Public Property Let Text(NewText As String)
   m_Text = NewText
   PropertyChanged "Text"
   SelEffect
End Property

Public Property Get TextBorderColor() As OLE_COLOR
   TextBorderColor = m_TextBorderColor
End Property

Public Property Let TextBorderColor(NewTextBorderColor As OLE_COLOR)
   m_TextBorderColor = NewTextBorderColor
   PropertyChanged "TextBorderColor"
   SelEffect
End Property

Public Property Get TextColor() As OLE_COLOR
   TextColor = m_TextColor
End Property

Public Property Let TextColor(NewTextColor As OLE_COLOR)
   m_TextColor = NewTextColor
   PropertyChanged "TextColor"
   SelEffect
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   TextStyle = PropBag.ReadProperty("TextStyle", m_def_TextStyle)
   Text = PropBag.ReadProperty("Text", Extender.name)
   TextBorderColor = PropBag.ReadProperty("TextBorderColor", m_def_TextBorderColor)
   TextColor = PropBag.ReadProperty("TextColor", m_def_TextColor)
   TextSize = PropBag.ReadProperty("TextSize", m_def_TextSz)
End Sub

Private Sub UserControl_Resize()
    SelEffect
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   With PropBag
   Call .WriteProperty("TextStyle", m_TextStyle, m_def_TextStyle)
   Call .WriteProperty("Text", m_Text, Extender.name)
   Call .WriteProperty("TextBorderColor", m_TextBorderColor, m_def_TextBorderColor)
   Call .WriteProperty("TextColor", m_TextColor, m_def_TextColor)
   Call .WriteProperty("TextSize", m_TextSz, m_def_TextSz)
   End With
End Sub

Private Sub SelEffect()
   Select Case TextStyle
      Case 0: Call Effect1
      Case 1: Call Effect2
      Case 2: Call Effect3
      Case 3: Call Effect4
      Case 4: Call Effect5
      Case 5: Call Effect6
      Case 6: Call Effect7
      Case 7: Call Effect8
      Case 8: Call Effect9
      Case 9: Call Effect10
      Case 10: Call Effect11
      Case 11: Call Effect12
   End Select
    UserControl.MaskPicture = UserControl.Image
End Sub
