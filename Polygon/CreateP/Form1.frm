VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Polygons"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   471
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   449
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4395
      TabIndex        =   18
      Text            =   "Picture1"
      Top             =   4785
      Width           =   1890
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Fill Color"
      Height          =   435
      Left            =   5130
      TabIndex        =   15
      Top             =   1680
      Width           =   1215
   End
   Begin VB.HScrollBar HS1 
      Height          =   255
      Left            =   135
      TabIndex        =   14
      Top             =   3150
      Width           =   3060
   End
   Begin VB.VScrollBar VS1 
      Height          =   2985
      Left            =   3225
      Max             =   1000
      TabIndex        =   13
      Top             =   135
      Width           =   225
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   135
      ScaleHeight     =   198
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   198
      TabIndex        =   11
      Top             =   120
      Width           =   3000
      Begin VB.PictureBox pic1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   3000
         Left            =   -30
         ScaleHeight     =   196
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   196
         TabIndex        =   12
         Top             =   -30
         Width           =   3000
      End
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open Image"
      Height          =   435
      Left            =   5130
      TabIndex        =   10
      Top             =   1170
      Width           =   1215
   End
   Begin VB.CommandButton cmdMag 
      Caption         =   "Magnify"
      Height          =   420
      Left            =   5130
      TabIndex        =   9
      Top             =   675
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   195
      Left            =   3390
      TabIndex        =   7
      Top             =   4065
      Width           =   195
   End
   Begin VB.TextBox arName 
      Height          =   285
      Left            =   4230
      TabIndex        =   5
      Text            =   "Position"
      Top             =   4365
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   3570
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3435
      Width           =   3090
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   450
      Left            =   5130
      TabIndex        =   2
      Top             =   150
      Width           =   1230
   End
   Begin VB.CommandButton cmdClipBoard 
      Caption         =   "Clipboard"
      Height          =   495
      Left            =   5115
      TabIndex        =   1
      Top             =   2490
      Width           =   1185
   End
   Begin VB.ComboBox cbo1 
      Height          =   315
      Left            =   3495
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   135
      Width           =   1590
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "(Where poly is to be drawn)"
      Height          =   255
      Left            =   4365
      TabIndex        =   19
      Top             =   5115
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Objects Name"
      Height          =   225
      Left            =   3345
      TabIndex        =   17
      Top             =   4830
      Width           =   1050
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   6375
      TabIndex        =   16
      Top             =   1680
      Width           =   330
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Include code to create the array"
      Height          =   210
      Left            =   3660
      TabIndex        =   8
      Top             =   4080
      Width           =   2670
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Array Name"
      Height          =   210
      Left            =   3345
      TabIndex        =   6
      Top             =   4440
      Width           =   885
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Click to set new Line                        Doubleclick to close the polygon"
      Height          =   480
      Left            =   3630
      TabIndex        =   4
      Top             =   6375
      Width           =   2520
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Private Type PointAPI
   X As Long
   Y As Long
End Type

Private Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As PointAPI, ByVal nCount As Long) As Long
Private Position() As PointAPI
Private Polyctr As Long
Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private strfileName As OPENFILENAME
Public FileSelected As Boolean

Private Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Dim MyFile As String
Dim CustomColors() As Byte

Private Sub Form_Load()
    Polyctr = -1
    cbo1.AddItem "16 x 16"
    cbo1.AddItem "32 x 32"
    cbo1.AddItem "100 x 100"
    cbo1.AddItem "150 x 150"
    cbo1.AddItem "200 x 200"
    cbo1.ListIndex = 4
    
    ReDim CustomColors(0 To 16 * 4 - 1) As Byte
    Dim i As Integer
    For i = LBound(CustomColors) To UBound(CustomColors)
        CustomColors(i) = 0
    Next i
    Label4.BackColor = pic1.FillColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload frmZoom
   Unload Me
End Sub

Private Sub Form_Terminate()
    Unload frmZoom
    Unload Me
End Sub

Private Sub cbo1_Click()
    Select Case cbo1.ListIndex
        Case 0
            picMain.Width = 16
            picMain.Height = 16
        Case 1
            picMain.Width = 32
            picMain.Height = 32
        Case 2
            picMain.Width = 100
            picMain.Height = 100
        Case 3
            picMain.Width = 150
            picMain.Height = 150
        Case 4
            picMain.Width = 200
            picMain.Height = 200
    End Select
    Text1.Text = ""
   
    VS1.Max = pic1.ScaleHeight - picMain.Height
    HS1.Max = pic1.ScaleWidth - picMain.Width
End Sub

Private Sub HS1_Change()
   pic1.Left = -HS1.Value
End Sub

Private Sub HS1_Scroll()
   pic1.Left = -HS1.Value
End Sub

Private Sub VS1_Change()
   pic1.Top = -VS1.Value
End Sub

Private Sub VS1_Scroll()
   pic1.Top = -VS1.Value
End Sub

Private Sub pic1_DblClick()
    Dim i As Integer
    If arName.Text = "" Then Exit Sub
    If Polyctr < 3 Then
      ' cmdClear_Click
       Exit Sub
    End If
    pic1.AutoRedraw = True
    Polygon pic1.hDC, Position(0), Polyctr
    pic1.AutoRedraw = False
    pic1.Refresh
    
    If Check1.Value = Checked Then
       Text1.Text = Text1.Text & " ' Place in Declares Section" & vbCrLf
       Text1.Text = Text1.Text & "Private Type POINTAPI" & vbCrLf
       Text1.Text = Text1.Text & "   x As Long" & vbCrLf
       Text1.Text = Text1.Text & "   y As Long" & vbCrLf
       Text1.Text = Text1.Text & "End Type" & vbCrLf
       Text1.Text = Text1.Text & "Private " & arName.Text & "() As POINTAPI" & vbCrLf
       Text1.Text = Text1.Text & "' End Declares Code" & vbCrLf
       Text1.Text = Text1.Text & vbCrLf
       Check1.Value = Unchecked
    End If
    
    If Text2.Text <> "" Then
       Text1.Text = Text1.Text & Text2.Text & ".FillStyle = 0" & vbCrLf
       Text1.Text = Text1.Text & Text2.Text & ".FillColor = " & Label4.BackColor & vbCrLf
    End If
    
    Text1.Text = Text1.Text & "Polyctr = " & Polyctr & vbCrLf
    Text1.Text = Text1.Text & "ReDim " & arName.Text & "(" & Polyctr & ")" & vbCrLf
    For i = 0 To Polyctr
        Text1.Text = Text1.Text & arName.Text & "(" & i & ").x = " & Position(i).X & vbCrLf
        Text1.Text = Text1.Text & arName.Text & "(" & i & ").y = " & Position(i).Y & vbCrLf
    Next i
    Text1.Text = Text1.Text & vbCrLf
    Polyctr = -1
    Unload frmZoom
End Sub

Private Sub pic1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim i As Long
    
    If Button = 2 Then
        Polyctr = -1
        pic1.Cls
    Else
        Polyctr = Polyctr + 1
        ReDim Preserve Position(Polyctr)
        Position(Polyctr).X = X
        Position(Polyctr).Y = Y
        If Polyctr > 0 Then
            For i = 1 To Polyctr
                pic1.Line (Position(i - 1).X, Position(i - 1).Y)-(Position(i).X, Position(i).Y)
            Next i
        End If
    End If
    
End Sub

Private Sub pic1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim i As Long
    
    If Polyctr > -1 Then
        pic1.Cls
        For i = 1 To Polyctr
            pic1.Line (Position(i - 1).X, Position(i - 1).Y)-(Position(i).X, Position(i).Y)
        Next i
        pic1.DrawMode = 6
        pic1.Line (Position(Polyctr).X, Position(Polyctr).Y)-(X, Y)
        pic1.DrawMode = 13
    End If
    
End Sub

Private Sub cmdMag_Click()
    frmZoom.Left = Form1.Left + Form1.Width + 50
    frmZoom.Top = Form1.Top + 1000
    frmZoom.Show
End Sub

Private Sub cmdClear_Click()
    pic1.Picture = LoadPicture()
    Polyctr = -1
    Text1.Text = ""
End Sub

Private Sub cmdOpen_Click()
   MyFile = fncGetFileNametoOpen("Open Any File", "All Files|*.*")
   If MyFile = "" Then Exit Sub
   pic1.Picture = LoadPicture(MyFile)
   
   VS1.Max = pic1.ScaleHeight - picMain.Height
   HS1.Max = pic1.ScaleWidth - picMain.Width
End Sub

Private Sub cmdClipBoard_Click()
    Clipboard.Clear
    Clipboard.SetText Text1.Text
End Sub

Private Sub Command4_Click()
    Dim NewColor As Long

    NewColor = ShowColor

    If NewColor <> -1 Then
        pic1.FillColor = NewColor
        Label4.BackColor = NewColor
    Else
        'MsgBox "You chose cancel"
    End If
End Sub

Private Sub DialogFilter(WantedFilter As String)
    Dim intLoopCount As Integer
    strfileName.lpstrFilter = ""
    For intLoopCount = 1 To Len(WantedFilter)
        If Mid(WantedFilter, intLoopCount, 1) = "|" Then strfileName.lpstrFilter = _
        strfileName.lpstrFilter + Chr(0) Else strfileName.lpstrFilter = _
        strfileName.lpstrFilter + Mid(WantedFilter, intLoopCount, 1)
    Next intLoopCount
    strfileName.lpstrFilter = strfileName.lpstrFilter + Chr(0)
End Sub

Private Function fncGetFileNametoOpen(Optional strDialogTitle As String = "Open", Optional strFilter As String = "All Files|*.*", Optional strDefaultExtention As String = "*.*") As String
    Dim lngReturnValue As Long
    Dim intRest As Integer
    strfileName.lpstrTitle = strDialogTitle
    strfileName.lpstrDefExt = strDefaultExtention
    DialogFilter (strFilter)
    strfileName.hInstance = App.hInstance
    strfileName.lpstrFile = Chr(0) & Space(259) ' --> will return Chr(0) & 259 spaces UNLESS a valid file is selected.
    strfileName.nMaxFile = 260 ' maximum length of a file name
    strfileName.flags = &H4
    strfileName.lStructSize = Len(strfileName)
    lngReturnValue = GetOpenFileName(strfileName)
    FileSelected = lngReturnValue ' --> CptnVics addition... must be done after the call to GetOpenFileName(strfileName)!
        'FileSelected will coerce this value (lngReturnValue) to boolean... true if a file was selected... false otherwise.
        'FileSelected could be dimensioned as a string... in which case it would return "1" if a file was selected... "0" if canceled
        'The boolean check takes less code... see the demo form.
    If FileSelected = True Then
       fncGetFileNametoOpen = strfileName.lpstrFile
    Else
       fncGetFileNametoOpen = ""
    End If
End Function

Private Function ShowColor() As Long
    Dim cc As CHOOSECOLOR
    'Dim Custcolor(16) As Long
    'Dim lReturn As Long

    'set the structure size
    cc.lStructSize = Len(cc)
    'Set the owner
    cc.hwndOwner = Me.hwnd
    'set the application's instance
    cc.hInstance = App.hInstance
    'set the custom colors (converted to Unicode)
    cc.lpCustColors = StrConv(CustomColors, vbUnicode)
    'no extra flags
    cc.flags = 0  'set to 0 = define custom colors unselected. 2= define custom colors selected
    
    'Show the 'Select Color'-dialog
    If CHOOSECOLOR(cc) <> 0 Then
        ShowColor = cc.rgbResult
        CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
    Else
        ShowColor = -1
    End If
End Function
