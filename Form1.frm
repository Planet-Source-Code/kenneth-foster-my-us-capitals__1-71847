VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "US Capitals"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13140
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0444
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   876
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkPlay 
      BackColor       =   &H00000040&
      Caption         =   "Play  Music"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   7740
      TabIndex        =   20
      Top             =   6765
      Width           =   1425
   End
   Begin VB.HScrollBar hsbVolume 
      Height          =   240
      Left            =   7830
      Max             =   1000
      Min             =   -1000
      TabIndex        =   19
      Top             =   7200
      Width           =   2595
   End
   Begin Project1.TextEffects TextEffects9 
      Height          =   1125
      Left            =   2250
      TabIndex        =   18
      Top             =   2760
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   1984
      TextStyle       =   2
      Text            =   "Press Start to Begin"
      TextBorderColor =   16777215
      TextSize        =   40
   End
   Begin Project1.TextEffects TextEffects8 
      Height          =   1080
      Left            =   5355
      TabIndex        =   17
      Top             =   6765
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   1905
      TextStyle       =   2
      Text            =   "Capitals"
      TextBorderColor =   16777215
      TextColor       =   32768
      TextSize        =   30
   End
   Begin Project1.TextEffects TextEffects7 
      Height          =   930
      Left            =   3375
      TabIndex        =   16
      Top             =   6795
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   1640
      TextStyle       =   2
      Text            =   "States"
      TextBorderColor =   16777215
      TextSize        =   30
   End
   Begin Project1.TextEffects TextEffects6 
      Height          =   870
      Left            =   1155
      TabIndex        =   15
      Top             =   6780
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   1535
      TextStyle       =   2
      Text            =   "United"
      TextBorderColor =   16777215
      TextColor       =   255
      TextSize        =   30
   End
   Begin Project1.TextEffects TextEffects5 
      Height          =   45
      Left            =   2715
      TabIndex        =   14
      Top             =   6840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   79
   End
   Begin Project1.TextEffects TextEffects4 
      Height          =   60
      Left            =   2910
      TabIndex        =   13
      Top             =   6810
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   106
   End
   Begin Project1.ccXPButton cmdExit 
      Height          =   360
      Left            =   11340
      TabIndex        =   12
      Top             =   6225
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   635
      Caption         =   "EXIT"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.TextEffects TE1 
      Height          =   1425
      Left            =   3315
      TabIndex        =   9
      Top             =   2565
      Visible         =   0   'False
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   2514
      TextStyle       =   2
      Text            =   "Fantastic"
      TextBorderColor =   16777215
      TextColor       =   12583104
      TextSize        =   60
   End
   Begin Project1.TextEffects TextEffects3 
      Height          =   285
      Left            =   12120
      TabIndex        =   8
      Top             =   7050
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   503
      Text            =   "Missed"
      TextBorderColor =   16777215
      TextColor       =   255
   End
   Begin Project1.TextEffects TextEffects2 
      Height          =   285
      Left            =   10590
      TabIndex        =   7
      Top             =   7050
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   503
      Text            =   "Correct"
      TextBorderColor =   16777215
   End
   Begin VB.PictureBox picStor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2865
      Left            =   4470
      ScaleHeight     =   191
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   182
      TabIndex        =   6
      Top             =   8025
      Width           =   2730
   End
   Begin Project1.ccXPButton cmdNewQuiz 
      Height          =   480
      Left            =   11805
      TabIndex        =   5
      Top             =   5550
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   847
      Caption         =   "New Quiz"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.ccXPButton cmdStart 
      Height          =   480
      Left            =   10815
      TabIndex        =   4
      Top             =   5550
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   847
      Caption         =   "START"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox lstCaps 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   10830
      TabIndex        =   2
      Top             =   4365
      Width           =   1950
   End
   Begin Project1.TextEffects TextEffects1 
      Height          =   285
      Left            =   11070
      TabIndex        =   1
      Top             =   0
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   503
      TextStyle       =   2
      Text            =   "Missed List"
      TextBorderColor =   16777215
      TextColor       =   255
   End
   Begin VB.ListBox lstMissed 
      Appearance      =   0  'Flat
      Height          =   3345
      Left            =   10485
      TabIndex        =   0
      Top             =   285
      Width           =   2640
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(via DirectX 7)"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   9165
      TabIndex        =   22
      Top             =   6765
      Width           =   1290
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Volume"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   8775
      TabIndex        =   21
      Top             =   7005
      Width           =   675
   End
   Begin VB.Label lblScoreCor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   10860
      TabIndex        =   10
      Top             =   6795
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   255
      Left            =   10860
      Top             =   6780
      Width           =   495
   End
   Begin VB.Label lblScoreInc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   12285
      TabIndex        =   11
      Top             =   6795
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   540
      Left            =   11775
      Shape           =   4  'Rounded Rectangle
      Top             =   5520
      Width           =   1050
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   11310
      Shape           =   4  'Rounded Rectangle
      Top             =   6180
      Width           =   915
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   570
      Left            =   10770
      Shape           =   4  'Rounded Rectangle
      Top             =   5505
      Width           =   990
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000005&
      Height          =   270
      Left            =   12285
      Top             =   6810
      Width           =   480
   End
   Begin VB.Line Line1 
      X1              =   699
      X2              =   699
      Y1              =   0
      Y2              =   454
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   1515
      Left            =   10785
      Top             =   3900
      Width           =   2040
   End
   Begin VB.Label lblState 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   10815
      TabIndex        =   3
      Top             =   3945
      Width           =   1965
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H000000FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   690
      Left            =   10560
      Top             =   6705
      Width           =   2520
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Ken Foster Mar 2009
'Free to use or abuse
'The random numbers code is not mine.
'DirectX sound code also is borrowed

Option Explicit
'Note: If using sound in another app,make sure to reference DirectX 7
   Dim aryStateData(49) As String
   Dim aryCapData(49) As String
   Dim aryImgData(49) As String
   
   Dim v_dx As New DirectX7
   Dim v_dmp As DirectMusicPerformance
   Dim v_dml As DirectMusicLoader
   Dim v_dms As DirectMusicSegment
   Dim v_dmss As DirectMusicSegmentState

   Dim vs_filename As String
   Dim vl_volume As Long

   'mix up states
   Dim a(49)
   Dim B(49)
   Dim P As Integer
   Dim w As Integer
   Dim x As Integer
   Dim y As Integer
   Dim mPlay As Boolean
   Dim RanNo(4) As Long   'mix up capital choices
   
   Private Declare Function TransparentBlt Lib "msimg32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Long

Private Sub Form_Load()
   LoadDataArray
   Array_Fill
   cmdNewQuiz.Visible = False
   Shape6.Visible = False
   cmdStart.Caption = "Start"
   Me.Show
   cmdStart.SetFocus
   Shape4.Visible = True
   w = 0
   
    On Local Error GoTo ErrSub
    Set v_dml = v_dx.DirectMusicLoaderCreate
    Set v_dmp = v_dx.DirectMusicPerformanceCreate
    
    Call v_dmp.Init(Nothing, hWnd)
    Call v_dmp.SetPort(-1, 1)
    vs_filename = App.Path & "\Jeopardy.mid"
    Exit Sub
ErrSub:
    Call ErrMess(Err.Number, Err.Description)
End Sub

Private Sub cmdStart_Click()
   Make_Quiz
   cmdStart.Visible = False
   Shape4.Visible = False
   cmdNewQuiz.Visible = True
   Shape6.Visible = True
   Beep
   TextEffects9.Visible = False
End Sub

Private Sub cmdExit_Click()
   Call v_dmp.Stop(v_dms, v_dmss, 0, 0)
   If mPlay = True Then Call v_dms.Unload(v_dmp)
   Unload Me
End Sub

Private Sub cmdNewQuiz_Click()
   
   Dim z As Integer
   ClearAll
   cmdStart.Enabled = True
   w = 0
   x = 0
   y = 0
   Array_Fill
   lblScoreCor.Caption = "0"
   lblScoreInc.Caption = "0"
   lblState.Caption = ""
   lstCaps.Clear
   cmdStart.Visible = True
   Shape4.Visible = True
   TE1.Visible = False
   Beep
   TextEffects9.Visible = True
End Sub

Private Sub ChkAnswer()
   
   Select Case lblState.Caption
      Case "Alabama"
         If lstCaps.Text = "Montgomery" Then
            LoadStatePic 0, "B", 454, 279
         Else
            LoadStatePic 0, "R", 454, 279
            lstMissed.AddItem "Alabama - " & aryCapData(0)
         End If
      Case "Alaska"
         If lstCaps.Text = "Juneau" Then
            LoadStatePic 1, "B", 49, 363
         Else
            LoadStatePic 1, "R", 49, 363
            lstMissed.AddItem "Alaska - " & aryCapData(1)
         End If
      Case "Arizonia"
         If lstCaps.Text = "Phoenix" Then
            LoadStatePic 2, "B", 109, 229
         Else
            LoadStatePic 2, "R", 109, 229
            lstMissed.AddItem "Arizonia - " & aryCapData(2)
         End If
      Case "Arkansas"
         If lstCaps.Text = "Little Rock" Then
            LoadStatePic 3, "B", 371, 259
         Else
            LoadStatePic 3, "R", 371, 259
            lstMissed.AddItem "Arkansas - " & aryCapData(3)
         End If
      Case "California"
         If lstCaps.Text = "Sacramento" Then
            LoadStatePic 4, "B", 25, 123
         Else
            LoadStatePic 4, "R", 25, 123
            lstMissed.AddItem "California - " & aryCapData(4)
         End If
      Case "Colorado"
         If lstCaps.Text = "Denver" Then
            LoadStatePic 5, "B", 192, 178
         Else
            LoadStatePic 5, "R", 192, 178
            lstMissed.AddItem "Colorado - " & aryCapData(5)
         End If
      Case "Connecticut"
         If lstCaps.Text = "Hartford" Then
            LoadStatePic 6, "B", 613, 141
         Else
            LoadStatePic 6, "R", 613, 141
            lstMissed.AddItem "Connecticut - " & aryCapData(6)
         End If
      Case "Delaware"
         If lstCaps.Text = "Dover" Then
            LoadStatePic 7, "B", 593, 185
         Else
            LoadStatePic 7, "R", 593, 185
            lstMissed.AddItem "Delaware - " & aryCapData(7)
         End If
      Case "Florida"
         If lstCaps.Text = "Tallahassee" Then
            LoadStatePic 8, "B", 465, 338
         Else
            LoadStatePic 8, "R", 465, 338
            lstMissed.AddItem "Florida - " & aryCapData(8)
         End If
      Case "Georgia"
         If lstCaps.Text = "Atlanta" Then
            LoadStatePic 9, "B", 485, 275
         Else
            LoadStatePic 9, "R", 485, 275
            lstMissed.AddItem "Georgia - " & aryCapData(9)
         End If
      Case "Hawaii"
         If lstCaps.Text = "Honolulu" Then
            LoadStatePic 10, "B", 159, 401
         Else
            LoadStatePic 10, "R", 159, 401
            lstMissed.AddItem "Hawaii - " & aryCapData(10)
         End If
      Case "Idaho"
         If lstCaps.Text = "Boise" Then
            LoadStatePic 11, "B", 112, 39
         Else
            LoadStatePic 11, "R", 112, 39
            lstMissed.AddItem "Idaho - " & aryCapData(11)
         End If
      Case "Illinois"
         If lstCaps.Text = "Springfield" Then
            LoadStatePic 12, "B", 407, 164
         Else
            LoadStatePic 12, "R", 407, 164
            lstMissed.AddItem "Illinois - " & aryCapData(12)
         End If
      Case "Indiana"
         If lstCaps.Text = "Indianapolis" Then
            LoadStatePic 13, "B", 452, 172
         Else
            LoadStatePic 13, "R", 452, 172
            lstMissed.AddItem "Indiana - " & aryCapData(13)
         End If
      Case "Iowa"
         If lstCaps.Text = "Des Moines" Then
            LoadStatePic 14, "B", 347, 150
         Else
            LoadStatePic 14, "R", 347, 150
            lstMissed.AddItem "Iowa - " & aryCapData(14)
         End If
      Case "Kansas"
         If lstCaps.Text = "Topeka" Then
            LoadStatePic 15, "B", 279, 202
         Else
            LoadStatePic 15, "R", 279, 202
            lstMissed.AddItem "Kansas - " & aryCapData(15)
         End If
      Case "Kentucky"
         If lstCaps.Text = "Frankfort" Then
            LoadStatePic 16, "B", 437, 213
         Else
            LoadStatePic 16, "R", 437, 213
            lstMissed.AddItem "Kentucky - " & aryCapData(16)
         End If
      Case "Louisiana"
         If lstCaps.Text = "Baton Rouge" Then
            LoadStatePic 17, "B", 378, 315
         Else
            LoadStatePic 17, "R", 378, 315
            lstMissed.AddItem "Louisiana - " & aryCapData(17)
         End If
      Case "Maine"
         If lstCaps.Text = "Augusta" Then
            LoadStatePic 18, "B", 627, 51
         Else
            LoadStatePic 18, "R", 627, 51
            lstMissed.AddItem "Maine - " & aryCapData(18)
         End If
      Case "Maryland"
         If lstCaps.Text = "Annopolis" Then
            LoadStatePic 19, "B", 551, 187
         Else
            LoadStatePic 19, "R", 551, 187
            lstMissed.AddItem "Maryland - " & aryCapData(19)
         End If
      Case "Massachusetts"
         If lstCaps.Text = "Boston" Then
            LoadStatePic 20, "B", 613, 126
         Else
            LoadStatePic 20, "R", 613, 126
            lstMissed.AddItem "Massachusetts - " & aryCapData(20)
         End If
      Case "Michigan"
         If lstCaps.Text = "Lansing" Then
            LoadStatePic 21, "B", 418, 86
         Else
            LoadStatePic 21, "R", 418, 86
            lstMissed.AddItem "Michigan - " & aryCapData(21)
         End If
      Case "Minnesota"
         If lstCaps.Text = "St. Paul" Then
            LoadStatePic 22, "B", 343, 59
         Else
            LoadStatePic 22, "R", 343, 59
            lstMissed.AddItem "Minnesota - " & aryCapData(22)
         End If
      Case "Missouri"
         If lstCaps.Text = "Jefferson City" Then
            LoadStatePic 23, "B", 357, 195
         Else
            LoadStatePic 23, "R", 357, 195
            lstMissed.AddItem "Missouri - " & aryCapData(23)
         End If
      Case "Mississippi"
         If lstCaps.Text = "Jackson" Then
            LoadStatePic 24, "B", 410, 281
         Else
            LoadStatePic 24, "R", 410, 281
            lstMissed.AddItem "Mississippi - " & aryCapData(24)
         End If
      Case "Montana"
         If lstCaps.Text = "Helena" Then
            LoadStatePic 25, "B", 144, 41
         Else
            LoadStatePic 25, "R", 144, 41
            lstMissed.AddItem "Montana - " & aryCapData(25)
         End If
      Case "No. Carolina"
         If lstCaps.Text = "Raleigh" Then
            LoadStatePic 26, "B", 501, 235
         Else
            LoadStatePic 26, "R", 501, 235
            lstMissed.AddItem "No. Carolina - " & aryCapData(32)
         End If
      Case "No. Dakota"
         If lstCaps.Text = "Bismarck" Then
            LoadStatePic 27, "B", 267, 61
         Else
            LoadStatePic 27, "R", 267, 61
            lstMissed.AddItem "No. Dakota - " & aryCapData(27)
         End If
      Case "Nebraska"
         If lstCaps.Text = "Lincoln" Then
            LoadStatePic 28, "B", 260, 153
         Else
            LoadStatePic 28, "R", 260, 153
            lstMissed.AddItem "Nebraska - " & aryCapData(28)
         End If
      Case "Nevada"
         If lstCaps.Text = "Carson City" Then
            LoadStatePic 29, "B", 67, 137
         Else
            LoadStatePic 29, "R", 67, 137
            lstMissed.AddItem "Nevada - " & aryCapData(29)
         End If
      Case "New Hampshire"
         If lstCaps.Text = "Concord" Then
            LoadStatePic 30, "B", 619, 90
         Else
            LoadStatePic 30, "R", 619, 90
            lstMissed.AddItem "New Hampshire - " & aryCapData(30)
         End If
      Case "New Jersy"
         If lstCaps.Text = "Trenton" Then
            LoadStatePic 31, "B", 598, 159
         Else
            LoadStatePic 31, "R", 598, 159
            lstMissed.AddItem "New Jersy - " & aryCapData(31)
         End If
      Case "New Mexico"
         If lstCaps.Text = "Santa Fe" Then
            LoadStatePic 32, "B", 179, 240
         Else
            LoadStatePic 32, "R", 179, 240
            lstMissed.AddItem "New Mexico - " & aryCapData(32)
         End If
      Case "New York"
         If lstCaps.Text = "Albany" Then
            LoadStatePic 33, "B", 540, 101
         Else
            LoadStatePic 33, "R", 540, 101
            lstMissed.AddItem "New York - " & aryCapData(33)
         End If
      Case "Ohio"
         If lstCaps.Text = "Columbus" Then
            LoadStatePic 34, "B", 484, 162
         Else
            LoadStatePic 34, "R", 484, 162
            lstMissed.AddItem "Ohio -" & aryCapData(34)
         End If
      Case "Oklahoma"
         If lstCaps.Text = "Oklahoma City" Then
            LoadStatePic 35, "B", 266, 249
         Else
            LoadStatePic 35, "R", 266, 249
            lstMissed.AddItem "Oklahoma - " & aryCapData(35)
         End If
      Case "Oregon"
         If lstCaps.Text = "Salem" Then
            LoadStatePic 36, "B", 33, 60
         Else
            LoadStatePic 36, "R", 33, 60
            lstMissed.AddItem "Oregon - " & aryCapData(36)
         End If
      Case "Pennsylvania"
         If lstCaps.Text = "Harrisburg" Then
            LoadStatePic 37, "B", 534, 151
         Else
            LoadStatePic 37, "R", 534, 151
            lstMissed.AddItem "Pennsylvania - " & aryCapData(37)
         End If
      Case "Rhode Island"
         If lstCaps.Text = "Providence" Then
            LoadStatePic 38, "B", 632, 140
         Else
            LoadStatePic 38, "B", 632, 140
            lstMissed.AddItem "Rhode Island - " & aryCapData(38)
         End If
      Case "So. Carolina"
         If lstCaps.Text = "Columbia" Then
            LoadStatePic 39, "B", 515, 269
         Else
            LoadStatePic 39, "R", 515, 269
            lstMissed.AddItem "So. Carolina - " & aryCapData(39)
         End If
      Case "So. Dakota"
         If lstCaps.Text = "Pierre" Then
            LoadStatePic 40, "B", 263, 108
         Else
            LoadStatePic 40, "R", 263, 108
            lstMissed.AddItem "So. Dakota - " & aryCapData(40)
         End If
      Case "Tennessee"
         If lstCaps.Text = "Nashville" Then
            LoadStatePic 41, "B", 427, 247
         Else
            LoadStatePic 41, "R", 427, 247
            lstMissed.AddItem "Tennessee - " & aryCapData(41)
         End If
      Case "Texas"
         If lstCaps.Text = "Austin" Then
            LoadStatePic 42, "B", 214, 256
         Else
            LoadStatePic 42, "R", 214, 256
            lstMissed.AddItem "Texas - " & aryCapData(42)
         End If
      Case "Utah"
         If lstCaps.Text = "Salt Lake City" Then
            LoadStatePic 43, "B", 130, 152
         Else
            LoadStatePic 43, "R", 130, 152
            lstMissed.AddItem "Utah - " & aryCapData(43)
         End If
      Case "Vermont"
         If lstCaps.Text = "Montelier" Then
            LoadStatePic 44, "B", 604, 95
         Else
            LoadStatePic 44, "R", 604, 95
            lstMissed.AddItem "Vermont - " & aryCapData(44)
         End If
      Case "Virginia"
         If lstCaps.Text = "Richmond" Then
            LoadStatePic 45, "B", 512, 197
         Else
            LoadStatePic 45, "R", 512, 197
            lstMissed.AddItem "Virginia - " & aryCapData(45)
         End If
      Case "Washington"
         If lstCaps.Text = "Olympia" Then
            LoadStatePic 46, "B", 56, 23
         Else
            LoadStatePic 46, "R", 56, 23
            lstMissed.AddItem "Washington - " & aryCapData(46)
         End If
      Case "Wisconsin"
         If lstCaps.Text = "Madison" Then
            LoadStatePic 47, "B", 389, 96
         Else
            LoadStatePic 47, "R", 389, 96
            lstMissed.AddItem "Wisconsin - " & aryCapData(47)
         End If
      Case "West Virginia"
         If lstCaps.Text = "Charleston" Then
            LoadStatePic 48, "B", 516, 190
         Else
            LoadStatePic 48, "R", 516, 190
            lstMissed.AddItem "West Virginia - " & aryCapData(48)
         End If
      Case "Wyoming"
         If lstCaps.Text = "Cheyenne" Then
            LoadStatePic 49, "B", 178, 112
         Else
            LoadStatePic 49, "R", 178, 112
            lstMissed.AddItem "Wyoming - " & aryCapData(49)
         End If
   End Select
End Sub

Private Sub Make_Quiz()
   Dim M As Integer
   Dim T As Integer
   Dim r As Integer
   Dim aryMix(4) As String
   
   lstCaps.Clear
   
   Randomize
   M = Int(49 * Rnd) 'picks the three other capitols
   T = Int(49 * Rnd)
   r = Int(49 * Rnd)
   '========================================================
   If M = T Or M = r Then M = M + 1
   If T = M Or T = r Then T = T + 1
   If r = M Or r = T Then r = r + 1
   '=====================================================
   P = B(w)
   lblState.Caption = aryStateData(P)
   
   Select Case P
         Case M: M = M + 2
         Case T: T = T + 2
         Case r: r = r + 2
   End Select
   ' if a choice is out of range then correct
   If M > 49 Then M = M - 6
   If T > 49 Then T = T - 20
   If r > 49 Then r = r - 45
   '======================================================
   RandomizeNumbers 0, 3  'mix up capital choices
   aryMix(0) = aryCapData(P) 'capitol that matches state
   aryMix(1) = aryCapData(M)
   aryMix(2) = aryCapData(T)
   aryMix(3) = aryCapData(r)
   
   lstCaps.AddItem aryMix(RanNo(0))
   lstCaps.AddItem aryMix(RanNo(1))
   lstCaps.AddItem aryMix(RanNo(2))
   lstCaps.AddItem aryMix(RanNo(3))
   If w = 49 Then GoTo here:
   w = w + 1 ' increment to get next state
here:
   If x + y = 50 Then
      If lblScoreCor.Caption = "50" Then TE1.Visible = True
      lblState.Caption = "End of Quiz"
      cmdStart.Enabled = False
      lstCaps.Clear
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call v_dmp.Stop(v_dms, v_dmss, 0, 0)
   If mPlay = True Then Call v_dms.Unload(v_dmp)
   Unload Me
End Sub

Private Sub lstCaps_Click()
   If lstCaps.Text = aryCapData(P) Then
      ChkAnswer
      lblState.Caption = ""
      lstCaps.Clear
      x = x + 1
      lblScoreCor.Caption = x
   Else
      ChkAnswer
      lblState.Caption = ""
      lstCaps.Clear
      y = y + 1
      lblScoreInc.Caption = y
   End If
   Make_Quiz
End Sub

Private Sub ClearAll()
   lstMissed.Clear
   Form1.Cls
End Sub

Private Sub Array_Fill()
   Dim MaxNumber As Integer
   Dim seq As Integer
   Dim MainLoop As Integer
   Dim ChosenNumber As Integer
   
   'Set the original array
   MaxNumber = 49
   For seq = 0 To MaxNumber
      a(seq) = seq
   Next seq
   'Main Loop (mix the numbers all up)
   Randomize (Timer)
   For MainLoop = MaxNumber To 0 Step -1
      ChosenNumber = Int(MainLoop * Rnd)
      B(MaxNumber - MainLoop) = a(ChosenNumber)
      a(ChosenNumber) = a(MainLoop)
   Next MainLoop
End Sub

Private Sub RandomizeNumbers(ByVal lngFrom As Long, ByVal lngTo As Long)
   Dim i As Long
   Dim j As Long
   Dim tmp As Long
   
   'Assign the numbers from x to xx
   For i = lngFrom To lngTo
      RanNo(i) = i
   Next i
   
   'Swap the numbers randomly with other numbers in the array
   For i = lngFrom To lngTo
      'Get a random number
      j = CLng((lngTo - lngFrom) * Rnd + lngFrom)
      
      'Swap the random position with the current position
      tmp = RanNo(i)
      RanNo(i) = RanNo(j)
      RanNo(j) = tmp
   Next i
End Sub

Private Sub LoadDataArray()
   
   aryStateData(0) = "Alabama"  ' US states-------------
   aryStateData(1) = "Alaska"
   aryStateData(2) = "Arizonia"
   aryStateData(3) = "Arkansas"
   aryStateData(4) = "California"
   aryStateData(5) = "Colorado"
   aryStateData(6) = "Connecticut"
   aryStateData(7) = "Delaware"
   aryStateData(8) = "Florida"
   aryStateData(9) = "Georgia"
   aryStateData(10) = "Hawaii"
   aryStateData(11) = "Idaho"
   aryStateData(12) = "Illinois"
   aryStateData(13) = "Indiana"
   aryStateData(14) = "Iowa"
   aryStateData(15) = "Kansas"
   aryStateData(16) = "Kentucky"
   aryStateData(17) = "Louisiana"
   aryStateData(18) = "Maine"
   aryStateData(19) = "Maryland"
   aryStateData(20) = "Massachusetts"
   aryStateData(21) = "Michigan"
   aryStateData(22) = "Minnesota"
   aryStateData(23) = "Missouri"
   aryStateData(24) = "Mississippi"
   aryStateData(25) = "Montana"
   aryStateData(26) = "No. Carolina"
   aryStateData(27) = "No. Dakota"
   aryStateData(28) = "Nebraska"
   aryStateData(29) = "Nevada"
   aryStateData(30) = "New Hampshire"
   aryStateData(31) = "New Jersy"
   aryStateData(32) = "New Mexico"
   aryStateData(33) = "New York"
   aryStateData(34) = "Ohio"
   aryStateData(35) = "Oklahoma"
   aryStateData(36) = "Oregon"
   aryStateData(37) = "Pennsylvania"
   aryStateData(38) = "Rhode Island"
   aryStateData(39) = "So. Carolina"
   aryStateData(40) = "So. Dakota"
   aryStateData(41) = "Tennessee"
   aryStateData(42) = "Texas"
   aryStateData(43) = "Utah"
   aryStateData(44) = "Vermont"
   aryStateData(45) = "Virginia"
   aryStateData(46) = "Washington"
   aryStateData(47) = "Wisconsin"
   aryStateData(48) = "West Virginia"
   aryStateData(49) = "Wyoming"
   
   aryCapData(0) = "Montgomery"   'US capitols-----------
   aryCapData(1) = "Juneau"
   aryCapData(2) = "Phoenix"
   aryCapData(3) = "Little Rock"
   aryCapData(4) = "Sacramento"
   aryCapData(5) = "Denver"
   aryCapData(6) = "Hartford"
   aryCapData(7) = "Dover"
   aryCapData(8) = "Tallahassee"
   aryCapData(9) = "Atlanta"
   aryCapData(10) = "Honolulu"
   aryCapData(11) = "Boise"
   aryCapData(12) = "Springfield"
   aryCapData(13) = "Indianapolis"
   aryCapData(14) = "Des Moines"
   aryCapData(15) = "Topeka"
   aryCapData(16) = "Frankfort"
   aryCapData(17) = "Baton Rouge"
   aryCapData(18) = "Augusta"
   aryCapData(19) = "Annopolis"
   aryCapData(20) = "Boston"
   aryCapData(21) = "Lansing"
   aryCapData(22) = "St. Paul"
   aryCapData(23) = "Jefferson City"
   aryCapData(24) = "Jackson"
   aryCapData(25) = "Helena"
   aryCapData(26) = "Raleigh"
   aryCapData(27) = "Bismarck"
   aryCapData(28) = "Lincoln"
   aryCapData(29) = "Carson City"
   aryCapData(30) = "Concord"
   aryCapData(31) = "Trenton"
   aryCapData(32) = "Santa Fe"
   aryCapData(33) = "Albany"
   aryCapData(34) = "Columbus"
   aryCapData(35) = "Oklahoma City"
   aryCapData(36) = "Salem"
   aryCapData(37) = "Harrisburg"
   aryCapData(38) = "Providence"
   aryCapData(39) = "Columbia"
   aryCapData(40) = "Pierre"
   aryCapData(41) = "Nashville"
   aryCapData(42) = "Austin"
   aryCapData(43) = "Salt Lake City"
   aryCapData(44) = "Montelier"
   aryCapData(45) = "Richmond"
   aryCapData(46) = "Olympia"
   aryCapData(47) = "Madison"
   aryCapData(48) = "Charleston"
   aryCapData(49) = "Cheyenne"
End Sub

Private Sub LoadStatePic(stateInt As Integer, Colr As String, L As Integer, T As Integer)
   DoEvents
   picStor.Picture = LoadPicture(App.Path & "\Images\" & stateInt & Colr & ".gif")
   TransparentBlt Form1.hdc, L, T, picStor.ScaleWidth, picStor.ScaleHeight, picStor.hdc, 0, 0, picStor.ScaleWidth, picStor.ScaleHeight, vbWhite
   Form1.Refresh
End Sub

Private Sub PlayMid()
    On Local Error GoTo ErrSub
    If vs_filename = "" Then Exit Sub
    Set v_dms = v_dml.LoadSegment(vs_filename)
    v_dms.SetRepeats 18             'sets the number of times to loop
    If StrConv(Right(vs_filename, 4), vbLowerCase) = ".mid" Then
        v_dms.SetStandardMidiFile
    End If
    
    Call v_dmp.SetMasterAutoDownload(True)
    Call v_dms.Download(v_dmp)
    
    Set v_dmss = v_dmp.PlaySegment(v_dms, 0, 0)
    v_dmss.GetRepeats      'go get loop number
    Call v_dmp.SetMasterVolume(hsbVolume.Value)
    mPlay = True
    Exit Sub
ErrSub:
    Call ErrMess(Err.Number, Err.Description)
End Sub

Private Sub StopMid()
    On Local Error GoTo ErrSub
    If v_dms Is Nothing Then Exit Sub
    Call v_dmp.Stop(v_dms, v_dmss, 0, 0)
    Call v_dms.Unload(v_dmp)
    Exit Sub
ErrSub:
    Call ErrMess(Err.Number, Err.Description)
End Sub

Private Sub chkPlay_Click()
   If chkPlay.Value = Checked Then
      PlayMid
   Else
      StopMid
   End If
End Sub

Private Sub hsbVolume_Change()
    On Local Error GoTo ErrSub
    Call v_dmp.SetMasterVolume(hsbVolume.Value)
    Exit Sub
ErrSub:
    Call ErrMess(Err.Number, Err.Description)
End Sub

Private Sub hsbVolume_Scroll()
    On Local Error GoTo ErrSub
    Call v_dmp.SetMasterVolume(hsbVolume.Value)
    Exit Sub
ErrSub:
    Call ErrMess(Err.Number, Err.Description)
End Sub

Sub ErrMess(eNumber, eDesc)
    Dim Msg As String
    Msg = "An error has been occured."
    Msg = Msg & Chr(13) & "(" & eNumber & ") - " & eDesc
    MsgBox Msg, vbCritical
    End
End Sub

