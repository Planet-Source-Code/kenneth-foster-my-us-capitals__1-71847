VERSION 5.00
Begin VB.Form frmZoom 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2250
   ControlBox      =   0   'False
   DrawWidth       =   5
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   150
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   150
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   270
      Left            =   1815
      TabIndex        =   1
      Top             =   1695
      Width           =   315
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   50
      Left            =   3150
      Top             =   2655
   End
   Begin VB.Label lblCoord 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   -30
      TabIndex        =   0
      Top             =   1995
      Width           =   3015
   End
End
Attribute VB_Name = "frmZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************************
'  Copyright © 2004, Alan Tucker, All Rights Reserved
'  Contact alan_usa@hotmail.com for usage restrictions
'**************************************************************************************************
Option Explicit

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const MOUSE_BUFFER = 300
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2

Private Type PointAPI
     X As Long
     Y As Long
End Type ' POINTAPI

Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, _
     ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
     ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, _
     ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
     ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
     ByVal dwRop As Long) As Long
     Private OldX As Integer
     Private OldY As Integer

Private Sub Form_DblClick()
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub lblCoord_DblClick()
   Unload Me
End Sub

Private Sub lblCoord_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      OldX = X
      OldY = Y
End Sub

Private Sub lblCoord_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then
      Me.Left = Me.Left + (X - OldX)
      Me.Top = Me.Top + (Y - OldY)
   End If

End Sub

Private Sub tmrUpdate_Timer()
     Dim m_Cursor As PointAPI
     Dim m_hDC As Long
     Dim lRtn As Long
     Dim X As Long
     Dim Y As Long
     Dim lScrHt As Long
     Dim lScrWt As Long
     Cls
     lScrHt = Screen.Height \ Screen.TwipsPerPixelY
     lScrWt = Screen.Width \ Screen.TwipsPerPixelX
     ' Get the position of the mouse cursor
     GetCursorPos m_Cursor
     ' update coordinates label
     lblCoord = "  X = " & m_Cursor.X & "  Y = " & m_Cursor.Y
     ' convert x and y positions into twips and add buffer.  Buffer necessary
     ' to create some space from the mouse cursor so that we don't see the corner
     ' of the zoom box in the magnification.
     ' If we are at the right of the screen
     If m_Cursor.X + (Me.Width \ Screen.TwipsPerPixelX) + (MOUSE_BUFFER \ Screen.TwipsPerPixelX) > lScrWt And _
          m_Cursor.Y + (Me.Height \ Screen.TwipsPerPixelY) + (MOUSE_BUFFER \ Screen.TwipsPerPixelY) > lScrHt Then
          X = (m_Cursor.X * Screen.TwipsPerPixelX) - (Me.Width + MOUSE_BUFFER)
          Y = (m_Cursor.Y * Screen.TwipsPerPixelY) - (Me.Height + MOUSE_BUFFER)
     ElseIf m_Cursor.X + (Me.Width \ Screen.TwipsPerPixelX) + _
          (MOUSE_BUFFER \ Screen.TwipsPerPixelX) > lScrWt Then
          X = (m_Cursor.X * Screen.TwipsPerPixelX) - (Me.Width + MOUSE_BUFFER)
          Y = m_Cursor.Y * Screen.TwipsPerPixelY + MOUSE_BUFFER
     ElseIf m_Cursor.Y + (Me.Height \ Screen.TwipsPerPixelY) + _
          (MOUSE_BUFFER \ Screen.TwipsPerPixelY) > lScrHt Then
          X = m_Cursor.X * Screen.TwipsPerPixelX + MOUSE_BUFFER
          Y = (m_Cursor.Y * Screen.TwipsPerPixelY) - (Me.Height + MOUSE_BUFFER)
     Else
          X = m_Cursor.X * Screen.TwipsPerPixelX + MOUSE_BUFFER
          Y = m_Cursor.Y * Screen.TwipsPerPixelY + MOUSE_BUFFER
     End If
     ' move the form with the cursor
    ' Me.Move X, Y, Me.Width, Me.Height
     ' Get the screen device context
     m_hDC = GetWindowDC(0)
     ' Blit the coordinates, passed in the api call, and stretch it into
     ' our form
     StretchBlt Me.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, _
        m_hDC, m_Cursor.X - 24, m_Cursor.Y - 24, 48, 48, vbSrcCopy
     ' Draw a box to make the form distinguishable from the background.  Set the forms
     ' forecolor to make changes to it.
     frmZoom.Line (0, 0)-(frmZoom.ScaleWidth - 1, frmZoom.ScaleHeight - 1), , B
     ' Bring the window to the top.
     Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
     ' release the screen's device context
     lRtn = ReleaseDC(0, m_hDC)
     ' If at coordinate 0, 0 then quit
     If m_Cursor.X = 0 And m_Cursor.Y = 0 Then
          Unload Me
          Set frmZoom = Nothing
     End If
     frmZoom.DrawWidth = 2
     frmZoom.ForeColor = vbBlack
     frmZoom.PSet (frmZoom.ScaleWidth / 2, frmZoom.ScaleHeight / 2)
     frmZoom.PSet (frmZoom.ScaleWidth / 2 + 4, frmZoom.ScaleHeight / 2)
     frmZoom.PSet (frmZoom.ScaleWidth / 2 - 4, frmZoom.ScaleHeight / 2)
     frmZoom.PSet (frmZoom.ScaleWidth / 2 + 8, frmZoom.ScaleHeight / 2)
     frmZoom.PSet (frmZoom.ScaleWidth / 2 - 8, frmZoom.ScaleHeight / 2)
     frmZoom.PSet (frmZoom.ScaleWidth / 2, frmZoom.ScaleHeight / 2 + 4)
     frmZoom.PSet (frmZoom.ScaleWidth / 2, frmZoom.ScaleHeight / 2 - 4)
     frmZoom.PSet (frmZoom.ScaleWidth / 2, frmZoom.ScaleHeight / 2 + 8)
     frmZoom.PSet (frmZoom.ScaleWidth / 2, frmZoom.ScaleHeight / 2 - 8)
End Sub

