VERSION 5.00
Begin VB.UserControl Duncan_XPButton 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF00FF&
   ClientHeight    =   825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1725
   DefaultCancel   =   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   55
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   115
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   -15
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Duncan_XPButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'What?
'A button that when Themes are enabled draws in XP style
'but when themes are removed reverts to an old 98 style button.

'Why?
'So that I can have a propper XP style button that I can use and
'see while developing.
'Because it doesnt require a manifest.
'Because I am making a full collection of base controls all
'like this and this is the latest.

'How?
'Uses a normal button for unthemed behaviour.
'Uses XP theme drawing to paint a button if possible.

'Behaviour
'It should behave just like a normal button.
'Please report any bugs.
'Picture works best if it is assigned an Icon rather than a gif or bmp

'Who?
'Thanks to Paul (programming god) Catton for his amazing subclassing work that has made this project possible
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=54117&lngWId=1

'When?
'Last Updated : June 2005

'to do:


'======================================================================================================================================================
'MY DECLARES FOR THIS CONTROL
'======================================================================================================================================================
'drawing the picture
Private Const DST_ICON = &H3
Private Const DST_BITMAP = &H4
Private Const DSS_DISABLED = &H20
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" _
       (ByVal hdc As Long, _
        ByVal hBrush As Long, _
        ByVal lpDrawStateProc As Long, _
        ByVal lParam As Long, _
        ByVal wParam As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal cX As Long, _
        ByVal cY As Long, _
        ByVal fuFlags As Long) As Long
'creating button shape
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'drawing the themed button
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Const FOCUSPADDING As Long = 4
Private Declare Function OpenThemeData Lib "uxtheme.dll" _
   (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" _
   (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeParentBackground Lib "uxtheme.dll" _
   (ByVal hwnd As Long, ByVal hdc As Long, prc As RECT) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal lHdc As Long, _
    ByVal iPartId As Long, ByVal iStateId As Long, _
    pRect As RECT, pClipRect As RECT) As Long
Private Declare Function GetThemeBackgroundContentRect Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal hdc As Long, _
    ByVal iPartId As Long, ByVal iStateId As Long, _
    pBoundingRect As RECT, pContentRect As RECT) As Long
Private Declare Function DrawThemeText Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, _
    ByVal iStateId As Long, ByVal pszText As Long, _
    ByVal iCharCount As Long, ByVal dwTextFlag As Long, _
    ByVal dwTextFlags2 As Long, pRect As RECT) As Long
Public Enum eButtonStyle
    Button = 0
    Toolbar = 1
End Enum
Private Enum eButtonState
    Normal = 1
    Hot = 2
    Pressed = 3
    Disabled = 4
    Defaulted = 5
End Enum
Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Private Enum DrawTextFlags
    DT_TOP = &H0
    DT_LEFT = &H0
    DT_CENTER = &H1
    DT_RIGHT = &H2
    DT_VCENTER = &H4
    DT_BOTTOM = &H8
    DT_WORDBREAK = &H10
    DT_SINGLELINE = &H20
    DT_EXPANDTABS = &H40
    DT_TABSTOP = &H80
    DT_NOCLIP = &H100
    DT_EXTERNALLEADING = &H200
    DT_CALCRECT = &H400
    DT_NOPREFIX = &H800
    DT_INTERNAL = &H1000
    DT_EDITCONTROL = &H2000
    DT_PATH_ELLIPSIS = &H4000
    DT_END_ELLIPSIS = &H8000
    DT_MODIFYSTRING = &H10000
    DT_RTLREADING = &H20000
    DT_WORD_ELLIPSIS = &H40000
    DT_NOFULLWIDTHCHARBREAK = &H80000
    DT_HIDEPREFIX = &H100000
    DT_PREFIXONLY = &H200000
End Enum

Public Enum eAlignment
    topleft = DT_TOP Or DT_LEFT Or DT_SINGLELINE  'left top
    topcenter = DT_TOP Or DT_CENTER Or DT_SINGLELINE 'top center
    topright = DT_TOP Or DT_RIGHT Or DT_SINGLELINE 'top right
    middleleft = DT_VCENTER Or DT_LEFT Or DT_SINGLELINE 'middle left
    middlecenter = DT_VCENTER Or DT_CENTER Or DT_SINGLELINE 'middle center
    middleright = DT_VCENTER Or DT_RIGHT Or DT_SINGLELINE 'middle right
    bottomleft = DT_BOTTOM Or DT_LEFT Or DT_SINGLELINE 'bottom left
    bottomcenter = DT_BOTTOM Or DT_CENTER Or DT_SINGLELINE 'bottom center
    bottomright = DT_BOTTOM Or DT_RIGHT Or DT_SINGLELINE 'bottom right
End Enum

'mouse position
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'INTERNAL USE
Private m_Hot As Boolean            'is the mouse over the control?
Private m_MouseDown As Boolean      'is the mouse down?
Private m_StateIdButton As eButtonState 'the current draw state of the button
Private m_Active As Boolean         'are we the active app? ie we dont process mouse over events unless we have focus
Private m_Caption As String            'text for the button
Private m_CaptionAlignment As eAlignment   'alignment for text
Private m_Enabled As Boolean        'Enabled?
Private m_UseThemes As Boolean  'Which button to use
Private m_HasFocus As Boolean       'Does button have focus?
Private m_DisplayAsDefault As Boolean
Private m_FocusRect As Boolean      'Do we draw the focus rectangle?
Private m_ButtonStyle As eButtonStyle   'Which theme to use when drawing
Private m_Picture As StdPicture
Private m_PictureAlignment As eAlignment    'alignment for picture
Private m_PicturePadding As Long    'How many pixels to indent
Private m_CaptionPadding As Long    'How many pixels to indent
'we dont want to send 2 click events if the user has mouse down
'over the control and presses enter so we set which input type
'has priority and only the latest event will be processed
Private m_EBP As eEventBeingProcessed
Private Enum eEventBeingProcessed
    None = 0
    Mouse = 1
    Keyboard = 2
End Enum
'Public events
Public Event Click()
'======================================================================================================================================================
'SUBCLASSING DECLARES
'======================================================================================================================================================
Private Enum eMsgWhen
  MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
  MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
  MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum
Private Type tSubData                                                                   'Subclass data type
  hwnd                               As Long                                            'Handle of the window being subclassed
  nAddrSub                           As Long                                            'The address of our new WndProc (allocated memory).
  nAddrOrig                          As Long                                            'The address of the pre-existing WndProc
  nMsgCntA                           As Long                                            'Msg after table entry count
  nMsgCntB                           As Long                                            'Msg before table entry count
  aMsgTblA()                         As Long                                            'Msg after table array
  aMsgTblB()                         As Long                                            'Msg Before table array
End Type
Private sc_aSubData()                As tSubData                                        'Subclass data array
Private Const ALL_MESSAGES           As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED             As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC            As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                      'Table A (after) entry count patch offset
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Window Messages
Private Const WM_ACTIVATE = &H6
Private Const WM_NCACTIVATE As Long = &H86
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_SYSCOLORCHANGE As Long = &H15
Private Const WM_THEMECHANGED As Long = &H31A

'//Mouse tracking declares
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum
Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                              As Long
    dwFlags                             As TRACKMOUSEEVENT_FLAGS
    hwndTrack                           As Long
    dwHoverTime                         As Long
End Type

'Subclass handler
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
Attribute zSubclass_Proc.VB_MemberFlags = "40"
'THIS MUST BE THE FIRST PUBLIC ROUTINE IN THIS FILE.
'That includes public properties also
'Parameters:
  'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
  'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
  'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
  'hWnd     - The window handle
  'uMsg     - The message number
  'wParam   - Message related data
  'lParam   - Message related data

    'Debug.Print WMbyName(uMsg) & " " & wParam
    Select Case uMsg
        Case WM_THEMECHANGED, WM_SYSCOLORCHANGE
            InitialiseThemes
            Refresh
        Case WM_MOUSELEAVE
            'called when mouse leaves the control
            If ProcessingMessages Then
                m_Hot = False
                m_MouseDown = False
                If Not m_EBP = Keyboard Then    'if there isnt a keyboard event stored, eg they are holding down the spacebar, then refresh
                    RefreshButton
                End If
            End If
        Case WM_ACTIVATE, WM_NCACTIVATE
            If wParam Then  '----------------------------------- Activated
                'Debug.Print "activated " & wParam & " " & lParam & " " & Now
                m_Active = True
                m_Hot = CheckForHot
            Else            '----------------------------------- Deactivated
                'Debug.Print "deactivated " & wParam & " " & lParam & " " & Now
                m_Active = False
                m_Hot = False
            End If
            Refresh
    End Select
End Sub

'======================================================================================================================================================
'Functions
'======================================================================================================================================================

Public Property Get Caption() As String
    Caption = m_Caption
End Property
Public Property Let Caption(sVal As String)
    If sVal <> m_Caption Then
        m_Caption = sVal
        UserControl.AccessKeys = GetAccessKey '---------- Set AccessKey property if desired
        Command1.Caption = m_Caption
        PropertyChanged "Caption"
        Refresh
    End If
End Property
Public Property Get CaptionAlignment() As eAlignment
    CaptionAlignment = m_CaptionAlignment
End Property
Public Property Let CaptionAlignment(eVal As eAlignment)
    If eVal <> m_CaptionAlignment Then
        m_CaptionAlignment = eVal
        PropertyChanged "CaptionAlignment"
        Refresh
    End If
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property
Public Property Let Enabled(bVal As Boolean)
    If bVal <> m_Enabled Then
        m_Enabled = bVal
        PropertyChanged "Enabled"
        Command1.Enabled = m_Enabled
        UserControl.Enabled = m_Enabled
        RefreshButton
    End If
End Property
Public Property Get FocusRect() As Boolean
    FocusRect = m_FocusRect
End Property
Public Property Let FocusRect(bVal As Boolean)
    If bVal <> m_FocusRect Then
        m_FocusRect = bVal
        PropertyChanged "FocusRect"
        RefreshButton
    End If
End Property
Public Property Get ButtonStyle() As eButtonStyle
    ButtonStyle = m_ButtonStyle
End Property
Public Property Let ButtonStyle(eVal As eButtonStyle)
    If eVal <> m_ButtonStyle Then
        m_ButtonStyle = eVal
        PropertyChanged "ButtonStyle"
        RefreshButton
    End If
End Property
Public Property Get hwnd() As Long
    If m_UseThemes Then
        hwnd = UserControl.hwnd
    Else
        hwnd = Command1.hwnd
    End If
End Property

Public Property Get Picture() As StdPicture
    Set Picture = m_Picture
End Property
Public Property Set Picture(ByVal oVal As StdPicture)
    Set m_Picture = oVal
    Set Command1.Picture = oVal
    Refresh
    PropertyChanged "Picture"
End Property
Public Property Get PictureAlignment() As eAlignment
    PictureAlignment = m_PictureAlignment
End Property
Public Property Let PictureAlignment(eVal As eAlignment)
    If eVal <> m_PictureAlignment Then
        m_PictureAlignment = eVal
        PropertyChanged "PictureAlignment"
        Refresh
    End If
End Property

Public Property Get PicturePadding() As Long
    PicturePadding = m_PicturePadding
End Property
Public Property Let PicturePadding(lVal As Long)
    If lVal <> m_PicturePadding Then
        m_PicturePadding = lVal
        PropertyChanged "PicturePadding"
        Refresh
    End If
End Property
Public Property Get CaptionPadding() As Long
    CaptionPadding = m_CaptionPadding
End Property
Public Property Let CaptionPadding(lVal As Long)
    If lVal <> m_CaptionPadding Then
        m_CaptionPadding = lVal
        PropertyChanged "CaptionPadding"
        Refresh
    End If
End Property
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property
Public Property Set Font(ByVal fVal As Font)
    Set UserControl.Font = fVal
    Set Command1.Font = fVal
    PropertyChanged "Font"
    Refresh
End Property

'----------------
'PUBLIC FUNCTIONS
'----------------

Public Sub Refresh()
    m_StateIdButton = -1    'set to invalid value for force update
    RefreshButton
End Sub

Private Sub RefreshButton()
    If m_UseThemes Then
        Command1.Visible = False
        DrawThemeButton
    Else
        UserControl.Cls
        Command1.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
        Command1.Visible = True
        Command1.Refresh
    End If
End Sub


Private Sub DrawThemeButton()
    Dim hTheme As Long
    Dim lPartId As Long
    Dim lStateId As eButtonState
    Dim tR As RECT
    Dim tTextR As RECT
    Dim tIconR As RECT
    Dim tImlR As RECT
    Dim tFocusR As RECT
    Dim retval As Long
    Dim L As Long
    Dim T As Long
    
    tR.Left = 0
    tR.Top = 0
    tR.Right = UserControl.ScaleWidth
    tR.Bottom = UserControl.ScaleHeight
   
    lPartId = 1
    
    If m_Enabled Then
        If m_Hot Then
            'the mouse is over us
            'is it pressed or not
            If m_MouseDown Then
                lStateId = Pressed
            Else
                lStateId = Hot
            End If
        Else
            'draw normal
            If m_DisplayAsDefault And m_ButtonStyle = Button Then
                lStateId = Defaulted
            Else
                lStateId = Normal
            End If
        End If
    Else
        lStateId = Disabled
    End If
    
    
    
    'if state has changed
    'or we are in design mode
    'If ((lStateId <> m_StateIdButton) Or (Not UserControl.Ambient.UserMode)) Then
        'state has changed - redraw the control
        m_StateIdButton = lStateId
        If m_ButtonStyle = Button Then
            hTheme = OpenThemeData(UserControl.hwnd, StrPtr("BUTTON"))
        Else
            hTheme = OpenThemeData(UserControl.hwnd, StrPtr("TOOLBAR"))
        End If
        
        If hTheme <> 0 Then
            UserControl.Cls
            'retval = DrawThemeParentBackground( _
                UserControl.hWnd, _
                UserControl.hdc, _
                tR)
    
            retval = DrawThemeBackground(hTheme, _
                UserControl.hdc, _
                lPartId, _
                lStateId, _
                tR, tR)
         
         
            If Len(m_Caption) > 0 Then
                retval = GetThemeBackgroundContentRect( _
                    hTheme, _
                    UserControl.hdc, _
                    lPartId, _
                    lStateId, _
                    tR, _
                    tTextR)
                    
                 Select Case m_CaptionAlignment
                 Case bottomleft
                    tTextR.Bottom = tTextR.Bottom - m_CaptionPadding - 2
                    tTextR.Left = tTextR.Left + m_CaptionPadding
                 Case bottomcenter
                    tTextR.Bottom = tTextR.Bottom - m_CaptionPadding - 2
                 Case bottomright
                    tTextR.Bottom = tTextR.Bottom - m_CaptionPadding - 2
                    tTextR.Right = tTextR.Right - m_CaptionPadding
                 Case middleleft
                    tTextR.Left = tTextR.Left + m_CaptionPadding
                 Case middleright
                    tTextR.Right = tTextR.Right - m_CaptionPadding
                 Case topcenter
                    tTextR.Top = tTextR.Top + m_CaptionPadding
                 Case topleft
                    tTextR.Left = tTextR.Left + m_CaptionPadding
                    tTextR.Top = tTextR.Top + m_CaptionPadding
                 Case topright
                    tTextR.Top = tTextR.Top + m_CaptionPadding
                    tTextR.Right = tTextR.Right - m_CaptionPadding
                 End Select
                 
                retval = DrawThemeText( _
                   hTheme, _
                   UserControl.hdc, _
                   lPartId, _
                   lStateId, _
                   StrPtr(m_Caption), _
                   -1, _
                   m_CaptionAlignment, _
                   0, _
                   tTextR)
            End If
            
            If FocusRect And m_Enabled Then
                If m_HasFocus Then
                    Dim lSpacer As Long
                    lSpacer = 4
                    tFocusR.Top = tR.Top + FOCUSPADDING
                    tFocusR.Left = tR.Left + FOCUSPADDING
                    tFocusR.Right = tR.Right - FOCUSPADDING
                    tFocusR.Bottom = tR.Bottom - FOCUSPADDING
                    If tFocusR.Bottom > tFocusR.Top And tFocusR.Right > tFocusR.Left Then
                        DrawFocusRect UserControl.hdc, tFocusR
                    End If
                End If
            End If
            
            DrawPicture

            CloseThemeData hTheme
        End If
    'End If
    
End Sub

Private Sub Command1_Click()
    If m_Enabled And m_Active Then
        RaiseEvent Click
    End If
End Sub

'-----------------
'PRIVATE FUNCTIONS
'-----------------
Private Function CheckForButton() As Boolean
    'lets you know if the left mouse button is down
    Dim retval As Long
    retval = GetKeyState(vbKeyLButton)  'returns a negative value while the button is being depressed
    If retval < False Then
        CheckForButton = True
    End If
End Function

Private Function CheckForHot() As Boolean
    'lets you know if the pointer is over the control
    Dim P As POINTAPI
    Dim H As Long
    'get position of cursor
    GetCursorPos P
    'Get the window under that position
    H = WindowFromPoint(P.X, P.Y)
    If H = UserControl.hwnd Then
        CheckForHot = True
    End If
    
End Function

Private Sub SetWindowRegion()
    'Trims the usercontrol into a rounded shape
    Dim RGN As Long
    RGN = CreateRoundRectRgn(0, 0, UserControl.ScaleWidth + 1, UserControl.ScaleHeight + 1, 2, 2)
    'Apply the region
    SetWindowRgn UserControl.hwnd, RGN, True
    'clean up
    DeleteObject RGN
End Sub

Private Function ActivityIsInControl(X As Single, Y As Single) As Boolean
    'called from the mouse move and mouse up functions
    'and lets you know if the mouse is over the control
    'it wont be for example if the user pressed down and then
    'dragged outside of the control
    With UserControl
        If X >= 0 And X <= .ScaleWidth And Y >= 0 And Y <= .ScaleHeight Then
            ActivityIsInControl = True
        End If
    End With
End Function

Private Function GetAccessKey() As String
    'Extracts and returns the AccessKey
    'Function sourced from LiTe Templer of PSC
    Dim lPos    As Long
    Dim lLen    As Long
    Dim lSearch As Long
    Dim sChr    As String
    lLen = Len(m_Caption)
    If lLen = 0 Then Exit Function
    lPos = 1
    Do While lPos + 1 < lLen
        lSearch = InStr(lPos, m_Caption, "&")
        If lSearch = 0 Or lSearch = lLen Then Exit Do
        sChr = LCase$(Mid$(m_Caption, lSearch + 1, 1))
        If sChr = "&" Then
            lPos = lSearch + 2
        Else
            GetAccessKey = sChr
            Exit Do
        End If
    Loop
End Function

Private Function InitialiseThemes() As Boolean
    'tests to see if themes are available
    'if it can then we can draw the XP buttons
    'if it cant then we use the old 98 style buttons
    On Error GoTo Whoops
    Dim hTheme As Long

    'opening and closing theme
    hTheme = OpenThemeData(UserControl.hwnd, StrPtr("BUTTON"))
    If hTheme <> 0 Then
        m_UseThemes = True
    Else
        m_UseThemes = False
    End If
    CloseThemeData hTheme
    
Whoops:
    'Not theme enabled
End Function

Private Sub TrackMouseLeave()
    'Starts tracking the mouse
    'When the mouse leaves the control the WM_MOUSELEAVE message will be sent
    'Doesnt work for transparent windows :(
    On Error GoTo Errs
    Dim tme As TRACKMOUSEEVENT_STRUCT
    With tme
        .cbSize = Len(tme)
        .dwFlags = TME_LEAVE
        .hwndTrack = UserControl.hwnd
    End With
    Call TrackMouseEvent(tme) '---- Track the mouse leaving the indicated window via subclassing
Errs:
End Sub

Private Function ProcessingMessages() As Boolean
    'are we procesing messages from the UserControl?
    If m_Enabled And m_Active And m_UseThemes Then
        ProcessingMessages = True
    End If
End Function

Private Sub DrawPicture()
    Dim picWidth As Long
    Dim picHeight As Long
    Dim Top As Single
    Dim Left As Single
    Dim middle As Single    'vertical
    Dim Bottom As Single
    Dim Right As Single
    Dim center As Single    'horizontal
    Dim drawTop As Single
    Dim drawLeft As Single
    Dim drawFlags As Long
    
    Const Padding As Long = 4
    
    If Not m_Picture Is Nothing Then
        'get picture dimensions
        picWidth = UserControl.ScaleX(m_Picture.Width, vbHimetric, vbPixels)
        picHeight = UserControl.ScaleY(m_Picture.Height, vbHimetric, vbPixels)
        'calc positioning
        Top = Padding
        Left = Padding
        middle = (UserControl.ScaleHeight - picHeight) / 2
        Bottom = UserControl.ScaleHeight - (picHeight + Padding)
        Right = UserControl.ScaleWidth - (picWidth + Padding)
        center = (UserControl.ScaleWidth - picWidth) / 2
        
        'assign positioning
        Select Case m_PictureAlignment
        Case topleft
            drawTop = Top + m_PicturePadding
            drawLeft = Left + m_PicturePadding
        Case topcenter
            drawTop = Top + m_PicturePadding
            drawLeft = center
        Case topright
            drawTop = Top + m_PicturePadding
            drawLeft = Right - m_PicturePadding
        Case middleright
            drawTop = middle
            drawLeft = Right + m_PicturePadding
        Case middleleft
            drawTop = middle
            drawLeft = Left + m_PicturePadding
        Case bottomleft
            drawTop = Bottom - m_PicturePadding
            drawLeft = Left + m_PicturePadding
        Case bottomcenter
            drawTop = Bottom - m_PicturePadding
            drawLeft = center
        Case bottomright
            drawTop = Bottom - m_PicturePadding
            drawLeft = Right - m_PicturePadding
        Case Else
            'middlecenter and unknown
            drawTop = middle
            drawLeft = center
        End Select
        
        If m_Picture.Type = vbPicTypeIcon Then
            drawFlags = DST_ICON
        Else
            'presume its a bitmap (vbPicTypeBitmap)
            drawFlags = DST_BITMAP
        End If
        If Not m_Enabled Then drawFlags = drawFlags Or DSS_DISABLED
        DrawState UserControl.hdc, 0, 0, m_Picture.Handle, 0, drawLeft, drawTop, picWidth, picHeight, drawFlags
    End If
End Sub


'------------
'USER CONTROL
'------------
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    If ProcessingMessages Then
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    m_DisplayAsDefault = Ambient.DisplayAsDefault
    Command1.Default = Ambient.DisplayAsDefault
    Command1.Cancel = UserControl.Extender.Cancel
    RefreshButton
End Sub

Private Sub UserControl_EnterFocus()
    m_HasFocus = True
    Refresh
End Sub

Private Sub UserControl_ExitFocus()
    m_HasFocus = False
    m_MouseDown = False
    m_EBP = None
    Refresh
End Sub

Private Sub UserControl_InitProperties()
    UserControl.AutoRedraw = True
    InitialiseThemes
    Set UserControl.Font = UserControl.Ambient.Font
    UserControl.BackColor = UserControl.Parent.BackColor
    m_Caption = Extender.Name
    m_CaptionAlignment = middlecenter
    m_PictureAlignment = middleleft
    m_FocusRect = False
    m_Enabled = True
    Refresh
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If ProcessingMessages Then
        Select Case KeyCode
        Case 32, 13 'space,enter
            m_EBP = Keyboard
            m_MouseDown = True
            m_Hot = True
            Refresh
        Case 37, 38             'Left Arrow and Up keys
            SendKeys "+{TAB}"
        Case 39, 40             'Right Arrow and Down keys
            SendKeys "{TAB}"
        Case Else
            'Debug.Print KeyCode & " down"
        End Select
    End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If ProcessingMessages Then
        Select Case KeyCode
        Case 32, 13 'space,enter
            If m_EBP = Keyboard Then
                RaiseEvent Click
                m_EBP = None
                m_MouseDown = CheckForButton
                m_Hot = CheckForHot
                Refresh
            End If
        Case Else
            'Debug.Print KeyCode & " up"
        End Select
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ProcessingMessages Then
        If Button = 1 Then
            m_MouseDown = True
            m_EBP = Mouse
            RefreshButton
       End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ProcessingMessages Then
        If Button = 1 Then
            If m_EBP = Mouse Then
                If ActivityIsInControl(X, Y) Then
                    RaiseEvent Click
                    m_EBP = None
                End If
            End If
        End If
        m_MouseDown = 0
        RefreshButton
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ProcessingMessages Then
        If ActivityIsInControl(X, Y) Then
            Call TrackMouseLeave
            m_Hot = True
        Else
            m_Hot = False
        End If
        If Not m_EBP = Keyboard Then
            RefreshButton
        End If
    End If
End Sub

Private Sub UserControl_Resize()
    If m_UseThemes Then
        SetWindowRegion
    End If
    RefreshButton
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    InitialiseThemes
    If Ambient.UserMode Then
        Call Subclass_Start(UserControl.hwnd)
        Call Subclass_AddMsg(UserControl.hwnd, WM_MOUSELEAVE, MSG_AFTER)

        Call Subclass_Start(UserControl.Parent.hwnd)
        'Call Subclass_AddMsg(UserControl.Parent.hwnd, ALL_MESSAGES, MSG_BEFORE)
        If UserControl.Parent.MDIChild Then
            Call Subclass_AddMsg(UserControl.Parent.hwnd, WM_NCACTIVATE, MSG_AFTER)
        Else
            Call Subclass_AddMsg(UserControl.Parent.hwnd, WM_ACTIVATE, MSG_AFTER)
        End If
        Call Subclass_AddMsg(UserControl.Parent.hwnd, WM_SYSCOLORCHANGE, MSG_AFTER)
        Call Subclass_AddMsg(UserControl.Parent.hwnd, WM_THEMECHANGED, MSG_AFTER)
    End If
    
    With PropBag
        Set Font = .ReadProperty("Font", Ambient.Font)
        FocusRect = .ReadProperty("FocusRect", True)
        ButtonStyle = .ReadProperty("ButtonStyle", 0)
        Caption = .ReadProperty("Caption", "")
        CaptionAlignment = .ReadProperty("CaptionAlignment", middlecenter)
        m_Enabled = .ReadProperty("Enabled", True)
        Command1.Enabled = .ReadProperty("Enabled", True)
        Set Picture = .ReadProperty("Picture", Nothing)
        PictureAlignment = .ReadProperty("PictureAlignment", middleleft)
        PicturePadding = .ReadProperty("PicturePadding", 0)
        CaptionPadding = .ReadProperty("CaptionPadding", 0)
    End With
    
    UserControl.BackColor = UserControl.Parent.BackColor
    Command1.Default = UserControl.Ambient.DisplayAsDefault
    Command1.Cancel = UserControl.Extender.Cancel
    RefreshButton
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Font", UserControl.Font, Ambient.Font
        .WriteProperty "FocusRect", FocusRect, True
        .WriteProperty "ButtonStyle", ButtonStyle, 0
        .WriteProperty "Caption", Caption, ""
        .WriteProperty "CaptionAlignment", CaptionAlignment, middlecenter
        .WriteProperty "Enabled", Enabled, True
        .WriteProperty "Picture", Picture, Nothing
        .WriteProperty "PictureAlignment", PictureAlignment, middleleft
        .WriteProperty "PicturePadding", PicturePadding, 0
        .WriteProperty "CaptionPadding", CaptionPadding, 0
    End With
End Sub

Private Sub UserControl_Terminate()
    On Error GoTo Errs
    If Ambient.UserMode Then Call Subclass_StopAll
    Debug.Print "UC terminated"
Errs:
End Sub




'========================================================================================
'Subclass routines below here - The programmer may call any of the following Subclass_??? routines
'======================================================================================================================================================
'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
On Error GoTo Errs
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
  'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
Errs:
End Sub

'Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
On Error GoTo Errs

'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
  'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to be removed from the before, after or both callback tables
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
Errs:
End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
  Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
On Error GoTo Errs
'Parameters:
  'lng_hWnd  - The handle of the window to be subclassed
'Returns;
  'The sc_aSubData() index
  Const CODE_LEN              As Long = 200                                             'Length of the machine code in bytes
  Const FUNC_CWP              As String = "CallWindowProcA"                             'We use CallWindowProc to call the original WndProc
  Const FUNC_EBM              As String = "EbMode"                                      'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
  Const FUNC_SWL              As String = "SetWindowLongA"                              'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
  Const MOD_USER              As String = "user32"                                      'Location of the SetWindowLongA & CallWindowProc functions
  Const MOD_VBA5              As String = "vba5"                                        'Location of the EbMode function if running VB5
  Const MOD_VBA6              As String = "vba6"                                        'Location of the EbMode function if running VB6
  Const PATCH_01              As Long = 18                                              'Code buffer offset to the location of the relative address to EbMode
  Const PATCH_02              As Long = 68                                              'Address of the previous WndProc
  Const PATCH_03              As Long = 78                                              'Relative address of SetWindowsLong
  Const PATCH_06              As Long = 116                                             'Address of the previous WndProc
  Const PATCH_07              As Long = 121                                             'Relative address of CallWindowProc
  Const PATCH_0A              As Long = 186                                             'Address of the owner object
  Static aBuf(1 To CODE_LEN)  As Byte                                                   'Static code buffer byte array
  Static pCWP                 As Long                                                   'Address of the CallWindowsProc
  Static pEbMode              As Long                                                   'Address of the EbMode IDE break/stop/running function
  Static pSWL                 As Long                                                   'Address of the SetWindowsLong function
  Dim I                       As Long                                                   'Loop index
  Dim j                       As Long                                                   'Loop index
  Dim nSubIdx                 As Long                                                   'Subclass data index
  Dim sHex                    As String                                                 'Hex code string
  
'If it's the first time through here..
  If aBuf(1) = 0 Then
  
'The hex pair machine code representation.
    sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
           "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
           "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
           "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"

'Convert the string from hex pairs to bytes and store in the static machine code buffer
    I = 1
    Do While j < CODE_LEN
      j = j + 1
      aBuf(j) = Val("&H" & Mid$(sHex, I, 2))                                            'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
      I = I + 2
    Loop                                                                                'Next pair of hex characters
    
'Get API function addresses
    If Subclass_InIDE Then                                                              'If we're running in the VB IDE
      aBuf(16) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      aBuf(17) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                           'Get the address of EbMode in vba6.dll
      If pEbMode = 0 Then                                                               'Found?
        pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                                         'VB5 perhaps
      End If
    End If
    
    pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                                'Get the address of the CallWindowsProc function
    pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                                'Get the address of the SetWindowLongA function
    ReDim sc_aSubData(0 To 0) As tSubData                                               'Create the first sc_aSubData element
  Else
    nSubIdx = zIdx(lng_hWnd, True)
    If nSubIdx = -1 Then                                                                'If an sc_aSubData element isn't being re-cycled
      nSubIdx = UBound(sc_aSubData()) + 1                                               'Calculate the next element
      ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                              'Create a new sc_aSubData element
    End If
    
    Subclass_Start = nSubIdx
  End If

  With sc_aSubData(nSubIdx)
    .hwnd = lng_hWnd                                                                    'Store the hWnd
    .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                                       'Allocate memory for the machine code WndProc
    .nAddrOrig = SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrSub)                          'Set our WndProc in place
    Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)                              'Copy the machine code from the static byte array to the code array in sc_aSubData
    Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                                        'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
    Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                                     'Original WndProc address for CallWindowProc, call the original WndProc
    Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                                           'Patch the relative address of the SetWindowLongA api function
    Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                                     'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
    Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                                           'Patch the relative address of the CallWindowProc api function
    Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))                                     'Patch the address of this object instance into the static machine code buffer
  End With
Errs:
End Function

'Stop all subclassing
Private Sub Subclass_StopAll()
On Error GoTo Errs
  Dim I As Long
  
  I = UBound(sc_aSubData())                                                             'Get the upper bound of the subclass data array
  Do While I >= 0                                                                       'Iterate through each element
    With sc_aSubData(I)
      If .hwnd <> 0 Then                                                                'If not previously Subclass_Stop'd
        Call Subclass_Stop(.hwnd)                                                       'Subclass_Stop
      End If
    End With
    I = I - 1                                                                           'Next element
  Loop
Errs:
End Sub

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
On Error GoTo Errs
'Parameters:
  'lng_hWnd  - The handle of the window to stop being subclassed
  With sc_aSubData(zIdx(lng_hWnd))
    Call SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrOrig)                                 'Restore the original WndProc
    Call zPatchVal(.nAddrSub, PATCH_05, 0)                                              'Patch the Table B entry count to ensure no further 'before' callbacks
    Call zPatchVal(.nAddrSub, PATCH_09, 0)                                              'Patch the Table A entry count to ensure no further 'after' callbacks
    Call GlobalFree(.nAddrSub)                                                          'Release the machine code memory
    .hwnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
    .nMsgCntB = 0                                                                       'Clear the before table
    .nMsgCntA = 0                                                                       'Clear the after table
    Erase .aMsgTblB                                                                     'Erase the before table
    Erase .aMsgTblA                                                                     'Erase the after table
  End With
Errs:
End Sub

'=======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.

'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
On Error GoTo Errs
  Dim nEntry  As Long                                                                   'Message table entry index
  Dim nOff1   As Long                                                                   'Machine code buffer offset 1
  Dim nOff2   As Long                                                                   'Machine code buffer offset 2
  
  If uMsg = ALL_MESSAGES Then                                                           'If all messages
    nMsgCnt = ALL_MESSAGES                                                              'Indicates that all messages will callback
  Else                                                                                  'Else a specific message number
    Do While nEntry < nMsgCnt                                                           'For each existing entry. NB will skip if nMsgCnt = 0
      nEntry = nEntry + 1
      
      If aMsgTbl(nEntry) = 0 Then                                                       'This msg table slot is a deleted entry
        aMsgTbl(nEntry) = uMsg                                                          'Re-use this entry
        Exit Sub                                                                        'Bail
      ElseIf aMsgTbl(nEntry) = uMsg Then                                                'The msg is already in the table!
        Exit Sub                                                                        'Bail
      End If
    Loop                                                                                'Next entry

    nMsgCnt = nMsgCnt + 1                                                               'New slot required, bump the table entry count
    ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                        'Bump the size of the table.
    aMsgTbl(nMsgCnt) = uMsg                                                             'Store the message number in the table
  End If

  If When = eMsgWhen.MSG_BEFORE Then                                                    'If before
    nOff1 = PATCH_04                                                                    'Offset to the Before table
    nOff2 = PATCH_05                                                                    'Offset to the Before table entry count
  Else                                                                                  'Else after
    nOff1 = PATCH_08                                                                    'Offset to the After table
    nOff2 = PATCH_09                                                                    'Offset to the After table entry count
  End If

  If uMsg <> ALL_MESSAGES Then
    Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                                    'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
  End If
  Call zPatchVal(nAddr, nOff2, nMsgCnt)                                                 'Patch the appropriate table entry count
Errs:
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
  zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
  Debug.Assert zAddrFunc                                                                'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Worker sub for Subclass_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
On Error GoTo Errs
  Dim nEntry As Long
  
  If uMsg = ALL_MESSAGES Then                                                           'If deleting all messages
    nMsgCnt = 0                                                                         'Message count is now zero
    If When = eMsgWhen.MSG_BEFORE Then                                                  'If before
      nEntry = PATCH_05                                                                 'Patch the before table message count location
    Else                                                                                'Else after
      nEntry = PATCH_09                                                                 'Patch the after table message count location
    End If
    Call zPatchVal(nAddr, nEntry, 0)                                                    'Patch the table message count to zero
  Else                                                                                  'Else deleteting a specific message
    Do While nEntry < nMsgCnt                                                           'For each table entry
      nEntry = nEntry + 1
      If aMsgTbl(nEntry) = uMsg Then                                                    'If this entry is the message we wish to delete
        aMsgTbl(nEntry) = 0                                                             'Mark the table slot as available
        Exit Do                                                                         'Bail
      End If
    Loop                                                                                'Next entry
  End If
Errs:
End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
On Error GoTo Errs
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
  zIdx = UBound(sc_aSubData)
  Do While zIdx >= 0                                                                    'Iterate through the existing sc_aSubData() elements
    With sc_aSubData(zIdx)
      If .hwnd = lng_hWnd Then                                                          'If the hWnd of this element is the one we're looking for
        If Not bAdd Then                                                                'If we're searching not adding
          Exit Function                                                                 'Found
        End If
      ElseIf .hwnd = 0 Then                                                             'If this an element marked for reuse.
        If bAdd Then                                                                    'If we're adding
          Exit Function                                                                 'Re-use it
        End If
      End If
    End With
    zIdx = zIdx - 1                                                                     'Decrement the index
  Loop
  
'  If Not bAdd Then
'    Debug.Assert False                                                                  'hWnd not found, programmer error
'  End If
Errs:

'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
  zSetTrue = True
  bValue = True
End Function

'END Subclassing Code===================================================================================


