VERSION 5.00
Begin VB.UserControl TygemButton 
   ClientHeight    =   1305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1800
   BeginProperty Font 
      Name            =   "±¼¸²"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   KeyPreview      =   -1  'True
   ScaleHeight     =   1305
   ScaleWidth      =   1800
   ToolboxBitmap   =   "TygemButton.ctx":0000
   Windowless      =   -1  'True
   Begin VB.Timer tmrMouse 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1200
      Top             =   0
   End
   Begin VB.Line lSplit 
      Visible         =   0   'False
      X1              =   960
      X2              =   960
      Y1              =   480
      Y2              =   960
   End
   Begin VB.Image imgOverlay 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1065
   End
   Begin VB.Shape pgFocusRect 
      BorderColor     =   &H00404040&
      BorderStyle     =   3  'Á¡
      Height          =   255
      Left            =   1200
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Left            =   1200
      Top             =   600
      Width           =   240
   End
   Begin VB.Image imgSkin 
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "TygemButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const m_def_Enabled = True
Dim m_Enabled As Boolean

Const m_def_Caption = ""
Dim m_Caption As String

'Const m_def_BackColor = &H8000000F
'Dim m_BackColor As OLE_COLOR

Const m_def_FontName = "±¼¸²"
Dim m_FontName As String

Const m_def_FontSize = 9
Dim m_FontSize As Integer

Const m_def_SplitLeft = False
Dim m_SplitLeft As Boolean

Const m_def_SplitRight = False
Dim m_SplitRight As Boolean

Const m_def_Default = False
Dim m_Default As Boolean

Const m_def_Skin = 1
Dim m_Skin As ButtonSkin

Enum ButtonIconPositions
    IconPositionLeft = 0
    IconPositionRight = 1
    IconPositionCenter = 4
End Enum

Const m_def_IconPosition = 0
Dim m_IconPosition As ButtonIconPositions

Enum ButtonSkin
    System = 0
    LiveBaduk = 1
    Diskeeper = 2
    Blue = 3
    Green = 4
End Enum

Enum ButtonState
    Normal = 1
    Hover = 2
    Pressed = 3
    Disabled = 4
    Focused = 5
End Enum

Dim m_Icon As IPictureDisp
Dim IconLeft%, IconTop%

Dim DrawNormalState As ButtonState
Dim IsPressed As Boolean

Event Click()
'Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private bHovering As Boolean
Private bMouseDown As Boolean

Public CommandButtonControlHandle As Long

Private Sub MouseOut()
    If bMouseDown Then Exit Sub
    bHovering = False
    If DrawNormalState = Focused Then
        lblCaption.ForeColor = ButtonSkinCaptionColor((m_Skin - 1) * 5 + 5)
    Else
        lblCaption.ForeColor = ButtonSkinCaptionColor((m_Skin - 1) * 5 + 1)
    End If
    DrawSkin DrawNormalState
    tmrMouse.Enabled = False
End Sub

Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    SetEnabled
End Property

Private Sub SetEnabled()
    If Not m_Enabled Then tmrMouse.Enabled = False
    If m_Enabled Then
        If DrawNormalState = Focused Then
            lblCaption.ForeColor = ButtonSkinCaptionColor((m_Skin - 1) * 5 + 5)
        Else
            lblCaption.ForeColor = ButtonSkinCaptionColor((m_Skin - 1) * 5 + 1)
        End If
        DrawNormalState = Normal
    Else
        lblCaption.ForeColor = ButtonSkinCaptionColor((m_Skin - 1) * 5 + 4)
        DrawNormalState = Disabled
    End If
    DrawSkin DrawNormalState
End Sub

Property Get Default() As Boolean
    Default = m_Default
End Property

Property Let Default(New_Default As Boolean)
    m_Default = New_Default
    DrawSkin DrawNormalState
End Property

Property Get SplitLeft() As Boolean
    SplitLeft = m_SplitLeft
End Property

Property Let SplitLeft(ByVal New_SplitLeft As Boolean)
    m_SplitLeft = New_SplitLeft
    PropertyChanged "SplitLeft"
    SetSplitButton
End Property

Property Get SplitRight() As Boolean
    SplitRight = m_SplitRight
End Property

Property Let SplitRight(ByVal New_SplitRight As Boolean)
    m_SplitRight = New_SplitRight
    PropertyChanged "SplitRight"
    SetSplitButton
End Property

Private Sub SetSplitButton()
    lSplit.Visible = m_SplitLeft
    DrawSkin DrawNormalState
End Sub

Property Get Skin() As ButtonSkin
    Skin = m_Skin
End Property

Property Let Skin(New_Skin As ButtonSkin)
    m_Skin = New_Skin
    RefreshSkin True
End Property

Property Get Caption() As String
    Caption = m_Caption
End Property

Property Let Caption(ByVal New_Caption As String)
    m_Caption = Trim$(New_Caption)
    PropertyChanged "Caption"
    If Not m_Icon Is Nothing Then
        Select Case m_IconPosition
            Case IconPositionRight
                lblCaption = m_Caption & "  "
            Case IconPositionCenter
                lblCaption = m_Caption
            Case Else
                lblCaption = "  " & m_Caption
        End Select
    Else
        lblCaption = m_Caption
    End If
End Property

Property Get FontName() As String
    FontName = m_FontName
End Property

Property Let FontName(ByVal New_FontName As String)
    m_FontName = New_FontName
    PropertyChanged "FontName"
    SetCaptionFont
End Property

Private Sub SetCaptionFont()
    If m_Skin = 0 Then Exit Sub
    lblCaption.Font.Name = m_FontName
    lblCaption.Font.Bold = ButtonSkinBold(m_Skin)
    lblCaption.Font.Italic = False
    UserControl.Font.Name = lblCaption.Font.Name
    UserControl.Font.Bold = lblCaption.Font.Bold
    UserControl.Font.Italic = lblCaption.Font.Italic
End Sub

Property Get FontSize() As String
    FontSize = m_FontSize
End Property

Property Let FontSize(ByVal New_FontSize As String)
    m_FontSize = New_FontSize
    PropertyChanged "FontSize"
    lblCaption.Font.Size = m_FontSize
    UserControl.Font.Size = m_FontSize
End Property

Property Get ButtonIcon() As IPictureDisp
    Set ButtonIcon = m_Icon
End Property

Property Set ButtonIcon(ByVal New_Icon As IPictureDisp)
    Set m_Icon = New_Icon
    PropertyChanged "ButtonIcon"
    SetIcon
End Property

Property Get ButtonIconPosition() As ButtonIconPositions
    ButtonIconPosition = m_IconPosition
End Property

Property Let ButtonIconPosition(ByVal New_ButtonIconPosition As ButtonIconPositions)
    m_IconPosition = New_ButtonIconPosition
    If Not m_Icon Is Nothing Then imgIcon.Left = GetIconPosition()
End Property

Private Function GetIconPosition() As ButtonIconPositions
    Select Case m_IconPosition
        Case IconPositionRight
            GetIconPosition = UserControl.Width - 2 * Screen.TwipsPerPixelX - ScaleX(m_Icon.Width, vbHimetric, vbTwips)
        Case IconPositionCenter
            GetIconPosition = UserControl.Width / 2 - ScaleX(m_Icon.Width, vbHimetric, vbTwips) / 2
        Case Else
            GetIconPosition = 3 * Screen.TwipsPerPixelX
    End Select
End Function

Private Sub SetIcon()
    If Not m_Icon Is Nothing Then
        IconTop = UserControl.Height / 2 - ScaleY(m_Icon.Height, vbHimetric, vbTwips) / 2
        imgIcon.Top = IconTop
        IconLeft = GetIconPosition()
        imgIcon.Left = IconLeft
        Select Case m_IconPosition
            Case IconPositionRight
                lblCaption = Trim$(lblCaption) & "  "
            Case IconPositionCenter
                lblCaption = Trim$(lblCaption)
            Case Else
                lblCaption = "  " & Trim$(lblCaption)
        End Select
    End If
    Set imgIcon.Picture = m_Icon
End Sub

'Property Get BackColor() As OLE_COLOR
'    BackColor = UserControl.BackColor
'End Property

'Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
'    m_BackColor = New_BackColor
'    UserControl.BackColor = New_BackColor
'    PropertyChanged "BackColor"
'End Property

'Private Sub imgOverlay_Click()
'    If m_Enabled Then RaiseEvent Click
'End Sub

Private Function IsHovering() As Boolean
    Static lpPos As POINTAPI
    Static lhWnd As Long
    GetCursorPos lpPos
    lhWnd = WindowFromPoint(lpPos.X, lpPos.Y)
    IsHovering = (lhWnd = CommandButtonControlHandle)
End Function

Private Sub tmrMouse_Timer()
    Static lpPos As POINTAPI
    Static lhWnd As Long
    GetCursorPos lpPos
    lhWnd = WindowFromPoint(lpPos.X, lpPos.Y)
    If (Not IsHovering()) And bHovering Then MouseOut
End Sub

Private Sub UserControl_GotFocus()
    pgFocusRect.Visible = True
End Sub

Private Sub UserControl_Initialize()
    bMouseDown = False
    DrawNormalState = Normal
    IsPressed = False
End Sub

Sub ShowAsPressed()
    lblCaption.ForeColor = ButtonSkinCaptionColor((m_Skin - 1) * 5 + 3)
    lblCaption.Tag = "mousedown"
    If ButtonSkinInsetLabel(m_Skin) Then
        lblCaption.Left = Screen.TwipsPerPixelX
        lblCaption.Top = (UserControl.Height - UserControl.TextHeight(lblCaption)) / 2 + 5 + Screen.TwipsPerPixelY
        imgIcon.Left = IconLeft + Screen.TwipsPerPixelX
        imgIcon.Top = IconTop + Screen.TwipsPerPixelY
    End If
    IsPressed = True
    DrawSkin Pressed
End Sub

Sub ShowAsUnpressed()
    lblCaption.Tag = ""
    If ButtonSkinInsetLabel(m_Skin) Then
        lblCaption.Left = 0
        lblCaption.Top = (UserControl.Height - UserControl.TextHeight(lblCaption)) / 2
        imgIcon.Left = IconLeft
        imgIcon.Top = IconTop
    End If
    IsPressed = False
    If DrawNormalState = Focused Then
        lblCaption.ForeColor = ButtonSkinCaptionColor((m_Skin - 1) * 5 + 5)
    Else
        lblCaption.ForeColor = ButtonSkinCaptionColor((m_Skin - 1) * 5 + 1)
    End If
    DrawSkin DrawNormalState
End Sub

Private Sub imgOverlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Not m_Enabled) Or Button <> 1 Then Exit Sub
    bMouseDown = True
    ShowAsPressed
    'RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
 
Private Sub imgOverlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not m_Enabled Then Exit Sub
    tmrMouse.Enabled = -1
    If Not bHovering Then
        bHovering = True
        DrawSkin Hover
        If lblCaption.Tag <> "mousedown" Then lblCaption.ForeColor = ButtonSkinCaptionColor((m_Skin - 1) * 5 + 2)
    End If
    'RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
 
Private Sub imgOverlay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Not m_Enabled) Or Button <> 1 Then Exit Sub
    bMouseDown = False
    ShowAsUnpressed
    bHovering = False
    If IsHovering() Then RaiseEvent Click
    'RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_InitProperties()
    'm_Caption = Ambient.DisplayName
    'm_BackColor = &H8000000F
    m_Enabled = True
    'm_SplitLeft = False
    'm_SplitRight = False
    'UserControl.BackColor = &H8000000F
    'lblCaption.Caption = m_Caption
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not m_Enabled Then Exit Sub
    If Not bMouseDown Then
        If KeyCode = 32 Then ShowAsPressed
    End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not m_Enabled Then Exit Sub
    If Not bMouseDown Then
        If KeyCode = 32 Then
            ShowAsUnpressed
            RaiseEvent Click
        End If
    End If
End Sub

Private Sub UserControl_LostFocus()
    pgFocusRect.Visible = False
End Sub

Private Sub RefreshSkin(Optional Redraw As Boolean = False)
    If m_Skin = 0 Then Exit Sub
    lblCaption.Font.Bold = ButtonSkinBold(m_Skin)
    UserControl.Font.Bold = lblCaption.Font.Bold
    If DrawNormalState = Focused Then
        lblCaption.ForeColor = ButtonSkinCaptionColor((m_Skin - 1) * 5 + 5)
    ElseIf Not m_Enabled Then
        lblCaption.ForeColor = ButtonSkinCaptionColor((m_Skin - 1) * 5 + 4)
    Else
        lblCaption.ForeColor = ButtonSkinCaptionColor((m_Skin - 1) * 5 + 1)
    End If
    lSplit.BorderColor = ButtonSkinSplitColor(m_Skin)
    If Redraw Then DrawSkin DrawNormalState
End Sub

Private Sub DrawSkin(Optional ByVal State As ButtonState = Normal)
    Dim pic As StdPicture
    If State = Normal And m_Default Then State = Focused
    Set pic = ButtonSkinTexture(State + (m_Skin - 1) * 5)
    If pic Is Nothing Or m_Skin = 0 Then Exit Sub
    
    Dim srcW As Long, srcH As Long
    Dim dstW As Long, dstH As Long
    Dim B As Long

    srcW = ScaleX(pic.Width, vbHimetric, vbPixels)
    srcH = ScaleY(pic.Height, vbHimetric, vbPixels)
    dstW = ScaleX(UserControl.Width, vbTwips, vbPixels)
    dstH = ScaleY(UserControl.Height, vbTwips, vbPixels)
    B = ButtonSkinBorder(m_Skin)

    Dim effectiveSrcW As Long
    Dim srcXOffset As Long

    If m_SplitLeft Then
        effectiveSrcW = srcW - 16
        srcXOffset = 0
    ElseIf m_SplitRight Then
        effectiveSrcW = 16
        srcXOffset = srcW - 16
    Else
        effectiveSrcW = srcW
        srcXOffset = 0
    End If

    Dim hScreenDC As Long, hMemDC As Long, hBmp As Long, hOld As Long
    hScreenDC = GetDC(0)
    hMemDC = CreateCompatibleDC(hScreenDC)
    hBmp = CreateCompatibleBitmap(hScreenDC, dstW, dstH)
    hOld = SelectObject(hMemDC, hBmp)

    Dim hSrcBmp As Long
    hSrcBmp = pic.Handle
    If hSrcBmp = 0& Then GoTo Cleanup

    Dim hSrcDC As Long, hOldSrc As Long
    hSrcDC = CreateCompatibleDC(hScreenDC)
    hOldSrc = SelectObject(hSrcDC, hSrcBmp)

    If effectiveSrcW > 2 * B And srcH > 2 * B Then
        StretchBlt hMemDC, 0, 0, B, B, hSrcDC, srcXOffset, 0, B, B, vbSrcCopy
        StretchBlt hMemDC, 0, dstH - B, B, B, hSrcDC, srcXOffset, srcH - B, B, B, vbSrcCopy
        StretchBlt hMemDC, dstW - B, 0, B, B, hSrcDC, srcXOffset + effectiveSrcW - B, 0, B, B, vbSrcCopy
        StretchBlt hMemDC, dstW - B, dstH - B, B, B, hSrcDC, srcXOffset + effectiveSrcW - B, srcH - B, B, B, vbSrcCopy
    End If

    StretchBlt hMemDC, 0, B, B, dstH - 2 * B, hSrcDC, srcXOffset, B, B, srcH - 2 * B, vbSrcCopy
    StretchBlt hMemDC, dstW - B, B, B, dstH - 2 * B, hSrcDC, srcXOffset + effectiveSrcW - B, B, B, srcH - 2 * B, vbSrcCopy
    StretchBlt hMemDC, B, 0, dstW - 2 * B, B, hSrcDC, srcXOffset + B, 0, effectiveSrcW - 2 * B, B, vbSrcCopy   ' top
    StretchBlt hMemDC, B, dstH - B, dstW - 2 * B, B, hSrcDC, srcXOffset + B, srcH - B, effectiveSrcW - 2 * B, B, vbSrcCopy

    StretchBlt hMemDC, B, B, dstW - 2 * B, dstH - 2 * B, hSrcDC, srcXOffset + B, B, effectiveSrcW - 2 * B, srcH - 2 * B, vbSrcCopy

    SelectObject hSrcDC, hOldSrc
    DeleteDC hSrcDC

    Set imgSkin.Picture = CreatePicture(hBmp, 1&)

Cleanup:
    SelectObject hMemDC, hOld
    DeleteDC hMemDC
    ReleaseDC 0&, hScreenDC
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    imgOverlay.Width = UserControl.Width
    imgOverlay.Height = UserControl.Height
    SetSplitButton
    lblCaption.Left = 0
    lblCaption.Top = (UserControl.Height - UserControl.TextHeight(lblCaption)) / 2
    lblCaption.Width = UserControl.Width
    IconTop = UserControl.Height / 2 - ScaleY(m_Icon.Height, vbHimetric, vbTwips) / 2
    imgIcon.Top = IconTop
    IconLeft = GetIconPosition()
    imgIcon.Left = IconLeft
    pgFocusRect.Top = 30
    pgFocusRect.Left = 30
    pgFocusRect.Width = UserControl.Width - 60
    pgFocusRect.Height = UserControl.Height - 60
    lSplit.Y1 = 15
    lSplit.Y2 = UserControl.Height - 15
    lSplit.X1 = UserControl.Width - 15
    lSplit.X2 = lSplit.X1
    DrawSkin DrawNormalState
End Sub

Function GetFocusRect() As Shape
    Set GetFocusRect = pgFocusRect
End Function

'Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'm_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    'm_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    'm_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    'm_FontName = PropBag.ReadProperty("FontName", m_def_FontName)
    'm_FontSize = PropBag.ReadProperty("FontSize", m_def_FontSize)
    'Set m_Icon = PropBag.ReadProperty("ButtonIcon", Nothing)
    'm_SplitLeft = PropBag.ReadProperty("SplitLeft", m_def_SplitLeft)
    'm_SplitRight = PropBag.ReadProperty("SplitRight", m_def_SplitRight)
    'SetEnabled
    'lblCaption.Caption = Trim$(m_Caption)
    'UserControl.BackColor = m_BackColor
    'SetIcon
    'SetSplitLeft
    'SetSplitRight
    'SetCaptionFont
'End Sub

'Private Sub UserControl_Terminate()
'    tmrMouse.Enabled = 0
'End Sub

'Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    'Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    'Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    'Call PropBag.WriteProperty("FontName", m_FontName, m_def_FontName)
    'Call PropBag.WriteProperty("FontSize", m_FontSize, m_def_FontSize)
    'Call PropBag.WriteProperty("SplitLeft", m_SplitLeft, m_def_SplitLeft)
    'Call PropBag.WriteProperty("SplitRight", m_SplitRight, m_def_SplitRight)
    'On Error Resume Next
    'Call PropBag.WriteProperty("ButtonIcon", m_Icon, Nothing)
'End Sub

