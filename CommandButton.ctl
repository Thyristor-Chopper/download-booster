VERSION 5.00
Begin VB.UserControl CommandButton 
   ClientHeight    =   1005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1770
   DefaultCancel   =   -1  'True
   BeginProperty Font 
      Name            =   "±¼¸²"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1005
   ScaleWidth      =   1770
   Begin VB.CommandButton cmdButton 
      Caption         =   "Command1"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin prjDownloadBooster.TygemButton tygButton 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
   End
End
Attribute VB_Name = "CommandButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IBSSubclass

Private Declare Function ActivateVisualStyles Lib "uxtheme.dll" Alias "SetWindowTheme" (ByVal hWnd As Long, Optional ByVal pszSubAppName As Long = 0&, Optional ByVal pszSubIdList As Long = 0&) As Long
Private Declare Function DeactivateVisualStyles Lib "uxtheme.dll" Alias "SetWindowTheme" (ByVal hWnd As Long, Optional ByRef pszSubAppName As String = " ", Optional ByRef pszSubIdList As String = " ") As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDestination As Long, ByVal pSource As Long, ByVal Length As Long)
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const BS_DEFPUSHBUTTON As Long = &H1

Const GWL_STYLE As Long = -16&

Const BM_SETIMAGE = &HF7
Const IMAGE_BITMAP = 0
Const IMAGE_ICON = 1

Const BCM_FIRST As Long = &H1600
Const BCM_GETIDEALSIZE As Long = (BCM_FIRST + 1)
Const BCM_SETIMAGELIST As Long = (BCM_FIRST + 2)
Const BCM_GETIMAGELIST As Long = (BCM_FIRST + 3)

Private Type BUTTON_IMAGELIST
    hImageList As Long
    RCMargin As RECT
    uAlign As Long
End Type

Private Const BUTTON_IMAGELIST_ALIGN_LEFT As Long = 0
Private Const BUTTON_IMAGELIST_ALIGN_RIGHT As Long = 1
Private Const BUTTON_IMAGELIST_ALIGN_CENTER As Long = 4
Public Enum IconAlignment
    IconAlignmentLeft = BUTTON_IMAGELIST_ALIGN_LEFT
    IconAlignmentRight = BUTTON_IMAGELIST_ALIGN_RIGHT
    IconAlignmentCenter = BUTTON_IMAGELIST_ALIGN_CENTER
End Enum

Const m_def_Enabled = True
Dim m_Enabled As Boolean

Dim m_Caption As String

Const m_def_BackColor = &H8000000F
Dim m_BackColor As OLE_COLOR

Const m_def_IsTygemButton As Boolean = False
Dim m_IsTygemButton As Boolean

Const m_def_VisualStyles As Boolean = True
Dim m_VisualStyles As Boolean

Const m_def_RoundButton As Boolean = False
Dim m_RoundButton As Boolean

Const m_def_Transparent As Boolean = False
Dim m_Transparent As Boolean

Const m_def_IconPosition = IconAlignment.IconAlignmentLeft
Dim m_IconPosition As IconAlignment

Dim m_ImageList As ImageList

Dim m_Font As StdFont

Private CommandButtonTransparentBrush As Long

Event Click()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)

Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    SetEnabled
End Property

Private Sub SetEnabled()
    cmdButton.Enabled = m_Enabled
    tygButton.Enabled = m_Enabled
End Sub

Property Get IsTygemButton() As Boolean
    IsTygemButton = m_IsTygemButton
End Property

Property Let IsTygemButton(ByVal New_IsTygemButton As Boolean)
    m_IsTygemButton = New_IsTygemButton
    PropertyChanged "IsTygemButton"
    SetIsTygemButton
End Property

Private Sub SetIsTygemButton()
    tygButton.Visible = m_IsTygemButton
    SetRgn
End Sub

Property Get VisualStyles() As Boolean
    VisualStyles = m_VisualStyles
End Property

Property Let VisualStyles(ByVal New_VisualStyles As Boolean)
    m_VisualStyles = New_VisualStyles
    PropertyChanged "VisualStyles"
    SetVisualStyles
End Property

Property Get RoundButton() As Boolean
    RoundButton = m_RoundButton
End Property

Property Let RoundButton(ByVal New_RoundButton As Boolean)
    m_RoundButton = New_RoundButton
    PropertyChanged "RoundButton"
    SetRgn
End Property

Property Get Transparent() As Boolean
    Transparent = m_Transparent
End Property

Property Let Transparent(ByVal New_Transparent As Boolean)
    m_Transparent = New_Transparent
    PropertyChanged "Transparent"
End Property

Private Sub SetVisualStyles()
    If m_VisualStyles Then
        ActivateVisualStyles cmdButton.hWnd
    Else
        DeactivateVisualStyles cmdButton.hWnd
    End If
    SetRgn
End Sub

Property Get Default() As Boolean
    Default = Extender.Default
End Property

Property Let Default(ByVal New_Default As Boolean)
    Extender.Default = New_Default
    SetDefault
End Property

Private Sub SetDefault()
    Dim CurrentStyle&
    CurrentStyle = GetWindowLong(cmdButton.hWnd, GWL_STYLE)
    SetWindowLong cmdButton.hWnd, GWL_STYLE, IIf(Ambient.DisplayAsDefault, CurrentStyle Or BS_DEFPUSHBUTTON, CurrentStyle And (Not BS_DEFPUSHBUTTON))
End Sub

Property Get Cancel() As Boolean
    Cancel = Extender.Cancel
End Property

Property Let Cancel(ByVal New_Cancel As Boolean)
    Extender.Cancel = New_Cancel
End Property

Property Get Caption() As String
    Caption = m_Caption
End Property

Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    SetCaption
End Property

Private Sub SetCaption()
    cmdButton.Caption = m_Caption
    tygButton.Caption = m_Caption
End Sub

Property Get ImageList() As ImageList
    Set ImageList = m_ImageList
End Property

Property Set ImageList(New_ImageList As ImageList)
    Set m_ImageList = New_ImageList
    SetImageList
End Property

Private Sub SetImageList()
    If m_ImageList Is Nothing Then
        SendMessage cmdButton.hWnd, BM_SETIMAGE, IMAGE_BITMAP, ByVal 0&
        SendMessage cmdButton.hWnd, BM_SETIMAGE, IMAGE_ICON, ByVal 0&
    Else
        Dim BTNIML As BUTTON_IMAGELIST
        BTNIML.hImageList = m_ImageList.hImageList
        If m_IconPosition = IconAlignmentLeft Then
            BTNIML.RCMargin.Left = 0
        ElseIf IconAlignmentRight Then
            BTNIML.RCMargin.Right = 0
        End If
        BTNIML.uAlign = m_IconPosition
        SendMessage cmdButton.hWnd, BCM_SETIMAGELIST, 0&, ByVal VarPtr(BTNIML)
        UserControl.Refresh
        cmdButton.Refresh
    End If
End Sub

Property Get IconPosition() As IconAlignment
    IconPosition = m_IconPosition
End Property

Property Let IconPosition(ByVal New_IconPosition As IconAlignment)
    m_IconPosition = New_IconPosition
    PropertyChanged "IconPosition"
    SetImageList
End Property

Property Get Font() As StdFont
    Set Font = m_Font
End Property

Property Set Font(ByVal New_Font As StdFont)
    Set m_Font = New_Font
    PropertyChanged "Font"
    SetFont
End Property

Private Sub SetFont()
    If m_Font Is Nothing Then
        Set m_Font = New StdFont
        m_Font.Name = "±¼¸²"
        m_Font.Size = 9
    End If
    Set cmdButton.Font = m_Font
    tygButton.FontName = m_Font.Name
    tygButton.FontSize = m_Font.Size
End Sub

Property Get BackColor() As OLE_COLOR
    BackColor = cmdButton.BackColor
End Property

Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    SetBackColor
End Property

Private Sub SetBackColor()
    cmdButton.BackColor = m_BackColor
    'tygButton.BackColor = m_BackColor
End Sub

Private Sub cmdButton_Click()
    RaiseEvent Click
End Sub

Private Sub cmdButton_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub cmdButton_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub cmdButton_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub cmdButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub cmdButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub cmdButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub IBSSubclass_UnsubclassIt()
    DetachMessage Me, UserControl.hWnd, WM_CTLCOLORSTATIC
    DetachMessage Me, UserControl.hWnd, WM_CTLCOLORBTN
End Sub

Private Function IBSSubclass_WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
    On Error Resume Next
 
    Select Case uMsg
        Case WM_CTLCOLORSTATIC, WM_CTLCOLORBTN
            IBSSubclass_WindowProc = CallOldWindowProc(hWnd, uMsg, wParam, lParam)
            If m_Transparent = True Then
                SetBkMode wParam, 1&
                Dim hDCBmp As Long
                Dim hBmp As Long, hBmpOld As Long
                With UserControl
                    If Not CommandButtonTransparentBrush Then
                        hDCBmp = CreateCompatibleDC(wParam)
                        If hDCBmp Then
                            hBmp = CreateCompatibleBitmap(wParam, .ScaleWidth / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY)
                            If hBmp Then
                                Dim hWndParent As Long
                                hWndParent = GetParent(.hWnd)
                                If (GetWindowLong(hWndParent, GWL_EXSTYLE) And WS_EX_LAYOUTRTL) = WS_EX_LAYOUTRTL Then SetLayout hDCBmp, LAYOUT_RTL
                                hBmpOld = SelectObject(hDCBmp, hBmp)
                                Dim WndRect As RECT, P As POINTAPI
                                GetWindowRect .hWnd, WndRect
                                MapWindowPoints hWnd_DESKTOP, hWndParent, WndRect, 2&
                                P.X = WndRect.Left
                                P.Y = WndRect.Top
                                SetViewportOrgEx hDCBmp, -P.X, -P.Y, P
                                SendMessage hWndParent, WM_PAINT, hDCBmp, ByVal 0&
                                SetViewportOrgEx hDCBmp, P.X, P.Y, P
                                CommandButtonTransparentBrush = CreatePatternBrush(hBmp)
                                SelectObject hDCBmp, hBmpOld
                                DeleteObject hBmp
                            End If
                            DeleteDC hDCBmp
                        End If
                    End If
                End With
                If CommandButtonTransparentBrush Then IBSSubclass_WindowProc = CommandButtonTransparentBrush
            End If
            Exit Function
    End Select
    
    IBSSubclass_WindowProc = CallOldWindowProc(hWnd, uMsg, wParam, lParam)
End Function

Private Function IBSSubclass_MsgResponse(ByVal hWnd As Long, ByVal uMsg As Long) As EMsgResponse
    IBSSubclass_MsgResponse = emrConsume
End Function

Private Sub tygButton_Click()
    RaiseEvent Click
End Sub

Private Sub tygButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub tygButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub tygButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    RaiseEvent Click
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If PropertyName = "DisplayAsDefault" Then
        SetDefault
        cmdButton.Refresh
    End If
End Sub

Private Sub UserControl_Initialize()
    cmdButton.Top = 0
    cmdButton.Left = 0
    tygButton.Top = 0
    tygButton.Left = 0
    
    AttachMessage Me, UserControl.hWnd, WM_CTLCOLORSTATIC
    AttachMessage Me, UserControl.hWnd, WM_CTLCOLORBTN
End Sub

Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
    m_Caption = Ambient.DisplayName
    m_BackColor = m_def_BackColor
    Set m_ImageList = Nothing
    m_IsTygemButton = m_def_IsTygemButton
    cmdButton.Caption = m_Caption
    tygButton.Caption = m_Caption
    m_VisualStyles = m_def_VisualStyles
    Set m_Font = New StdFont
    m_Font.Name = "±¼¸²"
    m_Font.Size = 9
    m_RoundButton = m_def_RoundButton
    m_Transparent = m_def_Transparent
End Sub

Private Sub UserControl_Resize()
    cmdButton.Height = UserControl.Height
    tygButton.Height = UserControl.Height
    
    cmdButton.Width = UserControl.Width
    tygButton.Width = UserControl.Width
    
'    If (Not m_IsTygemButton) And m_Transparent Then
'        If CommandButtonTransparentBrush Then
'            DeleteObject CommandButtonTransparentBrush
'            CommandButtonTransparentBrush = 0&
'        End If
'        RedrawWindow cmdButton.hWnd, 0&, 0&, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE
'    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    Set m_ImageList = PropBag.ReadProperty("ImageList", Nothing)
    m_IconPosition = PropBag.ReadProperty("IconPosition", m_def_IconPosition)
    m_IsTygemButton = PropBag.ReadProperty("IsTygemButton", m_def_IsTygemButton)
    m_VisualStyles = PropBag.ReadProperty("VisualStyles", m_def_VisualStyles)
    Set m_Font = PropBag.ReadProperty("Font", Nothing)
    m_RoundButton = PropBag.ReadProperty("RoundButton", m_def_RoundButton)
    m_Transparent = PropBag.ReadProperty("Transparent", m_def_Transparent)
    
    SetEnabled
    SetCaption
    SetBackColor
    SetImageList
    SetIsTygemButton
    SetVisualStyles
    SetFont
    SetDefault
    
    SetRgn
End Sub

Private Sub UserControl_Terminate()
    DetachMessage Me, UserControl.hWnd, WM_CTLCOLORSTATIC
    DetachMessage Me, UserControl.hWnd, WM_CTLCOLORBTN
    If CommandButtonTransparentBrush Then
        DeleteObject CommandButtonTransparentBrush
        CommandButtonTransparentBrush = 0&
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Caption", m_Caption, Ambient.DisplayName)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ImageList", m_ImageList, Nothing)
    Call PropBag.WriteProperty("IsTygemButton", m_IsTygemButton, m_def_IsTygemButton)
    Call PropBag.WriteProperty("VisualStyles", m_VisualStyles, m_def_VisualStyles)
    Call PropBag.WriteProperty("Font", m_Font, Nothing)
    Call PropBag.WriteProperty("IconPosition", m_IconPosition, m_def_IconPosition)
    Call PropBag.WriteProperty("RoundButton", m_RoundButton, m_def_RoundButton)
    Call PropBag.WriteProperty("Transparent", m_Transparent, m_def_Transparent)
End Sub

Sub Refresh()
    If CommandButtonTransparentBrush Then
        DeleteObject CommandButtonTransparentBrush
        CommandButtonTransparentBrush = 0&
    End If
    UserControl.Refresh
    cmdButton.Refresh
    RedrawWindow UserControl.hWnd, 0&, 0&, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
    SetRgn
End Sub

Private Sub SetRgn()
    Dim RC As RECT
    Dim Rgn&, Rgn1&, Rgn2&, Rgn3&, Rgn4&, Rgn5&, Rgn6&, Rgn7&, Rgn8&
    
    If m_IsTygemButton Then
        GetWindowRect UserControl.hWnd, RC
        Rgn = CreateRectRgn(0, 0, RC.Right - RC.Left, RC.Bottom - RC.Top)
        Rgn1 = CreateRectRgn(0, 0, 2, 1) '¿ÞÂÊ À§
        Rgn2 = CreateRectRgn(0, 1, 1, 2)
        Rgn3 = CreateRectRgn(RC.Right - RC.Left - 2, 0, RC.Right - RC.Left, 1) '¿À¸¥ÂÊ À§
        Rgn4 = CreateRectRgn(RC.Right - RC.Left - 1, 1, RC.Right - RC.Left, 2)
        Rgn5 = CreateRectRgn(0, RC.Bottom - RC.Top - 1, 2, RC.Bottom - RC.Top) '¿ÞÂÊ ¾Æ·¡
        Rgn6 = CreateRectRgn(0, RC.Bottom - RC.Top - 2, 1, RC.Bottom - RC.Top - 1)
        Rgn7 = CreateRectRgn(RC.Right - RC.Left - 2, RC.Bottom - RC.Top - 1, RC.Right - RC.Left, RC.Bottom - RC.Top) '¿À¸¥ÂÊ ¾Æ·¡
        Rgn8 = CreateRectRgn(RC.Right - RC.Left - 1, RC.Bottom - RC.Top - 2, RC.Right - RC.Left, RC.Bottom - RC.Top - 1)
        CombineRgn Rgn, Rgn, Rgn1, RGN_DIFF
        CombineRgn Rgn, Rgn, Rgn2, RGN_DIFF
        CombineRgn Rgn, Rgn, Rgn3, RGN_DIFF
        CombineRgn Rgn, Rgn, Rgn4, RGN_DIFF
        CombineRgn Rgn, Rgn, Rgn5, RGN_DIFF
        CombineRgn Rgn, Rgn, Rgn6, RGN_DIFF
        CombineRgn Rgn, Rgn, Rgn7, RGN_DIFF
        CombineRgn Rgn, Rgn, Rgn8, RGN_DIFF
        SetWindowRgn UserControl.hWnd, Rgn, True
        DeleteObject Rgn
        DeleteObject Rgn1
        DeleteObject Rgn2
        DeleteObject Rgn3
        DeleteObject Rgn4
        DeleteObject Rgn5
        DeleteObject Rgn6
        DeleteObject Rgn7
        DeleteObject Rgn8
    ElseIf m_RoundButton And (Not m_VisualStyles) Then
        GetWindowRect UserControl.hWnd, RC
        Rgn = CreateRectRgn(0, 0, RC.Right - RC.Left, RC.Bottom - RC.Top)
        Rgn1 = CreateRectRgn(0, 0, 2, 1)
        Rgn2 = CreateRectRgn(0, 1, 1, 2)
        Rgn3 = CreateRectRgn(RC.Right - RC.Left - 2, 0, RC.Right - RC.Left, 1)
        Rgn4 = CreateRectRgn(RC.Right - RC.Left - 1, 1, RC.Right - RC.Left, 2)
        Rgn5 = CreateRectRgn(0, RC.Bottom - RC.Top - 1, 2, RC.Bottom - RC.Top)
        Rgn6 = CreateRectRgn(0, RC.Bottom - RC.Top - 2, 1, RC.Bottom - RC.Top - 1)
        Rgn7 = CreateRectRgn(RC.Right - RC.Left - 2, RC.Bottom - RC.Top - 1, RC.Right - RC.Left, RC.Bottom - RC.Top)
        Rgn8 = CreateRectRgn(RC.Right - RC.Left - 1, RC.Bottom - RC.Top - 2, RC.Right - RC.Left, RC.Bottom - RC.Top - 1)
        CombineRgn Rgn, Rgn, Rgn1, RGN_DIFF
        CombineRgn Rgn, Rgn, Rgn2, RGN_DIFF
        CombineRgn Rgn, Rgn, Rgn3, RGN_DIFF
        CombineRgn Rgn, Rgn, Rgn4, RGN_DIFF
        CombineRgn Rgn, Rgn, Rgn5, RGN_DIFF
        CombineRgn Rgn, Rgn, Rgn6, RGN_DIFF
        CombineRgn Rgn, Rgn, Rgn7, RGN_DIFF
        CombineRgn Rgn, Rgn, Rgn8, RGN_DIFF
        SetWindowRgn UserControl.hWnd, Rgn, True
        DeleteObject Rgn
        DeleteObject Rgn1
        DeleteObject Rgn2
        DeleteObject Rgn3
        DeleteObject Rgn4
        DeleteObject Rgn5
        DeleteObject Rgn6
        DeleteObject Rgn7
        DeleteObject Rgn8
    Else
        SetWindowRgn UserControl.hWnd, 0&, True
    End If
End Sub
