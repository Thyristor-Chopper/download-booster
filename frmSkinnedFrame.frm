VERSION 5.00
Begin VB.Form frmSkinnedFrame 
   Appearance      =   0  '평면
   BackColor       =   &H80000005&
   BorderStyle     =   0  '없음
   ClientHeight    =   2145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSkinnedFrame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timCaptionMonitor 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3720
      Top             =   960
   End
   Begin VB.Timer timMinimizeHover 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3720
      Top             =   1440
   End
   Begin VB.Timer timMaximizeHover 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3240
      Top             =   1440
   End
   Begin VB.Timer timCloseHover 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2760
      Top             =   1440
   End
   Begin VB.PictureBox pbRight 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '없음
      ClipControls    =   0   'False
      Height          =   930
      Left            =   4560
      MousePointer    =   9  'W E 크기 조정
      ScaleHeight     =   930
      ScaleWidth      =   75
      TabIndex        =   16
      Top             =   480
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox pbBottomMiddle 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '없음
      ClipControls    =   0   'False
      Height          =   60
      Left            =   240
      MousePointer    =   7  'N S크기 조정
      ScaleHeight     =   60
      ScaleWidth      =   2055
      TabIndex        =   15
      Top             =   1920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox pbLeft 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '없음
      ClipControls    =   0   'False
      Height          =   930
      Left            =   0
      MousePointer    =   9  'W E 크기 조정
      ScaleHeight     =   930
      ScaleWidth      =   75
      TabIndex        =   14
      Top             =   480
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox pbTopRight 
      BorderStyle     =   0  '없음
      ClipControls    =   0   'False
      Height          =   480
      Left            =   3360
      ScaleHeight     =   480
      ScaleWidth      =   1260
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   1260
      Begin VB.Image imgMinimizeButton 
         Height          =   315
         Left            =   120
         Top             =   90
         Width           =   315
      End
      Begin VB.Image imgMaximizeButton 
         Height          =   315
         Left            =   480
         Top             =   90
         Width           =   315
      End
      Begin VB.Image imgCloseButton 
         Height          =   315
         Left            =   840
         Top             =   90
         Width           =   315
      End
      Begin VB.Label lblResizeRight 
         BackStyle       =   0  '투명
         Height          =   495
         Left            =   1185
         MousePointer    =   9  'W E 크기 조정
         TabIndex        =   13
         Top             =   75
         Width           =   75
      End
      Begin VB.Label lblResizeTopRight 
         BackStyle       =   0  '투명
         Height          =   75
         Left            =   960
         MousePointer    =   6  'NE SW 크기 조정
         TabIndex        =   12
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblResizeTop 
         BackStyle       =   0  '투명
         Height          =   75
         Index           =   1
         Left            =   0
         MousePointer    =   7  'N S크기 조정
         TabIndex        =   11
         Top             =   0
         Width           =   960
      End
   End
   Begin VB.PictureBox pbTopLeft 
      BorderStyle     =   0  '없음
      ClipControls    =   0   'False
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   570
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   570
      Begin VB.Image imgControlMenu 
         Height          =   240
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lblResizeLeft 
         BackStyle       =   0  '투명
         Height          =   390
         Left            =   0
         MousePointer    =   9  'W E 크기 조정
         TabIndex        =   9
         Top             =   75
         Width           =   75
      End
      Begin VB.Label lblResizeTopLeft 
         BackStyle       =   0  '투명
         Height          =   75
         Left            =   0
         MousePointer    =   8  'NW SE 크기 조정
         TabIndex        =   8
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblResizeTop 
         BackStyle       =   0  '투명
         Height          =   75
         Index           =   2
         Left            =   240
         MousePointer    =   7  'N S크기 조정
         TabIndex        =   7
         Top             =   0
         Width           =   330
      End
   End
   Begin VB.PictureBox pbTopMiddle 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '없음
      ClipControls    =   0   'False
      Height          =   480
      Left            =   570
      ScaleHeight     =   480
      ScaleWidth      =   2250
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   2250
      Begin VB.Label lblResizeTop 
         BackStyle       =   0  '투명
         Height          =   75
         Index           =   0
         Left            =   0
         MousePointer    =   7  'N S크기 조정
         TabIndex        =   4
         Top             =   0
         Width           =   2295
      End
      Begin VB.Label lblCaption 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Tag             =   "nocolorsizechange"
         Top             =   150
         Width           =   2295
      End
      Begin VB.Label lblCaptionShadow 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   15
         TabIndex        =   5
         Tag             =   "nocolorsizechange"
         Top             =   165
         Width           =   2295
      End
   End
   Begin VB.PictureBox pbBottomLeft 
      BorderStyle     =   0  '없음
      ClipControls    =   0   'False
      Height          =   135
      Left            =   840
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox pbBottomRight 
      BorderStyle     =   0  '없음
      ClipControls    =   0   'False
      Height          =   135
      Left            =   1200
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "frmSkinnedFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#If DISABLEFRAMESKIN Then
#Else

Option Explicit

Public TargetForm As Object
Dim IsWindowActive As Byte
Dim CurrentSkin As WindowSkin
Dim Terminated As Boolean
Dim LastRect As RECT
Dim VisibilitySet As Boolean

Implements IBSSubclass

Sub Init(Target As Object)
    Terminated = False
    Set TargetForm = Target
    If CurrentWindowSkin > 0 Then SetSkin CurrentWindowSkin
End Sub

Sub ReloadSkin(Optional NewSkin As WindowSkin)
    If NewSkin = 0 Then NewSkin = CurrentSkin
    SetSkin NewSkin
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CurrentSkin > 0 Then SetSkin -1
    Set TargetForm = Nothing
End Sub

Sub IBSSubclass_UnsubclassIt()
    If InIDE Or (TargetForm Is Nothing) Then Exit Sub
    DetachMessage Me, TargetForm.hWnd, WM_NCPAINT
    DetachMessage Me, TargetForm.hWnd, WM_MOVE
    DetachMessage Me, TargetForm.hWnd, WM_NCCALCSIZE
    DetachMessage Me, TargetForm.hWnd, WM_NCHITTEST
    DetachMessage Me, TargetForm.hWnd, WM_NCACTIVATE
    DetachMessage Me, TargetForm.hWnd, WM_SIZE
End Sub

Private Function IBSSubclass_MsgResponse(ByVal hWnd As Long, ByVal uMsg As Long) As EMsgResponse
    IBSSubclass_MsgResponse = emrConsume
End Function

Private Function IBSSubclass_WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, wParam As Long, lParam As Long, bConsume As Boolean) As Long
    On Error Resume Next
    
    If Terminated Or CurrentSkin <= System Then GoTo defproc
    
    Dim RC As RECT
    
    Select Case uMsg
        Case WM_NCPAINT
            If IsChild() Then SetFramePos
        
            IBSSubclass_WindowProc = 0&
            Exit Function
        Case WM_MOVE
            SetFramePos
            
            IBSSubclass_WindowProc = 0&
            Exit Function
        Case WM_NCCALCSIZE
            If wParam <> 0 Then
                CopyMemory RC, ByVal lParam, Len(RC)
        
                RC.Top = RC.Top + WindowSkinCaptionHeight(0)
                RC.Left = RC.Left + WindowSkinBorderSize(0)
                RC.Right = RC.Right - WindowSkinBorderSize(1)
                RC.Bottom = RC.Bottom - WindowSkinBorderSize(2)
        
                CopyMemory ByVal lParam, RC, Len(RC)
        
                IBSSubclass_WindowProc = 0&
                Exit Function
            End If
        Case WM_NCHITTEST
            Dim X As Long, Y As Long
            X = (lParam And &HFFFF&)
            Y = ((lParam \ &H10000) And &HFFFF&)
            If X And &H8000& Then X = X Or &HFFFF0000
            If Y And &H8000& Then Y = Y Or &HFFFF0000
        
            GetWindowRect hWnd, RC
        
            Dim hit As Long: hit = HTCLIENT
        
            If X >= RC.Left And X < RC.Left + WindowSkinBorderSize(0) Then
                hit = HTLEFT
            ElseIf X < RC.Right And X >= RC.Right - WindowSkinBorderSize(1) Then
                hit = HTRIGHT
            End If
        
            If Y >= RC.Top And Y < RC.Top + WindowSkinBorderSize(2) Then
                If hit = HTLEFT Then
                    hit = HTTOPLEFT
                ElseIf hit = HTRIGHT Then
                    hit = HTTOPRIGHT
                Else
                    hit = HTTOP
                End If
            ElseIf Y < RC.Bottom And Y >= RC.Bottom - WindowSkinBorderSize(2) Then
                If hit = HTLEFT Then
                    hit = HTBOTTOMLEFT
                ElseIf hit = HTRIGHT Then
                    hit = HTBOTTOMRIGHT
                Else
                    hit = HTBOTTOM
                End If
            End If
        
            If Y >= RC.Top + WindowSkinBorderSize(2) And Y < RC.Top + WindowSkinCaptionHeight(0) Then
                If hit = HTCLIENT Then hit = HTCAPTION
            End If
        
            IBSSubclass_WindowProc = hit
            Exit Function
        Case WM_NCACTIVATE
            Select Case wParam
                Case WA_ACTIVE
                    IsWindowActive = 1
                Case WA_INACTIVE
                    IsWindowActive = 0
            End Select
            SetFramePos
            SetSkinTextures
            SetSizableSkinTextures
            
            IBSSubclass_WindowProc = 1&
            Exit Function
        Case WM_SIZE
            OnResize
    End Select
    
defproc:
    IBSSubclass_WindowProc = CallOldWindowProc(hWnd, uMsg, wParam, lParam)
End Function

Private Sub OnResize()
    SetSizableSkinTextures
    SetWindowFrameRgn
    lblCaption.Width = pbTopMiddle.Width
    lblCaptionShadow.Width = pbTopMiddle.Width
    lblResizeTop(0).Width = pbTopMiddle.Width
    SetFramePos
End Sub

Private Sub SetFramePos()
    If CurrentSkin <= System Or Terminated = True Or TargetForm.Visible = False Then Exit Sub
    Dim RC As RECT
    GetWindowRect TargetForm.hWnd, RC
    Dim InsertAfter&
    If IsWindowActive Then
        InsertAfter = -1
    Else
        InsertAfter = -2
    End If
    Dim Flags As Long: Flags = 0& 'SWP_FRAMECHANGED
    If IsChild() Then
        Flags = Flags Or SWP_NOZORDER
        Dim RCP As RECT
        Dim Parent As Long
        Parent = GetParent(TargetForm.hWnd)
        GetClientRect Parent, RCP
        Dim PT As POINTAPI
        PT.X = RCP.Left: PT.Y = RCP.Top
        ClientToScreen Parent, PT
        RC.Left = RC.Left - PT.X
        RC.Top = RC.Top - PT.Y
        RC.Right = RC.Right - PT.X
        RC.Bottom = RC.Bottom - PT.Y
    End If
    
    Dim dwp As Long
    dwp = BeginDeferWindowPos(8)
    If dwp = 0& Then Exit Sub
    dwp = DeferWindowPos(dwp, pbTopLeft.hWnd, InsertAfter, RC.Left, RC.Top, pbTopLeft.Width \ 15, pbTopLeft.Height \ 15, Flags)
    dwp = DeferWindowPos(dwp, pbTopMiddle.hWnd, InsertAfter, RC.Left + pbTopLeft.Width \ 15, RC.Top, pbTopMiddle.Width \ 15, pbTopMiddle.Height \ 15, Flags)
    dwp = DeferWindowPos(dwp, pbTopRight.hWnd, InsertAfter, RC.Right - pbTopRight.Width \ 15, RC.Top, pbTopRight.Width \ 15, pbTopRight.Height \ 15, Flags)
    dwp = DeferWindowPos(dwp, pbLeft.hWnd, InsertAfter, RC.Left, RC.Top + pbTopLeft.Height \ 15, pbLeft.Width \ 15, pbLeft.Height \ 15, Flags)
    dwp = DeferWindowPos(dwp, pbBottomLeft.hWnd, InsertAfter, RC.Left, RC.Top + pbTopLeft.Height \ 15 + pbLeft.Height \ 15, pbBottomLeft.Width \ 15, pbBottomLeft.Height \ 15, Flags)
    dwp = DeferWindowPos(dwp, pbBottomMiddle.hWnd, InsertAfter, RC.Left + pbBottomLeft.Width \ 15, RC.Top + pbTopLeft.Height \ 15 + pbLeft.Height \ 15, pbBottomMiddle.Width \ 15, pbBottomMiddle.Height \ 15, Flags)
    dwp = DeferWindowPos(dwp, pbBottomRight.hWnd, InsertAfter, RC.Left + pbBottomLeft.Width \ 15 + pbBottomMiddle.Width \ 15, RC.Top + pbTopLeft.Height \ 15 + pbLeft.Height \ 15, pbBottomRight.Width \ 15, pbBottomRight.Height \ 15, Flags)
    dwp = DeferWindowPos(dwp, pbRight.hWnd, InsertAfter, RC.Right - pbRight.Width \ 15, RC.Top + pbTopRight.Height \ 15, pbRight.Width \ 15, pbRight.Height \ 15, Flags)
    EndDeferWindowPos dwp
    DeleteObject dwp
    
    If VisibilitySet = False Then
        pbTopLeft.Visible = True
        pbTopMiddle.Visible = True
        pbTopRight.Visible = True
        pbLeft.Visible = True
        pbRight.Visible = True
        pbBottomLeft.Visible = True
        pbBottomMiddle.Visible = True
        pbBottomRight.Visible = True
        VisibilitySet = True
    End If
End Sub

Private Function IsChild() As Boolean
    IsChild = (GetWindowLong(TargetForm.hWnd, GWL_STYLE) And WS_CHILD) <> 0
End Function

Private Sub SetSkinTextures()
    If CurrentSkin <= System Then Exit Sub
    
    Set pbTopLeft.Picture = WindowSkinTopLeft((CurrentSkin - 1) * 2 + IsWindowActive)
    pbTopLeft.Width = ScaleX(WindowSkinTopLeft((CurrentSkin - 1) * 2 + IsWindowActive).Width, vbHimetric, vbTwips)
    pbTopLeft.Height = ScaleY(WindowSkinTopLeft((CurrentSkin - 1) * 2 + IsWindowActive).Height, vbHimetric, vbTwips)
    
    pbTopMiddle.Height = ScaleY(WindowSkinTop((CurrentSkin - 1) * 2 + IsWindowActive).Height, vbHimetric, vbTwips)
    
    Set pbTopRight.Picture = WindowSkinTopRight((CurrentSkin - 1) * 2 + IsWindowActive)
    pbTopRight.Width = ScaleX(WindowSkinTopRight((CurrentSkin - 1) * 2 + IsWindowActive).Width, vbHimetric, vbTwips)
    pbTopRight.Height = ScaleY(WindowSkinTopRight((CurrentSkin - 1) * 2 + IsWindowActive).Height, vbHimetric, vbTwips)
    
    pbLeft.Width = ScaleX(WindowSkinLeft((CurrentSkin - 1) * 2 + IsWindowActive).Width, vbHimetric, vbTwips)
    pbRight.Width = ScaleX(WindowSkinRight((CurrentSkin - 1) * 2 + IsWindowActive).Width, vbHimetric, vbTwips)
    
    Set pbBottomLeft.Picture = WindowSkinBottomLeft((CurrentSkin - 1) * 2 + IsWindowActive)
    pbBottomLeft.Width = ScaleX(WindowSkinBottomLeft((CurrentSkin - 1) * 2 + IsWindowActive).Width, vbHimetric, vbTwips)
    pbBottomLeft.Height = ScaleY(WindowSkinBottomLeft((CurrentSkin - 1) * 2 + IsWindowActive).Height, vbHimetric, vbTwips)
    
    pbBottomMiddle.Height = ScaleY(WindowSkinBottom((CurrentSkin - 1) * 2 + IsWindowActive).Height, vbHimetric, vbTwips)
    
    Set pbBottomRight.Picture = WindowSkinBottomRight((CurrentSkin - 1) * 2 + IsWindowActive)
    pbBottomRight.Width = ScaleX(WindowSkinBottomRight((CurrentSkin - 1) * 2 + IsWindowActive).Width, vbHimetric, vbTwips)
    pbBottomRight.Height = ScaleY(WindowSkinBottomRight((CurrentSkin - 1) * 2 + IsWindowActive).Height, vbHimetric, vbTwips)
    
    Set imgCloseButton.Picture = WindowSkinClose((CurrentSkin - 1) * 4 + IsWindowActive)
    
    imgMaximizeButton.Enabled = IsMaximizeEnabled()
    SetMaxButtonTexture
    
    imgMinimizeButton.Enabled = IsMinimizeEnabled()
    If IsMinimizeEnabled() Then
        Set imgMinimizeButton.Picture = WindowSkinMinimize((CurrentSkin - 1) * 4 + IsWindowActive)
    Else
        Set imgMinimizeButton.Picture = WindowSkinMinimize((CurrentSkin - 1) * 4 + 0)
    End If
    
    lblCaptionShadow.Visible = IsWindowActive
    
    imgCloseButton.Top = pbTopRight.Height - imgCloseButton.Height - 3 * Screen.TwipsPerPixelY
    imgMaximizeButton.Top = pbTopRight.Height - imgMaximizeButton.Height - 3 * Screen.TwipsPerPixelY
    imgMinimizeButton.Top = pbTopRight.Height - imgMinimizeButton.Height - 3 * Screen.TwipsPerPixelY
    imgCloseButton.Left = pbTopRight.Width - WindowSkinBorderSize(0) * Screen.TwipsPerPixelX - imgCloseButton.Width - 2 * Screen.TwipsPerPixelX
    imgMaximizeButton.Left = imgCloseButton.Left - 2 * Screen.TwipsPerPixelX - imgMaximizeButton.Width
    imgMinimizeButton.Left = imgMaximizeButton.Left - 2 * Screen.TwipsPerPixelX - imgMinimizeButton.Width
End Sub

Private Sub SetMaxButtonTexture()
    If IsMaximizeEnabled() Then
        Set imgMaximizeButton.Picture = WindowSkinMaximize((CurrentSkin - 1) * 8 + IsWindowActive - (GetWindowState() = 2) * 4)
    Else
        Set imgMaximizeButton.Picture = WindowSkinMaximize((CurrentSkin - 1) * 8 + 0)
    End If
End Sub

Private Sub SetSizableSkinTextures()
    If GetWindowState() = 1 Or CurrentSkin <= System Then Exit Sub
    
    Dim RC As RECT
    GetWindowRect TargetForm.hWnd, RC
    
    Dim X&, Y&, Width&, Height&
    Dim Size&, hDC&
    
    Size = ScaleX(WindowSkinTop((CurrentSkin - 1) * 2 + IsWindowActive).Width, vbHimetric, vbPixels)
    hDC = GetDCFromPicture(WindowSkinTop((CurrentSkin - 1) * 2 + IsWindowActive))
    Height = pbTopMiddle.Height \ 15
    pbTopMiddle.Width = (RC.Right - RC.Left) * 15 - pbTopLeft.Width - pbTopRight.Width
    pbTopMiddle.Cls
    For X = 0 To pbTopMiddle.Width \ Screen.TwipsPerPixelX Step Size
        BitBlt pbTopMiddle.hDC, X, 0&, Size, Height, hDC, 0&, 0&, vbSrcCopy
    Next X
    DeleteDC hDC
    
    Size = ScaleY(WindowSkinLeft((CurrentSkin - 1) * 2 + IsWindowActive).Height, vbHimetric, vbPixels)
    hDC = GetDCFromPicture(WindowSkinLeft((CurrentSkin - 1) * 2 + IsWindowActive))
    Width = pbLeft.Width \ 15
    pbLeft.Height = (RC.Bottom - RC.Top) * 15 - pbTopLeft.Height - pbBottomLeft.Height
    pbLeft.Cls
    For Y = 0 To pbLeft.Height \ Screen.TwipsPerPixelY Step Size
        BitBlt pbLeft.hDC, 0&, Y, Width, Size, hDC, 0&, 0&, vbSrcCopy
    Next Y
    DeleteDC hDC
    
    Size = ScaleX(WindowSkinBottom((CurrentSkin - 1) * 2 + IsWindowActive).Width, vbHimetric, vbPixels)
    hDC = GetDCFromPicture(WindowSkinBottom((CurrentSkin - 1) * 2 + IsWindowActive))
    Height = pbBottomMiddle.Height \ 15
    pbBottomMiddle.Width = (RC.Right - RC.Left) * 15 - pbBottomLeft.Width - pbBottomRight.Width
    pbBottomMiddle.Cls
    For X = 0 To pbBottomMiddle.Width \ Screen.TwipsPerPixelX Step Size
        BitBlt pbBottomMiddle.hDC, X, 0&, Size, Height, hDC, 0&, 0&, vbSrcCopy
    Next X
    DeleteDC hDC
    
    Size = ScaleY(WindowSkinRight((CurrentSkin - 1) * 2 + IsWindowActive).Height, vbHimetric, vbPixels)
    hDC = GetDCFromPicture(WindowSkinRight((CurrentSkin - 1) * 2 + IsWindowActive))
    Width = pbRight.Width \ 15
    pbRight.Height = (RC.Bottom - RC.Top) * 15 - pbTopRight.Height - pbBottomRight.Height
    pbRight.Cls
    For Y = 0 To pbRight.Height \ Screen.TwipsPerPixelY Step Size
        BitBlt pbRight.hDC, 0&, Y, Width, Size, hDC, 0&, 0&, vbSrcCopy
    Next Y
    DeleteDC hDC
End Sub

Sub SetSkin(NewSkin As WindowSkin, Optional ForceRedraw As Boolean = False)
    If NewSkin >= System Then CurrentSkin = NewSkin
    
    pbTopLeft.Visible = False
    pbTopMiddle.Visible = False
    pbTopRight.Visible = False
    pbLeft.Visible = False
    pbRight.Visible = False
    pbBottomLeft.Visible = False
    pbBottomMiddle.Visible = False
    pbBottomRight.Visible = False
    
    If NewSkin > System And Terminated = False Then
        AttachMessage Me, TargetForm.hWnd, WM_NCPAINT
        AttachMessage Me, TargetForm.hWnd, WM_MOVE
        AttachMessage Me, TargetForm.hWnd, WM_NCCALCSIZE
        AttachMessage Me, TargetForm.hWnd, WM_NCHITTEST
        AttachMessage Me, TargetForm.hWnd, WM_NCACTIVATE
        AttachMessage Me, TargetForm.hWnd, WM_SIZE
        
        Dim Redo As Boolean
        Redo = False
        
setupskin:
        lblCaption.Width = pbTopMiddle.Width
        lblCaptionShadow.Width = pbTopMiddle.Width
        lblResizeTop(0).Width = pbTopMiddle.Width
        
        lblCaption.Alignment = WindowSkinCaptionAlign(CurrentSkin - 1)
        lblCaptionShadow.Alignment = WindowSkinCaptionAlign(CurrentSkin - 1)
        If TypeOf TargetForm Is Form Then
            If Not TargetForm.Icon Is Nothing Then
                Set imgControlMenu.Picture = TargetForm.Icon
            End If
        End If
        SetSkinTextures
        SetResizeCursors
        
        VisibilitySet = False
        
        Dim Parent As Long
        Parent = GetParent(TargetForm.hWnd)
        SetParent pbTopLeft.hWnd, Parent
        SetParent pbTopMiddle.hWnd, Parent
        SetParent pbTopRight.hWnd, Parent
        SetParent pbLeft.hWnd, Parent
        SetParent pbBottomLeft.hWnd, Parent
        SetParent pbBottomMiddle.hWnd, Parent
        SetParent pbBottomRight.hWnd, Parent
        SetParent pbRight.hWnd, Parent
        
        SetWindowLong pbTopLeft.hWnd, GWL_EXSTYLE, (GetWindowLong(pbTopLeft.hWnd, GWL_EXSTYLE) And (Not WS_EX_APPWINDOW)) Or WS_EX_TOOLWINDOW
        SetWindowLong pbTopMiddle.hWnd, GWL_EXSTYLE, (GetWindowLong(pbTopMiddle.hWnd, GWL_EXSTYLE) And (Not WS_EX_APPWINDOW)) Or WS_EX_TOOLWINDOW
        SetWindowLong pbTopRight.hWnd, GWL_EXSTYLE, (GetWindowLong(pbTopRight.hWnd, GWL_EXSTYLE) And (Not WS_EX_APPWINDOW)) Or WS_EX_TOOLWINDOW
        SetWindowLong pbLeft.hWnd, GWL_EXSTYLE, (GetWindowLong(pbLeft.hWnd, GWL_EXSTYLE) And (Not WS_EX_APPWINDOW)) Or WS_EX_TOOLWINDOW
        SetWindowLong pbBottomLeft.hWnd, GWL_EXSTYLE, (GetWindowLong(pbBottomLeft.hWnd, GWL_EXSTYLE) And (Not WS_EX_APPWINDOW)) Or WS_EX_TOOLWINDOW
        SetWindowLong pbBottomMiddle.hWnd, GWL_EXSTYLE, (GetWindowLong(pbBottomMiddle.hWnd, GWL_EXSTYLE) And (Not WS_EX_APPWINDOW)) Or WS_EX_TOOLWINDOW
        SetWindowLong pbBottomRight.hWnd, GWL_EXSTYLE, (GetWindowLong(pbBottomRight.hWnd, GWL_EXSTYLE) And (Not WS_EX_APPWINDOW)) Or WS_EX_TOOLWINDOW
        SetWindowLong pbRight.hWnd, GWL_EXSTYLE, (GetWindowLong(pbRight.hWnd, GWL_EXSTYLE) And (Not WS_EX_APPWINDOW)) Or WS_EX_TOOLWINDOW
        
        If (IsChild() Or ForceRedraw) Then OnResize
    Else
        IBSSubclass_UnsubclassIt
        
        SetParent pbTopLeft.hWnd, Me.hWnd
        SetParent pbTopMiddle.hWnd, Me.hWnd
        SetParent pbTopRight.hWnd, Me.hWnd
        SetParent pbLeft.hWnd, Me.hWnd
        SetParent pbBottomLeft.hWnd, Me.hWnd
        SetParent pbBottomMiddle.hWnd, Me.hWnd
        SetParent pbBottomRight.hWnd, Me.hWnd
        SetParent pbRight.hWnd, Me.hWnd
        
        If NewSkin < 0 Then
            Terminated = True
            Exit Sub
        End If
        
        SetWindowRgn TargetForm.hWnd, 0&, True
    End If
        
    timCaptionMonitor.Enabled = (NewSkin <> System) And TypeOf TargetForm Is Form
    
    'If NewSkin Then SetFramePos
    
'    SetWindowFrameRgn
'    SetWindowPos TargetForm.hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED
End Sub

Private Function IsResizable() As Boolean
    IsResizable = (GetWindowState() = 0 And ((GetWindowLong(TargetForm.hWnd, GWL_STYLE) And WS_SIZEBOX) <> 0))
End Function

Private Sub SetResizeCursors()
    Dim Active As Byte
    Active = -IsResizable
    lblResizeLeft.MousePointer = 9 * Active
    lblResizeTopLeft.MousePointer = 8 * Active
    lblResizeTop(0).MousePointer = 7 * Active
    lblResizeTop(1).MousePointer = 7 * Active
    lblResizeTop(2).MousePointer = 7 * Active
    lblResizeTopRight.MousePointer = 6 * Active
    lblResizeRight.MousePointer = 9 * Active
    pbLeft.MousePointer = 9 * Active
    pbRight.MousePointer = 9 * Active
    pbBottomLeft.MousePointer = 6 * Active
    pbBottomMiddle.MousePointer = 7 * Active
    pbBottomRight.MousePointer = 8 * Active
End Sub

Private Sub SetWindowFrameRgn()
    If CurrentSkin = System Then Exit Sub
    Dim RC As RECT, RCWin As RECT
    Dim Rgn&, Rgn1&, Rgn2&, Rgn3&, Rgn4&, Rgn5&, Rgn6&, Rgn7&, Rgn8&, Rgn9&
    Dim Width%, Height%
    
    Dim topLeft As POINTAPI
    GetWindowRect TargetForm.hWnd, RCWin
    If CompareRect(RCWin, LastRect) Then Exit Sub
    LastRect = RCWin
    GetClientRect TargetForm.hWnd, RC
    topLeft.X = RC.Left
    topLeft.Y = RC.Top
    ClientToScreen TargetForm.hWnd, topLeft
    topLeft.Y = topLeft.Y - RCWin.Top
    topLeft.X = topLeft.X - RCWin.Left
    Rgn = CreateRectRgn(topLeft.X + RC.Left, topLeft.Y + RC.Top, topLeft.X + RC.Right, topLeft.Y + RC.Bottom)
    SetWindowRgn TargetForm.hWnd, Rgn, True
    DeleteObject Rgn
    
    Select Case CurrentSkin
        Case Bluemetal, Greenmetal
            GetWindowRect pbTopLeft.hWnd, RC
            Width = RC.Right - RC.Left: Height = RC.Bottom - RC.Top
            Rgn = CreateRectRgn(0, 0, Width, Height)
            Rgn1 = CreateRectRgn(0, 0, 7, 1)
            Rgn2 = CreateRectRgn(0, 1, 5, 2)
            Rgn3 = CreateRectRgn(0, 2, 3, 3)
            Rgn4 = CreateRectRgn(0, 3, 2, 4)
            Rgn5 = CreateRectRgn(0, 4, 2, 5)
            Rgn6 = CreateRectRgn(0, 5, 1, 6)
            Rgn7 = CreateRectRgn(0, 6, 1, 7)
            CombineRgn Rgn, Rgn, Rgn1, RGN_DIFF
            CombineRgn Rgn, Rgn, Rgn2, RGN_DIFF
            CombineRgn Rgn, Rgn, Rgn3, RGN_DIFF
            CombineRgn Rgn, Rgn, Rgn4, RGN_DIFF
            CombineRgn Rgn, Rgn, Rgn5, RGN_DIFF
            CombineRgn Rgn, Rgn, Rgn6, RGN_DIFF
            CombineRgn Rgn, Rgn, Rgn7, RGN_DIFF
            DeleteObject Rgn1
            DeleteObject Rgn2
            DeleteObject Rgn3
            DeleteObject Rgn4
            DeleteObject Rgn5
            DeleteObject Rgn6
            DeleteObject Rgn7
            SetWindowRgn pbTopLeft.hWnd, Rgn, True
            DeleteObject Rgn
            
            GetWindowRect pbTopRight.hWnd, RC
            Width = RC.Right - RC.Left: Height = RC.Bottom - RC.Top
            Rgn = CreateRectRgn(0, 0, Width, Height)
            Rgn1 = CreateRectRgn(Width - 7, 0, Width, 1)
            Rgn2 = CreateRectRgn(Width - 5, 1, Width, 2)
            Rgn3 = CreateRectRgn(Width - 3, 2, Width, 3)
            Rgn4 = CreateRectRgn(Width - 2, 3, Width, 4)
            Rgn5 = CreateRectRgn(Width - 2, 4, Width, 5)
            Rgn6 = CreateRectRgn(Width - 1, 5, Width, 6)
            Rgn7 = CreateRectRgn(Width - 1, 6, Width, 7)
            CombineRgn Rgn, Rgn, Rgn1, RGN_DIFF
            CombineRgn Rgn, Rgn, Rgn2, RGN_DIFF
            CombineRgn Rgn, Rgn, Rgn3, RGN_DIFF
            CombineRgn Rgn, Rgn, Rgn4, RGN_DIFF
            CombineRgn Rgn, Rgn, Rgn5, RGN_DIFF
            CombineRgn Rgn, Rgn, Rgn6, RGN_DIFF
            CombineRgn Rgn, Rgn, Rgn7, RGN_DIFF
            DeleteObject Rgn1
            DeleteObject Rgn2
            DeleteObject Rgn3
            DeleteObject Rgn4
            DeleteObject Rgn5
            DeleteObject Rgn6
            DeleteObject Rgn7
            SetWindowRgn pbTopRight.hWnd, Rgn, True
            DeleteObject Rgn
            
            GetWindowRect pbBottomLeft.hWnd, RC
            Width = RC.Right - RC.Left: Height = RC.Bottom - RC.Top
            Rgn = CreateRectRgn(0, 0, Width, Height)
            Rgn1 = CreateRectRgn(0, Height - 2, 1, Height - 1)
            Rgn2 = CreateRectRgn(0, Height - 1, 2, Height)
            CombineRgn Rgn, Rgn, Rgn1, RGN_DIFF
            CombineRgn Rgn, Rgn, Rgn2, RGN_DIFF
            DeleteObject Rgn1
            DeleteObject Rgn2
            SetWindowRgn pbBottomLeft.hWnd, Rgn, True
            DeleteObject Rgn
            
            GetWindowRect pbBottomRight.hWnd, RC
            Width = RC.Right - RC.Left: Height = RC.Bottom - RC.Top
            Rgn = CreateRectRgn(0, 0, Width, Height)
            Rgn1 = CreateRectRgn(Width - 1, Height - 2, Width, Height - 1)
            Rgn2 = CreateRectRgn(Width - 2, Height - 1, Width, Height)
            CombineRgn Rgn, Rgn, Rgn1, RGN_DIFF
            CombineRgn Rgn, Rgn, Rgn2, RGN_DIFF
            DeleteObject Rgn1
            DeleteObject Rgn2
            SetWindowRgn pbBottomRight.hWnd, Rgn, True
            DeleteObject Rgn
    End Select
End Sub

Private Sub imgCloseButton_Click()
    On Error Resume Next
    Unload TargetForm
End Sub

Private Sub imgCloseButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    Set imgCloseButton.Picture = WindowSkinClose((CurrentSkin - 1) * 4 + 3)
    imgCloseButton.Tag = "down"
End Sub

Private Sub imgCloseButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imgCloseButton.Tag = "down" Then Exit Sub
    Set imgCloseButton.Picture = WindowSkinClose((CurrentSkin - 1) * 4 + 2)
    timCloseHover.Enabled = True
End Sub

Private Sub imgCloseButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    Set imgCloseButton.Picture = WindowSkinClose((CurrentSkin - 1) * 4 + IsWindowActive)
    imgCloseButton.Tag = ""
End Sub

Private Sub imgControlMenu_DblClick()
    On Error Resume Next
    Unload TargetForm
End Sub

Private Sub imgControlMenu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowControlMenu
End Sub

Private Sub ToggleMaximized()
    If Not IsMaximizeEnabled() Then Exit Sub
    If GetWindowState() = 2 Then ShowWindow TargetForm.hWnd, SW_SHOWNORMAL Else ShowWindow TargetForm.hWnd, SW_SHOWMAXIMIZED
    SetResizeCursors
    SetMaxButtonTexture
End Sub

Private Sub imgMaximizeButton_Click()
    ToggleMaximized
End Sub

Private Sub imgMaximizeButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    Set imgMaximizeButton.Picture = WindowSkinMaximize((CurrentSkin - 1) * 8 + 3 - (GetWindowState() = 2) * 4)
    imgMaximizeButton.Tag = "down"
End Sub

Private Sub imgMaximizeButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imgMaximizeButton.Tag = "down" Then Exit Sub
    Set imgMaximizeButton.Picture = WindowSkinMaximize((CurrentSkin - 1) * 8 + 2 - (GetWindowState() = 2) * 4)
    timMaximizeHover.Enabled = True
End Sub

Private Sub imgMaximizeButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    Set imgMaximizeButton.Picture = WindowSkinMaximize((CurrentSkin - 1) * 8 + IsWindowActive - (GetWindowState() = 2) * 4)
    imgMaximizeButton.Tag = ""
End Sub

Private Sub imgMinimizeButton_Click()
    ShowWindow TargetForm.hWnd, SW_SHOWMINIMIZED
End Sub

Private Sub imgMinimizeButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    Set imgMinimizeButton.Picture = WindowSkinMinimize((CurrentSkin - 1) * 4 + 3)
    imgMinimizeButton.Tag = "down"
End Sub

Private Sub imgMinimizeButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imgMaximizeButton.Tag = "down" Then Exit Sub
    Set imgMinimizeButton.Picture = WindowSkinMinimize((CurrentSkin - 1) * 4 + 2)
    timMinimizeHover.Enabled = True
End Sub

Private Sub imgMinimizeButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    Set imgMinimizeButton.Picture = WindowSkinMinimize((CurrentSkin - 1) * 4 + IsWindowActive)
    imgMinimizeButton.Tag = ""
End Sub

Private Sub lblCaption_DblClick()
    ToggleMaximized
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then ShowControlMenu
End Sub

Private Sub lblCaptionShadow_DblClick()
    ToggleMaximized
End Sub

Private Sub lblCaptionShadow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then ShowControlMenu
End Sub

Private Sub lblResizeLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not IsResizable Then Exit Sub
    If Button = 1 Then
        ReleaseCapture
        SendMessage TargetForm.hWnd, WM_NCLBUTTONDOWN, HTLEFT, 0&
    End If
End Sub

Private Sub lblResizeRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not IsResizable Then Exit Sub
    If Button = 1 Then
        ReleaseCapture
        SendMessage TargetForm.hWnd, WM_NCLBUTTONDOWN, HTRIGHT, 0&
    End If
End Sub

Private Sub lblResizeTop_DblClick(Index As Integer)
    If GetWindowState() = 2 Then ToggleMaximized
End Sub

Private Sub lblResizeTop_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not IsResizable Then Exit Sub
    If Button = 1 Then
        ReleaseCapture
        SendMessage TargetForm.hWnd, WM_NCLBUTTONDOWN, HTTOP, 0&
    End If
End Sub

Private Sub lblResizeTopLeft_DblClick()
    On Error Resume Next
    If GetWindowState() = 2 Or IsMaximizeEnabled() = False Then Unload TargetForm
End Sub

Private Sub lblResizeTopLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If GetWindowState() = 2 Or IsMaximizeEnabled() = False Then ShowControlMenu
End Sub

Private Sub lblResizeTopLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not IsResizable Then Exit Sub
    If Button = 1 Then
        ReleaseCapture
        SendMessage TargetForm.hWnd, WM_NCLBUTTONDOWN, HTTOPLEFT, 0&
    End If
End Sub

Private Sub lblResizeTopRight_DblClick()
    If GetWindowState() = 2 Then ToggleMaximized
End Sub

Private Sub lblResizeTopRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not IsResizable Then Exit Sub
    If Button = 1 Then
        ReleaseCapture
        SendMessage TargetForm.hWnd, WM_NCLBUTTONDOWN, HTTOPRIGHT, 0&
    End If
End Sub

Private Sub pbBottomLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not IsResizable Then Exit Sub
    If Button = 1 Then
        ReleaseCapture
        SendMessage TargetForm.hWnd, WM_NCLBUTTONDOWN, HTBOTTOMLEFT, 0&
    End If
End Sub

Private Sub pbBottomMiddle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not IsResizable Then Exit Sub
    If Button = 1 Then
        ReleaseCapture
        SendMessage TargetForm.hWnd, WM_NCLBUTTONDOWN, HTBOTTOM, 0&
    End If
End Sub

Private Sub DragWindow()
    ReleaseCapture
    SendMessage TargetForm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    SetFramePos
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then DragWindow
End Sub

Private Sub ShowControlMenu()
    Dim hCur As Long
    Dim PT As POINTAPI
    Dim hMenu As Long
    Dim Cmd As Long
    
    GetCursorPos PT
    hMenu = GetSystemMenu(TargetForm.hWnd, 0)
    Cmd = TrackPopupMenu(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD, PT.X + 1, PT.Y + 1, 0, TargetForm.hWnd, ByVal 0&)
    If Cmd <> 0 Then SendMessage TargetForm.hWnd, WM_SYSCOMMAND, Cmd, 0&
End Sub

Private Sub lblCaptionShadow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then DragWindow
End Sub

Private Sub pbBottomRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not IsResizable Then Exit Sub
    If Button = 1 Then
        ReleaseCapture
        SendMessage TargetForm.hWnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0&
    End If
End Sub

Private Sub pbLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not IsResizable Then Exit Sub
    If Button = 1 Then
        ReleaseCapture
        SendMessage TargetForm.hWnd, WM_NCLBUTTONDOWN, HTLEFT, 0&
    End If
End Sub

Private Sub pbRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not IsResizable Then Exit Sub
    If Button = 1 Then
        ReleaseCapture
        SendMessage TargetForm.hWnd, WM_NCLBUTTONDOWN, HTRIGHT, 0&
    End If
End Sub

Private Sub pbTopLeft_DblClick()
    On Error Resume Next
    Unload TargetForm
End Sub

Private Sub pbTopLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowControlMenu
End Sub

Private Sub pbTopLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pbTopMiddle_MouseMove Button, Shift, X, Y
End Sub

Private Sub pbTopLeft_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then ShowControlMenu
End Sub

Private Sub pbTopMiddle_DblClick()
    ToggleMaximized
End Sub

Private Sub pbTopMiddle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then DragWindow
End Sub

Private Sub pbTopMiddle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then ShowControlMenu
End Sub

Private Sub pbTopRight_DblClick()
    ToggleMaximized
End Sub

Private Sub pbTopRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pbTopMiddle_MouseMove Button, Shift, X, Y
End Sub

Private Sub pbTopRight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then ShowControlMenu
End Sub

Sub SetCaption(Caption As String)
    SetLabelText lblCaption, Caption
    SetLabelText lblCaptionShadow, Caption
End Sub

Private Function IsMouseOn(imgImage As Image) As Boolean
    Dim PT As POINTAPI
    Dim rectLeft As Single, rectTop As Single, rectRight As Single, rectBottom As Single
    
    GetCursorPos PT
    ScreenToClient pbTopRight.hWnd, PT
    PT.X = PT.X * Screen.TwipsPerPixelX
    PT.Y = PT.Y * Screen.TwipsPerPixelY
    
    rectLeft = imgImage.Left
    rectTop = imgImage.Top
    rectRight = imgImage.Left + imgImage.Width
    rectBottom = imgImage.Top + imgImage.Height
    
    IsMouseOn = (PT.X >= rectLeft And PT.X <= rectRight And PT.Y >= rectTop And PT.Y <= rectBottom)
End Function

Private Sub timCaptionMonitor_Timer()
    If lblCaption <> TargetForm.Caption Then
        SetLabelText lblCaption, TargetForm.Caption
        SetLabelText lblCaptionShadow, TargetForm.Caption
    End If
End Sub

Private Sub timCloseHover_Timer()
    If (Not IsMouseOn(imgCloseButton)) And imgCloseButton.Tag <> "down" Then
        Set imgCloseButton.Picture = WindowSkinClose((CurrentSkin - 1) * 4 + IsWindowActive)
        timCloseHover.Enabled = False
    End If
End Sub

Private Sub timMaximizeHover_Timer()
    If (Not IsMouseOn(imgMaximizeButton)) And imgMaximizeButton.Tag <> "down" Then
        Set imgMaximizeButton.Picture = WindowSkinMaximize((CurrentSkin - 1) * 8 + IsWindowActive - (GetWindowState() = 2) * 4)
        timMaximizeHover.Enabled = False
    End If
End Sub

Private Sub timMinimizeHover_Timer()
    If (Not IsMouseOn(imgMinimizeButton)) And imgMinimizeButton.Tag <> "down" Then
        Set imgMinimizeButton.Picture = WindowSkinMinimize((CurrentSkin - 1) * 4 + IsWindowActive)
        timMinimizeHover.Enabled = False
    End If
End Sub

Private Function GetWindowState() As VbAppWinStyle
    Dim wp As WINDOWPLACEMENT
    wp.Length = Len(wp)
    
    If GetWindowPlacement(TargetForm.hWnd, wp) Then
        Select Case wp.showCmd
            Case SW_SHOWMAXIMIZED
                GetWindowState = vbMaximized
            Case SW_SHOWMINIMIZED
                GetWindowState = vbMinimized
            Case Else
                GetWindowState = vbNormal
        End Select
    Else
        GetWindowState = vbNormal
    End If
End Function

Private Function IsMaximizeEnabled() As Boolean
    IsMaximizeEnabled = (GetWindowLong(TargetForm.hWnd, GWL_STYLE) And WS_MAXIMIZEBOX) <> 0
End Function

Private Function IsMinimizeEnabled() As Boolean
    IsMinimizeEnabled = (GetWindowLong(TargetForm.hWnd, GWL_STYLE) And WS_MINIMIZEBOX) <> 0
End Function

#End If
