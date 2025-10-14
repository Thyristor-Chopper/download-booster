VERSION 5.00
Begin VB.Form frmLiveBadukSkinProperties 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "스킨 설정"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5310
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLiveBadukSkinProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CheckBox chkEnableBorders 
      Caption         =   "쪽지 테두리 표시(&W)"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Frame fFrameBackground 
      Caption         =   "프레임 배경"
      Height          =   1335
      Left            =   120
      TabIndex        =   16
      Top             =   3720
      Width           =   5055
      Begin VB.OptionButton optBackgroundColor 
         Caption         =   "단색(&O):"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1695
      End
      Begin prjDownloadBooster.CommandButtonW cmdSelectFrameTexture 
         Height          =   330
         Left            =   2160
         TabIndex        =   21
         Top             =   900
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         Caption         =   "선택(&L)..."
      End
      Begin VB.OptionButton optFrameTexture 
         Caption         =   "텍스처(&U):"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optFrameTransparent 
         Caption         =   "반투명(&R)"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.Label lblFrameBackgroundColorSelect 
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   19
         Top             =   600
         Width           =   615
      End
      Begin VB.Shape pgFrameBackgroundColor 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   255
         Left            =   2160
         Shape           =   4  '둥근 사각형
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame fText 
      Caption         =   "글자"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.CheckBox chkEnableFontSize 
         Caption         =   "라벨 글자 크기(&S):"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Value           =   1  '확인
         Width           =   1935
      End
      Begin VB.CheckBox chkBold 
         Caption         =   "라벨 굵게(&B)"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Value           =   1  '확인
         Width           =   2655
      End
      Begin prjDownloadBooster.SpinBox sbFontSize 
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   1320
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         Min             =   6
         Max             =   14
         Value           =   10
         ThousandsSeparator=   0   'False
         AllowOnlyNumbers=   -1  'True
         TextAlignment   =   1
      End
      Begin VB.CheckBox chkShadowColor 
         Caption         =   "라벨 그림자 색(&H):"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Value           =   1  '확인
         Width           =   1935
      End
      Begin VB.CheckBox chkTextColor 
         Caption         =   "라벨 글자 색(&X):"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblSelectTextColor 
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblSelectShadow 
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblSelectContentColor 
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   960
         Width           =   615
      End
      Begin VB.Shape pgContentColor 
         BackColor       =   &H00000000&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   255
         Left            =   2160
         Shape           =   4  '둥근 사각형
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "현황 글자 색(&P):"
         Height          =   180
         Left            =   390
         TabIndex        =   5
         Top             =   1005
         Width           =   1350
      End
      Begin VB.Shape pgText 
         BackColor       =   &H00000000&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   255
         Left            =   2160
         Shape           =   4  '둥근 사각형
         Top             =   240
         Width           =   615
      End
      Begin VB.Shape pgShadow 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   255
         Left            =   2160
         Shape           =   4  '둥근 사각형
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame fFrameColor 
      Caption         =   "프레임 테두리"
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   5055
      Begin prjDownloadBooster.CommandButtonW cmdSelectTexture 
         Height          =   330
         Left            =   2160
         TabIndex        =   15
         Top             =   930
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         Caption         =   "선택(&S)..."
      End
      Begin VB.OptionButton optTexture 
         Caption         =   "텍스처(&E):"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optColor 
         Caption         =   "단색(&C):"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optTransparent 
         Caption         =   "반투명(&T)"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.Label lblSelectFrameColor 
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   600
         Width           =   615
      End
      Begin VB.Shape pgFrameColor 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   255
         Left            =   2160
         Shape           =   4  '둥근 사각형
         Top             =   600
         Width           =   615
      End
   End
   Begin prjDownloadBooster.CommandButtonW cmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   3000
      TabIndex        =   24
      Top             =   5520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      Caption         =   "취소"
   End
   Begin prjDownloadBooster.CommandButtonW cmdOK 
      Default         =   -1  'True
      Height          =   345
      Left            =   960
      TabIndex        =   23
      Top             =   5520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      Caption         =   "확인"
   End
End
Attribute VB_Name = "frmLiveBadukSkinProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If DISABLEFRAMESKIN Then
#Else
Public SkinnedFrame As frmSkinnedFrame
#End If

Implements IBSSubclass

Private Function IBSSubclass_MsgResponse(ByVal hWnd As Long, ByVal uMsg As Long) As EMsgResponse
    IBSSubclass_MsgResponse = emrConsume
End Function

Private Sub IBSSubclass_UnsubclassIt()
    DetachMessage Me, fFrameColor.hWnd, WM_PRINTCLIENT
    DetachMessage Me, fFrameBackground.hWnd, WM_PRINTCLIENT
End Sub

Private Function IBSSubclass_WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
    On Error Resume Next
 
    Select Case uMsg
        Case WM_PRINTCLIENT
            SendMessage hWnd, WM_PAINT, wParam, ByVal 0&
            
            IBSSubclass_WindowProc = 0&
            Exit Function
    End Select
    
    IBSSubclass_WindowProc = CallOldWindowProc(hWnd, uMsg, wParam, lParam)
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    frmOptions.LiveBadukMemoSkinShadowColor = CLng(pgShadow.BackColor)
    frmOptions.LiveBadukMemoSkinFrameColor = CLng(pgFrameColor.BackColor)
    frmOptions.LiveBadukMemoSkinFrameType = IIf(optTransparent.Value, "transparent", IIf(optColor.Value, "solidcolor", "texture"))
    frmOptions.LiveBadukMemoSkinTextColor = CLng(pgText.BackColor)
    frmOptions.LiveBadukMemoSkinEnableShadow = chkShadowColor.Value
    frmOptions.LiveBadukMemoSkinEnableTextColor = chkTextColor.Value
    frmOptions.LiveBadukMemoSkinEnableBorder = chkEnableBorders.Value
    frmOptions.LiveBadukMemoSkinFrameBackgroundType = IIf(optFrameTransparent.Value, "transparent", IIf(optBackgroundColor.Value, "solidcolor", "texture"))
    frmOptions.LiveBadukMemoSkinFrameBackgroundColor = CLng(pgFrameBackgroundColor.BackColor)
    frmOptions.LiveBadukMemoSkinContentTextColor = CLng(pgContentColor.BackColor)
    frmOptions.LiveBadukMemoSkinEnableLabelFontSize = chkEnableFontSize.Value
    frmOptions.LiveBadukMemoSkinLabelFontSize = sbFontSize.Value
    frmOptions.LiveBadukMemoSkinLabelFontBold = chkBold.Value
    
    frmOptions.ColorChanged = True
    frmOptions.FontChanged = True
    frmOptions.ProgressSkinChanged = True
    frmOptions.cmdApply.Enabled = -1
    Unload Me
End Sub

Private Sub cmdSelectFrameTexture_Click()
    ShowFileDialog 6, frmOptions.LiveBadukMemoSkinFrameBackground
End Sub

Private Sub cmdSelectTexture_Click()
    ShowFileDialog 5, frmOptions.LiveBadukMemoSkinFrameTexture
End Sub

Private Sub Form_Load()
    InitForm Me
    
    pgShadow.BackColor = frmOptions.LiveBadukMemoSkinShadowColor
    pgFrameColor.BackColor = frmOptions.LiveBadukMemoSkinFrameColor
    pgText.BackColor = frmOptions.LiveBadukMemoSkinTextColor
    pgFrameBackgroundColor.BackColor = frmOptions.LiveBadukMemoSkinFrameBackgroundColor
    pgContentColor.BackColor = frmOptions.LiveBadukMemoSkinContentTextColor
    chkShadowColor.Value = frmOptions.LiveBadukMemoSkinEnableShadow
    chkTextColor.Value = frmOptions.LiveBadukMemoSkinEnableTextColor
    Select Case frmOptions.LiveBadukMemoSkinFrameType
        Case "solidcolor"
            optColor.Value = True
        Case "texture"
            optTexture.Value = True
        Case Else
            optTransparent.Value = True
    End Select
    Select Case frmOptions.LiveBadukMemoSkinFrameBackgroundType
        Case "texture"
            optFrameTexture.Value = True
        Case "solidcolor"
            optBackgroundColor.Value = True
        Case Else
            optFrameTransparent.Value = True
    End Select
    chkEnableBorders.Value = frmOptions.LiveBadukMemoSkinEnableBorder
    
    chkEnableFontSize.Value = frmOptions.LiveBadukMemoSkinEnableLabelFontSize
    sbFontSize.Value = frmOptions.LiveBadukMemoSkinLabelFontSize
    chkBold.Value = frmOptions.LiveBadukMemoSkinLabelFontBold
    
    chkEnableBorders.Enabled = (DPI = 96)
    
    Set cmdOK.ImageList = frmDummyForm.imgOK
    Set cmdCancel.ImageList = frmDummyForm.imgCancel
    
    tr fText, "Text"
    tr chkTextColor, "Label te&xt color:"
    tr chkShadowColor, "Label s&hadow color:"
    tr fFrameColor, "Frame"
    tr optTransparent, "Semi-&transparent"
    tr optColor, "&Color:"
    tr optTexture, "T&exture:"
    tr cmdSelectTexture, "&Select..."
    tr cmdOK, "OK"
    tr cmdCancel, "Cancel"
    tr Me, "Skin Settings"
    tr fFrameBackground, "Frame Background"
    tr optFrameTransparent, "Semi-t&ransparent"
    tr optFrameTexture, "Text&ure:"
    tr chkEnableBorders, "Sho&w borders"
    tr cmdSelectFrameTexture, "Se&lect..."
    tr optBackgroundColor, "C&olor:"
    tr Label1, "&Progress text color:"
    tr chkEnableFontSize, "Label text &size:"
    tr chkBold, "&Bold text"
    
    AttachMessage Me, fFrameColor.hWnd, WM_PRINTCLIENT
    AttachMessage Me, fFrameBackground.hWnd, WM_PRINTCLIENT
End Sub

Private Sub Form_Unload(Cancel As Integer)
    IBSSubclass_UnsubclassIt

#If DISABLEFRAMESKIN Then
#Else
    Unload SkinnedFrame
#End If
End Sub

Private Sub lblFrameBackgroundColorSelect_Click()
    Dim Color As OLE_COLOR
    Color = ShowColorDialog(Me.hWnd, True, pgFrameBackgroundColor.BackColor)
    If Color = -1 Then Exit Sub
    pgFrameBackgroundColor.BackColor = Color
    optBackgroundColor.Value = True
End Sub

Private Sub lblSelectContentColor_Click()
    Dim Color As OLE_COLOR
    Color = ShowColorDialog(Me.hWnd, True, pgContentColor.BackColor, True)
    If Color = -1 Then Exit Sub
    pgContentColor.BackColor = Color
End Sub

Private Sub lblSelectFrameColor_Click()
    Dim Color As OLE_COLOR
    Color = ShowColorDialog(Me.hWnd, True, pgFrameColor.BackColor)
    If Color = -1 Then Exit Sub
    pgFrameColor.BackColor = Color
    optColor.Value = True
End Sub

Private Sub lblSelectShadow_Click()
    Dim Color As OLE_COLOR
    Color = ShowColorDialog(Me.hWnd, True, pgShadow.BackColor, True)
    If Color = -1 Then Exit Sub
    pgShadow.BackColor = Color
    chkShadowColor.Value = 1
End Sub

Private Sub lblSelectTextColor_Click()
    Dim Color As OLE_COLOR
    Color = ShowColorDialog(Me.hWnd, True, pgText.BackColor, True)
    If Color = -1 Then Exit Sub
    pgText.BackColor = Color
    chkTextColor.Value = 1
End Sub

Private Sub sbFontSize_Change()
    chkEnableFontSize.Value = 1
End Sub

