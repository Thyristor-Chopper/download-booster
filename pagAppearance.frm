VERSION 5.00
Begin VB.Form pagAppearance 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "화면 배색"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7035
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "pagAppearance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.PictureBox pbPreview 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000001&
      Enabled         =   0   'False
      Height          =   2085
      Left            =   120
      ScaleHeight     =   2025
      ScaleWidth      =   6735
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "nobgdraw"
      Top             =   120
      Width           =   6795
      Begin VB.PictureBox pbBackground 
         Height          =   1320
         Left            =   240
         ScaleHeight     =   1260
         ScaleWidth      =   3975
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "nobgdraw"
         Top             =   120
         Width           =   4035
         Begin VB.Frame FrameW5 
            Caption         =   "다운로드 현황"
            Height          =   555
            Left            =   180
            TabIndex        =   4
            Top             =   375
            Width           =   3615
            Begin prjDownloadBooster.ProgressBar pbSampleClassic 
               Height          =   225
               Left            =   120
               Tag             =   "novisualstylechange"
               Top             =   240
               Visible         =   0   'False
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   397
               VisualStyles    =   0   'False
               Enabled         =   0   'False
               Value           =   5
               Step            =   10
            End
            Begin prjDownloadBooster.ProgressBar pbSample 
               Height          =   225
               Left            =   120
               Tag             =   "novisualstylechange"
               Top             =   240
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   397
               Enabled         =   0   'False
               Value           =   4
               Step            =   10
               State           =   3
            End
         End
         Begin VB.TextBox TextBoxW1 
            Height          =   270
            Left            =   1140
            TabIndex        =   3
            Top             =   60
            Width           =   2415
         End
         Begin VB.CheckBox CheckBoxW1 
            Caption         =   "완료 후 열기"
            Height          =   255
            Left            =   180
            TabIndex        =   5
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox txtSampleClassic 
            Height          =   270
            Left            =   1140
            TabIndex        =   41
            Top             =   60
            Visible         =   0   'False
            Width           =   2415
         End
         Begin prjDownloadBooster.CommandButtonW cmdSample 
            Height          =   285
            Left            =   2220
            TabIndex        =   6
            TabStop         =   0   'False
            Tag             =   "notygchange"
            Top             =   960
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            Caption         =   "다운로드"
         End
         Begin VB.Label Label11 
            BackStyle       =   0  '투명
            Caption         =   "파일 주소:"
            Height          =   255
            Left            =   180
            TabIndex        =   2
            Top             =   120
            Width           =   975
         End
         Begin VB.Shape pgPatternPreview 
            BackColor       =   &H8000000F&
            BackStyle       =   1  '투명하지 않음
            BorderStyle     =   0  '투명
            Height          =   255
            Left            =   0
            Top             =   0
            Width           =   255
         End
         Begin VB.Image imgPreview 
            Height          =   135
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   135
         End
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "배경"
      Height          =   1200
      Left            =   3570
      TabIndex        =   21
      Top             =   2580
      Width           =   3345
      Begin VB.ComboBox lvPatterns 
         Height          =   300
         Left            =   945
         Style           =   2  '드롭다운 목록
         TabIndex        =   23
         Top             =   150
         Width           =   1395
      End
      Begin VB.ComboBox lvBackgrounds 
         Height          =   300
         Left            =   945
         Style           =   2  '드롭다운 목록
         TabIndex        =   26
         Top             =   495
         Width           =   1395
      End
      Begin VB.ComboBox cbImagePosition 
         Height          =   300
         Left            =   945
         Style           =   2  '드롭다운 목록
         TabIndex        =   29
         Top             =   840
         Width           =   1395
      End
      Begin VB.CheckBox chkCenter 
         Caption         =   "중앙(&C)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2370
         TabIndex        =   30
         Top             =   870
         Width           =   945
      End
      Begin prjDownloadBooster.CommandButtonW cmdChooseBackground 
         Height          =   300
         Left            =   2370
         TabIndex        =   27
         Top             =   495
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         ImageListAlignment=   4
      End
      Begin VB.Label lblFillColorSelect 
         BackStyle       =   0  '투명
         Height          =   255
         Left            =   2370
         TabIndex        =   24
         Top             =   180
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "위치(&S):"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   885
         Width           =   840
      End
      Begin VB.Label Label21 
         BackStyle       =   0  '투명
         Caption         =   "무늬(&P):"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label22 
         BackStyle       =   0  '투명
         Caption         =   "배경(&A):"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   555
         Width           =   855
      End
      Begin VB.Shape pgPatternColor 
         BackColor       =   &H00000000&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   255
         Left            =   2370
         Shape           =   4  '둥근 사각형
         Top             =   180
         Width           =   615
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "스킨"
      Height          =   1200
      Left            =   3570
      TabIndex        =   31
      Top             =   3870
      Width           =   3345
      Begin VB.ComboBox cbFrameSkin 
         Height          =   300
         Left            =   870
         Style           =   2  '드롭다운 목록
         TabIndex        =   33
         Top             =   150
         Width           =   1785
      End
      Begin VB.ComboBox cbSkin 
         Height          =   300
         Left            =   870
         Style           =   2  '드롭다운 목록
         TabIndex        =   36
         Top             =   495
         Width           =   1785
      End
      Begin VB.ComboBox cbProgressSkin 
         Height          =   300
         Left            =   870
         Style           =   2  '드롭다운 목록
         TabIndex        =   39
         Top             =   840
         Width           =   1785
      End
      Begin prjDownloadBooster.CommandButtonW cmdAdvancedSkin 
         Height          =   300
         Left            =   2700
         TabIndex        =   37
         Top             =   495
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         ImageListAlignment=   4
      End
      Begin prjDownloadBooster.CommandButtonW cmdAdvancedProgressSkin 
         Height          =   300
         Left            =   2700
         TabIndex        =   40
         Top             =   840
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         ImageListAlignment=   4
      End
      Begin prjDownloadBooster.CommandButtonW cmdAdvancedFrameSkin 
         Height          =   300
         Left            =   2700
         TabIndex        =   34
         Top             =   150
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         ImageListAlignment=   4
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "단추(&O):"
         Height          =   180
         Left            =   120
         TabIndex        =   35
         Top             =   540
         Width           =   705
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "창(&W):"
         Height          =   180
         Left            =   120
         TabIndex        =   32
         Top             =   210
         Width           =   540
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '투명
         Caption         =   "현황(&R):"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   885
         Width           =   855
      End
   End
   Begin VB.ComboBox cbTheme 
      Height          =   300
      Left            =   960
      Style           =   2  '드롭다운 목록
      TabIndex        =   8
      Top             =   2235
      Width           =   3075
   End
   Begin prjDownloadBooster.CommandButtonW cmdDeleteTheme 
      Height          =   300
      Left            =   5580
      TabIndex        =   10
      Top             =   2235
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      Caption         =   "삭제(&D)"
   End
   Begin prjDownloadBooster.CommandButtonW cmdSaveTheme 
      Height          =   300
      Left            =   4155
      TabIndex        =   9
      Top             =   2235
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      Caption         =   "저장(&V)..."
   End
   Begin VB.Frame fForeColor 
      Caption         =   "글자색"
      Height          =   1200
      Left            =   120
      TabIndex        =   16
      Top             =   3870
      Width           =   3345
      Begin VB.CheckBox chkForeColorMainOnly 
         Caption         =   "메인 창에만 적용(&N)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   840
         Width           =   2805
      End
      Begin VB.OptionButton optUserFore 
         Caption         =   "사용자 지정(&U)"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   540
         Width           =   1575
      End
      Begin VB.OptionButton optSystemFore 
         Caption         =   "시스템 색상(&Y)"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblSelectFore 
         BackStyle       =   0  '투명
         Height          =   255
         Left            =   2415
         TabIndex        =   19
         Top             =   540
         Width           =   600
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000012&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   255
         Left            =   2415
         Shape           =   4  '둥근 사각형
         Top             =   240
         Width           =   600
      End
      Begin VB.Shape pgFore 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   255
         Left            =   2415
         Shape           =   4  '둥근 사각형
         Top             =   540
         Width           =   600
      End
   End
   Begin VB.Frame fBackColor 
      Caption         =   "배경색"
      Height          =   1200
      Left            =   120
      TabIndex        =   11
      Top             =   2580
      Width           =   3345
      Begin VB.CheckBox chkBackColorMainOnly 
         Caption         =   "메인 창에만 적용(&O)"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   840
         Width           =   2805
      End
      Begin VB.OptionButton optSystemColor 
         Caption         =   "시스템 색상(&S)"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optUserColor 
         Caption         =   "사용자 지정(&C)"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   540
         Width           =   1575
      End
      Begin VB.Label lblSelectColor 
         BackStyle       =   0  '투명
         Height          =   255
         Left            =   2415
         TabIndex        =   14
         Top             =   540
         Width           =   600
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000F&
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   255
         Left            =   2415
         Shape           =   4  '둥근 사각형
         Top             =   240
         Width           =   600
      End
      Begin VB.Shape pgColor 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         Height          =   255
         Left            =   2415
         Shape           =   4  '둥근 사각형
         Top             =   540
         Width           =   600
      End
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "테마(&T):"
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   690
   End
End
Attribute VB_Name = "pagAppearance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Loaded As Boolean
Public ColorChanged As Boolean
Public ImageChanged As Boolean
Public VisualStyleChanged As Boolean
Dim SkinChanged As Boolean
Public ProgressSkinChanged As Boolean
Dim PatternChanged As Boolean
Public ChangedBackgroundPath$
Dim DoLoadTheme As Boolean
Dim BackgroundParentDir As String
Dim PatternsSplit() As String
Dim PatternL As Byte, PatternU As Byte

Public RoundClassicButtons As Byte, DisableVisualStyle As Byte
Public LiveBadukMemoSkinShadowColor&, LiveBadukMemoSkinFrameColor&, LiveBadukMemoSkinFrameType$, LiveBadukMemoSkinTextColor&, LiveBadukMemoSkinEnableShadow As Byte, LiveBadukMemoSkinEnableTextColor As Byte, LiveBadukMemoSkinEnableBorder As Byte, LiveBadukMemoSkinFrameBackgroundType$, LiveBadukMemoSkinFrameBackgroundColor&, LiveBadukMemoSkinContentTextColor&, LiveBadukMemoSkinFrameTexture$, LiveBadukMemoSkinFrameBackground$, LiveBadukMemoSkinLabelFontSize As Integer, LiveBadukMemoSkinLabelFontBold As Byte, LiveBadukMemoSkinEnableLabelFontSize As Byte
Public ClassicFrame As Byte, NoDWMFrame As Byte

Implements IBSSubclass

Private Sub cbImagePosition_Click()
    chkCenter.Enabled = (cbImagePosition.ListIndex >= 1 And cbImagePosition.ListIndex <= 3 And lvBackgrounds.ListIndex <> 0)
    If Loaded Then
        EnableApply Me
        ImageChanged = True
    End If
End Sub

Private Sub Form_Load()
    AttachMessage Me, fBackColor.hWnd, WM_PRINTCLIENT
    AttachMessage Me, fForeColor.hWnd, WM_PRINTCLIENT
End Sub

Private Sub Form_Unload(Cancel As Integer)
    IBSSubclass_UnsubclassIt
End Sub

Private Function IBSSubclass_MsgResponse(ByVal hWnd As Long, ByVal uMsg As Long) As EMsgResponse
    IBSSubclass_MsgResponse = emrConsume
End Function

Private Sub IBSSubclass_UnsubclassIt()
    DetachMessage Me, fBackColor.hWnd, WM_PRINTCLIENT
    DetachMessage Me, fForeColor.hWnd, WM_PRINTCLIENT
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

Sub Initialize()
    InitForm Me, True
    Loaded = False
    
    PatternsSplit = Split("*.jpg;*.jpeg;*.jpe;*.jfif;*.gif;*.bmp;*.dib;*.png;*.wmf;*.emf;*.ico;*.cur;*.tif;*.tiff;*.rle", ";")
    PatternL = LBound(PatternsSplit)
    PatternU = UBound(PatternsSplit)
    
    pgPatternPreview.Width = pbBackground.Width
    pgPatternPreview.Height = pbBackground.Height
    
    pbBackground.Enabled = False
    SetPreviewPosition
    
    AddItemToComboBox cbFrameSkin, t("시스템 스타일", "System style")
#If DISABLEFRAMESKIN Then
#Else
    AddItemToComboBox cbFrameSkin, t("금속 파랑", "Blue metal")
    AddItemToComboBox cbFrameSkin, t("금속 초록", "Green metal")
    AddItemToComboBox cbFrameSkin, "Windows XP"
#End If
    
    AddItemToComboBox cbSkin, t("시스템 스타일", "System style")
    AddItemToComboBox cbSkin, t("라이브바둑 쪽지", "LiveBaduk memo")
    AddItemToComboBox cbSkin, t("디스크키퍼 2011", "Diskeeper 2011")
    AddItemToComboBox cbSkin, t("광택 파랑", "Glossy blue")
    AddItemToComboBox cbSkin, t("광택 초록", "Glossy green")
    
    AddItemToComboBox cbProgressSkin, t("(없음)", "(None)")
    AddItemToComboBox cbProgressSkin, t("기본 스킨", "Default skin")
    AddItemToComboBox cbProgressSkin, t("라이브바둑 쪽지", "LiveBaduk memo")
    
    AddItemToComboBox lvPatterns, t("(없음)", "(None)")
    AddItemToComboBox lvPatterns, t("수평선", "Horizontal lines")
    AddItemToComboBox lvPatterns, t("수직선", "Vertical lines")
    AddItemToComboBox lvPatterns, t("하향 대각선", "NW-SE lines")
    AddItemToComboBox lvPatterns, t("상향 대각선", "NE-SW lines")
    AddItemToComboBox lvPatterns, t("교차", "Grid")
    AddItemToComboBox lvPatterns, t("대각선 교차", "X grid")
    
    AddItemToComboBox cbImagePosition, t("늘이기", "Stretch")
    AddItemToComboBox cbImagePosition, t("높이 맞춤", "Fit to height")
    AddItemToComboBox cbImagePosition, t("너비 맞춤", "Fit to width")
    AddItemToComboBox cbImagePosition, t("원본 크기", "True size")
    AddItemToComboBox cbImagePosition, t("바둑판식", "Tile")
    
    DoLoadTheme = False
    ClearComboBox cbTheme
    AddItemToComboBox cbTheme, t("수정된 테마", "Modified theme")
    cbTheme.ListIndex = 0
    On Error Resume Next
    Dim ThemeList() As String
    If GetSubkeys(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\DownloadBooster\Options\Themes", ThemeList) Then
        Dim CurrentTheme$: CurrentTheme = GetSetting("DownloadBooster", "Options", "Theme", "")
        Dim k&
        For k = LBound(ThemeList) To UBound(ThemeList)
            AddItemToComboBox cbTheme, ThemeList(k)
            If ThemeList(k) = CurrentTheme Then cbTheme.ListIndex = cbTheme.ListCount - 1
        Next k
    End If
    DoLoadTheme = True
    
    LoadTheme
End Sub

Private Sub LoadTheme(Optional ByVal ThemeName As String = "")
    Dim Section$
    If ThemeName = "" Then Section = "Options" Else Section = "Options\Themes\" & ThemeName
    On Error Resume Next
    
    Dim clrBackColor As Long
    clrBackColor = GetSetting("DownloadBooster", Section, "BackColor", DefaultBackColor)
    If clrBackColor < 0 Or clrBackColor > 16777215 Then
        optSystemColor.Value = True
        pgColor.BackColor = &H8000000F
    Else
        pgColor.BackColor = clrBackColor
        optUserColor.Value = True
    End If
    FrameW5.BackColor = pgColor.BackColor
    CheckBoxW1.BackColor = pgColor.BackColor
    pbBackground.BackColor = pgColor.BackColor
    pgPatternPreview.BackColor = pgColor.BackColor
    
    cbSkin.ListIndex = GetSetting("DownloadBooster", Section, "ButtonSkin", 4)
    cbFrameSkin.ListIndex = GetSetting("DownloadBooster", Section, "WindowSkin", 0)
    cbProgressSkin.ListIndex = GetSetting("DownloadBooster", Section, "ProgressFrameSkin", 1)
    
    DisableVisualStyle = CByte(GetSetting("DownloadBooster", Section, "DisableVisualStyle", 0))
    RoundClassicButtons = CByte(GetSetting("DownloadBooster", Section, "RoundClassicButtons", 0))
    ClassicFrame = CByte(GetSetting("DownloadBooster", Section, "UseClassicThemeFrame", 0))
    NoDWMFrame = CByte(GetSetting("DownloadBooster", Section, "DisableDWMWindow", DefaultDisableDWMWindow))
    
    LiveBadukMemoSkinShadowColor = CLng(GetSetting("DownloadBooster", Section, "LiveBadukMemoSkinShadowColor", 16777215))
    LiveBadukMemoSkinFrameColor = CLng(GetSetting("DownloadBooster", Section, "LiveBadukMemoSkinFrameColor", 11194780))
    LiveBadukMemoSkinFrameType = LCase(GetSetting("DownloadBooster", Section, "LiveBadukMemoSkinFrameType", "solidcolor"))
    LiveBadukMemoSkinTextColor = CLng(GetSetting("DownloadBooster", Section, "LiveBadukMemoSkinTextColor", 2111255))
    LiveBadukMemoSkinEnableShadow = CByte(GetSetting("DownloadBooster", Section, "LiveBadukMemoSkinEnableShadow", 1))
    LiveBadukMemoSkinEnableTextColor = CByte(GetSetting("DownloadBooster", Section, "LiveBadukMemoSkinEnableTextColor", 1))
    LiveBadukMemoSkinEnableBorder = CByte(GetSetting("DownloadBooster", Section, "LiveBadukMemoSkinEnableBorder", 1))
    LiveBadukMemoSkinFrameBackgroundType = LCase(GetSetting("DownloadBooster", Section, "LiveBadukMemoSkinFrameBackgroundType", "transparent"))
    LiveBadukMemoSkinFrameBackgroundColor = CLng(GetSetting("DownloadBooster", Section, "LiveBadukMemoSkinFrameBackgroundColor", 16777215))
    LiveBadukMemoSkinContentTextColor = CLng(GetSetting("DownloadBooster", Section, "LiveBadukMemoSkinContentTextColor", 0))
    LiveBadukMemoSkinFrameTexture = GetSetting("DownloadBooster", Section, "LiveBadukMemoSkinFrameTexture", "")
    LiveBadukMemoSkinFrameBackground = GetSetting("DownloadBooster", Section, "LiveBadukMemoSkinFrameBackground", "")
    LiveBadukMemoSkinLabelFontSize = CInt(GetSetting("DownloadBooster", Section, "LiveBadukMemoSkinLabelFontSize", 10))
    LiveBadukMemoSkinLabelFontBold = CByte(GetSetting("DownloadBooster", Section, "LiveBadukMemoSkinLabelFontBold", 0))
    LiveBadukMemoSkinEnableLabelFontSize = CByte(GetSetting("DownloadBooster", Section, "LiveBadukMemoSkinEnableLabelFontSize", 0))
    
    If ClassicFrame Then RemoveVisualStyles pbBackground.hWnd
    
    cmdSample.RoundButton = RoundClassicButtons
    cmdSample.VisualStyles = (DisableVisualStyle = 0)
    cmdSample.IsTygemButton = cbSkin.ListIndex > 0
    txtSampleClassic.Visible = (DisableVisualStyle <> 0)
    pbSampleClassic.Visible = (DisableVisualStyle <> 0)
    
    lvPatterns.ListIndex = CInt(GetSetting("DownloadBooster", Section, "FormFillStyle", 4))
    
    pgPatternColor.BackColor = CLng(GetSetting("DownloadBooster", Section, "FormFillColor", 14544344))
    pgPatternPreview.FillColor = pgPatternColor.BackColor
    pgPatternPreview.FillStyle = lvPatterns.ListIndex + 1
    
    Dim clrForeColor As Long
    clrForeColor = GetSetting("DownloadBooster", Section, "ForeColor", -1)
    If clrForeColor < 0 Or clrForeColor > 16777215 Then
        optSystemFore.Value = True
        pgFore.BackColor = &H80000012
        If DisableVisualStyle = 0 Then
            ActivateVisualStyles CheckBoxW1.hWnd
            ActivateVisualStyles FrameW5.hWnd
        Else
            RemoveVisualStyles CheckBoxW1.hWnd
            RemoveVisualStyles FrameW5.hWnd
        End If
    Else
        optUserFore.Value = True
        pgFore.BackColor = clrForeColor
        RemoveVisualStyles CheckBoxW1.hWnd
        RemoveVisualStyles FrameW5.hWnd
        CheckBoxW1.ForeColor = pgFore.BackColor
        FrameW5.ForeColor = pgFore.BackColor
    End If
    Label11.ForeColor = pgFore.BackColor
    'CheckBoxW1.Transparent = True
    'FrameW5.Transparent = True
    
    ChangedBackgroundPath = GetSetting("DownloadBooster", Section, "BackgroundImagePath", "")
    LoadBackgroundList (ThemeName = ""), (GetSetting("DownloadBooster", Section, "UseBackgroundImage", 0) <> 0)
    Dim imgpos As Byte: imgpos = GetSetting("DownloadBooster", Section, "ImagePosition", 1)
    If imgpos > 3 And imgpos <= 6 Then
        imgpos = imgpos - 3: chkCenter.Value = 1
    ElseIf imgpos = 7 Then
        imgpos = 4
    End If
    cbImagePosition.ListIndex = imgpos
    cbImagePosition_Click
    
    chkBackColorMainOnly = GetSetting("DownloadBooster", Section, "BackColorMainOnly", 0)
    chkForeColorMainOnly = GetSetting("DownloadBooster", Section, "ForeColorMainOnly", 0)
End Sub

Sub LoadBackgroundList(Optional ByVal OnLoad As Boolean = False, Optional BackgroundImageEnabled As Boolean)
    Dim BackgroundPath$: BackgroundPath = ChangedBackgroundPath
    ClearComboBox lvBackgrounds
    AddItemToComboBox lvBackgrounds, t("(없음)", "(None)")
    Dim SelectedIndex&: SelectedIndex = 0
    BackgroundParentDir = GetParentFolderName(BackgroundPath)
    If Right$(BackgroundParentDir, 1) <> "\" Then BackgroundParentDir = BackgroundParentDir & "\"
    If FolderExists(BackgroundParentDir) Then
        Dim CurrentPattern$
        Dim CurrentBackgroundLcase$, FileNameLcase$
        If BackgroundImageEnabled Then
            CurrentBackgroundLcase = LCase(GetFilename(BackgroundPath))
            AddItemToComboBox lvBackgrounds, CurrentBackgroundLcase
            SelectedIndex = 1
        End If
        Dim FileName$, PatternMatched As Boolean
        Dim li&: li = 1
        Dim i As Byte
        Dim WFD As WIN32_FIND_DATA
        Dim hFind As Long
        hFind = FindFirstFile(BackgroundParentDir & "*.*", WFD)
        If hFind <> INVALID_HANDLE_VALUE Then
            Do
                If li > 50 Then Exit Do
                If WFD.dwFileAttributes And vbDirectory Then GoTo NextItem
                FileName = Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
                FileNameLcase = LCase(FileName)
                PatternMatched = False
                For i = PatternL To PatternU
                    CurrentPattern = PatternsSplit(i)
                    PatternMatched = (FileNameLcase Like CurrentPattern)
                    If PatternMatched Then Exit For
                Next i
                If Not PatternMatched Then GoTo NextItem
                If BackgroundImageEnabled = True And CurrentBackgroundLcase = FileNameLcase Then GoTo NextItem
                AddItemToComboBox lvBackgrounds, FileName
                li = li + 1
NextItem:
            Loop While FindNextFile(hFind, WFD)
            FindClose hFind
        End If
    End If
    lvBackgrounds.ListIndex = SelectedIndex
End Sub

Private Sub SetPreviewPosition()
    Dim Left%, Top%, Width%, Height%
    Left = 30
    Top = 6
    Width = 4035
    Height = 1320
    pbBackground.BorderStyle = 0
    SetWindowLong pbBackground.hWnd, GWL_STYLE, GetWindowLong(pbBackground.hWnd, GWL_STYLE) Or WS_BORDER Or WS_OVERLAPPED Or WS_CAPTION Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_SYSMENU
    SetWindowText pbBackground.hWnd, App.Title
    pbBackground.Top = Top * 15 + 15 + 30
    pbBackground.Left = Left * 15
    imgPreview.Width = Width
    imgPreview.Height = Height
    pbBackground.Width = Width + PaddedBorderWidth * 15 + DialogBorderWidth * 30
    pbBackground.Height = Height + PaddedBorderWidth * 15 + DialogBorderWidth * 30 + CaptionHeight * 15
    RedrawPreview
End Sub

Sub RedrawPreview()
    DoEvents
    pbBackground.Refresh
'    cmdSample.Refresh
'    Dim ctrl As Control
'    On Error Resume Next
'    For Each ctrl In Me.Controls
'        If ctrl.Container Is pbBackground Then
'            ctrl.Refresh
'            DoEvents
'        End If
'    Next ctrl
'    FrameW5.Refresh
'    CheckBoxW1.Refresh
End Sub

Private Sub cbSkin_Click()
    cmdSample.IsTygemButton = (cbSkin.ListIndex > 0)
    If cmdSample.IsTygemButton Then cmdSample.GetTygemButton().Skin = cbSkin.ListIndex
    'cmdSample.Refresh
    Dim ctrl As Control
    On Error Resume Next
    For Each ctrl In Me.Controls
        If ctrl.Container Is pbBackground And ctrl.Name <> "cmdSample" And ctrl.Name <> "pbSample" And ctrl.Name <> "pbSampleClassic" And ctrl.Name <> "txtSampleClassic" Then
            ctrl.VisualStyles = cmdSample.VisualStyles
        End If
    Next ctrl
    If Loaded Then
        EnableApply Me
        SkinChanged = True
        VisualStyleChanged = True
'        If cbSkin.ListIndex = 2 And DPI <> 96 Then
'            MsgBox t("이 스킨의 일부 요소는 96 DPI(100% 배율)에서만 표시됩니다.", "Some of the elements of this skin only works in 96 DPI (100% size)."), 48
'        End If
    End If
    If optUserFore.Value Then
        RemoveVisualStyles CheckBoxW1.hWnd
        RemoveVisualStyles FrameW5.hWnd
        CheckBoxW1.ForeColor = pgFore.BackColor
        FrameW5.ForeColor = pgFore.BackColor
    End If
    'cmdAdvancedSkin.Enabled = (cbSkin.ListIndex = 0)
    cmdSample.RoundButton = (GetSetting("DownloadBooster", "Options", "RoundClassicButtons", 0) <> 0)
End Sub

Private Sub lvBackgrounds_Click()
    On Error GoTo nopicture
    Dim BackgroundPath$
    BackgroundPath = BackgroundParentDir
    If Right$(BackgroundParentDir, 1) <> "\" Then BackgroundPath = BackgroundPath & "\"
    BackgroundPath = BackgroundPath & lvBackgrounds.Text
    If lvBackgrounds.ListIndex = 0 Then
nopicture:
        Set imgPreview.Picture = Nothing
    Else
        Set imgPreview.Picture = LoadPictureFromFile(BackgroundPath)
    End If
    'cmdSample.Refresh
    ChangedBackgroundPath = BackgroundPath
    If lvBackgrounds.ListIndex = 0 Then
        cbImagePosition.Enabled = False
        chkCenter.Enabled = False
        Label2.Enabled = False
    Else
        cbImagePosition.Enabled = True
        chkCenter.Enabled = (cbImagePosition.ListIndex >= 1 And cbImagePosition.ListIndex <= 3)
        Label2.Enabled = True
    End If
    If Loaded Then
        EnableApply Me
        ImageChanged = True
        RedrawPreview
    End If
End Sub
