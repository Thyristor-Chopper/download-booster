VERSION 5.00
Begin VB.Form pagGeneral 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "일반"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7050
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "pagGeneral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame fDownload 
      Caption         =   "다운로드 설정"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6780
      Begin VB.CheckBox chkOpenWhenComplete 
         Caption         =   "완료 후 파일 열기(&O)"
         Height          =   255
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   2520
      End
      Begin VB.CheckBox chkOpenDirWhenComplete 
         Caption         =   "완료 후 폴더 열기(&P)"
         Height          =   255
         Left            =   3720
         TabIndex        =   2
         Top             =   240
         Width           =   2655
      End
      Begin VB.CheckBox chkAlwaysResume 
         Caption         =   "항상 이어받기(&A)"
         Height          =   255
         Left            =   3720
         TabIndex        =   4
         Top             =   480
         Width           =   2295
      End
      Begin VB.CheckBox chkRememberURL 
         Caption         =   "파일 주소 기억(&M)"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   480
         Width           =   2415
      End
      Begin VB.CheckBox chkAutoRetry 
         Caption         =   "오류 시 자동 재시도(&U)"
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   720
         Width           =   2610
      End
      Begin VB.ComboBox cbWhenExist 
         Height          =   300
         Left            =   3210
         Style           =   2  '드롭다운 목록
         TabIndex        =   9
         Top             =   1230
         Width           =   2040
      End
      Begin VB.CheckBox chkAutoYtdl 
         Caption         =   "지원되는 링크에서 자동으로 youtube-dl 사용(&Y)"
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   960
         Width           =   4815
      End
      Begin VB.CheckBox chkUseServerModified 
         Caption         =   "서버의 수정한 날짜 반영(&D)"
         Height          =   255
         Left            =   3720
         TabIndex        =   6
         Top             =   720
         Width           =   2895
      End
      Begin prjDownloadBooster.SpinBox txtMaxThreadCount 
         Height          =   270
         Left            =   3210
         TabIndex        =   11
         Top             =   1560
         Width           =   720
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Min             =   2
         Max             =   655
         Value           =   25
         AllowOnlyNumbers=   -1  'True
         TextAlignment   =   1
         HideSelection   =   0   'False
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "중복 파일명 처리(&D):"
         Height          =   180
         Left            =   1080
         TabIndex        =   8
         Tag             =   "nocolorchange"
         Top             =   1275
         Width           =   1710
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "최대 스레드 개수(&X):"
         Height          =   180
         Left            =   1080
         TabIndex        =   10
         Top             =   1605
         Width           =   1710
      End
      Begin VB.Label Label15 
         BackStyle       =   0  '투명
         Caption         =   "개 (다시 시작 필요)"
         Height          =   255
         Left            =   3960
         TabIndex        =   12
         Top             =   1605
         Width           =   2055
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   120
         Picture         =   "pagGeneral.frx":000C
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame fInterface 
      Caption         =   "인터페이스"
      Height          =   2895
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   6780
      Begin VB.ComboBox cbLanguage 
         Height          =   300
         Left            =   2640
         Style           =   2  '드롭다운 목록
         TabIndex        =   21
         Top             =   1470
         Width           =   1455
      End
      Begin VB.CheckBox chkAlwaysOnTop 
         Caption         =   "항상 위에 표시(&W)"
         Height          =   255
         Left            =   840
         TabIndex        =   14
         Top             =   240
         Width           =   1920
      End
      Begin VB.CheckBox chkLazyElapsed 
         Caption         =   "첫 바이트 수신 후 경과 시간 계산(&C)"
         Height          =   255
         Left            =   840
         TabIndex        =   16
         Top             =   480
         Width           =   4095
      End
      Begin VB.CheckBox chkForceOldDialog 
         Caption         =   "윈도우 3.1 대화 상자 사용(&S)"
         Height          =   255
         Left            =   840
         TabIndex        =   17
         Top             =   720
         Width           =   4095
      End
      Begin VB.CheckBox chkAllowDuplicates 
         Caption         =   "일괄 처리 목록에 중복 항목 허용(&I)"
         Height          =   255
         Left            =   840
         TabIndex        =   18
         Top             =   960
         Width           =   4095
      End
      Begin VB.OptionButton optLinePerScroll 
         Caption         =   "한 줄씩(&N)"
         Height          =   255
         Left            =   2640
         TabIndex        =   27
         Top             =   2175
         Width           =   1395
      End
      Begin VB.OptionButton optScreenPerScroll 
         Caption         =   "한 화면씩(&R)"
         Height          =   255
         Left            =   4200
         TabIndex        =   28
         Top             =   2175
         Width           =   1500
      End
      Begin VB.ComboBox cbStartupPosition 
         Height          =   300
         Left            =   2640
         Style           =   2  '드롭다운 목록
         TabIndex        =   30
         Top             =   2490
         Width           =   2535
      End
      Begin VB.CheckBox chkShowProgressInTaskbar 
         Caption         =   "작업 표시줄에 다운로드 진행률 표시(&R)"
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   1200
         Width           =   5055
      End
      Begin VB.ComboBox cbFont 
         Height          =   300
         Left            =   2640
         TabIndex        =   24
         Top             =   1830
         Width           =   1455
      End
      Begin VB.CheckBox chkAeroWindow 
         Caption         =   "유리 창 효과 사용(&G)"
         Height          =   255
         Left            =   3720
         TabIndex        =   15
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "언어(&L):"
         Height          =   255
         Left            =   1080
         TabIndex        =   20
         Top             =   1515
         Width           =   975
      End
      Begin VB.Image Image3 
         Height          =   405
         Left            =   120
         Picture         =   "pagGeneral.frx":044E
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label16 
         BackStyle       =   0  '투명
         Caption         =   "(다시 시작 필요)"
         Height          =   255
         Left            =   4200
         TabIndex        =   22
         Top             =   1530
         Width           =   1575
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "스레드 스크롤(&H):"
         Height          =   180
         Left            =   1080
         TabIndex        =   26
         Top             =   2220
         Width           =   1470
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "시작 창 위치(&T):"
         Height          =   180
         Left            =   1080
         TabIndex        =   29
         Top             =   2550
         Width           =   1350
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "글꼴(&F):"
         Height          =   180
         Left            =   1080
         TabIndex        =   23
         Top             =   1875
         Width           =   675
      End
      Begin VB.Label lblFontPreview 
         Alignment       =   2  '가운데 맞춤
         BorderStyle     =   1  '단일 고정
         Height          =   270
         Left            =   4200
         TabIndex        =   25
         Top             =   1845
         Width           =   1575
      End
   End
End
Attribute VB_Name = "pagGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Loaded As Boolean
Dim FontChanged As Boolean

Implements IBSSubclass

Private Sub Form_Load()
    AttachMessage Me, fInterface.hWnd, WM_PRINTCLIENT
End Sub

Private Sub Form_Unload(Cancel As Integer)
    IBSSubclass_UnsubclassIt
End Sub

Private Function IBSSubclass_MsgResponse(ByVal hWnd As Long, ByVal uMsg As Long) As EMsgResponse
    IBSSubclass_MsgResponse = emrConsume
End Function

Private Sub IBSSubclass_UnsubclassIt()
    DetachMessage Me, fInterface.hWnd, WM_PRINTCLIENT
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
    'InitForm Me
    Loaded = False
    FontChanged = False

    AddItemToComboBox cbFont, "(" & t("기본값", "default") & ")"
    If t(1, 2) = 2 Then
        If FontExists("Tahoma") Then AddItemToComboBox cbFont, "Tahoma"
        If FontExists("Segoe UI") Then AddItemToComboBox cbFont, "Segoe UI"
    Else
        If FontExists("굴림") Then AddItemToComboBox cbFont, "굴림"
        If FontExists("돋움") Then AddItemToComboBox cbFont, "돋움"
        If FontExists("바탕") Then AddItemToComboBox cbFont, "바탕"
        If FontExists("궁서") Then AddItemToComboBox cbFont, "궁서"
        If FontExists("맑은 고딕") Then AddItemToComboBox cbFont, "맑은 고딕"
    End If
    
    AddItemToComboBox cbLanguage, t("자동", "Auto")
    AddItemToComboBox cbLanguage, "한국어"
    AddItemToComboBox cbLanguage, "English"
    
    AddItemToComboBox cbWhenExist, t("건너뛰기", "Skip")
    AddItemToComboBox cbWhenExist, t("덮어쓰기", "Overwrite")
    AddItemToComboBox cbWhenExist, t("자동 이름 변경", "Auto Rename")
    
    AddItemToComboBox cbStartupPosition, t("이전 위치 기억", "Remember last position")
    AddItemToComboBox cbStartupPosition, t("화면 가운데", "Center of the screen")
    AddItemToComboBox cbStartupPosition, t("자동", "Auto")
    
    chkAeroWindow.Enabled = IsDWMEnabled()
    chkAeroWindow.Visible = (WinVer >= 6!)
    
    chkAlwaysOnTop = -(MainFormOnTop)
    chkLazyElapsed = GetSetting("DownloadBooster", "Options", "LazyElapsed", 0)
    chkForceOldDialog = GetSetting("DownloadBooster", "Options", "ForceWin31Dialog", 0)
    chkRememberURL = GetSetting("DownloadBooster", "Options", "RememberURL", 1)
    chkAutoYtdl = GetSetting("DownloadBooster", "Options", "AutoDetectYtdlURL", 1)
    chkAllowDuplicates = GetSetting("DownloadBooster", "Options", "AllowDuplicatesInQueue", 0)
    txtMaxThreadCount = GetSetting("DownloadBooster", "Options", "MaxThreadCount", 25)
    optLinePerScroll = True
    optScreenPerScroll = (GetSetting("DownloadBooster", "Options", "ScrollOneScreen", 0) <> 0)
    chkOpenWhenComplete.Value = frmMain.chkOpenAfterComplete.Value
    chkOpenDirWhenComplete.Value = frmMain.chkOpenFolder.Value
    chkAlwaysResume.Value = frmMain.chkContinueDownload.Value
    chkAutoRetry.Value = frmMain.chkAutoRetry.Value
    Dim LangSet As String
    LangSet = GetSetting("DownloadBooster", "Options", "Language", "0")
    If LangSet = "0" Then
        cbLanguage.ListIndex = 0
    ElseIf LangSet = "1042" Then
        cbLanguage.ListIndex = 1
    Else
        cbLanguage.ListIndex = 2
    End If
    cbWhenExist.ListIndex = GetSetting("DownloadBooster", "Options", "WhenFileExists", 0)
    cbFont = Trim$(GetSetting("DownloadBooster", "Options", "Font", ""))
    If cbFont = "" Then cbFont = ("(" & t("기본값", "default") & ")")
    cbStartupPosition.ListIndex = GetSetting("DownloadBooster", "Options", "StartupPosition", 0)
    chkShowProgressInTaskbar.Value = GetSetting("DownloadBooster", "Options", "ShowProgressInTaskbar", 1)
    
    tr fDownload, "Download settings"
    tr fInterface, "Interface"
    tr chkOpenWhenComplete, "&Open file when complete"
    tr chkOpenDirWhenComplete, "O&pen folder when complete"
    tr chkAlwaysResume, "&Always resume"
    tr chkAutoRetry, "A&uto retry on error"
    tr Label3, "If filename alrea&dy exists:"
    tr Label14, "Ma&x. number of threads:"
    tr Label15, "(restart required)"
    tr chkForceOldDialog, "U&se Windows 3.1 dialogs"
    tr chkAutoYtdl, "Automatically use &youtube-dl for supported links"
    tr chkAlwaysOnTop, "Al&ways on top"
    tr chkAeroWindow, "Use Aero &glass window"
    tr chkLazyElapsed, "Elapsed time sin&ce first data receive"
    tr chkAllowDuplicates, "Allow dupl&icates in queue"
    tr Label13, "&Font:"
    tr Label16, Label15.Caption
    tr Label18, "T&hread scroll:"
    tr optLinePerScroll, "Per li&ne"
    tr optScreenPerScroll, "Pe&r screen"
    tr Label9, "S&tartup position:"
    tr chkShowProgressInTaskbar, "Show p&rogress in taskbar"
    
#If HIDEYTDL Then
    chkAutoYtdl.Visible = False
    fDownload.Height = fDownload.Height - chkAutoYtdl.Height
    'fDownload.Refresh
    fInterface.Top = fInterface.Top - chkAutoYtdl.Height
    'fInterface.Refresh
    Label3.Top = Label3.Top - chkAutoYtdl.Height
    cbWhenExist.Top = cbWhenExist.Top - chkAutoYtdl.Height
    Label14.Top = Label14.Top - chkAutoYtdl.Height
    txtMaxThreadCount.Top = txtMaxThreadCount.Top - chkAutoYtdl.Height
    Label15.Top = Label15.Top - chkAutoYtdl.Height
#End If
    
    If WinVer < 6.1! Then
        chkShowProgressInTaskbar.Visible = False
        Label1.Top = Label1.Top - chkShowProgressInTaskbar.Height
        cbLanguage.Top = cbLanguage.Top - chkShowProgressInTaskbar.Height
        Label16.Top = Label16.Top - chkShowProgressInTaskbar.Height
        Label18.Top = Label18.Top - chkShowProgressInTaskbar.Height
        optLinePerScroll.Top = optLinePerScroll.Top - chkShowProgressInTaskbar.Height
        optScreenPerScroll.Top = optScreenPerScroll.Top - chkShowProgressInTaskbar.Height
        Label9.Top = Label9.Top - chkShowProgressInTaskbar.Height
        cbStartupPosition.Top = cbStartupPosition.Top - chkShowProgressInTaskbar.Height
        fInterface.Height = fInterface.Height - chkShowProgressInTaskbar.Height
        Label13.Top = Label13.Top - chkShowProgressInTaskbar.Height
        cbFont.Top = cbFont.Top - chkShowProgressInTaskbar.Height
        lblFontPreview.Top = lblFontPreview.Top - chkShowProgressInTaskbar.Height
    End If
    
    Loaded = True
End Sub

Private Sub cbFont_Change()
    OnFontChange
End Sub

Private Sub cbFont_Click()
    OnFontChange
End Sub

Private Sub OnFontChange()
    On Error Resume Next
    
    Dim FontName$, FontSize%
    If FontExists(cbFont) Then
        FontName = cbFont
    ElseIf cbFont = t("(기본값)", "(default)") Then
        If t(1, 2) = 2 Then
            FontName = "Tahoma"
        Else
            FontName = DefaultFont
        End If
    End If
    If LenB(FontName) = 0 Then Exit Sub
    If LCase(FontName) = "tahoma" Or Left$(FontName, 7) = "Tahoma " Then FontSize = 8 Else FontSize = 9
    
    lblFontPreview.Font.Name = FontName
    lblFontPreview.Font.Size = FontSize
    lblFontPreview.Font.Bold = False
    lblFontPreview.Font.Italic = False
    lblFontPreview = FontSize & " pt. " & FontName
    
    If Loaded Then
        EnableApply Me
        FontChanged = True
    End If
End Sub
