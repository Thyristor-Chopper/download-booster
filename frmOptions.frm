VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "옵션"
   ClientHeight    =   11085
   ClientLeft      =   2760
   ClientTop       =   3855
   ClientWidth     =   14625
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11085
   ScaleWidth      =   14625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.PictureBox pbPanel 
      BorderStyle     =   0  '없음
      Enabled         =   0   'False
      Height          =   4425
      Index           =   4
      Left            =   7440
      ScaleHeight     =   4425
      ScaleWidth      =   7050
      TabIndex        =   4
      Top             =   5040
      Visible         =   0   'False
      Width           =   7050
      Begin VB.Frame FrameW6 
         Caption         =   "다운로드 알림"
         Height          =   975
         Left            =   120
         TabIndex        =   79
         Top             =   120
         Width           =   6810
         Begin prjDownloadBooster.CommandButtonW cmdBrowseCompleteSound 
            Height          =   300
            Left            =   5640
            TabIndex        =   84
            Top             =   570
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            ImageListAlignment=   4
         End
         Begin VB.TextBox txtCompleteSoundPath 
            Height          =   300
            Left            =   960
            TabIndex        =   82
            Top             =   570
            Width           =   3975
         End
         Begin VB.CheckBox chkBeepWhenComplete 
            Caption         =   "다운로드 완료(&B)"
            Height          =   255
            Left            =   720
            TabIndex        =   81
            Top             =   240
            Value           =   1  '확인
            Width           =   2535
         End
         Begin prjDownloadBooster.CommandButtonW cmdTestCompleteSound 
            Height          =   300
            Left            =   6240
            TabIndex        =   85
            Top             =   570
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   529
            Caption         =   "▶"
         End
         Begin prjDownloadBooster.CommandButtonW cmdResetCompleteSound 
            Height          =   300
            Left            =   5040
            TabIndex        =   83
            Top             =   570
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            ImageListAlignment=   4
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   120
            Picture         =   "frmOptions.frx":000C
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.Frame FrameW9 
         Caption         =   "메시지 상자"
         Height          =   3135
         Left            =   120
         TabIndex        =   80
         Top             =   1200
         Width           =   6810
         Begin prjDownloadBooster.CommandButtonW cmdBrowseAsterisk 
            Height          =   300
            Left            =   5640
            TabIndex        =   90
            Top             =   2730
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            ImageListAlignment=   4
         End
         Begin VB.TextBox txtAsterisk 
            Height          =   300
            Left            =   960
            TabIndex        =   88
            Top             =   2730
            Width           =   3975
         End
         Begin VB.CheckBox chkAsterisk 
            Caption         =   "일반 메시지(&A)"
            Height          =   255
            Left            =   720
            TabIndex        =   87
            Top             =   2400
            Value           =   1  '확인
            Width           =   1815
         End
         Begin prjDownloadBooster.CommandButtonW cmdTestAsterisk 
            Height          =   300
            Left            =   6240
            TabIndex        =   91
            Top             =   2730
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   529
            Caption         =   "▶"
         End
         Begin prjDownloadBooster.CommandButtonW cmdResetAsterisk 
            Height          =   300
            Left            =   5040
            TabIndex        =   89
            Top             =   2730
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            ImageListAlignment=   4
         End
         Begin prjDownloadBooster.CommandButtonW cmdBrowseExclamation 
            Height          =   300
            Left            =   5640
            TabIndex        =   96
            Top             =   2010
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            ImageListAlignment=   4
         End
         Begin VB.TextBox txtExclamation 
            Height          =   300
            Left            =   960
            TabIndex        =   94
            Top             =   2010
            Width           =   3975
         End
         Begin VB.CheckBox chkExclamation 
            Caption         =   "경고 메시지(&E)"
            Height          =   255
            Left            =   720
            TabIndex        =   93
            Top             =   1680
            Value           =   1  '확인
            Width           =   3135
         End
         Begin prjDownloadBooster.CommandButtonW cmdTestExclamation 
            Height          =   300
            Left            =   6240
            TabIndex        =   97
            Top             =   2010
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   529
            Caption         =   "▶"
         End
         Begin prjDownloadBooster.CommandButtonW cmdResetExclamation 
            Height          =   300
            Left            =   5040
            TabIndex        =   95
            Top             =   2010
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            ImageListAlignment=   4
         End
         Begin prjDownloadBooster.CommandButtonW cmdBrowseError 
            Height          =   300
            Left            =   5640
            TabIndex        =   102
            Top             =   1290
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            ImageListAlignment=   4
         End
         Begin VB.TextBox txtError 
            Height          =   300
            Left            =   960
            TabIndex        =   100
            Top             =   1290
            Width           =   3975
         End
         Begin VB.CheckBox chkError 
            Caption         =   "오류 메시지(&R)"
            Height          =   255
            Left            =   720
            TabIndex        =   99
            Top             =   960
            Value           =   1  '확인
            Width           =   1815
         End
         Begin prjDownloadBooster.CommandButtonW cmdTestError 
            Height          =   300
            Left            =   6240
            TabIndex        =   103
            Top             =   1290
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   529
            Caption         =   "▶"
         End
         Begin prjDownloadBooster.CommandButtonW cmdResetError 
            Height          =   300
            Left            =   5040
            TabIndex        =   101
            Top             =   1290
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            ImageListAlignment=   4
         End
         Begin prjDownloadBooster.CommandButtonW cmdBrowseQuestion 
            Height          =   300
            Left            =   5640
            TabIndex        =   108
            Top             =   570
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            ImageListAlignment=   4
         End
         Begin VB.TextBox txtQuestion 
            Height          =   300
            Left            =   960
            TabIndex        =   106
            Top             =   570
            Width           =   3975
         End
         Begin VB.CheckBox chkQuestion 
            Caption         =   "질문(&Q)"
            Height          =   255
            Left            =   720
            TabIndex        =   105
            Top             =   240
            Value           =   1  '확인
            Width           =   1815
         End
         Begin prjDownloadBooster.CommandButtonW cmdTestQuestion 
            Height          =   300
            Left            =   6240
            TabIndex        =   109
            Top             =   570
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   529
            Caption         =   "▶"
         End
         Begin prjDownloadBooster.CommandButtonW cmdResetQuestion 
            Height          =   300
            Left            =   5040
            TabIndex        =   107
            Top             =   570
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            ImageListAlignment=   4
         End
         Begin VB.Image Image8 
            Height          =   480
            Left            =   120
            Picture         =   "frmOptions.frx":0456
            Top             =   240
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox pbPanel 
      BorderStyle     =   0  '없음
      Enabled         =   0   'False
      Height          =   2385
      Index           =   2
      Left            =   7440
      ScaleHeight     =   2385
      ScaleWidth      =   7050
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   7050
      Begin VB.Frame FrameW3 
         Caption         =   "서버 연결"
         Height          =   1215
         Left            =   120
         TabIndex        =   37
         Top             =   120
         Width           =   6780
         Begin prjDownloadBooster.Slider trRequestInterval 
            Height          =   450
            Left            =   3000
            TabIndex        =   42
            Top             =   720
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   794
            Max             =   7
            Value           =   2
            ShowTip         =   0   'False
            Transparent     =   -1  'True
         End
         Begin VB.CheckBox chkIgnore300 
            Caption         =   "300번대 응답 코드 무시(&I)"
            Height          =   255
            Left            =   3720
            TabIndex        =   39
            Top             =   240
            Width           =   2415
         End
         Begin VB.CheckBox chkForceGet 
            Caption         =   "파일 검사 시 GET 요청(&Q)"
            Height          =   255
            Left            =   840
            TabIndex        =   40
            Top             =   480
            Width           =   2865
         End
         Begin VB.CheckBox chkNoRedirectCheck 
            Caption         =   "리다이렉트 검사 안 함(&R)"
            Height          =   255
            Left            =   840
            TabIndex        =   38
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label lblIntervalDisplay 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "(0.1초)"
            Height          =   180
            Left            =   5070
            TabIndex        =   43
            Top             =   840
            Width           =   570
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "스레드 요청 간격(&N):"
            Height          =   180
            Left            =   1080
            TabIndex        =   41
            Top             =   840
            Width           =   1725
         End
         Begin VB.Image Image4 
            Height          =   480
            Left            =   120
            Picture         =   "frmOptions.frx":0898
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Frame fHeaders 
         Caption         =   "사용자 지정 헤더"
         Height          =   795
         Left            =   120
         TabIndex        =   44
         Top             =   1440
         Width           =   6780
         Begin prjDownloadBooster.CommandButtonW cmdHeaders 
            Height          =   330
            Left            =   5040
            TabIndex        =   86
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            Caption         =   "헤더(&H)..."
         End
         Begin VB.Image Image5 
            Height          =   480
            Left            =   120
            Picture         =   "frmOptions.frx":0CDA
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label17 
            BackStyle       =   0  '투명
            Caption         =   "다운로드 중 서버에 요청할 때 전송할 헤더를 설정합니다."
            Height          =   495
            Left            =   840
            TabIndex        =   92
            Top             =   240
            Width           =   4095
         End
      End
   End
   Begin VB.PictureBox pbPanel 
      BorderStyle     =   0  '없음
      Height          =   5145
      Index           =   1
      Left            =   120
      ScaleHeight     =   5145
      ScaleWidth      =   7050
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   7050
      Begin VB.Frame Frame5 
         Caption         =   "인터페이스"
         Height          =   2895
         Left            =   120
         TabIndex        =   19
         Top             =   2160
         Width           =   6780
         Begin VB.CheckBox chkAeroWindow 
            Caption         =   "유리 창 효과 사용(&G)"
            Height          =   255
            Left            =   3720
            TabIndex        =   21
            Top             =   240
            Width           =   2535
         End
         Begin VB.ComboBox cbFont 
            Height          =   300
            Left            =   2640
            TabIndex        =   30
            Top             =   1830
            Width           =   1455
         End
         Begin VB.CheckBox chkShowProgressInTaskbar 
            Caption         =   "작업 표시줄에 다운로드 진행률 표시(&R)"
            Height          =   255
            Left            =   840
            TabIndex        =   25
            Top             =   1200
            Width           =   5055
         End
         Begin VB.ComboBox cbStartupPosition 
            Height          =   300
            Left            =   2640
            Style           =   2  '드롭다운 목록
            TabIndex        =   36
            Top             =   2490
            Width           =   2535
         End
         Begin VB.OptionButton optScreenPerScroll 
            Caption         =   "한 화면씩(&R)"
            Height          =   255
            Left            =   4200
            TabIndex        =   34
            Top             =   2175
            Width           =   1500
         End
         Begin VB.OptionButton optLinePerScroll 
            Caption         =   "한 줄씩(&N)"
            Height          =   255
            Left            =   2640
            TabIndex        =   33
            Top             =   2175
            Width           =   1395
         End
         Begin VB.CheckBox chkAllowDuplicates 
            Caption         =   "일괄 처리 목록에 중복 항목 허용(&I)"
            Height          =   255
            Left            =   840
            TabIndex        =   24
            Top             =   960
            Width           =   4095
         End
         Begin VB.CheckBox chkForceOldDialog 
            Caption         =   "윈도우 3.1 대화 상자 사용(&S)"
            Height          =   255
            Left            =   840
            TabIndex        =   23
            Top             =   720
            Width           =   4095
         End
         Begin VB.CheckBox chkLazyElapsed 
            Caption         =   "첫 바이트 수신 후 경과 시간 계산(&C)"
            Height          =   255
            Left            =   840
            TabIndex        =   22
            Top             =   480
            Width           =   4095
         End
         Begin VB.CheckBox chkAlwaysOnTop 
            Caption         =   "항상 위에 표시(&W)"
            Height          =   255
            Left            =   840
            TabIndex        =   20
            Top             =   240
            Width           =   1920
         End
         Begin VB.ComboBox cbLanguage 
            Height          =   300
            Left            =   2640
            Style           =   2  '드롭다운 목록
            TabIndex        =   27
            Top             =   1470
            Width           =   1455
         End
         Begin VB.Label lblFontPreview 
            Alignment       =   2  '가운데 맞춤
            BorderStyle     =   1  '단일 고정
            Height          =   270
            Left            =   4200
            TabIndex        =   31
            Top             =   1845
            Width           =   1575
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "글꼴(&F):"
            Height          =   180
            Left            =   1080
            TabIndex        =   29
            Top             =   1875
            Width           =   675
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "시작 창 위치(&T):"
            Height          =   180
            Left            =   1080
            TabIndex        =   35
            Top             =   2550
            Width           =   1350
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "스레드 스크롤(&H):"
            Height          =   180
            Left            =   1080
            TabIndex        =   32
            Top             =   2220
            Width           =   1470
         End
         Begin VB.Label Label16 
            BackStyle       =   0  '투명
            Caption         =   "(다시 시작 필요)"
            Height          =   255
            Left            =   4200
            TabIndex        =   28
            Top             =   1530
            Width           =   1575
         End
         Begin VB.Image Image3 
            Height          =   405
            Left            =   120
            Picture         =   "frmOptions.frx":111C
            Top             =   240
            Width           =   435
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "언어(&L):"
            Height          =   255
            Left            =   1080
            TabIndex        =   26
            Top             =   1515
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "다운로드 설정"
         Height          =   1935
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   6780
         Begin VB.CheckBox chkUseServerModified 
            Caption         =   "서버의 수정한 날짜 반영(&D)"
            Height          =   255
            Left            =   3720
            TabIndex        =   12
            Top             =   720
            Width           =   2895
         End
         Begin prjDownloadBooster.SpinBox txtMaxThreadCount 
            Height          =   270
            Left            =   3210
            TabIndex        =   17
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
         End
         Begin VB.CheckBox chkAutoYtdl 
            Caption         =   "지원되는 링크에서 자동으로 youtube-dl 사용(&Y)"
            Height          =   255
            Left            =   840
            TabIndex        =   13
            Top             =   960
            Width           =   4815
         End
         Begin VB.ComboBox cbWhenExist 
            Height          =   300
            Left            =   3210
            Style           =   2  '드롭다운 목록
            TabIndex        =   15
            Top             =   1230
            Width           =   2040
         End
         Begin VB.CheckBox chkAutoRetry 
            Caption         =   "오류 시 자동 재시도(&U)"
            Height          =   255
            Left            =   840
            TabIndex        =   11
            Top             =   720
            Width           =   2610
         End
         Begin VB.CheckBox chkRememberURL 
            Caption         =   "파일 주소 기억(&M)"
            Height          =   255
            Left            =   840
            TabIndex        =   9
            Top             =   480
            Width           =   2415
         End
         Begin VB.CheckBox chkAlwaysResume 
            Caption         =   "항상 이어받기(&A)"
            Height          =   255
            Left            =   3720
            TabIndex        =   10
            Top             =   480
            Width           =   2295
         End
         Begin VB.CheckBox chkOpenDirWhenComplete 
            Caption         =   "완료 후 폴더 열기(&P)"
            Height          =   255
            Left            =   3720
            TabIndex        =   8
            Top             =   240
            Width           =   2655
         End
         Begin VB.CheckBox chkOpenWhenComplete 
            Caption         =   "완료 후 파일 열기(&O)"
            Height          =   255
            Left            =   840
            TabIndex        =   7
            Top             =   240
            Width           =   2520
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   120
            Picture         =   "frmOptions.frx":134E
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label15 
            BackStyle       =   0  '투명
            Caption         =   "개 (다시 시작 필요)"
            Height          =   255
            Left            =   3960
            TabIndex        =   18
            Top             =   1605
            Width           =   2055
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "최대 스레드 개수(&X):"
            Height          =   180
            Left            =   1080
            TabIndex        =   16
            Top             =   1605
            Width           =   1710
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "중복 파일명 처리(&D):"
            Height          =   180
            Left            =   1080
            TabIndex        =   14
            Tag             =   "nocolorchange"
            Top             =   1275
            Width           =   1710
         End
      End
   End
   Begin VB.PictureBox pbPanel 
      BorderStyle     =   0  '없음
      ClipControls    =   0   'False
      Enabled         =   0   'False
      Height          =   1695
      Index           =   5
      Left            =   7440
      ScaleHeight     =   1695
      ScaleWidth      =   7050
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   7050
      Begin VB.Frame FrameW2 
         Caption         =   "프로그램 경로"
         ClipControls    =   0   'False
         Height          =   1455
         Left            =   120
         TabIndex        =   110
         Top             =   120
         Width           =   6780
         Begin VB.TextBox txtYtdlPath 
            Height          =   270
            Left            =   2520
            TabIndex        =   115
            Top             =   1080
            Width           =   4095
         End
         Begin VB.TextBox txtNodePath 
            Height          =   270
            Left            =   2520
            TabIndex        =   113
            Top             =   720
            Width           =   4095
         End
         Begin VB.Image imgIcon2 
            Height          =   480
            Left            =   120
            Picture         =   "frmOptions.frx":1790
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label6 
            BackStyle       =   0  '투명
            Caption         =   "기본값을 사용하려면 필드를 비워두십시오. 아래는 고급 사용자를 위한 것이며 일반적으로 변경할 필요가 없습니다."
            Height          =   480
            Left            =   840
            TabIndex        =   111
            Top             =   240
            Width           =   5775
         End
         Begin VB.Label Label7 
            BackStyle       =   0  '투명
            Caption         =   "&youtube-dl/yt-dlp:"
            Height          =   255
            Left            =   840
            TabIndex        =   114
            Top             =   1110
            Width           =   1695
         End
         Begin VB.Label Label4 
            BackStyle       =   0  '투명
            Caption         =   "N&ode.js:"
            Height          =   255
            Left            =   840
            TabIndex        =   112
            Top             =   750
            Width           =   1455
         End
      End
   End
   Begin VB.PictureBox pbPanel 
      BorderStyle     =   0  '없음
      Enabled         =   0   'False
      Height          =   5130
      Index           =   3
      Left            =   120
      ScaleHeight     =   5130
      ScaleWidth      =   7050
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   7050
      Begin VB.Frame Frame1 
         Caption         =   "배경색"
         Height          =   1200
         Left            =   120
         TabIndex        =   50
         Top             =   2580
         Width           =   3345
         Begin VB.OptionButton optUserColor 
            Caption         =   "사용자 지정(&C)"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   540
            Width           =   1575
         End
         Begin VB.OptionButton optSystemColor 
            Caption         =   "시스템 색상(&S)"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   1575
         End
         Begin VB.CheckBox chkBackColorMainOnly 
            Caption         =   "메인 창에만 적용(&O)"
            Height          =   255
            Left            =   360
            TabIndex        =   54
            Top             =   840
            Width           =   2805
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
         Begin VB.Label lblSelectColor 
            BackStyle       =   0  '투명
            Height          =   255
            Left            =   2415
            TabIndex        =   53
            Top             =   540
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
      Begin VB.Frame FrameW4 
         Caption         =   "글자색"
         Height          =   1200
         Left            =   120
         TabIndex        =   55
         Top             =   3870
         Width           =   3345
         Begin VB.OptionButton optSystemFore 
            Caption         =   "시스템 색상(&Y)"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optUserFore 
            Caption         =   "사용자 지정(&U)"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   540
            Width           =   1575
         End
         Begin VB.CheckBox chkForeColorMainOnly 
            Caption         =   "메인 창에만 적용(&N)"
            Enabled         =   0   'False
            Height          =   255
            Left            =   360
            TabIndex        =   59
            Top             =   840
            Width           =   2805
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
         Begin VB.Label lblSelectFore 
            BackStyle       =   0  '투명
            Height          =   255
            Left            =   2415
            TabIndex        =   58
            Top             =   540
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
      Begin VB.Frame FrameW1 
         BorderStyle     =   0  '없음
         Caption         =   "테마"
         Height          =   345
         Left            =   120
         TabIndex        =   45
         Top             =   2235
         Width           =   6795
         Begin prjDownloadBooster.CommandButtonW cmdDeleteTheme 
            Height          =   300
            Left            =   5460
            TabIndex        =   49
            Top             =   0
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            Caption         =   "삭제(&D)"
         End
         Begin prjDownloadBooster.CommandButtonW cmdSaveTheme 
            Height          =   300
            Left            =   4035
            TabIndex        =   48
            Top             =   0
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            Caption         =   "저장(&V)..."
         End
         Begin VB.ComboBox cbTheme 
            Height          =   300
            Left            =   840
            Style           =   2  '드롭다운 목록
            TabIndex        =   47
            Top             =   0
            Width           =   3075
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "테마(&T):"
            Height          =   180
            Left            =   0
            TabIndex        =   46
            Top             =   45
            Width           =   690
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "스킨"
         Height          =   1200
         Left            =   3570
         TabIndex        =   70
         Top             =   3870
         Width           =   3345
         Begin VB.ComboBox cbProgressSkin 
            Height          =   300
            Left            =   870
            Style           =   2  '드롭다운 목록
            TabIndex        =   77
            Top             =   840
            Width           =   1785
         End
         Begin prjDownloadBooster.CommandButtonW cmdAdvancedSkin 
            Height          =   300
            Left            =   2700
            TabIndex        =   75
            Top             =   495
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            ImageListAlignment=   4
         End
         Begin VB.ComboBox cbSkin 
            Height          =   300
            Left            =   870
            Style           =   2  '드롭다운 목록
            TabIndex        =   74
            Top             =   495
            Width           =   1785
         End
         Begin VB.ComboBox cbFrameSkin 
            Height          =   300
            Left            =   870
            Style           =   2  '드롭다운 목록
            TabIndex        =   72
            Top             =   150
            Width           =   1785
         End
         Begin prjDownloadBooster.CommandButtonW cmdAdvancedProgressSkin 
            Height          =   300
            Left            =   2700
            TabIndex        =   78
            Top             =   840
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            ImageListAlignment=   4
         End
         Begin prjDownloadBooster.CommandButtonW cmdAdvancedFrameSkin 
            Height          =   300
            Left            =   2700
            TabIndex        =   98
            Top             =   150
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            ImageListAlignment=   4
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '투명
            Caption         =   "현황(&R):"
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   885
            Width           =   855
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "창(&W):"
            Height          =   180
            Left            =   120
            TabIndex        =   71
            Top             =   210
            Width           =   540
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "단추(&O):"
            Height          =   180
            Left            =   120
            TabIndex        =   73
            Top             =   540
            Width           =   705
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "배경"
         Height          =   1200
         Left            =   3570
         TabIndex        =   60
         Top             =   2580
         Width           =   3345
         Begin VB.CheckBox chkCenter 
            Caption         =   "중앙(&C)"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2370
            TabIndex        =   69
            Top             =   870
            Width           =   945
         End
         Begin VB.ComboBox cbImagePosition 
            Height          =   300
            Left            =   945
            Style           =   2  '드롭다운 목록
            TabIndex        =   68
            Top             =   840
            Width           =   1395
         End
         Begin VB.ComboBox lvBackgrounds 
            Height          =   300
            Left            =   945
            Style           =   2  '드롭다운 목록
            TabIndex        =   65
            Top             =   495
            Width           =   1395
         End
         Begin VB.ComboBox lvPatterns 
            Height          =   300
            Left            =   945
            Style           =   2  '드롭다운 목록
            TabIndex        =   62
            Top             =   150
            Width           =   1395
         End
         Begin prjDownloadBooster.CommandButtonW cmdChooseBackground 
            Height          =   300
            Left            =   2370
            TabIndex        =   66
            Top             =   495
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            ImageListAlignment=   4
         End
         Begin VB.Label Label22 
            BackStyle       =   0  '투명
            Caption         =   "배경(&A):"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   555
            Width           =   855
         End
         Begin VB.Label Label21 
            BackStyle       =   0  '투명
            Caption         =   "무늬(&P):"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   210
            Width           =   855
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '투명
            Caption         =   "위치(&S):"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   885
            Width           =   840
         End
         Begin VB.Label lblFillColorSelect 
            BackStyle       =   0  '투명
            Height          =   255
            Left            =   2370
            TabIndex        =   63
            Top             =   180
            Width           =   615
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
      Begin VB.PictureBox pbPreview 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000001&
         Enabled         =   0   'False
         Height          =   2085
         Left            =   120
         ScaleHeight     =   2025
         ScaleWidth      =   6735
         TabIndex        =   120
         TabStop         =   0   'False
         Tag             =   "nobgdraw"
         Top             =   120
         Width           =   6795
         Begin VB.PictureBox pbBackground 
            Height          =   1320
            Left            =   240
            ScaleHeight     =   1260
            ScaleWidth      =   3975
            TabIndex        =   121
            TabStop         =   0   'False
            Tag             =   "nobgdraw"
            Top             =   120
            Width           =   4035
            Begin VB.TextBox txtSampleClassic 
               Height          =   270
               Left            =   1140
               TabIndex        =   104
               Top             =   60
               Visible         =   0   'False
               Width           =   2415
            End
            Begin VB.CheckBox CheckBoxW1 
               Caption         =   "완료 후 열기"
               Height          =   255
               Left            =   180
               TabIndex        =   122
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox TextBoxW1 
               Height          =   270
               Left            =   1140
               TabIndex        =   123
               Top             =   60
               Width           =   2415
            End
            Begin VB.Frame FrameW5 
               Caption         =   "다운로드 현황"
               Height          =   555
               Left            =   180
               TabIndex        =   124
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
                  Value           =   24
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
                  Value           =   24
                  Step            =   10
                  State           =   3
               End
            End
            Begin prjDownloadBooster.CommandButtonW cmdSample 
               Height          =   285
               Left            =   2220
               TabIndex        =   125
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
               TabIndex        =   119
               Top             =   120
               Width           =   975
            End
            Begin VB.Image imgPreview 
               Height          =   135
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   135
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
         End
      End
   End
   Begin prjDownloadBooster.CommandButtonW cmdApply 
      Height          =   360
      Left            =   9840
      TabIndex        =   118
      Top             =   120
      Width           =   1320
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   0   'False
      Caption         =   "적용(&A)"
   End
   Begin prjDownloadBooster.TabStrip tsTabStrip 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   661
      TabFixedWidth   =   53
      TabMinWidth     =   48
      Separators      =   0   'False
      TabScrollWheel  =   0   'False
      DoubleBuffer    =   0   'False
      InitTabs        =   "frmOptions.frx":1BD2
   End
   Begin prjDownloadBooster.CommandButtonW CancelButton 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   8400
      TabIndex        =   117
      Top             =   120
      Width           =   1320
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "취소"
   End
   Begin prjDownloadBooster.CommandButtonW OKButton 
      Default         =   -1  'True
      Height          =   360
      Left            =   6960
      TabIndex        =   116
      Top             =   120
      Width           =   1320
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "확인"
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'참고 자료:
'- https://www.vbforums.com/showthread.php?284592-Listview-StartLabelEdit-second-column-*RESOLVED*

Option Explicit

#If DISABLEFRAMESKIN Then
#Else
Public SkinnedFrame As frmSkinnedFrame
Public PreviewSkinnedFrame As frmSkinnedFrame
#End If

Dim Loaded As Boolean
Public GlassChanged As Boolean
Public ColorChanged As Boolean
Public ImageChanged As Boolean
Public VisualStyleChanged As Boolean
Dim SkinChanged As Boolean
Public ProgressSkinChanged As Boolean
Public FontChanged As Boolean
Dim PatternChanged As Boolean
Dim ScrollChanged As Boolean
Dim IntervalValues(7) As Single
Public ChangedBackgroundPath$
Dim PreviewControls(5) As Control
Dim DoLoadTheme As Boolean
Dim BackgroundParentDir As String
Dim PatternsSplit() As String
Dim PatternL As Byte, PatternU As Byte

Public RoundClassicButtons As Byte, DisableVisualStyle As Byte
Public LiveBadukMemoSkinShadowColor&, LiveBadukMemoSkinFrameColor&, LiveBadukMemoSkinFrameType$, LiveBadukMemoSkinTextColor&, LiveBadukMemoSkinEnableShadow As Byte, LiveBadukMemoSkinEnableTextColor As Byte, LiveBadukMemoSkinEnableBorder As Byte, LiveBadukMemoSkinFrameBackgroundType$, LiveBadukMemoSkinFrameBackgroundColor&, LiveBadukMemoSkinContentTextColor&, LiveBadukMemoSkinFrameTexture$, LiveBadukMemoSkinFrameBackground$, LiveBadukMemoSkinLabelFontSize As Integer, LiveBadukMemoSkinLabelFontBold As Byte, LiveBadukMemoSkinEnableLabelFontSize As Byte
Public ClassicFrame As Byte, NoDWMFrame As Byte

Implements IBSSubclass

Private Sub CancelButton_Click()
    Unload Me
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
    
    Dim i As Byte
    For i = LBound(PreviewControls) To UBound(PreviewControls)
        PreviewControls(i).Font.Name = FontName
        PreviewControls(i).Font.Size = FontSize
        PreviewControls(i).Font.Bold = False
        PreviewControls(i).Font.Italic = False
    Next i
    
    lblFontPreview = FontSize & " pt. " & FontName
    
    If Loaded Then
        cmdApply.Enabled = -1
        FontChanged = True
    End If
End Sub

Private Sub cbFont_Change()
    OnFontChange
End Sub

Private Sub cbFont_Click()
    OnFontChange
End Sub

Private Sub cbFrameSkin_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        SkinChanged = True
    End If
    
#If DISABLEFRAMESKIN Then
#Else
    PreviewSkinnedFrame.SetSkin cbFrameSkin.ListIndex
#End If
    
    cmdAdvancedFrameSkin.Enabled = (cbFrameSkin.ListIndex = 0 And chkAeroWindow = 0)
End Sub

Private Sub cbImagePosition_Click()
    chkCenter.Enabled = (cbImagePosition.ListIndex >= 1 And cbImagePosition.ListIndex <= 3 And chkAeroWindow = 0 And lvBackgrounds.ListIndex <> 0)
    If Loaded Then
        cmdApply.Enabled = -1
        ImageChanged = True
    End If
End Sub

Private Sub cbLanguage_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub cbProgressSkin_Click()
    cmdAdvancedProgressSkin.Enabled = (cbProgressSkin.ListIndex <> 0)
    If Loaded Then
        ProgressSkinChanged = True
        cmdApply.Enabled = -1
    End If
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
        cmdApply.Enabled = -1
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

Private Sub LoadTheme(Optional ByVal ThemeName As String = "")
    Dim Section$
    If ThemeName = "" Then Section = "Options" Else Section = "Options\Themes\" & ThemeName
    On Error Resume Next
    
    chkBeepWhenComplete = GetSetting("DownloadBooster", Section, "PlaySound", 1)
    txtCompleteSoundPath = Trim$(GetSetting("DownloadBooster", Section, "CompleteSoundPath", ""))
    
    chkAsterisk = GetSetting("DownloadBooster", Section, "EnableAsteriskSound", 1)
    chkExclamation = GetSetting("DownloadBooster", Section, "EnableExclamationSound", 1)
    chkError = GetSetting("DownloadBooster", Section, "EnableErrorSound", 1)
    chkQuestion = GetSetting("DownloadBooster", Section, "EnableQuestionSound", 1)
    txtAsterisk = GetSetting("DownloadBooster", Section, "AsteriskSound", "")
    txtExclamation = GetSetting("DownloadBooster", Section, "ExclamationSound", "")
    txtError = GetSetting("DownloadBooster", Section, "ErrorSound", "")
    txtQuestion = GetSetting("DownloadBooster", Section, "QuestionSound", "")
    
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
    
    cbFont = Trim$(GetSetting("DownloadBooster", Section, "Font", ""))
    If cbFont = "" Then cbFont = ("(" & t("기본값", "default") & ")")
    
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

Private Sub cbStartupPosition_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub cbTheme_Click()
    cmdDeleteTheme.Enabled = (cbTheme.ListIndex > 0)
    If cbTheme.ListIndex = 0 Then Exit Sub
    If Not Loaded Then Exit Sub
    
    On Error Resume Next
    
    Dim ThemeName$
    ThemeName = cbTheme.List(cbTheme.ListIndex)
    
    If DoLoadTheme Then
        VisualStyleChanged = True
        ImageChanged = True
        ColorChanged = True
        SkinChanged = True
        FontChanged = True
        ProgressSkinChanged = True
                
        LiveBadukMemoSkinShadowColor = CLng(GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinShadowColor", 16777215))
        LiveBadukMemoSkinFrameColor = CLng(GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinFrameColor", 11194780))
        LiveBadukMemoSkinFrameType = GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinFrameType", "solidcolor")
        LiveBadukMemoSkinTextColor = CLng(GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinTextColor", 2111255))
        LiveBadukMemoSkinEnableShadow = CByte(GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinEnableShadow", 1))
        LiveBadukMemoSkinEnableTextColor = CByte(GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinEnableTextColor", 1))
        LiveBadukMemoSkinEnableBorder = CByte(GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinEnableBorder", 1))
        LiveBadukMemoSkinFrameBackgroundType = GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinFrameBackgroundType", "transparent")
        LiveBadukMemoSkinFrameBackgroundColor = CLng(GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinFrameBackgroundColor", 16777215))
        LiveBadukMemoSkinContentTextColor = CLng(GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinContentTextColor", 0))
        LiveBadukMemoSkinFrameTexture = GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinFrameTexture", "")
        LiveBadukMemoSkinFrameBackground = GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinFrameBackground", "")
        LiveBadukMemoSkinLabelFontSize = CInt(GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinLabelFontSize", 10))
        LiveBadukMemoSkinLabelFontBold = CByte(GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinLabelFontBold", 0))
        LiveBadukMemoSkinEnableLabelFontSize = CByte(GetSetting("DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinEnableLabelFontSize", 0))
        
        LoadTheme ThemeName
    End If
    
    RedrawPreview
    
    cmdApply.Enabled = True
End Sub

Private Sub cbWhenExist_Click()
    If Loaded Then
        cmdApply.Enabled = -1
    End If
End Sub

Private Sub chkAeroWindow_Click()
    Dim Enable As Boolean: Enable = (chkAeroWindow = 0)
    chkBackColorMainOnly.Enabled = (Enable = True And optUserColor.Value = True)
    Label10.Enabled = Enable
    cbFrameSkin.Enabled = Enable
    cmdAdvancedFrameSkin.Enabled = Enable

    If Loaded Then
        GlassChanged = True
        cmdApply.Enabled = True
    End If
End Sub

Private Sub chkAllowDuplicates_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkAlwaysOnTop_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkAlwaysResume_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkAsterisk_Click()
    If Loaded Then cmdApply.Enabled = -1
    If Loaded Or chkAsterisk = 0 Then
        'EnableFrameControls fAsterisk, chkAsterisk, (chkAsterisk.Value = 1)
        txtAsterisk.Enabled = chkAsterisk
        cmdBrowseAsterisk.Enabled = chkAsterisk
        cmdTestAsterisk.Enabled = chkAsterisk
        cmdResetAsterisk.Enabled = chkAsterisk
    End If
End Sub

Private Sub chkAutoRetry_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

#If HIDEYTDL Then
#Else
Private Sub chkAutoYtdl_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub
#End If

Private Sub chkBackColorMainOnly_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ColorChanged = True
    End If
End Sub

Private Sub chkBeepWhenComplete_Click()
    If Loaded Then cmdApply.Enabled = -1
    If Loaded Or chkBeepWhenComplete = 0 Then
        'EnableFrameControls fCompleteSound, chkBeepWhenComplete, (chkBeepWhenComplete.Value = 1)
        txtCompleteSoundPath.Enabled = chkBeepWhenComplete
        cmdBrowseCompleteSound.Enabled = chkBeepWhenComplete
        cmdTestCompleteSound.Enabled = chkBeepWhenComplete
        cmdResetCompleteSound.Enabled = chkBeepWhenComplete
    End If
End Sub

Private Sub chkCenter_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ImageChanged = True
    End If
End Sub

Private Sub chkError_Click()
    If Loaded Then cmdApply.Enabled = -1
    If Loaded Or chkError = 0 Then
        'EnableFrameControls fError, chkError, (chkError.Value = 1)
        txtError.Enabled = chkError
        cmdBrowseError.Enabled = chkError
        cmdTestError.Enabled = chkError
        cmdResetError.Enabled = chkError
    End If
End Sub

Private Sub chkExclamation_Click()
    If Loaded Then cmdApply.Enabled = -1
    If Loaded Or chkExclamation = 0 Then
        'EnableFrameControls fExclamation, chkExclamation, (chkExclamation.Value = 1)
        txtExclamation.Enabled = chkExclamation
        cmdBrowseExclamation.Enabled = chkExclamation
        cmdTestExclamation.Enabled = chkExclamation
        cmdResetExclamation.Enabled = chkExclamation
    End If
End Sub

Private Sub chkForceGet_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkForceOldDialog_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkForeColorMainOnly_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ColorChanged = True
    End If
End Sub

Private Sub chkIgnore300_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkLazyElapsed_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkNoRedirectCheck_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkOpenDirWhenComplete_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkOpenWhenComplete_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkQuestion_Click()
    If Loaded Then cmdApply.Enabled = -1
    If Loaded Or chkQuestion = 0 Then
        'EnableFrameControls fQuestion, chkQuestion, (chkQuestion.Value = 1)
        txtQuestion.Enabled = chkQuestion
        cmdBrowseQuestion.Enabled = chkQuestion
        cmdTestQuestion.Enabled = chkQuestion
        cmdResetQuestion.Enabled = chkQuestion
    End If
End Sub

Private Sub chkRememberURL_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkShowProgressInTaskbar_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub chkUseServerModified_Click()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub cmdAdvancedFrameSkin_Click()
    frmSystemFrameProperties.Show vbModal
End Sub

Private Sub cmdAdvancedProgressSkin_Click()
    frmLiveBadukSkinProperties.Show vbModal, Me
End Sub

Private Sub cmdAdvancedSkin_Click()
    frmSystemSkinProperties.Show vbModal, Me
End Sub

Private Sub cmdApply_Click()
    SaveSetting "DownloadBooster", "Options", "RememberURL", chkRememberURL.Value
    SaveSetting "DownloadBooster", "Options", "NoRedirectCheck", chkNoRedirectCheck.Value
    SaveSetting "DownloadBooster", "Options", "ForceGet", chkForceGet.Value
    SaveSetting "DownloadBooster", "Options", "Ignore300", chkIgnore300.Value
    SaveSetting "DownloadBooster", "Options", "LazyElapsed", chkLazyElapsed.Value
    SaveSetting "DownloadBooster", "Options", "ForceWin31Dialog", chkForceOldDialog.Value
    SaveSetting "DownloadBooster", "Options", "AutoDetectYtdlURL", chkAutoYtdl.Value
    SaveSetting "DownloadBooster", "Options", "CompleteSoundPath", Trim$(txtCompleteSoundPath)
    SaveSetting "DownloadBooster", "Options", "AllowDuplicatesInQueue", chkAllowDuplicates.Value
    SaveSetting "DownloadBooster", "Options", "ScrollOneScreen", -optScreenPerScroll.Value
    SaveSetting "DownloadBooster", "Options", "BackColorMainOnly", chkBackColorMainOnly.Value
    SaveSetting "DownloadBooster", "Options", "ForeColorMainOnly", chkForeColorMainOnly.Value
    SaveSetting "DownloadBooster", "Options", "UseServerModifiedDate", chkUseServerModified.Value
    SaveSetting "DownloadBooster", "Options", "StartupPosition", cbStartupPosition.ListIndex
    SaveSetting "DownloadBooster", "Options", "ShowProgressInTaskbar", chkShowProgressInTaskbar.Value
    frmMain.pbTotalProgress.ShowInTaskBar = (chkShowProgressInTaskbar.Value = 1)
    If ScrollChanged Then
        frmMain.ScrollOneScreen = optScreenPerScroll.Value
        frmMain.trThreadCount_Scroll
        frmMain.pbProgressContainer.Top = 0
        frmMain.vsProgressScroll.Value = 0
        frmMain.pbProgressContainer.Refresh
        Dim LargeChange As Byte
        If optScreenPerScroll.Value Then LargeChange = 1 Else LargeChange = 10
        frmMain.vsProgressScroll.LargeChange = LargeChange
    End If
    If trRequestInterval.Value < 8 Then
        SaveSetting "DownloadBooster", "Options", "ThreadRequestInterval", CInt(IntervalValues(trRequestInterval.Value) * 1000)
        trRequestInterval.Max = 7
    End If
    
    If PatternChanged Then
        SaveSetting "DownloadBooster", "Options", "FormFillStyle", lvPatterns.ListIndex
        SaveSetting "DownloadBooster", "Options", "FormFillColor", pgPatternColor.BackColor
        frmMain.SetPattern
        frmMain.SetBackgroundPosition
    End If
    
    Dim NoDisable As Boolean
    NoDisable = False
    
    On Error GoTo maxtrdnotint
    If Not IsNumeric(txtMaxThreadCount.Value) Then
maxtrdnotint:
        MsgBox t("최대 쓰레드 개수는 정수여야 합니다.", "Maximum number of threads should be an integer."), 16
        NoDisable = True
        GoTo aftermaxtrdcheck
    ElseIf txtMaxThreadCount.Value > MAX_THREAD_COUNT_CONTROL Or txtMaxThreadCount.Text < 2& Then
        MsgBox t("최대 쓰레드 개수는 2개 이상 " & MAX_THREAD_COUNT_CONTROL & "개 이하여야 합니다.", "Maximum number of threads should range in 2-" & MAX_THREAD_COUNT_CONTROL & "."), 16
        NoDisable = True
    Else
        SaveSetting "DownloadBooster", "Options", "MaxThreadCount", txtMaxThreadCount.Value
        If txtMaxThreadCount.Value = 655 Then txtMaxThreadCount.Value = 655
'        If CInt(txtMaxThreadCount.Text) > 50 Then
'            MsgBox t("최대 쓰레드 개수가 너무 클 경우 실제 사용 개수와 관계없이 실행 속도가 느려질 수 있습니다.", "If the maximum number of threads is too high, the application might run slower."), 48
'        End If
    End If
aftermaxtrdcheck:
    On Error GoTo 0
    
    SaveSetting "DownloadBooster", "Options", "EnableAsteriskSound", chkAsterisk.Value
    SaveSetting "DownloadBooster", "Options", "EnableExclamationSound", chkExclamation.Value
    SaveSetting "DownloadBooster", "Options", "EnableErrorSound", chkError.Value
    SaveSetting "DownloadBooster", "Options", "EnableQuestionSound", chkQuestion.Value
    SaveSetting "DownloadBooster", "Options", "AsteriskSound", txtAsterisk
    SaveSetting "DownloadBooster", "Options", "ExclamationSound", txtExclamation
    SaveSetting "DownloadBooster", "Options", "ErrorSound", txtError
    SaveSetting "DownloadBooster", "Options", "QuestionSound", txtQuestion
    
    SaveSetting "DownloadBooster", "Options", "OpenWhenComplete", chkOpenWhenComplete.Value
    SaveSetting "DownloadBooster", "Options", "OpenFolderWhenComplete", chkOpenDirWhenComplete.Value
    SaveSetting "DownloadBooster", "Options", "PlaySound", chkBeepWhenComplete.Value
    SaveSetting "DownloadBooster", "Options", "ContinueDownload", chkAlwaysResume.Value
    SaveSetting "DownloadBooster", "Options", "AutoRetry", chkAutoRetry.Value
    SaveSetting "DownloadBooster", "Options", "WhenFileExists", cbWhenExist.ListIndex
    frmMain.cbWhenExist.ListIndex = cbWhenExist.ListIndex
    
    frmMain.chkOpenAfterComplete.Value = chkOpenWhenComplete.Value
    frmMain.chkOpenFolder.Value = chkOpenDirWhenComplete.Value
    frmMain.chkContinueDownload.Value = chkAlwaysResume.Value
    frmMain.chkAutoRetry.Value = chkAutoRetry.Value
    
    If optSystemColor.Value Then
        SaveSetting "DownloadBooster", "Options", "BackColor", "-1"
    ElseIf optUserColor.Value Then
        SaveSetting "DownloadBooster", "Options", "BackColor", CLng(pgColor.BackColor)
    End If
    If optSystemFore.Value Then
        SaveSetting "DownloadBooster", "Options", "ForeColor", "-1"
    ElseIf optUserFore.Value Then
        SaveSetting "DownloadBooster", "Options", "ForeColor", CLng(pgFore.BackColor)
    End If
    Dim DwmEnabled As Boolean: DwmEnabled = IsDWMEnabled()
    If DwmEnabled And GlassChanged Then
        SaveSetting "DownloadBooster", "Options", "UseAeroWindow", chkAeroWindow
        If chkAeroWindow = 1 Then frmMain.EnableGlassWindow _
        Else frmMain.DisableGlassWindow
    End If
    If ColorChanged Or VisualStyleChanged Or SkinChanged Or GlassChanged Then
        SaveSetting "DownloadBooster", "Options", "DisableVisualStyle", DisableVisualStyle
        SaveSetting "DownloadBooster", "Options", "RoundClassicButtons", RoundClassicButtons
        
        SaveSetting "DownloadBooster", "Options", "UseClassicThemeFrame", ClassicFrame
        SaveSetting "DownloadBooster", "Options", "DisableDWMWindow", NoDWMFrame
        
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinShadowColor", LiveBadukMemoSkinShadowColor
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinFrameColor", LiveBadukMemoSkinFrameColor
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinFrameType", LiveBadukMemoSkinFrameType
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinTextColor", LiveBadukMemoSkinTextColor
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinEnableShadow", LiveBadukMemoSkinEnableShadow
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinEnableTextColor", LiveBadukMemoSkinEnableTextColor
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinEnableBorder", LiveBadukMemoSkinEnableBorder
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinFrameBackgroundType", LiveBadukMemoSkinFrameBackgroundType
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinFrameBackgroundColor", LiveBadukMemoSkinFrameBackgroundColor
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinContentTextColor", LiveBadukMemoSkinContentTextColor
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinFrameTexture", LiveBadukMemoSkinFrameTexture
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinFrameBackground", LiveBadukMemoSkinFrameBackground
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinLabelFontSize", LiveBadukMemoSkinLabelFontSize
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinLabelFontBold", LiveBadukMemoSkinLabelFontBold
        SaveSetting "DownloadBooster", "Options", "LiveBadukMemoSkinEnableLabelFontSize", LiveBadukMemoSkinEnableLabelFontSize
    
        SaveSetting "DownloadBooster", "Options", "ButtonSkin", cbSkin.ListIndex
        SaveSetting "DownloadBooster", "Options", "WindowSkin", cbFrameSkin.ListIndex
        CurrentButtonSkin = cbSkin.ListIndex
        CurrentWindowSkin = cbFrameSkin.ListIndex
        
        SetFormBackgroundColor Me, True
        If chkAeroWindow = 1 And DwmEnabled Then
            SetFormBackgroundColor frmMain, True, frmMain.clrKey
        Else
            SetFormBackgroundColor frmMain, True
        End If
        RedrawPreview
        'cmdChooseBackground.Refresh
        frmMain.pbProgressContainer.Refresh
        frmMain.SetTextColors
        
        If NoDWMFrame And chkAeroWindow = 0 Then
            DisableDWMWindow Me.hWnd
            DisableDWMWindow frmMain.hWnd
        Else
            EnableDWMWindow Me.hWnd
            EnableDWMWindow frmMain.hWnd
        End If
        
#If DISABLEFRAMESKIN Then
#Else
        frmMain.SkinnedFrame.SetSkin CurrentWindowSkin
        SkinnedFrame.SetSkin CurrentWindowSkin
#End If
    End If
    If ProgressSkinChanged Or GlassChanged Then
        SaveSetting "DownloadBooster", "Options", "ProgressFrameSkin", cbProgressSkin.ListIndex
        frmMain.LoadLiveBadukSkin
        frmMain.SetupSplitButtons
    End If
'    If VisualStyleChanged Then
'        On Error Resume Next
'        DrawTabBackground Me, tsTabStrip, pbPanel
'        cmdChooseBackground.Refresh
'        cmdSample.Refresh
'        On Error GoTo 0
'    End If
    If cbLanguage.ListIndex = 0 Then
        SaveSetting "DownloadBooster", "Options", "Language", "0"
    ElseIf cbLanguage.ListIndex = 1 Then
        SaveSetting "DownloadBooster", "Options", "Language", 1042
    Else
        SaveSetting "DownloadBooster", "Options", "Language", 1033
    End If
    
    Dim SaveImgPos As Byte
    If chkCenter.Value <> 0 And cbImagePosition.ListIndex >= 1 And cbImagePosition.ListIndex <= 3 Then
        SaveImgPos = cbImagePosition.ListIndex + 3
    ElseIf cbImagePosition.ListIndex = 4 Then
        SaveImgPos = 7
    Else
        SaveImgPos = cbImagePosition.ListIndex
    End If
    SaveSetting "DownloadBooster", "Options", "ImagePosition", SaveImgPos
    frmMain.ImagePosition = SaveImgPos
    If ImageChanged Or GlassChanged Then
        SaveSetting "DownloadBooster", "Options", "UseBackgroundImage", -(lvBackgrounds.ListIndex <> 0)
        SaveSetting "DownloadBooster", "Options", "BackgroundImagePath", ChangedBackgroundPath
        frmMain.SetBackgroundImage
        frmMain.SetBackgroundPosition True
    End If
    
    If LenB(Trim$(txtNodePath)) Then
        If FileExists(Trim$(txtNodePath)) Then
            SaveSetting "DownloadBooster", "Options", "NodePath", Trim$(txtNodePath)
        Else
            MsgBox t("Node.js 경로가 존재하지 않습니다.", "Node.js path does not exist."), 16
            NoDisable = True
        End If
    Else
        SaveSetting "DownloadBooster", "Options", "NodePath", ""
    End If
    If LenB(Trim$(txtYtdlPath)) Then
        If FileExists(Trim$(txtYtdlPath)) Then
            SaveSetting "DownloadBooster", "Options", "YtdlPath", Trim$(txtYtdlPath)
        Else
            MsgBox t("Youtube-dl 경로가 존재하지 않습니다.", "Youtube-dl path does not exist."), 16
            NoDisable = True
        End If
    Else
        SaveSetting "DownloadBooster", "Options", "YtdlPath", ""
    End If
    
    If FontChanged Then
        cbFont = Trim$(cbFont)
        If LenB(cbFont) And cbFont <> ("(" & t("기본값", "default") & ")") And (Not FontExists(cbFont)) Then
            MsgBox t("지정한 글꼴이 존재하지 않습니다.", "The specified font does not exist."), vbCritical
            NoDisable = True
        Else
            If cbFont = ("(" & t("기본값", "default") & ")") Then
                SaveSetting "DownloadBooster", "Options", "Font", ""
            Else
                SaveSetting "DownloadBooster", "Options", "Font", cbFont
            End If
            SetFont Me, True
            SetFont frmMain, True
        End If
        InitPropertySheetDimensions Me, tsTabStrip, pbPanel, OKButton, CancelButton, cmdApply
    End If
    
    If lvBackgrounds.ListIndex <> 0 And LenB(GetSetting("DownloadBooster", "Options", "BackgroundImagePath", "")) = 0 Then
        MsgBox t("배경 그림이 선택되지 않았습니다.", "Background image is not selected."), 48
        SaveSetting "DownloadBooster", "Options", "UseBackgroundImage", "0"
        NoDisable = True
    End If
    
    Dim hSysMenu As Long
    Dim MII As MENUITEMINFO
    hSysMenu = GetSystemMenu(frmMain.hWnd, 0)
    MainFormOnTop = (chkAlwaysOnTop.Value = 1)
    Dim InsertAfter As Long
    If MainFormOnTop Then InsertAfter = hWnd_TOPMOST Else InsertAfter = hWnd_NOTOPMOST
    SetWindowPos frmMain.hWnd, InsertAfter, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    SetWindowPos Me.hWnd, InsertAfter, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    With MII
        .cbSize = Len(MII)
        .fMask = MIIM_STATE
        .fState = MFS_ENABLED
        If MainFormOnTop Then .fState = .fState Or MFS_CHECKED
    End With
    SetMenuItemInfo hSysMenu, 1000, 0, MII
    SaveSetting "DownloadBooster", "Options", "AlwaysOnTop", -MainFormOnTop
    
    If optUserFore.Value Then
        RemoveVisualStyles CheckBoxW1.hWnd
        RemoveVisualStyles FrameW5.hWnd
        CheckBoxW1.ForeColor = pgFore.BackColor
        FrameW5.ForeColor = pgFore.BackColor
    End If
    
    Dim SaveThemeName As String
    If cbTheme.ListIndex = 0 Then SaveThemeName = "" Else SaveThemeName = cbTheme.List(cbTheme.ListIndex)
    SaveSetting "DownloadBooster", "Options", "Theme", SaveThemeName
    
    RedrawPreview
    ResetChanged
    If Not NoDisable Then cmdApply.Enabled = 0
End Sub

Private Sub ResetChanged()
    ColorChanged = False
    ImageChanged = False
    VisualStyleChanged = False
    SkinChanged = False
    ProgressSkinChanged = False
    ScrollChanged = False
    FontChanged = False
    PatternChanged = False
    GlassChanged = False
End Sub

Private Sub cmdBrowseAsterisk_Click()
    Set Tags.BrowseTargetTextbox = txtAsterisk
    ShowFileDialog 4, Tags.BrowseTargetTextbox.Text, True
End Sub

Private Sub cmdBrowseCompleteSound_Click()
    Set Tags.BrowseTargetTextbox = txtCompleteSoundPath
    ShowFileDialog 4, Tags.BrowseTargetTextbox.Text, True
End Sub

Private Sub cmdBrowseError_Click()
    Set Tags.BrowseTargetTextbox = txtError
    ShowFileDialog 4, Tags.BrowseTargetTextbox.Text, True
End Sub

Private Sub cmdBrowseExclamation_Click()
    Set Tags.BrowseTargetTextbox = txtExclamation
    ShowFileDialog 4, Tags.BrowseTargetTextbox.Text, True
End Sub

Private Sub cmdBrowseQuestion_Click()
    Set Tags.BrowseTargetTextbox = txtQuestion
    ShowFileDialog 4, Tags.BrowseTargetTextbox.Text, True
End Sub

Private Sub cmdChooseBackground_Click()
    ShowFileDialog 3, GetSetting("DownloadBooster", "Options", "BackgroundImagePath", ""), ForceNewDialog:=True
End Sub

Private Sub cmdDeleteTheme_Click()
    If cbTheme.ListIndex = 0 Then Exit Sub
    On Error Resume Next
    If MsgBox(t("선택한 테마를 삭제하시겠습니까?", "Delete the selected theme?"), vbQuestion + vbYesNo) = vbYes Then
        DeleteSetting "DownloadBooster", "Options\Themes\" & cbTheme.List(cbTheme.ListIndex)
        cbTheme.RemoveItem cbTheme.ListIndex
        cbTheme.ListIndex = 0
    End If
End Sub

Private Sub cmdHeaders_Click()
    Tags.DownloadOptionsTargetForm = 3
    Load frmDownloadOptions
#If HIDEYTDL Then
#Else
    frmDownloadOptions.RemoveYtdlTab
#End If
    frmDownloadOptions.Show vbModal, Me
End Sub

Private Sub cmdResetAsterisk_Click()
    txtAsterisk = ""
End Sub

Private Sub cmdResetCompleteSound_Click()
    txtCompleteSoundPath = ""
End Sub

Private Sub cmdResetError_Click()
    txtError = ""
End Sub

Private Sub cmdResetExclamation_Click()
    txtExclamation = ""
End Sub

Private Sub cmdResetQuestion_Click()
    txtQuestion = ""
End Sub

Private Sub cmdSaveTheme_Click()
    Dim ThemeName$, DefThemeName$
    If cbTheme.ListIndex = 0 Then DefThemeName = "" Else DefThemeName = cbTheme.List(cbTheme.ListIndex)
    ThemeName = InputBoxEx(t("테마 이름을 입력하십시오.", "Choose your theme name."), t("테마 저장", "Save theme"), DefThemeName)
    If ThemeName = "" Then
        Exit Sub
    ElseIf InStr(ThemeName, "\") Then
        MsgBox t("테마 이름에 허용되지 않은 문자가 포함되어 있습니다.", "Theme name contains invalid characters."), 16
        Exit Sub
    ElseIf ThemeName = "수정된 테마" Or LCase(ThemeName) = "modified theme" Then
        MsgBox t("테마 이름이 올바르지 않습니다.", "Theme name is invalid."), 16
        Exit Sub
    End If
    
    On Error Resume Next
    DeleteSetting "DownloadBooster", "Options\Themes\" & ThemeName
    
    cbFont = Trim$(cbFont)
    If Not (cbFont <> "" And cbFont <> ("(" & t("기본값", "default") & ")") And (Not FontExists(cbFont))) Then
        If cbFont = ("(" & t("기본값", "default") & ")") Then
            SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "Font", ""
        Else
            SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "Font", cbFont
        End If
    End If
    
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "CompleteSoundPath", Trim$(txtCompleteSoundPath)
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "BackColorMainOnly", chkBackColorMainOnly.Value
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "ForeColorMainOnly", chkForeColorMainOnly.Value
    
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "FormFillStyle", lvPatterns.ListIndex
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "FormFillColor", pgPatternColor.BackColor
    
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "EnableAsteriskSound", chkAsterisk.Value
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "EnableExclamationSound", chkExclamation.Value
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "EnableErrorSound", chkError.Value
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "EnableQuestionSound", chkQuestion.Value
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "AsteriskSound", txtAsterisk
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "ExclamationSound", txtExclamation
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "ErrorSound", txtError
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "QuestionSound", txtQuestion
    
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "PlaySound", chkBeepWhenComplete.Value
    
    If optSystemColor.Value Then
        SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "BackColor", "-1"
    ElseIf optUserColor.Value Then
        SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "BackColor", CLng(pgColor.BackColor)
    End If
    If optSystemFore.Value Then
        SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "ForeColor", "-1"
    ElseIf optUserFore.Value Then
        SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "ForeColor", CLng(pgFore.BackColor)
    End If
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "ButtonSkin", cbSkin.ListIndex
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "WindowSkin", cbFrameSkin.ListIndex
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "ProgressFrameSkin", cbProgressSkin.ListIndex
    Dim SaveImgPos As Byte
    If cbImagePosition.ListIndex >= 1 And cbImagePosition.ListIndex <= 3 And chkCenter.Value <> 0 Then
        SaveImgPos = cbImagePosition.ListIndex + 3
    ElseIf cbImagePosition.ListIndex = 4 Then
        SaveImgPos = 7
    Else
        SaveImgPos = cbImagePosition.ListIndex
    End If
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "ImagePosition", SaveImgPos
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "UseBackgroundImage", -(lvBackgrounds.ListIndex <> 0)
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "BackgroundImagePath", ChangedBackgroundPath
    
    If lvBackgrounds.ListIndex <> 0 And LenB(GetSetting("DownloadBooster", "Options", "BackgroundImagePath", "")) = 0 Then
        SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "UseBackgroundImage", "0"
    End If
    
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "DisableVisualStyle", DisableVisualStyle
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "RoundClassicButtons", RoundClassicButtons
        
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "UseClassicThemeFrame", ClassicFrame
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "DisableDWMWindow", NoDWMFrame
    
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinShadowColor", LiveBadukMemoSkinShadowColor
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinFrameColor", LiveBadukMemoSkinFrameColor
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinFrameType", LiveBadukMemoSkinFrameType
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinTextColor", LiveBadukMemoSkinTextColor
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinEnableShadow", LiveBadukMemoSkinEnableShadow
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinEnableTextColor", LiveBadukMemoSkinEnableTextColor
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinEnableBorder", LiveBadukMemoSkinEnableBorder
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinFrameBackgroundType", LiveBadukMemoSkinFrameBackgroundType
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinFrameBackgroundColor", LiveBadukMemoSkinFrameBackgroundColor
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinContentTextColor", LiveBadukMemoSkinContentTextColor
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinFrameTexture", LiveBadukMemoSkinFrameTexture
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinFrameBackground", LiveBadukMemoSkinFrameBackground
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinLabelFontSize", LiveBadukMemoSkinLabelFontSize
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinLabelFontBold", LiveBadukMemoSkinLabelFontBold
    SaveSetting "DownloadBooster", "Options\Themes\" & ThemeName, "LiveBadukMemoSkinEnableLabelFontSize", LiveBadukMemoSkinEnableLabelFontSize
    
    Dim i%, ThemeFound As Boolean
    ThemeFound = False
    For i = 1 To cbTheme.ListCount - 1
        If cbTheme.List(i) = ThemeName Then
            ThemeFound = True
            cbTheme.ListIndex = i
        End If
    Next i
    If Not ThemeFound Then
        AddItemToComboBox cbTheme, ThemeName
        cbTheme.ListIndex = cbTheme.ListCount - 1
    End If
End Sub

Private Sub cmdTestAsterisk_Click()
    txtAsterisk = Trim$(txtAsterisk)
    If txtAsterisk = "" Then
        MessageBeep 64
    Else
        PlayWave txtAsterisk
    End If
End Sub

Private Sub cmdTestCompleteSound_Click()
    txtCompleteSoundPath = Trim$(txtCompleteSoundPath)
    If txtCompleteSoundPath = "" Then
        MessageBeep 64
    Else
        PlayWave txtCompleteSoundPath
    End If
End Sub

Private Sub cmdTestError_Click()
    txtError = Trim$(txtError)
    If txtError = "" Then
        MessageBeep 16
    Else
        PlayWave txtError
    End If
End Sub

Private Sub cmdTestExclamation_Click()
    txtExclamation = Trim$(txtExclamation)
    If txtExclamation = "" Then
        MessageBeep 48
    Else
        PlayWave txtExclamation
    End If
End Sub

Private Sub cmdTestQuestion_Click()
    txtQuestion = Trim$(txtQuestion)
    If txtQuestion = "" Then
        MessageBeep 32
    Else
        PlayWave txtQuestion
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    IBSSubclass_UnsubclassIt
    
#If DISABLEFRAMESKIN Then
#Else
    Unload PreviewSkinnedFrame
    Unload SkinnedFrame
#End If
End Sub

Private Function IBSSubclass_MsgResponse(ByVal hWnd As Long, ByVal uMsg As Long) As EMsgResponse
    IBSSubclass_MsgResponse = emrConsume
End Function

Private Sub IBSSubclass_UnsubclassIt()
    DetachMessage Me, Me.hWnd, WM_SETTINGCHANGE
    'DetachMessage Me, Me.hWnd, WM_THEMECHANGED
    DetachMessage Me, Frame5.hWnd, WM_PRINTCLIENT
    DetachMessage Me, Frame1.hWnd, WM_PRINTCLIENT
    DetachMessage Me, FrameW4.hWnd, WM_PRINTCLIENT
End Sub

Private Function IBSSubclass_WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
    On Error Resume Next
 
    Select Case uMsg
        Case WM_SETTINGCHANGE
            Select Case GetStrFromPtr(lParam)
                Case "WindowMetrics"
                    UpdateBorderWidth
                    SetPreviewPosition
                    'DrawTabBackground Me, tsTabStrip, pbPanel
            End Select
        'Case WM_THEMECHANGED
            'DrawTabBackground Me, tsTabStrip, pbPanel
        Case WM_PRINTCLIENT
            SendMessage hWnd, WM_PAINT, wParam, ByVal 0&
            
            IBSSubclass_WindowProc = 0&
            Exit Function
    End Select
    
    IBSSubclass_WindowProc = CallOldWindowProc(hWnd, uMsg, wParam, lParam)
End Function

Private Sub lblFillColorSelect_Click()
    Dim Color As OLE_COLOR
    Color = ShowColorDialog(Me.hWnd, True, pgPatternColor.BackColor, True)
    If Color = -1 Then Exit Sub
    pgPatternColor.BackColor = Color
    cmdApply.Enabled = -1
    pgPatternPreview.FillColor = pgPatternColor.BackColor
    PatternChanged = True
    RedrawPreview
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
        cmdApply.Enabled = -1
        ImageChanged = True
        RedrawPreview
    End If
End Sub

Private Sub lvPatterns_Click()
    pgPatternPreview.FillStyle = lvPatterns.ListIndex + 1
    If Loaded Then
        cmdApply.Enabled = -1
        PatternChanged = True
        RedrawPreview
    End If
End Sub

Private Sub optLinePerScroll_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ScrollChanged = True
    End If
End Sub

Private Sub optScreenPerScroll_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ScrollChanged = True
    End If
End Sub

Private Sub trRequestInterval_Change()
    trRequestInterval_Scroll
End Sub

Private Sub trRequestInterval_Scroll()
    If trRequestInterval.Value = 8 Then
        lblIntervalDisplay.Caption = "(" & t("사용자 지정", "Customized") & ")"
    Else
        lblIntervalDisplay.Caption = "(" & IntervalValues(trRequestInterval.Value) & t("초", " second" & IIf(IntervalValues(trRequestInterval.Value) = 1, "", "s")) & ")"
    End If
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub txtAsterisk_Change()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub txtCompleteSoundPath_Change()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub Form_Load()
    InitForm Me
    
    Loaded = False
    PatternsSplit = Split("*.jpg;*.jpeg;*.jpe;*.jfif;*.gif;*.bmp;*.dib;*.png;*.wmf;*.emf;*.ico;*.cur;*.tif;*.tiff;*.rle", ";")
    PatternL = LBound(PatternsSplit)
    PatternU = UBound(PatternsSplit)
    
#If HIDEYTDL Then
    If LaunchFromMemory Then tsTabStrip.Tabs.Remove 5
#End If

    Set PreviewControls(0) = Label11
    Set PreviewControls(1) = TextBoxW1
    Set PreviewControls(2) = FrameW5
    Set PreviewControls(3) = CheckBoxW1
    Set PreviewControls(4) = cmdSample
    Set PreviewControls(5) = lblFontPreview
    
    IntervalValues(0) = 0.01
    IntervalValues(1) = 0.05
    IntervalValues(2) = 0.1
    IntervalValues(3) = 0.3
    IntervalValues(4) = 0.5
    IntervalValues(5) = 1#
    IntervalValues(6) = 3#
    IntervalValues(7) = 5#
    
    Set cmdAdvancedFrameSkin.ImageList = frmMain.imgWrench
    Set cmdAdvancedSkin.ImageList = frmMain.imgWrench
    Set cmdAdvancedProgressSkin.ImageList = frmMain.imgWrench
    
    Set cmdChooseBackground.ImageList = frmMain.imgOpenFolder
    Set cmdBrowseCompleteSound.ImageList = frmMain.imgOpenFolder
    Set cmdBrowseAsterisk.ImageList = frmMain.imgOpenFolder
    Set cmdBrowseExclamation.ImageList = frmMain.imgOpenFolder
    Set cmdBrowseError.ImageList = frmMain.imgOpenFolder
    Set cmdBrowseQuestion.ImageList = frmMain.imgOpenFolder
    
    Set cmdResetCompleteSound.ImageList = frmMain.imgErase
    Set cmdResetAsterisk.ImageList = frmMain.imgErase
    Set cmdResetExclamation.ImageList = frmMain.imgErase
    Set cmdResetError.ImageList = frmMain.imgErase
    Set cmdResetQuestion.ImageList = frmMain.imgErase
    
    Set OKButton.ImageList = frmDummyForm.imgOK
    Set CancelButton.ImageList = frmDummyForm.imgCancel
    
    RemoveVisualStyles txtSampleClassic.hWnd
    
    InitPropertySheetDimensions Me, tsTabStrip, pbPanel, OKButton, CancelButton, cmdApply
    
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
    
    pbBackground.Enabled = False
    SetPreviewPosition
    
#If DISABLEFRAMESKIN Then
#Else
    Set PreviewSkinnedFrame = New frmSkinnedFrame
    PreviewSkinnedFrame.Init pbBackground
    PreviewSkinnedFrame.SetCaption App.Title
#End If
    
    AddItemToComboBox cbFrameSkin, t("시스템 스타일", "System style")
#If DISABLEFRAMESKIN Then
#Else
    AddItemToComboBox cbFrameSkin, t("금속 파랑", "Blue metal")
    AddItemToComboBox cbFrameSkin, t("금속 초록", "Green metal")
    AddItemToComboBox cbFrameSkin, "Windows XP"
#End If
    
    imgPreview.Top = 0
    imgPreview.Left = 0
    
    'DrawTabBackground Me, tsTabStrip, pbPanel
    
    AddItemToComboBox cbSkin, t("시스템 스타일", "System style")
    AddItemToComboBox cbSkin, t("라이브바둑 쪽지", "LiveBaduk memo")
    AddItemToComboBox cbSkin, t("디스크키퍼 2011", "Diskeeper 2011")
    AddItemToComboBox cbSkin, t("광택 파랑", "Glossy blue")
    AddItemToComboBox cbSkin, t("광택 초록", "Glossy green")
    
    AddItemToComboBox cbProgressSkin, t("(없음)", "(None)")
    AddItemToComboBox cbProgressSkin, t("기본 스킨", "Default skin")
    AddItemToComboBox cbProgressSkin, t("라이브바둑 쪽지", "LiveBaduk memo")
    
    AddItemToComboBox cbLanguage, t("자동", "Auto")
    AddItemToComboBox cbLanguage, "한국어"
    AddItemToComboBox cbLanguage, "English"
    
    AddItemToComboBox cbWhenExist, t("건너뛰기", "Skip")
    AddItemToComboBox cbWhenExist, t("덮어쓰기", "Overwrite")
    AddItemToComboBox cbWhenExist, t("자동 이름 변경", "Auto Rename")
    
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
    
    AddItemToComboBox cbStartupPosition, t("이전 위치 기억", "Remember last position")
    AddItemToComboBox cbStartupPosition, t("화면 가운데", "Center of the screen")
    AddItemToComboBox cbStartupPosition, t("자동", "Auto")
    
    chkAeroWindow.Enabled = IsDWMEnabled()
    chkAeroWindow.Visible = (WinVer >= 6!)
    
    LoadSettings
    
    tr chkCenter, "&Center"
    tr tsTabStrip.Tabs(1), "General"
    tr tsTabStrip.Tabs(2), "Connection"
    tr tsTabStrip.Tabs(3), "Appearance"
    tr tsTabStrip.Tabs(4), "Sounds"
#If HIDEYTDL Then
    If Not LaunchFromMemory Then tr tsTabStrip.Tabs(5), "Advanced"
#End If
    tr Label5, "P&rogress:"
    tr Frame1, "Background color"
    tr FrameW4, "Text color"
    tr Frame4, "Wallpaper"
    tr Label10, "&Window:"
    tr Frame2, "Download settings"
    tr Frame5, "Interface"
    tr chkRememberURL, "Re&member URL"
    tr optSystemColor, "&System color"
    tr optSystemFore, "S&ystem color"
    tr optUserColor, "&Custom color"
    tr optUserFore, "C&ustom color"
    tr Label1, "&Language:"
    tr OKButton, "OK"
    tr CancelButton, "Cancel"
    tr cmdApply, "&Apply"
    tr Me, "Options"
    tr Frame6, "Skin"
    tr chkOpenWhenComplete, "&Open file when complete"
    tr chkOpenDirWhenComplete, "O&pen folder when complete"
    tr chkBeepWhenComplete, "Download &complete"
    tr chkAlwaysResume, "&Always resume"
    tr chkAutoRetry, "A&uto retry on error"
    tr Label3, "If filename alrea&dy exists:"
    tr Label6, "Leave the field blank to use defaults. This option is for advanced users and there is no need to change for normal use."
    tr FrameW2, "Paths"
    tr cmdSample, "Download"
    tr Label2, "Po&sition:"
    tr Label8, "Butt&on:"
    tr fHeaders, "Headers"
    tr chkNoRedirectCheck, "Don't check fo&r redirects"
    tr chkForceGet, "Force GET re&quest on file check"
    tr chkIgnore300, "&Ignore 3XX reponse code"
    tr chkAlwaysOnTop, "Al&ways on top"
    tr chkAeroWindow, "Use Aero &glass window"
    tr chkLazyElapsed, "Elapsed time sin&ce first data receive"
    tr FrameW3, "Server connection"
    tr chkForceOldDialog, "U&se Windows 3.1 dialogs"
    tr chkAutoYtdl, "Automatically use &youtube-dl for supported links"
    tr Label11, "File URL:"
    tr FrameW5, "Download status"
    tr CheckBoxW1, "Open when done"
    tr chkAsterisk, "&Asterisk"
    tr chkExclamation, "&Exclamation"
    tr chkError, "E&rror"
    tr chkQuestion, "&Question"
    tr chkAllowDuplicates, "Allow dupl&icates in queue"
    tr Label13, "&Font:"
    tr Label14, "Ma&x. number of threads:"
    tr Label15, "(restart required)"
    tr Label16, Label15.Caption
    tr FrameW6, "Download notifications"
    tr FrameW9, "Message boxes"
    tr Label17, "Set the headers when requesting to the server on download."
    tr Label18, "T&hread scroll:"
    tr optLinePerScroll, "Per li&ne"
    tr optScreenPerScroll, "Pe&r screen"
    tr Label19, "Thread request i&nterval:"
    'tr cmdAdvancedSkin, "Ad&vanced..."
    tr chkBackColorMainOnly, "&Only apply to main window"
    tr chkForeColorMainOnly, "O&nly apply to main window"
    tr Label21, "&Patterns:"
    tr Label22, "W&allpaper:"
    tr FrameW1, "Theme"
    tr Label20, "&Theme:"
    tr cmdSaveTheme, "Sa&ve..."
    tr cmdDeleteTheme, "&Delete"
    tr chkUseServerModified, "Use server's mo&dified date"
    tr cmdHeaders, "&Headers..."
    tr Label9, "S&tartup position:"
    tr chkShowProgressInTaskbar, "Show p&rogress in taskbar"
    
    If LaunchFromMemory Then
        Label4.Enabled = 0
        txtNodePath.Enabled = 0
    End If
    
    AttachMessage Me, Me.hWnd, WM_SETTINGCHANGE
    'AttachMessage Me, Me.hWnd, WM_THEMECHANGED
    AttachMessage Me, Frame5.hWnd, WM_PRINTCLIENT
    AttachMessage Me, Frame1.hWnd, WM_PRINTCLIENT
    AttachMessage Me, FrameW4.hWnd, WM_PRINTCLIENT
    
    Set cmdSample.ImageList = frmMain.imgDownload
    
#If HIDEYTDL Then
    txtYtdlPath.Visible = False
    chkAutoYtdl.Visible = False
    Label7.Visible = False
    Frame2.Height = Frame2.Height - chkAutoYtdl.Height
    'Frame2.Refresh
    Frame5.Top = Frame5.Top - chkAutoYtdl.Height
    'Frame5.Refresh
    Label3.Top = Label3.Top - chkAutoYtdl.Height
    cbWhenExist.Top = cbWhenExist.Top - chkAutoYtdl.Height
    FrameW2.Height = FrameW2.Height - txtYtdlPath.Height - 120
    'FrameW2.Refresh
    Label14.Top = Label14.Top - chkAutoYtdl.Height
    txtMaxThreadCount.Top = txtMaxThreadCount.Top - chkAutoYtdl.Height
    Label15.Top = Label15.Top - chkAutoYtdl.Height
#End If

'    Tags.DownloadOptionsTargetForm = 3
'    Load frmDownloadOptions
'#If HIDEYTDL Then
'#Else
'    frmDownloadOptions.RemoveYtdlTab
'#End If
'    SetParent frmDownloadOptions.hWnd, fHeaders.hWnd
'    frmDownloadOptions.tsTabStrip.Visible = 0
'    frmDownloadOptions.OKButton.Visible = 0
'    frmDownloadOptions.CancelButton.Visible = 0
'    frmDownloadOptions.Show
'    frmDownloadOptions.Move 0, 0

    LoadTheme
    
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
        Frame5.Height = Frame5.Height - chkShowProgressInTaskbar.Height
        Label13.Top = Label13.Top - chkShowProgressInTaskbar.Height
        cbFont.Top = cbFont.Top - chkShowProgressInTaskbar.Height
        lblFontPreview.Top = lblFontPreview.Top - chkShowProgressInTaskbar.Height
    End If
    
    Loaded = True
End Sub

Private Sub LoadSettings()
    'SetWindowPos Me.hWnd, IIf(MainFormOnTop, hWnd_TOPMOST, hWnd_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    On Error Resume Next
    
    chkNoRedirectCheck = GetSetting("DownloadBooster", "Options", "NoRedirectCheck", 0)
    chkForceGet = GetSetting("DownloadBooster", "Options", "ForceGet", 1)
    chkIgnore300 = GetSetting("DownloadBooster", "Options", "Ignore300", 0)
    chkAlwaysOnTop = -(MainFormOnTop)
    chkLazyElapsed = GetSetting("DownloadBooster", "Options", "LazyElapsed", 0)
    chkForceOldDialog = GetSetting("DownloadBooster", "Options", "ForceWin31Dialog", 0)
    chkRememberURL = GetSetting("DownloadBooster", "Options", "RememberURL", 1)
    chkAutoYtdl = GetSetting("DownloadBooster", "Options", "AutoDetectYtdlURL", 1)
    chkAllowDuplicates = GetSetting("DownloadBooster", "Options", "AllowDuplicatesInQueue", 0)
    txtMaxThreadCount = GetSetting("DownloadBooster", "Options", "MaxThreadCount", 25)
    optLinePerScroll = True
    optScreenPerScroll = (GetSetting("DownloadBooster", "Options", "ScrollOneScreen", 0) <> 0)
    chkUseServerModified = GetSetting("DownloadBooster", "Options", "UseServerModifiedDate", 1)
    chkAeroWindow = GetSetting("DownloadBooster", "Options", "UseAeroWindow", 0)
    Select Case CInt(GetSetting("DownloadBooster", "Options", "ThreadRequestInterval", 100))
        Case 10
            trRequestInterval.Value = 0
        Case 50
            trRequestInterval.Value = 1
        Case 100
            trRequestInterval.Value = 2
        Case 300
            trRequestInterval.Value = 3
        Case 500
            trRequestInterval.Value = 4
        Case 1000
            trRequestInterval.Value = 5
        Case 3000
            trRequestInterval.Value = 6
        Case 5000
            trRequestInterval.Value = 7
        Case Else
            trRequestInterval.Max = 8
            trRequestInterval.Value = 8
    End Select
    trRequestInterval_Scroll
    
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
    pgPatternPreview.Width = pbBackground.Width
    pgPatternPreview.Height = pbBackground.Height
    
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
    
    txtNodePath = GetSetting("DownloadBooster", "Options", "NodePath", "")
    txtYtdlPath = GetSetting("DownloadBooster", "Options", "YtdlPath", "")
    
    cbStartupPosition.ListIndex = GetSetting("DownloadBooster", "Options", "StartupPosition", 0)
    chkShowProgressInTaskbar.Value = GetSetting("DownloadBooster", "Options", "ShowProgressInTaskbar", 1)
    
    tsTabStrip.Tabs(1).Selected = True
    ResetChanged
    cmdApply.Enabled = False
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 And IsKeyPressed(gksKeyboardctrl) Then
        NextTabPage tsTabStrip, IsKeyPressed(gksKeyboardShift)
    End If
End Sub

Private Sub lblSelectColor_Click()
    If chkAeroWindow = 1 Then Exit Sub
    Dim Color As OLE_COLOR
    Color = ShowColorDialog(Me.hWnd, True, pgColor.BackColor)
    If Color = -1 Then Exit Sub
    pgColor.BackColor = Color
    cmdApply.Enabled = -1
    optUserColor.Value = True
    ColorChanged = True
    pbBackground.BackColor = pgColor.BackColor
    pgPatternPreview.BackColor = pgColor.BackColor
    'cmdSample.Refresh
    RedrawPreview
    FrameW5.BackColor = pgColor.BackColor
    CheckBoxW1.BackColor = pgColor.BackColor
End Sub

Private Sub lblSelectFore_Click()
    Dim Color As OLE_COLOR
    Color = ShowColorDialog(Me.hWnd, True, pgFore.BackColor, True)
    If Color = -1 Then Exit Sub
    pgFore.BackColor = Color
    cmdApply.Enabled = -1
    optUserFore.Value = True
    ColorChanged = True
    ProgressSkinChanged = True
    Label11.ForeColor = pgFore.BackColor
    RemoveVisualStyles CheckBoxW1.hWnd
    RemoveVisualStyles FrameW5.hWnd
    CheckBoxW1.ForeColor = pgFore.BackColor
    FrameW5.ForeColor = pgFore.BackColor
End Sub

Private Sub OKButton_Click()
    If cmdApply.Enabled Then cmdApply_Click
    Unload Me
End Sub

Private Sub optSystemColor_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ColorChanged = True
    End If
    pbBackground.BackColor = &H8000000F
    pgPatternPreview.BackColor = pbBackground.BackColor
    'cmdSample.Refresh
    RedrawPreview
    chkBackColorMainOnly.Enabled = False
    FrameW5.BackColor = &H8000000F
    CheckBoxW1.BackColor = &H8000000F
End Sub

Private Sub optSystemFore_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ColorChanged = True
        ProgressSkinChanged = True
    End If
    Label11.ForeColor = &H80000012
    If DisableVisualStyle = 0 Then
        ActivateVisualStyles CheckBoxW1.hWnd
        ActivateVisualStyles FrameW5.hWnd
    Else
        RemoveVisualStyles CheckBoxW1.hWnd
        RemoveVisualStyles FrameW5.hWnd
    End If
    CheckBoxW1.ForeColor = &H80000012
    FrameW5.ForeColor = &H80000012
    chkForeColorMainOnly.Enabled = False
End Sub

Private Sub optUserColor_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ColorChanged = True
    End If
    pbBackground.BackColor = pgColor.BackColor
    pgPatternPreview.BackColor = pbBackground.BackColor
    'cmdSample.Refresh
    RedrawPreview
    chkBackColorMainOnly.Enabled = (chkAeroWindow = 0)
    FrameW5.BackColor = pgColor.BackColor
    CheckBoxW1.BackColor = pgColor.BackColor
End Sub

Private Sub optUserFore_Click()
    If Loaded Then
        cmdApply.Enabled = -1
        ColorChanged = True
        ProgressSkinChanged = True
    End If
    Label11.ForeColor = pgFore.BackColor
    RemoveVisualStyles CheckBoxW1.hWnd
    RemoveVisualStyles FrameW5.hWnd
    CheckBoxW1.ForeColor = pgFore.BackColor
    FrameW5.ForeColor = pgFore.BackColor
    chkForeColorMainOnly.Enabled = True
End Sub

Private Sub tsTabStrip_TabClick(TabItem As TbsTab)
    On Error Resume Next
    
    Static i As Byte, Show As Boolean
    For i = 1 To pbPanel.Count
        Show = (i = TabItem.Index)
        pbPanel(i).Visible = Show
        pbPanel(i).Enabled = Show
    Next i
    
    If TabItem.Index = 3 Then
        DoEvents
        RedrawPreview
    End If
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

Private Sub txtError_Change()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub txtExclamation_Change()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub txtMaxThreadCount_Change()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub txtNodePath_Change()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub txtQuestion_Change()
    If Loaded Then cmdApply.Enabled = -1
End Sub

Private Sub txtYtdlPath_Change()
    If Loaded Then cmdApply.Enabled = -1
End Sub
