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
         Begin VB.Label Label11 
            BackStyle       =   0  '투명
            Caption         =   "파일 주소:"
            Height          =   255
            Left            =   180
            TabIndex        =   2
            Top             =   120
            Width           =   975
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
   Begin VB.Frame FrameW4 
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
   Begin VB.Frame Frame1 
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
