VERSION 5.00
Begin VB.Form pagConnection 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "연결"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6990
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "pagConnection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame fHeaders 
      Caption         =   "사용자 지정 헤더"
      Height          =   795
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   6780
      Begin prjDownloadBooster.CommandButtonW cmdHeaders 
         Height          =   330
         Left            =   5040
         TabIndex        =   9
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         Caption         =   "헤더(&H)..."
         LazyCreate      =   -1  'True
      End
      Begin VB.Label Label17 
         BackStyle       =   0  '투명
         Caption         =   "다운로드 중 서버에 요청할 때 전송할 헤더를 설정합니다."
         Height          =   495
         Left            =   840
         TabIndex        =   8
         Top             =   240
         Width           =   4095
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   120
         Picture         =   "pagConnection.frx":000C
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame FrameW3 
      Caption         =   "서버 연결"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6780
      Begin VB.CheckBox chkNoRedirectCheck 
         Caption         =   "리다이렉트 검사 안 함(&R)"
         Height          =   255
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
      Begin VB.CheckBox chkForceGet 
         Caption         =   "파일 검사 시 GET 요청(&Q)"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   480
         Width           =   2865
      End
      Begin VB.CheckBox chkIgnore300 
         Caption         =   "300번대 응답 코드 무시(&I)"
         Height          =   255
         Left            =   3720
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
      Begin prjDownloadBooster.Slider trRequestInterval 
         Height          =   450
         Left            =   3000
         TabIndex        =   5
         Top             =   735
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   794
         Max             =   7
         Value           =   2
         ShowTip         =   0   'False
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   120
         Picture         =   "pagConnection.frx":044E
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "스레드 요청 간격(&N):"
         Height          =   180
         Left            =   1080
         TabIndex        =   4
         Top             =   840
         Width           =   1725
      End
      Begin VB.Label lblIntervalDisplay 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "(0.1초)"
         Height          =   180
         Left            =   5070
         TabIndex        =   6
         Top             =   840
         Width           =   570
      End
   End
End
Attribute VB_Name = "pagConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

