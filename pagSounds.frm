VERSION 5.00
Begin VB.Form pagSounds 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "효과음"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7020
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "pagSounds.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame FrameW9 
      Caption         =   "메시지 상자"
      Height          =   3135
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   6810
      Begin VB.CheckBox chkQuestion 
         Caption         =   "질문(&Q)"
         Height          =   255
         Left            =   720
         TabIndex        =   24
         Top             =   240
         Value           =   1  '확인
         Width           =   1815
      End
      Begin VB.TextBox txtQuestion 
         Height          =   300
         Left            =   960
         TabIndex        =   23
         Top             =   570
         Width           =   3975
      End
      Begin VB.CheckBox chkError 
         Caption         =   "오류 메시지(&R)"
         Height          =   255
         Left            =   720
         TabIndex        =   19
         Top             =   960
         Value           =   1  '확인
         Width           =   1815
      End
      Begin VB.TextBox txtError 
         Height          =   300
         Left            =   960
         TabIndex        =   18
         Top             =   1290
         Width           =   3975
      End
      Begin VB.CheckBox chkExclamation 
         Caption         =   "경고 메시지(&E)"
         Height          =   255
         Left            =   720
         TabIndex        =   14
         Top             =   1680
         Value           =   1  '확인
         Width           =   3135
      End
      Begin VB.TextBox txtExclamation 
         Height          =   300
         Left            =   960
         TabIndex        =   13
         Top             =   2010
         Width           =   3975
      End
      Begin VB.CheckBox chkAsterisk 
         Caption         =   "일반 메시지(&A)"
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   2400
         Value           =   1  '확인
         Width           =   1815
      End
      Begin VB.TextBox txtAsterisk 
         Height          =   300
         Left            =   960
         TabIndex        =   8
         Top             =   2730
         Width           =   3975
      End
      Begin prjDownloadBooster.CommandButtonW cmdBrowseAsterisk 
         Height          =   300
         Left            =   5640
         TabIndex        =   7
         Top             =   2730
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         ImageListAlignment=   4
      End
      Begin prjDownloadBooster.CommandButtonW cmdTestAsterisk 
         Height          =   300
         Left            =   6240
         TabIndex        =   10
         Top             =   2730
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   529
         Caption         =   "▶"
      End
      Begin prjDownloadBooster.CommandButtonW cmdResetAsterisk 
         Height          =   300
         Left            =   5040
         TabIndex        =   11
         Top             =   2730
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         ImageListAlignment=   4
      End
      Begin prjDownloadBooster.CommandButtonW cmdBrowseExclamation 
         Height          =   300
         Left            =   5640
         TabIndex        =   12
         Top             =   2010
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         ImageListAlignment=   4
      End
      Begin prjDownloadBooster.CommandButtonW cmdTestExclamation 
         Height          =   300
         Left            =   6240
         TabIndex        =   15
         Top             =   2010
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   529
         Caption         =   "▶"
      End
      Begin prjDownloadBooster.CommandButtonW cmdResetExclamation 
         Height          =   300
         Left            =   5040
         TabIndex        =   16
         Top             =   2010
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         ImageListAlignment=   4
      End
      Begin prjDownloadBooster.CommandButtonW cmdBrowseError 
         Height          =   300
         Left            =   5640
         TabIndex        =   17
         Top             =   1290
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         ImageListAlignment=   4
      End
      Begin prjDownloadBooster.CommandButtonW cmdTestError 
         Height          =   300
         Left            =   6240
         TabIndex        =   20
         Top             =   1290
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   529
         Caption         =   "▶"
      End
      Begin prjDownloadBooster.CommandButtonW cmdResetError 
         Height          =   300
         Left            =   5040
         TabIndex        =   21
         Top             =   1290
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         ImageListAlignment=   4
      End
      Begin prjDownloadBooster.CommandButtonW cmdBrowseQuestion 
         Height          =   300
         Left            =   5640
         TabIndex        =   22
         Top             =   570
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         ImageListAlignment=   4
      End
      Begin prjDownloadBooster.CommandButtonW cmdTestQuestion 
         Height          =   300
         Left            =   6240
         TabIndex        =   25
         Top             =   570
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   529
         Caption         =   "▶"
      End
      Begin prjDownloadBooster.CommandButtonW cmdResetQuestion 
         Height          =   300
         Left            =   5040
         TabIndex        =   26
         Top             =   570
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         ImageListAlignment=   4
      End
      Begin VB.Image Image8 
         Height          =   480
         Left            =   120
         Picture         =   "pagSounds.frx":000C
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame FrameW6 
      Caption         =   "다운로드 알림"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6810
      Begin VB.CheckBox chkBeepWhenComplete 
         Caption         =   "다운로드 완료(&B)"
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   240
         Value           =   1  '확인
         Width           =   2535
      End
      Begin VB.TextBox txtCompleteSoundPath 
         Height          =   300
         Left            =   960
         TabIndex        =   2
         Top             =   570
         Width           =   3975
      End
      Begin prjDownloadBooster.CommandButtonW cmdBrowseCompleteSound 
         Height          =   300
         Left            =   5640
         TabIndex        =   1
         Top             =   570
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         ImageListAlignment=   4
      End
      Begin prjDownloadBooster.CommandButtonW cmdTestCompleteSound 
         Height          =   300
         Left            =   6240
         TabIndex        =   4
         Top             =   570
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   529
         Caption         =   "▶"
      End
      Begin prjDownloadBooster.CommandButtonW cmdResetCompleteSound 
         Height          =   300
         Left            =   5040
         TabIndex        =   5
         Top             =   570
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         ImageListAlignment=   4
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "pagSounds.frx":044E
         Top             =   120
         Width           =   480
      End
   End
End
Attribute VB_Name = "pagSounds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
