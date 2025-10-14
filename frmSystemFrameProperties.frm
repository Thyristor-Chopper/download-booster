VERSION 5.00
Begin VB.Form frmSystemFrameProperties 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "스킨 설정"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSystemFrameProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CheckBox chkNoDWM 
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   4215
   End
   Begin VB.CheckBox chkDisableVisualStyle 
      Caption         =   "고전 스타일 제목 표시줄 사용(&C)"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
   Begin prjDownloadBooster.CommandButtonW cmdOK 
      Default         =   -1  'True
      Height          =   345
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      Caption         =   "확인"
   End
   Begin prjDownloadBooster.CommandButtonW cmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   2520
      TabIndex        =   2
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      Caption         =   "취소"
   End
End
Attribute VB_Name = "frmSystemFrameProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#If DISABLEFRAMESKIN Then
#Else
Public SkinnedFrame As frmSkinnedFrame
#End If

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    frmOptions.ClassicFrame = chkDisableVisualStyle.Value
    frmOptions.NoDWMFrame = chkNoDWM.Value
    
    frmOptions.VisualStyleChanged = True
    frmOptions.cmdApply.Enabled = True
    If chkDisableVisualStyle.Value = 0 Then
        ActivateVisualStyles frmOptions.pbBackground.hWnd
    Else
        RemoveVisualStyles frmOptions.pbBackground.hWnd
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    InitForm Me
    
    chkDisableVisualStyle.Value = frmOptions.ClassicFrame
    chkNoDWM.Value = frmOptions.NoDWMFrame
    If WinVer < 6! Then
        cmdOK.Top = cmdOK.Top - 240
        cmdCancel.Top = cmdCancel.Top - 240
        Me.Height = Me.Height - 240
        chkNoDWM.Visible = False
    End If
    
    Set cmdOK.ImageList = frmDummyForm.imgOK
    Set cmdCancel.ImageList = frmDummyForm.imgCancel
    
    tr Me, "Skin Settings"
    tr chkDisableVisualStyle, "Use &classic style title bar"
    If WinVer >= 6.2 Then
        chkNoDWM.Caption = t("Windows 7 기본 스타일 사용(&N)", "Use Wi&ndows 7 Basic Style")
    Else
        chkNoDWM.Caption = t("항상 Windows Aero 비활성화(&N)", "Always disable Wi&ndows Aero")
    End If
    tr cmdOK, "OK"
    tr cmdCancel, "Cancel"
End Sub

Private Sub Form_Unload(Cancel As Integer)
#If DISABLEFRAMESKIN Then
#Else
    Unload SkinnedFrame
#End If
End Sub
