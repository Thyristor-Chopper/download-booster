VERSION 5.00
Begin VB.Form frmSystemSkinProperties 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "스킨 설정"
   ClientHeight    =   1425
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
   Icon            =   "frmClassicSkinProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin prjDownloadBooster.CheckBoxW chkDisableVisualStyle 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      Caption         =   "컨트롤에 고전 스타일 사용(&C)"
   End
   Begin prjDownloadBooster.CommandButtonW cmdOK 
      Default         =   -1  'True
      Height          =   345
      Left            =   960
      TabIndex        =   2
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
      TabIndex        =   3
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      Caption         =   "취소"
   End
   Begin prjDownloadBooster.CheckBoxW chkRoundClassicButtons 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      Caption         =   "둥근 단추 사용(&U)"
   End
End
Attribute VB_Name = "frmSystemSkinProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SkinnedFrame As frmSkinnedFrame

Private Sub chkDisableVisualStyle_Click()
    chkRoundClassicButtons.Enabled = (-chkDisableVisualStyle.Value)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    frmOptions.RoundClassicButtons = chkRoundClassicButtons.Value
    frmOptions.DisableVisualStyle = chkDisableVisualStyle.Value
    
    frmOptions.VisualStyleChanged = True
    frmOptions.cmdApply.Enabled = True
    frmOptions.cmdSample.VisualStyles = (chkDisableVisualStyle.Value = 0)
    frmOptions.cmdSample.RoundButton = chkRoundClassicButtons.Value
    frmOptions.txtSampleClassic.Visible = chkDisableVisualStyle.Value
    frmOptions.pbSampleClassic.Visible = chkDisableVisualStyle.Value
    
    If frmOptions.optUserFore.Value Then
        frmOptions.CheckBoxW1.VisualStyles = False
        frmOptions.FrameW5.VisualStyles = False
    Else
        frmOptions.CheckBoxW1.VisualStyles = (frmOptions.DisableVisualStyle = 0)
        frmOptions.FrameW5.VisualStyles = (frmOptions.DisableVisualStyle = 0)
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    InitForm Me
    
    chkRoundClassicButtons.Value = frmOptions.RoundClassicButtons
    chkDisableVisualStyle.Value = frmOptions.DisableVisualStyle
    chkRoundClassicButtons.Enabled = -chkDisableVisualStyle.Value
    
    chkRoundClassicButtons.Visible = (frmOptions.cbSkin.ListIndex = 0)
    If frmOptions.cbSkin.ListIndex <> 0 Then
        cmdOK.Top = cmdOK.Top - 240
        cmdCancel.Top = cmdCancel.Top - 240
        Me.Height = Me.Height - 240
    End If
    
    Set cmdOK.ImageList = frmDummyForm.imgOK
    Set cmdCancel.ImageList = frmDummyForm.imgCancel
    
    tr Me, "Skin Settings"
    tr chkRoundClassicButtons, "&Use rounded buttons"
    tr chkDisableVisualStyle, "Use &classic style for form controls"
    tr cmdOK, "OK"
    tr cmdCancel, "Cancel"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload SkinnedFrame
End Sub
