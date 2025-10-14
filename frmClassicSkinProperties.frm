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
   Begin VB.CheckBox chkDisableVisualStyle 
      Caption         =   "컨트롤에 고전 스타일 사용(&C)"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4215
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
   Begin VB.CheckBox chkRoundClassicButtons 
      Caption         =   "둥근 단추 사용(&U)"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   4215
   End
End
Attribute VB_Name = "frmSystemSkinProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#If DISABLEFRAMESKIN Then
#Else
Public SkinnedFrame As frmSkinnedFrame
#End If

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
    
    If frmOptions.DisableVisualStyle <> 0 Or frmOptions.optUserFore.Value = True Then
        RemoveVisualStyles frmOptions.CheckBoxW1.hWnd
        RemoveVisualStyles frmOptions.FrameW5.hWnd
    Else
        ActivateVisualStyles frmOptions.CheckBoxW1.hWnd
        ActivateVisualStyles frmOptions.FrameW5.hWnd
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
#If DISABLEFRAMESKIN Then
#Else
    Unload SkinnedFrame
#End If
End Sub
