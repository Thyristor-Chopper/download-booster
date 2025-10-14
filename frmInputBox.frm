VERSION 5.00
Begin VB.Form frmInputBox 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "입력 상자"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInputBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin prjDownloadBooster.CommandButtonW cmdOK 
      Default         =   -1  'True
      Height          =   330
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      Enabled         =   0   'False
      Caption         =   "확인"
   End
   Begin prjDownloadBooster.CommandButtonW cmdCancel 
      Cancel          =   -1  'True
      Height          =   330
      Left            =   2520
      TabIndex        =   3
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      Caption         =   "취소"
   End
   Begin VB.TextBox txtInput 
      Height          =   270
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
   Begin VB.Image imgInformation 
      Height          =   360
      Left            =   90
      Picture         =   "frmInputBox.frx":000C
      Top             =   90
      Width           =   360
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   240
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  '투명
      Caption         =   "입력하십시오."
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If DISABLEFRAMESKIN Then
#Else
Public SkinnedFrame As frmSkinnedFrame
#End If

Public ResultID As String
Public InputBoxObject As frmInputBox

Private Sub cmdCancel_Click()
    If Functions.InputBoxResults Is Nothing Then Set Functions.InputBoxResults = New Collection
    If Exists(Functions.InputBoxResults, ResultID) Then Functions.InputBoxResults.Remove ResultID
    Functions.InputBoxResults.Add "", ResultID
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Functions.InputBoxResults Is Nothing Then Set Functions.InputBoxResults = New Collection
    If Exists(Functions.InputBoxResults, ResultID) Then Functions.InputBoxResults.Remove ResultID
    Functions.InputBoxResults.Add Trim$(txtInput.Text), ResultID
    Unload Me
End Sub

Private Sub Form_Activate()
    txtInput.SelStart = 0
    txtInput.SelLength = Len(txtInput.Text)
End Sub

Private Sub Form_Load()
    InitForm Me
    Set imgIcon.Picture = Train(RandInt(1, 2))
    
    Set cmdOK.ImageList = frmDummyForm.imgOK
    Set cmdCancel.ImageList = frmDummyForm.imgCancel
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not InputBoxObject Is Nothing Then
        Unload InputBoxObject
        Set InputBoxObject = Nothing
    End If
    
#If DISABLEFRAMESKIN Then
#Else
    Unload SkinnedFrame
#End If
End Sub

Private Sub txtInput_Change()
    cmdOK.Enabled = LenB(Trim$(txtInput.Text))
End Sub
