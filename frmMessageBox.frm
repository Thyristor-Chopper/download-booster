VERSION 5.00
Begin VB.Form frmMessageBox 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "메시지 상자"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   28440
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMessageBox.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   28440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdTryAgain 
      Caption         =   "다시 시도(&T)"
      Height          =   315
      Left            =   15240
      TabIndex        =   11
      Top             =   840
      Width           =   1455
   End
   Begin VB.Timer timeout 
      Enabled         =   0   'False
      Left            =   360
      Top             =   960
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "취소"
      Height          =   320
      Left            =   5880
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "아니요(&N)"
      Height          =   320
      Left            =   4320
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "예(&Y)"
      Height          =   320
      Left            =   2760
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin VB.OptionButton optNo 
      Caption         =   "아니요(&N)"
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   1320
      Width           =   1575
   End
   Begin VB.OptionButton optYes 
      Caption         =   "예(&Y)"
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "확인"
      Height          =   315
      Left            =   7440
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdRetry 
      Caption         =   "다시 시도(&R)"
      Height          =   315
      Left            =   9000
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "중단(&A)"
      Height          =   315
      Left            =   10560
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdIgnore 
      Caption         =   "무시(&I)"
      Height          =   315
      Left            =   12120
      TabIndex        =   6
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "계속(&C)"
      Height          =   315
      Left            =   13680
      TabIndex        =   12
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "도움말"
      Height          =   315
      Left            =   16800
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.Image imgIcon 
      Height          =   360
      Index           =   1
      Left            =   75
      Picture         =   "frmMessageBox.frx":000C
      Stretch         =   -1  'True
      Top             =   90
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgIcon 
      Height          =   360
      Index           =   3
      Left            =   75
      Picture         =   "frmMessageBox.frx":00C1
      Stretch         =   -1  'True
      Top             =   90
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgIcon 
      Height          =   360
      Index           =   2
      Left            =   75
      Picture         =   "frmMessageBox.frx":01DC
      Stretch         =   -1  'True
      Top             =   90
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgIcon 
      Height          =   360
      Index           =   4
      Left            =   75
      Picture         =   "frmMessageBox.frx":029A
      Stretch         =   -1  'True
      Top             =   90
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblContent 
      BackColor       =   &H00F8EFE5&
      BackStyle       =   0  '투명
      Caption         =   "내용"
      Height          =   495
      Left            =   960
      TabIndex        =   10
      Top             =   360
      Width           =   27255
   End
   Begin VB.Image imgTrain 
      Height          =   480
      Left            =   255
      Stretch         =   -1  'True
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmMessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public MsgBoxMode As Byte
Public MsgBoxResult As VbMsgBoxResultEx
Public ResultID As String
Public MessageBoxObject As frmMessageBox

#If False Then
Private Sub cmdAbort_Click()
    MsgBoxResult = vbAbort
    Unload Me
End Sub

Private Sub cmdContinue_Click()
    MsgBoxResult = vbContinue
    Unload Me
End Sub

Private Sub cmdIgnore_Click()
    MsgBoxResult = vbIgnore
    Unload Me
End Sub

Private Sub cmdTryAgain_Click()
    MsgBoxResult = vbTryAgain
    Unload Me
End Sub
#End If

Private Sub cmdNo_Click()
    MsgBoxResult = vbNo
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If MsgBoxMode = vbYesNoEx Then
        If optYes.Value = True Then
            MsgBoxResult = vbYes
        Else
            MsgBoxResult = vbNo
        End If
    Else
        MsgBoxResult = vbOK
    End If
    Unload Me
End Sub

Private Sub cmdRetry_Click()
    MsgBoxResult = vbRetry
    Unload Me
End Sub

Private Sub cmdYes_Click()
    MsgBoxResult = vbYes
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    MsgBoxResult = vbCancel
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Select Case MsgBoxMode
        Case vbOKOnly
            cmdOK.SetFocus
        Case vbYesNo
            cmdYes.SetFocus
        Case vbYesNoEx
            optNo.SetFocus
        Case vbYesNoCancel
            cmdCancel.SetFocus
        Case vbRetryCancel
            cmdRetry.SetFocus
#If False Then
        Case vbAbortRetryIgnore
            cmdAbort.SetFocus
        Case vbOKCancel
            cmdOK.SetFocus
        Case vbCancelTryContinue
            cmdCancel.SetFocus
#End If
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case 78 'N
            If optNo.Visible Then
                optNo.Value = True
                optNo_Click
                optNo.SetFocus
            ElseIf cmdNo.Visible Then
                cmdNo_Click
            End If
        Case 89 'Y
            If optYes.Visible Then
                optYes.Value = True
                optYes_Click
                optYes.SetFocus
            ElseIf cmdYes.Visible Then
                cmdYes_Click
            End If
        Case 82 'R
            If cmdRetry.Visible Then cmdRetry_Click
#If False Then
        Case 65 'A
            If cmdAbort.Visible Then cmdAbort_Click
        Case 73 'I
            If cmdIgnore.Visible Then cmdIgnore_Click
        Case 67 'C
            If cmdContinue.Visible Then cmdContinue_Click
        Case 84 'T
            If cmdTryAgain.Visible Then cmdTryAgain_Click
#End If
    End Select
End Sub

Private Sub Form_Load()
    InitForm Me
End Sub

Sub Init()
    Dim SystemMenu As Long
    SystemMenu = GetSystemMenu(Me.hWnd, 0&)
    DeleteMenu SystemMenu, 0&, MF_BYCOMMAND
    If MsgBoxMode = vbYesNo Or MsgBoxMode = vbAbortRetryIgnore Then
        DeleteMenu SystemMenu, SC_CLOSE, MF_BYCOMMAND
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then
        If MsgBoxMode = vbYesNo Or MsgBoxMode = vbAbortRetryIgnore Then
            Cancel = 1
            Exit Sub
        Else
            MsgBoxResult = vbCancel
        End If
    End If
    GetSystemMenu Me.hWnd, 1&
    If MsgBoxMode <> vbOKOnly Then
        If Exists(Functions.MsgBoxResults, ResultID) Then Functions.MsgBoxResults.Remove ResultID
        Functions.MsgBoxResults.Add MsgBoxResult, ResultID
    End If
    
    If Not MessageBoxObject Is Nothing Then
        Unload MessageBoxObject
        Set MessageBoxObject = Nothing
    End If
End Sub

Private Sub timeout_Timer()
    cmdOK_Click
End Sub

Private Sub optNo_Click()
    cmdOK.Enabled = True
End Sub

Private Sub optYes_Click()
    cmdOK.Enabled = True
End Sub
