VERSION 5.00
Begin VB.Form frmBatchAdd 
   Caption         =   "일괄 다운로드"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5985
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBatchAdd.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin prjDownloadBooster.CommandButtonW cmdAdvanced 
      Height          =   340
      Left            =   4560
      TabIndex        =   7
      Top             =   900
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      Caption         =   "옵션(&O)..."
   End
   Begin VB.TextBox txtSavePath 
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   3375
      Width           =   4215
   End
   Begin VB.TextBox txtURLs 
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   1
      Top             =   360
      Width           =   4215
   End
   Begin prjDownloadBooster.CommandButtonW cmdCancel 
      Cancel          =   -1  'True
      Height          =   340
      Left            =   4560
      TabIndex        =   6
      Top             =   510
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      Caption         =   "취소"
   End
   Begin prjDownloadBooster.CommandButtonW cmdOK 
      Height          =   340
      Left            =   4560
      TabIndex        =   5
      Top             =   120
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      Caption         =   "확인"
   End
   Begin prjDownloadBooster.CommandButtonW cmdBrowse 
      Height          =   330
      Left            =   4560
      TabIndex        =   4
      Top             =   3360
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      Caption         =   "찾아보기(&B)..."
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "저장 경로(&S):"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3165
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "각 줄에 파일 주소를 입력하십시오(&L)."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmBatchAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If DISABLEFRAMESKIN Then
#Else
Public SkinnedFrame As frmSkinnedFrame
#End If

Dim PrevKeyCode As Integer
Public HeaderCache As String

Implements IBSSubclass

Private Sub cmdAdvanced_Click()
    Tags.DownloadOptionsTargetForm = 1
    Dim DecodedHeaders As Collection
    Set DecodedHeaders = DecodeHeaderCache(HeaderCache)
    Set frmDownloadOptions.HeaderKeys = DecodedHeaders("keys")
    Set frmDownloadOptions.Headers = DecodedHeaders("Values")
    frmDownloadOptions.Show vbModal, Me
End Sub

Private Sub cmdBrowse_Click()
    txtSavePath.Text = Trim$(txtSavePath.Text)
    ShowFileDialog 2, txtSavePath.Text
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    txtSavePath.Text = Trim$(txtSavePath.Text)
    If Not FolderExists(txtSavePath.Text) Then
        MsgBox t("저장 경로가 존재하지 않습니다. [찾아보기] 기능으로 폴더를 찾아볼 수 있습니다.", "Save path does not exist. Use Broewse to browse folders."), 16
        Exit Sub
    End If
    txtSavePath.Text = FilterFilename(txtSavePath.Text, True)

    Dim URLs() As String
    URLs = Split(txtURLs.Text, vbCrLf)
    Dim ErrURLs$
    ErrURLs = ""
    Dim i&, MaxCountErrorShown As Boolean
    MaxCountErrorShown = False
    For i = 0& To UBound(URLs)
        If LenB(Trim$(URLs(i))) Then
            If frmMain.lvBatchFiles.ListItems.Count >= MAX_32BIT_SIGNED_INT Then
                If Not MaxCountErrorShown Then
                    MsgBox t("최대 일괄 다운로드 개수를 초과했습니다.", "Maximum number of items exceeded."), vbExclamation
                    MaxCountErrorShown = True
                End If
                GoTo adderrorurl
            End If
            If Not frmMain.AddBatchURLs(URLs(i), txtSavePath.Text, HeaderCache) Then
adderrorurl:
                ErrURLs = ErrURLs & URLs(i) & vbCrLf
            End If
        End If
    Next i
    
    If LenB(ErrURLs) Then
        txtURLs.Text = ErrURLs
        txtURLs.SelStart = 0
        txtURLs.SelLength = Len(txtURLs.Text)
        On Error Resume Next
        txtURLs.SetFocus
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    txtURLs.SetFocus
End Sub

Private Sub Form_Load()
    InitForm Me
    
    tr Me, "Batch download"
    tr cmdOK, "OK"
    tr cmdCancel, "Cancel"
    tr Label1, "Enter one UR&L per line:"
    tr Label2, "&Save to:"
    tr cmdBrowse, "&Browse..."
    tr cmdAdvanced, "&Options..."
    
    Set cmdOK.ImageList = frmDummyForm.imgOK
    Set cmdCancel.ImageList = frmDummyForm.imgCancel
    
    HeaderCache = ""
    
    On Error Resume Next
    Me.Width = GetSetting("DownloadBooster", "UserData", "BatchURLAddWidth", Me.Width - PaddedBorderWidth * 15 * 2) + PaddedBorderWidth * 15 * 2
    Me.Height = GetSetting("DownloadBooster", "UserData", "BatchURLAddHeight", Me.Height - PaddedBorderWidth * 15 * 2) + PaddedBorderWidth * 15 * 2
    
    AttachMessage Me, Me.hWnd, WM_GETMINMAXINFO
    AttachMessage Me, Me.hWnd, WM_SETTINGCHANGE
End Sub

Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    cmdOK.Left = Me.Width - PaddedBorderWidth * 15 * 2 - 1545
    cmdCancel.Left = cmdOK.Left
    cmdAdvanced.Left = cmdOK.Left
    txtURLs.Width = Me.Width - PaddedBorderWidth * 15 * 2 - 1890
    txtURLs.Height = Me.Height - PaddedBorderWidth * 15 * 2 - 1770
    Label2.Top = Me.Height - PaddedBorderWidth * 15 * 2 - 1140 - 30
    txtSavePath.Top = Me.Height - PaddedBorderWidth * 15 * 2 - 900 - 30
    cmdBrowse.Left = cmdOK.Left
    cmdBrowse.Top = Me.Height - PaddedBorderWidth * 15 * 2 - 945
    txtSavePath.Width = txtURLs.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.WindowState = 0 Then
        SaveSetting "DownloadBooster", "UserData", "BatchURLAddWidth", Me.Width - PaddedBorderWidth * 15 * 2
        SaveSetting "DownloadBooster", "UserData", "BatchURLAddHeight", Me.Height - PaddedBorderWidth * 15 * 2
    End If
    
    IBSSubclass_UnsubclassIt
    
#If DISABLEFRAMESKIN Then
#Else
    Unload SkinnedFrame
#End If
End Sub

Private Function IBSSubclass_MsgResponse(ByVal hWnd As Long, ByVal uMsg As Long) As EMsgResponse
    IBSSubclass_MsgResponse = emrConsume
End Function

Private Sub IBSSubclass_UnsubclassIt()
    DetachMessage Me, Me.hWnd, WM_GETMINMAXINFO
    DetachMessage Me, Me.hWnd, WM_SETTINGCHANGE
End Sub

Private Function IBSSubclass_WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
    On Error Resume Next
 
    Select Case uMsg
        Case WM_GETMINMAXINFO
            Dim lpMMI As MINMAXINFO
            CopyMemory lpMMI, ByVal lParam, Len(lpMMI)
            lpMMI.ptMinTrackSize.X = (5145 + PaddedBorderWidth * 15 * 2) / 15 * (DPI / 96)
            lpMMI.ptMinTrackSize.Y = (2310 + PaddedBorderWidth * 15 * 2) / 15 * (DPI / 96)
            lpMMI.ptMaxTrackSize.X = (Screen.Width + 1200) * (DPI / 96)
            lpMMI.ptMaxTrackSize.Y = (Screen.Height + 1200) * (DPI / 96)
            CopyMemory ByVal lParam, lpMMI, Len(lpMMI)
            
            IBSSubclass_WindowProc = 1&
            Exit Function
        Case WM_SETTINGCHANGE
            Select Case GetStrFromPtr(lParam)
                Case "WindowMetrics"
                    UpdateBorderWidth
                    Form_Resize
            End Select
    End Select
    
    IBSSubclass_WindowProc = CallOldWindowProc(hWnd, uMsg, wParam, lParam)
End Function

Private Sub txtURLs_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 13 Or KeyCode = 10) Then
        If (PrevKeyCode = 13 Or PrevKeyCode = 10) Then cmdOK_Click
        PrevKeyCode = KeyCode
    Else
        PrevKeyCode = 0
    End If
End Sub

Private Sub txtURLs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PrevKeyCode = 0
End Sub
