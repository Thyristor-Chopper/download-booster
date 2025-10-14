VERSION 5.00
Begin VB.Form frmEditBatch 
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5685
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditBatch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin prjDownloadBooster.CommandButtonW cmdOK 
      Default         =   -1  'True
      Height          =   330
      Left            =   2760
      TabIndex        =   9
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      Caption         =   "확인"
   End
   Begin prjDownloadBooster.CommandButtonW cmdCancel 
      Cancel          =   -1  'True
      Height          =   330
      Left            =   4200
      TabIndex        =   10
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      Caption         =   "취소"
   End
   Begin VB.Frame fInfo 
      Caption         =   " 파일 다운로드 정보 "
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin prjDownloadBooster.CommandButtonW cmdHeaders 
         Height          =   330
         Left            =   2160
         TabIndex        =   8
         Top             =   1920
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         Caption         =   "헤더(&H)..."
      End
      Begin prjDownloadBooster.CommandButtonW cmdYtdl 
         Height          =   330
         Left            =   360
         TabIndex        =   7
         Top             =   1920
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         Caption         =   "&youtube-dl..."
      End
      Begin prjDownloadBooster.CommandButtonW cmdBrowse 
         Height          =   330
         Left            =   3720
         TabIndex        =   5
         Top             =   1380
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         Caption         =   "찾아보기(&B)..."
      End
      Begin VB.TextBox txtFilePath 
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Width           =   4935
      End
      Begin VB.TextBox txtURL 
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   4935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "다운로드 옵션(&D):"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "저장 경로(&S):"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "파일 주소(&A):"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If DISABLEFRAMESKIN Then
#Else
Public SkinnedFrame As frmSkinnedFrame
#End If

Public OriginalURL As String
Public OriginalPath As String
Public EncodedHeaders As String
Public InitialFileName As String

Private Sub cmdBrowse_Click()
    ShowFileDialog 1, Trim$(txtFilePath.Text)
    
    If FolderExists(txtFilePath.Text) Then
        If Right$(txtFilePath.Text, 1) <> "\" Then txtFilePath.Text = txtFilePath.Text & "\"
        txtFilePath.Text = txtFilePath.Text & Tags.FileNameOnly
    End If
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHeaders_Click()
    Tags.DownloadOptionsTargetForm = 2
    Dim DecodedHeaders As Collection
    Set DecodedHeaders = DecodeHeaderCache(EncodedHeaders)
    Set frmDownloadOptions.HeaderKeys = DecodedHeaders("keys")
    Set frmDownloadOptions.Headers = DecodedHeaders("Values")
#If HIDEYTDL Then
#Else
    frmDownloadOptions.tsTabStrip.Tabs(2).Selected = True
#End If
    frmDownloadOptions.Show vbModal, Me
End Sub

Private Sub cmdOK_Click()
    txtURL.Text = Trim$(txtURL.Text)
    If Left$(txtURL.Text, 7) <> "http://" And Left$(txtURL.Text, 8) <> "https://" Then
        MsgBox t("주소가 올바르지 않습니다. 'http://' 또는 'https://'로 시작해야 합니다.", "Invalid address. Must start with 'http://' or 'https://'."), 16
        Exit Sub
    End If

    txtFilePath.Text = Trim$(txtFilePath.Text)
    Do While Replace(txtFilePath.Text, "\\", "\") <> txtFilePath.Text
        txtFilePath.Text = Replace(txtFilePath.Text, "\\", "\")
    Loop
    
    Dim OnlyFolderInPath As Boolean, FileNameModified As Boolean
    OnlyFolderInPath = False
    FileNameModified = True
    If FolderExists(txtFilePath.Text) Then
        If Right$(txtFilePath.Text, 1) <> "\" Then txtFilePath.Text = txtFilePath.Text & "\"
        txtFilePath.Text = txtFilePath.Text & Trim$(Tags.FileNameOnly)
        OnlyFolderInPath = True
    ElseIf Right$(txtFilePath.Text, 1) = "\" Or (Not FolderExists(GetParentFolderName(txtFilePath.Text))) Then
        MsgBox t("저장 경로가 존재하지 않습니다. [찾아보기] 기능으로 폴더를 찾아볼 수 있습니다.", "Save path does not exist. Use Broewse to browse folders."), 16
        Exit Sub
    ElseIf LCase(GetFilename(txtFilePath.Text)) = LCase(Tags.FileNameOnly) Then
        FileNameModified = False
    End If
    txtFilePath.Text = FilterFilename(txtFilePath.Text, True)
    
    On Error Resume Next
    Dim ParentFolderName As String
    ParentFolderName = GetParentFolderName(txtFilePath.Text)
    If Right$(ParentFolderName, 1) = "\" Then ParentFolderName = Left$(ParentFolderName, Len(ParentFolderName) - 1)
    frmMain.lvBatchFiles.SelectedItem.ListSubItems(2).Text = txtURL.Text
    If frmMain.lvBatchFiles.SelectedItem.ListSubItems(3).Text = t("완료", "Done") Then
        If txtURL.Text <> Trim$(OriginalURL) Then
            frmMain.lvBatchFiles.SelectedItem.ListSubItems(3).Text = t("대기", "Queued")
            frmMain.lvBatchFiles.SelectedItem.Checked = True
            frmMain.lvBatchFiles.SelectedItem.ForeColor = &H80000008
            frmMain.lvBatchFiles.SelectedItem.ListSubItems(1).ForeColor = &H80000008
            frmMain.lvBatchFiles.SelectedItem.ListSubItems(2).ForeColor = &H80000008
            frmMain.lvBatchFiles.SelectedItem.ListSubItems(3).ForeColor = &H80000008
            GoTo changeFilepath
        ElseIf txtFilePath.Text <> Trim$(OriginalPath) Then
            MsgBox t("다운로드가 이미 완료된 파일의 저장 경로가 수정되지 않았습니다.", "Save path has not been changed because it's already downloaded."), 64
        End If
    Else
changeFilepath:
        frmMain.lvBatchFiles.SelectedItem.Text = Replace(txtFilePath.Text, ParentFolderName & "\", "", 1, 1)
        frmMain.lvBatchFiles.SelectedItem.ListSubItems(1).Text = txtFilePath.Text
        If txtFilePath.Text <> Trim$(OriginalPath) And (Not OnlyFolderInPath) And FileNameModified Then
            frmMain.lvBatchFiles.SelectedItem.ListSubItems(4).Text = "N"
        End If
    End If
    frmMain.lvBatchFiles.SelectedItem.ListSubItems(5).Text = EncodedHeaders
    
    Unload Me
End Sub

#If HIDEYTDL Then
#Else
Private Sub cmdYtdl_Click()
    Tags.DownloadOptionsTargetForm = 2
    Dim DecodedHeaders As Collection
    Set DecodedHeaders = DecodeHeaderCache(EncodedHeaders)
    Set frmDownloadOptions.HeaderKeys = DecodedHeaders("keys")
    Set frmDownloadOptions.Headers = DecodedHeaders("Values")
    frmDownloadOptions.Show vbModal, Me
End Sub
#End If

Private Sub Form_Activate()
    On Error Resume Next
    txtURL.SetFocus
End Sub

Private Sub Form_Load()
    InitForm Me
    
    tr cmdOK, "OK"
    tr cmdCancel, "Cancel"
    tr cmdBrowse, "&Browse..."
    tr fInfo, " File download information "
    tr Label1, "File &address:"
    tr Label2, "&Save to:"
    tr Label3, "&Download options:"
    tr cmdHeaders, "&Headers..."
    
    EncodedHeaders = frmMain.lvBatchFiles.SelectedItem.ListSubItems(5).Text
    InitialFileName = frmMain.lvBatchFiles.SelectedItem.Text
    Me.Caption = t("편집", "Edit") & " - " & InitialFileName
    
    Set cmdOK.ImageList = frmDummyForm.imgOK
    Set cmdCancel.ImageList = frmDummyForm.imgCancel
    
#If HIDEYTDL Then
    cmdYtdl.Visible = False
    cmdHeaders.Left = cmdYtdl.Left
#End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
#If DISABLEFRAMESKIN Then
#Else
    Unload SkinnedFrame
#End If
End Sub
