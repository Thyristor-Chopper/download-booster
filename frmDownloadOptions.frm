VERSION 5.00
Begin VB.Form frmDownloadOptions 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "다운로드 설정"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11490
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDownloadOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   11490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.PictureBox pbPanel 
      BorderStyle     =   0  '없음
      Enabled         =   0   'False
      Height          =   4095
      Index           =   2
      Left            =   4920
      ScaleHeight     =   4095
      ScaleWidth      =   6405
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   6405
      Begin prjDownloadBooster.CommandButtonW cmdEditHeaderName 
         Height          =   330
         Left            =   3600
         TabIndex        =   20
         Top             =   3660
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         Enabled         =   0   'False
         Caption         =   "이름 변경(&R)"
      End
      Begin VB.TextBox txtEdit 
         Height          =   255
         Left            =   3600
         TabIndex        =   22
         Top             =   720
         Visible         =   0   'False
         Width           =   2535
      End
      Begin prjDownloadBooster.CommandButtonW cmdDeleteHeader 
         Height          =   330
         Left            =   2220
         TabIndex        =   19
         Top             =   3660
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         Enabled         =   0   'False
         Caption         =   "삭제(&D)"
      End
      Begin prjDownloadBooster.CommandButtonW cmdEditHeaderValue 
         Height          =   330
         Left            =   4980
         TabIndex        =   21
         Top             =   3660
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         Enabled         =   0   'False
         Caption         =   "편집(&E)"
      End
      Begin prjDownloadBooster.CommandButtonW cmdAddHeader 
         Height          =   330
         Left            =   840
         TabIndex        =   18
         Top             =   3660
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         Caption         =   "추가(&A)"
      End
      Begin prjDownloadBooster.ListView lvHeaders 
         Height          =   3015
         Left            =   840
         TabIndex        =   17
         Top             =   600
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   5318
         VisualTheme     =   1
         View            =   3
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HideSelection   =   0   'False
         ShowLabelTips   =   -1  'True
         HighlightColumnHeaders=   -1  'True
         AutoSelectFirstItem=   0   'False
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  '투명
         Caption         =   "이 헤더는 이번에만 적용됩니다. 영구적으로 변경하려면 옵션에서 변경하십시오. 이 헤더는 옵션의 헤더보다 우선합니다."
         Height          =   495
         Left            =   840
         TabIndex        =   16
         Top             =   180
         Width           =   5475
      End
      Begin VB.Image imgIcon1 
         Height          =   480
         Left            =   120
         Picture         =   "frmDownloadOptions.frx":000C
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.PictureBox pbPanel 
      BorderStyle     =   0  '없음
      Height          =   3135
      Index           =   1
      Left            =   120
      ScaleHeight     =   3135
      ScaleWidth      =   4695
      TabIndex        =   1
      Top             =   600
      Width           =   4695
      Begin VB.Frame FrameW1 
         BorderStyle     =   0  '없음
         Height          =   3135
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   4695
         Begin VB.CheckBox chkAutoYtdl 
            Caption         =   "youtube-dl 사용 여부 자동 결정(&T)"
            Height          =   255
            Left            =   360
            TabIndex        =   3
            Top             =   120
            Width           =   4455
         End
         Begin VB.OptionButton optDisableYtdl 
            Caption         =   "youtube-dl 사용 안 함(&D)"
            Height          =   255
            Left            =   360
            TabIndex        =   4
            Top             =   480
            Value           =   -1  'True
            Width           =   4335
         End
         Begin VB.OptionButton optUseYtdl 
            Caption         =   "youtube-dl 사용(&U)"
            Height          =   255
            Left            =   360
            TabIndex        =   5
            Top             =   840
            Width           =   1935
         End
         Begin VB.Frame fYtdl 
            Caption         =   "        "
            Height          =   2175
            Left            =   120
            TabIndex        =   6
            Top             =   870
            Width           =   4455
            Begin VB.ComboBox txtFormat 
               Height          =   300
               Left            =   1200
               TabIndex        =   8
               Top             =   360
               Width           =   2895
            End
            Begin VB.ComboBox cbBitRate 
               Height          =   300
               Left            =   1560
               Style           =   2  '드롭다운 목록
               TabIndex        =   15
               Top             =   1800
               Width           =   1215
            End
            Begin VB.OptionButton optCBR 
               Caption         =   "&CBR:"
               Height          =   255
               Left            =   840
               TabIndex        =   14
               Top             =   1800
               Width           =   735
            End
            Begin VB.ComboBox cbAudioFormat 
               Height          =   300
               Left            =   2040
               Style           =   2  '드롭다운 목록
               TabIndex        =   11
               Top             =   1080
               Width           =   2055
            End
            Begin VB.ComboBox cbVBR 
               Height          =   300
               Left            =   1560
               Style           =   2  '드롭다운 목록
               TabIndex        =   13
               Top             =   1440
               Width           =   1215
            End
            Begin VB.OptionButton optVBR 
               Caption         =   "&VBR:"
               Height          =   255
               Left            =   840
               TabIndex        =   12
               Top             =   1440
               Value           =   -1  'True
               Width           =   735
            End
            Begin VB.CheckBox chkExtractAudio 
               Caption         =   "음원만 추출(&E)"
               Height          =   255
               Left            =   360
               TabIndex        =   9
               Top             =   720
               Width           =   3015
            End
            Begin VB.Label Label4 
               BackStyle       =   0  '투명
               Caption         =   "오디오 형식(&A):"
               Height          =   255
               Left            =   600
               TabIndex        =   10
               Top             =   1125
               Width           =   1455
            End
            Begin VB.Label Label3 
               BackStyle       =   0  '투명
               Caption         =   "포맷(&F):"
               Height          =   255
               Left            =   360
               TabIndex        =   7
               Top             =   390
               Width           =   855
            End
         End
      End
   End
   Begin prjDownloadBooster.TabStrip tsTabStrip 
      Height          =   390
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   688
      TabMinWidth     =   48
      DoubleBuffer    =   0   'False
      InitTabs        =   "frmDownloadOptions.frx":044E
   End
   Begin prjDownloadBooster.CommandButtonW CancelButton 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   4320
      TabIndex        =   24
      Top             =   120
      Width           =   1320
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "취소"
   End
   Begin prjDownloadBooster.CommandButtonW OKButton 
      Default         =   -1  'True
      Height          =   360
      Left            =   2880
      TabIndex        =   23
      Top             =   120
      Width           =   1320
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "확인"
   End
End
Attribute VB_Name = "frmDownloadOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If DISABLEFRAMESKIN Then
#Else
Public SkinnedFrame As frmSkinnedFrame
#End If

Dim SelectedListItem As LvwListItem
Public Headers As Collection
Public HeaderKeys As Collection

'Implements IBSSubclass

Private Sub Form_Unload(Cancel As Integer)
    'IBSSubclass_UnsubclassIt
    
#If DISABLEFRAMESKIN Then
#Else
    Unload SkinnedFrame
#End If
End Sub

'Private Function IBSSubclass_MsgResponse(ByVal hWnd As Long, ByVal uMsg As Long) As EMsgResponse
'    IBSSubclass_MsgResponse = emrConsume
'End Function
'
'Private Sub IBSSubclass_UnsubclassIt()
'    DetachMessage Me, Me.hWnd, WM_SETTINGCHANGE
'    DetachMessage Me, Me.hWnd, WM_THEMECHANGED
'End Sub
'
'Private Function IBSSubclass_WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
'    On Error Resume Next
'
'    Select Case uMsg
'        Case WM_SETTINGCHANGE
'            Select Case GetStrFromPtr(lParam)
'                Case "WindowMetrics"
'                    DrawTabBackground Me, tsTabStrip, pbPanel
'            End Select
'        Case WM_THEMECHANGED
'            DrawTabBackground Me, tsTabStrip, pbPanel
'    End Select
'
'    IBSSubclass_WindowProc = CallOldWindowProc(hWnd, uMsg, wParam, lParam)
'End Function

Private Sub CancelButton_Click()
    Unload Me
End Sub

#If HIDEYTDL Then
#Else
Private Sub cbAudioFormat_Click()
    optUseYtdl_Click
End Sub

Private Sub chkAutoYtdl_Click()
    optDisableYtdl.Enabled = (chkAutoYtdl.Value = 0)
    optUseYtdl.Enabled = optDisableYtdl.Enabled
    optUseYtdl_Click
End Sub

Private Sub chkExtractAudio_Click()
    optUseYtdl_Click
End Sub
#End If

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 And IsKeyPressed(gksKeyboardctrl) Then
        NextTabPage tsTabStrip, IsKeyPressed(gksKeyboardShift)
    End If
End Sub

Private Sub OKButton_Click()
    On Error Resume Next
    Dim i&
    
    If Tags.DownloadOptionsTargetForm = 3 Then
        SaveSetting "DownloadBooster", "Options\Headers", "_Dummy_", "x" '오류 방지
        DeleteSetting "DownloadBooster", "Options\Headers"
        For i = 1 To lvHeaders.ListItems.Count
            If LenB(Trim$(lvHeaders.ListItems(i).Text)) Then SaveSetting "DownloadBooster", "Options\Headers", Trim$(lvHeaders.ListItems(i).Text), lvHeaders.ListItems(i).ListSubItems(1).Text
        Next i
        BuildHeaderCache
    Else
        For i = 1& To Headers.Count
            Headers.Remove 1
        Next i
        For i = 1& To HeaderKeys.Count
            HeaderKeys.Remove 1
        Next i
        
        Dim HeaderCache$
        
        If lvHeaders.ListItems.Count Then
            Dim RawHeaders$
            RawHeaders = ""
            For i = 1& To lvHeaders.ListItems.Count
                If LenB(Trim$(lvHeaders.ListItems(i).Text)) And (Not Exists(Headers, CStr(Trim$(lvHeaders.ListItems(i).Text)))) Then
                    Headers.Add CStr(lvHeaders.ListItems(i).ListSubItems(1).Text), CStr(Trim$(lvHeaders.ListItems(i).Text))
                    HeaderKeys.Add CStr(Trim$(lvHeaders.ListItems(i).Text))
                    RawHeaders = RawHeaders & LCase(Trim$(lvHeaders.ListItems(i).Text)) & ": " & lvHeaders.ListItems(i).ListSubItems(1).Text & vbLf
                End If
            Next i
            If Right$(RawHeaders, 1) = vbLf Then RawHeaders = Left$(RawHeaders, Len(RawHeaders) - 1)
            HeaderCache = btoa(RawHeaders)
        Else
            HeaderCache = ""
        End If
        
        Select Case Tags.DownloadOptionsTargetForm
            Case 0
                SessionHeaderCache = HeaderCache
            Case 1
                frmBatchAdd.HeaderCache = HeaderCache
            Case 2
                frmEditBatch.EncodedHeaders = HeaderCache
        End Select
    End If
    
#If HIDEYTDL Then
#Else
    SaveSetting "DownloadBooster", "Options", "AutoDetectYtdlURL", chkAutoYtdl.Value

    frmMain.ytdlEnabled = optUseYtdl.Value
    If optUseYtdl.Value Or chkAutoYtdl.Value Then
        frmMain.ytdlFormat = Replace(txtFormat.Text, " ", "")
        If frmMain.ytdlFormat = txtFormat.List(0) Then frmMain.ytdlFormat = ""
        frmMain.ytdlExtractAudio = (chkExtractAudio.Value = 1)
        frmMain.ytdlAudioFormat = cbAudioFormat.ListIndex
        frmMain.ytdlAudioBitrateType = IIf(optVBR.Value, AudioBitrateType.VBR, AudioBitrateType.CBR)
        frmMain.ytdlAudioVBR = CByte(cbVBR.Text)
        frmMain.ytdlAudioCBR = CInt(Left$(cbBitRate.Text, InStr(cbBitRate.Text, " ") - 1))
    End If
#End If
    
    Unload Me
End Sub

#If HIDEYTDL Then
#Else
Private Sub optDisableYtdl_Click()
    optUseYtdl_Click
End Sub

Private Sub optUseYtdl_Click()
    Dim ctrl As Control
    On Error Resume Next
    For Each ctrl In fYtdl.ContainedControls
        ctrl.Enabled = (optUseYtdl.Value Or chkAutoYtdl.Value = 1)
    Next ctrl
    
    If optUseYtdl.Value Or chkAutoYtdl.Value = 1 Then
        Label4.Enabled = (chkExtractAudio.Value = 1)
        cbAudioFormat.Enabled = (chkExtractAudio.Value = 1)
        optVBR.Enabled = (chkExtractAudio.Value = 1) And cbAudioFormat.ListIndex = 1
        optCBR.Enabled = (chkExtractAudio.Value = 1) And cbAudioFormat.ListIndex = 1
        cbVBR.Enabled = (chkExtractAudio.Value = 1) And cbAudioFormat.ListIndex = 1
        cbBitRate.Enabled = (chkExtractAudio.Value = 1) And cbAudioFormat.ListIndex = 1
        If chkExtractAudio.Value = 1 And cbAudioFormat.ListIndex = 1 Then
            cbVBR.Enabled = optVBR.Value
            cbBitRate.Enabled = optCBR.Value
        End If
    End If
End Sub

Private Sub optCBR_Click()
    optUseYtdl_Click
End Sub

Private Sub optVBR_Click()
    optUseYtdl_Click
End Sub
#End If

Private Sub Form_Load()
    InitForm Me
    
    Dim i As Byte
    InitPropertySheetDimensions Me, tsTabStrip, pbPanel, OKButton, CancelButton
'#If HIDEYTDL Then
'    DrawTabBackground Me, tsTabStrip, pbPanel, False
'#Else
'    DrawTabBackground Me, tsTabStrip, pbPanel
'#End If
    
    On Error Resume Next
    
    Set lvHeaders.SmallIcons = frmDummyForm.imgFiles
    
    Set OKButton.ImageList = frmDummyForm.imgOK
    Set CancelButton.ImageList = frmDummyForm.imgCancel

#If HIDEYTDL Then
#Else
    'FrameW1.Refresh
    'fYtdl.Refresh
    
    AddItemToComboBox cbAudioFormat, t("자동", "Auto") & " (M4A/Opus)"
    AddItemToComboBox cbAudioFormat, "MP3"
    AddItemToComboBox cbAudioFormat, "WAV"
    AddItemToComboBox cbAudioFormat, "FLAC"
    cbAudioFormat.ListIndex = 0
    
    For i = 0 To 9
         AddItemToComboBox cbVBR, i
    Next i
    cbVBR.ListIndex = 0
    
    AddItemToComboBox cbBitRate, "8 kbps"
    AddItemToComboBox cbBitRate, "16 kbps"
    AddItemToComboBox cbBitRate, "24 kbps"
    AddItemToComboBox cbBitRate, "32 kbps"
    AddItemToComboBox cbBitRate, "40 kbps"
    AddItemToComboBox cbBitRate, "48 kbps"
    AddItemToComboBox cbBitRate, "56 kbps"
    AddItemToComboBox cbBitRate, "64 kbps"
    AddItemToComboBox cbBitRate, "80 kbps"
    AddItemToComboBox cbBitRate, "96 kbps"
    AddItemToComboBox cbBitRate, "112 kbps"
    AddItemToComboBox cbBitRate, "128 kbps"
    AddItemToComboBox cbBitRate, "144 kbps"
    AddItemToComboBox cbBitRate, "160 kbps"
    AddItemToComboBox cbBitRate, "192 kbps"
    AddItemToComboBox cbBitRate, "224 kbps"
    AddItemToComboBox cbBitRate, "256 kbps"
    AddItemToComboBox cbBitRate, "320 kbps"
    cbBitRate.ListIndex = 14
    
    AddItemToComboBox txtFormat, t("자동", "Auto")
    AddItemToComboBox txtFormat, "233"
    AddItemToComboBox txtFormat, "234"
    AddItemToComboBox txtFormat, "256"
    AddItemToComboBox txtFormat, "140"
    AddItemToComboBox txtFormat, "251"
    AddItemToComboBox txtFormat, "380"
    AddItemToComboBox txtFormat, "328"
    AddItemToComboBox txtFormat, "258"
    AddItemToComboBox txtFormat, "269"
    AddItemToComboBox txtFormat, "160"
    AddItemToComboBox txtFormat, "230"
    AddItemToComboBox txtFormat, "134"
    AddItemToComboBox txtFormat, "18"
    AddItemToComboBox txtFormat, "605"
    AddItemToComboBox txtFormat, "243"
    AddItemToComboBox txtFormat, "231"
    AddItemToComboBox txtFormat, "135"
    AddItemToComboBox txtFormat, "hls_mp3_0_0"
    AddItemToComboBox txtFormat, "http_mp3_0_0"
    AddItemToComboBox txtFormat, "hls_opus_0_0"
    txtFormat.ListIndex = 0
    
    chkAutoYtdl.Value = GetSetting("DownloadBooster", "Options", "AutoDetectYtdlURL", 1)

    tr chkAutoYtdl, "Automatically use &youtube-dl for supported links"
    tr optDisableYtdl, "Never use youtube-&dl for all links"
    tr optUseYtdl, "Always &use youtube-dl for all links"
    optUseYtdl.Width = t(optUseYtdl.Width, 2775)
    tr chkExtractAudio, "&Extract audio"
    tr Label4, "&Audio format:"
    tr Label3, "&Format:"
#End If
    
    tr tsTabStrip.Tabs(2), "Headers"
    
    tr OKButton, "OK"
    tr CancelButton, "Cancel"
    
    tr cmdAddHeader, "&Add"
    tr cmdDeleteHeader, "&Delete"
    tr cmdEditHeaderName, "&Rename"
    tr cmdEditHeaderValue, "&Edit"
    
    lvHeaders.ColumnHeaders.Add , , t("이름", "Name"), 2055
    lvHeaders.ColumnHeaders.Add , , t("값", "Value"), 2775
    
    LoadSettings
    
#If HIDEYTDL Then
    RemoveYtdlTab
#End If

#If HIDEYTDL Then
#Else
    optUseYtdl_Click
#End If
    
'    AttachMessage Me, Me.hWnd, WM_SETTINGCHANGE
'    AttachMessage Me, Me.hWnd, WM_THEMECHANGED
End Sub

Sub RemoveYtdlTab()
    tsTabStrip.Tabs(2).Selected = True
    tsTabStrip.Tabs.Remove 1
    pbPanel(2).Visible = True
    pbPanel(1).Visible = False
    pbPanel(2).Enabled = True
End Sub

Private Sub LoadSettings()
    'SetWindowPos Me.hWnd, IIf(MainFormOnTop, hWnd_TOPMOST, hWnd_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    Dim i&
    
    Me.Caption = t("다운로드 설정", "Download settings")
    Select Case Tags.DownloadOptionsTargetForm
        Case 0
            lblDescription.Caption = t(lblDescription.Caption, "Headers here are only applied in this session. Go to options to change them permanently.")
        Case 1
            Me.Caption = Me.Caption & " - " & t("일괄 다운로드", "Batch Download")
            lblDescription.Caption = t("일괄 다운로드할 파일들에 접속할 때 요청할 헤더를 지정하십시오.", "Specify the headers for this batch download.")
        Case 2
            Me.Caption = Me.Caption & " - " & frmEditBatch.InitialFileName
            lblDescription.Caption = t("이 파일 다운로드 시에 요청할 헤더를 지정하십시오.", "Specify the headers when requesting this file to download.")
        Case 3
            Me.Caption = t("사용자 지정 헤더", "Headers")
            lblDescription.Caption = t("다운로드 중 서버에 요청할 때 전송할 헤더를 설정합니다. [다운로드 설정]에서 설정한 헤더가 우선적으로 적용됩니다.", "Set the headers when requesting to the server on download. Headers set in Download Options have higher priority.")
            
            Dim GlobalHeaders() As String
            GlobalHeaders = GetAllSettings("DownloadBooster", "Options\Headers")
            For i = LBound(GlobalHeaders) To UBound(GlobalHeaders)
                lvHeaders.ListItems.Add(, , GlobalHeaders(i, 0), , 1).ListSubItems.Add , , GlobalHeaders(i, 1)
            Next i
            
            GoTo headerloadcomplete
    End Select
    
    Dim Header
    For Each Header In HeaderKeys
        lvHeaders.ListItems.Add(, , Header, , 1).ListSubItems.Add , , Headers(CStr(Header))
    Next Header
    
headerloadcomplete:

#If HIDEYTDL Then
#Else
    optUseYtdl.Value = frmMain.ytdlEnabled
    txtFormat.Text = Replace(frmMain.ytdlFormat, " ", "")
    If LenB(txtFormat.Text) = 0 Then txtFormat.ListIndex = 0
    chkExtractAudio.Value = -(frmMain.ytdlExtractAudio)
    cbAudioFormat.ListIndex = frmMain.ytdlAudioFormat
    IIf(frmMain.ytdlAudioBitrateType = CBR, optCBR, optVBR).Value = True
    cbVBR.ListIndex = frmMain.ytdlAudioVBR
    For i = 0 To cbBitRate.ListCount - 1
        If CInt(Left$(cbBitRate.List(i), InStr(cbBitRate.List(i), " ") - 1)) = frmMain.ytdlAudioCBR Then
            cbBitRate.ListIndex = i
            Exit For
        End If
    Next i
#End If
End Sub

#If HIDEYTDL Then
#Else
Private Sub tsTabStrip_TabClick(TabItem As TbsTab)
    On Error Resume Next
    Static i As Byte, Show As Boolean
    For i = 1 To pbPanel.Count
        Show = (i = TabItem.Index)
        pbPanel(i).Visible = Show
        pbPanel(i).Enabled = Show
    Next i
End Sub
#End If

Private Sub cmdAddHeader_Click()
    lvHeaders.SetFocus
    Set lvHeaders.SelectedItem = lvHeaders.ListItems.Add(, , "", , 1)
    lvHeaders.SelectedItem.ListSubItems.Add , , ""
    lvHeaders.StartLabelEdit
End Sub

Private Sub cmdDeleteHeader_Click()
    If Not lvHeaders.SelectedItem Is Nothing Then
        If lvHeaders.SelectedItem.Selected Then
            lvHeaders.ListItems.Remove lvHeaders.SelectedItem.Index
        End If
    End If
End Sub

Private Sub cmdEditHeaderName_Click()
    On Error Resume Next
    lvHeaders.SetFocus
    lvHeaders.StartLabelEdit
End Sub

Private Sub cmdEditHeaderValue_Click()
    On Error GoTo exitsub
    If lvHeaders.SelectedItem Is Nothing Then GoTo exitsub
    Set SelectedListItem = lvHeaders.SelectedItem
    Dim SubItemLeft As Integer
    SubItemLeft = SelectedListItem.ListSubItems(1).Left
    'Dim SecondColumnWidth As Integer
    'SecondColumnWidth = lvHeaders.ColumnHeaders(2).Width
    With txtEdit
        .Top = lvHeaders.Top + SelectedListItem.Top + 15
        .Left = lvHeaders.Left + Max(SubItemLeft, 0) + 30
        .Width = Min(Min(lvHeaders.ColumnHeaders(2).Width, lvHeaders.Width - SubItemLeft - 60), lvHeaders.Width - 60) - (-CBool(GetWindowLong(lvHeaders.hWnd, GWL_STYLE) And WS_VSCROLL)) * ScrollBarWidth * 15 'Max((ScrollBarWidth * 15 - Max(lvHeaders.Width - (SecondColumnWidth + SubItemLeft), 0)), 0)
        .Text = SelectedListItem.ListSubItems(1).Text
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    OKButton.Default = 0
    CancelButton.Cancel = 0
exitsub:
End Sub

Private Sub txtEdit_LostFocus()
    On Error Resume Next
    SelectedListItem.ListSubItems(1).Text = txtEdit.Text
    txtEdit.Visible = False
    Set SelectedListItem = Nothing
    OKButton.Default = -1
    CancelButton.Cancel = -1
End Sub
 
Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Select Case KeyAscii
        Case 10, 13
            SelectedListItem.ListSubItems(1).Text = txtEdit.Text
endedit:
            txtEdit.Visible = False
            Set SelectedListItem = Nothing
            OKButton.Default = -1
            CancelButton.Cancel = -1
            lvHeaders.SetFocus
        Case 27
            GoTo endedit
    End Select
End Sub

Private Sub lvHeaders_AfterLabelEdit(Cancel As Boolean, NewString As String)
    NewString = Trim$(NewString)
    If NewString = "" Then
invalidname:
        Cancel = True
        MsgBox t("헤더 이름이 잘못되었습니다.", "Invalid header name."), 16
        Exit Sub
    End If
    
    Dim i%
    For i = 1 To Len(NewString)
        Select Case Mid$(NewString, i, 1)
            Case "a" To "z", "A" To "Z", "0" To "9", "-", "_"
            Case Else
                GoTo invalidname
        End Select
    Next i
    
    For i = 1 To lvHeaders.ListItems.Count
        If LCase(lvHeaders.ListItems(i).Text) = LCase(NewString) Then
            Cancel = True
            MsgBox t("해당 이름이 이미 존재합니다.", "Duplicate header name."), 16
            Exit Sub
            Exit For
        End If
    Next i
End Sub

Private Sub lvHeaders_ItemDblClick(Item As LvwListItem, ByVal Button As Integer)
    If Item.Selected And Button = 1 Then cmdEditHeaderValue_Click
End Sub

Private Sub lvHeaders_ItemSelect(Item As LvwListItem, ByVal Selected As Boolean)
    On Error GoTo justdisable
    If Selected Then
        cmdDeleteHeader.Enabled = -1
        cmdEditHeaderName.Enabled = -1
        cmdEditHeaderValue.Enabled = -1
        Exit Sub
    End If
justdisable:
    cmdDeleteHeader.Enabled = 0
    cmdEditHeaderName.Enabled = 0
    cmdEditHeaderValue.Enabled = 0
End Sub

