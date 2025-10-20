Attribute VB_Name = "PropertySheets"
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PropertySheet Lib "comctl32.dll" Alias "PropertySheetA" (psh As PROPSHEETHEADER) As Long
Private Declare Function CreatePropertySheetPage Lib "comctl32.dll" Alias "CreatePropertySheetPageA" (psp As PROPSHEETPAGE) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontW" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsW" (ByVal hDC As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long

Private Const LOGPIXELSY As Long = 90&
Private Const FW_NORMAL As Long = 400&
Private Const DEFAULT_CHARSET As Long = 1&
Private Const OUT_DEFAULT_PRECIS As Long = 0&
Private Const CLIP_DEFAULT_PRECIS As Long = 0&
Private Const DEFAULT_QUALITY As Long = 0&
Private Const DEFAULT_PITCH As Long = 0&
Private Const FF_DONTCARE As Long = 0&
Private Const WS_CHILD As Long = &H40000000
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_EX_CONTROLPARENT As Long = &H10000
Private Const WM_INITDIALOG As Long = &H110&
Private Const WM_DESTROY As Long = &H2&
Private Const WM_NOTIFY As Long = &H4E&
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_SYSKEYDOWN As Long = &H104
Private Const WM_USER As Long = &H400&
Private Const WM_MENUCHAR As Long = 288&
'Private Const WM_MENUSELECT As Long = 287&
Private Const MF_SYSMENU As Long = 8192&
Private Const PSH_PROPTITLE As Long = &H1&
Private Const PSH_USECALLBACK As Long = &H100&
Private Const PSP_USETITLE As Long = &H8&
Private Const PSP_DLGINDIRECT As Long = &H1&
Private Const PSM_CHANGED As Long = (WM_USER + 104&)
Private Const PSM_UNCHANGED As Long = (WM_USER + 109&)
Private Const PSM_SETTITLE As Long = (WM_USER + 111&)
Private Const PSM_GETTABCONTROL As Long = (WM_USER + 116&)
Private Const PSM_ISDIALOGMESSAGE As Long = (WM_USER + 117&)
Private Const BM_CLICK As Long = &HF5&
Private Const DS_CONTROL As Long = &H2000&
Private Const PSN_FIRST As Long = -200&
Private Const PSN_APPLY As Long = PSN_FIRST - 2&
Private Const PSN_RESET As Long = PSN_FIRST - 3&
Private Const PSN_SETACTIVE As Long = PSN_FIRST - 0&
Private Const PSN_KILLACTIVE As Long = PSN_FIRST - 1&
Private Const PSNRET_NOERROR As Long = 0&
Private Const PSNRET_INVALID As Long = 1&
Private Const PSCB_INITIALIZED As Long = 1&
Private Const DWLP_MSGRESULT As Long = 0&
Private Const GWL_STYLE As Long = -16&
Private Const GWL_EXSTYLE As Long = -20&
'Private Const GWL_ID As Long = -12&
Private Const GWL_WNDPROC As Long = -4&
Private Const SPI_GETNONCLIENTMETRICS As Long = 41&
Private Const SW_SHOW As Long = 5&
Private Const BN_CLICKED As Long = 0&
Private Const BS_RADIOBUTTON As Long = 4&
Private Const BS_AUTORADIOBUTTON As Long = 9&

Private Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Integer
    tmLastChar As Integer
    tmDefaultChar As Integer
    tmBreakChar As Integer
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

Private Type LOGFONT
    LFHeight As Long
    LFWidth As Long
    LFEscapement As Long
    LFOrientation As Long
    LFWeight As Long
    LFItalic As Byte
    LFUnderline As Byte
    LFStrikeOut As Byte
    LFCharset As Byte
    LFOutPrecision As Byte
    LFClipPrecision As Byte
    LFQuality As Byte
    LFPitchAndFamily As Byte
    LFFaceName As String * 32
End Type

Private Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONT
    iSmCaptionWidth As Long
    iSmCaptionHeight As Long
    lfSmCaptionFont As LOGFONT
    iMenuHeight As Long
    iMenuWidth As Long
    lfMenuFont As LOGFONT
    lfStatusFont As LOGFONT
    lfMessageFont As LOGFONT
End Type

Private Type NMHDR
    hWndFrom As Long
    idFrom As Long
    code As Long
End Type

Private Type PROPSHEETPAGE
    dwSize As Long
    dwFlags As Long
    hInstance As Long
    pResource As Long
    hIcon As Long
    pszTitle As String
    pfnDlgProc As Long
    lParam As Long
    pfnCallback As Long
    pcRefParent As Long
    'pszHeaderTitle As Long
    'pszHeaderSubTitle As Long
    'hActCtx As Long
    'hbmHeader As Long
End Type

Private Type PROPSHEETHEADER
    dwSize As Long
    dwFlags As Long
    hwndParent As Long
    hInstance As Long
    hIcon As Long
    pszCaption As String
    nPages As Long
    nStartPage As Long
    phpage As Long
    pfnCallback As Long
    'hbmWatermark As Long
    'hplWatermark As Long
    'hbmHeader As Long
End Type

Private Type DLGTEMPLATE
    Style As Long
    dwExtendedStyle As Long
    cdit As Integer
    X As Integer
    Y As Integer
    CX As Integer
    CY As Integer
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Dim DialogWidth&, DialogHeight&
Dim OldSheetProc As Long
Dim PageForms As Collection
Dim hwndMap As Collection
Dim hdlgMap As Collection
Dim Forms As Collection
Dim CurrentPage As Long

Sub EnableApply(Page As Form)
    SendMessage GetParent(hdlgMap(CStr(Page.hWnd))), PSM_CHANGED, hdlgMap(CStr(Page.hWnd)), 0&
End Sub

Sub DisableApply(Page As Form)
    SendMessage GetParent(hdlgMap(CStr(Page.hWnd))), PSM_UNCHANGED, hdlgMap(CStr(Page.hWnd)), 0&
End Sub

Private Function PixelsToDialogUnits(ByVal px As Long, Mode As Byte) As Integer
    Dim hDC As Long
    Dim hFont As Long
    Dim hOld As Long
    Dim tm As TEXTMETRIC
    Dim logHeight As Long
    Dim FontName$, FontSize&
    Dim ncm As NONCLIENTMETRICS
    Dim lpy As Long
    ncm.cbSize = 340&

    hDC = GetDC(0&)
    lpy = GetDeviceCaps(hDC, LOGPIXELSY)
    If Mode = 1 And SystemParametersInfo(SPI_GETNONCLIENTMETRICS, ncm.cbSize, ncm, 0&) Then
        FontName = Left$(ncm.lfMessageFont.LFFaceName, InStr(ncm.lfMessageFont.LFFaceName, vbNullChar) - 1)
        FontSize = CLng(-ncm.lfMessageFont.LFHeight * 72 / lpy)
    Else
        FontName = "MS Shell Dlg"
        FontSize = 8&
    End If
    If Mode = 1 Then logHeight = -((FontSize * lpy) \ 72&)
    hFont = CreateFont(logHeight, 0&, 0&, 0&, FW_NORMAL, 0&, 0&, 0&, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH Or FF_DONTCARE, StrPtr(FontName))
    hOld = SelectObject(hDC, hFont)
    GetTextMetrics hDC, tm
    If Mode = 0 Then
        PixelsToDialogUnits = CInt((px * 4&) \ tm.tmAveCharWidth)
    Else
        PixelsToDialogUnits = CInt((px * 8&) \ tm.tmHeight)
    End If
    SelectObject hDC, hOld
    DeleteObject hFont
    ReleaseDC 0&, hDC
End Function

Private Function ControlsOf(frmForm As Form) As Collection
    Set ControlsOf = New Collection
    Dim ctrl As Control
    Dim i As Long
    Dim inserted As Boolean
    Dim ctrlTabIndex As Integer
    
    For Each ctrl In frmForm.Controls
        inserted = False
        
        On Error Resume Next
        Err.Clear
        ctrlTabIndex = ctrl.TabIndex
        If Err.Number <> 0 Then ctrlTabIndex = -9999
        On Error GoTo 0
        
        If ControlsOf.Count = 0 Then
            ControlsOf.Add ctrl
        Else
            For i = 1 To ControlsOf.Count
                Dim cmpCtrl As Control
                Dim cmpTabIndex As Long
                
                On Error Resume Next
                Err.Clear
                cmpTabIndex = ControlsOf.Item(i).TabIndex
                If Err.Number <> 0 Then cmpTabIndex = -9999
                On Error GoTo 0
                
                If ctrlTabIndex < cmpTabIndex Then
                    ControlsOf.Add ctrl, Before:=i
                    inserted = True
                    Exit For
                End If
            Next i
            
            If Not inserted Then ControlsOf.Add ctrl
        End If
    Next ctrl
End Function

Private Function SheetWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case WM_MENUCHAR
            Dim KeyAscii As Integer, ch As String
            If HiWord(wParam) = MF_SYSMENU Then
                KeyAscii = LoWord(wParam)
                If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
                If KeyAscii >= 65 And KeyAscii <= 90 Then
                    ch = "&" & Chr$(KeyAscii)
                    On Error Resume Next
                    Dim ctrl As Control, ctrls As Collection, i&, Caption$
                    Set ctrls = ControlsOf(PageForms(CStr(hwndMap(CStr(CurrentPage)))))
                    For i = 1 To ctrls.Count
                        Set ctrl = ctrls(i)
                        Caption = ctrl.Caption
                        If InStr(Caption, ch) > 0 And ctrl.Visible Then
                            Do While TypeOf ctrl Is Label
                                i = i + 1&
                                Set ctrl = ctrls(i)
                            Loop
                            SetFocus ctrl.hWnd
                            If TypeOf ctrl Is CommandButton Then
                                SendMessage ctrl.hWnd, BM_CLICK, 0&, 0&
                            ElseIf TypeOf ctrl Is OptionButton Then
                                ctrl.Value = True
                            ElseIf TypeOf ctrl Is CheckBox Then
                                If ctrl.Value = 0 Then ctrl.Value = 1 Else ctrl.Value = 0
                            End If

                            SheetWndProc = 3&
                            Exit Function
                        End If
                    Next i
                End If
            End If
    End Select
    SheetWndProc = CallWindowProc(OldSheetProc, hWnd, uMsg, wParam, lParam)
End Function

Private Function PageDlgProc(ByVal hDlg As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim PageForm As Form
    On Error Resume Next
    Set PageForm = PageForms(CStr(hwndMap(CStr(hDlg))))
    On Error GoTo 0

    Select Case uMsg
        Case WM_INITDIALOG
            SetWindowLong hDlg, GWL_EXSTYLE, GetWindowLong(hDlg, GWL_EXSTYLE) Or WS_EX_CONTROLPARENT
            Dim ID As Long
            CopyMemory ID, ByVal (lParam + 28&), 4&
            Set PageForm = Forms(ID)
            PageForm.BorderStyle = 0
            PageForm.Caption = ""
            PageForms.Add PageForm, CStr(PageForm.hWnd)
            hwndMap.Add PageForm.hWnd, CStr(hDlg)
            hdlgMap.Add hDlg, CStr(PageForm.hWnd)
            PageForm.Left = 0
            PageForm.Top = 0
            PageForm.Width = DialogWidth
            PageForm.Height = DialogHeight
            Dim TabBackgroundHint As Long: TabBackgroundHint = GetThemeColor(PageForm.hWnd, "TAB", 9&, 1&, 3821&, &H8000000F)
            PageForm.BackColor = TabBackgroundHint
            On Error Resume Next
            PageForm.Initialize
            'On Error GoTo 0
            'On Error Resume Next
            '방법 1, 폼을 통째로 이동
            SetWindowLong PageForm.hWnd, GWL_STYLE, GetWindowLong(PageForm.hWnd, GWL_STYLE) Or WS_CHILD
            SetWindowLong PageForm.hWnd, GWL_EXSTYLE, GetWindowLong(PageForm.hWnd, GWL_EXSTYLE) Or WS_EX_CONTROLPARENT
            SetParent PageForm.hWnd, hDlg
            Dim ctrl As Control, hwndSwp As Long
            For Each ctrl In ControlsOf(PageForm)
                hwndSwp = ctrl.hWnd
                If TypeOf ctrl Is Frame Or TypeOf ctrl Is PictureBox Then
                    SetWindowLong ctrl.hWnd, GWL_EXSTYLE, GetWindowLong(ctrl.hWnd, GWL_EXSTYLE) Or WS_EX_CONTROLPARENT
                ElseIf TypeOf ctrl Is SpinBox Or TypeOf ctrl Is CommandButtonW Or TypeOf ctrl Is Slider Then
                    SetWindowLong ctrl.hWndUserControl, GWL_STYLE, GetWindowLong(ctrl.hWndUserControl, GWL_STYLE) Or WS_TABSTOP
                    hwndSwp = ctrl.hWndUserControl
                End If
                SetWindowPos hwndSwp, 1&, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE
                If (Not TypeOf ctrl.Container Is PictureBox) And (Not TypeOf ctrl Is PictureBox) And (Not TypeOf ctrl Is Shape) And (Not TypeOf ctrl Is ComboBox) And (Not TypeOf ctrl Is TextBox) And (Not TypeOf ctrl Is SpinBox) And ctrl.Tag <> "nobackcolorchange" And ctrl.Tag <> "nobackcolorchange novisualstylechange" Then ctrl.BackColor = TabBackgroundHint
            Next ctrl
            ShowWindow PageForm.hWnd, SW_SHOW
            
            '방법 2, 모든 컨트롤을 각각 이동
'            Dim ctrl As Control, Container As Frame
'            For Each ctrl In ControlsOf(PageForm)
'                If ctrl.Container Is PageForm Then
'                    SetParent ctrl.hWnd, hDlg
'                End If
'                If TypeOf ctrl Is Frame Then
'                    SetWindowLong ctrl.hWnd, GWL_EXSTYLE, GetWindowLong(ctrl.hWnd, GWL_EXSTYLE) Or WS_EX_CONTROLPARENT
'                ElseIf TypeOf ctrl Is SpinBox Or TypeOf ctrl Is CommandButtonW Or TypeOf ctrl Is Slider Then
'                    SetWindowLong ctrl.hWndUserControl, GWL_STYLE, GetWindowLong(ctrl.hWndUserControl, GWL_STYLE) Or WS_TABSTOP
'                End If
'            Next ctrl
            On Error GoTo 0
            
            PageDlgProc = 1&
            Exit Function
        Case WM_NOTIFY
            Dim pnmh As NMHDR
            CopyMemory pnmh, ByVal lParam, LenB(pnmh)
            Select Case pnmh.code
                Case PSN_APPLY
                    Dim ApplyResult As Boolean: ApplyResult = True
                    On Error Resume Next
                    ApplyResult = PageForm.ApplyChanges()
                    On Error GoTo 0
                    SetWindowLong hDlg, DWLP_MSGRESULT, -CLng(Not ApplyResult)
                    
                    PageDlgProc = 1&
                    Exit Function
                Case PSN_RESET
                    On Error Resume Next
                    PageForm.CancelClick
                    On Error GoTo 0
                    SetWindowLong hDlg, DWLP_MSGRESULT, 0&
                    
                    PageDlgProc = 0&
                    Exit Function
                Case PSN_SETACTIVE
                    CurrentPage = hDlg
                    On Error Resume Next
                    PageForm.Activate
                    On Error GoTo 0
                    SetWindowLong hDlg, DWLP_MSGRESULT, 0&
                    
                    PageDlgProc = 1&
                    Exit Function
                Case PSN_KILLACTIVE
                    On Error Resume Next
                    PageForm.Deactivate
                    On Error GoTo 0
                    SetWindowLong hDlg, DWLP_MSGRESULT, 0&
                    
                    PageDlgProc = 1&
                    Exit Function
            End Select
    End Select
    
    PageDlgProc = 0&
End Function

Private Function PropSheetCallback(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case PSCB_INITIALIZED
            OldSheetProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf SheetWndProc)
    End Select
    
    PropSheetCallback = 0&
End Function

Private Sub SetValue(ByRef Dest As Long, ByVal Value As Long)
    Dest = Value
End Sub

Sub ShowPropertySheetEx(Parent As Form, Title As String, ParamArray Pages())
    If IsArray(Pages(0)) Then Pages = Pages(0)
    
    Dim psp As PROPSHEETPAGE
    Dim pspHandle As Long
    Dim psh As PROPSHEETHEADER
    Dim pspArray() As Long
    ReDim pspArray(0 To UBound(Pages) / 4)
    Dim Ptr As Long
    Set hwndMap = New Collection
    Set PageForms = New Collection
    Set hdlgMap = New Collection
    Set Forms = New Collection
    DialogWidth = 0&: DialogHeight = 0&
    Dim hMems As New Collection
    
    Dim i As Byte
    For i = 0 To UBound(Pages) Step 4
        Dim hMem As Long, pMem As Long
        Dim dlg As DLGTEMPLATE
        dlg.Style = WS_CHILD 'Or DS_CONTROL
        'dlg.dwExtendedStyle = 0&
        'dlg.cdit = 0&
        'dlg.X = 0&
        'dlg.Y = 0&
        If DialogWidth < Pages(i + 2) Then DialogWidth = Pages(i + 2)
        If DialogHeight < Pages(i + 3) Then DialogHeight = Pages(i + 3)
        dlg.CX = PixelsToDialogUnits(Pages(i + 2) \ Screen.TwipsPerPixelX, 0)
        dlg.CY = PixelsToDialogUnits(Pages(i + 3) \ Screen.TwipsPerPixelY, 1)
        hMem = GlobalAlloc(&H40&, LenB(dlg) + 256)
        hMems.Add hMem
        pMem = GlobalLock(hMem)
        CopyMemory ByVal pMem, dlg, LenB(dlg)
        GlobalUnlock hMem
        
        psp.dwSize = LenB(psp)
        psp.dwFlags = PSP_DLGINDIRECT Or PSP_USETITLE
        psp.hInstance = App.hInstance
        psp.pResource = pMem
        SetValue psp.pfnDlgProc, AddressOf PageDlgProc
        psp.pszTitle = Pages(i + 1)
        psp.lParam = i / 4 + 1
        Forms.Add Pages(i)
        
        pspHandle = CreatePropertySheetPage(psp)
        pspArray(i / 4) = pspHandle
    Next i
    
    psh.dwSize = LenB(psh)
    psh.dwFlags = PSH_USECALLBACK 'Or PSH_PROPTITLE
    psh.hwndParent = Parent.hWnd
    psh.hInstance = App.hInstance
    psh.pszCaption = Title
    psh.nPages = UBound(Pages) / 4 + 1&
    psh.nStartPage = 0&
    psh.phpage = VarPtr(pspArray(0))
    SetValue psh.pfnCallback, AddressOf PropSheetCallback
    
    PropertySheet psh
    
    For i = 0 To UBound(Pages) Step 4
        GlobalFree hMems(1)
        hMems.Remove 1
        Unload Pages(i)
    Next i
End Sub

Sub ShowPropertySheet(Parent As Form, Title As String, ParamArray Pages())
    Dim NewPages(), i As Byte
    ReDim NewPages(0 To (UBound(Pages) + 1) * 4 - 1)
    For i = 0 To UBound(Pages)
        Set NewPages(i * 4) = Pages(i)
        NewPages(i * 4 + 1) = Pages(i).Caption
        NewPages(i * 4 + 2) = Pages(i).ScaleWidth
        NewPages(i * 4 + 3) = Pages(i).ScaleHeight
    Next i
    ShowPropertySheetEx Parent, Title, NewPages
End Sub
