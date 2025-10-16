Attribute VB_Name = "Functions"
'참고자료
'- https://www.vbforums.com/showthread.php?457171-RESOLVED-How-to-get-Desktop-Path-in-VB
'- https://www.vbforums.com/showthread.php?445574-Reading-shortcut-information
'- https://www.vbforums.com/showthread.php?430704-RESOLVED-Get-drive-size-space
'- https://www.codeguru.com/visual-basic/displaying-the-file-properties-dialog/
'- http://vbcity.com/forums/t/105530.aspx
'- https://www.vbforums.com/showthread.php?644597-How-to-get-String-from-Pointer-in-VB
'- https://www.vbforums.com/showthread.php?903019-UxTheme-dll-Visual-Styles-help-request-button-face-color-when-mouse-hovers-over

Option Explicit

Public Const MAX_THREAD_COUNT_CONTROL As Integer = 655& '679&
Public Const MAX_32BIT_SIGNED_INT As Long = 2147483647
Public Const PROPERTY_SHEET_BUTTON_WIDTH As Integer = 1320
Public Const PROPERTY_SHEET_BUTTON_HEIGHT As Integer = 360

Public MsgBoxResults As Collection
Public InputBoxResults As Collection

'Declare Function SetProcessDpiAwarenessContext Lib "user32" (ByVal context As Long) As Long
'Declare Function SetProcessDPIAware Lib "user32" () As Long
Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProcess As Long, ByRef Wow64Process As Long) As Long
Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
'Private Declare Function RtlGetVersion Lib "ntdll" (lpVersionInformation As OSVERSIONINFO) As Long
Declare Function DwmSetWindowAttribute Lib "dwmapi.dll" (ByVal hWnd As Long, ByVal dwAttribute As Long, ByRef pvAttribute As Long, ByVal cbAttribute As Long) As Long
Private Declare Function DwmIsCompositionEnabled Lib "dwmapi.dll" (ByRef pfEnabled As Long) As Long
'Declare Function DwmEnableComposition Lib "dwmapi.dll" (ByVal uCompositionAction As Long) As Long
'Declare Function DwmExtendFrameIntoClientArea Lib "dwmapi.dll" (ByVal hWnd As Long, Margin As MARGINS) As Long
'Declare Sub DwmGetColorizationParameters Lib "dwmapi.dll" Alias "#127" (ByRef Parameters As DWM_COLORIZATION_PARAMS)
'Declare Sub DwmSetColorizationParameters Lib "dwmapi.dll" Alias "#131" (ByRef Parameters As DWM_COLORIZATION_PARAMS, Optional ByVal DoNotSaveToRegistry As Boolean = True)
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32" Alias "RegEnumKeyExW" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As Long, ByRef lpcName As Long, Optional ByVal lpReserved As Long, Optional ByVal lpClass As Long, Optional ByRef lpcClass As Long, Optional ByVal lpftLastWriteTime As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32" Alias "RegQueryInfoKeyW" (ByVal hKey As Long, Optional ByVal lpClass As Long, Optional ByRef lpcClass As Long, Optional ByVal lpReserved As Long, Optional ByRef lpcSubKeys As Long, Optional ByRef lpcMaxSubKeyLen As Long, Optional ByRef lpcMaxClassLen As Long, Optional ByRef lpcValues As Long, Optional ByRef lpcMaxValueNameLen As Long, Optional ByRef lpcMaxValueLen As Long, Optional ByRef lpcbSecurityDescriptor As Long, Optional ByVal lpftLastWriteTime As Long) As Long
Private Declare Function SysReAllocStringLen Lib "oleaut32" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long, Optional ByVal Length As Long) As Long
Declare Function GetUserDefaultUILanguage Lib "kernel32" () As Integer
Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
'Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpMII As MENUITEMINFO) As Long
'Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpMII As MENUITEMINFO) As Long
'Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
'Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
'Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpMII As MENUITEMINFO) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
'Declare Function CheckMenuRadioItem Lib "user32" (ByVal hMenu As Long, ByVal un1 As Long, ByVal un2 As Long, ByVal un3 As Long, ByVal un4 As Long) As Long
Private Declare Function CryptBinaryToString Lib "crypt32" Alias "CryptBinaryToStringW" (ByVal pbBinary As Long, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As Long, ByRef pcchString As Long) As Long
Private Const CRYPT_STRING_BASE64 As Long = 1&
Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryW" (ByVal pszString As Long, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByRef pcbBinary As Long, ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
'Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ExpandEnvironmentStringsA Lib "kernel32" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
'Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
'Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
'Declare Function PrintWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcBlt As Long, ByVal nFlags As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function SysAllocStringByteLen Lib "oleaut32.dll" (Optional ByVal pszStrPtr As Long, Optional ByVal Length As Long) As String
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadID As Long) As Long
'Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function X_GetThemeColor Lib "uxtheme.dll" Alias "GetThemeColor" (ByVal hTheme As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal iPropId As Long, pColor As Long) As Long
Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Long
Private Declare Function IsThemeActive Lib "uxtheme.dll" () As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
'Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, Optional ByVal dwDuration As Long = 250) As Long
'Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, Optional ByVal bInvert As Long = 1&) As Long
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal lpRootPathName As String) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As LARGE_INTEGER, lpTotalNumberOfBytes As LARGE_INTEGER, lpTotalNumberOfFreeBytes As LARGE_INTEGER) As Long
Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (lpChooseColor As ChooseColorStruct) As Long
Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function ShellExecuteEx Lib "shell32" (ByRef s As SHELLEXECUTEINFO) As Long
Declare Function GetKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
'Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
'Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
'Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
'Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Declare Function ExitProcess Lib "kernel32" (ByVal ExitCode As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
'Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function SetLayout Lib "gdi32" (ByVal hDC As Long, ByVal dwLayout As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByRef lppt As Any, ByVal cPoints As Long) As Long
Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As POINTAPI) As Long
Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As Any, lpLastAccessTime As Any, lpLastWriteTime As Any) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As Any, lpLastAccessTime As Any, lpLastWriteTime As Any) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
'Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'Declare Sub IUnknown_AtomicRelease Lib "shlwapi" (ppUnk As Any)
Declare Function OleCreatePictureIndirect Lib "oleaut32" (lpPictDesc As PICTDESC, riid As IID, ByVal fOwn As Boolean, lplpvObj As IPicture) As Long
Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As Any, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal uFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprcRect As Any) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
'Declare Function Release Lib "ole32.dll" Alias "IUnknown_Release_Proxy" (pUnk As Any) As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function BeginDeferWindowPos Lib "user32" (ByVal nNumWindows As Long) As Long
Declare Function DeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long, ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal uFlags As Long) As Long
Declare Function EndDeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long) As Long
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Public Const INVALID_HANDLE_VALUE As Long = -1&

Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As Currency
    ftLastAccessTime As Currency
    ftLastWriteTime As Currency
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type

Public Const WS_SIZEBOX As Long = &H40000

Public Const WS_EX_APPWINDOW As Long = &H40000
Public Const WS_EX_TOOLWINDOW As Long = &H80&

Public Const TPM_LEFTALIGN As Long = &H0&
Public Const TPM_RETURNCMD As Long = &H100&

Public Const HTERROR As Long = -2       ' Error
Public Const HTTRANSPARENT As Long = -1 ' Transparent
Public Const HTNOWHERE As Long = 0      ' Outside window
Public Const HTCLIENT As Long = 1       ' Client area
Public Const HTCAPTION As Long = 2      ' Title bar
Public Const HTSYSMENU As Long = 3      ' System menu
Public Const HTGROWBOX As Long = 4      ' Size box (old)
Public Const HTSIZE As Long = HTGROWBOX
Public Const HTMENU As Long = 5         ' Menu
Public Const HTHSCROLL As Long = 6      ' Horizontal scroll bar
Public Const HTVSCROLL As Long = 7      ' Vertical scroll bar
Public Const HTMINBUTTON As Long = 8    ' Minimize button
Public Const HTMAXBUTTON As Long = 9    ' Maximize button
Public Const HTLEFT As Long = 10        ' Left border
Public Const HTRIGHT As Long = 11       ' Right border
Public Const HTTOP As Long = 12         ' Top border
Public Const HTTOPLEFT As Long = 13     ' Top-left corner
Public Const HTTOPRIGHT As Long = 14    ' Top-right corner
Public Const HTBOTTOM As Long = 15      ' Bottom border
Public Const HTBOTTOMLEFT As Long = 16  ' Bottom-left corner
Public Const HTBOTTOMRIGHT As Long = 17 ' Bottom-right corner
Public Const HTBORDER As Long = 18      ' Border (obsolete)
Public Const HTREDUCE As Long = HTMINBUTTON
Public Const HTZOOM As Long = HTMAXBUTTON
Public Const HTSIZEFIRST As Long = HTLEFT
Public Const HTSIZELAST As Long = HTBOTTOMRIGHT
Public Const HTOBJECT As Long = 19      ' Object in window
Public Const HTCLOSE As Long = 20       ' Close button
Public Const HTHELP As Long = 21        ' Help button

Public Const LVM_SETVIEW As Long = 4238&
Public Const LVM_SETIMAGELIST As Long = &H1003&
Public Const LVSIL_SMALL As Long = 1&
Public Const LVSIL_NORMAL As Long = 0&

'Public Const SHGFI_DISPLAYNAME As Long = &H200&
Public Const SHGFI_SYSICONINDEX As Long = &H4000&
Public Const SHGFI_ICON As Long = &H100&
Public Const SHGFI_LARGEICON As Long = &H0&
Public Const SHGFI_SMALLICON As Long = &H1&
Public Const SHGFI_USEFILEATTRIBUTES As Long = &H10&
Public Const SHGFI_TYPENAME As Long = &H400&

Private Const GMEM_MOVEABLE = &H2&

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_FRAMECHANGED As Long = &H20&

Enum WindowSkin
    'System = 0
    Bluemetal = 1
    Greenmetal = 2
End Enum

Enum SP_RESULTS
    SP_SUCCESS = 0
    SP_CREATEPIPEFAILED = &H80042B00
    SP_CREATEPROCFAILED = &H80042B01
End Enum

Type PROCESSINFO
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long                'LPBYTE
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 260
    szTypeName As String * 80
End Type

Type IID
    Data1       As Long
    Data2       As Integer
    Data3       As Integer
    Data4(7&)   As Byte
End Type

Type PICTDESC
    cbSize          As Long
    PicType         As Long
    hgdiObj         As Long
    hPalOrXYExt     As Long
    Reserved        As Long
End Type

Enum VbMsgBoxResultEx
'    vbAbort = 3
'    vbCancel = 2
'    vbIgnore = 5
'    vbNo = 7
'    vbOK = 1
'    vbRetry = 4
'    vbYes = 6
    vbTryAgain = 10
    vbContinue = 11
End Enum

Enum VbMsgBoxStyleEx
'    vbAbortRetryIgnore = 2
'    vbApplicationModal = 0
'    vbCritical = 16
'    vbDefaultButton1 = 0
'    vbDefaultButton2 = 256
'    vbDefaultButton3 = 512
'    vbDefaultButton4 = 768
'    vbExclamation = 48
'    vbInformation = 64
'    vbMsgBoxHelpButton = 16384
'    vbMsgBoxRight = 524288
'    vbMsgBoxRtlReading = 1048576
'    vbMsgBoxSetForeground = 65536
'    vbOKCancel = 1
'    vbOKOnly = 0
'    vbQuestion = 32
'    vbRetryCancel = 5
'    vbSystemModal = 4096
'    vbYesNo = 4
'    vbYesNoCancel = 3
    vbCancelTryContinue = 6
    vbYesNoEx = 7
End Enum

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(63) As Byte
    StandardDate(7) As Integer
    StandardBias As Long
    DaylightName(63) As Byte
    DaylightDate(7) As Integer
    DaylightBias As Long
End Type

Public Const WM_NOTIFY As Long = &H4E&
Public Const WM_MOVE As Long = &H3&
'Public Const WM_MOVING As Long = &H216&
Public Const WM_SETCURSOR As Long = &H20&
Public Const WM_NCPAINT As Long = &H85&
Public Const WM_COMMAND As Long = &H111&
Public Const WM_SIZE As Long = &H5&
'Public Const WM_SIZING As Long = &H214&
Public Const WM_GETMINMAXINFO As Long = &H24
Public Const WM_SYSCOMMAND As Long = &H112&
Public Const WM_INITMENU As Long = &H116&
Public Const WM_SETTINGCHANGE As Long = &H1A
Public Const WM_DWMCOMPOSITIONCHANGED As Long = &H31E&
Public Const WM_THEMECHANGED As Long = &H31A&
Public Const WM_DPICHANGED As Long = &H2E0&
Public Const WM_CTLCOLORSCROLLBAR As Long = &H137&
Public Const WM_CTLCOLORSTATIC As Long = &H138&
Public Const WM_CTLCOLORBTN As Long = &H135&
Public Const WM_PAINT As Long = &HF&
Public Const WM_PRINTCLIENT As Long = &H318&
Public Const hWnd_TOPMOST As Long = -1&
Public Const hWnd_NOTOPMOST As Long = -2&
Public Const WM_NCCALCSIZE As Long = &H83&
Public Const WM_NCHITTEST As Long = &H84&
Public Const WM_NCACTIVATE As Long = &H86&
Public Const WA_INACTIVE As Long = 0&
Public Const WA_ACTIVE As Long = 1&

Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWNORMAL = 1

Type WINDOWPLACEMENT
    Length As Long
    Flags As Long
    showCmd As Long
    ptMinPositionX As Long
    ptMinPositionY As Long
    ptMaxPositionX As Long
    ptMaxPositionY As Long
    rcNormalPositionLeft As Long
    rcNormalPositionTop As Long
    rcNormalPositionRight As Long
    rcNormalPositionBottom As Long
End Type

Private Type FILETIME
    LowDateTime As Long
    HighDateTime As Long
End Type
    
Private Type SYSTEMTIME
    Year As Integer
    Month As Integer
    DayOfWeek As Integer
    Day As Integer
    Hour As Integer
    Minute As Integer
    Second As Integer
    Milliseconds As Integer
End Type

Private Const CB_ERR As Long = -1&
Private Const CB_ADDSTRING As Long = &H143&
Private Const CB_RESETCONTENT As Long = &H14B&
Private Const CB_SETITEMDATA As Long = &H151&

Public Const hWnd_DESKTOP As Long = &H0&
Public Const LAYOUT_RTL As Long = &H1&
Public Const WS_EX_LAYOUTRTL As Long = &H400000, WS_EX_RTLREADING As Long = &H2000&
Public Const RDW_UPDATENOW As Long = &H100&, RDW_INVALIDATE As Long = &H1&, RDW_ERASE As Long = &H4&, RDW_ALLCHILDREN As Long = &H80&

Private Const PROCESS_ALL_ACCESS As Long = &H1F0FFF

Public Const GWL_STYLE As Long = -16&
'Public Const GWL_HWNDPARENT As Long = -8&

'Public Const DWM_EC_DISABLECOMPOSITION As Long = 0&
'Public Const DWM_EC_ENABLECOMPOSITION As Long = 1&

Public Const SND_APPLICATION = &H80
Public Const SND_ALIAS = &H10000
Public Const SND_ALIAS_ID = &H110000
Public Const SND_ASYNC = &H1
Public Const SND_FILENAME = &H20000
Public Const SND_LOOP = &H8
Public Const SND_MEMORY = &H4
Public Const SND_NODEFAULT = &H2
Public Const SND_NOSTOP = &H10
Public Const SND_NOWAIT = &H2000
Public Const SND_PURGE = &H40
Public Const SND_RESOURCE = &H40004
Public Const SND_SYNC = &H0

Public Const TMT_TEXTCOLOR As Long = 3803

Public Const WM_NCLBUTTONDOWN = &HA1

Enum ResourceType
    BITMAP = 2
    Icon = 3
    RCData = 10
    Manifest = 24
End Enum

Type POINTAPI
   X As Long
   Y As Long
End Type

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

'Type DWM_COLORIZATION_PARAMS
'    ColorBlue  As Byte
'    ColorGreen As Byte
'    ColorRed   As Byte
'    ColorAlpha As Byte
'
'    AfterGlowBlue  As Byte
'    AfterGlowGreen As Byte
'    AfterGlowRed   As Byte
'    AfterGlowAlpha As Byte
'
'    ColorBalance     As Long
'    AfterGlowBalance As Long
'    BlurBalance      As Long
'
'    StripesIntensity As Long
'
'    Opaque As Boolean
'End Type

'Type MARGINS
'    cxLeftWidth    As Long
'    cxRightWidth   As Long
'    cyTopHeight    As Long
'    cyBottomHeight As Long
'End Type

Enum AudioFormat
    Auto = 0
    MP3 = 1
    WAV = 2
    FLAC = 3
End Enum

Enum AudioBitrateType
    VBR = 0
    CBR = 1
End Enum

'Type CWPSTRUCT
'    lParam As Long
'    wParam As Long
'    Message As Long
'    hWnd As Long
'End Type

Public Const GWL_EXSTYLE As Long = -20&

Public Const RGN_DIFF As Long = 4&
Public Const RGN_OR As Long = 2&

Public Const WS_VISIBLE As Long = &H10000000
Public Const WS_VSCROLL As Long = &H200000
Public Const WS_TABSTOP As Long = &H10000
Public Const WS_THICKFRAME As Long = &H40000
Public Const WS_MAXIMIZE As Long = &H1000000
Public Const WS_MAXIMIZEBOX As Long = &H10000
Public Const WS_MINIMIZE As Long = &H20000000
Public Const WS_MINIMIZEBOX As Long = &H20000
Public Const WS_SYSMENU As Long = &H80000
Public Const WS_BORDER As Long = &H800000
Public Const WS_CAPTION As Long = &HC00000
Public Const WS_CHILD As Long = &H40000000
Public Const WS_CHILDWINDOW As Long = (WS_CHILD)
Public Const WS_CLIPCHILDREN As Long = &H2000000
Public Const WS_CLIPSIBLINGS As Long = &H4000000
Public Const WS_DISABLED As Long = &H8000000
Public Const WS_DLGFRAME As Long = &H400000
Public Const WS_EX_ACCEPTFILES As Long = &H10&
Public Const WS_EX_DLGMODALFRAME As Long = &H1&
Public Const WS_EX_NOPARENTNOTIFY As Long = &H4&
Public Const WS_EX_TOPMOST As Long = &H8&
Public Const WS_EX_TRANSPARENT As Long = &H20&
Public Const WS_EX_WINDOWEDGE As Long = &H100&
Public Const WS_EX_CLIENTEDGE As Long = &H200&
Public Const WS_EX_STATICEDGE As Long = &H20000
Public Const WS_GROUP As Long = &H20000
Public Const WS_HSCROLL As Long = &H100000
Public Const WS_ICONIC As Long = WS_MINIMIZE
Public Const WS_OVERLAPPED As Long = &H0&

Private Type ItemID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As ItemID
End Type

Enum DriveTypes
    DRIVE_UNKNOWN = 0
    DRIVE_NO_ROOT_DIR = 1
    DRIVE_REMOVABLE = 2
    DRIVE_FIXED = 3
    DRIVE_REMOTE = 4
    DRIVE_CDROM = 5    'can be a CD or a DVD
    DRIVE_RAMDISK = 6
End Enum


Private Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Private Const SW_SHOW = 5
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    ' optional fields
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Public Const CSIDL_DESKTOP = &H0
Public Const CSIDL_INTERNET = &H1
Public Const CSIDL_PROGRAMS = &H2
Public Const CSIDL_CONTROLS = &H3
Public Const CSIDL_PRINTERS = &H4
Public Const CSIDL_PERSONAL = &H5
Public Const CSIDL_FAVORITES = &H6
Public Const CSIDL_STARTUP = &H7
Public Const CSIDL_RECENT = &H8
Public Const CSIDL_SENDTO = &H9
Public Const CSIDL_BITBUCKET = &HA
Public Const CSIDL_STARTMENU = &HB
Public Const CSIDL_DESKTOPDIRECTORY = &H10
Public Const CSIDL_DRIVES = &H11
Public Const CSIDL_NETWORK = &H12
Public Const CSIDL_NETHOOD = &H13
Public Const CSIDL_FONTS = &H14
Public Const CSIDL_TEMPLATES = &H15
Public Const CSIDL_COMMON_STARTMENU = &H16
Public Const CSIDL_COMMON_PROGRAMS = &H17
Public Const CSIDL_COMMON_STARTUP = &H18
Public Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19
Public Const CSIDL_APPDATA = &H1A
Public Const CSIDL_PRINTHOOD = &H1B
Public Const CSIDL_ALTSTARTUP = &H1D
Public Const CSIDL_COMMON_ALTSTARTUP = &H1E
Public Const CSIDL_COMMON_FAVORITES = &H1F
Public Const CSIDL_INTERNET_CACHE = &H20
Public Const CSIDL_COOKIES = &H21
Public Const CSIDL_HISTORY = &H22

Public Const SC_MOVE = &HF010&
Public Const SC_RESTORE = &HF120&
Public Const SC_SIZE = &HF000&
Public Const SC_CLOSE = &HF060&
Global Const MF_BYPOSITION = &H400
Global Const MF_BYCOMMAND = &H0&
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Const ERROR_SUCCESS = 0
Const REG_SZ = 1
Const REG_DWORD = 4

Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
Const KEY_READ = &H20019

Public SessionHeaders As Collection
Public SessionHeaderKeys As Collection
Public HeaderCache As String
Public SessionHeaderCache As String

Private Type OSVERSIONINFO
    OSVSize         As Long
    dwVerMajor      As Long
    dwVerMinor      As Long
    dwBuildNumber   As Long
    PlatformID      As Long
    szCSDVersion    As String * 128
End Type

Public Const MIIM_STATE = &H1
Public Const MIIM_ID = &H2
Public Const MIIM_TYPE = &H10
Public Const MFT_SEPARATOR = &H800
Public Const MFT_STRING = &H0
Public Const MFS_ENABLED = &H0
Public Const MFS_GRAYED = &H3
Public Const MFS_DISABLED = MFS_GRAYED
Public Const MFS_CHECKED = &H8

Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hBmpChecked As Long
    hBmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Private Const VER_PLATFORM_WIN32s = 0&
Private Const VER_PLATFORM_WIN32_WINDOWS = 1&
Private Const VER_PLATFORM_WIN32_NT = 2&

Type ChooseColorStruct
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    RGBResult As Long
    lpCustColors As Long
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Const CC_RGBINIT = &H1&
Const CC_FULLOPEN = &H2&
Const CC_PREVENTFULLOPEN = &H4&
Const CC_SHOWHELP = &H8&
Const CC_ENABLEHOOK = &H10&
Const CC_ENABLETEMPLATE = &H20&
Const CC_ENABLETEMPLATEHANDLE = &H40&
Const CC_SOLIDCOLOR = &H80&
Const CC_ANYCOLOR = &H100&
Const CLR_INVALID = &HFFFF

'Enum MsgBoxExIcon
'    Critical = 16
'    Question = 32
'    Exclamation = 48
'    Information = 64
'    Doraemon = 128
'End Enum

Private Const VK_SHIFT As Long = &H10&
Private Const VK_CONTROL As Long = &H11&
Private Const VK_MENU As Long = &H12&
Private Const VK_CAPITAL As Long = &H14&
Private Const VK_NUMLOCK As Long = &H90&
Private Const VK_SCROLL As Long = &H91&

Enum GetKeyStateKeyboardCodes
    gksKeyboardShift = VK_SHIFT
    gksKeyboardctrl = VK_CONTROL
    gksKeyboardAlt = VK_MENU
    gksKeyboardCapsLock = VK_CAPITAL
    gksKeyboardNumLock = VK_NUMLOCK
    gksKeyboardScrollLock = VK_SCROLL
End Enum

'https://www.mrexcel.com/board/threads/test-if-shift-key-was-held-when-commandbutton-gets-clicked.194874/
Function IsKeyPressed(lKey As GetKeyStateKeyboardCodes) As Boolean
    Dim iResult As Integer
    iResult = GetKeyState(lKey)

    Select Case lKey
        Case gksKeyboardCapsLock, gksKeyboardNumLock, gksKeyboardScrollLock
            iResult = iResult And 1
        Case Else
            iResult = iResult And &H8000
    End Select

    IsKeyPressed = iResult
End Function

Sub DisableDWMWindow(hWnd As Long)
    On Error Resume Next
    DwmSetWindowAttribute hWnd, 2&, 1&, 4&
End Sub

Sub EnableDWMWindow(hWnd As Long)
    On Error Resume Next
    DwmSetWindowAttribute hWnd, 2&, 0&, 4&
End Sub

Function IsDWMAvailable() As Boolean
    On Error GoTo noavail
    DwmSetWindowAttribute 0&, 0&, 0&, 0&
    IsDWMAvailable = True
noavail:
End Function

Function IsDWMEnabled() As Boolean
    On Error GoTo nodwm
    Dim DwmEnabled&, ret&
    ret = DwmIsCompositionEnabled(DwmEnabled)
    If ret = 0& Then IsDWMEnabled = (DwmEnabled <> 0&)
nodwm:
End Function

'Sub ExtendDWMFrame(ByRef frmForm As Form, Top As Long, Right As Long, Bottom As Long, Left As Long)
'    On Error Resume Next
'    Dim Margin As MARGINS
'    Margin.cxLeftWidth = Left
'    Margin.cxRightWidth = Right
'    Margin.cyTopHeight = Top
'    Margin.cyBottomHeight = Bottom
'    DwmExtendFrameIntoClientArea frmForm.hWnd, Margin
'End Sub

Sub SetFormBackgroundColor(frmForm As Form, Optional DisableClassicTheme As Boolean = False)
    Dim clrBackColor As Long
    Dim clrForeColor As Long
    Dim DisableVisualStyle As Boolean
    Dim RoundButton As Boolean
    DisableVisualStyle = CBool(CInt(GetSetting("DownloadBooster", "Options", "DisableVisualStyle", 0)))
    clrBackColor = GetSetting("DownloadBooster", "Options", "BackColor", DefaultBackColor)
    RoundButton = (GetSetting("DownloadBooster", "Options", "RoundClassicButtons", 0) <> 0)
    If clrBackColor < 0 Or clrBackColor > 16777215 Then
        If frmForm.BackColor <> &H8000000F Then frmForm.BackColor = &H8000000F
        clrBackColor = &H8000000F
    ElseIf GetSetting("DownloadBooster", "Options", "BackColorMainOnly", 0) <> 0 And (Not frmForm Is frmMain) Then
        frmForm.BackColor = &H8000000F
        clrBackColor = &H8000000F
    Else
        frmForm.BackColor = clrBackColor
    End If
    Dim IsSystemColor As Boolean
    clrForeColor = GetSetting("DownloadBooster", "Options", "ForeColor", -1)
    IsSystemColor = (clrForeColor = -1)
    If clrForeColor < 0 Or clrForeColor > 16777215 Then
        If frmForm.ForeColor <> &H80000012 Then frmForm.ForeColor = &H80000012
        clrForeColor = &H80000012
    ElseIf GetSetting("DownloadBooster", "Options", "ForeColorMainOnly", 0) <> 0 And (Not frmForm Is frmMain) Then
        frmForm.ForeColor = &H80000012
        clrForeColor = &H80000012
        IsSystemColor = True
    Else
        frmForm.ForeColor = clrForeColor
    End If

    On Error Resume Next
    Dim ctrl As Control
    For Each ctrl In frmForm.Controls
        If TypeOf ctrl Is CheckBox Or TypeOf ctrl Is Frame Or TypeOf ctrl Is DriveListBox Or TypeOf ctrl Is FileListBox Or TypeOf ctrl Is DirListBox Or TypeOf ctrl Is TextBox Or TypeOf ctrl Is ComboBox Or TypeOf ctrl Is ImageCombo Or TypeOf ctrl Is ToolBar Or TypeOf ctrl Is PictureBox Or TypeOf ctrl Is Label Or TypeOf ctrl Is TabStrip Or TypeOf ctrl Is Slider Or TypeOf ctrl Is OptionButton Or TypeOf ctrl Is ProgressBar Or TypeOf ctrl Is FrameW Or TypeOf ctrl Is CommandButton Or TypeOf ctrl Is CommandButtonW Or TypeOf ctrl Is CheckBoxW Or TypeOf ctrl Is StatusBar Or TypeOf ctrl Is ListView Or TypeOf ctrl Is ListBox Then
            If TypeOf ctrl Is CommandButtonW And ctrl.Tag <> "notygchange" Then
                ctrl.IsTygemButton = CurrentButtonSkin > 0
                If CurrentButtonSkin = 0 Then ctrl.Refresh Else ctrl.GetTygemButton().Skin = CurrentButtonSkin
            End If
            If ctrl.Tag <> "novisualstylechange" And ctrl.Tag <> "nobackcolorchange novisualstylechange" Then
                If TypeOf ctrl Is CommandButton Or TypeOf ctrl Is DriveListBox Or TypeOf ctrl Is FileListBox Or TypeOf ctrl Is DirListBox Or TypeOf ctrl Is TextBox Or TypeOf ctrl Is ComboBox Or TypeOf ctrl Is CheckBox Or TypeOf ctrl Is Frame Then
                    If DisableVisualStyle Then
                        RemoveVisualStyles ctrl.hWnd
                        'If TypeOf ctrl Is CommandButton Then ctrl.Style = 1
                    Else
                        ActivateVisualStyles ctrl.hWnd
                        'If TypeOf ctrl Is CommandButton Then ctrl.Style = 0
                    End If
                Else
                    If (Not DisableVisualStyle) Then
                        If ctrl.Tag <> "nocolorchange" And ctrl.Tag <> "nocolorsizechange" And (Not frmForm Is frmOptions) And (Not frmForm Is frmDownloadOptions) And (Not IsSystemColor) And (TypeOf ctrl Is FrameW Or TypeOf ctrl Is CheckBoxW Or TypeOf ctrl Is OptionButton) Then
                            RemoveVisualStyles ctrl.hWnd
                            ctrl.VisualStyles = False
                            ctrl.RoundButton = RoundButton
                        Else
                            If Not (TypeOf ctrl Is PictureBox) Then ActivateVisualStyles ctrl.hWnd
                            ctrl.VisualStyles = True
                        End If
                    ElseIf Not (TypeOf ctrl Is PictureBox) Then
                        RemoveVisualStyles ctrl.hWnd
                        ctrl.VisualStyles = False
                        ctrl.RoundButton = RoundButton
                    End If
                End If
            End If
            If TypeOf ctrl Is DriveListBox Or TypeOf ctrl Is FileListBox Or TypeOf ctrl Is DirListBox Or TypeOf ctrl Is TextBox Or TypeOf ctrl Is ComboBox Or TypeOf ctrl Is ListView Or TypeOf ctrl Is ListBox Then GoTo nextfor
            If ctrl.Tag <> "nocolorchange" And ctrl.Tag <> "nocolorsizechange" And ctrl.ForeColor <> clrForeColor And (Not frmForm Is frmOptions) And (Not frmForm Is frmDownloadOptions) Then
                ctrl.ForeColor = clrForeColor
                If (Not IsSystemColor) And (TypeOf ctrl Is FrameW Or TypeOf ctrl Is CheckBoxW Or TypeOf ctrl Is OptionButton) Then
                    If Not (TypeOf ctrl Is PictureBox) Then RemoveVisualStyles ctrl.hWnd
                    ctrl.VisualStyles = False
                ElseIf (Not DisableVisualStyle) And ctrl.VisualStyles = False And ctrl.Tag <> "novisualstylechange" And ctrl.Tag <> "nobackcolorchange novisualstylechange" Then
                    If Not (TypeOf ctrl Is PictureBox) Then ActivateVisualStyles ctrl.hWnd
                    ctrl.VisualStyles = True
                End If
            End If
            If TypeOf ctrl Is PictureBox And (Not ctrl.Tag = "forcebgchange") Then
                If ctrl.AutoRedraw = True Then GoTo nextfor
            End If
            If ctrl.Tag <> "nobackcolorchange" And ctrl.Tag <> "nobackcolorchange novisualstylechange" And ctrl.BackColor <> clrBackColor Then
                ctrl.BackColor = clrBackColor
                If TypeOf ctrl Is CheckBoxW Or TypeOf ctrl Is FrameW Then
                    If ctrl.Transparent Then ctrl.Refresh
                End If
            End If
        End If
nextfor:
    Next ctrl

    SetClassicTheme frmForm, DisableClassicTheme
End Sub

Sub SetClassicTheme(frmForm As Form, Optional DisableClassicTheme As Boolean = False)
    If GetSetting("DownloadBooster", "Options", "UseClassicThemeFrame", 0) <> 0 Then
        RemoveVisualStyles frmForm.hWnd
    ElseIf DisableClassicTheme Then
        ActivateVisualStyles frmForm.hWnd
    End If
End Sub

Function ShowColorDialog(Optional ByVal hParent As Long, Optional ByVal bFullOpen As Boolean, Optional ByVal InitColor As OLE_COLOR, Optional ByVal SolidOnly As Boolean = False) As Long
    Dim CC As ChooseColorStruct
    Static aColorRef(15) As Long
    Dim lInitColor As Long

    If InitColor <> 0 Then
        If OleTranslateColor(InitColor, 0, lInitColor) Then
            lInitColor = CLR_INVALID
        End If
    End If

    If Not aColorRef(0) Then
        aColorRef(0) = RGB(233, 245, 236)
        aColorRef(1) = RGB(233, 237, 243)
        aColorRef(2) = RGB(185, 209, 234)
        aColorRef(3) = RGB(235, 233, 245)
        aColorRef(4) = RGB(252, 251, 224)
        aColorRef(5) = RGB(244, 232, 232)
        aColorRef(6) = RGB(248, 228, 244)
        aColorRef(7) = RGB(223, 233, 244)

        aColorRef(8) = RGB(249, 242, 230)
        aColorRef(9) = RGB(222, 235, 248)
        aColorRef(10) = RGB(227, 244, 232)
        aColorRef(11) = RGB(236, 230, 211)
        aColorRef(12) = RGB(212, 208, 200)
        aColorRef(13) = RGB(192, 192, 192)
        aColorRef(14) = 16777215
        aColorRef(15) = 0&
    End If

    Dim SolidColor As Long
    If SolidOnly Then
        SolidColor = CC_SOLIDCOLOR
    Else
        SolidColor = 0&
    End If

    With CC
        .lStructSize = Len(CC)
        .hWndOwner = hParent
        .lpCustColors = VarPtr(aColorRef(0))
        .RGBResult = lInitColor
        .Flags = SolidColor Or CC_ANYCOLOR Or CC_RGBINIT Or IIf(bFullOpen, CC_FULLOPEN, 0)
    End With

    If ChooseColor(CC) Then
        ShowColorDialog = CC.RGBResult
    Else
        ShowColorDialog = -1
    End If
End Function

Function GetKeyValue(ByVal KeyRoot As Long, KeyName As String, SubKeyRef As String, Optional Default As String = "") As Variant
    Dim i As Long                                           ' 루프 카운터
    Dim RC As Long                                          ' 반환 코드
    Dim hKey As Long                                        ' 열려 있는 레지스트리 키 처리
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' 레지스트리 키의 데이터 형식
    Dim tmpVal As String                                    ' 레지스트리 키 값을 임시로 저장
    Dim KeyValSize As Long                                  ' 레지스트리 키 변수의 크기
    Dim KeyVal
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    RC = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' 레지스트리 키를 엽니다.

    If (RC <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 오류를 처리합니다...

    tmpVal = String$(1024, 0)                             ' 변수의 크기를 할당합니다.
    KeyValSize = 1024                                       ' 변수 크기를 표시합니다.

    '------------------------------------------------------------
    ' 레지스트리 키 값을 읽어옵니다...
    '------------------------------------------------------------
    RC = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' 키 값을 가져오고 작성합니다.

    If (RC <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 오류를 처리합니다.

    If (Asc(Mid$(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95는 Null 종료 문자열을 추가합니다...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null을 찾았습니다. 문자열에서 추출합니다.
    Else                                                    ' WinNT는 Null 종료 문자열 추가하지 않습니다...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null을 찾지 못했습니다. 문자열에서만 추출합니다.
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' 데이터 형식을 검색합니다.
    Case REG_DWORD                                          ' 이진 단어 레지스트리 키 데이터 형식
        For i = Len(tmpVal) To 1 Step -1                    ' 각각 비트를 변환합니다.
            KeyVal = KeyVal + Hex(Asc(Mid$(tmpVal, i, 1)))   ' 값 문자를 문자별로 작성합니다.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' 이진 단어를 문자열로 변환합니다.
    Case Else                                               ' 문자열 레지스트리 키 데이터 형식
        KeyVal = tmpVal                                     ' 문자열 값을 복사합니다.
    End Select

    GetKeyValue = KeyVal
    RC = RegCloseKey(hKey)                                  ' 레지스트리 키를 닫습니다.
    Exit Function                                           ' 종료합니다.

GetKeyError:      ' 오류가 발생하면 지웁니다...
    GetKeyValue = Default
    RC = RegCloseKey(hKey)                                  ' 레지스트리 키를 닫습니다.
End Function

'https://www.vbforums.com/showthread.php?796771-RESOLVED-Help!-cannot-delete-registry-x64-subkeys&p=4894805
Function GetSubkeys(ByVal KeyRoot As Long, KeyName As String) As String()
    Dim KeysRev() As String, Keys() As String
    Dim hKey&, i&, j&, nBufferLen&, sBuffer$

    If RegOpenKeyEx(KeyRoot, KeyName, 0&, KEY_READ, hKey) <> ERROR_SUCCESS Then GoTo keyerr
    If RegQueryInfoKey(hKey, lpcSubKeys:=i, lpcMaxSubKeyLen:=nBufferLen) <> ERROR_SUCCESS Then GoTo keyerr
    SysReAllocStringLen VarPtr(sBuffer), Length:=nBufferLen
    ReDim KeysRev(0 To i - 1) As String
    j = 0&
    For i = i - 1& To 0& Step -1&
        nBufferLen = Len(sBuffer) + 1&
        If RegEnumKeyEx(hKey, i, StrPtr(sBuffer), nBufferLen) = ERROR_SUCCESS Then
            KeysRev(j) = Left$(sBuffer, nBufferLen)
            j = j + 1&
        End If
    Next i
    If j > 0& Then
        j = j - 1&
        ReDim Keys(0 To j) As String
        For i = j To 0& Step -1&
            Keys(j - i) = KeysRev(i)
        Next i
        GetSubkeys = Keys
        RegCloseKey hKey
        Exit Function
    End If

keyerr:
    RegCloseKey hKey
    GetSubkeys = Keys
End Function

Function Exists(oCol As Collection, vKey As String) As Boolean
    On Error Resume Next
    Err.Clear
    oCol.Item vKey
    Exists = (Err.Number = 0)
End Function

Function TextWidth(s As String, Optional ByVal FontName As String = "", Optional ByVal FontSize As Integer = -1, Optional FontBold As Boolean = False) As Single
    If FontSize = 0 Then
        TextWidth = 0
        Exit Function
    End If
    Dim UserFontName$
    UserFontName = Trim$(GetSetting("DownloadBooster", "Options", "Font", ""))
    On Error Resume Next
    If LangID <> 1042 Then
        If FontName = "" Then
            FontName = IIf(UserFontName = "", "Tahoma", UserFontName)
            If FontSize = -1 Then
                If LCase(FontName) = "tahoma" Then
                    FontSize = 8
                Else
                    FontSize = 9
                End If
            End If
        ElseIf FontSize = -1 Then
            FontSize = 9
        End If
    Else
        If FontName = "" Then FontName = IIf(UserFontName = "", DefaultFont, UserFontName)
        If FontSize = -1 Then
            If LCase(FontName) = "tahoma" Then
                FontSize = 8
            Else
                FontSize = 9
            End If
        End If
    End If
    frmDummyForm.Font.Name = FontName
    frmDummyForm.Font.Size = FontSize
    frmDummyForm.Font.Bold = FontBold
    TextWidth = frmDummyForm.TextWidth(s)
End Function

Function TextHeight(s As String, Optional ByVal FontName As String = "", Optional ByVal FontSize As Integer = -1, Optional FontBold As Boolean = False) As Single
    If FontSize = 0 Then
        TextHeight = 0
        Exit Function
    End If
    Dim UserFontName$
    UserFontName = Trim$(GetSetting("DownloadBooster", "Options", "Font", ""))
    On Error Resume Next
    If LangID <> 1042 Then
        If FontName = "" Then
            FontName = IIf(UserFontName = "", "Tahoma", UserFontName)
            If FontSize = -1 Then
                If LCase(FontName) = "tahoma" Then
                    FontSize = 8
                Else
                    FontSize = 9
                End If
            End If
        ElseIf FontSize = -1 Then
            FontSize = 9
        End If
    Else
        If FontName = "" Then FontName = IIf(UserFontName = "", DefaultFont, UserFontName)
        If FontSize = -1 Then
            If LCase(FontName) = "tahoma" Then
                FontSize = 8
            Else
                FontSize = 9
            End If
        End If
    End If
    frmDummyForm.Font.Name = FontName
    frmDummyForm.Font.Size = FontSize
    frmDummyForm.Font.Bold = FontBold
    TextHeight = frmDummyForm.TextHeight(s)
End Function

Function StrLen(s As String) As Integer
    StrLen = LenB(StrConv(s, vbFromUnicode))
End Function

Private Function CutLines(Text As String, ByVal Width As Single) As String()
    Dim Paragraphs() As String
    Dim ParagraphX As Long
    Dim Words() As String
    Dim WordX As Long
    Dim CutLine As String
    Dim NewCutLine As String
    Dim SingleWord As Boolean
    Dim ForceX As Long
    Dim Lines() As String
    Dim LineX As Long

    Paragraphs = Split(Text, vbNewLine)
    For ParagraphX = 0 To UBound(Paragraphs)
        Words = Split(Paragraphs(ParagraphX), " ")
        WordX = 0
        Do While WordX <= UBound(Words)
            Do
                If Len(CutLine) = 0 Then
                    NewCutLine = Words(WordX)
                    SingleWord = True
                Else
                    NewCutLine = NewCutLine & " " & Words(WordX)
                End If
                If TextWidth(NewCutLine) > Width Then Exit Do
                CutLine = NewCutLine
                WordX = WordX + 1
                SingleWord = False
            Loop While WordX <= UBound(Words)
            If SingleWord Then
                For ForceX = Len(Words(WordX)) - 1 To 1 Step -1
                    CutLine = Left$(Words(WordX), ForceX)
                    If TextWidth(CutLine) <= Width Then
                        Words(WordX) = Mid$(Words(WordX), ForceX + 1)
                        Exit For
                    End If
                Next
            End If
            ReDim Preserve Lines(LineX)
            Lines(LineX) = CutLine
            LineX = LineX + 1
            CutLine = vbNullString
        Loop
    Next
    CutLines = Lines
End Function

Function InputBoxEx(Prompt As String, Optional Title As String, Optional Default As String)
    If Title = "" Then Title = App.Title

    Dim InpBox As frmInputBox
    Set InpBox = New frmInputBox
    InpBox.ResultID = CStr(Rnd * 1E+15)
    Set InpBox.InputBoxObject = InpBox

    On Error Resume Next

    InpBox.cmdOK.Caption = t("확인", "OK")
    InpBox.cmdCancel.Caption = t("취소", "Cancel")

    InpBox.lblCaption = Prompt
    InpBox.txtInput.Text = Default
    InpBox.Caption = Title

    InpBox.Show vbModal
    InputBoxEx = InputBoxResults(InpBox.ResultID)
    InputBoxResults.Remove InpBox.ResultID
    Unload InpBox
    Set InpBox = Nothing
End Function

Function RandInt(StartNumber, EndNumber)
    RandInt = Int(Rnd * (EndNumber - StartNumber + 1)) + StartNumber
End Function

Function ShowMessageBox(ByVal Content As String, Optional ByVal Title As String, Optional Icon As VbMsgBoxStyleEx = 64, Optional IsModal As Boolean = True, Optional AlertTimeout As Integer = -1, Optional ByVal DefaultOption As VbMsgBoxResultEx = vbNo, Optional ByVal MsgBoxMode As VbMsgBoxStyleEx = vbOKOnly) As VbMsgBoxResultEx
    If Title = "" Then Title = App.Title
'    If GetSetting("DownloadBooster", "Options", "ForceNativeMessageBox", 0) <> 0 And MsgBoxMode <> vbYesNoEx And MsgBoxMode <> vbCancelTryContinue Then
'        ShowMessageBox = VBA.MsgBox(Content, Icon + MsgBoxMode, Title)
'        Exit Function
'    End If

    Dim MessageBox As frmMessageBox
    Set MessageBox = New frmMessageBox
    MessageBox.MsgBoxMode = MsgBoxMode
    MessageBox.ResultID = CStr(Rnd * 1E+15)
    Set MessageBox.MessageBoxObject = MessageBox

    On Error Resume Next
    'Dim NoIcon As Boolean
    'NoIcon = False
    
    Dim IconIndex As Byte
    Dim RandVal As Integer
    RandVal = RandInt(0, 999)
    If RandVal Then IconIndex = (RandVal Mod 2) + 1 Else IconIndex = 3
    Set MessageBox.imgTrain.Picture = Train(IconIndex)
    MessageBox.imgIcon(Icon / 16).Visible = True

    Content = Replace(Content, "&", "&&")
    Content = Replace(Content, vbCrLf & vbCrLf, vbCrLf & " " & vbCrLf)

    Dim i%
    Dim LineCount As Integer
    Dim LContent As Integer
    Dim MAX_WIDTH As Long
    MAX_WIDTH = Screen.Width / 2
    Content = Join(CutLines(Content, MAX_WIDTH), vbCrLf)
    LContent = 0
    LineCount = UBound(Split(Content, vbLf)) + 1
    Dim s%
    Dim ln$
    Dim CI%, c$
    Dim LineContent$
    For s = 0 To UBound(Split(Content, vbCrLf))
        LineContent = Split(Content, vbCrLf)(s)
        If TextWidth(LineContent) > LContent Then LContent = TextWidth(LineContent)
    Next s

    If LContent = 0 Then LContent = StrLen(Content)
    If LineCount > 1 Then MessageBox.lblContent.Top = 280

    Dim MsgBoxMinWidth As Integer
    Select Case MsgBoxMode
        Case vbYesNo, vbRetryCancel, vbOKCancel, vbYesNoEx
            MsgBoxMinWidth = 3480
        Case vbYesNoCancel, vbAbortRetryIgnore, vbCancelTryContinue
            MsgBoxMinWidth = 4920
        Case Else 'vbOKOnly
            MsgBoxMinWidth = 1920
    End Select

    'MessageBox.lblContent.Height = 185 * LineCount + 60
    MessageBox.Height = 1615 + LineCount * 180 - 300 + 190 - 60 + IIf(MsgBoxMode = vbYesNoEx, 735, 0)
    MessageBox.Caption = Title
    MessageBox.lblContent.Caption = Content
    MessageBox.Width = Max(2040 + LContent - 640 - 225, MsgBoxMinWidth)
'    If NoIcon Then
'        MessageBox.Width = Max(MessageBox.Width - 720, MsgBoxMinWidth)
'        MessageBox.lblContent.Top = MessageBox.lblContent.Top - 180
'        MessageBox.lblContent.Left = MessageBox.lblContent.Left - 720
'        MessageBox.Height = MessageBox.Height - 240
'    End If

    Select Case MsgBoxMode
        Case vbYesNo
            MessageBox.cmdYes.Left = MessageBox.Width / 2 - 810 - MessageBox.cmdYes.Width / 2
            MessageBox.cmdYes.Top = 840 + (LineCount * 185) - 350
            MessageBox.cmdNo.Left = MessageBox.Width / 2 - 810 - MessageBox.cmdYes.Width / 2 - 120 + MessageBox.cmdYes.Width + 240 - 30
            MessageBox.cmdNo.Top = 840 + (LineCount * 185) - 350
            If LineCount < 2 Then
                MessageBox.Height = MessageBox.Height + 180
                MessageBox.cmdYes.Top = MessageBox.cmdYes.Top + 180
                MessageBox.cmdNo.Top = MessageBox.cmdNo.Top + 180
            End If
'            If NoIcon Then
'                MessageBox.cmdYes.Top = MessageBox.cmdYes.Top - 210
'                MessageBox.cmdNo.Top = MessageBox.cmdNo.Top - 210
'            End If
        Case vbYesNoEx
            MessageBox.cmdOK.Left = MessageBox.Width / 2 - 810 - MessageBox.cmdOK.Width / 2
            MessageBox.cmdOK.Top = 840 + (LineCount * 185) - 350 + 705
            MessageBox.cmdCancel.Left = MessageBox.Width / 2 - 810 - MessageBox.cmdOK.Width / 2 - 120 + MessageBox.cmdOK.Width + 240 - 30
            MessageBox.cmdCancel.Top = 840 + (LineCount * 185) - 350 + 705
            MessageBox.optYes.Top = MessageBox.cmdOK.Top - 620
            MessageBox.optNo.Top = MessageBox.cmdOK.Top - 320
            If LineCount > 1 Then
                MessageBox.optYes.Top = MessageBox.optYes.Top - 80
                MessageBox.optNo.Top = MessageBox.optNo.Top - 80
            End If
            If IsEmpty(DefaultOption) Then
                MessageBox.optYes.Value = False
                MessageBox.optNo.Value = False
                MessageBox.cmdOK.Enabled = False
            ElseIf DefaultOption = vbYes Then
                MessageBox.optYes.Value = True
                MessageBox.cmdOK.Enabled = True
            Else
                MessageBox.optNo.Value = True
                MessageBox.cmdOK.Enabled = True
            End If
            If LineCount < 2 Then
                MessageBox.Height = MessageBox.Height + 180
                MessageBox.cmdOK.Top = MessageBox.cmdOK.Top + 180
                MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top + 180
            End If
'            If NoIcon Then
'                MessageBox.cmdOK.Top = MessageBox.cmdOK.Top - 210
'                MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top - 210
'                MessageBox.optYes.Top = MessageBox.optYes.Top - 210
'                MessageBox.optNo.Top = MessageBox.optNo.Top - 210
'            End If
        Case vbYesNoCancel
            MessageBox.cmdYes.Left = MessageBox.Width / 2 - 900 - MessageBox.cmdYes.Width
            MessageBox.cmdYes.Top = 840 + (LineCount * 185) - 350
            MessageBox.cmdNo.Left = MessageBox.Width / 2 - 810 + 15
            MessageBox.cmdNo.Top = 840 + (LineCount * 185) - 350
            MessageBox.cmdCancel.Left = MessageBox.Width / 2 - 900 + MessageBox.cmdYes.Width + 190 + 30
            MessageBox.cmdCancel.Top = 840 + (LineCount * 185) - 350
            If LineCount < 2 Then
                MessageBox.Height = MessageBox.Height + 180
                MessageBox.cmdYes.Top = MessageBox.cmdYes.Top + 180
                MessageBox.cmdNo.Top = MessageBox.cmdNo.Top + 180
                MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top + 180
            End If
'            If NoIcon Then
'                MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top - 210
'                MessageBox.cmdYes.Top = MessageBox.cmdYes.Top - 210
'                MessageBox.cmdNo.Top = MessageBox.cmdNo.Top - 210
'            End If
        Case vbRetryCancel
            MessageBox.cmdRetry.Left = MessageBox.Width / 2 - 810 - MessageBox.cmdRetry.Width / 2
            MessageBox.cmdRetry.Top = 840 + (LineCount * 185) - 350
            MessageBox.cmdCancel.Left = MessageBox.Width / 2 - 810 - MessageBox.cmdCancel.Width / 2 - 120 + MessageBox.cmdRetry.Width + 240 - 30
            MessageBox.cmdCancel.Top = 840 + (LineCount * 185) - 350
            If LineCount < 2 Then
                MessageBox.Height = MessageBox.Height + 180
                MessageBox.cmdRetry.Top = MessageBox.cmdRetry.Top + 180
                MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top + 180
            End If
'            If NoIcon Then
'                MessageBox.cmdRetry.Top = MessageBox.cmdRetry.Top - 210
'                MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top - 210
'            End If
#If False Then
        Case vbAbortRetryIgnore
            MessageBox.cmdAbort.Left = MessageBox.Width / 2 - 900 - MessageBox.cmdAbort.Width
            MessageBox.cmdAbort.Top = 840 + (LineCount * 185) - 350
            MessageBox.cmdRetry.Left = MessageBox.Width / 2 - 810 + 15
            MessageBox.cmdRetry.Top = 840 + (LineCount * 185) - 350
            MessageBox.cmdIgnore.Left = MessageBox.Width / 2 - 900 + MessageBox.cmdAbort.Width + 190 + 30
            MessageBox.cmdIgnore.Top = 840 + (LineCount * 185) - 350
            If LineCount < 2 Then
                MessageBox.Height = MessageBox.Height + 180
                MessageBox.cmdAbort.Top = MessageBox.cmdAbort.Top + 180
                MessageBox.cmdRetry.Top = MessageBox.cmdRetry.Top + 180
                MessageBox.cmdIgnore.Top = MessageBox.cmdIgnore.Top + 180
            End If
'            If NoIcon Then
'                MessageBox.cmdIgnore.Top = MessageBox.cmdIgnore.Top - 210
'                MessageBox.cmdAbort.Top = MessageBox.cmdAbort.Top - 210
'                MessageBox.cmdRetry.Top = MessageBox.cmdRetry.Top - 210
'            End If
        Case vbOKCancel
            MessageBox.cmdOK.Left = MessageBox.Width / 2 - 810 - MessageBox.cmdOK.Width / 2
            MessageBox.cmdOK.Top = 840 + (LineCount * 185) - 350
            MessageBox.cmdCancel.Left = MessageBox.Width / 2 - 810 - MessageBox.cmdCancel.Width / 2 - 120 + MessageBox.cmdOK.Width + 240 - 30
            MessageBox.cmdCancel.Top = 840 + (LineCount * 185) - 350
            If LineCount < 2 Then
                MessageBox.Height = MessageBox.Height + 180
                MessageBox.cmdOK.Top = MessageBox.cmdOK.Top + 180
                MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top + 180
            End If
'            If NoIcon Then
'                MessageBox.cmdOK.Top = MessageBox.cmdOK.Top - 210
'                MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top - 210
'            End If
        Case vbCancelTryContinue
            MessageBox.cmdCancel.Left = MessageBox.Width / 2 - 900 - MessageBox.cmdCancel.Width
            MessageBox.cmdCancel.Top = 840 + (LineCount * 185) - 350
            MessageBox.cmdTryAgain.Left = MessageBox.Width / 2 - 810 + 15
            MessageBox.cmdTryAgain.Top = 840 + (LineCount * 185) - 350
            MessageBox.cmdContinue.Left = MessageBox.Width / 2 - 900 + MessageBox.cmdCancel.Width + 190 + 30
            MessageBox.cmdContinue.Top = 840 + (LineCount * 185) - 350
            If LineCount < 2 Then
                MessageBox.Height = MessageBox.Height + 180
                MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top + 180
                MessageBox.cmdTryAgain.Top = MessageBox.cmdTryAgain.Top + 180
                MessageBox.cmdContinue.Top = MessageBox.cmdContinue.Top + 180
            End If
'            If NoIcon Then
'                MessageBox.cmdContinue.Top = MessageBox.cmdContinue.Top - 210
'                MessageBox.cmdCancel.Top = MessageBox.cmdCancel.Top - 210
'                MessageBox.cmdTryAgain.Top = MessageBox.cmdTryAgain.Top - 210
'            End If
#End If
        Case Else 'vbOKOnly
            MessageBox.cmdOK.Left = MessageBox.Width / 2 - 810 + 30
            MessageBox.cmdOK.Top = 840 + (LineCount * 185) - 350
            If LineCount < 2 Then
                MessageBox.Height = MessageBox.Height + 180
                MessageBox.cmdOK.Top = MessageBox.cmdOK.Top + 180
            End If
'            If NoIcon Then
'                MessageBox.cmdOK.Top = MessageBox.cmdOK.Top - 210
'            End If
    End Select
    
    MessageBox.lblContent.Height = MessageBox.Height

    Dim MessageSoundPath$
    Select Case Icon
        Case 48
            If GetSetting("DownloadBooster", "Options", "EnableExclamationSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "ExclamationSound", "")
            End If
        Case 16
            If GetSetting("DownloadBooster", "Options", "EnableErrorSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "ErrorSound", "")
            End If
        Case 64
            If GetSetting("DownloadBooster", "Options", "EnableAsteriskSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "AsteriskSound", "")
            End If
        Case 32
            If GetSetting("DownloadBooster", "Options", "EnableQuestionSound", 1) = 0 Then
                MessageSoundPath = "-"
            Else
                MessageSoundPath = GetSetting("DownloadBooster", "Options", "QuestionSound", "")
            End If
        Case Else
            MessageSoundPath = ""
    End Select
    If MessageSoundPath <> "-" Then PlayWave MessageSoundPath, FallbackSound:=Icon

    If MsgBoxMode = vbOKOnly And AlertTimeout >= 0 Then
        MessageBox.timeout.Interval = AlertTimeout
        MessageBox.timeout.Enabled = -1
    End If

    tr MessageBox.cmdOK, "OK"
    tr MessageBox.cmdCancel, "Cancel"
    tr MessageBox.cmdYes, "&Yes"
    tr MessageBox.cmdNo, "&No"
    tr MessageBox.cmdAbort, "&Abort"
    tr MessageBox.cmdRetry, "&Retry"
    tr MessageBox.cmdIgnore, "&Ignore"
    tr MessageBox.optYes, "&Yes"
    tr MessageBox.optNo, "&No"
    tr MessageBox.cmdTryAgain, "&Try Again"
    tr MessageBox.cmdContinue, "&Continue"

    MessageBox.cmdOK.Visible = (MsgBoxMode = vbOKOnly Or MsgBoxMode = vbYesNoEx Or MsgBoxMode = vbOKCancel)
    MessageBox.cmdCancel.Visible = (MsgBoxMode = vbYesNoEx Or MsgBoxMode = vbYesNoCancel Or MsgBoxMode = vbRetryCancel Or MsgBoxMode = vbOKCancel Or MsgBoxMode = vbCancelTryContinue)
    MessageBox.cmdYes.Visible = (MsgBoxMode = vbYesNo Or MsgBoxMode = vbYesNoCancel)
    MessageBox.cmdNo.Visible = (MsgBoxMode = vbYesNo Or MsgBoxMode = vbYesNoCancel)
    MessageBox.optYes.Visible = (MsgBoxMode = vbYesNoEx)
    MessageBox.optNo.Visible = (MsgBoxMode = vbYesNoEx)

    MessageBox.cmdAbort.Visible = (MsgBoxMode = vbAbortRetryIgnore)
    MessageBox.cmdRetry.Visible = (MsgBoxMode = vbAbortRetryIgnore Or MsgBoxMode = vbRetryCancel)
    MessageBox.cmdIgnore.Visible = (MsgBoxMode = vbAbortRetryIgnore)
    MessageBox.cmdContinue.Visible = (MsgBoxMode = vbCancelTryContinue)
    MessageBox.cmdTryAgain.Visible = (MsgBoxMode = vbCancelTryContinue)
    MessageBox.cmdHelp.Visible = False

    MessageBox.cmdCancel.Cancel = (MsgBoxMode = vbYesNoEx Or MsgBoxMode = vbYesNoCancel Or MsgBoxMode = vbRetryCancel Or MsgBoxMode = vbOKCancel Or MsgBoxMode = vbCancelTryContinue)
    MessageBox.cmdCancel.Default = False
    MessageBox.cmdYes.Cancel = False
    MessageBox.cmdYes.Default = False
    MessageBox.cmdNo.Cancel = False
    MessageBox.cmdNo.Default = False
    MessageBox.cmdOK.Cancel = (MsgBoxMode = vbOKOnly)
    MessageBox.cmdOK.Default = (MsgBoxMode = vbOKOnly Or MsgBoxMode = vbYesNoEx)

    MessageBox.Init
    If MsgBoxMode = vbOKOnly Then
        If IsModal Then
            MessageBox.Show vbModal
            Unload MessageBox
            Set MessageBox = Nothing
        Else
            MessageBox.Show
        End If
        ShowMessageBox = vbOK
    Else
        MessageBox.Show vbModal
        ShowMessageBox = MsgBoxResults(MessageBox.ResultID)
        MsgBoxResults.Remove MessageBox.ResultID
        Unload MessageBox
        Set MessageBox = Nothing
    End If
End Function

Function ConfirmEx(ByVal Content As String, Optional ByVal Title As String, Optional ByVal Icon As VbMsgBoxStyleEx = 32, Optional ByVal DefaultOption As VbMsgBoxResultEx = vbNo) As VbMsgBoxResultEx
    ConfirmEx = ShowMessageBox(Content, Title, Icon, DefaultOption:=DefaultOption, MsgBoxMode:=vbYesNoEx)
End Function

'https://www.vbforums.com/showthread.php?894947-How-to-test-if-a-font-is-available
Function FontExists(FontName As String) As Boolean
    On Error GoTo noexist
    If FontName = "" Then GoTo noexist
    With New StdFont
        .Name = FontName
        FontExists = (StrComp(.Name, FontName, vbTextCompare) = 0)
    End With
    Exit Function
noexist:
    FontExists = False
End Function

Function FolderExists(sFullPath As String) As Boolean
    On Error GoTo nonexist
    FolderExists = ((GetAttr(sFullPath) And (vbDirectory Or vbVolume)) <> 0)
    Exit Function
nonexist:
    FolderExists = False
End Function

Function Floor(ByVal floatval As Double, Optional ByVal decimalPlaces As Long) As Long
    Dim intval As Long
    intval = Round(floatval)
    If intval > floatval Then intval = intval - 1
    If decimalPlaces Then floatval = floatval / (10 ^ decimalPlaces)
    Floor = intval
End Function

Function ParseSize(ByVal Size As Double, Optional ByVal ShowBytes As Boolean = False, Optional Suffix As String = "") As String
    If Size < 0 Then
        ParseSize = "-"
        Exit Function
    End If

    On Error GoTo ErrLn4
    Dim ret#
    If Size >= 1024# * 1024# * 1024# * 1024# Then
        ret = Fix(Size / 1024# / 1024# / 1024# / 1024# * 100) / 100
        If ret >= 100# Then
            ret = Fix(ret)
        ElseIf ret >= 10# Then
            ret = Fix(ret * 10) / 10
        End If
        ParseSize = ret & "TB" & Suffix
    ElseIf Size >= 1024# * 1024# * 1024# Then
        ret = Fix(Size / 1024# / 1024# / 1024# * 100) / 100
        If ret >= 100# Then
            ret = Fix(ret)
        ElseIf ret >= 10# Then
            ret = Fix(ret * 10) / 10
        End If
        ParseSize = ret & "GB" & Suffix
    ElseIf Size >= 1024# * 1024# Then
        ret = Fix(Size / 1024# / 1024# * 100) / 100
        If ret >= 100# Then
            ret = Fix(ret)
        ElseIf ret >= 10# Then
            ret = Fix(ret * 10) / 10
        End If
        ParseSize = ret & "MB" & Suffix
    ElseIf Size >= 1024# Then
        ret = Fix(Size / 1024# * 100) / 100
        If ret >= 100# Then
            ret = Fix(ret)
        ElseIf ret >= 10# Then
            ret = Fix(ret * 10) / 10
        End If
        ParseSize = ret & "KB" & Suffix
    Else
        ParseSize = CStr(Size) & " " & t("바이트", "Bytes") & Suffix
    End If

    If Size >= (1024#) And ShowBytes Then
        ParseSize = ParseSize & " (" & Size & " " & t("바이트", "Bytes") & Suffix & ")"
    End If
    Exit Function
ErrLn4:
    ParseSize = "0 " & t("바이트", "Bytes") & Suffix
End Function

Function FilterFilename(FileName As String, Optional ByVal PreserveBackslash As Boolean) As String
    Dim Str As String
    Dim ret As String
    ret = ""
    Str = StrConv(FileName, vbProperCase)
    Dim i%
    For i = 1 To Len(Str)
        If Mid$(Str, i, 1) = "?" Then
            ret = ret & "_"
        Else
            ret = ret & Mid$(FileName, i, 1)
        End If
    Next i
    ret = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(ret, "?", "_"), "*", "_"), "|", "_"), """", "_"), ":", "_"), "<", "_"), ">", "_"), "/", "_")
    If Not PreserveBackslash Then
        ret = Replace(ret, "\", "_")
    ElseIf Mid$(ret, 2, 1) = "_" Then
        ret = Left$(ret, 1) & ":" & Right$(ret, Len(ret) - 2)
    End If
    FilterFilename = ret
End Function

'https://gist.github.com/jvarn/5e11b1fd741b5f79d8a516c9c2368f17
Function URLDecode(ByVal strIn As String, Optional Skip As Boolean = False) As String
    If Skip Then
        URLDecode = strIn
        Exit Function
    End If
    
    On Error GoTo ErrorHandler

    Dim sl As Long, tl As Long
    Dim Key As String, kl As Long
    Dim hh As String, Hi As String, hl As String
    Dim A As Long

    Key = "%"
    kl = Len(Key)
    sl = 1: tl = 1
    sl = InStr(sl, strIn, Key, vbTextCompare)
    Do While sl > 0
        If (tl = 1 And sl <> 1) Or tl < sl Then
            URLDecode = URLDecode & Mid$(strIn, tl, sl - tl)
        End If

        Select Case UCase(Mid$(strIn, sl + kl, 1))
            Case "U"
                A = val("&H" & Mid$(strIn, sl + kl + 1, 4))
                URLDecode = URLDecode & ChrW(A)
                sl = sl + 6
            Case "E"
                hh = Mid$(strIn, sl + kl, 2)
                A = val("&H" & hh)
                If A < 128 Then
                    sl = sl + 3
                    URLDecode = URLDecode & Chr(A)
                Else
                    Hi = Mid$(strIn, sl + 3 + kl, 2)
                    hl = Mid$(strIn, sl + 6 + kl, 2)
                    A = ((val("&H" & hh) And &HF) * 2 ^ 12) Or ((val("&H" & Hi) And &H3F) * 2 ^ 6) Or (val("&H" & hl) And &H3F)
                    URLDecode = URLDecode & ChrW(A)
                    sl = sl + 9
                End If
            Case Else
                hh = Mid$(strIn, sl + kl, 2)
                A = val("&H" & hh)
                If A < 128 Then
                    sl = sl + 3
                Else
                    Hi = Mid$(strIn, sl + 3 + kl, 2)
                    A = ((val("&H" & hh) - 194) * 64) + val("&H" & Hi)
                    sl = sl + 6
                End If
                URLDecode = URLDecode & ChrW(A)
        End Select

        tl = sl
        sl = InStr(sl, strIn, Key, vbTextCompare)
    Loop

    URLDecode = URLDecode & Mid$(strIn, tl)
    Exit Function

ErrorHandler:
    URLDecode = strIn
End Function

Function GetWindowsVersion() As Single
    Dim osv As OSVERSIONINFO
    osv.OSVSize = Len(osv)

    If GetVersionEx(osv) = 1 Then
        Select Case osv.PlatformID
'            Case VER_PLATFORM_WIN32s
'                GetWindowsVersion = 3.1
            Case VER_PLATFORM_WIN32_NT
                GetWindowsVersion = Round(osv.dwVerMajor + (CSng(osv.dwVerMinor) * 0.1), 1&)
                Build = osv.dwBuildNumber
                'If GetWindowsVersion >= 6.2 Then GetWindowsVersion = fWinVer()
'            Case VER_PLATFORM_WIN32_WINDOWS:
'                Select Case osv.dwVerMinor
'                    Case 0
'                        GetWindowsVersion = 4#
'                        Build = 950
'                    Case 90
'                        GetWindowsVersion = 4.9
'                        Build = 3000
'                    Case Else
'                        GetWindowsVersion = 4.1
'                        Build = 1998
'                End Select
            Case Else
                GetWindowsVersion = 3.1
        End Select
    Else
        GetWindowsVersion = 5.1
        Build = 2600
    End If
End Function

'Private Function fWinVer() As Single
'    Dim osv As OSVERSIONINFO
'    osv.OSVSize = Len(osv)
'    If GetVersionEx(osv) <> 1 Then
'        fWinVer = 5.1
'        Exit Function
'    End If
'
'    If osv.PlatformID = VER_PLATFORM_WIN32_NT Then
'        If RtlGetVersion(osv) <> 0 Then
'            fWinVer = 5.1
'            Exit Function
'        End If
'    End If
'
'    Build = osv.dwBuildNumber
'    fWinVer = Round(osv.dwVerMajor + (CSng(osv.dwVerMinor) * 0.1), 1&)
'End Function

Function t(k As String, e As String) As String
    If LangID = 1042 Then
        t = k
    Else
        t = e
    End If
End Function

Sub SetFont(frm As Form, Optional ByVal Force As Boolean = False)
    On Error Resume Next
    Dim LBEnabled As Boolean
    LBEnabled = CByte(GetSetting("DownloadBooster", "Options", "ProgressFrameSkin", 1)) > 0
    Dim FontName$, FontSize%, FontBold As Boolean
    FontName = Trim$(GetSetting("DownloadBooster", "Options", "Font", ""))
    If FontName = "" And LangID = 1042 Then
        If Force Or (DefaultFont <> "굴림") Then
            FontName = DefaultFont
        Else
            GoTo setlbfont
        End If
    End If
    If FontName = "" Then FontName = "Tahoma"
    If Not FontExists(FontName) Then
        If LangID = 1042 Then GoTo setlbfont
        FontName = "Tahoma"
    End If
    FontSize = 9
    If FontName = "Tahoma" Or Left$(FontName, 7) = "Tahoma " Then FontSize = 8
    frm.Font.Name = FontName
    frm.Font.Size = FontSize
    Dim ctrl As Control
    For Each ctrl In frm.Controls
        ctrl.Font.Name = FontName
        If ctrl.Tag <> "nocolorsizechange" And ctrl.Tag <> "nosizechange" Then ctrl.Font.Size = FontSize
        ctrl.FontName = FontName
        If ctrl.Tag <> "nocolorsizechange" And ctrl.Tag <> "nosizechange" Then ctrl.FontSize = FontSize
        If (Not ctrl Is frmMain.lblLBCaption) And (Not ctrl Is frmMain.lblLBCaptionShadow) And (Not ctrl Is frmMain.lblLBCaption2) And (Not ctrl Is frmMain.lblLBCaptionShadow2) Then
            ctrl.FontBold = False
            ctrl.Font.Bold = False
        End If
        ctrl.FontItalic = False
        ctrl.Font.Italic = False
    Next ctrl
setlbfont:
    If frm Is frmMain Then
        If LCase(frm.lblURL.Font.Name) = "tahoma" Then FontSize = 8 Else FontSize = 9
        If LBEnabled Then
            If GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinEnableLabelFontSize", "0") <> "0" Then
                FontSize = CInt(GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinLabelFontSize", 10))
            End If
            FontBold = GetSetting("DownloadBooster", "Options", "LiveBadukMemoSkinLabelFontBold", "0") <> "0"
        Else
            FontBold = False
        End If
        frm.lblURL.Font.Size = FontSize
        frm.lblFilePath.Font.Size = FontSize
        frm.lblThreadCountLabel.Font.Size = FontSize
        frm.lblURLShadow.Font.Size = FontSize
        frm.lblFilePathShadow.Font.Size = FontSize
        frm.lblThreadCountLabelShadow.Font.Size = FontSize
        frm.lblLBCaption.Font.Size = FontSize
        frm.lblLBCaption2.Font.Size = FontSize
        frm.lblLBCaptionShadow.Font.Size = FontSize
        frm.lblLBCaptionShadow2.Font.Size = FontSize
        frm.lblURL.Font.Bold = FontBold
        frm.lblFilePath.Font.Bold = FontBold
        frm.lblThreadCountLabel.Font.Bold = FontBold
        frm.lblURLShadow.Font.Bold = FontBold
        frm.lblFilePathShadow.Font.Bold = FontBold
        frm.lblThreadCountLabelShadow.Font.Bold = FontBold
        frm.lblLBCaption.Font.Bold = FontBold
        frm.lblLBCaption2.Font.Bold = FontBold
        frm.lblLBCaptionShadow.Font.Bold = FontBold
        frm.lblLBCaptionShadow2.Font.Bold = FontBold
    End If
End Sub

Function FormatTime(Sec) As String
    Dim Hour As Integer, Minutes As Integer, Seconds As Integer
    Dim ret As String
    If Sec >= 3600 Then
        ret = CStr(Floor(Sec / 3600)) & t("시간 ", " hours, ")
    Else
        ret = ""
    End If

    If Sec >= 60 Then
        ret = ret & Floor((Sec Mod 3600) / 60) & t("분 ", " minutes and ")
    End If
    ret = ret & (Sec Mod 60) & t("초", " seconds")
    FormatTime = ret
End Function

Function btoa(Str As String) As String
    On Error Resume Next
    Dim Data() As Byte
    Data = StrConv(Str, vbFromUnicode)
    Dim ss As String, s As Long
    ss = String$(2 * UBound(Data) + 6, 0)
    s = Len(ss) + 1
    CryptBinaryToString VarPtr(Data(0)), UBound(Data) + 1, CRYPT_STRING_BASE64, StrPtr(ss), s
    btoa = Left$(ss, s)
End Function

Sub BuildHeaderCache()
    Dim Headers() As String
    Dim RawHeaders As String
    RawHeaders = ""
    On Error GoTo dontbuild
    Headers = GetAllSettings("DownloadBooster", "Options\Headers")
    On Error GoTo 0
    Dim i%
    For i = LBound(Headers) To UBound(Headers)
        RawHeaders = RawHeaders & LCase(Headers(i, 0)) & ": " & Headers(i, 1) & vbLf
    Next i
    If Right$(RawHeaders, 1) = vbLf Then RawHeaders = Left$(RawHeaders, Len(RawHeaders) - 1)
    HeaderCache = btoa(RawHeaders)
dontbuild:
End Sub

Function DecodeHeaderCache(HeaderCache As String) As Collection
    Set DecodeHeaderCache = New Collection
    Dim Headers As Collection
    Dim HeaderKeys As Collection
    Set Headers = New Collection
    Set HeaderKeys = New Collection
    If LenB(HeaderCache) = 0 Then GoTo returncollection
    Dim RawHeaders$
    RawHeaders = StrConv(atob(HeaderCache), vbUnicode)
    Dim HeaderSplit() As String
    HeaderSplit = Split(RawHeaders, vbLf)
    Dim HeaderLine$
    Dim ColonPos%
    Dim i%
    For i = LBound(HeaderSplit) To UBound(HeaderSplit)
        HeaderLine = HeaderSplit(i)
        ColonPos = InStr(HeaderLine, ": ")
        If ColonPos < 1 Then GoTo continue
        Headers.Add Mid$(HeaderLine, ColonPos + 2), Left$(HeaderLine, ColonPos - 1)
        HeaderKeys.Add Left$(HeaderLine, ColonPos - 1)
continue:
    Next i
returncollection:
    DecodeHeaderCache.Add HeaderKeys, "keys"
    DecodeHeaderCache.Add Headers, "Values"
End Function

Function GetSpecialFolder(CSIDL As Long) As String
    Dim lngRetVal As Long
    Dim IDL As ITEMIDLIST
    Dim strPath As String
    lngRetVal = SHGetSpecialFolderLocation(100&, CSIDL, IDL)
    If lngRetVal = 0& Then
        strPath = Space$(512)
        lngRetVal = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal strPath)
        GetSpecialFolder = Left$(strPath, InStr(strPath, Chr$(0)) - 1)
    End If
End Function

Sub GetDiskSpace(sDrive As String, ByRef dblTotal As Double, ByRef dblFree As Double)
    Dim lResult As Long
    Dim liAvailable As LARGE_INTEGER
    Dim liTotal As LARGE_INTEGER
    Dim liFree As LARGE_INTEGER
    lResult = GetDiskFreeSpaceEx(sDrive, liAvailable, liTotal, liFree)
    dblTotal = CLargeInt(liTotal.LowPart, liTotal.HighPart)
    dblFree = CLargeInt(liFree.LowPart, liFree.HighPart)
End Sub

Private Function CLargeInt(Lo As Long, Hi As Long) As Double
    Dim dblLo As Double, dblHi As Double

    If Lo < 0 Then
        dblLo = 2 ^ 32 + Lo
    Else
        dblLo = Lo
    End If

    If Hi < 0 Then
        dblHi = 2 ^ 32 + Hi
    Else
        dblHi = Hi
    End If
    
    CLargeInt = dblLo + dblHi * 2 ^ 32
End Function

Sub ShellExecute(sFile As String, Optional Action As String = "open", Optional WorkingDirectory As String)
    Dim shinfo As SHELLEXECUTEINFO
    With shinfo
        .cbSize = LenB(shinfo)
        .lpFile = sFile
        .nShow = SW_SHOW
        If Action = "properties" Then .fMask = SEE_MASK_INVOKEIDLIST
        If LenB(WorkingDirectory) Then .lpDirectory = WorkingDirectory
        .lpVerb = Action
    End With
    ShellExecuteEx shinfo
End Sub

Function GetShortcutTarget(sPath As String) As String
    Dim shl As Shell, file As FolderItem, fld As shell32.Folder
    Dim lnk As ShellLinkObject, i As Long, folderPath As String
    Dim Shortcutname As String

    On Error GoTo exit_sub
    folderPath = GetParentFolderName(sPath)
    Set shl = New Shell
    Set fld = shl.NameSpace(folderPath)
    Set file = fld.Items.Item(GetFilename(sPath))
    If Err Then
        GetShortcutTarget = "."
        GoTo exit_sub
    Else
        If file.IsLink Then
            Set lnk = file.GetLink
            GetShortcutTarget = lnk.Path
            If Left$(GetShortcutTarget, 1) = """" And Right$(GetShortcutTarget, 1) = """" Then GetShortcutTarget = Mid$(GetShortcutTarget, 2, Len(GetShortcutTarget) - 2)
        Else
            GetShortcutTarget = "."
        End If
    End If
exit_sub:
    Set lnk = Nothing
    Set file = Nothing
    Set fld = Nothing
    Set shl = Nothing
End Function

Function atob(sText As String) As Byte()
    Dim lSize           As Long
    Dim dwDummy         As Long
    Dim baOutput()      As Byte
    lSize = Len(sText) + 1
    ReDim baOutput(0 To lSize - 1) As Byte
    CryptStringToBinary StrPtr(sText), Len(sText), CRYPT_STRING_BASE64, VarPtr(baOutput(0)), lSize, 0, dwDummy
    If lSize > 0 Then
        ReDim Preserve baOutput(0 To lSize - 1) As Byte
        atob = baOutput
    Else
        atob = vbNullString
    End If
End Function

Function FormatModified(DateTime) As String
    If t(1, 2) = 1 Then
        FormatModified = Replace(Replace(Format(DateTime, "yyyy-mm-dd AM/PM h:mm"), "AM", "오전"), "PM", "오후")
    Else
        FormatModified = Replace(Format(DateTime, "m-d-yyyy h:mm AM/PM"), "-", "/")
    End If
End Function

Function GetParentFolderName(ByVal Path As String) As String
    On Error GoTo errfso
    Do While Right$(Path, 1) = "\"
        Path = Left$(Path, Len(Path) - 1)
    Loop
    If InStrRev(Path, "\") = 0 Then GoTo errfso
    GetParentFolderName = Left$(Path, InStrRev(Path, "\") - 1)
    Do While Right$(GetParentFolderName, 1) = "\"
        GetParentFolderName = Left$(GetParentFolderName, Len(GetParentFolderName) - 1)
    Loop
    If Len(GetParentFolderName) = 2 And Right$(GetParentFolderName, 1) = ":" Then GetParentFolderName = GetParentFolderName & "\"
    Exit Function
errfso:
    GetParentFolderName = ""
End Function

Function GetFilename(ByVal Path As String) As String
    On Error GoTo errfso
    Do While Right$(Path, 1) = "\"
        Path = Left$(Path, Len(Path) - 1)
    Loop
    GetFilename = Mid$(Path, InStrRev(Path, "\") + 1)
    Exit Function
errfso:
    GetFilename = ""
End Function

Function GetExtensionName(ByVal Path As String) As String
    On Error GoTo errfso
    Path = GetFilename(Path)
    If InStrRev(Path, ".") = 0 Then GoTo errfso
    GetExtensionName = Mid$(Path, InStrRev(Path, ".") + 1)
    Exit Function
errfso:
    GetExtensionName = ""
End Function

Function ArrayIncludes(Target, toFind) As Boolean
    Dim i&
    For i = LBound(Target) To UBound(Target)
        If Target(i) = toFind Then
            ArrayIncludes = True
            Exit Function
        End If
    Next i
    ArrayIncludes = False
End Function

Function GetStrFromPtr(ByVal Ptr As Long) As String
    GetStrFromPtr = SysAllocStringByteLen(Ptr, lstrlen(Ptr))
End Function

Sub UpdateBorderWidth()
    Startup.DialogBorderWidth = GetSystemMetrics(8&)
    Startup.SizingBorderWidth = GetSystemMetrics(33&)
    Startup.PaddedBorderWidth = SizingBorderWidth - DialogBorderWidth
    Startup.CaptionHeight = GetSystemMetrics(31&)
    Startup.ScrollBarWidth = GetSystemMetrics(2&)
    WindowSkinBorderSize(-3) = SizingBorderWidth
    WindowSkinBorderSize(-2) = SizingBorderWidth
    WindowSkinBorderSize(-1) = SizingBorderWidth
    WindowSkinCaptionHeight(-1) = CaptionHeight + SizingBorderWidth
End Sub

Function ExpandEnvironmentStrings(strInput As String) As String
    Dim ret As Long
    Dim strOutput As String
    ret = ExpandEnvironmentStringsA(strInput, strOutput, ret)
    If ret = 0& Then
        ExpandEnvironmentStrings = strInput
        Exit Function
    End If
    strOutput = Space$(ret)
    ret = ExpandEnvironmentStringsA(strInput, strOutput, ret)
    ExpandEnvironmentStrings = strOutput
End Function

Function StartsWith(Str As String, s As String) As Boolean
    StartsWith = (Left$(Str, Len(s)) = s)
End Function

Function EndsWith(Str As String, s As String) As Boolean
    EndsWith = (Right$(Str, Len(s)) = s)
End Function

Function ExcludeParameters(URL As String) As String
    Dim InStrRet As Long
    InStrRet = InStr(URL, "?")
    If InStrRet Then
        ExcludeParameters = Left$(URL, InStrRet - 1)
    Else
        ExcludeParameters = URL
    End If
End Function

Function Col(Expression, ByRef IfFalse)
    If VarType(Expression) = vbString Then
        If Expression = "" Then
            Col = IfFalse
        Else
            Col = Expression
        End If
        Exit Function
    End If
    If Expression Then
        Col = Expression
    Else
        Col = IfFalse
    End If
End Function

#If HIDEYTDL Then
#Else
Function IsYtdlSupported(URL As String) As Boolean
    If EndsWith(LCase(ExcludeParameters(URL)), ".m3u8") Then
        IsYtdlSupported = True
        Exit Function
    End If

    Dim HostName$, InStrRet&
    InStrRet = InStr(URL, "://")
    If InStrRet Then
        HostName = Mid$(URL, InStrRet + 3)
    Else
        HostName = URL
    End If
    HostName = Left$(HostName, Col(InStr(HostName, "/"), Len(HostName) + 1) - 1)
    If InStr(HostName, ":") Then
        HostName = Left$(HostName, InStrRev(HostName, ":") - 1)
    End If
    HostName = LCase(HostName)

    IsYtdlSupported = ArrayIncludes(Array("youtube.com", "soundcloud.com", "ok.ru", "bilibili.tv", "dailymotion.com"), HostName)
End Function
#End If

Sub tr(ByRef ctrl As Object, EnglishCaption As String)
    'On Error Resume Next
    If LangID <> 1042 Then ctrl.Caption = EnglishCaption
End Sub

Function GetThemeColor(ByVal hWnd As Long, ClassList As String, Optional ByVal Part As Long = 0&, Optional ByVal State As Long = 0&, Optional ByVal Prop As Long = TMT_TEXTCOLOR, Optional ByVal DefaultColor As Long = 0&) As Long
    On Error GoTo returndefault
    Dim hTheme&, clr&

    If IsAppThemed() = 0& Or IsThemeActive() = 0& Then GoTo returndefault
    hTheme = OpenThemeData(hWnd, StrPtr(ClassList))
    If hTheme = 0& Then GoTo returndefault
    If X_GetThemeColor(hTheme, Part, State, Prop, clr) <> 0 Then GoTo returndefault
    CloseThemeData hTheme
    GetThemeColor = clr
    Exit Function

returndefault:
    If hTheme <> 0& Then CloseThemeData hTheme
    GetThemeColor = DefaultColor
End Function

Sub PlayWave(ByVal Path As String, Optional ByVal LoopWave As Boolean = False, Optional ByVal StopPreviousWave As Boolean = True, Optional ByVal FallbackSound As Long = -1&)
    If FileExists(Path) Then
        Dim Flags As Long
        Flags = SND_FILENAME Or SND_ASYNC
        If LoopWave Then Flags = Flags Or SND_LOOP
        If Not StopPreviousWave Then Flags = Flags Or SND_NOSTOP
        PlaySound Path, 0&, Flags
    ElseIf FallbackSound >= 0& Then
        MessageBeep FallbackSound
    End If
End Sub

Sub EnableFrameControls(ByRef fFrame As Control, ByRef Except As Control, Optional ByVal Enable As Boolean = True)
    Dim ctrl As Control
    For Each ctrl In fFrame.ContainedControls
        If Not ctrl Is Except Then ctrl.Enabled = Enable
    Next ctrl
End Sub

Function Max(ByVal L As Double, ByVal R As Double) As Double
    If L > R Then
        Max = L
    Else
        Max = R
    End If
End Function

Function Min(ByVal L As Double, ByVal R As Double) As Double
    If L < R Then
        Min = L
    Else
        Min = R
    End If
End Function

Function MsgBox(ByVal Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String) As VbMsgBoxResult
    If Title = "" Then Title = App.Title

    If Buttons > 70 Then
        GoTo nativemsgbox
    ElseIf Buttons < 16 Then
        MsgBox = ShowMessageBox(Prompt, Title, 0, MsgBoxMode:=Buttons)
    ElseIf (Buttons And vbInformation) = vbInformation Then
        MsgBox = ShowMessageBox(Prompt, Title, vbInformation, MsgBoxMode:=(Buttons And (Not vbInformation)))
    ElseIf (Buttons And vbExclamation) = vbExclamation Then
        MsgBox = ShowMessageBox(Prompt, Title, vbExclamation, MsgBoxMode:=(Buttons And (Not vbExclamation)))
    ElseIf (Buttons And vbQuestion) = vbQuestion Then
        MsgBox = ShowMessageBox(Prompt, Title, vbQuestion, MsgBoxMode:=(Buttons And (Not vbQuestion)))
    ElseIf (Buttons And vbCritical) = vbCritical Then
        MsgBox = ShowMessageBox(Prompt, Title, vbCritical, MsgBoxMode:=(Buttons And (Not vbCritical)))
    Else
        GoTo nativemsgbox
    End If

    Exit Function
nativemsgbox:
    MsgBox = VBA.MsgBox(Prompt, Buttons, Title)
End Function

Function Right(Str As String, Length As Long) As String
    On Error GoTo errproc
    Right = VBA.Right$(Str, Length)
    Exit Function
errproc:
    Right = ""
End Function

'https://www.vbforums.com/showthread.php?842795-Can-vb6-detect-the-DPI
Function GetDPI() As Long
    Dim hWndDesktop As Long
    Dim hDCDesktop As Long

    hWndDesktop = GetDesktopWindow()
    hDCDesktop = GetDC(hWndDesktop)
    GetDPI = GetDeviceCaps(hDCDesktop, 88&)
    ReleaseDC hWndDesktop, hDCDesktop
End Function

Sub UpdateDPI()
    DPI = GetDPI()
End Sub

Function Ceil(val)
    Dim Rounded
    Rounded = Round(val)
    If Rounded < val Then
        Ceil = Rounded + 1
    Else
        Ceil = Rounded
    End If
End Function

Function GetPictureWidth(pic As StdPicture, Optional ToScale As ScaleModeConstants = vbTwips) As Long
    GetPictureWidth = Round(frmDummyForm.ScaleX(pic.Width, vbHimetric, ToScale))
End Function

Function GetPictureHeight(pic As StdPicture, Optional ToScale As ScaleModeConstants = vbTwips) As Long
    GetPictureHeight = Round(frmDummyForm.ScaleY(pic.Height, vbHimetric, ToScale))
End Function

Sub ExtractResource(ByVal ResourceID As Integer, ByVal ResourceType As ResourceType, FileName As String)
    Dim ff As Integer
    Dim B() As Byte
    On Error Resume Next
    MkDir CachePath
    On Error GoTo 0

    If Not FileExists(CachePath & FileName) Then
        B = LoadResData(ResourceID, ResourceType)
        ff = FreeFile()
        Open CachePath & FileName For Binary Access Write As #ff
        Put #ff, , B
        Close #ff
        Erase B
    End If
End Sub

Function LoadResText(ByVal ResourceID As Integer, ByVal ResourceType As ResourceType) As String
    LoadResText = StrConv(LoadResData(ResourceID, ResourceType), vbUnicode)
End Function

'https://stackoverflow.com/questions/1230333
Sub AddItemToComboBox(cbComboBox As ComboBox, ByVal Text As String)
    SendMessage cbComboBox.hWnd, CB_ADDSTRING, 0&, ByVal Text
End Sub

Sub ClearComboBox(cbComboBox As ComboBox)
    SendMessage cbComboBox.hWnd, CB_RESETCONTENT, 0&, 0&
End Sub

'https://www.vbforums.com/showthread.php?704979
Sub SetFileDate(ByVal sFilename As String, ByVal dFileDate As Date)
    Dim lhwndFile As Long
    Dim tSystemTime As SYSTEMTIME
    Dim tLocalTime As FILETIME, lpLastWriteTime As FILETIME, lpCreationTime As FILETIME

    Dim uInfo As TIME_ZONE_INFORMATION
    Dim wMonth%, wDay%, wHour%
    If GetTimeZoneInformation(uInfo) > 0& Then
        wMonth = Month(dFileDate)
        wDay = Day(dFileDate)
        wHour = Hour(dFileDate)
        If (wMonth > uInfo.DaylightDate(5) Or (wMonth = uInfo.DaylightDate(5) And (wDay > uInfo.DaylightDate(0) Or (wDay = uInfo.DaylightDate(5) And wHour > uInfo.DaylightDate(2))))) And (Not (wMonth > uInfo.StandardDate(5) Or (wMonth = uInfo.StandardDate(5) And (wDay > uInfo.StandardDate(0) Or (wDay = uInfo.StandardDate(5) And wHour > uInfo.StandardDate(2)))))) Then
            dFileDate = DateAdd("n", -uInfo.DaylightBias, dFileDate)
        End If
    End If

    tSystemTime.Year = Year(dFileDate)
    tSystemTime.Month = Month(dFileDate)
    tSystemTime.Day = Day(dFileDate)
    tSystemTime.Hour = Hour(dFileDate)
    tSystemTime.Minute = Minute(dFileDate)
    tSystemTime.Second = Second(dFileDate)
    tSystemTime.Milliseconds = 50
    lhwndFile = CreateFile(sFilename, 256&, 1&, ByVal 0&, 3&, 0&, 0&)
    If lhwndFile Then
        SystemTimeToFileTime tSystemTime, tLocalTime
        LocalFileTimeToFileTime tLocalTime, lpLastWriteTime
        GetFileTime lhwndFile, lpCreationTime, 0&, 0&
        SetFileTime lhwndFile, lpCreationTime, 0&, lpLastWriteTime
        CloseHandle lhwndFile
    End If
End Sub

Sub NextTabPage(ByRef tsTabStrip As TabStrip, Optional ByVal Reverse As Boolean = False)
    On Error Resume Next
    Dim A%, B%, X%, Y%, Z%
    A = tsTabStrip.Tabs.Count
    B = tsTabStrip.SelectedItem.Index
    If Reverse Then X = 1 Else X = A
    If Reverse Then Y = A Else Y = 1
    If Reverse Then Z = -1 Else Z = 1
    If B = X Then
        tsTabStrip.Tabs(Y).Selected = True
    Else
        tsTabStrip.Tabs(B + Z).Selected = True
    End If
End Sub

Sub InitForm(ByRef frmForm As Form)
    On Error Resume Next
    If GetSetting("DownloadBooster", "Options", "DisableDWMWindow", DefaultDisableDWMWindow) <> 0 Then DisableDWMWindow frmForm.hWnd
    SetFormBackgroundColor frmForm
    SetFont frmForm
    If MainFormOnTop Then SetWindowPos frmForm.hWnd, hWnd_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE
    'If frmForm.BorderStyle = 2 Then Set frmForm.Icon = frmMain.Icon
#If DISABLEFRAMESKIN Then
#Else
    On Error GoTo noskin
    Set frmForm.SkinnedFrame = New frmSkinnedFrame
    frmForm.SkinnedFrame.Init frmForm
noskin:
    'On Error Resume Next
#End If
End Sub

Function GenerateSolidColor(ByVal Color As Long) As IPictureDisp
    frmDummyForm.pbDummy.Cls
    frmDummyForm.pbDummy.DrawWidth = 5
    frmDummyForm.pbDummy.Width = 15
    frmDummyForm.pbDummy.Height = 15
    frmDummyForm.pbDummy.Line (0, 0)-(15, 15), Color
    frmDummyForm.pbDummy.Refresh
    Set GenerateSolidColor = frmDummyForm.pbDummy.Image
    frmDummyForm.pbDummy.Cls
End Function

Sub OpenFolder(ByVal Path As String)
    If Not FolderExists(Path) Then Path = GetParentFolderName(Path)
    ShellExecute Path
End Sub

'Function RemoveQuotes(Path As String) As String
'    RemoveQuotes = Path
'    If Left$(RemoveQuotes, 1) = """" And Right$(RemoveQuotes, 1) = """" Then RemoveQuotes = Mid$(RemoveQuotes, 2, Len(RemoveQuotes) - 2)
'End Function

Sub ShowFileDialog(Optional TargetForm As Byte, Optional PresetPath As String, Optional ForceNewDialog As Boolean)
    Tags.BrowsePresetPath = PresetPath
    Tags.BrowseTargetForm = TargetForm
    If (Not ForceNewDialog) And GetSetting("DownloadBooster", "Options", "ForceWin31Dialog", "0") <> "0" Then
        frmBrowse.Show vbModal
    Else
        Dim Explorer As frmExplorer
        Set Explorer = New frmExplorer
        Explorer.Show vbModal
        Set Explorer = Nothing
    End If
End Sub

Function CreatePicture(Handle As Long, PicType As Long) As IPicture
    Dim uDesc As PICTDESC
    uDesc.cbSize = Len(uDesc)
    uDesc.PicType = PicType
    uDesc.hgdiObj = Handle
    OleCreatePictureIndirect uDesc, IPictureIID, True, CreatePicture
End Function

'Sub ShellGetFileInfo(Path As String, UseFileAttributes As Boolean, ByRef LargeIcon As IPicture, ByRef SmallIcon As IPicture, ByRef TypeName As String)
'    Dim SFI As SHFILEINFO
'    Dim PathPtr&, SfiPtr&, SfiSize&
'    Dim FlagUFA As Long
'    PathPtr = StrPtr(Path)
'    SfiPtr = VarPtr(SFI)
'    SfiSize = LenB(SFI)
'    FlagUFA = (-UseFileAttributes) * SHGFI_USEFILEATTRIBUTES
'
'    If SHGetFileInfo(PathPtr, 0&, SfiPtr, SfiSize, FlagUFA Or SHGFI_ICON Or SHGFI_LARGEICON Or SHGFI_TYPENAME) = 0 Then GoTo onfail
'    TypeName = SFI.szTypeName
'    Set LargeIcon = CreatePicture(SFI.hIcon, vbPicTypeIcon)
'    SHGetFileInfo PathPtr, 0&, SfiPtr, SfiSize, FlagUFA Or SHGFI_ICON Or SHGFI_SMALLICON
'    Set SmallIcon = CreatePicture(SFI.hIcon, vbPicTypeIcon)
'onfail:
'End Sub

Function GetExtensionDescription(Path As String) As String
    Dim SFI As SHFILEINFO
    If SHGetFileInfo(Path, 0&, SFI, LenB(SFI), SHGFI_USEFILEATTRIBUTES Or SHGFI_TYPENAME) = 0& Then GoTo onfail
    GetExtensionDescription = SFI.szTypeName
onfail:
End Function

Sub InitPropertySheetDimensions(frmForm As Form, tsTabStrip As TabStrip, Panels As Object, OKButton As CommandButtonW, CancelButton As CommandButtonW, Optional ApplyButton As CommandButtonW)
    Dim i As Byte
    Dim MaxWidth%, MaxHeight%
    Dim ClientLeft%, ClientTop%, Left%, Top%, Width%, Height%, ButtonTop%
    ClientLeft = tsTabStrip.ClientLeft
    ClientTop = tsTabStrip.ClientTop
    Left = tsTabStrip.Left
    Top = tsTabStrip.Top
    Width = tsTabStrip.Width
    Height = tsTabStrip.Height
    For i = Panels.LBound To Panels.UBound
        Panels(i).Top = ClientTop + Top
        Panels(i).Left = ClientLeft + Left
        If MaxWidth < Panels(i).Width Then MaxWidth = Panels(i).Width
        If MaxHeight < Panels(i).Height Then MaxHeight = Panels(i).Height
    Next i
    For i = Panels.LBound To Panels.UBound
        Panels(i).Width = MaxWidth
        Panels(i).Height = MaxHeight
    Next i
    Width = MaxWidth + (Width - tsTabStrip.ClientWidth)
    tsTabStrip.Width = Width
    Height = MaxHeight + (Height - tsTabStrip.ClientHeight)
    tsTabStrip.Height = Height
    ButtonTop = Top + Height + 60
    CancelButton.Top = ButtonTop
    OKButton.Top = ButtonTop
    Dim ButtonLeft%
    ButtonLeft = 120 + Width - PROPERTY_SHEET_BUTTON_WIDTH
    If ApplyButton Is Nothing Then
        CancelButton.Left = ButtonLeft
    Else
        ApplyButton.Top = ButtonTop
        ApplyButton.Left = ButtonLeft
        CancelButton.Left = ButtonLeft - 120 - PROPERTY_SHEET_BUTTON_WIDTH
    End If
    OKButton.Left = CancelButton.Left - 120 - PROPERTY_SHEET_BUTTON_WIDTH
    frmForm.Height = ButtonTop + PROPERTY_SHEET_BUTTON_HEIGHT + 540
    frmForm.Width = Width + 300
End Sub

Function RunNode(SP As ShellPipe, FileName As String, Script() As Byte, Optional ByVal Arguments As String, Optional Environment As String) As SP_RESULTS
    Arguments = "-e ""var c='';process.stdin.setEncoding('utf8');process.stdin.on('data',function(d){c+=d;});process.stdin.on('end',function(){module._compile(c,'" & FileName & "');});"" 0 " & Arguments
    If LaunchFromMemory Then
        RunNode = SP.RunInMemory(NodeJS, Arguments)
    Else
        Dim NodePath$: NodePath = GetSetting("DownloadBooster", "Options", "NodePath", "")
        If NodePath = "" Then NodePath = CachePath & NodeFileName
        RunNode = SP.Run("""" & NodePath & """ " & Arguments)
    End If
    If RunNode = SP_SUCCESS Then
        If SP.SendBytes(Script) = False Then
            RunNode = SP_CREATEPIPEFAILED
            SP.FinishChild 0, 0
        Else
            SP.ClosePipe
        End If
    End If
End Function

Sub StripTrailingNull(ByRef B() As Byte)
    Dim i As Long
    Dim newSize As Long
    newSize = -1
    For i = UBound(B) To 0 Step -1
        If B(i) <> 0 Then
            newSize = i
            Exit For
        End If
    Next i
    If newSize = -1 Then
        ReDim B(-1)
    Else
        ReDim Preserve B(newSize)
    End If
End Sub

'Function ConvertUTF8(ByRef Bytes() As Byte) As String
'    Dim i As Long
'    Dim CodePoint As Long
'    Dim c As Long
'    Dim ret As String
'    i = 0
'    Do While i <= UBound(Bytes)
'        c = Bytes(i)
'        If c < &H80 Then
'            ret = ret & ChrW(c)
'            i = i + 1
'        ElseIf (c And &HE0) = &HC0 Then
'            If i + 1 > UBound(Bytes) Then Exit Do
'            CodePoint = ((c And &H1F) * &H40) Or (Bytes(i + 1) And &H3F)
'            ret = ret & ChrW(CodePoint)
'            i = i + 2
'        ElseIf (c And &HF0) = &HE0 Then
'            If i + 2 > UBound(Bytes) Then Exit Do
'            CodePoint = ((c And &HF) * &H1000) Or ((Bytes(i + 1) And &H3F) * &H40) Or (Bytes(i + 2) And &H3F)
'            ret = ret & ChrW(CodePoint)
'            i = i + 3
'        ElseIf (c And &HF8) = &HF0 Then
'            If i + 3 > UBound(Bytes) Then Exit Do
'            CodePoint = ((c And &H7) * &H40000) Or ((Bytes(i + 1) And &H3F) * &H1000) Or ((Bytes(i + 2) And &H3F) * &H40) Or (Bytes(i + 3) And &H3F)
'            CodePoint = CodePoint - &H10000
'            ret = ret & ChrW(&HD800 + (CodePoint \ &H400)) & ChrW(&HDC00 + (CodePoint And &H3FF))
'            i = i + 4
'        Else
'            i = i + 1
'        End If
'    Loop
'    ConvertUTF8 = ret
'End Function

'Function MinifyScript(ByRef Script As String) As String
'    MinifyScript = Replace(Replace(Replace(Script, vbLf, " "), vbCr, ""), vbTab, "")
'End Function

'Function IsWOW64() As Boolean
'    On Error GoTo Not64
'
'    Dim bIsWow64 As Long
'    Dim ret As Long
'    ret = IsWow64Process(GetCurrentProcess(), bIsWow64)
'    If ret = 0 Then GoTo Not64
'    IsWOW64 = bIsWow64 <> 0
'    Exit Function
'
'Not64:
'    IsWOW64 = False
'End Function

'Function GetCurrentEXEPath(Optional Fallback As String) As String
'    Dim bufferSize As Long
'    Dim exePath As String
'    Dim ret As Long
'
'    bufferSize = 260
'    Do
'        exePath = String$(bufferSize, vbNullChar)
'        ret = GetModuleFileName(0, exePath, bufferSize)
'        If ret = 0 Then
'            GetCurrentEXEPath = Fallback
'            Exit Function
'        End If
'        If ret < bufferSize - 1 Then Exit Do
'        bufferSize = bufferSize * 2
'    Loop
'
'    GetCurrentEXEPath = Left$(exePath, ret)
'End Function

Function CreateRectRgnFix(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Dim DPIScale As Double: DPIScale = DPI / 96
    CreateRectRgnFix = CreateRectRgn(X1 * DPIScale, Y1 * DPIScale, X2 * DPIScale, Y2 * DPIScale)
End Function

Sub DrawTabBackground(frmForm As Form, tsTabStrip As TabStrip, pbPanel As Object, Optional Refresh As Boolean = True)
    On Error Resume Next
    Dim ctrl As Control
    Dim i As Byte
    
    For i = pbPanel.LBound To pbPanel.UBound
        tsTabStrip.DrawBackground pbPanel(i).hWnd, pbPanel(i).hDC
    Next i
    
    If Not Refresh Then Exit Sub
    For Each ctrl In frmForm.Controls
        If TypeOf ctrl Is FrameW Or TypeOf ctrl Is CheckBoxW Or TypeOf ctrl Is CommandButtonW Then
            ctrl.Refresh
        ElseIf TypeOf ctrl Is Slider Then
            ctrl.Refresh
            ctrl.VisualStyles = Not ctrl.VisualStyles
            ctrl.VisualStyles = Not ctrl.VisualStyles
        End If
    Next ctrl
End Sub

Sub SetLabelText(lbl As Label, ByVal fullText As String)
    Dim i As Long
    Dim tempText As String
    lbl = fullText
    If TextWidth(fullText, lbl.Font.Name, lbl.Font.Size, lbl.Font.Bold) > lbl.Width Then
        For i = Len(fullText) To 1 Step -1
            tempText = Left$(fullText, i) & "..."
            If TextWidth(tempText, lbl.Font.Name, lbl.Font.Size, lbl.Font.Bold) <= lbl.Width Then
                lbl = tempText
                Exit Sub
            End If
        Next i
        lbl = "..."
    End If
End Sub

Function CompareRect(R1 As RECT, R2 As RECT) As Boolean
    CompareRect = (R1.Bottom = R2.Bottom And R1.Left = R2.Left And R1.Right = R2.Right And R1.Top = R2.Top)
End Function

Function GetDCFromPicture(pic As StdPicture) As Long
    Dim hDCMem As Long
    hDCMem = CreateCompatibleDC(0&)
    SelectObject hDCMem, pic.Handle
    GetDCFromPicture = hDCMem
End Function
