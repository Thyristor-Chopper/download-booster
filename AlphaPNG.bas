Attribute VB_Name = "AlphaPNG"
'https://www.vbforums.com/showthread.php?896878-PNG-with-alpha-channel-into-standard-VB6-image-control

Option Explicit

Private Type GdiplusStartupInput
    GdiplusVersion              As Long
    DebugEventCallback          As Long
    SuppressBackgroundThread    As Long
    SuppressExternalCodecs      As Long
End Type
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef lpInput As GdiplusStartupInput, Optional ByRef lpOutput As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal sFilename As Long, hImage As Long) As Long
Private Type BitmapData
    Width               As Long
    Height              As Long
    Stride              As Long
    PixelFormat         As Long
    Scan0               As Long
    Reserved            As Long
End Type
Private Const ImageLockModeRead As Long = &H1&
Private Const PixelFormat32bppPARGB As Long = &HE200B
Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal hBitmap As Long, lpRect As Any, ByVal lFlags As Long, ByVal lPixelFormat As Long, uLockedBitmapData As BitmapData) As Long
Private Type BITMAPINFOHEADER
    biSize              As Long
    biWidth             As Long
    biHeight            As Long
    biPlanes            As Integer
    biBitCount          As Integer
    biCompression       As Long
    biSizeImage         As Long
    biXPelsPerMeter     As Long
    biYPelsPerMeter     As Long
    biClrUsed           As Long
    biClrImportant      As Long
End Type
Private Const DIB_RGB_COLORS As Long = 0&
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, lpBitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long, lpBitsOut As Long, ByVal hSection As Long, ByVal offset As Long) As Long
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal hBitmap As Long, uLockedBitmapData As BitmapData) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function CreateEnhMetaFileW Lib "gdi32" (ByVal hdcRef As Long, ByVal lpFileName As Long, lpRect As Any, ByVal lpDescription As Long) As Long
Private Type BLENDFUNCTION
    BlendOp             As Byte
    BlendFlags          As Byte
    SourceConstantAlpha As Byte
    AlphaFormat         As Byte
End Type
Private Declare Function GetMem4 Lib "msvbvm60" (ByRef Source As Any, ByRef dest As Any) As Long ' Always ignore the returned Value, it's useless.
Private Declare Function GdiAlphaBlend Lib "gdi32" (ByVal hdcDest As Long, ByVal xoriginDest As Long, ByVal yoriginDest As Long, ByVal wDest As Long, ByVal hDest As Long, ByVal hdcSrc As Long, ByVal xoriginSrc As Long, ByVal yoriginSrc As Long, ByVal wSrc As Long, ByVal hSrc As Long, ByVal ftn As Long) As Long
Private Declare Function CloseEnhMetaFile Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function IIDFromString Lib "ole32" (ByVal lpsz As Long, ByRef CLSID As IID) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long

Private Declare Function CreateStreamOnHGlobal Lib "ole32.dll" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As Long, hImage As Long) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (ByVal pStream As Long, ByVal lSize As Long, ByVal fRunMode As Long, riid As IID, ipic As StdPicture) As Long

'Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Const HORZSIZE   As Long = 4&
Const VERTSIZE   As Long = 6&
Const HORZRES    As Long = 8&
Const VERTRES    As Long = 10&
Const LOGPIXELSX As Long = 88&
Const LOGPIXELSY As Long = 90&
Const AC_SRC_OVER  As Byte = 0
Const AC_SRC_ALPHA As Byte = 1
Const GMEM_MOVEABLE = &H2

Function LoadPictureFromResource(ByVal ResourceID As Integer, ByVal ResourceType As ResourceType) As IPicture
    Set LoadPictureFromResource = LoadPictureFromBuffer(LoadResData(ResourceID, ResourceType))
End Function

Function LoadPictureFromBuffer(Buffer() As Byte) As IPicture
    'https://www.vbforums.com/showthread.php?805563-RESOLVED-GDI-Load-image-from-a-byte-array-into-PictureBox
    Dim o_hMem&, Length&, Stream As stdole.IUnknown
    Length = UBound(Buffer)
    o_hMem = GlobalAlloc(GMEM_MOVEABLE, Length)
    CopyMemory ByVal GlobalLock(o_hMem), ByVal VarPtr(Buffer(0)), Length
    GlobalUnlock o_hMem
    CreateStreamOnHGlobal o_hMem, 1&, Stream
    Set LoadPictureFromBuffer = LoadPngIntoPictureWithAlpha(StreamPtr:=ObjPtr(Stream))
    Set Stream = Nothing
    GlobalFree o_hMem
End Function

Function LoadPictureFromFile(Path As String) As IPicture
    On Error Resume Next
    Select Case LCase(GetExtensionName(Path))
        Case "cur", "wmf", "emf"
            Set LoadPictureFromFile = LoadPicture(Path)
        Case Else
            Set LoadPictureFromFile = LoadPngIntoPictureWithAlpha(StrPtr(Path))
    End Select
End Function

Private Function LoadPngIntoPictureWithAlpha(Optional PathPtr As Long, Optional StreamPtr As Long) As IPicture
    On Error GoTo loaderror
    Dim mlGdipToken As Long
    Dim StartupInput As GdiplusStartupInput
    StartupInput.GdiplusVersion = 1&
    GdiplusStartup mlGdipToken, StartupInput, 0&
    Dim hGdipImage As Long
    If PathPtr Then
        If GdipLoadImageFromFile(PathPtr, hGdipImage) Then GoTo loaderror
    Else
        If GdipLoadImageFromStream(StreamPtr, hGdipImage) Then GoTo loaderror
    End If
    Dim uData As BitmapData
    GdipBitmapLockBits hGdipImage, ByVal 0&, ImageLockModeRead, PixelFormat32bppPARGB, uData
    Dim hMemDC As Long
    hMemDC = CreateCompatibleDC(0&)
    Dim uHdr As BITMAPINFOHEADER
    uHdr.biSize = Len(uHdr)
    uHdr.biPlanes = 1
    uHdr.biBitCount = 32
    uHdr.biWidth = uData.Width
    uHdr.biHeight = -uData.Height
    uHdr.biSizeImage = uData.Stride * uData.Height
    Dim hDib As Long
    Dim lpBits As Long
    hDib = CreateDIBSection(hMemDC, uHdr, DIB_RGB_COLORS, lpBits, 0&, 0&)
    CopyMemory ByVal lpBits, ByVal uData.Scan0, uData.Stride * uData.Height
    GdipBitmapUnlockBits hGdipImage, uData
    GdipDisposeImage hGdipImage
    Dim hPrevDib As Long
    hPrevDib = SelectObject(hMemDC, hDib)
    Dim hEmfDC As Long
    hEmfDC = CreateEnhMetaFileW(0&, 0&, ByVal 0&, 0&)
    Dim Xscale As Double, Yscale As Double
    Xscale = CDbl(GetDeviceCaps(hEmfDC, HORZRES)) / CDbl(GetDeviceCaps(hEmfDC, HORZSIZE)) * 25.4 / CDbl(GetDeviceCaps(hEmfDC, LOGPIXELSX))
    Yscale = CDbl(GetDeviceCaps(hEmfDC, VERTRES)) / CDbl(GetDeviceCaps(hEmfDC, VERTSIZE)) * 25.4 / CDbl(GetDeviceCaps(hEmfDC, LOGPIXELSY))
    Dim bf As BLENDFUNCTION
    bf.BlendOp = AC_SRC_OVER
    bf.AlphaFormat = AC_SRC_ALPHA
    bf.SourceConstantAlpha = &HFF
    Dim ftn As Long
    GetMem4 bf, ftn
    GdiAlphaBlend hEmfDC, 0&, 0&, CLng(CDbl(uData.Width) * Xscale) + 1&, CLng(CDbl(uData.Height) * Yscale) + 1&, hMemDC, 0&, 0&, uData.Width, uData.Height, ftn
    SelectObject hMemDC, hPrevDib
    DeleteDC hMemDC
    DeleteObject hDib
    Set LoadPngIntoPictureWithAlpha = CreatePicture(CloseEnhMetaFile(hEmfDC), vbPicTypeEMetafile)
    GdiplusShutdown mlGdipToken
loaderror:
End Function

'Private Sub IUnknown_Release(ByVal pUnk As Long)
'    Dim vtbl As Long, pRelease As Long
'    If pUnk = 0& Then Exit Sub
'
'    CopyMemory vtbl, ByVal pUnk, 4&
'    CopyMemory pRelease, ByVal (vtbl + 8&), 4&
'    CallWindowProc pRelease, pUnk, 0&, 0&, 0&
'End Sub

Function LoadImageFromResource(ByVal ResID As Long, ByVal ResType As ResourceType) As StdPicture
    Dim B() As Byte
    Dim hMem As Long, pStream As Long
    Dim pic As StdPicture
    B = LoadResData(ResID, ResType)
    hMem = GlobalAlloc(GMEM_MOVEABLE, UBound(B) + 1)
    CopyMemory ByVal GlobalLock(hMem), B(0), UBound(B) + 1
    GlobalUnlock hMem
    CreateStreamOnHGlobal hMem, 1&, pStream
    OleLoadPicture pStream, UBound(B) + 1, True, IPictureIID, pic
    Set LoadImageFromResource = pic
    'IUnknown_Release pStream
    GlobalFree hMem
End Function
