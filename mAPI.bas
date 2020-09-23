Attribute VB_Name = "mAPI"

Option Explicit

Public Type POINTAPI
    X                   As Long
    Y                   As Long
End Type

Public Type RECT
    Left                As Long
    Top                 As Long
    Right               As Long
    Bottom              As Long
End Type

Public Type BITMAP '14 bytes
    bmType              As Long
    bmWidth             As Long
    bmHeight            As Long
    bmWidthBytes        As Long
    bmPlanes            As Integer
    bmBitsPixel         As Integer
    bmBits              As Long
End Type

Public Type BITMAPINFOHEADER
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

Public Type ICONINFO
    fIcon               As Long
    xHotspot            As Long
    yHotspot            As Long
    hbmMask             As Long
    hbmColor            As Long
End Type

Public Type OSVERSIONINFO
   dwOSVersionInfoSize  As Long
   dwMajorVersion       As Long
   dwMinorVersion       As Long
   dwBuildNumber        As Long
   dwPlatformID         As Long
   szCSDVersion         As String * 128 ' Maintenance string
End Type

Public Type Size
    cx                  As Long
    cy                  As Long
End Type

Public Const STRETCH_HALFTONE       As Long = &H4&

Public Const SND_ASYNC              As Long = &H1
Public Const SND_RESOURCE           As Long = &H40004

Public Const WM_GETSYSMENU          As Long = &H313

'window constansts
Public Const SWP_NOSIZE             As Long = &H1
Public Const SWP_NOMOVE             As Long = &H2
Public Const SWP_NOACTIVATE         As Long = &H10
Public Const SWP_SHOWWINDOW         As Long = &H40
Public Const TOPMOST_FLAGS          As Long = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST           As Long = -1
Public Const HWND_NOTOPMOST         As Long = -2
Public Const HWND_TOP               As Long = 0

Public Const WM_SETFOCUS            As Long = &H7
Public Const WM_ACTIVATE            As Long = &H6

Public Const SW_SHOWNORMAL          As Long = 1
Public Const SW_SHOWMINIMIZED       As Long = 2
Public Const SW_SHOWNOACTIVATE      As Long = 4
Public Const SW_SHOW                As Long = 5
Public Const SW_SHOWMINNOACTIVE     As Long = 7
Public Const SW_RESTORE             As Long = 9

Public Const API_TRUE               As Long = 1&

Public Const DT_CALCRECT            As Long = &H400
Public Const DT_CENTER              As Long = &H1
Public Const DT_EDITCONTROL         As Long = &H2000
Public Const DT_END_ELLIPSIS        As Long = &H8000
Public Const DT_LEFT                As Long = &H0
Public Const DT_MODIFYSTRING        As Long = &H10000
Public Const DT_NOCLIP              As Long = &H100
Public Const DT_NOPREFIX            As Long = &H800
Public Const DT_PATH_ELLIPSIS       As Long = &H4000
Public Const DT_RIGHT               As Long = &H2
Public Const DT_RTLREADING          As Long = &H20000
Public Const DT_WORDBREAK           As Long = &H10
Public Const DT_SINGLELINE          As Long = &H20

'CHM Help constants
Public Const HH_DISPLAY_TOC         As Long = &H1
Public Const HH_DISPLAY_INDEX       As Long = &H2
Public Const HH_DISPLAY_SEARCH      As Long = &H3
Public Const HH_DISPLAY_TOPIC       As Long = &H0
Public Const HH_SET_WIN_TYPE        As Long = &H4
Public Const HH_GET_WIN_TYPE        As Long = &H5
Public Const HH_GET_WIN_HANDLE      As Long = &H6
Public Const HH_DISPLAY_TEXT_POPUP  As Long = &HE
Public Const HH_HELP_CONTEXT        As Long = &HF
Public Const HH_TP_HELP_CONTEXTMENU As Long = &H10
Public Const HH_TP_HELP_WM_HELP     As Long = &H11

'Window Constants
Public AppInactive As Boolean
Public Const WM_ACTIVATEAPP   As Long = &H1C
Public Const WA_INACTIVE      As Long = 0
Public Const WA_ACTIVE        As Long = 1
Public Const GWL_WNDPROC        As Long = (-4)
Public Const WM_SIZE           As Long = &H5
Public Const WM_SIZING         As Long = &H214
Public Const WM_NCDESTROY      As Long = &H82
Public Const WM_EXITSIZEMOVE   As Long = &H232&
Public Const WM_SYSCOMMAND     As Long = &H112

'API Declares
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Public Declare Function HTMLHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function PlaySoundLong Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As Long, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetVolumeSerialNumber Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent&, ByVal hWndChildAfter&, ByVal lpClassName$, ByVal lpWindowName$) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd&, lpRect As RECT, ByVal bErase&) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function DrawTextUnicode Lib "user32" Alias "DrawTextW" (ByVal hdc As Long, ByVal lpArrPtr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function DrawFocusRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT) As Long
Public Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Public Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateDIBPatternBrushPt Lib "gdi32" (lpPackedDIB As Any, ByVal iUsage As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long
Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Public Declare Function StrFormatByteSize Lib "shlwapi" Alias "StrFormatByteSizeA" (ByVal dw As Long, ByVal pszBuf As String, ByVal cchBuf As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Public Function HiWord(lParam As Long) As Long
    'This is the HIWORD of the lParam:
    HiWord = lParam \ &H10000 And &HFFFF&
End Function

Public Function LoWord(lParam As Long) As Long
    'This is the LOWORD of the lParam:
    LoWord = lParam And &HFFFF&
End Function
