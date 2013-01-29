Attribute VB_Name = "Module1"
' This file contains constants used to assist in programming---------------------------------------------
' with the Olectra Chart controls.


' HugeValue is returned in some API calls when the control-----------------------------------------------
' can't determine an appropriate value.
'
Public Const ocHugeValue As Double = 1E+308


' Windows messages, taken out of the WINAPI.TXT file
'
' Constants for dealing with mouse events
'Public Const WM_MOUSEFIRST = &H200
'Public Const WM_MOUSEMOVE = &H200
'Public Const WM_LBUTTONDOWN = &H201
'Public Const WM_LBUTTONUP = &H202
'Public Const WM_LBUTTONDBLCLK = &H203
'Public Const WM_RBUTTONDOWN = &H204
'Public Const WM_RBUTTONUP = &H205
'Public Const WM_RBUTTONDBLCLK = &H206
'Public Const WM_MBUTTONDOWN = &H207
'Public Const WM_MBUTTONUP = &H208
'Public Const WM_MBUTTONDBLCLK = &H209
'Public Const WM_MOUSELAST = &H209

' Flags set when one of the mouse events is triggered
Public Const MK_LBUTTON = &H1
Public Const MK_MBUTTON = &H10
Public Const MK_RBUTTON = &H2

' Keyboard events for when a key is pressed/released
'Public Const WM_KEYDOWN = &H100
'Public Const WM_KEYUP = &H101

' Flags set when a keyboard event is triggered
Public Const MK_ALT = &H20
Public Const MK_CONTROL = &H8
Public Const MK_SHIFT = &H4

' The Virtual Key codes
Public Const VK_ESCAPE = &H1B  'The <Esc> key (ASCII Character 27)
Public Const VK_SHIFT = &H10   'The <Shift> key
Public Const VK_CONTROL = &H11 'The <Ctrl> key







'System API functions and declarations-----------------------------------------------------------------------

Public Type POINTAPI
    X As Long
    Y As Long
End Type
Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type
Type WINDOWPLACEMENT
        Length As Long
        Flags As Long
        showCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As RECT
End Type

Public cpoint As POINTAPI

Global Const gintMAX_SIZE% = 255                        ' 最大缓冲区大小

' ShowWindow() Commands
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_NORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
Public Const SW_MAX = 10


' SetWindowPos Flags
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering

Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

'Mouse Pointer Constants
Global Const gintMOUSE_DEFAULT% = 0

'MsgError() Constants
Global Const MSGERR_ERROR = 1
Global Const MSGERR_WARNING = 2


' SetWindowPos() hwndInsertAfter values
Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
' GetWindow() Constants
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_OWNER = 4
Public Const GW_CHILD = 5
Public Const GW_MAX = 5
' Window Messages
Public Const WM_NULL = &H0
Public Const WM_CREATE = &H1
Public Const WM_DESTROY = &H2
Public Const WM_MOVE = &H3
Public Const WM_SIZE = &H5

Public Const WM_ACTIVATE = &H6
'
'  WM_ACTIVATE state values

Public Const WA_INACTIVE = 0
Public Const WA_ACTIVE = 1
Public Const WA_CLICKACTIVE = 2

Public Const WM_SETFOCUS = &H7
Public Const WM_KILLFOCUS = &H8
Public Const WM_ENABLE = &HA
Public Const WM_SETREDRAW = &HB
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_PAINT = &HF
Public Const WM_CLOSE = &H10
Public Const WM_QUERYENDSESSION = &H11
Public Const WM_QUIT = &H12
Public Const WM_QUERYOPEN = &H13
Public Const WM_ERASEBKGND = &H14
Public Const WM_SYSCOLORCHANGE = &H15
Public Const WM_ENDSESSION = &H16
Public Const WM_SHOWWINDOW = &H18
Public Const WM_WININICHANGE = &H1A
Public Const WM_DEVMODECHANGE = &H1B
Public Const WM_ACTIVATEAPP = &H1C
Public Const WM_FONTCHANGE = &H1D
Public Const WM_TIMECHANGE = &H1E
Public Const WM_CANCELMODE = &H1F
Public Const WM_SETCURSOR = &H20
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_CHILDACTIVATE = &H22
Public Const WM_QUEUESYNC = &H23

Public Const WM_GETMINMAXINFO = &H24

Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type

Public Const WM_PAINTICON = &H26
Public Const WM_ICONERASEBKGND = &H27
Public Const WM_NEXTDLGCTL = &H28
Public Const WM_SPOOLERSTATUS = &H2A
Public Const WM_DRAWITEM = &H2B
Public Const WM_MEASUREITEM = &H2C
Public Const WM_DELETEITEM = &H2D
Public Const WM_VKEYTOITEM = &H2E
Public Const WM_CHARTOITEM = &H2F
Public Const WM_SETFONT = &H30
Public Const WM_GETFONT = &H31
Public Const WM_SETHOTKEY = &H32
Public Const WM_GETHOTKEY = &H33
Public Const WM_QUERYDRAGICON = &H37
Public Const WM_COMPAREITEM = &H39
Public Const WM_COMPACTING = &H41
Public Const WM_OTHERWINDOWCREATED = &H42               '  no longer suported
Public Const WM_OTHERWINDOWDESTROYED = &H43             '  no longer suported
Public Const WM_COMMNOTIFY = &H44                       '  no longer suported

' notifications passed in low word of lParam on WM_COMMNOTIFY messages
Public Const CN_RECEIVE = &H1
Public Const CN_TRANSMIT = &H2
Public Const CN_EVENT = &H4

Public Const WM_WINDOWPOSCHANGING = &H46
Public Const WM_WINDOWPOSCHANGED = &H47

Public Const WM_POWER = &H48
'
'  wParam for WM_POWER window message and DRV_POWER driver notification

Public Const PWR_OK = 1
Public Const PWR_FAIL = (-1)
Public Const PWR_SUSPENDREQUEST = 1
Public Const PWR_SUSPENDRESUME = 2
Public Const PWR_CRITICALRESUME = 3

Public Const WM_COPYDATA = &H4A
Public Const WM_CANCELJOURNAL = &H4B

Type COPYDATASTRUCT
        dwData As Long
        cbData As Long
        lpData As Long
End Type

Public Const WM_NCCREATE = &H81
Public Const WM_NCDESTROY = &H82
Public Const WM_NCCALCSIZE = &H83
Public Const WM_NCHITTEST = &H84
Public Const WM_NCPAINT = &H85
Public Const WM_NCACTIVATE = &H86
Public Const WM_GETDLGCODE = &H87
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONUP = &HA2
Public Const WM_NCLBUTTONDBLCLK = &HA3
Public Const WM_NCRBUTTONDOWN = &HA4
Public Const WM_NCRBUTTONUP = &HA5
Public Const WM_NCRBUTTONDBLCLK = &HA6
Public Const WM_NCMBUTTONDOWN = &HA7
Public Const WM_NCMBUTTONUP = &HA8
Public Const WM_NCMBUTTONDBLCLK = &HA9

Public Const WM_KEYFIRST = &H100
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CHAR = &H102
Public Const WM_DEADCHAR = &H103
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const WM_SYSCHAR = &H106
Public Const WM_SYSDEADCHAR = &H107
Public Const WM_KEYLAST = &H108
Public Const WM_INITDIALOG = &H110
Public Const WM_COMMAND = &H111
Public Const WM_SYSCOMMAND = &H112
Public Const WM_TIMER = &H113
Public Const WM_HSCROLL = &H114
Public Const WM_VSCROLL = &H115
Public Const WM_INITMENU = &H116
Public Const WM_INITMENUPOPUP = &H117
Public Const WM_MENUSELECT = &H11F
Public Const WM_MENUCHAR = &H120
Public Const WM_ENTERIDLE = &H121

Public Const WM_CTLCOLORMSGBOX = &H132
Public Const WM_CTLCOLOREDIT = &H133
Public Const WM_CTLCOLORLISTBOX = &H134
Public Const WM_CTLCOLORBTN = &H135
Public Const WM_CTLCOLORDLG = &H136
Public Const WM_CTLCOLORSCROLLBAR = &H137
Public Const WM_CTLCOLORSTATIC = &H138

Public Const WM_MOUSEFIRST = &H200
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MOUSELAST = &H209

Public Const WM_PARENTNOTIFY = &H210
Public Const WM_ENTERMENULOOP = &H211
Public Const WM_EXITMENULOOP = &H212
Public Const WM_MDICREATE = &H220
Public Const WM_MDIDESTROY = &H221
Public Const WM_MDIACTIVATE = &H222
Public Const WM_MDIRESTORE = &H223
Public Const WM_MDINEXT = &H224
Public Const WM_MDIMAXIMIZE = &H225
Public Const WM_MDITILE = &H226
Public Const WM_MDICASCADE = &H227
Public Const WM_MDIICONARRANGE = &H228
Public Const WM_MDIGETACTIVE = &H229
Public Const WM_MDISETMENU = &H230
Public Const WM_DROPFILES = &H233
Public Const WM_MDIREFRESHMENU = &H234


Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public Const WM_CLEAR = &H303
Public Const WM_UNDO = &H304
Public Const WM_RENDERFORMAT = &H305
Public Const WM_RENDERALLFORMATS = &H306
Public Const WM_DESTROYCLIPBOARD = &H307
Public Const WM_DRAWCLIPBOARD = &H308
Public Const WM_PAINTCLIPBOARD = &H309
Public Const WM_VSCROLLCLIPBOARD = &H30A
Public Const WM_SIZECLIPBOARD = &H30B
Public Const WM_ASKCBFORMATNAME = &H30C
Public Const WM_CHANGECBCHAIN = &H30D
Public Const WM_HSCROLLCLIPBOARD = &H30E
Public Const WM_QUERYNEWPALETTE = &H30F
Public Const WM_PALETTEISCHANGING = &H310
Public Const WM_PALETTECHANGED = &H311
Public Const WM_HOTKEY = &H312

Public Const WM_PENWINFIRST = &H380
Public Const WM_PENWINLAST = &H38F

' NOTE: All Message Numbers below 0x0400 are RESERVED.

Global Const gstrNULL$ = ""                             ' 空字符串

Global Const gstrSEP_DIR$ = "\"                         ' 目录分隔符
Public Const gstrSEP_REGKEY$ = "\"                      ' 注册关键字分隔符
Global Const gstrSEP_DRIVE$ = ":"                       ' 驱动器分隔符,例如 C:\
Global Const gstrSEP_DIRALT$ = "/"                      ' 备用目录分隔符
Global Const gstrSEP_EXT$ = "."                         ' 文件扩展名分隔符
Public Const gstrSEP_PROGID = "."

Public Type STARTUPINFO
    cb              As Long
    lpReserved      As Long
    lpDesktop       As Long
    lpTitle         As Long
    dwX             As Long
    dwY             As Long
    dwXSize         As Long
    dwYSize         As Long
    dwXCountChars   As Long
    dwYCountChars   As Long
    dwFillAttribute As Long
    dwFlags         As Long
    wShowWindow     As Integer
    cbReserved2     As Integer
    lpReserved2     As Long
    hStdInput       As Long
    hStdOutput      As Long
    hStdError       As Long
End Type

Public Type PROCESS_INFORMATION
    hProcess    As Long
    hThread     As Long
    dwProcessID As Long
    dwThreadID  As Long
End Type

Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    nReserved1 As Integer
    nReserved2 As Integer
    szPathName As String * 256
End Type

Type VERINFO                                            'Version FIXEDFILEINFO
    'There is data in the following two dwords, but it is for Windows internal
    '   use and we should ignore it
    Ignore(1 To 8) As Byte
    'Signature As Long
    'StrucVersion As Long
    FileVerPart2 As Integer
    FileVerPart1 As Integer
    FileVerPart4 As Integer
    FileVerPart3 As Integer
    ProductVerPart2 As Integer
    ProductVerPart1 As Integer
    ProductVerPart4 As Integer
    ProductVerPart3 As Integer
    FileFlagsMask As Long 'VersionFileFlags
    FileFlags As Long 'VersionFileFlags
    FileOS As Long 'VersionOperatingSystemTypes
    FileType As Long
    FileSubtype As Long 'VersionFileSubTypes
    'I've never seen any data in the following two dwords, so I'll ignore them
    Ignored(1 To 8) As Byte 'DateHighPart As Long, DateLowPart As Long
End Type

Type PROTOCOL
    strName As String
    strFriendlyName As String
End Type

Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Type LOCALESIGNATURE
    lsUsb(3)          As Long
    lsCsbDefault(1)   As Long
    lsCsbSupported(1) As Long
End Type
Type FONTSIGNATURE
    fsUsb(3) As Long
    fsCsb(1) As Long
End Type
Type CHARSETINFO
    ciCharset As Long
    ciACP     As Long
    fs        As FONTSIGNATURE
End Type

Global Const OF_EXIST& = &H4000&
Global Const OF_SEARCH& = &H400&
Global Const HFILE_ERROR% = -1

'
' Global variables used for silent and SMS installation
'
Public gfSilent As Boolean                              ' Whether or not we are doing a silent install
Public gstrSilentLog As String                          ' filename for output during silent install.
Public gfSMS As Boolean                                 ' Whether or not we are doing an SMS silent install
Public gstrMIFFile As String                            ' status output file for SMS
Public gfSMSStatus As Boolean                           ' status of SMS installation
Public gstrSMSDescription As String                     ' description string written to MIF file for SMS installation
Public gfNoUserInput As Boolean                         ' True if either gfSMS or gfSilent is True
Public gfDontLogSMS As Boolean                          ' Prevents MsgFunc from being logged to SMS (e.g., for confirmation messasges)
Global Const MAX_SMS_DESCRIP = 255                      ' SMS does not allow description strings longer than 255 chars.

'Variables for caching font values
Private m_sFont As String                   ' the cached name of the font
Private m_nFont As Integer                  ' the cached size of the font
Private m_nCharset As Integer               ' the cached charset of the font


Global Const gsSTARTMENUKEY As String = "$(Start Menu)"
Global Const gsPROGMENUKEY As String = "$(Programs)"
Global Const gsSTART As String = "StartIn"
Global Const gsPARENT As String = "Parent"

'
'List of available protocols
'
Global gProtocol() As PROTOCOL
Global gcProtocols As Integer
'
' AXDist.exe and wint351.exe needed.  These are self extracting exes
' that install other files not installed by setup1.
'
Public gfAXDist As Boolean
Global Const gstrFILE_AXDIST = "AXDIST.EXE"
Public gstrAXDISTInstallPath As String
Public gfAXDistChecked As Boolean
Public gfMDag As Boolean
Global Const gstrFILE_MDAG = "mdac_typ.exe"
Global Const gstrFILE_MDAGARGS = " /q:a /c:""setup.exe /QN1"""
Public gstrMDagInstallPath As String
Public gfWINt351 As Boolean
Global Const gstrFILE_WINT351 = "WINt351.EXE"
Public gstrWINt351InstallPath As String
Public gfWINt351Checked As Boolean
Declare Function GetUserDefaultLCID Lib "Kernel32" () As Long
Declare Sub GetLocaleInfoA Lib "Kernel32" (ByVal lLCID As Long, ByVal lLCTYPE As Long, ByVal strLCData As String, ByVal lDataLen As Long)

Declare Function GetDiskFreeSpace Lib "Kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
Declare Function GetSystemDirectory Lib "Kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetFullPathName Lib "Kernel32" Alias "GetFullPathNameA" (ByVal lpFilename As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal lSize As Long, ByVal lpFilename As String) As Long
Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFilename As String) As Long
Declare Function GetPrivateProfileSection Lib "Kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Declare Function GetPrivateProfileSectionNames Lib "Kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long


Declare Function VerInstallFile Lib "version.dll" Alias "VerInstallFileA" (ByVal Flags&, ByVal SrcName$, ByVal DestName$, ByVal SrcDir$, ByVal DestDir$, ByVal CurrDir As Any, ByVal TmpName$, lpTmpFileLen&) As Long
Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal sFile As String, lpLen) As Long
Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal sFile As String, ByVal lpIgnored As Long, ByVal lpSize As Long, ByVal lpBuf As Long) As Long
Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (ByVal lpBuf As Long, ByVal szReceive As String, lpBufPtr As Long, lLen As Long) As Long
Declare Function OSGetShortPathName Lib "Kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Declare Function GetDriveType32 Lib "Kernel32" Alias "GetDriveTypeA" (ByVal strWhichDrive As String) As Long
Declare Function GetTempFilename32 Lib "Kernel32" Alias "GetTempFileNameA" (ByVal strWhichDrive As String, ByVal lpPrefixString As String, ByVal wUnique As Integer, ByVal lpTempFilename As String) As Long
Declare Function GetTempPath Lib "Kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Global Const LB_FINDSTRINGEXACT = &H1A2
Global Const LB_ERR = (-1)

' Reboot system code
Public Const EWX_REBOOT = 2
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Public Const Flags = SWP_SHOWWINDOW Or SWP_FRAMECHANGED

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Sub SetActiveWindow Lib "user32" (ByVal hwnd1 As Long)
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd1 As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd1 As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd1 As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd1 As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function SetFocus Lib "user32" (ByVal hwnd1 As Long) As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Public Declare Function GetCursor Lib "user32" () As Long
Public Declare Function GetClipCursor Lib "user32" (lprc As RECT) As Long
Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Public brushhwnd As Long, curfilename As String, Searchkey As Integer, ArgNumber As Integer, argarray() As String 'ArgNumber记录命令行参数个数

    Function GetCommandLine(Optional MaxArgs)
        Dim C, CmdLine, CmdLnLen, InArg, I, NumArgs
        '检查是否提供了 MaxArgs 参数。
        If IsMissing(MaxArgs) Then MaxArgs = 10
        ' 使数组的大小合适。
        ReDim Preserve argarray(MaxArgs)
        NumArgs = 0: InArg = False
        '取得命令行参数。
        CmdLine = Command()
        CmdLnLen = Len(CmdLine)
        '以一次一个字符的方式取出命令行参数。
        For I = 1 To CmdLnLen
            C = Mid(CmdLine, I, 1)
     '检测是否为 space 或 tab。
            If (C <> " " And C <> vbTab) Then
                '若既不是 space 键，也不是 tab 键，则检测是否为参数内含之字符。
                If Not InArg Then
                '新的参数。检测参数是否过多。
                    If NumArgs = MaxArgs Then Exit For
                        NumArgs = NumArgs + 1
                        InArg = True
                    End If
                '将字符加到当前参数中。
                argarray(NumArgs) = argarray(NumArgs) + C
            Else
                '找到 space 或 tab字符,将 InArg 标志设置成 False。
                InArg = False
            End If
        Next I
        '调整数组大小使其刚好符合参数个数。
        ReDim Preserve argarray(NumArgs)
        '将数组和参数个数返回。
        GetCommandLine = argarray()
        ArgNumber = NumArgs
     End Function



'-----------------------------------------------------------
' 函数: StripTerminator
'
' 返回非零结尾的字符串。典型地，这是一个由 Windows API 调用返回的字符串。
'
' 入口: [strString] - 要删除结束符的字符串
'
' 返回: 传递的字符串减去尾部零以后的值。
'-----------------------------------------------------------
'
Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function


Public Sub loadbrush()
  Dim X, delaystart, delayend
  
  delaystart = Timer
  delayend = Timer
  On Error GoTo errorhandler2
  If checkbrush() = 0 Then
    If ArgNumber > 0 Then
       curfilename = argarray(1)
       If UCase$(Extension(curfilename)) <> "BMP" Then
         GoTo errorhandler1
       End If
       If FileExist(curfilename) > 0 Then
          If Not IsWindowsNT() Then
             X = Shell("pbrush.exe " + curfilename, 3)
          Else
             X = Shell("mspaint.exe " + curfilename, 3)
          End If
       Else
         If DirExists(GetPathName(curfilename)) = False Then
            If GetPathName(curfilename) <> "" Then MkDir GetPathName(curfilename)
         End If
         FileCopy App.Path + "\systmp.bmp", curfilename
         Reset  '刷新
         Do While FileExist(curfilename) = 0
           If ((delayend - delaystart) < 10) Then '10 seconds
             delayend = Timer
           Else
             'MsgBox "在系统中找不到‘画图’程序！"
            GoTo errorhandler2: Exit Do
           End If
         Loop
          If Not IsWindowsNT() Then
             X = Shell("pbrush.exe " + curfilename, 3)
          Else
             X = Shell("mspaint.exe " + curfilename, 3)
          End If
       End If
    Else
          If Not IsWindowsNT() Then
             X = Shell("pbrush.exe ", 3)
          Else
             X = Shell("mspaint.exe ", 3)
          End If
    End If
'    Do While ((delayend - delaystart) < 3)
'        delayend = Timer
'    Loop
    brushhwnd = 0
    Do While brushhwnd = 0
      Call GetHwnd(brushhwnd, "画图", "- 画图")
      'brushhwnd = FindWindow("MSPaintApp", "未命名 - 画图")
      'GetClassName(y, z, 30)
      If ((delayend - delaystart) < 10) Then '10 seconds
        delayend = Timer
      Else
        'MsgBox "在系统中找不到‘画图’程序！"
        GoTo errorhandler2: Exit Do
      End If
    Loop
    SetWindowPos brushhwnd, HWND_TOP, 0, 0, Screen.Width, Screen.Height, Flags
    'SendKeys "% X", True
    ShowWindow brushwnd, SW_SHOWMAXIMIZED
  Else
    SetWindowPos brushhwnd, HWND_TOP, 0, 0, Screen.Width, Screen.Height, Flags
   ' SendKeys "% X", True
   ' ShowWindow brushwnd, SW_HIDE
    ShowWindow brushwnd, SW_SHOWMAXIMIZED
   ' SetActiveWindow brushhwnd
    SetFocus brushhwnd
    'SendMessage brushhwnd, WM_LBUTTONDOWN, 2, 0
    'SendMessage brushhwnd, WM_LBUTTONUP, 2, 0
  End If
  Exit Sub
errorhandler1:
  MsgBox curfilename + " 文件名不正确 (应给出BMP扩展名和完整路径)！" + Chr$(10) + Chr$(13) + "很抱歉，系统将退出。"
  End   '系统退出
errorhandler2:
  MsgBox "文件名不正确 或 在系统中找不到‘画图’程序！" + Chr$(10) + Chr$(13) + "很抱歉，系统将退出。"
  End   '系统退出
End Sub

Public Function checkbrush() As Long
  On Error GoTo errorhandler1
  checkbrush = 0
  If brushhwnd > 0 Then
    checkbrush = 1
  End If
  If IsWindow(brushhwnd) = 0 Then
    checkbrush = 0
  End If
  Exit Function
errorhandler1:
  checkbrush = 0
End Function

Public Sub delay(t As Single)
  Dim delaystart, delayend
  If t <= 0 Then Exit Sub
  delaystart = Timer
  delayend = Timer
  Do While ((delayend - delaystart) < t)  't seconds
    delayend = Timer
  Loop
End Sub

Sub senddata()
 Dim WinStr As String
  loadbrush
'  SendKeys "%E"
  SetFocus brushhwnd
  WinStr = GETWINTEXT(brushhwnd)
  AppActivate WinStr

  'SetActiveWindow brushhwnd
  'Call delay(1)
  If brushhwnd > 0 Then
    Call PostMessage(brushhwnd, WM_KEYDOWN, &H11, &H1D0001) '< 0 Then MsgBox "DON'T SUCCESS!"
    Call PostMessage(brushhwnd, WM_KEYDOWN, &H11, &H401D0001)
    Call PostMessage(brushhwnd, WM_KEYDOWN, &H56, &H2F0001)
    Call PostMessage(brushhwnd, WM_KEYUP, &H56, &HC82F0001)
    Call PostMessage(brushhwnd, WM_KEYUP, &H11, &HC81D0001)
    
    Call SendMessage(brushhwnd, WM_INITMENU, &H34C, &H0)
    Call SendMessage(brushhwnd, WM_INITMENUPOPUP, &H354, &H1)
    Call SendMessage(brushhwnd, WM_COMMAND, &H1E125, &H7FF6B2)
    
    Call SendMessage(brushhwnd, WM_MENUSELECT, &H80900001, &HCE4)
    Call SendMessage(brushhwnd, WM_MENUSELECT, &H8080E125, &HCDC)
  End If
End Sub

Function GETWINTEXT(hwnd As Long) As String
  Dim strBuf As String
   
   GETWINTEXT = ""
   If hwnd <= 0 Then Return
   strBuf = Space$(gintMAX_SIZE)
   GetWindowText hwnd, strBuf, gintMAX_SIZE
   strBuf = StripTerminator$(strBuf)
   GETWINTEXT = strBuf
End Function

Sub AppSetFocus(hwnd As Long)
Dim strstr As String

  strstr = GETWINTEXT(hwnd)
  AppActivate strstr
  AppActivate strstr
  AppActivate strstr
End Sub

Public Function FileExist(name As String) As Integer
 Dim filen As Integer
 
    On Error GoTo err1
    filen = FreeFile
    Open name For Input As filen
    Close filen
    FileExist = 1
    Exit Function
err1:
'    MsgBox name & "不存在！"
    FileExist = 0
End Function

Sub GetHwnd(hwnd As Long, str1 As String, str2 As String)
  Dim strBuf As String
  Dim CurHwnd As Long
  str1 = UCase$(str1)
  str2 = UCase$(str2)
  
  Searchkey = 0
  hwnd = 0
  strBuf = Space$(gintMAX_SIZE)
  
  CurHwnd = Fsketch.hwnd   '任一个已知窗口句柄
  Do While GetWindow(CurHwnd, GW_HWNDNEXT) > 0
    CurHwnd = GetWindow(CurHwnd, GW_HWNDNEXT)
    strBuf = Space$(gintMAX_SIZE)
    GetWindowText CurHwnd, strBuf, gintMAX_SIZE
    strBuf = StripTerminator$(strBuf)
    If UCase$(strBuf) = str1 Or InStr(UCase$(strBuf), str2) > 0 Then
       Searchkey = 1
       hwnd = CurHwnd
       Exit Sub
    End If
  Loop
  CurHwnd = Fsketch.hwnd
  Do While GetWindow(CurHwnd, GW_HWNDPREV) > 0
    CurHwnd = GetWindow(CurHwnd, GW_HWNDPREV)
    strBuf = Space$(gintMAX_SIZE)
    GetWindowText CurHwnd, strBuf, gintMAX_SIZE
    strBuf = StripTerminator$(strBuf)
    If UCase$(strBuf) = str1 Or InStr(UCase$(strBuf), str2) > 0 Then
       Searchkey = 1
       hwnd = CurHwnd
       Exit Sub
    End If
  Loop
'  If Searchkey = 0 Then
'    MsgBox "在系统中找不到" & str1 & "程序！"
'  End If
End Sub


'-----------------------------------------------------------
' 函数: Extension
'
' 抽取文件名/路径名的扩展部分
'
' 入口: [strFilename] - 要获取扩展部分的文件/路径
'
' 返回: 扩展部分，如果存在；否则 gstrNULL
'-----------------------------------------------------------
'
Function Extension(ByVal strFilename As String) As String
    Dim intPos As Integer

    Extension = gstrNULL

    intPos = Len(strFilename)

    Do While intPos > 0
        Select Case Mid$(strFilename, intPos, 1)
            Case gstrSEP_EXT
                Extension = Mid$(strFilename, intPos + 1)
                Exit Do
            Case gstrSEP_DIR, gstrSEP_DIRALT
                Exit Do
            '结束 Case
        End Select

        intPos = intPos - 1
    Loop
End Function


'-----------------------------------------------------------
' 函数: GetPath
'
' 抽取文件名/路径名的路径部分
'
' 入口: [strFilename] - 要获取路径部分的文件/路径
'
' 返回: 扩展部分，如果存在；否则 gstrNULL
'-----------------------------------------------------------
'
Function GetPath(ByVal strFilename As String) As String
    Dim intPos As Integer

    GetPath = gstrNULL

    intPos = Len(strFilename)

    Do While intPos > 0
        Select Case Mid$(strFilename, intPos, 1)
            Case gstrSEP_DIR, gstrSEP_DIRALT
                GetPath = left$(strFilename, intPos)
                Exit Do
            '结束 Case
        End Select

        intPos = intPos - 1
    Loop
End Function

'-----------------------------------------------------------
' 函数: GetMainName
'
' 抽取文件名/路径名的主名部分
'
' 入口: [strFilename] - 要获取扩展部分的文件/路径
'
' 返回: 主名部分，如果存在；否则 gstrNULL
'-----------------------------------------------------------
'
Function GetMainName(ByVal strFilename As String) As String
    Dim intPos As Integer, intPos1 As Integer, Ext As String, Pat As String

    GetMainName = gstrNULL

    intPos = Len(strFilename)
    Ext = ""

    Do While intPos > 0
        Select Case Mid$(strFilename, intPos, 1)
            Case gstrSEP_EXT
                Ext = Mid$(strFilename, intPos + 1)
                Exit Do
            Case gstrSEP_DIR, gstrSEP_DIRALT
                Exit Do
            '结束 Case
        End Select

        intPos = intPos - 1
    Loop
    If intPos < 0 Then intPos = Len(strFilename)
        
    intPos1 = Len(strFilename)
    Pat = ""
    Do While intPos1 > 0
        Select Case Mid$(strFilename, intPos1, 1)
            Case gstrSEP_DIR, gstrSEP_DIRALT
                Pat = left$(strFilename, intPos1)
                Exit Do
            '结束 Case
        End Select

        intPos1 = intPos1 - 1
    Loop
    If intPos1 < 0 Then intPos1 = 0
    
    If intPos = Len(strFilename) Then
      GetMainName = Mid$(strFilename, intPos1 + 1)
    Else
      GetMainName = Mid$(strFilename, intPos1 + 1, intPos - intPos1 - 1)
    End If
      
    
End Function


'得到文件名中的路径部分
Function GetPathName(ByVal strFilename As String) As String
    Dim intPos As Integer
    Dim strPathOnly As String
    Dim dirTmp As DirListBox
    Dim I As Integer

    On Error Resume Next


    Err = 0
    
    intPos = Len(strFilename)

    '
    '改变所有 '/' 字符为 '\'
    '

    For I = 1 To Len(strFilename)
        If Mid$(strFilename, I, 1) = gstrSEP_DIRALT Then
            Mid$(strFilename, I, 1) = gstrSEP_DIR
        End If
    Next I

    If InStr(strFilename, gstrSEP_DIR) = intPos Then
        If intPos > 1 Then
            intPos = intPos - 1
        End If
    Else
        Do While intPos > 0
            If Mid$(strFilename, intPos, 1) <> gstrSEP_DIR Then
                intPos = intPos - 1
            Else
                Exit Do
            End If
        Loop
    End If

    If intPos > 0 Then
        strPathOnly = left$(strFilename, intPos)
        If right$(strPathOnly, 1) = gstrCOLON Then
            strPathOnly = strPathOnly & gstrSEP_DIR
        End If
    Else
        strPathOnly = CurDir$
    End If

    If right$(strPathOnly, 1) = gstrSEP_DIR Then
        strPathOnly = left$(strPathOnly, Len(strPathOnly) - 1)
    End If

    GetPathName = UCase(strPathOnly)
    
    Err = 0
End Function

'-----------------------------------------------------------
' 子程序: AddDirSep
' 添加一个尾随的目录路径分隔符 (反斜杠) 到路径名末端，除非已经存在分隔符
'
' 入口/出口: [strPathName] - 要添加分隔符的路径
'-----------------------------------------------------------
'
Sub AddDirSep(strPathName As String)
    If right(Trim(strPathName), Len(gstrSEP_URLDIR)) <> gstrSEP_URLDIR And _
       right(Trim(strPathName), Len(gstrSEP_DIR)) <> gstrSEP_DIR Then
        strPathName = RTrim$(strPathName) & gstrSEP_DIR
    End If
End Sub
'-----------------------------------------------------------
' 子程序: AddURLDirSep
' 添加一个尾随的 URL 路径分隔符 (正斜杠) 到 URL 末端，
' 除非已经存在分隔符 (或反斜杠)
'
' 入口/出口: [strPathName] - 要添加分隔符的路径
'-----------------------------------------------------------
'
Sub AddURLDirSep(strPathName As String)
    If right(Trim(strPathName), Len(gstrSEP_URLDIR)) <> gstrSEP_URLDIR And _
       right(Trim(strPathName), Len(gstrSEP_DIR)) <> gstrSEP_DIR Then
        strPathName = Trim(strPathName) & gstrSEP_URLDIR
    End If
End Sub

'-----------------------------------------------------------
' 函数: FileExists
' 判断是否存在指定的文件
'
' 入口: [strPathName] - 要检查的文件
'
' 返回: True，如果文件存在；否则为 False
'-----------------------------------------------------------
'
Function FileExists(ByVal strPathName As String) As Integer
    Dim intFileNum As Integer

    On Error Resume Next

    '
    ' 如果引用了字符串，删除引用
    '
    strPathName = strUnQuoteString(strPathName)
    '
    '删除所有尾随的目录分隔符
    '
    If right$(strPathName, 1) = gstrSEP_DIR Then
        strPathName = left$(strPathName, Len(strPathName) - 1)
    End If

    '
    '试图打开文件，本函数的返回值为 False
    '如果打开时出错，否则为 True
    '
    intFileNum = FreeFile
    Open strPathName For Input As intFileNum

    FileExists = IIf(Err = 0, True, False)

    Close intFileNum

    Err = 0
End Function

'-----------------------------------------------------------
' 函数: DirExists
'
' 判断是否存在指定的目录名。
' 本函数用于 (例如) 通过传递的如 'A:\'，判断安装软盘是否在驱动器中。
'
' 入口: [strDirName] - 要检查的目录名
'
' 返回: True，如果目录存在；否则为 False
'-----------------------------------------------------------
'
Public Function DirExists(ByVal strDirName As String) As Integer
    Const strWILDCARD$ = "*.*"

    Dim strDummy As String

    On Error Resume Next

    AddDirSep strDirName
    strDummy = Dir$(strDirName & strWILDCARD, vbDirectory)
    DirExists = Not (strDummy = gstrNULL)

    Err = 0
End Function

'-----------------------------------------------------------
' FUNCTION: GetFileSize
'
' 决定指定文件的大小（字节数）
'
' 入口: [strFileName] - 要得到其大小的文件名
'
' 返回值: 文件大小的字节数，如果出错则返回 -1。
'-----------------------------------------------------------
'
Function GetFileSize(strFilename As String) As Long
    On Error Resume Next

    GetFileSize = filelen(strFilename)

    If Err > 0 Then
        GetFileSize = -1
        Err = 0
    End If
End Function

Public Function strUnQuoteString(ByVal strQuotedString As String)
'
' 本子程序测试 strQuotedString 是否在引号中折行，如果是，删除之。
'
    strQuotedString = Trim(strQuotedString)

    If Mid$(strQuotedString, 1, 1) = gstrQUOTE And right$(strQuotedString, 1) = gstrQUOTE Then
        '
        ' 如果有引号，去掉引号。
        '
        strQuotedString = Mid$(strQuotedString, 2, Len(strQuotedString) - 2)
    End If
    strUnQuoteString = strQuotedString
End Function

'----------------------------------------------------------
' FUNCTION: GetWinPlatform
' Get the current windows platform.
' ---------------------------------------------------------
Public Function GetWinPlatform() As Long

    Dim osvi As OSVERSIONINFO
    Dim strCSDVersion As String
    osvi.dwOSVersionInfoSize = Len(osvi)
    If GetVersionEx(osvi) = 0 Then
        Exit Function
    End If
    GetWinPlatform = osvi.dwPlatformId
End Function



'-----------------------------------------------------------
' FUNCTION: GetDriveType
' Determine whether a disk is fixed, removable, etc. by
' calling Windows GetDriveType()
'-----------------------------------------------------------
'
Function GetDriveType(ByVal intDriveNum As Integer) As Integer
    '
    ' This function expects an integer drive number in Win16 or a string in Win32
    '
    Dim strDriveName As String
    
    strDriveName = Chr$(Asc("A") + intDriveNum) & gstrSEP_DRIVE & gstrSEP_DIR
    GetDriveType = CInt(GetDriveType32(strDriveName))
End Function

'-----------------------------------------------------------
' FUNCTION: ReadProtocols
' Reads the allowable protocols from the specified file.
'
' IN: [strInputFilename] - INI filename from which to read the protocols
'     [strINISection] - Name of the INI section
'-----------------------------------------------------------
Function ReadProtocols(ByVal strInputFilename As String, ByVal strINISection As String) As Boolean
    Dim intIdx As Integer
    Dim fOk As Boolean
    Dim strInfo As String
    Dim intOffset As Integer
    
    intIdx = 0
    fOk = True
    Erase gProtocol
    gcProtocols = 0
    
    Do
        strInfo = ReadIniFile(strInputFilename, strINISection, gstrINI_PROTOCOL & Format$(intIdx + 1))
        If strInfo <> vbNullString Then
            intOffset = InStr(strInfo, gstrCOMMA)
            If intOffset > 0 Then
                'The "ugly" name will be first on the line
                ReDim Preserve gProtocol(intIdx + 1)
                gcProtocols = intIdx + 1
                gProtocol(intIdx + 1).strName = left$(strInfo, intOffset - 1)
                
                '... followed by the friendly name
                gProtocol(intIdx + 1).strFriendlyName = Mid$(strInfo, intOffset + 1)
                If (gProtocol(intIdx + 1).strName = "") Or (gProtocol(intIdx + 1).strFriendlyName = "") Then
                    fOk = False
                End If
            Else
                fOk = False
            End If

            If Not fOk Then
                Exit Do
            Else
                intIdx = intIdx + 1
            End If
        End If
    Loop While strInfo <> vbNullString
    
    ReadProtocols = fOk
End Function

'-----------------------------------------------------------
' FUNCTION: ResolveResString
' Reads resource and replaces given macros with given values
'
' Example, given a resource number 14:
'    "Could not read '|1' in drive |2"
'   The call
'     ResolveResString(14, "|1", "TXTFILE.TXT", "|2", "A:")
'   would return the string
'     "Could not read 'TXTFILE.TXT' in drive A:"
'
' IN: [resID] - resource identifier
'     [varReplacements] - pairs of macro/replacement value
'-----------------------------------------------------------
'
Public Function ResolveResString(ByVal resID As Integer, ParamArray varReplacements() As Variant) As String
    Dim intMacro As Integer
    Dim strResString As String
    
    strResString = LoadResString(resID)
    
    ' For each macro/value pair passed in...
    For intMacro = LBound(varReplacements) To UBound(varReplacements) Step 2
        Dim strMacro As String
        Dim strValue As String
        
        strMacro = varReplacements(intMacro)
        On Error GoTo MismatchedPairs
        strValue = varReplacements(intMacro + 1)
        On Error GoTo 0
        
        ' Replace all occurrences of strMacro with strValue
        Dim intPos As Integer
        Do
            intPos = InStr(strResString, strMacro)
            If intPos > 0 Then
                strResString = left$(strResString, intPos - 1) & strValue & right$(strResString, Len(strResString) - Len(strMacro) - intPos + 1)
            End If
        Loop Until intPos = 0
    Next intMacro
    
    ResolveResString = strResString
    
    Exit Function
    
MismatchedPairs:
    Resume Next
End Function
'-----------------------------------------------------------
' SUB: GetLicInfoFromVBL
' Parses a VBL file name and extracts the license key for
' the registry and license information.
'
' IN: [strVBLFile] - must be a valid VBL.
'
' OUT: [strLicKey] - registry key to write license info to.
'                    This key will be added to
'                    HKEY_CLASSES_ROOT\Licenses.  It is a
'                    guid.
' OUT: [strLicVal] - license information.  Usually in the
'                    form of a string of cryptic characters.
'-----------------------------------------------------------
'
Public Sub GetLicInfoFromVBL(strVBLFile As String, strLicKey As String, strLicVal As String)
    Dim fn As Integer
    Const strREGEDIT = "REGEDIT"
    Const strLICKEYBASE = "HKEY_CLASSES_ROOT\Licenses\"
    Dim strTemp As String
    Dim posEqual As Integer
    Dim fLicFound As Boolean
    
    fn = FreeFile
    Open strVBLFile For Input Access Read Lock Read Write As #fn
    '
    ' Read through the file until we find a line that starts with strLICKEYBASE
    '
    fLicFound = False
    Do While Not EOF(fn)
        Line Input #fn, strTemp
        strTemp = Trim(strTemp)
        If left$(strTemp, Len(strLICKEYBASE)) = strLICKEYBASE Then
            '
            ' We've got the line we want.
            '
            fLicFound = True
            Exit Do
        End If
    Loop

    Close fn
    
    If fLicFound Then
        '
        ' Parse the data on this line to split out the
        ' key and the license info.  The line should be
        ' the form of:
        ' "HKEY_CLASSES_ROOT\Licenses\<lickey> = <licval>"
        '
        posEqual = InStr(strTemp, gstrASSIGN)
        If posEqual > 0 Then
            strLicKey = Mid$(Trim(left$(strTemp, posEqual - 1)), Len(strLICKEYBASE) + 1)
            strLicVal = Trim(Mid$(strTemp, posEqual + 1))
        End If
    Else
        strLicKey = vbNullString
        strLicVal = vbNullString
    End If
End Sub

 '-----------------------------------------------------------
 ' FUNCTION GetLongPathName
 '
 ' Retrieve the long pathname version of a path possibly
 '   containing short subdirectory and/or file names
 '-----------------------------------------------------------
 '
 'Function GetLongPathName(ByVal strShortPath As String) As String
 '   On Error GoTo 0
 '
 '   MakeLongPath (strShortPath)
 '   GetLongPathName = StripTerminator(strShortPath)
 'End Function
 
 '-----------------------------------------------------------
 ' FUNCTION GetShortPathName
 '
 ' Retrieve the short pathname version of a path possibly
 '   containing long subdirectory and/or file names
 '-----------------------------------------------------------
 '
 'Function GetShortPathName(ByVal strLongPath As String) As String
 '    Const cchBuffer = 300
 '    Dim strShortPath As String
 '    Dim lResult As Long

 '    On Error GoTo 0
 '    strShortPath = String(cchBuffer, Chr$(0))
 '    lResult = OSGetShortPathName(strLongPath, strShortPath, cchBuffer)
 '    If lResult = 0 Then
 '        'Error 53 ' File not found
 '        'Vegas#51193, just use the long name as this is usually good enough
 '        GetShortPathName = strLongPath
 '    Else
 '        GetShortPathName = StripTerminator(strShortPath)
 '    End If
 'End Function
 
'-----------------------------------------------------------
' FUNCTION: GetTempFilename
' Get a temporary filename for a specified drive and
' filename prefix
' PARAMETERS:
'   strDestPath - Location where temporary file will be created.  If this
'                 is an empty string, then the location specified by the
'                 tmp or temp environment variable is used.
'   lpPrefixString - First three characters of this string will be part of
'                    temporary file name returned.
'   wUnique - Set to 0 to create unique filename.  Can also set to integer,
'             in which case temp file name is returned with that integer
'             as part of the name.
'   lpTempFilename - Temporary file name is returned as this variable.
' RETURN:
'   True if function succeeds; false otherwise
'-----------------------------------------------------------
'
Function GetTempFilename(ByVal strDestPath As String, ByVal lpPrefixString As String, ByVal wUnique As Integer, lpTempFilename As String) As Boolean
    If strDestPath = vbNullString Then
        '
        ' No destination was specified, use the temp directory.
        '
        strDestPath = String(gintMAX_PATH_LEN, vbNullChar)
        If GetTempPath(gintMAX_PATH_LEN, strDestPath) = 0 Then
            GetTempFilename = False
            Exit Function
        End If
    End If
    lpTempFilename = String(gintMAX_PATH_LEN, vbNullChar)
    GetTempFilename = GetTempFilename32(strDestPath, lpPrefixString, wUnique, lpTempFilename) > 0
    lpTempFilename = StripTerminator(lpTempFilename)
End Function
'-----------------------------------------------------------
' FUNCTION: GetDefMsgBoxButton
' Decode the flags passed to the MsgBox function to
' determine what the default button is.  Use this
' for silent installs.
'
' IN: [intFlags] - Flags passed to MsgBox
'
' Returns: VB defined number for button
'               vbOK        1   OK button pressed.
'               vbCancel    2   Cancel button pressed.
'               vbAbort     3   Abort button pressed.
'               vbRetry     4   Retry button pressed.
'               vbIgnore    5   Ignore button pressed.
'               vbYes       6   Yes button pressed.
'               vbNo        7   No button pressed.
'-----------------------------------------------------------
'
Function GetDefMsgBoxButton(intFlags) As Integer
    '
    ' First determine the ordinal of the default
    ' button on the message box.
    '
    Dim intButtonNum As Integer
    Dim intDefButton As Integer
    
    If (intFlags And vbDefaultButton2) = vbDefaultButton2 Then
        intButtonNum = 2
    ElseIf (intFlags And vbDefaultButton3) = vbDefaultButton3 Then
        intButtonNum = 3
    Else
        intButtonNum = 1
    End If
    '
    ' Now determine the type of message box we are dealing
    ' with and return the default button.
    '
    If (intFlags And vbRetryCancel) = vbRetryCancel Then
        intDefButton = IIf(intButtonNum = 1, vbRetry, vbCancel)
    ElseIf (intFlags And vbYesNoCancel) = vbYesNoCancel Then
        Select Case intButtonNum
            Case 1
                intDefButton = vbYes
            Case 2
                intDefButton = vbNo
            Case 3
                intDefButton = vbCancel
            'End Case
        End Select
    ElseIf (intFlags And vbOKCancel) = vbOKCancel Then
        intDefButton = IIf(intButtonNum = 1, vbOK, vbCancel)
    ElseIf (intFlags And vbAbortRetryIgnore) = vbAbortRetryIgnore Then
        Select Case intButtonNum
            Case 1
                intDefButton = vbAbort
            Case 2
                intDefButton = vbRetry
            Case 3
                intDefButton = vbIgnore
            'End Case
        End Select
    ElseIf (intFlags And vbYesNo) = vbYesNo Then
        intDefButton = IIf(intButtonNum = 1, vbYes, vbNo)
    Else
        intDefButton = vbOK
    End If
    
    GetDefMsgBoxButton = intDefButton
    
End Function
'-----------------------------------------------------------
' FUNCTION: GetDiskSpaceFree
' Get the amount of free disk space for the specified drive
'
' IN: [strDrive] - drive to check space for
'
' Returns: Amount of free disk space, or -1 if an error occurs
'-----------------------------------------------------------
'
Function GetDiskSpaceFree(ByVal strDrive As String) As Long
    Dim strCurDrive As String
    Dim lDiskFree As Long

    On Error Resume Next

    '
    'Save the current drive
    '
    strCurDrive = left$(CurDir$, 2)

    '
    'Fixup drive so it includes only a drive letter and a colon
    '
    If InStr(strDrive, gstrSEP_DRIVE) = 0 Or Len(strDrive) > 2 Then
        strDrive = left$(strDrive, 1) & gstrSEP_DRIVE
    End If

    '
    'Change to the drive we want to check space for.  The DiskSpaceFree() API
    'works on the current drive only.
    '
    ChDrive strDrive

    '
    'If we couldn't change to the request drive, it's an error, otherwise return
    'the amount of disk space free
    '
    If Err <> 0 Or (strDrive <> left$(CurDir$, 2)) Then
        lDiskFree = -1
    Else
        Dim lRet As Long
        Dim lBytes As Long, lSect As Long, lClust As Long, lTot As Long
        
        lRet = GetDiskFreeSpace(vbNullString, lSect, lBytes, lClust, lTot)
        On Error Resume Next
        lDiskFree = (lBytes * lSect) * lClust
        If Err Then lDiskFree = 2147483647
    End If

    If lDiskFree = -1 Then
        MsgError Error$ & vbLf & vbLf & ResolveResString(resDISKSPCERR) & strDrive, vbExclamation, gstrTitle
    End If

    GetDiskSpaceFree = lDiskFree

    '
    'Cleanup by setting the current drive back to the original
    '
    ChDrive strCurDrive

    Err = 0
End Function

'-----------------------------------------------------------
' FUNCTION: GetUNCShareName
'
' Given a UNC names, returns the leftmost portion of the
' directory representing the machine name and share name.
' E.g., given "\\SCHWEIZ\PUBLIC\APPS\LISTING.TXT", returns
' the string "\\SCHWEIZ\PUBLIC"
'
' Returns a string representing the machine and share name
'   if the path is a valid pathname, else returns NULL
'-----------------------------------------------------------
'
Function GetUNCShareName(ByVal strFN As String) As Variant
    GetUNCShareName = Null
    If IsUNCName(strFN) Then
        Dim iFirstSeparator As Integer
        iFirstSeparator = InStr(3, strFN, gstrSEP_DIR)
        If iFirstSeparator > 0 Then
            Dim iSecondSeparator As Integer
            iSecondSeparator = InStr(iFirstSeparator + 1, strFN, gstrSEP_DIR)
            If iSecondSeparator > 0 Then
                GetUNCShareName = left$(strFN, iSecondSeparator - 1)
            Else
                GetUNCShareName = strFN
            End If
        End If
    End If
End Function

'-----------------------------------------------------------
' FUNCTION: GetWindowsSysDir
'
' Calls the windows API to get the windows\SYSTEM directory
' and ensures that a trailing dir separator is present
'
' Returns: The windows\SYSTEM directory
'-----------------------------------------------------------
'
Function GetWindowsSysDir() As String
    Dim strBuf As String

    strBuf = Space$(gintMAX_SIZE)

    '
    'Get the system directory and then trim the buffer to the exact length
    'returned and add a dir sep (backslash) if the API didn't return one
    '
    If GetSystemDirectory(strBuf, gintMAX_SIZE) > 0 Then
        strBuf = StripTerminator(strBuf)
        AddDirSep strBuf
        
        GetWindowsSysDir = strBuf
    Else
        GetWindowsSysDir = vbNullString
    End If
End Function
'-----------------------------------------------------------
' SUB: TreatAsWin95
'
' Returns True iff either we're running under Windows 95
' or we are treating this version of NT as if it were
' Windows 95 for registry and application loggin and
' removal purposes.
'-----------------------------------------------------------
'
Function TreatAsWin95() As Boolean
    If IsWindows95() Then
        TreatAsWin95 = True
    ElseIf NTWithShell() Then
        TreatAsWin95 = True
    Else
        TreatAsWin95 = False
    End If
End Function
'-----------------------------------------------------------
' FUNCTION: NTWithShell
'
' Returns true if the system is on a machine running
' NT4.0 or greater.
'-----------------------------------------------------------
'
Function NTWithShell() As Boolean

    If Not IsWindowsNT() Then
        NTWithShell = False
        Exit Function
    End If
    
    Dim osvi As OSVERSIONINFO
    Dim strCSDVersion As String
    osvi.dwOSVersionInfoSize = Len(osvi)
    If GetVersionEx(osvi) = 0 Then
        Exit Function
    End If
    strCSDVersion = StripTerminator(osvi.szCSDVersion)
    
    'Is this Windows NT 4.0 or higher?
    Const NT4MajorVersion = 4
    Const NT4MinorVersion = 0
    If (osvi.dwMajorVersion >= NT4MajorVersion) Then
        NTWithShell = True
    Else
        NTWithShell = False
    End If
    
End Function
'-----------------------------------------------------------
' FUNCTION: IsDepFile
'
' Returns true if the file passed to this routine is a
' dependency (*.dep) file.  We make this determination
' by verifying that the extension is .dep and that it
' contains version information.
'-----------------------------------------------------------
'
'Function fIsDepFile(strFilename As String) As Boolean
'    Const strEXT_DEP = "DEP"
'
'    fIsDepFile = False
'
'    If UCase(Extension(strFilename)) = strEXT_DEP Then
'        If GetFileVersion(strFilename) <> vbNullString Then
'            fIsDepFile = True
'        End If
'    End If
'End Function

'-----------------------------------------------------------
' FUNCTION: IsWin32
'
' Returns true if this program is running under Win32 (i.e.
'   any 32-bit operating system)
'-----------------------------------------------------------
'
Function IsWin32() As Boolean
    IsWin32 = (IsWindows95() Or IsWindowsNT())
End Function

'-----------------------------------------------------------
' FUNCTION: IsWindows95
'
' Returns true if this program is running under Windows 95
'   or successor
'-----------------------------------------------------------
'
Function IsWindows95() As Boolean
    Const dwMask95 = &H1&
    IsWindows95 = (GetWinPlatform() And dwMask95)
End Function

'-----------------------------------------------------------
' FUNCTION: IsWindowsNT
'
' Returns true if this program is running under Windows NT
'-----------------------------------------------------------
'
Function IsWindowsNT() As Boolean
    Const dwMaskNT = &H2&
    IsWindowsNT = (GetWinPlatform() And dwMaskNT)
End Function

'-----------------------------------------------------------
' FUNCTION: IsWindowsNT4WithoutSP2
'
' Determines if the user is running under Windows NT 4.0
' but without Service Pack 2 (SP2).  If running under any
' other platform, returns False.
'
' IN: [none]
'
' Returns: True if and only if running under Windows NT 4.0
' without at least Service Pack 2 installed.
'-----------------------------------------------------------
'
Function IsWindowsNT4WithoutSP2() As Boolean
    IsWindowsNT4WithoutSP2 = False
    
    If Not IsWindowsNT() Then
        Exit Function
    End If
    
    Dim osvi As OSVERSIONINFO
    Dim strCSDVersion As String
    osvi.dwOSVersionInfoSize = Len(osvi)
    If GetVersionEx(osvi) = 0 Then
        Exit Function
    End If
    strCSDVersion = StripTerminator(osvi.szCSDVersion)
    
    'Is this Windows NT 4.0?
    Const NT4MajorVersion = 4
    Const NT4MinorVersion = 0
    If (osvi.dwMajorVersion <> NT4MajorVersion) Or (osvi.dwMinorVersion <> NT4MinorVersion) Then
        'No.  Return False.
        Exit Function
    End If
    
    'If no service pack is installed, or if Service Pack 1 is
    'installed, then return True.
    Const strSP1 = "SERVICE PACK 1"
    If strCSDVersion = "" Then
        IsWindowsNT4WithoutSP2 = True 'No service pack installed
    ElseIf strCSDVersion = strSP1 Then
        IsWindowsNT4WithoutSP2 = True 'Only SP1 installed
    End If
End Function

'-----------------------------------------------------------
' FUNCTION: IsUNCName
'
' Determines whether the pathname specified is a UNC name.
' UNC (Universal Naming Convention) names are typically
' used to specify machine resources, such as remote network
' shares, named pipes, etc.  An example of a UNC name is
' "\\SERVER\SHARE\FILENAME.EXT".
'
' IN: [strPathName] - pathname to check
'
' Returns: True if pathname is a UNC name, False otherwise
'-----------------------------------------------------------
'
Function IsUNCName(ByVal strPathName As String) As Integer
    Const strUNCNAME$ = "\\//\"        'so can check for \\, //, \/, /\

    IsUNCName = ((InStr(strUNCNAME, left$(strPathName, 2)) > 0) And _
                 (Len(strPathName) > 1))
End Function
'-----------------------------------------------------------
' FUNCTION: LogSilentMsg
'
' If this is a silent install, this routine writes
' a message to the gstrSilentLog file.
'
' IN: [strMsg] - The message
'
' Normally, this routine is called inlieu of displaying
' a MsgBox and strMsg is the same message that would
' have appeared in the MsgBox

'-----------------------------------------------------------
'
Sub LogSilentMsg(strMsg As String)
    If Not gfSilent Then Exit Sub
    
    Dim fn As Integer
    
    On Error Resume Next
    
    fn = FreeFile
    
    Open gstrSilentLog For Append As fn
    Print #fn, strMsg
    Close fn
    Exit Sub
End Sub
'-----------------------------------------------------------
' FUNCTION: LogSMSMsg
'
' If this is a SMS install, this routine appends
' a message to the gstrSMSDescription string.  This
' string will later be written to the SMS status
' file (*.MIF) when the installation completes (success
' or failure).
'
' Note that if gfSMS = False, not message will be logged.
' Therefore, to prevent some messages from being logged
' (e.g., confirmation only messages), temporarily set
' gfSMS = False.
'
' IN: [strMsg] - The message
'
' Normally, this routine is called inlieu of displaying
' a MsgBox and strMsg is the same message that would
' have appeared in the MsgBox
'-----------------------------------------------------------
'
Sub LogSMSMsg(strMsg As String)
    If Not gfSMS Then Exit Sub
    '
    ' Append the message.  Note that the total
    ' length cannot be more than 255 characters, so
    ' truncate anything after that.
    '
    gstrSMSDescription = left(gstrSMSDescription & strMsg, MAX_SMS_DESCRIP)
End Sub

'-----------------------------------------------------------
' FUNCTION: MakePathAux
'
' Creates the specified directory path.
'
' No user interaction occurs if an error is encountered.
' If user interaction is desired, use the related
'   MakePathAux() function.
'
' IN: [strDirName] - name of the dir path to make
'
' Returns: True if successful, False if error.
'-----------------------------------------------------------
'
Function MakePathAux(ByVal strDirName As String) As Boolean
    Dim strPath As String
    Dim intOffset As Integer
    Dim intAnchor As Integer
    Dim strOldPath As String

    On Error Resume Next

    '
    'Add trailing backslash
    '
    If right$(strDirName, 1) <> gstrSEP_DIR Then
        strDirName = strDirName & gstrSEP_DIR
    End If

    strOldPath = CurDir$
    MakePathAux = False
    intAnchor = 0

    '
    'Loop and make each subdir of the path separately.
    '
    intOffset = InStr(intAnchor + 1, strDirName, gstrSEP_DIR)
    intAnchor = intOffset 'Start with at least one backslash, i.e. "C:\FirstDir"
    Do
        intOffset = InStr(intAnchor + 1, strDirName, gstrSEP_DIR)
        intAnchor = intOffset

        If intAnchor > 0 Then
            strPath = left$(strDirName, intOffset - 1)
            ' Determine if this directory already exists
            Err = 0
            ChDir strPath
            If Err Then
                ' We must create this directory
                Err = 0
#If LOGGING Then
                NewAction gstrKEY_CREATEDIR, """" & strPath & """"
#End If
                MkDir strPath
#If LOGGING Then
                If Err Then
                    LogError ResolveResString(resMAKEDIR) & " " & strPath
                    AbortAction
                    GoTo Done
                Else
                    CommitAction
                End If
#End If
            End If
        End If
    Loop Until intAnchor = 0

    MakePathAux = True
Done:
    ChDir strOldPath

    Err = 0
End Function


'-----------------------------------------------------------
' FUNCTION: MsgFunc
'
' Forces mouse pointer to default and calls VB's MsgBox
' function.  See also MsgError.
'
' IN: [strMsg] - message to display
'     [intFlags] - MsgBox function type flags
'     [strCaption] - caption to use for message box
' Returns: Result of MsgBox function
'-----------------------------------------------------------
'
Function MsgFunc(ByVal strMsg As String, ByVal intFlags As Integer, ByVal strCaption As String) As Integer
    Dim intOldPointer As Integer
  
    intOldPointer = Screen.MousePointer
    If gfNoUserInput Then
        MsgFunc = GetDefMsgBoxButton(intFlags)
        If gfSilent = True Then
            LogSilentMsg strMsg
        End If
        If gfSMS = True Then
            LogSMSMsg strMsg
            gfDontLogSMS = False
        End If
    Else
        Screen.MousePointer = gintMOUSE_DEFAULT
        MsgFunc = MsgBox(strMsg, intFlags, strCaption)
        Screen.MousePointer = intOldPointer
    End If
End Function

'-----------------------------------------------------------
' FUNCTION: MsgWarning
'
' Forces mouse pointer to default, calls VB's MsgBox
' function, and logs this error and (32-bit only)
' writes the message and the user's response to the
' logfile (32-bit only)
'
' IN: [strMsg] - message to display
'     [intFlags] - MsgBox function type flags
'     [strCaption] - caption to use for message box
'
' Returns: Result of MsgBox function
'-----------------------------------------------------------
'
Function MsgWarning(ByVal strMsg As String, ByVal intFlags As Integer, ByVal strCaption As String) As Integer
    MsgWarning = MsgError(strMsg, intFlags, strCaption, MSGERR_WARNING)
End Function
'-----------------------------------------------------------
' SUB: SetFormFont
'
' Walks through all controls on specified form and
' sets Font a font chosen according to the system locale
'
' IN: [frm] - Form whose control fonts need to be set.
'-----------------------------------------------------------
'
Public Sub SetFormFont(frm As Form)
    Dim ctl As Control
    Dim fntSize As Integer
    Dim fntName As String
    Dim fntCharset As Integer
    Dim oFont As StdFont
    
    ' some controls may fail, so we will do a resume next...
    '
    On Error Resume Next
    
    ' get the font name, size, and charset
    '
    GetFontInfo fntName, fntSize, fntCharset
    
    'Create a new font object
    Set oFont = New StdFont
    With oFont
        .name = fntName
        .size = fntSize
        .Charset = fntCharset
    End With
    ' Set the form's font
    Set frm.Font = oFont
    '
    ' loop through each control and try to set its font property
    ' this may fail, but our error handling is shut off
    '
    For Each ctl In frm.Controls
        Set ctl.Font = oFont
    Next
    '
    ' get out, reset error handling
    '
    Set ctl = Nothing
    On Error GoTo 0
    Exit Sub
       
End Sub

'-----------------------------------------------------------
' SUB:  GetFontInfo
'
' Gets the best font to use according the current system's
' locale.
'
' OUT:  [sFont] - name of font
'       [nFont] - size of font
'       [nCharset] - character set of font to use
'-----------------------------------------------------------
Private Sub GetFontInfo(sFont As String, nFont As Integer, nCharSet As Integer)
    Dim LCID    As Integer
    Dim PLangId As Integer
    Dim sLangId As Integer
    ' if font is set, used the cached values
    If m_sFont <> "" Then
        sFont = m_sFont
        nFont = m_nFont
        nCharSet = m_nCharset
        Exit Sub
    End If
    
    ' font hasn't been set yet, need to get it now...
    LCID = GetSystemDefaultLCID                 ' get current system LCID
    PLangId = PRIMARYLANGID(LCID)               ' get LCID's Primary language id
    sLangId = SUBLANGID(LCID)                   ' get LCID's Sub language id
    
    Select Case PLangId                         ' determine primary language id
    Case LANG_CHINESE
        If (sLangId = SUBLANG_CHINESE_TRADITIONAL) Then
            sFont = ChrW$(&H65B0) & ChrW$(&H7D30) & ChrW$(&H660E) & ChrW$(&H9AD4)   ' New Ming-Li
            nFont = 9
            nCharSet = CHARSET_CHINESEBIG5
        ElseIf (sLangId = SUBLANG_CHINESE_SIMPLIFIED) Then
            sFont = ChrW$(&H5B8B) & ChrW$(&H4F53)
            nFont = 9
            nCharSet = CHARSET_CHINESESIMPLIFIED
        End If
    Case LANG_JAPANESE
        sFont = ChrW$(&HFF2D) & ChrW$(&HFF33) & ChrW$(&H20) & ChrW$(&HFF30) & _
                ChrW$(&H30B4) & ChrW$(&H30B7) & ChrW$(&H30C3) & ChrW$(&H30AF)
        nFont = 9
        nCharSet = CHARSET_SHIFTJIS
    Case LANG_KOREAN
        If (sLangId = SUBLANG_KOREAN) Then
            sFont = ChrW$(&HAD74) & ChrW$(&HB9BC)
        ElseIf (sLangId = SUBLANG_KOREAN_JOHAB) Then
            sFont = ChrW$(&HAD74) & ChrW$(&HB9BC)
        End If
        nFont = 9
        nCharSet = CHARSET_HANGEUL
    Case Else
        sFont = "Tahoma"
        If Not IsFontSupported(sFont) Then
            'Tahoma is not on this machine.  This condition is very probably since
            'this is a setup program that may be run on a clean machine
            'Try Arial
            sFont = "Arial"
            If Not IsFontSupported(sFont) Then
                'Arial isn't even on the machine.  This is an unusual situation that
                'is caused by deliberate removal
                'Try system
                sFont = "System"
                'If system isn't supported, allow the default font to be used
                If Not IsFontSupported(sFont) Then
                    'If "System" is not supported, "IsFontSupported" will have
                    'output the default font in sFont
                End If
            End If
        End If
        nFont = 8
        ' set the charset for the users default system Locale
        nCharSet = GetUserCharset
    End Select
    m_sFont = sFont
    m_nFont = nFont
    m_nCharset = nCharSet
'-------------------------------------------------------
End Sub
'-------------------------------------------------------
'-----------------------------------------------------------
' FUNCTION: ReadIniFile
'
' Reads a value from the specified section/key of the
' specified .INI file
'
' IN: [strIniFile] - name of .INI file to read
'     [strSection] - section where key is found
'     [strKey] - name of key to get the value of
'
' Returns: non-zero terminated value of .INI file key
'-----------------------------------------------------------
'
Function ReadIniFile(ByVal strIniFile As String, ByVal strsection As String, ByVal strKey As String) As String
    Dim strBuffer As String
    Dim intPos As Integer

    '
    'If successful read of .INI file, strip any trailing zero returned by the Windows API GetPrivateProfileString
    '
    strBuffer = Space$(gintMAX_SIZE)
    
    If GetPrivateProfileString(strsection, strKey, vbNullString, strBuffer, gintMAX_SIZE, strIniFile) > 0 Then
        ReadIniFile = RTrim$(StripTerminator(strBuffer))
    Else
        ReadIniFile = vbNullString
    End If
End Function

'Try to convert a path to its long filename equivalent, but leave it unaltered
'   if we fail.
'Public Sub MakeLongPath(Path As String)
'    On Error Resume Next
'    Path = LongPath(Path)
'End Sub

'-----------------------------------------------------------
' FUNCTION: MsgError
'
' Forces mouse pointer to default, calls VB's MsgBox
' function, and logs this error and (32-bit only)
' writes the message and the user's response to the
' logfile (32-bit only)
'
' IN: [strMsg] - message to display
'     [intFlags] - MsgBox function type flags
'     [strCaption] - caption to use for message box
'     [intLogType] (optional) - The type of logfile entry to make.
'                   By default, creates an error entry.  Use
'                   the MsgWarning() function to create a warning.
'                   Valid types as MSGERR_ERROR and MSGERR_WARNING
'
' Returns: Result of MsgBox function
'-----------------------------------------------------------
'
Function MsgError(ByVal strMsg As String, ByVal intFlags As Integer, ByVal strCaption As String, Optional ByVal intLogType As Integer = MSGERR_ERROR) As Integer
    Dim iRet As Integer
    
    iRet = MsgFunc(strMsg, intFlags, strCaption)
    MsgError = iRet
#If LOGGING Then
    ' We need to log this error and decode the user's response.
    Dim strID As String
    Dim strLogMsg As String

    Select Case iRet
        Case vbOK
            strID = ResolveResString(resLOG_vbok)
        Case vbCancel
            strID = ResolveResString(resLOG_vbCancel)
        Case vbAbort
            strID = ResolveResString(resLOG_vbabort)
        Case vbRetry
            strID = ResolveResString(resLOG_vbretry)
        Case vbIgnore
            strID = ResolveResString(resLOG_vbignore)
        Case vbYes
            strID = ResolveResString(resLOG_vbyes)
        Case vbNo
            strID = ResolveResString(resLOG_vbno)
        Case Else
            strID = ResolveResString(resLOG_IDUNKNOWN)
        'End Case
    End Select

    strLogMsg = strMsg & vbLf & "(" & ResolveResString(resLOG_USERRESPONDEDWITH, "|1", strID) & ")"
    On Error Resume Next
    Select Case intLogType
        Case MSGERR_WARNING
            LogWarning strLogMsg
        Case MSGERR_ERROR
            LogError strLogMsg
        Case Else
            LogError strLogMsg
        'End Case
    End Select
#End If
End Function

'-----------------------------------------------------------
' FUNCTION: GetFileVersion
'
' Returns the internal file version number for the specified
' file.  This can be different than the 'display' version
' number shown in the File Manager File Properties dialog.
' It is the same number as shown in the VB5 SetupWizard's
' File Details screen.  This is the number used by the
' Windows VerInstallFile API when comparing file versions.
'
' IN: [strFilename] - the file whose version # is desired
'     [fIsRemoteServerSupportFile] - whether or not this file is
'          a remote ActiveX component support file (.VBR)
'          (Enterprise edition only).  If missing, False is assumed.
'
' Returns: The Version number string if found, otherwise
'          vbnullstring
'-----------------------------------------------------------
'
'Function GetFileVersion(ByVal strFilename As String, Optional ByVal fIsRemoteServerSupportFile As Boolean = False) As String
'    Dim sVerInfo As VERINFO
'    Dim strVer As String

'    On Error GoTo GFVError

    '
    'Get the file version into a VERINFO struct, and then assemble a version string
    'from the appropriate elements.
    '
'    If GetFileVerStruct(strFilename, sVerInfo, fIsRemoteServerSupportFile) = True Then
'        strVer = Format$(sVerInfo.FileVerPart1) & gstrDECIMAL & Format$(sVerInfo.FileVerPart2) & gstrDECIMAL
'        strVer = strVer & Format$(sVerInfo.FileVerPart3) & gstrDECIMAL & Format$(sVerInfo.FileVerPart4)
'        GetFileVersion = strVer
'    Else
'        GetFileVersion = vbNullString
'    End If
'
'    Exit Function
'
'GFVError:
'    GetFileVersion = vbNullString
'    Err = 0
'End Function

'------------------------------------------------------------
'- Language Functions...
'------------------------------------------------------------
Private Function PRIMARYLANGID(ByVal LCID As Integer) As Integer
    PRIMARYLANGID = (LCID And &H3FF)
End Function
Private Function SUBLANGID(ByVal LCID As Integer) As Integer
    SUBLANGID = (LCID / (2 ^ 10))
End Function

'-----------------------------------------------------------
' Function: IsFontSupported
'
' Validates a font name to make sure it is supported by
' on the current system.
'
' IN/OUT: [sFontName] - name of font to check, will also]
'         be set to the default font name if the provided
'         one is not supported.
'-----------------------------------------------------------
Private Function IsFontSupported(sFontName As String) As Boolean
    Dim oFont As StdFont
    On Error Resume Next
    Set oFont = New StdFont
    oFont.name = sFontName
    IsFontSupported = (UCase(oFont.name) = UCase(sFontName))
    sFontName = oFont.name
End Function

'Public Function LongPath(Path As String) As String
'    Dim oDesktop As IVBShellFolder
'    Dim nEaten As Long
'    Dim pIdl As Long
'    Dim sPath As String
'    Dim oMalloc As IVBMalloc

'    If Len(Path) > 0 Then
'        SHGetDesktopFolder oDesktop
'        oDesktop.ParseDisplayName 0, 0, Path, nEaten, pIdl, 0
'        sPath = String$(gintMAX_PATH_LEN, 0)
'        SHGetPathFromIDListA pIdl, sPath
'        SHGetMalloc oMalloc
'        oMalloc.Free pIdl
'        LongPath = StringFromBuffer(sPath)
'    End If
'End Function

'-----------------------------------------------------------
' FUNCTION: GetFileVerStruct
'
' Gets the file version information into a VERINFO TYPE
' variable
'
' IN: [strFilename] - name of file to get version info for
'     [fIsRemoteServerSupportFile] - whether or not this file is
'          a remote ActiveX component support file (.VBR)
'          (Enterprise edition only).  If missing, False is assumed.
' OUT: [sVerInfo] - VERINFO Type to fill with version info
'
' Returns: True if version info found, False otherwise
'-----------------------------------------------------------
'
'Function GetFileVerStruct(ByVal sFile As String, sVer As VERINFO, Optional ByVal fIsRemoteServerSupportFile As Boolean = False) As Boolean
'    Dim lVerSize As Long, lTemp As Long, lRet As Long
'    Dim bInfo() As Byte
'    Dim lpBuffer As Long
'    Const sEXE As String * 1 = "\"
'    Dim fFoundVer As Boolean
'
'    GetFileVerStruct = False
'    fFoundVer = False
'
'    If fIsRemoteServerSupportFile Then
'        GetFileVerStruct = GetRemoteSupportFileVerStruct(sFile, sVer)
'        fFoundVer = True
'    Else
        '
        'Get the size of the file version info, allocate a buffer for it, and get the
        'version info.  Next, we query the Fixed file info portion, where the internal
        'file version used by the Windows VerInstallFile API is kept.  We then copy
        'the fixed file info into a VERINFO structure.
        '
'        lVerSize = GetFileVersionInfoSize(sFile, lTemp)
'        ReDim bInfo(lVerSize)
'        If lVerSize > 0 Then
'            lRet = GetFileVersionInfo(sFile, lTemp, lVerSize, VarPtr(bInfo(0)))
'            If lRet <> 0 Then
'                lRet = VerQueryValue(VarPtr(bInfo(0)), sEXE & vbNullChar, lpBuffer, lVerSize)
'                If lRet <> 0 Then
'                    CopyMemory sVer, ByVal lpBuffer, lVerSize
'                    fFoundVer = True
'                    GetFileVerStruct = True
'                End If
'            End If
'        End If
'    End If
'    If Not fFoundVer Then
        '
        ' We were unsuccessful in finding the version info from the file.
        ' One possibility is that this is a dependency file.
        '
'        If UCase(Extension(sFile)) = gstrEXT_DEP Then
'            GetFileVerStruct = GetDepFileVerStruct(sFile, sVer)
'        End If
'    End If
'End Function

