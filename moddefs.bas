Attribute VB_Name = "moddefs"
'Definitions module

'Predefined control classes that you can create with CreateWindowEx (in lpClassName)
'BUTTON (CommandButton)
'LISTBOX
'COMBOBOX
'EDIT (TextBox)
'MDICLIENT (MDI client window)
'RichEdit - RichEdit v1.0
'RICHEDIT_CLASS - RichEdit v2.0
'SCROLLBAR
'STATIC (label)

'CreateWindowEx parameters:
'dwExStyle - extended window style, Long
'lpClassName - pointer to registered class name, can be String
'lpWindowName - pointer to window name, can be String
'dwStyle - window style, Long
'x - horizontal position of window, Long
'y - vertical position of window, Long
'nWidth - window width, Long
'nHeight - window height, Long
'hWndParent - handle to parent or owner window, Long
'hMenu - handle to menu, or child-window identifier
'hInstance - handle to application instance
'lpParam - pointer to window-creation data
 
Public Type POINTAPI
    x As Long
    y As Long
End Type
Public Type Msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type
Public Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type
Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(32) As Byte 'or LF_FACESIZE instead of 32
End Type

Public Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Public Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Public Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Public Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Public Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const GWL_WNDPROC = (-4)
'generic window (extended) styles:
Public Const WS_VISIBLE = &H10000000
Public Const WS_CHILD = &H40000000
Public Const WS_THICKFRAME = &H40000
Public Const WS_TABSTOP = &H10000
Public Const WS_BORDER = &H800000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_CAPTION = &HC00000 ' WS_BORDER Or WS_DLGFRAME
Public Const WS_SYSMENU = &H80000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
'uncomment the flags in the following constant to have Maximize and Minimize buttons
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME) 'Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_EX_CLIENTEDGE = &H200&
Public Const WS_EX_WINDOWEDGE = &H100&
Public Const WS_EX_STATICEDGE = &H20000
'STATIC (Label) style constant:
Public Const SS_CENTER = &H1&
'BUTTON style:
Public Const BS_LEFT = &H100& 'left-aligned text
Public Const BS_TOP = &H400& 'top-aligned text
Public Const BS_CHECKBOX = &H2&
Public Const BS_AUTOCHECKBOX = &H3&
Public Const BM_GETCHECK = &HF0
Public Const BM_CLICK = &HF5
'constants used in acting upon the ListBox:
Public Const LB_ADDSTRING = &H180
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETTEXT = &H189
Public Const LB_GETCURSEL = &H188
Public Const LB_GETTOPINDEX = &H18E
'constants used in creating the ListBox:
Public Const LBS_NOTIFY = &H1&
Public Const LBS_SORT = &H2&
Public Const WS_VSCROLL = &H200000
Public Const LBS_STANDARD = (LBS_NOTIFY Or WS_VSCROLL Or WS_BORDER)
'^of course that you can specify here the LBS_SORT flag, also
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_VSCROLL = &H115 'vertical scroll
Public Const WM_KEYUP = &H101 'emulate the end of _KeyPress or _Change
Public Const WM_LBUTTONUP = &H202 'emulate the end of _Click
Public Const WM_LBUTTONDOWN = &H201 'emulate the beginning of _Click
Public Const WM_SHOWWINDOW = &H18
Public Const WM_DESTROY = &H2 'aka Form_Unload
Public Const WM_SETFONT = &H30 'used in building text font for the new controls
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302

Public Const CS_VREDRAW = &H1
Public Const CS_HREDRAW = &H2
Public Const COLOR_WINDOW = 5
Public Const IDC_ARROW = 32512&
Public Const IDI_APPLICATION = 32512&
Public Const SW_SHOWNORMAL = 1
Public Const CW_USEDEFAULT = &H80000000
Public Const gClassName = "CustomClName"
Public Const gAppName = "Application caption"
Public gHwnd As Long
