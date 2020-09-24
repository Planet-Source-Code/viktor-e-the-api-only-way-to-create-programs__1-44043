Attribute VB_Name = "modfuncs"
'Functions module

Dim rfont As Long, tba() As Byte, balimit As Long, x As Long, FSize As Long, f As Long

Sub Main()
'The form is created and subclassed using code by Joseph Huntley;
'see "API Form by Joseph Huntley" here on PSC, contest winner.
'All other code by me
Dim wMsg As Msg
'Register window class name
If RegisterWindowClass = False Then Exit Sub
'Create window
If CreateWindows Then
'Loop will exit when WM_QUIT is sent to the window.
Do While GetMessage(wMsg, 0, 0, 0)
''TranslateMessage takes keyboard messages and converts them to WM_CHAR for easier processing
TranslateMessage wMsg
''Dispatchmessage calls the default window procedure
''to process the window message. (WndProc)
DispatchMessage wMsg
Loop
End If
Call UnregisterClass(gClassName, App.hInstance)
End Sub

Public Function RegisterWindowClass() As Boolean
Dim wc As WNDCLASS
'Registers the new window so we can use its class name:
'redraw entire window when movement or size adjustments modify the client area's width/height:
With wc
    .style = CS_HREDRAW Or CS_VREDRAW 'Specifies the class style(s). Styles can be combined by using the bitwise OR operator
    .lpfnwndproc = GetAddress(AddressOf FormProc) 'Address of (pointer to) the window procedure
    .hInstance = App.hInstance 'Handle to the instance that the window procedure of this class is within
    .hIcon = LoadIcon(0&, IDI_APPLICATION) 'Default application icon
    .hCursor = LoadCursor(0&, IDC_ARROW) 'Default arrow cursor
    .hbrBackground = COLOR_WINDOW 'Default color for window background
    .lpszClassName = gClassName 'Pointer to a null-terminated string or, if lpszClassName is a string, it specifies the window class name
End With
RegisterWindowClass = RegisterClass(wc) <> 0
End Function
Public Function CreateWindows() As Boolean
'Create actual window
gHwnd = CreateWindowEx(0, gClassName, gAppName, WS_OVERLAPPEDWINDOW, 100, 100, 340, 240, 0, 0, App.hInstance, ByVal 0&)
ShowWindow gHwnd, SW_SHOWNORMAL
CreateWindows = (gHwnd <> 0)
End Function
Public Function GetAddress(ByVal lngAddr As Long) As Long
'Used with AddressOf to return the address in memory of a procedure
GetAddress = lngAddr '&
End Function

Public Sub SetFontFor(ByVal hwnd As Long)
Dim LogicFont As LOGFONT
FSize = 8
With LogicFont
.lfHeight = (FSize * -20) / Screen.TwipsPerPixelY
'.lfEscapement = 0  'orientation; i.e. 90 degrees=900
'.lfItalic = 0 '>=1: italic, 0: not italic
End With
'Set font name:
tba = StrConv("Tahoma" & Chr$(0), vbFromUnicode)
'Build the font structure:
balimit = UBound(tba)
For x = 0 To balimit
    LogicFont.lfFaceName(x) = tba(x)
Next x
rfont = CreateFontIndirect(LogicFont)
'Apply the newly created font to the control text:
SendMessage hwnd, WM_SETFONT, rfont, 0
End Sub
Public Sub ReSubclassListBox(ByVal hwnd As Long)
'WARNING: FOR lb ONLY
AddItemsToListBox
SetWindowLong lb, GWL_WNDPROC, ProcOld3
'start intercepting messages sent to the new ListBox:
ProcOld3 = SetWindowLong(lb, GWL_WNDPROC, AddressOf WindowProc3)
End Sub
Public Sub AddItemsToListBox()
For f = 1 To 10
    SendMessage lb, LB_ADDSTRING, 0, ByVal CStr("Item #" & f)
Next f
SetFontFor lb
End Sub
