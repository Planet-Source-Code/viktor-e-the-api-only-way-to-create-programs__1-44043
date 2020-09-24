Attribute VB_Name = "modsubcl"
'Here we intercept the messages sent to the subclassed windows

'Now, it's obvious that for a form which must have a bunch of buttons, textboxes, lb's etc.,
'the creation of a subclassing procedure for each control is at least unpractical
'Right, who would do this instead of classical control manipulation ? Maybe those who want to get rid
'of some frx's... I don't know, take this code as is, for teaching purposes only... :)

'Imagine doing this in VC++...

Public ProcOld As Long 'used in subclassing the button
Public ProcOld2 As Long 'used in subclassing the textbox
Public ProcOld3 As Long 'used in subclassing the listbox
Public ProcOld4 As Long 'used in subclassing the checkbox

Public b As Long 'CommandButton created on the fly
Public lb As Long 'ListBox created on the fly
Public tb As Long 'TextBox created on the fly
Public lab As Long 'Label created on the fly
Dim lc As Long 'ListCount
Dim nchars As Long 'characters in the tb TextBox
Dim s As String 'used to retrieve ListBox selected item text
Dim sItem As Long 'the index of the selected ListBox item

Public Function WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'dedicated to messages sent to "b" CommandButton
Select Case iMsg
Case WM_LBUTTONUP
    'get the number of items in the "lb" ListBox...
    lc = SendMessage(lb, LB_GETCOUNT, 0, 0)
    '... and display it in the "tb" TextBox
    SetWindowText tb, ByVal CStr(lc & " items")
Case Else
'Ignore all other messages
End Select
WindowProc = CallWindowProc(ProcOld, hwnd, iMsg, wParam, lParam)
End Function
Public Function WindowProc2(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'dedicated to messages sent to "tb" TextBox
Select Case iMsg
Case WM_KEYUP
    'retrieve the length of "tb" text...
    nchars = SendMessage(tb, WM_GETTEXTLENGTH, 0, 0)
    '... and display it in the "lab" Label
    SetWindowText lab, ByVal CStr(nchars & " characters")
Case WM_COPY
    SetWindowText lab, ByVal CStr("Text copied to Clipboard")
Case WM_PASTE
    WindowProc2 = CallWindowProc(ProcOld2, hwnd, iMsg, wParam, lParam)
    SendMessage tb, WM_KEYUP, 0, 0
    Exit Function
Case Else
'Ignore all other messages
End Select
WindowProc2 = CallWindowProc(ProcOld2, hwnd, iMsg, wParam, lParam)
End Function
Public Function WindowProc3(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'dedicated to messages sent to "lb" ListBox
Select Case iMsg
Case WM_LBUTTONUP
    'get the "lb" ListIndex...
    sItem = SendMessage(lb, LB_GETCURSEL, 0, 0)
    '...and display the item text in the "lab" Label
    s = String(10, Chr$(0))
    SendMessage lb, LB_GETTEXT, sItem, ByVal s 'mandatory ByVal, otherwise crash!boom!bang!
    s = Left$(s, InStr(s, Chr$(0)) - 1)
    SetWindowText lab, "Selected: " & s
Case WM_VSCROLL
    s = String(10, Chr$(0))
    'get the first visible item text...
    sItem = SendMessage(lb, LB_GETTOPINDEX, 0, 0)
    SendMessage lb, LB_GETTEXT, sItem, ByVal s 'mandatory ByVal, otherwise crash!boom!bang!
    s = Left$(s, InStr(s, Chr$(0)) - 1)
    '...and display it in the "lab" Label
    SetWindowText lab, ByVal CStr("First visible item index: " & SendMessage(lb, LB_GETTOPINDEX, 0, 0) & " (" & s & ")")
Case Else
'Ignore all other messages
End Select
WindowProc3 = CallWindowProc(ProcOld3, hwnd, iMsg, wParam, lParam)
End Function
Public Function WindowProc4(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'dedicated to messages sent to "ch" CheckBox
Select Case iMsg
Case WM_LBUTTONUP
    Dim cState As Long
    cState = SendMessage(hwnd, BM_GETCHECK, 0, 0)
    If cState = 0 Then
        SetWindowText lab, "Checked"
        DestroyWindow lb
        lb = CreateWindowEx(WS_EX_CLIENTEDGE, "LISTBOX", "", WS_CHILD Or LBS_STANDARD Or WS_VISIBLE Or LBS_SORT, 20, 60, 120, 120, gHwnd, 0, App.hInstance, 0)
        ReSubclassListBox lb
    Else
        SetWindowText lab, "Unchecked"
        DestroyWindow lb
        lb = CreateWindowEx(WS_EX_CLIENTEDGE, "LISTBOX", "", WS_CHILD Or LBS_STANDARD Or WS_VISIBLE, 20, 60, 120, 120, gHwnd, 0, App.hInstance, 0)
        ReSubclassListBox lb
    End If
Case Else
'Ignore all other messages
End Select
WindowProc4 = CallWindowProc(ProcOld4, hwnd, iMsg, wParam, lParam)
End Function

Public Function FormProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'dedicated to messages sent to the new Form, of [class name]
Select Case iMsg
Case WM_SHOWWINDOW
    'set some... title on the form:
    SendMessage hwnd, WM_SETTEXT, 0, ByVal CStr("Lightweight API-based proggie") 'mandatory ByVal CStr for readable text
    'didn't set ScaleMode to pixels - this function makes the transformation from twips to pixels by itself
    b = CreateWindowEx(WS_EX_WINDOWEDGE, "BUTTON", "Click for ListCount", BS_LEFT Or BS_TOP Or WS_CHILD Or WS_VISIBLE, 20, 20, 120, 30, gHwnd, 0, App.hInstance, 0)
    lb = CreateWindowEx(WS_EX_CLIENTEDGE, "LISTBOX", "", WS_CHILD Or LBS_STANDARD Or WS_VISIBLE, 20, 60, 120, 120, gHwnd, 0, App.hInstance, 0)
    tb = CreateWindowEx(WS_EX_CLIENTEDGE, "EDIT", "Type here to get Len(.Text)", WS_CHILD Or WS_VISIBLE, 150, 20, 150, 20, gHwnd, 0, App.hInstance, 0)
    '^will be subclassed
    'an info label:
    lab = CreateWindowEx(WS_EX_STATICEDGE, "STATIC", "Plain Arial...", WS_CHILD Or WS_VISIBLE, 20, 180, 280, 20, gHwnd, 0, App.hInstance, 0)
    'a cool-looking... reminder label:
    lab2 = CreateWindowEx(WS_EX_WINDOWEDGE, "STATIC", "End the program only by clicking the Close button on the form", WS_CHILD Or WS_VISIBLE Or SS_CENTER Or WS_THICKFRAME, 150, 60, 150, 80, gHwnd, 0, App.hInstance, 0)
    '^you can play with CLIENT-, STATIC- and WINDOW EDGEs for cool windows...
    'a checkbox to (un)sort the ListBox:
    ch = CreateWindowEx(WS_EX_WINDOWEDGE, "BUTTON", "Sorted list", WS_CHILD Or WS_VISIBLE Or BS_AUTOCHECKBOX, 160, 150, 120, 20, gHwnd, 0, App.hInstance, 0)
    '^if the WS_VISIBLE flag isn't specified in the new window style,
    'you must call the ShowWindow API function to make the windows visible, such as:
    'ShowWindow blahblah, SW_NORMAL
    'add 10 items to the new ListBox:
    AddItemsToListBox
    'set a font for each of the created controls:
    SetFontFor b: SetFontFor lb: SetFontFor tb: SetFontFor lab2: SetFontFor ch
    'leave lab as it is, to see what a raw control text looks like, with a system font
    'start intercepting messages sent to the new button:
    ProcOld = SetWindowLong(b, GWL_WNDPROC, AddressOf WindowProc)
    'start intercepting messages sent to the new TextBox:
    ProcOld2 = SetWindowLong(tb, GWL_WNDPROC, AddressOf WindowProc2)
    'start intercepting messages sent to the new ListBox:
    ProcOld3 = SetWindowLong(lb, GWL_WNDPROC, AddressOf WindowProc3)
    'start intercepting messages sent to the new CheckBox:
    ProcOld4 = SetWindowLong(ch, GWL_WNDPROC, AddressOf WindowProc4)
    'subclassing for each window must be done only once and must be associated with an unsubclassing procedure
Case WM_DESTROY
    'send the WM_QUIT message to the form
    PostQuitMessage 0 '&
    'unregister its class:
    UnregisterClass gClassName, App.hInstance
Case Else
'Ignore all other messages
End Select
FormProc = DefWindowProc(hwnd, iMsg, wParam, lParam)
End Function
