<div align="center">

## System\-wide mouse and keyboard hook


</div>

### Description

Set system-wide mouse and keyboard hook and generate 'standard' VB events (System_MouseUp/Down/Move, System_KeyUp/Down) with standard parameters (Button, Shift, X, Y, KeyCode).
 
### More Info
 
Form to receive hook notification and hook flags

Though MSDN says that WH_JOURNALRECORD hook is thread defined, in w95/98 it allow system-wide hook when set ThreadID parameter of hook = 0. To run this code you need form with two multiline textboxes (Text1 and Text2) and one label (Label1).

Mouse/keyboard events with appropriate parameters

Works only with w95/98. Don't work with NT/2000.

This code use hook, don't stop sample from IDE, use Form [x] button.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ark](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ark.md)
**Level**          |Advanced
**User Rating**    |4.8 (29 globes from 6 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ark-system-wide-mouse-and-keyboard-hook__1-9758/archive/master.zip)

### API Declarations

```
'See bas module code
```


### Source Code

```
'---Bas module code---
Option Explicit
Public Enum HookFlags
  HFMouseDown = 1
  HFMouseUp = 2
  HFMouseMove = 4
  HFKeyDown = 8
  HFKeyUp = 16
End Enum
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)
Private Declare Function GetAsyncKeyState% Lib "user32" (ByVal vKey As Long)
Private Declare Function GetForegroundWindow& Lib "user32" ()
Private Declare Function GetWindowThreadProcessId& Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long)
Private Declare Function GetKeyboardLayout& Lib "user32" (ByVal dwLayout As Long)
Private Declare Function MapVirtualKeyEx Lib "user32" Alias "MapVirtualKeyExA" (ByVal uCode As Long, ByVal uMapType As Long, ByVal dwhkl As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOREDRAW = &H8
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MOUSEWHEEL = &H20A
Private Const WH_JOURNALRECORD = 0
Type EVENTMSG
   wMsg As Long
   lParamLow As Long
   lParamHigh As Long
'   msgTime As Long
'   hWndMsg As Long
End Type
Dim EMSG As EVENTMSG
Dim hHook As Long, frmHooked As Form, hFlags As Long
Public Function HookProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 If nCode < 0 Then
   HookProc = CallNextHookEx(hHook, nCode, wParam, lParam)
   Exit Function
 End If
 Dim i%, j%, k%
 CopyMemory EMSG, ByVal lParam, Len(EMSG)
 Select Case EMSG.wMsg
  Case WM_KEYDOWN
    If (hFlags And HFKeyDown) = HFKeyDown Then
     If GetAsyncKeyState(vbKeyShift) Then j = 1
     If GetAsyncKeyState(vbKeyControl) Then j = 2
     If GetAsyncKeyState(vbKeyMenu) Then j = 4
     Select Case (EMSG.lParamLow And &HFF)
         Case 0 To 31, 90 To 159
           k = (EMSG.lParamLow And &HFF)
         Case Else
           k = MapVirtualKeyEx(EMSG.lParamLow And &HFF, 2, GetKeyboardLayout(GetWindowThreadProcessId(GetForegroundWindow, 0)))
     End Select
     frmHooked.System_KeyDown k, j
    End If
  Case WM_KEYUP
    If (hFlags And HFKeyUp) = HFKeyUp Then
     If GetAsyncKeyState(vbKeyShift) Then j = 1
     If GetAsyncKeyState(vbKeyControl) Then j = 2
     If GetAsyncKeyState(vbKeyMenu) Then j = 4
     Select Case (EMSG.lParamLow And &HFF)
         Case 0 To 31, 90 To 159
           k = (EMSG.lParamLow And &HFF)
         Case Else
           k = MapVirtualKeyEx(EMSG.lParamLow And &HFF, 2, GetKeyboardLayout(GetWindowThreadProcessId(GetForegroundWindow, 0)))
     End Select
     frmHooked.System_KeyUp k, j
    End If
  Case WM_MOUSEWHEEL
     Debug.Print "MouseWheel"
  Case WM_MOUSEMOVE
    If (hFlags And HFMouseMove) = HFMouseMove Then
     If GetAsyncKeyState(vbKeyLButton) Then i = 1
     If GetAsyncKeyState(vbKeyRButton) Then i = 2
     If GetAsyncKeyState(vbKeyMButton) Then i = 4
     If GetAsyncKeyState(vbKeyShift) Then j = 1
     If GetAsyncKeyState(vbKeyControl) Then j = 2
     If GetAsyncKeyState(vbKeyMenu) Then j = 4
     frmHooked.System_MouseMove i, j, CSng(EMSG.lParamLow), CSng(EMSG.lParamHigh)
    End If
  Case WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN
    If (hFlags And HFMouseDown) = HFMouseDown Then
     If GetAsyncKeyState(vbKeyShift) Then i = 1
     If GetAsyncKeyState(vbKeyControl) Then i = 2
     If GetAsyncKeyState(vbKeyMenu) Then i = 4
     frmHooked.System_MouseDown 2 ^ ((EMSG.wMsg - 513) / 3), i, CSng(EMSG.lParamLow), CSng(EMSG.lParamHigh)
    End If
  Case WM_LBUTTONUP, WM_RBUTTONUP, WM_MBUTTONUP
    If (hFlags And HFMouseUp) = HFMouseUp Then
     If GetAsyncKeyState(vbKeyShift) Then i = 1
     If GetAsyncKeyState(vbKeyControl) Then i = 2
     If GetAsyncKeyState(vbKeyMenu) Then i = 4
     frmHooked.System_MouseUp 2 ^ ((EMSG.wMsg - 514) / 3), i, CSng(EMSG.lParamLow), CSng(EMSG.lParamHigh)
    End If
 End Select
 Call CallNextHookEx(hHook, nCode, wParam, lParam)
End Function
Public Sub SetHook(fOwner As Form, flags As HookFlags)
  hHook = SetWindowsHookEx(WH_JOURNALRECORD, AddressOf HookProc, 0, 0)
  Set frmHooked = fOwner
  hFlags = flags
  Window_SetAlwaysOnTop frmHooked.hwnd, True
End Sub
Public Sub RemoveHook()
  UnhookWindowsHookEx hHook
  Window_SetAlwaysOnTop frmHooked.hwnd, False
  Set frmHooked = Nothing
End Sub
Private Function Window_SetAlwaysOnTop(hwnd As Long, bAlwaysOnTop As Boolean) As Boolean
  Window_SetAlwaysOnTop = SetWindowPos(hwnd, -2 - bAlwaysOnTop, 0, 0, 0, 0, SWP_NOREDRAW Or SWP_NOSIZE Or SWP_NOMOVE)
End Function
'---End of bas module code---
'--------------------------------------------
'---Form code---
'Add two multiline TextBoxes (better with vertical scrollbar) and one Label at form
Private Sub Form_Load()
  SetHook Me, HFMouseDown + HFMouseUp + HFMouseMove + HFKeyDown + HFKeyUp
  Text1 = "Mouse activity log:"
  Text2 = "Keyboard activity log:"
End Sub
Public Sub System_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim s As String
  Select Case KeyCode
     Case 32 To 90, 160 To 255
        s = LCase(Chr$(KeyCode))
     Case Else
        s = "ASCII code " & KeyCode
  End Select
  If Shift = vbShiftMask Then s = UCase(s): s = s & " + Shift "
  If Shift = vbCtrlMask Then s = s & " + Ctrl "
  If Shift = vbAltMask Then s = s & " + Alt "
  Text2 = Text2 & vbCrLf & s & " down"
End Sub
Public Sub System_KeyUp(KeyCode As Integer, Shift As Integer)
  Dim s As String
  Select Case KeyCode
     Case 32 To 90, 160 To 255
        s = LCase(Chr$(KeyCode))
     Case Else
        s = "ASCII code " & KeyCode
  End Select
  If Shift = vbShiftMask Then s = UCase(s): s = s & " + Shift "
  If Shift = vbCtrlMask Then s = s & " + Ctrl "
  If Shift = vbAltMask Then s = s & " + Alt "
  Text2 = Text2 & vbCrLf & s & " up"
End Sub
Public Sub System_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim s As String
 If Button = vbLeftButton Then s = "Left Button "
 If Button = vbRightButton Then s = "Right Button "
 If Button = vbMiddleButton Then s = "Middle Button "
 If Shift = vbShiftMask Then s = s & "+ Shift "
 If Shift = vbCtrlMask Then s = s & "+ Ctrl "
 If Shift = vbAltMask Then s = s & "+ Alt "
 Text1 = Text1 & vbCrLf & s & "Down at pos (pixels): " & CStr(x) & " , " & CStr(y)
End Sub
Public Sub System_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim s As String
 If Button = vbLeftButton Then s = "Left Button "
 If Button = vbRightButton Then s = "Right Button "
 If Button = vbMiddleButton Then s = "Middle Button "
 If Shift = vbShiftMask Then s = s & "+ Shift "
 If Shift = vbCtrlMask Then s = s & "+ Ctrl "
 If Shift = vbAltMask Then s = s & "+ Alt "
 Text1 = Text1 & vbCrLf & s & "Up at pos (pixels): " & CStr(x) & " , " & CStr(y)
End Sub
Public Sub System_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim s As String
 If Button = vbLeftButton Then s = "Left Button "
 If Button = vbRightButton Then s = "Right Button "
 If Button = vbMiddleButton Then s = "Middle Button "
 If Shift = vbShiftMask Then s = s & "+ Shift "
 If Shift = vbCtrlMask Then s = s & "+ Ctrl "
 If Shift = vbAltMask Then s = s & "+ Alt "
 Label1 = "Mouse info" & vbCrLf & "X = " & x & " Y= " & y & vbCrLf
 If s <> "" Then Label1 = Label1 & "Extra Info: " & vbCrLf & s & "pressed"
End Sub
Private Sub Form_Unload(Cancel As Integer)
  RemoveHook
End Sub
'--End of form code--
```

