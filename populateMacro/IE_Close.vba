Attribute VB_Name = "IE_Close"
'Option Explicit
  
Private Declare Function PostMessage Lib "user32" _
Alias "PostMessageA" _
(ByVal hwnd As Long, _
ByVal wMsg As Long, _
ByVal wParam As Long, _
lParam As Any) As Long
  
Private Declare Function GetDesktopWindow Lib "user32" () As Long
  
Private Declare Function GetWindow Lib "user32" _
(ByVal hwnd As Long, _
ByVal wCmd As Long) As Long
  
Private Declare Function GetWindowText Lib "user32" _
Alias "GetWindowTextA" _
(ByVal hwnd As Long, _
ByVal lpString As String, _
ByVal cch As Long) As Long
  
Private Declare Function GetClassName Lib "user32" _
Alias "GetClassNameA" _
(ByVal hwnd As Long, _
ByVal lpClassName As String, _
ByVal nMaxCount As Long) _
As Long
  
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3
Private Const GW_OWNER = 4
Private Const GW_CHILD = 5
Private Const WM_CLOSE = &H10
  
Function FindWindowHwndLike(hWndStart As Long, _
ClassName As String, WindowTitle As String, _
level As Long, lHolder As Long) As Long
  
'finds the first window where the class name start with ClassName
'and where the Window title starts with WindowTitle, returns Hwnd
'----------------------------------------------------------------
  
Dim hwnd As Long
Dim sWindowTitle As String
Dim sClassName As String
Dim r As Long
  
'Initialize if necessary. This is only executed
'when level = 0 and hWndStart = 0, normally
'only on the first call to the routine.
  
If level = 0 Then
    If hWndStart = 0 Then
        hWndStart = GetDesktopWindow()
    End If
End If
  
'Increase recursion counter
level = level + 1

'Get first child window
hwnd = GetWindow(hWndStart, GW_CHILD)
  
Do Until hwnd = 0
  
    'Search children by recursion
    lHolder = FindWindowHwndLike(hwnd, ClassName, _
    WindowTitle, level, lHolder)
  
    'Get the window text
    sWindowTitle = Space$(255)
    r = GetWindowText(hwnd, sWindowTitle, 255)
    sWindowTitle = Left$(sWindowTitle, r)
  
    'get the class name
    sClassName = Space$(255)
    r = GetClassName(hwnd, sClassName, 255)
    sClassName = Left$(sClassName, r)
    
    If InStr(1, sWindowTitle, WindowTitle, vbBinaryCompare) > 0 And _
            sClassName Like ClassName & "*" Then
        FindWindowHwndLike = hwnd
        lHolder = hwnd
        Exit Function
    End If
  
    'Get next child window
    hwnd = GetWindow(hwnd, GW_HWNDNEXT)
Loop

FindWindowHwndLike = lHolder
  
End Function
  
Function CloseApp(ByVal strApp As String, _
ByVal strClass As String) As Long
  
'will find a window based on:
'the partial start of the Window title and/or
'the partial start of the Window class
'and then close that window
  
'for example, this will close Excel:
'CloseApp "", "XLM" and this will:
'CloseApp "Microsoft Excel", ""
'but this won't: CloseApp "", "LM"
'it will only close the first window that
'fulfills the criteria
'will return Hwnd if successfull, and 0 if not
'---------------------------------------------
  
Dim hwnd As Long
  
On Error GoTo ERROROUT
  
hwnd = FindWindowHwndLike(0, strClass, strApp, 0, 0)
  
If hwnd = 0 Then
    CloseApp = 0
    Exit Function
End If
  
'Post a message to the window to close itself
'--------------------------------------------
PostMessage hwnd, WM_CLOSE, 0&, 0&
CloseApp = hwnd
Exit Function
  
ERROROUT:
On Error GoTo 0
CloseApp = 0
End Function

