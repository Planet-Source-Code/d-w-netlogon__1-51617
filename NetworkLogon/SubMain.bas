Attribute VB_Name = "SubMain"
Option Explicit

Private Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function EnumChildWindows Lib "user32.dll" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Any) As Long 'ANY to send STRING

Private Const WM_SETTEXT = &HC

Public Function EnumChildProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
'this will only fire if target window is found
Static WinNum As Integer
WinNum = WinNum + 1

Debug.Print "# " & WinNum & " hWnd = " & hWnd

If WinNum = 6 Then
SendMessage hWnd, WM_SETTEXT, 0, ByVal "NetworkLogon" 'put username here
End If

If WinNum = 8 Then
SendMessage hWnd, WM_SETTEXT, 0, ByVal "NetworkPassword" 'put password here
EnumChildProc = 0
WinNum = 0
Exit Function
End If

EnumChildProc = 1
End Function

Public Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
Dim Slength As Long
Dim Wintext As String
Dim Retval As Long
Dim Buffer As String

Slength = GetWindowTextLength(hWnd) + 1
If Slength > 1 Then
Buffer = Space(Slength)
GetWindowText hWnd, Buffer, Slength
    
    If Left(Buffer, 11) = "Connect to " Then 'this is what my network logon caption says
    Debug.Print "This is the one >>" & hWnd & "<<"
    EnumWindowsProc = 0
    EnumChildWindows hWnd, AddressOf EnumChildProc, 0
    Exit Function
    End If

Wintext = Left(Buffer, Slength - 1)
Debug.Print "Text of window hWnd " & hWnd & " is " & Wintext
Else
Debug.Print "Window hWnd " & hWnd & " has no text"
End If
EnumWindowsProc = 1
End Function

Sub Main()
EnumWindows AddressOf EnumWindowsProc, 0
End Sub


