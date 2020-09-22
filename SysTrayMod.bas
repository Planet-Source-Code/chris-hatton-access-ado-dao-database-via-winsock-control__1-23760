Attribute VB_Name = "SysTrayMod"
Option Explicit
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    sTip As String * 64
    End Type
    Public Const NIM_ADD = &H0
    Public Const NIM_MODIFY = &H1
    Public Const NIM_DELETE = &H2
    Public Const NIF_MESSAGE = &H1
    Public Const NIF_ICON = &H2
    Public Const NIF_TIP = &H4
    Public Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    Public Const SW_RESTORE = 9
    Public Const SW_MINIMIZE = 6
    Public Const WM_MOUSEMOVE = &H200
    Public Const WM_LBUTTONDBLCLK = &H203
    Public Const WM_RBUTTONUP = &H205
