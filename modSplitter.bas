Attribute VB_Name = "modSplitter"
' Module      : modSplitter
' Description : Module to support splitter operations
' Designe  by Herman Pouls
' Data 15-11-2000

Option Explicit

'*
'* Declare function s found in gdi32.dll.
'*
Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As Long, ByVal lpOutput As Long, ByVal lpInitData As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long


'*
'* Basic types.
'*
Public Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type


Type POINTAPI
    X As Long
    Y As Long
    
End Type


'*
'* ShowWindow() Commands.
'*
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
Public Const SW_FORCEMINIMIZE = 11
Public Const SW_MAX = 11


'*
'* Old ShowWindow() Commands.
'*
Public Const HIDE_WINDOW = 0
Public Const SHOW_OPENWINDOW = 1
Public Const SHOW_ICONWINDOW = 2
Public Const SHOW_FULLSCREEN = 3
Public Const SHOW_OPENNOACTIVATE = 4


'*
'* Identifiers for the WM_SHOWWINDOW message.
'*
Public Const SW_PARENTCLOSING = 1
Public Const SW_OTHERZOOM = 2
Public Const SW_PARENTOPENING = 3
Public Const SW_OTHERUNZOOM = 4


'*
'* Window Styles.
'*
Public Const WS_OVERLAPPED = &H0&
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_TABSTOP = &H10000
Public Const WS_GROUP = &H20000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_THICKFRAME = &H40000
Public Const WS_SYSMENU = &H80000
Public Const WS_HSCROLL = &H100000
Public Const WS_VSCROLL = &H200000
Public Const WS_DLGFRAME = &H400000
Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_DISABLED = &H8000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_CHILD = &H40000000
Public Const WS_POPUP = &H80000000


'*
'* Common Window Styles.
'*
Public Const WS_CHILDWINDOW = (WS_CHILD)
Public Const WS_ICONIC = WS_MINIMIZE
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_TILED = WS_OVERLAPPED
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW


'*
'* Extended Window Styles.
'*
Public Const WS_EX_LEFT = &H0&
Public Const WS_EX_LTRREADING = &H0&
Public Const WS_EX_RIGHTSCROLLBAR = &H0&
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_TRANSPARENT = &H20&
Public Const WS_EX_MDICHILD = &H40&
Public Const WS_EX_TOOLWINDOW = &H80&
Public Const WS_EX_WINDOWEDGE = &H100&
Public Const WS_EX_CLIENTEDGE = &H200&
Public Const WS_EX_CONTEXTHELP = &H400&
Public Const WS_EX_RIGHT = &H1000&
Public Const WS_EX_RTLREADING = &H2000&
Public Const WS_EX_LEFTSCROLLBAR = &H4000&
Public Const WS_EX_CONTROLPARENT = &H10000
Public Const WS_EX_STATICEDGE = &H20000
Public Const WS_EX_APPWINDOW = &H40000

Public Const WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
Public Const WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)


'*
'* Window field offsets for GetWindowLong().
'*
Public Const GWL_WNDPROC = (-4)
Public Const GWL_HINSTANCE = (-6)
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_ID = (-12)
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_USERDATA = (-21)


'*
'* Window messages.
'*
Public Const WM_SETFONT = &H30
Public Const WM_USER = &H400


'*
'* SetWindowPos() Flags.
'*
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FRAMECHANGED = &H20      ' The frame changed: send WM_NCCALCSIZE.
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOOWNERZORDER = &H200    ' Don't do owner Z ordering.
Public Const SWP_NOSENDCHANGING = &H400   ' Don't send WM_WINDOWPOSCHANGING.
Public Const SWP_DEFERERASE = &H2000
Public Const SWP_ASYNCWINDOWPOS = &H4000

Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

Public Const HWND_TOP = (0)
Public Const HWND_BOTTOM = (1)
Public Const HWND_TOPMOST = (-1)
Public Const HWND_NOTOPMOST = (-2)


'*
'* Declare functions found in user32.dll.
'*
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpszClassName As String, ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long


Public Function ClientRectToScreen(ByVal hWnd As Long, lpRect As RECT) As Boolean
    Dim bSuccess As Boolean
    Dim ptTemp   As POINTAPI
    
    
    ' Convert the window coordinates to screen coordinates.
    ptTemp.X = lpRect.Left
    ptTemp.Y = lpRect.Top
    
    bSuccess = ClientToScreen(hWnd, ptTemp)
    
    lpRect.Left = ptTemp.X
    lpRect.Top = ptTemp.Y
    
    ptTemp.X = lpRect.Right
    ptTemp.Y = lpRect.Bottom
    
    bSuccess = (bSuccess And ClientToScreen(hWnd, ptTemp))
    
    lpRect.Right = ptTemp.X
    lpRect.Bottom = ptTemp.Y
    
     
End Function


Public Function DrawSplitterRect(ByVal hdc As Long, lpRect As RECT) As Boolean
    Dim bSuccess As Boolean
    Dim rcNew    As RECT
    
    
    ' Create a copy of the RECT structure.
    rcNew = lpRect
    
    
    ' Draw focus style rectangle.
    bSuccess = DrawFocusRect(hdc, rcNew)
    
    
    ' Resize rectangle (minus one pixel.
    rcNew.Bottom = rcNew.Bottom - 1
    rcNew.Left = rcNew.Left + 1
    rcNew.Right = rcNew.Right - 1
    rcNew.Top = rcNew.Top + 1
    

    ' Draw focus style redctangle.
    bSuccess = (bSuccess And DrawFocusRect(hdc, rcNew))
    
    
    ' Return success code.
    DrawSplitterRect = bSuccess
    
    
End Function


