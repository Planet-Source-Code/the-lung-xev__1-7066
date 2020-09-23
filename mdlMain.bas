Attribute VB_Name = "mdlMain"
Option Explicit
    
    Public apiError As Long
    Public Namespace As SWbemServices


Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long


    Public NOTIFYICONDATA As NOTIFYICONDATA
    Public Type NOTIFYICONDATA
            cbSize As Long
            hwnd As Long
            uID As Long
            uFlags As Long
            uCallbackMessage As Long
            hIcon As Long
            szTip As String * 64
    End Type


    Public Const NIF_ICON = &H2
    Public Const NIF_MESSAGE = &H1
    Public Const NIF_TIP = &H4
    Public Const NIF_STATE = &H8
    Public Const NIF_INFO = &H10

    Public Const NIM_ADD = &H0
    Public Const NIM_DELETE = &H2
    Public Const NIM_MODIFY = &H1
    Public Const NIM_SETFOCUS = &H3
    Public Const NIM_SETVERSION = &H4

    Public Const WM_CLOSE = &H10
    Public Const WM_DESTROY = &H2
    Public Const WM_LBUTTONDBLCLK = &H203
    Public Const WM_LBUTTONDOWN = &H201
    Public Const WM_LBUTTONUP = &H202
    Public Const WM_MBUTTONDBLCLK = &H209
    Public Const WM_MBUTTONDOWN = &H207
    Public Const WM_MBUTTONUP = &H208
    Public Const WM_MDIDESTROY = &H221
    Public Const WM_NCDESTROY = &H82
    Public Const WM_RBUTTONDBLCLK = &H206
    Public Const WM_RBUTTONDOWN = &H204
    Public Const WM_RBUTTONUP = &H205
    Public Const WM_SETTEXT = &HC
