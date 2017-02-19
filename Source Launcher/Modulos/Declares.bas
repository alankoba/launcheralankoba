Attribute VB_Name = "Declares"

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1
Public Declare Function SHCreateThread Lib "shlwapi.dll" (ByVal pfnThreadProc As Long, pData As Any, ByVal dwFlags As Long, ByVal pfnCallback As Long) As Long

Public Function MoverFormSemCaption(Botao As Integer, FormHwnd As Long) As Long
    If Botao = 1 Then
        Call ReleaseCapture
        MoverFormSemCaption = SendMessage(FormHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
End Function
