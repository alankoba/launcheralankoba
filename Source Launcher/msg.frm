VERSION 5.00
Begin VB.Form msg 
   BorderStyle     =   0  'None
   ClientHeight    =   2670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   ScaleHeight     =   2670
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "msg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Const SW_SHOWNORMAL = 1
Const MB_OK = &H0&
Const WM_CLOSE = &H10
Private Sub Form_Load()
msg.Hide
    Dim WinWnd As Long, Ret As String
    'Ask for a Window title
    Ret = "MU"
    'Search the window
    WinWnd = FindWindow(vbNullString, Ret)
    PostMessage WinWnd, WM_CLOSE, 0&, 0&
KillProcess (exe_nome)
MsgBox "Programa suspeito encontrado no sistema" + vbNewLine + _
"Nome da ameaça: " + det + vbNewLine + _
"Este alerta pode ser um falso-positivo. caso tenha duvidas entre em contato com nosso suporte", vbCritical, "Lite Guard Add-on | " + servername
End
End Sub
