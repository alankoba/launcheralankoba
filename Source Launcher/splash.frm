VERSION 5.00
Begin VB.Form splash 
   BorderStyle     =   0  'None
   Caption         =   "proc2"
   ClientHeight    =   2775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3330
   LinkTopic       =   "Form1"
   Picture         =   "splash.frx":0000
   ScaleHeight     =   2775
   ScaleWidth      =   3330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   240
      Top             =   600
   End
   Begin VB.Timer verifica 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2760
      Top             =   1320
   End
   Begin VB.Timer ocultame 
      Interval        =   1800
      Left            =   2760
      Top             =   1800
   End
   Begin VB.Timer buscando 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   2760
      Top             =   2280
   End
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProcess As Long, ByRef Wow64Process As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Const SW_SHOWNORMAL = 1
Const WM_CLOSE = &H10

Private Sub buscando_Timer()
If ProcessoExiste(exe_nome) Then
Else
End
End If
Call busca
Call buscac
Call protectexe
Call GetWinInfo
End Sub

Private Sub Form_Load()
App.TaskVisible = False
If hide_protect = 1 Then
    Dim Retas As Long
    IsWow64Process GetCurrentProcess, Retas
    If Retas = 0 Then
        Timer1.Enabled = True
    Else
      Timer1.Enabled = False
    End If
End If
End Sub

Private Sub ocultame_Timer()
Me.Hide
 ShellExecute Me.hWnd, vbNullString, exe_nome, "connect /u" + ip + " /p" + porta, vbNullString, SW_SHOWNORMAL
ocultame.Enabled = False
verifica.Enabled = True
buscando.Enabled = True
End Sub

Private Sub Timer1_Timer()
Call hides
End Sub

Private Sub verifica_Timer()
If ProcessoExiste("Verificador.exe") Then
Else
det = "verificador.exe não esta aberto ou foi bloqueado Metodo2"
msg.Show
End If
'If FindWindow(vbNullString, "proc1") Then
'Else
'det = "verificador.exe não esta aberto ou foi bloqueado Metodo1"
'msg.Show
'End If
End Sub
