VERSION 5.00
Begin VB.Form atualiza 
   BorderStyle     =   0  'None
   Caption         =   "Atualizando..."
   ClientHeight    =   1545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3975
   Icon            =   "atualiza.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   1545
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   1500
         Left            =   1080
         Top             =   240
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1200
         Left            =   600
         Top             =   240
      End
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   120
         Top             =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Aguarde enquanto subistituimos os arquivos."
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "O  jogar.exe sera aberto automaticamente apos o procedimento de atualização!"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   3255
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   0
         X2              =   0
         Y1              =   1440
         Y2              =   120
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   3720
         X2              =   0
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   3720
         X2              =   3720
         Y1              =   1440
         Y2              =   120
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   3720
         X2              =   0
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "Atualizando o:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
   End
End
Attribute VB_Name = "atualiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Const MF_BYPOSITION = &H400&
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim exenome As String
Private Sub RemoveMenus()
 Dim hMenu As Long
 hMenu = GetSystemMenu(hwnd, False)
 DeleteMenu hMenu, 6, MF_BYPOSITION
End Sub

Private Sub bar_GotFocus()

End Sub

Private Sub Form_Load()
exenome = "jogar.exe"
If Command = "" Then
ShellExecute Me.hwnd, vbNullString, exenome, vbNullString, vbNullString, SW_SHOWNORMAL
End
End If
Label1.Caption = Label1.Caption + " " + exenome
On Error Resume Next
If Command = "unrar" Then
Label1.Caption = "Atualizando o unrar.dll"
End If
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If FileLen(App.Path & "\jogar.rar") < 10000 Then
MsgBox "Download corrompido, por favor baixe nosso patch no site." + vbNewLine + "ou verifique sua conectividade com a internet!", vbCritical, "atualzia.exe"
End
Else
Call RARExtract("jogar.rar", App.Path & "/")
ShellExecute Me.hwnd, vbNullString, exenome, "update", vbNullString, SW_SHOWNORMAL
Kill "jogar.rar"
End
End If
Timer2.Enabled = False
End Sub

Private Sub Timer1_Timer()
    If Command = "jogar" Then
    Timer2.Enabled = True
    ElseIf Command = "unrar" Then
    Timer3.Enabled = True
    End If
End Sub

Private Sub Timer3_Timer()
    On Error Resume Next
    If FileLen(App.Path & "\unrar.rar") < 10000 Then
    MsgBox "Download corrompido, por favor baixe nosso patch no site." + vbNewLine + "ou verifique sua conectividade com a internet!", vbCritical, "atualzia.exe"
    End
    Else
     MkDir "temp"
    Call RARExtract("unrar.rar", "temp\")
    Sleep (2500)
    Name "unrar.dll" As "unrar2.dll"
    MoveFile "temp/unrar.dll", App.Path & "\unrar.dll"
    Sleep (500)
    DeleteFile "unrar2.dll"
    RmDir "temp"
    Kill "unrar.rar"
    ShellExecute Me.hwnd, vbNullString, exenome, "update", vbNullString, SW_SHOWNORMAL
    End
    End If
    Timer3.Enabled = False
End Sub

