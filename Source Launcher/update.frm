VERSION 5.00
Begin VB.Form update 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Atualizando..."
   ClientHeight    =   1770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8130
   Icon            =   "update.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "update.frx":0CCA
   ScaleHeight     =   1770
   ScaleWidth      =   8130
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timeout 
      Interval        =   15000
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1080
      Top             =   0
   End
   Begin Launcher.Progressbar_user bar 
      Height          =   195
      Left            =   400
      TabIndex        =   0
      Top             =   960
      Width           =   4540
      _ExtentX        =   8017
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   0
      Scrolling       =   9
   End
   Begin Launcher.Progressbar_user bartotal 
      Height          =   195
      Left            =   400
      TabIndex        =   1
      Top             =   720
      Width           =   4540
      _ExtentX        =   8017
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   0
      Scrolling       =   9
   End
   Begin VB.Label proc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "servername"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   440
      TabIndex        =   4
      Top             =   420
      Width           =   1335
   End
   Begin VB.Label tamanho 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Image img1 
      Height          =   195
      Index           =   0
      Left            =   7800
      Picture         =   "update.frx":30B16
      Top             =   120
      Width           =   210
   End
   Begin VB.Image img 
      Height          =   195
      Index           =   0
      Left            =   7560
      Picture         =   "update.frx":30D94
      Top             =   120
      Width           =   210
   End
   Begin VB.Label status 
      BackStyle       =   0  'Transparent
      Caption         =   "Conectando-se ao servidor de atualizações..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   400
      TabIndex        =   2
      Top             =   1200
      Width           =   3375
   End
End
Attribute VB_Name = "update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nun_server As Integer
Dim url1 As String
Dim url2 As String
Dim escolha As Integer
Private Declare Function IsNTAdmin Lib "advpack.dll" (ByVal dwReserved As Long, ByRef lpdwReserved As Long) As Long
Dim exenome As String
Dim exe_nome2 As String
Dim pagina2 As String
Dim total_m As Integer
Private Declare Function CreateRoundRectRgn Lib _
        "gdi32" (ByVal x1 As Long, ByVal y1 As _
        Long, ByVal x2 As Long, ByVal y2 As Long, _
        ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "User32" _
        (ByVal hWnd As Long, ByVal hRgn As Long, _
        ByVal bRedraw As Boolean) As Long
Private Declare Function GetClientRect Lib "User32" _
        (ByVal hWnd As Long, lpRect As RECT) As Long
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "User32" ()
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Dim updates_m As Integer ' numero de arquivos encontrados
Private Const FLAG_ICC_FORCE_CONNECTION = &H1
Private Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
Dim mCRC As New CRC
Dim arquivo As String
Dim total As Integer
Dim crc32() As String
Dim Caminho() As String
Dim caminho2() As String
Dim crc322() As String
Dim lista As String
Dim url As String
Dim downloas As Integer
Dim ativar As Integer
Dim i As Integer
Dim extrair_() As String

Sub Retangulo(m_hWnd As Long, Fator As Byte)
  Dim RGN As Long
  Dim RC As RECT
  Call GetClientRect(m_hWnd, RC)
  RGN = CreateRoundRectRgn(RC.Left, RC.Top, RC.Right, _
                           RC.Bottom, Fator, Fator)
  SetWindowRgn m_hWnd, RGN, True
End Sub

Private Sub Form_Initialize()
On Error Resume Next
Select Case Command
Case "update"
Kill "unrar2.dll"
RmDir "temp"
Kill "unrar.rar"
End Select
End Sub

Private Sub Form_Load()

On Error Resume Next


''''''''''''''''''''' Configurações
servername = "MU-PVP"
arquivo = "update.mupvp"
nun_server = 2
url1 = "http://mupvp.net/update/"
url2 = "http://mupvp.net/update/"
ip = "mu-pvp.sytes.net"
porta = "44405"
pagina = "http://mupvp.net/update/"
exe_nome = "pvp.exe"
'''''''''''''''''''''' Fim das configurações
'''' Repreenchendo dados
    If nun_server = 2 Then
        Randomize
         escolha = Rnd(1) * 2
        If escolha = 0 Or escolha = 1 Then
        url = url1
        ElseIf escolha = 2 Then
        url = url2
        Else
        url = url1
        End If
    Else
    url = url1
    End If
lista = url + arquivo + ".rar"
arquivo = App.Path & "\" + arquivo
Label3.Caption = servername
'''' Fim de todos os prenechiemntos
Call RetiraAcentos(CurDir)
Kill arquivo
Retangulo Me.hWnd, 18
If Dir$(App.Path & "\unrar.dll") = "" Then
Timer1.Enabled = False
MsgBox "A dependencia de unrar.dll não foi localizada!", vbCritical, "unrar.dll"
End
End If
exenome = "jogar"
Dim ativacomando As Integer
ativacomando = 0
'lista = "http://v8server.net/update/update.mulp.rar" ' invertido
'url = "http://v8server.net/update/" ' invertido
If ativacomando = 1 Then
    If App.EXEName <> exenome Then
    MsgBox "O Launcher precisa chamar: " + exenome + vbNewLine + "para o correto funcionamento!", vbCritical, App.Title
    End
    End If
End If

'SHCreateThread AddressOf ativacao, ByVal 0&, &H1, ByVal 0&
Timer1.Enabled = True
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoverFormSemCaption Button, hWnd
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Timer1.Enabled = False
'Set update = Nothing
End Sub

Private Sub img_Click(Index As Integer)
Me.WindowState = vbMinimized
End Sub

Private Sub img1_Click(Index As Integer)
On Error Resume Next
'bCancel = True
Kill "update.rar"
MsgBox "Atualização finalizada!", vbCritical, "Launcher"
End
End Sub





Private Sub timeout_Timer()
On Error Resume Next
If bar.Value = bar.Min And bartotal.Value = bartotal.Min Then
        MsgBox "Falha ao connectar no servidor de atualizações!", vbCritical, App.Title
       
       Timer1.Enabled = False
        inicia.Show
        Unload Me
        timeout.Enabled = False
        Exit Sub
        End If

End Sub

Private Sub Timer1_Timer()
'If App.PrevInstance Then
'MsgBox "O jogo ja se encontra aberto!", vbCritical, App.Title
'End
'End If
On Error Resume Next

     If InternetCheckConnection(lista, FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then
       timeout.Enabled = False
       Timer1.Enabled = False
        MsgBox "Falha ao connectar no servidor de atualizações!", vbCritical, App.Title
       
        inicia.Show vbNormal
        timeout.Enabled = False
        Unload Me
        Exit Sub
    End If

  '  bCancel = False
  '  SHCreateThread AddressOf DownloadAFile, ByVal 0&, &H1, ByVal 0&
Call DownloadAFile(CurDir + "\" + "update.rar", lista, False, True)

   Call RARExtract("update.rar", CurDir + "\" + "\")
   Kill "update.rar"
    
    If Dir$(arquivo) = "" Then
    MsgBox "Falha ao baixar o arquivo de atualização!" + vbNewLine + "Observação: A atualização pode estar desativada no momento!" + vbNewLine + "Verifique se a pasta do jogo não conten acentos ou caracteres especiais.", vbCritical, App.Title
          
        inicia.Show
     '   Set update = Nothing
        Unload Me
                Exit Sub
        End If

   total = ReadINI(arquivo, "UPDATE", "total")
    status.Caption = "Listando arquivos desatualizados..."
    status.Refresh
   '''


   ReDim crc32(0 To total)
ReDim Caminho(0 To total)
ReDim caminho2(0 To total)
ReDim crc322(0 To total)
For i = 0 To total

lKnownSize = ReadINI(arquivo, i, "len")
gets = ReadINI(arquivo, i, "len")
bartotal.Value = updates_m / (total / 100)

'
On Error Resume Next
Caminho(i) = ReadINI(arquivo, i, "caminho")
crc32(i) = ReadINI(arquivo, i, "crc32")
caminho2(i) = Caminho(i)
crc322(i) = mCRC.CRC(Caminho(i))
''''''
   If Caminho(i) = exenome + ".exe" And crc32(i) <> crc322(i) Then
status.Caption = "Baixando atualização do: " + exenome + ".exe"
    If Dir$(App.Path & "\atualiza.exe") = "" Then
    MsgBox "atualiza.exe não encontrado!", vbCritical, App.Title
    End
    End If
Call DownloadAFile(App.Path & "\jogar.rar", url + Caminho(i) + ".rar", False, True)
ShellExecute Me.hWnd, vbNullString, "atualiza.exe", "jogar", vbNullString, SW_SHOWNORMAL
End
End If
If Caminho(i) = "unrar.dll" And crc32(i) <> crc322(i) Then
status.Caption = "Baixando atualização do unrar.dll"
    If Dir$(App.Path & "\atualiza.exe") = "" Then
    MsgBox "atualiza.exe não encontrado!", vbCritical, App.Title
    End
    End If
Call DownloadAFile(App.Path & "\unrar.rar", url + Caminho(i) + ".rar", False, True)
ShellExecute Me.hWnd, vbNullString, "atualiza.exe", "unrar", vbNullString, SW_SHOWNORMAL
End
End If
updates_m = updates_m + 1
    If crc32(i) <> crc322(i) And Caminho(i) <> "unrar.dll" And Caminho(i) <> exenome + ".exe" Then
   status.Caption = "Baixando: " + Caminho(i)
   
   proc.Caption = CStr(updates_m) + "/" + CStr(total)
   ReDim extrair_(0 To total)
       extrair_(i) = GetFilePath(Caminho(i))
        Dim pastas() As String
        pastas() = Split(extrair_(i), "\")
      If ArquivoExiste(extrair_(i)) = False And extrair_(i) <> "" Then
        Dim tot As Integer
        
        tot = UBound(pastas)
            If tot = 1 Then
                If ArquivoExiste(pastas(0)) Then
            
            MkDir pastas(0) + "/" + pastas(1)
            Else
             MkDir pastas(0)
            MkDir pastas(0) + "/" + pastas(1)
                End If
            End If
                        If tot = 0 Then
            MkDir pastas(0)
            End If
                        If tot = 2 Then
                            If ArquivoExiste(pastas(0)) Then
                                If ArquivoExiste(pastas(1)) Then
                                MkDir pastas(0) + "/" + pastas(1) + "/" + pastas(2)
                                End If
                                MkDir pastas(0) + "/" + pastas(1)
            MkDir pastas(0) + "/" + pastas(1) + "/" + pastas(2)
            Else
            MkDir pastas(0)
            MkDir pastas(0) + "/" + pastas(1)
            MkDir pastas(0) + "/" + pastas(1) + "/" + pastas(2)
                            End If
            End If
            
            If tot = 3 Then
            On Error Resume Next
                MkDir pastas(0)
                MkDir pastas(0) + "/" + pastas(1)
                MkDir pastas(0) + "/" + pastas(1) + "/" + pastas(2)
                MkDir pastas(0) + "/" + pastas(1) + "/" + pastas(2) + "/" + pastas(3)
         End If
         If tot = 4 Then
         On Error Resume Next
                MkDir pastas(0)
                MkDir pastas(0) + "/" + pastas(1)
                MkDir pastas(0) + "/" + pastas(1) + "/" + pastas(2)
                MkDir pastas(0) + "/" + pastas(1) + "/" + pastas(2) + "/" + pastas(3)
                MkDir pastas(0) + "/" + pastas(1) + "/" + pastas(2) + "/" + pastas(3) + "/" + pastas(4)
         End If
        End If
Call DownloadAFile(App.Path & "\update.rar", url + Caminho(i) + ".rar", False, True)
 tamanho.Caption = CStr(i) + "/" + CStr(total)

            If extrair_(i) = "" Then
            Call RARExtract("update.rar", App.Path & "/")
            Else
            Call RARExtract("update.rar", extrair_(i) + "/")
            End If
   Kill "update.rar"

   If i = total Then
   status.Caption = "Atualização concluida!"
   status.Refresh
inicia.Show vbNormal
      timeout.Enabled = False
Exit For
  Exit Sub
Unload Me
  Exit Sub
End If
Else
status.Caption = "Verificando proximo arquivo...."
'tamanho.Caption = CStr(i) + "/" + CStr(total)
'tamanho.Refresh
'status.Refresh
bartotal.Value = bartotal.Max
If i = total Then
   status.Caption = "Atualização concluida!"
   status.Refresh
   tamanho.Caption = ""
   tamanho.Refresh
inicia.Show vbNormal
timeout.Enabled = False
'Me.Hide
Exit For
Unload Me
  Exit Sub
End If
End If
Next i
   Timer1.Enabled = False
End Sub












