VERSION 5.00
Begin VB.Form update 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Atualizando..."
   ClientHeight    =   5700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7785
   Icon            =   "Update.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Update.frx":0CCA
   ScaleHeight     =   5700
   ScaleWidth      =   7785
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer configura 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7080
      Top             =   480
   End
   Begin Launcher.PictureButton PictureButton1 
      Height          =   270
      Left            =   8280
      TabIndex        =   6
      Top             =   5160
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   476
      Picture         =   "Update.frx":106E0
      PictureHover    =   "Update.frx":11305
      PictureDown     =   "Update.frx":11F08
   End
   Begin VB.Timer timeout 
      Interval        =   15000
      Left            =   6120
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6600
      Top             =   480
   End
   Begin Launcher.Progressbar_user bar 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   4875
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   450
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
      Color           =   4210688
      Image           =   "Update.frx":12B2D
      Scrolling       =   1
      ShowText        =   -1  'True
   End
   Begin Launcher.Progressbar_user bartotal 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   4560
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   450
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
      Color           =   4210688
      Image           =   "Update.frx":149C5
      Scrolling       =   1
      ShowText        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   7080
      TabIndex        =   8
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   7440
      TabIndex        =   7
      Top             =   0
      Width           =   255
   End
   Begin VB.Label proc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6525
      TabIndex        =   5
      Top             =   4125
      Visible         =   0   'False
      Width           =   1050
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
      Left            =   480
      TabIndex        =   4
      Top             =   45
      Width           =   4455
   End
   Begin VB.Label tamanho 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   3
      Top             =   5235
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image img1 
      Height          =   270
      Index           =   0
      Left            =   8400
      Picture         =   "Update.frx":1685D
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image img 
      Height          =   270
      Index           =   0
      Left            =   7920
      Picture         =   "Update.frx":16CC7
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label status 
      BackStyle       =   0  'Transparent
      Caption         =   "Conectando-se ao servidor de atualizações..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   5235
      Width           =   4815
   End
End
Attribute VB_Name = "update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim acabou As Integer    ' pode abrir o launcher
Dim downlist As New Collection    ' lista a baixar
Dim downlen As New Collection    ' tamanho dos arquivos
Dim downname As New Collection    ' nome dos arquivos
Dim downext As New Collection    ' extrair para
Dim total_files As Integer    ' total de atualizações disponiveis

Dim nun_server As Integer
Dim url1 As String
Dim url2 As String
Dim escolha As Integer
Dim exenome As String
Dim exe_nome2 As String
Dim total_m As Integer
Private Const FLAG_ICC_FORCE_CONNECTION = &H1
Private Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
Dim mcrc As New CRC
Dim arquivo As String
Dim total As Integer
Dim crc32() As String
Dim Caminho() As String
Dim caminho2() As String
Dim crc322() As String
Dim lista As String
Dim Url As String
Dim downloas As Integer
Dim ativar As Integer
Dim I As Integer
Dim extrair_() As String
Dim ativacomando As Integer

Private Sub configura_Timer()
On Error Resume Next
servername = "MU ONLINE SERVER"    '- Nome que deve aparecer no launcher.
windowmode = 1
arquivo = "update.muonline"     '-Nome do arquivo de atualização (lembrar de colocar no gerador)
url1 = "http://update.muonline.com.br/"          '-Servidor padrão para á atualização
ip = "22.mudrop.net"
porta = "44405"
pagina = "http://muonline.com.br/launcher/"
exe_nome = "muonline.exe"
ativacomando = 1       '-Ativando essa opção o launcher só inicia com o nome "jogar.exe"
exenome = "jogar"
    

Url = url1

    lista = Url & arquivo & ".rar"
    arquivo = App.Path & "\" & arquivo
    Label3.Caption = servername

    Kill arquivo
    If Dir$(App.Path & "\unrar.dll") = vbNullString Then
        Timer1.Enabled = False
        MsgBox "Arquivo unrar.dll não encontrado.", vbCritical, "unrar.dll"
        End
    End If
    If ativacomando = 1 Then
        If App.EXEName <> exenome Then
            MsgBox "Não é possível alterar o nome do launcher.", vbCritical, "jogar.exe"
            End
        End If
    End If
    Timer1.Enabled = True
    configura.Enabled = False
End Sub

Private Sub Form_Initialize()

    On Error Resume Next
    If App.PrevInstance Then
        MsgBox "O launcher já está aberto.", vbCritical, App.Title
        End
    End If
    
    Select Case Command
    Case "update"
        Kill "unrar2.dll"
        RmDir "temp"
        Kill "unrar.rar"
    End Select
End Sub

Private Sub Form_Load()
    configura.Enabled = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xx = X
    yy = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With Me
            .Left = .Left - (xx - X)
            .Top = .Top - (yy - Y)
        End With
    End If
End Sub

Private Sub img_Click(Index As Integer)
    Me.WindowState = vbMinimized
End Sub

Private Sub img1_Click(Index As Integer)
    Kill "update.rar"
    MsgBox "Atualização concluída.", vbCritical, "Launcher"
    End
End Sub

Private Sub Label1_Click()
    On Error Resume Next
    'bCancel = True
    Kill "update.rar"
    MsgBox "Atualização concluída.", vbCritical, "Launcher"
    End
End Sub

Private Sub Label2_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub PictureButton1_Click()
'bCancel = True
    Kill "update.rar"
    MsgBox "Atualização concluída.", vbCritical, "Launcher"
    End
End Sub

Private Sub timeout_Timer()
    If bar.Value = bar.Min And bartotal.Value = bartotal.Min Then
        MsgBox "Servidor de atualizações indisponível no momento.", vbExclamation, App.Title

        Timer1.Enabled = False
        inicia.Show
        Unload Me
        timeout.Enabled = False
        Exit Sub
    End If

End Sub

Private Sub Timer1_Timer()
    On Error Resume Next

  
    Call DownloadAFile(App.Path & "\" & "update.rar", lista, False, True)
    Call RARExtract("update.rar", "")

    Kill "update.rar"
    If Dir$(arquivo) = "" Then
        MsgBox "Servidor de atualizações indisponível no momento.", vbExclamation, "Atenção..."

        inicia.Show
        Unload Me
        Exit Sub
    End If

    total = ReadINI(arquivo, "UPDATE", "total")
    bartotal.Max = total
    ReDim crc32(0 To total)
    ReDim Caminho(0 To total)
    ReDim caminho2(0 To total)
    ReDim crc322(0 To total)
    For I = 0 To total
        status.Caption = "Listando arquivos desatualizados -  " & I & " / " & total
        status.Refresh
        lKnownSize = (ReadINI(arquivo, I, "len"))
        gets = lKnownSize

        'bartotal.Value = i

        Caminho(I) = ReadINI(arquivo, I, "caminho")
        crc32(I) = ReadINI(arquivo, I, "crc32")
        caminho2(I) = Caminho(I)
        crc322(I) = mcrc.CRC(Caminho(I))
        '''''' Verificando se a atualização é do jogar.exe ou do unrar.dll
        If Caminho(I) = exenome & ".exe" And crc32(I) <> crc322(I) Then
            status.Caption = "Baixando atualização do: " + exenome + ".exe"
            If Dir$(App.Path & "\atualiza.exe") = "" Then
                MsgBox "O arquivo atualiza.exe não foi localizado.", vbCritical, App.Title
                End
            End If
            Call DownloadAFile(App.Path & "\jogar.rar", Url & Caminho(I) & ".rar", False, True)
            ShellExecute Me.hWnd, vbNullString, "atualiza.exe", "jogar", vbNullString, SW_SHOWNORMAL
            End
        End If
        ' atualiza.exe
        If Caminho(I) = "atualiza.exe" And crc32(I) <> crc322(I) Then
            status.Caption = "Baixando atualização do: atualiza.exe"
            Call DownloadAFile(App.Path & "\update.rar", Url & Caminho(I) & ".rar", False, True)
            Call RARExtract("update.rar", "")
            Kill "update.rar"
        End If
        '' atualzia.exe
        If Caminho(I) = "unrar.dll" And crc32(I) <> crc322(I) Then
            status.Caption = "Baixando atualização do unrar.dll"
            If Dir$(App.Path & "\atualiza.exe") = "" Then
                MsgBox "O arquivo atualiza.exe não foi localizado.", vbCritical, App.Title
                End
            End If
            Call DownloadAFile(App.Path & "\unrar.rar", Url & Caminho(I) & ".rar", False, True)
            ShellExecute Me.hWnd, vbNullString, "atualiza.exe", "unrar", vbNullString, SW_SHOWNORMAL
            End
        End If
        If crc32(I) <> crc322(I) And Caminho(I) <> "unrar.dll" And Caminho(I) <> exenome + ".exe" Then
            ReDim extrair_(0 To total)
            extrair_(I) = GetFilePath(Caminho(I))

            total_files = total_files + 1
            downlist.Add (Url + Caminho(I) + ".rar")
            downlen.Add (CStr(I) & "/" & CStr(total))
            downname.Add (Caminho(I))
            downext.Add (extrair_(I))

        Else
            If I >= total And downlist.Count = 0 Then
                inicia.Show vbNormal
                Timer1.Enabled = False
                status.Caption = "Atualização concluída."
                status.Refresh
                tamanho.Caption = ""
                tamanho.Refresh
                Unload Me
                Exit Sub
            End If
            bartotal.Value = bartotal.Max
        End If
    Next I

    Call BaixarArquivos
    Timer1.Enabled = False
End Sub

Function BaixarArquivos()
    For I = 1 To downlist.Count
        bartotal.Max = total_files
        bartotal.Value = I
        tamanho.Visible = True
        status.Caption = "Baixando: " + downname.Item(I)
        status.Refresh
        Call DownloadAFile(App.Path & "\update.rar", downlist(I), False, True)
       '  tamanho.Caption = CStr(i) & "/" & CStr(total_files)

        proc.Caption = CStr(I) + "/" + CStr(total_files)
        If downext.Item(I) = "" Then
            Call RARExtract("update.rar", "")
        Else
            Call RARExtract("update.rar", downext.Item(I) & "\")

        End If

        Kill "update.rar"
        If I >= total_files Then
            status.Caption = "Atualização concluída."
            status.Refresh
            inicia.Show vbNormal
            timeout.Enabled = False
            Exit For
            Unload Me
        End If

    Next I
End Function




