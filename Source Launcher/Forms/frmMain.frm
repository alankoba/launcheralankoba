VERSION 5.00
Begin VB.Form inicia 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Launcher"
   ClientHeight    =   5760
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   7680
   DrawMode        =   14  'Copy Pen
   FillStyle       =   0  'Solid
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0CCA
   ScaleHeight     =   5760
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin Launcher.PictureButton aplicar 
      Height          =   315
      Left            =   1795
      TabIndex        =   23
      Top             =   5205
      Visible         =   0   'False
      Width           =   1800
      _extentx        =   3175
      _extenty        =   556
      picture         =   "frmMain.frx":29B0
      picturehover    =   "frmMain.frx":3216
      picturedown     =   "frmMain.frx":3A7A
   End
   Begin Launcher.PictureButton cancelar 
      Height          =   315
      Left            =   3955
      TabIndex        =   22
      Top             =   5205
      Visible         =   0   'False
      Width           =   1800
      _extentx        =   3175
      _extenty        =   556
      picture         =   "frmMain.frx":43F0
      picturehover    =   "frmMain.frx":4C7E
      picturedown     =   "frmMain.frx":550A
   End
   Begin Launcher.PictureButton sair 
      Height          =   315
      Left            =   5030
      TabIndex        =   21
      Top             =   5205
      Width           =   1800
      _extentx        =   3175
      _extenty        =   556
      picture         =   "frmMain.frx":58A8
      picturehover    =   "frmMain.frx":5F7A
      picturedown     =   "frmMain.frx":664A
   End
   Begin Launcher.PictureButton opt 
      Height          =   315
      Left            =   2870
      TabIndex        =   20
      Top             =   5205
      Width           =   1800
      _extentx        =   3175
      _extenty        =   556
      picture         =   "frmMain.frx":6D14
      picturehover    =   "frmMain.frx":751E
      picturedown     =   "frmMain.frx":7D26
   End
   Begin Launcher.PictureButton jogar 
      Height          =   315
      Left            =   710
      TabIndex        =   19
      Top             =   5205
      Width           =   1800
      _extentx        =   3175
      _extenty        =   556
      picture         =   "frmMain.frx":864C
      picturehover    =   "frmMain.frx":8E76
      picturedown     =   "frmMain.frx":96A2
   End
   Begin VB.Timer configura 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5880
      Top             =   960
   End
   Begin VB.OptionButton op 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Desejo jogar em modo de tela cheia"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   330
      TabIndex        =   17
      Top             =   4350
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.OptionButton op 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Desejo jogar em modo janela"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   330
      TabIndex        =   16
      Top             =   4050
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Músicas"
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
      Left            =   1290
      TabIndex        =   14
      Top             =   2880
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Efeitos"
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
      Left            =   360
      TabIndex        =   13
      Top             =   2880
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.ComboBox cor 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2160
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox check_r 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   390
      MaxLength       =   10
      TabIndex        =   10
      Top             =   1260
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.HScrollBar volume 
      Height          =   255
      Left            =   360
      Max             =   10
      Min             =   1
      TabIndex        =   9
      Top             =   3600
      Value           =   1
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer botoes 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5880
      Top             =   480
   End
   Begin Launcher.PictureButton mi 
      Height          =   315
      Left            =   6950
      TabIndex        =   5
      Top             =   90
      Width           =   315
      _extentx        =   556
      _extenty        =   556
      picture         =   "frmMain.frx":9EC4
      picturehover    =   "frmMain.frx":A156
      picturedown     =   "frmMain.frx":A3EA
   End
   Begin Launcher.PictureButton close_ 
      Height          =   315
      Left            =   7310
      TabIndex        =   4
      Top             =   90
      Width           =   315
      _extentx        =   556
      _extenty        =   556
      picture         =   "frmMain.frx":A67C
      picturehover    =   "frmMain.frx":A94E
      picturedown     =   "frmMain.frx":AC2A
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MuOnline servidor privado"
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
      Left            =   480
      TabIndex        =   24
      Top             =   120
      Width           =   6255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Deseja jogar em modo janela?"
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
      Left            =   4620
      TabIndex        =   18
      Top             =   4500
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Volume dos efeitos:"
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
      Left            =   360
      TabIndex        =   15
      Top             =   3270
      Width           =   1485
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4350
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Opções"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   330
      TabIndex        =   7
      Top             =   570
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Volume:"
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
      Left            =   6000
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Configurações de audio:"
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
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label porcento 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   2160
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Este login sera preenchido automaticamente dentro do jogo:"
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
      Left            =   360
      TabIndex        =   1
      Top             =   930
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label lblResolucao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Configurações de vídeo:"
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
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Image img_Titulo 
      Height          =   450
      Left            =   0
      Picture         =   "frmMain.frx":AF08
      Top             =   0
      Width           =   7680
   End
End
Attribute VB_Name = "inicia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim note As Integer
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Private WithEvents m_WebControl As VBControlExtender
Attribute m_WebControl.VB_VarHelpID = -1
Dim executavel As String
Dim parametro As String
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Sub aplicar_Click()
    Call Salvar_Configurações
    'cancelar.Enabled = True
    aplicar.Visible = False
    cancelar.Visible = False
    m_WebControl.Visible = True
    Label16.Visible = False
    Text1.Visible = False
    Check1.Visible = False
    Check2.Visible = False
    Label5.Visible = False
    Label9.Visible = False
    volume.Visible = False
    porcento.Visible = False
    cor.Visible = False
    check_r.Visible = False
    lblResolucao.Visible = False
    Label3.Visible = False
    jogar.Visible = True
    opt.Visible = True
    sair.Visible = True
    op(o).Visible = False
    op(1).Visible = False

End Sub

Private Sub botoes_Timer()
    If jogar.Visible = False And aplicar.Visible = False Then
        jogar.Visible = True
        opt.Visible = True
        sair.Visible = True
        cancelar.Visible = False
    End If
End Sub

Private Sub cancelar_Click()
    aplicar.Visible = False
    cancelar.Visible = False
    m_WebControl.Visible = True
    Label16.Visible = False
    Text1.Visible = False
    Check1.Visible = False
    Check2.Visible = False
    Label5.Visible = False
    Label9.Visible = False
    volume.Visible = False
    porcento.Visible = False
    cor.Visible = False
    check_r.Visible = False
    lblResolucao.Visible = False
    Label3.Visible = False
    jogar.Visible = True
    opt.Visible = True
    sair.Visible = True

    op(o).Visible = False
    op(1).Visible = False

End Sub

Private Sub Check2_Click()
    If Check2.Value = Checked Then
        volume.Enabled = True
        volume.Value = 10
        porcento.Caption = "100%"
        porcento.Enabled = True
    Else
        volume.Enabled = False
        volume.Value = 1
        porcento.Caption = "0%"
        porcento.Enabled = False
    End If
End Sub

Private Sub close__Click()
    End
End Sub

Private Sub configura_Timer()
    check_r.AddItem "640x480"
    check_r.AddItem ("800x600")
    check_r.AddItem ("1024x768")
    check_r.AddItem ("1280x1024")

    check_r.List(0) = "640x480"
    check_r.List(1) = "800x600"
    check_r.List(2) = "1024x768"
    check_r.List(3) = "1280x1024"

    cor.AddItem "15bits"
    cor.AddItem "16bits"
    cor.AddItem "32bits"
    
    ActiveTransparency Me, True, True, 255, &HFF00FF

    Me.Caption = servername
    Unload update

    Set m_WebControl = Controls.Add("Shell.Explorer.2", "webctl", inicia)
    m_WebControl.Move 78, 450, 7520, 4580
    m_WebControl.Visible = True
    m_WebControl.object.Navigate pagina
    m_WebControl.object.silent = True

    If GetSettingLong(HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "Resolution") = 0 Then
        SaveSettingLong HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "Resolution", "0"
    End If
    configura.Enabled = False
End Sub

Private Sub Form_Load()
    configura.Enabled = True
    Label1.Caption = servername
    Me.Caption = servername
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoverFormSemCaption Button, hWnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub img_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoverFormSemCaption Button, hWnd
End Sub

Private Sub jogar_Click()
    Call function_jogar
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xx = X
    yy = Y
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoverFormSemCaption Button, hWnd
End Sub

Private Sub mi_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub opt_Click()
    botoes.Enabled = True
    aplicar.Visible = True
    cancelar.Visible = True
    m_WebControl.Visible = False
    Label16.Visible = True
    Text1.Visible = True
    Check1.Visible = True
    Check2.Visible = True
    Label5.Visible = True
    Label9.Visible = True
    volume.Visible = True
    porcento.Visible = True
    cor.Visible = True
    check_r.Visible = True
    lblResolucao.Visible = True
    Label3.Visible = True

    If windowmode = 1 Then
        op(0).Visible = True
        op(1).Visible = True
    End If


    Call Carregar_Configurações
    jogar.Visible = False
    sair.Visible = False
    opt.Visible = False
End Sub
Private Sub retornar_Click()
    ajuda_frame.Visible = False
    opt_frame.Visible = False
    inicio_frame.Visible = True
End Sub

Private Sub retornar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    retornar.Picture = hoves.retornar1.Picture
End Sub

Private Sub retornar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    retornar.Picture = hoves.retornar2.Picture
End Sub
Private Sub sair_Click()
    End
End Sub

Public Sub Carregar_Configurações()
    On Error Resume Next
    If GetSettingLong(HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "windowmode") = "1" Then
        op(0).Value = True
        op(1).Value = False
    Else
        op(0).Value = False
        op(1).Value = True
    End If
    Dim resolução As Long
    Text1.Text = GetSettingString(HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "ID")
    Check1.Value = GetSettingLong(HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "MusicOnOff")
    Check2.Value = GetSettingLong(HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "SoundOnOff")
    resolução = GetSettingLong(HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "Resolution")

    Select Case resolução
    Case "0"
        check_r.Text = "640x480"
    Case "1"
        check_r.Text = "800x600"
    Case "2"
        check_r.Text = "1024x768"
    Case "3"
        check_r.Text = "1280x1024"
    Case "4"
    End Select
    Dim volume_ As String
    volume_ = GetSettingLong(HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "volumeLevel")
    Select Case volume_
    Case "0"
        volume.Value = "1"
        porcento.Caption = CStr(volume.Value) + "0%"
    Case "1"
        volume.Value = "2"
        porcento.Caption = CStr(volume.Value) + "0%"
    Case "2"
        volume.Value = "3"
        porcento.Caption = CStr(volume.Value) + "0%"
    Case "3"
        volume.Value = "4"
        porcento.Caption = CStr(volume.Value) + "0%"
    Case "4"
        volume.Value = "5"
        porcento.Caption = CStr(volume.Value) + "0%"
    Case "5"
        volume.Value = "6"
        porcento.Caption = CStr(volume.Value) + "0%"
    Case "6"
        volume.Value = "7"
        porcento.Caption = CStr(volume.Value) + "0%"
    Case "7"
        volume.Value = "8"
        porcento.Caption = CStr(volume.Value) + "0%"
    Case "8"
        volume.Value = "9"
        porcento.Caption = CStr(volume.Value) + "0%"
    Case "9"
        volume.Value = "10"
        porcento.Caption = CStr(volume.Value) + "0%"
    End Select
    If Check2.Value = Checked Then
        volume.Enabled = True
    Else
        porcento.Caption = "0%"
        porcento.Enabled = False
        volume.Enabled = False
    End If
    Dim cor_ As String
    cor_ = GetSettingLong(HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "ColorDepth")
    Select Case cor_
    Case "0"
        cor.Text = "15bits"
    Case "1"
        cor.Text = "16bits"
    Case "2"
        cor.Text = "32bits"
    End Select

End Sub

Public Sub Salvar_Configurações()
    On Error Resume Next
    SaveSettingString HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "ID", Text1.Text
    SaveSettingLong HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "MusicOnOff", Check1.Value
    SaveSettingLong HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "SoundOnOff", Check2.Value
    If check_r.Text = "640x480" Then
        SaveSettingLong HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "Resolution", "0"
        'SaveSettingLong HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "ResolutionA", "0"
    ElseIf check_r.Text = "800x600" Then
        SaveSettingLong HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "Resolution", "1"
        'SaveSettingLong HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "ResolutionA", "0"
    ElseIf check_r.Text = "1024x768" Then
        SaveSettingLong HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "Resolution", "2"
        'SaveSettingLong HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "ResolutionA", "0"
    ElseIf check_r.Text = "1280x1024" Then
        SaveSettingLong HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "Resolution", "3"
        'SaveSettingLong HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "ResolutionA", "1"
    End If
    If op(0).Value = True Then
        SaveSettingLong HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "windowmode", "1"
    Else
        SaveSettingLong HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "windowmode", "0"
    End If
    If volume.Value = 0 Then
        SaveSettingLong HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "volumelevel", 0
    Else
        SaveSettingLong HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "volumelevel", volume.Value - 1
    End If
    Select Case cor.Text
    Case "15bits"
        SaveSettingLong HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "ColorDepth", "0"
    Case "16bits"
        SaveSettingLong HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "ColorDepth", "1"
    Case "32bits"
        SaveSettingLong HKEY_CURRENT_USER, "Software\Webzen\Mu\Config", "ColorDepth", "2"
    End Select
End Sub

Private Sub volume_Change()
    On Error Resume Next
    If volume.Value = 0 Then
        porcento.Caption = "0%"
    Else
        porcento.Caption = CStr(volume.Value) + "0%"
    End If
End Sub
Private Function function_jogar()
    On Error Resume Next
    If Dir$(App.Path & "\" + exe_nome) = "" Then
        MsgBox "O arquivo: " + exe_nome + " não foi localizado!", vbCritical, "Launcher"
    Else
            '  WinExec "muonline.exe -0", 1
            ShellExecute Me.hWnd, vbNullString, exe_nome, "connect /u" & ip & " /p" & porta, vbNullString, SW_SHOWNORMAL
            End
    End If
End Function
