VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerador de atualizações"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9495
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Gerar atualização"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   7800
      Width           =   1935
   End
   Begin VB.Timer tmr1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   12480
      Top             =   480
   End
   Begin VB.ListBox lst2 
      Enabled         =   0   'False
      Height          =   7470
      ItemData        =   "Form1.frx":030A
      Left            =   120
      List            =   "Form1.frx":0311
      TabIndex        =   2
      Top             =   120
      Width           =   9255
   End
   Begin VB.CommandButton cmdProcurar 
      Caption         =   "&Listar"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   7800
      Width           =   1335
   End
   Begin VB.ListBox lst1 
      Enabled         =   0   'False
      Height          =   255
      Left            =   9480
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   -120
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   7800
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1
Dim fso As New FileSystemObject
Dim cont As Integer
Dim crc32() As String
Public bytCancelaBusca As Byte
Dim exa As String
Dim exo As String
Dim len_() As String
Dim len2() As String
'Declares
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetCompressedFileSize Lib "kernel32" Alias "GetCompressedFileSizeA" (ByVal lpFileName As String, lpFileSizeHigh As Long) As Long

Private Sub cmdProcurar_Click()
    tmr1.Enabled = True
    bytCancelaBusca = 0
    
    Command1.Enabled = True
End Sub

Sub ProcuraPasta(xis As String)
'On Error GoTo erro
'esse "xis" é a variável 'caminho' passada como parâmetro para esta função.
    Dim Pasta As Folder
    Dim Arquivo As File
    
    Set PastaPai = fso.GetFolder(xis)
    
    If Not Mid(PastaPai, 4, Len(PastaPai) - 3) = "System Volume Information" Then
        
        Set TodosArquivos = PastaPai.Files
        
        For Each Arquivo In TodosArquivos
            
            lst2.AddItem Mid(Arquivo, Len(CurDir) + 9, 1000)
            
            DoEvents
            If bytCancelaBusca = 1 Then
'MsgBox "Busca cancelada.", vbInformation
                Exit Sub
            End If
            
        Next
        
        Set SubPastas = PastaPai.SubFolders
        
        For Each PastaAtual In SubPastas
            
            lst1.AddItem PastaAtual
            
'Faz uma nova busca, recursivamente.
            ProcuraPasta (PastaAtual.Path)
            
            DoEvents
            If bytCancelaBusca = 1 Then
                Exit Sub
            End If
            
        Next
        
    End If
    Set PastaPai = Nothing
    Set SubPastas = Nothing
End Sub

Private Sub Command1_Click()
    Dim hWndDesktop As Long
    Dim Rec As RECT
    hWndDesktop = GetDesktopWindow
    GetWindowRect hWndDesktop, Rec
    Label1.Caption = "Total de arquivos listados" + " " + CStr(lst2.ListCount - 1)


    Dim mCRC As New CRC
    
    Dim exe As String
    exe = "rar.exe"
    Dim par As String
    
    For I = lst2.ListIndex + 1 To lst2.ListCount - 1
    If Right(lst2.List(I), 4) = ".rar" Then
        MsgBox "Arquivos compactados encontrados na pasta.", vbCritical, "Atenção..."
        Exit Sub
    End If
    ReDim crc32(0 To lst2.ListCount - 1)
    ReDim len_(0 To lst2.ListCount - 1)
    
    crc32(I) = mCRC.CRC("update\" + lst2.List(I))
    
    WriteINI exa, "UPDATE", "total", I + 1
    WriteINI exa, I, "caminho", lst2.List(I)
    WriteINI exa, I, "crc32", crc32(I)
    
    len_(I) = CStr(FileLen("update\" + lst2.List(I)))
    WriteINI exa, I, "len", len_(I)
    par = "a -ac -cl -df -ep -m5" + " " + """" + "update\" + lst2.List(I) + ".rar" + """" + " " + """" + "update\" + lst2.List(I) + """"
'    If InStr(1, lst2.List(I), " ") Then
'    MsgBox "arquivos com espaços encontrados", vbCritical
'    End If
    ShellExecute hWndDesktop, vbNullString, "rar.exe", "a -ac -cl -df -ep -m5" + " " + """" + "update\" + lst2.List(I) + ".rar" + """" + " " + """" + "update\" + lst2.List(I) + """", vbNullString, 1
    Sleep (80)
    
    Next I
    On Error Resume Next
    FileCopy exa, "update\" + exo
    Dim par2 As String
    par2 = "a -ac -cl -df -ep -m5" + " " + """" + "update\" + exo + ".rar" + """" + " " + """" + "update\" + exo + """"
    ShellExecute Me.hwnd, vbNullString, exe, par2, vbNullString, SW_SHOWNORMAL
    Kill exa
End Sub
    
Private Sub Command2_Click()
    gerar.Show vbModal
End Sub

Private Sub Form_Load()
    exa = App.Path & "\" + ReadINI(App.Path & "\config.ini", "CONFIG", "name")
    exo = ReadINI(App.Path & "\config.ini", "CONFIG", "name")
    bytCancelaBusca = 0
End Sub
    
Private Sub tmr1_Timer()
    lst1.Clear
    lst2.Clear
    ProcuraPasta ("update")
    tmr1.Enabled = False
End Sub
Public Function ListBoxToArray(ctrl As Control) As String()
    Dim index As Long
    If ctrl.ListCount > 0 Then
        ReDim res(lst2.ListIndex + 1 To ctrl.ListCount - 1) As String
        
        For index = lst2.ListIndex + 1 To lst2.ListCount - 1
            res(index) = lst2.List(index)
        Next
    End If
    ListBoxToArray = res()
End Function
