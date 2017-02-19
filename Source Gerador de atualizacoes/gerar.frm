VERSION 5.00
Begin VB.Form gerar 
   Caption         =   "Tutorial gerando atualização."
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5985
   LinkTopic       =   "Form2"
   ScaleHeight     =   3255
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.Line Line2 
      X1              =   0
      X2              =   6000
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label8 
      Caption         =   "Observação: a pasta update. deve estar no mesmo diretorio do gerador.exe"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   5535
   End
   Begin VB.Label Label7 
      Caption         =   $"gerar.frx":0000
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   5655
   End
   Begin VB.Label Label6 
      Caption         =   "Atualizando Patch."
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   1680
      Width           =   4935
   End
   Begin VB.Line Line1 
      X1              =   6000
      X2              =   0
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label5 
      Caption         =   "4 - Copie todos os arquivos da pasta update para sua hospedagen."
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   5655
   End
   Begin VB.Label Label4 
      Caption         =   "Clique em ok na mensagen e não feche nenhuma janela que aparecer."
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   960
      Width           =   5415
   End
   Begin VB.Label Label3 
      Caption         =   "3 - Clique em gerar atualização."
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   5295
   End
   Begin VB.Label Label2 
      Caption         =   "2 - clique em listar  arquivos"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "1 - Coloque todos os seus arquivos na pasta Update/ "
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "gerar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
