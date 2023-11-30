VERSION 5.00
Begin VB.MDIForm MDIMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Livros"
   ClientHeight    =   11190
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   14370
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   11175
      Left            =   0
      ScaleHeight     =   11115
      ScaleWidth      =   14310
      TabIndex        =   0
      Top             =   0
      Width           =   14370
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6135
         Left            =   3480
         Picture         =   "MDIMain.frx":0000
         ScaleHeight     =   6135
         ScaleWidth      =   7215
         TabIndex        =   1
         Top             =   2040
         Width           =   7215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Livros"
         BeginProperty Font 
            Name            =   "Pangolin"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         TabIndex        =   2
         Top             =   960
         Width           =   14295
      End
   End
   Begin VB.Menu mnuCadastros 
      Caption         =   "Cadastros"
      Begin VB.Menu mnuLivros 
         Caption         =   "Livros"
      End
      Begin VB.Menu mnuCadastraAutor 
         Caption         =   "Autor"
      End
      Begin VB.Menu mnuCadastraAssunto 
         Caption         =   "Assuntos"
      End
   End
   Begin VB.Menu mnuRelatorio 
      Caption         =   "Relatorios"
      Begin VB.Menu mnuRelatorioLivro 
         Caption         =   "Livros"
      End
   End
   Begin VB.Menu mnuSair 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
ConectaBancoLivros
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
DesconectaBancoLivros
End Sub

Private Sub mnuCadastraAssunto_Click()
    frmAssuntoCadastro.Show vbModal
End Sub

Private Sub mnuCadastraAutor_Click()
    frmAutorCadastro.Show vbModal
End Sub

Private Sub mnuLivros_Click()
    
    frmLivroCadastro.Show vbModal
End Sub

Private Sub mnuRelatorioLivro_Click()
    frmRelatorioLivros.Show vbModal
End Sub

Private Sub mnuSair_Click()
DesconectaBancoLivros
End
End Sub
