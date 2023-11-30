VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmLivroCadastro 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   9435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13830
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   9435
   ScaleWidth      =   13830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCadAssunto 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   255
      Left            =   7920
      TabIndex        =   21
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton cmdCadAutor 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   255
      Left            =   3840
      TabIndex        =   20
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton cmdDeletar 
      Appearance      =   0  'Flat
      Caption         =   "Deletar"
      Height          =   735
      Left            =   3960
      MaskColor       =   &H00000000&
      TabIndex        =   19
      Top             =   8520
      Width           =   2895
   End
   Begin MSDataGridLib.DataGrid GridLivro 
      Height          =   3375
      Left            =   480
      TabIndex        =   18
      Top             =   5040
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdFechar 
      Appearance      =   0  'Flat
      Caption         =   "Fechar"
      Height          =   735
      Left            =   11880
      TabIndex        =   17
      Top             =   8520
      Width           =   1695
   End
   Begin VB.CommandButton cmdLimpar 
      Appearance      =   0  'Flat
      Caption         =   "Limpar"
      Height          =   735
      Left            =   8280
      TabIndex        =   16
      Top             =   8520
      Width           =   2895
   End
   Begin VB.CommandButton cmdAutorAdicionar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   255
      Left            =   3480
      TabIndex        =   15
      Top             =   2400
      Width           =   255
   End
   Begin VB.ComboBox cboAssunto 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4920
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   2400
      Width           =   2655
   End
   Begin VB.ListBox ListAutores 
      Height          =   2010
      Left            =   480
      TabIndex        =   12
      Top             =   2880
      Width           =   2655
   End
   Begin VB.ComboBox cboAutor 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   2400
      Width           =   2655
   End
   Begin VB.CommandButton cmdSalvar 
      Appearance      =   0  'Flat
      Caption         =   "Salvar"
      Height          =   735
      Left            =   480
      TabIndex        =   9
      Top             =   8520
      Width           =   2895
   End
   Begin VB.TextBox txtAno 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   11640
      MaxLength       =   4
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtEdicao 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   8760
      MaxLength       =   7
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox txtEditora 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4920
      MaxLength       =   40
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1200
      Width           =   3495
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   480
      MaxLength       =   40
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Assunto"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4920
      TabIndex        =   14
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Autor"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   480
      TabIndex        =   11
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ano de Publicação"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   11640
      TabIndex        =   8
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Edição"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   8760
      TabIndex        =   5
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Editora"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Título"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cadastro de Livros"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13815
   End
End
Attribute VB_Name = "frmLivroCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adorsDataGrid As New ADODB.Recordset
Private Sub cmdAutorAdicionar_Click()
If Not VerificaExisteAutorSelecionado(cboAutor) Then Exit Sub
AssociaAutoraoLivro cboAutor, ListAutores
End Sub

Private Sub cmdCadAssunto_Click()
CadastroRapidodeAssunto = True
frmAssuntoCadastro.Show vbModal
LivrosDataGrid GridLivro
CadastroRapidodeAssunto = False
CarregaAssuntosParaComboBox cboAssunto

End Sub

Private Sub cmdCadAutor_Click()
CadastroRapidodeAutor = True
frmAutorCadastro.Show vbModal
LivrosDataGrid GridLivro
CadastroRapidodeAutor = False
CarregaAutoresParaComboBox cboAutor
GridLivro.ReBind
End Sub

Private Sub cmdDeletar_Click()
If GridLivro.Row = -1 Then Exit Sub
If (MsgBox(ConfirmacaoDelecaoLivro, vbQuestion + vbYesNo) = vbYes) Then
    LivrosDataGridDelete GridLivro
End If
End Sub

Private Sub cmdFechar_Click()
Unload Me
End Sub

Private Sub cmdLimpar_Click()
LimparCampos
End Sub
Private Sub cmdSalvar_Click()
If Not Validacao Then
    MsgBox AlertadeValidacao, vbExclamation
    Exit Sub
End If
If (MsgBox(ConfirmacaoInclusaoLivro + txtTitulo.Text + ConfirmacaoInclusaoLivroComplemento, vbQuestion + vbYesNo) = vbYes) Then
    If (PreparaCadastroLivro(txtTitulo.Text, txtEditora.Text, txtEdicao.Text, txtAno.Text, ListAutores, cboAssunto, GridLivro)) Then
        LimparCampos
    End If
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
LimparCampos

If Not BancoLivrosConectado Then
    Unload Me
End If
CarregaAutoresParaComboBox cboAutor
CarregaAssuntosParaComboBox cboAssunto
LivrosDataGrid GridLivro
FormatGridLivros GridLivro

End Sub

Private Sub Form_Unload(Cancel As Integer)
FecharRecordsetDataGRid

End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
If Not ENumero(KeyAscii) Then
    KeyAscii = 0
End If
End Sub
Private Sub txtEdicao_KeyPress(KeyAscii As Integer)
If Not ENumero(KeyAscii) Then
    KeyAscii = 0
End If
End Sub
Function Validacao() As Boolean
If Len(txtTitulo.Text) = 0 Then Exit Function
If Len(txtEditora.Text) = 0 Then Exit Function
If Len(txtAno.Text) < 4 Then Exit Function
If Len(txtAno.Text) = 0 Then Exit Function
If ListAutores.ListCount = 0 Then Exit Function
If cboAssunto.ListIndex = -1 Then Exit Function
If cboAutor.ListIndex = -1 Then Exit Function
Validacao = True
End Function

Private Sub LimparCampos()
    txtTitulo.Text = ""
    txtEditora.Text = ""
    txtAno.Text = ""
    txtEdicao.Text = ""
    cboAutor.ListIndex = -1
    cboAssunto.ListIndex = -1
    ListAutores.Clear
End Sub
