VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAssuntoCadastro 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   13740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtAssunto 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   480
      MaxLength       =   40
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1560
      Width           =   3495
   End
   Begin VB.CommandButton cmdDeletar 
      Appearance      =   0  'Flat
      Caption         =   "Deletar"
      Height          =   735
      Left            =   3960
      MaskColor       =   &H00000000&
      TabIndex        =   4
      Top             =   8520
      Width           =   2895
   End
   Begin VB.CommandButton cmdFechar 
      Appearance      =   0  'Flat
      Caption         =   "Fechar"
      Height          =   735
      Left            =   11880
      TabIndex        =   3
      Top             =   8520
      Width           =   1695
   End
   Begin VB.CommandButton cmdLimpar 
      Appearance      =   0  'Flat
      Caption         =   "Limpar"
      Height          =   735
      Left            =   8280
      TabIndex        =   2
      Top             =   8520
      Width           =   2895
   End
   Begin VB.CommandButton cmdSalvar 
      Appearance      =   0  'Flat
      Caption         =   "Salvar"
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   8520
      Width           =   2895
   End
   Begin MSDataGridLib.DataGrid GridAssunto 
      Height          =   3375
      Left            =   480
      TabIndex        =   6
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
      Index           =   0
      Left            =   480
      TabIndex        =   7
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cadastro de Assuntos"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13830
   End
End
Attribute VB_Name = "frmAssuntoCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDeletar_Click()
If GridAssunto.Row = -1 Then Exit Sub
If (MsgBox(ConfirmacaoDelecaoAssunto, vbQuestion + vbYesNo) = vbYes) Then
    AssuntoDataGridDelete GridAssunto
End If
End Sub

Private Sub cmdFechar_Click()
Unload Me
End Sub
Private Sub LimparCampos()
txtAssunto.Text = ""
End Sub

Private Sub cmdLimpar_Click()
LimparCampos
End Sub

Private Sub cmdSalvar_Click()

    If (ValidarCricacaoAssunto(txtAssunto)) Then
        If (MsgBox(ConfirmacaoInclusaoAssunto + txtAssunto.Text + ConfirmacaoInclusaoAssuntoComplemento, vbQuestion + vbYesNo) = vbYes) Then
            If (PreparaCriacaoAssunto(txtAssunto.Text, GridAssunto)) Then
                FormatGridAssuntos GridAssunto
                LimparCampos
                
                If ECadastroRapidoAssunto Then
                    Unload Me
                End If
            End If
        End If
    Else
        MsgBox AlertadeValidacao, vbExclamation
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
LimparCampos
If Not BancoLivrosConectado Then Unload Me
AssuntoDataGrid GridAssunto
FormatGridAssuntos GridAssunto
End Sub


