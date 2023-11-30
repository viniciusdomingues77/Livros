VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Conecta_Banco_PRINCIPAL
End Sub
Function Conecta_Banco_PRINCIPAL() As Boolean
    Dim ADOCn As New ADODB.Connection
    Dim Banco As String
    
    On Error Resume Next
    
    sConnectionString = "DRIVER={Sql Server};SERVER=DESKTOP-PM4320J;DATABASE=Livros;UID=user_livro;PWD=123456;"

    With ADOCn
       .ConnectionString = sConnectionString
       .ConnectionTimeout = DataBaseTimeOut
       .CommandTimeout = CommandTimeout
      
       .Open
    End With
    
    If ADOCn.State = 0 Then
        Conecta_Banco_PRINCIPAL = False
    Else
        Conecta_Banco_PRINCIPAL = True
    End If

End Function
