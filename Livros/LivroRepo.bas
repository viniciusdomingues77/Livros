Attribute VB_Name = "LivroRepositorio"
Public adorsDataGrid As New ADODB.Recordset
Function LivrosDataGrid(LivroGrid As DataGrid) As ADODB.Recordset
    
    On Error GoTo Exception
    Set adors = New ADODB.Recordset
    FecharRecordsetDataGRid
    Dim SQL As String
    SQL = "select Codigo,Titulo,Editora,Edicao,AnoPublicacao, Assunto,Autor from LivrosListaAutoresVisao order by Titulo,AnoPublicacao"
    Set dbCommand = New ADODB.Command
    Set dbCommand.ActiveConnection = ADOCn
    dbCommand.CommandType = adCmdText
    dbCommand.CommandText = SQL
    adorsDataGrid.CursorLocation = adUseClient
    adorsDataGrid.Open SQL, ADOCn, adOpenDynamic, adLockPessimistic
    
    Set LivroGrid.DataSource = adorsDataGrid
    Set LivrosDataGrid = adorsDataGrid
    Exit Function
Exception:
    MsgBox ProblemasEncontrados, vbExclamation
End Function
Function LivrosReport() As ADODB.Recordset
    
    On Error GoTo Exception
    Set adors = New ADODB.Recordset
    FecharRecordsetDataGRid
    Dim SQL As String
    SQL = "select Codigo,Titulo,Editora,Edicao,AnoPublicacao, Assunto,Autor from LivrosListaAutoresVisao order by Titulo,AnoPublicacao"
    Set dbCommand = New ADODB.Command
    Set dbCommand.ActiveConnection = ADOCn
    dbCommand.CommandType = adCmdText
    dbCommand.CommandText = SQL
    adorsDataGrid.CursorLocation = adUseClient
    adorsDataGrid.Open SQL, ADOCn, adOpenDynamic, adLockPessimistic
    Set LivrosReport = adorsDataGrid
    
    Exit Function
Exception:
    MsgBox ProblemasEncontrados, vbExclamation
End Function
Public Function FecharRecordsetDataGRid()
    If adorsDataGrid.State = 1 Then adorsDataGrid.Close
   
End Function
Public Function DeletaLivro(CodigoLivro As String) As Boolean
    Dim SQLCmdLivro As New ADODB.Command
    On Error GoTo Exception
    SQLCmdLivro.ActiveConnection = ADOCn
    SQLCmdLivro.Parameters.Append SQLCmdLivro.CreateParameter("@CodLivro", adInteger, adParamInput, 4, CodigoLivro)
    SQLCmdLivro.CommandText = "dbo.DeletarLivro"
    SQLCmdLivro.CommandType = adCmdStoredProc
    ADOCn.BeginTrans
    TransacaoIniciada = True
    SQLCmdLivro.Execute
    ADOCn.CommitTrans
    DeletaLivro = True
    Exit Function
Exception:
    If TransacaoIniciada Then ADOCn.RollbackTrans
    MsgBox ProblemasEncontrados

End Function
Public Function CadastraLivro(Titulo As String, Editora As String, Edicao As Integer, Ano As Integer, CodAutores() As Integer, CodAssunto As Integer) As Boolean

On Error GoTo Exception
  

    Dim CodLivroNovo As Integer
    Dim TransacaoIniciada As Integer

    Dim SQLCmdLivro As New ADODB.Command
    Dim SQLCmdLivroAutor As New ADODB.Command
    Dim SQLCmdLivroAssunto As New ADODB.Command
    
    SQLCmdLivro.ActiveConnection = ADOCn
    SQLCmdLivroAutor.ActiveConnection = ADOCn
    SQLCmdLivroAssunto.ActiveConnection = ADOCn
    
    SQLCmdLivro.CommandText = "dbo.CriarLivro"
    SQLCmdLivro.CommandType = adCmdStoredProc
    
    SQLCmdLivro.Parameters.Append SQLCmdLivro.CreateParameter("@Titulo", adVarChar, adParamInput, 40, Titulo)
    SQLCmdLivro.Parameters.Append SQLCmdLivro.CreateParameter("@Editora", adVarChar, adParamInput, 40, Editora)
    SQLCmdLivro.Parameters.Append SQLCmdLivro.CreateParameter("@Edicao", adInteger, adParamInput, 4, Edicao)
    SQLCmdLivro.Parameters.Append SQLCmdLivro.CreateParameter("@AnoPublicacao", adVarChar, adParamInput, 4, Ano)
    SQLCmdLivro.Parameters.Append SQLCmdLivro.CreateParameter("@CodL", adInteger, adParamOutput, 4)
  

    
    SQLCmdLivro.ActiveConnection = ADOCn
    ADOCn.BeginTrans
    TransacaoIniciada = True
    SQLCmdLivro.Execute
    
    CodLivroNovo = SQLCmdLivro.Parameters("@CodL").Value

    For i = 0 To UBound(CodAutores)
        If CodAutores(i) = 0 Then
            
            Exit For
        End If
    
         
        SQLCmdLivroAutor.CommandText = "dbo.LivroAutorAssociar"
        SQLCmdLivroAutor.CommandType = adCmdStoredProc
        
        SQLCmdLivroAutor.Parameters.Append SQLCmdLivroAutor.CreateParameter("@CodAutor", adInteger, adParamInput, 4, CodAutores(i))
        SQLCmdLivroAutor.Parameters.Append SQLCmdLivroAutor.CreateParameter("@CodLivro", adInteger, adParamInput, 4, CodLivroNovo)
        SQLCmdLivroAutor.Execute
        
        For j = 0 To SQLCmdLivroAutor.Parameters.Count - 1
            SQLCmdLivroAutor.Parameters.Delete (0)
        Next
      
        
      
        
        
    Next
    SQLCmdLivroAssunto.CommandText = "dbo.LivroAssuntoAssociar"
    SQLCmdLivroAssunto.CommandType = adCmdStoredProc
    SQLCmdLivroAssunto.Parameters.Append SQLCmdLivroAssunto.CreateParameter("@CodAssunto", adInteger, adParamInput, 4, CodAssunto)
    SQLCmdLivroAssunto.Parameters.Append SQLCmdLivroAssunto.CreateParameter("@CodLivro", adInteger, adParamInput, 4, CodLivroNovo)
    SQLCmdLivroAssunto.Execute
 
    ADOCn.CommitTrans
    CadastraLivro = True
    Exit Function
Exception:
    If (TransacaoIniciada) Then
        ADOCn.RollbackTrans
    End If
    MsgBox MensagemCadastroComFalha, vbExclamation
    
End Function
