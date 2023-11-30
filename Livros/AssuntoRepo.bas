Attribute VB_Name = "AssuntoRepositorio"
Public Sub CarregaAssuntosParaComboBox(cbo As ComboBox)

On Error GoTo Excecao
    
    cbo.Clear

    SQL = "select CodAs as Codigo,Descricao as Assunto from Assunto order by Descricao "
    
    Set adors = New ADODB.Recordset
    adors.Open SQL, ADOCn, adOpenForwardOnly, adLockReadOnly
       
    If adors.EOF Then
    
    
        Exit Sub
    
    End If
    
    With cbo
                
        Do While Not adors.EOF
                      
            
            .AddItem adors("Assunto")
            .ItemData(.NewIndex) = CInt(adors("Codigo"))
            adors.MoveNext
        
        Loop
    End With
  
    adors.Close
    
    Exit Sub

Excecao:

    MsgBox "Erro ao carregar autores" & Err.Number & " - " & Err.Description, vbCritical

    Screen.MousePointer = vbDefault

End Sub
Public Function ValidarCricacaoAssunto(txtAssunto As TextBox) As Boolean
    If (Len(txtAssunto.Text) <= 2) Then Exit Function
    ValidarCricacaoAssunto = True
    
End Function
Public Function CadastraAssunto(Assunto As String) As Boolean
    
    On Error GoTo Exception
    Dim SQLCmdAssunto As New ADODB.Command
    
    SQLCmdAssunto.ActiveConnection = ADOCn

    
    SQLCmdAssunto.CommandText = "dbo.CriarAssunto"
    SQLCmdAssunto.CommandType = adCmdStoredProc
    
    SQLCmdAssunto.Parameters.Append SQLCmdAssunto.CreateParameter("@Assunto", adVarChar, adParamInput, 40, Assunto)
    

      
    SQLCmdAssunto.Execute
    CadastraAssunto = True
    Exit Function
Exception:
    MsgBox MensagemCadastroComFalha, vbExclamation

End Function
Public Function DeletaAssunto(CodigoAssunto As String) As Boolean
    Dim SQLCmdAssunto As New ADODB.Command
    On Error GoTo Exception
    SQLCmdAssunto.ActiveConnection = ADOCn
    SQLCmdAssunto.Parameters.Append SQLCmdAssunto.CreateParameter("@CodAssunto", adInteger, adParamInput, 4, CodigoAssunto)
    SQLCmdAssunto.CommandText = "dbo.DeletarAssunto"
    SQLCmdAssunto.CommandType = adCmdStoredProc
    ADOCn.BeginTrans
    TransacaoIniciada = True
    SQLCmdAssunto.Execute
    ADOCn.CommitTrans
    DeletaAssunto = True
    Exit Function
Exception:
    If TransacaoIniciada Then ADOCn.RollbackTrans
    MsgBox ProblemasEncontrados, vbExclamation

End Function
Function AssuntoDataGrid(AssuntoGrid As DataGrid) As ADODB.Recordset
    
    On Error GoTo Exception
    Set adors = New ADODB.Recordset
    FecharRecordsetDataGRid
    Dim SQL As String
    SQL = "select Codigo,Assunto from AssuntoVisao order by Assunto,Codigo"
    Set dbCommand = New ADODB.Command
    Set dbCommand.ActiveConnection = ADOCn
    dbCommand.CommandType = adCmdText
    dbCommand.CommandText = SQL
    adorsDataGrid.CursorLocation = adUseClient
    adorsDataGrid.Open SQL, ADOCn, adOpenDynamic, adLockPessimistic
    
    Set AssuntoGrid.DataSource = adorsDataGrid
    Set AssuntoDataGrid = adorsDataGrid
    Exit Function
Exception:
    MsgBox ProblemasEncontrados, vbExclamation
End Function
