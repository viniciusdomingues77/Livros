Attribute VB_Name = "AutorRepositorio"
Public Sub CarregaAutoresParaComboBox(cbo As ComboBox)

On Error GoTo Excecao
    
    cbo.Clear

    SQL = "select CodAu as Codigo,Nome as Autor from Autor order by Nome "
    
    Set adors = New ADODB.Recordset
    adors.Open SQL, ADOCn, adOpenForwardOnly, adLockReadOnly
       
    If adors.EOF Then
    
        
        Exit Sub
    
    End If
    
    With cbo
                
        Do While Not adors.EOF
                        
            .AddItem adors("Autor")
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
Public Function DeletaAutor(CodigoAutor As String) As Boolean
    Dim SQLCmdAutor As New ADODB.Command
    On Error GoTo Exception
    SQLCmdAutor.ActiveConnection = ADOCn
    SQLCmdAutor.Parameters.Append SQLCmdAutor.CreateParameter("@CodAutor", adInteger, adParamInput, 4, CodigoAutor)
    SQLCmdAutor.CommandText = "dbo.DeletarAutor"
    SQLCmdAutor.CommandType = adCmdStoredProc
    ADOCn.BeginTrans
    TransacaoIniciada = True
    SQLCmdAutor.Execute
    ADOCn.CommitTrans
    DeletaAutor = True
    Exit Function
Exception:
    If TransacaoIniciada Then ADOCn.RollbackTrans
    MsgBox ProblemasEncontrados, vbExclamation

End Function
Function AutorDataGrid(AutorGrid As DataGrid) As ADODB.Recordset
    
    On Error GoTo Exception
    Set adors = New ADODB.Recordset
    FecharRecordsetDataGRid
    Dim SQL As String
    SQL = "select Codigo,Autor from AutorVisao order by Autor,Codigo"
    Set dbCommand = New ADODB.Command
    Set dbCommand.ActiveConnection = ADOCn
    dbCommand.CommandType = adCmdText
    dbCommand.CommandText = SQL
    adorsDataGrid.CursorLocation = adUseClient
    adorsDataGrid.Open SQL, ADOCn, adOpenDynamic, adLockPessimistic
    
    Set AutorGrid.DataSource = adorsDataGrid
    Set AutorDataGrid = adorsDataGrid
    Exit Function
Exception:
    MsgBox ProblemasEncontrados, vbExclamation
End Function
Public Function CadastraAutor(NomeAutor As String) As Boolean
    
    On Error GoTo Exception
    Dim SQLCmdAutor As New ADODB.Command
    
    SQLCmdAutor.ActiveConnection = ADOCn

    
    SQLCmdAutor.CommandText = "dbo.CriarAutor"
    SQLCmdAutor.CommandType = adCmdStoredProc
    
    SQLCmdAutor.Parameters.Append SQLCmdAutor.CreateParameter("@NomeAutor", adVarChar, adParamInput, 40, NomeAutor)
    

      
    SQLCmdAutor.Execute
    CadastraAutor = True
    Exit Function
Exception:
    MsgBox MensagemCadastroComFalha, vbExclamation

End Function
