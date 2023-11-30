Attribute VB_Name = "LivroService"
Public Function PreparaCadastroLivro(Titulo As String, Editora As String, Edicao As String, Ano As String, listAutor As ListBox, cboAssunto As ComboBox, LivroGrid As DataGrid) As Boolean
    
    Dim Autores(0 To 50) As Integer
    With listAutor
        For i = 0 To .ListCount - 1
            Autores(i) = .ItemData(i)
        Next
    End With
    Dim CodAssunto As Integer
    With cboAssunto
        CodAssunto = .ItemData(.ListIndex)
    End With
    If (CadastraLivro(Titulo, Editora, CInt(Edicao), CInt(Ano), Autores, CodAssunto) = True) Then
        LivrosDataGrid LivroGrid
        FormatGridLivros LivroGrid
        LivroGrid.Refresh
        MsgBox MensagemCadastroSucesso, vbInformation
       
        PreparaCadastroLivro = True
    Else
        MsgBox MensagemCadastroComFalha, vbExclamation
    End If
End Function
Public Sub FormatGridLivros(GridLivros As DataGrid)
If GridLivros.Columns.Count = 7 Then GridLivros.Columns.Item(6).Width = 4500

End Sub
Public Function LinhasCabRelatorio() As String
LinhasCabRelatorio = String(18, " ") + "Codigo" + String(17, " ") + "Livro" + String(40, " ") + "Autor"
End Function
Public Function CarregaRelatorioLivros() As String
Dim bBuffer As String

Dim adorsReportGrid As ADODB.Recordset
    Set adorsReportGrid = LivrosReport
    Dim LinhasRelatorio As String
    Dim Titulo As String
    Dim Autor As String
    Dim AnoPub As String
    
    While Not adorsReportGrid.EOF
    
        Titulo = Left(Trim(adorsReportGrid!Titulo) + String(30, " "), 30)
        Autor = Left(IIf(IsNull(adorsReportGrid!Autor), String(30, " "), Trim(adorsReportGrid!Autor) + String(30, " ")), 30)
        Assunto = Left(IIf(IsNull(adorsReportGrid!Assunto), String(30, " "), Trim(adorsReportGrid!Assunto) + String(30, " ")), 30)
        AnoPub = Str(adorsReportGrid!AnoPublicacao)
        
        LinhasRelatorio = LinhasRelatorio + String(20, " ") + Format(Str(adorsReportGrid!Codigo), "00000") + String(20, " ") + Titulo + String(20, " ") + Autor + vbCrLf
        
        adorsDataGrid.MoveNext
    Wend
    
bBuffer = vbCrLf + LinhasRelatorio
CarregaRelatorioLivros = bBuffer
End Function
Public Sub AssociaAutoraoLivro(cboAutor As ComboBox, ListAutores As ListBox)

    Dim Autor As String
    Dim CodAutor As String
    With cboAutor
        CodAutor = .ItemData(.ListIndex)
        Autor = .List(.ListIndex)
        If (ExisteAutor(CodAutor, ListAutores)) Then
            MsgBox MensagemAssociacaoRealizada, vbExclamation
        Else
            With ListAutores
                .AddItem Autor
                .ItemData(.NewIndex) = CodAutor
               
            End With
        End If
    End With
End Sub
Function LivrosDataGridDelete(LivroGrid As DataGrid) As Boolean
    Dim i As Integer
    i = LivroGrid.Row
    adorsDataGrid.MoveFirst
    adorsDataGrid.Move i
    If (DeletaLivro(adorsDataGrid("Codigo"))) Then
        LivrosDataGrid LivroGrid
        LivroGrid.Refresh
        MsgBox MensagemDelecaoSucesso, vbInformation
    Else
        MsgBox ProblemasEncontrados, vbExclamation
    End If
   
End Function
