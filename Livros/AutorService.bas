Attribute VB_Name = "AutorService"
Global CadastroRapidodeAutor As Boolean
Public Function ExisteAutor(CodAutor As String, ListAutores As ListBox) As Boolean
    With ListAutores
        For i = 0 To .ListCount - 1
            If .ItemData(i) = CodAutor Then
                ExisteAutor = True
                Exit Function
            End If
        Next
    End With
End Function
Public Function ValidarCricacaoAutor(txtNomeAutor As TextBox) As Boolean
    If (Len(txtNomeAutor.Text) <= 2) Then Exit Function
    ValidarCricacaoAutor = True
    
End Function
Public Function PreparaCriacaoAutor(NomeAutor As String, AutorGrid As DataGrid) As Boolean
    If (CadastraAutor(NomeAutor)) Then
        PreparaCriacaoAutor = True
        AutorDataGrid AutorGrid
        FormatGridAutor AutorGrid
        MsgBox MensagemCadastroSucesso, vbInformation
    End If
End Function
Function AutorDataGridDelete(AutorGrid As DataGrid) As Boolean
    Dim i As Integer
    i = AutorGrid.Row
    adorsDataGrid.MoveFirst
    adorsDataGrid.Move i
    If (DeletaAutor(adorsDataGrid("Codigo"))) Then
        AutorDataGrid AutorGrid
        FormatGridAutor AutorGrid
        AutorGrid.Refresh
        MsgBox MensagemDelecaoSucesso, vbInformation
    End If
    
End Function
Function ECadastroRapidoAutor() As Boolean
    ECadastroRapidoAutor = CadastroRapidodeAutor
End Function
Function VerificaExisteAutorSelecionado(cbo As ComboBox) As Boolean
    If (cbo.ListIndex = -1) Then Exit Function
    VerificaExisteAutorSelecionado = True
    
End Function
Public Sub FormatGridAutor(GridAutor As DataGrid)
If GridAutor.Columns.Count = 2 Then GridAutor.Columns.Item(1).Width = 4500

End Sub
