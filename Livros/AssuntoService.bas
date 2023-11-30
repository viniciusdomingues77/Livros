Attribute VB_Name = "AssuntoService"
Global CadastroRapidodeAssunto As Boolean
Public Sub FormatGridAssuntos(GridAssuntos As DataGrid)
If GridAssuntos.Columns.Count = 2 Then GridAssuntos.Columns.Item(1).Width = 4500

End Sub
Public Function PreparaCriacaoAssunto(Assunto As String, AssuntoGrid As DataGrid) As Boolean
    If (CadastraAssunto(Assunto)) Then
        PreparaCriacaoAssunto = True
        AssuntoDataGrid AssuntoGrid
        MsgBox MensagemCadastroSucesso, vbInformation
    End If
End Function
Function ECadastroRapidoAssunto() As Boolean
    ECadastroRapidoAssunto = CadastroRapidodeAssunto
End Function
Function AssuntoDataGridDelete(AssuntoGrid As DataGrid) As Boolean
    Dim i As Integer
    i = AssuntoGrid.Row
    adorsDataGrid.MoveFirst
    adorsDataGrid.Move i
    If (DeletaAssunto(adorsDataGrid("Codigo"))) Then
        AssuntoDataGrid AssuntoGrid
        AssuntoGrid.Refresh
        MsgBox MensagemDelecaoSucesso, vbInformation
    End If
   
End Function
