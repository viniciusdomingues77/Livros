Attribute VB_Name = "DBConect"
Global ADOCn As New ADODB.Connection
Global BancoConectado As Boolean
Sub ConectaBancoLivros()
    Screen.MousePointer = vbHourglass
  
    On Error GoTo Excecao
    
    sConnectionString = "DRIVER={Sql Server};SERVER=DESKTOP-PM4320J;DATABASE=Livros;UID=user_livro;PWD=123456;"

    With ADOCn
       .ConnectionString = sConnectionString
       .ConnectionTimeout = DataBaseTimeOut
       .CommandTimeout = CommandTimeout
       .Open
    End With
    
    If ADOCn.State = 1 Then BancoConectado = True
      
    
    Screen.MousePointer = vbArrow
    Exit Sub
Excecao:
    Screen.MousePointer = vbArrow
    MsgBox "Não foi possível conecatar o banco " + Err.Number, vbCritical
    
End Sub
Public Sub DesconectaBancoLivros()
 On Error GoTo Excecao
  
  If ADOCn.State = 1 Then
    BancoConectado = True
    ADOCn.Close
  End If
 
  Exit Sub
Excecao:
    MsgBox "Não foi possível desconectar o banco " + Err.Number, vbCritical
End Sub
Public Function BancoLivrosConectado() As Boolean
    If BancoConectado Then
        BancoLivrosConectado = True
    Else
        ConectaBancoLivros
        If BancoConectado Then
            BancoLivrosConectado = True
        End If
    End If
End Function
