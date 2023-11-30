Attribute VB_Name = "Funcionalidades"

Function ENumero(KeyAscii As Integer) As Boolean
If (KeyAscii >= Asc("0")) And (KeyAscii <= Asc("9")) Then ENumero = True
End Function
