Function convDoub(ByVal strValor As String) As Double
    convDoub = 0#
    Dim char As String
    Dim isNegat As Boolean
    Dim digito$
    ParseNumeber = 0#
    isNegat = False
    For i = 1 To Len(strValor)
'////////////////////////////procurando negativo//////////////////////////////////////////
        char = Mid(strValor, i, 1)
        If (char = "-") Then
            digito = digito & char
'////////////////////////////procurando negativo//////////////////////////////////////////
'//////////////////////////procurando numeros 0--9////////////////////////////////////////
        ElseIf (char >= "0" And char <= "9") Then
            digito = digito & char
'///////////////////////////////troca ponto///////////////////////////////////////////////
        ElseIf (char = ".") Then
            
            digito = digito & ","
        End If
    Next i
'///////////////////////////////troca ponto///////////////////////////////////////////////
  
    convDoub = CDbl(digito)
End Function
