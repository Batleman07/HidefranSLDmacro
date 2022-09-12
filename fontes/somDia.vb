Function somaDia(dbArr() As Double) As String
    Dim aux As String
    '///////////////////////////////testa qual o diâmetro e soma 3mm e 5mm no comprimento ///////////////////////////////////////////////
    If (dbArr(0) = dbArr(1)) Then
        dbArr(1) = dbArr(1) + CDbl(3)
        dbArr(2) = dbArr(2) + CDbl(5)
        dbArr(1) = CInt(dbArr(1))
        dbArr(2) = CInt(dbArr(2))
            aux = "Ø" & CStr(dbArr(1)) & " X " & CStr(dbArr(2))
                somaDia = CStr(aux)
    ElseIf (dbArr(2) = dbArr(1)) Then
        dbArr(1) = dbArr(1) + CDbl(3)
        dbArr(0) = dbArr(0) + CDbl(5)
        dbArr(1) = CInt(dbArr(1))
        dbArr(0) = CInt(dbArr(0))
            aux = "Ø" & CStr(dbArr(1)) & " X " & CStr(dbArr(0))
                somaDia = CStr(aux)
    Else
        somaDia = ""
    End If
End Function
