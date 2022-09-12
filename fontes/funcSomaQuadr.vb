Function somaQuadr(dbArr() As Double) As String
    Dim aux As String
    Dim auxb As Double
    '///////////////////////////////somar 5mm em cada lado///////////////////////////////////////////////
    For i = 0 To UBound(dbArr())
        dbArr(i) = dbArr(i) + CDbl(5)
        dbArr(i) = CInt(dbArr(i))
    Next i
    
    
        aux = CStr(dbArr(0)) & " X " & CStr(dbArr(1)) & " X " & CStr(dbArr(2))
        somaQuadr = CStr(aux)
   
    
End Function