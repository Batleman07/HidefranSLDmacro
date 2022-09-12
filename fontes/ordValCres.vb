Function ordValCresc(x As Double, y As Double, z As Double) As Double()
'////////////////////////////variaveis//////////////////////////////////////////
    Dim dbArr(4) As Double
'////////////////////////////variaveis//////////////////////////////////////////

'////////////////////////////ordenar//////////////////////////////////////////
    For i = 0 To UBound(dbArr) - 1
        dbArr(0) = x
        dbArr(1) = y
        dbArr(2) = z
    Next i
    For i = 0 To UBound(dbArr) - 1
        If (dbArr(i) > dbArr(i + 1)) Then
            dbArr(3) = dbArr(i + 1)
            dbArr(i + 1) = dbArr(i)
            dbArr(i) = dbArr(3)
        End If
    Next i
'////////////////////////////ordenar//////////////////////////////////////////

    ordValCresc = dbArr()

End Function
