Function pgcd(ByVal int1 As Integer, ByVal int2 As Integer) As Integer

    If int2 = int1 Then
        pgcd = int1
        Exit Function
    ElseIf int1 > int2 Then
        pgcd = pgcd(int1 - int2, int2)
    Else
        pgcd = pgcd(int1, int2 - int1)
End If

End Function
