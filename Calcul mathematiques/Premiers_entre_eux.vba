Function premierentreEux(ByVal int1 As Integer, ByVal int2 As Integer) As Boolean

    If pgcd(int1, int2) = 1 Then
        premierentreEux = True
    Else
        premierentreEux = False
    End If

End Function
