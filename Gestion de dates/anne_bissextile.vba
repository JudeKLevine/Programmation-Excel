Function bissextile(years As Integer) As Boolean

    If years < 0 Then
        Debug.Print "IMPOSSIBLE"
        Exit Function
    ElseIf years Mod 4 = 0 Then
        bissextile = True
    End If

End Function
