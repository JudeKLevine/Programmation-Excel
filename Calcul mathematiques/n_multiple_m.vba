Function multiple(ByVal int1 As Integer, ByVal int2 As Integer) As Boolean

    If int1 Mod int2 = 0 Or int2 Mod int1 = 0 Then
        multiple = True
    End If

End Function
