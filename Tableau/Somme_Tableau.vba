Function Somme2Tab(T1() As Integer, T2() As Integer) As Integer()

    Dim REPONSE() As Integer
    Dim i As Integer
    ReDim REPONSE(LBound(T1) To UBound(T1))
    
    If UBound(T1) <> UBound(T2) And LBound(T1) <> LBound(T2) Then
        Err.Raise 5010, , "les tableau doivent avoir les memes indices"
    Else
        For i = LBound(T2) To UBound(T2)
            REPONSE(i) = T1(i) + T2(i)
        Next
    End If
    Somme2Tab = REPONSE
End Function
