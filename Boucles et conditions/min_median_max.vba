Function neme(ByVal rang As Integer, ByVal arg1 As Double, ByVal arg2 As Double, ByVal arg3 As Double) As Double
'renvoie selon la valeur de n le min, le max ou le median

    If rang < 1 Or rang > 3 Then Debug.Print "ERREUR": Exit Function
    Select Case rang
        Case 1
            neme = mini3(arg1, arg2, arg3)
        Case 2
            neme = Median(arg1, arg2, arg3)
        Case 3
            neme = maxi3(arg1, arg2, arg3)
    End Select

End Function
