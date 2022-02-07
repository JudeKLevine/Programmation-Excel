Function Myreplace(ByVal S As String, ByVal car1 As String, ByVal car2 As String) As String
    
    Do While InStr(1, S, car1) > 0
        S = Left(S, InStr(1, S, car1) - 1) & car2 & Right(S, Len(S) - InStr(1, S, car1) + 1 - Len(car1))
    Loop
    
    Myreplace = S
   
End Function
