Function Myreplace(ByVal S As String, ByVal car1 As String, ByVal car2 As String) As String
    
    Do While InStr(1, S, car1) > 0
        S = Replace(S, car1, car2)
    Loop
    Myreplace = S
   
End Function
