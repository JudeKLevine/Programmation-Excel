Function Exp_x(n As Integer, x As Integer) As Double
    'renvoie une valeur approchée de exp(x) en fonction de n
    
    Dim res As Double
    Dim i As Integer
    
    res = 1
    If n = 0 Then Exp_x = 1: Exit Function
    For i = 1 To n
        res = res + (x ^ i) / (Fact(i))
        Next
    Exp_x = res
        
End Function
