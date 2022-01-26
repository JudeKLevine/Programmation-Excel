Function SommeEntier(ByVal N As Integer) As Integer

    Dim Sum As Integer
    Dim i As Integer
    
    ' on pourra traiter le cas de N = 100
    For i = 0 To N
        Sum = Sum + i
    Next
    SommeEntier = Sum

End Function
