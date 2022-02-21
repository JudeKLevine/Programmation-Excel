Function getTable(ByVal last As Integer, _
                  ByVal min As Integer, _
                  ByVal max As Integer) As Integer()

    'Option Base 0 pas necessaire
    
    Dim TABLEAU() As Integer
    Dim i As Integer
    ReDim TABLEAU(last)
    
    For i = 0 To last
        TABLEAU(i) = Int((max - min + 1) * Rnd + min)
    Next
   
    getTable = TABLEAU
    
End Function
