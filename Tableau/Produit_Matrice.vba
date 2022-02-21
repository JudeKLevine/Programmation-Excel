Function Produit_Matrice(Mat1() As Integer, Mat2() As Integer) As Integer()

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim val As Integer
    
    If UBound(Mat1, 2) <> UBound(Mat2, 1) Or LBound(Mat1, 2) <> LBound(Mat2, 1) Then
        Err.Raise 5050, , "PRODUIT PAS POSSIBLE"
    Else
        Dim TABLEAU() As Integer
        ReDim TABLEAU(LBound(Mat1, 1) To UBound(Mat1), LBound(Mat2, 2) To UBound(Mat2, 2))
    
        For i = 0 To UBound(Mat1, 1)
            For j = 0 To UBound(Mat2, 2)
                val = 0
                For k = 0 To UBound(Mat2, 1)
                    val = val + Mat1(i, k) * Mat2(k, j)
                Next
            TABLEAU(i, j) = val
            Next
        Next
    End If
    Produit_Matrice = TABLEAU
    
End Function
