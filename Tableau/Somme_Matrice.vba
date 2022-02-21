Function Somme2Matrice(Mat1() As Integer, Mat2() As Integer) As Integer()
    
    Dim SOMME() As Integer
    Dim i As Integer
    Dim j As Integer
    
    If LBound(Mat1, 1) <> LBound(Mat2, 1) Or UBound(Mat1, 1) <> UBound(Mat2, 1) Or _
       LBound(Mat1, 2) <> LBound(Mat2, 2) Or UBound(Mat1, 2) <> UBound(Mat2, 2) Then
       Err.Raise 5050, , "PAS PDE MEME DIMENSION"
    Else
        ReDim SOMME(LBound(Mat1, 1) To UBound(Mat1, 1), LBound(Mat1, 2) To UBound(Mat1, 2))
        For i = LBound(Mat1, 1) To UBound(Mat1, 1)
            For j = LBound(Mat1, 2) To UBound(Mat1, 2)
                SOMME(i, j) = Mat1(i, j) + Mat2(i, j)
            Next
        Next
    End If

    Somme2Matrice = SOMME
    
End Function
