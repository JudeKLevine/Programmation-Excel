Function Transpose_Matrice(Mat() As Double) As Double()

    Dim Matrice() As Double
    Dim i As Integer
    Dim j As Integer
    
    ReDim Matrice(LBound(Mat, 2) To UBound(Mat, 2), LBound(Mat, 1) To UBound(Mat, 1))
    
    For i = LBound(Mat, 1) To UBound(Mat, 1)
        For j = LBound(Mat, 2) To UBound(Mat, 2)
            Matrice(j, i) = Mat(i, j)
        Next
    Next
    
    Transpose_Matrice = Matrice
    
End Function
