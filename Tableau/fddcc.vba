Debug.Print LBound(Mat1, 1)
    If LBound(Mat1, 1) <> LBound(Mat2, 1) Or UBound(Mat1, 1) <> UBound(Mat2, 1) Or _
       LBound(Mat1, 2) <> LBound(Mat2, 2) Or UBound(Mat1, 2) <> UBound(Mat2, 2) Then
       Err.Raise 5050, , "PAS POSSIBLE"
    Else
        ReDim SOMME(LBound(Mat1, 1) To UBound(Mat1, 1), LBound(Mat1, 2) To UBound(Mat1, 2))
        For i = LBound(Mat1, 1) To UBound(Mat1, 1)
            For j = LBound(Mat1, 2) To UBound(Mat1, 2)
                SOMME(i, j) = 1
            Next
        Next
    End If
    Debug.Print SOMME(1, 1)
    Somme2Matrice = 1
