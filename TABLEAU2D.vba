'--------------------------------------------
'----------------TABLEAU 2D------------------
'--------------------------------------------
Function getMatrice(ByVal n As Integer, ByVal m As Integer, _
                    ByVal a As Integer, ByVal b As Integer) As Integer()
Dim i As Integer
Dim j As Integer
Dim mat() As Integer
ReDim mat(n, m)
For i = 0 To n
    For j = 0 To m
    mat(i, j) = WorksheetFunction.RandBetween(a, b)
    Next j
Next i
getMatrice = mat
End Function
Sub test2()
Dim n As Integer
Dim m As Integer
Dim a As Integer
Dim b As Integer
Dim i As Integer
Dim j As Integer
Dim chaine As String
Dim mat() As Integer
n = Application.InputBox("entrer le nombre de ligne de la matrice", Type:=1)
m = Application.InputBox("entrer le nombre de colonne de la matrice", Type:=1)
a = Application.InputBox("entrer la borne inf", Type:=1)
b = Application.InputBox("entrer la borne sup", Type:=1)
mat = getMatrice(n, m, a, b)
chaine = ""
For i = 0 To n
    For j = 0 To m
    chaine = chaine & mat(i, j) '& vbNewLine
    Next j
Next i
MsgBox chaine
End Sub

Function somme2matrice(M1() As Integer, M2() As Integer) As Integer()
'hypothese: les 2 matrices doivent etre de meme taille
Dim mat() As Integer
ReDim mat(LBound(M1), UBound(M1))
Dim i As Integer
Dim j As Integer
For i = LBound(M1) To UBound(M1)
    For j = LBound(M1) To UBound(M2)
    mat(i, j) = M1(i, j) + M2(i, j)
    Next j
Next i
somme2matrice = mat
End Function
Sub test3()
Dim M1() As Integer
Dim M2() As Integer
Dim mat() As Integer
Dim i As Integer
Dim j As Integer
Dim chaine As String
M1 = getMatrice(2, 2, 1, 9)
M2 = getMatrice(2, 2, 1, 9)
mat = somme2matrice(M1, M2)
chaine = ""
For i = LBound(M1) To UBound(M2)
    For j = LBound(M1) To UBound(M1)
    chaine = chaine & mat(i, j)
    Next j
Next i
MsgBox chaine
    
End Sub
