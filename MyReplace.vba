Function MyReplace1(ByVal S As String, ByVal car1 As String, ByVal car2 As String) As String

'renvoie: remplace dans la chaine de caractere S le mot car1 par le mot car2

Dim n1 As Long 'le nombre de caracteres de S
Dim n2 As Long 'le nombre de caracteres de car1
Dim c As String
Dim p As Long 'la position du mot car1 dans S
n2 = Len(car1)
n1 = Len(S)
p = InStr(1, S, car1)
If p > 1 Then
    MyReplace1 = MyReplace1(Left(S, p - 1) & car2 & " " & Right(S, n1 - n2 - p), car1, car2)
ElseIf p = 1 Then
    MyReplace1 = car2 & Right(S, n1 - 1)
Else
    MyReplace1 = S
End If
End Function
Sub test5()
MsgBox MyReplace1("bonjour à mes freres, je veux manger mes bonbons et mes biscuits", "mes", "tes")
End Sub
