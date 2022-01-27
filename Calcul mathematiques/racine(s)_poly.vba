Function racineP(ByVal neme As Integer, _
                ByVal a As Double, _
                ByVal b As Double, _
                ByVal c As Double) As Double
                
    Dim Nb_racine As Integer
    Dim Dlt As Double
    
    Dlt = Delta(a, b, c)
    Nb_racine = Nbracine(Dlt)
    Select Case Nb_racine
        Case 0
            MsgBox "PAS DE SOLUTION"
        Case 1
            racineP = -b / (2 * a)
        Case 2
            Select Case neme
                Case 1
                    racineP = (-b - Sqr(Dlt)) / (2 * a)
                Case 2
                    racineP = (-b + Sqr(Dlt)) / (2 * a)
                Case Is > 2
                    Debug.Print "PAS PLUS DE DEUX SOLUTIONS"
            End Select
    End Select
 
End Function
