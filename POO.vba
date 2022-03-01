Option Explicit
Option Base 1


Type parametre
    min As Integer
    max As Integer
    operateur As String
End Type
    
Function recupereparametre() As parametre

    Dim PARA() As parametre
    Dim diff As Integer
    ReDim PARA(1)
    
    diff = Cells(3, 2).Value
    
    Select Case diff
        Case 0
            PARA.min = 1
            PARA.max = 10
            PARA.operateur = "+"
        Case 1
            PARA.max = 10
            PARA.min = 1
            PARA.operateur = "+*"
        Case 2
            PARA.max = 10
            PARA.min = 1
            PARA.operateur = "*/"
        Case 3
            PARA.max = 50
            PARA.min = 1
            PARA.operateur = "*/"
        Case Is > 3, Is < 0
            PARA.max = -1
            PARA.min = -1
            PARA.operateur = ""
        End Select
        
        recupereparametre = PARA
        
End Function
Function OperateurAleatoire(ByVal operateur As String) As String

    Dim rand As Integer
    rand = WorksheetFunction.RandBetween(1, Len(operateur))
    OperateurAleatoire = Mid(operateur, rand, 1)
End Function

Sub test()

    Dim Nbr_operation As Integer
    Dim diff As Integer
    Dim erreur As String
    Dim i As Integer
    Dim j As Integer
    Dim PARA() As parametre
    
   
    
End Sub
















Function Max_dif(ByVal dif As Integer) As Integer

    Dim max As Integer
    
    Select Case dif
        Case 0, 1, 2
            max = 10
        Case 3
            max = 50
        End Select
    Max_dif = max
End Function

Function Operation(ByVal dif As Integer, ByVal int1 As Integer, ByVal int2 As Integer) As Double
    
    Dim resultat As Double
    Dim rand As Integer
    Dim operateur As String
    
    Select Case dif
        Case 0
            resultat = int1 + int2
        Case 1
            rand = WorksheetFunction.RandBetween(0, 1)
            resultat = rand * (int1 + int2) + (1 - rand) * (int1 * int2)
        Case 2, 3
            rand = WorksheetFunction.RandBetween(0, 1)
            resultat = rand * (int1 / int2) + (1 - rand) * (int1 * int2)
        End Select
    
        Operation = resultat

End Function
Sub GenereCalclul()

    Dim Nbr_operation As Integer
    Dim diff As Integer
    Dim erreur As String
    Dim i As Integer
    Dim j As Integer
    Dim PARA() As parametre
    
    
    diff = Cells(3, 2).Value
    Nbr_operation = Cells(2, 2).Value
    
    Debug.Print diff
    Debug.Print Nbr_operation
    
    '''''''''''''''''''''''''''''''''''''''''''
    'If Diff < 0 Or Diff > 4 Then
    '    Err.Raise 5050, , "DEPASSE"
    'End If
    ''''''''''''''''''''''''''''''''''''''''''

    For i = 1 To Nbr_operation
        Cells(i + 4, 1).Value = WorksheetFunction.RandBetween(1, Max_dif(diff))
        Cells(i + 4, 3).Value = WorksheetFunction.RandBetween(1, Max_dif(diff))
        Cells(i + 4, 4).Value = "="
        Cells(i + 4, 5).Value = Operation(diff, Cells(i + 4, 1).Value, Cells(i + 4, 3).Value)
    Next

    
End Sub
