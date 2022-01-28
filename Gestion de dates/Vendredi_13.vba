Function Vendredi_13() As Integer

    Dim mois As Integer
    Dim annee As Integer
    Dim result As Integer
    result = 0
    
    Dim date_string As String
    Dim date_test As Date
    Dim da As Date
    
    For annee = 2001 To 2100
        For mois = 1 To 12
            date_string = 13 & "/" & mois & "/" & annee
            date_test = CDate(date_string)
            If Weekday(date_test, 2) = 5 Then
                result = result + 1
            End If
        Next
    Next
    
    Vendredi_13 = result
    
End Function

Sub test()
    Debug.Print Vendredi_13()
End Sub

