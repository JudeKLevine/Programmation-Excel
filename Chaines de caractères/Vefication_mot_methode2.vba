Function VerifiMots(ByVal S As String, ByVal mot1 As String, ByVal mot2 As String) As String

    Dim rang1 As Integer
    Dim rang2 As Integer
    
    rang1 = InStr(1, S, mot1)
    rang2 = InStr(1, S, mot2)
    
    If rang1 > 0 Then
        VerifiMots = mot1 & " " & "present au rang" & " " & rang1
    Else
        VerifiMots = mot1 & " " & "absent"
    End If
    
    If rang2 > 0 Then
        VerifiMots = VerifiMots & " " & ";" & " " & mot2 & " " & "present au rang" & " " & rang1
    Else
        VerifiMots = VerifiMots & " " & ";" & " " & mot2 & " " & "absent"
    End If
    
End Function
