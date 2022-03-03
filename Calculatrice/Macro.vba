Sub Zero()
    Cells(3, 3).Value = Cells(3, 3).Value & 0
End Sub
Sub Un()
    Cells(3, 3).Value = Cells(3, 3).Value & 1
End Sub
Sub Deux()
    Cells(3, 3).Value = Cells(3, 3).Value & 2
End Sub
Sub Trois()
    Cells(3, 3).Value = Cells(3, 3).Value & 3
End Sub
Sub Quatre()
    Cells(3, 3).Value = Cells(3, 3).Value & 4
End Sub
Sub Cinq()
    Cells(3, 3).Value = Cells(3, 3).Value & 5
End Sub
Sub Six()
    Cells(3, 3).Value = Cells(3, 3).Value & 6
End Sub
Sub Sept()
    Cells(3, 3).Value = Cells(3, 3).Value & 7
End Sub
Sub Huit()
    Cells(3, 3).Value = Cells(3, 3).Value & 8
End Sub
Sub Neuf()
    Cells(3, 3).Value = Cells(3, 3).Value & 9
End Sub
Sub Plus()
    Cells(3, 3).Value = Cells(3, 3).Value & "+"
End Sub
Sub Moins()
    Cells(3, 3).Value = Cells(3, 3).Value & "-"
End Sub
Sub Fois()
    Cells(3, 3).Value = Cells(3, 3).Value & "*"
End Sub
Sub Diviser()
    Cells(3, 3).Value = Cells(3, 3).Value & "/"
End Sub
Sub Point()
    Cells(3, 3).Value = Cells(3, 3).Value & "."
End Sub
Sub AC()
    Cells(3, 3).Value = ""
    Cells(5, 4).Value = ""
End Sub
Sub DEL()

    Dim chaine As String: chaine = Cells(3, 3).Value
    Dim taille As Integer
    taille = Len(chaine)
    If Len(chaine) > 0 Then
        chaine = Mid(chaine, 1, Len(chaine) - 1)
    End If
    
    Cells(3, 3).Value = chaine
    
End Sub
Sub Pourcent()

    Cells(5, 4).NumberFormat = "+0%; -0%"

End Sub
Sub Ce()

    Cells(5, 4).NumberFormat = "+0; -0"

End Sub
Sub Egale()
    Dim Con As String
    Con = "=" & Cells(3, 3)
    If InStr(1, Con, ",") > 0 Then
        Cells(5, 4).Value = Mid(Con, 2, Len(Con) - 1)
    Else
        Cells(5, 4).Value = Con
    End If
    Cells(3, 3).Value = ""
End Sub


