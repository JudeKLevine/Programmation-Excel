Function lendemain(jour As Integer, mois As Integer, annee As Integer) As String

    Dim N_jour As Integer
    Dim N_mois As Integer
    Dim N_annee As Integer

    'test sur les jours, les mois et les années
    If jour = nombreJours(mois, annee) And mois < 12 Then
        N_jour = 1
        N_mois = mois
        N_annee = annee
    ElseIf jour = nombreJours(mois, annee) And mois = 12 Then
        N_jour = 1
        N_mois = 1
        N_annee = annee + 1
    Else
        N_jour = jour + 1
        N_mois = mois
        N_annee = annee
    End If
    lendemain = "(" & N_jour & "," & N_mois & "," & N_annee & ")"

End Function
