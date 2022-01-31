'# Retrouver ce algorithme sur wikipedia
'# https://fr.wikipedia.org/wiki/Calcul_de_la_date_de_P%C3%A2ques

Function Dimache_Paques(ByVal annee As Integer) As Date
    
    Dim n As Integer
    Dim c As Integer
    Dim u As Integer
    Dim q As Integer
    Dim e As Integer
    Dim L As Integer
    Dim mois As Integer
    Dim jour As Integer
    
    n = annee Mod 19
    c = annee \ 100
    u = annee Mod 100
    q = (c - ((c + 8) \ 25) + 1) \ 3
    e = (19 * n + c - c \ 4 - q + 15) Mod 30
    L = (2 * (c Mod 4) + 2 * (u \ 4) - e - (u Mod 4) + 32) Mod 7
    mois = (e + L - 7 * ((n + 11 * e + 22 * L) \ 451) + 114) \ 31
    jour = (e + L - 7 * ((n + 11 * e + 22 * L) \ 451) + 114) Mod 31
   
   Debug.Print jour
   Debug.Print mois
   Dimache_Paques = DateSerial(annee, mois, jour + 1)
   
End Function


Sub test()
   
    Debug.Print Dimache_Paques(2020)
 
End Sub

    
   
   
   
   
   
   
