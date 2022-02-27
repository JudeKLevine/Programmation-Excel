Option Explicit
Option Base 1

Type enseignement
    id As Integer
    nom As String
End Type

'passage d’un tableau ByRef obligatoire
Public Sub initTab(ByRef tabMatiere() As enseignement)
'initialise le tableau tabMatiere avec 7 elements

ReDim tabMatiere(7)

'tabMatiere(1).id = 25: tabMatiere(1).nom = "Optimisation non lineaire"

With tabMatiere(1)
.id = 25
.nom = "Optimisation non lineaire"
End With
With tabMatiere(2)
    .id = 36
    .nom = "Programmation"
End With
With tabMatiere(3)
    .id = 58
    .nom = "Gestion financiere"
End With
With tabMatiere(4)
    .id = 11
    .nom = "Probabilites"
End With
With tabMatiere(5)
    .id = 7
    .nom = "Statistiques"
End With
With tabMatiere(6)
    .id = 61
    .nom = "Bases de Donnees"
End With
With tabMatiere(7)
    .id = 87
    .nom = "Decision dans l’incertain"
End With

End Sub
Function tabAsString(ByRef tabMatiere() As enseignement) As String

    Dim chaine As String
    Dim i As Integer
    
    For i = LBound(tabMatiere) To UBound(tabMatiere)
        chaine = chaine & i & "." & " " & tabMatiere(i).nom & vbNewLine
    Next
    
    tabAsString = chaine
    
End Function

Sub principal()

    Dim i As Integer
    Dim indice As Integer
    Dim msgErreur As String: msgErreur = ""
    Dim listeMatiere As String: listeMatiere = ""
    Dim TABLE() As enseignement
    ReDim TABLE(7)
    
    initTab TABLE
    listeMatiere = tabAsString(TABLE)
    
    Do
    indice = Application.InputBox(msgErreur & "Liste des matieres : " _
            & vbNewLine & listeMatiere & vbNewLine & _
            "Donner un index pour choisir une matiere : ", Type:=1)
    msgErreur = "Erreur d'indice" & vbNewLine
    Loop Until LBound(TABLE) <= indice And indice <= UBound(TABLE) Or indice = 0
    MsgBox "l'identifiant de la matiere est : " & TABLE(indice).id
    
End Sub



