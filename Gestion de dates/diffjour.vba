Function diffjour(ByVal D_choose As Date) As Integer

    Dim D_Now As Date           'Utiliser le bon format pour date D_choose
                                'il faut mettre de # avant et apres la date
    D_Now = Now                 'Exemple : #jj/mm/aaaa#
    diffjour = D_Now - D_choose

End Function
