User Function M415GRV()
    if (Inclui .or. Altera)
        if ApMsgYesNo("Manda E-mail do Or�amento para o Cliente?","Or�amento")
            U_GerOrc()
        endif
    endif
Return nil