User Function M415GRV()
    if (Inclui .or. Altera)
        if ApMsgYesNo("Manda E-mail do Orçamento para o Cliente?","Orçamento")
            U_GerOrc()
        endif
    endif
Return nil