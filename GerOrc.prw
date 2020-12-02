#Include "Protheus.ch"
#Include "RESTFUL.ch"
#Include "tbiconn.ch"

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³GERORC    ºAutor  ³Tiago Malta      º Data ³     31/07/17   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ função para gerar o html com os dados do orçamento         º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                         º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function GerOrc()

	Local oHtml
	Local cHtml    := ""
	Local cTotal   := ""
	Local cItens   := ""
	Local cAssunto := "Orçamento: "
	Local cBody	   := "Segue em anexo o orçamento do Cliente: " 
	Local cEmail   := ""
	Local cAttach  := ""
	Local nAxTotal := 0
	Local nValIcm  := 0
	Local nValIpi  := 0
	Local nAliqIpi := 0
	Local nTotIpi  := 0
	Local nTotIcm  := 0
	Local nAxTotI  := 0
	Local nTotResI := 0
	Local nTotAliIcm := 0  
	Local nTotBasRet := 0
	Local nFrete	 := 0
	Local lRetorno   := .T.
	Local cEndere    := SM0->(Alltrim(M0_ENDENT) + ', ' + Alltrim(M0_BAIRENT) + ', ' + Alltrim(M0_CIDENT) + ', ' + Alltrim(M0_ESTENT) + ', CEP: ' + Alltrim(M0_CEPENT))
	
	Private aCab     := {}
	Private aItens   := {}

	oHtml := TWFHtml():New("\workflow\orcamentos.html")

	SA1->( dbsetorder(1) )
	SA1->( dbseek( xFilial("SA1") + SCJ->CJ_CLIENTE + SCJ->CJ_LOJA ) )
	
	oHtml:cBuffer := StrTran( oHtml:cBuffer, "#LOGOCLI#"    , SuperGetMV("MV_LOGOLNK",.F.,"Logo.png"))	//	Incluir o link do Logo do Cliente
	oHtml:cBuffer := StrTran( oHtml:cBuffer, "#RAZAOSOCIAL#", SM0->M0_NOMECOM)
	AADD(aCab,SM0->M0_NOMECOM)
	
	oHtml:cBuffer := StrTran( oHtml:cBuffer, "#ENDERECO#"   , cEndere)
	AADD(aCab,cEndere)
	
	oHtml:cBuffer := StrTran( oHtml:cBuffer, "#SITE#"       , "https://alergoshop.com.br")	//	Incluir a Pagina WEB do Cliente

	oHtml:cBuffer := StrTran( oHtml:cBuffer, "numorc"    , SCJ->CJ_NUM )
	AADD(aCab,SCJ->CJ_NUM)
	
	cBody         += SA1->A1_NOME + ". "+CRLF
	AADD(aCab,SA1->A1_NOME)
	
	cBody         += "Pedido Número: "+alltrim(SCJ->CJ_NUM)
	oHtml:cBuffer := StrTran( oHtml:cBuffer, "emissao"   , Dtoc(SCJ->CJ_EMISSAO) )
	AADD(aCab,SCJ->CJ_EMISSAO)
	
	oHtml:cBuffer := StrTran( oHtml:cBuffer, "cliente"   , SCJ->CJ_CLIENTE+"/"+SCJ->CJ_LOJA + " - " + SA1->A1_NOME)
	AADD(aCab,SCJ->CJ_CLIENTE)
	AADD(aCab,SCJ->CJ_LOJA)
	
	oHtml:cBuffer := StrTran( oHtml:cBuffer, "cnpj"      , Transform(SA1->A1_CGC,"@R 99.999.999/9999-99"))
	AADD(aCab,SA1->A1_CGC)
	
	oHtml:cBuffer := StrTran( oHtml:cBuffer, "enderec"   , SA1->A1_END)
	AADD(aCab,SA1->A1_END)
	
	oHtml:cBuffer := StrTran( oHtml:cBuffer, "munic"     , SA1->A1_MUN)
	AADD(aCab,SA1->A1_MUN)
	
	oHtml:cBuffer := StrTran( oHtml:cBuffer, "bairro"    , SA1->A1_BAIRRO)
	AADD(aCab,SA1->A1_BAIRRO)
	
	oHtml:cBuffer := StrTran( oHtml:cBuffer, "uff"       , SA1->A1_EST)
	AADD(aCab,SA1->A1_EST)
	
	oHtml:cBuffer := StrTran( oHtml:cBuffer, "cepp"      , Transform(SA1->A1_CEP,"@R 99999-999") )
	AADD(aCab,SA1->A1_CEP)
	
	oHtml:cBuffer := StrTran( oHtml:cBuffer, "telefone"  , SA1->A1_TEL)
	AADD(aCab,SA1->A1_TEL)
	
	oHtml:cBuffer := StrTran( oHtml:cBuffer, "email"     , SA1->A1_EMAIL)
	AADD(aCab,SA1->A1_EMAIL)
	
	cEmail 		  := SA1->A1_EMAIL
	
	oHtml:cBuffer := StrTran( oHtml:cBuffer, "codbanco"  , Posicione("SA6",1,xFilial("SA6")+SCJ->CJ_BANCO,"A6_NOME") + " " + Alltrim(SCJ->CJ_AGENCIA) + " " + Alltrim(SCJ->CJ_CONTA))
	AADD(aCab,Posicione("SA6",1,xFilial("SA6")+SCJ->CJ_BANCO,"A6_NOME") + " AG.: " + Alltrim(SCJ->CJ_AGENCIA) + " CONTA: " + Alltrim(SCJ->CJ_CONTA))
	
	oHtml:cBuffer := StrTran( oHtml:cBuffer, "codseguro" , Transform(SCJ->CJ_SEGURO,"@E 99,999.99"))
	AADD(aCab,SCJ->CJ_SEGURO)
	
	oHtml:cBuffer := StrTran( oHtml:cBuffer, "condpgto"  , Posicione("SE4",1,xFilial("SE4")+SCJ->CJ_CONDPAG,"E4_DESCRI"))
	AADD(aCab,Posicione("SE4",1,xFilial("SE4")+SCJ->CJ_CONDPAG,"E4_DESCRI"))
	
	oHtml:cBuffer := StrTran( oHtml:cBuffer, "codtab"    , SCJ->CJ_TABELA)
	AADD(aCab,SCJ->CJ_TABELA)
	
	oHtml:cBuffer := StrTran( oHtml:cBuffer, "codven"    , SA1->A1_VEND)
	AADD(aCab,SA1->A1_VEND)
	
	oHtml:cBuffer := StrTran( oHtml:cBuffer, "nomevend"  , Posicione("SA3",1,xFilial("SA3")+SA1->A1_VEND,"A3_NOME"))
	AADD(aCab,Posicione("SA3",1,xFilial("SA3")+SA1->A1_VEND,"A3_NOME"))
	
	oHtml:cBuffer := StrTran( oHtml:cBuffer, "comivend"  , Transform(SA1->A1_COMIS,"@E 99,999.99"))
	AADD(aCab,SA1->A1_COMIS)
	//até alterar a tabela pelo configurador
	//oHtml:cBuffer := StrTran( oHtml:cBuffer, "fonetransp", Posicione("SA4",1,xFilial("SA4")+SCJ->CJ_TRANSP,"A4_TEL"))
	oHtml:cBuffer := StrTran( oHtml:cBuffer, "valfrete"	 , Transform(SCJ->CJ_FRETE,"@E 99,999.99"))
	AADD(aCab,SCJ->CJ_FRETE)
	
	nFrete  	  := SCJ->CJ_FRETE
	//oHtml:cBuffer := StrTran( oHtml:cBuffer, "nometransp", Posicione("SA4",1,xFilial("SA4")+SCJ->CJ_TRANSP,"A4_NOME"))

	SCK->( dbsetorder(1) )
	SCK->( dbseek( SCJ->CJ_FILIAL + SCJ->CJ_NUM ) )

	While SCK->( !EOF() ) .AND. SCK->CK_FILIAL == SCJ->CJ_FILIAL .AND. SCK->CK_NUM == SCJ->CJ_NUM
		nValIcm := 0						//Valor ICMS
		nAliIcm := 0						//alícota ICMS
		nValRetI:= 0						// Valor do ICMS Retido
		nBasRetI:= 0						// Base do ICMS Retido
		//nValIpi := 0
		//nAliqIpi:= 0
		//nValSol	:= 0
		BuscaImp(SCK->CK_ITEM,@nValIcm,@nAliIcm,@nValRetI,@nBasRetI)

		cItens += ' <tr> '
		cItens += '   <td width="40">' +SCK->CK_ITEM+'</td> '
		cItens += '   <td width="50">' +SCK->CK_PRODUTO+'</td> '
		cItens += '   <td width="200">'+SCK->CK_DESCRI+'</td> '
		cItens += '   <td width="30"><center>' +SCK->CK_UM+'</td> ' 
		cItens += '   <td width="75"><center>' +Transform(SCK->CK_PRCVEN,"@E 99,999,999,999.99")+'</center></td> '
		cItens += '   <td width="30"><center>' +Transform(SCK->CK_QTDVEN,"@E 99999999")+'</center></td> '
		cItens += '   <td width="58"><center>' +Transform(SCK->CK_VALOR,"@E 99,999,999,999.99")+'</center></td> '
		cItens += '   <td width="30"><center> - </td> '
		cItens += '   <td width="30"><center>' +Transform(nAliIcm,"@E 99,999,999,999.99")+'</center></td> ' //Ali Icms
		cItens += '   <td width="50"><center>' +Transform(nValIcm,"@E 99999999.99")+'</center></td> '
		//cItens += '   <td width="30"><center>' +Transform(nBasRetI,"@E 99,999,999,999.99")+'</center></td> '
		cItens += '   <td width="50"><center>' +Transform(nValRetI,"@E 99,999,999,999.99")+'</center></td> '
		cItens += '   <td width="90"><center>' +Transform(SCK->CK_VALOR + nValIcm + nValRetI, "@E 99,999,999,999.99")+'</center></td> '
		cItens += ' </tr> '
		
		AADD(aItens,{SCK->CK_ITEM   ,SCK->CK_PRODUTO,SCK->CK_DESCRI,SCK->CK_UM,SCK->CK_PRCVEN,SCK->CK_QTDVEN,SCK->CK_VALOR,;
		nAliIcm, nValIcm, nBasRetI, nValRetI, SCK->CK_NUM, SCK->CK_CLIENTE, SCK->CK_LOJA})

		nAxTotal += SCK->CK_VALOR
		//nTotIpi  += nValIpi
		nTotIcm  += nValIcm
		nTotResI += nValRetI
		nTotAliIcm += nAliIcm 
		nTotBasRet += nBasRetI
		nAxTotI  += SCK->CK_VALOR + nValIcm //+ nValRetI*/

		SCK->( dbskip() )
	Enddo

	cTotal := ' <tr style="    background-color: #dfe5f3;"> '
	cTotal += '  <td width="40"><strong>TOTAL:</strong></td> '
	cTotal += '  <td width="50">&nbsp;</td> '
	cTotal += '  <td width="200">&nbsp;</td> '
	cTotal += '  <td width="30">&nbsp;</td> '
	cTotal += '  <td width="75">&nbsp;</td> '
	cTotal += '  <td width="30">&nbsp;</td> '
	cTotal += '  <td width="58"><center>' +Transform(nAxTotal,"@E 99,999,999,999.99")+'</font></center></td> '
	cTotal += '  <td width="30"><center>' +Transform(nFrete,"@E 99,999,999,999.99")+'</font></center></td> '
	cTotal += '  <td width="30">&nbsp;</td> '
	cTotal += '  <td width="50"><center>' +Transform(nTotIcm,"@E 99,999,999,999.99")+'</font></center></td> '
	cTotal += '  <td width="30">&nbsp;</td> '
	//cTotal += '  <td width="50"><center>' +Transform(nTotResI,"@E 99,999,999,999.99")+'</font></center></td> '
	cTotal += '  <td width="90"><center>' +Transform(nAxTotI+nFrete,"@E 99,999,999,999.99")+'</font></center></td> '
	cTotal += '  </tr> '
	oHtml:cBuffer := StrTran( oHtml:cBuffer, "itens" , cItens+cTotal)

	cHtml := oHtml:cBuffer
	if !ExistDir("\workflow\Temp")
		MakeDir("\workflow\Temp")
	EndIf


	If MemoWrite("\workflow\Temp\ORCAMENTO.HTML", cHtml)
		cAttach := "\workflow\Temp\ORCAMENTO.HTML"
		lRetorno := U_FSENVMAIL(cAssunto, cBody, cEmail, cAttach)
		If lRetorno
			U_GravaOrc (aCab, aItens)
		EndIf
	Endif

Return cHtml


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³GERORC    ºAutor  ³Microsiga           º Data ³  07/31/17   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³                                                            º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
Static Function BuscaImp(nItem,nValIcm,nAliIcm,nValRetI,nBasRetI)

	Local aArea	  := GetArea()
	Local aRefImp := {}

	IF !Empty(SCK->CK_TES)
		MaFisIni(SA1->A1_COD,SA1->A1_LOJA,"C","N",SA1->A1_TIPO,aRefImp)
		MaFisIniLoad(Val(nItem))
		MaFisLoad("IT_PRODUTO" , SCK->CK_PRODUTO, Val(nItem))
		MaFisLoad("IT_TES"     , SCK->CK_TES    , Val(nItem))
		MaFisLoad("IT_QUANT"   , SCK->CK_QTDVEN , Val(nItem))
		MaFisLoad("IT_PRCUNI"  , SCK->CK_PRCVEN , Val(nItem))
		MaFisLoad("IT_VALMERC" , SCK->CK_VALOR  , Val(nItem))
		MaFisLoad("IT_BASEIPI" , SCK->CK_VALOR  , Val(nItem))
		MaFisLoad("IT_BASEICM" , SCK->CK_VALOR  , Val(nItem))
		MaFisRecal("IT_TES"    , Val(nItem))
		MaFisRecal("IT_ALIQICM", Val(nItem))
		MaFisRecal("IT_VALIPI" , Val(nItem))
		MaFisRecal("IT_VALICM" , Val(nItem))
		MaFisEndLoad(Val(nItem), 2)

		//nValIpi  := MaFisRet(Val(nItem), "IT_VALIPI")
		nValIcm  := MaFisRet(Val(nItem), "IT_VALICM")           // IT_VALICM
		nAliIcm  := MaFisRet(Val(nItem), "IT_ALIQICM")
		nValRetI := MaFisRet(Val(nItem), "LF_ICMSRET")  
		nBasRetI := MaFisRet(Val(nItem), "LF_BASERET")
		//nAliqIpi := MaFisRet(Val(nItem), "IT_ALIQIPI")
		//nValSol  := MaFisRet(Val(nItem), "IT_VALSOL" )
		
	Endif
	RestArea(aArea)

Return


//=================================================================================================
//	Botão dentro do Orçamento
//User Function MA415BUT

//Return({{" ", {|| U_GerOrc()}, "Envia Orçamento"}})


//=================================================================================================
//	Botão no Browse
User Function MA415MNU

	aAdd(aRotina, {"Envia Orçamento", "U_GerOrc()", 0 , 4, 0, })
	aAdd(aRotina, {"Gerar Pedido"   , "MATA416()", 0 , 4, 0, })

Return


//=================================================================================================
//	Grava orçamento para futura consulta e impressão de relatório

//User Function GravaOrc

//=================================================================================================

User Function GravaOrc(aCab, aItens)

Local aArea  := GetArea()
Local lGrava := .T.
Local x      := 0

DbSelectArea("Z10")
DbSetOrder(1)
If DbSeek(xFilial("Z10") + aCab[3] + aCab[6] + aCab[7])
 lGrava := .F.	
EndIf

Begin Transaction

RecLock("Z10",lGrava)

Z10->Z10_FILIAL   := xFilial("Z10")
Z10->Z10_NUM      := aCab[3]
Z10->Z10_NOMECOM  := aCab[1]
Z10->Z10_END      := aCab[2]
Z10->Z10_NOME     := aCab[4]
Z10->Z10_EMISSA   := dtoc(aCab[5])//aCab[5]
Z10->Z10_CLIENTE  := aCab[6]
Z10->Z10_LOJA     := aCab[7]
Z10->Z10_CGC      := aCab[8]
Z10->Z10_MUN      := aCab[10]
Z10->Z10_BAIRRO   := aCab[11]
Z10->Z10_EST      := aCab[12]
Z10->Z10_CEP      := aCab[13]
Z10->Z10_TEL      := aCab[14]
Z10->Z10_EMAIL    := aCab[15]
Z10->Z10_BANCO    := aCab[16]
Z10->Z10_SEGUR    := aCab[17]
Z10->Z10_DESCR    := aCab[18]
Z10->Z10_TABELA   := aCab[19]
Z10->Z10_VEND     := aCab[20]
Z10->Z10_COMIS    := Val(aCab[21])
Z10->Z10_FRETE    := aCab[22]
Z10->Z10_DATA     := dDataBase
	
MsUnLock()


DbSelectArea("Z11")
DbSetOrder(1)

For x := 1 To Len(aItens)

RecLock("Z11",lGrava)

	Z11->Z11_FILIAL   := xFilial("Z11")
	Z11->Z11_NUM      := aItens[x][12]
	Z11->Z11_CLIENTE  := aItens[x][13]
	Z11->Z11_LOJA     := aItens[x][14]
	Z11->Z11_ITEM     := aItens[x][1]
	Z11->Z11_PRODUTO  := aItens[x][2]
	Z11->Z11_DESCRI   := aItens[x][3]
	Z11->Z11_UM       := aItens[x][4]
	Z11->Z11_PRCVEN   := aItens[x][5]
	Z11->Z11_QTDVEN   := aItens[x][6]
	Z11->Z11_VALOR    := aItens[x][7]
	Z11->Z11_ALIICM   := aItens[x][8]
	Z11->Z11_VALICM   := aItens[x][9]
	Z11->Z11_BASERET  := aItens[x][10]
	Z11->Z11_VALRET   := aItens[x][11]
	Z11->Z11_DATA     := dDataBase
	
	MsUnLock()

Next

END TRANSACTION

RestArea(aArea)

Return