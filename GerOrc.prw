#Include "Protheus.ch"
#Include "RESTFUL.ch"
#Include "tbiconn.ch"
#Include "RwMake.ch"

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
	Local cEmail   := SA1->A1_EMAIL
	Local cAttach  := ""
	Local nAxTotal := 0
	Local nValIcm  := 0
	Local nTotIcm  := 0
	Local nAxTotI  := 0
	Local nTotResI := 0
	Local nTotAliIcm := 0
	Local nTotBasRet := 0
	Local nFrete	 := 0
	Local lRetorno   := .T.
	Local cEndere    := SM0->(Alltrim(M0_ENDENT) + ', ' + Alltrim(M0_BAIRENT) + ', ' + Alltrim(M0_CIDENT) + ', ' + Alltrim(M0_ESTENT) + ', CEP: ' + Alltrim(M0_CEPENT))
	Local aArea      := GetArea()
	Local aAreaBkp   := {}

	Private aCab     := {}
	Private aItens   := {}
	Private aImposto := {}
	Private nTotal   := 0

	If Empty(cEmail)
		MsgBox("O Orçamento não será enviado, pois esse Cliente não tem e-mail cadastrado")
	Else
		oHtml := TWFHtml():New("\workflow\orcamentos.html")

		oHtml:cBuffer := StrTran( oHtml:cBuffer, "#LOGOCLI#"    , SuperGetMV("MV_LOGOLNK",.F.,"\WorkFlow\Logo.png"))	//	Incluir o link do Logo do Cliente
		oHtml:cBuffer := StrTran( oHtml:cBuffer, "#RAZAOSOCIAL#", SM0->M0_NOMECOM)
		AADD(aCab,SM0->M0_NOMECOM)

		oHtml:cBuffer := StrTran( oHtml:cBuffer, "#ENDERECO#"   , cEndere)
		AADD(aCab,cEndere)

		oHtml:cBuffer := StrTran( oHtml:cBuffer, "#SITE#"       , "https://alergoshop.com.br")	//	Incluir a Pagina WEB do Cliente

		oHtml:cBuffer := StrTran( oHtml:cBuffer, "numorc"       , SCJ->CJ_NUM )
		AADD(aCab,SCJ->CJ_NUM)

		cBody += SA1->A1_NOME + ". "+CRLF
		AADD(aCab,SA1->A1_NOME)

		cBody += "Pedido Número: "+alltrim(SCJ->CJ_NUM)
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

		oHtml:cBuffer := StrTran( oHtml:cBuffer, "email"     , cEmail)
		AADD(aCab,cEmail)

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
		//oHtml:cBuffer := StrTran( oHtml:cBuffer, "nometransp", Posicione("SA4",1,xFilial("SA4")+SCJ->CJ_TRANSP,"A4_NOME")))
		aAreaBkp := GetArea()
		aImposto := Ma415Impos()	//	Ma415Imops(3)
		RestArea(aAreaBkp)
		SCK->( dbsetorder(1) )
		SCK->( dbseek( SCJ->CJ_FILIAL + SCJ->CJ_NUM ))
		While SCK->( !EOF() ) .AND. SCK->CK_FILIAL == SCJ->CJ_FILIAL .AND. SCK->CK_NUM == SCJ->CJ_NUM
			ny := Val(SCK->CK_ITEM)
			nValIcm := 0						//Valor ICMS
			nAliIcm := 0						//alícota ICMS
			nValRetI:= 0						// Valor do ICMS Retido
			nBasRetI:= 0						// Base do ICMS Retido
			cItens += ' <tr> '
			cItens += '   <td width="40">' +SCK->CK_ITEM+'</td> '
			cItens += '   <td width="50">' +SCK->CK_PRODUTO+'</td> '
			cItens += '   <td width="200">'+SCK->CK_DESCRI+'</td> '
			cItens += '   <td width="30"><center>' +SCK->CK_UM+'</td> '
			cItens += '   <td width="75"><center>' +Transform(SCK->CK_PRCVEN,"@E 99,999,999,999.99")+'</center></td> '
			cItens += '   <td width="30"><center>' +Transform(SCK->CK_QTDVEN,"@E 99999999")+'</center></td> '
			cItens += '   <td width="58"><center>' +Transform(SCK->CK_VALOR,"@E 99,999,999,999.99")+'</center></td> '
			cItens += '   <td width="30"><center> - </td> '//aAdd(aImposto,{nValIcm,nAliIcm,nValRetI,nBasRetI})
			cItens += '   <td width="30"><center>' +Transform(aImposto[ny, 2],"@E 99,999,999,999.99")+'</center></td> '
			cItens += '   <td width="50"><center>' +Transform(aImposto[ny, 1],"@E 99999999.99")+'</center></td> '
			cItens += '   <td width="50"><center>' +Transform(aImposto[ny, 3],"@E 99,999,999,999.99")+'</center></td> '
			cItens += '   <td width="90"><center>' +Transform(SCK->CK_VALOR + aImposto[ny, 1] + aImposto[ny, 3], "@E 99,999,999,999.99")+'</center></td> '
			cItens += ' </tr> '
			nValIcm  := aImposto[ny, 1]
			nAliIcm  := aImposto[ny, 2]
			nValRetI := aImposto[ny, 3]
			nBasRetI := aImposto[ny, 4]
			AADD(aItens,{SCK->CK_ITEM, SCK->CK_PRODUTO, SCK->CK_DESCRI, SCK->CK_UM, SCK->CK_PRCVEN, SCK->CK_QTDVEN, SCK->CK_VALOR, ;
				nAliIcm, nValIcm, nBasRetI, nValRetI, SCK->CK_NUM, SCK->CK_CLIENTE, SCK->CK_LOJA})
			nAxTotal += SCK->CK_VALOR
			nTotIcm  += nValIcm
			nTotResI += nValRetI
			nTotAliIcm += nAliIcm
			nTotBasRet += nBasRetI
			nAxTotI  += SCK->CK_VALOR + nValIcm //+ nValRetI*/
			SCK->( dbskip() )
		EndDo
		MaFisEnd()
		SCK->( dbsetorder(1) )
		SCK->( dbseek( SCJ->CJ_FILIAL + SCJ->CJ_NUM ) )
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
		cTotal += '  <td width="50"><center>' +Transform(nTotResI,"@E 99,999,999,999.99")+'</font></center></td> '
		cTotal += '  <td width="90"><center>' +Transform(nTotal+nFrete+nTotIcm,"@E 99,999,999,999.99")+'</font></center></td> '
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
		RestArea(aArea)
	Endif

Return cHtml


//=================================================================================================
//	Botão no Browse
User Function MA415MNU

	aAdd(aRotina, {"Envia Orçamento", "U_GerOrc()", 0 , 4, 0, })
	aAdd(aRotina, {"Gerar Pedido"   , "U_RedMata()", 0 , 4, 0, })

Return


//=================================================================================================
//	Grava orçamento para futura consulta e impressão de relatório
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
		Z10->Z10_NUM      := aCab[03]
		Z10->Z10_NOMECOM  := aCab[01]
		Z10->Z10_END      := aCab[02]
		Z10->Z10_NOME     := aCab[04]
		Z10->Z10_EMISSA   := dtoc(aCab[5])//aCab[5]
		Z10->Z10_CLIENTE  := aCab[06]
		Z10->Z10_LOJA     := aCab[07]
		Z10->Z10_CGC      := aCab[08]
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
			Z11->Z11_FILIAL  := xFilial("Z11")
			Z11->Z11_NUM     := aItens[x, 12]
			Z11->Z11_CLIENTE := aItens[x, 13]
			Z11->Z11_LOJA    := aItens[x, 14]
			Z11->Z11_ITEM    := aItens[x, 01]
			Z11->Z11_PRODUTO := aItens[x, 02]
			Z11->Z11_DESCRI  := aItens[x, 03]
			Z11->Z11_UM      := aItens[x, 04]
			Z11->Z11_PRCVEN  := aItens[x, 05]
			Z11->Z11_QTDVEN  := aItens[x, 06]
			Z11->Z11_VALOR   := aItens[x, 07]
			Z11->Z11_ALIICM  := aItens[x, 08]
			Z11->Z11_VALICM  := aItens[x, 09]
			Z11->Z11_BASERET := aItens[x, 10]
			Z11->Z11_VALRET  := aItens[x, 11]
			Z11->Z11_DATA    := dDataBase
			MsUnLock()
		Next
	END TRANSACTION
	RestArea(aArea)

Return


//=================================================================================================
User Function RedMata()

	Local aArea   := GetArea()

	SetFunName("MATA416")
	MATA416()
	SetFunName("GERORC")
	RestArea(aArea)

Return


//=================================================================================================
Static Function Ma415Impos

	Local nPrcLista := 0
	Local nValMerc  := 0
	Local nDesconto := 0
	Local nAcresFin := 0
	Local nQtdPeso  := 0
	Local dDataCnd  := SCJ->CJ_EMISSAO
	Local lProspect := .F.
	Local cTipo	  	:= ""
	Local cFilSB1  := xFilial("SB1")
	Local cFilSB2  := xFilial("SB2")
	Local cFilSF4  := xFilial("SF4")
	Local lDescSai := Iif(GetNewPar('MV_DESCSAI','1') == "2",.T.,.F.)
	Local nItem    := 1

	SCK->( dbsetorder(1) )
	SCK->( dbseek( SCJ->CJ_FILIAL + SCJ->CJ_NUM ))

	//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
	//³Posiciona Registros                          ³
	//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
	dbSelectArea("SA1")
	dbSetOrder(1)
	MsSeek(xFilial("SA1")+If(!Empty(SCJ->CJ_CLIENT),SCJ->CJ_CLIENT,SCJ->CJ_CLIENTE)+SCJ->CJ_LOJAENT)
	dbSelectArea("SE4")
	dbSetOrder(1)
	MsSeek(xFilial("SE4")+SCJ->CJ_CONDPAG)

	//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
	//³Inicializa a funcao fiscal                   ³
	//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
	If !Empty(SCJ->CJ_PROSPE) .And. !Empty(SCJ->CJ_LOJPRO)
		cTipo := Posicione("SUS",1,xFilial("SUS") + SCJ->CJ_PROSPE + SCJ->CJ_LOJPRO,"US_TIPO")
		lProspect := .T.
	Endif

	//Altera cFilAnt para calcular os impostos na filial de entrega.
	MaFisSave()
	MaFisEnd()
	MaFisIni(Iif(Empty(SCJ->CJ_CLIENT),SCJ->CJ_CLIENTE,SCJ->CJ_CLIENT),;// 1-Codigo Cliente/Fornecedor
		SCJ->CJ_LOJAENT,;		// 2-Loja do Cliente/Fornecedor
		"C",;				// 3-C:Cliente , F:Fornecedor
		"N",;				// 4-Tipo da NF
		Iif(lProspect,cTipo,Iif(SCJ->(ColumnPos("CJ_TIPOCLI")) > 0,SCJ->CJ_TIPOCLI,SA1->A1_TIPO)),;		// 5-Tipo do Cliente/Fornecedor
		Nil,;
		Nil,;
		Nil,;
		Nil,;
		"MATA461",;
		Nil,;
		Nil,;
		IiF(lProspect,SCJ->CJ_PROSPE+SCJ->CJ_LOJPRO,""))

	While SCK->( !EOF() ) .AND. SCK->CK_FILIAL == SCJ->CJ_FILIAL .AND. SCK->CK_NUM == SCJ->CJ_NUM
		//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
		//³ Verifica se a linha foi deletada                     ³
		//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
		SB2->(dbSetOrder(1))
		SB2->(MsSeek(cFilSB2 + SCK->CK_PRODUTO + SCK->CK_LOCAL))
		SF4->(dbSetOrder(1))
		SF4->(MsSeek(cFilSF4 + SCK->CK_TES))

		//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
		//³Calcula o preco de lista                     ³
		//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
		nValMerc  := SCK->CK_VALOR
		nPrcLista := SCK->CK_PRUNIT
		nQtdPeso  := 0
		If ( nPrcLista == 0 )
			nPrcLista := A410Arred(nValMerc/SCK->CK_QTDVEN,"CK_PRCVEN")
		EndIf
		nAcresFin := A410Arred(SCK->CK_PRCVEN*SE4->E4_ACRSFIN/100,"D2_PRCVEN")
		nValMerc  += A410Arred(nAcresFin*SCK->CK_QTDVEN,"D2_TOTAL")
		nDesconto := A410Arred(nPrcLista*SCK->CK_QTDVEN,"D2_DESCON")-nValMerc
		nDesconto := IIf(nDesconto==0,SCK->CK_VALDESC,nDesconto)
		nDesconto := Max(0,nDesconto)
		nPrcLista += nAcresFin

		//Para os outros paises, este tratamento e feito no programas que calculam os impostos.
		If cPaisLoc=="BRA" .or. lDescSai
			nValMerc  += nDesconto
		Endif

		//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
		//³Verifica a data de entrega para as duplicatas³
		//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
		If ( dDataCnd > SCK->CK_ENTREG .And. !Empty(SCK->CK_ENTREG) )
			dDataCnd := SCK->CK_ENTREG
		EndIf

		//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
		//³Agrega os itens para a funcao fiscal         ³
		//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
		MaFisAdd(SCK->CK_PRODUTO,;   	// 1-Codigo do Produto ( Obrigatorio )
			SCK->CK_TES,;	   	// 2-Codigo do TES ( Opcional )
			SCK->CK_QTDVEN,;  	// 3-Quantidade ( Obrigatorio )
			nPrcLista,;		  	// 4-Preco Unitario ( Obrigatorio )
			nDesconto,; 	// 5-Valor do Desconto ( Opcional )
			"",;	   			// 6-Numero da NF Original ( Devolucao/Benef )
			"",;				// 7-Serie da NF Original ( Devolucao/Benef )
			0,;					// 8-RecNo da NF Original no arq SD1/SD2
			0,;					// 9-Valor do Frete do Item ( Opcional )
			0,;					// 10-Valor da Despesa do item ( Opcional )
			0,;					// 11-Valor do Seguro do item ( Opcional )
			0,;					// 12-Valor do Frete Autonomo ( Opcional )
			nValMerc,;			// 13-Valor da Mercadoria ( Obrigatorio )
			0,;					// 14-Valor da Embalagem ( Opiconal )
			, , , , , , , , , , , , ,;
			SCK->CK_CLASFIS) // 28-Classificacao fiscal)
		SB1->(dbSetOrder(1))
		If SB1->(MsSeek(cFilSB1 + SCK->CK_PRODUTO))
			nQtdPeso := SCK->CK_QTDVEN*SB1->B1_PESO
		Endif

		//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
		//³Calculo do ISS                               ³
		//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
		If SA1->A1_INCISS == "N"
			If ( SF4->F4_ISS=="S" )
				nPrcLista := a410Arred(nPrcLista/(1-(MaAliqISS(nItem)/100)),"D2_PRCVEN")
				nValMerc  := a410Arred(nValMerc/(1-(MaAliqISS(nItem)/100)),"D2_PRCVEN")
				MaFisAlt("IT_PRCUNI",nPrcLista,nItem)
				MaFisAlt("IT_VALMERC",nValMerc,nItem)
			EndIf
		EndIf
		nValIcm  := MaFisRet(nItem, "IT_VALICM")           // IT_VALICM
		nAliIcm  := MaFisRet(nItem, "IT_ALIQICM")
		nValRetI := MaFisRet(nItem, "LF_ICMSRET")
		nBasRetI := MaFisRet(nItem, "LF_BASERET")
		aAdd(aImposto,{nValIcm,nAliIcm,nValRetI,nBasRetI})
		nItem ++
		SCK->( dbskip() )
	EndDo
	nTotal := MaFisRet(,"NF_TOTAL")

Return aImposto


//=================================================================================================
User Function GatOper()

	Local cOper := "01"
	Local cTesInt := ""

	M->CK_OPER := cOper
	cTesInt := MaTesInt(2, M->CK_OPER, M->CJ_CLIENTE, M->CJ_LOJA, "C", M->CK_PRODUTO, "CK_TES")
	M->CK_TES := cTesInt
	M->CK_CLASFIS := CodSitTri()

Return cOper
