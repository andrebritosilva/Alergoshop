#Include 'Protheus.ch'
#Include 'tbiconn.ch'
/*
{Protheus.doc} FSENVMAIL()
Envia email
@Author     Fernando Carvalho
@Since      16/08/2017
@Version    P12.7
@Project    Portal Protheus
@Param		 cAssunto
@Param		 cBody
@Param		 cEmail
*/
User Function FSENVMAIL(cAssunto, cBody, cEmail,cAttach/*caminho do arquivo a ser anexado*/,cMailConta,cUsuario,cMailServer,cMailSenha,lMailAuth,lUseSSL,lUseTLS,cCopia,cCopiaOculta)
	
	Local nMailPort		:= 0
	Local nAt			:= ""
	Local lRet 			:= .T.
	Local oServer		:= TMailManager():New()
	Local aAttach		:= {}
	Local nLoop			:= 0
	
	Default cAttach		:= ''
	Default cMailConta	:= SuperGetMV("MV_RELACNT")
	Default cUsuario	:= SubStr(cMailConta,1,At("@",cMailConta)-1)
	Default cMailServer	:= "smtp.ethosx.com"//AllTrim(SuperGetMv("MV_RELSERV"))
	Default cMailSenha	:= SuperGetMV("MV_RELPSW")
	Default lMailAuth	:= .T.//SuperGetMV("MV_RELAUTH",,.F.)
	Default lUseSSL		:= SuperGetMV("MV_RELSSL",,.F.)
	Default lUseTLS		:= SuperGetMV("MV_RELTLS",,.F.)
	Default cCopia		:= ''
	
	nAt			:= At(":",cMailServer)
	
	oServer:SetUseSSL(lUseSSL)
	oServer:SetUseTLS(lUseTLS)
	
	
	// Tratamento para usar a porta quando informada no mailserver
	If nAt > 0
		nMailPort	:= VAL(SUBSTR(ALLTRIM(cMailServer),At(":",cMailServer) + 1,Len(ALLTRIM(cMailServer)) - nAt))
		cMailServer	:= SUBSTR(ALLTRIM(cMailServer),1,At(":",cMailServer)-1)
		oServer:Init("", cMailServer, cMailConta, cMailSenha,0,nMailPort)
	Else
		oServer:Init("", cMailServer, cMailConta, cMailSenha,0,587)
	EndIf
	
	If oServer:SMTPConnect() != 0
		lRet := .F.
		alert("Servidor não conectou!"+CRLF+"Servidor: "+cMailServer+CRLF+"Verifique os dados cadastrados no Configurador."+CRLF+"Acesse Ambiente -> E-mail/Proxy -> Configurar")
	EndIf
	
	If lRet
		If lMailAuth
			
			//Tentar com conta e senha
			If oServer:SMTPAuth(cMailConta, cMailSenha) != 0
				
				//Tentar com usuário e senha
				If oServer:SMTPAuth(cUsuario, cMailSenha) != 0
					lRet := .F.
					alert("Autenticação do servidor não funcionou!"+CRLF+ "Conta: "+cMailConta+".  Usuário: "+cUsuario+".  Senha: "+cMailSenha+"."+CRLF+"Verifique os dados cadastrados no Configurador."+CRLF+"Acesse Ambiente -> E-mail/Proxy -> Configurar")
				EndIf
				
			EndIf
			
		EndIf
	EndIf
	
	If lRet
		
		oMessage				:= TMailMessage():New()
		
		oMessage:Clear()
		oMessage:cFrom			:= cMailConta
		oMessage:cTo			:= cEmail
		oMessage:cCc			:= cCopia
		oMessage:cBCC			:= cCopiaOculta
		oMessage:cSubject		:= cAssunto
		oMessage:cBody			:= cBody
		
		//oMessage:AttachFile( cAttach )
		aAttach	:= StrTokArr(cAttach, ';')
		
		For nLoop := 1 To Len(aAttach)
			oMessage:AttachFile( aAttach[nLoop] )
		Next
		//Envia o e-mail
		
		nErro := oMessage:Send( oServer )
  		If( nErro != 0 )
   			 //conout( "Não enviou o e-mail.", oServer:GetErrorString( nErro ) )
				MsgaInfo( oServer:GetErrorString( nErro ) ,"Não enviou o e-mail.")
    		Return
  		EndIf
		
	EndIf
	 
	//Desconecta do servidor
	oServer:SMTPDisconnect()
	if lRet 
		MsgInfo("Email enviado com sucesso!","Envia Orçamento")
	Else
		Alert("Email não enviado!")
	Endif
Return lRet