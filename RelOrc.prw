#Include 'Protheus.ch'
#Include 'FWMVCDef.ch'
 
Static cTitulo := "Histórico de Envio de Orçamentos"
 
User Function RelOrc()
    
Local aArea   := GetArea()
Local oBrowse

Private aSN3  := {}
 
oBrowse := FWMBrowse():New()

oBrowse:SetAlias("Z10")
 
oBrowse:SetDescription(cTitulo)
     
//Legendas
/*oBrowse:AddLegend( "Z10->Z10_FLAG != '1'", "RED"  ,  "Não Processado" )
oBrowse:AddLegend( "Z10->Z10_FLAG == '1'", "GREEN",  "Processado" )
 */
//Ativa a Browse
oBrowse:Activate()
 
RestArea(aArea)

Return Nil
 
/*---------------------------------------------------------------------*
 | Func:  MenuDef                                                      |
 | Autor: André Brito                                                |
 | Data:  17/08/2015                                                   |
 | Desc:  Criação do menu MVC                                          |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
 
Static Function MenuDef()

Local aRot := {}
 
//Adicionando opções
//ADD OPTION aRot TITLE 'Incluir'    ACTION 'VIEWDEF.AtuCtaAtf' OPERATION MODEL_OPERATION_INSERT ACCESS 0 //OPERATION X
ADD OPTION aRot TITLE 'Visualizar'   ACTION 'VIEWDEF.RelOrc' OPERATION MODEL_OPERATION_VIEW   ACCESS 0 //OPERATION 1
//ADD OPTION aRot TITLE 'Alterar'    ACTION 'VIEWDEF.AtuCtaAtf' OPERATION MODEL_OPERATION_UPDATE ACCESS 0 //OPERATION 3
ADD OPTION aRot TITLE 'Gerar Excel'  ACTION 'U_RelPlan()' OPERATION MODEL_OPERATION_UPDATE ACCESS 0 //OPERATION 4
ADD OPTION aRot TITLE 'Efetivar Pedido'  ACTION 'MATA416()' OPERATION MODEL_OPERATION_UPDATE ACCESS 0 //OPERATION 4
ADD OPTION aRot TITLE 'Excluir'      ACTION 'VIEWDEF.RelOrc' OPERATION MODEL_OPERATION_DELETE ACCESS 0 //OPERATION 5
 
Return aRot
 
/*---------------------------------------------------------------------*
 | Func:  ModelDef                                                     |
 | Autor: André Brito                                                |
 | Data:  17/08/2015                                                   |
 | Desc:  Criação do modelo de dados MVC                               |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
 
Static Function ModelDef()

Local oModel     := Nil
Local oStPai     := FWFormStruct(1, 'Z10')
Local oStFilho   := FWFormStruct(1, 'Z11')
Local aRel    := {}
 
//Criando o modelo e os relacionamentos
oModel := MPFormModel():New('PROCORC')
oModel:AddFields('Z10MASTER',/*cOwner*/,oStPai)
oModel:AddGrid('Z11DETAIL','Z10MASTER',oStFilho,/*bLinePre*/, /*bLinePost*/,/*bPre - Grid Inteiro*/,/*bPos - Grid Inteiro*/,/*bLoad - Carga do modelo manualmente*/)

oModel:SetRelation("Z11DETAIL", ;       
 					{{"Z11_FILIAL",'xFilial("Z11")'},;        
					{"Z11_NUM","Z10_NUM"  }}, ;       
					Z11->(IndexKey(1)))  
 
oModel:GetModel('Z11DETAIL'):SetUniqueLine({"Z11_FILIAL","Z11_NUM"})    //Não repetir informações ou combinações {"CAMPO1","CAMPO2","CAMPOX"}
oModel:SetPrimaryKey({})
 
//Setando as descrições
oModel:SetDescription("Histórico de Orçamentos")
oModel:GetModel('Z10MASTER'):SetDescription('Dados Cliente')
oModel:GetModel('Z11DETAIL'):SetDescription('Dados Orçamentos')

Return oModel
 
/*---------------------------------------------------------------------*
 | Func:  ViewDef                                                      |
 | Autor: André Brito                                                |
 | Data:  17/08/2015                                                   |
 | Desc:  Criação da visão MVC                                         |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
 
Static Function ViewDef()
    Local oView      := Nil
    Local oModel     := FWLoadModel('RelOrc')
    Local oStPai     := FWFormStruct(2, 'Z10')
    Local oStFilho   := FWFormStruct(2, 'Z11')
     
    //Criando a View
    oView := FWFormView():New()
    oView:SetModel(oModel)
     
    //Adicionando os campos do cabeçalho e o grid dos filhos
    oView:AddField('VIEW_Z10',oStPai,'Z10MASTER')
    oView:AddGrid('VIEW_Z11',oStFilho,'Z11DETAIL')
     
    //Setando o dimensionamento de tamanho
    oView:CreateHorizontalBox('CABEC',45)
    oView:CreateHorizontalBox('GRID',55)
     
    //Amarrando a view com as box
    oView:SetOwnerView('VIEW_Z10','CABEC')
    oView:SetOwnerView('VIEW_Z11','GRID')
     
    //Habilitando título
    oView:EnableTitleView('VIEW_Z10','Dados Cliente')
    oView:EnableTitleView('VIEW_Z11','Historico Orçamento')
    
Return oView

User Function GerCtasx()

RptStatus({|| GerCtaAtf()}, "Aguarde...", "Gravando contas nos ativos...")

Return

Static Function GerCtaAtf()

Return lRet

//==============================================================================================================


User Function RelPlan()

Local aArea     := GetArea()
Local lRet      := .F.
Local oExcelApp := Nil
Local cPath     := AllTrim(GetTempPath())//"C:\Temp"
Local cArquivo  := CriaTrab(,.F.)
Local x         := 0
Local cDirDocs  := MsDocPath()
Local oExcel
Local oExcelApp
Local _oPlan
Local aColunas   := {}
Local aLocais    := {}
Local cQuery     := ""
Local cAliAux    := GetNextAlias()
Local cNumDe     := ""
Local cNumAte    := ""
Local aPWiz      := {}
Local aRetWiz    := {}
Local dDtDe
Local dDtAte

aAdd(aPWiz,{ 1,"Orcamento de: "    ,Space(TamSX3("Z10_NUM")[1]) ,"","","ORC","",    ,.F.})
aAdd(aPWiz,{ 1,"Orcamento ate: "   ,Space(TamSX3("Z10_NUM")[1]) ,"","","ORC","",    ,.F.})
aAdd(aPWiz,{ 1,"Data de: "         ,Ctod("") ,"","",""   ,  ,60 ,.T.})
aAdd(aPWiz,{ 1,"Data de: "         ,Ctod("") ,"","",""   ,  ,60 ,.T.})

aAdd(aRetWiz,Space(TamSX3("Z10_NUM")[1]))
aAdd(aRetWiz,Space(TamSX3("Z10_NUM")[1]))
aAdd(aRetWiz,Ctod(""))
aAdd(aRetWiz,Ctod(""))

lRet := ParamBox(aPWiz,"AlergoShop",aRetWiz,,,,,,,,.T.,.T.) 

cNumDe    := Alltrim(aRetWiz[1]) 
cNumAte   := Alltrim(aRetWiz[2])
dDtDe     := aRetWiz[3]
dDtAte    := aRetWiz[4]

cQuery :="SELECT Z11_FILIAL, Z11_NUM," 
cQuery +=" Z11_ITEM," 
cQuery +=" Z11_PRODUT," 
cQuery +=" Z11_DESCRI," 
cQuery +=" Z11_UM," 
cQuery +=" Z11_PRCVEN," 
cQuery +=" Z11_QTDVEN," 
cQuery +=" Z11_VALOR," 
cQuery +=" Z11_ALIICM," 
cQuery +=" Z11_VALICM," 
cQuery +=" Z11_BASERE," 
cQuery +=" Z11_VALRET," 
cQuery +=" Z11_CLIENT," 
cQuery +=" Z11_LOJA," 
cQuery +=" Z11_DATA,"
cQuery +=" Z10_DATA,"
cQuery +=" Z10_END,"
cQuery +=" Z10_CLIENT,"
cQuery +=" Z10_NOME,"
cQuery +=" Z10_LOJA,"
cQuery +=" Z10_CGC,"
cQuery +=" Z10_MUN,"
cQuery +=" Z10_BAIRRO,"
cQuery +=" Z10_EST,"
cQuery +=" Z10_CEP,"
cQuery +=" Z10_TEL,"
cQuery +=" Z10_EMAIL,"
cQuery +=" Z10_BANCO,"
cQuery +=" Z10_DESCR,"
cQuery +=" Z10_VEND,"
cQuery +=" Z10_COMIS,"
cQuery +=" Z10_FRETE"
cQuery +=" FROM   " + RetSqlName("Z11") + " AS Z11" 
cQuery +=" JOIN " + RetSqlName("Z10") + " AS Z10" 
cQuery +=" ON Z11.Z11_NUM = Z10.Z10_NUM" 
cQuery +=" WHERE  Z11_NUM >= '" + cNumDe + "' AND Z11_NUM <= '" + cNumAte + "' AND (Z11_DATA   >= '" + DTOS(dDtDe)  + "' AND Z11_DATA <= '" + DTOS(dDtAte) + "') AND Z11.D_E_L_E_T_ = ' '" 
cQuery +=" AND Z10.D_E_L_E_T_ = ' '" 
cQuery +=" GROUP  BY Z11_FILIAL, Z11_NUM," 
cQuery +=" Z11_ITEM," 
cQuery +=" Z11_PRODUT," 
cQuery +=" Z11_DESCRI," 
cQuery +=" Z11_UM," 
cQuery +=" Z11_PRCVEN," 
cQuery +=" Z11_QTDVEN," 
cQuery +=" Z11_VALOR," 
cQuery +=" Z11_ALIICM," 
cQuery +=" Z11_VALICM," 
cQuery +=" Z11_BASERE," 
cQuery +=" Z11_VALRET," 
cQuery +=" Z11_CLIENT," 
cQuery +=" Z11_LOJA," 
cQuery +=" Z11_DATA, Z10_DATA,Z10_END,Z10_CLIENT,Z10_NOME,Z10_LOJA,Z10_CGC,Z10_MUN,Z10_BAIRRO,Z10_EST,Z10_CEP,Z10_TEL,Z10_EMAIL,Z10_BANCO,Z10_DESCR,Z10_VEND,Z10_COMIS,Z10_FRETE" 

cQuery := ChangeQuery(cQuery) 
 
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliAux,.T.,.T.)

oBrush1  := TBrush():New(, RGB(193,205,205))

oExcel  := FWMSExcel():New()
cAba    := "Orçamentos Enviados"
cTabela := "Orçamentos - AlergoShop"

// Criação de nova aba 
oExcel:AddworkSheet(cAba)

// Criação de tabela
oExcel:AddTable (cAba,cTabela)

// Criação de colunas 

oExcel:AddColumn(cAba,cTabela,"FILIAL"              ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"NUM. ORCAMENTO"      ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"ITEM ORCAMENTO"      ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"DAT. ORCAMENTO"      ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"ENDERECO"            ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"COD. CLIENTE"        ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"NOME"                ,2,1,.F.) 
oExcel:AddColumn(cAba,cTabela,"LOJA"                ,2,1,.F.)  
oExcel:AddColumn(cAba,cTabela,"CNPJ"                ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"MUNICIPIO"           ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"BAIRRO"              ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"ESTADO"              ,2,1,.F.) 
oExcel:AddColumn(cAba,cTabela,"CEP"                 ,2,1,.F.) 
oExcel:AddColumn(cAba,cTabela,"TELEFONE"            ,2,1,.F.) 
oExcel:AddColumn(cAba,cTabela,"E-MAIL"              ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"BANCO"               ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"DESCRICAO ORCAMENTO" ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"PRODUTO"             ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"DESCRICAO PRODUTO"   ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"UNIDADE MEDIDA"      ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"PRECO"               ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"QUANTIDADE"          ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"VALOR"               ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"ALIQ. ICM"           ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"VALOR ICM"           ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"BASE RETENCAO"       ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"VALOR RETIDO"        ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"VENDEDOR"            ,2,1,.F.) 
oExcel:AddColumn(cAba,cTabela,"COMISSAO"            ,2,1,.F.) 
oExcel:AddColumn(cAba,cTabela,"FRETE"               ,2,1,.F.) 


DbselectArea(cAliAux)
dbGoTop()

While !(cAliAux)->(Eof())
 	//oProcExc:IncRegua2("Gerando a planilha...")	
    // Criação de Linhas 
    cData := STOD((cAliAux)->Z10_DATA)
    cData := DTOC(cData)
    oExcel:AddRow(cAba,cTabela, { (cAliAux)->Z11_FILIAL ,;
                                  (cAliAux)->Z11_NUM   ,;
                                  (cAliAux)->Z11_ITEM  ,;
                                  cData  ,; 
                                  (cAliAux)->Z10_END   ,; 
                                  (cAliAux)->Z10_CLIENT ,;
                                  (cAliAux)->Z10_NOME   ,;
                                  (cAliAux)->Z10_LOJA ,;
                                  (cAliAux)->Z10_CGC ,;
                                  Alltrim((cAliAux)->Z10_MUN)     ,;
                                  Alltrim((cAliAux)->Z10_BAIRRO) ,;
                                  (cAliAux)->Z10_EST,;
                                  (cAliAux)->Z10_CEP  ,;
                                  (cAliAux)->Z10_TEL   ,;
                                  Alltrim((cAliAux)->Z10_EMAIL) ,;
                                  Alltrim((cAliAux)->Z10_BANCO) ,;
                                  Alltrim((cAliAux)->Z10_DESCR),;
                                  (cAliAux)->Z11_PRODUT ,;
                                  Alltrim((cAliAux)->Z11_DESCRI) ,;
                                  (cAliAux)->Z11_UM ,;
                                  Alltrim(TRANSFORM( (cAliAux)->Z11_PRCVEN, "@E 999,999.99")) ,;
                                  (cAliAux)->Z11_QTDVEN ,;
                                 Alltrim(TRANSFORM( (cAliAux)->Z11_VALOR, "@E 999,999,999.99")),;
                                  (cAliAux)->Z11_ALIICM ,;
                                  Alltrim(TRANSFORM( (cAliAux)->Z11_VALICM, "@E 999,999.99")) ,;
                                  (cAliAux)->Z11_BASERE ,;
                                  Alltrim(TRANSFORM( (cAliAux)->Z11_VALRET, "@E 999,999.99")) ,;
                                  (cAliAux)->Z10_VEND ,;
                                  Alltrim(TRANSFORM( (cAliAux)->Z10_COMIS, "@E 999,999.99")) ,;
                                 Alltrim(TRANSFORM( (cAliAux)->Z10_FRETE, "@E 999,999.99"))})

    (cAliAux)->(dbSkip())

End


DbCloseArea(cAliAux)

If !Empty(oExcel:aWorkSheet)

    oExcel:Activate()
    oExcel:GetXMLFile(cArquivo)
 
    CpyS2T("\SYSTEM\"+cArquivo, cPath)

    oExcelApp := MsExcel():New()
    oExcelApp:WorkBooks:Open(cPath + "\" + cArquivo) // Abre a planilha
	oExcelApp:SetVisible(.T.)
	
EndIf

MSGINFO( "Arquivo: " + cArquivo + " gerado no diretório " + cPath + "", "Concluido" )

RestArea(aArea)

Return ''