#Include 'Protheus.ch'
#Include 'Parmtype.ch'
#include "FWMVCDEF.CH"
#include "tbiconn.ch"
#include "topconn.ch"

//-------------------------------------------------------------------
/*/{Protheus.doc} LAUA0004
Tela para cadastro de faturas automaticas
@author  Sidney Sales
@since   05/12/19
@version 1.0
/*/
//-------------------------------------------------------------------
User Function LAUA0004()
	Local oBrowse
	aRotina := {}
	//Montagem do Browse principal	
	oBrowse := FWMBrowse():New()
	oBrowse:SetAlias('ZA7')
	oBrowse:SetDescription('Faturamento Automático')
	oBrowse:AddLegend( "ZA7_STATUS=='G'"										,"RED"		,	"Faturas geradas" )	// "Bloqueada"
	oBrowse:AddLegend( "ZA7_STATUS=='C'"										,"BLACK"	,	"Faturas canceladas" )	// "Vigente"
	oBrowse:AddLegend( "ZA7_STATUS=='P'"										,"YELLOW"	,	"Falha no processamento" )	// "Cancelada"
	oBrowse:SetMenuDef('LAUA0004')

	ZA7->(DbSetFilter({|| ZA7->ZA7_STATUS == 'G' }, "ZA7_STATUS == 'G' "))
	oBrowse:Activate()
	
	ZA7->(DBClearFilter())
Return

//Montagem do menu 
Static Function MenuDef()
	Local aRotina := {}

	aAdd( aRotina, { 'Incluir' 			, 'U_LAUA004A'			, 0, 3, 0, NIL 		} )
	aAdd( aRotina, { 'Ver Faturas'		, 'U_LAUA0007(.F.)'			, 0, 4, 0, NIL 	} )
	aAdd( aRotina, { 'Todas as Fat.'	, 'U_LAUA0007(.T.)'			, 0, 4, 0, NIL 	} )
	aAdd( aRotina, { 'Gerar Boletos'	, 'U_LAUF0006'			, 0, 4, 0, NIL 		} )
	aAdd( aRotina, { 'Apenas faturas'	, 'U_LAUA007L'			, 0, 3, 0, NIL 		} )
	aAdd( aRotina, { 'Cancelar Fatura'	, 'U_LAUF006A'			, 0, 4, 0, NIL 		} )
	aAdd( aRotina, { 'Importar CSV'		, 'U_LAUF0010'			, 0, 3, 0, NIL 		} )
//	aAdd( aRotina, { 'Excluir' 		, 'VIEWDEF.LAUA0004'	, 0, 5, 0, NIL } )
Return aRotina

//Construcao do mdelo
Static Function ModelDef()
	Local oModel
	Local oStruZA7 := FWFormStruct(1,"ZA7")

	oModel := MPFormModel():New("MD_ZA7") 
	oModel:addFields('MASTERZA7',,oStruZA7)
	oModel:SetPrimaryKey({'ZA7_FILIAL', 'ZA7_CODIGO'})

Return oModel

//Construcao da visualizacao
Static Function ViewDef()
	Local oModel := ModelDef()
	Local oView
	Local oStrZA7:= FWFormStruct(2, 'ZA7')

	oView := FWFormView():New()
	oView:SetModel(oModel)

	oView:AddField('FORM_ZA7' , oStrZA7,'MASTERZA7' ) 
	oView:CreateHorizontalBox( 'BOX_FORM_ZA7', 100)
	oView:SetOwnerView('FORM_ZA7','BOX_FORM_ZA7')

Return oView

//-------------------------------------------------------------------
/*/{Protheus.doc} LAUA004A
Funcao que monta tela de filtro para geracao das faturas
@author  Sidney Sales
@since   29/11/2019
@version 1.0
/*/
//-------------------------------------------------------------------
User Function LAUA004A()

	
	Local oSize 	:= FwDefSize():New()
	Local oLayer 	:= FWLayer():new()
	Local dEmissao
	Local oDlg  	:= FWDialogModal():New()
	Local dDtAtu  	:= dDataBase
	Private dDtEmissa  	:= dDataBase
	Private oBrowseSE1
	Private aRecnosSE1 := {}
	Private nTotFat := 0	
	Private nTotVlr := 0	
	Private nTotCTE := 0	
	Private cMark
	Private nRecNSE1
	Private cGrupo, cDescGrupo, cCliente, cNomeCliente,	cNatureza, cDescNat, cLojaCli
	Private cTpFrete := ""
	Private cTextoTit := ""
	Private cNatureza := Alltrim(SuperGetMV("MS_NATFATU", .F., "110112"))

	If Empty(FunName())
		PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01RN0001'
		oDlg:SetSize(400,600)	
	Else
		oDlg:enableAllClient()
	EndIf

	cMark 	:= GetMark()
	
	oDlg:SetTitle('Faturamento automático')
	oDlg:SetSubTitle('Geração de faturas')	
	
	oDlg:CreateDialog()
		
	oPanelModal := oDlg:GetPanelMain()
	
	oLayer := FwLayer():New()
	oLayer:Init(oPanelModal)

	//Linha dos Filtros
	oLayer:addLine("lin01",45,.T.)		
	oLayer:AddCollumn('col01', 60, .T., 'lin01')	
	oLayer:addWindow("col01","win01",'Filtros',100,.F.,.F.,{|| }, "lin01",{|| })
	oPanel01 := oLayer:getWinPanel('col01', 'win01', "lin01")
	
	cGrupo			:= Space(Len(SA1->A1_GRPVEN))
    oSay1			:= TSay():create(oPanel01, {||  'Grupo de Cliente' },02,05,,,,,,.T.,,,200,20)
	oGetGrupo		:= TGet():New( 010, 005, { | u | If( PCount() == 0, cGrupo, cGrupo := u ) },oPanel01, 050, 010, "@!",{|| Validar('grupo') }, 0, 16777215,,.F.,,.T.,,.F.,,.F.,.F.,,.F.,.F. ,"ACY","cGrupo",,,,.T.  )
	
	cDescGrupo		:= Space(150)
	oGetDescGrupo 	:= TGet():Create( oPanel01,{||cDescGrupo},10,60,180,10,"@!",,0,,,.F.,,.T.,,.F.,,.F.,.F.,,.T.,.F.,,cDescGrupo,,,, )    
	oGetDescGrupo:lActive := .F.

	cDoc		:= Space(LEN(SE1->E1_NUM))
	oSay1			:= TSay():create(oPanel01, {||  'Numero Doc' },02,240,,,,,,.T.,,,200,20)
	oGetDoc		:= TGet():New( 010, 240, { | u | If( PCount() == 0, cDoc, cDoc := u ) },oPanel01, 050, 010, "@!", , 0, 16777215,,.F.,,.T.,,.F.,,.F.,.F.,,.F.,.F. ,"","cDoc",,,,.T.  )	

    cCliente		:= Space(Len(SA1->A1_COD))
	
	oSay1			:= TSay():create(oPanel01, {||  'Cliente' },027,05,,,,,,.T.,,,200,20)
	oGetCliente		:= TGet():New( 035, 005, { | u | If( PCount() == 0, cCliente, cCliente := u ) },oPanel01, 050, 010, "@!",{|| Validar('cliente')}, 0, 16777215,,.F.,,.T.,,.F.,,.F.,.F.,,.F.,.F. ,"SA1","cCliente",,,,.T.  )	

    cLojaCli		:= Space(Len(SA1->A1_COD))
	oSay1			:= TSay():create(oPanel01, {||  'Loja' },027,60,,,,,,.T.,,,200,20)
	oGetCliente		:= TGet():New( 035, 60, { | u | If( PCount() == 0, cLojaCli, cLojaCli := u ) },oPanel01, 030, 010, "@!",{|| Validar('loja')}, 0, 16777215,,.F.,,.T.,,.F.,,.F.,.F.,,.F.,.F. ,"","cLojaCli",,,,.T.  )	

	cNomeCliente	:= Space(150)
	oGetNomeCliente	:= TGet():Create( oPanel01,{||cNomeCliente},35,95,220,10,"@!",,0,,,.F.,,.T.,,.F.,,.F.,.F.,,.T.,.F.,,cNomeCliente,,,, )    
	oGetNomeCliente:lActive := .F.	

	cTpFrete		:= Space(1)
	oSay1			:= TSay():create(oPanel01, {||  'Tipo Frete' },027,315,,,,,,.T.,,,200,20)
	oGetFrete		:= TGet():New( 035, 315, { | u | If( PCount() == 0, cTpFrete, cTpFrete := u ) },oPanel01, 020, 010, "@!", , 0, 16777215,,.F.,,.T.,,.F.,,.F.,.F.,,.F.,.F. ,"","cTpFrete",,,,.T.  )	

	dEmissDe		:= dDataBase
    oSay1			:= TSay():create(oPanel01, {||  'Emissao Inicial' },050,05,,,,,,.T.,,,200,20)
	oGetDataDe		:= TGet():New( 058, 005, { | u | If( PCount() == 0, dEmissDe , dEmissDe   := u ) },oPanel01, 060, 010, "@D",, 0, 16777215,,.F.,,.T.,,.F.,,.F.,.F.,,.F.,.F. ,,"dEmissDe",,,,.T.  )
	
	dEmissAte		:= dDataBase
    oSay1			:= TSay():create(oPanel01, {||  'Emissao Final' },050,075,,,,,,.T.,,,200,20)
	oGetDataDe		:= TGet():New( 058, 075, { | u | If( PCount() == 0, dEmissAte , dEmissAte   := u ) },oPanel01, 060, 010, "@D",, 0, 16777215,,.F.,,.T.,,.F.,,.F.,.F.,,.F.,.F. ,,"dEmissAte",,,,.T.  )
	
	lCtes 		:= .T.    
	lServico 	:= .F.    
 	lVenda 		:= .F.    
 	lOutros 	:= .F.    
   
   	oCheckCte := TCheckBox():New(58,140,'Conhecimentos',{|| lCtes	 },oPanel01,100,210,,{|| lCtes := !lCtes},,,,,,.T.,,,)
	oCheckSer := TCheckBox():New(58,192,'NF de Serviço',{|| lServico },oPanel01,100,210,,{|| lServico := !lServico},,,,,,.T.,,,)
   	oCheclVen := TCheckBox():New(58,245,'NF de Venda'  ,{|| lVenda	 },oPanel01,100,210,,{|| lVenda := !lVenda},,,,,,.T.,,,)
   	oCheclVen := TCheckBox():New(68,140,'Outros'  		,{|| lOutros	 },oPanel01,100,210,,{|| lOutros := !lOutros},,,,,,.T.,,,)

	oBtnFiltro:= TButton():New( 58, 295, "Filtrar",oPanel01,{|| Filtrar()}, 40,12,,,.F.,.T.,.F.,,.F.,,,.F. )   
	oBtnFiltro:= TButton():New( 010, 295, "Buscar Doc",oPanel01,{|| PesqDoc()}, 40,12,,,.F.,.T.,.F.,,.F.,,,.F. )   
	

	//Janela dos dados de geracao de fatura
	oLayer:AddCollumn('col02', 40, .T., 'lin01')	
	oLayer:addWindow("col02","win02",'Dados para Fatura',100,.F.,.F.,{|| }, "lin01",{|| })
	oPanel02 := oLayer:getWinPanel('col02', 'win02', "lin01")

	oSay1		:= TSay():create(oPanel02, {||  'Cedente' },02,05,,,,,,.T.,,,200,20)	
	cCedente	:= FWFilialName() 
	oGetCedente	:= TGet():Create( oPanel02,{||cCedente},10,05,155,10,"@!",,0,,,.F.,,.T.,,.F.,,.F.,.F.,,.T.,.F.,,cCedente,,,, )    
	oGetCedente:lActive := .F.	
    
	aTipoTitulo		:= {'Boleto(BOL)','Fatura(FT)','Deposito(TF)'}
	cCmbTipo 		:= aTipoTitulo[2]
    oSay1			:= TSay():create(oPanel02, {||  'Tipo de Titulo' },02,170,,,,,,.T.,,,200,20)
	oComboTipo		:= TComboBox():New(10,170,{|u|if(PCount()>0,cCmbTipo:=u,cCmbTipo)},aTipoTitulo,50,500,oPanel02,,{|| .T.},,,,.T.,,,,,,,,,'cCmbTipo')

	dVencto			:= dDataBase
    oSay1			:= TSay():create(oPanel02, {||  'Vencimento' },27,75,,,,,,.T.,,,200,20)
	oGetDataDe		:= TGet():New( 35, 75, { | u | If( PCount() == 0, dVencto , dVencto   := u ) },oPanel02, 50, 010, "@D",, 0, 16777215,,.F.,,.T.,,.F.,,.F.,.F.,,.F.,.F. ,,"dVencto",,,,.T.  )

    cNatureza		:= PADR(cNatureza,Len(SED->ED_CODIGO))
	// oSay1			:= TSay():create(oPanel02, {||  'Natureza' },27,55,,,,,,.T.,,,200,20)
	// oGetNat			:= TGet():New( 35, 55, { | u | If( PCount() == 0, cNatureza, cNatureza := u ) },oPanel02, 050, 010, "@!",{|| Validar('natureza')}, 0, 16777215,,.F.,,.T.,,.F.,,.F.,.F.,,.F.,.F. ,"SED","cNatureza",,,,.T.  )
	
	dDtEmissa		:= dDataBase
    oSay1			:= TSay():create(oPanel02, {||  'Emissao' },27,05,,,,,,.T.,,,200,20)
	oGetDataDe		:= TGet():New( 35, 05, { | u | If( PCount() == 0, dDtEmissa , dDtEmissa   := u ) },oPanel02, 50, 010, "@D",, 0, 16777215,,.F.,,.T.,,.F.,,.F.,.F.,,.F.,.F. ,,"dDtEmissa",,,,.T.  )



	oFont := TFont():New('Courier new',,-12,.T.)
	oSay1 := TSay():New(60,05,{||'Total Titulos: ' + "R$ " + Alltrim(Transform(nTotVlr, "@e 999,999,999.99"))+' - Total em saldo: ' + "R$ " + Alltrim(Transform(nTotFat, "@e 999,999,999.99")) +" - Qtd."+Alltrim(Transform(nTotCTE, "@e 99999999999999")) },oPanel02,,oFont,,,,.T.,IIF(nTotVlr != nTotFat, CLR_RED,CLR_BLACK),CLR_WHITE,200,20)
	// oSay1 := TSay():New(60,130,{||'Qtd. CTE: ' },oPanel02,,oFont,,,,.T.,CLR_BLACK,CLR_WHITE,200,20)
	oFont:Bold := .T.
	
	oFontTotal := TFont():New('Courier new',,-16,.T.)
	oFontTotal:Bold := .T.
	// oSayTotal  := TSay():New(60,68,{|| "R$ " + Alltrim(Transform(nTotFat, "@e 999,999,999.99")) },oPanel02,,oFontTotal,,,,.T.,CLR_BLUE,CLR_WHITE,200,20)
	// oSayTotal  := TSay():New(60,140,{|| "   " + Alltrim(Transform(nTotCTE, "@e 99999999999999")) },oPanel02,,oFontTotal,,,,.T.,CLR_BLUE,CLR_WHITE,200,20)


	//grid com as faturas temporarias
	oLayer:addLine("lin02",55,.T.)		
	oLayer:AddCollumn('col03', 30, .T., 'lin02')	
	oLayer:addWindow("col03","win03",'Clientes ',100,.F.,.F.,{|| }, "lin02",{|| })
	oPanel03 := oLayer:getWinPanel('col03', 'win03', "lin02")
	
	//retorna os dados
	aAux 	:= retDados()
	
	//Cria array com os campo de marcacao
	aMark	:= {{ "{|| if(self:oData:aArray[self:nAt, 1] == cMark, 'LBOK','LBNO') }", "{|| u_LAUA004B(.F.) }", "{|| u_LAUA004B(.T.)  }" }}	
	oBrowseSE1 	:= uFwBrowse():create(oPanel03,,aAux[2],,aAux[1],,aMark,, .F.)	

	oBrowseSE1:disableConfig()
	oBrowseSE1:disableFilter()
	oBrowseSE1:disableReport()
	oBrowseSE1:Activate()
	
	oLayer:AddCollumn('col04', 70, .T., 'lin02')	
	oLayer:addWindow("col04","win04",'Faturas Temporarias - Pressione F5 para MARCAR/DESMARCAR todos.',100,.F.,.F.,{|| }, "lin02",{|| })
	oPanel04 := oLayer:getWinPanel('col04', 'win04', "lin02")
	
	//Cria array com os campo de marcacao
	aMark1	:= {{ "{|| if(self:oData:aArray[self:nAt, 1] == cMark, 'LBOK','LBNO') }", "{|| u_LAUA004D(.F.) }", "{|| u_LAUA004D(.T.)  }" }}	
	aAux2 := retTits()
	oBrowseTIT 	:= uFwBrowse():create(oPanel04,,aAux2[2],,aAux2[1],,aMark1,, .F.)	
	
	oBrowseTIT:disableConfig()
	oBrowseTIT:disableFilter()
	oBrowseTIT:disableReport()
	oBrowseTIT:Activate()
	oBrowseSE1:SetChange({|| AtuTit() })
	SetaAtalhos()
	oDlg:AddButton('Fechar'				, { || oDlg:Deactivate() }					, 'Sair', , .T., .T., .T., ) 
	oDlg:AddButton('Confirmar'			, { || MsAguarde( { || Confirmar(@oDlg) }	, "Gerando faturas","Aguarde finalização do processamento." )}	, 'Confirmar', 	, .T., .T., .T., ) 
	oDlg:AddButton('Pré-fatura'			, { || MSAguarde( { || u_LAUA004E()}		, "Processando" ,"Exportando Pré-fatura", .F.) }				, 'Pré-fatura', , .T., .T., .T.,  )
	oDlg:AddButton('Dactes'				, { || MSAguarde( { || ExporXML(.F.,.T.)}	, "Processando" ,"Exportando DACTES"	, .F.) }				, 'Dactes', 	, .T., .T., .T.,  )
	oDlg:AddButton('XML'				, { || MSAguarde( { || ExporXML(.T.,.F.)}	, "Processando" ,"Exportando XML"		, .F.) }				, 'XML', , .T.	, .T., .T.,  )
	oDlg:AddButton('FATURAS INDIVIDUAIS'				, { || MSAguarde( { || FatuInd(@oDlg)}	, "Processando" ,"Gerando individuais."		, .F.) }				, 'FATURAS INDIVIDUAIS', , .T.	, .T., .T.,  )
	oDlg:Activate()	
	
	dDataBase := dDtAtu
Return

//-------------------------------------------------------------------
/*/{Protheus.doc} Valida
Valida o preenchimento dos campos
@author  Sidney Sales
@since   05/12/19
@version 1.0
/*/
//-------------------------------------------------------------------
Static Function Validar(cTipo)
	Local lRet := .T.
	
	//Valida o campo de grupo
	If cTipo == 'grupo'
		ACY->(DbSetOrder(1))
		If ACY->(DbSeek(xFilial('ACY') + cGrupo))
			cDescGrupo := ACY->ACY_DESCRI
		ElseIf !Empty(cGrupo)
			lRet := .F.
			MsgAlert('Grupo inválido.')
		EndIf
	//Valida o campo cliente
	ElseIf cTipo == 'cliente'
		SA1->(DbSetOrder(1))
		If SA1->(DbSeek(xFilial('SA1') + cCliente))
			cNomeCliente := SA1->A1_NOME
		ElseIf !Empty(cCliente)
			lRet := .F.
			MsgAlert('Cliente inválido')
		EndIf
	ElseIf cTipo == 'loja'
		SA1->(DbSetOrder(1))
		If SA1->(DbSeek(xFilial('SA1') + cCliente +cLojaCli ))
			cNomeCliente := SA1->A1_NOME
		ElseIf !Empty(cCliente)
			lRet := .F.
			MsgAlert('Cliente inválido')
		EndIf
	//Valida o cmapo de natureza
	ElseIf cTipo == 'natureza'
		SED->(DbSetOrder(1))
		If SED->(DbSeek(xFilial('SED') + cNatureza ))
			cDescNat := SED->ED_DESCRIC
		ElseIf !Empty(cDescNat)			
			lRet := .F.
			MsgAlert('Natureza inválida')
		EndIf
	EndIf
Return lRet


//-------------------------------------------------------------------
/*/{Protheus.doc} Confirmar
Funcao que faz a confirmacao, geracao das faturas
@author  Sidney Sales	
@since   05/12/19
@version 1.0
/*/
//-------------------------------------------------------------------
Static Function Confirmar(oDlg)
	Local i
	Local j
	Local nQtdFat	:= 0
	Local cFilAux := ""
	Private _cFilZA7 := cFilAnt
	Private dDataAtu := dDtEmissa
	
	dDataBase := dDataAtu	
	If Empty(cNatureza)
		MsgAlert('Por favor, preencha a natureza.')
		Return .T.
	EndIf

	If dVencto < dDataBase
		MsgAlert('O vencimento não pode ser menor que a data atual.')
		Return .T.
	EndIf
	
	If aScan(oBrowseSE1:oData:aArray, {|x| x[1] == cMark}) == 0
		MsgAlert('Selecione pelo menos uma pre fatura para ser gerada.')
		Return .T.
	EndIf

	If MsgYesNo('Confirma a geração das faturas selecionadas?')

		cCodZA7 := GetSxeNum('ZA7', 'ZA7_CODIGO')
		ConfirmSX8()

		//Cria o registro da ZA7, sera o agrupamento das faturas		
		RecLock('ZA7', .T.)
			ZA7->ZA7_FILIAL := xFilial('ZA7')
			ZA7->ZA7_CODIGO	:= cCodZA7
			ZA7->ZA7_DATA	:= dDataBase
			ZA7->ZA7_HORA	:= Time()				
			ZA7->ZA7_STATUS	:= "P"
			ZA7->ZA7_GRUPO	:= cGrupo
			ZA7->ZA7_CODSA1	:= IIF(Empty(cGrupo),cCliente,"")
			ZA7->ZA7_LOJA	:= IIF(Empty(cGrupo),cLojaCli,"")
		ZA7->(MsUnLock())
		cFilAux := cFilAnt
		
		//Percorre a grid das faturas que serao geradas por cliente
		For i := 1 to Len(oBrowseSE1:oData:aArray)
			cFilant := _cFilZA7
			dDataBase := dDataAtu 
			//Verifica se esta marcada
			If oBrowseSE1:oData:aArray[i, 1] == cMark

				//Pega o cliente atual
				cCliLoja := oBrowseSE1:oData:aArray[i, 2]
				
				//Verifica a possicao dele no array de RECNOS de SE1's
				nPosCli  := aScan(aRecnosSE1, {|x| Alltrim(x[1]) == cCliLoja })
				aAuxSE1	 := aRecnosSE1[nPosCli][2]
				nVlrFat	 := 0
				cNumFat := GetNumFat()
				_cMsgDate := ""
				//Percorre os titulos que foram agrupados para o cliente 
				For j := 1 to Len(aAuxSE1)
					
					//Posiciona o titulo e monta o filtro
					SE1->(DbGoTo(aAuxSE1[j]))								
					If Alltrim(SE1->E1_YFATSEQ) == Alltrim(cNumFat)
						loop
					EndIf

					RecLock('SE1', .F.)
						SE1->E1_YFATSEQ := cNumFat
					SE1->(MsUnLock())

					nVlrFat += SE1->E1_SALDO - RetImp()

					If SE1->E1_VENCTO <= Date()
						RecLock('SE1', .F.)
							SE1->E1_VENCTO 	:= SE1->E1_VENCTO 	+ 365
							SE1->E1_VENCREA := SE1->E1_VENCREA 	+ 365
						SE1->(MsUnLock())
					EndIf
				Next

				cFiltro := " E1_YFATSEQ = '"+cNumFat+"' "
				
				//Chama a funcao que ira incluir a liquidacao(fatura)
				lRet := liquidar(cFiltro, StrTran(cCliLoja, '/',''), nVlrFat, cCodZA7, _cFilZA7,cNumFat)
				

				//Caso ocorra algum erro, ira faturar apenas o que conseguiu, pergunta se continua ou se interrompe(mas na cancela nda. )
				If ! lRet 
					If MsgYesNo('Foram encontrados erros na geração da fatura para o cliente ' + Alltrim(SA1->A1_NOME) + ' deseja parar o processamento?')
						Exit
					EndIf
				Else
					nQtdFat++
				EndIf

			EndIf
		Next
		cFilAnt := cFilAux
		//Se tiver faturado algo, se nao, apaga o registro(ZA7)
		If nQtdFat > 0
			ApMsgInfo('Processamento Concluído. Faturas geradas: ' + cValToChar(nQtdFat))
			RecLock('ZA7', .F.)
				ZA7->ZA7_STATUS := 'G' // Gerado
			ZA7->(MsUnLock())
		Else
			MsgAlert('Nenhuma fatura foi gerada.')
			RecLock('ZA7', .F.)
				ZA7->(DbDelete())
			ZA7->(MsUnLock())
		EndIf
	Else
		MsgAlert('Operação cancelada pelo usuário.')
	EndIf

	oDlg:Deactivate()

Return .T.


//-------------------------------------------------------------------
/*/{Protheus.doc} liquidar
Chama funcao que faz a liquidacao(geracao da fatura)
@author  Sidney Sales
@since   05/12/19
@version 1.0
/*/
//-------------------------------------------------------------------
Static Function liquidar(cFiltro, cCliLoja, nVlrFat, cCodZA7, cFilAux,cFatSeq)
	Local aItens := {}
	Local aCabec := {}
	Local nMoeda := 1
	Local lRet   := .T.
	Local cNumSE1 := ""
	Local cPrefixo 	:= Alltrim(SuperGetMV('MS_PREFFAT', .F. ,'FAT'))
	Local cTipo		:= Alltrim(SuperGetMV('MS_TIPOFAT', .F. ,'FT' ))

	Private lMsErroAuto := .F.
	Default cFilAux := ""
	Default cFatSeq := ""

	//Tipo que sera gerado a fatura
	If cCmbTipo == 'Boleto(BOL)'
		cTipo 	:= 'BOL'
	ElseIf cCmbTipo =='Fatura(FT)'
		cTipo	:= 'FT'	
	ElseIf cCmbTipo == 'Deposito(TF)'
		cTipo	:= 'TF'
	EndIf
	
	//Posiciona o cliente
	SA1->(DbSeek(xFilial('SA1') + cCliLoja))
	
	//Pega o número do titulo
	cNumSE1	:= GetSxeNum('SE1','E1_NUM')
	ConfirmSX8()
	
	//Fatura que sera gerada
	Aadd(aItens,{;
		{"E1_FILIAL"    , cFilAux		},;
		{"E1_PREFIXO"   , cPrefixo     },;
		{'E1_NUM' 		, cNumSE1      },;
		{'E1_PARCELA' 	, '001'  	   },;
		{'E1_VENCTO' 	, dVencto      },;
		{'E1_VALOR' 	, nVlrFat      },;
		{'E1_VLCRUZ' 	, nVlrFat      }})

	//Cabecalho da fatura
	aCabec :={  {"cCondicao"    ,""   	  		},;
				{"cNatureza"    ,cNatureza   	},;
				{"E1_TIPO"      ,cTipo	 	 	},;
				{"cCliente"     ,SA1->A1_COD    },;
				{"nMoeda"       ,nMoeda     	},;
				{"FO0_YZA7"     ,cCodZA7     	},;
				{"cLoja"        ,SA1->A1_LOJA   }}
	aParcelas := Condicao( nVlrFat, "001",, dDataBase)
	//Executa a rotina automatica
	BEGIN TRANSACTION
		MSExecAuTo({|x,y,k,w,z| Fina460(x,y,k,w,z)}, ,aCabec , aItens , 3, cFiltro)  	

		//Verifica se houveram erros
		If lMsErroAuto
			MostraErro()
			lRet := .F.
		Else
			If !VldValor(cFatSeq)
				ApMsgInfo("Houve divergência nos valores gerados por essa fatura. Contate um analista.")
				DisarmTransaction()
				Return .F.
			EndIf
			PopYNLiq(FO0->FO0_NUMLIQ)
			u_LAUA004Z(FO0->FO0_NUMLIQ)
			
		EndIf
	END TRANSACTION
Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} LAUA004B
Funcaopara marcacao e desmarcacao da grid
@author  Sidney Sales
@since   05/12/19
@version 1.0
/*/
//-------------------------------------------------------------------
User Function LAUA004B(lTodos)

	Local i		:= 1
	Local aArray:= oBrowseSE1:oData:aArray 
	Local nAt	:= oBrowseSE1:nAt
	
	If lTodos
		For i := 1 to Len(aArray)
			if aArray[i, 1] == cMark
				aArray[i, 1] 	:= '  '
				aRecnosSE1 		:= {}
			else
				Filtrar()
				exit
			EndIf
		Next
		AtuTit()
	Else
		if aArray[oBrowseSE1:nAt, 1] == cMarK
			RemoTodos()
			aArray[oBrowseSE1:nAt, 1] := '  '
			AtuTit()
		else
			AdiTodos()
			aArray[oBrowseSE1:nAt, 1] := cMark	
			AtuTit()	
		EndIf           
	EndIf
	AtuTit()
	AtuTotais()
	oBrowseSE1:Refresh()
	
	//Atualiza o campo de totais
	

Return

//-------------------------------------------------------------------
/*/{Protheus.doc} retDados
Funcao que retorna os dados para a grid
@author  Sidney Sales
@since   05/12/19
@version 1.0
/*/
//-------------------------------------------------------------------

Static Function retDados()
	Local cQuery 	
	Local aHeader	:= {}
	Local aCols		:= {}	
		
	//Cabecalho da grid de titulos
	aAdd(aHeader,{''			,10, ''				})
	aAdd(aHeader,{'A1_COD'		,10, 'A1_COD'		})
	aAdd(aHeader,{'A1_CGC'		,10, 'A1_CGC'		})
	aAdd(aHeader,{'A1_NOME'		,30, 'A1_NOME'		})
	aAdd(aHeader,{'E1_VALOR'	,10, 'E1_VALOR'		})
	aAdd(aHeader,{'QTD'			,10, ''				})
	
	
	If !Empty(cGrupo) .OR. !Empty(cCliente)

		cQuery := " SELECT DISTINCT SE1.R_E_C_N_O_ AS RECNOSE1, E1_FILIAL, E1_SALDO, E1_VALOR, A1_NOME, E1_CLIENTE, E1_LOJA, A1_CGC FROM " + RetSqlTab('SE1')
		If !lOutros
			cQuery += " LEFT JOIN " + RetSqlTab('SC5') + " ON C5_FILIAL = E1_FILIAL AND C5_NUM = E1_PEDIDO AND SC5.D_E_L_E_T_ = SE1.D_E_L_E_T_ "
		EndIf
		cQuery += " INNER JOIN " + RetSqlTab('SA1') + " ON A1_FILIAL = LEFT(E1_FILIAL,2) AND E1_CLIENTE = A1_COD AND E1_LOJA = A1_LOJA AND SA1.D_E_L_E_T_ = SE1.D_E_L_E_T_ "
		cQuery += " WHERE SE1.D_E_L_E_T_ = '' AND SA1.D_E_L_E_T_ = '' "
		cQuery += " AND E1_EMISSAO BETWEEN '" + DtoS(dEmissDe) + "' AND '" + DtoS(dEmissAte) + "' "
		cQuery += " AND E1_SALDO > 0 "
		cQuery += " AND E1_BAIXA = '' "
		cQuery += " AND LEFT(E1_FILIAL,2) = '" + LEFT(xFilial('SE1'), 2) + "' "

		If !Empty(cGrupo)
			cQuery += " AND A1_GRPVEN = '" + PADR(cGrupo, len(SA1->A1_GRPVEN))+ "' "
		ElseIf !Empty(cCliente)
			cQuery += " AND A1_COD = '" + PADR(cCliente,len(SA1->A1_COD)) + "' "
			cQuery += " AND A1_LOJA = '" + PADR(cLojaCli,len(SA1->A1_LOJA)) + "' "
		EndIf

		If !Empty(cTpFrete) .AND. !lOutros
			cQuery += " AND C5_TPFRETE = '" + LEFT(cTpFrete, 1) + "' "
		EndIf
		
		cQuery += " AND E1_NUMLIQ = '' AND E1_NUMBOR = '' " //Titulos que nao foram liquidados
		
		cTipos := ''

		If lCtes
			cTipos := "'CTE'"
		EndIf

		If lVenda
			If Empty(cTipos)
				cTipos := "'NF '"
			Else
				cTipos += ",'NF '"
			EndIf
		EndIf

		If lServico
			If Empty(cTipos)
				cTipos := "'NFS'"
			Else
				cTipos += ",'NFS'"
			EndIf
		EndIf
		If  !lOutros
			cQuery += " AND E1_TIPO IN ("+ cTipos+ ")" 		
		Else
			cTpAux := "'AB-','FB-','FC-','FU-','IR-','IN-','IS-','PI-','CF-','CS-','FE-','IV-'"
			If !('NF ' $ cTipos)
				cTpAux += IIF(Empty(cTpAux),"'NF '",",'NF '")
			EndIf
			If !('NFS' $ cTipos)
				cTpAux += IIF(Empty(cTpAux),"'NFS'",",'NFS'")
			EndIf
			If !('CTE' $ cTipos)
				cTpAux += IIF(Empty(cTpAux),"'CTE'",",'CTE'")
			EndIf
			If !Empty(cTpAux)
				cQuery += " AND E1_TIPO NOT IN ("+cTpAux+")" 		
			EndIf
		EndIf

	
		cQuery += " ORDER BY E1_FILIAL, E1_CLIENTE, E1_LOJA "

		If Select('QRY') > 0
			QRY->(DbCloseArea())
		EndIf
		
		TcQuery cQuery New Alias 'QRY'

		aRecnosSE1 	:= {}
		cCliAtu		:= ''
		nValorCli	:= 0
		nTotFat		:= 0

		//Preenche o acols com os dados da query
		While QRY->(!Eof())
			cCliAtu	:= QRY->( E1_CLIENTE + '/' + E1_LOJA)
			If len(aRecnosSE1) > 0
				nPosAux := aScan(aRecnosSE1,{|x| AllTrim(x[1]) == cCliAtu })
				If nPosAux > 0
					aAdd(aRecnosSE1[nPosAux][2], QRY->RECNOSE1)	
					aRecnosSE1[nPosAux][3]	+= 	QRY->E1_SALDO
					aRecnosSE1[nPosAux][4]	+= 	QRY->E1_VALOR
					aCols[nPosAux][5] += QRY->E1_SALDO	
				Else
					
					aAdd(aRecnosSE1, { cCliAtu, {QRY->RECNOSE1}, QRY->E1_SALDO, QRY->E1_VALOR })						
					aAdd(aCols, {cMark, cCliAtu, QRY->A1_CGC, QRY->A1_NOME, QRY->E1_SALDO, 0 })
				EndIf
			Else	
				aAdd(aRecnosSE1, { cCliAtu, {QRY->RECNOSE1}, QRY->E1_SALDO, QRY->E1_VALOR })						
				aAdd(aCols, {cMark, cCliAtu, QRY->A1_CGC, QRY->A1_NOME, QRY->E1_SALDO, 0 })
			EndIf
			SE1->(DbGoTo(QRY->RECNOSE1))
			If SE1->E1_VALOR != SE1->E1_SALDO
				cTextoTit += "Titulo "+SE1->E1_NUM+"/"+SE1->E1_PREFIXO+" da filial "+ SE1->E1_FILIAL+ " e cliente "+SE1->(E1_CLIENTE +"/"+SE1->E1_LOJA)+" está com valor diferente do saldo."
				cTextoTit += "Saldo = "+Alltrim(Transform(SE1->E1_SALDO,"@R 999,999,999.99"))+" - Valor ="+Alltrim(Transform(SE1->E1_VALOR,"@R 999,999,999.99"))+CRLF
			EndIf
			
			QRY->(DbSkip())
		EndDo
	EndIf

	//Caso nao tenha dados no select, preencho uma linha vaiza
	If Empty(aCols)
		aAdd(aCols, {'  ', '', '', '', 0 , 0 })
	EndIf

Return {aHeader, aCols }

//-------------------------------------------------------------------
/*/{Protheus.doc} filtrar
Funcao que executa o filtro
@author  Sidney Sales
@since   05/12/19
@version 1.0
/*/
//-------------------------------------------------------------------
Static Function Filtrar
	Local aAux
	Local cCliAtual := ""
	//Pelo menos um tipo tem que ser selecionado
	If lCtes .OR. lVenda .OR. lServico .OR. lOutros
	
		MSAguarde( { || aAux := retDados()}, "Processando" ,"Aguarde busca dos CTES", .F.) 
		cCliAtual := StrTran(aAux[2][1][2],"/","")
		aAux2 	:= retTits(cCliAtual)
		
		oBrowseSE1:oData:aArray := aAux[2]
		oBrowseSE1:goTop()
		oBrowseSE1:gobottom()
		oBrowseSE1:Refresh()	
		oBrowseSE1:gobottom()
		oBrowseSE1:goTop()	

		MSAguarde( { ||  AtuTit()}, "Processando" ,"Aguarde busca dos CTES", .F.) 
		AtuTotais()
		
		If !Empty(cTextoTit)
			Aviso("Titulos com saldo e valor inconsistente.",cTextoTit, ,3)
		EndIf
		cTextoTit := ""
		oBrowseSE1:Refresh()
	Else
		MsgAlert('Selecione pelo menos um tipo de documento!')
	EndIf

Return

//-------------------------------------------------------------------
/*/{Protheus.doc} AtuTit
Atualiza grid de titulos
@author  Samuel Dantas
@since   27/12/2019
@version version
/*/
//-------------------------------------------------------------------
Static Function AtuTit()
	aAux 	:= oBrowseSE1:oData:aArray 
	nAt		:= oBrowseSE1:nAt
	cCliAtual := StrTran(aAux[nAt][2],"/","")
	aAux2 	:= retTits(cCliAtual)
	
	oBrowseTIT:oData:aArray := aAux2[2]
	oBrowseTIT:goTop()
	oBrowseTIT:gobottom()
	oBrowseTIT:Refresh()	
	oBrowseTIT:gobottom()
	oBrowseTIT:goTop()	
	oBrowseSE1:Refresh()
Return

//-------------------------------------------------------------------
/*/{Protheus.doc} AtuTotais
Funcao que atualiza os campos de totais
@author  Sidney Sales
@since   05/12/19
@version 1.0
/*/
//-------------------------------------------------------------------
Static Function AtuTotais
	Local aArray	:= oBrowseSE1:oData:aArray 
	Local aTits		:= oBrowseTIT:oData:aArray 
	Local nAt		:= oBrowseSE1:nAt
	Local nAtTit	:= oBrowseTIT:nAt
	Local i			:= 1

	nTotFat := 0
	nTotCTE := 0
	nTotVlr := 0
	
		//Percorre a grid
	If len(aRecnosSE1) > 0
		For i := 1 to Len(aRecnosSE1)
			nPos := aScan(aArray,{|x| AllTrim(x[2]) == aRecnosSE1[i][1]})
			If nPos > 0
				aArray[nPos][5] :=  aRecnosSE1[i][3]
				aArray[nPos][6] :=  len(aRecnosSE1[i][2])
			EndIf
			nTotFat += aRecnosSE1[i][3]
			nTotCTE += len(aRecnosSE1[i][2])
			nTotVlr += aRecnosSE1[i][4]
			//Verifica se esta marcado  
			// If oBrowseSE1:oData:aArray[i][1] == cMark	
			// 	nTotFat += oBrowseSE1:oData:aArray[i][5]
			// EndIf
		Next
		oBrowseSE1:oData:aArray := aArray
	EndIf
Return

//-------------------------------------------------------------------
/*/{Protheus.doc} LAUA004c
Funcao que abre a tela com as faturas
@author  Sidney Sales
@since   05/12/19
@version 1.0
/*/
//-------------------------------------------------------------------
User Function LAUA004c
	Local oBrowse := FWmBrowse():New()
	Local cFiltro := "FO0_FILIAL == ZA7->ZA7_FILIAL .AND. FO0->FO0_YZA7 == ZA7->ZA7_CODIGO"
	Private aRotina := {}
	Private LOPCAUTO := .T.
	//Filtra
	FO0->(DbSetFilter({|| &(cFiltro) }, cFiltro))
	
	oBrowse:SetAlias("FO0")
	oBrowse:SetDescription("Faturas geradas")   
	
	oBrowse:AddLegend( "FO0_STATUS=='2'"										,"RED"		,	OemToAnsi("Bloqueada") )	// "Bloqueada"
	oBrowse:AddLegend( "FO0_STATUS=='1' .AND. FO0_DTVALI >= dDatabase"			,"GREEN"	,	OemToAnsi("Vigente") )	// "Vigente"
	oBrowse:AddLegend( "FO0_STATUS=='3'"										,"BLUE"		,	OemToAnsi("Cancelada") )	// "Cancelada"
	oBrowse:AddLegend( "FO0_STATUS=='1' .AND. FO0_DTVALI < dDatabase"			,"YELLOW"	,	OemToAnsi("Vencida") )	// "Vencida"
	oBrowse:AddLegend( "FO0_STATUS=='4'"										,"WHITE"	,	OemToAnsi("Gerada") )	// "Gerada"
	oBrowse:AddLegend( "FO0_STATUS=='5'"										,"BLACK"	,	OemToAnsi("Encerrada") )	// "Encerrada"

	oBrowse:SetMenuDef('FINA460A')
	
	oBrowse:Activate()

	//Limpa o filtro
	FO0->(DBClearFilter())	

Return

//-------------------------------------------------------------------
/*/{Protheus.doc} retTits
Funcao que retorna os dados para a grid
@author  Samuel Dantas
@since   05/12/19
@version 1.0
/*/
//-------------------------------------------------------------------

Static Function retTits(cCliAtual)
	Local cQuery 	
	Local aHeader	:= {}
	Local aCols		:= {}	
	Default cCliAtual := ""	
	//Cabecalho da grid de titulos
	aAdd(aHeader,{''				,10, ''				})
	aAdd(aHeader,{'E1_FILIAL'		,10, 'E1_FILIAL'	})
	aAdd(aHeader,{'E1_PREFIXO'		,4 , 'E1_PREFIXO'	})
	aAdd(aHeader,{'E1_NUM'			,10, 'E1_NUM'		})
	aAdd(aHeader,{'E1_EMISSAO'		,10, 'E1_EMISSAO'	})
	aAdd(aHeader,{'E1_VENCTO'		,10, 'E1_VENCTO'	})
	aAdd(aHeader,{'E1_VENCREA'		,10, 'E1_VENCREA'	})
	aAdd(aHeader,{'E1_SALDO'		,10, 'E1_SALDO'		})
	aAdd(aHeader,{'E1_VALOR'		,10, 'E1_VALOR'		})
	aAdd(aHeader,{'E1_CLIENTE'	  	,10, 'E1_CLIENTE'	})
	aAdd(aHeader,{'E1_LOJA'		  	,10, 'E1_LOJA'		})
	aAdd(aHeader,{'IDENTIFICADOR' 	,10, ''		})
	
	
	If !Empty(cCliAtual)
		cQuery := " SELECT DISTINCT SE1.R_E_C_N_O_ AS RECNOSE1, E1_FILIAL, E1_PREFIXO, E1_NUM, E1_EMISSAO, E1_VENCTO, E1_VENCREA, E1_SALDO, E1_VALOR, E1_CLIENTE, E1_LOJA FROM " + RetSqlTab('SE1')
		If !lOutros
			cQuery += " LEFT JOIN " + RetSqlTab('SC5') + " ON C5_FILIAL = E1_FILIAL AND C5_NUM = E1_PEDIDO AND SC5.D_E_L_E_T_ = SE1.D_E_L_E_T_ "
		EndIf
		cQuery += " INNER JOIN " + RetSqlTab('SA1') + " ON A1_FILIAL = LEFT(E1_FILIAL,2) AND E1_CLIENTE = A1_COD AND E1_LOJA = A1_LOJA AND SA1.D_E_L_E_T_ = SE1.D_E_L_E_T_ "
		cQuery += " WHERE SE1.D_E_L_E_T_ = '' "
		cQuery += " AND E1_EMISSAO BETWEEN '" + DtoS(dEmissDe) + "' AND '" + DtoS(dEmissAte) + "' "
		cQuery += " AND E1_SALDO > 0 "
		cQuery += " AND E1_BAIXA = '' "
		cQuery += " AND LEFT(E1_FILIAL,2) = '" + LEFT(xFilial('SE1'), 2) + "' "

		cQuery += " AND E1_CLIENTE+E1_LOJA = '" +cCliAtual+ "' "	
		
		If !Empty(cTpFrete) .AND. !lOutros
			cQuery += " AND C5_TPFRETE = '" + LEFT(cTpFrete, 1) + "' "
		EndIf
		
		cQuery += " AND E1_NUMLIQ = '' AND E1_NUMBOR = '' " //Titulos que nao foram liquidados
		
		cTipos := ''

		If lCtes
			cTipos := "'CTE'"
		EndIf

		If lVenda
			If Empty(cTipos)
				cTipos := "'NF '"
			Else
				cTipos += ",'NF '"
			EndIf
		EndIf

		If lServico
			If Empty(cTipos)
				cTipos := "'NFS'"
			Else
				cTipos += ",'NFS'"
			EndIf
		EndIf
		If  !lOutros
			cQuery += " AND E1_TIPO IN ("+ cTipos+ ")" 		
		Else
			cTpAux := "'AB-','FB-','FC-','FU-','IR-','IN-','IS-','PI-','CF-','CS-','FE-','IV-'"
			If !('NF ' $ cTipos)
				cTpAux += IIF(Empty(cTpAux),"'NF '",",'NF '")
			EndIf
			If !('NFS' $ cTipos)
				cTpAux += IIF(Empty(cTpAux),"'NFS'",",'NFS'")
			EndIf
			If !('CTE' $ cTipos)
				cTpAux += IIF(Empty(cTpAux),"'CTE'",",'CTE'")
			EndIf
			If !Empty(cTpAux)
				cQuery += " AND E1_TIPO NOT IN ("+cTpAux+")" 		
			EndIf
		EndIf	

		cQuery += " ORDER BY E1_FILIAL, E1_PREFIXO, E1_NUM, E1_EMISSAO, E1_VENCTO, E1_VENCREA, E1_SALDO, E1_VALOR, E1_CLIENTE, E1_LOJA "

		If Select('QRY') > 0
			QRY->(DbCloseArea())
		EndIf
		
		TcQuery cQuery New Alias 'QRY'

		aAuxCli 	:= oBrowseSE1:oData:aArray 
		nAt			:= oBrowseSE1:nAt
		cCliAtu		:= aAuxCli[nAt][2]

		//Preenche o acols com os dados da query
		While QRY->(!Eof())
			lMarca := .F.
			nPosCli := aScan(aRecnosSE1,{|x| Alltrim(x[1]) == Alltrim(cCliAtu) })
			If nPosCli > 0
				If len(aRecnosSE1) > 0
					aAuxRecnos := aRecnosSE1[nPosCli][2]
					nPos := aScan(aAuxRecnos,{|x| x == QRY->RECNOSE1})
					If nPos > 0
						aAdd(aCols, {cMarK, QRY->E1_FILIAL, QRY->E1_PREFIXO, QRY->E1_NUM, StoD(QRY->E1_EMISSAO), StoD(QRY->E1_VENCTO), StoD(QRY->E1_VENCREA), QRY->E1_SALDO,QRY->E1_VALOR, QRY->E1_CLIENTE, QRY->E1_LOJA, QRY->RECNOSE1 })
						lMarca := .T.
					EndIf
				EndIf
			EndIf
			
			If !lMarca
				aAdd(aCols, {" ", QRY->E1_FILIAL, QRY->E1_PREFIXO, QRY->E1_NUM, StoD(QRY->E1_EMISSAO), StoD(QRY->E1_VENCTO), StoD(QRY->E1_VENCREA), QRY->E1_SALDO,QRY->E1_VALOR, QRY->E1_CLIENTE, QRY->E1_LOJA, QRY->RECNOSE1 })
			EndIf

			QRY->(DbSkip())
		EndDo
	EndIf
	
	//Caso nao tenha dados no select, preencho uma linha vaiza
	If Empty(aCols)
		aAdd(aCols, {'','  ', '', '', CtoD(''), CtoD(''), CtoD(''), 0, 0, "", "",0 })
	EndIf

Return {aHeader, aCols }

//-------------------------------------------------------------------
/*/{Protheus.doc} RemoRecno
description
@author  Samuel Dantas
@since   27/12/2019
@version version
/*/
//-------------------------------------------------------------------
Static Function RemoRecno(cCliAtu, nRecno)
	Local nPosCli 	:= 0
	Local nJ 		:= 1
	Local aRecAux 	:= {}
	Local nPosRec 	:= 0
	Local aDadosSE1 := oBrowseTIT:oData:aArray 

	nPosCli := aScan(aRecnosSE1,{|x| Alltrim(x[1]) == Alltrim(cCliAtu) })
	If nPosCli > 0
		If len(aRecnosSE1) > 0
			aAuxRecnos := aRecnosSE1[nPosCli][2]
			nPosRec := aScan(aAuxRecnos,{|x| x == nRecno})
			
			For nJ := 1 To len(aAuxRecnos)
				If aAuxRecnos[nJ] != nRecno
					aAdd(aRecAux, aAuxRecnos[nJ])
				EndIf
			Next
			If len(aRecAux) > 0
				aRecnosSE1[nPosCli][2] := aRecAux
				aRecnosSE1[nPosCli][3] -= aDadosSE1[oBrowseTIT:nAt][8]
				aRecnosSE1[nPosCli][4] -= aDadosSE1[oBrowseTIT:nAt][9]
			Else
				RemoTodos()
			EndIf
			
		EndIf
	EndIf
Return

//-------------------------------------------------------------------
/*/{Protheus.doc} AdicRecno
description
@author  Samuel Dantas
@since   27/12/2019
@version version
/*/
//-------------------------------------------------------------------
Static Function AdicRecno(cCliAtu, nRecno)
	Local nPosCli 	:= 0
	Local nJ 		:= 1
	Local aRecAux 	:= {}
	Local nPosRec 	:= 0
	Local aDadosSE1 := oBrowseTIT:oData:aArray 

	nPosCli := aScan(aRecnosSE1,{|x| Alltrim(x[1]) == Alltrim(cCliAtu) })
	If nPosCli > 0
		aAdd(aRecnosSE1[nPosCli][2],nRecno)
		aRecnosSE1[nPosCli][3] += aDadosSE1[oBrowseTIT:nAt][8]
		aRecnosSE1[nPosCli][4] += aDadosSE1[oBrowseTIT:nAt][9]
	Else
		aAdd(aRecnosSE1, { cCliAtu, {nRecno}, aDadosSE1[oBrowseTIT:nAt][8],aDadosSE1[oBrowseTIT:nAt][9]  })	
	EndIf
Return

//-------------------------------------------------------------------
/*/{Protheus.doc} RemoTodos
Remove todos os registros referente a um título
@author  Samuel Dantas
@since   27/12/2019
@version version
/*/
//-------------------------------------------------------------------
Static Function RemoTodos(cCliAtu)
	Local nJ	:= 1
	aAuxCli 	:= oBrowseSE1:oData:aArray 
	nAt			:= oBrowseSE1:nAt
	cCliAtu		:= aAuxCli[nAt][2]
	aRAux		:= {}

	nPosCli := aScan(aRecnosSE1,{|x| Alltrim(x[1]) == Alltrim(cCliAtu) })
	If nPosCli > 0
		For nJ := 1 To len(aRecnosSE1)
			If nPosCli != nJ
				aAdd(aRAux,aRecnosSE1[nJ])
			EndIf
		Next
		aRecnosSE1 := aRAux
	EndIf

Return

//-------------------------------------------------------------------
/*/{Protheus.doc} AdiTodos
Adiciona todos titulos de um deteminado cliente
@author  Samuel Dantas
@since   18/02/2020
@version version
/*/
//-------------------------------------------------------------------
Static Function AdiTodos(cCliAtu)
	Local aDadosSE1 := oBrowseTIT:oData:aArray 
	Local nJ		:= 1
	aAuxCli 	:= oBrowseSE1:oData:aArray 
	nAt			:= oBrowseSE1:nAt
	cCliAtu		:= aAuxCli[nAt][2]

	
	For nJ := 1 To len(aDadosSE1)
		nPosCli := aScan(aRecnosSE1,{|x| Alltrim(x[1]) == Alltrim(cCliAtu) })
		If nPosCli == 0
			aAdd(aRecnosSE1, { cCliAtu, {aDadosSE1[nJ][12]}, aDadosSE1[nJ][8],aDadosSE1[nJ][9] })	
		Else		
			aAdd(aRecnosSE1[nPosCli][2], aDadosSE1[nJ][12])
			aRecnosSE1[nPosCli][3] += aDadosSE1[nJ][8]
			aRecnosSE1[nPosCli][4] += aDadosSE1[nJ][9]
		EndIf
		
	Next

Return


//-------------------------------------------------------------------
/*/{Protheus.doc} LAUA004D
Funcao para marcacao e desmarcacao da grid
@author  Sidney Sales
@since   05/12/19
@version 1.0
/*/
//-------------------------------------------------------------------
User Function LAUA004D(lTodos)

	Local aArray	:= oBrowseSE1:oData:aArray 
	Local aTits		:= oBrowseTIT:oData:aArray 
	Local nAt		:= oBrowseSE1:nAt
	Local nAtTit	:= oBrowseTIT:nAt
	cCliAtu		:= aArray[nAt][2]

	If lTodos
		U_LAUA004B(.F.)
	Else
		if aTits[oBrowseTIT:nAt, 1] == cMarK
			RemoRecno(cCliAtu, aTits[nAtTit][12])
			aTits[oBrowseTIT:nAt, 1] := '  '
		else
			AdicRecno(cCliAtu, aTits[nAtTit][12])
			aTits[oBrowseTIT:nAt, 1] := cMark	
			aArray[nAt][1]  := cMarK
		EndIf           
	EndIf
	// AtuTit()
	oBrowseTIT:nAt := nAtTit
	AtuTotais()
	oBrowseTIT:Refresh()
	oBrowseSE1:Refresh()
	
	//Atualiza o campo de totais
	

Return
//-------------------------------------------------------------------
/*/{Protheus.doc} GetNumFat
Busca o ultimo numero disponivel para gerar numero de fatura disponivel
@author  author
@since   date
@version version
/*/
//-------------------------------------------------------------------
Static Function GetNumFat()
	Local cQry := ""
	Local cRet := ""

	cQry := " SELECT MAX(E1_YFATSEQ) as MAX FROM "  + RetSqlName('SE1') + " SE1"
	cQry += " WHERE SE1.D_E_L_E_T_ <> '*' "
	
	If Select('QRYMAX') > 0
		QRYMAX->(dbclosearea())
	EndIf
	
	TcQuery cQry New Alias 'QRYMAX'
	
	If QRYMAX->(!Eof())
		If Empty(QRYMAX->MAX)
			cRet := StrZero(1,TAMSX3("E1_YFATSEQ")[1])
		Else	
			cRet := Soma1(QRYMAX->MAX)
		EndIf
		QRYMAX->(dbSkip())
	Else 
		cRet := StrZero(1,TAMSX3("E1_YFATSEQ")[1])
	EndIf
	 
Return cRet

Static Function SetaAtalhos()
	SetKey(VK_F5, { || u_LAUA004D(.T.) })
	SetKey(VK_F7, { || u_LAUA004D(.F.) })
	
Return

//-------------------------------------------------------------------
/*/{Protheus.doc} LAUF010C
Método responsável pela impressão de pré-fatura(layout de fatura sem boleto)
@author  Samuel Dantas
@since   19/01/2020
@version version
/*/
//-------------------------------------------------------------------

User Function LAUA004E()
    Local cPath       := "c:\temp\"
    Local cDirFinal  := ""
    Local nTipoFat   := 1
    Local aParamBox  := {}
    Local aRet       := {}
    
    Private cStartPath := GetSrvProfString("StartPath","")
    Private aRet          := {}
    Private aRecnos     := {}
    Private nVlrFat     := 0
    Private dVencto     := dDataBase
    Private cLogo       := ""
    Private nRow        := 0
    Private nCols       := 0

    // Pergunta ao usuário o tipo de faturamento
    aAdd(aParamBox,{3,"Tipo de Faturamento", 1, {"Por Loja","Por Cliente"}, 50, "", .T.})
    
    If ParamBox(aParamBox,"Escolha o tipo de faturamento",@aRet)
        nTipoFat := aRet[1]
        
        If !ExistDir(cPath)
            MakeDir(cPath)
        EndIf
        
        cDirFinal := cGetFile("Todos | *.* ", OemToAnsi("Selecione o diretorio"),  ,  ,.F., GETF_LOCALHARD + GETF_RETDIRECTORY, .F.)
        
        If !Empty(cDirFinal)
            If nTipoFat == 1 // Por Loja
                For nX := 1 to len(aRecnosSE1)
                    PreFat(cDirFinal, aRecnosSE1[nX])
                Next
                ApMsgInfo("CSV com CTE foi gravado na pasta: "+cDirFinal+".")
            Else // Por Cliente
                PreFatCli(cDirFinal, aRecnosSE1)
                ApMsgInfo("CSV com CTE foi gravado na pasta: "+cDirFinal+".")
            EndIf
        EndIf
    EndIf
Return

/*------------------------------------
--  Função: Impressão por loja      --
--                                  --
--------------------------------------*/
Static Function PreFat(cDirFinal, aFatura)
    Local aAreaSA1 := SA1->(GetArea())
    Local nVlrTotal := 0
    Local cTexto := ""
    Local cNomRem := ";;"
    Local cNomDest := ";;"
    Local cSeries := Alltrim(SuperGetMv("MS_SERSNFS",.F.,"'NUC','021'"))
    
    Private nPag      := 1
    cTexto += "FILIAL;DOCUMENTO;SERIE;TOMADOR;TIPO;FRETE;CNPJ REM;NOME REM;CNPJ DEST;NOME DEST;EMISSÃO;VALOR;VALOR CARGA"+CRLF
    
    aRecnos := aFatura[2]
    cCliAtual := StrTran(aFatura[1],"/","")
    SA1->(DbSetOrder(1))
    SA1->(DbSeek(xFilial("SA1") + cCliAtual ))
    nVlrTotal := aFatura[3]
    aAreaSA1 := SA1->(GetArea())    
    
    For nZ := 1 To len(aRecnos)
        SE1->(DbGoTo(aRecnos[nZ])) 
        _cCgc := Alltrim(SA1->A1_CGC)
        
        nVlrCarga := 0
        If ! SE1->E1_PREFIXO $ cSeries
            DbSelectArea("ZA4")
            ZA4->(DbSetOrder(4))
            If ZA4->(DbSeek(SE1->E1_FILIAL + SE1->E1_NUM + SE1->E1_PREFIXO + "inclusao"))
                nVlrCarga := ZA4->ZA4_VLRCAR
            EndIf
        EndIf

        If !PosiTabs()
            cTexto += Alltrim(SE1->E1_FILIAL) + ";" + ;
                     Alltrim(SE1->E1_NUM) + ";" + ;
                     Alltrim(SE1->E1_PREFIXO) + ";" + ;
                     Transform(_cCgc, "@R 99.999.999/9999-99") + ";" + ;
                     SE1->E1_TIPO + ";" + ;
                     ";" + ;
                     cNomRem + ";" + ;
                     Alltrim(cNomDest) + ";" + ;
                     DtoC(SE1->E1_EMISSAO) + ";" + ;
                     Transform(SE1->E1_VALOR,"@E 99,999,999.99") + ";" + ;
                     Transform(nVlrCarga,"@E 99,999,999.99") + CRLF
        Else
            cFrete := IIF(SC5->C5_TPFRETE == 'C','CIF','FOB')
            If ! SE1->E1_PREFIXO $ cSeries
                SA1->(dbSetOrder(1))
                SA1->(DbSeek(xFilial("SA1") + DT6->( DT6->(DT6_CLIREM + DT6_LOJREM))))
                cNomRem := Transform( SA1->A1_CGC, "@R 99.999.999/9999-99" )+";"+SA1->A1_NOME
                SA1->(DbSeek(xFilial("SA1") + DT6->( DT6->(DT6_CLIDES + DT6_LOJDES))))
                cNomDest := Transform( SA1->A1_CGC, "@R 99.999.999/9999-99" )+";"+SA1->A1_NOME
            Else
                If PosiZA5()
                    SA1->(dbSetOrder(1))
                    SA1->(DbSeek(xFilial("SA1") + ZA5->( ZA5->(ZA5_REM + ZA5_REMLOJ))))
                    cNomRem := Transform( SA1->A1_CGC, "@R 99.999.999/9999-99" )+";"+SA1->A1_NOME
                    SA1->(DbSeek(xFilial("SA1") + ZA5->( ZA5->(ZA5_DEST + ZA5_DESLOJ))))
                    cNomDest := Transform( SA1->A1_CGC, "@R 99.999.999/9999-99" )+";"+SA1->A1_NOME
                EndIf
            EndIf
            cTexto += Alltrim(SE1->E1_FILIAL) + ";" + ;
                     Alltrim(SE1->E1_NUM) + ";" + ;
                     Alltrim(SE1->E1_PREFIXO) + ";" + ;
                     Transform(_cCgc, "@R 99.999.999/9999-99") + ";" + ;
                     SE1->E1_TIPO + ";" + ;
                     cFrete + ";" + ;
                     cNomRem + ";" + ;
                     Alltrim(cNomDest) + ";" + ;
                     DtoC(SE1->E1_EMISSAO) + ";" + ;
                     Transform(SE1->E1_VALOR,"@E 99,999,999.99") + ";" + ;
                     Transform(nVlrCarga,"@E 99,999,999.99") + CRLF
        EndIf
        SA1->(RestArea(aAreaSA1))
    Next
    cFileCsv := cDirFinal + AllTrim(SA1->A1_CGC)+".csv"
    Memowrite(cFileCsv, cTexto)
Return

/*------------------------------------
--  Função: Impressão por cliente   --
--                                  --
--------------------------------------*/
Static Function PreFatCli(cDirFinal, aFatura)
    Local aAreaSA1  := SA1->(GetArea())
    Local nVlrTotal := 0
    Local cTexto    := ""
    Local cNomRem   := ";;"
    Local cNomDest  := ";;"
    Local cSeries   := Alltrim(SuperGetMv("MS_SERSNFS",.F.,"'NUC','021'"))
    Local cCNPJBase := ""
    Local aClientes := {}
    Local nPos      := 0
    
    // Cabeçalho
    cTexto += "FILIAL;DOCUMENTO;SERIE;TOMADOR;TIPO;FRETE;CNPJ REM;NOME REM;CNPJ DEST;NOME DEST;EMISSÃO;VALOR;VALOR CARGA"+CRLF
    
    // Agrupa títulos por CNPJ base
    For nX := 1 to Len(aFatura)
        cCNPJBase := SubStr(aFatura[nX,1], 1, At("/",aFatura[nX,1])-1)
        nPos := aScan(aClientes, {|x| x[1] == cCNPJBase})
        
        If nPos == 0
            aAdd(aClientes, {cCNPJBase, {aFatura[nX]}})
        Else
            aAdd(aClientes[nPos,2], aFatura[nX])
        EndIf
    Next
    
    // Processa cada cliente
    For nI := 1 to Len(aClientes)
        aFaturaAtual := aClientes[nI,2]
        cCNPJBase := aClientes[nI,1]
        
        // Processa os títulos do cliente
        For nZ := 1 To Len(aFaturaAtual)
            aRecnos := aFaturaAtual[nZ,2]

            For nY := 1 To Len(aRecnos)
                SE1->(DbGoTo(aRecnos[nY]))
                
				// Posiciona na SA1 para cada título individualmente
        		SA1->(DbSetOrder(1))
        		SA1->(DbSeek(xFilial("SA1") + SE1->E1_CLIENTE + SE1->E1_LOJA))
        		_cCgc := Alltrim(SA1->A1_CGC)
				
                // Busca valor da carga
                nVlrCarga := 0
                If ! SE1->E1_PREFIXO $ cSeries
                    DbSelectArea("ZA4")
                    ZA4->(DbSetOrder(4))
                    If ZA4->(DbSeek(SE1->E1_FILIAL + SE1->E1_NUM + SE1->E1_PREFIXO + "inclusao"))
                        nVlrCarga := ZA4->ZA4_VLRCAR
                    EndIf
                EndIf
                
                // Processa igual à função original
                If !PosiTabs()
                    cTexto += Alltrim(SE1->E1_FILIAL) + ";" + ;
                             Alltrim(SE1->E1_NUM) + ";" + ;
                             Alltrim(SE1->E1_PREFIXO) + ";" + ;
                             Transform(_cCgc, "@R 99.999.999/9999-99") + ";" + ; // _cCgc posicionado
                             SE1->E1_TIPO + ";" + ;
                             ";" + ;
                             cNomRem + ";" + ;
                             Alltrim(cNomDest) + ";" + ;
                             DtoC(SE1->E1_EMISSAO) + ";" + ;
                             Transform(SE1->E1_VALOR,"@E 99,999,999.99") + ";" + ;
                             Transform(nVlrCarga,"@E 99,999,999.99") + CRLF
                Else
                    cFrete := IIF(SC5->C5_TPFRETE == 'C','CIF','FOB')
                    If ! SE1->E1_PREFIXO $ cSeries
                        SA1->(dbSetOrder(1))
                        SA1->(DbSeek(xFilial("SA1") + DT6->( DT6->(DT6_CLIREM + DT6_LOJREM))))
                        cNomRem := Transform( SA1->A1_CGC, "@R 99.999.999/9999-99" )+";"+SA1->A1_NOME
                        SA1->(DbSeek(xFilial("SA1") + DT6->( DT6->(DT6_CLIDES + DT6_LOJDES))))
                        cNomDest := Transform( SA1->A1_CGC, "@R 99.999.999/9999-99" )+";"+SA1->A1_NOME
                    Else
                        If PosiZA5()
                            SA1->(dbSetOrder(1))
                            SA1->(DbSeek(xFilial("SA1") + ZA5->( ZA5->(ZA5_REM + ZA5_REMLOJ))))
                            cNomRem := Transform( SA1->A1_CGC, "@R 99.999.999/9999-99" )+";"+SA1->A1_NOME
                            SA1->(DbSeek(xFilial("SA1") + ZA5->( ZA5->(ZA5_DEST + ZA5_DESLOJ))))
                            cNomDest := Transform( SA1->A1_CGC, "@R 99.999.999/9999-99" )+";"+SA1->A1_NOME
                        EndIf
                    EndIf

                    cTexto += Alltrim(SE1->E1_FILIAL) + ";" + ;
                             Alltrim(SE1->E1_NUM) + ";" + ;
                             Alltrim(SE1->E1_PREFIXO) + ";" + ;
                             Transform(_cCgc, "@R 99.999.999/9999-99") + ";" + ; // Alterado aqui
                             SE1->E1_TIPO + ";" + ;
                             cFrete + ";" + ;
                             cNomRem + ";" + ;
                             Alltrim(cNomDest) + ";" + ;
                             DtoC(SE1->E1_EMISSAO) + ";" + ;
                             Transform(SE1->E1_VALOR,"@E 99,999,999.99") + ";" + ;
                             Transform(nVlrCarga,"@E 99,999,999.99") + CRLF
                EndIf
            Next
        Next
        
        // Salva arquivo único por CNPJ base
        cFileCsv := cDirFinal + AllTrim(cCNPJBase) + ".csv"
        MemoWrite(cFileCsv, cTexto)
        
        // Limpa texto para próximo cliente
        cTexto := "FILIAL;DOCUMENTO;SERIE;TOMADOR;TIPO;FRETE;CNPJ REM;NOME REM;CNPJ DEST;NOME DEST;EMISSÃO;VALOR;VALOR CARGA"+CRLF
    Next
    
    SA1->(RestArea(aAreaSA1))
Return

//-------------------------------------------------------------------
/*/{Protheus.doc} ExporXML
Função responsável por gerar xml e/ou dactes em arquivo definido 
pelo usuário.
@author  Samuel Dantas
@since   19/01/2020
@version version
/*/
//-------------------------------------------------------------------
Static Function ExporXML(lXml, lDactes)
    
    Local cQuery := ""
    Local cFile := ""
    Local cFileRet := ""
    Local aAreaSE1 := SE1->(GetArea())
    Local aAreaSA1 := SA1->(GetArea())
    Local cPathRar := ""
    Local cDirFinal := cGetFile("Todos | *.* ", OemToAnsi("Selecione o diretorio"),  ,         ,.F.,GETF_LOCALHARD + GETF_RETDIRECTORY, .F.)  
    Local cCGC := ""
    Local _cFilAux := cFilAnt
    Default lXml := .F.
    Default lDactes := .F.
	If !Empty(cDirFinal)
		SA1->(DbSetOrder(1))
		aToZip  := {}
		cDirXML   := "\anexos\xmls\"
		cDirDacte := "\anexos\dactes\"
		For nZ := 1 to len(aRecnosSE1)
			aFaturas := aRecnosSE1[nZ][2]
			For nI := 1 To len(aFaturas)
				SE1->(DbGoTo(aFaturas[nI]))
				setEmpresa(SE1->E1_FILIAL)
				PosiTabs()
				SA1->(DbSeek(xFilial("SA1") + SE1->(E1_CLIENTE+E1_LOJA)))
				cCGC := Alltrim(SA1->A1_CGC)
				
				If ALLTRIM(SE1->E1_PREFIXO) $ '2,3'
					cQuery := " SELECT ZA4.R_E_C_N_O_ AS RECNO FROM "  + RetSqlName( 'ZA4') + " ZA4"
					cQuery += " WHERE ZA4.D_E_L_E_T_ <> '*' AND ZA4_CHAVE = '"+SF2->F2_CHVNFE+"' AND ZA4_ACAO = 'inclusao' "
				Else
					cQuery := " SELECT ZA5.R_E_C_N_O_ AS RECNO FROM "  + RetSqlName( 'ZA5') + " ZA5"
					cQuery += " WHERE ZA5.D_E_L_E_T_ <> '*' AND ZA5_ACAO = 'inclusao' "
					cQuery += " AND ZA5_NUMNF = '"+SF2->F2_DOC+"' "
					cQuery += " AND ZA5_SERNF = '"+SF2->F2_SERIE+"' "
					cQuery += " AND ZA5_CLIENT = '"+SF2->F2_CLIENTE+"' "
					cQuery += " AND ZA5_LOJA = '"+SF2->F2_LOJA+"' "
					cQuery += " AND ZA5_FILNFS = '"+SF2->F2_FILIAL+"' "
				EndIf

				If Select('QRY') > 0
					QRY->(dbclosearea())
				EndIf
				
				TcQuery cQuery New Alias 'QRY'
				
				While QRY->(!Eof())
					If ALLTRIM(SE1->E1_PREFIXO) $ '2,3'
						ZA4->(DbGoTo(QRY->RECNO))
						If lXml 
							Memowrite(cDirXMl + (cArqXML := Alltrim(ZA4->ZA4_CHAVE) + ".XML"), ZA4->ZA4_BODY)
							aAdd(aToZip,cDirXMl + cArqXML)   
						EndIf
						If lDactes
							u_uRTMSR31(ZA4->ZA4_BODY, cArqCte := Alltrim(ZA4->ZA4_CHAVE) + ".PDF", cDirDacte)                    
							aAdd(aToZip,cDirDacte + cArqCte)
						EndIf
					Else
						ZA5->(DbGoTo(QRY->RECNO))
						If lXml 
							Memowrite(cDirXMl + (cArqXML := Alltrim(ZA5->ZA5_FILNFS)+Alltrim(ZA5->ZA5_NUMNF) + ".XML"), ZA5->ZA5_BODY)
							aAdd(aToZip,cDirXMl + cArqXML)   
						EndIf
					EndIf
					QRY->(dbSkip())
				EndDo        
			Next
			setEmpresa(_cFilAux)
			cPathRar := IIF( lXml .AND. lDactes, "_DACTES_XML.zip", IIF(lDactes, "_DACTES.zip", "_XML.zip"))

			If FZip("\anexos\"+ cCGC + cPathRar, aToZip, "\anexos\") == 0
				cFile       := "\anexos\"+ cCGC + cPathRar
				If CpyS2T( cFile, cDirFinal )  
					// cFile := cDirFinal +  cNumLiq + cPathRar
					ApMsgInfo("Arquivo Salvo com sucesso.")
				EndIf
			Else
				MsgAlert('Erro ao tentar compactar xmls e dactes para envio. ')
			EndIf
		Next
    EndIf

    SE1->(RestArea(aAreaSE1))
    SA1->(RestArea(aAreaSA1))

Return cFile


Static Function GetMsg()
  Local oFont := TFont():New('Courier new',,-18,.T.)
  Local oDlgAux 
  
  Local oMultiGet, oDlgAux
	Local oPanel11, oPanel21, oPanel31
	Default cObserv	:= Space(100) 
	Default cTitulo 	:= "Observação via Email" 
	Default cTitulo2	:= "Observação para ser enviada por email:"
	Default lConfirmar:= .T.

	DEFINE MSDIALOG oDlgAux FROM 1,1 TO 250,450 TITLE cTitulo PIXEL

	@C(001),C(001) MSPANEL oPanel11 PROMPT "" SIZE C(001),C(015) OF oDlgAux
	oPanel11:align := CONTROL_ALIGN_TOP

	@C(001),C(001) MSPANEL oPanel21 PROMPT "" SIZE C(001),C(015) OF oDlgAux
	oPanel21:align := CONTROL_ALIGN_ALLCLIENT

	@C(001),C(001) MSPANEL oPanel31 PROMPT "" SIZE C(001),C(015) OF oDlgAux
	oPanel31:align := CONTROL_ALIGN_BOTTOM 

	@ C(002),C(005) Say cTitulo2 Size C(100),C(008) COLOR CLR_BLACK PIXEL OF oPanel11
	oMultiGet := TMultiget():Create(oPanel21,{|u| if(Pcount() > 0,cMsgEmail := u, cMsgEmail )},1,1,1,1,,nil,nil,nil,nil,.T.,nil,nil, ,nil,nil,,{ || .T.})	
	oMultiGet:align := CONTROL_ALIGN_ALLCLIENT

	If lConfirmar
		TButton():New(005, 005, "&Confirmar", oPanel31,{|| oDlgAux:End() },40,010,,,.F.,.T.,.F.,,.F.,,,.F.)
		TButton():New(005, 055, "&Fechar"	, oPanel31,{|| oDlgAux:End() },40,010,,,.F.,.T.,.F.,,.F.,,,.F.)
	Else
		TButton():New(005, 005, "&Fechar"	, oPanel31,{|| oDlgAux:End() },40,010,,,.F.,.T.,.F.,,.F.,,,.F.)		
	EndIf


	ACTIVATE MSDIALOG oDlgAux CENTERED
  
Return

//-------------------------------------------------------------------
/*/{Protheus.doc} LAUA004Z
Envia liquidação para o nucci
@author  Samuel Dantas
@since   18/02/2020
@version version
/*/
//-------------------------------------------------------------------
User Function LAUA004Z(cNumLiq)

	Local aHeader       := {}
    Local oDados
    Local cJson         := ''
    Local cDestino         := ''
    Local lRet         	:= .F.
	Default cNumLiq := "024786"
	
    If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '010101'
    EndIf
    
	oJson := U_LAUF010E(cNumLiq)
	If ValType(oJson) == "J"
	
		//https://lauto.nuccitms.com.br/custom/rest/lauto/get_motorista_totvs.php
		
		cEndereco           := SuperGetMv('MS_URLNUC', .F.,"https://lauto.nuccitms.com.br/custom/")
		oPost               := JsonObject():new()
		oRest               := fwRest():New(cEndereco)
		oResult             := JsonObject():new()
		
		cJson               := oJson:toJson()
		
		cDestino            := "rest/lauto/fatura_totvs.php"
			
		oRest:setPath(cDestino)    
		oRest:SetPostParams(cJson)

		aAdd(aHeader, "Content-Type: application/json")

		If oRest:Post(aHeader) 
			oResult:FromJson(oRest:GetResult())        
			ConOut(oResult['message'])       
			lRet := .T.
		Else
			lRet := .F.
		EndIf
	EndIf

Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} LAUA004W
Envia liquidação para o nucci
@author  author
@since   date
@version version
/*/
//-------------------------------------------------------------------
User Function LAUA004W(cNumLiq)
	Local cQuery := ""
	Local lRet := .F.
	private nSuccess := 0
	private nErro 	:= 0
	private cMsg := ""
	
	If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '010101'
    EndIf
	lRet := U_LAUA004Z(SE1->E1_NUMLIQ)
	
	If lRet 
		ApMsgInfo("Fatura enviada para NUCCI com sucesso.")
	Else
		ApMsgInfo("Falha ao enviar fatura para o NUCCI.")
	EndIf
	
Return

//-------------------------------------------------------------------
/*/{Protheus.doc} LAUA004Y
Método criado para envio de todas as faturas para o NUCCI
@author  Samuel Dantas
@since   18/02/2020
@version version
/*/
//-------------------------------------------------------------------
User Function LAUA004Y
	Local cQuery := ""
	private nSuccess := 0
	private nErro 	:= 0
	private cMsg := ""
	
	If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '010101'
    EndIf

	cQuery := " SELECT R_E_C_N_O_ AS RECNO FROM "  + RetSqlName('SE1') + " SE1 "
	cQuery += " WHERE SE1.D_E_L_E_T_ <> '*' AND E1_NUMLIQ <> ' ' "
	
	If Select('QRYAUX') > 0
		QRYAUX->(dbclosearea())
	EndIf
	
	TcQuery cQuery New Alias 'QRYAUX'

	SE1->(DbSetOrder(1))
	While QRYAUX->(!Eof())
		SE1->(DbGoTo(QRYAUX->(RECNO)))	
		If U_LAUA004Z(SE1->E1_NUMLIQ)
			nSuccess += 1
		Else
			nErro += 1
		EndIf
		QRYAUX->(dbSkip())
	EndDo
	
Return

//-------------------------------------------------------------------
/*/{Protheus.doc} LAUA004X
Enviar liquidação pro NUCCI de acordo com objeto json passado
@author  Samuel Dantas	
@since   04/02/2020
@version version
/*/
//-------------------------------------------------------------------
User Function LAUA004X(oJson)

	Local aHeader       := {}
    Local oDados
    Local cJson         := ''
    Local cDestino         := ''
    Local lRet         	:= .F.
	Default cNumLiq := "000154"
	
    If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '010101'
    EndIf
    
	If ValType(oJson) == "J"
	
		//https://lauto.nuccitms.com.br/custom/rest/lauto/get_motorista_totvs.php
		
		cEndereco           := SuperGetMv('MS_URLNUC', .F.,"https://lauto.nuccitms.com.br/custom/")
		oPost               := JsonObject():new()
		oRest               := fwRest():New(cEndereco)
		oResult             := JsonObject():new()
		
		cJson               := oJson:toJson()
		
		cDestino            := "rest/lauto/fatura_totvs.php"
			
		oRest:setPath(cDestino)    
		oRest:SetPostParams(cJson)

		aAdd(aHeader, "Content-Type: application/json")

		If oRest:Post(aHeader) 
			oResult:FromJson(oRest:GetResult())        
			ConOut(oResult['message'])       
			lRet := .T.
		Else
			lRet := .F.
		EndIf
	EndIf

Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} POPYNLiq
Método para popular dados no campo E1_YNLIQ
@author  Samuel Dantas
@since   18/02/2020
@version version
/*/
//-------------------------------------------------------------------
Static Function PopYNLiq(cNumLiq)
    
    Local _cNumLiq := ""
    Local cQuery := ""
	Local aAreaSE1 := SE1->(GetArea())
	
	
	Default cNumLiq := ""

    If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01RN0001'
    EndIf
    
    cQuery := " SELECT DISTINCT R_E_C_N_O_ AS RECNOSE1 FROM "  + RetSqlName('SE1') + " SE1"
    cQuery += " WHERE SE1.D_E_L_E_T_ <> '*' AND E1_NUMLIQ = '"+cNumLiq+"' "
    
    If Select('QRYLIQ') > 0
        QRYLIQ->(dbclosearea())
    EndIf
    
    TcQuery cQuery New Alias 'QRYLIQ'

    SE1->(DbSetOrder(1))
    While QRYLIQ->(!Eof())
        SE1->(DbGoTo(QRYLIQ->RECNOSE1))
        _cNumLiq := SE1->E1_NUMLIQ
        _cDocLiq := SE1->E1_NUM
        _cFilLiq := SE1->E1_FILIAL

		RecLock('SE1', .F.)
			SE1->E1_YZA7 := FO0->FO0_YZA7
		SE1->(MsUnLock())
		
        aChaves := GetTitFat()
		If ValType(aChaves) <> "A"
			aChaves := {}
		EndIf

        For nI := 1 To len(aChaves)
            SE1->(DbSeek(aChaves[nI]))
            RecLock('SE1', .F.)
                SE1->E1_YNLIQ := _cNumLiq
				SE1->E1_YDOCFAT := _cDocLiq
                SE1->E1_YFILFAT := _cFilLiq
            SE1->(MsUnLock())
        Next
        QRYLIQ->(dbSkip())
    EndDo

	SE1->(RestArea(aAreaSE1))
Return

//-------------------------------------------------------------------
/*/{Protheus.doc} ValidaLiq
Função que verifica as baixas geradas na liquidação
@author  Samuel Dantas
@since   17/02/2020	
@version version
/*/
//-------------------------------------------------------------------
Static Function VldValor(cNumFat)
	Local cMsg 		:= ""
	Local aAreaSE1 	:= SE1->(GetArea())
	Local aAreaSE5 	:= SE5->(GetArea())
	Local cQryBx 	:= ""
	Local nTotFatura 	:= 0
	Local lRet 	:= .F.
	Default cLiq := ""
	
	cQryBx := " SELECT R_E_C_N_O_ AS RECNOSE1 FROM SE1010 SE1 "
	cQryBx += " WHERE SE1.D_E_L_E_T_ <> '*' AND E1_YFATSEQ = '"+cNumFat+"'  "

	If Select('QRYBX') > 0
		QRYBX->(dbclosearea())
	EndIf

	TcQuery cQryBx New Alias 'QRYBX'

	While QRYBX->(!Eof())
		SE1->(DbGoTo(QRYBX->RECNOSE1))
		If SE1->E1_SALDO != 0
			return .F.
		EndIf
		nTotFatura += ( SE1->E1_VALOR - RetImp() )
		QRYBX->(DbSkip())
	EndDo

	SE1->(DBOrderNickname("FINALIQ"))
	If SE1->(DbSeek(FO0->(FO0_FILIAL+FO0_NUMLIQ)))
		If ROUND(SE1->E1_VALOR,2) == ROUND(nTotFatura, 2)
			lRet := .T.
		EndIf
	EndIf

	SE1->(RestArea(aAreaSE1))
	SE5->(RestArea(aAreaSE5))


Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} Confirmar
Funcao que faz a confirmacao, geracao das faturas
@author  Sidney Sales	
@since   05/12/19
@version 1.0
/*/
//-------------------------------------------------------------------
Static Function FatuInd(oDlg)
	Local i
	Local j
	Local nQtdFat	:= 0
	Local cFilAux := ""
	Private _cFilZA7 := cFilAnt
	Private dDataAtu := dDtEmissa

	If Empty(cNatureza)
		MsgAlert('Por favor, preencha a natureza.')
		Return .T.
	EndIf

	If dVencto < dDataBase
		MsgAlert('O vencimento não pode ser menor que a data atual.')
		Return .T.
	EndIf
	
	If aScan(oBrowseSE1:oData:aArray, {|x| x[1] == cMark}) == 0
		MsgAlert('Selecione pelo menos uma pre fatura para ser gerada.')
		Return .T.
	EndIf

	If !MsgYesNo('Tem certeza que deseja gerar fatura individual para cada CTE e/ou NFSE ?')
		Return .T.
	EndIf

	If MsgYesNo('Confirma a geração das faturas selecionadas?')

		cCodZA7 := GetSxeNum('ZA7', 'ZA7_CODIGO')
		ConfirmSX8()

		//Cria o registro da ZA7, sera o agrupamento das faturas		
		RecLock('ZA7', .T.)
			ZA7->ZA7_FILIAL := xFilial('ZA7')
			ZA7->ZA7_CODIGO	:= cCodZA7
			ZA7->ZA7_DATA	:= dDataBase
			ZA7->ZA7_HORA	:= Time()				
			ZA7->ZA7_STATUS	:= "P"
			ZA7->ZA7_GRUPO	:= cGrupo
			ZA7->ZA7_CODSA1	:= IIF(Empty(cGrupo),cCliente,"")
			ZA7->ZA7_LOJA	:= IIF(Empty(cGrupo),cLojaCli,"")
		ZA7->(MsUnLock())
		cFilAux := cFilAnt
		BEGIN TRANSACTION
		//Percorre a grid das faturas que serao geradas por cliente
		For i := 1 to Len(oBrowseSE1:oData:aArray)
			cFilant := _cFilZA7
			dDataBase := dDataAtu 
			//Verifica se esta marcada
			If oBrowseSE1:oData:aArray[i, 1] == cMark

				//Pega o cliente atual
				cCliLoja := oBrowseSE1:oData:aArray[i, 2]
				
				//Verifica a possicao dele no array de RECNOS de SE1's
				nPosCli  := aScan(aRecnosSE1, {|x| Alltrim(x[1]) == cCliLoja })
				aAuxSE1	 := aRecnosSE1[nPosCli][2]
				
				_cMsgDate := ""
				//Percorre os titulos que foram agrupados para o cliente 
				For j := 1 to Len(aAuxSE1)
					nVlrFat	 := 0
					cFilant := _cFilZA7
					cNumFat := GetNumFat()	
					//Posiciona o titulo e monta o filtro
					SE1->(DbGoTo(aAuxSE1[j]))								
					If Alltrim(SE1->E1_YFATSEQ) == Alltrim(cNumFat)
						loop
					EndIf

					RecLock('SE1', .F.)
						SE1->E1_YFATSEQ := cNumFat
					SE1->(MsUnLock())

					nVlrFat += SE1->E1_SALDO - RetImp()

					If SE1->E1_VENCTO <= Date()
						RecLock('SE1', .F.)
							SE1->E1_VENCTO 	:= SE1->E1_VENCTO 	+ 365
							SE1->E1_VENCREA := SE1->E1_VENCREA 	+ 365
						SE1->(MsUnLock())
					EndIf
					cFiltro := " E1_YFATSEQ = '"+cNumFat+"' "
				
					//Chama a funcao que ira incluir a liquidacao(fatura)
					lRet := liquidar(cFiltro, StrTran(cCliLoja, '/',''), nVlrFat, cCodZA7, _cFilZA7,cNumFat)
					
					//Caso ocorra algum erro, ira faturar apenas o que conseguiu, pergunta se continua ou se interrompe(mas na cancela nda. )
					If ! lRet 
						If MsgYesNo('Foram encontrados erros na geração da fatura para o cliente ' + Alltrim(SA1->A1_NOME) + ' deseja parar o processamento?')
							Exit
						EndIf
					Else
						nQtdFat++
					EndIf

				Next
				If ! lRet 
						Exit
				EndIf
			EndIf
		Next
		END TRANSACTION
		cFilAnt := cFilAux
		//Se tiver faturado algo, se nao, apaga o registro(ZA7)
		If nQtdFat > 0
			ApMsgInfo('Processamento Concluído. Faturas geradas: ' + cValToChar(nQtdFat))
			RecLock('ZA7', .F.)
				ZA7->ZA7_STATUS := 'G' // Gerado
			ZA7->(MsUnLock())
		Else
			MsgAlert('Nenhuma fatura foi gerada.')
			RecLock('ZA7', .F.)
				ZA7->(DbDelete())
			ZA7->(MsUnLock())
		EndIf
	Else
		MsgAlert('Operação cancelada pelo usuário.')
	EndIf

	oDlg:Deactivate()

Return .T.

//-------------------------------------------------------------------
/*/{Protheus.doc} RetImp
Retorna impostos
@author  Samuel dantas
@since   17/03/2020
@version version
/*/
//-------------------------------------------------------------------
Static Function RetImp
	Local cQry := ""
	Local nTotal := 0

	cQry := " SELECT SUM(E1_VALOR) AS TOTIMP FROM "  + RetSqlName('SE1') + " SE1"
	cQry += " WHERE SE1.D_E_L_E_T_ <> '*' AND E1_FILIAL = '"+SE1->E1_FILIAL+"' AND E1_NUM = '"+SE1->E1_NUM+"' AND E1_PREFIXO = '"+SE1->E1_PREFIXO+"' AND E1_CLIENTE = '"+SE1->E1_CLIENTE+"' AND E1_LOJA = '"+SE1->E1_LOJA+"' "
	cQry += " AND E1_TIPO IN ('AB-','FB-','FC-','FU-','IR-','IN-','IS-','PI-','CF-','CS-','FE-','IV-') "
	
	If Select('QRYIMP') > 0
		QRYIMP->(dbclosearea())
	EndIf
	
	TcQuery cQry New Alias 'QRYIMP'
	
	While QRYIMP->(!Eof())
		nTotal := QRYIMP->TOTIMP
		QRYIMP->(dbSkip())
	EndDo
	
Return nTotal


Static Function PesqDoc
    
    If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01RN0001'
    EndIf

    Private aPergs	  := {}
	Private aRet	  := {}
	Private aDados	  := {}
	Private nPosRecno := 14
    Private oDlgBusc
    Private oBrwBus
	Private oPanel0X
	Private dDatade := dDataBase
	Private dDataAte := dDataBase
	Private cCliente    := ""
	Private cLoja       := ""
	Private cGrpVen       := ""
	Private oPanelMod2
    Private aFields := {}
	Private cMark 	 := GetMark()
	Private oWindow 

	aDados := retPesq()
	
    //Cria a tela
    oDlgBusc := FwDialogModal():New()
  	oDlgBusc:SetTitle('Itens da fatura')
    oDlgBusc:SetSize(300,550)
	oDlgBusc:CreateDialog()	
	oDlgBusc:addCloseButton( , "Fechar")
	
	oPanelMod2	:= oDlgBusc:GetPanelMain()
	
	oLayerA := FwLayer():New()
	oLayerA:Init(oPanelMod2)
	
	//Linha dos Filtros
	oLayerA:addLine("lin01",100,.T.)		
	oLayerA:AddCollumn('col01', 100, .T., 'lin01')	
	oLayerA:addWindow("col01","win01",'Títulos',100,.F.,.F.,{|| }, "lin01",{|| })
	oPanel0X := oLayerA:getWinPanel('col01', 'win01', "lin01")

	//Cria array com os campo de marcacao
	oBrwBus 	:= uFwBrowse():create(oPanel0X,,aDados[2],,aDados[1],, ,, .F.)	
	oDlgBusc:AddButton('Selecionar'			, { ||  BuscarDoc() }	, 'Selecionar', 	, .T., .T., .T., )
	// oBrwBus:SetUseFilter()
	oBrwBus:Activate()
	 

   //Ativa a tela principal
	oDlgBusc:Activate()	
		                                 
Return

//-------------------------------------------------------------------
/*/{Protheus.doc} BuscarDoc
Busca documento na frid
@author  Samuel Dantas
@since   18/02/2020
@version version
/*/
//-------------------------------------------------------------------
Static Function BuscarDoc()
	Local aTits		:= oBrowseTIT:oData:aArray 
	Local nAtTit	:= oBrowseTIT:nAt
	Local aTitBus	:= oBrwBus:oData:aArray 
	Local nAtBus	:= oBrwBus:nAt
	
	If len(aTits) > 0 
		nPos := aScan(aTits,{|x| x[12] == aTitBus[nAtBus][12] })
		
		If aTitBus[nAtBus][12] > 0
			SE1->(DbGoTo(aTitBus[nAtBus][12]))
			If !Empty(SE1->E1_BAIXA)
				ApMsgInfo("Título "+cDoc+" está faturado. Fatura: "+SE1->E1_YDOCFAT+".")
				oDlgBusc:Deactivate()
			EndIf
		EndIf
		
		If nPos > 0
			oBrowseTIT:nAt := nPos
			oBrowseTIT:Refresh()
			oDlgBusc:Deactivate()
		Else
			ApMsgInfo("Título "+cDoc+" não encontrado.")
		EndIf
	EndIf
	oDlgBusc:Deactivate()
Return

//-------------------------------------------------------------------
/*/{Protheus.doc} retPesq
Funcao para retornar os dados para a grid
@author  Samuel Dantas
@since   08/02/18
@version 1.0
/*/
//-------------------------------------------------------------------
Static Function retPesq()
	Local cQuery 	:= ""
	Local aHeader	:= {}
	Local aCols		:= {}
	aAux 	:= oBrowseSE1:oData:aArray 
	nAt		:= oBrowseSE1:nAt
	cCliAtual := StrTran(aAux[nAt][2],"/","")
	
	//Cabecalho da grid de titulos
	aAdd(aHeader,{'E1_FILIAL'	,10, 'E1_FILIAL'	})
	aAdd(aHeader,{'E1_PREFIXO'	,10, 'E1_PREFIXO'	})
	aAdd(aHeader,{'E1_NUM'		,10, 'E1_NUM' 		})
	aAdd(aHeader,{'E1_TIPO'		,10, 'E1_TIPO' 		})
    aAdd(aHeader,{'E1_PARCELA'	,10, 'E1_PARCELA' 	})
	aAdd(aHeader,{'E1_EMISSAO'	,10, 'E1_EMISSAO' 	})
	aAdd(aHeader,{'E1_VENCTO'	,10, 'E1_VENCTO' 	})
    aAdd(aHeader,{'E1_CLIENTE'	,10, 'E1_CLIENTE' 	})
	aAdd(aHeader,{'E1_LOJA'	    ,10, 'E1_LOJA' 	})
	aAdd(aHeader,{'E1_SALDO'	,10, 'E1_SALDO' 	})
	aAdd(aHeader,{'E1_VALOR'	,10, 'E1_VALOR' 	})
	aAdd(aHeader,{'IDENTIFICADOR'	,10, '' 	})
    
	If !Empty(cCliAtual)
		cQuery := " SELECT DISTINCT SE1.R_E_C_N_O_ AS RECNOSE1, E1_FILIAL, E1_PREFIXO, E1_NUM, E1_EMISSAO, E1_VENCTO, E1_VENCREA, E1_SALDO, E1_VALOR, E1_CLIENTE, E1_LOJA FROM " + RetSqlTab('SE1')
		If !lOutros
			cQuery += " LEFT JOIN " + RetSqlTab('SC5') + " ON C5_FILIAL = E1_FILIAL AND C5_NUM = E1_PEDIDO AND SC5.D_E_L_E_T_ = SE1.D_E_L_E_T_ "
		EndIf
		cQuery += " INNER JOIN " + RetSqlTab('SA1') + " ON A1_FILIAL = LEFT(E1_FILIAL,2) AND E1_CLIENTE = A1_COD AND E1_LOJA = A1_LOJA AND SA1.D_E_L_E_T_ = SE1.D_E_L_E_T_ "
		cQuery += " WHERE SE1.D_E_L_E_T_ = '' "
		cQuery += " AND E1_EMISSAO BETWEEN '" + DtoS(dEmissDe) + "' AND '" + DtoS(dEmissAte) + "' "
		cQuery += " AND E1_SALDO > 0 "
		cQuery += " AND E1_BAIXA = '' "
		cQuery += " AND E1_NUM = '"+PADL(cDoc,9,"0")+"' "
		cQuery += " AND LEFT(E1_FILIAL,2) = '" + LEFT(xFilial('SE1'), 2) + "' "

		cQuery += " AND E1_CLIENTE+E1_LOJA = '" +cCliAtual+ "' "	
		
		If !Empty(cTpFrete) .AND. !lOutros
			cQuery += " AND C5_TPFRETE = '" + LEFT(cTpFrete, 1) + "' "
		EndIf
		
		cQuery += " AND E1_NUMLIQ = '' AND E1_NUMBOR = '' " //Titulos que nao foram liquidados
		
		cTipos := ''

		If lCtes
			cTipos := "'CTE'"
		EndIf

		If lVenda
			If Empty(cTipos)
				cTipos := "'NF '"
			Else
				cTipos += ",'NF '"
			EndIf
		EndIf

		If lServico
			If Empty(cTipos)
				cTipos := "'NFS'"
			Else
				cTipos += ",'NFS'"
			EndIf
		EndIf
		If  !lOutros
			cQuery += " AND E1_TIPO IN ("+ cTipos+ ")" 		
		Else
			cTpAux := "'AB-','FB-','FC-','FU-','IR-','IN-','IS-','PI-','CF-','CS-','FE-','IV-'"
			If !('NF ' $ cTipos)
				cTpAux += IIF(Empty(cTpAux),"'NF '",",'NF '")
			EndIf
			If !('NFS' $ cTipos)
				cTpAux += IIF(Empty(cTpAux),"'NFS'",",'NFS'")
			EndIf
			If !('CTE' $ cTipos)
				cTpAux += IIF(Empty(cTpAux),"'CTE'",",'CTE'")
			EndIf
			If !Empty(cTpAux)
				cQuery += " AND E1_TIPO NOT IN ("+cTpAux+")" 		
			EndIf
		EndIf	

		cQuery += " ORDER BY E1_FILIAL, E1_PREFIXO, E1_NUM, E1_EMISSAO, E1_VENCTO, E1_VENCREA, E1_SALDO, E1_VALOR, E1_CLIENTE, E1_LOJA "
		
		If Select('QRYSE1') > 0
			QRYSE1->(dbclosearea())
		EndIf
		
		TcQuery cQuery New Alias 'QRYSE1'
		
		
		
		While QRYSE1->(!Eof())
			SE1->(DbGoTo(QRYSE1->RECNOSE1))
			aAdd(aCols, { SE1->E1_FILIAL, SE1->E1_PREFIXO, SE1->E1_NUM, SE1->E1_TIPO,SE1->E1_PARCELA, SE1->E1_EMISSAO, SE1->E1_VENCTO,;
								SE1->E1_CLIENTE, SE1->E1_LOJA,; 
								SE1->E1_SALDO, SE1->E1_VALOR,SE1->(Recno())})
			QRYSE1->(dbSkip())
		EndDo
	EndIf
    //Caso nao tenha dados no select, preencho uma linha vazIa
	If Empty(aCols)
		aAdd(aCols, {'','  ', '', '', CtoD(''), CtoD(''), CtoD(''), 0, 0, "", "",0 })
	EndIf

Return { aHeader, aCols }
//-------------------------------------------------------------------
/*/{Protheus.doc} PosiZA5
Posiciona a tabela ZA5 (integração de nfse)
@author  Samuel Dantas
@since   02/04/2020
@version version
/*/
//-------------------------------------------------------------------
Static Function PosiZA5
	Local cQryZA5 := ""
	Local lRet    := .F.	
	Local cSerie  := ""
	
	cSerie := u_getSerie('', SE1->E1_FILIAL)
	If ValType(cSerie) <> "C"
		cSerie := ""
	EndIf
	
	cQryZA5 := " SELECT ZA5.R_E_C_N_O_ AS RECNO FROM "  + RetSqlName( 'ZA5') + " ZA5"
	cQryZA5 += " WHERE ZA5.D_E_L_E_T_ <> '*' AND ZA5_ACAO = 'inclusao' "
	cQryZA5 += " AND ZA5_NUMNF = '"+SE1->E1_NUM+"' "
	cQryZA5 += " AND ZA5_SERNF = '" + cSerie +"'"
	cQryZA5 += " AND ZA5_CLIENT = '"+SE1->E1_CLIENTE+"' "
	cQryZA5 += " AND ZA5_LOJA = '"+SE1->E1_LOJA+"' "
	cQryZA5 += " AND ZA5_FILNFS = '"+SE1->E1_FILIAL+"' "
	cQryZA5 += " ORDER BY ZA5_DATA,ZA5_HORA DESC "
	
	If Select('QRYZA5') > 0
		QRYZA5->(dbclosearea())
	EndIf
	
	TcQuery cQryZA5 New Alias 'QRYZA5'
	
	If QRYZA5->(!Eof())
		ZA5->(DbGoto(QRYZA5->RECNO))
		lRet    := .T.
		QRYZA5->(dbSkip())
	EndIf

Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} setEmpresa
Seta a empresa de acordo com o CNPJ
@author  Sidney Sales
@since   11/12/19
@version 1.0
/*/
//-------------------------------------------------------------------
Static Function setEmpresa(_cFilial)
    
    Local lRet := .F.
    
    OpenSM0(cEmpAnt)
    
    SM0->(DbGoTop())

    While SM0->(!Eof())
        If Alltrim(SM0->M0_CODFIL) == _cFilial
            cFilAnt := SM0->M0_CODFIL
            lRet    := .T.
            Exit
        EndIf
        SM0->(DbSkip())
    EndDo

Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} PosiTabs
Posiciona tabelas a partir da SE1
@author  Samuel Dantas
@since   20/12/2019
@version version
/*/
//-------------------------------------------------------------------
Static Function PosiTabs()
  Local cQuery := ""
  Local lRet   := .F.
  
  If SE1->E1_PREFIXO $ '2  ,3  '
    cQuery := " SELECT SC5.R_E_C_N_O_ AS RECNOSC5, SF2.R_E_C_N_O_ AS RECNOSF2, DT6.R_E_C_N_O_ AS RECNODT6 FROM "  + RetSqlName('SC5') + " SC5"
  Else
    cQuery := " SELECT SC5.R_E_C_N_O_ AS RECNOSC5, SF2.R_E_C_N_O_ AS RECNOSF2 FROM "  + RetSqlName('SC5') + " SC5"
  EndIf

  cQuery += " INNER JOIN  "+RetSqlName('SF2')+" SF2 ON F2_FILIAL = C5_FILIAL AND F2_DOC = C5_NOTA AND F2_SERIE = C5_SERIE AND F2_CLIENTE = C5_CLIENTE AND F2_LOJA = C5_LOJACLI AND SF2.D_E_L_E_T_ = SC5.D_E_L_E_T_ "
  If SE1->E1_PREFIXO $ '2  ,3  '
    cQuery += " INNER JOIN  "+RetSqlName('DT6')+" DT6 ON F2_FILIAL = DT6_FILIAL AND DT6_CHVCTE = F2_CHVNFE AND SF2.D_E_L_E_T_ = DT6.D_E_L_E_T_ "
  EndIF
  cQuery += " WHERE SC5.D_E_L_E_T_ <> '*'"
  cQuery += " AND C5_FILIAL = '"+SE1->E1_FILIAL+"' "
  cQuery += " AND C5_NUM = '"+SE1->E1_PEDIDO+"' "
  
  If Select('QRY') > 0
    QRY->(dbclosearea())
  EndIf
  
  TcQuery cQuery New Alias 'QRY'
  
  If QRY->(!Eof())
    SF2->(DbGoTo(QRY->RECNOSF2))
    SD2->(DbSetOrder(3)) //D2_FILIAL, D2_DOC, D2_SERIE, D2_CLIENTE, D2_LOJA, D2_COD, D2_ITEM, R_E_C_N_O_, D_E_L_E_T_
    SD2->(DbSeek(SF2->(F2_FILIAL + F2_DOC + F2_SERIE + F2_CLIENTE + F2_LOJA )))
    SC5->(DbGoTo(QRY->RECNOSC5))
    If SE1->E1_PREFIXO $ '2  ,3  '
      DT6->(DbGoTo(QRY->RECNODT6))
    EndIF
    
    // DT6->(DBOrderNickname("DT6CHVCTE"))
    // DT6->(DbSeek(xFilial("DT6") + SF2->F2_CHVNFE))

    lRet := .T.
  EndIf
  
Return lRet

Static Function GetTitFat()
  Local aChaves := {}
  Local cQuery := ""

  cQuery := " SELECT DISTINCT FK7_CHAVE FROM "  + RetSqlName('FO0') + " FO0"
  cQuery += " INNER JOIN "+RetSqlName("FO1")+" FO1 ON FO0_FILIAL = FO1_FILIAL AND FO0_PROCES = FO1_PROCES "
  cQuery += " INNER JOIN "+RetSqlName("FK7")+" FK7 ON LEFT(FK7_FILIAL,2) = LEFT(FO1_FILIAL,2) AND FK7_IDDOC = FO1_IDDOC "
  cQuery += " WHERE FO0.D_E_L_E_T_ <> '*' AND FO0_NUMLIQ = '"+SE1->E1_NUMLIQ+"' AND FK7_ALIAS = 'SE1' AND FO0_FILIAL = '"+SE1->E1_FILIAL+"' AND FO0_STATUS <> '5' "
  
  If Select('QRY') > 0
    QRY->(dbclosearea())
  EndIf
  
  TcQuery cQuery New Alias 'QRY'
  
  While QRY->(!Eof())
    aAdd(aChaves, StrTran(QRY->FK7_CHAVE, "|" , "") )
    QRY->(dbSkip())
  EndDo

Return aChaves
