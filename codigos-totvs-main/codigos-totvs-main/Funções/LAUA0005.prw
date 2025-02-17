#Include 'Protheus.ch'
#Include 'Parmtype.ch'
#include "FWMVCDEF.CH"
#include "TOPCONN.CH"
#INCLUDE 'Rwmake.ch'
#INCLUDE 'TbIconn.ch'

STATIC __lChgX5FIL := .T.

//-------------------------------------------------------------------
/*/{Protheus.doc} LAUA0005
Tela para verificao de integraes
@author  Sidney Sales
@since   12/12/2019
@version 1.0
/*/
//-------------------------------------------------------------------
User Function LAUA0005()
	Local oBrowse
	//Montagem do Browse principal	
	oBrowse := FWMBrowse():New()
	//LEGENDA
	oBrowse:AddLegend("ZA4->ZA4_STATUS=='A' ", "GREEN"   	, 'Aberto')
	oBrowse:AddLegend("ZA4->ZA4_STATUS=='C' ", "PINK"   	, 'Cancelamento')
	oBrowse:AddLegend("ZA4->ZA4_STATUS=='P' .AND. ZA4->ZA4_STAFIS=='P' ", "BLUE"    	, 'Processado')
	oBrowse:AddLegend("ZA4->ZA4_STATUS=='E'", "RED"			, 'Erro')
	oBrowse:AddLegend("ZA4->ZA4_STAFIS!='P'", "YELLOW"		, 'Pendente dados fiscais')

	oBrowse:SetAlias('ZA4')
	oBrowse:SetDescription('Integrao CTE - Protheus x Nucci')
	oBrowse:SetMenuDef('LAUA0005')
	oBrowse:Activate()
Return

//Montagem do menu 
Static Function MenuDef()
	Local aRotina := {}
	Local aOpcoes := {}
		
	aAdd( aRotina, { 'Visualizar'		, 'VIEWDEF.LAUA0005'	, 0, 2, 0, NIL } ) 
	aAdd( aRotina, { 'Incluir' 			, 'VIEWDEF.LAUA0005'	, 0, 3, 0, NIL } )
	aAdd( aRotina, { 'Alterar' 			, 'VIEWDEF.LAUA0005'	, 0, 4, 0, NIL } )
	aAdd( aRotina, { 'Imprimir' 		, 'VIEWDEF.LAUA0005'	, 0, 8, 0, NIL } )
	aAdd( aRotina, { 'Reabrir Erros' 	, 'U_LAUA005C()'		, 0, 8, 0, NIL } )
	
	aAdd( aRotina, { 'Processar CTe', 'U_LAUA005A()', 0, 8, 0, NIL } )

	aOpcoes := {}
	aAdd( aOpcoes, { 'Excluir'		, 'U_LAUF005B("E")', 0, 8, 0, NIL } )
	aAdd( aOpcoes, { 'Reprocessar'	, 'U_LAUF005B("R")', 0, 8, 0, NIL } )

	aAdd( aRotina, { 'Fiscal CTe', aOpcoes, 0, 8, 0, NIL } )
	
	

Return aRotina

//Construcao do mdelo
Static Function ModelDef()
	Local oModel
	Local oStruZA4 := FWFormStruct(1,"ZA4")

	oModel := MPFormModel():New("MD_ZA4") 
	oModel:addFields('MASTERZA4',,oStruZA4)
	oModel:SetPrimaryKey({'ZA4_FILIAL', 'ZA4_CODIGO'})

Return oModel

//Construcao da visualizacao
Static Function ViewDef()
	Local oModel := ModelDef()
	Local oView
	Local oStrZA4:= FWFormStruct(2, 'ZA4')

	oView := FWFormView():New()
	oView:SetModel(oModel)

	oView:AddField('FORM_ZA4' , oStrZA4,'MASTERZA4' ) 
	oView:CreateHorizontalBox( 'BOX_FORM_ZA4', 100)
	oView:SetOwnerView('FORM_ZA4','BOX_FORM_ZA4')

Return oView

//-------------------------------------------------------------------
/*/{Protheus.doc} LAUA005A
description
@author  author
@since   13/12/2019
@version 1.0
/*/
//-------------------------------------------------------------------
User Function LAUA005A()

	Local aPergs 		:= {}   
	Local aRet			:= {}     
	Local lSel, dDataDe, dDataAte

	Private oProcess
	Private nQtdCtes	:= 0
	
	aAdd(aPergs, {2,"Selecionado"	,'Sim'		,{'Sim','No'},50,'.T.',.T.})
	aAdd(aPergs, {1,"Data de"  		,dDataBase	,'@!','.T.','','.T.',50,.F.})
	aAdd(aPergs, {1,"Data Ate" 		,dDataBase	,'@!','.T.','','.T.',50,.F.})

	If ParamBox(aPergs ,"Processamento de CTe",aRet,,,,,,,.F.)      	
		lSel	:= aRet[1] == 'Sim'
		dDataDe	:= aRet[2] 
		dDataAte:= aRet[3] 
		dDataAte:= aRet[3] 
	Else
		MsgAlert("Operao cancelada", "Ateno")
		Return
	EndIf

	//Processar CTes
	If lSel
		MsAguarde( {|| U_LAUF0005(ZA4->(RecNo()))}, "Processando CTe","Aguarde finalizao do processamento." )	
		ApMsgInfo('Fim do processamento.')
	Else

 		oProcess := MsNewProcess():New( { || processa() }, "Processando CTes", "Aguarde...", .F. )	
		
		cQuery := " SELECT R_E_C_N_O_ AS RECNZA4, * FROM " + RetSqlTab('ZA4')
		cQuery += " WHERE D_E_L_E_T_ = '' AND ZA4_STATUS = 'A' "
		cQuery += " AND ZA4_DATA BETWEEN '" + DtoS(dDataDe) + "' AND '" + DtoS(dDataAte) + "' "

		If Select('QRY1') > 0
			QRY1->(DbCloseArea())
		EndIf
		
		TcQuery cQuery New Alias 'QRY1'
		
		Count to nQtdCtes

		If nQtdCtes == 0
			MsgAlert('No existem CTes para processamento no intervalo selecionado.')
			Return
		Else
			QRY1->(DbGoTop())
			oProcess:Activate() 
			ApMsgInfo('Fim do processamento.')
		EndIf

	EndIf
	
Return

User Function LAUA005E
	Local nOpc := 0
	Local cProcCte 
	Local cProcNfs 
	Private oProcess
	Private nQtdCtes	:= 0

	If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01PI0030'
		SetModulo( "SIGAFAT", "FAT" )
		InitPublic()
		SetsDefault()
	EndIf
	
	cProcCte := Alltrim(SuperGetMv("MS_PROCTE", .F.,"S"))
	cProcNfs := Alltrim(SuperGetMv("MS_PRONFS", .F.,"S"))
	
	If cProcCte == 'S'
		cQuery := " SELECT R_E_C_N_O_ AS RECNZA4, * FROM " + RetSqlTab('ZA4')
		cQuery += " WHERE D_E_L_E_T_ = '' AND ZA4_STATUS = 'A' "

		If Select('QRY1') > 0
			QRY1->(DbCloseArea())
		EndIf
		
		TcQuery cQuery New Alias 'QRY1'

		While QRY1->(!Eof())			
			U_LAUF0005(QRY1->RECNZA4)
			QRY1->(DbSkip())
		EndDo
	EndIf
	
	If cProcNfs == 'S'
		cQuery := " SELECT R_E_C_N_O_ AS RECNZA5, * FROM " + RetSqlTab('ZA5')
		cQuery += " WHERE D_E_L_E_T_ = '' AND ZA5_STATUS = 'A' "

		If Select('QRY2') > 0
			QRY2->(DbCloseArea())
		EndIf
		
		TcQuery cQuery New Alias 'QRY2'

		While QRY2->(!Eof())	
			U_LAUF0009(QRY2->RECNZA5, 3)
			QRY2->(DbSkip())
		EndDo
	EndIf

Return

Static Function processa
	Local nQtd := 1
	
	oProcess:SetRegua1(nQtdCtes)	

	While QRY1->(!Eof())
		oProcess:IncRegua1("Processando CTE " + cValToChar(nQtd) + '/' + cValToChar(nQtdCtes))					
		U_LAUF0005(QRY1->RECNZA4)
		nQtd++
		QRY1->(DbSkip())
	EndDo
	
Return

User Function LAUA005C
	Local cQuery
	
	cQuery := " UPDATE " + RetSqlName('ZA4') + " SET ZA4_STATUS = 'A' WHERE ZA4_STATUS = 'E' AND D_E_L_E_T_ = '' "	

	If TCSqlExec(cQuery) < 0	
		MsgAlert('No foi possvel atualizar os erros. Falha no update:  ' + TCSQLError())
	Else
		ApMsgInfo('Atualizado com sucesso.')
	EndIf

Return


//-------------------------------------------------------------------
/*/{Protheus.doc} LAUA005B
Mtodo usado para excluso e reprocessamento de dados fiscais da nota.
@author  Samuel Dantas 
@since   16/12/2019
@version version
/*/
//-------------------------------------------------------------------
User Function LAUA005B(lReprocessa)
    Local aRet := {}
    Local i, j
    Local bBlock := ErrorBlock()
    Private nOrdem := 0
    Private cChave := ""
    Private aTabs := {}
    Private oCTE
    Private cAlias
    Private cCNPJ
    Private cErro := ""
    Private cAviso := ""
    Default lReprocessa := .F.

	cQuery := " SELECT R_E_C_N_O_ AS RECNO FROM "  + RetSqlName('SF2') + " SF2"
	cQuery += " WHERE SF2.D_E_L_E_T_ <> '*' AND F2_CHVNFE = '"+ZA4->ZA4_CHAVE+"' "
	
	If Select('QRY') > 0
		QRY->(dbclosearea())
	EndIf
	
	TcQuery cQuery New Alias 'QRY'
	
	If QRY->(!Eof())
		SF2->(DbGoTo(QRY->RECNO))
		QRY->(dbSkip())
	EndIf
	If SF2->(!EoF())
		//Posiona o registro para processamento
		ZA4->(DbSeek(xFilial("ZA4") + SF2->F2_CHVNFE))
		cXml := ZA4->ZA4_BODY

		//Transforma o CTE em um objeto
		oCTE := xmlParser(cXml, "_", @cErro, @cAviso)    
		cCNPJ := oCTE:_CTEPROC:_CTE:_INFCTE:_EMIT:_CNPJ:TEXT
		
		If !setEmpresa(cCNPJ)
			return 'Cadastro de filial no localizado com o CNPJ ' + cCNPJ
		EndIf

		SA1->(DbSetOrder(1))
		If SA1->(DbSeek(xFilial("SA1") + SF2->F2_CLIENTE + SF2->F2_LOJA))
		EndIf

		BEGIN TRANSACTION
			aTabs := {}
				
			aAdd(aTabs, {'SF2', 1, SF2->(F2_FILIAL + F2_DOC + F2_SERIE + F2_CLIENTE + F2_LOJA), "SF2->(F2_FILIAL + F2_DOC + F2_SERIE + F2_CLIENTE + F2_LOJA)" })
			aAdd(aTabs, {'SD2', 3, SF2->(F2_FILIAL + F2_DOC + F2_SERIE + F2_CLIENTE + F2_LOJA), "SD2->(D2_FILIAL + D2_DOC + D2_SERIE + D2_CLIENTE + D2_LOJA)" })
			aAdd(aTabs, {'SF3', 4, SF2->(F2_FILIAL + F2_CLIENTE + F2_LOJA + F2_DOC + F2_SERIE), "SF3->(F3_FILIAL + F3_CLIEFOR + F3_LOJA + F3_NFISCAL + F3_SERIE)" }) 
			aAdd(aTabs, {'SFT', 1, SF2->F2_FILIAL + 'S' + SF2->(F2_SERIE + F2_DOC + F2_CLIENTE + F2_LOJA), "SFT->(FT_FILIAL + FT_TIPOMOV + FT_SERIE + FT_NFISCAL + FT_CLIEFOR + FT_LOJA)" })

			//Percorre o array de tabelas, preenchendo os campos extras
			For i := 1  to Len(aTabs)
				cAlias  := aTabs[i][1]
				nOrdem  := aTabs[i][2]
				cChave  := aTabs[i][3]
				
				//Retorna os dados que devem ser preenchidos na tabela
				aRet    := getArray(oCTE, cAlias, SF2->(Recno())) //getArray(oCTE, cAlias, nRecNoSF2)
				(cAlias)->(DbSetOrder(nOrdem))
				(cAlias)->(DbSeek(cChave))            
				
				//Perccorre os registro, na pratica, sempre sera um, porem, ja coloquei para evitar. 
				While (cAlias)->(!Eof()) .AND. cChave == &(aTabs[i][4])
					//Altera os campos
					RecLock(cAlias, .F.)
						For j := 1 to Len(aRet)
							cCampo := aRet[j][1]
							xValor := aRet[j][2]
							(cAlias)->&(cCampo) := Iif(!lReprocessa, GetVazio(cAlias,cCampo), xValor ) 
						Next
					(cAlias)->(MsUnLock())
					(cAlias)->(DbSkip())
				EndDo
			Next

			If lReprocessa
				//Cria o DT6, tabela do TMS, usada apenas para funcionar o registro fiscal
				aDT6 := getArray(oCTE, 'DT6', SF2->(Recno()))
				RecLock('DT6', .T.)            
					For j := 1 to Len(aDT6)
						cCampo := aDT6[j][1]
						xValor := aDT6[j][2]
						DT6->&(cCampo) := xValor
					Next
				DT6->(MsUnLock())
			Else
				//Exclui DT6
				DT6->(DBOrderNickname("DT6CHVCTE"))
				DT6->(DbSeek(xFilial("DT6") + SF2->F2_CHVNFE))
				RecLock('DT6', .F.)
					DT6->(DbDelete())
				DT6->(MsUnLock())
			EndIf

		END TRANSACTION 
	EndIf
    
Return


//-------------------------------------------------------------------
/*/{Protheus.doc} GetVazio
Retorna valor vazio de acordo com tipo
@author  Samuel Dantas
@since   16/12/2019
@version version
/*/
//-------------------------------------------------------------------
Static Function GetVazio(cTab,cCampo)
    Local xRet := (cTab)->&(cCampo)

    If ValType(xRet) == 'N'
        xRet := 0
    ElseIf ValType(xRet) == 'L'
        xRet := .F.
    ElseIf ValType(xRet) == 'C'
        xRet := ""
    ElseIf ValType(xRet) == 'M'
        xRet := ""
    ElseIf ValType(xRet) == 'D'
        xRet := CtoD("")
    EndIf

Return xRet
