#Include 'Protheus.ch'
#Include 'Parmtype.ch'
#include "FWMVCDEF.CH"
#include "TOPCONN.CH"

//-------------------------------------------------------------------
/*/{Protheus.doc} LAUA0006
Tela para verificao de integrações de NFSE
@author  Sidney Sales
@since   12/12/2019
@version 1.0
/*/
//-------------------------------------------------------------------
User Function LAUA0006()
	Local oBrowse
	//Montagem do Browse principal	
	oBrowse := FWMBrowse():New()
	//LEGENDA
	oBrowse:AddLegend("ZA5->ZA5_STATUS=='A'", "GREEN"   	, 'Aberto')
	oBrowse:AddLegend("ZA5->ZA5_STATUS=='P'", "BLUE"    	, 'Processado')
	oBrowse:AddLegend("ZA5->ZA5_STATUS=='E'", "RED"			, 'Erro')

	oBrowse:SetAlias('ZA5')
	oBrowse:SetDescription('Integração NFSE - Protheus x Nucci')
	oBrowse:SetMenuDef('LAUA0006')
	oBrowse:Activate()
Return

//Montagem do menu 
Static Function MenuDef()
	Local aRotina := {}
	Local aOpcoes := {}
		
	aAdd( aRotina, { 'Visualizar'		, 'VIEWDEF.LAUA0006'	, 0, 2, 0, NIL } )
	aAdd( aRotina, { 'Incluir' 			, 'VIEWDEF.LAUA0006'	, 0, 3, 0, NIL } )
	aAdd( aRotina, { 'Alterar' 			, 'VIEWDEF.LAUA0006'	, 0, 4, 0, NIL } )
	aAdd( aRotina, { 'Imprimir' 		, 'VIEWDEF.LAUA0006'	, 0, 8, 0, NIL } )
	aAdd( aRotina, { 'Reabrir Erros' 	, 'U_LAUA006C()'		, 0, 8, 0, NIL } )
	aAdd( aRotina, { 'Processar NFSe'	, 'U_LAUA006A(3)'		, 0, 8, 0, NIL } )
	aAdd( aRotina, { 'Excluir NFSe'	 	, 'U_LAUA006A(5)'		, 0, 8, 0, NIL } )
	aAdd( aRotina, { 'Forçar Processamento', 'U_LAUA006F()'		, 0, 8, 0, NIL } )
	aAdd( aRotina, { 'Arredonda ISS' 	, 'U_ControlISS()'		, 0, 8, 0, NIL } )

Return aRotina

//Construcao do mdelo
Static Function ModelDef()
	Local oModel
	Local oStruZA5 := FWFormStruct(1,"ZA5")

	oModel := MPFormModel():New("MD_ZA5") 
	oModel:addFields('MASTERZA5',,oStruZA5)
	oModel:SetPrimaryKey({'ZA5_FILIAL', 'ZA5_CODIGO'})

Return oModel

//Construcao da visualizacao
Static Function ViewDef()
	Local oModel := ModelDef()
	Local oView
	Local oStrZA5:= FWFormStruct(2, 'ZA5')

	oView := FWFormView():New()
	oView:SetModel(oModel)

	oView:AddField('FORM_ZA5' , oStrZA5,'MASTERZA5' ) 
	oView:CreateHorizontalBox( 'BOX_FORM_ZA5', 100)
	oView:SetOwnerView('FORM_ZA5','BOX_FORM_ZA5')

Return oView

//-------------------------------------------------------------------
/*/{Protheus.doc} LAUA006A
description
@author  author
@since   13/12/2019
@version 1.0
/*/
//-------------------------------------------------------------------
User Function LAUA006A(nOpc)

	Local aPergs 		:= {}   
	Local aRet			:= {}     
	Local lSel, dDataDe, dDataAte

	Private oProcess
	Private nQtdNfses	:= 0
	Private cPerg := PadR('U_LAUA0006', 10)
	
	aAdd(aPergs, {2,"Selecionado"	,'Sim'		,{'Sim','Não'},50,'.T.',.T.})

	If ParamBox(aPergs ,"Processamento de NFSe",aRet,,,,,,,.F.)      	
		lSel	:= aRet[1] == 'Sim'
		
	Else
		MsgAlert("Operação cancelada", "Atenção")
		Return
	EndIf

	//Processar CTes
	If lSel
		MsAguarde( {|| U_LAUF0009(ZA5->(RecNo()), nOpc)  }, "Processando NFSe","Aguarde finalização do processamento." )	
		ApMsgInfo('Fim do processamento.')
	Else
		ValidPerg()

		If ! Pergunte(cPerg, .T.)
			Return
		EndIf

		dDataDe	:= MV_PAR01
		dDataAte:= MV_PAR02 

 		oProcess := MsNewProcess():New( { || processa(nOpc) }, "Processando NFSe", "Aguarde...", .F. )	
		
		cQuery := " SELECT R_E_C_N_O_ AS RECNZA5, * FROM " + RetSqlTab('ZA5')
		cQuery += " WHERE D_E_L_E_T_ = '' "
		cQuery += " AND ZA5_DATA BETWEEN '" + DtoS(dDataDe) + "' AND '" + DtoS(dDataAte) + "' "
		cQuery += " AND ZA5_FILNFS BETWEEN '" + MV_PAR03 + "' AND '" + MV_PAR04 + "' "
		cQuery += " AND ZA5_CLIENT BETWEEN '" + MV_PAR05 + "' AND '" + MV_PAR06 + "' "
		cQuery += " AND ZA5_LOJA BETWEEN '" + MV_PAR07 + "' AND '" + MV_PAR08 + "' "
		cQuery += " AND ZA5_NUMNF BETWEEN '" + MV_PAR09 + "' AND '" + MV_PAR10 + "' "
		cQuery += " AND ZA5_STATUS = '" + IIF(MV_PAR11 == 1,'E','A') + "' "

		If Select('QRY1') > 0
			QRY1->(DbCloseArea())
		EndIf
		
		TcQuery cQuery New Alias 'QRY1'
		
		Count to nQtdNfses

		If nQtdNfses == 0
			MsgAlert('Não existem NFSes para processamento no intervalo selecionado.')
			Return
		Else
			QRY1->(DbGoTop())
			oProcess:Activate() 
			ApMsgInfo('Fim do processamento.')
		EndIf

	EndIf
	
Return

//Nova função para processamento forçado
User Function LAUA006F()
    // Variáveis para controle do processamento
    Local aPergs := {}   
    Local aRet := {}     
    Local lSel, dDataDe, dDataAte
    Private oProcess
    Private nQtdNfses := 0
    Private cPerg := PadR('U_LAUA0006', 10)
    
    // Adiciona pergunta se deseja processar registro selecionado
    aAdd(aPergs, {2,"Selecionado", 'Sim', {'Sim','Não'}, 50, '.T.', .T.})

    // Exibe tela de parâmetros
    If ParamBox(aPergs, "Processamento Forçado de NFSe", aRet,,,,,,,, .F.)      	
        lSel := aRet[1] == 'Sim'
    Else
        MsgAlert("Operação cancelada", "Atenção")
        Return
    EndIf

    // Processa registro selecionado
    If lSel
        MsAguarde({|| U_LAUF0009(ZA5->(RecNo()), 3, .T.)}, "Processando NFSe", "Aguarde finalização do processamento.")	
        ApMsgInfo('Fim do processamento.')
    Else
        // Processa por intervalo de datas
        ValidPerg()
        If !Pergunte(cPerg, .T.)
            Return
        EndIf

        dDataDe := MV_PAR01
        dDataAte := MV_PAR02 

        oProcess := MsNewProcess():New({|| processaForcado()}, "Processando NFSe", "Aguarde...", .F.)	
        
        // Query para buscar NFSes no intervalo
        cQuery := " SELECT R_E_C_N_O_ AS RECNZA5, * FROM " + RetSqlTab('ZA5')
        cQuery += " WHERE D_E_L_E_T_ = '' "
        cQuery += " AND ZA5_DATA BETWEEN '" + DtoS(dDataDe) + "' AND '" + DtoS(dDataAte) + "' "
        cQuery += " AND ZA5_FILNFS BETWEEN '" + MV_PAR03 + "' AND '" + MV_PAR04 + "' "
        cQuery += " AND ZA5_CLIENT BETWEEN '" + MV_PAR05 + "' AND '" + MV_PAR06 + "' "
        cQuery += " AND ZA5_LOJA BETWEEN '" + MV_PAR07 + "' AND '" + MV_PAR08 + "' "
        cQuery += " AND ZA5_NUMNF BETWEEN '" + MV_PAR09 + "' AND '" + MV_PAR10 + "' "
        cQuery += " AND ZA5_STATUS = '" + IIF(MV_PAR11 == 1,'E','A') + "' "

        If Select('QRY1') > 0
            QRY1->(DbCloseArea())
        EndIf
        
        TcQuery cQuery New Alias 'QRY1'
        Count to nQtdNfses

        If nQtdNfses == 0
            MsgAlert('Não existem NFSes para processamento no intervalo selecionado.')
            Return
        Else
            QRY1->(DbGoTop())
            oProcess:Activate() 
            ApMsgInfo('Fim do processamento.')
        EndIf
    EndIf
Return

// Função auxiliar para processar múltiplas NFSes
Static Function processaForcado()
    Local nQtd := 1
    
    oProcess:SetRegua1(nQtdNfses)	

    While QRY1->(!Eof())
        oProcess:IncRegua1("Processando NFSe " + cValToChar(nQtd) + '/' + cValToChar(nQtdNfses))					
        U_LAUF0009(QRY1->(RECNZA5), 3, .T.)
        nQtd++
        QRY1->(DbSkip())
    EndDo
Return

Static Function processa(nOpc)
	Local nQtd := 1
	
	oProcess:SetRegua1(nQtdNfses)	

	While QRY1->(!Eof())
		oProcess:IncRegua1("Processando NFSe " + cValToChar(nQtd) + '/' + cValToChar(nQtdNfses))					
		U_LAUF0009(QRY1->(RECNZA5), nOpc)
		nQtd++
		QRY1->(DbSkip())
	EndDo
	
Return

User Function LAUA006C
	Local cQuery
	
	cQuery := " UPDATE " + RetSqlName('ZA5') + " SET ZA5_STATUS = 'A' WHERE ZA5_STATUS = 'E' AND D_E_L_E_T_ = '' "	

	If TCSqlExec(cQuery) < 0	
		MsgAlert('Não foi possível atualizar os erros. Falha no update:  ' + TCSQLError())
	Else
		ApMsgInfo('Aatualizado com sucesso.')
	EndIf

Return

//-------------------------------------------------------------------
/*/{Protheus.doc} ValidPerg
Criacao das perguntas do relatorio
@author  author
@since   04/03/2020
@version 1.0
/*/
//-------------------------------------------------------------------
Static Function ValidPerg
	Local aRegs    := {}
	Local aAreaSX1 := SX1->(GetArea())
	Local i, j

	SX1->(DbSetOrder(1))

	// Numeracao dos campos:
	// 01 -> X1_GRUPO   02 -> X1_ORDEM    03 -> X1_PERGUNT  04 -> X1_PERSPA  05 -> X1_PERENG
	// 06 -> X1_VARIAVL 07 -> X1_TIPO     08 -> X1_TAMANHO  09 -> X1_DECIMAL 10 -> X1_PRESEL
	// 11 -> X1_GSC     12 -> X1_VALID    13 -> X1_VAR01    14 -> X1_DEF01   15 -> X1_DEFSPA1
	// 16 -> X1_DEFENG1 17 -> X1_CNT01    18 -> X1_VAR02    19 -> X1_DEF02   20 -> X1_DEFSPA2
	// 21 -> X1_DEFENG2 22 -> X1_CNT02    23 -> X1_VAR03    24 -> X1_DEF03   25 -> X1_DEFSPA3
	// 26 -> X1_DEFENG3 27 -> X1_CNT03    28 -> X1_VAR04    29 -> X1_DEF04   30 -> X1_DEFSPA4
	// 31 -> X1_DEFENG4 32 -> X1_CNT04    33 -> X1_VAR05    34 -> X1_DEF05   35 -> X1_DEFSPA5
	// 36 -> X1_DEFENG5 37 -> X1_CNT05    38 -> X1_F3       39 -> X1_GRPSXG

	aAdd(aRegs, {cPerg, "01", "Data de?"  		, "", "", "mv_ch1", 'D', 8, 0, 0, 'G', "", "MV_PAR01", "",  "", "", "", "", "",    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""})
	aAdd(aRegs, {cPerg, "02", "Data até?" 		, "", "", "mv_ch2", 'D', 8, 0, 0, 'G', "", "MV_PAR02", "",  "", "", "", "", "",    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""})
	aAdd(aRegs, {cPerg, "03", "Filial de?" 		, "", "", "mv_ch3", 'C', LEN(cFilAnt), 0, 0, 'G', "", "MV_PAR03", "",  "", "", "", "", "",    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "SM0", ""})
	aAdd(aRegs, {cPerg, "04", "Filial até?" 	, "", "", "mv_ch4", 'C', LEN(cFilAnt), 0, 0, 'G', "", "MV_PAR04", "",  "", "", "", "", "",    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "SM0", ""})
	aAdd(aRegs, {cPerg, "05", "Cliente De?" 	, "", "", "mv_ch5", 'C', LEN(SA1->A1_COD), 0, 0, 'G', "", "MV_PAR05", "",  "", "", "", "", "",    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "SA1", ""})
	aAdd(aRegs, {cPerg, "06", "Cliente até?" 	, "", "", "mv_ch6", 'C', LEN(SA1->A1_COD), 0, 0, 'G', "", "MV_PAR06", "",  "", "", "", "", "",    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "SA1", ""})
	aAdd(aRegs, {cPerg, "07", "Loja de?" 		, "", "", "mv_ch7", 'C', 4, 0, 0, 'G', "", "MV_PAR07", "",  "", "", "", "", "",    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""})
	aAdd(aRegs, {cPerg, "08", "Loja até?" 		, "", "", "mv_ch8", 'C', 4, 0, 0, 'G', "", "MV_PAR08", "",  "", "", "", "", "",    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""})
	aAdd(aRegs, {cPerg, "09", "Num de?" 		, "", "", "mv_ch9", 'C', 9, 0, 0, 'G', "", "MV_PAR09", "",  "", "", "", "", "",    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""})
	aAdd(aRegs, {cPerg, "10", "Num até?" 		, "", "", "mv_ch10", 'C', 9, 0, 0, 'G', "", "MV_PAR10", "",  "", "", "", "", "",    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""})
	aAdd(aRegs, {cPerg, "11", "Status?" 		, "", "", "mv_ch11", 'C', 1, 0, 0, 'G', "", "MV_PAR11", "E-ERRO",  "", "", "", "", "A-ABERTO",    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""})
	
	For i := 1 To Len(aRegs)
		If ! SX1->(DbSeek(cPerg+aRegs[i,2]))
			RecLock("SX1", .T.)

			For j :=1 to SX1->(FCount())
				If j <= Len(aRegs[i])
					SX1->(FieldPut(j,aRegs[i,j]))
				EndIf
			Next

			SX1->(MsUnlock())
		EndIf
	Next

	SX1->(RestArea(aAreaSX1))
Return
