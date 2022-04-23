#include "TOTVS.CH"
//-------------------------------------------------------------------
/*/{Protheus.doc} 
ExecAuto Fina811()
//FINA810
@author Thiago Prado
@since 17/11/2021
@version 12.1.27
@
/*/
//-------------------------------------------------------------------
function U_xFina811()
	//<<NAO MANIPULAR O FONTE>> APENAS CHUMBE OS VALORES NAS VARIAVEIS
	local a811Param           :={}         as array
	local cCliD               :="000001"   as character //Filtro Cliente DE
	local cLojD               :="01"       as character //Filtro Cliente loja DE
	local cCliAt              :="999999"   as character //Filtro Cliente loja Até
	local cLojAt              :="01"       as character //Filtro Cliente loja Até
	local cLay                :="000001"   as character //Carta de cobrança
	
	Private aRetorno :={} 
	private lMsErroAuto     := .F.         as logical
	private lAutoErrNoFile  := .T.         as logical

	
	//Preparação do Ambiente
	RpcSetEnv("01","01","Admin","","FIN","FINA811",{"SA1"})

	dbSelectArea("FWP")
	FWP->(DbGoTop())
	FWP->(DbSetOrder(1))//FWP_FILIAL, FWP_CODCRT, R_E_C_N_O_, D_E_L_E_T_ 
	if !FWP->(DbSeek(xFilial("FWP") + cLay))  // Valido se encontrou o layout da carta de cobrança
		conOut("Layout nao encontrado")
		RETURN
	endIf
	//Parametros padroes da rotina
	aAdd( a811Param,  cCliD         							)   //Código do cliente De (tipo caracter)
	aAdd( a811Param,  cLojD       								)   //Loja do cliente De (tipo caracter)
	aAdd( a811Param,  cCliAt         							)   //Código do cliente Até (tipo caracter)
	aAdd( a811Param,  cLojAt     								)   //Loja do cliente Até (tipo caracter)
	aAdd( a811Param,  dDataBase - 999                           )   //Data de emissão De (tipo data)
	aAdd( a811Param,  dDataBase + 180                           )   //Data de emissão Até (tipo data)
	aAdd( a811Param,  dDataBase                                 )   //Data de referência (tipo data)
	aAdd( a811Param,  0                                         )   //Valor de (tipo numérico)
	aAdd( a811Param,  999999999.99                              )   //Valor Até (tipo numérico)
	aAdd( a811Param,  '2'                                       )   //"2" (tipo caracter) - IMPORTANTE:  SEMPRE 2 PARA NÃO APRESENTAR TELA DE SELEÇÃO DE FILIAIS
	aAdd( a811Param,  '1'                                       )   //Títulos a vencer (tipo caracter  sendo "1" = Sim ou "2" = Não)
	aAdd( a811Param,  '1'                                       )   //Considera valor (tipo caracter sendo "1" = Total ou  "2" = Saldo)
	aAdd( a811Param,  '2'                                       )   //2" (tipo caracter) - IMPORTANTE: SEMPRE 2 PARA NÃO APRESENTAR TELA DE SELEÇÃO DE SITUAÇÃO DE COBRANÇA
	aAdd( a811Param,  {}                                        )   //{} (tipo array) - IMPORTANTE: ENVIAR ARRAY VAZIO
	aAdd( a811Param,  ''                                        )   //"" (tipo caracter) - IMPORTANTE: SEMPRE VAZIO
	aAdd( a811Param,  cLay                                      )   //Código do layout da carta a ser utilizado (tipo caracter)
    aAdd( a811Param,  '2'                                       )   //A1_MSBLQL


	MSExecAuto({|x,y| Fina811(x,y)},0,a811Param)

	if lMsErroAuto
		ConOut("Erro !")
		MostraErro()
		aLog := GetAutoGRLog()
	else
		//FwAlertSucess("Sucesso!")

		Alert("Sucesso!")
	endif

	RpcClearEnv()

Return


