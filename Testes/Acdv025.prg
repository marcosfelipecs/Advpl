#INCLUDE "Acdv025.ch" 
#INCLUDE 'PROTHEUS.CH'
#INCLUDE 'APVT100.CH'

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o    ³ ACDV025    ³ Autor ³ Anderson Rodrigues  ³ Data ³ 15/03/04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³Apontamento Producao PCP Mod2 - Este programa tem por       ³±±
±±³          ³objetivo realizar os apontamentos de Producao/Perda e Hrs   ³±±
±±³          ³improdutivas baseados no roteiro de operacoes               ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³SIGAACD                                                     ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
*/              
Function ACDV025()
Local bkey05
Local bkey09
Local cOP      := Space(Len(CBH->CBH_OP))
Local cOperacao:= Space(Len(CBH->CBH_OPERAC))
Local cTransac := Space(Len(CBH->CBH_TRANSA))
Local cRetPe   := ""    
Local lContinua:= .T.
Local lVolta   := .f.
Local lPiope   := ExistBlock("CB023IOPE") 

Private cOperador  := Space(Len(CB1->CB1_CODOPE))
Private cTM        := GetMV("MV_TMPAD")
Private cRoteiro   := Space(Len(SC2->C2_ROTEIRO))
Private cProduto   := Space(Len(SC2->C2_PRODUTO))
Private cLocPad    := Space(Len(SC2->C2_LOCAL))
Private cUltOper   := Space(Len(CBH->CBH_OPERAC))
Private cPriOper   := Space(Len(CBH->CBH_OPERAC))
Private cTipIni    := "1"
Private cUltApont  := " "
Private cApontAnt  := " "
Private nSldOPer   := 0
Private nQtdOP     := 0
Private aOperadores:= {}
Private lConjunto  := .f.
Private lFimIni    := .f.
Private lAutAskUlt := .f.
Private lVldOper   := .f.
Private lRastro    := GetMV("MV_RASTRO")  == "S" // Verifica se utiliza controle de Lote
Private lSGQTDOP   := GetMV("MV_SGQTDOP") == "1" // Sugere quantidade no inicio e no apontamento da producao
Private lInfQeIni  := GetMV("MV_INFQEIN") == "1" // Verifica se deve informar a quantidade no inicio da Operacao
Private lCBAtuemp  := GetMV("MV_CBATUD4") == "1" // Verifica se ajusta o empenho no inicio da producao
Private lVldQtdOP  := GetMV("MV_CBVQEOP") == "1" // Valida no inicio da operacao a quantidade informada com o saldo a produzir da mesma
Private lVldQtdIni := GetMV("MV_CBVLAPI") == "1" // Valida a quantidade do apontamento com a quantidade informada no inicio da Producao
Private lCfUltOper := GetMV("MV_VLDOPER") == "S" // Verifica se tem controle de operacoes
Private lOperador  := GetMV("MV_SOLOPEA",,"2") == "1" // Solicita o codigo do operador no apontamento 1-sim 2-nao (default)
Private lMod1      := .f.
Private lMsHelpAuto:= .f.
Private lMSErroAuto:= .f.
Private lPerdInf   := .F.
IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf
 
// -- Verifica se data do Protheus esta diferente da data do sistema.
DLDataAtu()

If IsTelnet()
	cOperador := CBRETOPE()
	If lPiope
		cRetPe := ExecBlock("CB023IOPE",.F.,.F.,{cOperador})
		If ValType(cRetPe)=="C"
			cOperador := cRetPe
			If ! CBVldOpe(cOperador)
				lContinua := .f.
			EndIf
		EndIf
	EndIf
	If lContinua .And. Empty(cOperador)
		CBALERT(STR0001,STR0002,.T.,3000,2)  //"Operador nao cadastrado"###"Aviso"
		lContinua := .f.
	EndIf
	If lContinua .And. VtModelo() == "RF"
		bkey05   := VTSetKey(05,{|| CB025Encer()},STR0017)       //"Encerrar"
		bkey09   := VTSetKey(09,{|| CB023Hist(cOP)},STR0003)       //"Informacoes"
	EndIf
EndIf

If lContinua .And. Empty(cTM)
	CBALERT(STR0004,STR0002,.T.,3000,2) //"Informe o tipo de movimentacao padrao - MV_TMPAD"###"Aviso"
	lContinua := .f.
EndIf

If lContinua .And. !lRastro .and. lCBAtuemp
	CBALERT(STR0005,STR0002,.T.,4000,2) //"O parametro MV_CBATUD4 so deve ser ativado quando o sistema controlar rastreabilidade"###"Aviso"
	lContinua := .f.
EndIf

If lContinua .And. (lVldQtdOP .or. lVldQtdIni .or. lCBAtuemp) .and. !lInfQeIni
	CBALERT(STR0006,STR0002,.T.,3000,2) //"O parametro MV_INFQEIN deve ser ativado"###"Aviso"
	lContinua := .f.
EndIf

While lContinua
	if lVT100B
		While .t.
			vtClear()
			@ 0, 00 vtSay STR0007 //Producao PCP MOD2
			If lOperador
				cOperador  := Space(Len(CB1->CB1_CODOPE))
				@ 1,00 VtSay STR0013 VtGet cOperador Valid CBVldOpe(cOperador) //"Operador:"
			EndIf
			@ 2,00 VTSAY STR0008 //"OP: "
			@ 2,04 VTGET cOP pict '@!'  Valid CB023OP(cOP) F3 "SC2" When Empty(cOP)
			vtRead
			if VTLastkey() != 27
				lVolta := .f.
				VTClear(1,0,3,19)
				@ 1,00 VTSAY STR0009 //"Operacao: "
				@ 1,10 VTGET cOperacao pict '@!' Valid CB023OPERAC(cOP,cOperacao);
					when iif(VTRow() == 1 .and. VTLastkey() == 5, (VTKeyboard(chr(27)),lVolta := .t.),.t.)
				@ 2,00 VTSAY STR0010 //"Transacao:"
				@ 2,11 VTGET cTransac pict '@!'  Valid CB023VTran(cOP,cOperacao,cOperador,cTransac) F3 "CBI"
				VtRead
			endif
			
			if lVolta
				loop
			endif
			exit				
		enddo	
	elseIf IsTelnet() .and. VtModelo() == "RF"
		VtClear()
		@ 0,00 VTSAY STR0007 //"Producao PCP MOD2"
		If lOperador
			cOperador  := Space(Len(CB1->CB1_CODOPE))
			@ 1,00 VtSay STR0013 VtGet cOperador Valid CBVldOpe(cOperador) //"Operador:"
		EndIf
		@ 2,00 VTSAY STR0008 //"OP: "
		@ 2,04 VTGET cOP pict '@!'  Valid CB023OP(cOP) F3 "SC2" When Empty(cOP)
		@ 4,00 VTSAY STR0009 //"Operacao: "
		@ 4,10 VTGET cOperacao pict '@!' Valid CB023OPERAC(cOP,cOperacao)
		@ 7,00 VTSAY STR0010 //"Transacao:"
		@ 7,11 VTGET cTransac pict '@!'  Valid CB023VTran(cOP,cOperacao,cOperador,cTransac) F3 "CBI"
		VtRead
		If VtLastKey() == 27
			Exit
		EndIf
	Else
		TerIsQuit()
		If TerProtocolo() == "GRADUAL"
			cOperador:= CB023IOPE() // Solcita o operador para Microterminal com porta paralela
			If TerEsc()
				Loop
			EndIf
		EndIf
		TerCls()
		If VtModelo() == "MT44"
			@ 0,00 TerSay STR0007 //"Producao PCP MOD2"
			@ 1,00 TerSay STR0008 //"OP: "
			@ 1,05 TerGetRead cOP pict "XXXXXXXXXXXXX"  Valid CB023OP(cOP)
			If TerEsc()
				If IsTelnet()
					Exit
				EndIf
				Loop
			EndIf
			@ 0,20 TerSay STR0009 //"Operacao: "
			@ 0,32 TerGetRead cOperacao pict 'XX' Valid CB023OPERAC(cOP,cOperacao)
			TerCls()
			If TerEsc()
				Loop
			EndIf
			@ 0,00 TerSay STR0010 //"Transacao:"
			@ 0,12 TerGetRead cTransac pict 'XX'  Valid CB023VTran(cOP,cOperacao,cOperador,cTransac)
		Else
			@ 0,00 TerSay STR0007 //"Producao PCP MOD2"
			@ 1,00 TerSay STR0008 //"OP: "
			@ 1,05 TerGetRead cOP pict 'XXXXXXXXXXXXX' Valid CB023OP(cOP)
			TerCls()
			If TerEsc()
				If IsTelnet()
					Exit // quando for executa pelo sigaacdt a rotina devera' retornar ao menu
				EndIf
				Loop
			Endif
			@ 0,00 TerSay STR0009 //"Operacao: "
			@ 0,12 TerGetRead cOperacao pict 'XX' Valid CB023OPERAC(cOP,cOperacao)
			TerCls()
			If TerEsc()
				Loop
			EndIf
			@ 0,00 TerSay STR0010 //"Transacao:"
			@ 0,12 TerGetRead cTransac pict 'XX'  Valid CB023VTran(cOP,cOperacao,cOperador,cTransac)
		EndIf
		If TerEsc()
			Loop
		EndIf
	EndIf
	if !lVolta
		lContinua := .f.
	endif
	cOP       := Space(Len(CBH->CBH_OP))
	cOperacao := Space(Len(CBH->CBH_OPERAC))
	cTransac  := Space(Len(CBH->CBH_TRANSA))
EndDo
If lContinua
	If IsTelnet() .and. VtModelo() == "RF"
		vtsetkey(05,bkey05)		
		vtsetkey(09,bkey09)
	Else
		TerIsQuit()
	EndIf
EndIf
Return

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o    ³  CB025GRV  ³ Autor ³ Anderson Rodrigues  ³ Data ³ 20/08/02 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³ Realiza gravacao dos arquivos para apontar a Producao      ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ SIGAACD                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
Function CB025GRV(cOP,cOperacao,cTransac,cProduto,cRecurso,cOperador,cTipAtu,nQtd,cLote,dValid,dDtIni,cHrIni,dDtFim,cHrFim,aCpsUsu,cParTot,dDtApont)
Local cCalend,cTempo2
Local cH6PT		:= ""
Local nTempoPar,nTempoTra
Local nPos,nMinutos,nTempo1,nTempo2
Local nSldSH6  := CB023SH6(cOP,cProduto,cOperacao)
Local aDadosSH6:= {}
Local aMata681 := {}
Local aPEAux	:= {} 
Local lACD25MOV := .T.
Local lP025MOV	:= ExistBlock("ACD25MOV")
Local lP025AUT 	:= ExistBlock("CB025AUT")
Local lP025GR 	:= ExistBlock("ACD025GR")
Local cMsgErro	:= "" 
Local nPosRotOr	:= 0 
Default cHrIni  := ""
Default cHrFim  := Left(Time(),5)
Default dDtIni  := CTOD("  /  /    ")
Default dDtFim  := dDataBase
Default dDtApont := dDataBase

If !(Empty(cParTot))
	cH6PT := cParTot
EndIf

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³Ponto de Entrada para verificar se executa mata681.³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
If lP025MOV
	lACD25MOV := ExecBlock("ACD25MOV",.F.,.F.)   
	If ValType(lACD25MOV) <> "L"
	 	lACD25MOV := .T.
	EndIf
Endif

If TerProtocolo() # "PROTHEUS"
	aDadosSH6:= CB023Dados(cOP,cProduto,cOperacao,cOperador) // --> Retorna array contendo as informacoes do ultimo apontamento no SH6
	If !Empty(aDadosSH6)
		dDtIni := aDadosSH6[1,4]
		cHrIni := aDadosSH6[1,5]
	Endif
	cTipo := "1" // -> Inicio da operacao para a OP
	CBH->(DbSetOrder(3))
	If ! CBH->(DbSeek(xFilial("CBH")+cOP+cTipo+cOperacao+cOperador))
		If Empty(DTOS(dDtIni)+cHrIni)
			CBALERT(STR0011,STR0002,.T.,3000,2,Nil)  //"OP inconsistente"###"Aviso"
			DisarmTransaction()
			Break
		Endif
	ElseIf (DTOS(CBH->CBH_DTINI)+CBH->CBH_HRINI) > (DTOS(dDtIni)+cHrIni)
		dDtIni:= CBH->CBH_DTINI
		cHrIni:= CBH->CBH_HRINI
	Endif
	If !lACD25MOV
		CB025CBH(cOP,cTipAtu,cOperacao,cTransac,cOperador,@dDtIni,@cHrIni)
	EndIf
	If dDtIni == dDtFim .and. cHrIni == cHrFim
		cHrFim:= Left(cHrFim,3)+StrZero(Val(Right(cHrFim,2))+1,2)
		If Right(cHrFim,2) == "60"
			cHrFim:= StrZero(Val(Left(cHrFim,2))+1,2)+":00"
			If Left(cHrFim,2)== "24"
				cHrFim:= "00:00"
				dDtFim++
			EndIf
		EndIf
	Endif
Endif
cCalend := GetMV("MV_CBCALEN") // Parametro onde e informado o calendario padrao que deve ser utilizado
If Empty(cCalend)
	cCalend := Posicione("SH1",1,xFilial("SH1")+cRecurso,"H1_CALEND")
Endif
nTempoPar := CB023Pausa(cOP,cOperacao,cRecurso,cOperador,dDtIni,cHrIni,dDataBase,cHrFim)
nTempoTra := IF(SuperGetMV("MV_USACALE",.F.,.T.),PmsHrsItvl(dDtIni,cHrIni,dDtFim,cHrFim,cCalend,"",cRecurso,.T.),A680Tempo(dDtIni,cHrIni,dDtFim,cHrFim))
nTempo1   := nTempoTra - nTempoPar
nTempo2   := Int(nTempo1)
nMinutos  := (nTempo1-nTempo2)*60
If nMinutos == 60
	nTempo2++
	nMinutos:= 0
Endif
cTempo2:= StrZero(nTempo2,3)+":"+StrZero(nMinutos,2)
If TerProtocolo() # "PROTHEUS"
	If IsTelnet() .and. VtModelo() == "RF"
		VtClear()
		VtSay(2,0,STR0012) //"Aguarde..."
	Else
		TerCls()
		TerSay(1,0,STR0012) //"Aguarde..."
	Endif
Endif

aAdd(aMata681,{"H6_OP", cOP              ,NIL})
aAdd(aMata681,{"H6_PRODUTO", cProduto    ,NIL})
aAdd(aMata681,{"H6_OPERAC" , cOperacao   ,NIL})
aAdd(aMata681,{"H6_RECURSO", cRecurso    ,NIL})
aAdd(aMata681,{"H6_DATAINI", dDtIni      ,NIL})
aAdd(aMata681,{"H6_HORAINI", cHrIni      ,NIL})
aAdd(aMata681,{"H6_DATAFIN", dDtFim      ,NIL})
aAdd(aMata681,{"H6_HORAFIN", cHrFim      ,NIL})
If SuperGetMV("MV_CBCALPR", .F., .T.) == .T.
	aAdd(aMata681,{"H6_TEMPO"  , cTempo2 ,NIL})
EndIf
aAdd(aMata681,{"H6_OPERADO", cOperador   ,NIL})
aAdd(aMata681,{"H6_DTAPONT", dDtApont   ,NIL})
If cTipAtu == "4"
	aAdd(aMata681,{"H6_QTDPROD", nQtd    ,NIL})
Elseif cTipAtu == "5"
	aAdd(aMata681,{"H6_QTDPERD" ,nQtd    ,NIL})
Endif
If !Empty(cH6PT)
	aAdd(aMata681,{"H6_PT"  , cH6PT      ,NIL})
Endif
aAdd(aMata681,{"H6_CBFLAG","1"           ,NIL}) // Flag que indica que foi gerado pelo ACD
If !lCfUltOper
	aAdd(aMata681,{"AUTASKULT",lAutAskUlt,NIL})
Endif

aadd(aMata681,{"H6_LOCAL",cLocPad    ,NIL})

If Rastro(SC2->C2_PRODUTO)
	aadd(aMata681,{"H6_LOTECTL",cLote    ,Nil})
	aadd(aMata681,{"H6_DTVALID",dValid   ,Nil})
EndIf
//-- Ponto de entrada que permite manipular o conteudo do array que sera passado para rotina automatica
//-- por isso deve ser usado com muito cuidado para nao descaracterizar
If lP025AUT
	aPEAux := aClone(aMata681)  
	aPEAux := ExecBlock("CB025AUT",.F.,.F.,{aPEAux,cOP,cOperacao,cTransac,cProduto,cRecurso,cOperador,cTipAtu,nQtd,dDtIni,cHrIni,dDtFim,cHrFim,cLote,dValid})
	If ValType(aPEAux)=="A" 
		aMata681 := aClone(aPEAux)
	EndIf
EndIf

If lACD25MOV
	lMsHelpAuto := .T.
	lMSErroAuto := .F.
	nModuloOld  := nModulo
	nModulo     := 4
	msExecAuto({|x|MATA681(x)},aMata681)
	nModulo     := nModuloOld
	
	lMsHelpAuto:=.F.
	If lMSErroAuto
		DisarmTransaction()
	
		// APT PENDENTE 
		nPosRotOr := aScan(aMata681,{|x| x[1] == "PENDENTE"}) //Verifica Tag no Apontamento
		IF nPosRotOr > 0 .and. aMata681[nPosRotOr][2] = "2"
			If IsTelNet()
				VTAlert(STR0019, STR0020,.t.,4000,4)
			Else
				HELP(' ',1,"ACDA080" ,,STR0020,2,0,,,,,, {STR0019})	
			EndIf
			IF FindFunction("ErrosApt")
			 	cMsgErro	:= ErrosApt()
			 EndIF
			IF FindFunction("a250GrvPnd")
				a250GrvPnd(aMata681,"MATA681", cMsgErro)			
			EndIf
		EndIf
	
		Break
	EndIf

EndIF

CB023CBH(cOP,cOperacao,cOperador,cTransac,Nil,dDtIni,cHrIni,dDtFim,cHrFim,cTipAtu,"ACDV023",0,nQtd,cRecurso,aCpsUsu,SH6->H6_LOTECTL,SH6->H6_NUMLOTE,SH6->H6_DTVALID,SH6->H6_DTAPONT)
CB023FIM(cOP,cProduto,cOperacao,cOperador,nQtd,dDtFim,cHrFim)
CB023HrImp(cOP,cOperacao,cRecurso,cOperador,dDtIni,cHrIni,dDtFim,cHrFim)

If lP025GR  // Executado apos a gravacao do apontamento da producao
	ExecBlock("ACD025GR",.F.,.F.,{cOp,cOperacao,cRecurso,cOperador,nQtd,cTransac})
EndIf

Return .t.
/*
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o	 ³CB025Encer    ³ Autor ³ Aecio Ferreira Gomes³ Data ³ 11/10/09 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Responsavel pelo Encerramento das Ops.						³±±
±±³          ³ 						                                        ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ 										                        ³±±
±±³          ³                                                              ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Uso		 ³ ACDV025                                                      ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
Static Function CB025Encer()

Local cOP      := Space(Len(SH6->H6_OP))
Local aMata681 := {}

While .T.
	
	If IsTelnet() .and. VtModelo() == "RF"
		VTCLEAR()
		@ 0,0 vtSay STR0014 //"Encerramento da OP"
		@ 1,0 VTSAY STR0008 //"OP: "
		@ 2,0 VtGet cOP pict '@!'  Valid CB023OP(cOP)  F3 "SC2" When Empty(cOP)
		VTREAD
		If vtLastKey() == 27
			Exit
		EndIf
	EndIf
	DbSelectArea("SH6")
	DbSetOrder(1)
	
	DBSetFilter( {|| cOP == SH6->H6_OP .AND. SH6->H6_PRODUTO == SC2->C2_PRODUTO }, " cOP == SH6->H6_OP .AND. SH6->H6_PRODUTO == SC2->C2_PRODUTO" ) //Verifica se o registro existe na tabela SH6
	DbGoTop()
	If !EOF() .And.  cOP == SH6->H6_OP .AND. SH6->H6_PRODUTO == SC2->C2_PRODUTO  // Se existir o registro nao encerra a OP.
		If ! VTYesNo(STR0015,STR0002,.T.)  //"Deseja encerrar a OP?"###"Aviso"
			cOP := Space(Len(SH6->H6_OP))
			VTGEtSetFocus('cOP')
			Loop
		EndIf	
		aadd(aMata681,{"H6_OP"      , SH6->H6_OP       ,NIL})
		aadd(aMata681,{"H6_PRODUTO" , SH6->H6_PRODUTO  ,NIL})
		aadd(aMata681,{"H6_SEQ"     , SH6->H6_SEQ      ,NIL})

		lMsHelpAuto := .T.
		lMSErroAuto := .F.
		nModuloOld  := nModulo
		nModulo     := 4
		
		MsExecAuto({|x,Y| MATA681(aMata681,7)})// "Encerra ordem de producao"

		nModulo     := nModuloOld
		lMsHelpAuto :=.F.

    Else
		VTAlert(STR0016,STR0002,.T.,3000)// "Nao existem apontamentos para a ordem de producao no arquivo de movimentos da producao","Aviso"
	EndIf   
	DBClearFilter()

	cOP := Space(Len(SH6->H6_OP))
	VTGEtSetFocus('cOP')
End

If lMSErroAuto         
	VTDispFile(NomeAutoLog(),.t.)
Endif

Return !lMSErroAuto         

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³CB025CBH   ³ Autor ³Isaias Florencio      ³ Data ³18/09/14  ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³          ³ Retorna a data e hora final do ultimo registro de log      ³±±
±±³Descri‡…o ³ na tabela CBH, em caso de o ponto de entrada ACD25MOV      ³±±
±±³          ³ estar ativo e retornar .F.                                 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ CB025GRV                                                   ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
Static Function CB025CBH(cOP,cTipAtu,cOperacao,cTransac,cOperador,dDtIni,cHrIni)
Local aAreaAnt 	:= GetArea()
Local aAreaCBH 	:= CBH->(GetArea())
Local cQuery		:= ""
Local cAliasCBH	:= GetNextAlias()

cQuery := "SELECT CBH.CBH_DTFIM AS DTFIM, CBH.CBH_HRFIM AS HRFIM "
cQuery += "FROM "+RetSqlName("CBH")+" CBH "
cQuery += "WHERE CBH.CBH_FILIAL = '"+xFilial('CBH')+"' AND CBH.CBH_OP = '"+ cOP +"' AND "
cQuery += "CBH.CBH_OPERAC = '"+ cOperacao +"' AND CBH.CBH_OPERAD = '"+ cOperador +"' AND "
cQuery += "CBH.CBH_TRANSA = '"+ cTransac +"' AND CBH.CBH_TIPO   = '"+ cTipAtu +"' AND CBH.D_E_L_E_T_ = ' ' "
cQuery += "ORDER BY CBH.R_E_C_N_O_ DESC "

cQuery := ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasCBH,.T.,.T.)

(cAliasCBH)->(DbGoTop())
If !(cAliasCBH)->(Eof())
	dDtIni := STOD((cAliasCBH)->DTFIM)
	cHrIni := (cAliasCBH)->HRFIM
EndIf
(cAliasCBH)->(dbCloseArea())

RestArea(aAreaCBH)
RestArea(aAreaAnt)
Return Nil
