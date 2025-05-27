#INCLUDE "acdv165.ch" 
#INCLUDE 'PROTHEUS.CH'
#INCLUDE 'APVT100.CH'


// variaveis usadas para POSICAO de campos nos arrays aItens e aItensCB9
STATIC _POSRECNO   :=1
STATIC _POSITEM    :=2
STATIC _POSCODPRO  :=3
STATIC _POSARMAZEM :=4
STATIC _POSENDERECO:=5
STATIC _POSSLDSEP  :=6
STATIC _POSQTDSEP  :=7
STATIC _POSLOTECTL :=8
STATIC _POSNUMLOTE :=9
STATIC _POSNUMSERIE:=10
STATIC _POSSUBVOL  :=11
STATIC _POSIDETI   :=12
STATIC _POSCFLOTE  :=13
STATIC _POSLOTESUG :=14
STATIC _POSSLOTESUG:=15
STATIC _lPulaItem
STATIC __nSem:=0

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡ao    ³ ACDV165  ³ Autor ³ ACD                   ³ Data ³ 03/01/02 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡ao ³ Pre-Separacao                                              ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ OBS      ³ CB7_STATUS:  0-Nao Iniciada 								  ³±±
±±³			 ³ 				1-Em Separacao								  ³±±
±±³			 ³ 				2-Separacao Finalizada 						  ³±±
±±³			 ³ 				3-Em processo embalagem 					  ³±±
±±³			 ³ 				4-Embalagem Finalizada						  ³±±
±±³			 ³ 				5-Gera Nota 								  ³±±
±±³			 ³ 				6-Imprime nota 								  ³±±
±±³			 ³ 				7-Imprime Volume 							  ³±±
±±³			 ³ 				8-Em processo embarque						  ³±±
±±³			 ³ 				9-Finalizado                                  ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
Function ACDV165()
Local 	nTamOrdSep	:= TamSX3("CB8_ORDSEP")[1]
Private cCodOpe  	:=CBRetOpe()
Private cOrdSep  	:= Space(nTamOrdSep)
Private aItens   	:= {}
Private lCB001   	:= UsaCB0('01')
Private lCB002  	:= UsaCB0('02')
Private cPictQtde	:= PesqPict("CB8","CB8_QTDORI")
Private aItensPallet:= {}
Private lIsPallet   := .T.
Private lForcaQtd   := GetMV("MV_CBFCQTD",,"2") =="1"
 
__nSem := 0 // variavel static do fonte para controle de semaforo

If Empty(cCodOpe)
	VTAlert(STR0001,STR0002,.T.,4000)  //"Operador nao cadastrado"###"Aviso"
	Return .F.
EndIf

While .T.
	VTClear()
	@ 0,0 VtSay STR0003 //"Pre-Separacao"
	cOrdSep := Space(nTamOrdSep)
	@ 1,0 VTSay STR0004 //'Informe o codigo:'
	@ 2,0 VTGet cOrdSep PICT "@!" F3 "CB7"  Valid VldCodSep()
	VTRead     
	If VTLastKey() == 27
		Exit
	EndIf
  	If ! MSCBFSem() //fecha o semaforo, somente um separador por ordem de separacao
  		VtAlert(STR0083,STR0002,.T.,4000,3) //"Ordem de Separacao ja esta em andamento...!"###"Aviso"
  		VtKeyboard(Chr(20))  // zera o get
	   	Return .F.
  	EndIf
	Separa()
	Exit
End
Return

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Funcao    ³ vldcodsep³ Autor ³ ACD                   ³ Data ³ 14/08/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descricao ³ valida o codigo da ordem de pre-separacao                  ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
Static Function VldCodSep()
Local nRecCB7 := 0

If Empty(cOrdSep)
	VtKeyBoard(chr(23))
	return .F.
EndIf

CB7->(DbSetOrder(1))
CB7->(DbSeek(xFilial("CB7")+cOrdSep))
If CB7->(Eof())
	VtAlert(STR0005,STR0002,.t.,4000,3)  //"Ordem de Pre-separacao nao encontrada."###"Aviso"
	VtKeyboard(Chr(20))  // zera o get
	Return .F.
EndIf
If ! "09*" $ CB7->CB7_TIPEXP
	VtAlert(STR0006,STR0007,.t.,4000,3)  //"Ordem de Pre-Separacao Invalida"###"Codigo Invalido"
	VtKeyboard(Chr(20))  // zera o get
	Return .F.
EndIf
If CB7->CB7_STATUS == "9"
	nRecCB7 := CB7->(Recno())
	CB7->(dbGoTop())
	While CB7->(!Eof())
		If CB7->CB7_PRESEP == cOrdSep
			VtAlert(STR0008+CB7->CB7_ORDSEP,STR0002,.T.,4000,3) //"Pre-Separacao possui Separacao gerada! Num. "###"Aviso"
			VtKeyboard(Chr(20)) //-- zera o get
			Return .F.
		EndIf
		CB7->(dbSkip())
	EndDo
	CB7->(dbGoTo(nRecCB7))
Else
	If CB7->CB7_STATUS # "0" .AND. CB7->CB7_STATPA == "1" .AND. CB7->CB7_CODOPE # cCodOpe  // SE ESTIVER EM SEPARACAO E PAUSADO SE DEVE VERIFICAR SE O OPERADOR E' O MESMO
		VtBeep(3)
		If !VTYesNo(STR0009+CB7->CB7_CODOPE+STR0010,STR0002,.T.) //"Ordem Separacao iniciada pelo operador "###". Deseja continuar ?"###"Aviso"
			VtKeyboard(Chr(20))  // zera o get
			Return .F.
		EndIf
	ElseIf CB7->CB7_STATUS # "0" .AND. CB7->CB7_STATPA == " " .and. CB7->CB7_CODOPE # cCodOpe  //Ordem Separacao ja esta em andamento...
		VtBeep(3)
		VtAlert(STR0011,STR0002,.t.,4000) //"Ordem de Separacao ja esta em andamento por outro operador!"###"Aviso"
		VtKeyboard(Chr(20))  // zera o get
		Return .F.
	EndIf
EndIf
RecLock("CB7")
If	CB7->CB7_STATUS == "0" .or. Empty(CB7->CB7_STATUS)
	CB7->CB7_STATUS := "1"  // em separacao
	CB7->CB7_DTINIS := dDataBase
	CB7->CB7_HRINIS := StrTran(Time(),":","")
EndIf
If !Empty(CB7->CB7_STATPA)  // se estiver pausado tira o STATUS  de pausa
	CB7->CB7_STATPA := " "
EndIf
CB7->CB7_CODOPE := cCodOpe
CB7->(MsUnlock())
Return .T.

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡Æo    ³ CarregaCB8³Autor ³ ACD                   ³ Data ³ 14/08/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡Æo ³ carrega os itens do cb8 para um array                      ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CarregaCB8(aVetor)
Local nPos      := 0
Local aItenstmp := {}
DbSelectArea("CB8")
CB8->(DbSetOrder(3))
CB8->(DbSeek(xFilial("CB8")+cOrdSep))
While ! CB8->(Eof()) .and. CB8->CB8_FILIAL == xFilial("CB8") .and. CB8->CB8_ORDSEP == cOrdSep
	If Empty(CB8->CB8_SALDOS) .or. ! Empty(CB8->CB8_OCOSEP)
		CB8->(DbSkip())
		Loop
	EndIf
	nPos := Ascan(aItenstmp,{|x| x[3]+x[4]+x[5]+x[8]+x[9]+x[10]+x[13] == CB8->(CB8_PROD+CB8_LOCAL+CB8->CB8_LCALIZ+CB8_LOTECT+CB8_NUMLOT+CB8_NUMSER+CB8_CFLOTE)})
	If  nPos== 0
		aadd(aItenstmp,{{CB8->(Recno())},;
		CB8->CB8_ITEM,;
		CB8->CB8_PROD,;
		CB8->CB8_LOCAL,;
		CB8->CB8_LCALIZ,;
		CB8->CB8_SALDOS,;
		CB8->CB8_SLDPRE,;
		CB8->CB8_LOTECT,;
		CB8->CB8_NUMLOT,;
		CB8->CB8_NUMSER,;
		NIL,;
		NIL,;
		CB8->CB8_CFLOTE })
	Else
		If Ascan(aItenstmp[nPos,1],{|x| x==CB8->(Recno())})==0
			aadd(aItenstmp[nPos,1],CB8->(Recno()))
		EndIf
		aItenstmp[nPos,6]+=CB8->CB8_SALDOS
		aItenstmp[nPos,7]+=CB8->CB8_SLDPRE
	EndIF
	CB8->(DbSkip())
EndDo
aItenstmp := aSort(aItenstmp,,,{|x,y| x[4]+x[5]+x[3]+x[8]+x[9]+x[10] < y[4]+y[5]+y[3]+y[8]+y[9]+y[10]})
If len(aItenstmp) == 0
	Return .F.
EndIf
aVetor := aclone(aItenstmp[1])
Return .T.


/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Funcao    ³ Separa    ³Autor ³ ACD                   ³ Data ³ 14/08/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descricao ³ carrega os itens do cb8 para um array                      ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
Static function Separa()
Local cCodEti
Local lMV_CONFEND  := GETMV("MV_CONFEND") =="1"  // =1 Solicita a leitura do endereco para conferencia
Local lMV_CFENDIG  := GetMV('MV_CFENDIG') =="1"  // =1 sempre solicita a leitura do endereco
Local lMV_OSEP2UN  := GetMV("MV_OSEP2UN") =="1"  // =1 Mostra para separar sempre pela 2 unidade de medida
Local bkey09       := VTSetKey(09,{|| Informa()},STR0012) //"Informacoes"
Local bkey24       := VTSetKey(24,{|| Estorna()},STR0013) //"Estorno"
Local lGeraTermino :=.T.
Local nX   		   := 0  
Local nTamLoc      := Tamsx3("B1_LOCPAD")[1]
Local nTamEnd      := TamSX3("BF_LOCALIZ")[1]

Private cArmazem   := Space(nTamLoc)
Private cEndereco  := Space(nTamEnd)
Private cDivItemPv := GetMV("MV_DIVERPV")  // codigo da divergencia
Private aLoteNew   := {}
Private aSLoteNew  := {}
Private aQtdLida   := {}

_lPulaItem := .f.

While .T.
	If ! CarregaCB8(aItens)
		If VTYesNo(STR0014,STR0015,.T.)  //'Pre-Separacao finalizada, confirma a saida?'###'Atencao'        
			If ExistBlock("ACD165FM")
				ExecBlock("ACD165FM")
			EndIf
			Exit
		Else
			Estorna()
			Loop
		EndIF
	EndIf
	aQtdLida := {}
	If Empty(aItens[_POSSLDSEP])  // se for branco significa que ja foi separado
		Loop
	EndIf
	If ! aItens[_POSARMAZEM]+aItens[_POSENDERECO] == cArmazem+cEndereco
		// Efetua a leitura da etiqueta de endereco
		If ! Endereco(lMV_CONFEND,lMV_CFENDIG,@cArmazem,@cEndereco)
			If QtdComp(aItens[_POSQTDSEP])<>QtdComp(aItens[_POSSLDSEP])
				If VTYesNo(STR0016,STR0015,.T.)  //'Confirma a saida sem geracao parciais das ordens?'###'Atencao'
					lGeraTermino:= .f.
					exit
				EndIf
				_lPulaItem:= .t.
				If ! PulaItem()
					loop
				EndIf
			EndIf
			If VTYesNo(STR0017,STR0015,.T.)  //'Confirma a saida?'###'Atencao'
				exit
			EndIf
			Loop
		EndIf
	EndIf
	//mostra na tela o produto a ser separada
	Tela(lMV_OSEP2UN)

	//solicita a leitura da etiqueta referente a produto
	If ! EtiqProduto()
		If QtdComp(aItens[_POSQTDSEP])<>QtdComp(aItens[_POSSLDSEP])
			If VTYesNo(STR0016,STR0015,.T.)  //'Confirma a saida sem geracao parciais das ordens?'###'Atencao'
				lGeraTermino:= .f.
				exit
			EndIf
			_lPulaItem:= .t.
			If ! PulaItem()
				loop
			EndIf
		EndIf
		If VTYesNo(STR0017,STR0015,.T.)  //'Confirma a saida?'###'Atencao'
			Exit
		EndIf
		Loop
	EndIf 
	 
	Reclock('CB7')
	CB7->CB7_STATUS := "1"  // inicio separacao
	CB7->(MsUnLock()) 
	 
	aArea:=GetArea()
	
	// Verifica se pula o primeiro item
	If Len(aItensPallet) = 0
		If PulaItem()
			Loop
		EndIf
	EndIf
	
	  // Gravacao por itens 
    For nX:= 1 to Len(aItensPallet)

    	DbSelectArea("CB0")
   		DbSeek(xFilial("CB0")+aItensPallet[nX])

		// verifica se pula o item
		If PulaItem()
			Loop
		EndIf
	
		If ! CBProdUnit(aItens[_POSCODPRO]) .or. lCB001
			cCodEti := CB0->CB0_CODETI
		Else
			cCodEti := NIL
		Endif
	
		If ! Grava(cCodEti,nX)
			Loop
		EndIf
		
		If ! Empty(aItens[_POSSLDSEP])
			Loop
		EndIf
		If lMV_CFENDIG
			cArmazem   := Space(nTamLoc)
			cEndereco  := Space(nTamEnd)
		EndIf
	Next nX	
	RestArea(aArea)
			
EndDo
If lGeraTermino
	AnalisaTermino()
Else
	Reclock('CB7')
	CB7->CB7_STATPA := "1"  // Pausa
	CB7->(MsUnLock())
EndIf
VTSetKey(09,bkey09)
VTSetKey(24,bkey24)
MSCBASem()// valor necessario para liberar o semaforo
Return

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Funcao    ³ Endereco ³ Autor ³ ACD                   ³ Data ³ 14/08/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descricao ³ Inicio do Envio para o endereco                            ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
Static Function Endereco(lMV_CONFEND,lMV_CFENDIG,cArmazem,cEndereco)
Local aTela    := VtSave()
Local cEtiqEnd := Space(20)
IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf

VTClear()
cArmazem   := Space(Tamsx3("B1_LOCPAD")[1])
cEndereco  := Space(TamSX3("BF_LOCALIZ")[1])
cEtiqEnd   := Space(20)
If lVT100B // GetMv("MV_RF4X20")
	@ 0,0 VTSay STR0018 //'Va para o endereco'
	@ 1,0 VTSay aItens[_POSARMAZEM]+'-'+aItens[_POSENDERECO]
	If !Localiza(aItens[_POSCODPRO]) .Or. ! lMV_CONFEND
		@ 3,0 VTPause STR0019 //'Enter para continuar'
	Else
		@ 2,0 VTSay STR0020 //'Leia o endereco'
		If lCB002
			@ 3,0 VTGet cEtiqEnd pict '@!' valid VldEnd(@cArmazem,@cEndereco,cEtiqEnd)
		Else
			@ 3,0 VTGet cArmazem pict '@!' valid ! Empty(cArmazem)
			@ 3,3 VTSay "-" VTGet cEndereco pict '@!' valid VtLastKey()==5 .or. VldEnd(@cArmazem,@cEndereco)
		EndIf
		VTRead
		If VTLastkey() == 27
			Return .f.
		EndIf
	EndIf
Else
	@ 1,0 VTSay STR0018 //'Va para o endereco'
	@ 2,0 VTSay aItens[_POSARMAZEM]+'-'+aItens[_POSENDERECO]
	If !Localiza(aItens[_POSCODPRO]) .Or. ! lMV_CONFEND
		@ 6,0 VTPause STR0019 //'Enter para continuar'
	Else
		@ 4,0 VTSay STR0020 //'Leia o endereco'
		If lCB002
			@ 5,0 VTGet cEtiqEnd pict '@!' valid VldEnd(@cArmazem,@cEndereco,cEtiqEnd)
		Else
			@ 5,0 VTGet cArmazem pict '@!' valid ! Empty(cArmazem)
			@ 5,3 VTSay "-" VTGet cEndereco pict '@!' valid VtLastKey()==5 .or. VldEnd(@cArmazem,@cEndereco)
		EndIf
		VTRead
		If VTLastkey() == 27
			Return .f.
		EndIf
	EndIf
Endif
cArmazem := aItens[_POSARMAZEM]
cEndereco:= aItens[_POSENDERECO]
Return .t.

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Funcao    ³ VldEnd   ³ Autor ³ ACD                   ³ Data ³ 14/08/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descricao ³ Validacao do Endereco                                      ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
Static Function VldEnd(cArmazem,cEndereco,cEtiqEnd)
Local aRet
VtClearBuffer()
If Empty(cEtiqEnd) .and. lCB002
	VtAlert(STR0021,STR0002,.t.,4000,3)  //"Etiqueta invalida"###"Aviso"
	VtKeyboard(Chr(20))  // zera o get
	Return .F.
EndIf
If Empty(cArmazem) // se cArmazem == Empty entao e' etiqueta de endereco com CB0
	aRet := CBRetEti(cEtiqEnd,'02')
	If len(aRet) == 0
		VtAlert(STR0021,STR0002,.t.,4000,3)  //"Etiqueta invalida"###"Aviso"
		VtKeyboard(Chr(20))  // zera o get
		Return .F.
	EndIf
	cArmazem  := aRet[2]   // somente esta sendo atribuido para posteriormnete ser comparado com aItens
	cEndereco := aRet[1]
EndIf
If !(cArmazem+cEndereco == aItens[_POSARMAZEM]+aItens[_POSENDERECO] )
	VtAlert(STR0022,STR0002,.t.,4000,3) 	 //"Endereco incorreto"###"Aviso"
	cArmazem  := Space(TamSX3("B2_LOCAL")[1])   // somente esta sendo atribuido para posteriormnete ser comparado com aItens
	cEndereco := Space(TamSX3("BF_LOCALIZ")[1])
	VTClearGet("cArmazem")
	VTClearGet("cEndereco")
	VTGetSetFocus("cArmazem")
	VtKeyboard(Chr(20))  // zera o get
	Return .F.
EndIf
If !CBEndLib(cArmazem,cEndereco) // verifica se o endereco esta liberado ou bloqueado
	VtAlert(STR0023,STR0002,.t.,4000,3) //"Endereco Bloqueado."###"Aviso"
	VTClearGet("cArmazem")
	VTClearGet("cEndereco")
	VTGetSetFocus("cArmazem")
	VtKeyboard(Chr(20))  // zera o get
	Return .F.
EndIf
Return .T.
//---TERMINO das funcoes referente ao endereco -------------------------

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Funcao    ³  Tela    ³ Autor ³ ACD                   ³ Data ³ 14/08/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descricao ³ Montagem da tela Principal                                 ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function Tela(lMV_OSEP2UN)
Local cSay
Local nQtdCx
Local nQtdRet := 0

VTClear()
cSay := STR0024 //'Separe '
If lMV_OSEP2UN
	SB1->(DbSetOrder(1))
	SB1->(DbSeek(xFilial("SB1")+aItens[_POSCODPRO]))
	nQtdCX := CBQEmb()
	If ExistBlock('CBRQEESP')
		nQtdRet := ExecBlock('CBRQEESP',,,SB1->B1_COD)
		If ValType(nQtdRet)=="N"
			nQtdCX  := nQtdRet
		EndIf
	EndIf
	If aItens[_POSSLDSEP]/nQtdCX < 1
		cSay +=Padr(Alltrim(Str(aItens[_POSSLDSEP],TamSx3("CB8_QTDORI")[1],TamSx3("CB8_QTDORI")[2]))+If(aItens[_POSSLDSEP]==1,STR0025,STR0026),20) //' item '###' itens '
	Else
		cSay +=Padr(Alltrim(Str(aItens[_POSSLDSEP]/nQtdCX,TamSx3("CB8_QTDORI")[1],TamSx3("CB8_QTDORI")[2]))+If(aItens[_POSSLDSEP]/nQtdCX==1,STR0027,STR0028),20) //" Volume"###" Volumes"
	EndIf
Else
	cSay +=Padr(Alltrim(Str(aItens[_POSSLDSEP],TamSx3("CB8_QTDORI")[1],TamSx3("CB8_QTDORI")[2]))+If(aItens[_POSSLDSEP]==1,STR0025,STR0026),20) //' item '###' itens '
EndIf
SB1->(DbSetOrder(1))
SB1->(DbSeek(xFilial("SB1")+aItens[_POSCODPRO]))

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³Ponto de Entrada na montagem da tela ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
If ExistBlock("A165TELA")
	ExecBlock("A165TELA",.F.,.F.)
Else	  
	@ 0,0 VTSay cSay
	@ 1,0 VTSay aItens[_POSCODPRO]
	@ 2,0 VTSay Left(SB1->B1_DESC,20)
	If Rastro(aItens[_POSCODPRO],"L")
		@ 3,0 VTSay STR0029 //'Lote       '
		@ 4,0 VTSay aItens[_POSLOTECTL]
	ElseIf Rastro(aItens[_POSCODPRO],"S")
		@ 3,0 VTSay STR0030 //'Lote       SubLote'
		@ 4,0 VTSay aItens[_POSLOTECTL]+' '+aItens[_POSNUMLOTE]
	EndIf
	If ! Empty(aItens[_POSNUMSERIE])
		@ 5,0 VTSay STR0031 //"Numero de Serie "
		@ 6,0 VTSay aItens[_POSNUMSERIE]
		@ 7,0 VTPause STR0019 //'Enter para continuar'
		VTClear()
	EndIf
EndIf	
Return

//---TERMINO das funcao de montagem da tela princiapal ------

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡ao    ³EtiqProduto³ Autor ³ ACD                   ³ Data ³ 14/08/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡ao ³Leitura das etiquetas de produto                             ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

//---INICIO das funcoes referente a leitura da etiqueta de produto------

Static Function EtiqProduto()
Local lGranel    := ! CBProdUnit(aItens[_POSCODPRO])  // granel
Local lSerie     := ! Empty(aItens[_POSNUMSERIE])
Local cEtiqCaixa := Space(TamSx3("CB0_CODET2")[1])
Local cEtiqAvulsa:= Space(TamSx3("CB0_CODET2")[1])
Local cEtiqProd  := Space(TamSx3("CB0_CODET2")[1])
Local cProduto   := IIf( FindFunction( 'AcdGTamETQ' ), AcdGTamETQ(), Space(48) )
Local nQtde      := 1
Local bKey16     := VtSetKey(16)
IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf

VtSetKey(16,{|| _lPulaItem:= .t.,VtKeyboard(CHR(13)) },STR0032)  // CTRL+P //"Pula Item"

// solicitando a leitura da etiqueta
If lGranel
	If lVT100B // GetMv("MV_RF4X20")
		@ 0,0 VTSay STR0033 //"Leia a caixa"
		@ 1,0 VtGet cEtiqCaixa pict '@!' Valid VldCaixa(cEtiqCaixa,nQtde)
	Else
		@ 6,0 VTSay STR0033 //"Leia a caixa"
		@ 7,0 VtGet cEtiqCaixa pict '@!' Valid VldCaixa(cEtiqCaixa,nQtde)
	Endif
	VTRead
	If VTLastkey() == 27
		VtSetKey(16, bKey16,"")
		Return .f.
	EndIf
	If lVT100B // GetMv("MV_RF4X20")
		@ 0,0 VTClear to 7,19
		@ 0,0 VTSay STR0034 //"Leia a etiq. avulsa"
		@ 1,0 VtGet cEtiqAvulsa pict '@!' Valid VldEtiqAvulsa(cEtiqAvulsa)
	Else
		@ 6,0 VTClear to 7,19
		@ 6,0 VTSay STR0034 //"Leia a etiq. avulsa"
		@ 7,0 VtGet cEtiqAvulsa pict '@!' Valid VldEtiqAvulsa(cEtiqAvulsa)
	Endif
Else
	If lCB001
		If lSerie
			@ 1,0 VTSay STR0035 //'Leia a etiqueta'
			@ 2,0 VTGet cEtiqProd pict '@!' Valid VldEti(cEtiqProd,nQtde)
		Else
			If lVT100B // GetMv("MV_RF4X20")
				@ 0,0 VTSay STR0035 //'Leia a etiqueta'
				@ 1,0 VTGet cEtiqProd pict '@!' Valid VldEti(cEtiqProd,nQtde)
			Else
				@ 6,0 VTSay STR0035 //'Leia a etiqueta'
				@ 7,0 VTGet cEtiqProd pict '@!' Valid VldEti(cEtiqProd,nQtde)
			Endif
		EndIf
	Else
		If lSerie
			@ 1,0 VTSay STR0036 VtGet nQtde pict cPictQtde valid VldQtde(nQtde,lSerie) when (lForcaQtd .or. VtLastkey()==5)  //'Qtde '
			@ 2,0 VTSay STR0037 //'Leia o produto'
			@ 3,0 VTGet cProduto pict '@!' VALID VTLastkey() == 5 .or. VldProd(cProduto,nQtde,lSerie)
		Else
			If lVT100B // GetMv("MV_RF4X20")
				@ 0,0 VTSay STR0036 VtGet nQtde pict cPictQtde valid VldQtde(nQtde,lSerie) when (lForcaQtd .or. VtLastkey()==5)  //'Qtde '
				@ 1,0 VTSay STR0037 //'Leia o produto'
				@ 2,0 VTGet cProduto pict '@!' VALID VTLastkey() == 5 .or. VldProd(cProduto,nQtde,lSerie)
			Else
				@ 5,0 VTSay STR0036 VtGet nQtde pict cPictQtde valid VldQtde(nQtde,lSerie) when (lForcaQtd .or. VtLastkey()==5)  //'Qtde '
				@ 6,0 VTSay STR0037 //'Leia o produto'
				@ 7,0 VTGet cProduto pict '@!' VALID VTLastkey() == 5 .or. VldProd(cProduto,nQtde,lSerie)
			Endif
		EndIf
	Endif
EndIf
VTRead
VtSetKey(16, bKey16,"")
If VTLastkey() == 27
	Return .f.
EndIf
Return .t.

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Funcao    ³  VldCaixa ³ Autor ³ ACD                   ³ Data ³ 14/08/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descricao ³Validacao das etiquetas a granel                             ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function VldCaixa(cEtiqAvulsa,nQtde)
Local cCodPro
Local aRet
Local aRetPE    := {}
Local lACD165VE := ExistBlock('ACD165VE')

If _lPulaItem
	Return .t.
EndIf
If Empty(cEtiqAvulsa)
	Return .f.
EndIf

If lCB001
	aRet := CBRetEti(cEtiqAvulsa,"01")
	If lACD165VE
		aRetPE := ExecBlock('ACD165VE',,,{aRet,nQtde})
		If	ValType(aRetPE)=="A"
			aRet := aRetPE
		EndIf
	EndIf
	If Len(aRet) == 0
		VtAlert(STR0021,STR0002,.t.,4000,3)  //"Etiqueta invalida"###"Aviso"
		VtKeyboard(Chr(20))  // zera o get
		Return .F.
	EndIf
	cCodPro := aRet[1]
	If ! Empty(aRet[2])
		VtAlert(STR0021,STR0002,.t.,4000,3)  //"Etiqueta invalida"###"Aviso"
		VtKeyboard(Chr(20))  // zera o get
		Return .F.
	EndIf
Else
	If !CBLoad128(@cEtiqAvulsa)
		VtKeyboard(Chr(20))  // zera o get
		Return .F.
	EndIf
	If ! CbRetTipo(cEtiqAvulsa) $ "EAN8OU13-EAN14-EAN128"
		VtAlert(STR0021,STR0002,.t.,4000,3)  //"Etiqueta invalida"###"Aviso"
		VtKeyboard(Chr(20))  // zera o get
		Return .F.
	EndIf
	aRet := CBRetEtiEan(cEtiqAvulsa)
	If lACD165VE
		aRetPE := ExecBlock('ACD165VE',,,{aRet,nQtde})
		If	ValType(aRetPE)=="A"
			aRet := aRetPE
		EndIf
	Endif
	If len(aRet) == 0
		VtAlert(STR0021,STR0002,.t.,4000,3)  //"Etiqueta invalida"###"Aviso"
		VtKeyboard(Chr(20))  // zera o get
		Return .F.
	EndIf
	cCodPro := aRet[1]
EndIf

If ! CBProdLib(cArmazem,cCodPro)
	VTKeyBoard(chr(20))
	Return .f.
EndIf
If !(aItens[_POSCODPRO] == cCodPro)
	VtAlert(STR0038,STR0002,.t.,4000,3)  //"Etiqueta de produto diferente"###"Aviso"
	VtKeyboard(Chr(20))  // zera o get
	Return .F.
EndIf
Return .T.

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Funcao    ³VldEtiqAvulsa³ Autor ³ ACD                   ³ Data ³ 14/08/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descricao ³Validacao das etiquetas avulsas                                ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function VldEtiqAvulsa(cEtiqAvulsa)
Local nQE
Local cLote    := Space(TamSX3("B8_LOTECTL")[1])
Local cSLote   := Space(TamSX3("B8_NUMLOTE")[1])

Local aEtiqueta:={}
If _lPulaItem
	Return .t.
EndIf
If Empty(cEtiqAvulsa)
	Return .f.
EndIf

If Len(CBRetEti(cEtiqAvulsa)) > 0
	VtAlert(STR0021,STR0002,.t.,4000,3)  //"Etiqueta invalida"###"Aviso"
	VtKeyboard(Chr(20))  // zera o get
	Return .F.
EndIf

CBGrvEti("01",{aItens[_POSCODPRO],0,cCodOpe},Padr(cEtiqAvulsa,10))
nQE  :=CBQtdEmb(aItens[_POSCODPRO])
If Empty(nQE)
	VtAlert(STR0039,STR0002,.t.,4000,3) //"Quantidade invalida"###"Aviso"
	VtKeyboard(Chr(20))  // zera o get
	RecLock('CB0',.f.)
	CB0->(DbDelete())
	CB0->(MSUnlock())
	Return .F.
EndIf
If nQE > aItens[_POSSLDSEP]
	VtAlert(STR0040,STR0002,.t.,4000,3)  //"Quantidade maior que solicitado"###"Aviso"
	VtKeyboard(Chr(20))  // zera o get
	RecLock('CB0',.f.)
	CB0->(DbDelete())
	CB0->(MSUnlock())
	Return .f.
EndIf
CBGrvEti("01",{NIL,nQE,NIL},Padr(cEtiqAvulsa,10))
If CBRastro(aItens[_POSCODPRO],@cLote,@cSLote)
	If aItens[_POSCFLOTE] == "1"
		If ! cLote+cSLote == aItens[_POSLOTECTL]+aItens[_POSNUMLOTE]
			VtAlert(STR0041,STR0002,.t.,4000,3)  //"Lote invalido"###"Aviso"
			VtKeyboard(Chr(20))  // zera o get
			Return .F.
		EndIf
	EndIf
	If  GetMv('MV_ESTNEG') =='N'
		If Localiza(aItens[_POSCODPRO])
			nSaldo := SaldoSBF(aItens[_POSARMAZEM],aItens[_POSENDERECO],aItens[_POSCODPRO],aItens[_POSNUMSERIE],cLote,cSLote,.T.)
		Else
			SB2->(DbSetOrder(1))
			SB2->(DbSeek(xFilial("SB2")+aItens[_POSCODPRO]+aItens[_POSARMAZEM]))
			nSaldo := SaldoSB2()
		EndIf
		If nQE > nSaldo+aItens[_POSSLDSEP] //Saldo disponivel "+" CB8->CB8_SALDOS... Saldo da ordem de separação.
			VtAlert(STR0042,STR0002,.t.,4000,3)  //"Saldo em estoque insuficiente"###"Aviso"
			VtKeyboard(Chr(20))  // zera o get
			Return .F.
		EndIf
	Else
		If ! CBExistLot(aItens[_POSCODPRO],aItens[_POSARMAZEM],aItens[_POSENDERECO],cLote,cSLote)
			VtAlert(STR0043,STR0002,.t.,4000,3)  //"Lote nao existe"###"Aviso"
			VtKeyboard(Chr(20))  // zera o get
			Return .F.
		EndIf
	EndIf
	aEtiqueta:= CBRetEti(Padr(cEtiqAvulsa,10),"01")
	aEtiqueta[16]:=cLote
	aEtiqueta[17]:=cSLote
	CBGrvEti("01",aEtiqueta,Padr(cEtiqAvulsa,10))
Else
	VTKeyBoard(chr(20))
	Return .F.
EndIf 

AADD(aLoteNew ,cLote)
AADD(aSLoteNew,cSLote)
AADD(aQtdLida ,nQE)
Return .T.

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Funcao    ³VldEti³ Autor ³ ACD                   ³ Data ³ 14/08/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descricao ³Validacao das etiquetas                                 ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
Static Function VldEti(cEtiqProd,nQtde)
Local cLote 	:= Space(TamSX3("B8_LOTECTL")[1])
Local cSLote	:= Space(TamSX3("B8_NUMLOTE")[1])
Local cNewEtiq  := ""
Local aRetPE    := {}
Local aEtiqueta := {}
Local aNewEtiq  := {}
Local aItensTemp:= {}
Local lACD165VE := ExistBlock('ACD165VE')
Local nX		:= 0 


//- Verifica se a etiqueta pertence a um PALLET
aItensPallet := CBItPallet(cEtiqProd)
	
If Len(aItensPallet) == 0
	aItensPallet:={cEtiqProd}
	lIsPallet := .F.
EndIf


For nX:= 1 to Len(aItensPallet)
	cEtiqProd:= aItensPallet[nX]
	If ! CarregaCB8(aItensTemp) .or. (aItensTemp[_POSARMAZEM]+aItensTemp[_POSENDERECO]+aItensTemp[_POSCODPRO]+aItensTemp[_POSLOTECTL]+aItensTemp[_POSNUMLOTE]+aItensTemp[_POSNUMSERIE] # ;
		aItens[_POSARMAZEM]+aItens[_POSENDERECO]+aItens[_POSCODPRO]+aItens[_POSLOTECTL]+aItens[_POSNUMLOTE]+aItens[_POSNUMSERIE])
		VtAlert(STR0044,STR0002,.T.,5000,3)  //"Separacao deste produto foi finalizado por outro usuario"###"Aviso"
		Return .T.
	EndIf
	If _lPulaItem
		Return .T.
	EndIf
	If Empty(cEtiqProd)
		Return .F.
	EndIf
	
	aEtiqueta:= CBRetEti(cEtiqProd,"01")
	If lACD165VE
		aRetPE := ExecBlock('ACD165VE',,,{aEtiqueta,nQtde})
		If	ValType(aRetPE)=="A"
			aEtiqueta := aRetPE
		EndIf
	EndIf
	If Empty(aEtiqueta)
		VtAlert(STR0021,STR0002,.t.,4000,3)  //"Etiqueta invalida"###"Aviso"
		VtKeyboard(Chr(20))  // zera o get
		Return .F.
	EndIf
	CB8->(DbSetOrder(1))
	If CB8->(DbSeek(xFilial("CB8")+cOrdSep+CB0->CB0_CODETI))
		VtAlert(STR0045,STR0002,.t.,4000,3)  //"Etiqueta ja foi lida"###"Aviso"
		VtKeyboard(Chr(20))  // zera o get
		Return .F.
	EndIf
	If A166VldCB9(aEtiqueta[1], cEtiqProd, .T.)
		VtAlert(STR0045,STR0002,.t.,4000,3)  //"Etiqueta ja foi lida"###"Aviso"
		VtKeyboard(Chr(20))  // zera o get
		Return .F.
	EndIf
	If ! aItens[_POSARMAZEM]+ aItens[_POSENDERECO] == aEtiqueta[10]+aEtiqueta[9]
		VtAlert(STR0046,STR0002,.t.,4000,3)  //"Endereco diferente"###"Aviso"
		VtKeyboard(Chr(20))  // zera o get
		Return .F.
	EndIf
	If !  aItens[_POSCODPRO] == aEtiqueta[1]
		VtAlert(STR0047,STR0002,.t.,4000,3)  //"Produto diferente"###"Aviso"
		VtKeyboard(Chr(20))  // zera o get
		Return .F.
	EndIf
	If ! CBProdLib(aEtiqueta[10],aEtiqueta[1])
		VTKeyBoard(chr(20))
		Return .F.
	EndIf

	cLote  := aEtiqueta[16]
	cSLote := aEtiqueta[17]
	If aItens[_POSCFLOTE] =="1"
		If ! cLote+cSLote == aItens[_POSLOTECTL]+aItens[_POSNUMLOTE]
			VtAlert(STR0041,STR0002,.t.,4000,3) //"Lote invalido"###"Aviso"
			VtKeyboard(Chr(20))  // zera o get
			Return .F.
		EndIf
	EndIf
	If CB0->CB0_NUMSER # aItens[_POSNUMSERIE]
		VtBeep(3)
		VtAlert(STR0048,STR0002,.t.,3000) //"Etiqueta Invalida !!!"###"Aviso"
		VtAlert(STR0049+aItens[_POSNUMSERIE]   ,STR0002,.t.,4000) //"Informe a etiqueta com o Numero de Serie "###"Aviso"
		VtKeyboard(Chr(20))  // zera o get
		Return .F.
	EndIf
	If ExistBlock('ACD170VE')
		aEtiqueta:=ExecBlock('ACD170VE',,,aEtiqueta)
		If Empty(aEtiqueta)
			Return .F.
		EndIf
	EndIf
	If CBQtdVar(aEtiqueta[1]) .and. aEtiqueta[2] > aItens[_POSSLDSEP]
		cNewEtiq:= GeraNewEti(aItens[_POSSLDSEP])
		aNewEtiq:= CBRetEti(cNewEtiq,"01")
	ElseIf aEtiqueta[2] > aItens[_POSSLDSEP]
		VtAlert(STR0050,STR0002,.t.,4000,3)  //"Quantidade maior que necessario"###"Aviso"
		VtKeyboard(Chr(20))  // zera o get
		Return .F.
	EndIf
	If Empty(aNewEtiq)
		AADD(aQtdLida ,aEtiqueta[2])
	   	AADD(aLoteNew ,cLote)
		AADD(aSLoteNew,cSLote)
	Else
		cEtiqProd:= CB0->CB0_CODETI
   		AADD(aQtdLida ,aNewEtiq[2])
	  	AADD(aLoteNew ,aNewEtiq[16])
	  	AADD(aSLoteNew,aNewEtiq[17])
	EndIf
Next nX	
Return .T.

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡ao    ³VldProd   ³ Autor ³ ACD                   ³ Data ³ 14/08/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡ao ³Validacao do produto                                        ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
Static Function VldProd(cProduto,nQtde,lSerie)
Local aRet   := {}
Local aRetPE := {}
Local nSaldo := 0
Local cLote  := Space(TamSX3("B8_LOTECTL")[1]) 
Local cSLote := Space(TamSX3("B8_NUMLOTE")[1])
Local cNumSer:= Space(TamSX3("BF_NUMSERI")[1])
Local lACD165VE:=ExistBlock('ACD165VE')

If _lPulaItem
	Return .T.
EndIf
If Empty(cProduto)
	Return .F.
EndIf

If !CBLoad128(@cProduto)
	VtKeyboard(Chr(20))  // zera o get
	Return .F.
EndIf

aItensPallet:={cProduto}

If ! CbRetTipo(cProduto) $ "EAN8OU13-EAN14-EAN128"
	VtAlert(STR0021,STR0002,.t.,4000,3)  //"Etiqueta invalida"###"Aviso"
	VtKeyboard(Chr(20))  // zera o get
	Return .F.
EndIf
aRet := CBRetEtiEan(cProduto)
If lACD165VE
	aRetPE := ExecBlock('ACD165VE',,,{aRet,nQtde})
	If	ValType(aRetPE)=="A"
		aRet := aRetPE
	EndIf
EndIf
If len(aRet) == 0
	VtAlert(STR0021,STR0002,.t.,4000,3)  //"Etiqueta invalida"###"Aviso"
	VtKeyboard(Chr(20))  // zera o get
	Return .F.
EndIf
If !aItens[_POSCODPRO] == aRet[1]
	VtAlert(STR0047,STR0002,.t.,4000,3)  //"Produto diferente"###"Aviso"
	VtKeyboard(Chr(20))  // zera o get
	Return .F.
EndIf
cLote := aRet[3]
If ! CBProdLib(cArmazem,aRet[1])
	VTKeyBoard(chr(20))
	Return .F.
EndIf
cNumSer:= aRet[5]	 // Numero de Serie
If ! CBRastro(aItens[_POSCODPRO],@cLote,@cSLote)
	VTKeyBoard(chr(20))
	Return .F.
EndIf

If aItens[_POSCFLOTE] =="1"
	If ! cLote+cSLote == aItens[_POSLOTECTL]+aItens[_POSNUMLOTE]
		VtAlert(STR0041,STR0002,.t.,4000,3)  //"Lote invalido"###"Aviso"
		VtKeyboard(Chr(20))  // zera o get
		Return .F.
	EndIf
EndIf
If lSerie
	If ! CBNumSer(@cNumSer,aItens[_POSNUMSERIE])
		VTKeyBoard(chr(20))
		Return .F.
	EndIf
EndIf
If aRet[2]*nQtde > aItens[_POSSLDSEP]
	VtAlert(STR0050,STR0002,.t.,4000,3)  //"Quantidade maior que necessario"###"Aviso"
	VtKeyboard(Chr(20))  // zera o get
	Return .F.
EndIf
If CB7->CB7_ORIGEM # "2"
	If GetMv('MV_ESTNEG') =='N'
		If Localiza(aItens[_POSCODPRO])
			nSaldo := SaldoSBF(aItens[_POSARMAZEM],aItens[_POSENDERECO],aItens[_POSCODPRO],aItens[_POSNUMSERIE],cLote,cSLote,.T.)
		Else
			SB2->(DbSetOrder(1))
			SB2->(DbSeek(xFilial("SB2")+aItens[_POSCODPRO]+aItens[_POSARMAZEM]))
			nSaldo := SaldoSB2()
		EndIf
		If aRet[2]*nQtde > nSaldo+aItens[_POSSLDSEP] //Saldo disponivel "+" CB8->CB8_SALDOS... Saldo da ordem de separação.
			VtAlert(STR0042,STR0002,.t.,4000,3)  //"Saldo em estoque insuficiente"###"Aviso"
			VtKeyboard(Chr(20))  // zera o get
			Return .F.
		EndIf
	Else
		If ! CBExistLot(aItens[_POSCODPRO],aItens[_POSARMAZEM],aItens[_POSENDERECO],cLote,cSLote)
			VtAlert(STR0043,STR0002,.t.,4000,3)  //"Lote nao existe"###"Aviso"
			VtKeyboard(Chr(20))  // zera o get
			Return .F.
		EndIf
	EndIf
EndIf
AADD(aQtdLida ,aRet[2]*nQtde)
AADD(aLoteNew ,cLote)
AADD(aSLoteNew,cSLote)
Return .t.

//---TERMINO das funcoes referente a leitura da etiqueta de produto-----

//---INICIO das funcoes referente ao pulo do item -----
/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Funcao    ³ PulaItem ³ Autor ³ ACD                   ³ Data ³ 14/08/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descricao ³ Pula o item                                                ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function PulaItem()
Local aRec, aSvTela,cOcoSep
Local nX := 0
If ! _lPulaItem
	Return .F.
EndIf
IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf
_lPulaItem := .F.

aSvTela := VtSave()
CB8->(DbGoto(aItens[_POSRECNO,1]))
cOcoSep := CB8->CB8_OCOSEP
CB4->(DbSetOrder(1))
CB4->(DbSeek(xFilial("CB4")+cOcoSep))
VTClear
If lVT100B // GetMv("MV_RF4X20")
	@ 0,0 VTSay STR0051 //'Informe o codigo'
	@ 1,0 VTSay STR0052 //'da divergencia:'
	@ 2,0 VtGet cOcoSep pict '@!' Valid VldOcoSep(cOcoSep) F3 "CB4"
Else
	@ 2,0 VTSay STR0051 //'Informe o codigo'
	@ 3,0 VTSay STR0052 //'da divergencia:'
	@ 4,0 VtGet cOcoSep pict '@!' Valid VldOcoSep(cOcoSep) F3 "CB4"
Endif
VtRead()
VtRestore(,,,,aSvTela)
If VtLastKey() == 27
	Return .F.
EndIf
For nX:= 1 to len(aItens[_POSRECNO])
	CB8->(DbGoto(aItens[_POSRECNO,nX]))
	RecLock("CB8")
	CB8->CB8_OCOSEP := cOcoSep
	CB8->(MsUnlock())
Next
If CB7->CB7_DIVERG # "1"   // marca divergencia na ORDEM DE SEPARACAO para que esta seja arrumada
	CB7->(RecLock("CB7"))
	CB7->CB7_DIVERG := "1"  // sim
	CB7->(MsUnlock())
EndIf
Return .T.

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡Æo    ³ VldOcoSep³ Autor ³ ACD                   ³ Data ³ 14/08/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡Æo ³ Validacao da ocorrencia informada no pulo do item          ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function VldOcoSep(cOcoSep)
Local lRet := .T.
If	Empty(cOcoSep)
	VtKeyBoard(chr(23))
EndIf

CB4->(DBSetOrder(1))
If	!CB4->(DbSeek(xFilial("CB4")+cOcoSep))
	VtAlert(STR0053,STR0002,.t.,4000,3)  //"Ocorrencia nao cadastrada"###"Aviso"
	VtKeyboard(Chr(20))  // zera o get
	lRet := .F.
EndIf
Return lRet

//---TERMINO das funcoes referente ao pulo do item -----

//---INICIO das funcoes referente a gravacao do item -----

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡Æo    ³   Grava  ³ Autor ³ ACD                   ³ Data ³ 14/08/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡Æo ³ Realiza a gravacao dos itens lidos                         ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function Grava(cCodCB0,nX)
Local nY := 0
Local nSaldoCB8 := 0
Local nLenNX    := Len(aLoteNew)
Local nTamVol   := TamSX3("CB9_VOLUME")[1]
DefauLt cCodCB0 := Space(10)

For nY:= 1 to len(aItens[_POSRECNO])
	CB8->(DbGoto(aItens[_POSRECNO,nY]))
	nSaldoCB8:= CB8->CB8_SALDOS
	If empty(nSaldoCB8)
		Loop
	EndIf
	CB9->(DbSetOrder(10)) // CB9_FILIAL+CB9_ORDSEP+CB9_ITESEP+CB9_PROD+CB9_LOCAL+CB9_LCALIZ+CB9_LOTECT+CB9_NUMLOT+
							 // CB9_NUMSER+CB9_LOTSUG+CB9_SLOTSU+CB9_VOLUME+CB9_CODETI+CB9_PEDIDO
	CB9->(DbSeek(xFilial("CB9")+CB8->(CB8_ORDSEP+CB8_ITEM+CB8_PROD+CB8_LOCAL+CB8_LCALIZ+ IIF(!lCB001,aLoteNew[nLenNX]+aSLoteNew[nLenNX],aLoteNew[nX]+aSLoteNew[nX])+;
									CB8_NUMSER+CB8_LOTECT+CB8_NUMLOT)+Space(nTamVol)+cCodCB0+CB8->CB8_PEDIDO))
	If ! CB9->(Eof())
		If lCB001
			VtAlert(STR0054,STR0002,.t.,4000,3)  //"Codigo ja Lido"###"Aviso"
			Return .F.
		EndIf
		RecLock("CB9",.F.)
	Else
		RecLock("CB9",.T.)
		CB9->CB9_FILIAL := xFilial("CB9")
		CB9->CB9_ORDSEP := CB7->CB7_ORDSEP
		CB9->CB9_CODETI := cCodCB0
		CB9->CB9_PROD   := CB8->CB8_PROD
		CB9->CB9_CODSEP := CB7->CB7_CODOPE
		CB9->CB9_ITESEP := CB8->CB8_ITEM
		CB9->CB9_SEQUEN := CB8->CB8_SEQUEN
		CB9->CB9_LOCAL  := CB8->CB8_LOCAL
		CB9->CB9_LCALIZ := CB8->CB8_LCALIZ
		If !(AllTrim(aItensPallet[nX]) == AllTrim(CB8->CB8_PROD))	//Sendo aItensPallet[nX] == CB8_PROD, nao existe tratamento de Pallet
	   																		//para o produto, dessa forma grava o Lote+Sublote conforme a CB8
	   		If !lCB001
		   		CB9->CB9_LOTECT := aLoteNew[nLenNX]
		  		CB9->CB9_NUMLOT := aSLoteNew[nLenNX]
		  	Else
		  		CB9->CB9_LOTECT := aLoteNew[nLenNX]
		  		CB9->CB9_NUMLOT := aSLoteNew[nLenNX]
		  	EndIf
	  	ElseIf !("10" $ CB7->CB7_TIPEXP)
	  		CB9->CB9_LOTECT := aLoteNew[Len(aLoteNew)]
	  		CB9->CB9_NUMLOT := aSLoteNew[Len(aSLoteNew)]
	 	Else
	 		CB9->CB9_LOTECT := CB8->CB8_LOTECT
	  		CB9->CB9_NUMLOT := CB8->CB8_NUMLOT
	  	EndIf
		CB9->CB9_NUMSER := CB8->CB8_NUMSER
		CB9->CB9_LOTSUG := CB8->CB8_LOTECT
		CB9->CB9_SLOTSU := CB8->CB8_NUMLOT
		CB9->CB9_PEDIDO := CB8->CB8_PEDIDO
	EndIf
	CB9->CB9_QTESEP += If(aQtdLida[nX] > nSaldoCB8,nSaldoCB8,aQtdLida[nX])
	CB9->CB9_STATUS := "1"  // EM ABERTO
	CB9->(MsUnlock())
	RecLock("CB8")
	CB8->CB8_SALDOS -=  If(aQtdLida[nX] > nSaldoCB8,nSaldoCB8,aQtdLida[nX])
	CB8->(MsUnlock())
	aItens[_POSSLDSEP] -= If(aQtdLida[nX] > nSaldoCB8,nSaldoCB8,aQtdLida[nX])
	aQtdLida[nX] -= If(aQtdLida[nX] > nSaldoCB8,nSaldoCB8,aQtdLida[nX])
	If aQtdLida[nX] == 0
		Exit
	Endif
Next
Return .T.

//---TERMINO das funcoes referente a gravacao do item -----

//---INICIO da funcao referente ao informa -----

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡Æo    ³ Informa  ³ Autor ³ ACD                   ³ Data ³ 14/08/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡Æo ³ Informa as etiquetas lidas                                 ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function Informa()
Local aCab,aSize,aTela := VTSAVE()
Local nOpc
Local aTemp:={}
Local nX := 0
Local aItenstmp:={}

VTCLear()
@ 0,0 VTSAY STR0012 //"Informacoes"
@ 1,0 VTSay STR0055              //'Selecione:'
nOpc:=VTaChoice(3,0,4,VTMaxCol(),{STR0056,STR0057})  //"A Separar"###"Separados"
VtRestore(,,,,aTela)
If VtLastKey() == 27
	Return
EndIf
VTCLear()
If nOpc ==1
	CB8->(DbSetOrder(3))
	CB8->(DbSeek(xFilial("CB8")+cOrdSep))
	While ! CB8->(Eof()) .and. CB8->CB8_FILIAL == xFilial("CB8") .and. CB8->CB8_ORDSEP == cOrdSep
		nPos := Ascan(aItenstmp,{|x| x[3]+x[4]+x[5]+x[8]+x[9]+X[12] == CB8->(CB8_PROD+CB8_LOCAL+CB8->CB8_LCALIZ+CB8_LOTECT+CB8_NUMLOT+CB8_CFLOTE)})
		If  nPos== 0
			aadd(aItenstmp,{{CB8->(Recno())},;
			CB8->CB8_ITEM,;
			CB8->CB8_PROD,;
			CB8->CB8_LOCAL,;
			CB8->CB8_LCALIZ,;
			CB8->CB8_SALDOS,;
			CB8->CB8_SLDPRE,;
			CB8->CB8_LOTECT,;
			CB8->CB8_NUMLOT,;
			NIL,;
			NIL,;
			CB8->CB8_CFLOTE })
		Else
			If Ascan(aItenstmp[nPos,1],{|x| x==CB8->(Recno())})==0
				aadd(aItenstmp[nPos,1],CB8->(Recno()))
			EndIf
			aItenstmp[nPos,6]+=CB8->CB8_SALDOS
			aItenstmp[nPos,7]+=CB8->CB8_SLDPRE
		EndIf
		CB8->(DbSkip())
	EndDo
	For nX:= 1 to Len(aItenstmp)
		aadd(aTemp,{aItenstmp[nX,_POSCODPRO],;
		Transform(aItenstmp[nX,_POSQTDSEP],cPictQtde),;
		Transform(aItenstmp[nX,_POSSLDSEP],cPictQtde),;
		aItenstmp[nX,_POSARMAZEM],;
		aItenstmp[nX,_POSENDERECO],;
		aItenstmp[nX,_POSLOTECTL],;
		aItenstmp[nX,_POSNUMLOTE]})
	Next
	aCab  := {STR0058,STR0059,STR0060,STR0061,STR0062,STR0063,STR0064} //"Produto"###"Quantidade"###"Saldo"###"Armazem"###"Endereco"###"Lote"###"Sub-Lote"
	aSize := {15,len(cPictQtde),len(cPictQtde),8,15,10,8}
	VTaBrowse(0,0,VtMaxRow(),vtMaxCol(),aCab,aTemp,aSize)
Else
	CB9->(DbSetOrder(6))
	CB9->(DbSeek(xFilial("CB9")+cOrdSep))
	While CB9->(! Eof() .and. CB9_FILIAL+CB9_ORDSEP == xFilial("CB9")+cOrdSep)
		nPos:= Ascan(aTemp,{|x| x[1]+x[3]+x[4]+x[5]+x[6]+x[7] == CB9->(CB9_PROD+CB9_LOCAL+CB9_LCALIZ+CB9_LOTECT+CB9_NUMLOT+CB9_CODETI)})
		If nPos == 0
			aadd(aTemp,{CB9->CB9_PROD,0,CB9->CB9_LOCAL,CB9->CB9_LCALIZ,CB9->CB9_LOTECT,CB9->CB9_NUMLOT,CB9->CB9_CODETI})
			nPos:= Len(aTemp)
		Endif
		aTemp[nPos,2]+= CB9->CB9_QTESEP
		CB9->(DbSkip())
	Enddo
	aTemp := aSort(aTemp,,,{|x,y| x[1]+x[3]+x[4]+x[5]+x[6]+x[7] < y[1]+y[3]+y[4]+y[5]+y[6]+y[7]})
	aCab  := {STR0058,STR0059,STR0061,STR0062,STR0063,STR0064,STR0065} //"Produto"###"Quantidade"###"Armazem"###"Endereco"###"Lote"###"Sub-Lote"###"Etiqueta"
	aSize := {15,len(cPictQtde),08,15,10,8,20}
	VTaBrowse(0,0,VtMaxRow(),vtMaxCol(),aCab,aTemp,aSize)
EndIf
VtRestore(,,,,aTela)
Return

//---TERMINO da funcao referente ao informa -----

//---INICIO da funcao que analisa se terminou e gera as ordens de separacoes  -----

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡Æo    ³AnalisaTermino³ Autor ³ ACD                   ³ Data ³ 14/08/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡Æo ³Analisa as Ordens de separacao a ser geradas                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function AnalisaTermino()
Local cOrigem    := CB7->CB7_ORIGEM
Local aSep       := {}
Local np1        := 0
Local nTamPed    := TamSX3("C6_NUM")[1]
Local aOrdSep    := {}
Local cPedido    := Space(nTamPed)
Local cArm       := Space(Tamsx3("B1_LOCPAD")[1])
Local cMsgParcial:= ""
Private cTipExp  := CB7->CB7_TIPEXP
Private cCond    := CB7->CB7_COND
Private cLojaEnt := CB7->CB7_LOJENT
Private cTransp  := CB7->CB7_TRANSP
Private cAgreg   := CB7->CB7_AGREG

CB8->(DbSetOrder(3))
CB8->(DbSeek(xFilial("CB8")+cOrdSep))
While !CB8->(Eof()) .and. CB8->CB8_FILIAL == xFilial("CB8") .and. CB8->CB8_ORDSEP == cOrdSep
	If !Empty(CB8->CB8_SALDOS) .Or. ! Empty(CB8->CB8_OCOSEP)
		cMsgParcial:=STR0066 //" Parciais"
	EndIf
	If	CB8->CB8_SLDPRE==CB8->CB8_SALDOS
		CB8->(DbSkip())
		Loop
	EndIf

	If	cOrigem == "1" // Pedido
		SC6->(dbSetOrder( 1 ))
		SC6->(DbSeek(xFilial("SC6") + CB8->CB8_PEDIDO + CB8->CB8_ITEM + CB8->CB8_PROD))
		If	"11*" $ cTipExp
			cPedido := Space(nTamPed)
		Else
			cPedido := CB8->CB8_PEDIDO
		EndIf
		If	"08*" $ cTipExp
			cArm :=Space(Tamsx3("B1_LOCPAD")[1])
		Else
			cArm := CB8->CB8_LOCAL
		EndIf
		nP1 := Ascan(aSep,{|x| x[1]+X[2]+x[3]+x[4] == cPedido+SC6->C6_CLI+SC6->C6_LOJA+cArm})
		If	nP1 == 0
			aadd(aSep,{cPedido,SC6->C6_CLI,SC6->C6_LOJA,cArm,{}})
			nP1 := Len(aSep)
		EndIf
		CB8->(aadd(aSep[nP1,5],{CB8_ITEM,CB8_SEQUEN,CB8_PROD,CB8_PEDIDO,CB8_LOCAL,CB8_SLDPRE,CB8_SALDOS,CB8_LCALIZ,CB8_LOTECT,CB8_NUMLOT,CB8_NUMSER,CB8_CFLOTE,Recno()}))
	Elseif cOrigem == "3" // Ordem de Producao
		SC2->(DbSetOrder(1))
		SC2->(DbSeek(xFilial("SC2")+CB8->CB8_OP))
		If	"08*" $ cTipExp  // gera Ordem de Separacao com todos os Armazens
			cArm :=Space(Tamsx3("B1_LOCPAD")[1])
		Else
			cArm := CB8->CB8_LOCAL
		EndIf
		nP1 := Ascan(aSep,{|x| x[1]+X[2] == CB8->CB8_OP+cArm})
		If	nP1 == 0
			aadd(aSep,{CB8->CB8_OP,cArm,{}})
			nP1 := Len(aSep)
		EndIf
		CB8->(aadd(aSep[nP1,3],{CB8_ITEM,CB8_SEQUEN,CB8_PROD,CB8_OP,CB8_LOCAL,CB8_SLDPRE,CB8_SALDOS,CB8_LCALIZ,CB8_LOTECT,CB8_NUMLOT,CB8_NUMSER,CB8_CFLOTE,Recno()}))
	Endif
	CB8->(DbSkip())
EndDo
If	Empty(aSep)
	Reclock('CB7')
	CB7->CB7_STATPA := "1"  // Pausa
	CB7->(MsUnLock())
	Return
EndIf

cTipExp := StrTran(cTipExp,"09*","")
If	!VTYesNo(STR0067+cMsgParcial+"?",STR0002,.T.) //"Gera as Ordens de Separacoes"###"Aviso"
	Reclock('CB7')
	CB7->CB7_STATPA := "1"  // Pausa
	CB7->(MsUnLock())
	Return
EndIf
VTMsg(STR0068) //"Aguarde..."

If	cOrigem == "1" // Pedido de Venda
	GeraOSPV(aSep,cMsgParcial)
ElseIf cOrigem == "3" // Ordem de Producao
	GeraOSOP(aSep,cMsgParcial)
EndIf
Return

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡ao    ³ GeraOSPV     ³ Autor ³ ACD                   ³ Data ³ 14/08/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡ao ³ Analisa as Ordens de separacao a ser geradas por Pedido        ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function GeraOSPV(aSep,cMsgParcial)
Local nX := 0
Local nY := 0
Local cNewOrdSep
Local cOrdSepAtu := CB7->CB7_ORDSEP
Local lACDV165F  := ExistBlock("ACDV165F")
Local nTamOper   := TamSX3("CB7_CODOPE")[1]

DbSelectArea("SC9")
DbSetOrder(1)

For nX:= 1 to Len(aSep)
	cNewOrdSep:= GetSX8Num( "CB7", "CB7_ORDSEP" )
	RecLock( "CB7", .T. )
	CB7->CB7_FILIAL := xFilial("CB7")
	CB7->CB7_ORDSEP := cNewOrdSep
	CB7->CB7_PEDIDO := aSep[nX,1]
	CB7->CB7_CLIENT := aSep[nX,2]
	CB7->CB7_LOJA   := aSep[nX,3]
	CB7->CB7_COND   := cCond
	CB7->CB7_LOJENT := cLojaEnt
	CB7->CB7_LOCAL  := aSep[nX,4]
	CB7->CB7_DTEMIS := dDataBase
	CB7->CB7_HREMIS := Time()
	CB7->CB7_STATUS := "0"
	CB7->CB7_CODOPE := Space(nTamOper)
	CB7->CB7_PRIORI := "1"
	CB7->CB7_ORIGEM := "1"
	CB7->CB7_TIPEXP := cTipExp
	CB7->CB7_TRANSP := cTransp
	CB7->CB7_AGREG  := cAgreg
	CB7->CB7_PRESEP := cOrdSep
	CB7->(MsUnLock())
	ConfirmSX8()
	For nY:= 1 to len(aSep[nX,5])
		RecLock( "CB8", .T. )
		CB8->CB8_FILIAL := xFilial("CB8")
		CB8->CB8_ORDSEP := cNewOrdSep
		CB8->CB8_ITEM   := aSep[nX,5,nY,1]
		CB8->CB8_PEDIDO := aSep[nX,5,nY,4]
		CB8->CB8_PROD   := aSep[nX,5,nY,3]
		CB8->CB8_LOCAL  := aSep[nX,5,nY,5]
		CB8->CB8_QTDORI := aSep[nX,5,nY,6]-aSep[nX,5,nY,7]
		CB8->CB8_SALDOS := aSep[nX,5,nY,6]-aSep[nX,5,nY,7]
		CB8->CB8_SALDOE := aSep[nX,5,nY,6]-aSep[nX,5,nY,7]
		CB8->CB8_LCALIZ := aSep[nX,5,nY,8]
		CB8->CB8_SEQUEN := aSep[nX,5,nY,2]
		CB8->CB8_LOTECT := aSep[nX,5,nY,9]
		CB8->CB8_NUMLOT := aSep[nX,5,nY,10]
		CB8->CB8_NUMSER := aSep[nX,5,nY,11]
		CB8->CB8_CFLOTE := If("10*" $ cTipExp,"1","2")
		CB8->(MsUnLock())
		CB8->(DbGoto(aSep[nX,5,nY,13]))
		RecLock("CB8")
		CB8->CB8_SLDPRE := aSep[nX,5,nY,7]
		CB8->CB8_SALDOS := aSep[nX,5,nY,7]
		CB8->(MsUnLock())
		
		//Atualiza SC9990
		SC9->(DbSeek(xFilial("SC9")+CB8->(CB8_PEDIDO+CB8_ITEM)))
		While SC9->(!Eof() .and. C9_FILIAL+C9_PEDIDO+C9_ITEM==xFilial("SC9")+CB8->(CB8_PEDIDO+CB8_ITEM))
			If ! Empty(SC9->C9_NFISCAL) .or. (SC9->C9_ORDSEP # cOrdSepAtu) .or. (CB8->CB8_PROD#SC9->C9_PRODUTO)
				SC9->(DbSkip())
				Loop
			EndIf
			RecLock("SC9",.F.)
			SC9->C9_ORDSEP := cNewOrdSep
			SC9->( MsUnLock() )
			SC9->(DbSkip())
		EndDo
	Next
	// --- Ponto de entrada apos o encerramento da geracao da Ordem de Separacao
	If lACDV165F
		lRet := ExecBlock("ACDV165F",.F.,.F.,{cNewOrdSep})
	EndIf
Next

CB7->(DbSetOrder(1))
CB7->(DbSeek(xFilial("CB7")+cOrdSep))
RecLock("CB7",.f.)
If Empty(cMsgParcial)
	CB7->CB7_STATUS := "9"
	CB7->CB7_DTFIMS := dDataBase
	CB7->CB7_HRFIMS := StrTran(Time(),":","")
Else
	CB7->CB7_STATPA := "1"
EndIf
CB7->(MsUnLock())
Return

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡ao    ³ GeraOSOP     ³ Autor ³ ACD                   ³ Data ³ 14/08/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡ao ³ Analisa as Ordens de separacao a ser geradas por OP's          ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function GeraOSOP(aSep,cMsgParcial)
Local nX := 0
lOCAL nY := 0
Local cNewOrdSep
Local lACDV165F  := ExistBlock("ACDV165F")
Local nTamOper   := TamSX3("CB7_CODOPE")[1]

For nX:= 1 to Len(aSep)
	cNewOrdSep:= GetSX8Num( "CB7", "CB7_ORDSEP" )
	RecLock("CB7", .T.)
	CB7->CB7_FILIAL := xFilial("CB7")
	CB7->CB7_ORDSEP := cNewOrdSep
	CB7->CB7_OP     := aSep[nX,1]
	CB7->CB7_LOCAL  := aSep[nX,2]
	CB7->CB7_DTEMIS := dDataBase
	CB7->CB7_HREMIS := Time()
	CB7->CB7_STATUS := "0"
	CB7->CB7_CODOPE := Space(nTamOper)
	CB7->CB7_PRIORI := "1"
	CB7->CB7_ORIGEM := "3"
	CB7->CB7_TIPEXP := cTipExp
	CB7->CB7_AGREG  := cAgreg
	CB7->CB7_PRESEP := cOrdSep
	CB7->(MsUnLock())
	ConfirmSX8()
	For nY:= 1 to len(aSep[nX,3])
		RecLock( "CB8", .T. )
		CB8->CB8_FILIAL := xFilial("CB8")
		CB8->CB8_ORDSEP := cNewOrdSep
		CB8->CB8_ITEM   := aSep[nX,3,nY,1]
		CB8->CB8_OP     := aSep[nX,3,nY,4]
		CB8->CB8_PROD   := aSep[nX,3,nY,3]
		CB8->CB8_LOCAL  := aSep[nX,3,nY,5]
		CB8->CB8_QTDORI := aSep[nX,3,nY,6]-aSep[nX,3,nY,7]
		CB8->CB8_SALDOS := aSep[nX,3,nY,6]-aSep[nX,3,nY,7]
		CB8->CB8_SALDOE := aSep[nX,3,nY,6]-aSep[nX,3,nY,7]
		CB8->CB8_LCALIZ := aSep[nX,3,nY,8]
		CB8->CB8_SEQUEN := aSep[nX,3,nY,2]
		CB8->CB8_LOTECT := aSep[nX,3,nY,9]
		CB8->CB8_NUMLOT := aSep[nX,3,nY,10]
		CB8->CB8_NUMSER := aSep[nX,3,nY,11]
		CB8->CB8_CFLOTE := If("10*" $ cTipExp,"1","2")
		CB8->(MsUnLock())
		CB8->(DbGoto(aSep[nX,3,nY,13]))
		RecLock("CB8")
		CB8->CB8_SLDPRE := aSep[nX,3,nY,7]
		CB8->CB8_SALDOS := aSep[nX,3,nY,7]
		CB8->(MsUnLock())
	Next
	// --- Ponto de entrada apos o encerramento da geracao da Ordem de Separacao
	If lACDV165F
		lRet := ExecBlock("ACDV165F",.F.,.F.,{cNewOrdSep})
	EndIf
Next

CB7->(DbSetOrder(1))
CB7->(DbSeek(xFilial("CB7")+cOrdSep))
RecLock("CB7",.f.)
If Empty(cMsgParcial)
	CB7->CB7_STATUS := "9"
	CB7->CB7_DTFIMS := dDataBase
	CB7->CB7_HRFIMS := StrTran(Time(),":","")
Else
	CB7->CB7_STATPA := "1"
EndIf
CB7->(MsUnLock())
Return

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡ao    ³  Estorna     ³ Autor ³ ACD                   ³ Data ³ 14/08/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡ao ³  Estorno da Separacao                                          ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function Estorna()
Local aTela
Local cEtiqEnd := Space(20)
Local cProduto := IIf( FindFunction( 'AcdGTamETQ' ), AcdGTamETQ(), Space(48) )
Local cVolume  := Space(10)
Local nQtde    := 1 

Private cArmazem := Space(TamSX3("B2_LOCAL")[1])
Private cEndereco:= Space(TamSX3("BF_LOCALIZ")[1])
Private cLoteNew := Space(TamSX3("B8_LOTECTL")[1])
Private cSLoteNew:= Space(TamSX3("B8_NUMLOTE")[1])
IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf

aTela := VTSave()
If lVT100B // GetMv("MV_RF4X20")
	While .T.
		VTClear()
		@ 0,0 VtSay STR0069 //"Estorno da leitura"
		@ 1,0 VTSay STR0020 //'Leia o endereco'
		If UsaCB0('02')
			@ 2,0 VTGet cEtiqEnd pict '@!' valid VldEndEst(@cArmazem,@cEndereco,cEtiqEnd) .and. iif(lVolta,(lVolta := .F., .T.), .T.)
		Else
			@ 2,0 VTGet cArmazem pict '@!' valid ! Empty(cArmazem) when iif(lVolta,(VTKeyBoard(chr(13)),lVolta := .F.,.T.),.T.)
			@ 2,3 VTSay "-" VTGet cEndereco pict '@!' valid VtLastKey()==5 .or. VldEndEst(@cArmazem,@cEndereco)
		EndIf
		VTRead
		
		If !(vtLastKey() == 27) //Segunda Tela
			VTClear
			cProduto   := IIf( FindFunction( 'AcdGTamETQ' ), AcdGTamETQ(), Space(48) )
			If ! UsaCB0('01')
				@ 0,0 VTSay STR0036 VtGet nQtde pict cPictQtde valid nQtde > 0 when VtLastkey()==5 .and. iif(vtRow() == 0 .and. vtLastKey() == 5,(VTKeyBoard(chr(27)),lVolta := .T.),.T.) //'Qtde '
			EndIf
			@ 1,0 VTSay STR0037 //'Leia o produto'
			@ 2,0 VTGet cProduto pict '@!' VALID VTLastkey() == 5 .or. VldProdEst(cProduto,@nQtde,cArmazem,cEndereco,cVolume) when iif(vtRow() == 0 .and. vtLastKey() == 5 .and. UsaCB0('01'),(VTKeyBoard(chr(27)),lVolta := .T.),.T.)
			VTRead
		EndIf
		
		if lVolta
			Loop
		endif
	
		Exit
	EndDo
Else
	VTClear()
	@ 0,0 VtSay STR0069 //"Estorno da leitura"
	@ 1,0 VTSay STR0020 //'Leia o endereco'
	If UsaCB0('02')
		@ 2,0 VTGet cEtiqEnd pict '@!' valid VldEndEst(@cArmazem,@cEndereco,cEtiqEnd)
	Else
		@ 2,0 VTGet cArmazem pict '@!' valid ! Empty(cArmazem)
		@ 2,3 VTSay "-" VTGet cEndereco pict '@!' valid VtLastKey()==5 .or. VldEndEst(@cArmazem,@cEndereco)
	EndIf
	cProduto   := IIf( FindFunction( 'AcdGTamETQ' ), AcdGTamETQ(), Space(48) )
	If ! UsaCB0('01')
		@ 5,0 VTSay STR0036 VtGet nQtde pict cPictQtde valid nQtde > 0 when VtLastkey()==5  //'Qtde '
	EndIf
	@ 6,0 VTSay STR0037 //'Leia o produto'
	@ 7,0 VTGet cProduto pict '@!' VALID VTLastkey() == 5 .or. VldProdEst(cProduto,@nQtde,cArmazem,cEndereco,cVolume)
	VTRead
EndIf
VTRestore(,,,,aTela)
Return

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡ao    ³VldEndEst ³ Autor ³ ACD                   ³ Data ³ 14/08/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡ao ³Validacao do Endereco na rotina de estorno                  ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function VldEndEst(cArmazem,cEndereco,cEtiqEnd)
Local aRet
Local lAchou
VtClearBuffer()
If Empty(cArmazem) // se Empty(cArmazem) entao e' etiqueta de endereco com CB0
	aRet := CBRetEti(cEtiqEnd,'02')
	If len(aRet) == 0
		VtAlert(STR0021,STR0002,.t.,4000,3)  //"Etiqueta invalida"###"Aviso"
		VtKeyboard(Chr(20))  // zera o get
		Return .F.
	EndIf
	cArmazem  := aRet[2]
	cEndereco := aRet[1]
EndIf
CB9->(DbSetOrder(1))
If ! CB9->(DbSeek(xFilial("CB9")+CB7->CB7_ORDSEP))
	VtAlert(STR0070,STR0002,.t.,4000,3)  //"Nao existe itens separados para esta Ordem de Pre-Separacao"###"Aviso"
	VTClearGet("cArmazem")
	VTClearGet("cEndereco")
	VTGetSetFocus("cArmazem")
	VtKeyboard(Chr(20))  // zera o get
	Return
EndIf
lAchou := .f.
While CB9->(! Eof() .and. xFilial("CB9")+CB7->CB7_ORDSEP==CB9_FILIAL+CB9_ORDSEP)
	If cArmazem+cEndereco == CB9->(CB9_LOCAL+CB9_LCALIZ)
		lAchou:= .t.
		Exit
	EndIf
	CB9->(DBSkip())
EndDo
If ! lAchou
	VtAlert(STR0022,STR0002,.t.,4000,3)  //"Endereco incorreto"###"Aviso"
	cArmazem  := Space(TamSX3("B2_LOCAL")[1])
	cEndereco := Space(TamSX3("BF_LOCALIZ")[1])
	VTClearGet("cArmazem")
	VTClearGet("cEndereco")
	VTGetSetFocus("cArmazem")
	VtKeyboard(Chr(20))  // zera o get
	Return .F.
EndIf
Return .T.

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡ao    ³VldProdEst³ Autor ³ ACD                   ³ Data ³ 14/08/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡ao ³Validacao do produto na rotina de estorno                   ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
Static function VldProdEst(cEProduto,nQtde,cArmazem,cEndereco,cVolume)
Local cTipo       := ""
Local cProduto 	  := ""
Local cSeek  	  := ""
Local aEtiqueta   := {}
Local aRet		  := {} 
Local aItensE 	  := {}
Local aItensE2 	  := {}
Local cLote  	  := Space(TamSX3("B8_LOTECTL")[1])
Local cSLote 	  := Space(TamSX3("B8_NUMLOTE")[1])
Local cNumSer	  := Space(TamSX3("BF_NUMSERI")[1])
Local nQtdItens   := 0
Local nQE  		  := 0
Local nPos 	  	  := 0
Local nTQtde   	  := 0
Local nQtdSaldo	  := 0
Local nSaldoCB9	  := 0
Local nQeDisp 	  := 0
Local nY 	  	  := 0
Local nW 		  := 0 
Local nX 		  := 0

Private nQtdLida  :=0

If Empty(cEProduto)
	Return .f.
EndIf

	
//- Verifica se a etiqueta pertence a um PALLET
aItensPallet := CBItPallet(cEProduto)
	
If Len(aItensPallet) == 0
	aItensPallet:={cEProduto}
	lIsPallet := .F.
EndIf

For nX :=1 To Len(aItensPallet)
	cEProduto :=aItensPallet[nX]
	If !CBLoad128(@cEProduto)
		VtKeyboard(Chr(20))  // zera o get
		Return .F.
	EndIf
	cTipo := CBRetTipo(cEProduto)
	If cTipo == "01"
		aEtiqueta:= CBRetEti(cEProduto,"01")
		If Empty(aEtiqueta)
			VtAlert(STR0021,STR0002,.t.,4000,3)  //"Etiqueta invalida"###"Aviso"
			VtKeyboard(Chr(20))  // zera o get
			Return .F.
		EndIf
		CB9->(DbSetorder(1))
		If ! CB9->(DbSeek(xFilial("CB9")+cOrdSep+Left(cEProduto,10)))
			VtAlert(STR0071,STR0002,.t.,4000,3)  //"Produto nao separado"###"Aviso"
			VtKeyboard(Chr(20))  // zera o get
			Return .F.
		EndIf
		cProduto := aEtiqueta[1]
		nQE      := aEtiqueta[2]
		cLote    := aEtiqueta[16]
		cSLote   := aEtiqueta[17]
		nQtdLida := nQE
		If cArmazem+cEndereco <> CB0->(CB0_LOCAL+CB0_LOCALI)
			VtAlert(STR0072+CB0->CB0_LOCAL+"-"+CB0->CB0_LOCALI,STR0002,.t.,5000,3)  //"Produto nao pertence a este endereco, o correto eh, "###"Aviso"
			VtKeyboard(Chr(20))  // zera o get
			Return .F.
		EndIf
		AADD(aLoteNew,cLote)
		AADD(aSLoteNew,cSLote)
	ElseIf cTipo $ "EAN8OU13-EAN14-EAN128"
		aRet := CBRetEtiEan(cEProduto)
		If len(aRet) == 0
			VtAlert(STR0021,STR0002,.t.,4000,3)  //"Etiqueta invalida"###"Aviso"
			VtKeyboard(Chr(20))  // zera o get
			Return .F.
		EndIf
		cProduto := aRet[1]
		nQE  :=CBQtdEmb(aRet[1])*nQtde
		If Empty(nQE)
			VtAlert(STR0039,STR0002,.t.,4000,3)  //"Quantidade invalida"###"Aviso"
			VtKeyboard(Chr(20))  // zera o get
			Return .F.
		EndIf
		cLote := aRet[3]
		If ! CBRastro(aRet[1],@cLote,@cSLote)
			VTKeyBoard(chr(20))
			Return .f.
		EndIf
		cNumSer:= aRet[5]	 // Numero de Serie
		If Empty(cNumSer) .and. CBSeekNumSer(cOrdSep,cProduto,nQtde)
			If ! VldQtde(nQtde,.T.)
				VtKeyboard(Chr(20))  // zera o get
				Return .F.
			EndIf
			If ! CBNumSer(@cNumSer)
				VTKeyBoard(chr(20))
				Return .f.
			EndIf
		Endif
	
		CB9->(DBSetOrder(9))
		If ! CB9->(DbSeek(xFilial("CB9")+cOrdSep+cProduto+cArmazem+cEndereco+cLote+cSLote+cNumSer))
			VtAlert(STR0073,STR0002,.t.,4000,3)  //"Item nao encontrado"###"Aviso"
			VtKeyboard(Chr(20))  // zera o get
			Return .F.
		EndIf
		nQeDisp:= CB9->(RetQeDisp(CB9_PROD,CB9_LOCAL,CB9_LCALIZ,CB9_LOTECT,CB9_NUMLOT))
		If nQE > nQeDisp
			VtAlert(STR0074,STR0002,.t.,4000,3)  //"Quantidade informada maior do que separada"###"Aviso"
			VtKeyboard(Chr(20))  // zera o get
			Return .F.
		EndIf
	Else
		VtAlert(STR0021,STR0002,.t.,4000,3)  //"Etiqueta invalida"###"Aviso"
		VtKeyboard(Chr(20))  // zera o get
		Return .F.
	EndIf
Next nX


IF ! VtYesNo(STR0075,STR0002,.t.)  //"Confirma o estorno?"###"Aviso"
	VtKeyboard(Chr(20))  // zera o get
	Return .F.
EndIf
nX:=0

For nX :=1 To Len(aItensPallet)	
	nQtdLida := nQE
	
	// montando array
	CB8->(DbSetOrder(3))
	CB8->(DbSeek(xFilial("CB8")+cOrdSep))
	While ! CB8->(Eof()) .and. CB8->CB8_FILIAL == xFilial("CB8") .and. CB8->CB8_ORDSEP == cOrdSep
		If CB8->(CB8_LOCAL+CB8_LCALIZ) # cArmazem+cEndereco
			CB8->(DbSkip())
			Loop
		EndIf
		If CB8->CB8_PROD # cProduto
			CB8->(DbSkip())
			Loop
		EndIf
		If CB8->CB8_QTDORI == CB8->CB8_SALDOS // Nao tem quantidade para estornar neste registro
			CB8->(DbSkip())
			Loop
		EndIf
		nPos := Ascan(aItensE,{|x| x[2]+x[3]+x[4]+x[5]+x[8]+x[9]+x[10]+x[13] == CB8->(CB8_ITEM+CB8_PROD+CB8_LOCAL+CB8->CB8_LCALIZ+CB8_LOTECT+CB8_NUMLOT+CB8_NUMSER+CB8_CFLOTE)})
		If  nPos == 0
			aadd(aItensE,{{CB8->(Recno())},;
			CB8->CB8_ITEM,;
			CB8->CB8_PROD,;
			CB8->CB8_LOCAL,;
			CB8->CB8_LCALIZ,;
			CB8->CB8_SALDOS,;
			CB8->CB8_QTDORI,;
			CB8->CB8_LOTECT,;
			CB8->CB8_NUMLOT,;
			CB8->CB8_NUMSER,;
			NIL,;
			NIL,;
			CB8->CB8_CFLOTE })
		Else
			If Ascan(aItensE[nPos,1],{|x| x == CB8->(Recno())}) == 0
				aadd(aItensE[nPos,1],CB8->(Recno()))
			EndIf
			aItensE[nPos,6]+= CB8->CB8_SALDOS
			aItensE[nPos,7]+= CB8->CB8_QTDORI
		EndIf
		CB8->(DbSkip())
	EndDo
	
	aItensE := aSort(aItensE,,,{|x,y| x[4]+x[5]+x[8]+x[9]+x[10] < y[4]+y[5]+y[8]+y[9]+y[10]})
	
	cIdCB0  := CB9->CB9_CODETI
	
	AADD(aLoteNew,cLote)
	AADD(aSLoteNew,cSLote)
	
	aItensE2 := aClone(aItensE)
	nQtdItens:= Len(aItensE2)
	
	For nY:= 1 to nQtdItens
		nPos:= ascan(aItensE2,{|x| X[_POSARMAZEM]+X[_POSENDERECO]+X[_POSLOTECTL]+X[_POSNUMLOTE]+X[_POSNUMSERIE] == ;
		CB9->(CB9_LOCAL+CB9_LCALIZ+CB9_LOTSUG+CB9_SLOTSUG+CB9_NUMSER)})
		If nPos == 0
			Loop
		EndIf
		For nW:= 1 to Len(aItensE2[nPos,_POSRECNO])
			CB8->(DbGoto(aItensE2[nPos,_POSRECNO,nW]))
			If UsaCB0("01")
				cSeek:= CB8->CB8_ORDSEP+cIdCB0
				CB9->(DbSetOrder(1))
				CB9->(DbSeek(xFilial("CB9")+cSeek))
			Else
				cSeek:= CB8->(CB8_ORDSEP+CB8_ITEM+CB8_PROD+CB8_LOCAL+CB8_LCALIZ+cLoteNew+cSLoteNew+CB8_NUMSER+CB8_LOTECT+CB8_NUMLOT)+Space(10)+cIdCB0
				CB9->(DbSetOrder(10))
				CB9->(DbSeek(xFilial("CB9")+cSeek))
			EndIf
			If CB9->(Eof())
				Loop
			EndIf
			nSaldoCB9:= CB9->CB9_QTESEP
			RecLock("CB9",.F.)
			CB9->CB9_QTESEP -= If(nQtdLida > nSaldoCB9,nSaldoCB9,nQtdLida)
			If Empty(CB9->CB9_QTESEP)
				CB9->(DbDelete())
			EndIf
			CB9->(MsUnlock())
			RecLock("CB8")
			CB8->CB8_SALDOS +=  If(nQtdLida > nSaldoCB9,nSaldoCB9,nQtdLida)
			CB8->(MsUnlock())
			aItensE2[nPos,_POSSLDSEP] += If(nQtdLida > nSaldoCB9,nSaldoCB9,nQtdLida)
			nQtdLida -= If(nQtdLida > nSaldoCB9,nSaldoCB9,nQtdLida)
			If nQtdLida == 0
				Exit
			Endif
		Next
		If nQtdLida == 0
			Exit
		EndIf
		If aItensE2[nPos,6] == aItensE2[nPos,7] // Nao ha mais quantidade para ser estornada deste item
			aDel(aItensE2,nPos)
			aSize(aItensE2,Len(aItensE2)-1)
		EndIf
	Next
Next nX
	
RecLock("CB7",.F.)
CB7->CB7_STATUS := '1'
CB7->(MsUnlock())
nQtde:= 1
VTGetRefresh('nQtde')
VtKeyboard(Chr(20))  // zera o get
Return .F.

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡ao    ³RetQeDisp ³ Autor ³ ACD                   ³ Data ³ 19/01/04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡ao ³Retorna a quantidade disponivel do produto para o estorno   ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
Static Function RetQeDisp(cProduto,cLocal,cEndereco,cLote,cSubLote)
Local nQuant  := 0
Local aAreaCB9:= CB9->(GetArea())
Local nRecCB9 := CB9->(Recno())

CB9->(DbSetOrder(9))
CB9->(DbSeek(xFilial("CB9")+cOrdSep+cProduto))
While CB9->(! Eof() .and. CB9_FILIAL+CB9_ORDSEP+CB9_PROD+CB9_LOCAL+CB9_LCALIZ+CB9_LOTECT+CB9_NUMLOT == xFilial("CB9")+cOrdSep+cProduto+cLocal+cEndereco+cLote+cSubLote)
	nQuant+= CB9->CB9_QTESEP
	CB9->(DbSkip())
EndDo

RestArea(aAreaCB9)
CB9->(DbGoTo(nRecCB9))
Return(nQuant)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o	 ³ GeraNewEti ³ Autor ³ Anderson Rodrigues  ³ Data ³ 15/11/02 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³ Gera Nova etiqueta para produtos de quantidade variavel    ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso		 ³ SigaACD                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
Static Function GeraNewEti(nQtd)
Local cEtiqueta:= CBProxCod("MV_CODCB0")
Local nRecno
Local nRecnoCB0 := CB0->(Recno())
Local aNewFields:={}

If ! CB5SetImp(CBRLocImp("MV_IACD01"),IsTelNet())
	VTBeep(3)
	VTAlert(STR0076,STR0077,.t.,3000)  //'Local de impressao nao configurado, MV_IACD01'###'Aviso'
	Return
EndIf

VTMsg(STR0080)  //"Imprimindo..."

CB0->(DbGoto(nRecnoCB0))                         
aNewFields :={	{"CB0_CODETI"	,cEtiqueta	},;
				{"CB0_QTDE"		,nQtd		},;
				{"CB0_ORIGEM"	,"CB7"		}}

nRecno:= CB0->(CBCopyRec(aNewFields))
CB0->(DbGoto(nRecno))

SB1->(DbSetOrder(1))
SB1->(DbSeek(xFilial("SB1")+CB0->CB0_CODPRO))
If ExistBlock('IMG01')
	ExecBlock("IMG01",,,{,,CB0->CB0_CODETI})
EndIf
If ExistBlock('IMG00')
	ExecBlock("IMG00",,,{"ACDV170",cOrdSep})
EndIf
MSCBCLOSEPRINTER()
CB0->(DbGoto(nRecnoCB0))
Reclock("CB0",.f.)
CB0->CB0_QTDE := CB0->CB0_QTDE - nQtd
CB0->(MSUnlock())
Return(cEtiqueta)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡ao    ³  VldQtde ³ Autor ³ Anderson Rodrigues    ³ Data ³ 29/06/04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡ao ³ Validacao da quantidade informada                          ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³SIGAACD                                                     ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
Static Function VldQtde(nQtde,lSerie)
Local lRet := .T.
Default lSerie:= .f.

If nQtde <= 0
	lRet := .f.
ElseIf lSerie .and. nQtde > 1
	VTAlert(STR0078,STR0002,.T.,2000) //"Quantidade invalida !!!"###"Aviso"
	VTAlert(STR0079,STR0002,.T.,4000) //"Quando se utiliza numero de serie a quantidade deve ser == 1"###"Aviso"
	lRet := .f.
EndIf
Return lRet

Static Function MSCBFSem()
Local nC:= 0
__nSem := -1
While __nSem  < 0
	__nSem  := MSFCreate("V165"+cCodOpe+".sem")
	IF  __nSem  < 0
		SLeep(50)
		nC++
		If nC == 3
			Return .F.
		EndIf
	EndIf
EndDo
FWrite(__nSem,STR0081+cCodOpe+STR0082+cOrdSep) //"Operador: "###" Ordem de Separacao: "
Return .T.

Static Function MSCBASem()
If __nSem > 0
	Fclose(__nSem)
	FErase("V165"+cCodOpe+STR0085)//".sem"
EndIf
Return 10
