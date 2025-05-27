#INCLUDE "Acdv150.ch" 
#INCLUDE "protheus.ch"
#INCLUDE "APVT100.CH"

Template function acdv150A(uPar1,uPar2,uPar3,uPar4,uPar5,uPar6,uPar7,uPar8,uPar9,uPar10)
Return acdv150A(uPar1,uPar2,uPar3,uPar4,uPar5,uPar6,uPar7,uPar8,uPar9,uPar10)
Template function acdv150B(uPar1,uPar2,uPar3,uPar4,uPar5,uPar6,uPar7,uPar8,uPar9,uPar10)
Return acdv150B(uPar1,uPar2,uPar3,uPar4,uPar5,uPar6,uPar7,uPar8,uPar9,uPar10)

Function acdv150A()  //somente para aramzens sem controle de CQ
Return acdv150(.f.)
Function acdv150B() // somente para armazen com controle de CQ
Return acdv150(.t.)

/*
‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±⁄ƒƒƒƒƒƒƒƒƒƒ¬ƒƒƒƒƒƒƒƒƒƒƒƒ¬ƒƒƒƒƒƒƒ¬ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒ¬ƒƒƒƒƒƒ¬ƒƒƒƒƒƒƒƒƒƒø±±
±±≥FunáÑo	 ≥ acdv150    ≥ Autor ≥ Sandro              ≥ Data ≥ 14/08/01 ≥±±
±±√ƒƒƒƒƒƒƒƒƒƒ≈ƒƒƒƒƒƒƒƒƒƒƒƒ¡ƒƒƒƒƒƒƒ¡ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒ¡ƒƒƒƒƒƒ¡ƒƒƒƒƒƒƒƒƒƒ¥±±
±±≥DescriáÑo ≥ Programa de transferencia (com Controle de Localizacao)    ≥±±
±±√ƒƒƒƒƒƒƒƒƒƒ≈ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒ¥±±
±±≥Parametros≥ExpL1: Indica se fara transf. de CQ. Default=.f.            ≥±±
±±√ƒƒƒƒƒƒƒƒƒƒ≈ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒ¥±±
±±≥ Uso      ≥ SigaACD                                                    ≥±±
±±¿ƒƒƒƒƒƒƒƒƒƒ¡ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒŸ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂ
*/   
Template function acdv150(uPar1,uPar2,uPar3,uPar4,uPar5,uPar6,uPar7,uPar8,uPar9,uPar10)
Return acdv150(uPar1,uPar2,uPar3,uPar4,uPar5,uPar6,uPar7,uPar8,uPar9,uPar10)

Function acdv150(lEndCQ)
Local aKey          := array(2)

Private nTamLote    := TamSX3("B8_LOTECTL")[1]
Private nTamSLote   := TamSX3("B8_NUMLOTE")[1]
Private cArmOri     := Space(Tamsx3("B1_LOCPAD") [1])
Private cEndOri     := Space(TamSX3("BF_LOCALIZ")[1])
Private cCB0EndOri  := Space(20)

Private cArmDes     := Space(Tamsx3("B1_LOCPAD") [1])
Private cEndDes     := Space(TamSX3("BF_LOCALIZ")[1])
Private cCB0EndDes  := Space(20)

Private cCB0ProdEnd := IIf( FindFunction( 'AcdGTamETQ' ), AcdGTamETQ(), Space(48) )
Private cProduto    := IIf( FindFunction( 'AcdGTamETQ' ), AcdGTamETQ(), Space(48) )

Private nQtde       := 1
Private nQtdeProd   := 1
Private cLote       := Space(nTamLote)
Private cSLote      := Space(nTamSLote)
Private cLoteDes    := Space(nTamLote)
Private cSLoteDes   := Space(nTamSLote)
Private dValLtDes	   := CTOD('')
Private cNumSerie   := Space(TamSX3("BF_NUMSERI")[1])
Private aLista      := {}
Private aHisEti     := {}
Private lMsErroAuto := .F.
Private nLin        := 0
Private lCQ         := lEndCQ
Private lForcaQtd   :=GetMV("MV_CBFCQTD",,"2") =="1"
Private lTrfLote 	  := SuperGetMV("MV_ACDTRLT",.F.,.F.) 
Private lVolta := .F.
IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf

aKey[1] := VTSetKey(24,{|| Estorna()},STR0045) //"Estorno"
aKey[2] := VTSetKey(09,{|| Informa()},STR0046) //"Informacoes"

While .t.
	VTClear
	nLin:= -1
	@ ++nLin,0 VTSAY STR0001 //"Transferencia"
	If !lVolta
		If ! GetEndOri()
			Exit
		EndIf
	EndIf
	GetProduto()
	If lVT100B /* GetMv("MV_RF4X20") */ .and. VTLastKey() == 27
		Exit
	EndIf
	GetEndDes()
	VTRead
	if lVolta
		Loop
	endif
	If vtLastKey() == 27
		If len(aLista) > 0 .and. ! VTYesNo(STR0002,STR0003) //'Confirma a saida?'###'Atencao'
			loop
		EndIf
		Exit
	EndIf
End
vtsetkey(24,aKey[1])
vtsetkey(09,aKey[2])
Return

//‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹
Static Function GetEndOri()
Local cPictNNR := PesqPict( "NNR","NNR_CODIGO")
Local lRet := .T.
If ! UsaCB0('01')
	@ ++nLin,0 VtSay STR0004 //'Endereco origem'
	If UsaCB0('02')
		@ ++nLin,0 VTGet cCB0EndOri pict '@!' Valid VldEndOri() when Empty(cCB0EndOri)
	Else
		@ ++nLin,0 VTGet cArmOri pict cPictNNR Valid ! Empty(cArmOri)  when Empty(cArmOri)
		If Len(cArmOri) > 2
			nLin++
		EndIf
		@ nLin,2 VTSay '-' VtGet cEndOri Pict '@!' Valid  VldEndOri() when Empty(cEndOri)
	EndIf
	VTRead()
	If VTLastkey() == 27
		lRet := .F.
	EndIf
	If lRet
		VTClear(1,0,2,19)
		nLin := 0
	EndIf
EndIf
Return lRet

//‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹
Static Function VldEndOri()
Local lRet := .T.
Local aEnd := {}
//⁄ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒø	
//≥ Ponto de Entrada utilizado para manipular o codigo da etiqueta lida ≥
//¿ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒŸ
If ExistBlock("ACD150ME")
	If UsaCB0("02")
		aEnd := ExecBlock("ACD150ME", .F., .F., {1,cCB0EndOri} )	
		If ValType(aEnd) == "A"
			cCB0EndOri := aEnd[1]
		EndIf
	Else
		aEnd := ExecBlock("ACD150ME", .F., .F., {1,cArmOri,cEndOri} )
		If ValType(aEnd) == "A"
			cArmOri := aEnd[1]
			cEndOri := aEnd[2]
		EndIf
	EndIf
	
	aEnd := {}
EndIf
If Empty( cCB0EndOri+cArmOri+cEndOri)
	VtAlert(STR0005, STR0013,.t.,4000,4) //'Endereco invalido'###'Aviso'
	VTKeyboard(chr(20))
	If ! UsaCb0("02")
		VTClearGet("cArmOri")
		VTGetSetFocus("cArmOri")
	EndIf
	lRet := .F.
EndIF
If lRet .And. UsaCB0('02')
	aEnd:= CbRetEti(cCB0EndOri,"02")
	If Empty(aEnd)
		VtAlert(STR0005, STR0013,.t.,4000,4) //'Endereco invalido'###'Aviso'
		VTKeyboard(chr(20))
		lRet := .F.
	EndIf
	cCB0EndOri:= aEnd[1]
	cArmOri   := aEnd[2]
	cEndOri   := aEnd[1]
	VTGetRefresh('cCB0EndOri')
EndIF

If lRet
	lRet := ValidEnd(cArmOri,cEndOri,1)
EndIf
Return lRet

//‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹
Static Function ValidEnd(cArmazem,cEndereco,nTipo)
Local lRet := .T.
Local cCpo := Upper(VtReadVar())
Local lSemEnd	:= .F.

SBE->(DbSetOrder(1))
If ! SBE->(MsSeek(xFilial("SBE")+cArmazem+cEndereco))
	If  cCpo =="CENDORI" 
		VtAlert(STR0007, STR0013,.t.,4000,4) //'Endereco nao encontrado'###'Aviso'
		VTKeyboard(chr(20))
		If ! UsaCb0("02")
			VTClearGet("cArmOri")
			VTClearGet("cEndOri")
			VTGetSetFocus("cArmOri")	
		EndIf
		lRet := .F.
	elseIf Empty(cEndereco) .and. !Empty(cArmazem)
		If NNR->(MsSeek(xFilial("NNR")+cArmazem))
			lRet := .T.
			lSemEnd := .T.
		else
			lRet := .F.
			VtAlert(STR0062, STR0013,.t.,4000,4)// armazem n„o localizado
		EndIf
	else
		VtAlert(STR0007, STR0013,.t.,4000,4) //'Endereco nao encontrado'###'Aviso'
		VTKeyboard(chr(20))
		If ! UsaCb0("02")
			VTClearGet("cArmDes")
			VTClearGet("cEndDes")
			VTGetSetFocus("cArmDes")
		EndIf

	Endif 
EndIf
If !lSemEnd 
	If lRet .And. ! CBEndLib(cArmazem,cEndereco) .Or. !ExistCpo("SBE",cArmazem+cEndereco,,,.F.) //5∫ parametro retira o Help
		VtAlert(STR0008, STR0013,.t.,4000,4) //'Endereco bloqueado'###'Aviso'
		VTKeyboard(chr(20))
		If ! UsaCb0("02")
			If cCpo =="CENDORI"
				VTClearGet("cArmOri")
				VTClearGet("cEndOri")
				VTGetSetFocus("cArmOri")
			Else
				VTClearGet("cArmDes")
				VTClearGet("cEndDes")
				VTGetSetFocus("cArmDes")
			EndIf
		EndIf
		lRet := .F.
	EndIf
EndIf
// ---- Ponto de Entrada para validar ENDERECO (ORIGEM E DESTINO)
If	lRet .And. ExistBlock("ACD150VE")
	lRet := ExecBlock("ACD150VE", .F., .F., {cArmazem, cEndereco, nTipo} )
	lRet := If(ValType(lRet)=="L",lRet,.T.)
EndIf
Return lRet

//‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹
Static Function GetProduto()
IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf
If UsaCB0('01') .and. UsaCB0('02')
	If lVT100B // GetMv("MV_RF4X20")
		lVolta := .F.
		++nLin
		@ ++nLin,0 VTSAY STR0009 //"Produto/Endereco"
		@ ++nLin,0 VTGET cCB0ProdEnd PICTURE "@!" valid VldProduto("01/02")
		VTRead
	Else
		@ ++nLin,0 VTSAY STR0009 //"Produto/Endereco"
		@ ++nLin,0 VTGET cCB0ProdEnd PICTURE "@!" valid VldProduto("01/02")
	EndIf
ElseIf UsaCB0('01') .and. ! UsaCB0('02')
	If lVT100B // GetMv("MV_RF4X20")
		lVolta := .F.
		++nLin
		@ ++nLin,0 VTSAY STR0010 //"Produto"
		@ ++nLin,0 VTGET cCB0ProdEnd PICTURE "@!" valid VldProduto("01")
		VTRead
	Else
		@ ++nLin,0 VTSAY STR0010 //"Produto"
		@ ++nLin,0 VTGET cCB0ProdEnd PICTURE "@!" valid VldProduto("01")
	EndIf
ElseIf ! UsaCB0('01')
	If lVT100B // GetMv("MV_RF4X20")
		lVolta := .F.
		@ ++nLin,0 VTSAY STR0011 //"Quantidade"
		@ ++nLin,0 VTGET nQtde PICTURE CBPictQtde() valid nQtde > 0 when (lForcaQtd .or. VTLastkey() == 5)
		VTRead
		VTClear()
		nLin := -1
		@ ++nLin,0 VTSAY STR0010 //"Produto"
		@ ++nLin,0 VTGET cProduto    PICTURE "@!" valid VldProduto("") when iif(vtRow() == 1 .and. vtLastKey() == 5,(VTKeyBoard(chr(27)),lVolta := .T.),.T.)
	Else
		@ ++nLin,0 VTSAY STR0011 //"Quantidade"
		@ ++nLin,0 VTGET nQtde PICTURE CBPictQtde() valid nQtde > 0 when (lForcaQtd .or. VTLastkey() == 5)
		@ ++nLin,0 VTSAY STR0010 //"Produto"
		@ ++nLin,0 VTGET cProduto    PICTURE "@!" valid VldProduto("")
	EndIf
EndIf
Return

//‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹
Static Function VldProduto(cTipo)
Local cTipId
Local cEtiqueta
Local cMvCq			:= GetMvNNR('MV_CQ','98')
Local aEtiqueta    :={}
Local nQE
Local nQtdeLida    :=0
Local nSaldo
Local nI
Local aListaBKP
Local aHisEtiBKP
Local aItensPallet
Local lIsPallet    := .t.
Local nP           := 0
Local lRet         := .T.
Local cV150PROD    := cProduto
Local nCodEtq2     := TamSX3("CB0_CODET2")[1]
Local nBFNumSer    := TamSX3("BF_NUMSERI")[1]
Local lAV150VPR    := ExistBlock("AV150VPR")
Local lForcaArm    := .F.
Local cCbEndCQ		 := GetMv("MV_CBENDCQ")
Local lWmsNew      := SuperGetMV("MV_WMSNEW",.F.,.F.)

If ExistBlock("V150PROD")
	ExecBlock("V150PROD", .F., .F., {cTipo})
	If !ValType(cProduto)=="C"
		cProduto := cV150PROD 
	EndIf
EndIf

If	"01" $ cTipo
	If	Empty(cCB0ProdEnd)
		If	cTipo=="01"
			Return .t.
		Else
			Return .f.
		EndIf
	EndIf
	cEtiqueta    := cCB0ProdEnd
	aItensPallet := CBItPallet(cCB0ProdEnd)
	lIsPallet    := .t.
	If	Len(aItensPallet) == 0
		aItensPallet:={cCB0ProdEnd}
		lIsPallet := .f.
	EndIf
	cTipId:=CBRetTipo(cCB0ProdEnd)
	If	cTipId == "01" .and. cTipId $ cTipo .or. lIsPallet
		aListaBKP := aClone(aLista)
		aHisEtiBKP:= aClone(aHisEti)

		Begin Sequence
		For nP := 1 to len(aItensPallet)
			cCB0ProdEnd :=  padr(aItensPallet[nP],nCodEtq2)
			aEtiqueta:= CBRetEti(cCB0ProdEnd,"01")
			If Empty(aEtiqueta)
				VTALERT(STR0012,STR0013 ,.T.,4000,3) //"Etiqueta invalida"###"Aviso"
				Break
			EndIf
			//--Valida se a etiqueta j· foi consumida por outro processo
			If CB0->CB0_STATUS $ "123"  
				VTBeep(2)
				VTAlert(STR0012,STR0013,.T.,4000,3) //"Etiqueta invalida"###"Aviso"
				Break
			EndIf			
			If ! lIsPallet .and. ! Empty(CB0->CB0_PALLET)
				VTALERT(STR0014,STR0013,.T.,4000,2) //"Etiqueta invalida, Produto pertence a um Pallet"###"Aviso"
				break
			EndIf
			If ! Empty(aEtiqueta[2]) .and. Ascan(aHisEti,cCB0ProdEnd) > 0
				VTALERT(STR0015,STR0013 ,.T.,4000,3) //"Etiqueta ja lida"###"Aviso"
				Break
			EndIf
			If lWmsNew .And. IntWms(aEtiqueta[1])
				VTALERT(STR0063,STR0013, .T.)//"Para o Novo WMS e Produtos com controle WMS usar: Menu Atualizacoes>Estoque WMS>Transferencia WMS" 
				Break
			EndIf
			If aEtiqueta[10] == cMvCq // asv verificar se eh possivel fazer este tipo de mov. no CQ
				VtAlert(STR0047, STR0013,.t.,4000,4) //'Aviso' //"Esta rotina nao trata armazem de CQ!"
				VtAlert(STR0048, STR0013,.t.,4000,4) //'Aviso' //"Utilize as rotinas de Envio/Baixa CQ!"
				Break
			EndIf
			If !Localiza(aEtiqueta[1])
				VTALERT(STR0016,STR0013,.T.,4000,3) //"Produto lido nao controla endereco!"###"Aviso"
				VTALERT(STR0017,STR0013,.T.,4000) //"Utilize rotina especifica ACDV151"###"Aviso"
				Break
			EndIf
			If !Empty(aEtiqueta[13])
				VTALERT(STR0054,STR0013 ,.T.,4000,3) //"Etiqueta utilizada em NF saida."###"Aviso"
				Break
			EndIf
			If Empty(aEtiqueta[2])
				aEtiqueta[2]:= 1
			EndIf
			cArmOri := aEtiqueta[10]
			cEndOri := aEtiqueta[9]
			cLote   := aEtiqueta[16]
			cSLote  := aEtiqueta[17]
			cNumSerie:=CB0->CB0_NUMSER
			If ! CBProdLib(cArmOri,aEtiqueta[1])
				Break
			EndIF
			If Empty(cEndOri)
				VTALERT(STR0018,STR0013 ,.T.,4000,3) //"Produto nao distribuido"###"Aviso"
				Break
			EndIf
			If ! ValidEnd(cArmOri,cEndOri,1)
				Break
			EndIf
			cProduto  := aEtiqueta[1]

			If Empty(cProduto)
				VTALERT(STR0051,STR0013,.T.,4000,3) //"Aviso" //"Inconsistencia na leitura da etiqueta"
				Break
			EndIf

			nQE:= 1
			If ! CBProdUnit(aEtiqueta[1])
				nQE := CBQtdEmb(aEtiqueta[1])
				If empty(nQE)
					Break
				EndIf
			EndIf
			nQtdeProd := aEtiqueta[2]*nQE
			nQtdeLida := 0
			aEval(aLista,{|x|  If(x[1]+x[3]+x[4]+x[5]+x[6]+x[7]==cProduto+cArmOri+cEndOri+cLote+cSLote+cNumSerie,nQtdeLida+=x[2],nil)})
			nSaldo := SaldoSBF(cArmOri,cEndOri,cProduto,cNumSerie,cLote,cSLote)
			If nQtdeProd+nQtdeLida >  nSaldo
				VTALERT(STR0019,STR0013 ,.T.,4000,3) //"Quantidade excede o saldo disponivel"###"Aviso"
				Break
			EndIf
			If	lAV150VPR
				lRet := ExecBlock("AV150VPR", .F., .F., {cEtiqueta})
				If	!If(ValType(lRet)=="L",lRet,.T.)
					Break
				EndIf
			EndIf
			TrataArray(cCB0ProdEnd)
		Next
		cProduto   := Space(nCodEtq2)
		VTKeyboard(chr(20))
		Return .f.
		End Sequence
		aLista := aClone(aListaBKP)
		aHisEti:= aClone(aHisEtiBKP)
		cProduto   := Space(nCodEtq2)
		VTKeyboard(chr(20))
		Return .f.
	ElseIf cTipId == "02" .and. cTipId $ cTipo
		aEtiqueta:= CBRetEti(cCB0ProdEnd,"02")
		If Empty(aEtiqueta)
			VTALERT(STR0012,STR0013 ,.T.,4000,3) //"Etiqueta invalida"###"Aviso"
			cProduto   := Space(nCodEtq2)
			VTKeyboard(chr(20))
			Return .f.
		EndIf
		If Empty(aLista)
			VTALERT(STR0020,STR0013 ,.T.,4000,3) //"Informe o produto"###"Aviso"
			cProduto   := Space(nCodEtq2)
			VTKeyboard(chr(20))
			Return .f.
		EndIf	
		cArmDes := aEtiqueta[2]
		cEndDes := aEtiqueta[1]

		If lCQ <> NIL
			If ! lCQ // transferencia de armazens normais exceto CQ
				If Trim(cArmDes) == cMvCq
					VtAlert(STR0047, STR0013,.t.,4000,4) //'Aviso' //"Esta rotina nao trata armazem de CQ!"
					VtAlert(STR0048, STR0013,.t.,4000,4) //'Aviso' //"Utilize as rotinas de Envio/Baixa CQ!"
					VTKeyboard(chr(20))
					Return .f.
				EndIf
			ELSE
				If Trim(cArmDes) <> cMvCq
					VtAlert(STR0049, STR0013,.t.,4000,4) //'Aviso' //"Armazem destino nao eh de CQ!"
					VTKeyboard(chr(20))
					Return .f.
				EndIf
				If !Empty(cCbEndCQ) .And. !(cArmDes+AllTrim(cEndDes)+";" $ cCbEndCQ)
					VtAlert(STR0050, STR0013,.t.,4000,4) //'Aviso' //"Armazem destino nao eh de Inspecao de CQ!"
					VTKeyboard(chr(20))
					Return .f.
				EndIf
			EndIf
		EndIf

		If Ascan(aLista,{|x| x[3]+x[4] ==cArmDes+cEndDes}) > 0
			VTALERT(STR0021,STR0013 ,.T.,4000,3) //"Endereco de origem igual ao destino"###"Aviso"
			cProduto   := Space(nCodEtq2)
			VTKeyboard(chr(20))
			Return .f.
		EndIF
		If ! ValidEnd(cArmDes,cEndDes,2)
			cProduto   := Space(nCodEtq2)
			VTKeyboard(chr(20))
			Return .f.
		EndIf
		For nI := 1 to Len(aLista)
			If ! CBProdLib(cArmDes,aLista[nI,1],.f.)
				VTALERT(STR0022+aLista[nI,1]+STR0023+cArmDes,STR0013,.t.,4000,2) //'Produto '###' bloqueado para inventario no armazem '###"Aviso"
				cProduto   := Space(nCodEtq2)
				VTKeyboard(chr(20))
				Return .f.
			EndIF
		Next
		GravaTransf()
		cProduto   := Space(nCodEtq2)
		VTKeyboard(chr(20))
		Return .f.
	Else
		VTALERT(STR0012,STR0013 ,.T.,4000,3) //"Etiqueta invalida"###"Aviso"
		cProduto   := Space(nCodEtq2)
		VTKeyboard(chr(20))
		Return .f.
	EndIf
Else
	If	Empty(cProduto) .And. Len(aLista) > 0
		Return .t.
	ElseIf	Empty(cProduto) .And. Len(aLista) <= 0
		VTKeyBoard(chr(23))
		Return .f.
	EndIf
	aItensPallet := CBItPallet(cProduto)
	If	Len(aItensPallet) == 0
		aItensPallet:={cProduto}
	EndIf

	aListaBKP := aClone(aLista)
	aHisEtiBKP:= aClone(aHisEti)
	Begin Sequence
	For nP := 1 to len(aItensPallet)
		cProduto :=  padr(aItensPallet[nP],nCodEtq2)
		If	! CBLoad128(@cProduto)
			Break
		EndIf
		cTipId:=CBRetTipo(cProduto)
		If	! cTipId $ "EAN8OU13-EAN14-EAN128"
			VTALERT(STR0024,STR0013,.T.,4000,3) //"Etiqueta invalida."###"Aviso"
			Break
		EndIf
		cEtiqueta := cProduto
		aEtiqueta := CBRetEtiEAN(cProduto)
		If	Empty(aEtiqueta) .or. Empty(aEtiqueta[2])
			VTALERT(STR0024,STR0013,.T.,4000,3) //"Etiqueta invalida."###"Aviso"
			Break
		EndIf
		If lWmsNew .And. IntWms(aEtiqueta[1])
			VTALERT(STR0063,STR0013,.T.)//"Para o Novo WMS e Produtos com controle WMS usar: Menu Atualizacoes>Estoque WMS>Transferencia WMS"
			Break
		EndIf
		If	! CBProdLib(cArmOri,aEtiqueta[1])
			Break
		EndIf
		nQE:= 1
		If	! CBProdUnit(aEtiqueta[1])
			nQE := CBQtdEmb(aEtiqueta[1])
			If	Empty(nQE)
				Break
			EndIf
		EndIf
		cLote := aEtiqueta[3]
		If	! CBRastro(aEtiqueta[1],@cLote,@cSLote)
			Break
		EndIf
		
		cLoteDes  := Space(nTamLote)
		cSLoteDes := Space(nTamSLote)
		dValLtDes := CTOD("")
		
		If lTrfLote .And. !Empty(cLote) 
			If	VTYesNo(STR0064,STR0013 ,.T.) //"Confirma transferencia?"###"Aviso"
				VTALERT(STR0055,"",.T.,3000,3) // "Digite o lote destino"
				If	! CBRastro(aEtiqueta[1],@cLoteDes,@cSLoteDes, @dValLtDes, .F.)
					Break
				EndIf
			Endif
		EndIf	
		If	!Localiza(aEtiqueta[1])
			VTALERT(STR0016,STR0013,.T.,4000,3) //"Produto lido nao controla endereco!"###"Aviso"
			VTALERT(STR0017,STR0013,.T.,4000) //"Utilize rotina especifica ACDV151"###"Aviso"
			Break
		EndIf
		cProduto  := aEtiqueta[1]
		If	Empty(cProduto)
			VTALERT(STR0051,STR0013,.T.,4000,3) //"Aviso" //"Inconsistencia na leitura da etiqueta"
			Break
		EndIf

		nQtdeProd := aEtiqueta[2]*nQtde*nQE
		If	Len(aEtiqueta) >= 5
			cNumSerie:=Padr(aEtiqueta[5],Len(Space(nBFNumSer)))
		EndIf
		If CBChkSer(aEtiqueta[1]) .and. ! CBNumSer(@cNumSerie,Nil,aEtiqueta)
			Break
		Endif
		nQtdeLida := 0
		aEval(aLista,{|x|  If(x[1]+x[3]+x[4]+x[5]+x[6]+x[7]==cProduto+cArmOri+cEndOri+cLote+cSLote+cNumSerie,nQtdeLida+=x[2],nil)})
		nSaldo := SaldoSBF(cArmOri,cEndOri,cProduto,cNumSerie,cLote,cSLote)
		If	nQtdeProd+nQtdeLida > nSaldo
			VTALERT(STR0019,STR0013 ,.T.,4000,3) //"Quantidade excede o saldo disponivel"###"Aviso"
			Break
		EndIf
		If	lAV150VPR
			lRet := ExecBlock("AV150VPR", .F., .F., {cEtiqueta})
			If	!If(ValType(lRet)=="L",lRet,.T.)
				Break
			EndIf
		EndIf
		TrataArray(Nil)
	Next
	If ExistBlock("V150FArm")
		lForcaArm := ExecBlock("V150FArm",.f.,.f.)
		If ValType(lForcaArm)<> "L"
			lForcaArm := .f.   
	    EndIf
	EndIf

	If lForcaArm .And. lForcaQtd 
		VTGetSetFocus("cArmDes")
	EndIf

	If	lForcaQtd .And. !lForcaArm
		VtClearGet('cProduto')
		VtGetSetFocus('cProduto')
		VtGetSetFocus('nQtde')
	Else
		VTKeyBoard(chr(20))
	EndIf
	
	If lForcaArm 
		VTGetSetFocus("cArmDes")
	Else
		nQtde := 1
		VTGetRefresh('nQtde')
	EndIf

	cProduto  := Space(nCodEtq2)
	cLote     := Space(nTamLote)
	cSLote    := Space(nTamSLote)
	cNumSerie := Space(nBFNumSer)
	Return .f.
	End Sequence
	aLista   := aClone(aListaBKP)
	aHisEti  := aClone(aHisEtiBKP)
	cProduto := Space(nCodEtq2)
	VTKeyboard(chr(20))
	Return .f.
EndIf
Return .f.

//‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹
Static Function GetEndDes()
Local cPictNNR := PesqPict( "NNR","NNR_CODIGO")
IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf
If lVT100B // GetMv("MV_RF4X20")
	VTClear(1,0,3,19)
	
	If Len(cArmDes) > 2
		nLin := 0
	Else
		nLin := 1
	EndIf
	
	If ! (UsaCB0('01') .and. UsaCB0('02'))
		@ ++nLin,0 VtSay STR0025 //'Endereco destino'
		If UsaCB0('02')
			@ ++nLin,0 VTGet cCB0EndDes Valid VTLastkey() == 5 .or. VldEndDes() ;
				when iif(vtRow() == 3 .and. vtLastKey() == 5,(VTKeyBoard(chr(27)),lVolta := .T.),.T.)
		Else
			@ ++nLin,0 VTGet cArmDes pict cPictNNR Valid VTLastkey() == 5 .or. ! Empty(cArmDes) ;
				when iif(vtRow() == 3 .and. vtCol() < 19 .and. vtLastKey() == 5,(VTKeyBoard(chr(27)),lVolta := .T.),.T.)
			If Len(cArmDes) > 2
				nLin++
			EndIf
			@ nLin,2 VTSay '-' VtGet cEndDes Pict '@!' Valid  VTLastkey() == 5 .or. VldEndDes()
		EndIf
	EndIf
Else
	If ! (UsaCB0('01') .and. UsaCB0('02'))
		@ ++nLin,0 VtSay STR0025 //'Endereco destino'
		If UsaCB0('02')
			@ ++nLin,0 VTGet cCB0EndDes Valid VTLastkey() == 5 .or. VldEndDes()
		Else
			@ ++nLin,0 VTGet cArmDes pict cPictNNR Valid VTLastkey() == 5 .or. ! Empty(cArmDes)
			If Len(cArmDes) > 2
				nLin++
			EndIf
			@ nLin,2 VTSay '-' VtGet cEndDes Pict '@!' Valid  VTLastkey() == 5 .or. VldEndDes()
		EndIf
	EndIf
EndIf
Return

//‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹
Static Function VldEndDes()
Local aEnd		:= {}
Local nI
Local cMvCq	:= GetMvNNR('MV_CQ','98')
Local cCbEndCQ	:= GetMv("MV_CBENDCQ")

//⁄ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒø	
//≥ Ponto de Entrada utilizado para manipular o codigo da etiqueta lida ≥
//¿ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒŸ
If ExistBlock("ACD150ME")
	If UsaCB0("02")
		aEnd := ExecBlock("ACD150ME", .F., .F., {2,cCB0EndDes} )
		If ValType(aEnd) == "A"
			cCB0EndDes := aEnd[1]
		EndIf
	Else
		aEnd := ExecBlock("ACD150ME", .F., .F., {2,cArmDes,cEndDes} )
		If ValType(aEnd) == "A"
			cArmDes := aEnd[1]
			cEndDes := aEnd[2]
		EndIf
	EndIf
	
	aEnd := {}
EndIf
If Empty( cCB0EndDes+cArmDes+cEndDes)
	VtAlert(STR0005, STR0013,.t.,4000,4) //'Endereco invalido'###'Aviso'
	VTKeyboard(chr(20))
	If ! UsaCb0("02")
		VTClearGet("cArmDes")
		VTGetSetFocus("cArmDes")
	EndIf
	Return .f.
EndIF
If UsaCB0('02')
	aEnd:= CbRetEti(cCB0EndDes,"02")
	If Empty(aEnd)
		VtAlert(STR0005, STR0013,.t.,4000,4) //'Endereco invalido'###'Aviso'
		VTKeyboard(chr(20))
		Return .f.
	EndIf
	cArmDes   := aEnd[2]
	cEndDes   := aEnd[1]
EndIF
If lCQ <> NIL
	If ! lCQ // transferencia de armanzens normais exceto CQ
		If Trim(cArmDes) == cMvCq
			VtAlert(STR0047, STR0013,.t.,4000,4) //'Aviso' //"Esta rotina nao trata armazem de CQ!"
			VtAlert(STR0048, STR0013,.t.,4000,4) //'Aviso' //"Utilize as rotinas de Envio/Baixa CQ!"
			VTKeyboard(chr(20))
			If ! UsaCb0("02")
				VTClearGet("cArmDes")
				VTClearGet("cEndDes")
				VTGetSetFocus("cArmDes")
			EndIf
			Return .f.
		EndIf
	ELSE
		If Trim(cArmDes) <> cMvCq
			VtAlert(STR0049, STR0013,.t.,4000,4) //'Aviso' //"Armazem destino nao eh de CQ!"
			VTKeyboard(chr(20))
			If ! UsaCb0("02")
				VTClearGet("cArmDes")
				VTClearGet("cEndDes")
				VTGetSetFocus("cArmDes")
			EndIf
			Return .f.
		EndIf
		If !Empty(cCbEndCQ) .And. !(cArmDes+AllTrim(cEndDes)+";" $ cCbEndCQ)
			VtAlert(STR0050, STR0013,.t.,4000,4) //'Aviso' //"Armazem destino nao eh de Inspecao de CQ!"
			VTKeyboard(chr(20))
			If ! UsaCb0("02")
				VTClearGet("cArmDes")
				VTClearGet("cEndDes")
				VTGetSetFocus("cArmDes")
			EndIf
			Return .f.
		EndIf
	EndIf
EndIf
if Len(aLista)<1
	VTALERT(STR0012+": "+STR0020,STR0013,.T.,4000,3) //"Etiqueta invalida"###"Informe o produto"###"Aviso"
	VTKeyboard(chr(20))
	If ! UsaCb0("02")
		VTClearGet("cArmDes")
		VTClearGet("cEndDes")
		VTGetSetFocus("cArmDes")
	EndIf
	Return .f.
Endif
For nI := 1 to Len(aLista)
	If ! CBProdLib(cArmDes,aLista[nI,1],.f.)
		VTALERT(STR0022+aLista[nI,1]+STR0023+cArmDes,STR0013,.t.,4000,2) //'Produto '###' bloqueado para inventario no armazem '###'Aviso'
		VTKeyboard(chr(20))
		If ! UsaCb0("02")
			VTClearGet("cArmDes")
			VTGetSetFocus("cArmDes")
		EndIf
		Return .f.
	EndIF
Next
If ! ValidEnd(cArmDes,cEndDes,2)
	Return .f.
EndIf
If (ValType(aLista[1]) == "A" .And. Len(aLista[1]) < 8 .And. Ascan(aLista,{|x| x[3]+x[4] == cArmDes+cEndDes}) > 0) .Or.;
   (ValType(aLista[1]) == "A" .And. Len(aLista[1]) >= 8 .And. Ascan(aLista,{|x| x[3]+x[4]+x[5]+x[6] == cArmDes+cEndDes+cLoteDes+cSLoteDes}) > 0)
	VTALERT(STR0021,STR0013 ,.T.,4000,3) //"Endereco de origem igual ao destino"###"Aviso"
	VTKeyboard(chr(20))
	If ! UsaCb0("02")
		VTClearGet("cArmDes")
		VTGetSetFocus("cArmDes")
	EndIf
	Return .f.
EndIF
If ! GravaTransf()
	Return .f.
EndIf
Return .t.

//‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹
Static Function TrataArray(cEtiqueta,lEstorno)
Local nPos		:= 0
Local nAchou	:= 0
Default lEstorno := .f.
If ! lEstorno
	If cEtiqueta <> NIL
		aadd(aHisEti,cEtiqueta)
	EndIf
	nPos := aScan(aLista,{|x| x[1]+x[3]+x[4]+x[5]+x[6]+x[7]+x[8]+x[9] == cProduto+cArmOri+cEndOri+cLote+cSLote+cNumSerie+cLoteDes+cSLoteDes})
	If Empty(nPos)
		aadd(aLista,{cProduto,nQtdeProd,cArmOri,cEndOri,cLote,cSLote,cNumSerie,cLoteDes,cSLoteDes,{{cEtiqueta,nQtdeProd}}})
		nPos := Len(aLista)
	Else
		aLista[nPos,2]+=nQtdeProd
		If (nAchou := aScan(aTail(aLista[nPos]),{|x| x[1] == cEtiqueta})) == 0	//-- Etiqueta lida anteriormente (qtd variavel)
			aAdd(aLista[nPos,10],{cEtiqueta,nQtdeProd})	//-- Adiciona
		Else
			aLista[nPos,10,nAchou,2] += nQtdeProd	//-- Adiciona
		EndIf
	EndIf

	If ExistBlock("AV150ARR")
		ExecBlock("AV150ARR",.F.,.F.,{cEtiqueta,nPos})
	EndIf
Else
	nPos := aScan(aLista,{|x| x[1]+x[3]+x[4]+x[5]+x[6]+x[7] == cProduto+cArmOri+cEndOri+cLote+cSLote+cNumSerie})
	aLista[nPos,2] -= nQtdeProd
	If Empty(aLista[nPos,2])	//-- Se zerou quantidade do item a transferir, remove
		aDel(aLista,nPos)
		aSize(aLista,len(aLista)-1)
	ElseIf cEtiqueta <> NIL .And. (nAchou := aScan(aLista[nPos,10],{|x| x[1] == cEtiqueta})) > 0	//-- Sen„o, remove etiqueta
		aDel(aLista[nPos,10],nAchou)
		aSize(aLista[nPos,10],len(aLista[nPos,10])-1)
	EndIf
	If cEtiqueta <> NIL
		nPos := aScan(aHisEti,cEtiqueta)
		aDel(aHisEti,nPos)
		aSize(aHisEti,len(aHisEti)-1)
	EndIf
EndIf
Return

//‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹
Static Function GravaTransf()
Local aSave
Local aTransf    :={}
Local dValid
Local dValidDest
Local lRet		 := .t.
Local nI	     	 	
Local lV150AUTO  := ExistBlock("V150AUTO")
Local lConfirma := .F.
	
Private nModulo  := 4

/* Ponto de entrada permite validar a continuidade do processo.*/
If	ExistBlock("AV150CF") 
	lRet := ExecBlock("AV150CF")
	If ValType(lRet)=="L" .And. !lRet 
		VTKeyboard(chr(20))
		Return lRet
	EndIf
EndIf

If	! VTYesNo(STR0026,STR0013 ,.T.) //"Confirma transferencia?"###"Aviso"
	VTKeyboard(chr(20))
	Return .F.
Else
	aSave := VTSAVE()
	VTClear()
	VTMsg(STR0027) //'Aguarde...'
	Begin Transaction
	lMsErroAuto := .F.
	lMsHelpAuto := .T.
	aTransf:=Array(len(aLista)+1)
	aTransf[1] := {"",dDataBase}
	For nI := 1 to Len(aLista)
		SB1->(DbSetOrder(1))
		If !SB1->(MsSeek(xFilial("SB1")+aLista[nI,1]))
			VTALERT(STR0052+aLista[nI,1]+STR0053,"Aviso",.T.,NIL,3) //'Produto "'###'" nao localizado no cadastro de produtos.'
			DisarmTransaction()
			Break
		EndIf
				
		If Rastro(aLista[nI,1])
			SB8->(DbSetOrder(3))
			SB8->(DbSeek(xFilial("SB8")+aLista[nI,1]+aLista[nI,3]+aLista[nI,5]+AllTrim(aLista[nI,6])))
			dValid := SB8->B8_DTVALID
			dValidDest := dValid
			
			If SB8->(DbSeek(xFilial("SB8")+aLista[nI,1]+cArmDes+aLista[nI,8]+aLista[nI,9]))
				dValidDest := SB8->B8_DTVALID
			EndIf
			
			If !lConfirma .And. !Empty(dValidDest) .And. dValidDest != dValid
				VTALERT(STR0058,STR0013,.T.,NIL,3) // "Existe(m) produto(s) com data de validade diferente no armazem de destino"
				If !VTYesNo(STR0059,STR0060,.t.) // "Sera assumida a data de validade do armazem de destino. Confirmar operacao?", "Lote Destino"
					VTALERT(STR0061,STR0013,.T.,NIL,3) // "Processo cancelado"
					DisarmTransaction()
					Break
				Else
					lConfirma := .T.
				EndIf
			EndIf
		EndIf

		aTransf[nI+1]:= {{"D3_COD"    , SB1->B1_COD                ,NIL}} 
		aAdd(aTransf[nI+1],{"D3_DESCRI" , SB1->B1_DESC               ,NIL}) 
		aAdd(aTransf[nI+1],{"D3_UM"     , SB1->B1_UM                 ,NIL}) 
		aAdd(aTransf[nI+1],{"D3_LOCAL"  , aLista[nI,3]               ,NIL}) 
		aAdd(aTransf[nI+1],{"D3_LOCALIZ", aLista[nI,4]           	  ,NIL}) 
		aAdd(aTransf[nI+1],{"D3_COD"    , SB1->B1_COD             	  ,NIL}) 
		aAdd(aTransf[nI+1],{"D3_DESCRI" , SB1->B1_DESC               ,NIL}) 
		aAdd(aTransf[nI+1],{"D3_UM"     , SB1->B1_UM             	  ,NIL}) 
		aAdd(aTransf[nI+1],{"D3_LOCAL"  , cArmDes             		  ,NIL}) 
		aAdd(aTransf[nI+1],{"D3_LOCALIZ", cEndDes               	  ,NIL}) 
		aAdd(aTransf[nI+1],{"D3_NUMSERI", aLista[nI,7]               ,NIL})//numserie 
		aAdd(aTransf[nI+1],{"D3_LOTECTL", aLista[nI,5]               ,NIL})//lote
		aAdd(aTransf[nI+1],{"D3_NUMLOTE", aLista[nI,6]               ,NIL})//sublote
		aAdd(aTransf[nI+1],{"D3_DTVALID", dValid                     ,NIL}) 
		aAdd(aTransf[nI+1],{"D3_POTENCI", criavar("D3_POTENCI")      ,NIL}) 
		aAdd(aTransf[nI+1],{"D3_QUANT"  , aLista[nI,2]               ,NIL}) 
		aAdd(aTransf[nI+1],{"D3_QTSEGUM", criavar("D3_QTSEGUM")      ,NIL}) 
		aAdd(aTransf[nI+1],{"D3_ESTORNO", criavar("D3_ESTORNO")      ,NIL}) 
		aAdd(aTransf[nI+1],{"D3_NUMSEQ" , criavar("D3_NUMSEQ")		  ,NIL}) 
		aAdd(aTransf[nI+1],{"D3_LOTECTL", IIF(!Empty(aLista[nI,8]),aLista[nI,8],aLista[nI,5])     ,NIL})
		aAdd(aTransf[nI+1],{"D3_DTVALID", dValid                     ,NIL}) 
		/*Ponto de entrada, permite manipular e ou acrescentar dados no array aTransf.*/
		If lV150AUTO
			aPEAux := aClone(aTransf)  
			aPEAux := ExecBlock("V150AUTO",.F.,.F.,{aTransf})
			If ValType(aPEAux)=="A" 
				aTransf := aClone(aPEAux)
			EndIf
		EndIf
		If ! UsaCB0("01")
			CBLog("02",{SB1->B1_COD,aLista[nI,2],aLista[nI,5],aLista[nI,6],aLista[nI,3],aLista[nI,4],cArmDes,cEndDes})
		EndIf
	Next
	MSExecAuto({|x| MATA261(x)},aTransf)

	If lMsErroAuto
		VTALERT(STR0028,STR0029,.T.,4000,3) //"Falha na gravacao da transferencia"###"ERRO"
		DisarmTransaction()
		Break
	Else
		If	ExistBlock("ACD150GR")
			ExecBlock("ACD150GR",.F.,.F.)
		EndIf
	EndIf
	End Transaction
	VtRestore(,,,,aSave)
	If	lMsErroAuto
		VTDispFile(NomeAutoLog(),.t.)
	Else
		If	ExistBlock("ACD150OK")
			ExecBlock("ACD150OK",.F.,.F.)
		EndIf
		cArmOri     := Space(Tamsx3("B1_LOCPAD") [1])
		cEndOri     := Space(TamSX3("BF_LOCALIZ")[1])
		cCB0EndOri  := Space(20)
		cArmDes     := Space(Tamsx3("B1_LOCPAD") [1])
		cEndDes     := Space(TamSX3("BF_LOCALIZ")[1])
		cCB0EndDes  := Space(20)
		cCB0ProdEnd := IIf( FindFunction( 'AcdGTamETQ' ), AcdGTamETQ(), Space(48) )
		cProduto    := IIf( FindFunction( 'AcdGTamETQ' ), AcdGTamETQ(), Space(48) )
		cLote       := Space(TamSX3("B8_LOTECTL")[1])
		cSLote      := Space(TamSX3("B8_NUMLOTE")[1])
		cNumSerie   := Space(TamSX3("BF_NUMSERI")[1])
		cLoteDes    := Space(TamSX3("B8_LOTECTL")[1])
		cSLoteDes   := Space(TamSX3("B8_NUMLOTE")[1])
		dValLtDes   := CTOD("")
		nQtde       := 1
		aLista      := {}
		aHisEti     := {}
	EndIf
EndIf
Return .t.


//‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹
Static Function Informa()

Local aCab  := {STR0010,STR0011,STR0032,STR0033,STR0034,STR0035,STR0036} //"Produto"###"Quantidade"###"Armazem"###"Endereco"###"Lote"###"SubLote"###"Num.Serie"
Local aSize := {15,16,7,15,10,7,20}
Local aSave := VTSAVE()
IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf
VtClear()
VTaBrowse(0,0,IIf(lVT100B /*GetMv("MV_RF4X20")*/ ,3,7),19,aCab,aLista,aSize)
VtRestore(,,,,aSave)
Return

//‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹
Static Function Estorna()
Local aTela
Local cEtiqueta
Local nQtde := 1
aTela := VTSave()
VTClear()
cEtiqueta := Space(TamSx3("CB0_CODET2")[1])
@ 00,00 VtSay Padc(STR0037,VTMaxCol()) //"Estorno da Leitura"
If	! UsaCB0('01')
	@ 1,00 VTSAY STR0006 VTGet nQtde pict CBPictQtde() when VTLastkey() == 5 //"Qtde."
EndIf
@ 02,00 VtSay STR0038 //"Etiqueta:"
@ 03,00 VtGet cEtiqueta pict "@!" Valid VldEstorno(cEtiqueta,@nQtde)
VtRead
vtRestore(,,,,aTela)
Return

//‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹
Static Function VldEstorno(cEtiqueta,nQtde)
Local nPos
Local aEtiqueta,nQE
Local aListaBKP    := aClone(aLista)
Local aHisEtiBKP   := aClone(aHisEti)
Local aItensPallet := CBItPallet(cEtiqueta)
Local lIsPallet    := .t.
Local lRet         := .t.
Local nP           := 0
Local nCodEti2     := TamSX3("CB0_CODET2")[1]
Local nBFNumSer    := TamSX3("BF_NUMSERI")[1]

If	Empty(cEtiqueta)
	Return .f.
EndIf

If	Len(aItensPallet) == 0
	aItensPallet:={cEtiqueta}
	lIsPallet := .f.
EndIf

Begin Sequence
For nP := 1 to len(aItensPallet)
	cEtiqueta:=padr(aItensPallet[nP],nCodEti2)
	If	UsaCB0("01")
		nPos := Ascan(aHisEti, {|x| AllTrim(x) == AllTrim(cEtiqueta)})
		If	nPos == 0
			VTALERT(STR0039,STR0013,.T.,4000,2) //"Etiqueta nao encontrada"###"Aviso"
			Break
		EndIf
		aEtiqueta:=CBRetEti(cEtiqueta,'01')
		cProduto := aEtiqueta[1]
		cArmOri  := aEtiqueta[10]
		cEndOri  := aEtiqueta[9]
		cLote    := aEtiqueta[16]
		cSlote   := aEtiqueta[17]
		cNumSerie:=CB0->CB0_NUMSER

		If Empty(aEtiqueta[2])
			aEtiqueta[2] := 1
		EndIf
		nQtde := 1

		If ! lIsPallet .and. ! Empty(CB0->CB0_PALLET)
			VTALERT(STR0014,STR0013,.T.,4000,2)    //"Etiqueta invalida, Produto pertence a um Pallet"###"Aviso"
			break
		EndIf
	Else
		If ! CBLoad128(@cEtiqueta)
			Break
		EndIf
		aEtiqueta := CBRetEtiEAN(cEtiqueta)
		If	Len(aEtiqueta) == 0
			VTALERT(STR0012,STR0013,.T.,4000,2) //"Etiqueta invalida"###"Aviso"
			Break
		EndIf
		cProduto := aEtiqueta[1]
		If	ascan(aLista,{|x| x[1] ==cProduto}) == 0
			VTALERT(STR0040,STR0013,.T.,4000,2) //"Produto nao encontrado"###"Aviso"
			Break
		EndIf
		cLote := aEtiqueta[3]
		If	Len(aEtiqueta) >=5
			cNumSerie:= padr(aEtiqueta[5],Len(Space(nBFNumSer)))
		EndIf
		If CBChkSer(aEtiqueta[1]) .And. ! CBNumSer(cNumSerie,Nil,aEtiqueta)
			Break
		Endif
	EndIf

	nQE := 1
	If ! CBProdUnit(cProduto)
		nQE := CBQtdEmb(cProduto)
		If	Empty(nQE)
			Break
		EndIf
	EndIf
	nQtdeProd:=nQtde*nQE*aEtiqueta[2]

	If ! Usacb0("01") .and. ! CBRastro(cProduto,@cLote,@cSLote)
		Break
	EndIf

	nPos := Ascan(aLista,{|x| x[1]+x[3]+x[4]+x[5]+x[6]+x[7] == cProduto+cArmOri+cEndOri+cLote+cSLote+cNumSerie})
	If	nPos == 0
		VTALERT(STR0041,STR0013,.T.,4000,2) //"Produto nao encontrado neste endereco"###"Aviso"
		Break
	EndIf
	If	aLista[nPos,2] < nQtdeProd
		VTALERT(STR0042,STR0013,.T.,4000,2) //"Quantidade excede o estorno"###"Aviso"
		Break
	EndIf
	If	UsaCB0("01")
		TrataArray(cEtiqueta,.t.)
	Else
		TrataArray(,.t.)
	EndIf
Next
If	! VTYesNo(STR0043,STR0044,.t.) //"Confirma o estorno?"###"ATENCAO"
	Break
EndIf
/* Ponto de entrada permite validar a continuidade do processo.*/
If ExistBlock("AV150EST")
	lRet := ExecBlock("AV150EST", .F., .F., {cEtiqueta,nQtde})
	If	!If(ValType(lRet)=="L",lRet,.T.)
		Break
	EndIf
EndIf

nQtde:= 1
VTGetRefresh("nQtde")
VTKeyboard(chr(20))
Return .f.
End Sequence
aLista := aClone(aListaBKP)
aHisEti:= aClone(aHisEtiBKP)
nQtde  := 1
VTGetRefresh("nQtde")
VTKeyBoard(chr(20))
Return .f.
