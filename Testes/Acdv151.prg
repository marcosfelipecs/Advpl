#INCLUDE "Acdv151.ch" 
#include "protheus.ch"
#INCLUDE 'APVT100.CH'

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o    ³ acdv151    ³ Autor ³ Sandro              ³ Data ³ 14/08/01 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³ Programa de transferencia (sem Controle de Localizacao) 	  ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ SigaACD                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/   

Function acdv151()
Local aKey        := array(3)

Private cArmOri   := Space(Tamsx3("B1_LOCPAD") [1])
Private cArmDes   := Space(Tamsx3("B1_LOCPAD") [1])
Private cCB0Prod  := Space(TamSx3("CB0_CODET2")[1])
Private cProduto  := IIf( FindFunction( 'AcdGTamETQ' ), AcdGTamETQ(), Space(48) )

Private nQtde     := 1
Private nQtdeProd := 1
Private cLote     := Space(TamSX3("B8_LOTECTL")[1])
Private cSLote    := Space(TamSX3("B8_NUMLOTE")[1])
Private cNumSerie := Space(TamSX3("BF_NUMSERI")[1])
Private aLista    := {}
Private aHisEti   := {}
Private lMsErroAuto := .F.
Private nLin:= 0
Private lVolta := .F.
IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf

akey[2] := VTSetKey(24,{|| Estorna()},STR0040) //"Estorna"
akey[3] := VTSetKey(09,{|| Informa()},STR0041) //"Informacoes"

While .t.
	VTClear
	nLin:= -1
	@ ++nLin,0 VTSAY STR0001 //"Transferencia"
	If !lVolta
		If ! GetArmOri()
			Exit
		EndIf
	EndIf
	GetProduto()
	If lVT100B /* GetMv("MV_RF4X20") */ .and. VTLastKey() == 27
		Exit
	EndIf
	GetArmDes()
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
vtsetkey(24,akey[1])
vtsetkey(09,akey[2])
Return

//ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
Static Function GetArmOri()

Local cPictNNR := PesqPict( "NNR","NNR_CODIGO")

If ! UsaCB0('01')
	If VtModelo()=="RF"
		@ ++nLin,0 VtSay STR0004 //'Armazem oris em'
		@ ++nLin,0 VTGet cArmOri pict cPictNNR Valid VldArmOri()
	Else
		@ 0,0 VtSay STR0004 //'Armazem oris em'
		@ 1,0 VTGet cArmOri pict cPictNNR Valid VldArmOri()
	EndIf
	VTRead()
	If VTLastkey() == 27
		Return .f.
	EndIf
	VTClear(1,0,2,19)
	nLin := 0
EndIf
Return .t.

//ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
Static Function VldArmOri()
Local lRet := .T.
If Empty(cArmOri)
	VtAlert(STR0016, STR0017,.t.,4000,4) //'Armazem invalido'###'Aviso'
	VTKeyboard(chr(20))
	Return .f.
EndIf
If Trim(cArmOri) == GetMvNNR('MV_CQ','98')
	VtAlert(STR0042, STR0017,.t.,4000,4) //"Esta rotina nao trata armazem de CQ!"###'Aviso'
	VtAlert(STR0043, STR0017,.t.,4000,4) //"Utilize as rotinas de Envio/Baixa CQ!"###'Aviso'
	VTKeyboard(chr(20))
	Return .f.
EndIf
// Ponto de Entrada para validar Armazem Origem
If ExistBlock("AV151VLD")
	lRet := ExecBlock("AV151VLD",.F.,.F.,{1,cArmOri})
EndIf
Return lRet

//ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
Static Function GetProduto()
IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf
If UsaCB0('01')
	If lVT100B // GetMv("MV_RF4X20")
		lVolta := .F.
		++nLin
		@ ++nLin,0 VTSAY STR0005 //"Produto"
		@ ++nLin,0 VTGET cCB0Prod PICTURE "@!" valid VldProduto("01")
		VTRead
	Else
		If VtModelo()=="RF"
			@ ++nLin,0 VTSAY STR0005 //"Produto"
			@ ++nLin,0 VTGET cCB0Prod PICTURE "@!" valid VldProduto("01")
		Else
			VTClear()
			@ 0,0 VTSAY STR0005 //"Produto"
			@ 1,0 VTGET cCB0Prod PICTURE "@!" valid VldProduto("01")
			VTRead()
			VTClear()
		EndIF
	EndIf
ElseIf ! UsaCB0('01')
	If lVT100B // GetMv("MV_RF4X20")
		lVolta := .F.
		@ ++nLin,0 VTSAY STR0006 //"Quantidade"
		@ ++nLin,0 VTGET nQtde PICTURE CBPictQtde() valid nQtde > 0
		VTRead
		VTClear()
		nLin := -1
		@ ++nLin,0 VTSAY STR0005 //"Produto"
		@ ++nLin,0 VTGET cProduto    PICTURE "@!" valid VldProduto("") when iif(vtRow() == 1 .and. vtLastKey() == 5,(VTKeyBoard(chr(27)),lVolta := .T.),.T.)
	Else
		If VtModelo()=="RF"
			@ ++nLin,0 VTSAY STR0006 //"Quantidade"
			@ ++nLin,0 VTGET nQtde PICTURE CBPictQtde() valid nQtde > 0 when VTLastKey() == 5
			@ ++nLin,0 VTSAY STR0005 //"Produto"
			@ ++nLin,0 VTGET cProduto    PICTURE "@!" valid VldProduto("")
		Else
			VTClear()
			@ 0,0 VTSAY STR0006 //"Quantidade"
			@ 1,0 VTGET nQtde PICTURE CBPictQtde() valid nQtde > 0
			VTRead()
			VTClear()
			@ 0,0 VTSAY STR0005 //"Produto"
			@ 1,0 VTGET cProduto    PICTURE "@!" valid VldProduto("")
			VTRead()
			VTClear()
		EndIf
	EndIf
EndIf
Return

//ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
Static Function VldProduto(cTipo)
Local cTipId      := ""
Local aEtiqueta   := {}
Local cEtiqueta   := cProduto
Local nQE         := 0
Local nQtdeLida   := 0
Local nSaldo      := 0
Local aListaBKP   := {}
Local aHisEtiBKP  := {}
Local aItensPallet:= {}
Local lIsPallet   := .t.
Local nP          := 0
Local nTamLote    := TamSX3("B8_LOTECTL")[1]
Local nTamSLote   := TamSX3("B8_LOTECTL")[1]
Local nTamSeri    := TamSX3("BF_NUMSERI")[1]
Local lEstNeg	  := SuperGetMV('MV_ESTNEG')=='N'
Local lAV151VPR   := ExistBlock("AV151VPR")

If "01" $ cTipo
	If Empty(cCB0Prod)
		Return .t.
	EndIf

	aItensPallet := CBItPallet(cCB0Prod)
	lIsPallet := .t.
	If len(aItensPallet) == 0
		aItensPallet:={cCB0Prod}
		lIsPallet := .f.
	EndIf
	cTipId:=CBRetTipo(cCB0Prod)
	If cTipId == "01" .and. cTipId $ cTipo .or. lIsPallet
		aListaBKP := aClone(aLista)
		aHisEtiBKP:= aClone(aHisEti)

		Begin Sequence
			For nP:= 1 to len(aItensPallet)
				cCB0Prod :=  padr(aItensPallet[nP],20)
				aEtiqueta:= CBRetEti(cCB0Prod,"01")
				If Empty(aEtiqueta)
					VTALERT(STR0007,STR0008 ,.T.,4000,3) //"Etiqueta invalida"###"Aviso"
					Break
				EndIf
				If ! lIsPallet .and. ! Empty(CB0->CB0_PALLET)
					VTALERT(STR0009,STR0008,.T.,4000,2) //"Etiqueta invalida, Produto pertence a um Pallet"###"Aviso"
					Break
				EndIf
				//--Valida se a etiqueta já foi consumida por outro processo
				If CB0->CB0_STATUS $ "123"  
					VTBeep(2)
					VTAlert(STR0007,STR0008,.T.,4000,3) //"Etiqueta invalida"###"Aviso"
					VTKeyBoard(chr(20))
					Return .f.
				EndIf				
				If ! Empty(aEtiqueta[2]) .and. Ascan(aHisEti,cCB0Prod) > 0
					VTALERT(STR0010,STR0008 ,.T.,4000,3) //"Etiqueta ja lida"###"Aviso"
					Break
				EndIf
				If aEtiqueta[10] == GetMvNNR('MV_CQ','98')
					VtAlert(STR0042, STR0017,.t.,4000,4) //"Esta rotina nao trata armazem de CQ!"###'Aviso'
					VtAlert(STR0043, STR0017,.t.,4000,4) //"Utilize as rotinas de Envio/Baixa CQ!"###'Aviso'
					Break
				EndIf
				If Localiza(aEtiqueta[1])
					VTALERT(STR0011,STR0008,.T.,4000,3) //"Produto lido controla endereco!"###"Aviso"
					VTALERT(STR0012,STR0008,.T.,4000) //"Utilize rotina especifica ACDV150"###"Aviso"
					Break
				EndIf
				If !Empty(aEtiqueta[13])
					VTALERT(STR0044,STR0008,.T.,4000,3) //"Etiqueta utilizada em NF saida."###"Aviso"
					Break
				EndIf
				If Empty(aEtiqueta[2])
					aEtiqueta[2]:= 1
				EndIf
				cArmOri := aEtiqueta[10]
				cLote   := aEtiqueta[16]
				cSLote  := aEtiqueta[17]
				cNumSerie:=CB0->CB0_NUMSER
				If ! CBProdLib(cArmOri,aEtiqueta[1])
					Break
				EndIf
				cProduto  := aEtiqueta[1]				
				nQE:= 1
				If ! CBProdUnit(aEtiqueta[1])
					nQE := CBQtdEmb(aEtiqueta[1])
					If empty(nQE)
						Break
					EndIf
				EndIf
				nQtdeProd := aEtiqueta[2]*nQE
				//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
				//³ Consiste Quantidade Negativa                           ³
				//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
				If lEstNeg
					nQtdeLida := 0
					aEval(aLista,{|x|  If(x[1]+x[3]+x[4]+x[5]+x[6]==cProduto+cArmOri+cLote+cSLote+cNumSerie,nQtdeLida+=x[2],nil)})
					//--> Verifica se o saldo do armz esta liberado
					SB2->(DbSetOrder(1))
					SB2->(DbSeek(xFilial("SB2")+cProduto+cArmOri))
					nSaldo := SaldoMov()
					If nQtdeProd+nQtdeLida >  nSaldo
						VTALERT(STR0013,STR0008 ,.T.,4000,3) //"Quantidade excede o saldo disponivel"###"Aviso"
						Break
					EndIf
				EndIf
				If lAV151VPR
					If ! ExecBlock("AV151VPR",.F.,.F.,cEtiqueta)
						Break
					EndIf
				EndIf
				TrataArray(cCB0Prod)
			Next
			VTKeyboard(chr(20))
			Return .f.
		End Sequence
		aLista := aClone(aListaBKP)
		aHisEti:= aClone(aHisEtiBKP)
		VTKeyboard(chr(20))
		Return .f.
	Else
		VTALERT(STR0007,STR0008 ,.T.,4000,3) //"Etiqueta invalida"###"Aviso"
		VTKeyboard(chr(20))
		Return .f.
	EndIf
Else
	If Empty(cProduto)
		Return .t.
	EndIf
	If ! CBLoad128(@cProduto)
		VTKeyboard(chr(20))
		Return .f.
	EndIf
	cTipId:=CBRetTipo(cProduto)
	If ! cTipId $ "EAN8OU13-EAN14-EAN128"
		VTALERT(STR0014,STR0008,.T.,4000,3) //"Etiqueta invalida."###"Aviso"
		VTKeyboard(chr(20))
		Return .f.
	EndIf
	aEtiqueta := CBRetEtiEAN(cProduto)
	If Empty(aEtiqueta) .or. Empty(aEtiqueta[2])
		VTALERT(STR0014,STR0008,.T.,4000,3) //"Etiqueta invalida."###"Aviso"
		VTKeyboard(chr(20))
		Return .f.
	EndIf
	If ! CBProdLib(cArmOri,aEtiqueta[1])
		VTKeyBoard(chr(20))
		Return .f.
	EndIf
	nQE:= 1
	If ! CBProdUnit(aEtiqueta[1])
		nQE := CBQtdEmb(aEtiqueta[1])
		If empty(nQE)
			VTKeyboard(chr(20))
			Return .f.
		EndIf
	EndIf
	cLote := aEtiqueta[3]
	If ! CBRastro(aEtiqueta[1],@cLote,@cSLote)
		VTKeyboard(chr(20))
		Return .f.
	EndIf
	If Localiza(aEtiqueta[1])
		VTALERT(STR0011,STR0008,.T.,4000,3) //"Produto lido controla endereco!"###"Aviso"
		VTALERT(STR0012,STR0008,.T.,4000) //"Utilize rotina especifica ACDV150"###"Aviso"
		VTKeyboard(chr(20))
		cLote     := Space(nTamLote)
		cSLote    := Space(nTamSLote)
		cNumSerie := Space(nTamSeri)
		Return .f.
	EndIf
	cProduto  := aEtiqueta[1]
	nQtdeProd := aEtiqueta[2]*nQtde*nQE
	If Len(aEtiqueta) >= 5
		cNumSerie:=Padr(aEtiqueta[5],Len(Space(nTamSeri)))
	EndIf
	If SuperGetMV('MV_ESTNEG')=='N'
		nQtdeLida := 0
		aEval(aLista,{|x|  If(x[1]+x[3]+x[4]+x[5]+x[6]==cProduto+cArmOri+cLote+cSLote+cNumSerie,nQtdeLida+=x[2],nil)})
		SB2->(DbSetOrder(1))
		SB2->(DbSeek(xFilial("SB2")+cProduto+cArmOri))
		nSaldo := SaldoMov()
		If nQtdeProd+nQtdeLida > nSaldo
			VTALERT(STR0013,STR0008 ,.T.,4000,3) //"Quantidade excede o saldo disponivel"###"Aviso"
			cLote     := Space(nTamLote)
			cSLote    := Space(nTamSLote)
			cNumSerie := Space(nTamSeri)
			VTKeyboard(chr(20))
			Return .f.
		EndIf
	EndIf
	If lAV151VPR
		If ! ExecBlock("AV151VPR",.F.,.F.,cEtiqueta)
			VTKeyboard(chr(20))
			Return .f.
		EndIf
	EndIf
	TrataArray(Nil)
	nQtde := 1
	VTGetRefresh('nQtde')
	VTKeyboard(chr(20))
	cLote     := Space(nTamLote)
	cSLote    := Space(nTamSLote)
	cNumSerie := Space(nTamSeri)
	Return .F.
EndIf
Return .t.

//ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
Static Function GetArmDes()
Local cPictNNR := PesqPict( "NNR","NNR_CODIGO")

IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf
If lVT100B /* GetMv("MV_RF4X20") */ .and. UsaCB0('01') 
	VTClear(1,0,3,19)
	nLin := 1
	@ ++nLin,0 VtSay STR0015 //'Armazem destino'
	@ ++nLin,0 VTGet cArmDes pict cPictNNR Valid VTLastkey() == 5 .or. (! Empty(cArmDes) .and. VldEndDes()) When !Empty(aLista) ;
		.and. iif(vtRow() == 3 .and. vtLastKey() == 5,(VTKeyBoard(chr(27)),lVolta := .T.),.T.)
Else
	If VtModelo()=="RF"
		@ ++nLin,0 VtSay STR0015 //'Armazem destino'
		@ ++nLin,0 VTGet cArmDes pict cPictNNR Valid VTLastkey() == 5 .or. (! Empty(cArmDes) .and. VldEndDes()) When !Empty(aLista)
	else
		@ 0,0 VtSay STR0015 //'Armazem destino'
		@ 1,0 VTGet cArmDes pict cPictNNR Valid VTLastkey() == 5 .or. (! Empty(cArmDes) .and. VldEndDes()) When !Empty(aLista)
	EndIf
EndIf
Return

//ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
Static Function VldEndDes()
Local nI
Local lRet := .T.
If Empty(cArmDes)
	VtAlert(STR0016, STR0017,.t.,4000,4) //'Armazem invalido'###'Aviso'
	VTKeyboard(chr(20))
	VTClearGet("cArmDes")
	VTGetSetFocus("cArmDes")
	Return .f.
EndIf
If Trim(cArmDes) == GetMvNNR('MV_CQ','98')
	VtAlert(STR0042, STR0017,.t.,4000,4) //"Esta rotina nao trata armazem de CQ!"###'Aviso'
	VtAlert(STR0043, STR0017,.t.,4000,4) //"Utilize as rotinas de Envio/Baixa CQ!"###'Aviso'
	VTKeyboard(chr(20))
	Return .f.
EndIf
For nI := 1 to Len(aLista)
	If ! CBProdLib(cArmDes,aLista[nI,1],.f.)
		VTALERT(STR0018+aLista[nI,1]+STR0019+cArmDes,STR0017,.t.,4000,2) //'Produto '###' bloqueado para inventario no armazem '###'Aviso'
		VTKeyboard(chr(20))
		If ! UsaCb0("02")
			VTClearGet("cArmDes")
			VTGetSetFocus("cArmDes")
		EndIf
		Return .f.
	EndIf
	SB2->(DbSetOrder(1))
	If !SB2->(MsSeek(xFilial('SB2')+aLista[nI,1]+cArmDes,.F.)) .And. SuperGetMV('MV_VLDALMO',.F.,'S') == 'S'
		VtAlert(STR0045+AllTrim(aLista[nI,1])+".")
		VTClearGet("cArmDes")
		VTGetSetFocus("cArmDes")
		Return .F.
	EndIf
Next
If Ascan(aLista,{|x| x[3] ==cArmDes}) > 0
	VTALERT(STR0020,STR0008 ,.T.,4000,3) //"Armazem de origem igual ao destino"###"Aviso"
	VTKeyboard(chr(20))
	If ! UsaCb0("02")
		VTClearGet("cArmDes")
		VTGetSetFocus("cArmDes")
	EndIf
	Return .f.
EndIf
// Ponto de Entrada para validar Armazem Destino
If ExistBlock("AV151VLD")
	lRet := ExecBlock("AV151VLD",.F.,.F.,{2,cArmDes})
EndIf

If lRet .And. !GravaTransf()
	lRet := .F.
EndIf

Return lRet


//ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
Static Function TrataArray(cEtiqueta,lEstorno)
Local nPos		:= 0
Local nAchou	:= 0
Default lEstorno := .f.
If ! lEstorno
	If cEtiqueta <> NIL
		aadd(aHisEti,cEtiqueta)
	EndIf	
	nPos := aScan(aLista,{|x| x[1]+x[3]+x[4]+x[5]+x[6] == cProduto+cArmOri+cLote+cSLote+cNumSerie})
	If Empty(nPos)
		aadd(aLista,{cProduto,nQtdeProd,cArmOri,cLote,cSLote,cNumSerie,NIL,NIL,NIL,{{cEtiqueta,nQtdeProd}}})
	Else
		aLista[nPos,2]+=nQtdeProd
		If (nAchou := aScan(aTail(aLista[nPos]),{|x| x[1] == cEtiqueta})) == 0	//-- Etiqueta lida anteriormente (qtd variavel)
			aAdd(aLista[nPos,10],{cEtiqueta,nQtdeProd})	//-- Adiciona
		Else
			aLista[nPos,10,nAchou,2] += nQtdeProd	//-- Adiciona
		EndIf
	EndIf
Else
	nPos := aScan(aLista,{|x| x[1]+x[3]+x[4]+x[5]+x[6] == cProduto+cArmOri+cLote+cSLote+cNumSerie})
	aLista[nPos,2] -= nQtdeProd
	If Empty(aLista[nPos,2])	//-- Se zerou quantidade do item a transferir, remove
		aDel(aLista,nPos)
		aSize(aLista,len(aLista)-1)
	ElseIf cEtiqueta <> NIL .And. (nAchou := aScan(aLista[nPos,10],{|x| x[1] == cEtiqueta})) > 0	//-- Senão, remove etiqueta
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

//ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
Static Function GravaTransf()
Local aSave
Local nI
Local aTransf:={}
Local cDoc     := ""
Local dValid
Local nTamEnd   := TamSX3("BF_LOCALIZ")[1]
Local nTamLoc   := TamSX3("B2_LOCAL") [1]
Local nTamLote  := TamSX3("B8_LOTECTL")[1]
Local nTamSLote := TamSX3("B8_LOTECTL")[1]
Local nTamSeri  := TamSX3("BF_NUMSERI")[1]
Local lV151AUTO := ExistBlock("V151AUTO")
Private nModulo := 4

If ! VTYesNo(STR0021,STR0008 ,.T.) //"Confirma transferencia?"###"Aviso"
	VTKeyboard(chr(20))
	Return .f.
EndIf

aSave     := VTSAVE()
VTClear()
VTMsg(STR0022) //'Aguarde...'
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³Ponto de entrada: Informar numero do documento.            ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
If ExistBlock('ACD151DC')
	cDoc := ExecBlock('ACD151DC',.F.,.F.)
	cDoc := If(ValType(cDoc)=="C",cDoc,"")
Endif
Begin Transaction
	lMsErroAuto := .F.
	lMsHelpAuto := .T.
	aTransf:=Array(len(aLista)+1)
	aTransf[1] := {cDoc,dDataBase}
	For nI := 1 to Len(aLista)
		SB1->(dbSeek(xFilial("SB1")+aLista[nI,1]))
		dValid := dDatabase+SB1->B1_PRVALID
		If Rastro(aLista[nI,1])
			SB8->(DbSetOrder(3))
			SB8->(DbSeek(xFilial("SB8")+aLista[nI,1]+aLista[nI,3]+aLista[nI,4]+AllTrim(aLista[nI,5])))
			dValid := SB8->B8_DTVALID
		EndIf
		aTransf[nI+1]:=  {{"D3_COD" , SB1->B1_COD				,NIL}}
		aAdd(aTransf[nI+1],{"D3_DESCRI" , SB1->B1_DESC				,NIL})
		aAdd(aTransf[nI+1],{"D3_UM"     , SB1->B1_UM				,NIL})
		aAdd(aTransf[nI+1],{"D3_LOCAL"  , aLista[nI,3]				,NIL})
		aAdd(aTransf[nI+1],{"D3_LOCALIZ", Space(nTamEnd)				,NIL})
		aAdd(aTransf[nI+1],{"D3_COD"    , SB1->B1_COD				,NIL})
		aAdd(aTransf[nI+1],{"D3_DESCRI" , SB1->B1_DESC				,NIL})
		aAdd(aTransf[nI+1],{"D3_UM"     , SB1->B1_UM				,NIL})
		aAdd(aTransf[nI+1],{"D3_LOCAL"  , cArmDes					,NIL})
		aAdd(aTransf[nI+1],{"D3_LOCALIZ", Space(nTamEnd)				,NIL})
		aAdd(aTransf[nI+1],{"D3_NUMSERI", aLista[nI,6]				,NIL})//numserie
		aAdd(aTransf[nI+1],{"D3_LOTECTL", aLista[nI,4]				,NIL})//lote
		aAdd(aTransf[nI+1],{"D3_NUMLOTE", aLista[nI,5]				,NIL})//sublote
		aAdd(aTransf[nI+1],{"D3_DTVALID", dValid					,NIL})
		aAdd(aTransf[nI+1],{"D3_POTENCI", criavar("D3_POTENCI")	,NIL})
		aAdd(aTransf[nI+1],{"D3_QUANT"  , aLista[nI,2]				,NIL})
		aAdd(aTransf[nI+1],{"D3_QTSEGUM", criavar("D3_QTSEGUM")	,NIL})
		aAdd(aTransf[nI+1],{"D3_ESTORNO", criavar("D3_ESTORNO")	,NIL})
		aAdd(aTransf[nI+1],{"D3_NUMSEQ" , criavar("D3_NUMSEQ")		,NIL})
		aAdd(aTransf[nI+1],{"D3_LOTECTL", aLista[nI,4]				,NIL})
		aAdd(aTransf[nI+1],{"D3_DTVALID", dValid					,NIL})
		/*Ponto de entrada, permite manipular e ou acrescentar dados no array aTransf.*/
		If lV151AUTO	
			aPEAux := aClone(aTransf)  
			aPEAux := ExecBlock("V151AUTO",.F.,.F.,{aTransf})
			If ValType(aPEAux)=="A" 
				aTransf := aClone(aPEAux)
			EndIf
		EndIf								
		If ! UsaCB0("01")
			CBLog("02",{SB1->B1_COD,aLista[nI,2],aLista[nI,4],aLista[nI,5],aLista[nI,3],,cArmDes})
		EndIf
	Next
	MSExecAuto({|x| MATA261(x)},aTransf)
	If lMsErroAuto
		VTALERT(STR0023,STR0024,.T.,4000,3) //"Falha na gravacao da transferencia"###"ERRO"
		DisarmTransaction()
		Break
	Else
		If ExistBlock("ACD151GR")
			ExecBlock("ACD151GR",.F.,.F.)
		EndIf
	EndIf
End Transaction
VtRestore(,,,,aSave)
If lMsErroAuto
	VTDispFile(NomeAutoLog(),.t.)
Else
	If ExistBlock("ACD151OK")
		ExecBlock("ACD151OK",.F.,.F.)
	EndIf
	cArmOri     := Space(nTamLoc)
	cEndOri     := Space(nTamEnd)
	cCB0ArmOri  := Space(20)
	cArmDes     := Space(nTamLoc)
	cCB0Prod    := Space(20)
	cProduto    := IIf( FindFunction( 'AcdGTamETQ' ), AcdGTamETQ(), Space(48) )
	cLote       := Space(nTamLote)
	cSLote      := Space(nTamSLote)
	cNumSerie   := Space(nTamSeri)
	nQtde       := 1
	aLista      := {}
	aHisEti     := {}
EndIf
Return .t.

//ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
Static Function Informa()
Local aCab  := {STR0005,STR0006,STR0027,STR0028,STR0029,STR0030} //"Produto"###"Quantidade"###"Armazem"###"Lote"###"SubLote"###"Num.Serie"
Local aSize := {15,16,7,10,7,20}
Local aSave := VTSAVE()
IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf
VtClear()
VTaBrowse(0,0,Iif(lVT100B /*GetMv("MV_RF4X20")*/ ,3,7),19,aCab,aLista,aSize)
VtRestore(,,,,aSave)
Return

//ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
Static Function Estorna()
Local aTela
Local cEtiqueta
Local nQtde := 1
aTela := VTSave()
VTClear()
cEtiqueta := IIf( FindFunction( 'AcdGTamETQ' ), AcdGTamETQ(), Space(48) )
If VtModelo()=="RF"
	@ 00,00 VtSay Padc(STR0031,VTMaxCol()) //"Estorno da Leitura"
EndIf
If ! UsaCB0('01')
	@ 1,00 VTSAY  STR0032 VTGet nQtde   pict CBPictQtde() when VTLastkey() == 5 //'Qtde.'
EndIf
If VtModelo()=="RF"
	@ 02,00 VtSay STR0033 //"Etiqueta:"
	@ 03,00 VtGet cEtiqueta pict "@!" Valid VldEstorno(cEtiqueta,@nQtde)
Else
	VtClear()
	@ 0,00 VtSay STR0033 //"Etiqueta:"
	@ 1,00 VtGet cEtiqueta pict "@!" Valid VldEstorno(cEtiqueta,@nQtde)
Endif
VtRead
vtRestore(,,,,aTela)
Return

//ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
Static Function VldEstorno(cEtiqueta,nQtde)
Local nPos
Local aEtiqueta,nQE
Local aListaBKP := aClone(aLista)
Local aHisEtiBKP:= aClone(aHisEti)
Local aItensPallet := CBItPallet(cEtiqueta)
Local lIsPallet := .t.
Local nP
Local nTamSeri := TamSX3("BF_NUMSERI")[1]

If Empty(cEtiqueta)
	Return .f.
EndIf

If len(aItensPallet) == 0
	aItensPallet:={cEtiqueta}
	lIsPallet := .f.
EndIf

Begin Sequence
	For nP:= 1 to len(aItensPallet)
		cEtiqueta:=padr(aItensPallet[nP], IIf( FindFunction( 'CBGetTamEtq' ), CBGetTamEtq(), 48 ) )
		If UsaCB0("01")
			nPos := Ascan(aHisEti, {|x| AllTrim(x) == AllTrim(cEtiqueta)})
			If nPos == 0
				VTALERT(STR0034,STR0008,.T.,4000,2) //"Etiqueta nao encontrada"###"Aviso"
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
			nQtde	   := 1

			If ! lIsPallet .and. ! Empty(CB0->CB0_PALLET)
				VTALERT(STR0009,STR0008,.T.,4000,2) //"Etiqueta invalida, Produto pertence a um Pallet"###"Aviso"
				Break
			EndIf
		Else
			If ! CBLoad128(@cEtiqueta)
				Return .f.
			EndIf
			aEtiqueta := CBRetEtiEAN(cEtiqueta)
			If Len(aEtiqueta) == 0
				VTALERT(STR0007,STR0008,.T.,4000,2) //"Etiqueta invalida"###"Aviso"
				VTKeyboard(chr(20))
				Return .f.
			EndIf
			cProduto := aEtiqueta[1]
			If ascan(aLista,{|x| x[1] ==cProduto}) == 0
				VTALERT(STR0035,STR0008,.T.,4000,2) //"Produto nao encontrado"###"Aviso"
				VTKeyboard(chr(20))
				Return .f.
			EndIf
			cLote := aEtiqueta[3]
			If len(aEtiqueta) >=5
				cNumSerie:= padr(aEtiqueta[5],Len(Space(nTamSeri)))
			EndIf
		EndIf

		nQE := 1
		If ! CBProdUnit(cProduto)
			nQE := CBQtdEmb(cProduto)
			If Empty(nQE)
				Break
			EndIf
		EndIf
		nQtdeProd:=nQtde*nQE*aEtiqueta[2]

		If ! Usacb0("01") .and. ! CBRastro(cProduto,@cLote,@cSLote)
			Break
		EndIf

		nPos := Ascan(aLista,{|x| x[1]+x[3]+x[4]+x[5]+x[6] == cProduto+cArmOri+cLote+cSLote+cNumSerie})
		If nPos == 0
			VTALERT(STR0036,STR0008,.T.,4000,2) //"Produto nao encontrado neste armazem"###"Aviso"
			Break
		EndIf
		If aLista[nPos,2] < nQtdeProd
			VTALERT(STR0037,STR0008,.T.,4000,2) //"Quantidade excede o estorno"###"Aviso"
			Break
		EndIf
		If UsaCB0("01")
			TrataArray(cEtiqueta,.t.)
		Else
			TrataArray(,.t.)
		EndIf
	Next
	If ! VTYesNo(STR0038,STR0039,.t.) //"Confirma o estorno?"###"ATENCAO"
		Break
	EndIf
	nQtde:= 1
	VTGetRefresh("nQtdePro")
	VTKeyboard(chr(20))
	Return .f.
End Sequence
aLista := aClone(aListaBKP)
aHisEti:= aClone(aHisEtiBKP)
nQtde  := 1
VTGetRefresh("nQtdePro")
VTKeyBoard(chr(20))
Return .f.
