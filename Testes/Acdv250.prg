#INCLUDE "Acdv250.ch" 
#include "protheus.ch"
#include "apvt100.ch"

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ ACDV250  ³ Autor ³ Henrique Gomes Oikawa  		     ³ Data ³ 31/03/04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Embarque Simples (R.F.)              						  		   ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Uso       ³ Rotina de embarque simples sobre os itens da Nota Fiscal de Saida	   ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/ 

Function ACDV250

Private cNota
Private cSerie
Private cCodOpe := CBRetOpe()
Private lUsaCB0 := UsaCB0('01')
Private lSLoteCBL := CBL->(FieldPos("CBL_SLOTE")) > 0

If Empty(cCodOpe)
	VTAlert(STR0001,STR0002,.T.,4000) //"Operador nao cadastrado"###"Aviso"
	Return .F.
EndIf   

While .T.
	cNota  := Space(TamSx3("F2_DOC")[1])
	cSerie := Space(SerieNfId("SF2",6,"F2_SERIE"))
	VTClear()
	@ 0,00 VTSAY STR0003 //'Embarque'
	@ 1,00 VTSAY STR0004 VTGet cNota pict '@!' Valid VldNota(cNota) F3 "CBK" //'Nota :'
	@ 2,00 VTSAY STR0005 VTGet cSerie pict '!!!' Valid VldNota(cNota,@cSerie,.T.) .or. VtLastkey()==5 //'Serie:'
	VTRead
   	If VtLastKey() == 27                    
		Exit
	EndIf
 	Embarque()
EndDo  

Return 

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±00±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³ VldNota  º Autor ³Henrique Gomes Oikawaº Data ³  31/03/04   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Valida a nota fiscal de saida                               º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ ACD                                                         º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
Static Function VldNota(cNota,cSerie,lSerie)

Default lSerie := .F.
Default cNota := ""
Default cSerie := ""

If Empty(cNota+cSerie)
	VTKeyBoard(chr(23))
	Return .F.
EndIf

If lSerie
	CBMULTDOC("SF2",cNota,@cSerie)
EndIf

SF2->(dbSetOrder(1))
If ! SF2->(dbSeek(xFilial()+cNota+cSerie))
	VTAlert(STR0006,STR0002,.T.,4000,2) //"Nota fiscal nao cadastrada"###"Aviso"
	VTKeyBoard(chr(20))
	Return .F.
EndIf
If Len(cNota+cSerie) < 9
	Return .t.
Endif
CBK->(DbSetOrder(1))
If CBK->(DbSeek(xFilial('CBK')+cNota+cSerie)) .AND. (CBK->(CBK_CLIENT+CBK_LOJA+DTOS(CBK_EMISSA))<>SF2->(F2_CLIENTE+F2_LOJA+DTOS(F2_EMISSAO)))
	//Se a nota gravada na tabela de Conferencia Embarque 'CBK' diferente da Nota fiscal, exclui CBK/CBL
	CBV250Del(SF2->(F2_DOC+F2_SERIE+F2_CLIENTE+F2_LOJA))
Endif
If CBK->(Eof())
	CBV250Grv(SF2->(F2_DOC+F2_SERIE+F2_CLIENTE+F2_LOJA))
Endif

If CBK->CBK_STATUS == "2"
	IF ! VTYesNo("Embarque finalizado,Confirma reabertura da conferência?",STR0002,.T.,4000) //'Embarque finalizado, "Confirma reabertura da conferência?"'###'Atencao'
		VtClearGet("cNota")  // Limpa o get
		VtClearGet("cSerie")  // Limpa o get
		VTGetSetFocus("cNota")
		Return .F.
	EndIf
	CBV250Del(SF2->(F2_DOC+F2_SERIE+F2_CLIENTE+F2_LOJA))
	CBV250Grv(SF2->(F2_DOC+F2_SERIE+F2_CLIENTE+F2_LOJA))
Endif

Return .t.


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³ CBV250Del  ºAutor  ³ TOTVS               º Data ³  31/03/04   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Exclui registros das tabelas CBK/CBL                          º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
Function CBV250Del(cChaveSF2)

CBK->(DbSetOrder(1))
CBK->(DbSeek(xFilial("CBK")+cChaveSF2))
While CBK->(!Eof() .AND. CBK_FILIAL+CBK_DOC+CBK_SERIE+CBK_CLIENT+CBK_LOJA == xFilial("CBK")+cChaveSF2)
	RecLock("CBK",.F.)
	CBK->(DbDelete())
	CBK->(MsUnlock())
	CBK->(DbSkip())
Enddo

CBL->(DbSetOrder(2))
CBL->(DbSeek(xFilial("CBL")+cChaveSF2))
While CBL->(!Eof() .AND. CBL_FILIAL+CBL_DOC+CBL_SERIE+CBL_CLIENT+CBL_LOJA == xFilial("CBL")+cChaveSF2)
	RecLock("CBL",.F.)
	CBL->(DbDelete())
	CBL->(MsUnlock())
	CBL->(DbSkip())
Enddo

Return

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³ CBV250Grv  ºAutor  ³ TOTVS               º Data ³  31/03/04   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Carrega os itens das tabelas CBK / CBL           			 º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
Function CBV250Grv(cChaveSF2)
Local aAreaSF2 := SF2->(GetArea())

SF2->(DbSetOrder(1))
SF2->(DbSeek(xFilial("SF2")+cChaveSF2))

//Grava cabecalho da Conferencia de Embarque:
RecLock("CBK",.T.)
CBK->CBK_FILIAL := xFilial('CBK')
CBK->CBK_DOC    := SF2->F2_DOC
//CBK->CBK_SERIE  := SF2->F2_SERIE
SerieNfId("CBK",1,"CBK_SERIE",,,,SF2->F2_SERIE)
CBK->CBK_CLIENT := SF2->F2_CLIENTE
CBK->CBK_LOJA   := SF2->F2_LOJA
CBK->CBK_EMISSA := SF2->F2_EMISSAO
CBK->CBK_STATUS := "0" //Embarque em andamento
CBK->(MsUnlock())

RestArea(aAreaSF2)
Return

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³ Embarque ºAutor  ³Henrique Gomes Oikawaº Data ³  31/03/04   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Faz o embarque dos itens da nota fiscal         			   º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
Static Function Embarque()
Local   bkey06
Local   bKey09
Local   bkey24
Local   cEtiqProd
Local   lUsaCB0 := UsaCB0('01')
Private lFimEmb := .f.
IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf

bKey06 := VTSetKey(06,{|| Faltas()},STR0009)    // CTRL+F //"Faltantes"
bkey09 := VTSetKey(09,{|| Informa()},STR0010) //"Informacoes"
bKey24 := VTSetKey(24,{|| Estorna()},STR0011)   // CTRL+X //"Estorno"
While .t.
	If lUsaCB0
		cEtiqProd := Space(TamSx3("CB0_CODET2")[1])
	Else
		cEtiqProd := IIf( FindFunction( 'AcdGTamETQ' ), AcdGTamETQ(), Space(48) )
	Endif
	If lVT100B // GetMv("MV_RF4X20")
		VTClear
		@ 1,00 VTSAY STR0012 //'Leia a etiqueta'
		@ 2,00 VTGET cEtiqProd pict '@!' Valid CBV250VEt(cEtiqProd)
	Else
		@ 4,00 VTSAY STR0012 //'Leia a etiqueta'
		@ 5,00 VTGET cEtiqProd pict '@!' Valid CBV250VEt(cEtiqProd)
	EndIf
	VTRead
	If lFimEmb
		Exit
	EndIf
	If VTLastkey() == 27
		If ! VTYesNo(STR0013,STR0008,.T.) //'Confirma a saida?'###'Atencao'
			Loop
		EndIf
		AtuCBK(.t.)
		Exit
	EndIf
EndDo
vtsetkey(06,bkey06)
vtsetkey(09,bkey09)
vtsetkey(24,bkey24)

Return



/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³ CBV250VEt ºAutor  ³Henrique Gomes Oikawaº Data ³  31/03/04   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Valida a etiqueta de produto								    º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
Function CBV250VEt(cEtiqProd,lEstorna,nQtdEmbMonit,oGetEtiq)
Local   cTipo
Local   cProduto     := Space(TamSX3("CBL_PROD")[01])
Local   cLote        := Space(TamSX3("CBL_LOTECT")[01])
Local   cSLote       := Space(IIF(lSLoteCBL,TamSX3("CBL_SLOTE")[01],TamSX3("D2_NUMLOTE")[01]))
Local   cNumSerie    := Space(TamSX3("CBL_NUMSER")[01])
Local   nQE          := 1
Local   aEtiqueta    := {}
Local   lUsaCB0      := UsaCB0('01')
Local   lRet         := .f.
Default lEstorna     := .f.
Default nQtdEmbMonit := 1

If Empty(cEtiqProd)
	If IsTelNet()
		Return .f.
	Else
		Return .t.
	Endif
EndIf

cTipo  := CBRetTipo(cEtiqProd)
If cTipo == "01"  // --> Etiqueta de Produto  com CB0
	aEtiqueta:= CBRetEti(cEtiqProd,"01")
	If Empty(aEtiqueta)
		CBAlert(STR0014,STR0002,.t.,4000,3) //"Etiqueta invalida"###"Aviso"
		If IsTelNet()
			VtClearGet("cEtiqProd")  // Limpa o get
		Endif
		Return .f.
	EndIf
	If !Empty(CB0->CB0_PALLET)
	   	CBALERT(STR0015,STR0002,.T.,4000,2) //"Etiqueta invalida, Produto pertence a um Pallet"###"AVISO"
		If IsTelNet()
			VtClearGet("cEtiqProd")  // Limpa o get
		Endif
		Return .f.
	Endif   
	CBL->(DbSetOrder(1))
	If CBL->(DbSeek(xFilial('CBL')+CBK->(CBK_DOC+CBK_SERIE)+AllTrim(cEtiqProd))) .and. !lEstorna
	   	CBALERT(STR0016,STR0002,.T.,4000,2) //"Etiqueta ja lida!"###"AVISO"
		If IsTelNet()
			VtClearGet("cEtiqProd")  // Limpa o get
		Endif
	   	Return .f.
	Endif
	cProduto  := aEtiqueta[01]
	nQE       := aEtiqueta[02] * nQtdEmbMonit
	cLote     := aEtiqueta[16]
	cSLote    := aEtiqueta[17]
	cNumSerie := aEtiqueta[23]
	If ! CBProdUnit(cProduto)
		nQE := CBQtdEmb(cProduto)
	EndIf
	If Empty(nQE)
		CBAlert(STR0017,STR0002,.t.,4000,3) //"Quantidade invalida!"###"Aviso"
		If IsTelNet()
			VtClearGet("cEtiqProd")  // Limpa o get
		Endif
		Return .f.
	EndIf      
Elseif cTipo $ "EAN8OU13-EAN14-EAN128" // --> Etiqueta EAN
	aEtiqueta := CBRetEtiEan(cEtiqProd)
	If Empty(aEtiqueta)
		CBAlert(STR0018,STR0002,.T.,4000,2) //"Etiqueta invalida."###"Aviso"
		If IsTelNet()
			VtClearBuffer()
			VTkeyBoard(chr(20))
		Endif
	 	Return .f.
	EndIf
	cProduto  := aEtiqueta[01]
	nQE 	  := If(aEtiqueta[2] == 0,1,aEtiqueta[2]) * nQtdEmbMonit
	cLote     := aEtiqueta[03]
	cNumSerie := aEtiqueta[05]
	If ! CBProdUnit(cProduto)
		nQE := CBQtdEmb(cProduto)
	EndIf
	If !CBRastro(cProduto,@cLote,@cSLote)
		Return .f.
	EndIf	
	If Empty(nQE)
		CBAlert(STR0017,STR0002,.t.,4000,3) //"Quantidade invalida!"###"Aviso"
		If IsTelNet()
			VtClearBuffer()
			VTkeyBoard(chr(20))
		Endif
		Return .f.
	EndIf      
Else
	CBAlert(STR0014,STR0002,.t.,4000,3) //"Etiqueta invalida"###"Aviso"
	If IsTelNet()
		VtKeyboard(Chr(20))  // zera o get
	Endif
	Return .F.
Endif

If !VldProd(cProduto,nQE,cLote,cNumSerie,If(lUsaCB0,CB0->CB0_CODETI,Space(TamSX3("CB0_CODETI")[1])),lEstorna,cSLote)
 	Return .f.
EndIf

If IsTelNet()
	VtClearBuffer()
	VTKeyBoard(chr(20))
	lRet := .t.
Else
	lRet := .f.
Endif
AtuCBK()  //Atualiza status de conferencia embarque
If IsInCallStack("ACDA150")
	ACDA150Sts(.t.)
	oGetEtiq:buffer := IIf( FindFunction( 'AcdGTamETQ' ), AcdGTamETQ(), Space(48) )
	cEtiq := IIf( FindFunction( 'AcdGTamETQ' ), AcdGTamETQ(), Space(48) )
Endif    

Return lRet

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³ VldProd  ºAutor  ³Henrique Gomes Oikawaº Data ³  31/03/04   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Valida o produto lido									   º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
Static Function VldProd(cProduto,nQE,cLote,cNumSerie,cEtiqProd,lEstorna,cSLote)
Local  aRetAnalise := {}
Local  nQtdeNota   := 0
Local  nQtdeEmb    := 0
Local  nQtdeNec    := 0
Local  cChave      := ""
Local  lRet		   := .T.	
Local  lSai		   := .F.
Local  lAchou      := .F.	
Default cSLote     := Space(IIF(lSLoteCBL,TamSX3("CBL_SLOTE")[01],TamSX3("D2_NUMLOTE")[01]))
CBL->(DbSetOrder(2))
If lEstorna                                            

	If lRet .And. !CBL->(DbSeek(xFilial("CBL")+CBK->(CBK_DOC+CBK_SERIE+CBK_CLIENT+CBK_LOJA)+cProduto+cLote+cNumSerie+cEtiqProd))
		CBAlert("Não existe saldo a coletar para estorno deste produto!","Aviso",.T.,4000,2)
		lSai:= CBYesNo("Deseja sair da tela de estorno?","Aviso")
		lRet:= .f.
	Endif       
	
	If lRet .And. CBL->CBL_QTDEMB < nQE
		CBAlert("A quantidade informada e superior a quantidade lida!","Aviso",.T.,4000,2)
		lRet := .f.
	Endif
	
	If lRet .And. !CBYesNo("Confirma estorno?","Aviso")
		lRet:= .f.
	Endif
	
	If lRet
		RecLock('CBL',.F.)
		If CBL->CBL_QTDEMB > nQE
			CBL->CBL_QTDEMB -= nQE
		Else
			CBL->(DbDelete())
		Endif
		CBL->(MsUnlock())
		
		AtuCBK()  //Atualiza status de conferencia embarque
	EndIf
	
	If IsTelNet()
		VtClearBuffer()
		VTkeyBoard(chr(20))
	Endif         
	
	Return lSai
Endif

aRetAnalise := AnalisaEmb(cProduto,cLote,cNumSerie,cSLote)
nQtdeNota 	:= aRetAnalise[01]
nQtdeEmb  	:= aRetAnalise[02]
nQtdeNec  	:= nQtdeNota - nQtdeEmb

If lRet .And. Empty(aRetAnalise[01])
	CBAlert(STR0019+Alltrim(cProduto)+STR0020,STR0002,.T.,4000,2) //"O produto "###" nao consta na nota!"###"Aviso"
	If IsTelNet()
		VtClearBuffer()
		VTkeyBoard(chr(20))
	Endif
  	lRet:= .f.
EndIf

If lRet .And. nQtdeNec <= 0
	CBAlert(STR0021,STR0002,.T.,4000,2) //"Nao existe saldo a coletar deste produto!"###"Aviso"
	If IsTelNet()
		VtClearBuffer()
		VTkeyBoard(chr(20))
	Endif
	lRet:= .f.
EndIf
	
If lRet .And. nQtdeNec < nQE
	CBAlert(STR0022,STR0002,.T.,4000,2) //"Quantidade maior que necessaria!"###"Aviso"
	If IsTelNet()
		VtClearBuffer()
		VTkeyBoard(chr(20))
	Endif
	lRet:=  .f.
EndIf

If lRet
	CBL->(dbSetOrder(2))
	CBL->(DbSeek(xFilial('CBL')+CBK->(CBK_DOC+CBK_SERIE+CBK_CLIENT+CBK_LOJA)))
	cChave := CBK->(CBK_FILIAL+CBK_DOC+CBK_SERIE+CBK_CLIENT+CBK_LOJA)
	While CBL->(CBL_FILIAL+CBL_DOC+CBL_SERIE+CBL_CLIENT+CBL_LOJA) == cChave
		// Atualiza a quantidade caso encontre Embarque para o Produto
		If lSLoteCBL
			If CBL->(CBL_FILIAL+CBL_DOC+CBL_SERIE+CBL_CLIENT+CBL_LOJA+CBL_PROD+CBL_LOTECT+CBL_SLOTE+CBL_NUMSER+CBL_CODETI) == cChave+cProduto+cLote+cSLote+cNumSerie+cEtiqProd
				RecLock('CBL',.F.)
				CBL->CBL_QTDEMB += nQE
				CBL->(MsUnlock())
				lAchou := .T.
				Exit
			EndIf
		Else
			If CBL->(CBL_FILIAL+CBL_DOC+CBL_SERIE+CBL_CLIENT+CBL_LOJA+CBL_PROD+CBL_LOTECT+CBL_NUMSER+CBL_CODETI) == cChave+cProduto+cLote+cNumSerie+cEtiqProd
				RecLock('CBL',.F.)
				CBL->CBL_QTDEMB += nQE
				CBL->(MsUnlock())
				lAchou := .T.
				Exit
			EndIf
		EndIf	
		CBL->(DbSkip())
	End
	// Caso nao encontre Embarque, gera um novo
	If !lAchou
		RecLock('CBL',.T.)
		CBL->CBL_FILIAL := xFilial('CBL')
		CBL->CBL_DOC    := CBK->CBK_DOC
		//CBL->CBL_SERIE  := CBK->CBK_SERIE
		SerieNfId("CBL",1,"CBL_SERIE",,,,CBK->CBK_SERIE)
		CBL->CBL_CLIENT := CBK->CBK_CLIENT
		CBL->CBL_LOJA   := CBK->CBK_LOJA
		CBL->CBL_PROD   := cProduto
		CBL->CBL_LOTECT := cLote
		If lSLoteCBL
			CBL->CBL_SLOTE := cSLote
		EndIf	
		CBL->CBL_NUMSER := cNumSerie
		CBL->CBL_CODETI := cEtiqProd
		CBL->CBL_QTDEMB += nQE
		CBL->(MsUnlock())
	EndIf
	AtuCBK()  //Atualiza status de conferencia embarque
EndIf

Return lRet

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o	 ³ Estorna	  ³ Autor ³Henrique Gomes Oikawa³ Data ³ 31/03/04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³ Estorno das etiquetas lidas                       		  ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
Static Function Estorna()
Local aTela	   := {}	
Local cEtiqueta                         
Local lUsaCB0  := UsaCB0('01')
Local lSai	   := .F.	

If Empty(cNota)
	Return
Endif

aTela := VTSave()
VTClear()                       
If lUsaCB0
   cEtiqueta := Space(TamSx3("CB0_CODET2")[1])
Else
   cEtiqueta := IIf( FindFunction( 'AcdGTamETQ' ), AcdGTamETQ(), Space(48) )
Endif
@ 00,00 VtSay Padc(STR0026,VTMaxCol()) //"Estorno da Leitura"
@ 02,00 VtSay STR0027 //"Etiqueta:"
@ 03,00 VtGet cEtiqueta pict "@!" Valid CBV250VEt(cEtiqueta,.t.)
VtRead                    

vtRestore(,,,,aTela)

Return 



/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o	 ³ Informa    ³ Autor ³Henrique Gomes Oikawa³ Data ³ 01/04/04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³ Mostra produtos que ja foram lidos                         ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
Static Function Informa()
Local aCab,aSize,aSave := VTSAVE()
Local aHist:={}
IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf

If Empty(cNota)
	Return
Endif

CBL->(dbSetOrder(2))
CBL->(DbSeek(xFilial('CBL')+CBK->(CBK_DOC+CBK_SERIE+CBK_CLIENT+CBK_LOJA)))
While CBL->(!Eof() .and. CBL_FILIAL+CBL_DOC+CBL_SERIE+CBL_CLIENT+CBL_LOJA == xFilial('SF2')+CBK->(CBK_DOC+CBK_SERIE+CBK_CLIENT+CBK_LOJA))
	If lSLoteCBL
		AADD(aHist,{CBL->CBL_CODETI,CBL->CBL_PROD,CBL->CBL_LOTECT,CBL->CBL_SLOTE,CBL->CBL_NUMSER,Transform(CBL->CBL_QTDEMB,"@E 999,999.99")})
	Else
		AADD(aHist,{CBL->CBL_CODETI,CBL->CBL_PROD,CBL->CBL_LOTECT,CBL->CBL_NUMSER,Transform(CBL->CBL_QTDEMB,"@E 999,999.99")})
	EndIf	
	CBL->(DbSkip())
EndDo
If lSLoteCBL
	aSort(aHist,,,{|x,y| x[2]+x[3]+x[4]+x[5]+x[1] < y[2]+y[3]+y[4]+y[5]+y[1] })
Else
	aSort(aHist,,,{|x,y| x[2]+x[3]+x[4]+x[1] < y[2]+y[3]+y[4]+y[1] })
EndIf

VTClear() 
@ 0,0 VTSay STR0030+cCodOpe //"Operador : "
aCab  := {STR0031,STR0028,"Lote","Num.Serie",STR0029} //"Etiqueta"###"Produto"###"Qtde"
aSize := {10,15,10,20,12}
If lVT100B // GetMv("MV_RF4X20")
	VTaBrowse(1,0,3,19,aCab,aHist,aSize)
Else
	VTaBrowse(1,0,7,19,aCab,aHist,aSize)
EndIf
VtRestore(,,,,aSave)
Return

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o	 ³ AtuCBK     ³ Autor ³Henrique Gomes Oikawa³ Data ³ 31/03/04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³ Atualiza o status da tabela de Cabecalho de Embarque (CBK) ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
Static Function AtuCBK(lFinal)
Local   aRetAnalise := {}
Local   nQtdeNota   := 0
Local   nQtdeEmb    := 0
Default lFinal      := .f.

aRetAnalise := AnalisaEmb()
nQtdeNota 	:= aRetAnalise[1]
nQtdeEmb  	:= aRetAnalise[2]

If nQtdeNota == nQtdeEmb
	lFimEmb := .t.
	CBAlert(STR0032,STR0002,.T.,4000) //"Embarque finalizado!"###"Aviso"
	RecLock("CBK",.F.)
	CBK->CBK_DTEMBQ := dDataBase
	CBK->CBK_STATUS := "2" //Embarque finalizado
	CBK->(MsUnlock())
Else
	RecLock("CBK",.F.)
	CBK->CBK_STATUS := "1" //Embarque em andamento
	CBK->(MsUnlock())
	If lFinal
		CBAlert(STR0033,STR0002,.T.,4000) //"Embarque em aberto!"###"Aviso"
	Endif
Endif
	
Return

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³ AnalisaEmb º Autor ³Henrique Gomes Oikawaº Data ³  31/03/04   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Analisa a necessidade de embarque do produto lido		     º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
Static Function AnalisaEmb(cProduto,cLote,cNumSerie,cSLote,lRetTodosPrd)
Local   nQtdPrdNF := 0
Local   nQtdPrdEm := 0
Local   nQtdeNota := 0
Local   nQtdeEmb  := 0
Local   aRet      := {}
Local   aRetNota  := {}
Local   aRetEmb   := {}

Default cProduto  := Space(TamSX3("CBL_PROD")[01])
Default cLote     := Space(TamSX3("CBL_LOTECT")[01])
Default cNumSerie := Space(TamSX3("CBL_NUMSER")[01])
Default cSLote    := Space(IIF(lSLoteCBL,TamSX3("CBL_SLOTE")[01],TamSX3("D2_NUMLOTE")[01]))
Default lRetTodosPrd := .F.

SD2->(DbSetOrder(3))
SD2->(DbSeek(xFilial('SD2')+CBK->(CBK_DOC+CBK_SERIE+CBK_CLIENT+CBK_LOJA)))
While SD2->(!Eof() .AND. D2_FILIAL+D2_DOC+D2_SERIE+D2_CLIENTE+D2_LOJA == xFilial('SF2')+CBK->(CBK_DOC+CBK_SERIE+CBK_CLIENT+CBK_LOJA))
	nPos := Ascan(aRetNota,{|x| x[1] == SD2->D2_COD})
	If nPos == 0
		aadd(aRetNota,{SD2->D2_COD,SD2->D2_QUANT})
	Else
		aRetNota[nPos,2] += SD2->D2_QUANT
	Endif
	If !Empty(cProduto) .AND. (SD2->(D2_COD+D2_LOTECTL+D2_NUMLOTE+D2_NUMSERI) == cProduto+cLote+cSLote+cNumSerie)
		nQtdPrdNF += SD2->D2_QUANT
	Endif
	nQtdeNota += SD2->D2_QUANT
	SD2->(DbSkip())
EndDo

CBL->(dbSetOrder(2))
CBL->(DbSeek(xFilial('CBL')+CBK->(CBK_DOC+CBK_SERIE+CBK_CLIENT+CBK_LOJA)))
While CBL->(!Eof() .and. CBL_FILIAL+CBL_DOC+CBL_SERIE+CBL_CLIENT+CBL_LOJA == xFilial('SF2')+CBK->(CBK_DOC+CBK_SERIE+CBK_CLIENT+CBK_LOJA))
	If lSLoteCBL
		If !Empty(cProduto) .AND. (CBL->(CBL_PROD+CBL_LOTECT+CBL_SLOTE+CBL_NUMSER) == cProduto+cLote+cSLote+cNumSerie)
			nQtdPrdEm += CBL->CBL_QTDEMB
		Endif
	Else
		If !Empty(cProduto) .AND. (CBL->(CBL_PROD+CBL_LOTECT+CBL_NUMSER) == cProduto+cLote+cNumSerie)
			nQtdPrdEm += CBL->CBL_QTDEMB
		Endif
	EndIf	
	nPos := Ascan(aRetEmb,{|x| x[1] == CBL->CBL_PROD})
	If nPos == 0
		aadd(aRetEmb,{CBL->CBL_PROD,CBL->CBL_QTDEMB})
	Else
		aRetEmb[nPos,2] += CBL->CBL_QTDEMB
	Endif
	nQtdeEmb += CBL->CBL_QTDEMB
	CBL->(DbSkip())
EndDo

If !Empty(cProduto)
	aRet := {nQtdPrdNF,nQtdPrdEm}
ElseIf !lRetTodosPrd
	aRet := {nQtdeNota,nQtdeEmb}
Else
	aRet := {aRetNota,aRetEmb}
Endif

Return aClone(aRet)

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o	 ³ Faltas 	  ³ Autor ³Henrique Gomes Oikawa³ Data ³ 31/03/04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³ Mostra produtos que faltam ser lidos                       ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
Static Function Faltas()
Local aSize,aSave := VTSAVE()
Local aItensNota	:= {}
Local aItensEmb	:= {}
Local aCab			:= {}
Local aItensFalta := {}
Local nI, nX
IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf

If Empty(cNota)
	Return
Endif

aItensAux  := AnalisaEmb(NIL,NIL,NIL,NIL,.T.)
aItensNota := aItensAux[1]
aItensEmb  := aItensAux[2]

For nI:=1 to Len(aItensNota)
	nX := Ascan(aItensEmb,{|x| x[01]==aItensNota[nI,1]})
	If nX > 0
		aadd(aItensFalta,{aItensNota[nI,01],AllTrim(Transform(aItensNota[nI,02]-aItensEmb[nX,02],"@E 999,999.99"))})
	Else
		aadd(aItensFalta,{aItensNota[nI,01],AllTrim(Transform(aItensNota[nI,02],"@E 999,999.99"))})
	Endif
Next

VTClear()                         
aCab  := {STR0028,STR0029} //"Produto"###"Qtde"
aSize := {15,12}
If lVT100B // GetMv("MV_RF4X20")
	VTaBrowse(0,0,3,19,aCab,aItensFalta,aSize)
Else
	VTaBrowse(0,0,7,19,aCab,aItensFalta,aSize)
EndIf
VtRestore(,,,,aSave)
Return
