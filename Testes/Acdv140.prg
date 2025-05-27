#INCLUDE "ACDV140.ch" 
#INCLUDE "protheus.ch"
#INCLUDE "apvt100.ch"


/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡Æo    ³Acdv140   ³ Autor ³ Ricardo               ³ Data ³ 27/07/01 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡Æo ³ Programa de Geracao de Volumes.                            ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³                                                            ³±±
±±³          ³                                                            ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ ACD                                                        ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
/*/  
Function Acdv140()
Local cNota   := ""
Local cSerie  := ""
Local cLoja   := ""
Local nQtdVol := 0
Local cVolume := ""
Local lVolta := .F.
Private cFornec := Space(06)
IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf

if lVT100B // GetMv("MV_RF4X20")
	While .T.
		If !lVolta
			cNota   := Space(TamSx3("F1_DOC")[1])
			cSerie  := Space(SerieNfId("SF1",6,"F1_SERIE"))
			cFornec := Space(TamSx3("F1_FORNECE")[1])
			cLoja   := Space(TamSx3("F1_LOJA")[1])
			cVolume := Space(10)
			nQtdVol := 0
		EndIf
		VTClear()
		@ 0,0  vtSay STR0001 //"Nota"
		@ 0,11 vtGet cNota pict '@!' when iif(lVolta,(VTKeyBoard(chr(13)),.T.),.T.) valid !Empty(cNota)
		@ 1,0  vtSay STR0002 //"Serie"
		@ 1,11 vtGet cSerie pict '!!!' when iif(lVolta,(VTKeyBoard(chr(13)),lVolta := .F.,.T.),.T.) valid Empty(cSerie) .or. CBMULTDOC("SF1",cNota,@cSerie)
		@ 2,0  vtSay STR0003 //"Fornecedor"
		@ 2,11 vtGet cFornec pict '@!'/* when iif(lVolta,(VTKeyBoard(chr(13)),lVolta := .F.,.T.),.T.)*/ valid !Empty(cFornec) .and. VTExistCPO("SA2",cFornec,1,STR0004+chr(13)+chr(10)+STR0005) F3 "FOR" //"Fornecedor nao"###"cadastrado!!!"
		VtRead
 		If !(vtLastKey() == 27)
			VtClear
			@ 0,0  vtSay STR0006 //"Loja"
			@ 0,11 vtGet cLoja pict '@!' when iif(vtRow() == 0 .and. vtLastKey() == 5,(VTKeyBoard(chr(27)),lVolta := .T.),.T.) valid !Empty(cLoja) .and. VTExistCPO("SA2",cFornec+cLoja,1,STR0007+chr(13)+chr(10)+STR0008) .And. VldForn(cFornec,cLoja,cNota,cSerie) .And. VldRemVnc(cNota, cSerie, cFornec, cLoja)//"Loja nao"###"cadastrada!!!"
			If ! UsaCB0("07")
				@ 1,0  vtSay STR0009 //"Qtd. Volumes"
				@ 1,11 vtGet nQtdVol pict '99999' valid vtlastkey()==5 .or. VldQtd(nQtdVol,cNota,cSerie,cFornec,cLoja)
			Else
				If GetMv("MV_REGVOL") =="0" // nao registra etiqueta avulsa
					@ 1,0  vtSay STR0009 //"Qtd. Volumes"
					@ 1,11 vtGet nQtdVol pict '99999' valid VldQtd(nQtdVol,cNota,cSerie,cFornec,cLoja)
				Else
					@ 1,0 vtSay STR0013 //"Etiqueta"
					@ 2,0 vtGet cVolume pict '@!' valid VldVolume(cVolume,cNota,cSerie,cFornec,cLoja)
				EndIf
			EndIf
			VtRead
		EndIf

		If lVolta
			Loop
		EndIf
		If vtLastKey() == 27
			Exit
		EndIf
	EndDo
else
	While .T.
		cNota   := Space(TamSx3("F1_DOC")[1])
		cSerie  := Space(SerieNfId("SF1",6,"F1_SERIE"))
		cFornec := Space(TamSx3("F1_FORNECE")[1])
		cLoja   := Space(TamSx3("F1_LOJA")[1])
		cVolume := Space(10)
		nQtdVol := 0
	
		VTClear()
		@ 0,0  vtSay STR0001 //"Nota"
		@ 0,11 vtGet cNota pict '@!' valid !Empty(cNota)
		@ 1,0  vtSay STR0002 //"Serie"
		@ 1,11 vtGet cSerie pict '!!!' valid Empty(cSerie) .or. CBMULTDOC("SF1",cNota,@cSerie)
		@ 2,0  vtSay STR0003 //"Fornecedor"
		@ 2,11 vtGet cFornec pict '@!' valid !Empty(cFornec) .and. VTExistCPO("SA2",cFornec,1,STR0004+chr(13)+chr(10)+STR0005) F3 "FOR" //"Fornecedor nao"###"cadastrado!!!"
		@ 3,0  vtSay STR0006 //"Loja"
		@ 3,11 vtGet cLoja pict '@!' valid !Empty(cLoja) .and. VTExistCPO("SA2",cFornec+cLoja,1,STR0007+chr(13)+chr(10)+STR0008) .And. VldForn(cFornec,cLoja,cNota,cSerie) .And. VldRemVnc(cNota, cSerie, cFornec, cLoja)//"Loja nao"###"cadastrada!!!"
	 
		If ! UsaCB0("07")
			@ 4,0  vtSay STR0009 //"Qtd. Volumes"
			@ 4,11 vtGet nQtdVol pict '99999' valid vtlastkey()==5 .or. VldQtd(nQtdVol,cNota,cSerie,cFornec,cLoja)
		Else
			If GetMv("MV_REGVOL") =="0" // nao registra etiqueta avulsa
				@ 4,0  vtSay STR0009 //"Qtd. Volumes"
				@ 4,11 vtGet nQtdVol pict '99999' valid VldQtd(nQtdVol,cNota,cSerie,cFornec,cLoja)
			Else
				@ 4,0 vtSay STR0013 //"Etiqueta"
				@ 5,0 vtGet cVolume pict '@!' valid VldVolume(cVolume,cNota,cSerie,cFornec,cLoja)
			EndIf
		EndIf
		VtRead
		If vtLastKey() == 27
			Exit
		EndIf
	EndDo
EndIf
Return .T.


Static Function VldForn(cForn,cLoja,cNota,cSerie)
If	ExistBlock('ACD140VF')
	If	! ExecBlock('ACD140VF',,,{cForn,cLoja,cNota,cSerie})
		VTKeyBoard(chr(20))
		Return .f.
	EndIf
EndIf
Return .t.


Static Function VldQtd(nQtdVol,cNota,cSerie,cFornec,cLoja)
Local i := 0
If	Empty(nQtdvol)
	Return .f.
EndIf
If	! VTYesNo(STR0010,STR0011,.T.) //"Imprime etiqueta ?"###"ATENCAO"
	VTKeyBoard(chr(20)) //Limpa o get
	Return .T.
EndIf
VtMsg(STR0012) //'Imprimindo'
CB5SetImp(CBRLocImp("MV_IACD02"),.T.)
For i := 1 To nQtdVol
	If	ExistBlock('IMG07')
		cResult :=ExecBlock("IMG07",,,{StrZero(i,10),cNota,cSerie,cFornec,cLoja,nQtdVol})
	EndIf
Next i
MSCBClosePrinter()
Return .t.


Static Function VldVolume(cVolume,cNota,cSerie,cFornec,cLoja)
If	Empty(cVolume)
	Return .t.
EndIf
If	! Empty(CBRetEti(cVolume))
	VtBeep(3)
	VTAlert(STR0014,STR0015,.t.,3000) //"Etiqueta ja registrada"###"Aviso"
	VtKeyBoard(chr(20))
	Return .f.
EndIf
CBGrvEti('07',{cVolume,cNota,cSerie,cFornec,cLoja},cVolume) 
VtKeyBoard(chr(20))
Return .F.

//-------------------------------------------------------------------
/*/{Protheus.doc} VldRemVnc
Validação de Notas fiscais vinculadas à remitos
@author marco.guimaraes
@since 29/04/14
@version 1.0
@param cNota, cSerie, cFornec, cLoja
@return lRet
/*/
//-------------------------------------------------------------------
Function VldRemVnc(cNota, cSerie, cFornec, cLoja)
Local lRet := .F.
Local cAliasSD1 := ""
Local nCntRemito := 0
Local cBlank := ""

lRet := SF1->(DbSeek(xFilial("SF1") + cNota + cSerie + cFornec + cLoja))

If lRet
	cAliasSD1 := GetNextAlias()
	BeginSql Alias cAliasSD1
		Select count(D1_REMITO) as nCntRemSQL
		FROM %table:SD1% SD1
		WHERE 
			D1_DOC = %exp:cNota%
			AND D1_REMITO = %exp:cBlank%
	EndSql
	
	nCntRemito := (cAliasSD1)->nCntRemSQL
	
	If nCntRemito > 0
		lRet := .T.
	Else
		VtBeep(3)
		VTAlert(STR0016,STR0015,.T.,3000) //"Nota apenas com remitos","AVISO"
		VtKeyBoard(chr(20))
		lRet := .F.
	EndIf
Else
	VtBeep(3)
	VTAlert(STR0017,STR0015,.T.,3000) //"Nota inválida","AVISO"
	VtKeyBoard(chr(20))	
	lRet := .F.
EndIf
	
Return lRet
