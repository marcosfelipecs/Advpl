#INCLUDE "Acdv060.ch" 
#INCLUDE "protheus.ch"
#INCLUDE "apvt100.ch"

Static __nSem:=0

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o	 ³ ACDV060    ³ Autor ³ Desenv. ACD         ³ Data ³ 17/04/01 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³ Distribuicao de produtos via coletor de radio frequencia   ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso		 ³ SigaACD                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/ 
Function ACDV060()
Local nOpc

If UsaCB0("01")
	ACDV0601(1)   // produto com cb0
Else
	VTCLear()
	@ 0,0 VTSay STR0001 // //'Selecione:'
	nOpc:=VTaChoice(1,0,3,VTMaxCol(),{STR0002,STR0003,STR0056}) //### //"Nota de Entrada"###"Producao"###"Sem Documento"
	VtClearBuffer()
	If nOpc == 1
		ACDV061()  // produto sem cb0 (Nota entrada)
	ElseIf nOpc == 2
		ACDV062()  // produto sem cb0 (Producao)
	ElseIf nOpc == 3
		ACDV063()  // produto sem cb0 e sem informar o codigo do documento
	EndIf
EndIf
Return NIL

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o	 ³ ACDV061    ³ Autor ³ Desenv. ACD         ³ Data ³ 17/04/01 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³ Distribuicao de produtos via coletor de radio frequencia   ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso		 ³ SigaACD                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
*/
Function ACDV061()
ACDV0601(2)
Return NIL

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o	 ³ ACDV062    ³ Autor ³ Desenv. ACD         ³ Data ³ 17/04/01 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³ Distribuicao de produtos via coletor de radio frequencia   ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso		 ³ SigaACD                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
*/       
Function ACDV062()
ACDV0601(3)
Return NIL

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o	 ³ ACDV063    ³ Autor ³ Desenv. ACD         ³ Data ³ 17/04/01 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³ Distribuicao de produtos via coletor de radio frequencia   ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso		 ³ SigaACD                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
*/
Function ACDV063()
If ! SIX->(DbSeek("SDA2"))
	VtAlert(STR0061+CRLF+"SDA"+CRLF+STR0062+CRLF+"2"+CRLF+STR0063+CRLF+"DA_FILIAL+DA_PRODUTO+DA_LOCAL+DA_LOTECTL",STR0064) // "Tabela" ### "Ordem" ### "Chave" ### "Indice obrigatorio"	
	Return
EndIf	
ACDV0601(4)
Return NIL

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o	 ³ ACDV0601    ³ Autor ³ Desenv. ACD         ³ Data ³ 17/04/01 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³ Distribuicao de produtos via coletor de radio frequencia   ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso		 ³ SigaACD                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
*/
// esta funcao esta sendo chamada tambem no acdv050
Function ACDV0601(nTipo,lEnvCQ)
Local oTpTab1	:= NIL
Local oTpTab2	:= NIL
Local bkey09
Local bkey24
Local nL := 0
Local nColArm
Local cPicEnd
Local cConsF3     := 'CBW' 
Local nTamCodEt2	 := TamSx3("CB0_CODET2")[1]
Local lVolta      := .F.
Local lACD060CA   := ExistBlock("ACD060CA")
Private lBranco   := .t.
Private cNota     := Space(TamSx3("F1_DOC")[1])
Private cSerie    := Space(SerieNfId("SF1",6,"F1_SERIE"))
Private cFornec   := Space(TamSx3("F1_FORNECE")[1])
Private cLoja     := Space(TamSx3("F1_LOJA")[1])
Private cDoc      := Space(TamSx3("D3_DOC")[1])
Private cProd     := Space( nTamCodEt2 )
Private nQtdePro  := 1
Private cArmazem  := Space(Tamsx3("B1_LOCPAD")[1])
Private cEndereco := Space(Tamsx3("BF_LOCALIZ")[1])
Private cEtiEnd   := Space(TamSx3("CB0_CODET2")[1])
Private aDist     := {}
Private aHisEti   := {}
Private cCondSF1  := "1 "   // variavel utilizada na consulta Sxb 'CBW'
Private lCQ
Private cArmEti   := space(Tamsx3("B1_LOCPAD")[1])
Private lForcaQtd := GetMV("MV_CBFCQTD",,"2") == "1" // Forca foco no GET Quantidade     
Private lItIguais := .F.
Private lArmazem  := .F.
Private lProdIg   := .F.
Private aPrDup    := {}
Private lForConf  := .T.


IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf
DEFAULT lEnvCQ    := .f.
DEFAULT cConsF3   := 'CBW' 

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³Ponto de entrada permitir forca o foco no campo quantidade após digitação do documento			   ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
If ExistBlock("V160FQTD")
	lForcaQtd := ExecBlock("V160FQTD",.F.,.F.)
	If ValType(lForcaQtd)<> "L"
		lForcaQtd := .f.   
    EndIf
EndIf

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³Ponto de entrada para alterar a consulta padrão F3 ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
If ExistBlock("ACD060F3")
	cConsF3 := ExecBlock("ACD060F3",.F.,.F.)
	If ValType(cConsF3)<> "C"
		cConsF3 := 'CBW'   
    EndIf
EndIf

nColArm := If(Len(cArmazem) < 4,3,6)
cPicEnd := If(Len(cArmazem) < 4,"@!","@!S11")
lCQ := lEnvCQ

oTpTab1 := FWTemporaryTable():New( "CABTMP" )

aStru :={{"CAB_NUMRF"	,"C",3,00}}

oTpTab1:SetFields( aStru )
oTpTab1:AddIndex("indice1", {"CAB_NUMRF"} )

oTpTab1:Create()

aStru:= {}
oTpTab2 := FWTemporaryTable():New( "ITETMP" )

aStru :={	{"ITE_RECNO","C",6,00},;
	{"ITE_FILIAL","C",2,00},;
	{"ITE_NUMSEQ","C",6,00},;
	{"ITE_QTD"   ,"N",12,4}}
	
oTpTab2:SetFields( aStru )
oTpTab2:AddIndex("indice1", {"ITE_RECNO","ITE_FILIAL","ITE_NUMSEQ"} )
oTpTab2:AddIndex("indice2", {"ITE_FILIAL","ITE_NUMSEQ"} )

oTpTab2:Create()

RegistraCab()

bkey09 := VTSetKey(09,{|| Informa()},STR0045) //"Informacoes"
bKey24 := VTSetKey(24,{|| Estorna()},STR0046)   // CTRL+X //"Estorno"
If nTipo == 1
	cProd     := Space(nTamCodEt2)
Else
	cProd     := IIf( FindFunction( 'AcdGTamETQ' ), AcdGTamETQ(), Space(48) )
EndIf

While .t.
	VTClear
	If ! lCQ
		@ 0,0 VTSAY STR0004 // //"Enderecamento"
	Else
		@ 0,0 VTSAY STR0043 //"Envio C.Q."
	EndIf
	nL := 0
	nQtdePro := 1
	If nTipo == 2 // nota
		If lVT100B
			//primeira tela
			@ 1,00 VTSAY  STR0005 VTGet cNota   	pict '@!'          	when VtLastkey()==5 .or. iif(lVolta,(VTKeyBoard(chr(13)),.T.),.T.) F3 cConsF3 // //'Nota '
			@ 1,14 VTSAY '-' VTGet cSerie       	pict '!!!'           when VtLastkey()==5 .or. iif(lVolta,(VTKeyBoard(chr(13)),.T.),.T.) Valid VldNota(@cNota,@cSerie,,,.T.)
			@ 2,00 VTSAY  STR0006 VTGet cFornec 	pict '@!' F3 'FOR'	when VtLastkey()==5 .or. iif(lVolta,(VTKeyBoard(chr(13)),.T.),.T.) Valid VldNota(cNota,cSerie,cFornec) // //'Forn.'
			@ 2,14 VTSAY '-' VTGet cLoja        	pict '@!'         	when VtLastkey()==5 .or. iif(lVolta .and. lForcaQtd,(VTKeyBoard(chr(13)),.T.),(lVolta := .F., .T.)) Valid (lBranco := .f.,VldNota(cNota,cSerie,cFornec,cLoja))
			@ 3,00 VTSAY STR0007 VTGet nQtdePro pict CBPictQtde()  valid nQtdePro > 0 when (lForcaQtd .and. lVolta, lVolta := .F.) //"Qtde."
			nL := 3
			VTRead
		Else
			@ 1,00 VTSAY  STR0005 VTGet cNota   	pict '@!'          	when Empty(cNota)   .or. VtLastkey()==5 F3 cConsF3 // //'Nota '
			@ 1,14 VTSAY '-' VTGet cSerie       	pict '!!!'          when (Empty(cSerie) .and. lBranco) .or. VtLastkey()==5 Valid VldNota(@cNota,@cSerie,,,.T.)
			@ 2,00 VTSAY  STR0006 VTGet cFornec 	pict '@!' F3 'FOR'	when Empty(cFornec) .or. VtLastkey()==5 Valid VldNota(cNota,cSerie,cFornec) // //'Forn.'
			@ 2,14 VTSAY '-' VTGet cLoja        	pict '@!'         	when Empty(cLoja)   .or. VtLastkey()==5 Valid (lBranco := .f.,VldNota(cNota,cSerie,cFornec,cLoja))
			@ 3,00 VTSAY STR0007  //"Qtde."
			@ 4,00 VTGet nQtdePro pict CBPictQtde()  valid nQtdePro > 0 when (lForcaQtd .or. VTLastkey() == 5)
			nL := 4
		Endif
	ElseIf nTipo == 3 // producao
			If lVT100B
				//primeira tela
				@ 1,00 VTSAY STR0008 VTGet cDoc      pict '@!'          when VtLastkey()==5 .or. iif(lVolta .and. lForcaQtd,(VTKeyBoard(chr(13)),.T.),(lVolta := .F., .T.)) Valid VldDoc(cDoc) // //'Documento '
				@ 2,00 VTSAY STR0007 VTGet nQtdePro  pict CBPictQtde()  valid nQtdePro > 0 when (lForcaQtd .and. lVolta, lVolta := .F.)//'Qtde.'
				VTRead
			Else
				@ 1,00 VTSAY STR0008 VTGet cDoc      pict '@!'          when VtLastkey()==5  Valid VldDoc(cDoc) // //'Documento '
				@ 2,00 VTSAY STR0007 VTGet nQtdePro  pict CBPictQtde()  valid nQtdePro > 0 when (lForcaQtd .or. VTLastkey() == 5)//'Qtde.'
				nL := 2
			EndIf
	ElseIf nTipo == 4
		If lVT100B
			//primeira tela
			@ 1,00 VTSAY STR0065 VTGet cArmEti  pict "@!" Valid ! Empty(cArmEti) when iif(lVolta .and. lForcaQtd,(VTKeyBoard(chr(13)),.T.),(lVolta := .F., .T.))  // "Armazem"
			@ 2,00 VTSAY STR0007 VTGet nQtdePro  pict CBPictQtde()  valid nQtdePro > 0 when (lForcaQtd .and. lVolta, lVolta := .F.)//'Qtde.'
			VTRead
		Else
			@ 1,00 VTSAY STR0065 VTGet cArmEti  pict "@!" Valid ! Empty(cArmEti) when iif(lVolta,(VTKeyBoard(chr(13)),lVolta := .F.,.T.),.T.)  // "Armazem"
			@ 2,00 VTSAY STR0007 VTGet nQtdePro  pict CBPictQtde()  valid nQtdePro > 0 when (lForcaQtd .or. VTLastkey() == 5)//'Qtde.'
			nL := 2
		Endif
	EndIf
	If UsaCB0("01") .and. UsaCB0("02")
		If lVT100B
			If !(vtLastKey() == 27)
				VTClear
			//segunda tela
				@ 0,0 VTSAY STR0009  // //'Etiqueta'
				@ 1,0 VTGET cProd PICTURE "@!" VALID VTLastkey() == 05 .or. Empty(cProd) .or.  VldEtiq() when iif(vtRow() == 1 .and. vtLastKey() == 5,(VTKeyBoard(chr(27)),lVolta := .T.),.T.)
			Endif
		Else
			@ ++nL,0 VTSAY STR0009  // //'Etiqueta'
			@ ++nL,0 VTGET cProd PICTURE "@!" VALID VTLastkey() == 05 .or. Empty(cProd) .or.  VldEtiq()
		Endif
	Else
		If lVT100B
			If !(vtLastKey() == 27)
				VTClear
			//segunda tela
				@ 0,0 VTSAY STR0010+': ' VTGET cProd PICTURE "@!" VALID VTLastkey() == 05 .or. Empty(cProd) .or. VldProd() when iif(vtRow() == 0 .and. vtLastKey() == 5,(VTKeyBoard(chr(27)),lVolta := .T.),.T.)
				@ 1,0 VTSAY STR0011  // //'Endereco'
				If UsaCB0("02")
					@ 2,0 VTGET cProd PICTURE "@!" VALID  VTLastkey() == 05 .or. VldEtiq("02")
				Else
					@ 2,0 VTGet cArmazem PICTURE "@!" Valid VTLastkey() == 05 .or. ! Empty(cArmazem)
					@ 2,nColArm VTSAY "-" VTGET cEndereco PICTURE cPicEnd VALID VtLastKey()==5 .or. VldEndereco()
				Endif
			Endif
		Else
			@ ++nL,0 VTSAY STR0010+': ' VTGET cProd PICTURE "@!" VALID VTLastkey() == 05 .or. Empty(cProd) .or. VldProd()
			@ ++nL,0 VTSAY STR0011  // //'Endereco'
			If UsaCB0("02")
				@ ++nL,0 VTGET cProd PICTURE "@!" VALID  VTLastkey() == 05 .or. VldEtiq("02")
			Else
				@ ++nL,0 VTGet cArmazem PICTURE "@!" Valid VTLastkey() == 05 .or. ! Empty(cArmazem)
				@   nL,nColArm VTSAY "-" VTGET cEndereco PICTURE cPicEnd VALID VtLastKey()==5 .or. VldEndereco()
			End
		EndIf
	EndIf
	VTREAD
	If lVolta
		Loop
	Endif
	If VTLASTKEY()==27
		If Empty(aDist) .or. VTYesNo(STR0012,STR0013,.T.)		  //### //"Saindo perdera o que foi lido, confirma saida?"###"ATENCAO"
			If lACD060CA
				ExecBlock("ACD060CA",.F.,.F.)
			EndIf
			Exit
		EndIf
	EndIf
	If nTipo == 1
		cProd     := Space(nTamCodEt2)
	Else
		cProd     := IIf( FindFunction( 'AcdGTamETQ' ), AcdGTamETQ(), Space(48) )
	EndIf

End

vtsetkey(09,bkey09)
vtsetkey(24,bkey24)
If UsaCB0("01")  
	CB0->(MsRUnlock())
EndIf

RegistraCab(.F.)
oTpTab1:Delete()
oTpTab2:Delete()

Return

//-------------------------------------------------------------------
/*/{Protheus.doc} VldNota()

@author Totvs
@version 1.0
@return NIL
/*/
//-------------------------------------------------------------------
Static Function VldNota(cNota,cSerie,cFornec,cLoja,lSerie)

Default cNota	:= ""
Default cSerie	:= ""
Default cFornec	:= ""
Default cLoja	:= "" 
Default lSerie 	:= .F.
Default lForConf 	:= .T.

If VtLastkey() == 05
	Return .t.
EndIf
SF1->(DbSetOrder(1))

If lSerie
	CBMULTDOC("SF1",cNota,@cSerie)
EndIf

If ! SF1->(MsSeek(xFilial('SF1')+cNota+cSerie+cFornec+cLoja))
	VTBEEP(2)
	VTALERT(STR0014,STR0015,.T.,3000)  //### //"Nota nao encontrada"###"AVISO"
	VTKeyBoard(chr(20))
	Return .f.
EndIf

lForConf := FornecConf(cFornec, cLoja)

If lForConf
	If GetMv("MV_CONFFIS")=="S" .AND. GetMv("MV_TPCONFF",.F.,"1")=="2" .AND. SF1->F1_STATCON<>"1"
		VTBEEP(2)
		VTAlert(STR0058 + AllTrim(SF1->F1_DOC) + "/" + AllTrim(SF1->&(SerieNfId("SF1",3,"F1_SERIE"))) + STR0059,STR0015,.T.,4000)
		VTKeyBoard(chr(20))
		Return .f.
	EndIf
EndIf

Return .t.

Static Function VldDoc(cDoc)
SD3->(DbSetOrder(2))
If ! SD3->(MsSeek(xFilial('SD3')+cDoc))
	VTBEEP(2)
	VTALERT(STR0016,STR0015,.T.,3000)  //### //"Documento nao encontrado"###"AVISO"
	VTKeyBoard(chr(20))
	Return .f.
Endif






Return .T.

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o	 ³ VldProd    ³ Autor ³ Totvs               ³ Data ³ 01/01/01 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³ Valida o Produto          								  ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso		 ³ ACDV060                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
Static function VldProd()
Local nPos      := 0
Local nX        := 0
Local nP        := 0
Local aEtiqueta := {}
Local cChavPesq := ""
Local aChavPesq := ""
Local cChave    := ""
Local cTipDis   := ""
Local cLote     := Space(TamSX3("B8_LOTECTL")[1])
Local cSLote    := Space(TamSX3("B8_NUMLOTE")[1])
Local aDistBKP  := aClone(aDist)
Local aHisEtiBKP:= aClone(aHisEti)
Local aGrava    :={}
Local aItensPallet := CBItPallet(cProd)
Local lIsPallet := .t.
Local lForcaArm := .f.
Local lConfFis  := AllTrim(SuperGetMV("MV_CONFFIS",.F.,"N")) == "S"
Local lAIC060VP   := ExistBlock("AIC060VPR")
Private cNumSeri  := Space(TamSX3("BF_NUMSERI")[1])

If len(aItensPallet) == 0
	aItensPallet:={cProd}
	lIsPallet := .f.
EndIf

begin Sequence
For nP:= 1 to len(aItensPallet)
	cProd :=  aItensPallet[nP]
	If UsaCB0("01")
		aEtiqueta := CBRetEti(cProd,"01")
		If Empty(aEtiqueta)
			VTBEEP(2)
			VTALERT(STR0017,STR0015,.T.,4000)  //### //"Etiqueta invalida."###"AVISO"
			Break
		EndIf
		If ! lIsPallet .and. ! Empty(CB0->CB0_PALLET)
			VTBeep(2)
			VTALERT(STR0047,STR0015,.T.,4000)   //"AVISO" //"Etiqueta invalida, Produto pertence a um Pallet"
			Break
		EndIf
		If ascan(aHisEti,{|x|x[1] == cProd}) > 0
			VTBEEP(2)
			VTALERT(STR0018,STR0015,.T.,4000)  //### //"Produto ja foi lido."###"AVISO"
			Break
		EndIf
		If !Empty(cArmEti) .and. aEtiqueta[10] # cArmEti
			VTBEEP(2)
			VTALERT(STR0048,STR0015,.T.,4000)  //"AVISO" //"Etiqueta invalida, armazem diferente!"
			Break
		EndIf
		If Empty(aEtiqueta[2])	//-- Se vazio, etiqueta de caixa (produto a granel)
			aEtiqueta[2] := 1
			nQtdePro := CBQtdEmb(aEtiqueta[1]) //-- Pede leitura da quantidade
			If Empty(nQtdePro)
				Break
			EndIf
		Else
			If !Empty(aEtiqueta[9]) .and. aEtiqueta[10] # almoxCQ()
				VTBEEP(2)
				VTALERT(STR0019+" "+aEtiqueta[9],STR0015,.T.,4000)  //### //"Produto ja foi enderecado."###"AVISO"
				Break
			EndIf		
			If lConfFis .And. Empty(aEtiqueta[4]+aEtiqueta[5]+aEtiqueta[6]+aEtiqueta[7]+aEtiqueta[11]+aEtiqueta[12])
				VTBEEP(2)
				VTALERT(STR0020,STR0015,.T.,4000)  //### //"Produto nao conferido"###"AVISO"
				Break
			EndIf		
			cNota    := aEtiqueta[4]
			cSerie   := aEtiqueta[5]
			cFornec  := aEtiqueta[6]
			cLoja    := aEtiqueta[7]
			cLote    := aEtiqueta[16]
			cSLote   := aEtiqueta[17]
			cNumseri := aEtiqueta[23]
			If CBChkSer(aEtiqueta[1]) .And. ! CBNumSer(@cNumseri,Nil,aEtiqueta)
				Break
			EndIf
			If CBChkSer(aEtiqueta[1]) .And. !VldEndSer(cNumseri,aDist)	
				Break
			EndIf	
			aEtiqueta[23] := cNumseri//Numero de Serie
		EndIf
		cArmEti  := aEtiqueta[10]  //Armazem a ser validado nas leituras das etiquetas
	Else
		If ! CBLoad128(@cProd)
			Break
		EndIf
		cTipId:=CBRetTipo(cProd)
		If ! cTipId $ "EAN8OU13-EAN14-EAN128"
			VTBEEP(2)
			VTALERT(STR0017,STR0015,.T.,4000) //### //"Etiqueta invalida."###"AVISO"
			Break
		EndIf
		aEtiqueta := CBRetEtiEAN(cProd)
		If Empty(aEtiqueta) .or. Empty(aEtiqueta[2])
			VTBEEP(2)
			VTALERT(STR0017,STR0015,.T.,4000)  //### //"Etiqueta invalida."###"AVISO"
			Break
		EndIf
		// Valida se a nota tem itens com o mesmo codigo de produto e armazens distintos 
		If !Empty(cNota) .And. !Empty(cSerie) .And. !Empty(cFornec) .And. !Empty(cLoja)
			lItIguais := VldItIguais(cNota,cSerie,cFornec,cLoja,aEtiqueta[1])
			IF	lArmazem .And. lProdIg
				InfArm()
			EndIF
		EndIf
		nQE:= 1
		If ! CBProdUnit(aEtiqueta[1])
			nQE := CBQtdEmb(aEtiqueta[1])
			If Empty(nQE)
				Break
			EndIf
		EndIf
		aEtiqueta[2]:=aEtiqueta[2]*nQE
		cLote := aEtiqueta[3]
		If ! CBRastro(aEtiqueta[1],@cLote,@cSLote)
			Break
		EndIf
		cNumseri := aEtiqueta[5]
		If CBChkSer(aEtiqueta[1]) .And. ! CBNumSer(@cNumseri,Nil,aEtiqueta)
			Break
		EndIf
		If CBChkSer(aEtiqueta[1]) .And. !VldEndSer(cNumseri,aDist)	
			Break
		EndIf
		aEtiqueta[3]:= cLote   //Lote   
		aEtiqueta[4]:= cSLote  //Sublote
		aEtiqueta[5]:= cNumseri//Numero de Serie
	EndIf
	If ! CBProdLib(cArmazem,aEtiqueta[1])
		Break
	EndIf
	// quando os elementos abaixo estiverem em branco e' porque nao foi conferido
	If lAIC060VP .and. ! ExecBlock("AIC060VPR",.F.,.F.,{cProd,aEtiqueta[2]*nQtdePro,aEtiqueta})
		Break
	EndIf
	
   If Empty(cNota+cSerie+cFornec+cLoja+cDoc) .and. ! UsaCB0("01") .And. CB0->CB0_ORIGEM <>"SD3" 
		If !Empty(cArmazem)
			cArmEti := cArmazem
		EndIf
		cTipDis   := "SDA"
		aChavPesq :=RetNSeqSDA(aEtiqueta[1],aEtiqueta[2]*nQtdePro,cLote,cSLote)
		If Empty(aChavPesq)
			VTBEEP(2)
			VTALERT(STR0021,STR0015,.T.,4000)  //### //"Nao tem saldo a distribuir"###"AVISO"
			Break
		EndIf    
	ElseIf Empty(cNota+cSerie+cFornec+cLoja+cDoc) .And. UsaCB0("01") .And. CB0->CB0_ORIGEM <>"SD3" 
		cTipDis   := "SDA"
		aChavPesq :=RetNSeqSDA(aEtiqueta[1],aEtiqueta[2]*nQtdePro,cLote,cSLote)
		If Empty(aChavPesq)
			VTBEEP(2)
			VTALERT(STR0021,STR0015,.T.,4000)  //### //"Nao tem saldo a distribuir"###"AVISO"
			Break
		EndIf
	ElseIf !Empty(cNota+cSerie+cFornec+cLoja) .And.If(!UsaCB0("01"),.T.,CB0->CB0_ORIGEM <>"SD3")
		cTipDis   := "SD1"
		cChave    :=cNota+cSerie+cFornec+cLoja
		aChavPesq :=RetNumSeq(cChave+aEtiqueta[1],aEtiqueta[2]*nQtdePro,cTipDis,cLote,cSLote)
		If Empty(aChavPesq)
			aChavPesq := RetNSeqSDA(aEtiqueta[1],aEtiqueta[2]*nQtdePro,cLote,cSLote)
			If !Empty(aChavPesq) 
				cTipDis   := "SDA"
			EndIf	
		EndIf
		If Empty(aChavPesq)
			VTBEEP(2)
			VTALERT(STR0021,STR0015,.T.,4000)  //### //"Nao tem saldo a distribuir"###"AVISO"
			Break
		EndIf
	Else 
		cTipDis   := "SD3"
		If UsaCB0("01")
			SD3->(DbSetOrder(4))
			SD3->(MsSeek(xFilial('SD3')+CB0->CB0_NUMSEQ))
			cChave    := SD3->D3_DOC
		Else
			cChave    := cDoc
		EndIf
		aChavPesq :=RetNumSeq(cChave+aEtiqueta[1],aEtiqueta[2]*nQtdePro,cTipDis,cLote,cSLote)

		If Empty(aChavPesq)
			aChavPesq := RetNSeqSDA(aEtiqueta[1],aEtiqueta[2]*nQtdePro,cLote,cSLote)
			If !Empty(aChavPesq)
				cTipDis := "SDA"
			EndIf
		EndIf

		If Empty(aChavPesq)
			VTBEEP(2)
			VTALERT(STR0021,STR0015,.T.,4000)  //### //"Nao tem saldo a distribuir"###"AVISO"
			Break
		EndIf
	EndIf                   
	//### //"Trava o produto para ser unico"###
	If UsaCB0("01") .And. !CB0->(SimpleLock())
		VTALERT(STR0060,STR0015,.T.,3000)
		Break
	EndIf 
	
	For nX := 1 to len(aChavPesq)
		cChavPesq := aChavPesq[nX,1]
		aadd(aHisEti,{cProd,aEtiqueta[1],cChavPesq})
		nPos      := aScan(aDist,{|x| x[1] == cChavPesq .and. x[2] == aEtiqueta[1]  })
		If nPos > 0 .and. Empty(cNumseri)
			aDist[nPos,3] += aChavPesq[nX,2]
			aadd(aDist[nPos,5],{cProd,aEtiqueta[2],CB0->CB0_CODETI})
		Else
			aadd(aDist,{cChavPesq,aEtiqueta[1],aChavPesq[nX,2],aChavPesq[nX,3],{{cProd,aEtiqueta[2],CB0->CB0_CODETI}},cLote,cSLote,cNumseri})
		EndIf
		aadd(aGrava,{xFilial(cTipDis),cChavPesq,aChavPesq[nX,2]})
	Next
Next
For nX:= 1 to len(aGrava)
	GravaQtd(aGrava[nX,1],aGrava[nX,2],aGrava[nX,3])
Next
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³Ponto de entrada, não apaga o conteúdo dos outros Gets e permite força o foco no Get armazém. |
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
If ExistBlock("V160FArm")
	lForcaArm := ExecBlock("V160FArm",.f.,.f.)
	If ValType(lForcaArm)<> "L"
		lForcaArm := .f.   
    EndIf
EndIf
If lForcaArm
 	VTGetSetFocus("cArmazem")
elseIf lForcaQtd
	VTGetSetFocus("nQtdePro")
	nQtdePro := 1
	VTGetRefresh("nQtdePro")
	cProd := Space(Len(cProd))
	VTGetRefresh("cProd")
Else
	nQtdePro := 1
	VTGetRefresh("nQtdePro")
	cProd := Space(Len(cProd))
	VTGetRefresh("cProd")
EndIf

Return .f.
end sequence
aDist  := aClone(aDistBKP)
aHisEti:= aClone(aHisEtiBKP)
nQtdePro := 1
VTGetRefresh("nQtdePro")
VTGetRefresh("cArmazem")
cProd := Space(Len(cProd))
VTGetRefresh("cProd")
Return .F.
/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o	 ³ RetNSeqSDA ³ Autor ³ Sandro              ³ Data ³ 25/05/09 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³ Retorna o saldo a endereçar 								  ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso		 ³ ACDV060                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
Static Function RetNSeqSDA(cProduto,nQtde,cLote,cSLote)
Local aNumSeq :={}
Local nSaldo  := nQtde
Local nQtdBx  := 0 
Local cAlias  := "SDA"
Local cQuery  := ""
Local aArea   := SDA->(GetArea())
Local aAreaSF1:= SF1->(GetArea())
Local aNF     := {}
Local nPos    := 0
Local nX
Default lForConf  := .T.

cAlias  := CriaTrab(Nil,.F.)          
cQuery	:= "  SELECT DA_FILIAL,DA_PRODUTO,DA_LOCAL,DA_LOTECTL,DA_NUMLOTE, "
cQuery  += "  DA_NUMSEQ,DA_SALDO,DA_DOC,DA_SERIE,DA_CLIFOR,DA_LOJA "
cQuery	+= "  FROM " +RetSqlName('SDA') + " SDA "
cQuery	+= "  WHERE SDA.DA_FILIAL  = '"+xFilial("SDA")+"' "
cQuery	+= "  AND SDA.DA_PRODUTO = '"+cProduto+"' "
cQuery	+= "  AND SDA.DA_LOCAL   = '"+cArmEti+"' "
If !Empty(cLote)
	cQuery	+= "  AND SDA.DA_LOTECTL = '"+cLote+"' "
EndIf
If !Empty(cSLote)
	cQuery	+= "  AND SDA.DA_NUMLOTE = '"+cSLote+"' "
EndIf
cQuery	+= "  AND SDA.DA_SALDO > 0 "
cQuery	+= "  AND SDA.D_E_L_E_T_ = ' ' "
cQuery := ChangeQuery(cQuery)
dbUseArea ( .T., "TOPCONN", TCGENQRY(,,cQuery), cAlias, .F., .T.)

While (cAlias)->(!Eof() .and. xFilial('SD1')+cProduto+cArmEti+cLote+cSLote == DA_FILIAL+DA_PRODUTO+DA_LOCAL+DA_LOTECTL+DA_NUMLOTE) 
	nQtdBx := RetSaldo((cAlias)->(xFilial("SDA")+DA_NUMSEQ),(cAlias)->DA_SALDO)
	If Empty(nQtdBx)
		(cAlias)->(DbSkip())
		Loop
	EndIf
	If nQtdBx > nSaldo
		nQtdBx :=nSaldo
	EndIf
	nPos := Ascan(aNF,{|x| x[01]+x[02]+x[03]+x[04] == (cAlias)->DA_DOC+(cAlias)->DA_SERIE+(cAlias)->DA_CLIFOR+(cAlias)->DA_LOJA})
	If nPos == 0
		aadd(aNF,{(cAlias)->DA_DOC,(cAlias)->DA_SERIE,(cAlias)->DA_CLIFOR,(cAlias)->DA_LOJA})
	Endif
	aadd(aNumSeq,{(cAlias)->DA_NUMSEQ,nQtdBx,'SDA'})
	nSaldo -=nQtdBx
	If Empty(nSaldo)
		Exit
	EndIf
	(cAlias)->(DbSkip())
EndDo
If nSaldo > 0
	aNumSeq :={}
EndIf 

For nX:=1 to Len(aNF)
	lForConf := FornecConf(aNF[nx,03], aNF[nx,04])
	If lForConf
		If !Empty(aNumSeq) .AND. !Empty(aNF) .AND. GetMv("MV_CONFFIS")=="S" .AND. GetMv("MV_TPCONFF",.F.,"1")=="2"
			SF1->(DbSetOrder(1))			
				If SF1->(DbSeek(xFilial("SF1")+aNF[nX,01]+aNF[nX,02]+aNF[nX,03]+aNF[nX,04])) .AND. SF1->F1_STATCON<>"1"
					VTAlert(STR0058 + AllTrim(aNF[nX,01]) + "/" + AllTrim(aNF[nX,02]) + iIf(UsaCB0("01"),STR0068,"")+STR0059,STR0015,.T.,4000)
					aNumSeq :={}
					Exit
				Endif				
		Endif
	EndIf
Next	
		
(cAlias)->(DbCloseArea())    

RestArea(aAreaSF1)
RestArea(aArea)
Return aNumSeq

Static Function RetNumSeq(cChave,nQtde,cTipDis,cLote,cSLote)
Local aNumSeq :={}
Local aArea   := SDA->(GetArea())
Local aAreaSD1:= SD1->(GetArea())
Local nSaldo  :=nQtde
Local nQtdBx  := 0
Local cNumSeq := ""
Local cLocal
Local cLocProc  := GetMvNNR('MV_LOCPROC','99')
Local cQuery    := ''
Local cQuery1   := ''
Local cQuery2   := ''
Local cAliasSDA := ''
Local lVldSaldo := .T.

If cTipDis == "SD1"
	SD1->(DbSetOrder(1))
	SD1->(MsSeek(xFilial('SD1')+cChave))
	If lProdIg .And. !lArmazem
				cAliasSDA  := GetNextAlias()          
				cQuery	:= "  SELECT DA_FILIAL,DA_PRODUTO,DA_LOCAL,DA_LOTECTL,DA_NUMLOTE, "
				cQuery += "  DA_NUMSEQ,DA_SALDO,DA_DOC,DA_SERIE,DA_CLIFOR,DA_LOJA "
				cQuery	+= "  FROM " +RetSqlName('SDA') + " SDA "
				cQuery	+= "  WHERE SDA.DA_FILIAL  = '"+xFilial("SDA")+"' "
				cQuery	+= "  AND SDA.DA_PRODUTO = '"+SD1->D1_COD+"' "
				cQuery	+= "  AND SDA.DA_LOTECTL = '"+cLote+"' "
				cQuery	+= "  AND SDA.DA_NUMLOTE = '"+cSLote+"' "
				cQuery	+= "  AND SDA.DA_DOC = '"+SD1->D1_DOC+"' "
				cQuery	+= "  AND SDA.DA_SERIE = '"+SD1->D1_SERIE+"' "
				cQuery	+= "  AND SDA.DA_CLIFOR = '"+SD1->D1_FORNECE+"' "
				cQuery	+= "  AND SDA.DA_LOJA = '"+SD1->D1_LOJA+"' "
				cQuery	+= "  AND SDA.DA_QTDORI >= '"+cValtochar(nQtde)+"' "
				cQuery	+= "  AND SDA.DA_SALDO >= '"+cValtochar(nQtde)+"' "
				cQuery	+= "  AND SDA.D_E_L_E_T_ = ' ' "
				cQuery := ChangeQuery(cQuery)
				dbUseArea ( .T., "TOPCONN", TCGENQRY(,,cQuery), cAliasSDA, .F., .T.)
			
				If Select(cAliasSDA) > 0
					SD1->(DbSetOrder(5))//D1_FILIAL+D1_COD+D1_LOCAL+D1_NUMSEQ
					SD1->(MsSeek(xFilial('SD1')+(cAliasSDA)->(DA_PRODUTO+DA_LOCAL+DA_NUMSEQ)))
					While SD1->(!Eof() .and. xFilial('SD1')+cChave == ;
						D1_FILIAL+D1_DOC+D1_SERIE+D1_FORNECE+D1_LOJA+D1_COD .And. D1_NUMSEQ ==(cAliasSDA)->DA_NUMSEQ)
						
						If ! ( SD1->D1_LOTECTL==cLote .and. SD1->D1_NUMLOTE ==cSLote )
							SD1->(DbSkip())
							(cAliasSDA)->(DbSkip())
							Loop
						EndIf
						If (lCQ 	.and. SD1->D1_LOCAL # AlmoxCQ()) //.or. (!lCQ .and. SD1->D1_LOCAL == AlmoxCQ())
							SD1->(DbSkip())
							(cAliasSDA)->(DbSkip())
							Loop
						EndIf
						If Empty(SD1->D1_NUMCQ) .or. lCQ
							nQtdBx := RetSaldo(SD1->(xFilial("SD1")+D1_NUMSEQ),(cAliasSDA)->DA_SALDO)
							cNumSeq := SD1->D1_NUMSEQ
						Else
							SD7->(DBSetOrder(2))
							If ! SD7->(MsSeek(xFilial('SD7')+SD1->(D1_NUMCQ+D1_COD+D1_LOCAL+D1_NUMSEQ)))
								SD1->(DbSkip())
								(cAliasSDA)->(DbSkip())
								Loop
							EndIf
							While SD7->(! Eof() .and. D7_FILIAL+D7_NUMERO+D7_PRODUTO+D7_LOCAL == xFilial('SD7')+SD1->(D1_NUMCQ+D1_COD+D1_LOCAL)) .And. nQtdBx < nSaldo
				                // --- QTD. LIBERADA PELO CQ
								If SD7->D7_TIPO == 1
									(cAliasSDA)->(MsSeek(xFilial('SDA')+SD1->D1_COD+SD7->D7_LOCDEST+SD7->D7_NUMSEQ))
									nQtdBx += RetSaldo(SD7->(xFilial("SD7")+D7_NUMSEQ),(cAliasSDA)->DA_SALDO)
									cNumSeq := SD7->D7_NUMSEQ
								EndIf
								SD7->(DbSkip())
								(cAliasSDA)->(DbSkip())
							EndDo
						EndIf
						If Empty(nQtdBx)
							SD1->(DbSkip())
							(cAliasSDA)->(DbSkip())
							Loop
						EndIf
						If nQtdBx > nSaldo
							nQtdBx :=nSaldo
						EndIf  
						If !Empty(cNumSeq)
							If Empty(SD1->D1_NUMCQ) .or. lCQ
								aAdd(aNumSeq,{cNumSeq,nQtdBx,'SD1'})
							Else			
								aAdd(aNumSeq,{cNumSeq,nQtdBx,'SD3'})
							EndIf
						EndIf
						nSaldo -=nQtdBx
						cNumSeq := ""
						If Empty(nSaldo)
							Exit
						EndIf
						SD1->(DbSkip())
						(cAliasSDA)->(DbSkip())
				EndDo
			EndIF
	ElseIf lProdIg .And. lArmazem
			cAliasSDA  := GetNextAlias()          
			cQuery1 	:= "  SELECT DA_FILIAL,DA_PRODUTO,DA_LOCAL,DA_LOTECTL,DA_NUMLOTE, "
			cQuery1 	+= "  DA_NUMSEQ,DA_SALDO,DA_DOC,DA_SERIE,DA_CLIFOR,DA_LOJA "
			cQuery1	+= "  FROM " +RetSqlName('SDA') + " SDA "
			cQuery1	+= "  WHERE SDA.DA_FILIAL  = '"+xFilial("SDA")+"' "
			cQuery1	+= "  AND SDA.DA_PRODUTO = '"+SD1->D1_COD+"' "
			cQuery1	+= "  AND SDA.DA_LOCAL   = '"+cArmazem+"' "
			cQuery1	+= "  AND SDA.DA_LOTECTL = '"+cLote+"' "
			cQuery1	+= "  AND SDA.DA_NUMLOTE = '"+cSLote+"' "
			cQuery1	+= "  AND SDA.DA_DOC = '"+SD1->D1_DOC+"' "
			cQuery1	+= "  AND SDA.DA_SERIE = '"+SD1->D1_SERIE+"' "
			cQuery1	+= "  AND SDA.DA_CLIFOR = '"+SD1->D1_FORNECE+"' "
			cQuery1	+= "  AND SDA.DA_LOJA = '"+SD1->D1_LOJA+"' "
			cQuery1	+= "  AND SDA.DA_QTDORI >= '"+cValtochar(nQtde)+"' "
			cQuery1	+= "  AND SDA.DA_SALDO >= '"+cValtochar(nQtde)+"' "
			cQuery1	+= "  AND SDA.D_E_L_E_T_ = ' ' "
			cQuery1 := ChangeQuery(cQuery1)
			dbUseArea ( .T., "TOPCONN", TCGENQRY(,,cQuery1), cAliasSDA, .F., .T.)
			
			If Select(cAliasSDA) > 0
				lVldSaldo := .F.
			EndIF
		
		
			IF lVldSaldo
				cAliasSDA  := GetNextAlias()          
				cQuery2	:= "  SELECT DA_FILIAL,DA_PRODUTO,DA_LOCAL,DA_LOTECTL,DA_NUMLOTE, "
				cQuery2   += "  DA_NUMSEQ,DA_SALDO,DA_DOC,DA_SERIE,DA_CLIFOR,DA_LOJA "
				cQuery2	+= "  FROM " +RetSqlName('SDA') + " SDA "
				cQuery2	+= "  WHERE SDA.DA_FILIAL  = '"+xFilial("SDA")+"' "
				cQuery2	+= "  AND SDA.DA_PRODUTO = '"+SD1->D1_COD+"' "
				cQuery2	+= "  AND SDA.DA_LOCAL   = '"+cArmazem+"' "
				cQuery2	+= "  AND SDA.DA_LOTECTL = '"+cLote+"' "
				cQuery2	+= "  AND SDA.DA_NUMLOTE = '"+cSLote+"' "
				cQuery2	+= "  AND SDA.DA_DOC = '"+SD1->D1_DOC+"' "
				cQuery2	+= "  AND SDA.DA_SERIE = '"+SD1->D1_SERIE+"' "
				cQuery2	+= "  AND SDA.DA_CLIFOR = '"+SD1->D1_FORNECE+"' "
				cQuery2	+= "  AND SDA.DA_LOJA = '"+SD1->D1_LOJA+"' "
				cQuery2	+= "  AND SDA.DA_SALDO > 0 "
				cQuery2	+= "  AND SDA.D_E_L_E_T_ = ' ' "
				cQuery2 := ChangeQuery(cQuery2)
				dbUseArea ( .T., "TOPCONN", TCGENQRY(,,cQuery2), cAliasSDA, .F., .T.)
			EndIF
			
			SD1->(DbSetOrder(5))//D1_FILIAL+D1_COD+D1_LOCAL+D1_NUMSEQ
			SD1->(MsSeek(xFilial('SD1')+(cAliasSDA)->(DA_PRODUTO+DA_LOCAL+DA_NUMSEQ)))
		
		
		While SD1->(!Eof() .and. xFilial('SD1')+cChave == ;
			D1_FILIAL+D1_DOC+D1_SERIE+D1_FORNECE+D1_LOJA+D1_COD .And. D1_NUMSEQ ==(cAliasSDA)->DA_NUMSEQ)
			
			If ! ( SD1->D1_LOTECTL==cLote .and. SD1->D1_NUMLOTE ==cSLote )
				SD1->(DbSkip())
				(cAliasSDA)->(DbSkip())
				Loop
			EndIf
			If (lCQ 	.and. SD1->D1_LOCAL # AlmoxCQ()) //.or. (!lCQ .and. SD1->D1_LOCAL == AlmoxCQ())
				SD1->(DbSkip())
				(cAliasSDA)->(DbSkip())
				Loop
			EndIf
			If Empty(SD1->D1_NUMCQ) .or. lCQ
				nQtdBx := RetSaldo(SD1->(xFilial("SD1")+D1_NUMSEQ),(cAliasSDA)->DA_SALDO)
				cNumSeq := SD1->D1_NUMSEQ
			Else
				SD7->(DBSetOrder(2))
				If ! SD7->(MsSeek(xFilial('SD7')+SD1->(D1_NUMCQ+D1_COD+D1_LOCAL+D1_NUMSEQ)))
					SD1->(DbSkip())
					(cAliasSDA)->(DbSkip())
					Loop
				EndIf
				While SD7->(! Eof() .and. D7_FILIAL+D7_NUMERO+D7_PRODUTO+D7_LOCAL == xFilial('SD7')+SD1->(D1_NUMCQ+D1_COD+D1_LOCAL)) .And. nQtdBx < nSaldo
	                // --- QTD. LIBERADA PELO CQ
					If SD7->D7_TIPO == 1
						(cAliasSDA)->(MsSeek(xFilial('SDA')+SD1->D1_COD+SD7->D7_LOCDEST+SD7->D7_NUMSEQ))
						nQtdBx += RetSaldo(SD7->(xFilial("SD7")+D7_NUMSEQ),(cAliasSDA)->DA_SALDO)
						cNumSeq := SD7->D7_NUMSEQ
					EndIf
					SD7->(DbSkip())
					(cAliasSDA)->(DbSkip())
				EndDo
			EndIf
			If Empty(nQtdBx)
				SD1->(DbSkip())
				(cAliasSDA)->(DbSkip())
				Loop
			EndIf
			If nQtdBx > nSaldo
				nQtdBx :=nSaldo
			EndIf  
			If !Empty(cNumSeq)
				If Empty(SD1->D1_NUMCQ) .or. lCQ
					aAdd(aNumSeq,{cNumSeq,nQtdBx,'SD1'})
				Else			
					aAdd(aNumSeq,{cNumSeq,nQtdBx,'SD3'})
				EndIf
			EndIf
			nSaldo -=nQtdBx
			cNumSeq := ""
			If Empty(nSaldo)
				Exit
			EndIf
			SD1->(DbSkip())
			(cAliasSDA)->(DbSkip())
		EndDo
	Else
		While SD1->(!Eof() .and. xFilial('SD1')+cChave == ;
			D1_FILIAL+D1_DOC+D1_SERIE+D1_FORNECE+D1_LOJA+D1_COD)
			If ! ( SD1->D1_LOTECTL==cLote .and. SD1->D1_NUMLOTE ==cSLote )
				SD1->(DbSkip())
				Loop
			EndIf
			If (lCQ 	.and. SD1->D1_LOCAL # AlmoxCQ()) //.or. (!lCQ .and. SD1->D1_LOCAL == AlmoxCQ())
				SD1->(DbSkip())
				Loop
			EndIf
			If Empty(SD1->D1_NUMCQ) .or. lCQ
				SDA->(MsSeek(xFilial('SDA')+SD1->(D1_COD+D1_LOCAL+D1_NUMSEQ)))
				nQtdBx := RetSaldo(SD1->(xFilial("SD1")+D1_NUMSEQ),SDA->DA_SALDO)
				cNumSeq := SD1->D1_NUMSEQ
			Else
				SD7->(DBSetOrder(1))
				If ! SD7->(MsSeek(xFilial('SD7')+SD1->(D1_NUMCQ+D1_COD+D1_LOCAL)))
					SD1->(DbSkip())
					Loop
				EndIf
				While SD7->(! Eof() .and. D7_FILIAL+D7_NUMERO+D7_PRODUTO+D7_LOCAL == xFilial('SD7')+SD1->(D1_NUMCQ+D1_COD+D1_LOCAL)) .And. nQtdBx < nSaldo
	                // --- QTD. LIBERADA PELO CQ
					If SD7->D7_TIPO == 1
						SDA->(MsSeek(xFilial('SDA')+SD1->D1_COD+SD7->D7_LOCDEST+SD7->D7_NUMSEQ))
						nQtdBx += RetSaldo(SD7->(xFilial("SD7")+D7_NUMSEQ),SDA->DA_SALDO)
						cNumSeq := SD7->D7_NUMSEQ
					EndIf
					SD7->(DbSkip())
				EndDo
			EndIf
			If Empty(nQtdBx)
				SD1->(DbSkip())
				Loop
			EndIf
			If nQtdBx > nSaldo
				nQtdBx :=nSaldo
			EndIf  
			If !Empty(cNumSeq)
				If Empty(SD1->D1_NUMCQ) .or. lCQ
					aAdd(aNumSeq,{cNumSeq,nQtdBx,'SD1'})
				Else			
					aAdd(aNumSeq,{cNumSeq,nQtdBx,'SD3'})
				EndIf
			EndIf
			nSaldo -=nQtdBx
			cNumSeq := ""
			If Empty(nSaldo)
				Exit
			EndIf
			SD1->(DbSkip())
		EndDo
	EndIf
ElseIf cTipDis == "SD3"
	SD3->(DbSetOrder(2))
	SD3->(MsSeek(xFilial('SD3')+cChave))
	While SD3->(!Eof() .and. xFilial('SD3')+cChave == D3_FILIAL+D3_DOC+D3_COD)
		If ! ( SD3->D3_LOTECTL==cLote .and. SD3->D3_NUMLOTE ==cSLote )
			SD3->(DbSkip())
			Loop
		EndIf
		If UsaCB0("01") .and. CB0->CB0_NUMSEQ <> SD3->D3_NUMSEQ
			SD3->(DbSkip())
			Loop
		EndIf
		If	( lCQ .and. SD3->D3_LOCAL #  AlmoxCQ()) .or. ;
			(!lCQ .and. SD3->D3_LOCAL == AlmoxCQ())
			SD3->(DbSkip())
			Loop
		EndIf
		SF5->(DbSetOrder(1))
		If ! SF5->(MsSeek(xFilial("SF5")+SD3->D3_TM))
			conout(STR0049) //'nao encontrou o tm no SF5'
		EndIf
		If SF5->F5_APROPR=="N" .and. Left(SD3->D3_TIPO,2)=="RE"
			cLocal:= cLocProc
			If Empty(cLocal)
				conout(STR0050) //'O local padrao para o armazem de processo nao esta preenchido - MV_LOCPROC'
			EndIf
		Else
			cLocal:= SD3->D3_LOCAL
		EndIf
		SDA->(MsSeek(xFilial('SDA')+SD3->(D3_COD+cLocal+D3_NUMSEQ)))
		nQtdBx := RetSaldo(SD3->(xFilial("SD3")+D3_NUMSEQ),SDA->DA_SALDO)
		If Empty(nQtdBx)
			SD3->(DbSkip())
			Loop
		EndIf
		If nQtdBx > nSaldo
			nQtdBx :=nSaldo
		EndIf
		aadd(aNumSeq,{SD3->D3_NUMSEQ,nQtdBx,'SD3'})
		nSaldo -=nQtdBx
		If Empty(nSaldo)
			Exit
		EndIf
		SD3->(DbSkip())
	EndDo
EndIf
If nSaldo > 0
	aNumSeq :={}
EndIf
RestArea(aAreaSD1)
RestArea(aArea)
Return aNumSeq


Static Function RetSaldo(cChave,nSaldo)
//Local nSaldo := SDA->DA_SALDO
ITETMP->(DBSetOrder(2))
ITETMP->(MsSeek(cChave))
While  ITETMP->( !Eof() .and. ITE_FILIAL+ITE_NUMSEQ == cChave )
	nSaldo -= ITETMP->ITE_QTD
	ITETMP->(DbSkip())
End
ITETMP->(DBSetOrder(1))
Return nSaldo

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o	 ³ VldEndereco³ Autor ³ Sandro              ³ Data ³ 01/01/01 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³ Valida o endereco         								  ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso		 ³ ACDV060                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
Static Function VldEndereco(lEtiEnd)
Local aEnd := {}
Local aLoc := {}
Local cCbEndCQ := GetMV("MV_CBENDCQ")
Default lEtiEnd := .F.

If ExistBlock("ACD060VE")
	lOk := ExecBlock("ACD060VE",.F.,.F.)
	If (ValType(lOk) == "L")
		If !lOk
			VTClearGet()
			VTClearGet("cArmazem")
			VTGetSetFocus("cArmazem")
			Return .F.
		EndIf
	EndIf
EndIf

If Empty(aDist)
	VTALERT(STR0017,STR0015,.T.,4000)   //"Etiqueta invalida."###"AVISO"
	VTClearGet()
	If lEtiEnd
		VTClearGet("cEtiEnd")
		VTGetSetFocus("cEtiEnd")
	Else
		VTClearGet("cArmazem")
		VTGetSetFocus("cArmazem")
	EndIf
	Return .f.
EndIf

If lEtiEnd
	If UsaCb0("02") .and. ! Empty(cEtiEnd)
		aEnd := CBRetEti(cEtiEnd,'02')
		If Empty(aEnd)
			VTBEEP(2)
			VTALERT(STR0044,STR0015,.T.,4000)  //### //"Endereco invalido"###"AVISO"
			VTKeyBoard(chr(20))
			VTClearGet()
			VTClearGet("cEtiEnd")
			VTGetSetFocus("cEtiEnd")
			Return .f.
		EndIf
		cEndereco := aEnd[1]
		cArmazem  := aEnd[2]
	EndIf
EndIf

If UsaCb0("01") .and. ! Empty(aHisEti)
	aLoc := CBRetEti(aHisEti[1,1],'01')
	If !lCQ .and. !Empty(CB0->CB0_LOCAL) .and. CB0->CB0_LOCAL <> cArmazem
		VTBEEP(2)
		VTALERT(STR0044,STR0015,.T.,4000)   //"Endereco invalido"###"AVISO"
		VTClearGet()
		If lEtiEnd
			VTClearGet("cEtiEnd")
			VTGetSetFocus("cEtiEnd")
		Else
			VTClearGet("cArmazem")
			VTGetSetFocus("cArmazem")
		EndIf
		Return .f.
	EndIf
	If !Empty(aLoc) .And. !Empty(aLoc[9])
		VTBEEP(2)
		VTALERT(STR0018,STR0015,.T.,4000)  //### //"Produto ja foi lido."###"AVISO"
		Return .F.	
	EndIf
EndIf

SBE->(DbSetOrder(1))
If ! SBE->(MsSeek(xFilial('SBE')+cArmazem+cEndereco)) // BY ERIKE
	VTBEEP(2)
	VTALERT(STR0022,STR0015,.T.,4000)   //"Endereco nao encontrado"###"AVISO"
	VTClearGet()
	If lEtiEnd
		VTClearGet("cEtiEnd")
		VTGetSetFocus("cEtiEnd")
	Else
		VTClearGet("cArmazem")
		VTGetSetFocus("cArmazem")
	EndIf
	Return .f.
EndIf
If lCQ
	If !Empty(cCbEndCQ) .And. !(cArmazem+Alltrim(cEndereco)+";" $ cCbEndCQ)
		VTBEEP(2)
		VTALERT(STR0044,STR0015,.T.,4000)   //"AVISO" //"Endereco invalido"
		VTClearGet()
		If lEtiEnd
			VTClearGet("cEtiEnd")
			VTGetSetFocus("cEtiEnd")
		Else
			VTClearGet("cArmazem")
			VTGetSetFocus("cArmazem")
		EndIf
		Return .f.
	EndIf
EndIf
If ! CBEndLib(cArmazem,cEndereco)
	VTBEEP(2)
	VTALERT(STR0023,STR0015,.T.,4000)   //"Endereco bloqueado"###"AVISO"
	VTClearGet()
	If lEtiEnd
		VTClearGet("cEtiEnd")
		VTGetSetFocus("cEtiEnd")
	Else
		VTClearGet("cArmazem")
		VTGetSetFocus("cArmazem")
	EndIf
	Return .f.
EndIf
VTBEEP(2)

If VTYesNo(STR0024,STR0013,.T.)   //"Confirma etiqueta de endereco?"###"ATENCAO	
	If ExistBlock("ACD060CF")
		ExecBlock("ACD060CF",.F.,.F.)
	EndIf
	If UsaCB0("01")  
		CB0->(MsUnLockAll())
	EndIf 
	Distribui(cEndereco,cArmazem)

	VTKeyBoard(chr(20))
	cArmazem  := Space(Tamsx3("B1_LOCPAD")[1])
	cEndereco := Space(TamSx3("BF_LOCALIZ")[1])
	cNota     := Space(TamSx3("F1_DOC")[1])
	cSerie    := Space(SerieNfId("SF1",6,"F1_SERIE"))
	cFornec   := Space(TamSx3("F1_FORNECE")[1])
	cLoja     := Space(TamSx3("F1_LOJA")[1])
	cDoc      := Space(TamSx3("D3_DOC")[1])
	cLote     := Space(TamSx3("D3_LOTECTL")[1])
	cSLote    := Space(TamSx3("D3_NUMLOTE")[1])
	cNumSeri  := Space(TamSx3("BF_NUMSERI")[1])
	cArmEti   := Space(Tamsx3("B1_LOCPAD")[1])
	cEtiEnd   := Space(TamSx3("CB0_CODET2")[1])
	lBranco   := .t.
	Return .t.
EndIf
Return .f.
/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o	 ³VldEtiq     ³ Autor ³ Desenv. ACD         ³ Data ³ 17/04/01 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³ Validacao do produto lido na etiqueta                      ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Retorno   ³ Retl = Retorna .T. se validacao foi ok                     ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametro ³ ExpC1 = Codigo da etiqueta de produto                      ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso		 ³ SigaACD                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
Static Function VldEtiq(cTipoObr) //funcao utilizado quando usacb0 no produto e endereco
Local cTipId   := ""
Local aItensPallet
Local lIsPallet := .t.
DEFAULT cTipoObr := "01/02"
If Empty(cProd)
	Return .f.
EndIf
If ExistBlock("ACD060ET")
	ExecBlock("ACD060ET",,,{cProd})
EndIf
aItensPallet := CBItPallet(cProd)
If len(aItensPallet) == 0
	lIsPallet := .f.
EndIf

cTipId:=CBRetTipo(cProd)
If lIsPallet .or. (cTipId =="01"  .and. cTipId $ cTipoObr )
	VldProd()
ElseIf cTipId =="02" .and. cTipId $ cTipoObr
	aEtiqueta := CBRetEti(cProd,"02")
	If Empty(aEtiqueta)
		VTBEEP(2)
		VTALERT(STR0017,STR0015,.T.,4000)  //### //"Etiqueta invalida."###"AVISO"
		VTKeyBoard(chr(20))
		Return .f.
	EndIf
	cEndereco:=aEtiqueta[1]
	cArmazem :=aEtiqueta[2]
	Return VLDEndereco()
Else
	VTBEEP(2)
	VTALERT(STR0017,STR0015,.T.,4000)  //### //"Etiqueta invalida."###"AVISO"
EndIf
VTKeyBoard(chr(20))
Return .f.
/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o	 ³Distribui   ³ Autor ³ Desenv. ACD         ³ Data ³ 17/04/01 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³ Grava a distribuicao no Sistema                            ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametro ³                                                            ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso		 ³ SigaACD                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
Static Function Distribui(cLocaliz,cLocal)
Local cEtiq    := space(10)
Local cItem    := ""
Local cDoc     := ""
Local cNumSeq  := ""
Local cSeek    := ""
Local cLote    := ""
Local cSLote   := ""
Local nI       := 0
Local nX       := 0
Local nIndex   := 0
Local aCab     := {}
Local aItens   := {}
Local aSave := VTSAVE()

Private lMSErroAuto := .F.
If lCQ
	cArmazem := AlmoxCQ()
EndIf

VTMSG(STR0025,1)  // //"Aguarde..."

Begin Transaction

For nI := 1 to Len(aDist)    
	If aDist[nI,4] == "SDA"
		cNumSeq := aDist[nI,1]
		cLote   := aDist[nI,6] 
		cSLote  := aDist[nI,7] 

      SDA->(DbSetOrder(1))
      SDA->(DbSeek(xFilial('SDA')+aDist[nI,2]+cLocal+cNumSeq))
		cDoc    := SDA->DA_DOC
		
	Else
		dbSelectArea(aDist[nI,4])
		cSeek  := xFilial(aDist[nI,4])+aDist[nI,1]
		dbSetOrder(4)
		If ! MsSeek(cSeek)
			Loop
		EndIf
		If aDist[nI,4] == "SD1"
			cDoc    := SD1->D1_DOC
			cNumSeq := SD1->D1_NUMSEQ
			cLote   := SD1->D1_LOTECTL
			cSLote  := SD1->D1_NUMLOTE
		Else
			cDoc    := SD3->D3_DOC
			cNumSeq := SD3->D3_NUMSEQ
			cLote   := SD3->D3_LOTECTL
			cSLote  := SD3->D3_NUMLOTE
		EndIf
	EndIf	
	AbreSemaf(cNumSeq)
	cItem := Item(nI,cLocal,cLocaliz,cNumSeq)
	aCAB  :={	{"DA_PRODUTO",	aDist[nI,2],	nil},;
				{"DA_LOCAL",	cLocal,			nil},;
				{"DA_NUMSEQ",	cNumSeq,		nil},; //relacionado ao campo D1_NUMSEQ
				{"DA_DOC",		cDoc,			nil}}  //Relacionado ao campo F1_DOC ou D1_DOC
	
	aITENS:={{	{"DB_ITEM",		cItem,	   		nil},;
				{"DB_LOCALIZ",	cLocaliz,  		nil},;
				{"DB_QUANT",	aDist[nI,3],	nil},;
				{"DB_DATA",		dDATABASE,		nil}}}
	
	If ! Empty(aDist[nI,8])
		SB1->(DbSetOrder(1))
		SB1->(MsSeek(xFilial("SB1")+aDist[nI,2]))
		If SB1->B1_QTDSER == IIf( TamSx3( 'B1_QTDSER' )[3] == "C", "2", 2 )
			aadd(aItens[1],{"DB_QTSEGUM",1 , nil})
		EndIf
		aadd(aItens[1],{"DB_NUMSERI",aDist[nI,8]	,nil})
	EndIf
	nModuloOld  := nModulo
	nModulo     := 4
	lMSHelpAuto := .T.
	lMSErroAuto := .F.
	lItIguais	:= .F.
	aPrDup		:= {}
	SX3->(DbSetOrder(1))
	
	msExecAuto({|x,y|mata265(x,y)},aCab,aItens)
	nModulo := nModuloOld
	lMSHelpAuto := .F.
	FechaSemaf(cNumSeq)
	If lMSErroAuto
		VTBEEP(2)
		VTALERT(STR0026,STR0027,.T.,4000)  //### //"Falha no processo de distribuicao."###"ERRO"
		DisarmTransaction()
		Exit
	Else
		If UsaCB0("01")
			For nX := 1 to len(aDist[nI,5])
				If Empty(aDist[nI,8])
					CBGrvEti("01",{NIL,NIL,NIL,NIL,NIL,NIL,NIL,NIL,cLocaliz,cLocaL,,cNumSeq},aDist[nI,5,nX,3])
				Else
					CBGrvEti("01",{NIL,NIL,NIL,NIL,NIL,NIL,NIL,NIL,cLocaliz,cLocaL,,cNumSeq,NIL,NIL,NIL,NIL,NIL,NIL,NIL,NIL,NIL,NIL,aDist[nI,8]},aDist[nI,5,nX,3]) // Atualiza o numero de serie
				EndIf
				CBLog("01",{aDist[nI,2],CB0->CB0_QTDE,cLote,cSLote,cLocal,cLocaliz,cNumSeq,cDoc,CB0->CB0_CODETI})
			Next
		Else
			CBLog("01",{aDist[nI,2],aDist[nI,3],cLote,cSLote,cLocal,cLocaliz,cNumSeq,cDoc,""})
		EndIf
		RegistraCab(.t.)
	EndIf
Next

End Transaction

If ExistBlock("ACD060GR")
	ExecBlock("ACD060GR",.F.,.F.)
EndIf

If lMsErroAuto
	VTDispFile(NomeAutoLog(),.t.)
Else
	If ExistBlock("ACD060OK")
		ExecBlock("ACD060OK",.F.,.F.)
	EndIf
	//Limpa Variaveis para permitir novo endereçamento de nota .
	aDist  := {}
	aHisEti:= {}
Endif
VTCLEAR
VtRestore(,,,,aSave)
Return

Static Function Item(nPos,cLocal,Localiz,cNumSeq)
Local cItem     := ""
SDB->(dbSetOrder(1))
If SDB->(MsSeek(xFilial("SDB")+aDist[nPos,2]+cLocal+aDist[nPos,1]))
	While SDB->(!EOF() .and. xFilial("SDB")+aDist[nPos,2]+cLocal+aDist[nPos,1] ==;
		DB_FILIAL+DB_PRODUTO+DB_LOCAL+DB_NUMSEQ)
		cItem := SDB->DB_ITEM
		SDB->(dbSkip())
	end
	cItem := strzero(val(cItem)+1,4)
Else
	cItem := "0001"
EndIf
Return cItem

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o	 ³ Informa    ³ Autor ³ Desenv. ACD         ³ Data ³ 30/05/01 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³ Mostra produtos que ja foram lidos                         ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametro ³                                                            ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso		 ³ ACDV060                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
Static Function Informa()
Local aCab,aSize,aSave := VTSAVE()
Local nX,nPos
Local aTemp:={}
VTClear()
If  UsaCB0("01")
	aCab  := {STR0028,STR0029}  //"Etiqueta"###"Produto"
	aSize := {10,16}
	aTemp := aClone(aHisEti)
Else
	aCab  := {STR0029,STR0030,STR0039,STR0040,STR0066}  //"Produto"###"Quantidade"###"Lote"###"SubLote" ### "Serie"
	aSize := {15,12,10,7,20}
	aHisEti:= {}
	For nx:= 1 to len(aDist)
		nPos := Ascan(aTemp,{|x| x[1] == aDist[nx,2] .and. x[3] == aDist[nx,6] .and. x[4] == aDist[nx,7] .and. x[5] == aDist[nx,8] })
		IF nPos == 0
			aadd(aTemp,{aDist[nx,2],aDist[nX,3],aDist[nX,6],aDist[nX,7],aDist[nX,8]})
		Else
			aTemp[nPos,2] += aDist[nX,3]
		endIf
	Next
EndIf
VTaBrowse(0,0,3,19,aCab,aTemp,aSize) //verificar
VtRestore(,,,,aSave)
Return


Static Function Estorna()
Local aTela
Local cEtiqueta
aTela := VTSave()
VTClear()
cEtiqueta := Space(TamSx3("CB0_CODET2")[1])
nQtdePro  := 1
@ 00,00 VtSay Padc(STR0033,VTMaxCol())  // //"Estorno da Leitura"
If ! UsaCB0('01')
	@ 1,00 VTSAY  STR0034 //'Qtde. '
	@ 1,05 VTGet nQtdePro   pict CBPictQtde() valid nQtdePro > 0 when (lForcaQtd .or. VTLastkey() == 5)//'Qtde.' //
EndIf
@ 02,00 VtSay STR0035  // //"Etiqueta:"
@ 03,00 VtGet cEtiqueta pict "@!" Valid VldEstorno(cEtiqueta,nQtdePro)
VtRead
vtRestore(,,,,aTela)
Return

Static Function VldEstorno(cEtiqueta,nQtdePro)
Local nPos,cKey,nQtd,cProd,nPosID
Local nX,nP
Local aEtiqueta,nSaldo,nQtdeBx
Local cLote     := Space(TamSX3("B8_LOTECTL")[1])
Local cSLote    := Space(TamSX3("B8_NUMLOTE")[1])
Local cNumSeri  := Space(TamSX3("BF_NUMSERI")[1])
Local aDistBKP  := aClone(aDist)
Local aHisEtiBKP:= aClone(aHisEti)
Local aGrava    :={}
Local aItensPallet := CBItPallet(cEtiqueta)
Local lIsPallet:= .t.

If Empty(cEtiqueta)
	Return .f.
EndIF

If len(aItensPallet) == 0
	aItensPallet:={cEtiqueta}
	lIsPallet := .f.
EndIf


Begin Sequence
For nP:= 1 to len(aItensPallet)
	cEtiqueta :=  aItensPallet[nP]
	If UsaCB0("01")
		nPos := Ascan(aHisEti, {|x| AllTrim(x[1]) == AllTrim(cEtiqueta)})
		If nPos == 0
			VTBeep(2)
			VTALERT(STR0036,STR0015,.T.,4000)   //### //"Etiqueta nao encontrada"###"AVISO"
			break
		EndIf
		If ! lIsPallet .and. ! Empty(CB0->CB0_PALLET)
			VTBeep(2)
			VTALERT(STR0047,STR0015,.T.,4000)   //"AVISO" //"Etiqueta invalida, Produto pertence a um Pallet"
			break
		EndIf
	Else
		If ! CBLoad128(@cEtiqueta)
			break
		EndIf
		aEtiqueta := CBRetEtiEAN(cEtiqueta)
		IF Len(aEtiqueta) == 0
			VTBeep(2)
			VTALERT(STR0038,STR0015,.T.,4000)   //### //"Etiqueta invalida"###"AVISO"
			break
		EndIf
		cLote := aEtiqueta[3]
		If ! CBRastro(aEtiqueta[1],@cLote,@cSLote)
			VTBeep(2)
			VTALERT(STR0041,STR0015,.T.,4000)   //"AVISO" //"Lote invalido"
			break
		EndIf
		cNumseri := aEtiqueta[5]
		If CBChkSer(aEtiqueta[1]) .And. ! CBNumSer(@cNumseri,Nil,aEtiqueta)
			Break
		EndIf
	EndIf

	If UsaCB0("01")
		//Estorno do aHisEti
		cKey := aHisEti[nPos,3]
		cProd:= aHisEti[nPos,2]
		nQtd := CBRetEti(cEtiqueta,'01')[2]
		aDel(aHisEti,nPos)
		aSize(aHisEti,Len(aHisEti)-1)
		//Estorno do aDist
		nPos := aScan(aDist,{|x| AllTrim(x[1]) == Alltrim(cKey) .and. x[2] == cProd})
		aadd(aGrava,{xFilial(aDist[nPos,4]),aDist[nPos,1],nQtd*-1})
		aDist[nPos,3] := aDist[nPos,3] - nQtd
		If Empty(aDist[nPos,3])
			aDel(aDist,nPos)
			aSize(aDist,Len(aDist)-1)
		Else
			nPosID := Ascan(aDist[nPos,5],{|x| Alltrim(x[1]) == Alltrim(cEtiqueta)})
			aDel(aDist[nPos,5],nPosID)
			aSize(aDist[nPos,5],Len(aDist[nPos,5])-1)
		EndIf
	Else
		cProd  := aEtiqueta[1]
		nQtde  := aEtiqueta[2]
		nSaldo := 0
		For nx:= 1 to len(aDist)
			If ! (aDist[nX,2] == cProd .And. aDist[nX,6] == cLote .And. aDist[nX,7] == cSLote .And. aDist[nX,8] == cNumseri)
				Loop
			EndIf
			nSaldo += aDist[nX,3]
			If nSaldo >= (nQtde*nQtdePro)
				Exit
			EndIf
		Next
		If  nSaldo < (nQtde*nQtdePro)
			VTBeep(2)
			VTALERT(STR0042,STR0015,.T.,4000)    //"AVISO" //"Saldo insuficiente"
			break
		EndIf
		nSaldo := (nQtde*nQtdePro)
		nQtdeBx:= 0
		For nx:= 1 to len(aDist)
			If nX > Len(aDist)
				Exit
			EndIF
			If ! (aDist[nX,2] == cProd .And. aDist[nX,6] == cLote .And. aDist[nX,7] == cSLote .And. aDist[nX,8] == cNumseri)
				Loop
			EndIf
			If nSaldo ==0
				Exit
			EndIf
			If aDist[nx,3] <= nSaldo
				nQtdeBx := aDist[nx,3]
			Else
				nQtdeBx := nSaldo
			EndIf
			aadd(aGrava,{xFilial(aDist[nx,4]),aDist[nx,1],nQtdeBx*-1})
			aDist[nx,3] := aDist[nx,3] - nQtdeBx
			nSaldo -= nQtdeBx
			If Empty(aDist[nx,3])
				aDel(aDist,nx)
				aSize(aDist,Len(aDist)-1)
				nX--
				Loop
			EndIf
		Next
	EndIf
Next

If ! VTYesNo(STR0037,STR0013,.t.)  //### //"Confirma o estorno desta Etiqueta?"###"ATENCAO"
	Break
EndIf

For nX:= 1 to len(aGrava)
	GravaQtd(aGrava[nX,1],aGrava[nX,2],aGrava[nX,3])
Next
If ExistBlock("ACD060ES")
	ExecBlock("ACD060ES",.F.,.F.,{cEtiqueta,nQtdePro})
EndIf
If lForcaQtd
	VtClearGet("cEtiqueta")  // Limpa o get
	VtGetSetFocus('nQtdePro')
Else
	VtKeyboard(Chr(20))  // zera o get
EndIf
Return .f.
End Sequence
aDist  := aClone(aDistBKP)
aHisEti:= aClone(aHisEtiBKP)
nQtdePro := 1
VTGetRefresh("nQtdePro")
VTKeyBoard(chr(20))
Return .f.


VtKeyboard(Chr(20))  // zera o get
Return .f.


Static Function CriaFile(cArq,cAlias)
Local aStru
Local cDrive   := 'DBFCDX'
Local cArquivo := RetArq(cDrive,cArq,.T.)
Local cIndice  := RetArq(cDrive,cArq,.F.)
If cAlias =="CABTMP"
	aStru :={{"CAB_NUMRF"	,"C",3,00}}
Else
	aStru :={	{"ITE_RECNO",	"C",6,00},;
				{"ITE_FILIAL",	"C",2,00},;
				{"ITE_NUMSEQ",	"C",6,00},;
				{"ITE_QTD",		"N",12,4} }
EndIf
dbCreate(cArquivo,aStru,cDrive)
dbUseArea(.T.,cDrive,cArquivo,cAlias,.F.,.F.) // Exclusivo
If cAlias =="CABTMP"
	INDEX ON CAB_NUMRF TAG &(RetFileName(cIndice)) TO &(FileNoExt(cArquivo))
Else
	INDEX ON ITE_RECNO+ITE_FILIAL+ITE_NUMSEQ TAG &(cArq+"1") TO &(FileNoExt(cArquivo))
	INDEX ON ITE_FILIAL+ITE_NUMSEQ TAG &(cArq+"2") TO &(FileNoExt(cArquivo))
EndIf
dbCloseArea()
Return

Static Function CloseFile(cArq,cAlias)
Local cDrive   := 'DBFCDX'
Local cArquivo := RetArq(cDrive,cArq,.T.)
Local cIndice  := RetArq(cDrive,cArq,.F.)

dbUseArea(.T.,cDrive,cArquivo,cAlias,.F.,.F.)
If ! neterr()
	dbCloseArea()
	FErase(cArquivo)
	FErase(cIndice)
	Return .f.
EndIf
Return .t.

Static Function RegistraCab(lRegistra)
DEFAULT lRegistra:= .T.
CABTMP->(DbGotop())
ITETMP->(DbSetOrder(1))
While !  CABTMP->(eof())
	If ! CABTMP->(Rlock())
		CABTMP->(DbSkip())
		Loop
	EndIf
	While ITETMP->(MsSeek(Str(CABTMP->(Recno()),6)))
		RecLock("ITETMP",.f.)
		ITETMP->(DBDelete())
		ITETMP->(MsUnLock())
	End
	CABTMP->(DBDelete())
	CABTMP->(MsUnLock())
	CABTMP->(DbSkip())
End
//- Elimina os itens que estão sobrando sem cabecalho
ITETMP->(DbGotop())
While ITETMP->(!Eof())
	CABTMP->(MsGoto(Val(ITETMP->(ITE_RECNO))))
	If CABTMP->(DELETED())
		RecLock("ITETMP",.f.)
		ITETMP->(DBDelete())
		ITETMP->(MsUnLock())
	EndIf
	ITETMP->(dbSkip())
EndDo
If lRegistra
	RecLock("CABTMP",.t.)
	CABTMP->CAB_NUMRF := VTNUMRF()
	CABTMP->(MsUnlock())
	RecLock("CABTMP",.f.)
EndIf
Return .t.

Static function GravaQtd(cFilTmp,cNumseq,qtde)
IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf

ITETMP->(DbSetOrder(1))
If ! ITETMP->(MsSeek(Str(CABTMP->(Recno()),6)+cFilTmp+cNumseq))
	RecLock("ITETMP",.t.)
	ITETMP->ITE_RECNO:= Str(CABTMP->(Recno()),6)
	ITETMP->ITE_FILIAL := cFilTmp
	ITETMP->ITE_NUMSEQ:=cNumSeq
Else
	RecLock("ITETMP",.f.)
EndIf
ITETMP->ITE_QTD   += Qtde
ITETMP->(MsUnLock())
Return

Static Function AbreSemaf(cNumSeq)
Local nC    := 0
Local aSave := VTSAVE()
__nSem := -1
While __nSem  < 0
	__nSem  := MSFCreate(cNumSeq+".SEM")
	IF  __nSem  < 0
		SLeep(50)
		nC++
		If nC == 60
			nC := 0
			conout(STR0051+cNumSeq) //'Semaforo fechado para o numseq '
			VTCLear()
			If lVT100B
				@ 0,3 VTSay STR0052 //'   Aguarde...'
				@ 2,3 VTSay STR0053 //'   Preparando'
				@ 3,3 VTSay STR0054 //'enderecamento...'
			Else
				@ 2,3 VTSay STR0052 //'   Aguarde...'
				@ 4,3 VTSay STR0053 //'   Preparando'
				@ 5,3 VTSay STR0054 //'enderecamento...'
			Endif
		EndIf
	Endif
End
FWrite(__nSem,STR0051+cNumSeq) //'Semaforo fechado para o numseq '
VtRestore(,,,,aSave)
Return

Static Function FechaSemaf(cNumSeq)
Fclose(__nSem)
FErase(cNumSeq+".SEM")
Return

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³VldEndSer ºAutor  ³Aecio Ferreira Gomes     º Data ³  18/06/09           º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³Valida existencia do numero de Serie.                                    º±±
±±º          ³                                                                         º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ ACDV060                                                                 º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
Static Function VldEndSer(cNumSerie,aDist)
Local lRet   := .t. 

SBF->(dbSetOrder(4))
If SBF->(DbSeek(xFilial("SBF")+SB1->B1_COD+cNumSerie))
	Help(" ",1,"NUMSERIEEX")
	lRet:=.F.
EndIf

/* Verifica se ja foi feito a leitura do numero de serie */
If CBChkSer(SB1->B1_COD) .And. Ascan(aDist,{|x| x[2]+x[8] == SB1->B1_COD+cNumseri}) > 0
	VTAlert(STR0057,STR0015,.T.,3000)//##"Esse numero de serie ja foi lido para esse produto"##AVISO"
	lRet:=.F.
EndIf                                   

If !lRet
	VtClearGet("cNumSerie")
	VTGetSetFocus("cNumSerie")		
EndIf	
Return lRet

/*/{Protheus.doc} VldItIguais
Valida se a Nota Fiscal utilizada no endereçamento
possui itens com o mesmo codigo de produto.

@author robson.ribeiro
@since 21/09/2015
@version 1.0
@param	cNota	- Numero da NF
		cSerie	- Serie da NF
		cFornec	- Codigo do Fornecedor
		cLoja	- Loja do Fonecedor
/*/

Static Function VldItIguais(cNota,cSerie,cFornec,cLoja,cProdut)

Local aArea		:= GetArea()
Local aAreaSD1	:= SD1->(GetArea())
Local aItensEnd	:= {}
Local nProdIg		:= 0
Local nArmazem   	:= 0
Local nX			:= 0
Local nY         	:= 0	

Default cProdut  := ""

SD1->(DbSetOrder(1))
SD1->(MsSeek(xFilial("SD1")+cNota+cSerie+cFornec+cLoja))
While cNota+cSerie+cFornec+cLoja == SD1->(D1_DOC+D1_SERIE+D1_FORNECE+D1_LOJA); 
		.And. SD1->D1_COD == cProdut
		//Quantidade de protudos iguais 
			nProdIg++
			If  nProdIg > 1 
				lProdIg := .T.
			EndIf
			aAdd(aItensEnd,{SD1->D1_ITEM,SD1->D1_COD,SD1->D1_QUANT,SD1->D1_LOCAL,SD1->D1_LOTECTL})
			SD1->(dbSkip())
EndDo

If Len(aItensEnd) > 1
	For nX := 1 to Len(aItensEnd)
		For nY := 1 to Len (aItensEnd)
			If (aItensEnd[nX][4] <> aItensEnd[nY][4]) .And. (nX <> nY) .And. (aItensEnd[nX][2]== cProdut) .And. (aItensEnd[nY][2]== cProdut)
				nArmazem++
				// Localizado armazens distintos no documento de entrada
				lArmazem:= .T.
			EndIF
		Next
	Next	
EndIf	

lRet := lProdIg

RestArea(aAreaSD1)
RestArea(aArea)

Return lRet

/*/{Protheus.doc} InfArm
Solita armazem para localização de registros iguais 
@author André Maximo
@since 22/04/2016
@version 1.0
@param 
/*/

Static Function InfArm()  

Local aSave := VTSAVE()

IF !Empty(cArmazem)
	cArmazem:= ""
EndIF

VTCLear()
@ 0,0 VTSay STR0065 
@ 0,8 VTGet cArmazem PICTURE "@!" Valid VTLastkey() == 05 .or. ! Empty(cArmazem) .And. VldArm(cArmazem)
VTREAD
VtRestore(,,,,aSave)

Return



/*/{Protheus.doc} VldArm
Valida armazem selecionado pelo usuário  
@author André Maximo
@since 22/04/2016
@version 1.0
@param 
/*/

Static Function VldArm(cArm,cTipo)

If Empty(cArm)
	Return .t.
EndIf

SBE->(DbSetOrder(1))
If ! SBE->(DbSeek(xFilial("SBE")+cArm))
	VTBeep(3)
	VTAlert(STR0067,STR0065,.t.,4000) //'Armazem nao existe'###'Armazem'
	VTKeyBoard(chr(20))
	Return .f.
EndIf

Return .t.

/*/{Protheus.doc} FornecConf
Valida se o Fornecedor nao faz Conferencia fisica 
@author jefferson.sousa
@since 28/09/2020
@version 1.0
@param	cFornec	- Codigo do Fornecedor
@param	cLoja	- Codigo da loja
/*/
Static Function FornecConf(cFornec,cLoja)
Local aAreaSA2 := SA2->(GetArea())
SA2->(dbSetOrder(1))
If  SA2->(dbSeek(xFilial("SA2")+cFornec+cLoja)) 
	lForConf := IIF(SA2->A2_CONFFIS # '3', .T., .F.)
EndIf
RestArea(aAreaSA2)

Return lForConf
