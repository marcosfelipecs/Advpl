#INCLUDE "Acdv153.ch" 
#include "protheus.ch"
#INCLUDE 'APVT100.CH'


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o	 ³ acdv152    ³ Autor ³ Sandro              ³ Data ³ 14/08/01 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³ Programa de Retorno de processo     							  ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso		 ³ SigaACD           	    								           ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/   
Template function acdv153(uPar1,uPar2,uPar3,uPar4,uPar5,uPar6,uPar7,uPar8,uPar9,uPar10)
Return acdv153(uPar1,uPar2,uPar3,uPar4,uPar5,uPar6,uPar7,uPar8,uPar9,uPar10)

Function acdv153()
Local aKey        := array(2)           
Private cArmOri   := Space(TamSX3("B2_LOCAL")[1])
Private cArmDes   := GetMvNNR('MV_LOCPROC','99')
Private cCB0Prod  := Space(TamSx3("CB0_CODET2")[1])
Private cProduto  := IIf( FindFunction( 'AcdGTamETQ' ), AcdGTamETQ(), Space(48) )

Private nQtde     := 1
Private nQtdeProd := 1 
Private cLote     := Space(TamSX3("B8_LOTECTL")[1])
Private cSLote    := Space(TamSX3("B8_NUMLOTE")[1])
Private cNumSerie := Space(TamSX3("BF_NUMSERI")[1])
Private aLista    := {}
Private aHisEti   := {}
Private aSaldos   := {}
Private lMsErroAuto := .F.
Private nLin:= 0

akey[1] := VTSetKey(24,{|| Estorna()},STR0040) //"Estorno"
akey[2] := VTSetKey(09,{|| Informa()},STR0041) //"Informacoes"

While .t.                         
   VTClear      
   nLin:= -1
	@ ++nLin,0 VTSAY STR0042 //"Retorno Processo"
	If ! GetArmOri()
  		Exit
	EndIf
	GetProduto()
	VTRead                
   If len(aLista) == 0 
      Exit
   End    
	If ! GravaTransf()
	   If VTYesNo(STR0002,STR0003) //'Confirma a saida?'###'Atencao'
	      Exit                                                          
	   EndIf
	EndIf      
End
vtsetkey(24,akey[1])
vtsetkey(09,akey[2])
Return      

Static Function GetArmOri()
If ! UsaCB0('01')
	cArmOri:= SuperGetMv("MV_CBARMPD",.F.,cArmOri)
	If ExistBlock("ACD152ARM") 
		cArmOri := ExecbLock("ACD152ARM",.F.,.F.)
	EndIf
	cArmOri:= If(Empty(cArmOri),CriaVar("B2_LOCAL"),cArmOri)
	@ ++nLin,0 VtSay STR0004 //'Armazem origem'
	@ ++nLin,0 VTGet cArmOri pict '@!' Valid ! Empty(cArmOri)  when empty(cArmOri)
	VTRead()                                       
	If VTLastkey() == 27
	   Return .f.
	EndIf 
	VTClear(1,0,2,19)
	nLin := 0
EndIf

Return .t.   

Static Function GetProduto()
Local lVolta := .F.
IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf
If UsaCB0('01')
	@ ++nLin,0 VTSAY STR0005 //"Produto"
	@ ++nLin,0 VTGET cCB0Prod PICTURE "@!" valid VldProduto("01")
ElseIf ! UsaCB0('01')
	If lVT100B //GetMv("MV_RF4X20")
		While .T.
			lVolta := .F.
			nLin := 1
			@ ++nLin,0 VTSAY STR0006 //"Quantidade"
			@ ++nLin,0 VTGET nQtde PICTURE CBPictQtde() valid nQtde > 0 //when VTLastKey() == 5
			VTRead
			
			If !(vtLastKey() == 27)
				VTClear(1,0,3,19)
				nLin := 1
				@ ++nLin,0 VTSAY STR0005 //"Produto"
				@ ++nLin,0 VTGET cProduto    PICTURE "@!" valid VldProduto("") ;
					when iif(vtRow() == 3 .and. vtLastKey() == 5,(VTKeyBoard(chr(27)),lVolta := .T.),.T.)
				VTRead
			EndIf
			
			if lVolta
				Loop
			endif
				
			Exit
		EndDo
	Else
		@ ++nLin,0 VTSAY STR0006 //"Quantidade"
		@ ++nLin,0 VTGET nQtde PICTURE CBPictQtde() valid nQtde > 0 when VTLastKey() == 5
		@ ++nLin,0 VTSAY STR0005 //"Produto"
		@ ++nLin,0 VTGET cProduto    PICTURE "@!" valid VldProduto("")
	EndIF
EndIf
Return

Static Function VldProduto(cTipo)
Local cTipId
Local aEtiqueta:={}
Local nQE      
Local nQtdeLida:=0           
Local nSaldo        
Local nI 
Local aListaBKP 
Local aHisEtiBKP
Local aSaldosBKP
Local aItensPallet 
Local lIsPallet 
Local nP  
Local cEtiqueta
Local nTamLote   := TamSX3("B8_LOTECTL")[1]
Local nTamSLote  := TamSX3("B8_NUMLOTE")[1]
Local nTamSeri   := TamSX3("BF_NUMSERI")[1]

If "01" $ cTipo
	If Empty(cCB0Prod)
    	Return .f.
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
		aSaldosBKP:= aClone(aSaldos)
	
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
					break			
				EndIf
				If ! Empty(aEtiqueta[2]) .and. Ascan(aHisEti,cCB0Prod) > 0
					VTALERT(STR0010,STR0008 ,.T.,4000,3) //"Etiqueta ja lida"###"Aviso"
					Break
				EndIf                    
				If Localiza(aEtiqueta[1])
					VTALERT(STR0011,STR0008,.T.,4000,3) //"Produto lido controla endereco!"###"Aviso"
					VTALERT(STR0012,STR0008,.T.,4000) //"Utilize rotina especifica ACDV150"###"Aviso"
					Break
				EndIf 
				SB1->(DbSetOrder(1))
				SB1->(DbSeek(xFilial("SB1")+aEtiqueta[1]))       
				If SB1->B1_APROPRI <> "I"
					VTALERT(STR0043,STR0008,.T.,4000)  //"Produto nao utiliza processo"###"Aviso"
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
				EndIF		
	            cProduto  := aEtiqueta[1]
				nQE:= 1                
				nQE := CBPedQtd(0)                   
				If empty(nQE)
					Break
				EndIf
		        nQtdeProd := nQE
	            nQtdeLida := 0
		        aEval(aLista,{|x|  If(x[1]+x[4]+x[5]+x[6]==cProduto+cLote+cSLote+cNumSerie,nQtdeLida+=x[2],nil)}) 
		  		SB2->(DbSetOrder(1))
		   		SB2->(DbSeek(xFilial()+cProduto+cArmDes)) //asv
		   		nSaldo := SaldoSB2(,.F.)
//		    	   nSaldo := SB2->B2_QATU
	            IF nQtdeProd+nQtdeLida >  nSaldo
		    		VTALERT(STR0013,STR0008 ,.T.,4000,3) //"Quantidade excede o saldo disponivel"###"Aviso"
					Break
		        EndIf               
		        If ExistBlock("AV153VPR") 
		        	cEtiqueta:= cCB0prod
					If ! ExecBlock("AV153VPR",.F.,.F.,cEtiqueta)
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
		aSaldos:= aClone(aSaldosBKP) 
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
	EndIF	
	nQE:= 1                 
	//nQE := CBPedQtd(0)                   
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
	nQtdeLida := 0
	aEval(aLista,{|x|  If(x[1]+x[4]+x[5]+x[6]==cProduto+cLote+cSLote+cNumSerie,nQtdeLida+=x[2],nil)}) 
	SB2->(DbSetOrder(1))
	SB2->(DbSeek(xFilial()+cProduto+cArmDes)) //asv   
	nSaldo := SaldoSB2(,.F.)
//     nSaldo := SB2->B2_QATU
 	IF nQtdeProd+nQtdeLida > nSaldo
		VTALERT(STR0013,STR0008 ,.T.,4000,3) //"Quantidade excede o saldo disponivel"###"Aviso"
		cLote     := Space(nTamLote)
		cSLote    := Space(nTamSLote)
		cNumSerie := Space(nTamSeri)
		VTKeyboard(chr(20))
		Return .f.
	EndIf
  	If ExistBlock("AV153VPR") 
		cEtiqueta:= cProduto  	
		If ! ExecBlock("AV153VPR",.F.,.F.,cEtiqueta)
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
	Return .f.
EndIf                           
Return .f.
      


Static Function TrataArray(cEtiqueta,lEstorno)
Local nPos 
Default lEstorno := .f.
If ! lEstorno                      
   If cEtiqueta <> NIL
		aadd(aHisEti,cEtiqueta)
		aadd(aSaldos,{cEtiqueta,nQtdeProd})
	EndIf	
	nPos := aScan(aLista,{|x| x[1]+x[3]+x[4]+x[5]+x[6] == cProduto+cArmOri+cLote+cSLote+cNumSerie})
	If Empty(nPos)
	   aadd(aLista,{cProduto,nQtdeProd,cArmOri,cLote,cSLote,cNumSerie})
	Else    
	   aLista[nPos,2]+=nQtdeProd
	EndIf 
Else    
	nPos := aScan(aLista,{|x| x[1]+x[3]+x[4]+x[5]+x[6] == cProduto+cArmOri+cLote+cSLote+cNumSerie})
	aLista[nPos,2] -= nQtdeProd
	If Empty(aLista[nPos,2])
		aDel(aLista,nPos)
		aSize(aLista,len(aLista)-1)
	EndIf		
   If cEtiqueta <> NIL
		nPos := aScan(aHisEti,cEtiqueta)
		aDel(aHisEti,nPos)
		aSize(aHisEti,len(aHisEti)-1)    

		nPos := aScan(aSaldos,{|x| x[1]==cEtiqueta})
		aDel(aSaldos,nPos)
		aSize(aSaldos,len(aSaldos)-1)    
	EndIf	                          
EndIf
Return 

Static Function GravaTransf()
Local aSave
Local nI,nX
Local aCab:={}
Local aItens:={}
Local aEtiqueta             
Local cArmOri2   := ""
Local dValid
Local cTm := GetMV("MV_TMCBDP")
Local nTamLoc   := TamSX3("B2_LOCAL")[1]
Local nTamLote  := TamSX3("B8_LOTECTL")[1]
Local nTamSLote := TamSX3("B8_NUMLOTE")[1]
Local nTamSeri  := TamSX3("BF_NUMSERI")[1]
Local nTamEnd   := TamSX3("BF_LOCALIZ")[1]
Private nModulo := 4    
    
If Empty(cTm) .or. cTM > "500"
	VTALERT(STR0044,STR0008 ,.T.,4000,3) //"Aviso"    //"Parametro MV_TMCBDP Invalido"
	Return .f.
EndIf	  
SF5->(DbSetOrder(1))
If ! SF5->(DbSeek(xFilial('SF5')+cTM))
	VTALERT(STR0045,STR0008 ,.T.,4000,5) //"Aviso"    //"Tipo de movimento referente ao parametro MV_TMCBDP nao cadastrado"
	Return .f.
EndIf
If SF5->F5_TIPO<>"D" 
	VTALERT(STR0046,STR0008 ,.T.,4000,5) //"Aviso"    //"Tipo de movimento referente ao parametro MV_TMCBDP nao eh de DEVOLUCAO"
	Return .f.
EndIf
If SF5->F5_APROPR<>"N"  
	VTALERT(STR0047,STR0008 ,.T.,4000,5) //"Aviso"    //"Tipo de movimento nao pode ser de apropriacao INDIRETA referente parametro MV_TMCBDP"
	Return .f.
EndIf

If ! VTYesNo(STR0048,STR0008 ,.T.)  //"Confirma o Retorno de processo"###"Aviso"
	Return .f.
EndIf	                
  
aSave     := VTSAVE()
VTClear()
VTMsg(STR0022) //'Aguarde...'
Begin Transaction
	lMsErroAuto := .F.
	lMsHelpAuto := .T.                      
	
  	aCab := {{"D3_DOC"	   ,ProxDoc()			,NIL},;
            {"D3_TM"    	,cTM		     		,NIL},;
            {"D3_EMISSAO"	,dDataBase			,Nil}} 

	For nI := 1 to Len(aLista)
		SB1->(dbSeek(xFilial()+aLista[nI,1]))
		If !Rastro(SB1->B1_COD)
			aadd(aItens,{	{"D3_COD"		 ,SB1->B1_COD ,NIL},;
							{"D3_LOCAL"    ,aLista[nI,3],NIL},;
            		  		{"D3_QUANT"	 ,aLista[nI,2],NIL}})
       Else
			aadd(aItens,{	{"D3_COD"      ,SB1->B1_COD ,NIL},;
							{"D3_LOCAL"  	 ,aLista[nI,3],NIL},;
	           		  	{"D3_QUANT"    ,aLista[nI,2],NIL},;
	           		  	{"D3_LOTECTL"  ,aLista[nI,4],NIL},;
	           		  	{"D3_NUMLOTE"  ,aLista[nI,5],NIL}})
	   EndIf	
		If ! UsaCB0("01")
			CBLog("02",{SB1->B1_COD,aLista[nI,2],aLista[nI,4],aLista[nI,5],aLista[nI,3],,cArmDes})
		EndIf						
   Next
	MSExecAuto({|x,y| MATA241(x,y)},aCab,aItens)
	If lMsErroAuto
		VTALERT(STR0023,STR0024,.T.,4000,3) //"Falha na gravacao da transferencia"###"ERRO"
		DisarmTransaction()
		Break
	EndIf
	If UsaCb0("01") .and. Len(aHiseti) > 0    
  	   For nx:= 1 to len(aHisEti)
     	   aEtiqueta := CBRetEti(aHisEti[nX],"01")
     	   cArmOri2   := aEtiqueta[10]
     	   aEtiqueta[2]  := aEtiqueta[2]+ aSaldos[nX,2]
     	   aEtiqueta[12] := SD3->D3_NUMSEQ
			//aEtiqueta[10] := cArmDes
			If CBProdUnit(aHisEti[nX])
	  	      CBGrvEti("01",aEtiqueta,aHisEti[nX])
	  	   EndIf   
		   CBLog("02",{CB0->CB0_CODPRO,CB0->CB0_QTDE,CB0->CB0_LOTE,CB0->CB0_SLOTE,cArmDes,,cArmOri2,,CB0->CB0_CODETI})
     	Next
	Else
	   
   EndIf                
   If ExistBlock("ACD153GR")
      ExecBlock("ACD153GR",.F.,.F.)
   EndIf                 
End Transaction
VtRestore(,,,,aSave)
If lMsErroAuto
	VTDispFile(NomeAutoLog(),.t.)
Else
   If ExistBlock("ACD153OK")
      ExecBlock("ACD153OK",.F.,.F.)
   EndIf
	cArmOri     := Space(nTamLoc)
	cEndOri     := Space(nTamEnd)
	cCB0ArmOri  := Space(20)
	cCB0Prod    := Space(20)
	cProduto    := IIf( FindFunction( 'AcdGTamETQ' ), AcdGTamETQ(), Space(48) )
	cLote       := Space(nTamLote)
	cSLote      := Space(nTamSLote)
	cNumSerie   := Space(nTamSeri)	
	nQtde       := 1
	aLista      := {}                     
	aHisEti     := {}                                    
	aSaldos     := {}
Endif
Return .t.


Static Function Informa()
Local aCab  := {STR0005,STR0006,STR0027,STR0028,STR0029,STR0030} //"Produto"###"Quantidade"###"Armazem"###"Lote"###"SubLote"###"Num.Serie"
Local aSize := {15,16,7,10,7,20}
Local aSave := VTSAVE()
VtClear()
VTaBrowse(0,0,7,19,aCab,aLista,aSize)
VtRestore(,,,,aSave)
Return

Static Function Estorna()
Local aTela        
Local cEtiqueta          
Local nQtde := 1
aTela := VTSave()
VTClear()                       
cEtiqueta := IIf( FindFunction( 'AcdGTamETQ' ), AcdGTamETQ(), Space(48) )
@ 00,00 VtSay Padc(STR0031,VTMaxCol()) //"Estorno da Leitura"
If ! UsaCB0('01')
	@ 1,00 VTSAY  STR0032 VTGet nQtde   pict CBPictQtde() when VTLastkey() == 5 //'Qtde.'
EndIf	
@ 02,00 VtSay STR0033 //"Etiqueta:"
@ 03,00 VtGet cEtiqueta pict "@!" Valid VldEstorno(cEtiqueta,@nQtde)
VtRead                                                       
vtRestore(,,,,aTela)
Return 

Static Function VldEstorno(cEtiqueta,nQtde)
Local nPos
Local aEtiqueta,nQE
Local aListaBKP := aClone(aLista)
Local aHisEtiBKP:= aClone(aHisEti)
Local aItensPallet := CBItPallet(cEtiqueta)
Local lIsPallet := .t.
Local nP
Local nTamSeri  := TamSX3("BF_NUMSERI")[1]

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
				break			
			EndIf
		Else
			If ! CBLoad128(@cEtiqueta)
				Return .f.
			EndIf                
			aEtiqueta := CBRetEtiEAN(cEtiqueta) 
		   IF Len(aEtiqueta) == 0
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
		nQE := CBPedQtd(0)                   
		nQtdeProd:=Qtde*nQE
		
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


Static Function CBPedQtd(nQE)
Local aSave                                                   
Local nQAux := nQE
IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf
DEFAULT nQE := 1     
aSave   := VTSAVE()
VTClear                   
@ IIF(lVT100B /*GetMv("MV_RF4X20")*/,0,1),0 VTSay STR0049 //"Produto Variavel?"
@ IIF(lVT100B /*GetMv("MV_RF4X20")*/,2,3),0 VtSay STR0006  //"Quantidade"
If !UsaCB0("01")
	@ IIF(lVT100B /*GetMv("MV_RF4X20")*/,3,4),0 VtGet nQE pict CBPictQtde() valid nQE > 0
Else
	@ IIF(lVT100B /*GetMv("MV_RF4X20")*/,3,4),0 VtGet nQE pict PesqPict("CB0","CB0_QTDE") valid nQE > 0
EndIf
VTREAD
VtRestore(,,,,aSave)	             
If VTLastKey() == 27 
   VTAlert(STR0050,STR0008,.t.,3000) //"Quantidade Invalida"###"Aviso"
  	nQE := 0
EndIf	        
Return nQE

Static Function ProxDoc()
Local aSvAlias   := GetArea()
Local aSvAliasD3 := SD3->(GetArea())
Local cDoc
dbSelectArea("SD3")
dbSetOrder(2)                                               
While .t.
	cDoc :=  NextNumero("SD3",2,"D3_DOC",.T.)
   If ! dbSeek(xFilial()+cDoc)
      exit
	EndIf
End     
RestArea(aSvAliasD3)
RestArea(aSvAlias)
Return cDoc

