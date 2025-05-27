#INCLUDE "acdv160.ch" 
#include "protheus.ch"
#include "apvt100.ch"


/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ACDV160   ³ Autor ³ Sandro                ³ Data ³ 02/01/02 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Desmonta embalagem                                         ³±±
±±³          ³                                                            ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ Generico                                                   ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÁÄÄÄÂÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ PROGRAMADOR  ³ DATA   ³ BOPS ³  MOTIVO DA ALTERACAO                   ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³              ³        ³      ³                                        ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
/*/    
Template function ACDV160(uPar1,uPar2,uPar3,uPar4,uPar5,uPar6,uPar7,uPar8,uPar9,uPar10)
Return ACDV160(uPar1,uPar2,uPar3,uPar4,uPar5,uPar6,uPar7,uPar8,uPar9,uPar10)

Function ACDV160()
Local cEtiq 

While .t.            
   cEtiq := Space(TamSx3("CB0_CODET2")[1])
 	VTClear()               
   @ 1,0 VTSay STR0001 //"Desmonta Embalagem"
   @ 2,0 VTSay STR0002 //"Etiqueta :"
   @ 3,0 VTGet cEtiq pict '@!' VALID VldEti(cEtiq)
   VTRead                      
   IF Empty(cEtiq) .or. VTLastKey() == 27
      Exit
   EndIf
Enddo
Return            

Static Function VldEti(cEtiq)
Local   aEtiqueta 
Local   nQE
Local   nRecnoCB0
Private aLog:={}
aEtiqueta := CBRetEti(cEtiq,"01")

If Empty(aEtiqueta)
   VTBeep(3)
   VTAlert(STR0003,STR0004,.t.,2000) //"Etiqueta invalida"###"Aviso"
   Return .t.
EndIf
If !CB0->(Rlock())
   VTBeep(3)
   VTAlert(STR0010,STR0004,.t.,2000) //"Registro em uso por outro usuario!"###"Aviso"
   Return .f.
EndIf
If ! CBProdUnit(aEtiqueta[1])
   VTBeep(3)
   VTAlert(STR0005,STR0004,.t.,4000) //"Produto a granel, etiqueta invalida"###"Aviso"
   Return .t.
EndIf
If ! CBProdLib(aEtiqueta[10],aEtiqueta[1])
	VTKeyBoard(chr(20))
	Return .f.
EndIf
nQE       :=CBQtdEmb(aEtiqueta[1])
If  aEtiqueta[2] == nQE
   VTBeep(3)
   VTAlert(STR0006,STR0004,.t.,2000) //"Embalagem ja desmontada"###"Aviso"
   Return .t.
EndIf    
nRecnoCB0:=CB0->(Recno())

If VTYesNo(STR0007,STR0004,.t.)  //"Confirma a impressao?"###"Aviso"
   Imprime(aEtiqueta)
   CB0->(DbGoto(nRecnoCB0))
   CB0->(MsUnlock())
Else
   If  VTYesNo(STR0011,STR0004,.t.)  //"Aviso" //"Registra Etiqueta"
	   RegistraLeitura(aEtiqueta) 
   EndIf
   CB0->(DbGoto(nRecnoCB0))
   CB0->(MsUnlock())
   Return .t.
EndIf   
Return



Static Function Imprime(aEtiPai)
Local nQE,nCopias
Local nX
Local cCodEtiPai:= CB0->CB0_CODETI
If ! CB5SetImp(CBRLocImp("MV_IACD02"),IsTelNet())
   VTAlert(STR0009,STR0004,.t.,2000) //"Codigo do tipo de impressao invalido"###"Aviso"
   Return .t.
EndIf                   

VtMsg(STR0008) //"Imprimindo"

nQE:=CBQtdEmb(aEtiPai[1]) 
If ! Empty(SB1->B1_SEGUM) .and. aEtiPai[2] # SB1->B1_CONV
   nQE:=SB1->B1_CONV    
EndIf
nCopias:= aEtiPai[2]/nQE 

CB0->(ExecBlock("IMG01",,,{nQE,,,nCopias,CB0_NFENT,CB0_SERIEE,CB0_FORNEC,CB0_LOJAFO,CB0_LOCAL,CB0_OP,CB0_NUMSEQ,CB0_LOTE,CB0_SLOTE,CB0_DTVLD,,,,,,CB0_LOCALI}))
/*
	For nX:=1 to nCopias
		CB0->(ExecBlock("IMG01",,,{nQE,,,1,CB0_NFENT,CB0_SERIEE,CB0_FORNEC,CB0_LOJAFO,CB0_LOCAL,CB0_OP,CB0_NUMSEQ,CB0_LOTE,CB0_SLOTE,CB0_DTVLD,CB0_LOCALI}))
		CB0->(aadd(aLog,{CB0_CODPRO,CB0_QTDE,CB0_LOTE,CB0_SLOTE,CB0_NFENT,CB0_SERIEE,CB0_FORNEC,CB0_LOJAFO,CB0_LOCAL,CB0_LOCALI,CB0_OP,CB0_CODETI,cCodEtiPai,CB0_NUMSEQ}))
	Next
*/

CBRetEti(cCodEtiPai,"01")
RecLock('CB0',.F.,.T.)
CB0->(dbDelete())
CB0->(MsUnlock())
If ExistBlock('IMG00')
   ExecBlock("IMG00",,,{"RACDI010PR"})
EndIf  
MSCBClosePrinter()
For nX:= 1 to Len(aLog)
	CbLog("11",aLog[nX])
Next
Return


Static Function RegistraLeitura(aEtiPai)
Local aTela:= VTSave()
Local cEtiqueta:= Space(TamSx3("CB0_CODET2")[1])
Local aHisEti:={}                                                    
Local cCodEtiPai:=CB0->CB0_CODETI
Local bkey09 := VTSetKey(09,{|| Informa(aHisEti)},STR0012) //"Informacoes"
Local bKey24 := VTSetKey(24,{|| Estorna(aHisEti)},STR0013)   // CTRL+X //"Estorno"
IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf

While .t.
	VTClear()               
	If lVT100B // GetMv("MV_RF4X20")
		@ 0,0 VTSay STR0014 //"Registro de Etiqueta"
		@ 1,0 VTSay STR0002 //"Etiqueta :"
		@ 2,0 VTGet cEtiqueta pict '@!' VALID VldEtiReg(cEtiqueta,aHisEti,aEtiPai,cCodEtiPai)
		@ 3,0 VTSay STR0015 //"Tecle ESC p/ Sair"
	Else
		@ 1,0 VTSay STR0014 //"Registro de Etiqueta"
		@ 2,0 VTSay STR0002 //"Etiqueta :"
		@ 3,0 VTGet cEtiqueta pict '@!' VALID VldEtiReg(cEtiqueta,aHisEti,aEtiPai,cCodEtiPai)
		@ 7,0 VTSay STR0015 //"Tecle ESC p/ Sair"
	EndIf
	VTRead                      
	If ! Empty(aHisEti) .and. ! VTYesNo(STR0016,STR0004,.t.)  //"Confirma a saida"###"Aviso"
	   Loop
	EndIf                 
	If ! VTYesNo(STR0017,STR0004,.t.) 	 //"Confirma o registro das etiquetas lidas"###"Aviso"
	   Exit
   EndIf	   
   If analisa(aHisEti,aEtiPai,cCodEtiPai)  
		Exit   
	EndIf 
EndDo	
VtRestore(,,,,aTela)

vtsetkey(09,bkey09)
vtsetkey(24,bkey24)
Return  

Static Function VldEtiReg(cEtiqueta,aHisEti,aEtiPai,cCodEtiPai)
Local aEtiqueta:={}
Local nPos:=0
Local nQE
Local nTot:= 0                 
Local nX:=0
If Empty(cEtiqueta)
   Return .f.
EndIf                            

aEtiqueta:= CBRetEti(cEtiqueta,"01")
If ! Empty(aEtiqueta) .and. ! Empty(CB0->CB0_NFENT)
	VtBeep(3)
	VtAlert(STR0018,STR0004,.t.,5000)  //"Processo invalido, Etiqueta ja registrada"###"Aviso"
	VtKeyboard(Chr(20))  // zera o get
	Return .F.
EndIf          

If ExistBlock("ACD160RE")
	cEtiqueta := ExecBlock("ACD160RE",,,{cEtiqueta,cCodEtiPai})
EndIf

If CBRetTipo(cEtiqueta) <>"01"
	VtBeep(3)
	VtAlert(STR0003,STR0004,.t.,3000)  //"Etiqueta invalida"###"Aviso"
	VtKeyboard(Chr(20))  // zera o get
	Return .F.
EndIf

nPos := Ascan(aHisEti,{|x| x[1] == cEtiqueta})
If ! Empty(nPos)
	VtBeep(3)
	VtAlert(STR0019,STR0004,.t.,4000)  //"Etiqueta ja lida"###"Aviso"
	VtKeyboard(Chr(20))  // zera o get
	Return .F.
EndIf
nQE:=CBQtdEmb(aEtiPai[1]) 
If ! Empty(SB1->B1_SEGUM) .and. aEtiPai[2] # SB1->B1_CONV
   nQE:= SB1->B1_CONV
EndIf   

For nX:=1 to len(aHisEti)
   nTot+=aHisEti[nX,2]
Next                  
If (nTot+nQE) > aEtiPai[2] // qtde original do pai
	VtBeep(3)
	VtAlert(STR0020,STR0021,.t.,6000)  //"Quantidade dos itens maior da Etiqueta Mestre"###"Inconsistencia"
	VtKeyboard(Chr(20))  // zera o get
	Return .F.
EndIf
If ExistBlock("ACD160VET") .and. ! ExecBlock("ACD160VET",.F.,.F.,{aEtiPai,cEtiqueta})
	VtKeyboard(Chr(20))  // zera o get
	Return .F.
EndIf                 

aadd(aHisEti,{cEtiqueta,nQE})
VtKeyboard(Chr(20))  // zera o get
Return .f.

Static Function Analisa(aHisEti,aEtiPai,cCodEtiPai)
Local nX:=0
Local nTot:= 0                 
Local _aEtiqueta := CBRetEti(cCodEtiPai)

For nX:=1 to len(aHisEti)
   nTot+=aHisEti[nX,2]
Next                  
If nTot > _aEtiqueta[2] // qtde original do pai
	VtAlert(STR0022,STR0021,.t.,6000,3)  //"Quantidade dos itens esta maior que a Etiqueta Mestre"###"Inconsistencia"
	Return .F.
ElseIf nTot < _aEtiqueta[2] // qtde original do pai
	VtAlert(STR0023,STR0021,.t.,6000,3)  //"Quantidade dos itens esta menor que a Etiqueta Mestre"###"Inconsistencia"
	Return .F.
EndIf
For nX:= 1 to len(aHisEti)
   _aEtiqueta[2]:= aHisEti[nX,2]    
   If Len(Alltrim(aHisEti[nX,1])) == TamSx3("CB0_CODETI")[1]
      CBGrvEti("01",aClone(_aEtiqueta),aHisEti[nX,1])
		CBRetEti(aHisEti[nX,1],"01")
	Else                    
	   _aEtiqueta[15] := aHisEti[nX,1]
      CBGrvEti("01",aClone(_aEtiqueta),aHisEti[nX,1])
	EndIf      
	CB0->(aadd(aLog,{CB0_CODPRO,CB0_QTDE,CB0_LOTE,CB0_SLOTE,CB0_NFENT,CB0_SERIEE,CB0_FORNEC,CB0_LOJAFO,CB0_LOCAL,CB0_LOCALI,CB0_OP,CB0_CODETI,cCodEtiPai,CB0_NUMSEQ}))   
Next
CBRetEti(cCodEtiPai,"01")
RecLock('CB0',.F.,.T.)
CB0->(dbDelete())
CB0->(MsUnlock())
For nX:= 1 to Len(aLog)
	CbLog("11",aLog[nX])
Next
Return .T.

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
±±³ Uso		 ³ ACDV060           	    								           ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
Static Function Informa(aHisEti)
Local aCab,aSize,aSave := VTSAVE()
Local aTemp:={}
VTClear()                         
aCab  := {STR0024,STR0025} //"Etiqueta"###"Quantidade"
aSize := {20,16}
aTemp := aClone(aHisEti)
VTaBrowse(0,0,VTMaxRow(),VTMaxCol(),aCab,aTemp,aSize)
VtRestore(,,,,aSave)
Return                            

                                                                
Static Function Estorna(aHisEti)
Local aTela        
Local cEtiqueta                         
aTela := VTSave()
VTClear()                       
cEtiqueta := Space(20)              
@ 00,00 VtSay Padc(STR0026,VTMaxCol())   //"Estorno da Leitura"
@ 02,00 VtSay STR0027 //"Etiqueta:"
@ 03,00 VtGet cEtiqueta pict "@!" Valid VldEstorno(cEtiqueta,aHisEti)
VtRead                                                       
vtRestore(,,,,aTela)
Return 


Static Function VldEstorno(cEtiqueta,aHisEti)
Local nPos
If Empty(cEtiqueta)
   Return .f.
EndIF
nPos := Ascan(aHisEti, {|x| AllTrim(x[1]) == AllTrim(cEtiqueta)})
If nPos == 0
  	VTBeep(2)
	VTALERT(STR0028,STR0004,.T.,4000)    //"Etiqueta nao encontrada"###"AVISO"
	VTKeyBoard(chr(20))          
	Return .f.
EndIf      
aDel(aHisEti,nPos)
aSize(aHisEti,Len(aHisEti)-1)   
VTKeyBoard(chr(20))                     
Return .f.

