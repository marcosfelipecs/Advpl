#INCLUDE "ACDV240.ch" 
#INCLUDE "protheus.ch"
#INCLUDE "apvt100.ch"


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o	 ³ ACDV240    ³ Autor ³ Desenv.    ACD      ³ Data ³ 24/02/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³ Extracao de etiquetas do Pallet                            ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso		 ³ SigaACD           	    								           ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
Template function ACDV240(uPar1,uPar2,uPar3,uPar4,uPar5,uPar6,uPar7,uPar8,uPar9,uPar10)
Return ACDV240(uPar1,uPar2,uPar3,uPar4,uPar5,uPar6,uPar7,uPar8,uPar9,uPar10)

Function ACDV240()
Local nOpc
VTCLear()
@ 0,0 VTSay STR0025 //'Selecione:'
nOpc:=VTaChoice(2,0,3,VTMaxCol(),{STR0026,STR0027})  //"Inclusao"###"Exclusao"
If nOpc == 1 // Inclusao itens no pallet
   ACDV241()  
ElseIf nOpc == 2 //  Exclusao itens no pallet  
   ACDV242()  
EndIf   
Return       

Template function ACDV241(uPar1,uPar2,uPar3,uPar4,uPar5,uPar6,uPar7,uPar8,uPar9,uPar10)
Return ACDV241(uPar1,uPar2,uPar3,uPar4,uPar5,uPar6,uPar7,uPar8,uPar9,uPar10)
Template function ACDV242(uPar1,uPar2,uPar3,uPar4,uPar5,uPar6,uPar7,uPar8,uPar9,uPar10)
Return ACDV242(uPar1,uPar2,uPar3,uPar4,uPar5,uPar6,uPar7,uPar8,uPar9,uPar10)

Function ACDV241()
   ACDV240X(.t.)
Return
Function ACDV242()
   ACDV240X(.f.)
Return


Static Function ACDV240X(lInclusao)
Local   bkey09 := VTSetKey(09,{|| ACD240Hist()},STR0028)// CTRL+I //"Informacoes"
Local   bKey24 := VTSetKey(24,{|| Estorna()},STR0029)   // CTRL+X //"Estorno"
Local   cTexto := " "
Local   aTela  := {}
Local   nLin   :=0
Local 	 lVolta  := .F.
Private aHisEti:= {}
Private cEti   := If(UsaCB0("01"),Space(TamSX3("CB0_CODET2")[1]), IIf( FindFunction( 'AcdGTamETQ' ), AcdGTamETQ(), Space(48) ) )			
Private cPallet:= If(UsaCB0("01"),Space(TamSX3("CB0_CODET2")[1]), IIf( FindFunction( 'AcdGTamETQ' ), AcdGTamETQ(), Space(48) ) )	
IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf		
VTClear()            


While .T.   
	nLin   :=0	
	@ nLin,0 VTSay STR0030 //"Manutencao de Pallet"
	@ ++nLin,0 VTSay If(lInclusao,STR0031	,STR0027) //"Inclusao:"###"Exclusao"
   @ ++nLin,0 VTSay STR0032   //"Pallet"
	@ ++nLin,0 VTGet cPallet pict '@!' Valid VldPallet(cPallet,lInclusao) When IIF(lVT100B /*GetMV("MV_RF4X20")*/,.T.,Empty(cPallet))
	If lVT100B // GetMv("MV_RF4X20")
		VTRead
		lVolta := .F.
		If VTLastKey() != 27
			VTClear
			nLin := 0
			@ nLin,0 VTSay STR0019   //"Produto"
			@ ++nLin,0 VTGet cEti pict '@!' Valid VldEti(cEti,lInclusao) .and. ! Empty(aHisEti) ;
				when iif(vtRow() == 1 .and. vtLastKey() == 5,(VTKeyBoard(chr(27)),lVolta := .T.),.T.)
		EndIF
	Else
   		@ ++nLin,0 VTSay STR0019   //"Produto"
		@ ++nLin,0 VTGet cEti pict '@!' Valid VldEti(cEti,lInclusao) .and. ! Empty(aHisEti) 
	EndIf
	VTRead                               
	If lVolta
		Loop
	EndIf
	If VTLastKey() == 27 .and. Empty(aHisEti)
	   Exit
	EndIf
   If Len(aHisEti) > 1
      cTexto:= STR0006 //"Confirma a extracao das etiquetas ?"
   Else
      cTexto:= STR0007 //"Confirma a extracao da etiqueta ?"
   Endif
   If VTYesNo(cTexto,STR0008,.t.) //"Pergunta"
	   Manutencao(lInclusao)  
	Else
      If VTYesNo(STR0004,STR0005,.t.) //'Aborta a operacao ?'###'Pergunta'
	      Exit   
   	EndIf
	   Loop   
	EndIf
	cPallet:= If(UsaCB0("01"),Space(TamSX3("CB0_CODET2")[1]), IIf( FindFunction( 'AcdGTamETQ' ), AcdGTamETQ(), Space(48) ) )			
	aHisEti:= {}
Enddo

Vtsetkey(09,bkey09)
Vtsetkey(24,bkey24)
Return

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o	 ³ VldPallet  ³ Autor ³ Anderson Rodrigues  ³ Data ³ 26/02/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³ Verifica integridade da etiqueta do Pallet                 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso		 ³ SigaACD           	    								           ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function VldPallet(cPallet,lInclusao)
Local aItensPallet      
If Empty(cPallet) 
	Return .f.
Endif
aItensPallet := CBItPallet(cPallet)
If lInclusao
   If ! CBRetTipo(cPallet)=="" 
		VTBeep(2)
		VTAlert(STR0012,STR0010,.T.,4000)    //"AVISO" //"Etiqueta Invalida"
		VtClearGet("cPallet")
	   Return .f.
   Else
      aHisEti := AClone( aItensPallet )
	EndIf
Else  //exclusao
	If Len(aItensPallet) ==0
		VTBeep(2)
		VTAlert(STR0009,STR0010,.T.,4000)    //"Etiqueta de Pallet nao encontrada"###"AVISO"
		VtClearGet("cPallet")
	   Return .f.
	Endif
EndIf
Return .t.


If Len(aItensPallet) > 0 
   Return .t.
Endif
Return .f.

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o	 ³ VldEti     ³ Autor ³ Anderson Rodrigues  ³ Data ³ 24/02/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³ Verifica integridade da etiqueta                           ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso		 ³ SigaACD           	    								           ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function VldEti(cEti,lInclusao)
Local nPos
Local aEtiqueta
Local aItensPallet      

If Empty(cEti) 
   Return .f.
EndIf   

nPos:= Ascan(aHisEti,{|x| x == cEti})
If nPos > 0	
   VTBeep(2)
	VTAlert(STR0011,STR0010,.t.,4000) //"Etiqueta ja informada"###"Aviso"
	VtClearGet("cEti")
	Return .f.
Endif

aItensPallet := CBItPallet(cEti)
If Len(aItensPallet) > 0
   VTBeep(2)
   VTAlert(STR0012,STR0010,.T.,4000)    //"Etiqueta invalida"###"AVISO"
   VtClearGet("cEti")
   Return .f.
Endif

aEtiqueta    := CBRetEti(cEti,"01")
If Empty(aEtiqueta)  
   VTBeep(2)
   VTAlert(STR0012,STR0010,.t.,4000)  //"Etiqueta invalida"###"Aviso"
   VtClearGet("cEti")
   Return .f.
EndIf
If ! lInclusao .and. Empty(CB0->CB0_PALLET)
   VTBeep(2)
   VTAlert(STR0013,STR0010,.t.,4000)  //"Etiqueta nao pertence a nenhum Pallet "###"Aviso"
   VtClearGet("cEti")
   Return .f.        
ElseIf lInclusao .and. ! Empty(CB0->CB0_PALLET)
   VTBeep(2)
   VTAlert(STR0033+CB0->CB0_PALLET,STR0010,.t.,4000) //"Etiqeuta ja pertece ao pallet "###"Aviso"
   VtClearGet("cEti")
   Return .f.        
Endif

If lInclusao .And. !LocalizPallet( aEtiqueta[10], aEtiqueta[9], aHisEti )
   VTBeep(2)
   VTAlert( STR0035 , STR0009,.t.,4000)  //"Produto em Armazem e/ou Endereco Diferente ao do Pallet.""###"Aviso"
   VtClearGet("cEti")
   Return .f.
EndIf

If ! lInclusao .and. Alltrim(CB0->CB0_PALLET) # Alltrim(cPallet)
   VTBeep(2)
   VTAlert(STR0014,STR0010,.t.,4000)  //"Pallet da etiqueta e diferente do Pallet informado"###"Aviso"
   VtClearGet("cEti")
   Return .f.
Endif
aadd(aHisEti,cEti)
VtClearGet("cEti")
Return .f.


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o	 ³ Estorna    ³ Autor ³ Anderson Rodrigues  ³ Data ³ 24/02/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³ Realiza o estorno da(s) etiqueta(s) informada(s)           ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso		 ³ SIGAACD            	    		  			                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function Estorna()
Local aTela := VTSave()
Local cEti  := If(UsaCB0("01"),Space(TamSX3("CB0_CODET2")[1]), IIf( FindFunction( 'AcdGTamETQ' ), AcdGTamETQ(), Space(48) ) )			

VTClear()
@ 00,00 VtSay STR0034 //"Estorno da Leitura"
@ 01,00 VtSay STR0018 //"Etiqueta"
@ 02,00 VtGet cEti pict "@!" Valid VldEstorno(cEti)
VtRead
VtRestore(,,,,aTela)
Return

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o	 ³ ACD240Hist ³ Autor ³ Anderson Rodrigues  ³ Data ³ 24/02/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³ Mostra as etiqueta(s) lid(a)s 					              ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso		 ³ SIGAACD           	    								  		     ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function ACD240Hist()
Local aSave := VTSAVE()
Local aCab  := {STR0018,STR0019,STR0020} //"Etiqueta"###"Produto"###"Quantidade"
Local	aSize := {20,15,10}
Local aProds:= {}
Local nX
IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf

CB0->(DbSetOrder(1))
For nX:= 1 to Len(aHisEti)
   If CB0->(DbSeek(xFilial("CB0")+aHisEti[nX]))       
      aadd(aProds,{CB0->CB0_CODETI,CB0->CB0_CODPRO,Str(CB0->CB0_QTDE,8,2)})   
   Endif
Next
   
VtClear()
@ 0,0 VTSay STR0021	 //"Etiqueta(s) Lida(s):"
VTaBrowse(2,0,IIf(lVT100B /*GetMv("MV_RF4X20")*/,3,7),19,aCab,aProds,aSize)
If VtLastKey() == 27
   VtRestore(,,,,aSave)
Endif
VtRestore(,,,,aSave)
Return

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o	 ³ VldEstorno ³ Autor ³ Anderson Rodrigues  ³ Data ³ 25/02/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³ Valida o estorno da Leitura da(s) etiqueta(s)              ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso		 ³ SIGAACD            	    		  			                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function VldEstorno(cEti)
Local nPos

If Empty(cEti)
	Return .f.
EndIF
nPos:= Ascan(aHisEti,{|x| AllTrim(x) == AllTrim(cEti)})
If nPos == 0
   VTBeep(2)
   VTALERT(STR0022,STR0010,.T.,3000) //"Etiqueta nao encontrada"###"AVISO"
   VtKeyboard(Chr(20))
   Return .f.
Endif
If ! VTYesNo(STR0023,STR0024,.t.) //"Confirma o estorno da etiqueta ?"###"ATENCAO"
   VtKeyboard(Chr(20))
   Return .f.
EndIf
While .t.
   nPos:= Ascan(aHisEti,{|x| AllTrim(x) == AllTrim(cEti)})
   If nPos == 0
      Exit
   Endif
   aDel(aHisEti,nPos)
   aSize(aHisEti,Len(aHisEti)-1)
   VtKeyboard(Chr(20))
Enddo
Return .f. 

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡„o	 ³ Manutencao ³ Autor ³ Anderson Rodrigues  ³ Data ³ 25/02/03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡„o ³ Retira a etiqueta do Pallet                                ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso		 ³ SIGAACD            	    		  			                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function Manutencao(lInclusao)  
Local nX

For nX:= 1 to Len(aHisEti)
   CBRetEti(aHisEti[nX],"01")
   If !CB0->(EOF())
      Reclock("CB0",.f.)           
      If lInclusao
	      CB0->CB0_PALLET:= cPallet
      Else
      	CB0->CB0_PALLET:= ""
      EndIf
      CB0->(MsUnlock())
   Endif
Next
Return
