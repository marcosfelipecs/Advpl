#include "rwmake.ch"          

/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±º Funcao   ³ CBA100DEL  º Autor ³ Anderson Rodrigues º Data ³Wed  10/07/02º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDescri‡„o ³ Faz o estorno automatico do C.Q (Caso exista prod. no C.Q) -	º±±
±±º			 ³ Somente Protheus.											º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ SIGAACD                                                    	º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
/*/
Function CBA100DEL()
Local aSD1CQ
Local cSeek
Local lRastro
Local nCounter
Local nX
Private lMsErroAuto := .f.
Private lMsHelpAuto := .t.

If !SuperGetMV("MV_CBPE020",.F.,.F.)
	Return .T.
EndIf

If Type("l103AUTO") =="L" .and. l103AUTO 
	Return .T.
EndIf

aSD1CQ := VerifCQ(SF1->F1_DOC,SF1->F1_SERIE,SF1->F1_FORNECE,SF1->F1_LOJA)
If Empty(aSD1CQ)
   Return .t.
EndIf
nCounter := 0	
For nX := 1 to len(aSD1CQ)
   cSeek:=xFilial("SDB")+aSD1CQ[nX,1]+aSD1CQ[nX,2]+aSD1CQ[nX,3]+SF1->F1_DOC+SF1->F1_SERIE+SF1->F1_FORNECE+SF1->F1_LOJA		
   lRastro  := Rastro(aSD1CQ[nX,1])
   SDB->(DbSetOrder(1))		
   If SDB->(DbSeek(cSeek))
      Do While SDB->(!EOF()) .And. SDB->(DB_FILIAL+DB_PRODUTO+DB_LOCAL+DB_NUMSEQ+DB_DOC+DB_SERIE+DB_CLIFOR+DB_LOJA) == cSeek   
         	If SDB->DB_TM > "500" .Or. SDB->DB_TIPO # "D"
			   SDB->(DbSkip())
			   Loop
			EndIf
			If lRastro .And. !(aSD1CQ[nX,5]==SDB->DB_LOTECTL)
			   SDB->(DbSkip())
			   Loop
			EndIf
			If SDB->DB_ESTORNO # "S"
			   nCounter++
			EndIf
			SDB->(DbSkip())
      Enddo
   EndIf
Next
If nCounter == 0
   Return .t.
EndIf
If ! Empty(GetMV("MV_CBCQEND"))
   If !  MSGBOX("Confirma o estorno do enderecamento de C.Q.","Atencao","YESNO")
      Return .f.
   EndIf
   Processa({||DistriCQ(SF1->F1_DOC,SF1->F1_SERIE,SF1->F1_FORNECE,SF1->F1_LOJA,.t.)})
EndIf

Return .t.
