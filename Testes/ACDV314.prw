#INCLUDE "PROTHEUS.CH"
#INCLUDE "APVT100.CH"
#INCLUDE "ACDV314.CH"

#DEFINE NEWLINE CHR(13)+CHR(10)

//-------------------------------------------------------------------
/*/{Protheus.doc} ACDV314
Tela de Apontamento de Produção

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Function ACDV314()
   Local lContinua      := .T.
   Local lRet           := .T.
   Local lSaiu          := .F.
   Private cCeTrab      := ""
   Private cNRORPO      := Space(TamSX3("CZH_NRORPO")[1])
   Private cIDAT        := Space(TamSX3("CZH_IDAT")[1])
   Private cIDATQO      := Space(TamSX3("CZH_IDATQO")[1])
   Private cCDMQ        := Space(TamSX3("CZH_CDMQ")[1])
   Private cCDTR        := Space(TamSX3("CZH_CDTR")[1])
   Private cCDOE        := Space(TamSX3("CZH_CDOE")[1])
   Private nQTAPRP      := 0
   Private cCDRF        := Space(TamSX3("CZH_CDRF")[1])
   Private cTPTR        := Space(TamSX3("CZH_TPTR")[1])
   Private nTela        := 1
   Private lFinIni      := .F. //Finaliza inicialização de apontamento de split
   Private lAlocado     := .F.
   Private dDtIn, dDtFi := Date()
   Private cHrIn, cHrFi := Time()
   Default lAutoMacao   := .F.

   IF !lAutoMacao
      VtClearBuffer()
   
      If SuperGetMv('MV_INTACD',.F.,"0") != "1"
         VTAlert(STR0048, STR0010) //"Integração com o ACD desativada. Verifique o parâmetro MV_INTACD" ### 'Erro'
         Return
      EndIf

      While lContinua .And. IsTelnet()
         cCDOE := CBRETOPE() //Operador logado
         If nTela == 1
            If !A314MonT1()
               lSaiu := .T.
               Exit
            Else
               nTela++
            Endif
         Endif

         If nTela == 2
            If !A314MonT2()
               nTela--
            Else
               nTela++
            Endif
         Endif

         If nTela == 3
            If !A314MonT3()
               nTela--
            Else
               nTela++
            Endif
         Endif

         If nTela == 4
            If !A314MonT5()
               nTela-=2
            Else
               nTela++
            Endif
         Endif

         If nTela == 5
            If !A314MonT6()
               nTela-=2
            Else
               lContinua := .F.
            Endif
         Endif
      End

      If !lSaiu
         Begin Transaction

            //Grava informações A314Grava
            lRet := A314Grava()

            If lRet
            //Grava informações SFC A314Auto
               lRet := A314Auto(cNRORPO,cIDAT,cIDATQO,cCDMQ)
            Endif

            If Existblock("ACDV314VAL")
               lRet := Execblock("ACDV314VAL",.F.,.F.,{cNRORPO,cIDAT,cIDATQO,cCDMQ,cCDTR,nQTAPRP})
            EndIf

            If !lRet
               DisarmTransaction()
            Endif

         End Transaction
         If lRet
            VTAlert(STR0018, STR0019) //"Movimentacao efetuada com sucesso." / "Sucesso"
         EndIf
      Endif
   ENDIF
Return

//-------------------------------------------------------------------
/*/{Protheus.doc} A314MonT1
Monta tela 1 do apontamento

@param

@return  lRet   Retorna se todos os dados foram informados
                        corretamente e/ou se houve cancelamento

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function A314MonT1()
Local nLin  := 0
Local lRet  := .T.

VtClear()

@nLin,00 VtSay STR0001 //"Apontamento Produção"
nLin+=2
@nLin,00 VtSay STR0002 //"O.P."
nLin++
@nLin,00 VtGet cNRORPO  Pict "@!"  Valid A314VlOp(cNRORPO) F3 "CYQ002"
nLin++
@nLin,00 VtSay STR0003 //"Operação"
nLin++
@nLin,00 VtGet cIDAT    Pict "@!"  Valid A314VlOpe(cNRORPO,cIDAT) F3 "CY9002"
nLin++
@nLin,00 VtSay STR0004 //"Split"
nLin++
@nLin,00 VtGet cIDATQO  Pict "@!"  Valid A314VlSp(cNRORPO,cIDAT,cIDATQO) F3 "CYY003"
VtRead

If VTLastkey() == 27  // Tecla ESC
   lRet := .F.
EndIf

Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} A314MonT2
Monta tela 2 do apontamento

@param

@return  lRet   Retorna se todos os dados foram informados
                        corretamente e/ou se houve cancelamento

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function A314MonT2()
Local nLin  := 0
Local lRet  := .T.

VtClear()
@nLin,00 VtSay STR0001 //"Apontamento Produção"
nLin+=2
@nLin,00 VtSay STR0005 //"Máquina"
nLin++
If lAlocado
   @nLin,00 VtSay cCDMQ
Else
   @nLin,00 VtGet cCDMQ  Pict "@!" F3 "CYB003" Valid A314VlMq(cCDMQ)
Endif
nLin++
@nLin,00 VtSay STR0006 //"Transação"
nLin++
@nLin,00 VtGet cCDTR Pict "@!"  Valid A314VlTr(cCDTR) F3 "CBI"
VtRead

If VTLastkey() == 27  // Tecla ESC
   lRet := .F.
EndIf

Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} A314MonT3
Monta tela 3 do apontamento

@param

@return  lRet   Retorna se todos os dados foram informados
                        corretamente e/ou se houve cancelamento

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function A314MonT3()
Local nLin  := 0
Local lRet  := .T.

//Se transação for Iniciar Produção ou Finalizar Produção, não monta Tela3
If cTPTR $ "1"
   Return .T.
Endif

VtClear()
@00,00 VtSay STR0001 //"Apontamento Produção"
@02,00 VtSay STR0007 //"Quantidade"
@03,00 VtGet nQTAPRP Pict "@E 9999999999,99" Valid A314VlQt()

//Se transação for Refugo ou Retrabalho solicita Motivo
If cTPTR == "5"
   @04,00 VtSay STR0008 //"Motivo"
   nLin++
   @05,00 VtGet cCDRF Pict "@!" Valid A314VlMt() F3 "CYO002"
Endif

VtRead()

If VTLastkey() == 27    // Tecla ESC
   lRet := .F.
EndIf

Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} A314VlQt
Valida campo quantidade reportada

@param

@return  lRet   Retorna se todos os dados foram informados
                        corretamente e/ou se houve cancelamento

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function A314VlQt()
Local lRet   := .T.
Local nQtdAp := 0
Local nQtdRf := 0
Local nQtdRp := 0
Local cTpTran

//Cálculo de quantidade total reportada até o momento
DbSelectArea("CZH")
CZH->(DbSetOrder(4))
CZH->(dbSeek(xFilial("CZH")+cNRORPO+cIDAT+cIDATQO+"1"))

While CZH->(!Eof()) .And.;
          CZH->CZH_FILIAL  == xFilial("CZH")   .And. ;
          CZH->CZH_NRORPO  == cNRORPO          .And. ;
          CZH->CZH_IDAT    == cIDAT            .And. ;
          CZH->CZH_IDATQO  == cIDATQO          .And. ;
          CZH->CZH_STTR    == "1"

   cTpTran := A314GtTr(CZH->CZH_CDTR,2)
   If cTpTran $ "24" //Aprovada
      nQtdAp += CZH->CZH_QTAPRP
   Endif

   If cTpTran $ "5" //Refugo
      nQtdRf += CZH->CZH_QTAPRP
   Endif

   CZH->(dbSkip())
End

nQtdRp := nQtdAp + nQtdRf + nQTAPRP

If !A314VlQtR(nQtdRp)
   Return .F.
Endif

Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} A314VlMt
Valida campo motivo de refugo/retrabalho

@param

@return    lRet     Retorna se todos os dados foram informados
                        corretamente e/ou se houve cancelamento

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function A314VlMt()
   Local lRet := IIf(ExistCpo("CYO",cCDRF), .T.,.F.)
   Default lAutoMacao := .F.

   If !lRet
      IF !lAutoMacao
         VTAlert(STR0009, STR0010) //"Motivo não cadastrado." / STR0010
      ENDIF
   Endif
Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} A314VlTr
Valida a transação informada

@param

@return  lRet   Retorna se é possível realizar a transação
                        informada

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function A314VlTr()
dDtIn := Date()
dDtFi := Date()
cHrIn := Time()
cHrFi := Time()

cTPTR := A314GtTr(cCDTR,1)
If Empty(cTPTR)
   CBAlert(STR0011,STR0010,.T.,3000,2,.t.) //"Transação da produção não cadastrada ou inválida."
   Return .f.
Endif

If cTPTR $ "1" .And. A314VlIn()
   Return .T.
Endif
If cTPTR $ "245" .And. A314VlAp()
   Return .T.
Endif
Return .F.

//-------------------------------------------------------------------
/*/{Protheus.doc} A314GtTr
Retorna tipo de transação equivalente ao código de transação passado

@param     cCdTrans Código da transação na CBI

@return  cTpTrans  Retorna tipo de transação. Os tipos sao: 1-Inicio,
                   2-Pausa c/,3-Pausa s/,4-Producao,5-Perda

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function A314GtTr(cCdTrans,nTipo)
Local cTpTrans := Space(1)

dbSelectArea("CBI")
CBI->(DbSetOrder(1))
If ! CBI->(DbSeek(xFilial("CBI")+cCdTrans))
   Return cTpTrans
EndIf

If CBI->CBI_TIPO <> "3" //Pausa sem apontamento
   cTpTrans := CBI->CBI_TIPO
Endif
If nTipo == 1
   lFinIni := IIf(CBI->CBI_FIMINI == "1",.T.,.F.)
Endif
Return cTpTrans

//-------------------------------------------------------------------
/*/{Protheus.doc} A314VlIn
Valida a possibilidade de inicialização do apontamento do split

@param

@return    lRet   Retorna se é possível inicializar o
                        apontamento do split

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function A314VlIn()
Local cTpTran
Local cMsg := ''
Default lAutoMacao := .F.
//Verifica se o apontamento do split já foi iniciado
DbSelectArea("CZH")
DbSetOrder(2)
If CZH->(dbSeek(xFilial("CZH")+cNRORPO+cIDAT+cIDATQO+cCDMQ+"1"))
   While CZH->(!Eof()) .And. xFilial("CZH")+cNRORPO+cIDAT+cIDATQO+cCDMQ+"1" == CZH->(CZH_FILIAL+CZH_NRORPO+CZH_IDAT+CZH_IDATQO+CZH_CDMQ+CZH_STTR)
      If Empty(CZH->CZH_DTRPED) .And. Empty(CZH->CZH_HRRPED)
         cTpTran := A314GtTr(CZH->CZH_CDTR,2)
         If cTpTran == cTPTR
            VTAlert(STR0014, STR0015) //"Apontamento ja iniciado" / "Atencao"
            Return .F.
         Endif
      Endif
      CZH->(DbSkip())
   End
EndIf
CZH->(DbCloseArea())

//Verifica se a máquina informada já possui um apontamento iniciado
DbSelectArea("CZH")
DbSetOrder(3)
If CZH->(dbSeek(xFilial("CZH")+cCDMQ))
   While CZH->(!Eof()) .And. xFilial("CZH")+cCDMQ == CZH->(CZH_FILIAL+CZH_CDMQ)
      If Empty(CZH->CZH_DTRPED) .And. Empty(CZH->CZH_HRRPED) .And. CZH->CZH_STTR != "2"
         cTpTran := A314GtTr(CZH->CZH_CDTR,2)
         If cTpTran == cTPTR
            cMsg := STR0016 + ; //"Maquina ja possui apontamento iniciado"
                    CHR(13) + CHR(10) + ;
                    STR0054 + AllTrim(CZH->CZH_NRORPO) + ; // "OP: "
                    CHR(13) + CHR(10) + ;
                    STR0003 + ": " + AllTrim(CZH->CZH_IDAT) // "Operação: "
            IF !lAutoMacao
               VTAlert(cMsg, STR0015)
            ENDIF
            Return .F.
         Endif
      Endif
      CZH->(DbSkip())
   End
EndIf
CZH->(DbCloseArea())

Return .T.

//-------------------------------------------------------------------
/*/{Protheus.doc} vldHrInic
Valida se o split pode ser iniciado no horário informado

@param

@return  lRet   Retorna se é possível fazer o
                        apontamento do split

@author  Lucas Konrad França
@since   01/02/2016
@version P11.8
/*/
//-------------------------------------------------------------------
Static Function vldHrInic()
	Local lRet       := .T.

   dbSelectArea("CYB")
   CYB->(dbSetOrder(1))
   If CYB->(dbSeek(cCDMQ))
      dbSelectArea("CYI")
      CYI->(dbSetOrder(1))
      CYI->(dbSeek(xFilial("CYI")+CYB->CYB_CDCETR))
   EndIf

   dbSelectArea("CYQ")
   CYQ->(dbSetOrder(1))
   CYQ->(dbSeek(xFilial("CYQ")+cNRORPO))

   dbSelectArea("CYY")
   CYY->(dbSetOrder(1))
   CYY->(dbSeek(xFilial("CYY")+cNRORPO+cIDAT+cIDATQO))

   If lRet .And. CYI->CYI_TPPC == '1' .AND. !CYB->CYB_LGOVRP .And. (CYQ->CYQ_TPRPOR == "2" .OR. (CYQ->CYQ_TPRPOR == "1" .AND. CYY->CYY_LGCERP))
      if Len(SFCA314VAP(cCDMQ,dDtIn,StoD("31129999"),cHrIn,"23:59:59")) > 0
         VTAlert(STR0055,STR0015) //"Já existe um apontamento de produção nesta data/hora. Início de produção não permitido."
         lRet := .F.
      Endif
   Endif

Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} A314VlAp
Valida a possibilidade de apontamento do split

@param

@return  lRet   Retorna se é possível fazer o
                        apontamento do split

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function A314VlAp()
Local lRet := .F.
Local cTpTran := Space(1)
Default lAutoMacao := .F.

//Verifica se o apontamento do split já foi iniciado
DbSelectArea("CZH")
DbSetOrder(2)
If CZH->(dbSeek(xFilial("CZH")+cNRORPO+cIDAT+cIDATQO+cCDMQ))
   While CZH->(!Eof()) .And. xFilial("CZH")+cNRORPO+cIDAT+cIDATQO+cCDMQ == CZH->(CZH_FILIAL+CZH_NRORPO+CZH_IDAT+CZH_IDATQO+CZH_CDMQ)
      cTpTran := A314GtTr(CZH->CZH_CDTR,2)
      If cTpTran == "1" .And. Empty(CZH->CZH_DTRPED) .And. Empty(CZH->CZH_HRRPED)
         lRet    := .T.
         dDtIn   := CZH->CZH_DTRPBG
         cHrIn   := CZH->CZH_HRRPBG
         Exit
      Endif
      CZH->(DbSkip())
   End
EndIf

If !lRet
   IF !lAutoMacao
      VTAlert(STR0017, STR0015)
   ENDIF
   Return lRet
Endif
Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} A314Grava
Grava informações na tabela CZH conforme transação selecionada

@param

@return

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function A314Grava()
Local oModelCYY
Local aData   := A314ADtIn(cNRORPO,cIDAT,cIDATQO,cCDMQ)
Local lRet    := .T.
Local cTpTran := Space(1)
Default lAutoMacao := .F.

If cTPTR == "1" //Inicializa split
   If !VTYesNo(STR0012,STR0013,.T.) //"Deseja iniciar o split?" / "Aviso"
      Return .F.
   Endif
Endif

If cTPTR == "2" .Or. lFinIni   //Finaliza split inicial
   dbSelectArea("CZH")
   dbSetOrder(4)
   If CZH->(dbSeek(xFilial("CZH")+cNRORPO+cIDAT+cIDATQO+"1"))
      While CZH->(!Eof()) .And. xFilial("CZH")+cNRORPO+cIDAT+cIDATQO+"1" == CZH->(CZH_FILIAL+CZH_NRORPO+CZH_IDAT+CZH_IDATQO+CZH_STTR)
         cTpTran := A314GtTr(CZH->CZH_CDTR,2)
         If cTpTran == "1" .And. Empty(CZH->CZH_DTRPED) .And. Empty(CZH->CZH_HRRPED)
            Exit
         Endif
         CZH->(DbSkip())
      End
   EndIf

   Reclock("CZH",.F.)
   CZH->CZH_DTRPED := dDtFi
   CZH->CZH_HRRPED := cHrFi
   CZH->(MsUnLock())
Endif

IF !lAutoMacao
   RecLock("CZH",.T.) // Cria Split de Movimentação
   CZH->CZH_FILIAL := xFilial("CZH")
   CZH->CZH_NRORPO := cNRORPO
   CZH->CZH_CDOE   := cCDOE
   CZH->CZH_IDAT   := cIDAT
   CZH->CZH_IDATQO := cIDATQO
   CZH->CZH_CDMQ   := cCDMQ
   CZH->CZH_CDTR   := cCDTR
   CZH->CZH_TPTR   := cTPTR
   CZH->CZH_STTR   := "1" //Pendente de envio ao Chão de Fábrica

   //CZH->CZH_CDLOSR
   //CZH->CZH_DTVLLO
   If cTPTR == "1" //Iniciar Produção
      CZH->CZH_DTRPBG := dDtIn
      CZH->CZH_HRRPBG := cHrIn
   Endif

   //Produção
   If cTPTR $ "245"
      CZH->CZH_DTRPBG    := aData[1]
      CZH->CZH_HRRPBG    := aData[2]
      If !GetMV('MV_SFCAPON') .And. cTPTR != "2" .And. !lFinIni
         CZH->CZH_DTRPED := CZH->CZH_DTRPBG
         CZH->CZH_HRRPED := SumTime(CZH->CZH_HRRPBG)
      Else
         CZH->CZH_DTRPED := dDtFi
         CZH->CZH_HRRPED := cHrFi
      Endif
      CZH->CZH_QTAPRP := nQTAPRP
   Endif

   If cTPTR $ "5"
      CZH->CZH_CDRF := cCDRF
   Endif

   CZH->(MsUnLock())
   CZH->(dbCloseArea())
ENDIF

If cTPTR == "1" .And. lRet //Se transação inicial, alocar máquina para o split
   dbSelectArea("CYY")
   CYY->(dbSetOrder(1))
   CYY->(dbGoTop())
   If CYY->(dbSeek(xFilial("CYY")+cNRORPO+cIDAT+cIDATQO))
      If Empty(CYY_CDMQ)
         oModelCYY := FWLoadModel( "SFCA315" )
         oModelCYY:SetOperation( 4 )
         oModelCYY:Activate()

         oModelCYY:SetValue( "CYYMASTER", "CYY_CDMQ", cCDMQ)
         oModelCYY:VldData()
         oModelCYY:CommitData()
         oModelCYY:DeActivate()
      Endif
   Endif
Endif

Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} A314VlOp
Retorna se ordem de produção é válida para apontamento

@param   cOrdem  Identifica a ordem de produção informada

@return  lRet   Retorna se valor informado é válido

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function A314VlOp(cOrdem)
   Local lRet := .T.
   Default lAutoMacao := .F.

   If Empty(cOrdem)
      IF !lAutoMacao
         VTAlert(STR0020, STR0010) //"Ordem de produção deve ser preenchida."
      ENDIF
      Return .F.
   EndIf

   DbSelectArea("CYQ")

   If !(CYQ->(dbSeek(xFilial("CYQ")+cOrdem)))
      VTAlert(STR0021, STR0010) // "Ordem de produção não existente."
      lRet := .F.
   Else
      If CYQ->CYQ_TPST $ "45"
         VTAlert(STR0022, STR0010) //"Ordem de produção com situação inválida."
         lRet := .F.
      Endif
   Endif
Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} A314VlMq
Retorna se máquina é válida para apontamento

@param   cMaquina Identifica a máquina informada

@return  lRet   Retorna se valor informado é válido

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function A314VlMq(cMaquina)
   Local lRet := .T.
   Default lAutoMacao := .F.

   If Empty(cMaquina)
      IF !lAutoMacao
         VTAlert(STR0023, STR0010) //"Maquina invalida"
      ENDIF
      Return .F.
   EndIf

   DbSelectArea("CYB")
   If !(CYB->(dbSeek(xFilial("CYB")+cMaquina)))
      VTAlert(STR0024, STR0010) //"Maquina nao existente"
      lRet := .F.
   Endif
Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} A314VlOpe
Retorna se a operação é válida para apontamento

@param      cOp        Identifica a ordem informada
            cOper      Identifica a operação informada

@return  lRet   Retorna se valor informado é válido

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function A314VlOpe(cOp, cOper)
   Local lRet := .T.
   Default lAutoMacao := .F.

   If Empty(cOper)
      IF !lAutoMacao
         VTAlert(STR0025, STR0010) //"Operacao invalida"
      ENDIF
      Return .F.
   EndIf

   DbSelectArea("CY9")
   CY9->(dbSetOrder(1))
   If !(CY9->(dbSeek(xFilial("CY9")+cOp+cOper)))
      VTAlert(STR0026, STR0010) //"Operação não existente para a ordem informada."
      lRet := .F.
   Endif
   //Atribui o centro de trabalho a variavel private para ser filtrada na seleção de máquina
   cCeTrab := CY9->(CY9_CDCETR)

Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} A314VlSp
Retorna se split é válido para apontamento

@param   cOp    Identifica a ordem de produção informada
         cOper  Identifica a operação informada
         cSplit Identifica o split informado

@return  lRet   Retorna se valor informado é válido

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Function A314VlSp(cOp,cOper,cSplit)
Local lRet := .T.
Local nQt  := (1/100)
Default lAutoMacao := .F.

dbSelectArea("CYY")
CYY->(dbSetOrder(1))
CYY->(dbGoTop())
If CYY->(dbSeek(xFilial("CYY")+cOp+cOper+cSplit))
   //Verifica se o status está finalizado(5) ou suspenso(6)
   If CYY->CYY_TPSTAT $ "56"
      VTAlert(STR0047, STR0010) //"Split já finalizado ou suspenso."
      lRet := .F.
   Else
      //Verifica se possui uma máquina alocada ao slit e grava o código da mesma
      cCDMQ := CYY->CYY_CDMQ
      If Empty(cCDMQ)
         lAlocado := .F.
      Else
         lAlocado := .T.
      Endif
   Endif
Else
   IF !lAutoMacao 
      VTAlert(STR0027, STR0010) //"Split inexistente"
   ENDIF
   lRet := .F.
Endif

If lRet
   //Valida operador - transação já iniciada por outro operador
   DbSelectArea("CZH")
   DbSetOrder(2)
   CZH->(dbSeek(xFilial("CZH")+cOp+cOper+cSplit))
   While CZH->(!Eof()) .And. CZH->CZH_NRORPO == cOp .And.;
      CZH->CZH_IDAT == cOper .And. CZH->CZH_IDATQO == cSplit
      If Empty(CZH->CZH_DTRPED) .And. Empty(CZH->CZH_HRRPED)
         //Inserido AllTrim para remover os espaços desnecessários
         If AllTrim(CZH->CZH_CDOE) != AllTrim(cCDOE)
            lRet := .F.
            VTAlert(STR0028, STR0010) //"Split já iniciado por outro operador."
            Exit
         Endif
      Endif
      CZH->(DbSkip())
   End
   CZH->(DbCloseArea())
Endif

If lRet
   //Valida se é possível apontar algo no split informado
   If !A314VlQtR(nQt)
      lRet := .F.
   Endif
Endif

Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} A314ADtIn
Retorna data de início para as transações de "Produção", "Refugo"
e "Pausar com apontamento"

@param   cOp    Identifica a ordem informada
         cOper  Identifica a operação informada
         cSplit Identifica o split informado
         cMaq   Identifica a máquina informada

@return  aData   Retorna data/hora inicial

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function A314ADtIn(cOp, cOper, cSplit, cMaq)
Local aData   := {"",""}
Local cTpTran := Space(1)

DbSelectArea("CZH")
DbSetOrder(4)

If CZH->(dbSeek(xFilial("CZH")+cOp+cOper+cSplit+"1"))
   While CZH->(!Eof()) .And. ;
         CZH->CZH_FILIAL == xFilial("CZH") .And. ;
         CZH->CZH_NRORPO == cOp .And. ;
         CZH->CZH_IDAT == cOper .And. ;
         CZH->CZH_IDATQO == cSplit .And. ;
         CZH->CZH_STTR == "1"

      cTpTran := A314GtTr(CZH->CZH_CDTR,2)
      If cTpTran == "1" //Se transação for "Iniciar Produção"
         If Empty(aData[1]) .Or.;
            (aData[1] < CZH->CZH_DTRPBG .Or. ;
            (aData[1] == CZH->CZH_DTRPBG .And. aData[2] < CZH->CZH_HRRPBG) )

            aData[1] := CZH->CZH_DTRPBG
            aData[2] := CZH->CZH_HRRPBG
         Endif
      Else
         If Empty(aData[1]) .Or.;
            (aData[1] < CZH->CZH_DTRPED .Or. ;
            (aData[1] == CZH->CZH_DTRPED .And. aData[2] < CZH->CZH_HRRPED) )

            aData[1] := CZH->CZH_DTRPED
            aData[2] := CZH->CZH_HRRPED
         Endif
      Endif

      CZH->(DbSkip())
   End
EndIf
Return aData

//-------------------------------------------------------------------
/*/{Protheus.doc} A314Auto
Envia movimentações para o módulo SFC

@param   cOp    Identifica a ordem informada
         cOper  Identifica a operação informada
         cSplit Identifica o split informado
         cMaq   Identifica a máquina informada

@return  lRet   Retorna se houve algum erro ou se a transação
                        foi efetuada

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function A314Auto(cOp, cOper, cSplit, cMaq)
Local nI
Local oModel
Local oModelCY0
Local oModelCYW
Local oModelCZ0
Local nQtdAp := 0
Local nQtdRf := 0
Local nQtdRt := 0
Local nQtdRp := 0
Local dDtFim
Local cHrFim
Local aRefugo    := {}
Local cTpTran    := Space(1)
Local aOperador  := {}
Local aFerrament := {}

Private lAutoErrNoFile := .F.
Private lMSErroAuto := .F.

If cTPTR != "2" .And. !lFinIni //Envia informações para o Chão de Fábrica
   Return .T.
Endif

If !IsInCallStack("SFCA310")
   Private _IsSFCA318 := .F.
Endif

aOperador  := A314CarOE()
aFerrament := A314CarFR()

//Posicionar no split para ativar o model
DbSelectArea("CYY")
CYY->(DbSetOrder(1))
If CYY->(!DbSeek(xFilial("CYY")+cOp+cOper+cSplit))
   Return .F.
Endif

// Instancia o modelo
oModel := FWLoadModel( "SFCA314" )
oModel:SetOperation( 3 ) //incluir apontamento
If !oModel:Activate()
   VTAlert(AllTrim(STR0010)+": "+oModel:GetErrorMessage()[6], oModel:GetErrorMessage()[5])
   Return .F.
EndIf

oModelCY0 := oModel:GetModel( "CY0DETAIL" )

DbSelectArea("CZH")
CZH->(DbSetOrder(4))

If CZH->(dbSeek(xFilial("CZH")+cOp+cOper+cSplit+"1"))
   While CZH->(!Eof()) .And.;
             CZH->CZH_FILIAL    == xFilial("CZH")   .And. ;
             CZH->CZH_NRORPO    == cOp              .And. ;
             CZH->CZH_IDAT      == cOper            .And. ;
             CZH->CZH_IDATQO    == cSplit           .And. ;
             CZH->CZH_STTR      == "1"

      cTpTran := A314GtTr(CZH->CZH_CDTR,2)
      If cTpTran == "1" //Inicialização
         oModel:SetValue("CYVMASTER","CYV_NRORPO",CZH->CZH_NRORPO)
         oModel:SetValue("CYVMASTER","CYV_IDAT"  ,CZH->CZH_IDAT)
         oModel:SetValue("CYVMASTER","CYV_IDATQO",CZH->CZH_IDATQO)
         oModel:SetValue("CYVMASTER","CYV_CDMQ"  ,CZH->CZH_CDMQ)
         oModel:SetValue("CYVMASTER","CYV_DTRPBG",CZH->CZH_DTRPBG)
         oModel:SetValue("CYVMASTER","CYV_HRRPBG",CZH->CZH_HRRPBG)
         dDtFim := CZH->CZH_DTRPED
         cHrFim := CZH->CZH_HRRPED
         oModel:SetValue("CYVMASTER","CYV_HRRPED",CZH->CZH_HRRPED)
         oModel:SetValue("CYVMASTER","CYV_DTRPED",CZH->CZH_DTRPED)
      Endif

      If cTpTran $ "24" //Aprovada
         nQtdAp += CZH->CZH_QTAPRP
      Endif

      If cTpTran == "5" //Refugo
         nQtdRf += CZH->CZH_QTAPRP

         A314AddRe(aRefugo,1,CZH->CZH_CDRF, CZH->CZH_QTAPRP)
      Endif

      CZH->(DbSkip())
   End
EndIf

nQtdRp = nQtdAp + nQtdRf
If nQtdRp <= 0
   Return .F.
Endif

If oModelCY0:Length() == 1
   oModelCY0:GoLine( 1 )
   oModelCY0:DeleteLine()
Endif

//Adiciona quantidades refugadas e retrabalhadas, e seus respectivos motivos
For nI := 1 To Len(aRefugo)
   oModelCY0:AddLine()
   oModelCY0:SetValue("CY0_CDRF",aRefugo[nI][1])
   If aRefugo[nI][2] > 0
      oModelCY0:SetValue("CY0_QTRF",aRefugo[nI][2])
   Endif
Next

oModel:SetValue("CYVMASTER","CYV_QTATRP",nQtdRp)
oModel:SetValue("CYVMASTER","CYV_QTATRF",nQtdRf)
oModel:SetValue("CYVMASTER","CYV_QTATRT",nQtdRt)
oModel:SetValue("CYVMASTER","CYV_QTATAP",nQtdAp)
oModel:SetValue("CYVMASTER","CYV_HRRPED",cHrFim)
oModel:SetValue("CYVMASTER","CYV_DTRPED",dDtFim)

oModelCYW := oModel:GetModel( "CYWDETAIL" )

For nI := 1 To Len(aOperador)
   If nI > 1
      oModelCYW:AddLine()
   Else
      oModelCYW:GoLine(nI)
   Endif
   If aOperador[nI][1] == '1'
      oModelCYW:SetValue("CYW_CDOE",aOperador[nI][2])
   Else
      oModelCYW:SetValue("CYW_CDGROE",aOperador[nI][2])
   Endif
   oModelCYW:SetValue("CYW_DTBGRP",dDtIn)
   oModelCYW:SetValue("CYW_HRBGRP",cHrIn)
   oModelCYW:SetValue("CYW_DTEDRP",dDtFi)
   oModelCYW:SetValue("CYW_HREDRP",cHrFi)
Next nI


oModelCZ0 := oModel:GetModel( "CZ0DETAIL" )

For nI := 1 To Len(aFerrament)
   If nI > 1
      oModelCZ0:AddLine()
   Else
      oModelCZ0:GoLine(nI)
   Endif
   oModelCZ0:SetValue("CZ0_CDFE",aFerrament[nI])
Next nI
// Valida o modelo
If oModel:VldData()
   VtClear()
   If !oModel:CommitData()
     aErro := oModel:GetErrorMessage()
     If !Empty(aErro[6])
        VTAlert(AllTrim(STR0010)+": "+oModel:GetErrorMessage()[6], oModel:GetErrorMessage()[5])
     Else
        VTAlert(STR0056,AllTrim(STR0010)) //"Ocorreram erros ao realizar o apontamento."
     EndIf
     Return .F.
   EndIf
Else
   VtClear()
   aErro := oModel:GetErrorMessage()
   VTAlert(AllTrim(STR0010)+": "+oModel:GetErrorMessage()[6], oModel:GetErrorMessage()[5])
   Return .F.
EndIf

//Atualiza informações de apontamento da tabela temporária CZH,
//caso seja uma finalização de apontamento
dbSelectArea("CZH")
CZH->(dbSetOrder(2))

If CZH->(dbSeek(xFilial("CZH")+cNRORPO+cIDAT+cIDATQO))
   While CZH->(!Eof()) .And. ;
         CZH->CZH_FILIAL == xFilial("CZH") .And. ;
         CZH->CZH_NRORPO == cNRORPO .And. ;
         CZH->CZH_IDAT == cIDAT .And. ;
         CZH->CZH_IDATQO == cIDATQO

      If CZH->CZH_STTR == "1"
         RecLock("CZH",.F.)
         CZH->CZH_STTR := "2" //Enviado ao Chão de Fábrica
         CZH->(MsUnLock())
      Endif
      CZH->(DbSkip())
   End
EndIf
Return .T.

//-------------------------------------------------------------------
/*/{Protheus.doc} A314AddRe
Adiciona Refugo/Retrabalho no array enviado, comparando os motivos
já utilizados para agrupá-los se necessário.

@param     aDadRe Identifica o array onde o item será adicionado
           cCdRf  Identifica o código do refugo/retrabalho
           nQtRf  Identifica a quantidade do refugo/retrabalho

@return  aDadRe   Retorna array com a informação adicionada

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function A314AddRe(aDadRe, nTipo, cCdRf, nQtRf)
Local nI     := 0
Local lAchou := .F.

For nI := 1 To Len(aDadRe)
   If aDadRe[nI][1] == cCdRf
      If nTipo == 1
         aDadRe[nI][2] += nQtRf
         aDadRe[nI][3] += 0
      Else
         aDadRe[nI][2] += 0
         aDadRe[nI][3] += nQtRf
      Endif
      lAchou := .T.
      Exit
   Endif
Next

If !lAchou
   If nTipo == 1
      aAdd(aDadRe,{cCdRf,nQtRf,0})
   Else
      aAdd(aDadRe,{cCdRf,0,nQtRf})
   Endif
Endif
Return aDadRe

//-------------------------------------------------------------------
/*/{Protheus.doc} A314VlQtR
Validação da quantidade informada

@param   nQtTotal   Quantidade total reportada

@return  lRet       Retorna se motivo será enviado para consulta padrão

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function A314VlQtR(nQtTotal)
Local lRet := .T.

dbSelectArea("CYQ")
CYQ->(dbSetOrder(1))
CYQ->(dbGoTop())
CYQ->(dbSeek(xFilial("CYQ")+cNRORPO))

If (CYQ->CYQ_TPRPOR == "2" .OR. CYQ->CYQ_TPRPOR == "3") // Reporte por Operação ou Ponto de Controle
   // Verificar se a operação anterior é externa, pois se for, não deverá ser verificada
   dbSelectArea("CY9")
   CY9->(dbSetOrder(1))
   CY9->(dbGoTop())
   CY9->(dbSeek(xFilial("CY9")+cNRORPO+CYD->CYD_IDATPV))

   If CY9->CY9_TPAT == "1"
      nQtdMax := SFCA314QRP(cNRORPO,cIDAT,cIDATQO,0,.T.,1,Date(),Time(),.F.,1)
   Else
      nQtdMax := nQtTotal
   Endif
   Pergunte('MTA680',.F.)
   If MV_PAR07 != 3
      If nQtdMax == 0 .AND. nQtTotal > 0
         VtAlert(STR0029 + NEWLINE +; //"Operação anterior ainda não foi reportada ou quantidade desejada não disponível para reporte na data informada. "
                 STR0002+": " + cNRORPO + NEWLINE +;
                 STR0003+": " + cIDAT + NEWLINE +;
                 STR0004+": " + cIDATQO, "SFCA314") //"Ordem: "###" Operação: "###" Split: "
         lRet := .F.
      Else
         If nQtdMax < nQtTotal
            VtAlert(STR0030 + NEWLINE +; //"Quantidade maior que a quantidade da operação anterior"
                    STR0002+": " + cNRORPO + NEWLINE +;
                    STR0003+": " + cIDAT + NEWLINE +;
                    STR0004+": " + cIDATQO + NEWLINE +;
                    STR0007+": " + STR(nQtdMax), "SFCA314") //"Ordem: "###" Operação: "###" Split: "
            lRet := .F.
         Endif
      EndIf
   Endif
Endif
Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} A314CarFR
Função para tratamento de dados necessários quando o tipo de tratamento
de tempo da operação é dependente de ferramenta.

@param

@return  aDados      Retorna array com as informações de ferramenta

@author  Lucas Konrad França
@since   21/01/2016
@version P118
/*/
//-------------------------------------------------------------------
Static Function A314CarFR()
Local aDados   := {}
Local nSFCFEA  := SuperGetMv("MV_SFCFEA",.F.,1)

Private cCdFer := Space(TamSX3("CYC_CDRC")[1])

dbSelectArea("CY9")
CY9->(dbSetOrder(1))
If CY9->(dbSeek(xFilial("CY9")+cNRORPO+cIDAT))
   If CY9->CY9_TPTE == "4" // Operação dependente de ferramenta.
      dbSelectArea("CYC")
      CYC->(dbSetOrder(1))
      If nSFCFEA == 1 //Primeira ferramenta cadastrada para a máquina
         If CYC->(dbSeek(xFilial("CYC")+cCDMQ+"2"))
            aAdd(aDados,CYC->CYC_CDRC)
         EndIf
      ElseIf nSFCFEA == 2 //Informa a ferramenta
         CYC->(dbSeek(xFilial("CYC")+cCDMQ+"2"))
         If !A314MonT7()
            Return aDados
         Else
            aAdd(aDados,cCdFer)
         EndIf
      ElseIf nSFCFEA == 3 //Todas as ferramentas cadastradas para a máquina
         If CYC->(dbSeek(xFilial("CYC")+cCDMQ+"2"))
            While CYC->( !Eof() ) .And. CYC->(CYC_FILIAL+CYC_CDMQ+CYC_TPRC) == xFilial("CYC")+cCDMQ+"2"
               aAdd(aDados,CYC->CYC_CDRC)
               CYC->(dbSkip())
            End
         EndIf
      EndIf
   EndIf
EndIf

Return aDados

//-------------------------------------------------------------------
/*/{Protheus.doc} ACDV314FER
Função de filtro para a consulta padrão CYC009

@param

@return  lRet - Exibe ou não o registro na consulta padrão.

@author  Lucas Konrad França
@since   21/01/2016
@version P118
/*/
//-------------------------------------------------------------------
Function ACDV314FER()
   Local lRet := .F.

   If CYC->CYC_CDMQ == cCDMQ .AND. CYC->CYC_TPRC == '2'
      lRet := .T.
   EndIf
Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} A314CarOE
Função para tratamento de dados necessários quando o tipo de reporte
do centro de trabalho em que o apontamento está acontecendo é por
operador ou equipe.

@param

@return  aDados   Retorna array com as informações de equipe/operador

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function A314CarOE()
Local aDados    := {}
Local cTpMod    := Space(1)
Local nSFCOEA   := SuperGetMv("MV_SFCOEA",.F.,1)
Private cCDOERP := Space(TamSX3("CYC_CDRC")[1])


dbSelectArea("CYB")
CYB->(dbSetOrder(1))
If CYB->(dbSeek(xFilial("CYB")+cCDMQ))
   DbSelectArea("CYI")
   CYI->(DbSetOrder(1))
   If CYI->(dbSeek(xFilial("CYI")+CYB->CYB_CDCETR))
      If CYI->CYI_TPMOD == "1" //Não reporta
         Return aDados
      Endif

      dbSelectArea("CYC")
      CYC->(dbSetOrder(1))
      CYC->(dbSeek(xFilial("CYC")+cCDMQ))

      If CYI->CYI_TPMOD == "2" //Reporta por operador
         cTpMod := "1"
      Endif

      If CYI->CYI_TPMOD == "3" //Reporta por equipe
         cTpMod := "3"
      Endif

   EndIf
EndIf

If nSFCOEA == 1 //Primeira equipe/operador cadastrada para a máquina
   If CYC->(dbSeek(xFilial("CYC")+cCDMQ+cTpMod))
      aAdd(aDados,{cTpMod,CYC->CYC_CDRC})
   EndIf
ElseIf nSFCOEA == 2 //Informa a equipe/operador
   //ABRE TELA PARA INFORMACAO DO OPERADOR OU EQUIPE
   If !A314MonT4(CYI->CYI_TPMOD)
      Return aDados
   Else
      If CYI->CYI_TPMOD == "2"
         aAdd(aDados,{'1',cCDOERP}) //REVER
      Else
         aAdd(aDados,{'3',cCDOERP}) //REVER
      Endif
      Return aDados
   Endif
ElseIf nSFCOEA == 3 //Todas as esquipes/operadores cadastrados para a máquina.
   If CYC->(dbSeek(xFilial("CYC")+cCDMQ+cTpMod))
      While CYC->( !Eof() ) .And. CYC->(CYC_FILIAL+CYC_CDMQ+CYC_TPRC) == xFilial("CYC")+cCDMQ+cTpMod
         aAdd(aDados,{cTpMod,CYC->CYC_CDRC})
         CYC->(dbSkip())
      End
   EndIf
EndIf
Return aDados

//-------------------------------------------------------------------
/*/{Protheus.doc} A314MonT4
Tela para informação de operador / equipe.

@param   cTPRP  Tipo de requisição
                1=Operador
                2=Equipe

@return  lRet   Retorna se os dados foram inseridos.

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function A314MonT4(cTPRP)
   Local nLin  := 0
   Local lRet  := .T.

   VtClear()
   @nLin,00 VtSay STR0001 //"Apontamento Produção"
   nLin++
   If cTPRP == '2'
      @nLin,00 VtSay STR0031 //"Reporte por Operador"
      nLin+=2
      @nLin,00 VtSay STR0032 //"Operador"
      nLin++
      @nLin,00 VtGet cCDOERP  Pict "@!" Valid A314VlOE(1,cCDOERP,cCDMQ) F3 "CYC007"
   Else
      @nLin,00 VtSay STR0033 //"Reporte por Equipe"
      nLin+=2
      @nLin,00 VtSay STR0034 //"Equipe"
      nLin++
      @nLin,00 VtGet cCDOERP  Pict "@!" Valid A314VlOE(2,cCDOERP,cCDMQ) F3 "CYC008"
   Endif

   VtRead()

   If VTLastkey() == 27    // Tecla ESC
      lRet := .F.
   EndIf
Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} A314VlOE
Função para validação de operador e equipe

@param   nTipo  Tipo de requisição
                  1=Operador
                  2=Equipe
         cCDOERP  Código da equipe/operador
         cCDMQ    Código da máquina

@return  lRet     Retorna se o operador/equipe é válido.

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function A314VlOE(nTipo,cCDOERP,cCDMQ)
Default lAutoMacao := .F.

   dbSelectArea("CYH")
   CYH->(dbSetOrder(3))

   If !(CYH->(dbSeek(xFilial("CYH")+cCDOERP)))
      If nTipo == 1
         IF !lAutoMacao
            VtAlert(STR0035, STR0010) //"Operador não cadastrado." ### "Erro"
         ENDIF
      Else
         VtAlert(STR0036, STR0010) //"Equipe não cadastrada." ### "Erro"
      Endif
      Return .F.
   Else
      If nTipo == 1
         If CYH->CYH_TPRC != '1'
            VtAlert(STR0037, STR0010) //"Recurso selecionado não é do tipo 'Operador'." ### "Erro"
            Return .F.
         Endif
      Endif
      If nTipo == 2
         If CYH->CYH_TPRC != '3'
            VtAlert(STR0038, STR0010) //"Recurso selecionado não é do tipo 'Equipe'." ### "Erro"
            Return .F.
         Endif
      Endif
   Endif

   dbSelectArea("CYC")
   CYC->(dbSetOrder(2))
   If !(CYC->(dbSeek(xFilial("CYC")+cCDMQ+CYH->CYH_IDRC)))
      If nTipo == 1
         VtAlert(STR0039, STR0010) //"Operador não relacionado à máquina informada." ### "Erro"
      Else
         VtAlert(STR0040, STR0010) //"Equipe não relacionado à máquina informada." ### "Erro"
      Endif
      Return .F.
   Endif

Return .T.

//-------------------------------------------------------------------
/*/{Protheus.doc} A314MonT5
Tela para solicitação de data/hora inicial

@param

@return  lRet   Retorna se a data/hora foi informada

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function A314MonT5()

If cTPTR != "1" //Inicializar split
   Return .T.
Endif

VtClear()
If !GetMV('MV_SFCAPON')
   @0,0 VTSAY STR0001 //Apontamento da Produção
   @2,0 VTSAY STR0041 //"Data Início"
   @3,0 VTGET dDtIn Pict "99/99/9999" Valid A314VlDtHr( dDtIn, 1, .F. )
   @4,0 VTSAY STR0042
   @5,0 VTGET cHrIn Pict "99:99:99" Valid A314VlDtHr( cHrIn, 2, .T. )
   VTREAD()
   If VTLastkey() == 27    // Tecla ESC
      Return .F.
   EndIf
EndIf

If !vldHrInic()
	Return .F.
EndIf

Return .T.

//-------------------------------------------------------------------
/*/{Protheus.doc} A314MonT6
Tela para solicitação de data/hora final

@param

@return  lRet   Retorna se a data/hora foi informada

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function A314MonT6()

If cTPTR != "2" .And. !lFinIni //Envia informações para o Chão de Fábrica
   Return .T.
Endif

VtClear()
If !GetMV('MV_SFCAPON')
   @0,0 VTSAY STR0001 //Apontamento da Produção
   @2,0 VTSAY STR0043 //"Data Final"
   @3,0 VTGET dDtFi Pict "99/99/9999" Valid A314VlDtHr( dDtFi, 1, .F. )  // "Data Fim"
   @4,0 VTSAY STR0044 //"Hora Final"
   @5,0 VTGET cHrFi Pict "99:99:99" Valid A314VlDtHr( cHrFi, 2, .F. ) // "Hora Fim"
   VTREAD()
   If VTLastkey() == 27    // Tecla ESC
      Return .F.
   EndIf
EndIf

Return .T.

//-------------------------------------------------------------------
/*/{Protheus.doc} A314VlDtHr
Tela para solicitação de data/hora final

@param     cCampo  Valor que será validado
           nTipo   1=Data
                   2=Hora

@return  lRet   Retorna se a data/hora é valida

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function A314VlDtHr(cCampo, nTipo, lValApon)
Local lRet := .T.

// Valida Data
If nTipo == 1
   If Day ( cCampo ) == 0
      VTAlert( STR0045, STR0010 ) //"Data inválida." ### "Erro"
      lRet := .F.
   EndIf

// Valida Hora
ElseIf nTipo == 2
   aHora := StrTokArr( cCampo, ':' )
   If len(aHora) < 3 .Or. !ValHora(aHora[1]) .Or. !ValHora(aHora[2]) .Or. !ValHora(aHora[3])
      VTAlert( STR0046, STR0010 ) //"Hora inválida." ### "Erro"
      lRet := .F.
   ElseIf Val(aHora[1]) > 23 .Or. Val(aHora[2]) > 59 .Or. Val(aHora[3]) > 59
      VTAlert( STR0046, STR0010 ) //"Hora inválida." ### "Erro"
      lRet := .F.
   EndIf
EndIf

If lRet .And. lValApon
   lRet := vldHrInic()
EndIf

Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} ValHora
Validação de campo tipo hora

@param   cHora  Campo hora

@return  lRet   Retorna se a hora é válida

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function ValHora(cHora)
Local nI := 0

For nI := 1 To Len(cHora)
   If !(SUBSTR(cHora,nI,1) $ "1234567890")
      Return .F.
   Endif
Next
Return .T.

//-------------------------------------------------------------------
/*/{Protheus.doc} ACDV314FOE
Filtro para Consulta Padrão Máquina x Operador, Máquina x Equipe

@param   cTipo  Tipo do recurso

@return  lRet   Retorna se o recurso é válido

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Function ACDV314FOE(cTipo)
Local lRet      := .T.

If CYC->CYC_CDMQ == cCDMQ .AND. CYC->CYC_TPRC == cTipo
   lRet := .T.
Else
   lRet := .F.
Endif

Return lRet

Static Function SumTime(cTime)
Return SFCXSegtoHour(SFCXHourToSeg(cTime)+1)

//Filtro de Maquina por Centro de Trabalho
Function ACDV314FCT()
   If CYB->CYB_CDCETR != cCeTrab
      Return .F.
   EndIf
Return .T.

//-------------------------------------------------------------------
/*/{Protheus.doc} A314MonT7
Tela para informação de ferramenta

@param
@return  lRet     Retorna se os dados foram inseridos.

@author  Lucas Konrad França
@since   21/01/2016
@version P118
/*/
//-------------------------------------------------------------------
Static Function A314MonT7()
   Local lRet     := .T.

   VtClear()
   @00,00 VtSay STR0001 //"Apontamento Produção"
   @01,00 VtSay STR0052 //"Reporte de Ferramenta"
   @03,00 VtSay STR0053 //"Ferramenta"
   @04,00 VtGet cCdFer Pict "@!" Valid A314VlFE() F3 "CYC009"
   VtRead()

   If VTLastkey() == 27    // Tecla ESC
      lRet := .F.
   EndIf
Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} A314VlFE
Validação da ferramenta informada.

@param
@return  lRet     Retorna se a ferramenta é válida.

@author  Lucas Konrad França
@since   21/01/2016
@version P118
/*/
//-------------------------------------------------------------------
Static Function A314VlFE()
Default lAutoMacao := .F.

   dbSelectArea("CYH")
   CYH->(dbSetOrder(3))
   If !(CYH->(dbSeek(xFilial("CYH")+cCdFer)))
      IF !lAutoMacao
         VtAlert(STR0051, STR0010) //"Ferramenta não cadastrada" ### "Erro"
      ENDIF
      Return .F.
   Else
      If CYH->CYH_TPRC != '2'
         VtAlert(STR0050, STR0010) //"Recurso selecionado não é do tipo 'Ferramenta'" ### "Erro"
         Return .F.
      Endif
   Endif

   dbSelectArea("CYC")
   CYC->(dbSetOrder(2))
   If !(CYC->(dbSeek(xFilial("CYC")+cCDMQ+CYH->CYH_IDRC)))
      VtAlert(STR0049, STR0010) //"Ferramenta não relacionada a máquina informada." ### "Erro"
      Return .F.
   Endif

Return .T.

//-------------------------------------------------------------------
/*/{Protheus.doc} ACDV314CYO
Função de filtro para a consulta padrão CYO002

@param

@return  lRet - Exibe ou não o registro na consulta padrão.

@author  Lucas Konrad França
@since   21/01/2016
@version P118
/*/
//-------------------------------------------------------------------
Function ACDV314CYO()
   Local lRet := .T.

   If !CYO->CYO_LGRFMP
      lRet := .F.
   EndIf
Return lRet
