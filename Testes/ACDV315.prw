#INCLUDE "PROTHEUS.CH"
#INCLUDE "APVT100.CH"
#INCLUDE "ACDV315.CH"

#DEFINE NEWLINE CHR(13)+CHR(10)

//-------------------------------------------------------------------
/*/{Protheus.doc} ACDV315
Tela de Divisão e Alocação de split

@author  Monique Madeira Pereira
@since   21/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Function ACDV315()
   Local aOpcoes := { STR0001, STR0002, STR0022, STR0031 } // "Alocar Split" ### "Dividir Split" ### "Estornar apontamento" ### "Cancelar Movimento"
   Local nMenu   := -1
   Private cCeTrab

   If SuperGetMv('MV_INTACD',.F.,"0") != "1"
      VTAlert(STR0029, STR0016) //"Integração com o ACD desativada. Verifique o parâmetro MV_INTACD" ### 'Erro'
      Return
   EndIf

   // Monta a tela
   @0,0 VTSAY STR0003 // "Funções Split"

   // Monta a lista de opcoes
   While nMenu < 0
      nMenu := VTAChoice( 2, 0, VTMaxRow(), VTMaxCol(), aOpcoes, , "ACDV315VLD", nMenu )
   End
Return Nil
//-------------------------------------------------------------------
/*/{Protheus.doc} ACDV315VLD
Validação da seleção do menu

@param   nModo  Modos do VTAChoice
                  0 - Inativo
                  1 - Tentativa de passar início da lista
                  2 - Tentativa de passar final da lista
                  3 - Normal
                  4 - Itens não selecionados
@param   nPosicao  Item selecionado em tela

@return nReturn Retorna 0 para sair da tela

@author  Monique Madeira Pereira
@since   21/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Function ACDV315VLD( nModo, nPosicao )
   Private cNRORPO   := Space(TamSX3("CYY_NRORPO")[1])
   Private cIDAT     := Space(TamSX3("CYY_IDAT")[1])
   Private cNRSQRP   := Space(TamSX3("CYV_NRSQRP")[1])
   Private cData     := DTOC(Date())
   Private cIDATQO   := Space(TamSX3("CYY_IDATQO")[1])
   Private cCDMQ     := Space(TamSX3("CYY_CDMQ")[1])
   Private nNovo     := 0
   Private nOriginal := 0
   Private nPrevista := 0

   If nModo == 3
      If VTLastkey() == 27       // Tecla ESC
         Return 0
      Else
         ACDV315T(nPosicao)
         Return 0
      EndIf
   EndIf

Return -1
//-------------------------------------------------------------------
/*/{Protheus.doc} ACDV315T
Tela de alocação/divisão de split

@param   nPosicao  Item selecionado em tela

@return

@author  Monique Madeira Pereira
@since   21/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Function ACDV315T(nPosicao)
   Local nLin  := 0
   Local lRet  := .T.
   Local oModelCYV

   VtClear()
   If nPosicao == 1
      @00,00 VtSay STR0004 //"Alocação de Split"
   ElseIf nPosicao == 2
      @00,00 VtSay STR0005 //"Divisão de Split"
   ElseIf nPosicao == 3
      //Condição para processo de estornar apontamento
      @00,00 VtSay STR0023 //"Estorno Apontamento"
      @02,00 VtSay STR0024 //"Sequencia Reporte"
      @03,00 VtGet cNRSQRP    Pict "@!"  F3 "CYV004" Valid ValidCampo(cNRSQRP, 1 )
      VtRead

      If VTLastkey() == 27   // Tecla ESC
         Return Nil
      Endif

      dbSelectArea("CYV")
      CYV->(dbSetOrder(1))
      CYV->(dbGoTop())
      CYV->(dbSeek(xFilial("CYV")+cNRSQRP))

      VtClear()

      @00,00 VtSay STR0023 //"Estorno Apontamento"

      @02,00 VtSay STR0006+": "+CYV->(CYV_NRORPO) //Ordem
      @03,00 VtSay STR0007+": "+CYV->(CYV_IDAT) //Operacao
      @04,00 VtSay STR0008+": "+CYV->(CYV_IDATQO) //Split
      @05,00 VtSay STR0009+": "+CYV->(CYV_CDMQ) //Maquina

      @06,00 VtSay STR0025
      @07,00 VtGet cData Pict "99/99/9999" Valid ValidCampo(cData, 2 ) //"Data Estorno"
      VtRead()

      If VTLastkey() == 27   // Tecla ESC
         Return Nil
      Endif

      SFCA313OK(cNRSQRP,CTOD(cData))

      VtAlert(STR0028,STR0012) //"Estorno do Apontamento efetuado com sucesso"

      Return Nil
   ElseIf nPosicao == 4
      @00,00 VtSay STR0032 //"Cancelar movimentos"
   Endif

   @02,00 VtSay STR0006 VtGet cNRORPO Pict "@!"  Valid A315VlOp(cNRORPO,nPosicao) F3 "CYQ001" //"Ordem"
   VtRead()

   If VTLastkey() == 27   // Tecla ESC
      Return Nil
   Endif
   @03,00 VtSay STR0007 VtGet cIDAT   Pict "@!"   Valid A315VlOpe(cNRORPO,cIDAT,nPosicao) F3 "CY9002" //"Operação"
   VtRead()
   If VTLastkey() == 27   // Tecla ESC
      Return Nil
   Endif

   //Se não tiver sido informada operação, e estiver na operação de cancelar movimentos
   //deve cancelar todos os movimentos da ordem. Nesse caso não solicita o código do split.
   If !(Empty(cIDAT) .And. nPosicao == 4)
      @04,00 VtSay STR0008 VtGet cIDATQO Pict "@!"   Valid A315VlSp(cNRORPO,cIDAT,cIDATQO,nPosicao) F3 "CYY003" //"Split"
      VtRead()
   EndIf

   If VTLastkey() == 27    // Tecla ESC
      Return Nil
   Endif
   
   If nPosicao == 4
      //Cancelamento de movimentos.
      If VtYesNo(STR0033 + AllTrim(cNRORPO) + STR0034 ,STR0035,.T.) //"Serão cancelados os movimentos realizados na ordem 'XXX'. Confirma?" ### "Atenção"
         a315Cancel(cNRORPO,cIDAT,cIDATQO)
         VtAlert(STR0036,STR0012) //"Movimentos cancelados com sucesso." ### "Sucesso"
      Else
         VtAlert(STR0037,STR0038) //"Processamento não efetuado." ### "Fim"
      EndIf
   Else
      //Posicionar no split selecionado
      dbSelectArea("CYY")
      CYY->(dbSetOrder(1))
      CYY->(dbGoTop())
      CYY->(dbSeek(xFilial("CYY")+cNRORPO+cIDAT+cIDATQO))
   
      cCeTrab := CYY->(CYY_CDCETR)
   
      nPrevista := CYY->CYY_QTAT
      nOriginal := if(CYY->CYY_QTATAP>0,CYY->CYY_QTATAP,CYY->CYY_QTAT)
      nNovo     := nPrevista - nOriginal
   
      If nPosicao == 1 //Alocação de split
         @05,00 VtSay STR0009 VtGet cCDMQ    Pict "@!" Valid A315VlMq(cCDMQ,CYY->CYY_CDCETR) F3 "CYB003" //"Máquina"
         VtRead()
         If VTLastkey() == 27    // Tecla ESC
            Return Nil
         Endif
      Endif
   
      @06,00 VtSay STR0010 //"Qtd Original"
      @07,00 VtGet nOriginal Pict "@E 999999999,9999"
      VtRead()
      If VTLastkey() == 27    // Tecla ESC
         Return Nil
      Endif
   
      lRet := IIf(nOriginal>nPrevista, nPrevista:=CYY->CYY_QTAT, nNovo:=nPrevista-nOriginal)
   
      If SFCA315OK(cNRORPO,cIDAT,cIDATQO,cCDMQ,nOriginal,nNovo,nPosicao,.F.)
         If nPosicao == 1
            VtAlert(STR0011,STR0012) //"Alocação de split efetuada com sucesso" ### "Sucesso"
         Else
            VtAlert(STR0013,STR0012) //"Divisão de split efetuada com sucesso."
         Endif
      Endif
   EndIf

Return Nil
//-------------------------------------------------------------------
/*/{Protheus.doc} A315VlOp
Retorna se ordem de produção é válida para apontamento

@param   cOrdem  Identifica a ordem de produção informada
@param   nOperac Identifica a operação que está sendo executada.
                 1 - Alocação de split
                 2 - Divisão de split
                 3 - Estorno de apontamento
                 4 - Cancelamento de movimentos

@return  lRet    Retorna se valor informado é válido

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function A315VlOp(cOrdem,nOperac)
   Local lRet      := .T.

   If Empty(cOrdem)
      VTAlert(STR0014, STR0016) //"Ordem de produção deve ser preenchida." ### "Erro"
      Return .F.
   EndIf

   DbSelectArea("CYQ")
   CYQ->(dbSetOrder(1))
   If !(CYQ->(dbSeek(xFilial("CYQ")+cOrdem)))
      VTAlert(STR0015, STR0016) // "Ordem de produção não existente." ### "Erro"
      lRet := .F.
   Endif

   If lRet .And. nOperac == 4
      If !ExistCZH(cOrdem)
         VTAlert(STR0039,STR0016) //"Não existem movimentos para cancelar nesta ordem de produção." ### "Erro"
         lRet := .F.
      EndIf
   EndIf

Return lRet
//-------------------------------------------------------------------
/*/{Protheus.doc} A315VlOpe
Retorna se a operação é válida para apontamento

@param     cOp      Identifica a ordem informada
           cOper    Identifica a operação informada
@param     nOperac  Identifica a operação que está sendo executada.
                      1 - Alocação de split
                      2 - Divisão de split
                      3 - Estorno de apontamento
                      4 - Cancelamento de movimentos
                 
@return  lRet   Retorna se valor informado é válido

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function A315VlOpe(cOp, cOper, nOperac)
Local lRet := .T.

If nOperac <> 4 .And. Empty(cOper)
   VTAlert(STR0017, STR0016) //"Operacao invalida" ### "Erro"
   Return .F.
EndIf

If nOperac == 4 .And. Empty(cOper)
    //Se for cancelamento de movimento, questiona se serão cancelados os movimentos de todas as operações.
    If VtYesNo(STR0040,STR0035,.T.) //"Operação não informada. Cancelar os movimentos de todas as operações da ordem?" ### "Atenção"
        lRet := .T.
    Else
        lRet := .F.
        VTAlert(STR0041,STR0016) // "Operação não informada." ### "Erro"
    EndIf
EndIf

If lRet .And. !Empty(cOper)
   DbSelectArea("CY9")
   CY9->(dbSetOrder(1))
   If !(CY9->(dbSeek(xFilial("CY9")+cOp+cOper)))
      VTAlert(STR0018, STR0016) //"Operação não existente para a ordem informada." ### "Erro"
      lRet := .F.
   Endif
EndIf

If lRet .And. !Empty(cOper) .And. nOperac == 4
    If !ExistCZH(cOp,cOper)
        VTAlert(STR0042,STR0016) //"Não existem movimentos para cancelar nesta ordem de produção/operação." ### "Erro"
        lRet := .F.
    EndIf
EndIf

Return lRet
//-------------------------------------------------------------------
/*/{Protheus.doc} A315VlSp
Retorna se split é válido para apontamento

@param   cOp      Identifica a ordem de produção informada
         cOper    Identifica a operação informada
         cSplit   Identifica o split informado
@param   nOperac  Identifica a operação que está sendo executada.
                      1 - Alocação de split
                      2 - Divisão de split
                      3 - Estorno de apontamento
                      4 - Cancelamento de movimentos

@return  lRet   Retorna se valor informado é válido

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function A315VlSp(cOp,cOper,cSplit,nOperac)
   Local lRet       := .T.
   Local cQuery     := ""
   Local cNextAlias := GetNextAlias()

   If Empty(cSplit) .And. nOperac == 4
      //Se for cancelamento de movimento, questiona se serão cancelados os movimentos de todos os splits.
      If VtYesNo(STR0043,STR0035,.T.) //"Split não informado. Cancelar os movimentos de todos os splits da operação?" ### "Atenção"
         Return .T.
      Else
        VtAlert(STR0045,STR0016) //"Split não informado." ### "Erro"
        Return .F.
      EndIf
   EndIf

   dbSelectArea("CYY")
   CYY->(dbSetOrder(1))
   CYY->(dbGoTop())
   If CYY->(dbSeek(xFilial("CYY")+cOp+cOper+cSplit))
      cCDMQ := CYY->CYY_CDMQ
   Else
      VTAlert(STR0019, STR0016) //"Split inexistente para a operação informada" ### "Erro"
      lRet := .F.
   Endif
   
   If lRet .And. nOperac <> 4
      //Verifica se a ordem já está com a produção iniciada.
      cQuery := " SELECT CZH.CZH_CDMQ "
      cQuery +=   " FROM " + RetSqlName("CZH") + " CZH "
      cQuery +=  " WHERE CZH.CZH_FILIAL = '" + xFilial("CZH") + "' "
      cQuery +=    " AND CZH.CZH_NRORPO = '" + cOp + "' "
      cQuery +=    " AND CZH.CZH_IDAT   = '" + cOper + "' "
      cQuery +=    " AND CZH.CZH_IDATQO = '" + cSplit + "' "
      cQuery +=    " AND CZH.CZH_TPTR   = '1' "
      cQuery +=    " AND CZH.CZH_STTR   = '1' "
      cQuery +=    " AND CZH.D_E_L_E_T_ = ' ' "
   
      cQuery := ChangeQuery(cQuery)
   
      dbUseArea( .T., 'TOPCONN', TcGenQry(,,cQuery), cNextAlias, .F., .F. )
      If (cNextAlias)->(!Eof())
         VTAlert(STR0030 + AllTrim((cNextAlias)->(CZH_CDMQ)) + ".",STR0016) //"Não é possível alocar este split. Ordem de produção já está iniciada na máquina ### " "Erro"
         lRet := .F.
      EndIf
      (cNextAlias)->(dbCloseArea())
   EndIf

   If lRet .And. nOperac == 4
      If !ExistCZH(cOp,cOper,cSplit)
         VTAlert(STR0044,STR0016) //"Não existem movimentos para cancelar nesta ordem de produção/operação/split." ### "Erro"
         lRet := .F.
      EndIf
   EndIf

Return lRet
//-------------------------------------------------------------------
/*/{Protheus.doc} A315VlMq
Retorna se máquina é válida para alocação

@param   cCDMQ   Identifica a máquina para alocação
         cCDCETR Identifica o centro de trabalho do split

@return  lRet   Retorna se valor informado é válido

@author  Monique Madeira Pereira
@since   01/09/2013
@version  P12
/*/
//-------------------------------------------------------------------
Static Function A315VlMq(cCDMQ, cCDCETR)
   dbSelectArea('CYB')
   CYB->(dbSetOrder(1))
   If !(CYB->(dbSeek(xFilial('CYB')+cCDMQ)))
      VtAlert(STR0020, STR0016) //"Máquina não cadastrada." ### "Erro"
      Return .F.
   Else
      If CYB->CYB_CDCETR != cCDCETR
         VtAlert(STR0021, STR0016) //"Máquina não pertece ao centro de trabalho da ordem de produção." ### "Erro"
         Return .F.
      Endif
   Endif
Return .T.
//-------------------------------------------------------------------
/*/{Protheus.doc} ValidCampo
Verifica se o campo é válido

@param   cCodigo Valor para ser validado
         nBusca  Tipo de Validação

@return  lRet   Retorna se valor informado é válido

@author  Ezequiel Ramos
@since   11/10/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function ValidCampo(cCodigo, nBusca)
   Local lRet := .T.

   If nBusca == 1
      dbSelectArea("CYV")
      CYV->(dbSetOrder(1))
      CYV->(dbGoTop())
      CYV->(dbSeek(xFilial("CYV")+cNRSQRP))

      If CYV->(EOF()) .OR. CYV->(CYV_LGRPEO) == .T.
         VTAlert( STR0026, STR0016 ) //"Sequência de Reporte Inválida"
         lRet := .F.
      EndIf
   ElseIf nBusca == 2 .AND. Day ( CTOD(cCodigo) ) == 0
      VTAlert( STR0027, STR0016 ) //"Data inválida."
      lRet := .F.
   EndIf
Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} ExistCZH
Verifica se existe registro na tabela CZH.

@param   cOp     Número da ordem de produção
         cOperac ID da operação
         cSplit  ID do split

@return  lRet   Indica se existe ou não registro na tabela CZH

@author  Lucas Konrad França
@since   03/09/2018
@version P12
/*/
//-------------------------------------------------------------------
Static Function ExistCZH(cOp,cOperac,cSplit)
    Local lRet      := .T.
    Local cAliasCZH := "BUSCACZH"
    Local cQuery    := ""
    Local aArea     := GetArea()

    cQuery := " SELECT COUNT(*) TOTAL "
    cQuery +=   " FROM " + RetSqlName("CZH") + " CZH "
    cQuery +=  " WHERE CZH.CZH_FILIAL = '" + xFilial("CZH") + "' "
    cQuery +=    " AND CZH.D_E_L_E_T_ = ' ' "
    cQuery +=    " AND CZH.CZH_STTR   = '1' "
    cQuery +=    " AND CZH.CZH_NRORPO = '" + cOp + "' "
    If !Empty(cOperac)
        cQuery += " AND CZH.CZH_IDAT = '" + cOperac + "' "
    EndIf
    If !Empty(cSplit)
        cQuery += " AND CZH.CZH_IDATQO = '" + cSplit + "' "
    EndIf

    cQuery := ChangeQuery(cQuery)
    dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasCZH,.T.,.T.)
      
    If (cAliasCZH)->(TOTAL) < 1
        lRet := .F.
    EndIf
    (cAliasCZH)->(dbCloseArea())

    RestArea(aArea)
Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} a315Cancel
Cancela os movimentos da tabela CZH, de acordo com as informações digitadas em tela.

@param   cOp     Número da ordem de produção
         cOperac ID da operação
         cSplit  ID do split

@return  Nil

@author  Lucas Konrad França
@since   03/09/2018
@version P12
/*/
//-------------------------------------------------------------------
Static Function a315Cancel(cOp,cOperac,cSplit)
    Local cAliasCZH := "PROCCZH"
    Local cQuery    := ""
    Local aArea     := GetArea()

    cQuery := " SELECT CZH.R_E_C_N_O_ REC "
    cQuery +=   " FROM " + RetSqlName("CZH") + " CZH "
    cQuery +=  " WHERE CZH.CZH_FILIAL = '" + xFilial("CZH") + "' "
    cQuery +=    " AND CZH.D_E_L_E_T_ = ' ' "
    cQuery +=    " AND CZH.CZH_STTR   = '1' "
    cQuery +=    " AND CZH.CZH_NRORPO = '" + cOp + "' "
    If !Empty(cOperac)
        cQuery += " AND CZH.CZH_IDAT = '" + cOperac + "' "
    EndIf
    If !Empty(cSplit)
        cQuery += " AND CZH.CZH_IDATQO = '" + cSplit + "' "
    EndIf

    cQuery := ChangeQuery(cQuery)
    dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasCZH,.T.,.T.)
      
    While !(cAliasCZH)->(EOF())
        CZH->(dbGoTo((cAliasCZH)->(REC)))

        RecLock("CZH",.F.)
            CZH->(dbDelete())
        CZH->(MsUnLock())

        (cAliasCZH)->(dbSkip())
    End

    (cAliasCZH)->(dbCloseArea())

    RestArea(aArea)
Return Nil