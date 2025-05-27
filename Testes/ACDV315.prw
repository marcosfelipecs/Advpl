#INCLUDE "PROTHEUS.CH"
#INCLUDE "APVT100.CH"
#INCLUDE "ACDV315.CH"

#DEFINE NEWLINE CHR(13)+CHR(10)

//-------------------------------------------------------------------
/*/{Protheus.doc} ACDV315
Tela de Divis�o e Aloca��o de split

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
      VTAlert(STR0029, STR0016) //"Integra��o com o ACD desativada. Verifique o par�metro MV_INTACD" ### 'Erro'
      Return
   EndIf

   // Monta a tela
   @0,0 VTSAY STR0003 // "Fun��es Split"

   // Monta a lista de opcoes
   While nMenu < 0
      nMenu := VTAChoice( 2, 0, VTMaxRow(), VTMaxCol(), aOpcoes, , "ACDV315VLD", nMenu )
   End
Return Nil
//-------------------------------------------------------------------
/*/{Protheus.doc} ACDV315VLD
Valida��o da sele��o do menu

@param   nModo  Modos do VTAChoice
                  0 - Inativo
                  1 - Tentativa de passar in�cio da lista
                  2 - Tentativa de passar final da lista
                  3 - Normal
                  4 - Itens n�o selecionados
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
Tela de aloca��o/divis�o de split

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
      @00,00 VtSay STR0004 //"Aloca��o de Split"
   ElseIf nPosicao == 2
      @00,00 VtSay STR0005 //"Divis�o de Split"
   ElseIf nPosicao == 3
      //Condi��o para processo de estornar apontamento
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
   @03,00 VtSay STR0007 VtGet cIDAT   Pict "@!"   Valid A315VlOpe(cNRORPO,cIDAT,nPosicao) F3 "CY9002" //"Opera��o"
   VtRead()
   If VTLastkey() == 27   // Tecla ESC
      Return Nil
   Endif

   //Se n�o tiver sido informada opera��o, e estiver na opera��o de cancelar movimentos
   //deve cancelar todos os movimentos da ordem. Nesse caso n�o solicita o c�digo do split.
   If !(Empty(cIDAT) .And. nPosicao == 4)
      @04,00 VtSay STR0008 VtGet cIDATQO Pict "@!"   Valid A315VlSp(cNRORPO,cIDAT,cIDATQO,nPosicao) F3 "CYY003" //"Split"
      VtRead()
   EndIf

   If VTLastkey() == 27    // Tecla ESC
      Return Nil
   Endif
   
   If nPosicao == 4
      //Cancelamento de movimentos.
      If VtYesNo(STR0033 + AllTrim(cNRORPO) + STR0034 ,STR0035,.T.) //"Ser�o cancelados os movimentos realizados na ordem 'XXX'. Confirma?" ### "Aten��o"
         a315Cancel(cNRORPO,cIDAT,cIDATQO)
         VtAlert(STR0036,STR0012) //"Movimentos cancelados com sucesso." ### "Sucesso"
      Else
         VtAlert(STR0037,STR0038) //"Processamento n�o efetuado." ### "Fim"
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
   
      If nPosicao == 1 //Aloca��o de split
         @05,00 VtSay STR0009 VtGet cCDMQ    Pict "@!" Valid A315VlMq(cCDMQ,CYY->CYY_CDCETR) F3 "CYB003" //"M�quina"
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
            VtAlert(STR0011,STR0012) //"Aloca��o de split efetuada com sucesso" ### "Sucesso"
         Else
            VtAlert(STR0013,STR0012) //"Divis�o de split efetuada com sucesso."
         Endif
      Endif
   EndIf

Return Nil
//-------------------------------------------------------------------
/*/{Protheus.doc} A315VlOp
Retorna se ordem de produ��o � v�lida para apontamento

@param   cOrdem  Identifica a ordem de produ��o informada
@param   nOperac Identifica a opera��o que est� sendo executada.
                 1 - Aloca��o de split
                 2 - Divis�o de split
                 3 - Estorno de apontamento
                 4 - Cancelamento de movimentos

@return  lRet    Retorna se valor informado � v�lido

@author  Monique Madeira Pereira
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function A315VlOp(cOrdem,nOperac)
   Local lRet      := .T.

   If Empty(cOrdem)
      VTAlert(STR0014, STR0016) //"Ordem de produ��o deve ser preenchida." ### "Erro"
      Return .F.
   EndIf

   DbSelectArea("CYQ")
   CYQ->(dbSetOrder(1))
   If !(CYQ->(dbSeek(xFilial("CYQ")+cOrdem)))
      VTAlert(STR0015, STR0016) // "Ordem de produ��o n�o existente." ### "Erro"
      lRet := .F.
   Endif

   If lRet .And. nOperac == 4
      If !ExistCZH(cOrdem)
         VTAlert(STR0039,STR0016) //"N�o existem movimentos para cancelar nesta ordem de produ��o." ### "Erro"
         lRet := .F.
      EndIf
   EndIf

Return lRet
//-------------------------------------------------------------------
/*/{Protheus.doc} A315VlOpe
Retorna se a opera��o � v�lida para apontamento

@param     cOp      Identifica a ordem informada
           cOper    Identifica a opera��o informada
@param     nOperac  Identifica a opera��o que est� sendo executada.
                      1 - Aloca��o de split
                      2 - Divis�o de split
                      3 - Estorno de apontamento
                      4 - Cancelamento de movimentos
                 
@return  lRet   Retorna se valor informado � v�lido

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
    //Se for cancelamento de movimento, questiona se ser�o cancelados os movimentos de todas as opera��es.
    If VtYesNo(STR0040,STR0035,.T.) //"Opera��o n�o informada. Cancelar os movimentos de todas as opera��es da ordem?" ### "Aten��o"
        lRet := .T.
    Else
        lRet := .F.
        VTAlert(STR0041,STR0016) // "Opera��o n�o informada." ### "Erro"
    EndIf
EndIf

If lRet .And. !Empty(cOper)
   DbSelectArea("CY9")
   CY9->(dbSetOrder(1))
   If !(CY9->(dbSeek(xFilial("CY9")+cOp+cOper)))
      VTAlert(STR0018, STR0016) //"Opera��o n�o existente para a ordem informada." ### "Erro"
      lRet := .F.
   Endif
EndIf

If lRet .And. !Empty(cOper) .And. nOperac == 4
    If !ExistCZH(cOp,cOper)
        VTAlert(STR0042,STR0016) //"N�o existem movimentos para cancelar nesta ordem de produ��o/opera��o." ### "Erro"
        lRet := .F.
    EndIf
EndIf

Return lRet
//-------------------------------------------------------------------
/*/{Protheus.doc} A315VlSp
Retorna se split � v�lido para apontamento

@param   cOp      Identifica a ordem de produ��o informada
         cOper    Identifica a opera��o informada
         cSplit   Identifica o split informado
@param   nOperac  Identifica a opera��o que est� sendo executada.
                      1 - Aloca��o de split
                      2 - Divis�o de split
                      3 - Estorno de apontamento
                      4 - Cancelamento de movimentos

@return  lRet   Retorna se valor informado � v�lido

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
      //Se for cancelamento de movimento, questiona se ser�o cancelados os movimentos de todos os splits.
      If VtYesNo(STR0043,STR0035,.T.) //"Split n�o informado. Cancelar os movimentos de todos os splits da opera��o?" ### "Aten��o"
         Return .T.
      Else
        VtAlert(STR0045,STR0016) //"Split n�o informado." ### "Erro"
        Return .F.
      EndIf
   EndIf

   dbSelectArea("CYY")
   CYY->(dbSetOrder(1))
   CYY->(dbGoTop())
   If CYY->(dbSeek(xFilial("CYY")+cOp+cOper+cSplit))
      cCDMQ := CYY->CYY_CDMQ
   Else
      VTAlert(STR0019, STR0016) //"Split inexistente para a opera��o informada" ### "Erro"
      lRet := .F.
   Endif
   
   If lRet .And. nOperac <> 4
      //Verifica se a ordem j� est� com a produ��o iniciada.
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
         VTAlert(STR0030 + AllTrim((cNextAlias)->(CZH_CDMQ)) + ".",STR0016) //"N�o � poss�vel alocar este split. Ordem de produ��o j� est� iniciada na m�quina ### " "Erro"
         lRet := .F.
      EndIf
      (cNextAlias)->(dbCloseArea())
   EndIf

   If lRet .And. nOperac == 4
      If !ExistCZH(cOp,cOper,cSplit)
         VTAlert(STR0044,STR0016) //"N�o existem movimentos para cancelar nesta ordem de produ��o/opera��o/split." ### "Erro"
         lRet := .F.
      EndIf
   EndIf

Return lRet
//-------------------------------------------------------------------
/*/{Protheus.doc} A315VlMq
Retorna se m�quina � v�lida para aloca��o

@param   cCDMQ   Identifica a m�quina para aloca��o
         cCDCETR Identifica o centro de trabalho do split

@return  lRet   Retorna se valor informado � v�lido

@author  Monique Madeira Pereira
@since   01/09/2013
@version  P12
/*/
//-------------------------------------------------------------------
Static Function A315VlMq(cCDMQ, cCDCETR)
   dbSelectArea('CYB')
   CYB->(dbSetOrder(1))
   If !(CYB->(dbSeek(xFilial('CYB')+cCDMQ)))
      VtAlert(STR0020, STR0016) //"M�quina n�o cadastrada." ### "Erro"
      Return .F.
   Else
      If CYB->CYB_CDCETR != cCDCETR
         VtAlert(STR0021, STR0016) //"M�quina n�o pertece ao centro de trabalho da ordem de produ��o." ### "Erro"
         Return .F.
      Endif
   Endif
Return .T.
//-------------------------------------------------------------------
/*/{Protheus.doc} ValidCampo
Verifica se o campo � v�lido

@param   cCodigo Valor para ser validado
         nBusca  Tipo de Valida��o

@return  lRet   Retorna se valor informado � v�lido

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
         VTAlert( STR0026, STR0016 ) //"Sequ�ncia de Reporte Inv�lida"
         lRet := .F.
      EndIf
   ElseIf nBusca == 2 .AND. Day ( CTOD(cCodigo) ) == 0
      VTAlert( STR0027, STR0016 ) //"Data inv�lida."
      lRet := .F.
   EndIf
Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} ExistCZH
Verifica se existe registro na tabela CZH.

@param   cOp     N�mero da ordem de produ��o
         cOperac ID da opera��o
         cSplit  ID do split

@return  lRet   Indica se existe ou n�o registro na tabela CZH

@author  Lucas Konrad Fran�a
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
Cancela os movimentos da tabela CZH, de acordo com as informa��es digitadas em tela.

@param   cOp     N�mero da ordem de produ��o
         cOperac ID da opera��o
         cSplit  ID do split

@return  Nil

@author  Lucas Konrad Fran�a
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