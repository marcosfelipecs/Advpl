//Bibliotecas
#Include "Protheus.ch"
#Include "TopConn.ch"
  
/*/{Protheus.doc} zTstBar
Função de exemplo de barras de processamento em AdvPL
@author Atilio
@since 28/10/2018
@version 1.0
@type function
@example u_zTstBar()
/*/
  
User Function zTstBar()
    Local aArea      := GetArea()
    Local lContinua  := .T.
    Local nTipoRegua := 0
    Local oProcess
    Private cQryAux  := ""
      
    //Monta a consulta de grupo de produtos
    cQryAux := " SELECT "                          + CRLF
    cQryAux += "     BM_GRUPO, "                    + CRLF
    cQryAux += "     BM_DESC "                      + CRLF
    cQryAux += " FROM "                            + CRLF
    cQryAux += "     SBM010 SBM "                   + CRLF
    cQryAux += " WHERE "                           + CRLF
    cQryAux += "     BM_FILIAL = ' ' "              + CRLF
    cQryAux += "     AND SBM.D_E_L_E_T_ = ' ' "     + CRLF
      
    //Enquanto houver testes
    While lContinua
        nTipoRegua := 0
        nTipoRegua := Aviso('Atenção', 'Qual tipo gostaria de Testar?', {'MsAguarde', 'MsNewProcess', 'MsgRun', 'RptStatus', 'Processa'}, 2)
          
        //Conforme botão selecionado, monta a régua
        If nTipoRegua == 1
            MsAguarde({|| fExemplo1()}, "Aguarde...", "Processando Registros...")
              
        ElseIf nTipoRegua == 2
            oProcess := MsNewProcess():New({|| fExemplo2(oProcess)}, "Processando...", "Aguarde...", .T.)
            oProcess:Activate()
              
        ElseIf nTipoRegua == 3
            fExemplo3()
              
        ElseIf nTipoRegua == 4
            RptStatus({|| fExemplo4()}, "Aguarde...", "Executando rotina...")
              
        ElseIf nTipoRegua == 5
            Processa({|| fExemplo5()}, "Filtrando...")
        EndIf
          
        lContinua := MsgYesNo("Continua testando?", "Atenção")
    EndDo
      
    RestArea(aArea)
Return
  
/*-----------------------------------------------------------*
 | Func.: fExemplo1                                          |
 | Desc.: Exemplo utilizando MsAguarde                       |
 *-----------------------------------------------------------*/
  
Static Function fExemplo1()
    Local aArea  := GetArea()
    Local nAtual := 0
    Local nTotal := 0
      
    //Executa a consulta
    TCQuery cQryAux New Alias "QRY_AUX"
      
    //Conta quantos registros existem, e seta no tamanho da régua
    Count To nTotal
      
    //Percorre todos os registros da query
    QRY_AUX->(DbGoTop())
    While ! QRY_AUX->(EoF())
          
        //Incrementa a mensagem na régua
        nAtual++
        MsProcTxt("Analisando registro " + cValToChar(nAtual) + " de " + cValToChar(nTotal) + "...")
          
        QRY_AUX->(DbSkip())
    EndDo
    QRY_AUX->(DbCloseArea())
      
    RestArea(aArea)
Return
  
/*-----------------------------------------------------------*
 | Func.: fExemplo2                                          |
 | Desc.: Exemplo utilizando MsNewProcess                    |
 *-----------------------------------------------------------*/
  
Static Function fExemplo2(oObj)
    Local aArea  := GetArea()
    Local nAtual := 0
    Local nTotal := 0
    Local nAtu2  := 0
    Local nTot2  := 90
      
    //Executa a consulta
    TCQuery cQryAux New Alias "QRY_AUX"
      
    //Conta quantos registros existem, e seta no tamanho da régua
    Count To nTotal
    oObj:SetRegua1(nTotal)
      
    //Percorre todos os registros da query
    QRY_AUX->(DbGoTop())
    While ! QRY_AUX->(EoF())
          
        //Incrementa a mensagem na régua
        nAtual++
        oObj:IncRegua1("Analisando registro " + cValToChar(nAtual) + " de " + cValToChar(nTotal) + "...")
          
        //Incrementando a régua 2
        oObj:SetRegua2(nTot2)
        For nAtu2 := 1 To nTot2
            oObj:IncRegua2("Posição " + cValToChar(nAtu2) + " de " + cValToChar(nTot2) + "...")
        Next
          
        QRY_AUX->(DbSkip())
    EndDo
    QRY_AUX->(DbCloseArea())
      
    RestArea(aArea)
Return
  
/*-----------------------------------------------------------*
 | Func.: fExemplo3                                          |
 | Desc.: Exemplo utilizando MsgRun                          |
 *-----------------------------------------------------------*/
  
Static Function fExemplo3()
    Local aArea  := GetArea()
    Local nTotal := 0
      
    //Executa a consulta
    TCQuery cQryAux New Alias "QRY_AUX"
      
    //Chamando a régua que irá executar o bloco de código (como um aEval, DbEval, etc)
    MsgRun("Lendo tabela...", "Título", {|| QRY_AUX->(DbEval({|x| nTotal++})) })
    QRY_AUX->(DbCloseArea())
      
    MsgInfo("Processado: " + cValToChar(nTotal) + " registro(s)", "Atenção")
      
    RestArea(aArea)
Return
  
/*-----------------------------------------------------------*
 | Func.: fExemplo4                                          |
 | Desc.: Exemplo utilizando RptStatus                       |
 *-----------------------------------------------------------*/
  
Static Function fExemplo4()
    Local aArea  := GetArea()
    Local nAtual := 0
    Local nTotal := 0
      
    //Executa a consulta
    TCQuery cQryAux New Alias "QRY_AUX"
      
    //Conta quantos registros existem, e seta no tamanho da régua
    Count To nTotal
    SetRegua(nTotal)
      
    //Percorre todos os registros da query
    QRY_AUX->(DbGoTop())
    While ! QRY_AUX->(EoF())
          
        //Incrementa a mensagem na régua
        nAtual++
        IncRegua()
          
        QRY_AUX->(DbSkip())
    EndDo
    QRY_AUX->(DbCloseArea())
      
    RestArea(aArea)
Return
  
/*-----------------------------------------------------------*
 | Func.: fExemplo5                                          |
 | Desc.: Exemplo utilizando Processa                        |
 *-----------------------------------------------------------*/
  
Static Function fExemplo5()
    Local aArea  := GetArea()
    Local nAtual := 0
    Local nTotal := 0
      
    //Executa a consulta
    TCQuery cQryAux New Alias "QRY_AUX"
      
    //Conta quantos registros existem, e seta no tamanho da régua
    Count To nTotal
    ProcRegua(nTotal)
      
    //Percorre todos os registros da query
    QRY_AUX->(DbGoTop())
    While ! QRY_AUX->(EoF())
          
        //Incrementa a mensagem na régua
        nAtual++
        IncProc("Analisando registro " + cValToChar(nAtual) + " de " + cValToChar(nTotal) + "...")
          
        QRY_AUX->(DbSkip())
    EndDo
    QRY_AUX->(DbCloseArea())
    RestArea(aArea)
Return