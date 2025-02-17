#INCLUDE "TOTVS.CH"
#INCLUDE "RESTFUL.CH"
#INCLUDE "TOPCONN.CH"
#INCLUDE "tbiconn.ch"

WSRESTFUL cte DESCRIPTION 'Post para integracao de CTE' FORMAT 'application/xml'
    WSMETHOD POST   DESCRIPTION 'Post para gravação de CTE'  WSSYNTAX '/acao/{}'
END WSRESTFUL

WSMETHOD POST WSSERVICE cte
    
    Local oResponse	:= JsonObject():new()     
    Local cBody     := self:getContent()
    Local lRet      := .T.
    Local cMensagem := ""
    Local lContinua := .T.
     
    ::SetContentType("application/json")   	
    nOpc := 2

    If len(::aURLParms) < 2
        oResponse["message"]    := 'acao e/ou chave nao enviada'
        oResponse["type"]       := "error"
        oResponse["status"]     := 400        
    Else
        cAcao   := ::aURLParms[1]
        cChave  := ::aURLParms[2]

        If ! Empty(cChave)

            If cAcao == 'inclusao' .OR. cAcao == 'cartacorrecao'
                If !(ExistZA4(cChave,cAcao))
                    cMensagem := PopulaZA4(cBody,cChave,cAcao)
                    If Empty(cMensagem)
                        oResponse["message"]    := cAcao + ' - concluido com sucesso'
                        oResponse["type"]       := "sucess"
                        oResponse["status"]     := 201
                    Else
                        oResponse["type"]       := "error"
                        oResponse["status"]     := 400
                        oResponse["message"]    := cAcao + ' - acao NAO permitida.'+cMensagem
                    EndIf
                Else
                    oResponse["type"]       := "error"
                    oResponse["status"]     := 400
                    oResponse["message"]    := "Chave ("+cChave+") já está cadastrada."
                EndIf
            Else
                If ! Empty(cAcao)
                
                    lExistZA4 := ExistZA4(cChave, 'inclusao')                    
                    
                    If cAcao == 'possocancelar' .OR. cAcao == 'cancelar' .OR. cAcao == 'inutilizar' .OR. cAcao == 'possoinutilizar'           

                        // If cAcao == 'possocancelar' .OR. cAcao == 'possoinutilizar' 
                        If cAcao == 'possocancelar' .OR. cAcao == 'possoinutilizar' .OR. cAcao == 'cancelar' .OR. cAcao == 'inutilizar'
                            // If !lExistZA4 .AND. cAcao == 'possocancelar'
                            If !lExistZA4 //.AND. cAcao == 'possocancelar'
                                oResponse["message"]    := cAcao + ' - acao NAO permitida. Título não está cadastrado no Protheus. '
                                oResponse["type"]       := "error"
                                oResponse["status"]     := 400 
                                lContinua := .F.
                            ElseIf (lExistZA4 .AND. QRY->ZA4_STATUS $ 'A,E') .OR. !lExistZA4 
                                oResponse["message"]    := cAcao + ' - acao permitida'
                                oResponse["type"]       := "sucess"
                                oResponse["status"]     := 201                        
                            ElseIf (lExistZA4 .AND. QRY->ZA4_STATUS == 'P') .AND. ! TemBaixa(cChave) //ValidaNota(cChave)
                                 //(analisar depois) como validar se pode cancelar/niutilizar?????
                                oResponse["message"]    := cAcao + ' - acao permitida'
                                oResponse["type"]       := "sucess"
                                oResponse["status"]     := 201      
                            ElseIf (lExistZA4 .AND. QRY->ZA4_STATUS == 'P')
                                //(analisar depois) como validar se pode cancelar/niutilizar?????
                                oResponse["message"]    := cAcao + ' - acao NAO permitida. Já feita a baixa do titulo. '
                                oResponse["type"]       := "error"
                                oResponse["status"]     := 400 
                                lContinua := .F.                                   
                            EndIf    

                            IF (cAcao == 'cancelar' .OR. cAcao == 'inutilizar')
                                If !(ExistZA4(cChave,cAcao))
                                    cMensagem := PopulaZA4(cBody,cChave,cAcao)
                                    If Empty(cMensagem)
                                        oResponse["message"]    := cAcao + ' - concluido com sucesso'
                                        oResponse["type"]       := "sucess"
                                        oResponse["status"]     := 201
                                    Else
                                        oResponse["type"]       := "error"
                                        oResponse["status"]     := 400
                                        oResponse["message"]    := cAcao + ' - acao NAO permitida.'+cMensagem
                                    EndIf

                                    IF !lContinua
                                        //TODO: Enviar email informando que está sendo cancelado ou inutilizado um CTE que no Protheus está com problemas
                                        cDest := SuperGetMv("P3_MAILCTE",,"thiago.brasil@lauto.com.br")
                                        cCopy := SuperGetMv("P3_COPYCTE",,"alexandre.varella@lauto.com.br")
                                        cMensagem := "A chave " + cChave + "está com uma inconsistência entre NUCCI x Protheus. Favor analisar! Possível erro: " + oResponse["message"]
                                        cFile := ""

                                        U_EnviaEmail("CTE LAUTO", cDest, cMensagem, cFile, cCopy, .T.)
                                    EndIf
                                Else
                                    oResponse["type"]       := "error"
                                    oResponse["status"]     := 400
                                    oResponse["message"]    := "Chave ("+cChave+") já está cadastrada."
                                ENDIF
                            EndIF

                        // ElseIf cAcao == 'cancelar' .OR. cAcao == 'inutilizar'                      
                        //     If !(ExistZA4(cChave,cAcao))
                        //         cMensagem := PopulaZA4(cBody,cChave,cAcao)
                        //         If Empty(cMensagem)
                        //             oResponse["message"]    := cAcao + ' - concluido com sucesso'
                        //             oResponse["type"]       := "sucess"
                        //             oResponse["status"]     := 201
                        //         Else
                        //             oResponse["type"]       := "error"
                        //             oResponse["status"]     := 400
                        //             oResponse["message"]    := cAcao + ' - acao NAO permitida.'+cMensagem
                        //         EndIf
                        //     Else
                        //         oResponse["type"]       := "error"
                        //         oResponse["status"]     := 400
                        //         oResponse["message"]    := "Chave ("+cChave+") já está cadastrada."
                        //     EndIf
                        EndIf                
                    EndIf
                Else
                    oResponse["message"]    := cAcao + ' - erro na solicitacao'
                    oResponse["type"]       := "error"
                    oResponse["status"]     := 400
                EndIf
            EndIf
        Else
            oResponse["message"]    := 'Chave nao enviada'
            oResponse["type"]       := "error"
            oResponse["status"]     := 400        
        EndIf
    EndIf

    self:SetResponse(EncodeUtf8(oResponse:toJson()))

Return lRet


//-------------------------------------------------------------------
/*/{Protheus.doc} SalvaCTE
Método utilzao para salvar xml das CTE para serem processados 
posteriormente
@author  Samuel Dantas
@since   19/11/2019
@version version
/*/
//-------------------------------------------------------------------
Static Function SalvaCTE (cXml, cChave, cAcao)
    Local lRet := .F.

    If !(ExistZA4(cChave,cAcao))
        RecLock('ZA4', .T.)                 
            ZA4->ZA4_BODY       := cXml
            ZA4->ZA4_DATA       := dDataBase
            ZA4->ZA4_HORA       := Time()
            ZA4->ZA4_CHAVE      := cChave
            ZA4->ZA4_ACAO       := cAcao
            ZA4->ZA4_STATUS     := 'A'
        ZA4->(MsUnLockAll())
        lRet := .T.
    EndIf
 
Return lRet

//------    -------------------------------------------------------------
/*/{Protheus.doc} ExistZA4
description
@author  Samuel Dantas
@since   19/11/2019
@version version
/*/
//-------------------------------------------------------------------
Static Function ExistZA4(cChave, cAcao)
    Local cQuery := ""
    Local lRet   := .F.

    cQuery := " SELECT ZA4_CHAVE, ZA4_STATUS, R_E_C_N_O_ AS RECNO FROM "  + RetSqlName('ZA4') + " ZA4"
    cQuery += " WHERE ZA4.D_E_L_E_T_ <> '*' AND RTRIM(ZA4_ACAO) = '"+cAcao+"' "
    cQuery += " AND RTRIM(ZA4_CHAVE) = '"+Alltrim(cChave)+"' "
    
    If Select('QRY') > 0
        QRY->(dbclosearea())
    EndIf
    
    TcQuery cQuery New Alias 'QRY'
    
    lRet :=  ! QRY->(Eof())
    
Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} ValidaNota
Valida se é possivel cancelar a nota.
@author  Samuel Dantas
@since   19/12/2019
@version version
/*/
//-------------------------------------------------------------------
Static Function ValidaNota(cChave)
    Local cQuery    := ""
    Local lRet      := .F.

    
    cQuery := " SELECT SE1.R_E_C_N_O_ AS RECNOSE1 FROM "  + RetSqlName('SF2') + " SF2 "
    cQuery += " INNER JOIN "  + RetSqlName('SE1') + " SE1 ON E1_FILIAL = F2_FILIAL AND E1_NUM = F2_DOC  "
    cQuery += " AND E1_PREFIXO = F2_SERIE AND E1_CLIENTE = F2_CLIENTE AND E1_LOJA = F2_LOJA AND SE1.D_E_L_E_T_ = SF2.D_E_L_E_T_"
    cQuery += " WHERE SF2.D_E_L_E_T_ <> '*' AND F2_CHVNFE = '"+cChave+"' AND E1_SALDO > 0  "
    
    If Select('QRYVLD') > 0
        QRYVLD->(dbclosearea())
    EndIf
    
    TcQuery cQuery New Alias 'QRYVLD'
    
    If QRYVLD->(!Eof())
        lRet := .T.
    EndIf
    
Return lRet


//-------------------------------------------------------------------
/*/{Protheus.doc} TEMBAIXA
Funcao para validar se a nfe sofreu ou nao baixas
@author  Sidney Sales
@since   15/04/2020
@version 1.0
/*/
//-------------------------------------------------------------------
Static Function TemBaixa(cChave)

    Local cQry   := ""
    Local lRet      := .F.

    cQry := " SELECT SE1.R_E_C_N_O_ AS RECNOSE1 FROM "  + RetSqlName('SF2') + " SF2 "
    cQry += " INNER JOIN "  + RetSqlName('SE1') + " SE1 ON E1_FILIAL = F2_FILIAL AND E1_NUM = F2_DOC  "
    cQry += " AND E1_PREFIXO = F2_SERIE AND E1_CLIENTE = F2_CLIENTE AND E1_LOJA = F2_LOJA AND SE1.D_E_L_E_T_ = SF2.D_E_L_E_T_"
    cQry += " WHERE SF2.D_E_L_E_T_ <> '*' AND F2_CHVNFE = '"+cChave+"' AND (E1_SALDO = 0 OR E1_BAIXA != '' ) "

    If Select('QRYBX') > 0
        QRYBX->(dbclosearea())
    EndIf

    TcQuery cQry New Alias 'QRYBX'

    If QRYBX->(!Eof())
        lRet := .T.
    EndIf

Return lRet

User Function LAUWS000

    If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01RN0001'
    EndIf
    
    ValidaNota('52191107189259001739570020000391161297098097')
    x:=1

Return

User Function TestCTE ()
    Local oResponse	:= JsonObject():new()     
    Local lContinua := .T.
    
    cAcao := "possocancelar"
    cChave := "31200107189259001810570020000111161732499053"

    If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01RN0001'
    EndIf
        
    nOpc := 2

    If cAcao == 'inclusao' .OR. cAcao == 'cartacorrecao'
        oResponse["message"]    := cAcao + ' - concluido com sucesso'
        oResponse["type"]       := "sucess"
        oResponse["status"]     := 201
        If !SalvaCTE(cBody, cChave, cAcao)
            oResponse["message"]    := "Chave ("+cChave+") já está cadastrada."
        EndIf
    Else
        If ! Empty(cAcao)
        
            lExistZA4 := ExistZA4(cChave, 'inclusao')                    
            
            If cAcao == 'possocancelar' .OR. cAcao == 'cancelar' .OR. cAcao == 'inutilizar' .OR. cAcao == 'possoinutilizar'           
                // If cAcao == 'possocancelar' .OR. cAcao == 'possoinutilizar' 
                If cAcao == 'possocancelar' .OR. cAcao == 'possoinutilizar' .OR. cAcao == 'cancelar' .OR. cAcao == 'inutilizar'  
                    // If !lExistZA4 .AND. cAcao == 'possocancelar'
                    If !lExistZA4 //.AND. cAcao == 'possocancelar'
                        oResponse["message"]    := cAcao + ' - acao NAO permitida. Título não está cadastrado no Protheus. '
                        oResponse["type"]       := "error"
                        oResponse["status"]     := 400 
                        lContinua := .F.
                    ElseIf (lExistZA4 .AND. QRY->ZA4_STATUS $ 'A,E') .OR. !lExistZA4 
                        oResponse["message"]    := cAcao + ' - acao permitida'
                        oResponse["type"]       := "sucess"
                        oResponse["status"]     := 201      
                    ElseIf (lExistZA4 .AND. QRY->ZA4_STATUS == 'P') .AND. ! TemBaixa(cChave) //ValidaNota(cChave)
                            //(analisar depois) como validar se pode cancelar/niutilizar?????
                        oResponse["message"]    := cAcao + ' - acao permitida'
                        oResponse["type"]       := "sucess"
                        oResponse["status"]     := 201      
                    ElseIf (lExistZA4 .AND. QRY->ZA4_STATUS == 'P')
                        //(analisar depois) como validar se pode cancelar/niutilizar?????
                        oResponse["message"]    := cAcao + ' - acao NAO permitida. Já feita a baixa do titulo. '
                        oResponse["type"]       := "error"
                        oResponse["status"]     := 400     
                        lContinua := .F.                               
                    EndIf                         


                    IF (cAcao == 'cancelar' .OR. cAcao == 'inutilizar')
                        oResponse["message"]    := cAcao + ' - concluido com sucesso'
                        oResponse["type"]       := "sucess"
                        oResponse["status"]     := 201                      
                        SalvaCTE(cBody, cChave, cAcao)

                        IF !lContinua
                            //TODO: Enviar email informando que está sendo cancelado ou inutilizado um CTE que no Protheus está com problemas
                            //Assunto, Destinatário, Mensagem, Anexos, Cópia, Cobrança?)
                            cDest := SuperGetMv("P3_MAILCTE",,"thiago.brasil@lauto.com.br")
                            cCopy := SuperGetMv("P3_COPYCTE",,"alexandre.varella@lauto.com.br")
                            cMensagem := "A chave " + cChave + "está com uma inconsistência entre NUCCI x Protheus. Favor analisar! Possível erro: " + oResponse["message"]
                            cFile := ""

                            U_EnviaEmail("CTE LAUTO", cDest, cMensagem, cFile, cCopy, .T.)
                        EndIf

                    Endif
                // ElseIf cAcao == 'cancelar' .OR. cAcao == 'inutilizar'     
                //     //(analisar depois) como validar se pode cancelar/niutilizar?????
                //     oResponse["message"]    := cAcao + ' - concluido com sucesso'
                //     oResponse["type"]       := "sucess"
                //     oResponse["status"]     := 201                      
                //     SalvaCTE(cBody, cChave, cAcao)
                EndIf                
            EndIf
        Else
            oResponse["message"]    := cAcao + ' - erro na solicitacao'
            oResponse["type"]       := "error"
            oResponse["status"]     := 400
        EndIf
    EndIf

Return


//-------------------------------------------------------------------
/*/{Protheus.doc} PopulaZA4
Realiza a leitura do xml do CTE, e popula dos dados do CTE na tabela ZA4
@author  Samuel Dantas
@since   08/03/2020
@version version
/*/
//-------------------------------------------------------------------
Static Function PopulaZA4(cXml, cChave, cAcao)
    Local aAreaSA1      := SA1->(GetArea())
    Local i             := 1
    Local cRet          := ""
    Local cCodCli       := ""
    Local cLojaCli      := ""
    Local cErro      := ""
    Local cAviso      := ""
    Local cTomador      := ""
    Local cCNPJ       := "" 
    Local cCNPJCli       := "" 
    Private cYNumSer    := ""
    Private cYNumNF     := ""
    Private cCHVNFE     := ""
    Private dEmissao  
    Private nVlrFrete   := 0
    Private nVlrCarga   := 0 // Nova variável para valor da carga
    
    Default cXml := ""
    //Transforma o CTE em um objeto
    oCTE := xmlParser(cXml, "_", @cErro, @cAviso)    
    
    If ValType(oCTE) == 'U'
        return 'Erro no objeto XML: ' + cErro
    EndIf

    //Inclusão
    If Alltrim(cAcao) == 'inclusao'
        aVars := {}
        aAdd(aVars, {"Serie CTE"        , "cYNumSer"    , WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_IDE:_SERIE:TEXT","string")        })
        aAdd(aVars, {"Numero CTE"       , "cYNumNF"     , WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_IDE:_NCT:TEXT","string")          })
        aAdd(aVars, {"CNPJ Emitente"    , "cCNPJ"       , WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_EMIT:_CNPJ:TEXT","string")        })
        aAdd(aVars, {"Chave CTE"        , "cCHVNFE"     , WSAdvValue(oCTE,"_CTEPROC:_PROTCTE:_INFPROT:_CHCTE:TEXT","string")        })
        aAdd(aVars, {"Emissao"          , "dEmissao"    , WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_IDE:_DHEMI:TEXT","string")        })
        aAdd(aVars, {"Valor Frete"      , "nVlrFrete"   , WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_VPREST:_VTPREST:TEXT","string")   })

        // Valor da carga é opcional
        nVlrCarga := WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_infCTeNorm:_infCarga:_vCarga:TEXT","string")
        If ValType(nVlrCarga) == "U"
            nVlrCarga := "0"
        EndIf

        //Valida se conseguiu recuperar as variaveis obrigatórias
        For i := 1 to Len(aVars)
            &(aVars[i][2]) := aVars[i][3]
            If ValType(&(aVars[i][2])) == 'U'
                Return "Tag " + aVars[i][1] + " não encontrada no xml do CTE"
            EndIf
        Next

        cCNPJ := oCTE:_CTEPROC:_CTE:_INFCTE:_EMIT:_CNPJ:TEXT
        dEmissao    := StoD(StrTran(Left(dEmissao,10),'-',''))
        
        If !setEmpresa(cCNPJ)
            return 'Cadastro de filial não localizado com o CNPJ ' + cCNPJ
        EndIf

        cTomador := WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_IDE:_TOMA3:_TOMA:TEXT","string")
        If ValType(cTomador) == 'U'
            cTomador := WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_IDE:_TOMA4:_TOMA:TEXT","string")
        EndIf

        If cTomador == 'U'
            Return "Tag Tomador não encontrada no xml do CTE"
        EndIf

        cCNPJCli :=  getTomador(cTomador, oCTE)
        //Valida o CNPJ tomador
        If Empty(cCNPJCli)
            Return 'Erro ao tentar obter o CNPJ('+cCNPJCli+') do tomador.'
        EndIf
        
        SA1->(dbSetOrder(3))
        If SA1->(DbSeek( xFilial("SA1") + PADR(cCNPJCli,TamSx3("A1_CGC")[1]) ))
            cCodCli  := SA1->A1_COD
            cLojaCli := SA1->A1_LOJA
        Else
            Return 'Tomador de CNPJ('+cCNPJCli+') não encontrado.'
        EndIf

        RecLock('ZA4', .T.)
            ZA4->ZA4_FILIAL := xFilial("ZA4")
            ZA4->ZA4_FILCTE := cFilant
            ZA4->ZA4_NUMCTE := PADL(Alltrim(cYNumNF),LEN(ZA4->ZA4_NUMCTE),"0")
            ZA4->ZA4_SERCTE := cYNumSer
            ZA4->ZA4_CLIENT := cCodCli
            ZA4->ZA4_LOJA   := cLojaCli
            ZA4->ZA4_CNPJ   := cCNPJCli
            ZA4->ZA4_BODY   := cXml
            ZA4->ZA4_DATA   := dDataBase
            ZA4->ZA4_HORA   := Time()
            ZA4->ZA4_CHAVE  := cChave
            ZA4->ZA4_ACAO   := cAcao
            ZA4->ZA4_STATUS := 'A'
            ZA4->ZA4_VLRSER := Val(nVlrFrete)
            ZA4->ZA4_VLRCAR := Val(nVlrCarga) // Se não existir a tag, grava zero
        ZA4->(MsUnLock())

    //Popula cancelamento
    ElseIf Alltrim(cAcao) == 'cancelar'
    
        cCnpj := SubStr(Alltrim(cChave), 7, 14)

        If ValType(cCNPJ) == "U"
            return '_procEventoCTe:_eventoCTe:_infEvento:_CNPJ INVALIDA ' 
        EndIf

        If !setEmpresa(cCNPJ)
            return 'Cadastro de filial não localizado com o CNPJ ' + cCNPJ
        EndIf

        cDtcanc := WSAdvValue(oCTE,"_PROCEVENTOCTE:_EVENTOCTE:_infEvento:_dhevento:TEXT","string") 
        If ValType(cDtcanc) == "U"
            cDtcanc := WSAdvValue(oCTE,"_retEventoCTe:_infEvento:_dhRegEvento:TEXT","string") 
            If ValType(cDtcanc) == "U"
                    cDtcanc := WSAdvValue(oCTE,"_PROCEVENTOCTE:_retEventoCTe:_infEvento:_dhRegEvento:TEXT","string") 
            EndIf
        EndIf

        If ValType(cDtcanc) == "U"
            Return "Data do evento não encontrada no XML." + CRLF
        EndIf

        dDtCanc := StoD(StrTran(cDtcanc,"-",""))

        cQuery := " SELECT R_E_C_N_O_ AS RECNO FROM "  + RetSqlName('ZA4') + " ZA4 "
        cQuery += " WHERE ZA4.D_E_L_E_T_ <> '*' AND ZA4_CHAVE = '"+cChave+"' AND ZA4_ACAO = 'inclusao' "
        
        If Select('QRYZA4') > 0
            QRYZA4->(dbclosearea())
        EndIf
        
        TcQuery cQuery New Alias 'QRYZA4'
        
        If QRYZA4->(!Eof())
            ZA4->(DbGoTo(QRYZA4->RECNO))
            cCodCli  := ZA4->ZA4_CLIENT   
            cLojaCli := ZA4->ZA4_LOJA
            cYNumNF  := ZA4->ZA4_NUMCTE   
            cYNumSer := ZA4->ZA4_SERCTE
            cCNPJCli := Posicione("SA1",1,xFilial("SA1") + ZA4->(ZA4_CLIENT+ZA4_LOJA),"A1_CGC")
            QRYZA4->(dbSkip())
        Else
            Return "Não foi encontrada requisição de inclusão para esta chave("+cChave+")."
        EndIf
        
        RecLock('ZA4', .T.)
            ZA4->ZA4_FILIAL := xFilial("ZA4")
            ZA4->ZA4_FILCTE := cFilant
            ZA4->ZA4_NUMCTE := PADL(Alltrim(cYNumNF),LEN(ZA4->ZA4_NUMCTE),"0")
            ZA4->ZA4_SERCTE := cYNumSer
            ZA4->ZA4_CLIENT := cCodCli
            ZA4->ZA4_LOJA   := cLojaCli
            ZA4->ZA4_CNPJ   := cCNPJCli
            ZA4->ZA4_BODY   := cXml
            ZA4->ZA4_DATA   := dDataBase
            ZA4->ZA4_HORA   := Time()
            ZA4->ZA4_CHAVE  := cChave
            ZA4->ZA4_ACAO   := cAcao
            ZA4->ZA4_STATUS := 'A'
        ZA4->(MsUnLock())
    //Popula inutilização
    ElseIf Alltrim(cAcao) = 'inutilizar'
        cStat       := WSAdvValue(oCTE,"_soap_Envelope:_soap_Body:_cteInutilizacaoCTResult:_retInutCTe:_infInut:_cStat:TEXT"   ,"string")
        cCNPJ       := WSAdvValue(oCTE,"_soap_Envelope:_soap_Body:_cteInutilizacaoCTResult:_retInutCTe:_infInut:_CNPJ:TEXT"    ,"string")
        cSerie      := WSAdvValue(oCTE,"_soap_Envelope:_soap_Body:_cteInutilizacaoCTResult:_retInutCTe:_infInut:_SERIE:TEXT"   ,"string")
        cYNumNFIni  := WSAdvValue(oCTE,"_soap_Envelope:_soap_Body:_cteInutilizacaoCTResult:_retInutCTe:_infInut:_nCTIni:TEXT"  ,"string")
        cYNumNFFim  := WSAdvValue(oCTE,"_soap_Envelope:_soap_Body:_cteInutilizacaoCTResult:_retInutCTe:_infInut:_nCTFin:TEXT"  ,"string")
        dDataRec    := WSAdvValue(oCTE,"_soap_Envelope:_soap_Body:_cteInutilizacaoCTResult:_retInutCTe:_infInut:_dhRecbto:TEXT","string")
        
        If ValType(cCNPJ) == "U"
            cStat       := WSAdvValue(oCTE,"_env_Envelope:_env_Body:_cteInutilizacaoCTResult:_retInutCTe:_infInut:_cStat:TEXT"   ,"string")
            cCNPJ       := WSAdvValue(oCTE,"_env_Envelope:_env_Body:_cteInutilizacaoCTResult:_retInutCTe:_infInut:_CNPJ:TEXT"    ,"string")
            cSerie      := WSAdvValue(oCTE,"_env_Envelope:_env_Body:_cteInutilizacaoCTResult:_retInutCTe:_infInut:_SERIE:TEXT"   ,"string")
            cYNumNFIni  := WSAdvValue(oCTE,"_env_Envelope:_env_Body:_cteInutilizacaoCTResult:_retInutCTe:_infInut:_nCTIni:TEXT"  ,"string")
            cYNumNFFim  := WSAdvValue(oCTE,"_env_Envelope:_env_Body:_cteInutilizacaoCTResult:_retInutCTe:_infInut:_nCTFin:TEXT"  ,"string")
            dDataRec    := WSAdvValue(oCTE,"_env_Envelope:_env_Body:_cteInutilizacaoCTResult:_retInutCTe:_infInut:_dhRecbto:TEXT","string")
            If ValType(cCNPJ) == "U"
                return "Tag de CNPJ não encontrada no XML."
            EndIf
            
        EndIf

        If !setEmpresa(cCNPJ)
            return 'Cadastro de filial não localizado com o CNPJ ' + cCNPJ
        EndIf
        cYNumNF := IIF(cYNumNFIni == cYNumNFFim,  PadL(cYNumNFIni, len(SF2->F2_DOC), "0"), cYNumNFIni+"*" )
        dDataRec    := StoD(StrTran(dDataRec,"-",""))

        RecLock('ZA4', .T.)
            ZA4->ZA4_FILIAL := xFilial("ZA4")
            ZA4->ZA4_FILCTE := cFilant
            ZA4->ZA4_NUMCTE := PADL(Alltrim(cYNumNF),LEN(ZA4->ZA4_NUMCTE),"0")
            ZA4->ZA4_SERCTE := cSerie
            ZA4->ZA4_BODY       := cXml
            ZA4->ZA4_DATA       := dDataBase
            ZA4->ZA4_HORA       := Time()
            ZA4->ZA4_CHAVE      := cChave
            ZA4->ZA4_ACAO       := cAcao
            ZA4->ZA4_STATUS     := 'A'
        ZA4->(MsUnLock())
    ElseIf cAcao == 'cartacorrecao'
        ZA4->(DbSetOrder(1))
        If ZA4->(DbSeek(xFilial('ZA4') + cChave + PADR("inclusao",len(ZA4->ZA4_ACAO))))
            cNumNf      := ZA4->ZA4_NUMCTE
            cYNumSer    := ZA4->ZA4_SERCTE
            cCodCli     := ZA4->ZA4_CLIENT
            cLojaCli    := ZA4->ZA4_LOJA
            _cFilial    := ZA4->ZA4_FILCTE
            
            RecLock('ZA4', .T.)                 
                ZA4->ZA4_BODY       := cXml
                ZA4->ZA4_DATA       := dDataBase
                ZA4->ZA4_HORA       := Time()
                ZA4->ZA4_CHAVE      := cChave
                ZA4->ZA4_ACAO       := cAcao
                ZA4->ZA4_STATUS     := 'A'
                ZA4->ZA4_NUMCTE     := cNumNf
                ZA4->ZA4_SERCTE     := cYNumSer
                ZA4->ZA4_CLIENT     := cCodCli
                ZA4->ZA4_LOJA       := cLojaCli
                ZA4->ZA4_FILCTE     := _cFilial
            ZA4->(MsUnLockAll())
        Else
           Return "Não houve inclusão para esta chave("+cChave+") no Protheus." 
        EndIf
    EndIf    
        
    SA1->(RestArea(aAreaSA1))
Return cRet

//-------------------------------------------------------------------
/*/{Protheus.doc} setEmpresa
Seta a empresa de acordo com o CNPJ
@author  Sidney Sales
@since   11/12/19
@version 1.0
/*/
//-------------------------------------------------------------------
Static Function setEmpresa(cCNPJ)
    
    Local lRet := .F.
    
    OpenSM0(cEmpAnt)
    
    SM0->(DbGoTop())

    While SM0->(!Eof())
        If SM0->M0_CGC == cCNPJ
            cFilAnt := SM0->M0_CODFIL
            lRet    := .T.
            Exit
        EndIf
        SM0->(DbSkip())
    EndDo

Return lRet


//-------------------------------------------------------------------
/*/{Protheus.doc} getTomador
Retorna o CGC do tomador
@author  Sidney Sales
@since   11/12/19
@version 1.0
/*/
//-------------------------------------------------------------------
Static Function getTomador(cTomador, oCTE)
    Local cRet      := ''
    Local aTomador  := {}

    aAdd(aTomador, {'0', '_CTEPROC:_CTE:_INFCTE:_REM:_CNPJ:TEXT'    })
    aAdd(aTomador, {'1', '_CTEPROC:_CTE:_INFCTE:_EXPED:_CNPJ:TEXT'  })
    aAdd(aTomador, {'2', '_CTEPROC:_CTE:_INFCTE:_RECEB:_CNPJ:TEXT'  })
    aAdd(aTomador, {'3', '_CTEPROC:_CTE:_INFCTE:_DEST:_CNPJ:TEXT'   })
    aAdd(aTomador, {'4', '_CTEPROC:_CTE:_INFCTE:_IDE:_TOMA4:_CNPJ:TEXT'})

    nPos    := aScan(aTomador,{|x| x[1] == cTomador})
    
    If nPos <> 0
        cRet := WSAdvValue(oCTE,aTomador[nPos][2],"string")
        If ValType(cRet) == 'U'
            aTomador := {}
            aAdd(aTomador, {'0', '_CTEPROC:_CTE:_INFCTE:_REM:_CPF:TEXT'    })
            aAdd(aTomador, {'1', '_CTEPROC:_CTE:_INFCTE:_EXPED:_CPF:TEXT'  })
            aAdd(aTomador, {'2', '_CTEPROC:_CTE:_INFCTE:_RECEB:_CPF:TEXT'  })
            aAdd(aTomador, {'3', '_CTEPROC:_CTE:_INFCTE:_DEST:_CPF:TEXT'   })
            aAdd(aTomador, {'4', '_CTEPROC:_CTE:_INFCTE:_IDE:_TOMA4:_CPF:TEXT'})

            cRet := WSAdvValue(oCTE,aTomador[nPos][2],"string")
            If ValType(cRet) == 'U'
                cRet   := ''
            EndIf
        EndIf
    EndIf

Return cRet

User Function TesCTE()
    Local oResponse	:= JsonObject():new()       
    Local cBody     := ""
    Local cMensagem := ""
    Local aPergs := {}
    Local aRet := {}

    If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01RN0001'
    EndIf

    aAdd(aPergs, {1,"Chave"				    ,Space(Len(ZA4->ZA4_CHAVE)),"@!",'.T.','','.T.',50,.T.})		
    aAdd(aPergs, {1,"Acao"				    ,Space(Len(ZA4->ZA4_ACAO)),"@!",'.T.','','.T.',50,.T.})		
    aAdd(aPergs, {6,"XML"		,Space(254)	 ,'@!'  ,'.T.','', 75, .F., "*.xml |*.xml","C:\"})	 
  
    If ParamBox(aPergs ,"Importação de XML NFS",aRet,,,.T.,50,50,,,.F.,.F.)
        cChave       := Alltrim(aRet[1])
        cAcao       := LOWER(Alltrim(aRet[2]))
        cFile       := Alltrim(aRet[3])
        cBody := MemoRead(cFile,.T.)
    Else
        Return
    EndIf

    cQueryZA4 := " SELECT TOP 1 R_E_C_N_O_ AS RECNOZA4 "
    cQueryZA4 += " FROM "  + RetSqlName('ZA4') + " ZA4"
    cQueryZA4 += " WHERE ZA4.D_E_L_E_T_ <> '*' AND ZA4_CHAVE <> '' "
    
    If Select('QRYZA4') > 0
        QRYZA4->(dbclosearea())
    EndIf
    
    TcQuery cQueryZA4 New Alias 'QRYZA4'
    If QRYZA4->(!Eof())
        ZA4->(DbGoTo(QRYZA4->RECNOZA4))
    
        cChave      := Alltrim(ZA4->ZA4_CHAVE)
        cAcao       := "inclusao" //LOWER(Alltrim(ZA4->ZA4_ACAO))
        // cBody       := ZA4->ZA4_BODY
    EndIf
    
    If !Empty(cChave)
        If cAcao == 'inclusao' .OR. cAcao == 'cartacorrecao'
            cMensagem := PopulaZA4(cBody,cChave,cAcao)
            If !(ExistZA4(cChave,cAcao))
                If Empty(cMensagem)
                    oResponse["message"]    := cAcao + ' - concluido com sucesso'
                    oResponse["type"]       := "sucess"
                    oResponse["status"]     := 201
                Else
                    oResponse["type"]       := "error"
                    oResponse["status"]     := 400
                    oResponse["message"]    := cAcao + ' - acao NAO permitida.'+cMensagem
                EndIf
            Else
                oResponse["type"]       := "error"
                oResponse["status"]     := 400
                oResponse["message"]    := "Chave ("+cChave+") já está cadastrada."
            EndIf
        Else
            If ! Empty(cAcao)
            
                lExistZA4 := ExistZA4(cChave, 'inclusao')                    
                
                If cAcao == 'possocancelar' .OR. cAcao == 'cancelar' .OR. cAcao == 'inutilizar' .OR. cAcao == 'possoinutilizar'           

                    If cAcao == 'possocancelar' .OR. cAcao == 'possoinutilizar' 
                        If !lExistZA4 .AND. cAcao == 'possocancelar'
                            oResponse["message"]    := cAcao + ' - acao NAO permitida. Título não está cadastrado no Protheus. '
                            oResponse["type"]       := "error"
                            oResponse["status"]     := 400 
                        ElseIf (lExistZA4 .AND. QRY->ZA4_STATUS $ 'A,E') .OR. !lExistZA4 
                            oResponse["message"]    := cAcao + ' - acao permitida'
                            oResponse["type"]       := "sucess"
                            oResponse["status"]     := 201                        
                        ElseIf (lExistZA4 .AND. QRY->ZA4_STATUS == 'P') .AND. ! TemBaixa(cChave) //ValidaNota(cChave)
                                //(analisar depois) como validar se pode cancelar/niutilizar?????
                            oResponse["message"]    := cAcao + ' - acao permitida'
                            oResponse["type"]       := "sucess"
                            oResponse["status"]     := 201      
                        ElseIf (lExistZA4 .AND. QRY->ZA4_STATUS == 'P')
                            //(analisar depois) como validar se pode cancelar/niutilizar?????
                            oResponse["message"]    := cAcao + ' - acao NAO permitida. Já feita a baixa do titulo. '
                            oResponse["type"]       := "error"
                            oResponse["status"]     := 400                                    
                        EndIf                        
                    ElseIf cAcao == 'cancelar' .OR. cAcao == 'inutilizar'                      
                        If !(ExistZA4(cChave,cAcao))
                            cMensagem := PopulaZA4(cBody,cChave,cAcao)
                            If Empty(cMensagem)
                                oResponse["message"]    := cAcao + ' - concluido com sucesso'
                                oResponse["type"]       := "sucess"
                                oResponse["status"]     := 201
                            Else
                                oResponse["type"]       := "error"
                                oResponse["status"]     := 400
                                oResponse["message"]    := cAcao + ' - acao NAO permitida.'+cMensagem
                            EndIf
                        Else
                            oResponse["type"]       := "error"
                            oResponse["status"]     := 400
                            oResponse["message"]    := "Chave ("+cChave+") já está cadastrada."
                        EndIf
                    EndIf                
                EndIf
            Else
                oResponse["message"]    := cAcao + ' - erro na solicitacao'
                oResponse["type"]       := "error"
                oResponse["status"]     := 400
            EndIf
        EndIf
    Else
        oResponse["message"]    := 'Chave nao enviada'
        oResponse["type"]       := "error"
        oResponse["status"]     := 400        
    EndIf
    

    MsgAlert(oResponse:toJson())
Return
