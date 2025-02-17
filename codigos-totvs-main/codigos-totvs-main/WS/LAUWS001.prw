#INCLUDE "TOTVS.CH"
#INCLUDE "RESTFUL.CH"
#INCLUDE "TOPCONN.CH"
#INCLUDE "tbiconn.ch"

WSRESTFUL nfse DESCRIPTION 'Post para integracao de NFSE' FORMAT 'application/xml'
    WSMETHOD POST   DESCRIPTION 'Post para gravação de NFSE'  WSSYNTAX '/acao/{}'
END WSRESTFUL

WSMETHOD POST WSSERVICE nfse
    Local oResponse	:= JsonObject():new()     
    Local cBody     := self:getContent()
    Local lRet      := .T.
    Local cMensagem := ""

    ::SetContentType("application/json")   	    

    If len(::aURLParms) < 2
        oResponse["message"]    := 'acao e/ou chave nao enviada'
        oResponse["type"]       := "error"
        oResponse["status"]     := 400        
    Else
        cAcao   := ::aURLParms[1]
        cChave  := ::aURLParms[2]

        If ! Empty(cChave)

            If cAcao == 'inclusao'
                If !ExistZA5(cChave, cAcao)
                    cMensagem := SalvaNFSE(cBody, cChave, cAcao)
                Else
                    cMensagem := "Chave ("+cChave+") já está cadastrada."
                EndIf

                If Empty(cMensagem)
                    oResponse["message"]    := cAcao + ' - concluido com sucesso'
                    oResponse["type"]       := "sucess"
                    oResponse["status"]     := 201
                Else
                    oResponse["message"]    := cAcao + ' - acao NAO permitida.' + cMensagem
                    oResponse["type"]       := "error"
                    oResponse["status"]     := 400
                EndIf
            Else
                If ! Empty(cAcao) .AND. (cAcao == 'inclusao' .OR. cAcao == 'possocancelar' .OR. cAcao == 'cancelar' .OR. cAcao == 'excluir' )
                    If ExistZA5(cChave, cAcao) .AND. cAcao != 'excluir'
                        cMensagem := "Chave ("+cChave+") já está cadastrada. acao = "+cAcao+" "  
                        oResponse["message"]    := cAcao + ' - acao NAO permitida.' + cMensagem
                        oResponse["type"]       := "error"
                        oResponse["status"]     := 400              
                    Else
                        lExistZA5 := ExistZA5(cChave, 'inclusao')                    
                        
                        If cAcao == 'possocancelar' .OR. cAcao == 'cancelar' .OR. cAcao == 'excluir'                     
                            If cAcao == 'possocancelar'
                                If  !lExistZA5
                                    //(analisar depois) como validar se pode cancelar/niutilizar??????
                                    oResponse["message"]    := cAcao + ' - acao NAO permitida. NF não encontrada no Protheus'
                                    oResponse["type"]       := "error"
                                    oResponse["status"]     := 400         
                                ElseIf (lExistZA5 .AND. QRY->ZA5_STATUS $ 'A,E')
                                    oResponse["message"]    := cAcao + ' - acao permitida'
                                    oResponse["type"]       := "sucess"
                                    oResponse["status"]     := 201
                                ElseIf(lExistZA5 .AND. QRY->ZA5_STATUS == 'P') .AND. ! TemBaixa(cChave) //ValidaNota(cChave)
                                    //(analisar depois) como validar se pode cancelar/niutilizar?????
                                    oResponse["message"]    := cAcao + ' - acao permitida'
                                    oResponse["type"]       := "sucess"
                                    oResponse["status"]     := 201                      
                                ElseIf (lExistZA5 .AND. QRY->ZA5_STATUS == 'P')
                                    //(analisar depois) como validar se pode cancelar/niutilizar??????
                                    oResponse["message"]    := cAcao + ' - acao NAO permitida. Título já foi baixado.'
                                    oResponse["type"]       := "error"
                                    oResponse["status"]     := 400                        
                                EndIf                        
                            ElseIf cAcao == 'cancelar'                             
                                If ! TemBaixa(cChave)                                
                                    cMensagem := SalvaNFSE(cBody, cChave, cAcao)
                                    If Empty(cMensagem)
                                        oResponse["message"]    := cAcao + ' - concluido com sucesso'
                                        oResponse["type"]       := "sucess"
                                        oResponse["status"]     := 201
                                    Else
                                        oResponse["message"]    := cAcao + ' - acao NAO permitida.' + cMensagem
                                        oResponse["type"]       := "error"
                                        oResponse["status"]     := 400
                                    EndIf
                                Else
                                    oResponse["message"]    := cAcao + ' - acao NAO permitida. Título já foi baixado.'
                                    oResponse["type"]       := "error"
                                    oResponse["status"]     := 400                        
                                EndIf
                            ElseIf cAcao == 'excluir'                
                                If ! TemBaixa(cChave)                                    
                                    cMensagem := SalvaNFSE(cBody, cChave, cAcao)
                                    If Empty(cMensagem)
                                        oResponse["message"]    := cAcao + ' - concluido com sucesso'
                                        oResponse["type"]       := "sucess"
                                        oResponse["status"]     := 201
                                    Else
                                        oResponse["message"]    := cAcao + ' - acao NAO permitida.' + cMensagem
                                        oResponse["type"]       := "error"
                                        oResponse["status"]     := 400
                                    EndIf
                                Else
                                    oResponse["message"]    := cAcao + ' - acao NAO permitida. Título já foi baixado.'
                                    oResponse["type"]       := "error"
                                    oResponse["status"]     := 400                                                        
                                EndIf
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
    EndIf

    self:SetResponse(EncodeUtf8(oResponse:toJson()))


Return lRet


// //-------------------------------------------------------------------
// /*/{Protheus.doc} SalvaNFSE
// Método utilzao para salvar xml das NFSE para serem processados 
// posteriormente
// @author  Samuel Dantas
// @since   19/11/2019
// @version version
// /*/
// //-------------------------------------------------------------------
// Static Function SalvaNFSE(cXml, cChave, cAcao)
//     Local lRet := .F.
//     If !ExistZA5(cChave, cAcao)
//         RecLock('ZA5', .T.)
//             ZA5->ZA5_BODY       := cXml
//             ZA5->ZA5_DATA       := dDataBase
//             ZA5->ZA5_HORA       := Time()
//             ZA5->ZA5_CHAVE      := cChave
//             ZA5->ZA5_ACAO       := cAcao
//             ZA5->ZA5_STATUS     := 'A'
//         ZA5->(MsUnLock())
//         lRet := .T.
//     EndIf
// Return  lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} ExistZA5
description
@author  Samuel Dantas
@since   19/11/2019
@version version
/*/
//-------------------------------------------------------------------
Static Function ExistZA5(cChave, cAcao)
    Local cQuery := ""
    Local lRet   := .F.

    cQuery := " SELECT * FROM "  + RetSqlName('ZA5') + " ZA5"
    cQuery += " WHERE ZA5.D_E_L_E_T_ <> '*' AND RTRIM(ZA5_ACAO) = '"+cAcao+"' "
    cQuery += " AND RTRIM(ZA5_CHAVE) = '"+Alltrim(cChave)+"' "
    
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
    Local cQryVLD    := ""
    Local lRet      := .F.
    Local cSerNFS := Alltrim(SuperGetMv("MS_SERNFSE",.F.,"NUC"))
    Local cSerie  := u_getSerie('NUC')

    cQryVLD := " SELECT SE1.R_E_C_N_O_ AS RECNOSE1 FROM "  + RetSqlName('ZA5') + " ZA5 "
    cQryVLD += " INNER JOIN "  + RetSqlName('SE1') + " SE1 ON E1_FILIAL = ZA5_FILNFS AND E1_NUM = ZA5_NUMNF  "
    cQryVLD += " AND E1_PREFIXO = ZA5_SERNF AND E1_CLIENTE = ZA5_CLIENT AND E1_LOJA = ZA5_LOJA AND SE1.D_E_L_E_T_ = ZA5.D_E_L_E_T_"
    cQryVLD += " WHERE ZA5.D_E_L_E_T_ <> '*' AND ZA5_CHAVE = '"+cChave+"' AND E1_SALDO > 0 AND E1_PREFIXO = '"+cSerie+"' AND E1_BAIXA = '' "
    
    If Select('QRYVLD') > 0
        QRYVLD->(dbclosearea())
    EndIf
    
    TcQuery cQryVLD New Alias 'QRYVLD'
    
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
    Local cSerNFS := Alltrim(SuperGetMv("MS_SERNFSE",.F.,"NUC"))
    Local cSerie  := u_getSerie('NUC')

    cQry := " SELECT SE1.R_E_C_N_O_ AS RECNOSE1 FROM "  + RetSqlName('ZA5') + " ZA5 "
    cQry += " INNER JOIN "  + RetSqlName('SE1') + " SE1 ON E1_FILIAL = ZA5_FILNFS AND E1_NUM = ZA5_NUMNF  "
    cQry += " AND E1_PREFIXO = ZA5_SERNF AND E1_CLIENTE = ZA5_CLIENT AND E1_LOJA = ZA5_LOJA AND SE1.D_E_L_E_T_ = ZA5.D_E_L_E_T_"
    cQry += " WHERE ZA5.D_E_L_E_T_ <> '*' AND ZA5_CHAVE = '"+cChave+"' AND E1_PREFIXO = '"+cSerie+"' AND (E1_SALDO = 0 OR E1_BAIXA != '' ) "

    If Select('QRYBX') > 0
        QRYBX->(dbclosearea())
    EndIf

    TcQuery cQry New Alias 'QRYBX'

    If QRYBX->(!Eof())
        lRet := .T.
    EndIf

Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} PopulaZA5
Alimenta os campos da tabela ZA5 a partir do xml
@author  Samuel Dantas
@since   08/03/2020
@version version
/*/
//-------------------------------------------------------------------
Static Function SalvaNFSE(cXml,_cChave,_cAcao)
    Local cAviso    := ''
    Local cErro     := ''
    Local cRet      := ""
    Local cCodCli   := ""
    Local cLojaCli   := ""
    Local cYNumNF   := ""
    Local cYNumSer := ""
    Local cCNPJToma := ""
    Local cCNPJEmit := ""
    Local cYNumSer := ""
    Local dEmissao
    Local nValor := 0
    Local i := 1
    Local cSerNFS := Alltrim(SuperGetMv("MS_SERNFSE",.F.,"NUC"))
    Default cAcao   := ""
    Default cXml    := ""
    Default _cChave := ""

    oNFSE := xmlParser(cXml, "_", @cErro, @cAviso)  
    
    SF2->(DbSetOrder(1))
    SF3->(DbSetOrder(5))
    If Alltrim(_cAcao) == 'inclusao'
        If ValType(oNFSE) == 'U'
            Return 'Erro ao processar objeto do XML do NFSE'
        EndIf

        oNFSE := oNFSE:_P_ENVIARLOTERPSENVIO:_P_LOTERPS:_P1_LISTARPS:_P1_RPS:_P1_INFRPS

        If ValType(oNFSE) == 'U'
            Return 'Erro ao processar objeto do XML do NFSE tag NS4 INFNFSE'
        EndIf

        aVars := {}
        aAdd(aVars, {"Numero NFse"      , "cYNumNF"     , WSAdvValue(oNFSE,"_P1_IDENTIFICACAORPS:_P1_NUMERO:TEXT","string")})
        aAdd(aVars, {"Serie"            , "cYNumSer"    , WSAdvValue(oNFSE,"_P1_IDENTIFICACAORPS:_P1_SERIE:TEXT","string")})
        aAdd(aVars, {"Emissao"          , "dEmissao"    , WSAdvValue(oNFSE,"_P1_DATAEMISSAO:TEXT","string")})
        // aAdd(aVars, {"Competenc"        , "dCompet"    , WSAdvValue(oNFSE,"_P1_DATACOMPETENCIA:TEXT","string")})
        aAdd(aVars, {"Valor Servico"    , "nValor"      , WSAdvValue(oNFSE,"_P1_SERVICO:_P1_VALORES:_P1_VALORSERVICOS:TEXT","string") })
        aAdd(aVars, {"Cnpj Emitente", "cCNPJEmit"  , WSAdvValue(oNFSE,"_P1_PRESTADOR:_P1_CNPJ:TEXT","string") })
        
        If ValType(WSAdvValue(oNFSE,"_P1_TOMADOR:_P1_IDENTIFICACAOTOMADOR:_P1_CPFCNPJ:_P1_CNPJ:TEXT","string")) == 'U'
            aAdd(aVars, {"Cnpj Tomador" , "cCNPJToma"  , WSAdvValue(oNFSE,"_P1_TOMADOR:_P1_IDENTIFICACAOTOMADOR:_P1_CPFCNPJ:_P1_CPF:TEXT","string") })
        Else
            aAdd(aVars, {"Cnpj Tomador" , "cCNPJToma"  , WSAdvValue(oNFSE,"_P1_TOMADOR:_P1_IDENTIFICACAOTOMADOR:_P1_CPFCNPJ:_P1_CNPJ:TEXT","string") })
        EndIf

        //Valida se conseguiu recuperar as variaveis
        For i := 1 to Len(aVars)
            &(aVars[i][2]) := aVars[i][3]
            If ValType(&(aVars[i][2])) == 'U'
                Return "Tag " + aVars[i][1] + " não encontrada no xml do CTE"
            EndIf
        Next

        cYNumNF     := PADL(cYNumNF,TAMSX3("F2_DOC")[1],"0")
        dEmissao    := StoD(StrTran(Left(dEmissao,10),'-',''))
        // dCompet     := StoD(StrTran(Left(dCompet,10),'-',''))       

        nValor      := Val(nValor)
        cYNumSer    := cSerNFS

        If ! setEmpresa(cCNPJEmit)
            return 'Cadastro de filial não localizado com o CNPJ ' + cCNPJEmit
        EndIf

        cYNumSer := u_getSerie(cYNumSer)

        SA1->(DbSetOrder(3))
        If SA1->(DbSeek(xFilial("SA1") + PADR(cCNPJToma,LEN(SA1->A1_CGC)) ))
            cCodCli := SA1->A1_COD
            cLojaCli := SA1->A1_LOJA
        Else    
            Return "Tomador de CNPJ "+cCNPJToma+" não está cadastrado no PROTHEUS."
        EndIf
        
        If ExistNum(cFilAnt, cYNumNF, cYNumSer, _cAcao)
            Return "A Acao "+cAcao+" já está cadastrada para a numeração "+cYNumNF+" para a Filial de CNPJ "+cCNPJEmit+", na chave '"+ZA5->ZA5_CHAVE+"'."  
        EndIf

        // If SF3->(DbSeek(xFilial("SF3") + cSerNFS +cYNumNF+SA1->(A1_COD+A1_LOJA)))
        //     Return "NFS "+cYNumNF+" ja está "+IIF(SF3->F3_CODRSEF == '101',"cancelada", "cadastrada")+" para a Filial de CNPJ "+cCNPJEmit+". "    
        // EndIf 

        //Transforma o NFSE em um objeto
        RecLock('ZA5', .T.)
            ZA5->ZA5_FILNFS := cFilAnt
            ZA5->ZA5_NUMNF  := cYNumNF
            ZA5->ZA5_SERNF  := cYNumSer
            ZA5->ZA5_CLIENT := cCodCli
            ZA5->ZA5_LOJA   := cLojaCli
            ZA5->ZA5_STATUS := 'A'
            ZA5->ZA5_BODY   := cXml
            ZA5->ZA5_DATA   := dDataBase
            ZA5->ZA5_HORA   := Time()
            ZA5->ZA5_CHAVE  := _cChave
            ZA5->ZA5_ACAO   := _cAcao
        ZA5->(MsUnLock())

    ElseIf Alltrim(_cAcao) == 'cancelar' .OR. Alltrim(_cAcao) == 'excluir'
        aVars := {}
        aAdd(aVars, {"Numero NFse"      , "cYNumNF"     , WSAdvValue(oNFSE,"_P_CancelarNfseEnvio:_PEDIDO:_P1_InfPedidoCancelamento:_P1_IdentificacaoNfse:_P1_NUMERO:TEXT","string")})
        aAdd(aVars, {"CNPJ"             , "cYCnpj"     , WSAdvValue(oNFSE,"_P_CancelarNfseEnvio:_PEDIDO:_P1_InfPedidoCancelamento:_P1_IdentificacaoNfse:_P1_CNPJ:TEXT","string")})
        aAdd(aVars, {"dData"             , "dData"     , WSAdvValue(oNFSE,"_P_CancelarNfseEnvio:_PEDIDO:_P1_InfPedidoCancelamento:_P1_IdentificacaoNfse:_P1_DataHora:TEXT","string")})
        
        //Valida se conseguiu recuperar as variaveis
        For i := 1 to Len(aVars)
            &(aVars[i][2]) := aVars[i][3]
            If ValType(&(aVars[i][2])) == 'U'
                Return "Tag " + aVars[i][1] + " não encontrada no xml do NFSE"
            EndIf
        Next

        lVldCNPJ := setEmpresa(cYCnpj)
        If !lVldCNPJ
            return 'Cadastro de filial não localizado com o CNPJ ' + cYCnpj
        EndIf
        
        cYNumSer := u_getSerie(cSerNFS)

        cQry := " SELECT ZA5_CLIENT, ZA5_LOJA FROM "  + RetSqlName('ZA5') + " ZA5"
        cQry += " WHERE ZA5.D_E_L_E_T_ <> '*' AND ZA5_ACAO = 'inclusao' AND ZA5_CHAVE = '"+_cChave+"' "
        
        If Select('QRYAUX') > 0
            QRYAUX->(dbclosearea())
        EndIf
        
        TcQuery cQry New Alias 'QRYAUX'
        
        If QRYAUX->(!Eof())
            cCodCli := QRYAUX->ZA5_CLIENT
            cLojaCli := QRYAUX->ZA5_LOJA
            QRYAUX->(dbSkip())
        Else
            Return "Chave "+_cChave+" com ação inclusão não foi encontrada no Protheus. Não foi possivel "+_cAcao+"."
        EndIf

        cYNumNF     := PADL(cYNumNF,TAMSX3("F2_DOC")[1],"0")
        dData := StoD(StrTran(Left(dData,10),'-',''))
        
        BEGIN TRANSACTION

            If ExistNum(cFilAnt, cYNumNF, cYNumSer, _cAcao)
                If Alltrim(_cAcao) != "excluir"
                    DisarmTransaction()
                    Return "A Acao "+cAcao+" já está cadastrada para a numeração "+cYNumNF+" para a Filial de CNPJ "+cYCnpj+", na chave '"+ZA5->ZA5_CHAVE+"'."  
                Else
                    If ZA5->(DbSeek(xFilial("ZA5") + PADR(_cChave, len(ZA5->ZA5_CHAVE)) + "excluir" ))
                        RecLock('ZA5', .F.)
                            ZA5->(DbDelete())
                        ZA5->(MsUnLock())
                    EndIf
                EndIf
            EndIf
            
        /*
        If _cAcao == "excluir"
            If ExistNum(cFilAnt, cYNumNF, cSerNFS, "cancelar")
                Return "Não é possível excluir pois existe um requisição de cancelamento para esta nota na chave "+ZA5->ZA5_CHAVE+"."  
            EndIf
        EndIf 
        */
            RecLock('ZA5', .T.)
                ZA5->ZA5_FILNFS := cFilAnt
                ZA5->ZA5_NUMNF  := cYNumNF
                ZA5->ZA5_SERNF  := cYNumSer
                ZA5->ZA5_CLIENT := cCodCli
                ZA5->ZA5_LOJA   := cLojaCli
                ZA5->ZA5_STATUS := 'A'
                ZA5->ZA5_BODY   := cXml
                ZA5->ZA5_DATA   := dDataBase
                ZA5->ZA5_HORA   := Time()
                ZA5->ZA5_CHAVE  := _cChave
                ZA5->ZA5_ACAO   := _cAcao
            ZA5->(MsUnLock())

            If Alltrim(_cAcao) == 'excluir'
                If ZA5->(DbSeek(xFilial("ZA5") + PADR(_cChave, len(ZA5->ZA5_CHAVE)) + "inclusao" ))
                    RecLock('ZA5', .F.)
                        ZA5->(DbDelete())
                    ZA5->(MsUnLock())
                    If ZA5->(DbSeek(xFilial("ZA5") + PADR(_cChave, len(ZA5->ZA5_CHAVE)) + "cancelar" ))
                        RecLock('ZA5', .F.)
                            ZA5->(DbDelete())
                        ZA5->(MsUnLock())
                    EndIf
                Else
                    DisarmTransaction()
                    Return "Chave "+_cChave+" com ação inclusão não foi encontrada no Protheus. Não foi possivel fazer a exclusão."     
                EndIf
            EndIf
        END TRANSACTION 
    EndIf

Return cRet

//-------------------------------------------------------------------
/*/{Protheus.doc} TesNFS
Método para realizar testes da integração
@author  Samuel Dantas
@since   31/03/2020
@version version
/*/
//-------------------------------------------------------------------
User Function TesNFS()
    Local oResponse	:= JsonObject():new()       
    Local cBody     := ""
    Local lRet      := .T.
    Local cMensagem := ""
    Local aPergs := {}
    Local aRet := {}

    If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01RN0001'
    EndIf
    aAdd(aPergs, {1,"Chave"				    ,Space(Len(ZA5->ZA5_CHAVE)),"@!",'.T.','','.T.',50,.T.})		
    aAdd(aPergs, {1,"Acao"				    ,Space(Len(ZA5->ZA5_ACAO)),"@!",'.T.','','.T.',50,.T.})		
    aAdd(aPergs, {6,"XML"		,Space(254)	 ,'@!'  ,'.T.','', 75, .F., "*.XML |*.XML","C:\"})	 
    
    If ParamBox(aPergs ,"Importação de XML NFS",aRet,,,.T.,50,50,,,.F.,.F.)
        cChave      := Alltrim(aRet[1])
        cAcao       := LOWER(Alltrim(aRet[2]))
        cFile       := Alltrim(aRet[3])
        cBody       := MemoRead(cFile,.T.)
    EndIf

    If ! Empty(cChave)
        If cAcao == 'inclusao'
            If !ExistZA5(cChave, cAcao)
                cMensagem := SalvaNFSE(cBody, cChave, cAcao)
            Else
                cMensagem := "Chave ("+cChave+") já está cadastrada."
            EndIf

            If Empty(cMensagem)
                oResponse["message"]    := cAcao + ' - concluido com sucesso'
                oResponse["type"]       := "sucess"
                oResponse["status"]     := 201
            Else
                oResponse["message"]    := cAcao + ' - acao NAO permitida.' + cMensagem
                oResponse["type"]       := "error"
                oResponse["status"]     := 400
            EndIf
        Else
            If ! Empty(cAcao) .AND. (cAcao == 'inclusao' .OR. cAcao == 'possocancelar' .OR. cAcao == 'cancelar' .OR. cAcao == 'excluir' )
                If ExistZA5(cChave, cAcao)    
                    cMensagem := "Chave ("+cChave+") já está cadastrada. acao = "+cAcao+" "  
                    oResponse["message"]    := cAcao + ' - acao NAO permitida.' + cMensagem
                    oResponse["type"]       := "error"
                    oResponse["status"]     := 400              
                Else
                    lExistZA5 := ExistZA5(cChave, 'inclusao')                    
                    
                    If cAcao == 'possocancelar' .OR. cAcao == 'cancelar' .OR. cAcao == 'excluir'                     
                        If cAcao == 'possocancelar'
                            If  !lExistZA5
                                //(analisar depois) como validar se pode cancelar/niutilizar??????
                                oResponse["message"]    := cAcao + ' - acao NAO permitida. NF não encontrada no Protheus'
                                oResponse["type"]       := "error"
                                oResponse["status"]     := 400         
                            ElseIf (lExistZA5 .AND. QRY->ZA5_STATUS $ 'A,E')
                                oResponse["message"]    := cAcao + ' - acao permitida'
                                oResponse["type"]       := "sucess"
                                oResponse["status"]     := 201
                            ElseIf(lExistZA5 .AND. QRY->ZA5_STATUS == 'P') .AND. ! TemBaixa(cChave) //ValidaNota(cChave)
                                //(analisar depois) como validar se pode cancelar/niutilizar?????
                                oResponse["message"]    := cAcao + ' - acao permitida'
                                oResponse["type"]       := "sucess"
                                oResponse["status"]     := 201                      
                            ElseIf (lExistZA5 .AND. QRY->ZA5_STATUS == 'P')
                                //(analisar depois) como validar se pode cancelar/niutilizar??????
                                oResponse["message"]    := cAcao + ' - acao NAO permitida. Título já foi baixado.'
                                oResponse["type"]       := "error"
                                oResponse["status"]     := 400                        
                            EndIf                        
                        ElseIf cAcao == 'cancelar'                            
                            cMensagem := SalvaNFSE(cBody, cChave, cAcao)
                            If Empty(cMensagem)
                                oResponse["message"]    := cAcao + ' - concluido com sucesso'
                                oResponse["type"]       := "sucess"
                                oResponse["status"]     := 201
                            Else
                                oResponse["message"]    := cAcao + ' - acao NAO permitida.' + cMensagem
                                oResponse["type"]       := "error"
                                oResponse["status"]     := 400
                            EndIf
                        ElseIf cAcao == 'excluir'                
                            cMensagem := SalvaNFSE(cBody, cChave, cAcao)
                            If Empty(cMensagem)
                                oResponse["message"]    := cAcao + ' - concluido com sucesso'
                                oResponse["type"]       := "sucess"
                                oResponse["status"]     := 201
                            Else
                                oResponse["message"]    := cAcao + ' - acao NAO permitida.' + cMensagem
                                oResponse["type"]       := "error"
                                oResponse["status"]     := 400
                            EndIf
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

//-------------------------------------------------------------------
/*/{Protheus.doc} ExistNum
Verifica se já existe requisição para essas numeração.
@author  Samuel Dantas
@since   31/03/2020
@version version
/*/
//-------------------------------------------------------------------
Static Function ExistNum(_cFilial, _cNum, _cSerie, _cAcao )
    Local cQryZA5 := ""
    Local lRet := .F.

    cQryZA5 := " SELECT R_E_C_N_O_ AS RECNO FROM "  + RetSqlName('ZA5') + " ZA5"
    cQryZA5 += " WHERE ZA5.D_E_L_E_T_ <> '*' AND ZA5_NUMNF = '"+_cNum+"' AND ZA5_FILNFS = '"+_cFilial+"' AND ZA5_SERNF = '"+_cSerie+"' AND ZA5_ACAO = '"+_cAcao+"' "
    
    If Select('QZA5') > 0
        QZA5->(dbclosearea())
    EndIf
    
    TcQuery cQryZA5 New Alias 'QZA5'
    
    If QZA5->(!Eof())
        ZA5->(DbGoTo(QZA5->RECNO))
        lRet := .T.
        QZA5->(dbSkip())
    EndIf
    
Return lRet

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
