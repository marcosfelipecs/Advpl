#INCLUDE 'Rwmake.ch'
#INCLUDE 'Protheus.ch'
#INCLUDE 'Tbiconn.ch'
#INCLUDE 'Topconn.ch'

//-------------------------------------------------------------------
/*/{Protheus.doc} LAUF0005
Funcao que processara os xmls e nfse
@author  Sidney Sales
@since   05/12/19
@version 1.0
/*/
//-------------------------------------------------------------------
User Function LAUF0005(nRecNoZA4)
    Local cRet
    Local cEmailSup := ""
    Private cErrorBlock := ''
    Private cYNumSer    := ''
    Private cYNumNF     := ''
    
    Default nRecNoZA4 := 0

    If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01RN0001'
    EndIf
    
    cEmailSup := Alltrim(SuperGetMv("MS_TIMAIL",.T.,"diego.cunha@lauto.com.br"))

    cQuery := " SELECT "    
    cQuery += " R_E_C_N_O_ AS RECZA4 FROM " + RetSqlTab('ZA4')
    cQuery += " WHERE D_E_L_E_T_ = '' "
    
    If nRecNoZA4 == 0
        cQuery += " AND ZA4_STATUS = 'A' "
    Else
        cQuery += " AND R_E_C_N_O_ = " + cValToChar(nRecNoZA4)
    EndIf

    If Select('QRYZA4') > 0
        QRYZA4->(DbCloseArea())
    EndIf

    TcQuery cQuery New Alias 'QRYZA4'
    
    //Percorre todos os registros e faz o processamento
    While QRYZA4->(!Eof())
        cRet := ProcessaCTE(QRYZA4->RECZA4, 'P')        
        
        RecLock('ZA4', .F.)
            ZA4->ZA4_STATUS := Iif(cRet != 'NF Gerada', 'E','P')
            ZA4->ZA4_ERRO   := Iif(cRet != 'NF Gerada', cRet,'')
            ZA4->ZA4_ERRFIS := Iif(!Empty(cErrorBlock), cErrorBlock, '')
            ZA4->ZA4_STAFIS := Iif(!Empty(cErrorBlock), 'E', 'P')
            // ZA4->ZA4_SERCTE := cYNumSer
            // ZA4->ZA4_NUMCTE := cYNumNF
        ZA4->(MsUnLock())
        If ZA4->ZA4_STATUS == 'E' .AND. !Empty(cEmailSup)
            U_EnviaEmail(" Processamento CTE - Erro", cEmailSup, "Foi encontrada a seguinte msg de erro ( ZA4_CHAVE = "+ZA4->ZA4_CHAVE+") :"+CRLF+Alltrim(ZA4->ZA4_ERRO), "", "", .F.)
        EndIf
        QRYZA4->(DbSkip())
    EndDo
 
Return

//-------------------------------------------------------------------
/*/{Protheus.doc} LAUF0005B
Usada no reprocessamento dos dados fiscais
@author  Samuel Dantas
@since   19/12/2019
@version version
/*/
//-------------------------------------------------------------------
User Function LAUF005B(cTipo)
    Local cTitulo := ''
    Default cTipo := ''

    If cTipo == 'P'
        cTitulo := "Processando nota"
    ElseIf cTipo == 'R'
        cTitulo := "Reprocessando dados fiscais"
    ElseIf cTipo == 'E'
        cTitulo := "Excluindo dados fiscais"
    Else
        cTipo := ""
    EndIf

    If !Empty(cTipo)
        MSAguarde( { || ProcessaCTE(ZA4->(Recno()), cTipo ) }, cTitulo ,"Processando...",.F.)
    EndIf
    
Return

//-------------------------------------------------------------------
/*/{Protheus.doc} ProcessaCTE
Funcao que faz o processamento dos CTEs incluindo os pedidos de venda
@author  Sidney Sales
@since   11/12/19
@version 1.0
/*/
//-------------------------------------------------------------------
Static Function ProcessaCTE(nRecZA4, cTipo )
    Local i
    Private cAviso    := ''
    Private cErro     := ''
    Private cArqErro	:= "erroauto.txt"    
    Private cCond     := SuperGetMv("MS_CONDTRA",.F.,"001")
    Private cCNPJ, cCHVNFE, cTes, oCTE
    Private i, j, nZ
    Private bBlock 	    := ErrorBlock()
    Private cMensagem 	:= ""
    Private cTpFrete 	:= ""
    Private aNFS 	    := {}
    
    //Posiona o registro para processamento
    ZA4->(DbGoTo(nRecZA4))
    cXml := ZA4->ZA4_BODY
    cXml := StrTran(cXml,"env:","soap:")
    
    //Valida o cadastro do tipo CTE no cadastro de tipos de titulos do financeiro
    If !  SX5->(DbSeek(xFilial('SX5') +  '05' + 'CTE'))
        return "Tipo CTE não localizado no cadastro de tipo de titulos"
    EndIf
    
    //Transforma o CTE em um objeto
    oCTE := xmlParser(cXml, "_", @cErro, @cAviso)    
            
    //Valida o objeto    
    If ValType(oCTE) == 'U'
        return 'Erro no objeto XML: ' + cErro
    EndIf

    If cTipo == 'P' .AND. Alltrim(ZA4->ZA4_ACAO) == 'cancelar' // Cancela nota
        ExcluiCTE()
        return cMensagem
    ElseIf cTipo == 'P' .AND. Alltrim(ZA4->ZA4_ACAO) == 'inutilizar' // Cancela nota
        InutCTE()
        return cMensagem
    EndIf

    aVars := {}
    aAdd(aVars, {"Serie CTE"        , "cYNumSer"    , WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_IDE:_SERIE:TEXT","string")})
    aAdd(aVars, {"Numero CTE"       , "cYNumNF"     , WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_IDE:_NCT:TEXT","string")  })
    aAdd(aVars, {"CNPJ Emitente"    , "cCNPJ"       , WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_EMIT:_CNPJ:TEXT","string")})
    aAdd(aVars, {"Chave CTE"        , "cCHVNFE"     , WSAdvValue(oCTE,"_CTEPROC:_PROTCTE:_INFPROT:_CHCTE:TEXT","string")})
    aAdd(aVars, {"CFOP"             , "cCFOP"       , WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_IDE:_CFOP:TEXT","string")})
    aAdd(aVars, {"Emissao"          , "dEmissao"    , WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_IDE:_DHEMI:TEXT","string")})
    aAdd(aVars, {"Valor Frete"      , "nVlrFrete"   , WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_VPREST:_VTPREST:TEXT","string")})
    aAdd(aVars, {"UF Inicial"       , "cUFORIG"     , WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_IDE:_UFINI:TEXT","string")})
    aAdd(aVars, {"UF Destino"       , "cUFDEST"     , WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_IDE:_UFFIM:TEXT","string")})
    
    If ValType(WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_REM:_CNPJ:TEXT","string")) != "U"
        aAdd(aVars, {"Remetente"     , "cCNPJRem"   , WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_REM:_CNPJ:TEXT","string")})
    ElseIf ValType(WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_REM:_CPF:TEXT","string")) != "U"
        aAdd(aVars, {"Remetente"     , "cCNPJRem"   , WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_REM:_CPF:TEXT","string")})
    Else
        Return "Tag _INFCTE:_REM:_CNPJ não encontrada no xml do CTE"
    EndIf
    
    aAdd(aVars, {"Nome Remetente"   , "cNomRem"     , WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_REM:_xNome:TEXT","string")})
    
    If ValType(WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_DEST:_CNPJ:TEXT","string")) != "U"
        aAdd(aVars, {"Destinatario"     , "cCnpjDest"   , WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_DEST:_CNPJ:TEXT","string")})
    ElseIf ValType(WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_DEST:_CPF:TEXT","string")) != "U"
        aAdd(aVars, {"Destinatario"     , "cCnpjDest"   , WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_DEST:_CPF:TEXT","string")})
    Else
        Return "Tag _INFCTE:_DEST:_CNPJ não encontrada no xml do CTE"
    EndIf

    aAdd(aVars, {"Nome Dest"        , "cNomDest"    , WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_DEST:_xNome:TEXT","string")})
    

    If Type('oProcess') != 'U'        	
    	oProcess:SetRegua2(4)	
        oProcess:IncRegua2("Validando XML...")		    
	EndIf

    //Valida se conseguiu recuperar as variaveis
    For i := 1 to Len(aVars)
        &(aVars[i][2]) := aVars[i][3]
        If ValType(&(aVars[i][2])) == 'U'
            Return "Tag " + aVars[i][1] + " não encontrada no xml do CTE"
        EndIf
    Next

    cTomador := WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_IDE:_TOMA3:_TOMA:TEXT","string")
    If ValType(cTomador) == 'U'
        cTomador := WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_IDE:_TOMA4:_TOMA:TEXT","string")
    EndIf

    If cTomador == 'U'
        Return "Tag Tomador não encontrada no xml do CTE"
    EndIf

    //Numero do CTE, preenche zeros a esquerda
    cYNumNF     := PADL(cYNumNF,TAMSX3("F2_DOC")[1],"0")
    
    If Type('oProcess') != 'U'        	
        oProcess:IncRegua2("Processando CTE " +  cYNumNF + '/' + cYNumSer)
	EndIf

    //Alteracao da data base
    dEmissao    := StoD(StrTran(Left(dEmissao,10),'-',''))
    dDatabase   := dEmissao

    //Pega o valor total do frete
    nVlrFrete   := Val(nVlrFrete)
    
    //dados dos vendedores, passa por parametro o array com os dados dos vendedores
    aXMLVend := WSAdvValue(oCTE,"_CTE:_INFCTE:_COMPL:_OBSCONT","string")
    aVends   := dadosVend(aXMLVend, nVlrFrete)

    //Seta a filial que corresponde ao CNPJ
    If ! setEmpresa(cCNPJ)
        return 'Cadastro de filial não localizado com o CNPJ ' + cCNPJ
    EndIf

    //Posiciona o cliente    
    cCNPJCli := getTomador(cTomador, oCTE)
    
    //Valida o CNPJ tomador
    If Empty(cCNPJCli)
        Return 'Erro ao tentar obter o CNPJ do tomador. 
    EndIf

    //Posiciona o cliente pelo CNPJ do tomador
    SA1->(dbSetOrder(3))
    If ! SA1->(DbSeek( xFilial("SA1") + PADR(cCNPJCli,TamSx3("A1_CGC")[1]) ))
        RecLock('ZA4', .F.)
            ZA4->ZA4_CNPJ := cCNPJCli
        ZA4->(MsUnLock())
        Return 'Cliente(tomador) não cadastrado no Protheus ' + cCNPJCli
    Else
        cCODCli := SA1->A1_COD
        cLojCli := SA1->A1_LOJA
        cNomCli := SA1->A1_NOME
    EndIf    
    
    RecLock('ZA4', .F.)
        ZA4->ZA4_FILCTE := cFilAnt
        ZA4->ZA4_CLIENT := SA1->A1_COD
        ZA4->ZA4_LOJA   := SA1->A1_LOJA
    ZA4->(MsUnLock())
    //Verifica o remetente, se nao tiver, GRAVA NA SA1
    SA1->(dbSetOrder(3))
    If ! SA1->(DbSeek( xFilial("SA1") + PADR(cCNPJRem,TamSx3("A1_CGC")[1]) ))
        cAux := GravaSA1(cCNPJRem,cNomRem)
        If Empty(cAux)
            cCODRem := SA1->A1_COD
            cLojRem := SA1->A1_LOJA
            cNomRem := SA1->A1_NOME
        Else
            Return "Remetente não cadastrado na (SA1)"+CRLF+cAux
        EndIf
    Else
        cCODRem := SA1->A1_COD
        cLojRem := SA1->A1_LOJA
        cNomRem := SA1->A1_NOME
    EndIf   

    //Verifica qual TES devera ser utilizada de acordo com o CFOP
    cCst := WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_IMP:_ICMS:_ICMS90:_CST:TEXT","string")//NP3-26/10/2022 - Ajuste de TES que não calcula ICM
    If ValType(cCst) == "U"
        cCst := ''
    EndIf
    cTes := getTes(cCFOP,cCst)
    cCst := WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_IMP:_ICMS:_ICMS45:_CST:TEXT","string")
    If ValType(cCst) != "U"
        cTes := AllTrim(SuperGetMv("MS_TES45", .F., "53J"))
    EndIf
    
    If Empty(cTes)
        Return 'TES não localizada para o CFOP ' + cCFOP
    EndIf

    //Verifica o destinatario, se nao tiver, nao usa        
    SA1->(dbSetOrder(3))
    If ! SA1->(DbSeek(xFilial('SA1') + cCnpjDest))
        cAux := GravaSA1(cCnpjDest,cNomDest)
        If Empty(cAux)
            cCODDes := SA1->A1_COD
            cLojDes := SA1->A1_LOJA
            cNomDes := SA1->A1_NOME
        Else
            Return "Destinatário não cadastrado na (SA1)"+CRLF+cAux
        EndIf
    Else
        cCODDes := SA1->A1_COD
        cLojDes := SA1->A1_LOJA
        cNomDes := SA1->A1_NOME
    EndIf
    
    aNFS := getNfs(oCTE)
    
    If (cCODRem+ cLojRem) == (cCODCli + cLojCli)
        cTpFrete := "C"
    Else
        cTpFrete := "F"
    EndIf

    cQuery := " SELECT R_E_C_N_O_ AS RECNO FROM "  + RetSqlName('SF2') + " SF2"
	cQuery += " WHERE SF2.D_E_L_E_T_ <> '*' AND F2_CHVNFE = '"+ZA4->ZA4_CHAVE+"' "
	
	If Select('QRY') > 0
		QRY->(dbclosearea())
	EndIf
	
	TcQuery cQuery New Alias 'QRY'

    If cTipo == 'P' .AND. Alltrim(ZA4->ZA4_ACAO) == 'inclusao' // Inclui|Processa a nota
        ProcCte(aVends)
    ElseIf cTipo == 'P' .AND. Alltrim(ZA4->ZA4_ACAO) == 'cancelar' // Cancela nota
        ExcluiCTE()
    ElseIf cTipo == 'P' .AND. Alltrim(ZA4->ZA4_ACAO) == 'inutilizar'// Inutiliza nota
        InutCTE()
    ElseIf cTipo == 'R'
        If QRY->(!EoF())
            ProcFis(.T., QRY->RECNO) //Reprocessa dados fiscais
        EndIf
    ElseIf cTipo == 'E'
        If QRY->(!EoF())
            ProcFis(.F., QRY->RECNO) // Exclui dados fiscais
        EndIf
    EndIf
    
Return cMensagem


//-------------------------------------------------------------------
/*/{Protheus.doc} ProcFis
Método usado para exclusão e reprocessamento de dados fiscais da nota.
@author  Samuel Dantas 
@since   16/12/2019
@version version
/*/
//-------------------------------------------------------------------
Static Function ProcFis(lReprocessa,nRecnoSF2)
    Local aRet := {}
    Local i, j, nz
    Private nOrdem := 0
    Private cChave := ""
    Private aTabs := {}
    Private cAlias
    Private cCAMPO := ""
    Default nRecnoSF2 := 0
    Default lReprocessa := .F.


    SF2->(DbGoTo(nRecnoSF2))
    If SF2->(!EoF())
        //Posiona o registro para processamento
        cXml := ZA4->ZA4_BODY

        //Transforma o CTE em um objeto
        oCTE := xmlParser(cXml, "_", @cErro, @cAviso)    
        cCNPJ := oCTE:_CTEPROC:_CTE:_INFCTE:_EMIT:_CNPJ:TEXT
        
        //NP3-26/10/2022 - validar esse ponto
        // If ! StaticCall(LAUF0005, setEmpresa, cCNPJ)
        //     return 'Cadastro de filial não localizado com o CNPJ ' + cCNPJ
        // EndIf

        SA1->(DbSetOrder(1))
        If SA1->(DbSeek(xFilial("SA1") + SF2->F2_CLIENTE + SF2->F2_LOJA))
        EndIf
        //funcao que valida os erros em tempo de execucao
        ErrorBlock( {|e| cErrorBlock := e:Description + e:ErrorStack })
        BEGIN SEQUENCE
            aTabs := {}
                
            aAdd(aTabs, {'SF2', 1, SF2->(F2_FILIAL + F2_DOC + F2_SERIE + F2_CLIENTE + F2_LOJA), "SF2->(F2_FILIAL + F2_DOC + F2_SERIE + F2_CLIENTE + F2_LOJA)" })
            aAdd(aTabs, {'SD2', 3, SF2->(F2_FILIAL + F2_DOC + F2_SERIE + F2_CLIENTE + F2_LOJA), "SD2->(D2_FILIAL + D2_DOC + D2_SERIE + D2_CLIENTE + D2_LOJA)" })
            aAdd(aTabs, {'SF3', 4, SF2->(F2_FILIAL + F2_CLIENTE + F2_LOJA + F2_DOC + F2_SERIE), "SF3->(F3_FILIAL + F3_CLIEFOR + F3_LOJA + F3_NFISCAL + F3_SERIE)" }) 
            aAdd(aTabs, {'SFT', 1, SF2->F2_FILIAL + 'S' + SF2->(F2_SERIE + F2_DOC + F2_CLIENTE + F2_LOJA), "SFT->(FT_FILIAL + FT_TIPOMOV + FT_SERIE + FT_NFISCAL + FT_CLIEFOR + FT_LOJA)" })
            
            //Percorre o array de tabelas, preenchendo os campos extras
            For i := 1  to Len(aTabs)
                cAlias  := aTabs[i][1]
                nOrdem  := aTabs[i][2]
                cChave  := aTabs[i][3]
                
                //Retorna os dados que devem ser preenchidos na tabela
                aRet    := getArray(oCTE, cAlias, nRecnoSF2)          
                (cAlias)->(DbSetOrder(nOrdem))
                (cAlias)->(DbSeek(cChave))            
                
                //Perccorre os registro, na pratica, sempre sera um, porem, ja coloquei para evitar. 
                While (cAlias)->(!Eof()) .AND. cChave == &(aTabs[i][4])
                        //Altera os campos
                        RecLock(cAlias, .F.)
                            For j := 1 to Len(aRet)
                                cCampo := aRet[j][1]
                                xValor := aRet[j][2]
                                If Alltrim(cCampo) != "F2_CHVNFE"
                                    xValor := aRet[j][2]
                                    (cAlias)->&(cCampo) := Iif(!lReprocessa, GetVazio(cAlias,cCampo), xValor ) 
                                ElseIf lReprocessa 
                                    (cAlias)->&(cCampo) := xValor
                                EndIf
                            Next
                        (cAlias)->(MsUnLock())
                    
                    (cAlias)->(DbSkip())
                EndDo
            Next
            SF2->(DbGoTo(nRecnoSF2))
            If lReprocessa .AND. Alltrim(ZA4->ZA4_ACAO) == 'inclusao'
                //Cria o DT6, tabela do TMS, usada apenas para funcionar o registro fiscal
                aDT6 := getArray(oCTE, 'DT6', nRecnoSF2)
                DT6->(DBOrderNickname("DT6CHVCTE"))
                If !DT6->(DbSeek(xFilial("DT6") + ZA4->ZA4_CHAVE))
                    RecLock('DT6', .T.)            
                        For j := 1 to Len(aDT6)
                            cCampo := aDT6[j][1]
                            xValor := aDT6[j][2]
                            DT6->&(cCampo) := xValor
                        Next
                    DT6->(MsUnLock())
                EndIf

                RecLock('ZA4', .F.)
                    ZA4->ZA4_STATUS := 'P'
                    ZA4->ZA4_STAFIS := 'P'
                ZA4->(MsUnLock())

                For nZ := 1 To len(aNFS)
                    If !ZA8->(DbSeek(xFilial("ZA8") + ZA4->ZA4_CHAVE + PADR(aNFS[nZ],Len(ZA4->ZA4_CHAVE)) ))
                        RecLock('ZA8', .T.)
                            ZA8->ZA8_CHVCTE := ZA4->ZA4_CHAVE
                            ZA8->ZA8_CHVNFE := aNFS[nZ]
                        ZA8->(MsUnLock())
                    EndIf
                Next

            Else
                //Exclui DT6
                DT6->(DBOrderNickname("DT6CHVCTE"))
                If DT6->(DbSeek(xFilial("DT6") + SF2->F2_CHVNFE))
                    RecLock('DT6', .F.)
                        DT6->(DbDelete())
                    DT6->(MsUnLock())
                    RecLock('ZA4', .F.)
                        ZA4->ZA4_STATUS := 'P'
                        ZA4->ZA4_STAFIS := 'A'
                    ZA4->(MsUnLock())
                EndIf
                
                For nZ := 1 To len(aNFS)
                    If ZA8->(DbSeek(xFilial("ZA8") + ZA4->ZA4_CHAVE + PADR(aNFS[nZ],Len(ZA4->ZA4_CHAVE)) ))
                        RecLock('ZA8', .F.)
                            ZA8->(DbDelete())
                        ZA8->(MsUnLock())
                    EndIf
                Next
                
            EndIf
            RECOVER
        END SEQUENCE 
    EndIf
	
Return


//-------------------------------------------------------------------
/*/{Protheus.doc} GetVazio
Retorna valor vazio de acordo com tipo
@author  Samuel Dantas
@since   16/12/2019
@version version
/*/
//-------------------------------------------------------------------
Static Function GetVazio(cTab,cCampo)
    Local xRet := (cTab)->&(cCampo)

    If ValType(xRet) == 'N'
        xRet := 0
    ElseIf ValType(xRet) == 'L'
        xRet := .F.
    ElseIf ValType(xRet) == 'C'
        xRet := ""
    ElseIf ValType(xRet) == 'M'
        xRet := ""
    ElseIf ValType(xRet) == 'D'
        xRet := CtoD("")
    EndIf

Return xRet


//-------------------------------------------------------------------
/*/{Protheus.doc} ProcCte
Funcao que faz o processamento dos CTEs incluindo os pedidos de venda
@author  Sidney Sales
@since   11/12/19
@version 1.0
/*/
//-------------------------------------------------------------------
Static Function ProcCte(aVends)
    Local oPedido	:=	TNYXPEDIDOS():New() //Classe da Newib agisrn
    Local i
    Local __lChgX5FIL := .T.
    Local nDiasVenc := SuperGetMv("MS_DVENCT", .F., 365)
    Local _cNatureza := Alltrim(SuperGetMv("MS_NATCTE",.F.,"110101"))
    nDiasVenc := IIF( ValType(nDiasVenc) == 'C', Val(Alltrim(nDiasVenc)),nDiasVenc  )
    oPedido:lMostrarErro := .F.

    //Prepapara as propriedades do objeto pedidos
	oPedido:SERIENF	        := cYNumSer                       //SERIE DA NOTA
	oPedido:TIPONF		    := 'N'						    //TIPO DA NOTA
	oPedido:FORMULNF	    := Space(len(SF2->F2_FORMUL))	//FORMULARIO PROPRIO	
    oPedido:aC6CUSTOMFIELDS := {}
    oPedido:aC5CUSTOMFIELDS := {}
    oPedido:aC5toF2			:= {}
    oPedido:lCRIASB2		:= .F.
    oPedido:cTIPOTIT		:= 'CTE'

    //Numero que sera usado para o CTE
    aAdd(oPedido:aNUMNFS, cYNumNF )
    
    oCrud := crud():new('SC5', nil, 1)	
    
    //Posiciona o cliente final
    SA1->(DbSetOrder(1))
    SA1->(DbSeek(xFilial('SA1') + cCODCli + cLojCli))

	oCrud:Set('C5_TIPO'		, 'N'	        )
	oCrud:Set('C5_EMISSAO'	, dDatabase     )	
	oCrud:Set('C5_CLIENTE'	, SA1->A1_COD	)
	oCrud:Set('C5_LOJACLI'	, SA1->A1_LOJA	)
	oCrud:Set('C5_TIPOCLI'	, SA1->A1_TIPO	)
	oCrud:Set('C5_CONDPAG'	, cCond         ) 	//CONDICAO DE PGTO
	oCrud:Set('C5_UFDEST'	, cUFDEST       ) 	//CONDICAO DE PGTO
	oCrud:Set('C5_UFORIG'	, cUFORIG       ) 	//CONDICAO DE PGTO
	oCrud:Set('C5_TPFRETE'	, cTpFrete      ) 	//CONDICAO DE PGTO
	oCrud:Set('C5_NATUREZ'	, _cNatureza      ) 	//CONDICAO DE PGTO

    For i := 1 to Len(aVends)
	    oCrud:Set('C5_VEND'  + cValToChar(i)	, aVends[i][4])//CONDICAO DE PGTO
	    oCrud:Set('C5_COMIS' + cValToChar(i)	, aVends[i][3])//CONDICAO DE PGTO
        aAdd(oPedido:aC5CUSTOMFIELDS, 'C5_VEND' + cValToChar(i))
        aAdd(oPedido:aC5CUSTOMFIELDS, 'C5_COMIS' + cValToChar(i))
    Next

    oCrud:addChild('SC6')	
    
    cItem := '00'
    
    oCrud:addLine("SC6")

    cItem       := Soma1(cItem)
    cProduto    := PADR(SuperGetMv('MS_PRODFRE',.F., 'FRETE          '), Len(SB1->B1_COD) )
    oCrud:set('C6_PRODUTO'	, cProduto)
    oCrud:set('C6_ITEM'		, cItem)
    oCrud:set('C6_QTDVEN'	, 1)
    oCrud:set('C6_QTDLIB'	, 0)
    oCrud:set('C6_PRCVEN'	, nVlrFrete)
    oCrud:set('C6_PRUNIT'	, nVlrFrete)
    oCrud:set('C6_TES'		, cTes) 
    oCrud:set('C6_CONTA'	, GetConta()) 

    aAdd(oPedido:SC5, oCrud) 

    If Type('oProcess') != 'U'        	
        oProcess:IncRegua2("Gerando NF " +  cYNumNF + '/' + cYNumSer)
	EndIf
    If SX5->(DbSeek(cFilAnt + "01"+PADR("2",LEN(SX5->X5_CHAVE))))
        RecLock('SX5', .F.)
            SX5->X5_DESCRI := cYNumNF
        SX5->(MsUnLock())
    EndIf
    
    If ! oPedido:INCLUIRNF()
        MostraErro(GetSrvProfString("Startpath","") , cArqErro )
        cMensagem := Alltrim(MemoRead( GetSrvProfString("Startpath","") + '\' + cArqErro ))
    Else    

        If Type('oProcess') != 'U'        	
            oProcess:IncRegua2("Atualizando dados fiscais " +  cYNumNF + '/' + cYNumSer)
        EndIf

        //Retorno para a rotina
        cMensagem := 'NF Gerada'
        nRecNoSF2 := oPedido:DOCS[1]:nRecNo

        RecLock('SE1', .F.)
            SE1->E1_PREFIXO  := cYNumSer
            SE1->E1_VENCTO  := dDataBase + nDiasVenc
            SE1->E1_VENCREA := dDataBase + nDiasVenc
            SE1->E1_YTPFRET := cTpFrete
        SE1->(MsUnLock())

        RecLock('SC5', .F.)
            SC5->C5_TPFRETE := cTpFrete
        SC5->(MsUnLock())
        //Posiciona o SF2 para montar os dados que serao usados para as tabelas
        // SF2->(DbGoTo(nRecNoSF2))        
        ProcFis(.T.,nRecNoSF2)    
    EndIf

Return cMensagem


//-------------------------------------------------------------------
/*/{Protheus.doc} getTes
Funcao que retorna a tes de acordo com o CFOP
@author  Sidney Sales
@since   11/12/19
@version 1.0
/*/
//-------------------------------------------------------------------

Static Function getTes(cCFOP,cCst)
    Local cRet := ''

    cQuery := " SELECT F4_CODIGO FROM " + RetSqlTab('SF4')
    cQuery += " WHERE D_E_L_E_T_ = ''"
    cQuery += " AND RIGHT(RTRIM(F4_CF),3) = '" + Right(cCFOP,3) + "' "
    cQuery += " AND F4_TIPO    = 'S'"  //TES DE SAIDA
    cQuery += " AND F4_DUPLIC  = 'S' " //TEM QUE GERAR DUPLICATA
    cQuery += " AND F4_ESTOQUE = 'N'"  //NAO PODE MOVIMENTAR ESTOQUE
    cQuery += " AND F4_FILIAL = '" + xFilial("SF4") + "' "
    //NP3-26/10/2022 - Ajuste de TES que não calcula ICM
    If AllTrim(cCst) == '90'
        cQuery += " AND F4_ICM = 'N' "
    EndIf

    If Select('QRYTES') > 0
        QRYTES->(DbCloseArea())
    EndIf

    TcQuery cQuery New Alias 'QRYTES'

    If QRYTES->(!Eof())
        cRet := QRYTES->F4_CODIGO
    EndIf

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

//-------------------------------------------------------------------
/*/{Protheus.doc} getArray
Funcao que retorna um arary com os campos e dados que deverao ser gravados nas tabelas
@author  Sidney Sales
@since   11/12/19
@version 1.0
/*/
//-------------------------------------------------------------------

Static Function getArray(oCTE, cTab, nRecNoSF2)
    Local aRet := {}
    Local i
    Local cAtivCPB   := SuperGetMv('MS_ATIVCPB',.F., '00000140')

    SF2->(DbGoTo(nRecNoSF2))
    nAliqCPB := SuperGetMv('MS_ALIQCPB', .F., 1.5)  

    If cTab == 'SF2'
        aAdd(aRet, {'F2_CHVNFE' , oCTE:_CTEPROC:_PROTCTE:_INFPROT:_CHCTE:TEXT})        
        aAdd(aRet, {'F2_HORA'   , SubStr(oCTE:_CTEPROC:_CTE:_INFCTE:_IDE:_DHEMI:TEXT, 12,5) })
        aAdd(aRet, {'F2_UFORIG' , oCTE:_CTEPROC:_CTE:_INFCTE:_IDE:_UFINI:TEXT })
        aAdd(aRet, {'F2_UFDEST' , oCTE:_CTEPROC:_CTE:_INFCTE:_IDE:_UFFIM:TEXT })        
        aAdd(aRet, {'F2_EST'    , oCTE:_CTEPROC:_CTE:_INFCTE:_IDE:_UFFIM:TEXT })    
        nAliqIcms := WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_IMP:_ICMS:_ICMS00:_pICMS:TEXT","string")
        nValIcms  := WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_IMP:_ICMS:_ICMS00:_vICMS:TEXT","string")
        If ValType( nAliqIcms) != 'U' .AND. ValType( nValIcms) != 'U'
            aAdd(aRet, {'F2_ALIQICM', Val(nAliqIcms) })        
            aAdd(aRet, {'F2_VALICM', Val(nValIcms) })        
        EndIf                      
    ElseIf cTab == 'SD2'
        aAdd(aRet, {'D2_CF'     , oCTE:_CTEPROC:_CTE:_INFCTE:_IDE:_CFOP:TEXT}) 
        aAdd(aRet, {'D2_EST'    , oCTE:_CTEPROC:_CTE:_INFCTE:_IDE:_UFFIM:TEXT })                                  
        aAdd(aRet, {'D2_ALIQCPB', nAliqCPB})
        aAdd(aRet, {'D2_VALCPB' , Round(SF2->F2_BASIMP6 * (nAliqCPB / 100),2) })
        aAdd(aRet, {'D2_BASECPB', SF2->F2_BASIMP6 })
        aAdd(aRet, {'D2_CONTA', GetConta() })
        nAliqIcms := WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_IMP:_ICMS:_ICMS00:_pICMS:TEXT","string")
        nValIcms  := WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_IMP:_ICMS:_ICMS00:_vICMS:TEXT","string")
        If ValType( nAliqIcms) != 'U' .AND. ValType( nValIcms) != 'U'
            aAdd(aRet, {'D2_PICM', Val(nAliqIcms) })        
            aAdd(aRet, {'D2_VALICM', Val(nValIcms) })        
        EndIf  

    ElseIf cTab == 'SF3'
        aAdd(aRet, {'F3_CHVNFE' , oCTE:_CTEPROC:_PROTCTE:_INFPROT:_CHCTE:TEXT})
        aAdd(aRet, {'F3_CFO'    , oCTE:_CTEPROC:_CTE:_INFCTE:_IDE:_CFOP:TEXT})         
        aAdd(aRet, {'F3_ESPECIE', 'CTE'})         
        aAdd(aRet, {'F3_HORNFE' , SubStr(oCTE:_CTEPROC:_CTE:_INFCTE:_IDE:_DHEMI:TEXT, 12,5) })
        aAdd(aRet, {'F3_ESTADO' , oCTE:_CTEPROC:_CTE:_INFCTE:_IDE:_UFFIM:TEXT })      
        aAdd(aRet, {'F3_CLIDEST', cCODDes })
        aAdd(aRet, {'F3_LOJDEST', cLojDes })
        
        IF ValType(WSAdvValue(oCTE,"oCTE:_CTEPROC:_CTE:_INFCTE:_IMP:_VTOTTRIB:TEXT","string")) != "U"
            aAdd(aRet, {'F3_OBSICM', Val(oCTE:_CTEPROC:_CTE:_INFCTE:_IMP:_VTOTTRIB:TEXT) })
        EndIf
        
        //CAMPOS CALCULADOS
        aAdd(aRet, {'F3_CSTPIS' , '01'})
        aAdd(aRet, {'F3_CSTCOF' , '01'})
        
        SF2->(DbGoTo(nRecNoSF2))
        nAliqCPB := SuperGetMv('MS_ALIQCPB', .F., 1.5)
        
        aAdd(aRet, {'F3_ALIQCPB', nAliqCPB })
        aAdd(aRet, {'F3_VALCPB' , Round(SF2->F2_BASIMP6 * (nAliqCPB / 100),2)})
        aAdd(aRet, {'F3_BASECPB', SF2->F2_BASIMP6 })
        
        aAdd(aRet, {'F3_CODRSEF', oCTE:_CTEPROC:_PROTCTE:_INFPROT:_CSTAT:TEXT })
        aAdd(aRet, {'F3_ATIVCPB' , cAtivCPB})
        aAdd(aRet, {'F3_CONTA' , GetConta()})
        

        nAliqIcms := WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_IMP:_ICMS:_ICMS00:_pICMS:TEXT","string")
        nValIcms  := WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_IMP:_ICMS:_ICMS00:_vICMS:TEXT","string")
        If ValType( nAliqIcms) != 'U' .AND. ValType( nValIcms) != 'U'
            aAdd(aRet, {'F3_ALIQICM', Val(nAliqIcms) })        
            aAdd(aRet, {'F3_VALICM', Val(nValIcms) })        
        EndIf  

        If AllTrim(ZA4->ZA4_ACAO) == 'cancelar'
            aAdd(aRet, {'F3_DTCANC', StoD(StrTran(oCTE:_CTEPROC:_protCTe:_infProt:_dhRecbto:TEXT,"-","")) })    
        ElseIf AllTrim(ZA4->ZA4_ACAO) == 'inutilizar'   
            aAdd(aRet, {'F3_DTCANC', StoD(StrTran(oCTE:_CTEPROC:_protCTe:_infProt:_dhRecbto:TEXT,"-","")) })    
            aAdd(aRet, {'F3_CODRSEF', "102" })
            aAdd(aRet, {'F3_OBSERV', "NF INUTILIZADA" })
            aAdd(aRet, {'F3_CHVNFE ', "" })     
        EndIf

    ElseIf cTab == 'SFT'        
        aAdd(aRet, {'FT_CHVNFE' , oCTE:_CTEPROC:_PROTCTE:_INFPROT:_CHCTE:TEXT})
        aAdd(aRet, {'FT_CFOP'   , oCTE:_CTEPROC:_CTE:_INFCTE:_IDE:_CFOP:TEXT})         
        aAdd(aRet, {'FT_HORNFE' , SubStr(oCTE:_CTEPROC:_CTE:_INFCTE:_IDE:_DHEMI:TEXT, 12,5) })
        aAdd(aRet, {'FT_ESTADO' , oCTE:_CTEPROC:_CTE:_INFCTE:_IDE:_UFFIM:TEXT })      
        aAdd(aRet, {'FT_CLIDEST', cCODDes })
        aAdd(aRet, {'FT_LOJDEST', cLojDes })
        aAdd(aRet, {'FT_VALCONT', Val(oCTE:_CTEPROC:_CTE:_INFCTE:_VPREST:_VTPREST:TEXT) })
        aAdd(aRet, {'FT_TOTAL'  , Val(oCTE:_CTEPROC:_CTE:_INFCTE:_VPREST:_VTPREST:TEXT) })
               
        // aAdd(aRet, {'FT_OBSERV' , oCTE:_CTEPROC:_CTE:_INFCTE:_IMP:_VTOTTRIB:TEXT })
        
        nValMerc  := WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_INFCTENORM:_INFCARGA:_VCARGA:TEXT","string")
        If ValType( nValMerc) != 'U'
            For i := 1 to Len(oCTE:_CTEPROC:_CTE:_INFCTE:_INFCTENORM:_INFCARGA:_INFQ)
                If Alltrim(oCTE:_CTEPROC:_CTE:_INFCTE:_INFCTENORM:_INFCARGA:_INFQ[i]:_TPMED:TEXT) == 'PESO REAL'        
                    aAdd(aRet, {'FT_PESO'   , Val(oCTE:_CTEPROC:_CTE:_INFCTE:_INFCTENORM:_INFCARGA:_INFQ[i]:_QCARGA:TEXT)})
                EndIf
                If Alltrim(oCTE:_CTEPROC:_CTE:_INFCTE:_INFCTENORM:_INFCARGA:_INFQ[i]:_TPMED:TEXT) == "QTDE DE VOLUMES"
                    aAdd(aRet, {'FT_QUANT'   , Val(oCTE:_CTEPROC:_CTE:_INFCTE:_INFCTENORM:_INFCARGA:_INFQ[i]:_QCARGA:TEXT)})
                EndIf         
            Next
        EndIf

        aAdd(aRet, {'FT_ATIVCPB' , cAtivCPB})
        aAdd(aRet, {'FT_CSTPIS'  , '01'})
        
        nAliqPis    := GetMV('MV_TXPIS')        

        aAdd(aRet, {'FT_BASEPIS' , SF2->F2_BASIMP6})
        aAdd(aRet, {'FT_ALIQPIS' , nAliqPis})
        aAdd(aRet, {'FT_VALPIS'  , Round(SF2->F2_BASIMP6 * nAliqPis / 100,2)})
        
        nAliqCof    := GetMV('MV_TXCOFIN')
        
        aAdd(aRet, {'FT_CSTCOF'  , '01'})
        aAdd(aRet, {'FT_BASECOF' , SF2->F2_BASIMP5})
        aAdd(aRet, {'FT_ALIQCOF' , nAliqCof})
        aAdd(aRet, {'FT_VALCOF'  , Round(SF2->F2_BASIMP5 * nAliqCof / 100,2)})
        
        aAdd(aRet, {'FT_CONTA'  , GetConta()})

        aAdd(aRet, {'FT_ALIQCPB', nAliqCPB })
        aAdd(aRet, {'FT_VALCPB' , Round(SF2->F2_BASIMP6 * 1.5 / 100,2)})
        aAdd(aRet, {'FT_BASECPB', SF2->F2_BASIMP6 })        

        aAdd(aRet, {'FT_CTIPI' , SuperGetMv('MS_CTIPI', .F., '53') })        
        aAdd(aRet, {'FT_GRPCST', SuperGetMv('MS_GRPCST', .F., '999') })     
        
        nAliqIcms := WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_IMP:_ICMS:_ICMS00:_pICMS:TEXT","string")
        nValIcms  := WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_IMP:_ICMS:_ICMS00:_vICMS:TEXT","string")
        If ValType( nAliqIcms) != 'U' .AND. ValType( nValIcms) != 'U'
            aAdd(aRet, {'FT_ALIQICM', Val(oCTE:_CTEPROC:_CTE:_INFCTE:_IMP:_ICMS:_ICMS00:_pICMS:TEXT) })        
            aAdd(aRet, {'FT_VALICM', Val(oCTE:_CTEPROC:_CTE:_INFCTE:_IMP:_ICMS:_ICMS00:_vICMS:TEXT) })        
        EndIf

        aAdd(aRet, {'FT_CLASFIS', StrZero(Val(SFT->FT_CLASFIS),3) })        
        
        If Alltrim(ZA4->ZA4_ACAO) == 'cancelar'
            aAdd(aRet, {'FT_DTCANC', StoD(StrTran(oCTE:_CTEPROC:_PROTCTE:_infProt:_dhRecbto:TEXT,"-","")) })        
         ElseIf AllTrim(ZA4->ZA4_ACAO) == 'inutilizar'   
            aAdd(aRet, {'FT_DTCANC', StoD(StrTran(oCTE:_CTEPROC:_protCTe:_infProt:_dhRecbto:TEXT,"-","")) })    
            aAdd(aRet, {'FT_OBSERV', "NF INUTILIZADA" })
            aAdd(aRet, {'FT_CHVNFE ', "" })     
        EndIf

    ElseIf cTab == 'DT6'

        aAdd(aRet, {'DT6_FILIAL' , cFilAnt})        
        aAdd(aRet, {'DT6_FILDOC' , cFilAnt})        
        aAdd(aRet, {'DT6_FILORI' , cFilAnt})        
        aAdd(aRet, {'DT6_FILVGA' , cFilAnt})        
        aAdd(aRet, {'DT6_CHVCTE' , oCTE:_CTEPROC:_PROTCTE:_INFPROT:_CHCTE:TEXT})        
        aAdd(aRet, {'DT6_TIPO'   , 'CTE'})        
        aAdd(aRet, {'DT6_SERIE'  , cYNumSer })
        aAdd(aRet, {'DT6_DOC'    , cYNumNF})
        aAdd(aRet, {'DT6_VALIMP' , SF2->F2_VALICM })
        aAdd(aRet, {'DT6_DATEMI', StoD(StrTran(Left(oCTE:_CTEPROC:_CTE:_INFCTE:_IDE:_DHEMI:TEXT,10),'-','')) })
        aAdd(aRet, {'DT6_HOREMI' , StrTran(SubStr(oCTE:_CTEPROC:_CTE:_INFCTE:_IDE:_DHEMI:TEXT, 12,5),':','') })
        aAdd(aRet, {'DT6_TIPTRA' , oCTE:_CTEPROC:_CTE:_INFCTE:_IDE:_MODAL:TEXT })
        
        aAdd(aRet, {'DT6_CDRDES' , oCTE:_CTEPROC:_CTE:_INFCTE:_IDE:_CMUNFIM:TEXT })
        aAdd(aRet, {'DT6_CDRCAL' , oCTE:_CTEPROC:_CTE:_INFCTE:_IDE:_CMUNFIM:TEXT })
        
        cDescri := oCTE:_CTEPROC:_CTE:_INFCTE:_IDE:_xMunIni:TEXT 
        cCodMun := oCTE:_CTEPROC:_CTE:_INFCTE:_IDE:_CMUNINI:TEXT
        cPais  := WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_REM:_enderReme:_CPAIS:TEXT","string")
        If ValType( cPais) == 'U' 
            cPais   := "105"
        EndIf
        cEst    := oCTE:_CTEPROC:_CTE:_INFCTE:_IDE:_UFIni:TEXT
        
        cCDRORI := GrvGrpReg(cDescri, cCodMun, cPais, cEst)
        aAdd(aRet, {'DT6_CDRORI' , cCDRORI })
        cDescri := oCTE:_CTEPROC:_CTE:_INFCTE:_IDE:_xMunFIM:TEXT 
        cCodMun := oCTE:_CTEPROC:_CTE:_INFCTE:_IDE:_CMUNFIM:TEXT
        cPais  := WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_dest:_enderDest:_CPAIS:TEXT","string")
        If ValType( cPais) == 'U' 
            cPais   := "105"
        EndIf
        cEst    := oCTE:_CTEPROC:_CTE:_INFCTE:_IDE:_UFFIM:TEXT
        
        cCDRDES := GrvGrpReg(cDescri, cCodMun, cPais, cEst)
        cCDRCAL := cCDRDES
        aAdd(aRet, {'DT6_CDRDES' , cCDRDES })
        aAdd(aRet, {'DT6_CDRCAL' , cCDRCAL })

        aAdd(aRet, {'DT6_CLICAL' , cCODCli })
        aAdd(aRet, {'DT6_LOJCAL' , cLojCli })
        aAdd(aRet, {'DT6_NOMCAL' , cNomCli })

        aAdd(aRet, {'DT6_CLIREM' , cCODRem })
        aAdd(aRet, {'DT6_LOJREM' , cLojRem })
        aAdd(aRet, {'DT6_NOMREM' , cNomRem })

        aAdd(aRet, {'DT6_CLIDES' , cCODDes })
        aAdd(aRet, {'DT6_LOJDES' , cLojDes })
        aAdd(aRet, {'DT6_NOMDES' , cNomDes })

        aAdd(aRet, {'DT6_CLIDEV' , cCODDes })
        aAdd(aRet, {'DT6_LOJDEV' , cLojDes })

        aAdd(aRet, {'DT6_VALTOT' , Val(oCTE:_CTEPROC:_CTE:_INFCTE:_VPREST:_VTPREST:TEXT) })
        aAdd(aRet, {'DT6_VALFAT' , Val(oCTE:_CTEPROC:_CTE:_INFCTE:_VPREST:_VTPREST:TEXT) })
        aAdd(aRet, {'DT6_VALFRE' , Val(oCTE:_CTEPROC:_CTE:_INFCTE:_VPREST:_VTPREST:TEXT) })
        
        nAcresFret := 0
        
        //Calcula o valor do acrescimo ao frete
        If Valtype(oCTE:_CTEPROC:_CTE:_INFCTE:_VPREST:_COMP) == "A"
            For i := 1 to oCTE:_CTEPROC:_CTE:_INFCTE:_VPREST:_COMP
                If oCTE:_CTEPROC:_CTE:_INFCTE:_VPREST:_COMP[i]:_XNOME:TEXT != "FRETE PESO"
                    nAcresFret += Val(oCTE:_CTEPROC:_CTE:_INFCTE:_VPREST:_VTPREST:TEXT)
                EndIf            
            Next
        Else
            If oCTE:_CTEPROC:_CTE:_INFCTE:_VPREST:_COMP:_xNome:TEXT != "FRETE PESO"
                nAcresFret += Val(oCTE:_CTEPROC:_CTE:_INFCTE:_VPREST:_COMP:_VCOMP:TEXT)
            EndIf
        EndIF

        aAdd(aRet, {'DT6_ACRESC', nAcresFret })
        
        nValMerc  := WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_INFCTENORM:_INFCARGA:_VCARGA:TEXT","string")
        If ValType( nValMerc) != 'U'
            aAdd(aRet, {'DT6_VALMER' , Val(oCTE:_CTEPROC:_CTE:_INFCTE:_INFCTENORM:_INFCARGA:_VCARGA:TEXT) })
            //Percorre o arary para achar o peso e a quantidade
            For i := 1 to Len(oCTE:_CTEPROC:_CTE:_INFCTE:_INFCTENORM:_INFCARGA:_INFQ)
                If Alltrim(oCTE:_CTEPROC:_CTE:_INFCTE:_INFCTENORM:_INFCARGA:_INFQ[i]:_TPMED:TEXT) == 'PESO REAL'        
                    aAdd(aRet, {'DT6_PESO'   , Val(oCTE:_CTEPROC:_CTE:_INFCTE:_INFCTENORM:_INFCARGA:_INFQ[i]:_QCARGA:TEXT)})
                    aAdd(aRet, {'DT6_PESOM3' , Val(oCTE:_CTEPROC:_CTE:_INFCTE:_INFCTENORM:_INFCARGA:_INFQ[i]:_QCARGA:TEXT)})
                EndIf
                If Alltrim(oCTE:_CTEPROC:_CTE:_INFCTE:_INFCTENORM:_INFCARGA:_INFQ[i]:_TPMED:TEXT) == "QTDE DE VOLUMES"
                    aAdd(aRet, {'DT6_QTDVOL'   , Val(oCTE:_CTEPROC:_CTE:_INFCTE:_INFCTENORM:_INFCARGA:_INFQ[i]:_QCARGA:TEXT)})
                    aAdd(aRet, {'DT6_VOLORI'   , Val(oCTE:_CTEPROC:_CTE:_INFCTE:_INFCTENORM:_INFCARGA:_INFQ[i]:_QCARGA:TEXT)})
                EndIf         
            Next
        EndIf
        aAdd(aRet, {'DT6_PROCTE'   , oCTE:_CTEPROC:_PROTCTE:_INFPROT:_NPROT:TEXT })
        aAdd(aRet, {'DT6_RETCTE'   , oCTE:_CTEPROC:_PROTCTE:_INFPROT:_NPROT:TEXT })

        cStatusCTE := oCTE:_CTEPROC:_PROTCTE:_INFPROT:_CSTAT:TEXT

        //autorizado,cancelado
        If cStatusCTE $ '100,101'    
            cSITCTE := '2'        
        //denegado
        ElseIf cStatusCTE $ '110'
            cSITCTE := '3'        
        //contingencia
        ElseIf cStatusCTE $ '108,109'
            cSITCTE := '4'
        EndIf

        aAdd(aRet, {'DT6_SITCTE' , cSITCTE })
        aAdd(aRet, {'DT6_CODOBS' , "" })
        aAdd(aRet, {'DT6_CODOBS' , "" })
        aAdd(aRet, {'DT6_TIPFRE' , IIF( cTpFrete == 'C', '1', '2') })
        cOBS := WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_COMPL:_XOBS:TEXT","string")
       

    EndIf

Return aRet

//-------------------------------------------------------------------
/*/{Protheus.doc} GetNfs
G
@author  Samuel Dantas
@since   17/12/2019
@version version
/*/
//-------------------------------------------------------------------
Static Function getNfs(oCTE)
    Local aRet := {}
    Local xAux 
    Local i
    
    xAux := WSAdvValue(oCTE,"_CTEPROC:_CTE:_INFCTE:_INFCTENORM:_INFDOC:_INFNFE","array")
    
    If ValType(xAux) != 'U'
        If ValType(xAux) == 'A'
            For i := 1 to Len(xAux)
                aAdd(aRet, xAux[i]:_CHAVE:TEXT)
            Next
        ElseIf ValType(xAux) == 'O'
            aAdd(aRet, xAux:_CHAVE:TEXT)
        EndIf
    EndIf

Return aRet 

//-------------------------------------------------------------------
/*/{Protheus.doc} ExcluiCTE
Exclui CTE
@author  Samuel Dantas
@since   17/12/2019
@version version
/*/
//-------------------------------------------------------------------
Static Function ExcluiCTE()
    
    Local oPedido	:=	TNYXPEDIDOS():New() //Classe da Newib agisrn
    Local aAreaZA4 := ZA4->(GetArea())
    Local lRetorno := .F.
    Local cCNPJ := WSAdvValue(oCTE,"_procEventoCTe:_eventoCTe:_infEvento:_CNPJ:TEXT","string") 
    
    cCnpj := SubStr(Alltrim(ZA4->ZA4_CHAVE), 7, 14)

    If ValType(cCNPJ) == "U"
        cMensagem := '_procEventoCTe:_eventoCTe:_infEvento:_CNPJ INVALIDA ' 
        return '_procEventoCTe:_eventoCTe:_infEvento:_CNPJ INVALIDA ' 
    EndIf

    If ! setEmpresa(cCNPJ)
        cMensagem := 'Cadastro de filial não localizado com o CNPJ ' + cCNPJ
        return 'Cadastro de filial não localizado com o CNPJ ' + cCNPJ
    EndIf
    
    //Valida antes de cancelar
    lRetorno := VldExist()
    BEGIN TRANSACTION
    If lRetorno
        cQryExl := " SELECT R_E_C_N_O_ AS RECNO FROM "  + RetSqlName('SF2') + " SF2"
        cQryExl += " WHERE SF2.D_E_L_E_T_ <> '*' AND F2_CHVNFE = '"+ZA4->ZA4_CHAVE+"' "
        
        If Select('QRYEXL') > 0
            QRYEXL->(dbclosearea())
        EndIf
        
        TcQuery cQryExl New Alias 'QRYEXL'
        
        If QRYEXL->(!Eof())
            nRecnoAux := QRYEXL->RECNO
            SF2->(DbGoTo(QRYEXL->RECNO))
                oPedido:DOC         := SF2->F2_DOC
                oPedido:SERIENF     := SF2->F2_SERIE
                oPedido:CLIENTENF   := SF2->F2_CLIENTE
                oPedido:LOJACLINF   := SF2->F2_LOJA
                oPedido:TIPONF      := SF2->F2_TIPO
                oPedido:lCRIASB2      := .F.
                If ! oPedido:EXCLUIRNF()
                    MostraErro(GetSrvProfString("Startpath","") , cArqErro )
                    cMensagem := Alltrim(MemoRead( GetSrvProfString("Startpath","") + '\' + cArqErro ))
                Else
                    oCTE := xmlParser(ZA4->ZA4_BODY, "_", @cErro, @cAviso)    
                    //Cancela fiscal
                    SF2->(DbGoTo(nRecnoAux))    
                    cMensagem := CancFis()
                    If Empty(cMensagem)
                        cMensagem := 'NF Gerada'
                    Else
                        DisarmTransaction()
                    EndIf
                    
                EndIf
        EndIF
    Else
        cMensagem := "Nota não foi incluída no Protheus, portanto não pode ser cancelada. Não foi encontrado registro (ZA4) de inclusão para esta nota."
    EndIf
    END TRANSACTION
Return cMensagem

//-------------------------------------------------------------------
/*/{Protheus.doc} GravaSA1
Grava destinatário, caso necessário
@author  Samuel Dantas
@since   17/12/2019
@version version
/*/
//-------------------------------------------------------------------
Static Function GravaSA1(cCNPJ, cNome)
    Local aVetor        := {}
    Local lMsErroAuto   := .F.
    Local cRet := ""
    Local cPessoa := IIF(LEN(ALLTRIM(cCNPJ)) == 14, 'J', 'F')
    Local cCodSa1 := IIF( cPessoa == 'J', SUBSTR(ALLTRIM(cCNPJ),1,8), SUBSTR(ALLTRIM(cCNPJ),1,9) )
    Local cCodLoj := IIF( cPessoa == 'J', SUBSTR(ALLTRIM(cCNPJ),9,4), '0000' )
        
    RecLock('SA1', .T.)
        SA1->A1_FILIAL := xFilial("SA1")
        SA1->A1_COD := cCodSa1
        SA1->A1_LOJA := cCodLoj
        SA1->A1_CGC := cCNPJ
        SA1->A1_NOME := cNome
        SA1->A1_NREDUZ := cNome
        SA1->A1_PESSOA := IIF(LEN(ALLTRIM(cCNPJ)) == 14, 'J', 'F')
    SA1->(MsUnLock())
            
Return cRet

//-------------------------------------------------------------------
/*/{Protheus.doc} VAlida e cria ZA4 de inclusão
Antes realizar o cancelamento valida a existência de um registro de
inclusão na tabela ZA4.
@author  Samuel Dantas
@since   18/12/2019
@version version
/*/
//-------------------------------------------------------------------
Static Function VldExist()
    Local cQuery    := ""
    Local aAreaZA4  := ZA4->(GetArea())
    Local lRet      := .F.
    Local cMsg      := ""
    Local cChvZA4   := ZA4->ZA4_CHAVE
    Local cBody     := ""
    Local cChave    := ""
    //Verifica se existe cadastro de inclusão de nota 
    //Vindo a partir da intergração com o NUCCI
    cQryZA4u := " SELECT * FROM "  + RetSqlName('ZA4') + " ZA4"
    cQryZA4u += " WHERE ZA4.D_E_L_E_T_ <> '*' AND ZA4_CHAVE = '"+ZA4->ZA4_CHAVE+"' "
    cQryZA4u += " AND ZA4_ACAO = 'inclusao' AND ZA4_STATUS = 'P' "
    
    If Select('QRYZA4U') > 0
        QRYZA4U->(dbclosearea())
    EndIf
    
    TcQuery cQryZA4u New Alias 'QRYZA4U'
    
    If QRYZA4U->(!Eof())
        lRet := .T.
        QRYZA4U->(dbSkip())
    EndIf

    
Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} InutCTE
Inutilização de CTE
@author  Samuel Dantas
@since   18/12/2019
@version version
/*/
//-------------------------------------------------------------------
Static Function InutCTE() 
    
    Local aCabec := {}
    Local aItens := {}
    Local aLinha := {}
    Local cProduto := GetMv("MV_INUTPRO") //PADR(SuperGetMv('MV_INUTPRO',.F., 'FRETE          '), Len(SB1->B1_COD) )
    Local cCliInu  := GetMv("MV_INUTCLI") //
    Local cLojInu  := GetMv("MV_INUTLOJ") //
    Local cTesInu  := GetMv("MV_INUTTES") 
    
    PRIVATE lMsErroAuto := .F.

    cStat       := WSAdvValue(oCTE,"_soap_Envelope:_soap_Body:_cteInutilizacaoCTResult:_retInutCTe:_infInut:_cStat:TEXT"   ,"string")
    cCNPJ       := WSAdvValue(oCTE,"_soap_Envelope:_soap_Body:_cteInutilizacaoCTResult:_retInutCTe:_infInut:_CNPJ:TEXT"    ,"string")
    cSerie      := WSAdvValue(oCTE,"_soap_Envelope:_soap_Body:_cteInutilizacaoCTResult:_retInutCTe:_infInut:_SERIE:TEXT"   ,"string")
    cYNumNFIni  := WSAdvValue(oCTE,"_soap_Envelope:_soap_Body:_cteInutilizacaoCTResult:_retInutCTe:_infInut:_nCTIni:TEXT"  ,"string")
    cYNumNFFim  := WSAdvValue(oCTE,"_soap_Envelope:_soap_Body:_cteInutilizacaoCTResult:_retInutCTe:_infInut:_nCTFin:TEXT"  ,"string")
    dDataRec    := WSAdvValue(oCTE,"_soap_Envelope:_soap_Body:_cteInutilizacaoCTResult:_retInutCTe:_infInut:_dhRecbto:TEXT","string")
    If ValType(cCNPJ) == "U"
        cMensagem := "XML inválido entre em contato o TI NUCCI para resolução."+CRLF+" A 'tag_soap_Envelope:_soap_Body:_cteInutilizacaoCTResult:_retInutCTe:_infInut:_CNPJ' não está presente no XML. " 
        return cMensagem
    EndIf

    If ! setEmpresa(cCNPJ)
        cMensagem := "XML inválido entre em contato o TI NUCCI para resolução."+CRLF+" A 'tag_soap_Envelope:_soap_Body:_cteInutilizacaoCTResult:_retInutCTe:_infInut:_CNPJ' não está presente no XML " 
        return 'Cadastro de filial não localizado com o CNPJ ' + cCNPJ
    EndIf

    cYNumNFIni  := PadL(cYNumNFIni, len(SF2->F2_DOC), "0")
    cYNumNFFim  := PadL(cYNumNFFim, len(SF2->F2_DOC), "0")
    dDataRec    := StoD(StrTran(dDataRec,"-",""))
    dDataBase := dDataRec
    cMensagem := ""
    While cYNumNFIni <= cYNumNFFim
        
        aCabec := {}
        aItens := {}

        aadd(aCabec,{"F2_TIPO"     , "N"            })
        aadd(aCabec,{"F2_FORMUL"   , "N"            })
        aadd(aCabec,{"F2_DOC"      , cYNumNFIni     })
        aadd(aCabec,{"F2_SERIE"    , cSerie         })
        aadd(aCabec,{"F2_EMISSAO"  , dDataBase      })
        aadd(aCabec,{"F2_CLIENTE"  , PADR(cCliInu,LEN(SA1->A1_COD))        })
        aadd(aCabec,{"F2_LOJA"     , cLojInu        })
        aadd(aCabec,{"F2_ESPECIE"  , "CTE"          })
        aadd(aCabec,{"F2_COND"     , "001"          })
        aadd(aCabec,{"F2_DESCONT"  , 0              })
        aadd(aCabec,{"F2_FRETE"    , 0              })
        aadd(aCabec,{"F2_SEGURO"   , 0              })
        aadd(aCabec,{"F2_DESPESA"  , 0              })
    
        
        aLinha := {}
        aadd(aLinha,{"D2_COD"   , cProduto          , Nil})
        aadd(aLinha,{"D2_ITEM"  , StrZero(1,2)      , Nil})
        aadd(aLinha,{"D2_QUANT" , 1                 , Nil})
        aadd(aLinha,{"D2_PRCVEN", 0.01                 , Nil})
        aadd(aLinha,{"D2_TOTAL" , 0.01                 , Nil})
        aadd(aLinha,{"D2_TES"   , cTesInu           , Nil})
        aadd(aItens,aLinha)

        MSExecAuTo({|x,y| MATA920(x,y)},aCabec, aItens, 3)

        If !lMsErroAuto
            SF2->(DbSetOrder(1))
            If SF2->(DbSeek(xFilial("SF2") + PADR(cYNumNFIni,Len(SF2->F2_DOC)) + PADR(cSerie,Len(SF2->F2_SERIE)) + PADR(cCliInu,Len(SF2->F2_CLIENTE)) + PADR(cLojInu,Len(SF2->F2_LOJA))))
                ErrorBlock( {|e| cErrorBlock := e:Description + e:ErrorStack })
                BEGIN SEQUENCE
                    cXml := ZA4->ZA4_BODY
                    cXml := StrTran(cXml,"env:","soap:")
                    oCTE := xmlParser(cXml, "_", @cErro, @cAviso)    
                    InutFis(.T.,SF2->(Recno()))
                    SD2->(DbSetOrder(3))
                    If SD2->(DbSeek(xFilial("SD2") + SF2->(F2_DOC + F2_SERIE + F2_CLIENTE + F2_LOJA + PADR(cProduto,LEN(SD2->D2_COD))) ))
                        RecLock('SD2', .F.)
                            SD2->(DbDelete())
                        SD2->(MsUnLock())
                    EndIf
                    RecLock('SF2', .F.)
                        SF2->(DbDelete())
                    SF2->(MsUnLock())
                        
                    cMensagem := "NF Gerada"

                    RECOVER
                END SEQUENCE 
                
            EndIf
        Else
            MostraErro(GetSrvProfString("Startpath","") , cArqErro )
            cMensagem := Alltrim(MemoRead( GetSrvProfString("Startpath","") + '\' + cArqErro ))
        EndIf
        cYNumNFIni := Soma1(cYNumNFIni)

    EndDo

Return 

//-------------------------------------------------------------------
/*/{Protheus.doc} GrvGrpReg
Grava grupo de região na tabela DUY, usado no módulo do TMS.
@author  Samuel Dantas
@since   18/12/2019
@version version
/*/
//-------------------------------------------------------------------
Static Function GrvGrpReg(cDescri, cCodMun, cPais, cEst)
    Local cQuery    := ""
    Local cCodDUY   := ""

    cQuery := " SELECT R_E_C_N_O_ AS RECNO FROM "  + RetSqlName('DUY') + " DUY"
    cQuery += " WHERE DUY.D_E_L_E_T_ <> '*' AND DUY_CODMUN = '"+Right(cCodMun,5)+"' AND DUY_EST = '"+cEst+"' "
    
    If Select('QRY') > 0
        QRY->(dbclosearea())
    EndIf
    
    TcQuery cQuery New Alias 'QRY'
    DUY->(DbSetOrder(1))
    If QRY->(Eof())
        RecLock('DUY', .T.)
            DUY->DUY_GRPVEN := GetSeqDUY()
            DUY->DUY_DESCRI := cDescri
            DUY->DUY_PAIS   := cPais
            DUY->DUY_CODMUN := Right(cCodMun,5)
            DUY->DUY_EST    := cEst   
        DUY->(MsUnLock())
        ConfirmSx8()
    Else
        DUY->(DbGoTo(QRY->RECNO))
    EndIf

    cCodDUY := DUY->DUY_GRPVEN

Return cCodDUY

//-------------------------------------------------------------------
/*/{Protheus.doc} GetSeqDUY
Busca uma nova sequencia
@author  Samuel Dantas
@since   19/12/2019
@version version
/*/
//-------------------------------------------------------------------
Static Function GetSeqDUY() 
    Local cRet := StrZero(0, len(DUY->DUY_GRPVEN))

    cQry := " SELECT MAX(DUY_GRPVEN) AS MAX FROM "  + RetSqlName('DUY') + " DUY"
    cQry += " WHERE DUY.D_E_L_E_T_ <> '*' "
    
    If Select('QRYDUY') > 0
        QRYDUY->(dbclosearea())
    EndIf
    
    TcQuery cQry New Alias 'QRYDUY'
    
    If QRYDUY->(!Eof())
        cRet :=  IIF(Empty(QRYDUY->MAX), cRet , Soma1(QRYDUY->MAX))
        QRYDUY->(dbSkip())
    EndIf

Return cRet

//-------------------------------------------------------------------
/*/{Protheus.doc} dadosvend
Recupera de um objeto xml os dados dos vendedores enviados na integracao
com o nucci
@author  Sidney Sales
@since   26/12/19
@version 1.0
/*/
//-------------------------------------------------------------------

Static Function dadosVend(aXMLVend, nTotal)
    Local aRet := {}
    Local i 
    Local aAux := {'C','N','V'} //CPF, Nome e Valor
    
    //So trata se receber um array
    If ValType(aXMLVend) == 'A'
        
        //Cria uma matriz que ira guardar os vendedores - CPF, Nome, Valor, codinterno, percentual
        aRet := Array(3,5) 
        
        //Percorre o array pegando os dados
        For i := 1 to Len(aXMLVend)

            //Campo descritivo que representa o vendedor
            cCampo := WSAdvValue(aXMLVEND[i],"_XCAMPO:TEXT","string")            

            //Campo que representa o valor do campo para o vendedor
            cTexto := WSAdvValue(aXMLVEND[i],IIF(IsInCallStack('ProcessaNFSE'),"_P1_XTEXTO:TEXT" , "_XTEXTO:TEXT"),"string")            

            //Verifica se sao campos validos
            If ValType(cCampo) == 'C' .AND. ValType(cTexto) == 'C'

                //Verifica se e o campo customizado, podem haver outros tipos de campo
                If Left(cCampo,7) == 'LAUTOV_'

                    //Recupera o numero do vendedor
                    nVend       := Val(SubStr(cCampo, 8,1))

                    //Pega o tipo do campo
                    cTpCampo    := Right(cCampo,1)

                    //Veriica a posicao que ira colocar no array
                    nPosCampo   := aScan(aAux, cTpCampo)

                    //Se os dados forem validos, adiciona no retorno
                    If nVend > 0 .AND. nPosCampo > 0
                        aRet[nVend][nPosCampo] := cTexto                        
                    EndIf
                
                EndIf
            EndIf
        Next

    EndIf
    
    //Preenche o campo de codigo do vendedor
    For i := 1 to Len(aRet)
        If ValType(aRet[i][1]) != 'U' .AND. ValType(aRet[i][2]) != 'U' .AND. ValType(aRet[i][3]) != 'U'
            nValComis  := Val(aRet[i][3])
            //Calcula o valor do perecentual da comissao
            If ValType(nValComis) != 'U' .AND. nValComis > 0
                nPerComis  := Round(nValComis / nTotal * 100,2)
            Else
                nPerComis  := 0
            EndIf
            aRet[i][3] := nPerComis
            aRet[i][4] := retCodVend(aRet[i])
        EndIf
    Next

Return aRet

//-------------------------------------------------------------------
/*/{Protheus.doc} retCodVEnd
    (long_description)
    @type  Static Function
    @author Sidney Sales
    @since   27/12/2019
    @version 1.0
/*/
//-------------------------------------------------------------------
Static Function retCodVEnd(aDados)
    Local cQuery 
    Local cRet
    
    cQuery := " SELECT A3_COD FROM " + RetSqlTab('SA3') 
    cQuery += " WHERE " + RetSqlDel('SA3')
    cQuery += " AND RTRIM(A3_CGC) = '" + Alltrim(aDados[1]) + "' "
    
    If Select('QRYSA3') > 0
        QRYSA3->(DbCloseArea())
    EndIf
        
    TcQuery cQuery New Alias 'QRYSA3'

    If QRYSA3->(!Eof())
        cRet := QRYSA3->A3_COD
    Else
        cRet := GETSXENUM('SA3', 'A3_COD')
        ConfirmSx8()
        RecLock('SA3', .T.)
            SA3->A3_FILIAL  := xFilial('SA3')
            SA3->A3_COD     := cRet
            SA3->A3_CGC     := aDados[1]
            SA3->A3_NOME    := aDados[2]
            SA3->A3_NREDUZ  := aDados[2]
        SA3->(MsUnLock())
    EndIf

Return cRet

Static Function GetConta()
    Local aContas   := {}
    Local cRet      := Alltrim(SuperGetMv("MS_CTACTE", .f., ''))

Return cRet

Static Function CancFis()
    Local cCNPJ     := ""
    Local cDtcanc   := ""
    Local cCodigo   := ""

    Local aAreaSF3 := SF3->(GetArea())
    Local aAreaSF2 := SF2->(GetArea())
    Local aAreaSFT := SFT->(GetArea())
    Local aAreaDT6 := DT6->(GetArea())
    Local aAreaSE1 := SE1->(GetArea())
    Local cRetorno     := ""
    
    SE1->(DbSetOrder(1))    
    SF3->(DbSetOrder(4))    
    SFT->(DbSetOrder(1))    
    DT6->(DbSetOrder(1))    

    cDtcanc := WSAdvValue(oCTE,"_PROCEVENTOCTE:_EVENTOCTE:_infEvento:_dhevento:TEXT","string") 
    If ValType(cDtcanc) == "U"
        cDtcanc := WSAdvValue(oCTE,"_retEventoCTe:_infEvento:_dhRegEvento:TEXT","string") 
        If ValType(cDtcanc) == "U"
             cDtcanc := WSAdvValue(oCTE,"_PROCEVENTOCTE:_retEventoCTe:_infEvento:_dhRegEvento:TEXT","string") 
        EndIf
    EndIf

    If ValType(cDtcanc) == "U"
        cRetorno := "Data do evento não encontrada no XML." + CRLF
    EndIf
    
    dDtCanc := StoD(StrTran(cDtcanc,"-",""))

    If SFT->(DbSeek(SF2->F2_FILIAL + 'S' + SF2->(F2_SERIE + F2_DOC + F2_CLIENTE + F2_LOJA)))
        RecLock('SFT', .F.)
            SFT->FT_DTCANC := dDtCanc
        SFT->(MsUnLock())
    Else
        cRetorno := "SFT não encontrada." + CRLF
    EndIf 

    If SF3->(DbSeek(SF2->(F2_FILIAL + F2_CLIENTE + F2_LOJA + F2_DOC + F2_SERIE)))
        RecLock('SF3', .F.)
            SF3->F3_DTCANC := dDtCanc
            SF3->F3_CODRSEF := "101" 
        SF3->(MsUnLock())
    Else
        cRetorno := "SF3 não encontrada." + CRLF
    EndIf

     //Exclui DT6
    DT6->(DBOrderNickname("DT6CHVCTE"))
    If DT6->(DbSeek(xFilial("DT6") + SF2->F2_CHVNFE))
        RecLock('DT6', .F.)
            DT6->(DbDelete())
        DT6->(MsUnLock())
    EndIf
    SE1->(DbSetOrder(2))
    If SE1->(DbSeek(xFilial("SE1") + SF2->(F2_CLIENTE + F2_LOJA + F2_SERIE + F2_DOC ) ))
        lMsErroAuto := .F.
        aVetor := {}

        aAdd(aVetor, {"E1_FILIAL"   , SF2->F2_FILIAL                     } )
        aAdd(aVetor, {"E1_NUM"      , SF2->F2_DOC                        } )
        aAdd(aVetor, {"E1_PREFIXO"  , SF2->F2_SERIE                      } )
        aAdd(aVetor, {"E1_CLIENTE"  , SF2->F2_CLIENTE                    } )
        aAdd(aVetor, {"E1_LOJA"     , SF2->F2_LOJA                        } )
        aAdd(aVetor, {"E1_TIPO"     , PADR(SF2->F2_ESPECIE,LEN(SE1->E1_TIPO)) } )

        MSExecAuTo({|x,y|FINA040(x,y)},aVetor,5)
        If lMsErroAuto
            cRetorno := "Error ao cancelar título (SE1)"
        EndIf
    Else
        cRetorno := "Error ao cancelar título (SE1)"
    EndIf


    SF3->(RestArea(aAreaSF3))
    SFT->(RestArea(aAreaSFT))
    SF2->(RestArea(aAreaSF2))
    SE1->(RestArea(aAreaSE1))
    DT6->(RestArea(aAreaDT6))

Return cRetorno


Static Function InutFIS()
    Local cCNPJ     := ""
    Local cDtcanc   := ""
    Local cCodigo   := ""

    Local aAreaSF3 := SF3->(GetArea())
    Local aAreaSF2 := SF2->(GetArea())
    Local aAreaSFT := SFT->(GetArea())
    Local aAreaDT6 := DT6->(GetArea())
    Local aAreaSE1 := SE1->(GetArea())
    Local cRetorno     := ""
    
    SE1->(DbSetOrder(1))    
    SF3->(DbSetOrder(4))    
    SFT->(DbSetOrder(1))    
    
              
    If SFT->(DbSeek(SF2->F2_FILIAL + 'S' + SF2->(F2_SERIE + F2_DOC + F2_CLIENTE + F2_LOJA)))
        RecLock('SFT', .F.)
            SFT->FT_DTCANC := StoD(StrTran(oCTE:_soap_Envelope:_soap_Body:_cteInutilizacaoCTResult:_retInutCTe:_infInut:_dhRecbto:TEXT,"-",""))
            SFT->FT_OBSERV := "NF INUTILIZADA"
            SFT->FT_CHVNFE := "'"
        SFT->(MsUnLock())  
    Else
        cRetorno := "SFT não encontrada." + CRLF
    EndIf 

    If SF3->(DbSeek(SF2->(F2_FILIAL + F2_CLIENTE + F2_LOJA + F2_DOC + F2_SERIE)))
        RecLock('SF3', .F.)
            SF3->F3_DTCANC := StoD(StrTran(oCTE:_soap_Envelope:_soap_Body:_cteInutilizacaoCTResult:_retInutCTe:_infInut:_dhRecbto:TEXT,"-",""))
            SF3->F3_CODRSEF := "102" 
            SF3->F3_OBSERV := "NF INUTILIZADA" 
            SF3->F3_CHVNFE := "" 
        SF3->(MsUnLock())
    Else
        cRetorno := "SF3 não encontrada." + CRLF
    EndIf

    SE1->(DbSetOrder(2))
    If SE1->(DbSeek(xFilial("SE1") + SF2->(F2_CLIENTE + F2_LOJA + F2_SERIE + F2_DOC ) ))
        lMsErroAuto := .F.
        aVetor := {}

        aAdd(aVetor, {"E1_FILIAL"   , SF2->F2_FILIAL                     } )
        aAdd(aVetor, {"E1_NUM"      , SF2->F2_DOC                        } )
        aAdd(aVetor, {"E1_PREFIXO"  , SF2->F2_SERIE                      } )
        aAdd(aVetor, {"E1_CLIENTE"  , SF2->F2_CLIENTE                    } )
        aAdd(aVetor, {"E1_LOJA"     , SF2->F2_LOJA                        } )
        aAdd(aVetor, {"E1_TIPO"     , PADR(SF2->F2_ESPECIE,LEN(SE1->E1_TIPO)) } )

        MSExecAuTo({|x,y|FINA040(x,y)},aVetor,5)
        If lMsErroAuto
            cRetorno := "Error ao salvar SE1"
        EndIf
    EndIf


    SF3->(RestArea(aAreaSF3))
    SFT->(RestArea(aAreaSFT))
    SF2->(RestArea(aAreaSF2))
    SE1->(RestArea(aAreaSE1))

Return cRetorno


User Function LAUF5CAN
    
    If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01RN0001'
    EndIf
    
    
    cQryAux :=  "SELECT ZA4.R_E_C_N_O_ AS RECNO FROM "+RetSqlName("ZA4")+ " ZA4 "
    cQryAux += " WHERE  ZA4_ACAO = 'inclusao' AND ZA4_CHAVE = '23200107189259000686570020000682791957718900' AND ZA4.D_E_L_E_T_ <> '*'  "

    If Select('QRYAUX') > 0
        QRYAUX->(dbclosearea())
    EndIf    

    TcQuery cQryAux New Alias 'QRYAUX'

    While QRYAUX->(!Eof())
        U_LAUF0005(QRYAUX->RECNO)
        QRYAUX->(dbSkip())
    EndDo


Return
