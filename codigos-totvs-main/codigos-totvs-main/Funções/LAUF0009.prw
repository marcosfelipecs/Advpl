#INCLUDE 'Rwmake.ch'
#INCLUDE 'Protheus.ch'
#INCLUDE 'TbIconn.ch'
#INCLUDE 'Topconn.ch'

//-------------------------------------------------------------------
/*/{Protheus.doc} LAUF0009
Funcao para processamento das notas fiscais de servicos
@author  Sidney Sales
@since   16/12/2019
@version 1.0
/*/
//-------------------------------------------------------------------
User Function LAUF0009(nRecnoZA5, nOpc,lForceProcess)    

    Private cErrorBlock := ''
    Private cYNumSer    := ''
    Private cYNumNF     := ''
    Private lForceProc  := IIf(ValType(lForceProcess)=="L", lForceProcess, .F.)


    If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01RN0001'
        nRecnoZA5 := 252984
        nOpc := 3
    EndIf
    cEmailSup := Alltrim(SuperGetMv("MS_TIMAIL",.T.,"diego.cunha@lauto.com.br"))
    ZA5->(DbGoTo(nRecnoZA5))
    nOpc := IIF( Alltrim(ZA5->ZA5_ACAO) == 'inclusao', 3, 5 )		 
    cRet := ProcessaNFSE(nRecnoZA5, nOpc)        
    RecLock('ZA5', .F.)
        If nOpc == 3
            ZA5->ZA5_STATUS := Iif(cRet != 'NFse Gerada', 'E','P')
            ZA5->ZA5_ERRO   := Iif(cRet != 'NFse Gerada', cRet,'')
        Else
            ZA5->ZA5_STATUS := Iif(cRet != 'NFse excluida', 'E','P')
            ZA5->ZA5_ERRO   := Iif(cRet != 'NFse excluida', cRet,'')
        EndIf
    ZA5->(MsUnLock())

    If ZA5->ZA5_STATUS == 'E' .AND. !Empty(cEmailSup)
        U_EnviaEmail(" Processamento CTE - Erro", cEmailSup, "Foi encontrada a seguinte msg de erro ( ZA5_CHAVE = "+ZA5->ZA5_CHAVE+") :"+CRLF+Alltrim(ZA5->ZA5_ERRO), "", "", .F.)
    EndIf
Return

Static Function ProcessaNFSE(nRecnoZA5, nOpc)
    Local cAviso    := ''
    Local cErro     := ''
    Local oPedido	:=	TNYXPEDIDOS():New() //Classe da Newib agisrn
    Local cArqErro	:= "erroauto.txt"    
    Local cCond     := SuperGetMv("MS_CONDTRA",.F.,"001")    
    Local cTes      := SuperGetMv("MS_TESCTE",.F.,"60A")  
    Local dMVULMES   := SuperGetMv("MV_ULMES")
    Local dMVDBLQMOV := SuperGetMV("MV_DBLQMOV")
    Local oPedido	:= TNYXPEDIDOS():New() //Classe da Newib agisrn
    Local lMostrarErro := .F.
    Local aRegSD2 := {}
    Local aRegSE1 := {}
    Local aRegSE2 := {}
    Local aAreaSA1 := SA1->(GetArea())
    Local cMsg := ""
    Local lRet := .F.
    Local cCodRetIs := ""
    Local i := 1
    Local cProduto := ""
    Private _cNatureza := ""
    Private lForceProcess := lForceProc // Adiciona variável private
    
    dDataAux := dDatabase
    ZA5->(DbGoTo(nRecnoZA5))

    If nOpc == 5
        SF2->(DbSetOrder(1))//F2_FILIAL, F2_DOC, F2_SERIE, F2_CLIENTE, F2_LOJA
        
        // Seta a database
        cMensagem := AtuZA5() 
        // cFIlAnt := ZA5->ZA5_FILNFS       
        If !Empty(cMensagem)
            Return cMensagem
        EndIf
        BEGIN TRANSACTION
            If SF2->(DbSeek(xFilial('SF2') + ZA5->(ZA5_NUMNF + ZA5->ZA5_SERNF) ))
                nTipo := GetMv("MV_TIPOPRZ")
                nDias := GetMv("MV_EXCNFS")
                PutMV("MV_TIPOPRZ",2)
                PutMV("MV_EXCNFS",5000)
                aHeader := {}
                // dDatabase := SF2->F2_EMISSAO
                //Exclui efetivamente a nota
                nRecnoSF2 := SF2->(Recno())
                If SF2->F2_EMISSAO <= dMVULMES .Or. SF2->F2_EMISSAO <= dMVDBLQMOV
                    lRet := .F.
                    DisarmTransaction()
                    cMensagem := "Exclusão não autorizada. Data emissão da nota inferior a data de fechamento do estoque/compras (MV_ULMES: " + DtoC(dMVULMES) + ") / (MV_DBLQMOV: " + DtoC(dMVDBLQMOV) + "). Para prosseguir, configure o parâmetro com data anterior a emissão da Nota (" + DtoC(SF2->F2_EMISSAO) + ")."
                    Return cMensagem
                EndIf

                If MaCanDelF2("SF2",SF2->(RecNo()),@aRegSD2,@aRegSE1,@aRegSE2)
                    lRet := SF2->(MaDelNFS(aRegSD2,aRegSE1,aRegSE2,.F.,.F.,.F.,.F.))
                EndIf

                If !lRet
                    cMensagem := "Exclusão não autorizada"
                Else
                    cMensagem := ExcluRegs(nRecnoSF2)
                    If Empty(cMensagem)
                        cMensagem := 'NFse excluida'
                    Else 
                        DisarmTransaction()
                    EndIf
                EndIf
                PutMV("MV_TIPOPRZ",nTipo)
                PutMV("MV_EXCNFS",nDias)
            Else        
                If Alltrim(ZA5->ZA5_ACAO) == 'excluir'
                    SET DELETED OFF	
                    If SF2->(DbSeek(xFilial('SF2') + ZA5->(ZA5_NUMNF + ZA5->ZA5_SERNF) ))
                        nRecnoSF2 := SF2->(Recno())
                        ExcluRegs(nRecnoSF2, .T.)
                        cMensagem := 'NFse excluida'
                    Else
                        cMensagem := "NF Não localizada para exclusão"
                    EndIf
                    SET DELETED ON
                Else
                    cMensagem := "NF Não localizada para exclusão"
                EndIf                

            EndIf
        
            dDatabase := dDataAux
        
        END TRANSACTION
    Else
        cXml := ZA5->ZA5_BODY
        
        //Transforma o NFSE em um objeto
        oNFSE := xmlParser(cXml, "_", @cErro, @cAviso)  
        oNFSEAux   := oNFSE
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
        aAdd(aVars, {"Quantidade"       , "nQuant"      , WSAdvValue(oNFSE,"_P1_IDENTIFICACAORPS:_P1_TIPO:TEXT","string") })
        aAdd(aVars, {"Valor Servico"    , "nValor"      , WSAdvValue(oNFSE,"_P1_SERVICO:_P1_VALORES:_P1_VALORSERVICOS:TEXT","string") })
        aAdd(aVars, {"Servico"          , "cServico"    , WSAdvValue(oNFSE,"_P1_SERVICO:_P1_ITEMLISTASERVICO:TEXT","string") })
        aAdd(aVars, {"ValorPis"         , "nVlrPis"     , WSAdvValue(oNFSE,"_P1_SERVICO:_P1_VALORES:_P1_VALORPIS:TEXT","string") })
        aAdd(aVars, {"Valor Cofins"     , "nVlrCofins"  , WSAdvValue(oNFSE,"_P1_SERVICO:_P1_VALORES:_P1_VALORCOFINS:TEXT","string") })
        aAdd(aVars, {"Valor INSS"       , "nVlrInss"    , WSAdvValue(oNFSE,"_P1_SERVICO:_P1_VALORES:_P1_VALORINSS:TEXT","string") })
        aAdd(aVars, {"Valor IR"         , "nVlrIr"      , WSAdvValue(oNFSE,"_P1_SERVICO:_P1_VALORES:_P1_VALORIR:TEXT","string") })
        aAdd(aVars, {"Valor CSLL"       , "nVlrCsll"    , WSAdvValue(oNFSE,"_P1_SERVICO:_P1_VALORES:_P1_VALORCSLL:TEXT","string") })
        aAdd(aVars, {"Valor ISS"        , "nVlrISS"    , WSAdvValue(oNFSE,"_P1_SERVICO:_P1_VALORES:_P1_VALORISS:TEXT","string") })
        aAdd(aVars, {"Aliquito"         , "nAliqISS"    , WSAdvValue(oNFSE,"_P1_SERVICO:_P1_VALORES:_P1_ALIQUOTA:TEXT","string") })
        aAdd(aVars, {"Valor ISS Retido" , "nVlrIssRe"   , WSAdvValue(oNFSE,"_P1_SERVICO:_P1_VALORES:_P1_VALORISSRETIDO:TEXT","string") })
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
        Next i
        
        SB1->(DbSetOrder(1))
        
        cYNumNF     := IIF(LEN(cYNumNf) > 9, RIGHT(cYNumNf, 9), cYNumNf)
        cYNumNF     := PADL(cYNumNF,TAMSX3("F2_DOC")[1],"0")
        
        dEmissao    := StoD(StrTran(Left(dEmissao,10),'-',''))
        dCompet := WSAdvValue(oNFSE,"_P1_DATACOMPETENCIA:TEXT","string")
        
        If ValType(dCompet) != "U"
            dCompet     := StoD(StrTran(Left(dCompet,10),'-',''))
            If dCompet != dEmissao
                dEmissao := dCompet
            EndIf
        EndIf
        
        dDatabase   := dEmissao

        nQuant      := Val(nQuant)
        nValor      := Val(nValor)
        nVlrPis     := Val(nVlrPis)
        nVlrCofins  := Val(nVlrCofins)
        nVlrInss    := Val(nVlrInss)
        nVlrIr      := Val(nVlrIr)
        nVlrCsll    := Val(nVlrCsll)
        nVlrIssRe   := Val(nVlrIssRe)
        nVlrIss     := Val(nVlrIss)
        nAliqISS    := Val(nAliqISS)
        cYNumSer    := "NUC"

        If !setEmpresa(cCNPJEmit)
            return 'Cadastro de filial não localizado com o CNPJ ' + cCNPJ
        EndIf
        
        cYNumSer    := u_getSerie(cYNumSer)
                
        cCNPJREM := WSAdvValue(oNFSE,"_P1_REMETENTE:_P1_IDENTIFICACAOREMETENTE:_P1_CPFCNPJ:_P1_CNPJ:TEXT","string") 
        cCNPJDes := WSAdvValue(oNFSE,"_P1_DESTINATARIO:_P1_IDENTIFICACAODESTINATARIO:_P1_CPFCNPJ:_P1_CNPJ:TEXT","string") 
        
        cCNPJREM := IIF(ValType(cCNPJREM) != "U", Alltrim(cCNPJREM), "" )
        cCNPJDes := IIF(ValType(cCNPJREM) != "U", Alltrim(cCNPJDEs), "" )

        cCodREM := ""
        cLojREM := ""
        cCodDes := ""
        cLojDes := ""

        If len(cCNPJDes) > 10 .AND. len(cCNPJREM) > 10
            cCodREM := IIF(len(cCNPJREM) == 14, SUBSTR(ALLTRIM(cCNPJREM),1,8), SUBSTR(ALLTRIM(cCNPJREM),1,9) )
            cLojREM := IIF(len(cCNPJREM) == 14, SUBSTR(ALLTRIM(cCNPJREM),9,4), "0000")
            cCodDes := IIF(len(cCNPJDes) == 14, SUBSTR(ALLTRIM(cCNPJDes),1,8), SUBSTR(ALLTRIM(cCNPJDes),1,9) )
            cLojDes := IIF(len(cCNPJDes) == 14, SUBSTR(ALLTRIM(cCNPJDes),9,4), "0000")
        EndIf
        
        
            
        SA1->(dbSetOrder(3))
         //grava os dados da nfse
        RecLock('ZA5', .F.)    
            If SA1->(DbSeek( xFilial("SA1") + PADR(cCNPJToma,TamSx3("A1_CGC")[1]) ))
                ZA5->ZA5_CLIENT := SA1->A1_COD
                ZA5->ZA5_LOJA   := SA1->A1_LOJA
            EndIf
            ZA5->ZA5_REM    := cCodREM
            ZA5->ZA5_REMLOJ := cLojREM
            ZA5->ZA5_DEST   := cCodDes
            ZA5->ZA5_DESLOJ := cLojDes
            ZA5->ZA5_SERNF  := cYNumSer
            ZA5->ZA5_NUMNF  := cYNumNF
            ZA5->ZA5_FILNFS := cFIlAnt
        ZA5->(MsUnLock())
        
        If "." $ cServico
            cServico := Alltrim(cServico)
        Else
            cServico := SubStr(cServico,1,2)+"."+SubStr(cServico,3,4)
        EndIf
        
        If SX5->(DbSeek(cFilAnt + "01"+PADR(cYNumSer,LEN(SX5->X5_CHAVE))))
            RecLock('SX5', .F.)
                SX5->X5_DESCRI := cYNumNF
            SX5->(MsUnLock())
        EndIf
        ZA9->(DbSetOrder(1))
        
        cQryZA9 := " SELECT ZA9.R_E_C_N_O_ AS RECNOZA9 FROM "  + RetSqlName('ZA9') + " ZA9 "
        cQryZA9 += " INNER JOIN "+RetSqlName('SB1')+" SB1 ON LEFT(ZA9_FILIAL,2) = B1_FILIAL AND B1_COD = ZA9_PROD AND ZA9.D_E_L_E_T_ = SB1.D_E_L_E_T_ "
        cQryZA9 += " INNER JOIN "+RetSqlName('SED')+" SED ON LEFT(ZA9_FILIAL,2) = ED_FILIAL AND ED_CODIGO = ZA9_NATURE AND ZA9.D_E_L_E_T_ = SED.D_E_L_E_T_ "
        cQryZA9 += " WHERE ZA9.D_E_L_E_T_ <> '*' AND ZA9_FILIAL = '"+cFIlAnt+"' AND LEFT(ZA9_PROD,5) = '"+cServico+"' "
        
        If nVlrPis > 0
            cQryZA9 += " AND B1_PIS = '1' "
            cMsg += "Não encontrado no cadastro produto x natureza (ZA9), produto com B1_PIS = '1' ." + CRLF
        Else
            cMsg += "Não encontrado no cadastro produto x natureza (ZA9), produto com B1_PIS = '2' ." + CRLF
            cQryZA9 += " AND B1_PIS = '2'"
        EndIf
        If nVlrCofins > 0
            cQryZA9 += " AND B1_COFINS = '1' "
            cMsg += "Não encontrado no cadastro produto x natureza (ZA9), produto com B1_COFINS = '1'." + CRLF
        Else
            cMsg += "Não encontrado no cadastro produto x natureza (ZA9), produto com B1_COFINS = '2'." + CRLF
            cQryZA9 += " AND B1_COFINS = '2' "
        EndIf

        If nAliqISS > 0 .AND. nAliqISS != 0.05
            cQryZA9 += " AND B1_ALIQISS = "+cValToChar(nAliqISS*100)+" "
            cMsg += "Não encontrado no cadastro produto x natureza (ZA9), produto com B1_ALIQISS = '"+cValToChar(nAliqISS*100)+"' ." + CRLF
        ElseIf nAliqISS == 0.05
            cQryZA9 += " AND (B1_ALIQISS = "+cValToChar(nAliqISS*100)+" OR B1_ALIQISS = 0) "
        EndIf
        
        If nVlrCsll > 0
            cQryZA9 += " AND B1_CSLL = '1' "
            cMsg += "Não encontrado no cadastro produto x natureza (ZA9), produto com B1_CSLL = '1' ." + CRLF
        Else
            cMsg += "Não encontrado no cadastro produto x natureza (ZA9), produto com B1_CSLL = '2' ." + CRLF
            cQryZA9 += " AND B1_CSLL = '2' "
        EndIf
        If nVlrInss > 0
            cQryZA9 += " AND B1_INSS = 'S' "
            cMsg += "Não encontrado no cadastro produto x natureza (ZA9), produto com B1_INSS = 'S' ." + CRLF
        Else
            cMsg += "Não encontrado no cadastro produto x natureza (ZA9), produto com B1_INSS = 'N' ." + CRLF
            cQryZA9 += " AND B1_INSS = 'N' "
        EndIf
        If nVlrIr > 0
            cQryZA9 += " AND B1_IRRF = 'S' "
            cMsg += "Não encontrado no cadastro produto x natureza (ZA9), produto com B1_IRRF = 'S' ." + CRLF
        Else
            cMsg += "Não encontrado no cadastro produto x natureza (ZA9), produto com B1_IRRF = 'N' ." + CRLF
            cQryZA9 += " AND B1_IRRF = 'N' "
        EndIf
        
        If nVlrIssRe > 0
            cQryZA9 += " AND ED_CALCISS = 'S' "
            cMsg += "Não encontrado no cadastro produto x natureza (ZA9), natureza com ED_CALCISS = 'S' ." + CRLF
        Else
            cQryZA9 += " AND ED_CALCISS = 'N' "
            cMsg += "Não encontrado no cadastro produto x natureza (ZA9), natureza com ED_CALCISS = 'N' ." + CRLF
        EndIf

        If Select('QRYZA9') > 0
            QRYZA9->(dbclosearea())
        EndIf
                    
        TcQuery cQryZA9 New Alias 'QRYZA9'
        
        If QRYZA9->(!Eof())
            ZA9->(DbGoTo(QRYZA9->RECNOZA9))
            _cNatureza := ZA9->ZA9_NATURE
            If ! SB1->(DbSeek(xFilial('SB1') + ZA9->ZA9_PROD))
                return 'Serviço ('+cServico+') não localizado no cadastro do Protheus (SB1)'  
            else
                cProduto := SB1->B1_COD
            EndIf
        Else
            return cMsg + 'Serviço ('+cServico+') não localizado no cadastro do Protheus (ZA9)'   
        EndIf

        //Posiciona o cliente pelo CNPJ do tomador
        SA1->(dbSetOrder(3))
        If ! SA1->(DbSeek( xFilial("SA1") + PADR(cCNPJToma,TamSx3("A1_CGC")[1]) ))
            Return 'Cliente(tomador) não cadastrado no Protheus ' + cCNPJToma
        EndIf  

    

        aAreaSA1 := SA1->(GetArea())
        cNatuAux := SA1->A1_NATUREZ
        cTes := SB1->B1_TS

        //grava os dados da nfse
        RecLock('ZA5', .F.)
            ZA5->ZA5_CLIENT := SA1->A1_COD
            ZA5->ZA5_LOJA   := SA1->A1_LOJA
            ZA5->ZA5_SERNF  := cYNumSer
            ZA5->ZA5_NUMNF  := cYNumNF
            ZA5->ZA5_FILNFS := cFIlAnt
        ZA5->(MsUnLock())

        aXMLVend := WSAdvValue(oNFSEAUX,"_P_ENVIARLOTERPSENVIO:_P_LOTERPS:_P1_COMPL:_P1_OBSCONT","string")
        aVends   := dadosVend(aXMLVend, nValor)
        oPedido:lMostrarErro := .F.

        //Prepapara as propriedades do objeto pedidos
        oPedido:SERIENF	        := cYNumSer                       //SERIE DA NOTA
        oPedido:TIPONF		    := 'N'						    //TIPO DA NOTA
        oPedido:FORMULNF	    := Space(len(SF2->F2_FORMUL))	//FORMULARIO PROPRIO	
        oPedido:aC6CUSTOMFIELDS := {}
        oPedido:aC5CUSTOMFIELDS := {}
        oPedido:aC5toF2			:= {}
        oPedido:lCRIASB2		:= .F.
        oPedido:cTIPOTIT		:= 'NFS'

        //Numero que sera usado para o CTE
        aAdd(oPedido:aNUMNFS, cYNumNF )
        
        oCrud := crud():new('SC5', nil, 1)	

        oCrud:Set('C5_TIPO'		, 'N'	        )
        oCrud:Set('C5_EMISSAO'	, dDatabase     )	
        oCrud:Set('C5_CLIENTE'	, SA1->A1_COD	)
        oCrud:Set('C5_LOJACLI'	, SA1->A1_LOJA	)
        oCrud:Set('C5_TIPOCLI'	, SA1->A1_TIPO	)
        oCrud:Set('C5_CONDPAG'	, cCond         ) 	//CONDICAO DE PGTO
        oCrud:Set('C5_NATUREZ'	, _cNatureza         ) 	//CONDICAO DE PGTO
        
        SED->(DbSetOrder(1))
        If SED->(DbSeek(xFilial('SED') + PADR(_cNatureza,len(SED->ED_CODIGO)) ))
            cCodRetIs := IIF(SED->ED_CALCISS == 'S', '1','2') //1
            aAdd(oPedido:aC5CUSTOMFIELDS, 'C5_RECISS')
            oCrud:Set('C5_RECISS'	, cCodRetIs)
        EndIF
        
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
        
        // oCrud:set('C6_PRODUTO'	, SB1->B1_COD)
        oCrud:set('C6_PRODUTO'	, cProduto)
        oCrud:set('C6_ITEM'		, cItem)
        oCrud:set('C6_QTDVEN'	, nQuant)
        oCrud:set('C6_QTDLIB'	, nQuant)
        oCrud:set('C6_PRCVEN'	, nValor)
        oCrud:set('C6_PRUNIT'	, nValor)
        oCrud:set('C6_TES'		, cTes) 
        oCrud:set('C6_CF'		, "999") 

        aAdd(oPedido:SC5, oCrud) 
        BEGIN TRANSACTION

            RecLock('SA1', .F.)
                SA1->A1_NATUREZ := ZA9->ZA9_NATURE
            SA1->(MsUnLock())
            
            If ! oPedido:INCLUIRNF()
                MostraErro(GetSrvProfString("Startpath","") , cArqErro )
                cMensagem := Alltrim(MemoRead( GetSrvProfString("Startpath","") + '\' + cArqErro ))
                
            Else    
                nRecNoSF2 := oPedido:DOCS[1]:nRecNo
                cRetorno := ValidaNat(nRecNoSF2, lForceProcess) // Passa o parâmetro
                
                If !Empty(cRetorno)
                    cMensagem := cRetorno
                    DisarmTransaction()
                Else
                    //Retorno para a rotina
                    cMensagem := 'NFse Gerada'
                EndIf
            EndIf

            SA1->(RestArea(aAreaSA1))
            
            RecLock('SA1', .F.)
                SA1->A1_NATUREZ := cNatuAux
            SA1->(MsUnLock())

        END TRANSACTION
    EndIF

Return cMensagem


//-------------------------------------------------------------------
/*/{Protheus.doc} VALIDANat
Valida natureza
@author  Samuel Dantas
@since   02/01/2020
@version version
/*/
//-------------------------------------------------------------------
Static Function ValidaNat(nRecnoSF2,lForceProcess)
    Local cRet := ""
    Local aAreaSF2 := SF2->(GetArea())
    Local aAreaSFT := SFT->(GetArea())
    Local aAreaSF3 := SF3->(GetArea())
    Local aAreaSD2 := SD2->(GetArea())
    Local aAreaSED := SED->(GetArea())
    Local nDiasVenc := SuperGetMv("MS_DVENCT", .F., 365)
    Default lForceProcess := .F.
    nDiasVenc := IIF( ValType(nDiasVenc) == 'C', Val(Alltrim(nDiasVenc)),nDiasVenc  )
    If nRecnoSF2 > 0 
        SF2->(DbGoTo(nRecnoSF2))
        
        SFT->(DbSetOrder(1))
        SF3->(DbSetOrder(5))
        SD2->(DbSetOrder(3))
        SC6->(DbSetOrder(4))
        nAliqCPB := SuperGetMv('MS_ALIQCPB', .F., 1.5)
        
        cConta := RetConta()
        If !SD2->(DbSeek(xFilial("SD2") + SF2->(F2_DOC + F2_SERIE + F2_CLIENTE + F2_LOJA)))
            cRet += "SD2 não encontra para o doc "+SF2->F2_DOC +"/"+SF2->F2_SERIE+"."
            Return cRet
        Else
            RecLock('SD2', .F.)
                SD2->D2_ALIQCPB := nAliqCPB
                SD2->D2_VALCPB  := Round(SF2->F2_BASIMP6 * (nAliqCPB / 100),2)
                SD2->D2_BASECPB := SF2->F2_BASIMP6
                SD2->D2_CONTA   := cConta	
            SD2->(MsUnLock())
            
        EndIf

        If !SFT->(DbSeek(xFilial("SFT") + SF2->("S" + F2_SERIE + F2_DOC + F2_CLIENTE + F2_LOJA)))
            cRet += "SFT não encontra para o doc "+SFT->FT_NFISCAL +"/"+SF2->FT_SERIE+"."
            Return cRet
        Else
            RecLock('SFT', .F.)
                SFT->FT_ATIVCPB := "00000140"
                SFT->FT_ALIQCPB := nAliqCPB
                SFT->FT_VALCPB  := Round(SF2->F2_BASIMP6 * (nAliqCPB / 100),2)
                SFT->FT_BASECPB := SF2->F2_BASIMP6
                SFT->FT_CONTA := cConta
            SFT->(MsUnLock())
        EndIf
        
        If !SF3->(DbSeek(xFilial("SF3") + SF2->( F2_SERIE + F2_DOC + F2_CLIENTE + F2_LOJA)))
            cRet += "SF3 não encontra para o doc "+SF2->F3_DOC +"/"+SF2->F3_SERIE+"."
            Return cRet
        Else
            RecLock('SF3', .F.)
                SF3->F3_ALIQCPB := nAliqCPB
                SF3->F3_VALCPB  := Round(SF2->F2_BASIMP6 * (nAliqCPB / 100),2)
                SF3->F3_BASECPB := SF2->F2_BASIMP6
                SF3->F3_ESPECIE := "NFS"
                SF3->F3_CONTA   := cConta
            SF3->(MsUnLock())
        EndIf

        If SC6->(DbSeek(xFilial("SC6") + SF2->(F2_DOC+F2_SERIE) ))
            RecLock('SC6', .F.)
                SC6->C6_CONTA := cConta
            SC6->(MsUnLock())
        EndIF

        RecLock('SE1', .F.)
            SE1->E1_PREFIXO := u_getSerie('NUC')
            SE1->E1_VENCTO  := dDataBase + nDiasVenc
            SE1->E1_VENCREA := dDataBase + nDiasVenc
        SE1->(MsUnLock())

        RecLock('SF2', .F.)
            SF2->F2_ESPECIE := "NFS"
            SF2->F2_PREFIXO := u_getSerie('NUC')
        SF2->(MsUnLock())

        //Validação do PIS <p1:ValorPis> D2_VALPIS F2_VALPIS FT_VRETPIS
        If SD2->D2_VALPIS != nVlrPis
            cRet += "Divegência entre o valor do PIS destacado no xml e valor cálculado no Protheus. (D2_VALPIS)." +" Valor gerado: "+ Transform(SD2->D2_VALPIS,"@E 999,999,999.99") +". Valor no xml: " + Transform(nVlrPis,"@E 999,999,999.99") + CRLF
        EndIf
        If SF2->F2_VALPIS != nVlrPis
            cRet += "Divegência entre o valor do PIS destacado no xml e valor cálculado no Protheus. (F2_VALPIS)." +" Valor gerado: "+ Transform(SF2->F2_VALPIS,"@E 999,999,999.99") +". Valor no xml: " + Transform(nVlrPis,"@E 999,999,999.99") + CRLF
        EndIf
        If SFT->FT_VRETPIS != nVlrPis
            cRet += "Divegência entre o valor do PIS destacado no xml e valor cálculado no Protheus. (FT_VRETPIS)." +" Valor gerado: "+ Transform(SFT->FT_VRETPIS,"@E 999,999,999.99")+". Valor no xml: " + Transform( nVlrPis,"@E 999,999,999.99") + CRLF
        EndIf

        //Validação do Cofins  D2_VALCOF F2_VALCOFI FT_VRETCOF
        If SD2->D2_VALCOF != nVlrCofins
            cRet += "Divegência entre o valor do COFINS destacado no xml e valor cálculado no Protheus. (D2_VALCOF)." +" Valor gerado: "+ Transform(SD2->D2_VALCOF,"@E 999,999,999.99") +". Valor no xml: " + Transform(nVlrCofins,"@E 999,999,999.99") + CRLF
        EndIf
        If SF2->F2_VALCOFI != nVlrCofins
            cRet += "Divegência entre o valor do COFINS destacado no xml e valor cálculado no Protheus. (F2_VALCOFI)." +" Valor gerado: "+ Transform(SF2->F2_VALCOFI,"@E 999,999,999.99")+". Valor no xml: "+ Transform( nVlrCofins,"@E 999,999,999.99") + CRLF
        EndIf
        If SFT->FT_VRETCOF != nVlrCofins
            cRet += "Divegência entre o valor do COFINS destacado no xml e valor cálculado no Protheus. (FT_VRETCOF)." +" Valor gerado: "+ Transform(SFT->FT_VRETCOF,"@E 999,999,999.99")+". Valor no xml: "+ Transform( nVlrCofins,"@E 999,999,999.99") + CRLF
        EndIf

        //Validação do Cofins <p1:ValorPis> D2_VALPIS F2_VALPIS FT_VRETPIS
        If SD2->D2_VALINS != nVlrInss
            cRet += "Divegência entre o valor do INSS destacado no xml e valor cálculado no Protheus. (D2_VALINS)." +" Valor gerado: "+ Transform(SD2->D2_VALINS,"@E 999,999,999.99") +". Valor no xml: " + Transform(nVlrInss,"@E 999,999,999.99") + CRLF
        EndIf
        If SF2->F2_VALINSS != nVlrInss
            cRet += "Divegência entre o valor do INSS destacado no xml e valor cálculado no Protheus. (F2_VALINSS)." +" Valor gerado: "+ Transform(SF2->F2_VALINSS,"@E 999,999,999.99")+". Valor no xml: " + Transform( nVlrInss,"@E 999,999,999.99") + CRLF    
        EndIf
        If SFT->FT_VALINS != nVlrInss
            cRet += "Divegência entre o valor do INSS destacado no xml e valor cálculado no Protheus. (FT_VALINS)." +" Valor gerado: "+ Transform(SFT->FT_VALINS,"@E 999,999,999.99") +". Valor no xml: " + Transform(nVlrInss,"@E 999,999,999.99") + CRLF
        EndIf

        //Validação do Cofins <p1:ValorPis> D2_VALPIS F2_VALPIS FT_VRETPIS
        If SD2->D2_VALIRRF != nVlrIr
            cRet += "Divegência entre o valor do IRRF destacado no xml e valor cálculado no Protheus. (D2_VALIRRF)." +" Valor gerado: "+ Transform(SD2->D2_VALIRRF,"@E 999,999,999.99")+". Valor no xml: " + Transform( nVlrIr,"@E 999,999,999.99") + CRLF
        EndIf
        If SF2->F2_VALIRRF != nVlrIr
            cRet += "Divegência entre o valor do IRRF destacado no xml e valor cálculado no Protheus. (F2_VALIRRF)." +" Valor gerado: "+ Transform(SF2->F2_VALIRRF,"@E 999,999,999.99")+". Valor no xml: " + Transform( nVlrIr,"@E 999,999,999.99") + CRLF
        EndIf
        If SFT->FT_VALIRR != nVlrIr
            cRet += "Divegência entre o valor do IRRF destacado no xml e valor cálculado no Protheus. (FT_VALIRR)." +" Valor gerado: "+ Transform(SFT->FT_VALIRR,"@E 999,999,999.99") +". Valor no xml: " + Transform(nVlrIr,"@E 999,999,999.99") + CRLF
        EndIf

        //Validação do Cofins <p1:ValorPis> D2_VALPIS F2_VALPIS FT_VRETPIS
        If SD2->D2_VALCSL != nVlrCsll
            cRet += "Divegência entre o valor do CSLL destacado no xml e valor cálculado no Protheus. (D2_VALCSL)." +" Valor gerado: "+ Transform(SD2->D2_VALCSL,"@E 999,999,999.99") +". Valor no xml: " + Transform(nVlrCsll,"@E 999,999,999.99") + CRLF
        EndIf
        If SF2->F2_VALCSLL != nVlrCsll
            cRet += "Divegência entre o valor do CSLL destacado no xml e valor cálculado no Protheus. (F2_VALCSLL)." +" Valor gerado: "+ Transform(SF2->F2_VALCSLL,"@E 999,999,999.99")+". Valor no xml: " + Transform( nVlrCsll,"@E 999,999,999.99") + CRLF
        EndIf
        If SFT->FT_VRETCSL != nVlrCsll
            cRet += "Divegência entre o valor do CSLL destacado no xml e valor cálculado no Protheus. (FT_VRETCSL)." +" Valor gerado: "+ Transform(SFT->FT_VRETCSL,"@E 999,999,999.99")+". Valor no xml: " + Transform( nVlrCsll,"@E 999,999,999.99") + CRLF
        EndIf
        
        //Validação do ISS <p1:ValorPis> D2_VALPIS F2_VALPIS FT_VRETPIS
        If !lForceProcess .AND. SD2->D2_VALISS != nVlrIss 
            cRet += "Divegência entre o valor do ISS destacado no xml e valor cálculado no Protheus. (D2_VALISS)." +" Valor gerado: "+ Transform(SD2->D2_VALISS,"@E 999,999,999.99") +". Valor no xml: " + Transform(nVlrIss,"@E 999,999,999.99") + CRLF
        EndIf

        If nVlrIssRe > 0
            If SED->(DbSeek(xFilial("SED") + PADR(_cNatureza,LEN(SED->ED_CODIGO))))
                If SED->ED_CALCISS != 'S'
                    cRet += "Natureza: "+_cNatureza+" Filial: "+cFilial+".O campo ED_CALCISS não pode ser DIFERENTE DE S. "
                EndIf
            Else
                cRet += "Natureza não encontrada. "
            EndIf
            
            If SF4->(DbSeek(xFilial("SF4") + SD2->D2_TES))   
                If SF4->F4_ISS != 'S'
                    cRet += "O campo F4_ISS não pode ser diferente de S "
                EndIf
                If !(SF4->F4_LFISS $ 'T,I,O')
                    cRet += "O campo F4_LFISS deve estar entre T I O "
                EndIf
            EndIf
        EndIf
    EndIf

    SF2->(RestArea(aAreaSF2))
    SFT->(RestArea(aAreaSFT))
    SF3->(RestArea(aAreaSF3))
    SD2->(RestArea(aAreaSD2))
    SED->(RestArea(aAreaSED))

Return cRet

Static Function RetConta()
    Local aContas   := {}
    Local cRet      := Alltrim(SuperGetMv("MS_CTANFSE", .f., ''))

Return cRet


User Function LAUF009C
    Local cQuery := ""
    // If Empty(FunName())
    //     PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01RN0001'
    // EndIf

    cQuery := " SELECT R_E_C_N_O_ AS RECNO FROM "  + RetSqlName('ZA5') + " ZA5"
    cQuery += " WHERE ZA5.D_E_L_E_T_ <> '*' AND ZA5_STATUS = 'A' "
    
    If Select('QRYAUX') > 0
        QRYAUX->(dbclosearea())
    EndIf
    
    TcQuery cQuery New Alias 'QRYAUX'
    
    While QRYAUX->(!Eof())
        U_LAUF0009(QRYAUX->RECNO,3)
        QRYAUX->(dbSkip())
    EndDo
    

Return



Static Function AtuZA5()
    Local cAviso    := ''
    Local cErro     := ''
    Local oPedido	:=	TNYXPEDIDOS():New() //Classe da Newib agisrn
    Local cRet      := ""
    Local i := 1
    cXml := ZA5->ZA5_BODY
    oNFSE := xmlParser(cXml, "_", @cErro, @cAviso)  
    
    aVars := {}
    aAdd(aVars, {"Numero NFse"      , "cYNumNF"     , WSAdvValue(oNFSE,"_P_CancelarNfseEnvio:_PEDIDO:_P1_InfPedidoCancelamento:_P1_IdentificacaoNfse:_P1_NUMERO:TEXT","string")})
    aAdd(aVars, {"CNPJ"             , "cYCnpj"      , WSAdvValue(oNFSE,"_P_CancelarNfseEnvio:_PEDIDO:_P1_InfPedidoCancelamento:_P1_IdentificacaoNfse:_P1_CNPJ:TEXT","string")})
    aAdd(aVars, {"dData"             , "dData"      , WSAdvValue(oNFSE,"_P_CancelarNfseEnvio:_PEDIDO:_P1_InfPedidoCancelamento:_P1_IdentificacaoNfse:_P1_DataHora:TEXT","string")})
    
    //Valida se conseguiu recuperar as variaveis
    For i := 1 to Len(aVars)
        &(aVars[i][2]) := aVars[i][3]
        If ValType(&(aVars[i][2])) == 'U'
            Return "Tag " + aVars[i][1] + " não encontrada no xml do NFSE"
        EndIf
    Next
    
    If !setEmpresa(cYCnpj)
        return 'Cadastro de filial não localizado com o CNPJ ' + cYCnpj
    EndIf
    //Transforma o NFSE em um objeto
    RecLock('ZA5', .F.)
        ZA5->ZA5_FILNFS := cFilAnt
        //ZA5->ZA5_NUMNF := cYNumNF
        ZA5->ZA5_NUMNF := IIF(LEN(cYNumNf) > 9, RIGHT(cYNumNf, 9), cYNumNf) //pega somente o 9 últimos numeros da nf
        ZA5->ZA5_SERNF := u_getSerie('NUC')
    ZA5->(MsUnLock())
    
    dData := StoD(StrTran(Left(dData,10),'-',''))
    
    If Empty(dData) .AND. Alltrim(ZA5->ZA5_ACAO) == 'cancelar'
        return 'Conteúdo da data de cancelamento esta invalido.'
    EndIf
    
    dDataBase := dData

Return cRet

//-------------------------------------------------------------------
/*/{Protheus.doc} ExcluRegs
Exclui de nota fiscal das tabelas SF3, SFT, SC6, SD2, CT6
@author  Samuel Dantas
@since   11/02/2020
@version version
/*/
//-------------------------------------------------------------------
Static Function ExcluRegs(nRecnoSF2, lForceExclui)
    Local aAreaSF3  := SF3->(GetArea())
    Local aAreaSFT  := SFT->(GetArea())
    Local aAreaSE1  := SE1->(GetArea())
    Local aAreaSD2  := SD2->(GetArea())
    Local aAreaSC5  := SC5->(GetArea())
    Local aAreaSC6  := SC6->(GetArea())
    Local aAreaZA5  := ZA5->(GetArea())
    Local nRecSC5   := 0
    Local nRecSC6   := 0
    Local nRecSF3   := 0
    Local nRecSFT   := 0
    Local nRecSE1   := 0
    Local nRecSD2   := 0
    Local cRetMsg   := ""
    Local lRet      := .T.
    Default nRecnoSF2 := 0
    Default lForceExclui := .F.
    
    ZA5->(DbSetOrder(1))
    SF3->(DbSetOrder(5))
    SFT->(DbSetOrder(1))
    SE1->(DbSetOrder(1))
    SD2->(DbSetOrder(3))
    SC6->(DbSetOrder(4))
    SC5->(DbSetOrder(1))
    BEGIN TRANSACTION
    SF2->(DbGoTo(nRecnoSF2))
    If SF2->(!EoF())
        If Alltrim(ZA5->ZA5_ACAO) == 'excluir'
            If SC6->(DbSeek(xFilial("SC6") + SF2->(F2_DOC+F2_SERIE)))
                If SC5->(DbSeek(xFilial("SC5") + SC6->C6_NUM))
                    nRecSC5 := SC5->(Recno())
                    RecLock('SC5', .F.)
                        SC5->(DbDelete())
                    SC5->(MsUnLock())
                EndIf

                nRecSC6 := SC6->(Recno())
                RecLock('SC6', .F.)
                    SC6->(DbDelete())
                SC6->(MsUnLock())
            EndIf
        EndIf

        //Se nao for exclusao forcada, deleta o SE1
        If ! lForceExclui
            cSerie := u_getSerie('NUC')
            cQrySE1 := " SELECT R_E_C_N_O_ AS RECNOSE1 FROM "  + RetSqlName('SE1') + " SE1"
            cQrySE1 += " WHERE SE1.D_E_L_E_T_ <> '*' AND E1_NUM = '"+SF2->F2_DOC+"' "
            cQrySE1 += " AND E1_PREFIXO = '"+cSerie+"'"
            cQrySE1 += " AND E1_TIPO = 'NFS' AND E1_FILIAL = '"+SF2->F2_FILIAL+"' "
            
            If Select('QRYSE1') > 0
                QRYSE1->(dbclosearea())
            EndIf
            
            TcQuery cQrySE1 New Alias 'QRYSE1'

            if QRYSE1->(!Eof())
                SE1->(DbGoTo(QRYSE1->(RECNOSE1)))
                nRecSE1 := SE1->(Recno())
                cSeekSE1 := SE1->(E1_FILIAL+E1_PREFIXO+E1_NUM+E1_PARCELA)
                SE1->(DbSeek(cSeekSE1))
                While SE1->(!EoF()) .AND. cSeekSE1 == SE1->(E1_FILIAL+E1_PREFIXO+E1_NUM+E1_PARCELA)
                    If Empty(SE1->E1_BAIXA)
                        RecLock('SE1', .F.)
                            SE1->(DbDelete())
                        SE1->(MsUnLock())
                    Else
                        lRet     := .F.
                        cRetMsg += "Titulo foi baixado."
                        DisarmTransaction()
                        Return cRetMsg
                    EndIf
                    SE1->(DbSkip())
                EndDo
            Else
                lRet     := .F.
                cRetMsg += "Titulo não encontrado."
                DisarmTransaction()
                Return cRetMsg
            EndIF
        EndIf

        //Exclui livros fiscais        
        If SF3->(DbSeek(cSeekSF3 := xFilial("SF3") + SF2->(F2_SERIE+F2_DOC)))
            While SF3->(!Eof()) .AND. cSeekSF3 == SF3->(F3_FILIAL + F3_SERIE + F3_NFISCAL)
                nRecSF3 := SF3->(Recno())
                If Alltrim(ZA5->ZA5_ACAO) == 'excluir'
                    RecLock('SF3', .F.)
                        SF3->(DbDelete())
                    SF3->(MsUnLock())
                Else
                    RecLock('SF3', .F.)
                        SF3->F3_CODRSEF = '101'
                    SF3->(MsUnLock())                
                EndIf
                SF3->(DbSkip()) 
            EndDo        
        Else
            lRet     := .F.
            cRetMsg += "Erro ao deletar registros da tabela SF3"
            DisarmTransaction()
            Return cRetMsg
        EndIf

        //Exclui livros fiscais
        If SFT->(DbSeek(cSeekSFT := xFilial("SFT") + "S" + SF2->(F2_SERIE+F2_DOC)))
            While SFT->(!Eof()) .AND. cSeekSFT == SFT->(FT_FILIAL + FT_TIPOMOV + FT_SERIE + FT_NFISCAL)
                nRecSFT :=  SFT->(Recno())
                If Alltrim(ZA5->ZA5_ACAO) == 'excluir'
                    RecLock('SFT', .F.)
                        SFT->(DbDelete())
                    SFT->(MsUnLock())
                EndIf
                SFT->(DbSkip())
            EndDo
        Else
            lRet     := .F.
            cRetMsg += "Erro ao deletar registros da tabela SFT"
            DisarmTransaction()
            Return cRetMsg
        EndIf
        

        If lRet .AND. Alltrim(ZA5->ZA5_ACAO) == 'excluir'
            RecLock('ZB2', .T.)
                ZB2->ZB2_FILIAL := cFilAnt
                ZB2->ZB2_DATA   := Date()
                ZB2->ZB2_HORA   := Time()
                ZB2->ZB2_NUMNF  := SF2->F2_DOC
                ZB2->ZB2_SERIE  := SF2->F2_SERIE
                ZB2->ZB2_RECSE1 := nRecSE1
                ZB2->ZB2_RECSFT := nRecSFT
                ZB2->ZB2_RECSF3 := nRecSF3
                ZB2->ZB2_RECSC5 := nRecSC5
                ZB2->ZB2_RECSC6 := nRecSC6
                ZB2->ZB2_RECSD2 := nRecSD2
                ZB2->ZB2_CHVNUC := ZA5->ZA5_CHAVE
            ZB2->(MsUnLock())
        EndIf

    EndIf
    END TRANSACTION
    SF3->(RestArea(aAreaSF3))
    SFT->(RestArea(aAreaSFT))
    SE1->(RestArea(aAreaSE1))
    SD2->(RestArea(aAreaSD2))
    SC5->(RestArea(aAreaSC5))
    SC6->(RestArea(aAreaSC6))
    
Return cRetMsg


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
