//Bibliotecas
#Include "Protheus.ch"
#Include "TopConn.ch"
 
//Constantes
#Define STR_PULA    Chr(13)+Chr(10)
 
/*/{Protheus.doc} zTstExc1
Função que cria um exemplo de FWMsExcel
@author Atilio
@since 06/08/2016
@version 1.0
    @example
    u_zTstExc1()
/*/
 
User Function zTstExc1()
    Local aArea        := GetArea()
    Local cQuery        := ""
    Local oFWMsExcel
    Local oExcel
    Local cArquivo    := GetTempPath()+'zTstExc1.xml'
 
    //Pegando os dados
    cQuery := " SELECT "                                                    + STR_PULA
    cQuery += "     SB1.B1_COD, "                                            + STR_PULA
    cQuery += "     SB1.B1_DESC, "                                        + STR_PULA
    cQuery += "     SB1.B1_TIPO, "                                        + STR_PULA
    cQuery += "     SBM.BM_GRUPO, "                                        + STR_PULA
    cQuery += "     SBM.BM_DESC, "                                        + STR_PULA
    cQuery += "     SBM.BM_PROORI "                                        + STR_PULA
    cQuery += " FROM "                                                    + STR_PULA
    cQuery += "     "+RetSQLName('SB1')+" SB1 "                            + STR_PULA
    cQuery += "     INNER JOIN "+RetSQLName('SBM')+" SBM ON ( "        + STR_PULA
    cQuery += "         SBM.BM_FILIAL = '"+FWxFilial('SBM')+"' "        + STR_PULA
    cQuery += "         AND SBM.BM_GRUPO = B1_GRUPO "                    + STR_PULA
    cQuery += "         AND SBM.D_E_L_E_T_='' "                            + STR_PULA
    cQuery += "     ) "                                                        + STR_PULA
    cQuery += " WHERE "                                                    + STR_PULA
    cQuery += "     SB1.B1_FILIAL = '"+FWxFilial('SBM')+"' "            + STR_PULA
    cQuery += "     AND SB1.D_E_L_E_T_ = '' "                            + STR_PULA
    cQuery += " ORDER BY "                                                + STR_PULA
    cQuery += "     SB1.B1_COD "                                            + STR_PULA
    TCQuery cQuery New Alias "QRYPRO"
     
    //Criando o objeto que irá gerar o conteúdo do Excel
    oFWMsExcel := FWMSExcel():New()
     
    //Aba 01 - Teste
    oFWMsExcel:AddworkSheet("Aba 1 Teste") //Não utilizar número junto com sinal de menos. Ex.: 1-
        //Criando a Tabela
        oFWMsExcel:AddTable("Aba 1 Teste","Titulo Tabela")
        //Criando Colunas
        oFWMsExcel:AddColumn("Aba 1 Teste","Titulo Tabela","Col1",1,1) //1 = Modo Texto
        oFWMsExcel:AddColumn("Aba 1 Teste","Titulo Tabela","Col2",2,2) //2 = Valor sem R$
        oFWMsExcel:AddColumn("Aba 1 Teste","Titulo Tabela","Col3",3,3) //3 = Valor com R$
        oFWMsExcel:AddColumn("Aba 1 Teste","Titulo Tabela","Col4",1,1)
        //Criando as Linhas
        oFWMsExcel:AddRow("Aba 1 Teste","Titulo Tabela",{11,12,13,sToD('20140317')})
        oFWMsExcel:AddRow("Aba 1 Teste","Titulo Tabela",{21,22,23,sToD('20140217')})
        oFWMsExcel:AddRow("Aba 1 Teste","Titulo Tabela",{31,32,33,sToD('20140117')})
        oFWMsExcel:AddRow("Aba 1 Teste","Titulo Tabela",{41,42,43,sToD('20131217')})
     
        //Criando as Linhas... Enquanto não for fim da query
        While !(QRYPRO->(EoF()))
            oFWMsExcel:AddRow("Aba 2 Produtos","Produtos",{;
                                                                    QRYPRO->B1_COD,;
                                                                    QRYPRO->B1_DESC,;
                                                                    QRYPRO->B1_TIPO,;
                                                                    QRYPRO->BM_GRUPO,;
                                                                    QRYPRO->BM_DESC,;
                                                                    Iif(QRYPRO->BM_PROORI == '0', 'Não Original', 'Original');
            })
         
            //Pulando Registro
            QRYPRO->(DbSkip())
        EndDo
     
    //Ativando o arquivo e gerando o xml
    oFWMsExcel:Activate()
    oFWMsExcel:GetXMLFile(cArquivo)
         
    //Abrindo o excel e abrindo o arquivo xml
    oExcel := MsExcel():New()             //Abre uma nova conexão com Excel
    oExcel:WorkBooks:Open(cArquivo)     //Abre uma planilha
    oExcel:SetVisible(.T.)                 //Visualiza a planilha
    oExcel:Destroy()                        //Encerra o processo do gerenciador de tarefas
     
    QRYPRO->(DbCloseArea())
    RestArea(aArea)
Return
