#include "totvs.ch"
#include "protheus.ch"
#include "TOPCONN.CH"

User Function ZFIN002()
    
    LOCAL cPerg  := "ZFIN002"
    LOCAL cAlias := GetNextAlias()
    PRIVATE aCab   := {}
    PRIVATE aDados := {}

    Pergunte(cPerg,.T.,"Relatório Notas Fiscais de Serviço")

    BEGINSQL Alias cAlias
        SELECT D2_FILIAL, D2_CLIENTE, D2_LOJA,D2_DOC, D2_SERIE, D2_EMISSAO, D2_TOTAL, F3_RECISS, F3_VALICM,
        D2_VALCSL, D2_VALINS, D2_VALIRRF, D2_VALCOF, D2_VALPIS
        FROM %table:SD2% SD2
        INNER JOIN %Table:SF3% SF3 ON D2_FILIAL+D2_DOC+D2_SERIE+D2_EMISSAO = F3_FILIAL+F3_NFISCAL+F3_SERIE+F3_EMISSAO
        WHERE SD2.D_E_L_E_T_ = ''
        AND SF3.D_E_L_E_T_ = ''
        AND SD2.D2_SERIE = 'NUC'
        AND D2_EMISSAO >= %exp:(MV_PAR01)%
        AND D2_EMISSAO <= %exp:(MV_PAR02)%
    ENDSQL

    AADD(aCab, {"Filial"		,"C", 09, 0})
    AADD(aCab, {"Cliente"	    ,"C", 06, 0})
    AADD(aCab, {"Loja"		    ,"C", 02, 0})
    AADD(aCab, {"Num NFS"   	,"C", 09, 0})
    AADD(aCab, {"Serie"		    ,"C", 03, 0})
    AADD(aCab, {"Dt Emissao" 	,"D", 08, 0})
    AADD(aCab, {"Valor"	        ,"N", 18, 2})
    AADD(aCab, {"Recolhe ISS"	,"N", 18, 2})
    AADD(aCab, {"ICMS"          ,"N", 18, 2})
    AADD(aCab, {"CSL"	        ,"N", 18, 2})
    AADD(aCab, {"INSS"	        ,"N", 18, 2})
    AADD(aCab, {"IRRF"	        ,"N", 18, 2})
    AADD(aCab, {"COFINS"        ,"N", 18, 2})
    AADD(aCab, {"PIS"	        ,"N", 18, 2})
    
    

    While !(cAlias)->(EOF())
        AADD( aDados,{D2_FILIAL, D2_CLIENTE, D2_LOJA, D2_DOC, D2_SERIE, D2_EMISSAO, D2_TOTAL, F3_RECISS, F3_VALICM,D2_VALCSL, D2_VALINS, D2_VALIRRF, D2_VALCOF, D2_VALPIS} )
        DBSKIP()
    end

    MsgRun("Favor Aguardar.....", "Exportando os Registros para o Excel",{||DlgToExcel({{"GETDADOS","RELATÓRIO NOTAS FISCAIS DE SERVIÇO",aCab,aDados}})})
    DbCloseArea() 
Return 
