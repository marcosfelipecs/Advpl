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
        SELECT DISTINCT D2_FILIAL, D2_CLIENTE, D2_LOJA,D2_DOC, D2_SERIE, D2_EMISSAO, D2_TOTAL, F3_RECISS, F3_VALICM,
        D2_VALCSL, D2_VALINS, D2_VALIRRF, D2_VALCOF, D2_VALPIS, F3_OBSERV, F3_DTCANC
        FROM %table:SD2% SD2
        INNER JOIN %Table:SF3% SF3 ON D2_FILIAL+D2_DOC+D2_SERIE+D2_EMISSAO = F3_FILIAL+F3_NFISCAL+F3_SERIE+F3_EMISSAO
        WHERE SF3.D_E_L_E_T_ = ''
        AND D2_EMISSAO >= %exp:(MV_PAR01)%
        AND D2_EMISSAO <= %exp:(MV_PAR02)%
        AND SD2.D2_SERIE = 'NUC' OR SD2.D2_SERIE = '021' 
    ENDSQL

    AADD(aCab, {"FILIAL"		,"C", 09, 0})
    AADD(aCab, {"CLIENTE"	    ,"C", 06, 0})
    AADD(aCab, {"LOJA"		    ,"C", 02, 0})
    AADD(aCab, {"NUM NFS"   	,"C", 09, 0})
    AADD(aCab, {"SÉRIE"		    ,"C", 03, 0})
    AADD(aCab, {"DT EMISSÃO" 	,"D", 08, 0})
    AADD(aCab, {"VALOR"	        ,"N", 18, 2})
    AADD(aCab, {"REC ISS"	    ,"N", 18, 2})
    AADD(aCab, {"ICMS"          ,"N", 18, 2})
    AADD(aCab, {"CSL"	        ,"N", 18, 2})
    AADD(aCab, {"INSS"	        ,"N", 18, 2})
    AADD(aCab, {"IRRF"	        ,"N", 18, 2})
    AADD(aCab, {"COFINS"        ,"N", 18, 2})
    AADD(aCab, {"PIS"	        ,"N", 18, 2})
    AADD(aCab, {"OBSERV"	    ,"C", 18, 0})
    AADD(aCab, {"DT CANC"    	,"D", 08, 0})
    
    

    While !(cAlias)->(EOF())
        AADD( aDados,{D2_FILIAL, D2_CLIENTE, D2_LOJA, D2_DOC, D2_SERIE, D2_EMISSAO, D2_TOTAL, F3_RECISS, F3_VALICM,D2_VALCSL, D2_VALINS, D2_VALIRRF, D2_VALCOF, D2_VALPIS, F3_OBSERV, F3_DTCANC} )
        DBSKIP()
    end

    MsgRun("Por favor aguardar...", "Exportando Notas para o Excel",{||DlgToExcel({{"GETDADOS","RELATÓRIO NOTAS FISCAIS DE SERVIÇO",aCab,aDados}})})
    DbCloseArea() 
Return 
