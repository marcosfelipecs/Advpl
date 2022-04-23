#include "totvs.ch"
#include "protheus.ch"
#include "TOPCONN.CH"

//-------------------------------------------------------------------
/*/{Protheus.doc} zRelMF01
Relatório de retenções de notas fiscais de serviço
@author Marcos Felipe
@since 08/03/2021
@version 1.0.0
/*/
//-------------------------------------------------------------------


User Function ZRELMF01()

    
    LOCAL cPerg  := "ZRELMF01"
    LOCAL cAlias := GetNextAlias()
    PRIVATE aCab   := {}
    PRIVATE aDados := {}


    Pergunte(cPerg,.T.,"Relatório Notas Fiscais de Serviço")

    BEGINSQL Alias cAlias
        COLUMN D2_EMISSAO as DATE
        SELECT DISTINCT D2_FILIAL, D2_CLIENTE, D2_LOJA,D2_DOC, D2_SERIE, D2_EMISSAO, D2_TOTAL, F3_RECISS, F3_VALICM,
        D2_VALCSL,D2_VALINS, D2_VALIRRF, D2_VALCOF,D2_VALPIS, F3_OBSERV
        FROM %table:SD2% SD2
        INNER JOIN %table:SF3% SF3 
        ON F3_FILIAL+F3_NFISCAL+F3_SERIE+F3_EMISSAO = D2_FILIAL+D2_DOC+D2_SERIE+D2_EMISSAO
        WHERE SF3.D_E_L_E_T_ = ''
        AND D2_FILIAL >= %exp:(MV_PAR01)%
        AND D2_FILIAL <= %exp:(MV_PAR02)%
        AND D2_EMISSAO >= %exp:(MV_PAR03)%
        AND D2_EMISSAO <= %exp:(MV_PAR04)%
        AND (D2_SERIE IN ('NUC','021','3','RPS'))
        ORDER BY D2_FILIAL,D2_DOC,D2_EMISSAO
    ENDSQL

    AADD(aCab, {"FILIAL"		,"C", 09, 0})
    AADD(aCab, {"CLIENTE"	    ,"C", 06, 0})
    AADD(aCab, {"LOJA"		    ,"C", 02, 0})
    AADD(aCab, {"NUMERO"       	,"C", 09, 0})
    AADD(aCab, {"SÉRIE"		    ,"C", 03, 0})
    AADD(aCab, {"EMISSÃO"   	,"D", 08, 0})
    AADD(aCab, {"VALOR"	        ,"N", 18, 2})
    AADD(aCab, {"RET_ISS"	    ,"N", 18, 2})
    AADD(aCab, {"ICMS"          ,"N", 18, 2})
    AADD(aCab, {"CSL"	        ,"N", 18, 2})
    AADD(aCab, {"INSS"	        ,"N", 18, 2})
    AADD(aCab, {"IRRF"	        ,"N", 18, 2})
    AADD(aCab, {"COFINS"        ,"N", 18, 2})
    AADD(aCab, {"PIS"	        ,"N", 18, 2})
    AADD(aCab, {"OBSERVAÇÃO"    ,"C", 40, 0})
    

    While !(cAlias)->(EOF())
        AADD( aDados,{D2_FILIAL, D2_CLIENTE, D2_LOJA, D2_DOC, D2_SERIE, D2_EMISSAO, D2_TOTAL, F3_RECISS, F3_VALICM, D2_VALCSL, D2_VALINS, D2_VALIRRF, D2_VALCOF, D2_VALPIS,F3_OBSERV,} )
        DBSKIP()
    end

    MsgRun("Por favor aguardar...", "Exportando Notas para o Excel",{||DlgToExcel({{"GETDADOS","RELATÓRIO NOTAS FISCAIS DE SERVIÇO",aCab,aDados}})})
    DbCloseArea() 

Return 
