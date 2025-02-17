#INCLUDE 'Rwmake.ch'
#INCLUDE 'Protheus.ch'
#INCLUDE 'TbIconn.ch'
#INCLUDE 'Topconn.ch'
//-------------------------------------------------------------------
/*/{Protheus.doc} Untitled-1
Retorna qual serie deverá ser usada em NFES, por enquanto
apenas uma filial muda de serie todos os anos
Será centralizada neste User Function, pois com a virada da R33, não se utiliza mais StaticCall.
Ela era centralizada na LAUF0009
@author  Jerry Junior
@since   22/02/2022
@version 1.0
/*/
//-------------------------------------------------------------------
User Function getSerie(cSerie, cFilRef)    
    Local cRet
    Default cSerie := ''
    Default cFilRef := cFilAnt
    If Empty(cSerie)
        cRet := Alltrim(SuperGetMv("MS_SERNFSE",.F.,"NUC"))
    Else
        cRet := cSerie
    EndIf
    
    //A serie da filial especifica mudará todos os anos a partir de 2021
    If cFilRef == '01BA0008'
        If DtoS(dDataBase) >= '20210101'
            cRet    := Alltrim(SuperGetMV("MS_SERDANO", .F., "021", cFilRef))
        EndIf
    EndIf

    //A serie da filial especifica mudará todos os anos a partir de 2021, provavalmente será necessário criar um cadastro serie x filial x ano
    If cFilRef == '01SE0023'
        If DtoS(dDataBase) >= '20210301'
            cRet    := Alltrim(SuperGetMV("MS_SERDANO", .F., "021", cFilRef))
        EndIf
    EndIf

    //A serie da filial especifica mudará todos os anos a partir de 2023, provavalmente será necessário criar um cadastro serie x filial x ano
    If cFilRef == '01PA0032'
        If DtoS(dDataBase) >= '20230601'
            cRet    := Alltrim(SuperGetMV("MS_SERDANO", .F., "N23", cFilRef))
        EndIf
    EndIf

    //A serie da filial especifica mudará todos os anos a partir de 2024, provavalmente será necessário criar um cadastro serie x filial x ano
    If cFilRef == '01BA0037'
        If DtoS(dDataBase) >= '20240101'
            cRet    := Alltrim(SuperGetMV("MS_SERDANO", .F., "N24", cFilRef))
        EndIf
    EndIf

Return cRet
