#INCLUDE 'Rwmake.ch'
#INCLUDE 'Protheus.ch'
#INCLUDE 'TbIconn.ch'
#INCLUDE 'Topconn.ch'
#include 'fileio.ch'

//-------------------------------------------------------------------
/*/{Protheus.doc} LAUF010A
Ler csv e gerar fatura 
@author  Samuel Dantas
@since   23/12/2019
@version version
/*/
//-------------------------------------------------------------------
User Function LAUF0010
    Local cFile         := ""
    Local aResponse     := {}
    Local aParamBox     := {}
    Local cSelTomador   := '' // Tomador selecionado
    If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01RN0001'
    EndIf
    
    
    Private aRet          := {}
    Private aRecnos     := {}
    Private cMsgErro    := ""
    Private nVlrFat     := 0
    Private cCliente    := ""
    Private cLoja       := ""
    Private cNatureza   := Alltrim(SuperGetMV("MS_NATFATU", .F., "110112"))
    Private dVencto     := dDataBase

    // cFile := cGetFile('Arquivo CSV', 'Selecione a arquivo csv',,, .F., 16)        
    SA1->(DbSetOrder(1))
    SED->(DbSetOrder(1))
    aAdd(aParamBox,{6,"Buscar arquivo",Space(254),"","","",50,.F.,"Todos os arquivos (*.csv) |*.csv"})
    aAdd(aParamBox,{1,"Data Venc."  ,Ctod(Space(8)),"","","","",50,.F.})
    aAdd(aParamBox, {1,"Tomador " 	      	, Space(TAMSX3("A1_COD")[1])	, '','' ,'SA1','.T.',50,.F. })
    aAdd(aParamBox, {1,"Loja " 	      	, Space(TAMSX3("A1_LOJA")[1])	, '','U_L10VLSA1()' ,'','.T.',50,.F. })
    
    If ParamBox(aParamBox,"Importão de fatura",@aResponse)
        cFile       := AllTrim(aResponse[1])
        dVencto       := aResponse[2]
        If !Empty(cFile)
            MSAguarde( { || aRet := LerCsv(cFile) }, "Leitura de CSV" ,"Processando... Aguarde a leitura do CSV.", .T.)    
        EndIf
        If len(aRet) > 1
            if !Empty(MV_PAR03) .AND. !Empty(MV_PAR04)
                cSelTomador := SA1->A1_CGC
            endIf
            MSAguarde( { || Processar(aRet, cSelTomador) }, "Gerando fatura" ,"Processando...", .T.)    
        EndIf
    Endif

Return

Static Function Processar(aRet, cSelTomador)
    Local nI := 1
    Local aRetorno := {}
    Local aFatCGCs := {}
    Local nJ     := 1
    Local nPosCnpj     := 1
    Local cCNPJ     := ''
    default cSelTomador := ''

    aCampos := aRet[1]
    nFilial     := 1
    nNum        := 2
    nSerie      := 3
    nTomador    := 4
    nTipo       := 5
    
    //Divide Por CNPJ
    For nI := 2 To len(aRet)
        cCNPJ := IIF(!Empty(cSelTomador) .AND. !Empty(aRet[nI][1]), cSelTomador, aRet[nI][nTomador])
        If !Empty(cCNPJ)
            If BuscaTomador(cCNPJ)
                if len(aFatCGCs) == 0
                    aAdd(aFatCGCs, {cCNPJ, {aRet[nI]}})
                else
                    nPosCnpj := aScan(aFatCGCs,{|x| AllTrim(x[1]) == Alltrim(cCNPJ)})
                    if nPosCnpj > 0
                        aAdd(aFatCGCs[nPosCnpj][2], aRet[nI])
                    else
                        aAdd(aFatCGCs, {cCNPJ, {aRet[nI]}})
                    endIf
                    
                endIf
            Else
                cMsgErro := "Cliente  de CNPJ "+cCNPJ+" NÃO ENCONTRADO."+ CRLF    
                ApMsgInfo(cMsgErro)
                return
            EndIf
        endIf
    Next
    _cDir := ''
    for nJ := 1 to len(aFatCGCs)
        aRet := aFatCGCs[nJ][2]
        cCNPJ := aFatCGCs[nJ][1]
        //lote da fatura par filtro
        cNumFat := GetNumFat()
        ConfirmSX8()
        cMsgErro := ''
        For nI := 2 To len(aRet)
            _cFilial    := aRet[nI][nFilial]
            _cNum       := aRet[nI][nNum]
            _cPrefixo   := aRet[nI][nSerie]
            // cCNPJ := aRet[nI][nTomador]
            cTipo := aRet[nI][nTipo]
            If Empty(_cFilial) .AND. Empty(_cNum) .AND. Empty(_cPrefixo)
            
            ElseIf BuscaTomador(cCNPJ)
                cCliLoja := SA1->(A1_COD + A1_LOJA)
                aRetorno := AddRecno(_cFilial,_cNum, _cPrefixo, nI, cTipo)
                If !aRetorno[1]
                    cMsgErro += aRetorno[2]
                EndIf
            Else
                cMsgErro := "Cliente  de CNPJ "+cCNPJ+" NÃO ENCONTRADO."+ CRLF    
            EndIf
            
        Next

        If len(aRecnos) > 0 .AND. Empty(cMsgErro)
            cCodZA7 := GetSxeNum('ZA7', 'ZA7_CODIGO')
            ConfirmSX8()
            
            cFiltro := " E1_YFATSEQ = '"+cNumFat+"' "
            
            RecLock('ZA7', .T.)
                ZA7->ZA7_FILIAL := xFilial('ZA7')
                ZA7->ZA7_CODIGO	:= cCodZA7
                ZA7->ZA7_DATA	:= dDataBase
                ZA7->ZA7_HORA	:= Time()				
                ZA7->ZA7_STATUS	:= "P"
            ZA7->(MsUnLock())
            
            lRet := liquidar(cFiltro, StrTran(cCliLoja, '/',''), nVlrFat, cCodZA7)

            If lRet
                RecLock('ZA7', .F.)
                    ZA7->ZA7_STATUS := 'G'
                ZA7->(MsUnLock())
                ApMsgInfo("Fatura incluída com sucesso.")
            Endif

        ElseIf !Empty(cMsgErro)
            ApMsgInfo("Não foi possivel gerar a fatura. Selecione uma pasta para salvar arquivo com log de erro.")
            if Empty(_cDir)
                _cDir := cGetFile("Todos | *.* ", OemToAnsi("Selecione o diretorio"),  ,         ,.F.,GETF_LOCALHARD + GETF_RETDIRECTORY, .F.)    
            endIf
            cFileName := StrTran(cCNPJ,'.','')
            cFileName := StrTran(cFileName,'-','')
            cFileName := StrTran(cFileName,'/','')
            cArqFim := "logfat_"+StrTran(cFileName,'.','')+".txt"
            MemoWrit( _cDir  +  cArqFim, cMsgErro  )
        EndIf
    next

Return

//-------------------------------------------------------------------
/*/{Protheus.doc} liquidar
Chama funcao que faz a liquidacao(geracao da fatura)
@author  Sidney Sales
@since   05/12/19
@version 1.0
/*/
//-------------------------------------------------------------------
Static Function liquidar(cFiltro, cCliLoja, nVlrFat, cCodZA7)
	Local aItens := {}
	Local aCabec := {}
	Local nMoeda := 1
	Local lRet   := .T.

	Local cPrefixo 	:= Alltrim(SuperGetMV('MS_PREFFAT', .F. ,'FAT'))
	Local cTipo		:= Alltrim(SuperGetMV('MS_TIPOFAT', .F. ,'FT' ))

	Private lMsErroAuto := .F.
	
    cTipo	:= 'FT'	
	SA1->(DbSetOrder(1))
    //Posiciona o cliente
	SA1->(DbSeek(xFilial('SA1') + cCliLoja))
	
	//Pega o número do titulo
	cNumSE1	:= GetSxeNum('SE1','E1_NUM')
	ConfirmSX8()
	
	//Fatura que sera gerada
	Aadd(aItens,{;
		{"E1_PREFIXO"   , cPrefixo     },;
		{'E1_NUM' 		, cNumSE1      },;
		{'E1_PARCELA' 	, '001'  	   },;
		{'E1_VENCTO' 	, dVencto      },;
		{'E1_VLCRUZ' 	, nVlrFat      }})

	//Cabecalho da fatura
   	aCabec :={  {"cCondicao"    ,""   	  		},;
                {"cNatureza"    ,cNatureza   	},;
                {"E1_TIPO"      ,cTipo	 	 	},;
                {"cCliente"     ,SA1->A1_COD    },;
                {"nMoeda"       ,nMoeda     	},;
                {"FO0_YZA7"     ,cCodZA7     	},;
                {"cLoja"        ,SA1->A1_LOJA   }}
	BEGIN TRANSACTION
	//Executa a rotina automatica
	MSExecAuTo({|x,y,k,w,z| Fina460(x,y,k,w,z)}, ,aCabec , aItens , 3, cFiltro)  	
    
        //Verifica se houveram erros
        If lMsErroAuto
            MostraErro()
            lRet := .F.
        Else
            If !VldValor(cNumFat)
                ApMsgInfo("Houve divergência nos valores gerados por essa fatura. Contate um analista.")
                DisarmTransaction()
                Return .F.
            EndIf
            PopYNLiq(FO0->FO0_NUMLIQ)
            u_LAUA004Z(FO0->FO0_NUMLIQ)
            
        EndIf
    END TRANSACTION
Return lRet


Static Function AddRecno(_cFilial,_cNum, _cPrefixo, nLinha, cTipo)
    Local lRet := .F.
    Local cMsg := " - Nota ou título não encontrado para a chave "+_cFilial+" - "+_cNum+"/"+_cPrefixo+cTipo+"."+ CRLF
    Local nPosRec := 0
    Default _cFilial    := ""
    Default _cNum       := ""
    Default _cPrefixo   := ""

    cQuery := " SELECT SE1.R_E_C_N_O_ AS RECNOSE1 FROM "  + RetSqlName('SE1') + " SE1 "
    cQuery += " WHERE SE1.D_E_L_E_T_ <> '*' AND E1_NUMLIQ = '' "
    cQuery += " AND E1_FILIAL = '" + PADR(_cFilial,LEN(SE1->E1_FILIAL))+"' "
    cQuery += " AND E1_NUM = '" + PADL(Alltrim(_cNum),LEN(SE1->E1_NUM), "0")+"' "
    cQuery += " AND E1_PREFIXO = '" + PADR(Alltrim(_cPrefixo),LEN(SE1->E1_PREFIXO), "")+"' "
    cQuery += " AND E1_TIPO = '" + PADR(Alltrim(cTipo),LEN(SE1->E1_TIPO), "")+"' "
    // If Alltrim(_cPrefixo) != 'TCK'
    //     cQuery += " AND E1_TIPO = '" + Iif(Alltrim(_cPrefixo) == 'NUC', 'NFS', 'CTE' )+"' "
    // Else
    //     cQuery += " AND E1_TIPO <> 'FT ' "
    // EndIf
    // cQuery += " AND E1_CLIENTE = '" +SA1->(A1_COD)+"' "
    // cQuery += " AND E1_LOJA = '" +SA1->(A1_LOJA)+"' "
    
    If Select('QRY') > 0
        QRY->(dbclosearea())
    EndIf
    
    TcQuery cQuery New Alias 'QRY'
    
    If QRY->(!Eof())
        SE1->(DbGoTo(QRY->RECNOSE1))
        If SE1->E1_SALDO > 0 .AND. Empty(SE1->E1_BAIXA)
            
            nPosRec := Iif (len(aRecnos) == 0, 0, aScan(aRecnos,{|x| x == QRY->RECNOSE1}) )
            
            If nPosRec < 1 
                aAdd(aRecnos,QRY->RECNOSE1)
                nVlrFat += SE1->E1_SALDO - RetImp()
                RecLock('SE1', .F.)
                    If SE1->E1_VENCTO <= Date()	
                        SE1->E1_VENCTO 	:= SE1->E1_VENCTO 	+ 365
                        SE1->E1_VENCREA := SE1->E1_VENCREA 	+ 365
					EndIf
                    SE1->E1_YFATSEQ := cNumFat
                SE1->(MsUnLock())
            Endif
            lRet := .T.
            
        Else
            cMsg := " Nota ou título "+_cFilial+" - "+_cNum+"/"+_cPrefixo+cTipo+" está baixado ou em liquidação("+SE1->E1_YDOCFAT+")."+ CRLF
        EndIf
    EndIf
    
Return {lRet,cMsg}

//-------------------------------------------------------------------
/*/{Protheus.doc} LerCsv
Função responsável por fazer a leitura do arquivo CSV. Para importação.
@author  Samuel Dantas 
@since   10/09/2019
@version version
/*/
//-------------------------------------------------------------------
Static Function LerCsv(cFile)
	// *-------------------------*
	Local cLinha  	:= ""
	Local nLin    	:= 1
	Local aDados  	:= {}
	Local nHandle   := 0 
	Local cTexto    := ""
    
    Default cFile   := ""
	
    Private nLinTit := 0
	Private aRet    := {}
       
    //cTexto := MemoRead(cFile)
    //aDados := StrTokArr2( cTexto, CRLF )

    oFile := FWFileReader():New(cFile)
    If !oFile:Open()
        MsgStop("Erro na abertura do arquivo.")
        Return
    Endif
    cLine := oFile:GetLine()
    While !(Empty(cLine))
        aAux := StrTokArr2( cLine, ";" ,.T.)
        aAdd(aDados, aAux)
        cLine := oFile:GetLine()
    Enddo

    oFile:Close()   

Return aDados



//-------------------------------------------------------------------
/*/{Protheus.doc} BuscaTomador
description
@author  Samuel Dantas
@since   date
@version version
/*/
//-------------------------------------------------------------------
Static Function BuscaTomador(cCNPJ)
    Local lRet := .F.
    Default cCNPJ := ""

    SA1->(DbSetOrder(3))
    cCNPJ := StrTran(cCNPJ,"/","")
    cCNPJ := StrTran(cCNPJ,"\","")
    cCNPJ := StrTran(cCNPJ,".","")
    cCNPJ := StrTran(cCNPJ,",","")
    cCNPJ := StrTran(cCNPJ,"-","")
    If SA1->(DbSeek(xFilial("SA1") + cCNPJ ))
        lRet := .T.
    EndIf

Return lRet



User Function LAUF010C
    Private _cDir := ""
    _cDir := cGetFile("Todos | *.* ", OemToAnsi("Selecione o diretorio"),  ,         ,.F.,GETF_LOCALHARD + GETF_RETDIRECTORY, .F.)    
    If !Empty(_cDir)
        MSAguarde( { || ExporCSV() }, "Fatura CSV" ,"Processando...",.F.)
    EndIf
Return

//-------------------------------------------------------------------
/*/{Protheus.doc} LAUF010C
description
@author  author
@since   date
@version version
/*/
//-------------------------------------------------------------------
Static Function ExporCSV()
    Local aChaves := {}
    Local cQuery := ""
    Local aAreaSE1 := SE1->(GetArea())
    Local aRet := {}
    Local cNomRem   := ""
    Local cNomDest  := ""
    Local cTexto  := ""
    Local nI := 1

    aChaves := GetTitFat()
    cTexto += "FILIAL"+";"+"DOCUMENTO"+";"+"SERIE"+";"+"TIPO"+";"+"TOMADOR"+";"+"LOJA TOMADOR"+";"+"FRETE"+";"+"CNPJ REM"+";"+"LOJA REM"+";"+"NOME REM"+";"+"CNPJ DEST"+";"+"LOJA DEST"+";"+"NOME DEST"+";"+"EMISSÃO"+";"+"VENCIMENTO"+";"+"ICMS"+";"+"VALOR BRUT"+";"+"ISS (-)"+";"+"VALOR LIQ"+CRLF
    SA1->(DbSeek(xFilial("SA1") + SE1->( SE1->(E1_CLIENTE + E1_LOJA))))
    _cCgc := Alltrim(_cCgc)
    For nI := 1 To len(aChaves)
        SE1->(dbSeek(aChaves[nI]))
        PosiTabs()
        
        cFrete := IIF(SC5->C5_TPFRETE == 'C','CIF','FOB')
        If SE1->E1_TIPO == 'CTE'
            SA1->(dbSetOrder(1))
            SA1->(DbSeek(xFilial("SA1") + DT6->( DT6->(DT6_CLIREM + DT6_LOJREM))))
            cNomRem := Transform( SA1->A1_CGC, "@R 99.999.999/9999-99" )+";"+SA1->A1_LOJA+";"+SA1->A1_NOME
            SA1->(DbSeek(xFilial("SA1") + DT6->( DT6->(DT6_CLIDES + DT6_LOJDES))))
            cNomDest := Transform( SA1->A1_CGC, "@R 99.999.999/9999-99" )+";"+SA1->A1_LOJA+";"+SA1->A1_NOME
        Else
            cFrete := ""
            If PosiZA5()
                SA1->(dbSetOrder(1))
                SA1->(DbSeek(xFilial("SA1") + ZA5->( ZA5->(ZA5_REM + ZA5_REMLOJ))))
                cNomRem := Transform( SA1->A1_CGC, "@R 99.999.999/9999-99" )+";"+SA1->A1_LOJA+";"+SA1->A1_NOME
                SA1->(DbSeek(xFilial("SA1") + ZA5->( ZA5->(ZA5_DEST + ZA5_DESLOJ))))
                cNomDest := Transform( SA1->A1_CGC, "@R 99.999.999/9999-99" )+";"+SA1->A1_LOJA+";"+SA1->A1_NOME
            EndIf
        EndIf
        cTexto += Alltrim(SE1->E1_FILIAL)+";"+Alltrim(SE1->E1_NUM)+";"+Alltrim(SE1->E1_SERIE)+";"+SE1->E1_TIPO+";"+Transform( _cCgc, "@R 99.999.999/9999-99" )+";"+Alltrim(SE1->E1_LOJA)+";"+cFrete+";"+Alltrim(cNomRem)+";"+Alltrim(cNomDest)+";"+DtoC(SE1->E1_EMISSAO)+";"+DtoC(SE1->E1_VENCTO)+";"+Transform(SF2->F2_VALICM,"@E 999,999.99")+";"+Transform(SE1->E1_VALOR,"@E 99,999,999.99")+";"+Transform(SF2->F2_VALISS,"@E 999,999.99")+";"+Transform(SE1->E1_VALOR - SF2->F2_VALISS,"@E 999,999.99") + CRLF 
    Next

    SE1->(RestArea(aAreaSE1))
    
    MemoWrit( _cDir  +  "FaturaCSV.csv", cTexto  )

    ApMsgInfo("Fatura salva com sucesso.")
Return 

//-------------------------------------------------------------------
/*/{Protheus.doc} PosiTabs
Posiciona tabelas a partir da SE1
@author  Samuel Dantas
@since   20/12/2019
@version version
/*/
//-------------------------------------------------------------------
Static Function PosiTabs()
  Local cQuery := ""
  Local lRet   := .F.

  cQuery := " SELECT SC5.R_E_C_N_O_ AS RECNOSC5, SF2.R_E_C_N_O_ AS RECNOSF2, DT6.R_E_C_N_O_ AS RECNODT6 FROM "  + RetSqlName('SC5') + " SC5"
  cQuery += " INNER JOIN  "+RetSqlName('SF2')+" SF2 ON F2_FILIAL = C5_FILIAL AND F2_DOC = C5_NOTA AND F2_SERIE = C5_SERIE AND F2_CLIENTE = C5_CLIENTE AND F2_LOJA = C5_LOJACLI AND SF2.D_E_L_E_T_ = SC5.D_E_L_E_T_ "
  cQuery += " INNER JOIN  "+RetSqlName('DT6')+" DT6 ON F2_FILIAL = DT6_FILIAL AND DT6_CHVCTE = F2_CHVNFE AND SF2.D_E_L_E_T_ = DT6.D_E_L_E_T_ "
  cQuery += " WHERE SC5.D_E_L_E_T_ <> '*'"
  cQuery += " AND C5_FILIAL = '"+SE1->E1_FILIAL+"' "
  cQuery += " AND C5_NUM = '"+SE1->E1_PEDIDO+"' AND C5_FILIAL = '"+SE1->E1_FILIAL+"' "
  
  If Select('QRY') > 0
    QRY->(dbclosearea())
  EndIf
  
  TcQuery cQuery New Alias 'QRY'
  
  If QRY->(!Eof())
    SF2->(DbGoTo(QRY->RECNOSF2))
    SC5->(DbGoTo(QRY->RECNOSC5))
    DT6->(DbGoTo(QRY->RECNODT6))
    // DT6->(DBOrderNickname("DT6CHVCTE"))
    // DT6->(DbSeek(xFilial("DT6") + SF2->F2_CHVNFE))

    lRet := .T.
  EndIf
  
Return lRet


Static Function GetNFs(cChave)
  Local nQtd := 0
  Local cRet := ""

  cQuery := " SELECT * FROM "  + RetSqlName('ZA8') + " ZA8"
  cQuery += " WHERE ZA8.D_E_L_E_T_ <> '*' "
  cQuery += " AND ZA8_CHVCTE = '"+cChave+"' "
  
  If Select('QRY') > 0
    QRY->(dbclosearea())
  EndIf
  
  TcQuery cQuery New Alias 'QRY'
  
  While QRY->(!Eof())
    nQtd += 1
    cRet := QRY->ZA8_CHVNFE
    QRY->(dbSkip())
  EndDo
  If !Empty(cRet)
    cRet := sUBstR(cRet,26,9)
    If nQtd >  1
      cRet += "*"
    EndIf
  EndIF
Return cRet

//-------------------------------------------------------------------
/*/{Protheus.doc} GetTitulos
s
@author  Samuel Dantas
@since   15/01/2020
@version 1.0
/*/
//-------------------------------------------------------------------
User Function LAUF010E(cNum)

	Local aJson		:= {}
	Local aAux		:= {}
	Local oJsonFim	:= JsonObject():new()
	Local oCTE	    := JsonObject():new()
	Local oFatura	:= JsonObject():new()
    Local aAreaSE1
    Local nI := 1
    Default cNum    := '003026'
    
    If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01RN0001'
    EndIf

    aAreaSE1 := SE1->(GetArea())    

    SE1->(DbSetOrder(1))
    cQryJson := " SELECT R_E_C_N_O_ AS RECNOSE1 FROM "  + RetSqlName('SE1') + " SE1"
    cQryJson += " WHERE SE1.D_E_L_E_T_ <> '*' AND E1_NUMLIQ = '"+cNum+"' "
    
    If Select('QRYSE1') > 0
        QRYSE1->(dbclosearea())
    EndIf
    
    TcQuery cQryJson New Alias 'QRYSE1'
    
    While QRYSE1->(!Eof())
        SE1->(DbGoTo(QRYSE1->RECNOSE1))
        QRYSE1->(dbSkip())
    EndDo

    If setEmpresa(SE1->E1_FILIAL)
    EndIf

    oJsonFim['filial'] 	            := SE1->E1_FILIAL
    oJsonFim['cnpj_emissor']        := SM0->M0_CGC 
    oJsonFim['numero'] 	            := SE1->E1_NUM
    oJsonFim['serie']	            := SE1->E1_PREFIXO
    oJsonFim['cnpj_cliente']        := POSICIONE("SA1", 1 , xFilial("SA1") +SE1->(E1_CLIENTE + E1_LOJA), "A1_CGC" )
    oJsonFim['tipo']	            := SE1->E1_TIPO
    oJsonFim['emissao']	            := DtoC(SE1->E1_EMISSAO)
    oJsonFim['vencimento']	        := DtoC(SE1->E1_VENCTO)
    oJsonFim['valor']	            := SE1->E1_VALOR
    oJsonFim['nosso_numero']	    := Iif( SE1->E1_PORTADO == "001", SE1->E1_YNN, SE1->E1_NUMBCO)
    oJsonFim['banco']	            := SE1->E1_PORTADO
    oJsonFim['agencia']	            := SE1->E1_AGEDEP
    oJsonFim['conta_corrente']	    := SE1->E1_CONTA
    oJsonFim['ctes']	            := {}    
    oJsonFim['data_cancelamento']   := ""
    
    aChaves := GetTitFat()

    For nI := 1 To len(aChaves)
        
        SE1->(DbSeek(aChaves[nI]))        
        PosiTabs()
        
        If ! setEmpresa(SE1->E1_FILIAL)
        EndIf        
        
        //Se tiver chave, e uma nfe, preenche com ela
        If !Empty(SF2->F2_CHVNFE)
            aAdd(oJsonFim['ctes'], SF2->F2_CHVNFE)
        Else
            //Se nao tiver, e nfse, tenta setar o ZA5 e pegar a chave
            If PosiZA5()
                aAdd(oJsonFim['ctes'], ZA5->ZA5_CHAVE)
            Else
                aAdd(oJsonFim['ctes'], '')
            EndIf
        EndIf

    Next
	
    SE1->(RestArea(aAreaSE1))

Return oJsonFim

Static Function GetTitFat()
  Local aChaves := {}
  Local cQuery := ""

  cQuery := " SELECT DISTINCT FK7_CHAVE FROM "  + RetSqlName('FO0') + " FO0"
  cQuery += " INNER JOIN "+RetSqlName("FO1")+" FO1 ON FO0_FILIAL = FO1_FILIAL AND FO0_PROCES = FO1_PROCES "
  cQuery += " INNER JOIN "+RetSqlName("FK7")+" FK7 ON LEFT(FK7_FILIAL,2) = LEFT(FO1_FILIAL,2) AND FK7_IDDOC = FO1_IDDOC "
  cQuery += " WHERE FO0.D_E_L_E_T_ <> '*' AND FO0_NUMLIQ = '"+SE1->E1_NUMLIQ+"' AND FO0_FILIAL = '"+SE1->E1_FILIAL+"' AND FK7_ALIAS = 'SE1' "
  
  If Select('QRY') > 0
    QRY->(dbclosearea())
  EndIf
  
  TcQuery cQuery New Alias 'QRY'
  
  While QRY->(!Eof())
    aAdd(aChaves, StrTran(QRY->FK7_CHAVE, "|" , "") )
    QRY->(dbSkip())
  EndDo

Return aChaves

//-------------------------------------------------------------------
/*/{Protheus.doc} setEmpresa
Seta a empresa de acordo com o CNPJ
@author  Sidney Sales
@since   11/12/19
@version 1.0
/*/
//-------------------------------------------------------------------
Static Function setEmpresa(_cFilial)
    
    Local lRet := .F.
    
    OpenSM0(cEmpAnt)
    
    SM0->(DbGoTop())

    While SM0->(!Eof())
        If Alltrim(SM0->M0_CODFIL) == _cFilial
            cFilAnt := SM0->M0_CODFIL
            lRet    := .T.
            Exit
        EndIf
        SM0->(DbSkip())
    EndDo

Return lRet

 
//-------------------------------------------------------------------
/*/{Protheus.doc} POPYNLiq
Método para popular dados no campo E1_YNLIQ
@author  Samuel Dantas
@since   18/02/2020
@version version
/*/
//-------------------------------------------------------------------
User Function PopYNLiq(cNumLiq)
    
    Local _cNumLiq := ""
    Local _cDocLiq := ""
    Local _cFilLiq := ""
    Local _cZA7    := ""
    Local cQuery := ""
	Local aAreaSE1 
	
	
	Default cNumLiq := ""

    If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01RN0001'
    EndIf
    aAreaSE1 := SE1->(GetArea())
    cQuery := " SELECT DISTINCT SE1.R_E_C_N_O_ AS RECNOSE1, FO0.R_E_C_N_O_ AS RECNOFO0 FROM "  + RetSqlName('SE1') + " SE1 "
    cQuery += " INNER JOIN "+RetSqlName('FO0')+" FO0 ON FO0_FILIAL = E1_FILIAL AND E1_NUMLIQ = FO0_NUMLIQ AND FO0.D_E_L_E_T_ = SE1.D_E_L_E_T_ "
    cQuery += " WHERE SE1.D_E_L_E_T_ <> '*' AND E1_NUMLIQ <> '' AND E1_EMISSAO = '20200401' "
    
    If Select('QRYLIQ') > 0
        QRYLIQ->(dbclosearea())
    EndIf
    
    TcQuery cQuery New Alias 'QRYLIQ'
    BEGIN TRANSACTION
    SE1->(DbSetOrder(1))
    FO0->(DbSetOrder(1))
    While QRYLIQ->(!Eof())
        SE1->(DbGoTo(QRYLIQ->RECNOSE1))
        FO0->(DbGoTo(QRYLIQ->RECNOFO0))
        _cNumLiq := SE1->E1_NUMLIQ
        _cDocLiq := SE1->E1_NUM
        _cFilLiq := SE1->E1_FILIAL
        RecLock('SE1', .F.)
            SE1->E1_YZA7 := FO0->FO0_YZA7
        SE1->(MsUnLock())
        // aChaves := GetTitFat()
        // For nI := 1 To len(aChaves)
        //     If SE1->(DbSeek(aChaves[nI]))
        //         RecLock('SE1', .F.)
        //             // SE1->E1_YNLIQ := _cNumLiq
        //             SE1->E1_YDOCFAT := _cDocLiq
        //             SE1->E1_YFILFAT := _cFilLiq
        //         SE1->(MsUnLock())
        //     EndIf
        // Next
        QRYLIQ->(dbSkip())
    EndDo
    END TRANSACTION
	SE1->(RestArea(aAreaSE1))
Return        

User Function LAUF010L
    
    If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01RN0001'
    EndIf
    
    nTot := 0
    cQuery := " SELECT A1_COD, A1_LOJA, COUNT(*) AS TOTAL FROM "  + RetSqlName('SA1') + " SA1 "
    cQuery += " INNER JOIN "  + RetSqlName('SF2') + " SF2 ON F2_CLIENTE = A1_COD AND F2_LOJA = A1_LOJA AND SF2.D_E_L_E_T_ = SA1.D_E_L_E_T_ "
    cQuery += " WHERE SA1.D_E_L_E_T_ <> '*' AND A1_YTOMADO <> 'T' GROUP BY A1_COD,A1_LOJA"
    
    If Select('QRY') > 0
        QRY->(dbclosearea())
    EndIf

    cQuery := ChangeQuery(cQuery)
    TcQuery cQuery New Alias 'QRY'
    
    SA1->(DbSetOrder(1))
    While QRY->(!Eof())
        nTot += 1
        // SA1->(DbGoTo(QRY->(RECNO)))
        // GERAITEM()
        QRY->(dbSkip())
    EndDo
    

Return

User Function LAUF010W
    
    If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01RN0001'
    EndIf
    
    
    cQuery := " SELECT DISTINCT SA1.R_E_C_N_O_ AS RECNO FROM "  + RetSqlName('SA1') + " SA1 "
    cQuery += " WHERE SA1.D_E_L_E_T_ <> '*' AND A1_YTOMADO = 'T' "
    
    If Select('QRY') > 0
        QRY->(dbclosearea())
    EndIf
    cQuery := ChangeQuery(cQuery)
    TcQuery cQuery New Alias 'QRY'
    
    SA1->(DbSetOrder(1))
    While QRY->(!Eof())
        SA1->(DbGoTo(QRY->(RECNO)))
        GERAITEM()
        QRY->(dbSkip())
    EndDo
    

Return

Static Function GERAITEM()

    Local cCodItem	:= "C"+SA1->A1_CGC
    Local cDescItem	:= SA1->A1_NOME
    Local cCredDev	:= "2"
    
    If "06626253" == LEFT(SA1->A1_CGC,8)
        cCodItem	:= "PM"+SA1->A1_CGC
    EndIF
   
    If SA1->A1_YTOMADO
        // Inclusao no Item Contabil. //Cliente
        U_fIncCTD(cCodItem, cDescItem, cCredDev)
    EndIf

Return


User Function PopEmail()
    Local aParamBox := {}
    Local aResponse := {}
    Local aRet := {}
    Local nI := 1
    If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01RN0001'
    EndIf
    
    
    cFile       := "C:\temp\cadastro_comercial.csv"
    

    aRet := LerCsv(cFile) 

    SA1->(DbSetOrder(3))
    For nI := 2 To len(aRet)
        cCgc        := Alltrim(aRet[nI][1])
        cCpf        := ""
        cEmailCob   := StrTran(aRet[nI][2], ",",  ";"    )

        aEmailCob   := StrTokArr2(cEmailCob,";")

        cCpf := IIF(len(cCgc) <= 11 , PADL(cCgc,11,"0"), ""  )
        cCgc := IIF( Empty(cCpf) , PADL(cCgc,14,"0"), cCpf  )
        If SA1->(DbSeek(xFilial("SA1") + PADR(cCgc,LEN(SA1->A1_CGC)) ))
            RecLock('SA1', .F.)
                SA1->A1_XEMCOB := cEmailCob
            SA1->(MsUnLock())
            
        EndIf
    Next

    
    
Return


User Function LerCanc()
    Local cErro     := ""
    Local cAviso    := ""
    Local nTotal    := 0
    Local cTexto    := ""
    If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01RN0001'
    EndIf


    cQuery := " SELECT ZA4.R_E_C_N_O_ AS RECNOZA4, SFT.R_E_C_N_O_ AS RECNOSFT, SF3.R_E_C_N_O_ AS RECNOSF3 FROM "  + RetSqlName('ZA4') + " ZA4"
    cQuery += " INNER JOIN "+RetSqlName('SF3')+" SF3 ON F3_CHVNFE = ZA4_CHAVE AND ZA4.D_E_L_E_T_ = SF3.D_E_L_E_T_ "
    cQuery += " INNER JOIN "+RetSqlName('SFT')+" SFT ON FT_CHVNFE = F3_CHVNFE AND SFT.D_E_L_E_T_ = SF3.D_E_L_E_T_ "
    cQuery += " WHERE ZA4.D_E_L_E_T_ <> '*' AND F3_CODRSEF = '101' AND ZA4_ACAO = 'cancelar' AND F3_DTCANC = '' "
    
    If Select('QRY') > 0
        QRY->(dbclosearea())
    EndIf
    
    TcQuery cQuery New Alias 'QRY'
    begin TRANSACTION
    While QRY->(!Eof())
        ZA4->(DbGoTo(QRY->RECNOZA4))
        oCTE := xmlParser(ZA4->ZA4_BODY, "_", @cErro, @cAviso)    
        If ValType(oCTE) != 'U'
            
            cDtcanc := WSAdvValue(oCTE,"_PROCEVENTOCTE:_EVENTOCTE:_infEvento:_dhevento:TEXT","string") 
            If ValType(cDtcanc) == "U"
                cDtcanc := WSAdvValue(oCTE,"_PROCEVENTOCTE:_retEventoCTe:_infEvento:_dhRegEvento:TEXT","string") 
                If ValType(cDtcanc) == "U"
                    QRY->(dbSkip())
                    loop
                EndIf
            EndIf

            dDtCanc := StoD(StrTran(cDtcanc,"-",""))
            SFT->(DbGoTo(QRY->RECNOSFT))
            If SFT->(!Eof())
                RecLock('SFT', .F.)
                    SFT->FT_DTCANC := dDtCanc
                SFT->(MsUnLock())
            EndIf
            SF3->(DbGoTo(QRY->RECNOSF3))
            If SF3->(!Eof())
                RecLock('SF3', .F.)
                    SF3->F3_DTCANC := dDtCanc
                SF3->(MsUnLock())
            EndIf
            
            If Empty(dDtCanc)
            x := 1
            EndIf
            
            cTexto += ZA4->ZA4_CHAVE +CRLF
            nTotal += 1
        EndIf
        
        QRY->(dbSkip())
    EndDo
    x := 1
    END TRANSACTION
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
        ZA5->ZA5_NUMNF := cYNumNF
        ZA5->ZA5_SERNF := "NUC"
    ZA5->(MsUnLock())
    
    dData := StoD(StrTran(Left(dData,10),'-',''))
    
    If Empty(dData) .AND. Alltrim(ZA5->ZA5_ACAO) == 'cancelar'
        return 'Conteúdo da data de cancelamento está vazio.'
    EndIf
    
    dDataBase := dData

Return cRet

//-------------------------------------------------------------------
/*/{Protheus.doc} LAUF010J
Repopula data de cancelamento
@author  Samuel Dantas  
@since   02/03/2020
@version version
/*/
//-------------------------------------------------------------------
User Function LAUF010J
    Local cErro := ""
    Local cAviso := ""
    cTexto := "FILIAL;NUM;SERIE;CLIENTE;LOJA;TIPO"+ CRLF
    If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01RN0001'
    EndIf

    cQry := " SELECT SF3.R_E_C_N_O_ AS RECNOSF3, SE1.R_E_C_N_O_ AS RECNOSE1 FROM SE1010 SE1  "
    cQry += " INNER JOIN SF3010 SF3 ON F3_FILIAL = E1_FILIAL AND E1_NUM = F3_NFISCAL AND E1_PREFIXO = F3_SERIE AND E1_TIPO = F3_ESPECIE AND F3_CLIEFOR = E1_CLIENTE AND F3_LOJA= E1_LOJA AND SF3.D_E_L_E_T_ = SE1.D_E_L_E_T_  "
    // cQry += " INNER JOIN SFT010 SFT ON F3_FILIAL = FT_FILIAL AND F3_CHVNFE = FT_CHVNFE AND SF3.D_E_L_E_T_ = SFT.D_E_L_E_T_ "
    cQry += " WHERE SE1.D_E_L_E_T_ <> '*' AND SF3.D_E_L_E_T_ <> '*'  "
    cQry += " AND E1_PREFIXO = '2' AND E1_TIPO = 'CTE' AND F3_OBSERV = 'NF CANCELADA' and E1_SALDO > 0 AND F3_CHVNFE <> '' AND E1_BAIXA = ''  "

    // cQry := " SELECT SF3.R_E_C_N_O_ AS RECNOSF3, SFT.R_E_C_N_O_ AS RECNOSFT FROM "  + RetSqlName('SF3') + " SF3 "
    // cQry += " INNER JOIN "  + RetSqlName('SFT') + " SFT ON F3_CHVNFE = FT_CHVNFE AND FT_SERIE = F3_SERIE AND SF3.D_E_L_E_T_ = SFT.D_E_L_E_T_ "
    // cQry += " WHERE SF3.D_E_L_E_T_ <> '*' AND F3_OBSERV = 'NF CANCELADA' AND F3_CODRSEF = '101' AND F3_SERIE = '2' AND F3_DTCANC = '' "
    
    If Select('QRY') > 0
        QRY->(dbclosearea())
    EndIf
    
    TcQuery cQry New Alias 'QRY'
    BEGIN TRANSACTION
    ZA4->(DbSetOrder(1))
    While QRY->(!Eof())
        SE1->(DbGoTo(QRY->RECNOSE1))
        SF3->(DbGoTo(QRY->RECNOSF3))
        SFT->(DbGoTo(QRY->RECNOSFT))
        cFilant := SE1->E1_FILIAL 
        cTexto += SE1->(E1_FILIAL+";"+E1_NUM+";"+E1_PREFIXO+";"+E1_CLIENTE+";"+E1_LOJA+";"+E1_TIPO) + CRLF

        If ZA4->(DbSeek(xFilial("ZA4") + SF3->F3_CHVNFE + "cancelar" ))
            oCTE := xmlParser(Alltrim(ZA4->ZA4_BODY), "_", @cErro, @cAviso)    
            
            cDtcanc := WSAdvValue(oCTE,"_PROCEVENTOCTE:_EVENTOCTE:_infEvento:_dhevento:TEXT","string") 
            If ValType(cDtcanc) == "U"
                cDtcanc := WSAdvValue(oCTE,"_retEventoCTe:_infEvento:_dhRegEvento:TEXT","string") 
                If ValType(cDtcanc) == "U"
                    cDtcanc := WSAdvValue(oCTE,"_PROCEVENTOCTE:_retEventoCTe:_infEvento:_dhRegEvento:TEXT","string") 
                EndIf
            EndIf

            If ValType(cDtcanc) == "U"
                cRetorno := "Data do evento não encontrada no XML." + CRLF
                DisarmTransaction()
                return
            EndIf
            
            dDtCanc := StoD(StrTran(cDtcanc,"-",""))
            RecLock('ZA4', .F.)
                ZA4->ZA4_STATUS := 'P'
            ZA4->(MsUnLock())
            
            RecLock('SFT', .F.)
                SFT->FT_DTCANC := dDtCanc
            SFT->(MsUnLock())
            
            RecLock('SF3', .F.)
                SF3->F3_DTCANC := dDtCanc
                SF3->F3_CODRSEF := "101" 
            SF3->(MsUnLock())

            //Exclui DT6
            DT6->(DBOrderNickname("DT6CHVCTE"))
            If DT6->(DbSeek(xFilial("DT6") + SF2->F2_CHVNFE))
                RecLock('DT6', .F.)
                    DT6->(DbDelete())
                DT6->(MsUnLock())
            EndIf
            
            lMsErroAuto := .F.
            aVetor := {}

            // RecLock('SE1', .F.)
            //     SE1->(DbRecall())
            // SE1->(MsUnLock())

            aAdd(aVetor, {"E1_FILIAL"   , SE1->E1_FILIAL                     } )
            aAdd(aVetor, {"E1_NUM"      , SE1->E1_NUM                        } )
            aAdd(aVetor, {"E1_PREFIXO"  , SE1->E1_PREFIXO                      } )
            aAdd(aVetor, {"E1_CLIENTE"  , SE1->E1_CLIENTE                    } )
            aAdd(aVetor, {"E1_LOJA"     , SE1->E1_LOJA                        } )
            aAdd(aVetor, {"E1_TIPO"     , SE1->E1_TIPO} )

            MSExecAuTo({|x,y|FINA040(x,y)},aVetor,5)

            If lMsErroAuto
                DisarmTransaction()
                Return
            EndIf

            RecLock('SF3', .F.)
                SF3->F3_DTCANC :=  dDtCanc
                SF3->F3_CODRSEF :=  '101'
            SF3->(MsUnLock())
            RecLock('SFT', .F.)
                SFT->FT_DTCANC :=  dDtCanc
            SFT->(MsUnLock())
        Else    
            DisarmTransaction()
            Return
        EndIf

        QRY->(dbSkip())
    EndDo

    // MemoWrite("C:\TEMP\resultado.csv", cTexto)
    
    END TRANSACTION
Return

User Function AjustCad
    If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01RN0001'
    EndIf
    cQuery := " SELECT R_E_C_N_O_ AS RECNO FROM "  + RetSqlName('ZA5') + " ZA5"
    cQuery += " WHERE ZA5.D_E_L_E_T_ <> '*'  AND ZA5_STATUS = 'E' "
    
    If Select('QRY') > 0
        QRY->(dbclosearea())
    EndIf
    
    TcQuery cQuery New Alias 'QRY'
    
    While QRY->(!Eof())
        ZA5->(DbGoTo(QRY->RECNO))    
        
        If  "ja cadastrada" $ Alltrim(ZA5->ZA5_ERRO)
            RecLock('ZA5', .F.)
                ZA5->ZA5_STATUS := 'P'
            ZA5->(MsUnLock())
                
        EndIf
        QRY->(dbSkip())
    EndDo
    
Return

//-------------------------------------------------------------------
/*/{Protheus.doc} ValidaLiq
Função que verifica as baixas geradas na liquidação
@author  Samuel Dantas
@since   17/02/2020	
@version version
/*/
//-------------------------------------------------------------------
Static Function VldValor(cNumFat)
	Local cMsg 		:= ""
	Local aAreaSE1 	:= SE1->(GetArea())
	Local aAreaSE5 	:= SE5->(GetArea())
	Local cQryBx 	:= ""
	Local nTotFatura 	:= 0
	Local lRet 	:= .F.
	Default cLiq := ""
	
	cQryBx := " SELECT R_E_C_N_O_ AS RECNOSE1 FROM SE1010 SE1 "
	cQryBx += " WHERE SE1.D_E_L_E_T_ <> '*' AND E1_YFATSEQ = '"+cNumFat+"'  "

	If Select('QRYBX') > 0
		QRYBX->(dbclosearea())
	EndIf

	TcQuery cQryBx New Alias 'QRYBX'

	While QRYBX->(!Eof())
		SE1->(DbGoTo(QRYBX->RECNOSE1))
		If SE1->E1_SALDO != 0
			return .F.
		EndIf
		nTotFatura += ( SE1->E1_VALOR - RetImp() )
		QRYBX->(DbSkip())
	EndDo

	SE1->(DBOrderNickname("FINALIQ"))
	If SE1->(DbSeek(FO0->(FO0_FILIAL+FO0_NUMLIQ)))
		If ROUND(SE1->E1_VALOR,2) == ROUND(nTotFatura, 2)
			lRet := .T.
		EndIf
	EndIf

	SE1->(RestArea(aAreaSE1))
	SE5->(RestArea(aAreaSE5))


Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} PopRemDest
Método responsável pela importação dos dados de remetente e destinatário/
na tabela de integração das NFSE(ZA5)
@author  Samuel Dantas  
@since   02/04/2020
@version version
/*/
//-------------------------------------------------------------------
User Function PopRemDest

    Local cFile         := ""
    Local aRet     := {}
    Local aParamBox     := {}

    Private aRet          := {}
    Private aRecnos     := {}

    
    aAdd(aParamBox,{6,"Buscar arquivo",Space(254),"","","",50,.F.,"Todos os arquivos (*.csv) |*.csv"})
    
    
    If ParamBox(aParamBox,"Importão de fatura",@aRet)
        cFile       := AllTrim(aRet[1])
        If !Empty(cFile)
            MSAguarde( { || aRet := LerCsv(cFile) }, "Leitura de CSV" ,"Processando... Aguarde a leitura do CSV.", .T.)    
        EndIf
        If len(aRet) > 1
            MSAguarde( { || ProcZa5(aRet) }, "Popula table ZA5" ,"Processando...", .T.)    
        EndIf
    Endif

Return


Static Function ProcZa5(aRet)
    Local nI := 2
    Local nPosFil   := 1 //Posição no arquivo csv
    Local nPosNum   := 2 //Posição no arquivo csv
    Local nPosRem   := 3 //Posição no arquivo csv
    Local nPosDes   := 4 //Posição no arquivo csv
    Default aRet := {}

    BEGIN TRANSACTION
        For nI := 2 To len(aRet)
            _cCNPJFil := aRet[nI][nPosFil]
            _cNum     := aRet[nI][nPosNum]
            _cCNPJRem := aRet[nI][nPosRem]
            _cCNPJDes := aRet[nI][nPosDes]

            //Trata string antes de busca a empresa
            _cCNPJFil := StrTran(_cCNPJFil,"/","")
            _cCNPJFil := StrTran(_cCNPJFil,"\","")
            _cCNPJFil := StrTran(_cCNPJFil,".","")
            _cCNPJFil := StrTran(_cCNPJFil,",","")
            _cCNPJFil := StrTran(_cCNPJFil,"-","")

            //Trata string antes de busca a remetente
            _cCNPJRem := StrTran(_cCNPJRem,"/","")
            _cCNPJRem := StrTran(_cCNPJRem,"\","")
            _cCNPJRem := StrTran(_cCNPJRem,".","")
            _cCNPJRem := StrTran(_cCNPJRem,",","")
            _cCNPJRem := StrTran(_cCNPJRem,"-","")

            //Trata string antes de busca a destinatario
            _cCNPJDes := StrTran(_cCNPJDes,"/","")
            _cCNPJDes := StrTran(_cCNPJDes,"\","")
            _cCNPJDes := StrTran(_cCNPJDes,".","")
            _cCNPJDes := StrTran(_cCNPJDes,",","")
            _cCNPJDes := StrTran(_cCNPJDes,"-","")

            If setEmpresa(_cCNPJFil)
                ApMsgInfo("Filial "+_cCNPJFil+" não encontrada.")
                DisarmTransaction()
            EndIf

            //Busca registro na tabela de importação
            cQry := " SELECT R_E_C_N_O_ AS RECNO FROM "  + RetSqlName('ZA5') + " ZA5 "
            cQry += " WHERE ZA5.D_E_L_E_T_ <> '*' AND ZA5_FILNFS = '"+cFilAnt+"' AND ZA5_NUMNF = '"+PADL(Alltrim(_cNum),9,"0")+"' " 
            cQry += " ORDER BY ZA5_DATA,ZA5_HORA DESC "
            
            If Select('QRYZA5') > 0
                QRYZA5->(dbclosearea())
            EndIf
            
            TcQuery cQry New Alias 'QRYZA5'
            
            If QRYZA5->(!Eof())
                ZA5->(DbGoTo(QRYZA5->RECNO))
                
                cCodREM := ""
                cLojREM := ""
                cCodDes := ""
                cLojDes := ""

                If (len(_cCNPJDes) == 11 .OR. len(_cCNPJDES) == 14) .AND. ( len(_cCNPJREM) == 11 .OR. len(_cCNPJREM) == 14 )
                    cCodREM := IIF(len(_cCNPJREM) == 14, SUBSTR(ALLTRIM(_cCNPJREM),1,8), SUBSTR(ALLTRIM(_cCNPJREM),1,9) )
                    cLojREM := IIF(len(_cCNPJREM) == 14, SUBSTR(ALLTRIM(_cCNPJREM),9,4), "0000")
                    cCodDes := IIF(len(_cCNPJDes) == 14, SUBSTR(ALLTRIM(_cCNPJDes),1,8), SUBSTR(ALLTRIM(_cCNPJDes),1,9) )
                    cLojDes := IIF(len(_cCNPJDes) == 14, SUBSTR(ALLTRIM(_cCNPJDes),9,4), "0000")
                    
                    RecLock('ZA5', .F.)
                        ZA5->ZA5_REM    := cCodREM
                        ZA5->ZA5_REMLOJ := cLojREM
                        ZA5->ZA5_DEST   := cCodDes
                        ZA5->ZA5_DESLOJ := cLojDes
                    ZA5->(MsUnLock())
                Else
                    ApMsgInfo("CNPJ/CPF de remetente ou destinatário inválido."+CRLF+"Remetente: "+_cCNPJREM+CRLF+"Destinatário: "+_cCNPJDES)
                    DisarmTransaction()
                EndIf

                QRYZA5->(dbSkip())
            EndIf
            
        Next
    END TRANSACTION
Return


User Function PopSX6
    Local aContas   := {}
    Local cRet      := {}
    Local nI := 1
    If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01RN0001'
    EndIf
    
    
    aAdd( aContas, {"01RN0001", "31100100001" })
    aAdd( aContas, {"01RN0002", "31100100002" })
    aAdd( aContas, {"01PE0003", "31100100003" })
    aAdd( aContas, {"01PB0005", "31100100004" })
    aAdd( aContas, {"01CE0006", "31100100005" })
    aAdd( aContas, {"01AL0007", "31100100006" })
    aAdd( aContas, {"01BA0008", "31100100007" })
    aAdd( aContas, {"01SP0012", "31100100008" })
    aAdd( aContas, {"01PR0013", "31100100009" })
    aAdd( aContas, {"01GO0017", "31100100010" })
    aAdd( aContas, {"01RJ0010", "31100100011" })
    aAdd( aContas, {"01PB0016", "31100100012" })
    aAdd( aContas, {"01MG0018", "31100100013" })
    aAdd( aContas, {"01PB0019", "31100100014" })
    aAdd( aContas, {"01AM0020", "31100100015" })
    aAdd( aContas, {"01RN0015", "31100100016" })
    aAdd( aContas, {"01SE0023", "31100100017" })
    aAdd( aContas, {"01CE0024", "31100100018" })
    aAdd( aContas, {"01PE0025", "31100100019" })
    aAdd( aContas, {"01ES0026", "31100100020" })
    aAdd( aContas, {"01CE0027", "31100100021" })
    aAdd( aContas, {"01RS0028", "31100100022" })
    aAdd( aContas, {"01MA0029", "31100100023" })
    aAdd( aContas, {"01PI0030", "31100100024" })
    aAdd( aContas, {"01PA0032", "31100100025" })
    aAdd( aContas, {"01PE0033", "31100100026" })
    aAdd( aContas, {"01CE0031", "31100100027" })
    aAdd( aContas, {"01GO0022", "31100100028" })
    aAdd( aContas, {"01CE0035", "31100100029" })
    aAdd( aContas, {"01SP0034", "31100100030" })
    aAdd( aContas, {"01AL0036", "31100100031" })

    BEGIN TRANSACTION
    For nI := 1 To len(aContas)
        RecLock('SX6', .T.)
            SX6->X6_FIL := aContas[nI][1]
            SX6->X6_VAR := 'MS_CTACTE'
            SX6->X6_TIPO := 'C'
            SX6->X6_DESCRIC := 'Conta CTE '+ aContas[nI][1]
            SX6->X6_DSCSPA := 'Conta CTE '+ aContas[nI][1]
            SX6->X6_DSCENG := 'Conta CTE '+ aContas[nI][1]
            SX6->X6_DESC1 := 'Conta CTE '+ aContas[nI][1]
            SX6->X6_CONTEUD := aContas[nI][2]
            SX6->X6_CONTSPA := aContas[nI][2]
            SX6->X6_CONTENG := aContas[nI][2]
        SX6->(MsUnLock())
        
    Next
    END TRANSACTION
Return

//-------------------------------------------------------------------
/*/{Protheus.doc} CopXML
Busca XML no servidor do cofre sieg e copia para o server no CLOUD
@author  Samuel Dantas
@since   03/04/2020
@version version
/*/
//-------------------------------------------------------------------
User Function CopXML()
    Local cTexto := ""
    Local nI := 0
    Local aArquivos := {}

    If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01RN0001'
    EndIf
    
    cEmailSup := Alltrim(SuperGetMv("MS_TIMAIL",.F.,"diego.cunha@lauto.com.br"))
    cPathNFE := Alltrim(SuperGetMV("MS_PATHNFE",.F., "C:\sieg_nfexml\"))
    cPathCTE := Alltrim(SuperGetMV("MS_PATHCTE",.F., "C:\sieg_nfexml\"))
    cPathSer := Alltrim(SuperGetMV("MS_PATHSER",.F., "\testexml\" ))
    
    If !ExistDir(cPathNFE+"enviados_"+DtoS(dDataBase)+"")
        MakeDir(cPathNFE+"enviados_"+DtoS(dDataBase)+"")
    EndIf
    If !ExistDir(cPathCTE+"enviados_"+DtoS(dDataBase)+"")
        MakeDir(cPathCTE+"enviados_"+DtoS(dDataBase)+"")
    EndIf
    //Busca os dados do XML NFE do COFRE SIEG
    aArquivos := {}
    aDados := ADir(cPathNFE+"*.xml", @aArquivos)
    If len(aArquivos) < 1 
        cTexto += "Não foram encontrados arquivos para NFE." +CRLF
    EndIf
    //Percorre arquivos e copia para o protheus data
    For nI := 1 To len(aArquivos)
        cFile := aArquivos[nI]
        // ApMsgInfo(cFile)
        If !CpyT2S( cPathNFE+cFile, cPathSer)
            cTexto += cFile+ " - erro ao enviar." +CRLF
        Else
            CpyS2T( cPathSer+cFile, cPathNFE+"enviados_"+DtoS(dDataBase)+"\")
            FErase(cPathNFE+cFile)
        EndIf
    Next

    //Busca os dados do XML CTE do COFRE SIEG
    aArquivos := {}
    aDados := ADir(cPathCTE+"*.xml", @aArquivos)
    //Percorre arquivos e copia para o protheus data
    For nI := 1 To len(aArquivos)
        cFile := aArquivos[nI]
        // ApMsgInfo(cFile)
        If !CpyT2S( cPathCTE+cFile, cPathSer)
            cTexto += cFile+ " - erro ao enviar." +CRLF
        Else
            CpyS2T( cPathSer+cFile, cPathCTE+"enviados_"+DtoS(dDataBase)+"\")
            FErase(cPathCTE+cFile)
        EndIf
    Next

    If len(aArquivos) < 1 
        cTexto += "Não foram encontrados arquivos para CTE." +CRLF
    EndIf

    // Caso tenha algum erro envia email ao suporte
    If !Empty(cTexto)
        U_EnviaEmail("Envio de XML(COFRE SIEG)", cEmailSup, cTexto, "", "", .F.)
    Else
        U_EnviaEmail("Envio de XML(COFRE SIEG)", cEmailSup, "XML Importados.", "", "", .F.)
    EndIf

Return



/*/{Protheus.doc} PopNFat
    (long_description)
    @type  User Function
    @author Samuel Dantas
    @since 11/05/2020
    @version version
/*/
User function PopNFat(param_name)
    Local cQuery := ""
    Local aChaves := {}
    If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01RN0001'
    EndIf
    
    cQry := " SELECT DISTINCT SE1.R_E_C_N_O_ AS RECNOSE1, E5_FILIAL, E5_DOCUMEN FROM "  + RetSqlName('SE1') + " SE1 "
    cQry += " INNER JOIN "  + RetSqlName('SE5') + " SE5 ON E1_FILIAL = E5_FILORIG AND E5_NUMERO = E1_NUM AND E1_PREFIXO = E5_PREFIXO AND E5_CLIFOR = E1_CLIENTE AND E5_LOJA = E1_LOJA AND E1_TIPO = E5_TIPO "
    cQry += " WHERE SE1.D_E_L_E_T_ <> '*' AND E5_TIPODOC = 'BA' AND E5_PREFIXO = 'TCK' AND E1_EMISSAO < '20200101' AND E1_YDOCFAT = '' AND E1_BAIXA <> '' GROUP BY E5_FILIAL, E5_DOCUMEN, SE1.R_E_C_N_O_ "
    
    If Select('QRYFAT') > 0
        QRYFAT->(dbclosearea())
    EndIf
    
    TcQuery cQry New Alias 'QRYFAT'
    BEGIN TRANSACTION 
        While QRYFAT->(!Eof())
            cNumFat := ""
            SE1->(DbOrderNickname( "FINALIQ" ))
            If SE1->(DbSeek(QRYFAT->(E5_FILIAL+PADR(E5_DOCUMEN,LEN(SE1->E1_NUMLIQ)))))
                cNumFat := SE1->E1_NUM
                If QRYFAT->RECNOSE1 > 0
                    SE1->(DbGoTo(QRYFAT->RECNOSE1))
                    If Empty(SE1->E1_YDOCFAT)
                        RecLock('SE1', .F.)
                            SE1->E1_YDOCFAT := cNumFat
                        SE1->(MsUnLock())
                    ENDIF
                EndIf
            EndIf

            QRYFAT->(dbSkip())
        EndDo
    END TRANSACTION 

Return 


USer Function popza()
    Local aAreaSA1      
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
    
    If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01RN0001'
    EndIf
    
    
    aAreaSA1      := SA1->(GetArea())
    cQuery := " SELECT R_E_C_N_O_ AS RECNO FROM "  + RetSqlName('ZA4') + " ZA4"
    cQuery += " WHERE ZA4.D_E_L_E_T_ <> '*' AND ZA4_ACAO IN ('cancelar','inutilizar') AND ZA4_NUMCTE = '' "
    
    If Select('QRY') > 0
        QRY->(dbclosearea())
    EndIf
    
    TcQuery cQuery New Alias 'QRY'
    
    While QRY->(!Eof())

        ZA4->(DbGoTo(QRY->RECNO))

        cAcao := Alltrim(ZA4->ZA4_ACAO)
        cChave := Alltrim(ZA4->ZA4_CHAVE)
        cXml := ZA4->ZA4_BODY

        //Transforma o CTE em um objeto
        oCTE := xmlParser(cXml, "_", @cErro, @cAviso)    
        
        If ValType(oCTE) == 'U'
            return 'Erro no objeto XML: ' + cErro
        EndIf


        If Alltrim(cAcao) == 'cancelar'
    
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

            ZA4->(DbGoTo(QRY->RECNO))
            
            If Alltrim(ZA4->ZA4_ACAO) == 'cancelar'
                RecLock('ZA4', .F.)
                    ZA4->ZA4_FILCTE := cFilant
                    ZA4->ZA4_NUMCTE := PADL(Alltrim(cYNumNF),LEN(ZA4->ZA4_NUMCTE),"0")
                    ZA4->ZA4_SERCTE := cYNumSer
                    ZA4->ZA4_CLIENT := cCodCli
                    ZA4->ZA4_LOJA   := cLojaCli
                    ZA4->ZA4_CNPJ   := cCNPJCli
                ZA4->(MsUnLock())
            EndIf
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

            RecLock('ZA4', .F.)
                ZA4->ZA4_FILCTE := cFilant
                ZA4->ZA4_NUMCTE := PADL(Alltrim(cYNumNF),LEN(ZA4->ZA4_NUMCTE),"0")
                ZA4->ZA4_SERCTE := cSerie
            ZA4->(MsUnLock())
        EndIf

        QRY->(dbSkip())
    EndDo
    
Return


User Function Analise()
    Local nCont := 0
    Local nDif := 0
    Local cErro := ""
    Local cAviso := ""
    Local cTexto := ""
    Local i := 1
    If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01RN0001'
    EndIf


    cQuery := " SELECT ZA5.R_E_C_N_O_ AS RECNOZA5, SD2.R_E_C_N_O_ AS RECNOSD2 FROM "  + RetSqlName('SD2') + " SD2"
    cQuery += " INNER JOIN "  + RetSqlName('ZA5') + " ZA5 ON D2_FILIAL = ZA5_FILNFS AND D2_DOC = ZA5_NUMNF AND D2_SERIE = ZA5_SERNF AND SD2.D_E_L_E_T_ = ZA5.D_E_L_E_T_ "
    cQuery += " WHERE SD2.D_E_L_E_T_ <> '*' AND D2_SERIE = 'NUC' AND ZA5_ACAO = 'inclusao' "
    
    If Select('QRY') > 0
        QRY->(dbclosearea())
    EndIf
    
    TcQuery cQuery New Alias 'QRY'
    
    While QRY->(!Eof())
        SD2->(DbGoTo(QRY->RECNOSD2))
        ZA5->(DbGoTo(QRY->RECNOZA5))

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
        
        aAdd(aVars, {"Valor ISS"        , "nVlrISS"    , WSAdvValue(oNFSE,"_P1_SERVICO:_P1_VALORES:_P1_VALORISS:TEXT","string") })
       

        //Valida se conseguiu recuperar as variaveis
        For i := 1 to Len(aVars)
            &(aVars[i][2]) := aVars[i][3]
            If ValType(&(aVars[i][2])) == 'U'
                Return "Tag " + aVars[i][1] + " não encontrada no xml do CTE"
            EndIf
        Next
        
        nVlrIss     := Val(nVlrIss)
        If SD2->D2_VALISS != nVlrIss
            cTexto += SD2->(D2_FILIAL+";"+D2_DOC+";"+D2_SERIE)+ CRLF
            nDif += 1
        EndIf
        QRY->(dbSkip())
    EndDo
    
Return nCont

//-------------------------------------------------------------------
/*/{Protheus.doc} GetNumFat
Busca o ultimo numero disponivel para gerar numero de fatura disponivel
@author  author
@since   date
@version version
/*/
//-------------------------------------------------------------------
Static Function GetNumFat()
	Local cQry := ""
	Local cRet := ""

	cQry := " SELECT MAX(E1_YFATSEQ) as MAX FROM "  + RetSqlName('SE1') + " SE1"
	cQry += " WHERE SE1.D_E_L_E_T_ <> '*' "
	
	If Select('QRYMAX') > 0
		QRYMAX->(dbclosearea())
	EndIf
	
	TcQuery cQry New Alias 'QRYMAX'
	
	If QRYMAX->(!Eof())
		If Empty(QRYMAX->MAX)
			cRet := StrZero(1,TAMSX3("E1_YFATSEQ")[1])
		Else	
			cRet := Soma1(QRYMAX->MAX)
		EndIf
		QRYMAX->(dbSkip())
	Else 
		cRet := StrZero(1,TAMSX3("E1_YFATSEQ")[1])
	EndIf
	 
Return cRet

//-------------------------------------------------------------------
/*/{Protheus.doc} RetImp
Retorna impostos
@author  Samuel dantas
@since   17/03/2020
@version version
/*/
//-------------------------------------------------------------------
Static Function RetImp
	Local cQry := ""
	Local nTotal := 0

	cQry := " SELECT SUM(E1_VALOR) AS TOTIMP FROM "  + RetSqlName('SE1') + " SE1"
	cQry += " WHERE SE1.D_E_L_E_T_ <> '*' AND E1_FILIAL = '"+SE1->E1_FILIAL+"' AND E1_NUM = '"+SE1->E1_NUM+"' AND E1_PREFIXO = '"+SE1->E1_PREFIXO+"' AND E1_CLIENTE = '"+SE1->E1_CLIENTE+"' AND E1_LOJA = '"+SE1->E1_LOJA+"' "
	cQry += " AND E1_TIPO IN ('AB-','FB-','FC-','FU-','IR-','IN-','IS-','PI-','CF-','CS-','FE-','IV-') "
	
	If Select('QRYIMP') > 0
		QRYIMP->(dbclosearea())
	EndIf
	
	TcQuery cQry New Alias 'QRYIMP'
	
	While QRYIMP->(!Eof())
		nTotal := QRYIMP->TOTIMP
		QRYIMP->(dbSkip())
	EndDo
	
Return nTotal

//-------------------------------------------------------------------
/*/{Protheus.doc} POPYNLiq
Método para popular dados no campo E1_YNLIQ
@author  Samuel Dantas
@since   18/02/2020
@version version
/*/
//-------------------------------------------------------------------
Static Function PopYNLiq(cNumLiq)
    
    Local _cNumLiq := ""
    Local cQuery := ""
	Local aAreaSE1 := SE1->(GetArea())
	Local nI := 1
	
	Default cNumLiq := ""

    If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01RN0001'
    EndIf
    
    cQuery := " SELECT DISTINCT R_E_C_N_O_ AS RECNOSE1 FROM "  + RetSqlName('SE1') + " SE1"
    cQuery += " WHERE SE1.D_E_L_E_T_ <> '*' AND E1_NUMLIQ = '"+cNumLiq+"' "
    
    If Select('QRYLIQ') > 0
        QRYLIQ->(dbclosearea())
    EndIf
    
    TcQuery cQuery New Alias 'QRYLIQ'

    SE1->(DbSetOrder(1))
    While QRYLIQ->(!Eof())
        SE1->(DbGoTo(QRYLIQ->RECNOSE1))
        _cNumLiq := SE1->E1_NUMLIQ
        _cDocLiq := SE1->E1_NUM
        _cFilLiq := SE1->E1_FILIAL

		RecLock('SE1', .F.)
			SE1->E1_YZA7 := FO0->FO0_YZA7
		SE1->(MsUnLock())
		
        aChaves := GetTitFat()
		If ValType(aChaves) <> "A"
			aChaves := {}
		EndIf

        For nI := 1 To len(aChaves)
            SE1->(DbSeek(aChaves[nI]))
            RecLock('SE1', .F.)
                SE1->E1_YNLIQ := _cNumLiq
				SE1->E1_YDOCFAT := _cDocLiq
                SE1->E1_YFILFAT := _cFilLiq
            SE1->(MsUnLock())
        Next
        QRYLIQ->(dbSkip())
    EndDo

	SE1->(RestArea(aAreaSE1))
Return

//-------------------------------------------------------------------
/*/{Protheus.doc} PosiZA5
Posiciona a tabela ZA5 (integração de nfse)
@author  Samuel Dantas
@since   02/04/2020
@version version
/*/
//-------------------------------------------------------------------
Static Function PosiZA5
	Local cQryZA5 := ""
	Local lRet    := .F.	
	Local cSerie  := ""
	
	cSerie := u_getSerie('', SE1->E1_FILIAL)
	If ValType(cSerie) <> "C"
		cSerie := ""
	EndIf
	
	cQryZA5 := " SELECT ZA5.R_E_C_N_O_ AS RECNO FROM "  + RetSqlName( 'ZA5') + " ZA5"
	cQryZA5 += " WHERE ZA5.D_E_L_E_T_ <> '*' AND ZA5_ACAO = 'inclusao' "
	cQryZA5 += " AND ZA5_NUMNF = '"+SE1->E1_NUM+"' "
	cQryZA5 += " AND ZA5_SERNF = '" + cSerie +"'"
	cQryZA5 += " AND ZA5_CLIENT = '"+SE1->E1_CLIENTE+"' "
	cQryZA5 += " AND ZA5_LOJA = '"+SE1->E1_LOJA+"' "
	cQryZA5 += " AND ZA5_FILNFS = '"+SE1->E1_FILIAL+"' "
	cQryZA5 += " ORDER BY ZA5_DATA,ZA5_HORA DESC "
	
	If Select('QRYZA5') > 0
		QRYZA5->(dbclosearea())
	EndIf
	
	TcQuery cQryZA5 New Alias 'QRYZA5'
	
	If QRYZA5->(!Eof())
		ZA5->(DbGoto(QRYZA5->RECNO))
		lRet    := .T.
		QRYZA5->(dbSkip())
	EndIf

Return lRet


user Function L10VLSA1()
    Local cCodCli   := PADR(MV_PAR03, LEN(SA1->A1_COD))
    Local cCodLoja  := PADR(MV_PAR04, LEN(SA1->A1_LOJA))
    Local lRet      := .F.
    
    SA1->(DbSetOrder(1))
    If SA1->(DbSeek(xFilial('SA1') + cCodCli + cCodLoja))
        lRet := .T.
    endIf

    If Empty(cCodCli) .AND. Empty(cCodLoja)
        lRet := .T.
    endIf

    if !lRet
        ApMsgInfo('Cliente/loja não encontrado.')
    endIf
    
return lRet
