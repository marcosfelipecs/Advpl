#INCLUDE 'Rwmake.ch'
#INCLUDE 'Protheus.ch'
#INCLUDE 'TbIconn.ch'
#INCLUDE 'Topconn.ch'
#include 'DIRECTRY.CH'
//-------------------------------------------------------------------
/*/{Protheus.doc} LAUA0007
Títulos de liquidação
@author  Samuel Dantas
@since   26/12/2019
@version 1.0
/*/
//-------------------------------------------------------------------
User Function LAUA0007(lTodos)
	Private aRotina := {}
	
	private LOPCAUTO := .F.
    Private cCadastro := 'Titulos gerados'
    Private aRotina := {{'Visualizar'       ,'U_LAUA007A()'     ,0  , 2     },;
                        {"Imprimir BOL"     ,"U_LAUF0007()"     ,0  , 2     },;
                        {"Exportar CSV"     ,"u_LAUF010C()"     ,0  , 2     },;
                        {"Apenas fatura"    ,"u_LAUA007h(.F.)"  ,0  , 2     },;
                        {"Exp. XML"         ,"u_LAUA007C(.F., 'XML')"       ,0,2},;
                        {"Exp. Dactes"      ,"u_LAUA007C(.F., 'DACTES')"    ,0,2},;
                        {"Email Fatura"     ,"u_uFA740not()"     ,0,2},;
                        {'Cancelar Fat'     ,'u_LAUA007I()'     ,0,2},;
                        {'Enviar P/NUCCI'   ,'u_LAUA004W(SE1->E1_NUMLIQ)'     ,0,2}}                
                        

    cFiltro := " SE1->(E1_FILIAL+E1_YZA7) == '"+ZA7->(ZA7_FILIAL+ZA7_CODIGO)+"' "
    if !lTodos 
        SE1->(DbSetFilter({|| &(cFiltro) }, cFiltro))
    Else
        SE1->(DbSetFilter({|| SE1->E1_NUMLIQ != " " }, "SE1->E1_NUMLIQ != ' '"))
    EndIf
    
    
    Private cDelFunc := '.T.' // Validacao para a exclusao. Pode-se utilizar ExecBlock

    Private cString := 'SE1'
    
    DbSelectArea(cString)
    DbSetOrder(1)
    DbSelectArea(cString)

    mBrowse( 6,1,22,75,cString)
    
    SE1->(DBClearFilter())	
Return

User Function LAUA007B
    Local aLegenda := {}

    AADD(aLegenda,{"BR_VERDE"		, "Em aberto"	})
	AADD(aLegenda,{"BR_AZUL"		, "Baixa parcial"	})
	AADD(aLegenda,{"BR_VERMELHO"	, "Baixado"	})
	BrwLegenda(cCadastro, "Legenda", aLegenda)
Return NIL

//-------------------------------------------------------------------
/*/{Protheus.doc} BuscaSE1
Busca SE1 da liquidação e gera o filtro.
@author  Samuel Dantas
@since   19/01/2020
@version version
/*/
//-------------------------------------------------------------------
Static Function BuscaSE1()
    Local cRecnos   := ""
    Local cQry      := ""
    cFiltro := ""

    cQry := " SELECT FO0_NUMLIQ FROM "  + RetSqlName('FO0') + " FO0 "
    cQry += " WHERE FO0.D_E_L_E_T_ <> '*' AND FO0_YZA7 = '"+ZA7->ZA7_CODIGO+"' AND FO0_FILIAL = '"+ZA7->ZA7_FILIAL+"' "
    
    If Select('QRY') > 0
        QRY->(dbclosearea())
    EndIf
    
    TcQuery cQry New Alias 'QRY'
    
    While QRY->(!Eof())
        If Empty(cFiltro)
            cFiltro += " SE1->E1_NUMLIQ == '"+QRY->FO0_NUMLIQ+"' "
        Else
            cFiltro += " .OR. SE1->E1_NUMLIQ == '"+QRY->FO0_NUMLIQ+"' "
        EndIf

        QRY->(dbSkip())
    EndDo
    
Return cFiltro

//-------------------------------------------------------------------
/*/{Protheus.doc} LAUA007A
Posiciona FO0 e abre tela de visualização da liquidação
@author  author
@since   date
@version version
/*/
//-------------------------------------------------------------------
User Function LAUA007A()
    Local cQuery := ""

    If !Empty(SE1->E1_NUMLIQ )
        cQuery := " SELECT R_E_C_N_O_ AS RECNO FROM "  + RetSqlName('FO0') + " FO0 "
        cQuery += " WHERE FO0.D_E_L_E_T_ <> '*' AND FO0_NUMLIQ = '"+SE1->E1_NUMLIQ+"' AND FO0_FILIAL = '"+SE1->E1_FILIAL+"' "
        
        If Select('QRY') > 0
            QRY->(dbclosearea())
        EndIf
        
        TcQuery cQuery New Alias 'QRY'
        
        If QRY->(!Eof())
            FO0->(DbGoTo(QRY->RECNO))
            F460VerSim()

            QRY->(dbSkip())
        EndIf
    Else
        ApMsgInfo("Título não foi gerado a partir da rotina de liquidação.")    
    EndIf
        
Return

User Function LAUA007h
    Local cPathFAT := ""

    MSAguarde( { || cPathFAT := u_LAUA007D( .F. ) }, "Fatura" ,"Imprimindo Fatura", .F.)

Return

//-------------------------------------------------------------------
/*/{Protheus.doc} LAUA007C
Função utilizada para solicitação geração dos arquivos 
(xml, dactes e pré-fatura) das CTE
@author  Samuel Dantas
@since   19/01/2020
@version version
/*/
//-------------------------------------------------------------------
User Function LAUA007c(lEnviaEmail , cArqs, lVarios )
    Local cFile := ""
    Local cPathXML := ""
    Local cPathFAT := ""
    Private cDirFinal   := ""
    Default cArqs   := ""
    Default lVarios   := .F.

    If lEnviaEmail .AND. Empty(cArqs)
        cArqs := U_LAUA007F('2')
    Else

    EndIf
    lDactes     := IIf( "DACTES" $ cArqs    , .T. , .F. )
    lXml        := IIf( "XML" $ cArqs       , .T. , .F. )
    lPreFatura  := IIf( "FATURA" $ cArqs    , .T. , .F. )
    
    If !Empty(cArqs)    
        CriaPaths()
        If !lPreFatura
            cDirFinal   := cGetFile("Todos | *.* ", OemToAnsi("Selecione o diretorio"),  ,         ,.F.,GETF_LOCALHARD + GETF_RETDIRECTORY, .F.)  
        EndIf

        If lPreFatura
            MSAguarde( { || cPathFAT := u_LAUA007D( lEnviaEmail ) }, "Fatura" ,"Imprimindo Fatura",.F.)
        EndIf
        If lXml .OR. lDactes
            MSAguarde( { || cPathXML := ExporXML(cDirFinal, lXml, lDactes, lEnviaEmail) }, "Exportação de XML" ,"Exportando xml...",.F.)
        EndIf

        cFile := cPathFAT
        cFile := Iif(!Empty(cFile),cFile + ";"+cPathXML, cPathXML )
        
        If lEnviaEmail
            SA1->(DbSetOrder(1))
            SA1->(DbSeek(xFilial("SA1") + SE1->(E1_CLIENTE + E1_LOJA) ))
            cMsgEmail :=  "Prezado "+Alltrim(SA1->A1_NOME)+", " +CRLF
            cMsgEmail +=  CRLF
            cMsgEmail +=  "Segue em anexo Fatura e seus documentos." +CRLF
            cMsgEmail +=  CRLF
            cMsgEmail +=  "Fatura nº: "+Alltrim(SE1->E1_NUM)+"/"+Alltrim(SE1->E1_PREFIXO)+" " +CRLF
            cMsgEmail +=  "Vencimento: "+DtoC(SE1->E1_VENCTO)+" " +CRLF
            cMsgEmail +=  CRLF
            cMsgEmail +=  "Confirmar o recebimento no telefone e/ou email abaixo: " +CRLF
            cMsgEmail +=  CRLF
            cMsgEmail +=  "Departamento Financeiro, " +CRLF
            cMsgEmail +=  "L'AUTO CARGO" +CRLF
            cMsgEmail +=  "Tel.: (84) 3086-4084 " +CRLF
            If !lVarios
                GetMsg()
            EndIf

            If !Empty(cFile)
                cCC := AllTrim(SuperGetMv("MS_BOLMAIL",.F., "samuel.batista@agisrn.com"))
                cTeste := AllTrim(SuperGetMv("MS_TESMAIL",.F., "S"))
                cMailTeste := AllTrim(SuperGetMv("MS_MAILTES",.F., "samuel.batista@agisrn.com"))
                //Email amarrado ?
                If cTeste != 'S'
                    If ! U_EnviaEmail("FATURA LAUTO", SA1->A1_XEMCOB, cMsgEmail, cFile, cCC, .T.)
                    MsgAlert('Erro ao enviar email.')
                    Endif
                Else
                    If ! U_EnviaEmail("FATURA LAUTO", cMailTeste, cMsgEmail, cFile, "", .T.)
                    MsgAlert('Erro ao enviar email.')
                    Endif
                EndIf
            EndIf
        EndIF
    EndIf

Return

//-------------------------------------------------------------------
/*/{Protheus.doc} ExporXML
Função responsável por gerar xml e/ou dactes em arquivo definido 
pelo usuário.
@author  Samuel Dantas
@since   19/01/2020
@version version
/*/
//-------------------------------------------------------------------
Static Function ExporXML(cDirFinal, lXml, lDactes, lEnviaEmail)
    
    Local cQuery := ""
    Local cNumLiq := Alltrim(SE1->E1_NUM) + StrTran(Time(),":","")
    Local cFile := ""
    Local cFileRet := ""
    Local aAreaSE1 := SE1->(GetArea())
    Local cPathRar := ""
    Local _cFilAux := cFilAnt
    Default cDirFinal := ""
    Default lXml := .F.
    Default lDactes := .F.
    Default lEnviaEmail := .F.
    
    
    aChaves := GetTitFat()
    aToZip  := {}
    If !ExistDir(cDirFinal+"xmls")
        MakeDir(cDirFinal+"xmls\")
    EndIf
    If !ExistDir(cDirFinal+"dactes")
        MakeDir(cDirFinal+"dactes\")
    EndIf

    cDirXML   := cDirFinal+"xmls\"
    cDirDacte := cDirFinal+"dactes\"
    
    For nI := 1 To len(aChaves)
        SE1->(DbSeek(aChaves[nI]))
        setEmpresa(SE1->E1_FILIAL)
        PosiTabs()

        If ALLTRIM(SE1->E1_PREFIXO) $ '2,3'
            cQuery := " SELECT ZA4.R_E_C_N_O_ AS RECNO FROM "  + RetSqlName( 'ZA4') + " ZA4"
            cQuery += " WHERE ZA4.D_E_L_E_T_ <> '*' AND ZA4_CHAVE = '"+SF2->F2_CHVNFE+"' AND ZA4_ACAO = 'inclusao' "
        Else
            cQuery := " SELECT ZA5.R_E_C_N_O_ AS RECNO FROM "  + RetSqlName( 'ZA5') + " ZA5"
            cQuery += " WHERE ZA5.D_E_L_E_T_ <> '*' AND ZA5_ACAO = 'inclusao' "
            cQuery += " AND ZA5_NUMNF = '"+SF2->F2_DOC+"' "
            cQuery += " AND ZA5_SERNF = '"+SF2->F2_SERIE+"' "
            cQuery += " AND ZA5_CLIENT = '"+SF2->F2_CLIENTE+"' "
            cQuery += " AND ZA5_LOJA = '"+SF2->F2_LOJA+"' "
            cQuery += " AND ZA5_FILNFS = '"+SF2->F2_FILIAL+"' "
        EndIf

        If Select('QRY') > 0
            QRY->(dbclosearea())
        EndIf
        
        TcQuery cQuery New Alias 'QRY'
        
        While QRY->(!Eof())
            If ALLTRIM(SE1->E1_PREFIXO) $ '2,3'
                ZA4->(DbGoTo(QRY->RECNO))
                If lXml 
                    Memowrite(cDirXMl + (cArqXML := Alltrim(ZA4->ZA4_CHAVE) + ".XML"), ZA4->ZA4_BODY)
                    aAdd(aToZip,cDirXMl + cArqXML)   
                EndIf
                If lDactes
                    u_uRTMSR31(ZA4->ZA4_BODY, cArqCte := Alltrim(ZA4->ZA4_CHAVE) + ".PDF", cDirDacte)                    
                    aAdd(aToZip,cDirDacte + cArqCte)
                EndIf
            Else
                ZA5->(DbGoTo(QRY->RECNO))
                If lXml 
                    Memowrite(cDirXMl + (cArqXML := Alltrim(ZA5->ZA5_FILNFS)+Alltrim(ZA5->ZA5_NUMNF) + ".XML"), ZA5->ZA5_BODY)
                    aAdd(aToZip,cDirXMl + cArqXML)   
                EndIf
            EndIf
            QRY->(dbSkip())
        EndDo        
        
    Next
    setEmpresa(_cFilAux)
    cPathRar := IIF( lXml .AND. lDactes, "_DACTES_XML.zip", IIF(lDactes, "_DACTES.zip", "_XML.zip"))

    If FZip(cDirFinal+ cNumLiq + cPathRar, aToZip, cDirFinal) == 0
        
        CpyT2S( cDirFinal+ cNumLiq + cPathRar, '\anexos\')

        cFile := '\anexos\' +  cNumLiq + cPathRar
        // Deleta arquivos dentro do diretorio
        aEval(Directory(cDirFinal+"dactes\*.pdf")   , { |aFile| FERASE(cDirFinal+"dactes\"+aFile[1]) })
        aEval(Directory(cDirFinal+"xmls\*.xml")     , { |aFile| FERASE(cDirFinal+"xmls\"+aFile[1])  })
        
        // remove o diretorio
        DirRemove( cDirFinal+"dactes\" )
        DirRemove( cDirFinal+"xmls\" )
        If !lEnviaEmail
            ApMsgInfo("Arquivo Salvo com sucesso.")
        EndIf
        
    Else
        MsgAlert('Erro ao tentar compactar xmls e dactes para envio. ')
    EndIf

    SE1->(RestArea(aAreaSE1))

Return cFile

//-------------------------------------------------------------------
/*/{Protheus.doc} LAUF010C
Método responsável pela impressão de pré-fatura(layout de fatura sem boleto)
@author  Samuel Dantas
@since   19/01/2020
@version version
/*/
//-------------------------------------------------------------------
User Function LAUA007D(lEnviaEmail)
    Local cFileRet := ""
    Local cPath       := "c:\temp\"
    Local aAreaSE1 := SE1->(GetArea())
    
    
    Private cTpImpBol   := GetMv("MV_XTPBOL")
    Private cQualBco := ""
    Private cNossoDg := ""
    Private cStgTipo := "'"
    Private bOrigCB  := .F.
    Private cPerg    := "BOLETO"
    Private lEnd     := .F.
    Private aReturn  := {"Banco",;				// [1]= Tipo de Formulário
                        1,;           			// [2]= Número de Vias
                        "Administração",;		// [3]= Destinatário
                        2,;						// [4]= Formato 1=Comprimido 2=Normal
                        2,;						// [5]= Mídia 1=Disco 2=Impressora
                        1,;						// [6]= Porta LPT1, LPT2, Etc
                        "",;						// [7]= Expressão do Filtro
                        1}						// [8]= ordem a ser selecionada
                        
    Private cTitulo    := "Boleto de Cobrança com Código de Barras"
    Private cStartPath := GetSrvProfString("StartPath","")
    Private nTpImp     := 0
    Private nPosPDF    := 0
    Private aLinDig    := {}
    Private aRet          := {}
    Private aRecnos     := {}
    Private cMsgErro    := ""
    Private nVlrFat     := 0
    Private cCliente    := ""
    Private cLoja       := ""
    Private cNatureza   := ""
    Private dVencto     := dDataBase
    Private nNumPag   := 1
    Private lBx       := .F.
    Private cDirGer   := ""
    Private cNumBor   := ""
    Private cBanco    := ""
    Private cCmpLv    := ""
    Private cNN       := ""
    Private cCart     := ""
    Private cNNum     := ""
    Private cConvenio := ""
    Private cLogo     := ""
    Private nDesc     := 0
    Private nJurMul   := 0
    Private nRow      := 0
    Private nCols     := 0
    Private nWith     := 0
    Private cDirServ  := "\anexos\boletos\"
    Private cMsgEmail := ""
    Private aChaves   := {}
    Default lEnviaEmail := .F.

    cStartPath := AllTrim(cStartPath) + "logo_bancos\"

    oPrint := FwMSPrinter():New("fatura"+Alltrim(SE1->E1_NUM), 6, .T., cPath, lEnviaEmail, .F., , , , .T., , !lEnviaEmail )
    oPrint:SetResolution(72)
    oPrint:SetMargin(5,5,5,5)
    oPrint:SetPortrait()	
    oPrint:StartPage()
    u_fnImpres(@oPrint, {}, {}, {}, {}, {}, {}, {}, '', .T. )
    
    If lEnviaEmail
        oPrint:cPathPdf := cPath
    EndIf

    cDirFinal := oPrint:cPathPdf
    cFileRet := cPath + StrTran(oPrint:cFileName,".rel",".pdf")
    // File2Printer(StrTran(oPrint:cFileName,".rel",".PD_"),"PDF")
    oPrint:Preview()     // Visualiza antes de imprimir
    CpyT2S( cFileRet, cDirServ )
    cFileRet := cDirServ + StrTran(oPrint:cFileName,".rel",".pdf")
    SE1->(RestArea(aAreaSE1))
Return cFileRet

//-------------------------------------------------------------------
/*/{Protheus.doc} LAUA007F
Tela de seleção dos tipos de arquivos (dactes, xml, pré-fatura) que serão gerados
@author  Samuel Dantas
@since   19/01/2020
@version version
/*/
//-------------------------------------------------------------------
User Function LAUA007F(cTipoFat)
	Default cTipoFat := '1' // 1 - Boleto, 2 - Pós fatura, 3 - Pré-fatura
	Private cRetorno := ""
    Private aPergs	  := {}
	Private aRet	  := {}
	Private aDados	  := {}
	Private nPosRecno := 13
    Private oDlgArq
    Private oBrowse
	Private oPanel01
	Private oPanelModal
    Private aFields := {}
	Private cMark 	 := GetMark()
	Private oWindow 	
	
	aDados := retDados(cTipoFat)
    
    //Cria a tela
    oDlgArq := FwDialogModal():New()
  	oDlgArq:SetTitle('Arquivos à acrescentar')
    oDlgArq:SetSize(300,550)
	oDlgArq:CreateDialog()	
	
	oPanelModal	:= oDlgArq:GetPanelMain()
	
	oLayer := FwLayer():New()
	oLayer:Init(oPanelModal)
	
	//Linha dos Filtros
	oLayer:addLine("lin01",100,.T.)		
	oLayer:AddCollumn('col01', 100, .T., 'lin01')	
	oLayer:addWindow("col01","win01",'Títulos',100,.F.,.F.,{|| }, "lin01",{|| })
	oPanel01 := oLayer:getWinPanel('col01', 'win01', "lin01")

	//Cria array com os campo de marcacao
	aMark	:= {{ "{|| if(self:oData:aArray[self:nAt, 1] == cMark, 'LBOK','LBNO') }", "{|| u_LAUA007G(.F.) }", "{|| u_LAUA007G(.T.)  }" }}	
	oBrowse 	:= uFwBrowse():create(oPanel01,,aDados[2],,aDados[1],, aMark,, .F.)	

	oBrowse:disableConfig()
	oBrowse:disableFilter()
	oBrowse:disableReport()

	oBrowse:Activate()
	 

    oDlgArq:AddButton('Fechar'		, {|| oDlgArq:Deactivate() }, 'Sair', , .T., .T., .T., ) 
	oDlgArq:AddButton('Confirmar'	, {|| Confirmar() }, 'Confirmar',, .T., .T., .T., )    //Ativa a tela principal
	oDlgArq:Activate()	
		                                 
Return cRetorno

//-------------------------------------------------------------------
/*/{Protheus.doc} Confirma
Retorna string com tipos de arquivos a serem enviados via email
@author  Samuel Dantas
@since   13/01/2019
@version version
/*/
//-------------------------------------------------------------------
Static Function Confirmar()
  Local aArray    := oBrowse:oData:aArray 
	Local nAt	      := oBrowse:nAt

  For nK := 1 To len(aArray)
    If aArray[nk][1] == cMarK
      cRetorno += IIF(Empty(cRetorno), aArray[nk][2], ","+Alltrim(aArray[nk][2]))
      
    EndIf
  Next
  oDlgArq:Deactivate()
Return


//-------------------------------------------------------------------
/*/{Protheus.doc} LAUA007G
Funcao que marca/desmarca os itens da grid
@author  Samuel Dantas
@since   08/02/19
@version 1.0
/*/
//-------------------------------------------------------------------
User Function LAUA007G(lTodos)
	Local aArray:= oBrowse:oData:aArray 
	Local nAt	:= oBrowse:nAt
	If lTodos
		For i := 1 to Len(aArray)
			if aArray[i, 1] == cMark
				aArray[i, 1] := '  '
			else
				aArray[i, 1] := cMark		
			EndIf
		Next
	Else
		if aArray[oBrowse:nAt, 1] == cMarK
			aArray[oBrowse:nAt, 1] := '  '
		else
			aArray[oBrowse:nAt, 1] := cMark		
		EndIf           
	EndIf

	oBrowse:Refresh()

Return

//-------------------------------------------------------------------
/*/{Protheus.doc} retdados
Funcao para retornar os dados para a grid
@author  Samuel Dantas
@since   08/02/18
@version 1.0
/*/
//-------------------------------------------------------------------
Static Function retDados(cTipoFat)
	Local cQuery 	:= ""
	Local aHeader	:= {}
	Local aCols		:= {}
    Default cTipoFat := '1'

    //Cabecalho da grid de titulos
    aAdd(aHeader,{' '	      ,2  , '' })
    aAdd(aHeader,{'Arquivo'	  ,10 , '' })
    If cTipoFat == '1'
        aAdd(aCols, { cMarK, "BOLETO" })
        aAdd(aCols, { "  ", "XML"    })
        aAdd(aCols, { "  ", "DACTES" })
    ElseIf cTipoFat == '2'
        aAdd(aCols, { cMarK, "FATURA" })
        aAdd(aCols, { "  ", "XML"    })
        aAdd(aCols, { "  ", "DACTES" })
    ElseIf cTipoFat == '3'
        aAdd(aCols, { cMarK, "XML"    })
        aAdd(aCols, { cMarK, "DACTES" })
    EndIf

Return { aHeader, aCols }


Static Function GetMsg()
  Local oFont := TFont():New('Courier new',,-18,.T.)
  Local oDlg 
  
  Local oMultiGet, oDlg
	Local oPanel1, oPanel2, oPanel3
	Default cObserv	:= Space(100) 
	Default cTitulo 	:= "Observação via Email" 
	Default cTitulo2	:= "Observação para ser enviada por email:"
	Default lConfirmar:= .T.

	DEFINE MSDIALOG oDlg FROM 1,1 TO 250,450 TITLE cTitulo PIXEL

	@C(001),C(001) MSPANEL oPanel1 PROMPT "" SIZE C(001),C(015) OF oDlg
	oPanel1:align := CONTROL_ALIGN_TOP

	@C(001),C(001) MSPANEL oPanel2 PROMPT "" SIZE C(001),C(015) OF oDlg
	oPanel2:align := CONTROL_ALIGN_ALLCLIENT

	@C(001),C(001) MSPANEL oPanel3 PROMPT "" SIZE C(001),C(015) OF oDlg
	oPanel3:align := CONTROL_ALIGN_BOTTOM 

	@ C(002),C(005) Say cTitulo2 Size C(100),C(008) COLOR CLR_BLACK PIXEL OF oPanel1
	oMultiGet := TMultiget():Create(oPanel2,{|u| if(Pcount() > 0,cMsgEmail := u, cMsgEmail )},1,1,1,1,,nil,nil,nil,nil,.T.,nil,nil, ,nil,nil,,{ || .T.})	
	oMultiGet:align := CONTROL_ALIGN_ALLCLIENT

	If lConfirmar
		TButton():New(005, 005, "&Confirmar", oPanel3,{|| oDlg:End() },40,010,,,.F.,.T.,.F.,,.F.,,,.F.)
		TButton():New(005, 055, "&Fechar"	, oPanel3,{|| oDlg:End() },40,010,,,.F.,.T.,.F.,,.F.,,,.F.)
	Else
		TButton():New(005, 005, "&Fechar"	, oPanel3,{|| oDlg:End() },40,010,,,.F.,.T.,.F.,,.F.,,,.F.)		
	EndIf


	ACTIVATE MSDIALOG oDlg CENTERED
  
Return

User Function LAUA007I()
    MSAguarde( { ||  u_LAUA007W() }, "Cancelamento de fatura" ,"Cancelando fatura",.F.)
Return


User Function LAUA007W()
    Local oJson 
    Local cFilAux1 := ""
    private LOPCAUTO := .F.
    // oJson := U_LAUF010E(SE1->E1_NUMLIQ)
    // oJson['data_cancelamento'] := dDatabase
    cFilAux1 := cFilAnt
    cFilAnt := SE1->E1_FILIAL
    BEGIN TRANSACTION
        lRet:= ExecCanc()
    END TRANSACTION
    cFilAnt := cFilAux1
    // U_LAUA004X(oJson)
Return


User Function LAUA007L
    
    If Empty(FunName())
        PREPARE ENVIRONMENT EMPRESA '01' FILIAL '01RN0001'
    EndIf

    Private aPergs	  := {}
	Private aRet	  := {}
	Private aDados	  := {}
	Private nPosRecno := 14
    Private oDlg
    Private oBrowse
	Private oPanel01
	Private dDatade := dDataBase
	Private dDataAte := dDataBase
	Private cCliente    := ""
	Private cLoja       := ""
	Private cGrpVen       := ""
	Private oPanelModal
    Private aFields := {}
	Private cMark 	 := GetMark()
	Private oWindow 
    Private cPerg := PadR('U_LAUA0007', 10)

	ValidPerg()
    
    If ! Pergunte(cPerg, .T.)
        Return
    EndIf
    
    dDatade  := MV_PAR01
    dDataAte := MV_PAR02
    cCliente := MV_PAR03
    cLoja    := MV_PAR04
    cGrpVen    := MV_PAR05

	aDados := retFats()
	
    //Cria a tela
    oDlg := FwDialogModal():New()
  	oDlg:SetTitle('Envio de faturas')
    oDlg:SetSize(300,550)
	oDlg:CreateDialog()	
	oDlg:addCloseButton( , "Fechar")
	
	oPanelModal	:= oDlg:GetPanelMain()
	
	oLayer := FwLayer():New()
	oLayer:Init(oPanelModal)
	
	//Linha dos Filtros
	oLayer:addLine("lin01",100,.T.)		
	oLayer:AddCollumn('col01', 100, .T., 'lin01')	
	oLayer:addWindow("col01","win01",'Títulos',100,.F.,.F.,{|| }, "lin01",{|| })
	oPanel01 := oLayer:getWinPanel('col01', 'win01', "lin01")

	//Cria array com os campo de marcacao
	aMark	:= {{ "{|| if(self:oData:aArray[self:nAt, 1] == cMark, 'LBOK','LBNO') }", "{|| u_LAUA007T(.F.) }", "{|| u_LAUA007T(.T.)  }" }}	
	oBrowse 	:= uFwBrowse():create(oPanel01,,aDados[2],,aDados[1],, aMark,, .F.)	

	oBrowse:disableConfig()
	oBrowse:disableFilter()
	oBrowse:disableReport()
    oDlg:AddButton('Enviar Fats.'			, { || MsAguarde( { || u_LAUA007U(@oDlg) }	, "Envio de faturas","Aguarde envio das faturas." )}	, 'Enviar Fats.', 	, .T., .T., .T., ) 
	oBrowse:Activate()
	 

   //Ativa a tela principal
	oDlg:Activate()	
		                                 
Return

//-------------------------------------------------------------------
/*/{Protheus.doc} LAUA007U
Funcao que marca/desmarca os itens da grid
@author  Samuel Dantas
@since   08/02/19
@version 1.0
/*/
//-------------------------------------------------------------------
User Function LAUA007U(oDlg)
	Local aArray:= oBrowse:oData:aArray 
	Local nAt	:= oBrowse:nAt
	Local i := 1
    Local cArqs := ""
    Local aTitulos := {}
    Local nOp := 0
    Local cBorAnt := ""
    Local cBorAtu := ""
    Local lFirst := .T.
    Local lBorDif := .F.
    nOp :=  Aviso("Impressão e/ou envio de fatura.", "Deseja imprimir/enviar boleto ou apenas fatura?", {"Boleto", "Apenas Fatura", "Cancelar"}, 2, "Selecione uma ação. Ao selecionar impressão de boleto, será gerado um borderô para esta fatura.")
    If nOp == 2 // Faturas
        cArqs := U_LAUA007F('2')
        cArqs += IIF( !("FATURA" $ cArqs ), ",FATURA", "" )
    ElseIf nOp := 1 // boletos
        // cArqs += IIF( !("BOLETO" $ cArqs ), ",BOLETO", "" )
    Else
        oDlg:oOwner:End()	
    EndIf
    
    For i := 1 to Len(aArray)
        
        if aArray[i, 1] == cMark
            cBorAnt := SE1->E1_NUMBOR
            SE1->(DbGoTo(aArray[i][nPosRecno]))
            
            If lFirst 
                cBorAnt := SE1->E1_NUMBOR
                lFirst := .F.
            EndIf

            cBorAtu := SE1->E1_NUMBOR
            
            If cBorAtu != cBorAnt
                lBorDif := .T.
            EndIf

            cFilAnt := SE1->E1_FILIAL
            aAdd(aTitulos, {{"E1_FILIAL" ,SE1->E1_FILIAL }, {"E1_PREFIXO" ,SE1->E1_PREFIXO }, {"E1_NUM" ,SE1->E1_NUM }, {"E1_PARCELA" ,SE1->E1_PARCELA }, {"E1_TIPO" ,SE1->E1_TIPO }} )
            If nOp == 2
                U_LAUA007C(.T.,cArqs,.T.)
            EndIf
        EndIf
    Next
    
    If nOp == 1 .AND. len(aTitulos) > 0 .AND. !lBorDif
        U_LAUF0007(aTitulos)
    EndIf
    
    //Se tiver titulos de borderôs diferentes não deixa excutar
    If lBorDif
        ApMsgInfo("Não é possível reenviar conjunto de títulos de borderô difentes. Selecione títulos SEM borderô OU EM MESMO borderô.")
    EndIf

	oDlg:oOwner:End()	

Return

//-------------------------------------------------------------------
/*/{Protheus.doc} LAUA007T
Funcao que marca/desmarca os itens da grid
@author  Samuel Dantas
@since   08/02/19
@version 1.0
/*/
//-------------------------------------------------------------------
User Function LAUA007T(lTodos)
	Local aArray:= oBrowse:oData:aArray 
	Local nAt	:= oBrowse:nAt
	If lTodos
		For i := 1 to Len(aArray)
			if aArray[i, 1] == cMark
				aArray[i, 1] := '  '
			else
				aArray[i, 1] := cMark		
			EndIf
		Next
	Else
		if aArray[oBrowse:nAt, 1] == cMarK
			aArray[oBrowse:nAt, 1] := '  '
		else
			aArray[oBrowse:nAt, 1] := cMark		
		EndIf           
	EndIf

	oBrowse:Refresh()

Return

//-------------------------------------------------------------------
/*/{Protheus.doc} retFats
Funcao para retornar os dados para a grid
@author  Samuel Dantas
@since   08/02/18
@version 1.0
/*/
//-------------------------------------------------------------------
Static Function retFats()
	Local cQuery 	:= ""
	Local aHeader	:= {}
	Local aCols		:= {}

    cQrySE1 := " SELECT DISTINCT SE1.R_E_C_N_O_ AS RECNOSE1 FROM "  + RetSqlName('SE1') + " SE1"
    cQrySE1 += " INNER JOIN "  + RetSqlName('SA1') + " SA1 ON A1_FILIAL = LEFT(E1_FILIAL,2) AND A1_COD = E1_CLIENTE AND A1_LOJA = E1_LOJA "
    cQrySE1 += " WHERE SE1.D_E_L_E_T_ <> '*' AND E1_PREFIXO = 'FAT' "
    cQrySE1 += " AND E1_EMISSAO BETWEEN '"+DtoS(dDatade)+"' AND '"+DtoS(dDataAte)+"' "
    
    If !Empty(cCliente)
        cQrySE1 += " AND E1_CLIENTE = '"+cCliente+"' "
    EndIf
    If !Empty(cLoja)
        cQrySE1 += " AND E1_LOJA = '"+cLoja+"' "
    EndIf

    If !Empty(cGrpVen)
        cQrySE1 += " AND A1_GRPVEN = '"+cGrpVen+"' "
    EndIf
    
    If Select('QRYSE1') > 0
        QRYSE1->(dbclosearea())
    EndIf
    
    TcQuery cQrySE1 New Alias 'QRYSE1'
    
	
	//Cabecalho da grid de titulos
	aAdd(aHeader,{''			,2, '' 				})
	aAdd(aHeader,{'E1_FILIAL'	,10, 'E1_FILIAL'	})
	aAdd(aHeader,{'E1_PREFIXO'	,10, 'E1_PREFIXO'	})
	aAdd(aHeader,{'E1_NUM'		,10, 'E1_NUM' 		})
	aAdd(aHeader,{'E1_TIPO'		,10, 'E1_TIPO' 		})
    aAdd(aHeader,{'E1_PARCELA'	,10, 'E1_PARCELA' 	})
	aAdd(aHeader,{'E1_EMISSAO'	,10, 'E1_EMISSAO' 	})
	aAdd(aHeader,{'E1_VENCTO'	,10, 'E1_VENCTO' 	})
    aAdd(aHeader,{'E1_CLIENTE'	,10, 'E1_CLIENTE' 	})
	aAdd(aHeader,{'E1_LOJA'	    ,10, 'E1_LOJA' 	})
	aAdd(aHeader,{'E1_SALDO'	,10, 'E1_SALDO' 	})
	aAdd(aHeader,{'E1_VALOR'	,10, 'E1_VALOR' 	})
	aAdd(aHeader,{'Email'	    ,10, '' 	})
	aAdd(aHeader,{'IDENTIFICADOR'	,10, '' 	})
    
      
    While QRYSE1->(!Eof())
        SE1->(DbGoTo(QRYSE1->RECNOSE1))
        cEmails := Posicione("SA1",1,xFilial("SA1")+SE1->(E1_CLIENTE+E1_LOJA), "A1_XEMCOB")
        aAdd(aCols, {'  ', SE1->E1_FILIAL, SE1->E1_PREFIXO, SE1->E1_NUM, SE1->E1_TIPO,SE1->E1_PARCELA, SE1->E1_EMISSAO, SE1->E1_VENCTO,;
							 SE1->E1_CLIENTE, SE1->E1_LOJA,; 
							 SE1->E1_SALDO, SE1->E1_VALOR,cEmails,SE1->(Recno())})
        QRYSE1->(dbSkip())
    EndDo

    If len(aCols) == 0
        aAdd(aCols, {'', '','','','','',CtoD(''),CtoD(''),'','',0,0,"",0}  )
    EndIf

Return { aHeader, aCols }


//-------------------------------------------------------------------
/*/{Protheus.doc} ValidPerg
Criacao das perguntas do relatorio
@author  author
@since   11/03/2020
@version 1.0
/*/
//-------------------------------------------------------------------
Static Function ValidPerg
    Local aRegs    := {}
    Local aAreaSX1 := SX1->(GetArea())
    Local i, j

    SX1->(DbSetOrder(1))

    // Numeracao dos campos:
    // 01 -> X1_GRUPO   02 -> X1_ORDEM    03 -> X1_PERGUNT  04 -> X1_PERSPA  05 -> X1_PERENG
    // 06 -> X1_VARIAVL 07 -> X1_TIPO     08 -> X1_TAMANHO  09 -> X1_DECIMAL 10 -> X1_PRESEL
    // 11 -> X1_GSC     12 -> X1_VALID    13 -> X1_VAR01    14 -> X1_DEF01   15 -> X1_DEFSPA1
    // 16 -> X1_DEFENG1 17 -> X1_CNT01    18 -> X1_VAR02    19 -> X1_DEF02   20 -> X1_DEFSPA2
    // 21 -> X1_DEFENG2 22 -> X1_CNT02    23 -> X1_VAR03    24 -> X1_DEF03   25 -> X1_DEFSPA3
    // 26 -> X1_DEFENG3 27 -> X1_CNT03    28 -> X1_VAR04    29 -> X1_DEF04   30 -> X1_DEFSPA4
    // 31 -> X1_DEFENG4 32 -> X1_CNT04    33 -> X1_VAR05    34 -> X1_DEF05   35 -> X1_DEFSPA5
    // 36 -> X1_DEFENG5 37 -> X1_CNT05    38 -> X1_F3       39 -> X1_GRPSXG

    aAdd(aRegs, {cPerg, "01", "Data de?"  , "", "", "mv_ch1", 'D', 8, 0, 0, 'G', "", "MV_PAR01", "",  "", "", "", "", "",    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""})
    aAdd(aRegs, {cPerg, "02", "Data até?" , "", "", "mv_ch2", 'D', 8, 0, 0, 'G', "", "MV_PAR02", "",  "", "", "", "", "",    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""})
    aAdd(aRegs, {cPerg, "03", "Cliente?" , "", "", "mv_ch3", 'C', LEN(SA1->A1_COD), 0, 0, 'G', "", "MV_PAR03", "",  "", "", "", "", "",    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "SA1", ""})
    aAdd(aRegs, {cPerg, "04", "Loja?" , "", "", "mv_ch4", 'C', 4, 0, 0, 'G', "", "MV_PAR04", "",  "", "", "", "", "",    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""})
    aAdd(aRegs, {cPerg, "05", "Grupo?" , "", "", "mv_ch5", 'C', 6, 0, 0, 'G', "", "MV_PAR05", "",  "", "", "", "", "",    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "AYC", ""})

    For i := 1 To Len(aRegs)
        If ! SX1->(DbSeek(cPerg+aRegs[i,2]))
            RecLock("SX1", .T.)

            For j :=1 to SX1->(FCount())
                If j <= Len(aRegs[i])
                    SX1->(FieldPut(j,aRegs[i,j]))
                EndIf
            Next

            SX1->(MsUnlock())
        EndIf
    Next

    SX1->(RestArea(aAreaSX1))
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
  
  If SE1->E1_PREFIXO $ '2  ,3  '
    cQuery := " SELECT SC5.R_E_C_N_O_ AS RECNOSC5, SF2.R_E_C_N_O_ AS RECNOSF2, DT6.R_E_C_N_O_ AS RECNODT6 FROM "  + RetSqlName('SC5') + " SC5"
  Else
    cQuery := " SELECT SC5.R_E_C_N_O_ AS RECNOSC5, SF2.R_E_C_N_O_ AS RECNOSF2 FROM "  + RetSqlName('SC5') + " SC5"
  EndIf

  cQuery += " INNER JOIN  "+RetSqlName('SF2')+" SF2 ON F2_FILIAL = C5_FILIAL AND F2_DOC = C5_NOTA AND F2_SERIE = C5_SERIE AND F2_CLIENTE = C5_CLIENTE AND F2_LOJA = C5_LOJACLI AND SF2.D_E_L_E_T_ = SC5.D_E_L_E_T_ "
  If SE1->E1_PREFIXO $ '2  ,3  '
    cQuery += " INNER JOIN  "+RetSqlName('DT6')+" DT6 ON F2_FILIAL = DT6_FILIAL AND DT6_CHVCTE = F2_CHVNFE AND SF2.D_E_L_E_T_ = DT6.D_E_L_E_T_ "
  EndIF
  cQuery += " WHERE SC5.D_E_L_E_T_ <> '*'"
  cQuery += " AND C5_FILIAL = '"+SE1->E1_FILIAL+"' "
  cQuery += " AND C5_NUM = '"+SE1->E1_PEDIDO+"' "
  
  If Select('QRY') > 0
    QRY->(dbclosearea())
  EndIf
  
  TcQuery cQuery New Alias 'QRY'
  
  If QRY->(!Eof())
    SF2->(DbGoTo(QRY->RECNOSF2))
    SD2->(DbSetOrder(3)) //D2_FILIAL, D2_DOC, D2_SERIE, D2_CLIENTE, D2_LOJA, D2_COD, D2_ITEM, R_E_C_N_O_, D_E_L_E_T_
    SD2->(DbSeek(SF2->(F2_FILIAL + F2_DOC + F2_SERIE + F2_CLIENTE + F2_LOJA )))
    SC5->(DbGoTo(QRY->RECNOSC5))
    If SE1->E1_PREFIXO $ '2  ,3  '
      DT6->(DbGoTo(QRY->RECNODT6))
    EndIF
    
    // DT6->(DBOrderNickname("DT6CHVCTE"))
    // DT6->(DbSeek(xFilial("DT6") + SF2->F2_CHVNFE))

    lRet := .T.
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

Static Function GetTitFat()
  Local aChaves := {}
  Local cQuery := ""

  cQuery := " SELECT DISTINCT FK7_CHAVE FROM "  + RetSqlName('FO0') + " FO0"
  cQuery += " INNER JOIN "+RetSqlName("FO1")+" FO1 ON FO0_FILIAL = FO1_FILIAL AND FO0_PROCES = FO1_PROCES "
  cQuery += " INNER JOIN "+RetSqlName("FK7")+" FK7 ON LEFT(FK7_FILIAL,2) = LEFT(FO1_FILIAL,2) AND FK7_IDDOC = FO1_IDDOC "
  cQuery += " WHERE FO0.D_E_L_E_T_ <> '*' AND FO0_NUMLIQ = '"+SE1->E1_NUMLIQ+"' AND FK7_ALIAS = 'SE1' AND FO0_FILIAL = '"+SE1->E1_FILIAL+"' AND FO0_STATUS <> '5' "
  
  If Select('QRY') > 0
    QRY->(dbclosearea())
  EndIf
  
  TcQuery cQuery New Alias 'QRY'
  
  While QRY->(!Eof())
    aAdd(aChaves, StrTran(QRY->FK7_CHAVE, "|" , "") )
    QRY->(dbSkip())
  EndDo

Return aChaves

Static Function CriaPaths()
  
  If !ExistDir("c:temp")
    MakeDir("c:\temp\")
  EndIf
  
  If !ExistDir("C:\TEMP")
    MakeDir("C:\TEMP\")
  EndIf

  If !ExistDir("\anexos")
    MakeDir("\anexos\")
  EndIf

  If !ExistDir("\anexos\boletos")
    MakeDir("\anexos\boletos\")
  EndIf

  If !ExistDir("\anexos\dactes")
    MakeDir("\anexos\dactes\")
  EndIf

  If !ExistDir("\anexos\xmls")
    MakeDir("\anexos\xmls\")
  EndIf

  If !ExistDir("c:temp\dactes")
    MakeDir("c:\temp\dactes\")
  EndIf

  If !ExistDir("c:temp\xmls")
    MakeDir("c:\temp\xmls\")
  EndIf


Return

//-------------------------------------------------------------------
/*/{Protheus.doc} ExecCanc
Executa cancelamento de fatura via execauto
@author  Samuel Dantas
@since   18/02/2020
@version version
/*/
//-------------------------------------------------------------------
Static Function ExecCanc()
    Local _cNumLiq           := ""
    Local lRet              := .T.
    Private LOPCAUTO        := .T.
    Private lMsErroAuto     := .F.

    _cFilAux    := cFilAnt
    cFilAnt     := SE1->E1_FILIAL
    _cNumLiq     := SE1->E1_NUMLIQ
    // u_FA460Can()
    MSExecAuTo({|x,y,k,w,z,p| FINA460(x,y,k,w,z,p)}, ,{} , {} , 5, "", SE1->E1_NUMLIQ)  	
    
    cFilAnt := _cFilAux
    
    If lMsErroAuto
        MostraErro()
        lRet := .F.
        return
    Else 
        If !u_ValidaCanc(_cNumLiq)
            lRet := .F.
            ApMsgInfo("Cancelamento da liquidação não foi efetuado. Entre em contato com um analista.(LAUF0006)")
            // DisarmTransaction()
        EndIf
    EndIf

Return lRet
