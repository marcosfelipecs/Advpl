#INCLUDE "PROTHEUS.CH"
#INCLUDE "APWIZARD.CH"

// #########################################################################################
// Projeto:
// Modulo :
// Fonte  : Import
// ---------+-------------------+-----------------------------------------------------------
// Data     | Autor             | Descricao
// ---------+-------------------+-----------------------------------------------------------
// 28/06/13 | Karlos Morais     | Programa responsavel para importação de dados das tabelas
//          |                   |   - Cliente (SA1)
//          |                   |   - Fornecedor (SA2)
//          |                   |   - Contas a Receber (SE1)
//          |                   |   - Contas a Pagar (SE2)
// ---------+-------------------+-----------------------------------------------------------
// #########################################################################################

User Function zImport()

	// ######################################################
	//		Variaveis
	// ######################################################

	Local oWizard
	Local nMetGlob
	Local nMetParc
	Local oRadioArq
	Local nRadioArq	:= 1
	Local cText
	Local cFile		:= Replicate( " ", 80 )
	Local cHeader 		:= "Importação de dados"
	Local cTpArq		:= "Delimitado (*.csv)|*.CSV|"
	Local cDelim		:= AllTrim(SuperGetMV("MV_TPDELI",.F.,';'))
	Local nLinCabec	:= 1	// Padrão sem linha de cabeçalho
	Local cCabec		:= ""	// String com o cabeçalho do arquivo original, se houver
	Local nQtdCab		:= 1	// String com o cabeçalho do arquivo original, se houver
	Local cNmAlias		:= "Clientes (SA1)"
	Local cTipo		:= "1"

	Private INCLUI	:= .T.
	Private ALTERA	:= .F.

	cText 	:= 	 "Esta rotina tem por objetivo importar registros, através " + ;
				 "de um arquivo padrão CSV (delimitado) , e armazena-los na tabela "+ ;
				 "correspondente do sistema."+ CRLF + ;
				 "Os nomes das colunas devem ser os mesmos nomes de campos a serem atualizados."+ CRLF + CRLF + ;
				 "Ao final da importação será gerado um arquivo de log contendo as "+ ;
				 "inconsistências."

	// ######################################################
	//		Primeiro Painel - Abertura
	// ######################################################

	DEFINE WIZARD oWizard		TITLE "Importação de dados" ;
								HEADER cHeader ;
								MESSAGE "Apresentação." ;
								TEXT cText ;
								NEXT { || .T. } ;
								FINISH {|| .T.} PANEL

	// ######################################################
	//		Segundo Painel - Arquivo e Contrato
	// ######################################################

	CREATE PANEL oWizard 	HEADER cHeader ;
							MESSAGE "Selecione os tabela que deseja importar" ;
							BACK {|| .T. } ;
							NEXT {|| .T. } ;
							FINISH {|| .F. } ;
							PANEL

	oPanel := oWizard:GetPanel( 2 )

	@ 15, 08 GROUP oGrpCon 	TO 120, 230 LABEL "Cadastro a ser importado" OF oPanel PIXEL DESIGN

	@ 25,35 Radio oRadioArq Var nRadioArq Items	"Clientes (SA1)",;
														"Fornecedores (SA2)",;
														"Contas a Receber - em aberto (SE1)",;
														"Contas a Pagar - em aberto (SE2)",;
														"Funcionario (SRA)";
														3D 	Size 170,10 Of oPanel PIXEL DESIGN ;
														ON CHANGE ImpChgRadio(nRadioArq,@cNmAlias)

	// ######################################################
	//		Segundo Painel - Arquivo e Contrato
	// ######################################################

	CREATE PANEL oWizard 	HEADER cHeader ;
							MESSAGE "Selecione o arquivo para importação." ;
							BACK {|| .T. } ;
							NEXT {|| ! empty( cDelim ) .and. ! empty( cFile ) } ;
							FINISH {|| .F. } ;
							PANEL

	oPanel := oWizard:GetPanel( 3 )

	@ 10, 08 GROUP oGrpCon 	TO 40, 280 LABEL "Selecione um arquivo." ;
								OF oPanel ;
								PIXEL ;
								DESIGN

	@ 20, 15 MSGET oArq 	VAR cFile WHEN .F. OF oPanel SIZE 140, 10 PIXEL ;
							MESSAGE "Utilize o botão ao lado para selecionar" ;

	DEFINE SBUTTON oButArq 	FROM 21, 160 ;
								TYPE 14 ;
								ACTION cFile := cGetFile(cTpArq, , 0, "SERVIDOR\", .T., GETF_LOCALHARD + GETF_NETWORKDRIVE) ;
								OF oPanel ;
								ENABLE


	@ 50, 08 GROUP oGrpCon 	TO 130, 280 LABEL "Informe as configurações do arquivo." ;
								OF oPanel ;
								PIXEL ;
								DESIGN

  	@ 60,20 SAY "Delimitador" OF oPanel SIZE 35,8 PIXEL
	@ 60,60 MSGET oDelim	VAR cDelim  ;
							PICTURE "@!" ;
							VALID !empty(cDelim) ;
							MESSAGE "Informe um delimitador de campo." ;
							OF oPanel SIZE 10,8 PIXEL

  	@ 80,20 SAY "Tipo" OF oPanel SIZE 35,8 PIXEL
	@ 80,60 COMBOBOX oTipo Var cTipo ITEMS {"1=Somente Log","2=Log + Importação"} SIZE 200,010 OF oPanel PIXEL

	// ######################################################
	//		Terceiro Painel - Confirmacao  / Processamento
	// ######################################################

	CREATE PANEL oWizard 	HEADER cHeader ;
							MESSAGE "Confirmação dos dados e início de processamento." ;
							BACK {|| .T. } ;
							NEXT {|| .T. } ;
							FINISH {|| .F. } ;
							PANEL

	oPanel := oWizard:GetPanel( 4 )

	@ 010, 010 SAY "Arquivo" OF oPanel SIZE 140, 8 PIXEL
	@ 010, 050 SAY cFile  OF oPanel SIZE 140, 8 COLOR CLR_HBLUE PIXEL

	@ 030, 010 SAY  "Delimitador" OF oPanel SIZE 140, 8 PIXEL
	@ 030, 050 SAY  cDelim  OF oPanel SIZE 140, 8 COLOR CLR_HBLUE PIXEL


	@ 050, 010 SAY  "Alias" OF oPanel SIZE 140, 8 PIXEL
	@ 050, 050 SAY  cNmAlias  OF oPanel SIZE 140, 8 COLOR CLR_HBLUE PIXEL


	@ 070, 010 SAY  "Tipo Proc.:" OF oPanel SIZE 140, 8 PIXEL
	@ 070, 050 SAY  IIf(cTipo=="1","Somente Log","Log+Importação")  OF oPanel SIZE 140, 8 COLOR CLR_HBLUE PIXEL


	// ######################################################
	//		Quarto Painel - Processamento
	// ######################################################

	CREATE PANEL oWizard 	HEADER cHeader ;
							MESSAGE "Processamento da Importação." ;
							BACK {|| .F. } ;
							NEXT {|| .T. } ;
							FINISH {|| .T. } ;
							EXEC {|| CursorWait(), IMPCADPro( oMetGlob, nRadioArq, cFile, cDelim, cTipo ), CursorArrow() } ;
							PANEL

	oPanel := oWizard:GetPanel( 5 )

	@ 25, 30 SAY "Importação" OF oPanel SIZE 140, 8 PIXEL
	@ 40, 30 METER oMetGlob 	VAR nMetGlob ;
								TOTAL 100 ;
								SIZE 224,10 OF oPanel PIXEL UPDATE DESIGN ;
								BARCOLOR CLR_BLACK,CLR_WHITE ;
								COLOR CLR_WHITE,CLR_BLACK ;
							 	NOPERCENTAGE

	ACTIVATE WIZARD oWizard CENTER

Return NIL


// #########################################################################################
// Programa  : Import
// ---------+-------------------+-----------------------------------------------------------
// Data     | Autor             | Descricao
// ---------+-------------------+-----------------------------------------------------------
// 28/06/13 | Karlos Morais     |
// ---------+-------------------+-----------------------------------------------------------
// #########################################################################################

Static Function IMPCADPro( oMetGlob, nRadioArq, cFile, cDelim,cTipo )

	// ######################################################
	//		Declaracao de Variaveis
	// ######################################################

	Local aArea		:= GetArea()
	Local lFirst		:= .T.
	Local cLinha 		:= ""
	Local aHeader		:= {}
	Local nHdl			:= 	0
	Local cEnvServ		:= GetEnvServer()
	Local cIniFile		:= GetADV97()
	Local cEnd			:= GetPvProfString(cEnvServ,"StartPath","",cIniFile)
	Local cDtHr		:= DtoS(dDataBase)+"-"+Substr(time(),1,2)+"-"+Substr(time(),4,2)+"-"+Substr(time(),7,2)
	Local cPath		:= "\IMPORT\"
	Local cTipoLog		:= "Import_"
	Local cNomeLog		:=	cPath+cTipoLog+cDtHr+"_Log.txt"
	Local cArq			:=	cEnd+cNomeLog
	Local cLin			:= ""
	Local cCdAlias		:= ""
	Local nQtReg		:= 0
	Local nQtNOk		:= 0
	Local nQtOk		:= 0
	Local aLog			:= {}
	Local lGrava		:= (cTipo == "2")
	Local cRotina		:= ""
	Local nCont		:= 0

	MAKEDIR(cEnd+cPath)

	// ######################################################
	//		Validacao do arquivo para importacao
	// ######################################################

	If !File(cFile) .OR. Empty(cFile)
		ApMsgStop("Problemas com arquivo informado!")
		RestArea(aArea)
		Return
	EndIf

	// ######################################################
	//		Identifica Alias de importacao
	// ######################################################

	Do Case
		Case nRadioArq == 1		// "Clientes (SA1)",;
			cCdAlias	:= "SA1"
			cRotina	:= "MATA030"

		Case nRadioArq == 2		// "Fornecedores (SA2)",;
			cCdAlias	:= "SA2"
			cRotina	:= "MATA020"

		Case nRadioArq == 3		// "Contas a Receber - em aberto (SE1)",;
			cCdAlias	:= "SE1"
			cRotina	:= "FINA040"

		Case nRadioArq == 4		// "Contas a Pagar - em aberto (SE2)"
			cCdAlias	:= "SE2"
			cRotina	:= "FINA050"

		Case nRadioArq == 5		// "Funcionario (SRA)"
			cCdAlias	:= "SRA"
			cRotina	:= "GPEA010"

		OtherWise
			ApMsgStop("Nao existe tratamento para importação deste tipo de arquivo!")
			Return
	EndCase

	// ######################################################
	//		Inicia Log
	// ######################################################

	AAdd(aLog, Replicate( '=', 80 ) )
	AAdd(aLog, 'INICIANDO O LOG - I M P O R T A C A O   D E   D A D O S' )
	AAdd(aLog, Replicate( '-', 80 ) )
	AAdd(aLog, 'DATABASE...........: ' + DtoC( dDataBase ) )
	AAdd(aLog, 'DATA...............: ' + DtoC( Date() ) )
	AAdd(aLog, 'HORA...............: ' + Time() )
	AAdd(aLog, 'ENVIRONMENT........: ' + GetEnvServer() )
	AAdd(aLog, 'PATCH..............: ' + GetSrvProfString( 'StartPath', '' ) )
	AAdd(aLog, 'ROOT...............: ' + GetSrvProfString( 'RootPath', '' ) )
	AAdd(aLog, 'VERSÃO.............: ' + GetVersao() )
	AAdd(aLog, 'MÓDULO.............: ' + 'SIGA' + cModulo )
	AAdd(aLog, 'EMPRESA / FILIAL...: ' + SM0->M0_CODIGO + '/' + SM0->M0_CODFIL )
	AAdd(aLog, 'NOME EMPRESA.......: ' + Capital( Trim( SM0->M0_NOME ) ) )
	AAdd(aLog, 'NOME FILIAL........: ' + Capital( Trim( SM0->M0_FILIAL ) ) )
	AAdd(aLog, 'USUÁRIO............: ' + SubStr( cUsuario, 7, 15 ) )
	AAdd(aLog, 'TABELA IMPORT......: ' + cCdAlias )
	AAdd(aLog, 'ARQUIVO IMPORT.....: ' + cFile )
	AAdd(aLog, 'DELIMITADOR........: ' + cDelim )
	AAdd(aLog, 'MODO PROCESSAMENTO.: ' + IIf(lGrava,"Atualizacao","Simulacao") )
	AAdd(aLog, Replicate( ':', 80 ) )
	AAdd(aLog, '' )

	AAdd(aLog, "Import = INICIO - Data "+DtoC(dDataBase)+ " as "+Time() )

	// ######################################################
	//		Leitura do arquivo
	// ######################################################

	FT_FUSE(cFile)

	nTot	:= FT_FLASTREC()
	nAtu	:= 0

	oMetGlob:SetTotal(nTot)
	CursorWait()

	FT_FGOTOP()

	While !FT_FEOF()

		nAtu++
		oMetGlob:Set(nAtu)

		cLinha := LeLinha() //FT_FREADLN()

		If Empty(cLinha)
			FT_FSKIP()
			Loop
		EndIf

		// ######################################################
		//		Tratamento de colunas
		// ######################################################

		aCols	:= {}
		aCols	:= TrataCols(cLinha,cDelim)

		If lFirst

			aHeader := aClone(aCols)
			lFirst := .F.

			// ######################################################
			//		Valida nomes das colunas
			// ######################################################

			cCpos := ImpVldCols(cCdAlias,aHeader)

			If !Empty(cCpos)
				ApMsgStop("Problemas na estrutura do arquivo, faltam as seguintes colunas "+cCpos)
				Return
			EndIf

		Else

			nQtReg++

			// ######################################################
			//		Validacao de campos obrigatorios
			// ######################################################

			cMsg := ImpObrigat(cCdAlias,aCols,aHeader)

			If !Empty(cMsg)
				AtuLog("NO MOT: CAMPOS OBRIGATORIOS - REGISTRO IGNORADO - "+cMsg,@aLog,nAtu)
				nQtNOk++
				FT_FSKIP()
				Loop
			EndIf

			// ######################################################
			//		Chamada de rotina automatica de inclusao
			// ######################################################

			If lGrava

				aRet := {}
				aRet := ImpGrava(cCdAlias,cRotina,aCols,aHeader)

				If aRet[1]
					nQtOk++
					AtuLog("OK MOT:REGISTRO INCLUIDO"+aRet[2],@aLog,nAtu)
				Else
					AtuLog("NO MOT: PROBLEMAS NA GRAVACAO ROTINA AUTOMATICA - "+cRotina+" - "+aRet[2],@aLog,nAtu)
					nQtNOk++
				EndIF

			Else

				nQtOk++
				AtuLog("OK MOT:REGISTRO INCLUIDO",@aLog,nAtu)

			EndIf
		EndIf

		FT_FSKIP()
	End

	FT_FUSE()

	AAdd(aLog, "Import = Total de Registros = "+ Alltrim(Str(nQtReg)))
	AAdd(aLog, "Import = Registros Nao importados = "+ Alltrim(Str(nQtNOk)))
	AAdd(aLog, "Import = Registros importados = "+ Alltrim(Str(nQtOk)))
	AAdd(aLog, "Import = FIM Data "+DtoC(dDataBase)+ " as "+Time() )

	// ######################################################
	//		Finaliza arquivo de Log
	// ######################################################

	nHdl  := 	fCreate(cArq)

	If nHdl == -1
		MsgAlert("O arquivo  "+cArq+" nao pode ser criado!","Atencao!")
		fClose(nHdl)
		fErase(cArq)
		RestArea(aArea)
	 	Return()
	EndIf

	For nCont:=1 to Len(aLog)

		cLin += aLog[nCont] + CHR(13)+CHR(10)

		If fWrite(nHdl,cLin,Len(cLin)) != Len(cLin)
			fClose(nHdl)
		    fErase(cArq)
		    cLin:=""
			RestArea(aArea)
		    Return()
		EndIf

		cLin:=""

	Next

	fClose(nHdl)

	ApMsgInfo("Verifique arquivo de log "+cArq)

	RestArea(aArea)

Return


// #########################################################################################
// Programa  : AtuLog
// ---------+-------------------+-----------------------------------------------------------
// Data     | Autor             | Descricao
// ---------+-------------------+-----------------------------------------------------------
// 28/06/13 | Karlos Morais     |
// ---------+-------------------+-----------------------------------------------------------
// #########################################################################################

Static Function AtuLog(cMsg,aLog,nAtu)

	AAdd(aLog, " Import = Linha " + StrZero(nAtu,12) + " = " + " LOG = " + cMsg)

Return


// #########################################################################################
// Programa  : LeLinha
// ---------+-------------------+-----------------------------------------------------------
// Data     | Autor             | Descricao
// ---------+-------------------+-----------------------------------------------------------
// 28/06/13 | Karlos Morais     |
// ---------+-------------------+-----------------------------------------------------------
// #########################################################################################

Static Function LeLinha()

	Local cLinhaTmp	:= ""
	Local cLinhaM100	:= ""

	cLinhaTmp := FT_FReadLN()

	If !Empty(cLinhaTmp)

		cIdent:= Substr(cLinhaTmp,1,1)

		If Len(cLinhaTmp) < 1023
			cLinhaM100 := cLinhaTmp
		Else

			cLinAnt := cLinhaTmp
			cLinhaM100 += cLinAnt

			Ft_FSkip()

			cLinProx:= Ft_FReadLN()

			If Len(cLinProx) >= 1023 .and. Substr(cLinProx,1,1) <> cIdent

				While Len(cLinProx) >= 1023 .and. Substr(cLinProx,1,1) <> cIdent .and. !Ft_fEof()

					cLinhaM100 += cLinProx

					Ft_FSkip()

					cLinProx := Ft_fReadLn()

					If Len(cLinProx) < 1023 .and. Substr(cLinProx,1,1) <> cIdent
						cLinhaM100 += cLinProx
					Endif

				Enddo

			Else
				cLinhaM100 += cLinProx
			Endif

		Endif

	Endif

Return(cLinhaM100)


// #########################################################################################
// Programa  : TrataCols
// ---------+-------------------+-----------------------------------------------------------
// Data     | Autor             | Descricao
// ---------+-------------------+-----------------------------------------------------------
// 28/06/13 | Karlos Morais     |
// ---------+-------------------+-----------------------------------------------------------
// #########################################################################################

Static Function TrataCols(cLinha,cSep)

	Local aRet		:= {}
	Local nPosSep	:= 0

	nPosSep	:= At(cSep,cLinha)

	While nPosSep <> 0

		AAdd(aRet, SubStr(cLinha,1,nPosSep-1)  )

		cLinha		:= SubStr(cLinha,nPosSep+1)
	 	nPosSep	:= At(cSep,cLinha)

	EndDo

	AAdd(aRet, cLinha )

Return aRet


// #########################################################################################
// Programa  : RetCol
// ---------+-------------------+-----------------------------------------------------------
// Data     | Autor             | Descricao
// ---------+-------------------+-----------------------------------------------------------
// 28/06/13 | Karlos Morais     |
// ---------+-------------------+-----------------------------------------------------------
// #########################################################################################

Static Function RetCol(cCpo,aCols,aHeader)

	Local cRet	:= ""
	Local nPos	:= 0
	Local aSX3Area	:= SX3->(GetArea())

	nPos := AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim(cCpo)) })

	If !Empty(nPos)

		If Upper(AllTrim(aCols[nPos])) <> "NULL"

			DbSelectArea("SX3")
			DbSetOrder(2)

			If MsSeek(cCpo)
				If SX3->X3_TIPO == "D"
					cRet := StoD(AllTrim(aCols[nPos]))
				ElseIf SX3->X3_TIPO == "N"
					cRet := Val(AllTrim(aCols[nPos]))
				Else
					cRet := PadR(Upper(AllTrim(aCols[nPos])),TamSX3(cCpo)[1])
				EndIf
			Else
				cRet := Upper(AllTrim(aCols[nPos]))
			EndIf

		EndIf

	EndIf

	SX3->(RestArea(aSX3Area))

Return cRet


// #########################################################################################
// Programa  : ImpVldCols
// ---------+-------------------+-----------------------------------------------------------
// Data     | Autor             | Descricao
// ---------+-------------------+-----------------------------------------------------------
// 28/06/13 | Karlos Morais     |
// ---------+-------------------+-----------------------------------------------------------
// #########################################################################################

Static Function ImpVldCols(cCdAlias,aHeader)

	Local cRet 	:= ""
	Local cFilSA1 	:= ""

	Do Case

		Case cCdAlias == "SA1"
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("A1_LOJA")) })	== 0
				cRet += IIf(Empty(cRet),"","/")+"A1_LOJA"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("A1_NOME")) })	== 0
				cRet += IIf(Empty(cRet),"","/")+"A1_NOME"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("A1_NREDUZ")) })	== 0
				cRet += IIf(Empty(cRet),"","/")+"A1_NREDUZ"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("A1_END")) })		== 0
				cRet += IIf(Empty(cRet),"","/")+"A1_END"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("A1_TIPO")) })	== 0
				cRet += IIf(Empty(cRet),"","/")+"A1_TIPO"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("A1_EST")) })		== 0
				cRet += IIf(Empty(cRet),"","/")+"A1_EST"
			EndIf

		Case cCdAlias == "SA2"
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("A2_LOJA")) })	== 0
				cRet += IIf(Empty(cRet),"","/")+"A2_LOJA"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("A2_NOME")) })	== 0
				cRet += IIf(Empty(cRet),"","/")+"A2_NOME"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("A2_NREDUZ")) })	== 0
				cRet += IIf(Empty(cRet),"","/")+"A2_NREDUZ"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("A2_END")) })		== 0
				cRet += IIf(Empty(cRet),"","/")+"A2_END"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("A2_EST")) })		== 0
				cRet += IIf(Empty(cRet),"","/")+"A2_EST"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("A2_MUN")) })		== 0
				cRet += IIf(Empty(cRet),"","/")+"A2_MUN"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("A2_TIPO")) })	== 0
				cRet += IIf(Empty(cRet),"","/")+"A2_TIPO"
			EndIf

		Case cCdAlias == "SE1"
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("E1_NUM")) })		== 0
				cRet += IIf(Empty(cRet),"","/")+"E1_NUM"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("E1_TIPO")) })	== 0
				cRet += IIf(Empty(cRet),"","/")+"E1_TIPO"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("E1_NATUREZ")) })	== 0
				cRet += IIf(Empty(cRet),"","/")+"E1_NATUREZ"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("E1_CLIENTE")) })	== 0
				cRet += IIf(Empty(cRet),"","/")+"E1_CLIENTE"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("E1_LOJA")) })	== 0
				cRet += IIf(Empty(cRet),"","/")+"E1_LOJA"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("E1_EMISSAO")) })	== 0
				cRet += IIf(Empty(cRet),"","/")+"E1_EMISSAO"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("E1_VENCTO")) })	== 0
				cRet += IIf(Empty(cRet),"","/")+"E1_VENCTO"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("E1_VENCREA")) })	== 0
				cRet += IIf(Empty(cRet),"","/")+"E1_VENCREA"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("E1_VALOR")) })	== 0
				cRet += IIf(Empty(cRet),"","/")+"E1_VALOR"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("E1_VLCRUZ")) })	== 0
				cRet += IIf(Empty(cRet),"","/")+"E1_VLCRUZ"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("E1_PORCJUR")) })	== 0
				cRet += IIf(Empty(cRet),"","/")+"E1_PORCJUR"
			EndIf

		Case cCdAlias == "SE2"
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("E2_NUM")) })		== 0
				cRet += IIf(Empty(cRet),"","/")+"E2_NUM"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("E2_TIPO")) })	== 0
				cRet += IIf(Empty(cRet),"","/")+"E2_TIPO"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("E2_NATUREZ")) })	== 0
				cRet += IIf(Empty(cRet),"","/")+"E2_NATUREZ"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("E2_FORNECE")) })	== 0
				cRet += IIf(Empty(cRet),"","/")+"E2_FORNECE"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("E2_LOJA")) })	== 0
				cRet += IIf(Empty(cRet),"","/")+"E2_LOJA"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("E2_EMISSAO")) })	== 0
				cRet += IIf(Empty(cRet),"","/")+"E2_EMISSAO"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("E2_VENCTO")) })	== 0
				cRet += IIf(Empty(cRet),"","/")+"E2_VENCTO"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("E2_VENCREA")) })	== 0
				cRet += IIf(Empty(cRet),"","/")+"E2_VENCREA"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("E2_VALOR")) })	== 0
				cRet += IIf(Empty(cRet),"","/")+"E2_VALOR"
			EndIf
			If AScan(aHeader,{|x| Upper(AllTrim(x)) == Upper(Alltrim("E2_VLCRUZ")) })	== 0
				cRet += IIf(Empty(cRet),"","/")+"E2_VLCRUZ"
			EndIf

	EndCase

Return cRet


// #########################################################################################
// Programa  : ImpObrigat
// ---------+-------------------+-----------------------------------------------------------
// Data     | Autor             | Descricao
// ---------+-------------------+-----------------------------------------------------------
// 28/06/13 | Karlos Morais     |
// ---------+-------------------+-----------------------------------------------------------
// #########################################################################################

Static Function ImpObrigat(cCdAlias,aCols,aHeader)

	Local cRet := ""

	Do Case

		Case cCdAlias == "SA1"
			If Empty(RetCol("A1_LOJA",aCols,aHeader))
				cRet += " / Coluna A1_LOJA esta vazia! "
			EndIf // Removido devido o campo ser gerado automaticamente
			If Empty(RetCol("A1_NOME",aCols,aHeader))
				cRet += " / Coluna A1_NOME esta vazia! "
			EndIf
			If Empty(RetCol("A1_NREDUZ",aCols,aHeader))
				cRet += " / Coluna A1_NREDUZ esta vazia! "
			EndIf
			If Empty(RetCol("A1_END",aCols,aHeader))
				cRet += " / Coluna A1_END esta vazia! "
			EndIf
			If Empty(RetCol("A1_TIPO",aCols,aHeader))
				cRet += " / Coluna A1_TIPO esta vazia! "
			EndIf
			If Empty(RetCol("A1_EST",aCols,aHeader))
				cRet += " / Coluna A1_EST esta vazia! "
			EndIf

		Case cCdAlias == "SA2"
			If Empty(RetCol("A2_LOJA",aCols,aHeader))
				cRet += " / Coluna A2_LOJA esta vazia! "
			EndIf // Removido devido o campo ser gerado automaticamente
			If Empty(RetCol("A2_NOME",aCols,aHeader))
				cRet += " / Coluna A2_NOME esta vazia! "
			EndIf
			If Empty(RetCol("A2_NREDUZ",aCols,aHeader))
				cRet += " / Coluna A2_NREDUZ esta vazia! "
			EndIf
			If Empty(RetCol("A2_END",aCols,aHeader))
				cRet += " / Coluna A2_END esta vazia! "
			EndIf
			If Empty(RetCol("A2_EST",aCols,aHeader))
				cRet += " / Coluna A2_EST esta vazia! "
			EndIf
			If Empty(RetCol("A2_MUN",aCols,aHeader))
				cRet += " / Coluna A2_MUN esta vazia! "
			EndIf
			If Empty(RetCol("A2_TIPO",aCols,aHeader))
				cRet += " / Coluna A2_TIPO esta vazia! "
			EndIf

		Case cCdAlias == "SE1"
			If Empty(RetCol("E1_NUM",aCols,aHeader))
				cRet += " / Coluna E1_NUM esta vazia! "
			EndIf
			If Empty(RetCol("E1_TIPO",aCols,aHeader))
				cRet += " / Coluna E1_TIPO esta vazia! "
			EndIf
			If Empty(RetCol("E1_NATUREZ",aCols,aHeader))
				cRet += " / Coluna E1_NATUREZ esta vazia! "
			EndIf
			If Empty(RetCol("E1_CLIENTE",aCols,aHeader))
				cRet += " / Coluna E1_CLIENTE esta vazia! "
			EndIf
			If Empty(RetCol("E1_LOJA",aCols,aHeader))
				cRet += " / Coluna E1_LOJA esta vazia! "
			EndIf
			If Empty(RetCol("E1_EMISSAO",aCols,aHeader))
				cRet += " / Coluna E1_EMISSAO esta vazia! "
			EndIf
			If Empty(RetCol("E1_VENCTO",aCols,aHeader))
				cRet += " / Coluna E1_VENCTO esta vazia! "
			EndIf
			If Empty(RetCol("E1_VENCREA",aCols,aHeader))
				cRet += " / Coluna E1_VENCREA esta vazia! "
			EndIf
			If Empty(RetCol("E1_VALOR",aCols,aHeader))
				cRet += " / Coluna E1_VALOR esta vazia! "
			EndIf
			If Empty(RetCol("E1_VLCRUZ",aCols,aHeader))
				cRet += " / Coluna E1_VLCRUZ esta vazia! "
			EndIf

		Case cCdAlias == "SE2"
			If Empty(RetCol("E2_NUM",aCols,aHeader))
				cRet += " / Coluna E2_NUM esta vazia! "
			EndIf
			If Empty(RetCol("E2_TIPO",aCols,aHeader))
				cRet += " / Coluna E2_TIPO esta vazia! "
			EndIf
			If Empty(RetCol("E2_NATUREZ",aCols,aHeader))
				cRet += " / Coluna E2_NATUREZ esta vazia! "
			EndIf
			If Empty(RetCol("E2_FORNECE",aCols,aHeader))
				cRet += " / Coluna E2_FORNECE esta vazia! "
			EndIf
			If Empty(RetCol("E2_LOJA",aCols,aHeader))
				cRet += " / Coluna E2_LOJA esta vazia! "
			EndIf
			If Empty(RetCol("E2_EMISSAO",aCols,aHeader))
				cRet += " / Coluna E2_EMISSAO esta vazia! "
			EndIf
			If Empty(RetCol("E2_VENCTO",aCols,aHeader))
				cRet += " / Coluna E2_VENCTO esta vazia! "
			EndIf
			If Empty(RetCol("E2_VENCREA",aCols,aHeader))
				cRet += " / Coluna E2_VENCREA esta vazia! "
			EndIf
			If Empty(RetCol("E2_VALOR",aCols,aHeader))
				cRet += " / Coluna E2_VALOR esta vazia! "
			EndIf
			If Empty(RetCol("E2_VLCRUZ",aCols,aHeader))
				cRet += " / Coluna E2_VLCRUZ esta vazia! "
			EndIf

	EndCase

Return cRet


// #########################################################################################
// Programa  : ImpGrava
// ---------+-------------------+-----------------------------------------------------------
// Data     | Autor             | Descricao
// ---------+-------------------+-----------------------------------------------------------
// 28/06/13 | Karlos Morais     |
// ---------+-------------------+-----------------------------------------------------------
// #########################################################################################

Static Function ImpGrava(cCdAlias,cRotina,aCols,aHeader)

	Local nX				:= 0
	Local cRotAuto			:= ""
	Local lOk				:= .F.
	Local cMsg				:= ""
	Local lGeraNumSeq		:= .T.
	Local cArqErro			:= "ERRO_AUTO.TXT"
	Local lTemFilial		:= .F.
	Local cCpoFilial 		:= IIf( SubStr(cCdAlias,1,1) == "S", SubStr(cCdAlias,2,2), cCdAlias) + "_FILIAL"
	Local cFilAlias		:= xFilial(cCdAlias)
	Private lMsHelpAuto	:= .T.
	Private lMsErroAuto	:= .F.
	Private aReg			:= {}

	// ######################################################
	//		Monta array com os campos do registro
	// ######################################################

	For nX:=1 to Len(aHeader)

		AAdd(aReg, {	Upper(Alltrim(aHeader[nX])),;
						RetCol(Alltrim(aHeader[nX]),aCols,aHeader),;
						Nil} )

		// ######################################################
		//		Verifica se possui campo sequencial informado
		// ######################################################

		If cCdAlias == "SA1" .AND. Upper(Alltrim(aHeader[nX])) == "A1_COD" .AND. !Empty(RetCol(Alltrim(aHeader[nX]),aCols,aHeader))
			lGeraNumSeq	:= .F.
		EndIf
		If cCdAlias == "SA2" .AND. Upper(Alltrim(aHeader[nX])) == "A2_COD" .AND. !Empty(RetCol(Alltrim(aHeader[nX]),aCols,aHeader))
			lGeraNumSeq	:= .F.
		EndIf

		// ######################################################
		//	Verifica se informou filial no arquivo
		// ######################################################

		If Upper(Alltrim(aHeader[nX])) == cCpoFilial
			lTemFilial := .T.
	    EndIf

	Next

	If lGeraNumSeq

		If cCdAlias == "SA1"
			AAdd(aReg, {	"A1_COD",;
							GetSxeNum("SA1","A1_COD"),;
							Nil} )
			ConfirmSx8()

			AAdd(aReg, {	"A1_LOJA",;
							"01",;
							Nil} )

		EndIf

		If cCdAlias == "SA2"
			AAdd(aReg, {	"A2_COD",;
							GetSxeNum("SA2","A2_COD"),;
							Nil} )
			ConfirmSx8()

			AAdd(aReg, {	"A2_LOJA",;
							"01",;
							Nil} )

		EndIf

	EndIf

	/*If !lTemFilial
		DbSelectArea(cCdAlias)
		AAdd(aReg, {	cCpoFilial,;
						cFilAlias,;
						Nil} )
	EndIf*/

	// ######################################################
	//		Chamada da rotina automatica
	// ######################################################

	DbSelectArea(cCdAlias)

	cRotAuto := "MSExecAuto({|x,y| "+cRotina+"(x,y)},aReg,3)"
	&cRotAuto

	If lMsErroAuto
		MostraErro( GetSrvProfString("Startpath","") , cArqErro )
		cMsg := MemoRead( GetSrvProfString("Startpath","") + '\' + cArqErro )
	Else
		if cCdAlias == "SA1"
			RecLock("SA1",.F.)
	    	SA1->A1_FILIAL := aReg[1][2]    
	      	Msunlock()
	    ELSEIF cCdAlias == "SA2"		    	
			RecLock("SA2",.F.)
	    	SA2->A2_FILIAL := aReg[1][2]    
	      	Msunlock()
	   ELSEIF cCdAlias == "SE1"		    	
			RecLock("SE1",.F.)
	    	SE1->E1_FILIAL := aReg[1][2]    
	      	Msunlock()
	   ELSEIF cCdAlias == "SE2"		    	
			RecLock("SE2",.F.)
	    	SE2->E2_FILIAL := aReg[1][2]    
	      	Msunlock()
	    ENDIF 	    		
		lOk := .T.
	EndIf

Return {lOk, cMsg }


// #########################################################################################
// Programa  : ImpChgRadio
// ---------+-------------------+-----------------------------------------------------------
// Data     | Autor             | Descricao
// ---------+-------------------+-----------------------------------------------------------
// 28/06/13 | Karlos Morais     |
// ---------+-------------------+-----------------------------------------------------------
// #########################################################################################

Static Function ImpChgRadio(nRadioArq,cNmAlias)

	Do Case
		Case nRadioArq == 1
			cNmAlias := "Clientes (SA1)"
		Case nRadioArq == 2
			cNmAlias := "Fornecedores (SA2)"
		Case nRadioArq == 3
			cNmAlias := "Contas a Receber - em aberto (SE1)"
		Case nRadioArq == 4
			cNmAlias := "Contas a Pagar - em aberto (SE2)"
		Case nRadioArq == 5
			cNmAlias := "Funcionarios (SRA)"
		OtherWise
			cNmAlias := "Processamento nao implementado"
	EndCase

Return