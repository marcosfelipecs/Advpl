#INCLUDE 'TOTVS.CH'
#INCLUDE 'PROTHEUS.CH'
#INCLUDE 'APVT100.CH'
#INCLUDE 'ACDV317.CH'
#INCLUDE 'FWMVCDEF.CH'

//------------------------------------------------------------------
/*/{Protheus.doc} ACDV317
Tela de Apontamento de Parada Geral

@author  Ana Carolina Tome Klock
@since   01/09/2013
@version P12
/*/
//------------------------------------------------------------------
Function ACDV317()
//Codigo da Area Selecionado

Private cArea      := ""
Private cDenMotivo := Space(TamSX3('CYN_DSSP')[1])
Private cConfirm   := Space(1)

If SuperGetMv('MV_INTACD',.F.,"0") != "1"
   VTAlert(STR0028, STR0009) //"Integração com o ACD desativada. Verifique o parâmetro MV_INTACD" ### 'Erro'
   Return
EndIf

ACDV317I()

Return Nil

//-------------------------------------------------------------------
/*/{Protheus.doc} ACDV317I
Método para inclusão de um apontamento de parada geral

@author  Ana Carolina Tome Klock
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function ACDV317I()
Local aDados  :={ { 'CZ7_CDARPO' , InicPadrao(1)                 } ,;
                  { 'CZ7_CDCETR', Space(TamSX3('CYI_CDCETR')[1]) } ,;
                  { 'CZ7_CDSP'  , InicPadrao(3)                  } ,;
                  { 'CZ7_DTBGSP', DTOC(Date())                   } ,;
                  { 'CZ7_HRBGSP', Time()                         } ,;
                  { 'CZ7_DTEDSP', DTOC(Date())                   } ,;
                  { 'CZ7_HREDSP', Time()                         } }

Private aCDCETR := {}
IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf

if LVT100B 
	// Monta a tela
	VTCLEAR()
	VTClearBuffer()
	
	@0,0 VTSAY STR0001 // "Apontam Parada Geral"
	@1,0 VTSAY STR0005+":" // "Área de Produção"
	@2,0 VTGET aDados[1][2] Pict "@!" F3 "CYA001" Valid ValidCampo(@aDados[1][2], 1, aDados[1][1] )
	VTREAD()
	If VTLastkey() == 27  // Tecla ESC
	   Return Nil
	Else
	   aCDCETR := GetCDCETR(aDados[1][2], @aDados)
	   cArea   := Iif(Len(aCDCETR)>0,aCDCETR[1],"")
	EndIf
	
	// Monta a tela
	If len(aCDCETR) > 1
	   aDados[2][2] := InicPadrao(2)
	
	   @3,0 VTSAY STR0006+":" // "Centro de Trabalho"
	   @3,VTCOL() VTGET aDados[2][2] Pict "@!" F3 "CYI002" Valid ValidCampo(@aDados[2][2], 2, aDados[2][1] )
	   VTREAD()
	   If VTLastkey() == 27    // Tecla ESC
	      Return Nil
	   EndIf
	EndIf
	
	// Monta a tela
	VTCLEAR()
	VTClearBuffer()
	
	//@0,0 VTSAY STR0001 // "Apontam Parada Geral"
	If len(aCDCETR) > 1
	   @0,0 VTSAY STR0006 + ": " + aDados[2][2] // "Centro de Trabalho"
	Else
	   @0,0 VTSAY STR0005 + ": " + aDados[1][2] // "Área de Produção"
	EndIf
	@1,0 VTSAY STR0007+":" VTGET aDados[3][2] Pict "@!" F3 "CYN002" Valid ValidCampo(@aDados[3][2], 3, aDados[3][1] ) // "Motivo Parada"
	VTREAD()
	//@2,0 VTSAY cDenMotivo
	If VTLastkey() == 27    // Tecla ESC
	   Return Nil
	EndIf
	
	@2,0 VTSAY STR0016+":" VTGET aDados[4][2] Pict "99/99/9999" Valid ValidCampo(@aDados[4][2], 5, aDados[4][1] ) // "Data Início"
	VTREAD()
	If VTLastkey() == 27    // Tecla ESC
	   Return Nil
	EndIf
	@3,0 VTSAY STR0017+":" VTGET aDados[5][2] Pict "99:99:99" Valid ValidCampo(@aDados[5][2], 6, aDados[5][1] ) // "Hora Início"
	VTREAD()
	If VTLastkey() == 27    // Tecla ESC
	   Return Nil
	EndIf
	
	//Imprime novamente a parte superior da tela, e solicita a data/hora final da parada.
	VTCLEAR()
	VTClearBuffer()
	
	//@0,0 VTSAY STR0001 // "Apontam Parada Geral"
	If len(aCDCETR) > 1
	   @0,0 VTSAY STR0006 + ": " + aDados[2][2] // "Centro de Trabalho"
	Else
	   @0,0 VTSAY STR0005 + ": " + aDados[1][2] // "Área de Produção"
	EndIf
	//@3,0 VTSAY cDenMotivo
	If VTLastkey() == 27    // Tecla ESC
	   Return Nil
	EndIf
	
	@1,0 VTSAY STR0031 VTGET aDados[6][2] Pict "99/99/9999" ;
	                       Valid ValidCampo(@aDados[6][2], 8, aDados[6][1], aDados ) // "Data Fim:"
	VTREAD()
	If VTLastkey() == 27    // Tecla ESC
	   Return Nil
	EndIf
	@2,0 VTSAY STR0032 VTGET aDados[7][2] Pict "99:99:99" ;
	                       Valid ValidCampo(@aDados[7][2], 9, aDados[7][1], aDados ) // "Hora Fim:"
	VTREAD()
	If VTLastkey() == 27    // Tecla ESC
	   Return Nil
	EndIf
	
	@3,0 VTSAY STR0020 VTGET cConfirm Pict "@!" Valid ValidCampo(@cConfirm, 7, '' )
	VTREAD()
	If VTLastkey() == 27    // Tecla ESC
	   Return Nil
	EndIf

else

	// Monta a tela
	VTCLEAR()
	VTClearBuffer()
	
	@0,0 VTSAY STR0001 // "Apontam Parada Geral"
	@2,0 VTSAY STR0005+":" // "Área de Produção"
	@3,0 VTGET aDados[1][2] Pict "@!" F3 "CYA001" Valid ValidCampo(@aDados[1][2], 1, aDados[1][1] )
	VTREAD()
	If VTLastkey() == 27  // Tecla ESC
	   Return Nil
	Else
	   aCDCETR := GetCDCETR(aDados[1][2], @aDados)
	   cArea   := Iif(Len(aCDCETR)>0,aCDCETR[1],"")
	EndIf
	
	// Monta a tela
	If len(aCDCETR) > 1
	   aDados[2][2] := InicPadrao(2)
	
	   @4,0 VTSAY STR0006+":" // "Centro de Trabalho"
	   @5,0 VTGET aDados[2][2] Pict "@!" F3 "CYI002" Valid ValidCampo(@aDados[2][2], 2, aDados[2][1] )
	   VTREAD()
	   If VTLastkey() == 27    // Tecla ESC
	      Return Nil
	   EndIf
	EndIf
	
	// Monta a tela
	VTCLEAR()
	VTClearBuffer()
	
	@0,0 VTSAY STR0001 // "Apontam Parada Geral"
	If len(aCDCETR) > 1
	   @2,0 VTSAY STR0006 + ": " + aDados[2][2] // "Centro de Trabalho"
	Else
	   @2,0 VTSAY STR0005 + ": " + aDados[1][2] // "Área de Produção"
	EndIf
	@3,0 VTSAY STR0007+":" VTGET aDados[3][2] Pict "@!" F3 "CYN002" Valid ValidCampo(@aDados[3][2], 3, aDados[3][1] ) // "Motivo Parada"
	VTREAD()
	@4,0 VTSAY cDenMotivo
	If VTLastkey() == 27    // Tecla ESC
	   Return Nil
	EndIf
	
	@5,0 VTSAY STR0016+":" VTGET aDados[4][2] Pict "99/99/9999" Valid ValidCampo(@aDados[4][2], 5, aDados[4][1] ) // "Data Início"
	VTREAD()
	If VTLastkey() == 27    // Tecla ESC
	   Return Nil
	EndIf
	@6,0 VTSAY STR0017+":" VTGET aDados[5][2] Pict "99:99:99" Valid ValidCampo(@aDados[5][2], 6, aDados[5][1] ) // "Hora Início"
	VTREAD()
	If VTLastkey() == 27    // Tecla ESC
	   Return Nil
	EndIf
	
	//Imprime novamente a parte superior da tela, e solicita a data/hora final da parada.
	VTCLEAR()
	VTClearBuffer()
	
	@0,0 VTSAY STR0001 // "Apontam Parada Geral"
	If len(aCDCETR) > 1
	   @2,0 VTSAY STR0006 + ": " + aDados[2][2] // "Centro de Trabalho"
	Else
	   @2,0 VTSAY STR0005 + ": " + aDados[1][2] // "Área de Produção"
	EndIf
	@3,0 VTSAY cDenMotivo
	If VTLastkey() == 27    // Tecla ESC
	   Return Nil
	EndIf
	
	@5,0 VTSAY STR0031 VTGET aDados[6][2] Pict "99/99/9999" ;
	                       Valid ValidCampo(@aDados[6][2], 8, aDados[6][1], aDados ) // "Data Fim:"
	VTREAD()
	If VTLastkey() == 27    // Tecla ESC
	   Return Nil
	EndIf
	@6,0 VTSAY STR0032 VTGET aDados[7][2] Pict "99:99:99" ;
	                       Valid ValidCampo(@aDados[7][2], 9, aDados[7][1], aDados ) // "Hora Fim:"
	VTREAD()
	If VTLastkey() == 27    // Tecla ESC
	   Return Nil
	EndIf
	
	@7,0 VTSAY STR0020 VTGET cConfirm Pict "@!" Valid ValidCampo(@cConfirm, 7, '' )
	VTREAD()
	If VTLastkey() == 27    // Tecla ESC
	   Return Nil
	EndIf
endif

// Realiza a gravacao
If cConfirm $ "SY"
   AUTO317( aDados )
Endif
Return Nil

//-------------------------------------------------------------------
/*/{Protheus.doc} AUTO317
Método para gravação dos dados no SFC

@param   aDados  Array contendo os dados que serão enviados

@author  Ana Carolina Tome Klock
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function AUTO317( aDados )
   Local oModel, nI
   Default aDados := {}

   // Instancia o modelo
   oModel := FWLoadModel( 'SFCA317' )
   oModel:SetOperation( MODEL_OPERATION_INSERT )
   oModel:Activate()

   // Preenche os dados

   aDados[4][2] := CTOD( aDados[4][2] )
   aDados[6][2] := CTOD( aDados[6][2] )
   For nI := 1 to len(aDados)
      If aDados[nI,1] == "CZ7_CDSP"
         oModel:SetValue('CZ7MASTER', aDados[nI,1], AllTrim(aDados[nI,2]) )
      Else
         oModel:SetValue('CZ7MASTER', aDados[nI,1], aDados[nI,2] )
      Endif
   Next

   // Valida o modelo
   If oModel:VldData()
      If !oModel:CommitData()
         VTClear()
         Return .F.
      EndIf
      VTClear()
      VTAlert( STR0033,STR0022) //"Apontamento de parada geral realizado com sucesso." , "Sucesso"
   Else
      aErro := oModel:GetErrorMessage()
      VTClear()
      VTAlert( STR0015 +oModel:GetErrorMessage()[6], oModel:GetErrorMessage()[5]) // 'Erro: '
      Return .F.
   EndIf

   // Desativa o modelo
   oModel:DeActivate()
Return .T.

//-------------------------------------------------------------------
/*/{Protheus.doc} GetCDCETR
Preenche uma vetor com todas os centros de trabalho da maquina

@param   cCDARPO  Codigo da Area de Producao selecionada
@param   aDados   Array com os dados informados em tela.
@return  aRet     Retorna um vetor com o codigo de todos os centros

@author  Ana Carolina Tome Klock
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function GetCDCETR( cCDARPO, aDados )
Local aRet := {}
Local cCT  := ""

dbSelectArea( 'CYI' )
CYI->( dbSetOrder(2) )
If CYI->( dbSeek( xFilial('CYI')+cCDARPO ) )
   While CYI->( !EOF() ) .And. CYI->(CYI_FILIAL+CYI_CDARPO) == xFilial('CYI')+cCDARPO
      aAdd( aRet, CYI->CYI_CDARPO )
      cCT := CYI->CYI_CDCETR
      CYI->( dbSkip() )
   End
EndIf
CYI->(dbCloseArea())
If Len(aRet) == 1
   aDados[2,2] := cCT
EndIf
Return aRet

//-------------------------------------------------------------------
/*/{Protheus.doc} ACDV317FIL
Realiza o filtro na Consulta Padrão especificada

@return  lRet   Retorna .T. ou .F. conforme filtro

@author  Ana Carolina Tome Klock
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Function ACDV317FIL()
Local lRet := .F.

If IsInCallStack( 'ACDV317I' )
   If aScan(aCDCETR, { |x| AllTrim(x) ==  AllTrim(CYI->CYI_CDCETR) } )  > 0
      lRet := .T.
   EndIf
EndIf

Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} ValidCampo
Valida se o valor informado é valido para o campo

@param   cCodigo Codigo que sera validado
@param   nBusca  Identifica o que sera validado
                  1 = Area de Producao
                  2 = Centro de Trabalho

@return  lRet  Retorna se o codigo e valido ou nao

@author  Ana Carolina Tome Klock
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function ValidCampo( cCodigo, nBusca, cTabela, aDados )
   Local lRet       := .T.
   Local cPontoEntr := 'ACDV317AP' // Ponto de Entrada de Area de Producao
   Default nBusca   := 1

   // Ponto de Entrada de Centro de Trabalho
   If nBusca == 2
      cPontoEntr := 'ACDV317CT'
   // Ponto de Entrada de Parada
   ElseIf nBusca == 3
      cPontoEntr := 'ACDV317SP'
   // Ponto de Entrada de Parada Geral
   ElseIf nBusca == 4
      cPontoEntr := 'ACDV317PG'
   EndIf

   // Executa ponto de entrada
   If ExistBlock(cPontoEntr)
      cCodigo := ExecBlock(cPontoEntr,.F.,.F.)
      If ValType(cCodigo) # "C"
         cCodigo := ""
      EndIf
   EndIf

   // Valida Area de Producao
   If nBusca == 1
      dbSelectArea( 'CYA' )
      CYA->( dbSetOrder(1) )
      If CYA->( !dbSeek( xFilial('CYA')+cCodigo ))
         VTAlert( STR0011, STR0009 ) // "Área de Produção inválida" ### "Erro"
         lRet := .F.
      EndIf

   // Valida Centro de Trabalho
   ElseIf nBusca == 2
      dbSelectArea( 'CYI' )
      CYI->( dbSetOrder(1) )
      //Se a informação de Centro de Trabalho estiver vazia, atribui a parada à Área
      If !Empty(cCodigo) .And. CYI->( !dbSeek( xFilial('CYI')+cCodigo ))
         VTAlert( STR0012, STR0009 ) // "Centro de Trabalho inválido" ### "Erro"
         lRet := .F.
      EndIf

   // Valida Parada
   ElseIf nBusca == 3
      dbSelectArea( 'CYN' )
      CYN->( dbSetOrder(1) )
      If CYN->( !dbSeek( xFilial('CYN')+cCodigo ) )
         VTAlert( STR0013, STR0009 ) // "Parada geral inválida" ### "Erro"
         lRet := .F.
      Else
         cDenMotivo := CYN->CYN_DSSP
      EndIf

   // Valida Parada
   ElseIf nBusca == 4
      dbSelectArea( 'CZ7' )
      CZ7->( dbSetOrder(1) )
      If CZ7->( !dbSeek( xFilial('CZ7')+cCodigo ) )
         VTAlert( STR0013, STR0009 ) // "Parada geral inválida" ### "Erro"
         lRet := .F.
      ElseIf CZ7->CZ7_DTEDSP != CTOD('31/12/9999')
         VTAlert( STR0014, STR0009 ) // "Parada geral finalizada" ### "Erro"
         lRet := .F.
      EndIf

   // Valida Data
   ElseIf nBusca == 5
      If Day ( CTOD(cCodigo) ) == 0
         VTAlert( STR0026, STR0009 ) //"Data inválida."
         lRet := .F.
      EndIf

   // Valida Hora
   ElseIf nBusca == 6
      aHora := StrTokArr( cCodigo, ':' )
      If len(aHora) < 3 .Or. !ValHora(aHora[1]) .Or. !ValHora(aHora[2]) .Or. !ValHora(aHora[3])
         VTAlert( STR0027, STR0009 ) //"Hora inválida."
         lRet := .F.
      ElseIf Val(aHora[1]) > 23 .Or. Val(aHora[2]) > 59 .Or. Val(aHora[3]) > 59
         VTAlert( STR0027, STR0009 ) //"Hora inválida."
         lRet := .F.
      EndIf

   ElseIf nBusca == 7
      If !(cCodigo $ "SNY")
         VTAlert(STR0025, STR0009 ) //"Opção inválida para confirmação."
         lRet := .F.
      Endif
   ElseIf nBusca == 8
      //Data final
      If Day( CTOD(cCodigo) ) == 0
         VTAlert( STR0026, STR0009 ) //"Data inválida."
         lRet := .F.
      EndIf
      If lRet .And. cToD(cCodigo) < cToD(aDados[4,2])
         VTAlert(STR0030, STR0009) //"Data final não pode ser inferior a data inicial."
         lRet := .F.
      EndIf

   ElseIf nBusca == 9
      //Hora final
      aHora := StrTokArr( cCodigo, ':' )
      If len(aHora) < 3 .Or. !ValHora(aHora[1]) .Or. !ValHora(aHora[2]) .Or. !ValHora(aHora[3])
         VTAlert( STR0027, STR0009 ) //"Hora inválida."
         lRet := .F.
      ElseIf Val(aHora[1]) > 23 .Or. Val(aHora[2]) > 59 .Or. Val(aHora[3]) > 59
         VTAlert( STR0027, STR0009 ) //"Hora inválida."
         lRet := .F.
      EndIf
      If lRet .And. aDados[4,2] == aDados[6,2] .And. cCodigo <= aDados[5,2]
         VTAlert(STR0029, STR0009) //"Hora final deve ser posterior a hora inicial."
         lRet := .F.
      EndIf
   EndIf

   //Remove os espaços sobrando
   If nBusca != 7
      cCodigo := AllTrim(cCodigo)
   EndIf
   //Adiciona a quantidade correta de espaços no final da string(conforme o campo no banco de dados)
   If cTabela != ""
      cCodigo := cCodigo + SPACE(TamSX3(cTabela)[1] - Len(cCodigo))
   EndIf

Return lRet

Static Function ValHora(cHora)
   Local nI := 0

   For nI := 1 To Len(cHora)
      If !(SUBSTR(cHora,nI,1) $ "1234567890")
         Return .F.
      Endif
   Next
Return .T.

//-------------------------------------------------------------------
/*/{Protheus.doc} InicPadrao
Retorna o unico valor valido como valor padrao

@param   nBusca  Identifica o que sera buscado
                  1 = Area de Producao
                  2 = Parada

@return  cRet    Retorna o codigo desejado

@author  Ana Carolina Tome Klock
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function InicPadrao( nBusca )
Local cRet     := ''
Default nBusca := 1

// Inicializador Padrao da Area de Producao
If nBusca == 1
   dbSelectArea( 'CYA' )
   CYA->( dbSeek(xFilial('CYA')) )
   While CYA->( !EOF() ) .And. CYA->CYA_FILIAL == xFilial('CYA')
      If Empty(cRet)
         cRet := CYA->CYA_CDARPO
      Else
         Return SPACE(TamSX3('CYA_CDARPO')[1]+1) //+1 para não mudar de campo automaticamente ao terminar de digitar
      EndIf
      CYA->( dbSkip() )
   End

// Inicializador Padrao do Centro de Trabalho
ElseIf nBusca == 2
   If len(aCDCETR) == 1
      cRet := aCDCETR[1]
   Else
      Return SPACE(TamSX3('CYI_CDCETR')[1]+1) //+1 para não mudar de campo automaticamente ao terminar de digitar
   EndIf

// Inicializador Padrao da Parada
ElseIf nBusca == 3
   dbSelectArea( 'CYN' )
   CYN->( dbSeek(xFilial("CYN")) )
   While CYN->( !EOF() ) .And. CYN->CYN_FILIAL == xFilial("CYN")
      If Empty(cRet)
         cRet := CYN->CYN_CDSP
      Else
         Return SPACE(TamSX3('CYN_CDSP')[1]+1) //+1 para não mudar de campo automaticamente ao terminar de digitar
      EndIf
      CYN->( dbSkip() )
   End
EndIf

Return cRet
//Filtro de Centro de Trabalho por Area
Function ACDV317FTA()
   If CYI->CYI_CDARPO != cArea
      Return .F.
   EndIf
Return .T.