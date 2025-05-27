#INCLUDE 'TOTVS.CH'
#INCLUDE 'PROTHEUS.CH'
#INCLUDE 'APVT100.CH'
#INCLUDE 'ACDV311.CH'

//------------------------------------------------------------------
/*/{Protheus.doc} ACDV311
Tela de Apontamento de Parada

@author  Ana Carolina Tome Klock
@since   01/09/2013
@version P12
/*/
//------------------------------------------------------------------
Function ACDV311()
Local aOpcoes := { STR0002, STR0003, STR0004 } // "Iniciar Parada" ### "Finalizar Parada" ### "Cancelar Parada"
Local nMenu   := -1

If SuperGetMv('MV_INTACD',.F.,"0") != "1"
   VTAlert(STR0027, STR0012) //"Integração com o ACD desativada. Verifique o parâmetro MV_INTACD" ### 'Erro'
   Return
EndIf

// Monta a tela
@0,0 VTSAY STR0001 // "Apont Parada"

// Monta a lista de opcoes
While nMenu < 0
   nMenu := VTAChoice( 2, 0, VTMaxRow(), VTMaxCol(), aOpcoes, , "ACDV311VLD", nMenu )
End

Return Nil

//-------------------------------------------------------------------
/*/{Protheus.doc} ACDV311VLD
Validação da seleção do menu

@param      nModo       Modos do VTAChoice
                        0 - Inativo
                        1 - Tentativa de passar início da lista
                        2 - Tentativa de passar final da lista
                        3 - Normal
                        4 - Itens não selecionados
@param      nPosicao    Item selecionado em tela

@return nReturn Retorna 0 para sair da tela

@author     Ana Carolina Tome Klock
@since      01/09/2013
@version    P12
/*/
//-------------------------------------------------------------------
Function ACDV311VLD( nModo, nPosicao )
Private aCDMQSP    := GetCDMQSP()
Private cDenMotivo := Space(TamSX3("CYN_DSSP")[1])
Private cDenMaq    := Space(TamSX3("CYB_DSMQ")[1])
Private cConfirm   := STR0028 //"SIM"
Default lAutoMacao := .F.

cConfirm := SUBSTR(cConfirm, 1,1)

If nModo == 3
   IF !lAutoMacao
      If VTLastkey() == 27     // Tecla ESC
         Return 0
      Else
         If nPosicao == 1      // Inclusao
            ACDV311I()
         ElseIf nPosicao == 2  // Finalizacao
            ACDV311F()
         ElseIf nPosicao == 3  // Cancelamento
            ACDV311C()
         EndIf
         Return 0
      EndIf
   ENDIF
EndIf

Return -1

//-------------------------------------------------------------------
/*/{Protheus.doc} ACDV311I
Método para inclusão de um apontamento de parada

@author  Ana Carolina Tome Klock
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function ACDV311I()
Local aDados  :={ { 'CYX_CDMQ' , InicPadrao(1) } ,;
                  { 'CYX_CDSP' , InicPadrao(2) } ,;
                  { 'CYX_DTBGSP', DTOC(Date()) } ,;
                  { 'CYX_HRBGSP', Time()       } }

If len(aDados[1][2]) == 0
   VTAlert( STR0018, STR0012 ) // 'Todas as maquinas estao paradas' ### 'Erro'
Else
   // Monta a tela
   VTCLEAR()
   VTClearBuffer()

   @0,0 VTSAY STR0002 // "Iniciar Parada"
   @2,0 VTSAY STR0005+":" VTGET aDados[1][2] Pict "@!" F3 "CYB311" Valid ValidCampo( aDados[1][2], 1 ) // "Máquina"
   VTREAD()
   If VTLastkey() == 27 // Tecla ESC
      Return Nil
   EndIf

   @3,0 VTSAY STR0006+":" VTGET aDados[2][2] Pict "@!" F3 "CYN002" Valid ValidCampo( aDados[2][2], 2 ) // "Motivo Parada"
   VTREAD()
   @4,0 VTSAY cDenMotivo
   If VTLastkey() == 27  // Tecla ESC
      Return Nil
   EndIf

   // Caso não aponte on-line
   If !GetMV('MV_SFCAPON')
      @5,0 VTSAY STR0007+":" VTGET aDados[3][2] Pict "99/99/9999" Valid ValidCampo( aDados[3][2], 3 ) // "Data Início"
      VTREAD()
      If VTLastkey() == 27 // Tecla ESC
         Return Nil
      EndIf
      @6,0 VTSAY STR0008+":" VTGET aDados[4][2] Pict "99:99:99" Valid ValidCampo( aDados[4][2], 4 ) // "Hora Início"
      VTREAD()
      If VTLastkey() == 27  // Tecla ESC
         Return Nil
      EndIf
   EndIf

   @7,0 VTSAY STR0019 VTGET cConfirm Pict "@!" Valid ValidCampo( cConfirm, 5 )
   VTREAD()
   If VTLastkey() == 27  // Tecla ESC
      Return Nil
   EndIf

   // Realiza a gravacao
   If cConfirm $ "SY"
      AUTO311( aDados, 3 )
   Endif
EndIf

Return Nil

//-------------------------------------------------------------------
/*/{Protheus.doc} ACDV311F
Método para finalização de um apontamento de parada

@author  Ana Carolina Tome Klock
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function ACDV311F()
Local aDados :={ { 'CYX_CDMQ' , InicPadrao(1)  } ,;
                 { 'CYX_DTEDSP', DTOC(Date())  } ,;
                 { 'CYX_HREDSP', Time()        } }

If len(aDados[1][2]) == 0
   VTAlert( STR0011, STR0012 ) // 'Sem paradas em andamento!' ### 'Erro'
Else
   // Monta a tela
   VTCLEAR()
   VTClearBuffer()

   @0,0 VTSAY STR0003 // "Finalizar Parada"
   @2,0 VTSAY STR0005+":" VTGET aDados[1][2] Pict "@!" F3 "CYB311" Valid ValidCampo( aDados[1][2], 1 ) // "Máquina"
   VTREAD()
   @3,0 VTSAY cDenMaq

   If VTLastkey() == 27    // Tecla ESC
      Return Nil
   EndIf

   // Caso não aponte on-line
   If !GetMV('MV_SFCAPON')
      @4,0 VTSAY STR0009+":" VTGET aDados[2][2] Pict "99/99/9999" Valid ValidCampo( aDados[2][2], 3 )  // "Data Fim"
      VTREAD()
      If VTLastkey() == 27    // Tecla ESC
         Return Nil
      EndIf
      @5,0 VTSAY STR0010+":" VTGET aDados[3][2] Pict "99:99:99" Valid ValidCampo( aDados[3][2], 4 ) // "Hora Fim"
      VTREAD()
      If VTLastkey() == 27    // Tecla ESC
         Return Nil
      EndIf
   EndIf

   @7,0 VTSAY STR0019 VTGET cConfirm Pict "@!" Valid ValidCampo( cConfirm, 5 )
   VTREAD()
   If VTLastkey() == 27    // Tecla ESC
      Return Nil
   EndIf

   // Realiza a gravacao
   If cConfirm $ "SY"
      AUTO311( aDados, 4 )
   Endif
EndIf

Return 0

//-------------------------------------------------------------------
/*/{Protheus.doc} ACDV311C
Método para cancelamento de um apontamento de parada

@author    Ana Carolina Tome Klock
@since     01/09/2013
@version   P12
/*/
//-------------------------------------------------------------------
Static Function ACDV311C()
Local aDados :={ { 'CYX_CDMQ' , InicPadrao(1) }}

If len(aDados[1][2]) == 0
   VTAlert( STR0011, STR0012 ) // 'Sem paradas em andamento!' ### 'Erro'
Else
   // Monta a tela
   VTCLEAR()
   VTClearBuffer()

   @0,0 VTSAY STR0004 //"Cancelar Parada"
   @2,0 VTSAY STR0005+":" VTGET aDados[1][2] Pict "@!" F3 "CYB311" Valid ValidCampo(aDados[1][2], 1)  // "Máquina"
   VTREAD()
   @3,0 VTSAY cDenMaq

   If VTLastkey() == 27    // Tecla ESC
      Return Nil
   EndIf

   @7,0 VTSAY STR0019 VTGET cConfirm Pict "@!" Valid ValidCampo( cConfirm, 5 )
   VTREAD()
   If VTLastkey() == 27    // Tecla ESC
      Return Nil
   EndIf

   // Realiza a gravacao
   If cConfirm $ "SY"
      AUTO311( aDados, 5 )
   Endif
EndIf

Return 0

//-------------------------------------------------------------------
/*/{Protheus.doc} AUTO311
Método para gravação dos dados no SFC

@param      aDados   Array contendo os dados que serão enviados
@param      nOpc     Tipo da operação que será efetuada
                     3 - Inclusão de Apontamento
                     4 - Finalização de Apontamento
                     5 - Cancelamento de Apontamento

@author  Ana Carolina Tome Klock
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function AUTO311( aDados, nOpc )
Local oModel, nI
Default aDados  := {}
Default nOpc    := 2
Default lAutoMacao := .F.

// Busca o dado para alterar ou excluir
If nOpc == 4 .Or. nOpc == 5
   dbSelectArea( 'CYX' )
   CYX->( dbSetOrder(2) )
   If !CYX->( dbSeek( xFilial('CYX')+aDados[1,2]+DTOS(CTOD('31/12/9999')) ) )
      Return .F.
   EndIf
EndIf

// Instancia o modelo
oModel := FWLoadModel( 'SFCA311' )
oModel:SetOperation( nOpc )
oModel:Activate()

// Preenche os dados
If nOpc == 3
   aDados[3][2] := CTOD( aDados[3][2] )
   For nI := 1 to len(aDados)
      If aDados[nI,1] == "CYX_CDSP"
         oModel:SetValue('CYXMASTER', aDados[nI,1], AllTrim(aDados[nI,2]) )
      Else
         IF !lAutoMacao
            oModel:SetValue('CYXMASTER', aDados[nI,1], aDados[nI,2] )
         ENDIF
      Endif
   Next

ElseIf nOpc == 4
   aDados[2][2] := CTOD( aDados[2][2] )
   For nI := 2 to len(aDados)
      If aDados[nI,1] == "CYX_CDSP"
         oModel:SetValue('CYXMASTER', aDados[nI,1], AllTrim(aDados[nI,2]) )
      Else
         oModel:SetValue('CYXMASTER', aDados[nI,1], aDados[nI,2] )
      Endif
   Next
EndIf

// Valida o modelo
If oModel:VldData()
   If !oModel:CommitData()
      IF !lAutoMacao
         VTClear()
      ENDIF
      Return .F.
   EndIf
   If nOpc == 3
      VTClear()
      VTAlert( STR0020, STR0023 ) //"Apontamento de parada efetuado." / "Sucesso"
   Endif
   If nOpc == 4
      VTClear()
      VTAlert( STR0021, STR0023 ) //"Finalização do apontamento de parada efetuada." / "Sucesso"
   Endif
   If nOpc == 5
      VTClear()
      VTAlert( STR0022, STR0023 ) //"Cancelamento de apontamento de parada efetuado." / "Sucesso"
   Endif
Else
   aErro := oModel:GetErrorMessage()
   VTClear()
   VTAlert( STR0013 + oModel:GetErrorMessage()[6], oModel:GetErrorMessage()[5]) // "Erro: "
   Return .F.
EndIf

// Desativa o modelo
oModel:DeActivate()

Return .T.

//-------------------------------------------------------------------
/*/{Protheus.doc} GetCDMQSP
Preenche uma vetor com todas as maquinas paradas no momento

@return  aRet   Retorna um vetor com o codigo de todas as maquinas paradas

@author  Ana Carolina Tome Klock
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function GetCDMQSP()
Local aRet := {}

dbSelectArea( 'CYX' )
CYX->( dbGoTop() )
While CYX->( !EOF() )
   If CYX->CYX_DTEDSP == CTOD('31/12/9999')
      aAdd( aRet, CYX->CYX_CDMQ )
   EndIf
   CYX->( dbSkip() )
End
CYX->(dbCloseArea())

Return aRet

//-------------------------------------------------------------------
/*/{Protheus.doc} ACDV311FIL
Realiza o filtro na Consulta Padrão especificada

@return  lRet   Retorna .T. ou .F. conforme filtro

@author  Ana Carolina Tome Klock
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Function ACDV311FIL()
Local lRet := .F.

If aScan(aCDMQSP, { |x| AllTrim(x) ==  AllTrim(CYB->CYB_CDMQ) } )  > 0
   If IsInCallStack('ACDV311F') .Or. IsInCallStack('ACDV311C')
      lRet := .T.
   EndIf
Else
   If IsInCallStack('ACDV311I')
      lRet := .T.
   EndIf
EndIf

Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} ValidCampo
Valida se o valor informado é valido para o campo

@param       cCodigo   Codigo que sera validado
@param       nBusca      Identifica o que sera validado
                  1 = Maquina
                  2 = Parada
                  3 = Data
                  4 = Hora

@return  lRet  Retorna se o codigo e valido ou nao

@author  Ana Carolina Tome Klock
@since   01/09/2013
@version  P12
/*/
//-------------------------------------------------------------------
Static Function ValidCampo( cCodigo, nBusca )
Local lRet         := .T.
Local cPontoEntr   := 'ACDV311MQ' // Ponto de Entrada de Maquina
Default nBusca     := 1
Default lAutoMacao := .F.

// Ponto de Entrada de Parada
If nBusca == 2
   cPontoEntr := 'ACDV311SP'
EndIf

// Executa ponto de entrada
If ExistBlock(cPontoEntr)
   cCodigo := ExecBlock(cPontoEntr,.F.,.F.)
   If ValType(aButtons) # "C"
      cCodigo := ""
   EndIf
EndIf

// Valida Máquina
If nBusca == 1
   dbSelectArea( 'CYB' )
   CYB->( dbSetOrder(1) )
   If CYB->( dbSeek( xFilial('CYB')+cCodigo ))
      If IsInCallStack( 'ACDV311I' )
         If aScan(aCDMQSP, { |x| AllTrim(x) ==  AllTrim(cCodigo) } )  > 0
            VTAlert( STR0014, STR0012 ) // "Máquina já possui parada em andamento" ### "Erro"
            lRet := .F.
         EndIf
      ElseIf IsInCallStack('ACDV311F') .Or. IsInCallStack('ACDV311C')
         If aScan(aCDMQSP, { |x| AllTrim(x) ==  AllTrim(cCodigo) } )  == 0
            VTAlert( STR0015, STR0012 ) // "Máquina não possui parada em andamento" ## "Erro"
            lRet := .F.
          EndIf
      EndIf
      cDenMaq := CYB->CYB_DSMQ
   Else
      VTAlert( STR0016, STR0012 ) // "Máquina não cadastrada" ### "Erro"
      lRet := .F.
   EndIf

// Valida Parada
ElseIf nBusca == 2
   If Len(AllTrim(cCodigo)) > TamSX3("CYN_CDSP")[1]
      VTAlert( STR0017, STR0012 ) // "Motivo Parada não cadastrado" ### "Erro"
      lRet := .F.
   Endif

   If lRet
      dbSelectArea( 'CYN' )
      CYN->( dbSetOrder(1) )
      If CYN->( !dbSeek( xFilial('CYN')+cCodigo ) )
         IF !lAutoMacao
            VTAlert( STR0017, STR0012 ) // "Motivo Parada não cadastrado" ### "Erro"
         ENDIF
         lRet := .F.
      Else
         cDenMotivo := CYN->CYN_DSSP
      EndIf
   Endif

// Valida Data
ElseIf nBusca == 3
   If Day ( CTOD(cCodigo) ) == 0
      VTAlert( STR0025, STR0012 ) //"Data inválida."
      lRet := .F.
   EndIf

// Valida Hora
ElseIf nBusca == 4
   aHora := StrTokArr( cCodigo, ':' )
   If len(aHora) < 3 .Or. !ValHora(aHora[1]) .Or. !ValHora(aHora[2]) .Or. !ValHora(aHora[3])
      VTAlert( STR0026, STR0012 ) //"Hora inválida."
      lRet := .F.
   ElseIf Val(aHora[1]) > 23 .Or. Val(aHora[2]) > 59 .Or. Val(aHora[3]) > 59
      VTAlert( STR0026, STR0012 ) //"Hora inválida."
      lRet := .F.
   EndIf

// Valida Confirmação
ElseIf nBusca == 5
   If !(cCodigo $ "SNY")
      VTAlert( STR0024, STR0009 ) //"Opção inválida para confirmação."
      lRet := .F.
   Endif
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

@param       nBusca      Identifica o que sera buscado
                  1 = Maquina
                  2 = Parada

@return  cRet   Retorna o codigo desejado

@author  Ana Carolina Tome Klock
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function InicPadrao( nBusca )
Local cRet     := ''
Default nBusca := 1

// Inicializador Padrao da Maquina
If nBusca == 1

   // Maquinas nao paradas
   If IsInCallStack('ACDV311I')
      dbSelectArea( 'CYB' )
      CYB->( dbSeek(xFilial("CYB")) )
      While CYB->( !EOF() ) .And. CYB->CYB_FILIAL == xFilial("CYB")
         If aScan(aCDMQSP, { |x| AllTrim(x) ==  AllTrim(CYB->CYB_CDMQ) } )  == 0
            If Empty(cRet)
               cRet := CYB->CYB_CDMQ
            ElseIf !Empty(cRet)
               Return SPACE(TamSX3('CYB_CDMQ')[1])
            EndIf
         EndIf
         CYB->( dbSkip() )
      End

   // Maquinas paradas
   ElseIf IsInCallStack('ACDV311F') .Or. IsInCallStack('ACDV311C')
      If len(aCDMQSP) == 1
         cRet := aCDMQSP[1]
      ElseIf len(aCDMQSP) > 1
         cRet := SPACE(TamSX3("CYB_CDMQ")[1])
      EndIf
   EndIf

// Inicializador Padrao da Parada
ElseIf nBusca == 2
   // Busca as paradas
   dbSelectArea( 'CYN' )
   If CYN->( dbSeek(xFilial("CYN")) )
	   While CYN->( !EOF() ) .And. CYN->CYN_FILIAL == xFilial("CYN")
	      If Empty(cRet)
	         cRet := CYN->CYN_CDSP
	      Else
	         Return SPACE(TamSX3("CYN_CDSP")[1]+1)
	      EndIf
	      CYN->( dbSkip() )
	   End
	Else
		Return SPACE(TamSX3("CYN_CDSP")[1]+1)
	EndIf
EndIf

Return cRet
