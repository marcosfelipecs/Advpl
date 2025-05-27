#INCLUDE 'TOTVS.CH'
#INCLUDE 'PROTHEUS.CH'
#INCLUDE 'APVT100.CH'
#INCLUDE 'ACDV312.CH'

//-----------------------------------------------------------------
/*/{Protheus.doc} ACDV312
Tela de Apontamento de Item Controle

@author  Ana Carolina Tome Klock
@since   01/09/2013
@version P12
/*/
//-----------------------------------------------------------------
Function ACDV312()
Local aDados  :={ { 'CYT_CDMQ', InicPadrao(1) } ,;
                  { 'CYT_DTRP', Date()        } ,;
                  { 'CYT_HRRP', Time()        } }
Local xAux    := Space(250)
Local cTabela := Space(TamSX3('CY2_DSRS')[1])
Local nI
Local nX := 1
Local aItControl := {}
Local cTipo
Default lAutoMacao := .F.

IF !lAutoMacao
   // Monta a tela
   VTCLEAR()
   VTClearBuffer()
ENDIF

If SuperGetMv('MV_INTACD',.F.,"0") != "1"
   IF !lAutoMacao
      VTAlert(STR0024, STR0005) //"Integração com o ACD desativada. Verifique o parâmetro MV_INTACD" ### 'Erro'
   ENDIF
   Return
EndIf

// Valida se existem maquinas com itens
If len(aDados[1][2]) == 0
   VTAlert( STR0006, STR0005 ) // "Sem maquinas com itens controle" ### "Erro")
Else
   @0,0 VTSAY STR0001 // "Apontam Item Controle"
   @2,0 VTSAY STR0002 // "Máquina"
   @3,0 VTGET aDados[1][2] Pict "@!" F3 "CYB312" VALID ValidCampo(aDados[1][2], 1)
   VTREAD()
   If VTLastkey() == 27    // Tecla ESC
      Return Nil
   EndIf
EndIf

dbSelectArea('CYE')
CYE->( dbSetOrder(1) )
If CYE->( dbSeek(xFilial('CYE')+aDados[1][2]) )
   While CYE->( !EOF() ) .And. CYE->CYE_CDMQ == aDados[1][2]
      aAdd(aItControl,CYE->CYE_CDVF)
      CYE->(dbSkip())
   End
Endif

VTClearBuffer()

While nX <= Len(aItControl)

   If nX < 1
      If VTYesNo(STR0015,"ACDV312") //"Deseja cancelar o apontamento de item controle?"
         Return Nil
      Else
         nX := 1
      Endif
   Endif

   // Monta a tela
   VTCLEAR()

   @0,0 VTSAY STR0001 // "Apont Item Controle"
   @2,0 VTSAY STR0003 + AllTrim(aItControl[nX]) // "Item Controle: "

   dbSelectArea('CYE')
   CYE->( dbSetOrder(1) )
   CYE->( dbSeek(xFilial('CYE')+aDados[1][2]+aItControl[nX]) )

   dbSelectArea('CYK')
   CYK->( dbSetOrder(1) )
   If CYK->( dbSeek( xFilial('CYK')+CYE->CYE_CDVF ) )
      @3,0 VTSAY AllTrim(CYK->CYK_DSVF)
      //Verifica o tipo do item controle e informa na tela
      Do Case
         Case AllTrim(CYK->CYK_TPRS) == "1"
            cTipo := STR0018 //Númerico
         Case AllTrim(CYK->CYK_TPRS) == "2"
            cTipo := STR0019 //Tabela
         Case AllTrim(CYK->CYK_TPRS) == "3"
            cTipo := STR0020 //Observação
         Case AllTrim(CYK->CYK_TPRS) == "4"
            cTipo := STR0021 //Data
         OtherWise
            cTipo := STR0022 //Nenhum Tipo Valido
      EndCase
      @4,0 VTSAY AllTrim(cTipo)

      If CYK->CYK_TPRS == '2' // Tabela Resultado
         nI := 1
         While nI <= CYE->CYE_NRMX
            @5,0 VTSAY STR0004 VTGET cTabela Pict "@!" F3 "CY2001" Valid ValidItemC( cTabela, CYE->CYE_CDVF, @aDados, CYK->CYK_CDTB, nX )  // "Valor: "
            VTREAD()

            // Validacoes da Tabela Resultado
            If CYE->CYE_LGVFOB .And. Empty(cTabela) .And. nI == 1
               VTAlert( STR0007, STR0005 ) // "Item Controle Obrigatório" ### "Erro"
            ElseIf Empty(cTabela)
               nI := CYE->CYE_NRMX+1
            Else
               If AScan(StrTokArr(xAux,";"), AllTrim(cTabela)) == 0
                  xAux := AllTrim(xAux) + AllTrim(cTabela) + ";"
                  cTabela := Space(TamSX3('CY2_DSRS')[1])
                  nI++
               Else
                  VTAlert( STR0007, STR0005 ) // "Item Controle Obrigatório" ### "Erro"
               EndIf
            EndIf

            // Adiciona resultados ao aDados
            If nI > CYE->CYE_NRMX
               xAux := Substr( xAux, 0, len(xAux)-1 )
               DelItContr(@aDados,nX)
               aAdd( aDados, {{ 'CYT_CDVF', CYE->CYE_CDVF }, {'CYT_DSRS', xAux }, {'nX',nX} } )
            EndIf
            VTClearBuffer()
         End
      Else
         @5,0 VTSAY STR0004 VTGET xAux Pict "@!" VALID ValidItemC( xAux, CYE->CYE_CDVF, @aDados, , nX ) // "Valor: "
         VTREAD()
      EndIf
   EndIf

   // Limpa variavel auxiliar
   xAux := Space(250)

   If VTLastkey() == 27    // Tecla ESC
      nX--
   Else
      nX++
   EndIf
End

// Realiza a gravacao
If len(aDados) > 3
   AUTO312( aDados )
Else
   VTAlert( STR0008, STR0005 ) // "Nenhum item controle informado" ### "Erro"
EndIf

Return Nil

//-------------------------------------------------------------------
/*/{Protheus.doc} AUTO312
Método para a inclusão dos dados no SFC

@param   aDados  Array contendo os dados que serão enviados

@author  Ana Carolina Tome Klock
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function AUTO312( aDados )
Local oModel, nI, nCount
Default aDados := {}
Default lAutoMacao := .F.

// Instancia o modelo
oModel := FWLoadModel( 'SFCA312' )
oModel:SetOperation( 3 )
oModel:Activate()

// Preenche os dados
For nI := 1 to 3
   IF !lAutoMacao
      oModel:SetValue('CYTMASTER', aDados[nI,1], aDados[nI,2] )
   ENDIF
Next

// Preenche as linhas
For nI := 4 to len(aDados)
   nCount := 1
   While nCount <= oModel:GetModel('CYTDETAIL'):Length()
      If oModel:GetModel('CYTDETAIL'):GetValue(aDados[nI,1,1], nCount) == aDados[nI,1,2]
         oModel:GetModel('CYTDETAIL'):GoLine(nCount)
         If AllTrim(aDados[nI,2,1]) == "CYT_DSRS"
            oModel:LoadValue('CYTDETAIL', aDados[nI,2,1], aDados[nI,2,2] )
         Else
            oModel:SetValue('CYTDETAIL', aDados[nI,2,1], aDados[nI,2,2] )
         EndIf
         Exit
      EndIf
      nCount++
   End
Next

// Valida o modelo
If oModel:VldData()
   IF !lAutoMacao
      VtClear()
      If !oModel:CommitData()
         VtClear()
         Return .F.
      EndIf
      VTAlert( STR0016, STR0017 ) //"Apontamento de item controle efetuado com sucesso." ### "Sucesso"
   ENDIF
Else
   IF !lAutoMacao
      VtClear()
      aErro := oModel:GetErrorMessage()
      VTAlert( STR0009 + oModel:GetErrorMessage()[6], oModel:GetErrorMessage()[5]) // "Erro: "
      Return .F.
   ENDIF
EndIf

// Desativa o modelo
oModel:DeActivate()

Return .T.

//-------------------------------------------------------------------
/*/{Protheus.doc} ACDV312FIL
Realiza o filtro na Consulta Padrão especificada

@return  lRet   Retorna .T. ou .F. conforme filtro

@author  Ana Carolina Tome Klock
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Function ACDV312FIL()
Local lRet := .F.

If CYE->(dbSeek(xFilial("CYE")+CYB->CYB_CDMQ))
   lRet := .T.
EndIf

Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} ACDV312CY2
Realiza o filtro na Consulta Padrão CY2001

@return  lRet     Retorna .T. ou .F. conforme filtro

@author  Lucas Konrad França
@since      13/01/2016
@version    P11
/*/
//-------------------------------------------------------------------
Function ACDV312CY2()
Local lRet  := .F.

If CY2->CY2_CDTB == CYK->CYK_CDTB
   lRet := .T.
EndIf

Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} ValidCampo
Valida se o valor informado é valido para o campo

@param   cCodigo Codigo que sera validado
@param   nBusca  Identifica o que sera validado
                 1 = Maquina

@return  lRet    Retorna se o codigo e valido ou nao

@author  Ana Carolina Tome Klock
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function ValidCampo( cCodigo, nBusca )
Local lRet       := .T.
Local cPontoEntr := 'ACDV312MQ' // Ponto de Entrada de Maquina
Default nBusca   := 1

// Executa ponto de entrada
If ExistBlock(cPontoEntr)
   cCodigo := ExecBlock(cPontoEntr,.F.,.F.)
   If ValType(cCodigo) # "C"
      cCodigo := ""
   EndIf
EndIf

// Valida Máquina
If nBusca == 1
   dbSelectArea( 'CYE' )
   CYE->( dbSetOrder(1) )
   If CYE->( !dbSeek( xFilial('CYB')+cCodigo ))
      VTAlert( STR0010, STR0005 ) // "Máquina sem item controle" ### "Erro"
      lRet := .F.
   EndIf
EndIf

Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} ValidItemC
Valida o item controle informado para a maquina

@param   xAux    Valor informado para o item controle
@param   cCDVF   Codigo do item controle a ser validado
@param   aDados  Vetor onde sera inserido caso esteja valido
@param   cTabela Tabela utilizada quando o tipo for Tabela Resultado

@return  cRet   Retorna se o codigo e valido ou nao

@author  Ana Carolina Tome Klock
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function ValidItemC( xAux, cCDVF, aDados, cTabela, nX )
Local lRet  := .T.
Local nI    := 0

dbSelectArea( 'CYE' )
CYE->( dbSetOrder(1) )
If CYE->( dbSeek( xFilial('CYE')+aDados[1][2]+cCDVF ) )
   If CYE->CYE_LGVFOB .And. Empty(xAux)
      VTAlert( STR0007, STR0005 ) // "Item Controle obrigatório" ### "Erro"
      lRet := .F.
   Else
      dbSelectArea( 'CYK' )
      CYK->( dbSetOrder(1) )
      If CYK->( dbSeek( xFilial('CYK')+cCDVF ) )
         If CYK->CYK_TPRS == '1' // Numerico
            //Verifica se foram informados apenas números.
            For nI := 1 To Len(AllTrim(xAux))
              If !IsDigit(SubStr(xAux,nI,Len(AllTrim(xAux))))
                 VTAlert(STR0023,STR0005) //"Valor digitado não é válido. Digite somente números." ### 'Erro'
                 lRet := .F.
                 Exit
              EndIf
            Next nI

            If lRet
               //Se o valor estiver dentro da faixa de valor máximo e minimo ou não for informado(igual a 0)
               If (Val(xAux) >= CYE->CYE_VLVFMI .And. Val(xAux) <= CYE->CYE_VLVFMX) .Or. Val(xAux) == 0
                  DelItContr(@aDados,nX)
                  aAdd( aDados, {{ 'CYT_CDVF', cCDVF }, {'CYT_VLRS', Val(xAux)}, {'nX', nX}})
               ElseIf !Val(xAux) == 0
                  VTAlert( STR0011 + AllTrim(STR(CYE->CYE_VLVFMI)) + STR0012 +  ;
                          AllTrim(STR(CYE->CYE_VLVFMX)), STR0005 ) // 'Valor deve estar entre ' ### ' e ' ### 'Erro'
                  lRet := .F.
               EndIf
            EndIf
         ElseIf CYK->CYK_TPRS == '2'//Tabela
            lRet := .F.
            dbSelectArea( 'CY2' )
            CY2->( dbSetOrder(2) )
            If CY2->( dbSeek( xFilial('CY2')+cTabela ) )
               While CY2->( !EOF() )
                  If CY2->CY2_CDTB == cTabela .And. (CY2->CY2_DSRS == xAux .Or. Empty(xAux))
                     lRet := .T.
                     Exit
                  EndIf
                  CY2->( dbSkip() )
               End

               If !lRet
                  VTAlert( STR0013, STR0005 ) // 'Valor nao consta na tabela resultado' ### 'Erro'
               EndIf
            EndIf
         ElseIf CYK->CYK_TPRS == '3' .And. !Empty(xAux) // Observacao
            DelItContr(@aDados,nX)
            aAdd( aDados, {{ 'CYT_CDVF', cCDVF }, {'CYT_DSRS', xAux }, {'nX', nX}} )
         ElseIf CYK->CYK_TPRS == '4' // Data
            If    (CTOD(xAux) >= DATE()-CYE->CYE_QTDYMI .And. CTOD(xAux) <= DATE()+CYE->CYE_QTDYMX)
               DelItContr(@aDados,nX)
               aAdd( aDados, {{ 'CYT_CDVF', cCDVF }, {'CYT_DTRS', CTOD(xAux) }, {'nX', nX}} )
            ElseIf !Empty(xAux)
               VTAlert( STR0014 + DTOC(DATE()-CYE->CYE_QTDYMI) + STR0012 + ;
                       DTOC(DATE()+CYE->CYE_QTDYMX), STR0005 ) // 'Data deve estar entre ' ### ' e ' ### 'Erro'
               lRet := .F.
            EndIf
         EndIf
      EndIf
   EndIf
EndIf

Return lRet

//-------------------------------------------------------------------
/*/{Protheus.doc} DelItContr
Excluir elementos do Array na posição informada

@param   aDados  Array de dados
         nX      Posição do Item controle a ser excluído

@return

@author  Ana Carolina Tome Klock
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function DelItContr(aDados,nX)
Local nJ       := 0
Local nNumElem := 0
Local nTotal   := Len(aDados)

For nJ := 4 To nTotal
   If aDados[nJ][3][2] == nX
      aDel(aDados,nJ)
      nNumElem++
   Endif
Next nJ

aSize(aDados,nTotal-nNumElem)
Return

//-------------------------------------------------------------------
/*/{Protheus.doc} InicPadrao
Retorna o unico valor valido como valor padrao

@param   nBusca  Identifica o que sera buscado
                 1 = Maquina

@return  cRet   Retorna o codigo desejado

@author  Ana Carolina Tome Klock
@since   01/09/2013
@version P12
/*/
//-------------------------------------------------------------------
Static Function InicPadrao( nBusca )
Local cRet     := ''
Local cQuery   := ''
Local cAlias   := GetNextAlias()

Default nBusca := 1

// Inicializador Padrao da Maquina
If nBusca == 1

   cQuery := " SELECT COUNT(*) TOTAL "
   cQuery +=   " FROM ( SELECT DISTINCT CYE.CYE_CDMQ "
   cQuery +=            " FROM " + RetSqlName("CYE") + " CYE "
   cQuery +=           " WHERE CYE.CYE_FILIAL = '" + xFilial("CYE") + "' "
   cQuery +=             " AND CYE.D_E_L_E_T_ = ' ' ) t "

   cQuery := ChangeQuery(cQuery)

   dbUseArea( .T., 'TOPCONN', TcGenQry(,,cQuery), cAlias, .F., .F. )
   If (cAlias)->(TOTAL) == 1
      dbSelectArea( 'CYE' )
      CYE->( dbGoTop() )
      cRet := CYE->CYE_CDMQ
   Else
      cRet := SPACE(TamSX3('CYB_CDMQ')[1])
   EndIf
   (cAlias)->(dbCloseArea())
EndIf

Return cRet
