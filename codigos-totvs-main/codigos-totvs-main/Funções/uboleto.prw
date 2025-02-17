#Include "Protheus.ch"
#Include "TopConn.ch"
#Include "rwmake.ch"
#INCLUDE "RPTDEF.CH"
#INCLUDE "FWPrintSetup.ch"
#Include "Directry.ch"
#INCLUDE 'parmtype.ch'
#INCLUDE "APWIZARD.CH"
#include "TOTVS.CH"
/*/{Protheus.doc} BOLETO
  Função BOLETO
  @param nTela    :=  1 - Usa tela com parametros próprios
                      2 - Usa tela com parametros de terceiros
                      3 - Não usar tela, impressão direita 
         aParam  := Se não usar tela para seleção passar os parametros de pergunta de 1 a 26.
         pTipo   := 1 - Impressão impressora
                    2 - Impressão PDF
         pGerArq := 1 - Gerar todos os boletos em um único arquivo, apenas para geração em PDF
                    2 - Gerar os boletos individual por parcelas em vários arquivos, apenas para geração em PDF
            
  @return Não retorna nada
  @author Totvs Nordeste
  @owner Totvs S/A
  @version Protheus 10, Protheus 11
  @since 27/10/2015 
  @sample
// BOLETO - User Function função impressão de boletos dos Bancarios (Genérico)
  U_BOLETO()
  Return
  @obs Rotina de Impressão de Boletoss
  @project 
  @menu \SIGAFIN\Atualização\Específico\Boleto
  @history
  27/10/2015 - Desenvolvimento da Rotina.
/*/
User Function uBOLETO(nTela,aParam,pTipo,pGerArq)
  Local aRegs     := {}
  Local aTitulos  := {}
  Local nLastKey  := 0
  Local nId       := 0
  Local nId1      := 0
  
  Local cParte    := ""
  Local cTamanho  := "M"
  Local cDesc1    := "Este programa tem como objetivo efetuar a impressão do"
  Local cDesc2    := "Boleto de Cobrança com código de barras, conforme os"
  Local cDesc3    := "parâmetros definidos pelo usuário"
  Local cString   := "SE1"
  Local wnrel     := "BOLETO"
  
  Default nTela   := 1
  Private cTpImpBol := GetMv("MV_XTPBOL")
  Private cQualBco := ""
  Private cNossoDg := ""
  Private cStgTipo := "'"
  Private cTpImpre := pTipo    // 1 - Impressora ou 2 - PDF
  Private cGerArq  := pGerArq  // Forma de geração do boleto
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

  nTela      := IIf(ValType(nTela) != "N",1,nTela)  
  cStartPath := AllTrim(cStartPath) + "logo_bancos\"

 // --------------
  For nId := 1 To Len(cTpImpBol)
      cParte := Substr(cTpImpBol,nId,1)
      
      If cParte == ","
         While nId1 < 3
           cStgTipo += " "
           
           nId1++  
         EndDo     
         cStgTipo += "','"
         nId1      := 0
       else
         cStgTipo += Substr(cTpImpBol,nId,1)
         nId1++
      EndIf      
  Next

  While nId1 < 3
    cStgTipo += " "
           
    nId1++  
  EndDo     
  
  cStgTipo += "'"
 // -------------- 
 
  fnCriaSx1(aRegs)
  
  If nTela == 1            // Usa tela com parametro 
    If Pergunte(cPerg,.T.)
       MsgRun("Títulos a Receber","Selecionando registros para processamento",{|| fnCallReg(@aTitulos,@nTela)})

       If Len(aTitulos) > 0		
	      // Monta tela de seleção dos registros que deverão gerar o boleto
          fnCallTela(@aTitulos)
       EndIf   
    EndIf
   elseIf nTela <> 1                      // Usa tela com parametros de terceiros
          mv_par01 := aParam[01]          // Prefixo Inicial
          mv_par02 := aParam[02]          // Prefixo Final
          mv_par03 := aParam[03]          // Numero Inicial
          mv_par04 := aParam[04]          // Numero Final
          mv_par05 := aParam[05]          // Parcela Inicial
          mv_par06 := aParam[06]          // Parcela Final
          mv_par07 := aParam[07]          // Tipo Inicial
          mv_par08 := aParam[08]          // Tipo Final
          mv_par09 := aParam[09]          // Cliente Inicial
          mv_par10 := aParam[10]          // Cliente Final
          mv_par11 := aParam[11]          // Loja Inicial
          mv_par12 := aParam[12]          // Loja Final
          mv_par13 := aParam[13]          // Emissão Inicial
          mv_par14 := aParam[14]          // Emissão Final
          mv_par15 := aParam[15]          // Vencimento Inicial
          mv_par16 := aParam[16]          // Vencimento Final
          mv_par17 := aParam[17]          // Natureza Inicial
          mv_par18 := aParam[18]          // Natureza Final
          mv_par19 := aParam[19]          // Banco
          mv_par20 := aParam[20]          // Agência
          mv_par21 := aParam[21]          // Conta
          mv_par22 := aParam[22]          // Subconta
          mv_par23 := aParam[23]          // Tipo do processo: 1 - Gerar, 2 - Reimpressão ou 3 - Regerar
          mv_par24 := aParam[24]          // Diretório
          mv_par25 := aParam[25]          // Gerar boleto: 1 - Sim ou 2 - Não
          mv_par26 := aParam[26]          // Tipo do boleto: 1 - Reduzido ou 2 - Completo
          mv_par27 := PADR(aParam[27],LEN(SE1->E1_NUMBOR)) // NUmero da borderô
          mv_par28 := PADR(aParam[28],LEN(SE1->E1_NUMLIQ)) // NUmero da borderô
          
          MsgRun("Títulos a Receber","Selecionando registros para processamento",{|| fnCallReg(@aTitulos,@nTela)})
          
          If Len(aTitulos) > 0		
            If nTela == 3         // Impressão sem tela com parametros de terceiros
               RptStatus({|lEnd| ImpBol(aTitulos)}, cTitulo)
             else          
              // Monta tela de seleção dos registros que deverão gerar o boleto
               fnCallTela(@aTitulos)
            EndIf   
          EndIf 
  EndIf   
Return

/*---------------------------------------------
--  Função: Pesquisa títulos para impressão  --
--          de boleto.                       --
-----------------------------------------------*/
Static Function fnCallReg(aTitulos,nTela)
  Local cQuery  := ""
 /*  
  If (Empty(mv_par19) .or. Empty(mv_par20)) .and. mv_par23 == 2
     Aviso("ATENÇÃO","Parametros: Banco e Agência não preenchidos.",{"OK"})
     Return
  EndIf
 
 // --- Validar se escolher o boleto reduzido sem cliente
  If mv_par26 == 1 .and. Empty(Alltrim(mv_par09))
     Aviso("ATENÇÃO","Parametros: Boleto reduzido escolhido sem o cliente.",{"OK"})
     Return
   elseIf mv_par26 == 1
          mv_par10 := mv_par09
          mv_par12 := mv_par11
  EndIf
 // -----------------------------------------------------
  */
  cQuery := " Select SE1.E1_PREFIXO, SE1.E1_NUM, SE1.E1_PARCELA, SE1.E1_TIPO, SE1.E1_NATUREZ,"
  cQuery += "        SE1.E1_CLIENTE, SE1.E1_LOJA, SE1.E1_NOMCLI, SE1.E1_EMISSAO, SE1.E1_VENCTO,"
  cQuery += "        SE1.E1_VENCREA, SE1.E1_VALOR, SE1.E1_HIST, SE1.E1_PORTADO, SE1.E1_AGEDEP, SE1.E1_CONTA,"
  cQuery += "        SE1.E1_XSUBCTA, SE1.E1_NUMBCO, R_E_C_N_O_ AS E1_REGSE1 "
  cQuery += "  from " + RetSqlName("SE1") + " SE1 "
  
  If Empty(mv_par27)
    cQuery += "    Where SE1.E1_FILIAL = '" + xFilial("SE1") + "'"
    cQuery += "      and SE1.E1_PREFIXO between '" + mv_par01 + "' and '" + mv_par02 + "'"
    cQuery += "      and SE1.E1_NUM     between '" + mv_par03 + "' and '" + mv_par04 + "'"
    cQuery += "      and SE1.E1_PARCELA between '" + mv_par05 + "' and '" + mv_par06 + "'"
    cQuery += "      and SE1.E1_TIPO    between '" + mv_par07 + "' and '" + mv_par08 + "'"
    cQuery += "      and SE1.E1_CLIENTE between '" + mv_par09 + "' and '" + mv_par10 + "'"
    cQuery += "      and SE1.E1_LOJA    between '" + mv_par11 + "' and '" + mv_par12 + "'"
    cQuery += "      and SE1.E1_EMISSAO between '" + DToS(mv_par13) + "' and '" + DToS(mv_par14) + "'"
    cQuery += "      and SE1.E1_VENCTO  between '" + DToS(mv_par15) + "' and '" + DToS(mv_par16) + "'"
    cQuery += "      and SE1.E1_NATUREZ between '" + mv_par17 + "' and '" + mv_par18 + "'"
    cQuery += "      and SE1.E1_SALDO > 0"
    
  Else
    cQuery += "      Where SE1.E1_NUMBOR = '"+mv_par27+"' "
  EndIf
  cQuery += "      and SE1.E1_TIPO in (" + cStgTipo + ") "

  If !Empty(mv_par28)
    cQuery += "      and SE1.E1_NUMLIQ = '"+mv_par28+"' "
  EndIf

  If mv_par23 == 2
      If Empty(mv_par27)
        cQuery += " and SE1.E1_NUMBCO <> ' '"
      EndIf
     
    cQuery += " and SE1.E1_PORTADO = '" + mv_par19 + "'"
    cQuery += " and SE1.E1_AGEDEP  = '" + mv_par20 + "'"
    cQuery += " and SE1.E1_CONTA   = '" + mv_par21 + "'"
  elseIf mv_par23 == 1
        cQuery += " and SE1.E1_NUMBCO = ' '"
  elseIf mv_par23 <> 3 .AND. Empty(mv_par27)
      cQuery += " and SE1.E1_NUMBCO <> ' '"  
      cQuery += " and SE1.E1_PORTADO = '" + mv_par19 + "'"
      cQuery += " and SE1.E1_AGEDEP  = '" + mv_par20 + "'"
      cQuery += " and SE1.E1_CONTA   = '" + mv_par21 + "'"
  ElseIf mv_par23 == 3
      cQuery += " and SE1.E1_PORTADO = '" + mv_par19 + "'"
      cQuery += " and SE1.E1_AGEDEP  = '" + mv_par20 + "'"
      cQuery += " and SE1.E1_CONTA   = '" + mv_par21 + "'"
  EndIf

  cQuery += " and SE1.E1_TIPO not in ('" + MVABATIM + "')"
  cQuery += " and SE1.D_E_L_E_T_ <> '*'"
  cQuery += " Order By SE1.E1_PREFIXO, SE1.E1_NUM, SE1.E1_PARCELA, SE1.E1_TIPO, SE1.E1_CLIENTE, SE1.E1_LOJA "

  If Select("FINR01A") > 0
     dbSelectArea("FINR01A")
     FINR01A->(dbCloseAea())
  EndIf

  TCQuery cQuery New Alias "FINR01A"
  TCSetField("FINR01A", "E1_EMISSAO", "D",08,0)
  TCSetField("FINR01A", "E1_VENCTO" , "D",08,0)
  TCSetField("FINR01A", "E1_VENCREA", "D",08,0)
  TCSetField("FINR01A", "E1_VALOR"  , "N",14,2)
  TCSetField("FINR01A", "E1_REGSE1" , "N",10,0)

  dbSelectArea("FINR01A")
  
  While ! FINR01A->(Eof())
   aAdd(aTitulos, {IIf(nTela == 3,.T.,.F.),;  // 01 = Mark
                   FINR01A->E1_PORTADO,;      // 02 = Portado
                   FINR01A->E1_PREFIXO,;      // 03 = Prefixo do Título
		           FINR01A->E1_NUM,;          // 04 = Número do Título
		           FINR01A->E1_PARCELA,;      // 05 = Parcela do Título
		           FINR01A->E1_TIPO,;         // 06 = Tipo do Título
                   FINR01A->E1_NATUREZ,;      // 07 = Natureza do Título
		           FINR01A->E1_CLIENTE,;      // 08 = Cliente do título
		           FINR01A->E1_LOJA,;         // 09 = Loja do Cliente
		           FINR01A->E1_NOMCLI,;       // 10 = Nome do Cliente
		           Posicione("SA1",1,xFilial("SA1") + FINR01A->E1_CLIENTE + FINR01A->E1_LOJA,"A1_XEMCOB"),;  // 11 = Email de cobrança do cliente
		           FINR01A->E1_EMISSAO,;      // 12 = Data de Emissão do Título
		           FINR01A->E1_VENCTO,;       // 13 = Data de Vencimento do Título
		           FINR01A->E1_VENCREA,;      // 14 = Data de Vencimento Real do Título
		           FINR01A->E1_VALOR,;        // 15 = Valor do Título
		           FINR01A->E1_HIST,;         // 16 = Histótico do Título
		           FINR01A->E1_NUMBCO,;       // 17 = Nosso Número
		           FINR01A->E1_REGSE1,;       // 18 = Número do registro no arquivo
		           FINR01A->E1_AGEDEP,;       // 19 = Agência
		           FINR01A->E1_CONTA,;        // 20 = Conta
		           FINR01A->E1_XSUBCTA})      // 21 = SubConta
    
    FINR01A->(dbSkip())
  EndDo

  If Len(aTitulos) == 0
     aAdd(aTitulos, {.F.,"","","","","","","","","","","","","",0,"",0,"","","",""})
  EndIf

  dbSelectArea("FINR01A")
  FINR01A->(dbCloseArea())
Return

/*=============================================
--  Função: Cria tela de escolha do título   --
--          para impressão.                  --
===============================================*/    
Static Function fnCallTela(aTitulos)
  Local aScreen  := GetScreenRes()
  Local oSize 	:= FwDefSize():New()
  Local bCancel  := {|| RFINR01A(oDlg,@lRetorno,aTitulos) }
  Local bOk      := {|| RFINR01B(oDlg,@lRetorno,aTitulos) }
  Local aAreaAtu := GetArea()
  Local aLabel   := {" ",;
  	                 "Portador",;
  	                 "Prefixo",;
  	                 "Número",;
  	                 "Parcela",;
  	                 "Tipo",;
  	                 "Natureza",;
  	                 "Cliente",;
  	                 "Loja",;
  	                 "Nome Cliente",;
  	                 "EMail",;
  	                 "Emissão",;
  	                 "Vencimento",;
  	                 "Venc.Real",;
  	                 "Valor",;
  	                 "Histórico",;
  	                 "Nosso Número"}
  	                 
  Local aBotao   := {}
  Local lRetorno := .T.
  Local lMark    := .F.
  Local cList1
  Local oDlg
  Local oList1
  Local oMark

  Private oOk	   := LoadBitMap(GetResources(),"LBOK")
  Private oNo    := LoadBitMap(GetResources(),"NADA")
  Private nQtSel := 0
  Private nTtSel := 0
  
 // --- Pegar posição da tela
  oSize:aMargins     := { 0, 0, 0, 0 }        // Espaco ao lado dos objetos 0, entre eles 3
  oSize:aWindSize[3] := (oMainWnd:nClientHeight * 0.99)		
  oSize:lProp        := .F.                   // Proporcional
  oSize:Process()                             // Dispara os calculos

  aAdd(aBotao,{"S4WB011N",{|| U_fnVisReg("SA1",SA1->(aTitulos[oList1:nAt,08] + aTitulos[oList1:nAt,09]),2)},"[F11] - Visualiza Cliente","Cliente"})
  aAdd(aBotao,{"S4WB011N",{|| U_fnVisReg("SE1",SE1->(aTitulos[oList1:nAt,03] + aTitulos[oList1:nAt,04] +; 
                                                     aTitulos[oList1:nAt,05] + aTitulos[oList1:nAt,06] +;
                                                     aTitulos[oList1:nAt,08] + aTitulos[oList1:nAt,09]),2)},"[F12] - Visualiza Título","Título"})
  
  SetKey(VK_F11,{|| IIf(Len(aTitulos) > 0,U_fnVisReg("SA1",SA1->(aTitulos[oList1:nAt,08] + aTitulos[oList1:nAt,09]),2),; 
                                          MsgAlert("Não existe registro selecionado."))})

  SetKey(VK_F12,{|| IIf(Len(aTitulos) > 0,U_fnVisReg("SE1",SE1->(aTitulos[oList1:nAt,03] + aTitulos[oList1:nAt,04] +; 
                                                                 aTitulos[oList1:nAt,05] + aTitulos[oList1:nAt,06] +;
                                                                 aTitulos[oList1:nAt,08] + aTitulos[oList1:nAt,09]),2),;
                                          MsgAlert("Não existe registro selecionado."))})

  Define MsDialog oDlg Title cTitulo From oSize:aWindSize[1],oSize:aWindSize[2] To oSize:aWindSize[3],oSize:aWindSize[4];
           Pixel STYLE nOR( WS_VISIBLE, WS_POPUP ) Of oMainWnd Pixel //"Importação de Tabelas"

    @ 015,005 CHECKBOX oMark VAR lMark PROMPT "Marca Todos" FONT oDlg:oFont PIXEL SIZE 80,09 OF oDlg;
		ON CLICK (aEval(aTitulos, {|x,y| aTitulos[y,1] := lMark}), oList1:Refresh(), fnBOLSel(aTitulos))
    
    @ 030,003 LISTBOX oList1 VAR cList1 Fields HEADER ;
                                               aLabel[01],;
                                               aLabel[02],;
                                               aLabel[03],;
                                               aLabel[04],;
                                               aLabel[05],;
                                               aLabel[06],;
                                               aLabel[07],;
                                               aLabel[08],;
                                               aLabel[09],;
                                               aLabel[10],;
                                               aLabel[11],;
                                               aLabel[12],;
                                               aLabel[13],;
                                               aLabel[14],;
                                               aLabel[15],;
                                               aLabel[16],;
                                               aLabel[17] ;
		Size (oSize:aWindSize[4] - 685),(oSize:aWindSize[3] - 385) NOSCROLL PIXEL
		
    oList1:SetArray(aTitulos)
    oList1:bLine := {|| {If(aTitulos[oList1:nAt,01], oOk, oNo),;
		                 aTitulos[oList1:nAt,02],;
		                 aTitulos[oList1:nAt,03],;
		                 aTitulos[oList1:nAt,04],;
		                 aTitulos[oList1:nAt,05],;
		                 aTitulos[oList1:nAt,06],;
		                 aTitulos[oList1:nAt,07],;
		                 aTitulos[oList1:nAt,08],;
		                 aTitulos[oList1:nAt,09],;
		                 aTitulos[oList1:nAt,10],;
		                 aTitulos[oList1:nAt,11],;
		                 aTitulos[oList1:nAt,12],;
		                 aTitulos[oList1:nAt,13],;
		                 Transform(aTitulos[oList1:nAt,14],"@E 999,999,999.99"),;
		                 Transform(aTitulos[oList1:nAt,15],"@E 999,999,999.99"),;
		                 aTitulos[oList1:nAt,16],;
		                 aTitulos[oList1:nAt,17]}}

    oList1:blDblClick := {|| aTitulos[oList1:nAt,01] := !aTitulos[oList1:nAt,01], oList1:Refresh(), fnBOLSel(aTitulos)}
    oList1:cToolTip   := "Duplo click para marcar/desmarcar o título"
    oList1:Refresh()

    oSayTx1 := TSay():New((oSize:aWindSize[3] - 340),010,{|| "Selecionados:"},oDlg,,,,,,.T.,CLR_BLUE) 
    oSayQtd := TSay():New((oSize:aWindSize[3] - 340),060,{|| Transform(nQtSel,"@E 999,999,999")},oDlg,,,,,,.T.,CLR_BLUE)
 
    oSayTx2 := TSay():New((oSize:aWindSize[3] - 340),100,{|| "Total:"},oDlg,,,,,,.T.,CLR_BLUE) 
    oSayTot := TSay():New((oSize:aWindSize[3] - 340),130,{|| Transform(nTtSel,"@E 999,999,999.99")},oDlg,,,,,,.T.,CLR_BLUE)
    
    oBtImp := TButton():New(013,080,"Impressão",oDlg,{|| RFINR01B(oDlg,@lRetorno,aTitulos,2)},35,13,,,.F.,.T.,.F.,,.F.,,,.F.)
    oBtEma := TButton():New(013,125,"E-mail"   ,oDlg,{|| RFINR01B(oDlg,@lRetorno,aTitulos,6)},35,13,,,.F.,.T.,.F.,,.F.,,,.F.)
    oBtCan := TButton():New(013,170,"Fechar"   ,oDlg,{|| RFINR01A(oDlg,@lRetorno,aTitulos)},35,13,,,.F.,.T.,.F.,,.F.,,,.F.)
  Activate MsDialog oDlg Centered //ON INIT EnchoiceBar(oDlg,bOk,bcancel,,aBotao)
Return(lRetorno)

/*----------------------------------
--  Função: Fechamento da tela    --
--                                --
------------------------------------*/
Static Function RFINR01A(oDlg,lRetorno, aTitulos)
  lRetorno := .F.

  oDlg:End()
Return(lRetorno)

/*----------------------------------------------
--  Função: Conta os registros selecionados.  --
--                                            --
------------------------------------------------*/
Static Function fnBOLSel(aTitulos)
  Local nId := 0
  
  nQtSel := 0
  nTtSel := 0
  
  For nId := 1 to Len(aTitulos)
      If aTitulos[nId][01]
         nTtSel += aTitulos[nId][15]
         nQtSel++
      EndIf   
  Next
  
  ObjectMethod(oSayQtd,"SetText('" + Transform(nQtSel  ,"@E 999,999,999") + "')")         
  ObjectMethod(oSayTot,"SetText('" + Transform(nTtSel  ,"@E 999,999,999.99") + "')")         
Return

/*-----------------------------------------
--  Função: Chamar Impressão de boleto.  --
--                                       --
-------------------------------------------*/
Static Function RFINR01B(oDlg,lRetorno, aTitulos, pTpImp)
  Local nLoop		:= 0
  Local nContador	:= 0

  lRetorno := .T.
  nTpImp   := pTpImp
  
  For nLoop := 1 To Len(aTitulos)
    If aTitulos[nLoop,1]
       nContador++
    EndIf
  Next

  If nContador > 0
     RptStatus( {|lEnd| ImpBol(aTitulos) }, cTitulo)
	else
     lRetorno := .F.
  EndIf

  oDlg:End()
Return(lRetorno)

/*==================================
--  Função: Visualizar título.    --
--                                --
====================================*/
User Function fnVisReg(cAlias, cRecAlias, nOpcEsc)
  Local aAreaAtu    := GetArea()
  Local aAreaAux    := (cAlias)->(GetArea())
  
  Private cCadastro := ""

  If ! Empty(cRecAlias)
     dbSelectArea(cAlias)
     (cAlias)->(dbSetOrder(1))
     (cAlias)->(dbSeek(xFilial(cAlias) + cRecAlias))
	
	 AxVisual(cAlias,(cAlias)->(Recno()),nOpcEsc)
  EndIf

  RestArea(aAreaAux)
  RestArea(aAreaAtu)
Return

/*-------------------------------------
--  Função: Impressão de boleto.     --
--                                   --
---------------------------------------*/
Static Function ImpBol(aTitulos)
  Local aEmpresa := {AllTrim(SM0->M0_NOMECOM),;                                   //[01] Nome da Empresa
                     AllTrim(SM0->M0_ENDENT),;                                    //[02] Endereço
                     AllTrim(SM0->M0_BAIRENT),;                                   //[03] Bairro
                     AllTrim(SM0->M0_CIDENT),;                                    //[04] Cidade
                     SM0->M0_ESTENT,;                                             //[05] Estado
                     "CEP: " + Transform(SM0->M0_CEPENT, "@R 99999-999"),;        //[06] CEP
                     "PABX/FAX: " + SM0->M0_TEL,;                                 //[07] Telefones
                     "CNPJ: " + Transform(SM0->M0_CGC, "@R 99.999.999/9999-99"),; //[08] CGC
                     "I.E.: " + Transform(SM0->M0_INSC, SuperGetMv("MV_IEMASC",.F.,"@R 999.999.999.999"))}	//[09] I.E
	
  Local aCB_RN_NN	:= {}
  Local aDadTit	:= {}
  Local aBanco	:= {}
  Local aSacado	:= {}
  Local aVlBol    := {}

// No máximo 8 elementos com 80 caracteres para cada linha de mensagem
  Local aBolTxt  := {"","","","","","","",""}
  Local nSaldo   := 0
  Local nLoop    := 0
  Local cNumCta  := ""
  Local cChvSA6  := ""
  Local cChvSEE  := ""
  Local cNmPDF   := ""
  Local lGerBor  := .F.
  Local aAreaSA1 := {}
  Local cFilAux := cFilAnt
  Private oPrint
  Private nNumPag   := 1
  Private lBx       := .F.
  Private cDirGer   := AllTrim(mv_par24) + IIf(Substr(AllTrim(mv_par24),Len(AllTrim(mv_par24)),1) == "\","","\")
  Private cNumBor   := IIf(mv_par25 == 1,Soma1(GetMV("MV_NUMBORR"),6),0)
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
  Private cTitErr   := ''
  Private nCols     := 0
  Private nWith     := 0
  Private cDirServ  := "\anexos\boletos\"
  Private cMsgEmail := ""
  Private aChaves   := {}
  Private cTipoArq  := ""

  
 // --- nTpImp = Tipo da impressão: 2 - Spool ou 6 - PDF (envio de e-mail) 
  If ! nTpImp == 6 .and. ! cTpImpre == "2"
     nPosPDF := 0
    //  oPrint:= TMSPrinter():New("Boleto Laser")
     oPrint := FwMSPrinter():New("fatura"+StrTran(Time(),":",""),6,.T.,"c:\temp\",.F.,.F.,,,,.T.,,.T.)
     oPrint:SetResolution(72)
     oPrint:SetMargin(5,5,5,5)
     oPrint:SetPortrait()	
    //  oPrint:SetPortrait()								// ou SetLandscape()
    //  oPrint:StartPage()									// Inicia uma nova página
    //  oPrint:Setup()
   elseIf cTpImpre == "2"                               // Saida em PDF
        If cGerArq == "1"
           nPosPDF := 0
     
          // Nome do PDF: "Bol_" + Codigo Cliente + Loja Cliente + Banco + Prefixo + Titulo
           cNmPDF := "Bol_" + Substr(aSacado[2],1,TamSX3("A1_COD")[1]) + "_" +;
                       Substr(aSacado[2],(TamSX3("A1_COD")[1] + 2),TamSX3("A1_LOJA")[1]) +;
                       "_" + mv_par19 + "_" + AllTrim(aTitulos[01][03]) +;
                       "_" + AllTrim(aTitulos[01][04]) + "_TODAS" 
          // --------------------------------------------------           
      
           oPrint := FwMSPrinter():New(cNmPDF,6,.T.,cDirGer,.T.,.F.,,,,.T.,,.F.)
           oPrint:SetResolution(72)
           oPrint:SetMargin(5,5,5,5)
           oPrint:SetPortrait()								// ou SetLandscape()
        EndIf  
  EndIf
  If nTpImp == 6
    
  EndIf
// Fazer pergunta de Centimetro ou Polegada

  nTipo := 1 /* Aviso(	"Impressão",;
		"Escolha o método de impressão.",;
		{"&Centimetro","&Polegada"},,;
		"A T E N Ç Ã O" ) */

  SetRegua(Len(aTitulos))

  cNumBor := Replicate("0",6-Len(Alltrim(cNumBor))) + Alltrim(cNumBor)

  While ! MayIUseCode("SE1"+xFilial("SE1")+cNumBor)      // verifica se esta na memoria, sendo usado
	// busca o proximo numero disponivel 
	cNumBor := Soma1(cNumBor)
  EndDo
  
  If len(aTitulos) > 0 .AND. nTpImp == 6
    dbSelectArea("SE1")
    SE1->(dbGoTo(aTitulos[1,18]))

    dbSelectArea("SA1")
    SA1->(dbSetOrder(1))
    SA1->(dbSeek(xFilial("SA1") + SE1->E1_CLIENTE + SE1->E1_LOJA))
    
  
  
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
    
    If !(IsInCallStack("U_LAUA0004") .AND. !Empty(ZA7->ZA7_GRUPO)) .OR.  IsInCallStack("U_LAUA0007")
      GetMsg()
    EndIf

  EndIf
  
  If nTpImp == 6
    cTipoArq := U_LAUA007F('1')
    If Empty(cTipoArq)
      Return NIL
    EndIf
  EndIf
 
 // Faz loop no array com os títulos a serem impressos
  For nLoop := 1 To Len(aTitulos)
     

      IncRegua("Titulo: "+aTitulos[nLoop,02]+"/"+aTitulos[nLoop,03]+"/"+aTitulos[nLoop,04])
	 // Se estiver marcado, imprime
      If aTitulos[nLoop,01]
         dbSelectArea("SE1")
         SE1->(dbGoTo(aTitulos[nLoop,18]))
         setEmpresa(SE1->E1_FILIAL)
         dbSelectArea("SA1")
         SA1->(dbSetOrder(1))
         SA1->(dbSeek(xFilial("SA1") + SE1->E1_CLIENTE + SE1->E1_LOJA))
         aAreaSA1 := SA1->(GetArea())
         
         dbSelectArea("SA6")
         SA6->(dbSetOrder(1))
         
         If ! Empty(aTitulos[nLoop,02]) .and. mv_par23 == 2
            cChvSA6 := aTitulos[nLoop,02] + aTitulos[nLoop,19] + aTitulos[nLoop][20]
          else  
            cChvSA6 := PADR(mv_par19,LEN(SA6->A6_COD)) + PADR(mv_par20,LEN(SA6->A6_AGENCIA)) + PADR(mv_par21,LEN(SA6->A6_NUMCON))
         EndIf
         
         If ! SA6->(dbSeek(xFilial("SA6") + cChvSA6))
            Aviso("Emissão do Boleto","Banco não localizado no cadastro!",{"&Ok"},,;
                  "Banco: " + SubStr(cChvSA6,1,3) + "/" + SubStr(cChvSA6,4,5) + "/" + SubStr(cChvSA6,9,10))
            Loop
         EndIf
        
        If IsInCallStack("U_LAUA0004") .AND. !Empty(ZA7->ZA7_GRUPO) .AND. !IsInCallStack("U_LAUA0007")
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
        EndIf
		     //Posiciona na Configuração do Banco
         dbSelectArea("SEE")
         SEE->(dbGoTop())
         SEE->(dbSetOrder(1))

         If ! Empty(aTitulos[nLoop,02]) .AND. Empty(mv_par22)
            cChvSEE := PADR(aTitulos[nLoop,02],len(SEE->EE_CODIGO)) + PADR(aTitulos[nLoop,19],len(SEE->EE_AGENCIA)) + PADR(aTitulos[nLoop][20],len(SEE->EE_CONTA)) + PADR(aTitulos[nLoop][21],len(SEE->EE_SUBCTA)) 
         else  
            cChvSEE := PADR(mv_par19,len(SEE->EE_CODIGO)) + PADR(mv_par20,len(SEE->EE_AGENCIA)) + PADR(mv_par21,len(SEE->EE_CONTA)) + PADR(mv_par22,len(SEE->EE_SUBCTA))
         EndIf
         
         If ! SEE->(dbSeek(xFilial("SEE") + cChvSEE))
            Aviso("Emissão do Boleto",	"Configuração dos parâmetros do banco não localizado no cadastro!",;
					{"&Ok"},,"Banco: " + Substr(cChvSEE,1,3) + "/" + SubStr(cChvSEE,4,5) + "/" +;
					SubStr(cChvSEE,9,10) + "/" + SubStr(cChvSEE,19,3))
            Loop
          else
            cLogo   := AllTrim(SEE->EE_XLOGO)
            aLinDig := {}
            
            aAdd(aLinDig, AllTrim(SEE->EE_XNNUM))     // Formatação do Nosso Numero
            aAdd(aLinDig, AllTrim(SEE->EE_XDGNN))     // Formatação para calculo no digito do nosso numero
            aAdd(aLinDig, AllTrim(SEE->EE_XMTNN))     // Montagem do Nosso Numero para o boleto
            aAdd(aLinDig, AllTrim(SEE->EE_XCRN1))     // Formatação da primeiro parte
            aAdd(aLinDig, AllTrim(SEE->EE_XCRN2))     // Formatação da segunda parte
            aAdd(aLinDig, AllTrim(SEE->EE_XCRN3))     // Formatação da terceira parte
            aAdd(aLinDig, AllTrim(SEE->EE_XCRN4))     // Formatação da quarta parte
            aAdd(aLinDig, AllTrim(SEE->EE_XCPLV))     // Formatação para Campo livre com digito
         EndIf
			
         dbSelectArea("SE1")
         cNumCta := IIf(AllTrim(SA6->A6_COD) == "237",StrZero(Val(AllTrim(SA6->A6_NUMCON)),7),AllTrim(SA6->A6_NUMCON))
        
         aBanco := {AllTrim(SA6->A6_COD),;                                                             // [01] Numero do Banco
                    SA6->A6_NREDUZ,;                                                                   // [02] Nome do Banco
                    IIf(Len(AllTrim(SA6->A6_AGENCIA)) < 4,StrZero(Val(AllTrim(SA6->A6_AGENCIA)),4),;
                                                          AllTrim(SA6->A6_AGENCIA)),;                  // [03] Agência
                    cNumCta,;                                                                          // [04] Conta Corrente
                    SubStr(SA6->A6_DVCTA,At("-",SA6->A6_DVCTA)+1,1),;                                  // [05] Dígito da conta corrente
                    AllTrim(SEE->EE_CODCART),;                                                         // [06] Codigo da Carteira
                    SEE->EE_XDVBCO,;                                                                   // [07] Dígito do Banco
                    SA6->A6_DVAGE,;                                                                    // [08] Digito da Agência
                    IIf(AllTrim(SA6->A6_COD) $ ("341/104"),AllTrim(SEE->EE_CODEMP),StrZero(Val(SEE->EE_CODEMP),7)),;// [09] Convêncio com o Banco
                    IIf(SEE->EE_TPCOBRA == "1",;
                      AllTrim(SEE->EE_CODCART),"SR")}// [10] Tipo da Carteira

			If Empty(SA1->A1_ENDCOB)
				aSacado := {AllTrim(SA1->A1_NOME),;						                 // [1] Razão Social
				            AllTrim(SA1->A1_COD ) + "-" + SA1->A1_LOJA,;                 // [2] Código
				            AllTrim(SA1->A1_END ) + " - " + AllTrim(SA1->A1_BAIRRO),;    // [3] Endereço
				            AllTrim(SA1->A1_MUN ),;                                      // [4] Cidade
				            SA1->A1_EST,;                                                // [5] Estado
				            SA1->A1_CEP,;                                                // [6] CEP
				            SA1->A1_CGC,;                                                // [7] CGC
				            SA1->A1_PESSOA}                                              // [8] PESSOA

			 else
				aSacado := {AllTrim(SA1->A1_NOME),;                                        // [1] Razão Social
				            AllTrim(SA1->A1_COD ) + "-" + SA1->A1_LOJA,;                   // [2] Código
                          AllTrim(SA1->A1_ENDCOB) + " - " + AllTrim(SA1->A1_BAIRROC),;   // [3] Endereço
                          AllTrim(SA1->A1_MUNC),;                                        // [4] Cidade
                          SA1->A1_ESTC,;                                                 // [5] Estado
                          SA1->A1_CEPC,;                                                 // [6] CEP
                          SA1->A1_CGC,;                                                  // [7] CGC
                          SA1->A1_PESSOA}                                                // [8] PESSOA
			Endif

		// Define o valor do título considerando Acréscimos e Decréscimos
		   aVlBol  := U_fnSldBol(SE1->E1_PREFIXO,SE1->E1_NUM,SE1->E1_PARCELA,SE1->E1_CLIENTE,SE1->E1_LOJA)
           nSaldo  := aVlBol[01]                // Valor do documento
           nJurMul := aVlBol[03]                // Valor de Mora/Multa (Acréscimos)
           nDesc   := aVlBol[04]                // Valor do desconto (Decréscimos)

		// Define o Nosso Número
			If ! Empty(SE1->E1_NUMBCO) .and. mv_par23 == 2
          If aBanco[1] != '001'
            cNNum	:= Substr(AllTrim(SE1->E1_NUMBCO),1,(len(Alltrim(SE1->E1_NUMBCO)) - 1))
          Else
            cNNum	:= Right(SE1->E1_NUMBCO, LEN(SEE->EE_FAXATU))
          EndIf
				bRetImp := .T. 
			 else              
              bRetImp := .F.
              cNNum   := AllTrim(SEE->EE_FAXATU)
              
			
				dbSelectArea("SEE")
				RecLock("SEE",.F.)
				  Replace SEE->EE_FAXATU with Soma1(Alltrim(SEE->EE_FAXATU),11)
				SEE->(MsUnLock())
			EndIf         
			
			dbSelectArea("SE1")

		  // ---- Monta codigo de barras
          aCB_RN_NN := Ret_cBarra(Subs(aBanco[1],1,3),;			// [01]-Banco+Fixo 9
                                  aBanco[3],;						// [02]-Agencia
                                  aBanco[4],;	    				// [03]-Conta
                                  aBanco[5],;						// [04]-Digito Conta
                                  aBanco[6],;						// [05]-Carteira
                                  cNNum,;							// [06]-Nosso Número
                                  nSaldo,;						// [07]-Valor do Título
                                  SE1->E1_VENCTO,;             // [08]-Vencimento
                                  aBanco[9],;                  // [09]-Convêncio
                                  SEE->EE_XMODELO,;            // [10]-Modulo para calculo do digito verificador do Nosso Número
                                  SEE->EE_XPESO)               // [11]-Peso para calcular o digito do nosso número modulo 11
 
          dbSelectArea("SE1")

          aDadTit := {AllTrim(E1_NUM) + IIf(Empty(E1_PARCELA),"","/" + E1_PARCELA),;  // [01] Número do título
                      E1_EMISSAO,;                                                    // [02] Data da emissão do título
                      dDataBase,;                                                     // [03] Data da emissão do boleto
                      E1_VENCTO,;                                                     // [04] Data do vencimento
                      nSaldo,;                                                        // [05] Valor do título
                      aCB_RN_NN[3],;                                                  // [06] Nosso número (Ver fórmula para calculo)
                      E1_PREFIXO,;                                                    // [07] Prefixo da NF
                      SEE->EE_X_ESPDC,;                                               // [08] Tipo do Titulo
                      E1_HIST,;                                                       // [09] HISTORICO DO TITULO
                      aCB_RN_NN[4],;                                                  // [10] Nosso numero para gravação na "SE1"
                      SEE->EE_XLOCPG}                                                 // [11] Mensagem de Local de Pagamento                                  

			aBolTxt := {"","","","","","","",""}
		
          If SEE->EE_X_MULTA > 0
             aBolTxt[1] := "Multa de R$ " + Alltrim(Transform(((aDadTit[5] * SEE->EE_X_MULTA) / 100),"@E 999.99")) + " após o vencimento."
          EndIf
			
          If SEE->EE_X_JURME > 0
             aBolTxt[2] := "Juros de R$ " + AllTrim(Transform(((aDadTit[5] * SEE->EE_X_JURME) / 100),"@E 999.99")) + " ao dia."
          EndIf

          If Val(SEE->EE_DIASPRT) > 0
             aBolTxt[3] := "Título sujeito a protesto após " + SEE->EE_DIASPRT + " dias de vencimento."
          EndIf
			
          If ! Empty(Alltrim(SEE->EE_FORMEN1))
             aBolTxt[5] := Substr(AllTrim(&(SEE->EE_FORMEN1)),1,(Len(AllTrim(&(SEE->EE_FORMEN1))) )) 
          EndIf
			
          If ! Empty(Alltrim(SEE->EE_FORMEN2))
             aBolTxt[6] := Substr(AllTrim(SEE->EE_FORMEN2),2,(Len(AllTrim(SEE->EE_FORMEN2)) - 2 ))
          EndIf

          If ! Empty(Alltrim(SEE->EE_FOREXT1))
             aBolTxt[7] := SubStr(AllTrim(SEE->EE_FOREXT1),1,(Len(AllTrim(SEE->EE_FOREXT1)) )) 
          EndIf
			
          If ! Empty(Alltrim(SEE->EE_FOREXT2))
		      aBolTxt[8] := SubStr(AllTrim(SEE->EE_FOREXT2),1,(Len(AllTrim(SEE->EE_FOREXT2))))
          EndIf

		  // Sempre Incremento a mensagem de não receber após vencimento
         //  aBolTxt[8]	:= "SR. CAIXA, NÃO RECEBER APÓS O VENCIMENTO"

         // Valida se é impressão em PDF para envio de E-mail ou impressão somente PDF
          If nTpImp == 6 .or. cTpImpre == "2" 
            cGerArq := "2"
             If cGerArq == "2"
                nPosPDF := 0
                cDirGer := "C:\temp\"
                If !ExistDir("C:\temp")
                  MakeDir("c:\temp\")
                EndIf
               // Nome do PDF: "Bol_" + Codigo Cliente + Loja Cliente + Banco + Prefixo + Titulo + Parcela
                cNmPDF := "Bol" + Substr(aSacado[2],1,TamSX3("A1_COD")[1]) + "" +;
                          Substr(aSacado[2],(TamSX3("A1_COD")[1] + 2),TamSX3("A1_LOJA")[1]) +;
                          "" + Alltrim(mv_par19) + "" + AllTrim(aTitulos[nLoop,3]) +;
                          "" + AllTrim(aTitulos[nLoop,4]) + "" + AllTrim(aTitulos[nLoop,5]) 
                // --------------------------------------------------           
                // nFlags := PD_ISTOTVSPRINTER + PD_DISABLEPAPERSIZE + PD_DISABLEPREVIEW + PD_DISABLEMARGIN
                // oSetup := FWPrintSetup():New(nFlags, "FILIAL " + cFilAnt)

                oPrint := FwMSPrinter():New(cNmPDF       ,6      ,.T.,cDirGer    ,.T.,.F.,,,,.T.,,.F.)
                oPrint:SetResolution(72)
                oPrint:SetMargin(5,5,5,5)
          
                oPrint:SetPortrait()								// ou SetLandscape()
             EndIf
             	
             oPrint:StartPage()
 
             If mv_par26 == 1                              
                fnImprRd(oPrint,aEmpresa,aDadTit,aBanco,aSacado,aBolTxt,aCB_RN_NN,cNNum)     // Impressão boleto reduzido
              else  
                u_fnImpres(oPrint,aEmpresa,aDadTit,aBanco,aSacado,aBolTxt,aCB_RN_NN,cNNum)     // Impressão boleto completo
             EndIf   

             If nTpImp == 6      // Se não for direito PDF
                CriaPaths()
                cFilePrint := cNmPDF + ".PD_"
                cFile := cDirGer + cNmPDF+ ".pdf"     
                If File( cFile )
                  FErase(cFile)
                EndIf        
                File2Printer(cFilePrint,"PDF")
                
                oPrint:cPathPDF := cDirGer
                oPrint:EndPage()     // Finaliza a página
                oPrint:Preview()     // Visualiza antes de imprimir
                CpyT2S( cFile, cDirServ)
                FErase(cFile)
                If "BOLETO" $ cTipoArq
                  cFile := cDirServ + cNmPDF+ ".pdf" 
                Else 
                  cFile := ""
                EndIf

                // cDirDacte := "\anexos\dactes\"
                // cDirXML   := "\anexos\xmls\"
               If "DACTES" $ cTipoArq .OR. "XML" $ cTipoArq
                  cFileXDac := ExporXML(cDirGer, 'XML' $ cTipoArq, 'DACTES' $ cTipoArq, .T.)
                  
                  If !Empty(cFileXDac)
                    cFile += IIF(!Empty(cFile) ,";"+ cFileXDac, cFileXDac)
                  Else
                    MsgAlert('Erro ao tentar compactar xmls e dactes para envio. ')
                  EndIf
                EndIf

                SA1->(RestArea(aAreaSA1))
                
                If !EnvPCli(cFile)
                // 'CLIENTE;NOME;FILIAL;NUMERO;PREFIXO;TIPO;PARCELA' + CRLF
                  cTitErr += Alltrim(SA1->A1_COD)+";"+Alltrim(SA1->A1_NOME)+";"+SE1->E1_FILIAL + ';' + SE1->E1_NUM + ';' + SE1->E1_PREFIXO + ";" + SE1->E1_TIPO + CRLF
                EndIf

                FErase(cFile)                
                FreeObj(oPrint)      // Destruir objeto

             EndIf   
           else
             oPrint:StartPage()
 
             If mv_par26 == 1
                fnImprRd(oPrint,aEmpresa,aDadTit,aBanco,aSacado,aBolTxt,aCB_RN_NN,cNNum)     // Impressão boleto reduzido
              else  
                u_fnImpres(oPrint,aEmpresa,aDadTit,aBanco,aSacado,aBolTxt,aCB_RN_NN,cNNum)     // Impressão boleto completo
             EndIf   
          EndIf
		  // -----------------------------------------------
		  	
          dbSelectArea("SE1")
          SE1->(dbGoTo(aTitulos[nLoop,18]))			
          If Empty(SE1->E1_NUMBCO)
            Reclock("SE1",.F.)
              Replace SE1->E1_PORTADOR with mv_par19
              Replace SE1->E1_AGEDEP   with mv_par20
              Replace SE1->E1_CONTA    with mv_par21
              Replace SE1->E1_XSUBCTA  with mv_par22
            SE1->(MsUnlock())
          EndIf
         // ---- Geração de Bordero 
          If mv_par23 <> 2 .and. mv_par25 == 1
             cNumBor := Replicate("0",6-Len(Alltrim(cNumBor))) + Alltrim(cNumBor)

             While ! MayIUseCode("SE1" + xFilial("SE1") + cNumBor)      // verifica se esta na memoria, sendo usado
	           // busca o proximo numero disponivel 
	            cNumBor := Soma1(cNumBor)
             EndDo

             fnGrvBrd()
  
             PutMv("MV_NUMBORR", cNumBor)
             
             lGerBor := .T.
          EndIf
         // ------------------------  
		EndIf
  Next 

  If lGerBor             
     Aviso("ATENÇÃO","Bordero - " + cNumBor + " gerado com sucesso...",{"OK"})
  EndIf
  
  If ! nTpImp == 6 .and. ! cTpImpre == "2"
     oPrint:EndPage()     // Finaliza a página
     oPrint:Preview()     // Visualiza antes de imprimir
  ElseIf nTpImp != 6
    // ApMsgInfo("Envio concluído com sucesso.")
  EndIf
  setEmpresa(cFilAux)

  If !Empty(cTitErr)
    
    If !ExistDir("c:temp")
      MakeDir("c:\temp\")
    EndIf

    cTitErr := 'CLIENTE;NOME;FILIAL;NUMERO;PREFIXO;TIPO' + CRLF + cTitErr
    
    MemoWrite('c:\temp\fatura_nao_enviada.csv', cTitErr)
    
    ApMsgInfo('Alguma(s) fatura(s) não foram enviadas, verifique o cadastro dos clientes. Lista de clientes está disponível em: c:\temp\fatura_nao_enviada.csv')
  EndIf

Return(Nil)

/*------------------------------------
--  Função: Impressão dos dados.    --
--                                  --
--------------------------------------*/
User Function fnImpres(oPrint,aEmpresa,aDadTit,aBanco,aSacado,aBolTxt,aCB_RN_NN,cNNum, lSoFatura)
  Local nI       := 0
  Local nZ       := 0
  Local i        := 1
  Local cBmp     := ""
  Local oFont07  := TFont():New("Arial Narrow",9, -08,.T.,.F.,5,.T.,5,.T.,.F.)
  Local oFont07n := TFont():New("Arial Narrow",9, -08,.T.,.T.,5,.T.,5,.T.,.F.)
  Local oFont08  := TFont():New("Arial Narrow"       ,9, -10,.T.,.F.,5,.T.,5,.T.,.F.)
  Local oFont08n := TFont():New("Arial Narrow"       ,9, -10,.T.,.T.,5,.T.,5,.T.,.F.)
  Local oFont11c := TFont():New("Courier New" ,9, -11,.T.,.T.,5,.T.,5,.T.,.F.)
  Local oFont10  := TFont():New("Arial Narrow"       ,9, -10,.T.,.T.,5,.T.,5,.T.,.F.)
  Local oFont12 := TFont():New("Arial Narrow",12,-12 ,.T.,.F.,5,.T.,5,.T.,.F.)
  Local oFont12n := TFont():New("Arial Narrow",12,-12 ,.T.,.T.,5,.T.,5,.T.,.F.)
  Local oFont14n := TFont():New("Courier New",14,-14 ,.T.,.T.,5,.T.,5,.T.,.F.)
  Local oFont15  := TFont():New("Arial Narrow"       ,9, -15,.T.,.F.,5,.T.,5,.T.,.F.)    
  Local oFont15n := TFont():New("Arial Narrow"       ,9, -15,.T.,.T.,5,.T.,5,.T.,.F.)    
  Local oFont20  := TFont():New("Arial Narrow"       ,9, -18,.T.,.T.,5,.T.,5,.T.,.F.)
  Local aCamposFil := { "M0_NOME", "M0_NOMECOM", "M0_CGC", "M0_TEL", "M0_FAX", "M0_ENDCOB", "M0_CIDCOB", "M0_ESTCOB", "M0_CEPCOB", "M0_COMPCOB","M0_BAIRCOB" }
  Local aAreaSE1 := SE1->(GetArea())
  Local aAreaSA1 := SA1->(GetArea())
  Local nSaldo    := 0
  Local nJurMul   := 0
  Local nDesc   := 0
  Private nPag      := 1
  Default lSoFatura := .F.
  SA1->(DbSetOrder(1))
  SA1->(DbSeek(xFilial("SA1") +SE1->(E1_CLIENTE+E1_LOJA)))
  aAreaSA1 := SA1->(GetArea())
  //Variaveis usadas em array
  static NOME := 1
  static NOMECOM := 2
  static CGC := 3
  static TEL := 4
  static FAX := 5
  static ENDCOB := 6
  static CIDCOB := 7
  static ESTCOB := 8
  static CEPCOB := 9
  static COMPCOB := 10
  static BAIRCOB := 11
  aEmpresa := {AllTrim(SM0->M0_NOMECOM),;                                   //[01] Nome da Empresa
                     AllTrim(SM0->M0_ENDENT),;                                    //[02] Endereço
                     AllTrim(SM0->M0_BAIRENT),;                                   //[03] Bairro
                     AllTrim(SM0->M0_CIDENT),;                                    //[04] Cidade
                     SM0->M0_ESTENT,;                                             //[05] Estado
                     "CEP: " + Transform(SM0->M0_CEPENT, "@R 99999-999"),;        //[06] CEP
                     "PABX/FAX: " + SM0->M0_TEL,;                                 //[07] Telefones
                     "CNPJ: " + Transform(SM0->M0_CGC, "@R 99.999.999/9999-99"),; //[08] CGC
                     "I.E.: " + Transform(SM0->M0_INSC, SuperGetMv("MV_IEMASC",.F.,"@R 999.999.999.999"))}	//[09] I.E

  cBmp	:= cStartPath + cLogo + ".bmp"
  
  nSizeHorz := oprint:nHorzRes()
  nSizeVert := 3000
  nXFim := 2275
  
  nRow1 := 50
  nRow2 := nSizeVert/5 
  nRow3 := nRow2 + nRow2/3
  nRow4 := 3*nRow2 
  nRow5 := 5*nRow2 - 15
  If !Empty(SE1->E1_NUMLIQ)
    aVlBolAux  := U_fnSldBol(SE1->E1_PREFIXO,SE1->E1_NUM,SE1->E1_PARCELA,SE1->E1_CLIENTE,SE1->E1_LOJA)
    nSaldo  := aVlBolAux[01]                // Valor do documento
    nJurMul := aVlBolAux[03]                // Valor de Mora/Multa (Acréscimos)
    nDesc   := aVlBolAux[04]                // Valor do desconto (Decréscimos)

    cBitmap := R110ALogo()
    _cFileLogo	:= GetSrvProfString('Startpath','') + cBitmap
    oPrint:SayBitmap(-100 , 40,_cFileLogo,800,600)
    // //Box linha 1 coluna esquerda
    // oPrint:line(  nRow1, 10, nRow1, nXFim/2 , 	RGB( 147, 166, 255 ), "20" )
    oPrint:Line( nRow1+15, 10, nRow1+15, nXFim/2, CLR_BLACK ,"10") // Linha horizontal superior
    oPrint:Line( nRow1, 10, nRow2, 10, CLR_BLACK ,"-9") // Linha vertical esquerda
    oPrint:Line( nRow2, 10, nRow2, nXFim/2, CLR_BLACK ,"-9") // Linha horizontal inferior
    oPrint:Line( nRow1, nXFim/2, nRow2, nXFim/2, CLR_BLACK ,"-9") // Linha vertical Direita
    
    
    _aFilial := FWSM0Util():GetSM0Data( cEmpAnt , cFilAnt , aCamposFil )
    oPrint:Say(nRow1 + 320 , 50, Alltrim(_aFilial[NOMECOM][2]), oFont15n)
    Alltrim(Transform(_aFilial[CGC][2],"@R 99.999.999/9999-99"))

    oPrint:Say(nRow1 + 395 , 50, "CNPJ: "+Alltrim(Transform(_aFilial[CGC][2],"@R 99.999.999/9999-99")), oFont15)
    oPrint:Say(nRow1 + 430 , 50, Alltrim(_aFilial[ENDCOB][2]) + " - "+Alltrim(_aFilial[BAIRCOB][2]), oFont15)
    oPrint:Say(nRow1 + 465 , 50, Alltrim(_aFilial[CIDCOB][2]) + " - "+Alltrim(_aFilial[ESTCOB][2])+" - CEP: "+Alltrim(Transform(_aFilial[CEPCOB][2],"@R 99999-999")), oFont15)
    oPrint:Say(nRow1 + 500 , 50, "Fone: "+Alltrim(_aFilial[TEL][2]), oFont15) 

    //Box linha 1 coluna 1 | direita
    oPrint:Line( nRow1+15, (nXFim/2) + 20, nRow1+15, nXFim, CLR_BLACK ,"10") // Linha horizontal superior
    oPrint:Line( nRow1, (nXFim/2) + 20, nRow2/2, (nXFim/2) + 20, CLR_BLACK ,"-9") // Linha vertical esquerda
    oPrint:Line( nRow2/2, (nXFim/2) + 20, nRow2/2, nXFim, CLR_BLACK ,"-9") // Linha horizontal inferior
    oPrint:Line( nRow1, nXFim, nRow2/2, nXFim, CLR_BLACK ,"-9") // Linha vertical Direita

    oPrint:Say(nRow1 + 70 , (nXFim/2) + 40, AllTrim(SA1->A1_NOME), oFont15n)
    oPrint:Say(nRow1 + 105 , (nXFim/2) + 40, "CNPJ: "+Alltrim(Transform(SA1->A1_CGC,"@R 99.999.999/9999-99")) , oFont15)
    oPrint:Say(nRow1 + 140 , (nXFim/2) + 40, AllTrim(SA1->A1_END) +" - "+AllTrim(SA1->A1_BAIRRO), oFont15)
    oPrint:Say(nRow1 + 175 , (nXFim/2) + 40, "CEP: "+Alltrim(Transform(SA1->A1_CEP,"@R 99999-999")) + " - "+AllTrim(SA1->A1_MUN)+" - "+AllTrim(SA1->A1_EST), oFont15)
    oPrint:Say(nRow1 + 210 , (nXFim/2) + 40, AllTrim(SA1->A1_TEL), oFont15) 

    nIni2 := ((nXFim/2) + 20 ) + ((nXFim/2) + 20 ) / 3
    nIni3 := nXFim - ( ((nXFim/2) + 20 ) / 3)

    //Box linha 1 coluna 2 | direita
    oPrint:Line( (nRow2/2)+35, (nXFim/2) + 20, (nRow2/2)+35, nXFim, CLR_BLACK ,"10") // Linha horizontal superior
    oPrint:Line( (nRow2/2)+15, (nXFim/2) + 20, nRow2/2 + nRow2/4 , (nXFim/2) + 20, CLR_BLACK ,"-9") // Linha vertical esquerda
    oPrint:Line( nRow2/2 + nRow2/4 , (nXFim/2) + 20, nRow2/2 + nRow2/4 , nXFim, CLR_BLACK ,"-9") // Linha horizontal inferior
    oPrint:Line( (nRow2/2)+15, nXFim, nRow2/2 + nRow2/4 , nXFim, CLR_BLACK ,"-9") // Linha vertical Direita

    oPrint:Line( (nRow2/2)+15, nIni2, nRow2/2 + nRow2/4 , nIni2, CLR_BLACK ,"-9") // Linha vertical esquerda
    oPrint:Line( (nRow2/2)+15, nIni3, nRow2/2 + nRow2/4 , nIni3, CLR_BLACK ,"-9") // Linha vertical esquerda

    oPrint:Say((nRow2/2) + 80 , (nXFim/2) + 50, "NÚMERO", oFont12n)
    oPrint:Say((nRow2/2) + 125 , (nXFim/2) + 100, SE1->E1_NUM , oFont14n)
    oPrint:Say((nRow2/2) + 80 , nIni2 + 30, "EMISSÃO", oFont12n)
    oPrint:Say((nRow2/2) + 125 , nIni2 + 80, DtoC(SE1->E1_EMISSAO) , oFont14n)
    oPrint:Say((nRow2/2) + 80 , nIni3 + 30, "VENCIMENTO", oFont12n)
    oPrint:Say((nRow2/2) + 125 , nIni3 + 80, DtoC(SE1->E1_VENCTO) , oFont14n)

    //Box linha 1 coluna 3 | direita
    oPrint:Line( (nRow2/2 + nRow2/4)+35, (nXFim/2) + 20, (nRow2/2 + nRow2/4)+35, nIni2 - 10, CLR_BLACK ,"10") // Linha horizontal superior
    oPrint:Line( (nRow2/2 + nRow2/4)+15, (nXFim/2) + 20, nRow2 , (nXFim/2) + 20, CLR_BLACK ,"-9") // Linha vertical esquerda
    oPrint:Line( nRow2 , (nXFim/2) + 20, nRow2 , nIni2 - 10, CLR_BLACK ,"-9") // Linha horizontal inferior
    oPrint:Line( (nRow2/2 + nRow2/4)+15, nIni2 - 10, nRow2 , nIni2 - 10, CLR_BLACK ,"-9") // Linha vertical Direita

    oPrint:Say((nRow2/2 + nRow2/4) + 125 , (nXFim/2) + 30, Alltrim(Transform(nDesc,"@E 99,999,999.99"))  , oFont14n)
    oPrint:Say((nRow2/2 + nRow2/4) + 125 , nIni2 + 30, Alltrim(Transform(nJurMul,"@E 99,999,999.99")) , oFont14n)
    oPrint:Say((nRow2/2 + nRow2/4) + 125 , nIni3 + 30, Alltrim(Transform(nSaldo,"@E 99,999,999.99")) , oFont14n)

    //Box linha 1 coluna 4 | direita
    oPrint:Line( (nRow2/2 + nRow2/4)+35, nIni2, (nRow2/2 + nRow2/4)+35, nIni3 - 10, CLR_BLACK ,"10") // Linha horizontal superior
    oPrint:Line( (nRow2/2 + nRow2/4)+15, nIni2, nRow2 , nIni2, CLR_BLACK ,"-9") // Linha vertical esquerda
    oPrint:Line( nRow2 , nIni2, nRow2 , nIni3 - 10, CLR_BLACK ,"-9") // Linha horizontal inferior
    oPrint:Line( (nRow2/2 + nRow2/4)+15, nIni3 - 10, nRow2 , nIni3 - 10, CLR_BLACK ,"-9") // Linha vertical Direita

    //Box linha 1 coluna 5 | direita
    oPrint:Line( (nRow2/2 + nRow2/4)+35, nIni3, (nRow2/2 + nRow2/4)+35, nXFim, CLR_BLACK ,"10") // Linha horizontal superior
    oPrint:Line( (nRow2/2 + nRow2/4)+15, nIni3, nRow2 , nIni3, CLR_BLACK ,"-9") // Linha vertical esquerda
    oPrint:Line( nRow2 , nIni3, nRow2 , nXFim, CLR_BLACK ,"-9") // Linha horizontal inferior
    oPrint:Line( (nRow2/2 + nRow2/4)+15, nXFim, nRow2 , nXFim, CLR_BLACK ,"-9") // Linha vertical Direita

    //Box linha 2
    oPrint:Line( nRow2 + 5 + 25 + 15,  10, nRow2 + 5 + 25 + 15, nXFim, CLR_BLACK ,"10") // Linha horizontal superior
    oPrint:Line( nRow2 + 5 + 25,  10, nRow3+ 20, 10, CLR_BLACK ,"-9") // Linha vertical esquerda
    oPrint:Line( nRow3 + 20,  10, nRow3 + 20, nXFim, CLR_BLACK ,"-9") // Linha horizontal inferior
    oPrint:Line( nRow2 + 5 + 25, nXFim, nRow3+ 20, nXFim, CLR_BLACK ,"-9") // Linha vertical Direita
    
    //Box linha 3
    // oPrint:Line( nRow3 + 25 + 25 + 15,  10, nRow3 + 25 + 25 +15, nXFim, CLR_BLACK ,"10") // Linha horizontal superior
    // oPrint:Line( nRow3 + 25 + 25,  10, nRow4+ 20, 10, CLR_BLACK ,"-9") // Linha vertical esquerda
    // oPrint:Line( nRow4 + 20,  10, nRow4 + 20, nXFim, CLR_BLACK ,"-9") // Linha horizontal inferior
    // oPrint:Line( nRow3 + 25 + 25, nXFim, nRow4+ 20, nXFim, CLR_BLACK ,"-9") // Linha vertical Direita
    
    aChaves := GetTitFat()
    nLinha := nRow3 + 120

    aAreaSE1 := SE1->(GetArea())
    aAreaSF2 := SF2->(GetArea())
    aAreaSA1 := SA1->(GetArea())
    aAreaSC5 := SC5->(GetArea())
    aAreaSC6 := SC6->(GetArea())
    
    SE1->(dbSetOrder(1))
    oPrint:Say(nLinha, 25, "TIPO"   , oFont12n ) 
    oPrint:Say(nLinha, 150, "NÚMERO" , oFont12n ) 
    oPrint:Say(nLinha, 350, "EMISSÃO", oFont12n ) 
    oPrint:Say(nLinha, 550, "FILIAL" , oFont12n ) 
    oPrint:Say(nLinha, 700, "FRETE"  , oFont12n ) 
    oPrint:Say(nLinha, 925, "NF"     , oFont12n ) 
    oPrint:Say(nLinha, 1250, "REMETENTE", oFont12n ) 
    oPrint:Say(nLinha, 1600, "DESTINATARIO", oFont12n ) 
    oPrint:Say(nLinha, 1920, "ICMS", oFont12n ) 
    oPrint:Say(nLinha, 2100, "VALOR", oFont12n ) 

    oPrint:Line( nLinha - 40,  530, nLinha + 10, 530, CLR_BLACK ,"-9") // Linha vertical esquerda
    oPrint:Line( nLinha - 40,  680, nLinha + 10, 680, CLR_BLACK ,"-9") // Linha vertical esquerda
    oPrint:Line( nLinha - 40,  810, nLinha + 10, 810, CLR_BLACK ,"-9") // Linha vertical esquerda
    oPrint:Line( nLinha - 40,  1100, nLinha + 10, 1100, CLR_BLACK ,"-9") // Linha vertical esquerda
    oPrint:Line( nLinha - 40,  1500, nLinha + 10, 1500, CLR_BLACK ,"-9") // Linha vertical esquerda
    oPrint:Line( nLinha - 40,  1880, nLinha + 10, 1880, CLR_BLACK ,"-9") // Linha vertical esquerda
    oPrint:Line( nLinha - 40,  2040, nLinha + 10, 2040, CLR_BLACK ,"-9") // Linha vertical esquerda
    
    oPrint:Line(nLinha + 10,  10, nLinha+10, nXFim, CLR_BLACK ,"-9") // Linha horizontal inferiorZZZZZZZ
    
    nLinha += 40
    For nZ := 1 To len(aChaves)
      SE1->(dbSeek(aChaves[nZ]))
      // oPrint:Say(nLinha, 45, SE1->E1_TIPO +Space(1)+SE1->E1_NUM +Space(1)+ DtoC(SE1->E1_EMISSAO)+Space(1)+SE1->E1_FILIAL+Space(1)+"CIF"+Space(5)+"NF"+Space(10)+"REMETENTE"+Space(10)+"DESTINATARIO"+Space(10)+"2.06"+Space(1)+Alltrim(Transform(SE1->E1_VALOR,"@E 99,999,999.99")) , oFont12n ) 
      oPrint:Say(nLinha, 25, SE1->E1_TIPO   , oFont12n ) 
      oPrint:Say(nLinha, 150, SE1->E1_NUM , oFont12n ) 
      oPrint:Say(nLinha, 350, DtoC(SE1->E1_EMISSAO), oFont12n ) 
      oPrint:Say(nLinha, 540, SE1->E1_FILIAL , oFont12n ) 
      If !PosiTabs()
        oPrint:Say(nLinha, 690, ""  , oFont12n ) 
        oPrint:Say(nLinha, 820, ""     , oFont12n ) 
        oPrint:Say(nLinha, 1110, "", oFont12n ) 
        oPrint:Say(nLinha, 1510, "", oFont12n ) 
        oPrint:Say(nLinha, 1890, "", oFont12n ) 
      Else
        If Alltrim(SE1->E1_TIPO) == 'CTE'
          oPrint:Say(nLinha, 690, IIF(AllTrim(SC5->C5_TPFRETE) == 'C','CIF',IIF(AllTrim(SC5->C5_TPFRETE) == 'F', 'FOB', '') )  , oFont12n ) 
          oPrint:Say(nLinha, 820, GetNFs(SF2->F2_CHVNFE), oFont12n ) 
          SA1->(dbSetOrder(1))
          SA1->(DbSeek(xFilial("SA1") + DT6->( DT6->(DT6_CLIREM + DT6_LOJREM))))
          oPrint:Say(nLinha, 1110, PADR(SA1->A1_NOME,25), oFont10 ) 
          SA1->(DbSeek(xFilial("SA1") + DT6->( DT6->(DT6_CLIDES + DT6_LOJDES))))
          oPrint:Say(nLinha, 1510, PADR(SA1->A1_NOME,25), oFont10 ) 
        Else
          If PosiZA5()
            SA1->(dbSetOrder(1))
            SA1->(DbSeek(xFilial("SA1") + ZA5->( ZA5->(ZA5_REM + ZA5_REMLOJ))))
            oPrint:Say(nLinha, 1110, PADR(SA1->A1_NOME,25), oFont10 ) 
            SA1->(DbSeek(xFilial("SA1") + ZA5->( ZA5->(ZA5_DEST + ZA5_DESLOJ))))
            oPrint:Say(nLinha, 1510, PADR(SA1->A1_NOME,25), oFont10 ) 
          EndIf
        EndIf
        oPrint:Say(nLinha, 1888, Transform(SF2->F2_VALICM,"@E 99,999.99"), oFont12n ) 
      EndIf
      SA1->(RestArea(aAreaSA1))
      oPrint:Say(nLinha, 2055, Transform(SE1->E1_VALOR - RetImp(),"@E 99,999,999.99"), oFont12n )

      oPrint:Line(nLinha + 10,  10, nLinha+10, nXFim, CLR_BLACK ,"-9") // Linha horizontal inferior

      oPrint:Line( nLinha - 40,  530, nLinha + 10, 530, CLR_BLACK ,"-9") // Linha vertical esquerda
      oPrint:Line( nLinha - 40,  680, nLinha + 10, 680, CLR_BLACK ,"-9") // Linha vertical esquerda
      oPrint:Line( nLinha - 40,  810, nLinha + 10, 810, CLR_BLACK ,"-9") // Linha vertical esquerda
      oPrint:Line( nLinha - 40,  1100, nLinha + 10, 1100, CLR_BLACK ,"-9") // Linha vertical esquerda
      oPrint:Line( nLinha - 40,  1500, nLinha + 10, 1500, CLR_BLACK ,"-9") // Linha vertical esquerda
      oPrint:Line( nLinha - 40,  1880, nLinha + 10, 1880, CLR_BLACK ,"-9") // Linha vertical esquerda
      oPrint:Line( nLinha - 40,  2040, nLinha + 10, 2040, CLR_BLACK ,"-9") // Linha vertical esquerda
      
      If nLinha > 2800 .AND. nPag == 1
        //Box linha 3
        oPrint:Line( nRow3 + 25 + 25 + 15,  10, nRow3 + 25 + 25 +15, nXFim, CLR_BLACK ,"10") // Linha horizontal superior
        oPrint:Line( nRow3 + 25 + 25,  10, nLinha +10, 10, CLR_BLACK ,"-9") // Linha vertical esquerda
        oPrint:Line( nLinha +10,  10, nLinha +10, nXFim, CLR_BLACK ,"-9") // Linha horizontal inferior
        oPrint:Line( nRow3 + 25 + 25, nXFim, nLinha + 10, nXFim, CLR_BLACK ,"-9") // Linha vertical Direita
        nPag += 1
        nLinha := 100
        //Descrição das seções
        oPrint:Say(nRow1 + 25 , 350, "DADOS DO BENEFICIÁRIO", oFont14n,,CLR_WHITE)
        oPrint:Say(nRow1 + 25 , 1550, "DADOS DO PAGADOR", oFont14n,,CLR_WHITE)
        oPrint:Say( (nRow2/2) + 35 + 10 , 1560, "DADOS DA FATURA", oFont14n,,CLR_WHITE)
        oPrint:Say( (nRow2/2 + nRow2/4)+ 35 + 10 , 1275, "DESCONTO", oFont14n,,CLR_WHITE)
        oPrint:Say( (nRow2/2 + nRow2/4)+ 35 + 10 , 1625, "ACRÉSCIMO", oFont14n,,CLR_WHITE)
        oPrint:Say( (nRow2/2 + nRow2/4)+ 35 + 10 , 2025, "VALOR", oFont14n,,CLR_WHITE)
        oPrint:Say(  nRow2 + 5 + 25 + 25 , 1000, "OBSERVAÇÕES DA FATURA", oFont14n,,CLR_WHITE)
        oPrint:Say( nRow3 + 25 + 25 + 25 , 900, "DEMONSTRATIVO RESUMIDO DA FATURA", oFont14n,,CLR_WHITE)

        oPrint:EndPage()
        oPrint:StartPage()
      ElseIf nLinha > 2800
        //Box linha 3
        oPrint:Line( 100,  10, 100, nXFim, CLR_BLACK ,"-9") // Linha horizontal superior
        // oPrint:Line( 100,  10, nLinha - 40, 10, CLR_BLACK ,"-9") // Linha vertical esquerda
        // oPrint:Line( nLinha - 40,  10, nLinha - 40, nXFim, CLR_BLACK ,"-9") // Linha horizontal inferior
        // oPrint:Line( 100, nXFim, nLinha - 40, nXFim, CLR_BLACK ,"-9") // Linha vertical Direita
        oPrint:Line( 100,  10, nLinha + 10, 10, CLR_BLACK ,"-9") // Linha vertical esquerda
        oPrint:Line( 100, nXFim, nLinha + 10, nXFim, CLR_BLACK ,"-9") // Linha vertical Direita
        nPag += 1
        nLinha := 100
        oPrint:EndPage()
        oPrint:StartPage()
      EndIf
      nLinha += 40
    Next

    If nRow4 + 20 > nLinha .AND. nPag == 1
      oPrint:Line( nRow3 + 25 + 25 + 15,  10, nRow3 + 25 + 25 + 15, nXFim, CLR_BLACK ,"10") // Linha horizontal superior
      oPrint:Line( nRow3 + 25 + 25,  10, nRow4 + 20, 10, CLR_BLACK ,"-9") // Linha vertical esquerda
      oPrint:Line( nRow4 + 20,  10, nRow4 + 20, nXFim, CLR_BLACK ,"-9") // Linha horizontal inferior
      oPrint:Line( nRow3 + 25 + 25, nXFim, nRow4 + 20, nXFim, CLR_BLACK ,"-9") // Linha vertical Direita
      //Descrição das seções
      oPrint:Say(nRow1 + 25 , 350, "DADOS DO BENEFICIÁRIO", oFont14n,,CLR_WHITE)
      oPrint:Say(nRow1 + 25 , 1550, "DADOS DO PAGADOR", oFont14n,,CLR_WHITE)
      oPrint:Say( (nRow2/2) + 35 + 10 , 1560, "DADOS DA FATURA", oFont14n,,CLR_WHITE)
      oPrint:Say( (nRow2/2 + nRow2/4)+ 35 + 10 , 1275, "DESCONTO", oFont14n,,CLR_WHITE)
      oPrint:Say( (nRow2/2 + nRow2/4)+ 35 + 10 , 1625, "ACRÉSCIMO", oFont14n,,CLR_WHITE)
      oPrint:Say( (nRow2/2 + nRow2/4)+ 35 + 10 , 2025, "VALOR", oFont14n,,CLR_WHITE)
      oPrint:Say(  nRow2 + 5 + 25 + 25 , 1000, "OBSERVAÇÕES DA FATURA", oFont14n,,CLR_WHITE)
      oPrint:Say( nRow3 + 25 + 25 + 25 , 900, "DEMONSTRATIVO RESUMIDO DA FATURA", oFont14n,,CLR_WHITE)
    ElseIf nRow4 + 20 > nLinha .AND. nPag > 1
      oPrint:Line( 100,  10, 100, nXFim, CLR_BLACK ,"-9") // Linha horizontal superior
      oPrint:Line( 100,  10, nRow4 + 20, 10, CLR_BLACK ,"-9") // Linha vertical esquerda
      oPrint:Line( nRow4 + 20,  10, nRow4 + 20, nXFim, CLR_BLACK ,"-9") // Linha horizontal inferior
      oPrint:Line( 100, nXFim, nRow4 + 20, nXFim, CLR_BLACK ,"-9") // Linha vertical Direita  
    ElseIf nPag == 1
      
      oPrint:Line( nRow3 + 25 + 25 + 15,  10, nRow3 + 25 + 25 + 15, nXFim, CLR_BLACK ,"10") // Linha horizontal superior
      oPrint:Line( nRow3 + 25 + 25,  10, nLinha - 30, 10, CLR_BLACK ,"-9") // Linha vertical esquerda
      oPrint:Line( nLinha - 30,  10, nLinha - 30, nXFim, CLR_BLACK ,"-9") // Linha horizontal inferior
      oPrint:Line( nRow3 + 25 + 25, nXFim, nLinha - 30, nXFim, CLR_BLACK ,"-9") // Linha vertical Direita
      //Descrição das seções
      oPrint:Say(nRow1 + 25 , 350, "DADOS DO BENEFICIÁRIO", oFont14n,,CLR_WHITE)
      oPrint:Say(nRow1 + 25 , 1550, "DADOS DO PAGADOR", oFont14n,,CLR_WHITE)
      oPrint:Say( (nRow2/2) + 35 + 10 , 1560, "DADOS DA FATURA", oFont14n,,CLR_WHITE)
      oPrint:Say( (nRow2/2 + nRow2/4)+ 35 + 10 , 1275, "DESCONTO", oFont14n,,CLR_WHITE)
      oPrint:Say( (nRow2/2 + nRow2/4)+ 35 + 10 , 1625, "ACRÉSCIMO", oFont14n,,CLR_WHITE)
      oPrint:Say( (nRow2/2 + nRow2/4)+ 35 + 10 , 2025, "VALOR", oFont14n,,CLR_WHITE)
      oPrint:Say(  nRow2 + 5 + 25 + 25 , 1000, "OBSERVAÇÕES DA FATURA", oFont14n,,CLR_WHITE)
      oPrint:Say( nRow3 + 25 + 25 + 25 , 900, "DEMONSTRATIVO RESUMIDO DA FATURA", oFont14n,,CLR_WHITE)
      
      nPag += 1
      oPrint:EndPage()
      oPrint:StartPage()
    Else
      oPrint:Line( 100,  10, 100, nXFim, CLR_BLACK ,"-9") // Linha horizontal superior
      oPrint:Line( 100,  10, nLinha - 30, 10, CLR_BLACK ,"-9") // Linha vertical esquerda
      oPrint:Line( nLinha - 30,  10, nLinha - 30, nXFim, CLR_BLACK ,"-9") // Linha horizontal inferior
      oPrint:Line( 100, nXFim, nLinha - 30, nXFim, CLR_BLACK ,"-9") // Linha vertical Direita
      nPag += 1
      oPrint:EndPage()
      oPrint:StartPage()
      
    EndIf

    SE1->(RestArea(aAreaSE1))
    If !IsInCallStack('U_LAUA007D')
      nRow3 := -10
      nYBarra := 02850
      For nI := 100 to 2300 step 50
          oPrint:Line(nRow3+1874, nI, nRow3+1874, nI+30) 										// Linha Pontilhada
      Next nI
    
    EndIf
    
  Else
    nYBarra := 01070
    nRow3 := -1800
  EndIf
  If !IsInCallStack('U_LAUA007D')
    // oPrint:SayAlign ( nRow4 + 22 + 25 , 2, Replicate("-",200), oFont15n, nXFim, 100  )

    // //Box linha 4
    // oPrint:Line( nRow4 + 40 + 25 ,  10, nRow4 + 40 + 25 , nXFim, CLR_BLACK ,"-9") // Linha horizontal superior
    // oPrint:Line( nRow4 + 40 + 25 ,  10, nRow5+ 20, 10, CLR_BLACK ,"-9") // Linha vertical esquerda
    // oPrint:Line( nRow5 + 20,  10, nRow5 + 20, nXFim, CLR_BLACK ,"-9") // Linha horizontal inferior
    // oPrint:Line( nRow4 + 40 + 25 , nXFim, nRow5+ 20, nXFim, CLR_BLACK ,"-9") // Linha vertical Direita
    
    //--------------------------------------------------------------------------------------------------------------//
    // Terceiro Bloco - Ficha de Compensação                                                                        //
    //--------------------------------------------------------------------------------------------------------------//
    

    

      oPrint:Line(nRow3 + (2000 - nPosPDF), 100 - 75 , nRow3 + (2000 - nPosPDF), 2300 - 75)
      oPrint:Line(nRow3 + (2000 - nPosPDF), 500 - 75, nRow3 + (1920 - nPosPDF), 0500 - 75)
      oPrint:Line(nRow3 + (2000 - nPosPDF), 710 - 75, nRow3 + (1920 - nPosPDF), 0710 - 75)

      oPrint:SayBitMap(nRow3 + 1884,100 - 75 ,cBmp,380, (110 - nPosPDF))	    					   // Nome do Banco
      oPrint:Say(nRow3 +60 + 1925, 513 - 75, aBanco[01] + "-" + aBanco[07], oFont20)                // Numero do Banco + Dígito
      oPrint:Say(nRow3 +60 + 1925, 755 - 75,aCB_RN_NN[02], oFont20)	                               // Linha Digitavel do Codigo de Barras

      oPrint:Line(nRow3 + (2100 - nPosPDF), 100 - 75 , nRow3 + (2100 - nPosPDF), 2300 - 75 )
      oPrint:Line(nRow3 + (2200 - nPosPDF), 100 - 75 , nRow3 + (2200 - nPosPDF), 2300 - 75 )
      oPrint:Line(nRow3 + (2270 - nPosPDF), 100 - 75 , nRow3 + (2270 - nPosPDF), 2300 - 75 )
      oPrint:Line(nRow3 + (2340 - nPosPDF), 100 - 75 , nRow3 + (2340 - nPosPDF), 2300 - 75 )

      oPrint:Line(nRow3 + (2200 - nPosPDF), 0500 - 75, nRow3 + (2340 - nPosPDF), 0500 - 75 )
      oPrint:Line(nRow3 + (2270 - nPosPDF), 0750 - 75, nRow3 + (2340 - nPosPDF), 0750 - 75 )
      oPrint:Line(nRow3 + (2200 - nPosPDF), 1000 - 75, nRow3 + (2340 - nPosPDF), 1000 - 75 )
      oPrint:Line(nRow3 + (2200 - nPosPDF), 1300 - 75, nRow3 + (2270 - nPosPDF), 1300 - 75 )
      oPrint:Line(nRow3 + (2200 - nPosPDF), 1480 - 75, nRow3 + (2340 - nPosPDF), 1480 - 75 )

      oPrint:Say(nRow3 + 25 + 2000, 100 - 75,"Local de Pagamento",oFont08)
      oPrint:Say(nRow3 + 25 + 2030, 300 - 75,aDadTit[11]         ,oFont12n)

    /*  If aBanco[1] == "104"
        oPrint:Say(nRow3 + 25 + 2020, 400, "PREFERENCIALMENTE NAS CASAS LOTÉRICAS ATÉ O VALOR LIMITE", oFont10)
      else
        oPrint:Say(nRow3 + 25 + 2010, 400, "ATÉ O VENCIMENTO, PREFERENCIALMENTE NO " + aBanco[2], oFont10)
        oPrint:Say(nRow3 + 25 + 2050, 400, "APÓS O VENCIMENTO, SOMENTE NO " + aBanco[2], oFont10)
      EndIf*/	  
              
      oPrint:Say(nRow3 + 25 + 2000, 1810 - 75,"Vencimento", oFont08)
      
      cString := StrZero(Day(aDadTit[4]),2) + "/" + StrZero(Month(aDadTit[4]),2) + "/" + Right(Str(Year(aDadTit[4])),4)
      nCol	   := 1910 + (374 - (len(cString) * 22))
      
      oPrint:Say(nRow3 + 25 + 2040, nCol - 75, cString, oFont11c)           // Vencimento
      
      oPrint:Say(nRow3 + 25 + 2100, 100 - 75, "Beneficiário", oFont08) 

      // If SEE->EE_YFACTOR == 'S' .AND. ! Empty(SEE->EE_YBENEFI)
      //   oPrint:Say(nRow3 + 25 + 2128, 0100 - 75, AllTrim(SEE->EE_YBENEFI), oFont10)                 // Nome da Empresa
      //   oPrint:Say(nRow3 + 25 + 2160, 0100 - 75, AllTrim(SEE->EE_YENDBEN), oFont08)                 // Endereco Beneficiario
      // else
        oPrint:Say(nRow3 + 25 + 2128, 100 - 75, AllTrim(aEmpresa[01]) + " - " + aEmpresa[08], oFont10)        // Nome + CNPJ
        oPrint:Say(nRow3 + 25 + 2160, 100 - 75, AllTrim(aEmpresa[02]) + " - " + AllTrim(aEmpresa[03]) + " - " +;
                                    AllTrim(aEmpresa[04]) + "/" + aEmpresa[05], oFont08)          // Endereço da empresa
      // EndIf                               

      oPrint:Say(nRow3 + 25 + 2100, 1810, "Agência/Código Beneficiário", oFont08)

      If Empty(aBanco[5])
        cString := aBanco[3] + "/" + aBanco[4]    // Agencia + Cód. Beneficiário + Dígito
      elseIf aBanco[1] == "104" .or. aBanco[1] == "033"
              cString := aBanco[3] + "/" + aBanco[9]
            elseIf aBanco[1] == "341"
                  cString := aBanco[3] + "/" + AllTrim(aBanco[4]) + "-" + aBanco[5]
                else        
                  cString := aBanco[3] + IIf(Empty(AllTrim(aBanco[8])),"","-") + aBanco[8] + " / " +;
                              aBanco[4] + IIf(aBanco[1] == "422","","-") + aBanco[5]
      EndIf

      nCol	:= 1910 + (373 - (len(cString) * 22))
      
      oPrint:Say(nRow3 + 25 + 2140, nCol - 75, cString, oFont11c)                         // Agência + Cod. Beneficiário

      oPrint:Say(nRow3 + 25 + 2200, 0100 - 75, "Data do Documento", oFont08)
      oPrint:Say(nRow3 + 25 + 2230, 0140 - 75, StrZero(Day(aDadTit[2]),2) + "/" + StrZero(Month(aDadTit[2]),2) +;
                                    "/" + Right(Str(Year(aDadTit[2])),4), oFont10)	 // Vencimento

      oPrint:Say(nRow3 + 25 + 2200, 0505 - 75, "Nro.Documento", oFont08)
      oPrint:Say(nRow3 + 25 + 2230, 0605 - 75, aDadTit[01], oFont10)	                      // Prefixo + Numero + Parcela

      oPrint:Say(nRow3 + 25 + 2200, 1005 - 75, "Espécie Doc.", oFont08)
      oPrint:Say(nRow3 + 25 + 2230, 1090 - 75, aDadTit[08], oFont10)                       // Tipo do Titulo

      oPrint:Say(nRow3 + 25 + 2200, 1305 - 75, "Aceite", oFont08)
      oPrint:Say(nRow3 + 25 + 2230, 1390 - 75, "N", oFont10)

      oPrint:Say(nRow3 + 25 + 2200, 1485 - 75, "Data do Processamento", oFont08)
      oPrint:Say(nRow3 + 25 + 2230, 1550 - 75, StrZero(Day(aDadTit[03]),2) + "/" + StrZero(Month(aDadTit[03]),2) +;
                                    "/" + Right(Str(Year(aDadTit[03])),4), oFont10)   // Data impressao

      oPrint:Say(nRow3 + 25 + 2200, 1810 - 75, "Nosso Número", oFont08)

      cString := aDadTit[6]
      nCol	   := 1910 + (373 - (len(cString) * 22))
      
      oPrint:Say(nRow3 + 25 + 2230, nCol - 75, cString, oFont11c)	// Nosso Número
      oPrint:Say(nRow3 + 25 + 2270, 100 - 75, "Uso do Banco", oFont08)
      oPrint:Say(nRow3 + 25 + 2270, 505 - 75, "Carteira", oFont08)
      
      //If aBanco[1] == "033"
        //oPrint:Say(nRow3 + 25 + 2300,505 - 75,aBanco[10],oFont07)
      //else
        oPrint:Say(nRow3 + 25 + 2300, 565 - 75,aBanco[10],oFont10)
      //EndIf
      
      oPrint:Say(nRow3 + 25 + 2270, 755 - 75, IIf(aBanco[1] == "104","Espécie Moeda","Espécie"), oFont08)
      oPrint:Say(nRow3 + 25 + 2300, 825 - 75, "R$", oFont10)

      oPrint:Say(nRow3 + 25 + 2270, 1005 - 75, "Qtde Moeda", oFont08)
      oPrint:Say(nRow3 + 25 + 2270, 1485 - 75, "Valor", oFont08)

      oPrint:Say(nRow3 + 25 + 2270, 1810 - 75, "Valor do Documento", oFont08)
      
      cString := Alltrim(Transform(aDadTit[05],"@E 99,999,999.99"))
      nCol	   := 1910 + (374 - (len(cString) * 22))
      
      oPrint:Say(nRow3 + 25 + 2300, nCol - 75, cString, oFont11c)	    // Valor do Documento

      If aBanco[1] == "104"
        oPrint:Say(nRow3 + 25 + 2340, 0100 - 75,"Instruções (Texto de Responsabilidade do Beneficiário):", oFont08)
      else  
        oPrint:Say(nRow3 + 25 + 2340, 0100 - 75, "Instruções (Todas informações deste bloqueto são de exclusiva " +;
                                        "responsabilidade do beneficiário)", oFont08)
      EndIf
                                  
      If Len(aBolTxt) > 0
        oPrint:Say(nRow3 + 25 + 2375, 0100 - 75, aBolTxt[1], oFont08n)	// 1a. Linha Instrução
        oPrint:Say(nRow3 + 25 + 2415, 0100 - 75, aBolTxt[2], oFont08n)	// 2a. Linha Instrução
        oPrint:Say(nRow3 + 25 + 2454, 0100 - 75, aBolTxt[3], oFont08n)	// 3a. Linha Instrução
        oPrint:Say(nRow3 + 25 + 2494, 0100 - 75, aBolTxt[4], oFont08)	// 4a. Linha Instrução
        oPrint:Say(nRow3 + 25 + 2534, 0100 - 75, aBolTxt[5], oFont08)	// 5a. Linha Instrução
        oPrint:Say(nRow3 + 25 + 2574, 0100 - 75, aBolTxt[6], oFont08)	// 6a. Linha Instrução
        oPrint:Say(nRow3 + 25 + 2614, 0100 - 75, aBolTxt[7], oFont08)	// 7a. Linha Instrução
        oPrint:Say(nRow3 + 25 + 2654, 0100 - 75, aBolTxt[8], oFont08)	// 8a. Linha Instrução
      else
      oPrint:Say(nRow3 + 25 + 2375, 0100 - 75, aDadTit[9], oFont08)	   // 1a. Linha Instrução
      oPrint:Say(nRow3 + 25 + 2655, 0100 - 75, aBolTxt[8], oFont08)	   // 8a. Linha Instrução
      EndIf

      oPrint:Say(nRow3 + 25 + 2340, 1810 - 75, "(-)Desconto/Abatimento", oFont08)
      oPrint:Say(nRow3 + 25 + 2410, 1810 - 75, "(-)Outras Deduções", oFont08)
      oPrint:Say(nRow3 + 25 + 2480, 1810 - 75, "(+)Mora/Multa", oFont08)
      oPrint:Say(nRow3 + 25 + 2550, 1810 - 75, "(+)Outros Acréscimos", oFont08)
      oPrint:Say(nRow3 + 25 + 2620, 1810 - 75, "(=)Valor Cobrado",	oFont08)

      oPrint:Say(nRow3 + 25 + 2690, 0100 - 75, IIf(aBanco[1] == "104","Pagador","Nome do Pagador"), oFont08)
      oPrint:Say(nRow3 + 25 + 2690, 0550 - 75, "(" + aSacado[2] + ") " + aSacado[1], oFont08n)	// Nome Cliente + Código

      If Empty(aSacado[7]) 
        oPrint:Say(nRow3 + 25 + 2690, 1850 - 75, "CPF/CNPJ NAO CADASTRADO", oFont08)
      elseIf aSacado[8] == "J" .and. ! Empty(aSacado[7])
              oPrint:Say(nRow3 + 25 + 2690, 1850 - 75, "CNPJ: " + Transform(aSacado[7],"@R 99.999.999/9999-99"), oFont08)	// CGC
            elseIf aSacado[8] == "F" .and. ! Empty(aSacado[7])
                  oPrint:Say(nRow3 + 25 + 2690, 1850 - 75, "CPF: " + Transform(aSacado[7],"@R 999.999.999-99"), oFont08)	// CPF
      EndIf

      oPrint:Say(nRow3 + 25 + 2723, 0550 - 75, aSacado[3], oFont08)	// Endereço

      If Empty(aSacado[6])
        oPrint:Say(nRow3 + 25 + 2763, 0550 - 75, "CEP NAO CADASTRADO - " + aSacado[4] + " - " + aSacado[5], oFont08)
      else
        oPrint:Say(nRow3 + 25 + 2763, 0550 - 75, Transform(aSacado[6],"@R 99999-999") + " - " + aSacado[4] + " - " + ;
                                        aSacado[5], oFont08)	// CEP + Cidade + Estado
      EndIf

      oPrint:Say(nRow3 + 25 + 2763, 1850 - 75, aDadTit[6], oFont08)	// Carteira + Nosso Número
      oPrint:Say(nRow3 + 20 + 2815, 0100 - 75, "Sacador/Avalista", oFont07)
      //Se for banco do tipo Factoring ira imprimir os dados da luzarte(empresa)
      // If SEE->EE_YFACTOR == 'S'
      //   oPrint:Say(nRow3 + 25 + 2815 + 60 , 100, AllTrim(aEmpresa[01]) + " - " + aEmpresa[08], oFont08n)             // Nome + CNPJ
      //   oPrint:Say(nRow3 + 25 + 2815 + 90, 100, AllTrim(aEmpresa[02]) + " - " + AllTrim(aEmpresa[03]) + " - " +;
      //                                 AllTrim(aEmpresa[04]) + "/" + aEmpresa[05], oFont08)                // Endereço da empresa
      // EndIf

      oPrint:Say(nRow3 + 20 + 2815, 1850 - 75, "Código de Baixa", oFont07)

      oPrint:Say(nRow3 + 20 + 2855, 1500 - 75,"Autenticação Mecânica - Ficha de Compensação",		oFont07)		// Texto Fixo

      oPrint:Line(nRow3 + (2000 - nPosPDF), 1800 - 75 , nRow3 + (2690 - nPosPDF), 1800 - 75)
      oPrint:Line(nRow3 + (2410 - nPosPDF), 1800 - 75 , nRow3 + (2410 - nPosPDF), 2300 - 75)
      oPrint:Line(nRow3 + (2480 - nPosPDF), 1800 - 75 , nRow3 + (2480 - nPosPDF), 2300 - 75)
      oPrint:Line(nRow3 + (2550 - nPosPDF), 1800 - 75 , nRow3 + (2550 - nPosPDF), 2300 - 75)
      oPrint:Line(nRow3 + (2620 - nPosPDF), 1800 - 75 , nRow3 + (2620 - nPosPDF), 2300 - 75)
      oPrint:Line(nRow3 + (2690 - nPosPDF), 0100 - 75 , nRow3 + (2690 - nPosPDF), 2300 - 75)
      oPrint:Line(nRow3 + (2850 - nPosPDF), 0100 - 75 , nRow3 + (2850 - nPosPDF), 2300 - 75)

      If nTpImp == 2
        If nTipo = 2
          oPrint:Int25(nYBarra,0070,aCB_RN_NN[1],0.73,40,.F.,.F., oFont08)	// Código de Barras
        else
          oPrint:Int25(nYBarra,0070,aCB_RN_NN[1],0.73,40,.F.,.F., oFont08)	
        EndIf
      else     
        oPrint:Int25(nYBarra,0070,aCB_RN_NN[1],0.73,40,.F.,.F., oFont08)	
      EndIf

    // Calculo do nosso numero mais o digito verificador, para ser gravado no campo E1_NUMBCO // Humberto / Liberato
        
      If ! bRetImp
        dbSelectArea("SE1")

        RecLock("SE1",.F.)
          Replace SE1->E1_PORTADO with Subs(aBanco[1],1,3) 
          Replace SE1->E1_NUMBCO  with cNossoDg
        SE1->(MsUnlock())
      EndIf
    EndIf
    oPrint:EndPage()                                   // Finaliza a página
    
  // cFilePrint := oPrint:cPathPDF + "boletos.PD_"
  // File2Printer(cFilePrint,"PDF")
 
Return

/*----------------------------------------------
--  Função: Impressão do boleto Reduzido.     --
--                                            --
------------------------------------------------*/
Static Function fnImprRd(oPrint, aEmpresa, aDadTit, aBanco, aSacado, aBolTxt, aCB_RN_NN, cNNum)
  Local nI         := 0
  Local cStartPath := GetSrvProfString("StartPath","")
  Local cBmp       := ""

 //Parametros de TFont.New()
 //1.Nome da Fonte (Windows)
 //3.Tamanho em Pixels
 //5.Bold (T/F)
  Local oFont06  := TFont():New("Arial Narrow",9,06,.T.,.F.,5,.T.,5,.T.,.F.)
  Local oFont07  := TFont():New("Arial Narrow",9,07,.T.,.F.,5,.T.,5,.T.,.F.)
  Local oFont07n := TFont():New("Arial Narrow",9,07,.T.,.T.,5,.T.,5,.T.,.F.)
  Local oFont08c := TFont():New("Courier New" ,9,08,.T.,.T.,5,.T.,5,.T.,.F.)
  Local oFont08c := TFont():New("Courier New" ,9,08,.T.,.T.,5,.T.,5,.T.,.F.)
  Local oFont12n := TFont():New("Arial Narrow",12,14 ,.T.,.T.,5,.T.,5,.T.,.F.)
  Local oFont11  := TFont():New("Arial Narrow",14,11,.T.,.F.,5,.T.,5,.T.,.F.)
 // -----------------------
 
 
 
  cStartPath := AllTrim(cStartPath) + "logo_bancos"
 
  If SubStr(cStartPath, Len(cStartPath), 1) <> "\"
     cStartPath += "\"
  EndIf

  cBmp	:= cStartPath + cLogo + ".bmp"
 	
  If nNumPag == 1
     nRow  := 0
     nCols := 0
     nWith := 0
   elseIf nNumPag > 3
		   oPrint:StartPage()   // Inicia uma nova página
		   nNumPag := 1
		   nRow    := 0
          nCols   := 0
          nWith   := 0
        else
          nRow  += 1050
          nCols := 0
          nWith := 0
  EndIf
	
  nNumPag++

 // ---- Canhoto
  oPrint:Line(nRow + 150, 100, nRow + 150, 600)
  oPrint:Line(nRow + 270, 100, nRow + 270, 600)

  oPrint:Line(nRow + 335, 100, nRow + 335, 600)
  oPrint:Line(nRow + 400, 100, nRow + 400, 600)
  oPrint:Line(nRow + 465, 100, nRow + 465, 600)
  oPrint:Line(nRow + 530, 100, nRow + 530, 600)
  oPrint:Line(nRow + 595, 100, nRow + 595, 600)
  oPrint:Line(nRow + 660, 100, nRow + 660, 600)
  oPrint:Line(nRow + 725, 100, nRow + 725, 600)
  oPrint:Line(nRow + 790, 100, nRow + 790, 600)
  oPrint:Line(nRow + 855, 100, nRow + 855, 600)
  oPrint:Line(nRow + 920, 100, nRow + 920, 600)
	
 // ---- Linha Pontilhada
  For nI := 100 To 1030 Step 10
      oPrint:Line(nRow + nI + 50, 700, nRow + nI + 50, 702)
  Next nI
	
 // ---- Boleto (Horizontal)
  oPrint:Line(nRow + 150, 800, nRow + 150, 2300)
  oPrint:Line(nRow + 225, 800, nRow + 225, 2300)
  oPrint:Line(nRow + 300, 800, nRow + 300, 2300)
  oPrint:Line(nRow + 375, 800, nRow + 375, 2300)
  oPrint:Line(nRow + 450, 800, nRow + 450, 2300)
  oPrint:Line(nRow + 750, 800, nRow + 750, 2300)
  oPrint:Line(nRow + 920, 800, nRow + 920, 2300)

 // ---- Traços Direita - Horizontal
  oPrint:Line(nRow + 510, 1900, nRow + 510, 2300)
  oPrint:Line(nRow + 570, 1900, nRow + 570, 2300)
  oPrint:Line(nRow + 630, 1900, nRow + 630, 2300)
  oPrint:Line(nRow + 690, 1900, nRow + 690, 2300)

 // ---- Vertical
  oPrint:Line(nRow + 300,  995, nRow + 450,  995)
  oPrint:Line(nRow + 375, 1130, nRow + 450, 1130)
  oPrint:Line(nRow + 300, 1280, nRow + 450, 1280)
  oPrint:Line(nRow + 300, 1430, nRow + 375, 1430)
  oPrint:Line(nRow + 225, 1580, nRow + 450, 1580)
  oPrint:Line(nRow + 150, 1900, nRow + 750, 1900)
  
 // ---- Traços Banco - Vertical
  oPrint:Line(nRow + 080, 1180, nRow + 150, 1180)
  oPrint:Line(nRow + 080, 1325, nRow + 150, 1325)
	
 // ---- Texto Canhoto
  oPrint:SayBitMap(nRow + 050,160,cBmp,330,90)					         // Logo Canhoto
	
  oPrint:Say(nRow + 155,110,"Beneficiário",oFont07)			

  // If SEE->EE_YFACTOR == 'S' .AND. ! Empty(SEE->EE_YBENEFI)
	//   oPrint:Say(nRow + 180, 110, AllTrim(SEE->EE_YBENEFI), oFont06)                 // Nome da Empresa
  // else
	  oPrint:Say(nRow + 180,110, AllTrim(aEmpresa[01]), oFont06)             // Nome 
	  oPrint:Say(nRow + 210,110, AllTrim(aEmpresa[02]) + " - " + AllTrim(aEmpresa[03]) + " - " +;
	                             AllTrim(aEmpresa[04]) + "/" + aEmpresa[05], oFont06)                // Endereço da empresa
	  oPrint:Say(nRow + 240,110, AllTrim(aEmpresa[08]), oFont06)             // CNPJ
  // EndIf
  oPrint:Say(nRow + 275,110,"Nro.Documento",oFont07)
  oPrint:Say(nRow + 310,600,aDadTit[1]     ,oFont08c,,,,1)	             // Prefixo + Numero + Parcela
	
  oPrint:Say(nRow + 340,110,"Vencimento",oFont07)
  
  cString := StrZero(Day(aDadTit[4]),2) + "/" + StrZero(Month(aDadTit[4]),2) + "/" + Right(Str(Year(aDadTit[4])),4)
  nCol	   := 150 + (374 - (Len(cString) * 22))
  
  oPrint:Say(nRow + 375,600,cString,oFont08c,,,,1)                      // Vencimento
	
  oPrint:Say(nRow + 405,110,"Agência/Código Beneficiario",oFont07)

  If aBanco[1] == "104" .or. aBanco[1] == "033"
     cString := AllTrim(aBanco[3]) + "/" + AllTrim(aBanco[9])
   elseIf aBanco[1] == "341" 
          cString := AllTrim(aBanco[3]) + "/" + AllTrim(aBanco[4]) + "-" + AllTrim(aBanco[5])
        else
          cString := AllTrim(aBanco[3]) + IIf(Empty(AllTrim(aBanco[8])),"","-") + AllTrim(aBanco[8]) + " / " +;
                     AllTrim(aBanco[4]) + IIf(aBanco[1] == "422","","-") + AllTrim(aBanco[5])
  EndIf

  nCol	:= 150 + (374 - (Len(cString) * 22))
  
  oPrint:Say(nRow + 440,600,cString,oFont08c,,,,1)
	
  oPrint:Say(nRow + 470,110,"Nosso Número",oFont07)
  
  cString := AllTrim(aDadTit[6])
  nCol    := 150 + (374 - (Len(cString) * 22))

  oPrint:Say(nRow + 505,600,cString,oFont08c,,,,1)	             // Nosso Número
	
  oPrint:Say(nRow + 535,110,"Valor do Documento",oFont07)
	
  cString := AllTrim(Transform(aDadTit[5],"@E 99,999,999.99"))
  nCol    := 150 + (374 - (Len(cString) * 22))
  oPrint:Say(nRow + 568,600,cString,oFont08c,,,,1)	
	
  oPrint:Say(nRow + 600,110,"(-)Desconto/Abatimento",oFont07)
  
  If nDesc > 0
     cString := Alltrim(Transform(nDesc,"@E 99,999,999.99"))
     nCol    := 1950+(374-(len(cString)*22))
  
     oPrint:Say(nRow + 633,600,cString, oFont08c,,,,1)	
  EndIf
  	
  oPrint:Say(nRow + 665,110,"(-)Outras Deduções",oFont07)
  
  oPrint:Say(nRow + 730,110,"(+)Mora/Multa",oFont07)	
  
  If nJurMul > 0
     cString := Alltrim(Transform(nJurMul,"@E 99,999,999.99"))
     nCol    := 1950+(374-(len(cString)*22))
     
     oPrint:Say(nRow + 763,600,cString,oFont08c,,,,1)	
  EndIf
  
  oPrint:Say(nRow + 795,110,"(+)Outros Acréscimos",oFont07)
  oPrint:Say(nRow + 860,110,"(=)Valor Cobrado",oFont07)
  	
  oPrint:Say(nRow + 925,110,"Pagador:",oFont07)
  oPrint:Say(nRow + 960,150,aSacado[1],oFont08c)               

  If Empty(aSacado[7]) 
     oPrint:Say(nRow + 995,150,"CPF/CNPJ NAO CADASTRADO",oFont08)
     
   elseIf aSacado[8] == "J" .and. ! Empty(aSacado[7])
          oPrint:Say(nRow + 995,150,"CNPJ: " + Transform(aSacado[7],"@R 99.999.999/9999-99"),oFont07)     // CGC
          
        elseIf aSacado[8] == "F" .and. ! Empty(aSacado[7])
               oPrint:Say(nRow + 995,150,"CPF: " + Transform(aSacado[7],"@R 999.999.999-99"),oFont07)  // CPF
  EndIf
	
 // -----------------------
 // ---- Texto do Boleto
 // -----------------------
  oPrint:SayBitMap(nRow + 050,800,cBmp,330,090)                           // Logo Boleto
  
  oPrint:Say(nRow + 095,1212,aBanco[1] + "-" + aBanco[7],oFont12n)        // Numero do Banco + Dígito
  oPrint:Say(nRow + 100,1335,aCB_RN_NN[2],oFont11)                        // Linha Digitavel do Codigo de Barras
	
  oPrint:Say(nRow + 155,810,"Local de Pagamento",oFont07)
  oPrint:Say(nRow + 190,850,aDadTit[11]         ,oFont07)

/*If aBanco[1] == "237"
     oPrint:Say(nRow + 240,850,"Pagável preferencialmente na Rede Bradesco ou Bradesco Expresso",oFont07)
   else  
     oPrint:Say(nRow + 240,850,"PAGAVEL EM QUALQUER BANCO ATÉ O VENCIMENTO",oFont07)
  EndIf*/   
	
  oPrint:Say(nRow + 155,1910,"Vencimento",oFont07)
  
  cString := StrZero(Day(aDadTit[4]),2) + "/" + StrZero(Month(aDadTit[4]),2) + "/" + Right(Str(Year(aDadTit[4])),4)
  nCol    := 1950 + (374 - (Len(cString) * 22))

  oPrint:Say(nRow + 190,nCol,cString,oFont08c)	                                                // Vencimento
	
  oPrint:Say(nRow + 230,810,"Beneficiário",oFont07)
                        
  // If SEE->EE_YFACTOR == 'S' .AND. ! Empty(SEE->EE_YBENEFI)
	//   oPrint:Say(nRow + 235, 930, AllTrim(SEE->EE_YBENEFI), oFont08n)                 // Nome da Empresa
  // Else
  oPrint:Say(nRow + 235,930, AllTrim(aEmpresa[01]), oFont07n)             // Nome 
  oPrint:Say(nRow + 270,850, AllTrim(aEmpresa[02]) + " - " + AllTrim(aEmpresa[03]) + " - " +;
                             AllTrim(aEmpresa[04]) + "/" + aEmpresa[05], oFont07)                // Endereço da empresa
  oPrint:Say(nRow + 230,1585,"CNPJ",oFont07)
  oPrint:Say(nRow + 265,1630,Substr(aEmpresa[8],7,(Len(aEmpresa[8]) - 7)),oFont07)		          // CNPJ
  // EndIf	
  oPrint:Say(nRow + 230,1910,"Agência/Código Beneficiário",oFont07)
	
  If aBanco[1] == "104" .or. aBanco[1] == "033"
     cString := AllTrim(aBanco[3]) + "/" + AllTrim(aBanco[9])
   elseIf aBanco[1] == "341" 
          cString := AllTrim(aBanco[3]) + "/" + AllTrim(aBanco[4]) + "-" + AllTrim(aBanco[5])
        else
          cString := AllTrim(aBanco[3]) + IIf(Empty(AllTrim(aBanco[8])),"","-") + AllTrim(aBanco[8]) + " / " +;
                     AllTrim(aBanco[4]) + IIf(aBanco[1] == "422","","-") + AllTrim(aBanco[5])
  EndIf
	
  nCol	:= 1950 + (374 - (Len(cString) * 22))
  oPrint:Say(nRow + 265,nCol,cString,oFont08c)	// Agência + Cod. Cedente
	
  oPrint:Say(nRow + 305,810,"Data do Documento",oFont07)
  oPrint:Say(nRow + 340,850,StrZero(Day(aDadTit[2]),2) + "/" + StrZero(Month(aDadTit[2]),2) +;
                            "/" + Right(Str(Year(aDadTit[2])),4),oFont07)               	// Vencimento
	
  oPrint:Say(nRow + 305,1000,"Nro.Documento",oFont07)
  oPrint:Say(nRow + 340,1040,aDadTit[1],oFont07)	                                          // Prefixo + Numero + Parcela
	
  oPrint:Say(nRow + 305,1285,"Espécie Doc.",oFont07)
  oPrint:Say(nRow + 340,1325,aDadTit[8],oFont07)                                           // Tipo do Titulo
	
  oPrint:Say(nRow + 305,1435,"Aceite",oFont07)
  oPrint:Say(nRow + 340,1475,"N",oFont07)
	
  oPrint:Say(nRow + 305,1585,"Data do Processamento",oFont07)
  oPrint:Say(nRow + 340,1625,StrZero(Day(aDadTit[3]),2) + "/" + StrZero(Month(aDadTit[3]),2) +;
		                      "/" + Right(Str(Year(aDadTit[3])),4),oFont07)                // Data impressao
	
  oPrint:Say(nRow + 305,1910,"Nosso Número",oFont07)
	
  cString := AllTrim(aDadTit[6])
  nCol    := 1950 + (374 - (Len(cString) * 22))
  
  oPrint:Say(nRow + 340,nCol,cString,oFont08c)	                                         // Nosso Número
	
  oPrint:Say(nRow + 380, 810,"Uso do Banco",oFont07)
  
  oPrint:Say(nRow + 380,1000,"Carteira"    ,oFont07)

  //If aBanco[1] == "033"
     //oPrint:Say(nRow + 415,1000,aBanco[10],oFont07)
   //else  
     oPrint:Say(nRow + 415,1040,aBanco[10],oFont07)
  //EndIf
     
  oPrint:Say(nRow + 380,1135,"Espécie"     ,oFont07)
  oPrint:Say(nRow + 415,1175,"R$"          ,oFont07)
  oPrint:Say(nRow + 380,1285,"Quantidade"  ,oFont07)
  oPrint:Say(nRow + 380,1585,"Valor"       ,oFont07)
	
  oPrint:Say(nRow + 380,1910,"Valor do Documento",oFont07)
	
  cString := AllTrim(Transform(aDadTit[5],"@E 99,999,999.99"))
  nCol    := 2350 - 85
  
  oPrint:Say(nRow + 415,nCol,cString,oFont08c)	                                        // Valor do Documento

  oPrint:Say(nRow + 455,0810,"Instruções (Todas informações deste bloqueto são de exclusiva " +;
                             "responsabilidade do beneficiário)", oFont07)
                             
  If Len(aBolTxt) > 0
     oPrint:Say(nRow + 500,0820,aBolTxt[1], oFont08c)	// 1a Linha Instrução
     oPrint:Say(nRow + 545,0820,aBolTxt[2], oFont08c)	// 2a. Linha Instrução
     oPrint:Say(nRow + 590,0820,aBolTxt[3], oFont08c)	// 3a. Linha Instrução
     oPrint:Say(nRow + 635,0820,aBolTxt[4], oFont07)	// 4a Linha Instrução
     oPrint:Say(nRow + 680,0820,aBolTxt[5], oFont07)	// 5a. Linha Instrução
     oPrint:Say(nRow + 725,0820,aBolTxt[6], oFont07)	// 6a. Linha Instrução
     oPrint:Say(nRow + 770,0820,aBolTxt[7], oFont07)	// 7a. Linha Instrução
     oPrint:Say(nRow + 815,0820,aBolTxt[8], oFont07)	// 8a. Linha Instrução
   else
	 oPrint:Say(nRow + 500,0820,aDadTit[9], oFont07)	// 1a Linha Instrução
	 oPrint:Say(nRow + 545,0820,aBolTxt[8], oFont07)	// 8a. Linha Instrução
  EndIf

  oPrint:Say(nRow + 455,1905,"(-)Desconto/Abatimento",oFont07)

  If nDesc > 0  
     cString := Alltrim(Transform(nDesc,"@E 99,999,999.99"))
     nCol    := 2350 - 85
  
     oPrint:Say(nRow + 460,nCol,cString,oFont08c)                                     
  EndIf
  
  oPrint:Say(nRow + 515,1905,"(-)Outras Deduções",oFont07)
  oPrint:Say(nRow + 575,1905,"(+)Mora/Multa",oFont07)
  
  If nJurMul > 0
     cString := Alltrim(Transform(nJurMul,"@E 99,999,999.99"))
     nCol    := 2350 - 85
  
     oPrint:Say(nRow + 580,nCol,cString,oFont08c)
  EndIf
  
  oPrint:Say(nRow + 635,1905,"(+)Outros Acréscimos",oFont07)
  oPrint:Say(nRow + 695,1905,"(=)Valor Cobrado",oFont07)
   
  oPrint:Say(nRow + 755,810,"Pagador:",oFont07)
  oPrint:Say(nRow + 790,850,aSacado[1] + Space(05) + " - " + IIf(Empty(aSacado[7]),"CPF/CNPJ NAO CADASTRADO",;
                                                              IIf(aSacado[8] == "J","CNPJ: " + Transform(aSacado[7],"@R 99.999.999/9999-99"),;
                                                                                    "CPF: " + Transform(aSacado[7],"@R 999.999.999-99"))),oFont07)
  oPrint:Say(nRow + 825,850,aSacado[3],oFont07)	                          // Endereço

  If Empty(aSacado[6])
     oPrint:Say(nRow + 855,850,"CEP NAO CADASTRADO - " + aSacado[4] + " - " + aSacado[5],oFont07)
   else
     oPrint:Say(nRow + 855,850,Transform(aSacado[6],"@R 99999-999") + " - " + aSacado[4] + "/" + aSacado[5],oFont07)	// CEP + Cidade + Estado
  EndIf

  oPrint:Say(nRow + 0890,0810,"Avalista:",oFont07)
  oPrint:Say(nRow + 0930,2065,"Autenticação Mecânica",oFont07)
  oPrint:Say(nRow + 0960,2065,"Ficha de Compensação",oFont07)
 // ----- Código de Barras
 
// MSBAR("INT25",13.0,1.0,aCB_RN_NN[1],oPrint,.F.,Nil,Nil,0.013,0.7,Nil,Nil,"A",.F.)				// Código de Barras
  
  If nNumPag == 2
    If cTpImpre == "2"    // Envia PDF
       oPrint:FWMSBAR("INT25",7.9,7,aCB_RN_NN[1],oPrint,.T.,,.T.,0.023,1.16,.T.,"Arial",NIL,.F.,2,2,.F.)
     else
       MSBAR("INT25",7.9,7,aCB_RN_NN[1],oPrint,.F.,Nil,Nil,0.023,1.16,Nil,Nil,"A",.F.)				// Código de Barras
    EndIf
   elseIf nNumPag == 3
         If cTpImpre == "2"    // Envia PDF
            oPrint:FWMSBAR("INT25",16.9,7,aCB_RN_NN[1],oPrint,.T.,,.T.,0.023,1.16,.T.,"Arial",NIL,.F.,2,2,.F.)
          else 
            MSBAR("INT25",16.9,7,aCB_RN_NN[1],oPrint,.F.,Nil,Nil,0.023,1.16,Nil,Nil,"A",.F.)			// Código de Barras
          // MSBAR("INT25",17.9,7,aCB_RN_NN[1],oPrint,.F.,Nil,Nil,0.021,1.13,Nil,Nil,"A",.F.)
         EndIf
        elseIf nNumPag == 4
             If cTpImpre == "2"    // Envia PDF
                oPrint:FWMSBAR("INT25",25.7,7,aCB_RN_NN[1],oPrint,.T.,,.T.,0.023,1.16,.T.,"Arial",NIL,.F.,2,2,.F.)
              else 
                MSBAR("INT25",25.7,7,aCB_RN_NN[1],oPrint,.F.,Nil,Nil,0.023,1.16,Nil,Nil,"A",.F.)	// Código de Barras
             // MSBAR("INT25",27.3,7  ,aCB_RN_NN[1],oPrint,.F.,Nil,Nil,0.021,1.13,Nil,Nil,"A",.F.)
             EndIf
  EndIf
    
  If ! bRetImp
     dbSelectArea("SE1")

     RecLock("SE1",.F.)
       Replace SE1->E1_PORTADO with Subs(aBanco[1],1,3) 
       Replace SE1->E1_NUMBCO  with cNossoDg
     SE1->(MsUnlock())
  EndIf
  
  If nNumPag > 3
     oPrint:EndPage()
  EndIf

//  // --- Gravar boleto em PDF
//   If nTpImp == 6 
//      cFilePrint := cDirGer + "BOL_" + Substr(aSacado[2],1,TamSX3("A1_COD")[1]) +;
//                    SubStr(aSacado[2],(TamSX3("A1_COD")[1] + 1),TamSX3("A1_LOJA")[1]) + "_" + DToS(aDadTit[4]) + ".PD_"
//      File2Printer(cFilePrint,"PDF")
//      oPrint:cPathPDF := cDirGer 
//   EndIf
Return

/*----------------------------------------------
--  Função: Calculo do digito pelo Modulo10.  --
--                                            --
------------------------------------------------*/
Static Function Modulo10(cData)
  Local L,D,P := 0
  Local B     := .F.

  L := Len(cData)
  B := .T.
  D := 0

  While L > 0
	 P := Val(SubStr(cData, L, 1))
	 
	 If (B)
		 P := P * 2
	 	 If P > 9
           P := P - 9
		 EndIf
 	 EndIf
	 
	 D := D + P
	 L := L - 1
	 B := !B
  EndDo
  
  D := 10 - (Mod(D,10))
  
  If D == 10
     D := 0
  EndIf
Return(D)

/*----------------------------------------------
--  Função: Calculo do digito pelo Modulo11.  --
--                                            --
------------------------------------------------*/
Static Function Modulo11(cData,nPeso,cOrig)
  Local L, D, P := 0

  L := Len(cdata)
  D := 0
  P := 1

  While L > 0
    P := P + 1
    D := D + (Val(SubStr(cData, L, 1)) * P)

    If P = nPeso
       P := 1
    EndIf
    
    L := L - 1
  EndDo

  If cQualBco == "033" .and. Alltrim(cOrig) == "NN"
     If mod(D,11) < 2
        Return(0)
      elseIf mod(D,11) == 10
             Return(1)
     EndIf        
  EndIf            
     
  D := 11 - (mod(D,11))

  If cQualBco == "104"
    If D > 9
      If bOrigCB     
         D := 1
       else
         D := 0
      EndIf     
     elseIf (D == 0 .Or. D == 1 .Or. D == 10)
            D := 1
    EndIf
   elseIf cQualBco == "237"
 	    If Alltrim(cOrig) == 'NN'
          If D == 11 //Se o resto for 11, o digito verificador será 0
             D := 0
		   EndIf
		
		   If D == 10 //Se o resto for 10, o dígito verificador será P
             D := "P"
		   EndIf
	     else	
          If D == 0 .or. D == 1 .or. D == 10 .or. D == 11
             D := 1
          EndIf
	    EndIf
	  else
         If D == 0 .or. D == 1 .or. D == 10 .or. D == 11
           D := 1
         EndIf
  EndIf
Return(D)

/*---------------------------------------------------------
--  Função: Montar código de barra.                      --
--          Campo Livre:                                 --
--            Caixa - Conta                              --
--                    Digito da conta                    --
--                    Nosso numero (1:3)                 --
--                    Carteira (1:1)                     --
--                    Nosso numero (4:3)                 --
--                    Carteira (2:1)                     --
--                    Nosso numero (7:9)                 --
--            BRADESCO - Agencia - tamanho 4             --
--                       Carteira - tamanho 2            --
--                       Nosso numero                    --
--                       Conta - tamanho 7 (sem digito)  --
-----------------------------------------------------------*/
Static Function Ret_cBarra(pBanco,pAgencia,pConta,pDacCC,pCart,pNNum,pValor,pVencto,pConvenio,pModDig,pPesoDig)
  Local nId         := 0
  Local nId1        := 0

  Private cBanco      := pBanco
  Private cAgencia    := pAgencia
  Private cConta      := pConta
  Private cDacCC      := pDacCC
  Private cCart       := pCart
  Private nValor      := pValor
  Private dVencto     := pVencto
  Private cConvenio   := pConvenio
  Private cModDig     := pModDig
  Private nPesoDig    := pPesoDig
  Private nDvnn       := 0
  Private nDvcb       := 0
  Private nDv         := 0
  Private nDvCl       := 0
  Private cNNRet      := ""
  Private cNNSE1      := ""
  Private cCB         := ""
  Private cS          := ""
  Private cCmpLv      := ""
  Private cFator      := StrZero(dVencto - CToD("29/05/2022"),4) //StrZero(dVencto - CToD("07/10/97"),4)
  Private cValorFinal := StrZero((nValor*100),10) //StrZero(Int(nValor*100),10)

  cNNum    := pNNum
  cQualBco := cBanco
  bOrigCB  := .F.
   
 // ---- Nosso Numero
 // -----------------
  If cVersao == "11"            
     aLinDig[01] := fnResolP11(aLinDig[01])
  EndIf
  
  cNN := &(aLinDig[01])

  If ! Empty(aLinDig[02]) 
     If cVersao == "11"            
        aLinDig[02] := fnResolP11(aLinDig[02])
     EndIf

     cS := &(aLinDig[02])

     If cModDig == "11"
        nDvnn := modulo11(cS,nPesoDig,"NN")
      else            
        nDvnn := modulo10(cS)
     EndIf                    
  EndIf
      
  If cVersao == "11"            
     aLinDig[03] := fnResolP11(aLinDig[03])
  EndIf
  
  cNNRet   := &(aLinDig[03]) 
  cNNSE1   := cNNRet

  If ValType(nDvnn) == "N"
    If cBanco == '001'
      cNossoDg := StrZero(Val(AllTrim(cNNum)),TamSX3("E1_NUMBCO")[1])
    Else
     cNossoDg := StrZero(Val(AllTrim(cNNum) + AllTrim(Str(nDvnn))),TamSX3("E1_NUMBCO")[1])
    EndIF
   else
     If cBanco == '001'
      cNossoDg := StrZero(Val(AllTrim(cNNum)),(TamSX3("E1_NUMBCO")[1]))
     Else
      cNossoDg := StrZero(Val(AllTrim(cNNum)),(TamSX3("E1_NUMBCO")[1] - 1))
      cNossoDg := cNossoDg + nDvnn
     EndIf 
  EndIf
  
 // ---- Campo Livre
 // ----------------
  If ! Empty(aLinDig[08])
     If cVersao == "11"            
        aLinDig[08] := fnResolP11(aLinDig[08])
     EndIf
  
     cS     := &(aLinDig[08])
     cCmpLv := &(aLinDig[08])
     nDvCl := modulo11(cS,9,"")
  EndIf

 // ---- Campo 1
 // ------------
  If cVersao == "11"            
     aLinDig[04] := fnResolP11(aLinDig[04])
  EndIf
 
  cS  := &(aLinDig[04])
  nDv := modulo10(cS)
  cRN1 := SubStr(cS,1,5) + "." + SubStr(cS,6,4) + AllTrim(Str(nDv)) + " "

 // ---- Campo 2
 // ------------
  If cVersao == "11"            
     aLinDig[05] := fnResolP11(aLinDig[05])
  EndIf
 
  cS   := &(aLinDig[05])
  nDv  := modulo10(cS)
  cRN2 := cRN1 + SubStr(cS,1,5) + "." + SubStr(cS,6,5) + AllTrim(Str(nDv)) + " "

 // ---- Campo 3
 // ------------
  If cVersao == "11"            
     aLinDig[06] := fnResolP11(aLinDig[06])
  EndIf
 
  cS  := &(aLinDig[06])
  nDv := modulo10(cS)
  cRN3 := cRN2 + SubStr(cS,1,5) + "." + SubStr(cS,6,5) + AllTrim(Str(nDv)) + " "

 // ---- Campo 4
 // ------------
  bOrigCB := .T.

  If cVersao == "11"            
     aLinDig[07] := fnResolP11(aLinDig[07])
  EndIf
  
  cS      := &(aLinDig[07])
  nDvcb   := modulo11(cS,9,"")
  cCB     := SubStr(cS,1,4) + AllTrim(Str(nDvcb)) + SubStr(cS,5,39)
     
  cRN4 := cRN3 + AllTrim(Str(nDvcb)) + " "

 // ---- Campo 5
 // ------------ 
  cRN5 := cRN4 + cFator + StrZero((nValor * 100),14-Len(cFator))
  
  RecLock('SE1', .F.)
    SE1->E1_YNN := cNNSE1
  SE1->(MsUnLock())

Return({cCB,cRN5,cNNRet,cNNSE1})
   
/*====================================================
--  Função: Converter variavél da linha digitável   --
--          para string. PROTHEUS 11.               --
======================================================*/
Static Function fnResolP11(pString)
  Local nId     := 0
  Local nId1    := 0
  Local cString := pString
  Local cResult := ""
  
  Private cVariavel := "" 
     
  For nId := 1 To Len(cString)
      If Substr(cString,nId,1) == "#"
         cVariavel := ""
         
         nId++
         
         For nId1 := nId To Len(cString)
             If SubStr(cString,nId1,1) == "#"
                nId := nId1 + 1
                Exit
             EndIf
                
             cVariavel += Substr(cString,nId1,1) 
         Next
                
         If cVariavel == "NDVNN" .or. cVariavel == "NDVCL"
            cResult += IIf(ValType(&(cVariavel)) == "C","'" + &(cVariavel) + "'",AllTrim(Str(&(cVariavel))))
          else
            cResult += "'" + &(cVariavel) + "'"
         EndIf
         
         If nId > Len(cString)
            Exit
         EndIf   
      EndIf
      
      cResult += SubStr(cString,nId,1)
  Next 
Return cResult  

/*==================================
--  Função: Gravação do Bordero   --
--                                --
====================================*/
Static Function fnGrvBrd()
  If ! Empty(SE1->E1_NUMBOR)
     dbSelectArea("SEA")
     SEA->(dbSetOrder(1))
     
     If SEA->(dbSeek(xFilial("SEA") + SE1->E1_NUMBOR + SE1->E1_PREFIXO + SE1->E1_NUM + SE1->E1_PARCELA + SE1->E1_TIPO))
        RecLock("SEA",.F.)
          dbDelete()
        SEA->(MsUnlock())
     EndIf
  EndIf        
     
  RecLock("SEA",.T.)
    Replace SEA->EA_FILIAL  with xFilial("SEA")
    Replace SEA->EA_NUMBOR  with cNumBor
    Replace SEA->EA_DATABOR with dDataBase
    Replace SEA->EA_PORTADO with mv_par19
    Replace SEA->EA_AGEDEP  with mv_par20
    Replace SEA->EA_NUMCON  with mv_par21
    Replace SEA->EA_NUM     with SE1->E1_NUM
    Replace SEA->EA_PARCELA with SE1->E1_PARCELA
    Replace SEA->EA_PREFIXO with SE1->E1_PREFIXO
    Replace SEA->EA_TIPO    with SE1->E1_TIPO
    Replace SEA->EA_CART    with "R"
    Replace SEA->EA_SITUACA with "1"
    Replace SEA->EA_FILORIG with SE1->E1_FILORIG
    Replace SEA->EA_SITUANT with "0"
    Replace SEA->EA_ORIGEM  with ""
  SEA->(MsUnlock())
  
  FKCOMMIT()
				
  RecLock("SE1",.F.)
    Replace SE1->E1_SITUACA with "1"
    Replace SE1->E1_NUMBOR  with cNumBor
    Replace SE1->E1_DATABOR with dDataBase
    Replace SE1->E1_MOVIMEN with dDataBase

  // DDA - Debito Direto Autorizado
    If SE1->E1_OCORREN $ "53/52"
       Replace SE1->E1_OCORREN with "01"
    Endif
  // ------------------------------
  SE1->(MsUnlock())
Return

/*--------------------------------
--  Função: Cria pergunta.      --
--                              --
----------------------------------*/
Static Function fnCriaSx1(aRegs)
  Local aAreaAtu := GetArea()
  Local aAreaSX1 := SX1->(GetArea())
  Local nJ		   := 0
  Local nY       := 0

 // ---- Monta array com as perguntas
  aAdd(aRegs,{cPerg,"01","Prefixo Inicial   ","","","mv_ch1","C",TamSX3("E1_PREFIXO")[1] ,0,0,"G","","MV_PAR01","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""})
  aAdd(aRegs,{cPerg,"02","Prefixo Final     ","","","mv_ch2","C",TamSX3("E1_PREFIXO")[1] ,0,0,"G","","MV_PAR02","","","","ZZZ","","","","","","","","","","","","","","","","","","","","","","","","",""})
  aAdd(aRegs,{cPerg,"03","Numero Inicial    ","","","mv_ch3","C",TamSX3("E1_NUM")[1]     ,0,0,"G","","MV_PAR03","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""})
  aAdd(aRegs,{cPerg,"04","Numero Final      ","","","mv_ch4","C",TamSX3("E1_NUM")[1]     ,0,0,"G","","MV_PAR04","","","","ZZZZZZZZZZ","","","","","","","","","","","","","","","","","","","","","","","","",""})
  aAdd(aRegs,{cPerg,"05","Parcela Inicial   ","","","mv_ch5","C",TamSX3("E1_PARCELA")[1] ,0,0,"G","","MV_PAR05","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""})
  aAdd(aRegs,{cPerg,"06","Parcela Final     ","","","mv_ch6","C",TamSX3("E1_PARCELA")[1] ,0,0,"G","","MV_PAR06","","","","ZZZ","","","","","","","","","","","","","","","","","","","","","","","","",""})
  aAdd(aRegs,{cPerg,"07","Tipo Inicial      ","","","mv_ch7","C",TamSX3("E1_TIPO")[1]    ,0,0,"G","","MV_PAR07","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""})
  aAdd(aRegs,{cPerg,"08","Tipo Final        ","","","mv_ch8","C",TamSX3("E1_TIPO")[1]    ,0,0,"G","","MV_PAR08","","","","ZZZ","","","","","","","","","","","","","","","","","","","","","","","","",""})
  aAdd(aRegs,{cPerg,"09","Cliente Inicial   ","","","mv_ch9","C",TamSX3("A1_COD")[1]     ,0,0,"G","","MV_PAR09","","","","","","","","","","","","","","","","","","","","","","","","","SA1","","","",""})
  aAdd(aRegs,{cPerg,"10","Cliente Final     ","","","mv_cha","C",TamSX3("A1_COD")[1]     ,0,0,"G","","MV_PAR10","","","","","","","","","","","","","","","","","","","","","","","","","SA1","","","",""})
  aAdd(aRegs,{cPerg,"11","Loja Inicial      ","","","mv_chb","C",TamSX3("A1_LOJA")[1]    ,0,0,"G","","MV_PAR11","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""})
  aAdd(aRegs,{cPerg,"12","Loja Final        ","","","mv_chc","C",TamSX3("A1_LOJA")[1]    ,0,0,"G","","MV_PAR12","","","","ZZ","","","","","","","","","","","","","","","","","","","","","","","","",""})
  aAdd(aRegs,{cPerg,"13","Emissao Inicial   ","","","mv_chd","D",08,0,0,"G","","MV_PAR13","","","","01/01/05","","","","",	"","","","","","","","","","","","","","","","","","","","",""})
  aAdd(aRegs,{cPerg,"14","Emissao Final     ","","","mv_che","D",08,0,0,"G","","MV_PAR14","","","","31/12/05","","","","","","","","","","","","","","","","","","","","","","","","",""})
  aAdd(aRegs,{cPerg,"15","Vencimento Inicial","","","mv_chf","D",08,0,0,"G","","MV_PAR15","","","","01/01/05","","","","","","","","","","","","","","","","","","","","","","","","",""})
  aAdd(aRegs,{cPerg,"16","Vencimento Final  ","","","mv_chg","D",08,0,0,"G","","MV_PAR16","","","","31/12/05","","","","","","","","","","","","","","","","","","","","","","","","",""})
  aAdd(aRegs,{cPerg,"17","Natureza Inicial  ","","","mv_chh","C",10,0,0,"G","","MV_PAR17","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""})
  aAdd(aRegs,{cPerg,"18","Natureza Final    ","","","mv_chi","C",10,0,0,"G","","MV_PAR18","","","","ZZZZZZZZZZ","","","","","","","","","","","","","","","","","","","","","","","","",""})
  aAdd(aRegs,{cPerg,"19","Banco Cobranca    ","","","mv_chj","C",TamSX3("A6_COD")[1]     ,0,0,"G","","MV_PAR19","","","","","","","","","","","","","","","","","","","","","","","","","XSEE","","","",""})
  aAdd(aRegs,{cPerg,"20","Agencia Cobranca  ","","","mv_chk","C",TamSX3("A6_AGENCIA")[1] ,0,0,"G","","MV_PAR20","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""})
  aAdd(aRegs,{cPerg,"21","Conta Cobranca    ","","","mv_chl","C",TamSX3("EE_CONTA")[1]   ,0,0,"G","","MV_PAR21","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""})
  aAdd(aRegs,{cPerg,"22","Sub-Conta         ","","","mv_chm","C",TamSX3("EE_SUBCTA")[1]  ,0,0,"G","","MV_PAR22","","","","001","","","","","","","","","","","","","","","","","","","","","","","","",""})
  aAdd(aRegs,{cPerg,"23","Tipo Processo     ","","","mv_chn","C",01,0,0,"C","","MV_PAR23","1- Gerar","1- Gerar","1- Gerar","","","2- Reimpressão","2- Reimpressão","2- Reimpressão","","","3- Regerar","3- Regerar","3- Regerar","","","","","","","","","","","","","","","",""})
  aAdd(aRegs,{cPerg,"24","Diretorio         ","","","mv_cho","C",40,0,0,"G","","MV_PAR24","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""})
  aAdd(aRegs,{cPerg,"25","Gerar Bordero     ","","","mv_chp","C",01,0,0,"C","","MV_PAR25","Sim","Sim","Sim","","","Não","Não","Não","","","","","","","","","","","","","","","","","","","","",""})
  aAdd(aRegs,{cPerg,"26","Tipo Boleto       ","","","mv_chq","C",01,0,0,"C","","MV_PAR26","1- Reduzido","1- Reduzido","1- Reduzido","","","2- Completo","2- Completo","2- Completo","","","","","","","","","","","","","","","","","","","","",""})

  dbSelectArea("SX1")
  SX1->(dbSetOrder(1))

  For nY := 1 To Len(aRegs)
    If ! SX1->(dbSeek(padr(cPerg,10)+aRegs[nY,2]))
		RecLock("SX1",.T.)
		  For nJ := 1 To FCount()
			 If nJ <= Len(aRegs[nY])
				FieldPut(nJ,aRegs[nY,nJ])
			 EndIf
		  Next
		SX1->(MsUnlock())
    EndIf
  Next

  RestArea(aAreaSX1)
  RestArea(aAreaAtu)
Return

/*================================
--  Função: Envio de E-Mail.    --
--                              --
==================================*/
User Function ufnEnvBol(pTitulos)
  Local nId    := 0
  Local nId1   := 0
  Local nPos   := 0
  Local aFiles := Directory(cDirGer + "*.PDF","D")
  Local aNmArq := {}
  Local aDados := {}
  Local cGerPDF := ""
  
  Local oHtml,oProcess
          // Nome do PDF: "Bol_" + Codigo Cliente + Loja Cliente + Banco + Prefixo + Titulo + Parcela
    
  AEVAL(aFiles, {|file| aAdd(aNmArq, file[F_NAME])})

 // Separar os clientes
  For nId := 1 To Len(pTitulos)
      If pTitulos[nId][01] .and. ! Empty(pTitulos[nId][11])
         nPos := aScan(aDados, {|x| x[1] == pTitulos[nId][08]})
         
         If nPos > 0
            Loop
         EndIf    

         aAdd(aDados, {pTitulos[nId][08],;
                       pTitulos[nId][10],;
                       pTitulos[nId][04],;
                       pTitulos[nId][05],;
                       pTitulos[nId][14],;
                       pTitulos[nId][11]})
      EndIf                 
  Next
 // -------------------
  
  aSort(aDados,,,{|x,y| x[08] < y[08]})  
  
  For nId := 1 To Len(aDados)
      cGerPDF := "bol_" + aDados[nId][01]
                     
      nPos := aScan(aNmArq, {|x| Substr(x,1,(TamSX3("A1_COD")[1] + 4)) == Upper(cGerPDF)})
         
      If nPos > 0 
         oProcess := TWFProcess():New("000001","Envio de Boleto")
         oProcess:NewTask("Inicio","\workflow\WFBoleto.htm")

         oHtml  := oProcess:oHtml
         cEmail := aDados[nId][06] 		   

         oHtml:ValByName("cCliente", aDados[nId][02])
         oHtml:ValByName("cNum"    , aDados[nId][03])
         oHtml:ValByName("cParcela", aDados[nId][04])
         oHtml:ValByName("cVencto" , aDados[nId][05])
         oHtml:ValByName("cEmpresa", SM0->M0_NOME)
  
       // Start do WorkFlow
       //_user := Subs(cUsuario,7,15)
         oProcess:ClientName("Administrador")

         For nId1 := 1 To Len(aNmArq)
             If Substr(aNmArq[nId1],1,(TamSX3("A1_COD")[1] + 4)) == Upper(cGerPDF) 
                __CopyFile(cDirGer + aNmArq[nId1], "\Workflow\Boleto_PDF\" + aNmArq[nId1])  // Copiar para pasta de Processado o arquivo já processado

                oProcess:AttachFile("\Workflow\Boleto_PDF\" + aNmArq[nId1])
                fErase(cDirGer + aNmArq[nId1])                               // Deletar da pasta de Recebido o arquivo processado
             EndIf
         Next
             
         oProcess:cTo      := cEmail
         subj              := "Boleto(s)"
         oProcess:cSubject := subj

         oProcess:Start()
  		
         WfSendMail()
      EndIf
  Next          
Return

/*==================================
--  Função: Saldo do boleto.      --
--                                --
====================================*/
USer Function fnSldBol(cPrefixo,cNum,cParcela,cCliente,cLoja)
// Retorna o Saldo de um título
	Local aRet		:= {0,0,0,0}
	Local nVlrAbat	:= 0
	Local nAcresc	:= 0
	Local nDecres	:= 0
	Local nSaldo	:= 0

// Pega os Default dos parâmetros
	cPrefixo	:= Iif(cPrefixo == Nil, SE1->E1_PREFIXO, cPrefixo)
	cNum		:= Iif(cNum == Nil, SE1->E1_NUM, cNum)
	cParcela	:= Iif(cParcela == Nil, SE1->E1_PARCELA, cParcela)
	cCliente	:= Iif(cCliente == Nil, SE1->E1_CLIENTE, cCliente)
	cLoja		:= Iif(cLoja == Nil, SE1->E1_LOJA, cLoja)

// Pega o valor dos abatimentos para o título
	nVlrAbat	:= SomaAbat(cPrefixo,cNum,cParcela,"R",1,,cCliente,cLoja)

// Pega o valor de acréscimos e decrescimos paa o título
	nAcresc := SE1->E1_ACRESC
	nDecres := SE1->E1_DECRESC

// Define o saldo do título
	nSaldo	:= (SE1->E1_SALDO - nVlrAbat - nDecres) + nAcresc

// Monta Vetor com o retorno
	aRet := {nSaldo,nVlrAbat,nAcresc,nDecres}
Return(aRet)

//-------------------------------------------------------------------
/*/{Protheus.doc} R110ALogo
Busca logo de acordo com a filial ou empresa
@author  Samuel Dantas
@since   04/04/2019
@version version
/*/
//-------------------------------------------------------------------
Static Function R110ALogo()

Local cRet := "LGRL"+SM0->M0_CODIGO+SM0->M0_CODFIL+".BMP" // Empresa+Filial

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Se nao encontrar o arquivo com o codigo do grupo de empresas ³
//³ completo, retira os espacos em branco do codigo da empresa   ³
//³ para nova tentativa.                                         ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
If !File( cRet )
	cRet := "LGRL" + AllTrim(SM0->M0_CODIGO) + SM0->M0_CODFIL+".BMP" // Empresa+Filial
EndIf

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Se nao encontrar o arquivo com o codigo da filial completo,  ³
//³ retira os espacos em branco do codigo da filial para nova    ³
//³ tentativa.                                                   ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
If !File( cRet )
	cRet := "LGRL"+SM0->M0_CODIGO + AllTrim(SM0->M0_CODFIL)+".BMP" // Empresa+Filial
EndIf

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Se ainda nao encontrar, retira os espacos em branco do codigo³
//³ da empresa e da filial simultaneamente para nova tentativa.  ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
If !File( cRet )
	cRet := "LGRL" + AllTrim(SM0->M0_CODIGO) + AllTrim(SM0->M0_CODFIL)+".BMP" // Empresa+Filial
EndIf

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Se nao encontrar o arquivo por filial, usa o logo padrao     ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
If !File( cRet )
	cRet := "LGRL"+SM0->M0_CODIGO+".BMP" // Empresa
EndIf

Return cRet

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

//-------------------------------------------------------------------
/*/{Protheus.doc} EnvPCli
description
@author  Samuel Dantas
@since   04/12/2019
@version version
/*/
//-------------------------------------------------------------------
Static Function EnvPCli(cFile)
  Local lRet := .T.
  Local cCC := AllTrim(SuperGetMv("MS_BOLMAIL",.F., "samuel.batista@agisrn.com"))
  Local cTeste := AllTrim(SuperGetMv("MS_TESMAIL",.F., "S"))
  Local cMailTeste := AllTrim(SuperGetMv("MS_MAILTES",.F., "samuel.batista@agisrn.com"))
  
  lRet := U_EnviaEmail("FATURA LAUTO", Alltrim(SA1->A1_XEMCOB), cMsgEmail, cFile, cCC, .T.)
  // lRet := U_EnviaEmail("FATURA LAUTO", Alltrim("testenjbn.oteste.com.br"), cMsgEmail, cFile, '', .T.)

Return lRet

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
/*/{Protheus.doc} getZA4
Busca notas na ZA4
@author  Samuel Dantas
@since   04/02/2020
@version version
/*/
//-------------------------------------------------------------------
Static Function getZA4(_cFilial,cNumCTe, cSerCte, _cCliente, _cLoja)
  Local cQuery
  Local nRECZA4 := 0
  
  cQuery := " SELECT ZA4.R_E_C_N_O_ AS RECZA4 FROM " + RetSqlTab('ZA4')
  cQuery += " INNER JOIN " + RetSqlTab('SF2')+" ON ZA4_CHAVE = F2_CHVNFE "
  cQuery += " WHERE " + RetSqlDel('ZA4')
  cQuery += " AND F2_DOC = '" + cNumCTe + "' AND F2_SERIE = '" + cSerCte + "' "
  cQuery += " AND F2_FILIAL = '" + _cFilial + "' AND F2_CLIENTE = '" + _cCliente + "' "
  cQuery += " AND F2_LOJA = '" + _cLoja + "' "
  cQuery += " AND ZA4_ACAO = 'inclusao' "

  If Select('QRYZA4') > 0
    QRYZA4->(DbCloseArea())
  EndIf

  TCQuery cQuery New Alias 'QRYZA4'

  If QRYZA4->(!Eof())
    nRECZA4 := QRYZA4->RECZA4
  EndIf

Return nRECZA4

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
	
	cSerie := u_getSerie('NUC')
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
