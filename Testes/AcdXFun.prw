#INCLUDE "AcdXFun.ch"
#INCLUDE "protheus.ch"
#INCLUDE "apvt100.ch"

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    � CBGrvEti � Autor � Sandro                � Data � 05/02/01 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Gera ID de etiqueta                                        ���
�������������������������������������������������������������������������Ĵ��
���Parametros� cTipo     = Tipo da etiqueta 01,02,03,04                   ���
���          � aConteudo = Conteudo correspondente ao cTipo obs. abaixo   ���
���          � cID       = Codigo Id                                      ���
�������������������������������������������������������������������������Ĵ��
���Retorno   � cID                                                        ���
�������������������������������������������������������������������������Ĵ��
���Uso       � SIGAACD                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
Tipo '01' produto
conteudo := cod.produto,
qtde,
usuario,
NF entrada,
serie entrada,
Fornecedor,
loja,
Localizacao,
Almoxarifado,
OP,
Num Seq,
NF Saida,
Serie Saida,
Etiqueta do Cliente,
Lote,
SubLote,
Data de validade
Centro de Custo
Tipo '02' localizacao
conteudo := cod. localizacao
Tipo '03' dispositivo de  movimentacao
conteudo := cod. do dispostivo
Tipo '04' usuario
conteudo := cod. do usuario
Tipo '05' volume
conteudo := cod.volume,
num.pedido,
nota,
serie
Tipo '06'
conteudo := cod.transportadora
Tipo '07' volume
conteudo := cod.volume,
NFENT,
SERIEE,
FORNEC,
LOJAFO
*/

Function CBGrvEti(cTipo,aConteudo,cID)
Local lNew
DbSelectArea('CB0')
IF cID == NIL
	While .t.
		cID := Padr(CBProxCod('MV_CODCB0'),10)
		If ! CB0->(DbSeek(xFilial("CB0")+cID))
			exit
		EndIf
	EndDo
	lNew := .t.
	RecLock('CB0',lNew)
	CB0->CB0_DTNASC := dDataBase
Else
	If Len(Alltrim(cID)) <=  TamSx3("CB0_CODETI")[1]   // Codigo Interno
		CB0->(DbSetOrder(1))
		CB0->(DbSeek(xFilial("CB0")+cID))
		lNew := ! CB0->(DbSeek(xFilial("CB0")+Padr(cID,TamSx3("CB0_CODETI")[1])))
	ELSEIf Len(Alltrim(cID)) ==  TamSx3("CB0_CODET2")[1]-1   // Codigo Interno  pelo codigo do cliente
		CB0->(DbSetOrder(2))
		lNew := ! CB0->(DbSeek(xFilial("CB0")+Padr(cID,TamSx3("CB0_CODET2")[1])))
		CB0->(DbSetOrder(1))
	EndIf

	RecLock('CB0',lNew)
	If !lNew .and. cTipo # CB0->CB0_TIPO
		Return NIL
	EndIf
EndIf
CB0->CB0_FILIAL := xFilial("CB0")
If lNew
	CB0->CB0_CODETI := cID
	CB0->CB0_DTNASC := dDataBase
	CB0->CB0_TIPO   := cTipo
Endif


IF cTipo == '01'    // produto
	CB0->CB0_CODPRO := CBChk(aConteudo,1,CB0_CODPRO)
	CB0->CB0_QTDE   := CBChk(aConteudo,2,CB0_QTDE)
	CB0->CB0_USUARI:= CBChk(aConteudo,3,CB0_USUARI)
	CB0->CB0_NFENT  := CBChk(aConteudo,4,CB0_NFENT)
	SerieNfId("CB0",1,'CB0_SERIEE',,,CBChk(aConteudo,5,CB0_SERIEE))
	CB0->CB0_FORNEC := CBChk(aConteudo,6,CB0_FORNEC)
	CB0->CB0_LOJAFO := CBChk(aConteudo,7,CB0_LOJAFO)
	CB0->CB0_PEDCOM := CBChk(aConteudo,8,CB0_PEDCOM)
	CB0->CB0_LOCALI := CBChk(aConteudo,9,CB0_LOCALI)
	CB0->CB0_LOCAL  := CBChk(aConteudo,10,CB0_LOCAL)
	CB0->CB0_OP     := CBChk(aConteudo,11,CB0_OP)
	CB0->CB0_NUMSEQ := CBChk(aConteudo,12,CB0_NUMSEQ)
	CB0->CB0_NFSAI  := CBChk(aConteudo,13,CB0_NFSAI)
	SerieNfId("CB0",1,'CB0_SERIES',,,CBChk(aConteudo,5,CB0_SERIES))
	CB0->CB0_CODET2 := CBChk(aConteudo,15,CB0_CODET2)
	CB0->CB0_LOTE   := CBChk(aConteudo,16,CB0_LOTE)
	CB0->CB0_SLOTE  := CBChk(aConteudo,17,CB0_SLOTE)
	CB0->CB0_DTVLD  := CBChk(aConteudo,18,CB0_DTVLD)
	CB0->CB0_CC     := CBChk(aConteudo,19,CB0_CC)
	CB0->CB0_LOCORI := CBChk(aConteudo,20,CB0_LOCORI)
	CB0->CB0_PALLET := CBChk(aConteudo,21,CB0_PALLET)
	CB0->CB0_OPREQ  := CBChk(aConteudo,22,CB0_OPREQ)
	CB0->CB0_NUMSER := CBChk(aConteudo,23,CB0_NUMSER)
	CB0->CB0_ORIGEM := CBChk(aConteudo,24,CB0_ORIGEM)
	CB0->CB0_ITNFE  := CBChk(aConteudo,25,CB0_ITNFE)

	If Type('cProgImp')=="C" .and. cProgImp=="ACDV130"
		CB0_STATUS := "3" // Disponivel para devolu��o
	EndIf
ElseIf cTipo == '02' // LOCALIZACAO
	CB0->CB0_LOCALI := CBChk(aConteudo,1,CB0_LOCALI)
	CB0->CB0_LOCAL  := CBChk(aConteudo,2,CB0_LOCAL )
ElseIf cTipo == '03' // UNITIZADOR
	CB0->CB0_DISPID := CBChk(aConteudo,1,CB0_DISPID)
ElseIf cTipo == '04' // USUARIO
	CB0->CB0_USUARI:= CBChk(aConteudo,1,CB0_USUARI)
ElseIf cTipo == '05' // VOLUME
	CB0->CB0_VOLUME := CBChk(aConteudo,1,CB0_VOLUME)
	CB0->CB0_PEDVEN := CBChk(aConteudo,2,CB0_PEDVEN)
	CB0->CB0_NFSAI  := CBChk(aConteudo,3,CB0_NFSAI)
	SerieNfId("CB0",1,'CB0_SERIES',,,CBChk(aConteudo,5,CB0_SERIES))
ElseIf cTipo == '06'
	CB0->CB0_TRANSP := CBChk(aConteudo,1,CB0_TRANSP)
ElseIf cTipo == '07' // VOLUME
	CB0->CB0_VOLUME := CBChk(aConteudo,1,CB0_VOLUME)
	CB0->CB0_NFENT  := CBChk(aConteudo,2,CB0_NFENT)
	SerieNfId("CB0",1,'CB0_SERIEE',,,CBChk(aConteudo,3,CB0_SERIEE))
	CB0->CB0_FORNEC := CBChk(aConteudo,4,CB0_FORNEC)
	CB0->CB0_LOJAFO := CBChk(aConteudo,5,CB0_LOJAFO)
EndIf
MsUnLock()
Return cID

Static Function CBChk(aConteudo,nItem,xDef)
local uRet := xDef
If nItem <= len(aConteudo) .and. aConteudo[nItem] <> NIL
	uRet:= aConteudo[nItem]
EndIf
Return uRet

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    � CBRetTipo� Autor � Sandro                � Data � 05/02/01 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Retorna o tipo do ID                                       ���
�������������������������������������������������������������������������Ĵ��
���Parametros� cID    - Numero do ID da Etiqueta                          ���
�������������������������������������������������������������������������Ĵ��
���Retorno   � Tipo do ID                                                 ���
�������������������������������������������������������������������������Ĵ��
���Uso       � SIGAACD                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Function CBRetTipo(cID)
Local cTipo   := ""
Local aArea   := SB1->(GetArea())
Private	cCodAux := SuperGETMV("MV_CODCB0")  // Foi declarado como private, pois podera ser usada no ponto de entrada CBRETTIPO

If	ExistBlock("CBRETTIPO")
	If	ExecBlock("CBRETTIPO",.F.,.F.,{cID})
		Return "EAN8OU13"
	EndIf
EndIf

// Retorna tipo da etiqueta
If	ExistBlock("CBRETIP2")
		cTipo:=ExecBlock("CBRETIP2",.F.,.F.,{cID})
		If Valtype(cTipo) == "C"
		Return cTipo
	EndIf
EndIf

If Len(Alltrim(cID)) == 8 .or. Len(Alltrim(cID)) == 13
	Return "EAN8OU13"
ElseIf Len(Alltrim(cID)) == 14 // verificar o digito
	Return "EAN14"
ElseIf Len(Alltrim(cID)) > TamSx3("B1_COD")[1] .and.  ! UsaCB0('01') //.and. ! Empty(CBAnalisa128(cID))
	Return "EAN128"
Else
	If (UsaCB0('01') .or. UsaCB0('02') .or. UsaCB0('03') .or. UsaCB0('04') .or. UsaCB0('05') .or. UsaCB0('06')) .and. Len(Alltrim(cID)) ==  Len(Alltrim(cCodAux))   // Codigo Interno
		CB0->(DbSetOrder(1))
		If CB0->(DbSeek(xFilial("CB0")+Padr(cID,TamSx3("CB0_CODETI")[1])))
			cTipo := CB0->CB0_TIPO
		EndIf
		Return cTipo
	ELSEIf Len(Alltrim(cID)) ==  TamSx3("CB0_CODET2")[1]-1 .And. (UsaCB0('01') .or. UsaCB0('02') .or. UsaCB0('03') .or. UsaCB0('04') .or. UsaCB0('05') .or. UsaCB0('06'))  // Codigo Interno  pelo codigo do cliente
		CB0->(DbSetOrder(2))
		If CB0->(DbSeek(xFilial("CB0")+Padr(cID,TamSx3("CB0_CODET2")[1])))
			cTipo := CB0->CB0_TIPO
		EndIf
		CB0->(DbSetOrder(1))
		Return cTipo
	EndIf
	SB1->(DbSetOrder(5))
	If SB1->(DbSeek(xFilial('SB1')+Padr(cID,TamSx3("B1_COD")[1])))
		RestArea(aArea)
		Return "EAN8OU13" // O codigo de barras especifico do cliente terah o mesmo comportamento que um codigo EAN8OU13
	EndIf
	//-- Tratamento adicionado para contorno da remo��o do gatilho do B1_COD para B1_CODBAR
	//-- (sem isto, n�o funciona ler o c�digo do produto na etiqueta)
	SB1->(DbSetOrder(1))
	If SB1->(DbSeek(xFilial('SB1')+Padr(cID,TamSx3("B1_COD")[1]))) .And. Empty(SB1->B1_CODBAR)
		RestArea(aArea)
		Return "EAN8OU13" // O codigo de barras especifico do cliente terah o mesmo comportamento que um codigo EAN8OU13
	EndIf	
	//Valida��o especifica para o modulo SIGALOJA
	SLK->(DbSetOrder(1))
	If SLK->(DbSeek(xFilial('SLK')+Padr(cID,TamSx3("LK_CODBAR")[1]))) //Codigo de Barra(LOJA210)
		RestArea(aArea)
		Return "EAN8OU13"
	EndIf
	RestArea(aArea)
EndIf
Return ""

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    � CBRetEti � Autor � Sandro                � Data � 05/02/01 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Retorna conteudo da etiqueta                               ���
�������������������������������������������������������������������������Ĵ��
���Parametros� cID    - Numero do ID da Etiqueta                          ���
���          � cTipId - Tipo do ID                                        ���
�������������������������������������������������������������������������Ĵ��
���Retorno   � aConteudo                                                  ���
�������������������������������������������������������������������������Ĵ��
���Uso       � SIGAACD                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Function CBRetEti( cID, cTipId, lInv, lReq, cOrdSep )
Local aRet   := {}
Local aRetPE := {}
Local lRetPE := .F.
Local lSumQtdEtq:= .F.

DEFAULT lInv := .F.
DEFAULT lReq := .F.
DEFAULT cOrdSep := ''

lSumQtdEtq:= !( Empty( cOrdSep ) )


//-- Ponto de Entrada : Na leitura do codigo da etiqueta, executado antes das validacoes da etiqueta.
If	ExistBlock("CBRETET2")
	aRetPE := ExecBlock("CBRETET2",.F.,.F.,{cID,cTipId})
	If Valtype(aRetPE) == "A" .And. Len(aRetPE) > 0
		Return aRetPE
	EndIf
EndIf

If Len(Alltrim(cID)) <=  TamSx3("CB0_CODETI")[1]   // Codigo Interno
	CB0->(DbSetOrder(1))
	CB0->(DbSeek(xFilial("CB0")+Padr(cID,TamSx3("CB0_CODETI")[1])))
ElseIf Len(Alltrim(cID)) ==  TamSx3("CB0_CODET2")[1]-1   // Codigo Interno  pelo codigo do cliente
	CB0->(DbSetOrder(2))
	CB0->(DbSeek(xFilial("CB0")+Padr(cID,TamSx3("CB0_CODET2")[1])))
	CB0->(DbSetOrder(1))
Else
	Return aRet
EndIf
If CB0->(Eof())
	Return aRet
EndIf
If ! Empty(CB0->CB0_OPREQ+CB0->CB0_CC)  .and. ! lReq // ENCERRADO POR REQUISICAO
	Return aRet
EndIf
If CB0->CB0_STATUS=="2" .and. ! lInv // ENCERRADO POR INVENTARIO
	Return aRet
EndIf

//-- Ponto de Entrada : Na leitura do codigo da etiqueta
If	ExistBlock("CBRETETI")
	aRetPE := ExecBlock("CBRETETI",.F.,.F.,{cID,cTipId})
	If Valtype(aRetPE) == "A" .And. Len(aRetPE) > 0
		Return aRetPE
	EndIf
EndIf
//-- Ponto de Entrada para validar a etiqueta
If	ExistBlock("CBVRETIQ")
	lRetPE := ExecBlock("CBVRETIQ")
	lRetPE := If(Valtype(lRetPE)=="L",lRetPE,.T.)
	If ! lRetPE
		Return aRet
	EndIf
EndIf

If cTipId <> NIL .and. CB0->CB0_TIPO # cTipId
	Return aRet
EndIf
If CB0->CB0_TIPO == '01'    // produto
	aadd(aRet,CB0->CB0_CODPRO)    //1-codigo do produto
	If !( lSumQtdEtq )
		aadd(aRet,CB0->CB0_QTDE)      //2-quantidade
	Else
		aadd(aRet, SumQtdEti( cOrdSep, CB0->CB0_CODPRO ) ) //2-quantidade
	EndIf	
	aadd(aRet,CB0->CB0_USUARI)   //3-codigo do usuario
	aadd(aRet,CB0->CB0_NFENT)     //4-nota fiscal de entrada
	aadd(aRet,CB0->CB0_SERIEE)    //5-serie da NF de entrada
	aadd(aRet,CB0->CB0_FORNEC)    //6-codigo do fornecedor
	aadd(aRet,CB0->CB0_LOJAFO)    //7-Loja do fornecedor
	aadd(aRet,CB0->CB0_PEDCOM)    //8-Pedido de compra
	aadd(aRet,CB0->CB0_LOCALI)    //9-Localizacao
	aadd(aRet,CB0->CB0_LOCAL)     //10-Almoxarifado
	aadd(aRet,CB0->CB0_OP)        //11-OP
	aadd(aRet,CB0->CB0_NUMSEQ)    //12-Numero de Sequencia
	aadd(aRet,CB0->CB0_NFSAI)     //13-nota fiscal de SAIDA
	aadd(aRet,CB0->CB0_SERIES)    //14-serie da NF de SAIDA
	aadd(aRet,CB0->CB0_CODET2)    //15-codigo da etiqueta do cliente
	aadd(aRet,CB0->CB0_LOTE)      //16-Lote
	aadd(aRet,CB0->CB0_SLOTE)     //17-Sub-Lote
	aadd(aRet,CB0->CB0_DTVLD)     //18-data de validade
	aadd(aRet,CB0->CB0_CC)        //19-centro de custo
	aadd(aRet,CB0->CB0_LOCORI)    //20-Armazem Original
	aadd(aRet,CB0->CB0_PALLET)    //21-Codigo do Pallet
	aadd(aRet,CB0->CB0_OPREQ)     //22-OP para qual o produto foi requisitado
	aadd(aRet,CB0->CB0_NUMSER)    //23-Numero de Serie
	aadd(aRet,CB0->CB0_ORIGEM)    //24-Origem
	aadd(aRet,CB0->CB0_ITNFE) 	  //25-Item Nota Fiscal de entrada

ElseIf CB0->CB0_TIPO == '02' // LOCALIZACAO
	aRet := {CB0->CB0_LOCALI,CB0->CB0_LOCAL}
ElseIf CB0->CB0_TIPO == '03' // UNITIZADOR
	aRet := {CB0->CB0_DISPID}
ElseIf CB0->CB0_TIPO == '04' // USUARIO
	aRet := {CB0->CB0_USUARI}
ElseIf CB0->CB0_TIPO == '05' //VOLUME
	aRet := {CB0->CB0_VOLUME,;
		CB0->CB0_PEDVEN,;
		CB0->CB0_NFSAI,;
		CB0->CB0_SERIES}
ElseIf CB0->CB0_TIPO == '06'
	aRet := {CB0->CB0_TRANSP}
ElseIf CB0->CB0_TIPO == '07' //VOLUME
	aRet := {CB0->CB0_VOLUME,;
		CB0->CB0_NFENT,;     //3-nota fiscal de entrada
		CB0->CB0_SERIEE,;    //4-serie da NF de entrada
		CB0->CB0_FORNEC,;    //5-codigo do fornecedor
		CB0->CB0_LOJAFO}     //6-Loja do fornecedor
EndIf
Return aRet
/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    � CBQEmb   � Autor � Henrique              � Data � 01/06/09 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Responsavel pelo retorno da quantidade por embalagem na    ���
���			 � leitura da etiquta										  ���
�������������������������������������������������������������������������Ĵ��
���Retorno   � nQE                                                        ���
�������������������������������������������������������������������������Ĵ��
���Uso       � SIGAACD                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Function CBQEmb()
SB5->(DbSetOrder(1))
If ! SB5->(DbSeek(xFilial("SB5")+SB1->B1_COD)) .or. Empty(SB5->B5_QEL)
	If RetArqProd(SB1->B1_COD)
		nQE := SB1->B1_QE
	Else
		nQE := SBZ->BZ_QE
	EndIf
Else
	nQE :=SB5->B5_QEL
EndIf

If Empty(nQE)
	nQE := 1
EndIf
Return nQE
/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    � CBQEmbI  � Autor � Henrique              � Data � 01/06/09 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Responsavel pelo retorno da quantidade por embalagem na    ���
���			 � impress�o da etiquta										  ���
�������������������������������������������������������������������������Ĵ��
���Retorno   � nQE                                                        ���
�������������������������������������������������������������������������Ĵ��
���Uso       � SIGAACD                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Function CBQEmbI()
SB5->(DbSetOrder(1))
If ! SB5->(DbSeek(xFilial("SB5")+SB1->B1_COD)) .or. Empty(SB5->B5_QEI)
	If RetArqProd(SB1->B1_COD)
		nQE := SB1->B1_QE
	Else
		nQE := SBZ->BZ_QE
	EndIf
Else
	nQE :=SB5->B5_QEI
EndIf

If Empty(nQE)
	nQE := 1
EndIf
Return nQE

/*
������������������������������������������������������������������������������
������������������������������������������������������������������������������
��������������������������������������������������������������������������Ŀ��
���Fun��o    � CBGrvQEmbI� Autor � Henrique              � Data � 01/06/09 ���
��������������������������������������������������������������������������Ĵ��
���Descri��o � Grava o ajuste da quantidade por embalagem na impressao da  ���
���			 � da etiquta	        									   ���
��������������������������������������������������������������������������Ĵ��
���Retorno   �                                                             ���
��������������������������������������������������������������������������Ĵ��
���Uso       � SIGAACD                                                     ���
���������������������������������������������������������������������������ٱ�
������������������������������������������������������������������������������
������������������������������������������������������������������������������
*/
Function CBGrvQEmbI(nQtde)
Local aAreaAnt
SB5->(DbSetOrder(1))
If ! SB5->(DbSeek(xFilial("SB5")+SB1->B1_COD)) .or. Empty(SB5->B5_QEI)
	If RetArqProd(SB1->B1_COD)
		SB1->(RecLock("SB1",.F.))
		SB1->B1_QE:=nQtde
		SB1->(MsUnLock())
	Else
		aAreaAnt := GetArea()
		SBZ->(RecLock("SBZ",.F.))
		SBZ->BZ_QE:=nQtde
		SBZ->(MsUnLock())
		RestArea(aAreaAnt)
	EndIf
Else
	SB5->(RecLock("SB5",.F.))
	SB5->B5_QEI:=nQtde
	SB5->(MsUnLock())
EndIf
Return


/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    � CBRetEtiEan� Autor � Sandro              � Data � 05/02/01 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Retorna array com o codigo do produto e qtde               ���
�������������������������������������������������������������������������Ĵ��
���Parametros� cID    - Numero do ID da Etiqueta                          ���
�������������������������������������������������������������������������Ĵ��
���Retorno   � array                                                      ���
�������������������������������������������������������������������������Ĵ��
���Uso       � SIGAACD                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Function CBRetEtiEAN(cID)
Local aRet      := {}
Local lEAN13OU8 := .F.
Local lEAN14    := .F.
Local lEAN128   := .F.
Local cCodBar   := ''
Local cDL       := ''
Local nQE       := 0
Local nPos      := 0
Local nOrdemSB1 := 0
Local aEan128   := {}
Local lEAN12813OU8 := .F.
Local lEAN12814    := .F.
Local lEAN12814VAR := .f.
Local cUnDespacho  := ''
Local nQtdeDespacho:= 0
Local cLote     := ''
Local dValid    := ctod('')
Local cNumSerie := Space(20)
Local uAux

If	ExistBlock("CBRETEAN")
	// Retorno devera ser um array conforme abaixo:
	// {codigo do produto,quantidade,lote,data de validade, numero de serie}
	aRet := ExecBlock("CBRETEAN",,,{cID})
	If	Len(aRet) > 0
		Return aRet
	EndIf
EndIf

//Se n�o � EAN 128, primeiro verifica se o usu�rio bipou o c�digo do produto
If Len(Alltrim(cID)) <= TamSX3("B1_COD")[1]
	SB1->(DbSetOrder(1))
	If SB1->(DbSeek(xFilial('SB1')+cID))
		SB5->(DbSetOrder(1))
		SB5->(DbSeek(xFilial("SB5")+SB1->B1_COD))
		If SB5->B5_TIPUNIT <> '0' //produtos com controle unit�rio
			nQE   := CBQEmb()
		Else
			nQE   := 1
		EndIf
		//Se encontrar pelo c�digo do produto, retorna direto
		Return {SB1->B1_COD,nQE,Padr(cLote,TamSX3("CB8_LOTECT")[1]),dValid,Padr(cNumSerie,20)}
	EndIf
EndIf

If Len(Alltrim(cID)) == 8  .or. Len(Alltrim(cID)) == 13
	cCodBar := Alltrim(cID)
	lEAN13OU8 :=.T.
ElseIf Len(Alltrim(cID)) == 14
	cCodBar := Subs(Alltrim(cID),2,12)
	cUnDespacho := Left(cID,1) //-- checar o digito
	If Left(cCodBar,5) =="00000"
		cCodBar := Subs(cCodBar,6)
	EndIf
	lEAN14 := .T.
ElseIf Len(Alltrim(cID)) > TamSX3("B1_COD")[1] .and. ! UsaCB0('01')
	aEan128 := CBAnalisa128(cID)
	If ! Empty(aEan128)
		lEAN128 := .T.
		nPos := Ascan(aEan128,{|x| x[1] == "01"})
		If nPos > 0
			cCodBar:= Subst(aEan128[nPos,2],2,12)
			cDL := Left(aEan128[nPos,2],1)
		EndIf
		nPos := Ascan(aEan128,{|x| x[1] == "02"})
		If nPos > 0
			cCodBar:= Subst(aEan128[nPos,2],2,12)
			cDL := Left(aEan128[nPos,2],1)
		EndIf
		nPos := Ascan(aEan128,{|x| x[1] == "8006"})
		If nPos > 0
			cCodBar:= Subst(aEan128[nPos,2],2,12)
			cDL := Left(aEan128[nPos,2],1)
		EndIf
		If cDL $ "12345678"
			cUnDespacho := cDL
			lEAN12814 := .T.
		ElseIf cDL =="0"
			lEAN12813OU8 := .T.
		ElseIf cDL =="9"
			lEAN12814VAR := .T.
		EndIf
		If Left(cCodBar,5) =="00000"
			cCodBar := Subs(cCodBar,6)
		EndIf
	EndIf
Else
	cCodBar := Alltrim(cID)
	lEAN13ou8 := .T.
EndIf
If ! lEAN13ou8 .And. ! lEAN14 .and. !lEAN128 .or. Empty(cCodBar)
	Return {}
EndIf
nOrdemSB1:= SB1->(IndexOrd())
SB1->(DbSetOrder(5))
SB1->(DBSeek(xFilial("SB1")+cCodBar))
SB1->(DbSetOrder(nOrdemSB1))
If SB1->(Eof())
	dbSelectArea("SLK")
	SLK->( dbSetOrder(1) )
	If SLK->( DBSeek(xFilial("SLK")+cCodBar) )
		aRet := {LK_CODIGO, LK_QUANT,Padr(cLote,TamSX3("CB8_LOTECT")[1]),dValid,Padr(cNumSerie,20)}
		Return aRet
	Else
		Return aRet
	EndIf
EndIf
SB5->(DbSetOrder(1))
SB5->(DBSeek(xFilial("SB5")+SB1->B1_COD))
 If lEAN13ou8
	If SB5->B5_TIPUNIT <> '0' //produtos com controle unitario
		nQE   := CBQEmb()
	Else
		nQE   := 1
	EndIf
ElseIf lEAN14
	nQtdeDespacho := SB5->(FieldGet(FieldPos("B5_EAN14"+cUnDespacho)))
	nQE := nQtdeDespacho
ElseIf lEAN128
	nPos := Ascan(aEan128,{|x| x[1] == "30"})  // Qtde variavel
	If nPos > 0
		nQtdeDespacho:= Val(aEan128[nPos,2])
	EndIf
	nPos := Ascan(aEan128,{|x| x[1] == "37"}) // Qtde de itens comerciais
	If nPos > 0
		nQE:= Val(aEan128[nPos,2])
		If lEAN12814
			nQE:= nQE*SB5->(FieldGet(FieldPos("B5_EAN14"+cUnDespacho)))
		ElseIf lEAN12814VAR
			If ! Empty(nQtdeDespacho)
				nQE:= nQE*nQtdeDespacho
			EndIf
		EndIf
	Else
		nQE := nQtdeDespacho
	EndIf
	nPos := Ascan(aEan128,{|x| x[1] == "10"})  // lote
	If nPos > 0
		cLote := aEan128[nPos,2]
	EndIf
	nPos := Ascan(aEan128,{|x| x[1] == "15"})  // data de durabilidade
	If nPos > 0
		uAux:= right(aEan128[nPos,2],2)+'/'+Subs(aEan128[nPos,2],3,2)+'/'+right(aEan128[nPos,2],2)
		If Left(uAux,2) =="00"
			uAux := "01"+Subs(uAux,3)
			dValid := ctod(StrZero(LastDay(ctod(uAux)),2)+Subs(uAux,3))
		Else
			dValid := ctod(uAux)
		EndIf
	EndIf
	nPos := Ascan(aEan128,{|x| x[1] == "17"})  // data de validade
	If nPos > 0
		uAux:= right(aEan128[nPos,2],2)+'/'+Subs(aEan128[nPos,2],3,2)+'/'+right(aEan128[nPos,2],2)
		If Left(uAux,2) =="00"
			uAux := "01"+Subs(uAux,3)
			dValid := ctod(StrZero(LastDay(ctod(uAux)),2)+Subs(uAux,3))
		Else
			dValid := ctod(uAux)
		EndIf
	EndIf
	nPos := Ascan(aEan128,{|x| x[1] == "21"})  // numero de serie
	If nPos > 0
		cNumSerie := aEan128[nPos,2]
	EndIf
EndIf
aRet := {SB1->B1_COD,nQE,Padr(cLote,TamSX3("CB8_LOTECT")[1]),dValid,Padr(cNumSerie,20)}
Return aRet

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    �CBAnalisa128� Autor � Sandro              � Data � 05/02/01 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Analisa string seperando os AIs                            ���
�������������������������������������������������������������������������Ĵ��
���Parametros� cID    - Numero do ID da Etiqueta                          ���
�������������������������������������������������������������������������Ĵ��
���Retorno   � array                                                      ���
�������������������������������������������������������������������������Ĵ��
���Uso       � SIGAACD                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Function CBAnalisa128(cID)
Local aAI
Local aResultado:={}
Local nX
Local nPos,nPos2
Local lVolta := .f.
Local lErro := .t.
aAI :=MsCbTabEan()
lVolta := .f.
While len(cId) >0
	lErro := .t.
	For nX:= 2 to 4
		If left(cId,1) == chr(29)
			cId := Subs(cId,2)
		EndIf

		npos:= ascan(aAI,{|x| Alltrim(x[1]) == Left(cID,nX)})
		If nPos > 0
			lErro := .f.
			If aAI[nPos,4] // variavel
				nPos2:= at(chr(29),cID)
				nPos2:= If(nPos2==0,len(cID)+1,nPos2)
				nPos2:= nPos2-nX
				aadd(aResultado,{aAI[nPos,1],Subs(cID,nX+1,nPos2-1),aAI[nPos,5]})
				//conout(aResultado[len(aResultado),2])
			Else
				nPos2:=aAI[nPos,3]
				aadd(aResultado,{aAI[nPos,1],Subs(cID,nX+1,nPos2),aAI[nPos,5]})
				//conout(aResultado[len(aResultado),2])
			EndIf
			cId := Subs(cId,nX+nPos2+1)
			If left(cId,1) == chr(29)
				cId := Subs(cId,2)
			EndIf
			lVolta:= .t.
			exit
		EndIf
	Next
	If lErro
		aResultado:={}
		Exit
	EndIf
	If ! lVolta
		exit
	endIf
End
Return aResultado
/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    � CBDigVer   � Autor � Sandro              � Data � 05/02/01 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � calcula digito verificador para ean13,ean8 e ean14         ���
�������������������������������������������������������������������������Ĵ��
���Parametros� cCodigo    - Numero do ID da Etiqueta                      ���
�������������������������������������������������������������������������Ĵ��
���Retorno   � digito                                                     ���
�������������������������������������������������������������������������Ĵ��
���Uso       � SIGAACD                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Function CBDigVer(cCodigo)
Local nPeso:= 3
Local nSoma:= 0
Local nx
For nx:= len(cCodigo) to 1 step -1
	nSoma += Val(Subs(cCodigo,nx,1))*nPeso
	nPeso :=If(nPeso== 3,1,3)
Next
nSoma := Val(Right(Alltrim(Str(nSoma,5)),1))
nDig  := 10-nSoma
If nDig== 10
	nDig:= 0
EndIf
Return Str(nDig,1)

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    � CBLoad128  � Autor � Sandro              � Data � 05/02/01 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Entra com a segunda parte do codigo 128                    ���
�������������������������������������������������������������������������Ĵ��
���Parametros� cID    - Numero do ID da Etiqueta                          ���
�������������������������������������������������������������������������Ĵ��
���Retorno   � array                                                      ���
�������������������������������������������������������������������������Ĵ��
���Uso       � SIGAACD                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/

Function CBLoad128(cID)
Local aSave := VTSave()
Local cID2 := space(64)
Local lRet:= .f.
Local lTem2Parte:= SuperGetmv("MV_EANP2",.F.,.F.)

IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf

If ! lTem2Parte
	Return .T.
EndIf
VTClear()

If lVT100B // GetMv("MV_RF4X20")
	 @ 0,0 VTSay STR0001 //"Leitura da segunda"
	 @ 1,0 VTSay STR0002 //"parte do codigo"
	 @ 2,0 VTGet cID2 pict "@!" //valid If(cID==cID2,(VTAlert(STR0003,STR0004,.t.,3000),.f.),.t.) //"Leitura invalida"###"Atencao"
ElseIf VTModelo()=="RF"
	 @ 2,0 VTSay STR0001 //"Leitura da segunda"
	 @ 3,0 VTSay STR0002 //"parte do codigo"
	 @ 4,0 VTGet cID2 pict "@!" //valid If(cID==cID2,(VTAlert(STR0003,STR0004,.t.,3000),.f.),.t.) //"Leitura invalida"###"Atencao"
Else
	 @ 0,0 VTSay STR0005 //"Complemento codigo"
	 @ 1,0 VTGet cID2 pict "@!" //valid If(cID==cID2,(VTAlert(STR0003,STR0004,.t.,3000),.f.),.t.) //"Leitura invalida"###"Atencao"
EndIf

VTRead
If VTLastkey() # 27
	cID := Alltrim(cID)+Alltrim(cID2)
	lRet:= .t.
EndIf
VTRestore(,,,,aSave)
vtclearBuffer()
Return lRet
/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    � CBEndLib � Autor � Sandro                � Data � 05/02/01 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Retorna se o endereco esta liberado para movimentacao      ���
�������������������������������������������������������������������������Ĵ��
���Parametros� cArmazem e cEndereco                                       ���
�������������������������������������������������������������������������Ĵ��
���Retorno   � Logico                                                     ���
�������������������������������������������������������������������������Ĵ��
���Uso       � SIGAACD                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Function CBEndLib(cArmazem,cEndereco)
Local aArea   := SBE->(GetArea())
Local dDtInv  := Posicione("SBE",1,xFilial("SBE")+cArmazem+cEndereco,"BE_DTINV")
Local lVlDtInv:= If (GetMv("MV_VLDTINV") == "1",.T.,.F.)
RestArea(aArea)
Return If(lVlDtInv,dDtInv # dDatabase,Empty(dDtInv))

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    � CBProxCod� Autor � Sandro                � Data � 05/02/01 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Retorna proximo ID                                         ���
�������������������������������������������������������������������������Ĵ��
���Parametros� cCodPar  - codigo do parametro                             ���
�������������������������������������������������������������������������Ĵ��
���Retorno   � ID                                                         ���
�������������������������������������������������������������������������Ĵ��
���Uso       � SIGAACD                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Function CBProxCod(cCodPar)
Local nX
Local cCar
Local lRes
Local cCodAnt
Local nC:=0
Local cNomeArq := "proxcod.sem"
Local oFWriter := FwFileWriter():New( cNomeArq ,.t.)

//-- Valida/cria sem�foro para evitar que duas threads peguem o mesmo valor
While !oFWriter:Create()
	SLeep(50)
	nC++
	If nC == 60
		nC := 0
		conout(STR0006) //'Semaforo fechado '
	EndIf
End

cCodAnt := GETMV( cCodPar )
oFWriter:Write( cCodAnt )

If cCodPar =="MV_CODCB0" .and. GetNewPar("MV_CBNEWID","0") =="0"
	For nx:= Len(cCodAnt) to 1 step -1
		cCar := Subs(cCodAnt,nX,1)
		lRes:= SomaCar(@cCar)
		cCodAnt := Stuff(cCodAnt,nX,1,cCar)
		If lRes
			exit
		EndIf
	Next nx
Else
	If GetNewPar("MV_CB0ALFA","1") =="0" .and. cCodAnt == Repl("9",len(cCodAnt))
		cCodAnt := Repl("0",len(cCodAnt))
	Else
		cCodAnt := Soma1(cCodAnt,Len(cCodAnt))
	EndIf
EndIF

PutMv(cCodPar,cCodAnt)
oFWriter:Close()
oFWriter:Erase()
Return cCodAnt

Static Function SomaCar(cCar)
Local cAnt:= ':['
Local cNew:= 'A0'
cCar := Chr(Asc(cCar)+1)
If cCar $ cAnt
	cCar := Subs(cNew,At(cCar,cAnt),1)
EndIF
Return (cCar # '0')



/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    �CB5SetImp � Autor � Sandro                � Data � 05/02/01 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � configura o MSCBPRINTER conforme a tabela CB5              ���
�������������������������������������������������������������������������Ĵ��
���Parametros� cCod - codigo do CB5                                       ���
���          � lVerServer - .t. aceita impressao somente no server        ���
���          � nDensidade -  densidade da impressao                       ���
���          � nTam  - Tamanho da etiqueta                                ���
���          � cPorta - porta de impresao                                 ���
�������������������������������������������������������������������������Ĵ��
���Retorno   � ID                                                         ���
�������������������������������������������������������������������������Ĵ��
���Uso       � SIGAACD                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
lVerServer // parametro nao mais utilizado
*/
Function CB5SetImp(cCod,lVerServer,nDensidade,nTam,cPorta)
Local cModelo,lTipo,nPortIP,cServer,cEnv,cFila,lDrvWin

If Empty(cCod)
	Return .f.
EndIf
If ! CB5->(DbSeek(xFilial("CB5")+cCod))
	Return .f.
EndIf
cModelo :=Trim(CB5->CB5_MODELO)
If cPorta ==NIL
	If CB5->CB5_TIPO == '4'
		cPorta:= "IP"
	Else
		IF CB5->CB5_PORTA $ "12345"
			cPorta  :='COM'+CB5->CB5_PORTA+':'+CB5->CB5_SETSER
		EndIf
		IF CB5->CB5_LPT $ "12345"
			cPorta  :='LPT'+CB5->CB5_LPT+':'
		EndIf
	EndIf
EndIf

lTipo   :=CB5->CB5_TIPO $ '12'
nPortIP :=Val(CB5->CB5_PORTIP)
cServer :=Trim(CB5->CB5_SERVER)
cEnv    :=Trim(CB5->CB5_ENV)
cFila   := NIL
If CB5->CB5_TIPO=="3"
	cFila := Alltrim(Tabela("J3",CB5->CB5_FILA,.F.))
EndIf
nBuffer := CB5->CB5_BUFFER
lDrvWin := (CB5->CB5_DRVWIN =="1")
MSCBPRINTER(cModelo,cPorta,nDensidade,nTam,lTipo,nPortIP,cServer,cEnv,nBuffer,cFila,lDrvWin,Trim(CB5->CB5_PATH))
MSCBCHKSTATUS(CB5->CB5_VERSTA =="1")

Return .t.


/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    � UsaCB0   � Autor � Eduardo Motta         � Data � 20/11/01 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Retorna se CB0 e' usado ou nao para determinado campo      ���
�������������������������������������������������������������������������Ĵ��
���Parametros� cTipo     = Tipo da etiqueta 01,02,03,04                   ���
�������������������������������������������������������������������������Ĵ��
���Retorno   � lRet                                                       ���
�������������������������������������������������������������������������Ĵ��
���Uso       � SIGAACD                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
Tipo '01' produto
Tipo '02' localizacao/Endereco
Tipo '03' dispositivo de  movimentacao
Tipo '04' usuario
Tipo '05' volume saida
Tipo '06' transportadora
Tipo '07' volume entrada
*/
Function UsaCB0(cTipo)
Return (cTipo $ SuperGetMv("MV_ACDCB0",.F.," "))

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    � CBQtdEmb � Autor � Sandro                � Data � 20/11/01 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Retorna a qtde de embalagem                                ���
�������������������������������������������������������������������������Ĵ��
���Parametros� cProduto  = codigo do produto                              ���
�������������������������������������������������������������������������Ĵ��
���Retorno   � nQE                                                        ���
�������������������������������������������������������������������������Ĵ��
���Uso       � SIGAACD                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/

Function CBQtdEmb(cProduto,nQE)
Local aSave
Local cPictQtde := CBPictQtde()
Local nQtdPE    := 0
DEFAULT nQE := 1                           
  

IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf

If	ExistBlock("CBRQEESP")
	nQtdPE := ExecBlock("CBRQEESP",.f.,.f.,cProduto)
	nQE := If(ValType(nQtdPE)=="N",nQtdPE,nQE)
	Return nQE
EndIf
SB1->(dbSetOrder(1))
SB1->(dbSeek(xFilial("SB1")+cProduto))
SB5->(dbSetOrder(1))
If ! SB5->(DBSeek(xFilial("SB5")+SB1->B1_COD)) .or. SB5->B5_TIPUNIT <> '0' //produtos com controle unitario
	nQE := CBQEmb()
Else  							 //produtos com a necessidade de ser embalado

	if IsTelNet()                  
		aSave := VTSAVE()
		VTClear
		If lVT100B // GetMv("MV_RF4X20")
			@ 0,0 VTSay STR0007 //"Produto a granel!"
			@ 2,0 VtSay STR0008  //"Quantidade"
			@ 3,0 VtGet nQE pict cPictQtde valid nQE > 0
		Elseif VTModelo()=="RF"
			@ 1,0 VTSay STR0007 //"Produto a granel!"
			@ 3,0 VtSay STR0008  //"Quantidade"
			@ 4,0 VtGet nQE pict cPictQtde valid nQE > 0
		Else         //  12345678901234567890
			@ 0,00 VTSay STR0007 //"Produto a granel!"
			@ 1,00 VtSay STR0009 VtGet nQE pict cPictQtde valid nQE > 0 //"Qtd:"
		EndIf
		VTREAD
		VtRestore(,,,,aSave)
		If	VTLastKey() == 27
			VTAlert(STR0010,STR0011,.t.,3000) //"Quantidade Invalida"###"Aviso"
			nQE := 0
		EndIf
	Else

	EndIf
EndIf
Return nQE

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    �CBProdUnit� Autor � Sandro                � Data � 20/11/01 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Verfica se o produto e unitario ou granel                  ���
�������������������������������������������������������������������������Ĵ��
���Parametros� cProduto  = codigo do produto                              ���
�������������������������������������������������������������������������Ĵ��
���Retorno   � nQE                                                        ���
�������������������������������������������������������������������������Ĵ��
���Uso       � SIGAACD                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Function CBProdUnit(cProduto)
Local aArea:= SB5->(GetArea())
Local lUnit
SB5->(DbSetOrder(1))
If ! SB5->(DbSeek(xFilial("SB5")+cProduto))
	lUnit := .t.
Else
	lUnit := SB5->B5_TIPUNIT <> '0'
EndIf
RestArea(aArea)
Return lUnit
/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    �CBPictQtde� Autor � Sandro                � Data � 20/11/01 ���
�������������������������������������������������������������������������Ĵ��
���Retorno   � Retorna a picture para a quantidade                        ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Function CBPictQtde()
Static cPicture := NIL
If cPicture==NIL
	cPicture := PesqPict("SB2","B2_QATU")
EndIf
Return cPicture
/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    �CBRetOpe  � Autor � Sandro                � Data � 20/11/01 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Atualiza o codigo do operador                              ���
�������������������������������������������������������������������������Ĵ��
���Parametros� cCodUsr,cCodOpe                                            ���
�������������������������������������������������������������������������Ĵ��
���Retorno   � Logico                                                     ���
�������������������������������������������������������������������������Ĵ��
���Uso       � SIGAACD                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Function CBRetOpe()
Local cCodUsr:=__cUserID
Local cCodOpe:=""
Local aArea:=GetArea()
Local aSX3 := SX3->(GetArea())
Local cQuery := ""
Local nRecnoCB1 := NIL
SX3->(DbSetOrder(1))
If ! SX3->(DbSeek('CB1'))
	RestArea(aSX3)
	Return cCodOpe
EndIf
RestArea(aSX3)
CB1->(DbSetOrder(2))

cQuery := "SELECT CB1_CODOPE, R_E_C_N_O_ RECCB1 FROM "+RetSqlName("CB1")
cQuery += " WHERE CB1_FILIAL = '"+xFilial("CB1")+"' AND "
cQuery += " CB1_CODUSR = '"+cCodUsr+"' AND "
cQuery += " CB1_STATUS = '1' AND "
cQuery += " D_E_L_E_T_ = ' '"
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"CB1TOP")

If CB1TOP->(!EOF())
	nRecnoCB1	:= CB1TOP->RECCB1
	cCodOpe 	:= CB1TOP->CB1_CODOPE
EndIf

CB1TOP->(DbCloseArea())

//Posiciona no operador caso exista.
If nRecnoCB1#NIL
	CB1->(DbGoto(nRecnoCB1))
EndIf

CB1->(DbSetOrder(1))
RestArea(aArea)
Return cCodOpe
/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    �CBVldOpe  � Autor � Sandro                � Data � 20/11/01 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Validacao do operador do ACDSTD                            ���
�������������������������������������������������������������������������Ĵ��
���Retorno   � Logico                                                     ���
�������������������������������������������������������������������������Ĵ��
���Uso       � SIGAACD Generico                                           ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Function CBVldOpe(cCodOpe)
Local lRet := .T.
Local cQuery    := ""
Local nRecnoCB1 := NIL

CB1->(DbSetOrder(1))

	cQuery := "SELECT R_E_C_N_O_ RECCB1 FROM "+RetSqlName("CB1")
	cQuery += " WHERE CB1_FILIAL = '"+xFilial("CB1")+"' AND "
	cQuery += " CB1_CODOPE = '"+cCodOpe+"' AND "
	cQuery += " CB1_STATUS = '1' AND "
	cQuery += " D_E_L_E_T_ = ' '"
	cQuery := ChangeQuery(cQuery)
	dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"CB1TOP")

	If CB1TOP->(!EOF())
		nRecnoCB1 := CB1TOP->RECCB1
		Conout(nRecnoCB1)
	EndIf

	CB1TOP->(DbCloseArea())

	//Posiciona no operador caso exista.
	If nRecnoCB1#NIL
		CB1->(DbGoto(nRecnoCB1))
	EndIf

	If nRecnoCB1 == NIL .Or. CB1->(cCodOpe <> CB1_CODOPE)
		lRet := .f.
	EndIf

If !lRet
	CBAlert(STR0012,STR0013,.T.,3000,2) //'Operador invalido'###'Aviso'
EndIf

Return lRet



/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    � CBRastro � Autor � Sandro                � Data � 20/11/01 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Verifica se o produto tem rastro                           ���
�������������������������������������������������������������������������Ĵ��
���Retorno   � Logico - .t. / .f.                                         ���
�������������������������������������������������������������������������Ĵ��
���Uso       � SIGAACD                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/

Function CBRastro(cProduto,cLote,cSubLote,dValid,lEmpty,lNextLote)
Local aSave
Local aRetAux		:= {}
Local lAltera		:= .T.
IF !Type("lVT100B") == "L"
	Private lVT100B := .F.
EndIf
DEFAULT lEmpty		:= .F.
DEFAULT lNextLote 	:= .F.


If ! Rastro(cProduto)
	Return .t.
EndIf

If	ExistBlock("CBRastro")
	// o retorno devera ser um array conforme abaixo:
	//{lote,sub-lote,data de validade}
	aRetAux := ExecBlock("CBRastro",.F.,.F.,{cProduto,cLote,cSubLote,dValid})
	If	Valtype(aRetAux)=='A' .And. Len(aRetAux) >= 3
		cLote    := If(Valtype(aRetAux[1])=='C',aRetAux[1],cLote)
		cSubLote := If(Valtype(aRetAux[2])=='C',aRetAux[2],cSubLote)
		dValid   := If(Valtype(aRetAux[3])=='D',aRetAux[3],dValid)
		If Len(aRetAux) >= 4
			lAltera  := If(Valtype(aRetAux[4])=='L',aRetAux[4],lAltera)
		EndIf
	EndIf
EndIf

aSave   := VTSAVE()

If lAltera
	If lVT100B // GetMv("MV_RF4X20")
		VTClear
		@ 0,0 VtSay STR0015  //"Lote"
		@ 1,0 VtGet cLote /*pict '@!'*/  valid If(lEmpty,.t.,!Empty(cLote).Or. lNextLote) when Empty(cLote)
		If Rastro(cProduto,"S")	.and. cSubLote <> NIL
			@ 2,0 VtSay STR0016  //"Sub-Lote"
			@ 3,0 VtGet cSubLote pict '@!' valid If(lEmpty,.t.,!Empty(cSubLote))
		EndIf
		VTRead
		VTClear()
		If dValid <> NIL
			@ 0,0 VtSay STR0017  //"Validade"
			@ 1,0 VtGet dValid pict '@D' valid If(lEmpty,.t.,!Empty(dValid))
			VTREAD
		EndIf
		
	ElseIf VTModelo()=="RF"
		VTClear
		@ 0,0 VTSay STR0014 //"Produto com rastro"
		@ 2,0 VtSay STR0015  //"Lote"
		@ 3,0 VtGet cLote /*pict '@!'*/  valid If(lEmpty,.t.,!Empty(cLote).Or. lNextLote) when Empty(cLote)
		If Rastro(cProduto,"S")	.and. cSubLote <> NIL
			@ 5,0 VtSay STR0016  //"Sub-Lote"
			@ 6,0 VtGet cSubLote pict '@!' valid If(lEmpty,.t.,!Empty(cSubLote))
		EndIf
		If dValid <> NIL
			@ 5,0 VtSay STR0017  //"Validade"
			@ 6,0 VtGet dValid pict '@D' valid If(lEmpty,.t.,!Empty(dValid))
		EndIf
		VTREAD
	Else
		VTClear
		@ 0,0 VTSay STR0014 //"Produto com rastro"
		@ 1,0 VtSay STR0018  	 //"Lote:"
		@ 1,7 VtGet cLote pict '@!'  valid If(lEmpty,.t.,!Empty(cLote)) when Empty(cLote)
		VTREAD
		If Rastro(cProduto,"S")	.and. cSubLote <> NIL
			VTClear(1,0,VTMaxCol(),VTMaxRow())
			@ 1,00 VtSay STR0019  //"Sub-Lote:"
			@ 1,10 VtGet cSubLote pict '@!' valid If(lEmpty,.t.,!Empty(cSubLote))
			VTREAD
		EndIf
		If dValid <> NIL
			VTClear(1,0,VTMaxCol(),VTMaxRow())
			@ 1,0 VtSay STR0020  //"Validade:"
			@ 1,10 VtGet dValid pict '@D' valid If(lEmpty,.t.,!Empty(dValid))
			VTREAD
		EndIf
	EndIf
EndIf
VtRestore(,,,,aSave)
If VTLastKey() == 27
	VTAlert(STR0021,STR0011,.t.,3000) //"Lote invalido"###"Aviso"
	Return .f.
EndIf

//������������������������������������������������������������Ŀ
//�Gera Lote automatico considerando a formula B1_FORMLOT caso �
//�exista                                                      �
//��������������������������������������������������������������
If lNextLote .And. Empty(cLote)
	cLote := NextLote(SC2->C2_PRODUTO,"L")
EndIf
Return .t.

/*
�������������������������������������������������������������������������������
�������������������������������������������������������������������������������
���������������������������������������������������������������������������Ŀ��
���Fun�ao    � CBChkSer   � Autor � SIGAACD               � Data � 17/03/08 ���
���������������������������������������������������������������������������Ĵ��
���Descri�ao �Verifica se o produto usa  numero de Serie                    ���
���������������������������������������������������������������������������Ĵ��
��� Uso      �SIGAACD                                                       ���
����������������������������������������������������������������������������ٱ�
�������������������������������������������������������������������������������
�������������������������������������������������������������������������������
*/
Function CBChkSer(cProduto)
Local aAreaSB5:=SB5->(GetArea())
Local lRet := .f.
SB5->(DbSetOrder(1))
If	Localiza(cProduto) .And. SB5->(DbSeek(xFilial("SB5")+cProduto)) .AND. (SB5->B5_NSERIE == "S")
	lRet := .t.
EndIf
RestArea(aAreaSB5)

Return lRet
/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    � CBNumSer � Autor � Anderson Rodrigues    � Data � 30/06/04 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Valida o numero de serie informado                         ���
�������������������������������������������������������������������������Ĵ��
���Retorno   � Logico - .t. / .f.                                         ���
�������������������������������������������������������������������������Ĵ��
���Uso       � SIGAACD                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Function CBNumSer(cNumSer,cSerie,aEtiqueta,lEstorno)
Local aSave    := {}
Local lCBVldNS := .F.
Local lCont    := .T.
Local lRet     := .T.

Private cPENumSer:= cNumSer
Default cSerie := Space(Len(CB8->CB8_NUMSER))
Default lEstorno := .F.

If ExistBlock("CBVLDNS")
	// Ponto de entrada que permite validacao especifica do numero de serie:
	lCBVldNS := ExecBlock("CBVLDNS",.F.,.F.,{cNumSer,cSerie,aEtiqueta})
	lCont    := .F.
	lRet     := (If(ValType(lCBVldNS) == "L",lCBVldNS,.F.))
	cNumSer  := cPENumSer
EndIf

If	lCont
	If	Empty(cNumSer)
		aSave:= VTSAVE()
		VTClear()
		If VTModelo()=="RF"
			@ 0,0 VTSay STR0022 //"Leitura do Numero"
			@ 1,0 VTSay STR0023 //"de Serie"
			@ 3,0 VTGet cNumSer pict '@!' Valid VldNumSer(cNumSer,cSerie,lEstorno)
		Else
			@ 0,0 VTSay STR0024 //"Numero de Serie"
			@ 1,0 VTGet cNumSer pict '@!' Valid VldNumSer(cNumSer,cSerie,lEstorno)
		EndIf
		VTREAD
		VtRestore(,,,,aSave)
		If VTLastKey() == 27
			VTAlert(STR0025,STR0011,.t.,3000) //"Numero de Serie invalido"###"Aviso"
			lRet := .F.
		EndIf
	ElseIf !VldNumSer(cNumSer,cSerie,lEstorno)
		lRet:= .F.
	EndIf
EndIf
Return(lRet)
/*
������������������������������������������������������������������������������
������������������������������������������������������������������������������
��������������������������������������������������������������������������Ŀ��
���Fun��o    � VldNumSer � Autor � Anderson Rodrigues    � Data � 30/06/04 ���
��������������������������������������������������������������������������Ĵ��
���Descri��o � Funcao Auxiliar chamada pela Funcao CBNumSer                ���
��������������������������������������������������������������������������Ĵ��
���Uso       � SIGAACD                                                     ���
���������������������������������������������������������������������������ٱ�
������������������������������������������������������������������������������
������������������������������������������������������������������������������
*/

Static Function VldNumSer(cNumSer,cSerie,lEstorno)
Local lRet 		:= .T.
Local lDif		:= (cNumSer # cSerie)
Local lSubNSer	:= .F.
Local cSubNSer	:= SuperGetMV("MV_SUBNSER",.F.,'1')

Default lEstorno := .F.

If ! Empty(cSerie) .And. Empty(cNumSer)
	lRet :=  .F.
Endif

If lRet
	lSubNSer := !lEstorno .And. lDif .And. IsIncallStack('ACDV166') .And. cSubNSer $ '2|3' .And. CB7->CB7_ORIGEM == "1"
EndIf

If lSubNSer
	If cSubNSer == '3'
		lRet := VTYesNo(STR0074,"SUBNSER",.T.)//""O numero de serie selecionado pelo sistema esta diferente do numero de serie lido, Deseja efetuar a troca?"
	Endif

	If lRet
		 lRet := CBVSUBNSER(CB8->CB8_PROD,CB8->CB8_LOCAL,cNumSer,cSerie)
	EndIf
ElseIf lRet
	If Empty(cNumSer)
		lRet :=  .T.
	ElseIf !Empty(cSerie) .And. lDif
		VtBeep(3)
		VtAlert(STR0025,STR0011,.t.,3000) //"Numero de Serie Invalido"###"Aviso"
		VtAlert(STR0026+cSerie,STR0011,.t.,4000) //"Informe o Numero de Serie "###"Aviso"
		VtKeyboard(Chr(20))  // zera o get
		lRet := .F.
	Endif
Endif

Return lRet

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    � CBCopyRec� Autor � Sandro/Motta          � Data � 20/11/01 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � duplica registro dentro da tabela                          ���
�������������������������������������������������������������������������Ĵ��
���Retorno   � numero do recno duplicado                                  ���
�������������������������������������������������������������������������Ĵ��
���Uso       � SIGAACD                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/

Function CBCopyRec(aCampos)
Local nRegOri:= Recno()
Local nRecno
Local aDados:={}
Local nX
Local cCampo
Local nPos
For nX := 1 to FCount()
	cCampo := FieldName(nX)
	nPos := Ascan(aCampos,{|x| Alltrim(x[1]) == Alltrim(cCampo)})
	If Empty(nPos)
		aadd(aDados,FieldGet(nX))
	Else
		aadd(aDados,aCampos[nPos,2])
	EndIf
Next
RecLock(Alias(),.t.)
For nX := 1 to FCount()
	FieldPut(nX,aDados[nX])
Next
MsUnLock()
nRecno := Recno()
Dbgoto(nRegOri)
Return nRecno

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    �CBExistLot� Autor � Sandro/Motta          � Data � 20/11/01 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Verifica a existencia do endereco com lote                 ���
�������������������������������������������������������������������������Ĵ��
���Retorno   � logico                                                     ���
�������������������������������������������������������������������������Ĵ��
���Uso       � SIGAACD                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Function CBExistLot(cProduto,cArmazem,cEndereco,cLote,cSLote)
Local lExist	:= .F.
Local lQuery	:= .F.
Local cQuery	:= ""
Local cAliasSBF := "SBF"
Local cChave	:= ""

If ! Rastro(cProduto)
	Return .t.
EndIf

SBF->(DBSetOrder(1))

lQuery		:= .T.
cAliasSBF 	:= "SBFQRY"
cQuery := "SELECT BF_LOTECTL, BF_NUMLOTE, R_E_C_N_O_ RECSBF FROM "+RetSqlName("SBF")
cQuery += " WHERE BF_FILIAL = '"+xFilial("SBF")+"' AND "
cQuery += " BF_LOCAL = '"+cArmazem+"' AND "
cQuery += " BF_LOCALIZ = '"+cEndereco+"' AND "
cQuery += " BF_PRODUTO = '"+cProduto+"' AND "
cQuery += " BF_LOTECTL = '"+cLote+"' AND "
cQuery += " BF_NUMLOTE = '"+cSLote+"' AND "
cQuery += " D_E_L_E_T_ = ' '"
cQuery := ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSBF)

While (cAliasSBF)->(! Eof() .and. If(lQuery,.T.,xFilial("SBF")+cArmazem+cEndereco+cProduto ==;
		BF_FILIAL+BF_LOCAL+BF_LOCALIZ+BF_PRODUTO) )
	If cLote+cSLote == (cAliasSBF)->(BF_LOTECTL+BF_NUMLOTE)
		lExist := .t.
		Exit
	EndIf
	(cAliasSBF)->(DbSkip())
End

If lQuery
	(cAliasSBF)->(DbCloseArea())
EndIf

If !lExist
	If Rastro(cProduto,"S")
		cChave := cProduto+cArmazem+cLote+cSLote
	Else
		cChave := cProduto+cArmazem+cLote
	EndIf

	DbSelectArea("SB8")
	SB8->(DbSetOrder(3))
	If SB8->(DbSeek(xFilial("SB8")+cChave))
		Return .t.
	EndIf
EndIf
Return lExist

Function CBChkMsg(cCodOpe)
Local aArea	:= GetArea()
Local nL		:= VTRow()
Local nC		:= VTCol()
Local lReverso:= VtReverso()
Local aTela
Local cRotina :=''
Local cRotMsg :=''
Local nPos    := 0
CB1->(DbSeek(xFilial('CB1')+cCodOpe))
cRotina := CB1->CB1_ROTINA

If IsVtMenu()
	CBF->(DbSetOrder(2))
	If CBF->(DbSeek(xFilial("CBF")+Space(6)+"1"))
		MSCBFSem("vtmsg.sem","")
		While CBF->(! Eof() .and. CBF_FILIAL+CBF_PARA+CBF_STATUS==xFilial("CBF")+Space(6)+"1")
			cRotMsg :=Alltrim(CBF->CBF_ROTINA)
			nPos := at("(",cRotMsg)
			If nPos > 0
				cRotMsg:= Left(cRotMsg,nPos-1)
			EndIf
			If Empty(cRotina) .OR. cRotMsg $ cRotina
				CBF->(RecLock("CBF"))
				CBF->CBF_PARA:= cCodOpe
				CBF->(MsUnlock())
				Exit
			EndIf
			CBF->(DbSkip())
		EndDo
		MSCBASem("vtmsg.sem")
	EndIf
EndIf


CBF->(DbSetOrder(2))
CBF->(DbSeek(xFilial("CBF")+cCodOpe+"1"))
While CBF->(! Eof() .and. CBF_FILIAL+CBF_PARA+CBF_STATUS==xFilial("CBF")+cCodOpe+"1")
	If IsVtMenu()
		CBF->(RecLock("CBF"))
		CBF->CBF_STATUS:= '2'
		CBF->(MsUnlock())
		aTela := VTSave()
		VtAlert(CBF->CBF_MSG,STR0027+CBF->CBF_DE,.F.,NIL,3)  //'DE:'
		If  CBF->CBF_RESPON#"0" .and. VtYesNo(STR0028,STR0029,.t.) //'Deseja responder'###'Mensagem'
			ACDV200(CBF->CBF_DE)
		EndIf
		If ! Empty(CBF->CBF_ROTINA) .AND. VtYesNo(STR0030,STR0029,.t.) //'Deseja executar a tarefa agora?'###'Mensagem'
			If ! Empty(CBF->CBF_KEYB)
				VTKeyBoard(Alltrim(CBF->CBF_KEYB))
			EndIf
			cRotMsg := Alltrim(CBF->CBF_ROTINA)
			cRotMsg +=If(Empty(at("(",cRotMsg)),"()","")
			VTAtuSem("SIGAACD",cRotMsg+" - ["+Alltrim(CBF->CBF_MSG)+']')
			If !(FindFunction(Alltrim(cRotMsg)))
				  VTAlert(STR0050+cRotMsG+STR0051,STR0004,.t.,4000)
				  Exit
			Else
				&(Alltrim(cRotMsg))
			EndIf
			VTAtuSem("SIGAACD","")
		Else
			CBF->(RecLock("CBF"))
			CBF->CBF_PENDEN:= 'X'
			CBF->(MsUnlock())
		EndIf
		VtRestore(,,,,aTela)
	Else
		If Empty(CBF->CBF_ROTINA)
			CBF->(RecLock("CBF"))
			CBF->CBF_STATUS:= '2'
			CBF->(MsUnlock())
			VtAlert(CBF->CBF_MSG,STR0027+CBF->CBF_DE,.T.,NIL,3)  //'DE:'
			If  CBF->CBF_RESPON#"0" .and. VtYesNo(STR0028,STR0029,.t.) //'Deseja responder'###'Mensagem'
				ACDV200(CBF->CBF_DE)
			EndIf
		EndIf
	EndIf
	CBF->(DbSkip())
EndDo
VtReverso(lReverso)
RestArea(aArea)
VtSay(nL,nC)
Return NIL

Function CBSendMsg(cPara,cMensagem,lResp,cRotina,cKeyb)
DEFAULT lResp  := .t.
DEFAULT cRotina:=""
DEFAULT cKeyb  := ""

RecLock("CBF",.t.)
CBF->CBF_FILIAL:= xFilial('CBF')
CBF->CBF_DE    := CbRetOpe()
CBF->CBF_PARA  := cPara
CBF->CBF_MSG   := cMensagem
CBF->CBF_STATUS:= '1'
CBF->CBF_DATA  := dDataBase
CBF->CBF_HORA  := Time()
CBF->CBF_DATAI := Inverte(dDataBase)
CBF->CBF_HORAI := Inverte(Time())
CBF->CBF_RESPON:= If(lResp,"1","0")
CBF->CBF_ROTINA:= cRotina
CBF->CBF_KEYB  := cKeyb
CBF->(MsUnlock())
Return

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    �CBProdLib � Autor � Sandro/Motta          � Data � 20/11/01 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Verifica se o produto esta bloqueado para inventario       ���
�������������������������������������������������������������������������Ĵ��
���Retorno   � logico                                                     ���
�������������������������������������������������������������������������Ĵ��
���Uso       � SIGAACD                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/

Function CBProdLib(cArmazem,cProduto,lMsg)
Local lBlq   := BlqInvent(cProduto,cArmazem)
Local lBlqAmz:= AvalBlqLoc(cProduto,cArmazem,Nil,.F.)

Default lMsg := .T.

If lBlq .And. lMsg
	VTBeep(3)
	VTAlert(STR0031,STR0013,.T.,4000) //'Produto bloqueado para inventario'###'Aviso'
ElseIf lBlqAmz .And. lMsg
	VTBeep(3)
	VTAlert(STR0049,STR0013,.T.,4000) //'Produto bloqueado para o armazem'###'Aviso'
EndIf

Return !lBlq

Function CBLog(cEvento,aDados)
Local nPos   := 0
Local nX     := 0
Local aEstru :={	{"01",{"CBG_CODPRO","CBG_QTDE","CBG_LOTE","CBG_SLOTE","CBG_ARM","CBG_END","CBG_NUMSEQ","CBG_DOC","CBG_CODETI","CBG_OBS"}},; // Enderecamento
	{"02",{"CBG_CODPRO","CBG_QTDE","CBG_LOTE","CBG_SLOTE","CBG_ARM","CBG_END","CBG_ARMDES","CBG_ENDDES","CBG_CODETI","CBG_OBS"}},; //transferencia
	{"03",{"CBG_CODPRO","CBG_QTDE","CBG_LOTE","CBG_SLOTE","CBG_ARM","CBG_END","CBG_NUMSEQ","CBG_DOC","CBG_CODETI","CBG_OBS"}},;  //Baixa CQ
	{"04",{"CBG_CODPRO","CBG_QTDE","CBG_LOTE","CBG_SLOTE","CBG_ARM","CBG_END","CBG_CODINV","CBG_CODCON","CBG_CODETI","CBG_OBS"}},;  //Inventario
	{"05",{"CBG_CODPRO","CBG_QTDE","CBG_LOTE","CBG_NOTAE","CBG_SERIEE","CBG_FORN","CBG_LOJFOR","CBG_ARM","CBG_CODETI","CBG_OBS"}},;  //Conferencia
	{"06",{"CBG_CODPRO","CBG_QTDE","CBG_LOTE","CBG_SLOTE","CBG_ARM","CBG_END","CBG_OP","CBG_CC","CBG_TM","CBG_CODETI","CBG_OBS"}},;  //requisicao
	{"07",{"CBG_CODPRO","CBG_QTDE","CBG_LOTE","CBG_SLOTE","CBG_ARM","CBG_END","CBG_OP","CBG_CODETI","CBG_ETIAUX","CBG_OBS"}},;  //Divisao Etiqueta
	{"08",{"CBG_CODPRO","CBG_QTDE","CBG_LOTE","CBG_SLOTE","CBG_ARM","CBG_END","CBG_OP","CBG_CODETI","CBG_ETIAUX","CBG_OBS"}},;  //Preparacao Enderecamento
	{"09",{"CBG_CODPRO","CBG_QTDE","CBG_LOTE","CBG_SLOTE","CBG_ARM","CBG_END","CBG_CODETI","CBG_OP","CBG_NOTAS","CBG_SERIES","CBG_CLI","CBG_LOJCLI","CBG_ORDSEP","CBG_VOLUME","CBG_SUBVOL","CBG_OBS"}},;  //Expedicao
	{"10",{"CBG_CODPRO","CBG_QTDE","CBG_LOTE","CBG_SLOTE","CBG_ARM","CBG_END","CBG_OP","CBG_CC","CBG_TM","CBG_CODETI","CBG_OBS"}},; //Devolucao
	{"11",{"CBG_CODPRO","CBG_QTDE","CBG_LOTE","CBG_SLOTE","CBG_NOTAE","CBG_SERIEE","CBG_FORN","CBG_LOJFOR","CBG_ARM","CBG_END","CBG_OP","CBG_CODETI","CBG_ETIAUX","CBG_NUMSEQ","CBG_OBS"}}} //Desmonta Embalagem
Local cCodOpe:= CBRetOpe()
Local cParam := GetNewPar("MV_LOGACD"," ")
Local cNoLog := GetNewPar("MV_NLOGACD"," ")

If Empty(cParam)
	Return
EndIf
//ponto de entrada
If ExistBlock("CBLOGALT")
	ExecBlock("CBLOGALT",.f.,.f.,aEstru)
End
If Alltrim(cParam) # "*" .and. ! cEvento $ cParam
	Return
EndIf
If cEvento $ cNoLog
	Return
EndIf
nPos := Ascan(aEstru,{|x| x[1] == cEvento})
If Empty(nPos)
	Return
EndIf

CBG->(DbSetOrder(1))
RecLock("CBG",.t.)
CBG->CBG_FILIAL:= xFilial("CBG")
CBG->CBG_EVENTO:= cEvento
CBG->CBG_CODOPE:= cCodOpe
CBG->CBG_USUARI:= __cUserID
CBG->CBG_DATA  := dDataBase
CBG->CBG_HORA  := Time()
For nX:= 1 to len(aEstru[nPos,2])
	If nX <= len(aDados) .and. aDados[nX] #NIL
		CBG->(FieldPut(FieldPos(aEstru[nPos,2,nX]),aDados[nX]))
	EndIf
Next
CBG->(MsUnLock())
Return

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o	 � CBRetMonit � Autor � Anderson Rodrigues  � Data � 21/08/02 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Retorna Array com as informacoes de Monitoramento da OP    ���
�������������������������������������������������������������������������Ĵ��
��� Uso		 � SIGAACD                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/

Function CBRetMonit(cOP)
Local aRetCBH  := {}
Local cDescTra := Space(Len(CBI->CBI_DESCRI))

CBH->(DBSetOrder(1))
If ! CBH->(DBSeek(xFilial("CBH")+cOP))
	Return aRetCBH
Endif
While ! CBH->(EOF()) .and. CBH->CBH_FILIAL+CBH->CBH_OP == xFilial("CBH")+cOP
	cDescTra := Posicione('CBI',1,xFilial("CBI")+CBH->(CBH_TRANSA+CBH_TIPO),"CBI_DESCRI")
	CBH->(aadd(aRetCBH,{CBH_OP,CBH_TRANSA,cDescTra,CBH_OPERAC,Str(CBH_QTD,10,2),CBH_DTINI,CBH_HRINI,CBH_TIPO}))
	CBH->( DBSkip() )
Enddo
aRetCBH:=aSort(aRetCBH)
Return aRetCBH

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    �CBQtdVar  � Autor � Anderson Rodrigues    � Data � 01/08/02 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Verfica se o produto tem quantidade variavel               ���
�������������������������������������������������������������������������Ĵ��
���Parametros� cCodProd  = codigo do produto                              ���
�������������������������������������������������������������������������Ĵ��
���Retorno   � Logico -  Se .T. indica que a quantidade e variavel        ���
�������������������������������������������������������������������������Ĵ��
���Uso       � SIGAACD                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/

Function CBQtdVar(cCodProd)
Local aArea:= SB5->(GetArea())
Local lRet

SB5->(DbSetOrder(1))
If ! SB5->(DbSeek(xFilial("SB5")+cCodProd))
	lRet := .f.
EndIf
If SB5->B5_QTDVAR == "1"
	lRet := .t.
Else
	lRet := .f.
Endif
If ExistBlock("CBQTDVAR")
	lRetPE := ExecBlock("CBQTDVAR",.F.,.F.,{cCodProd})
	If	Valtype(lRetPE) == 'L'
		lRet := lRetPE
	EndIf
EndIf
RestArea(aArea)
Return lRet

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    �CBProdxEnd� Autor � Anderson Rodrigues    � Data � 01/08/02 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o �Retorna Array com os Enderecos existentes para o Produto    ���
�������������������������������������������������������������������������Ĵ��
���Parametros� cCodProd  = Codigo do produto                              ���
�������������������������������������������������������������������������Ĵ��
���Uso       � SIGAACD                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/

Function CBProdxEnd(cCodProd,cArm,cNumSeq)
Local   aRetCBJ := {}

CBJ->(DbSetOrder(1))
If CBJ->(DbSeek(xFilial("SBJ")+cCodProd))
	If cArm == AlmoxCQ()
		SD7->(DbSetOrder(3))
		If SD7->(Dbseek(xFilial("SD7")+cCodProd+cNumSeq))
			cArm := SD7->D7_LOCDEST
		Else
			lRetSD7 := .f.
		Endif
	EndIf
	While ! CBJ->(EOF()) .and. CBJ->CBJ_FILIAL+CBJ->CBJ_CODPRO+CBJ->CBJ_ARMAZ == xFilial('CBJ')+cCodProd+cArm
		CBJ->(aadd(aRetCBJ,{CBJ->CBJ_ENDERE,Str(0,6,2)}))
		CBJ->( DbSkip() )
	Enddo
Endif
Return aRetCBJ

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    �CBImpEti  � Autor � Anderson Rodrigues    � Data � 01/08/02 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Verfica se imprime etiqueta do   produto                   ���
�������������������������������������������������������������������������Ĵ��
���Parametros� cCodProd  = codigo do produto                              ���
�������������������������������������������������������������������������Ĵ��
���Retorno   � Logico -  Se .T. indica que imprime                        ���
�������������������������������������������������������������������������Ĵ��
���Uso       � SIGAACD                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/

Function CBImpEti(cCodProd)
Local aArea	    := SB5->(GetArea())
Local lRet 	    := .T.
Local lCBETIQUE := ExistBlock("CBETIQUE")
Local lRetPE    := .F.

SB5->(DbSetOrder(1))
If ! SB5->(DbSeek(xFilial("SB5")+cCodProd))
	lRet := .t.
EndIf
If SB5->B5_IMPETI == "2"  // nao IMPRIME
	lRet := .f.
Else // sim
	lRet := .t.
Endif

If lCBETIQUE
	lRetPE := ExecBlock("CBETIQUE",.F.,.F.,{cCodProd,lRet})
	If ValType(lRetPE) == "L"
		lRet := lRetPE
	EndIf
EndIf

RestArea(aArea)
Return lRet

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    � CBArmProc  � Autor � Desenv.    ACD      � Data � 24/12/02 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Verifica se o produto vai para o armazem de processo (99)  ���
�������������������������������������������������������������������������Ĵ��
��� Uso      � SigaACD                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/

Function CBArmProc(cProduto,cTM)
Local cSB1Aprop,cSF5Aprop
Local lRet:= .f.

SB1->(DbSetOrder(1))
If ! SB1->(DbSeek(xFilial("SB1")+cProduto))
	Return(lRet)
Endif
cSB1Aprop:= SB1->B1_APROPRI
If cSB1Aprop == "I"
	SF5->(DbSetOrder(1))
	If ! SF5->(DbSeek(xFilial("SF5")+cTM))
		Return(lRet)
	Endif
	cSF5Aprop:= SF5->F5_APROPR
	If cSF5Aprop == "N"
		lRet:= .t.
	Endif
Endif
Return(lRet)

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �ACDXFUN   �Autor  �ACD                 � Data �  01/21/03   ���
�������������������������������������������������������������������������͹��
���Desc.     � Funcao para ser utilizando em trace                        ���
���          �                                                            ���
�������������������������������������������������������������������������͹��
���Uso       � AP6                                                        ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Function CBTrace(cConteudo,cConteudo2)
Local nHandle2
Local cArquivo := "CB"+Right(Dtos(dDatabase),6)+".log"
Local nDif
Local oFile		:= Nil

DEFAULT cConteudo :=""
Static nSeconds := 0

If ! File(cArquivo)
	oFile := FwFileWriter():New( cArquivo, .T. )
	If !oFile:Create()
		Return
	EndIf

Else
	If ( nHandle2 := oFile:nHandLen == -1 )
		Return
	EndIf

EndIf

If nSeconds == 0
	nDif := 0
Else
	nDif := ( Seconds() - nSeconds )
EndIf

oFile:Write( Alltrim( FunName() ) + ":"+Alltrim(ProcName(1))+"("+Alltrim(Str(ProcLine(1),5))+ ")"+STR0046+Time()+STR0047+Str(nDif,8,2)+If(! Empty(cConteudo),STR0048+cConteudo,"" ) + chr(13) + chr(10) ) //" Atual:"###" Diferenca:"###" Obs: "
If cConteudo2 <> NIL
	oFile:Write( VarInfo( cConteudo, cConteudo2,,.F.) )
EndIf

nSeconds:= Seconds()
oFile:Close()
Return


/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �CBItPallet�Autor  �Sandro              � Data �  02/21/03   ���
�������������������������������������������������������������������������͹��
���Desc.     �Retorna array com as etiquetas contidas no pallet           ���
���          �                                                            ���
�������������������������������������������������������������������������͹��
���Uso       � ACDSTD                                                     ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Function CBItPallet(cID)
Local aItens := {}
Local aRetPE := {}
Local aArea  := CB0->(GetArea())
cId := Padr(cID,TamSx3("CB0_PALLET")[1])
If Empty(cID)
	Return {}
EndIf
If UsaCB0("01")
	CB0->(DbSetOrder(5))
	If CB0->(DbSeek(xFilial("CB0")+cId))
		While CB0->(!Eof() .and. CB0_FILIAL+CB0_PALLET == xFilial("CB0")+cId)
			aadd(aItens,CB0->CB0_CODETI)
			CB0->(DbSkip())
		EndDo
	EndIf
EndIf
If	ExistBlock("CBPALLET")
	aRetPE := ExecBlock("CBPALLET",.F.,.F.,{cID})
	If	Valtype(aRetPE) == 'A'
		aItens := aRetPE
	EndIf
EndIf
RestArea(aArea)
Return aItens

/*
������������������������������������������������������������������������������
������������������������������������������������������������������������������
��������������������������������������������������������������������������Ŀ��
���Fun��o     � CBTotOP    � Autor � Desenv.    ACD      � Data � 19/11/02 ���
��������������������������������������������������������������������������Ĵ��
���Descri��o  � Retorna a quantidade necessaria do componente para atender ���
���           � a producao da OP informada                                 ���
��������������������������������������������������������������������������Ĵ��
���Parametros � cOP   -> Numero da OP a ser analisada                      ���
���           � cComp -> Codigo do produto / componente a ser analisado	   ���
��������������������������������������������������������������������������Ĵ��
��� Uso       � SigaACD                                                    ���
���������������������������������������������������������������������������ٱ�
������������������������������������������������������������������������������
������������������������������������������������������������������������������
*/

Function CBTotOP(cOP,cComp)
Local nQtdTotSC2:= 0

SC2->(DbSetOrder(1))
If SC2->(DbSeek(xFilial("SC2")+cOP))
	SG1->(DbSetOrder(2))
	If SG1->(DbSeek(xFilial("SG1")+cComp+SC2->C2_PRODUTO))
		nQtdTotSC2:= nQtdTotSC2+(SC2->C2_QUANT*SG1->G1_QUANT)
	Else
		nQtdTotSC2:= nQtdTotSC2+SC2->C2_QUANT
	EndIf
EndIf
Return(nQtdTotSC2)

/*
������������������������������������������������������������������������������
������������������������������������������������������������������������������
��������������������������������������������������������������������������Ŀ��
���Fun��o     � CBRatReq   � Autor � Desenv.    ACD      � Data � 19/11/02 ���
��������������������������������������������������������������������������Ĵ��
���Descri��o  � Realiza o Rateio das quantidades requisitadas entre as OP's���
��������������������������������������������������������������������������Ĵ��
���PARAMETROS � cOP      -> Numero da O.P a ser analisada                  ���
���           � cProduto -> Codigo do Produto a ser analisado              ���
���           � nQtdReq  -> Quantidade total requisitada a ser rateada     ���
���           � nQtdNec  -> Quantidade total necessaria do produto para    ���
���           �             atender a producao de todas as OP's            ���
��������������������������������������������������������������������������Ĵ��
��� Uso       � SigaACD                                                    ���
���������������������������������������������������������������������������ٱ�
������������������������������������������������������������������������������
������������������������������������������������������������������������������
*/

Function CBRatReq(cOP,cProduto,nQtdReq,nQtdNec)
Local nQtdRat := 0
Local nPercRat:= 0

SC2->(DbSetOrder(1))
If SC2->(DbSeek(xFilial("SC2")+cOP))
	SG1->(DbSetOrder(2))
	If SG1->(DbSeek(xFilial("SG1")+cProduto+SC2->C2_PRODUTO))
		nPercRat:= Round(((SC2->C2_QUANT*SG1->G1_QUANT)/nQtdNec)*100,4)
		nQtdRat := Round((nQtdReq*nPercRat)/100, TamSX3('B2_QATU')[2])
	Else
		nPercRat:= Round(((SC2->C2_QUANT)/nQtdNec)*100,4)
		nQtdRat := Round((nQtdReq*nPercRat)/100, TamSX3('B2_QATU')[2])
	EndIF
Endif
Return(nQtdRat)

/*
������������������������������������������������������������������������������
������������������������������������������������������������������������������
��������������������������������������������������������������������������Ŀ��
���Fun��o     � CBRLocImp  � Autor � Desenv.    ACD      � Data � 23/04/03 ���
��������������������������������������������������������������������������Ĵ��
���Descri��o  � Retorno o codigo do local de impressao                     ���
��������������������������������������������������������������������������Ĵ��
���PARAMETROS � cParam   -> codigo do parametro  ex. MV_IACD01, MV_IACD02  ���
��������������������������������������������������������������������������Ĵ��
��� Uso       � SigaACD                                                    ���
���������������������������������������������������������������������������ٱ�
������������������������������������������������������������������������������
������������������������������������������������������������������������������
*/

Function CBRLocImp(cParam)
Local nTamCod	:= TamSX3("CB5_CODIGO")[1]
Local cConteudo	:= GetMV(cParam)
If len(Alltrim(cConteudo)) > nTamCod
	cConteudo := &cConteudo
EndIf
Return PadR(cConteudo,nTamCod)

/*
������������������������������������������������������������������������������
������������������������������������������������������������������������������
��������������������������������������������������������������������������Ŀ��
���Fun��o     � BALSetPar  � Autor � Desenv.    ACD      � Data � 30/07/03 ���
��������������������������������������������������������������������������Ĵ��
���Descri��o  � Seta as configuracoes para Balanca Toledo                  ���
��������������������������������������������������������������������������Ĵ��
���PARAMETROS � cCOM       = Define a porta Serial                         ���
���           � cConector  = Define o conector a sr utilizado              ���
���           � cCanal     = Define qual canal do conector                 ���
���           � cBalanca   = Define a balanca ligada no canal do coletor.  ���
��������������������������������������������������������������������������Ĵ��
��� Uso       � SigaACD                                                    ���
���������������������������������������������������������������������������ٱ�
������������������������������������������������������������������������������
������������������������������������������������������������������������������
*/
Function CBBalToledo(cCOM, cConector, cCanal, cBalanca)
Local cBruto,cLiq,cTara
Local nRet
Local nRetDll		:= ExecInDLLOpen("TOLEDO.DLL")
Local aRet 			:= {0,0,0}
DEFAULT cCOM		:= '1'
DEFAULT cConector	:= '5'
DEFAULT cCanal		:= '1'
DEFAULT cBalanca	:= '-8132'
// cCOM+cConector+cCanal+cBalanca+spac(12)
//	154-8132____________
// 12345678901234567890 = 20
If !Left(cBalanca,1) $ '+-'
	cBalanca := '-'+AllTrim(cBalanca)
EndIf

cBruto := cCOM+spac(12)+cConector+cCanal+cBalanca
cLiq   := cBruto
cTara  := cBruto

If nRetDll== -1
	Final(STR0032) //"TOLEDO.DLL Nao Localizada. Ela deve estar no diretorio REMOTE."
Else
	nRet := ExeDLLRun2(nRetDll, 0, @cBruto )
	nRet := ExeDLLRun2(nRetDll, 1, @cLiq )
	nRet := ExeDLLRun2(nRetDll, 2, @cTara )
	ExecInDLLClose(nRetDll)
	aRet[1] := Val(cBruto)
	aRet[2] := Val(cLiq)
	aRet[3] := Val(cTara)
EndIf
Return aRet

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun�ao    � CBLogExp � Autor � Anderson Rodrigues    � Data � 12/01/04 ���
�������������������������������������������������������������������������Ĵ��
���Descri�ao � Grava Log dos Eventos da Expedicao                         ���
�������������������������������������������������������������������������Ĵ��
��� Uso      � SIGAACD                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/

Function CBLogExp(cOrdSep)
Local nX
Local cStatus
Local aLogExp:={}

CB7->(DbSetOrder(1))
CB7->(DbSeek(xFilial("CB7")+cOrdSep))
If Empty(CB7->CB7_STATUS) .or. CB7->CB7_STATUS == "0"
	cStatus:= STR0033 //"Nao iniciado"
ElseIf CB7->CB7_STATUS == "1"
	cStatus:= STR0034 //"Em separacao"
ElseIf CB7->CB7_STATUS == "2"
	cStatus:= STR0035 //"Separacao finalizada"
ElseIf CB7->CB7_STATUS == "3"
	cStatus:= STR0036 //"Em processo de embalagem"
ElseIf CB7->CB7_STATUS == "4"
	cStatus:= STR0037 //"Embalagem Finalizada"
ElseIf CB7->CB7_STATUS == "5"
	cStatus:= STR0038 //"Nota gerada"
ElseIf CB7->CB7_STATUS == "6"
	cStatus:= STR0039 //"Nota impressa"
ElseIf CB7->CB7_STATUS == "7"
	cStatus:= STR0040 //"Volume impresso"
ElseIf CB7->CB7_STATUS == "8"
	cStatus:= STR0041 //"Em processo de embarque"
ElseIf CB7->CB7_STATUS == "9"
	cStatus:=  STR0042 //"Finalizado"
EndIf

CB9->(DBSetOrder(1))
CB9->(DbSeek(xFilial("CB9")+cOrdSep))
While CB9->(! Eof() .and. CB9_FILIAL+CB9_ORDSEP == xFilial("CB9")+cOrdSep)
	CB8->(DbSetOrder(1))
	CB8->(DbSeek(xFilial("CB8")+cOrdSep+CB9->CB9_ITESEP+CB9->CB9_SEQUEN+CB9->CB9_PROD))
	//Alimenta array do log...
	aadd(aLogExp,{	CB9->CB9_PROD,;
		CB9->CB9_QTESEP,;
		CB9->CB9_LOTECT,;
		CB9->CB9_NUMLOT,;
		CB9->CB9_LOCAL,;
		CB9->CB9_LCALIZ,;
		CB9->CB9_CODETI,;
		CB8->CB8_OP,;
		CB7->CB7_NOTA,;
		CB7->CB7_SERIE,;
		CB7->CB7_CLIENT,;
		CB7->CB7_LOJA,;
		cOrdSep,;
		CB9->CB9_VOLUME,;
		CB9->CB9_SUBVOL,;
		cStatus})
	CB9->(DbSkip())
EndDo

For nX :=1 to Len(aLogExp)
	CBLog("09",{aLogExp[nX,01],aLogExp[nX,02],aLogExp[nX,03],aLogExp[nX,04],aLogExp[nX,05],aLogExp[nX,06],aLogExp[nX,07],aLogExp[nX,08],aLogExp[nX,09],;
		aLogExp[nX,10],aLogExp[nX,11],aLogExp[nX,12],aLogExp[nX,13],aLogExp[nX,14],aLogExp[nX,15],aLogExp[nX,16]})
Next
Return

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun�ao    �CBFlagSC5 � Autor � ACD                   � Data � 12/01/04 ���
�������������������������������������������������������������������������Ĵ��
���Descri�ao �Grava o status do embarque no Pedido de Vendas (flag)       ���
�������������������������������������������������������������������������Ĵ��
��� Uso      �SIGAACD                                                     ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/

Function CBFlagSC5(cFlag,cOrdSep)
Local  cPedido:=Space(6)

CB7->(DbSetOrder(1))
CB7->(DbSeek(xFilial("CB7")+cOrdSep))
SD2->(dbSetOrder(3))
If ! Empty(CB7->CB7_PEDIDO) //por pedido de venda
	cPedido := CB7->CB7_PEDIDO
ElseIf ! Empty(CB7->CB7_NOTA+CB7->CB7_SERIE+CB7->CB7_CLIENT+CB7->CB7_LOJA) // por nota
	SD2->(DbSeek(xFilial("SD2")+CB7->CB7_NOTA+CB7->CB7_SERIE+CB7->CB7_CLIENT+CB7->CB7_LOJA))
	cPedido:=SD2->D2_PEDIDO
Else // op
	Return
EndIf
SC5->(DbSetOrder(1))
If SC5->(DbSeek(xFilial("SC5")+cPedido))
	RecLock("SC5",.f.)
	SC5->C5_PREPEMB := cFlag
	SC5->(MsUnlock())
EndIf
Return

Function CBACDOK()
RPCSetType(3)
Return

/*/
������������������������������������������������������������������������������
������������������������������������������������������������������������������
��������������������������������������������������������������������������ͻ��
��� Funcao   �CBOrdemSix � Autor � Anderson Rodrigues � Data � 15/04/04    ���
��������������������������������������������������������������������������͹��
���Descri��o �Retorna a Ordem do Indice referente ao NickName informado    ���
��������������������������������������������������������������������������͹��
���Uso       �GENERICO                                                     ���
��������������������������������������������������������������������������ͼ��
������������������������������������������������������������������������������
������������������������������������������������������������������������������
/*/

Function CBOrdemSix(cAlias,cNickName)
Local nOrdem:= 1

DbSelectArea("SIX")
DbSetOrder(1)
SIX->(DbSeek(cAlias))
While ! SIX->(EOF()) .and. SIX->INDICE == cAlias
	If Alltrim(SIX->NICKNAME) == cNickName
		nOrdem:= Val(SIX->ORDEM)
		Exit
	Endif
	SIX->(DbSkip())
Enddo
Return(nOrdem)
/*
�������������������������������������������������������������������������������
�������������������������������������������������������������������������������
���������������������������������������������������������������������������Ŀ��
���Fun��o    � CBAlert    � Autor � Anderson Rodrigues  � Data � 06/04/04   ���
���������������������������������������������������������������������������Ĵ��
���Descri��o � Tratamento do Alert para RF, Microterminal e Protheus        ���
���������������������������������������������������������������������������Ĵ��
��� Uso      �SIGAACD                                                       ���
����������������������������������������������������������������������������ٱ�
�������������������������������������������������������������������������������
�������������������������������������������������������������������������������
*/

Function CBAlert(cTexto,cTitulo,lCenter,nTime,nBeep,lLimpa)
Default lLimpa  := .f.
Default nBeep   := 0

If TerProtocolo() == 'PROTHEUS'
	MsgAlert(cTexto)
ElseIf TerProtocolo() == 'VT100' .and. VtModelo() == "RF"
	VtAlert(cTexto,cTitulo,lCenter,nTime,nBeep)
	If lLimpa
		VtKeyBoard(chr(20))
	EndIf
Else
	If Len(cTexto) > If(VtModelo() == "MT44",60,30)
		TerAlert(cTexto,cTitulo,Nil,nBeep)
	Else
		TerAlert(cTexto,cTitulo,nTime,nBeep)
	EndIf
EndIf
Return
/*
�������������������������������������������������������������������������������
�������������������������������������������������������������������������������
���������������������������������������������������������������������������Ŀ��
���Fun��o	 � CBYesNo    � Autor � Anderson Rodrigues  � Data � 08/04/04   ���
���������������������������������������������������������������������������Ĵ��
���Descri��o � Tratamento do YesNo para RF e Microterminal                  ���
���������������������������������������������������������������������������Ĵ��
��� Uso      �SIGAACD                                                       ���
����������������������������������������������������������������������������ٱ�
�������������������������������������������������������������������������������
�������������������������������������������������������������������������������
*/

Function CBYesNo(cMensagem,cTitulo,lCenter)
If TerProtocolo() == 'PROTHEUS'
	Return MsgYesNo(cMensagem)
ElseIf TerProtocolo() == 'VT100' .and. VtModelo() == "RF"
	Return VtYesNo(cMensagem,cTitulo,lCenter)
Else
	Return TerYesNo(cMensagem,cTitulo)
EndIf
Return .t.

/*/
������������������������������������������������������������������������������
������������������������������������������������������������������������������
��������������������������������������������������������������������������ͻ��
���Fun��o    �CBRetRecur � Autor � Anderson Rodrigues � Data �  05/05/04   ���
��������������������������������������������������������������������������͹��
���Descri��o �Retorna o recurso de acordo com os parametros informados     ���
��������������������������������������������������������������������������͹��
���Uso       � SIGAACD/GENERICO                                            ���
��������������������������������������������������������������������������ͼ��
������������������������������������������������������������������������������
������������������������������������������������������������������������������
/*/

Function CBRetRecur(cOP,cOperacao,lModelo1)
Local cRecurso  := " "
Local cProduto  := " "
Local cRoteiro  := " "
Default lModelo1:= .f.
cOP := Padr(cOP,len(SH8->H8_OP))

If lModelo1
	SH8->(DbSetOrder(1))
	If ! SH8->(DbSeek(xFilial("SH8")+cOP+cOperacao))
		Return(cRecurso)
	Endif
	cRecurso:= SH8->H8_RECURSO
Else
	SC2->(DbSetOrder(1))
	If ! SC2->(DbSeek(xFilial("SC2")+Padr(cOP,Len(SH6->H6_OP))))
		Return(cRecurso)
	Endif
	cProduto:= SC2->C2_PRODUTO
	If !Empty(SC2->C2_ROTEIRO)
		cRoteiro := SC2->C2_ROTEIRO
	Else
		SB1->(DbSetorder(1))
		If SB1->(DbSeek(xFilial('SB1')+cProduto)) .And. !Empty(SB1->B1_OPERPAD)
			cRoteiro := SB1->B1_OPERPAD
		Else
			cRoteiro := StrZero(1, Len(SG2->G2_CODIGO))
		EndIf
	Endif
	If a630SeekSG2(1,cProduto,xFilial("SG2")+cProduto+cRoteiro+cOperacao)
		cRecurso:= SG2->G2_RECURSO
	Endif
Endif
Return(cRecurso)

/*
�������������������������������������������������������������������������������
�������������������������������������������������������������������������������
���������������������������������������������������������������������������Ŀ��
���Fun�ao    �CBSeekNumSer� Autor � Anderson Rodrigues    � Data � 06/07/04 ���
���������������������������������������������������������������������������Ĵ��
���Descri�ao �Verifica se o produto tem numero de Serie separado no CB9     ���
���������������������������������������������������������������������������Ĵ��
��� Uso      �SIGAACD - Rotinas de Expedicao                                ���
����������������������������������������������������������������������������ٱ�
�������������������������������������������������������������������������������
�������������������������������������������������������������������������������
*/
Function CBSeekNumSer(cOrdSep,cProduto)
Local nOrdem:= CB9->(IndexOrd())
Local nRecno:= CB9->(Recno())
Local lRet  := .f.

CB9->(DBSetOrder(9))
CB9->(DbSeek(xFilial("CB9")+cOrdSep))
While CB9->(! Eof() .and. CB9_FILIAL+CB9_ORDSEP == xFilial("CB9")+cOrdSep)
	If CB9->CB9_PROD # cProduto
		CB9->(DbSkip())
		Loop
	Endif
	If Empty(CB9->CB9_NUMSER)
		CB9->(DbSkip())
		Loop
	Endif
	lRet:= .t.
	Exit
Enddo
CB9->(DbSetOrder(nOrdem))
CB9->(DbGoto(nRecno))
Return(lRet)

/*
�������������������������������������������������������������������������������
�������������������������������������������������������������������������������
���������������������������������������������������������������������������Ŀ��
���Fun�ao    �CBAtuLote   � Autor � Andre Anjos           � Data � 13/02/08 ���
���������������������������������������������������������������������������Ĵ��
���Descri�ao �Atualiza a numeracao de lote no CB0 apos classificacao de     ���
���          �pre-nota.                                                     ���
���������������������������������������������������������������������������Ĵ��
���Parametros�cLote  = Numero de lote gravado no SD1                        ���
���          �cSLote = Numero de sub-lote gravado no SD1                    ���
���          �dDtVld = Data de validade gravada no SD1                      ���
���������������������������������������������������������������������������Ĵ��
��� Uso      �SIGAACD - Rotinas de Recebimento                              ���
����������������������������������������������������������������������������ٱ�
�������������������������������������������������������������������������������
�������������������������������������������������������������������������������
*/
Function CBAtuLote(cLote,cSLote,dDtVld)
Local aArea := GetArea()

dbSelectArea("CB0")
dbSetOrder(9)//CB0_FILIAL+CB0_FORNEC+CB0_LOJAFO+CB0_CODPRO+CB0_NFENT+CB0_SERIEE+CB0_ITNFE
dbSeek(xFilial("CB0")+SD1->(D1_FORNECE+D1_LOJA+D1_COD+D1_DOC+D1_SERIE+D1_ITEM))
While !EOF() .And. CB0->(CB0_FILIAL+CB0_FORNEC+CB0_LOJAFO+CB0_CODPRO+CB0_NFENT+CB0_SERIEE+CB0_ITNFE) == ;
							xFilial("CB0")+SD1->(D1_FORNECE+D1_LOJA+D1_COD+D1_DOC+D1_SERIE+D1_ITEM)
	RecLock("CB0",.F.)
	Replace CB0_LOTE With cLote
	If Rastro(CB0_CODPRO,"S")
		Replace CB0_SLOTE With cSLote
	EndIf
	Replace CB0_DTVLD With dDtVld
	MsUnLock()

	dbSkip()
End

RestArea(aArea)
Return

/*
�������������������������������������������������������������������������������
�������������������������������������������������������������������������������
���������������������������������������������������������������������������Ŀ��
���Fun�ao    �CBnfeDistAut� Autor � Andre Anjos           � Data � 24/04/08 ���
���������������������������������������������������������������������������Ĵ��
���Descri�ao �Realiza a distribuicao autom�tico a partir de uma movimentacao���
���          �de NFE ou producao                                            ���
���������������������������������������������������������������������������Ĵ��
���Parametros�cCod   = Codigo do produto                                    ���
���          �cDoc   = Documento                                            ���
���          �cSerie = Serie da NFE                                         ���
���          �cForn  = Codigo do cliente/fornecedor                         ���
���          �cLoja  = Loja do cliente/fornecedor                           ���
���������������������������������������������������������������������������Ĵ��
��� Uso      �SIGAACD - Documento de entrada                                ���
����������������������������������������������������������������������������ٱ�
�������������������������������������������������������������������������������
�������������������������������������������������������������������������������
*/
Function CBNFEDistAut(cCod,cDoc,cSerie,cCliFor,cLoja)
Local aArea    := GetArea()
Local cDistAut := SuperGetMv("MV_DISTAUT",.F.," ")
Local cMvCq	:= GetMvNNR('MV_CQ','98')

dbSelectArea("CB0")
dbSetOrder(6)
If dbSeek(xFilial("CB0")+cDoc+cSerie+cCliFor+cLoja+cCod)
	While !EOF() .And. CB0->(CB0_FILIAL+CB0_NFENT+CB0_SERIEE+CB0_FORNEC+CB0_LOJAFO+CB0_CODPRO) ==;
						xFilial("CB0")+cDoc+cSerie+cCliFor+cLoja+cCod
		If CB0->CB0_CODPRO == cCod
			RecLock("CB0",.F.)
			Replace CB0_LOCAL  With cMvCq
			Replace CB0_LOCALI With Substr(cDistAut,3,TamSX3("CB0_LOCALI")[1])
			Replace CB0_NUMSER With Substr(cDistAut,18,TamSX3("CB0_NUMSER")[1])
			MsUnLock()
		EndIf
		CB0->(dbSkip())
	End
EndIf

RestArea(aArea)
Return

/*
�������������������������������������������������������������������������������
�������������������������������������������������������������������������������
���������������������������������������������������������������������������Ŀ��
���Fun�ao    �CBOPDistAut � Autor � Andre Anjos           � Data � 24/04/08 ���
���������������������������������������������������������������������������Ĵ��
���Descri�ao �Processa a distribuicao automatica na tabela CBO para entradas���
���          �por producao.                                                 ���
���������������������������������������������������������������������������Ĵ��
���Parametros�cCod    = Codigo do produto                                   ���
���          �cOP     = Codigo da OP                                        ���
���          �cNumSeq = Numero sequencial do mov. de producao               ���
���������������������������������������������������������������������������Ĵ��
��� Uso      �SIGAACD - Apontamento de producao                             ���
����������������������������������������������������������������������������ٱ�
�������������������������������������������������������������������������������
�������������������������������������������������������������������������������
*/
Function CBOPDistAut(cCod,cOP,cNumSeq)
Local aArea := GetArea()
Local cDistAut := SuperGetMv("MV_DISTAUT",.F.," ")
Local cMvCq	:= GetMvNNR('MV_CQ','98')

BeginSQL Alias "CB0TMP"
	SELECT R_E_C_N_O_ CB0RECNO
	FROM %Table:CB0%
	WHERE %NotDel% AND
		CB0_FILIAL = %xFilial:CB0% AND
		CB0_CODPRO = %Exp:cCod% AND
		CB0_OP = %Exp:cOP% AND
		CB0_NUMSEQ = %Exp:cNumSeq%
EndSQL
While !CB0TMP->(EOF())
	CB0->(dbGoTo(CB0TMP->CB0RECNO))

	RecLock("CB0",.F.)
	Replace CB0_LOCAL  With cMvCq
	Replace CB0_LOCALI With Substr(cDistAut,3,TamSX3("CB0_LOCALI")[1])
	Replace CB0_NUMSER With Substr(cDistAut,18,TamSX3("CB0_NUMSER")[1])
	CB0->(MsUnLock())
	
	CB0TMP->(dbSkip())
EndDo
CB0TMP->(dbCloseArea())

RestArea(aArea)
Return

/*
�������������������������������������������������������������������������������
�������������������������������������������������������������������������������
���������������������������������������������������������������������������Ŀ��
���Fun�ao    �CBAtuItNFE  � Autor � Paulo Fco. Cruz Neto  � Data � 14/10/09 ���
���������������������������������������������������������������������������Ĵ��
���Descri�ao �Atualiza o campo CB0_ITNFE antes da classificacao da nota de  ���
���          �entrada.                                                      ���
���������������������������������������������������������������������������Ĵ��
��� Uso      �SIGAACD - Rotinas de Recebimento                              ���
����������������������������������������������������������������������������ٱ�
�������������������������������������������������������������������������������
�������������������������������������������������������������������������������
*/
Function CBAtuItNFE()
Local aArea  := {}
Local aCB0	 := {}
Local aSD1	 := {}
Local cImpIp := SuperGetMv("MV_IMPIP",.F.,"3")
Local cQuery := ""
Local lAtuali:= Inclui .Or. Altera .Or. If(Type("l103Class") != "U" .And. ValType("l103Class") == "L",l103Class,.F.)

//�������������������������������������������������������������������������������Ŀ
//� Atualiza os campos ref. a NFE quando o parametro MV_IMPIP = 1 ou A2_IMPIP = 1 �
//���������������������������������������������������������������������������������
 If (SA2->A2_IMPIP == "1" .Or. cImpIp == "1") .And. UsaCB0("01") .And. lAtuali
	aArea := GetArea()
	aCB0  := CB0->(GetArea())
	aSD1  := SD1->(GetArea())

	cQuery := "SELECT SD1.D1_DOC,SD1.D1_SERIE,SD1.D1_FORNECE,SD1.D1_LOJA,SD1.D1_PEDIDO,SD1.D1_ITEMPC,SD1.D1_COD,SD1.D1_ITEM,SD1.D1_LOCAL "
	cQuery += "FROM " + RetSqlName("SD1") + " SD1 "
	cQuery += "WHERE SD1.D1_FILIAL = '" + xFilial("SD1") + "' "
	cQuery += "AND SD1.D1_DOC = '" 		+ SF1->F1_DOC + "' "
	cQuery += "AND SD1.D1_SERIE = '" 	+ SF1->F1_SERIE + "' "
	cQuery += "AND SD1.D1_FORNECE = '" 	+ SF1->F1_FORNECE + "' "
	cQuery += "AND SD1.D1_LOJA = '" 	+ SF1->F1_LOJA + "' "
	cQuery += "AND SD1.D_E_L_E_T_ <> '*' "
	cQuery += "ORDER BY " + SqlOrder("SD1.D1_PEDIDO+SD1.D1_ITEMPC")

	cQuery := ChangeQuery(cQuery)
	DbUseArea(.T., "TOPCONN", TCGenQry(,,cQuery), 'SD1_TRB', .F., .T.)

	DbSelectArea("SD1_TRB")
	SD1_TRB->(DbGotop())
	While SD1_TRB->(!Eof())
		cQuery := "SELECT CB0.R_E_C_N_O_ RECNO "
		cQuery += "FROM " + RetSqlName("CB0") + " CB0 "
		cQuery += "WHERE CB0.CB0_FILIAL = '" + xFilial("CB0") + "' "
		cQuery += "AND CB0.CB0_FORNEC = '" 	+ SD1_TRB->D1_FORNECE + "' "
		cQuery += "AND CB0.CB0_LOJAFO = '" 	+ SD1_TRB->D1_LOJA + "' "
		cQuery += "AND CB0.CB0_PEDCOM = '" 	+ SD1_TRB->D1_PEDIDO + SD1_TRB->D1_ITEMPC + "' "
		cQuery += "AND CB0.CB0_NFENT = '' "
		cQuery += "AND CB0.CB0_SERIEE = '' "
		cQuery += "AND CB0.CB0_LOCAL = '' "
		cQuery += "AND CB0.CB0_ITNFE = '' "
		cQuery += "AND CB0.D_E_L_E_T_ <> '*' "
		cQuery += "ORDER BY " + SqlOrder("CB0.CB0_PEDCOM")

		cQuery := ChangeQuery(cQuery)
		DbUseArea(.T., "TOPCONN", TCGenQry(,,cQuery), 'CB0_TRB', .F., .T.)

		DbSelectArea("CB0_TRB")
		CB0_TRB->(DbGotop())
		While CB0_TRB->(!Eof())
			CB0->(DbGoTo(CB0_TRB->RECNO))
			RecLock("CB0",.F.)
			CB0->CB0_NFENT	:= SD1_TRB->D1_DOC
			//CB0->CB0_SERIEE	:= SD1_TRB->D1_SERIE
			SerieNfId("CB0",1,,,,SD1_TRB->D1_SERIE)
			CB0->CB0_LOCAL	:= SD1_TRB->D1_LOCAL
			CB0->CB0_ITNFE	:= SD1_TRB->D1_ITEM
			CB0->(MsUnLock())

			CB0_TRB->(DbSkip())
		EndDo
		CB0_TRB->(DbCloseArea())
		SD1_TRB->(DbSkip())
	EndDo
	SD1_TRB->(DbCloseArea())

	 RestArea(aArea)
	 CB0->(RestArea(aCB0))
	 SD1->(RestArea(aSD1))
EndIf

Return
/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �CBMonRF   �Autor  �Microsiga           � Data �  01/04/10   ���
�������������������������������������������������������������������������͹��
���Desc.     �Monitor Radio Frequencia                                    ���
���          �                                                            ���
�������������������������������������������������������������������������͹��
���Uso       �ACD							                              ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Function CBMonRF()
Local oDlg
Local oLbx
Local aLbx 	   :={}
Local oTimer
Local aButtons :={}
Local aSize    := MsAdvSize()

aAdd(aButtons,{"FRTONLINE"  	,{|| Monitora(oLbx,aLbx)}				,STR0057		}) // "Monitor"
aAdd(aButtons,{"CRITICA"  		,{|| Mensagem(oLbx,aLbx)}	  			,STR0058 		}) // "Mensagem"
aAdd(aButtons,{"AFASTAMENTO"	,{|| Desconect(oLbx,aLbx,oTimer)}   	,STR0059		}) // "Desconecta"

DEFINE MSDIALOG oDlg TITLE STR0060 FROM aSize[7],0 TO aSize[6],aSize[5] OF oMainWnd Pixel // "Monitor de RF"

	@ 01,01 LISTBOX oLbx FIELDS HEADER STR0061,STR0062,STR0063,STR0064,STR0065,STR0066,STR0067," " SIZES {15,20,20,20,10,10,10,10} SIZE 490,095 OF oDlg PIXEL	 // "Coletor","Usuario","Data","Hora","Tamanho","Programa Inicial","Rotina"
	oLbx:align := CONTROL_ALIGN_ALLCLIENT
	oLbx:bLDblClick := {|| Monitora(oLbx,aLbx)}

	CarregaItens(oLbx,aLbx)
	DEFINE TIMER oTimer INTERVAL 1000 ACTION AtuTela(oLbx,aLbx,oTimer) OF oDlg

ACTIVATE MSDIALOG oDlg ON INIT (	EnchoiceBar(oDlg, {|| oDlg:End()  },{|| oDlg:End()},,aButtons ), AtuTela(oLbx,aLbx,oTimer),oTimer:Activate())


Return nil



Static Function AtuTela(oLbx,aLbx,oTimer)
oTimer:Deactivate()
CarregaItens(oLbx,aLbx)
oTimer:Activate()
Return



Static Function CarregaItens(oLbx,aLbx)
Local nX,nPos
Local cLinha := Space(70)
Local cNumCol := ''
Local cUsuario:= ''
Local cData   := ''
Local cHora   := ''
Local cProgIni:= ''
Local cRotina := ''
Local nHTemp
Local aColetor 	:= Directory("VT*.SEM")
Local oFile		:= Nil

aLbx := {}
For nX := 1 to Len( aColetor )
	oFile  := FwFileReader():New( aColetor[nX,1] )
	nHTemp := oFile:nHandLen
	oFile:Close()
	If nHTemp > 0
		FErase(aColetor[nX,1])
	EndIf
Next

aColetor := Directory("VT*.SEM")
For nX := 1 to Len(aColetor)
	cLinha  := Memoread(aColetor[nX,1])
	cNumCol := Left(cLinha,3)
	cUsuario:= Subs(cLinha,4,25)
	cData   := stod(Subs(cLinha,29,8))
	cHora   := Subs(cLinha,37,8)
	cSize   := Str(Val(Subs(cLinha,45,03))+1,3)+" X "+Str(Val(Subs(cLinha,48,03))+1,3)
	cProgIni:= Subs(cLinha,51,8)
	cRotina := Subs(cLinha,59,30)
	nPos    := AsCan(aLbx,{|x|x[1]==cNumCol})
	If Empty(nPos)
		aadd(aLbx,{cNumcol,cUsuario,cData,cHora,cSize,cProgIni,cRotina,""})
	Else
		aadd(aLbx,{cNumcol,cUsuario,cData,cHora,cSize,cProgIni,cRotina,""})
	EndIf
Next
If Empty(aLbx)
	aadd(aLbx, {'','','','','','','',''})
EndIF

oLbx:SetArray( aLbx )
oLbx:bLine := {|| {aLbx[oLbx:nAt,1],aLbx[oLbx:nAt,2],aLbx[oLbx:nAt,3],aLbx[oLbx:nAt,4],aLbx[oLbx:nAt,5],aLbx[oLbx:nAt,6],aLbx[oLbx:nAt,7],aLbx[oLbx:nAt,8]} }
oLbx:Refresh()
Return
//================

Static Function Mensagem(oLbx,aLbx)
Local oMemo
Local cMemo
Local oFont
Local oDlgMsg

DEFINE FONT oFont NAME "Mono AS" SIZE 8,20

If Len(aLbx)==1 .AND. Empty(aLbx[oLbx:nAt,1])
	Return
EndIf

DEFINE MSDIALOG oDlgMsg FROM 0,0 TO 100,300  Pixel TITLE OemToAnsi(STR0052+aLbx[oLbx:nAt,1]) // "Mensagem para o coletor "
	@ 0,0 GET oMemo  VAR cMemo MEMO SIZE 150,30 OF oDlgMsg PIXEL
	TButton():New( 035,001, STR0053, oDlgMsg, {|| MemoWrite('VT'+aLbx[oLbx:nAt,1]+'.MSG',cMemo),oDlgMsg:End()}, 38, 11,,, .F., .t., .F.,, .F.,,, .F. ) // "Enviar"
	TButton():New( 035,111, STR0054, oDlgMsg, {|| oDlgMsg:End()}, 38, 11,,, .F., .t., .F.,, .F.,,, .F. ) // "Sair"
	oMemo:oFont:=oFont
ACTIVATE MSDIALOG oDlgMsg CENTERED

Return


//==========================================================
Static Function Monitora(oLbx,aLbx)
Local cFile   := 'VT'+aLbx[oLbx:nAt,1]+'.MON'
Local nMaxRow := Val(SubStr(aLbx[oLbx:nAt,5],1,3))
Local nMaxCol := Val(SubStr(aLbx[oLbx:nAt,5],7,3))
Local nBottom := 294
Local nRight  := 400
Local nI,nJ
Local nLin    := 3
Local nCol    := 3
Local bBlock
Private lSai := .t.

oFont  := TFont():New( "Mono AS", 16, 24, .F.,.T.,,,,,,,,,,, )
DEFINE FONT oFont2 NAME "Mono AS" SIZE 16,24 UNDERLINE BOLD

If Len(aLbx)==1 .AND. Empty(aLbx[oLbx:nAt,1])
	Return
EndIf
//Restaura a possicao
If nMaxRow== 8 .and. nMaxCol==20
	nBottom := 294
	nRight  := 400
ElseIf nMaxRow== 2 .and. nMaxCol==20
	nBottom := 105
	nRight  := 400
ElseIf nMaxRow== 2 .and. nMaxCol==40
	nBottom := 105
	nRight  := 790
EndIf

VTScrToFile(cFile,{{},{}})
DEFINE MSDIALOG oDlg TITLE OemToAnsi(STR0055+aLbx[oLbx:nAt,1]) FROM 00,00 TO nBottom,nRight PIXEL // "Monitorando Coletor "
TButton():New(if(nMaxrow==8,134,38),005, STR0054, oDlg, {|| lSai := .t.}, 38, 11,,, .F., .t., .F.,, .F.,,, .F. ) // "Sair"
aSayVt := Array(nMaxRow,nMaxCol,2)
For nI := 1 to nMaxRow
	nCol := 3
	For nJ := 1 to nMaxCol
		aSayVt[nI,nJ,2] := " "
		bBlock := &("{||aSayVt["+Str(nI,4)+","+Str(nJ,4)+",2]   }")
		aSayVt[nI,nJ,1] := TSay():New( nLin, nCol, bBlock,,,oFont, .F., .F., .F., .T.,,, 15, 17, .F., .F., .F., .F., .F. )
		nCol+=10
	Next
	nLin+=16
Next
ACTIVATE MSDIALOG oDlg CENTERED ON INIT Gerencia(cFile,aLbx[oLbx:nAt,1],oDlg)

While FErase(cFile) == -1 .and. file(cFile)
	Sleep(10)
End
Return


Static Function Gerencia(cFile,cNumTer,oDlg)
Local nI := 0
nI := 0
If lSai
	lSai := .F.
	While !lSai
		ProcessMessage()
		AtuMon(cFile,cNumTer,oDlg)
		sleep(1000)
	EndDo
	oDlg:End()
Else
	lSai := .T.
EndIf
Return .T.


Static Function AtuMon(cFile,cNumTer,oDlg)
Local nI,nJ
Local aOldSay := {,,}
Local aLoad
Local __aScreen
Local __aReverso
Local nHTemp
Local oFile		:= Nil
Local cFileVt	:= 'VT'+cNumTer+'.SEM'

IF ! file( cFileVt )
	lSai := .t.
	Return
EndIf

oFile := FwFileReader():New( cFileVt )
nHTemp := oFile:nHandLen
oFile:Close()
If nHTemp > 0
	FErase( cFileVt )
	lSai := .t.
	Return
EndIf

aLoad := VTFileToScr(cFile)
IF Len(aLoad) < 2
	Return
ENDIF
__aScreen  := aLoad[1]
__aReverso := aLoad[2]
For nI:= 1 to Len(__aScreen)
  If nI > Len(aSayVt) .or. lSai
	  Exit
  EndIf
  For nJ := 1 to Len(__aScreen[nI])
	  If nJ > Len(aSayVt[nI]) .or. lSai
		  Exit
	  EndIf
	  aOldSay[1] := aSayVt[nI,nJ,2]
	  aOldSay[2] := aSayVt[nI,nJ,1]:NCLRTEXT
	  aSayVt[nI,nJ,2] := SubStr(__aScreen[nI],nJ,1)
	  If SubStr(__aReverso[nI],nJ,1) == "0"
		  aSayVt[nI,nJ,1]:NCLRTEXT := CLR_BLACK
		  aSayVt[nI,nJ,1]:OFONT    := oFont
	  Else
		  aSayVt[nI,nJ,1]:NCLRTEXT := CLR_HRED //WHITE
		  aSayVt[nI,nJ,1]:OFONT    := oFont2
	  EndIf
	  If aSayVt[nI,nJ,2] # aOldSay[1] .or. aOldSay[2] # aSayVt[nI,nJ,1]:NCLRTEXT  // somente faz o refresh quando necessario
		  aSayVt[nI,nJ,1]:Refresh()
	  EndIf
  Next nJ
Next nI

sleep(10)
PROCESSMESSAGE()
Return .t.


Static Function Desconect(oLbx,aLbx,oTimer)
Local cFile
Local nCont := 0

If Empty(aLbx[oLbx:nAt,1]) .AND. Len(aLbx)==1
	Return
EndIf

If !  MsgYesNo(STR0056)// 'Confirma a desconexao do coletor selecionado?'
	Return
Else
	cFile := 'VT'+aLbx[oLbx:nAt,1]+'.MON'
	IF file( cFile )
		While FErase(cFile) == -1 .and. file(cFile)
			Sleep(10)
			nCont += 1
			If nCont > 100
				Alert (STR0075)
				Return
			EndIf
		End
	Endif
EndIf

cFile := 'VT'+aLbx[oLbx:nAt,1]+'.FIM'
MemoWrite(cFile,'fim')
AtuTela(oLbx,aLbx,oTimer)
Return
/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    � CB_SXESXF� Autor � Ricardo Berti         � Data �18/03/2011���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Funcao generica para geracao de chave e verificacao da     ���
���			 � existencia quando ja houver a sequencia gerada			  ���
�������������������������������������������������������������������������Ĵ��
���Parametros� ExpC1 = Caracter contendo o alias                          ���
���			 � ExpC2 = Caracter contendo o campo referencia      		  ���
���			 � ExpC3 = Caracter contendo o alias do SX8     			  ���
���			 � ExpC4 = Caracter contendo a ordem do alias    			  ���
�������������������������������������������������������������������������Ĵ��
���Retorno   �cChave - Chave com o numero sequencial gerado               ���
�������������������������������������������������������������������������Ĵ��
���Uso       �Generico - Modulos ACD                                      ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Function CB_SXESXF(cAlias,cCampo,cAliasSX8,cOrdem)

Local aArea    := (cAlias)->(GetArea())
Local nSaveSX8 := GetSx8Len()
Local cChave   := ""

//����������������������������������������������������������Ŀ
//� Gera a chave de ligacao									 �
//������������������������������������������������������������
dbSelectArea(cAlias)
dbSetOrder(cOrdem)
cChave := GetSxeNum(cAlias,cCampo,cAliasSX8,cOrdem)
While dbSeek(xFilial(cAlias)+cChave)
	While ( GetSx8Len() > nSaveSX8 )
		ConfirmSX8()
	 EndDo
	 cChave := GetSxeNum(cAlias,cCampo,cAliasSX8,cOrdem)
EndDo

RestArea(aArea)
Return(cChave)

/*���������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Fun��o    � EstCBED1 � Autor � Bruno Schmidt         � Data � 29/06/12 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Funcao generica limpar os registro da CBE quando for 	  ���
���			 � excluido um documento de entrada que possua conferencia    ���
�������������������������������������������������������������������������Ĵ��
���Parametros� ExpC1 = Caracter contendo o numero da Documento            ���
���			 � ExpC2 = Caracter contendo a Serie do Documento      		  ���
���			 � ExpC3 = Caracter contendo o Codigo do fornecedor     	  ���
���			 � ExpC4 = Caracter contendo a Loja do Fornecedor			  ���
���			 � ExpC5 = Caracter contendo o Codigo do Produto  			  ���
�������������������������������������������������������������������������Ĵ��
���Uso       �Generico - Modulos ACD                                      ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
���������������������������������������������������������������������������*/
Function EstCBED1(cNota,cSerie,cFornec,cLoja,cProduto)

Local aArea		:= GetArea()
Local cFilAux	:= xFilial('CBE')

Default cProduto:= ""

CBE->(DbSetOrder(2))
If !Empty(cProduto)
	If CBE->(DBSeek(cFilAux+cNota+cSerie+cFornec+cLoja+cProduto))
		RecLock("CBE")
		CBE->(DbDelete())
		CBE->(MsUnlock())
	EndIf
Else
	If CBE->(DBSeek(cFilAux+cNota+cSerie+cFornec+cLoja))
		While CBE->(cFilAux+CBE_NOTA+CBE_SERIE+CBE_FORNEC+CBE_LOJA) == cFilAux+cNota+cSerie+cFornec+cLoja
			RecLock("CBE")
			CBE->(DbDelete())
			CBE->(MsUnlock())
			CBE->(dbSkip())
		EndDo
	EndIf
EndIf

RestArea(aArea)
Return

// -------------------------------------------------------------------------------------
/*/{Protheus.doc} CBVSUBNSER
Valida a troca do numero de serie selecionado pelo sistema na libera��o do PV;
 pelo numero de serie lido pelo operador no ato da separacao

@param: Nil
@author: Aecio Ferreira Gomes
@since: 25/09/2013
@Obs: VldNumSer
/*/
// -------------------------------------------------------------------------------------
Function CBVSUBNSER(cProduto,cLocal,cNSerLido,cNSerSug)
Local lRet 		:= .T.
Local aSvSC6	:= SC6->(GetArea())
Local aSvSC9	:= SC9->(GetArea())
Local aSvCB9	:= CB9->(GetArea())
Local cAlias1	:= "TMPNSSUG"
Local cAlias2	:= "TMPNSLIDO"
Local cSubNSer	:= SuperGetMV("MV_SUBNSER",.F.,'1')

// Caso exista, o alias � fechado para que seja recriado com o novo filtro.
If Select(cAlias1) > 0
	(cAlias1)->(dbCloseArea())
EndIf

If Select(cAlias2) > 0
	(cAlias2)->(dbCloseArea())
EndIf

// Filtra item da tabela SDC referente ao numero de s?rie sugerido pelo sistema
BeginSQL Alias cAlias1

	SELECT
		DC_PEDIDO||DC_ITEM AS PVSUG
		,SDC.R_E_C_N_O_ AS REG
		,SDC.*
	FROM
		%table:SDC% SDC
	WHERE
		SDC.DC_FILIAL = %xFilial:SDC% AND
		SDC.DC_PRODUTO = %Exp:cProduto% AND
		SDC.DC_LOCAL = %Exp:cLocal% AND
		SDC.DC_NUMSERI = %Exp:cNSerSug% AND
		SDC.%notDel%
EndSQL

If (cAlias1)->REG > 0
	If SC6->(dbSeek(xFilial("SC6")+(cAlias1)->PVSUG))
		If !Empty(SC6->C6_NUMSERI)
			VTAlert(STR0068,STR0004,.T.,3000)// ###"N�o � possivel fazer a troca, pois o numero de serie do O.S. nao foi selecionado aleatoriamente pelo sistema"##"Aten��o"
			lRet := .F.
		EndIf
	EndIf
EndIf

If lRet
	// Filtra item da tabela SDC referente ao numero de s?rie lido pelo operador caso exista
	BeginSQL Alias cAlias2

		SELECT
			DC_PEDIDO||DC_ITEM AS PVLIDO
			,SDC.R_E_C_N_O_ AS REG
			,SDC.*
		FROM
			%table:SDC% SDC
		WHERE
			SDC.DC_FILIAL = %xFilial:SDC% AND
			SDC.DC_PRODUTO = %Exp:cProduto% AND
			SDC.DC_LOCAL = %Exp:cLocal% AND
			SDC.DC_NUMSERI = %Exp:cNSerLido% AND
			SDC.DC_ORIGEM = %Exp:'SC6'% AND
			SDC.%notDel%
	EndSQL

	If (cAlias2)->REG > 0

		If (cAlias2)->DC_ORIGEM # "SC6"
			VTAlert(STR0069,STR0004,.T.,3000) //### "A origem da reserva do numero de serie selecionado n?pertece a um pedido de vendas, selecione outro numero de serie"
			lRet := .F.
		EndIf

		If lRet .And. SC6->(dbSeek(xFilial("SC6")+(cAlias2)->PVLIDO))
			If !Empty(SC6->C6_NUMSERI)
				VTAlert(STR0070,STR0004,.T.,3000) //### "O numero de serie selecionado foi reservado manualmente para outro pedido de vendas, selecione outro numero de serie"
				lRet := .F.
			EndIf
		EndIf

		If lRet .And. SC9->(dbSeek(xFilial("SC9")+(cAlias2)->PVLIDO))
			If !Empty(SC9->C9_ORDSEP)
				CB9->(dbSetOrder(9)) //CB9_FILIAL+CB9_ORDSEP+CB9_PROD+CB9_LOCAL+CB9_LCALIZ+CB9_LOTECT+CB9_NUMLOT+CB9_NUMSER
				If CB9->(dbSeek(xFilial("CB9")+SC9->C9_ORDSEP+(cAlias2)->(DC_PRODUTO+DC_LOCAL+DC_LOCALIZ+DC_LOTECTL+DC_NUMLOTE+DC_NUMSERI)))
					VTAlert(STR0071,STR0004,.T.,3000) //### "O numero de serie selecionado pertence a outra ordem de separacao, selecione outro numero de serie"
					lRet := .F.
				EndIf
			EndIf
		EndIf

		If lRet
			If cSubNSer == '3'
				lRet := VTYesNo(STR0071,STR0004,.T.) //### "O numero de s�rie selecionado esta reservado para outro pedido, Deseja efetuar a troca?"
			EndIf
		EndIf

	Else
		SBF->(dbSetOrder(4))
		If !SBF->(dbSeek(xFilial("SBF")+cProduto+cNSerLido)) .Or.;
			(SBF->(dbSetOrder(1)), SBF->(SaldoSBF(BF_LOCAL,BF_LOCALIZ,BF_PRODUTO,BF_NUMSERI,BF_LOTECTL,BF_NUMLOTE) ) < 1)
			VTAlert(STR0073,STR0004,.t.,3000) //### "Nao existe saldo em estoque para o numero de serie selecionado"
			lRet := .F.
		ElseIf !(Empty(SBF->BF_EMPENHO)) .Or. !Empty(SBF->BF_QEMPPRE)
			VTAlert(STR0073,STR0004,.t.,3000) //### "Nao existe saldo em estoque para o numero de serie selecionado"
		Endif
	EndIf

EndIf

// Fecha os alias temporarios que contem os registros de troca do numero de serie
If !lRet
	If Select(cAlias1) > 0
		(cAlias1)->(dbCloseArea())
	EndIf

	If Select(cAlias2) > 0
		(cAlias2)->(dbCloseArea())
	EndIf
EndIf

RestArea(aSvSC6)
RestArea(aSvSC9)
RestArea(aSvCB9)
Return(lRet)

// -------------------------------------------------------------------------------------
/*/{Protheus.doc} CBMULTDOC
Cria uma Tela de escolhe de documento quando na base do cliente tem o mesmo documento e a mesma
serie

@param: Nil
@author:Flavio Lopes Rasta
@since: 25/09/2013

/*/
// -------------------------------------------------------------------------------------
Function CBMULTDOC(cAlias,cNota,cSerie)
Local aSave       	:= VTSave()
Local aArea			:=	GetArea()
Local aCab        	:=	{"Documento","Serie","Dt. Emiss�o"}
Local aItens      	:= {}
Local aSerAux			:= {}
Local aSizes    		:= {10,3,10}

VTClear
If cAlias == "SF1"
	SF1->(DbSetOrder(1))
	SF1->(DbSeek(xFilial("SF1")+cNota))

	While !SF1->(Eof()) .And. SF1->F1_DOC == cNota
		If SubStr(SF1->F1_SERIE,1,3) == cSerie
			aAdd(aItens,{SF1->F1_DOC,SerieNfId("SF1",2,"F1_SERIE"),SF1->F1_EMISSAO})
			aAdd(aSerAux,SF1->F1_SERIE)
		EndIf
		SF1->(DbSkip())
	End
ElseIf cAlias == "SF2"
	SF2->(DbSetOrder(1))
	SF2->(DbSeek(xFilial("SF2")+cNota))

	While !SF2->(Eof()) .And. SF2->F2_DOC == cNota
		If SubStr(SF2->F2_SERIE,1,3) == cSerie
			aAdd(aItens,{SF2->F2_DOC,SerieNfId("SF2",2,"F2_SERIE"),SF2->F2_EMISSAO})
			aAdd(aSerAux,SF2->F2_SERIE)
		EndIf
		SF2->(DbSkip())
	End
EndIf

nPos := Len(aItens)
If npos == 1
	cSerie:= aSerAux[nPos]
ElseIf npos > 1
	nPos := VTaBrowse(0,0,7,15,aCab,aItens,aSizes,,nPos)
	cSerie:= aSerAux[nPos]
EndIf

Restarea(aArea)
VTRestore(,,,,aSave)

Return .T.

// -------------------------------------------------------------------------------------
/*/{Protheus.doc} ACDCB9Ser
Valida o numero de serie informado na ordem de separa��o e atualiza variavel de ajuste de saldo
da SDC (Tabela de Composicao de empenho)
@param:
@author:Andre Maximo
@since: 03/06/2016
/*/
// -------------------------------------------------------------------------------------

Function ACDCB9Ser(cProd,cArm,nQtdLib,nQtdLib2,cLOTECTL,cNUMLOTE,cLOCALIZ,cNUMSERI,lUsaVenc,dDataBase,cORDSEP,cPed,aTravas)

Local aArea    	:= GetArea()
Local aSaldos    	:= {}
Local aSalTotal		:= {}
Local cFilialCB9 	:= xFilial("CB9")
Default nQtdLib  	:= 0
Default nQtdLib2 	:= 0
Default cProd		:= " "
Default cArm		:= " "
Default cLOTECTL	:= " "
Default cNUMLOTE	:= " "
Default cLOCALIZ	:= " "
Default cNUMSERI 	:= " "
Default cORDSEP	:= " "
Default lUsaVenc	:= .F.


dbSelectArea("CB9")
CB9->(DbSetOrder(1))
If CB9->(DBSeek(cFilialCB9+cORDSEP))
	While CB9->(! Eof() .and. CB9_FILIAL+CB9_ORDSEP ==;
									cFilialCB9+cORDSEP)
		If CB9->CB9_PEDIDO == cPed .And. CB9->CB9_PROD == cProd .And. CB9->CB9_ITESEP == SC9->C9_ITEM
			cNUMSERI:= CB9->CB9_NUMSER
			aSaldos:= SldPorLote(CB9->CB9_PROD,CB9->CB9_LOCAL,CB9->CB9_QTESEP,nQtdLib2,CB9->CB9_LOTECT,CB9->CB9_NUMLOT,CB9->CB9_LCALIZ,CB9->CB9_NUMSER,NIL,NIL,NIL,lUsaVenc,nil,nil,dDataBase)
			aEval(aSaldos,{|x|aAdd(aSalTotal,x)})
		EndIf
		CB9->(DbSkip())
	EndDo
Else
	aSalTotal:= SldPorLote(cProd,cArm,nQtdLib,nQtdLib2,cLOTECTL,cNUMLOTE,cLOCALIZ,cNUMSERI,aTravas,NIL,NIL,lUsaVenc,nil,nil,dDataBase)
EndIf

RestArea(aArea)
Return aSalTotal
/*/{Protheus.doc} CBGetTamEtq
	(long_description)
	@type  Function
	@author Paulo V. Beraldo
	@since Jun/2019
	@version 1.00
    @param param, param_type, param_descr
	@return nRet, Numeric, Tamanho do Campo C�digo de Etiqueta
	@example
	(examples)
	@see (links_or_references)
/*/
Function CBGetTamEtq()
Local nRet		:= 0
Local nTamEtq	:= 48
Local lCBTamEtq	:= ExistBlock( 'CBTAMETQ' )
If lCBTamEtq
    nRet	:= ExecBlock( 'CBTAMETQ', .F., .F.,{ nTamEtq } )
	nRet	:= IIf( ValType( nRet ) <> 'N', nTamEtq, nRet )
Else
	nRet	:= nTamEtq
EndIf
Return nRet
/*/{Protheus.doc} AcdGTamETQ
	(long_description)
	@type  Function
	@author Paulo V. Beraldo
	@since Jun/2019
	@version 1.00
    @param param, param_type, param_descr
	@return cRet, String, Spaco com o Tamanho do Campo C�digo de Etiqueta
	@example
	(examples)
	@see (links_or_references)
/*/
Function AcdGTamETQ()
Return Space( CBGetTamEtq() )

/*/{Protheus.doc} SumQtdEti( cOrdSep, cProduto )
	Funcao Responsavel por Somar a Quantidade Total de Etiquetas de um mesmo Produto Vinculadas a uma Ordem de Separacao
	@type  Static Function
	@author Paulo V. Beraldo
	@since Jul/2020
	@version 1.00
	@param cOrdSep	, Caracter, Codigo da Ordem de Separacao
	@param cProduto	, Caracter, Codigo do Produto
	@return nRet	, Integer , Quantidade Total de Etiquetas Impressas para um Produto Vinculado a uma Ordem de Separacao
	@example
	nQuant := SumQtdEti( '000010', '0000001010' )
/*/
Static Function SumQtdEti( cOrdSep, cProduto )
Local nRet		:= 0
Local nQtdEti	:= 0
Local nQtdEmb	:= 0
Local aArea		:= GetArea()
Local cQuery	:= Nil
Local cTmpAlias	:= GetNextAlias()

cQuery := " SELECT DISTINCT CB0.CB0_CODETI, CB0.CB0_CODPRO, CB0.CB0_QTDE, SUM( CB9.CB9_QTEEMB ) CB9_QTEEMB "
cQuery += " FROM "+ RetSQLName( 'CB0' ) +" CB0 "
cQuery += " 	INNER JOIN "+ RetSQLName( 'CB9' ) +" CB9 ON ( CB9.CB9_FILIAL = '"+ FWxFilial( 'CB0' ) +"' "
cQuery += " 		AND CB0.CB0_FILIAL = '"+ FWxFilial( 'CB9' ) +"'  
cQuery += " 		AND CB0.CB0_CODPRO = CB9.CB9_PROD 
cQuery += " 		AND CB0.CB0_CODETI = CB9.CB9_CODETI )
cQuery += " WHERE CB9.CB9_ORDSEP = '"+ cOrdSep +"' "
cQuery += " 	AND CB9.CB9_PROD = '"+ cProduto +"' "
cQuery += " 	AND CB9.D_E_L_E_T_ = ' ' "
cQuery += " 	AND CB0.D_E_L_E_T_ = ' ' "
cQuery += " GROUP BY CB0.CB0_CODPRO, CB0.CB0_CODETI, CB0.CB0_QTDE "
cQuery += " ORDER BY 1, 2 "

cQuery := ChangeQuery( cQuery )
dbUseArea( .T., __cRdd, TcGenQry( Nil, Nil, cQuery ), cTmpAlias, .T., .F.  )

dbSelectArea( cTmpAlias )
( cTmpAlias )->( dbEval( { || ( nQtdEti += ( cTmpAlias )->CB0_QTDE, nQtdEmb += ( cTmpAlias )->CB9_QTEEMB ) },,{ || !( cTmpAlias )->( Eof() ) } ) )
nRet := nQtdEti - nQtdEmb

IIf( Select( cTmpAlias ) > 0 , ( cTmpAlias )->( dbCloseArea() ), Nil )
RestArea( aArea )
Return nRet
