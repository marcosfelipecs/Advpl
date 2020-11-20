#Include 'Protheus.ch'
#include "FWMVCDEF.CH"
#include "topconn.ch"
#include "totvs.ch"

User Function modalTest()

Local oModal
Local oContainer

pDataDe  := ctod("  /  /  ")
pDataAte  := ctod("  /  /  ")
 
    oModal  := FWDialogModal():New()       
    oModal:SetEscClose(.T.)
    oModal:setTitle("Exportação de NFS-e ")
     
    //Seta a largura e altura da janela em pixel
    oModal:setSize(140, 170)
 
    oModal:createDialog()
    oModal:addCloseButton(nil, "Fechar")
    oContainer := TPanel():New( ,,, oModal:getPanelMain() )
    oContainer:Align := CONTROL_ALIGN_ALLCLIENT
     
   // TSay():New(1,1,{|| "Teste "},oContainer,,,,,,.T.,,,30,20,,,,,,.T.)

    TSay():New(oContainer, {||  'Emissao Inicial' },050,05,,,,,,.T.,,,200,20)
    //oGetDataDe		:= TGet():New( 058, 005, { | u | If( PCount() == 0, dEmissDe , dEmissDe   := u ) },oDlg, 060, 010, "@D",, 0, 16777215,,.F.,,.T.,,.F.,,.F.,.F.,,.F.,.F. ,,"dEmissDe",,,,.T.  )
	
	//dEmissAte		:= dDataBase
    //oSay  		:= TSay():create(oModal, {||  'Emissao Final' },050,075,,,,,,.T.,,,200,20)
	//oGetDataAte		:= TGet():New( 058, 075, { | u | If( PCount() == 0, dEmissAte , dEmissAte   := u ) },oModal, 060, 010, "@D",, 0, 16777215,,.F.,,.T.,,.F.,,.F.,.F.,,.F.,.F. ,,"dEmissAte",,,,.T.  )

    //oModal:addButtons({{"", "Exportar"  , {|| lConfirma := .T. , oModal:Deactivate() }, "Clique aqui para exportar",,.T.,.T.}})
    
    oModal:Activate()

Return
