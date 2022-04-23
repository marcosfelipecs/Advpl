#Include 'Protheus.ch'
#include "topconn.ch"
#include "totvs.ch"

User Function TestPergunt()

Local aPergs   := {}
Local dDataDe  := FirstDate(Date())
Local dDataAt  := LastDate(Date())
 
aAdd(aPergs, {1, "Emissão De",  dDataDe,  "", ".T.", "", ".T.", 80,  .F.})
aAdd(aPergs, {1, "Emissão Até", dDataAt,  "", ".T.", "", ".T.", 80,  .T.})
 
ParamBox(aPergs, "Exportar NFS-e")

//----------------------------------------------------------------------------------------------------------------//
// Exportacao de dados para o Excel.
//----------------------------------------------------------------------------------------------------------------//

static function TExcel3()

Local cAlias := GetNextAlias() 
Private _aCabec := {}
Private _aDados := {}

_aCabec := {"Filial","Número Pedido","Produto","Valor Total","Emissão"}           

//Consulta SQL
 BeginSql Alias cAlias
		SELECT C7_FILIAL, C7_NUM, C7_DESCRI, C7_TOTAL, C7_EMISSAO
		FROM  %table:SC7990%
		WHERE SC7990.%NotDel%
		AND C7_EMISSAO >= %Exp:DTOS(dDataDe)% 
		AND C7_EMISSAO <= %Exp:DTOS(dDataAt)% 
		//INNER JOIN  %table:SB1990% ON SB2990.B2_COD = SB1990.B1_COD
EndSql 

While !(cAlias)->(Eof())

	aAdd(_aDados,{C7_FILIAL,C7_NUM,C7_DESCRI,C7_TOTAL,C7_EMISSAO})
	
	(cAlias)->(dbSkip()) //PASSAR PARA O PRÓXIMO REGISTRO                                     
end

//JOGO TODO CONTEÚDO DO ARRAY PARA O EXCEL
MsgRun("Favor Aguardar.....", "Exportando os Registros para o Excel",;
{||DlgtoExcel({{"ARRAY","Relatório Notas Fiscais", _aCabec, _aDados}})})
	                                          
(cAlias)->(dbClosearea())	
close(odlg1)

return
