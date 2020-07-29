#include "topconn.ch"
#include "protheus.ch"                                                     

//ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
//±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
//±±ºPrograma  | zExpExcel  º  Autor ³ Marcos Felipe  º  Data ³  28/07/20   º±±
//±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
//±±ºDescricao ³ Rotina dados da SF2/SD2 para o Excel                       º±±
//±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
//±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
//ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß

User Function zExpExcel()

Local cAlias := GetNextAlias() 

Private aCabec := {} //ARRAY DO CABEÇALHO
Private aDados := {} //ARRAY QUE ARMAZENARÁ OS DADOS

//COMEÇO A MINHA CONSULTA SQL
BeginSql Alias cAlias
		SELECT B1_COD, B1_DESC, B2_LOCAL, B2_QATU FROM  %table:SB2990%
		INNER JOIN  %table:SB1990% ON SB2990.B2_COD = SB1990.B1_COD

EndSql //FINALIZO A MINHA QUERY

//CABEÇALHO

aCabec := {"CODIGO","DESCRICAO","LOCAL","QUANTIDADE ATUAL"}

While !(cAlias)->(Eof())

	aAdd(aDados,{B1_COD,B1_DESC, B2_LOCAL, B2_QATU})
	
	(cAlias)->(dbSkip()) //PASSAR PARA O PRÓXIMO REGISTRO                                     
enddo

//JOGO TODO CONTEÚDO DO ARRAY PARA O EXCEL
MsgRun("Favor Aguardar.....", "Exportando os Registros para o Excel",;
{||DlgtoExcel({{"ARRAY","Relat. Saldo de produtos", aCabec, aDados}})})
	                                          
(cAlias)->(dbClosearea())	

return


