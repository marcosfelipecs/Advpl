#INCLUDE "TOPCONN.CH"
#INCLUDE "PROTHEUS.CH"                                                      

//Protheus-Construindo um relatório ADVPL em 5 minutos
User Function SQLEXCEL()
Local cAlias := GetNextAlias() //Declarei meu ALIAS

Private aCabec := {} //ARRAY DO CABEÇALHO
Private aDados := {} //ARRAY QUE ARMAZENARÁ OS DADOS

//COMEÇO A MINHA CONSULTA SQL
BeginSql Alias cAlias
		SELECT B1_COD, B1_DESC, B2_LOCAL, B2_QATU FROM  %table:SB2%  SB2
		INNER JOIN  %table:SB1%  SB1 ON SB2.B2_COD = B1_COD AND B2_FILIAL = B1_FILIAL AND SB1.%notdel%  
		WHERE SB2.%notdel%  

EndSql //FINALIZO A MINHA QUERY
	
//CABEÇALHO
aCabec := {"CODIGO","DESCRICAO","LOCAL","QUANTIDADE ATUAL"}

While !(cAlias)->(Eof())

	aAdd(aDados,{B1_COD,B1_DESC, B2_LOCAL, B2_QATU})
	
	(cAlias)->(dbSkip()) //PASSAR PARA O PRÓXIMO REGISTRO                                     
enddo

//JOGO TODO CONTEÚDO DO ARRAY PARA O EXCEL
DlgtoExcel({{"ARRAY","Relatório saldo físico de produtos por armazem", aCabec, aDados}})
	                                          
(cAlias)->(dbClosearea())	

return


