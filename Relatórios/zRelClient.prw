#include "protheus.ch"

User Function tReport1()
   //Informando o vetor com as ordens utilizadas pelo relatório
   MPReport("MPREPORT1","SA1","Relacao de Clientes","Este relatório irá imprimir a relacao de clientes",{"Por Codigo","Alfabetica","Por "+RTrim(RetTitle("A1_CGC"))})
Return
 
User Function tReport2()
   //Informando para função carregar os índices do Dicionário de Índices (SIX) da tabela
   MPReport("MPREPORT2","SA1","Relacao de Clientes","Este relatório irá imprimir a relacao de clientes",,.T.)
Return
