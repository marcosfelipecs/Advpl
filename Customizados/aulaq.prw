#include "protheus.ch"
/*
+===========================================+
| Programa: Cálculo do Fatorial |
| Autor : Microsiga Software S.A. |
| Data : 02 de outubro de 2001 |
+===========================================+
*/
User Function CalcFator()
Local nCnt
Local nResultado := 1 // Resultado do fatorial
Local nFator := 5 // Número para o cálculo
// Cálculo do fatorial
For nCnt := nFator To 1 Step -1
nResultado *= nCnt
Next nCnt
// Exibe o resultado na tela, através da função alert
MsgAlert("O fatorial de " + cValToChar(nFator) + ;
 " é " + cValToChar(nResultado))
// Termina o programa
Return( NIL )
