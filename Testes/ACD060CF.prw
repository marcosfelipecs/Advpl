#include 'totvs.ch'
#include 'topconn.ch'


/*/{Protheus.doc} ACD060CF
Ponto de entrada para corrigir problema de endereçamento pelo coletor quando a nota de entrada tem o mesmo produto e lote repetidamente nas linhas do
SD1, criando assim um SDA para cada linha.

@type function
@version  1.0
@author rt1025
@since 11/24/2023
/*/
User Function ACD060CF()
Local i,x,nPos,aNovosItens
Local aAreaCB0 := CB0->(GetArea())
Local aAreaSD1 := SD1->(GetArea())


aNovosItens := {}

IF LEN(aDist) > 0 .and. aDist[1][4] == 'SD1'
CB0->(DBSETORDER(1))

    For i := 1 to Len(aDist) // Percorrer o array de endereçamentos

        // a posição 5 do array contem as etiquetas que foram lidas.
        // através delas vou verificar se o numseq da posição 1 bate com o numseq gravado na etiqueta.
        x := 1
        While x <= Len(aDist[i][5])
        //For x := 1 to Len(aDist[i][5])
            lFLAG := .F.
            // Posiciona o CB0
            CB0->(DBGOTOP())
            IF CB0->(DBSEEK(XFILIAL("CB0")+aDist[i][5][x][1],.F.))
                // Se encontrou a etiqueta, verifica se o numseq bate
                // PROCURA NOTA FISCAL NO SD1 PARA PEGAR O NUMSEQ
                cNumSeq := POSICIONE('SD1',1,XFILIAL('SD1')+CB0->(CB0_NFENT+CB0_SERIEE+CB0_FORNECE+CB0_LOJAFO+CB0_CODPRO+CB0_ITNFE),"D1_NUMSEQ")
                IF cNumSeq <> aDist[i][1]
                    lFlag := .T.
                    // Se for diferente, temos que tirar a etiqueta do array 5, mas antes guardar os dados em um array auxiliar
                    aaux := aClone(aDist[i][5][x])
                    aDel(aDist[i][5],x) // Deleta a etiqueta
                    aSize(aDist[i][5],len(aDist[i][5])-1) // Diminui o array

                    // Desconta a quantidade da posição 3 do array
                    aDist[i][3] -= aaux[2]
                    
                    // Adiciona novo item no array adist com os dados da etiqueta removida

                    // Verifica se o numseq da etiqueta já existe no array
                    nPos := aScan(aDist,{|d| d[1] == cNumSeq})

                    IF nPos <= 0 // Não encontrou o numseq
                        // Adiciona entrada no aDist

                        nPos := aScan(aNovosItens,{|n| n[1] == cNumSeq})
                        IF nPos <= 0 // Não encontrou o numseq
                            aAdd(aNovosItens,{;
                                cNumSeq,;
                                CB0->CB0_CODPRO,;
                                CB0->CB0_QTDE,;
                                "SD1",;
                                {AAUX},;
                                CB0->CB0_LOTE,;
                                SPACE(TAMSX3("CB0_SLOTE")[2]),;
                                SPACE(TAMSX3("CB0_NUMSER")[2]);
                            })
                        ELSE
                            aNovosItens[nPos][3] += CB0->CB0_QTDE
                            aAdd(aNovosItens[nPos][5],aAux)
                        END IF

                    ELSE // encontrou o numseq

                        aDist[nPos][3] += CB0->CB0_QTDE
                        aAdd(aDist[nPos][5],aAux)

                    END IF

                END IF
            End if
            IF !lFlag
                x++
            End if
        End
        //Next

    Next

    IF Len(aNovosItens) > 0 // SE ARRAY DE NOVOS CONTIVER ITENS

        For i := 1 to len(aNovosItens)
            aAdd(aDist,aNovosItens[i])
        Next

    END IF
END IF

CB0->(RESTAREA(aAreaCB0))
SD1->(RESTAREA(aAreaSD1))
RETURN
