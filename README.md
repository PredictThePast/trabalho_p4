(1)alunos.py: automatiza a verificação de presenças dos alunos, atraves das faltas por aluno e aulas totais.\
Maior ou igual a 75%:	Presença suficiente:	Admitido\
Menor que 75%:	Demasiadas faltas:	Excluído

upgr_alunos.py: codigo melhorado para aparte 2 do projeto/
| Requisito              | Onde está                                                                                                                |
| ---------------------- | ------------------------------------------------------------------------------------------------------------------------ |
| Múltiplas funções      | carregar_dados, calcular_percentagem, classificar_aluno, processar_presencas, gerar_estatisticas, contar_admitidos, main |
| Variável global        | LIMITE_PRESENCA = 75.0                                                                                                   |
| Variáveis locais       | percentagem, situacao dentro de processar_presencas                                                                      |
| Passagem de argumentos | calcular_percentagem(presencas, total), classificar_aluno(percentagem, limite)                                           |
| Valores retornados     | calcular_percentagem e classificar_aluno devolvem valores usados noutras funções                                         |
| Função recursiva       | contar_admitidos conta com recursividade                                                                                 |
| Lambda                 | sorted(..., key=lambda x: x[1]) e filter(lambda p: ...)                                                                  |
| Built-ins              | sum(), max(), min(), len(), sorted(), filter(), enumerate()                                                              |
| Leitura/escrita Excel  | load_workbook, wb.save()                                                                                                 |


(2)credito.py: automatiza pedidos de credito, atraves da divida e do rendimento.\
Menor que 0.3:	Dívida baixa face ao rendimento:	Aprovado\
Entre 0.3 e 0.6:	Situação intermédia:	Análise\
Maior que 0.6:	Dívida muito alta:	Recusado

upgr_credito.py:/
| Requisito              | Onde está                                                                                                                         |
| ---------------------- | --------------------------------------------------------------------------------------------------------------------------------- |
| Múltiplas funções      | carregar_dados, converter_numero, calcular_racio, classificar_cliente, processar_credito, gerar_estatisticas, somar_dividas, main |
| Variáveis globais      | LIMITE_APROVADO, LIMITE_ANALISE                                                                                                   |
| Variáveis locais       | rendimento, divida, racio, decisao dentro de processar_credito                                                                    |
| Passagem de argumentos | calcular_racio(divida, rendimento), classificar_cliente(racio, lim_aprovado, lim_analise)                                         |
| Valores retornados     | converter_numero, calcular_racio, classificar_cliente devolvem valores usados noutras funções                                     |
| Função recursiva       | somar_dividas soma todas as dívidas recursivamente                                                                                |
| Lambda                 | sorted(..., key=lambda x: x[1]) e filter(lambda x: ...)                                                                           |
| Built-ins              | sum(), max(), min(), len(), sorted(), filter(), enumerate(), round()                                                              |
| Leitura/escrita Excel  | load_workbook, wb.save()                                                                                                          |

_resultado.xlsx: excel resultante dos codigos melhorados

