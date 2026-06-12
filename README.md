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
_Explicação do Sistema de Análise de Crédito

Este código implementa um sistema completo de análise de crédito baseada em ficheiros Excel, aplicando conceitos de programação orientada a objetos (POO) como encapsulamento, herança, polimorfismo e composição.
1. Configuração Global

A classe ConfiguracaoCredito define regras e utilidades globais:

    LIMITE_APROVADO = 0.3 → abaixo disto, crédito aprovado

    LIMITE_ANALISE = 0.6 → entre 0.3 e 0.6, vai para análise

    Método converter_numero() → garante que valores inválidos viram 0.0

2. Classe Cliente (Base)

Representa um cliente normal.
Principais responsabilidades:

    Guarda:

        Nome

        Rendimento

        Dívida

    Calcula automaticamente:

        Rácio dívida/rendimento

        Decisão de crédito

Lógica:

    raˊcio=dividarendimentoraˊcio=rendimentodivida​

    Classificação:

        < 0.3 → Aprovado

        < 0.6 → Análise

        >= 0.6 → Recusado

Inclui métodos getters e __repr__ para debug.
3. ClientePrioritario (Herança + Polimorfismo)

Extende Cliente, mas com vantagens:

    Tem um bónus de +0.1 nos limites

    Override do método classificar()

Exemplo:

    Cliente normal: limite aprovação = 0.3

    Cliente VIP: limite aprovação = 0.4

Ou seja, clientes VIP têm maior probabilidade de aprovação.
4. Carteira (Composição)

Gerencia vários clientes.
Funcionalidades:

    Adicionar clientes

    Listar clientes

    Gerar estatísticas:

        Média, máximo, mínimo de rácios

        Ranking por nível de dívida

        Contagem de aprovados/recusados

        Separação entre VIPs e normais

Extra:

    somar_dividas() usa recursividade para somar todos os rácios

5. AnalisadorExcel

Responsável por ler e escrever dados no Excel.
Fluxo:

    Abre ficheiro (openpyxl)

    Lê linhas:

        Nome

        Rendimento

        Dívida

        Tipo (Normal ou VIP)

    Cria objetos Cliente ou ClientePrioritario

    Guarda resultados:

        Coluna D → Rácio

        Coluna E → Decisão

    Guarda novo ficheiro

6. Função Main

Coordena tudo:

    Processa o Excel

    Guarda resultados

    Mostra estatísticas da carteira

    Demonstra polimorfismo:

        Mesmo input gera decisões diferentes (normal vs VIP)

    Calcula soma total dos rácios

Exemplo prático

Se um cliente tiver:

    Rendimento = 3000

    Dívida = 1500

Então:

    raˊcio=1500/3000=0.5raˊcio=1500/3000=0.5

Resultado:

    Cliente normal → Análise

    Cliente VIP → Aprovado (devido ao aumento dos limites)

Conclusão

O código demonstra:

    Separação clara de responsabilidades

    Reutilização com herança

    Polimorfismo (mesma função com comportamentos diferentes)

    Integração com Excel

    Estrutura semelhante a sistemas reais de negócio
