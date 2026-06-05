from openpyxl import load_workbook

# | Requisito              | Onde está                                                                                                                |
# | ---------------------- | ------------------------------------------------------------------------------------------------------------------------ |
# | Múltiplas funções      | carregar_dados, calcular_percentagem, classificar_aluno, processar_presencas, gerar_estatisticas, contar_admitidos, main |
# | Variável global        | LIMITE_PRESENCA = 75.0                                                                                                   |
# | Variáveis locais       | percentagem, situacao dentro de processar_presencas                                                                      |
# | Passagem de argumentos | calcular_percentagem(presencas, total), classificar_aluno(percentagem, limite)                                           |
# | Valores retornados     | calcular_percentagem e classificar_aluno devolvem valores usados noutras funções                                         |
# | Função recursiva       | contar_admitidos conta com recursividade                                                                                 |
# | Lambda                 | sorted(..., key=lambda x: x[1]) e filter(lambda p: ...)                                                                  |
# | Built-ins              | sum(), max(), min(), len(), sorted(), filter(), enumerate()                                                              |
# | Leitura/escrita Excel  | load_workbook, wb.save()                                                                                                 |


# ─────────────────────────────────────────
# VARIÁVEL GLOBAL — limite mínimo de presença
# Justificação: é um valor de referência usado
# por várias funções, por isso é global.
# ─────────────────────────────────────────
LIMITE_PRESENCA = 75.0


# ─── 1. Carregar dados do Excel ───────────
def carregar_dados(caminho):
    """Lê o ficheiro Excel e devolve a worksheet."""
    wb = load_workbook(caminho)
    ws = wb["Sheet1"]
    return wb, ws


# ─── 2. Calcular percentagem (retorna valor) ──
def calcular_percentagem(presencas, total):
    """Calcula e devolve a percentagem de presenças."""
    if total and total > 0:
        return round((presencas / total) * 100, 1)
    return 0.0


# ─── 3. Classificar aluno (retorna valor) ────
def classificar_aluno(percentagem, limite=LIMITE_PRESENCA):
    """Devolve 'Admitido' ou 'Excluído' com base no limite."""
    return "Admitido" if percentagem >= limite else "Excluído"


# ─── 4. Processar todas as linhas ────────────
def processar_presencas(ws):
    """Percorre as linhas e escreve percentagem e situação."""
    ws["D1"] = "Percentagem"
    ws["E1"] = "Situação"

    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        aluno, presencas, total = row[0], row[1], row[2]

        # variáveis locais — só existem dentro deste ciclo
        percentagem = calcular_percentagem(presencas, total)
        situacao = classificar_aluno(percentagem)

        ws[f"D{i}"] = percentagem
        ws[f"E{i}"] = situacao


# ─── 5. Gerar estatísticas (built-ins) ───────
def gerar_estatisticas(ws):
    """Usa built-ins para calcular estatísticas gerais."""
    percentagens = [
        calcular_percentagem(ws[f"B{i}"].value, ws[f"C{i}"].value)
        for i in range(2, ws.max_row + 1)
    ]

    print("\n Estatísticas:")
    print(f"  Média de presenças : {round(sum(percentagens) / len(percentagens), 1)}%")
    print(f"  Máximo             : {max(percentagens)}%")
    print(f"  Mínimo             : {min(percentagens)}%")
    print(f"  Total de alunos    : {len(percentagens)}")

    # lambda — ordenar alunos por percentagem (do maior para o menor)
    alunos = [
        (ws[f"A{i}"].value, calcular_percentagem(ws[f"B{i}"].value, ws[f"C{i}"].value))
        for i in range(2, ws.max_row + 1)
    ]
    alunos_ordenados = sorted(alunos, key=lambda x: x[1], reverse=True)

    print("\n Ranking de presenças:")
    for pos, (nome, perc) in enumerate(alunos_ordenados, start=1):
        print(f"  {pos}º {nome} — {perc}%")


# ─── 6. Contar admitidos (recursiva) ─────────
def contar_admitidos(lista, index=0):
    """Conta recursivamente quantos alunos estão admitidos."""
    if index == len(lista):
        return 0
    atual = 1 if lista[index] >= LIMITE_PRESENCA else 0
    return atual + contar_admitidos(lista, index + 1)


# ─── 7. Função principal ──────────────────────
def main(caminho="alunos.xlsx"):
    wb, ws = carregar_dados(caminho)

    processar_presencas(ws)
    wb.save("alunos_resultado.xlsx") 
    print("Ficheiro atualizado com sucesso.")

    gerar_estatisticas(ws)

    percentagens = [
        calcular_percentagem(ws[f"B{i}"].value, ws[f"C{i}"].value)
        for i in range(2, ws.max_row + 1)
    ]
    excluidos = list(filter(lambda p: p < LIMITE_PRESENCA, percentagens))
    print(f"\n Alunos excluídos: {len(excluidos)}")
    print(f" Alunos admitidos (recursivo): {contar_admitidos(percentagens)}")


main()