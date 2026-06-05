from openpyxl import load_workbook


# | Requisito              | Onde está                                                                                                                         |
# | ---------------------- | --------------------------------------------------------------------------------------------------------------------------------- |
# | Múltiplas funções      | carregar_dados, converter_numero, calcular_racio, classificar_cliente, processar_credito, gerar_estatisticas, somar_dividas, main |
# | Variáveis globais      | LIMITE_APROVADO, LIMITE_ANALISE                                                                                                   |
# | Variáveis locais       | rendimento, divida, racio, decisao dentro de processar_credito                                                                    |
# | Passagem de argumentos | calcular_racio(divida, rendimento), classificar_cliente(racio, lim_aprovado, lim_analise)                                         |
# | Valores retornados     | converter_numero, calcular_racio, classificar_cliente devolvem valores usados noutras funções                                     |
# | Função recursiva       | somar_dividas soma todas as dívidas recursivamente                                                                                |
# | Lambda                 | sorted(..., key=lambda x: x[1]) e filter(lambda x: ...)                                                                           |
# | Built-ins              | sum(), max(), min(), len(), sorted(), filter(), enumerate(), round()                                                              |
# | Leitura/escrita Excel  | load_workbook, wb.save()                                                                                                          |


# ─────────────────────────────────────────
# VARIÁVEIS GLOBAIS — limites de decisão
# Justificação: são valores de referência usados
# por várias funções, por isso são globais.
# ─────────────────────────────────────────
LIMITE_APROVADO = 0.3
LIMITE_ANALISE  = 0.6


# ─── 1. Carregar dados do Excel ───────────
def carregar_dados(caminho):
    """Abre o ficheiro Excel e devolve workbook e worksheet."""
    wb = load_workbook(caminho)
    ws = wb["Sheet1"]
    return wb, ws


# ─── 2. Converter valor para float (retorna valor) ───
def converter_numero(valor):
    """Converte um valor para float. Devolve 0.0 se inválido."""
    try:
        return float(valor)
    except (TypeError, ValueError):
        return 0.0


# ─── 3. Calcular rácio (retorna valor) ───────
def calcular_racio(divida, rendimento):
    """Calcula e devolve o rácio dívida/rendimento."""
    if rendimento > 0:
        return round(divida / rendimento, 2)
    return 1.0


# ─── 4. Classificar cliente ───────────────
def classificar_cliente(racio, lim_aprovado=LIMITE_APROVADO, lim_analise=LIMITE_ANALISE):
    """Devolve a decisão de crédito com base no rácio."""
    if racio < lim_aprovado:
        return "Aprovado"
    elif racio < lim_analise:
        return "Análise"
    else:
        return "Recusado"


# ─── 5. Processar todas as linhas ─────────
def processar_credito(ws):
    """Percorre as linhas e escreve rácio e decisão no Excel."""
    ws["D1"] = "Rácio Dívida"
    ws["E1"] = "Decisão"

    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        # variáveis locais — só existem dentro deste ciclo
        rendimento = converter_numero(row[1])
        divida     = converter_numero(row[2])

        racio    = calcular_racio(divida, rendimento)
        decisao  = classificar_cliente(racio)

        ws[f"D{i}"] = racio
        ws[f"E{i}"] = decisao


# ─── 6. Gerar estatísticas (built-ins) ────
def gerar_estatisticas(ws):
    """Calcula e imprime estatísticas gerais dos clientes."""
    racios = [
        calcular_racio(converter_numero(ws[f"C{i}"].value), converter_numero(ws[f"B{i}"].value))
        for i in range(2, ws.max_row + 1)
    ]

    print("\n Estatísticas:")
    print(f"  Total de clientes  : {len(racios)}")
    print(f"  Rácio médio        : {round(sum(racios) / len(racios), 2)}")
    print(f"  Rácio máximo       : {max(racios)}")
    print(f"  Rácio mínimo       : {min(racios)}")

    # lambda — ordenar clientes do menos endividado para o mais
    clientes = [
        (ws[f"A{i}"].value, calcular_racio(
            converter_numero(ws[f"C{i}"].value),
            converter_numero(ws[f"B{i}"].value)
        ))
        for i in range(2, ws.max_row + 1)
    ]
    ordenados = sorted(clientes, key=lambda x: x[1])

    print("\nRanking (menos endividado → mais endividado):")
    for pos, (nome, racio) in enumerate(ordenados, start=1):
        print(f"  {pos}º {nome} — rácio: {racio}")

    # filter() — listar apenas os recusados
    recusados = list(filter(lambda x: x[1] >= LIMITE_ANALISE, clientes))
    print(f"\n Clientes recusados: {len(recusados)}")
    aprovados = list(filter(lambda x: x[1] < LIMITE_APROVADO, clientes))
    print(f" Clientes aprovados: {len(aprovados)}")


# ─── 7. Função recursiva ──────────────────
def somar_dividas(ws, index, total=0):
    """Soma recursivamente todas as dívidas dos clientes."""
    if index > ws.max_row:
        return total
    divida = converter_numero(ws[f"C{index}"].value)
    return somar_dividas(ws, index + 1, total + divida)


# ─── 8. Função principal ──────────────────
def main(caminho="credito.xlsx"):
    wb, ws = carregar_dados(caminho)

    processar_credito(ws)
    wb.save("credito_resultado.xlsx")
    print("Ficheiro guardado como credito_resultado.xlsx")

    gerar_estatisticas(ws)

    total_dividas = somar_dividas(ws, 2)
    print(f"\n Total de dívidas em carteira: {total_dividas}€")


main()