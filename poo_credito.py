from openpyxl import load_workbook

# ═══════════════════════════════════════════════════════════════
#  1. Configuração global (encapsulada)
# ═══════════════════════════════════════════════════════════════
class ConfiguracaoCredito:
    LIMITE_APROVADO = 0.3
    LIMITE_ANALISE  = 0.6

    @staticmethod
    def converter_numero(valor):
        try:
            return float(valor)
        except (TypeError, ValueError):
            return 0.0


# ═══════════════════════════════════════════════════════════════
#  2. Classe base — Cliente
# ═══════════════════════════════════════════════════════════════
class Cliente:
    def __init__(self, nome, rendimento_raw, divida_raw):
        self._nome        = nome
        self.__rendimento = ConfiguracaoCredito.converter_numero(rendimento_raw)
        self.__divida     = ConfiguracaoCredito.converter_numero(divida_raw)
        self._racio       = Cliente.calcular_racio(self.__divida, self.__rendimento)
        self._decisao     = self.classificar()

    def get_nome(self):    return self._nome
    def get_racio(self):   return self._racio
    def get_decisao(self): return self._decisao

    @staticmethod
    def calcular_racio(divida, rendimento):
        if rendimento > 0:
            return round(divida / rendimento, 2)
        return 1.0

    def classificar(self):
        if self._racio < ConfiguracaoCredito.LIMITE_APROVADO:
            return "Aprovado"
        elif self._racio < ConfiguracaoCredito.LIMITE_ANALISE:
            return "Análise"
        return "Recusado"

    def __repr__(self):
        return f"Cliente({self._nome!r}, rácio={self._racio}, decisão={self._decisao!r})"


# ═══════════════════════════════════════════════════════════════
#  3. Herança + Polimorfismo — ClientePrioritario
# ═══════════════════════════════════════════════════════════════
class ClientePrioritario(Cliente):
    BONUS = 0.1

    def __init__(self, nome, rendimento_raw, divida_raw, motivo="VIP"):
        self._motivo = motivo
        super().__init__(nome, rendimento_raw, divida_raw)

    def classificar(self):
        lim_ap = ConfiguracaoCredito.LIMITE_APROVADO + self.BONUS
        lim_an = ConfiguracaoCredito.LIMITE_ANALISE  + self.BONUS
        if self._racio < lim_ap:
            return "Aprovado (Prioritário)"
        elif self._racio < lim_an:
            return "Análise (Prioritário)"
        return "Recusado (Prioritário)"


# ═══════════════════════════════════════════════════════════════
#  4. Composição — Carteira
# ═══════════════════════════════════════════════════════════════
class Carteira:
    def __init__(self):
        self.__clientes = []

    def adicionar(self, cliente):
        self.__clientes.append(cliente)

    def listar(self):
        return list(self.__clientes)

    def gerar_estatisticas(self):
        racios = [c.get_racio() for c in self.__clientes]
        n = len(racios)
        if n == 0:
            print("Carteira vazia.")
            return

        print("\nEstatísticas:")
        print(f"  Total de clientes  : {n}")
        print(f"  Rácio médio        : {round(sum(racios) / n, 2)}")
        print(f"  Rácio máximo       : {max(racios)}")
        print(f"  Rácio mínimo       : {min(racios)}")

        ordenados = sorted(self.__clientes, key=lambda c: c.get_racio())
        print("\nRanking (menos endividado → mais endividado):")
        for pos, c in enumerate(ordenados, start=1):
            print(f"  {pos}º {c.get_nome()} — rácio: {c.get_racio()}")

        recusados = list(filter(lambda c: "Recusado" in c.get_decisao(), self.__clientes))
        aprovados = list(filter(lambda c: "Aprovado" in c.get_decisao(), self.__clientes))
        print(f"\nClientes recusados: {len(recusados)}")
        print(f"Clientes aprovados : {len(aprovados)}")

        # ── Mostrar detalhes VIP ──────────────────
        vips = list(filter(lambda c: isinstance(c, ClientePrioritario), self.__clientes))
        normais = list(filter(lambda c: not isinstance(c, ClientePrioritario), self.__clientes))

        print(f"\nClientes VIP       : {len(vips)}")
        print(f"Clientes normais   : {len(normais)}")

        if vips:
            print("\nDetalhe clientes VIP (limites +0.1 aplicados):")
            for c in vips:
                print(f"  • {c.get_nome()} — rácio: {c.get_racio()} → {c.get_decisao()}")
            print("\nDetalhe clientes normais:")
            for c in normais:
                print(f"  • {c.get_nome()} — rácio: {c.get_racio()} → {c.get_decisao()}")

    def somar_dividas(self, lista=None, total=0.0):
        if lista is None:
            lista = self.__clientes[:]
        if not lista:
            return total
        return self.somar_dividas(lista[1:], total + lista[0].get_racio())


# ═══════════════════════════════════════════════════════════════
#  5. Composição — AnalisadorExcel
# ═══════════════════════════════════════════════════════════════
class AnalisadorExcel:
    def __init__(self, caminho):
        self.__caminho  = caminho
        self.__carteira = Carteira()
        self.__wb = None
        self.__ws = None

    def __carregar(self):
        self.__wb = load_workbook(self.__caminho)
        self.__ws = self.__wb["Sheet1"]

    def processar(self):
        self.__carregar()
        self.__ws["D1"] = "Rácio Dívida"
        self.__ws["E1"] = "Decisão"
        self.__ws["F1"] = "Tipo"          # ← cabeçalho da col F

        for i, row in enumerate(self.__ws.iter_rows(min_row=2, values_only=True), start=2):
            nome       = row[0]
            rendimento = row[1]
            divida     = row[2]
            tipo       = row[5] if len(row) > 5 and row[5] else "Normal"  # col F

            if tipo == "VIP":
                cliente = ClientePrioritario(nome, rendimento, divida, motivo="VIP")
            else:
                cliente = Cliente(nome, rendimento, divida)

            self.__carteira.adicionar(cliente)
            self.__ws[f"D{i}"] = cliente.get_racio()
            self.__ws[f"E{i}"] = cliente.get_decisao()

    def guardar(self, destino="credito_resultado.xlsx"):
        self.__wb.save(destino)
        print(f"Ficheiro guardado como {destino}")

    def get_carteira(self):
        return self.__carteira


# ═══════════════════════════════════════════════════════════════
#  6. Main
# ═══════════════════════════════════════════════════════════════
def main(caminho="credito.xlsx"):
    analisador = AnalisadorExcel(caminho)
    analisador.processar()
    analisador.guardar("credito_resultado.xlsx")

    carteira = analisador.get_carteira()
    carteira.gerar_estatisticas()

    print("\n── Demonstração de Polimorfismo ──")
    c_normal = Cliente("João Silva", 3000, 1500)
    c_vip    = ClientePrioritario("Maria Santos", 3000, 1500, motivo="Funcionário")
    print(f"  Cliente normal → {c_normal.get_decisao()}")
    print(f"  Cliente VIP    → {c_vip.get_decisao()}")

    total = carteira.somar_dividas()
    print(f"\nSoma total dos rácios em carteira: {round(total, 2)}")


main()