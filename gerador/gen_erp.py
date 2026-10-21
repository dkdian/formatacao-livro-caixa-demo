"""Gera relatorios de Livro Caixa FICTICIOS no mesmo layout de largura fixa do ERP.

Saida: demo\\erp_xlsx\\<mes>\\*.xlsx (temporarios) -> convertidos para .xls pelo Excel (to_xls.py)
       demo\\expected.json  (saldos esperados para validar o Power Query)
Tudo inventado: empresa, contas, fornecedores, clientes, valores.
"""
import datetime as dt
import json
import os
import random
import sys

import openpyxl

random.seed(23)
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, "demo")

EMPRESA = "EMPRESA EXEMPLO COMERCIO LTDA"
CNPJ = "00.000.000/0001-00"
SISTEMA = "ERP Ficticio Sistemas"

# codigo, nome no ERP, nome na planilha, desembolso, texto do arquivo, saldo inicial de maio
CONTAS = [
    ("1101010001", "CAIXA", "Dinheiro", "Dinheiro", "Caixa", 18450.32),
    ("1101010002", "CHEQUES DEV", "Cheque DEV", "Cheque", "Cheques DEV", 125800.00),
    ("1101010003", "CHEQUES PRE", "Cheque PRE", "Cheque", "Cheques PRE", 210300.50),
    ("1101020001", "BANCO ALFA 1001-0", "Banco Alfa 1001-0", "C/D em Conta", "Banco Alfa 1001-0", -12480.90),
    ("1101020002", "BANCO BETA 2002-0", "Banco Beta 2002-0", "C/D em Conta", "Banco Beta 2002-0", 35120.15),
]
C_CX, C_DEV, C_PRE, C_ALFA, C_BETA = [c[0] for c in CONTAS]

# codigo contrapartida -> (descricao completa, tipo, categoria)
CP = {
    "3101010001": ("RECEB CLIENTES DUPLICATA", "Receita", "Operacional"),
    "3101010002": ("RECEB CARTOES", "Receita", "Operacional"),
    "3101010003": ("VENDAS A VISTA", "Receita", "Operacional"),
    "3102010001": ("JUROS RECEBIDOS", "Receita", "Receita.Financeira"),
    "4101010001": ("FORNECEDORES DIVERSOS", "Gasto", "Custo.Operacional"),
    "4101010002": ("FRETES E CARRETOS", "Gasto", "Custo.Operacional"),
    "4102010001": ("JUROS PAGOS FORNECEDORES", "Gasto", "Despesa.Financeira"),
    "4102010002": ("TARIFAS BANCARIAS", "Gasto", "Despesa.Financeira"),
    "4102010003": ("IOF", "Gasto", "Despesa.Financeira"),
    "4102010004": ("JUROS PAGOS DESC CHEQUES PRE", "Gasto", "Despesa.Financeira"),
    "4103010001": ("SALARIOS", "Gasto", "Despesa.com.Pessoal"),
    "4103010002": ("ENERGIA ELETRICA", "Gasto", "Despesa.Fixa"),
    "4103010003": ("ALUGUEL", "Gasto", "Despesa.Fixa"),
    "4103010004": ("MATERIAL DE LIMPEZA", "Gasto", "Despesa.Variável"),
    "4104010001": ("PROPAGANDA E PUBLICIDADE", "Gasto", "Marketing.Comercial"),
    "4105010001": ("IMPOSTOS E TAXAS", "Gasto", "Tributo"),
    "5101010007": ("ENTRA BAIXA CHEQUE DEV", "Movimentação", "Entrada"),
    "5101010011": ("ENTRA TRANSF ENTRE CONTAS", "Movimentação", "Entrada"),
    "5101010020": ("SAI TRANSF ENTRE CONTAS", "Movimentação", "Saída"),
    "5101010030": ("ENTRA CAPT CHEQUES PRE", "Movimentação", "Entrada"),
    "5101010031": ("SAI DESCONTO CHEQUES PRE", "Movimentação", "Saída"),
    "5101010040": ("SAI DEVOLUCAO DE CHEQUES", "Movimentação", "Saída"),
    "6101010001": ("RETIRADA DE SOCIOS", "Sócio", "Retirada"),
    # novas (NAO entram no De-Para): demonstram a aba Novas_Contrapartidas
    "4102010009": ("TARIFA PIX", "Gasto", "Despesa.Financeira"),
    "4103010009": ("HONORARIOS CONTABEIS", "Gasto", "Despesa.Fixa"),
}
NOVAS = {"4102010009", "4103010009"}

FORNECEDORES = [
    ("FERRAGENS ORION LTDA", ["FERRAGENS ORION LTDA", "ORION FERRAG LTDA", "FERRAGENS ORION"]),
    ("TINTAS AURORA LTDA", ["TINTAS AURORA LTDA", "AURORA TINTAS E VERNIZES"]),
    ("MADEIRAS CEDRO LTDA", ["MADEIRAS CEDRO LTDA"]),
    ("CIMENTOS ARAUCARIA LTDA", ["CIMENTOS ARAUCARIA LTDA"]),
    ("ELETRICA FAROL LTDA", ["ELETRICA FAROL LTDA"]),
    ("METAIS CARDEAL LTDA", ["METAIS CARDEAL LTDA"]),
]
CLIENTES = ["COMERCIAL PIONEIRA LTDA", "CONSTRUTORA HORIZONTE LTDA", "MARCENARIA ESTRELA LTDA",
            "REFORMAS LIRIO LTDA", "OBRAS VALE VERDE LTDA", "CASA DOS SONHOS REFORMAS"]


def fmt(v):
    """1234.5 -> '1.234,50' (negativo: '1.234,50-')"""
    s = f"{abs(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return s + ("-" if v < 0 else "")


def dias_uteis(ano, mes):
    d = dt.date(ano, mes, 1)
    out = []
    while d.month == mes:
        if d.weekday() < 5:
            out.append(d)
        d += dt.timedelta(days=1)
    return out


class Gerador:
    def __init__(self, ano, mes, saldos_ini):
        self.ano, self.mes = ano, mes
        self.dias = dias_uteis(ano, mes)
        self.saldos_ini = dict(saldos_ini)
        self.reg = {c[0]: [] for c in CONTAS}
        self.seq = 1100000 + (mes - 5) * 9000

    def lanc(self):
        self.seq += random.randint(1, 4)
        return self.seq

    def add(self, conta, dia, linhas, cancelado=False):
        """linhas: lista de (historico, entrada, saida, cp)  -> um 'registro' (mesma sequencia)"""
        self.reg[conta].append({"dia": dia, "itens": linhas, "cancelado": cancelado})

    def dia(self, a=0, b=None):
        b = len(self.dias) - 1 if b is None else b
        return random.choice(self.dias[a:b + 1])

    # ---- tipos de lancamento ----
    def forn(self, conta, dia, valor=None, juros=True):
        nome, variantes = random.choice(FORNECEDORES)
        grafia = random.choice(variantes)
        doc = f"{random.randint(100, 9999):08d}"
        par = random.choice(["-1", "-2", "-3", "-A"])
        valor = valor or round(random.uniform(380, 9400), 2)
        itens = [(f"FORNECEDOR {doc} {par} - {grafia}", 0, valor, "4101010001")]
        if juros and random.random() < 0.55:
            j = round(valor * random.uniform(0.01, 0.03), 2)
            itens.append((f"JUROS PAGOS A FORNEC. {doc} {par} - {grafia}", 0, j, "4102010001"))
        self.add(conta, dia, itens)

    def forn_sem_nome(self, conta, dia):
        self.add(conta, dia, [("PAGAMENTO AVULSO MATERIAL DE OBRA", 0, round(random.uniform(150, 900), 2), "4101010001")])

    def cliente(self, conta, dia, valor=None):
        cli = random.choice(CLIENTES)
        doc = f"{random.randint(570000, 576999)}"
        par = random.choice(["26/01", "26/02", "26/03"])
        valor = valor or round(random.uniform(70, 3600), 2)
        self.add(conta, dia, [(f"RECB DUP 999 {doc}-{par} - {cli}", valor, 0, "3101010001")])

    def simples(self, conta, dia, hist, ent, sai, cp):
        self.add(conta, dia, [(hist, ent, sai, cp)])

    def transferencia(self, de, para, dia, valor, motivo):
        self.add(de, dia, [(f"SAI TRANSF REF {motivo}", 0, valor, "5101010020")])
        self.add(para, dia, [(f"ENTRA TRANSF REF {motivo}", valor, 0, "5101010011")])

    def montar(self, nova_contrapartida):
        du = self.dias
        mm = f"{self.mes:02d}/{str(self.ano)[2:]}"
        # ---- Caixa (dinheiro) ----
        for _ in range(17):
            self.simples(C_CX, self.dia(), "VENDA BALCAO DO DIA", round(random.uniform(1500, 6800), 2), 0, "3101010003")
        for _ in range(6):
            self.cliente(C_CX, self.dia())
        for _ in range(4):
            self.forn(C_CX, self.dia(), juros=False)
        self.simples(C_CX, du[3], f"PAGTO FRETE ENTREGA ROMANEIO {random.randint(100,999)}", 0, 780.00, "4101010002")
        self.simples(C_CX, du[11], f"PAGTO FRETE ENTREGA ROMANEIO {random.randint(100,999)}", 0, 640.00, "4101010002")
        self.simples(C_CX, du[8], "MATERIAL DE LIMPEZA E COPA", 0, 214.60, "4103010004")
        self.simples(C_CX, du[4], f"SALARIOS REF {mm} 1A QUINZENA", 0, 6850.00, "4103010001")
        self.simples(C_CX, du[15], f"SALARIOS REF {mm} 2A QUINZENA", 0, 7120.00, "4103010001")
        self.simples(C_CX, du[17], "RETIRADA SOCIO ADMINISTRADOR", 0, 5000.00, "6101010001")
        self.forn_sem_nome(C_CX, du[9])
        self.simples(C_CX, du[6], "PROPAGANDA RADIO LOCAL", 0, 850.00, "4104010001")
        for k in (2, 7, 13, 18):
            self.transferencia(C_CX, random.choice([C_ALFA, C_BETA]), du[k], round(random.choice([6000, 8000, 10000, 12000]), 2), "DEPOSITO EM CONTA")
        # ---- Cheques PRE ----
        for _ in range(12):
            n = f"756 {random.randint(3100,4999)} {random.randint(100,999)}"
            self.simples(C_PRE, self.dia(0, 14), f"ENTRA CAPT. DO Cheque : {n}", round(random.uniform(500, 9000), 2), 0, "5101010030")
        for k, v in ((5, 21000.00), (12, 18500.00), (19, 24200.00)):
            alvo = random.choice([C_ALFA, C_BETA])
            self.add(C_PRE, du[k], [("SAI DESCONTO CHEQUES PRE", 0, v, "5101010031")])
            self.add(alvo, du[k], [("ENTRA TRANSF REF DESCONTO CHEQUES", round(v * 0.985, 2), 0, "5101010011")])
            self.add(alvo, du[k], [("JUROS DESC CHEQUES", 0, round(v * 0.015, 2), "4102010004")])
        for k in (9, 16):
            self.simples(C_PRE, du[k], "SAI DEVOLUCAO CHEQUE P/ CLIENTE", 0, round(random.uniform(800, 2800), 2), "5101010040")
        # ---- Cheques DEV ----
        for _ in range(7):
            n = f"756 {random.randint(3100,4999)} {random.randint(100,999)}"
            v = round(random.uniform(300, 4200), 2)
            self.simples(C_DEV, self.dia(), f"ENTRA ESTORNO DO Ch.Dev : {n}", v, 0, "5101010007")
        for _ in range(5):
            n = f"033 {random.randint(1000,1999)} {random.randint(10,99)}"
            v = round(random.uniform(250, 2600), 2)
            self.simples(C_DEV, self.dia(), f"SAI BAIXA DO Parcial CH: {n}", 0, v, "5101010040")
        # ---- Bancos ----
        for conta in (C_ALFA, C_BETA):
            for _ in range(random.randint(13, 16)):
                self.cliente(conta, self.dia())
            for _ in range(random.randint(5, 6)):
                self.simples(conta, self.dia(), "RECEBIMENTO CARTOES MAQUINA 01", round(random.uniform(350, 1900), 2), 0, "3101010002")
            for _ in range(random.randint(8, 10)):
                self.forn(conta, self.dia())
            for d in du[::2]:
                self.simples(conta, d, "TARIFA COBRANCA", 0, random.choice([5.00, 6.00, 5.50]), "4102010002")
            for k in (3, 10, 17):
                self.simples(conta, du[k], "DEB IOF", 0, round(random.uniform(40, 330), 2), "4102010003")
        self.simples(C_ALFA, du[2], f"PAGTO ENERGIA ELETRICA REF {self.mes-1:02d}/{str(self.ano)[2:]}", 0, 3480.25, "4103010002")
        self.simples(C_ALFA, du[4], f"ALUGUEL GALPAO REF {mm}", 0, 7800.00, "4103010003")
        self.simples(C_BETA, du[9], f"IMPOSTOS E TAXAS GUIA {random.randint(100000,999999)}", 0, 4365.90, "4105010001")
        self.simples(C_BETA, du[14], f"PAGTO ICMS REF {self.mes-1:02d}/{str(self.ano)[2:]}", 0, 3112.40, "4105010001")
        self.transferencia(C_BETA, C_ALFA, du[6], 15000.00, "TRANSFERENCIA MESMA TITULARIDADE")
        # cancelados (como no ERP real: lancamento estornado fica na listagem)
        self.add(C_ALFA, du[5], [("DEB IOF", 0, 248.85, "4102010003")], cancelado=True)
        self.add(C_BETA, du[12], [("TARIFA COBRANCA", 0, 5.00, "4102010002")], cancelado=True)
        # contrapartidas novas (so no mes novo)
        if nova_contrapartida:
            self.simples(C_ALFA, du[7], "TARIFA PIX ENVIADO", 0, 3.20, "4102010009")
            self.simples(C_ALFA, du[8], "TARIFA PIX ENVIADO", 0, 3.20, "4102010009")
            self.simples(C_BETA, du[10], "TARIFA PIX ENVIADO", 0, 1.60, "4102010009")
            self.simples(C_BETA, du[13], "HONORARIOS CONTABEIS REF " + mm, 0, 1850.00, "4103010009")

    # ---- gera linhas do relatorio para cada conta ----
    def relatorio(self, conta):
        _, nome, *_ = next(c for c in CONTAS if c[0] == conta)
        ini = self.saldos_ini[conta]
        regs = sorted(self.reg[conta], key=lambda r: r["dia"])
        # numera os lancamentos na ordem do relatorio
        linhas_det = []  # tuplas ('det', lancto, hist, data, ent, sai, cp, multi, marker) | ('dia', saldo) | ('total', v)
        saldo = ini
        ent_t = sai_t = 0.0
        n_validas = 0
        resumo = {}
        dia_atual = None
        marker_alt = 0
        for r in regs:
            if dia_atual is not None and r["dia"] != dia_atual:
                linhas_det.append(("dia", round(saldo, 2)))
            dia_atual = r["dia"]
            ls = []
            for (hist, ent, sai, cp) in r["itens"]:
                ls.append(self.lanc())
                if not r["cancelado"]:
                    saldo += ent - sai
                    ent_t += ent
                    sai_t += sai
                    resumo.setdefault(cp, [0.0, 0.0, 0])
                    resumo[cp][0] += ent
                    resumo[cp][1] += sai
                    resumo[cp][2] += 1
            if not r["cancelado"]:
                n_validas += len(r["itens"])
            total_reg = sum(i[2] for i in r["itens"]) if len(r["itens"]) > 1 else None
            for idx, (hist, ent, sai, cp) in enumerate(r["itens"]):
                multi = len(hist) > 40
                if multi:
                    marker_alt ^= 1
                mk = ("*" if marker_alt else ":") if multi else " "
                linhas_det.append(("det", ls[idx], hist, r["dia"], ent, sai, cp, mk, r["cancelado"]))
            if total_reg is not None:
                linhas_det.append(("total", round(total_reg, 2)))
        linhas_det.append(("dia", round(saldo, 2)))
        return {"ini": ini, "fim": round(saldo, 2), "ent": round(ent_t, 2), "sai": round(sai_t, 2),
                "linhas": linhas_det, "resumo": resumo, "n": n_validas, "nome": nome}


def cabecalho(pag, conta, nome, periodo, emissao):
    L = []
    L.append({})
    L.append({1: "LISTAGEM DA CONTA P/PERÍODO"})
    L.append({}); L.append({})
    L.append({13: "LCX001", 15: f"Página: {pag}"})
    L.append({1: "Empresa:", 3: f"001 - {EMPRESA} ", 5: "CNPJ/CPF:", 7: CNPJ})
    L.append({}); L.append({}); L.append({})
    L.append({1: "Cidade:", 3: "CIDADE EXEMPLO - UF", 5: "Telefone:", 7: "(00) 0000-0000", 10: "Emissão:", 12: emissao})
    L.append({})
    L.append({3: f"Empresa : 001 / {conta}  - {nome}".ljust(66) + periodo})
    L.append({})
    L.append({1: "     Lancto    Historico                                  Dt.Docto              Entrada              Saida           S a l d o                        "})
    return L


DIV = "   +--------+-------------------------------------------+-----------+-------------------+-------------------+------------------+                                          "
RODAPE = f"Desenvolvido por {SISTEMA} - Rua Exemplo, 100 - Cidade Exemplo/UF - www.exemplo.com.br"


def linha_det(lancto, marker, texto, data, ent, sai, saldo, cp):
    t = (marker + " " + texto.ljust(40) + " ")
    d = (" " + data) if data else " " * 11
    e = (fmt(ent).rjust(18) + " ") if ent else " " * 19
    s = (fmt(sai).rjust(18) + " ") if sai else " " * 19
    sd = (fmt(saldo).rjust(17) + " ") if saldo is not None else " " * 18
    c = (" " + cp[0] + "  " + cp[1][:9]).ljust(42) if cp else " " * 42
    return "   |" + (str(lancto) if lancto else "").ljust(8) + "|" + t + "|" + d + "|" + e + "|" + s + "|" + sd + "|" + c


def escreve(ws, linhas):
    for i, d in enumerate(linhas, start=1):
        for col, val in d.items():
            ws.cell(row=i, column=col + 1, value=val)


def gera_arquivo(g, conta, caminho, periodo, emissao):
    rel = g.relatorio(conta)
    nome = rel["nome"]
    rows = []
    pag = 1
    rows += cabecalho(pag, conta, nome, periodo, emissao)
    rows.append({1: DIV})
    rows.append({1: "   |        |        Saldo Anterior =--))               |           |                   |                   |" + (fmt(rel["ini"]).rjust(17) + " ") + "|" + " " * 42})
    n_det = 0
    for ln in rel["linhas"]:
        # quebra de pagina (o ERP quebra em qualquer ponto, inclusive no meio do dia)
        if ln[0] == "det" and n_det >= 44:
            rows.append({1: RODAPE}); rows.append({}); rows.append({})
            pag += 1
            rows += cabecalho(pag, conta, nome, periodo, emissao)
            rows.append({1: DIV})
            n_det = 0
        if ln[0] == "det":
            _, lan, hist, dia, ent, sai, cp, mk, canc = ln
            data = dia.strftime("%d/%m/%Y")
            cpt = (cp, CP[cp][0])
            ent_txt = ent
            sai_txt = sai
            if len(hist) > 40:
                a, b = hist[:40], hist[40:]
                rows.append({1: linha_det(lan, mk, a, None, 0, 0, None, None)})
                n_det += 1
                txt = linha_det(None, mk, b, data, 0 if canc else ent, 0 if canc else sai, None, cpt)
            else:
                txt = linha_det(lan, mk, hist, data, 0 if canc else ent, 0 if canc else sai, None, cpt)
            if canc:  # [-Cancelado-] no lugar do valor (coluna 5 = Saida, 4 = Entrada)
                partes = txt.split("|")
                partes[5 if (sai or not ent) else 4] = "   [-Cancelado-]   "
                txt = "|".join(partes)
            rows.append({1: txt})
            n_det += 1
        elif ln[0] == "dia":
            rows.append({1: "   |        |          Saldo  do  Dia =--))             |           |                   |                   |" + (fmt(ln[1]).rjust(17) + " ") + "|" + " " * 42})
            n_det += 1
        elif ln[0] == "total":
            rows.append({1: "   |        |          --- Total do registro ----))    " + fmt(ln[1]) + ".| ..........|                   |                   |                  |" + " " * 42})
            n_det += 1
    rows.append({1: "   |        |                                           |           |          -        |          -        |         -        |" + " " * 42})
    rows.append({1: DIV})
    # bloco RESUMO por contrapartida (posicoes fixas: descricao col 5, codigo col 39)
    saldo = rel["ini"]
    for cp in sorted(rel["resumo"]):
        e, s, _n = rel["resumo"][cp]
        saldo += e - s
        linha = " " * 5 + CP[cp][0][:34].ljust(34) + cp + " " * 21
        linha += (fmt(e) if e else "").rjust(14) + " " * 2 + (fmt(s) if s else "").rjust(14) + (" " * 3) + fmt(saldo).rjust(15)
        rows.append({1: linha})
    rows.append({})
    rows.append({1: " " * 54 + "Saldo anterior  Total de Entradas   Total  de  Saidas       Saldo Atual"})
    rows.append({})
    rows.append({1: RODAPE})
    rows.append({}); rows.append({})
    pag += 1
    rows += cabecalho(pag, conta, nome, periodo, emissao)[:-1]
    rows.append({1: "     Lancto    Historico                                  Dt.Docto              Entrada              Saida           S a l d o"})
    rows.append({1: " " * 51 + "------------------ -----------------   ----------------- -----------------"})
    rows.append({1: "     Total do periodo-) " + periodo + " " * 12 + fmt(rel["ini"]).rjust(13) + fmt(rel["ent"]).rjust(18) + fmt(rel["sai"]).rjust(18) + fmt(rel["fim"]).rjust(18)})
    rows.append({1: f"   -- Data de impressao : {emissao[:10]} Hora : {emissao[-8:]} --"})
    rows.append({})
    rows.append({1: RODAPE})
    rows[0] = {0: " "}  # o ERP exporta a coluna A (vazia, mas dentro da area usada); sem isso o Power Query perde a 1a coluna
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Relatorio_LCX001"
    escreve(ws, rows)
    os.makedirs(os.path.dirname(caminho), exist_ok=True)
    wb.save(caminho)
    return {"conta": conta, "nome": nome, "ini": rel["ini"], "ent": rel["ent"], "sai": rel["sai"], "fim": rel["fim"], "linhas_validas": rel["n"], "cp": {k: v[2] for k, v in rel["resumo"].items()}}


def mes(ano, m, saldos_ini, nova, emissao, tag):
    g = Gerador(ano, m, saldos_ini)
    g.montar(nova)
    ult = dt.date(ano, m, 1)
    fim = (dt.date(ano + (m == 12), (m % 12) + 1, 1) - dt.timedelta(days=1))
    periodo = f"01/{m:02d}/{ano} a {fim.day:02d}/{m:02d}/{ano}"
    res = []
    for conta, nome_erp, nome_cf, deb, arq, _ in CONTAS:
        caminho = os.path.join(OUT, "erp_xlsx", tag, f"Livro Caixa - {arq} {tag}.xlsx")
        res.append(gera_arquivo(g, conta, caminho, periodo, emissao))
    return res


if __name__ == "__main__":
    ini = {c[0]: c[5] for c in CONTAS}
    maio = mes(2026, 5, ini, False, "02/06/2026 às 09:12:41", "Mai26")
    ini_jun = {r["conta"]: r["fim"] for r in maio}
    jun = mes(2026, 6, ini_jun, True, "01/07/2026 às 10:03:17", "Jun26")
    with open(os.path.join(OUT, "expected.json"), "w", encoding="utf-8") as f:
        json.dump({"Mai26": maio, "Jun26": jun, "contas": CONTAS, "cp": CP, "novas": sorted(NOVAS)}, f, ensure_ascii=False, indent=1)
    for tag, res in (("Mai26", maio), ("Jun26", jun)):
        print(tag)
        for r in res:
            print(f"  {r['nome']:<20} ini {r['ini']:>12,.2f}  ent {r['ent']:>12,.2f}  sai {r['sai']:>12,.2f}  fim {r['fim']:>12,.2f}  n={r['linhas_validas']}")
