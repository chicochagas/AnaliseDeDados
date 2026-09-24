"""
Corrige as inconsistências da planilha de vendas e gera uma planilha tratada.

Entrada : data/planilha_vendas.csv
Saída   : data/planilha_vendas_tratada.xlsx
          - Vendas Tratadas : dados corrigidos + coluna de alertas
          - Resumo          : indicadores calculados por fórmula
          - Log de Correções: cada alteração feita (valor original -> corrigido)
          - Pendências      : itens que exigem confirmação humana

Uso: python scripts/corrigir_vendas.py [entrada.csv] [saida.xlsx]
"""

import csv
import re
import sys
import unicodedata
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

RAIZ = Path(__file__).resolve().parent.parent
ENTRADA = Path(sys.argv[1]) if len(sys.argv) > 1 else RAIZ / "data" / "planilha_vendas.csv"
SAIDA = Path(sys.argv[2]) if len(sys.argv) > 2 else RAIZ / "data" / "planilha_vendas_tratada.xlsx"

# ---------------------------------------------------------------- domínios
REGIOES = ["Norte", "Nordeste", "Centro-Oeste", "Sudeste", "Sul"]
STATUS_VALIDOS = {
    "fechado ganho": "Fechado Ganho",
    "closed won": "Fechado Ganho",
    "ganho": "Fechado Ganho",
    "fechado perdido": "Fechado Perdido",
    "closed lost": "Fechado Perdido",
    "perdido": "Fechado Perdido",
}
MESES = {
    "jan": 1, "janeiro": 1, "fev": 2, "fevereiro": 2, "mar": 3, "março": 3, "marco": 3,
    "abr": 4, "abril": 4, "mai": 5, "maio": 5, "jun": 6, "junho": 6,
    "jul": 7, "julho": 7, "ago": 8, "agosto": 8, "set": 9, "setembro": 9,
    "out": 10, "outubro": 10, "nov": 11, "novembro": 11, "dez": 12, "dezembro": 12,
}
# Marcações de qualidade que não devem ficar em "Observações"
MARCADORES_QUALIDADE = {"duplicado", "possível duplicado", "receita ausente", "sem representante"}
PARTICULAS = {"da", "de", "do", "das", "dos", "e"}

log = []         # (ID, coluna, original, corrigido, regra)
pendencias = []  # (ID, tipo, descrição, ação sugerida)
alertas = defaultdict(list)


def registrar(pid, coluna, original, corrigido, regra):
    if str(original) != str(corrigido):
        log.append((pid, coluna, original, corrigido, regra))


def moeda_br(valor):
    return "R$ " + f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def sem_acento(txt):
    return "".join(c for c in unicodedata.normalize("NFD", txt) if unicodedata.category(c) != "Mn")


# ---------------------------------------------------------------- normalizadores
def normalizar_nome(valor):
    partes = valor.strip().lower().split()
    return " ".join(p if (p in PARTICULAS and i > 0) else p.capitalize() for i, p in enumerate(partes))


def normalizar_regiao(valor):
    chave = sem_acento(valor.strip().lower()).replace(" ", "-")
    for r in REGIOES:
        if sem_acento(r.lower()) == chave:
            return r
    return ""


def normalizar_status(valor):
    return STATUS_VALIDOS.get(valor.strip().lower(), "")


def converter_data(valor):
    v = valor.strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(v, fmt)
        except ValueError:
            pass
    m = re.fullmatch(r"(\d{1,2})\s+([a-zà-ú]+)\.?\s+(\d{4})", v.lower())
    if m and m.group(2) in MESES:
        return datetime(int(m.group(3)), MESES[m.group(2)], int(m.group(1)))
    return None


def converter_receita(valor):
    v = valor.replace("R$", "").strip()
    if not v:
        return None
    if "," in v:  # formato brasileiro 1.234,56
        v = v.replace(".", "").replace(",", ".")
    return float(v)


# ---------------------------------------------------------------- 1. leitura e padronização
with open(ENTRADA, encoding="utf-8", newline="") as f:
    brutos = list(csv.DictReader(f))

registros = []
for b in brutos:
    pid = int(b["ID Pedido"])
    r = {"ID Pedido": pid}

    r["Nome do Cliente"] = normalizar_nome(b["Nome do Cliente"])
    registrar(pid, "Nome do Cliente", b["Nome do Cliente"], r["Nome do Cliente"], "Capitalização padronizada")

    r["Região"] = normalizar_regiao(b["Região"])
    registrar(pid, "Região", b["Região"], r["Região"], "Grafia padronizada para lista oficial")

    r["Representante"] = normalizar_nome(b["Representante"])
    registrar(pid, "Representante", b["Representante"], r["Representante"], "Capitalização padronizada")

    data = converter_data(b["Data do Pedido"])
    r["Data do Pedido"] = data
    if data is None:
        pendencias.append((pid, "Data inválida", f"Data '{b['Data do Pedido']}' não reconhecida", "Corrigir manualmente"))
    else:
        registrar(pid, "Data do Pedido", b["Data do Pedido"], data.strftime("%d/%m/%Y"), "Data convertida para DD/MM/AAAA")

    r["Receita"] = converter_receita(b["Receita"])
    if r["Receita"] is not None:
        registrar(pid, "Receita", b["Receita"], r["Receita"], "Texto convertido em número")

    r["Produto"] = b["Produto"].strip()

    obs = b["Observações"].strip()
    r["Status"] = normalizar_status(b["Status"])
    if b["Status"].strip() and not r["Status"]:
        pendencias.append((pid, "Status desconhecido", f"Status '{b['Status']}' fora da lista", "Classificar manualmente"))
    if not r["Status"] and normalizar_status(obs):
        # status digitado na coluna errada
        r["Status"] = normalizar_status(obs)
        registrar(pid, "Status", b["Status"] or "(vazio)", r["Status"], "Status recuperado da coluna Observações")
        obs = ""
    else:
        regra = "Traduzido para português" if b["Status"].strip().lower() == "closed won" else "Capitalização padronizada"
        registrar(pid, "Status", b["Status"], r["Status"], regra)

    if obs.lower() in MARCADORES_QUALIDADE:
        registrar(pid, "Observações", obs, "(vazio)", "Marcação de qualidade removida (tratada em Alertas/Log)")
        obs = ""
    r["Observações"] = obs
    registros.append(r)

# ---------------------------------------------------------------- 2. duplicados
def chave(r):
    return (r["Nome do Cliente"].lower(), r["Produto"], r["Receita"])

grupos = defaultdict(list)
for r in registros:
    grupos[chave(r)].append(r)

removidos = set()
for itens in grupos.values():
    if len(itens) < 2:
        continue
    itens.sort(key=lambda x: x["ID Pedido"])
    vistos = {}
    for r in itens:
        if r["Data do Pedido"] in vistos:  # mesma data = duplicado exato
            orig = vistos[r["Data do Pedido"]]
            removidos.add(r["ID Pedido"])
            log.append((r["ID Pedido"], "(linha inteira)", "registro", "REMOVIDO",
                        f"Duplicado exato do pedido {orig['ID Pedido']}"))
        else:
            vistos[r["Data do Pedido"]] = r
    restantes = list(vistos.values())
    if len(restantes) > 1:
        ids = ", ".join(str(x["ID Pedido"]) for x in restantes)
        status = {x["Status"] for x in restantes}
        extra = " (status divergentes: " + " x ".join(sorted(status)) + ")" if len(status) > 1 else ""
        for x in restantes:
            alertas[x["ID Pedido"]].append(f"Possível duplicado: pedidos {ids}{extra}")
        pendencias.append((ids, "Possível duplicado",
                           f"{restantes[0]['Nome do Cliente']} / {restantes[0]['Produto']} / {moeda_br(restantes[0]['Receita'])} em datas diferentes{extra}",
                           "Confirmar com o comercial se é recompra ou lançamento repetido"))

registros = [r for r in registros if r["ID Pedido"] not in removidos]

# ---------------------------------------------------------------- 3. campos vazios (inferência)
regioes_por_rep = defaultdict(Counter)
reps_por_regiao = defaultdict(Counter)
for r in registros:
    if r["Representante"] and r["Região"]:
        regioes_por_rep[r["Representante"]][r["Região"]] += 1
        reps_por_regiao[r["Região"]][r["Representante"]] += 1

for r in registros:
    pid = r["ID Pedido"]
    if not r["Região"]:
        opcoes = regioes_por_rep.get(r["Representante"], {})
        if len(opcoes) == 1:
            r["Região"] = next(iter(opcoes))
            registrar(pid, "Região", "(vazio)", r["Região"],
                      f"Inferida: {r['Representante']} atua apenas em {r['Região']}")
            alertas[pid].append("Região inferida pelo representante")
        else:
            alertas[pid].append("Região ausente")
            pendencias.append((pid, "Região ausente", "Não foi possível inferir", "Preencher manualmente"))
    if not r["Representante"]:
        opcoes = reps_por_regiao.get(r["Região"], {})
        if len(opcoes) == 1:
            r["Representante"] = next(iter(opcoes))
            registrar(pid, "Representante", "(vazio)", r["Representante"],
                      f"Inferido: único representante da região {r['Região']}")
            alertas[pid].append("Representante inferido pela região")
        else:
            alertas[pid].append("Representante ausente")
            pendencias.append((pid, "Representante ausente", "Não foi possível inferir", "Preencher manualmente"))
    if r["Receita"] is None:
        alertas[pid].append("Receita ausente")
        pendencias.append((pid, "Receita ausente", f"Pedido '{r['Status']}' sem valor ({r['Nome do Cliente']})",
                           "Buscar valor no contrato/nota fiscal"))
    if not r["Status"]:
        alertas[pid].append("Status ausente")
        pendencias.append((pid, "Status ausente", "Sem status nem indicação em Observações", "Preencher manualmente"))

# ---------------------------------------------------------------- 4. planilha
FONTE = "Arial"
CAB = Font(name=FONTE, bold=True, color="FFFFFF")
CAB_FILL = PatternFill("solid", start_color="1F4E78")
NORMAL = Font(name=FONTE)
NEGRITO = Font(name=FONTE, bold=True)
ALERTA_FILL = PatternFill("solid", start_color="FFF2CC")
BORDA = Border(bottom=Side(style="thin", color="BFBFBF"))
MOEDA = '"R$" #,##0.00;-"R$" #,##0.00;"-"'


def cabecalho(ws, colunas, larguras):
    ws.append(colunas)
    for i, (c, w) in enumerate(zip(ws[1], larguras), 1):
        c.font, c.fill = CAB, CAB_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"


wb = Workbook()

# --- Vendas Tratadas
ws = wb.active
ws.title = "Vendas Tratadas"
colunas = ["ID Pedido", "Nome do Cliente", "Região", "Representante", "Data do Pedido",
           "Receita (R$)", "Produto", "Status", "Observações", "Alertas"]
cabecalho(ws, colunas, [11, 22, 14, 18, 14, 15, 22, 17, 20, 55])
for r in registros:
    ws.append([r["ID Pedido"], r["Nome do Cliente"], r["Região"], r["Representante"], r["Data do Pedido"],
               r["Receita"], r["Produto"], r["Status"], r["Observações"] or None,
               "; ".join(alertas[r["ID Pedido"]]) or None])
for linha in ws.iter_rows(min_row=2):
    tem_alerta = linha[9].value is not None
    for c in linha:
        c.font, c.border = NORMAL, BORDA
        if tem_alerta:
            c.fill = ALERTA_FILL
    linha[4].number_format = "DD/MM/YYYY"
    linha[5].number_format = MOEDA
ultima = ws.max_row
ws.auto_filter.ref = f"A1:J{ultima}"

# --- Resumo (fórmulas sobre a aba de dados)
rs = wb.create_sheet("Resumo")
rs.column_dimensions["A"].width = 34
rs.column_dimensions["B"].width = 18
rs.column_dimensions["C"].width = 18
V = "'Vendas Tratadas'!"
F_REC, F_STA, F_REG = f"{V}$F$2:$F${ultima}", f"{V}$H$2:$H${ultima}", f"{V}$C$2:$C${ultima}"

rs["A1"] = "Resumo da Planilha de Vendas Tratada"
rs["A1"].font = Font(name=FONTE, bold=True, size=14)
indicadores = [
    ("Total de pedidos", f"=COUNTA({V}$A$2:$A${ultima})", "0"),
    ("Receita total (todos os status)", f"=SUM({F_REC})", MOEDA),
    ("Receita Fechado Ganho", f'=SUMIFS({F_REC},{F_STA},"Fechado Ganho")', MOEDA),
    ("Pedidos Fechado Ganho", f'=COUNTIFS({F_STA},"Fechado Ganho")', "0"),
    ("Pedidos Fechado Perdido", f'=COUNTIFS({F_STA},"Fechado Perdido")', "0"),
    ("Taxa de conversão", "=IFERROR(B6/(B6+B7),0)", "0.0%"),
    ("Ticket médio (ganhos com receita)",
     f'=IFERROR(B5/COUNTIFS({F_STA},"Fechado Ganho",{F_REC},"<>"),0)', MOEDA),
    ("Pedidos sem receita", f'=COUNTBLANK({F_REC})', "0"),
    ("Pedidos com alerta", f'=COUNTA({V}$J$2:$J${ultima})', "0"),
]
for i, (rot, form, fmt) in enumerate(indicadores, start=3):
    rs.cell(i, 1, rot).font = NORMAL
    c = rs.cell(i, 2, form)
    c.font, c.number_format = NEGRITO, fmt

base = 3 + len(indicadores) + 1
for j, t in enumerate(["Região", "Receita Ganha", "Pedidos"], 1):
    c = rs.cell(base, j, t)
    c.font, c.fill = CAB, CAB_FILL
for k, reg in enumerate(REGIOES, start=base + 1):
    rs.cell(k, 1, reg).font = NORMAL
    c = rs.cell(k, 2, f'=SUMIFS({F_REC},{F_REG},A{k},{F_STA},"Fechado Ganho")')
    c.font, c.number_format = NORMAL, MOEDA
    rs.cell(k, 3, f"=COUNTIFS({F_REG},A{k})").font = NORMAL
tot = base + len(REGIOES) + 1
rs.cell(tot, 1, "Total").font = NEGRITO
c = rs.cell(tot, 2, f"=SUM(B{base + 1}:B{tot - 1})")
c.font, c.number_format = NEGRITO, MOEDA
rs.cell(tot, 3, f"=SUM(C{base + 1}:C{tot - 1})").font = NEGRITO
rs.cell(tot + 2, 1, "Obs.: possíveis duplicados foram mantidos e estão sinalizados; "
                    "confirme-os na aba Pendências antes de usar os totais.").font = Font(name=FONTE, italic=True, color="7F7F7F")

# --- Log de Correções
lg = wb.create_sheet("Log de Correções")
cabecalho(lg, ["ID Pedido", "Coluna", "Valor Original", "Valor Corrigido", "Regra Aplicada"], [11, 18, 24, 24, 60])
for item in sorted(log, key=lambda x: (x[0], x[1])):
    lg.append([item[0], item[1], str(item[2]), str(item[3]), item[4]])
for linha in lg.iter_rows(min_row=2):
    for c in linha:
        c.font = NORMAL
lg.auto_filter.ref = f"A1:E{lg.max_row}"

# --- Pendências
pd_ = wb.create_sheet("Pendências")
cabecalho(pd_, ["ID(s) Pedido", "Tipo", "Descrição", "Ação Sugerida", "Resolvido? (S/N)"], [16, 20, 70, 50, 16])
for p in pendencias:
    pd_.append([str(p[0]), p[1], p[2], p[3], None])
for linha in pd_.iter_rows(min_row=2):
    for c in linha:
        c.font = NORMAL
    linha[4].fill = PatternFill("solid", start_color="FFFF00")
pd_.cell(pd_.max_row + 2, 1, "Preencha a coluna amarela após validar cada item com a área comercial.").font = \
    Font(name=FONTE, italic=True, color="7F7F7F")

SAIDA.parent.mkdir(parents=True, exist_ok=True)
wb.save(SAIDA)

print(f"Linhas lidas: {len(brutos)} | removidas: {len(removidos)} | finais: {len(registros)}")
print(f"Correções registradas: {len(log)} | pendências: {len(pendencias)}")
print(f"Arquivo gerado: {SAIDA}")
