"""
Validações genéricas do resultado da limpeza (valem para qualquer CSV enviado).

Cada regra confere que um item do diagnóstico foi corrigido ou, quando não dá
para corrigir automaticamente, que ficou sinalizado nos Alertas.

Usado pelo app Streamlit e pela suíte pytest:
    validacoes = validar(brutos, resultado)
    -> [{"regra": str, "ok": bool, "detalhe": str}, ...]
"""

from collections import defaultdict
from datetime import datetime

from corrigir_vendas import MARCADORES_QUALIDADE, PARTICULAS, REGIOES, STATUS_VALIDOS

STATUS_OFICIAIS = set(STATUS_VALIDOS.values())
MARCADORES = MARCADORES_QUALIDADE | set(STATUS_VALIDOS)


def _nome_padronizado(nome):
    for i, palavra in enumerate(nome.split()):
        if i > 0 and palavra in PARTICULAS:
            continue
        if not (palavra[0].isupper() and palavra[1:] == palavra[1:].lower()):
            return False
    return True


def _resumo(ids, limite=8):
    ids = [str(i) for i in ids]
    return ", ".join(ids[:limite]) + (f" e mais {len(ids) - limite}" if len(ids) > limite else "")


def validar(brutos, resultado):
    registros = resultado["registros"]
    alertas = resultado["alertas"]
    log = resultado["log"]
    removidos = resultado["removidos"]
    resultados = []

    def regra(nome, falhas, detalhe_ok):
        resultados.append({
            "regra": nome,
            "ok": not falhas,
            "detalhe": detalhe_ok if not falhas else "Falhou em: " + _resumo(falhas),
        })

    def sinalizado(r, termo):
        return any(termo in a for a in alertas.get(r["ID Pedido"], []))

    # 1. Nenhum pedido perdido
    ids_entrada = {r["_original"]["ID Pedido"].strip() for r in registros} | {str(i) for i in removidos}
    ids_originais = {b["ID Pedido"].strip() for b in brutos}
    faltando = sorted(ids_originais - ids_entrada)
    regra("Nenhum pedido perdido na limpeza", faltando,
          f"{len(registros)} mantidos + {len(removidos)} duplicados removidos = {len(brutos)} lidos")

    # 2. Duplicados exatos removidos e registrados
    campos = ["Nome do Cliente", "Região", "Representante", "Data do Pedido", "Receita", "Produto", "Status"]
    vistos, repetidos = set(), []
    for r in registros:
        chave = tuple(str(r[c]).lower() for c in campos)
        if r["Data do Pedido"] is not None and chave in vistos:
            repetidos.append(r["ID Pedido"])
        vistos.add(chave)
    regra("Sem duplicados exatos", repetidos, f"{len(removidos)} duplicado(s) exato(s) removido(s)")

    no_log = {e["id"] for e in log if e["corrigido"] == "REMOVIDO"}
    regra("Duplicados removidos registrados no log", sorted(removidos - no_log, key=str),
          "Toda remoção tem justificativa no log")

    # 3. Possíveis duplicados sinalizados
    grupos = defaultdict(list)
    for r in registros:
        if r["Receita"] is not None and r["Nome do Cliente"]:
            grupos[(r["Nome do Cliente"].lower(), r["Produto"].lower(), r["Receita"])].append(r)
    suspeitos = [r for g in grupos.values() if len(g) > 1 for r in g]
    regra("Possíveis duplicados sinalizados",
          [r["ID Pedido"] for r in suspeitos if not sinalizado(r, "Possível duplicado")],
          f"{len(suspeitos)} pedido(s) em possível duplicidade sinalizado(s)")

    # 4. Padronização de texto
    for coluna in ("Nome do Cliente", "Representante"):
        regra(f"{coluna}: capitalização padronizada",
              [r["ID Pedido"] for r in registros if r[coluna] and not _nome_padronizado(r[coluna])],
              "Todos com iniciais maiúsculas")

    # 5. Campos obrigatórios válidos ou sinalizados
    regra("Região válida ou sinalizada",
          [r["ID Pedido"] for r in registros if r["Região"] not in REGIOES and not sinalizado(r, "Região")],
          "Todas na lista oficial: " + ", ".join(REGIOES))
    regra("Representante preenchido ou sinalizado",
          [r["ID Pedido"] for r in registros if not r["Representante"] and not sinalizado(r, "Representante")],
          "Nenhum pedido sem representante sem aviso")
    regra("Status válido ou sinalizado",
          [r["ID Pedido"] for r in registros if r["Status"] not in STATUS_OFICIAIS and not sinalizado(r, "Status")],
          "Somente 'Fechado Ganho' e 'Fechado Perdido'")

    # 6. Datas e receita
    regra("Datas convertidas ou sinalizadas",
          [r["ID Pedido"] for r in registros
           if not isinstance(r["Data do Pedido"], datetime) and not sinalizado(r, "Data")],
          "Todas as datas no formato DD/MM/AAAA")
    regra("Receita numérica ou sinalizada",
          [r["ID Pedido"] for r in registros
           if not isinstance(r["Receita"], (int, float)) and not sinalizado(r, "Receita")],
          "Receita armazenada como número")

    # 7. Observações limpas
    regra("Observações sem marcadores de erro",
          [r["ID Pedido"] for r in registros if r["Observações"].strip().lower() in MARCADORES],
          "Marcações de erro movidas para Alertas/Log")

    # 8. Rastreabilidade
    registradas = {(e["id"], e["coluna"]) for e in log}
    sem_log = []
    for r in registros:
        for coluna in ("Nome do Cliente", "Região", "Representante", "Status"):
            original = r["_original"][coluna].strip()
            if original != r[coluna] and r[coluna] and (r["ID Pedido"], coluna) not in registradas:
                sem_log.append(f"{r['ID Pedido']}/{coluna}")
    regra("Toda alteração registrada no log", sem_log, f"{len(log)} alterações rastreáveis")

    inferidos = {e["id"] for e in log if e["regra"].startswith("Inferid")}
    regra("Valores inferidos sinalizados",
          [pid for pid in inferidos if not any("inferid" in a for a in alertas.get(pid, []))],
          f"{len(inferidos)} pedido(s) com valor deduzido e marcado")

    return resultados
