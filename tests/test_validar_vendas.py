"""
Testes das validações genéricas (scripts/validar_vendas.py) e da leitura de CSV
enviados pelo app: separador ';', codificação Windows, valores inválidos.

Uso: python -m pytest tests -v
"""

import copy
import io
import sys
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ / "scripts"))

from corrigir_vendas import ErroEntrada, gerar_excel, ler_csv, limpar  # noqa: E402
from validar_vendas import validar  # noqa: E402

CABECALHO = "ID Pedido;Nome do Cliente;Região;Representante;Data do Pedido;Receita;Produto;Status;Observações"
CSV_SUJO = "\n".join([
    CABECALHO,
    "1;ana DA silva;SUL;BRUNO CARVALHO;01/02/2024;R$ 1.234,56;Hardware;fechado ganho;",
    "2;ana da silva;Sul;Bruno Carvalho;01/02/2024;1234,56;Hardware;Fechado Ganho;duplicado",
    "3;Beto Lima;Marte;Carla Dias;31/02/2024;abc;Consultoria;em negociação;",
    "4;Caio Reis;;Bruno Carvalho;2024-03-10;;Consultoria;;",
    "ABC;Dora Melo;Norte;;5 abr. 2024;R$ 500,00;Hardware;closed won;sem representante",
    "6;Eva Luz;Norte;Fabio Nunes;10/04/2024;R$ 700,00;Hardware;Fechado Perdido;",
])


def exemplo():
    brutos = ler_csv(RAIZ / "data" / "planilha_vendas.csv")
    return brutos, limpar(brutos)


def sujo(codificacao="utf-8"):
    brutos = ler_csv(CSV_SUJO.encode(codificacao))
    return brutos, limpar(brutos)


def falhas(validacoes):
    return [v["regra"] + ": " + v["detalhe"] for v in validacoes if not v["ok"]]


# ---------------------------------------------------------------- leitura
def test_le_csv_com_virgula_e_ponto_e_virgula():
    assert len(ler_csv(RAIZ / "data" / "planilha_vendas.csv")) == 30
    assert len(ler_csv(CSV_SUJO.encode())) == 6


def test_le_csv_em_codificacao_windows():
    brutos = ler_csv(CSV_SUJO.encode("cp1252"))
    assert brutos[0]["Região"] == "SUL"
    assert "Região" in brutos[0]


def test_colunas_ausentes_geram_erro_claro():
    with pytest.raises(ErroEntrada, match="Receita"):
        ler_csv("ID Pedido,Nome do Cliente\n1,Ana".encode())


def test_arquivo_vazio_gera_erro():
    with pytest.raises(ErroEntrada):
        ler_csv(b"   ")


# ---------------------------------------------------------------- validações
def test_exemplo_passa_em_todas_as_validacoes():
    brutos, resultado = exemplo()
    assert falhas(validar(brutos, resultado)) == []


def test_csv_sujo_passa_em_todas_as_validacoes():
    brutos, resultado = sujo()
    assert falhas(validar(brutos, resultado)) == []


def test_csv_sujo_correcoes_e_pendencias():
    _, resultado = sujo()
    regs = {r["ID Pedido"]: r for r in resultado["registros"]}
    assert resultado["removidos"] == {2}                     # duplicado exato do 1
    assert regs[1]["Nome do Cliente"] == "Ana da Silva"
    assert regs[1]["Receita"] == pytest.approx(1234.56)
    assert regs[4]["Região"] == "Sul"                         # inferida pelo representante
    assert regs["ABC"]["Status"] == "Fechado Ganho"
    assert regs["ABC"]["Observações"] == ""
    assert regs["ABC"]["Representante"] == "Fabio Nunes"      # único representante do Norte
    tipos = {p["tipo"] for p in resultado["pendencias"]}
    assert {"ID inválido", "Região inválida", "Data inválida", "Receita inválida",
            "Status desconhecido", "Receita ausente", "Status ausente"} <= tipos


def test_excel_gerado_em_memoria():
    _, resultado = sujo()
    buffer = io.BytesIO()
    gerar_excel(resultado, buffer)
    assert buffer.getvalue()[:2] == b"PK"  # arquivo xlsx (zip)


# ---------------------------------------------------------------- as validações detectam erros
@pytest.mark.parametrize("sabotagem, regra", [
    (lambda r: r["registros"][0].update({"Região": "sul"}), "Região válida ou sinalizada"),
    (lambda r: r["registros"][0].update({"Nome do Cliente": "JOÃO SILVA"}), "Nome do Cliente: capitalização padronizada"),
    (lambda r: r["registros"][0].update({"Status": "closed won"}), "Status válido ou sinalizado"),
    (lambda r: r["registros"][0].update({"Data do Pedido": "15/01/2024"}), "Datas convertidas ou sinalizadas"),
    (lambda r: r["registros"][0].update({"Observações": "duplicado"}), "Observações sem marcadores de erro"),
    (lambda r: r["registros"].append(copy.deepcopy(r["registros"][0])), "Sem duplicados exatos"),
    (lambda r: r["registros"].pop(), "Nenhum pedido perdido na limpeza"),
    (lambda r: r["alertas"].clear(), "Possíveis duplicados sinalizados"),
    (lambda r: r["log"].clear(), "Toda alteração registrada no log"),
])
def test_validacao_detecta_sabotagem(sabotagem, regra):
    brutos, resultado = exemplo()
    sabotagem(resultado)
    resultado_regra = next(v for v in validar(brutos, resultado) if v["regra"] == regra)
    assert not resultado_regra["ok"]
