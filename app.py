"""
App web para limpeza da planilha de vendas.

O usuário envia o CSV, o app aplica a limpeza (scripts/corrigir_vendas.py),
roda as validações (scripts/validar_vendas.py), mostra os erros e melhorias
e disponibiliza a planilha limpa em CSV ou Excel.

Uso: streamlit run app.py
"""

import io
import sys
from collections import Counter
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

RAIZ = Path(__file__).resolve().parent
sys.path.insert(0, str(RAIZ / "scripts"))

from corrigir_vendas import COLUNAS_OBRIGATORIAS, ErroEntrada, chave_id, gerar_excel, ler_csv, limpar  # noqa: E402
from validar_vendas import validar  # noqa: E402

EXEMPLO = RAIZ / "data" / "planilha_vendas.csv"
COR_CORRIGIDO = "#2a78d6"
COR_PENDENTE = "#eb6834"

st.set_page_config(page_title="Limpeza de Vendas", page_icon="🧹", layout="wide")


# ---------------------------------------------------------------- processamento
@st.cache_data(show_spinner=False)
def processar(conteudo: bytes):
    brutos = ler_csv(conteudo)
    resultado = limpar(brutos)
    return brutos, resultado, validar(brutos, resultado)


def tabela_limpa(resultado):
    alertas = resultado["alertas"]
    linhas = [{
        "ID Pedido": str(r["ID Pedido"]),
        "Nome do Cliente": r["Nome do Cliente"],
        "Região": r["Região"],
        "Representante": r["Representante"],
        "Data do Pedido": r["Data do Pedido"],
        "Receita": r["Receita"],
        "Produto": r["Produto"],
        "Status": r["Status"],
        "Observações": r["Observações"],
        "Alertas": "; ".join(alertas.get(r["ID Pedido"], [])),
    } for r in resultado["registros"]]
    df = pd.DataFrame(linhas, columns=["ID Pedido", "Nome do Cliente", "Região", "Representante", "Data do Pedido",
                                       "Receita", "Produto", "Status", "Observações", "Alertas"])
    df["Data do Pedido"] = pd.to_datetime(df["Data do Pedido"])
    return df


def gerar_csv(df, padrao_brasil):
    saida = df.copy()
    saida["Data do Pedido"] = saida["Data do Pedido"].dt.strftime("%d/%m/%Y")
    if padrao_brasil:
        return saida.to_csv(index=False, sep=";", decimal=",", float_format="%.2f").encode("utf-8-sig")
    return saida.to_csv(index=False, float_format="%.2f").encode("utf-8")


def gerar_xlsx(resultado):
    buffer = io.BytesIO()
    gerar_excel(resultado, buffer)
    return buffer.getvalue()


def grafico_barras(contagem, cor, titulo):
    df = pd.DataFrame(contagem.most_common(), columns=["Tipo", "Quantidade"])
    base = alt.Chart(df, title=titulo).encode(
        y=alt.Y("Tipo:N", sort="-x", title=None, axis=alt.Axis(labelLimit=280)),
        x=alt.X("Quantidade:Q", title=None, axis=alt.Axis(labels=False, ticks=False, grid=False, domain=False)),
        tooltip=[alt.Tooltip("Tipo:N"), alt.Tooltip("Quantidade:Q")],
    )
    barras = base.mark_bar(color=cor, cornerRadiusEnd=4, height=18)
    rotulos = base.mark_text(align="left", dx=6).encode(text="Quantidade:Q")
    return (barras + rotulos).properties(height=max(90, 34 * len(df)))


def destacar_alertas(linha):
    estilo = "background-color: #FFF2CC; color: #000000" if linha["Alertas"] else ""
    return [estilo] * len(linha)


# ---------------------------------------------------------------- cabeçalho e upload
st.title("🧹 Limpeza da Planilha de Vendas")
st.caption("Envie o CSV de vendas: o app corrige as inconsistências, valida o resultado "
           "e gera a planilha limpa para download.")

with st.expander("Formato esperado do arquivo"):
    st.markdown("CSV separado por **vírgula** ou **ponto e vírgula**, com as colunas:")
    st.code(", ".join(COLUNAS_OBRIGATORIAS), language=None)
    st.markdown(
        "**O que é corrigido automaticamente:** duplicados exatos, maiúsculas/minúsculas, grafia de regiões, "
        "status em inglês ou na coluna errada, datas em formatos diferentes, receita como texto, "
        "região/representante vazios (quando dá para deduzir) e marcações de erro em Observações.\n\n"
        "**O que vai para revisão:** possíveis duplicados, receita ausente ou inválida, datas inválidas e "
        "campos que não puderam ser deduzidos."
    )

col_upload, col_exemplo = st.columns([4, 1], vertical_alignment="bottom")
arquivo = col_upload.file_uploader("Arquivo CSV", type=["csv"])
if EXEMPLO.exists() and col_exemplo.button("Usar planilha de exemplo", width="stretch"):
    st.session_state["exemplo"] = True
if arquivo is not None:
    st.session_state["exemplo"] = False

if arquivo is not None:
    conteudo, nome = arquivo.getvalue(), Path(arquivo.name).stem
elif st.session_state.get("exemplo"):
    conteudo, nome = EXEMPLO.read_bytes(), EXEMPLO.stem
else:
    st.info("Envie um arquivo CSV para começar.")
    st.stop()

try:
    with st.spinner("Limpando e validando a planilha..."):
        brutos, resultado, validacoes = processar(conteudo)
except ErroEntrada as erro:
    st.error(f"Não foi possível ler o arquivo: {erro}")
    st.stop()

if not brutos:
    st.warning("O arquivo não tem linhas de dados.")
    st.stop()

log, pendencias, alertas = resultado["log"], resultado["pendencias"], resultado["alertas"]
df = tabela_limpa(resultado)

# ---------------------------------------------------------------- indicadores
aprovadas = sum(v["ok"] for v in validacoes)
if aprovadas == len(validacoes):
    st.success(f"✅ Limpeza validada: {aprovadas} de {len(validacoes)} regras aprovadas.")
else:
    st.error(f"❌ {len(validacoes) - aprovadas} regra(s) de validação falharam. Veja a aba **Validação**.")

ids_antes = {e["id"] for e in log} | set(alertas)
ids_depois = {pid for pid in alertas if pid not in resultado["removidos"]}
lidas = resultado["lidas"]

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Linhas lidas", lidas)
m2.metric("Erros encontrados", len(log) + len(pendencias), help="Correções automáticas + pendências de revisão")
m3.metric("Melhorias aplicadas", len(log), help="Alterações feitas automaticamente (detalhes no log)")
m4.metric("Pendências para revisão", len(pendencias), help="Itens que precisam de confirmação humana")
m5.metric("Linhas finais", len(resultado["registros"]),
          delta=f"-{len(resultado['removidos'])} duplicada(s)" if resultado["removidos"] else None,
          delta_color="off")

q1, q2 = st.columns(2)
q1.metric("Linhas com algum problema — antes", f"{len(ids_antes)} ({len(ids_antes) / lidas:.0%})")
q2.metric("Linhas que ainda pedem revisão — depois", f"{len(ids_depois)} ({len(ids_depois) / lidas:.0%})")

# ---------------------------------------------------------------- downloads
st.subheader("Baixar planilha limpa")
d1, d2, d3 = st.columns([2, 1, 1], vertical_alignment="bottom")
formato = d1.radio("Formato do CSV", ["Excel Brasil (; e vírgula decimal)", "Padrão (, e ponto decimal)"],
                   horizontal=True)
d2.download_button("⬇️ Baixar CSV", gerar_csv(df, formato.startswith("Excel")), f"{nome}_tratada.csv",
                   "text/csv", width="stretch")
d3.download_button("⬇️ Baixar Excel", gerar_xlsx(resultado), f"{nome}_tratada.xlsx",
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   type="primary", width="stretch")
st.caption("O Excel inclui as abas Vendas Tratadas, Resumo (com fórmulas), Log de Correções e Pendências.")

# ---------------------------------------------------------------- abas de detalhe
aba_erros, aba_dados, aba_log, aba_pend, aba_val, aba_orig = st.tabs(
    ["📊 Erros e melhorias", "✅ Planilha limpa", f"📝 Log ({len(log)})", f"⚠️ Pendências ({len(pendencias)})",
     f"🧪 Validação ({aprovadas}/{len(validacoes)})", "📄 Arquivo original"])

with aba_erros:
    g1, g2 = st.columns(2)
    corrigidos = Counter(e["categoria"] for e in log)
    pendentes = Counter(p["tipo"] for p in pendencias)
    with g1:
        if corrigidos:
            st.altair_chart(grafico_barras(corrigidos, COR_CORRIGIDO, "Corrigidos automaticamente"),
                            width="stretch")
        else:
            st.info("Nenhuma correção foi necessária.")
    with g2:
        if pendentes:
            st.altair_chart(grafico_barras(pendentes, COR_PENDENTE, "Pendentes de revisão"),
                            width="stretch")
        else:
            st.success("Nenhuma pendência: tudo foi corrigido automaticamente.")
    resumo = pd.DataFrame(
        [{"Tipo": t, "Quantidade": q, "Situação": "Corrigido"} for t, q in corrigidos.most_common()]
        + [{"Tipo": t, "Quantidade": q, "Situação": "Pendente"} for t, q in pendentes.most_common()])
    with st.expander("Ver como tabela"):
        st.dataframe(resumo, hide_index=True, width="stretch")

with aba_dados:
    so_alertas = st.toggle(f"Mostrar só linhas com alerta ({(df['Alertas'] != '').sum()})")
    exibir = df[df["Alertas"] != ""] if so_alertas else df
    st.dataframe(
        exibir.style.apply(destacar_alertas, axis=1), hide_index=True, width="stretch",
        column_config={
            "Data do Pedido": st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Receita": st.column_config.NumberColumn("Receita (R$)", format="R$ %.2f"),
            "Alertas": st.column_config.TextColumn(width="large"),
        })
    st.caption("Linhas em amarelo têm alerta: valor deduzido, possível duplicado ou dado ausente.")

with aba_log:
    df_log = pd.DataFrame(sorted(log, key=lambda e: (chave_id(e["id"]), e["coluna"])))
    if df_log.empty:
        st.info("Nenhuma alteração foi necessária.")
    else:
        df_log = df_log.rename(columns={"id": "ID Pedido", "coluna": "Coluna", "original": "Valor Original",
                                        "corrigido": "Valor Corrigido", "regra": "Regra", "categoria": "Categoria"})
        df_log = df_log.astype(str)
        categorias = st.multiselect("Filtrar por categoria", sorted(df_log["Categoria"].unique()))
        if categorias:
            df_log = df_log[df_log["Categoria"].isin(categorias)]
        st.dataframe(df_log, hide_index=True, width="stretch")

with aba_pend:
    if pendencias:
        st.dataframe(pd.DataFrame(pendencias).rename(columns={
            "ids": "ID(s) Pedido", "tipo": "Tipo", "descricao": "Descrição", "acao": "Ação Sugerida"}),
            hide_index=True, width="stretch")
    else:
        st.success("Nenhuma pendência.")

with aba_val:
    st.caption("Regras conferidas automaticamente após a limpeza (as mesmas usadas na suíte de testes).")
    st.dataframe(pd.DataFrame([{"Resultado": "✅" if v["ok"] else "❌", "Regra": v["regra"],
                                "Detalhe": v["detalhe"]} for v in validacoes]),
                 hide_index=True, width="stretch")

with aba_orig:
    st.caption("Dados exatamente como foram enviados.")
    st.dataframe(pd.DataFrame(brutos), hide_index=True, width="stretch")
