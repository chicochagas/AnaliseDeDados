"""
Teste de fumaça do app Streamlit: carrega a planilha de exemplo e confere
indicadores, validação e botões de download.

Uso: python -m pytest tests -v
"""

from pathlib import Path

import pytest
from streamlit.testing.v1 import AppTest

APP = Path(__file__).resolve().parent.parent / "app.py"


@pytest.fixture(scope="module")
def app():
    at = AppTest.from_file(str(APP), default_timeout=60).run()
    at.button[0].click().run()  # "Usar planilha de exemplo"
    return at


def test_app_sem_erros(app):
    assert not app.exception
    assert not app.error


def test_validacao_aprovada(app):
    assert "Limpeza validada" in app.success[0].value


def test_indicadores(app):
    metricas = {m.label: m.value for m in app.metric}
    assert metricas["Linhas lidas"] == "30"
    assert metricas["Linhas finais"] == "29"
    assert metricas["Melhorias aplicadas"] == "87"
    assert metricas["Pendências para revisão"] == "6"
    assert metricas["Erros encontrados"] == "93"


def test_tela_inicial_pede_arquivo():
    at = AppTest.from_file(str(APP), default_timeout=60).run()
    assert not at.exception
    assert "Envie um arquivo CSV" in at.info[0].value
