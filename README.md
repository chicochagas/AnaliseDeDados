# AnaliseDeDados

Limpeza e validação da planilha de vendas.

## Instalação

```
pip install -r requirements.txt
```

## App web

```
streamlit run app.py
```

Envie o CSV de vendas (ou use a planilha de exemplo). O app corrige as inconsistências, valida o resultado,
mostra os erros encontrados e as melhorias aplicadas, e permite baixar a planilha limpa em CSV ou Excel.

## Linha de comando

```
python scripts/corrigir_vendas.py [entrada.csv] [saida.xlsx]
```

Padrão: lê `data/planilha_vendas.csv` e gera `data/planilha_vendas_tratada.xlsx`.

## Testes

```
python -m pytest
```

| Arquivo | O que testa |
|---|---|
| `scripts/corrigir_vendas.py` | Limpeza (leitura, correções, geração do Excel) |
| `scripts/validar_vendas.py` | Regras de validação genéricas, usadas pelo app e pelos testes |
| `tests/test_corrigir_vendas.py` | Cada item do diagnóstico na planilha de exemplo |
| `tests/test_validar_vendas.py` | Validações com CSV sujo sintético e detecção de erros |
| `tests/test_app.py` | Teste de fumaça do app Streamlit |
