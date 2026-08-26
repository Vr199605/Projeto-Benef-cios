# -*- coding: utf-8 -*-
"""
app_streamlit.py
=================
Interface Streamlit para o processo de transferência PC ELEA -> Reimplantação.

Como executar:
    pip install streamlit openpyxl pandas
    streamlit run app_streamlit.py

O usuário faz upload do arquivo PC ELEA (base de origem) e, opcionalmente,
de um template de Reimplantação diferente do padrão. Em seguida clica em
"Processar" e recebe o botão de download do arquivo final preenchido.
"""

import io
import logging
import tempfile
from pathlib import Path

import streamlit as st

from etl_reimplantacao import run_etl, logger

DEFAULT_TEMPLATE_PATH = Path(__file__).parent / "Reimplantação_ELEA_EM_BRANCO.xlsx"

st.set_page_config(page_title="Reimplantação ELEA — Automação", page_icon="📊", layout="centered")

st.title("📊 Automação de Reimplantação — ELEA")
st.write(
    "Faça upload da planilha **PC ELEA** (base de dados de origem) para gerar "
    "automaticamente a planilha de **Reimplantação** já formatada e preenchida."
)

# -------------------------------------------------------------------------
# Captura de logs para exibir na tela (além do console)
# -------------------------------------------------------------------------
class StreamlitLogHandler(logging.Handler):
    def __init__(self):
        super().__init__()
        self.records = []

    def emit(self, record):
        self.records.append(self.format(record))


# -------------------------------------------------------------------------
# Uploads
# -------------------------------------------------------------------------
col1, col2 = st.columns(2)

with col1:
    origem_file = st.file_uploader(
        "1) Planilha de origem (PC ELEA)",
        type=["xlsx"],
        help="Arquivo com os dados brutos a serem transferidos (ex.: PC_ELEA_2408.xlsx).",
    )

with col2:
    usar_template_padrao = st.checkbox(
        "Usar template padrão do sistema", value=DEFAULT_TEMPLATE_PATH.exists()
    )
    template_file = None
    if not usar_template_padrao:
        template_file = st.file_uploader(
            "2) Template de destino (Reimplantação em branco)",
            type=["xlsx"],
            help="Arquivo com o layout/formatação a ser preservado.",
        )

processar = st.button("🚀 Processar", type="primary", disabled=origem_file is None)

# -------------------------------------------------------------------------
# Processamento
# -------------------------------------------------------------------------
if processar:
    if origem_file is None:
        st.error("Envie a planilha de origem (PC ELEA) antes de processar.")
        st.stop()

    if usar_template_padrao and not DEFAULT_TEMPLATE_PATH.exists():
        st.error(
            "Template padrão não encontrado ao lado do app. Desmarque a opção "
            "acima e envie um template manualmente."
        )
        st.stop()

    if not usar_template_padrao and template_file is None:
        st.error("Envie o template de destino ou marque a opção de usar o template padrão.")
        st.stop()

    log_handler = StreamlitLogHandler()
    log_handler.setFormatter(logging.Formatter("%(levelname)-8s | %(message)s"))
    logger.addHandler(log_handler)

    with st.spinner("Processando... isso pode levar alguns segundos."):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_dir = Path(tmp_dir)

            origem_path = tmp_dir / "origem.xlsx"
            origem_path.write_bytes(origem_file.getvalue())

            if usar_template_padrao:
                template_path = DEFAULT_TEMPLATE_PATH
            else:
                template_path = tmp_dir / "template.xlsx"
                template_path.write_bytes(template_file.getvalue())

            saida_path = tmp_dir / "Reimplantacao_ELEA_PREENCHIDA.xlsx"

            try:
                resumo = run_etl(str(origem_path), str(template_path), str(saida_path))
                output_bytes = saida_path.read_bytes()
                sucesso = True
            except Exception as exc:  # noqa: BLE001
                sucesso = False
                st.exception(exc)

    logger.removeHandler(log_handler)

    if sucesso:
        st.success("Processamento concluído com sucesso!")

        st.subheader("Resumo por aba")
        st.table(
            [{"Aba": aba, "Linhas processadas": qtd} for aba, qtd in resumo.items()]
        )

        st.download_button(
            label="⬇️ Baixar planilha preenchida",
            data=output_bytes,
            file_name="Reimplantacao_ELEA_PREENCHIDA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        with st.expander("Ver log detalhado do processamento"):
            st.code("\n".join(log_handler.records), language="text")
    else:
        with st.expander("Ver log detalhado do erro"):
            st.code("\n".join(log_handler.records), language="text")

st.divider()
st.caption(
    "⚠️ Alguns campos codificados (Sexo, Estado Civil) usam uma tabela de "
    "códigos assumida a partir do padrão dos próprios dados — confirme com a "
    "área de negócio antes de usar o arquivo gerado em produção. "
    "Ajustes de mapeamento de colunas ficam em `mapping_config.py`."
)
