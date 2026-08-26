# -*- coding: utf-8 -*-
"""
etl_reimplantacao.py
=====================
Motor de transferência de dados PC ELEA -> Reimplantação ELEA.

Uso via linha de comando:
    python etl_reimplantacao.py --origem PC_ELEA_2408.xlsx \
                                 --template Reimplantação_ELEA_EM_BRANCO.xlsx \
                                 --saida Reimplantacao_ELEA_PREENCHIDA.xlsx

Uso programático (também usado pelo app Streamlit):
    from etl_reimplantacao import run_etl
    resumo = run_etl("PC_ELEA_2408.xlsx",
                      "Reimplantação_ELEA_EM_BRANCO.xlsx",
                      "saida.xlsx")

Regras de projeto seguidas (ver mapping_config.py para ajustar mapeamentos):
  - O template de destino é aberto com openpyxl e preenchido célula a célula.
    NUNCA é recriado com pandas, para preservar 100% de estilos, larguras de
    coluna, cores e fórmulas pré-existentes.
  - Qualquer coluna do destino que não tenha mapeamento definido em
    mapping_config.py é deixada como está no template (herdando o valor
    constante e a formatação da "linha-modelo", ver DST_TEMPLATE_ROW).
  - Datas são convertidas para o formato inteiro DDMMYYYY já usado pelo
    template (célula com number_format '00000000'), nunca texto.
  - Valores nulos/NaN da origem nunca são gravados como a string "NaN" —
    viram célula vazia (None).
"""

from __future__ import annotations

import argparse
import copy
import logging
import re
from datetime import datetime
from typing import Any, Dict, List, Optional

import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

import mapping_config as cfg

# -----------------------------------------------------------------------
# Logging
# -----------------------------------------------------------------------
logger = logging.getLogger("etl_reimplantacao")
if not logger.handlers:
    _h = logging.StreamHandler()
    _h.setFormatter(logging.Formatter("%(asctime)s | %(levelname)-8s | %(message)s"))
    logger.addHandler(_h)
logger.setLevel(logging.INFO)


# =========================================================================
# 1. FUNÇÕES DE TRATAMENTO / CONVERSÃO DE DADOS
# =========================================================================

def _only_digits(value: Any) -> str:
    """Remove tudo que não for dígito. NaN/None viram string vazia."""
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return re.sub(r"\D", "", str(value))


def clean_text(value: Any) -> Optional[str]:
    """strip() em textos; NaN/None/"" viram None (célula em branco)."""
    if value is None:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    text = str(value).strip()
    return text if text else None


def to_int_or_none(value: Any) -> Optional[int]:
    """Converte para int quando possível; caso contrário, None (nunca "NaN")."""
    text = clean_text(value)
    if text is None:
        return None
    try:
        return int(float(text))
    except (ValueError, TypeError):
        return None


def data_ddmmyyyy(value: Any) -> Optional[int]:
    """
    Converte uma data no formato 'dd/mm/aaaa' (texto, como vem da origem)
    para inteiro DDMMYYYY, compatível com o number_format '00000000' já
    usado no template (ex.: 02/09/1957 -> 2091957 -> exibido 02091957).
    Datas inválidas ou placeholders (ex.: 01/01/0001) retornam None.
    """
    text = clean_text(value)
    if text is None:
        return None
    dt = None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            dt = datetime.strptime(text, fmt)
            break
        except ValueError:
            continue
    if dt is None and isinstance(value, datetime):
        dt = value
    if dt is None or dt.year < 1900:
        return None
    return int(dt.strftime("%d%m%Y"))


def cpf_base(value: Any) -> Optional[str]:
    """9 primeiros dígitos do CPF (sem os 2 dígitos verificadores)."""
    digits = _only_digits(value)
    if not digits:
        return None
    digits = digits.zfill(11)[-11:]
    return digits[:9]


def cpf_control(value: Any) -> Optional[str]:
    """2 dígitos verificadores do CPF."""
    digits = _only_digits(value)
    if not digits:
        return None
    digits = digits.zfill(11)[-11:]
    return digits[9:11]


def cep_base(value: Any) -> Optional[str]:
    digits = _only_digits(value)
    if not digits:
        return None
    digits = digits.zfill(8)[-8:]
    return digits[:5]


def cep_comp(value: Any) -> Optional[str]:
    digits = _only_digits(value)
    if not digits:
        return None
    digits = digits.zfill(8)[-8:]
    return digits[5:8]


def sexo_code(value: Any) -> Optional[int]:
    text = clean_text(value)
    if text is None:
        return None
    return cfg.SEXO_MAP.get(text.upper())


def estado_civil_code(value: Any) -> Optional[int]:
    text = clean_text(value)
    if text is None:
        return None
    return cfg.ESTADO_CIVIL_MAP.get(text.upper())


def sim_nao(value: Any) -> Optional[str]:
    text = clean_text(value)
    if text is None:
        return None
    text = text.upper()
    if text in ("S", "SIM", "Y", "YES", "1", "TRUE"):
        return "S"
    if text in ("N", "NAO", "NÃO", "0", "FALSE"):
        return "N"
    return text[:1]


# Handler especial de telefone: recebe o REGISTRO inteiro (não só um campo),
# pois precisa combinar DDD + telefone.
def _telefone_com_ddd(record: Dict[str, Any]) -> Optional[str]:
    ddd = _only_digits(record.get("endereco__DDD 1"))
    fone = _only_digits(record.get("endereco__TELEFONE 1"))
    if not fone:
        return None
    return f"{ddd}{fone}" if ddd else fone


SPECIAL_HANDLERS = {
    "cpf_base": lambda record, raw: cpf_base(raw),
    "cpf_control": lambda record, raw: cpf_control(raw),
    "cep_base": lambda record, raw: cep_base(raw),
    "cep_comp": lambda record, raw: cep_comp(raw),
    "data_ddmmyyyy": lambda record, raw: data_ddmmyyyy(raw),
    "sexo_code": lambda record, raw: sexo_code(raw),
    "estado_civil_code": lambda record, raw: estado_civil_code(raw),
    "sim_nao": lambda record, raw: sim_nao(raw),
    "telefone_com_ddd": lambda record, raw: _telefone_com_ddd(record),
}


# =========================================================================
# 2. LEITURA E UNIFICAÇÃO DA BASE DE ORIGEM (PC ELEA)
# =========================================================================

def _read_source_sheet(path: str, sheet_key: str) -> pd.DataFrame:
    """Lê uma aba de origem respeitando a linha de cabeçalho configurada."""
    sheet_name = cfg.SRC_SHEETS[sheet_key]
    header_row = cfg.SRC_HEADER_ROW[sheet_key]  # 1-based
    df = pd.read_excel(
        path,
        sheet_name=sheet_name,
        header=header_row - 1,  # pandas é 0-based
        dtype=str,  # mantém tudo como texto; conversões ficam explícitas
    )
    df.columns = [str(c).strip() for c in df.columns]
    # remove linhas totalmente vazias
    df = df.dropna(how="all")
    return df


def load_source_data(path: str) -> Dict[str, pd.DataFrame]:
    """Carrega todas as abas relevantes da base PC ELEA em DataFrames."""
    data = {}
    for key in cfg.SRC_SHEETS:
        df = _read_source_sheet(path, key)
        data[key] = df
        logger.info("Origem: aba '%s' lida com %d linhas.", cfg.SRC_SHEETS[key], len(df))
    return data


def build_titular_records(source: Dict[str, pd.DataFrame]) -> List[Dict[str, Any]]:
    """
    Une TITULAR + DOCS + DADOS PROF. + DADOS BANCO + DADOS PLANO + ENDERECO
    em uma lista de registros (um dicionário por titular), na MESMA ORDEM
    em que aparecem na aba TITULAR de origem.

    Cada campo do resultado fica com o prefixo "<aba>__<coluna original>",
    exatamente como referenciado em mapping_config.py.
    """
    key = cfg.TITULAR_JOIN_KEY
    base = source["titular"].copy()
    base = base[base[key].notna()].reset_index(drop=True)

    lookups = {}
    for sheet_key in ("docs", "prof", "banco", "plano", "endereco"):
        df = source[sheet_key]
        df = df[df[key].notna()]
        # em caso de duplicidade de chave, mantém o primeiro registro
        lookups[sheet_key] = df.drop_duplicates(subset=[key], keep="first").set_index(key)

    records = []
    for _, row in base.iterrows():
        cert = row[key]
        record: Dict[str, Any] = {f"titular__{col}": row[col] for col in base.columns}
        for sheet_key, lk in lookups.items():
            if cert in lk.index:
                extra = lk.loc[cert]
                for col in lk.columns:
                    record[f"{sheet_key}__{col}"] = extra[col]
            else:
                for col in lk.columns:
                    record[f"{sheet_key}__{col}"] = None
        records.append(record)
    return records


def build_dependente_records(source: Dict[str, pd.DataFrame]) -> List[Dict[str, Any]]:
    """
    A aba DEPENDENTES já é autossuficiente (traz plano, região, carência do
    próprio dependente), então cada linha vira 1 registro, na mesma ordem
    da planilha de origem.
    """
    df = source["dependentes"].copy()
    key_cols = cfg.DEPENDENTE_JOIN_KEY
    df = df[df[key_cols[0]].notna() & df[key_cols[1]].notna()].reset_index(drop=True)

    records = []
    for _, row in df.iterrows():
        record = {f"dep__{col}": row[col] for col in df.columns}
        records.append(record)
    return records


# =========================================================================
# 3. ESCRITA NO TEMPLATE DE DESTINO (openpyxl, célula a célula)
# =========================================================================

def _header_map(ws: Worksheet, header_row: int) -> Dict[str, int]:
    """Retorna {nome_da_coluna: índice_da_coluna(1-based)} lendo o cabeçalho."""
    mapping = {}
    for col_idx in range(1, ws.max_column + 1):
        header = ws.cell(row=header_row, column=col_idx).value
        if header is not None:
            mapping[str(header).strip()] = col_idx
    return mapping


def _copy_style(src_cell, dst_cell) -> None:
    dst_cell.font = copy.copy(src_cell.font)
    dst_cell.fill = copy.copy(src_cell.fill)
    dst_cell.border = copy.copy(src_cell.border)
    dst_cell.alignment = copy.copy(src_cell.alignment)
    dst_cell.number_format = src_cell.number_format
    dst_cell.protection = copy.copy(src_cell.protection)


def _resolve_value(record: Dict[str, Any], sheet_name: str, dst_col: str, source_key: str) -> Any:
    """Aplica o handler especial (se houver) ou retorna o valor limpo direto."""
    handler_name = cfg.SPECIAL_COLUMNS.get(sheet_name, {}).get(dst_col)
    raw = record.get(source_key)
    if handler_name:
        handler = SPECIAL_HANDLERS[handler_name]
        return handler(record, raw)
    # Sem handler especial: texto limpo (strip) por padrão.
    return clean_text(raw)


def fill_sheet(
    ws: Worksheet,
    sheet_name: str,
    records: List[Dict[str, Any]],
) -> int:
    """
    Preenche uma aba do template com os registros fornecidos, célula a
    célula, preservando a formatação da linha-modelo (DST_TEMPLATE_ROW)
    para todas as colunas não mapeadas e para linhas extras necessárias.

    Retorna a quantidade de linhas efetivamente escritas.
    """
    columns_map = cfg.SHEET_CONFIG[sheet_name]["columns"]
    header_idx = _header_map(ws, cfg.DST_HEADER_ROW)
    data_start = cfg.DST_DATA_START_ROW
    template_row = cfg.DST_TEMPLATE_ROW
    max_col = ws.max_column

    # Snapshot da linha-modelo (valor + estilo) ANTES de começarmos a
    # sobrescrevê-la (a linha-modelo é a própria primeira linha de dados).
    template_snapshot = []
    for col_idx in range(1, max_col + 1):
        cell = ws.cell(row=template_row, column=col_idx)
        template_snapshot.append((cell.value, cell))

    original_last_row = ws.max_row

    for i, record in enumerate(records):
        dest_row = data_start + i

        # Normaliza SEMPRE a linha (estilo + valor constante) a partir da
        # linha-modelo, mesmo para linhas já existentes no bloco
        # pré-preenchido do template. Isso é necessário porque o template
        # de origem apresenta pequenas inconsistências de formatação em
        # algumas linhas (ex.: colunas de data com number_format real de
        # data, como 'ddmmyyyy', em vez do padrão numérico '00000000'
        # usado na linha-modelo) — sem essa normalização, o Excel poderia
        # exibir valores de data absurdos nessas linhas específicas.
        if dest_row != template_row:
            for col_idx in range(1, max_col + 1):
                tmpl_value, tmpl_cell = template_snapshot[col_idx - 1]
                dst_cell = ws.cell(row=dest_row, column=col_idx)
                _copy_style(tmpl_cell, dst_cell)
                dst_cell.value = tmpl_value

        # Aplica os campos variáveis (mapeados) por cima dos valores padrão.
        for dst_col, source_key in columns_map.items():
            if dst_col not in header_idx:
                logger.warning(
                    "Aba '%s': coluna de destino '%s' não encontrada no "
                    "cabeçalho do template — verifique mapping_config.py.",
                    sheet_name, dst_col,
                )
                continue
            value = _resolve_value(record, sheet_name, dst_col, source_key)
            ws.cell(row=dest_row, column=header_idx[dst_col]).value = value

    # Limpa linhas sobressalentes do bloco pré-preenchido que ficaram sem
    # registro correspondente (evita "linhas fantasmas" no arquivo final).
    last_needed_row = data_start + len(records) - 1
    for row_idx in range(last_needed_row + 1, original_last_row + 1):
        for col_idx in range(1, max_col + 1):
            ws.cell(row=row_idx, column=col_idx).value = None

    return len(records)


# =========================================================================
# 4. ORQUESTRAÇÃO
# =========================================================================

def run_etl(source_path: str, template_path: str, output_path: str) -> Dict[str, int]:
    """
    Executa o processo completo: lê a origem, unifica os registros, abre o
    template de destino e preenche todas as abas configuradas.

    Retorna um resumo {nome_da_aba: linhas_processadas} para logging/UI.
    """
    summary: Dict[str, int] = {}

    logger.info("Lendo base de origem: %s", source_path)
    source = load_source_data(source_path)

    logger.info("Unificando registros de TITULAR...")
    titular_records = build_titular_records(source)
    logger.info("Total de titulares unificados: %d", len(titular_records))

    logger.info("Unificando registros de DEPENDENTE...")
    dependente_records = build_dependente_records(source)
    logger.info("Total de dependentes unificados: %d", len(dependente_records))

    families = {
        "titular": titular_records,
        "dependente": dependente_records,
    }

    logger.info("Abrindo template de destino: %s", template_path)
    wb = load_workbook(template_path)

    for sheet_name in cfg.DST_SHEETS:
        try:
            if sheet_name not in wb.sheetnames:
                logger.error("Aba '%s' não encontrada no template — pulando.", sheet_name)
                continue
            ws = wb[sheet_name]
            family = cfg.SHEET_CONFIG[sheet_name]["family"]
            records = families[family]
            n = fill_sheet(ws, sheet_name, records)
            summary[sheet_name] = n
            logger.info("Aba '%s': %d linhas processadas com sucesso.", sheet_name, n)
        except Exception:
            logger.exception("Falha ao processar a aba '%s'.", sheet_name)
            summary[sheet_name] = 0

    logger.info("Salvando arquivo final: %s", output_path)
    wb.save(output_path)
    logger.info("Concluído.")
    return summary


# =========================================================================
# 5. CLI
# =========================================================================

def _parse_args():
    parser = argparse.ArgumentParser(description="ETL PC ELEA -> Reimplantação ELEA")
    parser.add_argument("--origem", required=True, help="Caminho do arquivo PC ELEA (origem)")
    parser.add_argument("--template", required=True, help="Caminho do template em branco (destino)")
    parser.add_argument("--saida", required=True, help="Caminho do arquivo .xlsx a ser gerado")
    return parser.parse_args()


if __name__ == "__main__":
    args = _parse_args()
    resumo = run_etl(args.origem, args.template, args.saida)
    print("\nResumo do processamento:")
    for aba, qtd in resumo.items():
        print(f"  - {aba}: {qtd} linha(s)")
