# -*- coding: utf-8 -*-
"""
mapping_config.py
==================
Configuração central do sistema de transferência PC ELEA -> Reimplantação ELEA.

Este arquivo concentra TUDO que costuma mudar de um lote para outro:
  - nomes das abas de origem e destino
  - de qual coluna de origem vem cada coluna de destino
  - tabelas de código (sexo, estado civil) usadas para converter texto -> código
  - qual é a "chave" usada para casar (JOIN) as abas de origem entre si

Se os nomes de colunas/abas mudarem em um próximo lote, ajuste APENAS este
arquivo — o motor de ETL (etl_reimplantacao.py) não precisa ser tocado.

>>> IMPORTANTE - PONTOS QUE EXIGEM VALIDAÇÃO DE NEGÓCIO (marcados como TODO) <<<
Alguns campos do layout de destino são campos codificados (ex.: Sexo, Estado
Civil) e a base de origem traz esses dados como texto ("Masculino",
"Casado(a)") ou já como código numérico, dependendo da aba. Foi necessário
assumir uma tabela de códigos com base no padrão observado nos próprios dados
(ver seção CODE_MAPS). CONFIRME essas tabelas com a área de negócio/operadora
antes de usar o arquivo gerado em produção.
"""

# =============================================================================
# 1. NOMES DAS ABAS
# =============================================================================

SRC_SHEETS = {
    "titular": "POS. CADASTRAL (TITULAR)",
    "docs": "POS. CADASTRAL (DOCS)",
    "prof": "POS. CADASTRAL (DADOS PROF.)",
    "banco": "POS. CADASTRAL (DADOS BANCO)",
    "plano": "POS. CADASTRAL (DADOS PLANO)",
    "endereco": "POS. CADASTRAL (ENDERECO)",
    "dependentes": "POS. CADASTRAL (DEPENDENTES)",
}

# Linha (1-based) onde está o cabeçalho de cada aba de origem.
# A aba TITULAR tem 2 linhas de cabeçalho de arquivo antes do cabeçalho de
# colunas propriamente dito; as demais abas já começam o cabeçalho na linha 1.
SRC_HEADER_ROW = {
    "titular": 3,
    "docs": 1,
    "prof": 1,
    "banco": 1,
    "plano": 1,
    "endereco": 1,
    "dependentes": 1,
}

DST_SHEETS = [
    "Titular",
    "Tit Novo Nome",
    "Endereço",
    "Plano",
    "Angariador",
    "Dependente",
    "Dep Novo Nome",
]

# Linha do cabeçalho e linha onde começam os dados no template de destino.
# O template já vem com um "bloco" de linhas pré-preenchidas (linhas 3 em
# diante) contendo os valores constantes do lote (Tipo Mov, CIA, Contrato,
# Tipo Reg, etc.). A linha 3 é usada como "linha-modelo": tudo que não for
# explicitamente mapeado abaixo é copiado dela (valor + formatação) para cada
# linha nova.
DST_HEADER_ROW = 1
DST_DATA_START_ROW = 3
DST_TEMPLATE_ROW = 3  # linha usada como modelo de estilo/valores constantes

# Chave usada para casar (JOIN) as abas de origem ao nível do TITULAR.
TITULAR_JOIN_KEY = "NÚMERO DO CERTIFICADO"

# Chave composta usada para casar registros ao nível do DEPENDENTE.
DEPENDENTE_JOIN_KEY = ["NÚMERO DO CERTIFICADO", "CÓDIGO DO DEPENDENTE"]


# =============================================================================
# 2. TABELAS DE CÓDIGO (TODO: validar com a operadora / área de negócio)
# =============================================================================

# Observação: a própria base de origem já traz o SEXO e o ESTADO CIVIL do
# DEPENDENTE como código numérico (1/2). O mapeamento abaixo foi construído
# para ser compatível com essa mesma convenção, aplicando-a também ao TITULAR
# (que vem como texto). CONFIRME antes de usar em produção.
SEXO_MAP = {
    "MASCULINO": 1,
    "FEMININO": 2,
    "M": 1,
    "F": 2,
    "1": 1,
    "2": 2,
}

ESTADO_CIVIL_MAP = {
    "SOLTEIRO(A)": 1,
    "CASADO(A)": 2,
    "SEPARADO(A)": 3,
    "DESQUITADO(A)": 3,
    "DIVORCIADO(A)": 4,
    "VIÚVO(A)": 5,
    "VIUVO(A)": 5,
    "OUTROS": 6,
    "1": 1,
    "2": 2,
    "3": 3,
    "4": 4,
    "5": 5,
    "6": 6,
}


# =============================================================================
# 3. MAPEAMENTO DE COLUNAS POR ABA DE DESTINO
# =============================================================================
# Cada entrada é: "Nome da coluna no destino": "chave_no_registro_de_origem"
# (o "registro de origem" é um dicionário Python já unificado — ver etl.py).
#
# Qualquer coluna do destino que NÃO apareça aqui é deixada como está no
# template (ou seja, mantém o valor constante já pré-preenchido pela
# operadora: Tipo Mov, Tipo Doc, CIA, Contrato, N° Sub, Tipo Reg, Data Inc
# Vig, fillers, etc.) — exatamente como pedido: "não recriar, apenas inserir
# os dados nos campos variáveis".
#
# As chaves usadas aqui (ex.: "titular__NOME DO SEGURADO") seguem o padrão
# "<sheet_de_origem>__<coluna_original>" criado pelo etl.py ao unificar as
# abas de origem por titular.

TITULAR_COLUMNS = {
    "N° Cert": "titular__NÚMERO DO CERTIFICADO",
    "Matricula Esp": "titular__MATRÍCULA ESPECIAL",
    "Data Nasc": "titular__DATA DE NASCIMENTO",
    "Mat Func": "prof__MATRÍCULA FUNCIONAL",
    "Data Adm": "prof__DATA DE ADMISSÃO",
    "Sexo": "titular__SEXO DO SEGURADO",
    "Estado Civil": "titular__ESTADO CIVIL",
    "CPF": "titular__NÚMERO DO CPF",              # tratado especial (split)
    "CPF Controle": "titular__NÚMERO DO CPF",      # tratado especial (split)
    "Cargo Ocup": "prof__CARGO DE OCUPAÇÃO",
    "Nome Resp": "titular__NOME RESPONSÁVEL",
    "Vinculo Trab": "prof__VÍNCULO DE TRABALHO",
    # Sem correspondência clara na base de origem -> ficam em branco:
    # "Data Canc Futuro", "Cod Cond Geral", "Peso", "Altura"
}

TIT_NOVO_NOME_COLUMNS = {
    "N° Cert": "titular__NÚMERO DO CERTIFICADO",
    "Matricula Esp": "titular__MATRÍCULA ESPECIAL",
    "Nome Seg": "titular__NOME DO SEGURADO",
    "Nome Mae Tit": "titular__NOME DA MÃE",
}

ENDERECO_COLUMNS = {
    "N° Cert": "titular__NÚMERO DO CERTIFICADO",
    "Matricula Esp": "titular__MATRÍCULA ESPECIAL",
    "Logradouro": "endereco__LOGRADOURO",
    "Bairro": "endereco__BAIRRO",
    "Cidade": "endereco__CIDADE",
    "UF": "endereco__UF",
    "Cep": "endereco__CEP",         # tratado especial (split 5+3)
    "CepComp": "endereco__CEP",     # tratado especial (split 5+3)
    "TelDDD1": "endereco__DDD 1",
    "Tel Celular": "endereco__TELEFONE 1",  # combinado com DDD 1
    # "Tel Dado Adicional" sem correspondência direta -> em branco
}

PLANO_COLUMNS = {
    "N° Cert": "titular__NÚMERO DO CERTIFICADO",
    "Matricula Esp": "titular__MATRÍCULA ESPECIAL",
    "Cod Plano": "plano__PLANO",
    "Cod Região": "plano__REGIÃO",
    # "Tipo Reg", "Reemb Desp Não Comporv", "Vl Diaria", "Prazo Dias":
    # sem correspondência na base de origem -> em branco (TODO: confirmar
    # regra de negócio, pois "Tipo Reg" nesta aba não veio pré-preenchido
    # no template, ao contrário das demais abas)
}

ANGARIADOR_COLUMNS = {
    "N° Cert": "titular__NÚMERO DO CERTIFICADO",
    "Matricula Esp": "titular__MATRÍCULA ESPECIAL",
    "Cod Carencia": "plano__CÓDIGO DE CARÊNCIA",
    # "Cod Ag Angariador", "Preposto": sem fonte na base de origem
    # (dados de corretor/angariador não estão no PC ELEA) -> em branco
}

DEPENDENTE_COLUMNS = {
    "N° Cert": "dep__NÚMERO DO CERTIFICADO",
    "Matricula Esp": "dep__MATRÍCULA ESPECIAL",
    "Cod Dependente": "dep__CÓDIGO DO DEPENDENTE",
    "Data Nascimento": "dep__DATA DE NASCIMENTO",
    "Sexo": "dep__SEXO DO DEPENDENTE",
    "Estado Civil": "dep__ESTADO CIVIL DO DEPENDENTE",
    "Parentesco": "dep__GRAU DE PARENTESCO",
    "CPF": "dep__NÚMERO DO CPF DO DEPENDENTE",           # split
    "CPF Controle": "dep__NÚMERO DO CPF DO DEPENDENTE",   # split
    "Matr Esp Dep": "dep__MATRÍCULA ESPECIAL DEPENDENTE",
    "Filho Invalid": "dep__FILHO COM INVALIDEZ",
    # "Data Canc Agen", "Peso", "Altura", "Filho Universitário": em branco
}

DEP_NOVO_NOME_COLUMNS = {
    "N° Cert": "dep__NÚMERO DO CERTIFICADO",
    "Matricula Esp": "dep__MATRÍCULA ESPECIAL",
    "Cod Dependente": "dep__CÓDIGO DO DEPENDENTE",
    "Nome Dependente": "dep__NOME DO DEPENDENTE",
    "Nome Mae Dep": "dep__NOME DA MÃE DEPENDENTE",
}

# Agrupa tudo: para cada aba de destino, guarda o mapeamento de colunas e
# de qual "família" de registros de origem ela deve iterar ("titular" ou
# "dependente" — ver etl.py).
SHEET_CONFIG = {
    "Titular":        {"family": "titular",    "columns": TITULAR_COLUMNS},
    "Tit Novo Nome":  {"family": "titular",    "columns": TIT_NOVO_NOME_COLUMNS},
    "Endereço":       {"family": "titular",    "columns": ENDERECO_COLUMNS},
    "Plano":          {"family": "titular",    "columns": PLANO_COLUMNS},
    "Angariador":     {"family": "titular",    "columns": ANGARIADOR_COLUMNS},
    "Dependente":     {"family": "dependente", "columns": DEPENDENTE_COLUMNS},
    "Dep Novo Nome":  {"family": "dependente", "columns": DEP_NOVO_NOME_COLUMNS},
}

# Colunas que exigem tratamento especial (split / combinação) em vez de
# cópia direta de valor. A função correspondente mora em etl.py
# (SPECIAL_HANDLERS).
SPECIAL_COLUMNS = {
    "Titular": {
        "CPF": "cpf_base",
        "CPF Controle": "cpf_control",
        "Data Nasc": "data_ddmmyyyy",
        "Data Adm": "data_ddmmyyyy",
        "Sexo": "sexo_code",
        "Estado Civil": "estado_civil_code",
    },
    "Endereço": {
        "Cep": "cep_base",
        "CepComp": "cep_comp",
        "Tel Celular": "telefone_com_ddd",
    },
    "Dependente": {
        "CPF": "cpf_base",
        "CPF Controle": "cpf_control",
        "Data Nascimento": "data_ddmmyyyy",
        "Filho Invalid": "sim_nao",
    },
}
