# -*- coding: utf-8 -*-
"""
mapping_config.py
==================
Configuração central do sistema de transferência (PC ELEA / PC PIEMONTE /
outras bases "parecidas") -> Reimplantação ELEA.

------------------------------------------------------------------------
COMO FUNCIONA (leia antes de editar)
------------------------------------------------------------------------
Bases de origem diferentes (ex.: "PC ELEA", "PC PIEMONTE") trazem os MESMOS
tipos de informação (nome, CPF, data de nascimento, dependentes...) só que
organizados em abas e colunas com nomes diferentes — e às vezes com
problemas de acentuação no próprio arquivo.

Em vez de mapear "aba X, coluna Y" (o que quebra a cada arquivo novo com
layout diferente), o sistema usa CAMPOS CANÔNICOS: um nome lógico único
para cada informação (ex.: "CPF", "DATA_NASCIMENTO"), e uma lista de
possíveis apelidos (aliases) usados para reconhecer esse campo em qualquer
aba/coluna da planilha de origem, seja qual for o nome exato usado nela.

O motor de ETL (etl_reimplantacao.py):
  1. varre todas as abas da planilha de origem;
  2. identifica automaticamente a linha de cabeçalho de cada aba;
  3. compara (de forma tolerante a acento/abreviação/erro de encoding)
     cada cabeçalho encontrado com os apelidos abaixo;
  4. monta um registro "unificado" por titular e por dependente com os
     campos canônicos preenchidos, não importa em qual aba cada um estava.

Se uma nova base de origem chegar com uma coluna que os apelidos abaixo
não reconhecem, o sistema NÃO quebra: ele registra um aviso no log e deixa
o campo em branco no destino. Para "ensinar" o sistema a reconhecer essa
coluna nova, basta adicionar o texto dela na lista de aliases do campo
correspondente, abaixo — não é necessário mexer no motor de ETL.
"""

# =============================================================================
# 1. CAMPOS CANÔNICOS — TITULAR
# =============================================================================
# Cada campo canônico tem uma lista de "apelidos" (como esse dado costuma
# aparecer nos cabeçalhos das bases de origem). Ordem importa: o primeiro
# apelido encontrado tem prioridade (ex.: preferimos a versão "(Y2K)" de
# datas, que já vem com ano de 4 dígitos, sobre a versão curta).

TITULAR_FIELD_ALIASES = {
    "CERTIFICADO": ["NUMERO DO CERTIFICADO", "NUMERO CERTIFICADO", "N CERTIFICADO", "N CERT"],
    "SUBFATURA": ["NUMERO DA SUBFATURA", "N SUBFATURA", "N SUB"],
    "MATRICULA_ESPECIAL": ["NUMERO DA MATRICULA", "N DA MATRICULA", "MATRICULA ESPECIAL"],
    "NOME_SEGURADO": ["NOME DO SEGURADO", "NOME SEGURADO"],
    "NOME_MAE": ["NOME DA MAE", "NOME MAE"],
    "DATA_NASCIMENTO": ["DATA DE NASCIMENTO Y2K", "DATA NASCIMENTO Y2K", "DATA DE NASCIMENTO", "DATA NASCIMENTO"],
    "SEXO": ["SEXO DO SEGURADO", "SEXO SEGURADO", "SEXO"],
    "ESTADO_CIVIL": ["ESTADO CIVIL"],
    "CPF": ["NUMERO DO CPF", "NUMERO CPF", "CPF"],
    "CARGO_OCUPACAO": ["CARGO DE OCUPACAO", "CARGO OCUPACAO", "CARGO"],
    "DATA_ADMISSAO": ["DATA DE ADMISSAO Y2K", "DATA ADMISSAO Y2K", "DATA DE ADMISSAO", "DATA ADMISSAO"],
    "MATRICULA_FUNCIONAL": ["MATRICULA FUNCIONAL"],
    "VINCULO_TRABALHO": ["VINCULO DE TRABALHO", "VINCULO TRABALHO"],
    "NOME_RESPONSAVEL": ["NOME RESPONSAVEL"],
    "PLANO": ["PLANO"],
    "REGIAO": ["REGIAO"],
    "CODIGO_CARENCIA": ["CODIGO DE CARENCIA", "CODIGO CARENCIA"],
    "LOGRADOURO": ["LOGRADOURO"],
    "BAIRRO": ["BAIRRO"],
    "CIDADE": ["CIDADE"],
    "UF": ["UF"],
    "CEP": ["CEP"],
    "DDD1": ["DDD 1", "DDD1", "DDD"],
    "TELEFONE1": ["TELEFONE 1", "TELEFONE1", "TELEFONE"],
}

# Campo(s) usados para identificar se uma aba é "nível titular" (precisa ter
# a coluna de certificado) e qual aba é a "principal" (a que tem o nome do
# segurado — usada para definir a ordem/quantidade de registros).
TITULAR_KEY_FIELD = "CERTIFICADO"
TITULAR_PRIMARY_HINT_FIELD = "NOME_SEGURADO"


# =============================================================================
# 2. CAMPOS CANÔNICOS — DEPENDENTE
# =============================================================================

DEPENDENTE_FIELD_ALIASES = {
    "CERTIFICADO": ["NUMERO DO CERTIFICADO", "NUMERO CERTIFICADO", "N CERTIFICADO", "N CERT"],
    "CODIGO_DEPENDENTE": ["CODIGO DO DEPENDENTE", "CODIGO DEPENDENTE"],
    "MATRICULA_ESPECIAL": [
        "NUMERO DA MATRICULA DO DEPENDENTE", "NUMERO DA MATRICULA DEPENDENTE", "NUMERO DA MATRICULA",
        "MATRICULA ESPECIAL DO DEPENDENTE", "MATRICULA ESPECIAL DEPENDENTE", "MATRICULA ESPECIAL",
    ],
    "NOME_DEPENDENTE": ["NOME DO DEPENDENTE", "NOME DEPENDENTE"],
    "NOME_MAE_DEPENDENTE": ["NOME DA MAE DO DEPENDENTE", "NOME DA MAE DEPENDENTE", "NOME MAE DEPENDENTE"],
    "DATA_NASCIMENTO": ["DATA DE NASCIMENTO Y2K", "DATA NASCIMENTO Y2K", "DATA DE NASCIMENTO", "DATA NASCIMENTO"],
    "SEXO": ["SEXO DO DEPENDENTE", "SEXO DEPENDENTE", "SEXO"],
    "ESTADO_CIVIL": ["ESTADO CIVIL DO DEPENDENTE", "ESTADO CIVIL DEPENDENTE", "ESTADO CIVIL"],
    "GRAU_PARENTESCO": ["GRAU DE PARENTESCO", "GRAU PARENTESCO"],
    "CPF": ["NUMERO DO CPF DO DEPENDENTE", "NUMERO DO CPF", "NUMERO CPF", "CPF"],
    "FILHO_INVALIDEZ": ["FILHO COM INVALIDEZ", "FILHO INVALIDEZ"],
}

DEPENDENTE_KEY_FIELDS = ["CERTIFICADO", "CODIGO_DEPENDENTE"]
# Campo usado para identificar qual aba é a de dependentes.
DEPENDENTE_ROLE_HINT_FIELD = "CODIGO_DEPENDENTE"


# =============================================================================
# 3. TABELAS DE CÓDIGO (TODO: validar com a operadora / área de negócio)
# =============================================================================
# A própria base de origem já traz o SEXO e o ESTADO CIVIL do DEPENDENTE
# como código numérico (1/2) em ambos os layouts observados (ELEA e
# PIEMONTE). O mapeamento abaixo replica essa mesma convenção também para
# o TITULAR (que em algumas bases vem como texto, ex. "Masculino").
# CONFIRME antes de usar em produção.
SEXO_MAP = {
    "MASCULINO": 1,
    "FEMININO": 2,
    "M": 1,
    "F": 2,
    "1": 1,
    "01": 1,
    "2": 2,
    "02": 2,
}

ESTADO_CIVIL_MAP = {
    "SOLTEIRO(A)": 1,
    "CASADO(A)": 2,
    "SEPARADO(A)": 3,
    "DESQUITADO(A)": 3,
    "DIVORCIADO(A)": 4,
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
# 4. TEMPLATE DE DESTINO
# =============================================================================

DST_SHEETS = [
    "Titular",
    "Tit Novo Nome",
    "Endereço",
    "Plano",
    "Angariador",
    "Dependente",
    "Dep Novo Nome",
]

DST_HEADER_ROW = 1
DST_DATA_START_ROW = 3
DST_TEMPLATE_ROW = 3  # linha usada como modelo de estilo/valores constantes


# =============================================================================
# 5. MAPEAMENTO DE COLUNAS POR ABA DE DESTINO
# =============================================================================
# "Nome da coluna no destino": "CAMPO_CANONICO"
#
# Qualquer coluna do destino que NÃO apareça aqui é deixada como está no
# template (mantém o valor constante já pré-preenchido: Tipo Mov, Tipo Doc,
# CIA, Contrato, N° Sub, Tipo Reg, Data Inc Vig, fillers, etc.)

TITULAR_COLUMNS = {
    "N° Cert": "CERTIFICADO",
    "Matricula Esp": "MATRICULA_ESPECIAL",
    "Data Nasc": "DATA_NASCIMENTO",
    "Mat Func": "MATRICULA_FUNCIONAL",
    "Data Adm": "DATA_ADMISSAO",
    "Sexo": "SEXO",
    "Estado Civil": "ESTADO_CIVIL",
    "CPF": "CPF",              # tratado especial (split)
    "CPF Controle": "CPF",     # tratado especial (split)
    "Cargo Ocup": "CARGO_OCUPACAO",
    "Nome Resp": "NOME_RESPONSAVEL",
    "Vinculo Trab": "VINCULO_TRABALHO",
    # Sem correspondência garantida em todas as bases -> ficam em branco
    # quando a base de origem não tiver o dado: "Data Canc Futuro",
    # "Cod Cond Geral", "Peso", "Altura"
}

TIT_NOVO_NOME_COLUMNS = {
    "N° Cert": "CERTIFICADO",
    "Matricula Esp": "MATRICULA_ESPECIAL",
    "Nome Seg": "NOME_SEGURADO",
    "Nome Mae Tit": "NOME_MAE",
}

ENDERECO_COLUMNS = {
    "N° Cert": "CERTIFICADO",
    "Matricula Esp": "MATRICULA_ESPECIAL",
    "Logradouro": "LOGRADOURO",
    "Bairro": "BAIRRO",
    "Cidade": "CIDADE",
    "UF": "UF",
    "Cep": "CEP",         # tratado especial (split 5+3)
    "CepComp": "CEP",     # tratado especial (split 5+3)
    "TelDDD1": "DDD1",
    "Tel Celular": "TELEFONE1",  # combinado com DDD1 (tratado especial)
}

PLANO_COLUMNS = {
    "N° Cert": "CERTIFICADO",
    "Matricula Esp": "MATRICULA_ESPECIAL",
    "Cod Plano": "PLANO",
    "Cod Região": "REGIAO",
}

ANGARIADOR_COLUMNS = {
    "N° Cert": "CERTIFICADO",
    "Matricula Esp": "MATRICULA_ESPECIAL",
    "Cod Carencia": "CODIGO_CARENCIA",
}

DEPENDENTE_COLUMNS = {
    "N° Cert": "CERTIFICADO",
    "Matricula Esp": "MATRICULA_ESPECIAL",
    "Cod Dependente": "CODIGO_DEPENDENTE",
    "Data Nascimento": "DATA_NASCIMENTO",
    "Sexo": "SEXO",
    "Estado Civil": "ESTADO_CIVIL",
    "Parentesco": "GRAU_PARENTESCO",
    "CPF": "CPF",              # split
    "CPF Controle": "CPF",     # split
    "Matr Esp Dep": "MATRICULA_ESPECIAL",
    "Filho Invalid": "FILHO_INVALIDEZ",
}

DEP_NOVO_NOME_COLUMNS = {
    "N° Cert": "CERTIFICADO",
    "Matricula Esp": "MATRICULA_ESPECIAL",
    "Cod Dependente": "CODIGO_DEPENDENTE",
    "Nome Dependente": "NOME_DEPENDENTE",
    "Nome Mae Dep": "NOME_MAE_DEPENDENTE",
}

# Para cada aba de destino: a "família" de registro que ela usa
# ("titular" ou "dependente") e o mapeamento de colunas.
SHEET_CONFIG = {
    "Titular":        {"family": "titular",    "columns": TITULAR_COLUMNS},
    "Tit Novo Nome":  {"family": "titular",    "columns": TIT_NOVO_NOME_COLUMNS},
    "Endereço":       {"family": "titular",    "columns": ENDERECO_COLUMNS},
    "Plano":          {"family": "titular",    "columns": PLANO_COLUMNS},
    "Angariador":     {"family": "titular",    "columns": ANGARIADOR_COLUMNS},
    "Dependente":     {"family": "dependente", "columns": DEPENDENTE_COLUMNS},
    "Dep Novo Nome":  {"family": "dependente", "columns": DEP_NOVO_NOME_COLUMNS},
}

# Campos canônicos que exigem tratamento especial (split / combinação) em
# vez de cópia direta do valor. O handler correspondente mora em
# etl_reimplantacao.py (SPECIAL_HANDLERS).
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
        "Sexo": "sexo_code",
        "Estado Civil": "estado_civil_code",
    },
}
