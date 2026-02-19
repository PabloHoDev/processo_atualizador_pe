# Caminhos, nomes de abas, colunas e regras

"""
parametros_processo.py

Arquivo central de configurações do processo
processo_novas_obrigacoes_rpe
"""

# ===============================
# 📂 CAMINHOS DOS ARQUIVOS
# ===============================

CAMINHO_RPE = r"caminho/para/arquivo/RPE.xlsx"
CAMINHO_BASE_GERAL_PE = r"caminho/para/arquivo/Base_Geral_PE.xlsx"

# ===============================
# 📑 NOMES DAS ABAS
# ===============================

ABA_RPE = "RPE"

# Abas que NÃO devem ser consideradas na Base Geral PE
ABAS_IGNORADAS_BASE_PE = [
    "DADOS DE VALIDAÇÃO",
    "META INTELIGENTE",
    "LAKE",
]

# ===============================
# 📌 COLUNAS OFICIAIS DO PROCESSO
# ===============================

COLUNAS_PERMITIDAS = [
    "USUÁRIO",
    "PROCEDIMENTO",
    "NOME PRESTADOR",
    "UF PRESTADOR",
    "CIDADE PRESTADOR",
    "VALOR DEPÓSITO",
    "DATA DEPÓSITO",
    "OBRIGAÇÃO",
    "ÁREA DA PENDÊNCIA",
    "STATUS",
]

COLUNA_CHAVE = "OBRIGAÇÃO"

# ===============================
# 📏 CONFIGURAÇÕES DE VALIDAÇÃO
# ===============================

COLUNAS_OBRIGATORIAS = [
    "OBRIGAÇÃO",
    "UF PRESTADOR",
]

VALIDAR_DATA_DEPOSITO = True
VALIDAR_VALOR_DEPOSITO = True

# ===============================
# 🧾 CONFIGURAÇÕES DE LOG
# ===============================

GERAR_LOG_ARQUIVO = True
NOME_ARQUIVO_LOG = "log_processo_novas_obrigacoes.csv"

# ===============================
# ⚙ CONFIGURAÇÕES DE EXECUÇÃO
# ===============================

MODO_DEBUG = True
PERMITIR_REEXECUCAO_SEM_ERRO = True
