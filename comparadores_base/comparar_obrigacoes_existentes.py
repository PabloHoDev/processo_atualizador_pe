import pandas as pd
from configuracoes.parametros_processo import COLUNA_CHAVE


# ===============================
# 🔍 EXTRAIR OBRIGAÇÕES EXISTENTES
# ===============================

def extrair_obrigacoes_base(df_base: pd.DataFrame) -> set:
    """
    Retorna um conjunto com todas as obrigações existentes na base.
    """
    if COLUNA_CHAVE not in df_base.columns:
        raise ValueError(f"Coluna chave '{COLUNA_CHAVE}' não encontrada na base")

    obrigacoes = set(df_base[COLUNA_CHAVE].dropna().unique())

    return obrigacoes


# ===============================
# 🚀 COMPARAÇÃO PRINCIPAL
# ===============================

def comparar_obrigacoes(df_rpe: pd.DataFrame, df_base: pd.DataFrame) -> dict:
    """
    Compara RPE com Base Geral PE e separa registros novos e existentes.
    """

    obrigacoes_base = extrair_obrigacoes_base(df_base)

    novos = []
    existentes = []

    for _, linha in df_rpe.iterrows():
        obrigacao = linha.get(COLUNA_CHAVE, "")

        if obrigacao in obrigacoes_base:
            existentes.append(linha)
        else:
            novos.append(linha)

    df_novos = pd.DataFrame(novos)
    df_existentes = pd.DataFrame(existentes)

    return {
        "novos": df_novos,
        "existentes": df_existentes,
        "total_rpe": len(df_rpe),
        "total_novos": len(df_novos),
        "total_existentes": len(df_existentes)
    }