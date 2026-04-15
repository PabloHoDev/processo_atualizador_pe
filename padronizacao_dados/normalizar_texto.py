# Maiúsculo, trim, limpeza
import pandas as pd
import re


# ===============================
# 🔤 NORMALIZAR TEXTO INDIVIDUAL
# ===============================

def normalizar_texto(valor):
    """
    Normaliza um valor textual:
    - Remove espaços extras
    - Converte para maiúsculo
    - Remove múltiplos espaços internos
    - Trata valores nulos
    """
    if pd.isna(valor):
        return ""

    valor = str(valor)

    # Remove espaços nas extremidades
    valor = valor.strip()

    # Remove múltiplos espaços internos
    valor = re.sub(r"\s+", " ", valor)

    # Converte para maiúsculo
    valor = valor.upper()

    return valor


# ===============================
# 📊 NORMALIZAR COLUNAS DO DF
# ===============================

def normalizar_colunas_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normaliza apenas colunas do tipo texto (object)
    """
    df = df.copy()

    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].apply(normalizar_texto)

    return df


# ===============================
# 🎯 NORMALIZAR COLUNAS CHAVE
# ===============================

def normalizar_colunas_chave(df: pd.DataFrame, colunas_chave: list) -> pd.DataFrame:
    """
    Normaliza colunas específicas consideradas críticas para o processo
    """
    df = df.copy()

    for col in colunas_chave:
        if col in df.columns:
            df[col] = df[col].apply(normalizar_texto)

    return df


# ===============================
# 🚀 NORMALIZAÇÃO COMPLETA
# ===============================

def normalizar_dados(df: pd.DataFrame, colunas_chave: list) -> pd.DataFrame:
    """
    Pipeline completo de normalização
    """
    df = normalizar_colunas_dataframe(df)
    df = normalizar_colunas_chave(df, colunas_chave)

    return df