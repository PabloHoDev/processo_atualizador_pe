from pathlib import Path
import pandas as pd


# ===============================
# 📂 VERIFICAÇÃO DE ARQUIVO
# ===============================

def verificar_arquivo_existe(caminho: str) -> None:
    """
    Verifica se o arquivo existe no caminho informado.
    """
    if not Path(caminho).exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")


# ===============================
# 📑 LISTAR ABAS
# ===============================

def listar_abas(caminho_arquivo: str) -> list:
    """
    Retorna a lista de abas de um arquivo Excel.
    """
    verificar_arquivo_existe(caminho_arquivo)
    
    with pd.ExcelFile(caminho_arquivo) as xls:
        return xls.sheet_names


# ===============================
# 📄 VERIFICAR ABA
# ===============================

def verificar_aba_existe(caminho_arquivo: str, nome_aba: str) -> None:
    """
    Verifica se a aba existe no arquivo Excel.
    """
    abas = listar_abas(caminho_arquivo)
    
    if nome_aba not in abas:
        raise ValueError(f"Aba '{nome_aba}' não encontrada no arquivo.")


# ===============================
# 📥 LEITURA DE PLANILHA
# ===============================

def ler_planilha(caminho_arquivo: str, nome_aba: str) -> pd.DataFrame:
    """
    Lê uma aba específica de um arquivo Excel.
    """
    verificar_arquivo_existe(caminho_arquivo)
    verificar_aba_existe(caminho_arquivo, nome_aba)
    
    df = pd.read_excel(caminho_arquivo, sheet_name=nome_aba)
    
    return df


# ===============================
# 📊 VALIDAR COLUNAS
# ===============================

def validar_colunas(df: pd.DataFrame, colunas_esperadas: list) -> None:
    """
    Verifica se todas as colunas esperadas existem no DataFrame.
    """
    colunas_df = [col.upper() for col in df.columns]
    
    faltantes = [col for col in colunas_esperadas if col.upper() not in colunas_df]
    
    if faltantes:
        raise ValueError(f"Colunas ausentes: {faltantes}")


# ===============================
# 📤 SALVAR PLANILHA
# ===============================

def salvar_planilha(df: pd.DataFrame, caminho_arquivo: str, nome_aba: str) -> None:
    """
    Salva um DataFrame em uma aba de um arquivo Excel.
    Se o arquivo existir, sobrescreve a aba.
    """
    with pd.ExcelWriter(caminho_arquivo, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=nome_aba, index=False)