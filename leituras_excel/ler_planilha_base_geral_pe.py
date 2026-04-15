import pandas as pd

from configuracoes.parametros_processo import (
    CAMINHO_BASE_GERAL_PE,
    ABAS_IGNORADAS_BASE_PE,
    COLUNAS_PERMITIDAS
)

from utilitarios_excel.funcoes_excel import (
    listar_abas,
    ler_planilha,
    validar_colunas
)


# ===============================
# 📥 CARREGAR BASE COMPLETA
# ===============================

def carregar_base_geral_pe():
    """
    Carrega todas as abas válidas da Base Geral PE
    e retorna um único DataFrame consolidado.
    """

    abas = listar_abas(CAMINHO_BASE_GERAL_PE)

    dfs = []

    for aba in abas:
        if aba in ABAS_IGNORADAS_BASE_PE:
            continue

        try:
            df = ler_planilha(CAMINHO_BASE_GERAL_PE, aba)

            # ===============================
            # 📊 Padronizar colunas
            # ===============================
            df.columns = [col.strip().upper() for col in df.columns]

            # ===============================
            # ✔ Validar estrutura
            # ===============================
            validar_colunas(df, COLUNAS_PERMITIDAS)

            # ===============================
            # 📌 Adicionar coluna de origem
            # ===============================
            df["ABA_ORIGEM"] = aba

            dfs.append(df)

        except Exception as e:
            print(f"Erro ao ler aba '{aba}': {e}")

    if not dfs:
        raise ValueError("Nenhuma aba válida encontrada na Base Geral PE")

    df_final = pd.concat(dfs, ignore_index=True)

    return df_final