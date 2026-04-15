# Leitura da planilha RPE

from configuracoes.parametros_processo import (
    CAMINHO_RPE,
    ABA_RPE,
    COLUNAS_PERMITIDAS
)

from utilitarios_excel.funcoes_excel import (
    ler_planilha,
    validar_colunas
)


def carregar_dados_rpe():
    """
    Carrega e valida a estrutura da planilha RPE.
    
    Retorna:
        pd.DataFrame: Dados da planilha RPE
    """

    # ===============================
    # 📥 Leitura da planilha
    # ===============================
    df = ler_planilha(CAMINHO_RPE, ABA_RPE)

    # ===============================
    # 📊 Padronizar nomes de colunas
    # ===============================
    df.columns = [col.strip().upper() for col in df.columns]

    # ===============================
    # ✔ Validar colunas obrigatórias
    # ===============================
    validar_colunas(df, COLUNAS_PERMITIDAS)

    return df