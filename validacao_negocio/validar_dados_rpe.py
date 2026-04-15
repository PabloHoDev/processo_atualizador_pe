import pandas as pd
from configuracoes.parametros_processo import (
    COLUNAS_OBRIGATORIAS,
    VALIDAR_DATA_DEPOSITO,
    VALIDAR_VALOR_DEPOSITO
)


# ===============================
# 🔍 VALIDAÇÃO DE LINHA
# ===============================

def validar_linha(linha: pd.Series) -> dict:
    """
    Valida uma única linha da RPE.
    """
    erros = []
    avisos = []

    # ===============================
    # ✔ Campos obrigatórios
    # ===============================
    for col in COLUNAS_OBRIGATORIAS:
        if col not in linha or str(linha[col]).strip() == "":
            erros.append(f"Campo obrigatório ausente: {col}")

    # ===============================
    # 📅 Validação de data
    # ===============================
    if VALIDAR_DATA_DEPOSITO and "DATA DEPÓSITO" in linha:
        valor = linha["DATA DEPÓSITO"]
        if pd.notna(valor):
            try:
                pd.to_datetime(valor)
            except Exception:
                erros.append("Data Depósito inválida")

    # ===============================
    # 💰 Validação de valor
    # ===============================
    if VALIDAR_VALOR_DEPOSITO and "VALOR DEPÓSITO" in linha:
        valor = linha["VALOR DEPÓSITO"]
        if pd.notna(valor):
            try:
                float(valor)
            except Exception:
                erros.append("Valor Depósito inválido")

    # ===============================
    # 📊 Classificação final
    # ===============================
    if erros:
        status = "ERRO"
    elif avisos:
        status = "AVISO"
    else:
        status = "VALIDO"

    return {
        "status": status,
        "erros": erros,
        "avisos": avisos
    }


# ===============================
# 🚀 VALIDAÇÃO COMPLETA
# ===============================

def validar_dados(df: pd.DataFrame) -> pd.DataFrame:
    """
    Aplica validação em todo o DataFrame.
    Retorna um DataFrame com resultados de validação.
    """
    resultados = []

    for idx, linha in df.iterrows():
        resultado = validar_linha(linha)

        resultados.append({
            "indice": idx,
            "status": resultado["status"],
            "erros": "; ".join(resultado["erros"]),
            "avisos": "; ".join(resultado["avisos"])
        })

    df_resultado = pd.DataFrame(resultados)

    return df_resultado