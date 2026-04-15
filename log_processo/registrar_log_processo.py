import pandas as pd
from datetime import datetime

from configuracoes.parametros_processo import (
    GERAR_LOG_ARQUIVO,
    NOME_ARQUIVO_LOG
)


# ===============================
# 🧾 CRIAR ESTRUTURA DE LOG
# ===============================

def criar_estrutura_log():
    """
    Cria um DataFrame vazio com estrutura padrão de log.
    """
    colunas = [
        "data_hora",
        "tipo",
        "obrigacao",
        "uf",
        "acao",
        "status",
        "mensagem"
    ]

    return pd.DataFrame(columns=colunas)


# ===============================
# ➕ ADICIONAR REGISTRO AO LOG
# ===============================

def adicionar_log(df_log, tipo, obrigacao="", uf="", acao="", status="", mensagem=""):
    """
    Adiciona um novo registro ao log.
    """
    novo_registro = {
        "data_hora": datetime.now(),
        "tipo": tipo,
        "obrigacao": obrigacao,
        "uf": uf,
        "acao": acao,
        "status": status,
        "mensagem": mensagem
    }

    return pd.concat([df_log, pd.DataFrame([novo_registro])], ignore_index=True)


# ===============================
# 💾 SALVAR LOG EM ARQUIVO
# ===============================

def salvar_log(df_log):
    """
    Salva o log em arquivo CSV.
    """
    if not GERAR_LOG_ARQUIVO:
        return

    df_log.to_csv(NOME_ARQUIVO_LOG, index=False, sep=";", encoding="utf-8-sig")


# ===============================
# 🚀 LOG DE INSERÇÃO
# ===============================

def log_insercoes(df_log, resultado_insercao):
    """
    Registra inserções realizadas.
    """
    for erro in resultado_insercao.get("erros", []):
        df_log = adicionar_log(
            df_log,
            tipo="INSERCAO",
            obrigacao=erro.get("obrigacao"),
            acao="INSERIR",
            status="ERRO",
            mensagem=erro.get("erro")
        )

    return df_log


# ===============================
# 📊 LOG RESUMO FINAL
# ===============================

def log_resumo(df_log, total_processado, total_novos, total_existentes):
    """
    Registra resumo da execução.
    """
    mensagem = f"Processados: {total_processado} | Novos: {total_novos} | Existentes: {total_existentes}"

    df_log = adicionar_log(
        df_log,
        tipo="RESUMO",
        acao="FINAL",
        status="OK",
        mensagem=mensagem
    )

    return df_log