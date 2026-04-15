# Orquestrador do processo principal -- (o main)

from configuracoes.parametros_processo import COLUNA_CHAVE

# Leituras
from leituras_excel.ler_planilha_rpe import carregar_dados_rpe
from leituras_excel.ler_base_geral_pe import carregar_base_geral_pe

# Processamento
from padronizacao_dados.normalizar_textos import normalizar_dados
from validacoes_negocio.validar_dados_rpe import validar_dados
from comparadores_base.comparar_obrigacoes_existentes import comparar_obrigacoes

# Inserção
from insercao_base_pe.inserir_novas_obrigacoes import inserir_novas_obrigacoes

# Log
from logs_processo.registrar_log_processo import (
    criar_estrutura_log,
    adicionar_log,
    salvar_log,
    log_insercoes,
    log_resumo
)


def executar_processo():
    """
    Executa o fluxo completo do processo de novas obrigações.
    """

    log = criar_estrutura_log()

    try:
        # ===============================
        # 🚀 INÍCIO
        # ===============================
        log = adicionar_log(log, tipo="PROCESSO", mensagem="Início do processo")

        # ===============================
        # 📥 LEITURA RPE
        # ===============================
        df_rpe = carregar_dados_rpe()

        # ===============================
        # 📥 LEITURA BASE
        # ===============================
        df_base = carregar_base_geral_pe()

        # ===============================
        # 🧹 NORMALIZAÇÃO
        # ===============================
        df_rpe = normalizar_dados(df_rpe, [COLUNA_CHAVE, "UF PRESTADOR"])

        # ===============================
        # ✔ VALIDAÇÃO
        # ===============================
        resultado_validacao = validar_dados(df_rpe)

        # Filtrar apenas válidos
        indices_validos = resultado_validacao[
            resultado_validacao["status"] == "VALIDO"
        ]["indice"]

        df_validos = df_rpe.loc[indices_validos].reset_index(drop=True)

        # ===============================
        # 🧠 COMPARAÇÃO
        # ===============================
        resultado_comparacao = comparar_obrigacoes(df_validos, df_base)

        df_novos = resultado_comparacao["novos"]

        # ===============================
        # 📌 INSERÇÃO
        # ===============================
        resultado_insercao = inserir_novas_obrigacoes(df_novos)

        # ===============================
        # 🧾 LOG DE INSERÇÃO
        # ===============================
        log = log_insercoes(log, resultado_insercao)

        # ===============================
        # 📊 LOG RESUMO
        # ===============================
        log = log_resumo(
            log,
            total_processado=resultado_comparacao["total_rpe"],
            total_novos=resultado_comparacao["total_novos"],
            total_existentes=resultado_comparacao["total_existentes"]
        )

        log = adicionar_log(log, tipo="PROCESSO", status="OK", mensagem="Processo finalizado com sucesso")

    except Exception as e:
        log = adicionar_log(log, tipo="PROCESSO", status="ERRO", mensagem=str(e))

    finally:
        salvar_log(log)

    print("Processo finalizado.")


# ===============================
# ▶ EXECUÇÃO
# ===============================

if __name__ == "__main__":
    executar_processo()