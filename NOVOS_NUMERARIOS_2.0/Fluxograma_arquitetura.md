# 🏗️ Fluxograma da Arquitetura

```text
                                    ┌──────────────────────┐
                                    │      main.py         │
                                    │  (Ponto de Entrada)  │
                                    └──────────┬───────────┘
                                               │
                                               ▼
                              ┌────────────────────────────────┐
                              │ Validar arquivos abertos        │
                              │                                │
                              │ • RPE.csv                      │
                              │ • Pend_Especial_2026.xlsx      │
                              └────────────────┬───────────────┘
                                               │
                                               ▼
                                ┌───────────────────────────┐
                                │      backup.py            │
                                │ Criar backup automático   │
                                └───────────┬───────────────┘
                                            │
                                            ▼
                               ┌────────────────────────────┐
                               │      excel_online.py       │
                               │ Conectar ao Excel aberto   │
                               └───────────┬────────────────┘
                                           │
          ┌────────────────────────────────┼────────────────────────────────┐
          │                                │                                │
          ▼                                ▼                                ▼
 ┌──────────────────┐           ┌──────────────────┐           ┌────────────────────┐
 │ obrigacoes.py    │           │ assistencia.py   │           │ normalizacao.py    │
 │                  │           │                  │           │                    │
 │ Carrega todas as │           │ Carrega o        │           │ Padroniza textos   │
 │ OBRIGAÇÕES       │           │ DE_PARA          │           │                    │
 │ existentes       │           │ ASSISTÊNCIA      │           │                    │
 └────────┬─────────┘           └────────┬─────────┘           └─────────┬──────────┘
          │                              │                               │
          └───────────────┬──────────────┴───────────────┬───────────────┘
                          │                              │
                          ▼                              ▼
                ┌──────────────────────────────────────────────────┐
                │            Leitura do RPE.csv                    │
                │                                                  │
                │              pandas.DataFrame                    │
                └─────────────────────┬────────────────────────────┘
                                      │
                                      ▼
                     ┌──────────────────────────────────┐
                     │ Normalizar todas as colunas       │
                     │ (Maiúsculas, acentos, espaços)    │
                     └─────────────────┬────────────────┘
                                       │
                                       ▼
                    ┌─────────────────────────────────────┐
                    │ Existe no Dictionary de Obrigações? │
                    └───────────────┬─────────────────────┘
                                    │
                    ┌───────────────┴───────────────┐
                    │                               │
                  SIM                              NÃO
                    │                               │
                    ▼                               ▼
      ┌────────────────────────┐      ┌────────────────────────────┐
      │ Registrar no LOG       │      │ Agrupar por UF             │
      │ Status = IGNORADA      │      │                            │
      └────────────┬───────────┘      └──────────────┬─────────────┘
                   │                                  │
                   │                                  ▼
                   │                ┌─────────────────────────────────┐
                   │                │ Buscar ASSISTÊNCIA              │
                   │                │ via DE_PARA_ASSISTENCIA         │
                   │                └──────────────┬──────────────────┘
                   │                               │
                   │                               ▼
                   │                ┌─────────────────────────────────┐
                   │                │ Preencher automaticamente:      │
                   │                │                                 │
                   │                │ • DATA INÍCIO                  │
                   │                │ • REGIONAL                     │
                   │                │ • ASSIST. FILIAL               │
                   │                │ • ASSIST. MATRIZ               │
                   │                └──────────────┬──────────────────┘
                   │                               │
                   └───────────────┬───────────────┘
                                   │
                                   ▼
                    ┌────────────────────────────────────┐
                    │ insercao.py                        │
                    │                                    │
                    │ Inserir na aba correspondente      │
                    │ (CE, PB, RN, PE, AL, ...)          │
                    └────────────────┬───────────────────┘
                                     │
                                     ▼
                       ┌────────────────────────────┐
                       │ logger.py                  │
                       │                            │
                       │ Registrar resultado        │
                       │                            │
                       │ • Incluída                 │
                       │ • Ignorada                 │
                       │ • Erro                     │
                       └─────────────┬──────────────┘
                                     │
                                     ▼
                        ┌──────────────────────────┐
                        │ Processo Finalizado      │
                        │                          │
                        │ Backup criado            │
                        │ LOG atualizado           │
                        │ Excel atualizado         │
                        └──────────────────────────┘
```

RPE.csv
      │
      ▼
Filtrar novas obrigações
      │
      ▼
Agrupar por UF
      │
      ▼
Montar todas as linhas em memória
      │
      ▼
Gravar CE de uma vez
      │
      ▼
Gravar PB de uma vez
      │
      ▼
Gravar RN de uma vez
      │
      ▼
...