# processo_novas_obrigacoes_rpe
Projeto da Hapvida da área do Relacionamento - Pendência Especial.

iremos criar algumas automações em python para os processos da pendência espeical tratados em planilhas onlines.

# 📘 Documentação Técnica — Macro `AtualizarBasePE`

## 1. Visão Geral

A macro `AtualizarBasePE` tem como objetivo sincronizar os campos **Área da Pendência** e **Status** das abas do arquivo **Base Geral PE.xlsx**(dependendo do ano a Base pode ter o nome diferente, fazendo necessário a atualização no script para funcionar) com base em uma **base unificada** (`RPE.csv`).

A atualização é **controlada**, **rastreável por log** e **não destrutiva**, garantindo que apenas valores diferentes e não vazios sejam sobrescritos.

---

## 2. Escopo Funcional

### 2.1 O que a macro FAZ

* Lê registros da base `RPE.csv`.
* Localiza a **Obrigação** correspondente nas abas válidas do arquivo destino.
* Atualiza somente:

  * `ÁREA DA PENDÊNCIA`
  * `STATUS`
* Registra todas as ações em log.

### 2.2 O que a macro NÃO FAZ

* Não cria novas obrigações.
* Não remove registros.
* Não atualiza outras colunas.
* Não valida consistência de UF entre base e aba.
* Não trata duplicidades de obrigação.

---

## 3. Pré-requisitos de Execução

### 3.1 Arquivos obrigatórios abertos

| Tipo    | Nome                       |
| ------- | -------------------------- |
| Destino | `Base Geral PE.xlsx` |
| Base    | `RPE.csv`                  |

### 3.2 Abas obrigatórias

* Base unificada: `RPE`
* Log (criada automaticamente): `LOG_ATUALIZACAO`

---

## 4. Estrutura Esperada da Base (RPE.csv)

### 4.1 Cabeçalho

* Localizado na **linha 1**.

### 4.2 Colunas obrigatórias (nomes exatos)

| Coluna            | Uso                         |
| ----------------- | --------------------------- |
| Obrigação         | Chave principal de busca    |
| UF Prestador      | Registro informativo no log |
| Área da Pendência | Novo valor para atualização |
| Status            | Novo valor para atualização |

### 4.3 Regra de leitura

* Iteração inicia na linha 2.
* Linhas sem obrigação são ignoradas.

---

## 5. Estrutura Esperada do Arquivo Destino

### 5.1 Abas ignoradas

As seguintes abas **não participam da busca**:

* LOG_ATUALIZACAO
* VISÃO GERAL - ESTÁGIO
* DADOS DE VALIDAÇÃO
* INDISPONIVEL TEMPORARIAMENTE
* MG

Todas as demais abas são consideradas válidas.

### 5.2 Cabeçalho das abas válidas

* Localizado na **linha 4**.

### 5.3 Colunas obrigatórias (nomes exatos)

| Coluna            | Uso                      |
| ----------------- | ------------------------ |
| OBRIGAÇÃO         | Chave de correspondência |
| ÁREA DA PENDÊNCIA | Campo atualizável        |
| STATUS            | Campo atualizável        |

### 5.4 Dados

* Iniciam na **linha 5**.

---

## 6. Regra de Localização da Obrigação

1. A macro percorre as abas válidas sequencialmente.
2. Dentro de cada aba, percorre linha a linha.
3. A comparação é textual, com `Trim`.
4. A **primeira ocorrência encontrada encerra a busca**.

Caso não seja encontrada em nenhuma aba:

* Nenhuma atualização é realizada.
* Um log do tipo **AVISO** é registrado.

---

## 7. Regras de Atualização de Dados

### 7.1 Condições obrigatórias

Um campo **só é atualizado** se:

* O novo valor **não for vazio**.
* O novo valor **for diferente do valor atual**.

### 7.2 Campos atualizados

| Campo             | Origem   | Destino     |
| ----------------- | -------- | ----------- |
| Área da Pendência | Base RPE | Aba destino |
| Status            | Base RPE | Aba destino |

Campos vazios na base **nunca sobrescrevem valores existentes**.

---

## 8. Regras de Proteção de Abas

* A macro detecta se a aba destino está protegida.
* Se estiver:

  * Remove a proteção.
  * Executa a atualização.
  * Reaplica a proteção.

### 8.1 Permissões após reproteção

* Filtros: Permitidos
* Tabelas Dinâmicas: Permitidas
* Demais alterações estruturais: Bloqueadas

---

## 9. Log de Execução

### 9.1 Estrutura do Log

Cada linha da base gera um registro no log.

| Campo                 | Descrição                |
| --------------------- | ------------------------ |
| DataHora              | Momento da ação          |
| Aba                   | Aba afetada ou N/A       |
| Obrigação             | Código tratado           |
| UF                    | UF da base               |
| ColAtualizadas        | Campos alterados         |
| OldArea / NewArea     | Antes e depois           |
| OldStatus / NewStatus | Antes e depois           |
| Tipo                  | OK / SEM_MUDANCA / AVISO |
| Mensagem              | Detalhe da ação          |
| LinhaBase             | Linha de origem          |

---

## 10. Finalização

Ao concluir:

* StatusBar é restaurada.
* Mensagem de sucesso é exibida.

---

## 11. Considerações Técnicas

* A macro prioriza **segurança e rastreabilidade**.
* Performance é proporcional ao volume de abas × linhas.
* Alterações de nomenclatura quebram a execução.

---

📌 **Este documento reflete fielmente o comportamento atual da macro, sem otimizações ou melhorias aplicadas.**

---

# 🧱 Arquitetura do Novo Processo em Python — Inserção de Novas Obrigações (RPE → Base Geral PE)

## 12. Visão Geral da Arquitetura

Este novo processo será implementado em **Python**, de forma **independente** da macro VBA, respeitando todas as regras documentadas anteriormente.

O objetivo é **identificar, validar e inserir novas obrigações** presentes na planilha **RPE**, mas **ausentes na Base Geral PE**, mantendo padronização, performance e rastreabilidade.

---

## 13. Estrutura de Pastas do Projeto

```
processo_novas_obrigacoes_rpe/
│
├── executar_processo.py          # Orquestrador do processo
│
├── configuracoes/
│   └── parametros_processo.py    # Caminhos, nomes de abas, colunas e regras
│
├── leituras_excel/
│   ├── ler_planilha_rpe.py       # Leitura da planilha RPE
│   └── ler_base_geral_pe.py      # Leitura da Base Geral PE
│
├── padronizacao_dados/
│   └── normalizar_textos.py      # Maiúsculo, trim, limpeza
│
├── validacoes_negocio/
│   └── validar_dados_rpe.py      # Validação de campos obrigatórios
│
├── comparadores_base/
│   └── comparar_obrigacoes_existentes.py   # Evita duplicidade
│
├── insercao_base_pe/
│   └── inserir_novas_obrigacoes.py         # Inserção nas abas corretas
│
├── logs_processo/
│   └── registrar_log_processo.py           # Log estruturado
│
└── utilitarios_excel/
    └── funcoes_excel.py                    # Busca de colunas, abas, helpers


```

---

## 14. Responsabilidade de Cada Camada

### 14.1 excecutar_processo.py (Orquestração)

* Controla o fluxo completo do processo
* Chama os módulos na ordem correta
* Finaliza com resumo de execução

---

### 14.2 leitura_excel

**Responsabilidade:** leitura segura dos arquivos

* `ler_planilha_rpe.py`

  * Abre a planilha RPE
  * Localiza colunas por nome (não por posição)
  * Retorna DataFrame padronizado

* `ler_planilha_base_geral_pe.py`

  * Abre a Base Geral PE
  * Identifica abas válidas
  * Constrói estrutura de obrigações já existentes

---

### 14.3 padronizacao_dados
* `normalizar_texto.py`

**Responsabilidade:** padronização de dados

* Converte todos os textos para MAIÚSCULO
* Remove espaços extras
* Padroniza nomes de colunas

---

### 14.4 validacao_negocio
* `validar_dados_rpe.py`

**Responsabilidade:** garantir integridade dos dados

Valida, no mínimo:

* Obrigação (obrigatória)
* UF Prestador
* Área da Pendência
* Status
* Datas válidas
* Valores numéricos válidos

Define se a linha é:

* VÁLIDA
* AVISO
* ERRO (bloqueia inserção)

---

### 14.5 comparadores_base
* `comparar_obg_existente.py`

**Responsabilidade:** evitar duplicidades

* Constrói um `set` ou `dict` com todas as obrigações existentes
* Verifica se a obrigação da RPE já existe na Base Geral PE
* Marca como **NOVA** ou **DUPLICADA**

---

### 14.6 insercao_base_pe
* `inserir_novas_obgs.py`

**Responsabilidade:** inserir dados corretamente

* Determina a aba destino (mesmas regras da Base Geral PE)
* Insere os dados na próxima linha disponível
* Garante alinhamento correto das colunas

---

### 14.7 log_processo
* `registrar_log_processo.py`

**Responsabilidade:** rastreabilidade total

Cada linha processada gera log com:

* Data/Hora
* Obrigação
* Resultado (INSERIDO / DUPLICADO / ERRO)
* Aba destino
* Mensagem detalhada

---

## 15. Estratégia de Performance

* Uso de `pandas` para leitura e validação
* Uso de `set` / `dict` para busca O(1)
* Escrita em Excel em lote (evitar célula a célula)
* Processamento linear (O(n))

---

## 16. Integração com o Processo Existente

* Este processo **apenas insere novas obrigações**
* O processo `Index_ABAS_BG` permanece responsável por atualizações
* Os dois processos são independentes e complementares

---

📌 **Esta arquitetura garante escalabilidade, governança e facilidade de manutenção.**

---

# 🆕 12. Novo Processo — Inserção de Novas Obrigações (RPE → Base Geral PE)

## 12.1 Objetivo

Criar um processo **independente e seguro**, em **Python**, responsável por identificar **novas obrigações** presentes na planilha **RPE** que **ainda não existem** na **Base Geral PE**, inserindo-as corretamente nas abas de destino, respeitando **todas as regras já estabelecidas** neste documento.

Este processo **não substitui** a macro de atualização existente; ele **alimenta** a Base Geral PE com novas demandas.

---

## 12.2 Princípios do Processo

* Busca de colunas **por nome**, nunca por posição
* Normalização de texto (maiúsculo, trim)
* Uso de **estruturas de dados eficientes** (set e dict)
* Operação **idempotente** (rodar mais de uma vez não duplica dados)
* Log detalhado de todas as decisões

---

## 12.3 Colunas Permitidas (Origem → Destino)

Somente as colunas abaixo podem ser inseridas:

* USUÁRIO
* PROCEDIMENTO
* NOME PRESTADOR
* UF PRESTADOR
* CIDADE PRESTADOR
* VALOR DEPÓSITO
* DATA DEPÓSITO
* OBRIGAÇÃO
* ÁREA DA PENDÊNCIA
* STATUS

Qualquer coluna fora dessa lista é ignorada.

---

## 12.4 Fluxo Técnico do Processo (Python)

### Etapa 1 — Carregamento de Dados

1. Abrir a Base Geral PE.
2. Ler todas as abas válidas.
3. Abrir a planilha RPE.

---

### Etapa 2 — Normalização de Cabeçalhos

1. Converter todos os nomes de colunas para **MAIÚSCULO**.
2. Aplicar `strip()` nos nomes.
3. Criar um dicionário `colunas_rpe` mapeando nome → índice.

---

### Etapa 3 — Construção do Índice de Obrigações Existentes

1. Percorrer todas as abas válidas da Base Geral PE.
2. Ler a coluna **OBRIGAÇÃO**.
3. Normalizar os valores (maiúsculo + trim).
4. Armazenar todas as obrigações em um **set** (`obrigacoes_existentes`).

---

### Etapa 4 — Leitura da RPE (Novas Demandas)

Para cada linha da RPE:

1. Ignorar linhas sem OBRIGAÇÃO.
2. Normalizar a OBRIGAÇÃO.
3. Verificar se a obrigação já existe no set.

* Se existir → registrar log `JA_EXISTE` e pular.
* Se não existir → seguir para validação.

---

### Etapa 5 — Validação da Linha

Antes de inserir:

* OBRIGAÇÃO não vazia
* UF PRESTADOR não vazia
* DATA DEPÓSITO válida

Falhas:

* Erro crítico → não insere, registra `ERRO`
* Aviso → insere, registra `AVISO`

---

### Etapa 6 — Determinação da Aba de Destino

1. Aplicar as **mesmas regras de distribuição** já existentes.
2. Determinar a aba correta.
3. Se a aba não existir → `ERRO_DISTRIBUICAO`.

---

### Etapa 7 — Preparação do Registro

1. Criar um dicionário com as colunas permitidas.
2. Converter todos os textos para **MAIÚSCULO**.
3. Garantir tipos corretos (datas, valores).

---

### Etapa 8 — Inserção na Base Geral PE

1. Localizar a próxima linha vazia da aba destino.
2. Inserir os valores respeitando o cabeçalho.
3. Não sobrescrever fórmulas existentes.

---

### Etapa 9 — Atualização do Índice

* Adicionar a nova obrigação ao set `obrigacoes_existentes`.

---

### Etapa 10 — Log de Inserção

Registrar:

* DataHora
* Obrigação
* Aba destino
* Resultado (INSERIDO / JA_EXISTE / ERRO)
* Mensagem detalhada
* Linha RPE

---

### Etapa 11 — Finalização

* Salvar o arquivo Base Geral PE.
* Exportar ou atualizar log.
* Exibir resumo da execução.

---

## 12.5 Garantias do Processo

* Nenhuma obrigação duplicada é criada.
* Nenhuma regra existente é violada.
* O processo é seguro para reexecução.

📌 **Este processo foi desenhado para implementação direta em Python, usando pandas + openpyxl, com foco em segurança, performance e rastreabilidade.**

---

# 🚀 Propostas de Otimização — Performance, Arquitetura e Robustez

Esta seção descreve **otimizações possíveis**, classificadas por **nível de impacto** e **risco de mudança de comportamento**. Nenhuma proposta abaixo é aplicada automaticamente; todas são sugestões técnicas.

---

## 1. Otimizações de Performance (Alto Impacto, Baixo Risco)

### 1.1 Uso de Arrays para leitura da base

**Situação atual:**

* Leitura célula a célula da base `RPE.csv`.

**Proposta:**

* Carregar toda a base em um array VBA (`Variant`).

**Benefícios:**

* Redução drástica de acessos ao Excel.
* Ganho significativo de velocidade.

**Impacto funcional:** Nenhum.

---

### 1.2 Uso de Arrays nas abas destino

**Situação atual:**

* Leitura e escrita célula a célula.

**Proposta:**

* Ler cada aba válida em array uma única vez.
* Aplicar alterações no array.
* Escrever o array de volta ao final.

**Benefícios:**

* Performance até dezenas de vezes maior.

**Impacto funcional:** Nenhum.

---

## 2. Otimizações com Dicionários (Altíssimo Impacto)

### 2.1 Dicionário de obrigações (lookup O(1))

**Situação atual:**

* Busca sequencial por obrigação (O(n)).

**Proposta:**

* Criar um `Scripting.Dictionary` com chave:

  * `Obrigação` ou `Obrigação|Aba`
* Valor:

  * Linha no array

**Benefícios:**

* Elimina loops aninhados.
* Escala muito melhor para bases grandes.

**Impacto funcional:**

* **Pode alterar comportamento** se houver obrigações duplicadas.
* Deve ser combinado com regra explícita de prioridade.

---

### 2.2 Validação explícita de duplicidade

**Proposta adicional:**

* Detectar duplicidades no carregamento.
* Registrar no log como erro ou aviso.

**Benefícios:**

* Evita atualizações silenciosamente incorretas.

---

## 3. Otimizações Estruturais (Médio Impacto)

### 3.1 Modularização do código

**Situação atual:**

* Macro monolítica.

**Proposta:**
Dividir em procedimentos:

* `CarregarBase()`
* `CarregarAbasDestino()`
* `AtualizarRegistro()`
* `RegistrarLog()`

**Benefícios:**

* Código mais legível.
* Facilita testes e manutenção.

**Impacto funcional:** Nenhum.

---

### 3.2 Substituir `GoTo`

**Proposta:**

* Usar `Continue For` (ou lógica condicional).

**Benefícios:**

* Fluxo mais claro.
* Menos risco de erro futuro.

---

## 4. Otimizações de Robustez (Baixo Impacto, Alto Valor)

### 4.1 Normalização de texto

**Proposta:**

* Padronizar textos (Trim, UCase, Replace múltiplos espaços).

**Benefícios:**

* Menos falhas por variações mínimas.

---

### 4.2 Validação de colunas obrigatórias

**Proposta:**

* Verificar se todas as colunas foram encontradas antes de executar.

**Benefícios:**

* Falha controlada com mensagem clara.

---

## 5. Otimizações Operacionais

### 5.1 Backup automático

**Proposta:**

* Criar cópia do arquivo destino com timestamp antes da execução.

**Benefícios:**

* Rollback imediato.

---

### 5.2 Relatório-resumo final

**Proposta:**

* Total de registros:

  * Atualizados
  * Sem mudança
  * Não encontrados

**Benefícios:**

* Visão executiva imediata.

---

## 6. Sugestão de Roadmap de Evolução

**Fase 1 (segura):**

* Arrays
* Modularização
* Backup automático

**Fase 2 (controlada):**

* Dicionários
* Validação de duplicidade

**Fase 3 (maturidade):**

* Casos de teste
* Métricas automáticas

---

📌 **Nenhuma otimização altera o comportamento atual sem decisão explícita.**

---

# 🧪 Casos de Teste Esperados — Macro `AtualizarBasePE`

Esta seção define **cenários de teste funcionais e técnicos** para validar o comportamento da macro antes e após otimizações. Cada caso descreve **entrada**, **processamento esperado** e **resultado esperado (dados + log)**.

---

## 1. Casos de Teste Básicos (Comportamento Normal)

### CT-01 — Atualização completa (Área + Status)

**Entrada (Base):**

* Obrigação existente
* Área diferente do destino
* Status diferente do destino

**Resultado esperado:**

* Área atualizada
* Status atualizado
* Log:

  * Tipo: `OK`
  * ColAtualizadas: `Área da Pendência;Status;`

---

### CT-02 — Atualização parcial (somente Área)

**Entrada:**

* Área diferente
* Status igual

**Resultado esperado:**

* Apenas Área atualizada
* Log:

  * Tipo: `OK`
  * ColAtualizadas: `Área da Pendência;`

---

### CT-03 — Atualização parcial (somente Status)

**Entrada:**

* Área igual
* Status diferente

**Resultado esperado:**

* Apenas Status atualizado
* Log:

  * Tipo: `OK`
  * ColAtualizadas: `Status;`

---

### CT-04 — Nenhuma alteração necessária

**Entrada:**

* Área igual
* Status igual

**Resultado esperado:**

* Nenhuma célula alterada
* Log:

  * Tipo: `SEM_MUDANCA`
  * Mensagem: `Valores iguais`

---

## 2. Casos de Teste com Dados Vazios

### CT-05 — Área vazia na base

**Entrada:**

* Área vazia
* Status diferente

**Resultado esperado:**

* Área NÃO sobrescrita
* Status atualizado

---

### CT-06 — Status vazio na base

**Entrada:**

* Área diferente
* Status vazio

**Resultado esperado:**

* Área atualizada
* Status NÃO sobrescrito

---

### CT-07 — Área e Status vazios

**Entrada:**

* Ambos vazios

**Resultado esperado:**

* Nenhuma alteração
* Log: `SEM_MUDANCA`

---

## 3. Casos de Teste de Localização da Obrigação

### CT-08 — Obrigação inexistente

**Entrada:**

* Obrigação não presente em nenhuma aba válida

**Resultado esperado:**

* Nenhuma atualização
* Log:

  * Tipo: `AVISO`
  * Mensagem: `Obrigação não encontrada em nenhuma aba válida`

---

### CT-09 — Obrigação duplicada em abas diferentes

**Entrada:**

* Mesma obrigação em duas abas

**Resultado esperado (macro atual):**

* Atualiza a primeira ocorrência encontrada
* Nenhum aviso de duplicidade

---

## 4. Casos de Teste Estruturais

### CT-10 — Alteração no nome da coluna

**Entrada:**

* Coluna `STATUS` renomeada

**Resultado esperado:**

* Erro de execução ou falha na atualização
* Macro não prossegue corretamente

---

### CT-11 — Cabeçalho fora da linha esperada

**Entrada:**

* Cabeçalho deslocado

**Resultado esperado:**

* Colunas não localizadas
* Atualização não ocorre

---

## 5. Casos de Teste de Proteção de Aba

### CT-12 — Aba protegida sem senha

**Entrada:**

* Aba protegida
* SENHA_ABA vazia

**Resultado esperado:**

* Aba é desprotegida
* Atualização ocorre
* Aba é reprotegida

---

### CT-13 — Aba protegida com senha incorreta

**Entrada:**

* Aba protegida com senha diferente

**Resultado esperado:**

* Falha silenciosa
* Possível não atualização
* Log não acusa erro

---

## 6. Casos de Teste de Volume e Performance

### CT-14 — Base pequena

**Entrada:**

* < 100 registros

**Resultado esperado:**

* Execução rápida
* StatusBar atualiza corretamente

---

### CT-15 — Base grande

**Entrada:**

* > 50.000 registros

**Resultado esperado:**

* Execução lenta na versão atual
* Risco de travamento

---

## 7. Uso dos Casos de Teste

Estes casos devem ser utilizados para:

* Validar comportamento atual
* Comparar versões após otimização
* Garantir que refatorações não alterem regras existentes

---

📌 **Estes testes representam o contrato funcional da macro.**

---

# 🆕 Novo Processo — Inclusão de Novas Obrigações na Base Geral PE

Este processo complementa a macro `AtualizarBasePE`. Seu objetivo é **identificar obrigações novas existentes na planilha RPE que ainda não constam na Base Geral PE** e **inseri-las corretamente**, respeitando **todas as regras já documentadas**.

---

## 1. Objetivo do Processo

* Comparar a planilha **RPE** com a **Base Geral PE**.
* Identificar obrigações **existentes na RPE** e **ausentes na Base Geral PE**.
* Inserir essas obrigações como **novas demandas** na Base Geral PE.
* Manter as mesmas regras de:

  * distribuição por abas
  * estrutura
  * log
  * segurança

📌 **Nenhuma regra existente é alterada. Este é um processo adicional.**

---

## 2. Princípios Herdados (Leis do Sistema)

Este novo processo herda integralmente:

* Estrutura de abas válidas e ignoradas
* Regras de proteção/desproteção
* Dependência de nomes exatos de colunas
* Uso de log
* Tratamento textual com `Trim`
* Não sobrescrever dados existentes

---

## 3. Definição de “Obrigação Nova”

Uma obrigação é considerada **nova** quando:

* Existe na planilha **RPE**
* **Não existe em nenhuma aba válida** da Base Geral PE

A verificação é **global**, não por aba.

---

## 4. Fluxo Lógico do Processo

### 4.1 Leitura da Base Geral PE

* Percorrer todas as abas válidas.
* Coletar todas as obrigações existentes.
* Armazenar em estrutura de busca (conceitualmente um conjunto).

---

### 4.2 Leitura da RPE

* Percorrer todas as linhas válidas.
* Ignorar linhas sem obrigação.

---

### 4.3 Comparação

Para cada obrigação da RPE:

* Se **já existir** na Base Geral PE → ignorar
* Se **não existir** → marcar como nova demanda

---

## 5. Regra de Distribuição nas Abas

A inserção de novas obrigações deve respeitar **as mesmas regras de distribuição já existentes**, incluindo:

* Abas ignoradas permanecem ignoradas
* Estrutura de cabeçalho na linha 4
* Dados iniciando na linha 5

📌 **A regra exata de escolha da aba (por UF ou outra lógica) deve ser a mesma já utilizada no processo atual.**

---

## 6. Regra de Inserção

Para cada nova obrigação:

* Inserir uma **nova linha** ao final da aba destino
* Preencher obrigatoriamente:

  * OBRIGAÇÃO
  * ÁREA DA PENDÊNCIA (se disponível)
  * STATUS (se disponível)

Campos vazios na RPE permanecem vazios na Base Geral PE.

---

## 7. Regras de Segurança

* Se a aba estiver protegida:

  * Desproteger
  * Inserir registros
  * Reproteger

* O processo **não altera registros existentes**.

---

## 8. Log do Novo Processo

Cada obrigação da RPE deve gerar log:

### 8.1 Tipos de Log

| Tipo     | Situação                  |
| -------- | ------------------------- |
| INSERIDO | Nova obrigação adicionada |
| IGNORADO | Obrigação já existente    |
| AVISO    | Falha na inserção         |

### 8.2 Informações mínimas no log

* DataHora
* Obrigação
* UF
* Aba destino
* Tipo
* Mensagem

---

## 9. Ordem Recomendada de Execução

1. **Processo de Inclusão de Novas Obrigações** (este documento)
2. **Processo de Atualização** (`AtualizarBasePE`)

📌 Garante que novas demandas já entrem atualizadas.

---

## 10. Separação de Responsabilidades

| Processo    | Responsabilidade           |
| ----------- | -------------------------- |
| Inclusão    | Criar novas obrigações     |
| Atualização | Atualizar dados existentes |

Misturar os dois fluxos é **proibido por regra de segurança**.

---

## 11. Conclusão Técnica

Este novo processo:

* Preserva todas as regras existentes
* Evita duplicidades
* Mantém rastreabilidade
* Prepara a Base Geral PE para crescimento controlado

📌 **Este processo deve ser implementado como macro separada.**

---

## 🔧 Aprimoramentos do Processo de Inserção de Novas Obrigações (RPE → Base Geral PE)

### 1️⃣ Padronização e Normalização de Dados

* Todas as comparações devem ocorrer **após normalização**:

  * Converter textos da RPE para **MAIÚSCULO** antes de qualquer comparação ou inserção.
  * Remover espaços extras no início/fim.
  * Padronizar acentuação (se aplicável ao processo principal).

📌 Justificativa:

* A RPE utiliza *capitalização padrão* (primeira letra maiúscula).
* A Base Geral PE utiliza **MAIÚSCULO TOTAL**.
* A normalização elimina falsos negativos na verificação de existência.

---

### 2️⃣ Busca de Colunas por Nome (Não por Posição)

* Nenhuma coluna deve ser acessada por índice fixo.
* O processo deve:

  * Ler o cabeçalho da planilha RPE.
  * Criar um **mapeamento dinâmico de colunas** (ex: dicionário `{nome_coluna: índice}`).
  * Utilizar esse mapeamento para leitura dos dados.

✔ Benefícios:

* Tolerância a mudanças de ordem das colunas.
* Redução de risco operacional.
* Maior vida útil do processo.

---

### 3️⃣ Colunas Autorizadas para Transferência (RPE → Base Geral PE)

Somente as colunas abaixo devem ser lidas da RPE e inseridas na Base Geral PE:

```
Usuário,
Procedimento,
Nome Prestador,
UF Prestador,
Cidade Prestador,
Valor Depósito,
Data Depósito,
Obrigação,
Área da Pendência,
Status
```

⚠️ Qualquer outra coluna existente na RPE deve ser **ignorada automaticamente**.

---

### 4️⃣ Regra de Identificação de Nova Obrigação (Reforçada)

Uma obrigação da RPE será considerada **nova** se:

* Após normalização (MAIÚSCULO + TRIM),
* O valor da coluna **OBRIGAÇÃO**
* **Não existir** na coluna **OBRIGAÇÃO** da Base Geral PE.

📌 A verificação deve usar:

* Estrutura de dados em memória (ex: `set` ou `dict`) para performance.

---

### 5️⃣ Validação Obrigatória Antes da Inserção

Antes de inserir qualquer nova obrigação:

* Verificar preenchimento obrigatório de:

  * Obrigação
  * UF Prestador
  * Nome Prestador
* Validar:

  * Data Depósito em formato válido
  * Valor Depósito numérico

❌ Falha em qualquer validação:

* Não inserir a linha
* Registrar erro detalhado em log

---

### 6️⃣ Inserção Controlada na Base Geral PE

* Inserir as novas obrigações apenas ao final da base unificada.
* Não sobrescrever dados existentes.
* Preservar:

  * Fórmulas
  * Formatação
  * Estrutura das abas

---

### 7️⃣ Integração com o Processo Principal

Após inseridas:

* As novas obrigações passam automaticamente a ser tratadas pelo processo padrão `Index_ABAS_BG`.
* Nenhuma regra nova de distribuição será criada.
* A lógica de UF permanece única e centralizada.

---

### 8️⃣ Logging Específico do Processo RPE

Criar registro dedicado contendo:

* Data/hora da execução
* Total de registros lidos da RPE
* Total de novas obrigações inseridas
* Total de obrigações já existentes
* Total de erros

📄 Log recomendado: aba `LOG_INSERCAO_RPE`

---

### 🔐 Princípios Garantidos com os Aprimoramentos

* Robustez contra mudanças na RPE
* Comparação confiável entre bases
* Zero impacto em obrigações existentes
* Alta rastreabilidade
* Performance escalável
