# 🚀 Index_Novos_Numerarios

## 📌 Objetivo

O projeto **Index_Novos_Numerarios** tem como objetivo automatizar a inclusão de **novas obrigações (numerários)** provenientes do sistema (`RPE.csv`) para a base operacional **Pend_Especial_2026.xlsx**, distribuindo automaticamente cada demanda para sua respectiva aba (UF) e preservando toda a estrutura existente da planilha.

O foco principal do projeto é:

- Evitar inclusão de obrigações duplicadas;
- Automatizar o preenchimento de colunas auxiliares;
- Reduzir trabalho manual;
- Garantir rastreabilidade através de logs;
- Criar backup automático antes de qualquer alteração;
- Operar sobre a própria planilha aberta no Excel, permitindo convivência com usuários trabalhando simultaneamente.

---

# 🏗 Arquitetura do Projeto

```
Index_Novos_Numerarios/

│

├── main.py

├── config.py

├── utils.py

├── backup.py

├── logger.py

├── normalizacao.py

├── assistencia.py

├── obrigacoes.py

├── insercao.py

├── excel_online.py

└── models.py
```

---

# 📂 Estrutura dos arquivos

## main.py

Responsável por orquestrar toda a execução.

Fluxo:

```
Inicialização

↓

Validar Excel

↓

Criar Backup

↓

Carregar RPE

↓

Carregar Obrigações

↓

Carregar Assistências

↓

Processar Novos Numerários

↓

Salvar LOG

↓

Finalizar
```

---

## config.py

Centraliza todas as configurações do projeto.

Exemplo:

```python
NOME_RPE = "RPE.csv"

NOME_PLANILHA = "Pend_Especial_2026.xlsx"

ABA_RPE = "RPE"

ABA_DEPARA = "DE_PARA_ASSISTENCIA"

ABA_LOG = "LOG_TRAMITE"
```

---

## utils.py

Funções auxiliares reutilizadas em todo projeto.

Exemplos:

- localizar workbook aberto;
- localizar worksheet;
- obter última linha;
- obter última coluna;
- verificar existência de aba;
- funções genéricas.

---

## backup.py

Responsável por criar automaticamente uma cópia da planilha online antes do processamento.

Formato:

```
Pend_Especial_2026_BACKUP_2026-06-15_1430.xlsx
```

---

## logger.py

Responsável por registrar toda movimentação realizada.

Será criada (ou atualizada) a aba:

```
LOG_TRAMITE
```

Contendo:

| Data/Hora | Obrigação | UF | Status | Observação |
|------------|----------|----|------------|----------------|
|15/06/2026|123456|CE|Incluída| |
|15/06/2026|654321|PB|Ignorada|Já existente|
|15/06/2026|777888|RN|Erro|UF inválida|

---

## normalizacao.py

Responsável pela padronização dos textos.

Regras:

- Upper Case
- Trim
- Remoção de acentos
- Remoção de caracteres especiais
- Remoção de múltiplos espaços

Exemplo:

```
João Pessoa

↓

JOAO PESSOA
```

---

## assistencia.py

Responsável pelo carregamento do DE/PARA.

Origem:

```
DE_PARA_ASSISTENCIA
```

Colunas:

```
UF PRESTADOR

CIDADE PRESTADOR

ASSIST. FILIA

ASSIST. MATRIZ
```

Busca realizada utilizando:

```
UF + CIDADE
```

Caso não encontre:

```
herdar valores da linha anterior.
```

---

## obrigacoes.py

Responsável por carregar todas as obrigações existentes.

Será criado um Dictionary contendo:

```
OBRIGAÇÃO

↓

True
```

Assim qualquer consulta será praticamente instantânea.

---

## insercao.py

Responsável pela inclusão efetiva dos novos numerários.

Este módulo:

- identifica aba destino;
- localiza última linha;
- preserva formatação;
- preenche colunas importadas;
- preenche colunas automáticas;
- consulta assistência;
- grava informações.

---

## excel_online.py

Responsável pela comunicação com o Excel aberto.

Não trabalha diretamente sobre arquivo salvo.

Opera utilizando a instância já aberta do Excel.

Isso permite:

- trabalhar com usuários simultaneamente;
- evitar perda de alterações;
- utilizar a memória atual do Excel.

---

## models.py

Centraliza modelos utilizados pelo projeto.

Exemplo:

```python
@dataclass
class LogProcessamento:

    obrigacao: str

    uf: str

    status: str

    detalhe: str
```

---

# 📥 Origem dos dados

Arquivo:

```
RPE.csv
```

Aba:

```
RPE
```

---

# 📤 Destino

Arquivo:

```
Pend_Especial_2026.xlsx
```

As abas possuem exatamente o nome da UF.

Exemplo:

```
CE

PB

RN

PE

AL

...
```

---

# 🔑 Chave de comparação

Toda verificação será realizada utilizando:

```
OBRIGAÇÃO
```

Caso já exista em qualquer aba:

```
não inserir.
```

---

# 📊 Colunas importadas

Serão buscadas pelo nome.

A posição das colunas não interfere.

## Relação

```
USUÁRIO

COD. PROCEDIMENTO

PROCEDIMENTO

NOME DO PRESTADOR

UF PRESTADOR

CIDADE PRESTADOR

VALOR DEPÓSITO

DATA DEPÓSITO

OBRIGAÇÃO

ÁREA DA PENDÊNCIA

STATUS
```

---

# ⚙ Colunas preenchidas automaticamente

## DATA INÍCIO

Recebe:

```
Data atual
```

---

## REGIONAL

Recebe:

```
Valor da linha anterior
```

---

## ASSIST. FILIA

Consulta:

```
DE_PARA_ASSISTENCIA
```

Caso não encontre:

```
herda linha anterior.
```

---

## ASSIST. MATRIZ

Consulta:

```
DE_PARA_ASSISTENCIA
```

Caso não encontre:

```
herda linha anterior.
```

---

# 🚫 Critérios para NÃO inserir

- Obrigação já existente;
- UF inexistente;
- Aba da UF inexistente;
- erro estrutural identificado durante processamento.

Todos serão registrados no LOG.

---

# 📝 Processo completo

```
Início

↓

Criar Backup

↓

Carregar RPE

↓

Normalizar dados

↓

Carregar Obrigações

↓

Carregar DE_PARA_ASSISTENCIA

↓

Filtrar apenas novas obrigações

↓

Agrupar por UF

↓

Inserir nas respectivas abas

↓

Preencher Data Início

↓

Preencher Regional

↓

Preencher Assistências

↓

Gerar LOG

↓

Finalizar
```

---

# ⚡ Estratégia de Performance

O projeto foi desenhado para minimizar acessos ao Excel.

Utiliza:

- pandas
- xlwings
- Dictionary
- processamento em memória
- agrupamento por UF
- escrita otimizada

O objetivo é que apenas o resultado final seja refletido nas planilhas abertas.

---

# 🔒 Segurança

Antes de qualquer alteração será criado automaticamente um backup.

Nenhuma informação original será perdida.

---

# 🎯 Objetivo final

Transformar um processo manual e sujeito a erros em uma rotina automatizada, rápida, rastreável e segura, garantindo que todas as novas obrigações sejam distribuídas corretamente entre as planilhas operacionais da empresa.