📘 Módulo de Configuração — parametros_processo.py
📌 Visão Geral

O arquivo parametros_processo.py é o módulo central de configuração do sistema processo_novas_obrigacoes_rpe.

Ele concentra todas as definições estruturais do processo, eliminando valores fixos (hardcoded) espalhados pelo código.

Sua função é garantir:

Centralização de regras estruturais

Facilidade de manutenção

Organização arquitetural

Escalabilidade futura

Este módulo não contém lógica de negócio.

🏗 Papel na Arquitetura do Projeto

Dentro do fluxo do sistema, este arquivo é responsável por definir:

Caminhos dos arquivos utilizados

Nome oficial das abas

Abas que devem ser ignoradas

Colunas oficiais do processo

Coluna chave de identificação

Regras de validação

Configurações de log

Flags de execução (debug e controle de reprocessamento)

Ele atua como ponto único de verdade estrutural do sistema.

📂 Estrutura do Arquivo

O arquivo está dividido em blocos organizados por responsabilidade.

1️⃣ Caminhos dos Arquivos

Define os arquivos principais utilizados pelo sistema:

Planilha RPE (entrada)

Base Geral PE (validação e inserção)

Esses caminhos podem ser alterados sem necessidade de modificar a lógica dos módulos.

2️⃣ Nomes das Abas

Define:

Nome oficial da aba da RPE

Lista de abas que devem ser ignoradas na Base Geral PE

Exemplo de abas ignoradas:

LOG

RESUMO

CONFIG

Isso evita que o sistema processe abas indevidas.

3️⃣ Colunas Oficiais do Processo

Contém:

Lista de colunas permitidas

Coluna chave (OBRIGAÇÃO)

Essa padronização garante:

Integridade estrutural

Consistência entre planilhas

Segurança na comparação e inserção de dados

4️⃣ Configurações de Validação

Define regras como:

Quais colunas são obrigatórias

Se valores financeiros devem ser validados

Se datas devem ser verificadas

Essas regras podem ser ativadas ou desativadas via flag, sem alterar código operacional.

5️⃣ Configuração de Log

Controla:

Se o sistema deve gerar log

Nome do arquivo de log

Permite rastreabilidade e auditoria do processo.

6️⃣ Configurações de Execução

Define comportamentos do sistema, como:

MODO_DEBUG

Permissão de reexecução sem erro

Essas flags permitem diferenciar ambientes de desenvolvimento e produção.

🔄 Como o Módulo é Utilizado

Os demais módulos importam as configurações da seguinte forma:

from configuracoes import parametros_processo as config

print(config.CAMINHO_RPE)
print(config.COLUNA_CHAVE)


Essa abordagem:

Mantém o código limpo

Evita repetição

Padroniza acesso às configurações

🧠 Boas Práticas Aplicadas

Este módulo segue princípios de arquitetura limpa:

Separação entre configuração e lógica

Centralização de parâmetros

Redução de acoplamento

Facilidade de manutenção

Escalabilidade

🚀 Benefícios Estratégicos

Com essa estrutura:

Mudança de coluna → altera apenas aqui

Mudança de aba → altera apenas aqui

Mudança de regra → altera apenas aqui

Mudança de caminho → altera apenas aqui

Sem impacto estrutural no restante do sistema.

⚠ Diretrizes Importantes

Este arquivo:

❌ Não deve conter funções operacionais

❌ Não deve conter lógica de negócio

❌ Não deve conter manipulação de dados

Ele existe exclusivamente para configuração do sistema.

Qualquer regra de processamento deve estar nos módulos específicos, como:

carregadores

validadores

comparadores

inseridores

geradores de log

📍 Conclusão

parametros_processo.py é o núcleo estrutural do sistema.

Ele garante organização, previsibilidade e evolução sustentável do projeto, transformando o código de um simples script para um sistema estruturado e profissional.