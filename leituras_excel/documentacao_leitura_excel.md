processo_novas_obrigacoes_rpe/leituras_excel/README.md
📘 Módulo leituras_excel
📌 Propósito

O módulo leituras_excel é responsável por realizar a leitura controlada das planilhas utilizadas no processo.

Ele faz a ponte entre:

Arquivos físicos (Excel)

Estrutura interna do sistema (DataFrames)

Seu papel é garantir que os dados entrem no sistema de forma organizada, validada estruturalmente e previsível.

🏗 Papel na Arquitetura

Este módulo:

Utiliza parâmetros definidos em configuracoes

Utiliza funções genéricas de utilitarios_excel

Retorna dados prontos para normalização e validação

Ele não aplica regra de negócio.

Ele apenas lê e organiza.

📂 Estrutura Interna
leituras_excel/
├── ler_planilha_rpe.py
└── ler_base_geral_pe.py
📄 1. ler_planilha_rpe.py

Responsável por:

Localizar o arquivo da RPE

Ler a aba correta

Garantir que a aba exista

Retornar DataFrame estruturado

Responsabilidades:

✔ Validar existência do arquivo
✔ Validar existência da aba definida
✔ Carregar dados com pandas
✔ Padronizar nomes de colunas (se necessário)
✔ Retornar estrutura pronta para validação

Não deve fazer:

❌ Validar regra de negócio
❌ Inserir dados
❌ Comparar com base
❌ Gerar log

📄 2. ler_base_geral_pe.py

Responsável por:

Localizar a Base Geral PE

Listar abas válidas

Ignorar abas administrativas

Retornar estrutura consolidada para comparação

Responsabilidades:

✔ Validar existência do arquivo
✔ Identificar abas válidas
✔ Ignorar abas definidas como administrativas
✔ Carregar dados de múltiplas abas
✔ Retornar estrutura consolidada

🔄 Fluxo de Funcionamento

O fluxo interno deste módulo segue a lógica:

Recebe caminho definido em configuracoes

Usa funções de utilitarios_excel

Valida estrutura básica

Retorna DataFrame(s) organizados

Ele não toma decisões.

Ele prepara o terreno para quem toma.

🧠 Princípios Aplicados

Este módulo segue:

Separação de responsabilidade

Desacoplamento da lógica de negócio

Reutilização de utilitários

Padronização de entrada de dados

🚫 O que NÃO deve existir aqui

❌ Regras de validação de obrigação

❌ Comparação de registros

❌ Inserção na base

❌ Tratamento de duplicidade

❌ Decisão de aba de inserção

Se houver decisão de negócio aqui, o módulo está incorreto.

🚀 Benefícios Arquiteturais

Com essa separação:

Erros de leitura ficam isolados

Mudança de estrutura de planilha é tratada aqui

O restante do sistema permanece estável

Testes unitários ficam mais simples

📌 Conclusão

O módulo leituras_excel é o ponto de entrada dos dados no sistema.

Ele garante que tudo que entra esteja:

Estruturalmente correto

Organizado

Pronto para ser tratado pelas próximas camadas

Sem ele, o sistema dependeria de leitura espalhada e frágil.