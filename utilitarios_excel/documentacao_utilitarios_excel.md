📘 Módulo de Utilitários Excel — utilitarios_excel
📌 Visão Geral

O módulo utilitarios_excel contém funções auxiliares responsáveis por operações genéricas relacionadas a arquivos Excel.

Ele é um módulo de infraestrutura técnica, não de regra de negócio.

Seu objetivo é:

Centralizar manipulação de planilhas

Evitar duplicação de código

Padronizar leitura e escrita de dados

Garantir robustez no tratamento de Excel

🏗 Papel na Arquitetura do Projeto

Dentro do fluxo processo_novas_obrigacoes_rpe, este módulo:

Apoia carregadores (carregadores)

Apoia inseridores (inseridores)

Apoia validadores

Apoia comparadores

Ele atua como uma camada técnica intermediária entre o sistema e os arquivos Excel.

🎯 Responsabilidade do Módulo

Este módulo deve conter funções como:

Listar abas de uma planilha

Verificar existência de aba

Buscar índice de coluna por nome

Validar estrutura mínima de planilha

Padronizar leitura com pandas

Salvar planilha com segurança

Criar aba se não existir

Remover abas ignoradas

Ele não deve conter lógica específica de RPE ou Base Geral PE.

📂 Arquivo Principal
utilitarios_excel/
└── funcoes_excel.py

🧠 Tipos de Funções Esperadas
📑 Manipulação de Abas

listar_abas(caminho_arquivo)

verificar_aba_existe(caminho_arquivo, nome_aba)

filtrar_abas_validas(lista_abas, abas_ignoradas)

📊 Manipulação de Colunas

buscar_indice_coluna(df, nome_coluna)

validar_colunas_existentes(df, colunas_obrigatorias)

padronizar_nomes_colunas(df)

📥 Leitura de Planilha

ler_planilha(caminho, aba)

ler_todas_abas_validas(caminho, abas_ignoradas)

Sempre com tratamento de erro estruturado.

📤 Escrita de Planilha

salvar_planilha(df, caminho, aba)

anexar_dados_em_aba_existente(df, caminho, aba)

criar_aba_se_nao_existir(caminho, aba)

🔐 Diretrizes Técnicas

Este módulo deve:

Utilizar pandas

Utilizar openpyxl quando necessário

Tratar exceções de arquivo inexistente

Tratar exceções de aba inexistente

Retornar erros claros e padronizados

Nunca imprimir diretamente no console (usar log quando necessário)

🚫 O que NÃO deve existir aqui

❌ Regras de validação de RPE

❌ Lógica de matching de obrigação

❌ Decisão de onde inserir dados

❌ Tratamento específico de área de pendência

Este módulo é neutro e reutilizável.

🧱 Benefícios Arquiteturais

Separar utilitários Excel garante:

Código mais limpo

Módulos de negócio mais simples

Testabilidade isolada

Manutenção facilitada

Reaproveitamento em outros projetos

🔄 Relação com Outros Módulos
Módulo	Dependência
carregadores	Usa para leitura
validadores	Usa para verificar estrutura
comparadores	Usa para garantir consistência
inseridores	Usa para escrita
logger	Pode usar para exportação
🚀 Evolução Futura

Este módulo poderá futuramente incluir:

Controle de versão de planilha

Backup automático antes de escrita

Controle de concorrência

Suporte a CSV

Suporte a múltiplos formatos

📍 Conclusão

utilitarios_excel é a base técnica da manipulação de planilhas do sistema.

Ele garante padronização, segurança e desacoplamento entre lógica de negócio e operações de arquivo.

Sem ele, o projeto vira um conjunto de scripts frágeis.

Com ele, o projeto se mantém estruturado e escalável.