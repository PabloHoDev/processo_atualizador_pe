📘 Módulo logs_processo
📌 Propósito

O módulo logs_processo é responsável por registrar de forma estruturada tudo o que acontece durante a execução do sistema.

Ele garante:

Rastreabilidade
Auditoria
Transparência
Capacidade de diagnóstico
🎯 Responsabilidade Principal

Este módulo deve registrar:

Execução do processo
Registros processados
Registros inseridos
Registros ignorados
Erros encontrados
Avisos gerados
⚠️ Problema que resolve

Sem log, você não consegue:

Saber o que foi inserido
Identificar erros
Auditar o processo
Confiar na automação

O sistema vira uma “caixa preta”.

📂 Estrutura do Módulo
logs_processo/
└── registrar_log_processo.py
📄 Arquivo: registrar_log_processo.py

Responsável por:

Criar estrutura de log
Registrar eventos do processo
Exportar log para arquivo (CSV ou outro formato)
🔄 Fluxo de Funcionamento
1. Início do processo → registrar início
2. Durante execução → registrar eventos
3. Ao final → consolidar informações
4. Gerar arquivo de log
📊 Tipos de Log
✔ Log de Execução
Data/hora de início
Data/hora de fim
Status geral
✔ Log de Processamento
Total de registros processados
Total de registros válidos
Total de registros inválidos
Total de novos registros
✔ Log de Inserção
Quais registros foram inseridos
Em qual aba
Em qual linha
✔ Log de Erros
Erros de validação
Erros de leitura
Erros de inserção
✔ Log de Avisos
Dados incompletos mas não críticos
Ajustes realizados
📌 Estrutura Recomendada (CSV)

Exemplo de colunas:

DataHora,Tipo,Obrigacao,UF,Acao,Status,Mensagem

Exemplo de registros:

2026-04-14 14:00,INSERCAO,OBR123,SP,INSERIDO,OK,Registro inserido com sucesso
2026-04-14 14:01,VALIDACAO,OBR456,RJ,ERRO,ERRO,UF não informada
🧠 Boas Práticas

✔ Logs devem ser claros e objetivos
✔ Evitar mensagens genéricas
✔ Padronizar tipos (OK, ERRO, AVISO)
✔ Não depender de print no console
✔ Sempre registrar erros críticos

🚫 O que NÃO deve existir aqui
❌ Regras de negócio
❌ Comparação de dados
❌ Inserção em planilha
❌ Leitura de arquivos
❌ Decisão de fluxo

Esse módulo apenas registra.

🔄 Dependências

Depende de:

configuracoes → nome do arquivo e flags
Dados vindos de todos os módulos anteriores
🧪 Critérios de Teste

Testar com:

Execução completa sem erros
Execução com erros de validação
Execução com inserções
Execução sem novas obrigações
🔐 Garantias do Módulo
Todo o processo é rastreável
Erros são identificáveis
A execução é auditável
🚀 Impacto no Sistema

Com logs bem estruturados:

Você ganha controle total
Pode confiar na automação
Facilita manutenção
Facilita evolução do sistema

Sem logs:

Você fica no escuro
📌 Conclusão

O módulo logs_processo transforma o sistema de uma automação simples em um processo confiável e auditável.

Ele não executa o processo.
Ele conta a história do que aconteceu.