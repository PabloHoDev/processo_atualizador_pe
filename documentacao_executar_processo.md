📘 Orquestrador do Processo — executar_processo.py
📌 Propósito

O arquivo executar_processo.py é o orquestrador principal do sistema.

Ele é responsável por:

Controlar o fluxo completo
Chamar os módulos na ordem correta
Garantir que cada etapa seja executada
Consolidar o resultado final
🎯 Responsabilidade Principal

Este arquivo deve:

Integrar todos os módulos
Executar o fluxo do início ao fim
Controlar erros
Garantir execução previsível
⚠️ Importante

Este arquivo:

👉 NÃO contém lógica de negócio

Ele apenas coordena.

Se começar a ter regra dentro dele → está errado.

🔄 Fluxo de Execução
1. Carregar configurações
2. Ler planilha RPE
3. Normalizar dados
4. Validar dados
5. Ler Base Geral PE
6. Comparar obrigações
7. Identificar novas
8. Inserir novas obrigações
9. Registrar log
10. Finalizar processo
🧠 Papel na Arquitetura

Ele é o “maestro”.

Cada módulo é um instrumento.

Se um instrumento falha → ele precisa saber lidar.

🔐 Responsabilidades Críticas
- Garantir ordem correta de execução
- Interromper processo em caso de erro crítico
- Registrar falhas no log
- Evitar execução parcial inconsistente
🚫 O que NÃO deve existir aqui
❌ Validação de dados
❌ Comparação de registros
❌ Manipulação de Excel
❌ Regras de inserção

Tudo isso pertence aos módulos.

🧪 Critérios de Teste
- Execução completa com sucesso
- Execução com erro de validação
- Execução sem novas obrigações
- Execução com múltiplas inserções
📌 Conclusão

executar_processo.py não faz o trabalho pesado.

Ele garante que o trabalho aconteça da forma certa.