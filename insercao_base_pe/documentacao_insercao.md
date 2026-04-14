📘 Módulo insercao_base_pe

📌 Propósito

O módulo insercao_base_pe é responsável por realizar a inserção controlada de novas obrigações na Base Geral PE.

Ele é o ponto onde o sistema efetivamente altera os dados.

Por isso, é um dos módulos mais sensíveis e críticos de todo o projeto.

🎯 Responsabilidade Principal

Este módulo deve garantir que:

- Apenas obrigações novas sejam inseridas
- Nenhum dado existente seja sobrescrito
- A inserção ocorra na aba correta
- A estrutura da planilha seja preservada
- O processo seja seguro e previsível
⚠️ Risco que esse módulo controla

Sem esse módulo bem definido, o sistema pode:
- Duplicar registros
- Inserir na aba errada
- Quebrar fórmulas existentes
- Corromper a base geral
- Perder rastreabilidade

Este módulo é o “ponto de impacto” do sistema.

📂 Estrutura do Módulo
insercao_base_pe/
└── inserir_novas_obrigacoes.py
📄 Arquivo: inserir_novas_obrigacoes.py

Responsável por:

- Receber registros já validados e classificados como novos
- Determinar a aba correta de inserção
- Inserir os dados na próxima linha disponível
- Garantir integridade da planilha

🔄 Fluxo de Funcionamento
1. Receber lista de novas obrigações
2. Identificar aba de destino (regra de negócio)
3. Localizar próxima linha vazia
4. Inserir dados nas colunas corretas
5. Preservar fórmulas e estrutura
6. Retornar resultado da inserção
🧠 Regras Fundamentais
✔ Inserir apenas registros novos

A decisão de “novo ou não” já vem do módulo comparadores_base.

Este módulo não decide, apenas executa.

✔ Não sobrescrever dados existentes

Nunca:
- Alterar células já preenchidas
- Modificar registros antigos
✔ Respeitar a estrutura da aba
- Não alterar cabeçalhos
- Não alterar fórmulas
- Não alterar formatação
✔ Inserção sempre no final

Os novos registros devem ser inseridos:

Na próxima linha disponível
Após o último registro válido
✔ Seguir regras de distribuição

A escolha da aba deve seguir:

UF Prestador (ou regra definida)
Mesma lógica do processo atual (macro/VBA)
🔐 Garantias do Módulo

Este módulo deve garantir:

Integridade da base
Consistência dos dados
Previsibilidade do comportamento
Segurança na escrita
🚫 O que NÃO deve existir aqui
❌ Validação de dados
❌ Normalização de texto
❌ Comparação de obrigações
❌ Regras de duplicidade
❌ Leitura de planilha

Essas responsabilidades pertencem a outros módulos.

🔄 Dependências do Módulo

Este módulo depende de:

configuracoes → regras e colunas
utilitarios_excel → escrita em Excel
comparadores_base → saber o que é novo
🧪 Critérios de Teste

Este módulo deve ser testado com:

- Inserção de 1 nova obrigação
- Inserção de múltiplas obrigações
- Inserção em diferentes abas
- Verificação de não sobrescrita
- Execução repetida (idempotência)
🚀 Impacto no Sistema

Se este módulo estiver correto:

- O sistema será confiável
- A base permanecerá íntegra
- O processo será automatizável

Se estiver errado:

Todo o sistema perde valor
📌 Conclusão

O módulo insercao_base_pe é o responsável por transformar decisão em ação.

Ele não pensa.
Ele executa — com precisão e segurança.