📘 Módulo comparadores_base
📌 Propósito

O módulo comparadores_base é responsável por identificar se uma obrigação da RPE:

Já existe na Base Geral PE
Ou é uma nova obrigação que deve ser inserida

Ele é o responsável por evitar duplicidade de dados.

🎯 Responsabilidade Principal

Este módulo deve:

- Comparar dados da RPE com a Base Geral PE
- Identificar registros existentes
- Identificar registros novos
- Retornar listas separadas para decisão posterior
⚠️ Problema que resolve

Sem esse módulo, o sistema pode:

- Inserir registros duplicados
- Perder controle da base
- Gerar inconsistência nos dados
- Comprometer relatórios e análises

Este módulo garante integridade lógica.

📂 Estrutura do Módulo
comparadores_base/
└── comparar_obrigacoes_existentes.py
📄 Arquivo: comparar_obrigacoes_existentes.py

Responsável por executar a comparação entre:

Dados da RPE (já validados e padronizados)
Dados da Base Geral PE (já carregados)
🔄 Fluxo de Funcionamento
1. Receber dados válidos da RPE
2. Receber base consolidada da Base Geral PE
3. Comparar utilizando coluna chave (OBRIGAÇÃO)
4. Identificar correspondências
5. Separar registros
6. Retornar resultado estruturado
🧠 Regra Principal de Comparação

A comparação deve ser baseada em:

Coluna chave: OBRIGAÇÃO

⚠️ Importante:

Os dados já devem estar padronizados (maiúsculo, sem espaços, etc).

Este módulo não corrige dados — ele apenas compara.

📊 Saída Esperada

O módulo deve retornar algo como:

{
    "novos": [...],
    "existentes": [...]
}

Ou em formato estruturado:

{
    "novos": lista_de_registros_novos,
    "existentes": lista_de_registros_existentes,
    "total_processado": 100,
    "total_novos": 20,
    "total_existentes": 80
}
📏 Regras Fundamentais
✔ Não alterar dados

Este módulo:

Não modifica valores
Não normaliza
Não valida

Ele apenas compara.

✔ Comparação determinística

A mesma entrada deve sempre gerar o mesmo resultado.

Sem aleatoriedade. Sem ambiguidade.

✔ Performance

A comparação deve ser eficiente:

Evitar loops desnecessários
Preferir estruturas como set ou dict

(Principalmente para grandes volumes)

🚫 O que NÃO deve existir aqui
❌ Inserção de dados
❌ Validação de dados
❌ Leitura de planilhas
❌ Padronização de texto
❌ Escolha de aba

Se estiver decidindo “onde inserir”, está errado.

🔄 Dependências

Depende de:

- padronizacao_dados → dados limpos
- validacoes_negocio → dados válidos
- leituras_excel → base carregada
🧪 Critérios de Teste

Testar com:

- Todos registros novos
- Todos registros existentes
- Mistura de ambos
- Dados duplicados na RPE
- Dados com pequenas variações (já normalizados)
🔐 Garantias do Módulo
- Nenhuma obrigação existente será duplicada
- Novas obrigações serão corretamente identificadas
- O fluxo de inserção será confiável
🚀 Impacto no Sistema

Se esse módulo for sólido:

- O sistema ganha precisão
- A base se mantém íntegra
- A automação se torna confiável

Se falhar:

Você automatiza erro em escala

📌 Conclusão

O módulo comparadores_base é o responsável por decidir:

👉 “Isso já existe ou não?”

Ele não executa ações.
Ele define o que deve ser feito.