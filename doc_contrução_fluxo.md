📌 Ordem Oficial de Implementação

1️⃣ configuracoes
2️⃣ utilitarios_excel
3️⃣ padronizacao_dados
4️⃣ leituras_excel
5️⃣ comparadores_base
6️⃣ validacoes_negocio
7️⃣ insercao_base_pe
8️⃣ logs_processo
9️⃣ executar_processo
🔟 testes finais

📌 PLANO DE EXECUÇÃO — processo_novas_obrigacoes_rpe
🎯 Objetivo

Implementar o pipeline completo de inserção de novas obrigações (RPE → Base Geral PE) em Python, garantindo:

-Segurança

-Idempotência

-Performance

-Rastreabilidade

-Separação de responsabilidades

🔵 FASE 1 — Preparação do Ambiente (Fundação Técnica)
1.1 Criar o repositório

Criar estrutura de pastas definida

Criar ambiente virtual (venv)

Criar .gitignore

Criar README inicial

Definir versão do Python (ex: 3.11+)

1.2 Definir dependências

Bibliotecas base:

pandas

openpyxl

python-dotenv (opcional)

loguru ou logging padrão

Criar requirements.txt

📌 Só depois disso começamos código.

🔵 FASE 2 — Construção da Base Estrutural (Infraestrutura)
Ordem correta de implementação:

2.1 configuracoes/parametros_processo.py

Começamos aqui porque:

Centraliza regras

Evita hardcode

Define colunas permitidas

Define nomes de abas válidas

Nada deve nascer sem passar por configuração.

2.2 utilitarios_excel/funcoes_excel.py

Criar:

Função para localizar colunas por nome

Função para validar existência de aba

Funções auxiliares reutilizáveis

Isso será usado por quase todos os módulos.

2.3 padronizacao_dados/normalizar_textos.py

Implementar:

Função normalize_text()

Função normalize_columns()

Toda leitura passará por aqui.

Sem padronização → sem comparação confiável.

🔵 FASE 3 — Camada de Leitura (Input Layer)
3.1 leituras_excel/ler_base_geral_pe.py

Objetivo:

Ler todas as abas válidas

Extrair coluna OBRIGAÇÃO

Construir set de obrigações existentes

Retorno esperado:

{
    "obrigacoes_existentes": set(...),
    "abas_validas": {...},
    "estrutura_colunas": {...}
}

3.2 leituras_excel/ler_planilha_rpe.py

Objetivo:

Ler RPE

Mapear colunas permitidas

Retornar DataFrame já normalizado

Sem validação ainda — apenas leitura limpa.

🔵 FASE 4 — Camada de Regras de Negócio
4.1 comparadores_base/comparar_obrigacoes_existentes.py

Função:

Recebe obrigação

Retorna:

"NOVA"

"DUPLICADA"

Baseado no set já construído.

Performance O(1).

4.2 validacoes_negocio/validar_dados_rpe.py

Implementar regras:

Obrigação obrigatória

UF obrigatória

Data válida

Valor numérico

Estrutura mínima

Retorno estruturado:

{
    "status": "VALIDO | AVISO | ERRO",
    "mensagens": [...]
}


Nada é inserido antes de passar aqui.

🔵 FASE 5 — Camada de Inserção
5.1 insercao_base_pe/inserir_novas_obrigacoes.py

Responsável por:

Determinar aba destino (regras já existentes)

Encontrar próxima linha vazia

Inserir registro

Atualizar set de obrigações existentes

Inserção deve ser:

Controlada

Em lote

Sem sobrescrever fórmulas

🔵 FASE 6 — Logging Profissional
6.1 logs_processo/registrar_log_processo.py

Criar logger estruturado:

Cada linha processada gera:

timestamp

obrigação

aba destino

resultado

mensagem

linha RPE

Resumo final:

Total lido

Total inserido

Total duplicado

Total erro

🔵 FASE 7 — Orquestração Final
7.1 executar_processo.py

Este será criado por último.

Fluxo:

Carregar configurações

Ler Base Geral PE

Ler RPE

Loop pelas linhas da RPE:

Normalizar

Validar

Verificar duplicidade

Inserir se necessário

Logar resultado

Salvar arquivo

Exibir resumo final

Orquestrador não contém regra.
Apenas coordena.

🔵 FASE 8 — Testes Controlados

Antes de considerar finalizado:

Testes obrigatórios:

RPE com obrigação nova

RPE com obrigação duplicada

RPE com campo obrigatório vazio

RPE com data inválida

RPE com coluna fora de ordem

Execução repetida (idempotência)

Se passar nesses testes → sistema aprovado.

🔵 FASE 9 — Hardening (Profissionalização Final)

Antes de declarar “produção”:

Adicionar tratamento de exceções globais

Garantir rollback seguro em erro crítico

Proteger contra arquivo aberto em Excel

Criar logs de erro separados

Versionar (ex: v1.0.0)

📌 Ordem Oficial de Implementação

1️⃣ configuracoes
2️⃣ utilitarios_excel
3️⃣ padronizacao_dados
4️⃣ leituras_excel
5️⃣ comparadores_base
6️⃣ validacoes_negocio
7️⃣ insercao_base_pe
8️⃣ logs_processo
9️⃣ executar_processo
🔟 testes finais

🏁 Critério de Conclusão do Projeto

Consideramos o fluxo finalizado quando:

Nenhuma duplicidade é criada

Logs estão completos

Processo é idempotente

Execução repetida gera mesmo resultado

Código está modular e legível

README está atualizado

Versão está marcada