📘 Módulo de Normalização — normalizar_textos.py
📌 Propósito

O arquivo normalizar_textos.py é responsável por garantir que todos os dados textuais estejam em um formato padronizado antes de serem validados ou comparados.

🎯 Responsabilidade Principal

Este módulo deve:
- Limpar textos
- Padronizar formato
- Garantir consistência
⚠️ Problema que resolve

Sem normalização:

"Prefeitura" ≠ "PREFEITURA"
" João " ≠ "JOÃO"
"SP" ≠ " sp "

Isso gera:
- Falhas na comparação
- Duplicidade falsa
- Dados inconsistentes
🔄 Fluxo de Atuação

Este módulo atua logo após a leitura:

Leitura → Normalização → Validação → Comparação
🧠 Funções Esperadas
- normalizar_texto(valor)
- Remove espaços laterais
- Converte para maiúsculo
- Trata valores nulos
- normalizar_dataframe(df)
- Aplica normalização em todas as colunas relevantes
- Mantém tipos corretos
- padronizar_colunas_chave(df)

Aplica normalização nas colunas críticas:
- OBRIGAÇÃO
- UF PRESTADOR
- NOME PRESTADOR
📏 Regras de Normalização

✔ Trim (remoção de espaços)
✔ Uppercase (maiúsculo)
✔ Remoção de múltiplos espaços internos
✔ Tratamento de None

🚫 O que NÃO deve existir aqui
❌ Validação de dados
❌ Comparação
❌ Inserção
❌ Regras de negócio
🧪 Critérios de Teste
- Texto com espaços
- Texto em minúsculo
- Texto com múltiplos espaços
- Valores nulos
- Mistura de formatos
🔐 Garantias do Módulo
- Dados consistentes
- Comparação confiável
- Redução de duplicidade falsa
📌 Conclusão

Sem normalização, o sistema falha silenciosamente.

Com normalização, o sistema se torna previsível.

🏁 RESUMO FINAL

Agora você tem:

✅ Arquitetura completa
✅ TODOS os módulos documentados
✅ Fluxo definido
✅ Responsabilidades claras