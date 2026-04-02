# Controle Financeiro — Google Apps Script

Planilha de controle financeiro pessoal gerada automaticamente no Google Sheets via Google Apps Script. Sem dependências externas — funciona 100% dentro do Google Sheets.

## Funcionalidades

- **Abas mensais** com resumo de entradas e saídas por categoria
- **Log de transações** com dropdown de categorias e seletor de data
- **Flag "Parcela?"** no log para identificar pagamentos de dívidas
- **Aba Dívidas** para acompanhar parcelas e financiamentos (status Ativa/Quitada automático)
- **Atualizar categorias** sem perder dados (reconstrói resumo, preserva log)
- **Aba "Como usar"** com instruções completas (criada automaticamente)
- **Células protegidas** com fundo cinza e aviso ao editar
- **Locale pt_BR**: datas dd/mm/aaaa, moeda R$, vírgula decimal

## Como usar

1. Crie um Google Sheets novo
2. Vá em **Extensões > Apps Script**
3. Cole o conteúdo de `planilha.gs` e salve
4. Recarregue a planilha
5. Menu **Financeiro > Criar mês atual**

### Preenchimento diário

No **LOG** (parte de baixo de cada aba mensal):

| Coluna | O que preencher |
|--------|----------------|
| **A** — Data | Seletor de data disponível |
| **B** — Descrição | Ex: "Supermercado Extra" |
| **C** — Categoria | Dropdown (Salário, Alimentação, etc.) |
| **D** — Valor | Sempre positivo — a categoria determina se é entrada ou saída |
| **E** — Parcela? | Sim/Não — marque "Sim" se for pagamento de parcela |

Os totais por categoria no resumo (parte de cima) atualizam automaticamente.

### Menu Financeiro

| Ação | Descrição |
|------|-----------|
| Criar mês atual | Cria a aba do mês corrente |
| Novo mês... | Cria qualquer mês/ano (ex: Jan/2027) |
| Resumo do mês | Exibe totais de entradas, saídas e saldo |
| Atualizar categorias | Recria o resumo com categorias atuais (log preservado) |
| Criar / atualizar aba Dívidas | Cria ou atualiza a aba de parcelas e financiamentos |
| Como usar (abrir aba) | Abre a aba com instruções completas |

## Personalizar categorias

Edite os arrays `CAT_ENTRADA` e `CAT_SAIDA` no topo do script. Depois:

- **Mês novo** (sem dados): crie normalmente — a nova categoria já aparece no resumo e no dropdown.
- **Mês existente (com dados)**: use **Financeiro > Atualizar categorias**. O resumo é reconstruído e os dados do log são **preservados e migrados** automaticamente se o layout mudou.

> **Atenção:** "Criar mês atual" e "Novo mês" **não recriam** abas existentes — se a aba já existe, apenas navegam até ela. Para recomeçar um mês, exclua a aba manualmente antes.

## Dívidas e parcelas

Use **Financeiro > Criar / atualizar aba Dívidas**.

| Você preenche | Calculado automaticamente |
|---------------|--------------------------|
| **Descrição** — ex: "Geladeira Nubank" | **Valor mensal** = Valor total ÷ Parcelas |
| **Valor total** — ex: R$ 3.600 | **Restantes** = Parcelas − Parcelas pagas |
| **Parcelas** — ex: 12 | **Saldo devedor** = Valor mensal × Restantes |
| **Início** — ex: Jan/2026 | **Status** = Ativa ou Quitada (automático) |
| **Parcelas pagas** — atualize todo mês | |

A aba Dívidas é independente das abas mensais — atualize "Parcelas pagas" manualmente a cada mês. No log mensal, lance o pagamento usando uma categoria normal (ex: "Cartão de crédito") e marque "Parcela? = Sim".

## Locale e formatação

O script configura automaticamente o locale para **pt_BR**. Todas as fórmulas usam **ponto-e-vírgula** como separador (padrão brasileiro).

## Licença

MIT
