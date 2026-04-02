# Controle Financeiro — Google Apps Script

Planilha de controle financeiro pessoal e empresarial gerada automaticamente no Google Sheets via Google Apps Script. Sem dependências externas — funciona 100% dentro do Google Sheets.

**Duas versões disponíveis:**

| Versão | Arquivo | Para quem |
|--------|---------|-----------|
| **Completa** | `planilha_financeira.gs` | Freelancers e empresários brasileiros (PJ/Lucro Presumido) que faturam para o exterior. Dashboard, gráficos, budget, investimentos, PJ. |
| **Simples** | `planilha_simples.gs` | Qualquer pessoa que só quer anotar entradas e saídas do mês. Sem complicação. |

---

## Versão Simples (`planilha_simples.gs`)

Cria uma aba mensal com duas seções (Entradas e Saídas) e um log de transações. Os totais por categoria atualizam automaticamente. Inclui aba **Dívidas** para acompanhar parcelas e financiamentos.

### Como usar
1. Crie um Google Sheets novo
2. Vá em **Extensões > Apps Script**
3. Cole o conteúdo de `planilha_simples.gs` e salve
4. Recarregue a planilha
5. Menu **Financeiro > Criar mês atual**

### Personalizar categorias
Edite os arrays `CAT_ENTRADA` e `CAT_SAIDA` no topo do script. Depois:

- **Mês novo** (sem dados): crie normalmente — a nova categoria já aparece no resumo e no dropdown.
- **Mês existente (com dados)**: use **Financeiro > Atualizar categorias**. O resumo (seções com totais) e os dropdowns são reconstruídos automaticamente. **Os dados do log são preservados.**

### Dívidas e parcelas
1. Menu **Financeiro > Criar / atualizar aba Dívidas**
2. Na aba "Dívidas", preencha para cada parcela/financiamento:

| Você preenche | Calculado automaticamente |
|---------------|--------------------------|
| **Descrição** — ex: "Geladeira Nubank" | **Valor mensal** = Valor total ÷ Parcelas |
| **Valor total** — ex: R$ 3.600 | **Pagas** = pagamentos no log + Anteriores |
| **Parcelas** — ex: 12 | **Restantes** = Parcelas − Pagas |
| **Início** — ex: Jan/2026 | **Saldo devedor** = Valor mensal × Restantes |
| **Anteriores** — parcelas já pagas antes da planilha | **Status** = Ativa ou Quitada (automático) |

- A descrição vira uma **categoria no dropdown** do log mensal
- Ao lançar um pagamento com essa categoria, "Pagas" atualiza sozinho
- Dívidas **quitadas** (Restantes ≤ 0) saem automaticamente do dropdown
- Para dívidas já em andamento, preencha **"Anteriores"** com as parcelas já pagas antes de usar a planilha

A linha de **Totais** no final mostra o compromisso mensal total e o saldo devedor total. Colunas em cinza contêm fórmulas — não edite.

---

## Versão Completa (`planilha_financeira.gs`)

### Funcionalidades

- **Abas mensais** (Jan–Dez) com resumo e log de transações
- **Finanças pessoais (PF)**: entradas, gastos fixos e variáveis com comparativo budget × real, investimentos por tipo
- **Finanças empresariais (PJ/CNPJ)**: faturamento, impostos (GPS, IRRF, IRPJ, CSLL, DARF), custos e saldo PJ
- **Saldo Anterior**: snapshot mensal de saldos PF + PJ (conta corrente, renda fixa/variável, cripto)
- **Dashboard** com visão consolidada do ano, gráficos e acumulado YTD
- **Gráficos**: saldo PF/PJ (linha), entradas vs gastos (barras), PJ (barras), ativos financeiros (linha), gastos fixos/variáveis (pizza), donut de gastos por categoria em cada aba mensal
- **Estrutura flexível**: adicionar ou remover categorias nos arrays ajusta o layout automaticamente
- **Células protegidas**: fórmulas ficam em cinza e exibem aviso se editadas acidentalmente
- **Aba Dívidas**: acompanhamento de parcelas e financiamentos com cálculo automático de saldo devedor
- **Fechar/Reabrir mês**: bloqueia edição de meses finalizados com indicador visual (aba verde)
- **Seletor de data** no log de transações
- **Locale pt_BR**: datas no formato dd/mm/aaaa, valores em R$, separador decimal vírgula
- **Menu "Financeiro"** com 11 ações úteis

## Configuração inicial

1. Crie uma planilha nova no Google Sheets
2. Vá em **Extensões > Apps Script**
3. Apague o código existente e cole o conteúdo de `planilha_financeira.gs`
4. Salve (Ctrl+S / Cmd+S)
5. Recarregue a planilha — o menu **Financeiro** aparecerá automaticamente
6. Clique em **Financeiro > Criar planilha completa (ano inteiro)**

## Menu Financeiro

| Ação | Descrição |
|------|-----------|
| Criar planilha completa | Cria as 12 abas mensais + Dashboard para o ano configurado |
| Novo mês... | Cria uma aba para qualquer mês/ano informado |
| Criar próximo mês automaticamente | Detecta o próximo mês a partir de hoje e cria a aba |
| Ir para o mês atual | Navega para a aba do mês corrente |
| Copiar budget do mês anterior | Copia os valores de budget do mês anterior |
| Atualizar categorias | Atualiza as listas de categorias em todas as abas mensais |
| Resumo do mês atual | Exibe um resumo rápido dos totais da aba ativa |
| Verificar meses do ano | Mostra quais meses existem e quais estão faltando |
| Fechar mês | Bloqueia a aba mensal contra edição acidental (aba fica verde) |
| Reabrir mês | Desbloqueia a aba mensal para edição novamente |
| Criar / atualizar aba Dívidas | Cria ou atualiza a aba de acompanhamento de parcelas e financiamentos |
| Como usar (abrir aba) | Abre a aba "Como usar" com instruções completas |

## Como preencher

### Diariamente — Log de transações (seção no final da aba)
- **Coluna A** — Data (seletor de data disponível)
- **Coluna B** — Descrição
- **Coluna C** — Categoria (dropdown)
- **Coluna D** — Valor (número positivo; a categoria determina se é entrada ou saída)

A categoria selecionada direciona o valor automaticamente para a seção correta no resumo acima.

### Mensalmente — Budget
- Preencha a coluna B nas seções **Gastos Fixos** e **Gastos Variáveis**
- Use **Financeiro > Copiar budget do mês anterior** para reaproveitar os valores

### Valores manuais
- **Saldo Anterior** (seção no topo, coluna C): saldo em conta corrente PF/PJ, renda fixa, renda variável, cripto e outros ativos — atualize todo mês para acompanhar sua evolução
- **Rendimento do mês** (seção Investimentos, coluna C): ganho ou perda com investimentos no mês

### Dívidas e parcelas
Use **Financeiro > Criar / atualizar aba Dívidas**. Funciona igual à versão simples — preencha Descrição, Valor total, Parcelas e Início. Parcelas pagas são contadas automaticamente a partir do log mensal. Veja a tabela detalhada na seção da versão simples acima.

### Células em cinza
Contêm fórmulas automáticas — não edite. Um aviso de confirmação aparece se você tentar.

## Configuração no script

No topo de `planilha_financeira.gs`, ajuste as constantes antes de rodar:

```javascript
const ANO = 2026; // Ano a ser gerado

const CAT_FIXO = [   // Categorias de gastos fixos
  'Moradia / Aluguel + IPTU',
  // ...adicione ou remova conforme necessário
];
```

Após alterar categorias:

- **Mês novo** (sem dados): crie normalmente — as novas categorias já aparecem no resumo e no dropdown.
- **Mês existente (com dados)**: use **Financeiro > Atualizar categorias**. O resumo (seções com totais) e os dropdowns são reconstruídos automaticamente. **Os dados do log são preservados.**

> **Atenção:** recriar uma aba existente ("Novo mês..." ou "Criar planilha completa") **apaga todos os dados** daquela aba, incluindo o log.

O layout das abas é calculado dinamicamente a partir dos arrays de categorias. Adicionar ou remover itens em qualquer array (ex: `CAT_FIXO`, `CAT_INVESTIMENTO`, `ITEMS_POS_FINANCEIRA`) ajusta automaticamente todas as posições de linha ao recriar.

## Novo ano

1. Abra **Extensões > Apps Script** e altere `const ANO = 2027`
2. Salve e volte para a planilha
3. Execute **Financeiro > Criar planilha completa**

O Dashboard existente é arquivado automaticamente como "Dashboard 2026" e um novo Dashboard 2027 é criado. Você mantém o histórico de todos os anos anteriores como abas separadas.

## Locale e formatação de números

O script configura automaticamente o locale da planilha para **pt_BR** ao rodar "Criar planilha completa". Isso garante:

- Datas no formato `dd/mm/aaaa`
- Moeda com símbolo `R$`
- Separador decimal vírgula (`1.234,56`)

Todas as fórmulas do script usam **ponto-e-vírgula** como separador de argumentos (padrão pt_BR), por exemplo: `=SUMIF($E:$E;"E";$C:$C)`. Isso é compatível com qualquer planilha em locale brasileiro.

## Contexto fiscal (PJ Lucro Presumido)

Categorias de custos pré-configuradas para empresas de serviços que faturam em moeda estrangeira:
- GPS e IRRF (mensal)
- IRPJ e CSLL (trimestral)
- DARF IRRF — Lucros e Dividendos
- PIS/COFINS não incluídos (isentos para serviços ao exterior)

## Licença

MIT
