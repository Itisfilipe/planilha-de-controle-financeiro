# Controle Financeiro — Google Apps Script

A Google Sheets financial control spreadsheet generator for Brazilian freelancers and business owners (PJ/Lucro Presumido). Built entirely with Google Apps Script — no external dependencies.

## Features

- **Monthly tabs** (Jan–Dez) with full summary + transaction log
- **Personal finances (PF)**: income, fixed/variable expenses with budget vs. actual comparison, investments, net worth tracking
- **Business finances (PJ/CNPJ)**: revenue, taxes (GPS, IRRF, IRPJ, CSLL, DARF), costs, PJ balance
- **Dashboard** with yearly overview across all months
- **Flexible structure**: insert or remove rows in any section without breaking formulas (SUMIF-based, not range-based)
- **Protected cells**: formula cells are gray and trigger a warning if accidentally edited
- **Date picker** on transaction log entries
- **pt_BR locale**: dates as dd/mm/yyyy, currency as R$
- **Custom "Financeiro" menu** with useful actions

## Setup

1. Create a new Google Sheets spreadsheet
2. Go to **Extensions > Apps Script**
3. Delete any existing code and paste the contents of `planilha_financeira.gs`
4. Save (Ctrl+S / Cmd+S)
5. Reload the spreadsheet — the **Financeiro** menu will appear
6. Click **Financeiro > Criar planilha completa (ano inteiro)**

## Menu Actions

| Action | Description |
|--------|-------------|
| Criar planilha completa | Creates all 12 monthly tabs + Dashboard for the configured year |
| Novo mês... | Creates a single month tab (any month/year) |
| Criar próximo mês automaticamente | Detects the next month from today and creates it |
| Ir para o mês atual | Navigates to the current month's tab |
| Copiar budget do mês anterior | Copies budget values from the previous month |
| Atualizar dropdowns | Refreshes category dropdowns in all monthly tabs |
| Resumo do mês atual | Shows a quick summary of totals for the active month |
| Verificar meses do ano | Shows which months exist and which are missing |
| Instruções de uso | In-spreadsheet help text |

## How to Fill Data

### Daily: Transaction Log (rows 62+)
- **Column A** — Date (date picker available)
- **Column B** — Description
- **Column C** — Category (dropdown)
- **Column D** — Amount (positive number; the category determines if it's income or expense)

The category you select automatically routes the amount to the correct summary section above.

### Monthly: Budget
- Fill column B in the **Gastos Fixos** and **Gastos Variáveis** sections
- Use **Financeiro > Copiar budget do mês anterior** to reuse last month's values

### Manual Values
- **Rendimento do mês** (row 49, col C): investment gain or loss for the month
- **Patrimônio** (rows 52–54, col C): current value of each asset

### Cells in Gray
These contain automatic formulas — do not edit. A warning dialog will appear if you try.

## Configuration

At the top of `planilha_financeira.gs`, adjust these constants before running:

```javascript
const ANO = 2026; // Year to generate

const CAT_FIXO = [   // Fixed expense categories
  'Moradia / Aluguel + IPTU',
  // ...add or remove as needed
];
```

After changing categories, run **Financeiro > Atualizar dropdowns** to update all existing tabs.

## New Year

1. Change `const ANO = 2027` in the script
2. Run **Financeiro > Criar planilha completa**
3. The existing Dashboard is automatically archived as "Dashboard 2026"

## Tax Context (PJ Lucro Presumido)

Pre-configured cost categories for Brazilian service companies billing foreign clients:
- GPS e IRRF (monthly)
- IRPJ and CSLL (quarterly)
- DARF IRRF — Lucros e Dividendos
- PIS/COFINS not included (exempt for foreign-currency services)

## License

MIT
