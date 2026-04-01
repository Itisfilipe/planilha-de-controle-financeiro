# Controle Financeiro — Google Apps Script

Planilha de controle financeiro pessoal e empresarial gerada automaticamente no Google Sheets via Google Apps Script. Desenvolvida para freelancers e empresários brasileiros (PJ/Lucro Presumido) que faturam para clientes no exterior.

Sem dependências externas — funciona 100% dentro do Google Sheets.

## Funcionalidades

- **Abas mensais** (Jan–Dez) com resumo e log de transações
- **Finanças pessoais (PF)**: entradas, gastos fixos e variáveis com comparativo budget × real, investimentos e patrimônio
- **Finanças empresariais (PJ/CNPJ)**: faturamento, impostos (GPS, IRRF, IRPJ, CSLL, DARF), custos e saldo PJ
- **Dashboard** com visão consolidada do ano
- **Estrutura flexível**: adicionar ou remover linhas em qualquer seção não quebra os cálculos (baseado em SUMIF + tags, não em intervalos fixos)
- **Células protegidas**: fórmulas ficam em cinza e exibem aviso se editadas acidentalmente
- **Seletor de data** no log de transações
- **Locale pt_BR**: datas no formato dd/mm/aaaa, valores em R$, separador decimal vírgula
- **Menu "Financeiro"** com 9 ações úteis

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
| Atualizar dropdowns | Atualiza as listas de categorias em todas as abas mensais |
| Resumo do mês atual | Exibe um resumo rápido dos totais da aba ativa |
| Verificar meses do ano | Mostra quais meses existem e quais estão faltando |
| Instruções de uso | Exibe as instruções de preenchimento |

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
- **Posição Financeira** (seção no topo, coluna C): saldo em conta corrente PF/PJ, renda fixa, renda variável, cripto e outros ativos — atualize todo mês para acompanhar sua evolução
- **Rendimento do mês** (seção Investimentos, coluna C): ganho ou perda com investimentos no mês
- **Patrimônio** (seção Patrimônio, coluna C): valor atual de cada bem físico (imóvel, carro, etc.)

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

Após alterar categorias, execute **Financeiro > Atualizar dropdowns** para propagar a mudança.

O layout das abas é calculado dinamicamente a partir dos arrays de categorias. Adicionar ou remover itens em qualquer array (ex: `CAT_FIXO`, `CAT_INVESTIMENTO`, `ITEMS_POS_FINANCEIRA`) ajusta automaticamente todas as posições de linha — basta recriar a planilha.

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
