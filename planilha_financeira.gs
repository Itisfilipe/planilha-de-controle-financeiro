/**
 * CONTROLE FINANCEIRO — Google Apps Script
 *
 * PRIMEIRA VEZ:
 * 1. Crie um Google Sheets novo
 * 2. Extensões > Apps Script
 * 3. Cole este código e salve (Ctrl+S)
 * 4. No menu "Financeiro" → "Criar planilha completa"
 *
 * MESES SEGUINTES:
 * Menu "Financeiro" → "Novo mês..." → ex: Jan/2027
 *
 * ADICIONAR CATEGORIA NOVA:
 * 1. Adicione o nome no array da seção correspondente (CAT_FIXO, CAT_VARIAVEL, etc.)
 * 2. Menu "Financeiro" → "Atualizar dropdowns"
 * 3. Insira a linha na aba do mês, copie a tag da coluna E da linha vizinha
 *
 * TAGS DE SEÇÃO (coluna E, invisível):
 * E=Entradas  F=Fixos  V=Variáveis  PJF=PJ Faturamento  PJC=PJ Custos  IA=Aporte  PAT=Patrimônio
 */

// ─── CONFIGURAÇÃO ─────────────────────────────────────────────────────────────

const ANO = 2026;

const MESES = [
  { nome: 'Janeiro',   abrev: 'Jan' },
  { nome: 'Fevereiro', abrev: 'Fev' },
  { nome: 'Março',     abrev: 'Mar' },
  { nome: 'Abril',     abrev: 'Abr' },
  { nome: 'Maio',      abrev: 'Mai' },
  { nome: 'Junho',     abrev: 'Jun' },
  { nome: 'Julho',     abrev: 'Jul' },
  { nome: 'Agosto',    abrev: 'Ago' },
  { nome: 'Setembro',  abrev: 'Set' },
  { nome: 'Outubro',   abrev: 'Out' },
  { nome: 'Novembro',  abrev: 'Nov' },
  { nome: 'Dezembro',  abrev: 'Dez' },
];

const LOG_ROW = 72; // Primeira linha de dados do log em cada aba mensal

// Tags de seção — coluna E oculta. Totais usam SUMIF nessas tags:
// inserir/remover linhas não quebra nada; basta copiar a tag da linha vizinha.
const TAG = {
  entrada:      'E',
  fixo:         'F',
  variavel:     'V',
  pjFat:        'PJF',
  pjCusto:      'PJC',
  invAporte:    'IA',
  pat:          'PAT',
  posFinanceira:'POS',
};

// Itens da seção Posição Financeira — valores manuais (snapshot do saldo atual)
const ITEMS_POS_FINANCEIRA = [
  'Conta corrente / Poupança',
  'Renda Fixa (CDB, LCI, Tesouro)',
  'Renda Variável (Ações, FIIs, ETFs)',
  'Criptomoedas',
  'Outros ativos financeiros',
];

const COR = {
  titulo:     '#1a1a2e',
  tituloFonte:'#ffffff',
  secao:      '#2c3e50',
  secaoFonte: '#ffffff',
  total:      '#dde8f0',
  saldo:      '#1a1a2e',
  saldoFonte: '#ffffff',
  pj:         '#1a4731',
  pjFonte:    '#ffffff',
  pjTotal:    '#d4edda',
  logHeader:  '#34495e',
  logFonte:   '#ffffff',
  verdeClaro: '#c8e6c9',
  verdeFonte: '#1b5e20',
  vermClaro:  '#ffcdd2',
  vermFonte:  '#b71c1c',
  protegido:  '#eeeeee', // fundo cinza = célula com fórmula, não editar
};

const FMT_BRL = 'R$ #,##0.00';

// Colunas de dados no dashboard (exclui cols 7 e 11 que são separadores visuais)
const DASH_COLS = [2, 3, 4, 5, 6, 8, 9, 10, 12];

// ─── CATEGORIAS ───────────────────────────────────────────────────────────────
// Fonte única de verdade: os arrays de seção definem os nomes.
// CATEGORIAS é construída a partir deles — renomear aqui atualiza tudo.

const CAT_ENTRADA = [
  'Pró-labore',
  'Distribuição de lucros',
  'Aluguel recebido',
  'Rendimento investimentos',
  'Outros entrada',
];

const CAT_FIXO = [
  'Moradia / Aluguel + IPTU',
  'Condomínio',
  'Energia Elétrica',
  'Água',
  'Internet',
  'Celular',
  'Seguro carro / vida',
  'Assinaturas',
];

const CAT_VARIAVEL = [
  'Mercado / Feira',
  'Refeições fora',
  'Combustível / Carro',
  'Farmácia / Saúde',
  'Lazer',
  'Vestuário',
  'Educação / Cursos',
  'Outros',
];

const CAT_PJ_CUSTO = [
  'GPS e IRRF',
  'IRPJ',
  'CSLL',
  'DARF IRRF - Lucros e Dividendos',
  'Contador',
  'Softwares / Ferramentas',
  'Plano de Saúde PJ',
  'Outros PJ',
  'Pró-labore PJ',
  'Distribuição de lucros PJ',
];

const CATEGORIAS = [
  ...CAT_ENTRADA,
  ...CAT_FIXO,
  ...CAT_VARIAVEL,
  'Faturamento PJ',
  ...CAT_PJ_CUSTO,
  'Aporte',
];

// ─── MENU ─────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Financeiro')
    .addItem('Criar planilha completa (ano inteiro)',    'criarPlanilhaFinanceira')
    .addSeparator()
    .addItem('Novo mês...',                              'criarNovoMes')
    .addItem('Criar próximo mês automaticamente',        'criarProximoMes')
    .addItem('Ir para o mês atual',                      'irParaMesAtual')
    .addSeparator()
    .addItem('Copiar budget do mês anterior',            'copiarBudgetMesAnterior')
    .addItem('Atualizar dropdowns de todas as abas',     'atualizarDropdowns')
    .addSeparator()
    .addItem('Resumo do mês atual',                      'resumoMesAtual')
    .addItem('Verificar meses do ano',                   'verificarMesesAno')
    .addSeparator()
    .addItem('Instruções de uso',                        'mostrarInstrucoes')
    .addToUi();
}

// ─── CRIAR PLANILHA COMPLETA ──────────────────────────────────────────────────

function criarPlanilhaFinanceira() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const abasExistentes = MESES.filter(({ abrev }) => ss.getSheetByName(`${abrev}/${ANO}`));
  const aviso = abasExistentes.length > 0
    ? `Já existem ${abasExistentes.length} aba(s) de ${ANO} — serão apagadas e recriadas.`
    : `Serão criadas 12 abas mensais (Jan/${ANO} a Dez/${ANO}) + Dashboard.`;

  const ok = ui.alert(
    `Criar planilha completa — ${ANO}`,
    `O que será feito:\n• ${aviso}\n• Cada aba terá resumo, budget e log de transações.\n\nContinuar?`,
    ui.ButtonSet.YES_NO
  );
  if (ok !== ui.Button.YES) return;

  MESES.forEach(({ nome, abrev }) => {
    montarAbaMensal(getOrCreateSheet(ss, `${abrev}/${ANO}`), nome, ANO);
  });

  criarDashboard(ss);

  ['Planilha1', 'Sheet1'].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s && ss.getSheets().length > 1) try { ss.deleteSheet(s); } catch (e) {}
  });

  const dash = ss.getSheetByName('Dashboard');
  if (dash) { ss.setActiveSheet(dash); ss.moveActiveSheet(1); }

  SpreadsheetApp.flush();
  ui.alert('Planilha criada com sucesso!');
}

// ─── NOVO MÊS ─────────────────────────────────────────────────────────────────

function criarNovoMes() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const resposta = ui.prompt(
    'Criar Novo Mês',
    'Cria uma nova aba mensal com resumo, budget e log de transações.\n\nDigite no formato Abrev/Ano — exemplo: Jan/2027',
    ui.ButtonSet.OK_CANCEL
  );
  if (resposta.getSelectedButton() !== ui.Button.OK) return;

  const partes = resposta.getResponseText().trim().split('/');
  if (partes.length !== 2) { ui.alert('Formato inválido. Use: Jan/2027'); return; }

  const abrevInput = partes[0].trim();
  const anoInput   = parseInt(partes[1].trim(), 10);
  const idx        = MESES.findIndex(m => m.abrev.toLowerCase() === abrevInput.toLowerCase());

  if (idx === -1) {
    ui.alert(`Mês inválido: "${abrevInput}"\nUse: Jan, Fev, Mar, Abr, Mai, Jun, Jul, Ago, Set, Out, Nov, Dez`);
    return;
  }
  if (isNaN(anoInput) || anoInput < 2020 || anoInput > 2100) {
    ui.alert(`Ano inválido: "${partes[1]}"`);
    return;
  }

  const { abrev, nome } = MESES[idx];
  const nomeMes = `${abrev}/${anoInput}`;

  if (ss.getSheetByName(nomeMes)) {
    const ok = ui.alert(`A aba "${nomeMes}" já existe.`, 'Recriar do zero? (dados serão perdidos)', ui.ButtonSet.YES_NO);
    if (ok !== ui.Button.YES) return;
  }

  const sheetNova = getOrCreateSheet(ss, nomeMes);
  montarAbaMensal(sheetNova, nome, anoInput);
  ss.setActiveSheet(sheetNova);
  ui.alert(`Aba "${nomeMes}" criada com sucesso!`);
}

// ─── IR PARA MÊS ATUAL ────────────────────────────────────────────────────────

function irParaMesAtual() {
  const ui      = SpreadsheetApp.getUi();
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const hoje    = new Date();
  const nomeMes = `${MESES[hoje.getMonth()].abrev}/${hoje.getFullYear()}`;
  const sheet   = ss.getSheetByName(nomeMes);

  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    ui.alert(
      'Mês atual não encontrado',
      `A aba "${nomeMes}" não existe ainda.\nUse Financeiro › "Novo mês..." para criá-la.`,
      ui.ButtonSet.OK
    );
  }
}

// ─── COPIAR BUDGET DO MÊS ANTERIOR ────────────────────────────────────────────

function copiarBudgetMesAnterior() {
  const ui         = SpreadsheetApp.getUi();
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const sheetAtual = ss.getActiveSheet();
  const nomeAtual  = sheetAtual.getName();

  if (!/^[A-Za-z]{3}\/\d{4}$/.test(nomeAtual)) {
    ui.alert('Abra uma aba mensal antes de usar esta função.\nExemplo: Jan/2026');
    return;
  }

  const [abrevAtual, anoStr] = nomeAtual.split('/');
  const anoAtual    = parseInt(anoStr, 10);
  const idxAtual    = MESES.findIndex(m => m.abrev === abrevAtual);
  if (idxAtual === -1) { ui.alert(`Abreviação de mês não reconhecida: "${abrevAtual}"`); return; }
  const idxAnterior = idxAtual === 0 ? 11 : idxAtual - 1;
  const anoAnterior = idxAtual === 0 ? anoAtual - 1 : anoAtual;
  const nomeAnt     = `${MESES[idxAnterior].abrev}/${anoAnterior}`;
  const sheetAnt    = ss.getSheetByName(nomeAnt);

  if (!sheetAnt) { ui.alert(`Aba "${nomeAnt}" não encontrada.`); return; }

  const ok = ui.alert(
    'Copiar Budget',
    `Copiar budget de "${nomeAnt}" → "${nomeAtual}"?\nOs valores de budget atuais serão substituídos.`,
    ui.ButtonSet.YES_NO
  );
  if (ok !== ui.Button.YES) return;

  // Constrói mapa label → budget a partir do mês anterior
  const linhasAnt  = sheetAnt.getLastRow();
  const tagsAnt    = sheetAnt.getRange(1, 5, linhasAnt).getValues();
  const labelsAnt  = sheetAnt.getRange(1, 1, linhasAnt).getValues();
  const budgetsAnt = sheetAnt.getRange(1, 2, linhasAnt).getValues();

  const mapaB = {};
  tagsAnt.forEach((t, i) => {
    if (t[0] === TAG.fixo || t[0] === TAG.variavel) {
      const label = labelsAnt[i][0];
      if (label && budgetsAnt[i][0] !== '') mapaB[label] = budgetsAnt[i][0];
    }
  });

  // Aplica no mês atual por label (robusto a linhas inseridas/removidas)
  const linhasAtual  = sheetAtual.getLastRow();
  const tagsAtual    = sheetAtual.getRange(1, 5, linhasAtual).getValues();
  const labelsAtual  = sheetAtual.getRange(1, 1, linhasAtual).getValues();
  let copiados = 0;

  tagsAtual.forEach((t, i) => {
    if ((t[0] === TAG.fixo || t[0] === TAG.variavel) && mapaB[labelsAtual[i][0]] !== undefined) {
      sheetAtual.getRange(i + 1, 2).setValue(mapaB[labelsAtual[i][0]]);
      copiados++;
    }
  });

  ui.alert(`Budget copiado! ${copiados} categoria(s) atualizada(s).`);
}

// ─── ATUALIZAR DROPDOWNS ──────────────────────────────────────────────────────

function atualizarDropdowns() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const ok = ui.alert(
    'Atualizar Dropdowns',
    'Atualiza a lista de categorias no log de transações em todas as abas mensais.\n\nNenhum dado existente será apagado. Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (ok !== ui.Button.YES) return;

  const validacao = SpreadsheetApp.newDataValidation()
    .requireValueInList(CATEGORIAS, true)
    .setAllowInvalid(false)
    .build();

  let count = 0;
  ss.getSheets().forEach(sheet => {
    if (/^[A-Za-z]{3}\/\d{4}$/.test(sheet.getName())) {
      sheet.getRange(`C${LOG_ROW}:C2000`).setDataValidation(validacao);
      count++;
    }
  });

  ui.alert(`Dropdowns atualizados em ${count} aba(s).`);
}

// ─── CRIAR PRÓXIMO MÊS ───────────────────────────────────────────────────────

function criarProximoMes() {
  const ui    = SpreadsheetApp.getUi();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const hoje  = new Date();
  const idx   = hoje.getMonth();                          // 0–11
  const anoH  = hoje.getFullYear();
  const idxP  = (idx + 1) % 12;
  const anoP  = idx === 11 ? anoH + 1 : anoH;
  const nomeP = `${MESES[idxP].abrev}/${anoP}`;

  const ok = ui.alert(
    'Criar Próximo Mês',
    `O que será feito:\n• Criar a aba "${nomeP}" com resumo, budget e log de transações.` +
    (ss.getSheetByName(nomeP) ? `\n\nAviso: a aba "${nomeP}" já existe e será recriada (dados serão perdidos).` : '') +
    '\n\nContinuar?',
    ui.ButtonSet.YES_NO
  );
  if (ok !== ui.Button.YES) return;

  const sheetProx = getOrCreateSheet(ss, nomeP);
  montarAbaMensal(sheetProx, MESES[idxP].nome, anoP);
  ss.setActiveSheet(sheetProx);
  ui.alert(`Aba "${nomeP}" criada com sucesso!`);
}

// ─── RESUMO DO MÊS ATUAL ─────────────────────────────────────────────────────

function resumoMesAtual() {
  const ui    = SpreadsheetApp.getUi();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const nome  = sheet.getName();

  if (!/^[A-Za-z]{3}\/\d{4}$/.test(nome)) {
    ui.alert('Resumo indisponível', 'Abra uma aba mensal (ex: Jan/2026) antes de usar esta função.', ui.ButtonSet.OK);
    return;
  }

  const lastRow  = sheet.getLastRow();
  const tags     = sheet.getRange(1, 5, lastRow).getValues(); // col E
  const reais    = sheet.getRange(1, 3, lastRow).getValues(); // col C

  const sumif = tag => tags.reduce((s, [t], i) => t === tag ? s + (Number(reais[i][0]) || 0) : s, 0);

  const fmt = v => 'R$ ' + v.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  const entradas  = sumif(TAG.entrada);
  const fixos     = sumif(TAG.fixo);
  const variaveis = sumif(TAG.variavel);
  const aportes   = sumif(TAG.invAporte);
  const saldoPF   = entradas - fixos - variaveis - aportes;
  const pjFat     = sumif(TAG.pjFat);
  const pjCusto   = sumif(TAG.pjCusto);
  const saldoPJ   = pjFat - pjCusto;
  const posTotal  = sumif(TAG.posFinanceira);

  const linha = (label, valor) => `\n  ${label.padEnd(24)} ${fmt(valor)}`;

  ui.alert(
    `Resumo — ${nome}`,
    '── PESSOAL (PF) ─────────────────────────' +
    linha('Entradas',         entradas)  +
    linha('Gastos Fixos',     fixos)     +
    linha('Gastos Variáveis', variaveis) +
    linha('Aportes',          aportes)   +
    linha('SALDO DO MÊS',    saldoPF)   +
    '\n\n── PJ / CNPJ ────────────────────────────' +
    linha('Faturamento PJ',   pjFat)     +
    linha('Custos PJ',        pjCusto)   +
    linha('SALDO PJ',         saldoPJ)   +
    '\n\n── POSIÇÃO FINANCEIRA ───────────────────' +
    linha('Total ativos financeiros', posTotal),
    ui.ButtonSet.OK
  );
}

// ─── VERIFICAR MESES DO ANO ───────────────────────────────────────────────────

function verificarMesesAno() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const existentes = [], faltantes = [];
  MESES.forEach(({ abrev }) => {
    (ss.getSheetByName(`${abrev}/${ANO}`) ? existentes : faltantes).push(abrev);
  });

  ui.alert(
    `Meses do ano ${ANO}`,
    `Criados (${existentes.length}/12):\n  ${existentes.join(', ') || '—'}` +
    `\n\nFaltando (${faltantes.length}/12):\n  ${faltantes.join(', ') || 'Nenhum — todos os meses existem!'}` +
    (faltantes.length > 0 ? '\n\nUse Financeiro › "Novo mês..." para criar os que faltam.' : ''),
    ui.ButtonSet.OK
  );
}

// ─── INSTRUÇÕES DE USO ────────────────────────────────────────────────────────

function mostrarInstrucoes() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Instruções de uso — Controle Financeiro',
    'PREENCHIMENTO DIÁRIO\n' +
    '  • Anote cada transação no LOG (linhas a partir de 62):\n' +
    '    Data | Descrição | Categoria | Valor\n' +
    '  • A Categoria (dropdown) determina em qual seção o valor aparece.\n' +
    '  • Gastos lançados como positivos; o sistema subtrai automaticamente.\n\n' +
    'BUDGET MENSAL\n' +
    '  • Preencha a coluna B nas seções Gastos Fixos e Variáveis.\n' +
    '  • Use "Copiar budget do mês anterior" para reaproveitar os valores.\n\n' +
    'VALORES MANUAIS\n' +
    '  • Rendimento do mês (linha 49): ganho ou perda de investimentos.\n' +
    '  • Patrimônio (linhas 52–54): valor atual de cada bem físico.\n' +
    '  • Posição Financeira (linhas 60–64): saldo em conta, investimentos,\n' +
    '    renda fixa/variável, cripto — atualize todo mês para acompanhar.\n\n' +
    'CÉLULAS EM CINZA\n' +
    '  • Contêm fórmulas automáticas — não edite.\n' +
    '  • Um aviso aparecerá se você tentar editar uma dessas células.\n\n' +
    'NOVO ANO\n' +
    '  • Mude a constante ANO no script e use "Criar planilha completa".\n' +
    '  • O Dashboard antigo é arquivado automaticamente.\n\n' +
    'NOVA CATEGORIA\n' +
    '  • Adicione o nome no array correto no script (CAT_FIXO, etc.).\n' +
    '  • Use "Atualizar dropdowns" para propagar a mudança.',
    ui.ButtonSet.OK
  );
}

// ─── ABA MENSAL ───────────────────────────────────────────────────────────────

/*
 * Layout fixo de linhas:
 *  1        — Título
 *  3        — ENTRADAS            tag: E
 *  4–8      — itens (CAT_ENTRADA)
 *  9        — TOTAL ENTRADAS
 *  11       — GASTOS FIXOS        tag: F
 *  12–19    — itens (CAT_FIXO)
 *  20       — TOTAL FIXOS
 *  22       — GASTOS VARIÁVEIS    tag: V
 *  23–30    — itens (CAT_VARIAVEL)
 *  31       — TOTAL VARIÁVEIS
 *  33       — PJ / CNPJ           tag PJF: linha 34 / tag PJC: linhas 35–44
 *  45       — SALDO PJ
 *  47       — INVESTIMENTOS       tag: IA (só linha 48)
 *  48       — Aportado no mês
 *  49       — Rendimento do mês
 *  51       — PATRIMÔNIO          tag: PAT
 *  52–54    — bens
 *  55       — TOTAL PATRIMÔNIO
 *  57       — SALDO DO MÊS
 *  59       — POSIÇÃO FINANCEIRA  tag: POS
 *  60–64    — itens (ITEMS_POS_FINANCEIRA)
 *  65       — TOTAL ATIVOS FINANCEIROS
 *  70       — LOG (título)        → LOG_ROW - 2
 *  71       — LOG (cabeçalhos)    → LOG_ROW - 1
 *  72+      — LOG (dados)         → LOG_ROW = 72
 */
function montarAbaMensal(sheet, mesNome, ano) {
  sheet.setConditionalFormatRules([]);
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 20);  // E — tag de seção (invisível)

  // ── Título ─────────────────────────────────────────────────────────────────
  sheet.setRowHeight(1, 42);
  sheet.getRange(1, 1, 1, 4).merge()
    .setValue(`${mesNome} / ${ano}`)
    .setBackground(COR.titulo).setFontColor(COR.tituloFonte)
    .setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // ── ENTRADAS ───────────────────────────────────────────────────────────────
  cabecalhoSecao(sheet, 3, 'ENTRADAS', COR.secao, COR.secaoFonte, ['', '', 'Real', '']);
  CAT_ENTRADA.forEach((cat, i) => {
    const row = 4 + i;
    linhaItem(sheet, row, cat, TAG.entrada, null, sumifCategoria(row), null);
  });
  linhaTotalSecao(sheet, 9, 'TOTAL ENTRADAS', TAG.entrada, null, false);

  // ── GASTOS FIXOS ───────────────────────────────────────────────────────────
  montarSecaoGastos(sheet, 11, 'GASTOS FIXOS', 12, CAT_FIXO, TAG.fixo, 20, 'TOTAL FIXOS');

  // ── GASTOS VARIÁVEIS ───────────────────────────────────────────────────────
  montarSecaoGastos(sheet, 22, 'GASTOS VARIÁVEIS', 23, CAT_VARIAVEL, TAG.variavel, 31, 'TOTAL VARIÁVEIS');

  // CF cobre diff de fixos e variáveis (range amplo acomoda linhas inseridas)
  formatacaoDiferenca(sheet, 'D12:D100');

  // ── PJ / CNPJ ──────────────────────────────────────────────────────────────
  cabecalhoSecao(sheet, 33, 'PJ / CNPJ', COR.pj, COR.pjFonte, ['', '', 'Real', '']);

  linhaItem(sheet, 34, 'Faturamento PJ', TAG.pjFat, null, sumifCategoria(34), null);

  CAT_PJ_CUSTO.forEach((cat, i) => {
    const row = 35 + i;
    linhaItem(sheet, row, cat, TAG.pjCusto, null, sumifCategoria(row), null);
  });

  // linha 45 — saldo PJ = faturamento − todos os custos
  sheet.getRange(45, 1, 1, 4).setBackground(COR.pjTotal);
  sheet.getRange(45, 1).setValue('SALDO PJ').setFontWeight('bold');
  sheet.getRange(45, 3)
    .setFormula(`=SUMIF($E:$E,"${TAG.pjFat}",$C:$C)-SUMIF($E:$E,"${TAG.pjCusto}",$C:$C)`)
    .setFontWeight('bold').setNumberFormat(FMT_BRL);
  formatacaoDiferenca(sheet, 'C45:C45');

  // ── INVESTIMENTOS ──────────────────────────────────────────────────────────
  cabecalhoSecao(sheet, 47, 'INVESTIMENTOS', COR.secao, COR.secaoFonte, ['', '', 'Valor', '']);

  // linha 48 — aportado: do log, tag IA (subtraído no Saldo do Mês)
  linhaItem(sheet, 48, 'Aportado no mês', TAG.invAporte, null,
    `=SUMIF($C$${LOG_ROW}:$C,"Aporte",$D$${LOG_ROW}:$D)`, null);

  // linha 49 — rendimento: entrada manual, verde=ganho / vermelho=perda
  sheet.getRange(49, 1).setValue('Rendimento do mês');
  sheet.getRange(49, 3).setNumberFormat(FMT_BRL);
  setTag(sheet, 49, '');
  formatacaoDiferenca(sheet, 'C49:C49');

  // ── PATRIMÔNIO ─────────────────────────────────────────────────────────────
  cabecalhoSecao(sheet, 51, 'PATRIMÔNIO', COR.secao, COR.secaoFonte, ['', '', 'Valor atual', '']);

  ['Apartamento', 'Lote', 'Carro'].forEach((item, i) => {
    const row = 52 + i;
    linhaItem(sheet, row, item, TAG.pat, null, null, null);
    sheet.getRange(row, 3).setNumberFormat(FMT_BRL);
  });

  sheet.getRange(55, 1, 1, 4).setBackground(COR.total);
  sheet.getRange(55, 1).setValue('TOTAL PATRIMÔNIO').setFontWeight('bold');
  sheet.getRange(55, 3)
    .setFormula(`=SUMIF($E:$E,"${TAG.pat}",$C:$C)`)
    .setFontWeight('bold').setNumberFormat(FMT_BRL);

  // ── SALDO DO MÊS ───────────────────────────────────────────────────────────
  sheet.setRowHeight(57, 40);
  sheet.getRange(57, 1, 1, 4).setBackground(COR.saldo);
  sheet.getRange(57, 1).setValue('SALDO DO MÊS')
    .setFontColor(COR.saldoFonte).setFontWeight('bold').setFontSize(12);
  sheet.getRange(57, 3)
    .setFormula(
      `=SUMIF($E:$E,"${TAG.entrada}",$C:$C)` +
      `-SUMIF($E:$E,"${TAG.fixo}",$C:$C)` +
      `-SUMIF($E:$E,"${TAG.variavel}",$C:$C)` +
      `-SUMIF($E:$E,"${TAG.invAporte}",$C:$C)`
    )
    .setFontColor(COR.saldoFonte).setFontWeight('bold').setFontSize(12)
    .setNumberFormat(FMT_BRL);
  formatacaoDiferenca(sheet, 'C57:C57');

  // ── POSIÇÃO FINANCEIRA ─────────────────────────────────────────────────────
  cabecalhoSecao(sheet, 59, 'POSIÇÃO FINANCEIRA', COR.secao, COR.secaoFonte, ['', '', 'Saldo atual', '']);

  ITEMS_POS_FINANCEIRA.forEach((item, i) => {
    const row = 60 + i;
    linhaItem(sheet, row, item, TAG.posFinanceira, null, null, null);
    sheet.getRange(row, 3).setNumberFormat(FMT_BRL);
  });

  const posEnd = 59 + ITEMS_POS_FINANCEIRA.length; // row 64
  sheet.getRange(posEnd + 1, 1, 1, 4).setBackground(COR.total);
  sheet.getRange(posEnd + 1, 1).setValue('TOTAL ATIVOS FINANCEIROS').setFontWeight('bold');
  sheet.getRange(posEnd + 1, 3)
    .setFormula(`=SUMIF($E:$E,"${TAG.posFinanceira}",$C:$C)`)
    .setFontWeight('bold').setNumberFormat(FMT_BRL);

  // ── LOG DE TRANSAÇÕES ──────────────────────────────────────────────────────
  sheet.setRowHeight(LOG_ROW - 2, 32);
  sheet.getRange(LOG_ROW - 2, 1, 1, 4).merge()
    .setValue('LOG DE TRANSAÇÕES')
    .setBackground(COR.titulo).setFontColor(COR.tituloFonte)
    .setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.setRowHeight(LOG_ROW - 1, 28);
  ['Data', 'Descrição', 'Categoria', 'Valor'].forEach((h, i) => {
    sheet.getRange(LOG_ROW - 1, i + 1)
      .setValue(h)
      .setBackground(COR.logHeader).setFontColor(COR.logFonte)
      .setFontWeight('bold').setHorizontalAlignment('center');
  });

  sheet.getRange(`A${LOG_ROW}:A2000`).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(`D${LOG_ROW}:D2000`).setNumberFormat(FMT_BRL);

  sheet.getRange(`C${LOG_ROW}:C2000`).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(CATEGORIAS, true)
      .setAllowInvalid(false)
      .build()
  );

  // Data picker na coluna A do log
  sheet.getRange(`A${LOG_ROW}:A2000`).setDataValidation(
    SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build()
  );

  aplicarCinzaFormulas(sheet);
  aplicarProtecao(sheet);

  sheet.setFrozenRows(1);
  sheet.getRange('A1:D57').setVerticalAlignment('middle');
}

// ─── DASHBOARD ────────────────────────────────────────────────────────────────

function criarDashboard(ss) {
  let sheet = ss.getSheetByName('Dashboard');

  if (sheet) {
    // Arquiva o dashboard existente se for de outro ano
    const matchAno = sheet.getRange(1, 1).getValue().toString().match(/\d{4}/);
    const anoExistente = matchAno ? parseInt(matchAno[0], 10) : null;
    if (anoExistente && anoExistente !== ANO) {
      sheet.setName(`Dashboard ${anoExistente}`);
      sheet = ss.insertSheet('Dashboard');
    } else {
      limparAba(sheet);
    }
  } else {
    sheet = ss.insertSheet('Dashboard');
  }

  [80, 140, 130, 140, 120, 140, 20, 150, 150, 140, 20, 170].forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  sheet.setRowHeight(1, 48);
  sheet.getRange(1, 1, 1, 12).merge()
    .setValue(`DASHBOARD — ${ANO}`)
    .setBackground(COR.titulo).setFontColor(COR.tituloFonte)
    .setFontWeight('bold').setFontSize(14)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.setRowHeight(2, 30);
  ['Mês', 'Entradas PF', 'Gastos Fixos', 'Gastos Variáveis', 'Aportes', 'Saldo PF',
   '', 'Faturamento PJ', 'Custos PJ', 'Saldo PJ', '', 'Ativos Financeiros'].forEach((h, i) => {
    sheet.getRange(2, i + 1)
      .setValue(h)
      .setBackground(COR.secao).setFontColor(COR.secaoFonte)
      .setFontWeight('bold').setHorizontalAlignment('center');
  });

  MESES.forEach(({ abrev }, idx) => {
    const row = idx + 3;
    const aba = `${abrev}/${ANO}`;
    const s   = (tag, col) => `SUMIF('${aba}'!E:E,"${tag}",'${aba}'!${col}:${col})`;

    sheet.getRange(row, 1).setValue(abrev);
    sheet.getRange(row, 2).setFormula(`=${s(TAG.entrada,'C')}`);
    sheet.getRange(row, 3).setFormula(`=${s(TAG.fixo,'C')}`);
    sheet.getRange(row, 4).setFormula(`=${s(TAG.variavel,'C')}`);
    sheet.getRange(row, 5).setFormula(`=${s(TAG.invAporte,'C')}`);
    sheet.getRange(row, 6).setFormula(
      `=${s(TAG.entrada,'C')}-${s(TAG.fixo,'C')}-${s(TAG.variavel,'C')}-${s(TAG.invAporte,'C')}`
    );
    sheet.getRange(row, 8).setFormula(`=${s(TAG.pjFat,'C')}`);
    sheet.getRange(row, 9).setFormula(`=${s(TAG.pjCusto,'C')}`);
    sheet.getRange(row, 10).setFormula(`=${s(TAG.pjFat,'C')}-${s(TAG.pjCusto,'C')}`);
    sheet.getRange(row, 12).setFormula(`=${s(TAG.posFinanceira,'C')}`);

    DASH_COLS.forEach(col => sheet.getRange(row, col).setNumberFormat(FMT_BRL));
    sheet.getRange(row, 1, 1, 12).setBackground(idx % 2 === 0 ? '#f7f9fc' : '#ffffff');
  });

  const totalRow = 3 + MESES.length;
  sheet.getRange(totalRow, 1).setValue('TOTAL').setFontWeight('bold');
  DASH_COLS.forEach(col => {
    const letra = colLetter(col);
    sheet.getRange(totalRow, col)
      .setFormula(`=SUM(${letra}3:${letra}${totalRow - 1})`)
      .setNumberFormat(FMT_BRL).setFontWeight('bold');
  });
  sheet.getRange(totalRow, 1, 1, 12).setBackground('#e8ecf0').setFontWeight('bold');

  formatacaoDiferenca(sheet, `F3:F${totalRow}`);
  formatacaoDiferenca(sheet, `J3:J${totalRow}`);
  sheet.setFrozenRows(2);
}

// ─── HELPERS ──────────────────────────────────────────────────────────────────

// Monta uma seção de gastos com budget/real/diferença (Fixos ou Variáveis)
function montarSecaoGastos(sheet, cabRow, titulo, startRow, itens, tag, totalRow, labelTotal) {
  cabecalhoSecao(sheet, cabRow, titulo, COR.secao, COR.secaoFonte, ['', 'Budget', 'Real', 'Diferença']);
  itens.forEach((cat, i) => {
    const row = startRow + i;
    linhaItem(sheet, row, cat, tag, null, sumifCategoria(row), `=IF(B${row}="","",B${row}-C${row})`);
    sheet.getRange(row, 2).setNumberFormat(FMT_BRL);
  });
  linhaTotalSecao(sheet, totalRow, labelTotal, tag,
    `=SUMIF($E:$E,"${tag}",$B:$B)-SUMIF($E:$E,"${tag}",$C:$C)`, true);
}

function linhaItem(sheet, row, label, tag, budgetFormula, realFormula, diffFormula) {
  sheet.getRange(row, 1).setValue(label);
  setTag(sheet, row, tag);
  if (budgetFormula) sheet.getRange(row, 2).setFormula(budgetFormula).setNumberFormat(FMT_BRL);
  if (realFormula)   sheet.getRange(row, 3).setFormula(realFormula).setNumberFormat(FMT_BRL);
  if (diffFormula)   sheet.getRange(row, 4).setFormula(diffFormula).setNumberFormat(FMT_BRL);
}

// Total via SUMIF na tag — robusto a linhas inseridas/removidas
function linhaTotalSecao(sheet, row, label, tag, diffFormula, showBudget) {
  sheet.getRange(row, 1, 1, 4).setBackground(COR.total);
  sheet.getRange(row, 1).setValue(label).setFontWeight('bold');
  if (showBudget) {
    sheet.getRange(row, 2)
      .setFormula(`=SUMIF($E:$E,"${tag}",$B:$B)`)
      .setFontWeight('bold').setNumberFormat(FMT_BRL);
  }
  sheet.getRange(row, 3)
    .setFormula(`=SUMIF($E:$E,"${tag}",$C:$C)`)
    .setFontWeight('bold').setNumberFormat(FMT_BRL);
  if (diffFormula) {
    sheet.getRange(row, 4).setFormula(diffFormula).setFontWeight('bold').setNumberFormat(FMT_BRL);
  }
}

// Retorna a fórmula SUMIF que busca a categoria do log pelo label da linha
function sumifCategoria(row) {
  return `=SUMIF($C$${LOG_ROW}:$C,A${row},$D$${LOG_ROW}:$D)`;
}

// Tag de seção na coluna E — texto branco sobre fundo branco (invisível ao usuário)
function setTag(sheet, row, tag) {
  sheet.getRange(row, 5).setValue(tag).setFontColor('#ffffff').setBackground('#ffffff');
}

function cabecalhoSecao(sheet, row, titulo, bg, fontColor, colLabels) {
  sheet.setRowHeight(row, 28);
  colLabels.forEach((label, i) => {
    const cell = sheet.getRange(row, i + 1);
    cell.setBackground(bg).setFontColor(fontColor).setFontWeight('bold');
    if (label) cell.setValue(label).setHorizontalAlignment('center');
  });
  sheet.getRange(row, 1).setValue(titulo).setHorizontalAlignment('left');
}

// Formatação condicional: positivo=verde, negativo=vermelho
// Chamado após setConditionalFormatRules([]) em montarAbaMensal — não acumula
function formatacaoDiferenca(sheet, rangeStr) {
  const range  = sheet.getRange(rangeStr);
  const regras = sheet.getConditionalFormatRules();
  regras.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground(COR.verdeClaro).setFontColor(COR.verdeFonte)
      .setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground(COR.vermClaro).setFontColor(COR.vermFonte)
      .setRanges([range]).build()
  );
  sheet.setConditionalFormatRules(regras);
}

// Obtém ou cria uma aba pelo nome, limpando se já existir
function getOrCreateSheet(ss, nome) {
  const existing = ss.getSheetByName(nome);
  if (!existing) return ss.insertSheet(nome);
  limparAba(existing);
  return existing;
}

// Limpa aba completamente incluindo formatação condicional e proteções
function limparAba(sheet) {
  sheet.clearContents();
  sheet.clearFormats();
  sheet.clearNotes();
  sheet.setConditionalFormatRules([]);
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
}

// Fundo cinza nas células com fórmulas automáticas — indica "não editar"
function aplicarCinzaFormulas(sheet) {
  const g           = COR.protegido;
  const entradasEnd = 3  + CAT_ENTRADA.length;   // 8
  const fixosEnd    = 11 + CAT_FIXO.length;      // 19
  const varEnd      = 22 + CAT_VARIAVEL.length;  // 30
  const pjCustoEnd  = 34 + CAT_PJ_CUSTO.length;  // 44

  const posEnd = 59 + ITEMS_POS_FINANCEIRA.length; // 64

  // Labels (col A) — alimentam SUMIF; renomear quebraria os cálculos
  [
    `A4:A${entradasEnd}`,
    `A12:A${fixosEnd}`,
    `A23:A${varEnd}`,
    `A34:A${pjCustoEnd}`,
    'A48', 'A49', 'A52:A54',
    `A60:A${posEnd}`,
  ].forEach(r => sheet.getRange(r).setBackground(g));

  // Fórmulas SUMIF (col C) — calculadas automaticamente a partir do log
  [
    `C4:C${entradasEnd}`,
    `C12:C${fixosEnd}`,
    `C23:C${varEnd}`,
    `C34:C${pjCustoEnd}`,
    'C48',
  ].forEach(r => sheet.getRange(r).setBackground(g));

  // Fórmulas de diferença (col D) — CF sobrepõe com verde/vermelho quando há valor;
  // cinza fica visível quando a célula está vazia ou com saldo zero.
  [
    `D12:D${fixosEnd}`,
    `D23:D${varEnd}`,
  ].forEach(r => sheet.getRange(r).setBackground(g));
}

// Proteção da aba: avisa ao editar células de fórmula
// Áreas editáveis (sem aviso): budget B, rendimento C49, patrimônio C52:C54, log
function aplicarProtecao(sheet) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());

  const fixosEnd = 11 + CAT_FIXO.length;
  const varEnd   = 22 + CAT_VARIAVEL.length;

  const protection = sheet.protect()
    .setDescription('Esta célula contém uma fórmula automática. Edite apenas as áreas em branco e o log.');
  protection.setWarningOnly(true);
  const posEnd = 59 + ITEMS_POS_FINANCEIRA.length; // 64

  protection.setUnprotectedRanges([
    sheet.getRange(`B12:B${fixosEnd}`),         // Budget Gastos Fixos
    sheet.getRange(`B23:B${varEnd}`),            // Budget Gastos Variáveis
    sheet.getRange('C49'),                       // Rendimento do mês
    sheet.getRange('C52:C54'),                   // Valores de Patrimônio
    sheet.getRange(`C60:C${posEnd}`),            // Posição Financeira
    sheet.getRange(`A${LOG_ROW}:D2000`),         // Log de transações
  ]);
}

// Legenda lateral (colunas G–H) — instruções de preenchimento
function adicionarLegenda(sheet) {
  const COL = 7; // coluna G

  const blocos = [
    { texto: 'COMO PREENCHER',                                         cabecalho: true },
    { texto: 'LOG DE TRANSAÇÕES',                                        secao: true },
    { col1: 'Onde:',        col2: `Linhas abaixo da linha ${LOG_ROW - 1} (LOG)` },
    { col1: 'Colunas:',     col2: 'Data | Descrição | Categoria | Valor' },
    { col1: 'Categoria:',   col2: 'determina em qual seção o valor aparece no resumo' },
    { texto: 'BUDGET',                                                  secao: true },
    { col1: 'Onde:',        col2: 'Coluna B — Gastos Fixos e Variáveis' },
    { col1: 'Copiar:',      col2: 'Financeiro > Copiar budget do mês anterior' },
    { texto: 'VALORES MANUAIS',                                         secao: true },
    { col1: 'Rendimento:',  col2: 'Coluna C, linha "Rendimento do mês" (linha 49)' },
    { col1: 'Patrimônio:',  col2: 'Coluna C, linhas Apartamento / Lote / Carro' },
    { texto: 'CELULAS EM CINZA',                                        secao: true },
    { col1: '',             col2: 'Contêm fórmulas — não edite.' },
    { col1: '',             col2: 'Um aviso aparece se você tentar editar.' },
  ];

  blocos.forEach(({ texto, col1, col2, cabecalho, secao }, i) => {
    const row = 3 + i;

    if (cabecalho) {
      sheet.setRowHeight(row, 30);
      sheet.getRange(row, COL, 1, 2).merge()
        .setValue(texto)
        .setBackground(COR.secao).setFontColor(COR.secaoFonte)
        .setFontWeight('bold').setFontSize(10)
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
    } else if (secao) {
      sheet.getRange(row, COL, 1, 2).merge()
        .setValue(texto)
        .setBackground(COR.total).setFontColor('#2c3e50')
        .setFontWeight('bold').setFontSize(9)
        .setVerticalAlignment('middle');
      sheet.setRowHeight(row, 22);
    } else {
      sheet.getRange(row, COL)
        .setValue(col1)
        .setBackground('#f8f9fa').setFontColor('#555555')
        .setFontSize(8).setFontWeight('bold');
      sheet.getRange(row, COL + 1)
        .setValue(col2)
        .setBackground('#f8f9fa').setFontColor('#333333')
        .setFontSize(8).setWrap(true);
      sheet.setRowHeight(row, 20);
    }
  });
}

// Converte número de coluna para letra: 1→A, 26→Z, 27→AA, ...
function colLetter(n) {
  let result = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    result   = String.fromCharCode(65 + rem) + result;
    n        = Math.floor((n - 1) / 26);
  }
  return result;
}
