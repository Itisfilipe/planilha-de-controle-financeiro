/**
 * CONTROLE FINANCEIRO — Google Apps Script
 * Versão: 1.0.0
 *
 * Planilha de controle financeiro pessoal no Google Sheets.
 *
 * COMO USAR:
 * 1. Crie um Google Sheets novo
 * 2. Extensões > Apps Script
 * 3. Cole este código e salve (Ctrl+S)
 * 4. Recarregue a planilha
 * 5. Menu "Financeiro" → "Criar mês atual"
 *
 * PREENCHIMENTO:
 * No LOG (parte de baixo da aba), preencha:
 *   Data | Descrição | Categoria | Valor | Parcela? (Sim/Não)
 * Os totais por categoria atualizam automaticamente.
 */

// ─── CONFIGURAÇÃO ─────────────────────────────────────────────────────────────

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

// ─── CATEGORIAS ───────────────────────────────────────────────────────────────
// Edite estes arrays para personalizar. Depois use "Atualizar categorias".
// CAT_FIXO = gastos recorrentes (moradia, assinaturas, etc.)
// CAT_VARIAVEL = gastos que mudam todo mês (alimentação, lazer, etc.)

const CAT_ENTRADA = [
  'Salário',
  'Freelance',
  'Outros entrada',
];

const CAT_FIXO = [
  'Moradia / Aluguel',
  'Condomínio / IPTU',
  'Energia',
  'Água',
  'Internet',
  'Celular',
  'Assinaturas',
  'Plano de saúde',
  'Seguro Carro',
  'Seguro Vida',
  'Mesada Carol',
  'Academia',
];

const CAT_VARIAVEL = [
  'Cartão de crédito',
  'Mercado / Feira',
  'Refeições fora',
  'Transporte / Combustível',
  'Farmácia',
  'Educação / Cursos',
  'Lazer / Entretenimento',
  'Vestuário',
  'Trabalho',
  'Presentes',
  'Outros',
];

const CATEGORIAS = [...CAT_ENTRADA, ...CAT_FIXO, ...CAT_VARIAVEL];

// Tags internas (coluna E, invisível ao usuário)
const TAG = { entrada: 'E', fixo: 'F', variavel: 'V' };

// LOG_ROW calculado a partir das categorias
const LOG_ROW = 3
  + (1 + CAT_ENTRADA.length + 1) + 1     // ENTRADAS
  + (1 + CAT_FIXO.length + 1) + 1        // GASTOS FIXOS
  + (1 + CAT_VARIAVEL.length + 1) + 1    // GASTOS VARIÁVEIS
  + 1                                     // SALDO DO MÊS
  + 4;                                    // gap(2) + título log(1) + cabeçalho(1)

// ─── CORES E FORMATOS ─────────────────────────────────────────────────────────

const COR = {
  titulo:     '#1a1a2e',
  tituloFonte:'#ffffff',
  secao:      '#2c3e50',
  secaoFonte: '#ffffff',
  total:      '#dde8f0',
  saldo:      '#1a1a2e',
  saldoFonte: '#ffffff',
  logHeader:  '#34495e',
  logFonte:   '#ffffff',
  verdeClaro: '#c8e6c9',
  verdeFonte: '#1b5e20',
  vermClaro:  '#ffcdd2',
  vermFonte:  '#b71c1c',
  protegido:  '#eeeeee',
};

const FMT_BRL = 'R$ #,##0.00';

// ─── MENU ─────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Financeiro')
    .addItem('Novo mês',                           'criarMesAtual')
    .addSeparator()
    .addItem('Ver resumo',                          'resumoMes')
    .addItem('Dívidas',                             'criarAbaDividas')
    .addItem('Como usar',                           'criarAbaComoUsar')
    .addSeparator()
    .addItem('Atualizar categorias',                'atualizarDropdowns')
    .addToUi();
}

// ─── CRIAR MÊS ATUAL ─────────────────────────────────────────────────────────

function criarMesAtual() {
  const ui   = SpreadsheetApp.getUi();
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const hoje = new Date();
  const mes  = MESES[hoje.getMonth()];
  const ano  = hoje.getFullYear();
  const nome = `${mes.abrev}/${ano}`;

  if (ss.getSheetByName(nome)) {
    ui.alert(`A aba "${nome}" já existe. Use-a diretamente ou exclua-a manualmente antes de recriar.`);
    ss.setActiveSheet(ss.getSheetByName(nome));
    return;
  }

  const ok = ui.alert(
    'Criar Mês Atual',
    `Criar a aba "${nome}" com resumo e log de transações?`,
    ui.ButtonSet.YES_NO
  );
  if (ok !== ui.Button.YES) return;

  try { ss.setSpreadsheetLocale('pt_BR'); } catch (e) {}
  try { ss.setSpreadsheetTimeZone('America/Sao_Paulo'); } catch (e) {}

  const sheet = getOrCreateSheet(ss, nome);
  montarAba(sheet, mes.nome, ano);

  // Cria aba "Como usar" se não existir
  if (!ss.getSheetByName('Como usar')) criarAbaComoUsar();

  ss.setActiveSheet(sheet);

  // Remove aba padrão se possível
  ['Planilha1', 'Sheet1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 1) try { ss.deleteSheet(s); } catch (e) {}
  });

  SpreadsheetApp.flush();
  ui.alert(`Aba "${nome}" criada! Anote suas transações no LOG.`);
}

// ─── RESUMO DO MÊS ───────────────────────────────────────────────────────────

function resumoMes() {
  const ui    = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const nome  = sheet.getName();

  if (!/^[A-Za-z]{3}\/\d{4}$/.test(nome)) {
    ui.alert('Abra uma aba mensal (ex: Jan/2026) antes de usar esta função.');
    return;
  }

  const lastRow = sheet.getLastRow();
  const tags    = sheet.getRange(1, 5, lastRow).getValues();
  const reais   = sheet.getRange(1, 3, lastRow).getValues();

  const sumif = tag => tags.reduce((s, [t], i) => t === tag ? s + (Number(reais[i][0]) || 0) : s, 0);
  const fmt   = v => 'R$ ' + v.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  const entradas  = sumif(TAG.entrada);
  const fixos     = sumif(TAG.fixo);
  const variaveis = sumif(TAG.variavel);
  const saldo     = entradas - fixos - variaveis;

  ui.alert(
    `Resumo — ${nome}`,
    `Entradas:          ${fmt(entradas)}\n` +
    `Gastos Fixos:      ${fmt(fixos)}\n` +
    `Gastos Variáveis:  ${fmt(variaveis)}\n` +
    `─────────────────────────\n` +
    `Saldo:             ${fmt(saldo)}`,
    ui.ButtonSet.OK
  );
}

// ─── ATUALIZAR DROPDOWNS ──────────────────────────────────────────────────────

function atualizarDropdowns() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const ok = ui.alert(
    'Atualizar Categorias',
    'Atualiza dropdowns e resumo (entradas/gastos) em todas as abas mensais.\n\nOs dados do log serão preservados. Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (ok !== ui.Button.YES) return;

  let count = 0;
  ss.getSheets().forEach(sheet => {
    if (/^[A-Za-z]{3}\/\d{4}$/.test(sheet.getName())) {
      reconstruirResumo(sheet);
      count++;
    }
  });

  ui.alert(`Categorias e dropdowns atualizados em ${count} aba(s).`);
}

// ─── ABA COMO USAR ────────────────────────────────────────────────────────────

function criarAbaComoUsar() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Como usar');
  if (sheet) {
    ss.setActiveSheet(sheet);
    return;
  }

  sheet = ss.insertSheet('Como usar');
  sheet.setColumnWidth(1, 800);
  sheet.setTabColor('#2196F3');

  const linhas = [
    ['COMO USAR — CONTROLE FINANCEIRO'],
    [''],
    ['COMO FUNCIONA'],
    ['Cada mês tem sua própria aba (ex: Abr/2026). A aba é dividida em duas partes:'],
    ['  • RESUMO (parte de cima) — totais automáticos por categoria, calculados a partir do log'],
    ['  • LOG (parte de baixo) — onde você anota cada transação do mês'],
    [''],
    ['O resumo tem três seções: ENTRADAS (dinheiro que entra), GASTOS FIXOS e GASTOS VARIÁVEIS.'],
    ['A categoria de cada transação no log determina em qual seção ela aparece.'],
    ['O SALDO DO MÊS = Entradas − Gastos Fixos − Gastos Variáveis, calculado automaticamente.'],
    [''],
    ['PREENCHIMENTO DIÁRIO'],
    ['No LOG, preencha uma linha para cada transação:'],
    ['  • Coluna A — Data (seletor de data disponível)'],
    ['  • Coluna B — Descrição livre (ex: "Supermercado Extra", "Salário março")'],
    ['  • Coluna C — Categoria (dropdown: Salário, Alimentação, Transporte, etc.)'],
    ['  • Coluna D — Valor (sempre POSITIVO — o sistema sabe se é entrada ou saída pela categoria)'],
    ['  • Coluna E — Parcela? (Sim/Não) — marque "Sim" se for pagamento de parcela de dívida'],
    [''],
    ['IMPORTANTE: o valor é sempre positivo. Se você gastou R$ 50 no mercado, coloque 50 (não -50).'],
    ['A categoria "Alimentação" está em GASTOS VARIÁVEIS, então o sistema já sabe que é um gasto.'],
    [''],
    ['O QUE NÃO FUNCIONA AUTOMATICAMENTE'],
    ['  • A aba Dívidas NÃO atualiza sozinha a partir do log — você precisa atualizar'],
    ['    "Parcelas pagas" manualmente na aba Dívidas a cada mês.'],
    ['  • Não há integração entre meses — cada aba é independente.'],
    ['  • Não há dashboard ou gráficos (ainda) — use o "Resumo do mês" no menu.'],
    [''],
    ['MENU FINANCEIRO'],
    ['  • Novo mês — cria a aba do mês atual (se já existir, navega até ela)'],
    ['  • Ver resumo — popup com totais rápidos (entradas, gastos fixos/variáveis, saldo)'],
    ['  • Dívidas — cria ou atualiza a aba de parcelas e financiamentos'],
    ['  • Como usar — abre esta aba'],
    ['  • Atualizar categorias — recria o resumo se você mudou os arrays no código'],
    ['    (dados do log são preservados e migrados automaticamente)'],
    [''],
    ['PERSONALIZAR CATEGORIAS'],
    ['  As categorias são definidas no código (Extensões > Apps Script):'],
    ['    CAT_ENTRADA  = [\'Salário\', \'Freelance\', \'Outros entrada\']'],
    ['    CAT_FIXO     = [\'Moradia\', \'Assinaturas\', \'Cartão de crédito\', ...]'],
    ['    CAT_VARIAVEL = [\'Alimentação\', \'Transporte\', \'Lazer\', ...]'],
    ['  Para adicionar ou remover: edite o array, salve, e use Financeiro > Atualizar categorias.'],
    ['  O resumo é reconstruído e os dados do log são preservados.'],
    ['  Se o layout mudou (mais ou menos categorias), o log é migrado automaticamente.'],
    [''],
    ['DÍVIDAS E PARCELAS'],
    ['  A aba Dívidas é INDEPENDENTE das abas mensais — serve para acompanhar'],
    ['  parcelas e financiamentos ao longo do tempo.'],
    [''],
    ['  Para cada dívida, preencha:'],
    ['    Descrição (ex: "Geladeira Nubank") | Valor total | Parcelas | Início'],
    ['  E atualize todo mês:'],
    ['    Parcelas pagas — quantas parcelas você já pagou no total'],
    [''],
    ['  O sistema calcula automaticamente:'],
    ['    Valor mensal = Valor total ÷ Parcelas'],
    ['    Restantes = Parcelas − Parcelas pagas'],
    ['    Saldo devedor = Valor mensal × Restantes'],
    ['    Status = "Ativa" ou "Quitada" (quando Restantes ≤ 0)'],
    [''],
    ['  No log mensal, lance o pagamento da parcela como saída normal'],
    ['  (ex: categoria "Cartão de crédito") e marque "Parcela? = Sim".'],
    ['  Isso garante que o gasto aparece no saldo do mês E você sabe que é parcela.'],
    [''],
    ['CÉLULAS EM CINZA'],
    ['  Contêm fórmulas automáticas — NÃO edite.'],
    ['  Se tentar editar, um aviso aparecerá pedindo confirmação.'],
    ['  Áreas editáveis: log de transações e Parcelas pagas na aba Dívidas.'],
    [''],
    ['DICA: esta aba pode ser excluída sem afetar a planilha. Para recriá-la,'],
    ['use Financeiro > Como usar (abrir aba).'],
  ];

  sheet.getRange(1, 1, linhas.length, 1).setValues(linhas)
    .setFontFamily('Google Sans').setVerticalAlignment('middle');

  // Título
  sheet.getRange(1, 1).setFontSize(14).setFontWeight('bold')
    .setBackground(COR.titulo).setFontColor(COR.tituloFonte);
  sheet.setRowHeight(1, 42);

  // Seções
  [3, 12, 23, 29, 37, 46, 65].forEach(r => {
    sheet.getRange(r, 1).setFontSize(11).setFontWeight('bold')
      .setFontColor(COR.secao);
  });

  sheet.setFrozenRows(1);
  ss.setActiveSheet(sheet);
}

// ─── MONTAR ABA ───────────────────────────────────────────────────────────────

/*
 * Layout (linhas calculadas a partir dos arrays):
 *  1        — Título
 *  3        — ENTRADAS (header + itens + total)
 *            — GASTOS FIXOS (header + itens + total)
 *            — GASTOS VARIÁVEIS (header + itens + total)
 *            — SALDO DO MÊS
 *  LOG_ROW  — início do log de transações
 */
function montarAba(sheet, mesNome, ano) {
  sheet.setConditionalFormatRules([]);
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 80);

  // ── Título ─────────────────────────────────────────────────────────────────
  sheet.setRowHeight(1, 42);
  sheet.getRange(1, 1, 1, 4).merge()
    .setValue(`${mesNome} / ${ano}`)
    .setBackground(COR.titulo).setFontColor(COR.tituloFonte)
    .setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // ── Resumo + log (título, cabeçalhos, validações) ───────────────────────────
  reconstruirResumo(sheet);
  sheet.setFrozenRows(1);
}

// Reconstrói as seções de resumo (entradas, gastos fixos, gastos variáveis, saldo) sem tocar no log.
// Chamado por montarAba (aba nova) e atualizarDropdowns (aba existente).
function reconstruirResumo(sheet) {
  // Detecta posição antiga do log (pode diferir se categorias mudaram)
  const totalRows = sheet.getLastRow();
  let oldLogDataRow = 0;
  if (totalRows > 0) {
    const colA = sheet.getRange(1, 1, totalRows, 1).getValues();
    for (let i = 0; i < colA.length; i++) {
      if (colA[i][0] === 'LOG DE TRANSAÇÕES') { oldLogDataRow = i + 3; break; }
    }
  }

  // Se o log mudou de posição, migra os dados
  let savedLogData = null;
  if (oldLogDataRow > 0 && oldLogDataRow !== LOG_ROW && totalRows >= oldLogDataRow) {
    savedLogData = sheet.getRange(oldLogDataRow, 1, totalRows - oldLogDataRow + 1, 5).getValues();
    sheet.getRange(oldLogDataRow - 2, 1, totalRows - oldLogDataRow + 3, 5).clearContent().clearFormat();
  }

  // Limpa tudo entre título (row 1) e log dados (LOG_ROW), incluindo gap rows e log title/headers
  const clearEnd = Math.min(LOG_ROW - 1, totalRows);
  if (clearEnd >= 2) {
    sheet.getRange(2, 1, clearEnd - 1, 5).clearContent().clearFormat()
      .clearDataValidations().setBackground(null);
  }
  sheet.setConditionalFormatRules([]);

  // Posições calculadas
  const entHeader = 3;
  const entStart  = entHeader + 1;
  const entEnd    = entHeader + CAT_ENTRADA.length;
  const entTotal  = entEnd + 1;

  const fixHeader = entTotal + 2;
  const fixStart  = fixHeader + 1;
  const fixEnd    = fixHeader + CAT_FIXO.length;
  const fixTotal  = fixEnd + 1;

  const varHeader = fixTotal + 2;
  const varStart  = varHeader + 1;
  const varEnd    = varHeader + CAT_VARIAVEL.length;
  const varTotal  = varEnd + 1;

  const saldoRow  = varTotal + 2;

  // ── Seções: Entradas, Gastos Fixos, Gastos Variáveis ────────────────────
  const secoes = [
    { header: entHeader, start: entStart, cats: CAT_ENTRADA, tag: TAG.entrada, titulo: 'ENTRADAS',         totalLabel: 'TOTAL ENTRADAS'  },
    { header: fixHeader, start: fixStart, cats: CAT_FIXO,    tag: TAG.fixo,    titulo: 'GASTOS FIXOS',     totalLabel: 'TOTAL FIXOS'     },
    { header: varHeader, start: varStart, cats: CAT_VARIAVEL,tag: TAG.variavel,titulo: 'GASTOS VARIÁVEIS', totalLabel: 'TOTAL VARIÁVEIS' },
  ];

  secoes.forEach(({ header, start, cats, tag, titulo, totalLabel }) => {
    cabecalho(sheet, header, titulo, ['', '', 'Real', '']);
    cats.forEach((cat, i) => {
      const row = start + i;
      sheet.getRange(row, 1).setValue(cat);
      setTag(sheet, row, tag);
      sheet.getRange(row, 3)
        .setFormula(`=SUMIF($C$${LOG_ROW}:$C;A${row};$D$${LOG_ROW}:$D)`)
        .setNumberFormat(FMT_BRL);
    });
    linhaTotal(sheet, start + cats.length, totalLabel, tag);
  });

  // ── SALDO DO MÊS ──────────────────────────────────────────────────────────
  sheet.setRowHeight(saldoRow, 40);
  sheet.getRange(saldoRow, 1, 1, 4).setBackground(COR.saldo);
  sheet.getRange(saldoRow, 1).setValue('SALDO DO MÊS')
    .setFontColor(COR.saldoFonte).setFontWeight('bold').setFontSize(12);
  sheet.getRange(saldoRow, 3)
    .setFormula(
      `=SUMIF($E:$E;"${TAG.entrada}";$C:$C)` +
      `-SUMIF($E:$E;"${TAG.fixo}";$C:$C)` +
      `-SUMIF($E:$E;"${TAG.variavel}";$C:$C)`
    )
    .setFontColor(COR.saldoFonte).setFontWeight('bold').setFontSize(12)
    .setNumberFormat(FMT_BRL);
  formatacaoCondicional(sheet, `C${saldoRow}:C${saldoRow}`);

  // ── Formatação ─────────────────────────────────────────────────────────────
  const g = COR.protegido;
  [
    `A${entStart}:A${entEnd}`,
    `C${entStart}:C${entEnd}`,
    `A${fixStart}:A${fixEnd}`,
    `C${fixStart}:C${fixEnd}`,
    `A${varStart}:A${varEnd}`,
    `C${varStart}:C${varEnd}`,
  ].forEach(r => sheet.getRange(r).setBackground(g));

  // Proteção com aviso
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  const protection = sheet.protect()
    .setDescription('Células com fórmula — edite apenas o LOG.');
  protection.setWarningOnly(true);
  protection.setUnprotectedRanges([
    sheet.getRange(`A${LOG_ROW}:E2000`),
  ]);

  sheet.getRange(1, 1, saldoRow, 4).setVerticalAlignment('middle');

  // ── Reescreve título e cabeçalhos do log (podem ter sido apagados na migração)
  sheet.setRowHeight(LOG_ROW - 2, 32);
  sheet.getRange(LOG_ROW - 2, 1, 1, 5).merge()
    .setValue('LOG DE TRANSAÇÕES')
    .setBackground(COR.titulo).setFontColor(COR.tituloFonte)
    .setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.setRowHeight(LOG_ROW - 1, 28);
  ['Data', 'Descrição', 'Categoria', 'Valor', 'Parcela?'].forEach((h, i) => {
    sheet.getRange(LOG_ROW - 1, i + 1)
      .setValue(h)
      .setBackground(COR.logHeader).setFontColor(COR.logFonte)
      .setFontWeight('bold').setHorizontalAlignment('center');
  });

  // ── Formatos e validações do log ─────────────────────────────────────────
  sheet.getRange(`A${LOG_ROW}:A2000`).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(`D${LOG_ROW}:D2000`).setNumberFormat(FMT_BRL);

  sheet.getRange(`C${LOG_ROW}:C2000`).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(CATEGORIAS, true)
      .setAllowInvalid(false)
      .build()
  );

  sheet.getRange(`E${LOG_ROW}:E2000`).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Sim', 'Não'], true)
      .setAllowInvalid(true)
      .build()
  );

  sheet.getRange(`A${LOG_ROW}:A2000`).setDataValidation(
    SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build()
  );

  // Restaura dados do log migrados (se o log mudou de posição)
  if (savedLogData) {
    sheet.getRange(LOG_ROW, 1, savedLogData.length, savedLogData[0].length).setValues(savedLogData);
  }
}

// ─── ABA DÍVIDAS ──────────────────────────────────────────────────────────────

function criarAbaDividas() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const ok = ui.alert(
    'Dívidas e Parcelas',
    'Criar (ou atualizar) a aba "Dívidas" para acompanhar parcelas e financiamentos?\n\nDados existentes serão preservados.',
    ui.ButtonSet.YES_NO
  );
  if (ok !== ui.Button.YES) return;

  let sheet = ss.getSheetByName('Dívidas');
  const isNew = !sheet;
  if (isNew) {
    sheet = ss.insertSheet('Dívidas');
  } else {
    // Salva dados manuais (A:C e E:F) antes de limpar
    const lr = sheet.getLastRow();
    let savedData = null;
    if (lr >= 3) {
      const all = sheet.getRange(3, 1, lr - 2, 6).getValues(); // A-F
      savedData = all.filter(row => row[0] !== '' && row[0] !== 'TOTAIS');
    }
    // Limpa tudo exceto título (row 1)
    if (lr >= 2) {
      sheet.getRange(2, 1, lr - 1, sheet.getMaxColumns()).clearContent().clearFormat()
        .clearDataValidations().setBackground(null);
    }
    sheet.setConditionalFormatRules([]);
    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    // Restaura dados manuais
    if (savedData && savedData.length > 0) {
      sheet.getRange(3, 1, savedData.length, 6).setValues(savedData);
    }
  }

  // A=Descrição B=Valor total C=Parcelas D=Valor mensal E=Início
  // F=Parcelas pagas G=Restantes H=Saldo devedor I=Status
  const headers = ['Descrição', 'Valor total', 'Parcelas', 'Valor mensal',
    'Início', 'Parcelas pagas', 'Restantes', 'Saldo devedor', 'Status'];
  [220, 130, 80, 130, 100, 110, 100, 140, 80].forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // Título
  sheet.setRowHeight(1, 42);
  if (isNew) {
    sheet.getRange(1, 1, 1, 9).merge()
      .setValue('DÍVIDAS E PARCELAS')
      .setBackground(COR.titulo).setFontColor(COR.tituloFonte)
      .setFontWeight('bold').setFontSize(13)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  }

  // Cabeçalhos (row 2)
  sheet.setRowHeight(2, 28);
  headers.forEach((h, i) => {
    sheet.getRange(2, i + 1)
      .setValue(h)
      .setBackground(COR.secao).setFontColor(COR.secaoFonte)
      .setFontWeight('bold').setHorizontalAlignment('center');
  });

  // Formatos
  sheet.getRange('B3:B500').setNumberFormat(FMT_BRL);
  sheet.getRange('D3:D500').setNumberFormat(FMT_BRL);
  sheet.getRange('H3:H500').setNumberFormat(FMT_BRL);

  for (let r = 3; r <= 100; r++) {
    // D = Valor mensal = Valor total / Parcelas
    sheet.getRange(r, 4).setFormula(`=IF(AND(B${r}<>"";C${r}<>"");B${r}/C${r};"")`);
    // G = Restantes = Parcelas - Pagas
    sheet.getRange(r, 7).setFormula(`=IF(AND(C${r}<>"";F${r}<>"");C${r}-F${r};"")`);
    // H = Saldo devedor = Valor mensal * Restantes
    sheet.getRange(r, 8).setFormula(`=IF(AND(D${r}<>"";G${r}<>"");D${r}*G${r};"")`);
    // I = Status
    sheet.getRange(r, 9).setFormula(`=IF(A${r}="";"";IF(G${r}<=0;"Quitada";"Ativa"))`);
  }

  // Totais
  const totalRow = 102;
  sheet.getRange(totalRow, 1).setValue('TOTAIS').setFontWeight('bold');
  sheet.getRange(totalRow, 1, 1, 9).setBackground(COR.total);
  sheet.getRange(totalRow, 4).setFormula('=SUM(D3:D101)').setFontWeight('bold').setNumberFormat(FMT_BRL);
  sheet.getRange(totalRow, 6).setFormula('=SUM(F3:F101)').setFontWeight('bold');
  sheet.getRange(totalRow, 8).setFormula('=SUM(H3:H101)').setFontWeight('bold').setNumberFormat(FMT_BRL);

  // Cinza nas colunas com fórmula (D, G, H, I)
  ['D3:D101', 'G3:G101', 'H3:H101', 'I3:I101'].forEach(r =>
    sheet.getRange(r).setBackground(COR.protegido)
  );

  // Formatação condicional: Quitada = verde, Ativa = vermelho claro
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Quitada')
      .setBackground(COR.verdeClaro).setFontColor(COR.verdeFonte)
      .setRanges([sheet.getRange('I3:I101')]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Ativa')
      .setBackground(COR.vermClaro).setFontColor(COR.vermFonte)
      .setRanges([sheet.getRange('I3:I101')]).build(),
  ]);

  // Proteção — editável: A, B, C, E, F
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  const protection = sheet.protect()
    .setDescription('Colunas D, G, H e I contêm fórmulas.');
  protection.setWarningOnly(true);
  protection.setUnprotectedRanges([
    sheet.getRange('A3:C101'),
    sheet.getRange('E3:F101'),
  ]);

  sheet.setFrozenRows(2);
  ss.setActiveSheet(sheet);
  ui.alert(
    'Aba "Dívidas" pronta!\n\n' +
    '• Preencha: Descrição, Valor total, Parcelas, Início e Parcelas pagas.\n' +
    '• Atualize "Parcelas pagas" todo mês.\n' +
    '• Status muda para "Quitada" automaticamente quando Restantes ≤ 0.'
  );
}

// ─── HELPERS ──────────────────────────────────────────────────────────────────

function cabecalho(sheet, row, titulo, labels) {
  sheet.setRowHeight(row, 28);
  labels.forEach((label, i) => {
    const cell = sheet.getRange(row, i + 1);
    cell.setBackground(COR.secao).setFontColor(COR.secaoFonte).setFontWeight('bold');
    if (label) cell.setValue(label).setHorizontalAlignment('center');
  });
  sheet.getRange(row, 1).setValue(titulo).setHorizontalAlignment('left');
}

function linhaTotal(sheet, row, label, tag) {
  sheet.getRange(row, 1, 1, 4).setBackground(COR.total);
  sheet.getRange(row, 1).setValue(label).setFontWeight('bold');
  sheet.getRange(row, 3)
    .setFormula(`=SUMIF($E:$E;"${tag}";$C:$C)`)
    .setFontWeight('bold').setNumberFormat(FMT_BRL);
}

function setTag(sheet, row, tag) {
  sheet.getRange(row, 5).setValue(tag).setFontColor('#ffffff').setBackground('#ffffff');
}

function formatacaoCondicional(sheet, rangeStr) {
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

function getOrCreateSheet(ss, nome) {
  return ss.getSheetByName(nome) || ss.insertSheet(nome);
}
