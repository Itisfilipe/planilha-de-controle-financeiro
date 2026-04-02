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

const CAT_ENTRADA = [
  'Salário',
  'Freelance',
  'Outros entrada',
];

const CAT_SAIDA = [
  'Moradia',
  'Alimentação',
  'Transporte',
  'Trabalho',
  'Saúde',
  'Educação',
  'Lazer',
  'Assinaturas',
  'Cartão de crédito',
  'Vestuário',
  'Outros',
];

const CATEGORIAS = [...CAT_ENTRADA, ...CAT_SAIDA];

// Tags internas (coluna E, invisível ao usuário)
const TAG = { entrada: 'E', saida: 'S' };

// LOG_ROW calculado a partir das categorias
const LOG_ROW = 3
  + (1 + CAT_ENTRADA.length + 1) + 1   // ENTRADAS: header + items + total + gap
  + (1 + CAT_SAIDA.length + 1) + 1     // SAÍDAS: header + items + total + gap
  + 1                                   // SALDO DO MÊS
  + 4;                                  // gap(2) + título log(1) + cabeçalho(1)

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
    .addItem('Criar mês atual',                  'criarMesAtual')
    .addItem('Novo mês...',                       'criarNovoMes')
    .addSeparator()
    .addItem('Resumo do mês',                     'resumoMes')
    .addItem('Atualizar categorias',               'atualizarDropdowns')
    .addSeparator()
    .addItem('Criar / atualizar aba Dívidas',     'criarAbaDividas')
    .addSeparator()
    .addItem('Como usar (abrir aba)',              'criarAbaComoUsar')
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

// ─── NOVO MÊS ────────────────────────────────────────────────────────────────

function criarNovoMes() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const resposta = ui.prompt(
    'Novo Mês',
    'Digite no formato Abrev/Ano — exemplo: Jan/2026',
    ui.ButtonSet.OK_CANCEL
  );
  if (resposta.getSelectedButton() !== ui.Button.OK) return;

  const partes = resposta.getResponseText().trim().split('/');
  if (partes.length !== 2) { ui.alert('Formato inválido. Use: Jan/2026'); return; }

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

  const { abrev, nome: mesNome } = MESES[idx];
  const nomeAba = `${abrev}/${anoInput}`;

  if (ss.getSheetByName(nomeAba)) {
    ui.alert(`A aba "${nomeAba}" já existe. Use-a diretamente ou exclua-a manualmente antes de recriar.`);
    ss.setActiveSheet(ss.getSheetByName(nomeAba));
    return;
  }

  try { ss.setSpreadsheetLocale('pt_BR'); } catch (e) {}

  const sheet = getOrCreateSheet(ss, nomeAba);
  montarAba(sheet, mesNome, anoInput);
  ss.setActiveSheet(sheet);
  ui.alert(`Aba "${nomeAba}" criada!`);
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

  const entradas = sumif(TAG.entrada);
  const saidas   = sumif(TAG.saida);
  const saldo    = entradas - saidas;

  ui.alert(
    `Resumo — ${nome}`,
    `Entradas:   ${fmt(entradas)}\n` +
    `Saídas:     ${fmt(saidas)}\n` +
    `─────────────────────\n` +
    `Saldo:      ${fmt(saldo)}`,
    ui.ButtonSet.OK
  );
}

// ─── ATUALIZAR DROPDOWNS ──────────────────────────────────────────────────────

function atualizarDropdowns() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const ok = ui.alert(
    'Atualizar Categorias',
    'Atualiza dropdowns e resumo (entradas/saídas) em todas as abas mensais.\n\nOs dados do log serão preservados. Continuar?',
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
      reconstruirResumo(sheet);
      sheet.getRange(`C${LOG_ROW}:C2000`).setDataValidation(validacao);
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
    ['O resumo tem duas seções: ENTRADAS (dinheiro que entra) e SAÍDAS (dinheiro que sai).'],
    ['A categoria de cada transação no log determina em qual seção ela aparece.'],
    ['O SALDO DO MÊS = Total Entradas − Total Saídas, calculado automaticamente.'],
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
    ['A categoria "Alimentação" está em SAÍDAS, então o sistema já sabe que é um gasto.'],
    [''],
    ['O QUE NÃO FUNCIONA AUTOMATICAMENTE'],
    ['  • A aba Dívidas NÃO atualiza sozinha a partir do log — você precisa atualizar'],
    ['    "Parcelas pagas" manualmente na aba Dívidas a cada mês.'],
    ['  • Não há integração entre meses — cada aba é independente.'],
    ['  • Não há dashboard ou gráficos (ainda) — use o "Resumo do mês" no menu.'],
    [''],
    ['MENU FINANCEIRO'],
    ['  • Criar mês atual — cria a aba do mês corrente (não sobrescreve se já existir)'],
    ['  • Novo mês... — cria qualquer mês/ano (ex: Jan/2027, Dez/2025)'],
    ['  • Resumo do mês — popup com totais rápidos (entradas, saídas, saldo)'],
    ['  • Atualizar categorias — recria o resumo com as categorias atuais do script'],
    ['    (dados do log são preservados e migrados se o layout mudou)'],
    ['  • Criar / atualizar aba Dívidas — cria ou atualiza a aba de parcelas'],
    ['  • Como usar (abrir aba) — abre esta aba'],
    [''],
    ['PERSONALIZAR CATEGORIAS'],
    ['  As categorias são definidas no código (Extensões > Apps Script):'],
    ['    CAT_ENTRADA = [\'Salário\', \'Freelance\', \'Outros entrada\']'],
    ['    CAT_SAIDA = [\'Moradia\', \'Alimentação\', \'Transporte\', ...]'],
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
  [3, 12, 23, 29, 38, 46, 65].forEach(r => {
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
 *  3        — ENTRADAS (header)
 *  4–N      — itens entrada
 *  N+1      — TOTAL ENTRADAS
 *  N+3      — SAÍDAS (header)
 *  N+4–M    — itens saída
 *  M+1      — TOTAL SAÍDAS
 *  M+3      — SALDO DO MÊS
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

  // ── Resumo (entradas, saídas, saldo) ───────────────────────────────────────
  reconstruirResumo(sheet);

  // ── LOG DE TRANSAÇÕES ──────────────────────────────────────────────────────
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

  sheet.setFrozenRows(1);
}

// Reconstrói as seções de resumo (entradas, saídas, saldo) sem tocar no log.
// Chamado por montarAba (aba nova) e atualizarDropdowns (aba existente).
function reconstruirResumo(sheet) {
  // Detecta posição antiga do log (pode diferir se categorias mudaram)
  const totalRows = sheet.getLastRow();
  let oldLogDataRow = 0;
  const searchRows = Math.min(totalRows, LOG_ROW + 20);
  if (searchRows > 0) {
    const colA = sheet.getRange(1, 1, searchRows, 1).getValues();
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

  // Limpa área do resumo (entre título e log title bar)
  const clearEnd = Math.min(LOG_ROW - 4, totalRows);
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

  const saiHeader = entTotal + 2;
  const saiStart  = saiHeader + 1;
  const saiEnd    = saiHeader + CAT_SAIDA.length;
  const saiTotal  = saiEnd + 1;

  const saldoRow  = saiTotal + 2;

  // ── ENTRADAS ───────────────────────────────────────────────────────────────
  cabecalho(sheet, entHeader, 'ENTRADAS', ['', '', 'Real', '']);
  CAT_ENTRADA.forEach((cat, i) => {
    const row = entStart + i;
    sheet.getRange(row, 1).setValue(cat);
    setTag(sheet, row, TAG.entrada);
    sheet.getRange(row, 3)
      .setFormula(`=SUMIF($C$${LOG_ROW}:$C;A${row};$D$${LOG_ROW}:$D)`)
      .setNumberFormat(FMT_BRL);
  });
  linhaTotal(sheet, entTotal, 'TOTAL ENTRADAS', TAG.entrada);

  // ── SAÍDAS ─────────────────────────────────────────────────────────────────
  cabecalho(sheet, saiHeader, 'SAÍDAS', ['', '', 'Real', '']);
  CAT_SAIDA.forEach((cat, i) => {
    const row = saiStart + i;
    sheet.getRange(row, 1).setValue(cat);
    setTag(sheet, row, TAG.saida);
    sheet.getRange(row, 3)
      .setFormula(`=SUMIF($C$${LOG_ROW}:$C;A${row};$D$${LOG_ROW}:$D)`)
      .setNumberFormat(FMT_BRL);
  });
  linhaTotal(sheet, saiTotal, 'TOTAL SAÍDAS', TAG.saida);

  // ── SALDO DO MÊS ──────────────────────────────────────────────────────────
  sheet.setRowHeight(saldoRow, 40);
  sheet.getRange(saldoRow, 1, 1, 4).setBackground(COR.saldo);
  sheet.getRange(saldoRow, 1).setValue('SALDO DO MÊS')
    .setFontColor(COR.saldoFonte).setFontWeight('bold').setFontSize(12);
  sheet.getRange(saldoRow, 3)
    .setFormula(`=SUMIF($E:$E;"${TAG.entrada}";$C:$C)-SUMIF($E:$E;"${TAG.saida}";$C:$C)`)
    .setFontColor(COR.saldoFonte).setFontWeight('bold').setFontSize(12)
    .setNumberFormat(FMT_BRL);
  formatacaoCondicional(sheet, `C${saldoRow}:C${saldoRow}`);

  // ── Formatação ─────────────────────────────────────────────────────────────
  const g = COR.protegido;
  [
    `A${entStart}:A${entEnd}`,
    `C${entStart}:C${entEnd}`,
    `A${saiStart}:A${saiEnd}`,
    `C${saiStart}:C${saiEnd}`,
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
  const existing = ss.getSheetByName(nome);
  if (!existing) return ss.insertSheet(nome);
  existing.clearContents();
  existing.clearFormats();
  existing.clearNotes();
  existing.getRange(1, 1, existing.getMaxRows(), existing.getMaxColumns()).clearDataValidations();
  existing.setConditionalFormatRules([]);
  existing.getCharts().forEach(c => existing.removeChart(c));
  existing.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  return existing;
}
