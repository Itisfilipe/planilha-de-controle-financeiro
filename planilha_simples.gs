/**
 * CONTROLE FINANCEIRO SIMPLES — Google Apps Script
 *
 * Versão simplificada para quem só quer anotar e acompanhar
 * entradas e saídas do mês, sem complicação.
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
 *   Data | Descrição | Categoria | Valor
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

// Retorna CATEGORIAS + descrições da aba Dívidas (se existir)
function getTodasCategorias(ss) {
  const dividas = ss.getSheetByName('Dívidas');
  if (!dividas) return CATEGORIAS;
  const lastRow = dividas.getLastRow();
  if (lastRow < 3) return CATEGORIAS;
  const descs = dividas.getRange(3, 1, lastRow - 2, 1).getValues()
    .map(([d]) => d)
    .filter(d => d && d !== 'TOTAIS');
  return [...CATEGORIAS, ...descs];
}

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

  const ok = ui.alert(
    'Criar Mês Atual',
    `Criar a aba "${nome}" com resumo e log de transações?` +
    (ss.getSheetByName(nome) ? `\n\nAviso: a aba "${nome}" já existe e será recriada.` : ''),
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
    const ok = ui.alert(`A aba "${nomeAba}" já existe.`, 'Recriar? (dados serão perdidos)', ui.ButtonSet.YES_NO);
    if (ok !== ui.Button.YES) return;
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
    .requireValueInList(getTodasCategorias(SpreadsheetApp.getActiveSpreadsheet()), true)
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
    ['COMO USAR — CONTROLE FINANCEIRO SIMPLES'],
    [''],
    ['PREENCHIMENTO DIÁRIO'],
    ['No LOG (parte de baixo de cada aba mensal), preencha uma linha por transação:'],
    ['  • Coluna A — Data (seletor de data disponível)'],
    ['  • Coluna B — Descrição (ex: "Supermercado Extra")'],
    ['  • Coluna C — Categoria (escolha no dropdown: Salário, Alimentação, etc.)'],
    ['  • Coluna D — Valor (sempre positivo; a categoria determina se é entrada ou saída)'],
    [''],
    ['Os totais por categoria no resumo (parte de cima) atualizam automaticamente.'],
    [''],
    ['MENU FINANCEIRO'],
    ['  • Criar mês atual — cria a aba do mês corrente'],
    ['  • Novo mês... — cria qualquer mês/ano (ex: Jan/2027)'],
    ['  • Resumo do mês — exibe totais de entradas, saídas e saldo'],
    ['  • Atualizar categorias — recria o resumo com as categorias atuais (log preservado)'],
    ['  • Criar / atualizar aba Dívidas — acompanhe parcelas e financiamentos'],
    [''],
    ['PERSONALIZAR CATEGORIAS'],
    ['  1. Abra Extensões > Apps Script'],
    ['  2. Edite os arrays CAT_ENTRADA e CAT_SAIDA no topo do código'],
    ['  3. Salve e volte para a planilha'],
    ['  4. Use Financeiro > Atualizar categorias'],
    ['  O resumo é reconstruído com as novas categorias. Dados do log são preservados.'],
    [''],
    ['DÍVIDAS E PARCELAS'],
    ['  1. Use Financeiro > Criar / atualizar aba Dívidas'],
    ['  2. Preencha: Descrição, Valor total, Parcelas e Início'],
    ['  3. A descrição da dívida vira automaticamente uma categoria no dropdown do log'],
    ['  4. Ao lançar um pagamento no log mensal usando essa categoria, "Parcelas pagas" atualiza sozinho'],
    ['  5. Restantes e Saldo devedor são calculados automaticamente'],
    ['  Se adicionar novas dívidas depois, use "Atualizar categorias" para atualizar os dropdowns.'],
    [''],
    ['CÉLULAS EM CINZA'],
    ['  Contêm fórmulas automáticas — não edite.'],
    ['  Um aviso aparecerá se você tentar modificar essas células.'],
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
  [3, 12, 19, 26, 34].forEach(r => {
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
  sheet.setColumnWidth(5, 20);

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
      .requireValueInList(getTodasCategorias(SpreadsheetApp.getActiveSpreadsheet()), true)
      .setAllowInvalid(false)
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
  // Limpa a área do resumo (entre título e log), preservando o log
  // LOG_ROW - 4: para antes do título "LOG DE TRANSAÇÕES" (que está em LOG_ROW - 2)
  sheet.getRange(2, 1, LOG_ROW - 4, 5).clearContent().clearFormat()
    .clearDataValidations().setBackground(null);
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
    sheet.getRange(`A${LOG_ROW}:D2000`),
  ]);

  sheet.getRange(1, 1, saldoRow, 4).setVerticalAlignment('middle');
}

// ─── ABA DÍVIDAS ──────────────────────────────────────────────────────────────

function criarAbaDividas() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const ok = ui.alert(
    'Dívidas e Parcelas',
    'Criar (ou atualizar) a aba "Dívidas"?\n\n' +
    '• Descrições na aba Dívidas viram categorias no log mensal.\n' +
    '• Ao lançar um pagamento no log com essa categoria, "Parcelas pagas" atualiza automaticamente.\n' +
    '• Dados existentes serão preservados.',
    ui.ButtonSet.YES_NO
  );
  if (ok !== ui.Button.YES) return;

  let sheet = ss.getSheetByName('Dívidas');
  const isNew = !sheet;
  if (isNew) sheet = ss.insertSheet('Dívidas');

  const headers = ['Descrição', 'Valor total', 'Parcelas', 'Valor mensal', 'Início', 'Parcelas pagas', 'Restantes', 'Saldo devedor'];
  [220, 130, 80, 130, 100, 120, 100, 140].forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // Título
  sheet.setRowHeight(1, 42);
  if (isNew) {
    sheet.getRange(1, 1, 1, 8).merge()
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

  // Fórmula COUNTIF: conta pagamentos no log de todos os meses
  const ano = new Date().getFullYear();
  const countif = (r) => MESES.map(({ abrev }) => {
    const aba = `${abrev}/${ano}`;
    return `IFERROR(COUNTIF('${aba}'!C$${LOG_ROW}:C;A${r});0)`;
  }).join('+');

  // Fórmulas automáticas para cada linha de dados
  for (let r = 3; r <= 100; r++) {
    // D = Valor mensal = Valor total / Parcelas
    sheet.getRange(r, 4).setFormula(`=IF(AND(B${r}<>"";C${r}<>"");B${r}/C${r};"")`);
    // F = Parcelas pagas = contagem automática no log de todos os meses
    sheet.getRange(r, 6).setFormula(`=IF(A${r}="";"";${countif(r)})`);
    // G = Restantes = Parcelas - Pagas
    sheet.getRange(r, 7).setFormula(`=IF(AND(C${r}<>"";F${r}<>"");C${r}-F${r};"")`);
    // H = Saldo devedor = Valor mensal * Restantes
    sheet.getRange(r, 8).setFormula(`=IF(AND(D${r}<>"";G${r}<>"");D${r}*G${r};"")`);
  }

  // Totais
  const totalRow = 102;
  sheet.getRange(totalRow, 1).setValue('TOTAIS').setFontWeight('bold');
  sheet.getRange(totalRow, 1, 1, 8).setBackground(COR.total);
  sheet.getRange(totalRow, 4).setFormula('=SUM(D3:D101)').setFontWeight('bold').setNumberFormat(FMT_BRL);
  sheet.getRange(totalRow, 6).setFormula('=SUM(F3:F101)').setFontWeight('bold');
  sheet.getRange(totalRow, 8).setFormula('=SUM(H3:H101)').setFontWeight('bold').setNumberFormat(FMT_BRL);

  // Formatação cinza nas colunas com fórmula (D, F, G, H)
  ['D3:D101', 'F3:F101', 'G3:G101', 'H3:H101'].forEach(r =>
    sheet.getRange(r).setBackground(COR.protegido)
  );

  // Proteção — editável: A (descrição), B (valor total), C (parcelas), E (início)
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  const protection = sheet.protect()
    .setDescription('Colunas D, F, G e H contêm fórmulas. Edite A, B, C e E.');
  protection.setWarningOnly(true);
  protection.setUnprotectedRanges([
    sheet.getRange('A3:C101'),
    sheet.getRange('E3:E101'),
  ]);

  sheet.setFrozenRows(2);
  ss.setActiveSheet(sheet);

  // Atualiza dropdowns em todas as abas para incluir as descrições de dívidas
  const validacao = SpreadsheetApp.newDataValidation()
    .requireValueInList(getTodasCategorias(ss), true)
    .setAllowInvalid(false)
    .build();
  ss.getSheets().forEach(s => {
    if (/^[A-Za-z]{3}\/\d{4}$/.test(s.getName())) {
      s.getRange(`C${LOG_ROW}:C2000`).setDataValidation(validacao);
    }
  });

  ui.alert(
    'Aba "Dívidas" pronta!\n\n' +
    '• Preencha: Descrição, Valor total, Parcelas e Início.\n' +
    '• A descrição vira uma categoria no log mensal.\n' +
    '• Ao lançar pagamentos no log com essa categoria, "Parcelas pagas" atualiza sozinho.\n' +
    '• Use "Atualizar categorias" se adicionar novas dívidas depois.'
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
