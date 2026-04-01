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
 * POS=Saldo Anterior  E=Entradas  F=Fixos  V=Variáveis  PJF=PJ Faturamento  PJC=PJ Custos  IA=Aporte
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

// Tags de seção — coluna E oculta. Totais usam SUMIF nessas tags:
// inserir/remover linhas não quebra nada; basta copiar a tag da linha vizinha.
const TAG = {
  entrada:      'E',
  fixo:         'F',
  variavel:     'V',
  pjFat:        'PJF',
  pjCusto:      'PJC',
  invAporte:    'IA',
  posFinanceira:'POS',
};

// Itens da seção Saldo Anterior — valores manuais (snapshot do saldo atual)
const ITEMS_POS_FINANCEIRA = [
  'Conta corrente PF',
  'Renda Fixa (CDB, LCI, Tesouro)',
  'Renda Variável (Ações, FIIs, ETFs)',
  'Criptomoedas',
  'Outros ativos financeiros',
  'Conta corrente PJ',
  'Investimentos PJ',
];

const CAT_INVESTIMENTO = [
  'Aporte Renda Fixa',
  'Aporte Renda Variável',
  'Aporte Cripto',
  'Aporte Filho',
  'Aporte Outros',
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

const DASH_HEADERS = [
  'Mês', 'Entradas PF', 'Gastos Fixos', 'Gastos Variáveis', 'Aportes', 'Saldo PF',
  '', 'Faturamento PJ', 'Custos PJ', 'Saldo PJ', '', 'Ativos Financeiros',
];

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
  'Seguro carro',
  'Seguro vida',
  'Empregada',
  'Mesada',
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
  'Aluguel PJ',
  'Condomínio PJ',
  'Gastos com escritório',
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
  ...CAT_INVESTIMENTO,
];

// LOG_ROW — primeira linha de dados no log, derivado de calcLayout().
// saldoRow(1) + gap(2) + título(1) + cabeçalho(1) = +5 → LOG_ROW é a linha seguinte.
const LOG_ROW = calcLayout().saldoRow + 5;

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
    .addItem('Fechar mês (bloquear edição)',             'fecharMes')
    .addItem('Reabrir mês (desbloquear edição)',         'reabrirMes')
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
    `O que será feito:\n• ${aviso}\n• Locale configurado para pt_BR (datas dd/mm/aaaa, moeda R$, vírgula decimal).\n• Cada aba terá resumo, budget e log de transações.\n\nContinuar?`,
    ui.ButtonSet.YES_NO
  );
  if (ok !== ui.Button.YES) return;

  try { ss.setSpreadsheetLocale('pt_BR'); } catch (e) {}
  try { ss.setSpreadsheetTimeZone('America/Sao_Paulo'); } catch (e) {}

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
    '\n\n── SALDO ANTERIOR ───────────────────────' +
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

// ─── FECHAR / REABRIR MÊS ────────────────────────────────────────────────────

// Valida que a aba ativa é mensal (ex: Jan/2026) e retorna { ui, sheet, nome }.
// Retorna null (e mostra alerta) se a aba não é mensal.
function obterAbaMensal() {
  const ui    = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const nome  = sheet.getName();

  if (!/^[A-Za-z]{3}\/\d{4}$/.test(nome)) {
    ui.alert('Abra uma aba mensal (ex: Jan/2026) antes de usar esta função.');
    return null;
  }
  return { ui, sheet, nome };
}

function fecharMes() {
  const ctx = obterAbaMensal();
  if (!ctx) return;
  const { ui, sheet, nome } = ctx;

  const ok = ui.alert(
    'Fechar Mês',
    `Bloquear a aba "${nome}"?\n\nTodas as células ficarão protegidas contra edição acidental. ` +
    'Use "Reabrir mês" para desbloquear depois.',
    ui.ButtonSet.YES_NO
  );
  if (ok !== ui.Button.YES) return;

  // Remove proteções existentes e aplica proteção total (sem ranges editáveis)
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  const protection = sheet.protect()
    .setDescription(`Mês fechado — "${nome}" está bloqueado para edição.`);
  protection.setWarningOnly(true);

  sheet.setTabColor('#4CAF50'); // verde = fechado
  ui.alert(`Mês "${nome}" fechado! A aba está protegida contra edições.`);
}

function reabrirMes() {
  const ctx = obterAbaMensal();
  if (!ctx) return;
  const { ui, sheet, nome } = ctx;

  const ok = ui.alert(
    'Reabrir Mês',
    `Desbloquear a aba "${nome}" para edição?`,
    ui.ButtonSet.YES_NO
  );
  if (ok !== ui.Button.YES) return;

  aplicarProtecao(sheet, calcLayout());
  sheet.setTabColor(null);
  ui.alert(`Mês "${nome}" reaberto! Áreas editáveis restauradas.`);
}

// ─── INSTRUÇÕES DE USO ────────────────────────────────────────────────────────

function mostrarInstrucoes() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Instruções de uso — Controle Financeiro',
    'PREENCHIMENTO DIÁRIO\n' +
    `  • Anote cada transação no LOG (linhas a partir de ${LOG_ROW}):\n` +
    '    Data | Descrição | Categoria | Valor\n' +
    '  • A Categoria (dropdown) determina em qual seção o valor aparece.\n' +
    '  • Gastos lançados como positivos; o sistema subtrai automaticamente.\n\n' +
    'BUDGET MENSAL\n' +
    '  • Preencha a coluna B nas seções Gastos Fixos e Variáveis.\n' +
    '  • Use "Copiar budget do mês anterior" para reaproveitar os valores.\n\n' +
    'VALORES MANUAIS\n' +
    '  • Saldo Anterior (seção no topo): saldo em conta, investimentos,\n' +
    '    renda fixa/variável, cripto — atualize todo mês para acompanhar.\n' +
    '  • Rendimento do mês (seção Investimentos): ganho ou perda.\n\n' +
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
 * Layout dinâmico — posições calculadas a partir dos arrays de categorias.
 * Adicionar/remover itens nos arrays (ITEMS_POS_FINANCEIRA, CAT_FIXO, etc.)
 * ajusta todas as linhas automaticamente, incluindo LOG_ROW.
 *
 * Ordem das seções:
 *  1        — Título
 *  SALDO ANTERIOR      (tag: POS)   — snapshot de saldos PF + PJ
 *  ENTRADAS            (tag: E)
 *  GASTOS FIXOS        (tag: F)     — com budget
 *  GASTOS VARIÁVEIS    (tag: V)     — com budget
 *  PJ / CNPJ           (tag: PJF / PJC)
 *  INVESTIMENTOS       (tag: IA)    — aportes por tipo + rendimento manual
 *  SALDO DO MÊS
 *  LOG DE TRANSAÇÕES   — começa em LOG_ROW (calculado)
 */
function montarAbaMensal(sheet, mesNome, ano) {
  sheet.setConditionalFormatRules([]);
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 20);  // E — tag de seção (invisível)

  // ── Posições de linha — calculadas a partir dos arrays de categorias ──────
  const L = calcLayout();

  // ── Título ─────────────────────────────────────────────────────────────────
  sheet.setRowHeight(1, 42);
  sheet.getRange(1, 1, 1, 4).merge()
    .setValue(`${mesNome} / ${ano}`)
    .setBackground(COR.titulo).setFontColor(COR.tituloFonte)
    .setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // ── SALDO ANTERIOR ─────────────────────────────────────────────────────────
  cabecalhoSecao(sheet, L.posHeader, 'SALDO ANTERIOR', COR.secao, COR.secaoFonte, ['', '', 'Saldo atual', '']);
  ITEMS_POS_FINANCEIRA.forEach((item, i) => {
    linhaItem(sheet, L.posStart + i, item, TAG.posFinanceira, null, null, null);
    sheet.getRange(L.posStart + i, 3).setNumberFormat(FMT_BRL);
  });
  sheet.getRange(L.posTotal, 1, 1, 4).setBackground(COR.total);
  sheet.getRange(L.posTotal, 1).setValue('TOTAL ATIVOS FINANCEIROS').setFontWeight('bold');
  sheet.getRange(L.posTotal, 3)
    .setFormula(`=SUMIF($E:$E;"${TAG.posFinanceira}";$C:$C)`)
    .setFontWeight('bold').setNumberFormat(FMT_BRL);

  // ── ENTRADAS ───────────────────────────────────────────────────────────────
  cabecalhoSecao(sheet, L.entHeader, 'ENTRADAS', COR.secao, COR.secaoFonte, ['', '', 'Real', '']);
  CAT_ENTRADA.forEach((cat, i) => {
    linhaItem(sheet, L.entStart + i, cat, TAG.entrada, null, sumifCategoria(L.entStart + i), null);
  });
  linhaTotalSecao(sheet, L.entTotal, 'TOTAL ENTRADAS', TAG.entrada, null, false);

  // ── GASTOS FIXOS ───────────────────────────────────────────────────────────
  montarSecaoGastos(sheet, L.fixHeader, 'GASTOS FIXOS', L.fixStart, CAT_FIXO, TAG.fixo, L.fixTotal, 'TOTAL FIXOS');

  // ── GASTOS VARIÁVEIS ───────────────────────────────────────────────────────
  montarSecaoGastos(sheet, L.varHeader, 'GASTOS VARIÁVEIS', L.varStart, CAT_VARIAVEL, TAG.variavel, L.varTotal, 'TOTAL VARIÁVEIS');

  // CF cobre diff de fixos e variáveis
  formatacaoDiferenca(sheet, `D${L.fixStart}:D${L.varTotal}`);

  // ── PJ / CNPJ ──────────────────────────────────────────────────────────────
  cabecalhoSecao(sheet, L.pjHeader, 'PJ / CNPJ', COR.pj, COR.pjFonte, ['', '', 'Real', '']);
  linhaItem(sheet, L.pjFatRow, 'Faturamento PJ', TAG.pjFat, null, sumifCategoria(L.pjFatRow), null);
  CAT_PJ_CUSTO.forEach((cat, i) => {
    linhaItem(sheet, L.pjCustoStart + i, cat, TAG.pjCusto, null, sumifCategoria(L.pjCustoStart + i), null);
  });
  sheet.getRange(L.pjSaldoRow, 1, 1, 4).setBackground(COR.pjTotal);
  sheet.getRange(L.pjSaldoRow, 1).setValue('SALDO PJ').setFontWeight('bold');
  sheet.getRange(L.pjSaldoRow, 3)
    .setFormula(`=SUMIF($E:$E;"${TAG.pjFat}";$C:$C)-SUMIF($E:$E;"${TAG.pjCusto}";$C:$C)`)
    .setFontWeight('bold').setNumberFormat(FMT_BRL);
  formatacaoDiferenca(sheet, `C${L.pjSaldoRow}:C${L.pjSaldoRow}`);

  // ── INVESTIMENTOS ──────────────────────────────────────────────────────────
  cabecalhoSecao(sheet, L.invHeader, 'INVESTIMENTOS', COR.secao, COR.secaoFonte, ['', '', 'Valor', '']);
  CAT_INVESTIMENTO.forEach((cat, i) => {
    linhaItem(sheet, L.invStart + i, cat, TAG.invAporte, null, sumifCategoria(L.invStart + i), null);
  });
  linhaTotalSecao(sheet, L.invTotal, 'TOTAL APORTES', TAG.invAporte, null, false);

  // Rendimento do mês — entrada manual, verde=ganho / vermelho=perda
  sheet.getRange(L.invRendRow, 1).setValue('Rendimento do mês');
  sheet.getRange(L.invRendRow, 3).setNumberFormat(FMT_BRL);
  setTag(sheet, L.invRendRow, '');
  formatacaoDiferenca(sheet, `C${L.invRendRow}:C${L.invRendRow}`);

  // ── SALDO DO MÊS ───────────────────────────────────────────────────────────
  sheet.setRowHeight(L.saldoRow, 40);
  sheet.getRange(L.saldoRow, 1, 1, 4).setBackground(COR.saldo);
  sheet.getRange(L.saldoRow, 1).setValue('SALDO DO MÊS')
    .setFontColor(COR.saldoFonte).setFontWeight('bold').setFontSize(12);
  sheet.getRange(L.saldoRow, 3)
    .setFormula(
      `=SUMIF($E:$E;"${TAG.entrada}";$C:$C)` +
      `-SUMIF($E:$E;"${TAG.fixo}";$C:$C)` +
      `-SUMIF($E:$E;"${TAG.variavel}";$C:$C)` +
      `-SUMIF($E:$E;"${TAG.invAporte}";$C:$C)`
    )
    .setFontColor(COR.saldoFonte).setFontWeight('bold').setFontSize(12)
    .setNumberFormat(FMT_BRL);
  formatacaoDiferenca(sheet, `C${L.saldoRow}:C${L.saldoRow}`);

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

  aplicarCinzaFormulas(sheet, L);
  aplicarProtecao(sheet, L);

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, LOG_ROW - 3, 4).setVerticalAlignment('middle');
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
  DASH_HEADERS.forEach((h, i) => {
    sheet.getRange(2, i + 1)
      .setValue(h)
      .setBackground(COR.secao).setFontColor(COR.secaoFonte)
      .setFontWeight('bold').setHorizontalAlignment('center');
  });

  MESES.forEach(({ abrev }, idx) => {
    const row = idx + 3;
    const aba = `${abrev}/${ANO}`;
    const s   = (tag, col) => `SUMIF('${aba}'!E:E;"${tag}";'${aba}'!${col}:${col})`;

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

  // ── GRÁFICOS ─────────────────────────────────────────────────────────────
  // Remove gráficos existentes para evitar acumulação ao recriar
  sheet.getCharts().forEach(c => sheet.removeChart(c));

  const mesesRange  = sheet.getRange('A3:A14');   // labels: Jan–Dez
  const chartRow1   = totalRow + 3;               // primeira linha de gráficos
  const chartRow2   = chartRow1 + 20;             // segunda linha de gráficos

  // Chart A — Saldo PF + PJ (linha)
  sheet.insertChart(sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(mesesRange)
    .addRange(sheet.getRange('F3:F14'))  // Saldo PF
    .addRange(sheet.getRange('J3:J14'))  // Saldo PJ
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setPosition(chartRow1, 1, 0, 0)
    .setOption('title', 'Saldo Mensal — PF vs PJ')
    .setOption('width', 700).setOption('height', 350)
    .setOption('legend', { position: 'bottom' })
    .setOption('curveType', 'function')
    .setOption('series', { 0: { color: '#2196F3' }, 1: { color: '#4CAF50' } })
    .build());

  // Chart B — Entradas vs Gastos (barras agrupadas)
  sheet.insertChart(sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(mesesRange)
    .addRange(sheet.getRange('B3:B14'))  // Entradas
    .addRange(sheet.getRange('C3:C14'))  // Fixos
    .addRange(sheet.getRange('D3:D14'))  // Variáveis
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setPosition(chartRow1, 7, 0, 0)
    .setOption('title', 'Entradas vs Gastos PF')
    .setOption('width', 700).setOption('height', 350)
    .setOption('legend', { position: 'bottom' })
    .setOption('isStacked', false)
    .setOption('series', { 0: { color: '#4CAF50' }, 1: { color: '#FF9800' }, 2: { color: '#F44336' } })
    .build());

  // Chart C — Faturamento vs Custos PJ (barras)
  sheet.insertChart(sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(mesesRange)
    .addRange(sheet.getRange('H3:H14'))  // Faturamento PJ
    .addRange(sheet.getRange('I3:I14'))  // Custos PJ
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setPosition(chartRow2, 1, 0, 0)
    .setOption('title', 'Faturamento vs Custos PJ')
    .setOption('width', 700).setOption('height', 350)
    .setOption('legend', { position: 'bottom' })
    .setOption('series', { 0: { color: '#2196F3' }, 1: { color: '#E91E63' } })
    .build());

  // Chart D — Evolução Ativos Financeiros (linha)
  sheet.insertChart(sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(mesesRange)
    .addRange(sheet.getRange('L3:L14'))  // Ativos Financeiros
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setPosition(chartRow2, 7, 0, 0)
    .setOption('title', 'Evolução — Ativos Financeiros')
    .setOption('width', 700).setOption('height', 350)
    .setOption('legend', { position: 'none' })
    .setOption('curveType', 'function')
    .setOption('series', { 0: { color: '#9C27B0' } })
    .build());

  // Chart E — Gastos Fixos vs Variáveis (pizza/donut)
  // Dados auxiliares: 2 linhas com totais anuais para gerar apenas 2 fatias
  const chartRow3 = chartRow2 + 20;
  const pieDataRow = chartRow3 - 2;
  sheet.getRange(pieDataRow, 1).setValue('Gastos Fixos');
  sheet.getRange(pieDataRow, 2).setFormula(`=C${totalRow}`);
  sheet.getRange(pieDataRow + 1, 1).setValue('Gastos Variáveis');
  sheet.getRange(pieDataRow + 1, 2).setFormula(`=D${totalRow}`);

  sheet.insertChart(sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange(pieDataRow, 1, 2, 2))
    .setPosition(chartRow3, 1, 0, 0)
    .setOption('title', 'Gastos Fixos vs Variáveis (Ano)')
    .setOption('width', 500).setOption('height', 350)
    .setOption('pieHole', 0.4)
    .setOption('legend', { position: 'bottom' })
    .setOption('slices', { 0: { color: '#FF9800' }, 1: { color: '#F44336' } })
    .build());

  // ── ACUMULADO NO ANO ──────────────────────────────────────────────────────
  const acumHeader = chartRow3 + 20;
  sheet.getRange(acumHeader, 1, 1, 12).merge()
    .setValue('ACUMULADO NO ANO')
    .setBackground(COR.titulo).setFontColor(COR.tituloFonte)
    .setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.getRange(acumHeader + 1, 1, 1, 12).setValues([DASH_HEADERS]);
  sheet.getRange(acumHeader + 1, 1, 1, 12)
    .setBackground(COR.secao).setFontColor(COR.secaoFonte)
    .setFontWeight('bold').setHorizontalAlignment('center');

  MESES.forEach(({ abrev }, idx) => {
    const row  = acumHeader + 2 + idx;
    const dRow = 3 + idx; // corresponding data row in the main table
    sheet.getRange(row, 1).setValue(abrev);
    DASH_COLS.forEach(col => {
      const letra = colLetter(col);
      sheet.getRange(row, col)
        .setFormula(`=SUM(${letra}$3:${letra}${dRow})`)
        .setNumberFormat(FMT_BRL);
    });
    sheet.getRange(row, 1, 1, 12).setBackground(idx % 2 === 0 ? '#f7f9fc' : '#ffffff');
  });

  sheet.setFrozenRows(2);
}

// ─── HELPERS ──────────────────────────────────────────────────────────────────

// Monta uma seção de gastos com budget/real/diferença (Fixos ou Variáveis)
function montarSecaoGastos(sheet, cabRow, titulo, startRow, itens, tag, totalRow, labelTotal) {
  cabecalhoSecao(sheet, cabRow, titulo, COR.secao, COR.secaoFonte, ['', 'Budget', 'Real', 'Diferença']);
  itens.forEach((cat, i) => {
    const row = startRow + i;
    linhaItem(sheet, row, cat, tag, null, sumifCategoria(row), `=IF(B${row}="";"";B${row}-C${row})`);
    sheet.getRange(row, 2).setNumberFormat(FMT_BRL);
  });
  linhaTotalSecao(sheet, totalRow, labelTotal, tag,
    `=SUMIF($E:$E;"${tag}";$B:$B)-SUMIF($E:$E;"${tag}";$C:$C)`, true);
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
      .setFormula(`=SUMIF($E:$E;"${tag}";$B:$B)`)
      .setFontWeight('bold').setNumberFormat(FMT_BRL);
  }
  sheet.getRange(row, 3)
    .setFormula(`=SUMIF($E:$E;"${tag}";$C:$C)`)
    .setFontWeight('bold').setNumberFormat(FMT_BRL);
  if (diffFormula) {
    sheet.getRange(row, 4).setFormula(diffFormula).setFontWeight('bold').setNumberFormat(FMT_BRL);
  }
}

// Retorna a fórmula SUMIF que busca a categoria do log pelo label da linha
function sumifCategoria(row) {
  return `=SUMIF($C$${LOG_ROW}:$C;A${row};$D$${LOG_ROW}:$D)`;
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
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
  sheet.setConditionalFormatRules([]);
  sheet.getCharts().forEach(c => sheet.removeChart(c));
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
}

// Calcula posições de linha de todas as seções a partir dos arrays de categorias.
// Chamado por montarAbaMensal, aplicarCinzaFormulas e aplicarProtecao.
function calcLayout() {
  const L = {};

  L.posHeader    = 3;
  L.posStart     = L.posHeader + 1;
  L.posEnd       = L.posHeader + ITEMS_POS_FINANCEIRA.length;
  L.posTotal     = L.posEnd + 1;

  L.entHeader    = L.posTotal + 2;
  L.entStart     = L.entHeader + 1;
  L.entEnd       = L.entHeader + CAT_ENTRADA.length;
  L.entTotal     = L.entEnd + 1;

  L.fixHeader    = L.entTotal + 2;
  L.fixStart     = L.fixHeader + 1;
  L.fixEnd       = L.fixHeader + CAT_FIXO.length;
  L.fixTotal     = L.fixEnd + 1;

  L.varHeader    = L.fixTotal + 2;
  L.varStart     = L.varHeader + 1;
  L.varEnd       = L.varHeader + CAT_VARIAVEL.length;
  L.varTotal     = L.varEnd + 1;

  L.pjHeader     = L.varTotal + 2;
  L.pjFatRow     = L.pjHeader + 1;
  L.pjCustoStart = L.pjFatRow + 1;
  L.pjCustoEnd   = L.pjFatRow + CAT_PJ_CUSTO.length;
  L.pjSaldoRow   = L.pjCustoEnd + 1;

  L.invHeader    = L.pjSaldoRow + 2;
  L.invStart     = L.invHeader + 1;
  L.invEnd       = L.invHeader + CAT_INVESTIMENTO.length;
  L.invTotal     = L.invEnd + 1;
  L.invRendRow   = L.invTotal + 1;

  L.saldoRow     = L.invRendRow + 2;

  return L;
}

// Fundo cinza nas células com fórmulas automáticas — indica "não editar"
function aplicarCinzaFormulas(sheet, L) {
  const g = COR.protegido;

  [
    // Labels (col A) — alimentam SUMIF; renomear quebraria os cálculos
    `A${L.posStart}:A${L.posEnd}`,
    `A${L.entStart}:A${L.entEnd}`,
    `A${L.fixStart}:A${L.fixEnd}`,
    `A${L.varStart}:A${L.varEnd}`,
    `A${L.pjFatRow}:A${L.pjCustoEnd}`,
    `A${L.invStart}:A${L.invEnd}`, `A${L.invRendRow}`,
    // Fórmulas SUMIF (col C) — calculadas automaticamente a partir do log
    `C${L.entStart}:C${L.entEnd}`,
    `C${L.fixStart}:C${L.fixEnd}`,
    `C${L.varStart}:C${L.varEnd}`,
    `C${L.pjFatRow}:C${L.pjCustoEnd}`,
    `C${L.invStart}:C${L.invEnd}`,
    // Fórmulas de diferença (col D)
    `D${L.fixStart}:D${L.fixEnd}`,
    `D${L.varStart}:D${L.varEnd}`,
  ].forEach(r => sheet.getRange(r).setBackground(g));
}

// Proteção da aba: avisa ao editar células de fórmula
// Áreas editáveis: posição financeira, budget, rendimento, patrimônio, log
function aplicarProtecao(sheet, L) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());

  const protection = sheet.protect()
    .setDescription('Esta célula contém uma fórmula automática. Edite apenas as áreas em branco e o log.');
  protection.setWarningOnly(true);

  protection.setUnprotectedRanges([
    sheet.getRange(`C${L.posStart}:C${L.posEnd}`),     // Saldo Anterior
    sheet.getRange(`B${L.fixStart}:B${L.fixEnd}`),      // Budget Gastos Fixos
    sheet.getRange(`B${L.varStart}:B${L.varEnd}`),      // Budget Gastos Variáveis
    sheet.getRange(`C${L.invRendRow}`),                  // Rendimento do mês
    sheet.getRange(`A${LOG_ROW}:D2000`),                 // Log de transações
  ]);
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
