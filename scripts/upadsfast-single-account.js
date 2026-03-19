// ============================================================
//  UPADSFAST — Google Ads Bulk Upload Script V5
//  Formato: Google Ads Editor Export (154 colunas)
//  Metodo: newFileUpload com blob CSV — aba unica "Google Ads"
//
//  A planilha deve ter UMA aba chamada "Google Ads" com as
//  154 colunas no formato exato do Google Ads Editor Export
//  (ver referencia: planilhavalidagoogle - Copia.csv).
//
//  Ordem das linhas na planilha (respeite esta ordem):
//    1. Campaign rows
//    2. Ad Group rows
//    3. Keyword rows
//    4. Ad (RSA) rows
//    5. Sitelink rows
//    6. Structured Snippet rows
//    7. Callout rows
//
//  O script faz upload em 2 fases:
//    Fase 1 — Campaigns + Ad Groups (aguarda 90s para indexacao)
//    Fase 2 — Keywords, Ads, Sitelinks, Snippets, Callouts
//
//  Ref: developers.google.com/google-ads/scripts/docs/features/bulk-upload
// ============================================================

// ─── CONFIGURACAO ────────────────────────────────────────────────────────────

// PASSO 1: Cole aqui a URL da sua Planilha Google
var SPREADSHEET_URL = 'COLE_A_URL_DA_SUA_PLANILHA_AQUI';

// PASSO 2: Modo de execucao
// true  = apenas visualiza, nao cria nada (recomendado na primeira vez)
// false = cria as campanhas de verdade no Google Ads
var PREVIEW_MODE = true;

// PASSO 3: Nome da aba na planilha (nao alterar se usar o template padrao)
var SHEET_NAME = 'Google Ads';

// Tempo de espera entre Fase 1 e Fase 2 (em milissegundos)
// Necessario para o Google indexar campanhas e grupos antes de criar filhos
var SLEEP_BETWEEN_PHASES = 90000; // 90 segundos

// ─── NAO ALTERE ABAIXO DESTA LINHA ──────────────────────────────────────────

// Indices das colunas discriminantes (0-based) — usados para detectar tipo de linha
// [2]   Campaign Type      → nao vazio = linha de Campaign (ex: "Search")
// [30]  Ad Group           → nao vazio = pode ser Ad Group, Keyword, ou Ad
// [63]  Criterion Type     → nao vazio = Keyword ("Broad"/"Phrase"/"Exact")
// [100] Ad type            → nao vazio = Ad RSA ("Responsive search ad")
// [142] Link Text          → nao vazio = Sitelink
// [145] Header             → nao vazio = Structured Snippet
// [147] Callout text       → nao vazio = Callout

var COL_CAMPAIGN_TYPE  = 2;
var COL_AD_GROUP       = 30;
var COL_CRITERION_TYPE = 63;
var COL_AD_TYPE        = 100;
var COL_LINK_TEXT      = 142;
var COL_HEADER         = 145;
var COL_CALLOUT_TEXT   = 147;

function main() {
  var ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  var sheet = ss.getSheetByName(SHEET_NAME);

  Logger.log('========================================');
  Logger.log('UPADSFAST V5 — Bulk Upload (Google Ads Editor Format)');
  Logger.log('Planilha: ' + ss.getName());
  Logger.log('Aba: ' + SHEET_NAME);
  Logger.log('Modo: ' + (PREVIEW_MODE ? 'PREVIEW — nenhuma alteracao sera feita' : 'APLICAR — criando campanhas agora'));
  Logger.log('========================================');

  if (!sheet) {
    Logger.log('[ERRO] Aba "' + SHEET_NAME + '" nao encontrada.');
    Logger.log('Verifique se o nome da aba esta correto e tente novamente.');
    return;
  }

  var values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    Logger.log('[PULAR] Aba sem dados (apenas cabecalho ou vazia).');
    return;
  }

  Logger.log('Total de linhas (incluindo cabecalho): ' + values.length);
  Logger.log('Total de colunas: ' + values[0].length);

  var headerRow = values[0];
  var dataRows  = values.slice(1);

  // ── Separar linhas por fase ──────────────────────────────────────────────
  // Fase 1: Campaign rows + Ad Group rows (entidades pai — devem existir primeiro)
  // Fase 2: Keyword, Ad, Sitelink, Snippet, Callout rows (entidades filho)

  var phase1Rows = [headerRow];
  var phase2Rows = [headerRow];

  for (var i = 0; i < dataRows.length; i++) {
    var row = dataRows[i];

    var campaignType  = String(row[COL_CAMPAIGN_TYPE]  || '').trim();
    var adGroup       = String(row[COL_AD_GROUP]       || '').trim();
    var criterionType = String(row[COL_CRITERION_TYPE] || '').trim();
    var adType        = String(row[COL_AD_TYPE]        || '').trim();
    var linkText      = String(row[COL_LINK_TEXT]      || '').trim();
    var snippetHeader = String(row[COL_HEADER]         || '').trim();
    var calloutText   = String(row[COL_CALLOUT_TEXT]   || '').trim();

    // Linha de Campaign: tem campaignType preenchido (ex: "Search"), sem Ad Group
    var isCampaign = campaignType !== '' &&
                     adGroup === '' &&
                     criterionType === '' &&
                     adType === '' &&
                     linkText === '' &&
                     snippetHeader === '' &&
                     calloutText === '';

    // Linha de Ad Group: tem adGroup mas sem os demais discriminantes de filho
    var isAdGroup = adGroup !== '' &&
                    campaignType === '' &&
                    criterionType === '' &&
                    adType === '' &&
                    linkText === '' &&
                    snippetHeader === '' &&
                    calloutText === '';

    if (isCampaign || isAdGroup) {
      phase1Rows.push(row);
    } else {
      phase2Rows.push(row);
    }
  }

  Logger.log('');
  Logger.log('Fase 1 (Campanhas + Grupos): ' + (phase1Rows.length - 1) + ' linha(s)');
  Logger.log('Fase 2 (Palavras-chave, Anuncios, Extensoes): ' + (phase2Rows.length - 1) + ' linha(s)');

  // ── Fase 1: Campaign + Ad Group ───────────────────────────────────────────
  if (phase1Rows.length > 1) {
    Logger.log('');
    Logger.log('>>> Fase 1: Enviando Campanhas e Grupos de Anuncios...');
    uploadRows(phase1Rows, 'upadsfast-phase1-campaigns-adgroups.csv');

    if (phase2Rows.length > 1) {
      var sleepSecs = Math.round(SLEEP_BETWEEN_PHASES / 1000);
      Logger.log('  Aguardando ' + sleepSecs + 's para indexacao das entidades pai...');
      Utilities.sleep(SLEEP_BETWEEN_PHASES);
    }
  }

  // ── Fase 2: Keywords, Ads, Extensoes ─────────────────────────────────────
  if (phase2Rows.length > 1) {
    Logger.log('');
    Logger.log('>>> Fase 2: Enviando Palavras-chave, Anuncios e Extensoes...');
    uploadRows(phase2Rows, 'upadsfast-phase2-keywords-ads-extensions.csv');
  }

  Logger.log('');
  Logger.log('========================================');
  Logger.log('UPADSFAST V5: Finalizado.');
  Logger.log('Verifique em: Ferramentas > Acoes em massa > Uploads');
  Logger.log('========================================');
}

// ─── uploadRows ──────────────────────────────────────────────────────────────
// Converte as linhas para CSV e faz upload via newFileUpload

function uploadRows(rows, fileName) {
  var csv    = buildCsv(rows);
  var blob   = Utilities.newBlob(csv, 'text/csv', fileName);
  var upload = AdsApp.bulkUploads().newFileUpload(blob);
  upload.forCampaignManagement();

  try {
    if (PREVIEW_MODE) {
      upload.preview();
      Logger.log('  [PREVIEW] ' + fileName + ': ' + (rows.length - 1) + ' linha(s) enviada(s).');
    } else {
      upload.apply();
      Logger.log('  [OK] ' + fileName + ': ' + (rows.length - 1) + ' linha(s) aplicada(s).');
    }
  } catch (e) {
    Logger.log('  [ERRO] ' + fileName + ': ' + e.message);
  }
}

// ─── buildCsv ────────────────────────────────────────────────────────────────
// Constroi string CSV (comma-separated) a partir de uma matriz de valores.
// Celulas com virgulas, aspas duplas ou quebras de linha sao envolvidas em aspas.

function buildCsv(rows) {
  var lines = [];
  for (var i = 0; i < rows.length; i++) {
    var cells = [];
    for (var j = 0; j < rows[i].length; j++) {
      cells.push(escapeCell(rows[i][j]));
    }
    lines.push(cells.join(','));
  }
  return lines.join('\n');
}

// ─── escapeCell ──────────────────────────────────────────────────────────────
// Normaliza um valor de celula para CSV valido:
//   - null/undefined → string vazia
//   - "none" (case-insensitive) → string vazia (Google Ads rejeita "None")
//   - contem virgula, aspas ou \n → envolve em aspas duplas, escapa aspas internas

function escapeCell(value) {
  if (value == null || value === undefined) return '';

  var text = String(value).trim();

  if (text.toLowerCase() === 'none') return '';

  if (text.indexOf(',') > -1 ||
      text.indexOf('"')  > -1 ||
      text.indexOf('\n') > -1 ||
      text.indexOf('\r') > -1) {
    text = '"' + text.replace(/"/g, '""') + '"';
  }

  return text;
}
