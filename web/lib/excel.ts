import ExcelJS from "exceljs";
import type { GeneratedCopy, CopyContext } from "./templates";

// ─── STYLE HELPERS ───────────────────────────────────────────────────────────

function styleHeader(cell: ExcelJS.Cell) {
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E79" } };
  cell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 10 };
  cell.alignment = { horizontal: "center", vertical: "middle" };
  cell.border = {
    top: { style: "thin", color: { argb: "FFCCCCCC" } },
    bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
    left: { style: "thin", color: { argb: "FFCCCCCC" } },
    right: { style: "thin", color: { argb: "FFCCCCCC" } },
  };
}

function styleData(cell: ExcelJS.Cell, wrapText = false) {
  cell.font = { size: 10 };
  cell.alignment = { horizontal: "left", vertical: "middle", wrapText };
  cell.border = {
    top: { style: "thin", color: { argb: "FFCCCCCC" } },
    bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
    left: { style: "thin", color: { argb: "FFCCCCCC" } },
    right: { style: "thin", color: { argb: "FFCCCCCC" } },
  };
}

function addHeaders(ws: ExcelJS.Worksheet, headers: string[]) {
  ws.getRow(1).height = 22;
  headers.forEach((h, i) => {
    const cell = ws.getRow(1).getCell(i + 1);
    cell.value = h;
    styleHeader(cell);
  });
}

// col index 146 (0-based) = "Snippet Values" → needs wrapText for multiline values
function addDataRow(ws: ExcelJS.Worksheet, rowIndex: number, values: string[]) {
  values.forEach((v, i) => {
    const cell = ws.getRow(rowIndex).getCell(i + 1);
    cell.value = v;
    styleData(cell, i === 146);
  });
}

function autoWidth(ws: ExcelJS.Worksheet, min = 8, max = 40) {
  ws.columns.forEach(col => {
    let maxLen = min;
    col.eachCell?.({ includeEmpty: true }, cell => {
      const text = String(cell.value ?? "");
      const len = text.split("\n")[0].length + 2;
      if (len > maxLen) maxLen = len;
    });
    col.width = Math.min(maxLen, max);
  });
}

// ─── TYPES ───────────────────────────────────────────────────────────────────

export interface MatrixItem {
  copy: GeneratedCopy;
  ctx: CopyContext;
  language: string;
  url: string;
}

// ─── 154 COLUMN HEADERS — Google Ads Editor Export Format ────────────────────
// Source: planilhavalidagoogle - Copia.csv (100% valid Google Ads reference)
// Column order MUST be preserved exactly — used verbatim by newFileUpload

const GOOGLE_ADS_HEADERS: string[] = [
  "Campaign", "Labels", "Campaign Type", "Networks", "Budget", "Budget type",
  "EU political ads", "Standard conversion goals", "Customer acquisition", "Languages",
  "Bid Strategy Type", "Bid Strategy Name", "Target CPA", "Desktop Bid Modifier",
  "Mobile Bid Modifier", "Tablet Bid Modifier", "TV Screen Bid Modifier", "Start Date",
  "End Date", "Broad match keywords", "Ad Schedule", "Ad rotation", "Content exclusions",
  "Targeting method", "Exclusion method", "Audience targeting", "Flexible Reach",
  "AI Max", "Text customization", "Final URL expansion", "Ad Group", "Max CPC",
  "Max CPM", "Max CPV", "Target CPV", "Percent CPC", "Target CPM", "Target ROAS",
  "Target CPC", "Display Network Custom Bid Type", "Optimized targeting",
  "Strict age and gender targeting", "Search term matching", "Ad Group Type", "Channels",
  "Audience name", "Age demographic", "Gender demographic", "Income demographic",
  "Parental status demographic", "Remarketing audience segments", "Interest categories",
  "Life events", "Custom audience segments", "Detailed demographics",
  "Remarketing audience exclusions", "Tracking template", "Final URL suffix",
  "Custom parameters", "Age", "Bid Modifier", "Final URL", "Final mobile URL",
  "Criterion Type", "Gender", "ID", "Location", "Reach", "Location groups", "Radius",
  "Unit", "Account keyword type", "Keyword", "First page bid", "Top of page bid",
  "First position bid", "Quality score", "Landing page experience", "Expected CTR",
  "Ad relevance", "Promotion target", "Occasion", "Language", "Currency",
  "Discount modifier", "Monetary discount", "Percent discount", "Orders over amount",
  "Promotion code", "Barcode type", "Barcode number", "QR code text", "Promotion start",
  "Promotion end", "Terms and conditions", "Link to additional terms",
  "Upgraded extension", "Source", "Link source", "Image Size", "Ad type",
  "Headline 1", "Headline 1 position", "Headline 2", "Headline 2 position",
  "Headline 3", "Headline 3 position", "Headline 4", "Headline 4 position",
  "Headline 5", "Headline 5 position", "Headline 6", "Headline 6 position",
  "Headline 7", "Headline 7 position", "Headline 8", "Headline 8 position",
  "Headline 9", "Headline 9 position", "Headline 10", "Headline 10 position",
  "Headline 11", "Headline 11 position", "Headline 12", "Headline 12 position",
  "Headline 13", "Headline 13 position", "Headline 14", "Headline 14 position",
  "Headline 15", "Headline 15 position",
  "Description 1", "Description 1 position", "Description 2", "Description 2 position",
  "Description 3", "Description 3 position", "Description 4", "Description 4 position",
  "Path 1", "Path 2", "IP address", "Link Text", "Description Line 1",
  "Description Line 2", "Header", "Snippet Values", "Callout text",
  "Campaign Status", "Ad Group Status", "Status", "Approval Status", "Ad strength", "Comment",
];

// ─── ROW BUILDERS (0-based index mapping) ────────────────────────────────────
// Each builder returns a 154-element string array.
// Key index reference (0-based):
//   [0]   Campaign           [2]  Campaign Type     [3]  Networks
//   [4]   Budget             [5]  Budget type       [9]  Languages
//   [10]  Bid Strategy Type  [12] Target CPA        [17] Start Date
//   [18]  End Date           [19] Broad match kws   [20] Ad Schedule
//   [21]  Ad rotation        [22] Content exclusions[23] Targeting method
//   [24]  Exclusion method   [25] Audience targeting[26] Flexible Reach
//   [27]  AI Max             [28] Text customization[29] Final URL expansion
//   [30]  Ad Group           [31] Max CPC            [32] Max CPM
//   [34]  Target CPV         [36] Target CPM         [39] DN Custom Bid Type
//   [40]  Optimized targeting[41] Strict age/gender  [42] Search term matching
//   [43]  Ad Group Type      [44] Channels           [61] Final URL
//   [63]  Criterion Type     [72] Keyword            [96] Upgraded extension
//   [97]  Source             [98] Link source        [100] Ad type
//   [101+2i] Headline i+1    [102+2i] Headline i+1 position  (i=0..14)
//   [131+2i] Description i+1 [132+2i] Description i+1 position (i=0..3)
//   [139] Path 1             [140] Path 2            [142] Link Text
//   [143] Description Line 1 [144] Description Line 2[145] Header
//   [146] Snippet Values     [147] Callout text
//   [148] Campaign Status    [149] Ad Group Status   [150] Status

function emptyRow(): string[] {
  return new Array(154).fill("");
}

function buildCampaignRow(copy: GeneratedCopy, ctx: CopyContext, today: string): string[] {
  const r = emptyRow();
  r[0]  = copy.campaign;
  r[2]  = "Search";
  r[3]  = "Google search";
  r[4]  = String(Number(ctx.budget) || 10);
  r[5]  = "Daily";
  r[6]  = "Doesn't have EU political ads"; // valor exato do CSV de referência válido
  r[7]  = "Account-level";
  r[8]  = "Bid equally";
  r[9]  = "All";
  r[10] = "Maximize conversions";
  if (ctx.target_cpa && Number(ctx.target_cpa) > 0) {
    r[12] = ctx.target_cpa;
  }
  r[17] = today;
  // r[18] End Date — deixar vazio (sem data de término)
  r[19] = "On";
  r[20] = "[]";
  r[21] = "Optimize for clicks";
  r[22] = "[]";
  r[23] = "Location of presence";
  r[24] = "Location of presence";
  r[25] = "Audience segments";
  r[26] = "Audience segments";
  r[27] = "Disabled";
  r[28] = "Disabled";
  r[29] = "Disabled";
  r[148] = "Enabled";
  return r;
}

function buildAdGroupRow(copy: GeneratedCopy): string[] {
  const r = emptyRow();
  r[0]  = copy.campaign;
  r[9]  = "All";
  r[25] = "Audience segments";
  r[26] = "Audience segments;Genders;Ages;Parental status;Household incomes";
  r[30] = copy.adGroup;
  r[31] = "0.01";
  r[32] = "0.01";
  r[34] = "0.01";
  r[36] = "0.01";
  r[39] = "None";
  r[40] = "Disabled";
  r[41] = "Disabled";
  r[42] = "Disabled";
  r[43] = "Standard";
  r[44] = "[]";
  r[148] = "Enabled";
  r[149] = "Enabled";
  return r;
}

function buildKeywordRow(copy: GeneratedCopy, keyword: string, matchType: string): string[] {
  const r = emptyRow();
  r[0]   = copy.campaign;
  r[30]  = copy.adGroup;
  r[63]  = matchType;   // "Broad" | "Phrase" | "Exact"
  r[72]  = keyword;
  r[148] = "Enabled";
  r[149] = "Enabled";
  r[150] = "Enabled";
  return r;
}

function buildAdRow(copy: GeneratedCopy, url: string): string[] {
  const r = emptyRow();
  r[0]   = copy.campaign;
  r[30]  = copy.adGroup;
  r[61]  = url;
  r[100] = "Responsive search ad";

  // Headlines 1-15: alternating [value, position] pairs starting at index 101
  // Position " -" means free rotation (no pinning)
  const headlines = [...copy.headlines];
  while (headlines.length < 15) headlines.push("");
  for (let i = 0; i < 15; i++) {
    r[101 + i * 2] = headlines[i];
    r[102 + i * 2] = headlines[i] ? " -" : "";
  }

  // Descriptions 1-4: alternating [value, position] pairs starting at index 131
  const descriptions = [...copy.descriptions];
  while (descriptions.length < 4) descriptions.push("");
  for (let i = 0; i < 4; i++) {
    r[131 + i * 2] = descriptions[i];
    r[132 + i * 2] = descriptions[i] ? " -" : "";
  }

  r[139] = copy.path1;
  r[140] = copy.path2;
  r[148] = "Enabled";
  r[149] = "Enabled";
  r[150] = "Enabled";
  return r;
}

function buildSitelinkRow(
  copy: GeneratedCopy,
  linkText: string,
  sitelinkUrl: string,
  d1: string,
  d2: string,
): string[] {
  const r = emptyRow();
  r[0]   = copy.campaign;
  r[17]  = "[]";          // Start Date — [] = sem agendamento (referência usa este valor)
  r[18]  = "[]";          // End Date — [] = sem data de término (documentação confirma)
  r[20]  = "[]";          // Ad Schedule
  r[61]  = sitelinkUrl;   // Final URL
  r[96]  = "[]";          // Upgraded extension
  r[97]  = "Advertiser";  // Source
  r[98]  = "Advertiser";  // Link source
  r[142] = linkText;      // Link Text
  r[143] = d1;            // Description Line 1
  r[144] = d2;            // Description Line 2
  r[148] = "Enabled";     // Campaign Status
  r[150] = "Enabled";     // Status
  return r;
}

function buildSnippetRow(
  copy: GeneratedCopy,
  header: string,
  values: string[],
): string[] {
  const r = emptyRow();
  r[0]   = copy.campaign;
  r[96]  = "[]";                   // Upgraded extension
  r[97]  = "Advertiser";           // Source
  r[98]  = "Advertiser";           // Link source
  r[145] = header;                 // Header (e.g. "Types")
  r[146] = values.join("\n");      // Snippet Values — newline-separated, quoted in TSV
  r[148] = "Enabled";              // Campaign Status
  r[150] = "Enabled";              // Status
  return r;
}

function buildCalloutRow(copy: GeneratedCopy, callout: string): string[] {
  const r = emptyRow();
  r[0]   = copy.campaign;
  r[17]  = "[]";          // Start Date — [] = sem agendamento (referência usa este valor)
  r[18]  = "[]";          // End Date — [] = sem data de término (documentação confirma)
  r[20]  = "[]";          // Ad Schedule
  r[96]  = "[]";          // Upgraded extension
  r[97]  = "Advertiser";  // Source
  r[98]  = "Advertiser";  // Link source
  r[147] = callout;       // Callout text
  r[148] = "Enabled";     // Campaign Status
  r[150] = "Enabled";     // Status
  return r;
}

// ─── MAIN EXPORT ─────────────────────────────────────────────────────────────
// Gera planilha com UMA aba "Google Ads" no formato exato do Google Ads Editor
// Export (154 colunas). Este formato é aceito diretamente por newFileUpload.
//
// Ordem das linhas (respeita dependências de entidade do Google Ads):
//   1. Campaign rows  (campanha deve existir antes do grupo)
//   2. Ad Group rows  (grupo deve existir antes de anúncios e palavras-chave)
//   3. Keyword rows
//   4. Ad (RSA) rows
//   5. Sitelink rows
//   6. Structured Snippet rows
//   7. Callout rows

export async function generateXlsx(items: MatrixItem[]): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Google Ads");
  ws.properties.tabColor = { argb: "FF1A3A6B" };

  addHeaders(ws, GOOGLE_ADS_HEADERS);

  const today = new Date().toISOString().slice(0, 10); // YYYY-MM-DD
  let rowIndex = 2;

  // ── 1. Campaigns (unique) ─────────────────────────────────────────────────
  const seenCampaigns = new Set<string>();
  for (const { copy, ctx } of items) {
    if (!seenCampaigns.has(copy.campaign)) {
      seenCampaigns.add(copy.campaign);
      addDataRow(ws, rowIndex++, buildCampaignRow(copy, ctx, today));
    }
  }

  // ── 2. Ad Groups (unique) ─────────────────────────────────────────────────
  const seenGroups = new Set<string>();
  for (const { copy } of items) {
    const key = `${copy.campaign}|${copy.adGroup}`;
    if (!seenGroups.has(key)) {
      seenGroups.add(key);
      addDataRow(ws, rowIndex++, buildAdGroupRow(copy));
    }
  }

  // ── 3. Keywords ───────────────────────────────────────────────────────────
  const seenKeywords = new Set<string>();
  for (const { copy, ctx } of items) {
    const keyword = ctx.product.toLowerCase();
    const key = `${copy.campaign}|${copy.adGroup}|${keyword}`;
    if (!seenKeywords.has(key)) {
      seenKeywords.add(key);
      addDataRow(ws, rowIndex++, buildKeywordRow(copy, keyword, "Broad"));
    }
  }

  // ── 4. Ads (RSA) ──────────────────────────────────────────────────────────
  for (const { copy, url } of items) {
    addDataRow(ws, rowIndex++, buildAdRow(copy, url));
  }

  // ── 5. Sitelinks ──────────────────────────────────────────────────────────
  const seenSitelinks = new Set<string>();
  for (const { copy } of items) {
    for (const s of copy.sitelinks) {
      const key = `${copy.campaign}|${s.text}`;
      if (!seenSitelinks.has(key)) {
        seenSitelinks.add(key);
        addDataRow(ws, rowIndex++, buildSitelinkRow(copy, s.text, s.url, s.d1, s.d2));
      }
    }
  }

  // ── 6. Structured Snippets ────────────────────────────────────────────────
  const seenSnippets = new Set<string>();
  for (const { copy } of items) {
    const key = `${copy.campaign}|${copy.snippetHeader}`;
    if (!seenSnippets.has(key)) {
      seenSnippets.add(key);
      addDataRow(ws, rowIndex++, buildSnippetRow(copy, copy.snippetHeader, copy.snippetValues));
    }
  }

  // ── 7. Callouts ───────────────────────────────────────────────────────────
  const seenCallouts = new Set<string>();
  for (const { copy } of items) {
    for (const callout of copy.callouts) {
      const key = `${copy.campaign}|${callout}`;
      if (!seenCallouts.has(key)) {
        seenCallouts.add(key);
        addDataRow(ws, rowIndex++, buildCalloutRow(copy, callout));
      }
    }
  }

  ws.views = [{ state: "frozen", ySplit: 1 }];
  autoWidth(ws);

  const arrayBuffer = await wb.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer);
}
