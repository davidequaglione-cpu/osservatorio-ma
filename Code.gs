// ── Osservatorio M&A Abruzzo — Google Apps Script backend ────────────────────
// Incollare in: Google Sheet › Estensioni › Apps Script
// Deploy: Web App, esegui come "Me", accesso "Anyone"
// Copiare l'URL del deployment nel campo APPS_SCRIPT_URL di index.html

const COLS = [
  "deal_id","data_annuncio","data_closing","anno","fonte_primaria",
  "tipologia_operazione","deal_origination","quota_acquisita_pct",
  "equity_investito_mln","enterprise_value_mln","ev_ebitda_x","ev_sales_x",
  "modalita_pagamento","addon_flag","platform_company","cross_border_flag",
  "acquirente_nome","acquirente_tipo","acquirente_paese","acquirente_estero",
  "coinvestitori","advisor_legale_acquirente","advisor_fin_acquirente",
  "target_nome","target_cf_piva","target_provincia","target_comune",
  "settore_pem","codice_ateco","attivita_descrizione","target_dimensione",
  "fatturato_t0_mln","fatturato_t1_mln","fatturato_t2_mln",
  "fatturato_cagr_pct","ebitda_mln","ebitda_margin_pct","dipendenti",
  "affidabilita","url_fonte","data_inserimento","note",
  "compilatore_nome","compilatore_email","compilatore_studio",
];

function doGet() {
  return ContentService.createTextOutput("Osservatorio M&A — endpoint attivo");
}

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = JSON.parse(e.postData.contents);

    // Intestazioni se il foglio è vuoto
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(COLS);
      const hdr = sheet.getRange(1, 1, 1, COLS.length);
      hdr.setFontWeight("bold");
      hdr.setBackground("#1B4F8A");
      hdr.setFontColor("#FFFFFF");
      sheet.setFrozenRows(1);
    }

    // Controlla duplicati su deal_id
    const existing = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 1).getValues().flat();
    if (existing.includes(data.deal_id)) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: "duplicate", deal_id: data.deal_id }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    sheet.appendRow(COLS.map(col => data[col] ?? ""));

    // Auto-resize colonne principali
    sheet.autoResizeColumns(1, COLS.length);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", deal_id: data.deal_id }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
