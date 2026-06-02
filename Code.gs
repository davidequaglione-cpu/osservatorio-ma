// Osservatorio M&A Abruzzo - Google Apps Script backend standalone
// Deploy: Web App, esegui come "Me", accesso "Anyone".

const SHEET_NAME = "Segnalazioni";
const SPREADSHEET_NAME = "Osservatorio M&A Abruzzo - Segnalazioni";
const SPREADSHEET_ID_PROPERTY = "OSSERVATORIO_MA_SPREADSHEET_ID";

const COLS = [
  "deal_id","form_version","data_annuncio","data_closing",
  "anno","fonte_primaria","tipologia_operazione","deal_origination",
  "quota_acquisita_pct","equity_investito_mln","enterprise_value_mln","ev_ebitda_x",
  "ev_sales_x","equity_investito_fascia","enterprise_value_fascia","equity_value_mln",
  "prezzo_pagato_mln","ev_ebit_x","pe_x","pbv_x",
  "premio_predeal_pct","modalita_pagamento","addon_flag","platform_company",
  "cross_border_flag","razionale_principale","razionale_secondario","integrazione_verticale_flag",
  "consolidamento_orizzontale_flag","espansione_geografica_flag","acquisizione_tecnologia_flag","passaggio_generazionale_flag",
  "distress_rescue_flag","diversificazione_flag","geografia_deal","target_sede_operativa_abruzzo_flag",
  "acquirente_sede_operativa_abruzzo_flag","prossimita_territoriale","acquirente_nome","acquirente_tipo",
  "acquirente_paese","acquirente_estero","coinvestitori","advisor_legale_acquirente",
  "advisor_fin_acquirente","buyer_settore","buyer_quotato_flag","buyer_ricavi_mln",
  "buyer_ebitda_mln","buyer_ebit_mln","buyer_utile_netto_mln","buyer_pfn_mln",
  "buyer_debt_ebitda_x","buyer_roa_pct","buyer_roe_pct","buyer_ros_pct",
  "buyer_export_pct","buyer_brevetti_flag","target_nome","target_cf_piva",
  "target_provincia","target_comune","settore_pem","codice_ateco",
  "attivita_descrizione","target_dimensione","fatturato_fascia","fatturato_t0_mln",
  "fatturato_t1_mln","fatturato_t2_mln","fatturato_cagr_pct","ebitda_mln",
  "ebitda_margin_pct","ebit_mln","utile_netto_mln","pfn_mln",
  "debt_ebitda_x","roa_pct","roe_pct","ros_pct",
  "export_pct","brevetti_flag","dipendenti","sinergie_costo_attese_mln",
  "sinergie_ricavo_attese_mln","sinergie_fiscali_flag","sinergie_finanziarie_flag","timing_sinergie_mesi",
  "costi_integrazione_attesi_mln","rischi_due_diligence","rischio_antitrust_flag","management_target_permane_flag",
  "fondatore_permane_flag","investitore_finanziario_flag","ricavi_post_12_mln","ricavi_post_24_mln",
  "ricavi_post_36_mln","ebitda_post_12_mln","ebitda_post_24_mln","ebitda_post_36_mln",
  "ebit_post_12_mln","ebit_post_24_mln","ebit_post_36_mln","utile_post_12_mln",
  "utile_post_24_mln","utile_post_36_mln","pfn_post_12_mln","pfn_post_24_mln",
  "pfn_post_36_mln","dipendenti_post_12","dipendenti_post_24","dipendenti_post_36",
  "sinergie_realizzate_pct","tempi_sinergie_effettivi_mesi","costi_integrazione_effettivi_mln","integrazione_it_stato",
  "brand_integrato_flag","rete_commerciale_integrata_flag","ristrutturazioni_chiusure_flag","impairment_goodwill_flag",
  "write_off_flag","dismissione_successiva_flag","acquisizioni_correttive_flag","giudizio_successo_deal",
  "success_score","affidabilita","url_fonte","data_inserimento",
  "note","compilatore_nome","compilatore_email","compilatore_studio",
];

function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(SPREADSHEET_ID_PROPERTY);
  if (!id) {
    const ss = SpreadsheetApp.create(SPREADSHEET_NAME);
    id = ss.getId();
    props.setProperty(SPREADSHEET_ID_PROPERTY, id);
  }
  return SpreadsheetApp.openById(id);
}

function getSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(COLS);
    const hdr = sheet.getRange(1, 1, 1, COLS.length);
    hdr.setFontWeight("bold");
    hdr.setBackground("#1B4F8A");
    hdr.setFontColor("#FFFFFF");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function csvEscape_(value) {
  const text = value === null || value === undefined ? "" : String(value);
  return '"' + text.replace(/"/g, '""') + '"';
}

function doGet(e) {
  const ss = getSpreadsheet_();
  const params = (e && e.parameter) ? e.parameter : {};
  if (params.format === "csv" || params.export === "csv") {
    const sheet = getSheet_();
    const values = sheet.getDataRange().getValues();
    const csv = values.map(row => row.map(csvEscape_).join(",")).join("\n");
    return ContentService
      .createTextOutput(csv)
      .setMimeType(ContentService.MimeType.CSV);
  }

  return ContentService
    .createTextOutput(JSON.stringify({
      status: "ok",
      message: "Osservatorio M&A endpoint attivo",
      spreadsheet_url: ss.getUrl(),
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const sheet = getSheet_();
    const data = JSON.parse(e.postData.contents);

    const existing = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 1).getValues().flat();
    if (existing.includes(data.deal_id)) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: "duplicate", deal_id: data.deal_id }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    sheet.appendRow(COLS.map(col => data[col] ?? ""));
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
