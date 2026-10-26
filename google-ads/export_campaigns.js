/**
 * ================================================================
 *  Kit Painel Multicanal — Coletor Google Ads
 *  AF Tecnologia e Marketing
 * ================================================================
 *  ONDE RODA: dentro do Google Ads (NÃO é o Apps Script da planilha).
 *    Google Ads -> Ferramentas -> Ações em massa -> Scripts -> (+) novo script.
 *
 *  O QUE FAZ: puxa as campanhas dos últimos 30 dias e grava na aba "GoogleAds"
 *    da MESMA planilha do painel (custo, cliques, conversões, etc.), por dia.
 *
 *  CONFIGURAÇÃO (2 passos):
 *   1) Compartilhe a Google Sheet do painel (permissão de Editor) com a conta
 *      Google que administra este Google Ads.
 *   2) Cole o ID da planilha em SHEET_ID abaixo (o trecho da URL entre /d/ e /edit).
 *
 *  DEPOIS: Autorizar (Executar/Visualizar) e AGENDAR para rodar 1x por dia.
 *  Ref.: https://developers.google.com/google-ads/scripts/docs/reference/adsapp/adsapp_report
 * ================================================================
 */

var SHEET_ID = 'COLE_AQUI_O_ID_DA_PLANILHA';
var TAB      = 'GoogleAds';

function main() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB) || ss.insertSheet(TAB);
  sheet.clearContents();
  sheet.appendRow(['data', 'campanha', 'impressions', 'clicks', 'custo', 'conversoes', 'valor_conversao']);

  // GAQL: campanhas com entrega nos últimos 30 dias
  var report = AdsApp.report(
    "SELECT segments.date, campaign.name, metrics.impressions, metrics.clicks, " +
    "metrics.cost_micros, metrics.conversions, metrics.conversions_value " +
    "FROM campaign " +
    "WHERE segments.date DURING LAST_30_DAYS AND metrics.impressions > 0");

  var rows = report.rows();
  var buf = [];
  while (rows.hasNext()) {
    var row = rows.next();
    buf.push([
      row['segments.date'],                       // yyyy-mm-dd
      row['campaign.name'],
      Number(row['metrics.impressions']),
      Number(row['metrics.clicks']),
      Number(row['metrics.cost_micros']) / 1e6,   // micros -> reais
      Number(row['metrics.conversions']),
      Number(row['metrics.conversions_value'])
    ]);
  }
  if (buf.length) {
    sheet.getRange(2, 1, buf.length, 7).setValues(buf);
  }
  Logger.log('GoogleAds: ' + buf.length + ' linhas gravadas.');
}
