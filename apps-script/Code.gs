/*
 * ================================================================
 *  Painel de Performance (white-label) — Backend (Meta + GA4 + Comercial)
 *  AF Tecnologia e Marketing
 * ================================================================
 *  Este script faz 3 coisas:
 *   1) updateMeta()  -> puxa a mídia da Meta (diário, 30 dias) p/ a aba "Meta".
 *   2) doGet()       -> a API que o PAINEL lê (junta Meta + Comercial em JSON).
 *   3) doPost()      -> a API que RECEBE os dados do comercial pela web e
 *                       grava na aba "Comercial".
 *
 *  ---------- CONFIGURAÇÃO (fazer 1 vez) ----------
 *  A) Propriedades do Script (Engrenagem -> Propriedades do script):
 *       META_TOKEN       = token de LEITURA (usuário de sistema, ads_read)
 *       META_ACCOUNT_ID  = act_XXXXXXXXXXXX
 *       SHEET_ID         = (o ID da sua Google Sheet, da URL .../d/ESTE_ID/edit)
 *       GA4_PROPERTY_ID  = (opcional) só o número da propriedade GA4, ex.: 123456789
 *     Para GA4: Serviços (+) -> "Analytics Data". Para Google Ads: usar google_ads_script.js (à parte).
 *  B) Rode  primeiraConfiguracao()  1x (cria abas + gatilho + 1ª carga).
 *  C) Implantar como app da web:  Implantar -> Nova implantação ->
 *       tipo "App da Web"  ·  Executar como: Eu  ·  Quem tem acesso: Qualquer pessoa.
 *       Copie a URL que termina em  /exec  e me envie (vai no painel).
 *  D) Quando publicar nova versão do código, use "Gerenciar implantações" ->
 *       editar -> Nova versão (para a mesma URL continuar valendo).
 * ================================================================
 */

var API_VERSION = 'v20.0';
var DATE_PRESET = 'last_30d';
var META_TAB = 'Meta';
var COM_TAB  = 'Comercial';
var ADS_TAB  = 'Anuncios';
var GA4_TAB  = 'GA4';
var GADS_TAB = 'GoogleAds';   // preenchida pelo google_ads_script.js
var VENDAS_TAB = 'Vendas';    // camada normalizada de vendas (canal + valor + margem)
var VENDAS_COLS = ['data','canal','veiculo','valor','margem','fonte'];
// MODO_CONVERSAO (Propriedade do Script): ecommerce | inside_sales | hibrido  (padrão inside_sales)
//   ecommerce   -> a venda (valor + canal) vem do GA4
//   inside_sales-> a venda vem do comercial (aba Vendas), com o canal marcado
//   hibrido     -> soma as duas trilhas (uma venda entra por uma só)
// colunas da aba Comercial (o comercial preenche do pre_venda em diante)
var COM_COLS = ['data','pre_venda','retorno_vendedor','em_contato_efetivo','em_negociacao','vendas','negociacoes_perdidas','leads_perdidos','receita','margem'];
// mapa coluna_planilha -> campo_do_painel
var COM_MAP = {pre_venda:'preVenda',retorno_vendedor:'retorno',em_contato_efetivo:'emContato',em_negociacao:'emNeg',vendas:'vendas',negociacoes_perdidas:'negPerdidas',leads_perdidos:'leadsPerdidos',receita:'receita',margem:'margem'};

function primeiraConfiguracao(){ garantirAbaComercial(); atualizarTudo(); garantirGatilhoDiario(); }
function atualizarTudo(){ updateMeta(); updateAnuncios_(); coletarGA4_(); }

/* ---------------- 1) META -> aba "Meta" ---------------- */
function updateMeta(){
  var props=PropertiesService.getScriptProperties();
  var token=props.getProperty('META_TOKEN'), acct=props.getProperty('META_ACCOUNT_ID');
  if(!token||!acct) throw new Error('Configure META_TOKEN e META_ACCOUNT_ID nas Propriedades do Script.');
  var fields=['spend','impressions','clicks','actions','video_thruplay_watched_actions','video_p75_watched_actions','video_p95_watched_actions'].join(',');
  var url='https://graph.facebook.com/'+API_VERSION+'/'+acct+'/insights?level=account&time_increment=1&date_preset='+DATE_PRESET+'&fields='+fields+'&limit=500&access_token='+encodeURIComponent(token);
  var rows=fetchAll_(url);
  var sh=sheet_(META_TAB); sh.clearContents();
  var header=['data','investimento','entradas','impressions','clicks','v3','thru','vp75','vp95','ctr','cpm','hook','hold','ret75','ret95'];
  var data=[header];
  rows.forEach(function(r){
    var imp=num_(r.impressions), clk=num_(r.clicks), spend=num_(r.spend);
    var leads=actionVal_(r.actions,'onsite_conversion.messaging_conversation_started_7d');
    var v3=actionVal_(r.actions,'video_view');
    var thru=firstVal_(r.video_thruplay_watched_actions), vp75=firstVal_(r.video_p75_watched_actions), vp95=firstVal_(r.video_p95_watched_actions);
    data.push([r.date_start,spend,leads,imp,clk,v3,thru,vp75,vp95, imp?100*clk/imp:0, imp?1000*spend/imp:0, imp?100*v3/imp:0, imp?100*thru/imp:0, imp?100*vp75/imp:0, imp?100*vp95/imp:0]);
  });
  sh.getRange(1,1,data.length,header.length).setValues(data);
  sh.getRange('A1').setNote('Atualizado: '+new Date());
}
function fetchAll_(url){var out=[],next=url,g=0;while(next&&g<50){g++;var res=UrlFetchApp.fetch(next,{muteHttpExceptions:true});var j=JSON.parse(res.getContentText());if(j.error)throw new Error('Meta API: '+j.error.message);(j.data||[]).forEach(function(d){out.push(d)});next=(j.paging&&j.paging.next)?j.paging.next:null;}return out;}
function actionVal_(a,t){if(!a)return 0;for(var i=0;i<a.length;i++)if(a[i].action_type===t)return num_(a[i].value);return 0;}
function firstVal_(a){return (a&&a.length)?num_(a[0].value):0;}
function num_(v){v=Number(v);return isNaN(v)?0:v;}
function ss_(){var id=PropertiesService.getScriptProperties().getProperty("SHEET_ID");if(id)return SpreadsheetApp.openById(id);var a=SpreadsheetApp.getActiveSpreadsheet();if(a)return a;throw new Error("Sem planilha: crie uma Google Sheet e adicione a propriedade SHEET_ID (o ID está na URL da planilha).");}
function sheet_(n){var ss=ss_();return ss.getSheetByName(n)||ss.insertSheet(n);}

/* ---------------- 1b) META por anúncio -> aba "Anuncios" ---------------- */
function updateAnuncios_(){
  var props=PropertiesService.getScriptProperties();
  var token=props.getProperty('META_TOKEN'), acct=props.getProperty('META_ACCOUNT_ID');
  if(!token||!acct) throw new Error('Configure META_TOKEN e META_ACCOUNT_ID.');
  var tk=encodeURIComponent(token), base='https://graph.facebook.com/'+API_VERSION+'/'+acct;
  // orçamento diário + status por CONJUNTO (daily_budget vem em centavos)
  var adsetInfo={};
  fetchAll_(base+'/adsets?fields=id,daily_budget,effective_status&limit=500&access_token='+tk)
    .forEach(function(a){ adsetInfo[a.id]={budget:num_(a.daily_budget)/100,status:a.effective_status}; });
  // orçamento por CAMPANHA (fallback p/ campanhas com orçamento no nível da campanha - CBO)
  var campBudget={};
  fetchAll_(base+'/campaigns?fields=id,daily_budget&limit=500&access_token='+tk)
    .forEach(function(c){ campBudget[c.id]=num_(c.daily_budget)/100; });
  // status por ANÚNCIO
  var adStatus={};
  fetchAll_(base+'/ads?fields=id,effective_status&limit=1000&access_token='+tk)
    .forEach(function(a){ adStatus[a.id]=a.effective_status; });
  // insights por ANÚNCIO, dia a dia
  var fields=['ad_id','ad_name','adset_id','adset_name','campaign_id','spend','actions'].join(',');
  var rows=fetchAll_(base+'/insights?level=ad&time_increment=1&date_preset='+DATE_PRESET+'&fields='+fields+'&limit=500&access_token='+tk);
  var header=['data','ad_id','anuncio','adset_id','conjunto','status','orcamento_dia','orc_key','spend','conversas'];
  var out=[header];
  rows.forEach(function(r){
    var aset=adsetInfo[r.adset_id]||{}, abud=aset.budget||0;
    var orc = abud>0 ? abud : (campBudget[r.campaign_id]||0);
    var orcKey = abud>0 ? r.adset_id : r.campaign_id;
    var conv=actionVal_(r.actions,'onsite_conversion.messaging_conversation_started_7d');
    out.push([r.date_start, r.ad_id, r.ad_name, r.adset_id, r.adset_name, (adStatus[r.ad_id]||''), orc, orcKey, num_(r.spend), conv]);
  });
  var sh=sheet_(ADS_TAB); sh.clearContents();
  sh.getRange(1,1,out.length,header.length).setValues(out);
  sh.getRange('A1').setNote('Atualizado: '+new Date());
}
function adsData_(){
  var t=readTab_(ADS_TAB); if(!t.rows.length) return [];
  var h=t.h, out=[];
  t.rows.forEach(function(r){ out.push({
    date:iso_(r[h['data']]), ad_id:String(r[h['ad_id']]), anuncio:r[h['anuncio']],
    adset_id:String(r[h['adset_id']]), conjunto:r[h['conjunto']], status:r[h['status']],
    orc:num_(r[h['orcamento_dia']]), orcKey:String(r[h['orc_key']]),
    spend:num_(r[h['spend']]), conv:num_(r[h['conversas']]) }); });
  return out;
}

/* ---------------- 1c) GA4 -> aba "GA4" ----------------
 * Pré-requisitos:
 *  - Propriedade do Script:  GA4_PROPERTY_ID = 123456789  (só o número; ver Admin da propriedade GA4)
 *  - No editor: Serviços (+) -> adicionar "Analytics Data" (liga a Google Analytics Data API do GA4)
 *  - A conta Google que roda o script precisa ter acesso de leitura à propriedade GA4.
 * O relatório vem por DIA e por CANAL (sessionDefaultChannelGroup): Paid Social, Paid Search,
 * Organic Search, Direct, Referral, etc. — é isso que permite ver "de onde veio" no painel.
 */
function coletarGA4_(){
  var propId=PropertiesService.getScriptProperties().getProperty('GA4_PROPERTY_ID');
  if(!propId){ Logger.log('GA4_PROPERTY_ID não configurado — pulando GA4.'); return; }
  if(typeof AnalyticsData==='undefined'){ Logger.log('Ative o serviço avançado "Analytics Data" (Serviços +).'); return; }
  var req={
    dateRanges:[{startDate:'30daysAgo',endDate:'today'}],
    dimensions:[{name:'date'},{name:'sessionDefaultChannelGroup'}],
    metrics:[{name:'totalUsers'},{name:'newUsers'},{name:'sessions'},{name:'engagedSessions'},{name:'conversions'},{name:'totalRevenue'},{name:'ecommercePurchases'}],
    orderBys:[{dimension:{dimensionName:'date'}}],
    limit:100000
  };
  var resp=AnalyticsData.Properties.runReport(req,'properties/'+propId);
  var header=['data','canal','usuarios','novos_usuarios','sessoes','sessoes_engajadas','conversoes','receita','compras'];
  var out=[header];
  (resp.rows||[]).forEach(function(r){
    var d=r.dimensionValues[0].value; // formato yyyymmdd
    var iso=d.slice(0,4)+'-'+d.slice(4,6)+'-'+d.slice(6,8);
    var m=r.metricValues;
    out.push([iso, r.dimensionValues[1].value, num_(m[0].value), num_(m[1].value), num_(m[2].value), num_(m[3].value), num_(m[4].value), num_(m[5].value), num_(m[6].value)]);
  });
  var sh=sheet_(GA4_TAB); sh.clearContents();
  sh.getRange(1,1,out.length,header.length).setValues(out);
  sh.getRange('A1').setNote('Atualizado: '+new Date());
}
function ga4Data_(){
  var t=readTab_(GA4_TAB); if(!t.rows.length) return [];
  var h=t.h, out=[];
  t.rows.forEach(function(r){ out.push({
    date:iso_(r[h['data']]), canal:r[h['canal']], usuarios:num_(r[h['usuarios']]),
    novos:num_(r[h['novos_usuarios']]), sessoes:num_(r[h['sessoes']]),
    engajadas:num_(r[h['sessoes_engajadas']]), conversoes:num_(r[h['conversoes']]),
    receita:num_(r[h['receita']]), compras:num_(r[h['compras']]) }); });
  return out;
}
/* ---------------- 1d) Google Ads (aba preenchida pelo google_ads_script.js) ---------------- */
function gadsData_(){
  var t=readTab_(GADS_TAB); if(!t.rows.length) return [];
  var h=t.h, out=[];
  t.rows.forEach(function(r){ out.push({
    date:iso_(r[h['data']]), campanha:r[h['campanha']], custo:num_(r[h['custo']]),
    clicks:num_(r[h['clicks']]), impressions:num_(r[h['impressions']]),
    conversoes:num_(r[h['conversoes']]), valor:num_(r[h['valor_conversao']]) }); });
  return out;
}
/* ---------------- 1e) Camada normalizada de VENDAS (depende do MODO_CONVERSAO) ---------------- */
function vendasData_(){
  var modo=(PropertiesService.getScriptProperties().getProperty('MODO_CONVERSAO')||'inside_sales').toLowerCase();
  var out=[];
  if(modo==='ecommerce'||modo==='hibrido'){ // venda automática vinda do GA4 (receita por canal)
    var g=readTab_(GA4_TAB);
    if(g.rows.length){ var h=g.h; g.rows.forEach(function(r){ var val=num_(r[h['receita']]); if(val>0)
      out.push({date:iso_(r[h['data']]), canal:r[h['canal']], veiculo:'', valor:val, margem:0, fonte:'ga4'}); }); }
  }
  if(modo==='inside_sales'||modo==='hibrido'){ // venda marcada pelo comercial (aba Vendas)
    var v=readTab_(VENDAS_TAB);
    if(v.rows.length){ var h2=v.h; v.rows.forEach(function(r){ if(!r[h2['data']])return;
      out.push({date:iso_(r[h2['data']]), canal:r[h2['canal']], veiculo:r[h2['veiculo']], valor:num_(r[h2['valor']]), margem:num_(r[h2['margem']]), fonte:r[h2['fonte']]||'comercial'}); }); }
  }
  return out;
}

/* ---------------- 2) API de leitura (painel) ---------------- */
function doGet(e){
  var p=(e&&e.parameter)||{};
  if(p.ads) return json_(adsData_());
  if(p.ga4) return json_(ga4Data_());
  if(p.gads) return json_(gadsData_());
  if(p.vendas) return json_(vendasData_());
  return json_(mergedData_());
}

function readTab_(name){var sh=ss_().getSheetByName(name);if(!sh||sh.getLastRow()<2)return {h:{},rows:[]};var v=sh.getDataRange().getValues();var h={};v[0].forEach(function(x,i){h[String(x).trim()]=i});return {h:h,rows:v.slice(1)};}
function iso_(v){if(v instanceof Date)return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd');return String(v).slice(0,10);}
function mergedData_(){
  var meta=readTab_(META_TAB), com=readTab_(COM_TAB), byDate={};
  meta.rows.forEach(function(r){var d=iso_(r[meta.h['data']]);if(!d)return;var o=byDate[d]=byDate[d]||{date:d};
    o.investimento=num_(r[meta.h['investimento']]);o.entradas=num_(r[meta.h['entradas']]);o.impressions=num_(r[meta.h['impressions']]);o.clicks=num_(r[meta.h['clicks']]);
    o.v3=num_(r[meta.h['v3']]);o.thru=num_(r[meta.h['thru']]);o.vp75=num_(r[meta.h['vp75']]);o.vp95=num_(r[meta.h['vp95']]);});
  com.rows.forEach(function(r){var d=iso_(r[com.h['data']]);if(!d)return;var o=byDate[d]=byDate[d]||{date:d};
    for(var k in COM_MAP){ if(com.h[k]!=null) o[COM_MAP[k]]=num_(r[com.h[k]]); }});
  var keys=['investimento','entradas','impressions','clicks','v3','thru','vp75','vp95','preVenda','retorno','emContato','emNeg','vendas','negPerdidas','leadsPerdidos','receita','margem'];
  var out=[]; Object.keys(byDate).sort().forEach(function(d){var o=byDate[d];keys.forEach(function(k){if(o[k]==null)o[k]=0});out.push(o)});
  return out;
}

/* ---------------- 3) API de escrita (comercial via web) ---------------- */
function doPost(e){
  var rec;
  try{ rec=JSON.parse(e.postData.contents); }catch(err){ return json_({ok:false,error:'JSON inválido'}); }
  if(!rec.date) return json_({ok:false,error:'faltou a data'});
  if(rec.tipo==='venda'){ appendVenda_(rec); return json_({ok:true}); }
  upsertComercial_(rec);
  return json_({ok:true});
}
function appendVenda_(rec){
  var sh=sheet_(VENDAS_TAB);
  if(sh.getLastRow()===0) sh.getRange(1,1,1,VENDAS_COLS.length).setValues([VENDAS_COLS]);
  sh.appendRow([rec.date, rec.canal||'', rec.veiculo||'', num_(rec.valor), num_(rec.margem), 'comercial']);
}
function garantirAbaComercial(){var sh=sheet_(COM_TAB);if(sh.getLastRow()===0)sh.getRange(1,1,1,COM_COLS.length).setValues([COM_COLS]);}
function upsertComercial_(rec){
  garantirAbaComercial();
  var sh=sheet_(COM_TAB), last=sh.getLastRow();
  var dates = last>1 ? sh.getRange(2,1,last-1,1).getValues().map(function(r){return iso_(r[0])}) : [];
  var row = COM_COLS.map(function(c){ return c==='data' ? rec.date : num_(rec[c]); });
  var idx = dates.indexOf(rec.date);
  if(idx>=0) sh.getRange(idx+2,1,1,COM_COLS.length).setValues([row]);
  else sh.appendRow(row);
}

function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function garantirGatilhoDiario(){var t=ScriptApp.getProjectTriggers().some(function(x){return x.getHandlerFunction()==='atualizarTudo'});if(!t)ScriptApp.newTrigger('atualizarTudo').timeBased().everyDays(1).atHour(6).create();}
