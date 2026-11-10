/**
 * 生産管理しくみ  データ置き場（Googleスプレッドシート＋Apps Script） ver 2026-11-10
 *
 * 【設置手順】
 * 1. 会社のGoogleアカウントで新しいスプレッドシートを作る
 * 2. 拡張機能 → Apps Script に、このファイルの中身を貼り付けて保存
 * 2-2. 下の KEY（合言葉）を社内だけで決めた文字に変える。HTMLの 設定 → 合言葉 にも同じ文字を入れる
 * 3. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *      実行ユーザー：自分 ／ アクセスできるユーザー：全員（または会社内の全員）
 * 3-2. Apps Script の画面で「ファイル ＋ → HTML」を押し、名前を index にして、index.html の中身を全部貼って保存
 * 4. 出てきた「ウェブアプリのURL（…/exec）」を、全タブレットのブラウザで開く（ホーム画面に追加すると便利）
 *      → 画面がGoogleから出て、最初から本番。各タブレットの 設定 で「タブレット名・役割・合言葉」だけ入れる
 *      （index.html を直接開いて使うときは、設定 → Apps Script URL にこのURLを貼る）
 * 5. シート「品番マスタ」を作り、A列=客先 / B列=名称 / C列=品番・背番 を入れる（1行目は見出し。HTMLの「ひな形」と同じ形）
 * 6. 毎日の実績Excel保存：setupDailyTrigger を1回だけ実行（Googleドライブのフォルダ名は下の EXPORT_FOLDER_NAME）
 *
 * ※項目を増やしたため、前のバージョンで作った「かんばん実績」シートがある場合は、シートを削除してから使う（自動で作り直す）
 * ※コードや index（画面）を変えたら「デプロイを管理 → 編集 → 新バージョン」で再デプロイが必要（URLは変わらない）
 */
var KEY = 'kyoho-change-me';   // ★合言葉（HTML側の設定と同じにする）
var SHEET_NAME = 'かんばん実績';
var MASTER_NAME = '品番マスタ';
var EXPORT_FOLDER_NAME = '生産管理_実績Excel';
var EXPORT_HOUR = 22;   // 毎日この時刻（時）に当日分をExcelで保存
var TZ = 'Asia/Tokyo';
var HEAD = ['id','日付','タブレット名','品番','協豊背番号','箱数（かんばん枚数）','うちQRなし','開始','完了','作業時間(分)','到着予定','状態','準備OK','検査完了','t0','t1','eta','tr','td','u','kb','stage','maker','qty','tp','ins','adj','ext'];
var ST_OUT = {wait:'待ち', ready:'準備OK', done:'検査完了', cancel:'取消'};
var ST_IN = {'待ち':'wait', '準備OK':'ready', '検査完了':'done', '取消':'cancel'};
var SCAN_ROWS = 3000; // 検索・一覧で読む最大行数（新しい方から）

function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function sheet_(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(SHEET_NAME);
  if (!s) {
    s = ss.insertSheet(SHEET_NAME);
    s.appendRow(HEAD); s.setFrozenRows(1);
    s.getRange('A:A').setNumberFormat('@'); s.getRange('B:B').setNumberFormat('@');
    s.getRange('D:E').setNumberFormat('@'); s.getRange('G:G').setNumberFormat('0');
    s.getRange('H:I').setNumberFormat('@'); s.getRange('K:K').setNumberFormat('@'); s.getRange('M:N').setNumberFormat('@');
  }
  if (s.getLastColumn() < HEAD.length) s.getRange(1, 1, 1, HEAD.length).setValues([HEAD]);   // 列を増やしたときに見出しを更新
  return s;
}
function hm_(ms){ return ms ? Utilities.formatDate(new Date(ms), TZ, 'HH:mm') : ''; }
function ymd_(ms){ return Utilities.formatDate(new Date(ms), TZ, 'yyyy-MM-dd'); }

function jobToRow_(j){
  return [String(j.id), ymd_(j.t1 || j.tp), j.tab || '', String(j.hin), String(j.sei || ''), j.n || 0, j.nq || 0,
    hm_(j.t0), hm_(j.t1), j.t1 ? Math.round((j.t1 - j.t0) / 6000) / 10 : '', hm_(j.eta), ST_OUT[j.st] || j.st, hm_(j.tr), hm_(j.td),
    j.t0 || '', j.t1 || '', j.eta || '', j.tr || '', j.td || '', j.u, JSON.stringify(j.kb || []),
    j.stage == null ? 3 : j.stage, j.maker || '', j.qty == null ? '' : j.qty, j.tp || '', j.ins || 0, JSON.stringify(j.adj || []), JSON.stringify({tab2: j.tab2 || '', ta2: j.ta2 || null, tab3: j.tab3 || '', ses: j.ses || [], n2: j.n2 == null ? null : j.n2, only: j.only || 0, tab0: j.tab0 || '', pz: j.pz || null, paus: j.paus || [], ord: j.ord == null ? null : j.ord})];
}
function rowToJob_(r){
  var o = {id:String(r[0]), tab:r[2], hin:String(r[3]), sei:String(r[4]), n:r[5], nq:r[6], st:ST_IN[r[11]] || 'wait',
    t0:r[14] || null, t1:r[15] || null, eta:r[16] || null, tr:r[17] || undefined, td:r[18] || undefined, u:r[19], kb:safeParse_(r[20]),
    stage:r[21] === '' ? 3 : r[21], maker:String(r[22] || ''), qty:r[23] === '' ? null : r[23], tp:r[24] || null, ins:Number(r[25]) || 0, adj:safeParse_(r[26])};
  var x = {}; try { x = JSON.parse(r[27]) || {}; } catch (e) {}
  o.tab2 = x.tab2 || ''; o.ta2 = x.ta2 || null; o.tab3 = x.tab3 || ''; o.ses = x.ses || []; o.n2 = x.n2 == null ? null : x.n2; o.only = x.only || 0; o.tab0 = x.tab0 || ''; o.pz = x.pz || null; o.paus = x.paus || []; if (x.ord != null) o.ord = x.ord;   // 検査の順番（長押しで入れ替え）
  return o;
}
function safeParse_(t){ try { return JSON.parse(t); } catch (e) { return []; } }

function readRows_(s){
  var last = s.getLastRow();
  if (last < 2) return {start: 2, rows: []};
  var start = Math.max(2, last - SCAN_ROWS + 1);
  return {start: start, rows: s.getRange(start, 1, last - start + 1, HEAD.length).getValues()};
}

function doPost(e){
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ok: false, error: String(err)}); }
  return json_(post_(body));
}
function post_(body){
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (body.key !== KEY) return {ok: false, error: '合言葉がちがいます'};
    if (body.action !== 'upsert') return {ok: false, error: 'unknown action'};
    var s = sheet_(), rd = readRows_(s), idx = {};
    rd.rows.forEach(function(r, i){ idx[String(r[0])] = rd.start + i; });
    var adds = [];
    body.jobs.forEach(function(j){
      var row = jobToRow_(j), at = idx[String(j.id)];
      if (at) {
        var cur = s.getRange(at, 20).getValue();
        if (Number(cur) > Number(j.u)) return;           // サーバー側の方が新しければ上書きしない
        s.getRange(at, 1, 1, HEAD.length).setValues([row]);
      } else adds.push(row);
    });
    if (adds.length) s.getRange(s.getLastRow() + 1, 1, adds.length, HEAD.length).setValues(adds);
    return {ok: true, saved: body.jobs.length};
  } catch (err) {
    return {ok: false, error: String(err)};
  } finally { lock.releaseLock(); }
}

function doGet(e){
  var p = (e && e.parameter) || {};
  if (!p.action && !p.key) return page_();   // URLをそのまま開いたときは画面（index）を出す
  return json_(get_(p));
}
/* 画面：Apps Script の HTMLファイル「index」（index.html の中身）を出す。全タブレットはこのURLを開くだけ */
function page_(){
  return HtmlService.createHtmlOutputFromFile('index').setTitle('生産管理しくみ')
    .addMetaTag('viewport', 'width=device-width,initial-scale=1,viewport-fit=cover');
}
/* Googleから開いた画面からの呼び出し（google.script.run）。中身は doGet / doPost と同じ */
function api(req){
  req = req || {};
  return JSON.stringify(req.body ? post_(req.body) : get_(req.params || {}));
}
function get_(p){
  try {
    if (p.key !== KEY) return {ok: false, error: '合言葉がちがいます'};
    var a = p.action || 'list';
    if (a === 'master') {
      var m = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_NAME);
      if (!m || m.getLastRow() < 2) return {ok: true, master: []};
      var v = m.getRange(2, 1, m.getLastRow() - 1, 3).getValues();
      return {ok: true, master: v.filter(function(r){ return r[1] !== '' || r[2] !== ''; })
        .map(function(r){ return {cust: String(r[0]), name: String(r[1]), hin: String(r[2])}; })};
    }
    if (a === 'list') {
      var rows = readRows_(sheet_()).rows, date = p.date, lim = Date.now() - 2 * 864e5;
      var out = rows.filter(function(r){ return r[0] !== ''; }).map(rowToJob_).filter(function(j){
        return date ? ymd_(j.t1 || j.tp) === date : ((j.st !== 'done' && j.st !== 'cancel') || (j.td || j.t1 || j.tp) >= lim);
      });
      return {ok: true, jobs: out};
    }
    return {ok: false, error: 'unknown action'};
  } catch (err) { return {ok: false, error: String(err)}; }
}

/* ===== 実績シート（タブレットごとに 段取り①+掛け／検査用段取り／検査／異常記録 を横に並べる。1箱当たりは 分/箱） ===== */
function rng_(a, b){ return a && b ? hm_(a) + '～' + hm_(b) : ''; }
function perBox_(ms, n){ return n > 0 && ms > 0 ? Math.round(ms / 60000 / n * 10) / 10 : ''; }
function buildJisseki_(sh, date, rows, andons){   // 1行＝1つの品番の流れ。各工程の欄にボタンを押したタブレット
  sh.setName('実績');
  var E5 = ['', '', '', '', ''], H = '箱数（かんばん枚数）', P = '１箱当たり出来高（分/箱）';
  var out = [['実績　' + date].concat(new Array(19).join('-').split('-')),
    ['', '段取り①+掛け', '', '', '', '', '検査用段取り', '', '', '', '', '検査', '', '', '', '', '異常記録', '', '', ''],
    ['No', 'タブレット', '品番・背番', H, '時間', P, 'タブレット', '品番・背番', H, '時間', P, 'タブレット', '品番・背番', H, '時間', P, 'タブレット', '品番', '数量', '選択項目']];
  var L = rows.map(function(j){ return {k: j.only ? j.tp : (j.t1 || j.tp), j: j}; })
    .concat(andons.map(function(a){ return {k: a.tp, a: a}; })).sort(function(x, y){ return x.k - y.k; });
  L.forEach(function(e, i){
    var no = i + 1, j = e.j, a = e.a;
    if (a) { out.push([no].concat(E5, E5, E5, [a.tab, '', '', '呼び出し：' + a.hin + (a.td ? '（対応 ' + hm_(a.td) + ' ' + (a.sei || '') + '）' : '（未対応）')])); return; }
    var n1 = Number(j.qty == null ? j.n : j.qty), n2 = Number(j.n2 == null ? j.n : j.n2), n3 = Number(j.ins) || 0;
    var b1 = j.only ? E5 : [j.tab, j.hin, n1, rng_(j.t0, j.t1), perBox_(j.t1 - j.t0, n1)];
    var b2 = j.ta2 ? [j.tab2 || '', j.hin, n2, rng_(j.t1, j.ta2), perBox_(j.ta2 - j.t1, n2)] : E5;
    var ses = j.ses || [], ms = ses.reduce(function(x, c){ return x + (c[1] - c[0]); }, 0), nm = j.hin + (j.only ? '（検査のみ）' : '');
    var b3 = ses.length ? [j.tab3 || '', nm, n3, rng_(ses[0][0], ses[ses.length - 1][1]), perBox_(ms, n3)] : (j.only ? [j.tab0 || '', nm, '', '', ''] : E5);
    var ev = [];
    (j.adj || []).forEach(function(x){ ev.push([x[0], x[4] || j.tab3 || '', x[2] + ' ' + (x[1] > 0 ? '+' : '') + x[1] + (x[3] ? '（' + x[3] + '）' : '')]); });
    (j.paus || []).forEach(function(p){ ev.push([p[0], p[1], '一時中断']); });
    if (j.only) ev.push([j.tp, j.tab0 || '', '検査のみ入力 ' + n1 + '箱']);
    ev.sort(function(x, y){ return x[0] - y[0]; });
    var tabs = []; ev.forEach(function(x){ if (x[1] && tabs.indexOf(x[1]) < 0) tabs.push(x[1]); });
    var dsum = (j.adj || []).reduce(function(t, x){ return t + x[1]; }, 0);
    var b4 = ev.length ? [tabs.join('／'), j.hin, (j.adj || []).length ? dsum : '', ev.map(function(x){ return x[2]; }).join(' ／ ')] : ['', '', '', ''];
    out.push([no].concat(b1, b2, b3, b4));
  });
  sh.getRange(1, 1, out.length, 20).setValues(out);
  ['B2:F2', 'G2:K2', 'L2:P2', 'Q2:T2'].forEach(function(r){ sh.getRange(r).merge(); });
  sh.getRange(1, 1).setFontSize(14).setFontWeight('bold');
  sh.getRange(2, 1, 2, 20).setFontWeight('bold').setBackground('#dbeafe').setHorizontalAlignment('center').setWrap(true);
  sh.setFrozenRows(3);
}

/* ===== 毎日の実績をExcelでGoogleドライブへ自動保存 ===== */
function setupDailyTrigger(){
  ScriptApp.getProjectTriggers().forEach(function(t){ if (t.getHandlerFunction() === 'dailyExport') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailyExport').timeBased().everyDays(1).atHour(EXPORT_HOUR).inTimezone(TZ).create();
}
function dailyExport(){ exportDay_(ymd_(Date.now())); }
function exportDay_(date){   // 手動で過去日を出したいときは exportDay_('2026-11-09') を実行
  var jobs = readRows_(sheet_()).rows.filter(function(r){ return r[0] !== ''; }).map(rowToJob_)
    .filter(function(j){ return j.t1 && Number(j.stage) >= 2 && Number(j.stage) < 9 && ymd_(j.t1) === date; }).sort(function(a, b){ return a.t1 - b.t1; });
  var allRows = readRows_(sheet_()).rows.filter(function(r){ return r[0] !== ''; }).map(rowToJob_);
  var andons = allRows.filter(function(j){ return Number(j.stage) === 9 && j.tp && ymd_(j.tp) === date; }).sort(function(a, b){ return a.tp - b.tp; });
  if (!jobs.length && !andons.length) return;
  var tmp = SpreadsheetApp.create('実績_' + date);
  buildJisseki_(tmp.getSheets()[0], date, allRows.filter(function(j){ return j.t1 && Number(j.stage) >= 2 && Number(j.stage) < 9 && j.st !== 'cancel' && ymd_(j.t1) === date; }), andons);
  var d = tmp.insertSheet('実績明細');
  var det = [['日付','タブレット名','品番','協豊背番号','箱数（かんばん枚数）','うちQRなし','開始（最初のQR読取）','完了','作業時間（分）','到着予定','準備OK','検査完了','状態']];
  var tab = {}, hin = {};
  jobs.forEach(function(j){
    var m = Math.round((j.t1 - j.t0) / 6000) / 10;
    det.push([date, j.tab, j.hin, j.sei, j.n, j.nq || 0, hm_(j.t0), hm_(j.t1), m, hm_(j.eta), hm_(j.tr), hm_(j.td), ST_OUT[j.st]]);
    tab[j.tab] = tab[j.tab] || [j.tab, 0, 0, 0]; tab[j.tab][1]++; tab[j.tab][2] += j.n; tab[j.tab][3] += m;
    hin[j.hin] = hin[j.hin] || [j.hin, j.sei, 0, 0]; hin[j.hin][2]++; hin[j.hin][3] += j.n;
  });
  d.getRange(1, 1, det.length, det[0].length).setNumberFormat('@').setValues(det);
  var t = tmp.insertSheet('タブレット別'), tv = [['タブレット名','件数','処理枚数','作業時間合計（分）']];
  Object.keys(tab).forEach(function(k){ tv.push([tab[k][0], tab[k][1], tab[k][2], Math.round(tab[k][3] * 10) / 10]); });
  t.getRange(1, 1, tv.length, 4).setValues(tv);
  var h = tmp.insertSheet('品番別'), hv = [['品番','協豊背番号','件数','処理枚数']];
  Object.keys(hin).forEach(function(k){ hv.push(hin[k]); });
  h.getRange(1, 1, hv.length, 4).setNumberFormat('@').setValues(hv);
  var a = tmp.insertSheet('呼び出し記録'), av = [['日付','呼び出し時刻','工程','タブレット名','理由','対応時刻','対応者','対応までの時間（分）','状態']];
  andons.forEach(function(j){ av.push([date, hm_(j.tp), j.maker, j.tab, j.hin, hm_(j.td), j.sei, j.td ? Math.round((j.td - j.tp) / 6000) / 10 : '', j.st === 'done' ? '対応済み' : j.st === 'cancel' ? '取消' : '呼出中']); });
  a.getRange(1, 1, av.length, 9).setNumberFormat('@').setValues(av);
  SpreadsheetApp.flush();
  var blob = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx',
    {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}}).getBlob().setName('実績_' + date + '.xlsx');
  var it = DriveApp.getFoldersByName(EXPORT_FOLDER_NAME);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(EXPORT_FOLDER_NAME);
  folder.createFile(blob);
  DriveApp.getFileById(tmp.getId()).setTrashed(true);
}
