/**
 * 駅徒歩 体感指標 — 集計用 Google Apps Script
 *
 * 公開ページ(index.html)の「比較表に追加」で送られてくる算出データを、
 * このスクリプトを紐づけたスプレッドシートに1行ずつ追記して蓄積する。
 *
 * 使い方:
 *  1. 集計用のスプレッドシートを新規作成
 *  2. 拡張機能 → Apps Script を開き、このファイルの内容を貼り付けて保存
 *  3. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *       - 実行ユーザー: 自分
 *       - アクセスできるユーザー: 全員
 *  4. 発行された /exec URL を index.html の LOG_ENDPOINT に貼る
 *
 * フロントは mode:'no-cors' の fire-and-forget で送信するため、
 * レスポンス本文は読まれない。CORS ヘッダの設定は不要。
 *
 * 列はヘッダ名で対応付ける。ペイロードに新しいキーが増えたら、
 * 既存シートの右端にその列を自動追加する（並びがズレない）。
 */

var SHEET_NAME = 'log';

// 新規シート時の優先的な列の並び。これ以外のキーは右端へ自動追加される。
var PREFERRED_COLUMNS = [
  'ts', 'name', 'dist', 'elev', 'net', 'ascent', 'descent', 'reliefRT', 'gradePct',
  'legal', 'taikan', 'elapsed', 'avgMETs', 'upMets', 'effMets', 'ratio', 'loadMult',
  'kRT', 'month', 'cc', 'wt', 'steps', 'signals', 'crossings', 'local',
  'stroller', 'carry', 'speed'
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.tryLock(10000); // 同時書き込みでヘッダ/行がズレないように直列化
    var data = JSON.parse(e.postData.contents);
    if (data.ts === undefined || data.ts === null || data.ts === '') {
      data.ts = new Date().toISOString();
    }

    var sheet = getSheet_();
    var header = readHeader_(sheet);

    // 新規シートなら優先順を種にする
    if (header.length === 0) {
      header = PREFERRED_COLUMNS.slice();
    }

    // ペイロードに有るが見出しに無いキーを右端へ追加
    var changed = false;
    Object.keys(data).forEach(function (k) {
      if (header.indexOf(k) === -1) { header.push(k); changed = true; }
    });

    // 見出し行を書く（空 or 列が増えたとき）
    if (sheet.getLastRow() === 0 || changed) {
      sheet.getRange(1, 1, 1, header.length).setValues([header]);
    }

    // 見出しの並びに合わせて値行を作る
    var row = header.map(function (k) {
      var v = data[k];
      return v === undefined || v === null ? '' : v;
    });
    sheet.appendRow(row);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function readHeader_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].filter(function (h) {
    return h !== '' && h !== null;
  });
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  return sheet;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
