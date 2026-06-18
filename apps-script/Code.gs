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
 */

var SHEET_NAME = 'log';

// 列の並び。フロントの payload のキーと対応させる。
var COLUMNS = [
  'ts', 'name', 'dist', 'elev', 'gradePct', 'legal', 'taikan', 'elapsed',
  'upMets', 'effMets', 'ratio', 'loadMult', 'kRT', 'month', 'cc',
  'wt', 'steps', 'signals', 'crossings', 'local', 'stroller', 'carry', 'speed'
];

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getSheet_();

    // シートが空ならヘッダ行を先に書く
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(COLUMNS);
    }

    // サーバー側で受信時刻を必ず付与（フロントの ts が無くても残る）
    if (data.ts === undefined || data.ts === null || data.ts === '') {
      data.ts = new Date().toISOString();
    }

    var row = COLUMNS.map(function (key) {
      var v = data[key];
      return v === undefined || v === null ? '' : v;
    });
    sheet.appendRow(row);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
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
