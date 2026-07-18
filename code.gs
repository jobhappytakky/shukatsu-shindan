/* ============================================================
   就活迷子タイプ診断｜GAS（設定配信＋ログ記録）
   使い方:
   1. 新規スプレッドシートを作成し、拡張機能→Apps Script にこのコードを貼る
   2. エディタ上部で関数「setup」を選んで実行（初回は権限承認）
      → 「リンク設定」「カード画像」「ログ」の3シートが自動作成される
   3. デプロイ→新しいデプロイ→種類「ウェブアプリ」
      実行ユーザー: 自分／アクセスできるユーザー: 全員
   4. 発行されたWebアプリURLを index.html / demo.html の GAS_URL に貼る
============================================================ */

const SHEET_LINK = 'リンク設定';
const SHEET_CARD = 'カード画像';
const SHEET_LOG  = 'ログ';

/* ---------- 初期セットアップ ---------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let link = ss.getSheetByName(SHEET_LINK);
  if (!link) {
    link = ss.insertSheet(SHEET_LINK);
    link.getRange(1, 1, 1, 5).setValues([['タイプキー', 'ボタン文言', 'リンク先URL', 'サムネ画像URL', '表示ON-OFF']]);
    const rows = [
      ['ikisaki', 'やりたいことをみてもらう',   'https://coconala.com/services/4191988', '', 'ON'],
      ['tonari',  '自分のペースをみてもらう',   'https://coconala.com/services/4191988', '', 'ON'],
      ['jishin',  '自分の強みをみてもらう',     'https://coconala.com/services/4191988', '', 'ON'],
      ['chizu',   '選ぶ軸をみてもらう',         'https://coconala.com/services/4191988', '', 'ON'],
      ['namae',   'この迷いの正体をみてもらう', 'https://coconala.com/services/4191988', '', 'ON'],
      ['zenbu',   '迷いをまるごとみてもらう',   'https://coconala.com/services/4191988', '', 'ON'],
      ['sekai',   '', '', '', 'OFF']
    ];
    link.getRange(2, 1, rows.length, 5).setValues(rows);
  }

  let card = ss.getSheetByName(SHEET_CARD);
  if (!card) {
    card = ss.insertSheet(SHEET_CARD);
    card.getRange(1, 1, 1, 2).setValues([['カード名', '画像URL（空ならリポジトリ同梱のライダー版を使用）']]);
    const names = ['愚者','魔術師','女教皇','女帝','皇帝','教皇','恋人','戦車','力','隠者','運命の輪','正義','吊るされた男','死神','節制','悪魔','塔','星','月','太陽','審判','世界'];
    card.getRange(2, 1, names.length, 1).setValues(names.map(n => [n]));
  }

  let log = ss.getSheetByName(SHEET_LOG);
  if (!log) {
    log = ss.insertSheet(SHEET_LOG);
    log.getRange(1, 1, 1, 5).setValues([['日時', 'イベント', '主タイプ', '副タイプ', 'カード']]);
  }
}

/* ---------- 設定配信（GET ?action=config） ---------- */
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'config') {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const out = { cta: {}, cards: {} };

    const link = ss.getSheetByName(SHEET_LINK);
    if (link && link.getLastRow() > 1) {
      const rows = link.getRange(2, 1, link.getLastRow() - 1, 5).getValues();
      rows.forEach(r => {
        const key = String(r[0]).trim();
        if (!key) return;
        out.cta[key] = {
          label: String(r[1] || ''),
          url: String(r[2] || ''),
          thumb: String(r[3] || ''),
          show: String(r[4]).trim().toUpperCase() !== 'OFF'
        };
      });
    }

    const card = ss.getSheetByName(SHEET_CARD);
    if (card && card.getLastRow() > 1) {
      const rows = card.getRange(2, 1, card.getLastRow() - 1, 2).getValues();
      rows.forEach(r => {
        const name = String(r[0]).trim();
        const url = String(r[1] || '').trim();
        if (name && url) out.cards[name] = url;
      });
    }

    return ContentService.createTextOutput(JSON.stringify(out))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput('ok');
}

/* ---------- ログ記録（POST） ---------- */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const log = ss.getSheetByName(SHEET_LOG);
    log.appendRow([
      new Date(),
      sanitize(data.event),
      sanitize(data.main),
      sanitize(data.sub),
      sanitize(data.card)
    ]);
  } catch (err) {}
  return ContentService.createTextOutput('ok');
}

/* 数式インジェクション対策＋50字制限 */
function sanitize(v) {
  let s = String(v == null ? '' : v).slice(0, 50);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}
