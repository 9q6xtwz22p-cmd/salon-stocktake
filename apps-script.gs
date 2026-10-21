/**
 * 棚卸しツールの保存先。Googleスプレッドシートを置き場にして、
 * ログイン不要で読み書きできるようにする。
 *
 * なぜこれか:
 *   claude.ai の Artifact はサインインした人にしか開かない。
 *   店舗のスタッフにアカウントを作らせずに同じ数字を見せたいので、
 *   HTMLは GitHub Pages、データはこのスプレッドシートに置く。
 *
 * 置き方（スプレッドシートを開いて 拡張機能 → Apps Script に貼る）:
 *   1. このファイルの中身を全部貼る
 *   2. KEY を自分で決めた合言葉に書き換える（空のままだと誰でも書ける）
 *   3. 上の「デプロイ」→「新しいデプロイ」→ 種類＝ウェブアプリ
 *        次のユーザーとして実行: 自分
 *        アクセスできるユーザー: **全員**  ← ここを変えないとログインを求められる
 *   4. 出てきた /exec で終わるURLを控える。これをページに入れる
 *   5. シートを作り直したくなったら setup() を一度だけ実行する
 *
 * **CORSの都合で、書き込みは Content-Type: text/plain で送る。**
 *   application/json にすると preflight(OPTIONS) が飛び、Apps Script はそれに答えない。
 */

var KEY = '';          // 合言葉。空なら誰でも書ける。必ず変えること
var TZ  = 'Asia/Tokyo';

// **文字として入れる列。** これを指定しないとスプレッドシートが勝手に変換する。
//   「2026-08」→ 日付、「6%」→ 0.06 になり、月のキーも別名も壊れる（実際に壊れた）。
var TEXT_COLS = { '店舗': [1, 2], '品目': [1, 2, 3, 4, 7], '棚卸': [1, 2, 3, 4], '月次': [1, 2, 7, 8],
  'アクセス': [1, 2, 3, 4], 'アンケート': [1, 2, 3, 4, 5, 6] };

var SHEETS = {
  店舗: ['id', '名前'],
  品目: ['店舗', 'id', '品目名', 'カテゴリ', '単価(税込)', '初期在庫', '別名'],
  棚卸: ['店舗', '月', '品目id', '品目名', '仕入', '月末在庫'],
  月次: ['店舗', '月', '売上(税込)', '材料費(税込)', '材料費率', '棚卸額(税込)', '記録者', '更新日時'],
  // **管理者だけが見る。** ページには出さない（スプレッドシートを開けるのは管理者だけ）。
  // 誰がいつ使ったかではなく、**棚卸しが終わったかを確かめるため**のもの。
  // 1人が何度も開くので、同じ人・同じ日・同じ操作は1行にまとめる（毎回足すと読めなくなる）。
  アクセス: ['日時', '名前', '店舗', '操作', '回数'],
  // アンケート。**1人1行で上書きする**（出し直せるように）。
  // 匿名にしない——自主の係は本人が引き受けるのが狙いで、
  // 匿名だと希望は集まるが責任が生まれない（2026-10-23 MTG）。
  アンケート: ['日時', '店舗', '名前', '設問', '希望', '自分で決めた分量']
};

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(function (name) {
    // **アクセス履歴は消さない。** setup() は作り直し用だが、履歴まで消すと記録が飛ぶ
    if (name === 'アクセス' && ss.getSheetByName(name)) return;
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.clear();
    sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  return 'シートを作りました: ' + Object.keys(SHEETS).join(' / ');
}

/* ---- 共通 ---- */
function sheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function rows_(name) {
  var sh = sheet_(name), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, SHEETS[name].length).getValues();
}
function write_(name, body) {
  var sh = sheet_(name), w = SHEETS[name].length;
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, w).clearContent();
  if (!body.length) return;
  // 値を入れる前に書式を文字にする。入れてからでは手遅れ
  (TEXT_COLS[name] || []).forEach(function (c) {
    sh.getRange(2, c, body.length, 1).setNumberFormat('@');
  });
  sh.getRange(2, 1, body.length, w).setValues(body);
}

function ym_(v) {
  // 既に日付として入っている古い行も読めるようにする
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM');
  return String(v || '');
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
function ok_(key) { return !KEY || String(key || '') === KEY; }
function num_(v) { return (v === '' || v === null || v === undefined) ? null : Number(v); }

/* ---- アンケート ---- */
function saveSurvey_(b) {
  var who = String(b.who || '').trim();
  if (!who) return { ok: false, error: 'name' };     // 記名でないと受け取らない
  var store = String(b.store || '');
  var sh = sheet_('アンケート');
  var stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');

  // **同じ人の前の回答は消してから入れ直す。** 出し直せないと、書き間違えたまま残る
  var last = sh.getLastRow();
  if (last > 1) {
    var rows = sh.getRange(2, 1, last - 1, SHEETS['アンケート'].length).getValues();
    for (var i = rows.length - 1; i >= 0; i--) {
      if (String(rows[i][2]) === who && String(rows[i][1]) === store) sh.deleteRow(i + 2);
    }
  }
  var body = (b.answers || []).map(function (a) {
    return [stamp, store, who, String(a.q || ''), String(a.want || ''), String(a.how || '')];
  });
  if (!body.length) return { ok: true, at: stamp, n: 0 };
  var from = sh.getLastRow() + 1;
  sh.getRange(from, 1, body.length, 6).setNumberFormat('@');
  sh.getRange(from, 1, body.length, 6).setValues(body);
  log_(who, store, 'アンケートを出した');
  return { ok: true, at: stamp, n: body.length };
}

/* ---- アクセス履歴（管理者だけが見る） ---- */
function log_(who, store, op) {
  who = String(who || '').trim();
  if (!who) return;                       // 名前を入れる前の読み込みは数えない
  try {
    var sh = sheet_('アクセス');
    var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    var last = sh.getLastRow();
    // 同じ人・同じ日・同じ操作なら、最終行を更新して回数を足す（行が増え続けないように）
    if (last > 1) {
      var r = sh.getRange(last, 1, 1, 5).getValues()[0];
      if (String(r[1]) === who && String(r[2]) === String(store || '') &&
          String(r[3]) === op && String(r[0]).slice(0, 10) === today) {
        sh.getRange(last, 1).setValue(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'));
        sh.getRange(last, 5).setValue(Number(r[4] || 1) + 1);
        return;
      }
    }
    sh.getRange(last + 1, 1, 1, 5).setNumberFormat('@');
    sh.getRange(last + 1, 1, 1, 5).setValues([[
      Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'),
      who, String(store || ''), op, 1]]);
  } catch (err) {
    // **記録に失敗しても本体は止めない。** 棚卸しが入らなくなる方が困る
  }
}

/* ---- 読み ---- */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!ok_(p.key)) return json_({ ok: false, error: 'key' });
  log_(p.who, p.store, '開いた');

  var stores = rows_('店舗').filter(function (r) { return r[0] !== ''; })
    .map(function (r) { return { id: String(r[0]), name: String(r[1] || r[0]) }; });

  var items = {};
  rows_('品目').forEach(function (r) {
    if (r[0] === '' || r[1] === '') return;
    var st = String(r[0]);
    items[st] = items[st] || { items: [], cats: [] };
    var cat = String(r[3] || 'その他');
    if (items[st].cats.indexOf(cat) < 0) items[st].cats.push(cat);
    items[st].items.push({
      id: String(r[1]), name: String(r[2]), cat: cat,
      price: num_(r[4]), opening: Number(r[5] || 0),
      alias: String(r[6] || '').split(/[,、]/).map(function (s) { return s.trim(); })
        .filter(function (s) { return s; })
    });
  });

  var months = {};
  rows_('棚卸').forEach(function (r) {
    if (r[0] === '' || r[1] === '' || r[2] === '') return;
    var k = String(r[0]) + '__' + ym_(r[1]);
    months[k] = months[k] || { entries: {} };
    var e2 = {};
    if (num_(r[4]) !== null) e2['in'] = num_(r[4]);
    if (num_(r[5]) !== null) e2.end = num_(r[5]);
    months[k].entries[String(r[2])] = e2;
  });
  rows_('月次').forEach(function (r) {
    if (r[0] === '' || r[1] === '') return;
    var k = String(r[0]) + '__' + ym_(r[1]);
    months[k] = months[k] || { entries: {} };
    months[k].sales = Number(r[2] || 0);
    months[k].who = String(r[6] || '');
    months[k].at = r[7] ? String(r[7]) : null;
  });

  return json_({ ok: true, stores: stores, items: items, months: months });
}

/* ---- 書き ---- */
function doPost(e) {
  var body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { }
  if (!ok_(body.key)) return json_({ ok: false, error: 'key' });

  if (body.op === 'survey') return json_(saveSurvey_(body));

  var OP = { month: '棚卸しを保存', items: '品目を編集', stores: '店舗を編集' };
  if (OP[body.op]) log_(body.who, body.store, OP[body.op]);

  if (body.op === 'month') return json_(saveMonth_(body));
  if (body.op === 'items') return json_(saveItems_(body));
  if (body.op === 'stores') return json_(saveStores_(body));
  return json_({ ok: false, error: 'op' });
}

function saveMonth_(b) {
  var store = String(b.store), ym = String(b.month);
  var names = {};
  rows_('品目').forEach(function (r) { if (String(r[0]) === store) names[String(r[1])] = String(r[2]); });

  var keep = rows_('棚卸').filter(function (r) {
    return !(String(r[0]) === store && ym_(r[1]) === ym);
  });
  var add = [];
  Object.keys(b.entries || {}).forEach(function (id) {
    var v = b.entries[id] || {};
    if (v['in'] === undefined && v.end === undefined) return;
    add.push([store, ym, id, names[id] || '', v['in'] === undefined ? '' : v['in'],
              v.end === undefined ? '' : v.end]);
  });
  keep.forEach(function (r) { r[1] = ym_(r[1]); });
  write_('棚卸', keep.concat(add));

  var stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
  var m = rows_('月次').filter(function (r) {
    return !(String(r[0]) === store && ym_(r[1]) === ym);
  });
  m.forEach(function (r) { r[1] = ym_(r[1]); });
  m.push([store, ym, Number(b.sales || 0), Number(b.cost || 0),
          b.rate === null || b.rate === undefined ? '' : Number(b.rate),
          Number(b.stockValue || 0), String(b.who || ''), stamp]);
  write_('月次', m);
  return { ok: true, at: stamp };
}

function saveItems_(b) {
  var store = String(b.store);
  var keep = rows_('品目').filter(function (r) { return String(r[0]) !== store; });
  var add = (b.items || []).map(function (it) {
    return [store, String(it.id), String(it.name), String(it.cat || 'その他'),
            it.price === null || it.price === undefined ? '' : Number(it.price),
            Number(it.opening || 0), (it.alias || []).join(',')];
  });
  write_('品目', keep.concat(add));
  return { ok: true };
}

function saveStores_(b) {
  write_('店舗', (b.stores || []).map(function (s) { return [String(s.id), String(s.name)]; }));
  return { ok: true };
}
