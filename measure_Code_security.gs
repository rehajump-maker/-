const SPREADSHEET_ID = '1sf4SDcT4qZFx_VplIHckmc1Ib4fBLErf8F-OrYVyshk';
const MEMBER_SHEET = '会員情報';
const RECORD_SHEET = '測定記録';
const SETTINGS_SHEET = '設定';
const MEMBER_HEADERS = ['id', 'memberId', 'name', 'sex', 'birth'];
const RECORD_HEADERS = ['id', 'memberId', 'date', 'valuesJson'];
const SETTINGS_HEADERS = ['key', 'value'];

/**
 * ============================================================
 * セキュリティ設定について
 * ============================================================
 * 初回のみ、関数「setupSecurity」を1回実行してください
 * （電子カルテと同じ「連携用シークレット」を設定します）。
 *
 * 通常時：電子カルテ経由（会員番号付きの一時的なリンク）でしか
 *         このシステムを開けません。
 *
 * 緊急時（電子カルテに不具合が起きている場合など）：
 *         スプレッドシートの「設定」シートを開き、
 *         SECURITY_ON の値を FALSE に書き換えるだけで、
 *         コードの変更・再デプロイなしにこの制限を一時解除できます。
 *         直ったら TRUE に戻してください。
 * ============================================================
 */
function setupSecurity() {
  PropertiesService.getScriptProperties().setProperty(
    "LINK_SECRET",
    "6ca10175-b568-4a9e-ba19-032eb8a768ffe7869e68bbb84d43b6496405dd92e7e2"
  );
  // 「設定」シートを作成し、初期値（セキュリティON）を入れる
  getSetting("SECURITY_ON", "TRUE");
  Logger.log("セキュリティ設定を初期化しました。LINK_SECRETを設定し、「設定」シートを作成しました。");
}

/* アプリ本体（Index.html）を配信する。
   iOS/iPadOSではメールやFilesで開いたローカルHTMLファイルはJavaScriptが
   動かないため、このように「実際にホストされたページ」として配信する
   必要がある。 */
function doGet(e) {
  const securityOn = getSetting("SECURITY_ON", "TRUE") !== "FALSE";
  if (securityOn) {
    const ok = verifyLinkToken(e.parameter.member, e.parameter.exp, e.parameter.sig);
    if (!ok) {
      return HtmlService.createHtmlOutput(
        '<!DOCTYPE html><html lang="ja"><head><meta charset="UTF-8">' +
        '<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
        '<title>アクセス制限</title></head><body style="font-family:sans-serif;' +
        'padding:60px 24px;text-align:center;color:#333;">' +
        '<h2>このページには直接アクセスできません</h2>' +
        '<p>電子カルテシステムの画面から「体力測定を開く」を選んで開いてください。</p>' +
        '</body></html>'
      ).setTitle("アクセス制限");
    }
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('体力測定記録カルテ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/* ---------- 連携トークンの検証 ---------- */
function verifyLinkToken(memberId, exp, sig) {
  if (!exp || !sig) return false;
  const secret = PropertiesService.getScriptProperties().getProperty("LINK_SECRET");
  if (!secret) return false;
  const mid = memberId || "";
  const expected = computeLinkSig(mid, exp, secret);
  if (expected !== sig) return false;
  if (Number(exp) < Math.floor(Date.now() / 1000)) return false;
  return true;
}
function computeLinkSig(memberId, exp, secret) {
  const raw = memberId + "|" + exp + "|" + secret;
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return digest.map(b => ((b < 0 ? b + 256 : b).toString(16)).padStart(2, "0")).join("");
}

/* ---------- 設定（ON/OFFなど）の読み書き ---------- */
function getSetting(key, defaultVal) {
  const sh = getSheet(SETTINGS_SHEET, SETTINGS_HEADERS);
  const rows = sh.getDataRange().getValues().slice(1);
  const row = rows.find(r => String(r[0]) === key);
  if (row) return String(row[1]);
  sh.appendRow([key, defaultVal]);
  return defaultVal;
}

/* ---------- クライアント（google.script.run）から呼ばれるAPI ---------- */

function apiGetAllData() {
  return { members: getMembers(), measurements: getMeasurements() };
}
function apiUpsertMember(member) {
  return withLock(function () { return upsertMember(member); });
}
function apiDeleteMember(id) {
  return withLock(function () { return deleteMember(id); });
}
function apiUpsertMeasurement(rec) {
  return withLock(function () { return upsertMeasurement(rec); });
}
function apiDeleteMeasurement(id) {
  return withLock(function () { return deleteMeasurement(id); });
}
function apiReplaceAll(members, measurements) {
  return withLock(function () { return replaceAll(members, measurements); });
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function getSheet(name, headers) {
  if (!SPREADSHEET_ID || SPREADSHEET_ID.indexOf('ここに') === 0) {
    throw new Error('SPREADSHEET_ID が設定されていません。コード冒頭の SPREADSHEET_ID にスプレッドシートのIDを入力し、保存→再デプロイしてください。');
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(headers);
  return sh;
}

function formatDate(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return v;
}

function safeParseJson(s) {
  try { return JSON.parse(s); } catch (e) { return {}; }
}

/* ---------- 読み取り ---------- */

function getMembers() {
  const sh = getSheet(MEMBER_SHEET, MEMBER_HEADERS);
  const rows = sh.getDataRange().getValues().slice(1);
  return rows
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      return { id: String(r[0]), memberId: String(r[1]), name: String(r[2]), sex: String(r[3]), birth: formatDate(r[4]) };
    });
}

function getMeasurements() {
  const sh = getSheet(RECORD_SHEET, RECORD_HEADERS);
  const rows = sh.getDataRange().getValues().slice(1);
  return rows
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      return { id: String(r[0]), memberId: String(r[1]), date: formatDate(r[2]), values: safeParseJson(r[3]) };
    });
}

/* ---------- 会員情報の書き込み ---------- */

function upsertMember(member) {
  const sh = getSheet(MEMBER_SHEET, MEMBER_HEADERS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === member.id) {
      sh.getRange(i + 1, 1, 1, MEMBER_HEADERS.length)
        .setValues([[member.id, member.memberId, member.name, member.sex, member.birth]]);
      return { ok: true };
    }
  }
  sh.appendRow([member.id, member.memberId, member.name, member.sex, member.birth]);
  return { ok: true };
}

function deleteMember(id) {
  const sh = getSheet(MEMBER_SHEET, MEMBER_HEADERS);
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === id) sh.deleteRow(i + 1);
  }
  // 紐づく測定記録も削除
  const rsh = getSheet(RECORD_SHEET, RECORD_HEADERS);
  const rdata = rsh.getDataRange().getValues();
  for (let i = rdata.length - 1; i >= 1; i--) {
    if (rdata[i][1] === id) rsh.deleteRow(i + 1);
  }
  return { ok: true };
}

/* ---------- 測定記録の書き込み ---------- */

function upsertMeasurement(rec) {
  const sh = getSheet(RECORD_SHEET, RECORD_HEADERS);
  const data = sh.getDataRange().getValues();
  const valuesJson = JSON.stringify(rec.values || {});
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === rec.id) {
      sh.getRange(i + 1, 1, 1, RECORD_HEADERS.length)
        .setValues([[rec.id, rec.memberId, rec.date, valuesJson]]);
      return { ok: true };
    }
  }
  sh.appendRow([rec.id, rec.memberId, rec.date, valuesJson]);
  return { ok: true };
}

function deleteMeasurement(id) {
  const sh = getSheet(RECORD_SHEET, RECORD_HEADERS);
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === id) sh.deleteRow(i + 1);
  }
  return { ok: true };
}

/* ---------- バックアップからの一括復元 ---------- */

function replaceAll(members, measurements) {
  const msh = getSheet(MEMBER_SHEET, MEMBER_HEADERS);
  msh.clearContents();
  msh.appendRow(MEMBER_HEADERS);
  (members || []).forEach(function (m) {
    msh.appendRow([m.id, m.memberId, m.name, m.sex, m.birth]);
  });

  const rsh = getSheet(RECORD_SHEET, RECORD_HEADERS);
  rsh.clearContents();
  rsh.appendRow(RECORD_HEADERS);
  (measurements || []).forEach(function (r) {
    rsh.appendRow([r.id, r.memberId, r.date, JSON.stringify(r.values || {})]);
  });
  return { ok: true };
}
