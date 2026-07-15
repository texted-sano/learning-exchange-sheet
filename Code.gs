/**
 * 学習交流シート（プロトタイプ）バックエンド
 * ------------------------------------------------------
 * このスクリプトは、Googleスプレッドシートに「コンテナバインド」して使います。
 * 教科書アプリ側（フロントエンド）は、このスクリプトを「ウェブアプリ」として
 * デプロイしたURL（exec URL）に対してAPIリクエストを送り、データの読み書きを行います。
 *
 * できること（コア機能＋最低限の本人確認）
 *  - 生徒が、公開された授業に対して「ふりかえり」を入力・保存する
 *  - 教員が、生徒の入力内容を一覧で確認し、コメント・10段階評価・いいねをつけて保存する
 *  - 生徒が、後日「きろくを見る」で自分の過去の入力＋先生のコメント等を見返す
 *  - 生徒は名簿のパスワード列が空のときだけ初回パスワード登録ができ、
 *    以降は名前＋パスワードでログインする（教師ログインと同じ考え方）
 *
 * 使い方（概要）
 *  1. 新しいGoogleスプレッドシートを作成する
 *  2. 拡張機能 > Apps Script を開き、このファイルの内容を貼り付けて保存する
 *  3. スプレッドシートをリロードすると「学習交流シート」メニューが出るので
 *     「シートを初期化」を実行する（設定/名簿/授業/回答シートが作られる）
 *  4. 「設定」シートにクラス名・教師パスワードなどを入力する
 *  5. 「名簿」シートに生徒の名前を入力する（パスワード列は空のままでよい＝生徒が自分で初回登録する）
 *  6. Apps Scriptエディタの「デプロイ」>「新しいデプロイ」>
 *     種類「ウェブアプリ」を選び、
 *       - 実行するユーザー: 自分
 *       - アクセスできるユーザー: 全員
 *     でデプロイする
 *  7. 発行されたウェブアプリURL（.../exec）を、index.html の
 *     「接続先URL」欄に入力する
 *
 * 注意：コードを書き換えたあとは、毎回「デプロイを管理」>「編集」>
 *      「新しいバージョン」でデプロイし直さないと変更が反映されません。
 *
 * セキュリティに関する注記：
 *  - パスワードはスプレッドシートに平文で保存されます（試作のための簡略化）。
 *  - 生徒の初回パスワード登録は「最初に名乗った人がそのパスワードを設定できる」方式です。
 *    本運用前には、教師が初期パスワードを配布する方式への変更も検討してください。
 */

const SHEET_SETTINGS = '設定';
const SHEET_ROSTER = '名簿';
const SHEET_LESSONS = '授業';
const SHEET_RESPONSES = '回答';
const SHEET_THEMES = 'テーマ';
const SETTING_STUDENT_TOKEN = '生徒用トークン';
const DEFAULT_THEMES = ['平和', '命', '多様性', '災害', '家族'];

// ============================================================
// 初期化（スプレッドシートのメニューから実行する）
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('学習交流シート')
    .addItem('シートを初期化', 'initializeSheets')
    .addToUi();
}

function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let settings = ss.getSheetByName(SHEET_SETTINGS);
  if (!settings) {
    settings = ss.insertSheet(SHEET_SETTINGS);
    settings.getRange('A1:B6').setValues([
      ['設定項目', '値'],
      ['クラス名', '6年1組'],
      ['担任名', ''],
      ['教師パスワード', ''],
      ['表示方法(本名/番号/匿名)', '本名'],
      [SETTING_STUDENT_TOKEN, ''],
    ]);
    settings.getRange('A1:B1').setFontWeight('bold');
    settings.setColumnWidth(1, 220);
    settings.setColumnWidth(2, 200);
  }

  let roster = ss.getSheetByName(SHEET_ROSTER);
  if (!roster) {
    roster = ss.insertSheet(SHEET_ROSTER);
    roster.getRange('A1:C1').setValues([['出席番号', '生徒名', 'パスワード']]);
    roster.getRange('A1:C1').setFontWeight('bold');
    roster.setColumnWidth(1, 100);
    roster.setColumnWidth(2, 160);
    roster.setColumnWidth(3, 160);
  }

  let lessons = ss.getSheetByName(SHEET_LESSONS);
  if (!lessons) {
    lessons = ss.insertSheet(SHEET_LESSONS);
    lessons.getRange('A1:H1').setValues([['授業ID', '日付', 'タイトル', '発問', '公開', 'テーマ', 'お気に入り', '削除済み']]);
    lessons.getRange('A1:H1').setFontWeight('bold');
    lessons.setColumnWidths(1, 8, 160);
  }

  let themes = ss.getSheetByName(SHEET_THEMES);
  if (!themes) {
    themes = ss.insertSheet(SHEET_THEMES);
    themes.getRange('A1').setValue('テーマ');
    themes.getRange('A1').setFontWeight('bold');
    themes.getRange(2, 1, DEFAULT_THEMES.length, 1).setValues(DEFAULT_THEMES.map((t) => [t]));
    themes.setColumnWidth(1, 160);
  }

  let responses = ss.getSheetByName(SHEET_RESPONSES);
  if (!responses) {
    responses = ss.insertSheet(SHEET_RESPONSES);
    responses.getRange('A1:I1').setValues([
      ['回答ID', '授業ID', '生徒名', '回答内容', '提出日時', '教師コメント', 'コメント日時', '評価', 'いいね'],
    ]);
    responses.getRange('A1:I1').setFontWeight('bold');
    responses.setColumnWidths(1, 9, 150);
  }

  SpreadsheetApp.getUi().alert(
    '初期化が完了しました。\n「設定」シートと「名簿」シートに内容を入力してください。\n' +
    '（名簿のパスワード列は空のままでOKです。生徒が初回ログイン時に自分で設定します）'
  );
}

// ============================================================
// リクエストの受け口
// ============================================================

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  let params = {};
  try {
    if (e.postData && e.postData.contents) {
      // フロントエンドは text/plain で JSON 文字列を送ってくる
      // （プリフライトを発生させないための工夫）
      params = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      params = e.parameter;
    }

    const action = params.action;
    let result;

    switch (action) {
      case 'ping':
        result = actionPing();
        break;
      case 'getRoster':
        result = actionGetRoster(params);
        break;
      case 'checkStudentAccount':
        result = actionCheckStudentAccount(params);
        break;
      case 'registerStudentPassword':
        result = actionRegisterStudentPassword(params);
        break;
      case 'studentLogin':
        result = actionStudentLogin(params);
        break;
      case 'teacherLogin':
        result = actionTeacherLogin(params);
        break;
      case 'registerTeacherPassword':
        result = actionRegisterTeacherPassword(params);
        break;
      case 'updateTeacherSettings':
        result = actionUpdateTeacherSettings(params);
        break;
      case 'getStudentToken':
        result = actionGetStudentToken(params);
        break;
      case 'getStudentTokenStatus':
        result = actionGetStudentTokenStatus(params);
        break;
      case 'getLessons':
        result = actionGetLessons(params);
        break;
      case 'createLesson':
        result = actionCreateLesson(params);
        break;
      case 'setLessonPublished':
        result = actionSetLessonPublished(params);
        break;
      case 'setLessonTheme':
        result = actionSetLessonTheme(params);
        break;
      case 'setLessonFavorite':
        result = actionSetLessonFavorite(params);
        break;
      case 'deleteLesson':
        result = actionDeleteLesson(params);
        break;
      case 'restoreLesson':
        result = actionRestoreLesson(params);
        break;
      case 'purgeLesson':
        result = actionPurgeLesson(params);
        break;
      case 'getDeletedLessons':
        result = actionGetDeletedLessons(params);
        break;
      case 'getThemes':
        result = actionGetThemes(params);
        break;
      case 'addTheme':
        result = actionAddTheme(params);
        break;
      case 'removeTheme':
        result = actionRemoveTheme(params);
        break;
      case 'getAllResponsesForStudent':
        result = actionGetAllResponsesForStudent(params);
        break;
      case 'getThemeEvaluationSummary':
        result = actionGetThemeEvaluationSummary(params);
        break;
      case 'submitReflection':
        result = actionSubmitReflection(params);
        break;
      case 'getResponsesForLesson':
        result = actionGetResponsesForLesson(params);
        break;
      case 'saveComment':
        result = actionSaveComment(params);
        break;
      case 'getMyRecords':
        result = actionGetMyRecords(params);
        break;
      case 'getRosterAdmin':
        result = actionGetRosterAdmin(params);
        break;
      case 'addStudent':
        result = actionAddStudent(params);
        break;
      case 'removeStudent':
        result = actionRemoveStudent(params);
        break;
      case 'resetStudentPassword':
        result = actionResetStudentPassword(params);
        break;
      case 'setStudentPassword':
        result = actionSetStudentPassword(params);
        break;
      default:
        throw new Error('不明な action です: ' + action);
    }

    return jsonOutput(Object.assign({ ok: true }, result));
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err.message || err) });
  }
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}

// ============================================================
// シート読み書きの共通ヘルパー
// ============================================================

function getSheet(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) {
    throw new Error(
      'シート「' + name + '」が見つかりません。スプレッドシートのメニューから' +
      '「学習交流シート > シートを初期化」を実行してください。'
    );
  }
  return sheet;
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data
    .slice(1)
    .filter((row) => row.some((cell) => cell !== '' && cell !== null))
    .map((row) => {
      const obj = {};
      headers.forEach((h, i) => (obj[h] = row[i]));
      return obj;
    });
}

function getSettingsMap() {
  const sheet = getSheet(SHEET_SETTINGS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const map = {};
  data.forEach((row) => {
    if (row[0]) map[row[0]] = row[1];
  });
  return map;
}

// ----- ごく簡易なレート制限（総当たり攻撃対策） -----
// 注：Apps Scriptの doGet/doPost からは呼び出し元のIPアドレスが取得できないため、
// 教師ログインは「アプリ全体」、生徒ログイン／登録は「名前ごと」にカウントする簡易版です。
// ロックされても既存データの破壊・流出は起きません（新たな照合・登録だけを一時停止します）。

function checkRateLimit(bucket) {
  const cache = CacheService.getScriptCache();
  if (cache.get(bucket + '_lock')) {
    throw new Error(
      '試行回数が多いため、しばらく時間（数分程度）をおいてからやり直してください。'
    );
  }
}

function recordFailedAttempt(bucket, maxAttempts, lockSeconds) {
  const cache = CacheService.getScriptCache();
  const countKey = bucket + '_count';
  const count = Number(cache.get(countKey) || '0') + 1;
  if (count >= maxAttempts) {
    cache.put(bucket + '_lock', '1', lockSeconds);
    cache.remove(countKey);
  } else {
    cache.put(countKey, String(count), lockSeconds);
  }
}

function clearRateLimit(bucket) {
  const cache = CacheService.getScriptCache();
  cache.remove(bucket + '_count');
  cache.remove(bucket + '_lock');
}

function checkTeacherPassword(params) {
  const BUCKET = 'teacherLogin';
  const MAX_ATTEMPTS = 5;
  const LOCK_SECONDS = 300; // 5分

  checkRateLimit(BUCKET);

  const settings = getSettingsMap();
  const correct = String(settings['教師パスワード'] || '');
  if (!correct) {
    throw new Error('「設定」シートに教師パスワードが設定されていません');
  }
  if (!params.password || String(params.password) !== correct) {
    recordFailedAttempt(BUCKET, MAX_ATTEMPTS, LOCK_SECONDS);
    throw new Error('教師パスワードが正しくありません');
  }
  clearRateLimit(BUCKET);
  return settings;
}

// 生徒用トークンの照合（生徒系アクション共通）
function checkStudentToken(params) {
  const settings = getSettingsMap();
  const token = String(settings[SETTING_STUDENT_TOKEN] || '');
  if (!token || String(params.token || '') !== token) {
    throw new Error('生徒用URLが無効です。教師に確認してください。');
  }
}

function setSettingValue(key, value) {
  const sheet = getSheet(SHEET_SETTINGS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

// 複数の設定項目をまとめて書き込む（シートの読み込みを1回にまとめて高速化する）
function setSettingValues(updates) {
  const sheet = getSheet(SHEET_SETTINGS);
  const data = sheet.getDataRange().getValues();
  Object.keys(updates).forEach((key) => {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === key) {
        sheet.getRange(i + 1, 2).setValue(updates[key]);
        return;
      }
    }
    sheet.appendRow([key, updates[key]]);
  });
}

// テーマシートが無い（旧スプレッドシート）場合は既定のテーマで自動作成する
function ensureThemeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_THEMES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_THEMES);
    sheet.getRange('A1').setValue('テーマ');
    sheet.getRange(2, 1, DEFAULT_THEMES.length, 1).setValues(DEFAULT_THEMES.map((t) => [t]));
  }
  return sheet;
}

function getThemeList() {
  const sheet = ensureThemeSheet();
  const data = sheet.getDataRange().getValues();
  return data
    .slice(1)
    .map((row) => String(row[0]).trim())
    .filter((t) => !!t);
}

// 授業シートに「テーマ」列が無い（旧スプレッドシート）場合は自動で追加する
function ensureLessonThemeColumn() {
  const sheet = getSheet(SHEET_LESSONS);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (header.indexOf('テーマ') === -1) {
    sheet.getRange(1, lastCol + 1).setValue('テーマ');
  }
}

// 授業シートに「お気に入り」列が無い（旧スプレッドシート）場合は自動で追加する
function ensureLessonFavoriteColumn() {
  const sheet = getSheet(SHEET_LESSONS);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (header.indexOf('お気に入り') === -1) {
    sheet.getRange(1, lastCol + 1).setValue('お気に入り');
  }
}

// 授業シートに「削除済み」列が無い（旧スプレッドシート）場合は自動で追加する
function ensureLessonDeletedColumn() {
  const sheet = getSheet(SHEET_LESSONS);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (header.indexOf('削除済み') === -1) {
    sheet.getRange(1, lastCol + 1).setValue('削除済み');
  }
}

function formatDate(d) {
  if (Object.prototype.toString.call(d) === '[object Date]') {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return d ? String(d) : '';
}

function formatDateTime(d) {
  if (Object.prototype.toString.call(d) === '[object Date]') {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  }
  return d ? String(d) : '';
}

function isTrue(v) {
  return v === true || v === 'true' || v === 'TRUE' || v === '公開';
}

// 生徒の識別子（生徒名があれば生徒名、無ければ「出席番号＋番」）
function getStudentIdentifier(row) {
  const name = row['生徒名'] && String(row['生徒名']).trim();
  if (name) return name;
  const number = row['出席番号'] && String(row['出席番号']).trim();
  return number ? number + '番' : '';
}

function findRosterRowIndex(data, headers, identifier) {
  const nameCol = headers.indexOf('生徒名');
  const numberCol = headers.indexOf('出席番号');
  for (let i = 1; i < data.length; i++) {
    const row = { '生徒名': data[i][nameCol], '出席番号': data[i][numberCol] };
    if (getStudentIdentifier(row) === identifier) return i;
  }
  return -1;
}

// ============================================================
// 各アクション
// ============================================================

// 接続確認＋クラス名の取得（ログイン画面に表示する）
function actionPing() {
  const settings = getSettingsMap();
  return {
    className: settings['クラス名'] || '',
    teacherName: settings['担任名'] || '',
    hasTeacherPassword: !!(settings['教師パスワード'] && String(settings['教師パスワード']).trim()),
  };
}

// 生徒ログイン画面の名前選択肢
function actionGetRoster(params) {
  checkStudentToken(params);
  const objs = sheetToObjects(getSheet(SHEET_ROSTER));
  const roster = objs.map((o) => getStudentIdentifier(o)).filter((n) => !!n);
  return { roster: roster };
}

// 選んだ生徒がパスワード登録済みかどうかを確認する
function actionCheckStudentAccount(params) {
  checkStudentToken(params);
  const name = params.name && String(params.name).trim();
  if (!name) throw new Error('名前を選んでください');

  const objs = sheetToObjects(getSheet(SHEET_ROSTER));
  const match = objs.find((o) => getStudentIdentifier(o) === name);
  if (!match) throw new Error('名簿に見つかりません');

  const pw = match['パスワード'];
  return { hasPassword: !!(pw && String(pw).trim()) };
}

// 初回パスワード登録（パスワード列が空のときだけ成功する）
function actionRegisterStudentPassword(params) {
  checkStudentToken(params);
  const name = params.name && String(params.name).trim();
  if (!name) throw new Error('名前を選んでください');
  const password = params.password;
  if (!password || String(password).length < 4) {
    throw new Error('パスワードは4文字以上で設定してください');
  }

  const bucket = 'studentRegister_' + name;
  checkRateLimit(bucket);

  const sheet = getSheet(SHEET_ROSTER);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const pwCol = headers.indexOf('パスワード');
  if (pwCol === -1) throw new Error('名簿シートに「パスワード」列がありません');

  const rowIndex = findRosterRowIndex(data, headers, name);
  if (rowIndex === -1) {
    recordFailedAttempt(bucket, 10, 600);
    throw new Error('名簿に見つかりません');
  }

  const existing = data[rowIndex][pwCol];
  if (existing && String(existing).trim()) {
    throw new Error('すでにパスワードが設定されています。ログインからお試しください。');
  }

  sheet.getRange(rowIndex + 1, pwCol + 1).setValue(String(password));
  clearRateLimit(bucket);
  const identifier = getStudentIdentifier({
    '生徒名': data[rowIndex][headers.indexOf('生徒名')],
    '出席番号': data[rowIndex][headers.indexOf('出席番号')],
  });
  return {
    studentName: identifier,
    lessons: getLessonsList('student'),
    records: getMyRecordsList(identifier),
  };
}

// 生徒ログイン（名前＋パスワード）
function actionStudentLogin(params) {
  checkStudentToken(params);
  const name = params.name && String(params.name).trim();
  if (!name) throw new Error('名前を選んでください');
  if (!params.password) throw new Error('パスワードを入力してください');

  const bucket = 'studentLogin_' + name;
  const MAX_ATTEMPTS = 6;
  const LOCK_SECONDS = 300; // 5分

  checkRateLimit(bucket);

  const objs = sheetToObjects(getSheet(SHEET_ROSTER));
  const match = objs.find((o) => getStudentIdentifier(o) === name);
  if (!match) throw new Error('名簿に見つかりません');

  const storedPw = match['パスワード'];
  if (!storedPw || !String(storedPw).trim()) {
    throw new Error('まだパスワードが設定されていません。先にパスワード登録を行ってください。');
  }
  if (String(params.password) !== String(storedPw)) {
    recordFailedAttempt(bucket, MAX_ATTEMPTS, LOCK_SECONDS);
    throw new Error('パスワードが正しくありません');
  }

  clearRateLimit(bucket);
  const identifier = getStudentIdentifier(match);
  return {
    studentName: identifier,
    lessons: getLessonsList('student'),
    records: getMyRecordsList(identifier),
  };
}

// 教師ログイン（パスワード確認のみ。セッションは持たず、毎回パスワードを送る方式）
function actionTeacherLogin(params) {
  const settings = checkTeacherPassword(params);
  return {
    className: settings['クラス名'] || '',
    teacherName: settings['担任名'] || '',
    lessons: getLessonsList('teacher'),
    themes: getThemeList(),
  };
}

// 教師の初回パスワード登録（パスワードが未設定のときだけ成功する）
function actionRegisterTeacherPassword(params) {
  const password = params.password;
  if (!password || String(password).length < 4) {
    throw new Error('パスワードは4文字以上で設定してください');
  }
  const settings = getSettingsMap();
  if (settings['教師パスワード'] && String(settings['教師パスワード']).trim()) {
    throw new Error('すでにパスワードが設定されています。ログインからお試しください。');
  }
  setSettingValue('教師パスワード', String(password));
  return {
    className: settings['クラス名'] || '',
    teacherName: settings['担任名'] || '',
    lessons: getLessonsList('teacher'),
    themes: getThemeList(),
  };
}

// 教師の設定（名前・クラス名・パスワード）を更新する
function actionUpdateTeacherSettings(params) {
  const settings = checkTeacherPassword(params);
  const updates = {};
  if (params.teacherName !== undefined) updates['担任名'] = String(params.teacherName);
  if (params.className !== undefined) updates['クラス名'] = String(params.className);
  if (params.newPassword) {
    if (String(params.newPassword).length < 4) {
      throw new Error('パスワードは4文字以上で設定してください');
    }
    updates['教師パスワード'] = String(params.newPassword);
  }
  if (Object.keys(updates).length > 0) {
    setSettingValues(updates);
  }
  const merged = Object.assign({}, settings, updates);
  return { className: merged['クラス名'] || '', teacherName: merged['担任名'] || '' };
}

// 授業一覧の中身を組み立てる（teacher は全件、それ以外は公開済みのみ）
function getLessonsList(role) {
  ensureLessonThemeColumn();
  ensureLessonFavoriteColumn();
  ensureLessonDeletedColumn();
  const sheet = getSheet(SHEET_LESSONS);
  const objs = sheetToObjects(sheet);
  let lessons = objs
    .filter((o) => !isTrue(o['削除済み']))
    .map((o) => ({
      lessonId: String(o['授業ID']),
      date: formatDate(o['日付']),
      title: o['タイトル'] || '',
      question: o['発問'] || '',
      published: isTrue(o['公開']),
      theme: o['テーマ'] || '',
      favorite: isTrue(o['お気に入り']),
    }));
  if (role !== 'teacher') {
    lessons = lessons.filter((l) => l.published);
  }
  lessons.sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0));
  return lessons;
}

// 授業一覧（teacher は全件、student は公開済みのみ）
function actionGetLessons(params) {
  if (params.role === 'teacher') {
    checkTeacherPassword(params);
  } else {
    checkStudentToken(params);
  }
  return { lessons: getLessonsList(params.role) };
}

// 授業を新規作成（教師のみ）
function actionCreateLesson(params) {
  checkTeacherPassword(params);
  if (!params.title) throw new Error('タイトルを入力してください');
  if (!params.question) throw new Error('発問（ふりかえりの問い）を入力してください');

  ensureLessonThemeColumn();
  ensureLessonFavoriteColumn();
  ensureLessonDeletedColumn();
  const sheet = getSheet(SHEET_LESSONS);
  const lessonId = 'L' + new Date().getTime();
  const published = isTrue(params.published);
  sheet.appendRow([
    lessonId,
    params.date || formatDate(new Date()),
    params.title,
    params.question,
    published,
    params.theme || '',
    false,
    false,
  ]);
  return { lessonId: lessonId };
}

// 授業の公開／非公開を切り替える（教師のみ）
function actionSetLessonPublished(params) {
  checkTeacherPassword(params);
  if (!params.lessonId) throw new Error('授業IDが必要です');
  const sheet = getSheet(SHEET_LESSONS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(params.lessonId)) {
      sheet.getRange(i + 1, 5).setValue(isTrue(params.published));
      return {};
    }
  }
  throw new Error('対象の授業が見つかりません');
}

// 授業のテーマを変更する（教師のみ）
function actionSetLessonTheme(params) {
  checkTeacherPassword(params);
  if (!params.lessonId) throw new Error('授業IDが必要です');
  const theme = params.theme ? String(params.theme).trim() : '';
  if (theme && getThemeList().indexOf(theme) === -1) {
    throw new Error('登録されていないテーマです');
  }
  ensureLessonThemeColumn();
  const sheet = getSheet(SHEET_LESSONS);
  const data = sheet.getDataRange().getValues();
  const themeCol = data[0].indexOf('テーマ');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(params.lessonId)) {
      sheet.getRange(i + 1, themeCol + 1).setValue(theme);
      return {};
    }
  }
  throw new Error('対象の授業が見つかりません');
}

// 授業のお気に入りを切り替える（教師のみ）
function actionSetLessonFavorite(params) {
  checkTeacherPassword(params);
  if (!params.lessonId) throw new Error('授業IDが必要です');
  ensureLessonFavoriteColumn();
  const sheet = getSheet(SHEET_LESSONS);
  const data = sheet.getDataRange().getValues();
  const favCol = data[0].indexOf('お気に入り');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(params.lessonId)) {
      sheet.getRange(i + 1, favCol + 1).setValue(isTrue(params.favorite));
      return {};
    }
  }
  throw new Error('対象の授業が見つかりません');
}

// 授業を削除する（論理削除。スプレッドシートにはデータを残し「削除済み」フラグだけ付ける・教師のみ）
function actionDeleteLesson(params) {
  checkTeacherPassword(params);
  if (!params.lessonId) throw new Error('授業IDが必要です');
  ensureLessonDeletedColumn();
  const sheet = getSheet(SHEET_LESSONS);
  const data = sheet.getDataRange().getValues();
  const delCol = data[0].indexOf('削除済み');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(params.lessonId)) {
      sheet.getRange(i + 1, delCol + 1).setValue(true);
      return {};
    }
  }
  throw new Error('対象の授業が見つかりません');
}

// 削除済みの授業を元に戻す（教師のみ）
function actionRestoreLesson(params) {
  checkTeacherPassword(params);
  if (!params.lessonId) throw new Error('授業IDが必要です');
  ensureLessonDeletedColumn();
  const sheet = getSheet(SHEET_LESSONS);
  const data = sheet.getDataRange().getValues();
  const delCol = data[0].indexOf('削除済み');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(params.lessonId)) {
      sheet.getRange(i + 1, delCol + 1).setValue(false);
      return {};
    }
  }
  throw new Error('対象の授業が見つかりません');
}

// 授業をスプレッドシートから完全に削除する（元に戻せない・教師のみ。生徒の回答データは「回答」シートに残る）
function actionPurgeLesson(params) {
  checkTeacherPassword(params);
  if (!params.lessonId) throw new Error('授業IDが必要です');
  const sheet = getSheet(SHEET_LESSONS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(params.lessonId)) {
      sheet.deleteRow(i + 1);
      return {};
    }
  }
  throw new Error('対象の授業が見つかりません');
}

// 削除済み（論理削除）の授業一覧を返す（教師のみ）
function actionGetDeletedLessons(params) {
  checkTeacherPassword(params);
  ensureLessonThemeColumn();
  ensureLessonFavoriteColumn();
  ensureLessonDeletedColumn();
  const objs = sheetToObjects(getSheet(SHEET_LESSONS));
  const lessons = objs
    .filter((o) => isTrue(o['削除済み']))
    .map((o) => ({
      lessonId: String(o['授業ID']),
      date: formatDate(o['日付']),
      title: o['タイトル'] || '',
      question: o['発問'] || '',
      published: isTrue(o['公開']),
      theme: o['テーマ'] || '',
      favorite: isTrue(o['お気に入り']),
    }));
  lessons.sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0));
  return { lessons: lessons };
}

// テーマの一覧（教師のみ）
function actionGetThemes(params) {
  checkTeacherPassword(params);
  return { themes: getThemeList() };
}

// テーマを新規追加（教師のみ）
function actionAddTheme(params) {
  checkTeacherPassword(params);
  const name = params.name && String(params.name).trim();
  if (!name) throw new Error('テーマ名を入力してください');

  const themes = getThemeList();
  if (themes.indexOf(name) !== -1) {
    throw new Error('同じテーマが既に登録されています');
  }
  ensureThemeSheet().appendRow([name]);
  return { themes: getThemeList() };
}

// テーマを削除（使用中の授業があるテーマは削除できない・教師のみ）
function actionRemoveTheme(params) {
  checkTeacherPassword(params);
  const name = params.name && String(params.name).trim();
  if (!name) throw new Error('テーマ名が必要です');

  ensureLessonThemeColumn();
  const lessonObjs = sheetToObjects(getSheet(SHEET_LESSONS));
  const inUse = lessonObjs.some((l) => String(l['テーマ'] || '') === name);
  if (inUse) {
    throw new Error('このテーマは使用中の授業があるため削除できません');
  }

  const sheet = ensureThemeSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === name) {
      sheet.deleteRow(i + 1);
      return { themes: getThemeList() };
    }
  }
  throw new Error('テーマが見つかりません');
}

// 生徒が「ふりかえり」を提出（既存があれば上書き更新）
function actionSubmitReflection(params) {
  checkStudentToken(params);
  if (!params.studentName) throw new Error('生徒名が必要です');
  if (!params.lessonId) throw new Error('授業IDが必要です');

  const sheet = getSheet(SHEET_RESPONSES);
  const data = sheet.getDataRange().getValues();
  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    if (
      String(data[i][1]) === String(params.lessonId) &&
      String(data[i][2]) === String(params.studentName)
    ) {
      sheet.getRange(i + 1, 4).setValue(params.text || '');
      sheet.getRange(i + 1, 5).setValue(now);
      return { updated: true };
    }
  }

  const responseId = 'R' + now.getTime();
  sheet.appendRow([
    responseId, params.lessonId, params.studentName, params.text || '', now, '', '', '', false,
  ]);
  return { created: true };
}

// 教師がある授業の全生徒の回答状況・内容を見る
function actionGetResponsesForLesson(params) {
  checkTeacherPassword(params);
  if (!params.lessonId) throw new Error('授業IDが必要です');

  const respObjs = sheetToObjects(getSheet(SHEET_RESPONSES)).filter(
    (o) => String(o['授業ID']) === String(params.lessonId)
  );
  const rosterNames = sheetToObjects(getSheet(SHEET_ROSTER))
    .map((o) => getStudentIdentifier(o))
    .filter((n) => !!n);

  const byName = {};
  respObjs.forEach((o) => {
    byName[o['生徒名']] = o;
  });

  const responses = rosterNames.map((name) => {
    const r = byName[name];
    return {
      studentName: name,
      submitted: !!r,
      text: r ? r['回答内容'] || '' : '',
      submittedAt: r ? formatDateTime(r['提出日時']) : '',
      comment: r ? r['教師コメント'] || '' : '',
      commentedAt: r ? formatDateTime(r['コメント日時']) : '',
      rating: r && r['評価'] !== '' && r['評価'] !== null && r['評価'] !== undefined ? Number(r['評価']) : null,
      liked: r ? isTrue(r['いいね']) : false,
    };
  });

  return { responses: responses };
}

// 教師がコメント・評価・いいねを保存する
function actionSaveComment(params) {
  checkTeacherPassword(params);
  if (!params.lessonId || !params.studentName) {
    throw new Error('授業IDと生徒名が必要です');
  }

  let rating = params.rating;
  if (rating !== undefined && rating !== null && String(rating).trim() !== '') {
    rating = Number(rating);
    if (isNaN(rating) || rating < 1 || rating > 10) {
      throw new Error('評価は1から10の範囲で指定してください');
    }
  } else {
    rating = '';
  }
  const liked = isTrue(params.liked);

  const sheet = getSheet(SHEET_RESPONSES);
  const data = sheet.getDataRange().getValues();
  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    if (
      String(data[i][1]) === String(params.lessonId) &&
      String(data[i][2]) === String(params.studentName)
    ) {
      sheet.getRange(i + 1, 6).setValue(params.comment || '');
      sheet.getRange(i + 1, 7).setValue(now);
      sheet.getRange(i + 1, 8).setValue(rating);
      sheet.getRange(i + 1, 9).setValue(liked);
      return {};
    }
  }
  throw new Error('対象の回答が見つかりません（生徒がまだ提出していない可能性があります）');
}

// 生徒の過去の記録（ふりかえり＋先生のコメント等）の中身を組み立てる
// 削除済み（論理削除）の授業に紐づく回答は除外する
function getMyRecordsList(studentName) {
  const respObjs = sheetToObjects(getSheet(SHEET_RESPONSES)).filter(
    (o) => o['生徒名'] === studentName
  );
  const lessonObjs = sheetToObjects(getSheet(SHEET_LESSONS));
  const lessonMap = {};
  const deletedLessonIds = {};
  lessonObjs.forEach((l) => {
    const id = String(l['授業ID']);
    lessonMap[id] = l;
    if (isTrue(l['削除済み'])) deletedLessonIds[id] = true;
  });

  const records = respObjs
    .filter((o) => {
      const id = String(o['授業ID']);
      // 授業が存在しない（完全削除など）／論理削除済みの授業に紐づく回答は生徒側に出さない
      return lessonMap[id] && !deletedLessonIds[id];
    })
    .map((o) => {
      const lesson = lessonMap[String(o['授業ID'])] || {};
      return {
        lessonId: String(o['授業ID']),
        title: lesson['タイトル'] || '(削除された授業)',
        date: formatDate(lesson['日付']),
        question: lesson['発問'] || '',
        theme: lesson['テーマ'] || '',
        text: o['回答内容'] || '',
        submittedAt: formatDateTime(o['提出日時']),
        comment: o['教師コメント'] || '',
        commentedAt: formatDateTime(o['コメント日時']),
        rating: o['評価'] !== '' && o['評価'] !== null && o['評価'] !== undefined ? Number(o['評価']) : null,
        liked: isTrue(o['いいね']),
      };
    });

  records.sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0));
  return records;
}

// 生徒が自分の過去の記録（ふりかえり＋先生のコメント等）を見る
function actionGetMyRecords(params) {
  checkStudentToken(params);
  if (!params.studentName) throw new Error('生徒名が必要です');
  return { records: getMyRecordsList(params.studentName) };
}

// 特定の生徒の全回答履歴をまとめて返す（テーマ別タブはフロント側で振り分ける・教師のみ）
function actionGetAllResponsesForStudent(params) {
  checkTeacherPassword(params);
  const studentName = params.studentName && String(params.studentName).trim();
  if (!studentName) throw new Error('生徒名が必要です');
  ensureLessonThemeColumn();
  return { records: getMyRecordsList(studentName) };
}

// 生徒用トークンが既に発行済みかどうかを確認する（教師のみ・新規発行はしない）
function actionGetStudentTokenStatus(params) {
  checkTeacherPassword(params);
  const settings = getSettingsMap();
  return { token: String(settings[SETTING_STUDENT_TOKEN] || '') };
}

// 特定の生徒の、テーマごとの平均評価（現在使用中のテーマだけを頂点にする・教師のみ）
function actionGetThemeEvaluationSummary(params) {
  checkTeacherPassword(params);
  const studentName = params.studentName && String(params.studentName).trim();
  if (!studentName) throw new Error('生徒名が必要です');

  ensureLessonThemeColumn();
  const lessonObjs = sheetToObjects(getSheet(SHEET_LESSONS));
  const themeByLessonId = {};
  const usedThemes = [];
  lessonObjs.forEach((l) => {
    if (isTrue(l['削除済み'])) return;
    const theme = String(l['テーマ'] || '').trim();
    themeByLessonId[String(l['授業ID'])] = theme;
    if (theme && usedThemes.indexOf(theme) === -1) usedThemes.push(theme);
  });

  const sums = {};
  const counts = {};
  sheetToObjects(getSheet(SHEET_RESPONSES))
    .filter((o) => o['生徒名'] === studentName)
    .forEach((o) => {
      const theme = themeByLessonId[String(o['授業ID'])];
      if (!theme) return;
      const rating = o['評価'];
      if (rating === '' || rating === null || rating === undefined) return;
      const n = Number(rating);
      if (isNaN(n)) return;
      sums[theme] = (sums[theme] || 0) + n;
      counts[theme] = (counts[theme] || 0) + 1;
    });

  const summary = usedThemes.map((theme) => ({
    theme: theme,
    average: counts[theme] ? Math.round((sums[theme] / counts[theme]) * 10) / 10 : 0,
  }));

  return { summary: summary };
}

// 生徒用URLのトークンを発行／再発行する（教師のみ）
function actionGetStudentToken(params) {
  checkTeacherPassword(params);

  const sheet = getSheet(SHEET_SETTINGS);
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === SETTING_STUDENT_TOKEN) {
      rowIndex = i;
      break;
    }
  }

  let token = rowIndex !== -1 ? String(data[rowIndex][1] || '') : '';
  if (!token || isTrue(params.regenerate)) {
    token = Utilities.getUuid().replace(/-/g, '');
    if (rowIndex !== -1) {
      sheet.getRange(rowIndex + 1, 2).setValue(token);
    } else {
      sheet.appendRow([SETTING_STUDENT_TOKEN, token]);
    }
  }

  return { token: token };
}

// 名簿の一覧（パスワードの値そのものは返さず、設定済みかどうかだけ返す）
function actionGetRosterAdmin(params) {
  checkTeacherPassword(params);
  const objs = sheetToObjects(getSheet(SHEET_ROSTER));
  const roster = objs
    .map((o) => ({
      number: o['出席番号'] !== undefined && o['出席番号'] !== null ? String(o['出席番号']) : '',
      name: getStudentIdentifier(o),
      hasName: !!(o['生徒名'] && String(o['生徒名']).trim()),
      hasPassword: !!(o['パスワード'] && String(o['パスワード']).trim()),
    }))
    .filter((s) => !!s.name);
  return { roster: roster };
}

// 生徒を名簿に追加する（パスワードは空のまま＝本人が初回ログイン時に設定。
// 生徒名を空にして出席番号だけで登録することもできる）
function actionAddStudent(params) {
  checkTeacherPassword(params);
  const name = params.name && String(params.name).trim();
  const number = params.number && String(params.number).trim();
  if (!name && !number) {
    throw new Error('生徒名か出席番号のどちらかを入力してください');
  }

  const sheet = getSheet(SHEET_ROSTER);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const identifier = getStudentIdentifier({ '生徒名': name, '出席番号': number });

  if (findRosterRowIndex(data, headers, identifier) !== -1) {
    throw new Error('同じ生徒（名前または出席番号）が既に登録されています');
  }

  sheet.appendRow([number || '', name || '', '']);
  return {};
}

// 生徒を名簿から削除する（過去の回答データ自体は「回答」シートに残る）
function actionRemoveStudent(params) {
  checkTeacherPassword(params);
  const name = params.name && String(params.name).trim();
  if (!name) throw new Error('生徒名が必要です');

  const sheet = getSheet(SHEET_ROSTER);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rowIndex = findRosterRowIndex(data, headers, name);
  if (rowIndex === -1) throw new Error('名簿に見つかりません');

  sheet.deleteRow(rowIndex + 1);
  return {};
}

// 教師がパスワードをリセットする（空にする＝次回ログイン時に本人が再登録）
function actionResetStudentPassword(params) {
  checkTeacherPassword(params);
  const name = params.name && String(params.name).trim();
  if (!name) throw new Error('生徒名が必要です');

  const sheet = getSheet(SHEET_ROSTER);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const pwCol = headers.indexOf('パスワード');
  const rowIndex = findRosterRowIndex(data, headers, name);
  if (rowIndex === -1) throw new Error('名簿に見つかりません');

  sheet.getRange(rowIndex + 1, pwCol + 1).setValue('');
  clearRateLimit('studentLogin_' + name);
  clearRateLimit('studentRegister_' + name);
  return {};
}

// 教師がパスワードを直接指定して設定する
function actionSetStudentPassword(params) {
  checkTeacherPassword(params);
  const name = params.name && String(params.name).trim();
  if (!name) throw new Error('生徒名が必要です');
  const newPassword = params.newPassword;
  if (!newPassword || String(newPassword).length < 4) {
    throw new Error('パスワードは4文字以上で設定してください');
  }

  const sheet = getSheet(SHEET_ROSTER);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const pwCol = headers.indexOf('パスワード');
  const rowIndex = findRosterRowIndex(data, headers, name);
  if (rowIndex === -1) throw new Error('名簿に見つかりません');

  sheet.getRange(rowIndex + 1, pwCol + 1).setValue(String(newPassword));
  clearRateLimit('studentLogin_' + name);
  return {};
}
