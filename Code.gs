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
    settings.getRange('A1:B5').setValues([
      ['設定項目', '値'],
      ['クラス名', '6年1組'],
      ['担任名', ''],
      ['教師パスワード', 'sensei123'],
      ['表示方法(本名/番号/匿名)', '本名'],
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
    lessons.getRange('A1:E1').setValues([['授業ID', '日付', 'タイトル', '発問', '公開']]);
    lessons.getRange('A1:E1').setFontWeight('bold');
    lessons.setColumnWidths(1, 5, 160);
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
        result = actionGetRoster();
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
      case 'getLessons':
        result = actionGetLessons(params);
        break;
      case 'createLesson':
        result = actionCreateLesson(params);
        break;
      case 'setLessonPublished':
        result = actionSetLessonPublished(params);
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

function findRosterRowIndex(data, headers, name) {
  const nameCol = headers.indexOf('生徒名');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][nameCol]).trim() === name) return i;
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
  };
}

// 生徒ログイン画面の名前選択肢
function actionGetRoster() {
  const objs = sheetToObjects(getSheet(SHEET_ROSTER));
  const roster = objs.map((o) => o['生徒名']).filter((n) => !!n);
  return { roster: roster };
}

// 選んだ生徒がパスワード登録済みかどうかを確認する
function actionCheckStudentAccount(params) {
  const name = params.name && String(params.name).trim();
  if (!name) throw new Error('名前を選んでください');

  const objs = sheetToObjects(getSheet(SHEET_ROSTER));
  const match = objs.find((o) => String(o['生徒名']).trim() === name);
  if (!match) throw new Error('名簿に見つかりません');

  const pw = match['パスワード'];
  return { hasPassword: !!(pw && String(pw).trim()) };
}

// 初回パスワード登録（パスワード列が空のときだけ成功する）
function actionRegisterStudentPassword(params) {
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
  return { studentName: data[rowIndex][headers.indexOf('生徒名')] };
}

// 生徒ログイン（名前＋パスワード）
function actionStudentLogin(params) {
  const name = params.name && String(params.name).trim();
  if (!name) throw new Error('名前を選んでください');
  if (!params.password) throw new Error('パスワードを入力してください');

  const bucket = 'studentLogin_' + name;
  const MAX_ATTEMPTS = 6;
  const LOCK_SECONDS = 300; // 5分

  checkRateLimit(bucket);

  const objs = sheetToObjects(getSheet(SHEET_ROSTER));
  const match = objs.find((o) => String(o['生徒名']).trim() === name);
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
  return { studentName: match['生徒名'] };
}

// 教師ログイン（パスワード確認のみ。セッションは持たず、毎回パスワードを送る方式）
function actionTeacherLogin(params) {
  checkTeacherPassword(params);
  const settings = getSettingsMap();
  return { className: settings['クラス名'] || '', teacherName: settings['担任名'] || '' };
}

// 授業一覧（teacher は全件、student は公開済みのみ）
function actionGetLessons(params) {
  const sheet = getSheet(SHEET_LESSONS);
  const objs = sheetToObjects(sheet);
  let lessons = objs.map((o) => ({
    lessonId: String(o['授業ID']),
    date: formatDate(o['日付']),
    title: o['タイトル'] || '',
    question: o['発問'] || '',
    published: isTrue(o['公開']),
  }));
  if (params.role !== 'teacher') {
    lessons = lessons.filter((l) => l.published);
  }
  lessons.sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0));
  return { lessons: lessons };
}

// 授業を新規作成（教師のみ）
function actionCreateLesson(params) {
  checkTeacherPassword(params);
  if (!params.title) throw new Error('タイトルを入力してください');
  if (!params.question) throw new Error('発問（ふりかえりの問い）を入力してください');

  const sheet = getSheet(SHEET_LESSONS);
  const lessonId = 'L' + new Date().getTime();
  const published = isTrue(params.published);
  sheet.appendRow([
    lessonId,
    params.date || formatDate(new Date()),
    params.title,
    params.question,
    published,
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

// 生徒が「ふりかえり」を提出（既存があれば上書き更新）
function actionSubmitReflection(params) {
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
    .map((o) => o['生徒名'])
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

// 生徒が自分の過去の記録（ふりかえり＋先生のコメント等）を見る
function actionGetMyRecords(params) {
  if (!params.studentName) throw new Error('生徒名が必要です');

  const respObjs = sheetToObjects(getSheet(SHEET_RESPONSES)).filter(
    (o) => o['生徒名'] === params.studentName
  );
  const lessonObjs = sheetToObjects(getSheet(SHEET_LESSONS));
  const lessonMap = {};
  lessonObjs.forEach((l) => {
    lessonMap[String(l['授業ID'])] = l;
  });

  const records = respObjs.map((o) => {
    const lesson = lessonMap[String(o['授業ID'])] || {};
    return {
      lessonId: String(o['授業ID']),
      title: lesson['タイトル'] || '(削除された授業)',
      date: formatDate(lesson['日付']),
      question: lesson['発問'] || '',
      text: o['回答内容'] || '',
      submittedAt: formatDateTime(o['提出日時']),
      comment: o['教師コメント'] || '',
      commentedAt: formatDateTime(o['コメント日時']),
      rating: o['評価'] !== '' && o['評価'] !== null && o['評価'] !== undefined ? Number(o['評価']) : null,
      liked: isTrue(o['いいね']),
    };
  });

  records.sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0));
  return { records: records };
}

// ============================================================
// 教師用：生徒の管理（スプレッドシートを直接開かずに操作できるようにする）
// ============================================================

// 名簿の一覧（パスワードの値そのものは返さず、設定済みかどうかだけ返す）
function actionGetRosterAdmin(params) {
  checkTeacherPassword(params);
  const objs = sheetToObjects(getSheet(SHEET_ROSTER));
  const roster = objs
    .filter((o) => !!o['生徒名'])
    .map((o) => ({
      number: o['出席番号'] !== undefined && o['出席番号'] !== null ? String(o['出席番号']) : '',
      name: o['生徒名'],
      hasPassword: !!(o['パスワード'] && String(o['パスワード']).trim()),
    }));
  return { roster: roster };
}

// 生徒を名簿に追加する（パスワードは空のまま＝本人が初回ログイン時に設定）
function actionAddStudent(params) {
  checkTeacherPassword(params);
  const name = params.name && String(params.name).trim();
  if (!name) throw new Error('生徒名を入力してください');

  const sheet = getSheet(SHEET_ROSTER);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('生徒名');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][nameCol]).trim() === name) {
      throw new Error('同じ名前の生徒が既に登録されています');
    }
  }

  sheet.appendRow([params.number || '', name, '']);
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
