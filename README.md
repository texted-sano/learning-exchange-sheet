# 学習交流シート（プロトタイプ）開発者向けREADME

生徒が授業の「ふりかえり」を入力・保存し、教員がコメント・10段階評価・いいねをつけ、生徒が後日見返せるツールです。データの保存先はGoogleスプレッドシート、バックエンドはGoogle Apps Script（GAS）、フロントエンドは単一のHTMLファイルです。

> 先生向けの導入手順は `先生向けマニュアル.md` を参照してください。本書は開発・保守を行う人向けです。

---

## 1. 全体構成

```
[index.html（フロントエンド）]
        │ fetch(POST, text/plain + JSON文字列)
        ▼
[Code.gs（GAS ウェブアプリ = HTTP API）]
        │ SpreadsheetApp
        ▼
[Googleスプレッドシート（データ本体）]
```

- **1スプレッドシート = 1クラス** の想定です。教師・クラス・生徒・授業・回答・テーマの情報はすべてスプレッドシートに保存され、`index.html` 側には個別設定を一切持ちません（全教師に同一ファイルを配布できます）。
- フロントエンドは `localStorage` 等を使わず、状態はメモリ内のみです。リロードすると接続・ログイン状態は消えます。
- 通信は常に `Content-Type: text/plain;charset=utf-8` のPOSTで、ボディにJSON文字列を載せます。これはGASウェブアプリに対してCORSプリフライト（OPTIONS）を発生させないための定石です。GAS側は `e.postData.contents` をJSONとしてパースします。

## 2. シート構成

`initializeSheets()`（スプレッドシートのメニュー「学習交流シート > シートを初期化」）が以下を作成します。**既存シートには一切手を付けません**（無いものだけ作る冪等設計）。

| シート | 列 | 備考 |
|---|---|---|
| 設定 | 設定項目 / 値 | クラス名、担任名、教師パスワード（初期値は空）、表示方法、生徒用トークン |
| 名簿 | 出席番号 / 生徒名 / パスワード | 生徒名は空欄可（出席番号のみの匿名運用に対応） |
| 授業 | 授業ID / 日付 / タイトル / 発問 / 公開 / テーマ / お気に入り / 削除済み | 授業IDは `L<epoch millis>`（`L`以降の数値が作成時刻ミリ秒。フロントはこれを作成時刻表示に利用）。削除は論理削除で、`削除済み` がTRUEの授業は生徒・教師の通常ビューから除外される |
| テーマ | テーマ | 初期値：平和・命・多様性・災害・家族 |
| 回答 | 回答ID / 授業ID / 生徒名 / 回答内容 / 提出日時 / 教師コメント / コメント日時 / 評価 / いいね | 回答IDは `R<epoch millis>`。評価は1〜10または空 |

### 自己修復（マイグレーション）処理

旧バージョンのスプレッドシートを新コードで動かした場合に備え、以下のヘルパーがアクセス時に不足分を自動補完します。

- `ensureThemeSheet()` … 「テーマ」シートが無ければ既定5テーマ入りで作成
- `ensureLessonThemeColumn()` … 「授業」シートに「テーマ」列が無ければヘッダーを追加
- `ensureLessonFavoriteColumn()` … 「授業」シートに「お気に入り」列が無ければヘッダーを追加
- `ensureLessonDeletedColumn()` … 「授業」シートに「削除済み」列が無ければヘッダーを追加
- `actionGetStudentToken` … 「設定」シートに「生徒用トークン」行が無ければ `appendRow` で追加

これらの `ensure*Column` は列名でヘッダー検索し、末尾に不足列を追加する方式のため、既存の列順・データには影響しません（`sheetToObjects` も列名アクセス）。

## 3. 認証・セキュリティモデル

試作品としての割り切りを含みます。**本運用前に必ず「7. 既知の制約」を確認してください。**

### 教師
- 「設定」シートの「教師パスワード」と照合。セッションは持たず、**教師系アクションは毎回パスワードをパラメータで送って検証**します（`checkTeacherPassword`）。
- パスワード未設定（空）の場合、フロントは接続直後に初回設定モーダルを表示し、`registerTeacherPassword` を呼びます。このアクションは**既にパスワードが設定済みなら必ず失敗**します（乗っ取り防止）。

### 生徒
- 二要素です：**①生徒用トークン**（クラス共通・教師が発行）＋ **②本人パスワード**（名簿の行ごと）。
- 生徒系アクションはすべて先頭で `checkStudentToken(params)` を実行します。トークン未発行・不一致なら即エラーです。exec URLを直接叩かれても、トークンが無ければ名簿の取得すらできません。
- トークンは `Utilities.getUuid()` のハイフン除去（32文字）。`getStudentToken` に `regenerate: 'true'` を渡すと再生成・上書きされ、**旧トークンの接続コードはその瞬間に無効**になります。
- 生徒の初回パスワード登録は「名簿のパスワード列が空のときだけ成功」します。

### 生徒の識別子
`getStudentIdentifier(row)` が単一のルールを提供します：**生徒名があれば生徒名、無ければ「出席番号＋番」（例：`5番`）**。回答シートの「生徒名」列にはこの識別子がそのまま入ります。名簿・回答・テーマ集計はすべてこの識別子で紐づくため、**運用開始後に生徒名を後から追記すると識別子が変わり、過去の回答と切れます**（既知の制約）。

### レート制限
`CacheService` による簡易実装です（IPアドレスはGASから取得できないため）。
- 教師ログイン：アプリ全体で5回失敗→5分ロック
- 生徒ログイン：名前ごとに6回失敗→5分ロック
- 生徒初回登録：名前ごとに10回失敗→10分ロック
- 教師による `resetStudentPassword` / `setStudentPassword` は該当生徒のロックを解除します

## 4. APIリファレンス

エンドポイントはウェブアプリの exec URL。リクエストは `{ "action": "...", ...params }` のJSON。レスポンスは成功時 `{ ok: true, ... }`、失敗時 `{ ok: false, error: "メッセージ" }`（HTTPステータスは常に200）。

認証列の凡例 … **T**: `password`（教師パスワード）必須 ／ **S**: `token`（生徒用トークン）必須 ／ −: 不要

| action | 認証 | 主なパラメータ | 返り値（抜粋） | 説明 |
|---|---|---|---|---|
| `ping` | − | | `className, teacherName, hasTeacherPassword` | 接続確認。初回モーダルの要否判定に使用 |
| `teacherLogin` | T | | `className, teacherName, lessons[], themes[]` | 授業一覧・テーマ一覧を同梱（通信削減） |
| `registerTeacherPassword` | − | `password` | 同上 | パスワード未設定時のみ成功 |
| `updateTeacherSettings` | T | `teacherName?, className?, newPassword?` | `className, teacherName` | まとめて更新 |
| `getStudentToken` | T | `regenerate?` | `token` | 未発行なら生成。`regenerate:'true'`で再生成 |
| `getStudentTokenStatus` | T | | `token` | 既存トークンの確認のみ（生成しない） |
| `getRoster` | S | | `roster[]`（識別子の配列） | 生徒ログインの名前選択肢 |
| `checkStudentAccount` | S | `name` | `hasPassword` | 初回登録かログインかの分岐 |
| `registerStudentPassword` | S | `name, password` | `studentName, lessons[], records[]` | 初回のみ。ログイン相当の情報を同梱 |
| `studentLogin` | S | `name, password` | 同上 | |
| `getLessons` | T/S | `role`（`teacher`なら要password、他は要token） | `lessons[]`（各要素に `theme, favorite` を含む） | teacherは全件、それ以外は公開済みのみ。**いずれも論理削除済み（削除済み=TRUE）は除外** |
| `createLesson` | T | `date?, title, question, published?, theme?` | `lessonId` | お気に入り・削除済みはFALSEで作成 |
| `setLessonPublished` | T | `lessonId, published` | | |
| `setLessonTheme` | T | `lessonId, theme` | | 作成後のテーマ変更。空文字で未設定化。登録済みテーマ以外はエラー |
| `setLessonFavorite` | T | `lessonId, favorite` | | お気に入りON/OFF |
| `deleteLesson` | T | `lessonId` | | **論理削除**（削除済み=TRUEを立てるだけ。行は残す） |
| `restoreLesson` | T | `lessonId` | | 論理削除の取り消し（削除済み=FALSE） |
| `purgeLesson` | T | `lessonId` | | **完全削除**（`deleteRow` で物理削除。元に戻せない）。回答データは残る |
| `getDeletedLessons` | T | | `lessons[]` | 論理削除済みの授業だけを返す（「削除した授業」ビュー用） |
| `getThemes` | T | | `themes[]` | |
| `addTheme` | T | `name` | `themes[]` | 重複は不可 |
| `removeTheme` | T | `name` | `themes[]` | **使用中の授業があるテーマは削除不可** |
| `submitReflection` | S | `studentName, lessonId, text` | `created` or `updated` | 既存回答があれば上書き |
| `getResponsesForLesson` | T | `lessonId` | `responses[]`（名簿全員分。未提出者は `submitted:false`） | |
| `saveComment` | T | `lessonId, studentName, comment?, rating?, liked?` | | 評価は1〜10か空。範囲外はエラー |
| `getMyRecords` | S | `studentName` | `records[]` | 各recordに `theme` を含む。**論理削除済み・授業が存在しない（完全削除された）回答は除外**（下記メモ参照） |
| `getAllResponsesForStudent` | T | `studentName` | `records[]` | テーマ別モーダル用の一括取得（タブ切替はフロント側でフィルタ）。`getMyRecordsList` 共用のため上記と同じ除外が効く |
| `getThemeEvaluationSummary` | T | `studentName` | `summary[]`（`{theme, average}`） | 使用中テーマのみ。評価なしは `average: 0`。**論理削除済みの授業は集計対象外** |
| `getRosterAdmin` | T | | `roster[]`（`number, name, hasName, hasPassword`） | パスワード値そのものは返さない |
| `addStudent` | T | `name?, number?`（どちらか必須） | | |
| `removeStudent` | T | `name`（識別子） | | 回答データは残る |
| `resetStudentPassword` | T | `name` | | 空にする＝本人が再登録 |
| `setStudentPassword` | T | `name, newPassword` | | 教師が直接指定 |

## 5. フロントエンド設計メモ

- **単一の `state` オブジェクト＋全画面再描画方式**（`render()` が `innerHTML` を組み直す素朴なSPA）。イベントは `onclick` 属性でグローバル関数を呼びます。
- **非同期ハンドラの規約**：`state.busy = true; render();` → try/catch → `finally{ state.busy = false; render(); }`。`finally` の `render()` を忘れると「読み込み中…」で固まります（過去に実際に発生）。
- **通信削減の方針**：ログイン応答に授業一覧（生徒はさらに記録）を同梱／書き込み成功後は再取得せずローカルのstateを直接更新（授業作成・公開切替・評価保存・生徒管理）。応答に同梱データが無い場合は個別取得にフォールバックし、フロントとGASのバージョンずれでも壊れないようにしています。
- **`Promise.allSettled`**：複数の並列取得は1つの失敗が他を巻き込まないようにallSettledを使います（例：生徒管理タブの 名簿＋テーマ＋トークン確認）。
- **モーダル**：高さが変動する「テーマ」「評価まとめ」「削除した授業」は `.modal-overlay.top-aligned` で上端固定（中央揃えだとタブ切替時に見出しが上下に踊るため）。これらのモーダル表示中はグローバルの読み込みオーバーレイを抑制し（`modalHandlesOwnLoading`）、モーダル内のインライン表示に任せます。
- **削除の扱い（論理削除）**：授業の削除は物理削除ではなく「削除済み」列のフラグ立てです。生徒側の2経路（`getLessonsList`＝「授業」タブ、`getMyRecordsList`＝「きろくを見る」）に加え、教師側の通常一覧・テーマ別モーダル・評価まとめからも一貫して除外し、教師は「削除した授業」モーダル（`getDeletedLessons`）でのみ確認・復活（`restoreLesson`）・完全削除（`purgeLesson`）します。`getMyRecordsList` は「削除済みフラグ」に加えて「対応する授業行が存在しないレコード」も除外するため、完全削除された授業の回答も生徒側には出ません（以前の `(削除された授業)` フォールバック表示は生徒側では実質使われません）。復活の一貫性のため、この除外方針は共用関数側にまとめています。
- **作成時刻の表示**：授業一覧カードの日付欄に「作成 HH:MM」を併記します。これは授業ID（`L<epoch millis>`）からフロント側で算出（`creationTimeFromLessonId`）しており、シートに時刻列は持ちません。過去データにもそのまま出ますが、これは「日付」列（先生が選んだ授業日）とは別物である点に注意。
- **カードの折りたたみ**：教師の提出状況カード（`state.expandedResponses`）に加え、生徒の「きろくを見る」も既定で畳んだ状態にし、`state.expandedRecords`（`lessonId` キー）で開閉します。「きろくを見る」タブを開くたびに `expandedRecords` をリセットして全カードを畳みます。
- **テーマ変更UI**：授業一覧の各カードのテーマは、他タブのタグ（`.theme-chip`）に寄せた `<select>` で、ラベルの右に横並び（`.lesson-theme-row`）。現在のテーマがテーマ一覧に無い（手動編集された旧データ等）場合は先頭に補完して選択状態を保ちます。
- **レーダーチャート**：`buildRadarSvg(categories, values, maxValue)` が素のSVG文字列を生成。頂点数はテーマ数に応じて可変、目盛りは1〜10全段階に点線円を描画します。
- **生徒モード**：`index.html#api=<encodeURIComponent(execURL)>&t=<token>` のハッシュで開くと接続画面・教師UIを出さず生徒ログインのみ表示。ローカルファイル運用ではURL共有が意味を持たないため、教師画面では **URLではなく「接続コード」（`api=...&t=...` のテキスト）** を発行し、生徒は接続画面の「生徒の方はこちら」に貼り付けて接続します。

## 6. 開発の流れ

1. `Code.gs` を修正 → GASエディタに貼り付けて保存
2. **「デプロイを管理」→ 編集（鉛筆）→ バージョン「新しいバージョン」→ デプロイ**
   - 保存しただけでは反映されません。「不明な action です」という既知エラーの原因は、ほぼ再デプロイ漏れです。
3. `index.html` を差し替えてブラウザで開く

デバッグ時は exec URL に直接POSTして `{ok:false, error}` の中身を見るのが早いです。`action=getRoster` をトークンなしで叩いてエラーになることが、生徒トークン保護の動作確認になります。

## 7. 既知の制約（本運用前チェックリスト）

| 項目 | 内容 | 対応案 |
|---|---|---|
| パスワード平文保存 | 名簿・設定シートにそのまま保存 | ハッシュ化、またはプラットフォーム側認証との連携 |
| 初回登録の先着問題 | 名簿に名前があれば誰でも初回パスワードを登録できる | 教師による初期パスワード配布方式、登録期間の限定 |
| 識別子の変更に弱い | 生徒名の後付け・改名で過去回答と切れる | 内部IDの導入（名簿にID列を追加し回答をIDで紐づけ） |
| レート制限が簡易 | IP不明のためアプリ全体／名前単位 | プラットフォーム側での対策 |
| 1シート=1クラス | 複数クラスはスプレッドシートを分ける運用 | クラスコードで接続先を切り替える仕組み |
| 同時書き込み | LockService未使用。同時操作で行ズレの理論的可能性 | LockServiceの導入 |
| 完全削除と回答の孤立 | 授業を完全削除（`purgeLesson`）しても「回答」シートの回答行は残る。孤立した回答は生徒・教師の記録から非表示になるだけでシート上には蓄積する | 回答も併せて削除するか、定期的な棚卸し |
| GASの応答速度 | 1リクエスト1秒前後が下限 | 根本対策は専用サーバーへの移行 |

## 8. 拡張のアイデア

- AIによる発問・コメント下書きの生成（Gemini API等）
- 授業中のリアルタイム進行（定期ポーリング等）
- 出版社サーバーへのバックエンド移行：フロントは「接続先URLを差し替えるだけ」で移行できるよう、API境界（本書4章）を維持したまま実装を置き換える設計を推奨
