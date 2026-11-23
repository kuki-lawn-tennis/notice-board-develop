# GitHub Pages / Google Apps Script の分離

## 配置

- `index.html` は GitHub Pages の公開ルートに配置します。
- `gas/Code.js` の内容をGASのスクリプトファイル（通常は `Code.gs`）、`gas/ApiBridge.html` をHTMLファイル `ApiBridge.html` として配置します。
- スプレッドシート、Drive、スクリプトプロパティなど、現在のGASプロジェクト設定は引き継いでください。

GitHub Pages とGASの間は、GASのHTML Serviceを隠しiframeとして読み込み、`postMessage`でRPCを行います。GAS側で `google.script.run` を呼び出すため、ブラウザーからGASへ直接 `fetch` する際のCORS制約を避けられます。ブリッジは許可したメソッドだけを実行し、親ページのオリジンも照合します。

## GASの設定とデプロイ

1. Apps Script の「プロジェクトの設定」→「スクリプト プロパティ」に `GITHUB_PAGES_ORIGINS` を追加します。値にはページのオリジンだけを指定します（例: `https://kuki-lawn-tennis.github.io`）。独自ドメインやローカル確認用URLがある場合はカンマ区切りで追加できます。パスや末尾の `/` は含めません。
2. `gas/Code.js` の内容をGASのスクリプトファイル（通常は `Code.gs`）、`gas/ApiBridge.html` をHTMLファイル `ApiBridge.html` に反映します。ウェブアプリとして新しいバージョンをデプロイし、実行ユーザーを自分、アクセスできるユーザーを公開サイトの利用者に設定します。管理操作は既存の管理者セッショントークンで保護されます。
3. デプロイされた `/exec` URLをコピーします。

## コート当番フォルダの設定

1. Apps Script のスクリプトプロパティ `COURT_DUTY_FOLDER_ID` に、保存先DriveフォルダのID（またはフォルダURL）を設定します。IDはフォルダURLの `/folders/` の後ろの値です。
2. ウェブアプリの実行ユーザー（上記設定ならデプロイした自分のGoogleアカウント）に、そのフォルダへのアクセス権があることを確認します。共有ドライブの場合も同じアカウントがメンバーである必要があります。
3. Apps Script エディターで `verifyCourtDutyFolderAccess` を選んで実行し、初回のGoogle Drive権限確認を完了します。成功すれば実行ログに確認メッセージが出ます。
4. ウェブアプリを新しいバージョンにデプロイします。

## お知らせ・イベント・スケジュールの装飾と添付

- 編集画面で重要度を「重要」または「通常」に設定できます。重要な投稿は公開画面で本文まで常時表示し、通常の投稿は従来どおりタイトルを開いて本文を表示します。既存の投稿は通常として扱われます。
- 管理画面の本文欄で太字、下線、取り消し線、斜体、リンク、文字色・背景色（各8色）を指定できます。本文は許可したタグ・色・リンクだけを保存・表示します。
- 1回の保存で最大5ファイル、合計10MBまでアップロードできます。PDF、Office文書、CSV、テキスト、画像に対応します。
- 編集画面で既存添付の「削除」を選び、投稿を保存すると、該当するDriveファイルをゴミ箱へ移動します。保存せずキャンセルした場合は削除されません。
- Apps Script のスクリプトプロパティに、別々のDriveフォルダを指定してください：`NOTICE_ATTACHMENTS_FOLDER_ID`（お知らせ）、`EVENT_ATTACHMENTS_FOLDER_ID`（イベント）、`SCHEDULE_ATTACHMENTS_FOLDER_ID`（スケジュール）。コート当番用フォルダとも分けます。
- Apps Script エディターで `verifyNoticeAttachmentFolderAccess`、`verifyEventAttachmentFolderAccess`、`verifyScheduleAttachmentFolderAccess` をそれぞれ実行し、書き込みとリンク共有の権限を確認してください。各関数は確認用ファイルを一時作成して共有を試し、最後にゴミ箱へ移動します。成功後、ウェブアプリを新しいバージョンでデプロイしてください。既存のシートには保存時に「添付ファイル」列が追加されます。
- 添付の保存API `saveItemWithAttachments` を使うため、`gas/Code.js` と `gas/ApiBridge.html` の両方を反映してから、新しいGASバージョンをデプロイしてください。古いブリッジではこのAPIが拒否されます。
- 添付は公開ページから開けるよう、アップロード時にDriveの共有範囲を「リンクを知っている全員（閲覧者）」にします。リンクを知る人は誰でもファイルを閲覧できるため、公開してよいファイルだけを添付してください。
- 投稿自体を削除してもDrive上の添付ファイルは自動削除されません。不要になったファイルは、添付専用フォルダで共有設定を解除して削除してください。

## GitHub Pages の設定

1. リポジトリの **Settings → Secrets and variables → Actions → Variables** で、`GAS_WEB_APP_URL` というRepository variableを作成し、手順2の `/exec` URLを設定します。これは秘密情報ではないため、SecretではなくVariableを使います。
2. **Settings → Pages → Build and deployment → Source** を **GitHub Actions** に設定します。
3. `main`へ`index.html`またはワークフローをpushすると、デプロイ時にURLがHTMLへ注入されます。Actionsタブから手動実行することもできます。値が未設定または不正な場合はデプロイを失敗させます。

ローカルで試す場合は、`index.html`の`GAS_WEB_APP_URL`プレースホルダーを一時的に `/exec` URLへ置き換え、HTTPサーバーからページを開いてください。そのオリジン（例: `http://localhost:8000`）は`GITHUB_PAGES_ORIGINS`に追加します。`file://`からは利用できません。

GASのHTML Serviceはブリッジを内側のiframeで実行するため、ブリッジから親ページへ接続通知を送り、そのウィンドウと `postMessage` で通信します。`Code.js` または `ApiBridge.html` を変更した場合は、Apps Script で新しいバージョンをデプロイしてください。
